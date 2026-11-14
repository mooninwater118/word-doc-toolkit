# -*- coding: utf-8 -*-
"""
文档批量处理工具 - 本地 Web 服务
功能：doc→docx 转换 / doc 文件清理 / 关键词查找 / 批量替换
"""
import os
import sys
import time
import json
import shutil
import threading
import traceback
import hashlib

from fastapi import FastAPI
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import List, Optional

import fuzzy

# ---------- 路径解析（兼容源码运行 / PyInstaller 打包）----------
# 打包成 exe 后：静态资源在临时解压目录（sys._MEIPASS，只读），
# 日志必须写到 exe 旁边（可写），绝不能写进临时目录，否则程序退出就没了。
if getattr(sys, 'frozen', False):
    _RES = sys._MEIPASS                       # PyInstaller 解压目录
    _APP = os.path.dirname(sys.executable)    # exe 所在目录
else:
    _RES = os.path.dirname(os.path.abspath(__file__))
    _APP = _RES

STATIC = os.path.join(_RES, 'static')


def _pick_logdir():
    """挑一个可写的日志目录。exe 放在只读位置（如 Program Files）时自动降级。"""
    cands = [
        os.path.join(_APP, 'logs'),
        os.path.join(os.path.expanduser('~'), 'DocToolkit', 'logs'),
        os.path.join(os.environ.get('TEMP') or '.', 'DocToolkit', 'logs'),
    ]
    for c in cands:
        try:
            os.makedirs(c, exist_ok=True)
            probe = os.path.join(c, '.write_test')
            with open(probe, 'w') as f:
                f.write('x')
            os.remove(probe)
            return c
        except Exception:
            continue
    return os.environ.get('TEMP') or '.'


LOGDIR = _pick_logdir()

# 本次运行的日志文件：程序一启动就建，每次运行一个独立文件（不按天合并）
RUN_ID = time.strftime('%Y%m%d_%H%M%S')
LOG_FILE = os.path.join(LOGDIR, 'run_%s.log' % RUN_ID)

WD_FORMAT_XML_DOCUMENT = 12
WD_REPLACE_NONE = 0
WD_REPLACE_ALL = 2
WD_WRAP_STOP = 0


def now():
    return time.strftime('%H:%M:%S')


# ============ 全局状态 ============
LOCK = threading.Lock()
STATE = {
    'running': False,
    'task': '',
    'task_label': '',
    'total': 0,
    'done': 0,
    'current': '',
    'logs': [],
    'seq': 0,
    'result': {},
    'cancel': False,
    'started_at': None,
    'finished_at': None,
    'logfile': LOG_FILE,
    # 卡死上报：某个文件已等待多久没有响应，供前端提示「点中断」
    'stuck_file': '',
    'stuck_seconds': 0,
    'stuck_since_ts': 0,
    'last_kill': {},
}


class Ctx:
    def __init__(self):
        self.lock = LOCK
        self._fh = None
        # 日志按天追加写，所有任务共用一个文件，不再每个任务覆盖
        lf = LOG_FILE
        with LOCK:
            STATE['logfile'] = lf
        try:
            self._fh = open(lf, 'a', encoding='utf-8')
        except Exception:
            self._fh = None

    def log(self, msg, level='info'):
        line = {'t': now(), 'level': level, 'msg': str(msg), 'seq': 0}
        with self.lock:
            STATE['seq'] += 1
            line['seq'] = STATE['seq']
            STATE['logs'].append(line)
            if len(STATE['logs']) > 5000:
                del STATE['logs'][:1000]
        if self._fh:
            try:
                self._fh.write('[%s][%s] %s\n' % (line['t'], level, msg))
                self._fh.flush()
            except Exception:
                pass

    def progress(self, done=None, total=None, current=None):
        with self.lock:
            if done is not None:
                STATE['done'] = done
            if total is not None:
                STATE['total'] = total
            if current is not None:
                STATE['current'] = current

    def check_cancel(self):
        with self.lock:
            return STATE['cancel']

    def close(self):
        if self._fh:
            try:
                self._fh.close()
            except Exception:
                pass


def write_run_header():
    """程序一启动就把日志文件建好，写入启动信息和引擎检测结果。

    之前引擎检测只在前端渲染，后端不落盘，所以日志文件里看不到那几行。
    现在启动就写，即使一次任务都没跑，点「下载」也能拿到完整启动记录。
    """
    lines = [
        '=' * 62,
        '文档批量处理工具  启动于 %s' % time.strftime('%Y-%m-%d %H:%M:%S'),
        '程序目录：%s' % mask_path(_APP),
        '日志目录：%s' % mask_path(LOGDIR),
        '-' * 62,
    ]
    try:
        info = detect_engines()
    except Exception as e:
        info = None
        lines.append('文档引擎检测失败：%s' % e)
    if info:
        lines.append('文档引擎检测：')
        lines.append('  WPS Office      %s' % ('可用' if info['wps_available'] else '未检测到'))
        if info['word_available']:
            lines.append('  Microsoft Word  可用')
        elif info['word_hijacked']:
            lines.append('  Microsoft Word  未安装（Word 接口被 WPS 接管）')
        else:
            lines.append('  Microsoft Word  未检测到')
        if info['has_engine']:
            if info['wps_available'] and info['word_available']:
                order = 'WPS → Word'
            elif info['wps_available']:
                order = 'WPS Office'
            else:
                order = 'Microsoft Word'
            lines.append('  使用顺序：%s' % order)
        else:
            lines.append('  !! 没有可用引擎，无法处理文档')
    lines.append('=' * 62)
    lines.append('')
    try:
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')
    except Exception:
        pass


def cleanup_old_logs(keep=30):
    """只保留最近 N 次运行的日志，避免目录里越堆越多"""
    try:
        files = [os.path.join(LOGDIR, f) for f in os.listdir(LOGDIR)
                 if f.startswith('run_') and f.endswith('.log')]
        files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        for old in files[keep:]:
            try:
                os.remove(old)
            except Exception:
                pass
    except Exception:
        pass


# ============ 卡顿监控 ============
class Watchdog:
    """监控单个文件的处理时长。

    COM 调用一旦被模态对话框挡住就永不返回，代码里无法打断它（没地方插检查点）。
    所以这里只做一件事：超时后在日志里喊一声，让人知道该去那台电脑上关弹窗。
    线程是 daemon，COM 调用期间会释放 GIL，所以它能照常跑。
    """

    def __init__(self, ctx, limit=None):
        self.ctx = ctx
        self.limit = limit or int(os.environ.get('TOOLKIT_FILE_TIMEOUT', '90'))
        self.state = {'cur': '', 't0': 0.0, 'stop': False, 'warned': False}
        threading.Thread(target=self._loop, daemon=True).start()

    def begin(self, name):
        self.state['cur'] = name
        self.state['t0'] = time.time()
        self.state['warned'] = False
        with LOCK:
            STATE['stuck_file'] = ''
            STATE['stuck_seconds'] = 0
            STATE['stuck_since_ts'] = 0

    def stop(self):
        self.state['stop'] = True
        with LOCK:
            STATE['stuck_file'] = ''
            STATE['stuck_seconds'] = 0
            STATE['stuck_since_ts'] = 0

    def _loop(self):
        while not self.state['stop']:
            time.sleep(3)
            cur = self.state['cur']
            if not cur or self.state['warned'] or self.state['stop']:
                continue
            el = time.time() - self.state['t0']
            if el > self.limit:
                self.state['warned'] = True
                with LOCK:
                    STATE['stuck_file'] = cur
                    STATE['stuck_seconds'] = el
                    STATE['stuck_since_ts'] = self.state['t0'] + self.limit
                self.ctx.log(
                    '「%s」已等待 %.0f 秒没有响应。通常是办公软件弹了对话框'
                    '（文件损坏是否修复 / 是否更新链接 / 受保护的视图 / 文件正在使用），'
                    '看不到弹窗就点界面上的「中断并清理进程」，它会结束残留的办公软件进程、'
                    '清掉 ~$ 锁文件，任务随即结束。' % (cur, el), 'warn')


# ============ 工具函数 ============
def collect_files(root, recursive, exts, ctx=None):
    """收集目标文件。ctx 可选，用于把「读取目录失败」写进日志。"""
    out = []
    if recursive:
        for dp, dns, fns in os.walk(root):
            dns[:] = [d for d in dns if not d.startswith('_')]
            for fn in fns:
                if fn.startswith('~$'):
                    continue
                if fn.lower().endswith(exts):
                    out.append(os.path.normpath(os.path.join(dp, fn)))
    else:
        try:
            for fn in os.listdir(root):
                p = os.path.join(root, fn)
                if os.path.isfile(p) and not fn.startswith('~$') and fn.lower().endswith(exts):
                    out.append(os.path.normpath(p))
        except Exception as e:
            if ctx:
                ctx.log('读取目录失败 %s << %s' % (root, str(e)[:120]), 'error')
    if not out and ctx:
        ctx.log('没有找到 %s 文件' % '/'.join(exts), 'warn')
    return sorted(out)


def doc_path(path):
    """把路径整理成办公软件真正认的形式，再交给 COM。

    Word 的 Documents.Open 会把传进来的字符串当 URL 解析：
    路径里只要出现正斜杠（用户从别处粘贴的 C:/xxx/yyy 就是这种，
    配置里手填的目录也一样），空格就会被编码成 %20，然后报
    「很抱歉，找不到您的文件 …… 1%20管理手册.doc」。

    实测（Office 2021）：
      反斜杠绝对路径 + 中文 + 空格 + # + %   全部正常
      同一路径改成正斜杠                     必然失败
    所以这里统一转绝对路径 + 反斜杠。WPS 对反斜杠绝对路径同样正常。
    """
    try:
        p = os.path.abspath(path)
    except Exception:
        p = path
    if os.sep != '/':
        p = p.replace('/', '\\')
    return p


def mask_path(p):
    """日志里不打印完整绝对路径，避免暴露用户名等隐私。

    C:\\Users\\DJ\\AppData\\Local\\... → %LOCALAPPDATA%\\...
    C:\\Program Files\\Microsoft Office\\... → %ProgramFiles%\\Microsoft Office\\...
    信息量保留，隐私去掉。用户把日志发给别人排查时也安全。

    比较前统一成正斜杠——COM 返回的是反斜杠，但从配置粘贴进来的路径常是正斜杠，
    不统一就会漏过。
    """
    try:
        p = str(p)
        norm = p.replace('\\', '/')

        def _hit(base):
            b = (base or '').replace('\\', '/').rstrip('/')
            return b if b and b != '~' and norm.lower().startswith(b.lower() + '/') else None

        h = _hit(os.path.expanduser('~'))
        if h:
            return '%USERPROFILE%\\' + norm[len(h) + 1:].replace('/', '\\')
        for env in ('LOCALAPPDATA', 'APPDATA', 'ProgramFiles(x86)', 'ProgramFiles', 'TEMP'):
            b = _hit(os.environ.get(env))
            if b:
                return '%' + env + '%\\' + norm[len(b) + 1:].replace('/', '\\')
    except Exception:
        pass
    return p


# ============ 文档引擎（WPS / 微软 Office）============
# 背景：WPS 为了兼容，会把 Word.Application 这个 ProgID 抢注到自己的 wps.exe 上，
# 而且 app.Name 会伪装成 "Microsoft Word"、Version 伪装成 "12.0"。
# 所以判断实际拿到的是谁，只能靠 app.Path。
PROGIDS = {
    'wps': 'KWPS.Application',
    'word': 'Word.Application',
}


def _reg_exe(progid):
    """读注册表拿到某个 ProgID 实际指向的 exe 路径，未注册返回 None"""
    import winreg

    def _val(root, sub):
        try:
            with winreg.OpenKey(root, sub) as k:
                return winreg.QueryValueEx(k, '')[0]
        except Exception:
            return None

    clsid = _val(winreg.HKEY_CLASSES_ROOT, progid + r'\CLSID')
    if not clsid:
        return None
    exe = _val(winreg.HKEY_CLASSES_ROOT, r'CLSID\%s\LocalServer32' % clsid)
    return exe or '?'


def detect_engines():
    """注册表层面探测本机有哪些文档引擎可用（不启动任何程序）"""
    wps_exe = _reg_exe(PROGIDS['wps'])
    word_exe = _reg_exe(PROGIDS['word'])
    word_low = (word_exe or '').lower()
    # Word.Application 指向 wps.exe / kingsoft 目录 → 被 WPS 接管，不是真 Word
    word_hijacked = bool(word_exe) and ('wps.exe' in word_low or 'kingsoft' in word_low)
    real_word = bool(word_exe) and not word_hijacked

    labels = []
    if wps_exe:
        labels.append('WPS Office')
    if real_word:
        labels.append('Microsoft Word')
    if not labels and word_hijacked:
        labels.append('WPS Office（Word 接口被 WPS 接管）')

    return {
        'wps_available': bool(wps_exe),
        'word_available': real_word,
        'word_hijacked': word_hijacked,
        'wps_exe': wps_exe or '',
        'word_exe': word_exe or '',
        'summary': '、'.join(labels) if labels else '未检测到 WPS 或 Microsoft Word',
        # 只统计「明确可用的引擎」。word_hijacked 单独存在时（Word 接口被 WPS 接管、
        # 但 KWPS.Application 又没注册）open_office 里两个都会跳过，所以不能算作有引擎，
        # 否则会出现「检测说有引擎、实际一个都启动不了」的矛盾。
        'has_engine': bool(wps_exe) or real_word,
    }


def identify_engine(app):
    """判断当前 COM 实例到底是 WPS 还是微软 Word。返回 (名称, 路径)"""
    try:
        path = app.Path or ''
    except Exception:
        path = ''
    low = path.lower()
    if 'kingsoft' in low or 'wps' in low:
        return 'WPS Office', path
    if 'microsoft office' in low or 'office1' in low or 'winword' in low or 'root\\office' in low:
        return 'Microsoft Word', path
    # 兜底：WPS 的 Build 形如 12.1.0.29158
    try:
        build = app.Build or ''
    except Exception:
        build = ''
    return '未知引擎', path or ('Build %s' % build)


def open_office():
    """按「先 WPS、后微软 Word」的顺序打开文档引擎，返回 (app, 引擎名, 路径)

    跳过逻辑很重要：本机只有 WPS 时，Word.Application 被 WPS 接管，
    强行调用会启动一个 47/48 文件都打不开的坏实例，所以 word_available=False 时直接跳过。
    """
    import win32com.client as win32

    info = detect_engines()
    if not info['has_engine']:
        raise RuntimeError('本机既没有装 WPS Office，也没有装 Microsoft Word，无法处理文档。')

    errors = []
    for key in ('wps', 'word'):
        if key == 'wps' and not info['wps_available']:
            continue
        if key == 'word' and not info['word_available']:
            continue
        progid = PROGIDS[key]
        try:
            app = win32.DispatchEx(progid)
        except Exception as e:
            errors.append('%s 启动失败（%s）' % (progid, str(e)[:80]))
            continue
        tune_app(app)
        name, where = identify_engine(app)
        return app, name, where

    raise RuntimeError('文档引擎启动失败。检测结果：%s\n尝试记录：%s'
                       % (info['summary'], ' | '.join(errors) or '没有可尝试的引擎'))


def tune_app(app):
    """抑制办公软件的自动化干扰：宏、链接更新、拼写检查、自动保存。

    这些不关掉，打开某些文档时会弹模态框；模态框一弹，后续 COM 调用就全部阻塞。
    """
    for attr, val in (('Visible', False), ('DisplayAlerts', 0),
                      ('AutomationSecurity', 1),      # msoAutomationSecurityLow，避免强制禁宏引发的异常态
                      ('DisplayRecentFiles', False)):
        try:
            setattr(app, attr, val)
        except Exception:
            pass
    try:
        opts = app.Options
    except Exception:
        return
    for opt, val in (('UpdateLinksAtOpen', False),
                     ('SaveInterval', 0),
                     ('CheckSpellingAsYouType', False),
                     ('CheckGrammarAsYouType', False),
                     # Word 专有：退出时别问「是否保存 Normal 模板」
                     ('SaveNormalPrompt', False)):
        try:
            setattr(opts, opt, val)
        except Exception:
            pass


def open_document(app, path, readonly=True):
    """打开文档，把能关的对话框都关掉。

    「卡在某个文件不动」通常是打开时弹了框：
      文件损坏是否修复 / 是否更新链接 / 编码选择 / 受保护的视图 / 宏安全提示

    参数说明（都是 Word 真实支持的，已实测）：
      ConfirmConversions=False  不弹格式转换询问
      OpenAndRepair=False       不自动修复，避免弹「是否修复损坏文档」
      NoEncodingDialog=True     不弹编码选择
      Visible=False             不在界面上显示文档窗口

    注：网上有些资料建议加 DisableFeatures=True，实测 Word 的 Documents.Open
    没有这个参数（那是 Excel Workbooks.Open 的参数），传进来会直接抛
    TypeError: unexpected keyword argument。别加。
    """
    return app.Documents.Open(
        FileName=doc_path(path),
        ConfirmConversions=False,
        ReadOnly=readonly,
        AddToRecentFiles=False,
        Revert=False,
        Visible=False,
        OpenAndRepair=False,
        NoEncodingDialog=True,
    )


def md5_of(path, chunk=1 << 20):
    """返回文件指纹。读取失败返回 None。
    注意：不能返回空串——两个读不了且大小相同的文件会因此被判为「内容完全相同」，
    进而误导用户删除不该删的文件（检查报告 S3）。"""
    h = hashlib.md5()
    try:
        with open(path, 'rb') as f:
            while True:
                b = f.read(chunk)
                if not b:
                    break
                h.update(b)
        return h.hexdigest()
    except Exception:
        return None


def human_size(n):
    for u in ('B', 'KB', 'MB', 'GB'):
        if n < 1024:
            return '%.1f %s' % (n, u)
        n /= 1024.0
    return '%.1f TB' % n


def safe_rel(root, p):
    try:
        return os.path.relpath(p, root)
    except Exception:
        return p


# ============ 任务：doc -> docx ============
def task_convert(ctx, root, recursive):
    docs = collect_files(root, recursive, ('.doc',), ctx=ctx)
    ctx.log('找到 %d 个 .doc 文件待转换' % len(docs))
    ctx.progress(0, len(docs))

    ok = fail = skip = 0
    fails = []
    app, eng_name, eng_where = open_office()
    ctx.log('文档引擎：%s  (%s)' % (eng_name, mask_path(eng_where)), 'ok')
    wd = Watchdog(ctx)
    try:
        for i, p in enumerate(docs, 1):
            if ctx.check_cancel():
                ctx.log('用户取消', 'warn')
                break
            rel = safe_rel(root, p)
            ctx.progress(i, len(docs), rel)
            wd.begin(rel)
            target = os.path.splitext(p)[0] + '.docx'
            if os.path.exists(target):
                skip += 1
                ctx.log('跳过（docx 已存在）%s' % rel, 'warn')
                continue
            doc = None
            try:
                doc = open_document(app, p, readonly=True)
                doc.SaveAs(doc_path(target), WD_FORMAT_XML_DOCUMENT)
                ok += 1
                ctx.log('转换成功 %s' % rel, 'ok')
            except Exception as e:
                fail += 1
                fails.append({'file': rel, 'err': str(e)[:200]})
                ctx.log('转换失败 %s << %s' % (rel, str(e)[:120]), 'error')
            finally:
                if doc is not None:
                    try:
                        doc.Close(0)
                    except Exception:
                        pass
                    doc = None      # 及时释放 COM 引用，避免批量处理时引用累积
    finally:
        wd.stop()
        try:
            app.Quit()
        except Exception:
            pass

    ctx.log('转换汇总：扫描 %d 个，成功 %d，跳过 %d，失败 %d'
            % (len(docs), ok, skip, fail), 'warn' if fail else 'ok')
    if ctx.check_cancel():
        ctx.log('（任务被中断，以上为已处理部分）', 'warn')
    return {'ok': ok, 'fail': fail, 'skip': skip, 'total': len(docs), 'fails': fails}


# ============ 任务：清理排查 ============
def task_cleanup_scan(ctx, root, recursive, check_pair, check_dup, check_lock):
    checks = []
    if check_pair:
        checks.append('同名 doc/docx')
    if check_dup:
        checks.append('内容重复')
    if check_lock:
        checks.append('临时锁文件')
    ctx.log('开始排查，检查项：%s' % ('、'.join(checks) if checks else '（无）'))
    all_files = collect_files(root, recursive, ('.doc', '.docx'), ctx=ctx)
    ctx.progress(0, len(all_files))

    candidates = {}
    reasons = {}
    n_pair = n_dup = n_lock = 0
    n_md5_fail = 0

    # 1) 同名 doc + docx 并存
    if check_pair:
        docx_bases = set()
        for p in all_files:
            if p.lower().endswith('.docx'):
                docx_bases.add(os.path.splitext(p)[0].lower())
        for p in all_files:
            if p.lower().endswith('.doc'):
                if os.path.splitext(p)[0].lower() in docx_bases:
                    candidates[p] = True
                    reasons[p] = '已有同名 docx，此 doc 为转换前原件'
                    n_pair += 1
        ctx.log('  ① 同名 doc/docx 并存：%d 个' % n_pair)

    # 2) 内容完全重复
    if check_dup:
        ctx.log('  ② 计算文件指纹...')
        groups = {}
        for i, p in enumerate(all_files, 1):
            if ctx.check_cancel():
                break
            ctx.progress(i, len(all_files), safe_rel(root, p))
            try:
                size = os.path.getsize(p)
            except Exception as e:
                ctx.log('    读不到文件大小，跳过 %s << %s'
                        % (safe_rel(root, p), str(e)[:60]), 'warn')
                continue
            h = md5_of(p)
            if h is None:
                n_md5_fail += 1
                ctx.log('    指纹计算失败，跳过 %s' % safe_rel(root, p), 'warn')
                continue
            groups.setdefault((size, h), []).append(p)
        for key, ps in groups.items():
            if len(ps) < 2:
                continue
            # 保留优先级：docx > doc；路径层级浅 > 深
            ps_sorted = sorted(ps, key=lambda x: (0 if x.lower().endswith('.docx') else 1, x.count(os.sep), x))
            keep = ps_sorted[0]
            for p in ps_sorted[1:]:
                candidates[p] = True
                reasons[p] = '与「%s」内容完全相同' % os.path.basename(keep)
                n_dup += 1
        msg = '  ② 内容完全相同：%d 个' % n_dup
        if n_md5_fail:
            msg += '（另有 %d 个无法计算指纹，已排除）' % n_md5_fail
        ctx.log(msg, 'warn' if n_md5_fail else 'info')

    # 3) 临时锁文件
    if check_lock:
        if recursive:
            walker = list(os.walk(root))
        else:
            try:
                walker = [(root, [], os.listdir(root))]
            except Exception as e:
                ctx.log('  读取目录失败 %s << %s' % (root, str(e)[:80]), 'error')
                walker = []
        for dp, dns, fns in walker:
            for fn in fns:
                if fn.startswith('~$'):
                    p = os.path.join(dp, fn)
                    if os.path.isfile(p):
                        candidates[p] = True
                        # 不写死 WPS：微软 Word 打开文档时同样会产生 ~$ 锁文件
                        reasons[p] = '办公软件临时锁文件（WPS/Word 共用）'
                        n_lock += 1
        ctx.log('  ③ 临时锁文件：%d 个' % n_lock)

    items = []
    total_size = 0
    for p in sorted(candidates):
        try:
            sz = os.path.getsize(p)
        except Exception:
            sz = 0
        total_size += sz
        items.append({
            'path': p,
            'rel': safe_rel(root, p),
            'size': sz,
            'size_text': human_size(sz),
            'reason': reasons.get(p, ''),
        })

    ctx.log('排查完成：%d 个候选，共 %s' % (len(items), human_size(total_size)), 'ok')
    return {'items': items, 'count': len(items), 'total_size': total_size,
            'total_size_text': human_size(total_size), 'root': root}


# ============ 送回收站（原生 API） ============
def move_to_recycle_bin(paths):
    """把一批文件送进 Windows 回收站。
    注意：SHFileOperationW 的返回码不可靠（实测返回 2 但文件已删除），
    所以统一用「文件是否还存在」来判定结果。
    """
    import ctypes
    from ctypes import wintypes

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [('hwnd', wintypes.HWND), ('wFunc', wintypes.UINT),
                    ('pFrom', wintypes.LPCWSTR), ('pTo', wintypes.LPCWSTR),
                    ('fFlags', ctypes.c_uint16), ('fAnyOperationsAborted', wintypes.BOOL),
                    ('hNameMappings', ctypes.c_void_p), ('lpszProgressTitle', wintypes.LPCWSTR)]

    FO_DELETE = 3
    FOF_ALLOWUNDO = 0x0040
    FOF_NOCONFIRMATION = 0x0010
    FOF_SILENT = 0x0004
    FOF_NOERRORUI = 0x0400

    absolute = [os.path.abspath(p) for p in paths if os.path.exists(p)]
    if not absolute:
        return [], []

    op = SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = FO_DELETE
    op.pFrom = '\0'.join(absolute) + '\0\0'
    op.pTo = None
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
    op.fAnyOperationsAborted = False
    try:
        ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    except Exception:
        pass

    good, bad = [], []
    for p in absolute:
        if os.path.exists(p):
            bad.append(p)
        else:
            good.append(p)
    return good, bad


# ============ 任务：清理删除 ============
def task_cleanup_delete(ctx, root, paths):
    root_abs = os.path.abspath(root)
    ctx.log('准备删除 %d 个文件（送回收站）' % len(paths))
    ctx.progress(0, len(paths))

    ok = 0
    fails = []
    checked = []
    for p in paths:
        ap = os.path.abspath(p)
        if not ap.lower().startswith(root_abs.lower() + os.sep):
            fails.append({'file': p, 'err': '不在处理目录内，已拒绝'})
            ctx.log('拒绝（越界）%s' % p, 'error')
            continue
        if not os.path.exists(ap):
            fails.append({'file': p, 'err': '文件已不存在'})
            ctx.log('跳过（文件已不存在）%s' % safe_rel(root, ap), 'warn')
            continue
        checked.append(ap)

    # 每批 20 个，稳妥起见分批处理
    batch = 20
    for i in range(0, len(checked), batch):
        if ctx.check_cancel():
            ctx.log('用户取消', 'warn')
            break
        chunk = checked[i:i + batch]
        ctx.progress(min(i + len(chunk), len(paths)), len(paths), safe_rel(root, chunk[0]))
        good, bad = move_to_recycle_bin(chunk)
        for p in good:
            ok += 1
            ctx.log('已送回收站 %s' % safe_rel(root, p), 'ok')
        for p in bad:
            fails.append({'file': p, 'err': '删除后文件仍存在（可能被占用）'})
            ctx.log('删除失败 %s' % safe_rel(root, p), 'error')

    ctx.log('删除汇总：提交 %d 个，成功 %d，失败 %d'
            % (len(paths), ok, len(fails)), 'warn' if fails else 'ok')
    if ctx.check_cancel():
        ctx.log('（任务被中断，以上为已处理部分）', 'warn')
    return {'ok': ok, 'fail': len(fails), 'fails': fails, 'total': len(paths)}


# ============ 任务：查找 ============
def _count_findnext(doc, kw):
    """主路径：Execute 一次拿到首个命中，之后用无参 Execute() 续找（等价 FindNext）。

    【为什么不能每轮重置范围】——这是微软 Word 下「任务卡死」的真凶：
    Word 里对同一个 rng 反复 Execute(完整参数) 并且用 SetRange(rng.End, end)
    挪范围，会一次次命中同一处：Execute 永远返回 True，但 rng.End 原地不动。
    旧代码靠「n > 100000 就停」兜底，而一次 Execute 在 Word 里约 40 毫秒，
    10 万次就是一小时以上——界面上没有任何输出、进度条不动，看着就是卡死，
    同时 WINWORD.EXE 把 CPU 吃满（实测一个文件空转 363 秒 CPU）。

    改成 FindNext 式续找后，命中数与大文件实测都对得上：
      人员能力考核评价表.doc  5 处（旧写法空转到 2000 才被截断）
      1 管理手册.doc         420 处 / 6.7 秒（旧写法空转到 2000 才被截断）
    """
    rng = doc.Content
    end = rng.End
    if end <= 0:
        return 0
    rng.SetRange(0, end)
    find = rng.Find
    n = 0
    last_start = -1
    hit = find.Execute(kw, True, False, False, False, False,
                       True, WD_WRAP_STOP, False, '', WD_REPLACE_NONE)
    while hit:
        n += 1
        if n > 100000:
            break
        # 起点没前进说明引擎在原地打转，立刻收手，绝不让它空转
        if rng.Start == last_start:
            break
        last_start = rng.Start
        hit = find.Execute()
    return n


def _count_by_reset(doc, kw):
    """备用路径：每轮重置查找范围（WPS 上一直用的老写法）。

    保留它是为了防 WPS：WPS 的 Find 语义和 Word 不完全一样，万一某个版本
    上 FindNext 式续找只返回第一处，就用这条老路兜回来。
    和老代码的区别只有一处——加了停滞检测，发动机打转就退出，不再空转。
    """
    rng = doc.Content
    end = rng.End
    if end <= 0:
        return 0
    rng.SetRange(0, end)
    find = rng.Find
    n = 0
    stall = 0
    while True:
        hit = find.Execute(kw, True, False, False, False, False,
                           True, WD_WRAP_STOP, False, '', WD_REPLACE_NONE)
        if not hit:
            break
        n += 1
        if n > 100000:
            break
        if rng.End >= end:
            break
        prev = rng.End
        rng.SetRange(rng.End, end)
        if rng.End == prev:
            stall += 1
            if stall > 3:
                break
    return n


def count_in_doc(doc, kw):
    """统计命中数。出错返回 -1。
    不能返回 0——那会把「COM 出错」伪装成「没找到」，用户以为文档里没有这个词（检查报告 S13）。
    """
    try:
        n = _count_findnext(doc, kw)
    except Exception:
        n = -1
    if n > 0:
        return n
    # 主路径给 0 或报错时，用老写法复核一遍，取较大的那个。
    # WPS 与 Word 的 Find 语义有差异，两条路都留着才两边都不吃亏。
    try:
        m = _count_by_reset(doc, kw)
    except Exception:
        m = -1
    if n < 0 and m < 0:
        return -1
    return max(n, m, 0)


def task_scan(ctx, root, recursive, keywords, fuzzy_mode=False):
    kws = [k for k in keywords if k]
    files = collect_files(root, recursive, ('.doc', '.docx'))
    ctx.log('扫描 %d 个文件，%d 个关键词%s'
            % (len(files), len(kws), '  【模糊模式】' if fuzzy_mode else ''))
    ctx.progress(0, len(files))

    patterns = {}
    if fuzzy_mode:
        for k in kws:
            try:
                patterns[k] = fuzzy.FuzzyPattern(k)
            except Exception as e:
                ctx.log('关键词「%s」编译失败：%s' % (k, e), 'error')
        ctx.log('模糊规则：日期格式等价 + 字符间空格容错')

    kw_total = {k: 0 for k in kws}
    kw_files = {k: 0 for k in kws}
    hits = []
    hit_files = 0
    fail = 0

    app, eng_name, eng_where = open_office()
    ctx.log('文档引擎：%s  (%s)' % (eng_name, mask_path(eng_where)), 'ok')
    wd = Watchdog(ctx)
    try:
        for i, p in enumerate(files, 1):
            if ctx.check_cancel():
                ctx.log('用户取消', 'warn')
                break
            rel = safe_rel(root, p)
            ctx.progress(i, len(files), rel)
            wd.begin(rel)
            doc = None
            try:
                doc = open_document(app, p, readonly=True)
            except Exception as e:
                fail += 1
                ctx.log('打不开 %s << %s' % (rel, str(e)[:100]), 'error')
                continue
            file_has = False
            ftext = None
            if fuzzy_mode:
                try:
                    ftext = doc.Content.Text or ''
                except Exception as e:
                    ftext = ''
                    ctx.log('读取文本失败 %s << %s' % (rel, str(e)[:80]), 'error')
            for k in kws:
                if fuzzy_mode:
                    fp = patterns.get(k)
                    n = len(fp.finditer(ftext)) if fp else 0
                else:
                    n = count_in_doc(doc, k)
                    if n < 0:
                        ctx.log('统计出错 %s [%s]' % (rel, k), 'error')
                        n = 0
                if n > 0:
                    kw_total[k] += n
                    kw_files[k] += 1
                    hits.append({'file': rel, 'path': p, 'keyword': k, 'count': n})
                    file_has = True
            if file_has:
                hit_files += 1
                ctx.log('命中 %s' % rel, 'ok')
            try:
                doc.Close(0)
            except Exception:
                pass
            doc = None          # 及时释放 COM 引用，避免批量处理时引用累积
    finally:
        wd.stop()
        try:
            app.Quit()
        except Exception:
            pass

    total = sum(kw_total.values())
    ctx.log('')
    for k in kws:
        ctx.log('关键词「%s」：命中 %d 处，涉及 %d 个文件' % (k, kw_total[k], kw_files[k]))
    msg = '查找汇总：扫描 %d 个文件，命中 %d 个，共 %d 处' % (len(files), hit_files, total)
    if fail:
        msg += '，%d 个文件打不开' % fail
    ctx.log(msg, 'warn' if fail else ('ok' if total else 'warn'))
    if ctx.check_cancel():
        ctx.log('（任务被中断，以上为已处理部分）', 'warn')
    return {
        'files': len(files),
        'hit_files': hit_files,
        'total_hits': total,
        'keywords': [{'kw': k, 'total': kw_total[k], 'files': kw_files[k]} for k in kws],
        'hits': hits,
        'fail': fail,
        'can_replace': total > 0,
    }


# ============ 任务：替换 ============
def _bsearch_range_start(doc, target, end_limit):
    """二分出一个「大概的」Range 起点，仅作为精确搜索用的提示值。

    len(Range(0,k).Text) 随 k 单调不减，所以能二分。但它不是严格一一对应：
    有的字符占位不占文本、有的位置占文本不占位置，所以结果只能是近似值
    （实测稳定偏 1），调用方必须再拿原文比对确认。
    """
    lo, hi = 0, max(1, end_limit)
    while lo < hi:
        mid = (lo + hi) // 2
        try:
            n = len(doc.Range(0, mid).Text or '')
        except Exception:
            return None
        if n > target:
            hi = mid
        else:
            lo = mid + 1
    return lo


# 上下文探针里不要碰的字符：这些在文本里有、Range 位置里没有（或反过来），
# 拿它们当锚点会让比对永远不相等。
_BAD_CTX = set('\x07\r\x0b\x0c\x01\x13\x14\x15')


def _context_of(text, s, e, width=20):
    """取命中片段左右的「干净」上下文，用来确认定位到的就是这一处。

    只认这些字符做锚点，是因为同一个词（比如同一串日期）在一个文档里往往
    重复出现十几次，光比对「读到的是不是这个词」无法区分是哪一处，
    会误把前一处改成后一处。
    """
    left = ''
    i = s
    while i > 0 and len(left) < width and text[i - 1] not in _BAD_CTX:
        i -= 1
        left = text[i] + left
    right = ''
    j = e
    while j < len(text) and len(right) < width and text[j] not in _BAD_CTX:
        right += text[j]
        j += 1
    return left, right


def _verify_at(doc, seg, k, left, right):
    """把 Range(k-左, k+片长+右) 读回来，与「左+片段+右」逐字比对"""
    a = k - len(left)
    b = k + len(seg) + len(right)
    if a < 0:
        return False
    try:
        return (doc.Range(a, b).Text or '') == (left + seg + right)
    except Exception:
        return False


def _find_hit_pos(doc, text, s, e, pm, cache, used, end_all):
    """定位单个命中的真实 Range 起点。失败返回 None。

    fuzzy.build_pos_map() 假设「只有 \\x07 不占 Range 位置」，这个假设在 WPS 上
    是对的，在 Word 上不成立。实测：
        1 管理手册.doc          len(文本) 42759 / End 41354 / 差 1405，而 \\x07 有 1450 个
        合格供方名录及供方评定.doc  偏差最大到 129 个字符
    也就是说部分 \\x07 占位、另有一部分 Range 位置在文本里没有对应字符，
    纯按个数换算必然错位——实测替换后会把命中词前面的空格吞掉、末尾留半个字。

    所以改成「猜 + 读回验证」，验证还带上左右上下文，确保改的就是这一处。
    """
    seg = text[s:e]
    length = e - s
    if length <= 0:
        return None
    left, right = _context_of(text, s, e)
    guess = pm[s]
    limit = end_all - length

    cands = []
    d = cache.get('delta')
    if d is not None:
        cands.append(guess + d)
    cands.append(guess)
    for step in range(1, 61):
        cands.append(guess + step)
        cands.append(guess - step)

    for k in cands:
        if k < 0 or k > limit or k in used:
            continue
        if _verify_at(doc, seg, k, left, right):
            cache['delta'] = k - guess
            return k

    # 小范围没找到：二分拿个提示值，再在附近扩大搜索
    try:
        hint = _bsearch_range_start(doc, s, end_all)
    except Exception:
        hint = None
    if hint is not None:
        for step in range(-60, 61):
            k = hint + step
            if k < 0 or k > limit or k in used:
                continue
            if _verify_at(doc, seg, k, left, right):
                return k
    return None


def fuzzy_replace_doc(doc, patterns, repls):
    """模糊模式替换一个文档。返回 (每个关键词的命中数, 实际替换成功数, 替换失败数)

    要点：
    - 全文只取一次，所有关键词共享
    - 先把每一处命中定位成真实 Range 位置，再按位置倒序回写
      （倒序是为了替换后不影响前面尚未处理的位置）
    - 多个关键词的命中若重叠，只处理位置靠后的那个
    - 命中数 ≠ 替换成功数，必须分别统计（检查报告 S1）
    """
    counts = {k: 0 for k in patterns}
    try:
        text = doc.Content.Text or ''
    except Exception:
        return counts, 0, 0
    if not text:
        return counts, 0, 0

    pm = fuzzy.build_pos_map(text)
    try:
        base = doc.Content.Start
    except Exception:
        base = 0
    try:
        end_all = doc.Content.End
    except Exception:
        end_all = len(text)

    all_hits = []
    for kw, fp in patterns.items():
        try:
            hits = fp.finditer(text)
        except Exception:
            hits = []
        counts[kw] = len(hits)
        rep = repls.get(kw, '')
        if rep:
            for s, e in hits:
                all_hits.append((s, e, rep))

    # 第一步：把文本索引全部换成真实 Range 位置（此时文档还没动，可以放心搜索）
    cache = {}
    used = set()
    located = []
    for s, e, rep in sorted(all_hits, key=lambda x: x[0]):
        k = _find_hit_pos(doc, text, s, e, pm, cache, used, end_all)
        if k is None:
            located.append((None, s, e, rep))
        else:
            used.add(k)
            located.append((k, s, e, rep))

    # 第二步：倒序回写（定位失败的排最后，仍按文本位置排，避免漏计失败数）
    located.sort(key=lambda x: (x[0] is None, -(x[0] if x[0] is not None else x[1])))
    last_start = len(text) + 1
    replaced = failed = 0
    for k, s, e, rep in located:
        if e > last_start:
            continue
        if k is None:
            failed += 1
            last_start = s
            continue
        try:
            doc.Range(base + k, base + k + (e - s)).Text = rep
            replaced += 1
        except Exception:
            failed += 1
        last_start = s
    return counts, replaced, failed


def task_replace(ctx, root, recursive, pairs, backup, fuzzy_mode=False):
    ps = [(a, b) for a, b in pairs if a]
    if not any(b for _a, b in ps):
        ctx.log('所有规则的「替换为」都是空的，没有可执行的替换', 'warn')
        return {'files': 0, 'changed': 0, 'total_hits': 0, 'saved': 0, 'fail': 0,
                'fails': [], 'details': [], 'backup_dir': ''}
    files = collect_files(root, recursive, ('.doc', '.docx'), ctx=ctx)
    ctx.log('扫描 %d 个文件，%d 条规则%s'
            % (len(files), len(ps), '  【模糊模式】' if fuzzy_mode else ''))
    ctx.progress(0, len(files))

    patterns = {}
    if fuzzy_mode:
        for a, _b in ps:
            try:
                patterns[a] = fuzzy.FuzzyPattern(a)
            except Exception as e:
                ctx.log('规则「%s」编译失败：%s' % (a, e), 'error')
        ctx.log('模糊规则：日期格式等价 + 字符间空格容错；替换值原样写入')

    backup_dir = ''
    if backup:
        backup_dir = os.path.join(root, '_备份_' + time.strftime('%Y%m%d_%H%M%S'))
        os.makedirs(backup_dir, exist_ok=True)
        ctx.log('备份目录：%s' % mask_path(backup_dir))

    ok = changed = total_hits = failed = 0
    details = []
    fails = []

    app, eng_name, eng_where = open_office()
    ctx.log('文档引擎：%s  (%s)' % (eng_name, mask_path(eng_where)), 'ok')
    wd = Watchdog(ctx)
    try:
        for i, p in enumerate(files, 1):
            if ctx.check_cancel():
                ctx.log('用户取消', 'warn')
                break
            rel = safe_rel(root, p)
            ctx.progress(i, len(files), rel)
            wd.begin(rel)
            doc = None
            try:
                doc = open_document(app, p, readonly=False)
            except Exception as e:
                failed += 1
                fails.append({'file': rel, 'err': '打不开'})
                ctx.log('打不开 %s' % rel, 'error')
                continue

            fh = 0
            done_n = 0
            replace_failed = 0
            if fuzzy_mode:
                try:
                    cnts, done_n, replace_failed = fuzzy_replace_doc(doc, patterns, dict(ps))
                    fh = sum(cnts.values())
                    if replace_failed:
                        ctx.log('替换部分失败 %s：命中 %d 处，成功 %d，失败 %d'
                                % (rel, fh, done_n, replace_failed), 'error')
                except Exception as e:
                    replace_failed += 1
                    ctx.log('模糊替换出错 %s << %s' % (rel, str(e)[:80]), 'error')
            else:
                for old, new in ps:
                    n = count_in_doc(doc, old)
                    if n < 0:
                        ctx.log('统计出错 %s [%s]' % (rel, old), 'error')
                        continue
                    if n <= 0:
                        continue
                    fh += n
                    if new:
                        try:
                            rng = doc.Content
                            hit = rng.Find.Execute(old, True, False, False, False, False,
                                                   True, WD_WRAP_STOP, False, new, WD_REPLACE_ALL)
                            if hit:
                                done_n += n
                            else:
                                replace_failed += 1
                                ctx.log('替换未生效 %s [%s]' % (rel, old), 'warn')
                        except Exception as e:
                            replace_failed += 1
                            ctx.log('替换出错 %s [%s] << %s' % (rel, old, str(e)[:80]), 'error')

            if fh > 0:
                changed += 1
                total_hits += fh
                details.append('%5d 处  %s' % (fh, rel))
                skip_save = False
                # 备份失败就不做替换，宁可不改也不能丢了退路（检查报告 S2）
                if backup:
                    try:
                        dst = os.path.join(backup_dir, rel.replace('\\', '_').replace('/', '_'))
                        shutil.copy2(p, dst)
                    except Exception as be:
                        skip_save = True
                        failed += 1
                        fails.append({'file': rel, 'err': '备份失败，已跳过替换'})
                        ctx.log('备份失败 %s << %s  —— 该文件保持原样，未做替换'
                                % (rel, str(be)[:80]), 'error')
                if not skip_save:
                    try:
                        doc.Save()
                        ok += 1
                        ctx.log('已替换 %s  （命中 %d 处，写入 %d 处）'
                                % (rel, fh, done_n if done_n else fh), 'ok')
                    except Exception as e:
                        failed += 1
                        fails.append({'file': rel, 'err': '保存失败'})
                        ctx.log('保存失败 %s << %s  —— 命中 %d 处但没写进磁盘'
                                % (rel, str(e)[:80], fh), 'error')
            try:
                doc.Close(0)
            except Exception:
                pass
            doc = None          # 及时释放 COM 引用，避免批量处理时引用累积
    finally:
        wd.stop()
        try:
            app.Quit()
        except Exception:
            pass

    ctx.log('替换汇总：扫描 %d 个文件，命中 %d 个文件共 %d 处，写入成功 %d 个，失败 %d 个'
            % (len(files), changed, total_hits, ok, failed), 'warn' if failed else 'ok')
    if ctx.check_cancel():
        ctx.log('（任务被中断，以上为已处理部分）', 'warn')
    return {'files': len(files), 'changed': changed, 'total_hits': total_hits,
            'saved': ok, 'fail': failed, 'fails': fails,
            'details': details, 'backup_dir': backup_dir}


# ============ 任务调度 ============
def start_task(name, label, fn, **kwargs):
    with LOCK:
        if STATE['running']:
            return False, '已有任务在执行，请等待完成'
        # 注意：logs 和 seq 绝对不清零，否则前端增量游标会失效、新日志永远看不到
        STATE.update({
            'running': True, 'task': name, 'task_label': label,
            'total': 0, 'done': 0, 'current': '',
            'result': {}, 'cancel': False, 'started_at': time.time(),
            'finished_at': None,
        })

    def runner():
        import pythoncom
        pythoncom.CoInitialize()
        ctx = Ctx()
        ctx.log('')
        ctx.log('────────────────  开始：%s  ────────────────' % label, 'head')
        t0 = time.time()
        try:
            res = fn(ctx, **kwargs)
            with LOCK:
                STATE['result'] = res or {}
            ctx.log('────────────────  完成：%s，耗时 %.1f 秒  ────────────────'
                    % (label, time.time() - t0), 'head')
        except Exception as e:
            ctx.log('任务异常：%s' % str(e)[:300], 'error')
            # traceback 逐行输出：整段塞一条日志会被 HTML 折叠成一行，完全没法看（检查报告 S10）
            tb = traceback.format_exc().rstrip().splitlines()
            for line in tb[-40:]:
                ctx.log('    ' + line, 'error')
            if len(tb) > 40:
                ctx.log('    ...（堆栈过长，仅显示最后 40 行）', 'error')
            with LOCK:
                STATE['result'] = {'error': str(e)}
            ctx.log('────────────────  结束（异常）：%s，耗时 %.1f 秒  ────────────────'
                    % (label, time.time() - t0), 'head')
        finally:
            ctx.close()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            with LOCK:
                STATE['running'] = False
                STATE['current'] = ''
                STATE['cancel'] = False
                STATE['finished_at'] = time.time()

    threading.Thread(target=runner, daemon=True).start()
    return True, ''


# ============ API ============
app = FastAPI()
app.mount('/static', StaticFiles(directory=STATIC), name='static')


class DirReq(BaseModel):
    root: str
    recursive: bool = True


class CleanupReq(BaseModel):
    root: str
    recursive: bool = True
    check_pair: bool = True
    check_dup: bool = True
    check_lock: bool = True


class DeleteReq(BaseModel):
    root: str
    paths: List[str]


class ScanReq(BaseModel):
    root: str
    recursive: bool = True
    keywords: List[str] = []
    fuzzy: bool = False


class ReplaceReq(BaseModel):
    root: str
    recursive: bool = True
    pairs: List[List[str]] = []
    backup: bool = True
    fuzzy: bool = False


@app.get('/', response_class=HTMLResponse)
def index():
    p = os.path.join(STATIC, 'index.html')
    with open(p, encoding='utf-8') as f:
        return f.read()


@app.get('/api/engines')
def api_engines():
    """探测本机可用的文档引擎（只读注册表，不启动任何程序）。
    前端页面加载时调用，成功才放开操作按钮。"""
    try:
        return detect_engines()
    except Exception as e:
        return {'summary': '检测失败：%s' % e, 'has_engine': False,
                'wps_available': False, 'word_available': False,
                'word_hijacked': False, 'wps_exe': '', 'word_exe': ''}


@app.get('/api/list-dir')
def api_list_dir(path: str = ''):
    """网页内目录浏览。path 为空 -> 返回盘符列表。"""
    path = (path or '').strip()

    if not path:
        drives = []
        for letter in 'CDEFGHIJKLMNOPQRSTUVWXYZ':
            d = letter + ':\\'
            try:
                if os.path.exists(d):
                    drives.append({'name': d, 'path': d})
            except Exception:
                pass
        return {'path': '', 'parent': '', 'dirs': drives, 'is_root': True}

    if not os.path.isdir(path):
        return {'error': '目录不存在：%s' % path}

    dirs = []
    SKIP = {'$recycle.bin', 'system volume information', 'recovery', 'config.msi'}
    try:
        for name in sorted(os.listdir(path), key=lambda x: x.lower()):
            if name.startswith('$') or name.startswith('.'):
                continue
            if name.lower() in SKIP:
                continue
            p = os.path.join(path, name)
            try:
                if os.path.isdir(p):
                    dirs.append({'name': name, 'path': p})
            except Exception:
                pass
    except PermissionError:
        return {'error': '没有权限访问这个目录'}
    except Exception as e:
        return {'error': str(e)}

    norm = os.path.normpath(path)
    if len(norm) <= 3 and norm[1:2] == ':':
        parent = ''
    else:
        parent = os.path.dirname(norm)

    return {'path': norm, 'parent': parent, 'dirs': dirs, 'is_root': False}


@app.post('/api/convert')
def api_convert(req: DirReq):
    if not os.path.isdir(req.root):
        return JSONResponse({'ok': False, 'msg': '目录不存在'}, status_code=400)
    ok, msg = start_task('convert', 'doc → docx 转换', task_convert,
                         root=req.root, recursive=req.recursive)
    return {'ok': ok, 'msg': msg}


@app.post('/api/cleanup-scan')
def api_cleanup_scan(req: CleanupReq):
    if not os.path.isdir(req.root):
        return JSONResponse({'ok': False, 'msg': '目录不存在'}, status_code=400)
    ok, msg = start_task('cleanup-scan', 'doc 文件清理排查', task_cleanup_scan,
                         root=req.root, recursive=req.recursive,
                         check_pair=req.check_pair, check_dup=req.check_dup,
                         check_lock=req.check_lock)
    return {'ok': ok, 'msg': msg}


@app.post('/api/cleanup-delete')
def api_cleanup_delete(req: DeleteReq):
    if not os.path.isdir(req.root):
        return JSONResponse({'ok': False, 'msg': '目录不存在'}, status_code=400)
    if not req.paths:
        return JSONResponse({'ok': False, 'msg': '没有选中文件'}, status_code=400)
    ok, msg = start_task('cleanup-delete', '清理文件（送回收站）', task_cleanup_delete,
                         root=req.root, paths=req.paths)
    return {'ok': ok, 'msg': msg}


@app.post('/api/scan')
def api_scan(req: ScanReq):
    if not os.path.isdir(req.root):
        return JSONResponse({'ok': False, 'msg': '目录不存在'}, status_code=400)
    kws = [k.strip() for k in req.keywords if k and k.strip()]
    if not kws:
        return JSONResponse({'ok': False, 'msg': '请至少填写一个查找关键词'}, status_code=400)
    ok, msg = start_task('scan', '关键词查找', task_scan,
                         root=req.root, recursive=req.recursive, keywords=kws,
                         fuzzy_mode=req.fuzzy)
    return {'ok': ok, 'msg': msg}


@app.post('/api/replace')
def api_replace(req: ReplaceReq):
    if not os.path.isdir(req.root):
        return JSONResponse({'ok': False, 'msg': '目录不存在'}, status_code=400)
    pairs = []
    for item in req.pairs:
        if not isinstance(item, (list, tuple)) or len(item) < 2:
            continue
        a, b = str(item[0]).strip(), str(item[1]).strip()
        if a:
            pairs.append((a, b))
    if not pairs:
        return JSONResponse({'ok': False, 'msg': '没有可替换的规则'}, status_code=400)
    ok, msg = start_task('replace', '批量替换', task_replace,
                         root=req.root, recursive=req.recursive,
                         pairs=pairs, backup=req.backup,
                         fuzzy_mode=req.fuzzy)
    return {'ok': ok, 'msg': msg}


@app.post('/api/cancel')
def api_cancel():
    with LOCK:
        STATE['cancel'] = True
    return {'ok': True}


@app.get('/api/status')
def api_status(since: int = 0):
    with LOCK:
        logs = [l for l in STATE['logs'] if l['seq'] > since]
        # 首次拉取（页面刚打开）只给最近 400 条，避免一次返回几千条
        if since == 0 and len(logs) > 400:
            logs = logs[-400:]
        # 卡死秒数实时算，不依赖 Watchdog 的 3 秒节拍
        stuck_since = STATE.get('stuck_since_ts', 0)
        stuck = STATE['stuck_seconds']
        if STATE['running'] and stuck_since:
            stuck = max(stuck, round(time.time() - stuck_since, 1))
        return {
            'running': STATE['running'],
            'task': STATE['task'],
            'task_label': STATE['task_label'],
            'total': STATE['total'],
            'done': STATE['done'],
            'current': STATE['current'],
            'result': STATE['result'],
            'seq': STATE['seq'],
            'logs': logs,
            'cancel': STATE['cancel'],
            'logfile': STATE['logfile'],
            'elapsed': round(time.time() - STATE['started_at'], 1) if STATE['started_at'] and STATE['running'] else None,
            'stuck_file': STATE['stuck_file'],
            'stuck_seconds': stuck if STATE['running'] else 0,
            'stuck_limit': int(os.environ.get('TOOLKIT_FILE_TIMEOUT', '90')),
            'last_kill': STATE.get('last_kill', {}),
        }


# ---------- 卡死中断：结束残留的办公软件进程 + 清理锁文件 ----------
OFFICE_IMAGES = ('WINWORD.EXE', 'EXCEL.EXE', 'POWERPNT.EXE',
                 'wps.exe', 'et.exe', 'wpp.exe')




# 「用户正开着的文档窗口」的顶层框架类名。自动化实例是 Visible=False 起的，
# 不会产生这些窗口；它卡在模态框时只有一个 #32770 对话框，不在此列。
# 覆盖 Word(OpusApp/_WwG)、Excel(XLMAIN)、PowerPoint(PPTFrameClass)、WPS 系。
DOC_WINDOW_CLASSES = {'opusapp', '_wwg', 'xlmain', 'pptframeclass',
                      'kwpsmainframe', 'ketmainframe', 'kwppmainframe'}


def _visible_window_pids(only_doc=False):
    """返回「拥有可见顶层窗口」的 PID 集合；查不出来返回 None（调用方据此降级）。

    only_doc=True 时只认文档主框架窗口（= 用户真的在看的文档）。

    为什么要区分：只看「有没有窗口」不够——自动化实例卡在模态对话框时也有一个
    可见窗口（类名 #32770），按「有窗口就跳过」会把真正卡死的那个放过去（对端实测踩过）。
    """
    import ctypes
    from ctypes import wintypes
    try:
        user32 = ctypes.windll.user32
    except Exception:
        return None
    found = set()
    cb_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def _cb(hwnd, _lparam):
        try:
            if not user32.IsWindowVisible(hwnd):
                return True
            if only_doc:
                buf = ctypes.create_unicode_buffer(256)
                if not user32.GetClassNameW(hwnd, buf, 256):
                    return True
                if buf.value.lower() not in DOC_WINDOW_CLASSES:
                    return True
            pid = wintypes.DWORD(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value:
                found.add(int(pid.value))
        except Exception:
            pass
        return True

    try:
        if not user32.EnumWindows(cb_type(_cb), 0):
            # EnumWindows 返回 0 通常是回调抛异常；此时「没找到窗口」不可信
            return None
    except Exception:
        return None
    return found


def _list_office_pids():
    """用 tasklist 拿名称+PID，不依赖 psutil（打包进 exe 时依赖越少越好）"""
    import subprocess
    out = {}
    try:
        txt = subprocess.run(['tasklist', '/FO', 'CSV', '/NH'],
                             capture_output=True, text=True, timeout=20,
                             encoding='utf-8', errors='replace').stdout
    except Exception:
        return out
    for line in txt.splitlines():
        parts = [x.strip('"') for x in line.split('","')]
        if len(parts) < 2:
            continue
        name, pid = parts[0].strip('"'), parts[1].strip('"')
        if name.upper() in [i.upper() for i in OFFICE_IMAGES]:
            out.setdefault(name, []).append(pid)
    return out


def _proc_cmdline(pid):
    """拿进程命令行。psutil 装了就精确识别，没装返回空串（退回看窗口的判据）。"""
    try:
        import psutil
        return ' '.join(psutil.Process(int(pid)).cmdline())
    except Exception:
        return ''


def _is_automation(pid):
    """是不是工具/脚本启动的自动化实例。True / False / **None=判断不了**。

    实测：COM 起的 Word 命令行一定是 `WINWORD.EXE /Automation -Embedding`，
    用户双击打开的是 `WINWORD.EXE "D:\\xxx.docx"`。
    比「有没有窗口」可靠——卡在模态对话框的实例也会有一个可见窗口，
    只看窗口会把它当成用户正在用的文档而跳过（实测踩过这个坑）。

    注意三态：拿不到命令行（psutil 没装/权限不足/进程已退出）时返回 None，
    不能当成 False——那会让调用方误以为「这是用户正在用的文档」而不敢杀，
    卡死的任务就永远解不开。调用方拿到 None 会降级用窗口类名判断。
    """
    low = _proc_cmdline(pid).lower()
    if not low:
        return None
    return '/automation' in low or '-embedding' in low


def _kill_pid(pid):
    import signal
    try:
        os.kill(int(pid), signal.SIGTERM)
        return True
    except Exception:
        return False


def _clean_locks(root, log):
    """删掉 root 下的 ~$ 锁文件。只有在确认没有任何办公软件进程后才调用。

    两道保守限制：不要整个盘（`D:\\`）当根、不要根目录当根——
    那会遍历整块盘，代价极高，也超出「清理本次任务目录」的本意。
    """
    removed = []
    if not root or not os.path.isdir(root):
        return removed
    drive, tail = os.path.splitdrive(os.path.abspath(root))
    if tail.strip('\\/') == '':
        log('清理范围是整个盘（%s），范围过大，已跳过锁文件清理' % root, 'warn')
        return removed
    for dp, _dns, fns in os.walk(root):
        for fn in fns:
            if fn.startswith('~$'):
                p = os.path.join(dp, fn)
                try:
                    os.remove(p)
                    removed.append(p)
                except Exception as e:
                    log('锁文件删不掉 %s << %s' % (safe_rel(root, p), str(e)[:60]), 'warn')
    return removed


def _push_log(msg, level='warn'):
    """不依赖任务上下文往日志里写一行（任务线程可能正被 COM 卡住）"""
    line = {'t': now(), 'level': level, 'msg': str(msg), 'seq': 0}
    with LOCK:
        STATE['seq'] += 1
        line['seq'] = STATE['seq']
        STATE['logs'].append(line)
    fh = None
    try:
        fh = open(STATE.get('logfile') or LOG_FILE, 'a', encoding='utf-8')
        fh.write('[%s][%s] %s\n' % (line['t'], level, msg))
        fh.flush()
    except Exception:
        pass
    finally:
        if fh:
            try:
                fh.close()
            except Exception:
                pass


class KillReq(BaseModel):
    root: str = ''
    force: bool = False


@app.post('/api/kill-stuck')
def api_kill_stuck(req: KillReq):
    """卡死中断：结束办公软件的自动化实例，清掉 ~$ 锁文件。

    杀掉进程后，被卡住的 COM 调用会立即返回错误，任务线程随之退出，
    不需要等任何超时。

    默认只结束「命令行带 /Automation -Embedding」的实例（工具自己起的）；
    用户手工打开文档的实例（命令行带文档路径）保留。force=True 时连它一起结束。

    判定是三级的，优先级从高到低：
      ① 命令行能看出是不是自动化实例（最可靠，psutil 在就有）
      ② 命令行拿不到时，看有没有「文档主框架窗口」——有=你在用，没有=后台实例
      ③ 连窗口都查不出来时，退回「有没有任何可见窗口」，且只杀无窗口的（最保守）
    """
    with LOCK:
        STATE['cancel'] = True
    _push_log('收到「中断并清理进程」请求%s，开始清理…'
              % ('（强制模式）' if req.force else ''), 'warn')

    doc_pids = _visible_window_pids(only_doc=True)
    any_pids = _visible_window_pids(only_doc=False)
    before = _list_office_pids()
    killed, skipped = [], []

    def _decide(pid):
        """返回 (是否结束, 原因)。宁可少杀，不能误杀用户没保存的文档。"""
        if req.force:
            return True, '强制模式'
        auto = _is_automation(pid)
        if auto is True:
            return True, '命令行含 /Automation，是工具自己起的实例'
        # auto 为 False（命令行明确是用户开的）或 None（拿不到命令行）：都要看窗口
        if doc_pids is not None:
            if pid in doc_pids:
                return False, '有可见文档窗口，像是你正在用的，没动它；确实要结束请用强制模式'
            return True, '没有文档窗口，判定为后台实例'
        if any_pids is not None:
            if pid in any_pids:
                return False, '有可见窗口，可能是你正在用的，没动它；确实要结束请用强制模式'
            return True, '没有任何可见窗口，判定为后台实例'
        return False, '判断不出用途（拿不到命令行也查不到窗口），为安全没动它'

    for name, pids in before.items():
        for pid in pids:
            pid = int(pid)
            do_kill, why = _decide(pid)
            if not do_kill:
                skipped.append({'name': name, 'pid': pid, 'why': why})
                continue
            if _kill_pid(pid):
                killed.append({'name': name, 'pid': pid, 'why': why})
            else:
                skipped.append({'name': name, 'pid': pid,
                                'why': '结束失败（可能已自行退出）'})

    for k in killed:
        _push_log('已结束 %s (PID %d) —— %s' % (k['name'], k['pid'], k['why']), 'error')
    for s in skipped:
        _push_log('跳过 %s (PID %d)：%s' % (s['name'], s['pid'], s['why']), 'warn')

    # 等进程真正退出
    for _ in range(20):
        if not _list_office_pids():
            break
        time.sleep(0.5)

    locks = []
    still = _list_office_pids()
    if still:
        _push_log('仍有办公软件进程在运行（%s），为安全起见不删 ~$ 锁文件'
                  % '、'.join('%s×%d' % (n, len(p)) for n, p in still.items()), 'warn')
    else:
        locks = _clean_locks(req.root, _push_log)
        if locks:
            _push_log('已清理 %d 个残留锁文件（~$）' % len(locks), 'warn')

    if not killed and not locks:
        _push_log('没有需要清理的进程或锁文件', 'warn')

    res = {'ok': True, 'killed': killed, 'skipped': skipped,
           'locks_removed': len(locks), 'still_running': still}
    with LOCK:
        STATE['last_kill'] = {'t': now(), 'killed': len(killed),
                              'locks': len(locks), 'skipped': len(skipped)}
    _push_log('清理完成：结束进程 %d 个，跳过 %d 个，删锁文件 %d 个'
              % (len(killed), len(skipped), len(locks)), 'warn')
    return res


@app.get('/api/download-log')
def download_log():
    lf = STATE.get('logfile', '')
    if lf and os.path.exists(lf):
        return FileResponse(lf, filename=os.path.basename(lf))
    return JSONResponse({'ok': False, 'msg': '没有日志文件'}, status_code=404)


def find_free_port(start=8770, tries=80):
    """从 start 开始找一个空闲端口。
    换一台机器 8770 可能被别的程序占了，写死端口会导致起不来。
    环境变量 TOOLKIT_PORT 可以强制指定。"""
    import socket
    env = os.environ.get('TOOLKIT_PORT')
    if env:
        try:
            return int(env)
        except ValueError:
            pass
    for p in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('127.0.0.1', p))
                return p
            except OSError:
                continue
    return start


if __name__ == '__main__':
    import uvicorn
    import webbrowser

    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

    port = find_free_port()

    # 启动就把日志文件建好（含引擎检测），不用等第一个任务才开始记录
    write_run_header()
    cleanup_old_logs()
    _info = detect_engines()

    print('=' * 58)
    print('  文档批量处理工具')
    print('  地址: http://127.0.0.1:%d' % port)
    print('  日志目录: %s' % mask_path(LOGDIR))
    print('  本次日志: %s' % os.path.basename(LOG_FILE))
    print('-' * 58)
    print('  文档引擎检测：')
    print('    WPS Office      %s' % ('可用' if _info['wps_available'] else '未检测到'))
    if _info['word_available']:
        print('    Microsoft Word  可用')
    elif _info['word_hijacked']:
        print('    Microsoft Word  未安装（Word 接口被 WPS 接管）')
    else:
        print('    Microsoft Word  未检测到')
    if _info['has_engine']:
        if _info['wps_available'] and _info['word_available']:
            order_text = 'WPS → Word'
        elif _info['wps_available']:
            order_text = 'WPS Office'
        else:
            order_text = 'Microsoft Word'
        print('    使用顺序：%s' % order_text)
    else:
        print('    !! 没有可用引擎，程序无法处理文档')
    print('=' * 58)
    print('  关闭这个窗口即可停止服务')
    print('=' * 58)

    if os.environ.get('NO_BROWSER') != '1':
        threading.Timer(1.2, lambda: webbrowser.open('http://127.0.0.1:%d' % port)).start()
    uvicorn.run(app, host='127.0.0.1', port=port, log_level='warning')
