"""方案AI 出表强制自检（2026-10-26 加；同日按独立审查修了 4 个 bug）。

为什么要它：技能里写着「出表前必跑 kb_tool check、有 ❌ 不能交付」，但模型可能不跑、或看到 ❌ 不改。
这个插件不靠模型自觉：
  transform_tool_result  每次工具调用后，找这一步新生成/改过的 xlsx，跑 check；有 ❌ 就把问题放进工具结果，
                         模型下一步一定看得到，要求改完重新生成
  transform_llm_output   这一轮结束时，本轮出过的表还有 ❌ → 在回复末尾标出来（内部业务平台里业务/方案部直接看到）
插件自己出错只记日志，不影响 agent 干活（Hermes 对插件钩子 fail-open）。
工作区：环境变量 FYYS_WORKSPACE，没设就用进程当前目录（systemd 的 WorkingDirectory）。

多人同时用（线上）：工作区是共用的，按修改时间扫到的表可能是别人的。只有一个会话在干活时全收；
几个会话同时在干活时，只认「这次工具调用的参数/结果里出现过文件名」的表。
"""
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
from pathlib import Path

log = logging.getLogger("plugins.fyys-selfcheck")
WS = Path(os.environ.get("FYYS_WORKSPACE") or os.getcwd())
KB_TOOL = WS / "scripts" / "kb_tool.py"
SKIP = ("/网页上传/", "/references/", "/exam/", "/.git/", "/~$")   # 用户上传的原件、资料库、考卷不查
MAX_SHOW = 12
ACTIVE_TTL = 3600            # 一轮超过 1 小时还没收工，就不算「正在干活」了

_lock = threading.Lock()
_last = {"sid": ""}          # 并行工具调用时 Hermes 不传 session_id（2026-10-26 考卷实测）
_active = {}                 # session_id → 这一轮开始时间（收工时删掉）
_sess = {}                   # session_id → {"since", "room", "files": {路径: 结果}, "seen": {路径: mtime}}
_checked = {}                # 路径 → (查的时间, 结果)：收工时按回复里提到的文件名再认一遍


def _state(sid):
    with _lock:
        return _sess.setdefault(sid or "", {"since": time.time(), "room": False, "files": {}, "seen": {}})


def _n_active():
    now = time.time()
    return sum(1 for t in _active.values() if now - t < ACTIVE_TTL)


def _new_xlsx(since):
    out = []
    for root, dirs, files in os.walk(WS):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d not in ("网页上传", "references", "exam", "node_modules")]
        for f in files:
            if f.endswith(".xlsx") and not f.startswith("~$"):
                p = os.path.join(root, f)
                try:
                    if os.path.getmtime(p) > since and not any(s in p for s in SKIP):
                        out.append(p)
                except OSError:
                    pass
    return out


def _core(name):
    """文件名去掉日期前缀和扩展名：「2026-10-26_示例项目_比单表.xlsx」→「示例项目_比单表」"""
    return re.sub(r"^\d{4}-?\d{2}-?\d{2}[_-]?", "", os.path.basename(name)).rsplit(".", 1)[0]


def _mentioned(p, text):
    c = _core(p)
    return bool(c) and (os.path.basename(p) in text or c[:12] in text)


def _check(path, room):
    """→ None（不是配单表）/ (❌数, ⚠数, ❌行, ⚠行) / ("ERR", 说明)"""
    cmd = [sys.executable, str(KB_TOOL), "check", path] + (["--会议室"] if room else [])
    r = subprocess.run(cmd, cwd=WS, capture_output=True, text=True, timeout=120)
    lines = r.stdout.splitlines()
    head = next((l for l in lines if l.startswith("自检")), "")
    # kb_tool 崩了（表没写完/坏了、资料库正在重写…）不能当「通过」——以前会显示「通过（无 ❌ / ⚠）」
    if r.returncode not in (0, 1) or not head or "Traceback" in r.stderr:
        return ("ERR", (r.stderr.strip().splitlines() or ["没有输出"])[-1][:200])
    if re.search(r"：0 行", head):            # 不是比单/配单表（没有型号列），不管
        return None
    bad = [l.strip() for l in lines if l.strip().startswith("❌")]
    warn = [l.strip() for l in lines if l.strip().startswith("⚠")]
    return len(bad), len(warn), bad, warn


def _note(rel, r):
    if r[0] == "ERR":
        return f"【系统自动自检】{rel}：⚠ 自检没跑成（{r[1]}）——这张表可能没写完或损坏，请重新生成后再交付"
    nb, nw, bad, warn = r
    if nb:
        return (f"【系统自动自检】{rel}：❌ {nb} 处、⚠ {nw} 处。有 ❌ 的表不能交付——逐条改掉后重新生成这张表"
                f"（系统会再自动查一遍），改不了的行留空并写「我司无对应产品」：\n" + "\n".join(bad[:MAX_SHOW])
                + (f"\n…另有 {nb - MAX_SHOW} 处" if nb > MAX_SHOW else "")
                + ("\n⚠ 项：要么改，要么在备注写理由：\n" + "\n".join(warn[:6]) if warn else ""))
    if nw:
        return f"【系统自动自检】{rel}：无 ❌，⚠ {nw} 处（要么改，要么在备注写理由）：\n" + "\n".join(warn[:MAX_SHOW])
    return f"【系统自动自检】{rel}：通过（无 ❌ / ⚠）"


def on_turn_start(session_id="", user_message="", **kw):
    _last["sid"] = session_id or _last["sid"]
    _active[session_id or ""] = time.time()
    st = _state(session_id)
    msg = str(user_message or "")
    st["since"] = time.time() - 2
    st["files"] = {}
    st["seen"] = {}
    st["t0"] = time.time()
    st["room"] = bool(re.search(r"会议室", msg) and re.search(r"\d+\s*(平|㎡|m2|平方)", msg))
    return None


def after_tool(tool_name="", result="", session_id="", **kw):
    if not KB_TOOL.exists():
        return None
    multi = _n_active() > 1
    sid = session_id or ("" if multi else _last["sid"])     # 几个会话同时干活时不乱猜是谁的
    st = _state(sid)
    # 文件时间戳用的是粗粒度时钟，可能比 time.time() 早几毫秒：往前多看 2 秒，再按修改时间去重
    t = time.time() - 2
    since = st["since"] if sid else time.time() - 120
    files = [p for p in _new_xlsx(since) if st["seen"].get(p) != os.path.getmtime(p)]
    st["since"] = t
    if multi:                                    # 只认这次工具调用里提到过的表
        ctx = f"{kw.get('args') or kw.get('arguments') or ''} {result}"
        files = [p for p in files if _mentioned(p, ctx)]
    for p in files:
        st["seen"][p] = os.path.getmtime(p)
    if not files:
        return None
    notes = []
    for p in files:
        try:
            r = _check(p, st["room"])
        except Exception as e:                  # 超时等：同样不能当通过
            r = ("ERR", f"{type(e).__name__}: {e}"[:200])
        if r is None:
            continue
        st["files"][p] = r
        _checked[p] = (time.time(), r)
        notes.append(_note(os.path.relpath(p, WS), r))
    if not notes:
        return None
    log.info("自检 %s：%s", sid, " | ".join(n.splitlines()[0] for n in notes))
    text = "\n\n".join(notes)
    try:                                         # 工具结果是 JSON 的，放进字段里（拼在后面 Hermes 就认不出命令失败了）
        obj = json.loads(result)
        if isinstance(obj, dict):
            obj["fyys_selfcheck"] = text
            return json.dumps(obj, ensure_ascii=False)
    except (TypeError, ValueError):
        pass
    return str(result) + "\n\n" + text


def _family(p):
    """同一张表的不同版本：「比单表.xlsx」「比单表_v2.xlsx」「比单表_修改.xlsx」算一家"""
    return os.path.join(os.path.dirname(p), re.sub(r"([_\-\s]*(v\d+|V\d+|修改\d*|改\d*|最终|final|新|\(\d+\)|（\d+）|\d))+$", "",
                                                   _core(p)))


def on_final(response_text="", session_id="", **kw):
    sid = session_id or ("" if _n_active() > 1 else _last["sid"])
    st = _state(sid)
    _active.pop(sid, None)
    files = dict(st["files"])
    t0 = st.get("t0", 0)
    for p, (t, r) in list(_checked.items()):     # 并行调用时没记到本会话、但回复里提到了的表
        if t >= t0 and p not in files and _mentioned(p, response_text):
            files[p] = r
    if not files:
        return None
    live = {p: v for p, v in files.items() if os.path.exists(p)}
    newest = {}                                  # 同一张表只看最新版本（改完另存 _v2 的，旧版不报）
    for p in live:
        f = _family(p)
        if os.path.getmtime(p) > newest.get(f, (0, ""))[0]:
            newest[f] = (os.path.getmtime(p), p)
    keep = {p: v for p, v in live.items() if newest[_family(p)][1] == p}
    left = {p: v for p, v in keep.items() if v[0] == "ERR" or v[0]}
    if not left:
        return response_text.rstrip() + "\n\n（系统自检：本次生成的表格无 ❌）"
    parts = []
    for p, v in left.items():
        if v[0] == "ERR":
            parts.append(f"- {os.path.basename(p)}：自检没跑成（{v[1]}），表可能没写完或损坏")
        else:
            parts.append(f"- {os.path.basename(p)}：❌ {v[0]} 处\n" + "\n".join("  " + b for b in v[2][:MAX_SHOW]))
    return (response_text.rstrip() + "\n\n---\n⛔ **系统自检未通过，以下问题还没改，这张表请人工核对后再用：**\n" + "\n".join(parts))


def register(ctx):
    ctx.register_hook("pre_llm_call", on_turn_start)
    ctx.register_hook("transform_tool_result", after_tool)
    ctx.register_hook("transform_llm_output", on_final)
