#!/usr/bin/env python3
"""比单三步走（2026-10-27 加）：大清单不再一行一行让模型搜，程序先打底稿，模型只判断拿不准的行，出表交给程序。

为什么：上次 21 行的单子模型调了 48 次、14 分钟；四五百行的工程量清单要几十分钟。时间都花在「每行都让模型搜一遍、
上下文越积越长」和「模型现写代码读表写表」上。

  python scripts/bidan.py draft 网页上传/客户清单.xlsx              # ① 打底稿：几秒出结果，只打印需要你判断的行
  python scripts/bidan.py pick 方案表格/xxx_底稿.json 12=HT-XXX 30=空:理由 31=HT-YYY:备注   # ② 你只决定拿不准的行（可一次多行，型号后可带备注）
  ⚠ 方括号里的 [12] 是底稿编号，不是客户原表行号（客户行号在后面括号里写着）
  python scripts/bidan.py show 方案表格/xxx_底稿.json [行号]        # 看某行的候选/依据；不带行号=还没决定的行
  python scripts/bidan.py export 方案表格/xxx_底稿.json            # ③ 出比单表（客户原表右侧加我司列）+ 自动自检

底稿里每行的「把握」：
  不做   我们不做的系统（LED屏/监控/门禁/土建/电脑…）→ 留空，备注写清品类
  高     工具第1名参数达标，且方案部经验明确（≥3 个项目、≥一半配它）或明显领先 → 直接填好（你觉得不对也可以 pick 改）
  待定   拿不准 → 必须由你 pick（看客户参数、候选、经验后决定；资料库没有合适的就 pick 空:理由）
  无候选 资料库里找不到同类在售 → 你确认是否真的没有（pick 空:理由 或 pick 型号）
"""
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kb_tool as K  # noqa: E402

HOST = re.compile(r"会议.{0,4}主机|手拉手.{0,4}主机|会议系统主机")
UNIT = re.compile(r"主席|代表|会议单元|发言单元|语音终端|(会议|发言)话筒(?!处理)")
# 会议系统配套：主机定了，这些也从同一系列里配（2026-10-28 655 题：无线收发器配成了手持话筒、扩展交互盒漏配）
KIT_ACC = [(re.compile(r"收发器|发射器|(?<![A-Z])AP(?![A-Z])|接入点"), re.compile(r"AP发射器|发射器")),
           (re.compile(r"扩展(主机|交互盒|盒)|交互盒"), re.compile(r"扩展主机")),       # 先找扩展主机（655 方案部配 HY200ME）
           (re.compile(r"扩展(主机|交互盒|盒)|交互盒"), re.compile(r"扩展盒")),         # 没有再找话筒扩展盒
           (re.compile(r"充电箱"), re.compile(r"充电箱"))]


def _brief(p):
    return f"{p.get('model')}「{p.get('device_name')}」{p.get('price') or '未定价'}元"


def _cands(r, kb, top=3):
    name = K.name_head(r["name"]) or r["name"]
    c, note = K.rank(name, r["params"][:600], kb=kb, top=top, section=r["section"], brand=r["brand"])
    return name, [p for _, p in c], [s for s, _ in c], note


def decide(r, kb):
    name, cands, scores, note = _cands(r, kb)
    sysn, ours, why = K.system_of(r["name"], r["params"], r["section"], r["brand"])
    d = {"id": 0, "sheet": r["sheet"], "row": r["row"], "name": name, "section": r["section"], "params": r["params"][:300],
         "system": sysn, "cands": [p["model"] for p in cands], "pick": None, "level": "", "why": "", "note": ""}
    if not ours:
        d.update(level="不做", why=f"{sysn}（{why}）", note=f"我司无对应产品（{sysn}，我司不做）")
        return d
    if sysn == "辅材" or re.search(r"网线|跳线|底盒|配管|穿线|线管|光缆|光纤|尾纤|配线架|理线器|机柜|交换机|水晶头|电源线|信号线|线材|"
                                    r"套管|软管|桥架|线槽|测试$|熔接|电缆", name):
        d.update(level="辅材", why="辅材/布线", note="辅材按需另配", cands=[p["model"] for p in cands])
        return d
    if not cands:
        d.update(level="无候选", why=note or "资料库里没有同类在售")
        return d
    top = cands[0]
    gaps = K.param_gaps(name, r["params"], top)
    exp, nproj = K.experience(name)
    share = (exp[0][1] / nproj) if exp and nproj else 0
    exp_ok = exp and exp[0][0] == top.get("model") and nproj >= 3 and share >= 0.5
    lead = len(scores) < 2 or scores[0] - scores[1] >= 1.5
    case = K.same_case(name, r["params"])
    case_ok = case and case[1] >= 0.7 and case[0] == top.get("model")
    if not gaps and (case_ok or exp_ok or (lead and sysn)):
        d.update(level="高", pick=top["model"],
                 why=(f"历史同样清单方案部配它（参数相似度 {case[1]:.2f}，{case[2]} 个项目）" if case_ok else
                      f"方案部 {exp[0][1]}/{nproj} 个项目配它" if exp_ok else "参数达标、明显领先"))
    else:
        d.update(level="待定", why="；".join(g[1] for g in gaps) or
                 (f"经验：{exp[0][0]}（{exp[0][1]}/{nproj} 项目）" if exp else "候选分不出高下"))
    return d


def kit_fix(rows, kb):
    """会议主机定了，主席/代表单元从同系列里挑（2026-10-26：U段主机配了5G单元）"""
    idx = K.index(kb)
    for sheet in {d["sheet"] for d in rows}:
        hosts = [d for d in rows if d["sheet"] == sheet and d["pick"] and HOST.search(d["name"])]
        if not hosts:
            continue
        hp = K.find_model(idx, hosts[0]["pick"])
        ser = K.series_of(hp[0]) if hp else ""
        if not ser:
            continue
        kit = [p for p in kb if K.avail(p) and K.series_of(p) == ser and not K.is_old(p.get("model"))]
        for d in rows:                           # 配套（收发器/扩展盒/充电箱）：同系列里名字对得上的，挑最便宜的
            if d["sheet"] != sheet or d["level"] == "不做" or d.get("decided") or HOST.search(d["name"]) \
                    or UNIT.search(d["name"]) or re.search(r"天线|话筒(?!处理)", d["name"]):
                continue
            hit = []
            for nx, rx in KIT_ACC:               # 按优先顺序找，找到就停
                if nx.search(d["name"]):
                    hit = sorted((p for p in kit if rx.search(str(p.get("device_name")))), key=lambda p: float(p.get("price") or 1e9))
                    if hit:
                        break
            if hit:
                d.update(pick=hit[0]["model"], level="高", why=f"和主机 {hosts[0]['pick']} 同系列配套（{ser}）",
                         cands=[hit[0]["model"]] + [m for m in d["cands"] if m != hit[0]["model"]][:2])
        for d in rows:
            if d["sheet"] != sheet or d["level"] == "不做" or not UNIT.search(d["name"]) or d.get("decided"):
                continue
            same = [p for _, p in K.rank(d["name"], d["params"], kb=kb, top=30)[0] if K.series_of(p) == ser]
            if not same:                         # 名字太泛（「智慧语音终端」）搜不到：直接在同系列话筒/单元里挑，最便宜的在前
                same = sorted((p for p in kit if re.search(r"话筒|单元", str(p.get("device_name"))) and not re.search(r"处理器|扩展", str(p.get("device_name")))),
                              key=lambda p: float(p.get("price") or 1e9))
            if same:
                # 名字里没写主席/代表的（「智慧语音终端」6 只），按代表单元配
                want = "主席" if "主席" in d["name"] else "代表"
                best = next((p for p in same if want and want in str(p.get("device_name"))), same[0])
                d.update(pick=best["model"], level="高", why=f"和主机 {hosts[0]['pick']} 同系列（{ser}）",
                         cands=[best["model"]] + [m for m in d["cands"] if m != best["model"]][:2])


def panel_fix(rows, kb):
    """紧跟中控主机的「控制面板/触摸面板」配主机同系列的触控面板：86 底盒/4 寸 → 4 寸，写了 7/10 寸按寸数。
    2026-10-29 方案部：DCS-A180 下一行配了音频处理器系列的 W86C 墙板，不配套。跟在音频处理器后面的 W86C 是对的（983），不动"""
    idx = K.index(kb)
    for a, d in zip(rows, rows[1:]):
        if a["sheet"] != d["sheet"] or not a["pick"] or d.get("decided") or d["level"] == "不做" \
                or not re.search(r"面板|触控屏|触摸屏|控制屏", d["name"]) or re.search(r"音量|灯光|调光", d["name"]):
            continue
        hp = K.find_model(idx, a["pick"])
        if not hp or not K.series_of(hp[0]).startswith("中控矩阵分布式") or "中控主机" not in str(hp[0].get("device_name")):
            continue
        ser = K.series_of(hp[0])
        tp = [p for p in kb if K.avail(p) and K.series_of(p) == ser and "面板" in str(p.get("device_name"))]
        if not tp:
            continue
        n = K._name_n(f"{d['name']} {d['params']}", r"寸|英寸")
        size = lambda p: K._name_n(str(p.get("device_name")), r"寸|英寸") or 0   # noqa: E731
        best = min(tp, key=lambda p: (abs(size(p) - (n or 4)), float(p.get("price") or 1e9)))
        d.update(pick=best["model"], level="高", why=f"和上一行中控主机 {a['pick']} 同系列的触控面板（{ser}）",
                 cands=[best["model"]] + [m for m in d["cands"] if m != best["model"]][:2])


IFACE_CARD = [("HDMI", "HD"), ("DVI", "DVI"), ("SDI", "SDI"), ("VGA", "VGA"), ("HDBaseT", "HDBT")]


def add_extras(doc, kb):
    """增加项（客户清单没列、但我司配套必须有的）：插卡式矩阵主机 → 输入卡 + 输出卡。
    2026-10-27 方案部给示例项目配 HT-PRA01-WF-HD8*8 时另加了 HD-4IN、HD-4OUT 两行，agent 没加。"""
    idx = K.index(kb)
    extras = doc.setdefault("extras", [])
    for d in doc["rows"]:
        if not d.get("pick"):
            continue
        hp = K.find_model(idx, d["pick"])
        if not hp or K.series_of(hp[0]) not in K.KIT_EXTRA or "矩阵主机" not in str(hp[0].get("device_name")):
            continue
        same_sheet = [x for x in doc["rows"] if x["sheet"] == d["sheet"]]
        cards = [p for p in K.load_kb() if K.avail(p) and K.series_of(p) in K.KIT_EXTRA[K.series_of(hp[0])]]
        want = next((k for name, k in IFACE_CARD if re.search(name, d["params"], re.I)), "HD")
        m = re.search(r"(\d+)\s*进\s*(\d+)\s*出", f"{hp[0].get('device_name')} {d['params']}")
        for kind, n in (("输入卡", int(m.group(1)) if m else 4), ("输出卡", int(m.group(2)) if m else 4)):
            if any(kind in x["name"] for x in same_sheet) or any(e["for"] == d["id"] and kind in e["name"] for e in extras):
                continue                                  # 客户清单里已经有板卡行 / 已加过
            pick = [p for p in cards if kind in str(p.get("device_name")) and f"-{want}-" in p["model"]] or \
                   [p for p in cards if kind in str(p.get("device_name"))]
            if not pick:
                continue
            per = int((re.search(r"(\d+)路", str(pick[0].get("device_name"))) or [0, 4])[1])
            extras.append({"for": d["id"], "sheet": d["sheet"], "name": f"（增加项）{kind}", "model": pick[0]["model"],
                           "qty": -(-n // per), "note": f"配套 {d['pick']}（插卡式主机必须配板卡）；{n} 路 ÷ 每卡 {per} 路，按实际信号路数核对"})
            print(f"  + 增加项：[{d['id']}] {d['name'][:12]} 配 {pick[0]['model']} × {-(-n // per)}（{kind}）")


ANALOG_SRC = re.compile(r"播放器|调谐器|收音|CD|DVD|前置放大|广播话筒|会议话筒|寻呼话筒(?!.*IP)|调音台|音源|卡座")


def add_encoder(doc, kb):
    """IP 网络广播里有模拟音源 → 整张清单要配一台模拟转网络的编码设备 HT-PAE01-1013TW（客户没列就加增加项）。
    2026-10-27 方案部：「若有模拟音频信号，整个清单需配置此设备，别人清单没有配的话，需要我们另外增加」"""
    idx = K.index(kb)
    picked = [(d, K.find_model(idx, d["pick"])[0]) for d in doc["rows"] if d.get("pick") and K.find_model(idx, d["pick"])]
    ip = [d for d, p in picked if K.is_ip_broadcast(p)]                 # 网络广播系统（纯模拟广播不需要，方案部 10-27）
    analog = [d for d, p in picked if K.is_analog_source(p)]             # 配了 8017D/8118 这类模拟音频设备
    has = any(p["model"] == "HT-PAE01-1013TW" or re.search(r"采集|编码", str(p.get("device_name"))) for _, p in picked) or \
        any(d["level"] != "不做" and re.search(r"音频采集|音频编码|广播.{0,6}编码|编码器|采集器", d["name"]) and not re.search(r"能耗|数据|视频", d["name"])
            for d in doc["rows"]) or any("1013TW" in e["model"] for e in doc.get("extras", []))   # 「能耗数据采集器」不算
    if ip and analog and not has and K.find_model(idx, "HT-PAE01-1013TW"):
        a = analog[0]
        doc.setdefault("extras", []).append({"for": a["id"], "sheet": a["sheet"], "name": "（增加项）音频采集器（模拟音频转网络）",
                                              "model": "HT-PAE01-1013TW", "qty": 1,
                                              "note": f"IP 网络广播里有模拟音源（{'、'.join(d['name'][:8] for d in analog[:3])}），需配模拟转网络编码设备"})
        print(f"  + 增加项：音频采集器 HT-PAE01-1013TW × 1（有模拟音源：{'、'.join(d['name'][:8] for d in analog[:3])}）")


def cmd_draft(path):
    kb = K.load_kb()
    rows = []
    for r in K.read_rows(path):
        if not (K.name_head(r["name"]) or "").strip():
            continue
        d = decide(r, kb)
        d["id"] = len(rows) + 1
        rows.append(d)
    kit_fix(rows, kb)
    panel_fix(rows, kb)
    doc = {"src": str(Path(path).resolve()), "rows": rows, "extras": []}
    add_extras(doc, kb)
    add_encoder(doc, kb)
    out = K.ROOT / "方案表格" / f"{Path(path).stem}_底稿.json"
    out.parent.mkdir(exist_ok=True)
    json.dump(doc, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    from collections import Counter
    c = Counter(d["level"] for d in rows)
    print(f"底稿 {out}（{len(rows)} 行：高 {c['高']}、待定 {c['待定']}、无候选 {c['无候选']}、辅材 {c['辅材']}、不做 {c['不做']}）")
    print(f"  不做 {c['不做']} 行、辅材 {c['辅材']} 行已默认留空（备注写了原因）；要改用 pick，看明细用 show 行号")
    eq = [d for d in rows if d["level"] not in ("不做", "辅材")]
    nop = [d for d in eq if len(re.sub(r"详见.{0,12}(参数|附表|表)|见附件|配套|同一品牌|\s", "", d.get("params") or "")) < 8]
    if eq and len(nop) >= len(eq) * 0.5:          # 一半以上设备行没参数：功率/4K/路数都没法核（2026-10-28 会议室 9.18 清单）
        print(f"  ⚠ 客户清单 {len(nop)}/{len(eq)} 行设备没有技术参数（空着或写「详见技术参数表」）→ "
              "回复第一句就提醒用户「请把技术参数表也传上来，下面先按设备名称配，4K/功率/路数可能不准」")
    _print_todo(rows, kb, limit=40)
    print(f"→ 用 pick 决定上面的行（如 `python scripts/bidan.py pick {out.relative_to(K.ROOT)} 12=HT-XXX 30=空:理由`），"
          f"高把握的行觉得不对也可以改；全部定完再 export。")


def _specs(p, n=6):
    """候选的主要参数（技术参数前几项），模型不用再自己写代码去资料库里挖"""
    key = re.compile(r"功率|路|通道|分辨率|像素|变焦|尺寸|寸|接口|灵敏度|输入|输出|频率|阻抗|声压")
    sp = sorted(p.get("technical_specs") or [], key=lambda x: not key.search(f"{x.get('key')}{x.get('value')}"))   # 要紧的排前面
    def _one(x):
        k = x.get("key")
        v = re.sub(r"\s+", " ", str(x.get("value")))[:40]
        return f"{k}:{v}"
    return "；".join(_one(x) for x in sp[:n])


def _keypts(p):
    """从资料库全文里提炼要和客户参数比的数：功放各阻抗单通道功率、额定功率、4K、路数、尺寸、光学变焦"""
    t, dn = K._ptext(p), str(p.get("device_name") or "")
    out = []
    aw = K._amp_w(t)
    if aw:
        out.append("、".join(f"{o}Ω 每通道 {w:g}W" for o, w in sorted(aw.items(), reverse=True)))
    rw = K._name_w(dn) or K._rated_w(t)
    if rw and not aw:
        out.append(f"额定 {rw:g}W")
    if re.search(r"摄像|终端|矩阵|编码|解码|录播|屏", dn):
        out.append("4K" if K.is_4k(t) and "2K" not in K.series_of(p) else ("1080P" if re.search("1080", t) else ""))
    for unit, lab in (("路|通道", "路"), ("寸|英寸", "寸")):
        v = K._name_n(dn, unit)
        if v:
            out.append(f"{v:g}{lab}")
    z = re.search(r"光学变焦[^0-9\n]{0,6}(\d{1,2})\s*倍|(\d{1,2})\s*倍\s*光学", t)
    if z:
        out.append(f"{z.group(1) or z.group(2)}倍光学变焦")
    return "，".join(x for x in out if x) or "（资料库里没写可比的数）"


def _print_todo(rows, kb, only=None, limit=None):
    idx = K.index(kb)
    todo = [d for d in rows if (only and d["id"] in only) or (not only and d["level"] in ("待定", "无候选") and not d.get("decided"))]
    if not todo:
        print("（没有需要判断的行）")
    if limit and len(todo) > limit:
        print(f"（需要判断的有 {len(todo)} 行，先列前 {limit} 行；决定完再用 show 看剩下的）")
        todo = todo[:limit]
    for d in todo:
        print(f"[{d['id']}]（客户表 {d['sheet']} 第{d['row']}行）{d['section'][:10] or '—'}｜{d['name'][:20]}｜{d['level']}：{d['why'][:60]}")
        if d["params"]:
            _cp = re.sub(r"\s+", " ", d["params"])[:600 if only else 120]
            print(f"     客户参数：{_cp}")
        for m in d["cands"][:3]:
            ps = K.find_model(idx, m)
            if ps:
                g = K.param_gaps(d["name"], d["params"], ps[0])
                print(f"     - {_brief(ps[0])}｜{K.series_of(ps[0])}" + (f"｜⚠{g[0][1]}" if g else ""))
                print(f"         要点：{_keypts(ps[0])}")          # 打底稿时就给要点：模型不用自己写代码去查
                if only:
                    print(f"         参数：{_specs(ps[0])}")


def _load(f):
    return json.load(open(f, encoding="utf-8"))


def cmd_pick(f, args):
    doc = _load(f)
    kb = K.load_kb()
    idx = K.index(kb)
    by = {d["id"]: d for d in doc["rows"]}
    for a in args:
        mx = re.match(r"\+(\d+)=([^*×]+)(?:[*×](\d+))?(?::(.*))?$", a)
        if mx and int(mx.group(1)) in by:              # 增加项：+行号=型号*数量:说明
            ps = K.find_model(idx, mx.group(2).strip())
            if not ps or not K.avail(ps[0]) or K.is_old(ps[0].get("model")):
                print(f"✗ 增加项 {mx.group(2)}：资料库没有/停单/旧系列，没加")
                continue
            d = by[int(mx.group(1))]
            doc.setdefault("extras", []).append({"for": d["id"], "sheet": d["sheet"], "name": f"（增加项）{ps[0]['device_name']}",
                                                  "model": ps[0]["model"], "qty": int(mx.group(3) or 1), "note": mx.group(4) or f"配套 [{d['id']}] {d['name']}"})
            print(f"+ 增加项：跟在 [{d['id']}] {d['name'][:12]} 后面 → {_brief(ps[0])} × {mx.group(3) or 1}")
            continue
        m = re.match(r"(\d+)=(.+)", a)
        if not m or int(m.group(1)) not in by:
            print(f"✗ 看不懂：{a}（格式：行号=型号 或 行号=空:理由）")
            continue
        d, v = by[int(m.group(1))], m.group(2).strip()
        if v.startswith("空"):
            d.update(pick=None, decided=True, note=v.split(":", 1)[1].strip() if ":" in v else "我司无对应产品")
            print(f"[{d['id']}] {d['name'][:16]} → 留空（{d['note']}）")
            continue
        note = ""
        if ":" in v or "：" in v:                     # 12=HT-XXX:备注（agent 反馈：选了型号也要能写备注）
            v, note = re.split(r"[:：]", v, 1)
            v, note = v.strip(), note.strip()
        ps = K.find_model(idx, v)
        if not ps:
            print(f"✗ [{d['id']}] {v}：资料库里没有，没改")
            continue
        p = ps[0]
        probs = []
        if not K.avail(p):
            probs.append(f"停单（{p.get('status')}）")
        if K.is_old(p.get("model")):
            probs.append("旧系列禁用")
        probs += [g[1] for g in K.param_gaps(d["name"], d["params"], p) if g[0] == "❌"]
        if probs:
            print(f"✗ [{d['id']}] {v}：{'；'.join(probs)}——没改，换一个")
            continue
        d.update(pick=p["model"], decided=True, **({"note": note} if note else {}))
        print(f"[{d['id']}] {d['name'][:16]} → {_brief(p)}" + (f"（备注：{note}）" if note else ""))
    before = {d["id"]: d["pick"] for d in doc["rows"]}
    kit_fix(doc["rows"], kb)                       # 主机定了：没决定的主席/代表单元自动配同系列
    panel_fix(doc["rows"], kb)
    add_extras(doc, kb)                            # 矩阵主机定了：自动加输入/输出卡（增加项）
    add_encoder(doc, kb)                           # IP 广播有模拟音源：自动加模拟转网络编码设备（增加项）
    for d in doc["rows"]:
        if d["pick"] != before[d["id"]]:
            d["decided"] = True
            print(f"[{d['id']}] {d['name'][:16]} → {d['pick']}（自动：{d['why']}）")
    json.dump(doc, open(f, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    left = [d for d in doc["rows"] if d["level"] in ("待定", "无候选") and not d.get("decided")]
    print(f"还有 {len(left)} 行没决定" + (f"：{[d['id'] for d in left][:30]}" if left else "，可以 export 了"))


def cmd_show(f, ids):
    doc = _load(f)
    _print_todo(doc["rows"], K.load_kb(), only=set(map(int, ids)) if ids else None)


def cmd_export(f):
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    doc = _load(f)
    kb = K.load_kb()
    idx = K.index(kb)
    left = [d["id"] for d in doc["rows"] if d["level"] in ("待定", "无候选") and not d.get("decided")]
    if left:
        print(f"✗ 还有 {len(left)} 行没决定：{left[:30]}——先 pick 完再 export")
        return 1
    wb = openpyxl.load_workbook(doc["src"])           # 直接在客户原表上加列：客户的行和列一行不少
    by = {}
    for d in doc["rows"]:
        by.setdefault(d["sheet"], {})[d["row"]] = d
    hdr = ["", "品牌", "我司型号", "我司参数", "备注", "不满足参数"]
    blue = PatternFill("solid", fgColor="DDEBF7")
    for ws in wb.worksheets:
        rows = by.get(ws.title)
        if not rows:
            continue
        c0 = ws.max_column + 1
        # 表头行 = 设备行上面、表头词最多的那一行（不能取「最近的非空行」——那常是「一、音频系统」分组行）
        hw = re.compile(r"名称|品名|货物|设备|产品|型号|数量|单位|参数|规格|品牌|单价|金额|备注|序号|特征|描述")
        cand = [(sum(1 for j in range(1, c0) if ws.cell(k, j).value and len(str(ws.cell(k, j).value)) < 16
                     and hw.search(str(ws.cell(k, j).value))), -k, k) for k in range(1, min(rows))]
        best = max(cand, default=(0, 0, 0))
        hrow = best[2] if best[0] >= 2 else max(1, min(rows) - 1)
        for j, h in enumerate(hdr):
            if h:
                c = ws.cell(hrow, c0 + j, h)
                c.font = Font(bold=True)
                c.fill = blue
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for rr, d in rows.items():
            if d["pick"]:
                p = K.find_model(idx, d["pick"])[0]
                full = "\n".join(x.get("full_text", "") for x in (p.get("full_official_specs") or []) if x.get("full_text"))
                gaps = "；".join(g[1] for g in K.param_gaps(d["name"], d["params"], p))
                note = d.get("note") or ""
                if "京东" in str(p.get("price_source") or ""):  # 外购辅材：型号带空格/民用品牌，交付要说清楚（agent 反馈）
                    note = ("外购辅材（京东参考价，非我司产品）；" + note).rstrip("；")
                vals = ["示例供应商", p["model"], full or p.get("device_name"), note, gaps]
            else:
                vals = ["", "", "", d.get("note") or "我司无对应产品", ""]
            for j, v in enumerate(vals, start=1):
                c = ws.cell(rr, c0 + j, v)
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ex = [e for e in doc.get("extras", []) if e["sheet"] == ws.title]
        if ex:                                         # 增加项：客户清单没列、我司配套必须有的，写在这张表最后
            ncol = next((j for j in range(1, c0) if re.search(r"名称|品名|货物|设备", str(ws.cell(hrow, j).value or ""))), 2)
            qcol = next((j for j in range(1, c0) if "数量" in str(ws.cell(hrow, j).value or "")), None)
            r0 = ws.max_row + 2
            t = ws.cell(r0, ncol, "增加项（客户清单未列，我司配套必须）")
            t.font = Font(bold=True)
            for k, e in enumerate(ex, start=1):
                p = K.find_model(idx, e["model"])[0]
                ws.cell(r0 + k, ncol, e["name"])
                if qcol:
                    ws.cell(r0 + k, qcol, e["qty"])
                full = "\n".join(x.get("full_text", "") for x in (p.get("full_official_specs") or []) if x.get("full_text"))
                for j, v in enumerate(["示例供应商", p["model"], full or p.get("device_name"), e["note"], ""], start=1):
                    ws.cell(r0 + k, c0 + j, v).alignment = Alignment(wrap_text=True, vertical="top")
        for j, w in zip(range(c0, c0 + 6), (2, 10, 22, 60, 24, 24)):
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    out = Path(f).with_name(Path(f).name.replace("_底稿.json", "_比单表.xlsx"))
    wb.save(out)
    print(f"比单表：{out}")
    iss, _ = K.cmd_check(str(out))
    return 1 if any(x[0] == "❌" for x in iss) else 0


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    cmd, arg = sys.argv[1], sys.argv[2]
    if cmd == "draft":
        cmd_draft(arg)
    elif cmd == "pick":
        cmd_pick(arg, sys.argv[3:])
    elif cmd == "show":
        cmd_show(arg, sys.argv[3:])
    elif cmd == "export":
        sys.exit(cmd_export(arg))
    else:
        print(__doc__)
        sys.exit(1)
