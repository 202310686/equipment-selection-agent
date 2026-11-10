#!/usr/bin/env python3
"""生成「方案部经验库」references/experience.json（2026-10-26 加）。

来源：内部业务平台「方案流转」里方案部人工做的成果表（stage=result、非系统配价副本），只读拷贝。
每一行「客户设备名称 → 方案部配的我司型号」记一条经验；kb_tool search/plan 查到类似的行会提示「方案部以前配过 X（N 个项目）」。
规则：
  - 只收资料库里现在能配的型号（停单、旧系列、查不到的不收，免得学到过时习惯）
  - 同一个需求单里同名同型号只算 1 票（一张大单子不能把经验带偏）
  - 排除名单（--exclude 需求单id，逗号分隔）：考卷题目必须排除，不然等于提前给了答案；
    另外和排除名单里答案内容高度重合（≥60% 的行相同）的单子也排除（同一项目换了需求单号）
用法：python scripts/exp_build.py 文件清单.tsv 文件根目录 [--exclude 281,429,...] [--exclude-answers 答案目录] [--out 输出.json]
  每晚由 内部同步任务 调用（先写临时文件，检查条数没少太多再替换）
  文件清单.tsv 每行：需求单id|文件id|文件名|相对路径|大小|日期
"""
import glob
import json
import os
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import kb_tool as K  # noqa: E402

HT = re.compile(r"HT-[A-Z0-9][A-Z0-9\-\.（）()/]{1,40}", re.I)


def head(name):  # 和 kb_tool.name_head 同口径
    return K.name_head(name)


def pairs(path, idx):
    out = []
    for r in K.read_rows(path):
        m = HT.search(r["model"])
        if not m:
            continue
        ps = K.find_model(idx, m.group().rstrip("（("))
        if not ps or not K.avail(ps[0]) or K.is_old(ps[0].get("model")):
            continue
        n = head(r["name"])
        if len(n) >= 2:
            out.append((n, ps[0]["model"], re.sub(r"\s+", " ", r["params"])[:120], r["section"][:20]))
    return out


def main():
    a = sys.argv[1:]
    lst, base = a[0], a[1]
    excl = set(filter(None, a[a.index("--exclude") + 1].split(","))) if "--exclude" in a else set()   # 空串不能当 id
    idx = K.index(K.load_kb())
    banned = []                                   # 排除名单答案里的 (名称, 型号) 集合，用来查同项目换单号
    if "--exclude-answers" in a:
        for f in glob.glob(os.path.join(a[a.index("--exclude-answers") + 1], "**", "*.xlsx"), recursive=True):
            try:
                banned.append({(n, m) for n, m, _, _ in pairs(f, idx)})
            except Exception:
                pass
    rows, skipped = [], {"排除名单": 0, "和考卷重合": 0, "读不了": 0, "没有我司型号": 0}
    for line in open(lst, encoding="utf-8"):
        rid, fid, name, rel, size, date = line.rstrip("\n").split("|")[:6]
        if rid in excl:
            skipped["排除名单"] += 1
            continue
        try:
            ps = pairs(os.path.join(base, rel), idx)
        except Exception:
            skipped["读不了"] += 1
            continue
        if not ps:
            skipped["没有我司型号"] += 1
            continue
        s = {(n, m) for n, m, _, _ in ps}
        # 两个方向都看：大单子把考题答案整份包进去（占考题 60%）也要排除；考题只有一两行时不按它排除别人
        if any(b and (len(s & b) / len(s) >= 0.6 or (len(b) >= 5 and len(s & b) / len(b) >= 0.6)) for b in banned):
            skipped["和考卷重合"] += 1
            continue
        seen = set()
        for n, m, p, sec in ps:
            if (n, m) in seen:
                continue
            seen.add((n, m))
            rows.append({"name": n, "model": m, "params": p, "section": sec, "rid": int(rid), "date": date})
    out = Path(a[a.index("--out") + 1]) if "--out" in a else K.ROOT / "references" / "experience.json"
    json.dump({"说明": "方案部经验库：方案流转人工成果里「客户设备名称→我司型号」，每晚重建；kb_tool search/plan 会引用",
               "生成": __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M"), "排除需求单": sorted(excl, key=int),
               "rows": rows}, open(out, "w", encoding="utf-8"), ensure_ascii=False)
    print(f"经验库 {out}：{len(rows)} 条，来自 {len({r['rid'] for r in rows})} 个需求单；跳过 {skipped}")


if __name__ == "__main__":
    main()
