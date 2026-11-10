#!/usr/bin/env python3
"""设备选型工具：找候选型号（search）+ 出表前自检（check）。2026-10-26 加。

为什么要它：历史纠正里最多的错——编型号、配错类型、漏配、功放功率不够、配了停单——
都出在「模型凭记忆/速查表直接写型号、出表前没有任何检查」。这个工具把能算的交给程序：
  search  在资料库里按设备名+参数找在售候选，模型只在候选里挑，不许凭空写型号
  check   出表后逐行检查，有 ❌ 必须改完再交付，⚠ 要在备注里说明

用法（在工作区根目录下）：
  python scripts/kb_tool.py search "监听音箱" "一体化壁挂式，内置网络音频解码，额定功率≥10W"   # 名称和参数分开传
  python scripts/kb_tool.py plan 网页上传/客户清单.xlsx              # 比单第一步：逐行判系统（我们做不做）+ 候选
  python scripts/kb_tool.py search "拼接控制器" --分组 "三、显示系统" --品牌 诺瓦   # 带上分组标题和客户品牌，先判系统再找
  python scripts/kb_tool.py search "IP网络解码终端" --all          # 连停单的也列出（只用于说明，不能配）
  python scripts/kb_tool.py model HT-PAD01-1009LTW                   # 查一个型号的完整信息
  python scripts/kb_tool.py kit HT-MIU01-MSMX50                      # 主机的同系列配套清单（会议单元/扩展/AP 只能从这里选）
  python scripts/kb_tool.py check 方案表格/xxx.xlsx                  # 出表前必跑
  python scripts/kb_tool.py check 方案表格/xxx.xlsx --会议室          # 按面积出的会议室方案：再核对固定项
资料库：references/category_products_full_with_internal_model.json（每晚从内部业务平台产品浏览同步，
available=false 的是停单/已下架，不能配）。
"""
import json
import math
import os
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent          # 工作区根目录（不写死绝对路径）
PRICE_W = float(os.environ.get("KB_PRICE_W", "0.2"))   # 价格权重：2026-10-28 扫 0~0.8，0.2 命中率最高（33.2%/48.3%）
KB_PATH = ROOT / "references" / "category_products_full_with_internal_model.json"

# 设备大类关键词：客户写的类别和我们配的类别对不上，多半是配错类型
TYPE_WORDS = ["消防", "前置", "采集", "功放", "音箱", "音柱", "喇叭", "话筒", "寻呼", "主机", "服务器", "终端", "采集器", "播放器",
              "调音台", "处理器", "时序器", "摄像机", "显示", "屏", "升降器", "交换机", "机柜", "软件", "灯",
              "天线", "接收", "解码", "矩阵", "中控", "录播", "控制器", "分配器"]
# 同义：客户写 A，配 B 也算同类
SAME = {"音箱": {"音箱", "音柱", "喇叭", "扬声器"}, "音柱": {"音箱", "音柱"}, "喇叭": {"音箱", "喇叭", "扬声器"},
        "主机": {"主机", "服务器", "控制器"}, "服务器": {"服务器", "主机"}, "终端": {"终端", "解码"},
        "解码": {"解码", "终端"}, "显示": {"显示", "屏"}, "屏": {"屏", "显示"}, "话筒": {"话筒", "麦"},
        "寻呼": {"寻呼", "话筒"}, "接收": {"接收", "话筒"},
        # 别家叫「IP消防报警采集器」，我们是「数字IP网络消防主机」（带16路报警接口），历史纠正：应配 HT-PAR01-1016BTW
        # 名字里带「消防」就能对上消防主机；「采集」本身不等同消防（IP 音频采集器我们没有，要标无对应）
        "消防": {"消防"}, "采集": {"采集"}}


# 非主营：这些留空是对的（技能规矩「非主营留空」），不算漏配
NON_MAIN = re.compile(r"监控|安防|闸机|人脸|门禁|电梯|枪型|枪机|球机|半球|网桥|电脑|主机电脑|PC|笔记本|打印|"
                      r"服务费|集成服务|安装|调试|施工|线缆|HDMI|网线|光纤|空调|家具|桌椅|UPS|配电|"
                      r"立杆|地插|延长线|吊架|LED.{0,4}屏|拼接屏|机柜|加密狗|LED模组|接收卡|发送卡|拼接(控制|处理)器", re.I)
# LED 显示屏系统（模组/接收卡/发送卡/拼接控制器/视频处理器，诺瓦、卡莱特等）我们没有，必须留空。
# 2026-10-26 方案部指出：「拼接控制器（诺瓦 V16）」被配成 HT-2010——那是分布式中控的「4K30拼接输出节点」，完全不是一类
LED_SYS = re.compile(r"LED|接收卡|发送卡|拼接(控制|处理)器|视频处理器", re.I)
LED_BRAND = re.compile(r"诺瓦|NovaStar|卡莱特|Colorlight|小鸟", re.I)    # 小鸟科技：LED 拼接器（2026-10-28 655 题备注写「小鸟」）

# ------------------------------------------------------------------ 先判系统，再找型号（2026-10-26 加）
# 前两个大错都是「按字面配」：「拼接控制器」配了分布式拼接节点、参数里的「5G」当成 5G 系统。
# 客户清单里有现成线索：分组标题（三、显示系统）、品牌（诺瓦/海康）、参数（红外夜视）。先判这行属于哪个系统：
#   我们不做的系统 → 直接留空；我们做的系统 → 只在这个系统的产品里优先找、配到别的系统要提醒。
# 资料库核对过：LED屏/投影/显示器电视/监控/门禁/电脑/UPS 我们没有；会议摄像机、视频会议终端、交换机机柜网线（外购辅材）有。
NOT_OURS = [   # (系统, 名称要命中的, 分组/品牌/参数也要指向它才算（None=不需要）, 分组是这些时不算)
    ("LED显示屏", r"LED(?!.{0,12}(灯|时钟))|接收卡|发送卡|LED.{0,6}(视频)?处理器|模组|全彩|(?<![A-Z])P[1-9](\.\d)?(?![0-9A-Z])|条型屏|条屏|会标屏|屏体", None, r"灯光|舞台"),
    ("LED显示屏", r"拼接(控制|处理)器|视频处理器|拼接器", r"显示|LED|诺瓦|NovaStar|卡莱特|Colorlight|小鸟|大屏", None),
    ("投影/幕布", r"投影(仪|机)|幕布|激光电视|^投影$", None, None),
    ("显示器/电视/拼接屏", r"显示器|显示屏|辅屏|电视机|液晶电视|拼接屏|拼接单元|LCD.{0,4}拼接|液晶拼接", None, None),
    ("监控", r"硬盘录像|NVR|DVR|监控|枪机|球机|半球|枪型|筒机|人脸抓拍", None, None),
    ("监控", r"摄像机|摄像头", r"监控|安防|海康|大华|宇视|红外|夜视|IP6[67]|枪|球机|半球", r"录播|教室|课堂|会议|跟踪|导播"),
    ("门禁/闸机/对讲", r"门禁|闸机|电梯|访客|可视对讲门", None, None),
    ("门禁/闸机/对讲", r"人脸识别", r"门禁|闸机|通道|考勤", None),
    ("电脑", r"电脑(?!控|灯)|笔记本|台式机|工作站|主机电脑|(?<![A-Z])PC(?![A-Z])", None, None),
    ("UPS/配电", r"UPS|不间断电源|配电箱|配电柜", None, r"灯光|舞台"),
    ("土建/管材/强电", r"土方|回填|挖沟|沟槽|基础$|接地(角钢|圆钢|扁钢|装置|极|体)|角钢|圆钢|扁钢|断路器|电力电缆|^SC\d+|^PVC\d*|JDG|KBG|线槽|桥架|"
                     r"镀锌钢管|穿线管|^钢管|手孔|人井|检查井|混凝土|凿槽|开槽|防火封堵|弱电箱基础|监控杆基础", None, None),
    ("施工/服务", r"安装费|调试费|施工|服务费|集成服务|人工费|运输费|培训费|^安装|^调试|安装调试|钢结构|实施|集成$|系统集成|调试$|试运行$|措施项目", None, None),
]
# 名称里有我们的设备词（「AI广播话筒（监控办公室）」「电脑灯」「无纸化会议投影终端」）→ 不按「不做」判，只是顺带提到位置/用途
OUR_STRONG = re.compile(r"话筒|音箱|音柱|功放|广播|会议|灯|调音|音频|寻呼|无纸化|中控|扩声|时序|号角|喇叭|天花|吸顶|录播|"
                        r"信息发布|公告|班牌|广告机|跟踪|图像处理|面板|AP|发射器|接收器|网络广播|会议系统")
# 我们做的系统 → 产品浏览一级系列
OURS = {
    "广播": ("天问广播", "网络广播", "智能广播", "广播配套"),
    "扩声": ("专业音箱", "专业功放", "专业周边", "音频处理", "无线话筒", "教学扩声"),
    "会议": ("有线会议", "无线会议", "无纸化会议", "会议预约发布"),
    "视频会议/摄像": ("视频会议摄像机", "录播系统", "教学一体机&会议平板"),
    "录播": ("录播系统", "视频会议摄像机"),
    "中控/矩阵": ("中控矩阵分布式",),
    "显示(会议平板)": ("教学一体机&会议平板",),
    "灯光": ("灯光系统",),
    "辅材": ("耗材辅材",),
}
# 只用不会混的词；「矩阵/切换器/编解码器」广播、扩声、中控都有，看不出来就不判（宁可不判，不误报）
OURS_WORDS = [
    ("灯光", r"(?<!LED)(?<!指示)灯(?!箱)|调光台|灯光控台|电脑控台|硅箱|影像处理|追光"),
    ("录播", r"录播"),
    ("会议", r"会议(系统)?(主机|单元)|手拉手|主席|代表单元|鹅颈|无纸化|表决|同传|升降器|会议话筒|视像跟踪"),
    ("广播", r"广播|寻呼|消防|定压|分区|IP.{0,8}(终端|音箱|音柱)|网络.{0,8}(终端|音箱|音柱)|校时|打铃|号角|草坪|音柱|替补|音频采集"),
    ("扩声", r"音箱|扬声器|功放|调音台|音频处理器|反馈|无线话筒|手持|领夹|头戴|返听|监听|低音|线阵|分频|天线|接收机|音频矩阵"),
    ("视频会议/摄像", r"视频会议|会议终端|MCU|云台|摄像机|摄像头"),
    ("显示(会议平板)", r"会议平板|智慧黑板|交互平板|触摸一体机|教学一体机"),
    ("中控/矩阵", r"中控主机|中央控制|视频矩阵|HDMI|无线投屏|分布式|交互管理|节点|KVM|视频切换|拼控|输入卡|输出卡"),
    ("辅材", r"(?<![天无有])线(?!阵)|机柜|交换机|路由|跳线|水晶头|插座|支架|吊架|配件|辅材"),
]
FAR = [{"广播", "扩声", "会议"}, {"视频会议/摄像", "录播", "中控/矩阵", "显示(会议平板)", "会议"}, {"灯光"}]   # 同组内互相配不算离谱


def name_head(name):
    """名称格第一句。名称格里常把整段参数也写进去（「反馈抑制器产品特点：…信号灯指示…」）"""
    raw = re.sub(r"^\s*\d*[.、]?\s*(设备名称|名称)\s*[:：]\s*", "", str(name or ""))
    raw = re.sub(r"\n\s*\n", "\n", raw.strip())                         # 「数字音频处理\n\n器」：先把空行去掉
    raw = re.sub(r"\n(?=[^\n:：]{1,5}(\n|$))", "", raw)                  # 「LED平板柔\n光灯」这种折行接回去
    return re.sub(r"\s+", "", re.split(r"\n|产品特点|技术参数|[:：]", raw)[0].strip() or raw[:30])[:30]


def system_of(name, params="", section="", brand=""):
    """判断客户这一行属于哪个系统 → (系统, 是否我们做, 依据)。名称最重要；分组/品牌/参数用来分辨同一个词在不同系统的意思。
    依据以「分组」开头的是弱判断（名称看不出来，只按分组推的），不拿来报跨系统。"""
    name = name_head(name)
    name, params, section, brand = (re.sub(r"\s+", "", str(x or "")) for x in (name, params, section, brand))
    ctx = f"{section} {brand} {params[:200]}"
    core = re.sub(r"[（(][^）)]*[）)]?", "", name) or name     # 括号里多是说明（「含安装调试」「PC端管理」），不拿来判「不做」
    strong = OUR_STRONG.search(name)
    for sysname, rx, need, unless in NOT_OURS:
        if strong and not (sysname == "施工/服务" and re.search(r"费$|^安装|^调试|^施工|^人工|调试$|试运行$|措施项目", core)):
            continue                                       # 「功放（含安装调试）」「网络广播系统集成」是设备；「音箱安装费」才是服务
        if re.search(rx, core, re.I) and (need is None or re.search(need, ctx, re.I)) and not (unless and re.search(unless, ctx)):
            why = f"名称「{name[:12]}」" + (f"+{'分组' if re.search(need, section, re.I) else '品牌/参数'}" if need else "")
            return sysname, False, why
    # 分组标题就是我们不做的系统（「门禁管理系统」「视频监控系统」「弱电专网系统」…）：组里的网线/底盒/配管也跟着不做
    if not strong and re.search(r"监控|门禁|停车|入侵|巡更|对讲|楼宇|建筑设备|能耗|综合布线|布线系统|专网|网络系统|计算机网络|安防|一卡通|梯控|电梯|无线对讲",
                                section):
        return f"其他系统（{section[:12]}）", False, f"分组「{section[:12]}」"
    if not strong and re.search(r"显示|大屏|LED屏", section) and re.search(r"LED|拼接|模组|接收卡|发送卡|屏体|会标屏|条屏|视频处理器", name):
        return "LED显示屏", False, f"分组「{section[:10]}」"
    # 「音柱」默认是广播音柱；参数是专业扩声（分频/低音单元/低频反射）又没写定压的，是会议室扩声音柱（1013：配成了广播定压音柱）
    if re.search(r"音柱|声柱", name) and not re.search(r"IP|网络|定压|广播|室外|草坪", name, re.I) \
            and re.search(r"分频|2-way|LFDriver|低频反射|倒相|线阵|线声源|bassreflex|低音单元|高音单元", params[:300], re.I) \
            and not re.search(r"定压|70V|100V", params[:300], re.I):
        return "扩声", True, "名称「音柱」+参数是专业扩声（分频/低音单元）"
    # 会议室里的音柱（「会议音柱」或在会议/扩声分组下）没写定压 → 会议室扩声，不是广播音柱（2026-10-29 方案部：40W 会议音柱配了 200W）
    if re.search(r"音柱|声柱", name) and not re.search(r"IP|网络|定压|广播|室外|草坪", name, re.I) \
            and re.search(r"会议|扩声|报告厅|多功能厅", f"{name} {section}") and not re.search(r"广播|背景音乐", section) \
            and not re.search(r"定压|70V|100V", params[:300], re.I):
        return "扩声", True, "名称「音柱」+会议室/扩声分组，没写定压"
    # 广播分组下的功放是广播功放（名称「数字功放」本身按扩声判，以前靠「定压」加分拉回广播；有了常用表加分得说清楚）
    if re.search(r"功放|功率放大", name) and re.search(r"广播|背景音乐", section) and not re.search(r"扩声|会议|报告厅|剧场|礼堂", section):
        return "广播", True, f"名称「{name[:12]}」+分组「{section[:10]}」"
    for sysname, rx in OURS_WORDS:
        if re.search(rx, name, re.I):
            return sysname, True, f"名称「{name[:12]}」"
    for sysname, rx in OURS_WORDS:                 # 名称看不出来：按分组标题推（弱）
        if section and re.search(rx, section, re.I):
            return sysname, True, f"分组「{section[:10]}」"
    return "", True, ""


def _note_brand(v):
    """客户表没有品牌列时，备注里常写品牌（「小鸟」「华为」）：只取短的、不带数字的一个词当品牌线索"""
    v = str(v or "").strip()
    return v if 1 < len(v) <= 10 and not re.search(r"\d|配套|含|需|按|见|详", v) else ""


SECTION = re.compile(r"^\s*([一二三四五六七八九十]+|[0-9]{1,2}|[A-Z])\s*[、.．:：]\s*\S|系统\s*$|部分\s*$")


def _rows_no_header(sheet, rows):
    """没有表头的清单：名称列 = 短中文名最多的列；参数列 = 平均最长的文本列。猜不出就不读。"""
    data = [r for r in rows if sum(1 for c in r if c not in (None, "")) >= 3]   # 只有一两个格子的是分组标题
    if len(data) < 3:
        return []
    ncol = max(len(r) for r in data)
    short = lambda c: isinstance(c, str) and re.fullmatch(r"\s*[一-龥A-Za-z0-9（）()\-、/ ]{2,20}\s*", c) and re.search(r"[一-龥]{2}", c) \
        and not re.fullmatch(r"\s*(个|台|套|只|米|项|批|组|根|块|对|副|樘|m|M)\s*", c)  # noqa: E731
    sc = [sum(1 for r in data if j < len(r) and short(r[j])) for j in range(ncol)]
    nc = max(range(ncol), key=lambda j: sc[j])
    if sc[nc] < len(data) * 0.6:
        return []
    ln = [sum(len(str(r[j])) for r in data if j < len(r) and isinstance(r[j], str)) / len(data) for j in range(ncol)]
    pc = max((j for j in range(ncol) if j != nc), key=lambda j: ln[j], default=None)
    out, sec = [], ""
    for i, r in enumerate(rows):
        filled = [str(c).strip() for c in r if c not in (None, "") and str(c).strip()]
        if 0 < len(filled) <= 2:                  # 分组标题行（「会议系统--15人会议室」「显示系统」）
            sec = filled[-1][:30]
            continue
        if nc < len(r) and short(r[nc]):
            out.append({"sheet": sheet, "row": i + 1, "name": str(r[nc]).strip(),
                        "params": str(r[pc] or "") if pc is not None and pc < len(r) else "", "brand": "", "model": "", "qty": "", "section": sec})
    return out


def read_rows(path):
    """读一张清单（客户原表或比单表都行）：逐行给出 名称/参数/品牌/型号/所在分组。分组标题行可以在任意列（序号列合并单元格常见）。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        if re.search(r"日志|说明|封面|汇总|目录", ws.title):
            continue
        rows = list(ws.iter_rows(values_only=True))
        # 表头 = 前 25 行里表头词最多的一行（比单表上面常有「设计单位/项目名称/客户单位」封面，不能见「名称」就认）
        hw = re.compile(r"名称|品名|货物|^设备$|^产品$|型号|数量|单位|参数|规格|品牌|单价|金额|备注|序号")
        best = max(((sum(1 for c in r if c and len(str(c)) < 16 and hw.search(str(c).strip())), -i) for i, r in enumerate(rows[:25])),
                   default=(0, 0))
        best = (best[0], -best[1])                    # 几行表头词一样多（分页重复表头）取第一行，不然第一页的设备整段漏读（2026-10-27 草地音箱）
        hi = best[1] if best[0] >= 3 else None
        if hi is None or not any(c and re.search(r"名称|品名|货物|设备|产品", str(c)) for c in rows[hi]):
            out += _rows_no_header(ws.title, rows)        # 没表头（招标工程量清单常见）：按内容猜名称列/参数列
            continue
        cells = [str(c or "") for c in rows[hi]]
        find = lambda pat, excl=None: next((j for j, c in enumerate(cells) if re.search(pat, c) and len(c) < 16  # noqa: E731
                                            and not (excl and re.search(excl, c))), None)
        col = {"name": find(r"名称|品名|货物|^设备$|^产品$", r"品牌|型号"), "params": find(r"参数|规格|性能|指标|要求|特征|描述", r"型号|编码"),   # 招标工程量清单叫「项目特征描述」
               "brand": find(r"品牌|厂家|厂商"), "qty": find(r"数量"), "note": find(r"^\s*备注\s*$")}
        h2, mc = find_header(rows[hi:hi + 1])
        col["model"] = mc["model"] if mc else None
        if col["name"] is None:
            continue
        sec = ""
        for ri, r in enumerate(rows[hi + 1:], start=hi + 2):
            g = lambda k: str(r[col[k]]).strip() if col[k] is not None and col[k] < len(r) and r[col[k]] is not None else ""  # noqa: E731
            filled = [str(c).strip() for c in r if c not in (None, "") and str(c).strip()]
            if not filled:
                continue
            if len(filled) <= 2 and any(SECTION.search(c) and not re.fullmatch(r"[\d.]+", c) and len(c) <= 30 and "\n" not in c
                                        and not re.search(r"^\s*\d+[.、．]\s*[^\s]{1,8}[:：]", c) for c in filled) and not g("model"):
                sec = next(c for c in filled if SECTION.search(c))[:30]
                continue
            name = g("name")
            if not name or re.fullmatch(r"合计|小计|总计.*|设备合计|项目名称|设备名称|名称|品名|序号", name):   # 分页重复的表头行
                continue
            out.append({"sheet": ws.title, "row": ri, "name": name, "params": g("params"), "brand": g("brand") or _note_brand(g("note")),
                        "model": g("model"), "qty": g("qty"), "section": sec})
    return out


# ------------------------------------------------------------------ 方案部经验库（2026-10-26 加）
# references/experience.json：方案流转里方案部人工成果「客户设备名称 → 我司型号」，exp_build.py 生成（每晚重建）。
# 方案部有固定习惯（前置放大器 100% 配 HT-PAR01-1010TW、寻呼话筒 100% 配 HT-PAE01-1032TW…），模型凭字面猜不到。
_EXP = None


def load_exp():
    global _EXP
    if _EXP is None:
        f = ROOT / "references" / "experience.json"
        _EXP = json.load(open(f, encoding="utf-8"))["rows"] if f.exists() else []
        for e in _EXP:
            e["_bg"] = bigrams(e["name"])
    return _EXP


# 只差一个字但完全不是一回事：「视频处理器」≠「音频处理器」（二字组重合 0.6，以前被当成相近）
CONFLICT = [("视频", "音频"), ("有线", "无线"), ("室内", "室外"), ("输入", "输出"), ("吸顶", "壁挂"), ("主席", "代表"),
            ("软件", "主机"), ("发射", "接收"), ("定压", "定阻"), ("手持", "领夹"), ("手持", "头戴")]


def _conflict(a, b):
    # 一边只有 x、另一边只有 y 才算对不上（产品名里两个词都有的——「输出卡（支持…输入）」——不算）
    return any((x in a and y not in a and y in b and x not in b) or (y in a and x not in a and x in b and y not in b) for x, y in CONFLICT)


_EXPC = {}


def experience(name, top=3, idx=None, min_sim=0.6):
    k = (name_head(name), top, min_sim)
    if k not in _EXPC:
        _EXPC[k] = _experience(name, top, idx, min_sim)
    return _EXPC[k]


def _experience(name, top=3, idx=None, min_sim=0.6):
    """方案部以前遇到名称相近的行（第一句二字组重合 ≥0.6）配过什么 → ([(型号, 项目数, 最近日期)], 相近行涉及的项目数)。
    只列现在还能配的型号；按用过的项目数排。"""
    n = name_head(name)
    nb = bigrams(n)
    if len(nb) < 1:
        return [], 0
    agg, projs = {}, set()
    for e in load_exp():
        sim = 1.0 if e["name"] == n else len(nb & e["_bg"]) / (len(nb | e["_bg"]) or 1)
        if sim < min_sim or (sim < 1.0 and _conflict(n, e["name"])):
            continue
        projs.add(e["rid"])
        a = agg.setdefault(e["model"], [set(), ""])
        a[0].add(e["rid"])
        a[1] = max(a[1], e["date"])
    idx = idx or index(load_kb())
    res = []
    for m, (rids, last) in sorted(agg.items(), key=lambda kv: -len(kv[1][0])):
        ps = find_model(idx, m)
        if ps and avail(ps[0]) and not is_old(m):
            res.append((m, len(rids), last))
        if len(res) >= top:
            break
    return res, len(projs)


def _norm(s):
    return re.sub(r"[\s\W_]+", "", str(s or "")).lower()


_CASEC = {}


def same_case(name, params, idx=None):
    """历史上同样的清单行（2026-10-28 用户：「有同样的方案时，不能直接找答案吗」）：
    名称一样、参数原文二字组相似度 ≥0.6 的方案部历史行，按相似度投票 → (型号, 最高相似度, 项目数) 或 None。
    考卷 16 题实测：相似度 ≥0.7 的 21 行里 20 行和方案部这次一样（95%）；0.6~0.7 约三分之二一样。"""
    k = (name_head(name), str(params)[:200])
    if k in _CASEC:
        return _CASEC[k]
    n, pb = _norm(name_head(name)), bigrams(_norm(params)[:150])
    res = None
    if n and len(pb) >= 15:
        votes = {}
        for e in load_exp():
            if _norm(e["name"]) != n:
                continue
            if "_pbg" not in e:
                e["_pbg"] = bigrams(_norm(e.get("params"))[:150])
            if not e["_pbg"]:
                continue
            sim = len(pb & e["_pbg"]) / len(pb | e["_pbg"])
            if sim >= 0.6:
                v = votes.setdefault(e["model"], [0.0, 0.0, set()])
                v[0] += sim
                v[1] = max(v[1], sim)
                v[2].add(e["rid"])
        idx = idx or index(load_kb())
        for m, (w, best, rids) in sorted(votes.items(), key=lambda kv: -kv[1][0]):
            ps = find_model(idx, m)
            if ps and avail(ps[0]) and not is_old(m):
                res = (m, round(best, 2), len(rids))
                break
    _CASEC[k] = res
    return res


def exp_line(name, idx=None):
    res, n = experience(name, idx=idx)
    if not res:
        return ""
    return (f"方案部经验：名称相近的行在 {n} 个项目里出现过，配过 "
            + "、".join(f"{m}（{k} 个项目，最近 {d}）" for m, k, d in res)
            + "——固定配法就照着来；要按功率/尺寸选的，以客户参数为准")


# 方案部常用设备表（2026-10-29 方案部给的「1-会议系统配单表」，966 个在售型号）：配单一般就从这里选，同档位优先。
# 表外的不算错（内部业务平台库里更全）。references/方案部常用设备.json
COMMON_W = float(os.environ.get("KB_COMMON_W", "3.0"))
_COMMON = None


def common_of(m):
    """型号在方案部常用表里 → {"表","分组","名称","备注"}，不在 → None"""
    global _COMMON
    if _COMMON is None:
        try:
            _COMMON = {N(k): v for k, v in json.load(open(ROOT / "references" / "方案部常用设备.json", encoding="utf-8"))["devices"].items()}
        except (OSError, ValueError, KeyError):
            _COMMON = {}
    return _COMMON.get(N(m))


def is_ip_broadcast(p):
    return family_of(p) == "广播" and bool(re.search(r"网络|IP", f"{p.get('device_name')} {series_of(p)}"))


def is_analog_source(p):
    """模拟音频设备：8017D/8118/8005/8028 这类「广播音源设备」，以及不带网络的调音台/前置放大/播放器/收音"""
    dn, ser = str(p.get("device_name")), series_of(p)
    return bool(ser.endswith("广播音源设备") or re.match(r"HT-(8017D|8118|8005|8028)", str(p.get("model")))
                or (re.search(r"调音台|前置放大|播放|收音|调谐", dn) and not re.search(r"网络|IP", f"{dn} {ser}")))


def family_of(p):
    top = series_of(p).split("/")[0]
    return next((k for k, v in OURS.items() if top in v), "")


def far_apart(row_sys, p, why=""):
    """客户这行的系统和我们配的产品差得远（如显示类配了广播产品）→ 提醒。只按分组推出来的系统不算数。"""
    ps = family_of(p)
    if why.startswith("分组"):
        return False
    if not row_sys or not ps or ps == row_sys or "辅材" in (row_sys, ps):
        return False
    return not any(row_sys in g and ps in g for g in FAR)


# 旧系列（配套系列）已废弃，2026-10-26 方案部再次确认「配单按禁止用旧系列为准」。
# 例外：HT-8005/8118/8028/8017D 是通用辅助设备，可搭配天问系列（见 bid 技能 SKILL.md）
# 资料库里归在「天问广播」系列的不算旧系列（如 HT-A9000FR 分控软件）
OLD_SERIES = re.compile(r"\s*HT-(A?9\d{3}|8\d{3}|\d{3,4}DB)", re.I)
OLD_OK = re.compile(r"\s*HT-(8005|8118|8028|8017D|9989T)", re.I)   # 9989T：方案部常用表里在用（2026-10-29 用户确认）
_IDX = None


def is_old(m):
    global _IDX
    if not OLD_SERIES.match(str(m or "")) or OLD_OK.match(str(m or "")):
        return False
    if _IDX is None:
        _IDX = index(load_kb())
    return not any(series_of(p).startswith("天问") for p in find_model(_IDX, m))


def N(s):
    return re.sub(r"[^0-9A-Z]", "", str(s or "").upper())


_KB = None


def load_kb():
    """资料库只读一次（plan 一张 80 行的清单以前要重读 160 次 10MB JSON，34 秒）"""
    global _KB
    if _KB is None:
        _KB = _load_kb()
    return _KB


def _load_kb():
    data = json.load(open(KB_PATH, encoding="utf-8"))
    out = []
    for c in data:
        for s in c["subcategories"]:
            for p in s["products"]:
                specs = " ".join(f"{x.get('key', '')}:{x.get('value', '')}" for x in p.get("technical_specs") or [])
                full = " ".join(x.get("full_text", "") for x in p.get("full_official_specs") or [])
                p["_text"] = f"{p.get('device_name', '')} {p.get('series', '')} {specs} {full}"
                p["_cat"] = f"{c['category_code']}/{s['subcategory_code']}"
                out.append(p)
    return out


_IDXC = {}


def index(kb):
    if id(kb) not in _IDXC:
        _IDXC[id(kb)] = _index(kb)
    return _IDXC[id(kb)]


def _index(kb):
    idx = {}
    for p in kb:
        for k in (p.get("model"), p.get("internal_model"), p.get("catalog_model")):
            if N(k):
                idx.setdefault(N(k), []).append(p)
    return idx


def find_model(idx, m):
    """型号查找，在售的排前面（同一型号在资料库里可能有停单、在售两条，如 HT-A9000FR——以前只看第一条会误报停单）"""
    return sorted(_find_model(idx, m), key=lambda p: not avail(p))


def _find_model(idx, m):
    """精确 → 去掉括号后缀 → 前缀（客户/输出里常带 (2.0) -S 之类后缀）"""
    n = N(m)
    if not n:
        return []
    if n in idx:
        return idx[n]
    n2 = N(re.sub(r"[（(].*?[)）]", "", str(m)))
    if n2 in idx:
        return idx[n2]
    if len(n) >= 7:
        c = [p for k, v in idx.items() if len(k) >= 7 and (k.startswith(n) or n.startswith(k)) for p in v]
        if len({id(x) for x in c}) <= 3:
            return c
    return []


def series_of(p):
    """产品浏览里的「系列/大类」，如「无线会议/五、无线U段手拉手系列」——配套判断以它为准，不看参数字眼。"""
    return "/".join(str(p.get("series") or "").split("/")[:2])


def link_of(ser):
    """会议类的连接方式：不同连接方式之间绝对不兼容。"""
    if "U段" in ser:
        return "无线U段"
    if re.search(r"5G|WIFI|WiFi", ser):
        return "5G WiFi"
    if ser.startswith("有线"):
        return "有线"
    if ser.startswith("无线"):
        return "无线"
    return ""


# 产品库里分在两个系列、但实际是一套的：查主机时把配套系列一起列出来
KIT_EXTRA = {"中控矩阵分布式/五、混合矩阵": ["中控矩阵分布式/六、混合矩阵板卡"]}


def avail(p):
    return p.get("available", True) is not False and p.get("status") not in ("停止配单", "已隐藏", "产品浏览里已没有")


def bigrams(s):
    s = re.sub(r"\s+", "", str(s or "")).replace("声柱", "音柱")     # 「线声源声柱」= 音柱（1013 方案部配 L4504）
    return {s[i:i + 2] for i in range(len(s) - 1)}


def numbers(s):
    """抽数值约束：功率W、尺寸寸、路数、通道 —— 用来给候选加减分"""
    return {(u, float(v)) for v, u in re.findall(r"(\d+(?:\.\d+)?)\s*(W|寸|路|通道|英寸)", str(s or ""), re.I)}


def types_in(s):
    return {w for w in TYPE_WORDS if w in str(s or "").replace("声柱", "音柱")}    # 「线声源声柱」= 音柱


def type_ok(name, p):
    """客户设备名里的类型词，候选的设备名里要有对应（同义也算）。客户名里没有类型词时不限制。"""
    qt = types_in(name)
    if not qt:
        return True
    pt = types_in(p.get("device_name"))
    return any(pt & SAME.get(t, {t}) for t in qt)


# 规格档位（2026-10-26 加）：留出评估里排错最多的是档位——「音柱180W」该配 RCLYZ83P180、「8路时序器」排成16路、
# 「12路调音台」排成14路、「3寸音柱」排成4寸。客户名称里的数字最要紧，参数里的次之；产品看名称、前几条参数、型号里的 P180 这类。
IFACES = [("HDMI", r"HDMI|(?<![A-Z])HD(?![A-Z])|HD输|HD-"), ("DVI", r"DVI"), ("SDI", r"SDI"), ("VGA", r"VGA"),
          ("HDBaseT", r"HDBaseT|HDBT"), ("Dante", r"Dante"), ("AES67", r"AES67"), ("光纤", r"光纤")]
FORMS = ["壁挂", "吸顶", "音柱", "草坪", "号角", "防爆", "手持", "领夹", "头戴", "鹅颈"]


_CN = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def _cn2num(m):
    """「十六」→16、「八」→8、「二十四」→24"""
    s = m.group(1)
    if "十" in s:
        a, _, b = s.partition("十")
        v = (_CN.get(a, 1) if a else 1) * 10 + (_CN.get(b, 0) if b else 0)
    else:
        v = _CN.get(s, 0)
    return f"{v}{m.group(2)}"


def specs_in(text):
    """{单位: {数值}}：W 功率、寸 尺寸、路 路数/通道（8进8出 记 8）。中文数字也认（八路、十六路、四通道）"""
    t = re.sub(r"([一二两三四五六七八九十]{1,3})(路|通道|寸)", _cn2num, str(text or ""))
    out = {}
    for v, u in re.findall(r"(\d+(?:\.\d+)?)\s*(W|瓦|寸|英寸|″|\"|路|通道)(?![a-z])", t, re.I):
        u = {"瓦": "W", "w": "W", "英寸": "寸", "″": "寸", '"': "寸", "通道": "路"}.get(u, u.upper() if u in "wW" else u)
        out.setdefault(u, set()).add(float(v))
    for a, b in re.findall(r"(\d{1,2})\s*进\s*(\d{1,2})\s*出", t):
        out.setdefault("路", set()).update({float(a), float(b)})
    return out


def _pspecs(p):
    if "_spec" not in p:
        head = " ".join([str(p.get("device_name") or "")] + [f"{x.get('key')}{x.get('value')}" for x in (p.get("technical_specs") or [])[:6]])
        sp = specs_in(head)
        m = re.search(r"P(\d{2,3})$", str(p.get("model") or ""))          # RCLYZ83P180 → 180W
        if m:
            sp.setdefault("W", set()).add(float(m.group(1)))
        p["_spec"] = sp
    return p["_spec"]


# ------------------------------------------------------------------ 硬参数（2026-10-27 方案部：「要 4K 的配了 2K」）
# 客户写明的硬指标，产品达不到就是不满足——方案部自己以前也常配 2K（经验库会往 2K 推），所以这条要压过经验加分
def _ptext(p):
    return f"{p.get('device_name') or ''} {series_of(p)} " + " ".join(
        f"{x.get('key')}:{x.get('value')}" for x in (p.get("technical_specs") or [])[:12]) + " " + " ".join(
        x.get("full_text", "") for x in (p.get("full_official_specs") or [])[:15])


def is_4k(text):
    """「24K/44.1K」是音频采样率，不算（320P 曾因此被当成 4K）"""
    return bool(re.search(r"(?<![0-9.])4K(?![bB])|2160P|3840\s*[x×*]\s*2160|超高清", str(text), re.I))


def _rated_w(text):
    """额定功率（W）。只认明确写「额定/输出功率」那一段；几档可选（「10W,20W两档」）取最高档；2×350W 记单通道 350。
    峰值/节目/最大功率不算。"""
    m = re.search(r"额定(?:输出)?功率[^0-9\n]{0,12}([^\n；;。]{0,40})", str(text or ""))
    if not m:
        return None
    seg = re.sub(r"[（(][^）)]*(PEAK|PMPO|峰值|最大)[^）)]*[）)]?", "", m.group(1), flags=re.I)   # 「200W（300W，PEAK）」取 200
    seg = re.split(r"峰值|节目|最大|连续|PEAK|PMPO", seg, flags=re.I)[0]
    vals = [float(b) for a, b in re.findall(r"(?:(\d+)\s*[×xX*]\s*)?(\d+(?:\.\d+)?)\s*W", seg)]
    return max(vals) if vals else None


def _lead_w(text):
    """参数开头直接写的瓦数：「40W 全频音柱，含壁装支架」「1.额定输出 650W」→ 40 / 650（只认开头一小段，免得抓到别的瓦数）"""
    m = re.match(r"\s*(?:\d[.、]\s*)?(?:额定)?(?:输出)?(?:功率)?\s*[:：]?\s*(\d{1,4})\s*W(?![a-zA-Z])", str(text or ""))
    return float(m.group(1)) if m else None


def _name_w(text):
    """名称里写的瓦数：「壁挂音箱（食堂，20W）」「600W专业功放」「2×200W」→ 取最大的单通道值"""
    vals = [float(b) for a, b in re.findall(r"(?<![\d.])(?:(\d+)\s*[×xX*]\s*)?(\d{1,4})\s*W(?![a-zA-Z])", str(text or ""))]
    return max(vals) if vals else None


def _amp_w(text):
    """功放单通道功率 {阻抗: 瓦}：「≥2×900W/8Ω」「立体声功率（8Ω）:2×600W」。桥接不算。"""
    out = {}
    t = re.sub(r"桥接[^,，；;\n]{0,30}", "", str(text or ""))
    for a, w, ohm in re.findall(r"(?:(\d+)\s*[×xX*]\s*)?(\d{2,5})\s*W\s*/\s*(\d)\s*Ω", t):
        out.setdefault(int(ohm), float(w))
    for ohm, a, w in re.findall(r"[（(]\s*(\d)\s*Ω\s*[)）][^0-9\n]{0,6}(?:(\d+)\s*[×xX*]\s*)?(\d{2,5})\s*W", t):
        out.setdefault(int(ohm), float(w))
    for ohm, w in re.findall(r"[@＠]\s*(\d)\s*Ω\s*[:：]?\s*(\d{2,5})\s*W", t):      # 「立体声@8Ω：200W×2」
        out.setdefault(int(ohm), float(w))
    return out


def _amp_ch(p):
    """功放通道数：名称里的「四通道/双通道」→ 额定功率写法「2×350W」→ 默认 1"""
    n = specs_in(str(p.get("device_name") or "")).get("路")
    if n:
        return max(n)
    m = re.search(r"(?:额定(?:输出)?|立体声)功率(?:[（(]\s*\d\s*Ω\s*[)）])?[^0-9\n]{0,12}(\d)\s*[×xX*]\s*\d+(?:\.\d+)?\s*W",
                  str(p.get("_text") or ""))
    return int(m.group(1)) if m else 1


def _lf_size(text):
    """低音（低频）单元尺寸（寸）：「低音单元/LF Driver：4x4.5”」「低频单元：12寸」「低音：2×10英寸」「单元组成:全频2×3寸」"""
    m = re.search(r"(?:低(?:音|频)|全频)(?:单元|喇叭|扬声器)?[^0-9\n]{0,25}?(?:\d+\s*[x×X*]\s*)?(\d{1,2}(?:\.\d)?)\s*(?:寸|英寸|[\"″”]|inch)",
                  str(text or ""))
    return float(m.group(1)) if m else None


def _name_n(text, unit):
    """名称里明确写的路数/寸数：「八路电源时序器」→8，「8寸音箱」→8（中文数字也认）"""
    t = re.sub(r"([一二两三四五六七八九十]{1,3})(路|通道|寸)", _cn2num, str(text or ""))
    if "寸" in unit:                             # 「双6.5寸配套单15线阵低音」：取最大的（低音单元）
        vals = [float(v) for v in re.findall(r"(?<![\d.、])(\d{1,2}(?:\.\d)?)\s*(?:寸|英寸)", t)]
        vals += [float(v) for v in re.findall(r"[单双](\d{1,2}(?:\.\d)?)(?=[^\d.寸]|$)", t)]
        return max(vals) if vals else None
    m = re.search(r"(?<![\d.、])(\d{1,3}(?:\.\d)?)\s*(?:" + unit + r")", t)
    return float(m.group(1)) if m else None


def is_pro_amp_row(name, section=""):
    """会议/扩声分组里的功放 = 扩声功放（专业功放系列），名称或分组写了广播/背景音乐的除外"""
    return bool(re.search(r"功放|功率放大", str(name)) and re.search(r"会议|扩声|报告厅|多功能厅|礼堂|剧场", f"{name} {section}")
                and not re.search(r"广播|背景音乐|消防", f"{name} {section}"))


def param_gaps(name, params, p):
    """→ [(级别, 说明)]：客户写明的数字/硬指标，产品看得出达不到的。数字要抽得准才报（不准宁可不报），报了就是 ❌ 必须改。
    2026-10-27 用户：「这种有参数的、数字的，你是不能犯这种错误的」。"""
    want = f"{name} {str(params)[:800]}"
    pt = _ptext(p)
    dn = str(p.get("device_name") or "")
    if re.search(r"2K", series_of(p)):        # 资料库系列就叫「2K会议摄像机」
        pt = re.sub(r"(?<![0-9.])4K|超高清", "", pt) + " 2K"
    gaps = []
    # 分辨率
    if is_4k(want) and re.search(r"摄像|摄像机|终端|编码|解码|矩阵|显示|屏|录播|投屏|分配|切换|延长", f"{name} {dn}") \
            and not is_4k(pt) and re.search(r"1080|2K|720", pt, re.I):
        gaps.append(("❌", "客户要 4K，这个型号最高 1080P/2K"))
    # 光学变焦
    zw = re.search(r"(\d{1,2})\s*倍\s*光学|光学变焦[^0-9\n]{0,6}(\d{1,2})\s*倍", want)
    zh = re.search(r"光学变焦[^0-9\n]{0,6}(\d{1,2})\s*倍|(\d{1,2})\s*倍\s*光学", pt)
    if zw and zh:
        w, h = int(zw.group(1) or zw.group(2)), int(zh.group(1) or zh.group(2))
        if h < w:
            gaps.append(("❌", f"客户要 {w} 倍光学变焦，这个型号 {h} 倍"))
    # 额定功率：只比功放/音箱类，只比「额定」对「额定」（或名称里写的瓦数）
    specs_txt = " ".join(f"{x.get('key')}:{x.get('value')}" for x in (p.get("technical_specs") or []))
    if re.search(r"功放", f"{name} {dn}") and _amp_w(params):
        cw, pw = _amp_w(params), _amp_w(specs_txt + " " + pt)
        both = [o for o in (8, 4) if o in cw and o in pw]   # 同一阻抗下比单通道，优先 8Ω
        if both and pw[both[0]] < cw[both[0]] * 0.9:
            gaps.append(("❌", f"客户要 {cw[both[0]]:g}W/{both[0]}Ω（单通道），这个型号 {pw[both[0]]:g}W/{both[0]}Ω"))
    elif re.search(r"功放|音箱|音柱|扬声器|喇叭|号角", f"{name} {dn}"):
        w = _name_w(name) or _rated_w(params) or _lead_w(params)
        h = _name_w(dn) or _rated_w(specs_txt) or _rated_w(pt)
        if w and h and h < w * 0.9:
            gaps.append(("❌", f"客户要额定 {w:g}W，这个型号额定 {h:g}W"))
        # 音箱也不能大太多：客户要 40W 会议音柱，配了 200W（2026-10-29 方案部：要按参数功率配）。
        # 广播音箱多档抽头（3/6/10/20W）取的是最高档，不比；只比专业音箱这类单一额定功率的
        elif w and h and w >= 5 and h > w * 2 and not re.search(r"功放", f"{name} {dn}") and family_of(p) != "广播":
            gaps.append(("❌", f"客户要额定 {w:g}W，这个型号额定 {h:g}W，大太多——音箱按客户功率就近配（功放也要跟着配大）"))
    # 音箱低音单元尺寸：参数里写的「低音单元：4x4.5"」「低频单元 12寸」（2026-10-28 1013：客户要 4.5 寸，配了 4 寸音柱没查出来）
    if re.search(r"音箱|音柱|扬声器|声柱|低频|全频", f"{name} {dn}"):
        w, h = _lf_size(params), _lf_size(specs_txt + " " + pt)
        if w and h and h < w - 0.3:
            gaps.append(("❌", f"客户要 {w:g} 寸低音单元，这个型号 {h:g} 寸"))
    # 路数、尺寸：只看两边名称里明确写的
    for unit, label, rx in (("路", "路", r"路|通道"), ("寸", "寸", r"寸|英寸")):
        w, h = _name_n(name, rx), _name_n(dn, rx)
        if w and h and (h < w if unit == "路" else abs(h - w) > 0.6):
            gaps.append(("❌", f"客户要 {w:g}{label}，这个型号 {h:g}{label}"))
    return gaps


def spec_score(name, params, p):
    ps = _pspecs(p)
    s = 0.0
    for src, w in ((name, 1.5), (str(params)[:300], 0.6)):
        for u, vals in specs_in(src).items():
            if u not in ps:
                continue
            for v in vals:
                d = min(abs(x - v) / max(v, 1) for x in ps[u])
                if d <= 0.1:
                    s += w
                elif d <= 0.35:
                    s += w / 2
                elif u == "W" and max(ps[u]) > v:          # 功率大一些问题不大
                    s += 0
                else:
                    s -= w * 0.8
    dn = str(p.get("device_name") or "")
    want = [f for f in FORMS if f in name]
    if want:
        s += 1.0 if any(f in dn for f in want) else (-1.2 if any(f in dn for f in FORMS) else 0)
    if re.search(r"100V|70V|定压", f"{name} {params}", re.I) and family_of(p) == "广播":
        s += 0.8
    # 接口类型：客户要 HDMI 的配 HD 卡，不配 DVI/SDI 卡（464 题方案部纠正）
    want = {k for k, rx in IFACES if re.search(rx, f"{name} {str(params)[:300]}", re.I)}
    have = {k for k, rx in IFACES if re.search(rx, dn, re.I)}
    if want and have:
        s += 1.0 if want & have else -1.0
    return s


def score(q, p, name=None):
    """name=客户设备名称（定类型、权重大）；q=名称+参数（排序参考）。参数里顺带提到的类型词不参与定类型。"""
    name = name if name is not None else q
    nb = bigrams(p.get("device_name"))
    s = 4.0 * len(bigrams(name) & nb) / (len(bigrams(name)) + 1)
    qb = bigrams(q)
    s += 1.0 * len(qb & bigrams(p["_text"][:1500])) / (len(qb) + 1)
    qt, pt = types_in(name), types_in(p.get("device_name"))
    if qt and pt:
        s += 1.5 if any(pt & SAME.get(t, {t}) for t in qt) else -1.5
    if re.search(r"IP|网络", q, re.I):
        s += 0.8 if re.search(r"IP|网络", p["_text"][:400], re.I) else -1.0
    if _conflict(name, str(p.get("device_name") or "")):   # 视频↔音频、有线↔无线…对不上
        s -= 2.0
    return s + spec_score(name, q[len(name):] if q.startswith(name) else q, p)


def brief(p):
    specs = "；".join(f"{x.get('key')}:{x.get('value')}" for x in (p.get("technical_specs") or [])[:4])
    st = "在售" if avail(p) else f"✗{p.get('status') or '不可配'}"
    pr = p.get("price") or "未定价"
    ser = series_of(p)
    lk = link_of(ser)
    cm = common_of(p.get("model"))
    star = f"｜★方案部常用（{cm['分组']}{('；' + cm['备注']) if cm.get('备注') else ''}）" if cm else ""
    return (f"{p.get('model')}｜{p.get('device_name')}{star}｜{st}｜单价 {pr}{('/' + p['unit']) if p.get('unit') else ''}"
            f"｜系列：{ser or '—'}{('（' + lk + '）') if lk else ''}｜{specs[:140]}")


def rank(name, params="", show_all=False, top=8, kb=None, section="", brand=""):
    """先判系统（system_of），再找型号：我们不做的系统直接提示留空；同系统的产品加分，差得远的减分。"""
    kb = kb or load_kb()
    q = f"{name} {params}".strip()
    sysn, ours, why = system_of(name, params, section, brand)
    if not ours and not show_all:
        return [], f"⛔ 这行属于「{sysn}」（依据：{why}），我司不做 → 型号留空，写「我司无对应产品」"
    pool = [p for p in kb if show_all or (avail(p) and not is_old(p.get("model")))]   # 旧系列禁用（方案部 2026-10-26）
    same = [p for p in pool if type_ok(name, p)]
    note = "" if same else "⚠ 资料库里没有同类型的在售产品，下面是名字/参数最接近的，很可能应写「我司无对应产品」"
    if sysn:
        note = (f"判定系统：{sysn}（{why}）" + ("；" + note if note else ""))
    weak = why.startswith("分组")
    exp, nproj = experience(name)
    share = {m: k / nproj for m, k, _ in exp} if nproj else {}
    # 会议类的连接方式看客户参数：写了射频/MHz/U段 → U段系列；写了 5G/WiFi → 5G 系列（1013：客户写 610-690MHz，配了 5G 系列）
    pw = str(params)
    want_link = ("无线U段" if re.search(r"MHz|射频|UHF|U段", pw, re.I) and not re.search(r"5G|5\.8G|WIFI", pw, re.I)
                 else "5G WiFi" if re.search(r"5G|5\.8G|WIFI", pw, re.I) and not re.search(r"MHz|UHF|U段", pw, re.I) else "")
    if not re.search(r"会议|主席|代表|列席", name):
        want_link = ""
    amp_pro = is_pro_amp_row(name, section)
    # 先在方案部常用表里找（2026-10-29 用户：方案部都用这张表）：非广播的行，同类型的表内设备大幅加分；表里没有同类的再从全库找
    # 只在明确是会议/扩声/中控…的行才先查表：系统没判出来、或在广播分组下、或参数写定压的，多半是广播（表里一个广播设备都没有），
    # 硬推会把草地音箱/10W 壁挂/调谐器推成会议设备（实测前 3 命中 51.7→47.5）
    table_first = (sysn not in ("", "广播", "辅材") and not weak and not re.search(r"广播|背景音乐", section)
                   and (amp_pro or not re.search(r"定压|70V|100V", str(params)[:300], re.I))   # 会议组功放写定压也是扩声功放
                   and not (re.search(r"壁挂|吸顶|天花|扬声器|喇叭", name) and (_name_w(name) or 99) <= 20))   # 10W 壁挂/吸顶 = 广播音箱
    same_ids = {id(p) for p in same}
    if table_first:
        note = "先用方案部常用表：★ 的优先（方案部做会议/扩声/中控等都用这张表），表里没有满足参数的再用全库" + ("；" + note if note else "")
    lf_want = _lf_size(params) if re.search(r"音箱|音柱|声柱|扬声器|低频|全频", name) else None
    case = same_case(name, params)                     # 历史上同样的清单行，方案部配的型号
    if case:
        note = (f"历史同样清单：方案部配过 {case[0]}（参数相似度 {case[1]:.2f}，{case[2]} 个项目）"
                + ("；" + note if note else ""))

    def sc(p):
        v = score(q, p, name)
        if lf_want:                                   # 音箱低音单元尺寸对得上的往前（不够的在 param_gaps 里已经压后）
            lf = _lf_size(_ptext(p))
            if lf and abs(lf - lf_want) <= 0.3:
                v += 1.5
        lk = link_of(series_of(p)) if want_link else ""
        if lk:
            v += 0.8 if lk == want_link else -0.8
        if PRICE_W and sysn and str(p.get("price") or "").replace(".", "", 1).isdigit():   # 系统都判不出时不拿价格挑
            # 分数相近时便宜的往前（方案部同档位常配便宜的）；封顶 0.35 分，只当「差不多时挑便宜」，不能把不对的类型推上来
            v -= min(0.35, PRICE_W * max(0.0, math.log10(float(p["price"]) / 100)))
        if sysn and not weak:
            v += 1.0 if family_of(p) == sysn else (-1.5 if far_apart(sysn, p, why) else 0)
        elif sysn and weak and family_of(p) == sysn:
            v += 0.6        # 只从分组推出来的系统（「公共广播系统」下的时序器）：同类里优先配这个系统的（1032 方案部配 1003ATW 不配 SR258）
        if amp_pro and family_of(p) == "广播":
            v -= 3.0        # 会议室/扩声里的功放配专业功放，不配广播功放（2026-10-29 方案部：会议系统功放配成了天问广播 1650ARTW）
        if re.search(r"会议.{0,4}主机|会议系统主机", name) and "升降" in series_of(p):
            v -= 2.5        # HT-PMS-CH700T 叫「视像跟踪会议主机」，其实是无纸化升降器系列，下面没有会议单元（agent 反馈，555 题）
        if _conflict(name, str(p.get("device_name") or "")):
            return v                                       # 关键字对不上（视频↔音频）：经验不加分
        v += 2.0 * share.get(p.get("model"), 0)            # 方案部常配的往前排
        if table_first and common_of(p.get("model")) and family_of(p) == sysn and (id(p) in same_ids or share.get(p.get("model"))):
            v += COMMON_W       # 方案部做方案就用常用表（非广播答案 95% 在表里）：同类型的表内设备先上；广播类表里没有，不加
        if case and p.get("model") == case[0]:
            v += 3.0 if case[1] >= 0.7 else 1.5           # 历史同样清单的配法：≥0.7 基本就是它（实测 95%）
        if any(lv == "❌" for lv, _ in param_gaps(name, params, p)):
            v -= 4.0                                       # 客户写明的数字达不到（4K/瓦数/路数/寸/变焦）：压到后面，经验加分也盖不过
        elif is_4k(f"{name} {params}") and is_4k(_ptext(p)):
            v += 1.0
        return v
    base = same or pool
    # 「IP网络多音源音频编解码器」「音频采集器」= 模拟转网络编码设备，方案部配 HT-PAE01-1013TW（41 个项目 + 300 题答案）
    if re.search(r"音频.{0,6}(编解码|编码|采集)|多音源.{0,6}编|音频采集", name) and not re.search(r"视频", name):
        enc = [p for p in pool if p.get("model") == "HT-PAE01-1013TW"]
        base = enc + [p for p in base if p not in enc]
        share = dict(share, **{"HT-PAE01-1013TW": max(share.get("HT-PAE01-1013TW", 0), 1.0)})
    # 「CVS信号综合交互管理主机」「分布式系统主机」= 分布式系统服务端主机（655 方案部配 EC-ALLINONE「智能控制交互系统」）
    if re.search(r"(信号|综合).{0,6}交互.{0,6}主机|交互管理主机|分布式.{0,6}(管理)?主机|分布式.{0,4}服务器", name) and "软件" not in name:
        srv = [p for p in pool if "分布式系统-服务端" in series_of(p) and re.search(r"主机|交互系统", str(p.get("device_name")))]
        base = srv + [p for p in base if p not in srv]
        share = dict(share, **{p.get("model"): max(share.get(p.get("model"), 0), 1.0) for p in srv})
    have = {id(p) for p in base}
    base = base + [p for p in pool if (p.get("model") in share or (case and p.get("model") == case[0])) and id(p) not in have]   # 方案部配过但没搜到的也补进来
    cands = sorted(((sc(p), p) for p in base), key=lambda x: -x[0])[:top]
    return cands, note


def cmd_search(name, params="", show_all=False, top=8, section="", brand=""):
    cands, note = rank(name, params, show_all, top, section=section, brand=brand)
    q = f"{name} {params}".strip()
    if not cands and note.startswith("⛔"):
        print(f"「{q[:60]}」：{note}")
        return
    print(f"「{q[:60]}」的候选（同类型、在售，按相似度；方案部常配的排前面）：")
    if note:
        print("  " + note)
    el = exp_line(name)
    if el:
        print("  " + el)
    for i, (s, p) in enumerate(cands, 1):
        print(f"  {i}. {brief(p)}")
    print("→ 只能从候选里选；都不合适就写「我司无对应产品」，**不要自己编型号**。")


def cmd_plan(path, top=3):
    """比单第一步：把客户清单逐行过一遍——这行属于哪个系统（依据分组标题/品牌/名称）、我们做不做、前几个候选。
    之后逐行定型号时以这张底稿为准；「⛔ 我司不做」的行直接留空写「我司无对应产品」。"""
    kb = load_kb()
    rows = read_rows(path)
    if not rows:
        print(f"✗ {path}：没读到清单行（找不到「名称」列），请用 search 逐行查")
        return
    print(f"{path}：{len(rows)} 行（[表名]第几行｜分组｜名称｜品牌 → 判定）")
    for r in rows:
        cands, note = rank(r["name"], r["params"], top=top, kb=kb, section=r["section"], brand=r["brand"])
        head = f"[{r['sheet']}]第{r['row']}行｜{r['section'][:12] or '—'}｜{r['name'][:20]}｜{r['brand'][:8] or '—'}"
        if not cands:
            print(f"{head} → {note}")
            continue
        print(f"{head} → {note.split('；')[0] or '系统看不出，按名称找'}")
        el = exp_line(r["name"])
        if el:
            print("    " + el)
        for i, (_, p) in enumerate(cands, 1):
            print(f"    {i}. {p.get('model')}｜{p.get('device_name')}｜单价 {p.get('price') or '未定价'}｜{series_of(p)}")
    print("→ 候选只是起点：逐行再用 search/model 核对参数；会议主机/矩阵等要用 kit 配同系列；不做的系统留空。")


def cmd_kit(m):
    """列出和某个主机同系列的全部在售产品（单元/扩展/AP/充电箱/线材…）。配套设备只能从这里选。"""
    kb = load_kb()
    ps = find_model(index(kb), m)
    if not ps:
        print(f"✗ {m}：资料库里没有这个型号")
        return
    host = ps[0]
    ser = series_of(host)
    print(f"{m}「{host.get('device_name')}」的系列：{ser}（{link_of(ser) or '—'}）")
    if not ser:
        print("  这个型号在资料库里没有系列信息，配套请人工确认")
        return
    sers = [ser] + KIT_EXTRA.get(ser, [])
    same = [p for p in kb if series_of(p) in sers and avail(p) and p is not host]
    if len(sers) > 1:
        print(f"（并含配套系列：{'、'.join(sers[1:])}——矩阵主机是空壳，必须配输入卡和输出卡）")
    print(f"同系列在售 {len(same)} 个（配套设备只能从这里选；系列不同的一律不能混配）：")
    for p in sorted(same, key=lambda x: str(x.get("device_name"))):
        print("  - " + brief(p))


def cmd_model(m):
    kb = load_kb()
    ps = find_model(index(kb), m)
    if not ps:
        print(f"✗ {m}：资料库里没有这个型号（可能是编的或写错了），用 search 找真实型号")
        return
    for p in ps:
        print(brief(p))
        for x in (p.get("full_official_specs") or [])[:30]:
            print("   ", x.get("full_text", "")[:200])


# ------------------------------------------------------------------ check
def watts_of(text):
    """从名称/参数里取单只音箱功率、功放额定功率（取最大的 W 值）"""
    ws = [float(v) for v in re.findall(r"(\d+(?:\.\d+)?)\s*W(?![a-z])", str(text or ""), re.I)]
    return max(ws) if ws else None


def qty_of(v):
    m = re.search(r"\d+(\.\d+)?", str(v or ""))
    return float(m.group()) if m else None


def find_header(rows):
    for i, r in enumerate(rows[:15]):
        cells = [str(v or "") for v in r]
        # 比单表左边是客户清单（客户的「规格型号」是别家的），右边才是我司 → 优先带 供应商/匹配/我司/配置 的列，否则取最右边的型号列
        mcs = [j for j, c in enumerate(cells) if re.search(r"型号", c) and len(c) < 16 and "客户" not in c and "原" not in c]
        pref = [j for j in mcs if re.search(r"供应商|匹配|我司|配置", cells[j])]
        mc = (pref or mcs or [None])[-1]
        if mc is not None:
            find = lambda pat: next((j for j, c in enumerate(cells) if re.search(pat, c) and len(c) < 16), None)  # noqa: E731
            return i, {"model": mc, "name": find(r"设备名称|项目名称|产品名称|品名|名称|^产品$"), "qty": find(r"数量"),
                       "price": find(r"^单价|含税单价|销售单价"), "params": next((j for j, c in enumerate(cells) if len(c) < 16 and re.search(
                       r"客户.*参数|技术参数|规格要求|招标参数|主要参数|参数要求|性能指标|规格参数|^参数$|^规格$|指标要求", c)
                       and not re.search(r"我司|供应商|不满足|偏离", c)), None),   # 「主要参数」以前没认，客户要 4K 也查不到
                       "note": find(r"备注|说明"), "ourname": find(r"我司.*名称|匹配.*名称")}
    return None, None


def guess_cols(rows, h, col):
    """客户清单表头不全（比单表第一行就是「五、会议系统」，表头只有右边我司那几列）→ 按左边各列的内容猜：
    参数=平均最长的文字列，名称=短中文列，数量=最右的纯数字列（序号列是第一个数字列，不算）。2026-10-29：没认出名称列，整表 check 什么都没查"""
    right = min(j for j in (col["model"], col.get("note")) if j is not None)
    right = min([right] + [j for j, c in enumerate(rows[h]) if c and re.search(r"品牌", str(c))])
    stats = {}
    for r in rows[h + 1:h + 80]:
        for j in range(min(right, len(r))):
            v = str(r[j]).strip() if r[j] is not None else ""
            if v:
                st = stats.setdefault(j, {"n": 0, "len": 0, "cn": 0, "num": 0})
                st["n"] += 1
                st["len"] += len(v)
                st["cn"] += len(re.findall(r"[一-龥]", v)) >= 2
                st["num"] += bool(re.fullmatch(r"\d+(\.\d+)?", v))
    if not stats:
        return col
    avg = {j: st["len"] / st["n"] for j, st in stats.items()}
    col = dict(col)
    if col.get("params") is None:
        j = max(avg, key=avg.get)
        if avg[j] >= 12:
            col["params"] = j
    if col.get("name") is None:
        cn = [j for j, st in stats.items() if j != col.get("params") and st["cn"] >= st["n"] * 0.6 and avg[j] <= 16 and st["n"] >= 3]
        if cn:
            col["name"] = max(cn, key=lambda j: (stats[j]["cn"], -j))
    if col.get("qty") is None:
        nums = sorted(j for j, st in stats.items() if st["num"] >= st["n"] * 0.8 and st["n"] >= 3)
        if len(nums) >= 2:
            col["qty"] = nums[-1]
    return col


def cmd_check(path, room=False):
    import openpyxl
    kb = load_kb()
    idx = index(kb)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    issues = []
    stat = Counter()
    wb_ip, wb_analog, wb_enc = False, [], False     # 整张清单：有没有 IP 广播 / 模拟音源 / 编码设备
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        h, col = find_header(rows)
        if h is None:
            continue
        if col["name"] is None or col["params"] is None:
            col = guess_cols(rows, h, col)
        load = 0.0
        amp = 0.0
        groups = {}                 # 分区 → [音箱W, 功放W]
        self_pw, sub = 0, []        # 网络音箱台数；副音箱 (分区, W)
        conf = []                   # 会议系统：(行位置, 主机/单元, 系列, 型号)
        matrix = {"host": [], "in": 0, "out": 0}   # 混合矩阵主机是空壳，必须配输入卡+输出卡
        prev = None                                # 上一行配的产品：墙面面板跟紧挨着的主机配套（中控主机 / 音频处理器）
        grp = next((str(c).strip()[:30] for r in rows[:h] for c in r
                    if c and SECTION.search(str(c)) and len(str(c).strip()) <= 30), "（未分区）")
        bcol = next((j for j, c in enumerate(rows[h]) if c and re.search(r"品牌", str(c)) and not re.search(r"我司|供应商", str(c))), None)
        ncol = next((j for j, c in enumerate(rows[h]) if c and str(c).strip() == "备注"), None)   # 客户原表的备注（第一个；我司备注在右边）
        get = lambda r, k: (r[col[k]] if col[k] is not None and col[k] < len(r) else None)  # noqa: E731
        for ri, r in enumerate(rows[h + 1:], start=h + 2):
            name = str(get(r, "name") or "").strip()
            model = str(get(r, "model") or "").strip()
            note = str(get(r, "note") or "")
            if not name and not model:
                continue
            filled = [str(c).strip() for c in r if c not in (None, "") and str(c).strip()]
            if not model and (re.fullmatch(r"[一二三四五六七八九十]+[、.．].*", name) or (
                    len(filled) <= 2 and any(SECTION.search(c) and len(c) <= 30 and not re.fullmatch(r"[\d.]+", c) for c in filled))):
                grp = (name if re.fullmatch(r"[一二三四五六七八九十]+[、.．].*", name) else next(c for c in filled if SECTION.search(c)))[:30]
                continue                                   # 分组标题行（可能在序号列）：后面的设备归这个分区
            if re.fullmatch(r"合计|小计|总计.*", name) and not model:
                continue
            stat["行"] += 1
            brand = str(r[bcol]).strip() if bcol is not None and bcol < len(r) and r[bcol] else ""
            if re.search(r"供应商|HT", brand):              # 比单表右侧「品牌」是我司填的，不是客户线索
                brand = ""
            if not brand and ncol is not None and ncol < len(r):
                brand = _note_brand(r[ncol])
            sysn, ours, why = system_of(name, str(get(r, "params") or ""), "" if grp == "（未分区）" else grp, brand)
            where = f"[{ws.title}] 第{ri}行 {name[:18]}"
            if not model or model in ("/", "-", "—"):
                if re.search(r"无对应|不配|客户自备|非我司|暂无", note + model):
                    stat["标明无对应"] += 1
                    continue
                if not ours:                               # 我们不做的系统（LED屏/监控/电脑/施工…）留空是对的
                    stat["非主营留空"] += 1
                    continue
                # 技能规矩：非主营（监控、电脑、线缆、服务…）留空是对的。只有「我们确实有同类在售产品」才算漏配
                q = f"{name} {str(get(r, 'params') or '')[:80]}"
                best = max(((score(q, p, name), p) for p in kb if avail(p) and type_ok(name, p)), key=lambda x: x[0], default=(0, None))
                qt = types_in(name)
                same_type = best[1] is not None and qt and (types_in(best[1].get("device_name")) & set().union(*(SAME.get(t, {t}) for t in qt)))
                if NON_MAIN.search(name):
                    stat["非主营留空"] += 1
                elif best[1] is not None and _conflict(name, str(best[1].get("device_name") or "")):
                    stat["非主营留空"] += 1                # 最像的那个只差关键字（视频处理器↔音频处理器），不算漏配
                elif best[0] >= 2.2 and same_type and "京东" in str(best[1].get("price_source") or ""):
                    issues.append(("⚠", "可配外购辅材", where, f"可以配外购件 {best[1].get('model')}「{best[1].get('device_name')}」（京东参考价，非公司报价）"))
                elif best[0] >= 2.2 and same_type:
                    # 中文设备名重点在后几个字（立杆/机柜/天线…），只碰到一个类型词不够——看整个名称有多像
                    nb = bigrams(name)
                    sim = len(nb & bigrams(best[1].get("device_name"))) / (len(nb) or 1)
                    sug = f"{best[1].get('model')}「{best[1].get('device_name')}」"
                    exp_hit = experience(name, min_sim=1.0)      # 名称完全一样的历史行
                    # 「视频处理器」≈「音频处理器」只差一个字（二字组重合 0.75）——以前会误报 ❌。
                    # 只有方案部在别的项目里真配过（经验库 ≥2 个项目）才硬拦，否则只提醒
                    if sim >= 0.5 and len(name) >= 4 and exp_hit[1] >= 3 and exp_hit[0] and exp_hit[0][0][1] / exp_hit[1] >= 0.6:
                        m0 = exp_hit[0][0]
                        issues.append(("❌", "漏配", where, f"方案部在 {exp_hit[1]} 个项目里给类似的行配过 {m0[0]}（{m0[1]} 个项目），"
                                       f"这行却留空；工具建议 {sug}（用 search 确认）"))
                    elif sim >= 0.5 and len(name) >= 4:
                        issues.append(("⚠", "可能漏配", where, f"也许可以配 {sug}，请确认；不合适就写「我司无对应产品」"))
                    elif sim >= 0.25:
                        issues.append(("⚠", "可能漏配", where, f"也许可以配 {sug}，请确认；不合适就写「我司无对应产品」"))
                    else:
                        stat["非主营留空"] += 1
                else:
                    stat["非主营留空"] += 1
                continue
            ms = re.findall(r"[A-Z]{1,3}-[A-Z0-9][A-Z0-9\-\.（）()]{2,40}", model, re.I) or [model]
            for extra in ms[1:]:                   # 一格写了几个型号（「HT-DP206/HT-XXX」）：后面的也要查
                if not re.match(r"\s*HT-", extra, re.I):
                    continue
                eps = find_model(idx, extra)
                if is_old(extra):
                    issues.append(("❌", "旧系列禁用", where, f"{extra} 是旧系列（已废弃），换天问系列"))
                elif not eps:
                    issues.append(("❌", "型号查不到", where, f"{extra} 资料库里没有（编的或写错）→ 用 search 重新找"))
                elif not avail(eps[0]):
                    issues.append(("❌", "停单/不可配", where, f"{extra} 状态：{eps[0].get('status')}，换同类在售型号"))
            if is_old(ms[0]):
                issues.append(("❌", "旧系列禁用", where, f"{ms[0]} 是旧系列（已废弃），换天问系列；天问配不出就留空写「我司无对应产品」"))
                continue
            ps = find_model(idx, ms[0])
            if not ps:
                if not re.match(r"\s*HT-", ms[0], re.I):
                    stat["外购/非我司"] += 1        # 不是 HT- 开头：别家品牌或外购件，不归资料库管
                    continue
                if re.search(r"[一-龥]", model) and not re.search(r"HT-", model, re.I):
                    stat["外购/非我司"] += 1        # 写的是外购品牌或说明，不是我司型号
                    continue
                issues.append(("❌", "型号查不到", where, f"{ms[0]} 资料库里没有（编的或写错）→ 用 search 重新找"))
                continue
            exact = N(ms[0]) in idx or N(re.sub(r"[（(].*?[)）]", "", ms[0])) in idx
            nm, nk = N(ms[0]), N(ps[0].get("model"))
            if not exact and nm.startswith(nk):     # 比资料库型号多出一截（「HT-PAR01-1010TW-PRO」→ 1010TW）：多写的后缀多半是编的
                tail = nm[len(nk):]                  # 表里写得比资料库短的（「HD8*8」读成「HD8」）是读取截断，不报
                lv = "❌" if len(tail) > 2 else "⚠"
                issues.append((lv, "型号写法不规范", where, f"表里写 {ms[0]}，资料库只有 {ps[0].get('model')}——改成资料库里的规范型号"
                               + ("（多出来的后缀资料库里没有，多半是编的）" if lv == "❌" else "")))
            p = ps[0]
            if not avail(p):
                issues.append(("❌", "停单/不可配", where, f"{ms[0]} 状态：{p.get('status')}，换同类在售型号"))
                continue
            stat["型号有效"] += 1
            if is_ip_broadcast(p):
                wb_ip = True
            if p.get("model") == "HT-PAE01-1013TW" or re.search(r"采集|编码", str(p.get("device_name"))):
                wb_enc = True
            if is_analog_source(p):
                wb_analog.append(f"{p.get('model')}")
            for lv_, msg_ in param_gaps(name, str(get(r, "params") or ""), p):      # 客户写明的硬指标达不到
                issues.append((lv_, "参数不满足", where, f"{ms[0]}「{p.get('device_name')}」：{msg_}——换达标的型号（search 带上这条参数）"))
            # 先判系统：客户这行是我们不做的系统（LED屏/监控/电脑…），却配了我司产品 → ❌（我司产品本身就是这类的不算，如 LED 显示终端）
            # 2026-10-26 方案部纠正：诺瓦「拼接控制器」配成了分布式「4K30拼接输出节点」HT-2010
            if not ours and system_of(p.get("device_name"))[1]:
                issues.append(("❌", "非主营系统", where, f"这行是「{sysn}」（依据：{why}），我司不做；{ms[0]}「{p.get('device_name')}」"
                               "不能配 → 型号留空，写「我司无对应产品」"))
            elif ours and far_apart(sysn, p, why):
                issues.append(("⚠", "系统可能不符", where, f"这行看着是「{sysn}」（{why}），配的 {ms[0]}「{p.get('device_name')}」"
                               f"属于「{family_of(p) or '—'}」，请确认"))
            ser = series_of(p)                                                  # 系列/大类，如「无线会议/五、无线U段手拉手系列」
            dn0 = str(p.get("device_name"))
            if ser.endswith("混合矩阵") and "矩阵主机" in dn0:
                matrix["host"].append((where, ms[0]))
            if ser.endswith("混合矩阵板卡"):
                matrix["in"] += "输入卡" in dn0
                matrix["out"] += "输出卡" in dn0
            # 紧跟在中控主机后面的面板 = 中控面板，只能用主机同系列的（2026-10-29 方案部：DCS-A180 下一行配了音频处理器系列的
            # W86C 墙板，不配套）。跟在音频处理器后面的 W86C 是对的（983 方案部答案），不管
            if prev and re.search(r"面板|触控屏|触摸屏|控制屏", name) and not re.search(r"音量|灯光|调光", name) \
                    and series_of(prev[0]).startswith("中控矩阵分布式") and "中控主机" in str(prev[0].get("device_name")) \
                    and ser != series_of(prev[0]):
                issues.append(("❌", "面板与中控主机不配套", where, f"{ms[0]}「{dn0}」是「{ser}」的，上一行中控主机 {prev[1]} 是"
                               f"「{series_of(prev[0])}」——面板要配主机同系列的（kit {prev[1]} 里的触控面板，86 底盒用 4 寸 TC-400）"))
            prev = (p, ms[0])
            if family_of(p) == "广播" and re.search(r"功放", dn0) and is_pro_amp_row(name, "" if grp == "（未分区）" else grp):
                issues.append(("❌", "功放配成了广播功放", where, f"这行在会议/扩声系统里，{ms[0]}「{dn0}」是广播功放（{ser}）——"
                               "会议室扩声配专业功放（search 带上 --分组），按客户功率选"))
            if re.match(r"(有线|无线)会议", ser) and re.search(r"主机|单元|话筒", str(p.get("device_name"))):
                conf.append((where, "主机" if "主机" in str(p.get("device_name")) else "单元", ser, ms[0]))
            # 类型核对：客户要的大类和我们配的大类
            want = types_in(name) | types_in(str(get(r, "params") or "")[:60])
            have = types_in(p.get("device_name")) | types_in(p.get("series"))
            nb_, pb_ = bigrams(name), bigrams(p.get("device_name"))
            near_same = nb_ and len(nb_ & pb_) / len(nb_) >= 0.7          # 名称几乎一样就不比类型（输出卡配输出卡）
            if want and have and not near_same and not any(have & SAME.get(t, {t}) for t in types_in(name) or want):
                issues.append(("⚠", "类型可能不符", where, f"客户要「{name[:16]}」，配的是「{p.get('device_name')}」({ms[0]})"))
            if (re.search(r"IP|网络", name, re.I) and not re.search(r"线|水晶头|跳线|辅材", name)
                    and not re.search(r"IP|网络|POE", p["_text"][:600], re.I)):   # 网线/线材不查网络功能
                issues.append(("⚠", "缺网络功能", where, f"客户要 IP/网络设备，{ms[0]}「{p.get('device_name')}」看不出网络功能"))
            # 价格核对
            pv = get(r, "price")
            if pv not in (None, "") and p.get("price"):
                try:
                    if abs(float(str(pv).replace(",", "")) - float(p["price"])) > 0.5:
                        issues.append(("⚠", "单价与资料库不同", where, f"{ms[0]} 表里 {pv}，资料库 {p['price']}"))
                except ValueError:
                    pass
            # 功放/音箱功率累计（同一张表内）
            q = qty_of(get(r, "qty")) or 1
            dn = f"{p.get('device_name')} {name}"
            # 单台功率：名称里写的 → 参数里的「额定功率」→ 前 3 条参数里的瓦数（2026-10-28：1025TW/1035TW 名称不带瓦数曾被算成 0，
            # 整表误报「功放功率不够」；HT-960 曾取到最大功率 80W 而不是额定 60W）
            w = (_name_w(p.get("device_name")) or _rated_w(p.get("_text"))
                 or watts_of(" ".join(f"{x.get('value')}" for x in (p.get("technical_specs") or [])[:3])))
            if re.search(r"功放", dn) and w:
                w *= _amp_ch(p)                  # 多通道功放按总输出算（「200W四通道」= 800W）
                amp += w * q
                groups.setdefault(grp, [0.0, 0.0])[1] += w * q
            elif re.search(r"网络音箱|有源", f"{p.get('device_name')}"):
                self_pw += 1                     # 网络音箱自带功放，不算功放负载；它带的副音箱也不算
            elif re.search(r"副音箱", f"{p.get('device_name')}") and w:
                sub.append((grp, w * q))
            elif re.search(r"音箱|音柱|喇叭|扬声器", dn) and w and w <= 120:
                w = min(w, _rated_w(get(r, "params")) or w)   # 多档音箱按客户要的档位接（客户 6-10W，配 10/20W 两档的按 10W 算）
                load += w * q
                groups.setdefault(grp, [0.0, 0.0])[0] += w * q
        # 会议主机和会议单元必须同一系列（U段/5G WiFi/有线 互不兼容）。2026-10-26 实测：U段主机配了 5G WiFi 单元
        # 分级：连接方式不同（U段 / 5G WiFi / 有线）→ ❌ 必定接不上；同连接方式但系列不同 → ⚠ 请确认
        hosts = [c for c in conf if c[1] == "主机"]
        for where, kind, ser, mdl in conf:
            if kind != "单元" or not hosts or any(h[2] == ser for h in hosts):
                continue
            h = hosts[0]
            if link_of(ser) and link_of(h[2]) and link_of(ser) != link_of(h[2]):
                issues.append(("❌", "配套系列不一致", where,
                               f"{mdl} 是「{ser}」（{link_of(ser)}），会议主机 {h[3]} 是「{h[2]}」（{link_of(h[2])}）——"
                               f"连接方式不同，接不上；用 kit {h[3]} 选同系列单元"))
            else:
                issues.append(("⚠", "配套系列不同", where,
                               f"{mdl} 是「{ser}」，会议主机 {h[3]} 是「{h[2]}」——同为{link_of(ser) or '会议'}但系列不同，请确认能否互通，"
                               f"拿不准就用 kit {h[3]} 选同系列"))
        for where, mdl in matrix["host"]:
            lack = [x for x, n in (("输入卡", matrix["in"]), ("输出卡", matrix["out"])) if not n]
            if lack:
                issues.append(("❌", "矩阵缺板卡", where, f"{mdl} 是插卡式矩阵主机（空壳），没配{'、'.join(lack)}就不能用；用 kit {mdl} 选板卡，按信号路数配"))
        if not self_pw:                          # 没有网络音箱时，副音箱只能接功放
            for g, w_ in sub:
                load += w_
                groups.setdefault(g, [0.0, 0.0])[0] += w_
        for g, (gl, ga) in groups.items():
            if len(groups) > 1 and gl and ga and ga < gl * 1.5:
                issues.append(("⚠", "分区功放功率可能不够", f"[{ws.title}] {g}",
                               f"该分区音箱约 {gl:.0f}W，功放约 {ga:.0f}W，需 ≥{gl * 1.5:.0f}W"))
        if amp and load and amp < load * 1.5:
            issues.append(("⚠", "功放功率可能不够", f"[{ws.title}] 整表",
                           f"音箱合计约 {load:.0f}W，功放合计约 {amp:.0f}W，低于 1.5 倍余量（需 ≥{load * 1.5:.0f}W）；"
                           "不同瓦数/分区的音箱要分开配功放"))
    # 网络广播里配了 8017D/8118 这类模拟音频设备，就要配模拟转网络编码设备 HT-PAE01-1013TW；纯模拟广播不需要
    # （2026-10-27 方案部确认，客户没列要另外加）→ ❌
    if wb_ip and wb_analog and not wb_enc:
        issues.append(("❌", "缺模拟转网络编码设备", "整表", f"有 IP 网络广播设备，又有模拟音源（{'、'.join(wb_analog[:3])}），"
                       "没配音频采集/编码设备 → 加增加项 HT-PAE01-1013TW（网络音频电源管理器，模拟音频转网络）"))
    if room:
        # 按面积出的会议室扩声方案：固定项（技能 Step 3a + 「8118 仍按面积档标配」）。2026-10-26 实测漏了 8118
        # 历史配单里时序器/8118 偶有不配，只给 ⚠，不合适就在备注写理由
        blob = " ".join(str(c) for ws in openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets
                        for r in ws.iter_rows(values_only=True) for c in r if c is not None)
        for rx, what in ((r"HT-DP(206|408)", "音频处理器 HT-DP206"), (r"HT-DF1", "反馈抑制器 HT-DF1"),
                         (r"HT-SR358", "电源时序器 HT-SR358"), (r"HT-Z605", "无线手持话筒 HT-Z605-SC"),
                         (r"HT-8118", "广播会议话筒 HT-8118")):
            if not re.search(rx, blob):
                issues.append(("⚠", "会议室固定项缺", "整表", f"没配{what}（会议室扩声固定项），确实不需要就在备注写理由"))
    bad = sum(1 for x in issues if x[0] == "❌")
    print(f"自检 {path}：{stat['行']} 行，型号有效 {stat['型号有效']}，标明无对应 {stat['标明无对应']}，"
          f"❌ {bad} 处，⚠ {len(issues) - bad} 处")
    for lv, kind, where, msg in issues:
        print(f"  {lv} {kind}｜{where}｜{msg}")
    if bad:
        print("→ 有 ❌ 不能交付：改完重新生成表格，再跑一遍 check。⚠ 要么改，要么在备注里写清理由。")
    return issues, stat


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    cmd = sys.argv[1]
    opt = {}                 # --分组 X --品牌 Y
    rest = []
    it = iter(sys.argv[2:])
    for a in it:
        if a in ("--分组", "--品牌"):
            opt[a[2:]] = next(it, "")
        elif not a.startswith("--"):
            rest.append(a)
    args = rest
    arg = " ".join(args)
    if cmd == "search":      # search "设备名称" ["关键参数"] [--分组 "三、显示系统"] [--品牌 诺瓦]
        cmd_search(args[0], " ".join(args[1:]), show_all="--all" in sys.argv, section=opt.get("分组", ""), brand=opt.get("品牌", ""))
    elif cmd == "plan":      # plan 客户清单.xlsx → 逐行判系统 + 候选
        cmd_plan(arg)
    elif cmd == "kit":       # kit 主机型号 → 同系列配套清单
        cmd_kit(arg)
    elif cmd == "model":
        cmd_model(arg)
    elif cmd == "check":
        iss, _ = cmd_check(arg, room="--会议室" in sys.argv)
        sys.exit(1 if any(x[0] == "❌" for x in iss) else 0)
    else:
        print(__doc__)
        sys.exit(1)
