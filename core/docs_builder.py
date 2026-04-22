"""生成程序使用说明书 (DOCX) 和 PPT (PPTX) 资源文件。
首次访问时按需生成；每次启动会重新生成保证内容最新。
"""
from __future__ import annotations

from io import BytesIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGBColor
from pptx.util import Inches, Pt as PPt

# ─────────────────────────────────────────
# 共用：功能介绍 / 计算逻辑文本
# ─────────────────────────────────────────

INTRO_BULLETS = [
    "批量解析 OFP 飞行计划 TXT，一次可处理 500+ 份",
    "自动按航线（起飞-目的-线路）分组，同航线下按机型 A319 → A320-214W → A320-251 顺序汇总",
    "自动生成 Word 报告（纵向 A4，含大标题 + 副标题 + 完整表格 + 边框）",
    "机场四字码与机型代号映射可在 config 目录的 JSON 文件里随时增删",
    "支持网页版（任何人浏览器打开即用）与本地版（双击启动）",
]

LOGIC_ROWS = [
    ("大标题", "文件名机场四字码 + 末位字母→线路", "S=南线 / N=北线 / W=W线"),
    ("月份", "文件名末两位数字", "ZSWX-ZWTL S07 → 7"),
    ("机型副标题", "TXT 第 3 行机号代号 → 映射表", "B306C → A319-115"),
    ("起飞重量", "TXT 中 TOW 后数字", "70000"),
    ("总加油量", "TXT 中 TOTL 后数字", "17700"),
    ("航程油量", "DEST 行 FUEL 列", "12283"),
    ("航程时间", "DEST 行 TIME 列", "04/40"),
    ("航线距离", "DEST 行 DIST 列", "1818"),
    ("最大业载", "TXT 中 AV PLD 后数字", "11118"),
    ("航路平均风", "ROUTE AVG WIND 后字段", "M038"),
    ("额外油", "TXT 中 XTRA 后数字", "768"),
    ("落地剩油", "TARGET ARRIVAL 后数字", "5000"),
    ("计算高度", "FLIGHT LEVEL 字段后续 4 行内最大 FL × 100", "FL301 → 30100 FT"),
    ("限重计算温度", "固定值", "0"),
    ("人数", "AV PLD ÷ 85 向下取整", "11118 / 85 = 130"),
]


# ─────────────────────────────────────────
# DOCX 说明书
# ─────────────────────────────────────────

def _h(doc: Document, text: str, size: int, bold: bool = True, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _para(doc: Document, text: str, size: int = 10):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size)


def build_manual_bytes() -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    style.font.size = Pt(10.5)

    _h(doc, "航线载量分析系统 · 使用说明书", 18, color=RGBColor(0x1A, 0x52, 0x76))
    _para(doc, "湖南航空公司 · 王迪 出品 · 技术支持：杨清云")

    _h(doc, "一、功能介绍", 14, color=RGBColor(0x1A, 0x52, 0x76))
    for s in INTRO_BULLETS:
        _para(doc, "• " + s)

    _h(doc, "二、操作流程", 14, color=RGBColor(0x1A, 0x52, 0x76))
    for i, step in enumerate([
        "打开网页 https://wangdi-payload.streamlit.app  或本地双击「启动.command」",
        "在「上传 TXT」区域拖入或选择 OFP TXT 文件（支持一次性 500+ 份批量上传）",
        "等待解析完成，预览数据表，确认无误",
        "点击「生成报告」按钮，再点「下载 Word」即可导出",
    ], 1):
        _para(doc, f"{i}. {step}")

    _h(doc, "三、字段抓取与计算逻辑", 14, color=RGBColor(0x1A, 0x52, 0x76))
    table = doc.add_table(rows=1 + len(LOGIC_ROWS), cols=3)
    hdr = table.rows[0].cells
    for i, t in enumerate(["输出字段", "数据来源 / 算法", "示例"]):
        hdr[i].text = t
        for r in hdr[i].paragraphs[0].runs:
            r.bold = True
            r.font.size = Pt(10)
    for r, (a, b, c) in enumerate(LOGIC_ROWS, 1):
        cells = table.rows[r].cells
        cells[0].text = a
        cells[1].text = b
        cells[2].text = c
        for cell in cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "宋体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
                    run.font.size = Pt(9.5)
    table.style = "Light Grid Accent 1"

    _h(doc, "四、文件命名规则", 14, color=RGBColor(0x1A, 0x52, 0x76))
    _para(doc, "示例：306C ZSWX-ZWTL S07.txt")
    _para(doc, "• 第一段：机号（如 306C）")
    _para(doc, "• 第二段：起飞机场-目的机场（ICAO 四字码）")
    _para(doc, "• 第三段（可选）：线路标识 + 月份，例如 S07 表示南线 7 月")

    _h(doc, "五、输出说明", 14, color=RGBColor(0x1A, 0x52, 0x76))
    _para(doc, "• Word 文档为纵向 A4 排版")
    _para(doc, "• 同一航线（起飞-目的-线路相同）汇总在同一个大标题下")
    _para(doc, "• 大标题下按机型 A319-115 → A320-214W → A320-251 顺序排列各机型表格")

    _h(doc, "六、自定义机场 / 机型映射", 14, color=RGBColor(0x1A, 0x52, 0x76))
    _para(doc, "如果遇到未录入的机场代号或机型代号，请编辑：")
    _para(doc, "• config/airports.json — ICAO 四字码到中文名映射")
    _para(doc, "• config/aircraft.json — 机号代号到机型名映射")

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────
# PPTX 介绍幻灯片
# ─────────────────────────────────────────

def _add_title_slide(prs: Presentation, title: str, subtitle: str):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle


def _add_bullets_slide(prs: Presentation, title: str, bullets: list[str]):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title
    body = slide.placeholders[1].text_frame
    body.text = bullets[0]
    for b in bullets[1:]:
        p = body.add_paragraph()
        p.text = b
        p.level = 0


def _add_table_slide(prs: Presentation, title: str, header: list[str], rows: list[list[str]]):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    n_cols = len(header)
    n_rows = len(rows) + 1
    left = Inches(0.4); top = Inches(1.3)
    width = Inches(9.2); height = Inches(5.6)
    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table
    for i, h in enumerate(header):
        c = tbl.cell(0, i)
        c.text = h
        for p in c.text_frame.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.size = PPt(12)
    for r, row in enumerate(rows, 1):
        for i, val in enumerate(row):
            c = tbl.cell(r, i)
            c.text = str(val)
            for p in c.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = PPt(10)


def build_ppt_bytes() -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    _add_title_slide(
        prs,
        "航线载量分析系统",
        "湖南航空公司 · 王迪 出品\n技术支持：杨清云",
    )
    _add_bullets_slide(prs, "一、系统功能", INTRO_BULLETS)
    _add_bullets_slide(prs, "二、操作流程", [
        "Step 1  打开网页或本地启动",
        "Step 2  拖入 OFP TXT（支持 500+ 份批量）",
        "Step 3  确认解析数据预览",
        "Step 4  点击「生成报告」",
        "Step 5  下载 Word 文档",
    ])
    _add_table_slide(
        prs,
        "三、字段抓取与计算逻辑",
        ["输出字段", "数据来源 / 算法", "示例"],
        [list(r) for r in LOGIC_ROWS[:13]],
    )
    _add_bullets_slide(prs, "四、文件命名规则", [
        "示例：306C ZSWX-ZWTL S07.txt",
        "第一段：机号（如 306C）",
        "第二段：起飞-目的（ICAO 四字码）",
        "第三段：线路标识(S/N/W) + 月份(两位数字)",
    ])
    _add_bullets_slide(prs, "五、自定义映射", [
        "config/airports.json — ICAO 四字码到中文名",
        "config/aircraft.json — 机号代号到机型名",
        "未录入的代号会原样显示，编辑 JSON 即可补充",
    ])
    _add_bullets_slide(prs, "联系方式", [
        "作者：王迪",
        "技术支持：杨清云",
        "网页版：https://wangdi-payload.streamlit.app",
    ])

    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()
