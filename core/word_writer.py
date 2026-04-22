"""把分组后的 FlightPlan 数据写成 Word 表格。
分组规则：(dep_icao, arr_icao, route_suffix, aircraft_type_code) 一组 -> 一张表 + 一个标题。
"""
from __future__ import annotations

import json
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from .parser import FlightPlan

# 表头列定义：(显示名, FlightPlan 字段名 或 None=月份)
COLUMNS = [
    ("月份", "month"),
    ("起飞重量\n（公斤）", "tow_kg"),
    ("总加油量\n（公斤）", "total_fuel_kg"),
    ("航程油量\n（公斤）", "trip_fuel_kg"),
    ("航程时间\n（时/分）", "trip_time"),
    ("航线距离\n（海里）", "trip_dist_nm"),
    ("最大业载\n（公斤）", "av_pld_kg"),
    ("航路平均风", "avg_wind"),
    ("额外油\n（公斤）", "extra_fuel_kg"),
    ("落地剩油\n（公斤）", "target_arrival_kg"),
    ("计算高度\n(FT)", "calc_alt_ft"),
    ("限重计算温度\n(℃)", "_zero"),
    ("人数", "pax_count"),
]

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"


def _load_json(name: str) -> dict:
    p = CONFIG_DIR / name
    if not p.exists():
        return {}
    data = json.loads(p.read_text(encoding="utf-8"))
    return {k: v for k, v in data.items() if not k.startswith("_")}


def _airport_name(icao: str, mapping: dict) -> str:
    return mapping.get(icao, icao)


def _aircraft_name(code: str, mapping: dict) -> str:
    return mapping.get(code, code or "未知机型")


def _set_cell_text(cell, text: str, bold: bool = False, size: int = 10):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    run = p.add_run(str(text))
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size)
    run.bold = bold


def _set_cell_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        from docx.oxml import OxmlElement
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    from docx.oxml import OxmlElement
    for edge in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "6")
        b.set(qn("w:color"), "000000")
        tc_borders.append(b)


def _value_for(fp: FlightPlan, attr: str):
    if attr == "_zero":
        return 0
    if attr == "calc_alt_ft":
        return fp.calc_alt_ft
    return getattr(fp, attr)


def build_doc(
    plans: Iterable[FlightPlan],
    airports: dict | None = None,
    aircraft: dict | None = None,
) -> Document:
    """把多份 FlightPlan 渲染成一个 docx，返回 Document。"""
    if airports is None:
        airports = _load_json("airports.json")
    if aircraft is None:
        aircraft = _load_json("aircraft.json")

    # 分组
    groups: dict[tuple, list[FlightPlan]] = defaultdict(list)
    order: list[tuple] = []
    for fp in plans:
        key = (fp.dep_icao, fp.arr_icao, fp.route_suffix, fp.aircraft_type_code)
        if key not in groups:
            order.append(key)
        groups[key].append(fp)

    doc = Document()
    # 页面横向、A4
    section = doc.sections[0]
    section.page_height, section.page_width = section.page_width, section.page_height
    section.left_margin = section.right_margin = Cm(1.5)
    section.top_margin = section.bottom_margin = Cm(1.5)

    # 设置默认中文字体
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    style.font.size = Pt(10.5)

    for idx, key in enumerate(order, 1):
        dep, arr, suffix, ac_code = key
        items = sorted(groups[key], key=lambda x: x.month)

        # 大标题：1. 无锡-吐鲁番（南线）
        title_text = f"{idx}. {_airport_name(dep, airports)}-{_airport_name(arr, airports)}"
        if suffix:
            title_text += f"（{suffix}）"
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = title_p.add_run(title_text)
        run.bold = True
        run.font.size = Pt(14)
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

        # 副标题（表头）：湖南航空公司A319-115飞机航线及载量分析
        sub = doc.add_paragraph()
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = sub.add_run(f"湖南航空公司{_aircraft_name(ac_code, aircraft)}飞机航线及载量分析")
        run.bold = True
        run.font.size = Pt(12)
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

        # 表格
        n_cols = len(COLUMNS)
        table = doc.add_table(rows=2 + len(items), cols=n_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # 第一行：信息行（合并）
        info_row = table.rows[0]
        info_row.cells[0].merge(info_row.cells[-1])
        sample = items[0]
        info = (
            f"机号：{sample.aircraft_reg}; "
            f"起飞机场：{sample.dep_icao}; "
            f"目的地机场：{sample.arr_icao}; "
            f"备降场：{sample.altn_icao}; "
            f"备降距离：{sample.altn_dist_nm}NM"
        )
        _set_cell_text(info_row.cells[0], info, bold=True, size=10)

        # 第二行：表头
        header_row = table.rows[1]
        for i, (name, _) in enumerate(COLUMNS):
            _set_cell_text(header_row.cells[i], name, bold=True, size=9)

        # 数据行
        for r, fp in enumerate(items, start=2):
            row = table.rows[r]
            for i, (_, attr) in enumerate(COLUMNS):
                _set_cell_text(row.cells[i], _value_for(fp, attr), size=10)

        # 边框
        for row in table.rows:
            for cell in row.cells:
                _set_cell_borders(cell)

        doc.add_paragraph()  # 段间距

    return doc


def build_bytes(plans: Iterable[FlightPlan], **kw) -> bytes:
    doc = build_doc(plans, **kw)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
