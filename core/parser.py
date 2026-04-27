"""TXT 飞行计划解析。所有字段均从前 ~80 行内抓取（保险阈值），CPT 表用于扫描最大 FL。"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, asdict
from pathlib import Path

SUFFIX_MAP = {"S": "南线", "N": "北线", "W": "W线"}

# 各机型核定最大载客人数（公司标准，pax_count 不得超过此上限）
MAX_PAX_BY_TYPE = {
    "A319-115": 144,
    "A320-214": 180,
    "A320-214W": 186,
    "A320-251": 186,
    "A321-211": 220,
}

_CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"


def _load_aircraft_map() -> dict:
    p = _CONFIG_DIR / "aircraft.json"
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return {k: v for k, v in data.items() if not k.startswith("_")}


@dataclass
class FlightPlan:
    source_file: str
    month: int                     # 月份（文件名末尾两位）
    route_suffix: str              # "南线" / "北线" / "W线" / ""
    dep_icao: str                  # ZSWX
    arr_icao: str                  # ZWTL
    altn_icao: str                 # ZWKM
    altn_dist_nm: int              # 331
    aircraft_reg: str              # 306C（去掉 B 前缀）
    aircraft_type_code: str        # B306C（原始）
    tow_kg: int                    # 70000
    total_fuel_kg: int             # TOTL = 17700
    trip_fuel_kg: int              # DEST 行 FUEL = 12283
    trip_time: str                 # "04/40"
    trip_dist_nm: int              # 1818
    av_pld_kg: int                 # 11118
    avg_wind: str                  # "M038"
    extra_fuel_kg: int             # XTRA = 768
    target_arrival_kg: int         # TARGET ARRIVAL 5000
    max_fl: int                    # 整条计划最大飞行高度层 = 301
    pax_count: int                 # av_pld // 85

    @property
    def calc_alt_ft(self) -> int:
        return self.max_fl * 100

    def to_row(self) -> dict:
        return asdict(self)


# ---------- 文件名解析 ----------

_FILENAME_RE = re.compile(
    r"""^\s*
        (?P<reg>[A-Z0-9]+)\s+
        (?P<dep>[A-Z]{4})-(?P<arr>[A-Z]{4})
        (?:\s+(?P<tag>[A-Z]?)(?P<month>\d{1,2}))?
        \s*$""",
    re.VERBOSE,
)


def parse_filename(stem: str) -> dict:
    """从文件名（不含扩展名）解析 dep/arr/月份/线路后缀。

    例： "306C ZSWX-ZWTL S07" -> dep=ZSWX, arr=ZWTL, month=7, suffix=南线
    """
    m = _FILENAME_RE.match(stem)
    if not m:
        # 兜底：用宽松规则再试
        toks = stem.strip().split()
        if len(toks) >= 2 and "-" in toks[1]:
            dep, arr = toks[1].split("-", 1)
            tag = ""
            month = 0
            if len(toks) >= 3:
                tail = toks[-1]
                tm = re.match(r"^([A-Z]?)(\d{1,2})$", tail)
                if tm:
                    tag = tm.group(1)
                    month = int(tm.group(2))
            return {
                "dep": dep,
                "arr": arr[:4],
                "month": month,
                "suffix": SUFFIX_MAP.get(tag, ""),
            }
        raise ValueError(f"无法解析文件名: {stem}")

    return {
        "dep": m.group("dep"),
        "arr": m.group("arr"),
        "month": int(m.group("month") or 0),
        "suffix": SUFFIX_MAP.get((m.group("tag") or "").upper(), ""),
    }


# ---------- TXT 解析 ----------

def _read_text(path: Path) -> str:
    for enc in ("utf-8", "gb18030", "latin-1"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_bytes().decode("utf-8", errors="replace")


def _find_int(pattern: str, text: str, default: int = 0) -> int:
    m = re.search(pattern, text)
    return int(m.group(1)) if m else default


def _find_str(pattern: str, text: str, default: str = "") -> str:
    m = re.search(pattern, text)
    return m.group(1).strip() if m else default


def parse_txt(path: str | Path) -> FlightPlan:
    p = Path(path)
    text = _read_text(p)
    fname = parse_filename(p.stem)

    # 第三行的机型代号
    lines = text.splitlines()
    aircraft_type_code = ""
    for ln in lines[:6]:
        m = re.search(r"\bB(\d{3}[A-Z])\b", ln)
        if m:
            aircraft_type_code = "B" + m.group(1)
            break
    aircraft_reg = aircraft_type_code[1:] if aircraft_type_code else ""

    # TOW
    tow = _find_int(r"\bTOW\s+(\d+)", text)

    # DEST 行：DEST <ICAO> <FUEL> <HH/MM> <DIST> <ARRIVE>
    dest_m = re.search(
        r"^DEST\s+([A-Z]{4})\s+(\d+)\s+(\d{2}/\d{2})\s+(\d+)",
        text, re.MULTILINE,
    )
    if not dest_m:
        raise ValueError(f"未找到 DEST 行: {p.name}")
    dest_icao = dest_m.group(1)
    trip_fuel = int(dest_m.group(2))
    trip_time = dest_m.group(3)
    trip_dist = int(dest_m.group(4))

    # ALTN 行
    altn_m = re.search(
        r"^ALTN\s+([A-Z]{4})\s+\d+\s+\d{2}/\d{2}\s+(\d+)",
        text, re.MULTILINE,
    )
    altn_icao = altn_m.group(1) if altn_m else ""
    altn_dist = int(altn_m.group(2)) if altn_m else 0

    # XTRA / TOTL / TARGET ARRIVAL
    extra = _find_int(r"^XTRA\s+(\d+)", text) or _find_int(r"\bXTRA\s+(\d+)", text)
    totl = _find_int(r"^TOTL\s+(\d+)", text) or _find_int(r"\bTOTL\s+(\d+)", text)
    target_arr = _find_int(r"TARGET\s+ARRIVAL\s+(\d+)", text)

    # AV PLD
    av_pld = _find_int(r"\bAV\s+PLD\s+(\d+)", text)

    # ROUTE AVG WIND
    avg_wind = _find_str(r"ROUTE\s+AVG\s+WIND\s+([A-Z]?\d+)", text)

    # 计算高度：从 "FLIGHT LEVEL" 字段开始抓取，因字符可能被挤到下一行
    # （甚至中间夹空行），扫描其后 4 行内所有 3 位 FL 数字取最大值。
    max_fl = 0
    fl_match = re.search(r"FLIGHT\s+LEVEL\b(.*)", text)
    if fl_match:
        start = fl_match.start()
        tail = text[start:]
        scan = "\n".join(tail.splitlines()[:4])
        for m in re.finditer(r"\b(\d{3})\b", scan):
            fl = int(m.group(1))
            if 100 <= fl <= 450 and fl > max_fl:
                max_fl = fl
    if max_fl == 0:
        # 兜底：扫描 CPT 表
        cpt_re = re.compile(r"^\s*\S+\s+(\d{3})\s+(?:\.\.|[MP]\d{2})\s", re.MULTILINE)
        for m in cpt_re.finditer(text):
            fl = int(m.group(1))
            if fl > max_fl:
                max_fl = fl

    pax = av_pld // 85 if av_pld else 0

    # 按机型核定最大载客人数限制 pax_count
    ac_map = _load_aircraft_map()
    type_name = ac_map.get(aircraft_type_code, "")
    max_pax = MAX_PAX_BY_TYPE.get(type_name)
    if max_pax is not None and pax > max_pax:
        pax = max_pax

    return FlightPlan(
        source_file=p.name,
        month=fname["month"],
        route_suffix=fname["suffix"],
        dep_icao=fname["dep"] or dest_icao,
        arr_icao=fname["arr"] or dest_icao,
        altn_icao=altn_icao,
        altn_dist_nm=altn_dist,
        aircraft_reg=aircraft_reg,
        aircraft_type_code=aircraft_type_code,
        tow_kg=tow,
        total_fuel_kg=totl,
        trip_fuel_kg=trip_fuel,
        trip_time=trip_time,
        trip_dist_nm=trip_dist,
        av_pld_kg=av_pld,
        avg_wind=avg_wind,
        extra_fuel_kg=extra,
        target_arrival_kg=target_arr,
        max_fl=max_fl,
        pax_count=pax,
    )
