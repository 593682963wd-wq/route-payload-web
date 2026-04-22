"""
航线载量分析系统 — Web 版
Flight Route Payload Analysis System — Web Edition
作者: 王迪
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from core.parser import parse_txt, FlightPlan  # noqa: E402
from core.word_writer import build_bytes, _load_json  # noqa: E402

# ══════════════════════════════════════════════════════════════════
APP_VERSION = "V 1.0.1"
AUTHOR = "王迪"
TECH_SUPPORT = "邵小隆"
# ══════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="✈ 航线载量分析系统",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "About": f"航线载量分析系统 {APP_VERSION}\n\n作者: {AUTHOR}",
    },
)

# ═══════════════════════════════════════════════════════════
# 主题样式 — 与机场障碍物分析系统统一
# ═══════════════════════════════════════════════════════════
st.markdown(
    """
<style>
    :root {
        --bg: #0a0e17;
        --panel: #0d1520;
        --panel-strong: #0d2137;
        --line: #1a3a5c;
        --line-strong: #1a5276;
        --accent: #4fc3f7;
        --accent-soft: #80d0f8;
        --text: #c0d8f0;
        --muted: #6d93b2;
        --ok: #66bb6a;
        --warn: #ffb74d;
    }
    html, body, [class*="css"] {
        font-family: "Menlo", "Consolas", "SF Mono", "Monaco", monospace;
        color: var(--text);
    }
    .stApp {
        background: radial-gradient(circle at 100% -5%, #113052 0%, var(--bg) 35%);
    }
    .main .block-container {
        max-width: 1400px;
        padding-top: 1.1rem;
        padding-bottom: 1.2rem;
    }
    .main-header {
        position: relative;
        text-align: center;
        padding: 0.8rem 0 0.4rem 0;
        border-bottom: 1px solid var(--line);
        margin-bottom: 0.8rem;
    }
    .main-header h1 {
        color: var(--accent);
        margin: 0;
        letter-spacing: 2px;
        font-size: 1.75rem;
        font-weight: 700;
    }
    .main-header p {
        color: var(--muted);
        margin: 0.2rem 0 0 0;
        font-size: 0.8rem;
        letter-spacing: 1px;
    }
    .header-meta {
        position: absolute;
        top: 6px; right: 8px;
        text-align: right; line-height: 1.55;
    }
    .header-meta .badge-version {
        display: inline-block;
        background: transparent;
        color: #50fa7b;
        border: 1px solid #50fa7b;
        border-radius: 12px;
        padding: 2px 14px;
        font-size: 0.78rem;
        font-weight: 700;
        font-family: "Menlo", monospace;
        letter-spacing: 2px;
        margin-bottom: 12px;
    }
    .header-meta table.credits {
        margin-left: auto;
        border-collapse: collapse;
    }
    .header-meta table.credits td {
        color: #4fc3f7;
        font-size: 0.78rem;
        font-weight: 700;
        font-family: "Menlo", monospace;
        letter-spacing: 1px;
        padding: 2px 0;
    }
    .header-meta table.credits td.t-label { text-align: right; padding-right: 2px; }
    .header-meta table.credits td.t-colon { text-align: center; padding: 0 2px; }
    .header-meta table.credits td.t-name  { text-align: left;  padding-left: 2px; }
    .panel {
        background: var(--panel);
        border: 1px solid var(--line);
        border-radius: 6px;
        padding: 0.85rem 1rem;
        margin: 0.35rem 0;
    }
    .panel-title {
        color: var(--accent);
        font-weight: 700;
        margin-bottom: 0.4rem;
        letter-spacing: 0.4px;
    }
    .info-box, .warn-box, .ok-box {
        border-radius: 5px;
        padding: 0.7rem 0.9rem;
        margin: 0.35rem 0;
        border-left: 3px solid var(--accent);
        background: var(--panel-strong);
        font-size: 0.84rem;
    }
    .warn-box { border-left-color: var(--warn); color: #e0c890; background: #1b1a10; }
    .ok-box   { border-left-color: var(--ok);   color: #aadab0; background: #0f1f17; }
    .step-num {
        display: inline-block;
        background: var(--accent);
        color: var(--bg);
        width: 26px; height: 26px;
        border-radius: 999px;
        text-align: center; line-height: 26px;
        font-weight: 700; margin-right: 8px;
    }
    section[data-testid="stSidebar"] { border-right: 1px solid var(--line); }
    section[data-testid="stSidebar"] > div { background: var(--panel); }
    .stButton > button, .stDownloadButton > button {
        background: var(--panel-strong) !important;
        color: var(--accent) !important;
        border: 1px solid var(--line-strong) !important;
        font-family: "Menlo", monospace !important;
        font-weight: 700 !important;
        letter-spacing: 1px !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: var(--line) !important;
        border-color: var(--accent) !important;
        color: var(--accent-soft) !important;
    }
    .stDataFrame { border: 1px solid var(--line); border-radius: 4px; }
    div[data-testid="stFileUploader"] {
        background: var(--panel);
        border: 1px dashed var(--line-strong);
        border-radius: 6px;
        padding: 0.4rem;
    }
    .metric-strip {
        display: flex; gap: 14px; flex-wrap: wrap;
        margin: 0.3rem 0 0.6rem 0;
    }
    .metric-card {
        background: var(--panel-strong);
        border: 1px solid var(--line);
        border-left: 3px solid var(--accent);
        border-radius: 4px;
        padding: .55rem .85rem;
        min-width: 120px;
    }
    .metric-card .v {
        color: var(--accent); font-size: 1.15rem; font-weight: 700;
        font-family: "Menlo", monospace;
    }
    .metric-card .l {
        color: var(--muted); font-size: .72rem; letter-spacing: 1px;
    }
</style>
""",
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════════════════
# 顶部标题
# ═══════════════════════════════════════════════════════════
st.markdown(
    f"""
<div class="main-header">
  <h1>✈ 航线载量分析系统</h1>
  <p>FLIGHT ROUTE PAYLOAD ANALYSIS · 批量解析 OFP TXT，输出 Word 表格</p>
  <div class="header-meta">
    <div class="badge-version">{APP_VERSION}</div>
    <table class="credits">
      <tr><td class="t-label">作者</td><td class="t-colon">:</td><td class="t-name">{AUTHOR}</td></tr>
      <tr><td class="t-label">技术支持</td><td class="t-colon">:</td><td class="t-name">{TECH_SUPPORT}</td></tr>
    </table>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════════════════
# 侧边栏
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("<div class='panel-title'>⚙ 配置中心</div>", unsafe_allow_html=True)
    airports = _load_json("airports.json")
    aircraft = _load_json("aircraft.json")
    st.markdown(
        f"<div class='panel'>已映射机场 <b style='color:#4fc3f7'>{len(airports)}</b> 个<br>"
        f"已映射机型 <b style='color:#4fc3f7'>{len(aircraft)}</b> 个</div>",
        unsafe_allow_html=True,
    )
    st.markdown("<div class='panel-title' style='margin-top:.8rem'>📂 配置文件</div>", unsafe_allow_html=True)
    st.code(str(ROOT / "config"), language="text")

    with st.expander("📖 文件名规则", expanded=False):
        st.markdown(
            "- 例：`306C ZSWX-ZWTL S07.txt`\n"
            "- 末尾 `S/N/W` → 南线/北线/W线\n"
            "- 末尾两位数字 → 月份"
        )

# ═══════════════════════════════════════════════════════════
# 步骤 1：上传
# ═══════════════════════════════════════════════════════════
st.markdown(
    "<div class='panel'><span class='step-num'>1</span>"
    "<b style='color:#4fc3f7;letter-spacing:1px'>上传 OFP TXT 文件</b>"
    "<span style='color:#6d93b2;font-size:.78rem;margin-left:8px'>"
    "支持一次性 500+ 份批量处理</span></div>",
    unsafe_allow_html=True,
)
files = st.file_uploader(
    "拖入或选择 TXT",
    type=["txt"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if not files:
    st.markdown(
        "<div class='info-box'>📥 请在上方拖入或选择 OFP TXT 文件以开始解析。</div>",
        unsafe_allow_html=True,
    )
    st.stop()

# ═══════════════════════════════════════════════════════════
# 步骤 2：解析
# ═══════════════════════════════════════════════════════════
st.markdown(
    "<div class='panel' style='margin-top:1rem'><span class='step-num'>2</span>"
    "<b style='color:#4fc3f7;letter-spacing:1px'>解析数据</b></div>",
    unsafe_allow_html=True,
)

plans: list[FlightPlan] = []
errors: list[tuple[str, str]] = []

progress = st.progress(0.0)
status = st.empty()
for i, uf in enumerate(files, 1):
    tmp = ROOT / "输入" / uf.name
    tmp.parent.mkdir(exist_ok=True)
    tmp.write_bytes(uf.getbuffer())
    try:
        plans.append(parse_txt(tmp))
    except Exception as e:
        errors.append((uf.name, str(e)))
    progress.progress(i / len(files))
    status.caption(f"⏳ 解析中… {i}/{len(files)}: {uf.name}")
status.empty()
progress.empty()

groups_set = {(p.dep_icao, p.arr_icao, p.route_suffix, p.aircraft_type_code) for p in plans}
st.markdown(
    f"""
<div class='metric-strip'>
  <div class='metric-card'><div class='v'>{len(files)}</div><div class='l'>FILES</div></div>
  <div class='metric-card'><div class='v'>{len(plans)}</div><div class='l'>PARSED OK</div></div>
  <div class='metric-card'><div class='v'>{len(errors)}</div><div class='l'>FAILED</div></div>
  <div class='metric-card'><div class='v'>{len(groups_set)}</div><div class='l'>GROUPS</div></div>
</div>
""",
    unsafe_allow_html=True,
)

if errors:
    with st.expander(f"⚠ {len(errors)} 个文件解析失败", expanded=False):
        for name, err in errors:
            st.markdown(f"<div class='warn-box'><b>{name}</b><br>{err}</div>", unsafe_allow_html=True)

if not plans:
    st.stop()

rows = []
for fp in plans:
    rows.append({
        "文件": fp.source_file,
        "月": fp.month,
        "线路": fp.route_suffix,
        "起飞": fp.dep_icao,
        "目的": fp.arr_icao,
        "机型": fp.aircraft_type_code,
        "TOW": fp.tow_kg,
        "总加油": fp.total_fuel_kg,
        "航程油": fp.trip_fuel_kg,
        "时间": fp.trip_time,
        "距离": fp.trip_dist_nm,
        "AVPLD": fp.av_pld_kg,
        "平均风": fp.avg_wind,
        "额外油": fp.extra_fuel_kg,
        "落地剩油": fp.target_arrival_kg,
        "计算高度": fp.calc_alt_ft,
        "人数": fp.pax_count,
    })
df = pd.DataFrame(rows)
st.dataframe(df, use_container_width=True, hide_index=True, height=320)

# ═══════════════════════════════════════════════════════════
# 步骤 3：生成 Word
# ═══════════════════════════════════════════════════════════
st.markdown(
    "<div class='panel' style='margin-top:1rem'><span class='step-num'>3</span>"
    "<b style='color:#4fc3f7;letter-spacing:1px'>生成 Word 报告</b></div>",
    unsafe_allow_html=True,
)

c1, c2, c3 = st.columns([1, 1, 3])
with c1:
    if st.button("📄 生成报告", type="primary", use_container_width=True):
        with st.spinner("生成中…"):
            data = build_bytes(plans, airports=airports, aircraft=aircraft)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.session_state["docx_bytes"] = data
        st.session_state["docx_name"] = f"航线载量分析_{ts}.docx"

if "docx_bytes" in st.session_state:
    with c2:
        st.download_button(
            "⬇ 下载 Word",
            data=st.session_state["docx_bytes"],
            file_name=st.session_state["docx_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with c3:
        st.markdown(
            f"<div class='ok-box'>✔ 已生成 <b>{st.session_state['docx_name']}</b> · "
            f"{len(st.session_state['docx_bytes']) / 1024:.1f} KB</div>",
            unsafe_allow_html=True,
        )
