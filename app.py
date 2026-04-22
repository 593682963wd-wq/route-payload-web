"""航线载量分析系统 — 网页版 / 本地版共用同一份代码。"""
from __future__ import annotations

import tempfile
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from core.parser import FlightPlan, parse_txt
from core.word_writer import build_doc, _load_json
from core.docs_builder import build_manual_bytes, build_ppt_bytes

APP_VERSION = "V 1.2.1"
AUTHOR = "王迪"
TECH_SUPPORT = "杨清云"

st.set_page_config(
    page_title="航线载量分析系统",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────
# 主题（暗蓝赛博风，统一风格）
# ─────────────────────────────────────────
st.markdown(
    """
<style>
:root{
  --bg:#0a0e17; --panel:#0d1520; --panel-strong:#0d2137;
  --line:#1a3a5c; --line-strong:#1a5276;
  --accent:#4fc3f7; --accent-2:#80d8ff;
  --text:#c0d8f0; --muted:#6d93b2;
  --ok:#66bb6a; --warn:#ffb74d; --bad:#ef5350;
}
html, body, [class*="css"]{
  font-family: "Menlo","Consolas","SF Mono","Monaco",monospace;
  color: var(--text);
}
.stApp{
  background: radial-gradient(circle at 100% -5%, #113052 0%, var(--bg) 35%);
}
.main .block-container{ max-width:1400px; padding-top:1.1rem; padding-bottom:1.4rem; }

/* 隐藏侧边栏控制按钮（侧边栏已为 collapsed） */
section[data-testid="stSidebar"]{ display:none !important; }
button[kind="header"][aria-label*="sidebar" i]{ display:none !important; }

/* ── 顶部栏 ── */
.main-header{
  position: relative;
  text-align: center;
  padding: 0.8rem 0 0.6rem 0;
  border-bottom: 1px solid var(--line);
  margin-bottom: 0.9rem;
}
.main-header h1{
  color: var(--accent);
  margin: 0;
  letter-spacing: 2px;
  font-size: 1.85rem;
  font-weight: 700;
}
.main-header p{
  color: var(--muted);
  margin: 0.25rem 0 0 0;
  font-size: 0.82rem;
  letter-spacing: 1px;
}
.header-meta{
  position: absolute;
  top: 6px;
  right: 8px;
  text-align: right;
  line-height: 1.55;
}
.header-meta .badge-version{
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
  margin-bottom: 10px;
}
.header-meta table.credits{ margin-left:auto; border-collapse: collapse; }
.header-meta table.credits td{
  color: #4fc3f7;
  font-size: 0.78rem;
  font-weight: 700;
  font-family: "Menlo", monospace;
  letter-spacing: 1px;
  padding: 2px 0;
}
.header-meta table.credits td.t-label{ text-align:right; padding-right:2px; }
.header-meta table.credits td.t-colon{ text-align:center; padding:0 2px; }
.header-meta table.credits td.t-name{ text-align:left; padding-left:2px; }

/* ── 上传区：超显眼大区块 ── */
.upload-hero{
  background: linear-gradient(135deg, rgba(79,195,247,.10) 0%, rgba(13,33,55,.6) 100%);
  border: 2px dashed var(--accent);
  border-radius: 14px;
  padding: 22px 28px 8px 28px;
  margin: 0 0 14px 0;
  box-shadow: 0 0 28px rgba(79,195,247,.15);
}
.upload-hero .hero-title{
  color: var(--accent);
  font-size: 1.45rem;
  font-weight: 700;
  letter-spacing: 1px;
  margin: 0 0 4px 0;
}
.upload-hero .hero-sub{
  color: var(--muted);
  font-size: 0.9rem;
  margin: 0 0 10px 0;
}

/* file_uploader 放大 */
[data-testid="stFileUploader"] section{
  background: var(--panel) !important;
  border: 1.5px dashed var(--accent) !important;
  border-radius: 10px !important;
  min-height: 130px;
  padding: 18px !important;
}
[data-testid="stFileUploader"] section *{ color: var(--text) !important; }
[data-testid="stFileUploader"] section button{
  background: var(--accent) !important;
  color: #0a0e17 !important;
  font-weight: 700 !important;
  border: none !important;
  padding: .55rem 1.4rem !important;
}

/* 步骤条 */
.step-card{
  background: var(--panel);
  border: 1px solid var(--line);
  border-left: 3px solid var(--accent);
  border-radius: 8px;
  padding: 10px 16px;
  margin: 14px 0 8px 0;
}
.step-num{
  display: inline-flex;
  width: 24px; height: 24px; border-radius: 50%;
  background: var(--accent); color: #0a0e17; font-weight: 700;
  align-items: center; justify-content: center;
  margin-right: 10px;
  font-size: 0.85rem;
}
.step-title{ color: var(--accent); font-weight: 600; font-size: 1rem; }

/* 指标卡片 */
.metric-strip{ display:flex; gap:12px; margin: 8px 0; flex-wrap:wrap; }
.metric-card{
  flex:1; min-width:140px;
  background: var(--panel-strong); border:1px solid var(--line);
  border-radius:8px; padding:12px 16px;
}
.metric-card .label{ color: var(--muted); font-size:.78rem; letter-spacing:1px; }
.metric-card .value{ color: var(--accent); font-size:1.7rem; font-weight:700; margin-top:2px; }

/* 按钮 */
.stButton > button, .stDownloadButton > button{
  background: var(--panel-strong) !important;
  color: var(--accent) !important;
  border: 1px solid var(--line-strong) !important;
  border-radius: 6px !important;
  font-weight: 600 !important;
  min-height: 2.4rem;
}
.stButton > button:hover, .stDownloadButton > button:hover{
  border-color: var(--accent) !important;
  background: #153d5e !important;
}
button[data-testid="stBaseButton-primary"]{
  background: linear-gradient(180deg, #1d6f3a 0%, #155a2c 100%) !important;
  color: #f0fff4 !important;
  border: 2px solid #50fa7b !important;
  font-weight: 700 !important;
  font-size: 1.05rem !important;
  min-height: 2.9rem !important;
  letter-spacing: 1px !important;
  box-shadow: 0 0 14px rgba(80,250,123,.3) !important;
}

h1,h2,h3,h4{ color: var(--accent) !important; letter-spacing: .5px; }
hr{ border-color: var(--line) !important; }

/* 表格 */
.stDataFrame{ border:1px solid var(--line) !important; border-radius:8px !important; }

/* 折叠面板 */
details, .streamlit-expander{
  background: var(--panel) !important;
  border:1px solid var(--line) !important;
  border-radius: 8px !important;
  margin: 8px 0 !important;
}
details summary, .streamlit-expanderHeader{
  color: var(--accent) !important;
  font-weight: 600 !important;
  padding: 10px 14px !important;
}
</style>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────
# 顶部标题（右上角徽标）
# ─────────────────────────────────────────
st.markdown(
    f"""
<div class="main-header">
    <div class="header-meta">
        <div class="badge-version">{APP_VERSION}</div>
        <table class="credits">
            <tr>
                <td class="t-label">系统开发</td>
                <td class="t-colon">：</td>
                <td class="t-name">{AUTHOR}</td>
            </tr>
            <tr>
                <td class="t-label">技术支持</td>
                <td class="t-colon">：</td>
                <td class="t-name">{TECH_SUPPORT}</td>
            </tr>
        </table>
    </div>
    <h1>✈ 航线载量分析系统</h1>
    <p>ROUTE PAYLOAD ANALYSIS SYSTEM · OFP 飞行计划批量解析</p>
</div>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────
# 顶部黄金区：上传区（最显眼）
# ─────────────────────────────────────────
st.markdown(
    """
<div class="upload-hero">
  <div class="hero-title">📥 第一步 · 在这里上传 OFP 飞行计划 TXT 文件</div>
  <div class="hero-sub">把文件直接拖到下方框内，或点击「Browse files」选择。支持一次性批量上传 500+ 份。</div>
</div>
""",
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    label="把 OFP TXT 文件拖到下方",
    type=["txt"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

# ─────────────────────────────────────────
# 数据加载
# ─────────────────────────────────────────
airports = _load_json("airports.json")
aircraft = _load_json("aircraft.json")

# ─────────────────────────────────────────
# 如果还没上传 → 显示快速提示并停止
# ─────────────────────────────────────────
if not uploaded:
    st.info("👆 请先在上方上传 TXT 文件。文件名示例：`306C ZSWX-ZWTL S07.txt`")

    # 快速三步指引
    st.markdown(
        """
<div style="display:grid; grid-template-columns: repeat(3,1fr); gap:12px; margin-top:10px;">
  <div class="step-card"><span class="step-num">1</span><span class="step-title">上传 TXT 文件</span><div style="color:var(--muted); font-size:.85rem; margin-top:6px;">把 OFP 飞行计划文件拖到上方</div></div>
  <div class="step-card"><span class="step-num">2</span><span class="step-title">查看解析结果</span><div style="color:var(--muted); font-size:.85rem; margin-top:6px;">系统自动按航线和机型分组</div></div>
  <div class="step-card"><span class="step-num">3</span><span class="step-title">下载 Word 报告</span><div style="color:var(--muted); font-size:.85rem; margin-top:6px;">点「生成报告」即可导出</div></div>
</div>
""",
        unsafe_allow_html=True,
    )

    # 文档下载（顶部黄金区之后，但仍较显眼）
    st.markdown("---")
    cdl1, cdl2, _ = st.columns([1.2, 1.2, 4])
    try:
        cdl1.download_button(
            "📕  使用说明书 (Word)",
            data=build_manual_bytes(),
            file_name=f"航线载量分析_使用说明书_{APP_VERSION}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        cdl2.download_button(
            "📊  介绍 PPT",
            data=build_ppt_bytes(),
            file_name=f"航线载量分析_介绍_{APP_VERSION}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
    except Exception as e:
        st.caption(f"资料生成失败：{e}")

    with st.expander("📖 系统功能介绍"):
        st.markdown(
            """
本系统用来批量处理 OFP 飞行计划 TXT，自动汇总成航线载量分析 Word 报告。

- **批量处理**：一次拖入数百份 TXT，秒级解析
- **航线归类**：相同的「起飞-目的-线路」自动汇总到同一大标题
- **机型分层**：每条航线下，按 A319-115 → A320-214W → A320-251 顺序出表
- **零配置**：机场/机型词典已预置，未识别代号在 `config/*.json` 里随时补充
- **网页 + 本地共用**：浏览器打开网页即用，本地双击「启动.command」也能跑
"""
        )
    with st.expander("🧮 字段抓取与计算逻辑"):
        st.markdown(
            """
| 输出字段 | 数据来源 / 算法 | 示例 |
|---|---|---|
| 大标题（航线） | 文件名机场四字码 + 末位字母→线路（S=南线 / N=北线 / W=W线） | ZSWX-ZWTL S07 → 无锡-吐鲁番（南线） |
| 月份 | 文件名末两位数字 | S07 → 7 |
| 机型副标题 | TXT 第 3 行机号代号 → `aircraft.json` 映射 | B306C → A319-115 |
| 起飞重量 | TXT 中 `TOW` 后数字 | 70000 |
| 总加油量 | TXT 中 `TOTL` 后数字 | 17700 |
| 航程油量 | DEST 行 FUEL 列 | 12283 |
| 航程时间 | DEST 行 TIME 列 | 04/40 |
| 航线距离 | DEST 行 DIST 列 | 1818 |
| 最大业载 | TXT 中 `AV PLD` 后数字 | 11118 |
| 航路平均风 | `ROUTE AVG WIND` 后字段 | M038 |
| 额外油 | TXT 中 `XTRA` 后数字 | 768 |
| 落地剩油 | `TARGET ARRIVAL` 后数字 | 5000 |
| 计算高度 | `FLIGHT LEVEL` 字段后 4 行内最大 FL × 100 | FL301 → 30100 FT |
| 限重计算温度 | 固定值 0 | 0 |
| 人数 | 最大业载 ÷ 85，向下取整 | 11118 ÷ 85 = 130 |
"""
        )
    with st.expander("❓ 常见问题 FAQ"):
        st.markdown(
            """
**Q1：上传的文件名有特殊字符，会出错吗？**  
A：不会。系统按 `机号 起飞-目的 线路+月份` 三段解析，中间允许空格。

**Q2：未识别的机场或机型怎么办？**  
A：编辑 `config/airports.json` 或 `config/aircraft.json`，加入对应映射后重启程序即可。

**Q3：网页版和本地版有什么区别？**  
A：完全一样的代码和功能。网页版无需安装；本地版双击「启动.command」断网也能用。

**Q4：FLIGHT LEVEL 解析为什么是后几行的数字？**  
A：OFP TXT 的换行有时把 FL 值挤到下一行；系统会从 `FLIGHT LEVEL` 关键词起往后扫描 4 行，自动取最大 3 位 FL 值。
"""
        )
    st.stop()

# ─────────────────────────────────────────
# Step 2 · 解析
# ─────────────────────────────────────────
st.markdown(
    '<div class="step-card"><span class="step-num">2</span><span class="step-title">数据解析与预览</span></div>',
    unsafe_allow_html=True,
)

plans: list[FlightPlan] = []
errors: list[tuple[str, str]] = []
with tempfile.TemporaryDirectory() as td:
    td_path = Path(td)
    for f in uploaded:
        try:
            tmp_path = td_path / f.name
            tmp_path.write_bytes(f.read())
            fp = parse_txt(tmp_path)
            plans.append(fp)
        except Exception as e:  # noqa: BLE001
            errors.append((f.name, str(e)))

route_keys = {(p.dep_icao, p.arr_icao, p.route_suffix) for p in plans}

st.markdown(
    f"""
<div class="metric-strip">
  <div class="metric-card"><div class="label">文件总数</div><div class="value">{len(uploaded)}</div></div>
  <div class="metric-card"><div class="label">解析成功</div><div class="value">{len(plans)}</div></div>
  <div class="metric-card"><div class="label">解析失败</div><div class="value">{len(errors)}</div></div>
  <div class="metric-card"><div class="label">航线分组</div><div class="value">{len(route_keys)}</div></div>
</div>
""",
    unsafe_allow_html=True,
)

if errors:
    with st.expander(f"⚠️ {len(errors)} 个文件解析失败 · 点击查看详情"):
        for name, msg in errors:
            st.error(f"**{name}** — {msg}")

if plans:
    rows = []
    for p in plans:
        rows.append({
            "文件": p.source_file,
            "月份": p.month,
            "机号": p.aircraft_reg,
            "机型": aircraft.get(p.aircraft_type_code, p.aircraft_type_code),
            "起飞机场": f"{p.dep_icao} {airports.get(p.dep_icao, '')}",
            "目的机场": f"{p.arr_icao} {airports.get(p.arr_icao, '')}",
            "线路": p.route_suffix or "-",
            "起飞重量": p.tow_kg,
            "总加油量": p.total_fuel_kg,
            "航程油量": p.trip_fuel_kg,
            "航程时间": p.trip_time,
            "航线距离": p.trip_dist_nm,
            "最大业载": p.av_pld_kg,
            "平均风": p.avg_wind,
            "额外油": p.extra_fuel_kg,
            "落地剩油": p.target_arrival_kg,
            "计算高度(FT)": p.calc_alt_ft,
            "人数": p.pax_count,
        })
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────
# Step 3 · 生成报告
# ─────────────────────────────────────────
st.markdown(
    '<div class="step-card"><span class="step-num">3</span><span class="step-title">生成 Word 报告</span></div>',
    unsafe_allow_html=True,
)

if not plans:
    st.warning("没有可用于生成报告的数据，请检查上传的文件。")
    st.stop()

col1, col2 = st.columns([1, 3])
with col1:
    gen = st.button("📄  生成报告", type="primary", use_container_width=True)

if gen:
    with st.spinner("正在生成 Word 报告..."):
        doc = build_doc(plans, airports=airports, aircraft=aircraft)
        buf = BytesIO()
        doc.save(buf)
        st.session_state["docx_bytes"] = buf.getvalue()
    st.success(f"✅ 报告已生成！共 {len(route_keys)} 条航线分组。")

if "docx_bytes" in st.session_state:
    st.download_button(
        "⬇️  下载 Word 报告",
        data=st.session_state["docx_bytes"],
        file_name=f"航线载量分析_{APP_VERSION}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

# ─────────────────────────────────────────
# 底部：资料下载 + 帮助
# ─────────────────────────────────────────
st.markdown("---")
b1, b2, _ = st.columns([1.2, 1.2, 4])
try:
    b1.download_button(
        "📕  使用说明书 (Word)",
        data=build_manual_bytes(),
        file_name=f"航线载量分析_使用说明书_{APP_VERSION}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
        key="dl_manual_bottom",
    )
    b2.download_button(
        "📊  介绍 PPT",
        data=build_ppt_bytes(),
        file_name=f"航线载量分析_介绍_{APP_VERSION}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        key="dl_ppt_bottom",
    )
except Exception as e:
    st.caption(f"资料生成失败：{e}")

st.markdown(
    f"<div style='text-align:center; color:var(--muted); font-size:.78rem; padding: 10px 0;'>"
    f"航线载量分析系统 {APP_VERSION} · 系统开发 {AUTHOR} · 技术支持 {TECH_SUPPORT}"
    f"</div>",
    unsafe_allow_html=True,
)
