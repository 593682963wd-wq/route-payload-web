"""航线载量分析 · 王迪 · 技术支持 杨清云

单一 Streamlit 应用：网页版 / 本地版共用同一份代码。
"""
from __future__ import annotations

import base64
from collections import defaultdict
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from core.parser import FlightPlan, parse_txt
from core.word_writer import build_doc, _load_json
from core.docs_builder import build_manual_bytes, build_ppt_bytes

APP_VERSION = "V 1.1.0"
AUTHOR = "王迪"
TECH_SUPPORT = "杨清云"

st.set_page_config(
    page_title="航线载量分析系统",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
# 主题（与机场障碍物分析保持一致：暗蓝赛博风）
# ─────────────────────────────────────────
st.markdown(
    """
<style>
:root{
  --bg:#0a0e17; --panel:#0d1520; --panel-strong:#0d2137; --line:#1a3a5c;
  --accent:#4fc3f7; --accent-2:#80d8ff; --text:#c0d8f0; --muted:#6d93b2;
  --good:#4caf50; --warn:#ffb74d; --bad:#ef5350;
}
.stApp{
  background: radial-gradient(ellipse at top, #0d1d33 0%, var(--bg) 70%) fixed;
  color: var(--text);
  font-family: 'Menlo','SF Mono','Roboto Mono',monospace;
}
section[data-testid="stSidebar"]{
  background: var(--panel) !important;
  border-right: 1px solid var(--line);
}
section[data-testid="stSidebar"] *{ color: var(--text) !important; }
h1,h2,h3,h4{ color: var(--accent) !important; letter-spacing:.5px; }
hr{ border-color: var(--line) !important; }

/* 头部 */
.app-header{
  background: linear-gradient(135deg, #0d2137 0%, #0a1525 100%);
  border:1px solid var(--line); border-radius:12px;
  padding: 22px 28px; margin-bottom: 18px;
  box-shadow: 0 0 24px rgba(79,195,247,.08);
}
.app-header .title{ font-size: 1.9rem; font-weight: 700; color: var(--accent); margin:0; letter-spacing:1px; }
.app-header .subtitle{ color: var(--muted); margin-top:4px; font-size:.95rem; }
.badge-version{
  display:inline-block; padding:3px 10px; border-radius:4px;
  background: var(--panel-strong); border:1px solid var(--accent);
  color: var(--accent); font-size:.8rem; margin-left:10px;
}
table.credits{ margin-top:12px; border-collapse:collapse; }
table.credits td{
  padding: 4px 16px 4px 0; color: var(--muted); font-size:.85rem; border:none;
}
table.credits td b{ color: var(--text); }

/* 步骤卡片 */
.step-card{
  background: var(--panel); border:1px solid var(--line);
  border-left: 3px solid var(--accent);
  border-radius:8px; padding:14px 18px; margin: 10px 0;
}
.step-num{
  display:inline-flex; width:26px; height:26px; border-radius:50%;
  background: var(--accent); color:#0a0e17; font-weight:700;
  align-items:center; justify-content:center; margin-right:10px;
}
.step-title{ color: var(--accent); font-weight:600; font-size:1.05rem; }

/* 指标条 */
.metric-strip{ display:flex; gap:12px; margin: 10px 0 6px; flex-wrap:wrap; }
.metric-card{
  flex:1; min-width:140px;
  background: var(--panel-strong); border:1px solid var(--line);
  border-radius:8px; padding:12px 16px;
}
.metric-card .label{ color: var(--muted); font-size:.78rem; text-transform:uppercase; letter-spacing:1px; }
.metric-card .value{ color: var(--accent); font-size:1.7rem; font-weight:700; margin-top:2px; }

/* 功能高亮网格 */
.feature-grid{ display:grid; grid-template-columns: repeat(auto-fit,minmax(220px,1fr)); gap:12px; margin:6px 0 12px; }
.feature-card{
  background: var(--panel); border:1px solid var(--line); border-radius:8px;
  padding:14px 16px; transition: all .2s;
}
.feature-card:hover{ border-color: var(--accent); box-shadow: 0 0 16px rgba(79,195,247,.15); transform: translateY(-2px); }
.feature-card .icon{ font-size:1.6rem; margin-bottom:6px; }
.feature-card .title{ color: var(--accent); font-weight:600; font-size:1rem; }
.feature-card .desc{ color: var(--muted); font-size:.85rem; margin-top:4px; }

/* 按钮 */
.stButton > button{
  background: var(--panel-strong) !important;
  color: var(--accent) !important;
  border: 1px solid var(--accent) !important;
  border-radius:6px !important; font-weight:600 !important;
  padding: .6rem 1.4rem !important;
  transition: all .2s !important;
}
.stButton > button:hover{
  background: var(--accent) !important;
  color: #0a0e17 !important;
  box-shadow: 0 0 18px rgba(79,195,247,.5) !important;
}
.stButton > button[kind="primary"]{
  background: linear-gradient(135deg, var(--accent) 0%, var(--accent-2) 100%) !important;
  color:#0a0e17 !important; font-size:1.05rem !important;
  padding: .85rem 2rem !important;
  box-shadow: 0 0 14px rgba(79,195,247,.4) !important;
  border:none !important;
}

.stDownloadButton > button{
  background: var(--panel-strong) !important;
  color: var(--accent) !important;
  border: 1px solid var(--accent) !important;
  border-radius:6px !important; font-weight:600 !important;
  padding:.6rem 1.4rem !important;
}
.stDownloadButton > button:hover{
  background: var(--accent) !important; color:#0a0e17 !important;
}

/* 上传区 */
[data-testid="stFileUploader"] section{
  background: var(--panel) !important; border:1.5px dashed var(--accent) !important;
  border-radius:10px !important;
}
[data-testid="stFileUploader"] section *{ color: var(--text) !important; }

/* 表格 */
.stDataFrame{ border:1px solid var(--line) !important; border-radius:8px !important; }

/* 折叠面板 */
details{ background: var(--panel) !important; border:1px solid var(--line) !important; border-radius:8px !important; margin: 8px 0 !important; }
details summary{ color: var(--accent) !important; font-weight:600; padding: 10px 14px !important; }
.streamlit-expanderHeader{ color: var(--accent) !important; }

/* alert */
div[data-baseweb="notification"]{ border-radius:8px !important; }
</style>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────
# 头部
# ─────────────────────────────────────────
st.markdown(
    f"""
<div class="app-header">
  <div class="title">✈️ 航线载量分析系统 <span class="badge-version">{APP_VERSION}</span></div>
  <div class="subtitle">OFP 飞行计划批量解析 · 自动汇总航线载量 · 一键生成 Word 报告</div>
  <table class="credits">
    <tr>
      <td><b>作者</b>：{AUTHOR}</td>
      <td><b>技术支持</b>：{TECH_SUPPORT}</td>
      <td><b>所属</b>：湖南航空</td>
      <td><b>版本</b>：{APP_VERSION}</td>
    </tr>
  </table>
</div>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────
# 功能高亮
# ─────────────────────────────────────────
st.markdown(
    """
<div class="feature-grid">
  <div class="feature-card">
    <div class="icon">📥</div>
    <div class="title">批量解析</div>
    <div class="desc">单次可处理 500+ 份 OFP TXT，秒级响应</div>
  </div>
  <div class="feature-card">
    <div class="icon">🧭</div>
    <div class="title">智能分组</div>
    <div class="desc">自动按起飞-目的-线路（南/北/W线）汇总</div>
  </div>
  <div class="feature-card">
    <div class="icon">🛩️</div>
    <div class="title">机型自动排序</div>
    <div class="desc">A319-115 → A320-214W → A320-251 顺序展现</div>
  </div>
  <div class="feature-card">
    <div class="icon">📄</div>
    <div class="title">Word 一键导出</div>
    <div class="desc">纵向 A4，标题 + 副标题 + 完整带边框表格</div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────
# 侧边栏
# ─────────────────────────────────────────
airports = _load_json("airports.json")
aircraft = _load_json("aircraft.json")

with st.sidebar:
    st.markdown("### 📊 系统信息")
    st.metric("机场词典", f"{len(airports)} 项")
    st.metric("机型词典", f"{len(aircraft)} 项")
    st.markdown("---")
    st.markdown("### 📁 配置文件")
    st.code("config/airports.json\nconfig/aircraft.json", language="text")
    st.caption("如遇未识别的代号，请编辑上述文件后重启程序")
    st.markdown("---")
    st.markdown("### 📖 资料下载")
    try:
        st.download_button(
            "📕 使用说明书 (Word)",
            data=build_manual_bytes(),
            file_name=f"航线载量分析_使用说明书_{APP_VERSION}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.download_button(
            "📊 介绍 PPT",
            data=build_ppt_bytes(),
            file_name=f"航线载量分析_介绍_{APP_VERSION}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
    except Exception as e:
        st.caption(f"资料生成失败：{e}")

# ─────────────────────────────────────────
# 功能介绍 + 计算逻辑（折叠）
# ─────────────────────────────────────────
with st.expander("📖 点这里查看 · 系统功能介绍", expanded=False):
    st.markdown(
        """
本系统专为湖南航空放行签派工作设计，用来批量处理 OFP 飞行计划 TXT 文件、自动汇总成航线载量分析 Word 报告。

**核心特点**

- **批量处理**：一次性拖入数百份 TXT，秒级完成解析
- **航线归类**：相同的「起飞-目的-线路」自动汇总到同一大标题
- **机型分层**：每条航线下，按 A319-115 → A320-214W → A320-251 的顺序分别出表
- **零配置上手**：机场和机型词典已预置，未识别代号在 `config/*.json` 里随时补充
- **网页与本地共用**：手机/电脑浏览器打开网页即用，本地双击「启动.command」也能跑
"""
    )

with st.expander("🧮 点这里查看 · 字段抓取与计算逻辑", expanded=False):
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

# ─────────────────────────────────────────
# Step 1：上传
# ─────────────────────────────────────────
st.markdown(
    '<div class="step-card"><span class="step-num">1</span><span class="step-title">上传 OFP 飞行计划 TXT 文件</span></div>',
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "请将 TXT 文件拖到下方（支持一次性批量上传 500+ 份）",
    type=["txt"],
    accept_multiple_files=True,
    label_visibility="visible",
)

if not uploaded:
    st.info("👆 请上传 TXT 文件后开始解析。文件名格式示例：306C ZSWX-ZWTL S07.txt")
    # 示例预览
    with st.expander("👀 输出示例预览（点击展开）", expanded=False):
        st.markdown(
            """
导出的 Word 报告大致结构：

```
1. 无锡-吐鲁番（南线）
    湖南航空公司A319-115飞机航线及载量分析
    ┌────────────────────────────────────┐
    │ 月份 │ 起飞重量 │ 总加油量 │ ... │ 人数 │
    │  7   │  70000  │  17700  │ ... │ 130 │
    └────────────────────────────────────┘
    湖南航空公司A320-214W飞机航线及载量分析
    ┌────────────────────────────────────┐
    │ ...                                │
    └────────────────────────────────────┘

2. 无锡-吐鲁番（北线）
    ...
```
"""
        )
    st.stop()

# ─────────────────────────────────────────
# Step 2：解析
# ─────────────────────────────────────────
st.markdown(
    '<div class="step-card"><span class="step-num">2</span><span class="step-title">数据解析与预览</span></div>',
    unsafe_allow_html=True,
)

import tempfile

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

# 分组数（外层航线）
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
# Step 3：生成报告
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
        use_container_width=False,
    )

# ─────────────────────────────────────────
# 底部
# ─────────────────────────────────────────
st.markdown("---")
with st.expander("❓ 常见问题 FAQ"):
    st.markdown(
        """
**Q1：上传的文件名有特殊字符，会出错吗？**  
A：不会。系统按 `机号 起飞-目的 线路+月份` 三段解析，中间允许空格。

**Q2：未识别的机场或机型怎么办？**  
A：编辑 `config/airports.json` 或 `config/aircraft.json`，加入对应映射后重启程序即可。

**Q3：网页版和本地版有什么区别？**  
A：完全一样的代码和功能。网页版无需安装、随时打开；本地版双击「启动.command」即可，断网也能用。

**Q4：FLIGHT LEVEL 解析为什么有时候是后几行的数字？**  
A：OFP TXT 的换行有时把 FL 值挤到下一行；系统会从 `FLIGHT LEVEL` 关键词开始往后扫描 4 行，自动取最大 3 位 FL 值。
"""
    )

st.markdown(
    f"<div style='text-align:center; color:var(--muted); font-size:.8rem; padding: 12px 0;'>"
    f"航线载量分析系统 {APP_VERSION} · {AUTHOR} 出品 · 技术支持 {TECH_SUPPORT}"
    f"</div>",
    unsafe_allow_html=True,
)
"""
航线载量分析系统 — Web 版
Flight Route Payload Analysis System — Web Edition
作者: 王迪
"""
