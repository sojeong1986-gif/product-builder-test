import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(
    page_title="동업사 본드포워드 현황",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── 인쇄 JS: 페이지 최상단에 삽입 ─────────────────────────────────
st.markdown("""
<script>
function doPrint() {
    window.parent.print();
}
</script>
""", unsafe_allow_html=True)

st.markdown("""
<style>
.stApp { background-color: #ffffff; }
[data-testid="stAppViewContainer"] { background-color: #ffffff; }
[data-testid="stHeader"] { background-color: #ffffff; }
.page-title { font-size: 22px; font-weight: 700; color: #1a2340; margin: 0; }
.page-subtitle { font-size: 13px; color: #8a94a6; margin-top: 4px; }
.kpi-card { background: #f7f9fc; border: 1px solid #e8edf3; border-radius: 12px; padding: 20px 22px; }
.kpi-label { font-size: 12px; color: #8a94a6; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
.kpi-value { font-size: 26px; font-weight: 700; color: #1a2340; line-height: 1.1; }
.kpi-sub { font-size: 12px; color: #8a94a6; margin-top: 4px; }
.kpi-badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; margin-top: 6px; }
.badge-up { background: #e8f5e9; color: #2e7d32; }
.badge-down { background: #ffebee; color: #c62828; }
.section-title { font-size: 15px; font-weight: 700; color: #1a2340; margin-bottom: 14px; padding-bottom: 8px; border-bottom: 1px solid #e8edf3; }
.styled-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.styled-table thead tr { background: #f0f4f9; }
.styled-table th { padding: 10px 14px; text-align: right; font-weight: 600; color: #4a5568; border-bottom: 1px solid #e8edf3; white-space: nowrap; }
.styled-table th:first-child { text-align: left; }
.styled-table td { padding: 9px 14px; text-align: right; color: #2d3748; border-bottom: 1px solid #f0f4f9; white-space: nowrap; }
.styled-table td:first-child { text-align: left; font-weight: 600; color: #1a2340; }
.styled-table tr:hover { background: #f7f9fc; }
.styled-table tr.손보-total { background: #e3f2fd; font-weight: 700; color: #1565c0; }
.styled-table tr.손보-total td { color: #1565c0; }
.styled-table tr.생보-total { background: #f3e5f5; font-weight: 700; color: #6a1b9a; }
.styled-table tr.생보-total td { color: #6a1b9a; }
.styled-table tr.grand-total { background: #1a2340; color: white; font-weight: 700; }
.styled-table tr.grand-total td { color: white; border-bottom: none; }

/* 공유 버튼 */
.share-bar { display: flex; gap: 8px; align-items: center; margin-top: 6px; }
.share-btn {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 6px 12px; border-radius: 6px; font-size: 12px;
    font-weight: 600; text-decoration: none; cursor: pointer; border: none;
}
.btn-print { background: #1a2340; color: white; }
.btn-kakao { background: #FEE500; color: #3C1E1E; }
.btn-copy  { background: #f0f4f9; color: #4a5568; border: 1px solid #e8edf3; }

@media print { .no-print { display: none !important; } }
</style>
""", unsafe_allow_html=True)

# ── 엑셀 파싱 ──────────────────────────────────────────────────────
@st.cache_data
def load_data(path="bond_forward.xlsx"):
    df_raw = pd.read_excel(path, header=None, sheet_name=0)
    header_row = df_raw.iloc[6]
    periods = []
    for v in header_row:
        if isinstance(v, str) and "년" in v and "월" in v:
            p = v.strip().replace("\xa0", "")
            if p not in periods:
                periods.append(p)

    손보_names = ["메리츠화재","삼성화재","DB손해보험","현대해상","KB손해보험",
                  "한화손해보험","롯데손해보험","흥국화재","NH농협손해보험"]
    생보_names = ["삼성생명","한화생명","교보생명","신한라이프","농협생명",
                  "DB생명","미래에셋생명","동양생명","흥국생명"]
    all_companies = 손보_names + 생보_names

    companies = {}
    for _, row in df_raw.iterrows():
        name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if name not in all_companies:
            continue
        asset = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
        balances = []
        for v in row.iloc[3:3+len(periods)]:
            try:
                clean = str(v).replace(",","").replace("\xa0","").replace(" ","")
                balances.append(float(clean))
            except:
                balances.append(0.0)
        구분 = "손보" if name in 손보_names else "생보"
        companies[name] = {"구분": 구분, "자산총계": asset, "잔액": balances}
    return companies, periods

try:
    RAW, PERIODS = load_data()
    if not RAW:
        st.error("❌ 엑셀에서 데이터를 읽지 못했습니다.")
        st.stop()
except FileNotFoundError:
    st.error("❌ `bond_forward.xlsx` 파일을 찾을 수 없습니다.")
    st.stop()

def make_df():
    rows = []
    for co, d in RAW.items():
        for i, p in enumerate(PERIODS):
            if i >= len(d["잔액"]): continue
            bal = d["잔액"][i]
            rows.append({"회사": co, "구분": d["구분"], "자산총계": d["자산총계"],
                "기간": p, "잔액": bal,
                "비중": round(bal / d["자산총계"] * 100, 2) if d["자산총계"] else 0})
    return pd.DataFrame(rows)

df_long = make_df()

# ── 헤더 + 버튼 ────────────────────────────────────────────────────
col_title, col_btn = st.columns([5, 3])
with col_title:
    st.markdown("""
    <div>
        <div class="page-title">📊 동업사 본드포워드 현황</div>
        <div class="page-subtitle">보험업계 금리선도(Bond Forward) 미결제약정 현황 모니터링</div>
    </div>
    """, unsafe_allow_html=True)
with col_btn:
    page_url = "https://sojeong1986-gif-product-builder-test-bond-forward-app-8hzloc.streamlit.app"
    kakao_msg = "동업사 본드포워드 현황 대시보드"
    st.markdown(f"""
    <div class="share-bar no-print" style="justify-content:flex-end; padding-top:14px;">
        <button class="share-btn btn-print" onclick="window.parent.print()">🖨️ 인쇄 / PDF</button>
        <a class="share-btn btn-kakao"
           href="https://sharer.kakao.com/talk/friends/picker/link?app_key=KAKAO_APP_KEY&lang=ko&url={page_url}&text={kakao_msg}"
           target="_blank">
           💬 카카오톡
        </a>
        <button class="share-btn btn-copy"
           onclick="navigator.clipboard.writeText('{page_url}').then(()=>alert('링크 복사 완료!'))">
           🔗 링크복사
        </button>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ── 필터 ───────────────────────────────────────────────────────────
fc1, fc2, fc3 = st.columns([2, 2, 4])
with fc1:
    sel_period = st.selectbox("📅 조회 기준 시점", PERIODS, index=len(PERIODS)-1)
with fc2:
    sel_type = st.selectbox("🏢 구분", ["전체", "손보", "생보"])
with fc3:
    sel_companies = st.multiselect("🔍 회사 선택 (미선택 시 전체)", options=list(RAW.keys()), default=[])

손보_bal = df_long[(df_long["기간"]==sel_period)&(df_long["구분"]=="손보")]["잔액"].sum()
생보_bal = df_long[(df_long["기간"]==sel_period)&(df_long["구분"]=="생보")]["잔액"].sum()
total_bal = 손보_bal + 생보_bal

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

# ── KPI 카드 ────────────────────────────────────────────────────────
pi = PERIODS.index(sel_period)
prev_idx = pi - 1

def chg_badge(cur, prev):
    if prev <= 0: return ""
    diff = cur - prev
    pct = diff / prev * 100
    cls = "badge-up" if diff >= 0 else "badge-down"
    arrow = "▲" if diff >= 0 else "▼"
    return f'<span class="kpi-badge {cls}">{arrow} {abs(pct):.1f}%</span>'

if prev_idx >= 0:
    prev_total = df_long[df_long["기간"]==PERIODS[prev_idx]]["잔액"].sum()
    prev_손보  = df_long[(df_long["기간"]==PERIODS[prev_idx])&(df_long["구분"]=="손보")]["잔액"].sum()
    prev_생보  = df_long[(df_long["기간"]==PERIODS[prev_idx])&(df_long["구분"]=="생보")]["잔액"].sum()
else:
    prev_total = prev_손보 = prev_생보 = 0

top_co = df_long[df_long["기간"]==sel_period].sort_values("비중", ascending=False).iloc[0]

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">생/손보 합계</div>
        <div class="kpi-value">{total_bal/1000:,.1f}조</div>
        <div class="kpi-sub">{total_bal:,.0f} 억원</div>
        {chg_badge(total_bal, prev_total)}
    </div>""", unsafe_allow_html=True)
with k2:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">손보 합계</div>
        <div class="kpi-value">{손보_bal/1000:,.1f}조</div>
        <div class="kpi-sub">{손보_bal:,.0f} 억원 · {손보_bal/total_bal*100:.1f}%</div>
        {chg_badge(손보_bal, prev_손보)}
    </div>""", unsafe_allow_html=True)
with k3:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">생보 합계</div>
        <div class="kpi-value">{생보_bal/1000:,.1f}조</div>
        <div class="kpi-sub">{생보_bal:,.0f} 억원 · {생보_bal/total_bal*100:.1f}%</div>
        {chg_badge(생보_bal, prev_생보)}
    </div>""", unsafe_allow_html=True)
with k4:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">자산대비 비중 1위</div>
        <div class="kpi-value" style="font-size:20px">{top_co['회사']}</div>
        <div class="kpi-sub">자산대비 {top_co['비중']:.2f}%</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

# ── 차트 + 상세테이블 ─────────────────────────────────────────────
chart_col, table_col = st.columns([5, 5])

with chart_col:
    st.markdown('<div class="section-title">📈 본드포워드 잔액 추이</div>', unsafe_allow_html=True)
    손보_trend = df_long[df_long["구분"]=="손보"].groupby("기간")["잔액"].sum().reindex(PERIODS)
    생보_trend = df_long[df_long["구분"]=="생보"].groupby("기간")["잔액"].sum().reindex(PERIODS)

    fig = go.Figure()
    if sel_companies:
        colors = px.colors.qualitative.Plotly
        for i, co in enumerate(sel_companies):
            vals = [df_long[(df_long["회사"]==co)&(df_long["기간"]==p)]["잔액"].sum() for p in PERIODS]
            fig.add_trace(go.Scatter(x=PERIODS, y=vals, name=co, mode="lines+markers",
                line=dict(color=colors[i % len(colors)], width=2), marker=dict(size=6)))
    else:
        fig.add_trace(go.Bar(x=PERIODS, y=손보_trend.values, name="손보", marker_color="#1565c0", opacity=0.85))
        fig.add_trace(go.Bar(x=PERIODS, y=생보_trend.values, name="생보", marker_color="#6a1b9a", opacity=0.85))
        fig.add_trace(go.Scatter(x=PERIODS, y=(손보_trend+생보_trend).values, name="합계",
            mode="lines+markers", line=dict(color="#e53935", width=2.5), marker=dict(size=7)))

    fig.update_layout(
        barmode="stack", plot_bgcolor="white", paper_bgcolor="white",
        font=dict(size=12, color="#2d3748"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    bgcolor="rgba(255,255,255,0.9)", bordercolor="#e8edf3", borderwidth=1),
        xaxis=dict(showgrid=False, tickfont=dict(size=11)),
        yaxis=dict(showgrid=True, gridcolor="#f0f4f9", tickformat=",", ticksuffix=" 억", tickfont=dict(size=11)),
        margin=dict(l=10, r=10, t=40, b=10), height=340, hovermode="x unified",
    )
    for trace in fig.data:
        if trace.type == "bar":
            trace.marker.color = ["#ff7043" if p==sel_period else
                ("#1565c0" if trace.name=="손보" else "#6a1b9a") for p in PERIODS]
            trace.marker.opacity = 1
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title" style="margin-top:4px">🥧 생/손보 비중</div>', unsafe_allow_html=True)
    fig2 = go.Figure(go.Pie(
        labels=["손보","생보"], values=[손보_bal, 생보_bal], hole=0.6,
        marker_colors=["#1565c0","#6a1b9a"], textinfo="label+percent", textfont_size=13,
    ))
    fig2.update_layout(
        showlegend=False, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=10, r=10, t=10, b=10), height=210,
        annotations=[dict(text=f"{total_bal/10000:.1f}조", x=0.5, y=0.5,
                          font_size=18, font_color="#1a2340", showarrow=False)]
    )
    st.plotly_chart(fig2, use_container_width=True)

with table_col:
    st.markdown(f'<div class="section-title">📋 상세 현황 ({sel_period})</div>', unsafe_allow_html=True)

    def render_group(구분):
        cos = [c for c, d in RAW.items() if d["구분"]==구분]
        if sel_companies:
            cos = [c for c in cos if c in sel_companies]
        rows_html = ""
        grp_bal = 0
        for co in cos:
            d = RAW[co]
            cur  = d["잔액"][pi] if pi < len(d["잔액"]) else 0
            prev = d["잔액"][pi-1] if pi > 0 and (pi-1) < len(d["잔액"]) else cur
            diff = cur - prev
            arrow = "▲" if diff > 0 else ("▼" if diff < 0 else "–")
            color = "#2e7d32" if diff > 0 else ("#c62828" if diff < 0 else "#888")
            ratio = cur / d["자산총계"] * 100 if d["자산총계"] else 0
            rows_html += f"""<tr>
                <td>{co}</td><td>{d['자산총계']:,.0f}</td><td>{cur:,.0f}</td>
                <td><span style='color:{color};font-size:11px'>{arrow} {abs(diff):,.0f}</span></td>
                <td>{ratio:.2f}%</td>
            </tr>"""
            grp_bal += cur
        share = grp_bal / total_bal * 100 if total_bal else 0
        cls = "손보-total" if 구분=="손보" else "생보-total"
        label = "손보" if 구분=="손보" else "생보"
        rows_html += f'<tr class="{cls}"><td>{label} 합계</td><td></td><td>{grp_bal:,.0f}</td><td></td><td>{share:.1f}%</td></tr>'
        return rows_html

    bg1, fg1 = "#e3f2fd", "#1565c0"
    bg2, fg2 = "#f3e5f5", "#6a1b9a"
    손보_rows = f'<tr><td colspan="5" style="background:{bg1};color:{fg1};font-weight:700;padding:6px 14px;font-size:12px">▶ 손해보험</td></tr>' + render_group("손보")
    생보_rows = f'<tr><td colspan="5" style="background:{bg2};color:{fg2};font-weight:700;padding:6px 14px;font-size:12px">▶ 생명보험</td></tr>' + render_group("생보")
    grand = f'<tr class="grand-total"><td>생/손보 합계</td><td></td><td>{total_bal:,.0f}</td><td></td><td>100%</td></tr>'
    header = """<table class="styled-table"><thead><tr>
        <th style="text-align:left">회사</th><th>자산총계(억)</th>
        <th>잔액(억)</th><th>전분기比</th><th>자산비중</th>
    </tr></thead><tbody>"""
    body = {"전체": 손보_rows+생보_rows+grand, "손보": 손보_rows, "생보": 생보_rows}
    st.markdown(header + body[sel_type] + "</tbody></table>", unsafe_allow_html=True)

# ── 기간별 전체 추이 테이블 ──────────────────────────────────────────
st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
st.markdown('<div class="section-title">📅 기간별 잔액 추이 (억원)</div>', unsafe_allow_html=True)

cos_show = list(RAW.keys())
if sel_companies:
    cos_show = sel_companies
elif sel_type != "전체":
    cos_show = [c for c, d in RAW.items() if d["구분"]==sel_type]

pivot_rows = ""
grp_totals = {"손보": [0]*len(PERIODS), "생보": [0]*len(PERIODS)}

for 구분, bg, fg, label, cls in [
    ("손보","#e3f2fd","#1565c0","손해보험","손보-total"),
    ("생보","#f3e5f5","#6a1b9a","생명보험","생보-total")
]:
    cos_g = [c for c in cos_show if RAW[c]["구분"]==구분]
    if not cos_g: continue
    pivot_rows += f'<tr><td colspan="{len(PERIODS)+2}" style="background:{bg};color:{fg};font-weight:700;padding:6px 14px;font-size:12px">▶ {label}</td></tr>'
    grp_sums = [0]*len(PERIODS)
    for co in cos_g:
        d = RAW[co]
        cells = ""
        for i, p in enumerate(PERIODS):
            bal = d["잔액"][i] if i < len(d["잔액"]) else 0
            grp_sums[i] += bal
            hl = "background:#fff8e1;font-weight:700;" if p==sel_period else ""
            cells += f'<td style="{hl}">{bal:,.0f}</td>'
        pivot_rows += f'<tr><td>{co}</td><td>{d["자산총계"]:,.0f}</td>{cells}</tr>'

    # 소계 행 — 조회시점 강조 + 구분 색상
    gcells = ""
    for i, (p, s) in enumerate(zip(PERIODS, grp_sums)):
        hl = "background:#fff8e1;" if p==sel_period else f"background:{bg};"
        gcells += f'<td style="{hl}font-weight:700;color:{fg}">{s:,.0f}</td>'
        grp_totals[구분][i] = s
    pivot_rows += f'<tr class="{cls}"><td>{"손보" if 구분=="손보" else "생보"} 소계</td><td></td>{gcells}</tr>'

# 생/손보 합계 행
total_cells = ""
for i, p in enumerate(PERIODS):
    t = grp_totals["손보"][i] + grp_totals["생보"][i]
    hl = "background:#2d3f6b;" if p==sel_period else ""
    total_cells += f'<td style="{hl}font-weight:700;">{t:,.0f}</td>'
pivot_rows += f'<tr class="grand-total"><td>생/손보 합계</td><td></td>{total_cells}</tr>'

ph = "<th style='text-align:left'>회사</th><th>자산총계</th>" + "".join([
    f'<th style="{"color:#e65100;font-weight:800;" if p==sel_period else ""}">{p}</th>' for p in PERIODS])
st.markdown(f'<table class="styled-table"><thead><tr>{ph}</tr></thead><tbody>{pivot_rows}</tbody></table>', unsafe_allow_html=True)

# ── 회사별 시계열 추이 차트 ──────────────────────────────────────────
st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
st.markdown('<div class="section-title">📉 회사별 본드포워드 미결제약정 시계열 추이</div>', unsafe_allow_html=True)

손보_names_list = [c for c, d in RAW.items() if d["구분"]=="손보"]
생보_names_list = [c for c, d in RAW.items() if d["구분"]=="생보"]

# 손보 선택
st.markdown("**손해보험사 선택**", unsafe_allow_html=False)
손보_cols = st.columns(len(손보_names_list))
if "sel_손보" not in st.session_state:
    st.session_state["sel_손보"] = list(손보_names_list)  # 기본 전체 선택
for i, co in enumerate(손보_names_list):
    with 손보_cols[i]:
        is_on = co in st.session_state["sel_손보"]
        btn_style = "background:#1565c0;color:white;border:none;" if is_on else "background:#f0f4f9;color:#4a5568;border:1px solid #e8edf3;"
        if st.button(co, key=f"btn_손보_{co}",
                     help=f"{'클릭하여 제외' if is_on else '클릭하여 추가'}",
                     use_container_width=True):
            if is_on:
                st.session_state["sel_손보"].remove(co)
            else:
                st.session_state["sel_손보"].append(co)
            st.rerun()

# 생보 선택
st.markdown("**생명보험사 선택**", unsafe_allow_html=False)
생보_cols = st.columns(len(생보_names_list))
if "sel_생보" not in st.session_state:
    st.session_state["sel_생보"] = list(생보_names_list)
for i, co in enumerate(생보_names_list):
    with 생보_cols[i]:
        is_on = co in st.session_state["sel_생보"]
        if st.button(co, key=f"btn_생보_{co}",
                     help=f"{'클릭하여 제외' if is_on else '클릭하여 추가'}",
                     use_container_width=True):
            if is_on:
                st.session_state["sel_생보"].remove(co)
            else:
                st.session_state["sel_생보"].append(co)
            st.rerun()

# 선택된 회사 버튼 강조 CSS
selected_all = st.session_state["sel_손보"] + st.session_state["sel_생보"]
btn_css = ""
for co in 손보_names_list:
    if co in st.session_state["sel_손보"]:
        btn_css += f"""
        div[data-testid="stButton"] button[kind="secondary"]:has(div:contains("{co}")) {{
            background: #1565c0 !important; color: white !important;
        }}"""
# Streamlit 버튼 on/off 색상 (전체 적용)
st.markdown("""
<style>
/* 기본 버튼 스타일 */
div[data-testid="stButton"] > button {
    border-radius: 20px !important;
    font-size: 12px !important;
    padding: 4px 10px !important;
    font-weight: 600 !important;
    transition: all 0.15s;
}
</style>
""", unsafe_allow_html=True)

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# 차트 2개 나란히
ts_left, ts_right = st.columns(2)

# 손보 시계열 차트
with ts_left:
    st.markdown('<div class="section-title">🔵 손해보험사</div>', unsafe_allow_html=True)
    colors_손보 = ["#1565c0","#1976d2","#42a5f5","#0d47a1","#29b6f6","#4fc3f7","#0288d1","#039be5","#0097a7"]
    fig_ts1 = go.Figure()
    sel_손보_now = st.session_state["sel_손보"]
    if not sel_손보_now:
        st.info("손해보험사를 위에서 선택해주세요.")
    else:
        for i, co in enumerate(손보_names_list):
            if co not in sel_손보_now:
                continue
            vals = [RAW[co]["잔액"][j] if j < len(RAW[co]["잔액"]) else 0 for j in range(len(PERIODS))]
            fig_ts1.add_trace(go.Scatter(
                x=PERIODS, y=vals, name=co, mode="lines+markers",
                line=dict(color=colors_손보[i % len(colors_손보)], width=2.5),
                marker=dict(size=7),
                hovertemplate=f"<b>{co}</b><br>%{{x}}: %{{y:,.0f}} 억<extra></extra>"
            ))
        fig_ts1.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            font=dict(size=12, color="#2d3748"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                        bgcolor="rgba(255,255,255,0.9)", bordercolor="#e8edf3", borderwidth=1),
            xaxis=dict(showgrid=False, tickfont=dict(size=11)),
            yaxis=dict(showgrid=True, gridcolor="#f0f4f9", tickformat=",", ticksuffix=" 억", tickfont=dict(size=11)),
            margin=dict(l=10, r=10, t=40, b=10), height=380, hovermode="x unified",
        )
        # 조회시점 강조선
        if sel_period in PERIODS:
            idx_x = PERIODS.index(sel_period)
            fig_ts1.add_shape(type="line",
                x0=idx_x, x1=idx_x, y0=0, y1=1, xref="x", yref="paper",
                line=dict(color="#ff7043", width=1.5, dash="dash"))
            fig_ts1.add_annotation(x=idx_x, y=1, xref="x", yref="paper",
                text="조회시점", showarrow=False, font=dict(color="#ff7043", size=10),
                yshift=8)
        st.plotly_chart(fig_ts1, use_container_width=True)

# 생보 시계열 차트
with ts_right:
    st.markdown('<div class="section-title">🟣 생명보험사</div>', unsafe_allow_html=True)
    colors_생보 = ["#6a1b9a","#7b1fa2","#8e24aa","#9c27b0","#ab47bc","#ce93d8","#4a148c","#ba68c8","#e1bee7"]
    fig_ts2 = go.Figure()
    sel_생보_now = st.session_state["sel_생보"]
    if not sel_생보_now:
        st.info("생명보험사를 위에서 선택해주세요.")
    else:
        for i, co in enumerate(생보_names_list):
            if co not in sel_생보_now:
                continue
            vals = [RAW[co]["잔액"][j] if j < len(RAW[co]["잔액"]) else 0 for j in range(len(PERIODS))]
            fig_ts2.add_trace(go.Scatter(
                x=PERIODS, y=vals, name=co, mode="lines+markers",
                line=dict(color=colors_생보[i % len(colors_생보)], width=2.5),
                marker=dict(size=7),
                hovertemplate=f"<b>{co}</b><br>%{{x}}: %{{y:,.0f}} 억<extra></extra>"
            ))
        fig_ts2.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            font=dict(size=12, color="#2d3748"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                        bgcolor="rgba(255,255,255,0.9)", bordercolor="#e8edf3", borderwidth=1),
            xaxis=dict(showgrid=False, tickfont=dict(size=11)),
            yaxis=dict(showgrid=True, gridcolor="#f0f4f9", tickformat=",", ticksuffix=" 억", tickfont=dict(size=11)),
            margin=dict(l=10, r=10, t=40, b=10), height=380, hovermode="x unified",
        )
        if sel_period in PERIODS:
            idx_x = PERIODS.index(sel_period)
            fig_ts2.add_shape(type="line",
                x0=idx_x, x1=idx_x, y0=0, y1=1, xref="x", yref="paper",
                line=dict(color="#ff7043", width=1.5, dash="dash"))
            fig_ts2.add_annotation(x=idx_x, y=1, xref="x", yref="paper",
                text="조회시점", showarrow=False, font=dict(color="#ff7043", size=10),
                yshift=8)
        st.plotly_chart(fig_ts2, use_container_width=True)

# ── 하단 ────────────────────────────────────────────────────────────
st.markdown("""
<div style="margin-top:32px;padding-top:16px;border-top:1px solid #e8edf3;
     display:flex;justify-content:space-between;align-items:center;">
    <span style="font-size:12px;color:#aab0bc;">
        ※ 자료: 각사 DART 공시 / 단위: 억원, % &nbsp; ※ 자산대비 비중 = 본드포워드 잔액 ÷ 자산총계
    </span>
    <span style="font-size:12px;color:#aab0bc;">메리츠화재 자산운용본부</span>
</div>
""", unsafe_allow_html=True)
