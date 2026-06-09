import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import urllib.parse

st.set_page_config(page_title="동업사 본드포워드 현황", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.stApp,[data-testid="stAppViewContainer"],[data-testid="stHeader"]{background:#fff;}
.page-title{font-size:22px;font-weight:700;color:#1a2340;margin:0;}
.page-subtitle{font-size:13px;color:#8a94a6;margin-top:4px;}
.kpi-card{background:#f7f9fc;border:1px solid #e8edf3;border-radius:12px;padding:20px 22px;}
.kpi-label{font-size:12px;color:#8a94a6;font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;}
.kpi-value{font-size:26px;font-weight:700;color:#1a2340;line-height:1.1;}
.kpi-sub{font-size:12px;color:#8a94a6;margin-top:4px;}
.kpi-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;margin-top:6px;}
.badge-up{background:#e8f5e9;color:#2e7d32;}
.badge-down{background:#ffebee;color:#c62828;}
.section-title{font-size:15px;font-weight:700;color:#1a2340;margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid #e8edf3;}
.styled-table{width:100%;border-collapse:collapse;font-size:15px;}
.styled-table thead tr{background:#f0f4f9;}
.styled-table th{padding:9px 10px;text-align:right;font-weight:700;color:#4a5568;border-bottom:2px solid #d1dce8;white-space:nowrap;font-size:14px;}
.styled-table th:first-child{text-align:left;padding-left:12px;}
.styled-table td{padding:7px 10px;text-align:right;color:#2d3748;border-bottom:1px solid #f0f4f9;white-space:nowrap;font-size:15px;line-height:1.3;}
.styled-table td:first-child{text-align:left;font-weight:600;color:#1a2340;padding-left:12px;}
.styled-table td:nth-child(2){
    color:#2d3748;font-size:15px;
    border-left:2px solid #e8edf3;border-right:2px solid #e8edf3;
    background:#fafbfc;}
.styled-table th:nth-child(2){
    color:#4a5568;font-size:14px;
    border-left:2px solid #e8edf3;border-right:2px solid #e8edf3;
    background:#f4f6f9;}
.styled-table tr:hover td{background:#f0f6ff;}
.styled-table tr:hover td:nth-child(2){background:#e8f0ff;}
.styled-table tr.손보-total{background:#dbeafe;font-weight:700;}
.styled-table tr.손보-total td{color:#1565c0;padding:8px 10px;font-size:15px;}
.styled-table tr.손보-total td:first-child{padding-left:12px;}
.styled-table tr.손보-total td:nth-child(2){color:#1565c0;background:#dbeafe;border-left:2px solid #bfdbfe;border-right:2px solid #bfdbfe;}
.styled-table tr.생보-total{background:#ede9fe;font-weight:700;}
.styled-table tr.생보-total td{color:#6a1b9a;padding:8px 10px;font-size:15px;}
.styled-table tr.생보-total td:first-child{padding-left:12px;}
.styled-table tr.생보-total td:nth-child(2){color:#6a1b9a;background:#ede9fe;border-left:2px solid #d8b4fe;border-right:2px solid #d8b4fe;}
.styled-table tr.grand-total{background:#1a2340;color:white;font-weight:700;}
.styled-table tr.grand-total td{color:white;border-bottom:none;padding:9px 10px;font-size:15px;}
.styled-table tr.grand-total td:first-child{padding-left:12px;}
.styled-table tr.grand-total td:nth-child(2){color:white;background:#1a2340;border-left:2px solid #2d3f6b;border-right:2px solid #2d3f6b;}
.share-btn{display:inline-flex;align-items:center;gap:5px;padding:7px 14px;border-radius:8px;font-size:13px;font-weight:600;text-decoration:none;cursor:pointer;border:none;}
.btn-kakao{background:#FEE500;color:#3C1E1E;}
.print-hint{font-size:12px;color:#8a94a6;display:flex;align-items:center;gap:6px;}
.print-hint kbd{background:#f0f4f9;border:1px solid #d1d9e6;border-radius:4px;padding:2px 7px;font-size:12px;color:#4a5568;}

/* 선택버튼 토글 — selected */
div[data-testid="stButton"].btn-selected > button {
    background-color: #e53935 !important;
    color: white !important;
    border-color: #e53935 !important;
}
div[data-testid="stButton"] > button {
    border-radius: 20px !important;
    font-size: 12px !important;
    padding: 4px 10px !important;
    font-weight: 600 !important;
}
@media print{.no-print{display:none!important;}}
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
            p = v.strip().replace("\xa0","")
            if p not in periods: periods.append(p)
    손보_names = ["메리츠화재","삼성화재","DB손해보험","현대해상","KB손해보험","한화손해보험","롯데손해보험","흥국화재","NH농협손해보험"]
    생보_names = ["삼성생명","한화생명","교보생명","신한라이프","농협생명","DB생명","미래에셋생명","동양생명","흥국생명"]
    companies = {}
    for _, row in df_raw.iterrows():
        name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if name not in 손보_names + 생보_names: continue
        asset = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
        balances = []
        for v in row.iloc[3:3+len(periods)]:
            try: balances.append(float(str(v).replace(",","").replace("\xa0","").replace(" ","")))
            except: balances.append(0.0)
        companies[name] = {"구분":"손보" if name in 손보_names else "생보","자산총계":asset,"잔액":balances}
    return companies, periods

try:
    RAW, PERIODS = load_data()
    if not RAW: st.error("❌ 엑셀에서 데이터를 읽지 못했습니다."); st.stop()
except FileNotFoundError:
    st.error("❌ `bond_forward.xlsx` 파일을 찾을 수 없습니다."); st.stop()

def make_df():
    rows = []
    for co, d in RAW.items():
        for i, p in enumerate(PERIODS):
            if i >= len(d["잔액"]): continue
            bal = d["잔액"][i]
            rows.append({"회사":co,"구분":d["구분"],"자산총계":d["자산총계"],"기간":p,"잔액":bal,
                "비중":round(bal/d["자산총계"]*100,2) if d["자산총계"] else 0})
    return pd.DataFrame(rows)

df_long = make_df()

# ── 헤더 ───────────────────────────────────────────────────────────
col_title, col_btn = st.columns([5, 3])
with col_title:
    st.markdown('<div class="page-title">📊 동업사 본드포워드 현황</div><div class="page-subtitle">보험업계 금리선도(Bond Forward) 미결제약정 현황 모니터링</div>', unsafe_allow_html=True)
with col_btn:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    st.components.v1.html("""
    <div style="display:flex;justify-content:flex-end;">
        <button onclick="window.parent.print()"
            style="padding:7px 16px;background:#1a2340;color:white;border:none;
                   border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;
                   white-space:nowrap;">
            🖨️ 인쇄 / PDF
        </button>
    </div>
    """, height=44)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ── 필터 ───────────────────────────────────────────────────────────
fc1, fc2, fc3 = st.columns([2,2,4])
with fc1: sel_period = st.selectbox("📅 조회 기준 시점", PERIODS, index=len(PERIODS)-1)
with fc2: sel_type = st.selectbox("🏢 구분", ["전체","손보","생보"])
with fc3: sel_companies = st.multiselect("🔍 회사 선택 (미선택 시 전체)", options=list(RAW.keys()), default=[])

손보_bal = df_long[(df_long["기간"]==sel_period)&(df_long["구분"]=="손보")]["잔액"].sum()
생보_bal = df_long[(df_long["기간"]==sel_period)&(df_long["구분"]=="생보")]["잔액"].sum()
total_bal = 손보_bal + 생보_bal
pi = PERIODS.index(sel_period)

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

# ── KPI ─────────────────────────────────────────────────────────────
def chg_badge(cur, prev):
    if prev <= 0: return ""
    diff = cur - prev; pct = diff/prev*100
    cls = "badge-up" if diff>=0 else "badge-down"; arrow = "▲" if diff>=0 else "▼"
    return f'<span class="kpi-badge {cls}">{arrow} {abs(pct):.1f}%</span>'

prev_idx = pi - 1
if prev_idx >= 0:
    prev_total = df_long[df_long["기간"]==PERIODS[prev_idx]]["잔액"].sum()
    prev_손보  = df_long[(df_long["기간"]==PERIODS[prev_idx])&(df_long["구분"]=="손보")]["잔액"].sum()
    prev_생보  = df_long[(df_long["기간"]==PERIODS[prev_idx])&(df_long["구분"]=="생보")]["잔액"].sum()
else:
    prev_total = prev_손보 = prev_생보 = 0

top_co = df_long[df_long["기간"]==sel_period].sort_values("비중", ascending=False).iloc[0]
k1,k2,k3,k4 = st.columns(4)
with k1: st.markdown(f'<div class="kpi-card"><div class="kpi-label">생/손보 합계</div><div class="kpi-value">{total_bal/1000:,.1f}조</div><div class="kpi-sub">{total_bal:,.0f} 억원</div>{chg_badge(total_bal,prev_total)}</div>', unsafe_allow_html=True)
with k2: st.markdown(f'<div class="kpi-card"><div class="kpi-label">손보 합계</div><div class="kpi-value">{손보_bal/1000:,.1f}조</div><div class="kpi-sub">{손보_bal:,.0f} 억 · {손보_bal/total_bal*100:.1f}%</div>{chg_badge(손보_bal,prev_손보)}</div>', unsafe_allow_html=True)
with k3: st.markdown(f'<div class="kpi-card"><div class="kpi-label">생보 합계</div><div class="kpi-value">{생보_bal/1000:,.1f}조</div><div class="kpi-sub">{생보_bal:,.0f} 억 · {생보_bal/total_bal*100:.1f}%</div>{chg_badge(생보_bal,prev_생보)}</div>', unsafe_allow_html=True)
with k4: st.markdown(f'<div class="kpi-card"><div class="kpi-label">자산대비 비중 1위</div><div class="kpi-value" style="font-size:20px">{top_co["회사"]}</div><div class="kpi-sub">자산대비 {top_co["비중"]:.2f}%</div></div>', unsafe_allow_html=True)

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

# ── 차트 + 테이블 ──────────────────────────────────────────────────
chart_col, table_col = st.columns([5,5])

with chart_col:
    st.markdown('<div class="section-title">📈 본드포워드 잔액 추이</div>', unsafe_allow_html=True)
    손보_trend = df_long[df_long["구분"]=="손보"].groupby("기간")["잔액"].sum().reindex(PERIODS)
    생보_trend = df_long[df_long["구분"]=="생보"].groupby("기간")["잔액"].sum().reindex(PERIODS)
    fig = go.Figure()
    if sel_companies:
        clrs = px.colors.qualitative.Plotly
        for i,co in enumerate(sel_companies):
            vals=[df_long[(df_long["회사"]==co)&(df_long["기간"]==p)]["잔액"].sum() for p in PERIODS]
            fig.add_trace(go.Scatter(x=PERIODS,y=vals,name=co,mode="lines+markers",line=dict(color=clrs[i%len(clrs)],width=2),marker=dict(size=6)))
    else:
        fig.add_trace(go.Bar(x=PERIODS,y=손보_trend.values,name="손보",marker_color="#1565c0",opacity=0.85))
        fig.add_trace(go.Bar(x=PERIODS,y=생보_trend.values,name="생보",marker_color="#6a1b9a",opacity=0.85))
        fig.add_trace(go.Scatter(x=PERIODS,y=(손보_trend+생보_trend).values,name="합계",mode="lines+markers",line=dict(color="#e53935",width=2.5),marker=dict(size=7)))
    fig.update_layout(barmode="stack",plot_bgcolor="white",paper_bgcolor="white",font=dict(size=12,color="#2d3748"),
        legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1,bgcolor="rgba(255,255,255,0.9)",bordercolor="#e8edf3",borderwidth=1),
        xaxis=dict(showgrid=False,tickfont=dict(size=11)),yaxis=dict(showgrid=True,gridcolor="#f0f4f9",tickformat=",",ticksuffix=" 억",tickfont=dict(size=11)),
        margin=dict(l=10,r=10,t=40,b=10),height=340,hovermode="x unified")
    # 조회시점 해당 바에 테두리로 강조 (색상 변경 없이)
    for trace in fig.data:
        if trace.type=="bar":
            trace.marker.color="#1565c0" if trace.name=="손보" else "#6a1b9a"
            trace.marker.opacity=1
            trace.marker.line.width=[3 if p==sel_period else 0 for p in PERIODS]
            trace.marker.line.color="#ff7043"
    st.plotly_chart(fig,use_container_width=True)

    st.markdown('<div class="section-title" style="margin-top:4px">🥧 생/손보 비중</div>', unsafe_allow_html=True)
    fig2=go.Figure(go.Pie(labels=["손보","생보"],values=[손보_bal,생보_bal],hole=0.6,
        marker_colors=["#1565c0","#6a1b9a"],textinfo="label+percent",textfont_size=13))
    fig2.update_layout(showlegend=False,plot_bgcolor="white",paper_bgcolor="white",margin=dict(l=10,r=10,t=10,b=10),height=210,
        annotations=[dict(text=f"{total_bal/10000:.1f}조",x=0.5,y=0.5,font_size=18,font_color="#1a2340",showarrow=False)])
    st.plotly_chart(fig2,use_container_width=True)

with table_col:
    st.markdown(f'<div class="section-title">📋 상세 현황 ({sel_period})</div>', unsafe_allow_html=True)
    def render_group(구분):
        cos=[c for c,d in RAW.items() if d["구분"]==구분]
        if sel_companies: cos=[c for c in cos if c in sel_companies]
        rows_html=""; grp_bal=0
        for co in cos:
            d=RAW[co]
            cur=d["잔액"][pi] if pi<len(d["잔액"]) else 0
            prev=d["잔액"][pi-1] if pi>0 and (pi-1)<len(d["잔액"]) else cur
            diff=cur-prev; arrow="▲" if diff>0 else("▼" if diff<0 else "–")
            color="#2e7d32" if diff>0 else("#c62828" if diff<0 else "#888")
            ratio=cur/d["자산총계"]*100 if d["자산총계"] else 0
            rows_html+=f'<tr><td>{co}</td><td>{d["자산총계"]:,.0f}</td><td>{cur:,.0f}</td><td><span style="color:{color};font-size:11px">{arrow} {abs(diff):,.0f}</span></td><td style="border-left:3px solid #ff7043;border-right:3px solid #ff7043;">{ratio:.2f}%</td></tr>'
            grp_bal+=cur
        share=grp_bal/total_bal*100 if total_bal else 0
        cls="손보-total" if 구분=="손보" else "생보-total"
        grp_asset = sum(RAW[c]["자산총계"] for c in cos)
        rows_html+=f'<tr class="{cls}"><td>{"주요 손보사 계" if 구분=="손보" else "주요 생보사 계"}</td><td>{grp_asset:,.0f}</td><td>{grp_bal:,.0f}</td><td></td><td style="border-left:3px solid #ff7043;border-right:3px solid #ff7043;">{share:.1f}%</td></tr>'
        return rows_html
    bg1,fg1="#e3f2fd","#1565c0"; bg2,fg2="#f3e5f5","#6a1b9a"
    s_rows=f'<tr><td colspan="5" style="background:{bg1};color:{fg1};font-weight:700;padding:6px 14px;font-size:12px">▶ 손해보험</td></tr>'+render_group("손보")
    g_rows=f'<tr><td colspan="5" style="background:{bg2};color:{fg2};font-weight:700;padding:6px 14px;font-size:12px">▶ 생명보험</td></tr>'+render_group("생보")
    total_asset = sum(d["자산총계"] for d in RAW.values())
    grand=f'<tr class="grand-total"><td>주요 생손보 합계</td><td>{total_asset:,.0f}</td><td>{total_bal:,.0f}</td><td></td><td style="border-left:3px solid #ff7043;border-right:3px solid #ff7043;">100%</td></tr>'
    header='<table class="styled-table"><thead><tr><th style="text-align:left">회사</th><th>자산총계(억)</th><th>잔액(억)</th><th>전분기比</th><th style="border-left:3px solid #ff7043;border-right:3px solid #ff7043;border-top:3px solid #ff7043;color:#ff7043;">자산비중</th></tr></thead><tbody>'
    body={"전체":s_rows+g_rows+grand,"손보":s_rows,"생보":g_rows}
    st.markdown(header+body[sel_type]+"</tbody></table>", unsafe_allow_html=True)

# ── 기간별 전체 추이 테이블 ──────────────────────────────────────────
st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
st.markdown('<div class="section-title">📅 기간별 잔액 추이 (억원)</div>', unsafe_allow_html=True)
cos_show=list(RAW.keys())
if sel_companies: cos_show=sel_companies
elif sel_type!="전체": cos_show=[c for c,d in RAW.items() if d["구분"]==sel_type]
pivot_rows=""; grp_totals={"손보":[0]*len(PERIODS),"생보":[0]*len(PERIODS)}
for 구분,bg,fg,label,cls in [("손보","#e3f2fd","#1565c0","손해보험","손보-total"),("생보","#f3e5f5","#6a1b9a","생명보험","생보-total")]:
    cos_g=[c for c in cos_show if RAW[c]["구분"]==구분]
    if not cos_g: continue
    pivot_rows+=f'<tr><td colspan="{len(PERIODS)+2}" style="background:{bg};color:{fg};font-weight:700;padding:6px 14px;font-size:12px">▶ {label}</td></tr>'
    grp_sums=[0]*len(PERIODS)
    for co in cos_g:
        d=RAW[co]; cells=""
        for i,p in enumerate(PERIODS):
            bal=d["잔액"][i] if i<len(d["잔액"]) else 0; grp_sums[i]+=bal
            hl="font-weight:700;border-left:3px solid #ff7043;border-right:3px solid #ff7043;" if p==sel_period else ""
            cells+=f'<td style="{hl}">{bal:,.0f}</td>'
        pivot_rows+=f'<tr><td>{co}</td><td>{d["자산총계"]:,.0f}</td>{cells}</tr>'
    gcells="".join([f'<td style="background:{bg};font-weight:700;color:{fg};font-size:15px;padding:11px 16px;{"border-left:3px solid #ff7043;border-right:3px solid #ff7043;" if p==sel_period else ""}">{s:,.0f}</td>' for p,s in zip(PERIODS,grp_sums)])
    pivot_rows+=f'<tr class="{cls}"><td>{"주요 손보사 계" if 구분=="손보" else "주요 생보사 계"}</td><td></td>{gcells}</tr>'
    for i in range(len(PERIODS)): grp_totals[구분][i]=grp_sums[i]
total_cells="".join([f'<td style="background:{"#2d3f6b" if p==sel_period else "#1a2340"};font-weight:700;font-size:15px;padding:11px 16px;{"border-left:3px solid #ff7043;border-right:3px solid #ff7043;" if p==sel_period else ""}">{grp_totals["손보"][i]+grp_totals["생보"][i]:,.0f}</td>' for i,p in enumerate(PERIODS)])
pivot_rows+=f'<tr class="grand-total"><td>생/손보 합계</td><td></td>{total_cells}</tr>'
ph="<th style='text-align:left'>회사</th><th>자산총계</th>"+"".join([f'<th style="{"color:#ff7043;font-weight:800;border-left:3px solid #ff7043;border-right:3px solid #ff7043;border-top:3px solid #ff7043;" if p==sel_period else ""}">{p}</th>' for p in PERIODS])
st.markdown(f'<table class="styled-table"><thead><tr>{ph}</tr></thead><tbody>{pivot_rows}</tbody></table>', unsafe_allow_html=True)

# ── 회사별 시계열 추이 ───────────────────────────────────────────────
st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
st.markdown('<div class="section-title">📉 회사별 본드포워드 미결제약정 시계열 추이</div>', unsafe_allow_html=True)

손보_list = [c for c,d in RAW.items() if d["구분"]=="손보"]
생보_list  = [c for c,d in RAW.items() if d["구분"]=="생보"]

# 회사별 고유 색상 (손보: 파란/청록 계열, 생보: 보라/분홍 계열, 모두 명확히 구분)
COLORS_손보 = ["#1565c0","#e53935","#2e7d32","#f57c00","#00838f","#6d4c41","#37474f","#ad1457","#558b2f"]
COLORS_생보  = ["#6a1b9a","#0277bd","#558b2f","#d84315","#00695c","#4527a0","#283593","#827717","#37474f"]

if "sel_손보" not in st.session_state: st.session_state["sel_손보"]=list(손보_list)
if "sel_생보" not in st.session_state: st.session_state["sel_생보"]=list(생보_list)

# ── 손보 토글 버튼 ─────────────────────────────────────────────────
st.markdown("**🔵 손해보험사 선택** (클릭하여 ON/OFF)")
손보_cols = st.columns(len(손보_list))
for i, co in enumerate(손보_list):
    with 손보_cols[i]:
        is_on = co in st.session_state["sel_손보"]
        color = COLORS_손보[i]
        # 선택 = 해당 회사 색상, 미선택 = 회색
        btn_color = color if is_on else "#e0e0e0"
        txt_color = "white" if is_on else "#999"
        st.markdown(f"""
        <div style="text-align:center;margin-bottom:4px;">
            <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{btn_color};margin-right:3px;"></span>
            <span style="font-size:11px;color:{txt_color};font-weight:{'700' if is_on else '400'}">{co}</span>
        </div>""", unsafe_allow_html=True)
        if st.button(f"{'✅' if is_on else '○'} {co}", key=f"ts_손보_{co}", use_container_width=True):
            if is_on: st.session_state["sel_손보"].remove(co)
            else: st.session_state["sel_손보"].append(co)
            st.rerun()

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ── 생보 토글 버튼 ─────────────────────────────────────────────────
st.markdown("**🟣 생명보험사 선택** (클릭하여 ON/OFF)")
생보_cols = st.columns(len(생보_list))
for i, co in enumerate(생보_list):
    with 생보_cols[i]:
        is_on = co in st.session_state["sel_생보"]
        color = COLORS_생보[i]
        btn_color = color if is_on else "#e0e0e0"
        txt_color = "white" if is_on else "#999"
        st.markdown(f"""
        <div style="text-align:center;margin-bottom:4px;">
            <span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{btn_color};margin-right:3px;"></span>
            <span style="font-size:11px;color:{txt_color};font-weight:{'700' if is_on else '400'}">{co}</span>
        </div>""", unsafe_allow_html=True)
        if st.button(f"{'✅' if is_on else '○'} {co}", key=f"ts_생보_{co}", use_container_width=True):
            if is_on: st.session_state["sel_생보"].remove(co)
            else: st.session_state["sel_생보"].append(co)
            st.rerun()

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# ── 시계열 차트 ────────────────────────────────────────────────────
ts_left, ts_right = st.columns(2)

def make_ts_chart(name_list, sel_list, color_list, title):
    fig = go.Figure()
    if not sel_list:
        return None
    for i, co in enumerate(name_list):
        if co not in sel_list: continue
        vals=[RAW[co]["잔액"][j] if j<len(RAW[co]["잔액"]) else 0 for j in range(len(PERIODS))]
        fig.add_trace(go.Scatter(x=PERIODS, y=vals, name=co, mode="lines+markers",
            line=dict(color=color_list[i], width=2.5), marker=dict(size=7, color=color_list[i]),
            hovertemplate=f"<b>{co}</b><br>%{{x}}: %{{y:,.0f}} 억<extra></extra>"))
    fig.update_layout(plot_bgcolor="white", paper_bgcolor="white", font=dict(size=12, color="#2d3748"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    bgcolor="rgba(255,255,255,0.9)", bordercolor="#e8edf3", borderwidth=1),
        xaxis=dict(showgrid=False, tickfont=dict(size=11)),
        yaxis=dict(showgrid=True, gridcolor="#f0f4f9", tickformat=",", ticksuffix=" 억", tickfont=dict(size=11)),
        margin=dict(l=10, r=10, t=40, b=10), height=380, hovermode="x unified")
    # 조회시점 강조
    if sel_period in PERIODS:
        idx_x = PERIODS.index(sel_period)
        fig.add_shape(type="line", x0=idx_x, x1=idx_x, y0=0, y1=1, xref="x", yref="paper",
            line=dict(color="#ff7043", width=1.5, dash="dash"))
        fig.add_annotation(x=idx_x, y=1.05, xref="x", yref="paper",
            text="조회시점", showarrow=False, font=dict(color="#ff7043", size=10))
    return fig

with ts_left:
    st.markdown('<div class="section-title">🔵 손해보험사 추이</div>', unsafe_allow_html=True)
    fig_s = make_ts_chart(손보_list, st.session_state["sel_손보"], COLORS_손보, "손보")
    if fig_s: st.plotly_chart(fig_s, use_container_width=True)
    else: st.info("손해보험사를 위에서 선택해주세요.")

with ts_right:
    st.markdown('<div class="section-title">🟣 생명보험사 추이</div>', unsafe_allow_html=True)
    fig_g = make_ts_chart(생보_list, st.session_state["sel_생보"], COLORS_생보, "생보")
    if fig_g: st.plotly_chart(fig_g, use_container_width=True)
    else: st.info("생명보험사를 위에서 선택해주세요.")

# ── 하단 ─────────────────────────────────────────────────────────
st.markdown("""
<div style="margin-top:32px;padding-top:16px;border-top:1px solid #e8edf3;display:flex;justify-content:space-between;align-items:center;">
    <span style="font-size:12px;color:#aab0bc;">※ 자료: 각사 DART 공시 / 단위: 억원, % &nbsp;|&nbsp; ※ 자산대비 비중 = 본드포워드 잔액 ÷ 자산총계 &nbsp;|&nbsp; ※ 자산총계는 조회 기준 시점 최근 분기 기준</span>
    <span style="font-size:12px;color:#aab0bc;">메리츠화재 자산운용본부</span>
</div>""", unsafe_allow_html=True)
