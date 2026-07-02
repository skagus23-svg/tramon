import streamlit as st
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
import altair as alt
import math, os, warnings, io, json
from datetime import date, timedelta
warnings.filterwarnings('ignore')

# ── Supabase 연결 ─────────────────────────────────────────────
def get_supabase():
    try:
        from supabase import create_client
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception:
        return None

sb = get_supabase()
LOCAL_MODE = sb is None

# ── 페이지 설정 ───────────────────────────────────────────────
st.set_page_config(page_title="트라몬 재고 관리", layout="wide", page_icon="🧳")
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #FFFFFF; }
[data-testid="stSidebar"] { background: #F5F7FA; }
[data-testid="stHeader"] { background: #FFFFFF; }
.stTabs [data-baseweb="tab"] { color: #666; }
.stTabs [aria-selected="true"] { color: #1A73E8 !important; border-bottom: 2px solid #1A73E8; }
.alert-red    { background:#FFF0F0; border-left:4px solid #E05252; padding:10px 14px; border-radius:6px; margin:4px 0; color:#333; }
.alert-yellow { background:#FFFBEA; border-left:4px solid #F0A020; padding:10px 14px; border-radius:6px; margin:4px 0; color:#333; }
.alert-green  { background:#F0FFF4; border-left:4px solid #4CAF50; padding:10px 14px; border-radius:6px; margin:4px 0; color:#333; }
.formula-box  { background:#F0F4FF; border:1px solid #C5D5F5; border-radius:8px; padding:16px 20px; font-family:monospace; font-size:15px; color:#1A1A2E; }
h1, h2, h3 { color: #1A1A2E; }
</style>
""", unsafe_allow_html=True)

# ── 상수 ─────────────────────────────────────────────────────
def to_short_month(dt):
    """datetime → '25.1' 형식"""
    if hasattr(dt, 'year'):
        return f"{str(dt.year)[2:]}.{dt.month}"
    return str(dt)

COLORS    = ['다크그레이', '블랙', '화이트']
SIZES     = ['20인치', '24인치', '28인치']
SKUS      = [f"{c}/{s}" for c in COLORS for s in SIZES]
COLOR_MOQ = 500   # 색상별 최소 발주량
SIZE_MOQ  = 200   # 사이즈별 최소 발주량

DATA_FILES = [
    ('data/ss_26.xlsx',      17),
    ('data/ss_26_25.xlsx',   17),
    ('data/ss_26_25_2.xlsx', 17),
    ('data/ss_26_1.xlsx',    18),
]

# ── 데이터 처리 ───────────────────────────────────────────────
def is_pouch(name, opt):
    if '파우치' in name or '스페이스' in name: return True
    if 'S+S+M' in name or 'S+M+L' in name: return True
    if 'S+S+M' in opt  or 'S+M+L' in opt:  return True
    return False

def get_size(name, opt, sold=''):
    for o in [str(opt), str(sold)]:
        if '28인치' in o or '75cm' in o or '사이즈: 28' in o or '사이즈 : 75cm' in o or '사이즈 : 28' in o: return '28인치'
        if '24인치' in o or '65cm' in o or '사이즈: 24' in o or '사이즈 : 65cm' in o or '사이즈 : 24' in o: return '24인치'
        if '20인치' in o or '55cm' in o or '사이즈: 20' in o or '사이즈 : 55cm' in o or '사이즈 : 20' in o: return '20인치'
    n = str(name)
    if '28인치' in n or '75cm' in n: return '28인치'
    if '24인치' in n or '65cm' in n: return '24인치'
    if '20인치' in n or '55cm' in n: return '20인치'
    return '기타'

def get_color(name, opt, sold=''):
    for o in [str(opt), str(sold)]:
        if '다크그레이' in o: return '다크그레이'
        if '블랙' in o: return '블랙'
        if '화이트' in o: return '화이트'
    n = str(name)
    if '다크그레이' in n: return '다크그레이'
    if '블랙' in n: return '블랙'
    if '화이트' in n: return '화이트'
    return '기타'

def parse_excel(file_obj, ncols):
    df = pd.read_excel(file_obj, sheet_name=0)
    if len(df.columns) == 17:
        df.columns = ['상품주문번호','주문번호','주문시각','주문상태','발송속성','풀필먼트',
                      '클레임상태','클레임처리','상품번호','상품명','옵션정보','수량',
                      '구매자명','구매자ID','수취인명','배송요청회원','배송요청비회원']
        df['판매옵션정보'] = ''
    elif len(df.columns) == 18:
        df.columns = ['상품주문번호','주문번호','주문시각','주문상태','발송속성','풀필먼트',
                      '클레임상태','클레임처리','상품번호','상품명','옵션정보','판매옵션정보','수량',
                      '구매자명','구매자ID','수취인명','배송요청회원','배송요청비회원']
    else:
        return pd.DataFrame()
    df['주문시각'] = pd.to_datetime(df['주문시각'])
    confirmed = df[df['주문상태'] == '구매확정'].copy()
    rows = []
    for _, r in confirmed.iterrows():
        name = str(r['상품명']); opt = str(r['옵션정보']); sold = str(r.get('판매옵션정보',''))
        if is_pouch(name, opt): continue
        sz = get_size(name, opt, sold); cl = get_color(name, opt, sold)
        if sz == '기타' or cl == '기타': continue
        rows.append({'날짜': r['주문시각'], 'SKU': f"{cl}/{sz}", '수량': r['수량']})
    return pd.DataFrame(rows)

@st.cache_data
def load_local_sales():
    all_rows = []
    for path, ncols in DATA_FILES:
        if not os.path.exists(path): continue
        try:
            df = parse_excel(path, ncols)
            all_rows.append(df)
        except Exception:
            continue
    if not all_rows: return pd.DataFrame(columns=['월','SKU','수량'])
    raw = pd.concat(all_rows, ignore_index=True)
    raw['월'] = raw['날짜'].dt.to_period('M')
    return raw.groupby(['월','SKU'])['수량'].sum().reset_index()

# ── 입고 이력 로드 ────────────────────────────────────────────
@st.cache_data
def load_stock_in():
    path = 'data/stock_in.csv'
    if not os.path.exists(path): return pd.DataFrame()
    df = pd.read_csv(path, encoding='utf-8-sig')
    df['날짜'] = pd.to_datetime(df['날짜'])
    return df

# ── 재고 DB ───────────────────────────────────────────────────
LOCAL_INV_PATH = 'data/current_inventory.json'

def get_inventory():
    if sb:
        rows = sb.table('inventory').select('sku,stock,updated_at').execute().data
        return {r['sku']: r['stock'] for r in rows}
    # 로컬 JSON 파일 사용
    if os.path.exists(LOCAL_INV_PATH):
        with open(LOCAL_INV_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {sku: 0 for sku in SKUS}

def set_inventory(sku, qty, note='수동입력'):
    if sb:
        sb.table('inventory').upsert({'sku': sku, 'stock': qty, 'updated_at': 'now()'}).execute()
        sb.table('stock_log').insert({'sku': sku, 'change_qty': qty, 'note': note}).execute()
    else:
        # 로컬 JSON 업데이트
        inv = get_inventory()
        inv[sku] = qty
        with open(LOCAL_INV_PATH, 'w', encoding='utf-8') as f:
            json.dump(inv, f, ensure_ascii=False, indent=2)

def get_stock_log():
    if sb:
        rows = sb.table('stock_log').select('*').order('created_at', desc=True).limit(100).execute().data
        return pd.DataFrame(rows)
    return pd.DataFrame()

# ── 수요예측 ─────────────────────────────────────────────────
@st.cache_data
def forecast_sku(monthly_tuple, sku, n=6):
    monthly = pd.DataFrame(monthly_tuple, columns=['월','SKU','수량'])
    monthly['월_dt'] = monthly['월'].apply(lambda x: pd.Period(x, 'M').to_timestamp())
    sub = monthly[monthly['SKU'] == sku].sort_values('월_dt').reset_index(drop=True)
    if len(sub) < 4: return pd.DataFrame(), 0
    sub['t'] = np.arange(len(sub))
    avg = sub['수량'].mean()
    seasonal = {m: sub[sub['월_dt'].dt.month == m]['수량'].mean() / avg
                for m in range(1,13) if len(sub[sub['월_dt'].dt.month == m]) > 0}
    model = LinearRegression().fit(sub[['t']], sub['수량'])
    last_t = sub['t'].iloc[-1]; last_dt = sub['월_dt'].iloc[-1]
    future = pd.date_range(last_dt + pd.offsets.MonthBegin(), periods=n, freq='MS')
    preds = []
    for i, dt in enumerate(future):
        tr = model.predict([[last_t + 1 + i]])[0]
        s = seasonal.get(dt.month, 1.0)
        preds.append({'월': dt, 'SKU': sku, '예측': max(0, round(0.75*tr + 0.25*avg*s))})
    return pd.DataFrame(preds), model.coef_[0]

LEAD_TIME = 45  # 발주 후 입고까지 일수 (고정)

# ── 계절 조정 발주 타이밍 계산 ────────────────────────────────
def get_seasonal_factors(monthly_raw, sku):
    """월별 계절 지수 계산 (1.0 = 평균)"""
    sub = monthly_raw[monthly_raw['SKU'] == sku].copy()
    if len(sub) == 0: return {m: 1.0 for m in range(1, 13)}
    sub['month'] = sub['월_dt'].dt.month
    avg = sub['수량'].mean()
    if avg == 0: return {m: 1.0 for m in range(1, 13)}
    factors = {}
    for m in range(1, 13):
        vals = sub[sub['month'] == m]['수량'].values
        factors[m] = float(vals.mean() / avg) if len(vals) > 0 else 1.0
    return factors

def calc_reorder_timing(current_stock, monthly_raw, sku, lead_time=LEAD_TIME):
    """
    계절 조정 일평균 수요로 소진일 시뮬레이션 → 발주 권장일 계산
    발주 권장일 = 소진 예정일 - lead_time
    """
    sub = monthly_raw[monthly_raw['SKU'] == sku]
    if len(sub) == 0: return None, None, None, 0
    avg_monthly = sub['수량'].mean()
    if avg_monthly == 0: return None, None, None, 0

    seasonal = get_seasonal_factors(monthly_raw, sku)
    avg_daily = avg_monthly / 30

    today = date.today()
    stock = float(current_stock)
    stockout_day = None

    for d in range(730):  # 최대 2년
        m = (today + timedelta(days=d)).month
        daily = avg_daily * seasonal.get(m, 1.0)
        stock -= daily
        if stock <= 0:
            stockout_day = d
            break

    if stockout_day is None:
        return None, None, None, avg_daily  # 2년 이상 여유

    stockout_date  = today + timedelta(days=stockout_day)
    reorder_days   = stockout_day - lead_time  # 오늘부터 발주까지 남은 일수
    reorder_date   = today + timedelta(days=max(0, reorder_days))

    return reorder_date, stockout_date, reorder_days, avg_daily

def calc_reorder_table(inventory, monthly_raw, lead_time=LEAD_TIME):
    rows = []
    for sku in SKUS:
        stock = inventory.get(sku, 0)
        reorder_date, stockout_date, reorder_days, avg_daily = calc_reorder_timing(
            stock, monthly_raw, sku, lead_time)

        seasonal = get_seasonal_factors(monthly_raw, sku)
        today_seasonal = seasonal.get(date.today().month, 1.0)
        adj_daily = avg_daily * today_seasonal

        if reorder_days is None:
            status = '✅ 2년+ 여유'; urgency = 3
        elif reorder_days < 0:
            status = '🔴 즉시 발주 (이미 지남)'; urgency = 0
        elif reorder_days == 0:
            status = '🔴 오늘 발주'; urgency = 0
        elif reorder_days <= 14:
            status = f'🟠 {reorder_days}일 후 발주'; urgency = 1
        elif reorder_days <= 30:
            status = f'🟡 {reorder_days}일 후 발주'; urgency = 2
        else:
            status = f'✅ {reorder_days}일 후 발주'; urgency = 3

        rows.append({
            'SKU': sku,
            '현재재고': stock,
            '일평균(계절조정)': round(adj_daily, 1),
            '소진예상일': stockout_date.strftime('%Y-%m-%d') if stockout_date else '2년+ 이후',
            '발주권장일': reorder_date.strftime('%Y-%m-%d') if reorder_date else '-',
            'D-day': reorder_days if reorder_days is not None else 999,
            '상태': status,
            '_urgency': urgency,
        })

    return pd.DataFrame(rows).sort_values('_urgency')

# ── 특허 수식: 최적 재고량 Rt ─────────────────────────────────
def calc_optimal_inventory(Pt, St, Bt, Dt, Vt, gamma, wx=1.0, wy=1.0):
    """
    Rt = Pt + wx·log(St+1) + wy·(Bt-Dt) + γ·cos(2π·Vt/100)
    Pt: 과거 재고량, St: 최근 판매량, Bt: 초기 주문량
    Dt: 예상 수요량, Vt: 여행객 증가율(%), γ: 진폭 조절 계수
    """
    log_term      = wx * math.log(St + 1)
    demand_term   = wy * (Bt - Dt)
    seasonal_term = gamma * math.cos(2 * math.pi * Vt / 100)
    Rt = Pt + log_term + demand_term + seasonal_term
    return max(0, round(Rt))

# ── Altair 차트 (x축 수평 레이블) ────────────────────────────
def make_line_chart(df, x_col, y_col, color_col=None, title=''):
    if color_col:
        chart = alt.Chart(df).mark_line(point=True).encode(
            x=alt.X(f'{x_col}:N', sort=None,
                    axis=alt.Axis(labelAngle=0, title='', labelFontSize=11)),
            y=alt.Y(f'{y_col}:Q', axis=alt.Axis(title='수량')),
            color=alt.Color(f'{color_col}:N'),
            tooltip=[x_col, color_col, y_col]
        )
    else:
        chart = alt.Chart(df).mark_line(point=True, color='#1A73E8').encode(
            x=alt.X(f'{x_col}:N', sort=None,
                    axis=alt.Axis(labelAngle=0, title='', labelFontSize=11)),
            y=alt.Y(f'{y_col}:Q', axis=alt.Axis(title='수량')),
            tooltip=[x_col, y_col]
        )
    return chart.properties(title=title, height=300).configure_view(strokeWidth=0)

# ═══════════════════════════════════════════════════════════════
# 메인 UI
# ═══════════════════════════════════════════════════════════════
st.title("🧳 트라몬 캐리어 재고 관리 시스템")
if LOCAL_MODE:
    st.warning("⚠️ 로컬 모드 — 재고 저장 기능은 Supabase 연결 후 활성화됩니다.", icon="🔒")

monthly_raw = load_local_sales()
if monthly_raw.empty or '월' not in monthly_raw.columns:
    st.info("📂 판매 데이터가 없습니다. 'Excel 업로드' 탭에서 엑셀 파일을 업로드해주세요.")
    monthly_raw = pd.DataFrame(columns=['월','SKU','수량','월_dt','월_label'])
else:
    monthly_raw['월_dt'] = monthly_raw['월'].apply(lambda x: pd.Period(x, 'M').to_timestamp())
    monthly_raw['월_label'] = monthly_raw['월_dt'].apply(to_short_month)

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 판매 추이", "🔮 수요예측", "📦 재고 관리", "🤖 AI 재고 최적화", "📂 엑셀 업로드"])

# ════════════ TAB 1: 판매 추이 ══════════════════════════════
with tab1:
    total = monthly_raw.groupby(['월_dt','월_label'])['수량'].sum().reset_index()
    total = total.sort_values('월_dt')

    recent3 = total.tail(3)['수량'].mean()
    prev3   = total.iloc[-6:-3]['수량'].mean() if len(total) >= 6 else recent3
    growth  = (recent3 - prev3) / prev3 * 100 if prev3 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("전체 누적 판매", f"{monthly_raw['수량'].sum():,}개")
    k2.metric("최근 3개월 평균", f"{recent3:,.0f}개/월", f"{growth:+.1f}%")
    k3.metric("최고 판매 SKU", monthly_raw.groupby('SKU')['수량'].sum().idxmax())
    k4.metric("최근월 판매", f"{total.iloc[-1]['수량']:,}개")

    st.markdown("---")
    st.subheader("전체 월별 판매 추이 + 향후 3개월 예측")

    # 전체 합계 3개월 예측
    monthly_tuple_t1 = [tuple(r) for r in monthly_raw[['월','SKU','수량']].itertuples(index=False)]
    fc_total_rows = []
    for sku in SKUS:
        fc, _ = forecast_sku(tuple(monthly_tuple_t1), sku, 3)
        if not fc.empty:
            fc_total_rows.append(fc)
    if fc_total_rows:
        fc_total = pd.concat(fc_total_rows).groupby('월')['예측'].sum().reset_index()
        fc_total['월_label'] = fc_total['월'].apply(to_short_month)
        fc_total['구분'] = '예측'
    else:
        fc_total = pd.DataFrame(columns=['월_label','예측','구분'])

    hist_total = total[['월_label','수량']].copy()
    hist_total['구분'] = '실적'
    hist_total = hist_total.rename(columns={'수량':'값'})
    fc_total_disp = fc_total[['월_label','예측','구분']].rename(columns={'예측':'값'}) if not fc_total.empty else pd.DataFrame()

    # 실적 라인 (파란색 실선)
    all_labels = list(hist_total['월_label']) + (list(fc_total_disp['월_label']) if not fc_total_disp.empty else [])
    hist_line = alt.Chart(hist_total).mark_line(point=True, color='#1A73E8', strokeWidth=2).encode(
        x=alt.X('월_label:N', sort=all_labels, axis=alt.Axis(labelAngle=0, title='')),
        y=alt.Y('값:Q', title='수량'),
        tooltip=['월_label','값']
    )
    layers = [hist_line]
    if not fc_total_disp.empty:
        fc_line = alt.Chart(fc_total_disp).mark_line(
            point=alt.OverlayMarkDef(color='#E08020', size=80),
            strokeDash=[6, 3], color='#E08020', strokeWidth=2
        ).encode(
            x=alt.X('월_label:N', sort=all_labels, axis=alt.Axis(labelAngle=0, title='')),
            y=alt.Y('값:Q'),
            tooltip=['월_label','값']
        )
        layers.append(fc_line)
        # 연결선 (마지막 실적 → 첫 예측)
        bridge = pd.DataFrame([
            {'월_label': hist_total.iloc[-1]['월_label'], '값': hist_total.iloc[-1]['값']},
            {'월_label': fc_total_disp.iloc[0]['월_label'], '값': fc_total_disp.iloc[0]['값']},
        ])
        bridge_line = alt.Chart(bridge).mark_line(strokeDash=[4,2], color='#AAAAAA').encode(
            x=alt.X('월_label:N', sort=all_labels),
            y='값:Q'
        )
        layers.append(bridge_line)

    trend_chart = alt.layer(*layers).properties(height=320).configure_view(strokeWidth=0)
    st.altair_chart(trend_chart, use_container_width=True)
    if not fc_total_disp.empty:
        st.caption("━━ 실적 (파랑)   ╌╌ 예측 (주황)")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("색상별 추이 + 예측")
        cm = monthly_raw.copy()
        cm['색상'] = cm['SKU'].str.split('/').str[0]
        cm_g = cm.groupby(['월_label','월_dt','색상'])['수량'].sum().reset_index().sort_values('월_dt')

        # 색상별 예측
        fc_color_rows = []
        if fc_total_rows:
            fc_color_df = pd.concat(fc_total_rows).copy()
            fc_color_df['색상'] = fc_color_df['SKU'].str.split('/').str[0]
            fc_color_g = fc_color_df.groupby(['월','색상'])['예측'].sum().reset_index()
            fc_color_g['월_label'] = fc_color_g['월'].apply(to_short_month)
            fc_color_g = fc_color_g.rename(columns={'예측':'수량'})
            fc_color_rows = fc_color_g

        hist_c = cm_g[['월_label','색상','수량']].assign(구분='실적')
        hist_line_c = alt.Chart(hist_c).mark_line(point=True, strokeWidth=2).encode(
            x=alt.X('월_label:N', sort=None, axis=alt.Axis(labelAngle=0, title='')),
            y=alt.Y('수량:Q'), color='색상:N', tooltip=['월_label','색상','수량']
        )
        layers_c = [hist_line_c]
        if len(fc_color_rows) > 0:
            fc_c_line = alt.Chart(fc_color_rows).mark_line(
                strokeDash=[5,3], strokeWidth=2
            ).encode(
                x=alt.X('월_label:N', sort=None, axis=alt.Axis(labelAngle=0, title='')),
                y='수량:Q', color='색상:N', tooltip=['월_label','색상','수량']
            )
            layers_c.append(fc_c_line)
        st.altair_chart(
            alt.layer(*layers_c).properties(height=280).configure_view(strokeWidth=0),
            use_container_width=True)

    with col2:
        st.subheader("사이즈별 추이 + 예측")
        sm = monthly_raw.copy()
        sm['사이즈'] = sm['SKU'].str.split('/').str[1]
        sm_g = sm.groupby(['월_label','월_dt','사이즈'])['수량'].sum().reset_index().sort_values('월_dt')

        fc_size_rows = pd.DataFrame()
        if fc_total_rows:
            fc_size_df = pd.concat(fc_total_rows).copy()
            fc_size_df['사이즈'] = fc_size_df['SKU'].str.split('/').str[1]
            fc_size_g = fc_size_df.groupby(['월','사이즈'])['예측'].sum().reset_index()
            fc_size_g['월_label'] = fc_size_g['월'].apply(to_short_month)
            fc_size_g = fc_size_g.rename(columns={'예측':'수량'})
            fc_size_rows = fc_size_g

        hist_s = sm_g[['월_label','사이즈','수량']].assign(구분='실적')
        hist_line_s = alt.Chart(hist_s).mark_line(point=True, strokeWidth=2).encode(
            x=alt.X('월_label:N', sort=None, axis=alt.Axis(labelAngle=0, title='')),
            y=alt.Y('수량:Q'), color='사이즈:N', tooltip=['월_label','사이즈','수량']
        )
        layers_s = [hist_line_s]
        if len(fc_size_rows) > 0:
            fc_s_line = alt.Chart(fc_size_rows).mark_line(strokeDash=[5,3], strokeWidth=2).encode(
                x=alt.X('월_label:N', sort=None, axis=alt.Axis(labelAngle=0, title='')),
                y='수량:Q', color='사이즈:N', tooltip=['월_label','사이즈','수량']
            )
            layers_s.append(fc_s_line)
        st.altair_chart(
            alt.layer(*layers_s).properties(height=280).configure_view(strokeWidth=0),
            use_container_width=True)

    st.subheader("SKU별 누적 판매 순위")
    sku_total = monthly_raw.groupby('SKU')['수량'].sum().sort_values(ascending=False).reset_index()
    sku_total.columns = ['SKU','총판매량']
    sku_total['비중'] = (sku_total['총판매량'] / sku_total['총판매량'].sum() * 100).round(1).astype(str) + '%'
    st.dataframe(sku_total, use_container_width=True, hide_index=True)

# ════════════ TAB 2: 수요예측 ════════════════════════════════
with tab2:
    st.subheader("향후 6개월 SKU별 수요 예측")
    st.caption("선형 추세 + 계절 지수 모델")

    monthly_tuple = [tuple(r) for r in monthly_raw[['월','SKU','수량']].itertuples(index=False)]
    fc_list = []; slopes = {}
    for sku in SKUS:
        fc, slope = forecast_sku(tuple(monthly_tuple), sku, 6)
        if not fc.empty:
            fc_list.append(fc); slopes[sku] = slope

    if fc_list:
        fc_df = pd.concat(fc_list, ignore_index=True)
        fc_pivot = fc_df.pivot_table(index='SKU', columns='월', values='예측', aggfunc='sum', fill_value=0)
        fc_pivot.columns = [f"{c.month}월" for c in fc_pivot.columns]
        fc_pivot['6개월합'] = fc_pivot.sum(axis=1)
        fc_pivot['추세'] = [f"{'▼' if slopes.get(s,0)<0 else '▲'} {abs(slopes.get(s,0)):.1f}/월" for s in fc_pivot.index]
        st.dataframe(fc_pivot, use_container_width=True)

        sel = st.selectbox("SKU 상세 차트", SKUS)
        hist = monthly_raw[monthly_raw['SKU']==sel][['월_dt','월_label','수량']].copy().sort_values('월_dt')
        fc_sel = fc_df[fc_df['SKU']==sel].copy()
        fc_sel['월_label'] = fc_sel['월'].apply(to_short_month)

        hist_chart = alt.Chart(hist).mark_line(point=True, color='#1A73E8').encode(
            x=alt.X('월_label:N', sort=None, axis=alt.Axis(labelAngle=0, title='')),
            y=alt.Y('수량:Q', title='수량'),
            tooltip=['월_label','수량']
        ).properties(title='실적')

        fc_chart = alt.Chart(fc_sel).mark_line(point=True, strokeDash=[6,3], color='#E08020').encode(
            x=alt.X('월_label:N', sort=None, axis=alt.Axis(labelAngle=0, title='')),
            y=alt.Y('예측:Q', title='수량'),
            tooltip=['월_label','예측']
        ).properties(title='예측')

        st.altair_chart((hist_chart + fc_chart).configure_view(strokeWidth=0).properties(height=320),
                        use_container_width=True)

        st.info(f"**{sel}** — 추세: 월 {slopes.get(sel,0):.1f}개 {'감소↓' if slopes.get(sel,0)<0 else '증가↑'} | 향후 6개월 합계: **{fc_df[fc_df['SKU']==sel]['예측'].sum()}개**")

# ════════════ TAB 3: 재고 관리 ═══════════════════════════════
with tab3:
    # ── 기본 재고 현황 ────────────────────────────────────────
    st.subheader("재고 현황 및 발주 관리")

    c1, c2 = st.columns([1, 3])
    with c1:
        lead_time   = st.number_input("리드타임 (일)", 1, 90, 14)
        safety_days = st.number_input("안전재고 (일치)", 1, 90, 30)

    inventory = get_inventory()
    monthly_tuple2 = [tuple(r) for r in monthly_raw[['월','SKU','수량']].itertuples(index=False)]

    daily_demand = {}; forecast_3m = {}
    for sku in SKUS:
        fc, _ = forecast_sku(tuple(monthly_tuple2), sku, 3)
        if not fc.empty:
            daily_demand[sku]  = fc['예측'].mean() / 30
            forecast_3m[sku]   = fc['예측'].sum()
        else:
            recent = monthly_raw[monthly_raw['SKU']==sku].tail(3)['수량'].mean()
            daily_demand[sku]  = (recent if not np.isnan(recent) else 0) / 30
            forecast_3m[sku]   = daily_demand[sku] * 90

    st.markdown("---")
    st.subheader("재고 수량 입력")

    alerts = []; order_data = []
    cols = st.columns(3)
    for i, sku in enumerate(SKUS):
        with cols[i % 3]:
            current = inventory.get(sku, 0)
            new_qty = st.number_input(f"**{sku}**", min_value=0, value=current, step=10, key=f"inv_{sku}")
            dd = daily_demand.get(sku, 0)
            days_left = (new_qty / dd) if dd > 0 else 999
            rop = dd * (lead_time + safety_days)
            rec_order = round(dd * 60)
            if new_qty == 0:
                st.markdown('<div class="alert-red">🔴 재고 없음 · 즉시 발주</div>', unsafe_allow_html=True)
                alerts.append((sku, '긴급', new_qty, 0, rec_order))
            elif new_qty < rop:
                st.markdown(f'<div class="alert-yellow">🟡 잔여 ~{round(days_left)}일 · 발주 권장</div>', unsafe_allow_html=True)
                alerts.append((sku, '권장', new_qty, round(days_left), rec_order))
            else:
                st.markdown(f'<div class="alert-green">✅ 잔여 ~{round(days_left)}일</div>', unsafe_allow_html=True)
            order_data.append({'SKU': sku, '현재재고': new_qty, '일평균수요': round(dd,1),
                               '예상소진(일)': round(days_left) if days_left<999 else '-',
                               '리오더포인트': round(rop), '권장발주량': rec_order,
                               '_new': new_qty, '_old': current})

    if st.button("💾 재고 저장", type="primary", disabled=LOCAL_MODE):
        saved = sum(1 for r in order_data if set_inventory(r['SKU'], r['_new']) or r['_new'] != r['_old'])
        st.success(f"✅ 업데이트 완료") if saved else st.info("변경된 항목 없음")

    if alerts:
        st.markdown("---")
        st.subheader("🚨 발주 알림")
        for sku, status, stock, days, qty in sorted(alerts, key=lambda x: x[1]):
            icon = "🔴" if status == '긴급' else "🟡"
            st.markdown(f"{icon} **{sku}** — {status} | 현재고 {stock}개 | 잔여 {days}일 | 권장발주 **{qty}개**")

    st.markdown("---")
    df_order = pd.DataFrame([{k:v for k,v in r.items() if not k.startswith('_')} for r in order_data])
    st.dataframe(df_order.set_index('SKU'), use_container_width=True)

    # ── AI 최적 재고량 (특허 수식 Rt) ─────────────────────────
    st.markdown("---")
    st.subheader("🤖 AI 최적 재고량 계산")
    st.markdown("""
<div class="formula-box">
Rt = Pt + wx · log(St + 1) + wy · (Bt − Dt) + γ · cos(2π · Vt / 100)
</div>
""", unsafe_allow_html=True)

    with st.expander("변수 설명", expanded=False):
        st.markdown("""
| 변수 | 단위 | 설명 |
|------|------|------|
| **Rt** | 개 | 최적화된 재고량 (산출값) |
| **Pt** | 개 | 과거 재고량 (이전 월 기준) |
| **St** | 개 | 최근 30일 실제 판매량 |
| **Bt** | 개 | 초기 주문량 (발주 예정량) |
| **Dt** | 개 | 예상 수요량 (수요예측부 산출) |
| **Vt** | % | 전월 대비 여행객 증가율 |
| **γ** | - | 여행객 증가율 변화에 대한 진폭 조절 계수 |
| **wx, wy** | - | 가중치 |

- **Bt − Dt > 0** → 과잉 재고 가능성 → 추가 발주 미실행
- **Bt − Dt < 0** → 재고 부족 가능성 → 추가 발주 실행
        """)

    rt_sku = st.selectbox("SKU 선택", SKUS, key='rt_sku')

    # 자동 입력값 (판매 데이터 기반)
    sku_history = monthly_raw[monthly_raw['SKU']==rt_sku].sort_values('월_dt')
    auto_Pt = int(inventory.get(rt_sku, 0))
    auto_St = int(sku_history.tail(1)['수량'].values[0]) if len(sku_history) > 0 else 0
    auto_Dt = int(round(forecast_3m.get(rt_sku, 0) / 3))

    col_a, col_b = st.columns(2)
    with col_a:
        Pt = st.number_input("Pt — 현재 재고량 (개)", min_value=0, value=auto_Pt, step=10, key='Pt')
        St = st.number_input("St — 최근 30일 판매량 (개)", min_value=0, value=auto_St, step=10, key='St')
        Bt = st.number_input("Bt — 발주 예정량 (개)", min_value=0, value=auto_Dt, step=10, key='Bt')
    with col_b:
        Dt = st.number_input("Dt — 예상 수요량 (개)", min_value=0, value=auto_Dt, step=10, key='Dt',
                             help="수요예측탭의 다음 달 예측값")
        Vt = st.number_input("Vt — 여행객 증가율 (%)", min_value=-100.0, max_value=200.0, value=5.0, step=1.0, key='Vt',
                             help="전월 대비 여행객 증가율. 증가시 양수, 감소시 음수")
        gamma = st.number_input("γ — 진폭 조절 계수", min_value=0.0, max_value=10.0, value=0.5, step=0.1, key='gamma')

    with st.expander("가중치 설정 (고급)", expanded=False):
        wx = st.slider("wx (로그 판매량 가중치)", 0.1, 5.0, 1.0, 0.1)
        wy = st.slider("wy (주문-수요 차이 가중치)", 0.1, 5.0, 1.0, 0.1)

    if st.button("🔢 최적 재고량 계산", type="primary"):
        Rt = calc_optimal_inventory(Pt, St, Bt, Dt, Vt, gamma, wx, wy)
        log_val      = wx * math.log(St + 1)
        demand_val   = wy * (Bt - Dt)
        seasonal_val = gamma * math.cos(2 * math.pi * Vt / 100)

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("최적 재고량 Rt", f"{Rt:,}개")
        r2.metric("로그항 wx·log(St+1)", f"{log_val:+.1f}")
        r3.metric("수급항 wy·(Bt-Dt)", f"{demand_val:+.1f}")
        r4.metric("계절항 γ·cos(2π·Vt/100)", f"{seasonal_val:+.1f}")

        diff = Bt - Dt
        if diff > 0:
            st.markdown(f'<div class="alert-yellow">🟡 Bt − Dt = <b>+{diff}</b> → 과잉 재고 가능성 → 추가 발주 <b>미실행</b> 권장 | 최적 재고: {Rt}개</div>', unsafe_allow_html=True)
        elif diff < 0:
            st.markdown(f'<div class="alert-red">🔴 Bt − Dt = <b>{diff}</b> → 재고 부족 가능성 → 추가 발주 <b>실행</b> 권장 | 필요 발주량: {abs(diff)}개 추가</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="alert-green">✅ Bt = Dt → 수급 균형 | 최적 재고: {Rt}개</div>', unsafe_allow_html=True)

        st.caption(f"계산식: {Rt} = {Pt} + {log_val:.2f} + {demand_val:.2f} + {seasonal_val:.2f}")

    # ── 입고 이력 ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📥 입고 이력")
    stock_in = load_stock_in()
    if not stock_in.empty:
        stock_in['월_label'] = stock_in['날짜'].apply(to_short_month)
        stock_in['날짜_str'] = stock_in['날짜'].dt.strftime('%Y-%m-%d')

        # SKU 필터
        sel_sku_in = st.selectbox("SKU 선택", ['전체'] + SKUS, key='stock_in_sku')
        if sel_sku_in != '전체':
            disp = stock_in[stock_in['SKU'] == sel_sku_in]
        else:
            disp = stock_in

        # 입고 바 차트 (SKU별 월간 합계)
        monthly_in = disp.copy()
        monthly_in['월_dt'] = monthly_in['날짜'].dt.to_period('M').dt.to_timestamp()
        monthly_in['월_label'] = monthly_in['월_dt'].apply(to_short_month)
        monthly_in_g = monthly_in.groupby(['월_label','월_dt','SKU'])['입고수량'].sum().reset_index().sort_values('월_dt')

        bar = alt.Chart(monthly_in_g).mark_bar().encode(
            x=alt.X('월_label:N', sort=None, axis=alt.Axis(labelAngle=0, title='')),
            y=alt.Y('입고수량:Q', title='입고수량'),
            color=alt.Color('SKU:N'),
            tooltip=['월_label','SKU','입고수량']
        ).properties(height=280).configure_view(strokeWidth=0)
        st.altair_chart(bar, use_container_width=True)

        # 테이블
        pivot_in = stock_in.pivot_table(index='날짜_str', columns='SKU', values='입고수량', aggfunc='sum', fill_value=0)
        pivot_in['합계'] = pivot_in.sum(axis=1)
        st.dataframe(pivot_in, use_container_width=True)

        # SKU별 총 입고 합계
        total_in = stock_in.groupby('SKU')['입고수량'].sum().reset_index()
        total_in.columns = ['SKU','총입고량']
        st.caption(f"총 입고: {total_in['총입고량'].sum():,}개")

    if not LOCAL_MODE:
        with st.expander("📒 재고 변경 이력"):
            log = get_stock_log()
            if not log.empty:
                st.dataframe(log[['sku','change_qty','note','created_at']], use_container_width=True)

# ════════════ TAB 4: AI 재고 최적화 ══════════════════════════
with tab4:
    st.subheader("🤖 AI 재고 최적화")
    st.caption(f"리드타임 {LEAD_TIME}일 고정 | 계절 조정 수요 기반 | MOQ: 색상별 {COLOR_MOQ}개 / 사이즈별 {SIZE_MOQ}개")

    inventory_now = get_inventory()

    # ── Section 1: 발주 타이밍 대시보드 ─────────────────────
    st.subheader("📅 발주 타이밍 — 언제 발주해야 하나?")
    st.caption(f"생산+배송 {LEAD_TIME}일 소요. 계절별 판매량 패턴 반영.")

    timing_df = calc_reorder_table(inventory_now, monthly_raw)

    # KPI 요약
    urgent  = timing_df[timing_df['_urgency'] == 0]
    soon    = timing_df[timing_df['_urgency'] == 1]
    warning = timing_df[timing_df['_urgency'] == 2]
    safe    = timing_df[timing_df['_urgency'] >= 3]

    t1, t2, t3, t4 = st.columns(4)
    t1.metric("🔴 즉시/지남", f"{len(urgent)}개")
    t2.metric("🟠 14일 이내", f"{len(soon)}개")
    t3.metric("🟡 30일 이내", f"{len(warning)}개")
    t4.metric("✅ 여유", f"{len(safe)}개")

    # 발주 타이밍 테이블
    display_timing = timing_df[['SKU','현재재고','일평균(계절조정)','소진예상일','발주권장일','D-day','상태']].copy()
    display_timing['D-day'] = display_timing['D-day'].apply(
        lambda x: f'D-{abs(x)} 초과' if x < 0 else (f'D-{x}' if x < 999 else '여유'))
    st.dataframe(display_timing.set_index('SKU'), use_container_width=True)

    # 알림 카드
    for _, row in timing_df.iterrows():
        if row['_urgency'] == 0:
            st.markdown(
                f'<div class="alert-red">🔴 <b>{row["SKU"]}</b> — {row["상태"]} | '
                f'현재고 {row["현재재고"]}개 | 소진 {row["소진예상일"]} | '
                f'<b>발주일: {row["발주권장일"]}</b></div>',
                unsafe_allow_html=True)
        elif row['_urgency'] == 1:
            st.markdown(
                f'<div class="alert-yellow">🟠 <b>{row["SKU"]}</b> — {row["상태"]} | '
                f'현재고 {row["현재재고"]}개 | 소진 {row["소진예상일"]} | '
                f'<b>발주일: {row["발주권장일"]}</b></div>',
                unsafe_allow_html=True)

    # 계절 지수 시각화
    with st.expander("📊 월별 계절 지수 (판매 패턴)"):
        sel_sku_s = st.selectbox("SKU 선택", SKUS, key='seasonal_sku')
        sf = get_seasonal_factors(monthly_raw, sel_sku_s)
        sf_df = pd.DataFrame({'월': [f'{m}월' for m in range(1,13)],
                              '계절지수': [round(sf[m], 2) for m in range(1,13)]})
        bar_s = alt.Chart(sf_df).mark_bar().encode(
            x=alt.X('월:N', sort=None, axis=alt.Axis(labelAngle=0)),
            y=alt.Y('계절지수:Q', scale=alt.Scale(domain=[0, 2.5])),
            color=alt.condition(
                alt.datum['계절지수'] >= 1.0,
                alt.value('#1A73E8'), alt.value('#9EC8F5')),
            tooltip=['월','계절지수']
        ).properties(height=240).configure_view(strokeWidth=0)
        st.altair_chart(bar_s, use_container_width=True)
        st.caption("1.0 = 평균. 1.5 = 평균보다 50% 더 팔림. 0.5 = 평균의 절반.")

    st.markdown("---")

    # ── Section 2: Rt 최적 재고 계산 ─────────────────────────
    st.subheader("🔢 AI 최적 재고량 계산 (Rt)")
    st.markdown("""
<div class="formula-box">
Rt = Pt + wx · log(St + 1) + wy · (Bt − Dt) + γ · cos(2π · Vt / 100)
</div>
""", unsafe_allow_html=True)
    st.markdown("")

    col_p1, col_p2, col_p3, col_p4 = st.columns(4)
    with col_p1:
        Vt_global = st.slider("Vt — 여행객 증가율 (%)", -50.0, 100.0, 5.0, 1.0)
    with col_p2:
        gamma_global = st.slider("γ — 진폭 조절 계수", 0.0, 10.0, 0.5, 0.1)
    with col_p3:
        wx_global = st.slider("wx — 판매량 가중치", 0.1, 5.0, 1.0, 0.1)
    with col_p4:
        wy_global = st.slider("wy — 수급 가중치", 0.1, 5.0, 1.0, 0.1)

    monthly_tuple3 = [tuple(r) for r in monthly_raw[['월','SKU','수량']].itertuples(index=False)]
    fc3m = {}
    for sku in SKUS:
        fc, _ = forecast_sku(tuple(monthly_tuple3), sku, 3)
        fc3m[sku] = fc['예측'].sum() if not fc.empty else 0

    def calc_all_rt(inventory, monthly_raw, forecast_3m, Vt, gamma, wx, wy):
        rows = []
        for sku in SKUS:
            Pt = inventory.get(sku, 0)
            last = monthly_raw[monthly_raw['SKU']==sku].sort_values('월_dt').tail(1)
            St = int(last['수량'].values[0]) if len(last) > 0 else 0
            Dt = max(1, round(forecast_3m.get(sku, 0) / 3))
            Rt = calc_optimal_inventory(Pt, St, Dt, Dt, Vt, gamma, wx, wy)
            diff = Rt - Pt
            판단 = '발주 필요 🔴' if diff > 20 else ('과잉 🟡' if diff < -50 else '적정 ✅')
            rows.append({'SKU': sku, 'Pt 현재재고': Pt, 'St 최근월판매': St,
                         'Dt 예측수요/월': Dt, 'Rt 최적재고': Rt,
                         '차이(Rt-Pt)': diff, '상태': 판단, '권장발주량': max(0, diff)})
        return pd.DataFrame(rows)

    rt_df = calc_all_rt(inventory_now, monthly_raw, fc3m, Vt_global, gamma_global, wx_global, wy_global)
    st.dataframe(rt_df.set_index('SKU'), use_container_width=True)

    # 현재재고 vs Rt 차트
    chart_df = pd.concat([
        rt_df[['SKU','Pt 현재재고']].rename(columns={'Pt 현재재고':'수량'}).assign(구분='현재 재고'),
        rt_df[['SKU','Rt 최적재고']].rename(columns={'Rt 최적재고':'수량'}).assign(구분='최적 재고(Rt)'),
    ])
    bar_cmp = alt.Chart(chart_df).mark_bar().encode(
        x=alt.X('SKU:N', axis=alt.Axis(labelAngle=0)),
        y=alt.Y('수량:Q'),
        color=alt.Color('구분:N', scale=alt.Scale(
            domain=['현재 재고','최적 재고(Rt)'], range=['#9EC8F5','#1A73E8'])),
        xOffset='구분:N', tooltip=['SKU','구분','수량']
    ).properties(height=300).configure_view(strokeWidth=0)
    st.altair_chart(bar_cmp, use_container_width=True)

    st.markdown("---")

    # ── Section 3: MOQ + 발주량 종합 ─────────────────────────
    st.subheader("📦 발주량 종합 (MOQ 참고)")
    st.caption(f"MOQ 미달이어도 발주 가능. 색상별 {COLOR_MOQ}개 / 사이즈별 {SIZE_MOQ}개는 참고 기준.")

    # 발주 타이밍 + Rt 권장량 합산
    combined = timing_df[['SKU','현재재고','발주권장일','D-day','상태']].merge(
        rt_df[['SKU','Rt 최적재고','권장발주량']], on='SKU')
    combined['색상'] = combined['SKU'].str.split('/').str[0]
    combined['사이즈'] = combined['SKU'].str.split('/').str[1]

    color_order = combined.groupby('색상')['권장발주량'].sum()
    size_order  = combined.groupby('사이즈')['권장발주량'].sum()

    combined['색상별 합계'] = combined['색상'].map(color_order)
    combined['사이즈별 합계'] = combined['사이즈'].map(size_order)

    def moq_note(row):
        notes = []
        if row['권장발주량'] > 0:
            if row['색상별 합계'] < COLOR_MOQ:
                notes.append(f"색상 {int(row['색상별 합계'])}/{COLOR_MOQ}")
            if row['사이즈별 합계'] < SIZE_MOQ:
                notes.append(f"사이즈 {int(row['사이즈별 합계'])}/{SIZE_MOQ}")
        return '⚠️ MOQ 미달 참고 (' + ', '.join(notes) + ')' if notes else ('✅' if row['권장발주량'] > 0 else '-')

    combined['MOQ 참고'] = combined.apply(moq_note, axis=1)
    combined['D-day'] = combined['D-day'].apply(
        lambda x: f'D-{abs(x)} 초과' if x < 0 else (f'D-{x}' if x < 999 else '여유'))

    st.dataframe(
        combined[['SKU','현재재고','발주권장일','D-day','Rt 최적재고','권장발주량','MOQ 참고']].set_index('SKU'),
        use_container_width=True)

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.markdown("**색상별 발주 합계 vs MOQ**")
        for color, total in color_order.items():
            if total > 0:
                status = "✅" if total >= COLOR_MOQ else f"⚠️ {int(COLOR_MOQ-total)}개 부족"
                st.markdown(f"**{color}** — {int(total):,}개 {status}")
                st.progress(min(1.0, total / COLOR_MOQ))
    with col_m2:
        st.markdown("**사이즈별 발주 합계 vs MOQ**")
        for size, total in size_order.items():
            if total > 0:
                status = "✅" if total >= SIZE_MOQ else f"⚠️ {int(SIZE_MOQ-total)}개 부족"
                st.markdown(f"**{size}** — {int(total):,}개 {status}")
                st.progress(min(1.0, total / SIZE_MOQ))

    st.markdown("---")

    # ── 이력 저장 ─────────────────────────────────────────────
    if st.button("💾 현재 계산 이력 저장", type="primary"):
        history_path = 'data/rt_history.csv'
        save_row = rt_df.copy()
        save_row['계산일시'] = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')
        save_row['Vt'] = Vt_global; save_row['gamma'] = gamma_global
        existing = pd.read_csv(history_path, encoding='utf-8-sig') if os.path.exists(history_path) else pd.DataFrame()
        pd.concat([existing, save_row], ignore_index=True).to_csv(history_path, index=False, encoding='utf-8-sig')
        st.success("✅ 이력 저장 완료")

    history_path = 'data/rt_history.csv'
    if os.path.exists(history_path):
        with st.expander("📒 과거 계산 이력"):
            hist_df = pd.read_csv(history_path, encoding='utf-8-sig')
            dates = hist_df['계산일시'].unique().tolist()
            sel_date = st.selectbox("날짜 선택", dates[::-1])
            st.dataframe(hist_df[hist_df['계산일시']==sel_date].set_index('SKU'), use_container_width=True)

# ════════════ TAB 5: 엑셀 업로드 ════════════════════════════
with tab5:
    st.subheader("📂 네이버 스마트스토어 주문 데이터 업로드")
    st.caption("스마트스토어 > 주문관리 > 발주·발송관리 > 엑셀 다운로드 후 업로드")

    uploaded = st.file_uploader("엑셀 파일 선택 (여러 개 동시 가능)", type=['xlsx','xls'], accept_multiple_files=True)
    if uploaded:
        all_parsed = []
        for f in uploaded:
            try:
                parsed = parse_excel(io.BytesIO(f.read()), 0)
                if not parsed.empty:
                    all_parsed.append(parsed)
                    st.success(f"✅ {f.name} — {len(parsed)}건 파싱 완료")
                else:
                    st.warning(f"⚠️ {f.name} — 유효한 데이터 없음")
            except Exception as e:
                st.error(f"❌ {f.name} — {e}")

        if all_parsed:
            combined = pd.concat(all_parsed, ignore_index=True)
            combined['월'] = combined['날짜'].dt.to_period('M').astype(str)
            monthly_up = combined.groupby(['월','SKU'])['수량'].sum().reset_index()
            st.subheader("업로드 데이터 미리보기")
            pivot = monthly_up.pivot_table(index='월', columns='SKU', values='수량', fill_value=0)
            st.dataframe(pivot, use_container_width=True)
            st.info(f"총 {combined['수량'].sum():,}개 | {combined['날짜'].min().strftime('%Y.%m')} ~ {combined['날짜'].max().strftime('%Y.%m')}")
