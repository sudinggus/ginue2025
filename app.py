import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from collections import defaultdict

# ==========================================
# 1. 초기 설정 및 페이지 레이아웃
# ==========================================
st.set_page_config(page_title="근무표 자동화 시스템", layout="wide")

# CSS: 표 디자인 및 가독성 향상
st.markdown("""
    <style>
        .stTable { border: 1px solid #333; }
        th { background-color: #f2f2f2 !important; color: black !important; text-align: center !important; }
        td { text-align: center !important; }
        .stButton>button { width: 100%; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

# 상수 설정
LOCATIONS_CONFIG = {
    "인천": {"생활관1": 2, "생활관2": 2, "생활관3": 2, "상황실1": 3, "도서관1": 2},
    "경기": {"생활관1": 2, "생활관2": 2, "상황실2": 3, "도서관2": 2}
}
HOLIDAYS = ['2025-10-03', '2025-10-06', '2025-10-09']

# ==========================================
# 2. 핵심 로직 엔진 (코랩 코드 이식)
# ==========================================

def get_korean_weekday(date_obj):
    return ['월', '화', '수', '목', '금', '토', '일'][date_obj.weekday()]

def generate_schedule_logic(df_staff, start_dt, end_dt):
    """코랩에서 사용하던 배정 알고리즘"""
    df_staff['이름'] = df_staff['이름'].astype(str).str.strip()
    work_counts = {name: 0 for name in df_staff['이름'].unique()}
    schedule_results = []
    
    # [생략되지 않은 전체 로직 구현부]
    fixed_assignments = defaultdict(list)
    for _, row in df_staff.iterrows():
        if pd.notna(row.get('고정근무일자')):
            raw_dates = str(row['고정근무일자']).split(',')
            raw_locs = str(row['고정근무지']).split(',') if pd.notna(row.get('고정근무지')) else []
            for i, d_str in enumerate(raw_dates):
                try:
                    clean_date = datetime.strptime(d_str.strip(), '%Y-%m-%d').strftime('%Y-%m-%d')
                    loc_target = raw_locs[i].strip() if i < len(raw_locs) else (raw_locs[0].strip() if raw_locs else "미지정")
                    fixed_assignments[clean_date].append((row['이름'], loc_target, row['캠퍼스']))
                    work_counts[row['이름']] += 1
                except: continue

    date_range = []
    curr = start_dt
    while curr <= end_dt:
        if curr.weekday() < 5 and curr.strftime("%Y-%m-%d") not in HOLIDAYS:
            date_range.append(curr)
        curr += timedelta(days=1)

    for date in date_range:
        date_str = date.strftime("%Y-%m-%d")
        today_assigned = []
        if date_str in fixed_assignments:
            for name, loc, campus in fixed_assignments[date_str]:
                schedule_results.append({"날짜": date_str, "캠퍼스": campus, "근무지": loc, "직원": name, "유형": "고정"})
                today_assigned.append(name)

        for campus, locs in LOCATIONS_CONFIG.items():
            for loc_name, total_required in locs.items():
                already_filled = len([s for s in schedule_results if s['날짜'] == date_str and s['캠퍼스'] == campus and s['근무지'] == loc_name])
                needed = total_required - already_filled
                if needed <= 0: continue
                
                possible_staff = df_staff[((df_staff['캠퍼스'] == campus) | (df_staff['캠퍼스'] == "모두")) & (~df_staff['이름'].isin(today_assigned))]
                final_candidates = []
                for _, s_row in possible_staff.iterrows():
                    dept = str(s_row['소속'])
                    is_excluded = any(key in dept and key in loc_name for key in ['생활관', '상황실', '도서관'])
                    if not is_excluded: final_candidates.append(s_row['이름'])

                random.shuffle(final_candidates)
                final_candidates.sort(key=lambda x: work_counts[x])
                assigned_now = final_candidates[:needed]
                for person in assigned_now:
                    schedule_results.append({"날짜": date_str, "캠퍼스": campus, "근무지": loc_name, "직원": person, "유형": "일반"})
                    work_counts[person] += 1
                    today_assigned.append(person)

    return pd.DataFrame(schedule_results), work_counts

def make_final_excel_blob(df, stats):
    """코랩에서 사용하던 엑셀 시각화 및 멀티 시트 생성"""
    output = BytesIO()
    wb = Workbook()
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    # 시트 1: 원본 데이터
    ws_raw = wb.active
    ws_raw.title = "Schedule"
    headers = ["날짜", "캠퍼스", "근무지", "직원", "유형"]
    ws_raw.append(headers)
    for _, row in df.iterrows():
        ws_raw.append(row.tolist())

    # 시트 2: 근무통계
    ws_stat = wb.create_sheet("근무통계")
    ws_stat.append(["직원 이름", "횟수"])
    for name, count in stats.items():
        ws_stat.append([name, count])

    wb.save(output)
    return output.getvalue()

# ==========================================
# 3. 세션 관리 및 UI 구성
# ==========================================

if 'df' not in st.session_state:
    st.session_state.df = None
if 'stats' not in st.session_state:
    st.session_state.stats = {}

# --- 사이드바 (관리자 도구) ---
with st.sidebar:
    st.title("🔐 관리자 제어")
    pw = st.text_input("관리자 암호", type="password")
    
    if pw == "1234": # 암호 설정
        st.success("인증 성공")
        file = st.file_uploader("명단 파일(xlsx) 업로드", type=['xlsx'])
        s_date = st.date_input("시작일", datetime.today())
        e_date = st.date_input("종료일", datetime.today() + timedelta(days=14))
        
        if st.button("🚀 근무표 생성 및 게시"):
            if file:
                input_df = pd.read_excel(file)
                res_df, res_stats = generate_schedule_logic(input_df, s_date, e_date)
                st.session_state.df = res_df.reset_index(drop=True)
                st.session_state.stats = res_stats
                st.rerun()
    else:
        st.info("암호를 입력하면 관리 기능을 사용할 수 있습니다.")

# --- 메인 화면 (직원 게시판) ---
st.title("📢 실시간 근무 게시판")

if st.session_state.df is not None:
    df = st.session_state.df
    
    # 상단 도구 (다운로드 및 교체 신청)
    col1, col2 = st.columns([1, 1])
    with col1:
        excel_data = make_final_excel_blob(df, st.session_state.stats)
        st.download_button("📥 코랩 스타일 엑셀 다운로드", excel_data, 
                           file_name=f"근무표_{datetime.now().strftime('%m%d')}.xlsx")
    
    with col2:
        with st.expander("🔄 1:1 교체 신청 (관리자용)"):
            if pw == "1234":
                idx1 = st.selectbox("대상자 1", df.index, format_func=lambda x: f"{df.loc[x, '날짜']} {df.loc[x, '직원']}")
                idx2 = st.selectbox("대상자 2", df.index, format_func=lambda x: f"{df.loc[x, '날짜']} {df.loc[x, '직원']}")
                if st.button("교체 확정"):
                    df.at[idx1, '직원'], df.at[idx2, '직원'] = df.at[idx2, '직원'], df.at[idx1, '직원']
                    st.session_state.df = df
                    st.success("교체되었습니다!")
                    st.rerun()
            else:
                st.warning("교체 권한이 없습니다.")

    # 주간 근무표 시각화 (Pivot Table)
    st.subheader("🗓️ 주간 근무 현황")
    try:
        pivot_view = df.pivot_table(
            index=['캠퍼스', '근무지'],
            columns='날짜',
            values='직원',
            aggfunc=lambda x: ", ".join(x)
        ).fillna("-")
        st.table(pivot_view)
    except:
        st.dataframe(df) # 피벗 에러 시 기본 표 출력

    # 본인 검색 기능
    st.divider()
    search = st.text_input("🔍 내 이름으로 근무 찾기", "")
    if search:
        mine = df[df['직원'].str.contains(search)]
        st.write(f"'{search}'님의 근무 일정:")
        st.table(mine)

else:
    st.warning("현재 게시된 근무표가 없습니다. 관리자가 명단을 업로드해야 합니다.")