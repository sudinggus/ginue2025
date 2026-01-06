import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from collections import defaultdict

# ==========================================
# 1. 페이지 설정 및 디자인
# ==========================================
st.set_page_config(page_title="근무표 자동화 시스템", layout="wide")

st.markdown("""
    <style>
        .stTable { border: 1px solid #333; font-size: 14px; }
        th { background-color: #F2F2F2 !important; color: black !important; font-weight: bold !important; text-align: center !important; border: 1px solid #333 !important; }
        td { border: 1px solid #333 !important; text-align: center !important; }
        .stButton>button { width: 100%; }
    </style>
""", unsafe_allow_html=True)

# 설정값
LOCATIONS_CONFIG = {
    "인천": {"생활관1": 2, "생활관2": 2, "생활관3": 2, "상황실1": 3, "도서관1": 2},
    "경기": {"생활관1": 2, "생활관2": 2, "상황실2": 3, "도서관2": 2}
}
HOLIDAYS = ['2025-10-03', '2025-10-06', '2025-10-09']

# ==========================================
# 2. 핵심 로직 (코랩 배정 엔진)
# ==========================================

def get_korean_weekday(date_obj):
    return ['월', '화', '수', '목', '금', '토', '일'][date_obj.weekday()]

def generate_schedule_logic(df_staff, start_dt, end_dt):
    df_staff['이름'] = df_staff['이름'].astype(str).str.strip()
    work_counts = {name: 0 for name in df_staff['이름'].unique()}
    schedule_results = []
    
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

# ==========================================
# 3. 엑셀 생성 (코랩 스타일 복원)
# ==========================================

def make_final_excel_blob(df, stats):
    output = BytesIO()
    wb = Workbook()
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    # 시트 1: 주간근무표 (시각화 시트)
    ws1 = wb.active
    ws1.title = "주간근무표"
    dates = sorted(df['날짜'].unique())
    curr_r = 1
    
    for d_str in dates:
        dt_obj = datetime.strptime(d_str, '%Y-%m-%d')
        ws1.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=6)
        cell = ws1.cell(row=curr_r, column=1, value=f"{d_str}({get_korean_weekday(dt_obj)}) 근무표")
        cell.alignment = center; cell.fill = header_fill; cell.font = Font(bold=True)
        curr_r += 1
        
        headers = ["캠퍼스", "도서관", "상황실", "생활관1", "생활관2", "생활관3"]
        for c_idx, h in enumerate(headers, 1):
            cell = ws1.cell(row=curr_r, column=c_idx, value=h)
            cell.alignment = center; cell.border = border; cell.fill = header_fill
        curr_r += 1
        
        for cp in ["인천", "경기"]:
            ws1.cell(row=curr_r, column=1, value=cp).border = border
            for c_idx, loc_b in enumerate(["도서관", "상황실", "생활관1", "생활관2", "생활관3"], 2):
                loc_f = loc_b + ("1" if cp=="인천" and "생활관" not in loc_b else ("2" if cp=="경기" and "생활관" not in loc_b else ""))
                names = df[(df['날짜']==d_str) & (df['캠퍼스']==cp) & (df['근무지']==loc_f)]['직원'].tolist()
                ws1.cell(row=curr_r, column=c_idx, value=", ".join(names)).border = border
                ws1.cell(row=curr_r, column=c_idx).alignment = center
            curr_r += 1
        curr_r += 1 # 공백행

    # 시트 2: 근무통계
    ws2 = wb.create_sheet("근무통계")
    ws2.append(["직원 이름", "총 근무 횟수"])
    for name, count in sorted(stats.items(), key=lambda x: x[1], reverse=True):
        ws2.append([name, count])

    wb.save(output)
    return output.getvalue()

# ==========================================
# 4. Streamlit UI (게시판 및 관리)
# ==========================================

if 'df' not in st.session_state: st.session_state.df = None
if 'stats' not in st.session_state: st.session_state.stats = {}

with st.sidebar:
    st.title("🔐 관리자 인증")
    pw = st.text_input("암호를 입력하세요", type="password")
    if pw == "1234":
        st.success("인증됨")
        file = st.file_uploader("명단 업로드", type=['xlsx'])
        s_d = st.date_input("시작일", datetime.today())
        e_d = st.date_input("종료일", datetime.today() + timedelta(days=7))
        if st.button("신규 근무표 생성"):
            if file:
                in_df = pd.read_excel(file)
                res_df, res_stats = generate_schedule_logic(in_df, s_d, e_d)
                st.session_state.df = res_df.reset_index(drop=True)
                st.session_state.stats = res_stats
                st.rerun()

st.title("📢 실시간 근무 게시판")

if st.session_state.df is not None:
    df = st.session_state.df
    
    # 상단 버튼 (엑셀 다운로드)
    excel_bin = make_final_excel_blob(df, st.session_state.stats)
    st.download_button(label="📥 코랩 스타일 엑셀 다운로드 (전체 시트 포함)", 
                       data=excel_bin, file_name="근무표_최종.xlsx")
    
    # 관리자 전용 교체 기능
    if pw == "1234":
        with st.expander("🔄 1:1 인원 교체"):
            c1, c2 = st.columns(2)
            idx1 = c1.selectbox("첫번째 셀", df.index, format_func=lambda x: f"{df.loc[x,'날짜']} {df.loc[x,'직원']}")
            idx2 = c2.selectbox("두번째 셀", df.index, format_func=lambda x: f"{df.loc[x,'날짜']} {df.loc[x,'직원']}")
            if st.button("교체 확정"):
                df.at[idx1, '직원'], df.at[idx2, '직원'] = df.at[idx2, '직원'], df.at[idx1, '직원']
                st.session_state.df = df
                st.rerun()

    # 웹 화면 시각화 (코랩 엑셀 시트와 동일한 구조)
    st.divider()
    dates = sorted(df['날짜'].unique())
    for d_str in dates:
        st.subheader(f"🗓️ {d_str} ({get_korean_weekday(datetime.strptime(d_str, '%Y-%m-%d'))})")
        disp = []
        for cp in ["인천", "경기"]:
            r = {"캠퍼스": cp}
            for loc_b in ["도서관", "상황실", "생활관1", "생활관2", "생활관3"]:
                loc_f = loc_b + ("1" if cp=="인천" and "생활관" not in loc_b else ("2" if cp=="경기" and "생활관" not in loc_b else ""))
                names = df[(df['날짜']==d_str) & (df['캠퍼스']==cp) & (df['근무지']==loc_f)]['직원'].tolist()
                r[loc_b] = ", ".join(names)
            disp.append(r)
        st.table(pd.DataFrame(disp))
else:
    st.info("관리자 메뉴에서 엑셀 파일을 업로드하여 근무표를 생성해주세요.")