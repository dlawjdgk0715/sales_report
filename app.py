# app.py  —  별도매출현황 자동생성 웹앱 (Streamlit)

import streamlit as st
import pandas as pd
import tempfile
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from processor import load_grid_files, apply_mapping, find_unmapped_vendors, aggregate_by_month_vendor, append_to_raw_data
from excel_gen import generate_excel
from config import ACCOUNT_LABEL, GAME_SUB_ORDER, ETC_SUB_ORDER

# ── 페이지 설정 ───────────────────────────────────────────────────
st.set_page_config(
    page_title="별도매출현황 자동생성",
    page_icon="📊",
    layout="wide",
)

st.title("📊 별도매출현황 자동생성")
st.caption("ERP에서 추출한 grid 파일을 업로드하면 별도매출현황.xlsx를 자동 생성합니다.")

# ── 세션 상태 초기화 ──────────────────────────────────────────────
if "extra_map" not in st.session_state:
    st.session_state.extra_map = {}   # {계정코드: {거래처명: (표시명, 세부구분, 비고)}}
if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "step" not in st.session_state:
    st.session_state.step = 1

# ── STEP 1: 파일 업로드 ───────────────────────────────────────────
st.header("① 파일 업로드", divider="blue")

col1, col2 = st.columns([2, 1])
with col1:
    grid_files = st.file_uploader(
        "grid 파일 업로드 (복수 선택 가능)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="grid_upload",
        help="ERP에서 추출한 GLDDCM00200 grid 파일을 모두 선택하세요."
    )
with col2:
    prev_file = st.file_uploader(
        "기존 별도매출현황.xlsx (선택 — 누적 갱신 시)",
        type=["xlsx"],
        key="prev_upload",
        help="이전 분기 파일을 업로드하면 Raw_Data를 누적 append합니다."
    )

if grid_files:
    st.success(f"✅ grid 파일 {len(grid_files)}개 업로드됨")
    for f in grid_files:
        st.write(f"  • {f.name}")

# ── STEP 2: 데이터 처리 및 신규 거래처 분류 ──────────────────────
if grid_files and st.button("🔍 데이터 분석 시작", type="primary", use_container_width=True):
    with st.spinner("grid 파일 읽는 중..."):
        # 임시 파일로 저장 후 처리
        tmp_paths = []
        for uf in grid_files:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            tmp.write(uf.read())
            tmp.close()
            tmp_paths.append(tmp.name)

        combined = load_grid_files(tmp_paths)
        for p in tmp_paths:
            os.unlink(p)

        if combined.empty:
            st.error("유효한 거래 데이터가 없습니다.")
            st.stop()

        # 기존 파일에서 Raw_Data 불러오기
        existing_raw = None
        if prev_file:
            try:
                existing_raw = pd.read_excel(prev_file, sheet_name="Raw_Data")
                st.info(f"기존 Raw_Data {len(existing_raw):,}행 로드됨")
            except Exception:
                st.warning("기존 파일에서 Raw_Data 시트를 찾지 못했습니다. 신규 생성합니다.")

        # 매핑 적용
        mapped = apply_mapping(combined, st.session_state.extra_map)

        # 누적 append
        full_raw = append_to_raw_data(existing_raw, mapped)

        st.session_state.raw_df = full_raw
        st.session_state.step = 2

    st.rerun()

# ── STEP 2: 신규 거래처 분류 UI ──────────────────────────────────
if st.session_state.step >= 2 and st.session_state.raw_df is not None:
    raw_df = st.session_state.raw_df

    # 매핑 재적용 (세션에 추가된 것 반영)
    raw_df = apply_mapping(
        raw_df.drop(columns=["매출구분","세부구분","표시명","월","비고_매핑"], errors="ignore"),
        st.session_state.extra_map
    )
    st.session_state.raw_df = raw_df

    unmapped = find_unmapped_vendors(raw_df)

    if unmapped:
        st.header("② 신규 거래처 분류", divider="orange")
        st.warning(f"⚠️ 매핑 정보가 없는 거래처 **{len(unmapped)}개**가 발견되었습니다. 분류해주세요.")

        # 분류 입력 UI
        sub_options = {
            "게임매출": GAME_SUB_ORDER,
            "기타수익": ETC_SUB_ORDER,
            "용역수익": ["경영관리수익"],
        }

        changes = {}
        for acct, vendor, label in unmapped:
            with st.expander(f"📌 [{label}] {vendor}", expanded=True):
                c1, c2, c3 = st.columns([2, 2, 3])
                with c1:
                    cat = st.selectbox(
                        "매출구분",
                        ["게임매출", "기타수익", "용역수익"],
                        key=f"cat_{acct}_{vendor}",
                        index=["게임매출","기타수익","용역수익"].index(label) if label in ["게임매출","기타수익","용역수익"] else 0
                    )
                with c2:
                    subs = sub_options.get(cat, ["기타"])
                    sub = st.selectbox("세부구분", subs, key=f"sub_{acct}_{vendor}")
                with c3:
                    display = st.text_input(
                        "표시명 (요약표에 표시될 이름)",
                        value=vendor[:20],
                        key=f"disp_{acct}_{vendor}"
                    )
                    note = st.text_input("비고", value="", key=f"note_{acct}_{vendor}")
                changes[(str(acct), vendor)] = (display, sub, note)

        if st.button("✅ 분류 저장 후 계속", type="primary"):
            for (acct, vendor), (display, sub, note) in changes.items():
                if acct not in st.session_state.extra_map:
                    st.session_state.extra_map[acct] = {}
                st.session_state.extra_map[acct][vendor] = (display, sub, note)
            st.session_state.step = 3
            st.rerun()
    else:
        st.session_state.step = 3

# ── STEP 3: Excel 생성 및 다운로드 ───────────────────────────────
if st.session_state.step >= 3 and st.session_state.raw_df is not None:
    st.header("③ Excel 생성 및 다운로드", divider="green")

    raw_df = apply_mapping(
        st.session_state.raw_df.drop(
            columns=["매출구분","세부구분","표시명","월","비고_매핑"], errors="ignore"
        ),
        st.session_state.extra_map
    )

    agg_df = aggregate_by_month_vendor(raw_df)

    # 데이터 미리보기
    months = sorted(
        agg_df["월"].dropna().unique().tolist(),
        key=lambda x: int(x.replace("월",""))
    )
    st.write(f"**처리 결과:** {len(months)}개월 데이터, 총 {len(raw_df):,}행")

    col1, col2 = st.columns(2)
    with col1:
        st.write("**포함 월:**", ", ".join(months))
    with col2:
        by_label = raw_df.groupby("매출구분")["대변"].sum()
        for label, val in by_label.items():
            st.write(f"  • {label}: {val:,.0f}원")

    # 생성 버튼
    if st.button("📥 별도매출현황.xlsx 생성", type="primary", use_container_width=True):
        with st.spinner("Excel 파일 생성 중..."):
            out_path = tempfile.mktemp(suffix=".xlsx")
            generate_excel(raw_df, agg_df, out_path)

        with open(out_path, "rb") as f:
            data = f.read()
        os.unlink(out_path)

        year = 2026
        if not raw_df.empty and "회계일" in raw_df.columns:
            dates = pd.to_datetime(raw_df["회계일"], errors="coerce").dropna()
            if not dates.empty:
                year = int(dates.dt.year.mode()[0])

        fname = f"별도매출현황_{year}_{'-'.join(months)}.xlsx"

        st.download_button(
            label=f"⬇️ {fname} 다운로드",
            data=data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.success("✅ 파일 생성 완료!")

    # 초기화 버튼
    if st.button("🔄 처음부터 다시", use_container_width=True):
        for key in ["extra_map", "raw_df", "step"]:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

# ── 사이드바: 사용 가이드 ─────────────────────────────────────────
with st.sidebar:
    st.header("📖 사용 가이드")
    st.markdown("""
**기본 사용법**
1. ERP에서 추출한 grid 파일 업로드 (여러 개 동시 선택 가능)
2. 신규 거래처가 있으면 분류 입력
3. Excel 파일 다운로드

**분기별 누적 갱신**
- 기존 파일을 함께 업로드하면 Raw_Data에 자동 누적됩니다

**생성 파일 구조**
- `거래처별요약`: 주요 게임매출 거래처 × 월별 (백만원)
- `N월`: 게임매출 + 기타수익 + 용역수익 상세
- `Raw_Data`: 전체 원본 데이터 누적

**수식 연결 구조**
- 모든 월별 시트의 금액은 Raw_Data를 SUMIFS로 참조
- Raw_Data만 갱신하면 월별 시트가 자동 업데이트됩니다
""")
    st.info("grid 파일 계정코드:\n- 4100100 = 게임매출\n- 4100300 = 용역수익\n- 4100500 = 기타수익")
