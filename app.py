# app.py  —  별도매출현황 자동생성 웹앱 (Streamlit)

import streamlit as st
import pandas as pd
import tempfile
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(__file__))

from processor import load_grid_files, apply_mapping, find_unmapped_vendors, aggregate_by_month_vendor, append_to_raw_data, read_legacy_file
from excel_gen import generate_excel
from config import ACCOUNT_LABEL, GAME_SUB_ORDER, ETC_SUB_ORDER

# ── 상수 ─────────────────────────────────────────────────────────
DEFAULT_BASE_PATH = os.path.join(os.path.dirname(__file__), "default_base.xlsx")

# ── 페이지 설정 ───────────────────────────────────────────────────
st.set_page_config(
    page_title="별도매출현황 자동생성",
    page_icon="📊",
    layout="wide",
)

st.title("📊 별도매출현황 자동생성")
st.caption("FI시스템 원장 파일을 업로드하면 별도매출현황.xlsx를 자동 생성합니다.")

# ── 세션 상태 초기화 ──────────────────────────────────────────────
if "extra_map"        not in st.session_state: st.session_state.extra_map        = {}
if "raw_df"           not in st.session_state: st.session_state.raw_df           = None
if "step"             not in st.session_state: st.session_state.step             = 1
if "use_default_base" not in st.session_state: st.session_state.use_default_base = True
if "custom_base_path" not in st.session_state: st.session_state.custom_base_path = None


# ── STEP 1: 파일 업로드 ───────────────────────────────────────────
st.header("① 파일 업로드", divider="blue")

# 1-A. 기존 별도매출현황 ──────────────────────────────────────────
st.subheader("(별도)매출현황 파일")

if st.session_state.use_default_base:
    try:
        _preview = pd.read_excel(DEFAULT_BASE_PATH, sheet_name="Raw_Data",
                                  dtype={"계정코드": str})
        _months  = sorted(_preview["월"].dropna().unique(),
                          key=lambda x: int(x.replace("월","")))
        st.success(
            f"✅ 기본 파일 로드됨 — "
            f"{len(_months)}개월 ({', '.join(_months)}), {len(_preview):,}행"
        )
    except Exception as e:
        st.warning(f"기본 파일을 읽을 수 없습니다: {e}")

    col_a, _ = st.columns([1, 3])
    with col_a:
        if st.button("✕ 기본 파일 해제", help="기본 파일 대신 직접 업로드합니다"):
            st.session_state.use_default_base = False
            if st.session_state.custom_base_path and os.path.exists(st.session_state.custom_base_path):
                os.unlink(st.session_state.custom_base_path)
            st.session_state.custom_base_path = None
            st.rerun()
else:
    prev_file = st.file_uploader(
        "(별도)매출현황 파일 업로드 (선택)",
        type=["xlsx"],
        key="prev_upload",
        help="직전 갱신한 별도매출현황 파일을 업로드하면 Raw_Data에 신규 데이터를 누적합니다."
    )
    col_a, _ = st.columns([1, 3])
    with col_a:
        if st.button("↩ 기본 파일 복원"):
            st.session_state.use_default_base = True
            st.rerun()

    if prev_file:
        if st.session_state.custom_base_path and os.path.exists(st.session_state.custom_base_path):
            os.unlink(st.session_state.custom_base_path)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(prev_file.read())
        tmp.close()
        st.session_state.custom_base_path = tmp.name
        try:
            _preview = pd.read_excel(tmp.name, sheet_name="Raw_Data", dtype={"계정코드": str})
            _months  = sorted(_preview["월"].dropna().unique(),
                              key=lambda x: int(x.replace("월","")))
            st.success(f"✅ {prev_file.name} — {len(_months)}개월 ({', '.join(_months)}), {len(_preview):,}행")
        except Exception:
            st.info("Raw_Data 시트 없음 — 처리 시 월별 시트에서 자동 변환합니다.")

st.divider()

# 1-B. 신규 원장 파일 ─────────────────────────────────────────────
st.subheader("원장 파일 (FI시스템 추출본)")
grid_files = st.file_uploader(
    "원장 파일 업로드 (복수 선택 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="grid_upload",
    help="FI시스템 > 계정별원장에서 당분기 기간, 컨텐츠제공수익·용역수익·기타수익 계정 선택 후 다운로드한 파일"
)
if grid_files:
    st.success(f"✅ 원장 파일 {len(grid_files)}개 업로드됨: " +
               ", ".join(f.name for f in grid_files))

# 세션 내 분류 이력 표시
if st.session_state.extra_map:
    total_saved = sum(len(v) for v in st.session_state.extra_map.values())
    st.info(f"💾 이번 세션에서 저장된 거래처 분류 {total_saved}건 — 파일 생성 시 자동 반영됩니다.")

st.divider()
if not grid_files:
    st.info("원장 파일을 업로드하면 분석을 시작할 수 있습니다.")

btn_analyze = st.button(
    "🔍 데이터 분석 시작",
    type="primary",
    use_container_width=True,
    disabled=not bool(grid_files),
)

if btn_analyze and grid_files:
    with st.spinner("원장 파일 읽는 중..."):
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

        existing_raw = None
        base_path = (DEFAULT_BASE_PATH if st.session_state.use_default_base
                     else st.session_state.custom_base_path)

        if base_path and os.path.exists(base_path):
            try:
                existing_raw = pd.read_excel(
                    base_path, sheet_name="Raw_Data",
                    dtype={"계정코드": str, "거래처코드": str}
                )
                _m = sorted(existing_raw["월"].dropna().unique(),
                            key=lambda x: int(x.replace("월","")))
                st.info(f"✅ 기존 Raw_Data {len(existing_raw):,}행 로드 ({', '.join(_m)})")
            except Exception:
                st.info("🔄 Raw_Data 시트 없음 — 월별 시트에서 자동 변환 중...")
                existing_raw = read_legacy_file(base_path)
                if existing_raw.empty:
                    st.warning("⚠️ 기존 파일에서 데이터를 읽지 못했습니다. 신규 생성합니다.")
                    existing_raw = None
                else:
                    st.success(f"✅ 구버전 파일에서 {len(existing_raw):,}행 변환 완료")

        mapped   = apply_mapping(combined, st.session_state.extra_map)
        full_raw = append_to_raw_data(existing_raw, mapped)

        st.session_state.raw_df = full_raw
        st.session_state.step   = 2

    st.rerun()


# ── STEP 2: 신규 거래처 분류 ─────────────────────────────────────
if st.session_state.step >= 2 and st.session_state.raw_df is not None:
    raw_df = apply_mapping(
        st.session_state.raw_df.drop(
            columns=["매출구분","세부구분","표시명","월","비고_매핑"], errors="ignore"
        ),
        st.session_state.extra_map
    )
    st.session_state.raw_df = raw_df

    unmapped = find_unmapped_vendors(raw_df)

    if unmapped:
        st.header("② 신규 거래처 분류", divider="orange")

        new_vendors   = [r for r in unmapped if r[3] == "미분류"]
        check_vendors = [r for r in unmapped if r[3] == "확인필요"]

        if new_vendors:
            st.warning(f"⚠️ 매핑 정보가 없는 **신규 거래처 {len(new_vendors)}개** — 분류해주세요.")
        if check_vendors:
            st.info(f"🔍 적요로 구분이 불분명한 **거래처 {len(check_vendors)}개** — 확인 후 분류해주세요.")

        sub_options = {
            "게임매출": GAME_SUB_ORDER,
            "기타수익": ETC_SUB_ORDER,
            "용역수익": ["경영관리수익"],
        }

        # 전체 펼치기/접기 토글
        if "expand_all" not in st.session_state:
            st.session_state.expand_all = False
        col_exp, _ = st.columns([1, 4])
        with col_exp:
            toggle_label = "📂 전체 접기" if st.session_state.expand_all else "📂 전체 펼치기"
            if st.button(toggle_label, key="toggle_expand"):
                st.session_state.expand_all = not st.session_state.expand_all
                st.rerun()

        changes = {}
        for acct, vendor, label, cur_sub, cur_note in unmapped:
            is_check = (cur_sub == "확인필요")
            icon = "🔍" if is_check else "📌"
            tag  = "적요 확인 필요" if is_check else "신규 거래처"

            # 세션에 이미 저장된 분류가 있으면 기본값으로 활용
            saved = st.session_state.extra_map.get(str(acct), {}).get(vendor)

            # 저장된 항목은 기본 접힘, 전체 펼치기 시 예외
            is_expanded = st.session_state.expand_all or not bool(saved)

            with st.expander(f"{icon} [{label}] {vendor}  —  {tag}", expanded=is_expanded):
                if saved:
                    st.caption(f"💾 이전 분류 저장됨: {saved[0]} / {saved[1]}")
                if is_check and cur_note:
                    st.caption(f"📄 {cur_note}")
                c1, c2, c3 = st.columns([2, 2, 3])
                with c1:
                    default_cat = saved[1] if saved and saved[1] in ["게임매출","기타수익","용역수익"] else label
                    cat = st.selectbox(
                        "매출구분",
                        ["게임매출", "기타수익", "용역수익"],
                        key=f"cat_{acct}_{vendor}",
                        index=["게임매출","기타수익","용역수익"].index(default_cat)
                              if default_cat in ["게임매출","기타수익","용역수익"] else 0
                    )
                with c2:
                    sub = st.selectbox("세부구분", sub_options.get(cat, ["기타"]),
                                       key=f"sub_{acct}_{vendor}")
                with c3:
                    default_display = saved[0] if saved else vendor[:20]
                    default_note    = saved[2] if saved else ""
                    display = st.text_input("표시명", value=default_display,
                                            key=f"disp_{acct}_{vendor}")
                    note    = st.text_input("비고", value=default_note,
                                            key=f"note_{acct}_{vendor}")
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

    months = sorted(
        agg_df["월"].dropna().unique().tolist(),
        key=lambda x: int(x.replace("월",""))
    )

    # ── 월별 합계 미리보기 테이블 ─────────────────────────────────
    st.write(f"**처리 결과:** {len(months)}개월 ({', '.join(months)}), 총 {len(raw_df):,}행")

    if not agg_df.empty:
        with st.expander("📊 월별 합계 미리보기", expanded=False):
            pivot = (
                agg_df.groupby(["매출구분","월"])["대변"]
                .sum()
                .unstack(fill_value=0)
                .reindex(columns=months, fill_value=0)
            )
            row_order = [r for r in ["게임매출","용역수익","기타수익"] if r in pivot.index]
            pivot = pivot.reindex(row_order)
            pivot["합계"] = pivot.sum(axis=1)
            pivot_display = pivot.applymap(lambda x: f"{x:,.0f}")
            st.dataframe(pivot_display, use_container_width=True)

    # 파일명
    today_str = date.today().strftime("%Y%m%d")
    if months:
        start_m = months[0].replace("월","")
        end_m   = months[-1].replace("월","")
        year    = 2026
        if not raw_df.empty and "회계일" in raw_df.columns:
            dates = pd.to_datetime(raw_df["회계일"], errors="coerce").dropna()
            if not dates.empty:
                year = int(dates.dt.year.mode()[0])
        fname = f"별도매출현황_{year}_{start_m}월-{end_m}월_{today_str}.xlsx"
    else:
        fname = f"별도매출현황_{today_str}.xlsx"

    st.write(f"**출력 파일명:** `{fname}`")

    if st.button("📥 Excel 파일 생성", type="primary", use_container_width=True):
        with st.spinner("Excel 파일 생성 중..."):
            out_path = tempfile.mktemp(suffix=".xlsx")
            generate_excel(raw_df, agg_df, out_path)

        with open(out_path, "rb") as f:
            data = f.read()
        os.unlink(out_path)

        st.download_button(
            label=f"⬇️ {fname} 다운로드",
            data=data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.success("✅ 파일 생성 완료!")

    if st.button("🔄 처음부터 다시", use_container_width=True):
        st.session_state.confirm_reset = True

    if st.session_state.get("confirm_reset"):
        st.warning("**초기화 범위를 선택해주세요.**")
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            if st.button("데이터만 초기화\n(분류 이력 유지)", use_container_width=True, key="reset_data"):
                for key in ["raw_df", "step", "confirm_reset", "expand_all"]:
                    st.session_state.pop(key, None)
                st.rerun()
        with col_r2:
            if st.button("전체 초기화\n(분류 이력 포함)", use_container_width=True, key="reset_all"):
                for key in ["extra_map", "raw_df", "step", "confirm_reset", "expand_all"]:
                    st.session_state.pop(key, None)
                st.rerun()
        with col_r3:
            if st.button("취소", use_container_width=True, key="reset_cancel"):
                st.session_state.confirm_reset = False
                st.rerun()


# ── 사이드바 ──────────────────────────────────────────────────────
with st.sidebar:
    st.header("📖 사용 가이드")
    st.markdown("""**기본 사용법**

① **원장 파일 업로드**
- 데이터 경로: FI시스템 > 계정별원장
- 검색 조건: 회계일(당분기), 계정(컨텐츠제공수익·용역수익·기타수익) 선택

⇒ 엑셀 파일 다운로드 후 업로드

② **(별도)매출현황 업로드**

직전 작성 파일 업로드

<span style="color:#E8820C;">※ 기본 파일 해제 후 신규 파일 업로드할 것</span>

③ **데이터 분석 시작** 클릭

④ 신규 거래처 있을 경우 분류 및 비고 작성

⑤ 엑셀 파일 생성 및 다운로드

---

**안내사항**
- Raw Data: 기존 파일에 누적 갱신됨

---

**파일 구조**
- `거래처별 요약` — 월별 주요 거래처 매출 요약
- `N월` — 월별 매출 내역 상세
- `Raw Data` — 원본 데이터
- `검색` — 조건 입력 시 관련 데이터 자동 필터링 (다중 조건·복수 선택 가능)

**파일명 규칙**
- `별도매출현황_연도_시작월-종료월_생성일자.xlsx`
- 예: `별도매출현황_2026_1월-6월_20260723.xlsx`
""", unsafe_allow_html=True)

    st.divider()

    # 계정코드 안내
    st.caption("**원장 파일 계정코드**")
    st.caption("4100100 : 컨텐츠제공수익 (게임매출)")
    st.caption("4100300 : 용역수익")
    st.caption("4100500 : 기타수익")

    st.divider()

    # 현재 기본 파일 상태
    st.caption("**현재 기본 파일**")
    if st.session_state.use_default_base and os.path.exists(DEFAULT_BASE_PATH):
        st.caption("📂 별도매출현황_2026_1월2월3월4월.xlsx")
    elif st.session_state.custom_base_path:
        st.caption("📂 사용자 업로드 파일 사용 중")
    else:
        st.caption("⬜ 기존 파일 없음 (신규 생성 모드)")

    # 세션 내 분류 이력
    if st.session_state.extra_map:
        total_saved = sum(len(v) for v in st.session_state.extra_map.values())
        st.divider()
        st.caption(f"**이번 세션 저장 분류:** {total_saved}건")
        for acct, vendors in st.session_state.extra_map.items():
            for vendor, (display, sub, _) in vendors.items():
                st.caption(f"  · {display} ({sub})")
