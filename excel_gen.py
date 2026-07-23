# excel_gen.py
# 별도매출현황 Excel 파일 생성
# 시트 순서: 거래처별요약 → N월 (월별) → Raw_Data

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
import pandas as pd
from config import TOP_GAME_VENDORS, GAME_SUB_ORDER, BIZ_SUB_ORDER, ETC_SUB_ORDER

# ── 색상 팔레트 ─────────────────────────────────────────────────
C_HEADER_GAME = "1F4E79"   # 게임매출 헤더 (진파랑)
C_HEADER_ETC  = "375623"   # 기타수익 헤더 (진초록)
C_HEADER_SVC  = "7030A0"   # 용역수익 헤더 (보라)
C_SUB_GAME    = "D6E4F0"   # 게임매출 소계 배경
C_SUB_ETC     = "E2EFDA"   # 기타수익 소계 배경
C_SUB_SVC     = "EAD7F7"   # 용역수익 소계 배경
C_TOTAL       = "FFF2CC"   # 총합계 배경
C_SECTION_BG  = "F2F2F2"   # 섹션 타이틀 배경
C_ROW_ALT     = "FAFAFA"   # 교차 행 배경

NUM_FMT  = '#,##0'         # 숫자 서식
FONT_NAME = "Arial"

thin = Side(style="thin", color="CCCCCC")
med  = Side(style="medium", color="888888")
THIN_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
MED_BORDER  = Border(left=med,  right=med,  top=med,  bottom=med)


def _font(bold=False, size=10, color="000000"):
    return Font(name=FONT_NAME, bold=bold, size=size, color=color)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _apply_header(ws, row, col, text, bg, fg="FFFFFF", bold=True, size=10):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = _font(bold=bold, size=size, color=fg)
    cell.fill = _fill(bg)
    cell.alignment = _align("center")
    cell.border = THIN_BORDER
    return cell

def _apply_num(ws, row, col, value_or_formula, bg=None):
    cell = ws.cell(row=row, column=col, value=value_or_formula)
    cell.number_format = NUM_FMT
    cell.alignment = _align("right")
    cell.border = THIN_BORDER
    if bg:
        cell.fill = _fill(bg)
    return cell

def _apply_label(ws, row, col, text, bg=None, bold=False, indent=0):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = _font(bold=bold)
    cell.alignment = Alignment(horizontal="left", vertical="center",
                               indent=indent, wrap_text=False)
    cell.border = THIN_BORDER
    if bg:
        cell.fill = _fill(bg)
    return cell


# ── Raw_Data 컬럼 위치 (1-based) ─────────────────────────────────
# 원본 30개 컬럼 뒤에 추가 컬럼
RAW_COL_COUNT = 30          # 원본 컬럼 수
COL_CREDIT = 7              # 대변 (G열)
COL_VENDOR = 11             # 거래처 (K열)
# 추가 컬럼
COL_LABEL   = RAW_COL_COUNT + 1   # 매출구분  (AE)
COL_SUB     = RAW_COL_COUNT + 2   # 세부구분  (AF)
COL_DISPLAY = RAW_COL_COUNT + 3   # 표시명    (AG)
COL_MONTH   = RAW_COL_COUNT + 4   # 월        (AH)
COL_NOTE    = RAW_COL_COUNT + 5   # 비고_매핑 (AI)

COL_DISPLAY_LETTER = get_column_letter(COL_DISPLAY)
COL_MONTH_LETTER   = get_column_letter(COL_MONTH)
COL_CREDIT_LETTER  = get_column_letter(COL_CREDIT)
COL_LABEL_LETTER   = get_column_letter(COL_LABEL)
COL_SUB_LETTER     = get_column_letter(COL_SUB)

RAW_SHEET = "Raw_Data"


def _sumifs_formula(month_str, display_name, category=None, sub=None):
    """
    Raw_Data 참조 SUMIFS 수식 생성
    """
    base = (
        f"SUMIFS({RAW_SHEET}!${COL_CREDIT_LETTER}:${COL_CREDIT_LETTER},"
        f"{RAW_SHEET}!${COL_DISPLAY_LETTER}:${COL_DISPLAY_LETTER},\"{display_name}\","
        f"{RAW_SHEET}!${COL_MONTH_LETTER}:${COL_MONTH_LETTER},\"{month_str}\""
    )
    if category:
        base += f",{RAW_SHEET}!${COL_LABEL_LETTER}:${COL_LABEL_LETTER},\"{category}\""
    if sub:
        base += f",{RAW_SHEET}!${COL_SUB_LETTER}:${COL_SUB_LETTER},\"{sub}\""
    return "=" + base + ")"


# ── 월별 시트 ────────────────────────────────────────────────────

def _write_section_header(ws, row, title, bg_color):
    """섹션 타이틀 행 (■ 게임매출 등)"""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    cell = ws.cell(row=row, column=1, value=f"■ {title}")
    cell.font = _font(bold=True, size=10, color="FFFFFF")
    cell.fill = _fill(bg_color)
    cell.alignment = _align("left")
    cell.border = THIN_BORDER
    for c in range(2, 5):
        ws.cell(row=row, column=c).fill = _fill(bg_color)
        ws.cell(row=row, column=c).border = THIN_BORDER
    return row + 1


def _write_col_headers(ws, row):
    headers = ["구분", "거래처명", "금액", "비고"]
    for i, h in enumerate(headers, 1):
        _apply_header(ws, row, i, h, "404040")
    return row + 1


def build_monthly_sheet(ws, month_str, agg_df, year):
    """
    월별 시트 작성
    agg_df: aggregate_by_month_vendor() 결과
    """
    # 타이틀
    ws.merge_cells("A1:C1")
    t = ws.cell(row=1, column=1, value=f"{year}년 {month_str} 매출현황")
    t.font = _font(bold=True, size=13)
    t.alignment = _align("center")
    unit = ws.cell(row=1, column=4, value="(단위: 원)")
    unit.alignment = _align("right")
    unit.font = _font(size=9, color="666666")

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 6   # 공백

    # 열 너비
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 40

    row = 3
    month_data = agg_df[agg_df["월"] == month_str] if not agg_df.empty else pd.DataFrame()

    # ── 게임매출 ─────────────────────────────────────────────────
    row = _write_section_header(ws, row, "게임매출", C_HEADER_GAME)
    row = _write_col_headers(ws, row)
    hdr_row = row - 1

    sub_start_rows = {}  # sub → first data row
    sub_end_rows   = {}  # sub → last data row

    for sub in GAME_SUB_ORDER:
        sub_vendors = (
            month_data[(month_data["매출구분"] == "게임매출") & (month_data["세부구분"] == sub)]
            if not month_data.empty else pd.DataFrame()
        )
        # 매핑에 있는 모든 벤더 + 실제 데이터에 있는 벤더 합집합
        from config import VENDOR_MAP
        known = {v[0] for v in VENDOR_MAP.get("4100100", {}).values() if v[1] == sub}
        actual_names = set(sub_vendors["표시명"].tolist()) if not sub_vendors.empty else set()
        all_names = sorted(known | actual_names)

        if not all_names:
            continue

        sub_start_rows[sub] = row
        for name in all_names:
            formula = _sumifs_formula(month_str, name, "게임매출", sub)
            note_series = (
                sub_vendors[sub_vendors["표시명"] == name]["비고_매핑"]
                if not sub_vendors.empty else pd.Series(dtype=str)
            )
            note = note_series.iloc[0] if not note_series.empty else ""
            bg = C_ROW_ALT if row % 2 == 0 else None
            _apply_label(ws, row, 1, sub, bg, indent=1)
            _apply_label(ws, row, 2, name, bg)
            _apply_num(ws, row, 3, formula, bg)
            _apply_label(ws, row, 4, note, bg)
            row += 1
        sub_end_rows[sub] = row - 1

    # 소계 행
    subtotal_rows = {}
    for sub in GAME_SUB_ORDER:
        if sub not in sub_start_rows:
            continue
        s, e = sub_start_rows[sub], sub_end_rows[sub]
        c_ref = get_column_letter(3)
        formula = f"=SUM({c_ref}{s}:{c_ref}{e})"
        _apply_label(ws, row, 1, f"{sub} 소계", C_SUB_GAME, bold=True)
        _apply_label(ws, row, 2, "", C_SUB_GAME)
        _apply_num(ws, row, 3, formula, C_SUB_GAME)
        ws.cell(row=row, column=3).font = _font(bold=True)
        _apply_label(ws, row, 4, "", C_SUB_GAME)
        subtotal_rows[sub] = row
        row += 1

    # 게임매출 합계
    game_total_row = row
    if subtotal_rows:
        refs = "+".join(f"C{r}" for r in subtotal_rows.values())
        game_formula = f"={refs}"
    else:
        game_formula = "=0"
    for c in range(1, 5):
        ws.cell(row=row, column=c).fill = _fill(C_SUB_GAME)
        ws.cell(row=row, column=c).border = THIN_BORDER
    _apply_label(ws, row, 1, "게임매출 합계", C_SUB_GAME, bold=True)
    _apply_label(ws, row, 2, "", C_SUB_GAME)
    _apply_num(ws, row, 3, game_formula, C_SUB_GAME)
    ws.cell(row=row, column=3).font = _font(bold=True)
    row += 2  # 공백 1행

    # ── 기타수익 ─────────────────────────────────────────────────
    row = _write_section_header(ws, row, "기타수익", C_HEADER_ETC)
    row = _write_col_headers(ws, row)

    etc_sub_rows = {}
    for sub in ETC_SUB_ORDER:
        sub_vendors = (
            month_data[(month_data["매출구분"] == "기타수익") & (month_data["세부구분"] == sub)]
            if not month_data.empty else pd.DataFrame()
        )
        from config import VENDOR_MAP as VM
        known = {v[0] for v in VM.get("4100500", {}).values() if v[1] == sub}
        actual_names = set(sub_vendors["표시명"].tolist()) if not sub_vendors.empty else set()
        all_names = sorted(known | actual_names)

        if not all_names:
            continue

        s_row = row
        for name in all_names:
            formula = _sumifs_formula(month_str, name, "기타수익", sub)
            note_series = (
                sub_vendors[sub_vendors["표시명"] == name]["비고_매핑"]
                if not sub_vendors.empty else pd.Series(dtype=str)
            )
            note = note_series.iloc[0] if not note_series.empty else ""
            bg = C_ROW_ALT if row % 2 == 0 else None
            _apply_label(ws, row, 1, sub, bg, indent=1)
            _apply_label(ws, row, 2, name, bg)
            _apply_num(ws, row, 3, formula, bg)
            _apply_label(ws, row, 4, note, bg)
            row += 1
        etc_sub_rows[sub] = (s_row, row - 1)

    # 기타수익 소계
    etc_total_row = row
    if etc_sub_rows:
        etc_formula = "=SUM(" + ",".join(f"C{s}:C{e}" for s, e in etc_sub_rows.values()) + ")"
    else:
        etc_formula = "=0"
    for c in range(1, 5):
        ws.cell(row=row, column=c).fill = _fill(C_SUB_ETC)
        ws.cell(row=row, column=c).border = THIN_BORDER
    _apply_label(ws, row, 1, "기타수익 합계", C_SUB_ETC, bold=True)
    _apply_label(ws, row, 2, "", C_SUB_ETC)
    _apply_num(ws, row, 3, etc_formula, C_SUB_ETC)
    ws.cell(row=row, column=3).font = _font(bold=True)
    row += 2

    # ── 용역수익 ─────────────────────────────────────────────────
    row = _write_section_header(ws, row, "용역수익", C_HEADER_SVC)
    row = _write_col_headers(ws, row)

    svc_vendors = (
        month_data[month_data["매출구분"] == "용역수익"]
        if not month_data.empty else pd.DataFrame()
    )
    from config import VENDOR_MAP as VM2
    known_svc = {v[0] for v in VM2.get("4100300", {}).values()}
    actual_svc = set(svc_vendors["표시명"].tolist()) if not svc_vendors.empty else set()
    all_svc = sorted(known_svc | actual_svc)

    svc_start = row
    for name in all_svc:
        formula = _sumifs_formula(month_str, name, "용역수익")
        note_series = (
            svc_vendors[svc_vendors["표시명"] == name]["비고_매핑"]
            if not svc_vendors.empty else pd.Series(dtype=str)
        )
        note = note_series.iloc[0] if not note_series.empty else ""
        bg = C_ROW_ALT if row % 2 == 0 else None
        _apply_label(ws, row, 1, "경영관리수익", bg, indent=1)
        _apply_label(ws, row, 2, name, bg)
        _apply_num(ws, row, 3, formula, bg)
        _apply_label(ws, row, 4, note, bg)
        row += 1
    svc_end = row - 1

    # 용역수익 합계
    svc_total_row = row
    svc_formula = f"=SUM(C{svc_start}:C{svc_end})" if all_svc else "=0"
    for c in range(1, 5):
        ws.cell(row=row, column=c).fill = _fill(C_SUB_SVC)
        ws.cell(row=row, column=c).border = THIN_BORDER
    _apply_label(ws, row, 1, "용역수익 합계", C_SUB_SVC, bold=True)
    _apply_label(ws, row, 2, "", C_SUB_SVC)
    _apply_num(ws, row, 3, svc_formula, C_SUB_SVC)
    ws.cell(row=row, column=3).font = _font(bold=True)
    row += 2

    # ── 총합계 ───────────────────────────────────────────────────
    for c in range(1, 5):
        ws.cell(row=row, column=c).fill = _fill(C_TOTAL)
        ws.cell(row=row, column=c).border = MED_BORDER
    _apply_label(ws, row, 1, "총  합  계", C_TOTAL, bold=True)
    ws.cell(row=row, column=1).font = _font(bold=True, size=11)
    ws.cell(row=row, column=1).alignment = _align("center")
    _apply_label(ws, row, 2, "", C_TOTAL)
    total_formula = f"=C{game_total_row}+C{etc_total_row}+C{svc_total_row}"
    _apply_num(ws, row, 3, total_formula, C_TOTAL)
    ws.cell(row=row, column=3).font = _font(bold=True, size=11)

    # 창 고정 (헤더 고정)
    ws.freeze_panes = "A3"


# ── 거래처별 요약 시트 ─────────────────────────────────────────────

def build_summary_sheet(ws, months, agg_df, year):
    """
    거래처별 요약 시트: 게임매출 상위 거래처 × 월별
    months: ["1월", "2월", ...] 오름차순
    """
    # 타이틀
    col_count = len(months) + 2
    t = ws.cell(row=1, column=1, value=f"{year}년 컴투스 매출현황 (게임매출 주요 거래처)")
    t.font = _font(bold=True, size=13)
    t.alignment = _align("center")

    ws.cell(row=2, column=col_count, value="(단위: 백만원)").alignment = _align("right")

    # 헤더 행
    row = 3
    _apply_header(ws, row, 1, "분류", C_HEADER_GAME)
    _apply_header(ws, row, 2, "거래처", C_HEADER_GAME)
    for i, m in enumerate(months, 3):
        _apply_header(ws, row, i, m, C_HEADER_GAME)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 20
    for i in range(3, col_count + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13

    row = 4
    # 게임매출 상위 7개 + 기타 + 합계
    vendor_rows = {}

    for vendor in TOP_GAME_VENDORS:
        bg = C_ROW_ALT if row % 2 == 0 else None
        _apply_label(ws, row, 1, "게임매출", bg)
        _apply_label(ws, row, 2, vendor, bg)
        for i, m in enumerate(months, 3):
            # 백만원 단위 SUMIFS ÷ 1,000,000
            f = f"=IFERROR({_sumifs_formula(m, vendor, '게임매출')[1:]}/1000000,0)"
            _apply_num(ws, row, i, f, bg)
        vendor_rows[vendor] = row
        row += 1

    # 기타 (합계 - 상위7)
    _apply_label(ws, row, 1, "게임매출", C_ROW_ALT)
    _apply_label(ws, row, 2, "기타", C_ROW_ALT)
    other_row = row
    for i, m in enumerate(months, 3):
        # 게임매출 전체 - 상위7 합
        total_f = f"SUMIFS({RAW_SHEET}!${COL_CREDIT_LETTER}:${COL_CREDIT_LETTER},{RAW_SHEET}!${COL_MONTH_LETTER}:${COL_MONTH_LETTER},\"{m}\",{RAW_SHEET}!${COL_LABEL_LETTER}:${COL_LABEL_LETTER},\"게임매출\")"
        known_refs = "+".join(
            f"SUMIFS({RAW_SHEET}!${COL_CREDIT_LETTER}:${COL_CREDIT_LETTER},"
            f"{RAW_SHEET}!${COL_DISPLAY_LETTER}:${COL_DISPLAY_LETTER},\"{v}\","
            f"{RAW_SHEET}!${COL_MONTH_LETTER}:${COL_MONTH_LETTER},\"{m}\","
            f"{RAW_SHEET}!${COL_LABEL_LETTER}:${COL_LABEL_LETTER},\"게임매출\")"
            for v in TOP_GAME_VENDORS
        )
        f = f"=IFERROR(({total_f}-({known_refs}))/1000000,0)"
        _apply_num(ws, row, i, f, C_ROW_ALT)
    row += 1

    # 합계 행
    for c in range(1, col_count + 1):
        ws.cell(row=row, column=c).fill = _fill(C_SUB_GAME)
        ws.cell(row=row, column=c).border = THIN_BORDER
    ws.cell(row=row, column=1, value="게임매출").font = _font(bold=True)
    ws.cell(row=row, column=1).fill = _fill(C_SUB_GAME)
    ws.cell(row=row, column=1).alignment = _align("left")
    ws.cell(row=row, column=2, value="합계").font = _font(bold=True)
    ws.cell(row=row, column=2).fill = _fill(C_SUB_GAME)
    ws.cell(row=row, column=2).alignment = _align("left")
    game_total_summary_row = row
    for i, m in enumerate(months, 3):
        f = (f"=IFERROR(SUMIFS({RAW_SHEET}!${COL_CREDIT_LETTER}:${COL_CREDIT_LETTER},"
             f"{RAW_SHEET}!${COL_MONTH_LETTER}:${COL_MONTH_LETTER},\"{m}\","
             f"{RAW_SHEET}!${COL_LABEL_LETTER}:${COL_LABEL_LETTER},\"게임매출\")/1000000,0)")
        _apply_num(ws, row, i, f, C_SUB_GAME)
        ws.cell(row=row, column=i).font = _font(bold=True)
    row += 1

    # 기타수익 행
    for c in range(1, col_count + 1):
        ws.cell(row=row, column=c).border = THIN_BORDER
    _apply_label(ws, row, 1, "기타수익")
    _apply_label(ws, row, 2, "합계")
    for i, m in enumerate(months, 3):
        f = (f"=IFERROR(SUMIFS({RAW_SHEET}!${COL_CREDIT_LETTER}:${COL_CREDIT_LETTER},"
             f"{RAW_SHEET}!${COL_MONTH_LETTER}:${COL_MONTH_LETTER},\"{m}\","
             f"{RAW_SHEET}!${COL_LABEL_LETTER}:${COL_LABEL_LETTER},\"기타수익\")/1000000,0)")
        _apply_num(ws, row, i, f)
    etc_summary_row = row
    row += 1

    # 용역수익 행
    for c in range(1, col_count + 1):
        ws.cell(row=row, column=c).border = THIN_BORDER
    _apply_label(ws, row, 1, "용역수익")
    _apply_label(ws, row, 2, "합계")
    for i, m in enumerate(months, 3):
        f = (f"=IFERROR(SUMIFS({RAW_SHEET}!${COL_CREDIT_LETTER}:${COL_CREDIT_LETTER},"
             f"{RAW_SHEET}!${COL_MONTH_LETTER}:${COL_MONTH_LETTER},\"{m}\","
             f"{RAW_SHEET}!${COL_LABEL_LETTER}:${COL_LABEL_LETTER},\"용역수익\")/1000000,0)")
        _apply_num(ws, row, i, f)
    svc_summary_row = row
    row += 1

    # 총합계 행
    for c in range(1, col_count + 1):
        ws.cell(row=row, column=c).fill = _fill(C_TOTAL)
        ws.cell(row=row, column=c).border = MED_BORDER
    ws.cell(row=row, column=1, value="합계").font = _font(bold=True, size=11)
    ws.cell(row=row, column=1).fill = _fill(C_TOTAL)
    ws.cell(row=row, column=1).alignment = _align("center")
    ws.cell(row=row, column=2, value="총계").font = _font(bold=True)
    ws.cell(row=row, column=2).fill = _fill(C_TOTAL)
    ws.cell(row=row, column=2).alignment = _align("left")
    for i in range(3, col_count + 1):
        f = f"={get_column_letter(i)}{game_total_summary_row}+{get_column_letter(i)}{etc_summary_row}+{get_column_letter(i)}{svc_summary_row}"
        _apply_num(ws, row, i, f, C_TOTAL)
        ws.cell(row=row, column=i).font = _font(bold=True, size=11)

    ws.freeze_panes = "C4"


# ── Raw_Data 시트 ─────────────────────────────────────────────────

def build_raw_data_sheet(ws, raw_df):
    """
    Raw_Data 시트: 원본 데이터 + 매핑 컬럼
    """
    if raw_df is None or raw_df.empty:
        ws.cell(row=1, column=1, value="데이터 없음")
        return

    # 컬럼 순서 정리
    original_cols = [c for c in raw_df.columns
                     if c not in ("매출구분", "세부구분", "표시명", "월", "비고_매핑")]
    extra_cols = ["매출구분", "세부구분", "표시명", "월", "비고_매핑"]
    all_cols = original_cols + [c for c in extra_cols if c in raw_df.columns]
    raw_df = raw_df[all_cols]

    # 헤더
    for c_idx, col in enumerate(all_cols, 1):
        cell = ws.cell(row=1, column=c_idx, value=col)
        is_extra = col in extra_cols
        bg = "1F4E79" if is_extra else "404040"
        cell.font = _font(bold=True, color="FFFFFF")
        cell.fill = _fill(bg)
        cell.alignment = _align("center")
        cell.border = THIN_BORDER

    # 데이터 행
    for r_idx, row_data in enumerate(raw_df.itertuples(index=False), 2):
        for c_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = THIN_BORDER
            cell.font = _font()
            if isinstance(val, (int, float)) and c_idx == COL_CREDIT:
                cell.number_format = NUM_FMT
                cell.alignment = _align("right")

    # 열 너비 자동 조정 (간단 버전)
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["D"].width = 14  # 회계일
    ws.column_dimensions["I"].width = 40  # 적요
    ws.column_dimensions["K"].width = 36  # 거래처

    # 추가 컬럼 강조
    for c_idx in range(RAW_COL_COUNT + 1, len(all_cols) + 1):
        ltr = get_column_letter(c_idx)
        ws.column_dimensions[ltr].width = 18

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(all_cols))}1"


# ── 메인 생성 함수 ────────────────────────────────────────────────

def generate_excel(raw_df, agg_df, output_path):
    """
    최종 Excel 파일 생성.
    raw_df: Raw_Data (매핑 컬럼 포함)
    agg_df: aggregate_by_month_vendor() 결과
    output_path: 저장 경로
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # 기본 Sheet 제거

    # 연도 추출
    months_in_data = sorted(
        agg_df["월"].dropna().unique().tolist(),
        key=lambda x: int(x.replace("월", ""))
    ) if not agg_df.empty else []

    year = 2026
    if not raw_df.empty and "회계일" in raw_df.columns:
        dates = pd.to_datetime(raw_df["회계일"], errors="coerce").dropna()
        if not dates.empty:
            year = dates.dt.year.mode()[0]

    # 1. 거래처별 요약 (첫 번째 시트)
    ws_sum = wb.create_sheet("거래처별요약")
    build_summary_sheet(ws_sum, months_in_data, agg_df, year)

    # 2. 월별 시트
    for m in months_in_data:
        ws_m = wb.create_sheet(m)
        build_monthly_sheet(ws_m, m, agg_df, year)

    # 3. Raw_Data (마지막)
    ws_raw = wb.create_sheet(RAW_SHEET)
    build_raw_data_sheet(ws_raw, raw_df)

    wb.save(output_path)
    return output_path
