# processor.py
# grid 파일 읽기 → 필터링 → 매핑 적용 → 통합 DataFrame 반환

import pandas as pd
from config import VENDOR_MAP, ACCOUNT_LABEL

EXCLUDE_NAMES = {"전일이월", "월계", "누계"}


def load_grid_files(file_paths):
    """
    여러 grid 파일을 읽어 하나의 DataFrame으로 합칩니다.
    각 파일에서 불필요한 행(전일이월, 월계, 누계) 및 거래처 없는 행 제거.
    """
    frames = []
    for path in file_paths:
        df = pd.read_excel(path, dtype={"계정코드": str, "거래처코드": str})
        # 헤더 정리
        df.columns = [str(c).strip() for c in df.columns]
        # 실제 거래 행만 유지
        df = df[
            ~df["계정명"].isin(EXCLUDE_NAMES) &
            df["거래처"].notna() &
            df["거래처"].str.strip().ne("")
        ].copy()
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    return combined


def apply_mapping(df, extra_map=None):
    """
    DataFrame에 매핑 정보(매출구분, 세부구분, 표시명, 월)를 추가합니다.
    extra_map: {계정코드: {거래처명: (표시명, 세부구분, 비고)}} — 웹앱에서 추가된 신규 분류
    """
    merged_map = {}
    for acct, mapping in VENDOR_MAP.items():
        merged_map[acct] = dict(mapping)
    if extra_map:
        for acct, mapping in extra_map.items():
            if acct not in merged_map:
                merged_map[acct] = {}
            merged_map[acct].update(mapping)

    # 계정코드 기준 매출구분
    df["매출구분"] = df["계정코드"].map(ACCOUNT_LABEL).fillna("기타")

    # 거래처 + 계정코드 기준 세부구분 / 표시명 / 비고
    def _map_row(row):
        acct = str(row["계정코드"])
        vendor = str(row["거래처"]).strip()
        acct_map = merged_map.get(acct, {})
        info = acct_map.get(vendor)
        if info:
            return pd.Series({"표시명": info[0], "세부구분": info[1], "비고_매핑": info[2]})
        return pd.Series({"표시명": vendor, "세부구분": "미분류", "비고_매핑": ""})

    mapped = df.apply(_map_row, axis=1)
    df = pd.concat([df, mapped], axis=1)

    # 월 컬럼 (예: "6월")
    df["회계일"] = pd.to_datetime(df["회계일"], errors="coerce")
    df["월"] = df["회계일"].dt.month.astype("Int64").astype(str) + "월"

    # 대변 숫자형 보장
    df["대변"] = pd.to_numeric(df["대변"], errors="coerce").fillna(0)

    return df


def find_unmapped_vendors(df):
    """
    매핑되지 않은 신규 거래처 목록 반환.
    Returns list of (계정코드, 거래처명, 매출구분) tuples.
    """
    unmapped = df[df["세부구분"] == "미분류"][["계정코드", "거래처", "매출구분"]].drop_duplicates()
    return unmapped.values.tolist()


def aggregate_by_month_vendor(df):
    """
    월 × 표시명 × 매출구분 × 세부구분 기준으로 대변 합산.
    Returns: pivot-ready DataFrame
    """
    agg = (
        df.groupby(["월", "매출구분", "세부구분", "표시명", "비고_매핑"], dropna=False)["대변"]
        .sum()
        .reset_index()
    )
    return agg


def append_to_raw_data(existing_df, new_df):
    """
    기존 Raw_Data에 새 데이터를 누적 append.
    동일 (전표번호 + 순번) 중복은 제거.
    """
    if existing_df is None or existing_df.empty:
        return new_df

    combined = pd.concat([existing_df, new_df], ignore_index=True)
    if "전표번호" in combined.columns and "순번" in combined.columns:
        combined = combined.drop_duplicates(subset=["전표번호", "순번"], keep="last")
    return combined.reset_index(drop=True)
