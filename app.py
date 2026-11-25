import calendar
import copy
import os
import warnings
import pandas as pd
import numpy as np
import json
import glob
import re
import openpyxl
from datetime import datetime, timedelta

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def find_latest_inventory_file(prefix):
    """폴더에서 prefix로 시작하는 .xlsx 파일 중 날짜(YYYYMMDD)가 가장 큰 것 반환"""
    files = glob.glob(f"{prefix}*.xlsx")
    if not files:
        return None
    def extract_date(fname):
        m = re.search(r'(\d{8})', fname)
        return int(m.group(1)) if m else 0
    return max(files, key=extract_date)


def find_latest_sales_file(keywords):
    """국가별 판매 파일(*SCM_재고관리*.xlsx) 중 국가 키워드가 포함된 파일을 모아
    수정시각(최종 저장 시간)이 가장 최신인 것을 반환.
    같은 이름으로 재다운로드하면 '(1)' 복사본이 생기므로 파일명이 아니라 mtime 기준."""
    candidates = [
        f for f in glob.glob("*SCM_재고관리*.xlsx")
        if not os.path.basename(f).startswith("~$")
        and any(kw in f for kw in keywords)
    ]
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)


def load_robust_excel(file_path, target_keyword='품목코드', sheet_name=0):
    if not os.path.exists(file_path):
        return None
    try:
        df_raw = pd.read_excel(file_path, header=None, sheet_name=sheet_name)
        header_idx = None
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.contains(target_keyword).any():
                header_idx = idx
                break
        if header_idx is not None:
            raw_cols = [str(x).strip() for x in df_raw.iloc[header_idx]]
            columns = []
            seen = {}
            for c in raw_cols:
                if c in seen:
                    seen[c] += 1
                    columns.append(f"{c}_{seen[c]}")
                else:
                    seen[c] = 0
                    columns.append(c)
            df_refined = df_raw.iloc[header_idx + 1:].copy()
            df_refined.columns = columns
            return df_refined.reset_index(drop=True)
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df.columns = [str(c).strip() for c in df.columns]
            return df
    except Exception as e:
        print(f"❌ '{file_path}' 시트 '{sheet_name}' 로드 중 에러: {e}")
        return None


def parse_int(val):
    if pd.notna(val) and str(val).strip() != '-':
        try:
            return int(float(str(val).strip().replace(',', '')))
        except:
            return 0
    return 0


def find_sku_col(df):
    """재고 파일들의 SKU 컬럼명 탐색. 2026-07-10 FBA 파일부터 'Seller SKU'가 'SKU'로 변경됨
    (FNSKU 컬럼이 추가되며 명칭 정리된 것으로 보임) — 일반/FBT/FBA 세 파일 모두 같은 변경
    리스크가 있어 공통 헬퍼로 둠."""
    for c in ('Seller SKU', 'SKU'):
        if c in df.columns:
            return c
    raise KeyError("'Seller SKU' 또는 'SKU' 컬럼을 찾을 수 없습니다.")


def assign_region(warehouse_name):
    wh_str = str(warehouse_name)
    if '미국' in wh_str:
        return 'us'
    elif '영국' in wh_str:
        return 'uk'
    # 스페인·이탈리아·프랑스 FBT 창고도 EU (2026-11-19 — 이전엔 '기타'로 빠져 EU 재고에서 누락)
    elif any(keyword in wh_str for keyword in ['독일', '네덜란드', '스페인', '이탈리아', '프랑스', '유럽', 'EU']):
        return 'eu'
    return '기타'


def read_first_matching_sheet(path, required_cols):
    """여러 시트 중 required_cols를 모두 포함한 첫 시트를 DataFrame으로 반환.
    재고 파일에 피벗 시트가 첫 번째로 저장된 경우 대응 (2026-07-09 일반_전체재고 파일에서 발생 —
    첫 시트가 '행 레이블' 피벗이라 기존 pd.read_excel(첫 시트 고정)이 조용히 실패)."""
    xls = pd.ExcelFile(path)
    for sn in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sn)
            df.columns = [str(c).strip() for c in df.columns]
            if all(c in df.columns for c in required_cols):
                return df
        except Exception:
            continue
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_integrated_inventory(file_gen, file_fbt, file_fba):
    """일반/FBT/FBA 통합 재고 파일 3종 읽어 국가별 창고재고 합계 산출"""
    all_rows = []

    def expand_multicode_rows(df_input):
        expanded = []
        for _, row in df_input.iterrows():
            code_str = str(row['품목코드']).strip()
            if ',' in code_str:
                sub_codes = [c.strip() for c in code_str.split(',') if c.strip()]
                for sc in sub_codes:
                    new_row = row.copy()
                    new_row['품목코드'] = sc
                    expanded.append(new_row)
            else:
                expanded.append(row)
        return pd.DataFrame(expanded).reset_index(drop=True)

    if file_gen and os.path.exists(file_gen):
        try:
            df_gen = read_first_matching_sheet(file_gen, ['창고', '가용 재고 (개)'])
            tmp = pd.DataFrame({
                '창고': df_gen['창고'],
                '품목코드': df_gen[find_sku_col(df_gen)].astype(str).str.strip().str.upper(),
                '가용재고': df_gen['가용 재고 (개)'].apply(parse_int),
                '상품명': df_gen['상품명'] if '상품명' in df_gen.columns else ''
            })
            tmp = expand_multicode_rows(tmp)
            all_rows.append(tmp)
            print(f"   > {file_gen} 로드 완료 ({len(tmp)}행)")
        except Exception as e:
            print(f"   > ⚠️ {file_gen} 처리 중 에러: {e}")

    if file_fbt and os.path.exists(file_fbt):
        try:
            df_fbt = read_first_matching_sheet(file_fbt, ['가용 재고 (개)'])
            tmp = pd.DataFrame({
                '창고': df_fbt['창고'],
                '품목코드': df_fbt[find_sku_col(df_fbt)].astype(str).str.strip().str.upper(),
                '가용재고': df_fbt['가용 재고 (개)'].apply(parse_int),
                '상품명': df_fbt['상품명'] if '상품명' in df_fbt.columns else ''
            })
            tmp = expand_multicode_rows(tmp)
            all_rows.append(tmp)
            print(f"   > {file_fbt} 로드 완료 ({len(tmp)}행)")
        except Exception as e:
            print(f"   > ⚠️ {file_fbt} 처리 중 에러: {e}")

    if file_fba and os.path.exists(file_fba):
        try:
            df_fba = read_first_matching_sheet(file_fba, ['가용 재고 (개)'])
            warehouse_col = '창고' if '창고' in df_fba.columns else (
                '창고재고그룹' if '창고재고그룹' in df_fba.columns else None
            )
            if warehouse_col is None:
                raise KeyError("'창고' 또는 '창고재고그룹' 컬럼을 찾을 수 없습니다.")
            tmp = pd.DataFrame({
                '창고': df_fba[warehouse_col],
                '품목코드': df_fba[find_sku_col(df_fba)].astype(str).str.strip().str.upper(),
                '가용재고': df_fba['가용 재고 (개)'].apply(parse_int),
                '상품명': df_fba['상품명'] if '상품명' in df_fba.columns else ''
            })
            tmp = expand_multicode_rows(tmp)
            all_rows.append(tmp)
            print(f"   > {file_fba} 로드 완료 ({len(tmp)}행)")
        except Exception as e:
            print(f"   > ⚠️ {file_fba} 처리 중 에러: {e}")

    if not all_rows:
        return pd.DataFrame(columns=['품목코드', 'us_stock', 'uk_stock', 'eu_stock'])

    df_total = pd.concat(all_rows, ignore_index=True)
    df_total = df_total.dropna(subset=['품목코드'])
    df_total['품목코드'] = df_total['품목코드'].astype(str).str.strip()
    global _LAST_INV_ROWS
    _LAST_INV_ROWS = df_total.copy()   # 리뉴얼(v2) 추가 권역 재고용 (US/UK/EU 필터 전 전체 창고)
    df_total['region'] = df_total['창고'].apply(assign_region)
    df_total = df_total[df_total['region'].isin(['us', 'uk', 'eu'])]

    df_grouped = df_total.groupby(['품목코드', 'region'])['가용재고'].sum().reset_index()
    df_pivot = df_grouped.pivot(index='품목코드', columns='region', values='가용재고').fillna(0).astype(int).reset_index()

    for region in ['us', 'uk', 'eu']:
        if region not in df_pivot.columns:
            df_pivot[region] = 0

    df_pivot = df_pivot.rename(columns={'us': 'us_stock', 'uk': 'uk_stock', 'eu': 'eu_stock'})
    return df_pivot[['품목코드', 'us_stock', 'uk_stock', 'eu_stock']]


def parse_kr_available(file_gen):
    """일반_전체재고 파일에서 국내(콜로세움) 창고의 SKU별 가용재고(F열) 추출.
    부킹 관리에서 '한국에서 보낼 수 있는 재고'의 베이스로 사용 (권역 현지 재고와 별개).
    """
    if not file_gen or not os.path.exists(file_gen):
        return {}
    try:
        df = read_first_matching_sheet(file_gen, ['창고', '가용 재고 (개)'])
        sku_col = find_sku_col(df)
    except Exception as e:
        print(f"   > ⚠️ 한국(콜로세움) 가용재고 파싱 실패: {e}")
        return {}
    if '창고' not in df.columns or '가용 재고 (개)' not in df.columns:
        print("   > ⚠️ 한국(콜로세움) 가용재고: 필요한 컬럼(창고/가용 재고)을 찾지 못함")
        return {}

    kr_map = {}
    for _, row in df.iterrows():
        if '콜로세움' not in str(row['창고']):
            continue
        code_str = str(row[sku_col]).strip().upper()
        if not code_str or code_str == 'NAN':
            continue
        qty = parse_int(row['가용 재고 (개)'])
        # 멀티코드(콤마 구분)는 기존 재고 로드와 동일하게 각 코드에 동일 수량 부여
        codes = [c.strip() for c in code_str.split(',')] if ',' in code_str else [code_str]
        for c in codes:
            if c:
                kr_map[c] = kr_map.get(c, 0) + qty
    return kr_map


def read_and_summarize_sales(file_path, region_prefix):
    if not os.path.exists(file_path):
        print(f"⚠️ '{file_path}' 파일이 없어 {region_prefix} 일판매 0 처리")
        return pd.DataFrame(columns=['품목코드', f'{region_prefix}_sales'])
    try:
        df_raw = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"❌ '{file_path}' 로드 에러: {e}")
        return pd.DataFrame(columns=['품목코드', f'{region_prefix}_sales'])

    sales_col_idx = None
    for r_idx in range(min(20, len(df_raw) - 1)):
        row_values = df_raw.iloc[r_idx].tolist()
        next_row_values = df_raw.iloc[r_idx + 1].tolist()
        for c_idx in range(min(len(row_values), len(next_row_values))):
            val_current = row_values[c_idx]
            val_next = next_row_values[c_idx]
            if pd.notna(val_current) and '당월 일평균 판매량' in str(val_current):
                if pd.notna(val_next) and '소계' in str(val_next):
                    sales_col_idx = c_idx
                    break
        if sales_col_idx is not None:
            break

    df = load_robust_excel(file_path, '품목코드')
    if df is None or '품목코드' not in df.columns:
        return pd.DataFrame(columns=['품목코드', f'{region_prefix}_sales'])

    df = df.dropna(subset=['품목코드'])
    df['품목코드'] = df['품목코드'].astype(str).str.strip()
    df = df[df['품목코드'] != '품목코드']

    explicit_sales_col = next((c for c in df.columns if '일평균' in str(c) and '판매' in str(c)), None)
    if explicit_sales_col:
        df[f'{region_prefix}_sales'] = df[explicit_sales_col].apply(parse_int)
    elif sales_col_idx is not None and sales_col_idx < len(df.columns):
        try:
            df[f'{region_prefix}_sales'] = df.iloc[:, sales_col_idx].apply(parse_int)
        except Exception:
            df[f'{region_prefix}_sales'] = 0
    else:
        df[f'{region_prefix}_sales'] = 0

    return df[['품목코드', f'{region_prefix}_sales']]


def _find_daily_layout(df_raw):
    """판매 시트에서 일자별 컬럼(datetime 헤더 셀)과 품목코드 컬럼 위치를 탐색.
    파일마다 헤더 행 위치가 다르므로(US 3행째, UK/EU 2행째) 앞쪽 행들을 스캔해서 찾는다.
    날짜 datetime 셀이 있는 그 컬럼이 해당 일자 블록의 '소계' 컬럼이다.
    """
    date_cols = []
    for r in range(min(6, len(df_raw))):
        row = df_raw.iloc[r].tolist()
        cols = [(i, v) for i, v in enumerate(row) if isinstance(v, datetime)]
        if len(cols) >= 5:  # 날짜 컬럼이 여러 개 있는 행이 일자별 헤더 (단일 datetime 셀 오탐 방지)
            date_cols = [(i, v.date()) for i, v in cols]
            break
    code_row, code_col = None, None
    for r in range(min(8, len(df_raw))):
        for i, v in enumerate(df_raw.iloc[r].tolist()):
            if str(v).strip() == '품목코드':
                code_row, code_col = r, i
                break
        if code_col is not None:
            break
    return date_cols, code_row, code_col


def parse_daily_sales(file_path):
    """판매 파일의 최신 2개 월 시트('26년 7월' 형식)에서 SKU별 일자별 판매량(소계) 추출.
    반환: {품목코드: {date: qty}}. 월 경계를 넘는 28일 롤링 평균 계산을 위해 두 달치를 합친다.
    """
    if not os.path.exists(file_path):
        return {}
    try:
        xls = pd.ExcelFile(file_path)
    except Exception:
        return {}

    monthly_sheets = []
    for s in xls.sheet_names:
        m = re.match(r'(\d{2})년\s*(\d{1,2})월$', str(s).strip())
        if m:
            monthly_sheets.append((2000 + int(m.group(1)), int(m.group(2)), s))
    monthly_sheets.sort(reverse=True)

    result = {}
    for year, month, sheet in monthly_sheets[:2]:
        try:
            df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
        except Exception:
            continue
        date_cols, code_row, code_col = _find_daily_layout(df_raw)
        if not date_cols or code_col is None:
            print(f"   > ⚠️ '{sheet}' 일자별 컬럼 탐색 실패 — 이 시트는 건너뜀")
            continue
        for r in range(code_row + 1, len(df_raw)):
            code = df_raw.iat[r, code_col]
            if pd.isna(code):
                continue
            code = str(code).strip().upper()
            if not code or code == '품목코드':
                continue
            entry = result.setdefault(code, {})
            for ci, d in date_cols:
                if ci >= df_raw.shape[1]:
                    continue
                qty = parse_int(df_raw.iat[r, ci])
                if qty:
                    entry[d] = entry.get(d, 0) + qty
    return result


def summarize_sales_28d(daily_map, region_prefix, today=None):
    """일자별 판매에서 '어제까지' 기준 다중 윈도우(7/14/28일) 평균 중 최댓값을 일판매로 산출.
    - 28일 단일 평균은 판매가 막 시작된 SKU(결품 후 재개, 신규 수요)를 과소평가해서
      항공 긴급을 늦게 감지하는 문제가 있었음 (2026-07-15). max 방식이라 보수적(재고 여유 방향).
    - 어떤 윈도우가 채택됐는지 {r}_sales_window(7/14/28)로 함께 반환 — 셀 툴팁 표시용.
    급증 감지: 최근 3일 평균이 max(10, 28일 평균×2) 이상이면 spike 플래그.
    반환 DataFrame: 품목코드 / {r}_sales / {r}_sales_window / {r}_last7(list) / {r}_spike(bool)
    """
    today = today or datetime.now().date()
    start = today - timedelta(days=28)
    rows = []
    for code, dates in daily_map.items():
        window = [dates.get(start + timedelta(days=i), 0) for i in range(28)]
        avg28 = sum(window) / 28
        avg14 = sum(window[-14:]) / 14
        avg7 = sum(window[-7:]) / 7
        # 동률이면 긴 윈도우 우선 (수치가 안정적이고, 툴팁에 '급증 반영'으로 오해되지 않도록)
        best_avg, best_window = avg28, 28
        if avg14 > best_avg:
            best_avg, best_window = avg14, 14
        if avg7 > best_avg:
            best_avg, best_window = avg7, 7
        recent3 = sum(window[-3:]) / 3
        spike = recent3 >= max(10, 2 * avg28)
        rows.append({
            '품목코드': code,
            f'{region_prefix}_sales': int(round(best_avg)),
            f'{region_prefix}_sales_window': best_window,
            f'{region_prefix}_last7': window[-7:],
            f'{region_prefix}_spike': bool(spike),
        })
    return pd.DataFrame(rows)


def build_group_sales(merged, region_daily, group_col='통합 제품명'):
    """통합 제품명 그룹(tab2)의 일판매를 '멤버 코드의 일자별 판매를 먼저 합산한 뒤' 윈도우 평균으로 산출.

    코드별 {r}_sales를 단순 합산하면 안 되는 이유: 2026-07-15부터 코드마다 채택 윈도우가
    달라질 수 있어(max(7·14·28일)), 구코드의 28일 평균 + 신코드의 7일 평균처럼 서로 다른
    기간의 수요가 더해져 과대 계상된다. 코드 전환 그룹에서 실제 발생
    (GLMX090→GLMX178: 3,094 + 2,151 = 5,245로 표시됐지만 실제 그룹 수요는 3,632).
    일자별로 먼저 합산하면 전환기에도 그룹의 실제 수요 곡선 하나만 남는다.

    반환: {region: DataFrame(group_col, {r}_sales, {r}_sales_window, {r}_last7, {r}_spike)}
          일자별 파싱이 실패한 권역(폴백 경로)은 결과에서 빠지며, 호출부는 기존 합산값을 유지한다.
    """
    out = {}
    if group_col not in merged.columns:
        return out
    # (코드, 그룹) 쌍 단위 중복 제거 — 마스터의 유령 중복 코드('GLMX183' vs 'GLMX183\n')로
    # merge 시 같은 코드 행이 여러 벌 생기면 한 그룹 안에서 같은 SKU 판매가 중복 합산된다.
    # 코드만으로 dedup하면 반대로 다른 그룹에 붙은 같은 코드가 통째로 누락되므로 쌍으로 dedup한다.
    pairs = merged[['품목코드', group_col]].dropna(subset=['품목코드']).copy()
    pairs['_code_key'] = pairs['품목코드'].astype(str).str.strip().str.upper()
    pairs = pairs.drop_duplicates(subset=['_code_key', group_col])
    for r, dmap in region_daily.items():
        if not dmap:
            continue
        group_daily = {}
        for code, gname in zip(pairs['_code_key'], pairs[group_col]):
            d = dmap.get(code)
            if not d:
                continue
            entry = group_daily.setdefault(str(gname), {})
            for dt, q in d.items():
                entry[dt] = entry.get(dt, 0) + q
        if group_daily:
            out[r] = summarize_sales_28d(group_daily, r).rename(columns={'품목코드': group_col})
    return out


def build_group_fcst(master_df, fcst_map):
    """마스터의 코드→통합 제품명 매핑으로 월별 판매예정을 그룹 단위 합산.

    merged 기반 집계(agg_fcst)를 쓰면 안 되는 이유: merged는 재고·판매·이동중 데이터가 있는
    코드로만 만들어지고 마스터는 거기에 left join되므로, **재고가 아직 없는 신코드는 행 자체가
    없어** 그 코드의 포캐스팅 물량이 통째로 누락된다. 제품 리뉴얼로 코드가 바뀌면
    (GLMX124 → GLMX124_V02) 신코드에 잡힌 향후 계획이 사라져 그룹 포캐스팅이 0으로 떨어짐
    (2026-07-22 발견). 마스터를 순회하면 재고 유무와 무관하게 전부 집계된다.

    반환: {통합 제품명: {"YYYY-MM": {"us","uk","eu"}}}
    """
    out = {}
    if master_df is None or '품목코드' not in master_df.columns or '통합 제품명' not in master_df.columns:
        return out
    for _, r in master_df.dropna(subset=['품목코드']).iterrows():
        gname = str(r.get('통합 제품명', '')).strip()
        if not gname or gname.lower() == 'nan':
            continue
        for part in str(r['품목코드']).split(','):
            code = part.strip().upper()
            months = fcst_map.get(code) if code else None
            if not months:
                continue
            g = out.setdefault(gname, {})
            for mk, regs in months.items():
                o = g.setdefault(mk, {'us': 0, 'uk': 0, 'eu': 0, 'jp': 0})
                for rg in ('us', 'uk', 'eu', 'jp'):
                    o[rg] += regs.get(rg, 0)
    return out


def read_sales_rolling(file_path, region_prefix):
    """일자별 파싱 기반 28일 롤링 평균으로 일판매 산출. 파싱 실패 시 기존 '당월 일평균' 방식으로 폴백.
    반환: (DataFrame, daily_map) — daily_map은 모달의 일별 추이 차트용.
    """
    daily_map = parse_daily_sales(file_path)
    if not daily_map:
        print(f"   > ⚠️ {region_prefix.upper()}: 일자별 파싱 실패 → 기존 '당월 일평균' 컬럼으로 폴백")
        return read_and_summarize_sales(file_path, region_prefix), {}

    df_new = summarize_sales_28d(daily_map, region_prefix)

    # 검증 로그: 기존 당월 일평균 방식과 비교 (둘 다 판매가 있는 SKU 기준)
    df_old = read_and_summarize_sales(file_path, region_prefix).rename(
        columns={f'{region_prefix}_sales': '_old_sales'})
    cmp = pd.merge(df_new[['품목코드', f'{region_prefix}_sales']], df_old, on='품목코드', how='inner')
    both = cmp[(cmp[f'{region_prefix}_sales'] > 0) & (cmp['_old_sales'] > 0)]
    if len(both) > 0:
        ratio = (both[f'{region_prefix}_sales'] / both['_old_sales']).mean()
        print(f"   > {region_prefix.upper()}: 28일 롤링 {len(df_new)}개 SKU 산출 "
              f"(기존 당월 일평균 대비 평균 비율 {ratio:.2f}, 비교 가능 {len(both)}개)")
    else:
        print(f"   > {region_prefix.upper()}: 28일 롤링 {len(df_new)}개 SKU 산출 (기존 방식과 비교 가능한 SKU 없음)")

    return df_new, daily_map


def read_transit_data(file_path):
    df = load_robust_excel(file_path, '품목코드')
    if df is None or '품목코드' not in df.columns:
        print(f"⚠️ 이동중 재고 파일('{file_path}') 없음")
        return pd.DataFrame(columns=['품목코드', 'us_air', 'us_sea', 'uk_air', 'uk_sea', 'eu_air', 'eu_sea'])
    df = df.dropna(subset=['품목코드'])
    df['품목코드'] = df['품목코드'].astype(str).str.strip()
    mapping = {
        'us_air': 'US항공', 'us_sea': 'US해상',
        'uk_air': 'UK항공', 'uk_sea': 'UK해상',
        'eu_air': 'EU항공', 'eu_sea': 'EU해상'
    }
    res_df = pd.DataFrame({'품목코드': df['품목코드']})
    for key, col_name in mapping.items():
        if col_name in df.columns:
            res_df[key] = df[col_name].apply(parse_int)
        else:
            alt_col = next((c for c in df.columns if key.split('_')[0].upper() in str(c) and ('항공' if 'air' in key else '해상') in str(c)), None)
            res_df[key] = df[alt_col].apply(parse_int) if alt_col else 0
    # 품목코드는 다른 원본(재고·판매·수불부)처럼 대문자로 통일 — 'GLMX226_v01'처럼 소문자로 적힌 행이
    # 마스터의 'GLMX226_V01'과 따로 노는 문제 방지 (2026-11-23). 대문자화 후 같은 코드는 합산.
    res_df['품목코드'] = res_df['품목코드'].str.upper()
    return res_df.groupby('품목코드', as_index=False).sum(numeric_only=True)


def parse_transit_schedule(file_path, year=2026):
    """ETA 파일에서 SKU별 입고 일정 [{region, mode, date, qty}] 추출"""
    if not os.path.exists(file_path):
        print(f"   > ⚠️ ETA 파일 없음: {file_path}")
        return {}

    schedule_map = {}

    target_sheets = [
        ('us', 'air', 'US항공'),
        ('us', 'sea', 'US해상'),
        ('uk', 'air', 'UK항공'),
        ('uk', 'sea', 'UK해상'),
        ('eu', 'air', 'EU항공'),
        ('eu', 'sea', 'EU해상'),
    ]

    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"   > ⚠️ ETA 파일 열기 실패: {e}")
        return {}

    for region, mode, sheet_keyword in target_sheets:
        matched_sheet = next((s for s in xls.sheet_names if s.startswith(sheet_keyword)), None)
        if matched_sheet is None:
            print(f"   > ⚠️ '{sheet_keyword}' 시트를 찾을 수 없습니다.")
            continue

        try:
            df_raw = pd.read_excel(file_path, sheet_name=matched_sheet, header=None)
        except Exception as e:
            print(f"   > ⚠️ '{matched_sheet}' 읽기 실패: {e}")
            continue

        header_row = None
        for r in range(min(10, len(df_raw))):
            row_vals = df_raw.iloc[r].tolist()
            if any('품목코드' in str(v) for v in row_vals):
                header_row = r
                break
        if header_row is None:
            print(f"   > ⚠️ '{matched_sheet}'에서 헤더를 찾지 못함")
            continue

        header = df_raw.iloc[header_row].tolist()
        code_col_idx = None
        date_cols = []
        tbd_cols = []   # ETA 미정 차수 컬럼 (예: '해상4차(TBD)', '레오스 특송(TBD)') — 날짜 없이 수량만 있음
        broken_cols = []  # 차수 열처럼 보이는데 날짜/TBD를 못 읽은 열 (예: '항공5차(7월15일)', '해상3차(7/32)')

        date_pattern = re.compile(r'\((\d{1,2})/(\d{1,2})\)')
        batch_pattern = re.compile(r'\d+\s*차|특송')  # 차수 열 판별 ('항공5차', '해상22차', '레오스 특송')

        for i, h in enumerate(header):
            if h is None or (isinstance(h, float) and pd.isna(h)):
                continue
            h_str = str(h)
            if '품목코드' in h_str and code_col_idx is None:
                code_col_idx = i
            m = date_pattern.search(h_str)
            if m:
                month = int(m.group(1))
                day = int(m.group(2))
                try:
                    eta_date = datetime(year, month, day).date()
                    date_cols.append((i, eta_date, h_str.strip()))
                except ValueError:
                    broken_cols.append(h_str.strip())  # 존재하지 않는 날짜 (예: 7/32)
            elif 'TBD' in h_str.upper():
                tbd_cols.append((i, h_str.strip()))
            elif batch_pattern.search(h_str):
                broken_cols.append(h_str.strip())

        if code_col_idx is None:
            print(f"   > ⚠️ '{matched_sheet}': 품목코드 컬럼 못 찾음")
            continue

        if broken_cols:
            # 형식이 깨진 차수 열은 입고 일정에서 빠지므로 반드시 경고 (정상 차수 열은 아래에서 그대로 파싱)
            print(f"   > ⚠️ '{matched_sheet}': 차수 열 {len(broken_cols)}개의 ETA 형식을 읽지 못해 입고 일정에서 제외됨 "
                  f"→ {', '.join(broken_cols[:5])}{' 외' if len(broken_cols) > 5 else ''} "
                  f"(형식: '항공5차(7/15)' 또는 '해상4차(TBD)')")

        if date_cols or tbd_cols:
            for r in range(header_row + 1, len(df_raw)):
                code = df_raw.iat[r, code_col_idx]
                if pd.isna(code):
                    continue
                sku = str(code).strip().upper()
                if not sku or sku == 'NAN' or sku == '품목코드':
                    continue

                for col_idx, eta_date, label in date_cols:
                    if col_idx >= df_raw.shape[1]:
                        continue
                    val = df_raw.iat[r, col_idx]
                    qty = parse_int(val)
                    if qty <= 0:
                        continue
                    schedule_map.setdefault(sku, []).append({
                        'region': region,
                        'mode': mode,
                        'date': eta_date.isoformat(),
                        'qty': qty,
                        'label': label
                    })

                for col_idx, label in tbd_cols:
                    if col_idx >= df_raw.shape[1]:
                        continue
                    qty = parse_int(df_raw.iat[r, col_idx])
                    if qty <= 0:
                        continue
                    schedule_map.setdefault(sku, []).append({
                        'region': region,
                        'mode': mode,
                        'date': None,
                        'qty': qty,
                        'label': label
                    })
            print(f"   > '{matched_sheet}' 파싱 완료 (도착일 컬럼 {len(date_cols)}개 + ETA미정 컬럼 {len(tbd_cols)}개)")
            continue

        # 대체 스키마: 여러 날짜 컬럼 대신 '입고 예정일' + '수량' 단일 컬럼 (예: EU항공)
        eta_date_col_idx = next((i for i, h in enumerate(header) if h is not None and '입고 예정일' in str(h)), None)
        qty_col_idx = next((i for i, h in enumerate(header) if h is not None and str(h).strip() == '수량'), None)

        if eta_date_col_idx is None or qty_col_idx is None:
            if not broken_cols:
                # 차수 열 자체가 없음 = 현재 이 경로로 이동 중인 재고가 없음 (정상 상태, 경고 아님)
                print(f"   > ℹ️ '{matched_sheet}': 이동중 차수 없음")
            continue

        row_count = 0
        for r in range(header_row + 1, len(df_raw)):
            code = df_raw.iat[r, code_col_idx]
            if pd.isna(code):
                continue
            sku = str(code).strip().upper()
            if not sku or sku == 'NAN' or sku == '품목코드':
                continue

            raw_date = df_raw.iat[r, eta_date_col_idx] if eta_date_col_idx < df_raw.shape[1] else None
            if pd.isna(raw_date):
                continue
            try:
                eta_date = pd.to_datetime(raw_date).date()
            except Exception:
                continue

            qty = parse_int(df_raw.iat[r, qty_col_idx]) if qty_col_idx < df_raw.shape[1] else 0
            if qty <= 0:
                continue

            schedule_map.setdefault(sku, []).append({
                'region': region,
                'mode': mode,
                'date': eta_date.isoformat(),
                'qty': qty
            })
            row_count += 1

        print(f"   > '{matched_sheet}' 파싱 완료 (입고 예정일+수량 방식, {row_count}건)")

    print(f"   > ETA 일정 추출 완료: 총 {len(schedule_map)}개 SKU의 입고 기록")
    return schedule_map


def parse_domestic_inbound(file_path):
    """
    국내 메인창고 입고 스케줄 파싱
    파일: 2026_SCM_영업채널별_판매_포캐스팅 / 탭: '입고 스케줄_상세'
    반환: { 품목코드: { 'total': 미입총계, 'items': [{month, week, qty}, ...] } }
    """
    if not os.path.exists(file_path):
        print(f"   > ⚠️ 입고 스케줄 파일 없음: {file_path}")
        return {}

    try:
        df_raw = pd.read_excel(file_path, sheet_name='입고 스케줄_상세', header=None)
    except Exception as e:
        print(f"   > ⚠️ 입고 스케줄 시트 읽기 실패: {e}")
        return {}

    if len(df_raw) < 5:
        print("   > ⚠️ 입고 스케줄 시트 데이터 부족")
        return {}

    row_month = df_raw.iloc[2].tolist()
    row_week  = df_raw.iloc[3].tolist()

    code_col_idx = None
    total_col_idx = None
    week_cols = []  # [(col_idx, '6월', '22~28...')]

    current_month = None
    for i in range(len(row_week)):
        m = row_month[i] if i < len(row_month) else None
        w = row_week[i] if i < len(row_week) else None

        if m is not None and m != '':
            m_str = str(m).strip()
            if m_str == '미입 총계':
                total_col_idx = i
                continue
            if '월' in m_str:
                current_month = m_str

        if w == '품목코드':
            code_col_idx = i
            continue

        if w is not None and isinstance(w, str):
            w_clean = w.replace('\n', ' ').strip()
            if w_clean == '소계':
                continue
            if current_month and ('~' in w_clean or '영업일' in w_clean):
                week_cols.append((i, current_month, w_clean))

    if code_col_idx is None or total_col_idx is None or not week_cols:
        print(f"   > ⚠️ 입고 스케줄 헤더 파싱 실패 (code={code_col_idx}, total={total_col_idx}, weeks={len(week_cols)})")
        return {}

    result = {}
    for r in range(4, len(df_raw)):
        row = df_raw.iloc[r].tolist()
        if code_col_idx >= len(row):
            continue
        code = row[code_col_idx]
        if pd.isna(code):
            continue
        code = str(code).strip().upper()
        if not code or code in ('NAN', '품목코드'):
            continue

        total_val = parse_int(row[total_col_idx]) if total_col_idx < len(row) else 0
        items = []
        for col_idx, month, week in week_cols:
            if col_idx >= len(row):
                continue
            qty = parse_int(row[col_idx])
            if qty > 0:
                items.append({'month': month, 'week': week, 'qty': qty})

        if total_val > 0 or items:
            result[code] = {'total': total_val, 'items': items}

    print(f"   > 입고 스케줄 파싱 완료: {len(result)}개 SKU")
    return result


def parse_product_cost(file_path, sheet_name='입고 스케줄_상세'):
    """입고 스케줄_상세 시트에서 품목코드별 원가(E열 '원가(vat+)') 추출 → 재고 금액 계산에 사용"""
    if not os.path.exists(file_path):
        return {}
    try:
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        print(f"   > ⚠️ 원가 파싱 실패: {e}")
        return {}

    header_row = None
    for r in range(min(6, len(df_raw))):
        if any('품목코드' in str(v) for v in df_raw.iloc[r].tolist()):
            header_row = r
            break
    if header_row is None:
        print("   > ⚠️ 원가 시트에서 헤더(품목코드)를 찾지 못함")
        return {}

    header = df_raw.iloc[header_row].tolist()
    code_col = next((i for i, h in enumerate(header) if str(h).strip() == '품목코드'), None)
    cost_col = next((i for i, h in enumerate(header) if '원가' in str(h)), None)
    if code_col is None or cost_col is None:
        print(f"   > ⚠️ 원가 컬럼 못 찾음 (code={code_col}, cost={cost_col})")
        return {}

    cost_map = {}
    for r in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[r].tolist()
        if code_col >= len(row):
            continue
        code = row[code_col]
        if pd.isna(code):
            continue
        code = str(code).strip().upper()
        if not code or code == '품목코드':
            continue
        cost = parse_int(row[cost_col]) if cost_col < len(row) else 0
        if cost > 0 and code not in cost_map:
            cost_map[code] = cost
    return cost_map


def parse_manufacturer(file_path, sheet_name='입고 스케줄_상세'):
    """입고 스케줄_상세 시트에서 품목코드별 제조사(A열 '제조사') 추출 → Slack 부족알림 제조사 그룹핑에 사용."""
    if not os.path.exists(file_path):
        return {}
    try:
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    except Exception:
        return {}
    header_row = None
    for r in range(min(6, len(df_raw))):
        if any('품목코드' in str(v) for v in df_raw.iloc[r].tolist()):
            header_row = r
            break
    if header_row is None:
        return {}
    header = df_raw.iloc[header_row].tolist()
    code_col = next((i for i, h in enumerate(header) if str(h).strip() == '품목코드'), None)
    mfr_col = next((i for i, h in enumerate(header) if str(h).strip() == '제조사'), None)
    if code_col is None or mfr_col is None:
        return {}
    mfr_map = {}
    for r in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[r].tolist()
        if code_col >= len(row):
            continue
        code = row[code_col]
        if pd.isna(code):
            continue
        code = str(code).strip().upper()
        mfr = str(row[mfr_col]).strip() if mfr_col < len(row) and pd.notna(row[mfr_col]) else ''
        if code and mfr and code not in mfr_map:
            mfr_map[code] = mfr
    return mfr_map


def lookup_manufacturer(code, mfr_map):
    """정확일치 우선, 없으면 접미사(-N, _V0N) 제거한 base 코드로 폴백."""
    c = str(code).strip().upper()
    if c in mfr_map:
        return mfr_map[c]
    base = re.sub(r'(_V0\d+|-\d+)$', '', c)
    return mfr_map.get(base, '')


def parse_product_status(file_path, sheet_name='입고 스케줄_상세'):
    """입고 스케줄_상세 시트에서 품목코드별 PLC 상태(운영/단종) 추출.
    제품명 셀에 취소선(strikethrough) 서식이 적용된 행은 단종으로 간주 (파일 작성자가 수기로 표시하는 규칙).
    """
    if not os.path.exists(file_path):
        return {}
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        print(f"   > ⚠️ PLC 상태 파싱 실패: {e}")
        return {}
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]

    header_row, code_col, name_col = None, None, None
    for r in range(1, min(7, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() == '품목코드':
                header_row, code_col = r, c
            if v is not None and str(v).strip() == '제품명':
                name_col = c
        if header_row is not None:
            break
    if header_row is None or code_col is None or name_col is None:
        print("   > ⚠️ PLC 상태: 품목코드/제품명 헤더를 찾지 못함")
        return {}

    status_map = {}
    for r in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(row=r, column=code_col).value
        if code is None:
            continue
        code = str(code).strip().upper()
        if not code or code == '품목코드':
            continue
        font = ws.cell(row=r, column=name_col).font
        status_map[code] = '단종' if (font and font.strike) else '운영'
    return status_map


def parse_forecast_sales(file_path, base_year=2026):
    """포캐스팅 파일의 'N월_판매예정' 탭들에서 SKU별 US/UK/EU 월별 판매예정 수량 파싱.

    시트 구조: '품목코드'가 있는 행=헤더('품목코드'/'수량'/'매출원가'), 그 바로 윗행=채널명, 헤더 다음 행부터 데이터.
    (2026-11 이전엔 헤더가 3행 고정이었으나, 맨 위에 메모 행('블프' 등)이 추가되며 4행으로 밀림 → 앞 8행에서 탐색)
    채널명이 '미국*'이면 US, '영국*'이면 UK, 'EU*'이면 EU로 합산 (MX/일본/글로벌 등 나머지 채널 제외).
    '마케팅(미국)' 같은 부서 소계 열은 '미국'으로 시작하지 않아 자동 제외 → 이중 합산 없음.
    시트명 'NN년 N월_판매예정'은 20NN년. 연도 없는 'N월_판매예정'은 오늘 기준 -5~+6개월 범위로 연도를 추정
    (예: 2026-11 기준 '11월'·'12월'=2026, '1월'~'5월'=2027). base_year는 추정 불가 시에만 쓰는 예비값.

    반환: ({품목코드(upper): {"YYYY-MM": {"us":n, "uk":n, "eu":n}}}, [월 키 정렬 리스트])
    """
    today = datetime.now()

    def infer_year(month):
        d = (month - today.month) % 12
        if d > 6:
            d -= 12  # 7개월 넘게 앞이면 지난 달로 간주 (예: 1월에 보는 '11월' 시트 = 작년 11월)
        return today.year + (today.month - 1 + d) // 12
    result = {}
    months_found = set()
    if not os.path.exists(file_path):
        print(f"   > ⚠️ 포캐스팅 파일 없음: {file_path}")
        return result, []
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    pat = re.compile(r'^(?:(\d{2})년\s*)?(\d{1,2})월_판매예정$')
    for sheet in wb.sheetnames:
        m = pat.match(sheet.strip())
        if not m:
            continue
        year = 2000 + int(m.group(1)) if m.group(1) else (infer_year(int(m.group(2))) or base_year)
        month_key = f"{year}-{int(m.group(2)):02d}"
        ws = wb[sheet]
        rows = ws.iter_rows(values_only=True)
        # 헤더 행('품목코드' 포함)을 앞 8행에서 탐색, 채널명은 그 바로 윗행
        head = []
        r2 = r3 = None
        for _ in range(8):
            try:
                cur = next(rows)
            except StopIteration:
                break
            if cur and any(str(v).strip() == '품목코드' for v in cur if v is not None) and head:
                r2, r3 = head[-1], cur
                break
            head.append(cur)
        if r3 is None:
            print(f"   > ⚠️ '{sheet}' 탭에서 '품목코드' 헤더 행을 찾지 못해 건너뜀")
            continue
        code_idx = next(i for i, v in enumerate(r3) if v is not None and str(v).strip() == '품목코드')
        region_cols = {'us': [], 'uk': [], 'eu': [], 'jp': []}
        for i, name in enumerate(r2):
            if not name or i >= len(r3) or str(r3[i]).strip() != '수량':
                continue
            nm = str(name).strip()
            if nm.startswith('미국'):
                region_cols['us'].append(i)
            elif nm.startswith('영국'):
                region_cols['uk'].append(i)
            elif nm.startswith('EU'):
                region_cols['eu'].append(i)
            elif nm.startswith('일본'):          # 일본 라쿠텐/큐텐/아마존 (마케팅(일본) 소계 열은 제외됨)
                region_cols['jp'].append(i)
        if not any(region_cols.values()):
            print(f"   > ⚠️ '{sheet}' 탭에서 미국/영국/EU 수량 컬럼을 찾지 못해 건너뜀")
            continue
        months_found.add(month_key)
        for row in rows:
            if not row or code_idx >= len(row) or row[code_idx] is None:
                continue
            code = str(row[code_idx]).strip().upper()
            if not code:
                continue
            entry = result.setdefault(code, {}).setdefault(month_key, {'us': 0, 'uk': 0, 'eu': 0, 'jp': 0})
            for reg, cols in region_cols.items():
                s = 0.0
                for c in cols:
                    v = row[c] if c < len(row) else None
                    try:
                        s += float(v or 0)
                    except (TypeError, ValueError):
                        pass
                entry[reg] += int(round(s))
    return result, sorted(months_found)


def build_pallet_map(path='마스터 계정.xlsx', sheet_name='패킹리스트'):
    """마스터 계정.xlsx의 '패킹리스트' 탭에서 품목코드별 팔렛트당 개입수 추출.
    헤더는 2행 고정, A열=품목코드, N열=아웃박스 입수량, V열=1PLT 카톤박스 총 적재수량.
    팔렛트당 개입수 = N × V.
    """
    if not os.path.exists(path):
        return {}
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=1)
    except Exception as e:
        print(f"   > ⚠️ 팔렛트 환산표 파싱 실패: {e}")
        return {}

    code_col = df.columns[0]
    box_col = next((c for c in df.columns if '아웃박스' in str(c) and '입수' in str(c)), None)
    plt_col = next((c for c in df.columns if '1PLT' in str(c) and '적재' in str(c)), None)
    if box_col is None or plt_col is None:
        print(f"   > ⚠️ 팔렛트 환산표 컬럼 못 찾음 (box={box_col}, plt={plt_col})")
        return {}

    pallet_map = {}
    for _, row in df.iterrows():
        code = row[code_col]
        if pd.isna(code):
            continue
        code = str(code).strip().upper()
        if not code:
            continue
        box_qty = parse_int(row[box_col])
        plt_qty = parse_int(row[plt_col])
        units = box_qty * plt_qty
        if units > 0:
            pallet_map[code] = units
    return pallet_map


def build_b2b_map(prefix='B2B할당_'):
    """B2B 사전할당 소형 엑셀(품목코드 + B2B물량 총량, 파일명에 날짜 포함) 자동 탐색·파싱.
    파일이 아직 없는 주에는 빈 딕셔너리를 반환해 정상 동작을 유지한다.
    """
    path = find_latest_inventory_file(prefix)
    if path is None:
        return {}, None
    df = load_robust_excel(path, '품목코드')
    if df is None or '품목코드' not in df.columns:
        print(f"   > ⚠️ B2B 사전할당 파일({path})에서 '품목코드' 컬럼을 찾지 못함")
        return {}, path

    qty_col = next((c for c in df.columns if 'B2B' in str(c) and ('물량' in str(c) or '수량' in str(c))), None)
    if qty_col is None:
        print(f"   > ⚠️ B2B 사전할당 파일({path})에서 물량 컬럼을 찾지 못함")
        return {}, path

    b2b_map = {}
    for _, row in df.dropna(subset=['품목코드']).iterrows():
        code = str(row['품목코드']).strip().upper()
        if not code:
            continue
        qty = parse_int(row[qty_col])
        if qty > 0:
            b2b_map[code] = b2b_map.get(code, 0) + qty
    return b2b_map, path


def build_master_lookup(master_df):
    name_dict = {}
    if master_df is None or '품목코드' not in master_df.columns or '제품명' not in master_df.columns:
        return name_dict
    for _, row in master_df.dropna(subset=['품목코드']).iterrows():
        raw_code = str(row['품목코드']).strip()
        name = str(row['제품명']).strip() if pd.notna(row.get('제품명')) else ""
        if not name or not raw_code:
            continue
        if ',' in raw_code:
            continue
        if name_dict.get(raw_code, "") == "":
            name_dict[raw_code] = name
    return name_dict


def lookup_product_name(code, name_dict):
    code_str = str(code).strip()
    if ',' in code_str:
        return ""
    return name_dict.get(code_str, "")


# 신제품(버전) 코드 판정: 코드 끝이 _V01, _V02 ... 형태 (예: GLMX238_V01)
NEW_PRODUCT_RE = re.compile(r'_V0\d+$', re.IGNORECASE)

def is_new_product_code(code):
    """마스터에 신제품으로 등록되는 _V0N 접미사 코드인지"""
    return bool(NEW_PRODUCT_RE.search(str(code).strip()))


def apply_sales_plc_override(df):
    """PLC 자동판정(취소선 기반) 보정: 지역 합계 판매량이 100개 이상이면 소스 파일 표시와 무관하게 '운영'으로 간주.
    실제로 잘 팔리고 있는데도 원본 파일의 취소선 표시(오기입 등) 때문에 '단종'으로 잘못 뜨는 사례를 보정하기 위함.
    """
    total_sales = df['us_sales'] + df['uk_sales'] + df['eu_sales']
    df.loc[(total_sales >= 100) & (df['plc_status'] != '운영'), 'plc_status'] = '운영'


def build_excluded_codes(master_df):
    """마스터 계정.xlsx의 '제외' 컬럼에 표시된 품목코드 집합 반환 (단종/미운영 등 리포트 제외 대상)"""
    excluded = set()
    if master_df is None or '제외' not in master_df.columns or '품목코드' not in master_df.columns:
        return excluded
    for _, row in master_df.dropna(subset=['품목코드']).iterrows():
        flag = str(row.get('제외', '')).strip().upper()
        if flag in ('Y', 'TRUE', '1', '제외', 'O'):
            raw_code = str(row['품목코드']).strip()
            if ',' in raw_code:
                for c in raw_code.split(','):
                    if c.strip():
                        excluded.add(c.strip().upper())
            else:
                excluded.add(raw_code.upper())
    return excluded


# 스냅샷 델타 비교에 쓰는 필드 (전체 필드를 임베드하면 파일이 비대해지므로 화면 델타에 쓰는 것만)
SNAPSHOT_DELTA_FIELDS = ['us_total', 'us_days', 'us_sales',
                         'uk_total', 'uk_days', 'uk_sales',
                         'eu_total', 'eu_days', 'eu_sales']


def update_snapshots(json_tab1, json_tab2, snap_dir="snapshots", keep_days=62, embed_days=30):
    """계산 완료된 결과(rawData/groupedData)를 날짜별로 동결 저장하고,
    최근 keep_days일치에서 델타 비교용 필드만 추려 history 객체로 반환한다.

    - full_YYYY-MM-DD.json: 그날의 전체 계산 결과 동결본. 같은 날 재실행 시 그날 최종 상태로 갱신.
      (다른 날짜는 절대 덮어쓰지 않음 — 과거 시점 박제가 목적.)
    - 판매가 나중에 소급 재집계돼도 과거 동결본은 안 바뀜 → '그날 대시보드가 뭘 위험이라 판정했나'가 보존됨.
    """
    os.makedirs(snap_dir, exist_ok=True)
    today = datetime.now().date()
    today_key = today.isoformat()

    # 1) 오늘 동결본 저장 (전체 필드)
    full_path = os.path.join(snap_dir, f"full_{today_key}.json")
    with open(full_path, 'w', encoding='utf-8') as f:
        json.dump({"date": today_key, "generated_at": datetime.now().isoformat(),
                   "tab1": json_tab1, "tab2": json_tab2}, f, ensure_ascii=False)

    # 2) 보관 기간 지난 동결본 삭제
    cutoff = today - timedelta(days=keep_days)
    for p in glob.glob(os.path.join(snap_dir, "full_*.json")):
        m = re.search(r'full_(\d{4}-\d{2}-\d{2})\.json$', p.replace('\\', '/'))
        if m:
            try:
                if datetime.strptime(m.group(1), "%Y-%m-%d").date() < cutoff:
                    os.remove(p)
            except ValueError:
                pass

    # 3) 남은 동결본에서 델타 필드만 추려 history 구성 (dashboard.html에 임베드용)
    hist = {"dates": [], "tab1": {}, "tab2": {}}
    embed_cutoff = (today - timedelta(days=embed_days)).isoformat()   # 파일은 keep_days 보관, 화면 델타는 최근 embed_days만
    for p in sorted(glob.glob(os.path.join(snap_dir, "full_*.json"))):
        m = re.search(r'full_(\d{4}-\d{2}-\d{2})\.json$', p.replace('\\', '/'))
        if not m:
            continue
        dkey = m.group(1)
        if dkey < embed_cutoff:
            continue
        try:
            with open(p, encoding='utf-8') as f:
                snap = json.load(f)
        except Exception:
            continue
        hist["dates"].append(dkey)
        for tab_key in ("tab1", "tab2"):
            hist[tab_key][dkey] = {
                str(r.get("code")): {fld: r.get(fld) for fld in SNAPSHOT_DELTA_FIELDS if fld in r}
                for r in snap.get(tab_key, [])
            }
    print(f"   > 스냅샷: {today_key} 동결 저장, 비교 가능 이력 {len(hist['dates'])}일치 ({hist['dates'][0] if hist['dates'] else '-'} ~ {hist['dates'][-1] if hist['dates'] else '-'})")
    return hist


def build_heavy_codes(master_df):
    """마스터 계정.xlsx의 '중량물' 컬럼에 표시된 품목코드 집합 반환.
    중량물은 항공 운임이 매우 비싸 안전재고 일수 +30일, 항공 경고 임계선 +30일을 적용한다 (2026-07-15).
    컬럼이 없으면 빈 집합 — 기능이 조용히 꺼진 상태로 기존과 동일하게 동작."""
    heavy = set()
    if master_df is None or '중량물' not in master_df.columns or '품목코드' not in master_df.columns:
        return heavy
    for _, row in master_df.dropna(subset=['품목코드']).iterrows():
        flag = str(row.get('중량물', '')).strip().upper()
        if flag in ('Y', 'TRUE', '1', '중량', '중량물', 'O'):
            raw_code = str(row['품목코드']).strip()
            for c in raw_code.split(','):
                if c.strip():
                    heavy.add(c.strip().upper())
    return heavy


HEAVY_SAFETY_EXTRA_DAYS = 30  # 중량물 안전재고 가산 일수 (template.html의 항공 임계선 +30과 세트)
US_SAFETY_DAYS = 120  # 미국 안전재고 일수 (2026-11-25 90→120일 변경, template_v2 DAYS_RULE.us.danger와 세트)


# ─────────────────────────────────────────────────────────────
# 리뉴얼(v2) 대시보드 전용: 층별화 등급(SA/A/B/C) + 제품 카테고리 (2026-11-19)
#   - 계산 결과는 JSON 행에 필드로만 추가됨 → 기존 template.html은 이 필드를 쓰지 않으므로 화면 영향 없음
#   - template_v2.html이 있으면 리뉴얼이 dashboard.html(정식), 기존 화면은 dashboard_legacy.html(로컬 비교용)
# ─────────────────────────────────────────────────────────────
GRADE_BANDS = [(40, 'SA'), (70, 'A'), (90, 'B'), (100, 'C')]  # 누적 비중 상한(%) → 등급


def prev_full_months(n=3, today=None):
    """오늘 기준 직전 n개 '마감된' 달 [(year, month), ...] (최근 달이 앞). 11/19 → 10월, 9월, 8월"""
    today = today or datetime.now().date()
    y, m = today.year, today.month
    out = []
    for _ in range(n):
        m -= 1
        if m == 0:
            y, m = y - 1, 12
        out.append((y, m))
    return out


def parse_monthly_sales_totals(file_path, months):
    """판매 파일의 지정 월 시트('26년 10월' 형식)에서 일자별 소계를 합산해 코드별 월간 판매(출고)량 반환.
    반환: ({품목코드: qty}, [읽은 시트명], [못 찾은 (y,m)])
    일자별 컬럼 탐색은 대시보드 일판매와 같은 _find_daily_layout()을 재사용 (같은 원본·같은 숫자)."""
    totals, used, missing = {}, [], []
    if not file_path or not os.path.exists(file_path):
        return totals, used, list(months)
    try:
        xls = pd.ExcelFile(file_path)
    except Exception:
        return totals, used, list(months)
    sheet_by_month = {}
    for s in xls.sheet_names:
        mm = re.match(r'(\d{2})년\s*(\d{1,2})월$', str(s).strip())
        if mm:
            sheet_by_month[(2000 + int(mm.group(1)), int(mm.group(2)))] = s
    for ym in months:
        sheet = sheet_by_month.get(ym)
        if not sheet:
            missing.append(ym)
            continue
        try:
            df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
        except Exception:
            missing.append(ym)
            continue
        date_cols, code_row, code_col = _find_daily_layout(df_raw)
        if not date_cols or code_col is None:
            print(f"   > ⚠️ '{sheet}' 일자별 컬럼 탐색 실패 — 층별화 집계에서 제외")
            missing.append(ym)
            continue
        # 시트에 다른 달 날짜 컬럼이 섞여 있어도 해당 월 날짜만 합산
        date_cols = [(ci, d) for ci, d in date_cols if (d.year, d.month) == ym]
        used.append(sheet)
        for r in range(code_row + 1, len(df_raw)):
            code = df_raw.iat[r, code_col]
            if pd.isna(code):
                continue
            code = str(code).strip().upper()
            if not code or code == '품목코드':
                continue
            for ci, _d in date_cols:
                if ci < df_raw.shape[1]:
                    q = parse_int(df_raw.iat[r, ci])
                    if q:
                        totals[code] = totals.get(code, 0) + q
    return totals, used, missing


def assign_grades(sales_by_key):
    """판매량 내림차순 누적 비중으로 SA/A/B/C 등급 부여 (파레토/ABC 분석).
    누적 비중(자기 포함) ≤40% SA, ≤70% A, ≤90% B, 나머지 C.
    층별화 엑셀(어크로스 비 층별화_261110.xlsx)의 IF(누적<=40%,"SA",...) 수식과 동일 — 예외 규칙 없음.
    판매 0인 품목은 비중 0% → C."""
    total = sum(q for q in sales_by_key.values() if q > 0)
    result = {}
    cum = 0.0
    for i, (key, q) in enumerate(sorted(sales_by_key.items(), key=lambda kv: -kv[1])):
        share = (q / total * 100) if (total > 0 and q > 0) else 0.0
        cum_after = cum + share
        if q <= 0:
            grade = 'C'
        else:
            grade = next((g for limit, g in GRADE_BANDS if cum_after <= limit + 1e-9), 'C')
        result[key] = {'grade': grade, 'share': round(share, 2), 'cum': round(cum_after, 1), 'qty': int(q)}
        cum = cum_after
    return result


BRAND_RULES = [('바버501', ['바버501']), ('메종드펜세', ['메종드펜세', '메종드 펜세']),
               ('닥터멜락신', ['닥터멜락신', '닥터 멜락신', '멜락신'])]
# 상품 유형 키워드 — 위에서부터 먼저 걸리는 규칙 채택 (구체적인 키워드를 앞에)
PRODUCT_TYPE_RULES = [
    ('세트', ['세트', '기획', 'GWP', '듀오']),
    ('부자재', ['쇼핑백', '샘플', '파우치']),
    ('방향제', ['프레셔너']),
    ('헤어', ['샴푸', '클리닉', '트리트먼트', '헤어']),
    ('쉐이빙', ['쉐이빙', '왁싱']),
    ('이너케어', ['청결제', '페미닌', '인티밋']),
    ('선케어', ['선스크린', '선세럼', '선크림', '썬', '선스틱']),
    ('패드', ['패드']),
    ('마스크/패치', ['마스크', '하이드로겔', '패치', '아이백', '코팩']),
    ('클렌저', ['클렌저', '클렌징', '클렌져', '폼']),
    ('립', ['립']),
    ('바디', ['바디']),
    ('아이크림', ['아이크림', '아이 크림']),
    ('토너', ['토너']),
    ('미스트', ['미스트', '스프레이']),
    ('앰플', ['앰플']),
    ('세럼', ['세럼', '에센스']),
    ('밤', ['밤']),
    ('크림', ['크림', '젤']),
    ('로션/올인원', ['올인원', '로션', '모이스처', '부스터']),
    ('디바이스', ['디바이스', '브러쉬', '소닉샷']),
]


def derive_brand(name):
    s = str(name or '')
    for brand, kws in BRAND_RULES:
        if any(k in s for k in kws):
            return brand
    return '기타'


def derive_product_type(name):
    s = str(name or '')
    for ptype, kws in PRODUCT_TYPE_RULES:
        if any(k in s for k in kws):
            return ptype
    return '기타'


# 제품명 자동분류 유형 → 구매팀 카테고리 파일의 유형 체계 (파일에 없는 코드 표기 통일용)
AUTO_TYPE_TO_FILE_TYPE = {
    '크림': '크림/아이크림', '아이크림': '크림/아이크림', '밤': '크림/아이크림',
    '세럼': '세럼/앰플', '앰플': '세럼/앰플',
    '토너': '스킨/토너', '미스트': '스킨/토너',
    '패드': '마스크/팩/패드/패치', '마스크/패치': '마스크/팩/패드/패치',
    '헤어': '헤어/바디', '바디': '헤어/바디', '쉐이빙': '헤어/바디', '이너케어': '헤어/바디',
    '선케어': '자외선차단제', '립': '립메이크업', '방향제': '퍼퓸/방향', '세트': 'GWP',
}


def load_category_file(folder=None):
    """구매팀 카테고리 정리 엑셀(`카테고리 정리*.xlsx`, mtime 최신)에서 품목코드 → {brand, line, ptype}.
    헤더에 '브랜드명'·'품목코드'·'유형'이 있는 시트를 사용. '대표 코드'도 같은 카테고리로 등록(통합 제품명 행 조회용).
    반환: ({코드: {...}}, 파일명 또는 None)"""
    folder = folder or os.path.dirname(os.path.abspath(__file__))
    files = [f for f in glob.glob(os.path.join(folder, '카테고리 정리*.xlsx')) if not os.path.basename(f).startswith('~$')]
    if not files:
        return {}, None
    path = max(files, key=os.path.getmtime)
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        print(f"   > ⚠️ 카테고리 파일 읽기 실패({os.path.basename(path)}): {e}")
        return {}, None
    for sn in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sn, header=None, nrows=None)
        if df.empty:
            continue
        head = [str(v).strip() if pd.notna(v) else '' for v in df.iloc[0].tolist()]
        if not all(h in head for h in ('브랜드명', '품목코드', '유형')):
            continue
        ci = {h: head.index(h) for h in ('브랜드명', '품목코드', '유형')}
        line_col = head.index('상품라인') if '상품라인' in head else None
        rep_col = head.index('대표 코드') if '대표 코드' in head else None

        def val(row, i):
            if i is None:
                return ''
            v = row[i]
            v = '' if pd.isna(v) else re.sub(r'\s+', ' ', str(v)).strip()
            return '' if v.lower() == 'nan' else v
        out, reps = {}, {}
        for row in df.iloc[1:].itertuples(index=False):
            row = list(row)
            info = {'brand': val(row, ci['브랜드명']), 'line': val(row, line_col), 'ptype': val(row, ci['유형'])}
            for c in val(row, ci['품목코드']).split(','):
                c = c.strip().upper()
                if c:
                    out.setdefault(c, info)
            rc = val(row, rep_col).upper()
            if rc:
                reps.setdefault(rc, info)
        for c, info in reps.items():
            out.setdefault(c, info)
        print(f"   > 카테고리 파일: {os.path.basename(path)} [{sn}] {len(out)}개 코드")
        return out, os.path.basename(path)
    print(f"   > ⚠️ 카테고리 파일에서 '브랜드명/품목코드/유형' 헤더 시트를 못 찾음: {os.path.basename(path)}")
    return {}, None


def build_category_lookup(master_df):
    """마스터 계정에 '브랜드명'/'상품 유형' 컬럼이 있으면 그 값을 우선 사용 (수기 관리),
    없거나 빈칸이면 제품명 키워드로 자동 분류. 반환: ({품목코드(upper): (브랜드, 상품유형)}, 브랜드컬럼有, 유형컬럼有)"""
    lookup = {}
    if master_df is None or '품목코드' not in master_df.columns:
        return lookup, False, False
    brand_col = next((c for c in master_df.columns if str(c).strip() in ('브랜드명', '브랜드')), None)
    type_col = next((c for c in master_df.columns if str(c).replace(' ', '') in ('상품유형', '제품유형')), None)
    for _, row in master_df.dropna(subset=['품목코드']).iterrows():
        name = row.get('제품명', '')
        b = str(row.get(brand_col, '') or '').strip() if brand_col else ''
        t = str(row.get(type_col, '') or '').strip() if type_col else ''
        if b.lower() == 'nan':
            b = ''
        if t.lower() == 'nan':
            t = ''
        val = (b or derive_brand(name), t or derive_product_type(name))
        for c in str(row['품목코드']).split(','):
            if c.strip():
                lookup.setdefault(c.strip().upper(), val)
    return lookup, brand_col is not None, type_col is not None


# 리뉴얼(v2) 권역 정의 — 순서 = 화면 표시 순서
#   full=True : 재고·이동중·일판매·안전재고·회전일 전체 블록 (기존 US/UK/EU)
#   full=False: 재고만 (판매/출고 실적 데이터가 아직 없는 권역 — 2026-11-19 사용자 결정 "재고만 우선 표시")
V2_REGIONS = [
    ('us', '미국', True), ('uk', '영국', True), ('eu', '유럽 연합', True),
    ('jp', '일본', True), ('cn', '중국', False), ('krb2b', '한국 B2B', False), ('hk', '홍콩', False),
    ('tw', '대만', False), ('sg', '싱가포르', False), ('ca', '캐나다', False), ('au', '호주', False),
    ('mx', '멕시코', False),
]
GRADE_REGIONS = ['us', 'uk', 'eu', 'jp']   # 현재 출고/판매 수량이 있는 권역 (나머지는 0으로 등급 계산)

# 일본(JP) — 2026-11-19 추가. 판매 파일 '2026_일본_SCM_재고관리*.xlsx' (US/UK/EU 판매 파일과 같은 양식).
# 재고 = 일본 창고(국제·쉽너지) 합계, 이동중 재고 데이터는 아직 없음(항공·해상이동 0).
# 안전재고 30일 / 항공 리드타임 10일 / 해상 리드타임 20일 (사용자 지정) — 화면 임계선·부킹 시간창은 template_v2.html JP 규칙.
SEED_PERIOD_MAX_AGE_DAYS = 21   # 최근 기간 수불부가 이보다 오래되면 3개월 월평균 seeding으로 대체
JP_SAFETY_DAYS = 30
JP_AIR_LEAD_DAYS = 10
JP_SEA_LEAD_DAYS = 20
# 일반/FBA/FBT 재고 파일의 창고명 → 추가 권역 (부분 문자열 매칭, 위에서부터 먼저 걸리는 것)
EXTRA_REGION_WAREHOUSE_KEYWORDS = [
    ('krb2b', ['국내']), ('jp', ['일본']), ('cn', ['중국']), ('hk', ['홍콩']), ('tw', ['대만']),
    ('sg', ['싱가포르']), ('ca', ['캐나다']), ('au', ['호주']), ('mx', ['멕시코']),
]
_LAST_INV_ROWS = None   # load_integrated_inventory가 읽은 전 창고 행 (추가 권역 재고 집계용)


def assign_extra_region(warehouse_name):
    s = str(warehouse_name)
    for region, kws in EXTRA_REGION_WAREHOUSE_KEYWORDS:
        if any(k in s for k in kws):
            return region
    return None


def build_extra_region_stock():
    """추가 권역(JP/CN/KR B2B/HK/TW/SG/CA/AU/MX) 코드별 창고 가용재고.
    KR B2B = 국내 창고 전체 합계 (KCTC·콜로세움·쿠팡·리브·품고·AGCA — 2026-11-19 사용자 결정).
    반환: ({품목코드: {region: qty}}, {region: [창고명...]})"""
    stock, warehouses = {}, {}
    df = _LAST_INV_ROWS
    if df is None or df.empty:
        return stock, warehouses
    for _, row in df.iterrows():
        region = assign_extra_region(row['창고'])
        if not region:
            continue
        code = str(row['품목코드']).strip().upper()
        if not code or code == 'NAN':
            continue
        warehouses.setdefault(region, set()).add(str(row['창고']))
        d = stock.setdefault(code, {})
        d[region] = d.get(region, 0) + int(row['가용재고'] or 0)
    return stock, {r: sorted(v) for r, v in warehouses.items()}


# ── 어크로스 비 수불부 (US/UK 출고량 원본) ──
LEDGER_REGION_KEYWORDS = [('us', ['미국']), ('uk', ['영국']), ('ca', ['캐나다'])]


def _ledger_year_month(path):
    """수불부 파일의 (연도 또는 None, 월). 연도 우선순위: 파일명 '26년 10월'/'2026년 10월'/'2026-10' → 상위 폴더명의 4자리 연도.
    연도를 못 찾으면 None (대상 3개월 중 같은 월로 매칭하되, 연도가 명시된 파일이 있으면 그쪽이 우선)."""
    name = os.path.basename(path)
    mm = re.search(r'(\d{2,4})\s*년\s*(\d{1,2})\s*월', name) or re.search(r'(20\d{2})[.\-_](\d{1,2})(?!\d)', name)
    if mm:
        y = int(mm.group(1))
        return (y + 2000 if y < 100 else y), int(mm.group(2))
    mm = re.search(r'(\d{1,2})\s*월', name)
    if not mm:
        return None, None
    folder_year = re.search(r'(20\d{2})', os.path.basename(os.path.dirname(path)))
    return (int(folder_year.group(1)) if folder_year else None), int(mm.group(1))


def parse_ledger_outbound(path):
    """수불부 1개 파일 → {품목코드: 출고량}. 출고량 = '출고 재고(EA)' 아래 '총합(EA)' − 'fake(seeding)'.
    같은 SKU가 여러 행이면 합산 (층별화 엑셀의 '중복항목 취합' 단계와 동일)."""
    return parse_ledger_file(path)[0]


def parse_ledger_file(path):
    """수불부 1개 파일 → ({품목코드: 출고량(총합−fake)}, {품목코드: fake(seeding) 수량})."""
    df = pd.read_excel(path, header=None)
    hr = None
    for r in range(min(10, len(df))):
        vals = [str(v).strip() for v in df.iloc[r].tolist()]
        if 'SKU' in vals and any('출고' in v for v in vals):
            hr = r
            break
    if hr is None:
        raise ValueError("헤더('SKU' + '출고 재고') 행을 찾지 못함")
    head = [str(v).strip() for v in df.iloc[hr].tolist()]
    sub = [str(v).strip().lower() for v in df.iloc[hr + 1].tolist()]
    code_col = head.index('SKU')
    out_col = next(i for i, v in enumerate(head) if '출고' in v)
    total_col = next((i for i in range(out_col, len(sub)) if '총합' in sub[i]), out_col)
    fake_col = next((i for i in range(out_col, len(sub)) if 'fake' in sub[i] or 'seeding' in sub[i]), None)
    result, fake = {}, {}
    for r in range(hr + 2, len(df)):
        code = df.iat[r, code_col]
        if pd.isna(code):
            continue
        code = str(code).strip().upper()
        if not code or code == 'SKU':
            continue
        f = parse_int(df.iat[r, fake_col]) if fake_col is not None else 0
        q = parse_int(df.iat[r, total_col]) - f
        result[code] = result.get(code, 0) + q
        if f:
            fake[code] = fake.get(code, 0) + f
    return result, fake


def _ledger_period(path):
    """기간 수불부 파일명 → (시작일, 종료일) 또는 None.
    예: '..._수불부_20261101~20261107.xlsx' → (2026-11-01, 2026-11-07), '..._수불부_20261101.xlsx' → 하루치."""
    name = os.path.basename(path)
    m = re.search(r'(?<!\d)(20\d{6})\s*[~\-–]\s*(20\d{6})(?!\d)', name)
    try:
        if m:
            a = datetime.strptime(m.group(1), '%Y%m%d').date()
            b = datetime.strptime(m.group(2), '%Y%m%d').date()
            return (a, b) if a <= b else (b, a)
        m = re.search(r'(?<!\d)(20\d{6})(?!\d)', name)
        if m:
            d = datetime.strptime(m.group(1), '%Y%m%d').date()
            return (d, d)
    except ValueError:
        return None
    return None


def _is_mon_sun(p):
    """월요일 시작·일요일 끝 7일짜리 기간 수불부인지 (워터폴 주차와 같은 단위 → 겹칠 때 우선)."""
    return p["start"].weekday() == 0 and p["days"] == 7


def _ledger_folder(folder=None):
    return folder or os.path.join(os.path.dirname(os.path.abspath(__file__)), '수불부')


def load_ledger_periods(folder=None):
    """수불부/ 폴더(하위 포함)의 기간 수불부(파일명에 YYYYMMDD~YYYYMMDD) 전부 로드 (2026-11-20 주 단위 수불부).
    같은 권역·같은 기간 파일이 여러 개면 mtime 최신 1개. 반환: [{region, start, end, days, out, fake, file}]."""
    folder = _ledger_folder(folder)
    if not os.path.isdir(folder):
        return []
    picked = {}
    for f in glob.glob(os.path.join(folder, '**', '*.xls*'), recursive=True):
        name = os.path.basename(f)
        if name.startswith('~$'):
            continue
        region = next((r for r, kws in LEDGER_REGION_KEYWORDS if any(k in name for k in kws)), None)
        per = _ledger_period(f)
        if not region or not per:
            continue
        key = (region, per)
        mt = os.path.getmtime(f)
        if key not in picked or mt > picked[key][0]:
            picked[key] = (mt, f)
    out = []
    for (region, (a, b)), (_mt, f) in sorted(picked.items(), key=lambda kv: kv[0]):
        try:
            q, fk = parse_ledger_file(f)
        except Exception as e:
            print(f"   > ⚠️ 기간 수불부 파싱 실패({os.path.basename(f)}): {e}")
            continue
        out.append({"region": region, "start": a, "end": b, "days": (b - a).days + 1,
                    "out": q, "fake": fk, "file": os.path.basename(f)})
    return out


def _month_from_periods(periods, region, ym):
    """월 수불부가 없는 달을 기간 수불부로 채울 수 있으면 ({code: 출고}, {code: fake}, [파일명]) 반환.
    조건: 그 달 안에 완전히 들어가는 기간 파일들이 겹치지 않고 1일~말일을 빠짐없이 덮을 때만 (월을 걸친 파일이 있으면 포기)."""
    y, m = ym
    first = datetime(y, m, 1).date()
    last = datetime(y, m, calendar.monthrange(y, m)[1]).date()
    ps = [p for p in periods if p["region"] == region and p["end"] >= first and p["start"] <= last]
    if not ps or any(p["start"] < first or p["end"] > last for p in ps):
        return None
    ps = sorted(ps, key=lambda p: p["start"])
    cur = first
    for p in ps:
        if p["start"] != cur:
            return None
        cur = p["end"] + timedelta(days=1)
    if cur != last + timedelta(days=1):
        return None
    q, fk = {}, {}
    for p in ps:
        for c, v in p["out"].items():
            q[c] = q.get(c, 0) + v
        for c, v in p["fake"].items():
            fk[c] = fk.get(c, 0) + v
    return q, fk, [p["file"] for p in ps]


def load_ledger_outbound(months, folder=None):
    """수불부/ 폴더(하위 폴더 포함)에서 직전 3개월(months)에 해당하는 권역×월 수불부만 골라 출고량 로드.
    같은 권역·월 파일이 여러 개면 연도가 명시된 파일 우선, 그다음 mtime 최신 1개. 3개월 밖 파일은 무시(보관용).
    반환: ({(region, (y,m)): {code: qty}}, {(region,(y,m)): 파일명}, {(region,(y,m)): {code: fake(seeding)}})"""
    folder = _ledger_folder(folder)
    data, files, fakes = {}, {}, {}
    if not os.path.isdir(folder):
        return data, files, fakes
    # 수불부는 계속 쌓아두는 폴더(월별 출고 추이용) — 하위 폴더(예: 수불부6\)도 읽고, 대상 3개월 파일만 골라 씀
    picked = {}   # (region, (y,m)) → (연도명시 여부, mtime, 경로)
    for f in glob.glob(os.path.join(folder, '**', '*.xls*'), recursive=True):
        name = os.path.basename(f)
        if name.startswith('~$'):
            continue
        region = next((r for r, kws in LEDGER_REGION_KEYWORDS if any(k in name for k in kws)), None)
        if _ledger_period(f):   # 기간(주 단위) 수불부는 아래에서 따로 처리 — 월 파일과 이중 집계 방지
            continue
        y, mo = _ledger_year_month(f)
        if not region or not mo:
            continue
        if y is not None:
            if (y, mo) not in months:
                continue
            ym = (y, mo)
        else:
            ym = next(((yy, m) for yy, m in months if m == mo), None)
            if not ym:
                continue
        rank = (y is not None, os.path.getmtime(f))
        key = (region, ym)
        if key not in picked or rank > picked[key][:2]:
            picked[key] = (rank[0], rank[1], f)
    picked = {k: v[2] for k, v in picked.items()}
    for key, f in picked.items():
        try:
            data[key], fakes[key] = parse_ledger_file(f)
            files[key] = os.path.basename(f)
        except Exception as e:
            print(f"   > ⚠️ 수불부 파싱 실패({os.path.basename(f)}): {e} — 이 달은 판매 파일로 대체")
    # 월 수불부가 없는 달: 그 달을 빈틈없이 덮는 기간 수불부가 있으면 합산해서 사용 (월 파일이 있으면 월 파일 우선)
    periods = load_ledger_periods(folder)
    for region in {p["region"] for p in periods}:
        for ym in months:
            if (region, ym) in data:
                continue
            got = _month_from_periods(periods, region, ym)
            if got:
                data[(region, ym)], fakes[(region, ym)], fl = got
                files[(region, ym)] = f"기간 수불부 {len(fl)}개 합산"
    return data, files, fakes


def build_unregistered_rows(df):
    """마스터 미등록(제품명 없음) + US/UK/EU 재고 있는 코드의 JSON 행. 이름은 재고 파일 상품명, 그룹은 코드 자신."""
    if df is None or df.empty:
        return []
    names = {}
    inv = _LAST_INV_ROWS
    if inv is not None and '상품명' in inv.columns:
        for code, nm in zip(inv['품목코드'], inv['상품명']):
            c = str(code).strip().upper()
            nm = str(nm or '').strip()
            if nm and nm.lower() != 'nan' and c not in names:
                names[c] = nm
    df = df.copy()
    df['제품명'] = df['품목코드'].apply(lambda c: names.get(str(c).strip().upper()) or str(c))
    df['통합 제품명'] = df['품목코드'].astype(str).str.strip()
    df['상품라인'] = '기타(미등록)'
    rows = build_json_rows(df)
    for r in rows:
        r['unregistered'] = True
    return rows


def _zero_row_template(rows):
    """재고·판매가 전부 0인 기존 행을 골라 행 양식으로 사용 (회전일 '여유' 등 표기 규칙을 build_json_rows와 동일하게)."""
    def is_zero(r):
        return all(not r.get(f'{k}_{f}') for k in ('us', 'uk', 'eu') for f in ('total', 'sales'))
    tpl = next((r for r in rows if is_zero(r)), rows[0] if rows else None)
    if tpl is None:
        return None
    out = {}
    for k, v in tpl.items():
        if isinstance(v, bool):
            out[k] = False
        elif isinstance(v, (int, float)):
            out[k] = 0
        elif isinstance(v, list):
            out[k] = []
        elif isinstance(v, dict):
            out[k] = {}
        elif k.endswith('_days'):
            out[k] = v          # 재고·판매 0 행의 회전일 표기 그대로
        else:
            out[k] = '' if v is not None else None
    return out


def _inject_extra_region_rows(json_tab1, json_tab2, master_df, extra_stock, group_extra, extra_keys, excluded_codes):
    """새 권역 재고가 있는데 대시보드 행이 없는 코드(tab1)·통합 제품명(tab2)을 0 행으로 추가.
    마스터 계정에 등록된 코드만 (제품명·그룹을 알 수 있어야 함). 반환: (tab1 추가 수, tab2 추가 수)"""
    if master_df is None or '품목코드' not in master_df.columns:
        return 0, 0
    info = {}
    for _, r in master_df.dropna(subset=['품목코드']).iterrows():
        for c in str(r['품목코드']).split(','):
            c = c.strip().upper()
            if c and c not in info:
                g = str(r.get('통합 제품명', '') or '').strip()
                info[c] = {'name': str(r.get('제품명', '') or '').strip(),
                           'group': '' if g.lower() == 'nan' else g,
                           'line': str(r.get('상품라인', '') or '').strip()}
    added1 = added2 = 0
    tpl1, tpl2 = _zero_row_template(json_tab1), _zero_row_template(json_tab2)
    have1 = {str(r['code']).strip().upper() for r in json_tab1}
    if tpl1:
        for c in sorted(extra_stock):
            if c in have1 or c in excluded_codes or c not in info:
                continue
            if not any(extra_stock[c].get(k, 0) > 0 for k in extra_keys):
                continue
            row = dict(tpl1)
            row.update({'code': c, 'name': info[c]['name'] or c, '통합 제품명': info[c]['group'],
                        '상품라인': '' if info[c]['line'].lower() == 'nan' else info[c]['line'],
                        'plc_status': '정보없음'})
            json_tab1.append(row)
            added1 += 1
    have2 = {str(r.get('통합 제품명', '')) for r in json_tab2}
    if tpl2:
        for g in sorted(group_extra):
            if g in have2 or not any(group_extra[g].get(k, 0) > 0 for k in extra_keys):
                continue
            rep = info.get(g.upper()) or next((v for v in info.values() if v['group'] == g), None)
            if not rep:
                continue
            row = dict(tpl2)
            row.update({'code': g, 'name': rep['name'] or g, '통합 제품명': g,
                        '상품라인': '' if rep['line'].lower() == 'nan' else rep['line'], 'plc_status': '정보없음'})
            json_tab2.append(row)
            added2 += 1
    return added1, added2


PAST_REGIONS = ['us', 'uk', 'eu', 'jp']
LEDGER_SALES_REGIONS = ('us', 'uk')   # 지난 실적 주판매를 수불부 출고(총합 − fake)로 보는 권역 — 나머지는 판매 파일


def past_week_ranges(today=None):
    """지난달 1일 ~ 어제를 월~일 주차로 자른 목록 [(시작, 끝)]. 첫 주는 지난달 1일부터, 마지막 주는 어제까지(이번 주 경과분)."""
    today = today or datetime.now().date()
    first = (today.replace(day=1) - timedelta(days=1)).replace(day=1)
    end_all = today - timedelta(days=1)
    out, cur = [], first
    while cur <= end_all:
        wend = min(cur + timedelta(days=6 - cur.weekday()), end_all)
        out.append((cur, wend))
        cur = wend + timedelta(days=1)
    return out


def _past_daily_seed(weeks):
    """과거 주차 기간의 권역별 일자별 수불부 값 → (fake, src, out).
    fake/out = {region: {code: {date: qty}}} (out = 출고 총합 − fake, 즉 seeding 뺀 출고), src = {region: {date: 'period'|'monthly'}}.
    기간(주 단위) 수불부가 덮는 날은 그 파일 fake ÷ 기간 일수, 아니면 그 달 월 수불부 fake ÷ 그 달 일수 (둘 다 없으면 없음)."""
    if not weeks:
        return {}, {}
    start, end = weeks[0][0], weeks[-1][1]
    months, d = [], start.replace(day=1)
    while d <= end:
        months.append((d.year, d.month))
        d = (d.replace(day=28) + timedelta(days=4)).replace(day=1)
    mout, _files, mfake = load_ledger_outbound(months)
    # 같은 날을 여러 기간 파일이 덮으면 월~일 7일 파일을 우선 (워터폴 주차와 같은 단위), 그다음 짧은 기간
    periods = sorted(load_ledger_periods(), key=lambda p: (not _is_mon_sun(p), p["days"], p["start"]))
    seed, src, outd = {}, {}, {}
    day = start
    while day <= end:
        for p in periods:
            if p["start"] <= day <= p["end"]:
                rg = p["region"]
                if src.get(rg, {}).get(day) == 'period':
                    continue    # 같은 날을 덮는 기간 파일이 여러 개면 먼저 잡힌 1개만
                src.setdefault(rg, {})[day] = 'period'
                for c, q in p["fake"].items():
                    e = seed.setdefault(rg, {}).setdefault(c, {})
                    e[day] = e.get(day, 0) + q / p["days"]
                for c, q in p["out"].items():
                    e = outd.setdefault(rg, {}).setdefault(c, {})
                    e[day] = e.get(day, 0) + q / p["days"]
        for (rg, ym), fk in mfake.items():
            if ym != (day.year, day.month) or src.get(rg, {}).get(day):
                continue
            mdays = calendar.monthrange(*ym)[1]
            src.setdefault(rg, {})[day] = 'monthly'
            for c, q in fk.items():
                e = seed.setdefault(rg, {}).setdefault(c, {})
                e[day] = e.get(day, 0) + q / mdays
            for c, q in mout.get((rg, ym), {}).items():
                e = outd.setdefault(rg, {}).setdefault(c, {})
                e[day] = e.get(day, 0) + q / mdays
        day += timedelta(days=1)
    return seed, src, outd


def _load_snapshot_stock(dates, snap_dir="snapshots"):
    """날짜별 동결본(full_YYYY-MM-DD.json)에서 권역 창고재고({r}_stock)만 → {date: {'tab1': {code: {r: qty}}, 'tab2': {...}}}."""
    out = {}
    for dt in dates:
        path = os.path.join(snap_dir, f"full_{dt.isoformat()}.json")
        if not os.path.exists(path):
            continue
        try:
            with open(path, encoding='utf-8') as f:
                snap = json.load(f)
        except Exception:
            continue
        out[dt] = {tk: {str(r.get('code')): {rg: r.get(f'{rg}_stock') for rg in PAST_REGIONS if r.get(f'{rg}_stock') is not None}
                        for r in snap.get(tk, [])} for tk in ('tab1', 'tab2')}
    return out


def attach_past_weeks(json_tab1, json_tab2, master_df, region_daily, today=None):
    """워터폴 위쪽에 보여줄 '지난달 1일 ~ 어제' 주차별 실적을 행마다 부착 (2026-11-20).
    행 필드 {r}_past = [[주판매, seeding, 기초재고, 기말재고], ...] (meta weeks 순서, 모르는 값은 None).
    - 주판매: US/UK = 수불부 출고(총합 − fake, 주간 파일 우선·없으면 월 파일 일할), EU/JP = 판매 파일 일자별 실적
      (층별화 등급과 같은 기준, 2026-11-20). 통합 제품명은 멤버 코드 합산
    - seeding: 수불부 fake (주 단위 파일 우선, 없으면 월 파일 일할) — US/UK만
    - 재고: 그 주 첫날(없으면 그 주 안의 가장 이른 날) 동결본의 창고재고, 기말 = 다음 주 기초 (마지막 주는 현재 재고)"""
    today = today or datetime.now().date()
    weeks = past_week_ranges(today)
    if not weeks:
        return {}
    seed, seed_src, ledger_out = _past_daily_seed(weeks)
    all_days = [weeks[0][0] + timedelta(days=i) for i in range((weeks[-1][1] - weeks[0][0]).days + 1)]
    snaps = _load_snapshot_stock(all_days + [today])
    # 주차별 기초재고에 쓸 스냅샷 날짜 (그 주 안의 가장 이른 날)
    snap_day = []
    for a, b in weeks:
        snap_day.append(next((d for d in sorted(snaps) if a <= d <= b), None))
    # 판매 데이터가 있는 날짜 범위 (권역별) — 범위 밖 주차는 '모름'(None)
    sales_range = {}
    for rg in PAST_REGIONS:
        ds = [d for dm in (region_daily.get(rg) or {}).values() for d in dm]
        if ds:
            sales_range[rg] = (min(ds), max(ds))

    code2group = {}
    if master_df is not None and '통합 제품명' in master_df.columns:
        for _, r in master_df.dropna(subset=['품목코드']).iterrows():
            g = str(r.get('통합 제품명', '') or '').strip()
            if g and g.lower() != 'nan':
                for c in str(r['품목코드']).split(','):
                    if c.strip():
                        code2group.setdefault(c.strip().upper(), g)
    members = {}
    for c, g in code2group.items():
        members.setdefault(g, set()).add(c)
    for row in json_tab1:
        g = str(row.get('통합 제품명', '') or '')
        if g:
            members.setdefault(g, set()).add(str(row.get('code', '')).strip().upper())

    def week_sum(daymap_list, a, b):
        tot, d = 0, a
        while d <= b:
            for dm in daymap_list:
                tot += dm.get(d, 0)
            d += timedelta(days=1)
        return tot

    def fill(row, codes, tab_key):
        key = str(row.get('code'))
        for rg in PAST_REGIONS:
            if rg in LEDGER_SALES_REGIONS:
                sdaily = [ledger_out[rg][c] for c in codes if rg in ledger_out and c in ledger_out[rg]]
            else:
                sdaily = [region_daily[rg][c] for c in codes if region_daily.get(rg) and c in region_daily[rg]]
            fdaily = [seed[rg][c] for c in codes if rg in seed and c in seed[rg]]
            vals, any_val = [], False
            for i, (a, b) in enumerate(weeks):
                has_src = rg in seed_src and any(seed_src[rg].get(a + timedelta(days=k)) for k in range((b - a).days + 1))
                if rg in LEDGER_SALES_REGIONS:
                    sales = int(round(week_sum(sdaily, a, b))) if has_src else None
                else:
                    rng = sales_range.get(rg)
                    sales = week_sum(sdaily, a, b) if rng and a >= rng[0] else None
                sd = int(round(week_sum(fdaily, a, b))) if has_src else None
                op = (snaps.get(snap_day[i]) or {}).get(tab_key, {}).get(key, {}).get(rg) if snap_day[i] else None
                if i + 1 < len(weeks):
                    nd = snap_day[i + 1]
                    cl = (snaps.get(nd) or {}).get(tab_key, {}).get(key, {}).get(rg) if nd else None
                else:
                    cl = row.get(f'{rg}_stock')
                if sales or sd or op or cl:
                    any_val = True
                vals.append([sales, sd, op, cl])
            if any_val:
                row[f'{rg}_past'] = vals

    for row in json_tab1:
        fill(row, {str(row.get('code', '')).strip().upper()}, 'tab1')
    for row in json_tab2:
        g = str(row.get('통합 제품명', '') or '')
        codes = set(members.get(g, set())) | {str(row.get('code', '')).strip().upper()}
        fill(row, codes, 'tab2')

    src_meta = {}
    for rg, dm in seed_src.items():
        lst = []
        for a, b in weeks:
            got = [dm.get(a + timedelta(days=k)) for k in range((b - a).days + 1)]
            kinds = set(got) - {None}
            kind = 'period' if kinds == {'period'} else 'monthly' if kinds == {'monthly'} else ('mixed' if kinds else None)
            lst.append([kind, sum(1 for g in got if g)] if kind else None)   # [종류, 수불부가 덮는 일수]
        src_meta[rg] = lst
    meta = {"weeks": [[a.isoformat(), b.isoformat()] for a, b in weeks],
            "ledger_sales_regions": list(LEDGER_SALES_REGIONS),
            "snap_days": [d.isoformat() if d else None for d in snap_day],
            "seed_src": src_meta,
            "sales_range": {rg: [a.isoformat(), b.isoformat()] for rg, (a, b) in sales_range.items()}}
    print(f"   > 지난 실적(워터폴): {weeks[0][0]} ~ {weeks[-1][1]} {len(weeks)}개 주차, "
          f"재고 스냅샷 {sum(1 for d in snap_day if d)}/{len(weeks)}주, seeding 권역 {', '.join(r.upper() for r in src_meta) or '-'}")
    return meta


def enrich_v2_fields(json_tab1, json_tab2, master_df, sales_files, excluded_codes, unreg_rows=None,
                     jp_sales_df=None, jp_daily=None):
    """층별화 등급 + 카테고리 + 추가 권역 재고를 tab1(개별 코드)/tab2(통합 제품명) JSON 행에 부착하고 메타 반환.
    등급 수량: US/UK = 어크로스 비 수불부 출고량(총합 − fake(seeding), 없는 달은 판매 파일), EU/JP = 판매 파일
    (2026-11-19 사용자 결정, 2026-11-20 재확인 — 판매 파일 통일안은 철회).
    화면에서 권역을 고르면 grade_universe로 브라우저가 등급을 다시 계산함."""
    months = prev_full_months(3)
    ledger, ledger_files, ledger_fake = load_ledger_outbound(months)
    region_qty = {r: {} for r in GRADE_REGIONS}
    sources = {}
    missing_all = {}
    for region in GRADE_REGIONS:
        path = sales_files.get(region)
        sales_totals_by_month = {}
        need_sales = [ym for ym in months if (region, ym) not in ledger]
        if need_sales:
            for ym in need_sales:
                t, _used, miss = parse_monthly_sales_totals(path, [ym])
                sales_totals_by_month[ym] = t
                if miss:
                    missing_all.setdefault(region, []).append(f"{ym[0]}-{ym[1]:02d}")
        src = []
        for ym in sorted(months):
            label = f"{ym[1]}월"
            if (region, ym) in ledger:
                totals = ledger[(region, ym)]       # 수불부 출고 = 총합 − fake(seeding) → seeding 중복 집계 없음
                src.append(f"{label} 수불부")
            else:
                totals = sales_totals_by_month.get(ym, {})
                src.append(f"{label} 판매파일")
            for c, q in totals.items():
                if c not in excluded_codes and q:
                    region_qty[region][c] = region_qty[region].get(c, 0) + q
        sources[region] = src

    # 마스터 미등록이지만 US/UK/EU 재고가 있는 코드 → 개별코드·통합제품명 양쪽에 자기 코드 이름의 행으로 추가
    for r in (unreg_rows or []):
        json_tab1.append(dict(r))
        json_tab2.append(dict(r))
    if unreg_rows:
        print(f"   > 마스터 미등록 + US/UK/EU 재고 있는 코드 {len(unreg_rows)}개 행 추가")

    # tab1 모수: 출고/판매가 있는 전 코드 + 대시보드 표시 코드
    all_codes = set()
    for r in GRADE_REGIONS:
        all_codes |= set(region_qty[r])
    for row in json_tab1:
        all_codes.add(str(row['code']).strip().upper())
    uni1 = {c: [region_qty[r].get(c, 0) for r in GRADE_REGIONS] for c in sorted(all_codes)}

    # tab2: 마스터의 코드→통합 제품명 매핑으로 그룹 합산 (재고 없는 코드의 판매도 포함 — build_group_fcst와 같은 방식)
    code2group = {}
    if master_df is not None and '통합 제품명' in master_df.columns:
        for _, r in master_df.dropna(subset=['품목코드']).iterrows():
            g = str(r.get('통합 제품명', '') or '').strip()
            if g and g.lower() != 'nan':
                for c in str(r['품목코드']).split(','):
                    if c.strip():
                        code2group.setdefault(c.strip().upper(), g)
    uni2 = {}
    for c, vec in uni1.items():
        g = code2group.get(c, c)   # 마스터 미등록 코드는 자기 코드가 곧 그룹 (미등록 행의 '통합 제품명'=코드)
        acc = uni2.setdefault(g, [0] * len(GRADE_REGIONS))
        for i, q in enumerate(vec):
            acc[i] += q
    for row in json_tab2:
        uni2.setdefault(str(row.get('통합 제품명', '')), [0] * len(GRADE_REGIONS))

    grades1 = assign_grades({k: sum(v) for k, v in uni1.items()})
    grades2 = assign_grades({k: sum(v) for k, v in uni2.items()})
    cat_lookup, has_brand_col, has_type_col = build_category_lookup(master_df)
    cat_file, cat_file_name = load_category_file()
    group_codes = {}
    for c, g in code2group.items():
        group_codes.setdefault(g, []).append(c)
    cat_miss = set()

    def apply_cat(row, codes):
        """브랜드·상품라인·상품 유형: 카테고리 정리 엑셀(구매팀) 우선 → 마스터 컬럼 → 제품명 자동분류."""
        info = next((cat_file[c] for c in codes if c in cat_file), None)
        if info:
            row['brand'] = info['brand'] or derive_brand(row['name'])
            row['ptype'] = info['ptype'] or derive_product_type(row['name'])
            if info['line']:
                row['상품라인'] = info['line']
            return
        if cat_file:
            cat_miss.add(codes[0])
        b, t = cat_lookup.get(codes[0]) or (derive_brand(row['name']), derive_product_type(row['name']))
        if cat_file:   # 파일에 없는 코드도 파일과 같은 유형 체계로 표기
            t = AUTO_TYPE_TO_FILE_TYPE.get(t, t)
        row['brand'], row['ptype'] = b, t

    # 추가 권역 재고 (tab2는 마스터 매핑으로 그룹 합산)
    extra_stock, extra_wh = build_extra_region_stock()
    # 창고 재고를 재고 파일 창고명으로 붙이는 권역: 재고만 권역 + JP(전체 블록이지만 재고 출처는 같은 방식)
    extra_keys = [r for r, _l, full in V2_REGIONS if not full] + ['jp']
    group_extra = {}
    for c, d in extra_stock.items():
        g = code2group.get(c)
        if g:
            acc = group_extra.setdefault(g, {})
            for r, q in d.items():
                acc[r] = acc.get(r, 0) + q

    # 새 권역(JP·KR B2B·HK…)에만 재고가 있는 코드/그룹도 행으로 표시 (판매·출고 없으면 재고 수량만 보임 — 2026-11-19)
    added1, added2 = _inject_extra_region_rows(json_tab1, json_tab2, master_df, extra_stock, group_extra,
                                               extra_keys, excluded_codes)
    print(f"   > 새 권역 재고만 있는 행 추가: 개별코드 {added1}개 / 통합제품명 {added2}개")

    # JP 일판매 (개별코드: 코드별 / 통합제품명: 멤버 코드 일자별 합산 후 재계산 — build_group_sales와 같은 방식)
    jp_code = {}
    if jp_sales_df is not None and len(jp_sales_df):
        for rec in jp_sales_df.to_dict('records'):
            jp_code[str(rec['품목코드']).strip().upper()] = rec
    jp_group = {}
    if jp_daily:
        gd = {}
        for c, d in jp_daily.items():
            g = code2group.get(str(c).strip().upper())
            if not g:
                continue
            e = gd.setdefault(g, {})
            for dt, q in d.items():
                e[dt] = e.get(dt, 0) + q
        if gd:
            for rec in summarize_sales_28d(gd, 'jp').to_dict('records'):
                jp_group[rec['품목코드']] = rec
    elif jp_code:   # 일자별 파싱 실패(당월 일평균 폴백) → 코드 합산
        for c, rec in jp_code.items():
            g = code2group.get(c)
            if g:
                acc = jp_group.setdefault(g, {'jp_sales': 0, 'jp_sales_window': 28, 'jp_last7': [0] * 7, 'jp_spike': False})
                acc['jp_sales'] += int(rec.get('jp_sales', 0) or 0)

    # 일본에서 팔리는데 재고가 0이라 행이 없는 코드/그룹도 표시 (결품 판매중 — US/UK/EU와 같은 규칙)
    jp_sell_codes = {c: {'jp': 1} for c, rec in jp_code.items() if (rec.get('jp_sales') or 0) > 0
                     or any((rec.get('jp_last7') or []))}
    jp_sell_groups = {g: {'jp': 1} for g, rec in jp_group.items() if (rec.get('jp_sales') or 0) > 0}
    a1, a2 = _inject_extra_region_rows(json_tab1, json_tab2, master_df, jp_sell_codes, jp_sell_groups,
                                       ['jp'], excluded_codes)
    if a1 or a2:
        print(f"   > 일본 판매 중인데 재고 0인 행 추가: 개별코드 {a1}개 / 통합제품명 {a2}개")

    # fake(seeding) 일평균 (권역별, 수불부 있는 US/UK만). 워터폴에서 주차 일수 × 일평균만큼 잔여재고에서 차감.
    #  ① 기간(주 단위) 수불부가 있으면 가장 최근 파일 1개의 fake ÷ 그 기간 일수 (2026-11-20 주 단위 전환)
    #     — 최근 파일 종료일이 SEED_PERIOD_MAX_AGE_DAYS보다 오래됐으면 ②로 대체
    #  ② 없으면 최근 3개월 월 수불부 fake 합계 ÷ 그 달들 일수
    seed_code, seed_meta = {}, {}
    periods = load_ledger_periods()
    today = datetime.now().date()
    for rg in sorted({k[0] for k in ledger_fake} | {p["region"] for p in periods}):
        rps = [p for p in periods if p["region"] == rg]
        latest = max(rps, key=lambda p: (p["end"], _is_mon_sun(p), p["start"])) if rps else None
        if latest and (today - latest["end"]).days <= SEED_PERIOD_MAX_AGE_DAYS:
            tot = {c: q for c, q in latest["fake"].items() if c not in excluded_codes}
            days = latest["days"]
            for c, q in tot.items():
                seed_code.setdefault(c, {})[rg] = q / days
            seed_meta[rg] = {"mode": "period", "start": latest["start"].isoformat(),
                             "end": latest["end"].isoformat(), "days": days, "file": latest["file"],
                             "total": int(sum(tot.values())), "codes": len(tot)}
            continue
        yms = sorted(ym for (r2, ym) in ledger_fake if r2 == rg)
        days = sum(calendar.monthrange(y, m)[1] for y, m in yms)
        if not days:
            continue
        tot = {}
        for ym in yms:
            for c, q in ledger_fake[(rg, ym)].items():
                if c not in excluded_codes:
                    tot[c] = tot.get(c, 0) + q
        for c, q in tot.items():
            seed_code.setdefault(c, {})[rg] = q / days
        seed_meta[rg] = {"mode": "monthly", "months": [f"{y}-{m:02d}" for y, m in yms], "days": days,
                         "total": int(sum(tot.values())), "codes": len(tot)}
    for rg, meta in seed_meta.items():
        if meta["mode"] == "period":
            print(f"   > seeding({rg.upper()}): 기간 수불부 {meta['start']}~{meta['end']} ({meta['days']}일) "
                  f"fake {meta['total']:,}개 / {meta['codes']}개 코드")
        else:
            print(f"   > seeding({rg.upper()}): 월 수불부 {', '.join(meta['months'])} 평균 ({meta['days']}일) fake {meta['total']:,}개")
    seed_group = {}
    for c, d in seed_code.items():
        g = code2group.get(c, c)
        acc = seed_group.setdefault(g, {})
        for rg, v in d.items():
            acc[rg] = acc.get(rg, 0) + v

    def apply_seed(row, d):
        for rg in GRADE_REGIONS:
            row[f'{rg}_seed_daily'] = round((d or {}).get(rg, 0), 2)

    def apply_jp(row, rec):
        """JP 전체 블록 필드: 이동중 0, 총재고=창고재고, 안전재고=일판매×30일(중량물 +30), 회전일."""
        rec = rec or {}
        sales = int(rec.get('jp_sales', 0) or 0)
        last7 = rec.get('jp_last7')
        row['jp_air'] = row['jp_sea'] = 0
        row['jp_total'] = int(row.get('jp_stock', 0) or 0)
        row['jp_sales'] = sales
        row['jp_sales_window'] = int(rec.get('jp_sales_window', 28) or 28)
        row['jp_last7'] = [int(x) for x in last7] if isinstance(last7, list) else [0] * 7
        row['jp_spike'] = bool(rec.get('jp_spike', False))
        heavy = HEAVY_SAFETY_EXTRA_DAYS if row.get('is_heavy') else 0
        row['jp_safety'] = int(sales * (JP_SAFETY_DAYS + heavy))
        row['jp_days'] = "여유" if sales <= 0 else str(int(row['jp_total'] / sales))

    empty = {'grade': 'C', 'share': 0, 'cum': 100, 'qty': 0}
    for row in json_tab1:
        key = str(row['code']).strip().upper()
        g = grades1.get(key, empty)
        row['grade'], row['grade_share'], row['grade_cum'], row['sales_3m'] = g['grade'], g['share'], g['cum'], g['qty']
        apply_cat(row, [key])
        es = extra_stock.get(key, {})
        for r in extra_keys:
            row[f'{r}_stock'] = int(es.get(r, 0))
        apply_jp(row, jp_code.get(key))
        apply_seed(row, seed_code.get(key))
    for row in json_tab2:
        gname = str(row.get('통합 제품명', ''))
        g = grades2.get(gname, empty)
        row['grade'], row['grade_share'], row['grade_cum'], row['sales_3m'] = g['grade'], g['share'], g['cum'], g['qty']
        rcode = str(row['code']).strip().upper()
        apply_cat(row, [rcode, gname.strip().upper()] + group_codes.get(gname, []))
        es = group_extra.get(gname) or extra_stock.get(str(row['code']).strip().upper(), {})
        for r in extra_keys:
            row[f'{r}_stock'] = int(es.get(r, 0))
        apply_jp(row, jp_group.get(gname) or jp_code.get(rcode))
        apply_seed(row, seed_group.get(gname) or seed_code.get(rcode))

    def count(rows):
        cnt = {'SA': 0, 'A': 0, 'B': 0, 'C': 0}
        for r in rows:
            cnt[r['grade']] += 1
        return cnt

    return {
        "months": [f"{y}-{m:02d}" for y, m in sorted(months)],
        "total_qty": int(sum(sum(v) for v in uni1.values())),
        "region_totals": {r: int(sum(region_qty[r].values())) for r in GRADE_REGIONS},
        "sources": sources,
        "ledger_files": {f"{r}-{y}-{m:02d}": n for (r, (y, m)), n in ledger_files.items()},
        "months_missing": missing_all,
        "bands": [{"grade": g, "upto": u} for u, g in GRADE_BANDS],
        "count_tab1": count(json_tab1),
        "count_tab2": count(json_tab2),
        "brand_source": (f"{cat_file_name} 우선" if cat_file_name else "") +
                        (" → 마스터 '브랜드명'" if has_brand_col else "") + " → 제품명 자동분류",
        "type_source": (f"{cat_file_name} 우선 (없는 코드 {len(cat_miss)}개는 자동분류)" if cat_file_name else
                        ("마스터 '상품 유형' 컬럼 우선, 빈칸은 제품명 자동분류" if has_type_col else "제품명 자동분류")),
        "category_missing": sorted(cat_miss),
        "regions": [{"key": r, "label": l, "full": full} for r, l, full in V2_REGIONS],
        "grade_regions": GRADE_REGIONS,
        "seed_meta": seed_meta,
        "jp_rule": {"safety_days": JP_SAFETY_DAYS, "air_lead_days": JP_AIR_LEAD_DAYS, "sea_lead_days": JP_SEA_LEAD_DAYS},
        "extra_warehouses": extra_wh,
        "universe_tab1": uni1,
        "universe_tab2": uni2,
    }


def build_json_rows(df):
    def calc_days(total, sales):
        if sales <= 0:
            return "여유"
        try:
            return str(int(total / sales))
        except:
            return "0"

    for r in ['us', 'uk', 'eu']:
        df[f'{r}_total'] = df[f'{r}_stock'] + df[f'{r}_air'] + df[f'{r}_sea']

    if 'is_heavy' not in df.columns:
        df['is_heavy'] = False
    df['is_heavy'] = df['is_heavy'].fillna(False).astype(bool)
    heavy_extra = df['is_heavy'].astype(int) * HEAVY_SAFETY_EXTRA_DAYS
    df['us_safety'] = df['us_sales'] * (US_SAFETY_DAYS + heavy_extra)
    df['uk_safety'] = df['uk_sales'] * (180 + heavy_extra)
    df['eu_safety'] = df['eu_sales'] * (180 + heavy_extra)

    df['us_days'] = df.apply(lambda row: calc_days(row['us_total'], row['us_sales']), axis=1)
    df['uk_days'] = df.apply(lambda row: calc_days(row['uk_total'], row['uk_sales']), axis=1)
    df['eu_days'] = df.apply(lambda row: calc_days(row['eu_total'], row['eu_sales']), axis=1)

    json_data = []
    for _, row in df.iterrows():
        json_data.append({
            "code": str(row['품목코드']),
            "name": str(row['제품명']),
            "상품라인": str(row.get('상품라인', '기타(미등록)')),
            "통합 제품명": str(row.get('통합 제품명', '미등록')),
            "us_stock": int(row['us_stock']),
            "us_air": int(row['us_air']),
            "us_sea": int(row['us_sea']),
            "us_total": int(row['us_total']),
            "us_sales": int(row['us_sales']),
            "us_safety": int(row['us_safety']),
            "us_days": str(row['us_days']),
            "uk_stock": int(row['uk_stock']),
            "uk_air": int(row['uk_air']),
            "uk_sea": int(row['uk_sea']),
            "uk_total": int(row['uk_total']),
            "uk_sales": int(row['uk_sales']),
            "uk_safety": int(row['uk_safety']),
            "uk_days": str(row['uk_days']),
            "eu_stock": int(row['eu_stock']),
            "eu_air": int(row['eu_air']),
            "eu_sea": int(row['eu_sea']),
            "eu_total": int(row['eu_total']),
            "eu_sales": int(row['eu_sales']),
            "eu_safety": int(row['eu_safety']),
            "eu_days": str(row['eu_days']),
            "plc_status": str(row.get('plc_status', '정보없음')),
            "us_value": int(row.get('us_value', 0)),
            "uk_value": int(row.get('uk_value', 0)),
            "eu_value": int(row.get('eu_value', 0)),
            "stock_value": int(row.get('stock_value', 0)),
            "pallet_units": int(row.get('pallet_units', 0)),
            "b2b_qty": int(row.get('b2b_qty', 0)),
            "kr_stock": int(row.get('kr_stock', 0)),
            "is_heavy": bool(row.get('is_heavy', False)),
            "us_spike": bool(row.get('us_spike', False)),
            "uk_spike": bool(row.get('uk_spike', False)),
            "eu_spike": bool(row.get('eu_spike', False)),
            "us_sales_window": int(row.get('us_sales_window', 28) or 28),
            "uk_sales_window": int(row.get('uk_sales_window', 28) or 28),
            "eu_sales_window": int(row.get('eu_sales_window', 28) or 28),
            "us_last7": [int(x) for x in (row.get('us_last7') if isinstance(row.get('us_last7'), list) else [0]*7)],
            "uk_last7": [int(x) for x in (row.get('uk_last7') if isinstance(row.get('uk_last7'), list) else [0]*7)],
            "eu_last7": [int(x) for x in (row.get('eu_last7') if isinstance(row.get('eu_last7'), list) else [0]*7)],
            "fcst": row.get('fcst') if isinstance(row.get('fcst'), dict) else {},
            "mfr": str(row.get('mfr', '') or '')
        })
    return json_data


def run_dashboard_generator():
    # 파일 경로 — 날짜 자동 탐색
    file_gen = find_latest_inventory_file("일반_전체재고_")
    file_fbt = find_latest_inventory_file("FBT_전체재고_")
    file_fba = find_latest_inventory_file("FBA_전체재고_")

    excel_us_sales = find_latest_sales_file(["미국"])
    excel_uk_sales = find_latest_sales_file(["영국", "UK"])
    excel_eu_sales = find_latest_sales_file(["EU"])
    excel_jp_sales = find_latest_sales_file(["일본"])   # 없으면 리뉴얼 화면의 JP 판매만 0 (빌드는 계속)
    excel_transit  = "2025_이동중_재고_현황.xlsx"
    excel_master   = "마스터 계정.xlsx"
    excel_inbound  = "2026_SCM_영업채널별_판매_포캐스팅_공유용.xlsx"

    template_path  = "template.html"
    output_path    = "dashboard.html"

    for name, path in [('일반재고', file_gen), ('FBT재고', file_fbt), ('FBA재고', file_fba),
                       ('US판매', excel_us_sales), ('UK판매', excel_uk_sales), ('EU판매', excel_eu_sales)]:
        if path is None:
            print(f"❌ {name} 파일을 찾을 수 없습니다.")
            return
        print(f"   > {name}: {path} 자동 선택")

    print("🔄 [1/7] 마스터 계정 기반 품목코드->제품명 사전 구축 중...")
    master_df = load_robust_excel(excel_master, '품목코드')
    global_name_dict = build_master_lookup(master_df)
    print(f"   > {len(global_name_dict)}개 단일코드-제품명 매핑 완료")
    excluded_codes = build_excluded_codes(master_df)
    print(f"   > 마스터 계정 '제외' 표시: {len(excluded_codes)}개 품목코드 리포트 제외 대상")
    heavy_codes = build_heavy_codes(master_df)
    if heavy_codes:
        print(f"   > 마스터 계정 '중량물' 표시: {len(heavy_codes)}개 품목코드 — 안전재고/항공 임계선 +{HEAVY_SAFETY_EXTRA_DAYS}일 적용")
    else:
        print("   > 마스터 계정에 '중량물' 컬럼 없음(또는 표시 0건) — 중량물 차등 미적용")

    print("🔄 [2/7] 통합 재고 파일(일반/FBT/FBA) 읽기 중...")
    stock_df = load_integrated_inventory(file_gen, file_fbt, file_fba)

    print("🔄 [3/7] 국가별 일판매 데이터 추출 중 (28일 롤링 평균)...")
    us_sales_df, us_daily = read_sales_rolling(excel_us_sales, "us")
    uk_sales_df, uk_daily = read_sales_rolling(excel_uk_sales, "uk")
    eu_sales_df, eu_daily = read_sales_rolling(excel_eu_sales, "eu")
    if excel_jp_sales:
        print(f"   > JP판매: {excel_jp_sales} 자동 선택")
        jp_sales_df, jp_daily = read_sales_rolling(excel_jp_sales, "jp")
    else:
        print("   > ℹ️ JP판매 파일(*일본*SCM_재고관리*.xlsx) 없음 — 일본 일판매 0으로 표시")
        jp_sales_df, jp_daily = None, {}

    print("🔄 [4/7] 이동중(In-Transit) 합산 데이터 결합 중...")
    transit_df = read_transit_data(excel_transit)

    print("🔄 [5/7] ETA(현지 입고 일정) 파일 파싱 중 (US/UK 항공·해상)...")
    eta_schedule = parse_transit_schedule(excel_transit, year=2026)

    print("🔄 [5.5/7] 국내 메인창고 입고 스케줄 파싱 중...")
    domestic_inbound = parse_domestic_inbound(excel_inbound)

    print("🔄 [5.6/7] 원가/PLC 상태 파싱 중...")
    product_cost = parse_product_cost(excel_inbound)
    product_status = parse_product_status(excel_inbound)
    product_mfr = parse_manufacturer(excel_inbound)
    print(f"   > 원가 매칭 가능 SKU: {len(product_cost)}개, PLC 상태 정보: {len(product_status)}개, 제조사 매핑: {len(product_mfr)}개")

    print("🔄 [5.65/7] 월별 판매예정(포캐스팅) 파싱 중...")
    fcst_map, fcst_months = parse_forecast_sales(excel_inbound)
    print(f"   > 판매예정: {len(fcst_months)}개월 ({', '.join(fcst_months)}), {len(fcst_map)}개 코드")

    # 포캐스팅에는 수량이 있는데 마스터 계정에 없는 코드 경고.
    # 제품 리뉴얼로 코드가 바뀌면(GLMX124 → GLMX124_V02) 포캐스팅은 신코드로 잡히는데
    # 마스터에 신코드가 없으면 그 물량이 대시보드/워터폴에서 조용히 0으로 빠진다 (2026-07-22 발견).
    fcst_unmapped = []
    if master_df is not None and '품목코드' in master_df.columns:
        master_code_set = set()
        for c in master_df['품목코드'].dropna().astype(str):
            for p in c.split(','):
                if p.strip():
                    master_code_set.add(p.strip().upper())
        for code, months in fcst_map.items():
            qty = sum(sum(v.values()) for v in months.values())
            if qty > 0 and code not in master_code_set:
                fcst_unmapped.append((code, qty))
        fcst_unmapped.sort(key=lambda x: -x[1])
    if fcst_unmapped:
        total_missing = sum(q for _, q in fcst_unmapped)
        print(f"   > ⚠️ 포캐스팅 물량이 있으나 마스터 계정에 없는 코드: {len(fcst_unmapped)}개 "
              f"(합계 {total_missing:,}개) — 이 물량은 워터폴/포캐스팅 비교에서 누락됨")
        print(f"     예시: {', '.join(c for c, _ in fcst_unmapped[:10])}"
              f"{' 외 다수' if len(fcst_unmapped) > 10 else ''}")
        print(f"     ↳ 해결: 마스터 계정.xlsx에 해당 코드를 추가하고 '통합 제품명'을 기존 제품과 동일하게 지정")

    print("🔄 [5.7/7] 부킹 관리용 팔렛트 환산표/B2B 사전할당/한국 가용재고 파싱 중...")
    pallet_map = build_pallet_map(excel_master)
    print(f"   > 팔렛트 환산표(패킹리스트): {len(pallet_map)}개 코드 매칭")
    kr_stock_map = parse_kr_available(file_gen)
    print(f"   > 한국(콜로세움) 가용재고: {len(kr_stock_map)}개 코드")
    b2b_map, b2b_path = build_b2b_map("B2B할당_")
    if b2b_path:
        print(f"   > B2B 사전할당: {b2b_path} 자동 선택, {len(b2b_map)}개 코드")
    else:
        print("   > B2B 사전할당: 파일 없음 (전량 0으로 처리)")

    # 데이터 병합
    merged = stock_df.copy()
    if merged.empty:
        merged = pd.DataFrame(columns=['품목코드', 'us_stock', 'uk_stock', 'eu_stock'])

    merged = pd.merge(merged, us_sales_df, on='품목코드', how='outer')
    merged = pd.merge(merged, uk_sales_df, on='품목코드', how='outer')
    merged = pd.merge(merged, eu_sales_df, on='품목코드', how='outer')
    merged = pd.merge(merged, transit_df,  on='품목코드', how='outer')

    # 재고 0 신제품(_V0N) 주입: 마스터엔 등록됐지만 재고·판매·이동중이 전혀 없어
    # merged에 아직 없는 신제품 코드를 빈 행으로 추가 → 재고가 들어오기 전에도 대시보드에 노출
    # (2026-08-27 사용자 결정: 신제품만 항상 표시, 오래된 죽은 코드는 계속 숨김)
    if master_df is not None and '품목코드' in master_df.columns:
        existing_codes = set(merged['품목코드'].dropna().astype(str).str.strip())
        new_prod_codes = sorted({
            c for c in (str(x).strip() for x in master_df['품목코드'].dropna())
            if c and ',' not in c and is_new_product_code(c) and c not in existing_codes
        })
        if new_prod_codes:
            merged = pd.concat([merged, pd.DataFrame({'품목코드': new_prod_codes})],
                               ignore_index=True)
            print(f"   > 🆕 재고 0 신제품(_V0N) {len(new_prod_codes)}개 주입 (대시보드 노출용)")

    fill_numeric = ['us_stock', 'uk_stock', 'eu_stock',
                    'us_sales', 'uk_sales', 'eu_sales',
                    'us_air', 'us_sea', 'uk_air', 'uk_sea', 'eu_air', 'eu_sea']
    for c in fill_numeric:
        if c not in merged.columns:
            merged[c] = 0
    merged[fill_numeric] = merged[fill_numeric].fillna(0)

    # 28일 롤링에서 나온 급증 플래그/최근7일 컬럼 결측 보정 (재고만 있고 판매 데이터 없는 SKU)
    for r in ['us', 'uk', 'eu']:
        spike_col, last7_col = f'{r}_spike', f'{r}_last7'
        if spike_col not in merged.columns:
            merged[spike_col] = False
        merged[spike_col] = merged[spike_col].apply(lambda v: bool(v) if pd.notna(v) else False)
        if last7_col not in merged.columns:
            merged[last7_col] = None
        merged[last7_col] = merged[last7_col].apply(lambda v: v if isinstance(v, list) else [0]*7)
        window_col = f'{r}_sales_window'
        if window_col not in merged.columns:
            merged[window_col] = 28
        merged[window_col] = merged[window_col].fillna(28).astype(int)

    merged['품목코드'] = merged['품목코드'].astype(str).str.strip()
    merged['제품명'] = merged['품목코드'].apply(lambda c: lookup_product_name(c, global_name_dict))

    if excluded_codes:
        before_count = len(merged)
        merged = merged[~merged['품목코드'].str.upper().isin(excluded_codes)]
        print(f"   > 제외 대상 필터링: {before_count - len(merged)}개 행 제외됨")

    # 원가/PLC 상태 부착 + 재고 금액 계산 (그룹핑 전, 개별 코드 단위로 계산해야 통합 제품명 합산 시 정확함)
    merged['unit_cost'] = merged['품목코드'].str.upper().map(product_cost).fillna(0).astype(int)
    merged['plc_status'] = merged['품목코드'].str.upper().map(product_status).fillna('정보없음')
    merged['mfr'] = merged['품목코드'].apply(lambda c: lookup_manufacturer(c, product_mfr))
    merged['pallet_units'] = merged['품목코드'].str.upper().map(pallet_map).fillna(0).astype(int)
    merged['b2b_qty'] = merged['품목코드'].str.upper().map(b2b_map).fillna(0).astype(int)
    merged['kr_stock'] = merged['품목코드'].str.upper().map(kr_stock_map).fillna(0).astype(int)
    merged['is_heavy'] = merged['품목코드'].str.upper().isin(heavy_codes)
    # 월별 판매예정(포캐스팅) 부착 — 코드 단위. tab2에서는 그룹 내 코드끼리 월별/권역별 합산됨.
    merged['fcst'] = merged['품목코드'].str.upper().map(lambda c: fcst_map.get(c, {}))
    for r in ['us', 'uk', 'eu']:
        merged[f'{r}_total'] = merged[f'{r}_stock'] + merged[f'{r}_air'] + merged[f'{r}_sea']
        merged[f'{r}_value'] = merged[f'{r}_total'] * merged['unit_cost']
    merged['stock_value'] = merged['us_value'] + merged['uk_value'] + merged['eu_value']

    print("🔄 [6/7] 마스터 계정 연동 및 그룹 정보 부착 중...")
    if master_df is not None:
        cols_to_use = ['품목코드']
        if '상품라인'   in master_df.columns: cols_to_use.append('상품라인')
        if '통합 제품명' in master_df.columns: cols_to_use.append('통합 제품명')

        master_mapping = master_df[cols_to_use].drop_duplicates(subset=['품목코드'])
        master_mapping['품목코드'] = master_mapping['품목코드'].astype(str).str.strip()

        merged['품목코드_str'] = merged['품목코드'].astype(str).str.strip()
        merged = pd.merge(merged, master_mapping,
                          left_on='품목코드_str', right_on='품목코드',
                          how='left', suffixes=('', '_master'))

        merged['상품라인']   = merged['상품라인'].fillna('기타(미등록)')   if '상품라인'   in merged.columns else '기타(미등록)'
        merged['통합 제품명'] = merged['통합 제품명'].fillna('미등록') if '통합 제품명' in merged.columns else '미등록'
    else:
        merged['상품라인']   = '기타(미등록)'
        merged['통합 제품명'] = '미등록'

    # 리뉴얼(v2): 마스터 미등록이라 이름이 없어도 US/UK/EU 재고(이동중 포함)가 있으면 표시하기 위해 따로 보관 (2026-11-19)
    _has_stock = sum(merged[f'{r}_{k}'] for r in ['us', 'uk', 'eu'] for k in ['stock', 'air', 'sea']) > 0
    unregistered_df = merged[(merged['제품명'].str.strip() == "") & _has_stock].copy()
    merged = merged[merged['제품명'].str.strip() != ""]

    print("🔄 [7/7] JSON 생성 및 HTML 발행 중...")

    tab1 = merged.copy()
    tab1['grand_total'] = tab1['us_stock'] + tab1['uk_stock'] + tab1['eu_stock'] + \
                          tab1['us_air'] + tab1['us_sea'] + \
                          tab1['uk_air'] + tab1['uk_sea'] + \
                          tab1['eu_air'] + tab1['eu_sea']
    # 재고+이동중이 0이어도 판매가 발생 중이거나 국내 재고가 있으면 표시 —
    # 결품 상태로 팔리는 SKU가 대시보드에서 아예 사라져 항공 긴급을 놓치는 문제 방지 (2026-07-15)
    tab1_sales = tab1['us_sales'] + tab1['uk_sales'] + tab1['eu_sales']
    stockout_selling = ((tab1['grand_total'] <= 0) & (tab1_sales > 0)).sum()
    # 신제품(_V0N)은 재고·판매가 0이어도 항상 표시 (2026-08-27)
    tab1_is_new = tab1['품목코드'].apply(is_new_product_code)
    tab1 = tab1[(tab1['grand_total'] > 0) | (tab1_sales > 0) | (tab1['kr_stock'] > 0) | tab1_is_new]
    if stockout_selling:
        print(f"   > 🚨 결품 판매중 SKU {stockout_selling}개 — 재고 0이지만 판매 발생, 대시보드에 표시됨")

    # PLC 자동판정(단종) 보정: 판매가 실제로 발생 중이면(합계 100개 이상) 소스 파일 표시와 무관하게 운영으로 간주
    apply_sales_plc_override(tab1)

    unmatched_codes = tab1.loc[tab1['unit_cost'] <= 0, '품목코드'].tolist()
    print(f"   > ⚠️ 원가 미매칭 SKU: {len(unmatched_codes)}개 (재고 금액 계산에서 제외됨)")
    if unmatched_codes:
        preview = ', '.join(unmatched_codes[:15])
        print(f"     예시: {preview}{' 외 다수' if len(unmatched_codes) > 15 else ''}")

    unmatched_pallet_codes = tab1.loc[tab1['pallet_units'] <= 0, '품목코드'].tolist()
    print(f"   > ⚠️ 팔렛트 환산표 미매칭 SKU: {len(unmatched_pallet_codes)}개 (부킹 관리에서 PLT 계산 제외됨)")
    if unmatched_pallet_codes:
        preview = ', '.join(unmatched_pallet_codes[:15])
        print(f"     예시: {preview}{' 외 다수' if len(unmatched_pallet_codes) > 15 else ''}")

    sort_cols = [c for c in ['상품라인', '통합 제품명'] if c in tab1.columns]
    if sort_cols:
        tab1 = tab1.sort_values(by=sort_cols).reset_index(drop=True)

    json_tab1 = build_json_rows(tab1.copy())

    sum_cols  = ['us_stock', 'uk_stock', 'eu_stock',
                 'us_sales', 'uk_sales', 'eu_sales',
                 'us_air', 'us_sea', 'uk_air', 'uk_sea', 'eu_air', 'eu_sea',
                 'us_value', 'uk_value', 'eu_value', 'stock_value', 'kr_stock']
    # 그룹 키는 '통합 제품명' 단독으로 사용한다.
    # 과거에는 ['상품라인', '통합 제품명']을 함께 키로 썼는데, 마스터 계정.xlsx에 같은 통합
    # 제품명인데 상품라인이 다르게(오타/공백) 입력된 그룹이 있으면 그 통합 제품명이 여러 개의
    # 파편 그룹으로 쪼개져 재고 합계·PLC 상태가 일부 SKU만 반영되는 문제가 있었음(예: GLMX179/GLMX191).
    group_key = ['통합 제품명']

    # PLC: 그룹 내 아무 코드나 '운영'이면 그룹 전체를 '운영'으로 본다 (마스터 계정 순서 기반 판정은
    # 마스터 데이터 정합성에 너무 민감해서 보류 — 2026-07-01 재논의 후 단순 규칙으로 되돌림)
    def agg_plc_status(values):
        vals = set(values)
        if '운영' in vals:
            return '운영'
        if '단종' in vals:
            return '단종'
        return '정보없음'

    def agg_last7(series):
        # 그룹 내 코드들의 최근 7일 리스트를 요소별 합산
        lists = [v for v in series if isinstance(v, list)]
        if not lists:
            return [0] * 7
        return [sum(vals) for vals in zip(*lists)]

    def agg_fcst(series):
        # 그룹 내 코드들의 월별 판매예정을 월/권역별로 합산
        out = {}
        for d in series:
            if not isinstance(d, dict):
                continue
            for mk, regs in d.items():
                o = out.setdefault(mk, {'us': 0, 'uk': 0, 'eu': 0, 'jp': 0})
                for r in ('us', 'uk', 'eu', 'jp'):
                    o[r] += regs.get(r, 0)
        return out

    merged_rep = merged.groupby(group_key, sort=False).agg(
        대표제품명=('제품명', 'first'),
        대표상품라인=('상품라인', 'first'),
        plc_status=('plc_status', agg_plc_status),
        us_spike=('us_spike', 'any'),
        uk_spike=('uk_spike', 'any'),
        eu_spike=('eu_spike', 'any'),
        is_heavy=('is_heavy', 'any'),
        us_last7=('us_last7', agg_last7),
        uk_last7=('uk_last7', agg_last7),
        eu_last7=('eu_last7', agg_last7),
        us_sales_window=('us_sales_window', 'min'),
        uk_sales_window=('uk_sales_window', 'min'),
        eu_sales_window=('eu_sales_window', 'min'),
        fcst=('fcst', agg_fcst),
    ).reset_index()
    merged_sum = merged.groupby(group_key, sort=False)[sum_cols].sum().reset_index()

    tab2 = pd.merge(merged_rep, merged_sum, on=group_key)

    # 그룹 포캐스팅도 마스터 매핑 기준으로 재계산 — merged에 없는 신코드(재고 미도착) 물량 포함
    group_fcst = build_group_fcst(master_df, fcst_map)
    if group_fcst:
        before_zero = sum(1 for v in tab2['fcst'] if not v)
        tab2['fcst'] = tab2['통합 제품명'].map(lambda g: group_fcst.get(str(g).strip(), {}))
        after_zero = sum(1 for v in tab2['fcst'] if not v)
        print(f"   > 그룹 포캐스팅: 마스터 매핑 기준 재집계 ({len(group_fcst)}개 그룹, 빈 그룹 {before_zero}→{after_zero})")

    # 그룹 일판매는 코드별 값 합산이 아니라 '일자별 합계'에서 재계산한다.
    # (코드별 채택 윈도우가 다르면 서로 다른 기간의 수요가 더해져 과대 계상됨 — build_group_sales 주석 참고)
    group_sales = build_group_sales(merged, {'us': us_daily, 'uk': uk_daily, 'eu': eu_daily})
    for r, gdf in group_sales.items():
        cols = [f'{r}_sales', f'{r}_sales_window', f'{r}_last7', f'{r}_spike']
        tab2 = tab2.drop(columns=[c for c in cols if c in tab2.columns])
        tab2 = pd.merge(tab2, gdf, on='통합 제품명', how='left')
        tab2[f'{r}_sales'] = tab2[f'{r}_sales'].fillna(0).astype(int)
        tab2[f'{r}_sales_window'] = tab2[f'{r}_sales_window'].fillna(28).astype(int)
        tab2[f'{r}_spike'] = tab2[f'{r}_spike'].apply(lambda v: bool(v) if pd.notna(v) else False)
        tab2[f'{r}_last7'] = tab2[f'{r}_last7'].apply(lambda v: v if isinstance(v, list) else [0]*7)

    tab2['품목코드'] = tab2['통합 제품명']
    tab2 = tab2.rename(columns={'대표제품명': '제품명', '대표상품라인': '상품라인'})

    # 통합 제품명(=대표코드)의 표시 이름은 '그 대표코드 자신의 제품명'을 쓴다.
    # 마스터의 '통합 제품명' 칸에는 항상 대표 품목코드가 들어있고(189개 전부 품목코드와 일치),
    # 그룹 표시 이름은 그 코드의 제품명이어야 한다는 규칙. groupby('제품명','first')는 merged(재고)
    # 순서상 먼저 온 코드의 제품명을 잡아, 대표코드와 다른 형제 SKU 이름이 뜨는 문제가 있었음
    # (예: GLMX179 그룹에 GLMX091 '…시멘리트…' 이름이 잘못 표시됨).
    if master_df is not None and '제품명' in master_df.columns and '품목코드' in master_df.columns:
        code2name = {}
        for _, mr in master_df.iterrows():
            code = str(mr.get('품목코드', '')).strip()
            name = str(mr.get('제품명', '')).strip()
            if code and name and name.lower() != 'nan' and code not in code2name:
                code2name[code] = name
        tab2['제품명'] = [
            code2name.get(str(g).strip()) or nm
            for g, nm in zip(tab2['통합 제품명'], tab2['제품명'])
        ]

    tab2['grand_total'] = tab2['us_stock'] + tab2['uk_stock'] + tab2['eu_stock'] + \
                          tab2['us_air'] + tab2['us_sea'] + \
                          tab2['uk_air'] + tab2['uk_sea'] + \
                          tab2['eu_air'] + tab2['eu_sea']
    tab2_sales = tab2['us_sales'] + tab2['uk_sales'] + tab2['eu_sales']
    # 신제품(_V0N) 멤버 코드를 가진 통합 제품명 그룹은 재고·판매 0이어도 항상 표시 (2026-08-27)
    new_group_names = set(
        merged.loc[merged['품목코드'].apply(is_new_product_code), '통합 제품명']
              .astype(str).str.strip()
    )
    tab2_is_new = tab2['통합 제품명'].astype(str).str.strip().isin(new_group_names) \
                  | tab2['품목코드'].apply(is_new_product_code)
    tab2 = tab2[(tab2['grand_total'] > 0) | (tab2_sales > 0) | (tab2['kr_stock'] > 0) | tab2_is_new]

    apply_sales_plc_override(tab2)

    if sort_cols:
        tab2 = tab2.sort_values(by=sort_cols).reset_index(drop=True)

    json_tab2 = build_json_rows(tab2.copy())

    # 리뉴얼(v2)용 층별화 등급 + 카테고리 — 실패해도 기존 대시보드 빌드는 계속
    grade_meta = {}
    json_tab1_v2, json_tab2_v2 = json_tab1, json_tab2
    if os.path.exists("template_v2.html"):
        print("🔄 [7.5/7] (리뉴얼) 층별화 등급·카테고리 계산 중 (직전 3개월 판매)...")
        try:
            # v2는 행을 추가(새 권역 재고만 있는 코드)하므로 복사본 사용 — 기존 대시보드·슬랙·dashboard_data.json은 원본 그대로
            json_tab1_v2, json_tab2_v2 = copy.deepcopy(json_tab1), copy.deepcopy(json_tab2)
            unreg_rows = build_unregistered_rows(unregistered_df)
            grade_meta = enrich_v2_fields(
                json_tab1_v2, json_tab2_v2, master_df,
                {'us': excel_us_sales, 'uk': excel_uk_sales, 'eu': excel_eu_sales, 'jp': excel_jp_sales},
                excluded_codes, unreg_rows=unreg_rows, jp_sales_df=jp_sales_df, jp_daily=jp_daily)
            print(f"   > 기간 {grade_meta['months'][0]} ~ {grade_meta['months'][-1]}, "
                  f"전제품 판매 {grade_meta['total_qty']:,}개 | 개별코드 {grade_meta['count_tab1']} | 통합제품명 {grade_meta['count_tab2']}")
            if grade_meta.get('category_missing'):
                miss = grade_meta['category_missing']
                print(f"   > ℹ️ 카테고리 파일에 없는 코드 {len(miss)}개 → 자동분류: {', '.join(miss[:15])}{' 외' if len(miss) > 15 else ''}")
            for r, src in grade_meta['sources'].items():
                print(f"   > {r.upper()} 층별화 수량: {', '.join(src)} (합계 {grade_meta['region_totals'][r]:,})")
            for r, miss in grade_meta['months_missing'].items():
                if miss:
                    print(f"   > ⚠️ {r.upper()} 판매 파일에 {', '.join(miss)} 시트가 없어 층별화 합계에서 빠짐")
        except Exception as e:
            print(f"   > ⚠️ 층별화 계산 실패(기존 대시보드에는 영향 없음): {e}")

    # 워터폴 데이터 (현지 ETA)
    waterfall_map = {}
    if '통합 제품명' in merged.columns:
        for unified_name, group in merged.groupby('통합 제품명'):
            if not unified_name or unified_name == '미등록':
                continue
            schedules = []
            for sku in group['품목코드'].astype(str).str.strip().str.upper().unique():
                if sku in eta_schedule:
                    schedules.extend(eta_schedule[sku])
            if schedules:
                waterfall_map[unified_name] = schedules
    print(f"   > 워터폴 데이터: {len(waterfall_map)}개 통합 제품명에 입고 일정 매핑")

    # 국내 입고 데이터 (통합 제품명별 합산)
    inbound_map = {}
    if '통합 제품명' in merged.columns:
        for unified_name, group in merged.groupby('통합 제품명'):
            if not unified_name or unified_name == '미등록':
                continue
            agg_total = 0
            agg = {}
            for sku in group['품목코드'].astype(str).str.strip().str.upper().unique():
                if sku not in domestic_inbound:
                    continue
                entry = domestic_inbound[sku]
                agg_total += entry['total']
                for it in entry['items']:
                    key = (it['month'], it['week'])
                    agg[key] = agg.get(key, 0) + it['qty']
            if agg_total > 0 or agg:
                items = [{'month': k[0], 'week': k[1], 'qty': v} for k, v in agg.items()]
                inbound_map[unified_name] = {'total': agg_total, 'items': items}
    print(f"   > 국내 입고 예정: {len(inbound_map)}개 통합 제품명에 매핑 완료")

    # 일별 판매 추이 (통합 제품명별, 최근 28일) — 모달의 '일별 판매' 탭용
    today = datetime.now().date()
    win_dates = [today - timedelta(days=28) + timedelta(days=i) for i in range(28)]
    region_daily = {'us': us_daily, 'uk': uk_daily, 'eu': eu_daily, 'jp': jp_daily or {}}
    daily_sales_map = {}
    if '통합 제품명' in merged.columns:
        for unified_name, group in merged.groupby('통합 제품명'):
            if not unified_name or unified_name == '미등록':
                continue
            skus = group['품목코드'].astype(str).str.strip().str.upper().unique()
            entry = {}
            has_any = False
            for r, dmap in region_daily.items():
                series = [0] * 28
                for sku in skus:
                    d = dmap.get(sku)
                    if not d:
                        continue
                    for i, dt in enumerate(win_dates):
                        q = d.get(dt, 0)
                        if q:
                            series[i] += q
                if any(series):
                    has_any = True
                entry[r] = series
            if has_any:
                entry['dates'] = [d.isoformat() for d in win_dates]
                daily_sales_map[unified_name] = entry
    print(f"   > 일별 판매 추이: {len(daily_sales_map)}개 통합 제품명에 매핑 완료")

    cost_meta = {
        "matched_count": int((tab1['unit_cost'] > 0).sum()),
        "unmatched_count": len(unmatched_codes),
        "unmatched_codes": unmatched_codes
    }

    pallet_meta = {
        "matched_count": int((tab1['pallet_units'] > 0).sum()),
        "unmatched_count": len(unmatched_pallet_codes),
        "unmatched_codes": unmatched_pallet_codes,
        "b2b_file": b2b_path,
        "b2b_matched_count": len(b2b_map)
    }

    if not os.path.exists(template_path):
        print("❌ template.html 파일이 없습니다.")
        return

    # 스냅샷 동결 저장 + 최근 30일 델타 비교 이력 (기존·리뉴얼 대시보드 공통, 1회만 저장)
    snapshot_history = update_snapshots(json_tab1, json_tab2)

    # 리뉴얼 워터폴: 지난달 1일 ~ 어제 주차별 실적(판매·seeding·당시 재고) — 실패해도 빌드 계속
    if grade_meta:
        try:
            grade_meta['past'] = attach_past_weeks(json_tab1_v2, json_tab2_v2, master_df, region_daily)
        except Exception as e:
            print(f"   > ⚠️ 지난 실적 계산 실패(빌드에는 영향 없음): {e}")

    # 플레이스홀더 → 주입값. 기존(template.html)·리뉴얼(template_v2.html)이 같은 데이터를 쓴다.
    # (템플릿에 없는 플레이스홀더는 replace가 아무 일도 안 하므로 기존 대시보드엔 영향 없음)
    replacements = [
        ("__DATA_FROM_PYTHON__", json.dumps(json_tab1, ensure_ascii=False)),
        ("__GROUPED_DATA_FROM_PYTHON__", json.dumps(json_tab2, ensure_ascii=False)),
        ("__WATERFALL_DATA_FROM_PYTHON__", json.dumps(waterfall_map, ensure_ascii=False)),
        ("__INBOUND_DATA_FROM_PYTHON__", json.dumps(inbound_map, ensure_ascii=False)),
        # 부킹 관리용: SKU(개별 코드) 단위 국내 입고 예정 (그룹 단위 inbound_map을 쓰면 그룹 내 코드마다 중복 계산됨)
        ("__INBOUND_SKU_DATA_FROM_PYTHON__", json.dumps(domestic_inbound, ensure_ascii=False)),
        ("__LAST_UPDATED_FROM_PYTHON__", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("__COST_META_FROM_PYTHON__", json.dumps(cost_meta, ensure_ascii=False)),
        ("__DAILY_SALES_FROM_PYTHON__", json.dumps(daily_sales_map, ensure_ascii=False)),
        ("__PALLET_META_FROM_PYTHON__", json.dumps(pallet_meta, ensure_ascii=False)),
        ("__SNAPSHOT_HISTORY_FROM_PYTHON__", json.dumps(snapshot_history, ensure_ascii=False)),
        ("__GRADE_META_FROM_PYTHON__", json.dumps(grade_meta, ensure_ascii=False)),
    ]

    def render(tpl_path, out_path, overrides=None):
        with open(tpl_path, 'r', encoding='utf-8') as f:
            content = f.read()
        for key, val in (overrides or []) + replacements:
            content = content.replace(key, val)
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(content)

    # 2026-11-20부터 리뉴얼(template_v2.html)이 정식 대시보드: dashboard.html(온라인 배포본)로 생성.
    # 기존 화면은 dashboard_legacy.html로 로컬에만 남김(비교용, git에 올리지 않음).
    # 리뉴얼 생성이 실패하면 기존 화면을 dashboard.html로 써서 배포가 멈추지 않게 함.
    if os.path.exists("template_v2.html"):
        render(template_path, "dashboard_legacy.html")
        try:
            render("template_v2.html", output_path, overrides=[
                ("__DATA_FROM_PYTHON__", json.dumps(json_tab1_v2, ensure_ascii=False)),
                ("__GROUPED_DATA_FROM_PYTHON__", json.dumps(json_tab2_v2, ensure_ascii=False)),
            ])
            print(f"🎉 리뉴얼 대시보드 → '{output_path}' 생성 완료! (기존 화면: dashboard_legacy.html)")
        except Exception as e:
            print(f"   > ⚠️ 리뉴얼 생성 실패 — 기존 화면으로 '{output_path}' 생성: {e}")
            render(template_path, output_path)
    else:
        render(template_path, output_path)

    # S&OP 리포트(sop_report.py)가 대시보드 HTML을 다시 파싱하지 않도록 데이터만 별도 저장
    dashboard_data_path = "dashboard_data.json"
    with open(dashboard_data_path, 'w', encoding='utf-8') as f:
        json.dump({
            "generated_at": datetime.now().isoformat(),
            "grouped_rows": json_tab2,
            # v9(판매량집계_포캐스팅비교)가 발주/예측에서 단종 SKU를 거르는 데 사용 —
            # 취소선 기반 PLC 판정 + 판매량 보정(apply_sales_plc_override)이 적용된 개별 코드 단위 상태
            "plc_by_code": {str(r['품목코드']).strip().upper(): str(r['plc_status'])
                            for _, r in tab1.iterrows() if str(r.get('plc_status', '')) in ('운영', '단종')}
        }, f, ensure_ascii=False)

    print(f"🎉 성공: '{output_path}' 생성 완료!")

    # 재고 부족 Slack 메시지 준비 (여기서는 저장만 — 발송은 update.bat 최종 업로드 후 1회)
    try:
        slack_domestic = parse_slack_domestic_stock(master_df)
        if slack_domestic:
            print(f"   > 국내창고(슬랙 알림용) 재고: {len(slack_domestic)}개 코드, 합계 {sum(slack_domestic.values()):,}")
        else:
            print("   > ℹ️ 국내재고/ 폴더에 파일 없음 — 국내재고 차감 없이 계산")
        save_slack_message(json_tab1, slack_domestic)
    except Exception as e:
        print(f"   > ⚠️ Slack 메시지 준비 실패(빌드에는 영향 없음): {e}")


# ─────────────────────────────────────────────────────────────
# 재고 부족 Slack 알림
# ─────────────────────────────────────────────────────────────
# 기준(대시보드와 동일): 권역별로 ① 항공긴급(회전일 ≤ US 45 / UK·EU 60, 중량물 +30)
#                         또는 ② 안전재고 미달(총재고 < 안전재고) 이면 "부족"으로 집계.
# 집계 단위 = 개별 품목코드(tab1). Webhook URL은 환경변수 SCM_SLACK_WEBHOOK 에서 읽음.
SLACK_REGION_META = [
    ('us', '미국', 'US', 45),
    ('uk', '영국', 'UK', 60),
    ('eu', '유럽', 'EU', 60),
]
SLACK_MAX_ROWS = 60  # 표에 나열할 최대 행 수 (권역 합산)

def _slack_days_val(v):
    if v is None or v == '여유':
        return float('inf')
    try:
        return int(v)
    except (ValueError, TypeError):
        return float('inf')

def _is_sunscreen(name):
    n = str(name).lower()
    return ('선크림' in n) or ('썬크림' in n) or ('sun' in n)

# 국내창고 재고원장에서 슬랙 알림에 포함할 창고코드 (용인메인/서브·안성죽산 등)
DOMESTIC_WH_CODES = {'WH507', 'WH508', 'WH611', 'WH612'}

def _detect_kctc_layout(df):
    """KCTC 엑셀의 헤더를 인식해 (헤더행, 코드열, 합산할 수량열 리스트, 양식명) 반환. 아니면 None.
    - 구양식(~2026-10): 'SKU' + '재고현황'(그 아래 '합계' = 정상+보류계열 전체) → 재고현황 열 1개.
    - 신양식(2026-10-22~, 파일명 '엑셀 저장_*'): 'SKU 번호' + '정상재고'(합계/가용/신규주문) +
      '보류재고'(보류/불량/수선/예약/파손/반품보류/유통기한만료/파손(셀러귀책)) 2단 헤더.
      정상재고 그룹의 '가용' 열만 사용 (신규주문·보류재고 제외 — 구양식 '합계'와 정의가 다름에 주의)."""
    for r in range(min(8, len(df) - 1)):
        vals = [str(v).strip() if pd.notna(v) else '' for v in df.iloc[r].tolist()]
        if 'SKU 번호' in vals and '정상재고' in vals and '보류재고' in vals:
            sub = [str(v).strip() if pd.notna(v) else '' for v in df.iloc[r + 1].tolist()]
            code_col = vals.index('SKU 번호')
            normal_start = vals.index('정상재고')
            hold_start = vals.index('보류재고')
            # 정상재고 그룹 안의 '가용' 서브헤더 열만 사용 (2026-10-22 사용자 지시: 신규주문·보류재고 제외)
            normal_avail = next((i for i in range(normal_start, hold_start) if sub[i] == '가용'), None)
            if normal_avail is None:
                return None
            return r, code_col, [normal_avail], '신양식'
        if 'SKU' in vals and any('재고현황' in v for v in vals):
            return r, vals.index('SKU'), [next(i for i, v in enumerate(vals) if '재고현황' in v)], '구양식'
    return None


def find_kctc_files(folder):
    """국내재고/ 폴더에서 KCTC 재고 파일 후보 [(path, df, layout)]를 반환.
    파일명에 'KCTC'가 있거나(구양식), 헤더 구조가 KCTC 양식이면(신양식 '엑셀 저장_*'처럼 이름에 KCTC 없음) 인정.
    재고원장 파일은 제외(구조가 다르고 크기가 커서 읽지 않음)."""
    found = []
    for f in (glob.glob(os.path.join(folder, '*.xls')) + glob.glob(os.path.join(folder, '*.xlsx'))):
        name = os.path.basename(f)
        if name.startswith('~$') or '재고원장' in name:
            continue
        try:
            df = pd.read_excel(f, header=None)
        except Exception as e:
            if 'KCTC' in name.upper():
                print(f"   > ⚠️ KCTC 파일 읽기 실패({name}): {e}")
            continue
        layout = _detect_kctc_layout(df)
        if layout is not None:
            found.append((f, df, layout))
        elif 'KCTC' in name.upper():
            # 예전엔 고정 열(E/L)로 폴백했으나, 양식이 바뀌면 조용히 엉뚱한 숫자를 읽으므로 건너뛰고 경고만 남김
            print(f"   > ⚠️ KCTC 헤더 인식 실패 — 이 파일은 건너뜀(양식 확인 필요): {name}")
    return found


def parse_kctc_stock(folder, master_codes):
    """국내재고/ 폴더의 KCTC 파일(3PL 물류센터 재고)에서 코드별 재고 합산.
    - 구양식/신양식 모두 지원 (_detect_kctc_layout). 수량 = 구양식 합계(정상+보류계열) / 신양식 정상재고 가용.
    - 구양식은 같은 SKU가 제조일/유통기한별로 여러 행이라 코드별로 합산한다. 신양식 맨 아래 '합계' 행은 코드가 비어 자동 스킵.
    - 재고원장(WH507/508/611/612)과 다른 창고라 합산해도 중복 아님.
    - KCTC 파일이 여러 벌이면 mtime 최신 1개만 사용(재다운로드 복사본·구/신양식 공존 시 중복합산 방지)."""
    candidates = find_kctc_files(folder)
    if not candidates:
        return {}
    f, df, (hr, code_col, qty_cols, kind) = max(candidates, key=lambda x: os.path.getmtime(x[0]))
    if len(candidates) > 1:
        print(f"   > ℹ️ KCTC 파일 {len(candidates)}개 중 최신본만 사용: {os.path.basename(f)}")
    stock = {}
    for r in range(hr + 1, len(df)):  # 서브헤더 행(합계/정상/보류…)은 코드 빈칸이라 자동 스킵
        row = df.iloc[r].tolist()
        code = row[code_col] if code_col < len(row) else None
        if pd.isna(code):
            continue
        code = str(code).strip().upper()
        if not code or code in ('SKU', 'SKU 번호', 'NAN'):
            continue
        qty = sum(parse_int(row[c]) for c in qty_cols if c < len(row))
        for c in (code.split(',') if ',' in code else [code]):
            c = c.strip()
            if c and (not master_codes or c in master_codes):
                stock[c] = stock.get(c, 0) + qty
    print(f"   > KCTC 재고({kind}): {len(stock)}개 코드, 합계 {sum(stock.values()):,} ({os.path.basename(f)})")
    return stock


def parse_slack_domestic_stock(master_df, folder=None):
    """슬랙 부족알림 전용 국내창고 재고.
    `국내재고/` 폴더의 엑셀에서 헤더를 인식해 '코드'=품목코드, '가용'=가용재고를 읽어 코드별 합산.
    - '창고코드' 컬럼이 있으면(재고원장) → DOMESTIC_WH_CODES(WH507/508/611/612) 창고만 합산.
    - '창고코드'가 없으면(과거 센터별 파일) → 전 행 합산.
    - KCTC 파일(3PL 물류센터)은 구조가 달라 parse_kctc_stock으로 별도 파싱 후 합산 (구양식 SKU/재고현황, 신양식 SKU 번호/정상재고/보류재고).
    마스터 계정에 있는 코드만 유지(그 외 제외). 대시보드/부킹의 콜로세움 kr_stock과는 별개."""
    if folder is None:
        folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), '국내재고')
    if not os.path.isdir(folder):
        return {}
    master_codes = set()
    if master_df is not None and '품목코드' in master_df.columns:
        for v in master_df['품목코드'].dropna().astype(str):
            for c in (v.split(',') if ',' in v else [v]):
                c = c.strip().upper()
                if c and c != 'NAN':
                    master_codes.add(c)
    stock = {}
    all_files = [f for f in sorted(glob.glob(os.path.join(folder, '*.xls')) +
                                   glob.glob(os.path.join(folder, '*.xlsx')))
                 if not os.path.basename(f).startswith('~$')
                 and 'KCTC' not in os.path.basename(f).upper()]  # KCTC는 별도 파서로 처리
    # 재고원장 파일이 있으면 '최신 1개'만 사용 (옛 파일이 쌓여도 중복합산 방지).
    # 없으면 과거 센터별 파일들을 합산(하위호환).
    ledger_files = [f for f in all_files if '재고원장' in os.path.basename(f)]
    if ledger_files:
        files = [max(ledger_files, key=os.path.getmtime)]
        print(f"   > 국내재고 소스: 재고원장 최신본 사용 → {os.path.basename(files[0])}")
    else:
        # 이름에 KCTC가 없는 신양식 KCTC 파일('엑셀 저장_*')도 센터 파일로 오인하지 않게 제외
        kctc_paths = {p for p, _, _ in find_kctc_files(folder)}
        files = [f for f in all_files if '재고원장' not in os.path.basename(f) and f not in kctc_paths]
    for f in files:
        try:
            df = pd.read_excel(f, header=None)
        except Exception as e:
            print(f"   > ⚠️ 국내재고 파일 읽기 실패({os.path.basename(f)}): {e}")
            continue
        # 헤더행: '코드' 셀이 있는 행 (재고원장/센터파일 공통)
        hr = header = None
        for r in range(min(10, len(df))):
            vals = [str(v).strip() for v in df.iloc[r].tolist()]
            if '코드' in vals:
                hr, header = r, vals
                break
        if hr is None:
            print(f"   > ⚠️ 국내재고 파일 헤더('코드') 못 찾음: {os.path.basename(f)}")
            continue
        code_col = header.index('코드')
        qty_col = next((i for i, h in enumerate(header) if h == '가용'), None)
        if qty_col is None:
            qty_col = next((i for i, h in enumerate(header) if '가용' in h), None)
        wh_col = next((i for i, h in enumerate(header) if h == '창고코드'), None)
        if qty_col is None:
            print(f"   > ⚠️ 국내재고 파일 '가용' 컬럼 못 찾음: {os.path.basename(f)}")
            continue
        for r in range(hr + 1, len(df)):
            row = df.iloc[r].tolist()
            if wh_col is not None:  # 재고원장: 지정 창고코드만
                wc = row[wh_col] if wh_col < len(row) else None
                if pd.isna(wc) or str(wc).strip() not in DOMESTIC_WH_CODES:
                    continue
            code = row[code_col] if code_col < len(row) else None
            if pd.isna(code):
                continue
            code = str(code).strip().upper()
            if not code or code == '코드':
                continue
            qty = parse_int(row[qty_col])
            for c in (code.split(',') if ',' in code else [code]):
                c = c.strip()
                if c and (not master_codes or c in master_codes):
                    stock[c] = stock.get(c, 0) + qty
    # KCTC(3PL 물류센터) 재고 합산 — 재고원장과 다른 창고라 중복 아님
    for c, q in parse_kctc_stock(folder, master_codes).items():
        stock[c] = stock.get(c, 0) + q
    return stock

def build_shortage_report(json_tab1, domestic_map=None):
    """tab1 JSON에서 부족 SKU를 '코드 단위로 국가 합산'해 단일 리스트로 반환.
    - 부족재고 수량 = 부족한 권역들의 필요량 합계 − 국내창고 재고
    - 필요시점 = 부족 권역 중 가장 급한 회전일(최소)
    - 사유(항공긴급 여부) = 부족 권역 중 하나라도 항공긴급이면 항공긴급
    - domestic_map: 국내창고(용인메인/서브/안성죽산 합산) 코드별 재고. 주면 이 값을 차감,
      없으면 대시보드 kr_stock(콜로세움)으로 폴백.
    MX 코드는 선크림(MX00086/87 등)만 포함하고 나머지는 제외."""
    items = []
    for row in json_tab1:
        code = str(row.get('code', ''))
        name = str(row.get('name', ''))
        # MX 코드는 선크림만 알림에 포함 (2026-08-28 사용자 요청)
        if code.upper().startswith('MX') and not _is_sunscreen(name):
            continue
        need_sum = 0
        min_days = float('inf')
        any_air = False
        is_short = False
        for region, _label, _short, base_thr in SLACK_REGION_META:
            days = _slack_days_val(row.get(f'{region}_days'))
            extra = 30 if row.get('is_heavy') else 0
            is_air = days <= base_thr + extra
            total = row.get(f'{region}_total', 0) or 0
            safety = row.get(f'{region}_safety', 0) or 0
            is_danger = safety > 0 and total < safety
            if not (is_air or is_danger):
                continue
            is_short = True
            need_sum += max(0, int(safety) - int(total)) if safety > 0 else 0
            min_days = min(min_days, days)
            any_air = any_air or is_air
        if not is_short:
            continue
        # 국내창고 재고를 코드 단위로 1회 차감 → 추가 발주 필요 순부족분
        if domestic_map is not None:
            kr_stock = int(domestic_map.get(str(code).strip().upper(), 0))
        else:
            kr_stock = int(row.get('kr_stock', 0) or 0)
        need_after_kr = max(0, need_sum - kr_stock)
        items.append({
            'code': code,
            'name': name,
            'mfr': str(row.get('mfr', '') or '').strip() or '(제조사 미상)',
            'days': min_days,
            'need': need_after_kr,        # 한국재고 차감 후 순부족(발주 필요분)
            'need_overseas': need_sum,    # 차감 전 해외 부족합
            'kr_stock': kr_stock,
            'covered': need_after_kr == 0,  # 한국재고로 전량 커버(발송만 하면 됨)
            'is_air': any_air,
            'plc': row.get('plc_status', ''),
        })
    return items

SLACK_FLAG = {'US': ':flag-us:', 'UK': ':flag-gb:', 'EU': ':flag-eu:'}

def _group_by_mfr(section_items):
    """제조사별로 묶어 (제조사, [항목...]) 리스트 반환.
    그룹 내 상품은 부족재고 수량 내림차순, 제조사 그룹은 그룹 최대 부족량 내림차순."""
    groups = {}
    for it in section_items:
        groups.setdefault(it['mfr'], []).append(it)
    for its in groups.values():
        its.sort(key=lambda x: x['need'], reverse=True)
    return sorted(groups.items(), key=lambda kv: max(x['need'] for x in kv[1]), reverse=True)

def _months_needed(days):
    """회전일(재고 소진까지 남은 일수)을 '몇 개월 안에 필요'로 환산."""
    if days == float('inf'):
        return "-"
    if days <= 0:
        return "즉시 필요"
    return f"{(int(days) + 29) // 30}개월 내 필요"

def format_shortage_message(items):
    """부족 리스트를 사유별 2개 섹션(① 항공긴급 ② 안전재고 미달)으로, 각 섹션 안은 제조사별 하위목록으로 변환.
    각 줄 = 상품코드 · 제품명 · 부족재고 수량 · 몇 개월 안에 필요. (국가 구분 없음) 부족 없으면 None."""
    if not items:
        return None
    air = sorted([it for it in items if it['is_air']], key=lambda x: x['days'])
    safe = sorted([it for it in items if not it['is_air']], key=lambda x: x['days'])
    today = datetime.now().strftime('%Y-%m-%d %H:%M')
    lines = [
        f":package: *재고 부족 알림* — {today} 빌드",
        f"항공긴급 {len(air)}건 · 안전재고 미달 {len(safe)}건 · _제조사별 · 상품코드·제품명·부족수량(해외+국내재고 차감후)·필요시점_ · MX코드는 선크림만 포함",
    ]

    def render(it):
        tag = " _(단종)_" if it['plc'] == '단종' else ""
        if it['covered']:
            qty = f"부족 *0* (국내재고 {it['kr_stock']:,}개로 커버 — 발송만 필요)"
        else:
            qty = f"부족 *{it['need']:,}*"
        return f"    `{it['code']}` {it['name']} · {qty} · {_months_needed(it['days'])}{tag}"

    def render_section(title, section_items):
        lines.append(f"\n{title}")
        if not section_items:
            lines.append("_해당 없음_")
            return
        for mfr, its in _group_by_mfr(section_items):
            lines.append(f"  ▸ *{mfr}* ({len(its)}건)")
            for it in its:
                lines.append(render(it))

    render_section(":airplane: *1. 항공긴급 (매우 긴급)*", air)
    render_section(":rotating_light: *2. 안전재고 미달*", safe)

    return "\n".join(lines)

SLACK_MSG_FILE = 'slack_message.txt'      # app.py 빌드 때 렌더한 부족 메시지를 여기에 저장
SLACK_PAUSE_FILE = 'slack_paused.txt'     # 존재하면 발송 정지
SLACK_MARKER_FILE = 'slack_last_sent.txt' # 마지막 발송 날짜 (하루 1회 가드)

def save_slack_message(json_tab1, domestic_map=None):
    """부족재고 메시지를 렌더해 파일로 저장(발송하지 않음).
    실제 발송은 update.bat의 최종 업로드(git push) 성공 후 `python app.py --send-slack`에서 1회만 수행."""
    items = build_shortage_report(json_tab1, domestic_map)
    message = format_shortage_message(items)
    if message is None:
        message = ":white_check_mark: *재고 부족 알림* — 항공긴급/안전재고 미달 SKU 없음"
    base_dir = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(base_dir, SLACK_MSG_FILE), 'w', encoding='utf-8') as f:
        f.write(message)
    print(f"   > 📝 Slack 부족 메시지 준비 완료 (부족 {len(items)}건) — 발송은 최종 업로드 후 1회")

def send_saved_slack_message():
    """저장된 부족 메시지(slack_message.txt)를 Slack으로 발송.
    - update.bat이 git push 성공 후 `python app.py --send-slack`로 호출 → '최종 업로드에만 1회' 발송
    - 하루 1회 가드(slack_last_sent.txt) + 일시정지(slack_paused.txt) + 웹훅 미설정 시 자동 스킵."""
    import urllib.request
    webhook = os.environ.get('SCM_SLACK_WEBHOOK', '').strip()
    if not webhook:
        print("ℹ️ Slack 웹훅(SCM_SLACK_WEBHOOK) 미설정 — 알림 건너뜀")
        return
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.exists(os.path.join(base_dir, SLACK_PAUSE_FILE)):
        print("⏸️ Slack 알림 일시정지 상태(slack_paused.txt) — 발송 건너뜀")
        return
    marker = os.path.join(base_dir, SLACK_MARKER_FILE)
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        if os.path.exists(marker):
            with open(marker, encoding='utf-8') as f:
                if f.read().strip() == today:
                    print("ℹ️ 오늘 이미 Slack 알림을 보냈습니다 — 재발송 생략 (하루 1회)")
                    return
    except Exception:
        pass
    msg_path = os.path.join(base_dir, SLACK_MSG_FILE)
    if not os.path.exists(msg_path):
        print("⚠️ 발송할 메시지(slack_message.txt)가 없습니다 — 먼저 app.py 빌드가 필요합니다")
        return
    with open(msg_path, encoding='utf-8') as f:
        message = f.read()
    payload = json.dumps({"text": message}).encode('utf-8')
    req = urllib.request.Request(webhook, data=payload,
                                 headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=15) as resp:
        resp.read()
    try:
        with open(marker, 'w', encoding='utf-8') as f:
            f.write(today)
    except Exception:
        pass
    print("✅ Slack 재고부족 알림 전송 완료 (최종 업로드 후 1회)")


if __name__ == "__main__":
    import sys
    if '--send-slack' in sys.argv:
        # update.bat이 git push 성공 후 호출 — 저장된 메시지를 1회 발송
        send_saved_slack_message()
    else:
        run_dashboard_generator()