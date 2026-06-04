import sys
import pandas as pd
from collections import Counter
from openpyxl.styles import PatternFill


# キーワード部分一致でバッジ基準を決定（L-caféプレフィックス等に対応）
# 上から順に最初に一致したルールを使用
BADGE_CRITERIA_RULES: list[tuple[list[str], dict]] = [
    (['jcpp', 'learner'],  {'Bronze': 10, 'Silver': 20, 'Gold': 30}),
    (['jcpp', 'provider'], {'Bronze': 15, 'Silver': 30, 'Gold': 45}),
]

DEFAULT_CRITERIA = {'Bronze': 15, 'Silver': 30, 'Gold': 45}

# バッジセル背景色（ARGB形式）
BADGE_FILLS = {
    'Gold':   PatternFill(fill_type='solid', fgColor='FFD700'),
    'Silver': PatternFill(fill_type='solid', fgColor='C0C0C0'),
    'Bronze': PatternFill(fill_type='solid', fgColor='CD7F32'),
}


def _norm(s: str) -> str:
    """スペース・全角スペースを正規化して比較用文字列を返す"""
    import re
    s = re.sub(r'[\s　]+', ' ', str(s)).strip()
    s = re.sub(r'\s+（', '（', s)   # （の直前のスペースを除去
    s = re.sub(r'）\s+', '）', s)   # ）の直後のスペースを除去
    return s


def _norm_sid(s: str) -> str:
    """学生番号を正規化（大文字小文字を統一）"""
    return _norm(s).upper()


def _find_col(df: pd.DataFrame, *keywords: str) -> str | None:
    kws = [k.lower() for k in keywords]
    for col in df.columns:
        c = col.lower()
        if all(k in c for k in kws):
            return col
    return None


def determine_badge(count: int, jcpp_type: str) -> str:
    jcpp_lower = str(jcpp_type).lower()
    criteria = DEFAULT_CRITERIA
    for keywords, crit in BADGE_CRITERIA_RULES:
        if all(k in jcpp_lower for k in keywords):
            criteria = crit
            break
    if count >= criteria['Gold']:   return 'Gold'
    if count >= criteria['Silver']: return 'Silver'
    if count >= criteria['Bronze']: return 'Bronze'
    return ''


def process(input_path: str, output_path: str, log=print) -> None:
    log('読み込み中...')
    xl = pd.ExcelFile(input_path)

    # ---- participation_ALL_clean シート読み込み ----
    target_sheets = ['participation_ALL_clean']
    available = [s for s in target_sheets if s in xl.sheet_names]
    if not available:
        raise ValueError(f'participation_ALL_cleanシートが見つかりません。シート一覧: {xl.sheet_names}')
    log(f'読み込みシート: {available}')
    df = pd.concat([xl.parse(s, header=0) for s in available], ignore_index=True)
    total_input = len(df)
    log(f'入力データ合計: {total_input} 行')

    # ---- OpenBadge_all_clean から登録日・メールを取得 ----
    # キー: (氏名, プログラム名)  登録日はタイムスタンプの最古値
    reg_lookup:   dict = {}   # (学生番号, prog) -> オープンバッジ登録日
    email_lookup: dict = {}   # (学生番号, prog) -> email

    reg_sheet = 'OpenBadge_all_clean'
    if reg_sheet in xl.sheet_names:
        df_reg  = xl.parse(reg_sheet, header=0)
        r_sid   = _find_col(df_reg, '学生番号') or _find_col(df_reg, 'student')
        r_prog  = _find_col(df_reg, 'プログラム')
        r_date  = _find_col(df_reg, 'オープンバッジ登録日') or _find_col(df_reg, '登録日')
        r_email = _find_col(df_reg, 'email') or _find_col(df_reg, 'メール')

        if r_sid and r_prog and r_date:
            df_reg[r_date] = pd.to_datetime(df_reg[r_date], errors='coerce')
            for _, row in df_reg.iterrows():
                sid  = _norm_sid(row[r_sid])
                prog = _norm(row[r_prog])
                date = row[r_date]
                key  = (sid, prog)
                # 最古の登録日を使用
                if pd.notna(date) and (key not in reg_lookup or date < reg_lookup[key]):
                    reg_lookup[key] = date
                # メールアドレス
                if r_email and pd.notna(row.get(r_email)) and key not in email_lookup:
                    email_lookup[key] = row[r_email]
            log(f'OpenBadge_all_clean: {len(reg_lookup)} 件の登録日を取得（学生番号＋プログラム名で照合）')
        else:
            log(f'警告: OpenBadge_all_cleanシートのカラムを特定できませんでした（列: {list(df_reg.columns)}）')
    else:
        log(f'警告: {reg_sheet} シートが見つかりません。登録日・メールは空欄になります。')
        log(f'  利用可能シート: {xl.sheet_names}')

    # ---- カラム検出 ----
    col_sid    = _find_col(df, '学生番号') or _find_col(df, '学籍番号')
    col_name   = _find_col(df, '氏名') or _find_col(df, 'name')
    col_date   = _find_col(df, '参加日')
    col_filter = _find_col(df, '1のみ') or _find_col(df, '結果統合', '計算')
    col_status = _find_col(df, 'badge_status') or _find_col(df, 'badge', 'status')
    col_exclude = _find_col(df, 'exclude_reason') or _find_col(df, '除外')

    missing = [n for n, c in [
        ('学生番号', col_sid), ('氏名', col_name), ('参加日', col_date),
        ('結果統合フィルター', col_filter), ('badge_status', col_status),
    ] if c is None]
    if missing:
        raise ValueError(f'以下のカラムが見つかりません: {missing}\n列名一覧: {list(df.columns)}')

    # ---- 結果統合列の値の分布を表示 ----
    filter_dist = df[col_filter].value_counts(dropna=False)
    log(f'結果統合列の値の分布: {filter_dist.to_dict()}')

    # ---- フィルター（結果統合=1 のみ） ----
    mask_valid  = df[col_filter] == 1
    df_excluded = df[~mask_valid].copy()
    df_valid    = df[mask_valid].copy()

    log(f'有効行: {len(df_valid)} 行 / 除外行（結果統合≠1）: {len(df_excluded)} 行')
    if len(df_excluded) > 0:
        excl_dist = df_excluded[col_filter].value_counts(dropna=False)
        log(f'  除外内訳: {excl_dist.to_dict()}')

    df_valid = df_valid.copy()
    df_valid[col_sid]  = df_valid[col_sid].astype(str).str.strip()
    df_valid[col_date] = pd.to_datetime(df_valid[col_date], errors='coerce')

    # NaN日付の件数を把握
    nan_date_count = df_valid[col_date].isna().sum()
    log(f'  参加日NaN（日付パース失敗）: {nan_date_count} 行')

    # 学生番号の異常値チェック
    abnormal_sid = df_valid[df_valid[col_sid].isin(['', 'nan', 'NaN', 'None', 'na', 'NA'])]
    if len(abnormal_sid) > 0:
        log(f'  ⚠ 学生番号が空/nan の行: {len(abnormal_sid)} 行 → これらは1つのグループにまとめられます')

    # 全体の行数チェック（フィルター前）
    log(f'  シート全行: {total_input} = 有効{len(df_valid)} + 除外{len(df_excluded)} → 合計{len(df_valid)+len(df_excluded)} {"✓" if len(df_valid)+len(df_excluded)==total_input else "← 不一致！"}')

    # ---- 学生×プログラム名ごとに集計 ----
    records: dict = {}
    for _, row in df_valid.iterrows():
        sid  = _norm_sid(row[col_sid])
        prog = _norm(row[col_status])
        key  = (sid, prog)                   # ← (学生番号, プログラム名) で一意に管理
        if key not in records:
            records[key] = {
                'name_counter': Counter(),
                'dates_valid':  [],
            }
        records[key]['name_counter'][row[col_name]] += 1
        if pd.notna(row[col_date]):
            records[key]['dates_valid'].append(row[col_date])

    log(f'{len(records)} 件（学生×プログラム）のデータを処理中...')

    rows = []
    total_before_dedup = 0
    total_after_dedup  = 0

    before_reg_removed = 0

    for (sid, jcpp), data in records.items():
        name = data['name_counter'].most_common(1)[0][0]

        # 登録日・メールを取得（学生番号＋プログラム名で照合）
        reg_date = reg_lookup.get((sid, jcpp), None)
        email    = email_lookup.get((sid, jcpp), '')

        # 登録日より前の参加日を除外
        all_dates = sorted(data['dates_valid'])
        if reg_date and pd.notna(reg_date):
            filtered_dates = [d for d in all_dates if d >= reg_date]
            before_reg_removed += len(all_dates) - len(filtered_dates)
        else:
            filtered_dates = all_dates

        before_dedup = len(filtered_dates)
        sorted_dates = filtered_dates
        after_dedup  = len(sorted_dates)
        total_before_dedup += before_dedup
        total_after_dedup  += after_dedup
        badge = determine_badge(after_dedup, jcpp)

        row_data: dict = {
            'No':                   '',
            'プログラム名':         jcpp,
            '学生番号':             sid,
            '氏名':                 name,
            '国籍':                 '',
            'メールアドレス':       email,
            'オープンバッジ登録日': reg_date if (reg_date is not None and pd.notna(reg_date)) else '',
            '取得badge':            badge,
            '活動日':               after_dedup,
        }
        for j, d in enumerate(sorted_dates, start=1):
            row_data[f'参加日{j}'] = d.strftime('%Y-%m-%d')
        rows.append(row_data)

    # ---- 件数差異のサマリー ----
    dedup_removed = total_before_dedup - total_after_dedup
    log('─' * 40)
    log(f'【件数チェック】')
    log(f'  入力合計行数          : {total_input}')
    log(f'  除外（結果統合≠1）   : {len(df_excluded)}')
    log(f'  有効行（結果統合=1）  : {len(df_valid)}')
    log(f'  うちNaN日付で除外     : {nan_date_count}')
    log(f'  登録日より前で除外    : {before_reg_removed}')
    log(f'  重複日付で除去（同日複数回）: {dedup_removed}')
    log(f'  活動日合計            : {total_after_dedup}')
    log(f'  検算（除外+登録前+重複+活動日）: {len(df_excluded) + before_reg_removed + dedup_removed + total_after_dedup} {"✓" if len(df_excluded)+before_reg_removed+dedup_removed+total_after_dedup==total_input else "← 不一致！"}')
    log('─' * 40)

    # ---- プログラム名順にソートしてNo採番 ----
    rows.sort(key=lambda r: str(r['プログラム名']))
    for i, r in enumerate(rows, start=1):
        r['No'] = i

    # ---- 除外データシート ----
    excl_cols   = [col_sid, col_name, col_date, col_filter]
    excl_rename = {col_sid: '学生番号', col_name: '氏名',
                   col_date: '参加日', col_filter: '結果統合値'}
    if col_exclude:
        excl_cols.append(col_exclude)
        excl_rename[col_exclude] = '除外理由'

    df_excl_out = (
        df_excluded[excl_cols]
        .rename(columns=excl_rename)
        .reset_index(drop=True)
    )

    # ---- Excel出力 ----
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        out_df = pd.DataFrame(rows)
        out_df.to_excel(writer, sheet_name='結果', index=False)
        df_excl_out.to_excel(writer, sheet_name='除外データ', index=False)

        # バッジ列に色を付ける
        ws = writer.sheets['結果']
        badge_col_idx = out_df.columns.get_loc('取得badge') + 1  # 1始まり
        for row_idx in range(2, len(rows) + 2):          # 2行目～（1行目はヘッダー）
            cell  = ws.cell(row=row_idx, column=badge_col_idx)
            fill  = BADGE_FILLS.get(cell.value)
            if fill:
                cell.fill = fill

    log(f'完了！ {len(rows)} 人 → {output_path}')
    log(f'  ※「除外データ」シートに {len(df_excl_out)} 行の未適用データを記録しました')


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('使い方: python main.py <入力.xlsx> <出力.xlsx>')
        sys.exit(1)
    process(sys.argv[1], sys.argv[2])
