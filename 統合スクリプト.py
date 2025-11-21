# -*- coding: utf-8 -*-
"""
累積リスト統合プログラム（週次更新専用・完全版v3）

【カラム統一形式】（マッチング結果に合わせる）
- JANコード_旧
- 商品名称（カナ）_旧
- JANコード_新
- 商品名称（カナ）_新
- 新JAN備考
- データソース
- 期間
- 処理日

【花王・プラネット更新時】
累積から古い花王・プラネット部分を削除 → 新しいリストを追加

Author: HIBI KEITA
Version: 3.0
"""

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime


# ========================================================================
# カラム名の定義
# ========================================================================

# 内部処理用（マッチング結果の形式）
INTERNAL_COLUMNS = {
    'jan_old': 'JANコード_旧',
    'name_old': '商品名称（カナ）_旧',
    'jan_new': 'JANコード_新',
    'name_new': '商品名称（カナ）_新',
    'note': '新JAN備考',
    'source': 'データソース',
    'period': '期間',
    'date': '処理日',
}

# 最終出力用（Excel: 全カラム）
OUTPUT_COLUMNS_EXCEL = [
    '旧JANコード',
    '旧商品名',
    '新JANコード',
    '新商品名',
    '新JAN備考',
    'データソース',
    '期間',
    '処理日',
]

# 最終出力用（CSV: システム読み込み用）
OUTPUT_COLUMNS_CSV = [
    '旧JANコード',
    '旧商品名',
    '新JANコード',
    '新商品名',
    '新JAN備考',
]

# 内部→出力のカラム名変換
OUTPUT_COLUMN_MAPPING = {
    'JANコード_旧': '旧JANコード',
    '商品名称（カナ）_旧': '旧商品名',
    'JANコード_新': '新JANコード',
    '商品名称（カナ）_新': '新商品名',
    '新JAN備考': '新JAN備考',
    'データソース': 'データソース',
    '期間': '期間',
    '処理日': '処理日',
}

# ========================================================================
# ファイル読み込み関数
# ========================================================================

def load_file_flexible(file_path: str) -> pd.DataFrame:
    """CSV/TSV/Excelを自動判別して読み込み"""
    p = Path(file_path)
    ext = p.suffix.lower()
    
    print(f"📖 ファイル読み込み中: {p.name}")
    
    if ext in ['.xlsx', '.xls', '.xlsm']:
        try:
            df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
            print(f"✅ Excel読み込み成功")
            return df
        except Exception as e:
            raise ValueError(f"Excel読み込みエラー: {e}")
    
    encodings = ['utf-8', 'shift_jis', 'cp932']
    delimiter_map = {'.csv': ',', '.tsv': '\t', '.txt': '\t'}
    
    if ext not in delimiter_map:
        raise ValueError(f"サポート外の拡張子: {ext}")
    
    delimiter = delimiter_map[ext]
    
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding, delimiter=delimiter, 
                           on_bad_lines='skip', dtype=str)
            print(f"✅ {ext}読み込み成功（{encoding}）")
            return df
        except UnicodeDecodeError:
            continue
    
    raise UnicodeDecodeError(f"すべてのエンコーディングで読み込み失敗")


# ========================================================================
# 花王・プラネット用正規化（マッチング結果の形式に合わせる）
# ========================================================================

def normalize_kao_planet(df: pd.DataFrame, file_name: str, period: str) -> pd.DataFrame:
    """
    花王・プラネットのカラムをマッチング結果形式に変換
    
    入力: 旧JANコード, 旧商品名, 新JANコード, 新商品名, 新JAN備考
    出力: JANコード_旧, 商品名称（カナ）_旧, JANコード_新, 商品名称（カナ）_新, 新JAN備考, ...
    """
    print(f"📝 花王・プラネット正規化中...")
    
    # カラム名変換
    column_mapping = {
        '旧JANコード': INTERNAL_COLUMNS['jan_old'],
        '旧商品名': INTERNAL_COLUMNS['name_old'],
        '新JANコード': INTERNAL_COLUMNS['jan_new'],
        '新商品名': INTERNAL_COLUMNS['name_new'],
        '新JAN備考': INTERNAL_COLUMNS['note'],
    }
    
    df_normalized = df.rename(columns=column_mapping)
    
    # 必須カラムの追加
    df_normalized[INTERNAL_COLUMNS['source']] = '花王・プラネット'
    df_normalized[INTERNAL_COLUMNS['period']] = period
    df_normalized[INTERNAL_COLUMNS['date']] = datetime.now().strftime('%Y-%m-%d')
    
    # 新JAN備考がない場合はファイル名を入れる
    if INTERNAL_COLUMNS['note'] not in df_normalized.columns:
        df_normalized[INTERNAL_COLUMNS['note']] = file_name
    else:
        # 空の場合のみファイル名を入れる
        df_normalized[INTERNAL_COLUMNS['note']] = df_normalized[INTERNAL_COLUMNS['note']].fillna(file_name)
    
    # 統一カラム順に並び替え
    output_cols = list(INTERNAL_COLUMNS.values())
    for col in output_cols:
        if col not in df_normalized.columns:
            df_normalized[col] = ''
    
    df_normalized = df_normalized[output_cols].copy()
    
    print(f"✅ 正規化完了: {len(df_normalized)}件")
    return df_normalized


# ========================================================================
# マッチング結果用正規化
# ========================================================================

def normalize_matching(df: pd.DataFrame, file_name: str) -> pd.DataFrame:
    """
    マッチング結果のカラムを統一形式に変換
    （既にマッチング形式のカラム名なので、追加カラムのみ処理）
    """
    print(f"📝 マッチング結果正規化中...")
    
    # 既存のカラム名マッピング（念のため）
    column_mapping = {
        'JANコード_旧': INTERNAL_COLUMNS['jan_old'],
        '商品名称（カナ）_旧': INTERNAL_COLUMNS['name_old'],
        'JANコード_新': INTERNAL_COLUMNS['jan_new'],
        '商品名称（カナ）_新': INTERNAL_COLUMNS['name_new'],
    }
    
    df_normalized = df.rename(columns=column_mapping)
    
    # 必須カラムの追加
    if INTERNAL_COLUMNS['source'] not in df_normalized.columns:
        df_normalized[INTERNAL_COLUMNS['source']] = 'マッチング'
    
    if INTERNAL_COLUMNS['period'] not in df_normalized.columns:
        df_normalized[INTERNAL_COLUMNS['period']] = ''
    
    if INTERNAL_COLUMNS['date'] not in df_normalized.columns:
        df_normalized[INTERNAL_COLUMNS['date']] = datetime.now().strftime('%Y-%m-%d')
    
    if INTERNAL_COLUMNS['note'] not in df_normalized.columns:
        df_normalized[INTERNAL_COLUMNS['note']] = file_name
    
    print(f"✅ 正規化完了: {len(df_normalized)}件")
    return df_normalized


# ========================================================================
# JANコードクレンジング
# ========================================================================

def clean_jan_codes(df: pd.DataFrame) -> pd.DataFrame:
    """JANコードを13桁に統一"""
    print("🧹 JANコードクレンジング中...")
    
    jan_cols = [INTERNAL_COLUMNS['jan_old'], INTERNAL_COLUMNS['jan_new']]
    
    for col in jan_cols:
        if col in df.columns:
            df[col] = (df[col].astype(str)
                      .str.replace(r'\D+', '', regex=True)
                      .str.zfill(13)
                      .str[:13])
    
    print("✅ クレンジング完了")
    return df


# ========================================================================
# 累積から古い花王・プラネットデータを削除
# ========================================================================

def remove_old_kao_planet(existing_df: pd.DataFrame, periods_to_keep: list) -> pd.DataFrame:
    """
    累積リストから指定期間以外の花王・プラネットデータを削除
    
    Args:
        existing_df: 既存の累積リスト
        periods_to_keep: 保持する期間のリスト（例: ["24年下", "25年上"]）
    
    Returns:
        古いデータを削除した累積リスト
    """
    if existing_df.empty:
        return existing_df
    
    source_col = UNIFIED_COLUMNS['source']
    period_col = UNIFIED_COLUMNS['period']
    
    if source_col not in existing_df.columns or period_col not in existing_df.columns:
        print("⚠️ データソース/期間列がないため、削除処理をスキップ")
        return existing_df
    
    print(f"🗑️ 古い花王・プラネットデータを削除中...")
    print(f"   保持する期間: {periods_to_keep}")
    
    # 花王・プラネット以外のデータはそのまま保持
    non_kao_planet = existing_df[
        ~existing_df[source_col].str.contains('花王|プラネット', case=False, na=False, regex=True)
    ]
    
    # 花王・プラネットで保持する期間のデータ
    kao_planet_keep = existing_df[
        existing_df[source_col].str.contains('花王|プラネット', case=False, na=False, regex=True) &
        existing_df[period_col].isin(periods_to_keep)
    ]
    
    # 削除される件数を計算
    kao_planet_all = existing_df[
        existing_df[source_col].str.contains('花王|プラネット', case=False, na=False, regex=True)
    ]
    removed_count = len(kao_planet_all) - len(kao_planet_keep)
    
    if removed_count > 0:
        print(f"   削除: {removed_count}件")
    
    result = pd.concat([non_kao_planet, kao_planet_keep], ignore_index=True)
    print(f"✅ 削除後: {len(result)}件")
    
    return result


# ========================================================================
# 重複削除関数（完全版）
# ========================================================================

def remove_duplicates_advanced(existing_df: pd.DataFrame, 
                               kao_planet_df: pd.DataFrame, 
                               matching_df: pd.DataFrame) -> pd.DataFrame:
    """
    3段階の優先順位付き重複削除
    
    優先順位: 累積（既存） > 花王・プラネット > マッチング
    """
    print("\n📦 データ統合・重複削除開始...")
    
    jan_old = UNIFIED_COLUMNS['jan_old']
    jan_new = UNIFIED_COLUMNS['jan_new']
    source_col = UNIFIED_COLUMNS['source']
    
    print(f"  累積データ: {len(existing_df)}件")
    print(f"  花王・プラネット（今週）: {len(kao_planet_df)}件")
    print(f"  マッチング（今週）: {len(matching_df)}件")
    
    # ステップ1: 累積内の花王・プラネット由来JANを抽出
    existing_kao_planet_jans = set()
    if not existing_df.empty and source_col in existing_df.columns:
        kao_planet_rows = existing_df[
            existing_df[source_col].str.contains('花王|プラネット', case=False, na=False, regex=True)
        ]
        existing_kao_planet_jans = set(kao_planet_rows[jan_new].dropna())
        print(f"  累積内の花王・プラネット由来JAN: {len(existing_kao_planet_jans)}件")
    
    # ステップ2: マッチング結果から累積内花王・プラネット重複を削除
    if not matching_df.empty and existing_kao_planet_jans:
        before = len(matching_df)
        matching_df = matching_df[~matching_df[jan_new].isin(existing_kao_planet_jans)].copy()
        removed = before - len(matching_df)
        if removed > 0:
            print(f"  ✂️ マッチング→累積内花王・プラネット重複削除: {removed}件")
    
    # ステップ3: 今週の花王・プラネットとマッチングの重複削除
    if not kao_planet_df.empty and not matching_df.empty:
        kao_planet_old_jans = set(kao_planet_df[jan_old].dropna())
        kao_planet_new_jans = set(kao_planet_df[jan_new].dropna())
        
        before = len(matching_df)
        matching_df = matching_df[~matching_df[jan_new].isin(kao_planet_new_jans)].copy()
        matching_df = matching_df[~matching_df[jan_old].isin(kao_planet_old_jans)].copy()
        removed = before - len(matching_df)
        if removed > 0:
            print(f"  ✂️ マッチング→今週花王・プラネット重複削除: {removed}件")
    
    # ステップ4: データ結合（優先順位順）
    all_data = pd.concat([existing_df, kao_planet_df, matching_df], ignore_index=True)
    print(f"  統合後: {len(all_data)}件")
    
    # ステップ5: 新JANで重複削除
    before = len(all_data)
    all_data = all_data.drop_duplicates(subset=[jan_new], keep='first')
    removed = before - len(all_data)
    if removed > 0:
        print(f"  🗑️ 新JAN重複削除: {removed}件")
    
    # ステップ6: 旧JAN=新JANのものを削除
    before = len(all_data)
    all_data = all_data[all_data[jan_old] != all_data[jan_new]].copy()
    removed = before - len(all_data)
    if removed > 0:
        print(f"  🔄 同一JAN削除: {removed}件")
    
    print(f"  ✅ 最終件数: {len(all_data)}件")
    return all_data


# ========================================================================
# メイン処理
# ========================================================================

def main():
    print("=" * 60)
    print("累積リスト統合プログラム - 週次更新（v3）")
    print("=" * 60)
    
    # ========== ステップ1: 累積リスト読み込み ==========
    print("\n【ステップ1】累積リストの読み込み")
    
    root = tk.Tk()
    root.withdraw()
    
    existing_df = pd.DataFrame()
    
    if messagebox.askyesno("累積リスト", "前週の累積リストがありますか？\n（初回は「いいえ」）"):
        existing_path = filedialog.askopenfilename(
            title="前週の累積リストを選択",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("すべて", "*.*")]
        )
        
        if existing_path:
            try:
                existing_df = load_file_flexible(existing_path)
                existing_df = clean_jan_codes(existing_df)
                print(f"✅ 累積リスト読み込み: {len(existing_df)}件")
            except Exception as e:
                messagebox.showerror("エラー", f"累積リスト読み込み失敗:\n{e}")
                root.destroy()
                return
    else:
        print("📂 新規作成モード")
    
    root.destroy()
    
    # ========== ステップ2: マッチング結果読み込み ==========
    print("\n【ステップ2】マッチング結果の読み込み")
    
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo("選択", "今週のマッチング結果ファイルを選択してください")
    
    matching_path = filedialog.askopenfilename(
        title="マッチング結果を選択",
        filetypes=[("Excel/CSV", "*.xlsx *.csv *.tsv"), ("すべて", "*.*")]
    )
    
    root.destroy()
    
    if not matching_path:
        messagebox.showwarning("キャンセル", "マッチング結果が選択されませんでした")
        return
    
    try:
        matching_df = load_file_flexible(matching_path)
        matching_df = normalize_matching(matching_df, Path(matching_path).name)
        matching_df = clean_jan_codes(matching_df)
        print(f"✅ マッチング結果読み込み: {len(matching_df)}件")
    except Exception as e:
        messagebox.showerror("エラー", f"マッチング結果読み込み失敗:\n{e}")
        return
    
    # ========== ステップ3: 花王・プラネット差し替えリスト読み込み ==========
    print("\n【ステップ3】花王・プラネット差し替えリストの読み込み")
    
    root = tk.Tk()
    root.withdraw()
    
    kao_planet_df = pd.DataFrame()
    
    choice = messagebox.askquestion(
        "花王・プラネット", 
        "今週、花王・プラネットの差し替えリストを更新しますか？\n"
        "（半年に1回程度）\n\n"
        "「はい」→ 新しいファイルを読み込み、古いデータを削除\n"
        "「いいえ」→ スキップ（マッチングのみ統合）"
    )
    
    if choice == 'yes':
        # 期間入力
        period = simpledialog.askstring(
            "期間入力",
            "花王・プラネットの期間を入力してください\n（例: 25年上、25年下）",
            initialvalue="25年上"
        )
        
        if not period:
            period = datetime.now().strftime('%Y年')
        
        # 保持する期間を入力
        keep_periods_str = simpledialog.askstring(
            "保持期間",
            "累積から保持する期間をカンマ区切りで入力\n（例: 24年下,25年上）\n\n"
            "※古い期間のデータは削除されます\n"
            "※空欄で全削除してから新規追加",
            initialvalue="24年下,25年上"
        )
        
        if keep_periods_str:
            periods_to_keep = [p.strip() for p in keep_periods_str.split(',')]
        else:
            periods_to_keep = []
        
        # 古いデータを削除
        existing_df = remove_old_kao_planet(existing_df, periods_to_keep)
        
        # 新しい花王・プラネットファイルを読み込み
        kao_planet_path = filedialog.askopenfilename(
            title="花王・プラネット差し替えリストを選択",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("すべて", "*.*")]
        )
        
        if kao_planet_path:
            try:
                kao_planet_df = load_file_flexible(kao_planet_path)
                kao_planet_df = normalize_kao_planet(
                    kao_planet_df, 
                    Path(kao_planet_path).name,
                    period
                )
                kao_planet_df = clean_jan_codes(kao_planet_df)
                print(f"✅ 花王・プラネット読み込み: {len(kao_planet_df)}件")
            except Exception as e:
                messagebox.showerror("エラー", f"花王・プラネット読み込み失敗:\n{e}")
                root.destroy()
                return
    else:
        print("📂 花王・プラネットなし（マッチングのみ統合）")
    
    root.destroy()
    
    # ========== ステップ4: 出力先選択 ==========
    print("\n【ステップ4】出力先の選択")
    
    root = tk.Tk()
    root.withdraw()
    
    output_dir = filedialog.askdirectory(title="保存先フォルダを選択")
    
    root.destroy()
    
    if not output_dir:
        messagebox.showwarning("キャンセル", "保存先が選択されませんでした")
        return
    
    output_dir = Path(output_dir)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_csv = output_dir / f"累積_差し替えリスト_{timestamp}.csv"
    output_excel = output_dir / f"累積_差し替えリスト_{timestamp}.xlsx"
    latest_csv = output_dir / "累積_差し替えリスト_最新.csv"
    latest_excel = output_dir / "累積_差し替えリスト_最新.xlsx"
    
    # ========== ステップ5: データ統合 ==========
    print("\n【ステップ5】データ統合・重複削除")
    
    final_df = remove_duplicates_advanced(existing_df, kao_planet_df, matching_df)
    
    # ========== ステップ6: 保存 ==========
    print("\n【ステップ6】ファイル保存")
    
    try:
        print(f"💾 CSV保存中（cp932）: {output_csv.name}")
        final_df.to_csv(output_csv, index=False, encoding='cp932', errors='replace')
        final_df.to_csv(latest_csv, index=False, encoding='cp932', errors='replace')
        
        print(f"💾 Excel保存中: {output_excel.name}")
        final_df.to_excel(output_excel, index=False, engine='openpyxl')
        final_df.to_excel(latest_excel, index=False, engine='openpyxl')
        
        print("✅ 保存完了")
    except Exception as e:
        messagebox.showerror("エラー", f"ファイル保存失敗:\n{e}")
        return
    
    # ========== 完了メッセージ ==========
    source_col = UNIFIED_COLUMNS['source']
    source_counts = final_df[source_col].value_counts().to_dict() if source_col in final_df.columns else {}
    
    summary = f"""
🎉 統合処理完了！

【累積データ】
総件数: {len(final_df)}件

【データソース別】
{chr(10).join([f'  {k}: {v}件' for k, v in source_counts.items()])}

【保存先】
📁 {output_dir}

【ファイル】
- {output_csv.name}
- {output_excel.name}
- 累積_差し替えリスト_最新.csv/xlsx
"""
    
    print(summary)
    messagebox.showinfo("完了", summary)

# ========================================================================
# 実行
# ========================================================================

if __name__ == "__main__":
    main()