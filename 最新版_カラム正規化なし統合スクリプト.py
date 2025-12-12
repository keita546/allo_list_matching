# -*- coding: utf-8 -*-
"""
累積リスト統合プログラム（週次更新専用・花王/プラネット自動振り分け版v5.1）

【変更点】
- 花王・プラネットリストを1ファイルから読み込み
- データソース列で自動振り分け
- 各データソースごとに独立した期間管理

Author: HIBI KEITA
Version: 5.1
"""

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime


# ========================================================================
# カラム定義
# ========================================================================

# 必須カラム（入力・出力共通）
REQUIRED_COLUMNS = ['旧JANコード', '旧商品名', '新JANコード', '新商品名', '新JAN備考']

# メタデータカラム
METADATA_COLUMNS = ['データソース', '期間', '処理日']

# Excel出力用（全カラム）
OUTPUT_COLUMNS_EXCEL = REQUIRED_COLUMNS + METADATA_COLUMNS

# CSV出力用（システム読み込み用）
OUTPUT_COLUMNS_CSV = REQUIRED_COLUMNS


# ========================================================================
# ファイル読み込み
# ========================================================================

def load_file_flexible(file_path: str, sheet_name: str = '確定') -> pd.DataFrame:
    """CSV/TSV/Excelを自動判別して読み込み"""
    p = Path(file_path)
    ext = p.suffix.lower()
    
    print(f"📖 ファイル読み込み中: {p.name}")
    
    # Excel
    if ext in ['.xlsx', '.xls', '.xlsm']:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', dtype=str)
            print(f"✅ Excel読み込み成功（シート: {sheet_name}）")
            return df
        except Exception as e:
            raise ValueError(f"Excel読み込みエラー: {e}")
    
    # CSV/TSV
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
# データ処理
# ========================================================================

def normalize_columns(df: pd.DataFrame, file_type: str = 'matching') -> pd.DataFrame:
    """
    カラム名を標準形式に正規化
    
    Args:
        df: 入力データ
        file_type: 'matching' or 'kao_planet'
    """
    print(f"📝 カラム正規化中: {file_type}")
    
    if file_type == 'kao_planet':
        # 花王・プラネットは固定形式
        if set(REQUIRED_COLUMNS).issubset(set(df.columns)):
            print("✅ 既に標準形式です")
            return df
        else:
            raise ValueError(f"花王・プラネットデータに必須カラムがありません: {set(REQUIRED_COLUMNS) - set(df.columns)}")
    
    elif file_type == 'matching':
        # マッチング結果は自動変換
        print(f"  元のカラム: {list(df.columns)}")
        
        # カラム名の正規化マッピング
        rename_map = {}
        
        for col in df.columns:
            # 旧JANコード → そのまま
            if '旧JAN' in col and 'コード' in col:
                rename_map[col] = '旧JANコード'
            
            # 旧商品名（漢字）※漢字なければカナ → 旧商品名
            elif '旧商品名' in col:
                rename_map[col] = '旧商品名'
            
            # 新JANコード → そのまま
            elif '新JAN' in col and 'コード' in col:
                rename_map[col] = '新JANコード'
            
            # 新商品名（漢字）※漢字なければカナ → 新商品名
            elif '新商品名' in col:
                rename_map[col] = '新商品名'
            
            # 新JAN備考 → そのまま
            elif '新JAN備考' in col or '備考' in col:
                rename_map[col] = '新JAN備考'
        
        print(f"  変換マップ: {rename_map}")
        
        # リネーム
        df = df.rename(columns=rename_map)
        
        # 必須カラムチェック（新JAN備考以外）
        required = ['旧JANコード', '旧商品名', '新JANコード', '新商品名']
        missing = [col for col in required if col not in df.columns]
        
        if missing:
            raise ValueError(f"必須カラムが見つかりません: {missing}\n元のカラム: {list(df.columns)}")
        
        # 新JAN備考がない場合は空列追加
        if '新JAN備考' not in df.columns:
            df['新JAN備考'] = ''
        
        print(f"  変換後: {list(df.columns)}")
        print("✅ カラム正規化完了")
        return df
    
    else:
        raise ValueError(f"不明なfile_type: {file_type}")


def split_kao_planet_list(df: pd.DataFrame) -> tuple:
    """
    花王・プラネットリストをデータソース列で振り分け
    
    Args:
        df: 花王・プラネットリスト（データソース列あり）
    
    Returns:
        (kao_df, planet_df): 花王データ、プラネットデータ
    """
    print("🔀 花王・プラネットデータを振り分け中...")
    
    if 'データソース' not in df.columns:
        raise ValueError("データソース列が見つかりません")
    
    # 花王データ
    kao_df = df[df['データソース'] == '花王'].copy()
    print(f"  花王: {len(kao_df)}件")
    
    # プラネットデータ
    planet_df = df[df['データソース'] == 'プラネット'].copy()
    print(f"  プラネット: {len(planet_df)}件")
    
    # その他（警告）
    other_df = df[~df['データソース'].isin(['花王', 'プラネット'])]
    if len(other_df) > 0:
        print(f"  ⚠️ その他（無視）: {len(other_df)}件")
        print(f"     データソース値: {other_df['データソース'].unique().tolist()}")
    
    print("✅ 振り分け完了")
    return kao_df, planet_df


def update_metadata(df: pd.DataFrame, period: str, note: str = '') -> pd.DataFrame:
    """
    メタデータを更新（データソースは既に入っている前提）
    
    Args:
        df: 入力データ（データソース列あり）
        period: 期間（例: 25年上）
        note: 新JAN備考に追加する文字列（空の場合のみ）
    """
    print(f"📝 メタデータ更新中...")
    
    df = df.copy()
    
    # 期間と処理日を更新
    df['期間'] = period
    df['処理日'] = datetime.now().strftime('%Y-%m-%d')
    
    # 新JAN備考が空の場合のみnoteを入れる
    if note and '新JAN備考' in df.columns:
        df.loc[df['新JAN備考'].isna() | (df['新JAN備考'] == ''), '新JAN備考'] = note
    
    # カラム順を統一
    df = df[OUTPUT_COLUMNS_EXCEL].copy()
    
    print(f"✅ 処理完了: {len(df)}件")
    return df


def add_metadata(df: pd.DataFrame, source: str, period: str = '', note: str = '') -> pd.DataFrame:
    """
    メタデータを追加（マッチング用）
    
    Args:
        df: 入力データ（既に標準カラム名）
        source: データソース（マッチング）
        period: 期間（例: 25年上）
        note: 新JAN備考に追加する文字列（空の場合のみ）
    """
    print(f"📝 メタデータ追加中: {source}")
    
    df = df.copy()
    
    # メタデータ追加
    df['データソース'] = source
    df['期間'] = period
    df['処理日'] = datetime.now().strftime('%Y-%m-%d')
    
    # 新JAN備考が空の場合のみnoteを入れる
    if note and '新JAN備考' in df.columns:
        df['新JAN備考'] = df['新JAN備考'].fillna(note).replace('', note)
    
    # カラム順を統一
    df = df[OUTPUT_COLUMNS_EXCEL].copy()
    
    print(f"✅ 処理完了: {len(df)}件")
    return df


def clean_jan_codes(df: pd.DataFrame) -> pd.DataFrame:
    """JANコードを13桁に統一"""
    print("🧹 JANコードクレンジング中...")
    
    jan_cols = ['旧JANコード', '新JANコード']
    
    for col in jan_cols:
        if col in df.columns:
            df[col] = (df[col].astype(str)
                      .str.replace(r'\D+', '', regex=True)
                      .str.zfill(13)
                      .str[:13])
    
    print("✅ クレンジング完了")
    return df


def remove_specific_source_data(existing_df: pd.DataFrame, source_name: str, periods_to_keep: list) -> pd.DataFrame:
    """
    累積リストから特定データソース（花王 or プラネット）の指定期間以外を削除
    
    Args:
        existing_df: 既存の累積リスト
        source_name: 削除対象のデータソース名（'花王' or 'プラネット'）
        periods_to_keep: 保持する期間のリスト（例: ["25年下", "26年上"]）
    """
    if existing_df.empty:
        return existing_df
    
    if 'データソース' not in existing_df.columns or '期間' not in existing_df.columns:
        print("⚠️ データソース/期間列がないため、削除処理をスキップ")
        return existing_df
    
    print(f"🗑️ 古い{source_name}データを削除中...")
    print(f"   保持する期間: {periods_to_keep}")
    
    # 対象データソース以外はそのまま保持
    non_target = existing_df[existing_df['データソース'] != source_name]
    
    # 対象データソースで保持する期間のデータ
    target_keep = existing_df[
        (existing_df['データソース'] == source_name) &
        (existing_df['期間'].isin(periods_to_keep))
    ]
    
    # 削除される件数を計算
    target_all = existing_df[existing_df['データソース'] == source_name]
    removed_count = len(target_all) - len(target_keep)
    
    if removed_count > 0:
        print(f"   {source_name}削除: {removed_count}件")
        print(f"   {source_name}保持: {len(target_keep)}件")
    
    result = pd.concat([non_target, target_keep], ignore_index=True)
    print(f"✅ 削除後合計: {len(result)}件")
    
    return result


def merge_and_deduplicate(existing_df: pd.DataFrame, 
                          kao_df: pd.DataFrame, 
                          planet_df: pd.DataFrame,
                          matching_df: pd.DataFrame) -> pd.DataFrame:
    """
    4段階の優先順位付き統合・重複削除
    
    優先順位: 累積（既存） > 花王 > プラネット > マッチング
    """
    print("\n📦 データ統合・重複削除開始...")
    
    print(f"  累積データ: {len(existing_df)}件")
    print(f"  花王（今週）: {len(kao_df)}件")
    print(f"  プラネット（今週）: {len(planet_df)}件")
    print(f"  マッチング（今週）: {len(matching_df)}件")
    
    # ステップ1: 累積内の花王・プラネット由来の新JANを抽出
    existing_kao_planet_jans = set()
    if not existing_df.empty and 'データソース' in existing_df.columns:
        kao_planet_rows = existing_df[
            existing_df['データソース'].isin(['花王', 'プラネット'])
        ]
        existing_kao_planet_jans = set(kao_planet_rows['新JANコード'].dropna())
        print(f"  累積内の花王・プラネット由来JAN: {len(existing_kao_planet_jans)}件")
    
    # ステップ2: マッチング結果から累積内花王・プラネット重複を削除
    if not matching_df.empty and existing_kao_planet_jans:
        before = len(matching_df)
        matching_df = matching_df[~matching_df['新JANコード'].isin(existing_kao_planet_jans)].copy()
        removed = before - len(matching_df)
        if removed > 0:
            print(f"  ✂️ マッチング→累積内花王・プラネット重複削除: {removed}件")
    
    # ステップ3: 今週の花王・プラネットとマッチングの重複削除
    new_kao_planet_df = pd.concat([kao_df, planet_df], ignore_index=True)
    
    if not new_kao_planet_df.empty and not matching_df.empty:
        kao_planet_old_jans = set(new_kao_planet_df['旧JANコード'].dropna())
        kao_planet_new_jans = set(new_kao_planet_df['新JANコード'].dropna())
        
        before = len(matching_df)
        matching_df = matching_df[~matching_df['新JANコード'].isin(kao_planet_new_jans)].copy()
        matching_df = matching_df[~matching_df['旧JANコード'].isin(kao_planet_old_jans)].copy()
        removed = before - len(matching_df)
        if removed > 0:
            print(f"  ✂️ マッチング→今週花王・プラネット重複削除: {removed}件")
    
    # ステップ4: データ結合（優先順位順）
    all_data = pd.concat([existing_df, kao_df, planet_df, matching_df], ignore_index=True)
    print(f"  統合後: {len(all_data)}件")
    
    # ステップ5: 新JANで重複削除（先頭優先）
    before = len(all_data)
    all_data = all_data.drop_duplicates(subset=['新JANコード'], keep='first')
    removed = before - len(all_data)
    if removed > 0:
        print(f"  🗑️ 新JAN重複削除: {removed}件")
    
    # ステップ6: 旧JAN=新JANの同一JANを削除
    before = len(all_data)
    all_data = all_data[all_data['旧JANコード'] != all_data['新JANコード']].copy()
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
    print("累積リスト統合プログラム - 週次更新（v5.1・花王/プラネット自動振り分け版）")
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
        matching_df = normalize_columns(matching_df, file_type='matching')
        matching_df = add_metadata(
            matching_df, 
            source='マッチング',
            note=Path(matching_path).name
        )
        matching_df = clean_jan_codes(matching_df)
        print(f"✅ マッチング結果読み込み: {len(matching_df)}件")
    except Exception as e:
        messagebox.showerror("エラー", f"マッチング結果読み込み失敗:\n{e}")
        return
    
    # ========== ステップ3: 花王・プラネットリスト読み込み ==========
    print("\n【ステップ3】花王・プラネットリストの読み込み")
    
    root = tk.Tk()
    root.withdraw()
    
    kao_df = pd.DataFrame()
    planet_df = pd.DataFrame()
    
    choice = messagebox.askquestion(
        "花王・プラネットリスト", 
        "今週、花王・プラネットリストを更新しますか？\n\n"
        "「はい」→ 新しいファイルを読み込み\n"
        "「いいえ」→ スキップ（既存データを保持）"
    )
    
    if choice == 'yes':
        # ファイル選択
        kao_planet_path = filedialog.askopenfilename(
            title="花王・プラネットリストを選択",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("すべて", "*.*")]
        )
        
        if not kao_planet_path:
            messagebox.showwarning("キャンセル", "花王・プラネットリストが選択されませんでした")
        else:
            try:
                # ファイル読み込み
                kao_planet_df = load_file_flexible(kao_planet_path)
                kao_planet_df = normalize_columns(kao_planet_df, file_type='kao_planet')
                kao_planet_df = clean_jan_codes(kao_planet_df)
                
                # データソース列チェック
                if 'データソース' not in kao_planet_df.columns:
                    raise ValueError("データソース列が見つかりません")
                
                # 花王とプラネットに振り分け
                kao_df, planet_df = split_kao_planet_list(kao_planet_df)
                
                # 花王の期間設定
                if not kao_df.empty:
                    kao_period = simpledialog.askstring(
                        "花王期間入力",
                        f"花王リストの期間を入力してください\n（例: 25年下,26年上）\n\n"
                        f"読み込んだ花王データ: {len(kao_df)}件",
                        initialvalue="25年下,26年上"
                    )
                    
                    if not kao_period:
                        kao_period = datetime.now().strftime('%Y年')
                    
                    # 保持する期間
                    kao_keep_periods_str = simpledialog.askstring(
                        "花王保持期間",
                        "累積から保持する花王の期間をカンマ区切りで入力\n"
                        "（例: 25年下,26年上）\n\n"
                        "※この期間以外の花王データは削除されます\n"
                        "※空欄で花王データ全削除してから新規追加",
                        initialvalue=kao_period
                    )
                    
                    if kao_keep_periods_str:
                        kao_periods_to_keep = [p.strip() for p in kao_keep_periods_str.split(',')]
                    else:
                        kao_periods_to_keep = []
                    
                    # 古い花王データ削除
                    existing_df = remove_specific_source_data(existing_df, '花王', kao_periods_to_keep)
                    
                    # メタデータ更新
                    kao_df = update_metadata(kao_df, kao_period, Path(kao_planet_path).name)
                    print(f"✅ 花王データ処理完了: {len(kao_df)}件")
                
                # プラネットの期間設定
                if not planet_df.empty:
                    planet_period = simpledialog.askstring(
                        "プラネット期間入力",
                        f"プラネットリストの期間を入力してください\n（例: 25年上,25年下）\n\n"
                        f"読み込んだプラネットデータ: {len(planet_df)}件",
                        initialvalue="25年上,25年下"
                    )
                    
                    if not planet_period:
                        planet_period = datetime.now().strftime('%Y年')
                    
                    # 保持する期間
                    planet_keep_periods_str = simpledialog.askstring(
                        "プラネット保持期間",
                        "累積から保持するプラネットの期間をカンマ区切りで入力\n"
                        "（例: 25年上,25年下）\n\n"
                        "※この期間以外のプラネットデータは削除されます\n"
                        "※空欄でプラネットデータ全削除してから新規追加",
                        initialvalue=planet_period
                    )
                    
                    if planet_keep_periods_str:
                        planet_periods_to_keep = [p.strip() for p in planet_keep_periods_str.split(',')]
                    else:
                        planet_periods_to_keep = []
                    
                    # 古いプラネットデータ削除
                    existing_df = remove_specific_source_data(existing_df, 'プラネット', planet_periods_to_keep)
                    
                    # メタデータ更新
                    planet_df = update_metadata(planet_df, planet_period, Path(kao_planet_path).name)
                    print(f"✅ プラネットデータ処理完了: {len(planet_df)}件")
                
            except Exception as e:
                messagebox.showerror("エラー", f"花王・プラネットリスト読み込み失敗:\n{e}")
                root.destroy()
                return
    else:
        print("📂 花王・プラネットリストなし（既存データ保持）")
    
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
    
    final_df = merge_and_deduplicate(existing_df, kao_df, planet_df, matching_df)
    
    # ========== ステップ6: 保存 ==========
    print("\n【ステップ6】ファイル保存")
    
    try:
        # CSV（システム読み込み用: メタデータなし）
        print(f"💾 CSV保存中（システム用・cp932）: {output_csv.name}")
        final_df[OUTPUT_COLUMNS_CSV].to_csv(output_csv, index=False, encoding='cp932', errors='replace')
        final_df[OUTPUT_COLUMNS_CSV].to_csv(latest_csv, index=False, encoding='cp932', errors='replace')
        
        # Excel（管理用: メタデータあり）
        print(f"💾 Excel保存中（管理用・全カラム）: {output_excel.name}")
        final_df[OUTPUT_COLUMNS_EXCEL].to_excel(output_excel, index=False, engine='openpyxl')
        final_df[OUTPUT_COLUMNS_EXCEL].to_excel(latest_excel, index=False, engine='openpyxl')
        
        print("✅ 保存完了")
    except Exception as e:
        messagebox.showerror("エラー", f"ファイル保存失敗:\n{e}")
        return
    
    # ========== 完了メッセージ ==========
    source_counts = final_df['データソース'].value_counts().to_dict() if 'データソース' in final_df.columns else {}
    
    # 期間別の集計も追加
    period_summary = ""
    if 'データソース' in final_df.columns and '期間' in final_df.columns:
        period_summary = "\n【データソース×期間別】\n"
        for source in ['花王', 'プラネット', 'マッチング']:
            source_data = final_df[final_df['データソース'] == source]
            if not source_data.empty:
                period_counts = source_data['期間'].value_counts().to_dict()
                period_summary += f"  {source}:\n"
                for period, count in period_counts.items():
                    period_summary += f"    - {period}: {count}件\n"
    
    summary = f"""
🎉 統合処理完了！

【累積データ】
総件数: {len(final_df)}件

【データソース別】
{chr(10).join([f'  {k}: {v}件' for k, v in source_counts.items()])}
{period_summary}
【保存先】
📁 {output_dir}

【ファイル】
- CSV（システム用）: {output_csv.name}
- Excel（管理用）: {output_excel.name}
- 累積_差し替えリスト_最新.csv/xlsx
"""
    
    print(summary)
    messagebox.showinfo("完了", summary)


# ========================================================================
# 実行
# ========================================================================

if __name__ == "__main__":
    main()