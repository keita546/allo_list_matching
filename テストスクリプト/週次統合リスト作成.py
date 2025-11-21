# -*- coding: utf-8 -*-
"""
累積リスト統合プログラム（週次更新専用・完全版）

【処理フロー】
Week1: 花王・プラネット（任意） + マッチング → 累積リスト
Week2以降: 前週の累積 + 今週のマッチング → 更新された累積リスト

【優先順位】
1. 前週の累積リスト（全データ保持）
2. 今週の花王・プラネット（半年に1回のみ）
3. 今週のマッチング結果

Author: HIBI KEITA
Version: 2.0
"""

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime


# ========================================================================
# ファイル読み込み関数
# ========================================================================

def load_file_flexible(file_path: str) -> pd.DataFrame:
    """
    CSV/TSV/Excelを自動判別して読み込み
    """
    p = Path(file_path)
    ext = p.suffix.lower()
    
    print(f"📖 ファイル読み込み中: {p.name}")
    
    # Excel
    if ext in ['.xlsx', '.xls', '.xlsm']:
        try:
            df = pd.read_excel(file_path, engine='openpyxl', dtype=str)
            print(f"✅ Excelファイル読み込み成功")
            return df
        except Exception as e:
            raise ValueError(f"Excel読み込みエラー: {e}")
    
    # CSV/TSV
    encodings = ['utf-8', 'shift_jis', 'cp932']
    delimiter_map = {
        '.csv': ',',
        '.tsv': '\t',
        '.txt': '\t',
    }
    
    if ext not in delimiter_map:
        raise ValueError(f"サポート外の拡張子: {ext}")
    
    delimiter = delimiter_map[ext]
    
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding, delimiter=delimiter, 
                           on_bad_lines='skip', dtype=str)
            print(f"✅ {ext}ファイル読み込み成功（{encoding}）")
            return df
        except UnicodeDecodeError:
            continue
        except Exception as e:
            raise Exception(f"ファイル読み込みエラー: {e}")
    
    raise UnicodeDecodeError(f"すべてのエンコーディングで読み込み失敗")


# ========================================================================
# データ正規化関数
# ========================================================================

def normalize_columns(df: pd.DataFrame, source: str) -> pd.DataFrame:
    """
    カラム名を統一フォーマットに変換
    
    統一フォーマット:
    - 旧JANコード
    - 旧商品名
    - 新JANコード
    - 新商品名
    - メーカー名称
    - 備考
    - 処理日
    - データソース
    """
    print(f"📝 カラム名正規化中（{source}）...")
    
    # カラム名のマッピング辞書
    column_mapping = {
        # マッチング結果のパターン
        'JANコード_旧': '旧JANコード',
        'JANコード_新': '新JANコード',
        '商品名称（カナ）_旧': '旧商品名',
        '商品名称（カナ）_新': '新商品名',
        'メーカー名称_旧': 'メーカー名称',
        'メーカー名称_新': 'メーカー名称',
        
        # 花王・プラネットのパターン
        '旧JAN': '旧JANコード',
        '新JAN': '新JANコード',
        
        # その他のバリエーション
        'JAN旧': '旧JANコード',
        'JAN新': '新JANコード',
    }
    
    # カラム名を変換
    df_normalized = df.rename(columns=column_mapping)
    
    # 必須カラムの確認と追加
    required_columns = {
        '旧JANコード': '',
        '旧商品名': '',
        '新JANコード': '',
        '新商品名': '',
        'メーカー名称': '',
        '備考': '',
        '処理日': datetime.now().strftime('%Y-%m-%d'),
        'データソース': source,
    }
    
    for col, default_val in required_columns.items():
        if col not in df_normalized.columns:
            df_normalized[col] = default_val
    
    # 統一カラムのみ抽出
    output_columns = list(required_columns.keys())
    df_normalized = df_normalized[output_columns].copy()
    
    print(f"✅ 正規化完了: {len(df_normalized)}件")
    return df_normalized


# ========================================================================
# JANコードクレンジング
# ========================================================================

def clean_jan_codes(df: pd.DataFrame) -> pd.DataFrame:
    """
    JANコードを13桁に統一
    """
    print("🧹 JANコードクレンジング中...")
    
    for col in ['旧JANコード', '新JANコード']:
        if col in df.columns:
            df[col] = (df[col].astype(str)
                      .str.replace(r'\D+', '', regex=True)  # 数字以外削除
                      .str.zfill(13)  # 13桁に0埋め
                      .str[:13])  # 13桁切り取り
    
    print("✅ クレンジング完了")
    return df


# ========================================================================
# 重複削除関数（完全版）
# ========================================================================

def remove_duplicates_advanced(existing_df: pd.DataFrame, 
                               kao_planet_df: pd.DataFrame, 
                               matching_df: pd.DataFrame) -> pd.DataFrame:
    """
    3段階の優先順位付き重複削除
    
    【優先順位】
    1. 既存の累積リスト（全データ保持）
    2. 花王・プラネット差し替えリスト（半年に1回）
    3. マッチング結果（毎週）
    
    【重複削除ロジック】
    - ステップ1: 累積内の花王・プラネット由来データの新JANを抽出
    - ステップ2: マッチング結果から、花王・プラネット由来と新JANが被るものを削除
    - ステップ3: 今週の花王・プラネットとマッチングで、旧JAN・新JANが被るものを削除（マッチング側を削除）
    - ステップ4: 3つのデータを結合（優先順位順）
    - ステップ5: 新JANで重複削除（最初に出現した行を残す = 優先順位が高い方を残す）
    - ステップ6: 旧JAN=新JANのものを削除
    """
    print("\n📦 データ統合・重複削除開始...")
    
    print(f"  累積データ: {len(existing_df)}件")
    print(f"  花王・プラネット（今週）: {len(kao_planet_df)}件")
    print(f"  マッチング（今週）: {len(matching_df)}件")
    
    # ========== ステップ1: 累積内の花王・プラネット由来JANを抽出 ==========
    existing_kao_planet_jans = set()
    if not existing_df.empty:
        if 'データソース' in existing_df.columns:
            # 表記揺れに対応（部分一致で判定）
            kao_planet_rows = existing_df[
                existing_df['データソース'].str.contains('花王|プラネット|KAO|PLANET', 
                                                        case=False, 
                                                        na=False, 
                                                        regex=True)
            ]
            existing_kao_planet_jans = set(kao_planet_rows['新JANコード'].dropna())
            print(f"  累積内の花王・プラネット由来JAN: {len(existing_kao_planet_jans)}件")
        else:
            print("  ⚠️ 累積リストに「データソース」列がありません（スキップ）")
    
    # ========== ステップ2: マッチング結果から花王・プラネット重複を削除 ==========
    if not matching_df.empty and existing_kao_planet_jans:
        before_count = len(matching_df)
        matching_df = matching_df[~matching_df['新JANコード'].isin(existing_kao_planet_jans)].copy()
        removed_count = before_count - len(matching_df)
        if removed_count > 0:
            print(f"  ✂️ マッチング→累積内花王・プラネット重複削除: {removed_count}件")
    
    # ========== ステップ3: 今週の花王・プラネットとマッチングの重複削除 ==========
    if not kao_planet_df.empty and not matching_df.empty:
        # 旧JAN・新JANの両方で突合
        kao_planet_old_jans = set(kao_planet_df['旧JANコード'].dropna())
        kao_planet_new_jans = set(kao_planet_df['新JANコード'].dropna())
        
        before_count = len(matching_df)
        
        # 新JANで重複しているものを削除
        matching_df = matching_df[~matching_df['新JANコード'].isin(kao_planet_new_jans)].copy()
        
        # 旧JANで重複しているものを削除
        matching_df = matching_df[~matching_df['旧JANコード'].isin(kao_planet_old_jans)].copy()
        
        removed_count = before_count - len(matching_df)
        if removed_count > 0:
            print(f"  ✂️ マッチング→今週花王・プラネット重複削除: {removed_count}件")
    
    # ========== ステップ4: 3つのデータを結合（優先順位順） ==========
    all_data = pd.concat([
        existing_df,      # 1位（最優先）
        kao_planet_df,    # 2位
        matching_df       # 3位（最低優先）
    ], ignore_index=True)
    
    print(f"  統合後: {len(all_data)}件")
    
    # ========== ステップ5: 新JANコードで重複削除 ==========
    before_dedup = len(all_data)
    all_data = all_data.drop_duplicates(subset=['新JANコード'], keep='first')
    after_dedup = len(all_data)
    
    removed_by_dedup = before_dedup - after_dedup
    if removed_by_dedup > 0:
        print(f"  🗑️ 新JAN重複削除: {removed_by_dedup}件")
    
    # ========== ステップ6: 旧JAN=新JANのものを削除 ==========
    before_same_jan = len(all_data)
    all_data = all_data[all_data['旧JANコード'] != all_data['新JANコード']].copy()
    after_same_jan = len(all_data)
    
    removed_same_jan = before_same_jan - after_same_jan
    if removed_same_jan > 0:
        print(f"  🔄 同一JAN削除: {removed_same_jan}件")
    
    print(f"  ✅ 最終件数: {len(all_data)}件")
    
    return all_data


# ========================================================================
# メイン処理
# ========================================================================

def main():
    """
    累積リスト統合メイン処理
    """
    print("=" * 60)
    print("累積リスト統合プログラム - 週次更新（完全版）")
    print("=" * 60)
    
    # ========== ステップ1: 累積リスト読み込み ==========
    print("\n【ステップ1】累積リストの読み込み")
    
    root = tk.Tk()
    root.withdraw()
    
    if messagebox.askyesno("累積リスト", "前週の累積リストがありますか？\n（初回は「いいえ」）"):
        existing_path = filedialog.askopenfilename(
            title="前週の累積リストを選択",
            filetypes=[
                ("Excel", "*.xlsx *.xls *.xlsm"),
                ("CSV", "*.csv"),
                ("すべて", "*.*")
            ]
        )
        
        if existing_path:
            try:
                existing_df = load_file_flexible(existing_path)
                # 既存データは正規化しない（データソース情報を保持）
                existing_df = clean_jan_codes(existing_df)
                
                # データソース列がない場合は追加
                if 'データソース' not in existing_df.columns:
                    print("  ⚠️ 「データソース」列がないため追加します")
                    existing_df['データソース'] = '累積（旧版）'
                
                print(f"✅ 累積リスト読み込み: {len(existing_df)}件")
                
            except Exception as e:
                messagebox.showerror("エラー", f"累積リスト読み込み失敗:\n{e}")
                root.destroy()
                return
        else:
            existing_df = pd.DataFrame()
    else:
        existing_df = pd.DataFrame()
        print("📂 新規作成モード")
    
    root.destroy()
    
    # ========== ステップ2: マッチング結果読み込み ==========
    print("\n【ステップ2】マッチング結果の読み込み")
    
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo("選択", "今週のマッチング結果ファイルを選択してください\n（人間が修正済みのもの）")
    
    matching_path = filedialog.askopenfilename(
        title="マッチング結果を選択",
        filetypes=[
            ("Excel", "*.xlsx *.xls *.xlsm"),
            ("CSV", "*.csv"),
            ("TSV", "*.tsv"),
            ("すべて", "*.*")
        ]
    )
    
    root.destroy()
    
    if not matching_path:
        messagebox.showwarning("キャンセル", "マッチング結果が選択されませんでした")
        return
    
    try:
        matching_df = load_file_flexible(matching_path)
        matching_df = normalize_columns(matching_df, 'マッチング')
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
    
    # 花王・プラネットは任意（半年に1回程度）
    choice = messagebox.askquestion(
        "花王・プラネット", 
        "今週、花王・プラネットの差し替えリストはありますか？\n"
        "（半年に1回程度の更新）\n\n"
        "「はい」→ ファイルを選択\n"
        "「いいえ」→ スキップ（マッチングのみ統合）",
        icon='question'
    )
    
    if choice == 'yes':
        kao_planet_path = filedialog.askopenfilename(
            title="花王・プラネット差し替えリストを選択",
            filetypes=[
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("すべて", "*.*")
            ]
        )
        
        if kao_planet_path:
            try:
                kao_planet_df = load_file_flexible(kao_planet_path)
                kao_planet_df = normalize_columns(kao_planet_df, '花王・プラネット')
                kao_planet_df = clean_jan_codes(kao_planet_df)
                print(f"✅ 花王・プラネット読み込み: {len(kao_planet_df)}件")
            except Exception as e:
                messagebox.showerror("エラー", f"花王・プラネット読み込み失敗:\n{e}")
                root.destroy()
                return
        else:
            print("📂 ファイル未選択（スキップ）")
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
    
    # タイムスタンプ付きファイル名
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_csv = output_dir / f"累積_差し替えリスト_{timestamp}.csv"
    output_excel = output_dir / f"累積_差し替えリスト_{timestamp}.xlsx"
    
    # 最新版（タイムスタンプなし）
    latest_csv = output_dir / "累積_差し替えリスト_最新.csv"
    latest_excel = output_dir / "累積_差し替えリスト_最新.xlsx"
    
    # ========== ステップ5: データ統合 ==========
    print("\n【ステップ5】データ統合・重複削除")
    
    final_df = remove_duplicates_advanced(existing_df, kao_planet_df, matching_df)
    
    # ========== ステップ6: 保存 ==========
    print("\n【ステップ6】ファイル保存")
    
    try:
        # CSV保存（cp932エンコード）
        print(f"💾 CSV保存中（cp932）: {output_csv.name}")
        final_df.to_csv(output_csv, index=False, encoding='cp932', errors='replace')
        final_df.to_csv(latest_csv, index=False, encoding='cp932', errors='replace')
        
        # Excel保存
        print(f"💾 Excel保存中: {output_excel.name}")
        final_df.to_excel(output_excel, index=False, engine='openpyxl')
        final_df.to_excel(latest_excel, index=False, engine='openpyxl')
        
        print("✅ 保存完了")
        
    except Exception as e:
        messagebox.showerror("エラー", f"ファイル保存失敗:\n{e}")
        return
    
    # ========== ステップ7: 完了メッセージ ==========
    
    # データソース別の集計
    source_counts = final_df['データソース'].value_counts().to_dict() if 'データソース' in final_df.columns else {}
    
    summary = f
    """
    統合処理完了
    
    【累積データ】
    総件数: {len(final_df)}件
    【内訳】
    既存累積: {len(existing_df)}件
    花王・プラネット（今週）: {len(kao_planet_df)}件
    マッチング（今週）: {len(matching_df)}件
    
    【データソース別】
    {chr(10).join([f'{k}: {v}件' for k, v in source_counts.items()])}
    
    【保存先】
    {output_dir}
    
    【ファイル】
    - {output_csv.name}（cp932）
    - {output_excel.name}
    - 累積_差し替えリスト_最新.csv
    - 累積_差し替えリスト_最新.xlsx
    
    """
    
    print(summary)
    messagebox.showinfo("完了", summary)


# ========================================================================
# 実行
# ========================================================================

if __name__ == "__main__":
    main()