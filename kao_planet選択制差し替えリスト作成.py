# -*- coding: utf-8 -*-
"""
花王・プラネットの商品差し替えリスト作成スクリプト（柔軟版）
- 花王のみ OK
- プラネットのみ OK  
- 花王+プラネット OK
- 上期のみ/下期のみも OK

author : HibiKeita
"""

from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# --- 0. 設定と初期化 ---
ROOT_DIR = Path(os.path.expanduser("~")/"Box/D0RM_RM_130_リテールテクノロジー研究部/新/103_棚割/002_Allo/001_社内/002_マニュアル関連/差し替えリスト/差し替えリスト出力先（バックアップ版）") 

# --- GUIでファイルを選択する関数 ---
def select_files(title, filetypes, multiple=False):
    root = tk.Tk()
    root.withdraw()

    file_paths = []
    if multiple:
        file_paths = filedialog.askopenfilenames(title=title, filetypes=filetypes)
    else:
        file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if file_path:
            file_paths = [file_path]
    
    root.destroy()
    return [Path(p) for p in file_paths]

# --- 出力フォルダを選択する関数 ---
def select_output_folder(title="結果を保存するフォルダを選択してください"):
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=title)
    root.destroy()
    return Path(folder_path) if folder_path else None

# --- 1. 花王データ読み込み関数（修復機能付き） ---
def load_kao(path):
    df = load_with_repair(
        path, 
        usecols=[6, 14, 41, 43], 
        skiprows=5, 
        header=None,
        dtype={14: str, 41: str}
    )
    df.columns = ['新商品名', '新JAN', '旧JAN', '旧商品名']
    df = df.dropna(subset=['旧JAN', '新JAN'])[['旧JAN', '旧商品名', '新JAN', '新商品名']]
    df['備考'] = path.name
    return df

# --- 2. プラネットクレンジング関数 ---
def clean_planet(df, mode):
    df.columns = df.columns.str.replace('ＪＡＮ', 'JAN')
    
    if mode == 'discontinue':
        required_cols = ['JANコード', '新JANコード', '廃番予定品', '新商品名'] 
        df = df.dropna(subset=required_cols)
        return df.rename(columns={
            'JANコード': '旧JAN',
            '新JANコード': '新JAN',
            '廃番予定品': '旧商品名',
            '新商品名': '新商品名'
        })[['旧JAN', '旧商品名', '新JAN', '新商品名']]
    else: # mode == 'new'
        df = df.dropna(subset=['JANコード', '旧JANコード']) 
        return df.rename(columns={'旧JANコード': '旧JAN', 'JANコード': '新JAN', '商品名全角': '新商品名'})[
            ['旧JAN', '新JAN', '新商品名']
        ]

# --- 3. 純粋新規品抽出 ---
def extract_unmatched(new_df, old_df):
    add = new_df[~new_df['新JAN'].isin(old_df['旧JAN'])].copy()
    add['旧商品名'] = ''
    return add[['旧JAN', '旧商品名', '新JAN', '新商品名']]

# --- 4. クレンジング前除外処理 ---
def exclude_kao(df, is_kao_col):
    return df[~df[is_kao_col].astype(str).str.startswith('4901301') & ~df[is_kao_col].astype(str).str.contains('花王株式会社')]

# --- 5. プラネット差し替えリスト生成（修復機能付き） ---
def process_planet_diff(planet_paths_dict):
    result = []
    for season, paths in planet_paths_dict.items():
        # 新規品と廃番品の両方が揃っている場合のみ処理
        if 'new' not in paths or 'disc' not in paths:
            print(f"⚠️ {season}のプラネットデータが不完全です（スキップ）")
            continue
        
        try:
            new_df = load_with_repair(
                paths['new'],
                dtype={'ＪＡＮコード': str, '旧ＪＡＮコード': str}
            )
            disc_df = load_with_repair(
                paths['disc'],
                dtype={'JANコード': str, '新JANコード': str, '廃番予定品': str, '新商品名': str}
            )

            new_df['備考'] = paths['new'].name
            disc_df['備考'] = paths['disc'].name

            new_df = exclude_kao(new_df, 'メーカーコード')
            disc_df = exclude_kao(disc_df, 'メーカー')
            
            new_clean = clean_planet(new_df, 'new')
            disc_clean = clean_planet(disc_df, 'discontinue')

            disc_not_in_new_by_new_jan = disc_clean[
                ~disc_clean['新JAN'].isin(new_clean['新JAN'])
            ].copy()

            final_disc_additions = disc_not_in_new_by_new_jan[
                ~disc_not_in_new_by_new_jan['旧JAN'].isin(new_clean['旧JAN'])
            ].copy()

            pure_new_items = extract_unmatched(new_clean, disc_clean)
            combined_planet_diff = pd.concat([pure_new_items, final_disc_additions], ignore_index=True)
            
            combined_planet_diff_with_notes = pd.merge(combined_planet_diff, new_df[['JANコード', '備考']], 
                                                       left_on='新JAN', right_on='JANコード', how='left')
            combined_planet_diff_with_notes = combined_planet_diff_with_notes.drop(columns='JANコード').rename(columns={'備考': '新JAN備考'})
            
            result.append(combined_planet_diff_with_notes)
            print(f"✅ {season}の処理完了（{len(combined_planet_diff_with_notes)}件）")
            
        except Exception as e:
            print(f"❌ {season}の処理に失敗しました: {e}")
            continue
    
    if result:
        return pd.concat(result, ignore_index=True)
    else:
        return pd.DataFrame()

# --- 6. クリーンアップ処理 ---
def finalize(df):
    df = df.rename(columns={'旧JAN': '旧JANコード', '新JAN': '新JANコード'})
    
    for col in ['旧JANコード', '新JANコード']:
        df[col] = (df[col].astype(str)
                            .str.replace(r'\D+', '', regex=True)
                            .replace('', pd.NA)
                            .apply(lambda x: str(x).zfill(13)[:13] if pd.notna(x) else pd.NA)) 

    df['旧商品名'] = df['旧商品名'].replace('', '該当文字列なし')
    df['新商品名'] = df['新商品名'].replace('', '該当文字列なし')

    return df[df['旧JANコード'] != df['新JANコード']].drop_duplicates()

# --- 7. メイン処理 ---
def main():
    print("=" * 60)
    print("花王・プラネット差し替えリスト作成（柔軟版）")
    print("=" * 60)
    
    # --- 花王データ選択（任意） ---
    all_kao_file_paths = []
    
    if messagebox.askyesno("花王データ", "花王の差し替えリストを処理しますか？"):
        # 上期
        if messagebox.askyesno("花王上期", "花王の上期ファイルはありますか？"):
            messagebox.showinfo("選択", "花王の上期新規品・廃止品リスト（複数選択可）", icon='info')
            kao_upper = select_files("花王上期を選択", [("Excelファイル", "*.xlsm *.xlsx")], multiple=True)
            all_kao_file_paths.extend(kao_upper)
        
        # 下期
        if messagebox.askyesno("花王下期", "花王の下期ファイルはありますか？"):
            messagebox.showinfo("選択", "花王の下期新規品・廃止品リスト（複数選択可）", icon='info')
            kao_lower = select_files("花王下期を選択", [("Excelファイル", "*.xlsm *.xlsx")], multiple=True)
            all_kao_file_paths.extend(kao_lower)
    
    print(f"📂 花王ファイル: {len(all_kao_file_paths)}件")
    
    # --- プラネットデータ選択（任意） ---
    planet_paths_selected = {}
    
    if messagebox.askyesno("プラネットデータ", "プラネットの差し替えリストを処理しますか？"):
        # 上期
        if messagebox.askyesno("プラネット上期", "プラネットの上期ファイルはありますか？"):
            messagebox.showinfo("選択", "プラネット上期新規品リスト", icon='info')
            new_upper = select_files("上期新製品リスト", [("Excelファイル", "*.xlsx")])
            
            if new_upper:
                messagebox.showinfo("選択", "プラネット上期廃番品リスト", icon='info')
                disc_upper = select_files("上期廃番品リスト", [("Excelファイル", "*.xlsx")])
                
                if disc_upper:
                    planet_paths_selected["上期"] = {"new": new_upper[0], "disc": disc_upper[0]}
                else:
                    messagebox.showwarning("スキップ", "上期廃番品が選択されなかったため、上期はスキップします")
        
        # 下期
        if messagebox.askyesno("プラネット下期", "プラネットの下期ファイルはありますか？"):
            messagebox.showinfo("選択", "プラネット下期新規品リスト", icon='info')
            new_lower = select_files("下期新製品リスト", [("Excelファイル", "*.xlsx")])
            
            if new_lower:
                messagebox.showinfo("選択", "プラネット下期廃番品リスト", icon='info')
                disc_lower = select_files("下期廃番品リスト", [("Excelファイル", "*.xlsx")])
                
                if disc_lower:
                    planet_paths_selected["下期"] = {"new": new_lower[0], "disc": disc_lower[0]}
                else:
                    messagebox.showwarning("スキップ", "下期廃番品が選択されなかったため、下期はスキップします")
    
    print(f"📂 プラネット期間: {len(planet_paths_selected)}期")
    
    # --- データがない場合は終了 ---
    if not all_kao_file_paths and not planet_paths_selected:
        messagebox.showwarning("データなし", "処理するファイルが1つも選択されませんでした", icon='warning')
        return

    # --- 出力フォルダ選択 ---
    output_dir = select_output_folder("結果を保存するフォルダを選択してください")
    if not output_dir:
        output_dir = ROOT_DIR
        messagebox.showinfo("デフォルト保存", f"デフォルト保存先: {output_dir}")

    # --- データ処理開始 ---
    combined_df = pd.DataFrame()

    # 花王データ処理
    if all_kao_file_paths:
        print("\n🔄 花王データ処理中...")
        kao_df = pd.concat([load_kao(p) for p in all_kao_file_paths], ignore_index=True)
        kao_df = kao_df.rename(columns={'備考': '新JAN備考'})
        combined_df = pd.concat([combined_df, kao_df], ignore_index=True)
        print(f"✅ 花王: {len(kao_df)}件")

    # プラネットデータ処理
    if planet_paths_selected:
        print("\n🔄 プラネットデータ処理中...")
        planet_diff_df = process_planet_diff(planet_paths_selected) 
        if not planet_diff_df.empty:
            combined_df = pd.concat([combined_df, planet_diff_df], ignore_index=True)
            print(f"✅ プラネット: {len(planet_diff_df)}件")
        else:
            print("⚠️ プラネットデータが生成されませんでした")

    if combined_df.empty:
        messagebox.showwarning("データなし", "結合できるデータがありませんでした", icon='warning')
        return

    # 最終処理
    print("\n🧹 最終クリーンアップ中...")
    final_df = finalize(combined_df)
    
    # 出力
    print(f"\n💾 ファイル出力中... 最終件数: {len(final_df)}件")
    final_df.to_csv(output_dir / "花王・プラネット差し替えリスト完成版.csv", index=False, encoding='cp932', errors='replace')
    final_df.to_excel(output_dir / "花王・プラネット差し替えリスト完成版.xlsx", index=False, engine='openpyxl')
    
    summary = f"""
🎉 差し替えリスト作成完了！

【処理内容】
花王: {len(all_kao_file_paths)}ファイル
プラネット: {len(planet_paths_selected)}期間
最終件数: {len(final_df)}件

【出力先】
{output_dir}

【ファイル】
- 花王・プラネット差し替えリスト完成版.csv
- 花王・プラネット差し替えリスト完成版.xlsx
"""
    
    print(summary)
    messagebox.showinfo("完了", summary, icon='info')

if __name__ == '__main__':
    main()