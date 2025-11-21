# -*- coding: utf-8 -*-
"""
花王・プラネットの商品差し替えリスト作成スクリプト（修復機能改善版）
author : HibiKeita
"""

from pathlib import Path
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Excelの自動修復に使うライブラリやで！
try:
    import win32com.client as win32
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    print("⚠️ win32comライブラリが見つかりません。Excelの自動修復機能は無効になります。")
    print("もし修復が必要なエラーが出たら、'pip install pywin32' でインストールしてや！")

# --- 0. 設定と初期化 ---

# 環境変数を使ってデフォルト保存先を指定します
ROOT_DIR = Path(os.path.expanduser("~")) / "Box/D0RM_RM_130_リテールテクノロジー研究部/新/103_棚割/002_Allo/001_社内/002_マニュアル関連/差し替えリスト/差し替えリスト出力先（バックアップ版）"

# --- Excel自動修復関数 ---
def repair_and_resave_excel(file_path):
    """
    WindowsのExcelアプリケーションを起動し、破損したExcelファイルを
    開いて修復し、上書き保存する
    """
    if not WIN32COM_AVAILABLE:
        return False
        
    p_file = Path(file_path)
    print(f"🛠️ Excelを起動して、ファイルを自動修復しています: {p_file.name}")
    
    excel = None
    
    try:
        pythoncom.CoInitialize()
    except:
        pass 
        
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        workbook = excel.Workbooks.Open(str(p_file.resolve()), UpdateLinks=False, ReadOnly=False)
        workbook.Save()
        workbook.Close(SaveChanges=False)
        
        print(f"✅ 修復と再保存が完了しました: {p_file.name}")
        return True
        
    except Exception as e:
        print(f"❌ 自動修復に失敗しました: {p_file.name} - エラー: {e}")
        return False
        
    finally:
        if excel is not None:
            excel.Quit()
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass 

# --- GUIでファイルを選択する関数 ---
def select_files(title, filetypes, multiple=False):
    root = tk.Tk()
    root.withdraw()

    file_paths = []
    if multiple:
        file_paths = filedialog.askopenfilenames(
            title=title,
            filetypes=filetypes
        )
    else:
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes
        )
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

# --- 改善版：修復失敗時も通常読み込みにフォールバック ---
def load_with_repair(path, **read_excel_kwargs):
    """
    Excelファイルを読み込む
    流れ：
    1. 通常通り読み込みを試みる
    2. 失敗したら修復を試みて再度読み込み
    3. 修復失敗または修復不可なら、ユーザーにエラー表示
    """
    p_file = Path(path)
    
    # ステップ1: 通常通り読み込みを試みる
    try:
        print(f"📖 ファイルを読み込み中: {p_file.name}")
        df = pd.read_excel(p_file, engine='openpyxl', **read_excel_kwargs)
        print(f"✅ 読み込み成功（修復不要）")
        return df
    except Exception as e:
        print(f"⚠️ 通常読み込み失敗: {e}")
        
        # ステップ2: 修復を試みる
        if WIN32COM_AVAILABLE:
            print(f"🔧 修復を試みています...")
            repair_success = repair_and_resave_excel(p_file)
            
            if repair_success:
                # 修復成功したら再度読み込み
                try:
                    print(f"📖 修復後のファイルを読み込み中...")
                    df = pd.read_excel(p_file, engine='openpyxl', **read_excel_kwargs)
                    print(f"✅ 修復・読み込み成功")
                    return df
                except Exception as retry_e:
                    print(f"❌ 修復後も読み込み失敗: {retry_e}")
                    messagebox.showerror("読み込みエラー", 
                        f"ファイル '{p_file.name}' の修復と読み込みに失敗しました。\n\n"
                        f"エラー詳細: {str(retry_e)}\n\n"
                        f"ファイルが破損している可能性があります。")
                    raise
            else:
                print(f"❌ 修復処理に失敗しました")
                messagebox.showerror("修復失敗", 
                    f"ファイル '{p_file.name}' の修復に失敗しました。\n\n"
                    f"ファイルが破損している可能性があります。")
                raise
        else:
            # win32comが無い場合
            print(f"⚠️ win32comが利用不可のため、修復できません")
            messagebox.showerror("修復不可", 
                f"ファイル '{p_file.name}' の読み込みに失敗しました。\n\n"
                f"修復ツール（win32com）が利用不可です。\n"
                f"管理者に連絡してください。")
            raise

# --- 1. 花王データ読み込み関数 ---
def load_kao(path):
    """改善版：修復失敗時もエラーで明示"""
    try:
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
    except Exception as e:
        print(f"❌ 花王ファイルの処理に失敗: {path.name}")
        raise

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

# --- 5. プラネット差し替えリスト生成 ---
def process_planet_diff(planet_paths_dict):
    """改善版：各ファイル処理でエラーをキャッチ、続行"""
    result = []
    
    for season, paths in planet_paths_dict.items():
        try:
            print(f"\n【{season}の処理】")
            
            # 新規品の読み込み
            new_df = load_with_repair(
                paths['new'],
                dtype={'ＪＡＮコード': str, '旧ＪＡＮコード': str}
            )
            new_df['備考'] = paths['new'].name
            
            # 廃番品の読み込み
            disc_df = load_with_repair(
                paths['disc'],
                dtype={'JANコード': str, '新JANコード': str, '廃番予定品': str, '新商品名': str}
            )
            disc_df['備考'] = paths['disc'].name
            
            # 花王関連のデータを除外
            new_df = exclude_kao(new_df, 'メーカーコード')
            disc_df = exclude_kao(disc_df, 'メーカー')
            
            # 以下、既存の処理と同じ
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
            
            combined_planet_diff_with_notes = pd.merge(
                combined_planet_diff, 
                new_df[['JANコード', '備考']], 
                left_on='新JAN', 
                right_on='JANコード', 
                how='left'
            )
            combined_planet_diff_with_notes = combined_planet_diff_with_notes.drop(columns='JANコード').rename(columns={'備考': '新JAN備考'})
            
            result.append(combined_planet_diff_with_notes)
            print(f"✅ {season}の処理完了（{len(combined_planet_diff_with_notes)}件）")
            
        except Exception as e:
            print(f"❌ {season}の処理に失敗しました。スキップします。")
            continue
    
    return pd.concat(result, ignore_index=True) if result else pd.DataFrame()

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
    # --- ファイル選択UI ---
    messagebox.showinfo("ファイル選択", "花王の上期新規品・廃止品リスト（複数選択可）を選んでください。", icon='info')
    kao_upper_period_file_paths = select_files("花王の上期新規品・廃止品リストを選択", [("Excelファイル", "*.xlsm *.xlsx")], multiple=True)
    if not kao_upper_period_file_paths:
        messagebox.showwarning("処理中断", "花王の上期ファイルが選択されませんでした。処理を中断します。", icon='warning')
        return

    messagebox.showinfo("ファイル選択", "花王の下期新規品・廃止品リスト（複数選択可）を選んでください。", icon='info')
    kao_lower_period_file_paths = select_files("花王の下期新規品・廃止品リストを選択", [("Excelファイル", "*.xlsm *.xlsx")], multiple=True)
    if not kao_lower_period_file_paths:
        messagebox.showwarning("処理中断", "花王の下期ファイルが選択されませんでした。処理を中断します。", icon='warning')
        return
    
    # プラネットの新規品・廃番品ファイルを期間ごとに選択し、辞書に格納します。
    planet_paths_selected = {}
    
    # 上期プラネットファイルの選択
    messagebox.showinfo("ファイル選択", "プラネットの上期新規品リストを選択してください。\n例: 新製品リスト（上期）", icon='info')
    new_planet_path_upper = select_files("上期 プラネット新製品リスト", [("Excelファイル", "*.xlsx")])
    if new_planet_path_upper:
        planet_paths_selected["上期"] = {"new": new_planet_path_upper[0]}
    else:
        messagebox.showwarning("処理中断", "プラネットの上期新規品リストが選択されませんでした。処理を中断します。", icon='warning')
        return

    messagebox.showinfo("ファイル選択", "プラネットの上期廃番品リストを選択してください。\n例: 廃番品リスト（上期）", icon='info')
    disc_planet_path_upper = select_files("上期 プラネット廃番品リスト", [("Excelファイル", "*.xlsx")])
    if disc_planet_path_upper:
        planet_paths_selected["上期"]["disc"] = disc_planet_path_upper[0]
    else:
        messagebox.showwarning("処理中断", "プラネットの上期廃番品リストが選択されませんでした。処理を中断します。", icon='warning')
        return

    # 下期プラネットファイルの選択
    messagebox.showinfo("ファイル選択", "プラネットの下期新規品リストを選択してください。\n例: 新製品リスト（下期）", icon='info')
    new_planet_path_lower = select_files("下期 プラネット新製品リスト", [("Excelファイル", "*.xlsx")])
    if new_planet_path_lower:
        planet_paths_selected["下期"] = {"new": new_planet_path_lower[0]}
    else:
        messagebox.showwarning("処理中断", "プラネットの下期新規品リストが選択されませんでした。処理を中断します。", icon='warning')
        return

    messagebox.showinfo("ファイル選択", "プラネットの下期廃番品リストを選択してください。\n例: 廃番品リスト（下期）", icon='info')
    disc_planet_path_lower = select_files("下期 プラネット廃番品リスト", [("Excelファイル", "*.xlsx")])
    if disc_planet_path_lower:
        planet_paths_selected["下期"]["disc"] = disc_planet_path_lower[0]
    else:
        messagebox.showwarning("処理中断", "プラネットの下期廃番品リストが選択されませんでした。処理を中断します。", icon='warning')
        return

    # 全ての選択された花王ファイルを結合
    all_kao_file_paths = kao_upper_period_file_paths + kao_lower_period_file_paths
    
    # 処理するデータがない場合は中断
    if not all_kao_file_paths and not planet_paths_selected:
        messagebox.showwarning("データなし", "処理するファイルが一つも選択されませんでした。", icon='warning')
        return

    # 出力フォルダを選択
    output_dir = select_output_folder("結果を保存するフォルダを選択してください")
    if not output_dir:
        output_dir = ROOT_DIR  # キャンセル時はROOT_DIRにデフォルト保存
        messagebox.showinfo("キャンセル", f"デフォルト保存先を使用します。{output_dir}")

    # --- データ処理開始 ---
    combined_df = pd.DataFrame()

    if all_kao_file_paths:
        kao_df = pd.concat([load_kao(p) for p in all_kao_file_paths], ignore_index=True)
        kao_df = kao_df.rename(columns={'備考': '新JAN備考'})
        combined_df = pd.concat([combined_df, kao_df], ignore_index=True)

    if planet_paths_selected:
        planet_diff_df = process_planet_diff(planet_paths_selected) 
        combined_df = pd.concat([combined_df, planet_diff_df], ignore_index=True)

    if combined_df.empty:
        messagebox.showwarning("データなし", "結合できるデータが一つもありませんでした。", icon='warning')
        return

    final_df = finalize(combined_df)
    
    final_df.to_csv(output_dir / "花王・プラネット差し替えリスト完成版.csv", index=False, encoding='cp932', errors='replace')
    final_df.to_excel(output_dir / "花王・プラネット差し替えリスト完成版.xlsx", index=False, engine='openpyxl')
    
    messagebox.showinfo("完了", f"🎉 差し替えリスト作成完了！CSVとExcelを出力しました。\n出力先: {output_dir}", icon='info')
    print("花王とプラネットのデータ統合と、差し替えリストの生成が完了しました。")

# スクリプトが直接実行された場合にのみ、main関数を呼び出します。
if __name__ == '__main__':
    main()