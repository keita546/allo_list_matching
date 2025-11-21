# -*- coding: utf-8 -*-
"""
花王・プラネットの商品差し替えリスト作成スクリプト（構造最適化版_プラネット結合対応_エンジン指定_GUI対応）
author : HibiKeita
"""

from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# --- 0. 設定と初期化 ---

# スクリプトのルートディレクトリを指定します。
# このパスは、ファイル選択がキャンセルされた場合などの、出力ファイルのデフォルト基準ディレクトリとして使用します。
ROOT_DIR = Path(os.path.expanduser("~")/"Box/D0RM_RM_130_リテールテクノロジー研究部/新/103_棚割/002_Allo/001_社内/002_マニュアル関連/差し替えリスト/差し替えリスト出力先（バックアップ版）") 

# --- GUIでファイルを選択する関数 ---
# ファイル選択ダイアログを表示し、ユーザーにファイルパスを選択させる機能を提供します。
# 理由：ファイルパスをコードに直接記述せず、実行時に柔軟にユーザーが指定できるようにするためです。
def select_files(title, filetypes, multiple=False):
    # Tkinterのルートウィンドウを作成します。これはダイアログの親となりますが、表示はしません。
    # 理由：ファイル選択のみが目的であり、余分なウィンドウ表示を避けるためです。
    root = tk.Tk()
    root.withdraw() # ウィンドウを非表示にします。

    file_paths = []
    if multiple:
        # 複数ファイル選択ダイアログを表示します。
        # 理由：複数のファイルを一度に選択し、まとめて処理したい場合に利用するためです。
        file_paths = filedialog.askopenfilenames(
            title=title,
            filetypes=filetypes
        )
    else:
        # 単一ファイル選択ダイアログを表示します。
        # 理由：単一のファイルを選択したい場合に、操作を簡略化するためです。
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes
        )
        if file_path: # ファイルが選択された場合のみリストに追加します。
            file_paths = [file_path]
    
    root.destroy() # ウィンドウを破棄し、メモリリソースを解放します。
    return [Path(p) for p in file_paths] # Pathオブジェクトに変換し、ファイルパス操作を容易にします。

# --- 出力フォルダを選択する関数 ---
# 結果のCSV/Excelファイルを保存するフォルダをユーザーに選択させます。
# 理由：出力先を固定せず、ユーザーが毎回自由に保存先を指定できるようにするためです。
def select_output_folder(title="結果を保存するフォルダを選択してください"):
    # Tkinterのルートウィンドウを作成し、非表示にします。
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=title) # フォルダ選択ダイアログを表示します。
    root.destroy() # ウィンドウを破棄します。
    # フォルダが選択された場合はPathオブジェクトを返し、キャンセルされた場合はNoneを返します。
    return Path(folder_path) if folder_path else None

# --- 1. 花王データ読み込み関数 ---
# 花王のExcelファイルを読み込み、必要な列を抽出・整形し、備考列を追加します。
# 理由：花王データ特有の形式（ヘッダ位置、列番号）に対応し、後続の処理に必要な形式に統一するためです。
def load_kao(path):
    # Excelファイルを読み込みます。
    # usecolsで必要な列（7, 15, 42, 44列目：0始まり）を選択し、skiprowsで最初の5行をスキップします。
    # 理由：花王ファイルの形式に基づき、データ本体のみを効率的に読み込むためです。
    # dtype={14: str, 41: str}により、JANコード列を文字列として読み込みます。
    # 理由：JANコードの先頭や末尾の0が数値型として扱われて消えるのを防ぐためです。
    df = pd.read_excel(path, usecols=[6, 14, 41, 43], skiprows=5, header=None, engine='openpyxl',
                        dtype={14: str, 41: str}) 
    
    # 読み込んだ列に分かりやすい名前を付けます。
    # 理由：コードの可読性を高め、後の処理でどの列が何を示すか一目でわかるようにするためです。
    df.columns = ['新商品名', '新JAN', '旧JAN', '旧商品名']
    
    # 旧JANと新JANのどちらかが欠損している行は除外します。
    # 理由：JANコードのペアがないと、差し替え情報として機能しないため、不要なデータを除去します。
    df = df.dropna(subset=['旧JAN', '新JAN'])[['旧JAN', '旧商品名', '新JAN', '新商品名']]

    # 元のファイル名を「備考」列として追加します。
    # 理由：データの出所を記録し、後のデータ追跡や重複チェック時に役立てるためです。
    df['備考'] = path.name # path.nameでファイル名のみを取得します。
    return df

# --- 2. プラネットクレンジング関数 ---
# プラネットの新規品・廃番品データを、花王データとの結合に向けて統一的に整形します。
# 理由：プラネットの新規品と廃番品でファイルの列構造が異なるため、この関数で共通の「旧JAN」「新JAN」構造に変換するためです。
def clean_planet(df, mode):
    # 列名に含まれる全角の「ＪＡＮ」を半角の「JAN」に統一します。
    # 理由：列名の表記ゆれを標準化し、後の処理での参照エラーやバグを未然に防ぐためです。
    df.columns = df.columns.str.replace('ＪＡＮ', 'JAN')
    
    if mode == 'discontinue':
        # 廃番品（discontinue）のデータを処理する場合です。
        # 必要な列を特定し、欠損値がある行を除外します。
        # 理由：差し替え情報（旧→新）として必須の要素が揃っているデータのみを抽出するためです。
        required_cols = ['JANコード', '新JANコード', '廃番予定品', '新商品名'] 
        df = df.dropna(subset=required_cols)
        
        # 列名を統一的な名前に変更します。
        # 理由：花王データや新規品データとの結合時に、列名の一貫性を保つためです。
        return df.rename(columns={
            'JANコード': '旧JAN',      # 廃番品（旧）のJANコードを旧JANとする
            '新JANコード': '新JAN',     # 廃番品（新）のJANコードを新JANとする
            '廃番予定品': '旧商品名',   # 廃番予定品の商品名を旧商品名とする
            '新商品名': '新商品名'    # 新しい商品名を新商品名とする
        })[['旧JAN', '旧商品名', '新JAN', '新商品名']] # 必要な列のみを抽出します。
    else: # mode == 'new'
        # 新規品（new）のデータを処理する場合です。
        # 差し替え情報として機能する、「JANコード（新）」と「旧JANコード」が揃っている行のみを抽出します。
        # 理由：新規品リストから、差し替え情報（旧JANコードが存在するもの）を特定するためです。
        df = df.dropna(subset=['JANコード', '旧JANコード']) 

        # 列名を統一的な名前に変更します。
        # 理由：新旧JANと新商品名の対応を明確にするためです。
        return df.rename(columns={'旧JANコード': '旧JAN', 'JANコード': '新JAN', '商品名全角': '新商品名'})[
            ['旧JAN', '新JAN', '新商品名']
        ]

# --- 3. 純粋新規品抽出 ---
# 新規品リスト（new_df）の中から、廃番品リスト（old_df）に旧商品として含まれていない「純粋な新規品」を抽出します。
# 理由：差し替えではない、真に市場に追加された新規品を特定し、重複なくリストに追加するためです。
def extract_unmatched(new_df, old_df):
    # 新規品リストの「新JAN」が、廃番品リストの「旧JAN」に含まれていない行を抽出します。
    # 理由：効率的なブールインデックス（~.isin()）を用いることで、「差し替え対象ではない新規品」を高速にフィルタリングするためです。
    add = new_df[~new_df['新JAN'].isin(old_df['旧JAN'])].copy()
    
    # 純粋新規品には対応する旧商品名がないため、列を空文字で初期化します。
    # 理由：後のデータ結合時に列の整合性を保ち、データ型の不一致を防ぐためです。
    add['旧商品名'] = ''
    return add[['旧JAN', '旧商品名', '新JAN', '新商品名']]

# --- 4. クレンジング前除外処理 ---
# プラネットのデータから、花王関連の商品を除外します。
# 理由：花王のデータは個別に読み込まれているため、プラネットリスト内の花王商品を排除することで、データの重複（ダブり）を避けるためです（MECEの原則）。
def exclude_kao(df, is_kao_col):
    # 指定された列（メーカーコードまたはメーカー名）が「4901301」（花王のメーカーコード）で始まるもの、または「花王株式会社」を含む行を**除外**します。
    # 理由：メーカーコードやメーカー名による確実な花王データのフィルタリングを行うためです。astype(str)は、列が数値型の場合でも文字列操作を可能にするためです。
    return df[~df[is_kao_col].astype(str).str.startswith('4901301') & ~df[is_kao_col].astype(str).str.contains('花王株式会社')]

# --- 5. プラネット差し替えリスト生成 ---
# 期間ごとのプラネット新規品・廃番品データを処理し、差し替えリストを生成します。
# 理由：期間ごとに異なるファイルからデータを読み込み、統一的なロジックで「差し替え情報」と「純粋新規品」を抽出・統合するためです。
def process_planet_diff(planet_paths_dict): # 選択されたファイルパスの辞書を受け取ります。
    result = []
    # 期間（上期、下期など）ごとにループ処理を行います。
    # 理由：各期間で同じ処理を自動化し、コードの記述量を削減するためです。
    for season, paths in planet_paths_dict.items():
        # 新規品と廃番品のExcelファイルを読み込みます。
        # 理由：JANコードの精度を保つため、関連列を文字列型（str）で読み込みます。
        new_df = pd.read_excel(paths['new'], engine='openpyxl',
                                 dtype={'ＪＡＮコード': str, '旧ＪＡＮコード': str})
        disc_df = pd.read_excel(paths['disc'], engine='openpyxl',
                                 dtype={'JANコード': str, '新JANコード': str, '廃番予定品': str, '新商品名': str})

        # 備考列としてファイル名を追加します。
        new_df['備考'] = paths['new'].name
        disc_df['備考'] = paths['disc'].name

        # 花王関連のデータを除外します。
        # 理由：データ重複の排除と、花王以外の差し替え情報に焦点を絞るためです。
        new_df = exclude_kao(new_df, 'メーカーコード')
        disc_df = exclude_kao(disc_df, 'メーカー')
        
        # プラネットのデータを整形（クリーンアップ）します。
        new_clean = clean_planet(new_df, 'new')
        disc_clean = clean_planet(disc_df, 'discontinue')

        # --- 新規品リスト優先の重複排除ロジック ---
        # 1. 新規品リストに新JANが記載されている廃番品データは、新規品リスト側を正として排除します。
        # 理由：新規品リストを優先的な差し替え情報ソースとみなし、廃番品リストとの重複を排除するためです。
        disc_not_in_new_by_new_jan = disc_clean[
            ~disc_clean['新JAN'].isin(new_clean['新JAN'])
        ].copy()

        # 2. 上記で残った廃番品データのうち、旧JANが新規品リストの旧JANと重複する場合も排除します。
        # 理由：新規品リストで既に「旧商品」として扱われているものを、廃番品リストから再度取り込まないようにするためです。
        final_disc_additions = disc_not_in_new_by_new_jan[
            ~disc_not_in_new_by_new_jan['旧JAN'].isin(new_clean['旧JAN'])
        ].copy()

        # 純粋な新規品を抽出します。
        # 理由：差し替えではない、真の新規品を特定するためです。
        pure_new_items = extract_unmatched(new_clean, disc_clean)

        # 「純粋な新規品」と「重複排除済みの廃番品からの差し替え情報」を結合します。
        # 理由：新規品優先の原則を保ちつつ、廃番品リストからの有効な情報も漏れなく取り込むためです。
        combined_planet_diff = pd.concat([pure_new_items, final_disc_additions], ignore_index=True)
        
        # 結合したデータに、元の新規品リストの備考列を結合し直します。
        # 理由：clean_planet処理で備考列が失われたため、どの新規品がどのファイルから来たかを記録し直すためです。
        # 結合キーには、元のnew_dfの'JANコード'とcombined_planet_diffの'新JAN'を使用します。
        combined_planet_diff_with_notes = pd.merge(combined_planet_diff, new_df[['JANコード', '備考']], 
                                                   left_on='新JAN', right_on='JANコード', how='left')
        combined_planet_diff_with_notes = combined_planet_diff_with_notes.drop(columns='JANコード').rename(columns={'備考': '新JAN備考'})
        
        result.append(combined_planet_diff_with_notes)
        
    return pd.concat(result, ignore_index=True)

# --- 6. クリーンアップ処理 ---
# 最終的なデータフレームの整形、JANコードのクリーンアップ、重複排除などを行います。
# 理由：最終的な出力形式の統一、JANコードのデータ品質保証、冗長なデータの排除を行うためです。
def finalize(df):
    # 列名を最終出力用に統一します。
    df = df.rename(columns={'旧JAN': '旧JANコード', '新JAN': '新JANコード'})
    
    # JANコード列をクリーンアップします。
    # 理由：JANコードは13桁の数値として統一的に扱われるべきですが、元のファイルでは文字列や数値、非数字文字が混入することがあるため、データ品質を保証するためです。
    for col in ['旧JANコード', '新JANコード']:
        df[col] = (df[col].astype(str)                                   # 全て文字列に変換。
                            .str.replace(r'\D+', '', regex=True)         # 正規表現で数字以外の文字を全て除去。
                            .replace('', pd.NA)                          # 数字除去の結果空文字になったものを欠損値に変換。
                            # 欠損値でなければ13桁に0埋め（zfill(13)）し、先頭から13文字を切り出し（[:13]）て確実に13桁にします。
                            .apply(lambda x: str(x).zfill(13)[:13] if pd.notna(x) else pd.NA)) 

    # 商品名の空文字を「該当文字列なし」に置換します。
    # 理由：空白のままよりも、「情報がない」ことを明示することで、ユーザーの誤解を防ぎ可読性を向上させるためです。
    df['旧商品名'] = df['旧商品名'].replace('', '該当文字列なし')
    df['新商品名'] = df['新商品名'].replace('', '該当文字列なし')

    # 旧JANコードと新JANコードが同じ行は除外します。
    # 理由：差し替えリストの目的である「何が何に変わったか」という情報を満たさない、自己差し替えの行は不要なデータであるためです。
    return df[df['旧JANコード'] != df['新JANコード']].drop_duplicates() # 重複行を削除します。

# --- 7. メイン処理 ---
# スクリプトの主要な処理の流れを管理し、全体を統括します。
def main():
    # --- ファイル選択UI ---
    # 花王の上期ファイルパスを選択
    messagebox.showinfo("ファイル選択", "花王の上期新規品・廃止品リスト（複数選択可）を選んでください。", icon='info')
    kao_upper_period_file_paths = select_files("花王の上期新規品・廃止品リストを選択", [("Excelファイル", "*.xlsm *.xlsx")], multiple=True)
    if not kao_upper_period_file_paths:
        messagebox.showwarning("処理中断", "花王の上期ファイルが選択されませんでした。処理を中断します。", icon='warning')
        return

    # 花王の下期ファイルパスを選択
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

    # 出力フォルダを選択
    output_dir = select_output_folder("結果を保存するフォルダを選択してください")
    if not output_dir:
        output_dir = ROOT_DIR #キャンセル時はROOT_DIRにデフォルト保存
        messagebox.showinfo("キャンセル", f"デフォルト保存先を使用します。{output_dir}")
        return

    # --- データ処理開始 ---
    combined_df = pd.DataFrame() # 空のデータフレームを初期化します。

    if all_kao_file_paths:
        # 花王の全てのファイルを読み込み、結合します。
        # 理由：複数の花王データを一括で処理し、後のプラネットデータとの統合に備えるためです。
        kao_df = pd.concat([load_kao(p) for p in all_kao_file_paths], ignore_index=True)
        # 備考列名を統一します。
        kao_df = kao_df.rename(columns={'備考': '新JAN備考'})
        combined_df = pd.concat([combined_df, kao_df], ignore_index=True) # 結合します。

    if planet_paths_selected:
        # プラネットの差し替えリストを生成します。
        # 理由：選択されたファイル情報に基づき、複雑な重複排除ロジックを含むプラネットデータの処理を実行するためです。
        planet_diff_df = process_planet_diff(planet_paths_selected) 
        combined_df = pd.concat([combined_df, planet_diff_df], ignore_index=True) # 結合します。

    if combined_df.empty:
        messagebox.showwarning("データなし", "結合できるデータが一つもありませんでした。", icon='warning')
        return

    # 最終的なデータクリーンアップと整形を行います。
    # 理由：データの品質保証と、最終的な出力形式への調整（JANコード整形、重複・自己差し替えの削除）のためです。
    final_df = finalize(combined_df)
    
    # 完成した差し替えリストをCSVファイルとして出力します。
    # encoding='cp932'とerrors='replace'は、日本語環境での文字化けを防ぎ、特殊文字があってもエラーにならずに出力するための設定です。
    final_df.to_csv(output_dir / "花王・プラネット差し替えリスト完成版.csv", index=False, encoding='cp932', errors='replace')

    # Excelファイルとしても出力します。
    # 理由：CSVだけでなくExcel形式でも提供することで、ユーザーの利用環境に応じた柔軟性を持たせるためです。
    final_df.to_excel(output_dir / "花王・プラネット差し替えリスト完成版.xlsx", index=False, engine='openpyxl')
    
    messagebox.showinfo("完了", f"🎉 差し替えリスト作成完了！CSVとExcelを出力しました。\n出力先: {output_dir}", icon='info')
    print("花王とプラネットのデータ統合と、差し替えリストの生成が完了しました。")

# スクリプトが直接実行された場合にのみ、main関数を呼び出します。
# 理由：このスクリプトをモジュールとして他のファイルからインポートした場合、main関数が意図せず実行されるのを防ぐためです。
if __name__ == '__main__':
    main()