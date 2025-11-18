# -*- coding: utf-8 -*-
"""
Created by HIBI KEITA
修正版：suffix付きカラム名の統一、エラーハンドリング強化
"""

# ----------------------------------------
# 1. ライブラリのインポート 
# ----------------------------------------

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import os
# 処理: 文字列の類似度計算ライブラリをインポート
# 理由: 商品名称の曖昧突合処理を正確かつ高速に行うため。
from fuzzywuzzy import fuzz 


# ----------------------------------------
# 2. コアロジック関数
# ----------------------------------------

def calculate_similarity(s1: str, s2: str) -> float:
    """二つの文字列の類似度（0.0〜1.0）を計算します。"""
    # 処理: fuzz.ratioで得た0-100のスコアを100で割って正規化
    # 理由: 類似度スコアを標準的な0.0〜1.0の範囲に揃えるため。
    return fuzz.ratio(s1, s2) / 100.0


def load_data(file_path: str, suffix: str) -> pd.DataFrame:
    """ファイルを読み込み、suffixをカラムに付与する関数"""
    # 処理: pathlibを使って拡張子を取得し、小文字に変換
    # 理由: 拡張子（例: .XLSX）の大文字・小文字によるファイル判別エラーを防ぎ、安定性を高めるため。
    p = Path(file_path)
    ext = p.suffix.lower()
    
    if ext == '.csv':
        # 処理: CSVファイルをshift_jis エンコーディングで読み込み、不揃い行をスキップ
        # 理由: 日本語の文字化けを防ぎ、データ行の不整合があっても処理を続行できるようにするため。
        df = pd.read_csv(file_path, encoding='shift_jis', delimiter=',', on_bad_lines='skip')
    
    elif ext in ['.xlsx', '.xls']:
        # 処理: Excelファイルを読み込み（自動で適切なエンコーディングが適用される）
        # 理由: Excelファイルをサポートし、ユーザーが複数のファイル形式を使用できるようにするため。
        df = pd.read_excel(file_path)
        
    else:
        # 処理: サポート外のファイル形式の場合、例外を発生させて処理を停止
        # 理由: 不正なファイル形式での処理を防ぎ、ユーザーに明確なエラーメッセージを通知するため。
        raise ValueError(f"サポート外のファイル形式です:{ext}")
    
    # 処理: DataFrame内の文字列 'NULL' を Pandasの欠損値 pd.NA に置換
    # 理由: CSV内に文字列として存在する 'NULL' を、正しく欠損値として扱えるようにデータ品質を向上させるため。
    df = df.replace('NULL', pd.NA)

    # 処理: ユーザーのレポート出力とキー突合に必須のカラムを強制的に存在させる
    # 理由: データの欠損や読み込み時のエラーでカラム自体が消えることによるKeyErrorを絶対に防ぐため。
    required_cols = [
        'メーカー名称', 
        '目付', 
        'ブランド名称', 
        '標準分類名(クラス)',
        '商品名称（カナ）',
        'JANコード',
    ]
    
    for col in required_cols:
        if col not in df.columns:
            # 処理: 存在しない場合、pd.NA（欠損値）でカラムを生成
            # 理由: レポート出力時にKeyErrorを出さず、欠損データとして処理を進めるため。
            df[col] = pd.NA
    
    # 処理: デバッグ用：読み込み後のカラム名をコンソールに出力
    # 理由: どのカラムが実際に存在しているか確認し、エラーの原因を特定するため。
    print(f"Load data columns ({suffix}): {df.columns.tolist()}") 

    # 処理: 読み込んだDataFrameの全カラム名に接尾辞（_新または_旧）を付与
    # 理由: 後の突合処理で新旧のカラム名を明確に区別するため、コードの正確性を確保する。
    df = df.add_suffix(suffix)
    
    # 処理: デバッグ用：suffix付きのカラム名をコンソールに出力
    # 理由: add_suffix() が正しく機能しているか確認するため。
    print(f"After add_suffix ({suffix}): {df.columns.tolist()}")
    
    return df


def process_master_data(old_path: str, new_path: str, output_dir: str):
    """
    旧マスタと新マスタを突合し、リニューアル品を抽出するメインロジック関数
    """
    # 処理: 絶対一致させるキーカラムを定義
    # 理由: 実際のデータに存在する 'メーカーコード', '標準分類コード(クラス)', 'ブランドコード' をキーとして使用し、
    #       正確な商品照合を行うため。
    key_columns: list = ['メーカーコード', '標準分類コード(クラス)', 'ブランドコード'] 
    
    try:
        # 処理: 旧マスタと新マスタを読み込み、それぞれに接尾辞 _旧、_新 を付与
        # 理由: 両者のカラム名を明確に区別し、後の処理で混同を防ぐため。
        df_old = load_data(old_path, '_旧')
        df_new = load_data(new_path, '_新')
        
        # 処理: 新品DFに '発売日_新' カラムが存在しない場合に追加
        # 理由: ユーザー要望のレポートに '発売日_新' が必要なため、存在しない場合は空で補完してKeyErrorを防ぐため。
        if '発売日_新' not in df_new.columns:
            df_new['発売日_新'] = pd.NA
        
        # 処理: デバッグ用：process_master_data 内で使用するカラムが存在しているか確認
        # 理由: エラーの原因を特定するため。
        print(f"df_old columns: {df_old.columns.tolist()}")
        print(f"df_new columns: {df_new.columns.tolist()}")
        print(f"Checking for required columns:")
        print(f"  'JANコード_旧' in df_old: {'JANコード_旧' in df_old.columns}")
        print(f"  'JANコード_新' in df_new: {'JANコード_新' in df_new.columns}")
        print(f"  '商品名称（カナ）_旧' in df_old: {'商品名称（カナ）_旧' in df_old.columns}")
        print(f"  '商品名称（カナ）_新' in df_new: {'商品名称（カナ）_新' in df_new.columns}")

        # ----------------------------------------
        # 1. 曖昧突合処理の核となる関数（旧品の1行を基準に処理）
        # ----------------------------------------
        
        # 処理: 旧品DFの各行（Series）を受け取り、対応する新品候補の結果（Series）を返す関数を定義
        # 理由: df_old.apply()で各旧品レコードに対して個別に処理を実行できるようにするため。
        def find_best_matches_for_row(old_row: pd.Series) -> pd.Series:
            
            # 処理: 旧品のカナ名称を取得
            # 理由: これを基準に新品との類似度を計算するため。
            old_product_name_kana = old_row['商品名称（カナ）_旧']
            
            # 処理: デバッグ用：処理対象の旧品を表示
            # 理由: どのレコードで処理が止まっているか確認するため。
            print(f"Processing old row with JAN: {old_row.get('JANコード_旧', 'N/A')}, Name: {old_product_name_kana}")
            
            # 処理: 絶対キーで新品候補を絞り込む条件を初期化
            # 理由: 全新品DFから、現在の旧品行とキーが一致するものだけを抽出することで、
            #       処理の対象を限定し、効率と正確性を高めるため。
            filter_condition = pd.Series(True, index=df_new.index) 
            
            # 処理: キーカラムごとにフィルタ条件を追加（AND結合）
            # 理由: 複数のキー条件をすべて満たす新品レコードだけを抽出するため。
            for col in key_columns:
                old_key_col = col + '_旧'
                new_key_col = col + '_新'
                # 処理: カラムが実際に存在しているかチェック
                # 理由: 存在しないカラムを参照してKeyErrorが発生するのを防ぐため。
                if old_key_col in old_row.index and new_key_col in df_new.columns:
                    old_value = old_row[old_key_col]
                    print(f"    Filtering by {new_key_col}: {old_value}")
                    
                    # 処理: 旧品のキー値が欠損値（nan）でない場合のみフィルタリング
                    # 理由: nan == nan は常にFalseになるため、欠損値の場合はフィルタリングをスキップするため。
                    if pd.notna(old_value):
                        filter_condition &= (df_new[new_key_col] == old_value)
                    else:
                        print(f"      Skipping filter (old value is nan)")

            print(f"    Filtered df_new size: {filter_condition.sum()}")
            
            # 処理: フィルタリングされた新品候補から、商品名とJANコードを抽出
            # 理由: 類似度計算に必要なデータセットを用意し、重複を削除するため。
            new_candidates = df_new[filter_condition][['商品名称（カナ）_新', 'JANコード_新']].dropna().drop_duplicates()
            print(f"    new_candidates size: {len(new_candidates)}") 
            
            # 処理: 候補がない、または旧品カナ名が空の場合の処理
            # 理由: 照合結果を「新規/削除品」として明確に分類するため。
            if new_candidates.empty or pd.isna(old_product_name_kana):
                # 処理: 候補がない場合の結果Seriesを返す
                # 理由: 旧品が新マスタに存在しない、または削除された可能性があるため。
                return pd.Series({
                    '新品候補リスト': '候補なし（キー不一致 or カナ名空）', 
                    '最高類似度': 0.0,
                    '照合結果': '新規 or 削除品',
                    '判定': '新規 or 削除品',
                })

            # ---------------------
            # 類似度計算とソート
            # ---------------------
            # 処理: 全ての新品候補と旧品カナ名称の類似度を計算し、タプルリストに格納
            # 理由: 類似度スコアと商品情報をセットで保持し、後でソートを容易にするため。
            similarities = [
                (calculate_similarity(old_product_name_kana, row['商品名称（カナ）_新']), row['商品名称（カナ）_新'], row['JANコード_新'])
                for _, row in new_candidates.iterrows()
            ]
            # 処理: 類似度が高い順にソート（降順）
            # 理由: 最高類似度の候補を上位に持ってくるため。
            similarities.sort(key=lambda x: x[0], reverse=True)
            
            # 処理: 最高類似度の新品候補のJANコードを取得
            # 理由: 最適な新品レコードを特定するため。
            best_match_jan = similarities[0][2] 
            
            # 処理: JANコードが一致する新品レコードをdf_newから抽出 (Seriesとして取得)
            # 理由: 最高類似度の新品の全情報を取得し、旧品行に結合するため。
            best_new_row = df_new[df_new['JANコード_新'] == best_match_jan].iloc[0]

            # 処理: 必要な新品番情報をSeriesとして整形
            # 理由: レポートに必要なカラムを抽出し、ユーザーの要望に合わせてカラム名を変更するため。
            new_info_cols = [
                'JANコード_新', '商品名称（カナ）_新', 'メーカー名称_新', 
                '標準分類名(クラス)_新', 'ブランド名称_新', '目付_新', '発売日_新' 
            ]
            
            # 処理: 存在するカラムのみを抽出（存在しないカラムはスキップ）
            # 理由: KeyErrorを防ぐため。
            existing_new_info_cols = [col for col in new_info_cols if col in best_new_row.index]
            new_info_for_report = best_new_row[existing_new_info_cols].rename(
                # 処理: '標準分類名(クラス)_新' をユーザー要望の '標準分類(クラス)_新' にリネーム
                # 理由: レポートの表記を統一し、ユーザーの期待に合わせるため。
                {'標準分類名(クラス)_新': '標準分類(クラス)_新'}
            )
            
            print(f"  new_info_for_report: {new_info_for_report.to_dict()}")
            
            # 処理: 上位3つの候補を「商品名 (類似度%)」の形式で一つの文字列にまとめる
            # 理由: 複数の候補を分かりやすくレポートの1セルに表示し、ユーザーの確認を容易にするため。
            top_candidates = similarities[:3]
            candidate_list_str = " | ".join([
                f"{name} ({score:.1%})" for score, name, _ in top_candidates
            ])
            # 処理: 最高スコアを取得
            # 理由: 最高類似度を別途レポート出力する必要があるため。
            best_score = similarities[0][0] 
            
            # 処理: 照合結果を80%を基準に判定
            # 理由: 高い類似度の場合は信頼度が高いと判断し、低い場合は手動確認が必要と案内するため。
            if best_score >= 0.8: 
                result = '高類似度候補あり (80%以上)'
            else:
                result = '低類似度 (80%未満・要手動確認)'

            # 処理: 分析結果と最高候補の情報を結合して返す
            # 理由: 最終レポートに旧品情報と新品番情報を一元化するため。
            result_series = pd.Series({
                '新品候補リスト': candidate_list_str, 
                '最高類似度': best_score,
                '照合結果': result,
                '判定': result,
            })
            # 処理: 分析結果Seriesと新品情報Seriesを結合
            # 理由: 1つのSeriesとして返すことで、df_old.apply()で適切に結合できるようにするため。
            return pd.concat([result_series, new_info_for_report])

        # ----------------------------------------
        # 2. 結果の実行と結合
        # ----------------------------------------
        
        # 処理: 旧品DFの各行に対して、find_best_matches_for_row 関数を適用し、結果を新しいDFとして生成
        # 理由: ループ処理より効率的なPandasのベクトル化処理（apply）で、旧品リストの全レコードを一括処理するため。
        print(f"Starting analysis on {len(df_old)} old records...")
        analysis_result = df_old.apply(find_best_matches_for_row, axis=1)
        print(f"Analysis completed successfully!")

        # 処理: 旧品DFと分析結果DFを結合（インデックスが合うためそのまま結合）
        # 理由: 元の旧品リストに、新しい分析結果カラムと最高類似度の新品情報を追加するため。
        print(f"analysis_result shape: {analysis_result.shape}")
        print(f"analysis_result columns: {analysis_result.columns.tolist()}")
        print(f"df_old shape: {df_old.shape}")
        
        final_df = pd.concat([df_old, analysis_result], axis=1)
        print(f"final_df columns: {final_df.columns.tolist()}")


        # ----------------------------------------
        # 3. 結果の出力
        # ----------------------------------------
        
        # 処理: 最終レポート出力に合わせて、元の旧品カラム名も修正
        # 理由: '標準分類名(クラス)_旧' を ユーザー要望の '標準分類(クラス)_旧' にリネームするため。
        final_df = final_df.rename(columns={
            '標準分類名(クラス)_旧': '標準分類(クラス)_旧',
        })
        
        # 処理: 最終レポートに必要なカラムをユーザーの希望リストに基づいて定義
        # 理由: ユーザーの要望に完全に合致させ、MECEで分かりやすいレポート構造とするため。
        report_columns = [
            '照合結果', 'JANコード_旧','商品名称（カナ）_旧','メーカー名称_旧','標準分類(クラス)_旧', 'ブランド名称_旧','目付_旧','最高類似度','判定', 
            'JANコード_新','商品名称（カナ）_新','メーカー名称_新','標準分類(クラス)_新', 'ブランド名称_新','目付_新','発売日_新',
        ]
        
        # 処理: 出力ファイルパスを設定
        # 理由: 結果を保存する場所を指定するため。
        output_file_path = Path(output_dir) / 'リニューアル品抽出結果.csv'
        # 処理: 結果をCSVファイルとして保存（shift_jisエンコーディング）
        # 理由: 処理結果を永続化し、実務で使えるデータとしてアウトプットするため。
        final_df[report_columns].to_csv(output_file_path, index=False, encoding='shift_jis')
        
        return output_file_path

    except Exception as e:
        # 処理: データ処理中に発生したエラーを捕捉し、詳細情報を含むエラーを再発生
        # 理由: 予期せぬエラーによるアプリの強制終了を防ぎ、堅牢なエラーハンドリングを行うことでツールの信頼性を保つため。
        raise Exception(f"データ処理中にエラーが発生しました。詳細: {e}")


# ----------------------------------------
# 3. GUIクラスの定義
# ----------------------------------------
# 処理: Tkinterのメインアプリケーションをクラスとして定義
# 理由: ウィジェット（部品）や変数をクラス内で一元管理することで、コードの可読性とメンテナンス性を高めるため。

class MasterMatcherApp:
    
    def __init__(self, master):
        # 処理: アプリケーションの基本設定（タイトル設定と変数初期化）
        # 理由: アプリケーションが起動した際に、ウィンドウを作成し、部品配置のための準備をするため。
        self.master = master
        master.title("商品リニューアル品抽出")
        
        # 処理: ファイルパスを保持するためのTkinter専用の変数（StringVar）を定義
        # 理由: Entryウィジェットと値を自動で同期させ、入力されたパスを簡単に取得できるようにするため。
        self.old_path_var = tk.StringVar()
        self.new_path_var = tk.StringVar()
        # 処理: 出力先フォルダの初期値をデスクトップに設定
        # 理由: ユーザーが毎回フォルダを選択する手間を省くための、実用的な初期値設定。
        self.output_dir_var = tk.StringVar(value=os.path.join(os.path.expanduser('~'), 'Desktop'))
        
        self.create_widgets(master)

    # 処理: GUIの部品（ウィジェット）を画面上に配置する
    # 理由: 初期設定と画面部品の作成ロジックを分離し、コードの可読性を向上させるため。
    def create_widgets(self, master):
        
        # 処理: 余白（padding）を持たせたメインフレームを作成
        # 理由: 見た目を整え、部品同士がくっつきすぎるのを防ぐため。
        main_frame = tk.Frame(master, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 旧マスタの選択
        # 処理: ラベルを配置して入力フォームの目的を説明
        # 理由: ユーザーが何を入力すべきか一目で分かるようにするため。
        tk.Label(main_frame, text="旧マスタファイル (.csv, .xlsx)").grid(row=0, column=0, sticky='w', pady=5)
        # 処理: Entryウィジェットとself.old_path_varを紐づけ
        # 理由: 参照ボタンで選択したファイルパスが、自動でこの入力欄に表示されるようにするため。
        tk.Entry(main_frame, textvariable=self.old_path_var, width=50).grid(row=1, column=0, padx=5, sticky='ew')
        # 処理: ファイル選択ボタンを配置し、commandに関数を設定
        # 理由: このボタンがクリックされたときに、ファイル選択ダイアログを表示するメソッドを実行するため。
        tk.Button(main_frame, text="参照", command=self.select_file_old).grid(row=1, column=1, padx=5)

        # 2. 新マスタの選択
        tk.Label(main_frame, text="新マスタファイル (.csv, .xlsx)").grid(row=2, column=0, sticky='w', pady=5)
        tk.Entry(main_frame, textvariable=self.new_path_var, width=50).grid(row=3, column=0, padx=5, sticky='ew')
        tk.Button(main_frame, text="参照", command=self.select_file_new).grid(row=3, column=1, padx=5)

        # 3. 出力先フォルダの選択
        tk.Label(main_frame, text="結果出力先フォルダ").grid(row=4, column=0, sticky='w', pady=5)
        tk.Entry(main_frame, textvariable=self.output_dir_var, width=50).grid(row=5, column=0, padx=5, sticky='ew')
        tk.Button(main_frame, text="参照", command=self.select_output_dir).grid(row=5, column=1, padx=5)

        # 4. 実行ボタン
        # 処理: 処理実行ボタンを画面下部に大きく配置し、commandに実行メソッドを設定
        # 理由: ユーザーが次に何をすべきかを明確にし、クリックでデータ処理ロジックを開始させるため。
        tk.Button(main_frame, text="リニューアル品抽出を実行！", command=self.execute_analysis, 
                  bg="skyblue", fg="white", font=('Helvetica', 12, 'bold')).grid(row=6, column=0, columnspan=2, pady=20, sticky='ew')

    # 処理: ファイル選択ダイアログを表示し、選択結果を変数にセットする共通メソッドを定義
    # 理由: 旧マスタと新マスタの選択ロジックを共通化し、コードの再利用性を高めるため。
    def select_file(self, path_var: tk.StringVar):
        # 処理: 複数のファイル形式に対応したファイルタイプを設定
        # 理由: ユーザーがCSV、Excelのどのファイルを選んでもGUIが対応できるよう、実用性を高めるため。
        file_types = [
            ("マスタファイル", "*.csv *.xlsx *.xls"),
            ("CSVファイル", "*.csv"),
            ("Excelファイル", "*.xlsx *.xls")
        ]
        
        # 処理: ファイル選択ダイアログを開く
        # 理由: ユーザーがGUIから直感的にファイルを選択できるようにするため。
        file_path = filedialog.askopenfilename(
            parent=self.master, 
            title="ファイルを選択してください", 
            filetypes=file_types
        )
        # 処理: ファイルが選択された場合、変数に設定
        # 理由: 選択されたパスをEntry欄に表示し、後の処理で使用できるようにするため。
        if file_path:
            path_var.set(file_path)

    def select_file_old(self):
        # 処理: 旧マスタ用のファイル選択メソッドを呼び出し
        # 理由: 旧マスタの参照ボタンがクリックされたときに実行するため。
        self.select_file(self.old_path_var)

    def select_file_new(self):
        # 処理: 新マスタ用のファイル選択メソッドを呼び出し
        # 理由: 新マスタの参照ボタンがクリックされたときに実行するため。
        self.select_file(self.new_path_var)

    def select_output_dir(self):
        # 処理: フォルダ選択ダイアログを表示し、選択結果を変数にセット
        # 理由: 結果を保存するフォルダパスを取得するため。
        folder_path = filedialog.askdirectory(parent=self.master, title="結果を保存するフォルダを選択してください")
        if folder_path:
            self.output_dir_var.set(folder_path)

    # 処理: 実行ボタンが押されたときに、データ処理ロジックを呼び出すメソッド
    # 理由: GUIの操作をトリガーとして、裏側の複雑な分析処理を実行するため。
    def execute_analysis(self):
        # 処理: GUIの変数から、ユーザーが入力したファイルパスを取得
        # 理由: データ処理関数が、実際に処理するファイルの場所を知るため。
        old_path = self.old_path_var.get()
        new_path = self.new_path_var.get()
        output_dir = self.output_dir_var.get()
        
        # 処理: 必要なファイルパスと出力先フォルダがすべて指定されているかを確認
        # 理由: 入力不足によるエラーを防ぐための、ツールの安定性向上に必須なチェック。
        if not all([old_path, new_path, output_dir]):
            messagebox.showwarning("入力不足", "旧マスタ、新マスタ、出力先フォルダをすべて選択してください。")
            return

        try:
            # 処理: 定義したデータ処理のメイン関数を呼び出し
            # 理由: データ分析のロジックを実行させるため。
            output_file_path = process_master_data(old_path, new_path, output_dir)
            
            # 処理: 処理成功のメッセージをポップアップで表示
            # 理由: ユーザーに処理の完了と結果の保存場所を明確に伝えるため。
            messagebox.showinfo("完了", f"✅ リニューアル品抽出が完了しました。\n結果ファイル: {output_file_path}")

        except Exception as e:
            # 処理: 処理中に発生したエラーメッセージを、ユーザーに分かりやすいポップアップで表示
            # 理由: ツールがエラーでクラッシュするのを防ぎ、問題の原因をユーザーに通知し、ツールの信頼性を保つため。
            messagebox.showerror("エラー", f"❌ 処理中にエラーが発生しました。\n詳細: {e}")

# ----------------------------------------
# 4. メイン処理（アプリ起動）
# ----------------------------------------
if __name__ == "__main__":
    # 処理: Tkinterのメインウィンドウを初期化し、アプリケーションを実行
    # 理由: Pythonスクリプトとして直接実行されたときにGUIを起動するための標準的な記述方法。
    root = tk.Tk()
    app = MasterMatcherApp(root)
    root.mainloop()