# -*- coding: utf-8 -*-
"""
統合マスタ管理システム（完全版・1ファイル）
マッチング処理 + 花王・プラネット処理 + 累積管理

Author: HIBI KEITA
Version: 1.0
"""

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from fuzzywuzzy import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows

# win32com（Excel修復用）
try:
    import win32com.client as win32
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


# ========================================================================
# 共通ユーティリティ関数
# ========================================================================

def calculate_similarity(s1: str, s2: str) -> float:
    """文字列の類似度計算（0.0〜1.0）"""
    if pd.isna(s1) or pd.isna(s2):
        return 0.0
    return fuzz.ratio(str(s1), str(s2)) / 100.0


def get_weight_range(weight):
    """目付から許容範囲を計算（90%~110%）"""
    try:
        w = float(weight)
        return w * 0.9, w * 1.1
    except (ValueError, TypeError):
        return None, None


# ========================================================================
# マッチング処理モジュール
# ========================================================================

def load_data_for_matching(file_path: str, suffix: str) -> pd.DataFrame:
    """
    マッチング用データ読み込み - エンコーディング自動判別強化版
    """
    p = Path(file_path)
    ext = p.suffix.lower()
    df = pd.DataFrame()
    
    # 【Excelファイル処理】: Excelはread_csvでは読めないから、別の関数で処理を分けるよ
    if ext in ['.xlsx', '.xls', '.xlsm']:
        try:
            # openpyxlエンジンでExcelを読み込んでいるよ
            df = pd.read_excel(file_path, engine='openpyxl')
            # 読み込み成功したらすぐに後続処理へ移るよ
            return df.replace('NULL', pd.NA).add_suffix(suffix)
        except Exception as e:
            # Excel特有のエラーが出たら、すぐにユーザーに知らせるよ
            raise ValueError(f"Excelファイル読み込みエラー: {e}")

    # 【CSV/TSV/TXTファイル処理】: テキストベースのファイルを処理するよ
    read_params = {
        '.csv': {'delimiter': ','},
        '.tsv': {'delimiter': '\t'}, # ここがTSV対応のキモ！区切り文字をタブに設定してるよ
        '.txt': {'delimiter': '\t'},
    }
    
    # サポート外の拡張子ならエラーを出すよ
    if ext not in read_params:
        raise ValueError(f"サポート外のファイル形式:{ext}")

    # 日本語データでよくあるエンコーディングのリストを定義しているよ
    encodings = ['utf-8', 'shift_jis', 'cp932', 'euc_jp']
    
    common_args = {
        'delimiter': read_params[ext]['delimiter'],
        'on_bad_lines': 'skip' # 不正な行はスキップして、処理を中断させないようにしてるよ
    }

    # 定義したエンコーディングを順番に試しているよ
    for encoding in encodings:
        try:
            # 試行錯誤でファイルを読み込んでいるよ
            df = pd.read_csv(file_path, encoding=encoding, **common_args)
            # print(f"✅ {ext}を'{encoding}'エンコーディングで読み込み成功。")
            break # 成功したらループを抜けるよ
        except UnicodeDecodeError:
            continue # デコードエラーなら次のエンコーディングを試すよ
        except Exception as e:
            raise Exception(f"ファイル読み込み中に予期せぬエラーが発生しました: {e}")
    else:
        # すべて失敗したらエラーを出すよ
        raise UnicodeDecodeError(f"すべてのエンコーディング({', '.join(encodings)})で'{ext}'ファイルの読み込みに失敗しました。")
    
    # 欠損値の統一処理をしているよ
    df = df.replace('NULL', pd.NA)
    
    # 後続処理に必要な必須カラムのリストだよ
    required_cols = [
        'メーカーコード', 'ブランドコード', '標準分類コード(タイプ)',
        '目付', 'ブランド名称', '標準分類名(クラス)',
        '商品名称（カナ）', 'JANコード', 'メーカー名称',
    ]
    
    # 必須カラムがない場合は、NAカラムを追加してエラーを回避しているよ
    for col in required_cols:
        if col not in df.columns:
            df[col] = pd.NA
    
    # 新旧マスタの区別をつけるために接尾辞を追加しているよ
    df = df.add_suffix(suffix)
    return df


def clean_initial_data(df: pd.DataFrame, suffix: str) -> pd.DataFrame:
    """初期データクレンジング"""
    maker_col = f'メーカー名称{suffix}'
    brand_col = f'ブランドコード{suffix}'
    type_col = f'標準分類コード(タイプ){suffix}'
    
    before_count = len(df)
    df_cleaned = df[
        ~(df[maker_col].isna() & df[brand_col].isna() & df[type_col].isna())
    ].copy()
    
    after_count = len(df_cleaned)
    print(f"【クレンジング{suffix}】{before_count}行 → {after_count}行")
    
    return df_cleaned


def preprocess_old_data(df_old: pd.DataFrame):
    """旧マスタ事前処理（インデックス化）"""
    df_processed = df_old.copy()
    df_processed['メーカー名称_旧'] = df_processed['メーカー名称_旧'].astype(str).str.strip()
    df_processed['ブランドコード_旧'] = df_processed['ブランドコード_旧'].astype(str).str.strip()
    df_processed['標準分類コード(タイプ)_旧'] = df_processed['標準分類コード(タイプ)_旧'].astype(str).str.strip()
    df_processed['目付_旧_float'] = pd.to_numeric(df_processed['目付_旧'], errors='coerce')
    
    df_brands_indexed = df_processed.set_index('ブランドコード_旧')
    df_multi_indexed = df_processed.set_index(['メーカー名称_旧', '標準分類コード(タイプ)_旧'])
    
    return df_processed, df_brands_indexed, df_multi_indexed


def find_best_match(new_row, df_old_processed, df_old_brands_indexed, df_old_multi_indexed, df_old_original):
    """1行の新品に対して最適な旧品を検索"""
    new_maker_name = str(new_row.get('メーカー名称_新')).strip() if pd.notna(new_row.get('メーカー名称_新')) else None
    new_brand = str(new_row.get('ブランドコード_新')).strip() if pd.notna(new_row.get('ブランドコード_新')) else None
    new_type = str(new_row.get('標準分類コード(タイプ)_新')).strip() if pd.notna(new_row.get('標準分類コード(タイプ)_新')) else None
    new_weight = new_row.get('目付_新')
    new_name = new_row.get('商品名称（カナ）_新')
    
    if new_maker_name == 'nan': new_maker_name = None
    if new_brand == 'nan': new_brand = None
    if new_type == 'nan': new_type = None
    
    skip_reasons = []
    matching_old = None
    
    # パターン1: ブランドあり
    if new_brand:
        try:
            matching_old = df_old_brands_indexed.loc[new_brand].copy()
            if isinstance(matching_old, pd.Series):
                matching_old = matching_old.to_frame().T
        except KeyError:
            matching_old = df_old_processed.iloc[0:0].copy()
        
        if matching_old.empty:
            return {'照合結果': '候補なし（ブランド不一致）', '最高類似度': 0.0, '判定': '✕', 
                    '候補': '', 'スキップ理由': '', '候補あり': False}
        
        if pd.notna(new_weight):
            min_w, max_w = get_weight_range(new_weight)
            if min_w and max_w:
                weight_filtered = matching_old[
                    (matching_old['目付_旧_float'] >= min_w) & 
                    (matching_old['目付_旧_float'] <= max_w)
                ]
                if weight_filtered.empty:
                    return {'照合結果': '候補なし（目付範囲外）', '最高類似度': 0.0, '判定': '✕',
                            '候補': '', 'スキップ理由': '', '候補あり': False}
                matching_old = weight_filtered
        else:
            skip_reasons.append('目付スキップ')
    
    # パターン2: ブランドなし → メーカー+タイプ
    elif new_maker_name and new_type:
        try:
            matching_old = df_old_multi_indexed.loc[(new_maker_name, new_type)].copy()
            if isinstance(matching_old, pd.Series):
                matching_old = matching_old.to_frame().T
        except KeyError:
            matching_old = df_old_processed.iloc[0:0].copy()
        
        if matching_old.empty:
            return {'照合結果': '候補なし（メーカー名称+タイプ不一致）', '最高類似度': 0.0, '判定': '✕',
                    '候補': '', 'スキップ理由': '', '候補あり': False}
        
        if pd.notna(new_weight):
            min_w, max_w = get_weight_range(new_weight)
            if min_w and max_w:
                weight_filtered = matching_old[
                    (matching_old['目付_旧_float'] >= min_w) & 
                    (matching_old['目付_旧_float'] <= max_w)
                ]
                if weight_filtered.empty:
                    return {'照合結果': '候補なし（目付範囲外）', '最高類似度': 0.0, '判定': '✕',
                            '候補': '', 'スキップ理由': '', '候補あり': False}
                matching_old = weight_filtered
        else:
            skip_reasons.append('目付スキップ')
    else:
        return {'照合結果': '候補なし（キーコード不足）', '最高類似度': 0.0, '判定': '✕',
                '候補': '', 'スキップ理由': '', '候補あり': False}
    
    # 名称一致で最高類似度を選択
    candidates = matching_old[['商品名称（カナ）_旧', 'JANコード_旧']].drop_duplicates()
    
    if candidates.empty:
        return {'照合結果': '候補なし（名称一致なし）', '最高類似度': 0.0, '判定': '✕',
                '候補': '', 'スキップ理由': '、'.join(skip_reasons) if skip_reasons else '', '候補あり': False}
    
    similarities = [
        (calculate_similarity(new_name, row['商品名称（カナ）_旧']), 
         row['商品名称（カナ）_旧'], 
         row['JANコード_旧'])
        for _, row in candidates.iterrows()
    ]
    similarities.sort(key=lambda x: x[0], reverse=True)
    
    best_score, best_name, best_jan = similarities[0]
    best_old_row = df_old_original[df_old_original['JANコード_旧'] == best_jan].iloc[0]
    
    if best_score >= 0.8:
        result = '高類似度候補あり (80%以上)'
        judgment = '○'
    else:
        result = '低類似度 (80%未満・要手動確認)'
        judgment = '✕'
    
    return {
        '照合結果': result,
        '最高類似度': best_score,
        '判定': judgment,
        '候補': f"{best_name}({best_score:.1%})",
        'スキップ理由': '、'.join(skip_reasons) if skip_reasons else '',
        '候補あり': True,
        'JANコード_旧': best_old_row.get('JANコード_旧', ''),
        '商品名称（カナ）_旧': best_old_row.get('商品名称（カナ）_旧', ''),
        'メーカー名称_旧': best_old_row.get('メーカー名称_旧', ''),
        '標準分類(クラス)_旧': best_old_row.get('標準分類名(クラス)_旧', ''),
        'ブランド名称_旧': best_old_row.get('ブランド名称_旧', ''),
        '目付_旧': best_old_row.get('目付_旧', ''),
        '発売日_旧': best_old_row.get('発売日_旧', ''),
    }


def run_matching_process(old_path: str, new_path: str) -> pd.DataFrame:
    """マッチング処理実行（統合用）"""
    print("\n📊 マッチング処理開始...")
    
    df_new = load_data_for_matching(new_path, '_新')
    df_old = load_data_for_matching(old_path, '_旧')
    
    df_new = clean_initial_data(df_new, '_新')
    df_old = clean_initial_data(df_old, '_旧')
    
    if '発売日_旧' not in df_old.columns:
        df_old['発売日_旧'] = pd.NA
    
    print("旧マスタ前処理中...")
    df_old_processed, df_old_brands_indexed, df_old_multi_indexed = preprocess_old_data(df_old)
    
    print(f"突合処理開始...（{len(df_new)}件）")
    
    results = []
    for idx, new_row in df_new.iterrows():
        if idx % 100 == 0:
            print(f"処理中... {idx}/{len(df_new)}")
        
        result = find_best_match(new_row, df_old_processed, df_old_brands_indexed, 
                                df_old_multi_indexed, df_old)
        results.append(result)
    
    analysis_result = pd.DataFrame(results)
    final_df = pd.concat([df_new.reset_index(drop=True), analysis_result], axis=1)
    
    final_df = final_df.rename(columns={
        '標準分類名(クラス)_新': '標準分類(クラス)_新',
    }, errors='ignore')
    
    # 候補ありのみフィルタ
    if '候補あり' in final_df.columns:
        final_df = final_df[final_df['候補あり'] == True].copy()
        final_df = final_df.drop('候補あり', axis=1)
    
    print(f"✅ マッチング処理完了: {len(final_df)}件")
    return final_df


# ========================================================================
# 花王・プラネット処理モジュール
# ========================================================================

def repair_and_resave_excel(file_path):
    """Excel自動修復"""
    if not WIN32COM_AVAILABLE:
        return False
    
    p_file = Path(file_path)
    print(f"🛠️ Excelファイル修復中: {p_file.name}")
    
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
        
        print(f"✅ 修復完了: {p_file.name}")
        return True
    except Exception as e:
        print(f"❌ 修復失敗: {e}")
        return False
    finally:
        if excel:
            excel.Quit()
        try:
            pythoncom.CoUninitialize()
        except:
            pass


def load_with_repair(path, **read_excel_kwargs):
    """修復機能付きExcel読み込み"""
    p_file = Path(path)
    
    try:
        df = pd.read_excel(p_file, engine='openpyxl', **read_excel_kwargs)
        return df
    except Exception as e:
        print(f"⚠️ 通常読み込み失敗: {e}")
        
        if WIN32COM_AVAILABLE:
            print("🔧 修復を試みています...")
            if repair_and_resave_excel(p_file):
                try:
                    df = pd.read_excel(p_file, engine='openpyxl', **read_excel_kwargs)
                    return df
                except Exception as retry_e:
                    print(f"❌ 修復後も読み込み失敗: {retry_e}")
                    raise
        raise


def load_kao(path):
    """花王データ読み込み"""
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


def clean_planet(df, mode):
    """プラネットデータクレンジング"""
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
    else:
        df = df.dropna(subset=['JANコード', '旧JANコード'])
        return df.rename(columns={
            '旧JANコード': '旧JAN',
            'JANコード': '新JAN',
            '商品名全角': '新商品名'
        })[['旧JAN', '新JAN', '新商品名']]


def extract_unmatched(new_df, old_df):
    """純粋新規品抽出"""
    add = new_df[~new_df['新JAN'].isin(old_df['旧JAN'])].copy()
    add['旧商品名'] = ''
    return add[['旧JAN', '旧商品名', '新JAN', '新商品名']]


def exclude_kao(df, is_kao_col):
    """花王データ除外"""
    return df[~df[is_kao_col].astype(str).str.startswith('4901301') & 
              ~df[is_kao_col].astype(str).str.contains('花王株式会社')]


def finalize_kao_planet(df):
    """花王・プラネットデータ最終クリーンアップ"""
    df = df.rename(columns={'旧JAN': '旧JANコード', '新JAN': '新JANコード'})
    
    for col in ['旧JANコード', '新JANコード']:
        df[col] = (df[col].astype(str)
                   .str.replace(r'\D+', '', regex=True)
                   .replace('', pd.NA)
                   .apply(lambda x: str(x).zfill(13)[:13] if pd.notna(x) else pd.NA))
    
    df['旧商品名'] = df['旧商品名'].replace('', '該当文字列なし')
    df['新商品名'] = df['新商品名'].replace('', '該当文字列なし')
    
    return df[df['旧JANコード'] != df['新JANコード']].drop_duplicates()


def run_kao_planet_process(kao_files, planet_paths_dict) -> pd.DataFrame:
    """花王・プラネット処理実行（統合用）"""
    print("\n🏭 花王・プラネット処理開始...")
    
    combined_df = pd.DataFrame()
    
    # 花王処理
    if kao_files:
        kao_df = pd.concat([load_kao(p) for p in kao_files], ignore_index=True)
        kao_df = kao_df.rename(columns={'備考': '新JAN備考'})
        combined_df = pd.concat([combined_df, kao_df], ignore_index=True)
    
    # プラネット処理
    if planet_paths_dict:
        planet_result = []
        for season, paths in planet_paths_dict.items():
            try:
                new_df = load_with_repair(paths['new'], dtype={'ＪＡＮコード': str, '旧ＪＡＮコード': str})
                new_df['備考'] = paths['new'].name
                
                disc_df = load_with_repair(paths['disc'], dtype={'JANコード': str, '新JANコード': str})
                disc_df['備考'] = paths['disc'].name
                
                new_df = exclude_kao(new_df, 'メーカーコード')
                disc_df = exclude_kao(disc_df, 'メーカー')
                
                new_clean = clean_planet(new_df, 'new')
                disc_clean = clean_planet(disc_df, 'discontinue')
                
                disc_not_in_new = disc_clean[~disc_clean['新JAN'].isin(new_clean['新JAN'])].copy()
                final_disc_additions = disc_not_in_new[~disc_not_in_new['旧JAN'].isin(new_clean['旧JAN'])].copy()
                
                pure_new_items = extract_unmatched(new_clean, disc_clean)
                
                combined_planet = pd.concat([pure_new_items, final_disc_additions], ignore_index=True)
                combined_planet_with_notes = pd.merge(
                    combined_planet,
                    new_df[['JANコード', '備考']],
                    left_on='新JAN',
                    right_on='JANコード',
                    how='left'
                ).drop(columns='JANコード').rename(columns={'備考': '新JAN備考'})
                
                planet_result.append(combined_planet_with_notes)
            except Exception as e:
                print(f"❌ {season}処理失敗: {e}")
                continue
        
        if planet_result:
            planet_df = pd.concat(planet_result, ignore_index=True)
            combined_df = pd.concat([combined_df, planet_df], ignore_index=True)
    
    if combined_df.empty:
        return pd.DataFrame()
    
    final_df = finalize_kao_planet(combined_df)
    print(f"✅ 花王・プラネット処理完了: {len(final_df)}件")
    
    return final_df


# ========================================================================
# 統合処理モジュール
# ========================================================================

def normalize_matching_columns(df: pd.DataFrame) -> pd.DataFrame:
    """マッチング結果を統一フォーマットに変換"""
    df_normalized = df.rename(columns={
        'JANコード_旧': '旧JANコード',
        'JANコード_新': '新JANコード',
        '商品名称（カナ）_旧': '旧商品名',
        '商品名称（カナ）_新': '新商品名',
        'メーカー名称_新': 'メーカー名称',
    }).copy()
    
    df_normalized['データソース'] = 'マッチング'
    df_normalized['処理日'] = datetime.now().strftime('%Y-%m-%d')
    df_normalized['備考'] = df_normalized.get('照合結果', '')
    
    return df_normalized[['旧JANコード', '旧商品名', '新JANコード', '新商品名',
                          'メーカー名称', '備考', '処理日', 'データソース']]


def normalize_kao_planet_columns(df: pd.DataFrame) -> pd.DataFrame:
    """花王・プラネット結果を統一フォーマットに変換"""
    df_normalized = df.copy()
    
    if 'メーカー名称' not in df_normalized.columns:
        df_normalized['メーカー名称'] = ''
    
    if '備考' not in df_normalized.columns:
        df_normalized['備考'] = df_normalized.get('新JAN備考', '')
    
    df_normalized['データソース'] = '花王・プラネット'
    df_normalized['処理日'] = datetime.now().strftime('%Y-%m-%d')
    
    return df_normalized[['旧JANコード', '旧商品名', '新JANコード', '新商品名',
                          'メーカー名称', '備考', '処理日', 'データソース']]


def merge_and_deduplicate(existing_df, df_kao_planet, df_matching) -> pd.DataFrame:
    """データ統合と重複削除"""
    print("\n📦 データ統合開始...")
    
    print(f"  既存: {len(existing_df)}件")
    print(f"  花王・プラネット: {len(df_kao_planet)}件")
    print(f"  マッチング: {len(df_matching)}件")
    
    all_data = pd.concat([existing_df, df_kao_planet, df_matching], ignore_index=True)
    print(f"  統合後: {len(all_data)}件")
    
    for col in ['旧JANコード', '新JANコード']:
        all_data[col] = (all_data[col].astype(str)
                         .str.replace(r'\D+', '', regex=True)
                         .str.zfill(13).str[:13])
    
    before_dedup = len(all_data)
    all_data = all_data.drop_duplicates(subset=['新JANコード'], keep='first')
    after_dedup = len(all_data)
    
    print(f"  重複削除: {before_dedup - after_dedup}件")
    
    all_data = all_data[all_data['旧JANコード'] != all_data['新JANコード']]
    print(f"  最終件数: {len(all_data)}件")
    
    return all_data


def load_existing_data(累積ファイル: Path) -> pd.DataFrame:
    """既存累積データ読み込み"""
    if 累積ファイル.exists():
        try:
            df = pd.read_csv(累積ファイル, dtype={'旧JANコード': str, '新JANコード': str}, encoding='utf-8')
            print(f"📂 既存データ: {len(df)}件")
            return df
        except Exception as e:
            print(f"⚠️ 既存データ読み込みエラー: {e}")
            return pd.DataFrame()
    else:
        print("📂 新規作成モード")
        return pd.DataFrame()


# ========================================================================
# メイン処理
# ========================================================================

def main():
    """統合システムメイン処理"""
    print("=" * 60)
    print("統合マスタ管理システム - 週次実行版")
    print("=" * 60)
    
    # 出力先選択
    root = tk.Tk()
    root.withdraw()
    output_dir = filedialog.askdirectory(title="累積データ保存先を選択")
    root.destroy()
    
    if not output_dir:
        messagebox.showwarning("キャンセル", "フォルダ未選択")
        return
    
    output_dir = Path(output_dir)
    累積ファイル_csv = output_dir / "累積_差し替えリスト.csv"
    累積ファイル_excel = output_dir / "累積_差し替えリスト.xlsx"
    
    # 既存データ読み込み
    existing_df = load_existing_data(累積ファイル_csv)
    
    # 花王・プラネット処理
    kao_files = []
    planet_paths = {}
    
    if messagebox.askyesno("花王処理", "花王ファイルを処理しますか？"):
        root = tk.Tk()
        root.withdraw()
        kao_paths = filedialog.askopenfilenames(title="花王ファイル選択（複数可）",
                                                filetypes=[("Excel", "*.xlsm *.xlsx")])
        root.destroy()
        kao_files = [Path(p) for p in kao_paths]
    
    if messagebox.askyesno("プラネット処理", "プラネットファイルを処理しますか？"):
        root = tk.Tk()
        root.withdraw()
        
        messagebox.showinfo("選択", "上期 新規品リストを選択")
        new_upper = filedialog.askopenfilename(title="上期 新規品", filetypes=[("Excel", "*.xlsx")])
        
        messagebox.showinfo("選択", "上期 廃番品リストを選択")
        disc_upper = filedialog.askopenfilename(title="上期 廃番品", filetypes=[("Excel", "*.xlsx")])
        
        messagebox.showinfo("選択", "下期 新規品リストを選択")
        new_lower = filedialog.askopenfilename(title="下期 新規品", filetypes=[("Excel", "*.xlsx")])
        
        messagebox.showinfo("選択", "下期 廃番品リストを選択")
        disc_lower = filedialog.askopenfilename(title="下期 廃番品", filetypes=[("Excel", "*.xlsx")])
        
        root.destroy()
        
        if new_upper and disc_upper:
            planet_paths["上期"] = {"new": Path(new_upper), "disc": Path(disc_upper)}
        if new_lower and disc_lower:
            planet_paths["下期"] = {"new": Path(new_lower), "disc": Path(disc_lower)}
    
    # 花王・プラネット処理実行
    df_kao_planet = pd.DataFrame()
    if kao_files or planet_paths:
        try:
            df_kao_planet_raw = run_kao_planet_process(kao_files, planet_paths)
            if not df_kao_planet_raw.empty:
                df_kao_planet = normalize_kao_planet_columns(df_kao_planet_raw)
        except Exception as e:
            print(f"❌ 花王・プラネット処理エラー: {e}")
    
    # マッチング処理
    df_matching = pd.DataFrame()
    if messagebox.askyesno("マッチング処理", "マッチング処理を実行しますか？"):
        root = tk.Tk()
        root.withdraw()
        
        # --- 💥 ここから修正箇所 💥 ---
        # 複数のファイル形式を確実に選択できるように filetypes のリストを修正しているよ
        # なんでそうしてるか: "*.xlsx *.csv*.tsv" のようなスペース区切りだとOSやPythonのバージョンによって
        # 認識されないことがあるから、拡張子ごとに個別のパターンとして指定するのが確実なんや
        all_filetypes = [
            ("Excelファイル", "*.xlsx *.xls *.xlsm"), # Excel系の拡張子をまとめてるよ
            ("CSVファイル", "*.csv"), # CSVファイルだよ
            ("TSVファイル", "*.tsv"), # TSVファイル（タブ区切り）だよ。これで選択リストに出てくるようになるよ
            ("全ファイル", "*.*"), # 念のためすべてのファイルを表示できるようにしてるよ
        ]
        
        messagebox.showinfo("選択", "旧マスタファイルを選択")
        old_path = filedialog.askopenfilename(title="旧マスタ", filetypes=all_filetypes) # 修正した filetypes を渡してるよ
        
        if not old_path:
            messagebox.showwarning("キャンセル", "旧マスタ未選択")
            root.destroy()
        else:
            messagebox.showinfo("選択", "新マスタファイルを選択")
            new_path = filedialog.askopenfilename(title="新マスタ", filetypes=all_filetypes) # 修正した filetypes を渡してるよ
            
            root.destroy()
            
            if new_path:
                try:
                    df_matching_raw = run_matching_process(old_path, new_path)
                    if not df_matching_raw.empty:
                        df_matching = normalize_matching_columns(df_matching_raw)
                except Exception as e:
                    print(f"❌ マッチング処理エラー: {e}")
    
    # データ統合
    if df_kao_planet.empty and df_matching.empty:
        messagebox.showwarning("データなし", "処理データなし")
        return
    
    final_df = merge_and_deduplicate(existing_df, df_kao_planet, df_matching)
    
    # 保存
    print("\n💾 保存中...")
    final_df.to_csv(累積ファイル_csv, index=False, encoding='utf-8')
    final_df.to_excel(累積ファイル_excel, index=False, engine='openpyxl')
    
    # 完了メッセージ
    summary = f"""
🎉 処理完了！

【累積データ】
総件数: {len(final_df)}件

【今回追加】
花王・プラネット: {len(df_kao_planet)}件
マッチング: {len(df_matching)}件

【保存先】
{累積ファイル_csv}
"""
    
    print(summary)
    messagebox.showinfo("完了", summary)


# ========================================================================
# 実行
# ========================================================================

if __name__ == "__main__":
    main()