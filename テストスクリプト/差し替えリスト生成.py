# -*- coding: utf-8 -*-
"""
花王・プラネットの商品差し替えリスト作成スクリプト（構造最適化版）
Created by hibi keita
"""

from pathlib import Path
import pandas as pd

# --- 0. 設定 ---
ROOT_DIR = Path("C:/Users/337475/Box/LTS様/■アルゴリズム関連/依頼事項")
Kao_PATHS = [
    ROOT_DIR / "花王新規品廃止品リスト/2025年春/2025年春新製品廃止品対比表_バーコードなし（1225).xlsm",
    ROOT_DIR / "花王新規品廃止品リスト/2024年秋/2024年秋新製品・廃止品対比表_バーコードなし（0705）.xlsm"
]
Planet_PATHS = {
    "2024秋": {
        "new": ROOT_DIR / "プラネット新規品廃止品リスト/2024年秋/新製品リスト_20241128_085316.xlsx",
        "disc": ROOT_DIR / "プラネット新規品廃止品リスト/2024年秋/廃番品リスト_20241128_085150.xlsx",
    },
    "2025春": {
        "new": ROOT_DIR / "プラネット新規品廃止品リスト/2025年春/新製品リスト_20250128_134956.xlsx",
        "disc": ROOT_DIR / "プラネット新規品廃止品リスト/2025年春/廃番品リスト_20250128_135031.xlsx",
    },
    "2025秋": {
        "new": ROOT_DIR / "プラネット新規品廃止品リスト/2025年秋/新製品リスト_2025秋版_仮.xlsx",
        "disc": ROOT_DIR / "プラネット新規品廃止品リスト/2025年秋/廃番品リスト_2025秋版_仮.xlsx",
    },
}

# --- 1. 花王データ読み込み関数 ---
def load_kao(path):
    df = pd.read_excel(path, usecols=[6, 14, 41, 43], skiprows=5, header=None)
    df.columns = ['新商品名', '新JAN', '旧JAN', '旧商品名']
    return df.dropna(subset=['旧JAN', '新JAN'])[['旧JAN', '旧商品名', '新JAN', '新商品名']]

# --- 2. プラネットクレンジング関数 ---
def clean_planet(df, mode):
    df.columns = df.columns.str.replace('ＪＡＮ', 'JAN')
    if mode == 'discontinue':
        df = df.dropna(subset=['JANコード', '廃番予定品'])
        return df.rename(columns={'JANコード': '旧JAN', '廃番予定品': '旧商品名'})[['旧JAN', '旧商品名']]
    else:
        df = df.dropna(subset=['JANコード', '旧JANコード', '商品名全角'])
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
def process_planet_diff():
    result = []
    for season, paths in Planet_PATHS.items():
        new_df = pd.read_excel(paths['new'])
        disc_df = pd.read_excel(paths['disc'])
        new_df = exclude_kao(new_df, 'メーカーコード')
        disc_df = exclude_kao(disc_df, 'メーカー')
        new_clean = clean_planet(new_df, 'new')
        disc_clean = clean_planet(disc_df, 'discontinue')
        diff = extract_unmatched(new_clean, disc_clean)
        result.append(diff)
    return pd.concat(result, ignore_index=True)

# --- 6. クリーンアップ処理 ---
def finalize(df):
    df = df.rename(columns={'旧JAN': '旧JANコード', '新JAN': '新JANコード'})
    for col in ['旧JANコード', '新JANコード']:
        df[col] = (df[col].astype(str)
                         .str.replace(r'\D+', '', regex=True)
                         .replace('', pd.NA)
                         .astype('Int64'))
    df['旧商品名'] = df['旧商品名'].replace('', '該当文字列なし')
    df['新商品名'] = df['新商品名'].replace('', '該当文字列なし')
    return df[df['旧JANコード'] != df['新JANコード']]

# --- 7. メイン処理 ---
def main():
    kao_df = pd.concat([load_kao(p) for p in Kao_PATHS], ignore_index=True)
    planet_df = process_planet_diff()
    all_df = pd.concat([kao_df, planet_df], ignore_index=True)
    final_df = finalize(all_df)
    final_df.to_csv(ROOT_DIR / "花王差し替えリスト完成版.csv", index=False, encoding='cp932')
    print("🎉 差し替えリスト作成完了！")

if __name__ == '__main__':
    main()
