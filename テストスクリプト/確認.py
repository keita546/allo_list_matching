# -*- coding: utf-8 -*-
"""
Created by HIBI KEITA
最適化版：新品→旧品マッチング
"""

import pandas as pd

df1 = pd.read_csv("C:/Users/337475/pythonプロジェクト/Allo差し替えリストプログラム/使用データ/旧品_花王18カテ 1.tsv", sep='\t')
df2 = pd.read_csv("C:/Users/337475/pythonプロジェクト/Allo差し替えリストプログラム/使用データ/新品_花王18カテ 1.tsv",sep='\t')

df1_nan = df1[df1['メーカーコード'].isna()]

print(df1)
print(df2)
print(df1_nan)