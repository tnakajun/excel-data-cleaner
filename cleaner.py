import os
import glob
import re
import unicodedata
import pandas as pd

# 列名のマッピング定義（左：統一後の標準名、右：各拠点で使われがちな表記候補）
COLUMN_MAPPING = {
    "日付": ["日付", "販売日", "取引日", "売上日"],
    "商品名": ["商品名", "品名", "商品"],
    "カテゴリー": ["カテゴリー", "カテゴリ", "分類", "種別"],
    "単価": ["単価", "販売価格", "価格", "定価"],
    "数量": ["数量", "販売数", "個数"],
    "担当者": ["担当者", "営業担当", "担当", "スタッフ"]
}

def clean_currency(value):
    """金額や数値の文字列から記号・単位・全角を除去して整数に変換"""
    if pd.isna(value):
        return 0
    # 全角英数を半角に変換（２ -> 2 など）
    val_str = unicodedata.normalize("NFKC", str(value))
    # 数字以外（¥, 円, カンマ, 空白等）をすべて除去
    cleaned = re.sub(r"[^\d]", "", val_str)
    return int(cleaned) if cleaned else 0

def clean_quantity(value):
    """数量の文字列から全角や空白を除去して整数に変換"""
    if pd.isna(value):
        return 0
    val_str = unicodedata.normalize("NFKC", str(value)).strip()
    cleaned = re.sub(r"[^\d]", "", val_str)
    return int(cleaned) if cleaned else 0

def normalize_date(value):
    """様々な形式の日付（2026年9月2日, 2026/09/01, 2026-09-01）を統一形式に変換"""
    if pd.isna(value):
        return ""
    val_str = str(value).strip()
    # 「年月日」表記をスラッシュに変換
    val_str = re.sub(r"(\d{4})年(\d{1,2})月(\d{1,2})日", r"\1/\2/\3", val_str)
    try:
        # Pandasのパーサーで日付型に変換して YYYY-MM-DD 形式の文字列にする
        return pd.to_datetime(val_str).strftime("%Y-%m-%d")
    except Exception:
        return val_str

def process_single_file(file_path):
    """1つのExcelファイルを読み込み、列名を標準化してクレンジングを行う"""
    df = pd.read_excel(file_path)
    
    # 拠点名をファイル名から取得（例: branch_tokyo.xlsx -> 東京）
    file_name = os.path.basename(file_path)
    branch_code = file_name.replace("branch_", "").replace(".xlsx", "")
    branch_map = {"tokyo": "東京支社", "osaka": "大阪支社", "fukuoka": "福岡支社"}
    branch_name = branch_map.get(branch_code, branch_code)

    # 1. 列名の標準化
    rename_dict = {}
    for col in df.columns:
        col_clean = str(col).strip()
        for standard_name, aliases in COLUMN_MAPPING.items():
            if col_clean in aliases:
                rename_dict[col] = standard_name
                break
    df = df.rename(columns=rename_dict)

    # 2. 必要な列の存在確認と型のクレンジング
    if "単価" in df.columns:
        df["単価"] = df["単価"].apply(clean_currency)
    if "数量" in df.columns:
        df["数量"] = df["数量"].apply(clean_quantity)
    if "日付" in df.columns:
        df["日付"] = df["日付"].apply(normalize_date)
    
    # 前後の空白を除去
    for text_col in ["商品名", "カテゴリー", "担当者"]:
        if text_col in df.columns:
            df[text_col] = df[text_col].astype(str).str.strip()

    # 3. 計算列と拠点名の追加
    df["売上金額"] = df["単価"] * df["数量"]
    df["拠点"] = branch_name

    # 標準の列順に並び替え
    target_columns = ["拠点", "日付", "担当者", "カテゴリー", "商品名", "単価", "数量", "売上金額"]
    available_cols = [c for c in target_columns if c in df.columns]
    return df[available_cols]

def load_and_clean_all(input_dir="data/input"):
    """指定フォルダ内の全Excelを走査し、統合された1つのDataFrameを返す"""
    files = glob.glob(os.path.join(input_dir, "*.xlsx"))
    if not files:
        raise FileNotFoundError(f"{input_dir} にExcelファイルが見つかりません。")

    cleaned_dfs = []
    for f in sorted(files):
        print(f"🔄 読み込み・クレンジング中: {os.path.basename(f)}")
        cleaned_df = process_single_file(f)
        cleaned_dfs.append(cleaned_df)

    # 全拠点のデータを縦方向に合体
    merged_df = pd.concat(cleaned_dfs, ignore_index=True)
    
    # 日付昇順でソート
    merged_df = merged_df.sort_values(by=["日付", "拠点"]).reset_index(drop=True)
    return merged_df

if __name__ == "__main__":
    # 単体テスト用実行ブロック
    result_df = load_and_clean_all()
    print("\n✅ 統合クレンジング完了！プレビュー:\n")
    print(result_df.to_string(index=False))