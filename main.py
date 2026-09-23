import sys
import time
from cleaner import load_and_clean_all
from aggregator import create_summary_report

def run_pipeline():
    print("==================================================")
    print("📊 Excel Data Cleaner & Aggregator Pipeline")
    print("==================================================")
    start_time = time.time()

    try:
        # 1. 読み込み & クレンジング
        print("\n[Step 1/2] 複数拠点Excelの検知・正規化クレンジング中...")
        cleaned_df = load_and_clean_all("data/input")
        print(f"-> 正常処理レコード数: {len(cleaned_df)} 件")

        # 2. 集計 & Excel出力
        print("\n[Step 2/2] 集計サマリーおよび装飾レポート生成中...")
        output_file = "data/output/monthly_sales_summary.xlsx"
        create_summary_report(cleaned_df, output_file)

        elapsed = time.time() - start_time
        print(f"\n✨ パイプライン処理完了! (所要時間: {elapsed:.2f}秒)")
        print(f"📁 出力先: {output_file}")

    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    run_pipeline()