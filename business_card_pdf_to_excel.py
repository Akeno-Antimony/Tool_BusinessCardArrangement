import os
import json
import time
import pymupdf  # 修正: fitz から pymupdf に変更
from google import genai  # 修正: google.generativeai から変更
from google.genai import types  # 型定義のインポート
from PIL import Image
import io
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from dotenv import load_dotenv

# ===== 設定 =====
load_dotenv()
API_KEY = os.environ.get("GEMINI_API_KEY")
client = genai.Client(api_key=API_KEY)
# ================

def select_pdf_files():
    """ファイル選択ダイアログを表示し、複数のPDFを選択する"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    file_paths = filedialog.askopenfilenames(
        title="名刺のPDFファイルを選択してください（複数選択可）",
        filetypes=[("PDFファイル", "*.pdf")]
    )
    return file_paths

def select_output_excel():
    """追記先のExcelファイルを選択する"""
    file_path = filedialog.askopenfilename(
        title="【2/2】追記先のExcelファイルを選択してください（キャンセルで新規作成）",
        filetypes=[("Excelファイル", "*.xlsx")]
    )
    return file_path

def pdf_to_image(pdf_path):
    """PDFの1ページ目を高解像度画像に変換する"""
    # 修正: fitz.open() から pymupdf.open() に変更
    doc = pymupdf.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=300)
    img_data = pix.tobytes("png")
    return Image.open(io.BytesIO(img_data))

def extract_business_card_info(image):
    """Gemini APIに画像を渡し、名刺情報をJSONで抽出する（エラー時のリトライ付き）"""
    prompt = """
    この画像はスキャンした名刺です。画像から情報を抽出し、以下のキーを持つJSON形式で出力してください。
    読み取れない項目は null にしてください。
    - company_name: 会社名
    - department: 部署名や役職
    - name: 氏名
    - email: メールアドレス
    - phone: 電話番号
    - address: 会社の住所
    """

    max_retries = 5  # 最大再試行回数
    retry_delay = 10  # 待機する秒数

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model='gemini-3.8-flash',  # 環境に合わせて変更してください
                contents=[prompt, image],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                )
            )
            return json.loads(response.text)

        except Exception as e:
            error_msg = str(e)
            # 503エラー（UNAVAILABLE）が含まれているか確認
            if "503" in error_msg or "UNAVAILABLE" in error_msg:
                if attempt < max_retries - 1:
                    print(f"      ※API混雑を検知。{retry_delay}秒待機して再試行します... (試行 {attempt + 1}/{max_retries})")
                    time.sleep(retry_delay)
                    continue

            # 503以外のエラー、または最大回数を超えた場合はエラーとして処理を中断
            raise e

def main():
    print("ファイル選択ダイアログを開いています...")
    pdf_paths = select_pdf_files()

    if not pdf_paths:
        print("ファイルが選択されませんでした。処理を終了します。")
        return

    print("出力先Excelの選択ダイアログを開いています...")
    existing_excel_path = select_output_excel()

    processed_filenames = set()
    if existing_excel_path:
        print(f"追記モード: '{os.path.basename(existing_excel_path)}' にデータを追加します。")
        try:
            temp_df = pd.read_excel(existing_excel_path)
            if '元PDFファイル名' in temp_df.columns:
                processed_filenames = set(temp_df['元PDFファイル名'].dropna().tolist())
        except Exception as e:
            print(f"※既存ファイルの履歴読み込みに失敗しました: {e}")
    else:
        print("新規作成モード: 新しいExcelファイルを作成します。")

    target_pdf_paths = []
    for path in pdf_paths:
        filename = os.path.basename(path)
        if filename in processed_filenames:
            print(f"スキップ: {filename} (処理済み)")
        else:
            target_pdf_paths.append(path)
            time.sleep(10) # sleep for 10 seconds between AI process

    if not target_pdf_paths:
        print("\n選択されたファイルはすべて処理済みでした。処理を終了します。")
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo("処理完了", "選択されたファイルはすべて処理済みでした。\n追加するデータはありません。")
        return

    print(f"\n{len(target_pdf_paths)}件のファイルが解析対象です。AIによる解析を開始します...")

    extracted_data = []

    for i, pdf_path in enumerate(pdf_paths):
        filename = os.path.basename(pdf_path)
        print(f"[{i+1}/{len(pdf_paths)}] 処理中: {filename}")

        try:
            img = pdf_to_image(pdf_path)
            data = extract_business_card_info(img)
            data['source_file'] = filename

            extracted_data.append(data)
            print(f"  -> 成功: {data.get('name', '名前なし')} ({data.get('company_name', '会社名なし')})")

        except Exception as e:
            print(f"  -> エラーが発生しました: {e}")

    if extracted_data:
        print("\nExcelファイルに書き込んでいます...")
        new_df = pd.DataFrame(extracted_data)

        columns_order = ['company_name', 'department', 'name', 'email', 'phone', 'address', 'source_file']
        new_df = new_df[[col for col in columns_order if col in new_df.columns]]

        new_df.rename(columns={
            'company_name': '会社名',
            'department': '部署・役職',
            'name': '氏名',
            'email': 'メールアドレス',
            'phone': '電話番号',
            'address': '住所',
            'source_file': '元PDFファイル名'
        }, inplace=True)

        # 保存処理（追記 or 新規作成）
        if existing_excel_path:
            try:
                # 既存のExcelを読み込んで結合
                existing_df = pd.read_excel(existing_excel_path)
                final_df = pd.concat([existing_df, new_df], ignore_index=True)
                output_filename = existing_excel_path
                final_df.to_excel(output_filename, index=False)
            except Exception as e:
                print(f"既存Excelの読み込み・保存に失敗しました。新規ファイルとして保存します。エラー: {e}")
                output_filename = "名刺データ一覧_新規.xlsx"
                new_df.to_excel(output_filename, index=False)
        else:
            output_filename = "名刺データ一覧.xlsx"
            new_df.to_excel(output_filename, index=False)

        print(f"完了しました！ {len(extracted_data)}件のデータを '{output_filename}' に保存しました。")

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo("処理完了", f"{len(extracted_data)}件の名刺データの処理が完了し、\n'{output_filename}' に保存しました。")
    else:
        print("\n有効なデータが1件も抽出できませんでした。")

if __name__ == "__main__":
    main()
