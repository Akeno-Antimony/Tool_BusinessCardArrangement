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

    max_retries = 3  # 最大再試行回数
    retry_delay = 5  # 待機する秒数

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

    print(f"{len(pdf_paths)}件のファイルが選択されました。AIによる解析を開始します...")

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
        print("\nExcelファイルを作成しています...")
        df = pd.DataFrame(extracted_data)

        columns_order = ['company_name', 'department', 'name', 'email', 'phone', 'address', 'source_file']
        df = df[[col for col in columns_order if col in df.columns]]

        df.rename(columns={
            'company_name': '会社名',
            'department': '部署・役職',
            'name': '氏名',
            'email': 'メールアドレス',
            'phone': '電話番号',
            'address': '住所',
            'source_file': '元PDFファイル名'
        }, inplace=True)

        output_filename = "名刺データ一覧.xlsx"
        df.to_excel(output_filename, index=False)

        print(f"完了しました！ {len(extracted_data)}件のデータを '{output_filename}' に保存しました。")

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo("処理完了", f"{len(extracted_data)}件の名刺データの処理が完了し、\n'{output_filename}' に保存しました。")
    else:
        print("\n有効なデータが1件も抽出できませんでした。")

if __name__ == "__main__":
    main()
