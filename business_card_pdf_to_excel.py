import os
import json
import fitz  # PyMuPDF
import google.generativeai as genai
from PIL import Image
import io
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from dotenv import load_dotenv

# ===== 設定 =====
load_dotenv()
API_KEY = os.environ.get("GEMINI_API_KEY")
genai.configure(api_key=API_KEY)
# ================

def select_pdf_files():
    """ファイル選択ダイアログを表示し、複数のPDFを選択する"""
    root = tk.Tk()
    root.withdraw() # メインウィンドウを隠す
    root.attributes("-topmost", True) # ダイアログを最前面に表示

    file_paths = filedialog.askopenfilenames(
        title="名刺のPDFファイルを選択してください（複数選択可）",
        filetypes=[("PDFファイル", "*.pdf")]
    )
    return file_paths

def pdf_to_image(pdf_path):
    """PDFの1ページ目を高解像度画像に変換する"""
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=300)
    img_data = pix.tobytes("png")
    return Image.open(io.BytesIO(img_data))

def extract_business_card_info(image):
    """Gemini APIに画像を渡し、名刺情報をJSONで抽出する"""
    model = genai.GenerativeModel('gemini-1.5-flash')

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

    response = model.generate_content(
        [prompt, image],
        generation_config=genai.GenerationConfig(
            response_mime_type="application/json",
        )
    )
    return json.loads(response.text)

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
            # 1. 画像化
            img = pdf_to_image(pdf_path)

            # 2. AIでテキスト・構造抽出
            data = extract_business_card_info(img)

            # 3. どのファイルから抽出したか分かるようにファイル名を付与
            data['source_file'] = filename

            extracted_data.append(data)
            print(f"  -> 成功: {data.get('name', '名前なし')} ({data.get('company_name', '会社名なし')})")

        except Exception as e:
            print(f"  -> エラーが発生しました: {e}")

    # 結果をExcelに保存
    if extracted_data:
        print("\nExcelファイルを作成しています...")
        df = pd.DataFrame(extracted_data)

        # 列の並び順を指定（抽出漏れでキーが存在しない場合も考慮）
        columns_order = ['company_name', 'department', 'name', 'email', 'phone', 'address', 'source_file']
        # DataFrameに存在する列だけを残して並び替え
        df = df[[col for col in columns_order if col in df.columns]]

        # 列名を日本語にリネーム
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
        # Excelとして出力 (openpyxlが必要)
        df.to_excel(output_filename, index=False)

        print(f"完了しました！ {len(extracted_data)}件のデータを '{output_filename}' に保存しました。")

        # 完了のポップアップ通知
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo("処理完了", f"{len(extracted_data)}件の名刺データの処理が完了し、\n'{output_filename}' に保存しました。")
    else:
        print("\n有効なデータが1件も抽出できませんでした。")

if __name__ == "__main__":
    main()
