import os
import streamlit as st
import pandas as pd
import pytesseract
from PIL import Image

# Tesseractのパス（Windows用）
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\h.yoshihiro\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# CSVファイルの保存先ディレクトリ
SAVE_DIR = r"C:\Users\h.yoshihiro\AppData\Local\Programs\Tesseract-OCR"
SAVE_PATH = os.path.join(SAVE_DIR, "outlook_contacts.csv")

def convert_to_outlook_format(df):
    """名刺データをOutlook用のCSVフォーマットに変換"""
    outlook_columns = ['First Name', 'Last Name', 'E-mail Address', 'Company', 'Job Title', 'Business Phone', 'Mobile Phone', 'Business Address']
    outlook_df = pd.DataFrame(columns=outlook_columns)
    
    for col in outlook_columns:
        outlook_df[col] = df.get(col, '')

    return outlook_df

def extract_text_from_image(image):
    """画像からOCRでテキストを抽出"""
    return pytesseract.image_to_string(image, lang='eng')

def parse_ocr_text(text):
    """OCRで取得したテキストをDataFrameに変換"""
    lines = text.split('\n')
    data = {'First Name': '', 'Last Name': '', 'E-mail Address': '', 'Company': '', 'Job Title': '', 'Business Phone': '', 'Mobile Phone': '', 'Business Address': ''}

    for line in lines:
        if '@' in line:
            data['E-mail Address'] = line.strip()
        elif any(c.isdigit() for c in line):
            if data['Mobile Phone'] == '':
                data['Mobile Phone'] = line.strip()
            else:
                data['Business Phone'] = line.strip()
        elif len(line.split()) > 1:
            if data['Company'] == '':
                data['Company'] = line.strip()
            else:
                data['Job Title'] = line.strip()
        else:
            if data['First Name'] == '':
                data['First Name'] = line.strip()
            else:
                data['Last Name'] = line.strip()

    return pd.DataFrame([data])

def main():
    st.title("名刺データ → Outlook インポート用変換ツール")
    
    uploaded_file = st.file_uploader("名刺の画像をアップロード (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg"])
    
    if uploaded_file is not None:
        image = Image.open(uploaded_file)
        st.image(image, caption="アップロードされた名刺", use_column_width=True)
        
        text = extract_text_from_image(image)
        st.text_area("OCR抽出テキスト", text, height=200)
        
        df = parse_ocr_text(text)
        st.write("抽出されたデータのプレビュー:")
        st.dataframe(df)
        
        outlook_df = convert_to_outlook_format(df)
        st.write("Outlook形式に変換:")
        st.dataframe(outlook_df)
        
        # 指定のフォルダに保存
        outlook_df.to_csv(SAVE_PATH, index=False)
        
        st.success(f"CSVファイルが {SAVE_PATH} に保存されました。")
        st.write(f"[ダウンロードするにはこちらをクリック](file:///{SAVE_PATH})")

if __name__ == "__main__":
    main()
