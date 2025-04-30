import subprocess
import os


def convert_docx_to_pdf(docx_path):
    """
    將 DOCX 檔案轉換為 PDF。

    Args:
        docx_path (str): DOCX 檔案的路徑。

    Returns:
        str: 轉換後的 PDF 檔案路徑，如果轉換失敗則為 None。
    """
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    try:
        subprocess.run([
            "soffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(docx_path), docx_path
        ], check=True, capture_output=True, text=True)  # 捕捉標準輸出和標準錯誤
        return pdf_path
    except subprocess.CalledProcessError as e:
        print(f"DOCX 轉 PDF 失敗: {e.stderr}")  # 輸出錯誤訊息到控制台
        return None