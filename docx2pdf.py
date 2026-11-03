# -*- coding: utf-8 -*-
"""用 Word COM 把 docx 转为 PDF（无 Word 界面时用头less COM）"""
import sys
import os

def convert(docx_path):
    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        pdf_path = os.path.splitext(docx_path)[0] + "_wpdf.pdf"
        # 17 = wdFormatPDF
        doc.ExportAsFixedFormat(
            os.path.abspath(pdf_path), 17,  # wdExportFormatPDF
            Item=0,  # 导出全部页
        )
        doc.Close(False)
        print(f"已转换: {pdf_path}")
        return pdf_path
    finally:
        word.Quit()

if __name__ == "__main__":
    for p in sys.argv[1:]:
        try:
            convert(p)
        except Exception as e:
            print(f"转换失败 {p}: {e}")
