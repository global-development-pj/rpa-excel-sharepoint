# -*- coding: utf-8 -*-
import os
import sys
import win32com.client

def convert_excel_to_pdf(excel_path, pdf_path):
    """
    Excel(.xlsx)ファイルを開き、全シートを1つのPDFにまとめて出力する (Windows専用: pywin32使用)
    - Excelで設定されている印刷範囲やレイアウトを反映
    """
    excel_path = os.path.abspath(excel_path)
    pdf_path   = os.path.abspath(pdf_path)

    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False  # 警告ダイアログを表示しない

    try:
        # ブックを開く (更新リンクなし, 読み取り専用)
        wb = excel_app.Workbooks.Open(excel_path, UpdateLinks=0, ReadOnly=1)

        # 定数 xlTypePDF (Excel のバージョンによっては 0 や 57 など)
        xlTypePDF = 0

        # PDF 出力 (全シート対象)
        wb.ExportAsFixedFormat(
            Type=xlTypePDF,
            Filename=pdf_path,
            Quality=0,  # 標準
            IncludeDocProperties=True,
            IgnorePrintAreas=False,  # シートの印刷範囲設定を尊重
            From=None, To=None,      # 全ページ
            OpenAfterPublish=False
        )
        print(f"[成功] {excel_path} → {pdf_path}")

    except Exception as e:
        print(f"[エラー] {excel_path}: {e}")

    finally:
        wb.Close(SaveChanges=False)
        excel_app.Quit()

def convert_all_excel_in_subfolders(folder_path):
    """
    指定フォルダ以下のサブフォルダも含めて .xlsx を再帰的に検索し、
    同名のPDFファイルを生成する
    """
    folder_path = os.path.abspath(folder_path)
    count = 0

    # os.walk で再帰的に探索
    for root, dirs, files in os.walk(folder_path):
        for fname in files:
            # 拡張子が .xlsx のファイルのみ対象 (必要に応じて .xlsm など追加)
            if fname.lower().endswith(".xlsx"):
                excel_file = os.path.join(root, fname)
                base_name  = os.path.splitext(fname)[0]
                pdf_name   = base_name + ".pdf"
                pdf_file   = os.path.join(root, pdf_name)

                convert_excel_to_pdf(excel_file, pdf_file)
                count += 1

    if count == 0:
        print("指定フォルダ以下に .xlsx ファイルが見つかりませんでした。")
    else:
        print(f"合計 {count} 件のファイルをPDF化しました。")

if __name__ == "__main__":
    # コンソール文字化け対策（オプション）
    # sys.stdout.reconfigure(encoding='utf-8')

    # 変換対象のフォルダを指定 (サブフォルダも含め再帰的に探す)
    target_folder = r"C:\path\to\excel_folder"
    convert_all_excel_in_subfolders(target_folder)

