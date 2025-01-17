# -*- coding: utf-8 -*-
import os
import glob
from openpyxl import load_workbook
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.units import cm
from io import BytesIO
from PIL import Image as PILImage

# ▼ 日本語フォントを使う場合（Windowsの例）
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Windows標準フォント「MSゴシック」を登録 (フォントファイルは環境に合わせて変更)
pdfmetrics.registerFont(TTFont('MS-Gothic', 'msgothic.ttc'))

def excel_to_pdf(excel_path, pdf_path):
    """
    指定したExcel(.xlsx)ファイル内の全シートを読み取り、
    A3横向きのPDFに出力するサンプル関数
    - セルの値をTableで出力
    - シート内の画像も取得して順番にPDFへ配置
    """
    # 1) Excelファイルを読み込む
    workbook = load_workbook(excel_path, data_only=True)
    
    # 2) ReportLabでPDFファイル準備 (A3横向き)
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A3)
    )
    story = []
    
    # 3) 標準スタイル取得 ＋ 日本語フォントを設定
    styles = getSampleStyleSheet()
    styles["Normal"].fontName = 'MS-Gothic'
    styles["Heading2"].fontName = 'MS-Gothic'
    
    # 4) すべてのシートを処理
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        
        # --- (A) シート見出し ---
        heading_text = f"シート名: {sheet_name}"
        heading = Paragraph(heading_text, styles["Heading2"])
        story.append(heading)
        story.append(Spacer(1, 0.5*cm))
        
        # --- (B) セルのデータをTable化 ---
        max_row = sheet.max_row
        max_col = sheet.max_column
        data = []
        
        for row in sheet.iter_rows(min_row=1, max_row=max_row,
                                   min_col=1, max_col=max_col,
                                   values_only=True):
            row_list = [(cell if cell is not None else "") for cell in row]
            data.append(row_list)
        
        if data:
            table = Table(data)
            table.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 1, colors.black),
                ('FONTNAME', (0,0), (-1,-1), 'MS-Gothic'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
            ]))
            
            story.append(table)
            story.append(Spacer(1, 0.5*cm))
        
        # --- (C) シート内の画像を取得して配置 ---
        # openpyxlのバージョンによっては非公開属性 sheet._images に画像が入っている
        for img_obj in getattr(sheet, '_images', []):
            # バイナリデータの取り出し
            img_data = img_obj._data()
            pil_img = PILImage.open(BytesIO(img_data))
            
            # Pillow -> ReportLab Image
            tmp_io = BytesIO()
            pil_img.save(tmp_io, format='PNG')
            tmp_io.seek(0)
            
            rl_img = Image(tmp_io)
            # 必要に応じて画像サイズを調整 (幅20cmに合わせる例)
            desired_width = 20 * cm
            ratio = desired_width / pil_img.width
            rl_img._restrictSize(desired_width, pil_img.height * ratio)
            
            story.append(rl_img)
            story.append(Spacer(1, 0.5*cm))
        
        # シート間の余白
        story.append(Spacer(1, 1*cm))
    
    # 5) PDF出力
    doc.build(story)
    print(f"[完了] {excel_path} -> {pdf_path}")


def convert_folder_excel_to_pdf(folder_path):
    """
    指定フォルダ内にある .xlsx ファイルをすべて検出して
    1ファイルずつPDFに変換する
    """
    # フォルダパス末尾の区切りを標準化
    folder_path = os.path.abspath(folder_path)
    
    # .xlsx のファイル一覧を取得
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    
    if not excel_files:
        print("指定フォルダに .xlsx ファイルが見つかりませんでした。")
        return
    
    print(f"対象ファイル数: {len(excel_files)}")
    
    for excel_file in excel_files:
        # PDFファイル名をExcelと同名で拡張子だけ .pdf に変更
        base_name = os.path.splitext(os.path.basename(excel_file))[0]
        pdf_name = base_name + ".pdf"
        pdf_path = os.path.join(folder_path, pdf_name)
        
        # Excel -> PDF 変換
        excel_to_pdf(excel_file, pdf_path)

    print("フォルダ内のすべての変換が完了しました。")


if __name__ == "__main__":
    # 例: "C:/path/to/excel_folder" のように指定
    target_folder = "C:/Users/YourName/Documents/excel_folder"
    convert_folder_excel_to_pdf(target_folder)

