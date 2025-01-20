# -*- coding: utf-8 -*-
import sys
import os
import glob
from io import BytesIO
from urllib.parse import quote  # WindowsパスをURLエンコードするため
from PIL import Image as PILImage

# ▼ Windowsコンソールの文字化け対策（PowerShellの場合は "chcp 65001" が推奨）
sys.stdout.reconfigure(encoding='utf-8')

from openpyxl import load_workbook
from openpyxl.utils import coordinate_from_string

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, Image, KeepInFrame
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.units import cm

# ▼ 日本語フォントを使う（Windows標準のMSゴシックを例に）
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('MS-Gothic', 'msgothic.ttc'))


def excel_to_pdf_keep_images_near(excel_path, pdf_path):
    """
    指定した Excel(.xlsx) ファイルを読み込み、
    A3横向きのPDFに変換する。
      - 最初のページの先頭に「元ファイルへのリンク」を追加
      - すべてのシートを行単位で出力
      - 画像とデータが分離しないよう、画像アンカー行の直後に画像を配置
      - KeepInFrame + 画像縮小で "too large on page X..." エラーを回避
      - Windows上の文字化け/豆腐文字を回避するためにMSゴシックを使用
    """
    
    wb = load_workbook(excel_path, data_only=True)
    
    # PDFドキュメント(A3横向き)の準備
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A3)
    )
    story = []
    styles = getSampleStyleSheet()
    
    # 日本語フォント設定(標準スタイル上書き)
    styles["Normal"].fontName = 'MS-Gothic'
    styles["Heading2"].fontName = 'MS-Gothic'
    
    # --- (1) 最初のページに元ファイルへのリンク ---
    abs_excel_path = os.path.abspath(excel_path)
    excel_file_url = "file:///" + quote(abs_excel_path.replace("\\", "/"))  # file:/// 形式
    
    link_text = f'<link href="{excel_file_url}">元ファイル: {os.path.basename(excel_path)}</link>'
    link_paragraph = Paragraph(link_text, styles["Normal"])
    story.append(link_paragraph)
    story.append(Spacer(1, 1 * cm))
    
    # --- (2) 各シートを処理 ---
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        
        # シート名の見出し
        heading_txt = f"シート: {sheet_name}"
        heading = Paragraph(heading_txt, styles["Heading2"])
        story.append(heading)
        story.append(Spacer(1, 0.5 * cm))
        
        max_row = sheet.max_row
        max_col = sheet.max_column
        
        # (A) 画像のアンカー行を把握する
        images_in_row = {}
        for img_obj in getattr(sheet, "_images", []):
            anchor = img_obj.anchor
            rowx = None
            
            # 1) anchor が "C5" のような文字列の場合
            if isinstance(anchor, str):
                # 'C5' -> coordinate_from_string で ('C', '5')
                coord = coordinate_from_string(anchor)
                rowx = int(coord[1])
            else:
                # 2) anchor が OneCellAnchor / TwoCellAnchor オブジェクトの場合
                #    バージョンによっては row_idx + 1
                rowx = anchor._from.row_idx + 1
            
            if rowx:
                images_in_row.setdefault(rowx, []).append(img_obj)
        
        # (B) 行ごとにテーブル＋画像を追加
        for r in range(1, max_row + 1):
            # 行データ作成 (空セルは "")
            row_values = []
            for c in range(1, max_col + 1):
                val = sheet.cell(row=r, column=c).value
                row_values.append("" if val is None else str(val))
            
            # 1行N列のテーブル
            row_table = Table([row_values])
            row_table.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 1, colors.black),
                ('FONTNAME', (0,0), (-1,-1), 'MS-Gothic'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
            ]))
            
            # テーブルが大きすぎるときに備え、KeepInFrame で縮小
            wrapped_table = KeepInFrame(
                width=38*cm, height=20*cm,
                content=[row_table],
                mode='shrink'
            )
            story.append(wrapped_table)
            
            # 行に画像があれば、直後に画像を追加
            if r in images_in_row:
                for img_obj in images_in_row[r]:
                    # バイナリデータの取得 (バージョンによって _data or _data())
                    if callable(img_obj._data):
                        img_data = img_obj._data()
                    else:
                        img_data = img_obj._data
                    
                    pil_img = PILImage.open(BytesIO(img_data))
                    
                    # 一旦PNGにしてReportLabのImageに変換
                    tmp_io = BytesIO()
                    pil_img.save(tmp_io, format='PNG')
                    tmp_io.seek(0)
                    
                    rl_img = Image(tmp_io)
                    
                    # 幅15cmに合わせて縮小
                    desired_width = 15 * cm
                    ratio = desired_width / pil_img.width
                    rl_img._restrictSize(desired_width, pil_img.height * ratio)
                    
                    # KeepInFrameでフレームからはみ出る場合は更に縮小
                    wrapped_img = KeepInFrame(
                        width=38*cm, height=20*cm,
                        content=[rl_img],
                        mode='shrink'
                    )
                    story.append(wrapped_img)
            
            # 行ごとの余白
            story.append(Spacer(1, 0.3 * cm))
        
        # シートごとの区切り
        story.append(Spacer(1, 1 * cm))
    
    # --- (3) PDF出力 ---
    doc.build(story)
    print(f"[完了] {excel_path} -> {pdf_path}")


def convert_all_excel_in_subfolders(base_folder):
    """
    base_folder 以下のサブフォルダも含めて再帰的に検索し、
    見つかった *.xlsx をすべて PDF に変換する。
    PDFファイルは各 .xlsx と同じフォルダに同名で作成。
    """
    base_folder = os.path.abspath(base_folder)
    
    count = 0
    for root, dirs, files in os.walk(base_folder):
        for filename in files:
            if filename.lower().endswith(".xlsx"):
                excel_path = os.path.join(root, filename)
                base_name = os.path.splitext(filename)[0]
                pdf_name = base_name + ".pdf"
                pdf_path = os.path.join(root, pdf_name)
                
                excel_to_pdf_keep_images_near(excel_path, pdf_path)
                count += 1
    
    if count == 0:
        print("指定フォルダ以下に .xlsx ファイルが見つかりませんでした。")
    else:
        print(f"合計 {count} 件の Excel ファイルをPDF化しました。")


if __name__ == "__main__":
    # 例: C:/path/to/root_folder
    target_folder = r"C:\Users\YourName\Documents\root_folder"
    convert_all_excel_in_subfolders(target_folder)



