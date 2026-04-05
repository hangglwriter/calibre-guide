"""
샘플 DOCX를 PDF로 변환 후, 디자인 요소 페이지를 이미지로 추출
1. Word COM → PDF 변환
2. PyMuPDF로 전체 페이지 이미지 추출
3. 디자인 요소가 있는 페이지 확인용
"""
import subprocess
import time
import os

# Step 1: Word COM으로 DOCX → PDF 변환
docx_path = r"D:\Sites\calibre-guide\docs\나도 책을 쓸 수 있을까 - 샘플원고.docx"
pdf_path = r"D:\Sites\calibre-guide\docs\sample_temp.pdf"
output_dir = r"D:\Sites\calibre-guide\docs\images"

print("=== Step 1: DOCX → PDF 변환 ===")

# 기존 Word 프로세스 확인
subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"],
               capture_output=True, text=True)
time.sleep(1)

import win32com.client

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx_path)
    # PDF로 저장 (ExportAsFixedFormat)
    doc.ExportAsFixedFormat(
        OutputFileName=pdf_path,
        ExportFormat=17,  # wdExportFormatPDF
        OptimizeFor=0,    # wdExportOptimizeForPrint
        CreateBookmarks=0
    )
    print(f"PDF 저장 완료: {pdf_path}")
    doc.Close(False)
finally:
    word.Quit()

time.sleep(1)

# Step 2: PyMuPDF로 전체 페이지 이미지 추출
print("\n=== Step 2: PDF → 페이지 이미지 추출 ===")
import fitz

pdf = fitz.open(pdf_path)
print(f"총 {len(pdf)} 페이지")

# 고해상도로 전체 페이지 렌더링 (DPI 200)
zoom = 200 / 72  # 200 DPI
mat = fitz.Matrix(zoom, zoom)

page_images = []
for i in range(len(pdf)):
    pix = pdf[i].get_pixmap(matrix=mat)
    img_path = os.path.join(output_dir, f"_temp_page_{i+1}.png")
    pix.save(img_path)
    page_images.append(img_path)
    print(f"  페이지 {i+1} → {img_path} ({pix.width}x{pix.height})")

pdf.close()

# 임시 PDF 삭제
os.remove(pdf_path)

print(f"\n총 {len(page_images)}개 페이지 이미지 생성 완료")
print("각 페이지를 확인해서 디자인 요소가 어디 있는지 찾아보세요")
