"""건강검진 PDF 텍스트 추출기

PDF에서 텍스트를 추출합니다.
- 텍스트 기반 PDF: PyMuPDF로 직접 추출
- 이미지 기반 PDF (스캔): PyMuPDF → 이미지 변환 → Tesseract OCR
"""
import sys
import os
import json
import fitz  # PyMuPDF
import pytesseract
from PIL import Image, ImageFilter, ImageEnhance
from io import BytesIO

# Windows Tesseract 경로 설정
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


def is_text_based_page(page):
    """페이지에 추출 가능한 텍스트가 있는지 확인"""
    text = page.get_text().strip()
    # 의미 있는 텍스트가 10자 이상이면 텍스트 기반으로 판단
    return len(text) > 10


def extract_text_from_page(page):
    """PyMuPDF로 텍스트 직접 추출"""
    return page.get_text()


def preprocess_image(img):
    """OCR 정확도 향상을 위한 이미지 전처리"""
    # 그레이스케일 변환
    img = img.convert("L")
    # 대비 강화
    img = ImageEnhance.Contrast(img).enhance(2.0)
    # 샤프닝
    img = img.filter(ImageFilter.SHARPEN)
    # 이진화 (Otsu 방식 근사)
    img = img.point(lambda x: 0 if x < 140 else 255, "1")
    return img


def extract_text_via_ocr(page, dpi=300):
    """페이지를 이미지로 렌더링 후 Tesseract OCR로 텍스트 추출"""
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.tobytes("png")
    img = Image.open(BytesIO(img_data))
    img = preprocess_image(img)
    text = pytesseract.image_to_string(
        img, lang="kor+eng",
        config="--psm 6"  # 균일한 블록 텍스트로 인식
    )
    return text


def extract_pdf(pdf_path):
    """PDF에서 전체 텍스트를 추출"""
    doc = fitz.open(pdf_path)
    filename = os.path.basename(pdf_path)
    total_pages = len(doc)

    print(f"파일: {filename}")
    print(f"총 페이지: {total_pages}")
    print("-" * 60)

    pages = []
    text_pages = 0
    ocr_pages = 0

    for i, page in enumerate(doc):
        if is_text_based_page(page):
            text = extract_text_from_page(page)
            method = "text"
            text_pages += 1
        else:
            text = extract_text_via_ocr(page)
            method = "ocr"
            ocr_pages += 1

        text = text.strip()
        if text:
            pages.append({
                "page": i + 1,
                "method": method,
                "content": text
            })

        # 진행률 표시
        progress = (i + 1) / total_pages * 100
        print(f"\r  [{progress:5.1f}%] 페이지 {i+1}/{total_pages} ({method})", end="", flush=True)

    print()
    print(f"  텍스트 추출: {text_pages}p / OCR 추출: {ocr_pages}p / 유효 페이지: {len(pages)}p")
    doc.close()

    return {
        "filename": filename,
        "total_pages": total_pages,
        "text_pages": text_pages,
        "ocr_pages": ocr_pages,
        "pages": pages
    }


def save_result(result, output_path):
    """추출 결과를 JSON으로 저장"""
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"  저장 완료: {output_path}")


def main():
    health_dir = os.path.dirname(os.path.abspath(__file__))

    # 인자가 있으면 해당 파일만, 없으면 폴더 내 모든 PDF
    if len(sys.argv) > 1:
        pdf_files = [sys.argv[1]]
    else:
        pdf_files = [
            os.path.join(health_dir, f)
            for f in sorted(os.listdir(health_dir))
            if f.lower().endswith(".pdf")
        ]

    if not pdf_files:
        print("PDF 파일이 없습니다.")
        return

    data_dir = os.path.join(health_dir, "data")
    os.makedirs(data_dir, exist_ok=True)

    for pdf_path in pdf_files:
        if not os.path.exists(pdf_path):
            print(f"파일을 찾을 수 없습니다: {pdf_path}")
            continue

        print("=" * 60)
        result = extract_pdf(pdf_path)

        # JSON 저장 (파일명 기반)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join(data_dir, f"{base_name}.json")
        save_result(result, output_path)
        print()

    print("=" * 60)
    print("전체 추출 완료!")


if __name__ == "__main__":
    main()
