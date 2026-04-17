"""건강검진 PDF 텍스트 추출기

PDF에서 텍스트를 추출합니다.
- 텍스트 기반 PDF: PyMuPDF로 직접 추출
- 이미지 기반 PDF (스캔): PyMuPDF → 이미지 변환 → OpenCV 전처리 → EasyOCR
"""
import sys
import os
import json
import numpy as np
import cv2
import fitz  # PyMuPDF
import easyocr

# EasyOCR 인스턴스 (한국어+영어, GPU 미사용)
reader = easyocr.Reader(["ko", "en"], gpu=False)


def is_text_based_page(page):
    """페이지에 추출 가능한 텍스트가 있는지 확인"""
    text = page.get_text().strip()
    return len(text) > 10


def extract_text_from_page(page):
    """PyMuPDF로 텍스트 직접 추출"""
    return page.get_text()


def preprocess_image(img_array):
    """OCR 정확도 향상을 위한 OpenCV 전처리

    1. CLAHE 대비 강화 (적응형 히스토그램 균등화)
    2. 가우시안 블러 (노이즈 제거)
    3. 언샤프 마스크 (선명도 강화)
    """
    # 그레이스케일 변환
    if len(img_array.shape) == 3:
        gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
    else:
        gray = img_array

    # CLAHE 대비 강화
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)

    # 가우시안 블러로 노이즈 제거
    blurred = cv2.GaussianBlur(enhanced, (3, 3), 0)

    # 언샤프 마스크로 선명도 강화
    sharpened = cv2.addWeighted(enhanced, 1.5, blurred, -0.5, 0)

    return sharpened


def extract_text_via_ocr(page, dpi=300):
    """페이지를 이미지로 렌더링 후 EasyOCR로 텍스트 추출"""
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)

    # PyMuPDF pixmap → numpy array
    img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n == 4:  # RGBA → BGR
        img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)
    elif pix.n == 1:  # grayscale → BGR
        img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2BGR)

    # OpenCV 전처리
    preprocessed = preprocess_image(img_array)

    # EasyOCR 실행
    results = reader.readtext(preprocessed, detail=0, paragraph=True)

    return "\n".join(results)


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

    data_dir = os.path.join(health_dir, "data", "raw")
    os.makedirs(data_dir, exist_ok=True)

    for pdf_path in pdf_files:
        if not os.path.exists(pdf_path):
            print(f"파일을 찾을 수 없습니다: {pdf_path}")
            continue

        print("=" * 60)
        result = extract_pdf(pdf_path)

        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join(data_dir, f"{base_name}.json")
        save_result(result, output_path)
        print()

    print("=" * 60)
    print("전체 추출 완료!")


if __name__ == "__main__":
    main()
