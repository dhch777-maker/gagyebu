import os
import uuid
from flask import Flask, request, jsonify, render_template, send_from_directory
import cv2
from extractor import process_photo, manual_process, inpaint_region
from PIL import Image
import base64
from io import BytesIO

app = Flask(__name__)

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
PROCESSED_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "processed")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".heic", ".webp"}
OUTPUT_SIZE = 1080


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "photo" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["photo"]
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": f"Unsupported format: {ext}"}), 400

    file_id = uuid.uuid4().hex[:12]
    original_name = f"{file_id}_original{ext}"
    original_path = os.path.join(UPLOAD_DIR, original_name)
    file.save(original_path)

    img = cv2.imread(original_path)
    if img is None:
        return jsonify({"error": "Cannot read image file"}), 400

    result = process_photo(img, output_size=OUTPUT_SIZE)

    if result is None or not result.success:
        return jsonify({
            "file_id": file_id,
            "original": f"/files/uploads/{original_name}",
            "processed": None,
            "hand_detected": False,
            "detection_method": result.detection_method if result else "none",
            "confidence": result.confidence if result else 0.0,
            "message": "작품 추출에 실패했습니다. 수동으로 지정해주세요.",
        })

    processed_name = f"{file_id}_cropped.jpg"
    processed_path = os.path.join(PROCESSED_DIR, processed_name)
    cv2.imwrite(processed_path, result.image, [cv2.IMWRITE_JPEG_QUALITY, 92])

    msg = "작품 추출 완료!"
    if result.hand_detected:
        msg += " (손 자동 보정됨)"

    return jsonify({
        "file_id": file_id,
        "original": f"/files/uploads/{original_name}",
        "processed": f"/files/processed/{processed_name}",
        "hand_detected": result.hand_detected,
        "detection_method": result.detection_method,
        "confidence": result.confidence,
        "message": msg,
    })


@app.route("/manual-crop", methods=["POST"])
def manual_crop():
    data = request.get_json()
    file_id = data.get("file_id")
    points = data.get("points")  # [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
    source = data.get("source", "original")  # "original" or "processed"

    if not file_id or not points or len(points) != 4:
        return jsonify({"error": "file_id와 4개의 꼭지점이 필요합니다"}), 400

    # Find source file
    if source == "processed":
        processed_name = f"{file_id}_cropped.jpg"
        source_path = os.path.join(PROCESSED_DIR, processed_name)
        if not os.path.exists(source_path):
            source_path = None
    else:
        source_path = None

    if source_path is None:
        # Fallback to original
        for f in os.listdir(UPLOAD_DIR):
            if f.startswith(file_id):
                source_path = os.path.join(UPLOAD_DIR, f)
                break

    if not source_path:
        return jsonify({"error": "파일을 찾을 수 없습니다"}), 404

    img = cv2.imread(source_path)
    if img is None:
        return jsonify({"error": "이미지를 읽을 수 없습니다"}), 400

    result = manual_process(img, points, output_size=OUTPUT_SIZE)

    processed_name = f"{file_id}_cropped.jpg"
    processed_path = os.path.join(PROCESSED_DIR, processed_name)
    cv2.imwrite(processed_path, result, [cv2.IMWRITE_JPEG_QUALITY, 92])

    return jsonify({
        "processed": f"/files/processed/{processed_name}",
        "message": "수동 보정 완료!",
    })


@app.route("/inpaint", methods=["POST"])
def inpaint():
    data = request.get_json()
    image_b64 = data.get("image")
    mask_b64 = data.get("mask")
    file_id = data.get("file_id")

    if not image_b64 or not mask_b64 or not file_id:
        return jsonify({"error": "image, mask, file_id가 필요합니다"}), 400

    # Decode base64 images
    img_data = base64.b64decode(image_b64.split(",")[1] if "," in image_b64 else image_b64)
    mask_data = base64.b64decode(mask_b64.split(",")[1] if "," in mask_b64 else mask_b64)

    img = Image.open(BytesIO(img_data)).convert("RGB")
    mask = Image.open(BytesIO(mask_data)).convert("L")

    # Run inpainting
    result = inpaint_region(img, mask)

    # Save to processed directory
    processed_name = f"{file_id}_cropped.jpg"
    processed_path = os.path.join(PROCESSED_DIR, processed_name)
    result.save(processed_path, "JPEG", quality=92)

    # Return as base64
    buf = BytesIO()
    result.save(buf, format="JPEG", quality=92)
    result_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

    return jsonify({
        "result_image": "data:image/jpeg;base64," + result_b64,
        "processed": f"/files/processed/{processed_name}",
        "message": "부분 보정 완료!",
    })


@app.route("/files/uploads/<filename>")
def serve_upload(filename):
    return send_from_directory(UPLOAD_DIR, filename)


@app.route("/files/processed/<filename>")
def serve_processed(filename):
    return send_from_directory(PROCESSED_DIR, filename)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
