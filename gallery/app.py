import os
import uuid
from flask import Flask, request, jsonify, render_template, send_from_directory
import cv2
import numpy as np
from extractor import process_photo, manual_crop

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
        return jsonify({"error": "Unsupported format: {}".format(ext)}), 400

    file_id = uuid.uuid4().hex[:12]
    original_name = "{}_original{}".format(file_id, ext)
    original_path = os.path.join(UPLOAD_DIR, original_name)
    file.save(original_path)

    img = cv2.imread(original_path)
    if img is None:
        return jsonify({"error": "Cannot read image file"}), 400

    result = process_photo(img, output_size=OUTPUT_SIZE)

    if result is None:
        return jsonify({
            "file_id": file_id,
            "original": "/files/uploads/{}".format(original_name),
            "processed": None,
            "message": "작품 영역을 자동으로 감지하지 못했습니다. 수동 크롭을 사용해주세요.",
        })

    processed_name = "{}_cropped.jpg".format(file_id)
    processed_path = os.path.join(PROCESSED_DIR, processed_name)
    cv2.imwrite(processed_path, result, [cv2.IMWRITE_JPEG_QUALITY, 92])

    return jsonify({
        "file_id": file_id,
        "original": "/files/uploads/{}".format(original_name),
        "processed": "/files/processed/{}".format(processed_name),
        "message": "작품 추출 완료!",
    })


@app.route("/files/uploads/<filename>")
def serve_upload(filename):
    return send_from_directory(UPLOAD_DIR, filename)


@app.route("/files/processed/<filename>")
def serve_processed(filename):
    return send_from_directory(PROCESSED_DIR, filename)


@app.route("/manual-crop", methods=["POST"])
def manual_crop_endpoint():
    data = request.get_json()
    file_id = data.get("file_id")
    corners = data.get("corners")  # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]

    if not file_id or not corners or len(corners) != 4:
        return jsonify({"error": "file_id and 4 corners required"}), 400

    # Find the original file
    original_path = None
    for fname in os.listdir(UPLOAD_DIR):
        if fname.startswith(file_id):
            original_path = os.path.join(UPLOAD_DIR, fname)
            break

    if not original_path or not os.path.exists(original_path):
        return jsonify({"error": "Original file not found"}), 404

    img = cv2.imread(original_path)
    corners_arr = np.array(corners, dtype=np.float32)

    result = manual_crop(img, corners_arr, output_size=OUTPUT_SIZE)

    processed_name = "{}_cropped.jpg".format(file_id)
    processed_path = os.path.join(PROCESSED_DIR, processed_name)
    cv2.imwrite(processed_path, result, [cv2.IMWRITE_JPEG_QUALITY, 92])

    return jsonify({
        "file_id": file_id,
        "processed": "/files/processed/{}".format(processed_name),
        "message": "수동 크롭 완료!",
    })


if __name__ == "__main__":
    app.run(debug=True, port=5000)
