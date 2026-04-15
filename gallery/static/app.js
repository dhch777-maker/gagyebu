document.addEventListener("DOMContentLoaded", function() {

var dropzone = document.getElementById("dropzone");
var fileInput = document.getElementById("fileInput");
var selectBtn = document.getElementById("selectBtn");
var spinner = document.getElementById("spinner");
var preview = document.getElementById("preview");
var originalImg = document.getElementById("originalImg");
var processedImg = document.getElementById("processedImg");
var message = document.getElementById("message");
var confirmBtn = document.getElementById("confirmBtn");
var manualBtn = document.getElementById("manualBtn");
var retryBtn = document.getElementById("retryBtn");
var failMessage = document.getElementById("failMessage");
var failText = document.getElementById("failText");
var failManualBtn = document.getElementById("failManualBtn");
var failRetryBtn = document.getElementById("failRetryBtn");

// Inpainting elements
var inpaintBtn = document.getElementById("inpaintBtn");
var inpaintSection = document.getElementById("inpaintSection");
var inpaintCanvas = document.getElementById("inpaintCanvas");
var inpaintCtx = inpaintCanvas.getContext("2d");
var brushSize = document.getElementById("brushSize");
var brushSizeLabel = document.getElementById("brushSizeLabel");
var inpaintUndoBtn = document.getElementById("inpaintUndoBtn");
var inpaintDoneBtn = document.getElementById("inpaintDoneBtn");
var inpaintCancelBtn = document.getElementById("inpaintCancelBtn");

// Manual crop elements
var manualSection = document.getElementById("manualSection");
var manualHint = document.getElementById("manualHint");
var manualCanvas = document.getElementById("manualCanvas");
var manualApplyBtn = document.getElementById("manualApplyBtn");
var manualResetBtn = document.getElementById("manualResetBtn");
var manualCancelBtn = document.getElementById("manualCancelBtn");
var manualCtx = manualCanvas.getContext("2d");

// Source toggle elements
var srcOriginalBtn = document.getElementById("srcOriginalBtn");
var srcProcessedBtn = document.getElementById("srcProcessedBtn");

// State
var currentFileId = null;
var currentOriginalUrl = null;
var currentProcessedUrl = null;
var manualSource = "original"; // "original" or "processed"
var manualPoints = [];
var manualImg = null;
var manualScale = 1;

// Inpainting state
var inpaintImg = null;
var inpaintScale = 1;
var inpaintDrawing = false;
var inpaintMaskCanvas = null;
var inpaintMaskCtx = null;
var inpaintHistory = [];

// Prevent browser from opening dropped files
document.addEventListener("dragover", function(e) { e.preventDefault(); });
document.addEventListener("drop", function(e) { e.preventDefault(); });

// Drag and drop
dropzone.addEventListener("dragover", function(e) {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.add("dragover");
});
dropzone.addEventListener("dragleave", function(e) {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove("dragover");
});
dropzone.addEventListener("drop", function(e) {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files.length > 0) {
        uploadFile(e.dataTransfer.files[0]);
    }
});

// Click to select
selectBtn.addEventListener("click", function(e) {
    e.stopPropagation();
    fileInput.click();
});
dropzone.addEventListener("click", function() { fileInput.click(); });
fileInput.addEventListener("change", function() {
    if (fileInput.files.length > 0) {
        uploadFile(fileInput.files[0]);
    }
});

function resetUI() {
    preview.hidden = true;
    failMessage.hidden = true;
    manualSection.hidden = true;
    spinner.hidden = true;
    dropzone.hidden = false;
    fileInput.value = "";
    confirmBtn.hidden = false;
    manualBtn.hidden = false;
    inpaintBtn.hidden = false;
    retryBtn.textContent = "다시하기";
    currentFileId = null;
    currentOriginalUrl = null;
    currentProcessedUrl = null;
    manualPoints = [];
    inpaintSection.hidden = true;
    inpaintHistory = [];
    inpaintUndoBtn.disabled = true;
}

function uploadFile(file) {
    dropzone.hidden = true;
    spinner.hidden = false;
    preview.hidden = true;
    failMessage.hidden = true;
    manualSection.hidden = true;

    var formData = new FormData();
    formData.append("photo", file);

    fetch("/upload", { method: "POST", body: formData })
        .then(function(resp) {
            return resp.json().then(function(data) {
                return { ok: resp.ok, data: data };
            });
        })
        .then(function(result) {
            spinner.hidden = true;

            if (!result.ok) {
                failText.textContent = result.data.error || "업로드 실패";
                failMessage.hidden = false;
                return;
            }

            currentFileId = result.data.file_id;
            currentOriginalUrl = result.data.original;
            originalImg.src = result.data.original;

            if (result.data.processed) {
                currentProcessedUrl = result.data.processed;
                processedImg.src = result.data.processed;
                message.textContent = result.data.message;
                message.style.color = "#2a7d2a";
                confirmBtn.hidden = false;
                manualBtn.hidden = false;
                inpaintBtn.hidden = false;
                preview.hidden = false;
            } else {
                failText.textContent = result.data.message;
                failMessage.hidden = false;
            }
        })
        .catch(function() {
            spinner.hidden = true;
            failText.textContent = "서버 연결에 실패했습니다.";
            failMessage.hidden = false;
        });
}

// --- Manual Crop ---

var manualMousePos = null; // Track mouse position for preview lines

function startManualCrop() {
    preview.hidden = true;
    failMessage.hidden = true;
    manualSection.hidden = false;
    manualPoints = [];
    manualMousePos = null;
    manualApplyBtn.disabled = true;
    manualSource = "original";
    srcOriginalBtn.classList.add("active");
    srcProcessedBtn.classList.remove("active");

    // Enable/disable processed button based on availability
    if (currentProcessedUrl) {
        srcProcessedBtn.disabled = false;
    } else {
        srcProcessedBtn.disabled = true;
    }

    updateManualHint();
    loadManualImage();
}

function loadManualImage() {
    var url = (manualSource === "processed" && currentProcessedUrl)
        ? currentProcessedUrl : currentOriginalUrl;

    manualImg = new Image();
    manualImg.onload = function() {
        var maxW = Math.min(700, window.innerWidth - 40);
        manualScale = Math.min(maxW / manualImg.width, 1);
        manualCanvas.width = Math.round(manualImg.width * manualScale);
        manualCanvas.height = Math.round(manualImg.height * manualScale);
        drawManualCanvas();
    };
    manualImg.src = url;
}

function updateManualHint() {
    var remaining = 4 - manualPoints.length;
    if (remaining > 0) {
        manualHint.textContent = "그림의 꼭지점을 클릭하세요 (남은 점: " + remaining + "개)";
    } else {
        manualHint.textContent = "4개 점 지정 완료! '적용' 버튼을 눌러주세요";
    }
}

function drawManualCanvas() {
    manualCtx.drawImage(manualImg, 0, 0, manualCanvas.width, manualCanvas.height);

    // Draw solid lines between confirmed points
    if (manualPoints.length > 1) {
        manualCtx.beginPath();
        manualCtx.setLineDash([]);
        manualCtx.moveTo(manualPoints[0].cx, manualPoints[0].cy);
        for (var i = 1; i < manualPoints.length; i++) {
            manualCtx.lineTo(manualPoints[i].cx, manualPoints[i].cy);
        }
        if (manualPoints.length === 4) {
            manualCtx.closePath();
        }
        manualCtx.strokeStyle = "rgba(74, 144, 217, 0.9)";
        manualCtx.lineWidth = 2;
        manualCtx.stroke();
    }

    // Draw dashed preview lines to mouse position
    if (manualMousePos && manualPoints.length > 0 && manualPoints.length < 4) {
        var last = manualPoints[manualPoints.length - 1];
        manualCtx.beginPath();
        manualCtx.setLineDash([6, 4]);
        manualCtx.strokeStyle = "rgba(220, 30, 30, 0.8)";
        manualCtx.lineWidth = 2;

        // Last point → mouse position
        manualCtx.moveTo(last.cx, last.cy);
        manualCtx.lineTo(manualMousePos.cx, manualMousePos.cy);

        // Mouse position → first point (closing preview, when 2+ points exist)
        if (manualPoints.length >= 2) {
            manualCtx.moveTo(manualMousePos.cx, manualMousePos.cy);
            manualCtx.lineTo(manualPoints[0].cx, manualPoints[0].cy);
        }

        manualCtx.stroke();
        manualCtx.setLineDash([]);
    }

    // Semi-transparent fill when 4 points completed
    if (manualPoints.length === 4) {
        manualCtx.beginPath();
        manualCtx.moveTo(manualPoints[0].cx, manualPoints[0].cy);
        for (var i = 1; i < manualPoints.length; i++) {
            manualCtx.lineTo(manualPoints[i].cx, manualPoints[i].cy);
        }
        manualCtx.closePath();
        manualCtx.fillStyle = "rgba(220, 30, 30, 0.25)";
        manualCtx.fill();
    }

    // Draw points
    for (var i = 0; i < manualPoints.length; i++) {
        var p = manualPoints[i];

        // Outer circle
        manualCtx.beginPath();
        manualCtx.arc(p.cx, p.cy, 10, 0, Math.PI * 2);
        manualCtx.fillStyle = "rgba(74, 144, 217, 0.85)";
        manualCtx.fill();
        manualCtx.strokeStyle = "white";
        manualCtx.lineWidth = 2;
        manualCtx.stroke();

        // Number
        manualCtx.fillStyle = "white";
        manualCtx.font = "bold 12px sans-serif";
        manualCtx.textAlign = "center";
        manualCtx.textBaseline = "middle";
        manualCtx.fillText((i + 1).toString(), p.cx, p.cy);
    }
}

manualCanvas.addEventListener("click", function(e) {
    if (manualPoints.length >= 4) return;

    var rect = manualCanvas.getBoundingClientRect();
    var cx = e.clientX - rect.left;
    var cy = e.clientY - rect.top;

    // Convert to original image coordinates
    var origX = cx / manualScale;
    var origY = cy / manualScale;

    manualPoints.push({ cx: cx, cy: cy, ox: origX, oy: origY });
    updateManualHint();
    drawManualCanvas();

    if (manualPoints.length === 4) {
        manualApplyBtn.disabled = false;
    }
});

// Mouse tracking for preview lines
manualCanvas.addEventListener("mousemove", function(e) {
    if (manualPoints.length === 0 || manualPoints.length >= 4) return;
    var rect = manualCanvas.getBoundingClientRect();
    manualMousePos = { cx: e.clientX - rect.left, cy: e.clientY - rect.top };
    drawManualCanvas();
});

manualCanvas.addEventListener("mouseleave", function() {
    manualMousePos = null;
    drawManualCanvas();
});

// Source toggle
srcOriginalBtn.addEventListener("click", function() {
    if (manualSource === "original") return;
    manualSource = "original";
    srcOriginalBtn.classList.add("active");
    srcProcessedBtn.classList.remove("active");
    manualPoints = [];
    manualApplyBtn.disabled = true;
    updateManualHint();
    loadManualImage();
});
srcProcessedBtn.addEventListener("click", function() {
    if (manualSource === "processed" || !currentProcessedUrl) return;
    manualSource = "processed";
    srcProcessedBtn.classList.add("active");
    srcOriginalBtn.classList.remove("active");
    manualPoints = [];
    manualApplyBtn.disabled = true;
    updateManualHint();
    loadManualImage();
});

manualApplyBtn.addEventListener("click", function() {
    if (manualPoints.length !== 4 || !currentFileId) return;

    manualSection.hidden = true;
    spinner.hidden = false;

    var points = manualPoints.map(function(p) { return [p.ox, p.oy]; });

    fetch("/manual-crop", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_id: currentFileId, points: points, source: manualSource })
    })
    .then(function(resp) {
        return resp.json().then(function(data) {
            return { ok: resp.ok, data: data };
        });
    })
    .then(function(result) {
        spinner.hidden = true;

        if (!result.ok) {
            failText.textContent = result.data.error || "수동 보정 실패";
            failMessage.hidden = false;
            return;
        }

        processedImg.src = result.data.processed + "?t=" + Date.now();
        message.textContent = result.data.message;
        message.style.color = "#2a7d2a";
        confirmBtn.hidden = false;
        manualBtn.hidden = false;
        preview.hidden = false;
    })
    .catch(function() {
        spinner.hidden = true;
        failText.textContent = "서버 연결에 실패했습니다.";
        failMessage.hidden = false;
    });
});

manualResetBtn.addEventListener("click", function() {
    manualPoints = [];
    manualApplyBtn.disabled = true;
    updateManualHint();
    drawManualCanvas();
});

manualCancelBtn.addEventListener("click", function() {
    manualSection.hidden = true;
    if (processedImg.src) {
        preview.hidden = false;
    } else {
        failMessage.hidden = false;
    }
});

// Manual crop buttons
manualBtn.addEventListener("click", startManualCrop);
failManualBtn.addEventListener("click", startManualCrop);

// Retry
retryBtn.addEventListener("click", resetUI);
failRetryBtn.addEventListener("click", resetUI);

// --- Inpainting Mode ---

function startInpaint() {
    preview.hidden = true;
    failMessage.hidden = true;
    manualSection.hidden = true;
    inpaintSection.hidden = false;
    inpaintHistory = [];
    inpaintUndoBtn.disabled = true;
    loadInpaintImage();
}

function loadInpaintImage() {
    var url = currentProcessedUrl || currentOriginalUrl;
    inpaintImg = new Image();
    inpaintImg.onload = function() {
        var maxW = Math.min(700, window.innerWidth - 40);
        inpaintScale = Math.min(maxW / inpaintImg.width, 1);
        inpaintCanvas.width = Math.round(inpaintImg.width * inpaintScale);
        inpaintCanvas.height = Math.round(inpaintImg.height * inpaintScale);

        // Create offscreen mask canvas
        inpaintMaskCanvas = document.createElement("canvas");
        inpaintMaskCanvas.width = inpaintCanvas.width;
        inpaintMaskCanvas.height = inpaintCanvas.height;
        inpaintMaskCtx = inpaintMaskCanvas.getContext("2d");
        inpaintMaskCtx.fillStyle = "black";
        inpaintMaskCtx.fillRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);

        drawInpaintCanvas();
    };
    inpaintImg.src = url + "?t=" + Date.now();
}

function drawInpaintCanvas() {
    inpaintCtx.drawImage(inpaintImg, 0, 0, inpaintCanvas.width, inpaintCanvas.height);

    // Overlay red tint where mask is white
    var maskData = inpaintMaskCtx.getImageData(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
    var overlay = inpaintCtx.getImageData(0, 0, inpaintCanvas.width, inpaintCanvas.height);
    for (var i = 0; i < maskData.data.length; i += 4) {
        if (maskData.data[i] > 128) {
            overlay.data[i] = Math.min(255, overlay.data[i] + 100);
            overlay.data[i + 1] = Math.max(0, overlay.data[i + 1] - 50);
            overlay.data[i + 2] = Math.max(0, overlay.data[i + 2] - 50);
            overlay.data[i + 3] = 200;
        }
    }
    inpaintCtx.putImageData(overlay, 0, 0);
}

function drawBrushCursor(x, y) {
    var r = parseInt(brushSize.value) * inpaintScale;
    inpaintCtx.beginPath();
    inpaintCtx.arc(x, y, r, 0, Math.PI * 2);
    inpaintCtx.strokeStyle = "rgba(255, 255, 255, 0.8)";
    inpaintCtx.lineWidth = 2;
    inpaintCtx.stroke();
}

function paintMask(x, y) {
    var r = parseInt(brushSize.value) * inpaintScale;
    inpaintMaskCtx.beginPath();
    inpaintMaskCtx.arc(x, y, r, 0, Math.PI * 2);
    inpaintMaskCtx.fillStyle = "white";
    inpaintMaskCtx.fill();
}

function sendInpaintRequest() {
    var spinnerEl = document.createElement("div");
    spinnerEl.className = "inpaint-spinner";
    spinnerEl.textContent = "AI 보정 중...";
    inpaintCanvas.parentElement.appendChild(spinnerEl);

    // Save current state to undo history
    inpaintHistory.push(inpaintImg.src);
    inpaintUndoBtn.disabled = false;

    // Get current image as base64 at original resolution
    var imgCanvas = document.createElement("canvas");
    imgCanvas.width = inpaintImg.naturalWidth;
    imgCanvas.height = inpaintImg.naturalHeight;
    var imgCtx = imgCanvas.getContext("2d");
    imgCtx.drawImage(inpaintImg, 0, 0);
    var imageB64 = imgCanvas.toDataURL("image/jpeg", 0.92);

    // Scale mask to original resolution
    var maskFullCanvas = document.createElement("canvas");
    maskFullCanvas.width = inpaintImg.naturalWidth;
    maskFullCanvas.height = inpaintImg.naturalHeight;
    var maskFullCtx = maskFullCanvas.getContext("2d");
    maskFullCtx.drawImage(inpaintMaskCanvas, 0, 0, maskFullCanvas.width, maskFullCanvas.height);
    var maskB64 = maskFullCanvas.toDataURL("image/png");

    fetch("/inpaint", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            image: imageB64,
            mask: maskB64,
            file_id: currentFileId
        })
    })
    .then(function(resp) {
        return resp.json().then(function(data) {
            return { ok: resp.ok, data: data };
        });
    })
    .then(function(result) {
        spinnerEl.remove();
        if (!result.ok) {
            alert(result.data.error || "보정 실패");
            return;
        }
        currentProcessedUrl = result.data.processed;
        inpaintImg = new Image();
        inpaintImg.onload = function() {
            inpaintMaskCtx.fillStyle = "black";
            inpaintMaskCtx.fillRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
            drawInpaintCanvas();
        };
        inpaintImg.src = result.data.result_image;
    })
    .catch(function() {
        spinnerEl.remove();
        alert("서버 연결에 실패했습니다.");
    });
}

// Inpainting canvas events
inpaintCanvas.addEventListener("mousedown", function(e) {
    inpaintDrawing = true;
    var rect = inpaintCanvas.getBoundingClientRect();
    paintMask(e.clientX - rect.left, e.clientY - rect.top);
    drawInpaintCanvas();
    drawBrushCursor(e.clientX - rect.left, e.clientY - rect.top);
});

inpaintCanvas.addEventListener("mousemove", function(e) {
    var rect = inpaintCanvas.getBoundingClientRect();
    var x = e.clientX - rect.left;
    var y = e.clientY - rect.top;
    if (inpaintDrawing) {
        paintMask(x, y);
    }
    drawInpaintCanvas();
    drawBrushCursor(x, y);
});

inpaintCanvas.addEventListener("mouseup", function() {
    if (!inpaintDrawing) return;
    inpaintDrawing = false;
    var maskData = inpaintMaskCtx.getImageData(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
    var hasWhite = false;
    for (var i = 0; i < maskData.data.length; i += 4) {
        if (maskData.data[i] > 128) { hasWhite = true; break; }
    }
    if (hasWhite) {
        sendInpaintRequest();
    }
});

inpaintCanvas.addEventListener("mouseleave", function() {
    if (inpaintDrawing) {
        inpaintDrawing = false;
        sendInpaintRequest();
    }
    drawInpaintCanvas();
});

brushSize.addEventListener("input", function() {
    brushSizeLabel.textContent = brushSize.value + "px";
});

inpaintBtn.addEventListener("click", startInpaint);

inpaintUndoBtn.addEventListener("click", function() {
    if (inpaintHistory.length === 0) return;
    var prevSrc = inpaintHistory.pop();
    if (inpaintHistory.length === 0) inpaintUndoBtn.disabled = true;
    inpaintImg = new Image();
    inpaintImg.onload = function() {
        inpaintMaskCtx.fillStyle = "black";
        inpaintMaskCtx.fillRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
        drawInpaintCanvas();
    };
    inpaintImg.src = prevSrc;
});

inpaintDoneBtn.addEventListener("click", function() {
    inpaintSection.hidden = true;
    processedImg.src = (currentProcessedUrl || "") + "?t=" + Date.now();
    message.textContent = "부분 보정 완료!";
    message.style.color = "#2a7d2a";
    confirmBtn.hidden = false;
    manualBtn.hidden = false;
    inpaintBtn.hidden = false;
    preview.hidden = false;
});

inpaintCancelBtn.addEventListener("click", function() {
    inpaintSection.hidden = true;
    if (currentProcessedUrl) {
        preview.hidden = false;
    } else {
        failMessage.hidden = false;
    }
});

// Confirm
confirmBtn.addEventListener("click", function() {
    message.textContent = "저장 완료! 갤러리에 등록되었습니다.";
    confirmBtn.hidden = true;
    manualBtn.hidden = true;
    inpaintBtn.hidden = true;
    retryBtn.textContent = "새 사진 올리기";
});

}); // end DOMContentLoaded
