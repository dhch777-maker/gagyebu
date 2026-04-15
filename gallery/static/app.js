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

// Manual crop elements
var manualSection = document.getElementById("manualSection");
var manualHint = document.getElementById("manualHint");
var manualCanvas = document.getElementById("manualCanvas");
var manualApplyBtn = document.getElementById("manualApplyBtn");
var manualResetBtn = document.getElementById("manualResetBtn");
var manualCancelBtn = document.getElementById("manualCancelBtn");
var manualCtx = manualCanvas.getContext("2d");

// State
var currentFileId = null;
var currentOriginalUrl = null;
var manualPoints = [];
var manualImg = null;
var manualScale = 1;

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
    retryBtn.textContent = "다시하기";
    currentFileId = null;
    currentOriginalUrl = null;
    manualPoints = [];
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
                processedImg.src = result.data.processed;
                message.textContent = result.data.message;
                message.style.color = "#2a7d2a";
                confirmBtn.hidden = false;
                manualBtn.hidden = false;
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

function startManualCrop() {
    preview.hidden = true;
    failMessage.hidden = true;
    manualSection.hidden = false;
    manualPoints = [];
    manualApplyBtn.disabled = true;
    updateManualHint();

    manualImg = new Image();
    manualImg.onload = function() {
        var maxW = Math.min(700, window.innerWidth - 40);
        manualScale = Math.min(maxW / manualImg.width, 1);
        manualCanvas.width = Math.round(manualImg.width * manualScale);
        manualCanvas.height = Math.round(manualImg.height * manualScale);
        drawManualCanvas();
    };
    manualImg.src = currentOriginalUrl;
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

    // Draw lines between points
    if (manualPoints.length > 1) {
        manualCtx.beginPath();
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

manualApplyBtn.addEventListener("click", function() {
    if (manualPoints.length !== 4 || !currentFileId) return;

    manualSection.hidden = true;
    spinner.hidden = false;

    var points = manualPoints.map(function(p) { return [p.ox, p.oy]; });

    fetch("/manual-crop", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_id: currentFileId, points: points })
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

// Confirm
confirmBtn.addEventListener("click", function() {
    message.textContent = "저장 완료! 갤러리에 등록되었습니다.";
    confirmBtn.hidden = true;
    manualBtn.hidden = true;
    retryBtn.textContent = "새 사진 올리기";
});

}); // end DOMContentLoaded
