var dropzone = document.getElementById("dropzone");
var fileInput = document.getElementById("fileInput");
var selectBtn = document.getElementById("selectBtn");
var spinner = document.getElementById("spinner");
var preview = document.getElementById("preview");
var originalImg = document.getElementById("originalImg");
var processedImg = document.getElementById("processedImg");
var message = document.getElementById("message");
var confirmBtn = document.getElementById("confirmBtn");
var retryBtn = document.getElementById("retryBtn");
var failMessage = document.getElementById("failMessage");
var failText = document.getElementById("failText");
var failRetryBtn = document.getElementById("failRetryBtn");

// Manual crop elements
var cropCanvas = document.getElementById("cropCanvas");
var cropCtx = cropCanvas.getContext("2d");
var cropConfirmBtn = document.getElementById("cropConfirmBtn");
var cropResetBtn = document.getElementById("cropResetBtn");
var manualPoints = [];
var currentFileId = null;
var currentOriginalUrl = null;

// Drag and drop
dropzone.addEventListener("dragover", function(e) {
    e.preventDefault();
    dropzone.classList.add("dragover");
});
dropzone.addEventListener("dragleave", function() {
    dropzone.classList.remove("dragover");
});
dropzone.addEventListener("drop", function(e) {
    e.preventDefault();
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
    spinner.hidden = true;
    dropzone.hidden = false;
    fileInput.value = "";
    confirmBtn.hidden = false;
    retryBtn.textContent = "다시하기";
    manualPoints = [];
    currentFileId = null;
    currentOriginalUrl = null;
    cropConfirmBtn.hidden = true;
    cropResetBtn.hidden = true;
}

function uploadFile(file) {
    dropzone.hidden = true;
    spinner.hidden = false;
    preview.hidden = true;
    failMessage.hidden = true;

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

            originalImg.src = result.data.original;
            currentFileId = result.data.file_id;
            currentOriginalUrl = result.data.original;

            if (result.data.processed) {
                processedImg.src = result.data.processed;
                message.textContent = result.data.message;
                message.style.color = "#2a7d2a";
                confirmBtn.hidden = false;
                preview.hidden = false;
            } else {
                failText.textContent = result.data.message;
                failMessage.hidden = false;
                showManualCropCanvas(result.data.original);
            }
        })
        .catch(function() {
            spinner.hidden = true;
            failText.textContent = "서버 연결에 실패했습니다.";
            failMessage.hidden = false;
        });
}

// Manual crop canvas
function showManualCropCanvas(originalUrl) {
    var img = new Image();
    img.onload = function() {
        var maxW = 700;
        var scale = img.width > maxW ? maxW / img.width : 1;
        cropCanvas.width = img.width * scale;
        cropCanvas.height = img.height * scale;
        cropCanvas.dataset.scale = String(1 / scale);
        cropCtx.drawImage(img, 0, 0, cropCanvas.width, cropCanvas.height);
        cropCanvas.hidden = false;
        manualPoints = [];
    };
    img.src = originalUrl;
}

cropCanvas.addEventListener("click", function(e) {
    if (manualPoints.length >= 4) return;
    var rect = cropCanvas.getBoundingClientRect();
    var x = e.clientX - rect.left;
    var y = e.clientY - rect.top;
    manualPoints.push([x, y]);

    // Draw point
    cropCtx.beginPath();
    cropCtx.arc(x, y, 6, 0, Math.PI * 2);
    cropCtx.fillStyle = "#e74c3c";
    cropCtx.fill();
    cropCtx.fillStyle = "white";
    cropCtx.font = "bold 12px sans-serif";
    cropCtx.fillText(String(manualPoints.length), x - 4, y + 4);

    if (manualPoints.length === 4) {
        // Draw outline
        cropCtx.beginPath();
        cropCtx.moveTo(manualPoints[0][0], manualPoints[0][1]);
        for (var i = 1; i < 4; i++) {
            cropCtx.lineTo(manualPoints[i][0], manualPoints[i][1]);
        }
        cropCtx.closePath();
        cropCtx.strokeStyle = "#e74c3c";
        cropCtx.lineWidth = 2;
        cropCtx.stroke();
        cropConfirmBtn.hidden = false;
        cropResetBtn.hidden = false;
    }
});

cropResetBtn.addEventListener("click", function() {
    manualPoints = [];
    cropConfirmBtn.hidden = true;
    cropResetBtn.hidden = true;
    var img = new Image();
    img.onload = function() {
        cropCtx.drawImage(img, 0, 0, cropCanvas.width, cropCanvas.height);
    };
    img.src = currentOriginalUrl;
});

cropConfirmBtn.addEventListener("click", function() {
    var scale = parseFloat(cropCanvas.dataset.scale);
    var scaledPoints = manualPoints.map(function(pt) {
        return [Math.round(pt[0] * scale), Math.round(pt[1] * scale)];
    });

    spinner.hidden = false;
    failMessage.hidden = true;

    fetch("/manual-crop", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_id: currentFileId, corners: scaledPoints }),
    })
    .then(function(resp) { return resp.json(); })
    .then(function(data) {
        spinner.hidden = true;

        if (data.processed) {
            processedImg.src = data.processed;
            message.textContent = data.message;
            message.style.color = "#2a7d2a";
            confirmBtn.hidden = false;
            preview.hidden = false;
        }
    })
    .catch(function() {
        spinner.hidden = true;
        failText.textContent = "수동 크롭 처리에 실패했습니다.";
        failMessage.hidden = false;
    });
});

// Retry
retryBtn.addEventListener("click", resetUI);
failRetryBtn.addEventListener("click", resetUI);

// Confirm
confirmBtn.addEventListener("click", function() {
    message.textContent = "저장 완료! 갤러리에 등록되었습니다.";
    confirmBtn.hidden = true;
    retryBtn.textContent = "새 사진 올리기";
});
