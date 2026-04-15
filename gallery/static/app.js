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
    spinner.hidden = true;
    dropzone.hidden = false;
    fileInput.value = "";
    confirmBtn.hidden = false;
    retryBtn.textContent = "다시하기";
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

            if (result.data.processed) {
                processedImg.src = result.data.processed;
                message.textContent = result.data.message;
                message.style.color = "#2a7d2a";
                confirmBtn.hidden = false;
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

// Retry
retryBtn.addEventListener("click", resetUI);
failRetryBtn.addEventListener("click", resetUI);

// Confirm
confirmBtn.addEventListener("click", function() {
    message.textContent = "저장 완료! 갤러리에 등록되었습니다.";
    confirmBtn.hidden = true;
    retryBtn.textContent = "새 사진 올리기";
});
