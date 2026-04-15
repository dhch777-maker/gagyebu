/* ===== app.js — Mobile 3-step Wizard ===== */

document.addEventListener("DOMContentLoaded", function () {

    // ── DOM refs ──────────────────────────────────────────────────────────
    var progressBar    = document.getElementById("progressBar");
    var progressLabel  = document.getElementById("progressLabel");
    var progressSteps  = document.querySelectorAll(".progress-step");

    // Step 1
    var step1          = document.getElementById("step1");
    var cameraBtn      = document.getElementById("cameraBtn");
    var galleryBtn     = document.getElementById("galleryBtn");
    var fileInput      = document.getElementById("fileInput");
    var galleryInput   = document.getElementById("galleryInput");

    // Loading
    var loadingView    = document.getElementById("loadingView");
    var loadingText    = document.getElementById("loadingText");
    var lsDetect       = document.getElementById("lsDetect");
    var lsTransform    = document.getElementById("lsTransform");
    var lsHand         = document.getElementById("lsHand");
    var lsFinish       = document.getElementById("lsFinish");

    // Step 2 (success)
    var step2          = document.getElementById("step2");
    var reviewImg      = document.getElementById("reviewImg");
    var handBadge      = document.getElementById("handBadge");
    var confidenceBadge= document.getElementById("confidenceBadge");
    var reviewMessage  = document.getElementById("reviewMessage");
    var confirmBtn     = document.getElementById("confirmBtn");
    var adjustCornersBtn = document.getElementById("adjustCornersBtn");
    var touchupBtn     = document.getElementById("touchupBtn");
    var retryBtn       = document.getElementById("retryBtn");

    // Step 2 (fail)
    var step2fail      = document.getElementById("step2fail");
    var failText       = document.getElementById("failText");
    var failManualBtn  = document.getElementById("failManualBtn");
    var failRetryBtn   = document.getElementById("failRetryBtn");

    // Step 3
    var step3          = document.getElementById("step3");
    var tabCorners     = document.getElementById("tabCorners");
    var tabBrush       = document.getElementById("tabBrush");
    var cornerMode     = document.getElementById("cornerMode");
    var brushMode      = document.getElementById("brushMode");
    var cornerHint     = document.getElementById("cornerHint");
    var cornerCanvas   = document.getElementById("cornerCanvas");
    var cornerApplyBtn = document.getElementById("cornerApplyBtn");
    var cornerResetBtn = document.getElementById("cornerResetBtn");
    var cornerCancelBtn= document.getElementById("cornerCancelBtn");
    var brushSize      = document.getElementById("brushSize");
    var brushSizeLabel = document.getElementById("brushSizeLabel");
    var brushCanvas    = document.getElementById("brushCanvas");
    var brushUndoBtn   = document.getElementById("brushUndoBtn");
    var brushDoneBtn   = document.getElementById("brushDoneBtn");
    var brushCancelBtn = document.getElementById("brushCancelBtn");

    // Saved
    var savedView      = document.getElementById("savedView");
    var newPhotoBtn    = document.getElementById("newPhotoBtn");

    var cornerCtx = cornerCanvas.getContext("2d");
    var brushCtx  = brushCanvas.getContext("2d");

    // ── State ─────────────────────────────────────────────────────────────
    var currentFileId        = null;
    var currentOriginalUrl   = null;
    var currentProcessedUrl  = null;

    // Corner mode
    var cornerImg   = null;
    var cornerScale = 1;
    var cornerPoints= [];
    var cornerMousePos = null;

    // Brush mode
    var brushImg        = null;
    var brushScale      = 1;
    var brushDrawing    = false;
    var brushMaskCanvas = null;
    var brushMaskCtx    = null;
    var brushHistory    = [];   // array of image src strings

    // ── Helpers ───────────────────────────────────────────────────────────
    var ALL_STEPS = [step1, loadingView, step2, step2fail, step3, savedView];

    function showStep(el, progressNum) {
        ALL_STEPS.forEach(function (s) { s.hidden = true; });
        el.hidden = false;

        if (progressNum) {
            progressBar.hidden = false;
            progressLabel.textContent = progressNum + "/3";
            progressSteps.forEach(function (ps, i) {
                ps.classList.remove("active", "done");
                var n = parseInt(ps.dataset.step);
                if (n < progressNum) ps.classList.add("done");
                else if (n === progressNum) ps.classList.add("active");
            });
        } else {
            progressBar.hidden = true;
        }
    }

    // Animate loading steps with staggered highlights
    var loadingTimer = null;
    var loadingStepEls = [lsDetect, lsTransform, lsHand, lsFinish];

    function startLoadingAnimation() {
        loadingStepEls.forEach(function (el) {
            el.classList.remove("active", "done");
        });
        var idx = 0;
        loadingTimer = setInterval(function () {
            if (idx > 0) loadingStepEls[idx - 1].classList.replace("active", "done");
            if (idx < loadingStepEls.length) {
                loadingStepEls[idx].classList.add("active");
                idx++;
            } else {
                clearInterval(loadingTimer);
            }
        }, 700);
    }

    function stopLoadingAnimation() {
        clearInterval(loadingTimer);
        loadingStepEls.forEach(function (el) {
            el.classList.remove("active");
            el.classList.add("done");
        });
    }

    // ── Client-side image resize (max 2048px) ─────────────────────────────
    function resizeImageFile(file, maxPx, callback) {
        var reader = new FileReader();
        reader.onload = function (e) {
            var img = new Image();
            img.onload = function () {
                var w = img.width, h = img.height;
                if (w <= maxPx && h <= maxPx) {
                    callback(file); // no resize needed
                    return;
                }
                var scale = Math.min(maxPx / w, maxPx / h);
                var canvas = document.createElement("canvas");
                canvas.width  = Math.round(w * scale);
                canvas.height = Math.round(h * scale);
                canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
                canvas.toBlob(function (blob) {
                    callback(new File([blob], file.name, { type: "image/jpeg" }));
                }, "image/jpeg", 0.92);
            };
            img.src = e.target.result;
        };
        reader.readAsDataURL(file);
    }

    // ── Upload ────────────────────────────────────────────────────────────
    function uploadFile(file) {
        showStep(loadingView, 1);
        startLoadingAnimation();

        resizeImageFile(file, 2048, function (resized) {
            var formData = new FormData();
            formData.append("photo", resized);

            fetch("/upload", { method: "POST", body: formData })
                .then(function (resp) {
                    return resp.json().then(function (data) {
                        return { ok: resp.ok, data: data };
                    });
                })
                .then(function (result) {
                    stopLoadingAnimation();
                    fileInput.value = "";
                    galleryInput.value = "";

                    if (!result.ok) {
                        failText.textContent = result.data.error || "업로드 실패";
                        showStep(step2fail, 2);
                        return;
                    }

                    currentFileId       = result.data.file_id;
                    currentOriginalUrl  = result.data.original;
                    currentProcessedUrl = result.data.processed || null;

                    if (result.data.processed) {
                        reviewImg.src = result.data.processed + "?t=" + Date.now();

                        // Hand badge
                        if (result.data.hand_detected) {
                            handBadge.hidden = false;
                        } else {
                            handBadge.hidden = true;
                        }

                        // Confidence badge
                        if (result.data.confidence != null) {
                            var pct = Math.round(result.data.confidence * 100);
                            confidenceBadge.textContent = "신뢰도 " + pct + "%";
                            confidenceBadge.hidden = false;
                        } else {
                            confidenceBadge.hidden = true;
                        }

                        reviewMessage.textContent = result.data.message || "";
                        showStep(step2, 2);
                    } else {
                        failText.textContent = result.data.message || "작품을 자동으로 찾지 못했어요";
                        showStep(step2fail, 2);
                    }
                })
                .catch(function () {
                    stopLoadingAnimation();
                    failText.textContent = "서버 연결에 실패했습니다.";
                    showStep(step2fail, 2);
                });
        });
    }

    // Button wiring — Step 1
    cameraBtn.addEventListener("click", function () { fileInput.click(); });
    fileInput.addEventListener("change", function () {
        if (fileInput.files.length > 0) uploadFile(fileInput.files[0]);
    });

    galleryBtn.addEventListener("click", function () { galleryInput.click(); });
    galleryInput.addEventListener("change", function () {
        if (galleryInput.files.length > 0) uploadFile(galleryInput.files[0]);
    });

    // Step 2 actions
    confirmBtn.addEventListener("click", function () {
        showStep(savedView);
    });

    retryBtn.addEventListener("click", function () {
        showStep(step1);
    });

    failRetryBtn.addEventListener("click", function () {
        showStep(step1);
    });

    adjustCornersBtn.addEventListener("click", function () {
        startCornerMode();
    });

    touchupBtn.addEventListener("click", function () {
        startBrushMode();
    });

    failManualBtn.addEventListener("click", function () {
        startCornerMode();
    });

    newPhotoBtn.addEventListener("click", function () {
        currentFileId       = null;
        currentOriginalUrl  = null;
        currentProcessedUrl = null;
        showStep(step1);
    });

    // ── Tab switching in Step 3 ───────────────────────────────────────────
    tabCorners.addEventListener("click", function () {
        tabCorners.classList.add("active");
        tabBrush.classList.remove("active");
        cornerMode.hidden = false;
        brushMode.hidden  = true;
    });

    tabBrush.addEventListener("click", function () {
        tabBrush.classList.add("active");
        tabCorners.classList.remove("active");
        brushMode.hidden  = false;
        cornerMode.hidden = true;
        // Lazy init brush canvas if not already loaded
        if (!brushImg) initBrushCanvas();
    });

    // ── Corner Mode ───────────────────────────────────────────────────────
    function startCornerMode() {
        cornerPoints   = [];
        cornerMousePos = null;
        cornerApplyBtn.disabled = true;
        tabCorners.classList.add("active");
        tabBrush.classList.remove("active");
        cornerMode.hidden = false;
        brushMode.hidden  = true;
        brushImg = null; // reset brush so it reloads
        showStep(step3, 3);
        loadCornerImage();
    }

    function loadCornerImage() {
        var url = currentProcessedUrl || currentOriginalUrl;
        cornerImg = new Image();
        cornerImg.onload = function () {
            var container = cornerCanvas.parentElement;
            var maxW = container.clientWidth  || (window.innerWidth);
            var maxH = container.clientHeight || Math.round(window.innerHeight * 0.55);
            cornerScale = Math.min(maxW / cornerImg.width, maxH / cornerImg.height, 1);
            cornerCanvas.width  = Math.round(cornerImg.width  * cornerScale);
            cornerCanvas.height = Math.round(cornerImg.height * cornerScale);
            drawCornerCanvas();
        };
        cornerImg.src = url + "?t=" + Date.now();
    }

    function updateCornerHint() {
        var left = 4 - cornerPoints.length;
        if (left > 0) {
            cornerHint.textContent = "꼭지점을 터치하세요 (남은 점: " + left + "개)";
        } else {
            cornerHint.textContent = "4개 완료! 적용 버튼을 누르세요";
        }
    }

    function drawCornerCanvas() {
        if (!cornerImg) return;
        cornerCtx.drawImage(cornerImg, 0, 0, cornerCanvas.width, cornerCanvas.height);

        var pts = cornerPoints;

        // Solid lines between placed points
        if (pts.length > 1) {
            cornerCtx.beginPath();
            cornerCtx.setLineDash([]);
            cornerCtx.moveTo(pts[0].cx, pts[0].cy);
            for (var i = 1; i < pts.length; i++) cornerCtx.lineTo(pts[i].cx, pts[i].cy);
            if (pts.length === 4) cornerCtx.closePath();
            cornerCtx.strokeStyle = "rgba(37,99,235,0.9)";
            cornerCtx.lineWidth = 2.5;
            cornerCtx.stroke();
        }

        // Dashed preview to mouse/touch position
        if (cornerMousePos && pts.length > 0 && pts.length < 4) {
            var last = pts[pts.length - 1];
            cornerCtx.beginPath();
            cornerCtx.setLineDash([6, 4]);
            cornerCtx.strokeStyle = "rgba(220,38,38,0.8)";
            cornerCtx.lineWidth = 2;
            cornerCtx.moveTo(last.cx, last.cy);
            cornerCtx.lineTo(cornerMousePos.cx, cornerMousePos.cy);
            if (pts.length >= 2) {
                cornerCtx.moveTo(cornerMousePos.cx, cornerMousePos.cy);
                cornerCtx.lineTo(pts[0].cx, pts[0].cy);
            }
            cornerCtx.stroke();
            cornerCtx.setLineDash([]);
        }

        // Semi-transparent fill when 4 points
        if (pts.length === 4) {
            cornerCtx.beginPath();
            cornerCtx.moveTo(pts[0].cx, pts[0].cy);
            for (var j = 1; j < pts.length; j++) cornerCtx.lineTo(pts[j].cx, pts[j].cy);
            cornerCtx.closePath();
            cornerCtx.fillStyle = "rgba(220,38,38,0.18)";
            cornerCtx.fill();
        }

        // Point circles
        for (var k = 0; k < pts.length; k++) {
            var p = pts[k];
            cornerCtx.beginPath();
            cornerCtx.arc(p.cx, p.cy, 12, 0, Math.PI * 2);
            cornerCtx.fillStyle = "rgba(37,99,235,0.85)";
            cornerCtx.fill();
            cornerCtx.strokeStyle = "white";
            cornerCtx.lineWidth = 2.5;
            cornerCtx.setLineDash([]);
            cornerCtx.stroke();
            cornerCtx.fillStyle = "white";
            cornerCtx.font = "bold 13px sans-serif";
            cornerCtx.textAlign = "center";
            cornerCtx.textBaseline = "middle";
            cornerCtx.fillText(String(k + 1), p.cx, p.cy);
        }
    }

    function getCanvasXY(canvas, e) {
        var rect = canvas.getBoundingClientRect();
        var scaleX = canvas.width  / rect.width;
        var scaleY = canvas.height / rect.height;
        var clientX, clientY;
        if (e.touches && e.touches.length > 0) {
            clientX = e.touches[0].clientX;
            clientY = e.touches[0].clientY;
        } else {
            clientX = e.clientX;
            clientY = e.clientY;
        }
        return {
            cx: (clientX - rect.left) * scaleX,
            cy: (clientY - rect.top)  * scaleY
        };
    }

    function cornerHandleTap(e) {
        e.preventDefault();
        if (cornerPoints.length >= 4) return;
        var pos = getCanvasXY(cornerCanvas, e);
        cornerPoints.push({
            cx: pos.cx, cy: pos.cy,
            ox: pos.cx / cornerScale,
            oy: pos.cy / cornerScale
        });
        updateCornerHint();
        drawCornerCanvas();
        if (cornerPoints.length === 4) cornerApplyBtn.disabled = false;
    }

    cornerCanvas.addEventListener("click",      cornerHandleTap);
    cornerCanvas.addEventListener("touchend",   cornerHandleTap);

    cornerCanvas.addEventListener("mousemove", function (e) {
        if (cornerPoints.length === 0 || cornerPoints.length >= 4) return;
        cornerMousePos = getCanvasXY(cornerCanvas, e);
        drawCornerCanvas();
    });

    cornerCanvas.addEventListener("mouseleave", function () {
        cornerMousePos = null;
        drawCornerCanvas();
    });

    cornerCanvas.addEventListener("touchmove", function (e) { e.preventDefault(); });

    cornerResetBtn.addEventListener("click", function () {
        cornerPoints = [];
        cornerApplyBtn.disabled = true;
        updateCornerHint();
        drawCornerCanvas();
    });

    cornerCancelBtn.addEventListener("click", function () {
        if (currentProcessedUrl) {
            showStep(step2, 2);
        } else {
            showStep(step2fail, 2);
        }
    });

    cornerApplyBtn.addEventListener("click", function () {
        if (cornerPoints.length !== 4 || !currentFileId) return;
        showStep(loadingView, 3);
        startLoadingAnimation();

        var points = cornerPoints.map(function (p) { return [p.ox, p.oy]; });

        fetch("/manual-crop", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ file_id: currentFileId, points: points, source: "original" })
        })
        .then(function (resp) {
            return resp.json().then(function (data) { return { ok: resp.ok, data: data }; });
        })
        .then(function (result) {
            stopLoadingAnimation();
            if (!result.ok) {
                failText.textContent = result.data.error || "꼭지점 보정 실패";
                showStep(step2fail, 2);
                return;
            }
            currentProcessedUrl = result.data.processed;
            reviewImg.src = result.data.processed + "?t=" + Date.now();
            reviewMessage.textContent = result.data.message || "꼭지점 보정 완료!";
            handBadge.hidden = true;
            confidenceBadge.hidden = true;
            showStep(step2, 2);
        })
        .catch(function () {
            stopLoadingAnimation();
            failText.textContent = "서버 연결에 실패했습니다.";
            showStep(step2fail, 2);
        });
    });

    // ── Brush Mode ────────────────────────────────────────────────────────
    function startBrushMode() {
        tabBrush.classList.add("active");
        tabCorners.classList.remove("active");
        brushMode.hidden  = false;
        cornerMode.hidden = true;
        brushHistory    = [];
        brushUndoBtn.disabled = true;
        showStep(step3, 3);
        initBrushCanvas();
    }

    function initBrushCanvas() {
        var url = currentProcessedUrl || currentOriginalUrl;
        brushImg = new Image();
        brushImg.onload = function () {
            var container = brushCanvas.parentElement;
            var maxW = container.clientWidth  || window.innerWidth;
            var maxH = container.clientHeight || Math.round(window.innerHeight * 0.55);
            brushScale = Math.min(maxW / brushImg.width, maxH / brushImg.height, 1);
            brushCanvas.width  = Math.round(brushImg.width  * brushScale);
            brushCanvas.height = Math.round(brushImg.height * brushScale);

            brushMaskCanvas = document.createElement("canvas");
            brushMaskCanvas.width  = brushCanvas.width;
            brushMaskCanvas.height = brushCanvas.height;
            brushMaskCtx = brushMaskCanvas.getContext("2d");
            brushMaskCtx.fillStyle = "black";
            brushMaskCtx.fillRect(0, 0, brushMaskCanvas.width, brushMaskCanvas.height);

            drawBrushCanvas();
        };
        brushImg.src = url + "?t=" + Date.now();
    }

    function drawBrushCanvas() {
        if (!brushImg) return;
        brushCtx.drawImage(brushImg, 0, 0, brushCanvas.width, brushCanvas.height);

        // Overlay red tint on masked areas
        var maskData = brushMaskCtx.getImageData(0, 0, brushMaskCanvas.width, brushMaskCanvas.height);
        var imgData  = brushCtx.getImageData(0, 0, brushCanvas.width, brushCanvas.height);
        for (var i = 0; i < maskData.data.length; i += 4) {
            if (maskData.data[i] > 128) {
                imgData.data[i]     = Math.min(255, imgData.data[i]     + 100);
                imgData.data[i + 1] = Math.max(0,   imgData.data[i + 1] - 50);
                imgData.data[i + 2] = Math.max(0,   imgData.data[i + 2] - 50);
            }
        }
        brushCtx.putImageData(imgData, 0, 0);
    }

    function paintBrushMask(x, y) {
        var r = parseInt(brushSize.value) * brushScale;
        brushMaskCtx.beginPath();
        brushMaskCtx.arc(x, y, r, 0, Math.PI * 2);
        brushMaskCtx.fillStyle = "white";
        brushMaskCtx.fill();
    }

    function drawBrushCursor(x, y) {
        var r = parseInt(brushSize.value) * brushScale;
        brushCtx.beginPath();
        brushCtx.arc(x, y, r, 0, Math.PI * 2);
        brushCtx.strokeStyle = "rgba(255,255,255,0.85)";
        brushCtx.lineWidth = 2;
        brushCtx.stroke();
    }

    function sendBrushInpaint() {
        // Show overlay spinner
        var overlay = document.createElement("div");
        overlay.className = "inpaint-overlay";
        overlay.innerHTML = '<div class="inpaint-overlay-text"><div class="loading-spinner"></div>AI 보정 중...</div>';
        brushCanvas.parentElement.appendChild(overlay);

        // Save undo snapshot (the current image src)
        brushHistory.push(brushImg.src);
        brushUndoBtn.disabled = false;

        // Full-res image canvas
        var imgCanvas = document.createElement("canvas");
        imgCanvas.width  = brushImg.naturalWidth;
        imgCanvas.height = brushImg.naturalHeight;
        imgCanvas.getContext("2d").drawImage(brushImg, 0, 0);
        var imageB64 = imgCanvas.toDataURL("image/jpeg", 0.92);

        // Full-res mask canvas
        var maskFull = document.createElement("canvas");
        maskFull.width  = brushImg.naturalWidth;
        maskFull.height = brushImg.naturalHeight;
        maskFull.getContext("2d").drawImage(brushMaskCanvas, 0, 0, maskFull.width, maskFull.height);
        var maskB64 = maskFull.toDataURL("image/png");

        fetch("/inpaint", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ image: imageB64, mask: maskB64, file_id: currentFileId })
        })
        .then(function (resp) {
            return resp.json().then(function (data) { return { ok: resp.ok, data: data }; });
        })
        .then(function (result) {
            overlay.remove();
            if (!result.ok) {
                alert(result.data.error || "보정 실패");
                return;
            }
            currentProcessedUrl = result.data.processed;
            var newImg = new Image();
            newImg.onload = function () {
                brushImg = newImg;
                // Reset mask
                brushMaskCtx.fillStyle = "black";
                brushMaskCtx.fillRect(0, 0, brushMaskCanvas.width, brushMaskCanvas.height);
                drawBrushCanvas();
            };
            newImg.src = result.data.result_image;
        })
        .catch(function () {
            overlay.remove();
            alert("서버 연결에 실패했습니다.");
        });
    }

    // Brush canvas events
    brushCanvas.addEventListener("mousedown", function (e) {
        e.preventDefault();
        brushDrawing = true;
        var pos = getCanvasXY(brushCanvas, e);
        paintBrushMask(pos.cx, pos.cy);
        drawBrushCanvas();
        drawBrushCursor(pos.cx, pos.cy);
    });

    brushCanvas.addEventListener("mousemove", function (e) {
        e.preventDefault();
        var pos = getCanvasXY(brushCanvas, e);
        if (brushDrawing) paintBrushMask(pos.cx, pos.cy);
        drawBrushCanvas();
        drawBrushCursor(pos.cx, pos.cy);
    });

    brushCanvas.addEventListener("mouseup", function (e) {
        e.preventDefault();
        if (!brushDrawing) return;
        brushDrawing = false;
        if (hasMaskContent()) sendBrushInpaint();
    });

    brushCanvas.addEventListener("mouseleave", function (e) {
        if (brushDrawing) {
            brushDrawing = false;
            if (hasMaskContent()) sendBrushInpaint();
        }
        drawBrushCanvas();
    });

    brushCanvas.addEventListener("touchstart", function (e) {
        e.preventDefault();
        brushDrawing = true;
        var pos = getCanvasXY(brushCanvas, e);
        paintBrushMask(pos.cx, pos.cy);
        drawBrushCanvas();
    }, { passive: false });

    brushCanvas.addEventListener("touchmove", function (e) {
        e.preventDefault();
        if (!brushDrawing) return;
        var pos = getCanvasXY(brushCanvas, e);
        paintBrushMask(pos.cx, pos.cy);
        drawBrushCanvas();
    }, { passive: false });

    brushCanvas.addEventListener("touchend", function (e) {
        e.preventDefault();
        if (!brushDrawing) return;
        brushDrawing = false;
        if (hasMaskContent()) sendBrushInpaint();
    }, { passive: false });

    function hasMaskContent() {
        if (!brushMaskCtx) return false;
        var data = brushMaskCtx.getImageData(0, 0, brushMaskCanvas.width, brushMaskCanvas.height).data;
        for (var i = 0; i < data.length; i += 4) {
            if (data[i] > 128) return true;
        }
        return false;
    }

    brushSize.addEventListener("input", function () {
        brushSizeLabel.textContent = brushSize.value + "px";
    });

    brushUndoBtn.addEventListener("click", function () {
        if (brushHistory.length === 0) return;
        var prevSrc = brushHistory.pop();
        if (brushHistory.length === 0) brushUndoBtn.disabled = true;
        var prev = new Image();
        prev.onload = function () {
            brushImg = prev;
            brushMaskCtx.fillStyle = "black";
            brushMaskCtx.fillRect(0, 0, brushMaskCanvas.width, brushMaskCanvas.height);
            drawBrushCanvas();
        };
        prev.src = prevSrc;
    });

    brushDoneBtn.addEventListener("click", function () {
        reviewImg.src = (currentProcessedUrl || "") + "?t=" + Date.now();
        reviewMessage.textContent = "보정 완료!";
        handBadge.hidden = true;
        confidenceBadge.hidden = true;
        showStep(step2, 2);
    });

    brushCancelBtn.addEventListener("click", function () {
        if (currentProcessedUrl) showStep(step2, 2);
        else showStep(step2fail, 2);
    });

    // ── Initial state ─────────────────────────────────────────────────────
    showStep(step1);

}); // end DOMContentLoaded
