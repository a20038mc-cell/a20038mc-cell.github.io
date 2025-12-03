document.addEventListener("DOMContentLoaded", () => {
    // ★重要: あなたのGASウェブアプリのURLに書き換えてください
    const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyvXnq6KYcJAkWdRD4w4rvywkqywAHgmfEAEr9bXurLH057XbQBHkB-zTKxkqyt2/exec';

    // OpenCVのロード完了を待つフラグ
    let isCvLoaded = false;
    window.Module = {
        onRuntimeInitialized: () => {
            isCvLoaded = true;
            console.log("OpenCV ready.");
            DOM.status.innerText = "書類を選択してください";
        }
    };

    // UI要素の取得
    const DOM = {
        modeSelection: document.getElementById('mode-selection'),
        appContainer: document.getElementById('app-container'),
        
        video: document.getElementById('video'),
        image: document.getElementById('uploaded-image'),
        canvas: document.getElementById('selection-canvas'),
        ctx: document.getElementById('selection-canvas').getContext('2d'),
        previewContainer: document.getElementById('preview-container'),

        btnHyoushi: document.getElementById('select-hyoushi'),
        btnShishutsu: document.getElementById('select-shishutsu'),
        btnCamera: document.getElementById('btn-camera-mode'),
        btnFile: document.getElementById('btn-file-mode'),
        fileInput: document.getElementById('file-input'),
        btnSave: document.getElementById('save-button'),
        btnBack: document.getElementById('back-button'),
        
        status: document.getElementById('status-label'),
        targetArea: document.getElementById('dynamic-target-buttons'),
        resultArea: document.getElementById('dynamic-result-labels'),
        
        // ★追加: カメラ選択UI要素
        cameraSelectGroup: document.getElementById('camera-select-group'),
        cameraSelect: document.getElementById('camera-select'),
    };

    // アプリの状態
    const State = {
        sheetName: null,      
        definitions: [],      
        currentTarget: null,  
        ocrResults: {},       
        stream: null,         
        isCameraMode: true,   
        isProcessing: false,  
        rafId: null,          
        isOCRReady: false,    
    };

    // 読み取り項目の定義 (前と同じ)
    const DEFINITIONS = {
        "OCR-Data": [
            { key: "dantai_name", label: "団体名称" },
            { key: "daihyousha", label: "代表者氏名" },
            { key: "kaikei_sekinin", label: "会計責任者" },
            { key: "jimutantou", label: "事務担当者" },
            { key: "date_koushutsu", label: "公出年月日" }
        ],
        "支払明細": [
            { key: "no", label: "番号(No)" },
            { key: "kingaku", label: "金額" },
            { key: "shishutsu_date", label: "支出年月日" },
            { key: "shishutsu_mokuteki", label: "支出の目的" },
            { key: "shishutsu_saki", label: "支出先名称" }
        ]
    };

    // --- 1. OCRエンジンの管理 (前と同じ) ---
    const OCR = {
        worker: null,
        init: async () => {
            if (State.isOCRReady) return;
            DOM.status.innerText = "OCRエンジン起動中...";
            
            try {
                OCR.worker = await Tesseract.createWorker('jpn', 1, {
                    logger: m => {
                        if (m.status === 'recognizing text') {
                            DOM.status.innerText = `読取中... ${(m.progress * 100).toFixed(0)}%`;
                        }
                    },
                });
                State.isOCRReady = true;
                DOM.status.innerText = "OCR準備完了";
            } catch (error) {
                 DOM.status.innerText = "OCRエンジンの起動に失敗しました。";
                 console.error("Tesseract initialization error:", error);
            }
        },
        recognize: async (canvas, options) => {
            if (!State.isOCRReady) await OCR.init();
            
            await OCR.worker.setParameters({
                tessedit_pageseg_mode: Tesseract.PSM.SINGLE_LINE,
                ...options
            });
            
            const { data: { text } } = await OCR.worker.recognize(canvas);
            return text;
        }
    };

    // --- 2. 画像処理と読取実行 (前と同じ) ---
    const Processor = {
        execute: async () => {
            if (State.isProcessing || !State.currentTarget || !isCvLoaded || !State.isOCRReady) return;
            
            const source = State.isCameraMode ? DOM.video : DOM.image;
            
            if (State.isCameraMode && DOM.video.readyState !== 4) {
                 DOM.status.innerText = "カメラの映像を待機中...";
                 return;
            }
            if (!State.isCameraMode && (!DOM.image.src || DOM.image.style.display === 'none')) return;

            State.isProcessing = true;
            let srcMat = null, grayMat = null, binMat = null, roiMat = null;

            try {
                const w = State.isCameraMode ? DOM.video.videoWidth : DOM.image.naturalWidth;
                const h = State.isCameraMode ? DOM.video.videoHeight : DOM.image.naturalHeight;
                
                if (!w || !h || w === 0 || h === 0) throw new Error("ソースサイズ取得失敗");

                const capCanvas = document.createElement('canvas');
                capCanvas.width = w;
                capCanvas.height = h;
                const capCtx = capCanvas.getContext('2d');
                capCtx.drawImage(source, 0, 0, w, h);

                srcMat = cv.imread(capCanvas);

                const displayW = DOM.canvas.offsetWidth;
                const displayH = DOM.canvas.offsetHeight;
                
                const scaleX = w / displayW;
                const scaleY = h / displayH;

                const rectW_disp = displayW * 0.6;
                const rectH_disp = 100;
                const rectX_disp = (displayW - rectW_disp) / 2;
                const rectY_disp = (displayH - rectH_disp) / 2;
                
                const rectX = Math.floor(rectX_disp * scaleX);
                const rectY = Math.floor(rectY_disp * scaleY);
                const rectWidth = Math.floor(rectW_disp * scaleX);
                const rectHeight = Math.floor(rectH_disp * scaleY);

                if (rectX < 0 || rectY < 0) throw new Error("計算エラー");

                let roiRect = new cv.Rect(rectX, rectY, rectWidth, rectHeight);
                roiMat = srcMat.roi(roiRect);

                grayMat = new cv.Mat();
                cv.cvtColor(roiMat, grayMat, cv.COLOR_RGBA2GRAY);
                
                binMat = new cv.Mat();
                cv.medianBlur(grayMat, binMat, 3);
                grayMat.delete();

                cv.adaptiveThreshold(binMat, binMat, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 15, 8);

                let dsize = new cv.Size(binMat.cols * 2, binMat.rows * 2);
                cv.resize(binMat, binMat, dsize, 0, 0, cv.INTER_LINEAR);

                const finalCanvas = document.createElement('canvas');
                cv.imshow(finalCanvas, binMat);

                const definition = State.definitions.find(d => d.key === State.currentTarget);
                const label = definition.label;
                let opts = { tessedit_char_whitelist: '' };

                if (label.includes("金額")) {
                    opts.tessedit_char_whitelist = '0123456789,¥円';
                } else if (label.includes("日付") || label.includes("年月日")) {
                    opts.tessedit_char_whitelist = '0123456789/.-年月日';
                } else if (label.includes("No")) {
                    opts.tessedit_char_whitelist = '0123456789';
                }

                const rawText = await OCR.recognize(finalCanvas, opts);
                const cleanText = rawText.replace(/\s+/g, '').trim();

                if (cleanText.length > 0) {
                    UIManager.setResult(State.currentTarget, cleanText);
                    DOM.status.innerText = `読取成功: ${cleanText}`;
                    
                    await new Promise(r => setTimeout(r, State.isCameraMode ? 1500 : 500));
                } else {
                    if (!State.isCameraMode) DOM.status.innerText = "文字が見つかりません (枠に合わせてください)";
                }

            } catch (err) {
                console.error("Processor Execution Error:", err);
                DOM.status.innerText = "処理エラー: コンソールを確認してください";
            } finally {
                if (srcMat) srcMat.delete();
                if (roiMat) roiMat.delete();
                if (binMat) binMat.delete();
                State.isProcessing = false;
            }
        }
    };

    // --- 3. UIの制御 (前と同じ) ---
    const UIManager = {
        init: (sheetName) => {
            State.sheetName = sheetName;
            State.definitions = DEFINITIONS[sheetName];
            State.ocrResults = {};
            State.currentTarget = null;

            DOM.modeSelection.style.display = 'none';
            DOM.appContainer.style.display = 'flex';

            DOM.targetArea.innerHTML = '';
            DOM.resultArea.innerHTML = '';
            
            State.definitions.forEach(def => {
                const btn = document.createElement('button');
                btn.textContent = def.label;
                btn.onclick = () => UIManager.selectTarget(def.key);
                DOM.targetArea.appendChild(btn);

                const div = document.createElement('div');
                div.className = 'result-item';
                div.innerHTML = `
                    <span>${def.label}:</span>
                    <input type="text" id="res-${def.key}" placeholder="未入力">
                `;
                DOM.resultArea.appendChild(div);
            });

            OCR.init();

            Actions.switchMode(true);
        },

        selectTarget: (key) => {
            if (State.currentTarget === key && State.isCameraMode) return;

            State.currentTarget = key;
            const label = State.definitions.find(d => d.key === key).label;
            
            Array.from(DOM.targetArea.children).forEach(btn => {
                btn.classList.remove('active-target');
                btn.style.backgroundColor = ''; 
                if (btn.textContent === label) {
                    btn.classList.add('active-target');
                    btn.style.backgroundColor = '#d1e7dd';
                } 
            });

            DOM.status.innerText = `「${label}」をスキャン中...`;

            if (!State.isCameraMode) {
                Processor.execute();
            }
        },

        setResult: (key, val) => {
            State.ocrResults[key] = val;
            const input = document.getElementById(`res-${key}`);
            if (input) input.value = val;
        }
    };

    // --- 4. アクション (カメラ/ファイル/保存) ---
    const Actions = {
        
        // ★新規追加: カメラデバイスを列挙し、ドロップダウンを構築する関数
        enumerateCameras: async (selectedDeviceId) => {
            try {
                const devices = await navigator.mediaDevices.enumerateDevices();
                const videoDevices = devices.filter(d => d.kind === 'videoinput');
                
                DOM.cameraSelect.innerHTML = '';
                
                // カメラが2つ以上ある場合にのみ選択UIを表示
                if (videoDevices.length > 1) {
                    DOM.cameraSelectGroup.style.display = 'block';
                } else {
                    DOM.cameraSelectGroup.style.display = 'none';
                }

                videoDevices.forEach((device, index) => {
                    const option = document.createElement('option');
                    option.value = device.deviceId;
                    option.text = device.label || `Camera ${index + 1}`; 
                    DOM.cameraSelect.appendChild(option);
                });
                
                // 起動したカメラを選択状態にする
                if (selectedDeviceId) DOM.cameraSelect.value = selectedDeviceId;

            } catch (err) {
                console.error("enumerateDevices Error:", err);
            }
        },

        // ★修正: カメラ選択UIの表示切り替えと startCamera 呼び出しの変更
        switchMode: (isCamera) => {
            State.isCameraMode = isCamera;
            
            DOM.btnCamera.classList.toggle('active', isCamera);
            DOM.btnFile.classList.toggle('active', !isCamera);
            
            if (isCamera) {
                // カメラモード
                DOM.image.style.display = 'none';
                DOM.video.style.display = 'block';
                // 選択肢がある場合は選択UIを表示
                DOM.cameraSelectGroup.style.display = DOM.cameraSelect.options.length > 1 ? 'block' : 'none';
                Actions.startCamera();
                if (!State.rafId) Actions.drawGuideLoop();
            } else {
                // ファイルモード
                Actions.stopCamera();
                DOM.video.style.display = 'none';
                DOM.image.style.display = DOM.image.src ? 'block' : 'none';
                DOM.cameraSelectGroup.style.display = 'none';
                DOM.status.innerText = DOM.image.src ? "画像読込完了。項目ボタンを押してOCR実行" : "画像をアップロードしてください";
            }
        },

        // ★修正: 選択されたカメラIDを使ってストリームを開始するロジック
        startCamera: async (deviceId) => {
            Actions.stopCamera(); 
            
            DOM.status.innerText = "カメラ起動中...";

            // 選択されたデバイスID、またはドロップダウンに値があればそのIDを使用
            const targetDeviceId = deviceId || (DOM.cameraSelect.options.length > 0 ? DOM.cameraSelect.value : null);

            const constraints = {
                video: {
                    // 特定のカメラIDが指定されている場合は、そのIDをexactで要求
                    deviceId: targetDeviceId ? { exact: targetDeviceId } : undefined,
                    // IDがない場合は、背面カメラを優先
                    facingMode: targetDeviceId ? undefined : "environment" 
                }
            };

            try {
                const stream = await navigator.mediaDevices.getUserMedia(constraints);
                
                DOM.video.srcObject = stream;
                DOM.video.play(); 
                State.stream = stream;
                DOM.status.innerText = "項目を選択してカメラを向けてください";

                // カメラ起動成功後、初めて利用可能デバイスを列挙し、ドロップダウンを更新する
                if (DOM.cameraSelect.options.length === 0) {
                     // 起動したストリームの最初のトラックからデバイスIDを取得
                     const activeTrack = stream.getVideoTracks()[0];
                     const activeDeviceId = activeTrack ? activeTrack.getSettings().deviceId : null;
                     
                     await Actions.enumerateCameras(activeDeviceId);
                }
                
            } catch (err) {
                console.error("Camera Access Error:", err);
                alert(`カメラを起動できませんでした。\n原因: ${err.name} - ${err.message}\n(ヒント: カメラへのアクセス権限を確認してください。設定が厳しすぎる可能性があります。)`);
                Actions.switchMode(false);
            }
        },

        stopCamera: () => {
            if (State.stream) {
                State.stream.getTracks().forEach(t => t.stop());
                State.stream = null;
            }
        },

        saveData: () => {
            if (!State.sheetName || Object.keys(State.ocrResults).length === 0) {
                alert("読み取り項目がありません。");
                return;
            }

            DOM.status.innerText = "データをGoogle Sheetsへ送信中...";
            
            State.definitions.forEach(def => {
                const input = document.getElementById(`res-${def.key}`);
                if (input && input.value) {
                    State.ocrResults[def.key] = input.value;
                }
            });

            fetch(GOOGLE_SCRIPT_URL, {
                method: 'POST',
                mode: 'no-cors', 
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    action: 'appendData',
                    sheetName: State.sheetName,
                    data: State.ocrResults
                })
            }).then(() => {
                DOM.status.innerText = "送信完了 (Sheetsを確認してください)";
                alert("送信完了しました");
            }).catch((e) => {
                 console.error("Fetch Error:", e);
                 DOM.status.innerText = "送信エラー (ネットワークまたはGAS設定の問題)";
                 alert("送信エラー");
            });
        },

        drawGuideLoop: () => {
            const target = State.isCameraMode ? DOM.video : DOM.image;
            
            const w = target.offsetWidth;
            const h = target.offsetHeight;

            if (w > 0 && h > 0 && target.style.display !== 'none') {
                DOM.canvas.width = w;
                DOM.canvas.height = h;
                
                DOM.canvas.style.width = `${w}px`;
                DOM.canvas.style.height = `${h}px`;

                const ctx = DOM.ctx;
                ctx.clearRect(0, 0, w, h);
                
                const rectW = w * 0.6;
                const rectH = 100;
                const x = (w - rectW) / 2;
                const y = (h - rectH) / 2;

                ctx.strokeStyle = "red";
                ctx.lineWidth = 3;
                ctx.strokeRect(x, y, rectW, rectH);
            }
            
            State.rafId = requestAnimationFrame(Actions.drawGuideLoop);
        }
    };

    // --- イベントリスナー設定 ---
    DOM.btnHyoushi.addEventListener('click', () => UIManager.init('OCR-Data'));
    DOM.btnShishutsu.addEventListener('click', () => UIManager.init('支払明細'));
    
    DOM.btnCamera.addEventListener('click', () => Actions.switchMode(true));
    DOM.btnFile.addEventListener('click', () => {
        Actions.switchMode(false);
        DOM.fileInput.click();
    }); 

    // ★新規追加: カメラ選択が変更されたときのイベントリスナー
    DOM.cameraSelect.addEventListener('change', () => {
        // 選択されたカメラIDを取得し、新しいストリームでカメラを再起動
        Actions.startCamera(DOM.cameraSelect.value); 
    });

    DOM.fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = (ev) => {
            DOM.image.src = ev.target.result;
            DOM.image.onload = () => {
                 DOM.image.style.display = 'block';
                 DOM.status.innerText = "画像読込完了。項目ボタンを押してOCR実行";
                 if (!State.rafId) Actions.drawGuideLoop();
            };
        };
        reader.readAsDataURL(file);
    });

    DOM.btnSave.addEventListener('click', Actions.saveData);
    
    DOM.btnBack.addEventListener('click', () => {
        Actions.stopCamera();
        DOM.appContainer.style.display = 'none';
        DOM.modeSelection.style.display = 'flex';
        cancelAnimationFrame(State.rafId);
        State.rafId = null;
        State.currentTarget = null;
    });

    // 監視ループ (カメラモード時のみ定期的にOCR試行)
    setInterval(() => {
        if (State.isCameraMode && State.currentTarget && !State.isProcessing && isCvLoaded && State.isOCRReady) {
            Processor.execute();
        }
    }, 1500);

    if (typeof cv === 'undefined') {
        DOM.status.innerText = "OpenCV.jsをロード中...";
    } else {
        DOM.status.innerText = "書類を選択してください";
    }
});
