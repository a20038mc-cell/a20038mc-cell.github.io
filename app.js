document.addEventListener("DOMContentLoaded", () => {
    // ★重要: あなたのGASウェブアプリのURLに書き換えてください
    const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyvXnq6KYcJAkWdRD4w4rvywkqywAHgmfEAEq8Ir9bXurLH057XbQBHkB-zTKxkqyt2/exec';

    // UI要素の取得
    const DOM = {
        modeSelection: document.getElementById('mode-selection'),
        appContainer: document.getElementById('app-container'),
        
        // プレビューエリア
        video: document.getElementById('video'),
        image: document.getElementById('uploaded-image'),
        canvas: document.getElementById('selection-canvas'),
        ctx: document.getElementById('selection-canvas').getContext('2d'),
        previewContainer: document.getElementById('preview-container'),

        // ボタン
        btnHyoushi: document.getElementById('select-hyoushi'),
        btnShishutsu: document.getElementById('select-shishutsu'),
        btnCamera: document.getElementById('btn-camera-mode'),
        btnFile: document.getElementById('btn-file-mode'),
        fileInput: document.getElementById('file-input'),
        btnSave: document.getElementById('save-button'),
        btnBack: document.getElementById('back-button'),
        
        // 表示エリア
        status: document.getElementById('status-label'),
        targetArea: document.getElementById('dynamic-target-buttons'),
        resultArea: document.getElementById('dynamic-result-labels')
    };

    // アプリの状態
    const State = {
        sheetName: null,      // 'OCR-Data' or '支払明細'
        definitions: [],      // 現在の項目の定義
        currentTarget: null,  // 今読み取っている項目のキー
        ocrResults: {},       // 読取結果
        stream: null,         // カメラストリーム
        isCameraMode: true,   // true=カメラ, false=ファイル
        isProcessing: false,  // 処理中フラグ
        rafId: null           // アニメーションフレームID
    };

    // 読み取り項目の定義
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
            { key: "kingaku", label: "金額" }, // 数字のみ
            { key: "shishutsu_date", label: "支出年月日" }, // 日付のみ
            { key: "shishutsu_mokuteki", label: "支出の目的" },
            { key: "shishutsu_saki", label: "支出先名称" }
        ]
    };

    // --- 1. OCRエンジンの管理 ---
    const OCR = {
        worker: null,
        init: async () => {
            if (OCR.worker) return;
            DOM.status.innerText = "OCRエンジン起動中...";
            
            // ★重要: workerPathを同じ階層に指定
            OCR.worker = await Tesseract.createWorker('jpn', 1, {
                logger: m => {
                    if (m.status === 'recognizing text') {
                        DOM.status.innerText = `読取中... ${(m.progress * 100).toFixed(0)}%`;
                    }
                },
                workerPath: 'worker.min.js' 
                // corePathは自動解決に任せるか、必要なら 'tesseract-core.wasm.js' を指定
            });
            DOM.status.innerText = "OCR準備完了";
        },
        recognize: async (canvas, options) => {
            if (!OCR.worker) await OCR.init();
            
            // パラメータをセット
            await OCR.worker.setParameters({
                tessedit_pageseg_mode: Tesseract.PSM.SINGLE_LINE, // 1行として読む
                ...options
            });
            
            const { data: { text } } = await OCR.worker.recognize(canvas);
            return text;
        }
    };

    // --- 2. 画像処理と読取実行 ---
    const Processor = {
        execute: async () => {
            if (State.isProcessing || !State.currentTarget) return;

            // ソースの特定
            const source = State.isCameraMode ? DOM.video : DOM.image;
            if (State.isCameraMode && DOM.video.readyState !== 4) return;
            if (!State.isCameraMode && (!DOM.image.src || DOM.image.style.display === 'none')) return;

            State.isProcessing = true;
            let srcMat = null, grayMat = null, binMat = null;

            try {
                // 現在の表示サイズを取得
                const w = source.offsetWidth || source.videoWidth;
                const h = source.offsetHeight || source.videoHeight;
                if (!w || !h) throw new Error("サイズ取得失敗");

                // OpenCV用にCanvasへ描画
                const capCanvas = document.createElement('canvas');
                capCanvas.width = w; 
                capCanvas.height = h;
                const capCtx = capCanvas.getContext('2d');
                capCtx.drawImage(source, 0, 0, w, h);

                srcMat = cv.imread(capCanvas);

                // 赤枠エリア (ROI) の計算: 中央 60% x 100px
                const rectW = Math.floor(w * 0.6);
                const rectH = 100;
                const rectX = Math.floor((w - rectW) / 2);
                const rectY = Math.floor((h - rectH) / 2);

                if (rectX < 0 || rectY < 0) throw new Error("画面サイズ不足");

                // 切り出し
                let roiRect = new cv.Rect(rectX, rectY, rectWidth, rectHeight);
                let roiMat = srcMat.roi(roiRect);

                // --- 画像処理パイプライン ---
                // 1. グレースケール
                grayMat = new cv.Mat();
                cv.cvtColor(roiMat, grayMat, cv.COLOR_RGBA2GRAY);
                roiMat.delete();

                // 2. ノイズ除去
                binMat = new cv.Mat();
                cv.medianBlur(grayMat, binMat, 3);

                // 3. 適応的二値化 (影に強い)
                cv.adaptiveThreshold(binMat, binMat, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 15, 8);

                // 4. 2倍拡大 (小さな文字対策)
                let dsize = new cv.Size(binMat.cols * 2, binMat.rows * 2);
                cv.resize(binMat, binMat, dsize, 0, 0, cv.INTER_LINEAR);

                // 確認用Canvas作成
                const finalCanvas = document.createElement('canvas');
                cv.imshow(finalCanvas, binMat);

                // --- 項目別パラメータ設定 ---
                const label = State.definitions.find(d => d.key === State.currentTarget).label;
                let opts = { tessedit_char_whitelist: '' }; // ホワイトリスト初期化

                if (label.includes("金額")) {
                    opts.tessedit_char_whitelist = '0123456789,¥';
                } else if (label.includes("日付") || label.includes("年月日")) {
                    opts.tessedit_char_whitelist = '0123456789/.-年月日';
                } else if (label.includes("No")) {
                    opts.tessedit_char_whitelist = '0123456789';
                }

                // OCR実行
                const rawText = await OCR.recognize(finalCanvas, opts);
                const cleanText = rawText.replace(/\s+/g, '').trim();

                if (cleanText.length > 0) {
                    UIManager.setResult(State.currentTarget, cleanText);
                    DOM.status.innerText = `読取成功: ${cleanText}`;
                    
                    // 連続読取防止 (カメラモードは少し長く待つ)
                    await new Promise(r => setTimeout(r, State.isCameraMode ? 1500 : 500));
                } else {
                    if (!State.isCameraMode) DOM.status.innerText = "文字が見つかりません (枠に合わせてください)";
                }

            } catch (err) {
                console.error(err);
                DOM.status.innerText = "処理エラー";
            } finally {
                // メモリ解放
                if (srcMat) srcMat.delete();
                if (grayMat) grayMat.delete();
                if (binMat) binMat.delete();
                State.isProcessing = false;
            }
        }
    };

    // --- 3. UIの制御 ---
    const UIManager = {
        init: (sheetName) => {
            State.sheetName = sheetName;
            State.definitions = DEFINITIONS[sheetName];
            State.ocrResults = {};
            State.currentTarget = null;

            // 画面切り替え
            DOM.modeSelection.style.display = 'none';
            DOM.appContainer.style.display = 'flex';

            // ボタン生成
            DOM.targetArea.innerHTML = '';
            DOM.resultArea.innerHTML = '';
            
            State.definitions.forEach(def => {
                // ターゲットボタン
                const btn = document.createElement('button');
                btn.textContent = def.label;
                btn.onclick = () => UIManager.selectTarget(def.key);
                DOM.targetArea.appendChild(btn);

                // 結果欄
                const div = document.createElement('div');
                div.className = 'result-item';
                div.innerHTML = `
                    <span>${def.label}:</span>
                    <input type="text" id="res-${def.key}" placeholder="未入力">
                `;
                DOM.resultArea.appendChild(div);
            });

            // デフォルトはカメラモード
            Actions.switchMode(true);
        },

        selectTarget: (key) => {
            State.currentTarget = key;
            const label = State.definitions.find(d => d.key === key).label;
            
            // ボタンのハイライト
            Array.from(DOM.targetArea.children).forEach(btn => {
                if (btn.textContent === label) {
                    btn.classList.add('active-target');
                    btn.style.backgroundColor = '#d1e7dd';
                } else {
                    btn.classList.remove('active-target');
                    btn.style.backgroundColor = '';
                }
            });

            DOM.status.innerText = `「${label}」をスキャン中...`;

            // ファイルモードなら、ボタンを押した瞬間に実行
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
        switchMode: (isCamera) => {
            State.isCameraMode = isCamera;
            
            if (isCamera) {
                // カメラモード
                DOM.image.style.display = 'none';
                DOM.video.style.display = 'block';
                Actions.startCamera();
            } else {
                // ファイルモード
                Actions.stopCamera();
                DOM.video.style.display = 'none';
                DOM.image.style.display = 'block';
                DOM.status.innerText = "画像をアップロードしてください";
            }
            // 赤枠描画開始
            if (!State.rafId) Actions.drawGuideLoop();
        },

        startCamera: async () => {
            DOM.status.innerText = "カメラ起動中...";
            try {
                const stream = await navigator.mediaDevices.getUserMedia({
                    video: { facingMode: "environment", width: { ideal: 1280 }, height: { ideal: 720 } }
                });
                DOM.video.srcObject = stream;
                State.stream = stream;
                DOM.status.innerText = "項目を選択してカメラを向けてください";
                
                // OpenCVロード待ち
                if (typeof cv === 'undefined' || !cv.Mat) {
                    const waitCv = setInterval(() => {
                        if (typeof cv !== 'undefined' && cv.Mat) {
                            clearInterval(waitCv);
                            DOM.status.innerText = "準備完了";
                        }
                    }, 500);
                }
            } catch (err) {
                console.error(err);
                alert("カメラを起動できませんでした。HTTPS接続か確認してください。");
                // 失敗したらファイルモードへ誘導
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
            if (!State.sheetName) return;
            DOM.status.innerText = "送信中...";
            
            fetch(GOOGLE_SCRIPT_URL, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    sheetName: State.sheetName,
                    data: State.ocrResults
                })
            }).then(() => {
                alert("送信完了しました");
                DOM.status.innerText = "送信完了";
            }).catch(() => {
                alert("送信エラー");
            });
        },

        // 赤枠を描画し続けるループ
        drawGuideLoop: () => {
            const target = State.isCameraMode ? DOM.video : DOM.image;
            
            // サイズ合わせ
            const w = target.offsetWidth || target.videoWidth;
            const h = target.offsetHeight || target.videoHeight;

            // 要素が見えていない場合はスキップ
            if (w && h && target.style.display !== 'none') {
                DOM.canvas.width = w;
                DOM.canvas.height = h;
                
                // コンテナの中央に配置 (CSSのtransformを考慮したサイズ設定)
                DOM.canvas.style.width = `${w}px`;
                DOM.canvas.style.height = `${h}px`;

                const ctx = DOM.ctx;
                ctx.clearRect(0, 0, w, h);
                
                // 赤枠の計算
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
    DOM.btnFile.addEventListener('click', () => DOM.fileInput.click()); // ボタン→input発火

    DOM.fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            Actions.switchMode(false); // ファイルモードへ強制切替
            DOM.image.src = ev.target.result;
            DOM.status.innerText = "画像読込完了。項目ボタンを押してOCR実行";
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
    });

    // 監視ループ (カメラモード時のみ定期的にOCR試行)
    setInterval(() => {
        if (State.isCameraMode && State.currentTarget && !State.isProcessing) {
            Processor.execute();
        }
    }, 1500);
});