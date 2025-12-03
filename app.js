document.addEventListener("DOMContentLoaded", () => {
    // ★重要: あなたのGASウェブアプリのURLに書き換えてください
    const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyvXnq6KYcJAkWdRD4w4rvywkqywAHgmfEAEr9bXurLH057XbQBHkB-zTKxkqyt2/exec';

    // OpenCVのロード完了を待つフラグ
    let isCvLoaded = false;
    window.Module = {
        onRuntimeInitialized: () => {
            isCvLoaded = true;
            console.log("OpenCV ready.");
            // OpenCVがロードされた後に、初期状態のステータスを設定
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
        rafId: null,          // アニメーションフレームID
        isOCRReady: false,    // Tesseract準備完了フラグ
    };

    // 読み取り項目の定義 (変更なし)
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
            if (State.isOCRReady) return;
            DOM.status.innerText = "OCRエンジン起動中...";
            
            try {
                // workerPathはTesseract.jsのバージョンや配置によって異なる場合があります
                OCR.worker = await Tesseract.createWorker('jpn', 1, {
                    logger: m => {
                        if (m.status === 'recognizing text') {
                            DOM.status.innerText = `読取中... ${(m.progress * 100).toFixed(0)}%`;
                        }
                    },
                    // workerPath: 'worker.min.js' // 必要であれば指定を解除
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
            
            // パラメータをセット
            await OCR.worker.setParameters({
                tessedit_pageseg_mode: Tesseract.PSM.SINGLE_LINE,
                ...options
            });
            
            const { data: { text } } = await OCR.worker.recognize(canvas);
            return text;
        }
    };

    // --- 2. 画像処理と読取実行 ---
    const Processor = {
        execute: async () => {
            // OpenCVとOCRの両方が必要
            if (State.isProcessing || !State.currentTarget || !isCvLoaded || !State.isOCRReady) return;
            
            const source = State.isCameraMode ? DOM.video : DOM.image;
            
            // ソースの有効性チェック
            if (State.isCameraMode && DOM.video.readyState !== 4) {
                 DOM.status.innerText = "カメラの映像を待機中...";
                 return;
            }
            if (!State.isCameraMode && (!DOM.image.src || DOM.image.style.display === 'none')) return;

            State.isProcessing = true;
            let srcMat = null, grayMat = null, binMat = null, roiMat = null;

            try {
                // 現在の表示サイズを取得 (videoWidth/videoHeightはカメラモードでの内部解像度)
                const w = State.isCameraMode ? DOM.video.videoWidth : DOM.image.naturalWidth;
                const h = State.isCameraMode ? DOM.video.videoHeight : DOM.image.naturalHeight;
                
                if (!w || !h || w === 0 || h === 0) throw new Error("ソースサイズ取得失敗");

                // OpenCV用にCanvasへ描画 (高解像度でキャプチャ)
                const capCanvas = document.createElement('canvas');
                capCanvas.width = w;
                capCanvas.height = h;
                const capCtx = capCanvas.getContext('2d');
                capCtx.drawImage(source, 0, 0, w, h);

                srcMat = cv.imread(capCanvas);

                // 赤枠エリア (ROI) の計算:
                // DOM要素のサイズではなく、キャプチャした画像の実サイズで計算する
                const displayW = DOM.canvas.offsetWidth;
                const displayH = DOM.canvas.offsetHeight;
                
                // 画面表示サイズと画像実サイズの比率を計算 (キャンバス描画用の赤枠に合わせる)
                const scaleX = w / displayW;
                const scaleY = h / displayH;

                // 赤枠は常に表示の中央 60% x 100px
                const rectW_disp = displayW * 0.6;
                const rectH_disp = 100;
                const rectX_disp = (displayW - rectW_disp) / 2;
                const rectY_disp = (displayH - rectH_disp) / 2;
                
                // 画像の実サイズにおけるROIを計算
                const rectX = Math.floor(rectX_disp * scaleX);
                const rectY = Math.floor(rectY_disp * scaleY);
                const rectWidth = Math.floor(rectW_disp * scaleX);
                const rectHeight = Math.floor(rectH_disp * scaleY);

                if (rectX < 0 || rectY < 0) throw new Error("計算エラー");

                // 切り出し
                let roiRect = new cv.Rect(rectX, rectY, rectWidth, rectHeight);
                roiMat = srcMat.roi(roiRect);

                // --- 画像処理パイプライン ---
                // 1. グレースケール
                grayMat = new cv.Mat();
                cv.cvtColor(roiMat, grayMat, cv.COLOR_RGBA2GRAY);
                
                // 2. ノイズ除去
                binMat = new cv.Mat();
                cv.medianBlur(grayMat, binMat, 3);
                grayMat.delete(); // 中間Matは解放

                // 3. 適応的二値化
                cv.adaptiveThreshold(binMat, binMat, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 15, 8);

                // 4. 2倍拡大
                let dsize = new cv.Size(binMat.cols * 2, binMat.rows * 2);
                cv.resize(binMat, binMat, dsize, 0, 0, cv.INTER_LINEAR);

                // 確認用Canvas作成 (OCRに渡す)
                const finalCanvas = document.createElement('canvas');
                cv.imshow(finalCanvas, binMat);

                // --- 項目別パラメータ設定 ---
                const definition = State.definitions.find(d => d.key === State.currentTarget);
                const label = definition.label;
                let opts = { tessedit_char_whitelist: '' }; // ホワイトリスト初期化

                if (label.includes("金額")) {
                    opts.tessedit_char_whitelist = '0123456789,¥円';
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
                    
                    // 連続読取防止
                    await new Promise(r => setTimeout(r, State.isCameraMode ? 1500 : 500));
                } else {
                    if (!State.isCameraMode) DOM.status.innerText = "文字が見つかりません (枠に合わせてください)";
                    // カメラモードでは状態をリセットしない
                }

            } catch (err) {
                console.error("Processor Execution Error:", err);
                DOM.status.innerText = "処理エラー: コンソールを確認してください";
            } finally {
                // メモリ解放
                if (srcMat) srcMat.delete();
                if (roiMat) roiMat.delete();
                // if (grayMat) grayMat.delete(); // grayMatは既に解放済み
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

            // ボタン生成と結果欄クリア
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

            // OCRエンジンを先に起動開始
            OCR.init();

            // デフォルトはカメラモード
            Actions.switchMode(true);
        },

        selectTarget: (key) => {
            // 現在のターゲットが同じなら何もしない
            if (State.currentTarget === key && State.isCameraMode) return;

            State.currentTarget = key;
            const label = State.definitions.find(d => d.key === key).label;
            
            // ボタンのハイライト
            Array.from(DOM.targetArea.children).forEach(btn => {
                btn.classList.remove('active-target');
                btn.style.backgroundColor = ''; 
                if (btn.textContent === label) {
                    btn.classList.add('active-target');
                    btn.style.backgroundColor = '#d1e7dd';
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
            
            // UIの状態更新
            DOM.btnCamera.classList.toggle('active', isCamera);
            DOM.btnFile.classList.toggle('active', !isCamera);
            
            if (isCamera) {
                // カメラモード
                DOM.image.style.display = 'none';
                DOM.video.style.display = 'block';
                Actions.startCamera();
                // 監視ループを再開
                if (!State.rafId) Actions.drawGuideLoop();
            } else {
                // ファイルモード
                Actions.stopCamera();
                DOM.video.style.display = 'none';
                DOM.image.style.display = DOM.image.src ? 'block' : 'none';
                DOM.status.innerText = DOM.image.src ? "画像読込完了。項目ボタンを押してOCR実行" : "画像をアップロードしてください";
            }
        },

        startCamera: async () => {
            if (State.stream) return; // すでに起動していたら何もしない
            DOM.status.innerText = "カメラ起動中...";
            try {
                // ★最重要の変更点: 最小限の要求 'video: true' で起動を試みる
                // facingMode: "environment" も削除し、ブラウザのデフォルトに任せる (通常は背面)
                const stream = await navigator.mediaDevices.getUserMedia({
                    video: true 
                    // 以前の設定: video: { facingMode: "environment", width: { ideal: 1280 }, height: { ideal: 720 } }
                });
                
                DOM.video.srcObject = stream;
                // play()は自動再生(autoplay)でカバーされますが、明示的に呼び出し
                DOM.video.play(); 
                State.stream = stream;
                DOM.status.innerText = "項目を選択してカメラを向けてください";
                
            } catch (err) {
                console.error("Camera Access Error:", err);
                // エラーの原因をコンソールに出力し、ユーザーにも分かりやすく伝える
                alert(`カメラを起動できませんでした。\n原因: ${err.name} - ${err.message}\n(ヒント: HTTPS接続か、ブラウザのカメラ権限を確認してください。)`);
                
                // 失敗したらファイルモードへ誘導し、カメラループを止める
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
            
            // 結果Inputの現在の値を取得して上書きする
            State.definitions.forEach(def => {
                const input = document.getElementById(`res-${def.key}`);
                if (input && input.value) {
                    State.ocrResults[def.key] = input.value;
                }
            });

            fetch(GOOGLE_SCRIPT_URL, {
                method: 'POST',
                // GASはCORS設定により 'no-cors' で送信するが、成功を確認できない点に注意
                mode: 'no-cors', 
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    action: 'appendData', // GAS側で処理しやすいようにactionを追加
                    sheetName: State.sheetName,
                    data: State.ocrResults
                })
            }).then(() => {
                // no-corsモードではPromiseが解決するだけで、実際の成功/失敗は確認できない
                DOM.status.innerText = "送信完了 (Sheetsを確認してください)";
                alert("送信完了しました");
            }).catch((e) => {
                 console.error("Fetch Error:", e);
                 DOM.status.innerText = "送信エラー (ネットワークまたはGAS設定の問題)";
                 alert("送信エラー");
            });
        },

        // 赤枠を描画し続けるループ
        drawGuideLoop: () => {
            const target = State.isCameraMode ? DOM.video : DOM.image;
            
            // 要素の表示サイズを取得
            const w = target.offsetWidth;
            const h = target.offsetHeight;

            // 要素が見えていて、サイズがある場合のみ描画
            if (w > 0 && h > 0 && target.style.display !== 'none') {
                DOM.canvas.width = w;
                DOM.canvas.height = h;
                
                // CSSでサイズを設定
                DOM.canvas.style.width = `${w}px`;
                DOM.canvas.style.height = `${h}px`;

                const ctx = DOM.ctx;
                ctx.clearRect(0, 0, w, h);
                
                // 赤枠の計算 (表示サイズの60% x 100px)
                const rectW = w * 0.6;
                const rectH = 100;
                const x = (w - rectW) / 2;
                const y = (h - rectH) / 2;

                // 赤枠描画
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
        Actions.switchMode(false); // ファイルモードへ切替
        DOM.fileInput.click();     // ファイル選択ダイアログを表示
    }); 

    DOM.fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        // ファイルモード時に画像が選択されたら、画像を表示
        const reader = new FileReader();
        reader.onload = (ev) => {
            DOM.image.src = ev.target.result;
            DOM.image.onload = () => {
                 DOM.image.style.display = 'block';
                 DOM.status.innerText = "画像読込完了。項目ボタンを押してOCR実行";
                 // 画像ロード後に赤枠描画を開始 (ループが止まっていた場合)
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
        // OpenCVとOCRが準備完了していて、カメラモード、ターゲット選択済み、処理中でない場合
        if (State.isCameraMode && State.currentTarget && !State.isProcessing && isCvLoaded && State.isOCRReady) {
            Processor.execute();
        }
    }, 1500);

    // 初期ロード時のステータス設定
    if (typeof cv === 'undefined') {
        DOM.status.innerText = "OpenCV.jsをロード中...";
    } else {
        DOM.status.innerText = "書類を選択してください";
    }

});
