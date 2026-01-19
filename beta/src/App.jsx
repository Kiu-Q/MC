import React, { useState, useRef, useEffect, useCallback } from 'react';
import { 
    FileSpreadsheet, Trash2, CheckCircle2, 
    ChevronRight, ChevronDown, ChevronUp, Download, ZoomIn, ZoomOut, 
    LayoutTemplate, ChevronLeft, ScanLine, GraduationCap, Hash, 
    X, Loader2, FolderInput, Files, BrainCircuit, Anchor, Search, FileText, HelpCircle
} from 'lucide-react';

// --- External Libraries (Dynamic Load) ---

// PDF.js
const PDF_LIB_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
const PDF_WORKER_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

// SheetJS (XLSX)
const XLSX_LIB_URL = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";

// Tesseract.js (OCR for Alignment)
const TESSERACT_LIB_URL = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";

const loadPdfLib = async () => {
    if (window.pdfjsLib) return window.pdfjsLib;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = PDF_LIB_URL;
        script.onload = () => {
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_URL;
            resolve(window.pdfjsLib);
        };
        script.onerror = reject;
        document.body.appendChild(script);
    });
};

const loadXlsxLib = async () => {
    if (window.XLSX) return window.XLSX;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = XLSX_LIB_URL;
        script.onload = () => resolve(window.XLSX);
        script.onerror = reject;
        document.body.appendChild(script);
    });
};

const loadTesseractLib = async () => {
    if (window.Tesseract) return window.Tesseract;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = TESSERACT_LIB_URL;
        script.onload = () => resolve(window.Tesseract);
        script.onerror = reject;
        document.body.appendChild(script);
    });
};

// Teachable Machine Loaders
const loadTeachableMachineLibs = async () => {
    if (window.tmImage) return window.tmImage;

    // 1. Load Tensorflow.js
    await new Promise((resolve, reject) => {
        if (window.tf) { resolve(); return; }
        const script = document.createElement('script');
        script.src = "https://cdn.jsdelivr.net/npm/@tensorflow/tfjs@latest/dist/tf.min.js";
        script.onload = resolve;
        script.onerror = reject;
        document.body.appendChild(script);
    });

    // 2. Load Teachable Machine Image
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = "https://cdn.jsdelivr.net/npm/@teachablemachine/image@latest/dist/teachablemachine-image.min.js";
        script.onload = () => resolve(window.tmImage);
        script.onerror = reject;
        document.body.appendChild(script);
    });
};

const App = () => {
    // --- State ---
    const [pages, setPages] = useState([]); // { id, imageUrl, width, height, results, regions, originalName }
    const [currentPageIndex, setCurrentPageIndex] = useState(0);
    const [pageInput, setPageInput] = useState("1");
    const [scale, setScale] = useState(1);
    const [selectedRegionId, setSelectedRegionId] = useState(null);
    const [isProcessing, setIsProcessing] = useState(false);
    const [progress, setProgress] = useState({ current: 0, total: 0, status: '' });
    const [showResultsSidebar, setShowResultsSidebar] = useState(true);
    const [toast, setToast] = useState(null);

    // Grading & Config
    const [answerKeyInput, setAnswerKeyInput] = useState("");
    const [parsedAnswerKey, setParsedAnswerKey] = useState({});
    const [isMarksSettingsOpen, setIsMarksSettingsOpen] = useState(false);
    const [marksConfig, setMarksConfig] = useState([{ start: 1, end: 100, marks: 1 }]);
    
    // AI Model State
    const [tmModel, setTmModel] = useState(null);
    
    // Batch Wizard State
    const [isBatchMode, setIsBatchMode] = useState(false);
    const [batchStep, setBatchStep] = useState(0); // 0: Upload, 1: Template
    const [batchPdfFiles, setBatchPdfFiles] = useState([]);
    const [previewPdfDoc, setPreviewPdfDoc] = useState(null);
    const [previewPdfPage, setPreviewPdfPage] = useState(1);
    const [previewTotalPages, setPreviewTotalPages] = useState(1);
    const [templateRegions, setTemplateRegions] = useState([]);
    const [isPdfLoading, setIsPdfLoading] = useState(false);
    const [autoAlign, setAutoAlign] = useState(true); // Default true per request
    const [anchorWord, setAnchorWord] = useState("Name");
    const [pagesPerExam, setPagesPerExam] = useState(1);

    // Refs
    const canvasRef = useRef(null);
    const templateCanvasRef = useRef(null);
    const folderInputRef = useRef(null);

    // --- Keyboard Shortcuts ---
    useEffect(() => {
        const handleKeyDown = (e) => {
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

            switch(e.key) {
                case 'Delete':
                case 'Backspace':
                    if (selectedRegionId && !isBatchMode) {
                        deleteSelectedRegion();
                    }
                    break;
                case 'ArrowLeft':
                    handlePageNavigation('prev');
                    break;
                case 'ArrowRight':
                    handlePageNavigation('next');
                    break;
                case '=':
                case '+':
                    setScale(s => Math.min(3, s + 0.1));
                    break;
                case '-':
                    setScale(s => Math.max(0.5, s - 0.1));
                    break;
                default:
                    break;
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [selectedRegionId, currentPageIndex, pages, isBatchMode]);

    // --- Effects ---
    useEffect(() => {
        if (toast) {
            const timer = setTimeout(() => setToast(null), 3000);
            return () => clearTimeout(timer);
        }
    }, [toast]);

    useEffect(() => {
        if (!isBatchMode && pages.length > 0 && pages[currentPageIndex]) {
            drawImageAndRegions();
            setPageInput((currentPageIndex + 1).toString());
        }
    }, [pages, currentPageIndex, scale, selectedRegionId, isBatchMode]);

    useEffect(() => {
        if (isBatchMode && batchStep === 1) {
            drawTemplateCanvas();
        }
    }, [batchStep, templateRegions, scale, previewPdfPage, isBatchMode]);

    // --- Special Answer Key Logic ---
    useEffect(() => {
        const rawText = answerKeyInput.toUpperCase();
        const tokens = rawText.split(/\s+/);
        
        const keyMap = {};
        let currentIndex = 1;

        tokens.forEach(token => {
            if (!/[E-Z]/.test(token)) {
                const matches = token.match(/[A-D]/g);
                if (matches) {
                    matches.forEach(char => {
                        keyMap[currentIndex] = char;
                        currentIndex++;
                    });
                }
            }
        });

        setParsedAnswerKey(keyMap);
    }, [answerKeyInput]);

    // --- Helpers ---
    const showToast = (message, type = 'success') => setToast({ message, type });

    const handlePageNavigation = (direction) => {
        if (direction === 'prev' && currentPageIndex > 0) setCurrentPageIndex(curr => curr - 1);
        if (direction === 'next' && currentPageIndex < pages.length - 1) setCurrentPageIndex(curr => curr + 1);
    };

    const deleteSelectedRegion = () => {
        if (!selectedRegionId) return;
        const newPages = [...pages];
        const page = newPages[currentPageIndex];
        page.regions = page.regions.filter(r => r.id !== selectedRegionId);
        // Renumber
        page.regions.forEach((r, idx) => r.label = `Q${idx + 1}`);
        setPages(newPages);
        setSelectedRegionId(null);
    };

    const getScore = (page) => {
        let score = 0;
        let total = 0;
        if (!page.results || !parsedAnswerKey) return { score: 0, total: 0 };
        
        Object.entries(page.results).forEach(([qLabel, res]) => {
            const qNum = parseInt(qLabel.replace(/\D/g, '')) || 0;
            if (qNum === 0) return;

            const config = marksConfig.find(c => qNum >= c.start && qNum <= c.end);
            const weight = config ? (parseFloat(config.marks) || 0) : 1;
            
            const correctAns = parsedAnswerKey[qNum.toString()];
            
            if (correctAns) {
                total += weight;
                if (res.label === correctAns) {
                    score += weight;
                }
            }
        });
        return { score, total };
    };

    // --- PDF Logic (Updated) ---
    const handleFileSelect = async (e) => {
        const files = Array.from(e.target.files).filter(f => f.type === 'application/pdf');
        if (files.length === 0) {
            showToast("No PDF files selected", "error");
            return;
        }

        setIsPdfLoading(true);
        try {
            const pdfjs = await loadPdfLib();
            
            // Use the first file to set up the template
            const firstFile = files[0];
            const arrayBuffer = await firstFile.arrayBuffer();
            const pdfDoc = await pdfjs.getDocument(arrayBuffer).promise;

            setBatchPdfFiles(files);
            setPreviewPdfDoc(pdfDoc);
            setPreviewTotalPages(pdfDoc.numPages);
            setPreviewPdfPage(1);
            setTemplateRegions([]);
            setBatchStep(1); // Go to Template Step
            setIsBatchMode(true);
        } catch (err) {
            console.error(err);
            showToast("Failed to load PDF preview", "error");
        } finally {
            setIsPdfLoading(false);
        }
    };

    const renderPdfToImage = async (pdfDoc, pageNum) => {
        const page = await pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale: 2.0 });
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({ canvasContext: context, viewport: viewport }).promise;
        return {
            dataUrl: canvas.toDataURL('image/jpeg', 0.8),
            width: viewport.width / 2.0,
            height: viewport.height / 2.0
        };
    };

    // --- Auto OCR Alignment Helper ---
    const findUniqueAnchor = async (worker, canvas) => {
        const dataUrl = canvas.toDataURL('image/jpeg', 0.8);
        const { data: { words } } = await worker.recognize(dataUrl);
        
        const wordCounts = {};
        const wordPositions = {};

        words.forEach(w => {
            const text = w.text.trim();
            if (text.length > 3 && /^[a-zA-Z]+$/.test(text) && w.confidence > 80) {
                if (!wordCounts[text]) {
                    wordCounts[text] = 0;
                    wordPositions[text] = { x: w.bbox.x0, y: w.bbox.y0 };
                }
                wordCounts[text]++;
            }
        });

        const candidates = Object.keys(wordCounts).filter(word => wordCounts[word] === 1);
        if (candidates.length === 0) return null;

        // Prefer top-left
        candidates.sort((a, b) => {
            const posA = wordPositions[a];
            const posB = wordPositions[b];
            return (posA.y - posB.y) || (posA.x - posB.x);
        });

        const bestWord = candidates[0];
        return { word: bestWord, pos: wordPositions[bestWord] };
    };

    const findWordPosition = async (worker, canvas, wordToFind) => {
        if (!wordToFind) return null;
        
        const dataUrl = canvas.toDataURL('image/jpeg', 0.7); 
        const { data: { words } } = await worker.recognize(dataUrl);
        
        const target = wordToFind.toLowerCase();
        const match = words.find(w => w.text.toLowerCase().trim() === target);
        
        if (match) {
            return { x: match.bbox.x0, y: match.bbox.y0 };
        }
        return null;
    };

    const handleBatchExtract = async () => {
        if (templateRegions.length === 0) {
            showToast("Please define at least one answer box.", "error");
            return;
        }

        setIsProcessing(true);
        setIsBatchMode(false);
        const newPages = [];
        const pdfjs = await loadPdfLib();

        let tesseractWorker = null;

        try {
            let anchorData = null; 

            // Initialize Alignment (Built-in, no toggle needed)
            setProgress({ current: 0, total: 0, status: "Initializing Auto-Alignment..." });
            await loadTesseractLib();
            tesseractWorker = await window.Tesseract.createWorker('eng');
            
            // 1. Analyze Template
            setProgress({ current: 0, total: 0, status: "Scanning template for anchor..." });
            const page = await previewPdfDoc.getPage(previewPdfPage);
            const viewport = page.getViewport({ scale: 2.0 });
            const canvas = document.createElement('canvas');
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            const ctx = canvas.getContext('2d');
            await page.render({ canvasContext: ctx, viewport }).promise;
            
            anchorData = await findUniqueAnchor(tesseractWorker, canvas);
            
            if (!anchorData) {
                console.warn("Could not find a unique anchor word. Alignment disabled.");
            } else {
                console.log(`Auto-Alignment Anchor: "${anchorData.word}" at`, anchorData.pos);
            }

            // 2. Determine Tasks (Single-File vs Multi-File)
            let tasks = [];
            
            if (batchPdfFiles.length === 1 && pagesPerExam > 0) {
                // Single File Mode: Iterate pages
                const file = batchPdfFiles[0];
                const arrayBuffer = await file.arrayBuffer();
                const doc = await pdfjs.getDocument(arrayBuffer).promise;
                
                for (let p = previewPdfPage; p <= doc.numPages; p += pagesPerExam) {
                    tasks.push({ 
                        doc, 
                        pageNum: p, 
                        originalName: `${file.name} (P${p})` 
                    });
                }
            } else {
                // Multiple Files Mode
                for (let i = 0; i < batchPdfFiles.length; i++) {
                    const file = batchPdfFiles[i];
                    const arrayBuffer = await file.arrayBuffer();
                    const doc = await pdfjs.getDocument(arrayBuffer).promise;
                    if (doc.numPages >= previewPdfPage) {
                        tasks.push({
                            doc,
                            pageNum: previewPdfPage,
                            originalName: file.name
                        });
                    }
                }
            }

            // 3. Process Tasks
            for (let i = 0; i < tasks.length; i++) {
                const task = tasks[i];
                setProgress({ current: i + 1, total: tasks.length, status: `Extracting ${task.originalName}...` });

                const page = await task.doc.getPage(task.pageNum);
                const viewport = page.getViewport({ scale: 2.0 });
                const canvas = document.createElement('canvas');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                const ctx = canvas.getContext('2d');
                await page.render({ canvasContext: ctx, viewport }).promise;

                // 3a. Calculate Offset
                let offsetX = 0;
                let offsetY = 0;
                
                if (anchorData && tesseractWorker) {
                    const targetPos = await findWordPosition(tesseractWorker, canvas, anchorData.word);
                    if (targetPos) {
                        offsetX = targetPos.x - anchorData.pos.x;
                        offsetY = targetPos.y - anchorData.pos.y;
                        
                        if (Math.abs(offsetX) > canvas.width * 0.2 || Math.abs(offsetY) > canvas.height * 0.2) {
                            offsetX = 0; offsetY = 0;
                        }
                    }
                }

                // 4. Adjust Regions (Direct Offset Application requested)
                const adjustedRegions = templateRegions.map(r => ({
                    ...r,
                    x: r.x + offsetX,
                    y: r.y + offsetY
                }));

                newPages.push({
                    id: Date.now() + i,
                    imageUrl: canvas.toDataURL('image/jpeg', 0.8),
                    width: viewport.width / 2.0,
                    height: viewport.height / 2.0,
                    originalName: task.originalName,
                    regions: adjustedRegions,
                    results: {}
                });
            }
            
            if (tesseractWorker) await tesseractWorker.terminate();

            setPages(newPages); 
            setCurrentPageIndex(0);
            showToast(`Imported ${newPages.length} pages`, "success");
        } catch (e) {
            console.error(e);
            showToast("Error extracting: " + e.message, "error");
        } finally {
            if (tesseractWorker) try { await tesseractWorker.terminate(); } catch(e){}
            setIsProcessing(false);
            setProgress({ current: 0, total: 0, status: '' });
        }
    };

    // --- Drawing Logic ---
    const getMousePos = (e, canvas) => {
        const rect = canvas.getBoundingClientRect();
        return {
            x: (e.clientX - rect.left) / scale,
            y: (e.clientY - rect.top) / scale
        };
    };

    const handleCanvasMouseDown = (e, isTemplate = false) => {
        const canvas = isTemplate ? templateCanvasRef.current : canvasRef.current;
        const currentRegions = isTemplate ? templateRegions : pages[currentPageIndex]?.regions;
        const setReg = isTemplate ? setTemplateRegions : (newRegs) => {
            const updatedPages = [...pages];
            updatedPages[currentPageIndex].regions = newRegs;
            setPages(updatedPages);
        };
        
        if (!canvas || !currentRegions) return;

        const pos = getMousePos(e, canvas);

        const clickedIndex = currentRegions.findIndex(r => 
            pos.x >= r.x && pos.x <= r.x + r.width &&
            pos.y >= r.y && pos.y <= r.y + r.height
        );

        if (clickedIndex >= 0) {
            if (e.button === 2 || e.shiftKey) { // Right click delete
                const newRegions = [...currentRegions];
                newRegions.splice(clickedIndex, 1);
                newRegions.forEach((r, idx) => r.label = `Q${idx + 1}`);
                setReg(newRegions);
                if (!isTemplate) setSelectedRegionId(null);
                return;
            } 
            
            if (!isTemplate) {
                setSelectedRegionId(currentRegions[clickedIndex].id);
                return;
            }
        }

        if (!isTemplate) setSelectedRegionId(null);

        const startX = pos.x;
        const startY = pos.y;
        
        const onMove = (moveEvent) => {
            const movePos = getMousePos(moveEvent, canvas);
            if (isTemplate) drawTemplateCanvas({ x: startX, y: startY, w: movePos.x - startX, h: movePos.y - startY });
            else drawImageAndRegions({ x: startX, y: startY, w: movePos.x - startX, h: movePos.y - startY });
        };

        const onUp = (upEvent) => {
            const upPos = getMousePos(upEvent, canvas);
            const w = upPos.x - startX;
            const h = upPos.y - startY;

            if (Math.abs(w) > 5 && Math.abs(h) > 5) {
                const newRegion = {
                    id: Date.now().toString(),
                    x: w > 0 ? startX : upPos.x,
                    y: h > 0 ? startY : upPos.y,
                    width: Math.abs(w),
                    height: Math.abs(h),
                    label: `Q${currentRegions.length + 1}`
                };
                setReg([...currentRegions, newRegion]);
            }
            
            if (isTemplate) drawTemplateCanvas();
            else drawImageAndRegions();

            canvas.removeEventListener('mousemove', onMove);
            window.removeEventListener('mouseup', onUp);
        };

        canvas.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
    };

    const drawInternal = async (canvas, regions, imageSource, ghostRect = null, isTemplate = false) => {
        if (!canvas) return;
        const ctx = canvas.getContext('2d');
        const img = new Image();
        img.src = imageSource;
        
        try {
            await new Promise((resolve, reject) => {
                img.onload = resolve;
                img.onerror = () => reject(new Error("Image load failed"));
            });
        } catch (err) {
            console.error(err);
            return;
        }

        canvas.width = img.width * scale;
        canvas.height = img.height * scale;
        ctx.scale(scale, scale);
        ctx.drawImage(img, 0, 0, img.width, img.height);

        regions.forEach(r => {
            const isSelected = !isTemplate && r.id === selectedRegionId;
            ctx.strokeStyle = isSelected ? '#58a6ff' : '#238636';
            ctx.lineWidth = 2;
            ctx.fillStyle = isSelected ? 'rgba(88, 166, 255, 0.2)' : 'rgba(35, 134, 54, 0.1)';
            ctx.fillRect(r.x, r.y, r.width, r.height);
            ctx.strokeRect(r.x, r.y, r.width, r.height);

            ctx.fillStyle = isSelected ? '#58a6ff' : '#238636';
            const labelW = ctx.measureText(r.label).width + 8;
            ctx.fillRect(r.x, r.y - 18, labelW, 18);
            ctx.fillStyle = '#ffffff';
            ctx.font = '12px monospace';
            ctx.textBaseline = 'middle';
            ctx.fillText(r.label, r.x + 4, r.y - 9);
        });

        if (ghostRect) {
            ctx.strokeStyle = '#d29922';
            ctx.lineWidth = 2;
            ctx.setLineDash([5, 3]);
            ctx.strokeRect(ghostRect.x, ghostRect.y, ghostRect.w, ghostRect.h);
            ctx.setLineDash([]);
        }
    };

    const drawImageAndRegions = (ghostRect = null) => {
        if (!pages[currentPageIndex]) return;
        drawInternal(canvasRef.current, pages[currentPageIndex].regions, pages[currentPageIndex].imageUrl, ghostRect, false);
    };

    const drawTemplateCanvas = async (ghostRect = null) => {
        if (!previewPdfDoc) return;
        const render = await renderPdfToImage(previewPdfDoc, previewPdfPage);
        drawInternal(templateCanvasRef.current, templateRegions, render.dataUrl, ghostRect, true);
    };

    // --- Upload Scan Helper (Proxy) ---
    const uploadScanResult = async (canvas, label, confidence) => {
        const timestamp = new Date().toISOString();
        const filename = `${timestamp.replace(/[:.]/g, '-')}_${label}_${confidence.toFixed(2)}`;
        const base64 = canvas.toDataURL('image/jpeg', 0.8);
        const FUNCTION_ENDPOINT = "/.netlify/functions/add-scan"; 

        try {
            const res = await fetch(FUNCTION_ENDPOINT, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename, image_data: base64, created_at: timestamp })
            });
            if (!res.ok) {
                const errText = await res.text();
                console.error("DB Upload Failed", res.status, errText);
            }
        } catch (e) {
            console.error("DB Upload Error", e);
        }
    };

    // --- Teachable Machine Scanning ---
    const handleBatchScan = async () => {
        if (pages.length === 0) return;
        setIsProcessing(true);
        const updatedPages = [...pages];
        try {
            let model = tmModel;
            if (!model) {
                setProgress({ current: 0, total: 0, status: "Loading AI Model..." });
                await loadTeachableMachineLibs();
                const URL = "https://teachablemachine.withgoogle.com/models/YDV-9wqBW/";
                model = await window.tmImage.load(URL + "model.json", URL + "metadata.json");
                setTmModel(model);
            }
            let totalRegions = 0;
            updatedPages.forEach(p => totalRegions += (p.regions ? p.regions.length : 0));
            let processedCount = 0;

            for (let i = 0; i < updatedPages.length; i++) {
                const page = updatedPages[i];
                if (!page.regions || page.regions.length === 0) continue;
                const img = new Image();
                img.src = page.imageUrl;
                await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = reject; });
                page.results = {};

                for (const region of page.regions) {
                    setProgress({ current: processedCount, total: totalRegions, status: `Scanning Page ${i+1}: ${region.label}...` });
                    const cCanvas = document.createElement('canvas');
                    cCanvas.width = region.width;
                    cCanvas.height = region.height;
                    const ctx = cCanvas.getContext('2d');
                    ctx.drawImage(img, region.x, region.y, region.width, region.height, 0, 0, region.width, region.height);
                    
                    const predictions = await model.predict(cCanvas);
                    const best = predictions.reduce((prev, current) => (prev.probability > current.probability) ? prev : current);
                    const CONFIDENCE_THRESHOLD = 0.5;
                    const finalLabel = best.probability >= CONFIDENCE_THRESHOLD ? best.className : "EMPTY";
                    
                    // Upload DB
                    uploadScanResult(cCanvas, finalLabel, best.probability);

                    page.results[region.label] = { label: finalLabel, confidence: best.probability };
                    processedCount++;
                }
            }
            setPages(updatedPages);
            showToast("AI Scanning Complete!", "success");
        } catch (e) {
            console.error(e);
            showToast("Scan failed: " + e.message, "error");
        } finally {
            setIsProcessing(false);
            setProgress({ current: 0, total: 0, status: '' });
        }
    };

    const handleExportExcel = async () => {
        try {
            const XLSX = await loadXlsxLib();
            const allQLabels = new Set();
            pages.forEach(p => { if (p.results) Object.keys(p.results).forEach(k => allQLabels.add(k)); });
            const sortedLabels = Array.from(allQLabels).sort((a, b) => {
                const nA = parseInt(a.replace(/\D/g, '')) || 0;
                const nB = parseInt(b.replace(/\D/g, '')) || 0;
                return nA - nB;
            });
            const header = ['Student File', 'Total Score', 'Max Score', 'Percentage', ...sortedLabels];
            const rows = pages.map((page, i) => {
                const { score, total } = getScore(page);
                const rowIndex = i + 2; 
                const row = [page.originalName || `Page ${i + 1}`, score, total, { t: 'n', f: `B${rowIndex}/C${rowIndex}`, z: '0.0%' }];
                sortedLabels.forEach(label => {
                    const res = page.results?.[label];
                    row.push(res ? res.label : '-');
                });
                return row;
            });
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
            const getColLetter = (colIndex) => {
                let letter = '';
                while (colIndex >= 0) { letter = String.fromCharCode((colIndex % 26) + 65) + letter; colIndex = Math.floor(colIndex / 26) - 1; }
                return letter;
            };
            ws['!autofilter'] = { ref: `A1:${getColLetter(header.length - 1)}1` };
            XLSX.utils.book_append_sheet(wb, ws, "Grades");
            XLSX.writeFile(wb, "Handwritten_Scan_Grades.xlsx");
            showToast("Excel exported successfully");
        } catch (e) {
            console.error(e);
            showToast("Failed to load Excel library", "error");
        }
    };

    // --- UI ---
    const theme = { bg: 'bg-[#0d1117]', sidebar: 'bg-[#161b22]', border: 'border-[#30363d]', text: 'text-[#c9d1d9]', textMuted: 'text-[#8b949e]', accent: 'text-[#58a6ff]', accentBg: 'bg-[#1f6feb]', success: 'text-[#3fb950]', buttonHover: 'hover:bg-[#21262d]', inputBg: 'bg-[#0d1117]' };

    if (isBatchMode) {
        return (
            <div className={`h-screen w-screen ${theme.bg} ${theme.text} flex flex-col font-sans`}>
                <style>{`:root, body, #root { height: 100%; width: 100%; margin: 0; padding: 0; max-width: none !important; }`}</style>
                <div className={`h-16 border-b ${theme.border} ${theme.sidebar} flex items-center justify-between px-6`}>
                    <h2 className={`font-bold flex items-center gap-2 ${theme.accent}`}><Files size={20}/> Batch PDF Import</h2>
                    <button onClick={() => setIsBatchMode(false)} className={`p-2 ${theme.buttonHover} rounded-md`}><X size={20}/></button>
                </div>
                <div className="flex-1 overflow-hidden flex">
                    {batchStep === 0 && (
                        <div className="flex-1 flex flex-col items-center justify-center p-10">
                            <div className={`border-2 border-dashed ${theme.border} rounded-xl ${theme.bg} p-12 flex flex-col items-center text-center max-w-lg w-full`}>
                                <FolderInput size={48} className={`${theme.textMuted} mb-4`}/>
                                <h3 className="text-xl font-semibold mb-2">Select PDF File(s)</h3>
                                <p className={`${theme.textMuted} mb-6 text-sm`}>Select one or more PDF files.<br/>(Single file = whole class, Multiple files = one student each)</p>
                                <input type="file" ref={folderInputRef} className="hidden" multiple accept="application/pdf" onChange={handleFileSelect} />
                                <button onClick={() => folderInputRef.current.click()} disabled={isPdfLoading} className={`px-6 py-3 ${theme.accentBg} hover:opacity-90 text-white rounded-md font-medium flex items-center gap-2`}>
                                    {isPdfLoading ? <Loader2 className="animate-spin"/> : <FolderInput size={18}/>} Choose Files
                                </button>
                            </div>
                        </div>
                    )}
                    {batchStep === 1 && (
                        <div className="flex-1 flex overflow-hidden">
                            <div className={`w-80 ${theme.sidebar} border-r ${theme.border} flex flex-col p-4`}>
                                <h3 className="font-semibold mb-4 text-[#e6edf3]">Template Setup</h3>
                                <div className="mb-4">
                                    <label className={`text-xs ${theme.textMuted} uppercase font-bold`}>Page Selection</label>
                                    <div className="flex items-center gap-2 mt-1">
                                        <button onClick={() => setPreviewPdfPage(p => Math.max(1, p - 1))} className={`p-1 bg-[#21262d] rounded border ${theme.border}`}><ChevronLeft size={16}/></button>
                                        <span className="text-sm font-mono">{previewPdfPage} / {previewTotalPages}</span>
                                        <button onClick={() => setPreviewPdfPage(p => Math.min(previewTotalPages, p + 1))} className={`p-1 bg-[#21262d] rounded border ${theme.border}`}><ChevronRight size={16}/></button>
                                    </div>
                                </div>
                                <div className="mb-4 space-y-3">
                                    <p className={`text-[10px] ${theme.textMuted} flex items-center gap-1`}><Anchor size={12}/> Auto-Align Enabled</p>
                                    {batchPdfFiles.length === 1 && (
                                        <div>
                                            <label className={`text-xs ${theme.textMuted} uppercase font-bold mb-1 block`}>Pages Per Student</label>
                                            <input type="number" min="1" value={pagesPerExam} onChange={(e) => setPagesPerExam(parseInt(e.target.value) || 1)} className={`w-full ${theme.inputBg} border ${theme.border} rounded p-2 text-sm text-[#c9d1d9] focus:border-[#58a6ff] outline-none`}/>
                                            <p className={`text-[10px] ${theme.textMuted} mt-1`}>Use if one file contains multiple exams.</p>
                                        </div>
                                    )}
                                </div>
                                <div className={`flex-1 overflow-y-auto mb-4 border ${theme.border} rounded-md ${theme.bg} p-2`}>
                                    <div className="space-y-1">
                                        {templateRegions.map((r, i) => (
                                            <div key={i} className={`flex justify-between items-center text-xs p-2 bg-[#21262d] rounded border ${theme.border}`}>
                                                <span className={`font-mono ${theme.accent}`}>{r.label}</span>
                                                <button onClick={() => { const n = [...templateRegions]; n.splice(i, 1); setTemplateRegions(n); }} className="text-[#f85149]"><Trash2 size={12}/></button>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                                <button onClick={handleBatchExtract} className={`w-full py-2 ${theme.accentBg} text-white rounded-md font-medium flex items-center justify-center gap-2`}><CheckCircle2 size={16}/> Apply Template</button>
                            </div>
                            <div className={`flex-1 ${theme.bg} overflow-auto flex items-center justify-center p-8 relative`}>
                                <div className={`absolute top-4 right-4 flex gap-1 ${theme.sidebar} p-1 rounded border ${theme.border} z-20 shadow-lg`}>
                                    <button onClick={() => setScale(s => Math.max(0.5, s - 0.1))} className={`p-1.5 ${theme.buttonHover} rounded ${theme.text}`}><ZoomOut size={16}/></button>
                                    <span className={`text-xs font-mono flex items-center px-2 ${theme.textMuted}`}>{Math.round(scale * 100)}%</span>
                                    <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className={`p-1.5 ${theme.buttonHover} rounded ${theme.text}`}><ZoomIn size={16}/></button>
                                </div>
                                <div className={`relative shadow-2xl border ${theme.border}`}>
                                    <canvas ref={templateCanvasRef} onMouseDown={(e) => handleCanvasMouseDown(e, true)} className="cursor-crosshair block"/>
                                </div>
                            </div>
                        </div>
                    )}
                </div>
            </div>
        );
    }

    return (
        <div className={`flex h-screen w-full ${theme.bg} ${theme.text} font-sans overflow-hidden`}>
            <style>{`:root, body, #root { height: 100%; width: 100%; margin: 0; padding: 0; max-width: none !important; }`}</style>
            <div className={`w-64 ${theme.sidebar} border-r ${theme.border} flex flex-col flex-shrink-0`}>
                <div className={`p-4 border-b ${theme.border} flex items-center gap-2`}><ScanLine className="text-[#238636]" /> <span className="font-bold text-sm">Handwritten MC Scanner</span></div>
                <div className={`p-3 border-b ${theme.border}`}>
                    <button onClick={() => { setIsBatchMode(true); setBatchStep(0); }} className={`w-full py-2 border border-dashed ${theme.border} rounded-md ${theme.buttonHover} ${theme.accent} text-xs font-medium flex items-center justify-center gap-2 transition-colors`}><FolderInput size={16} /> Import File(s)</button>
                </div>
                <div className="flex-1 overflow-y-auto p-2 space-y-2">
                    {pages.map((page, idx) => (
                        <div key={page.id} onClick={() => setCurrentPageIndex(idx)} className={`p-2 rounded-md border cursor-pointer flex items-center gap-3 group transition-all ${currentPageIndex === idx ? `bg-[#21262d] border-[#58a6ff]` : `${theme.bg} ${theme.border}`}`}>
                            <div className={`w-8 h-8 bg-[#21262d] rounded flex items-center justify-center ${theme.textMuted} overflow-hidden flex-shrink-0`}>{page.imageUrl ? <img src={page.imageUrl} className="w-full h-full object-cover"/> : <FileSpreadsheet size={14}/>}</div>
                            <div className="flex-1 min-w-0"><div className="text-xs font-medium text-[#e6edf3] truncate">{page.originalName || `Page ${idx+1}`}</div><div className={`flex items-center gap-2 text-[10px] ${theme.textMuted}`}><span>{getScore(page).score}/{getScore(page).total}</span></div></div>
                        </div>
                    ))}
                </div>
                {pages.length > 0 && (
                    <div className={`p-3 border-t ${theme.border}`}>
                        <button onClick={handleBatchScan} disabled={isProcessing} className="w-full py-2 bg-[#238636] hover:bg-[#2ea043] disabled:opacity-50 text-white rounded-md font-medium flex items-center justify-center gap-2 text-sm">
                            {isProcessing ? <Loader2 className="animate-spin w-4 h-4"/> : <><BrainCircuit size={16}/> Scan All (AI)</>}
                        </button>
                    </div>
                )}
            </div>
            <div className="flex-1 flex flex-col relative overflow-hidden bg-[#010409]">
                <div className={`h-12 border-b ${theme.border} ${theme.sidebar} flex items-center justify-between px-4`}>
                    <div className="flex items-center gap-2">
                        <div className={`flex items-center gap-1 bg-[#21262d] rounded p-0.5 border ${theme.border}`}>
                             <button onClick={() => handlePageNavigation('prev')} className={`p-1 ${theme.buttonHover} rounded ${theme.text}`}><ChevronLeft size={14}/></button>
                             <input className={`w-8 bg-transparent text-center text-xs font-mono focus:outline-none ${theme.text}`} value={pageInput} onChange={(e) => { setPageInput(e.target.value); const val = parseInt(e.target.value); if(val > 0 && val <= pages.length) setCurrentPageIndex(val-1); }}/>
                             <button onClick={() => handlePageNavigation('next')} className={`p-1 ${theme.buttonHover} rounded ${theme.text}`}><ChevronRight size={14}/></button>
                        </div>
                    </div>
                    <div className="flex items-center gap-2">
                        <div className={`h-6 w-px bg-[#30363d] mx-2`}></div>
                        <button onClick={() => setScale(s => Math.max(0.5, s - 0.1))} className={`p-1.5 ${theme.buttonHover} rounded text-[#7d8590] hover:text-[#c9d1d9]`}><ZoomOut size={18}/></button>
                        <span className={`text-xs font-mono w-10 text-center ${theme.textMuted}`}>{Math.round(scale * 100)}%</span>
                        <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className={`p-1.5 ${theme.buttonHover} rounded text-[#7d8590] hover:text-[#c9d1d9]`}><ZoomIn size={18}/></button>
                        <div className={`h-6 w-px bg-[#30363d] mx-2`}></div>
                        <a href="https://github.com/Kiu-Q/MC" target="_blank" rel="noreferrer" className={`p-1.5 ${theme.textMuted} hover:${theme.text} ${theme.buttonHover} rounded-md`}><HelpCircle size={18}/></a>
                        <button onClick={() => setIsMarksSettingsOpen(true)} className={`p-1.5 ${theme.textMuted} hover:${theme.text} ${theme.buttonHover} rounded-md`} title="Weighting"><Hash size={18}/></button>
                        <button onClick={handleExportExcel} className={`p-1.5 ${theme.success} ${theme.buttonHover} rounded-md`} title="Export"><Download size={18}/></button>
                        <button onClick={() => setShowResultsSidebar(!showResultsSidebar)} className={`p-1.5 ${theme.accent} ${theme.buttonHover} rounded-md`}><LayoutTemplate size={18}/></button>
                    </div>
                </div>
                <div className="flex-1 overflow-auto flex items-center justify-center p-8 relative">
                    <div className={`${theme.bg} shadow-2xl relative border ${theme.border}`}>
                        <canvas ref={canvasRef} onMouseDown={(e) => handleCanvasMouseDown(e, false)} className="cursor-crosshair block"/>
                         {pages.length === 0 && (
                            <div className={`absolute inset-0 flex flex-col items-center justify-center ${theme.textMuted} pointer-events-none`}><ScanLine size={48} className="mb-4 opacity-20"/><p>Import Folder to Start</p></div>
                        )}
                    </div>
                </div>
                {isProcessing && (
                    <div className={`absolute inset-0 ${theme.bg}/80 z-50 flex items-center justify-center backdrop-blur-sm`}>
                        <div className={`${theme.sidebar} border ${theme.border} p-6 rounded-xl shadow-2xl max-w-sm w-full text-center`}>
                            <Loader2 size={32} className={`animate-spin ${theme.accent} mx-auto mb-4`}/><h3 className="text-[#e6edf3] font-medium mb-1">{progress.status}</h3><div className={`w-full ${theme.border} h-1.5 rounded-full overflow-hidden mt-3 bg-gray-700`}><div className="h-full bg-[#238636] transition-all duration-300" style={{ width: `${(progress.current/progress.total)*100}%` }}></div></div><p className={`text-xs ${theme.textMuted} mt-2`}>{progress.current} / {progress.total} items</p>
                        </div>
                    </div>
                )}
            </div>
            {showResultsSidebar && (
                <div className={`w-72 ${theme.sidebar} border-l ${theme.border} flex flex-col flex-shrink-0 z-20`}>
                    <div className={`p-4 border-b ${theme.border}`}>
                        <h3 className="font-bold text-[#e6edf3] flex items-center gap-2 mb-4"><GraduationCap className="text-[#e3b341]" size={18}/> Grading</h3>
                        <div className="space-y-3">
                            <div>
                                <label className={`text-xs font-bold ${theme.textMuted} uppercase mb-1 block`}>Answer Key</label>
                                <textarea className={`w-full ${theme.inputBg} border ${theme.border} rounded p-2 text-xs ${theme.text} font-mono h-24 focus:border-[#58a6ff] outline-none resize-none`} placeholder="Paste key (e.g. 1 A 2 B)" value={answerKeyInput} onChange={(e) => setAnswerKeyInput(e.target.value)}/>
                            </div>
                            {pages[currentPageIndex] && (
                                <div className={`flex items-center justify-between p-3 ${theme.inputBg} border ${theme.border} rounded-md`}><span className={`text-xs ${theme.textMuted}`}>Page Score</span><span className="text-lg font-bold text-[#e6edf3]">{getScore(pages[currentPageIndex]).score} <span className={`text-xs ${theme.textMuted} font-normal ml-1`}>/ {getScore(pages[currentPageIndex]).total}</span></span></div>
                            )}
                        </div>
                    </div>
                    <div className="flex-1 overflow-y-auto p-2">
                         {pages[currentPageIndex]?.results && Object.keys(pages[currentPageIndex].results).length > 0 ? (
                            <div className="space-y-1">
                                {Object.entries(pages[currentPageIndex].results).sort((a, b) => parseInt(a[0].replace(/\D/g, '')) - parseInt(b[0].replace(/\D/g, ''))).map(([key, val]) => {
                                        const qNum = parseInt(key.replace(/\D/g, '')) || 0;
                                        const correctAns = parsedAnswerKey[qNum.toString()];
                                        const isCorrect = correctAns && val.label === correctAns;
                                        const isGraded = !!correctAns;
                                        return (
                                            <div key={key} className={`flex items-center justify-between px-3 py-2 rounded border text-xs ${!isGraded ? `${theme.inputBg} ${theme.border}` : isCorrect ? 'bg-[#238636]/10 border-[#238636]/40' : 'bg-[#da3633]/10 border-[#da3633]/40'}`}>
                                                <div className="flex items-center gap-3"><span className={`font-mono ${theme.textMuted} w-6`}>{key}</span><span className={`font-bold ${isGraded ? (isCorrect ? 'text-[#3fb950]' : 'text-[#f85149]') : theme.text}`}>{val.label}</span></div>
                                                {isGraded && !isCorrect && (<span className={`font-mono ${theme.textMuted}`}>Exp: {correctAns}</span>)}
                                            </div>
                                        );
                                })}
                            </div>
                         ) : (<div className={`text-center ${theme.textMuted} py-10 text-xs`}>No scan results</div>)}
                    </div>
                </div>
            )}
            {isMarksSettingsOpen && (
                <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50">
                    <div className={`${theme.sidebar} border ${theme.border} p-6 rounded-xl w-80 shadow-2xl`}>
                        <div className="flex justify-between items-center mb-4"><h3 className="font-bold text-[#e6edf3]">Question Weighting</h3><button onClick={() => setIsMarksSettingsOpen(false)}><X size={16} className={theme.textMuted}/></button></div>
                        <div className="space-y-2 max-h-60 overflow-y-auto mb-4">{marksConfig.map((conf, i) => (<div key={i} className="flex gap-2 items-center"><input type="number" value={conf.start} onChange={e => { const n = [...marksConfig]; n[i].start = parseInt(e.target.value); setMarksConfig(n); }} className={`w-14 ${theme.inputBg} border ${theme.border} rounded p-1 text-xs`}/><span className={theme.textMuted}>-</span><input type="number" value={conf.end} onChange={e => { const n = [...marksConfig]; n[i].end = parseInt(e.target.value); setMarksConfig(n); }} className={`w-14 ${theme.inputBg} border ${theme.border} rounded p-1 text-xs`}/><span className={theme.textMuted}>:</span><input type="number" value={conf.marks} onChange={e => { const n = [...marksConfig]; n[i].marks = e.target.value; setMarksConfig(n); }} className={`w-12 ${theme.inputBg} border ${theme.border} rounded p-1 text-xs`}/><button onClick={() => {const n=[...marksConfig]; n.splice(i,1); setMarksConfig(n)}} className="text-[#f85149]"><Trash2 size={12}/></button></div>))}</div>
                        <button onClick={() => setMarksConfig([...marksConfig, {start:0, end:0, marks:1}])} className={`w-full py-1.5 border border-dashed ${theme.border} ${theme.textMuted} text-xs ${theme.buttonHover} rounded mb-2`}>+ Add Range</button>
                        <button onClick={() => setIsMarksSettingsOpen(false)} className="w-full py-2 bg-[#238636] text-white rounded font-medium text-xs">Done</button>
                    </div>
                </div>
            )}
            {toast && (<div className={`fixed bottom-6 right-6 px-4 py-3 rounded-md shadow-lg text-white text-xs font-medium z-50 flex items-center gap-3 animate-fade-in-up ${toast.type === 'error' ? 'bg-[#da3633]' : 'bg-[#238636]'}`}>{toast.type === 'error' ? <X size={14}/> : <CheckCircle2 size={14}/>}{toast.message}</div>)}
        </div>
    );
};

export default App;