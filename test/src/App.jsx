import React, { useState, useRef, useEffect } from 'react';
import {
    Trash2, CheckCircle2, ChevronRight, ChevronLeft, 
    ZoomIn, ZoomOut, LayoutTemplate, ScanLine, 
    X, Loader2, FolderInput, Files, Anchor, 
    Scissors, Download, Image as ImageIcon, Layers
} from 'lucide-react';

// --- External Libraries (Dynamic Load) ---

// PDF.js
const PDF_LIB_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
const PDF_WORKER_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

// Tesseract.js (用于自动对齐)
const TESSERACT_LIB_URL = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";

// JSZip (用于打包下载裁切图)
const JSZIP_LIB_URL = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";

const loadPdfLib = async () => {
    // If already loaded, ensure worker is set
    if (window.pdfjsLib) {
        if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
             window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_URL;
        }
        return window.pdfjsLib;
    }

    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = PDF_LIB_URL;
        script.onload = () => {
            if (window.pdfjsLib) {
                window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_URL;
                resolve(window.pdfjsLib);
            } else {
                reject(new Error("PDF.js library loaded but object not found"));
            }
        };
        script.onerror = () => reject(new Error("Failed to load PDF.js script"));
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

const loadJsZipLib = async () => {
    if (window.JSZip) return window.JSZip;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = JSZIP_LIB_URL;
        script.onload = () => resolve(window.JSZip);
        script.onerror = reject;
        document.body.appendChild(script);
    });
};

const App = () => {
    // --- State ---
    // pages 结构: { id, imageUrl, width, height, originalName, regions: [{x,y,w,h}], crops: [{label, dataUrl}] }
    const [pages, setPages] = useState([]); 
    const [currentPageIndex, setCurrentPageIndex] = useState(0);
    const [pageInput, setPageInput] = useState("1");
    const [scale, setScale] = useState(1);
    const [selectedRegionId, setSelectedRegionId] = useState(null);
    const [isProcessing, setIsProcessing] = useState(false);
    const [progress, setProgress] = useState({ current: 0, total: 0, status: '' });
    const [showResultsSidebar, setShowResultsSidebar] = useState(true);
    const [toast, setToast] = useState(null);

    // Batch Wizard State
    const [isBatchMode, setIsBatchMode] = useState(false);
    const [batchStep, setBatchStep] = useState(0); // 0: Upload, 1: Template
    const [batchPdfFiles, setBatchPdfFiles] = useState([]);
    const [previewPdfDoc, setPreviewPdfDoc] = useState(null);
    const [previewPdfPage, setPreviewPdfPage] = useState(1);
    const [previewTotalPages, setPreviewTotalPages] = useState(1);
    const [templateRegions, setTemplateRegions] = useState([]);
    const [isPdfLoading, setIsPdfLoading] = useState(false);
    const [pagesPerExam, setPagesPerExam] = useState(1); // 如果一个PDF有多份文件

    // Refs
    const canvasRef = useRef(null);
    const templateCanvasRef = useRef(null);
    const folderInputRef = useRef(null);

    // --- Keyboard Shortcuts ---
    useEffect(() => {
        const handleKeyDown = (e) => {
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

            switch (e.key) {
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
        // Renumber labels
        page.regions.forEach((r, idx) => r.label = `Region ${idx + 1}`);
        setPages(newPages);
        setSelectedRegionId(null);
    };

    // --- PDF Logic ---
    const handleFileSelect = async (e) => {
        const selectedFiles = e.target.files;
        if (!selectedFiles || selectedFiles.length === 0) return;

        const files = Array.from(selectedFiles).filter(f => f.type === 'application/pdf');
        if (files.length === 0) {
            showToast("No PDF files selected", "error");
            e.target.value = ''; // Reset input to allow re-selection
            return;
        }

        setIsPdfLoading(true);
        try {
            const pdfjs = await loadPdfLib();

            // Use the first file to set up the template
            const firstFile = files[0];
            const arrayBuffer = await firstFile.arrayBuffer();
            
            // Add timeout protection for PDF loading
            const loadingTask = pdfjs.getDocument(arrayBuffer);
            const pdfDoc = await Promise.race([
                loadingTask.promise,
                new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout loading PDF - check connection")), 15000))
            ]);

            setBatchPdfFiles(files);
            setPreviewPdfDoc(pdfDoc);
            setPreviewTotalPages(pdfDoc.numPages);
            setPreviewPdfPage(1);
            setTemplateRegions([]);
            setBatchStep(1); // Go to Template Step
        } catch (err) {
            console.error(err);
            showToast(`Failed to load PDF: ${err.message}`, "error");
        } finally {
            setIsPdfLoading(false);
            e.target.value = ''; // Reset input to allow re-selection
        }
    };

    const renderPdfToImage = async (pdfDoc, pageNum) => {
        const page = await pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale: 2.0 }); // High res for processing
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        // FIX: Fill white background first to prevent black transparency
        context.fillStyle = '#FFFFFF';
        context.fillRect(0, 0, canvas.width, canvas.height);

        await page.render({ canvasContext: context, viewport: viewport }).promise;
        return {
            dataUrl: canvas.toDataURL('image/jpeg', 0.8),
            width: viewport.width, // FIX: Use Full Width
            height: viewport.height
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
            // Filter out small words or symbols to find good anchors
            if (text.length > 3 && /^[a-zA-Z0-9]+$/.test(text) && w.confidence > 80) {
                if (!wordCounts[text]) {
                    wordCounts[text] = 0;
                    wordPositions[text] = { x: w.bbox.x0, y: w.bbox.y0 };
                }
                wordCounts[text]++;
            }
        });

        // Find words that appear exactly once (unique anchors)
        const candidates = Object.keys(wordCounts).filter(word => wordCounts[word] === 1);
        if (candidates.length === 0) return null;

        // Sort by position (prefer top-left)
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
        // Relaxed matching
        const match = words.find(w => w.text.toLowerCase().trim().includes(target));

        if (match) {
            return { x: match.bbox.x0, y: match.bbox.y0 };
        }
        return null;
    };

    // --- Main Processing: Load PDFs & Align ---
    const handleBatchLoadAndAlign = async () => {
        if (templateRegions.length === 0) {
            showToast("Please define at least one crop region.", "error");
            return;
        }

        setIsProcessing(true);
        setIsBatchMode(false);
        const newPages = [];
        const pdfjs = await loadPdfLib();

        let tesseractWorker = null;

        try {
            let anchorData = null;

            // Initialize Alignment
            setProgress({ current: 0, total: 0, status: "Initializing Alignment Engine..." });
            await loadTesseractLib();
            tesseractWorker = await window.Tesseract.createWorker('eng');

            // 1. Analyze Template to find Anchor
            setProgress({ current: 0, total: 0, status: "Analyzing template for anchor point..." });
            const page = await previewPdfDoc.getPage(previewPdfPage);
            const viewport = page.getViewport({ scale: 2.0 });
            const canvas = document.createElement('canvas');
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            const ctx = canvas.getContext('2d');
            
            // FIX: Fill white background for Template Analysis
            ctx.fillStyle = '#FFFFFF';
            ctx.fillRect(0, 0, canvas.width, canvas.height);

            await page.render({ canvasContext: ctx, viewport }).promise;

            anchorData = await findUniqueAnchor(tesseractWorker, canvas);

            if (!anchorData) {
                console.warn("Could not find a unique anchor word. Alignment disabled.");
            } else {
                console.log(`Auto-Alignment Anchor: "${anchorData.word}" at`, anchorData.pos);
            }

            // 2. Build Task List
            let tasks = [];

            if (batchPdfFiles.length === 1 && pagesPerExam > 0) {
                // Single PDF mode
                const file = batchPdfFiles[0];
                const arrayBuffer = await file.arrayBuffer();
                const doc = await pdfjs.getDocument(arrayBuffer).promise;

                for (let p = previewPdfPage; p <= doc.numPages; p += pagesPerExam) {
                    tasks.push({
                        doc,
                        pageNum: p,
                        originalName: `${file.name.replace('.pdf', '')}_P${p}`
                    });
                }
            } else {
                // Multi PDF mode
                for (let i = 0; i < batchPdfFiles.length; i++) {
                    const file = batchPdfFiles[i];
                    const arrayBuffer = await file.arrayBuffer();
                    const doc = await pdfjs.getDocument(arrayBuffer).promise;
                    // Assuming page 1 is the target, or user selected page
                    const targetP = doc.numPages >= previewPdfPage ? previewPdfPage : 1;
                    tasks.push({
                        doc,
                        pageNum: targetP,
                        originalName: file.name.replace('.pdf', '')
                    });
                }
            }

            // 3. Process each page (Render & Align)
            for (let i = 0; i < tasks.length; i++) {
                const task = tasks[i];
                setProgress({ current: i + 1, total: tasks.length, status: `Loading & Aligning ${task.originalName}...` });

                const page = await task.doc.getPage(task.pageNum);
                const viewport = page.getViewport({ scale: 2.0 });
                const canvas = document.createElement('canvas');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                const ctx = canvas.getContext('2d');
                
                // FIX: Fill white background for processing pages
                ctx.fillStyle = '#FFFFFF';
                ctx.fillRect(0, 0, canvas.width, canvas.height);

                await page.render({ canvasContext: ctx, viewport }).promise;

                // 3a. Calculate Offset based on Anchor
                let offsetX = 0;
                let offsetY = 0;

                if (anchorData && tesseractWorker) {
                    const targetPos = await findWordPosition(tesseractWorker, canvas, anchorData.word);
                    if (targetPos) {
                        offsetX = targetPos.x - anchorData.pos.x;
                        offsetY = targetPos.y - anchorData.pos.y;

                        // Safety check: ignore massive jumps which might be errors
                        if (Math.abs(offsetX) > canvas.width * 0.3 || Math.abs(offsetY) > canvas.height * 0.3) {
                            offsetX = 0; offsetY = 0;
                        }
                    }
                }

                // 4. Apply Offset to Regions
                // FIX: Use 1:1 coordinates because canvas is High Res
                const adjustedRegions = templateRegions.map(r => ({
                    ...r,
                    x: r.x + offsetX,
                    y: r.y + offsetY
                }));

                newPages.push({
                    id: Date.now() + i,
                    imageUrl: canvas.toDataURL('image/jpeg', 0.8),
                    width: viewport.width, // FIX: Use Full Width
                    height: viewport.height,
                    originalName: task.originalName,
                    regions: adjustedRegions,
                    crops: []
                });
            }

            setPages(newPages);
            setCurrentPageIndex(0);
            showToast(`Imported ${newPages.length} documents`, "success");
        } catch (e) {
            console.error(e);
            showToast("Error importing: " + e.message, "error");
        } finally {
            if (tesseractWorker) try { await tesseractWorker.terminate(); } catch (e) { }
            setIsProcessing(false);
            setProgress({ current: 0, total: 0, status: '' });
        }
    };

    // --- Cropping Logic ---
    const handleCropExtraction = async () => {
        if (pages.length === 0) return;
        setIsProcessing(true);
        const updatedPages = [...pages];

        try {
            let processedCount = 0;
            const totalOps = updatedPages.reduce((acc, p) => acc + p.regions.length, 0);

            for (let i = 0; i < updatedPages.length; i++) {
                const page = updatedPages[i];
                if (!page.regions || page.regions.length === 0) continue;

                const img = new Image();
                img.src = page.imageUrl;
                await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = reject; });
                
                page.crops = []; // Reset crops

                for (const region of page.regions) {
                    setProgress({ current: processedCount + 1, total: totalOps, status: `Cropping ${page.originalName}: ${region.label}...` });
                    
                    // Create a canvas for the crop
                    const cCanvas = document.createElement('canvas');
                    
                    // FIX: Ratio is now 1:1 because regions are in High Res coordinates
                    const cropX = region.x;
                    const cropY = region.y;
                    const cropW = region.width;
                    const cropH = region.height;

                    // Skip invalid crops
                    if (cropW <= 0 || cropH <= 0) continue;

                    cCanvas.width = cropW;
                    cCanvas.height = cropH;
                    
                    const ctx = cCanvas.getContext('2d');
                    // FIX: Draw directly
                    ctx.drawImage(img, cropX, cropY, cropW, cropH, 0, 0, cropW, cropH);

                    page.crops.push({
                        label: region.label,
                        dataUrl: cCanvas.toDataURL('image/jpeg', 0.9)
                    });
                    processedCount++;
                }
            }
            setPages(updatedPages);
            setShowResultsSidebar(true);
            showToast("Cropping Complete!", "success");
        } catch (e) {
            console.error(e);
            showToast("Crop failed: " + e.message, "error");
        } finally {
            setIsProcessing(false);
            setProgress({ current: 0, total: 0, status: '' });
        }
    };

    const handleDownloadZip = async () => {
        setIsProcessing(true);
        setProgress({ current: 0, total: 0, status: "Generating ZIP..." });
        try {
            const JSZip = await loadJsZipLib();
            const zip = new JSZip();
            
            let count = 0;
            pages.forEach(page => {
                if(page.crops && page.crops.length > 0) {
                    page.crops.forEach(crop => {
                        // Data URL format: "data:image/jpeg;base64,....."
                        const base64Data = crop.dataUrl.split(',')[1];
                        const fileName = `${page.originalName}_${crop.label}.jpg`;
                        zip.file(fileName, base64Data, {base64: true});
                        count++;
                    });
                }
            });

            if (count === 0) {
                showToast("No cropped images to download.", "error");
                return;
            }

            const content = await zip.generateAsync({type: "blob"});
            const url = window.URL.createObjectURL(content);
            const a = document.createElement('a');
            a.href = url;
            a.download = "cropped_images.zip";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            showToast("ZIP Downloaded!", "success");
        } catch (e) {
            console.error(e);
            showToast("Zip creation failed", "error");
        } finally {
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

        // Check if clicking existing region (Selection/Delete)
        const clickedIndex = currentRegions.findIndex(r =>
            pos.x >= r.x && pos.x <= r.x + r.width &&
            pos.y >= r.y && pos.y <= r.y + r.height
        );

        if (clickedIndex >= 0) {
            if (e.button === 2 || e.shiftKey) { // Right click or Shift+Click to delete
                const newRegions = [...currentRegions];
                newRegions.splice(clickedIndex, 1);
                // Renumber
                newRegions.forEach((r, idx) => r.label = `Region ${idx + 1}`);
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

        // Drawing new region
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
                    label: `Region ${currentRegions.length + 1}`
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
            ctx.strokeStyle = isSelected ? '#58a6ff' : '#d29922'; // Orange/Yellow for crops
            ctx.lineWidth = 2;
            ctx.fillStyle = isSelected ? 'rgba(88, 166, 255, 0.2)' : 'rgba(210, 153, 34, 0.1)';
            ctx.fillRect(r.x, r.y, r.width, r.height);
            ctx.strokeRect(r.x, r.y, r.width, r.height);

            // Label background
            ctx.fillStyle = isSelected ? '#58a6ff' : '#d29922';
            const labelW = ctx.measureText(r.label).width + 8;
            ctx.fillRect(r.x, r.y - 18, labelW, 18);
            
            // Label text
            ctx.fillStyle = '#000000';
            ctx.font = 'bold 11px monospace';
            ctx.textBaseline = 'middle';
            ctx.fillText(r.label, r.x + 4, r.y - 9);
        });

        if (ghostRect) {
            ctx.strokeStyle = '#f85149';
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

    // --- UI Theme ---
    const theme = { bg: 'bg-[#0d1117]', sidebar: 'bg-[#161b22]', border: 'border-[#30363d]', text: 'text-[#c9d1d9]', textMuted: 'text-[#8b949e]', accent: 'text-[#58a6ff]', accentBg: 'bg-[#1f6feb]', success: 'text-[#3fb950]', buttonHover: 'hover:bg-[#21262d]', inputBg: 'bg-[#0d1117]' };

    // --- Batch Wizard UI ---
    if (isBatchMode) {
        return (
            <div className={`h-screen w-screen ${theme.bg} ${theme.text} flex flex-col font-sans`}>
                 <style>{`:root, body, #root { height: 100%; width: 100%; margin: 0; padding: 0; max-width: none !important; }`}</style>
                <div className={`h-16 border-b ${theme.border} ${theme.sidebar} flex items-center justify-between px-6`}>
                    <h2 className={`font-bold flex items-center gap-2 ${theme.accent}`}><Files size={20} /> Batch Crop Wizard</h2>
                    <button onClick={() => setIsBatchMode(false)} className={`p-2 ${theme.buttonHover} rounded-md`}><X size={20} /></button>
                </div>
                <div className="flex-1 overflow-hidden flex">
                    {batchStep === 0 && (
                        <div className="flex-1 flex flex-col items-center justify-center p-10">
                            <div className={`border-2 border-dashed ${theme.border} rounded-xl ${theme.bg} p-12 flex flex-col items-center text-center max-w-lg w-full`}>
                                <FolderInput size={48} className={`${theme.textMuted} mb-4`} />
                                <h3 className="text-xl font-semibold mb-2">Select PDF File(s)</h3>
                                <p className={`${theme.textMuted} mb-6 text-sm`}>Supports multiple files.<br />Perfect for extracting signatures or specific fields.</p>
                                <input type="file" ref={folderInputRef} className="hidden" multiple accept="application/pdf" onChange={handleFileSelect} />
                                <button onClick={() => folderInputRef.current.click()} disabled={isPdfLoading} className={`px-6 py-3 ${theme.accentBg} hover:opacity-90 text-white rounded-md font-medium flex items-center gap-2`}>
                                    {isPdfLoading ? <Loader2 className="animate-spin" /> : <FolderInput size={18} />} Choose Files
                                </button>
                            </div>
                        </div>
                    )}
                    {batchStep === 1 && (
                        <div className="flex-1 flex overflow-hidden">
                            <div className={`w-80 ${theme.sidebar} border-r ${theme.border} flex flex-col p-4`}>
                                <h3 className="font-semibold mb-4 text-[#e6edf3]">Crop Template</h3>
                                <div className="mb-4">
                                    <label className={`text-xs ${theme.textMuted} uppercase font-bold`}>Page Selection</label>
                                    <div className="flex items-center gap-2 mt-1">
                                        <button onClick={() => setPreviewPdfPage(p => Math.max(1, p - 1))} className={`p-1 bg-[#21262d] rounded border ${theme.border}`}><ChevronLeft size={16} /></button>
                                        <span className="text-sm font-mono">{previewPdfPage} / {previewTotalPages}</span>
                                        <button onClick={() => setPreviewPdfPage(p => Math.min(previewTotalPages, p + 1))} className={`p-1 bg-[#21262d] rounded border ${theme.border}`}><ChevronRight size={16} /></button>
                                    </div>
                                </div>
                                <div className="mb-4 space-y-3">
                                    <p className={`text-[10px] ${theme.textMuted} flex items-center gap-1`}><Anchor size={12} /> Auto-Alignment Active</p>
                                    {batchPdfFiles.length === 1 && (
                                        <div>
                                            <label className={`text-xs ${theme.textMuted} uppercase font-bold mb-1 block`}>Pages Per Document</label>
                                            <input type="number" min="1" value={pagesPerExam} onChange={(e) => setPagesPerExam(parseInt(e.target.value) || 1)} className={`w-full ${theme.inputBg} border ${theme.border} rounded p-2 text-sm text-[#c9d1d9] focus:border-[#58a6ff] outline-none`} />
                                            <p className={`text-[10px] ${theme.textMuted} mt-1`}>Split single PDF into multiple docs every N pages.</p>
                                        </div>
                                    )}
                                </div>
                                <div className={`flex-1 overflow-y-auto mb-4 border ${theme.border} rounded-md ${theme.bg} p-2`}>
                                    <p className="text-[10px] text-gray-500 mb-2">Draw boxes on the canvas...</p>
                                    <div className="space-y-1">
                                        {templateRegions.map((r, i) => (
                                            <div key={i} className={`flex justify-between items-center text-xs p-2 bg-[#21262d] rounded border ${theme.border}`}>
                                                <span className={`font-mono text-[#d29922]`}>{r.label}</span>
                                                <button onClick={() => { const n = [...templateRegions]; n.splice(i, 1); setTemplateRegions(n); }} className="text-[#f85149]"><Trash2 size={12} /></button>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                                <button onClick={handleBatchLoadAndAlign} className={`w-full py-2 ${theme.accentBg} text-white rounded-md font-medium flex items-center justify-center gap-2`}><CheckCircle2 size={16} /> Import & Align</button>
                            </div>
                            <div className={`flex-1 ${theme.bg} overflow-auto flex items-center justify-center p-8 relative`}>
                                <div className={`absolute top-4 right-4 flex gap-1 ${theme.sidebar} p-1 rounded border ${theme.border} z-20 shadow-lg`}>
                                    <button onClick={() => setScale(s => Math.max(0.5, s - 0.1))} className={`p-1.5 ${theme.buttonHover} rounded ${theme.text}`}><ZoomOut size={16} /></button>
                                    <span className={`text-xs font-mono flex items-center px-2 ${theme.textMuted}`}>{Math.round(scale * 100)}%</span>
                                    <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className={`p-1.5 ${theme.buttonHover} rounded ${theme.text}`}><ZoomIn size={16} /></button>
                                </div>
                                <div className={`relative shadow-2xl border ${theme.border}`}>
                                    <canvas ref={templateCanvasRef} onMouseDown={(e) => handleCanvasMouseDown(e, true)} className="cursor-crosshair block" />
                                </div>
                            </div>
                        </div>
                    )}
                </div>
            </div>
        );
    }

    // --- Main UI ---
    return (
        <div className={`flex h-screen w-full ${theme.bg} ${theme.text} font-sans overflow-hidden`}>
            <style>{`:root, body, #root { height: 100%; width: 100%; margin: 0; padding: 0; max-width: none !important; }`}</style>
            
            {/* Left Sidebar: Document List */}
            <div className={`w-64 ${theme.sidebar} border-r ${theme.border} flex flex-col flex-shrink-0`}>
                <div className={`p-4 border-b ${theme.border} flex items-center gap-2`}><ScanLine className="text-[#d29922]" /> <span className="font-bold text-sm">PDF Batch Cropper</span></div>
                <div className={`p-3 border-b ${theme.border}`}>
                    <button onClick={() => { setIsBatchMode(true); setBatchStep(0); }} className={`w-full py-2 border border-dashed ${theme.border} rounded-md ${theme.buttonHover} ${theme.accent} text-xs font-medium flex items-center justify-center gap-2 transition-colors`}><FolderInput size={16} /> New Batch</button>
                </div>
                <div className="flex-1 overflow-y-auto p-2 space-y-2">
                    {pages.map((page, idx) => (
                        <div key={page.id} onClick={() => setCurrentPageIndex(idx)} className={`p-2 rounded-md border cursor-pointer flex items-center gap-3 group transition-all ${currentPageIndex === idx ? `bg-[#21262d] border-[#58a6ff]` : `${theme.bg} ${theme.border}`}`}>
                            {/* Show first crop as thumbnail if available, else generic icon */}
                            <div className={`w-10 h-10 bg-[#0d1117] rounded flex items-center justify-center ${theme.textMuted} overflow-hidden flex-shrink-0 border ${theme.border}`}>
                                {page.crops && page.crops.length > 0 ? (
                                    <img src={page.crops[0].dataUrl} className="w-full h-full object-contain" />
                                ) : (
                                    <Files size={16} />
                                )}
                            </div>
                            <div className="flex-1 min-w-0">
                                <div className="text-xs font-medium text-[#e6edf3] truncate">{page.originalName}</div>
                                <div className={`text-[10px] ${theme.textMuted}`}>
                                    {page.crops ? page.crops.length : 0} Crops ready
                                </div>
                            </div>
                        </div>
                    ))}
                </div>
                {pages.length > 0 && (
                    <div className={`p-3 border-t ${theme.border}`}>
                        <button onClick={handleCropExtraction} disabled={isProcessing} className="w-full py-2 bg-[#d29922] hover:opacity-90 text-black disabled:opacity-50 rounded-md font-bold flex items-center justify-center gap-2 text-sm shadow-md">
                            {isProcessing ? <Loader2 className="animate-spin w-4 h-4" /> : <><Scissors size={16} /> Crop All Images</>}
                        </button>
                    </div>
                )}
            </div>

            {/* Main Center Area: Canvas */}
            <div className="flex-1 flex flex-col relative overflow-hidden bg-[#010409]">
                <div className={`h-12 border-b ${theme.border} ${theme.sidebar} flex items-center justify-between px-4`}>
                    <div className="flex items-center gap-2">
                        <div className={`flex items-center gap-1 bg-[#21262d] rounded p-0.5 border ${theme.border}`}>
                            <button onClick={() => handlePageNavigation('prev')} className={`p-1 ${theme.buttonHover} rounded ${theme.text}`}><ChevronLeft size={14} /></button>
                            <input className={`w-8 bg-transparent text-center text-xs font-mono focus:outline-none ${theme.text}`} value={pageInput} onChange={(e) => { setPageInput(e.target.value); const val = parseInt(e.target.value); if (val > 0 && val <= pages.length) setCurrentPageIndex(val - 1); }} />
                            <button onClick={() => handlePageNavigation('next')} className={`p-1 ${theme.buttonHover} rounded ${theme.text}`}><ChevronRight size={14} /></button>
                        </div>
                    </div>
                    <div className="flex items-center gap-2">
                        <div className={`h-6 w-px bg-[#30363d] mx-2`}></div>
                        <button onClick={() => setScale(s => Math.max(0.5, s - 0.1))} className={`p-1.5 ${theme.buttonHover} rounded text-[#7d8590] hover:text-[#c9d1d9]`}><ZoomOut size={18} /></button>
                        <span className={`text-xs font-mono w-10 text-center ${theme.textMuted}`}>{Math.round(scale * 100)}%</span>
                        <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className={`p-1.5 ${theme.buttonHover} rounded text-[#7d8590] hover:text-[#c9d1d9]`}><ZoomIn size={18} /></button>
                        <div className={`h-6 w-px bg-[#30363d] mx-2`}></div>
                        <button onClick={handleDownloadZip} className={`p-1.5 ${theme.success} ${theme.buttonHover} rounded-md`} title="Download ZIP"><Download size={18} /></button>
                        <button onClick={() => setShowResultsSidebar(!showResultsSidebar)} className={`p-1.5 ${theme.accent} ${theme.buttonHover} rounded-md`}><LayoutTemplate size={18} /></button>
                    </div>
                </div>
                <div className="flex-1 overflow-auto flex items-center justify-center p-8 relative">
                    <div className={`${theme.bg} shadow-2xl relative border ${theme.border}`}>
                        <canvas ref={canvasRef} onMouseDown={(e) => handleCanvasMouseDown(e, false)} className="cursor-crosshair block" />
                        {pages.length === 0 && (
                            <div className={`absolute inset-0 flex flex-col items-center justify-center ${theme.textMuted} pointer-events-none`}><ScanLine size={48} className="mb-4 opacity-20" /><p>Start a New Batch</p></div>
                        )}
                    </div>
                </div>
                {/* Processing Overlay */}
                {isProcessing && (
                    <div className={`absolute inset-0 ${theme.bg}/80 z-50 flex items-center justify-center backdrop-blur-sm`}>
                        <div className={`${theme.sidebar} border ${theme.border} p-6 rounded-xl shadow-2xl max-w-sm w-full text-center`}>
                            <Loader2 size={32} className={`animate-spin ${theme.accent} mx-auto mb-4`} />
                            <h3 className="text-[#e6edf3] font-medium mb-1">{progress.status}</h3>
                            <div className={`w-full ${theme.border} h-1.5 rounded-full overflow-hidden mt-3 bg-gray-700`}>
                                <div className="h-full bg-[#d29922] transition-all duration-300" style={{ width: `${(progress.current / progress.total) * 100}%` }}></div>
                            </div>
                            <p className={`text-xs ${theme.textMuted} mt-2`}>{progress.current} / {progress.total} items</p>
                        </div>
                    </div>
                )}
            </div>

            {/* Right Sidebar: Crop Results Preview */}
            {showResultsSidebar && (
                <div className={`w-72 ${theme.sidebar} border-l ${theme.border} flex flex-col flex-shrink-0 z-20`}>
                    <div className={`p-4 border-b ${theme.border} flex justify-between items-center`}>
                        <h3 className="font-bold text-[#e6edf3] flex items-center gap-2"><ImageIcon size={18} /> Crop Results</h3>
                    </div>
                    <div className="flex-1 overflow-y-auto p-4 space-y-4">
                        {pages[currentPageIndex]?.crops && pages[currentPageIndex].crops.length > 0 ? (
                            pages[currentPageIndex].crops.map((crop, idx) => (
                                <div key={idx} className={`space-y-2`}>
                                    <div className={`text-xs ${theme.textMuted} font-mono mb-1`}>{crop.label}</div>
                                    <div className={`bg-[url('data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI4IiBoZWlnaHQ9IjgiPgo8cmVjdCB3aWR0aD0iOCIgaGVpZ2h0PSI4IiBmaWxsPSIjMGMxMTE3Ii8+CjxyZWN0IHdpZHRoPSI0IiBoZWlnaHQ9IjQiIGZpbGw9IiMxNjFiMjIiLz4KPHJlY3QgeD0iNCIgeT0iNCIgd2lkdGg9IjQiIGhlaWdodD0iNCIgZmlsbD0iIzE2MWIyMiIvPgo8L3N2Zz4=')] border ${theme.border} rounded-md overflow-hidden`}>
                                        <img src={crop.dataUrl} className="w-full h-auto block" alt={crop.label} />
                                    </div>
                                </div>
                            ))
                        ) : (
                            <div className={`text-center ${theme.textMuted} py-10 text-xs flex flex-col items-center gap-2`}>
                                <Layers size={24} className="opacity-20" />
                                <p>No crops yet.<br/>Click "Crop All Images" to extract.</p>
                            </div>
                        )}
                    </div>
                    {pages.some(p => p.crops && p.crops.length > 0) && (
                         <div className={`p-4 border-t ${theme.border}`}>
                             <button onClick={handleDownloadZip} className={`w-full py-2 ${theme.success} border border-[#238636] hover:bg-[#238636]/10 rounded-md font-medium text-xs flex items-center justify-center gap-2`}>
                                 <Download size={14} /> Download ZIP
                             </button>
                         </div>
                    )}
                </div>
            )}
            
            {/* Toast Notification */}
            {toast && (<div className={`fixed bottom-6 right-6 px-4 py-3 rounded-md shadow-lg text-white text-xs font-medium z-50 flex items-center gap-3 animate-fade-in-up ${toast.type === 'error' ? 'bg-[#da3633]' : 'bg-[#238636]'}`}>{toast.type === 'error' ? <X size={14} /> : <CheckCircle2 size={14} />}{toast.message}</div>)}
        </div>
    );
};

export default App;