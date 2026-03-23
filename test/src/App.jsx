import React, { useState, useRef, useEffect } from 'react';
import {
    Trash2, ChevronRight, ChevronLeft, 
    ZoomIn, ZoomOut, X, Loader2, FolderInput, 
    Files, Scissors, Download, FilePlus, ScanLine, AlertCircle, Filter, Hash, FileUp, GitBranch, FileSpreadsheet
} from 'lucide-react';

// --- External Libraries ---
const PDF_LIB_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
const PDF_WORKER_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const TESSERACT_LIB_URL = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
const JSZIP_LIB_URL = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
const PDF_LIB_MANIP_URL = "https://unpkg.com/pdf-lib/dist/pdf-lib.min.js";

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

const loadPdfManipLib = async () => {
    if (window.PDFLib) return window.PDFLib;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = PDF_LIB_MANIP_URL;
        script.onload = () => resolve(window.PDFLib);
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

// --- Helper Functions ---

// Clean up canvas to free memory
const cleanupCanvas = (canvas) => {
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    canvas.width = 0;
    canvas.height = 0;
    canvas.remove();
};

// Render PDF page to canvas
const renderPdfToCanvas = async (pdfDoc, pageNum, renderScale = 2.0) => {
    try {
        const page = await pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale: renderScale });
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        const renderTask = page.render({ canvasContext: ctx, viewport });
        await renderTask.promise;
        
        page.cleanup();
        
        return { canvas, width: viewport.width, height: viewport.height };
    } catch (e) {
        console.error("Render error:", e);
        throw e;
    }
};

// OCR a page and extract word data for anchor detection
const ocrPageForAnchors = async (worker, canvas) => {
    const base64 = canvas.toDataURL('image/jpeg', 0.5);
    const { data: { words } } = await worker.recognize(base64);
    
    const counts = {};
    const positions = {};
    
    words.forEach(w => {
        const text = w.text.trim().toLowerCase();
        if (text.length > 4 && w.confidence > 80) {
            counts[text] = (counts[text] || 0) + 1;
            if (!positions[text]) {
                positions[text] = { x: w.bbox.x0, y: w.bbox.y0 };
            }
        }
    });
    
    return { words, counts, positions };
};

// Find universal anchor across all documents
const findUniversalAnchor = (perDocWordData) => {
    if (perDocWordData.length === 0) return null;
    
    // Get unique words from each document
    const uniqueSets = perDocWordData.map(doc => {
        const uniques = new Set();
        for (const [word, count] of Object.entries(doc.counts)) {
            if (count === 1) uniques.add(word);
        }
        return uniques;
    });
    
    // Find intersection - words that are unique in ALL documents
    const firstSet = uniqueSets[0];
    const universalUniques = uniqueSets.slice(1).reduce(
        (intersection, set) => {
            return new Set([...intersection].filter(word => set.has(word)));
        },
        firstSet
    );
    
    if (universalUniques.size === 0) return null;
    
    // Score candidates by position variance
    const candidates = [];
    for (const word of universalUniques) {
        const positions = perDocWordData.map(doc => doc.positions[word]);
        const confidences = perDocWordData.map(doc => {
            const match = doc.words.find(w => w.text.trim().toLowerCase() === word);
            return match ? match.confidence : 0;
        });
        
        // Calculate position variance (lower is better)
        const avgX = positions.reduce((sum, p) => sum + p.x, 0) / positions.length;
        const avgY = positions.reduce((sum, p) => sum + p.y, 0) / positions.length;
        const variance = positions.reduce((sum, p) => 
            sum + Math.pow(p.x - avgX, 2) + Math.pow(p.y - avgY, 2), 0
        ) / positions.length;
        
        // Average confidence
        const avgConfidence = confidences.reduce((sum, c) => sum + c, 0) / confidences.length;
        
        candidates.push({
            word,
            positions,
            variance,
            avgConfidence,
            avgY: positions[0].y  // For sorting
        });
    }
    
    // Sort by: lowest variance first, then topmost (by y)
    candidates.sort((a, b) => {
        if (a.variance !== b.variance) return a.variance - b.variance;
        return a.avgY - b.avgY;
    });
    
    const best = candidates[0];
    
    return {
        word: best.word,
        positions: best.positions,
        variance: best.variance,
        avgConfidence: best.avgConfidence
    };
};

// Parse class numbers from filter input
const parseClassNumbers = (input) => {
    if (!input || !input.trim()) return null;
    const numbers = new Set();
    const parts = input.split(/[\s,]+/).filter(p => p.trim());
    parts.forEach(p => {
        if (p.includes('-')) {
            const [start, end] = p.split('-').map(Number);
            if (!isNaN(start) && !isNaN(end)) {
                for (let i = Math.min(start, end); i <= Math.max(start, end); i++) numbers.add(i);
            }
        } else {
            const n = parseInt(p);
            if (!isNaN(n)) numbers.add(n);
        }
    });
    return numbers;
};

const App = () => {
    const [scale, setScale] = useState(0.8);
    const [isProcessing, setIsProcessing] = useState(false);
    const [toast, setToast] = useState(null);

    // Batch Wizard State
    const [isBatchMode, setIsBatchMode] = useState(false);
    const [batchStep, setBatchStep] = useState(0); 
    const [batchPdfFiles, setBatchPdfFiles] = useState([]);
    const processedFullPdfsRef = useRef([]); 
    const [previewPdfDoc, setPreviewPdfDoc] = useState(null);
    const [previewPdfPage, setPreviewPdfPage] = useState(1);
    const [previewTotalPages, setPreviewTotalPages] = useState(1);
    const [templateRegions, setTemplateRegions] = useState([]); 
    const [progress, setProgress] = useState({ current: 0, total: 100, message: '' });
    
    // File loading state for progressive UI reveal
    const [isLoadingFiles, setIsLoadingFiles] = useState(false);
    const [loadingProgress, setLoadingProgress] = useState({ current: 0, total: 0, fileName: '' });

    const templateCanvasRef = useRef(null);
    const folderInputRef = useRef(null);

    const theme = { bg: 'bg-[#0d1117]', sidebar: 'bg-[#161b22]', border: 'border-[#30363d]', text: 'text-[#c9d1d9]', accent: 'text-[#58a6ff]', accentBg: 'bg-[#1f6feb]' };

    useEffect(() => {
        if (isBatchMode && batchStep === 1 && previewPdfDoc) {
            drawTemplateCanvas();
        }
    }, [batchStep, templateRegions, previewPdfPage, scale, isBatchMode, previewPdfDoc]);

    const showToast = (message) => {
        setToast(message);
        setTimeout(() => setToast(null), 3000);
    };

    const handleFileSelect = async (e) => {
        const files = Array.from(e.target.files);
        if (!files.length) return;

        // Immediately show batch mode UI with loading state
        setIsBatchMode(true);
        setBatchStep(1);
        setIsLoadingFiles(true);
        setLoadingProgress({ current: 0, total: files.length, fileName: '' });

        try {
            const pdfjs = await loadPdfLib();
            const { PDFDocument, rgb, StandardFonts } = await loadPdfManipLib();
            
            let preliminaryDocs = [];
            let maxPages = 0;

            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                setLoadingProgress(prev => ({ ...prev, current: i + 1, fileName: file.name }));
                const buffer = await file.arrayBuffer();
                const doc = await pdfjs.getDocument({ data: new Uint8Array(buffer.slice(0)) }).promise;
                if (doc.numPages > maxPages) maxPages = doc.numPages;
                preliminaryDocs.push({ file, pageCount: doc.numPages, buffer });
            }

            const finalDocs = [];
            const fullPdfsToDownload = []; 

            for (const item of preliminaryDocs) {
                const pdfDoc = await PDFDocument.load(item.buffer.slice(0));
                let finalBytes;
                
                if (item.pageCount < maxPages) {
                    const pagesToAdd = maxPages - item.pageCount;
                    const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
                    const pages = pdfDoc.getPages();
                    const width = pages.length > 0 ? pages[0].getWidth() : 595.28;
                    const height = pages.length > 0 ? pages[0].getHeight() : 841.89;

                    for (let j = 0; j < pagesToAdd; j++) {
                        const newPage = pdfDoc.addPage([width, height]);
                        newPage.drawText(`This is a blank page`, {
                            x: width / 2 - 100,
                            y: height / 2,
                            size: 24,
                            font: font,
                            color: rgb(0.85, 0.85, 0.85),
                            opacity: 0.6,
                        });
                    }
                    finalBytes = await pdfDoc.save();
                } else {
                    finalBytes = await pdfDoc.save();
                }
                
            fullPdfsToDownload.push({ name: item.file.name, bytes: finalBytes });
            finalDocs.push({ bytes: new Uint8Array(finalBytes), name: item.file.name, pageCount: maxPages });
            
            if (finalDocs.length === 1) {
                const previewDoc = await pdfjs.getDocument({ data: new Uint8Array(finalBytes) }).promise;
                setPreviewPdfDoc(previewDoc);
            }
            
            item.buffer = null;
        }

        setBatchPdfFiles(finalDocs);
        processedFullPdfsRef.current = fullPdfsToDownload;
        setPreviewTotalPages(maxPages);
            showToast(`Loaded ${files.length} files.`);
        } catch (err) {
            console.error(err);
            showToast("Failed to process PDF files.");
            setIsBatchMode(false);
            setBatchStep(0);
        } finally {
            setIsLoadingFiles(false);
            setLoadingProgress({ current: 0, total: 0, fileName: '' });
        }
    };

    const downloadProcessedPdfsZip = async () => {
        if (processedFullPdfsRef.current.length === 0) return;
        setIsProcessing(true);
        try {
            const JSZip = await loadJsZipLib();
            const zip = new JSZip();
            processedFullPdfsRef.current.forEach(f => {
                zip.file(f.name, f.bytes);
            });
            const blob = await zip.generateAsync({type: "blob"});
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = "processed_pdfs.zip";
            link.click();
            setTimeout(() => URL.revokeObjectURL(url), 1000);
            processedFullPdfsRef.current = [];
            showToast("ZIP downloaded successfully.");
        } catch {
            showToast("Zip generation failed.");
        } finally {
            setIsProcessing(false);
        }
    };

    const handleBatchProcessAndExport = async () => {
        if (templateRegions.length === 0) return showToast("Define at least one region.");
        
        setIsProcessing(true);
        let worker = null;
        
        try {
            const pdfjs = await loadPdfLib();
            await loadTesseractLib();
            const JSZip = await loadJsZipLib();
            const { PDFDocument } = await loadPdfManipLib();
            const zip = new JSZip();
            
            worker = await window.Tesseract.createWorker('eng');
            
            // ===== Calculate total work units for unified progress =====
            const pagesWithRegions = [...new Set(templateRegions.map(r => r.page))].sort((a,b) => a-b);
            const anchorDetectionWork = pagesWithRegions.length * batchPdfFiles.length;
            const cropProcessingWork = batchPdfFiles.length;
            const exportWork = templateRegions.length;
            const totalWork = anchorDetectionWork + cropProcessingWork + exportWork;
            
            let completedWork = 0;
            const updateUnifiedProgress = (message) => {
                completedWork++;
                const pct = Math.round((completedWork / totalWork) * 100);
                setProgress({ current: pct, total: 100, message });
            };
            
            setProgress({ current: 0, total: 100, message: 'Starting anchor detection...' });
            
            // ===== PHASE 2: Universal Anchor Detection =====
            const anchorMap = {};
            
            for (const pageNum of pagesWithRegions) {
                const perDocWordData = [];
                
                for (let docIndex = 0; docIndex < batchPdfFiles.length; docIndex++) {
                    const item = batchPdfFiles[docIndex];
                    
                    const pdfDoc = await window.pdfjsLib.getDocument({ data: new Uint8Array(item.bytes) }).promise;
                    const render = await renderPdfToCanvas(pdfDoc, pageNum, 1.5);
                    
                    const wordData = await ocrPageForAnchors(worker, render.canvas);
                    perDocWordData.push({
                        docIndex,
                        words: wordData.words,
                        counts: wordData.counts,
                        positions: wordData.positions
                    });
                    
                    cleanupCanvas(render.canvas);
                    pdfDoc.destroy();
                    
                    updateUnifiedProgress(`Scanning page ${pageNum}, document ${docIndex + 1}/${batchPdfFiles.length}`);
                }
                
                const anchor = findUniversalAnchor(perDocWordData);
                
                if (anchor) {
                    anchorMap[pageNum] = {
                        word: anchor.word,
                        positions: anchor.positions,
                        variance: anchor.variance,
                        avgConfidence: anchor.avgConfidence
                    };
                    console.log(`Page ${pageNum}: Found anchor "${anchor.word}" (variance: ${anchor.variance.toFixed(2)}, confidence: ${anchor.avgConfidence.toFixed(1)}%)`);
                } else {
                    console.warn(`Page ${pageNum}: No universal anchor found`);
                    anchorMap[pageNum] = null;
                }
                
                perDocWordData.length = 0;
            }
            
            // Report anchor detection results
            const anchorCount = Object.values(anchorMap).filter(a => a !== null).length;
            showToast(`Found universal anchors for ${anchorCount}/${pagesWithRegions.length} pages`);
            
            // ===== PHASE 3: Batch Crop Processing =====
            const croppedImagesByRegion = {};
            for (let i = 0; i < templateRegions.length; i++) {
                croppedImagesByRegion[i] = [];
            }

            for (let docIndex = 0; docIndex < batchPdfFiles.length; docIndex++) {
                const item = batchPdfFiles[docIndex];
                const currentDoc = await pdfjs.getDocument({ data: new Uint8Array(item.bytes) }).promise;
                
                const baseName = item.name.replace(/\.[^/.]+$/, ""); 
                const lastTwoDigitsMatch = baseName.match(/(\d{2})$/);
                const classNumber = lastTwoDigitsMatch ? parseInt(lastTwoDigitsMatch[1]) : null;

                for (const pNum of pagesWithRegions) {
                    const regionsOnPage = templateRegions
                        .map((r, idx) => ({ ...r, idx }))
                        .filter(r => r.page === pNum);
                    
                    const applicableRegions = regionsOnPage.filter(r => {
                        const allowedSet = parseClassNumbers(r.filter);
                        if (!allowedSet) return true;
                        return classNumber !== null && allowedSet.has(classNumber);
                    });

                    if (applicableRegions.length === 0) continue;

                    const render = await renderPdfToCanvas(currentDoc, pNum);
                    let offX = 0, offY = 0;
                    
                    if (anchorMap[pNum]) {
                        const templatePos = anchorMap[pNum].positions[0];
                        const currentPos = anchorMap[pNum].positions[docIndex];
                        if (templatePos && currentPos) {
                            offX = (currentPos.x - templatePos.x);
                            offY = (currentPos.y - templatePos.y);
                        }
                    }

                    for (const r of applicableRegions) {
                        const cropX = (r.x * 2) + offX;
                        const cropY = (r.y * 2) + offY;
                        const cropW = r.width * 2;
                        const cropH = r.height * 2;

                        const c = document.createElement('canvas');
                        c.width = Math.max(1, cropW);
                        c.height = Math.max(1, cropH);
                        const ctx = c.getContext('2d');
                        ctx.drawImage(render.canvas, cropX, cropY, cropW, cropH, 0, 0, cropW, cropH);
                        
                        const base64Data = c.toDataURL('image/jpeg', 0.9);
                        
                        cleanupCanvas(c);
                        
                        croppedImagesByRegion[r.idx].push({
                            base64: base64Data,
                            width: cropW * 0.5,
                            height: cropH * 0.5
                        });
                    }
                    
                    cleanupCanvas(render.canvas);
                }
                
                currentDoc.destroy();
                await new Promise(resolve => setTimeout(resolve, 10));
                
                updateUnifiedProgress(`Cropping document ${docIndex + 1}/${batchPdfFiles.length}`);
            }
            
            // ===== PHASE 4: Export =====
            for (let idx = 0; idx < templateRegions.length; idx++) {
                const pdfDoc = await PDFDocument.create();
                
                for (const imgData of croppedImagesByRegion[idx]) {
                    const image = await pdfDoc.embedJpg(imgData.base64);
                    const page = pdfDoc.addPage([imgData.width, imgData.height]);
                    page.drawImage(image, {
                        x: 0,
                        y: 0,
                        width: imgData.width,
                        height: imgData.height
                    });
                }
                
                const pdfBytes = await pdfDoc.save();
                const fileName = `Q${idx + 1}_${templateRegions[idx].label}.pdf`;
                zip.file(fileName, pdfBytes);
                
                croppedImagesByRegion[idx] = null;
                
                updateUnifiedProgress(`Building PDF ${idx + 1}/${templateRegions.length}`);
            }
            
            const blob = await zip.generateAsync({type: "blob"});
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = "batch_questions_pdfs.zip";
            link.click();
            setTimeout(() => URL.revokeObjectURL(url), 1000);

            showToast("Success! ZIP downloaded.");
            setIsBatchMode(false);
            setBatchStep(0);
        } catch (err) {
            console.error("Processing failed:", err);
            showToast("Processing failed: " + err.message);
        } finally {
            if (worker) await worker.terminate();
            setIsProcessing(false);
            setProgress({ current: 0, total: 0, message: '' });
        }
    };

    const drawTemplateCanvas = async (ghost = null) => {
        const canvas = templateCanvasRef.current; 
        if (!canvas || !previewPdfDoc) return;
        
        const ctx = canvas.getContext('2d');
        const { canvas: pdfCanvas } = await renderPdfToCanvas(previewPdfDoc, previewPdfPage, 1.0);
        
        canvas.width = pdfCanvas.width * scale; 
        canvas.height = pdfCanvas.height * scale;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.save();
        ctx.scale(scale, scale);
        ctx.drawImage(pdfCanvas, 0, 0);
        
        templateRegions.filter(r => r.page === previewPdfPage).forEach((r) => {
            const regionIndex = templateRegions.indexOf(r);
            ctx.strokeStyle = '#d29922'; 
            ctx.lineWidth = 3 / scale;
            ctx.strokeRect(r.x, r.y, r.width, r.height);
            ctx.fillStyle = '#d29922';
            ctx.font = `bold ${12 / scale}px sans-serif`;
            ctx.fillText(`Q${regionIndex + 1}: ${r.label}`, r.x, r.y - (5 / scale));
        });
        
        if (ghost) { 
            ctx.strokeStyle = '#58a6ff'; 
            ctx.setLineDash([5 / scale, 5 / scale]); 
            ctx.lineWidth = 2 / scale;
            ctx.strokeRect(ghost.x, ghost.y, ghost.w, ghost.h); 
        }
        ctx.restore();
        
        cleanupCanvas(pdfCanvas);
    };

    const startDrag = (e) => {
        const canvas = templateCanvasRef.current;
        const rect = canvas.getBoundingClientRect();
        
        const startX = (e.clientX - rect.left) / scale;
        const startY = (e.clientY - rect.top) / scale;

        const onMove = (me) => {
            const curX = (me.clientX - rect.left) / scale;
            const curY = (me.clientY - rect.top) / scale;
            drawTemplateCanvas({ 
                x: Math.min(startX, curX), 
                y: Math.min(startY, curY), 
                w: Math.abs(curX - startX), 
                h: Math.abs(curY - startY) 
            });
        };

        const onUp = (ue) => {
            const ex = (ue.clientX - rect.left) / scale;
            const ey = (ue.clientY - rect.top) / scale;
            const w = Math.abs(ex - startX);
            const h = Math.abs(ey - startY);
            if (w > 10 && h > 10) {
                setTemplateRegions(prev => [...prev, { 
                    page: previewPdfPage, 
                    x: Math.min(startX, ex), 
                    y: Math.min(startY, ey), 
                    width: w, 
                    height: h, 
                    label: `Region_${prev.length + 1}`,
                    filter: '' 
                }]);
            }
            window.removeEventListener('mousemove', onMove);
            window.removeEventListener('mouseup', onUp);
            drawTemplateCanvas();
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
    };

    // Loading canvas component with fade in/out animation
    const LoadingCanvas = () => {
        const canvasRef = useRef(null);
        
        useEffect(() => {
            const canvas = canvasRef.current;
            if (!canvas) return;
            
            const ctx = canvas.getContext('2d');
            const width = 595;
            const height = 842;
            canvas.width = width;
            canvas.height = height;
            
            let opacity = 0;
            let increasing = true;
            let animationId;
            
            const draw = () => {
                ctx.clearRect(0, 0, width, height);
                ctx.fillStyle = '#FFFFFF';
                ctx.fillRect(0, 0, width, height);
                
                // Draw loading template pattern
                ctx.globalAlpha = 0.1;
                ctx.strokeStyle = '#58a6ff';
                ctx.lineWidth = 1;
                
                // Draw grid pattern
                for (let x = 0; x < width; x += 40) {
                    ctx.beginPath();
                    ctx.moveTo(x, 0);
                    ctx.lineTo(x, height);
                    ctx.stroke();
                }
                for (let y = 0; y < height; y += 40) {
                    ctx.beginPath();
                    ctx.moveTo(0, y);
                    ctx.lineTo(width, y);
                    ctx.stroke();
                }
                
                ctx.globalAlpha = 1;
                
                // Draw loading text with fade animation
                ctx.globalAlpha = opacity;
                ctx.fillStyle = '#58a6ff';
                ctx.font = 'bold 24px sans-serif';
                ctx.textAlign = 'center';
                ctx.fillText('Loading PDF...', width / 2, height / 2 - 20);
                
                ctx.font = '14px sans-serif';
                ctx.fillStyle = '#8b949e';
                if (loadingProgress.fileName) {
                    ctx.fillText(`Processing: ${loadingProgress.fileName}`, width / 2, height / 2 + 20);
                    ctx.fillText(`${loadingProgress.current} / ${loadingProgress.total} files`, width / 2, height / 2 + 45);
                }
                
                ctx.globalAlpha = 1;
                
                // Animate opacity
                if (increasing) {
                    opacity += 0.02;
                    if (opacity >= 1) {
                        opacity = 1;
                        increasing = false;
                    }
                } else {
                    opacity -= 0.02;
                    if (opacity <= 0.3) {
                        opacity = 0.3;
                        increasing = true;
                    }
                }
                
                animationId = requestAnimationFrame(draw);
            };
            
            draw();
            
            return () => {
                if (animationId) {
                    cancelAnimationFrame(animationId);
                }
            };
        }, [loadingProgress]);
        
        return (
            <canvas 
                ref={canvasRef} 
                className="shadow-2xl border border-[#30363d] bg-white transition-all"
                style={{ width: 595 * scale, height: 842 * scale }}
            />
        );
    };

    if (isBatchMode && batchStep === 1) {
        return (
            <div className={`h-screen w-screen ${theme.bg} ${theme.text} flex flex-col font-sans`}>
                <div className={`h-14 border-b ${theme.border} ${theme.sidebar} flex items-center justify-between px-6`}>
                    <div className="flex items-center gap-2 font-bold text-[#58a6ff]"><FilePlus size={20}/> Batch Workflow</div>
                    <div className="flex items-center gap-3">
                        <a href="/beta/" target="_self" rel="noopener noreferrer" className="p-2 rounded-lg hover:bg-blue-500/10 transition-all duration-200" title="Beta Version"><FileSpreadsheet size={24} className="text-[#58a6ff] drop-shadow-md" /></a>
                        <a href="/" target="_self" rel="noopener noreferrer" className="p-2 rounded-lg hover:bg-blue-500/10 transition-all duration-200" title="Main Version"><GitBranch size={24} className="text-[#58a6ff] drop-shadow-md" /></a>
                        <button onClick={() => { setIsBatchMode(false); setBatchStep(0); }} className="hover:text-white transition-colors"><X/></button>
                    </div>
                </div>
                <div className="flex-1 flex overflow-hidden">
                    <div className={`w-80 ${theme.sidebar} border-r ${theme.border} p-5 flex flex-col`}>
                        <div className={`mb-6 ${isLoadingFiles ? 'opacity-50 pointer-events-none' : ''}`}>
                            <h3 className="font-bold mb-2 text-sm text-gray-400">Editor Controls</h3>
                            <div className="grid grid-cols-2 gap-2 mb-4">
                                <button onClick={() => setScale(s => Math.max(0.1, s - 0.1))} className="flex items-center justify-center gap-2 py-2 bg-[#21262d] border border-[#30363d] rounded text-[10px] hover:bg-[#30363d]"><ZoomOut size={12}/> Zoom Out</button>
                                <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className="flex items-center justify-center gap-2 py-2 bg-[#21262d] border border-[#30363d] rounded text-[10px] hover:bg-[#30363d]"><ZoomIn size={12}/> Zoom In</button>
                            </div>
                            <div className="flex items-center justify-between bg-black/30 p-2 rounded">
                                <button onClick={() => setPreviewPdfPage(p => Math.max(1, p-1))} className="p-1 hover:bg-white/10 rounded"><ChevronLeft size={20}/></button>
                                <span className="text-xs font-mono">Page {previewPdfPage} / {previewTotalPages || '?'}</span>
                                <button onClick={() => setPreviewPdfPage(p => Math.min(previewTotalPages, p+1))} className="p-1 hover:bg-white/10 rounded"><ChevronRight size={20}/></button>
                            </div>
                        </div>

                        {/* Loading indicator for sidebar */}
                        {isLoadingFiles && (
                            <div className="mb-6 p-3 bg-[#1f6feb]/10 border border-[#1f6feb]/30 rounded-lg animate-pulse">
                                <div className="flex items-center gap-2 mb-2">
                                    <Loader2 className="animate-spin text-[#58a6ff]" size={16} />
                                    <h3 className="text-[10px] font-bold text-[#58a6ff] uppercase tracking-wider">Loading Files...</h3>
                                </div>
                                <p className="text-[10px] text-gray-400 mb-2 leading-tight">
                                    {loadingProgress.fileName || 'Preparing...'}
                                </p>
                                <div className="w-full bg-[#30363d] rounded-full h-1.5">
                                    <div 
                                        className="bg-[#58a6ff] h-1.5 rounded-full transition-all duration-300" 
                                        style={{ width: `${loadingProgress.total > 0 ? (loadingProgress.current / loadingProgress.total) * 100 : 0}%` }}
                                    ></div>
                                </div>
                                <p className="text-[9px] text-gray-500 mt-1 text-right">
                                    {loadingProgress.current} / {loadingProgress.total} files
                                </p>
                            </div>
                        )}

                        <div className={`${isLoadingFiles ? 'opacity-50 pointer-events-none blur-[2px]' : ''}`}>
                            {processedFullPdfsRef.current.length > 0 && (
                                <div className="mb-6 p-3 bg-[#1f6feb]/10 border border-[#1f6feb]/30 rounded-lg">
                                    <h3 className="text-[10px] font-bold text-[#58a6ff] uppercase tracking-wider mb-2">Padded Documents</h3>
                                    <p className="text-[10px] text-gray-400 mb-3 leading-tight">All documents have been unified to {previewTotalPages} pages. You can download the full processed files below.</p>
                                    <button 
                                        onClick={downloadProcessedPdfsZip} 
                                        className="w-full py-2 bg-[#238636] hover:bg-[#2ea043] rounded flex items-center justify-center gap-2 text-[11px] font-bold text-white transition-colors"
                                    >
                                        <Download size={12}/> Download Processed PDFs (ZIP)
                                    </button>
                                </div>
                            )}

                            <h3 className="font-bold mb-2 text-sm text-gray-400">Target Regions</h3>
                            <div className="flex-1 overflow-auto space-y-3 mb-4 pr-1 scrollbar-thin">
                                {templateRegions.length === 0 && <div className="text-[10px] text-gray-500 italic p-6 text-center border border-dashed border-[#30363d] rounded">Drag boxes on the document to define crop areas</div>}
                                {templateRegions.map((r, i) => (
                                    <div key={i} className="p-3 bg-[#0d1117] border border-[#30363d] rounded-lg space-y-2 group">
                                        <div className="flex justify-between items-center">
                                            <span className="text-[10px] font-bold text-[#58a6ff]">Q{i + 1}: {r.label} (Page {r.page})</span>
                                            <button onClick={() => setTemplateRegions(templateRegions.filter((_, idx) => idx !== i))} className="text-red-500 hover:scale-110 opacity-50 group-hover:opacity-100 transition-opacity"><Trash2 size={12}/></button>
                                        </div>
                                        <div className="relative">
                                            <Hash className="absolute left-2 top-2 text-gray-500" size={12}/>
                                            <input 
                                                type="text"
                                                placeholder="Student ID filter (e.g. 01, 05, 10-15)"
                                                value={r.filter}
                                                onChange={(e) => {
                                                    const value = e.target.value;
                                                    setTemplateRegions(prev => prev.map((region, idx) => idx === i ? { ...region, filter: value } : region));
                                                }}
                                                className="w-full bg-black/40 border border-[#30363d] rounded pl-7 pr-2 py-1.5 text-[10px] focus:outline-none focus:border-[#58a6ff]"
                                            />
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </div>

                        <button 
                            onClick={handleBatchProcessAndExport} 
                            className={`w-full py-4 ${isLoadingFiles ? 'bg-gray-600 cursor-not-allowed' : theme.accentBg} rounded-lg font-bold flex flex-col items-center justify-center gap-1 shadow-lg active:scale-95 transition-transform ${isLoadingFiles ? 'opacity-50' : ''}`}
                            disabled={isLoadingFiles}
                        >
                            <div className="flex items-center gap-2">
                                {isLoadingFiles ? <Loader2 className="animate-spin" size={18} /> : <Scissors size={18}/>} 
                                Process & Export Crops
                            </div>
                            <span className="text-[9px] font-normal opacity-70 italic text-center">Groups crops into Q1.pdf, Q2.pdf, etc.</span>
                        </button>
                    </div>
                    <div className="flex-1 bg-[#010409] overflow-auto flex items-start p-10 bg-[radial-gradient(#30363d_1px,transparent_1px)] bg-[size:20px_20px]">
                        <div className="flex-shrink-0 relative mx-auto">
                            {isLoadingFiles || !previewPdfDoc ? (
                                <LoadingCanvas />
                            ) : (
                                <canvas 
                                    ref={templateCanvasRef} 
                                    onMouseDown={startDrag} 
                                    className="shadow-2xl border border-[#30363d] bg-white cursor-crosshair transition-all" 
                                />
                            )}
                        </div>
                    </div>
                </div>
                {isProcessing && (
                    <div className="absolute inset-0 bg-black/85 flex flex-col items-center justify-center z-50 animate-in fade-in duration-300">
                        <div className="relative">
                            <Loader2 className="animate-spin text-[#58a6ff] mb-6" size={64} />
                        </div>
                        <h2 className="text-2xl font-bold mb-2 tracking-tight">{progress.message || 'Processing Documents'}</h2>
                        {progress.total > 0 && (
                            <div className="w-64 bg-[#30363d] rounded-full h-2 mt-4">
                                <div 
                                    className="bg-[#58a6ff] h-2 rounded-full transition-all duration-300" 
                                    style={{ width: `${Math.round((progress.current / progress.total) * 100)}%` }}
                                ></div>
                            </div>
                        )}
                        {progress.total > 0 && (
                            <p className="text-sm text-gray-400 mt-2">{Math.round((progress.current / progress.total) * 100)}% complete</p>
                        )}
                    </div>
                )}
            </div>
        );
    }

    return (
        <div className={`h-screen w-screen ${theme.bg} ${theme.text} flex items-center justify-center font-sans`}>
            <div className="text-center max-w-md p-12 bg-[#161b22] border border-[#30363d] rounded-3xl shadow-2xl">
                <div className="bg-[#58a6ff]/10 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-8">
                    <ScanLine size={40} className="text-[#58a6ff]" />
                </div>
                <h1 className="text-3xl font-bold mb-4 tracking-tight">PDF Multi-Cropper</h1>
                <p className="text-sm text-gray-400 mb-10 leading-relaxed">
                    Export your crops directly into categorized PDFs. Ideal for grouping specific sections from multiple source documents into a single file.
                </p>
                <input 
                    type="file" 
                    ref={folderInputRef} 
                    hidden 
                    multiple 
                    accept=".pdf" 
                    onChange={handleFileSelect} 
                />
                <button 
                    onClick={() => folderInputRef.current.click()} 
                    className={`w-full py-4 ${theme.accentBg} rounded-xl font-bold text-lg flex items-center justify-center gap-3 hover:scale-[1.03] transition-all hover:shadow-[0_0_20px_rgba(88,166,255,0.3)] active:scale-95`}
                >
                    <FilePlus size={24}/> Create New Batch
                </button>
            </div>
            {isProcessing && (
                <div className="absolute inset-0 bg-black/85 flex flex-col items-center justify-center z-50">
                    <Loader2 className="animate-spin text-[#58a6ff] mb-4" size={48} />
                    <p className="text-white font-medium">{progress.message || 'Loading PDFs...'}</p>
                </div>
            )}
            {toast && <div className="fixed bottom-8 right-8 px-6 py-3 bg-[#238636] text-white text-sm font-bold rounded-xl shadow-2xl animate-in slide-in-from-bottom-4 duration-300">{toast}</div>}
        </div>
    );
};

export default App;
