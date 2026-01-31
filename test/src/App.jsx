import React, { useState, useRef, useEffect } from 'react';
import {
    Trash2, ChevronRight, ChevronLeft, 
    ZoomIn, ZoomOut, X, Loader2, FolderInput, 
    Files, Scissors, Download, FilePlus, ScanLine, AlertCircle
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

const App = () => {
    const [pages, setPages] = useState([]); 
    const [currentPageIndex, setCurrentPageIndex] = useState(0);
    const [scale, setScale] = useState(1);
    const [isProcessing, setIsProcessing] = useState(false);
    const [progress, setProgress] = useState({ current: 0, total: 0, status: '' });
    const [toast, setToast] = useState(null);

    // Batch Wizard State
    const [isBatchMode, setIsBatchMode] = useState(false);
    const [batchStep, setBatchStep] = useState(0); 
    const [batchPdfFiles, setBatchPdfFiles] = useState([]);
    const [previewPdfDoc, setPreviewPdfDoc] = useState(null);
    const [previewPdfPage, setPreviewPdfPage] = useState(1);
    const [previewTotalPages, setPreviewTotalPages] = useState(1);
    const [templateRegions, setTemplateRegions] = useState([]); 
    const [needsDownloadPadding, setNeedsDownloadPadding] = useState([]);

    const canvasRef = useRef(null);
    const templateCanvasRef = useRef(null);
    const folderInputRef = useRef(null);

    const theme = { bg: 'bg-[#0d1117]', sidebar: 'bg-[#161b22]', border: 'border-[#30363d]', text: 'text-[#c9d1d9]', accent: 'text-[#58a6ff]', accentBg: 'bg-[#1f6feb]' };

    useEffect(() => {
        if (!isBatchMode && pages.length > 0 && pages[currentPageIndex]) drawMainCanvas();
    }, [pages, currentPageIndex, scale, isBatchMode]);

    useEffect(() => {
        if (isBatchMode && batchStep === 1 && previewPdfDoc) drawTemplateCanvas();
    }, [batchStep, templateRegions, previewPdfPage, isBatchMode]);

    const showToast = (message) => {
        setToast(message);
        setTimeout(() => setToast(null), 3000);
    };

    const renderPdfToDataUrl = async (pdfDoc, pageNum, renderScale = 2.0) => {
        const page = await pdfDoc.getPage(pageNum);
        const viewport = page.getViewport({ scale: renderScale });
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({ canvasContext: ctx, viewport }).promise;
        return { 
            dataUrl: canvas.toDataURL('image/jpeg', 0.8), 
            width: viewport.width, 
            height: viewport.height,
            canvas 
        };
    };

    const handleFileSelect = async (e) => {
        const files = Array.from(e.target.files);
        if (!files.length) return;
        
        setIsProcessing(true);
        const pdfjs = await loadPdfLib();
        const { PDFDocument, rgb } = await loadPdfManipLib();
        
        let preliminaryDocs = [];
        let maxPages = 0;
        let paddedFilesForDownload = [];

        try {
            setProgress({ current: 0, total: files.length, status: 'Analyzing document lengths...' });
            for (let i = 0; i < files.length; i++) {
                const buffer = await files[i].arrayBuffer();
                const doc = await pdfjs.getDocument({ data: new Uint8Array(buffer) }).promise;
                if (doc.numPages > maxPages) maxPages = doc.numPages;
                preliminaryDocs.push({ file: files[i], pageCount: doc.numPages });
            }

            const finalDocs = [];
            for (let i = 0; i < preliminaryDocs.length; i++) {
                const item = preliminaryDocs[i];
                const buffer = await item.file.arrayBuffer();
                let pdfBytes;

                if (item.pageCount < maxPages) {
                    const pdfDoc = await PDFDocument.load(buffer);
                    const pagesToAdd = maxPages - item.pageCount;
                    const firstPage = pdfDoc.getPages()[0];
                    const width = firstPage ? firstPage.getWidth() : 595.28;
                    const height = firstPage ? firstPage.getHeight() : 841.89;

                    for (let j = 0; j < pagesToAdd; j++) {
                        const newPage = pdfDoc.addPage([width, height]);
                        newPage.drawText('This is a blank page', {
                            x: width / 2 - 60,
                            y: height / 2,
                            size: 12,
                            color: rgb(0.7, 0.7, 0.7)
                        });
                    }
                    pdfBytes = await pdfDoc.save();
                    // Requirement: No "fixed_" prefix, keep original filename
                    paddedFilesForDownload.push({ name: item.file.name, data: pdfBytes });
                } else {
                    pdfBytes = buffer;
                }
                
                const dataCopy = new Uint8Array(pdfBytes);
                const readyDoc = await pdfjs.getDocument({ data: dataCopy }).promise;
                finalDocs.push({ 
                    doc: readyDoc, 
                    name: item.file.name.replace('.pdf', ''), 
                    pageCount: maxPages 
                });
            }

            setBatchPdfFiles(finalDocs);
            setPreviewPdfDoc(finalDocs[0].doc);
            setPreviewTotalPages(maxPages);
            setNeedsDownloadPadding(paddedFilesForDownload);
            setBatchStep(1);
        } catch (err) {
            console.error(err);
            showToast("Failed to process local PDF files.");
        } finally {
            setIsProcessing(false);
        }
    };

    const downloadFixedPdfs = async () => {
        const JSZip = await loadJsZipLib();
        const zip = new JSZip();
        needsDownloadPadding.forEach(f => zip.file(f.name, f.data));
        const blob = await zip.generateAsync({type: "blob"});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = "aligned_pdfs.zip";
        link.click();
        showToast("Downloaded updated PDFs. Save these to your computer.");
        setNeedsDownloadPadding([]);
    };

    const handleBatchProcess = async () => {
        if (templateRegions.length === 0) return showToast("Define at least one region.");
        setIsProcessing(true);
        const newPages = [];
        const pagesToExtract = [...new Set(templateRegions.map(r => r.page))].sort((a,b) => a-b);
        let tesseractWorker = null;

        try {
            await loadTesseractLib();
            tesseractWorker = await window.Tesseract.createWorker('eng');
            const anchorMap = {}; 
            for (const pNum of pagesToExtract) {
                const { canvas } = await renderPdfToDataUrl(previewPdfDoc, pNum);
                anchorMap[pNum] = await findUniqueAnchor(tesseractWorker, canvas);
            }

            for (const item of batchPdfFiles) {
                for (const pNum of pagesToExtract) {
                    const render = await renderPdfToDataUrl(item.doc, pNum);
                    let offX = 0, offY = 0;
                    if (anchorMap[pNum]) {
                        const pos = await findWordPosition(tesseractWorker, render.canvas, anchorMap[pNum].word);
                        if (pos) { offX = pos.x - anchorMap[pNum].pos.x; offY = pos.y - anchorMap[pNum].pos.y; }
                    }
                    newPages.push({
                        id: crypto.randomUUID(),
                        imageUrl: render.dataUrl,
                        width: render.width, height: render.height,
                        originalName: `${item.name}_P${pNum}`,
                        regions: templateRegions.filter(r => r.page === pNum).map(r => ({...r, x: r.x + offX, y: r.y + offY})),
                        crops: []
                    });
                }
            }
            setPages(newPages);
            setIsBatchMode(false);
        } catch (e) {
            showToast("Alignment engine error.");
        } finally {
            if (tesseractWorker) await tesseractWorker.terminate();
            setIsProcessing(false);
        }
    };

    const findUniqueAnchor = async (worker, canvas) => {
        const { data: { words } } = await worker.recognize(canvas);
        const counts = {}, positions = {};
        words.forEach(w => {
            const t = w.text.trim().toLowerCase();
            if (t.length > 4 && w.confidence > 85) { counts[t] = (counts[t] || 0) + 1; positions[t] = { x: w.bbox.x0, y: w.bbox.y0 }; }
        });
        const cand = Object.keys(counts).filter(w => counts[w] === 1).sort((a,b) => positions[a].y - positions[b].y);
        return cand.length ? { word: cand[0], pos: positions[cand[0]] } : null;
    };

    const findWordPosition = async (worker, canvas, target) => {
        const { data: { words } } = await worker.recognize(canvas);
        const match = words.find(w => w.text.toLowerCase().includes(target.toLowerCase()));
        return match ? { x: match.bbox.x0, y: match.bbox.y0 } : null;
    };

    const handleCropExtraction = async () => {
        setIsProcessing(true);
        const updated = [...pages];
        for (const page of updated) {
            const img = new Image(); img.src = page.imageUrl;
            await new Promise(res => img.onload = res);
            page.crops = page.regions.map(r => {
                const c = document.createElement('canvas'); c.width = r.width; c.height = r.height;
                const ctx = c.getContext('2d');
                ctx.drawImage(img, r.x, r.y, r.width, r.height, 0, 0, r.width, r.height);
                return { label: r.label, dataUrl: c.toDataURL('image/jpeg', 0.9) };
            });
        }
        setPages(updated);
        setIsProcessing(false);
    };

    const drawMainCanvas = async () => {
        const canvas = canvasRef.current; if (!canvas || !pages[currentPageIndex]) return;
        const ctx = canvas.getContext('2d'), img = new Image(); img.src = pages[currentPageIndex].imageUrl;
        await new Promise(res => img.onload = res);
        canvas.width = img.width * scale; canvas.height = img.height * scale;
        ctx.scale(scale, scale); ctx.drawImage(img, 0, 0);
        pages[currentPageIndex].regions.forEach(r => {
            ctx.strokeStyle = '#d29922'; ctx.lineWidth = 3; ctx.strokeRect(r.x, r.y, r.width, r.height);
            ctx.fillStyle = '#d29922'; ctx.font = 'bold 12px sans-serif'; ctx.fillText(r.label, r.x, r.y - 5);
        });
    };

    const drawTemplateCanvas = async (ghost = null) => {
        const canvas = templateCanvasRef.current; if (!canvas) return;
        const ctx = canvas.getContext('2d'), p = await renderPdfToDataUrl(previewPdfDoc, previewPdfPage);
        const img = new Image(); img.src = p.dataUrl;
        await new Promise(res => img.onload = res);
        canvas.width = img.width * scale; canvas.height = img.height * scale;
        ctx.scale(scale, scale); ctx.drawImage(img, 0, 0);
        templateRegions.filter(r => r.page === previewPdfPage).forEach(r => {
            ctx.strokeStyle = '#d29922'; ctx.lineWidth = 3; ctx.strokeRect(r.x, r.y, r.width, r.height);
        });
        if (ghost) { ctx.strokeStyle = '#58a6ff'; ctx.setLineDash([5, 5]); ctx.strokeRect(ghost.x, ghost.y, ghost.w, ghost.h); }
    };

    const startDrag = (e, isTemplate) => {
        const canvas = isTemplate ? templateCanvasRef.current : canvasRef.current;
        const rect = canvas.getBoundingClientRect();
        const startX = (e.clientX - rect.left) / scale, startY = (e.clientY - rect.top) / scale;
        const onMove = (me) => {
            const curX = (me.clientX - rect.left) / scale, curY = (me.clientY - rect.top) / scale;
            if (isTemplate) drawTemplateCanvas({ x: Math.min(startX, curX), y: Math.min(startY, curY), w: Math.abs(curX - startX), h: Math.abs(curY - startY) });
        };
        const onUp = (ue) => {
            const ex = (ue.clientX - rect.left) / scale, ey = (ue.clientY - rect.top) / scale;
            const w = Math.abs(ex - startX), h = Math.abs(ey - startY);
            if (w > 20 && h > 20 && isTemplate) setTemplateRegions([...templateRegions, { page: previewPdfPage, x: Math.min(startX, ex), y: Math.min(startY, ey), width: w, height: h, label: `Reg_${templateRegions.length+1}` }]);
            window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp);
        };
        window.addEventListener('mousemove', onMove); window.addEventListener('mouseup', onUp);
    };

    const handleDownloadZip = async () => {
        const JSZip = await loadJsZipLib(); const zip = new JSZip();
        pages.forEach(p => p.crops.forEach(c => zip.file(`${p.originalName}_${c.label}.jpg`, c.dataUrl.split(',')[1], {base64: true})));
        const blob = await zip.generateAsync({type: "blob"});
        const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = "crops.zip"; link.click();
    };

    if (isBatchMode) {
        return (
            <div className={`h-screen w-screen ${theme.bg} ${theme.text} flex flex-col font-sans`}>
                <div className={`h-14 border-b ${theme.border} ${theme.sidebar} flex items-center justify-between px-6`}>
                    <div className="flex items-center gap-2 font-bold text-[#58a6ff]"><FilePlus size={20}/> Batch Workflow</div>
                    <button onClick={() => setIsBatchMode(false)}><X/></button>
                </div>
                {batchStep === 0 ? (
                    <div className="flex-1 flex items-center justify-center">
                        <div className={`border-2 border-dashed ${theme.border} rounded-xl p-16 text-center max-w-md`}>
                            <FolderInput size={48} className="mx-auto mb-4 text-gray-500" />
                            <h3 className="text-xl font-bold mb-2">Select PDFs</h3>
                            <p className="text-sm text-gray-400 mb-6 italic">Browser security prevents direct file modification. We will provide updated versions to save back to your drive.</p>
                            <input type="file" ref={folderInputRef} hidden multiple accept=".pdf" onChange={handleFileSelect} />
                            <button onClick={() => folderInputRef.current.click()} className={`px-10 py-3 ${theme.accentBg} rounded-lg font-bold`}>Choose Files</button>
                        </div>
                    </div>
                ) : (
                    <div className="flex-1 flex overflow-hidden">
                        <div className={`w-80 ${theme.sidebar} border-r ${theme.border} p-5 flex flex-col`}>
                            {needsDownloadPadding.length > 0 && (
                                <div className="mb-6 p-4 bg-amber-500/10 border border-amber-500/30 rounded-lg">
                                    <div className="flex items-center gap-2 text-amber-400 font-bold mb-2 text-xs"><AlertCircle size={14}/> Action Required</div>
                                    <p className="text-[10px] text-gray-400 mb-3">Some files were shorter than others. Download the padded versions to keep your local drive in sync.</p>
                                    <button onClick={downloadFixedPdfs} className="w-full py-2 bg-amber-600 text-white text-[10px] font-bold rounded flex items-center justify-center gap-2"><Download size={12}/> Download Fixed PDFs</button>
                                </div>
                            )}
                            <h3 className="font-bold mb-4">Crop Regions</h3>
                            <div className="flex items-center justify-between mb-4 bg-black/20 p-2 rounded">
                                <button onClick={() => setPreviewPdfPage(p => Math.max(1, p-1))}><ChevronLeft/></button>
                                <span className="text-xs font-mono">Page {previewPdfPage} / {previewTotalPages}</span>
                                <button onClick={() => setPreviewPdfPage(p => Math.min(previewTotalPages, p+1))}><ChevronRight/></button>
                            </div>
                            <div className="flex-1 overflow-auto space-y-2 mb-4">
                                {templateRegions.map((r, i) => (
                                    <div key={i} className="text-[10px] p-2 bg-[#0d1117] border border-[#30363d] rounded flex justify-between items-center">
                                        <span>{r.label} (P{r.page})</span>
                                        <button onClick={() => setTemplateRegions(templateRegions.filter((_, idx) => idx !== i))} className="text-red-500"><Trash2 size={12}/></button>
                                    </div>
                                ))}
                            </div>
                            <button onClick={handleBatchProcess} className={`w-full py-3 ${theme.accentBg} rounded-lg font-bold flex items-center justify-center gap-2`}><Scissors size={18}/> Process All</button>
                        </div>
                        <div className="flex-1 bg-[#010409] overflow-auto flex items-center justify-center p-10">
                            <canvas ref={templateCanvasRef} onMouseDown={(e) => startDrag(e, true)} className="shadow-2xl border border-[#30363d]" />
                        </div>
                    </div>
                )}
                {isProcessing && <div className="absolute inset-0 bg-black/70 flex items-center justify-center z-50"><Loader2 className="animate-spin text-[#d29922]" /></div>}
            </div>
        );
    }

    return (
        <div className={`h-screen w-screen ${theme.bg} ${theme.text} flex overflow-hidden font-sans`}>
            <div className={`w-64 ${theme.sidebar} border-r ${theme.border} flex flex-col`}>
                <div className="p-4 border-b border-[#30363d] font-bold text-sm">PDF Crop Manager</div>
                <div className="p-2">
                    <button onClick={() => { setIsBatchMode(true); setBatchStep(0); }} className="w-full py-2 border border-dashed border-[#30363d] rounded text-[10px] text-gray-500 hover:text-white">+ START BATCH</button>
                </div>
                <div className="flex-1 overflow-auto p-2">
                    {pages.map((p, i) => (
                        <div key={p.id} onClick={() => setCurrentPageIndex(i)} className={`p-2 rounded cursor-pointer text-[10px] truncate ${currentPageIndex === i ? 'bg-[#21262d] text-white border border-[#58a6ff]' : 'text-gray-500 border border-transparent'}`}>
                            {p.originalName}
                        </div>
                    ))}
                </div>
                {pages.length > 0 && <div className="p-4 border-t border-[#30363d]"><button onClick={handleCropExtraction} className="w-full py-2 bg-[#d29922] text-black font-bold rounded text-xs">EXTRACT CROPS</button></div>}
            </div>
            <div className="flex-1 flex flex-col bg-[#010409]">
                <div className="h-12 border-b border-[#30363d] bg-[#161b22] flex items-center justify-between px-4 text-[10px]">
                    <div className="flex items-center gap-4">
                        <button onClick={() => setCurrentPageIndex(p => Math.max(0, p-1))}><ChevronLeft size={16}/></button>
                        <span>{currentPageIndex + 1} / {pages.length}</span>
                        <button onClick={() => setCurrentPageIndex(p => Math.min(pages.length-1, p+1))}><ChevronRight size={16}/></button>
                    </div>
                    <div className="flex items-center gap-4">
                        <button onClick={() => setScale(s => Math.max(0.2, s-0.1))}><ZoomOut size={16}/></button>
                        <span>{Math.round(scale*100)}%</span>
                        <button onClick={() => setScale(s => Math.min(4, s+0.1))}><ZoomIn size={16}/></button>
                        <button onClick={handleDownloadZip} className="text-[#3fb950] font-bold ml-4"><Download size={14}/> ZIP</button>
                    </div>
                </div>
                <div className="flex-1 overflow-auto flex items-center justify-center p-10"><canvas ref={canvasRef} className="shadow-2xl border border-[#30363d]" /></div>
            </div>
            <div className={`w-72 ${theme.sidebar} border-l border-[#30363d] p-4 flex flex-col overflow-auto gap-4`}>
                <div className="font-bold text-xs text-gray-500 uppercase">Preview</div>
                {pages[currentPageIndex]?.crops.map((c, i) => (
                    <div key={i} className="space-y-1"><div className="text-[9px] text-gray-500">{c.label}</div><img src={c.dataUrl} className="w-full border border-[#30363d] rounded" /></div>
                ))}
            </div>
            {toast && <div className="fixed bottom-6 right-6 px-4 py-2 bg-[#238636] text-white text-xs font-bold rounded shadow-xl">{toast}</div>}
        </div>
    );
};

export default App;