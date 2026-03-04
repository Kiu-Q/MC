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

const App = () => {
    const [scale, setScale] = useState(0.8);
    const [isProcessing, setIsProcessing] = useState(false);
    const [toast, setToast] = useState(null);

    // Batch Wizard State
    const [isBatchMode, setIsBatchMode] = useState(false);
    const [batchStep, setBatchStep] = useState(0); 
    const [batchPdfFiles, setBatchPdfFiles] = useState([]);
    const [processedFullPdfs, setProcessedFullPdfs] = useState([]); // Store full PDFs with added pages
    const [previewPdfDoc, setPreviewPdfDoc] = useState(null);
    const [previewPdfPage, setPreviewPdfPage] = useState(1);
    const [previewTotalPages, setPreviewTotalPages] = useState(1);
    const [templateRegions, setTemplateRegions] = useState([]); 

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
            await page.render({ canvasContext: ctx, viewport }).promise;
            return { canvas, width: viewport.width, height: viewport.height };
        } catch (e) {
            console.error("Render error:", e);
            throw e;
        }
    };

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

    const updateRegionFilter = (index, value) => {
        setTemplateRegions(prev => prev.map((r, i) => i === index ? { ...r, filter: value } : r));
    };

    const handleFileSelect = async (e) => {
        const files = Array.from(e.target.files);
        if (!files.length) return;

        setIsProcessing(true);
        try {
            const pdfjs = await loadPdfLib();
            const { PDFDocument, rgb, StandardFonts } = await loadPdfManipLib();
            
            let preliminaryDocs = [];
            let maxPages = 0;

            for (const file of files) {
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
                const readyDoc = await pdfjs.getDocument({ data: new Uint8Array(finalBytes) }).promise;
                finalDocs.push({ doc: readyDoc, name: item.file.name, pageCount: maxPages });
            }

            setBatchPdfFiles(finalDocs);
            setProcessedFullPdfs(fullPdfsToDownload);
            setPreviewPdfDoc(finalDocs[0].doc);
            setPreviewTotalPages(maxPages);
            setIsBatchMode(true);
            setBatchStep(1);
            showToast(`Loaded ${files.length} files.`);
        } catch (err) {
            console.error(err);
            showToast("Failed to process PDF files.");
        } finally {
            setIsProcessing(false);
        }
    };

    const downloadProcessedPdfsZip = async () => {
        if (processedFullPdfs.length === 0) return;
        setIsProcessing(true);
        try {
            const JSZip = await loadJsZipLib();
            const zip = new JSZip();
            processedFullPdfs.forEach(f => {
                zip.file(f.name, f.bytes);
            });
            const blob = await zip.generateAsync({type: "blob"});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = "processed_pdfs.zip";
            link.click();
            showToast("ZIP downloaded successfully.");
        } catch (e) {
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
            await loadTesseractLib();
            const JSZip = await loadJsZipLib();
            const { PDFDocument } = await loadPdfManipLib();
            const zip = new JSZip();
            
            worker = await window.Tesseract.createWorker('eng');
            
            const pagesToExtract = [...new Set(templateRegions.map(r => r.page))].sort((a,b) => a-b);
            const anchorMap = {}; 

            const regionPdfs = {};
            for (let i = 0; i < templateRegions.length; i++) {
                regionPdfs[i] = await PDFDocument.create();
            }

            for (const pNum of pagesToExtract) {
                const { canvas } = await renderPdfToCanvas(previewPdfDoc, pNum);
                anchorMap[pNum] = await findUniqueAnchor(worker, canvas);
            }

            for (const item of batchPdfFiles) {
                const baseName = item.name.replace(/\.[^/.]+$/, ""); 
                const lastTwoDigitsMatch = baseName.match(/(\d{2})$/);
                const classNumber = lastTwoDigitsMatch ? parseInt(lastTwoDigitsMatch[1]) : null;

                for (const pNum of pagesToExtract) {
                    const regionsOnPageIndices = templateRegions
                        .map((r, idx) => ({ ...r, idx }))
                        .filter(r => r.page === pNum);
                    
                    const applicableRegions = regionsOnPageIndices.filter(r => {
                        const allowedSet = parseClassNumbers(r.filter);
                        if (!allowedSet) return true;
                        return classNumber !== null && allowedSet.has(classNumber);
                    });

                    if (applicableRegions.length === 0) continue;

                    const render = await renderPdfToCanvas(item.doc, pNum);
                    let offX = 0, offY = 0;
                    
                    if (anchorMap[pNum]) {
                        const pos = await findWordPosition(worker, render.canvas, anchorMap[pNum].word);
                        if (pos) {
                            offX = (pos.x - anchorMap[pNum].pos.x);
                            offY = (pos.y - anchorMap[pNum].pos.y);
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
                        
                        const imgBytes = await new Promise(resolve => {
                            c.toBlob(blob => blob.arrayBuffer().then(resolve), 'image/jpeg', 0.9);
                        });

                        const pdfDoc = regionPdfs[r.idx];
                        const image = await pdfDoc.embedJpg(imgBytes);
                        
                        const page = pdfDoc.addPage([cropW * 0.5, cropH * 0.5]);
                        page.drawImage(image, {
                            x: 0,
                            y: 0,
                            width: cropW * 0.5,
                            height: cropH * 0.5,
                        });
                    }
                }
            }
            
            for (let i = 0; i < templateRegions.length; i++) {
                const pdfBytes = await regionPdfs[i].save();
                const fileName = `Q${i + 1}_${templateRegions[i].label}.pdf`;
                zip.file(fileName, pdfBytes);
            }
            
            const blob = await zip.generateAsync({type: "blob"});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = "batch_questions_pdfs.zip";
            link.click();

            showToast("Success! ZIP downloaded.");
            setIsBatchMode(false);
            setBatchStep(0);
        } catch (e) {
            console.error(e);
            showToast("Processing failed.");
        } finally {
            if (worker) await worker.terminate();
            setIsProcessing(false);
        }
    };

    const findUniqueAnchor = async (worker, canvas) => {
        const { data: { words } } = await worker.recognize(canvas);
        const counts = {}, positions = {};
        words.forEach(w => {
            const t = w.text.trim().toLowerCase();
            if (t.length > 4 && w.confidence > 80) { 
                counts[t] = (counts[t] || 0) + 1; 
                positions[t] = { x: w.bbox.x0, y: w.bbox.y0 }; 
            }
        });
        const cand = Object.keys(counts).filter(w => counts[w] === 1).sort((a,b) => positions[a].y - positions[b].y);
        return cand.length ? { word: cand[0], pos: positions[cand[0]] } : null;
    };

    const findWordPosition = async (worker, canvas, target) => {
        const { data: { words } } = await worker.recognize(canvas);
        const match = words.find(w => w.text.toLowerCase().includes(target.toLowerCase()));
        return match ? { x: match.bbox.x0, y: match.bbox.y0 } : null;
    };

    const drawTemplateCanvas = async (ghost = null) => {
        const canvas = templateCanvasRef.current; 
        if (!canvas || !previewPdfDoc) return;
        
        const ctx = canvas.getContext('2d');
        const { canvas: pdfCanvas } = await renderPdfToCanvas(previewPdfDoc, previewPdfPage, 1.0);
        
        // Logical size of the canvas at current scale
        canvas.width = pdfCanvas.width * scale; 
        canvas.height = pdfCanvas.height * scale;
        
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.save();
        ctx.scale(scale, scale);
        ctx.drawImage(pdfCanvas, 0, 0);
        
        templateRegions.filter(r => r.page === previewPdfPage).forEach((r, i) => {
            const regionIndex = templateRegions.indexOf(r);
            ctx.strokeStyle = '#d29922'; 
            ctx.lineWidth = 3 / scale; // Keep stroke thickness visually consistent
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
    };

    const startDrag = (e) => {
        const canvas = templateCanvasRef.current;
        const rect = canvas.getBoundingClientRect();
        
        // Calculate coordinate relative to the unscaled PDF dimensions
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
            drawTemplateCanvas(); // Final redraw to clear ghost
        };
        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
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
                        <div className="mb-6">
                            <h3 className="font-bold mb-2 text-sm text-gray-400">Editor Controls</h3>
                            <div className="grid grid-cols-2 gap-2 mb-4">
                                <button onClick={() => setScale(s => Math.max(0.1, s - 0.1))} className="flex items-center justify-center gap-2 py-2 bg-[#21262d] border border-[#30363d] rounded text-[10px] hover:bg-[#30363d]"><ZoomOut size={12}/> Zoom Out</button>
                                <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className="flex items-center justify-center gap-2 py-2 bg-[#21262d] border border-[#30363d] rounded text-[10px] hover:bg-[#30363d]"><ZoomIn size={12}/> Zoom In</button>
                            </div>
                            <div className="flex items-center justify-between bg-black/30 p-2 rounded">
                                <button onClick={() => setPreviewPdfPage(p => Math.max(1, p-1))} className="p-1 hover:bg-white/10 rounded"><ChevronLeft size={20}/></button>
                                <span className="text-xs font-mono">Page {previewPdfPage} / {previewTotalPages}</span>
                                <button onClick={() => setPreviewPdfPage(p => Math.min(previewTotalPages, p+1))} className="p-1 hover:bg-white/10 rounded"><ChevronRight size={20}/></button>
                            </div>
                        </div>

                        {processedFullPdfs.length > 0 && (
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
                                            onChange={(e) => updateRegionFilter(i, e.target.value)}
                                            className="w-full bg-black/40 border border-[#30363d] rounded pl-7 pr-2 py-1.5 text-[10px] focus:outline-none focus:border-[#58a6ff]"
                                        />
                                    </div>
                                </div>
                            ))}
                        </div>

                        <button onClick={handleBatchProcessAndExport} className={`w-full py-4 ${theme.accentBg} rounded-lg font-bold flex flex-col items-center justify-center gap-1 shadow-lg active:scale-95 transition-transform`}>
                            <div className="flex items-center gap-2"><Scissors size={18}/> Process & Export Crops</div>
                            <span className="text-[9px] font-normal opacity-70 italic text-center">Groups crops into Q1.pdf, Q2.pdf, etc.</span>
                        </button>
                    </div>
                    {/* Centered Scrollable Canvas Container */}
                    <div className="flex-1 bg-[#010409] overflow-auto flex items-start justify-center p-10 bg-[radial-gradient(#30363d_1px,transparent_1px)] bg-[size:20px_20px]">
                        <div className="flex-shrink-0 relative">
                            <canvas 
                                ref={templateCanvasRef} 
                                onMouseDown={startDrag} 
                                className="shadow-2xl border border-[#30363d] bg-white cursor-crosshair transition-all" 
                            />
                        </div>
                    </div>
                </div>
                {isProcessing && (
                    <div className="absolute inset-0 bg-black/85 flex flex-col items-center justify-center z-50 animate-in fade-in duration-300">
                        <div className="relative">
                            <Loader2 className="animate-spin text-[#58a6ff] mb-6" size={64} />
                            <ScanLine className="absolute inset-0 m-auto text-white opacity-50 animate-pulse" size={24} />
                        </div>
                        <h2 className="text-2xl font-bold mb-2 tracking-tight">Processing Documents</h2>
                        <p className="text-sm text-gray-400 max-w-xs text-center">Running OCR and aligning crops across all files...</p>
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
                    <p className="text-white font-medium">Loading PDFs...</p>
                </div>
            )}
            {toast && <div className="fixed bottom-8 right-8 px-6 py-3 bg-[#238636] text-white text-sm font-bold rounded-xl shadow-2xl animate-in slide-in-from-bottom-4 duration-300">{toast}</div>}
        </div>
    );
};

export default App;