import React, { useState, useRef, useEffect } from 'react';
import { 
    Trash2, CheckCircle2, ChevronRight, ChevronLeft, 
    ScanLine, X, Loader2, FolderInput, Files, Download, 
    MousePointer2, Layers, ZoomIn, ZoomOut, Anchor, Database
} from 'lucide-react';

// --- External Libraries (Dynamic Load) ---

const PDF_LIB_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
const PDF_WORKER_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const JSZIP_URL = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";

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

const loadJsZip = async () => {
    if (window.JSZip) return window.JSZip;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = JSZIP_URL;
        script.onload = () => resolve(window.JSZip);
        script.onerror = reject;
        document.body.appendChild(script);
    });
};

const TrainingApp = () => {
    // --- State ---
    const [step, setStep] = useState(0); // 0: Upload, 1: Template, 2: Processing
    const [pdfFiles, setPdfFiles] = useState([]);
    const [previewPdfDoc, setPreviewPdfDoc] = useState(null);
    const [previewPageNum, setPreviewPageNum] = useState(1);
    const [previewTotalPages, setPreviewTotalPages] = useState(1);
    const [templateRegions, setTemplateRegions] = useState([]);
    const [isLoading, setIsLoading] = useState(false);
    const [progress, setProgress] = useState({ current: 0, total: 0, status: '' });
    const [toast, setToast] = useState(null);
    
    // Config
    const [scale, setScale] = useState(1);
    const [autoAlign, setAutoAlign] = useState(true);

    // Refs
    const canvasRef = useRef(null);
    const folderInputRef = useRef(null);

    // --- Effects ---
    useEffect(() => {
        if (step === 1 && previewPdfDoc) {
            drawTemplate();
        }
    }, [step, templateRegions, scale, previewPageNum]);

    useEffect(() => {
        if (toast) {
            const t = setTimeout(() => setToast(null), 3000);
            return () => clearTimeout(t);
        }
    }, [toast]);

    const showToast = (msg, type = 'info') => setToast({ msg, type });

    // --- PDF Loading ---
    const handleFolderSelect = async (e) => {
        const files = Array.from(e.target.files).filter(f => f.type === 'application/pdf');
        if (files.length === 0) {
            showToast("No PDFs found", "error");
            return;
        }

        setIsLoading(true);
        try {
            const pdfjs = await loadPdfLib();
            await loadJsZip(); // Preload Zip lib

            // Load first PDF for template
            const firstFile = files[0];
            const buffer = await firstFile.arrayBuffer();
            const doc = await pdfjs.getDocument(buffer).promise;

            setPdfFiles(files);
            setPreviewPdfDoc(doc);
            setPreviewTotalPages(doc.numPages);
            setPreviewPageNum(1);
            setStep(1);
        } catch (err) {
            console.error(err);
            showToast("Failed to load PDF", "error");
        } finally {
            setIsLoading(false);
        }
    };

    // --- Database Download Logic ---
    const handleDownloadDbData = async () => {
        setIsLoading(true);
        setProgress({ current: 0, total: 100, status: "Fetching data from DB..." });

        try {
            // 1. Fetch Data
            const response = await fetch('/.netlify/functions/get-scans');
            
            // Check if response is HTML (Vite fallback) instead of JSON
            const contentType = response.headers.get("content-type");
            if (contentType && contentType.includes("text/html")) {
                throw new Error("Function not found. Run 'netlify dev' to enable backend.");
            }

            if (!response.ok) throw new Error(`Server error: ${response.status}`);
            
            const rows = await response.json();

            if (rows.length === 0) {
                showToast("No training data found in database.", "error");
                return;
            }

            setProgress({ current: 0, total: rows.length, status: "Organizing files..." });

            // 2. Prepare Zip
            const JSZip = await loadJsZip();
            const zip = new JSZip();
            const rootFolder = zip.folder("db_training_data");

            // Create A, B, C, D, E folders
            const folders = {
                'A': rootFolder.folder('A'),
                'B': rootFolder.folder('B'),
                'C': rootFolder.folder('C'),
                'D': rootFolder.folder('D'),
                'E': rootFolder.folder('E') // For EMPTY
            };

            // 3. Process Rows
            rows.forEach((row, index) => {
                // Filename format: timestamp_LABEL_confidence
                // e.g. 2023-10-10_A_0.99
                const parts = row.filename.split('_');
                
                let label = 'UNKNOWN';
                if (parts.length >= 3) {
                    label = parts[parts.length - 2];
                }
                
                // Map EMPTY to E
                if (label === 'EMPTY') label = 'E';

                // Ensure valid folder, fallback to UNKNOWN if label is weird
                const targetFolder = folders[label] || rootFolder.folder('UNKNOWN');

                // Decode Base64 (remove data:image/jpeg;base64, prefix if present)
                let base64Data = row.image_data;
                if (base64Data.includes(',')) {
                    base64Data = base64Data.split(',')[1];
                }

                if (base64Data) {
                    targetFolder.file(`${row.filename}.jpg`, base64Data, {base64: true});
                }
            });

            setProgress({ current: 100, total: 100, status: "Zipping..." });

            // 4. Download
            const content = await zip.generateAsync({ type: "blob" });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = "db_training_set.zip";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            // 5. Delete from DB using the same endpoint with DELETE method
            setProgress({ current: 100, total: 100, status: "Clearing database..." });
            const deleteResponse = await fetch('/.netlify/functions/get-scans', { method: 'DELETE' });
            
            if (deleteResponse.ok) {
                showToast(`Downloaded & Cleared ${rows.length} images.`, "success");
            } else {
                showToast(`Downloaded, but failed to clear DB (${deleteResponse.status}).`, "warning");
            }

        } catch (e) {
            console.error(e);
            showToast("Download failed: " + e.message, "error");
        } finally {
            setIsLoading(false);
            setProgress({ current: 0, total: 0, status: '' });
        }
    };

    // --- Template Drawing Logic ---
    const renderPage = async (doc, pageNum) => {
        const page = await doc.getPage(pageNum);
        const viewport = page.getViewport({ scale: 1.5 }); // Good resolution for display
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: ctx, viewport }).promise;
        return { 
            dataUrl: canvas.toDataURL('image/jpeg'), 
            width: viewport.width / 1.5, // Normalize 
            height: viewport.height / 1.5 
        };
    };

    const drawTemplate = async (ghostBox = null) => {
        if (!canvasRef.current || !previewPdfDoc) return;
        const ctx = canvasRef.current.getContext('2d');
        
        const { dataUrl, width, height } = await renderPage(previewPdfDoc, previewPageNum);
        
        const img = new Image();
        img.src = dataUrl;
        await new Promise(r => img.onload = r);

        canvasRef.current.width = width * scale;
        canvasRef.current.height = height * scale;
        ctx.scale(scale, scale);
        
        ctx.drawImage(img, 0, 0, width, height);

        templateRegions.forEach(r => {
            ctx.fillStyle = 'rgba(88, 166, 255, 0.2)';
            ctx.strokeStyle = '#58a6ff';
            ctx.lineWidth = 2;
            ctx.fillRect(r.x, r.y, r.width, r.height);
            ctx.strokeRect(r.x, r.y, r.width, r.height);
            
            ctx.fillStyle = '#58a6ff';
            const textWidth = ctx.measureText(r.label).width + 8;
            ctx.fillRect(r.x, r.y - 20, textWidth, 20);
            ctx.fillStyle = '#fff';
            ctx.font = 'bold 12px sans-serif';
            ctx.fillText(r.label, r.x + 4, r.y - 5);
        });

        if (ghostBox) {
            ctx.strokeStyle = '#d29922';
            ctx.setLineDash([5, 5]);
            ctx.strokeRect(ghostBox.x, ghostBox.y, ghostBox.w, ghostBox.h);
            ctx.setLineDash([]);
        }
    };

    // --- Mouse Interaction ---
    const handleMouseDown = (e) => {
        const canvas = canvasRef.current;
        const rect = canvas.getBoundingClientRect();
        const startX = (e.clientX - rect.left) / scale;
        const startY = (e.clientY - rect.top) / scale;
        
        const clickedIdx = templateRegions.findIndex(r => 
            startX >= r.x && startX <= r.x + r.width &&
            startY >= r.y && startY <= r.y + r.height
        );

        if (clickedIdx >= 0) {
            if (e.button === 2 || e.shiftKey) {
                const n = [...templateRegions];
                n.splice(clickedIdx, 1);
                n.forEach((r, i) => r.label = `Q${i + 1}`);
                setTemplateRegions(n);
                drawTemplate();
                return;
            }
        }

        let isDragging = true;

        const onMove = (em) => {
            if (!isDragging) return;
            const mx = (em.clientX - rect.left) / scale;
            const my = (em.clientY - rect.top) / scale;
            drawTemplate({ x: startX, y: startY, w: mx - startX, h: my - startY });
        };

        const onUp = (eu) => {
            isDragging = false;
            window.removeEventListener('mousemove', onMove);
            window.removeEventListener('mouseup', onUp);
            
            const upX = (eu.clientX - rect.left) / scale;
            const upY = (eu.clientY - rect.top) / scale;
            const w = upX - startX;
            const h = upY - startY;

            if (Math.abs(w) > 5 && Math.abs(h) > 5) {
                const region = {
                    id: Date.now(),
                    x: w < 0 ? upX : startX,
                    y: h < 0 ? upY : startY,
                    width: Math.abs(w),
                    height: Math.abs(h),
                    label: `Q${templateRegions.length + 1}`
                };
                setTemplateRegions([...templateRegions, region]);
            } else {
                drawTemplate(); 
            }
        };

        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
    };

    // --- Anchors (Simplified) ---
    const findAnchors = (ctx, width, height) => {
        const { data } = ctx.getImageData(0, 0, width, height);
        const threshold = 200; 
        const stopY = Math.floor(height * 0.5); 
        
        let leftMost = { x: width, y: 0 };
        let rightMost = { x: 0, y: 0 };
        let found = false;

        for (let y = height - 1; y >= stopY; y--) {
            for (let x = 0; x < width; x++) {
                const i = (y * width + x) * 4;
                if (data[i] < threshold) {
                    if (x < leftMost.x) leftMost = { x, y };
                    if (x > rightMost.x) rightMost = { x, y };
                    found = true;
                }
            }
        }
        return found ? { left: leftMost, right: rightMost } : null;
    };

    // --- CORE LOGIC: Crop & Zip (Local PDFs) ---
    const handleGenerateDataset = async () => {
        if (templateRegions.length === 0) {
            showToast("Define at least one answer box.", "error");
            return;
        }

        setStep(2); // Processing UI
        setIsLoading(true);

        try {
            const JSZip = await loadJsZip();
            const zip = new JSZip();
            const pdfjs = await loadPdfLib();

            const datasetFolder = zip.folder("dataset");
            templateRegions.forEach(r => datasetFolder.folder(r.label));

            // Template Anchors
            const tPage = await previewPdfDoc.getPage(previewPageNum);
            const tView = tPage.getViewport({ scale: 2.0 });
            const tCanvas = document.createElement('canvas');
            tCanvas.width = tView.width; tCanvas.height = tView.height;
            const tCtx = tCanvas.getContext('2d');
            await tPage.render({ canvasContext: tCtx, viewport: tView }).promise;
            const tAnchors = findAnchors(tCtx, tCanvas.width, tCanvas.height);

            for (let i = 0; i < pdfFiles.length; i++) {
                const file = pdfFiles[i];
                setProgress({ current: i + 1, total: pdfFiles.length, status: `Processing ${file.name}...` });

                const buffer = await file.arrayBuffer();
                const doc = await pdfjs.getDocument(buffer).promise;
                if (doc.numPages < previewPageNum) continue;

                const page = await doc.getPage(previewPageNum);
                const viewport = page.getViewport({ scale: 2.0 }); 
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                canvas.width = viewport.width; canvas.height = viewport.height;
                await page.render({ canvasContext: ctx, viewport }).promise;

                let offX = 0, offY = 0;
                if (autoAlign && tAnchors) {
                    const target = findAnchors(ctx, canvas.width, canvas.height);
                    if (target) {
                        offX = target.left.x - tAnchors.left.x;
                        offY = target.left.y - tAnchors.left.y;
                    }
                }

                for (const region of templateRegions) {
                    const rX = (region.x * 2.0) + offX;
                    const rY = (region.y * 2.0) + offY;
                    const rW = region.width * 2.0;
                    const rH = region.height * 2.0;

                    const cropCanvas = document.createElement('canvas');
                    cropCanvas.width = rW; cropCanvas.height = rH;
                    const cropCtx = cropCanvas.getContext('2d');
                    cropCtx.drawImage(canvas, rX, rY, rW, rH, 0, 0, rW, rH);

                    const blob = await new Promise(r => cropCanvas.toBlob(r, 'image/jpeg', 0.95));
                    const safeName = file.name.replace(/\.pdf$/i, '').replace(/[^a-z0-9]/gi, '_');
                    datasetFolder.folder(region.label).file(`${safeName}_${region.label}.jpg`, blob);
                }
            }

            setProgress({ current: pdfFiles.length, total: pdfFiles.length, status: "Compressing..." });
            const content = await zip.generateAsync({ type: "blob" });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = "local_dataset.zip";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            showToast("Dataset generated!", "success");
            setStep(1); 

        } catch (e) {
            console.error(e);
            showToast("Error: " + e.message, "error");
            setStep(1);
        } finally {
            setIsLoading(false);
        }
    };

    // --- UI (GitHub Dark Theme) ---
    const theme = {
        bg: 'bg-[#0d1117]',
        sidebar: 'bg-[#161b22]',
        border: 'border-[#30363d]',
        text: 'text-[#c9d1d9]',
        textMuted: 'text-[#8b949e]',
        accent: 'text-[#58a6ff]',
        accentBg: 'bg-[#1f6feb]',
        success: 'text-[#3fb950]',
        buttonHover: 'hover:bg-[#21262d]',
        inputBg: 'bg-[#0d1117]'
    };

    return (
        <div className={`h-screen w-screen ${theme.bg} ${theme.text} flex flex-col font-sans`}>
            {/* Header */}
            <div className={`h-16 border-b ${theme.border} ${theme.sidebar} flex items-center justify-between px-6`}>
                <div className="flex items-center gap-2">
                    <Layers className={theme.success} size={24}/>
                    <h1 className="font-bold text-lg">Dataset Creator</h1>
                </div>
                
                {/* Header Actions */}
                <div className="flex items-center gap-4">
                    {/* Database Download Button */}
                    <button 
                        onClick={handleDownloadDbData}
                        disabled={isLoading}
                        className={`px-3 py-1.5 ${theme.buttonHover} border ${theme.border} rounded text-xs flex items-center gap-2 text-[#e3b341] transition-all`}
                        title="Download collected training data"
                    >
                        {isLoading && step === 0 ? <Loader2 size={14} className="animate-spin"/> : <Database size={14}/>}
                        Download DB Data
                    </button>

                    {step === 1 && (
                        <div className="flex items-center gap-4 border-l border-[#30363d] pl-4">
                            <label className="flex items-center gap-2 text-xs cursor-pointer select-none">
                                <input 
                                    type="checkbox" 
                                    checked={autoAlign}
                                    onChange={e => setAutoAlign(e.target.checked)}
                                    className="accent-[#1f6feb]"
                                />
                                <span className={autoAlign ? theme.accent : theme.textMuted}>
                                    <Anchor size={14} className="inline mr-1"/>
                                    Auto-Align
                                </span>
                            </label>
                            <div className={`text-xs ${theme.textMuted} flex items-center gap-2`}>
                                <MousePointer2 size={14}/> Draw boxes
                            </div>
                        </div>
                    )}
                </div>
            </div>

            {/* Content Area */}
            <div className="flex-1 overflow-hidden relative">
                
                {/* STEP 0: Upload */}
                {step === 0 && (
                    <div className="h-full flex flex-col items-center justify-center p-10 animate-fade-in">
                        <div className={`border-2 border-dashed ${theme.border} rounded-xl ${theme.bg} p-12 flex flex-col items-center text-center max-w-lg w-full`}>
                            <FolderInput size={48} className={`${theme.textMuted} mb-4`}/>
                            <h3 className="text-xl font-semibold mb-2">Select Source PDFs</h3>
                            <p className={`${theme.textMuted} mb-6 text-sm`}>
                                Select folder containing answer sheets to extract crops locally.
                            </p>
                            
                            <input 
                                type="file" 
                                ref={folderInputRef}
                                className="hidden" 
                                multiple 
                                accept="application/pdf"
                                onChange={handleFolderSelect}
                                {...{ webkitdirectory: "", mozdirectory: "" }}
                            />
                            <button 
                                onClick={() => folderInputRef.current.click()}
                                disabled={isLoading}
                                className={`px-6 py-3 ${theme.accentBg} hover:opacity-90 text-white rounded-md font-medium flex items-center gap-2 transition-all`}
                            >
                                {isLoading ? <Loader2 className="animate-spin"/> : <FolderInput size={18}/>}
                                Select Folder
                            </button>
                        </div>
                    </div>
                )}

                {/* STEP 1: Template Editor */}
                {step === 1 && (
                    <div className="h-full flex">
                        <div className={`w-72 ${theme.sidebar} border-r ${theme.border} flex flex-col p-4 z-10`}>
                            <h3 className="font-semibold mb-4 text-[#e6edf3] flex items-center gap-2">
                                <ScanLine size={18} className={theme.accent}/> Template
                            </h3>

                            <div className="mb-6">
                                <label className={`text-xs ${theme.textMuted} uppercase font-bold`}>Target Page</label>
                                <div className="flex items-center gap-2 mt-2">
                                    <button 
                                        onClick={() => setPreviewPageNum(p => Math.max(1, p - 1))}
                                        className={`p-1.5 bg-[#21262d] rounded border ${theme.border} hover:border-[#8b949e]`}
                                    >
                                        <ChevronLeft size={16}/>
                                    </button>
                                    <span className="text-sm font-mono flex-1 text-center">{previewPageNum} / {previewTotalPages}</span>
                                    <button 
                                        onClick={() => setPreviewPageNum(p => Math.min(previewTotalPages, p + 1))}
                                        className={`p-1.5 bg-[#21262d] rounded border ${theme.border} hover:border-[#8b949e]`}
                                    >
                                        <ChevronRight size={16}/>
                                    </button>
                                </div>
                            </div>

                            <div className={`flex-1 overflow-y-auto mb-4 border ${theme.border} rounded-md ${theme.bg} p-2`}>
                                {templateRegions.length === 0 ? (
                                    <div className={`h-full flex flex-col items-center justify-center ${theme.textMuted} text-xs text-center p-4`}>
                                        <p>No boxes defined.</p>
                                        <p className="mt-2">Drag on the image to create answer zones.</p>
                                    </div>
                                ) : (
                                    <div className="space-y-1">
                                        {templateRegions.map((r, i) => (
                                            <div key={r.id} className={`flex justify-between items-center text-xs p-2 bg-[#21262d] rounded border ${theme.border} group`}>
                                                <span className={`font-mono ${theme.accent} font-bold`}>{r.label}</span>
                                                <div className="flex gap-2 text-[#8b949e]">
                                                    <button 
                                                        onClick={() => {
                                                            const n = [...templateRegions];
                                                            n.splice(i, 1);
                                                            setTemplateRegions(n);
                                                        }}
                                                        className="text-[#f85149] hover:bg-[#30363d] rounded px-1 opacity-0 group-hover:opacity-100"
                                                    >
                                                        <Trash2 size={12}/>
                                                    </button>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>

                            <button 
                                onClick={handleGenerateDataset}
                                className={`w-full py-3 bg-[#238636] hover:bg-[#2ea043] text-white rounded-md font-medium flex items-center justify-center gap-2 shadow-lg transition-all`}
                            >
                                <Download size={18}/>
                                Generate Dataset
                            </button>
                            <p className={`text-[10px] ${theme.textMuted} text-center mt-2`}>
                                Will process {pdfFiles.length} files
                            </p>
                        </div>

                        <div className={`flex-1 ${theme.bg} overflow-auto flex items-center justify-center p-8 relative`}>
                            <div className={`absolute top-4 right-4 flex gap-1 ${theme.sidebar} p-1 rounded border ${theme.border} z-20 shadow-lg`}>
                                <button onClick={() => setScale(s => Math.max(0.5, s - 0.1))} className={`p-1.5 ${theme.buttonHover} rounded ${theme.text}`} title="Zoom Out"><ZoomOut size={16}/></button>
                                <span className={`text-xs font-mono flex items-center px-2 ${theme.textMuted} select-none`}>{Math.round(scale * 100)}%</span>
                                <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className={`p-1.5 ${theme.buttonHover} rounded ${theme.text}`} title="Zoom In"><ZoomIn size={16}/></button>
                            </div>

                            <div className={`relative shadow-2xl border ${theme.border} bg-white`}>
                                <canvas 
                                    ref={canvasRef}
                                    onMouseDown={handleMouseDown}
                                    onContextMenu={(e) => e.preventDefault()}
                                    className="cursor-crosshair block"
                                />
                            </div>
                        </div>
                    </div>
                )}

                {/* STEP 2: Processing */}
                {step === 2 && (
                    <div className={`absolute inset-0 ${theme.bg} z-50 flex flex-col items-center justify-center`}>
                        <div className={`${theme.sidebar} border ${theme.border} p-8 rounded-2xl shadow-2xl max-w-md w-full text-center`}>
                            <div className="relative mb-6">
                                <div className="w-16 h-16 border-4 border-[#30363d] border-t-[#58a6ff] rounded-full animate-spin mx-auto"></div>
                                <div className="absolute inset-0 flex items-center justify-center font-bold text-xs text-[#58a6ff]">
                                    {Math.round((progress.current / progress.total) * 100)}%
                                </div>
                            </div>
                            <h2 className="text-xl font-bold text-[#e6edf3] mb-2">Generating Dataset</h2>
                            <p className={`${theme.textMuted} text-sm mb-6 font-mono`}>{progress.status}</p>
                            
                            <div className={`w-full ${theme.border} bg-[#0d1117] h-2 rounded-full overflow-hidden`}>
                                <div 
                                    className="h-full bg-[#238636] transition-all duration-300 ease-out" 
                                    style={{ width: `${(progress.current / progress.total) * 100}%` }}
                                ></div>
                            </div>
                            <p className={`text-xs ${theme.textMuted} mt-4`}>
                                {progress.current} / {progress.total} files processed
                            </p>
                        </div>
                    </div>
                )}
            </div>

            {toast && (
                <div className={`fixed bottom-6 right-6 px-4 py-3 rounded-md shadow-lg text-white text-sm font-medium z-50 flex items-center gap-3 animate-fade-in-up border ${theme.border}
                    ${toast.type === 'error' ? 'bg-[#da3633]' : 'bg-[#1f6feb]'}
                `}>
                    {toast.type === 'error' ? <X size={16}/> : <CheckCircle2 size={16}/>}
                    {toast.msg}
                </div>
            )}
        </div>
    );
};

export default TrainingApp;