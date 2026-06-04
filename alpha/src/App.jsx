import React, { useState, useRef, useEffect, useMemo } from 'react';
import { Upload, FileSpreadsheet, Plus, Trash2, CheckCircle2, ScanSearch, Settings, ChevronRight, ChevronDown, ChevronUp, Download, ZoomIn, ZoomOut, LayoutTemplate, ChevronLeft, Layers, ScanLine, GraduationCap, Hash, UserSquare2, X, Loader2, RotateCw, Menu, GitBranch, Camera, FileText, Signature, Save, BookmarkPlus } from 'lucide-react';

const App = () => {
    // State for multiple pages
    const [file, setFile] = useState(null);
    const [pages, setPages] = useState([]); // Array of { id, imageUrl, width, height, results, regions }
    const [currentPageIndex, setCurrentPageIndex] = useState(0);

    // New state for manual page input
    const [pageInput, setPageInput] = useState("1");

    const [scale, setScale] = useState(1);
    const [selectedRegionId, setSelectedRegionId] = useState(null);
    const [isProcessing, setIsProcessing] = useState(false);
    const [progress, setProgress] = useState({ current: 0, total: 0 }); // Progress state
    const [showResults, setShowResults] = useState(false);
    const [toast, setToast] = useState(null); // Toast notification state
    const [showIdModal, setShowIdModal] = useState(false);
    const [idModalIdx, setIdModalIdx] = useState(0);

    // Grading State
    const [answerKeyInput, setAnswerKeyInput] = useState("");
    const [parsedAnswerKey, setParsedAnswerKey] = useState({});

    // Marks State (Batch Sections)
    const [isMarksSettingsOpen, setIsMarksSettingsOpen] = useState(false);
    const [weightSections, setWeightSections] = useState([
        { id: 1, start: 1, end: 60, mark: 1 }
    ]);

    // Built-in default template — always present, cannot be edited or deleted
    const DEFAULT_TEMPLATE_ID = '__default__';
    const getDefaultTemplate = () => {
        // Use standard region ratios (derived from getStandardRegions w/ no offset)
        // Approximate ratios based on getStandardRegions constants
        const IMG_W = 1, IMG_H = 1;
        const idW = 0.015;
        const idY = 0.066;
        const idH = 0.135;
        const ansW = 0.165;
        const ansY = 0.273;
        const ansH = 0.468;
        return {
            id: DEFAULT_TEMPLATE_ID,
            name: 'Default',
            isBuiltIn: true,
            createdAt: '—',
            idRegions: [
                { xRatio: 0.435, yRatio: idY, wRatio: idW, hRatio: idH * 0.7, rows: 7, cols: 1, labels: ['1','2','3','4','5','6','7'], gapHeightRatio: 1, hasGaps: false, suffix: 'level' },
                { xRatio: 0.51, yRatio: idY, wRatio: idW, hRatio: idH * 0.6, rows: 6, cols: 1, labels: ['A','B','C','D','E','S'], gapHeightRatio: 1, hasGaps: false, suffix: 'letter' },
                { xRatio: 0.605, yRatio: idY, wRatio: idW, hRatio: idH, rows: 10, cols: 1, labels: ['0','1','2','3','4','5','6','7','8','9'], gapHeightRatio: 1, hasGaps: false, suffix: 'n1' },
                { xRatio: 0.645, yRatio: idY, wRatio: idW, hRatio: idH, rows: 10, cols: 1, labels: ['0','1','2','3','4','5','6','7','8','9'], gapHeightRatio: 1, hasGaps: false, suffix: 'n2' },
            ],
            answerRegions: [
                { xRatio: 0.43, yRatio: ansY, wRatio: ansW, hRatio: ansH, rows: 35, cols: 4, startQ: 1, labels: ['A','B','C','D'], gapHeightRatio: 0.6, hasGaps: true },
                { xRatio: 0.723, yRatio: ansY, wRatio: ansW, hRatio: ansH, rows: 35, cols: 4, startQ: 31, labels: ['A','B','C','D'], gapHeightRatio: 0.6, hasGaps: true },
            ],
            sourceXOffset: 0,
            sourceYOffset: 0,
            useAlignmentY: true,
            useAlignmentX: true,
        };
    };

    // Custom Template State
    const [templates, setTemplates] = useState(() => {
        try {
            const saved = localStorage.getItem('mc-sheet-templates');
            return saved ? JSON.parse(saved) : [];
        } catch { return []; }
    });
    const [lastAppliedTemplateId, setLastAppliedTemplateId] = useState(() => {
        try { return localStorage.getItem('mc-sheet-last-template') || DEFAULT_TEMPLATE_ID; }
        catch { return DEFAULT_TEMPLATE_ID; }
    });

    // Combined list: default always first, then user templates
    const allTemplates = useMemo(() => [getDefaultTemplate(), ...templates], [templates]);
    const [showSaveTemplateModal, setShowSaveTemplateModal] = useState(false);
    const [newTemplateName, setNewTemplateName] = useState('');
    const [isTemplateSectionOpen, setIsTemplateSectionOpen] = useState(false);

    // Template Builder State (Visual Drag-to-Draw)
    const [drawMode, setDrawMode] = useState(false); // true = user is drawing on canvas
    const [drawnBlocks, setDrawnBlocks] = useState([]); // { uid, xRatio, yRatio, wRatio, hRatio, blockType, rows, cols, labels, startQ, suffix, hasGaps, gapHeightRatio }
    const [currentDraw, setCurrentDraw] = useState(null); // { startX, startY, endX, endY } in image coords while dragging
    const [drawTemplateMode, setDrawTemplateMode] = useState('idle'); // 'idle' | 'naming'
    const [drawTemplateName, setDrawTemplateName] = useState('');
    const [selectedDrawnBlock, setSelectedDrawnBlock] = useState(null);
    const [drawNextQ, setDrawNextQ] = useState(1);
    const [editingTemplateId, setEditingTemplateId] = useState(null);
    const [drawLineY, setDrawLineY] = useState(null); // adjustable detected horizontal line Y (image coords)
    const [drawLineX, setDrawLineX] = useState(null); // adjustable detected vertical line X (image coords)
    const [isDraggingLineY, setIsDraggingLineY] = useState(false);
    const [isDraggingLineX, setIsDraggingLineX] = useState(false);
    const [useLineY, setUseLineY] = useState(true); // toggle: use horizontal alignment
    const [useLineX, setUseLineX] = useState(true); // toggle: use vertical alignment

    // Persist templates to localStorage
    useEffect(() => {
        localStorage.setItem('mc-sheet-templates', JSON.stringify(templates));
    }, [templates]);

    // Persist last-applied template id
    useEffect(() => {
        localStorage.setItem('mc-sheet-last-template', lastAppliedTemplateId);
    }, [lastAppliedTemplateId]);

    // --- Draw Mode Functions ---
    const startDrawMode = () => {
        if (!currentPage) return;
        setDrawMode(true);
        setDrawnBlocks([]);
        setCurrentDraw(null);
        setSelectedDrawnBlock(null);
        setDrawNextQ(1);
        setDrawTemplateMode('idle');
        setDrawTemplateName('');
        setSelectedRegionId(null);
        setDrawLineY(currentPage.detectedLineY || Math.round(currentPage.height * 0.227));
        setDrawLineX(currentPage.detectedLineX || Math.round(currentPage.width * 0.355));
        setIsDraggingLineY(false);
        setIsDraggingLineX(false);
        setUseLineY(true);
        setUseLineX(true);
    };

    // Sync alignment lines when navigating pages in draw mode
    useEffect(() => {
        if (drawMode && currentPage) {
            setDrawLineY(currentPage.detectedLineY || Math.round(currentPage.height * 0.227));
            setDrawLineX(currentPage.detectedLineX || Math.round(currentPage.width * 0.355));
            setIsDraggingLineY(false);
            setIsDraggingLineX(false);
        }
    }, [currentPageIndex, drawMode]);

    const cancelDrawMode = () => {
        setDrawMode(false);
        setDrawnBlocks([]);
        setCurrentDraw(null);
        setSelectedDrawnBlock(null);
        setDrawTemplateMode('idle');
        setDrawTemplateName('');
        setEditingTemplateId(null);
        setDrawLineY(null);
        setDrawLineX(null);
        setIsDraggingLineY(false);
        setIsDraggingLineX(false);
    };

    const updateDrawnBlock = (uid, field, value) => {
        setDrawnBlocks(prev => prev.map(b => {
            if (b.uid !== uid) return b;
            const updated = { ...b, [field]: value };
            if (field === 'blockType') {
                if (value === 'id') {
                    updated.labels = ['0','1','2','3','4','5','6','7','8','9'];
                    updated.rows = 10;
                    updated.cols = 1;
                    updated.suffix = updated.suffix || 'custom';
                    updated.hasGaps = false;
                    updated.gapHeightRatio = 1;
                } else if (value === 'answer') {
                    updated.labels = ['A','B','C','D'];
                    updated.rows = 5;
                    updated.cols = 4;
                    updated.startQ = updated.startQ || 1;
                    updated.hasGaps = false;
                    updated.gapHeightRatio = 1;
                }
            }
            if (field === 'hasGaps') {
                updated.gapHeightRatio = value ? 0.6 : 1;
            }
            // Store raw labels text — parsing happens on save
            if (field === 'labels') {
                updated.labelsText = value;
                const labels = value.split(',').map(s => s.trim()).filter(Boolean);
                updated.labels = labels;
                if (b.blockType === 'id') {
                    updated.rows = labels.length;
                } else if (b.blockType === 'answer' && labels.length > 0) {
                    updated.cols = labels.length;
                }
            }
            return updated;
        }));
    };

    const removeDrawnBlock = (uid) => {
        setDrawnBlocks(prev => prev.filter(b => b.uid !== uid));
        if (selectedDrawnBlock === uid) setSelectedDrawnBlock(null);
    };

    const saveDrawnTemplate = async () => {
        if (!drawTemplateName.trim() || drawnBlocks.length === 0) return;
        // Convert percentage-based ratios (0-100) to decimal ratios (0-1) for storage
        const convertToDecimal = (b) => {
            const { uid: _u, blockType: _t, labelsText: _lt, ...rest } = b;
            return {
                ...rest,
                xRatio: rest.xRatio / 100,
                yRatio: rest.yRatio / 100,
                wRatio: rest.wRatio / 100,
                hRatio: rest.hRatio / 100,
            };
        };
        const idRegions = drawnBlocks.filter(b => b.blockType === 'id').map(convertToDecimal);
        const answerRegions = drawnBlocks.filter(b => b.blockType === 'answer').map(convertToDecimal);
        const savedName = drawTemplateName.trim();

        // Capture source page offsets for re-alignment.
        // If the user adjusted the alignment lines in draw mode, use their adjusted
        // line positions to derive the offsets rather than the auto-detected ones.
        let xOffset = 0;
        let yOffset = 0;
        if (currentPage) {
            const auto = await detectOffsetsFromImage(currentPage.imageUrl, currentPage.width, currentPage.height);
            // Y axis: use adjusted line if toggle is on
            if (useLineY && drawLineY != null) {
                yOffset = drawLineY - (currentPage.height * 0.227);
            } else if (!useLineY) {
                yOffset = 0; // alignment disabled
            } else {
                yOffset = auto.yOffset;
            }
            // X axis: use adjusted line if toggle is on
            if (useLineX && drawLineX != null) {
                xOffset = drawLineX - (currentPage.width * 0.355);
            } else if (!useLineX) {
                xOffset = 0; // alignment disabled
            } else {
                xOffset = auto.xOffset;
            }
        }

        if (editingTemplateId) {
            setTemplates(prev => prev.map(t => {
                if (t.id !== editingTemplateId) return t;
                return { ...t, name: savedName, idRegions, answerRegions, sourceXOffset: xOffset, sourceYOffset: yOffset, useAlignmentY: useLineY, useAlignmentX: useLineX };
            }));
            setToast(`Template "${savedName}" updated!`);
        } else {
            const newTpl = {
                id: Date.now().toString(),
                name: savedName,
                createdAt: new Date().toISOString(),
                idRegions, answerRegions,
                sourceXOffset: xOffset,
                sourceYOffset: yOffset,
                useAlignmentY: useLineY,
                useAlignmentX: useLineX,
            };
            setTemplates(prev => [...prev, newTpl]);
            setToast(`Template "${savedName}" created!`);
        }
        setTimeout(() => setToast(null), 3000);
        cancelDrawMode();
    };


    const canvasRef = useRef(null);
    const containerRef = useRef(null);

    // Helper to get current page data safely
    const currentPage = pages[currentPageIndex];
    // Helper to get regions for rendering (default to empty if no page loaded)
    const currentRegions = currentPage?.regions || [];

    // Sync input with actual page index whenever navigation happens
    useEffect(() => {
        setPageInput(String(currentPageIndex + 1));
    }, [currentPageIndex]);

    // Helper to get mark for a specific question based on sections
    const getQuestionMark = (qNum) => {
        let mark = 1; // Default
        weightSections.forEach(section => {
            const sStart = section.start === '' ? 0 : section.start;
            const sEnd = section.end === '' ? 0 : section.end;
            const sMark = section.mark === '' ? 0 : section.mark;

            if (qNum >= sStart && qNum <= sEnd) {
                mark = parseFloat(sMark);
            }
        });
        return (isNaN(mark) || mark <= 0) ? 1 : mark;
    };

    // Helper to extract Student ID from results (dynamic — reads all ID regions)
    const getPageName = (page, index) => {
        if (!page.results) return `Page ${index + 1}`;

        const getVal = (id) => {
            const res = page.results[id];
            if (!res || res.length === 0) return '?';
            const found = res[0];
            if (found && found.label && found.label !== 'BLANK' && found.label !== 'MULT' && found.label !== '?') {
                return found.label;
            }
            return '?';
        };

        // Dynamically find all ID regions on this page
        const idRegions = (page.regions || []).filter(r => r.type === 'id');
        if (idRegions.length === 0) return `Page ${index + 1}`;

        const values = idRegions.map(r => getVal(r.id));
        if (values.every(v => v === '?')) return `Page ${index + 1}`;

        return values.join('');
    };

    // Helper: Get all ID region info for a page (dynamic)
    const getIdFieldsForPage = (pageIdx) => {
        const page = pages[pageIdx];
        if (!page) return [];
        return (page.regions || [])
            .filter(r => r.type === 'id')
            .map(r => {
                const suffix = r.id.split('_id_')[1] || '?';
                return {
                    regionId: r.id,
                    suffix,
                    label: suffix.toUpperCase(),
                    options: r.labels || ['0','1','2','3','4','5','6','7','8','9'],
                };
            });
    };

    // Load External Libraries (PDF.js and XLSX)
    useEffect(() => {
        // Load XLSX
        const xlsxScript = document.createElement('script');
        xlsxScript.src = "https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";
        xlsxScript.async = true;
        document.body.appendChild(xlsxScript);

        // Load PDF.js
        const pdfScript = document.createElement('script');
        pdfScript.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
        pdfScript.async = true;
        pdfScript.onload = () => {
            const pdfjsLib = window.pdfjsLib || window['pdfjs-dist/build/pdf'];
            if (pdfjsLib) {
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
                console.log("PDF.js loaded successfully");
            }
        };
        document.body.appendChild(pdfScript);

        return () => {
            if (document.body.contains(xlsxScript)) document.body.removeChild(xlsxScript);
            if (document.body.contains(pdfScript)) document.body.removeChild(pdfScript);
        };
    }, []);

    // --- Keyboard Shortcuts ---
    useEffect(() => {
        const handleKeyDown = (e) => {
            // Prevent interference if typing in inputs
            if (['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;

            if ((e.metaKey || e.ctrlKey) && e.key === 'o') {
                e.preventDefault();
                document.getElementById('file-upload-input')?.click();
            }
            if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') {
                e.preventDefault();
                if (pages.length > 0 && !isProcessing) runBatchDetection();
            }
            if ((e.metaKey || e.ctrlKey) && e.key === 'e') {
                e.preventDefault();
                if (showResults) exportExcel();
            }
            if (e.key === 'ArrowLeft') {
                setCurrentPageIndex(p => Math.max(0, p - 1));
            }
            if (e.key === 'ArrowRight') {
                setCurrentPageIndex(p => Math.min(pages.length - 1, p + 1));
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [pages, isProcessing, showResults]);

    // --- Strict Answer Key Parsing Logic ---
    useEffect(() => {
        const rawText = answerKeyInput.toUpperCase();
        // Split by whitespace to handle words/tokens individually
        const tokens = rawText.split(/\s+/);
        
        const keyMap = {};
        let currentIndex = 1;

        tokens.forEach(token => {
            // Logic:
            // 1. If token contains any letter [E-Z], it's a word like "Answer" or "Part", ignore it.
            // 2. If token only contains [A-D] and non-letters (numbers/punctuation), accept it.
            //    e.g. "1.A" -> OK. "ABCD" -> OK. "Answer" -> Skip.
            
            if (!/[E-Z]/.test(token)) {
                // Extract valid answer characters
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

    // --- Helper: Excel Column Name Generator (1->A, 2->B... 27->AA) ---
    const getExcelCol = (n) => {
        let s = "";
        while (n >= 0) {
            s = String.fromCharCode(n % 26 + 65) + s;
            n = Math.floor(n / 26) - 1;
        }
        return s;
    };

    // --- Helper: Row Layout Calculation ---
    const calculateRowLayout = (regionHeight, totalRows, gapRatio = 0.4, hasGaps = true) => {
        const configs = [];
        let totalUnits = 0;
        for (let i = 0; i < totalRows; i++) {
            const isGap = hasGaps && ((i + 1) % 6 === 0);
            totalUnits += isGap ? gapRatio : 1;
        }
        const pxPerUnit = regionHeight / totalUnits;
        let currentY = 0;
        for (let i = 0; i < totalRows; i++) {
            const isGap = hasGaps && ((i + 1) % 6 === 0);
            const rowHeight = (isGap ? gapRatio : 1) * pxPerUnit;
            configs.push({ y: currentY, h: rowHeight, isGap: isGap });
            currentY += rowHeight;
        }
        return configs;
    };

    // --- Helper: Line Detection ---
    const detectVerticalOffset = (ctx, width, height) => {
        const EXPECTED_Y_RATIO = 0.227;
        const searchStartY = Math.floor(height * 0.3);
        const searchEndY = Math.floor(height * 0.5);
        const searchStartX = Math.floor(width * 0.2);
        const searchWidth = Math.floor(width * 0.8);
        try {
            const pixels = ctx.getImageData(searchStartX, searchStartY, searchWidth, searchEndY - searchStartY);
            const data = pixels.data;
            const searchH = searchEndY - searchStartY;
            let maxDarkness = 0;
            let bestY = -1;
            for (let y = 0; y < searchH; y++) {
                let darkPixels = 0;
                for (let x = 0; x < searchWidth; x++) {
                    const idx = (y * searchWidth + x) * 4;
                    const val = (data[idx] + data[idx + 1] + data[idx + 2]) / 3;
                    if (val < 200) darkPixels++;
                }
                if (darkPixels > searchWidth * 0.4) {
                    if (darkPixels > maxDarkness) { maxDarkness = darkPixels; bestY = y; }
                }
            }
            if (bestY !== -1) {
                const detectedAbsoluteY = searchStartY + bestY;
                return detectedAbsoluteY - (height * EXPECTED_Y_RATIO);
            }
        } catch (e) { console.warn("Y-alignment failed", e); }
        return 0;
    };

    // --- Helper: Get absolute X position of detected vertical line ---
    const getDetectedLineX = (ctx, width, height) => {
        const EXPECTED_X_RATIO = 0.355;
        const searchStartX = Math.floor(width * 0.2);
        const searchEndX = Math.floor(width * 0.5);
        const searchStartY = Math.floor(height * 0.5);
        const searchHeight = Math.floor(height * 0.9);
        try {
            const pixels = ctx.getImageData(searchStartX, searchStartY, searchEndX - searchStartX, searchHeight);
            const data = pixels.data;
            const searchW = searchEndX - searchStartX;
            let maxDarkness = 0;
            let bestX = -1;
            for (let x = 0; x < searchW; x++) {
                let darkPixels = 0;
                for (let y = 0; y < searchHeight; y++) {
                    const idx = (y * searchW + x) * 4;
                    const val = (data[idx] + data[idx + 1] + data[idx + 2]) / 3;
                    if (val < 200) darkPixels++;
                }
                if (darkPixels > searchHeight * 0.5) {
                    if (darkPixels > maxDarkness) { maxDarkness = darkPixels; bestX = x; }
                }
            }
            if (bestX !== -1) {
                return searchStartX + bestX;
            }
        } catch (e) { console.warn("Line X detection failed", e); }
        return -1;
    };

    // --- Helper: Get absolute Y position of detected horizontal line ---
    const getDetectedLineY = (ctx, width, height) => {
        const searchStartY = Math.floor(height * 0.3);
        const searchEndY = Math.floor(height * 0.5);
        const searchStartX = Math.floor(width * 0.2);
        const searchWidth = Math.floor(width * 0.8);
        try {
            const pixels = ctx.getImageData(searchStartX, searchStartY, searchWidth, searchEndY - searchStartY);
            const data = pixels.data;
            const searchH = searchEndY - searchStartY;
            let maxDarkness = 0;
            let bestY = -1;
            for (let y = 0; y < searchH; y++) {
                let darkPixels = 0;
                for (let x = 0; x < searchWidth; x++) {
                    const idx = (y * searchWidth + x) * 4;
                    const val = (data[idx] + data[idx + 1] + data[idx + 2]) / 3;
                    if (val < 200) darkPixels++;
                }
                if (darkPixels > searchWidth * 0.4) {
                    if (darkPixels > maxDarkness) { maxDarkness = darkPixels; bestY = y; }
                }
            }
            if (bestY !== -1) {
                return searchStartY + bestY;
            }
        } catch (e) { console.warn("Line Y detection failed", e); }
        return -1;
    };

    const detectHorizontalOffset = (ctx, width, height) => {
        const EXPECTED_X_RATIO = 0.355;
        const searchStartX = Math.floor(width * 0.2);
        const searchEndX = Math.floor(width * 0.5);
        const searchStartY = Math.floor(height * 0.5);
        const searchHeight = Math.floor(height * 0.9);
        try {
            const pixels = ctx.getImageData(searchStartX, searchStartY, searchEndX - searchStartX, searchHeight);
            const data = pixels.data;
            const searchW = searchEndX - searchStartX;
            let maxDarkness = 0;
            let bestX = -1;
            for (let x = 0; x < searchW; x++) {
                let darkPixels = 0;
                for (let y = 0; y < searchHeight; y++) {
                    const idx = (y * searchW + x) * 4;
                    const val = (data[idx] + data[idx + 1] + data[idx + 2]) / 3;
                    if (val < 200) darkPixels++;
                }
                if (darkPixels > searchHeight * 0.5) {
                    if (darkPixels > maxDarkness) { maxDarkness = darkPixels; bestX = x; }
                }
            }
            if (bestX !== -1) {
                const detectedAbsoluteX = searchStartX + bestX;
                return detectedAbsoluteX - (width * EXPECTED_X_RATIO);
            }
        } catch (e) { console.warn("X-alignment failed", e); }
        return 0;
    };

    // --- Template Logic ---
    const getStandardRegions = (imgW, imgH, pageIdPrefix, xOffset = 0, yOffset = 0) => {
        const LABELS_ANS = ['A', 'B', 'C', 'D'];
        const ANS_BLOCK_W = imgW * 0.165;
        const ANS_START_Y = (imgH * 0.273) + yOffset;
        const ANS_BLOCK_H = imgH * 0.468;
        const ROWS_PER_BLOCK = 35;
        const X1 = (imgW * 0.43) + xOffset;
        const X2 = (imgW * 0.723) + xOffset;
        const ID_Y = (imgH * 0.066) + yOffset;
        const ID_H = imgH * 0.135;
        const ID_W_SMALL = imgW * 0.015;
        const ID_X_LEVEL = (imgW * 0.435) + xOffset;
        const ID_X_LETTER = (imgW * 0.51) + xOffset;
        const ID_X_N1 = (imgW * 0.605) + xOffset;
        const ID_X_N2 = (imgW * 0.645) + xOffset;

        return [
            { id: `${pageIdPrefix}_id_level`, type: 'id', x: ID_X_LEVEL, y: ID_Y, w: ID_W_SMALL, h: ID_H * 0.7, rows: 7, cols: 1, labels: ['1', '2', '3', '4', '5', '6', '7'], gapHeightRatio: 1, hasGaps: false },
            { id: `${pageIdPrefix}_id_letter`, type: 'id', x: ID_X_LETTER, y: ID_Y, w: ID_W_SMALL, h: ID_H * 0.6, rows: 6, cols: 1, labels: ['A', 'B', 'C', 'D', 'E', 'S'], gapHeightRatio: 1, hasGaps: false },
            { id: `${pageIdPrefix}_id_n1`, type: 'id', x: ID_X_N1, y: ID_Y, w: ID_W_SMALL, h: ID_H, rows: 10, cols: 1, labels: ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'], gapHeightRatio: 1, hasGaps: false },
            { id: `${pageIdPrefix}_id_n2`, type: 'id', x: ID_X_N2, y: ID_Y, w: ID_W_SMALL, h: ID_H, rows: 10, cols: 1, labels: ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'], gapHeightRatio: 1, hasGaps: false },
            { id: `${pageIdPrefix}_b1`, type: 'answer', x: X1, y: ANS_START_Y, w: ANS_BLOCK_W, h: ANS_BLOCK_H, rows: ROWS_PER_BLOCK, cols: 4, startQ: 1, labels: LABELS_ANS, gapHeightRatio: 0.6, hasGaps: true },
            { id: `${pageIdPrefix}_b2`, type: 'answer', x: X2, y: ANS_START_Y - 3, w: ANS_BLOCK_W, h: ANS_BLOCK_H, rows: ROWS_PER_BLOCK, cols: 4, startQ: 31, labels: LABELS_ANS, gapHeightRatio: 0.6, hasGaps: true }
        ];
    };

    // --- Offset Detection from Image URL (async helper) ---
    const detectOffsetsFromImage = (imageUrl) => {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);
                const yOffset = detectVerticalOffset(ctx, img.width, img.height);
                const xOffset = detectHorizontalOffset(ctx, img.width, img.height);
                resolve({ xOffset, yOffset });
            };
            img.onerror = () => resolve({ xOffset: 0, yOffset: 0 });
            img.src = imageUrl;
        });
    };

    // --- Template Conversion Helpers ---
    const regionsToTemplate = (regions, imgW, imgH) => {
        const idRegions = [];
        const answerRegions = [];
        regions.forEach(r => {
            const base = {
                xRatio: r.x / imgW,
                yRatio: r.y / imgH,
                wRatio: r.w / imgW,
                hRatio: r.h / imgH,
                rows: r.rows,
                cols: r.cols,
                labels: [...r.labels],
                gapHeightRatio: r.gapHeightRatio,
                hasGaps: r.hasGaps,
            };
            if (r.type === 'id') {
                base.suffix = r.id.split('_id_')[1];
                idRegions.push(base);
            } else if (r.type === 'answer') {
                base.startQ = r.startQ;
                answerRegions.push(base);
            }
        });
        return { idRegions, answerRegions };
    };

    const templateToRegions = (template, imgW, imgH, pageIdPrefix) => {
        const regions = [];
        template.idRegions.forEach((ir, idx) => {
            // Use stored suffix, or fall back to index-based name to avoid collisions
            const suffix = ir.suffix || `id${idx + 1}`;
            regions.push({
                id: `${pageIdPrefix}_id_${suffix}`,
                type: 'id',
                x: imgW * ir.xRatio,
                y: imgH * ir.yRatio,
                w: imgW * ir.wRatio,
                h: imgH * ir.hRatio,
                rows: ir.rows || 10,
                cols: ir.cols || 1,
                labels: (ir.labels && ir.labels.length > 0) ? [...ir.labels] : ['0','1','2','3','4','5','6','7','8','9'],
                gapHeightRatio: ir.gapHeightRatio ?? 1,
                hasGaps: ir.hasGaps ?? false,
            });
        });
        template.answerRegions.forEach((ar, i) => {
            const labels = (ar.labels && ar.labels.length > 0) ? [...ar.labels] : ['A','B','C','D'];
            regions.push({
                id: `${pageIdPrefix}_b${i + 1}`,
                type: 'answer',
                x: imgW * ar.xRatio,
                y: imgH * ar.yRatio,
                w: imgW * ar.wRatio,
                h: imgH * ar.hRatio,
                rows: ar.rows || 30,
                cols: ar.cols || labels.length || 4,
                startQ: ar.startQ || 1,
                labels: labels,
                gapHeightRatio: ar.gapHeightRatio ?? 1,
                hasGaps: ar.hasGaps ?? false,
            });
        });
        return regions;
    };

    // --- File Handling ---
    const handleFileUpload = async (e) => {
        const uploadedFile = e.target.files[0];
        if (!uploadedFile) return;
        setFile(uploadedFile);
        setPages([]);
        setCurrentPageIndex(0);
        setShowResults(false);
        if (uploadedFile.type === 'application/pdf') {
            await processPdf(uploadedFile);
        } else if (uploadedFile.type.startsWith('image/')) {
            processImage(uploadedFile);
        }
    };

    const processImage = (file) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);
                const yOffset = detectVerticalOffset(ctx, img.width, img.height);
                const xOffset = detectHorizontalOffset(ctx, img.width, img.height);
                const detectedLineY = getDetectedLineY(ctx, img.width, img.height);
                const detectedLineX = getDetectedLineX(ctx, img.width, img.height);
                setPages([{
                    id: 0,
                    imageUrl: e.target.result,
                    width: img.width,
                    height: img.height,
                    detectedLineY: detectedLineY > 0 ? detectedLineY : null,
                    detectedLineX: detectedLineX > 0 ? detectedLineX : null,
                    results: {},
                    regions: getStandardRegions(img.width, img.height, 'p0', xOffset, yOffset)
                }]);
                if (containerRef.current) {
                    const initialScale = Math.min(1, containerRef.current.clientWidth / img.width);
                    setScale(initialScale);
                }
            };
            img.src = e.target.result;
        };
        reader.readAsDataURL(file);
    };

    const processPdf = async (file) => {
        const pdfjsLib = window.pdfjsLib || window['pdfjs-dist/build/pdf'];
        if (!pdfjsLib) {
            alert("PDF processing library is still loading. Please wait a moment and try uploading again.");
            return;
        }
        setIsProcessing(true);
        setProgress({ current: 0, total: 0 }); // Reset progress
        try {
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
            const totalPages = pdf.numPages;
            const loadedPages = [];

            // Initialize progress
            setProgress({ current: 0, total: totalPages });

            for (let i = 1; i <= totalPages; i++) {
                // Update progress
                setProgress({ current: i, total: totalPages });

                const page = await pdf.getPage(i);
                const viewport = page.getViewport({ scale: 2.0 });
                const canvas = document.createElement('canvas');
                const context = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;
                await page.render({ canvasContext: context, viewport: viewport }).promise;
                const yOffset = detectVerticalOffset(context, canvas.width, canvas.height);
                const xOffset = detectHorizontalOffset(context, canvas.width, canvas.height);
                const detectedLineY = getDetectedLineY(context, canvas.width, canvas.height);
                const detectedLineX = getDetectedLineX(context, canvas.width, canvas.height);
                loadedPages.push({
                    id: i - 1,
                    imageUrl: canvas.toDataURL(),
                    width: canvas.width,
                    height: canvas.height,
                    detectedLineY: detectedLineY > 0 ? detectedLineY : null,
                    detectedLineX: detectedLineX > 0 ? detectedLineX : null,
                    results: {},
                    regions: getStandardRegions(canvas.width, canvas.height, `p${i - 1}`, xOffset, yOffset)
                });
            }
            setPages(loadedPages);
            if (loadedPages.length > 0) {
                if (containerRef.current) {
                    const initialScale = Math.min(1, containerRef.current.clientWidth / loadedPages[0].width);
                    setScale(initialScale);
                }
            }
        } catch (error) {
            console.error(error);
            alert("Failed to process PDF.");
        } finally {
            setIsProcessing(false);
        }
    };

    // --- Template CRUD ---
    const saveAsTemplate = async () => {
        if (!currentPage || !newTemplateName.trim()) return;
        const { idRegions, answerRegions } = regionsToTemplate(
            currentPage.regions, currentPage.width, currentPage.height
        );
        // Capture source page offsets so we can re-align when applying to other pages
        const { xOffset, yOffset } = await detectOffsetsFromImage(currentPage.imageUrl, currentPage.width, currentPage.height);
        const newTemplate = {
            id: Date.now().toString(),
            name: newTemplateName.trim(),
            createdAt: new Date().toISOString(),
            idRegions,
            answerRegions,
            sourceXOffset: xOffset,
            sourceYOffset: yOffset,
        };
        setTemplates(prev => [...prev, newTemplate]);
        setLastAppliedTemplateId(newTemplate.id);
        setNewTemplateName('');
        setShowSaveTemplateModal(false);
        setToast(`Template "${newTemplate.name}" saved!`);
        setTimeout(() => setToast(null), 3000);
    };

    const applyTemplate = async (template) => {
        setIsProcessing(true);
        setProgress({ current: 0, total: pages.length });
        const sourceX = template.sourceXOffset || 0;
        const sourceY = template.sourceYOffset || 0;
        const tplUseY = template.useAlignmentY !== false; // default true
        const tplUseX = template.useAlignmentX !== false; // default true
        const updatedPages = [];
        for (let i = 0; i < pages.length; i++) {
            setProgress({ current: i + 1, total: pages.length });
            const page = pages[i];
            const { xOffset, yOffset } = await detectOffsetsFromImage(page.imageUrl, page.width, page.height);
            // Differential shift: how much this page's content moved vs. the source page
            // Only shift on axes where alignment is enabled
            const dx = tplUseX ? (xOffset - sourceX) : 0;
            const dy = tplUseY ? (yOffset - sourceY) : 0;
            const baseRegions = templateToRegions(template, page.width, page.height, `p${page.id}`);
            const shiftedRegions = baseRegions.map(r => ({ ...r, x: r.x + dx, y: r.y + dy }));
            updatedPages.push({ ...page, regions: shiftedRegions, results: {} });
        }
        setPages(updatedPages);
        setIsProcessing(false);
        setShowResults(false);
        setLastAppliedTemplateId(template.id);
        setToast(`Template "${template.name}" applied & aligned to all pages!`);
        setTimeout(() => setToast(null), 3000);
    };

    // Auto-apply last used template (or default) whenever new pages are loaded
    const prevPagesRef = useRef(null);
    useEffect(() => {
        // Only auto-apply on initial load (when going from 0 pages → N pages)
        const prev = prevPagesRef.current;
        if (prev != null && prev === 0 && pages.length > 0) {
            const tplId = lastAppliedTemplateId || DEFAULT_TEMPLATE_ID;
            const tpl = tplId === DEFAULT_TEMPLATE_ID
                ? getDefaultTemplate()
                : templates.find(t => t.id === tplId);
            if (tpl) {
                // Apply silently without showing toast
                (async () => {
                    setIsProcessing(true);
                    setProgress({ current: 0, total: pages.length });
                    const sourceX = tpl.sourceXOffset || 0;
                    const sourceY = tpl.sourceYOffset || 0;
                    const useY = tpl.useAlignmentY !== false;
                    const useX = tpl.useAlignmentX !== false;
                    const updatedPages = [];
                    for (let i = 0; i < pages.length; i++) {
                        setProgress({ current: i + 1, total: pages.length });
                        const page = pages[i];
                        const { xOffset, yOffset } = await detectOffsetsFromImage(page.imageUrl, page.width, page.height);
                        const dx = useX ? (xOffset - sourceX) : 0;
                        const dy = useY ? (yOffset - sourceY) : 0;
                        const baseRegions = templateToRegions(tpl, page.width, page.height, `p${page.id}`);
                        updatedPages.push({ ...page, regions: baseRegions.map(r => ({ ...r, x: r.x + dx, y: r.y + dy })), results: {} });
                    }
                    setPages(updatedPages);
                    setIsProcessing(false);
                })();
            }
        }
        prevPagesRef.current = pages.length;
    }, [pages.length]);

    // Edit existing template in draw mode
    const editTemplate = (template) => {
        if (!currentPage) return;
        const allRegions = [
            ...(template.idRegions || []).map((r, i) => ({ ...r, blockType: 'id', uid: `edit_id_${Date.now()}_${i}` })),
            ...(template.answerRegions || []).map((r, i) => ({ ...r, blockType: 'answer', uid: `edit_ans_${Date.now()}_${i}` })),
        ];
        // Convert decimal ratios (0-1) to percentage (0-100) and load
        const loadedBlocks = allRegions.map(r => ({
            uid: r.uid,
            xRatio: r.xRatio * 100,
            yRatio: r.yRatio * 100,
            wRatio: r.wRatio * 100,
            hRatio: r.hRatio * 100,
            blockType: r.blockType,
            rows: r.rows || 5,
            cols: r.cols || 4,
            labels: r.labels || ['A','B','C','D'],
            startQ: r.startQ || 1,
            suffix: r.suffix || '',
            hasGaps: r.hasGaps ?? false,
            gapHeightRatio: r.gapHeightRatio ?? 1,
        }));
        // Compute next Q
        let maxQ = 1;
        loadedBlocks.forEach(b => {
            if (b.blockType === 'answer') {
                let qCount = 0;
                for (let r = 0; r < b.rows; r++) if (!b.hasGaps || (r+1)%6!==0) qCount++;
                const endQ = (b.startQ || 1) + qCount;
                if (endQ > maxQ) maxQ = endQ;
            }
        });
        // Enter draw mode with loaded blocks
        setDrawMode(true);
        setDrawnBlocks(loadedBlocks);
        setCurrentDraw(null);
        setSelectedDrawnBlock(null);
        setDrawNextQ(maxQ);
        setDrawTemplateMode('naming'); // skip straight to naming since we're editing
        setDrawTemplateName(template.name);
        setSelectedRegionId(null);
        // Track which template we're editing so save overwrites instead of duplicating
        setEditingTemplateId(template.id);
        setDrawLineY(currentPage.detectedLineY || Math.round(currentPage.height * 0.227));
        setDrawLineX(currentPage.detectedLineX || Math.round(currentPage.width * 0.355));
        setIsDraggingLineY(false);
        setIsDraggingLineX(false);
        setUseLineY(template.useAlignmentY !== false);
        setUseLineX(template.useAlignmentX !== false);
    };

    const deleteTemplate = (templateId) => {
        const tpl = templates.find(t => t.id === templateId);
        setTemplates(prev => prev.filter(t => t.id !== templateId));
        setToast(`Template${tpl ? ` "${tpl.name}"` : ''} deleted.`);
        setTimeout(() => setToast(null), 3000);
    };

    // --- Canvas Drawing ---
    useEffect(() => {
        if (!currentPage || !canvasRef.current) return;
        const canvas = canvasRef.current;
        const ctx = canvas.getContext('2d');
        const img = new Image();
        img.onload = () => {
            canvas.width = img.width * scale;
            canvas.height = img.height * scale;
            ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

            // In draw mode, dim existing regions and draw drawn blocks instead
            if (drawMode) {
                // Dim the image slightly
                ctx.fillStyle = 'rgba(0, 0, 0, 0.3)';
                ctx.fillRect(0, 0, canvas.width, canvas.height);

                // Draw existing regions very faintly
                const regionsToDraw = currentPage.regions || [];
                regionsToDraw.forEach(region => {
                    ctx.strokeStyle = 'rgba(125, 133, 144, 0.15)';
                    ctx.lineWidth = 1;
                    ctx.strokeRect(region.x * scale, region.y * scale, region.w * scale, region.h * scale);
                });

                // Draw drawn blocks
                drawnBlocks.forEach(block => {
                    const x = (currentPage.width * block.xRatio / 100) * scale;
                    const y = (currentPage.height * block.yRatio / 100) * scale;
                    const w = (currentPage.width * block.wRatio / 100) * scale;
                    const h = (currentPage.height * block.hRatio / 100) * scale;

                    const isSelected = block.uid === selectedDrawnBlock;
                    if (block.blockType === 'id') {
                        ctx.strokeStyle = isSelected ? '#58a6ff' : '#1f6feb';
                        ctx.fillStyle = isSelected ? 'rgba(31, 111, 235, 0.15)' : 'rgba(31, 111, 235, 0.08)';
                    } else {
                        ctx.strokeStyle = isSelected ? '#3fb950' : '#238636';
                        ctx.fillStyle = isSelected ? 'rgba(35, 134, 54, 0.15)' : 'rgba(35, 134, 54, 0.08)';
                    }
                    ctx.lineWidth = isSelected ? 3 : 2;
                    ctx.fillRect(x, y, w, h);
                    ctx.strokeRect(x, y, w, h);

                    // Draw label
                    ctx.fillStyle = isSelected ? '#fff' : '#e6edf3';
                    ctx.font = 'bold 14px sans-serif';
                    const labelText = block.blockType === 'id'
                        ? `ID: ${block.suffix || '?'}`
                        : `Q${block.startQ}+ (${block.rows}r×${block.cols}c)`;
                    // Background for text
                    const textW = ctx.measureText(labelText).width;
                    ctx.fillStyle = isSelected ? 'rgba(31, 111, 235, 0.9)' : 'rgba(0,0,0,0.7)';
                    ctx.fillRect(x, y - 20, textW + 10, 18);
                    ctx.fillStyle = '#fff';
                    ctx.fillText(labelText, x + 5, y - 6);
                });

                // Draw the detected/adjusted HORIZONTAL alignment line
                if (useLineY && drawLineY != null) {
                    const lineYScaled = drawLineY * scale;
                    ctx.strokeStyle = isDraggingLineY ? '#d29922' : '#f0883e';
                    ctx.lineWidth = isDraggingLineY ? 4 : 2;
                    ctx.setLineDash([10, 6]);
                    ctx.beginPath();
                    ctx.moveTo(0, lineYScaled);
                    ctx.lineTo(canvas.width, lineYScaled);
                    ctx.stroke();
                    ctx.setLineDash([]);
                    // Label
                    const labelText = '⟷ H-Line (drag to adjust)';
                    ctx.font = 'bold 12px sans-serif';
                    const textW = ctx.measureText(labelText).width;
                    ctx.fillStyle = isDraggingLineY ? 'rgba(210, 153, 34, 0.95)' : 'rgba(240, 136, 62, 0.9)';
                    ctx.fillRect(canvas.width - textW - 14, lineYScaled - 22, textW + 10, 18);
                    ctx.fillStyle = '#fff';
                    ctx.fillText(labelText, canvas.width - textW - 9, lineYScaled - 8);
                }

                // Draw the detected/adjusted VERTICAL alignment line
                if (useLineX && drawLineX != null) {
                    const lineXScaled = drawLineX * scale;
                    ctx.strokeStyle = isDraggingLineX ? '#d29922' : '#a371f7';
                    ctx.lineWidth = isDraggingLineX ? 4 : 2;
                    ctx.setLineDash([10, 6]);
                    ctx.beginPath();
                    ctx.moveTo(lineXScaled, 0);
                    ctx.lineTo(lineXScaled, canvas.height);
                    ctx.stroke();
                    ctx.setLineDash([]);
                    // Label (rotated vertically alongside the line)
                    const labelText = '⟷ V-Line (drag)';
                    ctx.font = 'bold 12px sans-serif';
                    const textW = ctx.measureText(labelText).width;
                    ctx.fillStyle = isDraggingLineX ? 'rgba(210, 153, 34, 0.95)' : 'rgba(163, 113, 247, 0.9)';
                    ctx.fillRect(lineXScaled + 6, 8, textW + 10, 18);
                    ctx.fillStyle = '#fff';
                    ctx.fillText(labelText, lineXScaled + 11, 21);
                }

                // Draw current rectangle being drawn
                if (currentDraw) {
                    const x = Math.min(currentDraw.startX, currentDraw.endX) * scale;
                    const y = Math.min(currentDraw.startY, currentDraw.endY) * scale;
                    const w = Math.abs(currentDraw.endX - currentDraw.startX) * scale;
                    const h = Math.abs(currentDraw.endY - currentDraw.startY) * scale;
                    ctx.strokeStyle = '#d29922';
                    ctx.fillStyle = 'rgba(210, 153, 34, 0.1)';
                    ctx.lineWidth = 2;
                    ctx.setLineDash([6, 4]);
                    ctx.fillRect(x, y, w, h);
                    ctx.strokeRect(x, y, w, h);
                    ctx.setLineDash([]);
                }
                return;
            }

            const regionsToDraw = currentPage.regions || [];
            regionsToDraw.forEach(region => {
                // GitHub Focus Blue for selection
                ctx.strokeStyle = region.id === selectedRegionId ? '#1f6feb' : '#da3633';
                ctx.lineWidth = 2;
                const x = region.x * scale;
                const y = region.y * scale;
                const w = region.w * scale;
                const h = region.h * scale;
                ctx.strokeRect(x, y, w, h);
                const rowLayouts = calculateRowLayout(h, region.rows, region.gapHeightRatio, region.hasGaps);
                ctx.lineWidth = 1;
                const cellW = w / region.cols;
                rowLayouts.forEach((row, idx) => {
                    const rowY = y + row.y;
                    const rowH = row.h;
                    if (row.isGap) {
                        ctx.fillStyle = 'rgba(110, 118, 129, 0.2)'; // Muted gray
                        ctx.fillRect(x, rowY, w, rowH);
                    } else {
                        if (idx > 0) {
                            ctx.strokeStyle = region.id === selectedRegionId ? 'rgba(31, 111, 235, 0.5)' : 'rgba(218, 54, 51, 0.3)';
                            ctx.beginPath();
                            ctx.moveTo(x, rowY);
                            ctx.lineTo(x + w, rowY);
                            ctx.stroke();
                        }
                    }
                });
                ctx.beginPath();
                ctx.strokeStyle = region.id === selectedRegionId ? 'rgba(31, 111, 235, 0.5)' : 'rgba(218, 54, 51, 0.3)';
                for (let i = 1; i < region.cols; i++) {
                    ctx.moveTo(x + i * cellW, y);
                    ctx.lineTo(x + i * cellW, y + h);
                }
                ctx.stroke();
                ctx.fillStyle = ctx.strokeStyle;
                ctx.font = 'bold 12px sans-serif';
                if (region.type === 'id') {
                    let label = "ID";
                    if (region.id.includes('level')) label = "Form";
                    if (region.id.includes('letter')) label = "Cls";
                    if (region.id.includes('n1')) label = "Ten";
                    if (region.id.includes('n2')) label = "Unit";
                    ctx.fillText(label, x, y - 4);
                } else {
                    let qCount = 0;
                    for (let r = 0; r < region.rows; r++) if ((r + 1) % 6 !== 0) qCount++;
                    ctx.fillText(`Q${region.startQ} - Q${region.startQ + qCount - 1}`, x, y - 8);
                }
                if (currentPage.results && currentPage.results[region.id]) {
                    currentPage.results[region.id].forEach((ans) => {
                        const drawRowIndex = (ans.rowIndex !== undefined) ? ans.rowIndex : ans.qNum;
                        const rLayout = rowLayouts[drawRowIndex];
                        if (!rLayout) return;
                        if (region.type === 'id') {
                            if (ans.label === 'BLANK') return;
                            let fillStyle = 'rgba(31, 111, 235, 0.5)'; // Blue
                            if (ans.label === 'MULT') fillStyle = 'rgba(210, 153, 34, 0.4)'; // Orange
                            ctx.fillStyle = fillStyle;
                            const padX = w * 0.10;
                            const padY = rLayout.h * 0.10;
                            ctx.fillRect(x + padX, y + rLayout.y + padY, w - (padX * 2), rLayout.h - (padY * 2));
                            return;
                        }
                        const correctAns = parsedAnswerKey[ans.qNum];
                        let fillStyle = 'rgba(35, 134, 54, 0.5)'; // Green Success
                        if (correctAns) {
                            if (ans.label === correctAns) fillStyle = 'rgba(35, 134, 54, 0.6)'; // Green
                            else fillStyle = 'rgba(248, 81, 73, 0.6)'; // Red Danger
                        }
                        if (ans.detectedIndex !== -1 && ans.label !== 'MULT' && ans.label !== 'BLANK') {
                            const bubbleX = x + (ans.detectedIndex * cellW);
                            const bubbleY = y + rLayout.y;
                            ctx.fillStyle = fillStyle;
                            ctx.fillRect(bubbleX + 2, bubbleY + 2, cellW - 4, rLayout.h - 4);
                        } else if (ans.label === 'MULT') {
                            const bubbleY = y + rLayout.y;
                            ctx.fillStyle = 'rgba(210, 153, 34, 0.4)'; // Orange
                            ctx.fillRect(x, bubbleY, w, rLayout.h);
                        }
                    });
                }
            });
        };
        img.src = currentPage.imageUrl;
    }, [currentPage, scale, selectedRegionId, parsedAnswerKey, drawMode, drawnBlocks, currentDraw, selectedDrawnBlock, drawLineY, drawLineX, isDraggingLineY, isDraggingLineX, useLineY, useLineX]);

    // --- Mouse Interactions ---
    const [isDragging, setIsDragging] = useState(false);
    const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
    const [activeRegionStart, setActiveRegionStart] = useState({ x: 0, y: 0 });

    const handleMouseDown = (e) => {
        if (!currentPage) return;
        const rect = canvasRef.current.getBoundingClientRect();
        const x = (e.clientX - rect.left) / scale;
        const y = (e.clientY - rect.top) / scale;

        if (drawMode) {
            const tolerance = Math.max(6, 12 / scale);
            // First: check if clicking near the HORIZONTAL alignment line
            if (useLineY && drawLineY != null) {
                if (Math.abs(y - drawLineY) <= tolerance) {
                    setIsDraggingLineY(true);
                    setSelectedDrawnBlock(null);
                    setCurrentDraw(null);
                    return;
                }
            }
            // Then: check if clicking near the VERTICAL alignment line
            if (useLineX && drawLineX != null) {
                if (Math.abs(x - drawLineX) <= tolerance) {
                    setIsDraggingLineX(true);
                    setSelectedDrawnBlock(null);
                    setCurrentDraw(null);
                    return;
                }
            }
            // Check if clicking on an existing drawn block
            const clickedBlock = drawnBlocks.find(b => {
                const bx = currentPage.width * b.xRatio / 100;
                const by = currentPage.height * b.yRatio / 100;
                const bw = currentPage.width * b.wRatio / 100;
                const bh = currentPage.height * b.hRatio / 100;
                return x >= bx && x <= bx + bw && y >= by && y <= by + bh;
            });
            if (clickedBlock) {
                setSelectedDrawnBlock(clickedBlock.uid);
                // Allow dragging the block
                setIsDragging(true);
                setDragStart({ x, y });
                setActiveRegionStart({
                    x: currentPage.width * clickedBlock.xRatio / 100,
                    y: currentPage.height * clickedBlock.yRatio / 100,
                });
                setSelectedDrawnBlockId(clickedBlock.uid);
                return;
            }
            // Start drawing a new rectangle
            setSelectedDrawnBlock(null);
            setCurrentDraw({ startX: x, startY: y, endX: x, endY: y });
            return;
        }

        const regionsToCheck = currentPage.regions || [];
        const clickedRegion = regionsToCheck.find(r => x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h);
        if (clickedRegion) {
            setSelectedRegionId(clickedRegion.id);
            setIsDragging(true);
            setDragStart({ x, y });
            setActiveRegionStart({ x: clickedRegion.x, y: clickedRegion.y });
        } else {
            setSelectedRegionId(null);
        }
    };

    const [selectedDrawnBlockId, setSelectedDrawnBlockId] = useState(null);

    const handleMouseMove = (e) => {
        const rect = canvasRef.current.getBoundingClientRect();
        const currentX = (e.clientX - rect.left) / scale;
        const currentY = (e.clientY - rect.top) / scale;

        if (drawMode && isDraggingLineY) {
            // Drag the horizontal alignment line
            setDrawLineY(Math.max(0, Math.min(currentPage.height, currentY)));
            return;
        }

        if (drawMode && isDraggingLineX) {
            // Drag the vertical alignment line
            setDrawLineX(Math.max(0, Math.min(currentPage.width, currentX)));
            return;
        }

        if (drawMode && currentDraw) {
            // Update drawing rectangle
            setCurrentDraw(prev => ({ ...prev, endX: currentX, endY: currentY }));
            return;
        }

        if (drawMode && isDragging && selectedDrawnBlockId) {
            // Drag existing drawn block
            const dx = currentX - dragStart.x;
            const dy = currentY - dragStart.y;
            setDrawnBlocks(prev => prev.map(b => {
                if (b.uid !== selectedDrawnBlockId) return b;
                return {
                    ...b,
                    xRatio: ((activeRegionStart.x + dx) / currentPage.width) * 100,
                    yRatio: ((activeRegionStart.y + dy) / currentPage.height) * 100,
                };
            }));
            return;
        }

        if (isDragging && selectedRegionId) {
            const dx = currentX - dragStart.x;
            const dy = currentY - dragStart.y;
            setPages(prevPages => prevPages.map((p, idx) => {
                if (idx === currentPageIndex) {
                    const updatedRegions = p.regions.map(r => {
                        if (r.id === selectedRegionId) {
                            return { ...r, x: activeRegionStart.x + dx, y: activeRegionStart.y + dy };
                        }
                        return r;
                    });
                    return { ...p, regions: updatedRegions, results: {} };
                }
                return p;
            }));
        }
    };

    const handleMouseUp = () => {
        if (drawMode && (isDraggingLineY || isDraggingLineX)) {
            setIsDraggingLineY(false);
            setIsDraggingLineX(false);
            return;
        }
        if (drawMode && currentDraw && currentPage) {
            // Finalize drawn rectangle
            const x = Math.min(currentDraw.startX, currentDraw.endX);
            const y = Math.min(currentDraw.startY, currentDraw.endY);
            const w = Math.abs(currentDraw.endX - currentDraw.startX);
            const h = Math.abs(currentDraw.endY - currentDraw.startY);
            if (w > 10 && h > 10) {
                const newBlock = {
                    uid: Date.now() + Math.random(),
                    xRatio: (x / currentPage.width) * 100,
                    yRatio: (y / currentPage.height) * 100,
                    wRatio: (w / currentPage.width) * 100,
                    hRatio: (h / currentPage.height) * 100,
                    blockType: 'answer',
                    rows: 5, cols: 4,
                    labels: ['A','B','C','D'],
                    startQ: drawNextQ,
                    suffix: '',
                    hasGaps: false,
                    gapHeightRatio: 1,
                };
                // Auto-count questions
                let qCount = 0;
                for (let r = 0; r < newBlock.rows; r++) if (!newBlock.hasGaps || (r+1) % 6 !== 0) qCount++;
                setDrawNextQ(prev => prev + qCount);
                setDrawnBlocks(prev => [...prev, newBlock]);
                setSelectedDrawnBlock(newBlock.uid);
            }
            setCurrentDraw(null);
            return;
        }
        setIsDragging(false);
    };

    const runBatchDetection = async () => {
        setIsProcessing(true);
        const newPages = [...pages];
        const totalPages = newPages.length;
        setProgress({ current: 0, total: totalPages });

        const processSinglePage = (imgSrc, pageObj) => {
            return new Promise((resolve) => {
                const img = new Image();
                img.onload = () => {
                    const canvas = document.createElement('canvas');
                    canvas.width = img.width;
                    canvas.height = img.height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0);
                    const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
                    const grayData = new Uint8Array(canvas.width * canvas.height);
                    for (let i = 0; i < pixels.data.length; i += 4) {
                        const r = pixels.data[i];
                        const g = pixels.data[i + 1];
                        const b = pixels.data[i + 2];
                        grayData[i / 4] = 0.299 * r + 0.587 * g + 0.114 * b;
                    }
                    const pageResults = {};
                    const pageRegions = pageObj.regions || [];
                    pageRegions.forEach(region => {
                        const cellW = region.w / region.cols;
                        const rowLayouts = calculateRowLayout(region.h, region.rows, region.gapHeightRatio, region.hasGaps);
                        if (region.type === 'id') {
                            const regionResults = [];
                            const rowScores = [];

                            // 1. Collect scores for all rows
                            rowLayouts.forEach((rowConfig, rowIndex) => {
                                const cx = Math.floor(region.x);
                                const cy = Math.floor(region.y + rowConfig.y);
                                const cw = Math.floor(cellW);
                                const ch = Math.floor(rowConfig.h);
                                // REDUCED PADDING for ID detection (10%) to catch edge marks
                                const paddingX = cw * 0.10;
                                const paddingY = ch * 0.10;
                                let darkPixelCount = 0;
                                let totalPixelCount = 0;
                                for (let py = cy + paddingY; py < cy + ch - paddingY; py++) {
                                    for (let px = cx + paddingX; px < cx + cw - paddingX; px++) {
                                        if (px < canvas.width && py < canvas.height) {
                                            const val = grayData[Math.floor(py) * canvas.width + Math.floor(px)];
                                            if (val < 220) darkPixelCount++;
                                            totalPixelCount++;
                                        }
                                    }
                                }
                                const fillRatio = totalPixelCount > 0 ? darkPixelCount / totalPixelCount : 0;
                                rowScores.push({ rowIndex, fillRatio, label: region.labels[rowIndex] });
                            });

                            // 2. Sort by fill ratio descending
                            rowScores.sort((a, b) => b.fillRatio - a.fillRatio);

                            // NEW LOGIC: Candidate System for ID
                            const MIN_ID_THRESHOLD = 0.45; // Lowered to 5% to catch very light marks
                            const candidates = rowScores.filter(s => s.fillRatio >= MIN_ID_THRESHOLD);

                            let label = 'BLANK';
                            let bestRow = -1;
                            let maxFill = rowScores[0].fillRatio;

                            if (candidates.length === 0) {
                                label = 'BLANK';
                            } else if (candidates.length === 1) {
                                label = candidates[0].label;
                                bestRow = candidates[0].rowIndex;
                            } else {
                                // Multiple candidates found
                                const winner = candidates[0];
                                const runnerUp = candidates[1];

                                // Check if Winner is dominant
                                // If difference is > 5%
                                const diff = winner.fillRatio - runnerUp.fillRatio;

                                if (diff > 0.05) {
                                    label = winner.label;
                                    bestRow = winner.rowIndex;
                                } else {
                                    label = 'MULT'; // Truly ambiguous
                                }
                            }

                            regionResults.push({ qNum: bestRow, rowIndex: bestRow, detectedIndex: 0, label: label, confidence: maxFill });
                            pageResults[region.id] = regionResults;
                        } else {
                            // ... (Answer scanning logic remains unchanged as requested)
                            const regionResults = [];
                            let validQuestionCount = 0;
                            rowLayouts.forEach((rowConfig, rowIndex) => {
                                if (rowConfig.isGap) return;
                                const colScores = [];
                                for (let c = 0; c < region.cols; c++) {
                                    const cx = Math.floor(region.x + c * cellW);
                                    const cy = Math.floor(region.y + rowConfig.y);
                                    const cw = Math.floor(cellW);
                                    const ch = Math.floor(rowConfig.h);
                                    const paddingX = cw * 0.15;
                                    const paddingY = ch * 0.15;
                                    let darkPixelCount = 0;
                                    let totalPixelCount = 0;
                                    for (let py = cy + paddingY; py < cy + ch - paddingY; py++) {
                                        for (let px = cx + paddingX; px < cx + cw - paddingX; px++) {
                                            if (px < canvas.width && py < canvas.height) {
                                                const val = grayData[Math.floor(py) * canvas.width + Math.floor(px)];
                                                if (val < 220) darkPixelCount++;
                                                totalPixelCount++;
                                            }
                                        }
                                    }
                                    const fillRatio = totalPixelCount > 0 ? darkPixelCount / totalPixelCount : 0;
                                    colScores.push({ index: c, fillRatio: fillRatio });
                                }
                                colScores.sort((a, b) => b.fillRatio - a.fillRatio);
                                const maxFill = colScores[0].fillRatio;
                                const minFill = colScores[colScores.length - 1].fillRatio;
                                const secondMaxFill = colScores.length > 1 ? colScores[1].fillRatio : 0;
                                let label = '';
                                let selectedIndex = -1;

                                // Robust Answer Logic — tuned for custom templates
                                if ((maxFill - minFill) < 0.08 || maxFill < 0.35) {
                                    label = 'BLANK';
                                }
                                else if ((maxFill - secondMaxFill) < 0.10) {
                                    label = 'MULT';
                                }
                                else {
                                    selectedIndex = colScores[0].index;
                                    label = region.labels[selectedIndex];
                                }

                                regionResults.push({ qNum: region.startQ + validQuestionCount, rowIndex: rowIndex, detectedIndex: selectedIndex, label: label, confidence: maxFill });
                                validQuestionCount++;
                            });
                            pageResults[region.id] = regionResults;
                        }
                    });
                    resolve(pageResults);
                };
                img.src = imgSrc;
            });
        };

        for (let i = 0; i < newPages.length; i++) {
            setProgress(prev => ({ ...prev, current: i + 1 }));
            const results = await processSinglePage(newPages[i].imageUrl, newPages[i]);
            newPages[i] = { ...newPages[i], results: results };
        }

        // Apply filtering based on answer key or auto-detect cutoff
        const keyCount = Object.keys(parsedAnswerKey).length;
        if (keyCount > 0) {
            // Use answer key length
            newPages.forEach(page => {
                Object.keys(page.results).forEach(key => {
                    if (/_b\d+$/.test(key)) page.results[key] = page.results[key].filter(r => r.qNum <= keyCount);
                });
            });
        }
        // NOTE: Auto-cutoff based on blank-detection has been removed.
        // It was causing all answers except Q1 to be filtered out when scanning
        // sheets with many blank responses (which is normal for MC sheets).
        // To filter to a specific question count, provide an Answer Key —
        // its length is used as the cutoff.

        setPages(newPages);
        setIsProcessing(false);
        setShowResults(true);

        // Auto-jump to first page with incomplete student ID
        const incompleteIndices = [];
        newPages.forEach((page, index) => {
            const name = getPageName(page, index);
            if (name.startsWith('Page ')) {
                incompleteIndices.push(index);
            }
        });

        if (incompleteIndices.length > 0) {
            setCurrentPageIndex(incompleteIndices[0]);
            setShowIdModal(true);
            setIdModalIdx(0);
        }
    };

    const exportExcel = () => {
        if (!window.XLSX) { alert("Excel library not ready."); return; }

        // Sort pages
        const sortedPages = [...pages].sort((a, b) => getPageName(a, a.id).localeCompare(getPageName(b, b.id), undefined, { numeric: true }));
        let maxQ = 0;

        // 1. Prepare Main Data (Student Rows)
        const resultsData = sortedPages.map((page, index) => {
            const pageName = getPageName(page, index);
            const rowData = { 'Student': pageName };
            const pageRegions = page.regions || [];
            pageRegions.filter(r => r.type === 'answer').sort((a, b) => a.startQ - b.startQ).forEach(region => {
                if (page.results && page.results[region.id]) {
                    page.results[region.id].forEach(row => {
                        rowData[`Q${row.qNum}`] = row.label;
                        if (row.qNum > maxQ) maxQ = row.qNum;
                    });
                }
            });
            return rowData;
        });

        const wb = window.XLSX.utils.book_new();
        const wsResults = window.XLSX.utils.json_to_sheet(resultsData);

        // Apply AutoFilter to the Student Data range (Headers + Student Rows)
        // This runs before adding the footer, so the range '!ref' covers exactly what we want to filter.
        if (resultsData.length > 0 && wsResults['!ref']) {
            wsResults['!autofilter'] = { ref: wsResults['!ref'] };
        }

        // 2. Footer Stats (using Excel formulas)
        const studentCount = resultsData.length;
        const startRow = 2;
        const endRow = 1 + studentCount;

        const footerRows = [
            [], [],
            ["Marks"],
            ["Total mark"],
            ["Percentage"],
            ["Answer"],
            ["A"], ["B"], ["C"], ["D"]
        ];

        const IDX_MARKS = 2;
        const IDX_AVG = 3;
        const IDX_PCT = 4;
        const IDX_KEY = 5;
        const IDX_A = 6;
        const IDX_B = 7;
        const IDX_C = 8;
        const IDX_D = 9;

        const rowNum_Marks = endRow + 3;
        const rowNum_Key = endRow + 6;

        let totalFullMark = 0;

        for (let q = 1; q <= maxQ; q++) {
            const colLetter = getExcelCol(q);
            const range = `${colLetter}${startRow}:${colLetter}${endRow}`;
            const qMark = getQuestionMark(q);
            totalFullMark += qMark;

            footerRows[IDX_MARKS].push(qMark);

            const key = parsedAnswerKey[q] || "-";
            footerRows[IDX_KEY].push(key);

            const keyRef = `${colLetter}${rowNum_Key}`;
            const avgFormula = `COUNTIF(${range},${keyRef})/${studentCount}`;
            footerRows[IDX_AVG].push({ t: 'n', f: avgFormula, z: '0.00' });

            const marksRef = `${colLetter}${rowNum_Marks}`;
            const ratioFormula = `COUNTIF(${range},${keyRef})/${studentCount}`;

            footerRows[IDX_AVG][footerRows[IDX_AVG].length - 1] = { t: 'n', f: `${ratioFormula}*${marksRef}`, z: '0.00' };

            footerRows[IDX_PCT].push({ t: 'n', f: ratioFormula, z: '0.0%' });

            footerRows[IDX_A].push({ t: 'n', f: `COUNTIF(${range},"A")/${studentCount}`, z: '0.0%' });
            footerRows[IDX_B].push({ t: 'n', f: `COUNTIF(${range},"B")/${studentCount}`, z: '0.0%' });
            footerRows[IDX_C].push({ t: 'n', f: `COUNTIF(${range},"C")/${studentCount}`, z: '0.0%' });
            footerRows[IDX_D].push({ t: 'n', f: `COUNTIF(${range},"D")/${studentCount}`, z: '0.0%' });
        }

        window.XLSX.utils.sheet_add_aoa(wsResults, footerRows, { origin: -1 });
        window.XLSX.utils.book_append_sheet(wb, wsResults, "All Results");

        // --- SHEET 2: SCORES (Per Student) ---
        const startColLet = "B";
        const endColLet = getExcelCol(maxQ);
        const refKeyRange = `'All Results'!$${startColLet}$${rowNum_Key}:$${endColLet}$${rowNum_Key}`;
        const refMarksRange = `'All Results'!$${startColLet}$${rowNum_Marks}:$${endColLet}$${rowNum_Marks}`;
        
        // Dynamic Full Mark Reference from Sheet 1 (Row `rowNum_Marks`)
        const refTotalMarks = `SUM(${refMarksRange})`;

        const scoresData = sortedPages.map((page, index) => {
            const studentRowIdx = index + 2; // Data starts at row 2 (1-based)
            const refStudentRange = `'All Results'!${startColLet}${studentRowIdx}:${endColLet}${studentRowIdx}`;

            // ** UPDATE: Student ID is now a dynamic formula referencing the main sheet **
            const idFormula = `'All Results'!A${studentRowIdx}`;

            const scoreFormula = `SUMPRODUCT(--(${refStudentRange}=${refKeyRange}), ${refMarksRange})`;
            const pctFormula = `B${index + 2}/${refTotalMarks}`;

            return {
                'Student': { t: 's', f: idFormula }, // Dynamic ID Formula
                'Score': { t: 'n', f: scoreFormula },
                'Percentage': { t: 'n', f: pctFormula, z: '0.0%' }
            };
        });

        const wsScores = window.XLSX.utils.json_to_sheet(scoresData);

        // Format Percentage Column (Col C) as %
        const range = window.XLSX.utils.decode_range(wsScores['!ref']);
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            const ref = window.XLSX.utils.encode_cell({ r: R, c: 2 }); // Col C is index 2
            if (!wsScores[ref]) continue;
            wsScores[ref].z = '0.0%';
            wsScores[ref].t = 'n';
        }

        window.XLSX.utils.book_append_sheet(wb, wsScores, "Scores");
        
        // --- SHEET 3: STATISTICS (New) ---
        // References to Scores Sheet:
        // Scores are in Col B. Student count is `studentCount`.
        // Range B2:B(studentCount+1)
        const scoreRange = `'Scores'!B2:B${studentCount + 1}`;
        const fullMarkRef = refTotalMarks; 

        const statsData = [
            ["Metric", "Value"],
            ["Full Mark", { t:'n', f: fullMarkRef }], 
            ["Max", { t:'n', f: `MAX(${scoreRange})` }],
            ["Min", { t:'n', f: `MIN(${scoreRange})` }],
            ["Mean", { t:'n', f: `AVERAGE(${scoreRange})`, z:'0.00' }],
            ["Median", { t:'n', f: `MEDIAN(${scoreRange})` }],
            ["SD", { t:'n', f: `STDEVP(${scoreRange})`, z:'0.00' }],
            ["Passing Score (50%)", { t:'n', f: `B2*0.5` }], // B2 is Full Mark
            ["No. of Passes", { t:'n', f: `COUNTIF(${scoreRange},">="&B8)` }], // B8 is Passing Score 50%
            ["Passing Rate", { t:'n', f: `B9/${studentCount}`, z:'0.0%' }],
            ["Passing Score (40%)", { t:'n', f: `B2*0.4` }], 
            ["No. of Passes", { t:'n', f: `COUNTIF(${scoreRange},">="&B11)` }], // B11 is Passing Score 40%
            ["Passing Rate", { t:'n', f: `B12/${studentCount}`, z:'0.0%' }]
        ];
        
        const wsStats = window.XLSX.utils.aoa_to_sheet(statsData);
        
        // Adjust column width for Stats sheet
        wsStats['!cols'] = [{ wch: 25 }, { wch: 15 }];
        
        window.XLSX.utils.book_append_sheet(wb, wsStats, "Statistics");

        window.XLSX.writeFile(wb, `${file ? file.name.replace(/\.[^/.]+$/, "") : "results"}_results.xlsx`);
        setToast("Results exported successfully!");
        setTimeout(() => setToast(null), 3000);
    };

    // --- Marks Settings Component ---
    const renderWeightSettings = () => (
        <div className="space-y-2">
            <button onClick={() => setIsMarksSettingsOpen(!isMarksSettingsOpen)} className="w-full flex items-center justify-between text-xs font-semibold uppercase text-[#7d8590] tracking-wider hover:text-[#e6edf3] bg-[#21262d] p-2 rounded border border-[#30363d] transition-colors">
                <div className="flex items-center gap-2"><Hash className="w-4 h-4" /> Question Weighting</div>
                {isMarksSettingsOpen ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
            </button>
            {isMarksSettingsOpen && (
                <div className="bg-[#161b22] p-2 rounded-lg border border-[#30363d] space-y-2">
                    {weightSections.map((section, idx) => (
                        <div key={section.id} className="flex items-center gap-1 text-xs">
                            <input type="number" min="1" className="w-12 p-1 border border-[#30363d] rounded text-center bg-[#0d1117] text-[#e6edf3]" value={section.start} onChange={(e) => {
                                const valStr = e.target.value;
                                if (valStr === '') {
                                    const newSecs = [...weightSections];
                                    newSecs[idx].start = '';
                                    setWeightSections(newSecs);
                                    return;
                                }
                                let val = parseInt(valStr);
                                if (val < 1) val = 1;
                                const newSecs = [...weightSections];
                                newSecs[idx].start = val;
                                setWeightSections(newSecs);
                            }} placeholder="Start" />
                            <span className="text-[#7d8590]">-</span>
                            <input type="number" min="1" className="w-12 p-1 border border-[#30363d] rounded text-center bg-[#0d1117] text-[#e6edf3]" value={section.end}
                                onChange={(e) => {
                                    const valStr = e.target.value;
                                    if (valStr === '') {
                                        const newSecs = [...weightSections];
                                        newSecs[idx].end = '';
                                        setWeightSections(newSecs);
                                        return;
                                    }
                                    let val = parseInt(valStr);
                                    if (val < 1) val = 1;
                                    const newSecs = [...weightSections];
                                    newSecs[idx].end = val;
                                    setWeightSections(newSecs);
                                }}
                                onBlur={() => {
                                    const newSecs = [...weightSections];
                                    const s = newSecs[idx].start;
                                    const e = newSecs[idx].end;
                                    if (s !== '' && e !== '' && e < s) {
                                        newSecs[idx].end = s;
                                        setWeightSections(newSecs);
                                    }
                                }}
                                placeholder="End" />
                            <span className="text-[#7d8590]">:</span>
                            <input type="number" min="0.1" step="0.5" className="w-10 p-1 border border-[#30363d] rounded text-center bg-[#0d1117] text-[#e6edf3]" value={section.mark} onChange={(e) => {
                                const valStr = e.target.value;
                                if (valStr === '') {
                                    const newSecs = [...weightSections];
                                    newSecs[idx].mark = '';
                                    setWeightSections(newSecs);
                                    return;
                                }
                                let val = parseFloat(valStr);
                                if (val <= 0) val = 1;
                                if (isNaN(val)) val = 1;
                                const newSecs = [...weightSections];
                                newSecs[idx].mark = val;
                                setWeightSections(newSecs);
                            }} placeholder="Mark" />
                            {weightSections.length > 1 && (
                                <button onClick={() => setWeightSections(weightSections.filter(s => s.id !== section.id))} className="text-[#f85149] hover:text-[#ff7b72] bg-transparent"><X className="w-3 h-3" /></button>
                            )}
                        </div>
                    ))}
                    <button onClick={() => {
                        const lastEnd = weightSections[weightSections.length - 1].end;
                        setWeightSections([...weightSections, { id: Date.now(), start: lastEnd + 1, end: lastEnd + 10, mark: 1 }]);
                    }} className="w-full py-1 text-xs bg-[#1f6feb]/10 text-[#58a6ff] rounded border border-[#1f6feb]/30 hover:bg-[#1f6feb]/20 transition-colors">+ Add Range</button>
                </div>
            )}
        </div>
    );

    // --- Template Section UI ---
    const renderTemplateSection = () => (
        <div className="space-y-2">
            <button onClick={() => setIsTemplateSectionOpen(!isTemplateSectionOpen)} className="w-full flex items-center justify-between text-xs font-semibold uppercase text-[#7d8590] tracking-wider hover:text-[#e6edf3] bg-[#21262d] p-2 rounded border border-[#30363d] transition-colors">
                <div className="flex items-center gap-2"><LayoutTemplate className="w-4 h-4" /> Templates</div>
                {isTemplateSectionOpen ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
            </button>

            {isTemplateSectionOpen && (
                <div className="bg-[#161b22] p-2 rounded-lg border border-[#30363d] space-y-2">
                    {allTemplates.length === 0 ? (
                        <p className="text-xs text-[#7d8590] text-center py-2 italic">No templates available.</p>
                    ) : (
                        <div className="space-y-1 max-h-48 overflow-y-auto">
                            {allTemplates.map(t => {
                                const totalQs = t.answerRegions.reduce((sum, ar) => {
                                    let q = 0; for (let r = 0; r < ar.rows; r++) if (!ar.hasGaps || (r+1)%6!==0) q++;
                                    return sum + q;
                                }, 0);
                                const isDefault = t.id === DEFAULT_TEMPLATE_ID;
                                const isLastUsed = t.id === lastAppliedTemplateId;
                                return (
                                    <div key={t.id} className="flex items-center gap-1">
                                        <button
                                            onClick={() => applyTemplate(t)}
                                            disabled={pages.length === 0}
                                            className={`flex-1 text-left px-2 py-1.5 rounded text-xs truncate disabled:opacity-40 disabled:cursor-not-allowed transition-colors flex items-center gap-1 ${isDefault ? 'bg-[#1f6feb]/10 border border-[#1f6feb]/30 text-[#58a6ff] hover:bg-[#1f6feb]/20' : 'bg-[#21262d] border border-[#30363d] hover:bg-[#30363d] text-[#c9d1d9]'}`}
                                            title={`Apply "${t.name}" to all pages${isLastUsed ? ' (currently active)' : ''}`}
                                        >
                                            {isLastUsed && <CheckCircle2 className="w-3 h-3 text-[#3fb950] flex-shrink-0" />}
                                            <span className="font-medium truncate">{t.name}</span>
                                            <span className="text-[#7d8590] ml-auto">({t.idRegions.length} ID, {totalQs}Q)</span>
                                        </button>
                                        {!isDefault && (
                                            <>
                                                <button onClick={() => editTemplate(t)} className="p-1 text-[#58a6ff] hover:text-[#79c0ff] bg-transparent rounded hover:bg-[#1f6feb]/10 transition-colors flex-shrink-0" title="Edit template" disabled={!currentPage}>
                                                    <Settings className="w-3 h-3" />
                                                </button>
                                                <button
                                                    onClick={() => deleteTemplate(t.id)}
                                                    className="p-1.5 text-[#f85149] hover:text-[#ff7b72] bg-transparent rounded hover:bg-[#f85149]/10 transition-colors flex-shrink-0"
                                                    title="Delete template"
                                                >
                                                    <Trash2 className="w-3 h-3" />
                                                </button>
                                            </>
                                        )}
                                    </div>
                                );
                            })}
                        </div>
                    )}
                    {/* Add New Template Button (under the list) */}
                    <button
                        onClick={startDrawMode}
                        disabled={!currentPage}
                        className="w-full py-2 text-xs font-medium text-[#d29922] bg-[#d29922]/10 rounded border border-[#d29922]/30 hover:bg-[#d29922]/20 transition-colors flex justify-center items-center gap-1.5 disabled:opacity-40 disabled:cursor-not-allowed"
                    >
                        <Plus className="w-3 h-3" /> Add New Template
                    </button>
                </div>
            )}
        </div>
    );

    // --- Manual ID Editing ---
    const updateIdResult = (regionId, newLabel) => {
        setPages(prevPages => prevPages.map((p, idx) => {
            if (idx === currentPageIndex) {
                const newResults = { ...p.results };
                const regionResults = [...(newResults[regionId] || [{ qNum: -1, rowIndex: -1, detectedIndex: 0, label: '?', confidence: 0 }])];
                regionResults[0] = { ...regionResults[0], label: newLabel };
                newResults[regionId] = regionResults;
                return { ...p, results: newResults };
            }
            return p;
        }));
    };

    const getIdValueByRegionId = (regionId) => {
        if (!currentPage || !currentPage.results) return '?';
        return getIdValueByRegionIdForPage(currentPageIndex, regionId);
    };

    const getIdValueByRegionIdForPage = (pageIdx, regionId) => {
        const page = pages[pageIdx];
        if (!page || !page.results) return '?';
        const res = page.results[regionId];
        if (!res || res.length === 0) return '?';
        const label = res[0].label;
        if (label && label !== 'BLANK' && label !== 'MULT' && label !== '?') return label;
        return '?';
    };

    const getIncompletePages = () => {
        const result = [];
        pages.forEach((page, index) => {
            if (!page.results) { result.push(index); return; }
            const idRegions = (page.regions || []).filter(r => r.type === 'id');
            if (idRegions.length === 0) { result.push(index); return; }
            const hasIncomplete = idRegions.some(r => {
                const res = page.results[r.id];
                if (!res || res.length === 0) return true;
                const label = res[0].label;
                return !label || label === '?' || label === 'BLANK' || label === 'MULT';
            });
            if (hasIncomplete) result.push(index);
        });
        return result;
    };

    const updateIdResultForPageByRegionId = (pageIdx, regionId, newLabel) => {
        setPages(prevPages => prevPages.map((p, idx) => {
            if (idx === pageIdx) {
                const newResults = { ...p.results };
                const regionResults = [...(newResults[regionId] || [{ qNum: -1, rowIndex: -1, detectedIndex: 0, label: '?', confidence: 0 }])];
                regionResults[0] = { ...regionResults[0], label: newLabel };
                newResults[regionId] = regionResults;
                return { ...p, results: newResults };
            }
            return p;
        }));
    };

    // --- Manual Answer Editing ---
    const cycleAnswer = (regionId, resultIndex) => {
        const cycleOrder = ['A', 'B', 'C', 'D', 'MULT', 'BLANK'];
        setPages(prevPages => prevPages.map((p, idx) => {
            if (idx === currentPageIndex) {
                const newResults = { ...p.results };
                const regionResults = [...newResults[regionId]];
                const current = regionResults[resultIndex];
                const currentLabel = current.label;
                const currentCycleIdx = cycleOrder.indexOf(currentLabel);
                const nextIdx = (currentCycleIdx + 1) % cycleOrder.length;
                const nextLabel = cycleOrder[nextIdx];

                let newDetectedIndex = current.detectedIndex;
                if (['A', 'B', 'C', 'D'].includes(nextLabel)) {
                    newDetectedIndex = ['A', 'B', 'C', 'D'].indexOf(nextLabel);
                } else {
                    newDetectedIndex = -1;
                }

                regionResults[resultIndex] = { ...current, label: nextLabel, detectedIndex: newDetectedIndex };
                newResults[regionId] = regionResults;
                return { ...p, results: newResults };
            }
            return p;
        }));
    };

    // --- UI Components ---
    return (
        <div className="flex h-screen bg-[#0d1117] text-[#e6edf3] font-sans overflow-hidden">
            <style>{`:root, body, #root { height: 100%; width: 100%; margin: 0; padding: 0; max-width: none !important; }`}</style>

            {/* Toast Notification */}
            {toast && (
                <div className="fixed bottom-4 right-4 bg-[#238636] text-white px-4 py-2 rounded shadow-lg flex items-center gap-2 z-50 animate-bounce">
                    <CheckCircle2 className="w-4 h-4 text-white" />
                    {toast}
                </div>
            )}

            {/* Floating Incomplete ID Badge */}
            {showResults && !showIdModal && getIncompletePages().length > 0 && (
                <button onClick={() => { setShowIdModal(true); setIdModalIdx(0); }} className="fixed top-4 right-4 z-40 bg-[#da3633] text-white px-3 py-1.5 rounded-full text-sm font-bold shadow-lg hover:bg-[#f85149] transition-colors flex items-center gap-2 animate-bounce">
                    <UserSquare2 className="w-4 h-4" />
                    {getIncompletePages().length} incomplete ID{getIncompletePages().length > 1 ? 's' : ''}
                </button>
            )}

            {/* ID Review Modal */}
            {showIdModal && (() => {
                const incomplete = getIncompletePages();
                if (incomplete.length === 0) { setShowIdModal(false); return null; }
                const pageIdx = incomplete[Math.min(idModalIdx, incomplete.length - 1)];
                const page = pages[pageIdx];
                return (
                    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm" onClick={() => setShowIdModal(false)}>
                        <div className="bg-[#0d1117] border border-[#30363d] rounded-xl shadow-2xl w-full max-w-lg mx-4" onClick={(e) => e.stopPropagation()}>
                            <div className="px-6 py-4 border-b border-[#30363d] flex justify-between items-center">
                                <h2 className="text-lg font-bold text-[#e6edf3] flex items-center gap-2">
                                    <UserSquare2 className="w-5 h-5 text-[#f85149]" />
                                    Fix Student ID
                                </h2>
                                <span className="text-xs text-[#7d8590] bg-[#21262d] px-2 py-1 rounded-full">{Math.min(idModalIdx + 1, incomplete.length)} of {incomplete.length}</span>
                            </div>
                            <div className="p-6 space-y-4">
                                {page && (
                                    <div className="flex flex-col gap-4">
                                        <p className="text-sm text-[#7d8590]">Page {pageIdx + 1} — Select the correct student ID:</p>
                                        {page.imageUrl && (() => {
                                            // Crop the entire upper portion of the page (above the detected line)
                                            const lineY = page.detectedLineY || Math.round(page.height * 0.25);
                                            const cropW = page.width;
                                            const cropH = lineY;
                                            return (
                                                <div className="w-full border border-[#30363d] rounded overflow-hidden bg-[#161b22]" style={{ aspectRatio: `${cropW}/${cropH}`, maxHeight: '250px', backgroundImage: `url(${page.imageUrl})`, backgroundSize: `100% ${page.height / cropH * 100}%`, backgroundPosition: '0% 0%', backgroundRepeat: 'no-repeat' }} />
                                            );
                                        })()}
                                        <div className="flex items-center gap-2 justify-center flex-wrap">
                                            {getIdFieldsForPage(pageIdx).map((field) => {
                                                const val = getIdValueByRegionIdForPage(pageIdx, field.regionId);
                                                const isIncomplete = val === '?';
                                                return (
                                                    <div key={field.regionId} className="flex flex-col items-center gap-1">
                                                        <span className="text-[10px] text-[#7d8590] uppercase">{field.label}</span>
                                                        <select value={val} onChange={(e) => updateIdResultForPageByRegionId(pageIdx, field.regionId, e.target.value)} className={`text-sm font-bold px-2 py-1.5 rounded border outline-none cursor-pointer ${isIncomplete ? 'bg-[#da3633]/20 border-[#da3633]/40 text-[#f85149]' : 'bg-[#161b22] border-[#238636]/40 text-[#3fb950]'}`}>
                                                            <option value="?" disabled>?</option>
                                                            {field.options.map((opt) => (<option key={opt} value={opt}>{opt}</option>))}
                                                        </select>
                                                    </div>
                                                );
                                            })}
                                        </div>
                                    </div>
                                )}
                            </div>
                            <div className="px-6 py-4 border-t border-[#30363d] flex justify-between items-center">
                                <button onClick={() => setIdModalIdx(prev => Math.max(0, prev - 1))} disabled={idModalIdx === 0} className="px-4 py-2 text-sm rounded-lg border border-[#30363d] bg-[#21262d] text-[#e6edf3] hover:bg-[#30363d] disabled:opacity-30 flex items-center gap-1">
                                    <ChevronLeft className="w-4 h-4" /> Prev
                                </button>
                                <button onClick={() => setShowIdModal(false)} className="px-4 py-2 text-sm rounded-lg border border-[#30363d] bg-[#21262d] text-[#7d8590] hover:text-[#e6edf3] hover:bg-[#30363d]">
                                    Skip
                                </button>
                                {idModalIdx < incomplete.length - 1 ? (
                                    <button onClick={() => setIdModalIdx(prev => prev + 1)} className="px-4 py-2 text-sm rounded-lg bg-[#1f6feb] text-white hover:bg-[#388bfd] flex items-center gap-1">
                                        Next <ChevronRight className="w-4 h-4" />
                                    </button>
                                ) : (
                                    <button onClick={() => setShowIdModal(false)} className="px-4 py-2 text-sm rounded-lg bg-[#238636] text-white hover:bg-[#2ea043] flex items-center gap-1">
                                        <CheckCircle2 className="w-4 h-4" /> Done
                                    </button>
                                )}
                            </div>
                        </div>
                    </div>
                );
            })()}

            {/* Draw Mode Toolbar (top of screen when drawing) */}
            {drawMode && (
                <div className="fixed top-0 left-0 right-0 z-40 bg-[#0d1117] border-b border-[#30363d] px-4 py-2 flex items-center justify-between shadow-lg">
                    <div className="flex items-center gap-3">
                        <LayoutTemplate className="w-5 h-5 text-[#d29922]" />
                        <span className="text-sm font-bold text-[#e6edf3]">Template Builder Mode</span>
                        <span className="text-xs text-[#7d8590] hidden lg:inline">— Click and drag on the page to draw blocks</span>
                        {/* Alignment toggles */}
                        <div className="flex items-center gap-1 ml-2 pl-3 border-l border-[#30363d]">
                            <button
                                onClick={() => setUseLineY(v => !v)}
                                className={`flex items-center gap-1 px-2 py-1 text-xs rounded border transition-colors ${useLineY ? 'bg-[#f0883e]/20 border-[#f0883e]/50 text-[#f0883e]' : 'bg-[#21262d] border-[#30363d] text-[#7d8590] hover:text-[#e6edf3]'}`}
                                title="Toggle horizontal alignment line"
                            >
                                <span className="font-bold">━</span> H-Line
                            </button>
                            <button
                                onClick={() => setUseLineX(v => !v)}
                                className={`flex items-center gap-1 px-2 py-1 text-xs rounded border transition-colors ${useLineX ? 'bg-[#a371f7]/20 border-[#a371f7]/50 text-[#a371f7]' : 'bg-[#21262d] border-[#30363d] text-[#7d8590] hover:text-[#e6edf3]'}`}
                                title="Toggle vertical alignment line"
                            >
                                <span className="font-bold">┃</span> V-Line
                            </button>
                        </div>
                    </div>
                    <div className="flex items-center gap-2">
                        {drawTemplateMode === 'naming' ? (
                            <>
                                <input
                                    type="text"
                                    placeholder="Template name..."
                                    value={drawTemplateName}
                                    onChange={(e) => setDrawTemplateName(e.target.value)}
                                    onKeyDown={(e) => { if (e.key === 'Enter') saveDrawnTemplate(); }}
                                    className="p-1.5 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] outline-none w-48"
                                    autoFocus
                                />
                                <button onClick={saveDrawnTemplate} disabled={!drawTemplateName.trim() || drawnBlocks.length === 0} className="px-3 py-1.5 text-sm rounded-lg bg-[#238636] text-white hover:bg-[#2ea043] disabled:opacity-50 flex items-center gap-1.5">
                                    <Save className="w-3.5 h-3.5" /> Save
                                </button>
                            </>
                        ) : (
                            <>
                                <span className="text-xs text-[#7d8590] bg-[#21262d] px-2 py-1 rounded">{drawnBlocks.length} block(s) drawn</span>
                                <button
                                    onClick={() => setDrawTemplateMode('naming')}
                                    disabled={drawnBlocks.length === 0}
                                    className="px-3 py-1.5 text-sm rounded-lg bg-[#238636] text-white hover:bg-[#2ea043] disabled:opacity-50 flex items-center gap-1.5"
                                >
                                    <Save className="w-3.5 h-3.5" /> Save Template
                                </button>
                            </>
                        )}
                        <button onClick={cancelDrawMode} className="px-3 py-1.5 text-sm rounded-lg border border-[#30363d] bg-[#21262d] text-[#f85149] hover:bg-[#30363d] flex items-center gap-1.5">
                            <X className="w-3.5 h-3.5" /> Cancel
                        </button>
                    </div>
                </div>
            )}

            {/* Draw Mode: Floating Block Config Panel */}
            {drawMode && selectedDrawnBlock && (() => {
                const block = drawnBlocks.find(b => b.uid === selectedDrawnBlock);
                if (!block) return null;
                return (
                    <div className="fixed right-4 top-20 z-40 w-72 bg-[#0d1117] border border-[#30363d] rounded-xl shadow-2xl flex flex-col max-h-[80vh]">
                        <div className="px-4 py-3 border-b border-[#30363d] flex justify-between items-center bg-[#161b22] rounded-t-xl">
                            <span className="text-sm font-bold text-[#e6edf3]">Configure Block</span>
                            <button onClick={() => setSelectedDrawnBlock(null)} className="text-[#7d8590] hover:text-[#e6edf3]">
                                <X className="w-4 h-4" />
                            </button>
                        </div>
                        <div className="p-4 space-y-3 overflow-y-auto">
                            {/* Block Type */}
                            <div>
                                <label className="text-xs text-[#7d8590] block mb-1">Block Type</label>
                                <select value={block.blockType} onChange={(e) => updateDrawnBlock(block.uid, 'blockType', e.target.value)} className="w-full p-1.5 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] outline-none">
                                    <option value="answer">✏️ Answer Block (MC Questions)</option>
                                    <option value="id">📋 ID Block (Student ID)</option>
                                </select>
                            </div>

                            {/* ID Block Fields */}
                            {block.blockType === 'id' && (
                                <>
                                    <div>
                                        <label className="text-xs text-[#7d8590] block mb-1">Suffix (ID name)</label>
                                        <input type="text" value={block.suffix} onChange={(e) => updateDrawnBlock(block.uid, 'suffix', e.target.value.replace(/\s/g, ''))} className="w-full p-1.5 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] outline-none" placeholder="e.g. level, class, n1" />
                                    </div>
                                    <div>
                                        <label className="text-xs text-[#7d8590] block mb-1">Labels (comma-separated)</label>
                                        <input type="text" value={block.labelsText ?? block.labels.join(', ')} onChange={(e) => updateDrawnBlock(block.uid, 'labels', e.target.value)} className="w-full p-1.5 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] outline-none" placeholder="0, 1, 2, 3..." />
                                    </div>
                                </>
                            )}

                            {/* Answer Block Fields */}
                            {block.blockType === 'answer' && (
                                <>
                                    <div className="grid grid-cols-2 gap-2">
                                        <div>
                                            <label className="text-xs text-[#7d8590] block mb-1">Start Q#</label>
                                            <input type="number" min="1" value={block.startQ} onChange={(e) => updateDrawnBlock(block.uid, 'startQ', parseInt(e.target.value) || 1)} className="w-full p-1.5 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] outline-none" />
                                        </div>
                                        <div>
                                            <label className="text-xs text-[#7d8590] block mb-1">Rows</label>
                                            <input type="number" min="1" value={block.rows} onChange={(e) => updateDrawnBlock(block.uid, 'rows', Math.max(1, parseInt(e.target.value) || 1))} className="w-full p-1.5 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] outline-none" />
                                        </div>
                                    </div>
                                    <div>
                                        <label className="text-xs text-[#7d8590] block mb-1">Labels (comma-separated)</label>
                                        <input type="text" value={block.labelsText ?? block.labels.join(', ')} onChange={(e) => updateDrawnBlock(block.uid, 'labels', e.target.value)} className="w-full p-1.5 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] outline-none" placeholder="A, B, C, D" />
                                    </div>
                                </>
                            )}

                            {/* Delete button */}
                            <button onClick={() => removeDrawnBlock(block.uid)} className="w-full py-2 text-xs font-medium text-[#f85149] bg-[#f85149]/10 rounded border border-[#f85149]/30 hover:bg-[#f85149]/20 transition-colors flex justify-center items-center gap-1.5">
                                <Trash2 className="w-3 h-3" /> Delete This Block
                            </button>
                        </div>
                    </div>
                );
            })()}

            {/* Save Current Layout Modal */}
            {showSaveTemplateModal && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm" onClick={() => { setShowSaveTemplateModal(false); setNewTemplateName(''); }}>
                    <div className="bg-[#0d1117] border border-[#30363d] rounded-xl shadow-2xl w-full max-w-md mx-4" onClick={(e) => e.stopPropagation()}>
                        <div className="px-6 py-4 border-b border-[#30363d]">
                            <h2 className="text-lg font-bold text-[#e6edf3] flex items-center gap-2">
                                <BookmarkPlus className="w-5 h-5 text-[#58a6ff]" />
                                Save as Template
                            </h2>
                        </div>
                        <div className="p-6 space-y-4">
                            <p className="text-sm text-[#7d8590]">Save the current region layout as a reusable template. You can apply it to other MC sheets later.</p>
                            <input
                                type="text"
                                placeholder="Template name..."
                                value={newTemplateName}
                                onChange={(e) => setNewTemplateName(e.target.value)}
                                onKeyDown={(e) => { if (e.key === 'Enter') saveAsTemplate(); }}
                                className="w-full p-2 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] focus:bg-[#161b22] transition-colors outline-none"
                                autoFocus
                            />
                            <div className="text-xs text-[#7d8590]">
                                This will save {currentRegions.length} region(s) from Page {currentPageIndex + 1}
                            </div>
                        </div>
                        <div className="px-6 py-4 border-t border-[#30363d] flex justify-end gap-2">
                            <button onClick={() => { setShowSaveTemplateModal(false); setNewTemplateName(''); }} className="px-4 py-2 text-sm rounded-lg border border-[#30363d] bg-[#21262d] text-[#e6edf3] hover:bg-[#30363d] transition-colors">
                                Cancel
                            </button>
                            <button onClick={saveAsTemplate} disabled={!newTemplateName.trim()} className="px-4 py-2 text-sm rounded-lg bg-[#1f6feb] text-white hover:bg-[#388bfd] disabled:opacity-50 flex items-center gap-2 transition-colors">
                                <Save className="w-4 h-4" /> Save Template
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* LEFT SIDEBAR */}
            <div className="w-80 bg-[#010409] border-r border-[#30363d] flex flex-col shadow-xl z-10">
                <div className="p-4 border-b border-[#30363d] bg-[#0d1117] text-[#e6edf3] flex justify-between items-start">
                    <div>
                        <h1 className="text-xl font-bold flex items-center gap-2">
                            <ScanSearch className="w-6 h-6 text-[#58a6ff]" />
                            MC Auto Grader
                        </h1>
                    </div>
                </div>

                <div className="flex-1 overflow-y-auto p-4 space-y-6">
                    {/* 1. Upload */}
                    <div className="space-y-2">
                        <label className="text-xs font-semibold uppercase text-[#7d8590] tracking-wider flex items-center gap-2"><Upload className="w-4 h-4" /> Load Sheet</label>
                        <div className="relative group">
                            <input id="file-upload-input" type="file" onChange={handleFileUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" accept=".pdf,.jpg,.jpeg,.png" />
                            <div className="border-2 border-dashed border-[#30363d] rounded-lg p-6 text-center group-hover:border-[#58a6ff] group-hover:bg-[#21262d] transition-colors">
                                <Upload className="w-8 h-8 mx-auto text-[#7d8590] mb-2 group-hover:text-[#58a6ff]" />
                                <span className="text-sm font-medium text-[#c9d1d9]">Upload PDF / Image</span>
                                <span className="block text-xs text-[#7d8590] mt-1">Ctrl+O</span>
                            </div>
                        </div>
                        {pages.length > 0 && <div className="text-xs text-center text-[#7d8590] mt-1">{pages.length} page(s) loaded</div>}
                    </div>

                    {/* 2. Key */}
                    <div className="space-y-2">
                        <label className="text-xs font-semibold uppercase text-[#7d8590] tracking-wider flex items-center gap-2">
                            <GraduationCap className="w-4 h-4" />
                            Answer Key
                        </label>
                        <textarea placeholder="Paste answers (e.g. '1.A 2.B', 'ABCDA')" className="w-full h-24 p-2 text-sm border border-[#30363d] rounded bg-[#0d1117] text-[#e6edf3] focus:bg-[#161b22] transition-colors outline-none font-mono placeholder-[#7d8590]" value={answerKeyInput} onChange={(e) => setAnswerKeyInput(e.target.value)} />
                        <div className="text-xs text-[#7d8590]">Found {Object.keys(parsedAnswerKey).length} answers.</div>
                    </div>

                    {/* 3. Weights */}
                    {renderWeightSettings()}

                    {/* 4. Custom Templates */}
                    {renderTemplateSection()}
                </div>

                {/* Footer */}
                <div className="p-4 border-t border-[#30363d] bg-[#0d1117] space-y-2">
                    {!showResults && (
                    <button onClick={runBatchDetection} disabled={pages.length === 0 || isProcessing} className="w-full bg-[#1f6feb] text-white py-3 rounded-lg font-medium hover:bg-[#388bfd] transition-colors disabled:opacity-50 flex justify-center items-center gap-2 shadow-md">
                        {isProcessing ? (
                            <><Loader2 className="w-4 h-4 animate-spin" /> Scanning {progress.current} / {progress.total}</>
                        ) : (
                            <><CheckCircle2 className="w-5 h-5" /> Scan Answers <span className="text-xs opacity-50 ml-1">(Ctrl+Enter)</span></>
                        )}
                    </button>
                    )}
                    {showResults && (
                        <button onClick={exportExcel} className="w-full bg-[#238636] text-white py-3 rounded-lg font-medium hover:bg-[#2ea043] transition-colors flex justify-center items-center gap-2 shadow-md">
                            <FileSpreadsheet className="w-5 h-5" /> Export Results <span className="text-xs opacity-50 ml-1">(Ctrl+E)</span>
                        </button>
                    )}
                    <div className="text-[10px] text-center text-[#7d8590] mt-2 italic">
                        Your file will not be stored in any servers. Auto Grader may make an error, please review it.
                    </div>
                </div>
            </div>

            {/* MAIN WORKSPACE */}
            <div className="flex-1 flex flex-col bg-[#0d1117] relative">
                <div className="h-14 bg-[#010409] border-b border-[#30363d] flex items-center px-4 justify-between">
                    <div className="flex items-center gap-4 text-sm text-[#7d8590] overflow-hidden">
                        {file ? <span className="font-semibold text-[#e6edf3] truncate max-w-[300px]">{file.name}</span> : <span>No file loaded</span>}
                        {pages.length > 1 && (
                            <div className="flex items-center bg-[#21262d] rounded-lg p-1 gap-2 ml-4 border border-[#30363d]">
                                <button onClick={() => { setCurrentPageIndex(p => Math.max(0, p - 1)); setSelectedRegionId(null); }} disabled={currentPageIndex === 0} className="p-1 hover:bg-[#30363d] rounded shadow-sm disabled:opacity-30 text-[#e6edf3] bg-[#21262d]"><ChevronLeft className="w-4 h-4" /></button>
                                <div className="flex items-center gap-1 text-xs font-mono text-[#e6edf3] min-w-[80px] justify-center px-2">
                                    <span>Page</span>
                                    <input
                                        type="text"
                                        className="w-10 text-center bg-[#0d1117] border border-[#30363d] rounded focus:border-[#58a6ff] outline-none transition-colors text-[#e6edf3] py-0.5"
                                        value={pageInput}
                                        onChange={(e) => setPageInput(e.target.value)}
                                        onKeyDown={(e) => {
                                            if (e.key === 'Enter') {
                                                const val = parseInt(pageInput);
                                                if (!isNaN(val) && val >= 1 && val <= pages.length) {
                                                    setCurrentPageIndex(val - 1);
                                                    setSelectedRegionId(null);
                                                } else {
                                                    setPageInput(String(currentPageIndex + 1));
                                                }
                                                e.currentTarget.blur();
                                            }
                                        }}
                                        onBlur={() => {
                                            const val = parseInt(pageInput);
                                            if (!isNaN(val) && val >= 1 && val <= pages.length) {
                                                setCurrentPageIndex(val - 1);
                                                setSelectedRegionId(null);
                                            } else {
                                                setPageInput(String(currentPageIndex + 1));
                                            }
                                        }}
                                    />
                                    <span>/ {pages.length}</span>
                                </div>
                                <button onClick={() => { setCurrentPageIndex(p => Math.min(pages.length - 1, p + 1)); setSelectedRegionId(null); }} disabled={currentPageIndex === pages.length - 1} className="p-1 hover:bg-[#30363d] rounded shadow-sm disabled:opacity-30 text-[#e6edf3] bg-[#21262d]"><ChevronRight className="w-4 h-4" /></button>
                            </div>
                        )}
                    </div>
                    <div className="flex items-center gap-2">
                        <button onClick={() => setScale(s => Math.max(0.2, s - 0.1))} className="p-1.5 hover:bg-[#30363d] rounded text-[#7d8590] hover:text-[#e6edf3] transition-colors bg-[#21262d] border border-[#30363d]"><ZoomOut className="w-5 h-5" /></button>
                        <span className="text-xs font-mono w-12 text-center text-[#7d8590]">{Math.round(scale * 100)}%</span>
                        <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className="p-1.5 hover:bg-[#30363d] rounded text-[#7d8590] hover:text-[#e6edf3] transition-colors bg-[#21262d] border border-[#30363d]"><ZoomIn className="w-5 h-5" /></button>
                    </div>
                </div>

                <div ref={containerRef} className="flex-1 overflow-auto p-8 flex justify-center items-start bg-[#0d1117]">
                    {isProcessing && pages.length === 0 ? (
                        <div className="w-full h-full flex flex-col items-center justify-center p-8">
                            <div className="w-full h-full bg-[#21262d] rounded-lg animate-pulse border border-[#30363d] flex flex-col p-6 gap-3">
                                <div className="h-4 bg-[#30363d] rounded w-3/4" />
                                <div className="h-4 bg-[#30363d] rounded w-1/2" />
                                <div className="flex-1 grid grid-cols-2 gap-2 mt-4">
                                    {[...Array(8)].map((_, i) => <div key={i} className="bg-[#30363d] rounded animate-pulse" />)}
                                </div>
                            </div>
                            <p className="text-sm text-[#7d8590] mt-4">Loading {progress.total > 0 ? `${progress.current} / ${progress.total} pages` : 'pages'}...</p>
                        </div>
                    ) : currentPage ? (
                        <div className="relative shadow-2xl border border-[#30363d]">
                            <canvas ref={canvasRef} onMouseDown={handleMouseDown} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp} onMouseLeave={handleMouseUp} className={`bg-white ${drawMode ? 'cursor-crosshair' : (selectedRegionId ? 'cursor-move' : 'cursor-default')}`} />
                        </div>
                    ) : (
                        <div className="flex flex-col items-center justify-center h-full text-[#7d8590]">
                            <Layers className="w-24 h-24 mb-4 opacity-20" />
                            <p>Upload a PDF to view pages</p>
                        </div>
                    )}
                </div>

                {/* Results Panel */}
                {showResults && currentPage && currentPage.results && (
                    <div className="absolute bottom-0 left-0 right-0 bg-[#010409] border-t border-[#30363d] shadow-xl max-h-[300px] flex flex-col transition-transform z-30">
                        <div className="px-4 py-2 border-b border-[#30363d] flex justify-between items-center bg-[#0d1117]">
                            <div className="flex items-center gap-2 flex-wrap">
                                <UserSquare2 className="w-4 h-4 text-[#58a6ff]" />
                                {getIdFieldsForPage(currentPageIndex).map((field) => {
                                    const val = getIdValueByRegionId(field.regionId);
                                    const isIncomplete = val === '?';
                                    return (
                                        <select
                                            key={field.regionId}
                                            value={val}
                                            onChange={(e) => updateIdResult(field.regionId, e.target.value)}
                                            className={`text-sm font-bold px-1.5 py-0.5 rounded border outline-none cursor-pointer ${isIncomplete ? 'bg-[#da3633]/20 border-[#da3633]/40 text-[#f85149]' : 'bg-[#161b22] border-[#30363d] text-[#e6edf3] hover:border-[#58a6ff]'}`}
                                            title={field.label}
                                        >
                                            <option value="?" disabled>?</option>
                                            {field.options.map((opt) => (
                                                <option key={opt} value={opt}>{opt}</option>
                                            ))}
                                        </select>
                                    );
                                })}
                            </div>
                            <button onClick={() => setShowResults(false)} className="text-[#7d8590] hover:text-[#e6edf3] bg-transparent"><ChevronDown /></button>
                        </div>
                        <div className="flex-1 overflow-auto p-4 bg-[#0d1117]">
                            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-2">
                                {currentRegions.filter(r => r.type === 'answer').sort((a, b) => a.startQ - b.startQ).flatMap(r =>
                                    (currentPage.results[r.id] || []).map((res, resIdx) => {
                                        let bgColor = 'bg-[#161b22] border-[#30363d]';
                                        let textColor = 'text-[#c9d1d9]';
                                        const correctAns = parsedAnswerKey[res.qNum];
                                        if (res.label === 'BLANK') { bgColor = 'bg-[#161b22] border-[#30363d]'; textColor = 'text-[#7d8590] italic'; }
                                        else if (res.label === 'MULT') { bgColor = 'bg-[#d29922]/20 border-[#d29922]/40'; textColor = 'text-[#d29922] font-bold'; }
                                        else if (correctAns) {
                                            if (res.label === correctAns) { bgColor = 'bg-[#238636]/20 border-[#238636]/40'; textColor = 'text-[#3fb950] font-bold'; }
                                            else { bgColor = 'bg-[#da3633]/20 border-[#da3633]/40'; textColor = 'text-[#f85149] font-bold'; }
                                        }
                                        return (
                                            <div key={`${r.id}-${resIdx}`} onClick={() => cycleAnswer(r.id, resIdx)} className={`flex items-center justify-between text-sm p-2 rounded border cursor-pointer hover:opacity-80 transition-opacity select-none ${bgColor}`}>
                                                <span className="font-mono text-[#7d8590] text-xs">Q{res.qNum}</span>
                                                <div className="flex items-center gap-2">
                                                    {correctAns && res.label !== correctAns && <span className="text-xs text-[#7d8590] line-through mr-1">{correctAns}</span>}
                                                    <span className={`font-bold ${textColor}`}>{res.label || '-'}</span>
                                                </div>
                                            </div>
                                        );
                                    })
                                )}
                            </div>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
};

export default App;