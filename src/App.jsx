import React, { useState, useRef, useEffect, useMemo } from 'react';
import { Upload, FileSpreadsheet, Plus, Trash2, CheckCircle2, ScanSearch, Settings, ChevronRight, ChevronDown, Download, ZoomIn, ZoomOut, LayoutTemplate, ChevronLeft, Layers, ScanLine, GraduationCap } from 'lucide-react';

const App = () => {
  // State for multiple pages
  const [file, setFile] = useState(null);
  const [pages, setPages] = useState([]); // Array of { id, imageUrl, width, height, results, regions }
  const [currentPageIndex, setCurrentPageIndex] = useState(0);
  
  const [scale, setScale] = useState(1);
  const [selectedRegionId, setSelectedRegionId] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showResults, setShowResults] = useState(false);
  
  // Grading State
  const [answerKeyInput, setAnswerKeyInput] = useState("");
  const [parsedAnswerKey, setParsedAnswerKey] = useState({});

  const canvasRef = useRef(null);
  const containerRef = useRef(null);

  // Helper to get current page data safely
  const currentPage = pages[currentPageIndex];
  // Helper to get regions for rendering (default to empty if no page loaded)
  const currentRegions = currentPage?.regions || [];

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
      // Initialize worker once library is loaded
      // In v3+, the global is usually window.pdfjsLib
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

  // --- Answer Key Parsing Logic ---
  useEffect(() => {
    // Extract only A, B, C, D characters (case insensitive)
    // Map them to Question 1, 2, 3...
    const matches = answerKeyInput.toUpperCase().match(/[A-D]/g);
    const keyMap = {};
    if (matches) {
        matches.forEach((char, index) => {
            keyMap[index + 1] = char;
        });
    }
    setParsedAnswerKey(keyMap);
  }, [answerKeyInput]);

  // --- Helper: Row Layout Calculation ---
  const calculateRowLayout = (regionHeight, totalRows, gapRatio = 0.4) => {
    const configs = [];
    
    let totalUnits = 0;
    for (let i = 0; i < totalRows; i++) {
        const isGap = (i + 1) % 6 === 0;
        totalUnits += isGap ? gapRatio : 1;
    }

    const pxPerUnit = regionHeight / totalUnits;

    let currentY = 0;
    for (let i = 0; i < totalRows; i++) {
        const isGap = (i + 1) % 6 === 0;
        const rowHeight = (isGap ? gapRatio : 1) * pxPerUnit;
        
        configs.push({
            y: currentY,
            h: rowHeight,
            isGap: isGap
        });
        currentY += rowHeight;
    }
    return configs;
  };

  // --- Helper: Line Detection for Alignment ---
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
                const val = (data[idx] + data[idx+1] + data[idx+2]) / 3;
                if (val < 200) darkPixels++;
            }
            if (darkPixels > searchWidth * 0.5) {
                 if (darkPixels > maxDarkness) {
                     maxDarkness = darkPixels;
                     bestY = y;
                 }
            }
        }

        if (bestY !== -1) {
            const detectedAbsoluteY = searchStartY + bestY;
            const expectedAbsoluteY = height * EXPECTED_Y_RATIO;
            return detectedAbsoluteY - expectedAbsoluteY;
        }
    } catch (e) { console.warn("Y-alignment failed", e); }
    return 0;
  };

  const detectHorizontalOffset = (ctx, width, height) => {
    const EXPECTED_X_RATIO = 0.355; 

    const searchStartX = Math.floor(width * 0.2);
    const searchEndX = Math.floor(width * 0.5);
    const searchStartY = Math.floor(height * 0.5);
    const searchHeight = Math.floor(height * 0.9); // Search small band in header
    
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
                const val = (data[idx] + data[idx+1] + data[idx+2]) / 3;
                if (val < 200) darkPixels++;
            }
            if (darkPixels > searchHeight * 0.5) {
                 if (darkPixels > maxDarkness) {
                     maxDarkness = darkPixels;
                     bestX = x;
                 }
            }
        }

        if (bestX !== -1) {
            const detectedAbsoluteX = searchStartX + bestX;
            const expectedAbsoluteX = width * EXPECTED_X_RATIO;
            return detectedAbsoluteX - expectedAbsoluteX;
        }
    } catch (e) { console.warn("X-alignment failed", e); }
    return 0;
  };

  // --- Template Logic ---
  const getStandardRegions = (imgW, imgH, pageIdPrefix, xOffset = 0, yOffset = 0) => {
    const LABELS = ['A', 'B', 'C', 'D'];
    const BLOCK_W = imgW * 0.165; 
    const BASE_START_Y = imgH * 0.273;
    const START_Y = BASE_START_Y + yOffset;
    const BLOCK_H = imgH * 0.468; 
    const ROWS_PER_BLOCK = 35; 
    const BASE_X1 = imgW * 0.43;
    const BASE_X2 = imgW * 0.723;
    const X1 = BASE_X1 + xOffset;
    const X2 = BASE_X2 + xOffset;

    return [
      {
        id: `${pageIdPrefix}_b1`,
        x: X1,
        y: START_Y,
        w: BLOCK_W,
        h: BLOCK_H,
        rows: ROWS_PER_BLOCK,
        cols: 4,
        startQ: 1,
        labels: LABELS,
        gapHeightRatio: 0.6
      },
      {
        id: `${pageIdPrefix}_b2`,
        x: X2,
        y: START_Y-3,
        w: BLOCK_W,
        h: BLOCK_H,
        rows: ROWS_PER_BLOCK,
        cols: 4,
        startQ: 31,
        labels: LABELS,
        gapHeightRatio: 0.6
      }
    ];
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

        setPages([{
            id: 0,
            imageUrl: e.target.result,
            width: img.width,
            height: img.height,
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

    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
      const totalPages = pdf.numPages;
      const loadedPages = [];

      for (let i = 1; i <= totalPages; i++) {
          const page = await pdf.getPage(i);
          const viewport = page.getViewport({ scale: 2.0 }); 

          const canvas = document.createElement('canvas');
          const context = canvas.getContext('2d');
          canvas.height = viewport.height;
          canvas.width = viewport.width;

          await page.render({ canvasContext: context, viewport: viewport }).promise;
          
          const yOffset = detectVerticalOffset(context, canvas.width, canvas.height);
          const xOffset = detectHorizontalOffset(context, canvas.width, canvas.height);

          loadedPages.push({
              id: i - 1,
              imageUrl: canvas.toDataURL(),
              width: canvas.width,
              height: canvas.height,
              results: {},
              regions: getStandardRegions(canvas.width, canvas.height, `p${i-1}`, xOffset, yOffset)
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
      console.error("Error processing PDF:", error);
      alert("Failed to process PDF. Please try a different file.");
    } finally {
        setIsProcessing(false);
    }
  };

  const resetTemplate = () => {
      if (!currentPage) return;
      
      const img = new Image();
      img.onload = () => {
          const canvas = document.createElement('canvas');
          canvas.width = img.width;
          canvas.height = img.height;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0);
          
          const yOffset = detectVerticalOffset(ctx, img.width, img.height);
          const xOffset = detectHorizontalOffset(ctx, img.width, img.height);
          const newRegions = getStandardRegions(currentPage.width, currentPage.height, `p${currentPageIndex}`, xOffset, yOffset);
          
          setPages(prev => prev.map((p, idx) => {
              if (idx === currentPageIndex) {
                  return { ...p, regions: newRegions, results: {} };
              }
              return p;
          }));
      };
      img.src = currentPage.imageUrl;
  }

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

        const regionsToDraw = currentPage.regions || [];
        regionsToDraw.forEach(region => {
            ctx.strokeStyle = region.id === selectedRegionId ? '#3b82f6' : '#ef4444';
            ctx.lineWidth = 2;
            const x = region.x * scale;
            const y = region.y * scale;
            const w = region.w * scale;
            const h = region.h * scale;
            
            ctx.strokeRect(x, y, w, h);
            
            const rowLayouts = calculateRowLayout(h, region.rows, region.gapHeightRatio);
            
            ctx.lineWidth = 1;
            const cellW = w / region.cols;

            rowLayouts.forEach((row, idx) => {
                const rowY = y + row.y;
                const rowH = row.h;

                if (row.isGap) {
                    ctx.fillStyle = 'rgba(200, 200, 200, 0.3)';
                    ctx.fillRect(x, rowY, w, rowH);
                } else {
                    if (idx > 0) {
                        ctx.strokeStyle = region.id === selectedRegionId ? 'rgba(59, 130, 246, 0.5)' : 'rgba(239, 68, 68, 0.3)';
                        ctx.beginPath();
                        ctx.moveTo(x, rowY);
                        ctx.lineTo(x + w, rowY);
                        ctx.stroke();
                    }
                }
            });

            ctx.beginPath();
            ctx.strokeStyle = region.id === selectedRegionId ? 'rgba(59, 130, 246, 0.5)' : 'rgba(239, 68, 68, 0.3)';
            for (let i = 1; i < region.cols; i++) {
                ctx.moveTo(x + i * cellW, y);
                ctx.lineTo(x + i * cellW, y + h);
            }
            ctx.stroke();

            ctx.fillStyle = ctx.strokeStyle;
            ctx.font = 'bold 14px sans-serif';
            let qCount = 0;
            for(let r=0; r<region.rows; r++) if((r+1)%6 !== 0) qCount++;
            ctx.fillText(`Q${region.startQ} - Q${region.startQ + qCount - 1}`, x, y - 8);
            
            // Visualize results with grading colors
            if (currentPage.results && currentPage.results[region.id]) {
                currentPage.results[region.id].forEach((ans) => {
                    const correctAns = parsedAnswerKey[ans.qNum];
                    let fillStyle = 'rgba(34, 197, 94, 0.5)'; // Default Green

                    if (correctAns) {
                        if (ans.label === correctAns) {
                            fillStyle = 'rgba(34, 197, 94, 0.6)'; // Correct: Green
                        } else {
                            fillStyle = 'rgba(239, 68, 68, 0.6)'; // Incorrect: Red
                        }
                    }

                    if (ans.detectedIndex !== -1 && ans.label !== 'MULT' && ans.label !== 'BLANK') {
                        const relativeQ = ans.qNum - region.startQ;
                        const gapsBefore = Math.floor(relativeQ / 5);
                        const visualRowIndex = relativeQ + gapsBefore;
                        
                        if (visualRowIndex < rowLayouts.length) {
                            ctx.fillStyle = fillStyle;
                            const rLayout = rowLayouts[visualRowIndex];
                            const bubbleX = x + (ans.detectedIndex * cellW);
                            const bubbleY = y + rLayout.y;
                            ctx.fillRect(bubbleX + 2, bubbleY + 2, cellW - 4, rLayout.h - 4);
                        }
                    } else if (ans.label === 'MULT') {
                        // Mark multiple rows
                        const relativeQ = ans.qNum - region.startQ;
                        const gapsBefore = Math.floor(relativeQ / 5);
                        const visualRowIndex = relativeQ + gapsBefore;
                        
                        if (visualRowIndex < rowLayouts.length) {
                            const rLayout = rowLayouts[visualRowIndex];
                            const rowY = y + rLayout.y;
                            ctx.fillStyle = 'rgba(249, 115, 22, 0.4)'; // Orange for MULT
                            ctx.fillRect(x, rowY, w, rLayout.h);
                        }
                    }
                });
            }
        });
    };
    img.src = currentPage.imageUrl;

  }, [currentPage, scale, selectedRegionId, parsedAnswerKey]); 


  // --- Mouse Interactions ---
  const [isDragging, setIsDragging] = useState(false);
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
  const [activeRegionStart, setActiveRegionStart] = useState({ x: 0, y: 0 });

  const handleMouseDown = (e) => {
    if (!currentPage) return;
    const rect = canvasRef.current.getBoundingClientRect();
    const x = (e.clientX - rect.left) / scale;
    const y = (e.clientY - rect.top) / scale;

    const regionsToCheck = currentPage.regions || [];
    const clickedRegion = regionsToCheck.find(r => 
      x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h
    );

    if (clickedRegion) {
      setSelectedRegionId(clickedRegion.id);
      setIsDragging(true);
      setDragStart({ x, y });
      setActiveRegionStart({ x: clickedRegion.x, y: clickedRegion.y });
    } else {
      setSelectedRegionId(null);
    }
  };

  const handleMouseMove = (e) => {
    const rect = canvasRef.current.getBoundingClientRect();
    const currentX = (e.clientX - rect.left) / scale;
    const currentY = (e.clientY - rect.top) / scale;

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

  const handleMouseUp = (e) => {
    setIsDragging(false);
  };

  const runBatchDetection = async () => {
    setIsProcessing(true);
    const newPages = [...pages];

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
                    const regionResults = [];
                    const cellW = region.w / region.cols;
                    const rowLayouts = calculateRowLayout(region.h, region.rows, region.gapHeightRatio);
                    
                    let validQuestionCount = 0;

                    rowLayouts.forEach((rowConfig) => {
                        if (rowConfig.isGap) return;

                        const colScores = [];
                        for (let c = 0; c < region.cols; c++) {
                            const cx = Math.floor(region.x + c * cellW);
                            const cy = Math.floor(region.y + rowConfig.y);
                            const cw = Math.floor(cellW);
                            const ch = Math.floor(rowConfig.h);
                            const paddingX = cw * 0.30; 
                            const paddingY = ch * 0.30;
                            
                            let darkPixelCount = 0;
                            let totalPixelCount = 0;

                            for (let py = cy + paddingY; py < cy + ch - paddingY; py++) {
                                for (let px = cx + paddingX; px < cx + cw - paddingX; px++) {
                                    if (px < canvas.width && py < canvas.height) {
                                        const val = grayData[Math.floor(py) * canvas.width + Math.floor(px)];
                                        if (val < 200) darkPixelCount++;
                                        totalPixelCount++;
                                    }
                                }
                            }
                            const fillRatio = totalPixelCount > 0 ? darkPixelCount / totalPixelCount : 0;
                            colScores.push({ index: c, fillRatio: fillRatio });
                        }

                        // Determine Answer based on new logic
                        colScores.sort((a, b) => b.fillRatio - a.fillRatio);
                        const maxFill = colScores[0].fillRatio;
                        const minFill = colScores[colScores.length - 1].fillRatio;
                        const secondMaxFill = colScores[1].fillRatio;
                        
                        let label = '';
                        let selectedIndex = -1;
                        
                        // Rule 1: BLANK (Similar areas for all 4 OR max fill is very low)
                        // Using 10% threshold
                        if ((maxFill - minFill) < 0.1 || maxFill < 0.55) {
                            label = 'BLANK';
                        }
                        else if ((maxFill - secondMaxFill) < 0.05) {
                            label = 'MULT';
                        }
                        // Rule 3: Single Answer
                        else {
                            selectedIndex = colScores[0].index;
                            label = region.labels[selectedIndex];
                        }

                        regionResults.push({
                            qNum: region.startQ + validQuestionCount,
                            detectedIndex: selectedIndex,
                            label: label,
                            confidence: maxFill
                        });
                        validQuestionCount++;
                    });
                    pageResults[region.id] = regionResults;
                });

                // Post-Processing: Filtering based on Answer Key
                const keyCount = Object.keys(parsedAnswerKey).length;

                if (keyCount > 0) {
                    // Logic A: If Answer Key exists, strictly limit to the number of keys provided
                    Object.keys(pageResults).forEach(key => {
                        pageResults[key] = pageResults[key].filter(r => r.qNum <= keyCount);
                    });
                } else {
                    // Logic B: No Answer Key - Use auto-trim for trailing blanks
                    const allFlatResults = Object.values(pageResults).flat();
                    let maxAnsweredQ = 0;
                    allFlatResults.forEach(r => {
                        if (r.label !== 'BLANK') {
                            if (r.qNum > maxAnsweredQ) maxAnsweredQ = r.qNum;
                        }
                    });
                    Object.keys(pageResults).forEach(key => {
                        pageResults[key] = pageResults[key].filter(r => r.qNum <= maxAnsweredQ);
                    });
                }

                resolve(pageResults);
            };
            img.src = imgSrc;
        });
    };

    for (let i = 0; i < newPages.length; i++) {
        const results = await processSinglePage(newPages[i].imageUrl, newPages[i]);
        newPages[i] = { ...newPages[i], results: results };
    }

    setPages(newPages);
    setIsProcessing(false);
    setShowResults(true);
  };

  const exportExcel = () => {
    if (!window.XLSX) {
      alert("Excel library is still loading. Please try again in a moment.");
      return;
    }
    
    // --- 1. Prepare Main Data (Student Rows) ---
    // We need to determine the max question number to iterate correctly
    let maxQ = 0;
    const resultsData = pages.map((page, index) => {
        const rowData = { 'Page': index + 1 };
        const pageRegions = page.regions || [];
        
        pageRegions.sort((a,b) => a.startQ - b.startQ).forEach(region => {
            if (page.results && page.results[region.id]) {
                page.results[region.id].forEach(row => {
                    rowData[`Q${row.qNum}`] = row.label;
                    if (row.qNum > maxQ) maxQ = row.qNum;
                });
            }
        });
        return rowData;
    });

    // --- 2. Calculate Statistics for Footer ---
    const stats = {}; // Structure: { 1: { correct: 0, total: 0, counts: {A:0...} }, ... }
    
    // Initialize stats
    for(let q=1; q<=maxQ; q++) {
        stats[q] = { correct: 0, total: 0, counts: { A:0, B:0, C:0, D:0, MULT:0, BLANK:0 } };
    }

    // Aggregate Data
    pages.forEach((page) => {
        const pageRegions = page.regions || [];
        pageRegions.forEach(region => {
            if (page.results && page.results[region.id]) {
                page.results[region.id].forEach(row => {
                    const qNum = row.qNum;
                    const label = row.label; // "A", "B", "MULT", "BLANK"
                    const correctAns = parsedAnswerKey[qNum];
                    
                    if (stats[qNum]) {
                        stats[qNum].total++;
                        
                        // Count option frequency
                        // Normalize label to ensure it matches keys
                        if (stats[qNum].counts[label] !== undefined) {
                            stats[qNum].counts[label]++;
                        }

                        // Check correctness
                        if (correctAns && label === correctAns) {
                            stats[qNum].correct++;
                        }
                    }
                });
            }
        });
    });

    // --- 3. Construct Footer Rows ---
    const footerRows = [];
    
    // Spacing
    footerRows.push([]);
    footerRows.push([]);

    // Headers
    const rowQNum = ["Question No"];
    const rowAvg = ["Average Score"];
    const rowPct = ["Percentage"];
    const rowKey = ["Correct Answer"];
    const rowA = ["% Selecting A"];
    const rowB = ["% Selecting B"];
    const rowC = ["% Selecting C"];
    const rowD = ["% Selecting D"];

    for (let q=1; q<=maxQ; q++) {
        const s = stats[q];
        const total = s.total || 1; // Prevent divide by zero
        
        rowQNum.push(q);
        
        // Score stats
        const avg = s.correct / total;
        rowAvg.push(avg.toFixed(2));
        rowPct.push(`${(avg * 100).toFixed(1)}%`);
        
        // Correct Key
        rowKey.push(parsedAnswerKey[q] || "-");
        
        // Distribution stats
        rowA.push(`${((s.counts.A / total) * 100).toFixed(0)}%`);
        rowB.push(`${((s.counts.B / total) * 100).toFixed(0)}%`);
        rowC.push(`${((s.counts.C / total) * 100).toFixed(0)}%`);
        rowD.push(`${((s.counts.D / total) * 100).toFixed(0)}%`);
    }

    footerRows.push(rowQNum);
    footerRows.push(rowAvg);
    footerRows.push(rowPct);
    footerRows.push(rowKey);
    footerRows.push(rowA);
    footerRows.push(rowB);
    footerRows.push(rowC);
    footerRows.push(rowD);

    // --- 4. Generate Workbook ---
    const wb = window.XLSX.utils.book_new();
    
    // SHEET 1: RESULTS & STATS
    const wsResults = window.XLSX.utils.json_to_sheet(resultsData);
    // Append footer rows to the bottom of the sheet
    window.XLSX.utils.sheet_add_aoa(wsResults, footerRows, { origin: -1 });
    window.XLSX.utils.book_append_sheet(wb, wsResults, "All Results");

    // SHEET 2: SCORES (Per Student)
    const scoresData = pages.map((page, index) => {
        let correct = 0;
        let incorrect = 0;
        let blank = 0;
        let total = 0;

        const pageRegions = page.regions || [];
        pageRegions.forEach(region => {
            if (page.results && page.results[region.id]) {
                page.results[region.id].forEach(row => {
                    total++;
                    const correctAns = parsedAnswerKey[row.qNum];
                    if (correctAns) {
                        if (row.label === correctAns) {
                            correct++;
                        } else {
                            incorrect++;
                        }
                    }
                    if (row.label === 'BLANK') blank++;
                });
            }
        });

        // Calculate score based on available keys or total detected
        const totalScorable = Object.keys(parsedAnswerKey).length > 0 ? Object.keys(parsedAnswerKey).length : total;
        const percentage = totalScorable > 0 ? ((correct / totalScorable) * 100).toFixed(1) : 0;

        return {
            'Page': index + 1,
            'Score': correct,
            'Percentage': `${percentage}%`,
        };
    });

    const wsScores = window.XLSX.utils.json_to_sheet(scoresData);
    window.XLSX.utils.book_append_sheet(wb, wsScores, "Scores");

    window.XLSX.writeFile(wb, "omr_graded_results.xlsx");
  };

  // --- UI Components ---
  return (
    <div className="flex h-screen bg-slate-50 text-slate-800 font-sans overflow-hidden">
      
      {/* LEFT SIDEBAR - CONTROLS */}
      <div className="w-80 bg-white border-r border-slate-200 flex flex-col shadow-xl z-10">
        <div className="p-4 border-b border-slate-200 bg-slate-900 text-white">
          <h1 className="text-xl font-bold flex items-center gap-2">
            <ScanSearch className="w-6 h-6 text-blue-400" />
            Auto Grader
          </h1>
          <p className="text-xs text-slate-400 mt-1">Specialized for Ho Fung College</p>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-6">
          
          {/* 1. Upload Section */}
          <div className="space-y-2">
            <label className="text-xs font-semibold uppercase text-slate-500 tracking-wider">1. Load Sheet</label>
            <div className="relative group">
              <input 
                type="file" 
                onChange={handleFileUpload} 
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                accept=".pdf,.jpg,.jpeg,.png"
              />
              <div className="border-2 border-dashed border-slate-300 rounded-lg p-6 text-center group-hover:border-blue-500 group-hover:bg-blue-50 transition-colors">
                <Upload className="w-8 h-8 mx-auto text-slate-400 mb-2 group-hover:text-blue-500" />
                <span className="text-sm font-medium text-slate-600">Upload PDF / Image</span>
              </div>
            </div>
            {pages.length > 0 && (
                <div className="text-xs text-center text-slate-500 mt-1">
                    {pages.length} page(s) loaded
                </div>
            )}
          </div>

          {/* 2. Answer Key Input */}
          <div className="space-y-2">
            <label className="text-xs font-semibold uppercase text-slate-500 tracking-wider flex items-center gap-2">
                <GraduationCap className="w-4 h-4"/>
                Answer Key
            </label>
            <textarea
                placeholder="Paste answers here (e.g. '1. A 2. B' or 'ABCDE...')"
                className="w-full h-24 p-2 text-sm border border-slate-300 rounded focus:ring-2 focus:ring-blue-500 outline-none font-mono"
                value={answerKeyInput}
                onChange={(e) => setAnswerKeyInput(e.target.value)}
            />
            <div className="text-xs text-slate-400">
                Found {Object.keys(parsedAnswerKey).length} answers.
            </div>
          </div>

          {/* Instructions */}
          <div className="p-4 bg-blue-50 text-blue-800 text-sm rounded-lg border border-blue-100">
             <h3 className="font-bold mb-2 flex items-center gap-2">
                 <ScanLine className="w-4 h-4"/>
                 Auto-Align Active
             </h3>
             <p className="text-xs mb-2">Detects header line to fix X/Y shifts.</p>
             <ol className="list-decimal list-inside space-y-1 text-xs opacity-80">
                 <li>Upload PDF.</li>
                 <li><strong>Check alignment</strong> on each page.</li>
                 <li>(Optional) Enter Answer Key.</li>
                 <li>Click Scan Answers.</li>
             </ol>
             <button onClick={resetTemplate} className="mt-3 text-xs text-blue-600 underline hover:text-blue-800">
                 Recalculate Alignment (This Page)
             </button>
          </div>

          {/* Region List */}
          {currentRegions.length > 0 && (
            <div className="space-y-2">
              <label className="text-xs font-semibold uppercase text-slate-500 tracking-wider">Blocks (Page {currentPageIndex + 1})</label>
              {currentRegions.sort((a,b) => a.startQ - b.startQ).map((r, i) => {
                 let qCount = 0;
                 for(let k=0; k<r.rows; k++) if((k+1)%6 !== 0) qCount++;

                 return (
                  <button
                    key={r.id}
                    onClick={() => setSelectedRegionId(r.id)}
                    className={`w-full text-left px-3 py-2 rounded text-sm flex justify-between items-center ${selectedRegionId === r.id ? 'bg-blue-100 text-blue-700 border border-blue-200' : 'bg-white border hover:bg-slate-50'}`}
                  >
                    <span>Questions {r.startQ} - {r.startQ + qCount - 1}</span>
                    <ChevronRight className="w-4 h-4 opacity-50" />
                  </button>
                 );
              })}
            </div>
          )}
        </div>

        {/* Action Footer */}
        <div className="p-4 border-t border-slate-200 bg-slate-50 space-y-2">
            <button 
                onClick={runBatchDetection}
                disabled={pages.length === 0 || isProcessing}
                className="w-full bg-slate-900 text-white py-3 rounded-lg font-medium hover:bg-slate-800 transition-colors disabled:opacity-50 disabled:cursor-not-allowed flex justify-center items-center gap-2"
            >
                {isProcessing ? 'Processing...' : (
                    <>
                        <CheckCircle2 className="w-5 h-5" />
                        Scan Answers ({pages.length})
                    </>
                )}
            </button>
            {showResults && (
                <button 
                    onClick={exportExcel}
                    className="w-full bg-green-600 text-white py-3 rounded-lg font-medium hover:bg-green-700 transition-colors flex justify-center items-center gap-2"
                >
                    <FileSpreadsheet className="w-5 h-5" />
                    Export Graded Results
                </button>
            )}
        </div>
      </div>

      {/* MAIN WORKSPACE */}
      <div className="flex-1 flex flex-col bg-slate-100 relative">
        
        {/* Toolbar */}
        <div className="h-12 bg-white border-b border-slate-200 flex items-center px-4 justify-between">
           <div className="flex items-center gap-4 text-sm text-slate-600">
               {file ? <span className="font-semibold text-slate-900 truncate max-w-[200px]">{file.name}</span> : <span>No file loaded</span>}
               {pages.length > 1 && (
                   <div className="flex items-center bg-slate-100 rounded-lg p-1 gap-2 ml-4">
                       <button 
                         onClick={() => {
                             setCurrentPageIndex(p => Math.max(0, p - 1));
                             setSelectedRegionId(null);
                         }}
                         disabled={currentPageIndex === 0}
                         className="p-1 hover:bg-white rounded shadow-sm disabled:opacity-30"
                       >
                           <ChevronLeft className="w-4 h-4"/>
                       </button>
                       <span className="text-xs font-mono min-w-[60px] text-center">
                           Page {currentPageIndex + 1} / {pages.length}
                       </span>
                       <button 
                         onClick={() => {
                             setCurrentPageIndex(p => Math.min(pages.length - 1, p + 1));
                             setSelectedRegionId(null);
                         }}
                         disabled={currentPageIndex === pages.length - 1}
                         className="p-1 hover:bg-white rounded shadow-sm disabled:opacity-30"
                       >
                           <ChevronRight className="w-4 h-4"/>
                       </button>
                   </div>
               )}
           </div>
           <div className="flex items-center gap-2">
               <button onClick={() => setScale(s => Math.max(0.2, s - 0.1))} className="p-1.5 hover:bg-slate-100 rounded text-slate-600"><ZoomOut className="w-5 h-5"/></button>
               <span className="text-xs font-mono w-12 text-center">{Math.round(scale * 100)}%</span>
               <button onClick={() => setScale(s => Math.min(3, s + 0.1))} className="p-1.5 hover:bg-slate-100 rounded text-slate-600"><ZoomIn className="w-5 h-5"/></button>
           </div>
        </div>

        {/* Canvas Area */}
        <div 
          ref={containerRef}
          className="flex-1 overflow-auto p-8 flex justify-center items-start"
        >
          {currentPage ? (
             <div className="relative shadow-2xl">
                <canvas 
                    ref={canvasRef}
                    onMouseDown={handleMouseDown}
                    onMouseMove={handleMouseMove}
                    onMouseUp={handleMouseUp}
                    className={`bg-white ${selectedRegionId ? 'cursor-move' : 'cursor-default'}`}
                />
             </div>
          ) : (
             <div className="flex flex-col items-center justify-center h-full text-slate-400">
                <Layers className="w-24 h-24 mb-4 opacity-20" />
                <p>Upload a PDF to view pages</p>
             </div>
          )}
        </div>

        {/* Results Panel (Slide up) */}
        {showResults && currentPage && currentPage.results && (
            <div className="absolute bottom-0 left-0 right-0 bg-white border-t border-slate-200 shadow-xl max-h-[300px] flex flex-col transition-transform">
                <div className="px-4 py-2 border-b border-slate-100 flex justify-between items-center bg-slate-50">
                    <h3 className="font-bold text-slate-700 flex items-center gap-2">
                        Detected Results 
                        <span className="text-xs font-normal text-slate-500 bg-slate-100 px-2 py-0.5 rounded">Page {currentPageIndex + 1}</span>
                    </h3>
                    <button onClick={() => setShowResults(false)} className="text-slate-400 hover:text-slate-600"><ChevronDown/></button>
                </div>
                <div className="flex-1 overflow-auto p-4">
                    <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-2">
                        {currentRegions.sort((a,b) => a.startQ - b.startQ).flatMap(r => currentPage.results[r.id] || []).map((res, idx) => {
                            let bgColor = 'bg-slate-50 border-slate-100';
                            let textColor = 'text-slate-900';
                            const correctAns = parsedAnswerKey[res.qNum];

                            // Color Coding Logic
                            if (res.label === 'BLANK') {
                                bgColor = 'bg-slate-100 border-slate-200';
                                textColor = 'text-slate-400 italic';
                            } else if (res.label === 'MULT') {
                                bgColor = 'bg-orange-50 border-orange-200';
                                textColor = 'text-orange-600 font-bold';
                            } else if (correctAns) {
                                if (res.label === correctAns) {
                                    bgColor = 'bg-green-50 border-green-200';
                                    textColor = 'text-green-700 font-bold';
                                } else {
                                    bgColor = 'bg-red-50 border-red-200';
                                    textColor = 'text-red-600 font-bold';
                                }
                            }

                            return (
                                <div key={idx} className={`flex items-center justify-between text-sm p-2 rounded border ${bgColor}`}>
                                    <span className="font-mono text-slate-500 text-xs">Q{res.qNum}</span>
                                    <div className="flex items-center gap-2">
                                        {correctAns && res.label !== correctAns && (
                                            <span className="text-xs text-slate-400 line-through mr-1">{correctAns}</span>
                                        )}
                                        <span className={`font-bold ${textColor}`}>
                                            {res.label || '-'}
                                        </span>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </div>
            </div>
        )}

      </div>
      
      {/* Hidden Imports for PDF.js and XLSX */}
      <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/tailwindcss/2.2.19/tailwind.min.css" />
    </div>
  );
};

export default App;