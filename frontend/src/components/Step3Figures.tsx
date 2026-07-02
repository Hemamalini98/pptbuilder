import React, { useEffect, useRef, useState } from 'react';
import { useStore } from '../store';
import { Loader2 } from 'lucide-react';
import { toast } from 'sonner';
import { type AltTextEntry } from '../store';

const AltTextField: React.FC<{
  figId: string;
  stored: string | undefined;
  excelEntry: AltTextEntry | undefined;
  onUpdate: (id: string, text: string) => void;
}> = ({ figId, stored, excelEntry, onUpdate }) => {
  const hasExcel = !!excelEntry;
  const [value, setValue] = useState(stored ?? '');

  // Sync inward when store value changes from outside (e.g. Apply button)
  useEffect(() => {
    setValue(stored ?? '');
  }, [stored]);

  return (
    <div className="space-y-0.5">
      <div className="flex items-center justify-between">
        <span className="font-semibold text-[#64748b] text-[7px] uppercase tracking-wider">
          Alt Text {hasExcel && <span className="text-[#22d3ee]">✓</span>}
        </span>
        {hasExcel && !stored && (
          <button
            onClick={() => onUpdate(figId, excelEntry.alt_text_short)}
            className="text-[7px] text-[#38bdf8] hover:text-white border-none bg-none cursor-pointer px-1"
            title="Apply alt text from Excel"
          >
            Apply
          </button>
        )}
      </div>
      <textarea
        value={value}
        placeholder={hasExcel ? 'Click Apply or type…' : 'No alt text — upload Excel'}
        rows={3}
        onChange={(e) => {
          setValue(e.target.value);
          onUpdate(figId, e.target.value);
        }}
        className="w-full bg-[#0f172a] border border-[#334155] rounded px-1.5 py-1 text-[#e2e8f0] text-[8.5px] outline-none focus:border-[#38bdf8] resize-none leading-relaxed placeholder:text-[#475569]"
      />
      {hasExcel && excelEntry.decorative && (
        <span className="text-[7px] text-amber-400 font-semibold">Decorative</span>
      )}
    </div>
  );
};

const PdfThumbnail: React.FC<{ doc: any; pageNum: number; active: boolean; onClick: () => void }> = ({ doc, pageNum, active, onClick }) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [isVisible, setIsVisible] = useState(false);

  useEffect(() => {
    const observer = new IntersectionObserver(
      ([entry]) => {
        if (entry.isIntersecting) {
          setIsVisible(true);
          observer.disconnect();
        }
      },
      { rootMargin: '250px' }
    );
    if (containerRef.current) {
      observer.observe(containerRef.current);
    }
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    if (!doc || !canvasRef.current || !isVisible) return;
    let isCancelled = false;

    const renderThumbnail = async () => {
      try {
        const page = await doc.getPage(pageNum);
        const viewport = page.getViewport({ scale: 0.2 });
        const canvas = canvasRef.current;
        if (!canvas || isCancelled) return;

        const context = canvas.getContext('2d');
        if (!context) return;

        canvas.width = viewport.width;
        canvas.height = viewport.height;

        await page.render({
          canvasContext: context,
          viewport: viewport,
        }).promise;
      } catch (err) {
        console.error('Thumbnail render error:', err);
      }
    };

    renderThumbnail();
    return () => {
      isCancelled = true;
    };
  }, [doc, pageNum, isVisible]);

  return (
    <div
      ref={containerRef}
      onClick={onClick}
      className={`relative rounded overflow-hidden cursor-pointer transition-all border-2 flex-shrink-0 bg-[#0f172a] ${
        active
          ? 'border-[#38bdf8]'
          : 'border-transparent hover:border-zinc-700'
      }`}
      style={{ aspectRatio: '3 / 4', width: '100%' }}
    >
      {isVisible ? (
        <canvas ref={canvasRef} className="w-full block" />
      ) : (
        <div className="w-full h-full bg-zinc-900/60 animate-pulse flex items-center justify-center text-[10px] text-zinc-500 font-bold">
          Pg {pageNum}
        </div>
      )}
      <div className="absolute bottom-0 left-0 right-0 bg-black/65 text-[#94a3b8] text-[9.5px] text-center py-0.5 select-none font-bold">
        {pageNum}
      </div>
    </div>
  );
};

export const Step3Figures: React.FC = () => {
  const {
    sourcePdfPages,
    pdfUrl,
    currentPdfPage,
    figures,
    pdfCaptions,
    altTextEntries,
    altTextLoading,
    detectedChapter,
    addFigure,
    renameFigure,
    updateFigureCaption,
    updateFigureAltText,
    deleteFigure,
    uploadAltTextExcel,
    convertDeck,
    isConverting,
    conversionProgress,
    slides,
  } = useStore();

  const altTextFileRef = useRef<HTMLInputElement>(null);

  // Build a lookup map from figure_key → entry for fast access
  const altTextMap = React.useMemo(() => {
    const m: Record<string, AltTextEntry> = {};
    altTextEntries.forEach((e) => { m[e.figure_key] = e; });
    return m;
  }, [altTextEntries]);

  // Resolve alt text entry for a figure label.
  // - "Figure 1.1" → chapter is already embedded as the first number; direct lookup only.
  // - "Figure 1"   → no chapter in label; use detectedChapter from filename as fallback.
  const resolveAltEntry = React.useCallback((label: string): AltTextEntry | undefined => {
    const key = label.toLowerCase().trim();
    if (altTextMap[key]) return altTextMap[key];

    const dotFmt = key.match(/^(figure|table)\s+\d+\.\d+/);
    if (dotFmt) return undefined; // chapter already in label — no further fallback

    // Single number format e.g. "figure 1" — use filename-detected chapter
    if (detectedChapter != null) {
      const m = key.match(/^(figure|table)\s+(\d+)$/);
      if (m) {
        const fallbackKey = `${m[1]} ${detectedChapter}.${m[2]}`;
        if (altTextMap[fallbackKey]) return altTextMap[fallbackKey];
      }
    }
    return undefined;
  }, [altTextMap, detectedChapter]);

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  
  const [pdfDoc, setPdfDoc] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [extracting, setExtracting] = useState(false);
  const [zoom, setZoom] = useState(1.0);
  const [showOnlyMentioned, setShowOnlyMentioned] = useState(true);
  
  // Selection box state
  const [isDrawing, setIsDrawing] = useState(false);
  const [startPos, setStartPos] = useState({ x: 0, y: 0 });
  const [currentBox, setCurrentBox] = useState<{ x: number; y: number; w: number; h: number } | null>(null);
  const [showAltWarning, setShowAltWarning] = useState(false);

  // Extract all figures/tables referenced in the source presentation
  const mentionedRefs = React.useMemo(() => {
    const refs = new Set<string>();
    if (!slides) return refs;
    slides.forEach((slide) => {
      slide.shapes.forEach((shape: any) => {
        const text = (shape.textBody?.paragraphs || [])
          .map((para: any) => para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '')
          .join(' ')
          .trim();
        if (text) {
          const figRegex = /\b(figure|fig\.?|f\.?)\s*([\d.-]+)/gi;
          let match;
          while ((match = figRegex.exec(text)) !== null) {
            refs.add(`figure ${match[2]}`.toLowerCase());
            refs.add(`fig ${match[2]}`.toLowerCase());
          }
          const tabRegex = /\b(table|tab\.?|t\.?)\s*([\d.-]+)/gi;
          while ((match = tabRegex.exec(text)) !== null) {
            refs.add(`table ${match[2]}`.toLowerCase());
            refs.add(`tab ${match[2]}`.toLowerCase());
          }
        }
      });
    });
    return refs;
  }, [slides]);

  const filteredCaptions = React.useMemo(() => {
    if (!showOnlyMentioned) return pdfCaptions;
    return pdfCaptions.filter(cap => {
      const label = (cap.label || '').toLowerCase();
      return mentionedRefs.has(label);
    });
  }, [pdfCaptions, showOnlyMentioned, mentionedRefs]);

  const ZOOM_STOPS = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 2.0, 3.0];

  // Load PDF document
  useEffect(() => {
    let active = true;
    let timer: NodeJS.Timeout | null = null;

    const loadPdf = async () => {
      const pdfjsLib = (window as any).pdfjsLib;
      if (!pdfjsLib) {
        timer = setTimeout(loadPdf, 100);
        return;
      }
      
      if (!pdfUrl) return;

      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

      if (!active) return;
      setLoading(true);
      try {
        const loadingTask = pdfjsLib.getDocument(pdfUrl);
        const doc = await loadingTask.promise;
        if (active) {
          setPdfDoc(doc);
        }
      } catch (err) {
        console.error('Error loading PDF:', err);
        toast.error('Failed to parse PDF document structure.');
      } finally {
        if (active) {
          setLoading(false);
        }
      }
    };

    loadPdf();

    return () => {
      active = false;
      if (timer) clearTimeout(timer);
    };
  }, [pdfUrl]);

  // Load captions if empty (e.g. on page refresh or pre-existing upload)
  useEffect(() => {
    const fetchCaptions = async () => {
      if (pdfUrl && pdfCaptions.length === 0) {
        try {
          const res = await fetch('/api/pdf/captions');
          const data = await res.json();
          if (data.ok) {
            useStore.setState({ pdfCaptions: data.captions || [] });
          }
        } catch (e) {
          console.error("Failed to fetch captions:", e);
        }
      }
    };
    fetchCaptions();
  }, [pdfUrl, pdfCaptions.length]);

  // Render current page when page changes or zoom changes
  useEffect(() => {
    const renderPage = async () => {
      if (!pdfDoc || !canvasRef.current) return;
      setLoading(true);
      try {
        const page = await pdfDoc.getPage(currentPdfPage + 1);
        const canvas = canvasRef.current;
        const context = canvas.getContext('2d');
        if (!context) return;

        // Auto-scale to fit container size
        const padding = 32;
        const containerWidth = Math.max(300, (containerRef.current?.clientWidth || 700) - padding);
        const containerHeight = Math.max(300, (containerRef.current?.clientHeight || 500) - padding);
        const baseViewport = page.getViewport({ scale: 1.0 });
        
        const scaleX = containerWidth / baseViewport.width;
        const scaleY = containerHeight / baseViewport.height;
        const newScale = Math.min(scaleX, scaleY) * zoom;

        // Sharp high-DPI rendering for Mac Retina displays
        const dpr = window.devicePixelRatio || 1;
        const viewport = page.getViewport({ scale: newScale * dpr });
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.style.width = `${viewport.width / dpr}px`;
        canvas.style.height = `${viewport.height / dpr}px`;

        const renderContext = {
          canvasContext: context,
          viewport: viewport,
        };
        await page.render(renderContext).promise;
      } catch (err) {
        console.error('Error rendering page:', err);
      } finally {
        setLoading(false);
      }
    };
    renderPage();
  }, [pdfDoc, currentPdfPage, zoom]);

  // Handle Drag Selection
  const handleMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    if (!canvasRef.current || loading || extracting) return;
    const rect = canvasRef.current.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    setIsDrawing(true);
    setStartPos({ x, y });
    setCurrentBox({ x, y, w: 0, h: 0 });
  };

  const handleMouseMove = (e: React.MouseEvent<HTMLDivElement>) => {
    if (!isDrawing || !canvasRef.current || !currentBox) return;
    const rect = canvasRef.current.getBoundingClientRect();
    const x = Math.max(0, Math.min(e.clientX - rect.left, rect.width));
    const y = Math.max(0, Math.min(e.clientY - rect.top, rect.height));

    const w = x - startPos.x;
    const h = y - startPos.y;

    setCurrentBox({
      x: w < 0 ? x : startPos.x,
      y: h < 0 ? y : startPos.y,
      w: Math.abs(w),
      h: Math.abs(h),
    });
  };

  const handleMouseUp = async () => {
    if (!isDrawing || !currentBox) return;
    setIsDrawing(false);
  };

  const executeExtraction = async () => {
    if (!currentBox) return;
    setExtracting(true);
    const cropPromise = new Promise(async (resolve, reject) => {
      try {
        const formData = new FormData();
        formData.append('page', String(currentPdfPage));
        formData.append('scale', String(1.0)); // Send base scaling relative to original PDF points
        
        // Calculate original PDF coordinates based on active viewport aspect ratio
        if (!pdfDoc || !canvasRef.current) return;
        const page = await pdfDoc.getPage(currentPdfPage + 1);
        const baseViewport = page.getViewport({ scale: 1.0 });
        const canvasRect = canvasRef.current.getBoundingClientRect();
        
        const coordScaleX = baseViewport.width / canvasRect.width;
        const coordScaleY = baseViewport.height / canvasRect.height;
        
        formData.append('x0', String(currentBox.x * coordScaleX));
        formData.append('y0', String(currentBox.y * coordScaleY));
        formData.append('x1', String((currentBox.x + currentBox.w) * coordScaleX));
        formData.append('y1', String((currentBox.y + currentBox.h) * coordScaleY));

        const res = await fetch('/api/extract', {
          method: 'POST',
          body: formData,
        });
        const data = await res.json();
        
        if (data.url) {
          addFigure({
            url: data.url,
            page: data.page,
            filename: data.filename,
          });
          resolve(data);
        } else {
          reject(new Error('No URL returned'));
        }
      } catch (err) {
        reject(err);
      }
    });

    toast.promise(cropPromise, {
      loading: 'Extracting figure...',
      success: 'Figure extracted successfully!',
      error: 'Extraction failed. Make sure coordinates are in bounds.',
    });

    try {
      await cropPromise;
    } catch (e) {
      console.error(e);
    } finally {
      setExtracting(false);
      setCurrentBox(null);
    }
  };

  const clearSelection = () => {
    setCurrentBox(null);
  };

  const handleZoomOut = () => {
    const prev = [...ZOOM_STOPS].reverse().find(z => z < zoom - 0.01);
    if (prev !== undefined) setZoom(prev);
  };

  const handleZoomIn = () => {
    const next = ZOOM_STOPS.find(z => z > zoom + 0.01);
    if (next !== undefined) setZoom(next);
  };

  const handleZoomSelect = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const value = e.target.value;
    if (value === 'fit') {
      setZoom(1.0);
    } else {
      setZoom(parseFloat(value));
    }
  };

  const getPdfFilename = () => {
    if (!pdfUrl) return 'Loading...';
    return pdfUrl.substring(pdfUrl.lastIndexOf('/') + 1) || 'source.pdf';
  };

  return (
    <div className="flex h-[calc(100vh-140px)] w-full overflow-hidden text-slate-200 bg-[#0f172a] rounded-[var(--radius-custom)] border border-[#334155]">
      
      {/* ── Sidebar (Left) ── */}
      <aside className="w-44 flex-shrink-0 bg-[#1e293b] flex flex-col border-r border-[#334155] overflow-hidden">
        <div className="p-3 border-b border-[#334155] text-left">
          <span className="block font-bold text-xs text-[#f1f5f9] truncate" title={getPdfFilename()}>
            {getPdfFilename()}
          </span>
          <span className="text-[#64748b] text-[10px] block mt-0.5">{sourcePdfPages} pages</span>
        </div>
        <div className="flex-1 overflow-y-auto p-2.5 space-y-3.5 thumbs">
          {Array.from({ length: sourcePdfPages }).map((_, idx) => (
            <PdfThumbnail
              key={idx}
              doc={pdfDoc}
              pageNum={idx + 1}
              active={currentPdfPage === idx}
              onClick={() => useStore.setState({ currentPdfPage: idx })}
            />
          ))}
        </div>
      </aside>

      {/* ── Main Panel (Center) ── */}
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden main">
        
        {/* Toolbar */}
        <div className="bg-[#1e293b] border-b border-[#334155] px-4 py-2.5 flex items-center gap-3.5 flex-shrink-0 toolbar">
          <button
            onClick={executeExtraction}
            disabled={!currentBox || extracting}
            className="border-none rounded-md px-3.5 py-1.5 text-xs font-bold bg-[#0284c7] hover:bg-[#0369a1] text-white disabled:bg-[#1e3a4f] disabled:text-[#475569] disabled:cursor-default transition-all cursor-pointer"
          >
            ⬇ Extract
          </button>
          
          <button
            onClick={clearSelection}
            disabled={!currentBox}
            className="border-none rounded-md px-3.5 py-1.5 text-xs font-bold bg-[#334155] hover:bg-[#475569] text-[#94a3b8] disabled:opacity-40 disabled:cursor-default transition-all cursor-pointer"
          >
            ✕ Clear
          </button>

          <div className="w-[1px] bg-[#334155] h-5 sep"></div>

          <div className="flex items-center gap-1.5 zoom-grp">
            <button
              onClick={handleZoomOut}
              className="bg-[#334155] hover:bg-[#475569] text-white w-6 h-6 rounded flex items-center justify-center font-bold text-sm cursor-pointer border-none"
            >
              −
            </button>
            <select
              value={zoom}
              onChange={handleZoomSelect}
              className="bg-[#334155] border border-[#475569] text-[#e2e8f0] px-1.5 rounded text-xs cursor-pointer h-6 min-w-[70px] outline-none"
            >
              <option value="0.25">25%</option>
              <option value="0.5">50%</option>
              <option value="0.75">75%</option>
              <option value="1">100%</option>
              <option value="1.25">125%</option>
              <option value="1.5">150%</option>
              <option value="2">200%</option>
              <option value="3">300%</option>
              <option value="fit">Fit width</option>
            </select>
            <button
              onClick={handleZoomIn}
              className="bg-[#334155] hover:bg-[#475569] text-white w-6 h-6 rounded flex items-center justify-center font-bold text-sm cursor-pointer border-none"
            >
              +
            </button>
          </div>

          <div className="w-[1px] bg-[#334155] h-5 sep"></div>

          {/* Alt-Text Excel Upload */}
          <input
            ref={altTextFileRef}
            type="file"
            accept=".xlsx,.xls"
            className="hidden"
            onChange={async (e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              e.target.value = '';
              const p = uploadAltTextExcel(file);
              toast.promise(p, {
                loading: 'Parsing alt-text Excel…',
                success: `Alt texts loaded`,
                error: 'Failed to parse Excel',
              });
            }}
          />
          <button
            onClick={() => altTextFileRef.current?.click()}
            disabled={altTextLoading}
            title="Upload alt-text Excel"
            className="border-none rounded-md px-3 py-1.5 text-xs font-bold bg-[#1e3a4f] hover:bg-[#0284c7] text-[#38bdf8] hover:text-white disabled:opacity-50 disabled:cursor-default transition-all cursor-pointer flex items-center gap-1.5"
          >
            {altTextLoading ? <Loader2 className="w-3 h-3 animate-spin" /> : '♿'}
            Alt Text {altTextEntries.length > 0 && detectedChapter != null && (() => {
                const chapterCount = altTextEntries.filter(e => parseInt(e.chapter) === detectedChapter).length;
                return <span className="text-[9px] opacity-70">({chapterCount} CH{detectedChapter})</span>;
              })()}
          </button>

          <div className="w-[1px] bg-[#334155] h-5 sep"></div>

          <span className="text-[10px] text-[#64748b] flex-1 text-left font-mono coords">
            {currentBox
              ? `${Math.round(currentBox.w)} × ${Math.round(currentBox.h)} px selected`
              : 'Drag on the page below to select a region'}
          </span>

          <span className="text-[10px] text-[#475569] font-bold whitespace-nowrap page-info">
            Page {currentPdfPage + 1} / {sourcePdfPages}
          </span>
        </div>

        {/* Page Area */}
        <div className="flex-1 overflow-auto bg-[#374151] p-5 flex min-h-0 page-area" ref={containerRef}>
          {loading && (
            <div className="absolute inset-0 bg-[#0f172a]/75 flex items-center justify-center z-10 backdrop-blur-xs">
              <Loader2 className="w-8 h-8 text-[#0284c7] animate-spin" />
            </div>
          )}

          <div
            onMouseDown={handleMouseDown}
            onMouseMove={handleMouseMove}
            onMouseUp={handleMouseUp}
            className="relative cursor-crosshair select-none flex-shrink-0 m-auto page-wrap"
          >
            <canvas ref={canvasRef} className="block shadow-2xl bg-white max-w-full" />
            
            {/* Draw Area Box */}
            <div className="absolute inset-0 pointer-events-none">
              {currentBox && (
                <div
                  className="absolute border-2 border-[#38bdf8] bg-[#38bdf8]/10 rounded-xs shadow-[0_0_8px_rgba(56,189,248,0.4)]"
                  style={{
                    left: currentBox.x,
                    top: currentBox.y,
                    width: currentBox.w,
                    height: currentBox.h,
                  }}
                >
                  <div className="absolute right-0 bottom-0 bg-[#38bdf8] text-black text-[8px] px-1 font-bold">
                    Crop Region
                  </div>
                  {/* Select handles */}
                  <div className="absolute -top-1 -left-1 w-2.5 h-2.5 bg-[#38bdf8] rounded-full border border-white"></div>
                  <div className="absolute -top-1 -right-1 w-2.5 h-2.5 bg-[#38bdf8] rounded-full border border-white"></div>
                  <div className="absolute -bottom-1 -left-1 w-2.5 h-2.5 bg-[#38bdf8] rounded-full border border-white"></div>
                  <div className="absolute -bottom-1 -right-1 w-2.5 h-2.5 bg-[#38bdf8] rounded-full border border-white"></div>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* ── Extractions Panel (Right) ── */}
      <aside className={`w-56 bg-[#1e293b] border-l border-[#334155] flex flex-col overflow-hidden ext-panel ${figures.length > 0 ? 'open' : ''}`}>
        <div className="p-3 border-b border-[#334155] flex items-center justify-between flex-shrink-0 ext-panel-hdr">
          <h4 className="text-[10px] font-black tracking-wider text-[#64748b] uppercase">Extractions</h4>
          <button
            onClick={() => {
              figures.forEach(f => deleteFigure(f.id));
              toast.info('Deleted all extractions');
            }}
            className="border-none bg-[#334155] text-[#94a3b8] text-[9px] font-bold px-2 py-0.5 rounded hover:bg-[#ef4444] hover:text-white cursor-pointer transition-colors btn-clear-all"
          >
            Delete all
          </button>
        </div>

        <div className="p-2 border-b border-[#334155] bg-[#1e293b] flex items-center justify-between text-[9.5px]">
          <label className="flex items-center space-x-1.5 cursor-pointer text-slate-300 select-none">
            <input
              type="checkbox"
              checked={showOnlyMentioned}
              onChange={(e) => setShowOnlyMentioned(e.target.checked)}
              className="rounded bg-[#0f172a] border-[#334155] text-[#38bdf8] focus:ring-0 focus:ring-offset-0 w-3 h-3 cursor-pointer"
            />
            <span>Show mentioned only ({filteredCaptions.length})</span>
          </label>
        </div>

        <div className="flex-1 overflow-y-auto p-2.5 space-y-3.5 ext-list">
          {figures.map((fig) => {
            // Captions already picked by OTHER figures — exclude from this card's dropdown
            const usedCaptions = new Set(
              figures.filter(f => f.id !== fig.id && f.caption).map(f => f.caption!)
            );
            const availableCaptions = filteredCaptions.filter(
              c => !usedCaptions.has(c.text) || c.text === fig.caption
            );

            const figuresList = availableCaptions.filter(c => {
              const lbl = (c.label || '').toLowerCase();
              return lbl.startsWith('figure') || lbl.startsWith('fig');
            });
            const tablesList = availableCaptions.filter(c => {
              const lbl = (c.label || '').toLowerCase();
              return lbl.startsWith('table') || lbl.startsWith('tab');
            });
            const othersList = availableCaptions.filter(c => {
              const lbl = (c.label || '').toLowerCase();
              return !lbl.startsWith('figure') && !lbl.startsWith('fig') && !lbl.startsWith('table') && !lbl.startsWith('tab');
            });

            return (
              <div key={fig.id} className="rounded-md border border-[#334155] bg-[#0f172a] flex flex-col overflow-hidden ext-card">
                <div className="h-28 bg-[#0f172a] flex items-center justify-center p-1.5 ext-card-img">
                  <img src={fig.url} alt={fig.name} className="max-w-full max-h-full object-contain" />
                </div>
                <div className="px-2 pt-2 pb-0 space-y-2">
                  <select
                    value={fig.caption || ""}
                    onChange={(e) => {
                      const selectedVal = e.target.value;
                      const matchingCaption = pdfCaptions.find(c => c.text === selectedVal);
                      if (matchingCaption) {
                        renameFigure(fig.id, matchingCaption.label);
                        updateFigureCaption(fig.id, matchingCaption.text, matchingCaption.credit, matchingCaption.runs, matchingCaption.creditRuns);
                        // Auto-sync alt text from Excel whenever caption changes
                        const altEntry = resolveAltEntry(matchingCaption.label);
                        if (altEntry) {
                          updateFigureAltText(fig.id, altEntry.alt_text_short);
                        }
                      } else {
                        updateFigureCaption(fig.id, "", "");
                      }
                    }}
                    className={`w-full bg-[#0f172a] rounded px-1.5 py-1 text-[9.5px] outline-none text-ellipsis overflow-hidden whitespace-nowrap border ${
                      fig.caption
                        ? 'border-[#334155] text-[#e2e8f0] focus:border-[#38bdf8]'
                        : 'border-amber-500/60 text-amber-400 focus:border-amber-400'
                    }`}
                  >
                    <option value="">-- Select Caption --</option>
                    {figuresList.length > 0 && (
                      <optgroup label="Figures" className="bg-[#1e293b] text-slate-300 font-semibold text-[9px]">
                        {figuresList.map((cap) => (
                          <option key={cap.id} value={cap.text} className="bg-[#0f172a] text-[#e2e8f0] text-[9px]">
                            {cap.label}: {cap.text.length > 25 ? cap.text.substring(0, 25) + '...' : cap.text}
                          </option>
                        ))}
                      </optgroup>
                    )}
                    {tablesList.length > 0 && (
                      <optgroup label="Tables" className="bg-[#1e293b] text-slate-300 font-semibold text-[9px]">
                        {tablesList.map((cap) => (
                          <option key={cap.id} value={cap.text} className="bg-[#0f172a] text-[#e2e8f0] text-[9px]">
                            {cap.label}: {cap.text.length > 25 ? cap.text.substring(0, 25) + '...' : cap.text}
                          </option>
                        ))}
                      </optgroup>
                    )}
                    {othersList.length > 0 && (
                      <optgroup label="Others" className="bg-[#1e293b] text-slate-300 font-semibold text-[9px]">
                        {othersList.map((cap) => (
                          <option key={cap.id} value={cap.text} className="bg-[#0f172a] text-[#e2e8f0] text-[9px]">
                            {cap.label}: {cap.text.length > 25 ? cap.text.substring(0, 25) + '...' : cap.text}
                          </option>
                        ))}
                      </optgroup>
                    )}
                  </select>

                {fig.credit && (
                  <div className="text-[8.5px] text-slate-400 bg-slate-900/50 p-1.5 rounded border border-slate-800/80 leading-normal">
                    <span className="font-semibold text-slate-500 block text-[7px] uppercase tracking-wider mb-0.5">Credit</span>
                    {fig.credit}
                  </div>
                )}

                {/* Alt Text block — figures only, not tables */}
                {!fig.name.toLowerCase().startsWith('table') && (
                  <AltTextField
                    figId={fig.id}
                    stored={fig.alt_text}
                    excelEntry={resolveAltEntry(fig.name)}
                    onUpdate={updateFigureAltText}
                  />
                )}
              </div>
              <div className="flex items-center gap-1.5 p-2 bg-[#1e293b] border-t border-[#334155] ext-card-foot">
                <input
                  type="text"
                  value={fig.name}
                  onChange={(e) => renameFigure(fig.id, e.target.value)}
                  className="flex-1 bg-[#0f172a] border border-[#334155] rounded px-1.5 py-0.5 text-[#e2e8f0] text-[9.5px] font-mono outline-none focus:border-[#38bdf8] name-input"
                />
                <button
                  onClick={() => {
                    deleteFigure(fig.id);
                    toast.success('Deleted extraction');
                  }}
                  className="border-none bg-none text-slate-500 hover:text-red-500 font-bold text-sm cursor-pointer px-1 flex items-center justify-center btn-del"
                  title="Delete figure"
                >
                  ×
                </button>
              </div>
            </div>
            );
          })}
        </div>

        <div className="p-3 border-t border-[#334155] bg-[#1e293b] flex-shrink-0">
          <button
            onClick={() => {
              const missing = figures.filter(f => !f.alt_text?.trim());
              if (missing.length > 0) {
                setShowAltWarning(true);
              } else {
                convertDeck(4);
              }
            }}
            disabled={isConverting}
            className={`w-full py-2.5 font-bold rounded text-xs transition-all shadow-md flex items-center justify-center gap-2 ${
              isConverting
                ? 'bg-[#0369a1] text-white/70 cursor-not-allowed'
                : 'bg-[#0284c7] hover:bg-[#0369a1] text-white cursor-pointer'
            }`}
          >
            {isConverting ? (
              <>
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
                Processing… {conversionProgress}%
              </>
            ) : (
              'Proceed to Review'
            )}
          </button>
        </div>
      </aside>

      {/* Full-screen processing overlay */}
      {/* Alt Text warning modal */}
      {showAltWarning && (() => {
        const missing = figures.filter(f => !f.alt_text?.trim());
        return (
          <div className="absolute inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center">
            <div className="bg-[#1e293b] border border-[#334155] rounded-xl shadow-2xl w-80 p-5 space-y-4">
              <div className="flex items-start gap-3">
                <span className="text-amber-400 text-xl mt-0.5">⚠</span>
                <div>
                  <h3 className="text-sm font-bold text-[#f1f5f9]">Missing Alt Text</h3>
                  <p className="text-[11px] text-[#94a3b8] mt-1 leading-relaxed">
                    {missing.length} figure{missing.length > 1 ? 's are' : ' is'} missing alt text:
                  </p>
                  <ul className="mt-1.5 space-y-0.5 max-h-28 overflow-y-auto">
                    {missing.map(f => (
                      <li key={f.id} className="text-[10px] text-amber-300 font-mono">• {f.name || 'Unnamed'}</li>
                    ))}
                  </ul>
                </div>
              </div>
              <div className="flex gap-2 pt-1">
                <button
                  onClick={() => setShowAltWarning(false)}
                  className="flex-1 py-2 rounded text-xs font-bold bg-[#334155] hover:bg-[#475569] text-[#e2e8f0] cursor-pointer border-none transition-colors"
                >
                  Go Back
                </button>
                <button
                  onClick={() => { setShowAltWarning(false); convertDeck(4); }}
                  className="flex-1 py-2 rounded text-xs font-bold bg-[#0284c7] hover:bg-[#0369a1] text-white cursor-pointer border-none transition-colors"
                >
                  Proceed Anyway
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {isConverting && (
        <div className="absolute inset-0 z-50 bg-white/80 backdrop-blur-sm flex flex-col items-center justify-center gap-6">
          <Loader2 className="w-14 h-14 text-[#0284c7] animate-spin" />
          <div className="text-center space-y-3 w-72">
            <h3 className="text-base font-bold text-[var(--color-navy)]">Applying Layout Styles</h3>
            <p className="text-xs text-[var(--color-muted)]">Formatting shapes, fonts and inserting figures…</p>
            <div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden border border-slate-200">
              <div
                className="bg-[#0284c7] h-full transition-all duration-300 ease-out"
                style={{ width: `${conversionProgress}%` }}
              />
            </div>
            <p className="text-xs font-bold text-[#0284c7]">{conversionProgress}%</p>
          </div>
        </div>
      )}
    </div>
  );
};
