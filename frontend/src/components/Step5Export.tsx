import React, { useState, useEffect } from 'react';
import { useStore } from '../store';
import { Download, RefreshCw, AlertTriangle, FileJson, Check, LayoutGrid, BarChart2, ImageOff, Activity, ChevronDown, ChevronUp, FileWarning } from 'lucide-react';
import { toast } from 'sonner';

const getNormalizedRefName = (text: string) => {
  const match = text.match(/\binsert\s+(figure|fig\.?|f\.?)\s*([\d.-]+)/i);
  if (match) {
    return `figure ${match[2]}`;
  }
  const matchTab = text.match(/\binsert\s+(table|tab\.?|t\.?)\s*([\d.-]+)/i);
  if (matchTab) {
    return `table ${matchTab[2]}`;
  }
  return null;
};

const isShapeMissingFigure = (text: string, slideIndex: number, figures: any[]) => {
  const clean = text.trim();
  const refName = getNormalizedRefName(clean);
  if (refName) {
    const isMapped = figures.some(
      (f) => f.name.toLowerCase().trim() === refName && f.mappedTo?.slideIndex === slideIndex
    );
    return !isMapped;
  }
  return /\binsert\s+(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart|image)(?:\s*[\d.-]+)?/i.test(clean);
};

export const Step5Export: React.FC = () => {
  const {
    figures,
    slides,
    resetSession,
    setStep,
    setCurrentSlideIndex,
    setFocusedShapeIndex,
    customerName,
    projectName
  } = useStore();

  const [reportTab, setReportTab] = useState<'style' | 'figures' | 'accessibility' | 'content'>('style');
  const [expandedSubCategories, setExpandedSubCategories] = useState<Record<string, boolean>>({});
  const [selectedSeverity, setSelectedSeverity] = useState<string>('Error');
  const [figureDiag, setFigureDiag] = useState<{ missing: string[]; unplaced: string[] } | null>(null);
  const [accessibilityDiag, setAccessibilityDiag] = useState<{ issues: string[] } | null>(null);
  const [reportData, setReportData] = useState<any[]>([]);
  const [contentLoss, setContentLoss] = useState<{
    slides: { slide: number; missing_text: string[] }[];
    input_slide_count: number;
    output_slide_count: number;
    missing_slide_count: number;
  } | null>(null);
  const [activeFilter, setActiveFilter] = useState<string | null>(null);

  useEffect(() => {
    fetch('/api/figure-diagnostics')
      .then(r => r.json())
      .then(d => { if (d.ok) setFigureDiag({ missing: d.missing, unplaced: d.unplaced }); })
      .catch(() => {});

    fetch('/api/accessibility-report')
      .then(r => r.json())
      .then(d => { if (d.ok) setAccessibilityDiag({ issues: d.issues }); })
      .catch(() => {});

    fetch('/api/report-data')
      .then(r => r.json())
      .then(d => { if (d.ok) setReportData(d.changes); })
      .catch(() => {});

    fetch('/api/content-loss-report')
      .then(r => r.json())
      .then(d => {
        if (d.ok) {
          setContentLoss({
            slides: d.slides,
            input_slide_count: d.input_slide_count,
            output_slide_count: d.output_slide_count,
            missing_slide_count: d.missing_slide_count,
          });
        }
      })
      .catch(() => {});
  }, []);

  const totalSlides = slides?.length || 0;
  
  // Calculate mapped figures and unmapped figures
  const mappedCount = figures.filter((f) => f.mappedTo !== null).length;
  const unmappedCount = figures.length - mappedCount;

  // Calculate missing placeholders count
  let emptyPlaceholdersCount = 0;
  slides?.forEach((slide) => {
    slide.shapes.forEach((shape) => {
      const isImgPlaceholder =
        (shape.shapeType && (shape.shapeType === '13' || shape.shapeType.includes('Picture') || shape.shapeType.includes('Image'))) ||
        (shape.placeholder?.type && (shape.placeholder.type.includes('PICTURE') || shape.placeholder.type.includes('BITMAP')));
      if (isImgPlaceholder && !shape.imageUrl) {
        emptyPlaceholdersCount++;
      } else if (!shape.imageUrl) {
        const text = (shape.textBody?.paragraphs || [])
          .map((para: any) => para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '')
          .join(' ')
          .trim();
        const slideIndex = (slides || []).findIndex((s) => s.slide_id === slide.slide_id);
        const isFigRef = isShapeMissingFigure(text, slideIndex, figures);
        if (isFigRef) {
          emptyPlaceholdersCount++;
        }
      }
    });
  });

  const handleExportPpt = () => {
    toast.success("Downloading final styled presentation...");
    window.open('/api/download', '_blank');
  };

  const handleDownloadExcel = () => {
    toast.success("Generating and downloading Excel report...");
    const url = `/api/download-excel?customerName=${encodeURIComponent(customerName)}&projectName=${encodeURIComponent(projectName)}`;
    window.open(url, '_blank');
  };

  const handleDownloadMappingJson = () => {
    try {
      const mapping = figures.map((fig) => ({
        figureName: fig.name,
        filename: fig.filename,
        mappedTo: fig.mappedTo
          ? {
              slideIndex: fig.mappedTo.slideIndex,
              shapeIndex: fig.mappedTo.shapeIndex,
            }
          : null,
      }));

      const blob = new Blob([JSON.stringify(mapping, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'deckforge_mappings.json';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      toast.success("Mapping JSON downloaded successfully!");
    } catch (err) {
      toast.error("Failed to generate mapping file.");
    }
  };

  const handleNavigateToSlide = (slideIdx: number, shapeIdx?: number) => {
    setCurrentSlideIndex(slideIdx);
    if (shapeIdx !== undefined) {
      setFocusedShapeIndex(shapeIdx);
    } else {
      setFocusedShapeIndex(null);
    }
    setStep(4);
    toast.info(`Navigated to Slide ${slideIdx + 1}`);
  };

  return (
    <div className="space-y-8 max-w-5xl mx-auto pb-8">
      {/* Header Banner */}
      <div className="text-left space-y-1.5 border-b border-[var(--color-border)] pb-4">
        <div className="flex items-center space-x-2">
          <span className="text-[10px] font-bold bg-[var(--color-amber)] text-white px-2.5 py-0.5 rounded-full uppercase tracking-wider">
            Ready for Export
          </span>
        </div>
        <h2 className="text-2xl font-black text-[var(--color-navy)] tracking-tight">
          Review & Download Deck
        </h2>
        <p className="text-[var(--color-muted)] text-xs">
          Inspect slide completion validation, download your styled PowerPoint, or save the figure mapping JSON.
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Side: Summary & Actions (2 Cols) */}
        <div className="lg:col-span-2 space-y-6">
          <div className="surface-card p-6 space-y-4 shadow-sm border border-[var(--color-border)]">
            <h3 className="text-sm font-bold uppercase tracking-wider text-[var(--color-navy)] flex items-center space-x-2">
              <Check className="w-4 h-4 text-emerald-500" />
              <span>Compilation Validation Summary</span>
            </h3>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {/* Card 1: Total Slides */}
              <div className="p-4 bg-[var(--color-cream)]/30 border border-[var(--color-border)] rounded-[var(--radius-custom)] flex items-center space-x-3.5">
                <div className="w-10 h-10 rounded-lg bg-[var(--color-navy)]/10 text-[var(--color-navy)] flex items-center justify-center flex-shrink-0">
                  <FileJson className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-[10px] font-bold uppercase tracking-wider text-[var(--color-muted)]">Total Styled Slides</div>
                  <div className="text-xl font-black text-[var(--color-navy)] mt-0.5">{totalSlides}</div>
                </div>
              </div>

              {/* Card 2: Figures Mapped */}
              <div className="p-4 bg-[var(--color-cream)]/30 border border-[var(--color-border)] rounded-[var(--radius-custom)] flex items-center space-x-3.5">
                <div className="w-10 h-10 rounded-lg bg-emerald-50 text-[var(--color-success)] flex items-center justify-center flex-shrink-0">
                  <Check className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-[10px] font-bold uppercase tracking-wider text-[var(--color-muted)]">Auto-Inserted Figures</div>
                  <div className="text-xl font-black text-[var(--color-success)] mt-0.5">{mappedCount}</div>
                </div>
              </div>

              {/* Card 3: Empty Boxes */}
              <div className="p-4 bg-[var(--color-cream)]/30 border border-[var(--color-border)] rounded-[var(--radius-custom)] flex items-center space-x-3.5">
                <div className={`w-10 h-10 rounded-lg flex items-center justify-center flex-shrink-0 ${emptyPlaceholdersCount > 0 ? 'bg-amber-50 text-amber-600' : 'bg-emerald-50 text-emerald-600'}`}>
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-[10px] font-bold uppercase tracking-wider text-[var(--color-muted)]">Empty Image Placeholders</div>
                  <div className={`text-xl font-black mt-0.5 ${emptyPlaceholdersCount > 0 ? 'text-amber-600' : 'text-emerald-600'}`}>{emptyPlaceholdersCount}</div>
                </div>
              </div>

              {/* Card 4: Unused crops */}
              <div className="p-4 bg-[var(--color-cream)]/30 border border-[var(--color-border)] rounded-[var(--radius-custom)] flex items-center space-x-3.5">
                <div className="w-10 h-10 rounded-lg bg-zinc-100 text-zinc-500 flex items-center justify-center flex-shrink-0">
                  <Download className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-[10px] font-bold uppercase tracking-wider text-[var(--color-muted)]">Unused Cropped Crops</div>
                  <div className="text-xl font-black text-zinc-500 mt-0.5">{unmappedCount}</div>
                </div>
              </div>
            </div>
          </div>

          {/* Quick Action Deck */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            <button
              onClick={handleExportPpt}
              className="py-4 px-4 bg-[var(--color-navy)] hover:bg-[var(--color-navy-light)] text-white font-bold rounded-[var(--radius-custom)] transition-all cursor-pointer shadow-md flex items-center justify-center space-x-2 active:scale-[0.99] text-xs"
            >
              <Download className="w-4 h-4" />
              <span>Download Presentation</span>
            </button>

            <button
              onClick={handleDownloadExcel}
              className="py-4 px-4 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-[var(--radius-custom)] transition-all cursor-pointer shadow-md flex items-center justify-center space-x-2 active:scale-[0.99] text-xs"
            >
              <Download className="w-4 h-4" />
              <span>Download Excel Report</span>
            </button>

            <button
              onClick={handleDownloadMappingJson}
              className="py-4 px-4 bg-white hover:bg-[var(--color-cream)] border border-[var(--color-border)] text-[var(--color-navy)] font-bold rounded-[var(--radius-custom)] transition-all cursor-pointer shadow-sm flex items-center justify-center space-x-2 active:scale-[0.99] text-xs"
            >
              <FileJson className="w-4 h-4 text-[var(--color-amber)]" />
              <span>Save Mappings JSON</span>
            </button>
          </div>

          <button
            onClick={() => {
              resetSession();
              toast.info("Session reset. Ready for a new PPT formatting.");
            }}
            className="w-full py-3 bg-white hover:bg-neutral-50 border border-[var(--color-border)] text-zinc-600 font-semibold rounded-[var(--radius-custom)] transition-all cursor-pointer flex items-center justify-center space-x-2 text-xs"
          >
            <RefreshCw className="w-3.5 h-3.5 text-zinc-400" />
            <span>Reset and Start New Compilation</span>
          </button>
        </div>

        {/* Right Side: Slide Checklist (1 Col) */}
        <div className="lg:col-span-1 surface-card p-6 flex flex-col h-full max-h-[380px] border border-[var(--color-border)] shadow-sm">
          <h3 className="text-xs font-black uppercase tracking-wider text-[var(--color-navy)] pb-2 border-b border-[var(--color-border)] mb-3 flex items-center space-x-2">
            <LayoutGrid className="w-4 h-4 text-[var(--color-amber)]" />
            <span>Slide Validation Checklist</span>
          </h3>
          <div className="flex-1 overflow-y-auto space-y-2.5 pr-1">
            {slides?.map((slide, idx) => {
              const missingRefs: { label: string; shapeIndex: number }[] = [];
              slide.shapes.forEach((s) => {
                const isPicturePh = ((s.shapeType && (s.shapeType === '13' || s.shapeType.includes('Picture') || s.shapeType.includes('Image'))) ||
                  (s.placeholder?.type && (s.placeholder.type.includes('PICTURE') || s.placeholder.type.includes('BITMAP'))));
                  
                if (isPicturePh && !s.imageUrl) {
                  missingRefs.push({ label: "Picture Box", shapeIndex: s.index });
                } else if (!s.imageUrl) {
                  const text = (s.textBody?.paragraphs || [])
                    .map((para: any) => para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '')
                    .join(' ')
                    .trim();
                  
                  const isInsertRef = /^insert\s+(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart|image)(?:\s*([\d.-]+))?(?:\s+here)?[.:]?$/i;
                  
                  const match = text.match(isInsertRef);
                  if (match) {
                    const figNum = match[2] || '';
                    missingRefs.push({ label: `${match[1].toUpperCase()}${figNum ? ' ' + figNum : ''}`, shapeIndex: s.index });
                  } else {
                    const isFigRef = isShapeMissingFigure(text, idx, figures);
                    if (isFigRef) {
                      missingRefs.push({ label: text.length > 25 ? text.substring(0, 25) + '...' : text, shapeIndex: s.index });
                    }
                  }
                }
              });

              return (
                <div
                  key={slide.slide_id}
                  onClick={() => handleNavigateToSlide(idx)}
                  className="flex items-center justify-between p-2 border border-[var(--color-border)] rounded-[var(--radius-custom)] bg-white hover:border-amber-400 hover:bg-amber-50/10 cursor-pointer transition-all gap-2"
                >
                  <span className="text-xs font-semibold text-[var(--color-navy)] flex-shrink-0">
                    Slide {idx + 1}
                  </span>
                  {missingRefs.length > 0 ? (
                    <div className="flex flex-col items-end gap-1 flex-1 min-w-0">
                      {missingRefs.map((refObj, rIdx) => (
                        <span
                          key={rIdx}
                          onClick={(e) => {
                            e.stopPropagation();
                            handleNavigateToSlide(idx, refObj.shapeIndex);
                          }}
                          className="text-[8px] bg-rose-50 border border-rose-200 text-rose-700 px-1.5 py-0.5 rounded font-bold uppercase tracking-wide truncate max-w-full text-right hover:bg-rose-100 transition-all cursor-pointer"
                          title={`Click to focus missing ${refObj.label}`}
                        >
                          ⚠️ Missing {refObj.label}
                        </span>
                      ))}
                    </div>
                  ) : (
                    <span className="text-[9px] bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded-full flex items-center space-x-1 font-bold border border-emerald-200">
                      <Check className="w-3 h-3" />
                      <span>Complete</span>
                    </span>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Style & Diagnostics Report — tabbed */}
      <div className="surface-card border border-[var(--color-border)] shadow-sm overflow-hidden">
        {/* Tab Bar */}
        <div className="flex border-b border-[var(--color-border)] bg-zinc-50">
          <button
            onClick={() => setReportTab('style')}
            className={`flex items-center gap-2 px-5 py-3 text-xs font-bold uppercase tracking-wider transition-all border-b-2 ${
              reportTab === 'style'
                ? 'border-[var(--color-amber)] text-[var(--color-navy)] bg-white'
                : 'border-transparent text-zinc-400 hover:text-zinc-600 hover:bg-zinc-100'
            }`}
          >
            <BarChart2 className="w-3.5 h-3.5" />
            Style Change Report
          </button>
          <button
            onClick={() => setReportTab('figures')}
            className={`flex items-center gap-2 px-5 py-3 text-xs font-bold uppercase tracking-wider transition-all border-b-2 relative ${
              reportTab === 'figures'
                ? 'border-[var(--color-amber)] text-[var(--color-navy)] bg-white'
                : 'border-transparent text-zinc-400 hover:text-zinc-600 hover:bg-zinc-100'
            }`}
          >
            <ImageOff className="w-3.5 h-3.5" />
            Figure Diagnostics
            {figureDiag && figureDiag.missing.length > 0 && (
              <span className="ml-1 bg-rose-500 text-white text-[9px] font-black rounded-full px-1.5 py-0.5">
                {figureDiag.missing.length}
              </span>
            )}
          </button>
          <button
            onClick={() => setReportTab('accessibility')}
            className={`flex items-center gap-2 px-5 py-3 text-xs font-bold uppercase tracking-wider transition-all border-b-2 relative ${
              reportTab === 'accessibility'
                ? 'border-[var(--color-amber)] text-[var(--color-navy)] bg-white'
                : 'border-transparent text-zinc-400 hover:text-zinc-600 hover:bg-zinc-100'
            }`}
          >
            <Activity className="w-3.5 h-3.5" />
            Accessibility Report
            {accessibilityDiag && accessibilityDiag.issues.length > 0 && (
              <span className="ml-1 bg-amber-500 text-white text-[9px] font-black rounded-full px-1.5 py-0.5">
                {accessibilityDiag.issues.length}
              </span>
            )}
          </button>
          <button
            onClick={() => setReportTab('content')}
            className={`flex items-center gap-2 px-5 py-3 text-xs font-bold uppercase tracking-wider transition-all border-b-2 relative ${
              reportTab === 'content'
                ? 'border-[var(--color-amber)] text-[var(--color-navy)] bg-white'
                : 'border-transparent text-zinc-400 hover:text-zinc-600 hover:bg-zinc-100'
            }`}
          >
            <FileWarning className="w-3.5 h-3.5" />
            Content Loss Check
            {contentLoss && (contentLoss.slides.length > 0 || contentLoss.missing_slide_count > 0) && (
              <span className="ml-1 bg-rose-500 text-white text-[9px] font-black rounded-full px-1.5 py-0.5">
                {contentLoss.slides.length + contentLoss.missing_slide_count}
              </span>
            )}
          </button>
        </div>

        {/* Content Loss Tab */}
        {reportTab === 'content' && (
          <div className="p-6 space-y-6 max-h-[620px] overflow-y-auto">
            <div>
              <h4 className="flex items-center gap-2 text-sm font-black text-rose-600 uppercase tracking-wider mb-2">
                <FileWarning className="w-4 h-4" />
                Input vs. Output Content Loss
              </h4>
              <p className="text-xs text-[var(--color-muted)] mb-4">
                Flags text present in your uploaded PPTX that's missing from the styled output —
                a sign the conversion dropped content, not just restyled it.
              </p>

              {!contentLoss ? (
                <div className="text-xs text-zinc-400 italic py-6 text-center">Loading…</div>
              ) : (
                <>
                  {contentLoss.missing_slide_count > 0 && (
                    <div className="mb-4 p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800">
                      <span className="font-bold">{contentLoss.missing_slide_count}</span> slide(s) from the
                      input ({contentLoss.input_slide_count} total) are missing entirely from the output
                      ({contentLoss.output_slide_count} total).
                    </div>
                  )}

                  {contentLoss.slides.length === 0 && contentLoss.missing_slide_count === 0 ? (
                    <div className="flex items-center gap-2 text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg p-3">
                      <Check className="w-4 h-4" />
                      No content loss detected — every text fragment from the input appears in the output.
                    </div>
                  ) : (
                    <div className="space-y-2">
                      {contentLoss.slides.map((s) => (
                        <div key={s.slide} className="p-3 border border-rose-100 bg-rose-50/30 rounded-lg">
                          <div className="text-[11px] font-black text-rose-700 uppercase tracking-wide mb-1.5">
                            Slide {s.slide}
                          </div>
                          <ul className="space-y-1">
                            {s.missing_text.map((t, i) => (
                              <li key={i} className="text-xs text-zinc-700 pl-3 border-l-2 border-rose-300">
                                {t}
                              </li>
                            ))}
                          </ul>
                        </div>
                      ))}
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        )}

        {/* Style Report Tab */}
        {reportTab === 'style' && (() => {
          const TAG_MAP: Record<string, { cls: string, label: string }> = {
            'Font family': { cls: 'bg-blue-50 text-blue-700 border border-blue-200', label: 'Font' },
            'Font size': { cls: 'bg-amber-50 text-amber-700 border border-amber-200', label: 'Size' },
            'Color': { cls: 'bg-purple-50 text-purple-700 border border-purple-200', label: 'Color' },
            'Bold': { cls: 'bg-gray-50 text-gray-700 border border-gray-200', label: 'Bold' },
            'Alignment': { cls: 'bg-cyan-50 text-cyan-700 border border-cyan-200', label: 'Align' },
            'Line spacing': { cls: 'bg-slate-50 text-slate-700 border border-slate-200', label: 'Line Spc' },
            'Space before': { cls: 'bg-slate-50 text-slate-700 border border-slate-200', label: 'Spc Before' },
            'Space after': { cls: 'bg-slate-50 text-slate-700 border border-slate-200', label: 'Spc After' }
          };

          const FILTER_MAP: Record<string, string[]> = {
            'Font': ['Font family'],
            'Size': ['Font size'],
            'Color': ['Color'],
            'Bold': ['Bold'],
            'Alignment': ['Alignment'],
            'Spacing': ['Line spacing', 'Space before', 'Space after']
          };

          const phLabel = (t: string) => {
            if (/center.title/i.test(t)) return 'Center Title';
            if (/subtitle/i.test(t)) return 'Subtitle';
            if (/title/i.test(t)) return 'Title';
            if (/body|object|text/i.test(t)) return 'Body/Content';
            return t;
          };

          // Calculate stats
          let totalChanges = 0;
          const propertiesChanged = new Set<string>();
          reportData.forEach((s: any) =>
            s.placeholders.forEach((p: any) =>
              p.paras.forEach((q: any) => {
                totalChanges += q.changes.length;
                q.changes.forEach((c: any) => propertiesChanged.add(c.prop));
              })
            )
          );

          // Filter data
          const filtered = reportData.map((slide: any) => {
            if (!activeFilter) return slide;
            const allowedProps = FILTER_MAP[activeFilter];

            const newPlaceholders = slide.placeholders.map((ph: any) => {
              const newParas = ph.paras.map((para: any) => {
                const newChanges = para.changes.filter((c: any) => allowedProps.includes(c.prop));
                if (newChanges.length > 0) {
                  return { ...para, changes: newChanges };
                }
                return null;
              }).filter(Boolean);

              if (newParas.length > 0) {
                return { ...ph, paras: newParas };
              }
              return null;
            }).filter(Boolean);

            if (newPlaceholders.length > 0) {
              return { ...slide, placeholders: newPlaceholders };
            }
            return null;
          }).filter(Boolean);

          return (
            <div className="flex flex-col h-[620px] bg-white">
              {/* Header Stats */}
              <div className="flex items-center justify-between px-6 py-4 bg-zinc-900 text-white rounded-t-sm shadow-sm">
                <div>
                  <h3 className="text-sm font-bold tracking-wide uppercase">Style & Layout Differences</h3>
                  <div className="text-[10px] text-zinc-400 mt-0.5">Comparing template cascades to content outputs</div>
                </div>
                <div className="flex gap-6 text-center">
                  <div>
                    <div className="text-lg font-black font-mono leading-none">{reportData.length}</div>
                    <div className="text-[9px] text-zinc-400 uppercase tracking-wider mt-1">Slides Changed</div>
                  </div>
                  <div className="border-r border-zinc-800 h-6 self-center" />
                  <div>
                    <div className="text-lg font-black font-mono leading-none">{totalChanges}</div>
                    <div className="text-[9px] text-zinc-400 uppercase tracking-wider mt-1">Total Diff</div>
                  </div>
                  <div className="border-r border-zinc-800 h-6 self-center" />
                  <div>
                    <div className="text-lg font-black font-mono leading-none">{propertiesChanged.size}</div>
                    <div className="text-[9px] text-zinc-400 uppercase tracking-wider mt-1">Props Affected</div>
                  </div>
                </div>
              </div>

              {/* Filter Bar */}
              <div className="flex items-center gap-2 px-6 py-3 border-b border-zinc-200 bg-zinc-50 overflow-x-auto text-xs">
                <span className="font-bold text-zinc-500 uppercase tracking-wider text-[10px] mr-1">Filter changes:</span>
                {['Font', 'Size', 'Color', 'Bold', 'Alignment', 'Spacing'].map((filterName) => {
                  const isActive = activeFilter === filterName;
                  return (
                    <button
                      key={filterName}
                      onClick={() => setActiveFilter(isActive ? null : filterName)}
                      className={`px-2.5 py-1 rounded text-xs font-bold transition-all ${
                        isActive
                          ? 'bg-zinc-900 text-white shadow-sm'
                          : 'bg-white border border-zinc-200 text-zinc-600 hover:bg-zinc-100'
                      }`}
                    >
                      {filterName}
                    </button>
                  );
                })}
                {activeFilter && (
                  <button
                    onClick={() => setActiveFilter(null)}
                    className="ml-auto text-[10px] font-bold text-zinc-400 hover:text-zinc-600 uppercase"
                  >
                    Clear Filter
                  </button>
                )}
              </div>

              {/* Content Scrollable List */}
              <div className="flex-1 overflow-y-auto p-6 space-y-4 bg-zinc-50/50">
                {filtered.length === 0 ? (
                  <div className="text-center py-12 text-zinc-400 text-xs italic">
                    No visual style differences match the selected filter.
                  </div>
                ) : (
                  filtered.map((slide: any) => {
                    const slideChangesCount = slide.placeholders.reduce(
                      (acc: number, ph: any) => acc + ph.paras.reduce((pAcc: number, para: any) => pAcc + para.changes.length, 0),
                      0
                    );

                    return (
                      <div key={slide.slide} className="bg-white border border-zinc-200 rounded-md shadow-sm overflow-hidden">
                        {/* Slide Header */}
                        <div className="flex items-center justify-between px-4 py-2.5 bg-zinc-900 text-white">
                          <div className="flex items-center gap-2">
                            <span className="text-[10px] font-bold uppercase tracking-wider border border-white/20 rounded px-1.5 py-0.5">
                              Slide {slide.slide}
                            </span>
                            <span className="text-[11px] text-zinc-300">Layout styling cascade updated</span>
                          </div>
                          <span className="text-[10px] font-mono text-zinc-400">
                            {slideChangesCount} change{slideChangesCount !== 1 ? 's' : ''}
                          </span>
                        </div>

                        {/* Slide Placeholders */}
                        {slide.placeholders.map((ph: any, phIdx: number) => (
                          <div key={phIdx} className="border-t border-zinc-100 first:border-t-0">
                            <div className="flex items-center gap-2 px-4 py-2 bg-zinc-50/70 border-b border-zinc-100">
                              <span className="text-[10px] font-extrabold uppercase tracking-wide text-amber-600 bg-amber-50 border border-amber-200/60 rounded px-1.5 py-0.5">
                                {phLabel(ph.type)}
                              </span>
                              <span className="text-[10px] font-mono text-zinc-400">{ph.name}</span>
                            </div>

                            {/* Paragraphs */}
                            {ph.paras.map((para: any, paraIdx: number) => (
                              <div key={paraIdx} className="p-4 border-b border-zinc-50 last:border-b-0">
                                <div className="text-xs text-zinc-400 italic mb-2 relative pl-3 border-l-2 border-zinc-200">
                                  {para.text ? `"${para.text}${para.text.length >= 55 ? '...' : ''}"` : '(empty paragraph)'}
                                </div>
                                <div className="space-y-1.5">
                                  {para.changes.map((c: any, cIdx: number) => {
                                    const tagInfo = TAG_MAP[c.prop] || { cls: 'bg-zinc-50 text-zinc-600 border border-zinc-200', label: c.prop };
                                    const isColor = c.is_color && c.after && c.after.startsWith('#');

                                    return (
                                      <div key={cIdx} className="grid grid-cols-[70px_1fr_12px_1fr] gap-2 items-center text-xs">
                                        <span className={`text-[9px] font-bold uppercase tracking-wider text-center py-0.5 rounded truncate ${tagInfo.cls}`}>
                                          {tagInfo.label}
                                        </span>
                                        <span className="text-zinc-400 line-through truncate flex items-center gap-1">
                                          {c.is_color && c.before && c.before.startsWith('#') && (
                                            <span className="w-2.5 h-2.5 rounded-full border border-black/10 flex-shrink-0" style={{ backgroundColor: c.before }} />
                                          )}
                                          {c.before || '—'}
                                        </span>
                                        <span className="text-zinc-300 font-bold text-center">→</span>
                                        <span className="text-emerald-600 font-semibold truncate flex items-center gap-1">
                                          {isColor && (
                                            <span className="w-2.5 h-2.5 rounded-full border border-black/10 flex-shrink-0" style={{ backgroundColor: c.after }} />
                                          )}
                                          {c.after || '—'}
                                        </span>
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            ))}
                          </div>
                        ))}
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          );
        })()}

        {/* Figure Diagnostics Tab */}
        {reportTab === 'figures' && (
          <div className="p-6 space-y-6">
            {/* Skipped / Missing figures */}
            <div>
              <h4 className="flex items-center gap-2 text-sm font-black text-rose-600 uppercase tracking-wider mb-3">
                <AlertTriangle className="w-4 h-4" />
                Skipped Requested Figures
                <span className="text-[10px] font-semibold text-rose-400 normal-case tracking-normal">(placeholder in slide but no crop provided)</span>
              </h4>
              {figureDiag && figureDiag.missing.length > 0 ? (
                <div className="flex flex-wrap gap-2">
                  {figureDiag.missing.map((name, i) => (
                    <span key={i} className="font-mono text-xs bg-rose-50 border border-rose-200 text-rose-700 px-3 py-1.5 rounded font-bold uppercase">
                      {name}
                    </span>
                  ))}
                </div>
              ) : (
                <p className="text-xs text-emerald-600 font-semibold flex items-center gap-1.5">
                  <Check className="w-3.5 h-3.5" /> All requested figures were successfully placed.
                </p>
              )}
            </div>

            <div className="border-t border-[var(--color-border)]" />

            {/* Unused / Unplaced crops */}
            <div>
              <h4 className="flex items-center gap-2 text-sm font-black text-slate-600 uppercase tracking-wider mb-3">
                <FileJson className="w-4 h-4" />
                Unused Cropped Figures
                <span className="text-[10px] font-semibold text-slate-400 normal-case tracking-normal">(crop exists but no matching slide placeholder)</span>
              </h4>
              {figureDiag && figureDiag.unplaced.length > 0 ? (
                <div className="flex flex-wrap gap-2">
                  {figureDiag.unplaced.map((name, i) => (
                    <span key={i} className="font-mono text-xs bg-slate-50 border border-slate-200 text-slate-600 px-3 py-1.5 rounded font-bold uppercase">
                      {name}
                    </span>
                  ))}
                </div>
              ) : (
                <p className="text-xs text-emerald-600 font-semibold flex items-center gap-1.5">
                  <Check className="w-3.5 h-3.5" /> No unused cropped figures — every crop was placed.
                </p>
              )}
            </div>

            {!figureDiag && (
              <p className="text-xs text-zinc-400 italic">Loading figure diagnostics...</p>
            )}
          </div>
        )}

        {/* Accessibility Report Tab */}
        {reportTab === 'accessibility' && (() => {
          // Group by Severity, then Category using nested structured JSON
          const groupedIssues = (accessibilityDiag?.issues || []).reduce((acc: Record<string, Record<string, any[]>>, issue: any) => {
            const severity = issue.severity || 'Warning';
            const category = issue.category || 'General Issues';
            acc[severity] = acc[severity] || {};
            acc[severity][category] = acc[severity][category] || [];
            acc[severity][category].push({ slide: issue.slide, detail: issue.detail });
            return acc;
          }, {});

          const severities = ['Error', 'Warning', 'Tip'];
          
          // Pre-calculate counts for each severity
          const counts = severities.reduce((acc, sev) => {
            const categories = groupedIssues[sev] || {};
            const totalIssues = Object.values(categories).reduce((sum, list) => sum + list.length, 0);
            const uniqueSlides = Array.from(
              new Set(
                Object.values(categories)
                  .flatMap((list) => list.map((i) => i.slide))
                  .filter((s) => s != null)
              )
            ).length;
            acc[sev] = { totalIssues, uniqueSlides };
            return acc;
          }, {} as Record<string, { totalIssues: number, uniqueSlides: number }>);

          const cardStyles: Record<string, { activeBorder: string, border: string, bg: string, hoverBg: string, activeBg: string, text: string, iconColor: string, title: string }> = {
            'Error': {
              activeBorder: 'border-rose-500 ring-2 ring-rose-500/20',
              border: 'border-rose-100',
              bg: 'bg-rose-50/20',
              hoverBg: 'hover:bg-rose-50/50',
              activeBg: 'bg-rose-50/60',
              text: 'text-rose-800',
              iconColor: 'text-rose-500',
              title: 'Errors'
            },
            'Warning': {
              activeBorder: 'border-amber-500 ring-2 ring-amber-500/20',
              border: 'border-amber-100',
              bg: 'bg-amber-50/20',
              hoverBg: 'hover:bg-amber-50/50',
              activeBg: 'bg-amber-50/60',
              text: 'text-amber-800',
              iconColor: 'text-amber-500',
              title: 'Warnings'
            },
            'Tip': {
              activeBorder: 'border-blue-500 ring-2 ring-blue-500/20',
              border: 'border-blue-100',
              bg: 'bg-blue-50/20',
              hoverBg: 'hover:bg-blue-50/50',
              activeBg: 'bg-blue-50/60',
              text: 'text-blue-800',
              iconColor: 'text-blue-500',
              title: 'Tips'
            }
          };

          return (
            <div className="p-6 space-y-6 max-h-[620px] overflow-y-auto">
              <div>
                <h4 className="flex items-center gap-2 text-sm font-black text-amber-600 uppercase tracking-wider mb-4">
                  <Activity className="w-4 h-4" />
                  Accessibility Issues Found
                </h4>

                {/* Horizontal Cards */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-6">
                  {severities.map((sev) => {
                    const isActive = selectedSeverity === sev;
                    const style = cardStyles[sev];
                    const count = counts[sev];

                    return (
                      <button
                        key={sev}
                        type="button"
                        onClick={() => setSelectedSeverity(sev)}
                        className={`p-4 border rounded-lg text-left transition-all cursor-pointer flex flex-col justify-between h-24 ${
                          isActive 
                            ? `${style.activeBorder} ${style.activeBg}` 
                            : `border-zinc-200 bg-white ${style.hoverBg}`
                        }`}
                      >
                        <div className="flex items-center justify-between w-full">
                          <span className={`text-[11px] font-black uppercase tracking-wider ${style.text}`}>
                            {style.title}
                          </span>
                          <AlertTriangle className={`w-4 h-4 ${style.iconColor}`} />
                        </div>
                        <div className="mt-2">
                          <span className="text-xl font-black text-zinc-800 block leading-none">
                            {count.totalIssues}
                          </span>
                          <span className="text-[10px] text-zinc-400 font-semibold block mt-1">
                            {count.uniqueSlides > 0 ? `on ${count.uniqueSlides} slide${count.uniqueSlides > 1 ? 's' : ''}` : 'no issues found'}
                          </span>
                        </div>
                      </button>
                    );
                  })}
                </div>

                {/* Subcategories list for the selected severity card */}
                {(() => {
                  const categories = groupedIssues[selectedSeverity] || {};
                  const categoryEntries = Object.entries(categories);

                  if (categoryEntries.length === 0) {
                    return (
                      <div className="text-center py-10 border border-dashed border-zinc-200 rounded-lg bg-zinc-50/50">
                        <Check className="w-8 h-8 text-emerald-500 mx-auto mb-2" />
                        <p className="text-xs text-zinc-500 font-bold">No {selectedSeverity.toLowerCase()}s found! Slide content meets standard requirements.</p>
                      </div>
                    );
                  }

                  return (
                    <div className="space-y-3 animate-in fade-in duration-200">
                      {categoryEntries.map(([categoryName, items], cIdx) => {
                        const subKey = `${selectedSeverity}-${categoryName}`;
                        const isSubExpanded = expandedSubCategories[subKey] || false;

                        return (
                          <div key={cIdx} className="border border-zinc-200 rounded-md overflow-hidden bg-white shadow-xs">
                            <button 
                              type="button"
                              onClick={() => setExpandedSubCategories(prev => ({ ...prev, [subKey]: !prev[subKey] }))}
                              className="w-full bg-zinc-50/50 px-3.5 py-2.5 border-b border-zinc-100 flex items-center justify-between hover:bg-zinc-100/30 transition-colors cursor-pointer text-left"
                            >
                              <div className="flex items-center gap-1.5">
                                {isSubExpanded ? <ChevronUp className="w-3.5 h-3.5 text-zinc-400" /> : <ChevronDown className="w-3.5 h-3.5 text-zinc-400" />}
                                <span className="text-[11px] font-bold text-zinc-700 uppercase tracking-wide">{categoryName}</span>
                              </div>
                              <span className="text-[9px] font-bold bg-zinc-200 text-zinc-750 px-2 rounded-full">
                                {items.length} issue{items.length > 1 ? 's' : ''}
                              </span>
                            </button>

                            {isSubExpanded && (
                              <div className="p-3 bg-white border-t border-zinc-50">
                                <ul className="space-y-1.5 list-none pl-0">
                                  {items.map((item: any, mIdx: number) => (
                                    <li
                                      key={mIdx}
                                      onClick={() => item.slide && handleNavigateToSlide(item.slide - 1)}
                                      className={`text-zinc-650 text-[11px] p-1.5 rounded hover:bg-amber-50 hover:text-amber-900 transition-all ${
                                        item.slide ? 'cursor-pointer' : ''
                                      }`}
                                      title={item.slide ? `Click to inspect Slide ${item.slide}` : undefined}
                                    >
                                      {item.slide ? (
                                        <span className="font-bold mr-1 text-[var(--color-navy)] bg-[var(--color-cream)] px-1.5 py-0.5 rounded text-[9px] uppercase tracking-wide">
                                          Slide {item.slide}
                                        </span>
                                      ) : null}
                                      <span className="font-medium">{item.detail}</span>
                                    </li>
                                  ))}
                                </ul>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}
              </div>
            </div>
          );
        })()}
      </div>
    </div>
  );
};
