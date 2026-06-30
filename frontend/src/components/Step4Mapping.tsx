import React, { useRef, useState, useEffect } from 'react';
import { useStore } from '../store';
import { LayoutGrid, Play, AlertTriangle } from 'lucide-react';

export const Step4Mapping: React.FC = () => {
  const {
    slides,
    currentSlideIndex,
    setCurrentSlideIndex,
    setStep
  } = useStore();

  const canvasContainerRef = useRef<HTMLDivElement>(null);
  const [canvasWidth, setCanvasWidth] = useState(600);
  const [canvasHeightBound, setCanvasHeightBound] = useState(400);
  const [reportData, setReportData] = useState<any[]>([]);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [windowSize, setWindowSize] = useState({ width: window.innerWidth, height: window.innerHeight });
  const [showModifications, setShowModifications] = useState(false);
  const [filterMissing, setFilterMissing] = useState(false);

  const slideWidthPt = 960;
  const slideHeightPt = 540;
  const zoom = 1.0;

  const getChecklistItems = (slide: any) => {
    if (!slide?.shapes) return [];
    const items: { label: string; done: boolean; detail?: string }[] = [];
    slide.shapes.forEach((shape: any) => {
      const text = (shape.textBody?.paragraphs || [])
        .map((para: any) => para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '')
        .join(' ')
        .trim();
      const isFigRef = /\b(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart)\s*[\d.]+/i.test(text) || /\binsert\s+(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart|image)(?:\s*[\d.]+)?/i.test(text);
      if (shape.imageUrl) {
        const fileName = shape.imageUrl.split('/').pop() || 'image';
        items.push({
          label: `Mapped Figure / Table`,
          done: true,
          detail: fileName
        });
      } else if (isFigRef) {
        items.push({
          label: text.length > 40 ? text.substring(0, 40) + '...' : text,
          done: false,
          detail: 'Placeholder text not replaced'
        });
      }
    });
    return items;
  };

  useEffect(() => {
    fetch('/api/report-data')
      .then((r) => r.json())
      .then((data) => {
        if (data.ok && data.changes) {
          setReportData(data.changes);
        }
      })
      .catch((err) => console.error("Failed to load report data:", err));
  }, []);

  useEffect(() => {
    const handleResize = () => {
      setWindowSize({ width: window.innerWidth, height: window.innerHeight });
    };
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setIsFullscreen(false);
      }
    };
    window.addEventListener('resize', handleResize);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('resize', handleResize);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, []);

  useEffect(() => {
    const updateSize = () => {
      if (canvasContainerRef.current) {
        setCanvasWidth(canvasContainerRef.current.clientWidth - 16);
        setCanvasHeightBound(canvasContainerRef.current.clientHeight - 72);
      }
    };
    updateSize();
    const timer = setTimeout(updateSize, 150);
    window.addEventListener('resize', updateSize);
    return () => {
      window.removeEventListener('resize', updateSize);
      clearTimeout(timer);
    };
  }, [slides]);

  const scaleX = canvasWidth / slideWidthPt;
  const scaleY = canvasHeightBound / slideHeightPt;
  const baseScale = Math.min(scaleX, scaleY);
  const scale = baseScale * zoom;
  const canvasHeight = slideHeightPt * scale;
  const scaledWidth = slideWidthPt * scale;

  const fsViewportWidth = windowSize.width - 80;
  const fsViewportHeight = windowSize.height - 180;
  const fsScale = Math.min(fsViewportWidth / slideWidthPt, fsViewportHeight / slideHeightPt);

  const activeScale = isFullscreen ? fsScale : scale;
  const activeWidth = isFullscreen ? (slideWidthPt * fsScale) : scaledWidth;
  const activeHeight = isFullscreen ? (slideHeightPt * fsScale) : canvasHeight;

  const currentSlide = slides?.[currentSlideIndex];

  const renderSlideContent = (scaleVal: number, widthVal: number, heightVal: number) => {
    if (!currentSlide) return null;
    return (
      <div
        className="relative shadow-2xl transition-all duration-300 select-none overflow-hidden flex-shrink-0"
        style={{
          width: widthVal,
          height: heightVal,
          minWidth: widthVal,
          minHeight: heightVal,
          backgroundColor: currentSlide.backgroundColor || '#ffffff',
        }}
      >
        {currentSlide.shapes.map((shape) => {
          const left = shape.position.x_pt * scaleVal;
          const top = shape.position.y_pt * scaleVal;
          const width = shape.size.width_pt * scaleVal;
          const height = shape.size.height_pt * scaleVal;

          const isImagePlaceholder =
            (shape.shapeType && (shape.shapeType === '13' || shape.shapeType.includes('Picture') || shape.shapeType.includes('Image'))) ||
            (shape.placeholder?.type && (shape.placeholder.type.includes('PICTURE') || shape.placeholder.type.includes('BITMAP'))) ||
            shape.imageUrl !== undefined;

          // Determine background fill and borders
          const hasFill = shape.fill && shape.fill !== 'none';
          const bgStyle = hasFill ? { backgroundColor: shape.fill } : {};

          return (
            <div
              key={shape.index}
              className={`absolute transition-all select-none overflow-hidden placeholder-box ${isImagePlaceholder
                ? shape.imageUrl
                  ? "border border-emerald-400 bg-emerald-50/5 z-20"
                  : "border-2 border-dashed border-[var(--color-amber)] bg-amber-50/5 z-20"
                : "border border-transparent z-10"
                }`}
              style={{
                left,
                top,
                width,
                height,
                ...bgStyle,
              }}
            >
              {shape.imageUrl ? (
                <div className="w-full h-full bg-black/5 flex items-center justify-center">
                  <img
                    src={shape.imageUrl}
                    alt={shape.shapeName}
                    className="w-full h-full object-contain"
                  />
                </div>
              ) : isImagePlaceholder ? (
                <div className="w-full h-full flex flex-col items-center justify-center text-center p-2 text-[var(--color-amber)] bg-amber-50/5">
                  <span className="text-[9px] font-bold uppercase tracking-wider">
                    Figure Placeholder
                  </span>
                </div>
              ) : (
                <div className="w-full h-full p-2.5 overflow-hidden text-left">
                  {shape.textBody?.paragraphs && shape.textBody.paragraphs.length > 0 ? (
                    <div className="space-y-1">
                      {shape.textBody.paragraphs.map((para: any, pIdx: number) => {
                        const runText = para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '';
                        if (!runText.trim()) return null;

                        const firstRun = para.runs?.[0];
                        const isTitle = shape.shapeName.includes('Title') || shape.placeholder?.type?.includes('TITLE');

                        // Slightly scale down browser fonts by 15% (0.85 factor) to matches PowerPoint native typography margins
                        const fontScaleFactor = 0.85;

                        // Determine bullet level and prefixing
                        const level = para.level || 0;
                        const isBullet = level > 0 || (shape.shapeName.includes('Content') && !isTitle);
                        const indentPadding = level * 16 + (isBullet ? 12 : 0);

                        const paraStyle: React.CSSProperties = {
                          fontSize: firstRun?.fontSize_pt
                            ? `${firstRun.fontSize_pt * scaleVal * fontScaleFactor}px`
                            : `${(isTitle ? 20 : 10.5) * scaleVal}px`,
                          fontFamily: firstRun?.fontFamily || 'inherit',
                          fontWeight: firstRun?.bold || isTitle ? 'bold' : 'normal',
                          color: firstRun?.color || (isTitle ? '#ffffff' : '#1e293b'),
                          textAlign: para.alignment === 'ctr' ? 'center' : para.alignment === 'r' ? 'right' : 'left',
                          paddingLeft: `${indentPadding * scaleVal}px`,
                          textIndent: isBullet ? `-${12 * scaleVal}px` : undefined,
                        };

                        return (
                          <p key={pIdx} style={paraStyle} className="leading-snug whitespace-pre-wrap flex items-start">
                            {isBullet && (
                              <span
                                className="select-none font-bold mr-1.5 inline-block text-center"
                                style={{
                                  width: `${12 * scaleVal}px`,
                                  fontSize: firstRun?.fontSize_pt
                                    ? `${firstRun.fontSize_pt * scaleVal * fontScaleFactor * 0.9}px`
                                    : `${10.5 * scaleVal * 0.9}px`,
                                  color: firstRun?.color || '#475569'
                                }}
                              >
                                •
                              </span>
                            )}
                            <span className="flex-1">{runText}</span>
                          </p>
                        );
                      })}
                    </div>
                  ) : (
                    <span style={{ fontSize: `${9 * scaleVal}px` }} className="text-zinc-400/50 font-medium italic">{shape.shapeName}</span>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div className="space-y-2 max-w-full mx-auto h-[calc(100vh-140px)] flex flex-col">
      {/* Top Header Controls */}
      <div className="flex flex-col md:flex-row md:items-center md:justify-between border-b border-[var(--color-border)] pb-2 gap-2">
        <div className="text-left">
          <h2 className="text-xs font-black uppercase tracking-wider text-[var(--color-navy)]">
            Review Styled Presentation
          </h2>
          <p className="text-[11px] text-[var(--color-muted)] mt-0.5">
            Verify the layout template styling and auto-inserted figures before exporting.
          </p>
        </div>

        <div className="flex items-center space-x-2.5">
          <button
            onClick={() => setStep(3)}
            className="px-4 py-2 border border-[var(--color-border)] hover:bg-[var(--color-cream)] text-[var(--color-navy)] text-xs font-bold rounded-[var(--radius-custom)] transition-all cursor-pointer"
          >
            ← Back
          </button>
          <button
            onClick={() => setStep(5)}
            className="px-4 py-2 bg-[var(--color-navy)] hover:bg-[var(--color-navy-light)] text-white text-xs font-bold rounded-[var(--radius-custom)] transition-all cursor-pointer shadow-sm"
          >
            Proceed to Export →
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 flex-1 overflow-hidden">
        {/* Left Pane: Slide Thumbnail List (2 Cols) */}
        <div className="lg:col-span-2 surface-card p-3 flex flex-col overflow-hidden">
          <h3 className="text-xs font-bold text-[var(--color-navy)] border-b border-[var(--color-border)] pb-2 mb-2 flex items-center space-x-1.5">
            <LayoutGrid className="w-4 h-4 text-[var(--color-amber)]" />
            <span>Slides ({slides?.length || 0})</span>
          </h3>

          <div className="mb-3 pb-2 border-b border-[var(--color-border)]">
            <label className="flex items-center space-x-2 text-[10px] text-[var(--color-navy)] font-bold cursor-pointer select-none">
              <input
                type="checkbox"
                checked={filterMissing}
                onChange={(e) => setFilterMissing(e.target.checked)}
                className="rounded border-zinc-300 text-[var(--color-amber)] focus:ring-[var(--color-amber)]"
              />
              <span>Missing files only</span>
            </label>
          </div>

          <div className="flex-1 overflow-y-auto space-y-2 pr-1">
            {slides?.map((slide, idx) => {
              // Check if this slide has any un-mapped figures/tables/charts
              const hasMissingFigures = slide.shapes.some((shape: any) => {
                if (shape.imageUrl) return false;
                const text = (shape.textBody?.paragraphs || [])
                  .map((para: any) => para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '')
                  .join(' ')
                  .trim();
                return /\b(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart)\s*[\d.]+/i.test(text) || /\binsert\s+(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart|image)(?:\s*[\d.]+)?/i.test(text);
              });

              if (filterMissing && !hasMissingFigures) {
                return null;
              }

              const active = currentSlideIndex === idx;

              return (
                <button
                  key={slide.slide_id}
                  onClick={() => setCurrentSlideIndex(idx)}
                  className={`w-full flex items-center justify-between p-2 border rounded-[var(--radius-custom)] text-left hover:bg-[var(--color-cream)] transition-all ${active
                    ? "border-[var(--color-amber)] bg-[var(--color-cream)] ring-2 ring-[var(--color-amber)]/25"
                    : "border-[var(--color-border)] bg-white"
                    }`}
                >
                  <p className="text-[11px] font-semibold text-[var(--color-navy)] truncate flex-1">
                    Slide {idx + 1}
                  </p>
                  {hasMissingFigures && (
                    <span title="Unmapped figure or table reference found on this slide">
                      <AlertTriangle className="w-3 h-3 text-red-500 flex-shrink-0 ml-1" />
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        </div>

        {/* Middle Pane: Slide Canvas (7 Cols / 10 Cols depending on sidebar toggle) */}
        <div
          ref={canvasContainerRef}
          className={`surface-card p-4 flex flex-col justify-between items-center overflow-hidden transition-all duration-300 ${showModifications ? 'lg:col-span-7' : 'lg:col-span-10'
            }`}
        >
          {/* Controls header */}
          <div className="w-full flex flex-col sm:flex-row sm:items-center sm:justify-between border-b border-[var(--color-border)] pb-2 mb-3 gap-2">
            <span className="text-xs font-bold text-[var(--color-navy)]">
              Slide {currentSlideIndex + 1} of {slides?.length || 1}
            </span>

            <div className="flex items-center space-x-2">
              <div className="flex items-center space-x-1.5 bg-[var(--color-cream)] px-2.5 py-1 rounded border border-[var(--color-border)]">
                <span className="text-[10px] font-semibold text-[var(--color-navy)]">Scale:</span>
                <span className="text-[10px] font-bold text-[var(--color-amber)]">{Math.round(scale * 100)}%</span>
              </div>

              <button
                onClick={() => setShowModifications(!showModifications)}
                className="px-3 py-1.5 bg-zinc-100 hover:bg-zinc-200 border border-zinc-300 text-zinc-800 text-[11px] font-bold rounded flex items-center space-x-1 cursor-pointer transition-all shadow-sm"
                title={showModifications ? "Hide slide audit sidebar" : "Show slide audit sidebar"}
              >
                <span>{showModifications ? 'Hide Audits' : 'Show Audits'}</span>
              </button>

              <button
                onClick={() => setIsFullscreen(true)}
                className="px-3 py-1.5 bg-[var(--color-navy)] hover:bg-[var(--color-navy-light)] text-white text-[11px] font-bold rounded flex items-center space-x-1 cursor-pointer transition-all shadow-sm"
                title="Full Screen Slideshow"
              >
                <Play className="w-3.5 h-3.5 fill-current" />
                <span>Present</span>
              </button>
            </div>
          </div>

          {/* Styled Canvas Area with overflow scroll support */}
          <div className="flex-1 w-full flex items-center justify-center bg-zinc-900 border border-[var(--color-border)] rounded-[var(--radius-custom)] p-1 overflow-auto">
            {currentSlide ? (
              renderSlideContent(scale, scaledWidth, canvasHeight)
            ) : (
              <p className="text-xs text-[var(--color-muted)]">No active slide structure loaded.</p>
            )}
          </div>
        </div>

        {/* Right Pane: Formatting Modifications Report (3 Cols) */}
        {showModifications && (
          <div className="lg:col-span-3 surface-card p-4 flex flex-col overflow-hidden border border-[var(--color-border)] shadow-sm bg-[var(--color-cream)]/10 animate-in slide-in-from-right duration-200">
            <h4 className="text-xs font-black uppercase tracking-wider text-[var(--color-navy)] flex items-center space-x-1.5 border-b border-[var(--color-border)] pb-2 mb-3">
              <LayoutGrid className="w-3.5 h-3.5 text-[var(--color-amber)]" />
              <span>Slide Modifications</span>
            </h4>

            {/* Slide Validation Checklist Panel */}
            {currentSlide && (
              <div className="bg-white border border-[var(--color-border)] rounded-[var(--radius-custom)] p-3 mb-4 shadow-xs text-left">
                <h5 className="text-[11px] font-bold text-[var(--color-navy)] uppercase tracking-wider mb-2 pb-1 border-b border-zinc-150 flex items-center justify-between">
                  <span>Slide Checklist</span>
                  <span className="text-[9px] font-medium text-[var(--color-muted)]">Validation Status</span>
                </h5>
                <div className="space-y-2">
                  {(() => {
                    const checklist = getChecklistItems(currentSlide);
                    if (checklist.length === 0) {
                      return (
                        <div className="text-[10px] text-[var(--color-muted)] italic py-1">
                          No figure or table insertions required on this slide.
                        </div>
                      );
                    }
                    return checklist.map((item, idx) => (
                      <div key={idx} className="flex items-start justify-between text-[10px] py-1 border-b border-zinc-50 last:border-0">
                        <div className="flex-1 min-w-0 pr-2">
                          <span className="font-semibold text-zinc-700 block truncate">{item.label}</span>
                          {item.detail && <span className="text-[9px] text-zinc-400 block truncate">{item.detail}</span>}
                        </div>
                        <span className={`px-2 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider ${item.done
                          ? 'bg-emerald-50 text-emerald-700 border border-emerald-250'
                          : 'bg-rose-50 text-rose-700 border border-rose-250'
                          }`}>
                          {item.done ? 'Done' : 'Not Done'}
                        </span>
                      </div>
                    ));
                  })()}
                </div>
              </div>
            )}

            {/* Missing Figures Validation Section */}
            {(() => {
              const missingForActive = currentSlide ? currentSlide.shapes.filter((shape: any) => {
                if (shape.imageUrl) return false;
                const text = (shape.textBody?.paragraphs || [])
                  .map((para: any) => para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '')
                  .join(' ')
                  .trim();
                return /\b(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart)\s*[\d.]+/i.test(text) || /\binsert\s+(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart|image)(?:\s*[\d.]+)?/i.test(text);
              }) : [];

              if (missingForActive.length === 0) return null;

              return (
                <div className="space-y-2 mb-3">
                  {missingForActive.map((shape: any) => {
                    const textSnippet = (shape.textBody?.paragraphs || [])
                      .map((para: any) => para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '')
                      .join(' ')
                      .trim();
                    return (
                      <div
                        key={shape.index}
                        className="bg-rose-50 border-l-4 border-rose-500 p-2.5 rounded shadow-xs text-left"
                      >
                        <div className="flex items-center space-x-1.5 mb-1 text-rose-800">
                          <AlertTriangle className="w-3.5 h-3.5 text-rose-500 flex-shrink-0" />
                          <span className="text-[9px] font-black uppercase tracking-wider">Unmapped Placeholder</span>
                        </div>
                        <p className="text-[10px] text-rose-700 font-medium leading-relaxed">
                          Slide contains un-replaced reference text:
                          <strong className="block mt-0.5 bg-rose-100/60 px-1.5 py-0.5 rounded font-mono text-[9px] break-words text-rose-900">
                            "{textSnippet.length > 80 ? textSnippet.substring(0, 80) + '...' : textSnippet}"
                          </strong>
                        </p>
                      </div>
                    );
                  })}
                </div>
              );
            })()}

            <div className="flex-1 overflow-y-auto space-y-3 pr-1">
              {(() => {
                const currentSlideChanges = reportData.find((item) => item.slide === currentSlideIndex + 1);
                return currentSlideChanges && currentSlideChanges.placeholders.length > 0 ? (
                  currentSlideChanges.placeholders.map((ph: any) => (
                    <div key={ph.idx} className="space-y-1.5 text-left border-b border-[var(--color-border)] last:border-0 pb-3 mb-3 last:pb-0 last:mb-0">
                      <div className="flex items-center space-x-2">
                        <span className="text-[9px] font-bold bg-[var(--color-navy)] text-white px-1.5 py-0.5 rounded uppercase tracking-wider">
                          {ph.type.replace(/ \(\d+\)/, '')}
                        </span>
                        <span className="text-[10px] text-[var(--color-muted)] font-semibold truncate block max-w-[140px]" title={ph.name}>{ph.name}</span>
                      </div>
                      {ph.paras.map((p: any, pIdx: number) => (
                        <div key={pIdx} className="pl-2 border-l-2 border-[var(--color-border)] ml-1 space-y-1">
                          <p className="text-[10px] font-medium text-zinc-500 italic">"{p.text.substring(0, 50)}..."</p>
                          <div className="flex flex-col gap-1.5 mt-1">
                            {p.changes.map((c: any, cIdx: number) => (
                              <span key={cIdx} className="text-[9px] bg-white border border-[var(--color-border)] text-zinc-700 px-2 py-0.5 rounded flex flex-wrap items-center font-semibold">
                                <span className="text-[8px] uppercase tracking-wider text-[var(--color-amber)] mr-1">{c.prop}:</span>
                                <span className="line-through text-zinc-400 mr-1">{c.before}</span>
                                <span className="text-zinc-300 mr-1">→</span>
                                <span className="text-emerald-700 font-black">{c.after}</span>
                              </span>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  ))
                ) : (
                  <div className="text-[11px] text-[var(--color-muted)] italic flex flex-col items-center justify-center h-full py-8 text-center">
                    <span>✨ Slide layout formatting matches the template master styles exactly.</span>
                  </div>
                );
              })()}
            </div>
          </div>
        )}
      </div>

      {/* Fullscreen Slideshow Overlay */}
      {isFullscreen && (
        <div className="fixed inset-0 bg-zinc-950/98 z-[100] flex flex-col justify-between p-6 select-none animate-in fade-in duration-200">
          {/* Top Header */}
          <div className="flex items-center justify-between text-white border-b border-zinc-800 pb-3">
            <div className="flex items-center space-x-3">
              <span className="text-sm font-bold bg-zinc-800 px-3 py-1 rounded">
                Slide {currentSlideIndex + 1} of {slides?.length || 1}
              </span>
              <span className="text-xs text-zinc-400 font-medium">PowerPoint Slideshow Mode (Press Esc to Exit)</span>
            </div>
            <button
              onClick={() => setIsFullscreen(false)}
              className="px-4 py-1.5 bg-red-600 hover:bg-red-700 text-white text-xs font-bold rounded cursor-pointer transition-all shadow"
            >
              Exit Slideshow
            </button>
          </div>

          {/* Main Large Slide Canvas */}
          <div className="flex-1 flex items-center justify-center p-4 overflow-auto">
            {renderSlideContent(activeScale, activeWidth, activeHeight)}
          </div>

          {/* Bottom Navigation */}
          <div className="flex items-center justify-center space-x-6 pb-2">
            <button
              onClick={() => setCurrentSlideIndex(Math.max(0, currentSlideIndex - 1))}
              disabled={currentSlideIndex === 0}
              className="px-5 py-2 bg-zinc-800 hover:bg-zinc-700 text-white disabled:opacity-40 rounded cursor-pointer font-semibold text-sm transition-all"
            >
              Previous Slide
            </button>
            <button
              onClick={() => setCurrentSlideIndex(Math.min((slides?.length || 1) - 1, currentSlideIndex + 1))}
              disabled={currentSlideIndex === (slides?.length || 1) - 1}
              className="px-5 py-2 bg-zinc-800 hover:bg-zinc-700 text-white disabled:opacity-40 rounded cursor-pointer font-semibold text-sm transition-all"
            >
              Next Slide
            </button>
          </div>
        </div>
      )}


    </div>
  );
};
