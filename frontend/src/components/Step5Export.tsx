import React, { useState, useEffect } from 'react';
import { useStore } from '../store';
import { Download, RefreshCw, AlertTriangle, FileJson, Check, LayoutGrid, BarChart2, ImageOff, Activity, ChevronDown, ChevronUp } from 'lucide-react';
import { toast } from 'sonner';

export const Step5Export: React.FC = () => {
  const {
    figures,
    slides,
    resetSession
  } = useStore();

  const [reportTab, setReportTab] = useState<'style' | 'figures' | 'accessibility'>('style');
  const [expandedSubCategories, setExpandedSubCategories] = useState<Record<string, boolean>>({});
  const [selectedSeverity, setSelectedSeverity] = useState<string>('Error');
  const [figureDiag, setFigureDiag] = useState<{ missing: string[]; unplaced: string[] } | null>(null);
  const [accessibilityDiag, setAccessibilityDiag] = useState<{ issues: string[] } | null>(null);

  useEffect(() => {
    fetch('/api/figure-diagnostics')
      .then(r => r.json())
      .then(d => { if (d.ok) setFigureDiag({ missing: d.missing, unplaced: d.unplaced }); })
      .catch(() => {});
      
    fetch('/api/accessibility-report')
      .then(r => r.json())
      .then(d => { if (d.ok) setAccessibilityDiag({ issues: d.issues }); })
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
        const isFigRef = /\b(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart)\s*[\d.]+/i.test(text) || /\binsert\s+(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart|image)(?:\s*[\d.]+)?/i.test(text);
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
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <button
              onClick={handleExportPpt}
              className="py-4 px-6 bg-[var(--color-navy)] hover:bg-[var(--color-navy-light)] text-white font-bold rounded-[var(--radius-custom)] transition-all cursor-pointer shadow-md flex items-center justify-center space-x-2.5 active:scale-[0.99]"
            >
              <Download className="w-5 h-5" />
              <span>Download Styled Presentation</span>
            </button>

            <button
              onClick={handleDownloadMappingJson}
              className="py-4 px-6 bg-white hover:bg-[var(--color-cream)] border border-[var(--color-border)] text-[var(--color-navy)] font-bold rounded-[var(--radius-custom)] transition-all cursor-pointer shadow-sm flex items-center justify-center space-x-2.5 active:scale-[0.99]"
            >
              <FileJson className="w-5 h-5 text-[var(--color-amber)]" />
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
              const missingRefs: string[] = [];
              slide.shapes.forEach((s) => {
                const isPicturePh = ((s.shapeType && (s.shapeType === '13' || s.shapeType.includes('Picture') || s.shapeType.includes('Image'))) ||
                  (s.placeholder?.type && (s.placeholder.type.includes('PICTURE') || s.placeholder.type.includes('BITMAP'))));
                  
                if (isPicturePh && !s.imageUrl) {
                  missingRefs.push("Picture Box");
                } else if (!s.imageUrl) {
                  const text = (s.textBody?.paragraphs || [])
                    .map((para: any) => para.runs ? para.runs.map((r: any) => r.sampleText || '').join('') : '')
                    .join(' ')
                    .trim();
                  
                  const match = text.match(/\b(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart)\s*[\d.]+/i) || text.match(/\binsert\s+(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart|image)(?:\s*[\d.]+)?/i);
                  if (match) {
                    missingRefs.push(match[0].toUpperCase());
                  } else {
                    const isFigRef = /\b(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart)\s*[\d.]+/i.test(text) || /\binsert\s+(figure|fig\.?|f\.?|table|tab\.?|t\.?|chart|image)(?:\s*[\d.]+)?/i.test(text);
                    if (isFigRef) {
                      missingRefs.push(text.length > 25 ? text.substring(0, 25) + '...' : text);
                    }
                  }
                }
              });

              return (
                <div
                  key={slide.slide_id}
                  className="flex items-center justify-between p-2 border border-[var(--color-border)] rounded-[var(--radius-custom)] bg-white hover:border-zinc-300 transition-all gap-2"
                >
                  <span className="text-xs font-semibold text-[var(--color-navy)] flex-shrink-0">
                    Slide {idx + 1}
                  </span>
                  {missingRefs.length > 0 ? (
                    <div className="flex flex-col items-end gap-1 flex-1 min-w-0">
                      {missingRefs.map((ref, rIdx) => (
                        <span key={rIdx} className="text-[8px] bg-rose-50 border border-rose-200 text-rose-700 px-1.5 py-0.5 rounded font-bold uppercase tracking-wide truncate max-w-full text-right" title={`Missing ${ref}`}>
                          ⚠️ Missing {ref}
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
        </div>

        {/* Style Report Tab */}
        {reportTab === 'style' && (
          <div className="w-full h-[620px]">
            <iframe
              src="/api/report"
              title="Style Difference Report"
              className="w-full h-full border-none"
            />
          </div>
        )}

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
                                <ul className="list-disc pl-4 space-y-1">
                                  {items.map((item: any, mIdx: number) => (
                                    <li key={mIdx} className="text-zinc-650 text-[11px]">
                                      {item.slide ? <span className="font-bold mr-1">Slide {item.slide}:</span> : null}
                                      {item.detail}
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
