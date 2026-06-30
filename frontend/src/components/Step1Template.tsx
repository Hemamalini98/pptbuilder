import React, { useEffect, useRef, useState } from 'react';
import { useStore } from '../store';
import { Upload, ChevronDown, Check, Layout, FileText, X } from 'lucide-react';
import { toast } from 'sonner';

export const Step1Template: React.FC = () => {
  const {
    savedTemplates,
    selectedTemplate,
    templateStyles,
    templateLoading,
    fetchTemplates,
    uploadTemplateFile,
    selectTemplate,
    setStep,
    customerName,
    projectName,
    setCustomerName,
    setProjectName,
    customers,
    fetchCustomers
  } = useStore();

  const [dropdownOpen, setDropdownOpen] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragActive, setIsDragActive] = useState(false);
  const [selectedLayoutIdx, setSelectedLayoutIdx] = useState<number>(0);
  const [hoveredPlaceholderIdx, setHoveredPlaceholderIdx] = useState<number | null>(null);

  useEffect(() => {
    fetchTemplates();
    fetchCustomers();
  }, [fetchTemplates, fetchCustomers]);

  useEffect(() => {
    setSelectedLayoutIdx(0);
    setHoveredPlaceholderIdx(null);
  }, [selectedTemplate]);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setIsDragActive(true);
    } else if (e.type === "dragleave") {
      setIsDragActive(false);
    }
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith('.pptx')) {
        await handleUpload(file);
      } else {
        toast.error("Please upload a valid .pptx file.");
      }
    }
  };

  const handleFileInput = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      await handleUpload(e.target.files[0]);
    }
  };

  const handleUpload = async (file: File) => {
    try {
      await uploadTemplateFile(file);
      toast.success("Template parsed and loaded successfully!");
    } catch (err) {
      toast.error("Failed to parse the template. Make sure it's a valid presentation file.");
    }
  };

  // Extract layout colors or compute statistics
  const layoutsCount = templateStyles?.slideLayouts ? templateStyles.slideLayouts.length : 0;
  
  // Calculate average shapes per layout
  let placeholderCount = 0;
  if (templateStyles?.slideLayouts) {
    templateStyles.slideLayouts.forEach((layout: any) => {
      if (layout.placeholders && Array.isArray(layout.placeholders)) {
        placeholderCount += layout.placeholders.length;
      }
    });
  }



  return (
    <div className="space-y-3 max-w-7xl mx-auto">
      <div className="flex items-center justify-between border-b border-[var(--color-border)] pb-2">
        <h2 className="text-xs font-black uppercase tracking-wider text-[var(--color-navy)]">
          Select Layout Template
        </h2>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-stretch">
        {/* Left Side: Upload & Saved Selection */}
        <div className="surface-card p-6 flex flex-col justify-start space-y-6 h-[580px]">
          {/* Customer Selection & Project Details */}
          <div className="grid grid-cols-2 gap-4 flex-shrink-0 border-b border-zinc-100 pb-4">
            <div className="space-y-1.5 text-left">
              <label className="block text-[10px] font-black uppercase tracking-wider text-[var(--color-navy)]">
                Customer <span className="text-red-500">*</span>
              </label>
              <select
                value={customerName}
                onChange={(e) => setCustomerName(e.target.value)}
                className="w-full px-3 py-2 bg-[var(--color-cream)] border border-[var(--color-border)] rounded-[var(--radius-custom)] focus:outline-none focus:ring-2 focus:ring-[var(--color-amber)] text-xs font-semibold cursor-pointer"
              >
                <option value="">Select Customer...</option>
                {customers.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>

            <div className="space-y-1.5 text-left">
              <label className="block text-[10px] font-black uppercase tracking-wider text-[var(--color-navy)]">
                Project Name <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                placeholder="Enter Project Name..."
                value={projectName}
                onChange={(e) => setProjectName(e.target.value)}
                className="w-full px-3 py-2 bg-[var(--color-cream)] border border-[var(--color-border)] rounded-[var(--radius-custom)] focus:outline-none focus:ring-2 focus:ring-[var(--color-amber)] text-xs font-semibold"
              />
            </div>
          </div>

          <div className="space-y-4 flex-shrink-0">
            <label className="block text-sm font-semibold text-[var(--color-navy)]">
              Choose from Saved Templates
            </label>
            
            <div className="relative">
              <div className="relative flex items-center">
                <button
                  type="button"
                  onClick={() => setDropdownOpen(!dropdownOpen)}
                  className="w-full flex items-center justify-between pl-4 pr-10 py-3 bg-[var(--color-cream)] border border-[var(--color-border)] rounded-[var(--radius-custom)] focus:outline-none focus:ring-2 focus:ring-[var(--color-amber)] text-sm font-medium transition-all cursor-pointer text-left"
                >
                  <span className="truncate">{selectedTemplate ? selectedTemplate.name : "Select a template..."}</span>
                  <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-[var(--color-muted)] pointer-events-none" />
                </button>
                {selectedTemplate && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      // Clear template selection in store
                      useStore.setState({ selectedTemplate: null, templateStyles: null });
                    }}
                    className="absolute right-8 top-1/2 -translate-y-1/2 text-[var(--color-muted)] hover:text-rose-500 transition-colors p-1 cursor-pointer flex items-center justify-center"
                    title="Clear selection"
                  >
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>

              {dropdownOpen && (
                <div className="absolute z-10 w-full mt-2 bg-white border border-[var(--color-border)] rounded-[var(--radius-custom)] shadow-lg max-h-60 overflow-y-auto">
                  {savedTemplates.length === 0 ? (
                    <div className="px-4 py-3 text-sm text-[var(--color-muted)] text-center">
                      No templates saved yet.
                    </div>
                  ) : (
                    savedTemplates.map((tpl) => (
                      <button
                        key={tpl.filename}
                        onClick={() => {
                          selectTemplate(tpl.filename);
                          setDropdownOpen(false);
                          toast.success(`Selected template: ${tpl.name}`);
                        }}
                        className="w-full flex items-center justify-between px-4 py-3 hover:bg-[var(--color-cream)] text-left text-sm transition-all"
                      >
                        <span className="font-medium text-[var(--color-navy)]">{tpl.name}</span>
                        {selectedTemplate?.filename === tpl.filename && (
                          <Check className="w-4 h-4 text-[var(--color-success)]" />
                        )}
                      </button>
                    ))
                  )}
                </div>
              )}
            </div>
          </div>

          {!selectedTemplate && (
            <>
              <div className="relative">
                <div className="absolute inset-0 flex items-center" aria-hidden="true">
                  <div className="w-full border-t border-[var(--color-border)]"></div>
                </div>
                <div className="relative flex justify-center text-xs uppercase">
                  <span className="bg-white px-2 text-[var(--color-muted)] font-medium">Or upload new master</span>
                </div>
              </div>

              {/* Drag & Drop Area */}
              <div
                onDragEnter={handleDrag}
                onDragOver={handleDrag}
                onDragLeave={handleDrag}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                className={`dashed-drop flex flex-col items-center justify-center p-8 text-center cursor-pointer min-h-[180px] ${
                  isDragActive ? "drag-active border-[var(--color-amber)] bg-amber-50/10" : ""
                }`}
              >
                <input
                  type="file"
                  ref={fileInputRef}
                  onChange={handleFileInput}
                  accept=".pptx"
                  className="hidden"
                />
                {templateLoading ? (
                  <div className="space-y-3">
                    <div className="w-8 h-8 border-4 border-[var(--color-amber)] border-t-transparent rounded-full animate-spin mx-auto"></div>
                    <p className="text-sm font-semibold text-[var(--color-navy)]">Parsing slides styling...</p>
                  </div>
                ) : (
                  <div className="space-y-4">
                    <div className="w-12 h-12 bg-cream rounded-full flex items-center justify-center mx-auto text-[var(--color-navy)]">
                      <Upload className="w-6 h-6" />
                    </div>
                    <div>
                      <p className="text-sm font-semibold text-[var(--color-navy)]">
                        Drag and drop your template `.pptx`
                      </p>
                      <p className="text-xs text-[var(--color-muted)] mt-1">
                        PowerPoint Presentation up to 50MB
                      </p>
                    </div>
                  </div>
                )}
              </div>
            </>
          )}

          {selectedTemplate && (
            <div className="flex-1 border border-dashed border-zinc-200 rounded-[var(--radius-custom)] bg-zinc-50/50 p-6 flex flex-col items-center justify-center text-center space-y-3">
              <Check className="w-8 h-8 text-emerald-500" />
              <p className="text-xs font-bold text-[var(--color-navy)] uppercase tracking-wider">Template Loaded</p>
              <p className="text-[11px] text-[var(--color-muted)] max-w-[240px]">
                You have successfully loaded the layout styles. Review the details in the Layout Explorer on the right, or click the button below to proceed to source upload.
              </p>
            </div>
          )}
        </div>

        {/* Right Side: Preview Card & Layout Explorer */}
        <div className="surface-card p-6 flex flex-col justify-between bg-gradient-to-br from-white to-[var(--color-cream)] overflow-hidden h-[580px]">
          <div className="space-y-4 flex-1 flex flex-col min-h-0 overflow-hidden">
            <div className="flex justify-between items-center border-b border-zinc-150 pb-2">
              <h3 className="text-sm font-black uppercase tracking-wider text-[var(--color-navy)]">Template Profile</h3>
              {selectedTemplate && (
                <span className="text-[10px] font-bold text-[var(--color-amber)] bg-amber-50 px-2 py-0.5 rounded border border-amber-200 uppercase tracking-wide truncate max-w-[150px]">
                  {selectedTemplate.name}
                </span>
              )}
            </div>
            
            {selectedTemplate && templateStyles ? (() => {
              const layouts = templateStyles.slideLayouts || [];
              const activeLayout = layouts[selectedLayoutIdx] || layouts[0] || { placeholders: [] };
              const slideWidth = templateStyles.slide_width_pt || 960;
              const slideHeight = templateStyles.slide_height_pt || 540;

              return (
                <div className="flex-1 flex flex-col space-y-4 overflow-y-auto pr-1 min-h-0">
                  {/* Summary row */}
                  <div className="grid grid-cols-2 gap-3 flex-shrink-0">
                    <div className="p-2.5 bg-white border border-[var(--color-border)] rounded-[var(--radius-custom)] flex items-center space-x-2.5 shadow-xs">
                      <Layout className="w-4 h-4 text-[var(--color-navy)] flex-shrink-0" />
                      <div>
                        <div className="text-[9px] text-[var(--color-muted)] font-semibold uppercase">Layouts</div>
                        <div className="text-sm font-black text-[var(--color-navy)]">{layoutsCount}</div>
                      </div>
                    </div>
                    <div className="p-2.5 bg-white border border-[var(--color-border)] rounded-[var(--radius-custom)] flex items-center space-x-2.5 shadow-xs">
                      <FileText className="w-4 h-4 text-[var(--color-amber)] flex-shrink-0" />
                      <div>
                        <div className="text-[9px] text-[var(--color-muted)] font-semibold uppercase">Placeholders</div>
                        <div className="text-sm font-black text-[var(--color-navy)]">{placeholderCount}</div>
                      </div>
                    </div>
                  </div>

                  {/* Layout Selection Buttons */}
                  <div className="flex-shrink-0 space-y-1">
                    <div className="text-[9px] font-black uppercase tracking-wider text-[var(--color-navy)] mb-1">
                      Choose Master Layout to Inspect
                    </div>
                    <div className="flex flex-wrap gap-1.5 max-h-[85px] overflow-y-auto pb-1 pr-1">
                      {layouts.map((lay: any, idx: number) => {
                        const isActive = selectedLayoutIdx === idx;
                        return (
                          <button
                            key={idx}
                            type="button"
                            onClick={() => {
                              setSelectedLayoutIdx(idx);
                              setHoveredPlaceholderIdx(null);
                            }}
                            className={`px-2.5 py-1 text-[9px] font-bold rounded border transition-all cursor-pointer ${
                              isActive
                                ? 'bg-[var(--color-navy)] text-white border-[var(--color-navy)] shadow-sm'
                                : 'bg-white text-[var(--color-navy)] border-[var(--color-border)] hover:bg-[var(--color-cream)]'
                            }`}
                          >
                            {lay.layoutType || lay.layoutName}
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  {/* Visual micro-preview canvas */}
                  <div className="flex-shrink-0">
                    <div className="text-[9px] font-black uppercase tracking-wider text-[var(--color-navy)] mb-1">
                      Visual Micro-Canvas Layout
                    </div>
                    <div 
                      className="w-full bg-zinc-900 border border-zinc-800 rounded-[var(--radius-custom)] overflow-hidden relative shadow-inner flex-shrink-0 select-none aspect-[16/9]"
                    >
                      {activeLayout.placeholders?.map((ph: any, phIdx: number) => {
                        const leftPct = (ph.position.x_pt / slideWidth) * 100;
                        const topPct = (ph.position.y_pt / slideHeight) * 100;
                        const widthPct = (ph.size.width_pt / slideWidth) * 100;
                        const heightPct = (ph.size.height_pt / slideHeight) * 100;
                        const isHovered = hoveredPlaceholderIdx === phIdx;

                        return (
                          <div
                            key={phIdx}
                            onMouseEnter={() => setHoveredPlaceholderIdx(phIdx)}
                            onMouseLeave={() => setHoveredPlaceholderIdx(null)}
                            className={`absolute border rounded flex items-center justify-center p-0.5 text-[7px] text-center transition-all ${
                              isHovered 
                                ? 'border-[var(--color-amber)] bg-[var(--color-amber)]/25 z-10 scale-[1.02] shadow-[0_0_8px_#f5822a] font-bold text-amber-100' 
                                : 'border-white/20 bg-white/5 text-white/50'
                            }`}
                            style={{
                              left: `${leftPct}%`,
                              top: `${topPct}%`,
                              width: `${widthPct}%`,
                              height: `${heightPct}%`,
                            }}
                          >
                            <span className="truncate max-w-full leading-none scale-[0.85]">
                              {ph.placeholder?.type || 'SHAPE'}
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Placeholder details list */}
                  <div className="space-y-1.5 pr-1 flex-shrink-0">
                    <div className="text-[9px] font-black uppercase tracking-wider text-[var(--color-navy)] mb-1">
                      Placeholder Elements Detail
                    </div>
                    <div className="space-y-1.5">
                      {activeLayout.placeholders && activeLayout.placeholders.length > 0 ? (
                        activeLayout.placeholders.map((ph: any, phIdx: number) => {
                          const isHovered = hoveredPlaceholderIdx === phIdx;
                          return (
                            <div
                              key={phIdx}
                              onMouseEnter={() => setHoveredPlaceholderIdx(phIdx)}
                              onMouseLeave={() => setHoveredPlaceholderIdx(null)}
                              className={`text-[10px] p-2 border rounded flex justify-between items-center transition-all ${
                                isHovered ? 'bg-amber-50/50 border-[var(--color-amber)] shadow-xs' : 'bg-white border-zinc-100'
                              }`}
                            >
                              <div className="text-left">
                                <span className="font-bold text-[var(--color-navy)] block uppercase tracking-wide text-[9px]">{ph.placeholder?.type || 'SHAPE'}</span>
                                <span className="text-[8px] text-[var(--color-muted)] block truncate max-w-[160px]">ID/Name: {ph.shapeName || 'Unnamed'}</span>
                              </div>
                              <div className="text-right text-[8px] text-zinc-400 font-mono">
                                <div>X:{Math.round(ph.position.x_pt)}pt | Y:{Math.round(ph.position.y_pt)}pt</div>
                                <div>W:{Math.round(ph.size.width_pt)}pt | H:{Math.round(ph.size.height_pt)}pt</div>
                              </div>
                            </div>
                          );
                        })
                      ) : (
                        <div className="text-[10px] text-zinc-400 italic py-2 text-left">
                          No placeholders defined in this master layout.
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              );
            })() : (
              <div className="flex flex-col items-center justify-center py-12 text-center space-y-3 flex-1">
                <Layout className="w-12 h-12 text-neutral-300" />
                <p className="text-sm text-[var(--color-muted)]">
                  Select or upload a template to see layout properties.
                </p>
              </div>
            )}
          </div>

          <button
            onClick={() => setStep(2)}
            disabled={!selectedTemplate || !customerName || !projectName.trim()}
            className="w-full mt-4 py-3 px-4 bg-[var(--color-navy)] hover:bg-[var(--color-navy-light)] disabled:bg-neutral-200 disabled:text-neutral-400 text-white font-bold rounded-[var(--radius-custom)] transition-all cursor-pointer flex items-center justify-center space-x-2 text-xs flex-shrink-0"
          >
            <span>Proceed to Upload Source</span>
          </button>
        </div>
      </div>
    </div>
  );
};
