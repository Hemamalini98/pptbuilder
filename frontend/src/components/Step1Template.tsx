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
    setStep
  } = useStore();

  const [dropdownOpen, setDropdownOpen] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragActive, setIsDragActive] = useState(false);

  useEffect(() => {
    fetchTemplates();
  }, [fetchTemplates]);

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
        <div className="surface-card p-6 flex flex-col justify-between space-y-6">
          <div className="space-y-4">
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
        </div>

        {/* Right Side: Preview Card */}
        <div className="surface-card p-6 flex flex-col justify-between bg-gradient-to-br from-white to-[var(--color-cream)]">
          <div className="space-y-5">
            <h3 className="text-lg font-bold text-[var(--color-navy)]">Template Profile</h3>
            
            {selectedTemplate && templateStyles ? (
              <div className="space-y-5">
                <div className="p-4 bg-white/60 border border-[var(--color-border)] rounded-[var(--radius-custom)] space-y-3">
                  <div className="flex items-center space-x-3 text-sm">
                    <FileText className="w-5 h-5 text-[var(--color-amber)]" />
                    <span className="font-semibold text-[var(--color-navy)] truncate">
                      {selectedTemplate.name}.pptx
                    </span>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="p-3 bg-white border border-[var(--color-border)] rounded-[var(--radius-custom)] flex items-center space-x-3">
                    <Layout className="w-5 h-5 text-[var(--color-navy)]" />
                    <div>
                      <div className="text-xs text-[var(--color-muted)]">Layouts</div>
                      <div className="text-lg font-bold text-[var(--color-navy)]">{layoutsCount}</div>
                    </div>
                  </div>

                  <div className="p-3 bg-white border border-[var(--color-border)] rounded-[var(--radius-custom)] flex items-center space-x-3">
                    <Layout className="w-5 h-5 text-[var(--color-amber)]" />
                    <div>
                      <div className="text-xs text-[var(--color-muted)]">Placeholders</div>
                      <div className="text-lg font-bold text-[var(--color-navy)]">{placeholderCount}</div>
                    </div>
                  </div>
                </div>


              </div>
            ) : (
              <div className="flex flex-col items-center justify-center py-12 text-center space-y-3">
                <Layout className="w-12 h-12 text-neutral-300" />
                <p className="text-sm text-[var(--color-muted)]">
                  Select or upload a template to see layout properties.
                </p>
              </div>
            )}
          </div>

          <button
            onClick={() => setStep(2)}
            disabled={!selectedTemplate}
            className="w-full mt-6 py-3 px-4 bg-[var(--color-navy)] hover:bg-[var(--color-navy-light)] disabled:bg-neutral-300 text-white font-semibold rounded-[var(--radius-custom)] transition-all cursor-pointer flex items-center justify-center space-x-2"
          >
            <span>Proceed to Upload Source</span>
          </button>
        </div>
      </div>
    </div>
  );
};
