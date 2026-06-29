import React, { useRef, useState } from 'react';
import { useStore } from '../store';
import { Upload, FileText, Check, AlertCircle, Loader2 } from 'lucide-react';
import { toast } from 'sonner';

export const Step2Upload: React.FC = () => {
  const {
    inputPptName,
    sourcePdfName,
    selectedTemplate,
    isConverting,
    conversionProgress,
    uploadInputPptFile,
    uploadPdfFile,
    setStep
  } = useStore();

  const pptInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);

  const [pptDrag, setPptDrag] = useState(false);
  const [pdfDrag, setPdfDrag] = useState(false);

  const handleDrag = (e: React.DragEvent, type: 'ppt' | 'pdf') => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      if (type === 'ppt') setPptDrag(true);
      else setPdfDrag(true);
    } else if (e.type === "dragleave") {
      if (type === 'ppt') setPptDrag(false);
      else setPdfDrag(false);
    }
  };

  const handleDrop = async (e: React.DragEvent, type: 'ppt' | 'pdf') => {
    e.preventDefault();
    e.stopPropagation();
    if (type === 'ppt') setPptDrag(false);
    else setPdfDrag(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      if (type === 'ppt') {
        if (file.name.endsWith('.pptx')) {
          await uploadPpt(file);
        } else {
          toast.error("Please drop a valid .pptx file.");
        }
      } else {
        if (file.name.endsWith('.pdf')) {
          await uploadPdf(file);
        } else {
          toast.error("Please drop a valid .pdf file.");
        }
      }
    }
  };

  const uploadPpt = async (file: File) => {
    try {
      await uploadInputPptFile(file);
      toast.success("Presentation content uploaded successfully!");
    } catch (err) {
      toast.error("Failed to upload presentation.");
    }
  };

  const uploadPdf = async (file: File) => {
    try {
      await uploadPdfFile(file);
      toast.success("PDF source document uploaded successfully!");
    } catch (err) {
      toast.error("Failed to upload PDF source.");
    }
  };

  return (
    <div className="space-y-3 max-w-4xl mx-auto">
      <div className="flex items-center justify-between border-b border-[var(--color-border)] pb-2">
        <h2 className="text-xs font-black uppercase tracking-wider text-[var(--color-navy)]">
          Upload Content & Figures Source
        </h2>
      </div>

      {isConverting ? (
        <div className="surface-card p-12 flex flex-col items-center justify-center text-center space-y-6 animate-pulse">
          <Loader2 className="w-16 h-16 text-[var(--color-amber)] animate-spin" />
          <div className="space-y-2 w-full max-w-md">
            <h3 className="text-lg font-bold text-[var(--color-navy)]">Applying Layout Master Styles</h3>
            <div className="w-full bg-[var(--color-cream)] rounded-full h-2 overflow-hidden border border-[var(--color-border)]">
              <div
                className="bg-[var(--color-amber)] h-full transition-all duration-300"
                style={{ width: `${conversionProgress}%` }}
              ></div>
            </div>
            <p className="text-xs text-[var(--color-muted)]">Formatting shapes, fonts, and slides... {conversionProgress}%</p>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* PPT Upload Card */}
          <div className="surface-card p-6 flex flex-col justify-between space-y-4">
            <div>
              <h3 className="text-lg font-bold text-[var(--color-navy)] flex items-center space-x-2">
                <FileText className="w-5 h-5 text-[var(--color-amber)]" />
                <span>1. Raw Content PPTX</span>
              </h3>
              <p className="text-xs text-[var(--color-muted)] mt-1">
                The input presentation containing raw text, slide titles, and basic tables.
              </p>
            </div>

            <div
              onDragEnter={(e) => handleDrag(e, 'ppt')}
              onDragOver={(e) => handleDrag(e, 'ppt')}
              onDragLeave={(e) => handleDrag(e, 'ppt')}
              onDrop={(e) => handleDrop(e, 'ppt')}
              onClick={() => pptInputRef.current?.click()}
              className={`dashed-drop flex flex-col items-center justify-center p-6 text-center cursor-pointer min-h-[160px] ${
                pptDrag ? "drag-active border-[var(--color-amber)]" : ""
              }`}
            >
              <input
                type="file"
                ref={pptInputRef}
                onChange={(e) => e.target.files?.[0] && uploadPpt(e.target.files[0])}
                accept=".pptx"
                className="hidden"
              />
              {inputPptName ? (
                <div className="space-y-2">
                  <div className="w-10 h-10 bg-emerald-50 text-[var(--color-success)] rounded-full flex items-center justify-center mx-auto">
                    <Check className="w-5 h-5" />
                  </div>
                  <p className="text-sm font-semibold text-[var(--color-navy)] truncate max-w-[240px] mx-auto">
                    {inputPptName}
                  </p>
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      // Clear in store
                      useStore.setState({ inputPptName: null });
                    }}
                    className="text-xs text-red-500 hover:underline"
                  >
                    Remove
                  </button>
                </div>
              ) : (
                <div className="space-y-3">
                  <Upload className="w-8 h-8 text-[var(--color-muted)] mx-auto" />
                  <div>
                    <span className="text-sm font-semibold text-[var(--color-navy)]">
                      Drop content presentation
                    </span>
                    <p className="text-xs text-[var(--color-muted)] mt-1">.pptx files only</p>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* PDF Upload Card */}
          <div className="surface-card p-6 flex flex-col justify-between space-y-4">
            <div>
              <h3 className="text-lg font-bold text-[var(--color-navy)] flex items-center space-x-2">
                <FileText className="w-5 h-5 text-[var(--color-navy)]" />
                <span>2. Source PDF Document</span>
              </h3>
              <p className="text-xs text-[var(--color-muted)] mt-1">
                The document containing tables, charts, or images to extract and insert.
              </p>
            </div>

            <div
              onDragEnter={(e) => handleDrag(e, 'pdf')}
              onDragOver={(e) => handleDrag(e, 'pdf')}
              onDragLeave={(e) => handleDrag(e, 'pdf')}
              onDrop={(e) => handleDrop(e, 'pdf')}
              onClick={() => pdfInputRef.current?.click()}
              className={`dashed-drop flex flex-col items-center justify-center p-6 text-center cursor-pointer min-h-[160px] ${
                pdfDrag ? "drag-active border-[var(--color-amber)]" : ""
              }`}
            >
              <input
                type="file"
                ref={pdfInputRef}
                onChange={(e) => e.target.files?.[0] && uploadPdf(e.target.files[0])}
                accept=".pdf"
                className="hidden"
              />
              {sourcePdfName ? (
                <div className="space-y-2">
                  <div className="w-10 h-10 bg-emerald-50 text-[var(--color-success)] rounded-full flex items-center justify-center mx-auto">
                    <Check className="w-5 h-5" />
                  </div>
                  <p className="text-sm font-semibold text-[var(--color-navy)] truncate max-w-[240px] mx-auto">
                    {sourcePdfName}
                  </p>
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      useStore.setState({ sourcePdfName: null, sourcePdfPages: 0, pdfUrl: null });
                    }}
                    className="text-xs text-red-500 hover:underline"
                  >
                    Remove
                  </button>
                </div>
              ) : (
                <div className="space-y-3">
                  <Upload className="w-8 h-8 text-[var(--color-muted)] mx-auto" />
                  <div>
                    <span className="text-sm font-semibold text-[var(--color-navy)]">
                      Drop source document
                    </span>
                    <p className="text-xs text-[var(--color-muted)] mt-1">.pdf files only</p>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {!isConverting && (
        <div className="flex justify-between items-center bg-white/60 p-4 border border-[var(--color-border)] rounded-[var(--radius-custom)]">
          <div className="flex items-center space-x-2 text-xs text-[var(--color-muted)]">
            <AlertCircle className="w-4 h-4" />
            <span>
              Using template master: <strong>{selectedTemplate ? selectedTemplate.name : "None selected"}</strong>
            </span>
          </div>
          <button
            onClick={() => setStep(3)}
            disabled={!inputPptName || !sourcePdfName}
            className="px-6 py-3 bg-[var(--color-navy)] hover:bg-[var(--color-navy-light)] disabled:bg-neutral-300 text-white font-semibold rounded-[var(--radius-custom)] transition-all cursor-pointer shadow-md"
          >
            Proceed to PDF Figures
          </button>
        </div>
      )}
    </div>
  );
};
