import React from 'react';
import { useStore } from './store';
import { Step1Template } from './components/Step1Template';
import { Step2Upload } from './components/Step2Upload';
import { Step3Figures } from './components/Step3Figures';
import { Step4Mapping } from './components/Step4Mapping';
import { Step5Export } from './components/Step5Export';
import { Toaster, toast } from 'sonner';
import { LayoutGrid, ArrowLeft, ArrowRight, ShieldCheck } from 'lucide-react';

const steps = [
  { number: 1, label: 'Template Master' },
  { number: 2, label: 'Upload Content' },
  { number: 3, label: 'PDF Figures' },
  { number: 4, label: 'Review Presentation' },
  { number: 5, label: 'Export Deck' },
];

function App() {
  const { step, setStep, selectedTemplate, inputPptName, sourcePdfName, convertDeck } = useStore();

  const handleNext = async () => {
    if (step === 3) {
      try {
        await convertDeck(4);
      } catch (err) {
        toast.error("Failed to compile styled presentation with figures.");
      }
    } else if (step < 5) {
      setStep(step + 1);
    }
  };

  const handleBack = () => {
    if (step > 1) setStep(step - 1);
  };

  // Determine if next button should be enabled in the footer
  const isNextDisabled = () => {
    if (step === 1 && !selectedTemplate) return true;
    if (step === 2 && (!inputPptName || !sourcePdfName)) return true;
    return false;
  };

  return (
    <div className="min-h-screen bg-[var(--color-cream)] text-[var(--color-navy)] flex flex-col font-sans">
      {/* Merged Header & Stepper Bar */}
      <header className="bg-white border-b border-[var(--color-border)] py-2.5 px-6 flex items-center justify-between sticky top-0 z-50 shadow-sm gap-4">
        {/* Logo and Brand */}
        <div className="flex items-center space-x-2 flex-shrink-0">
          <div className="w-8 h-8 bg-[var(--color-navy)] rounded-[var(--radius-custom)] flex items-center justify-center text-white">
            <LayoutGrid className="w-4.5 h-4.5" />
          </div>
          <div>
            <h1 className="text-sm font-bold tracking-tight text-[var(--color-navy)]">PPT Builder</h1>
          </div>
        </div>

        {/* Compact Horizontal Stepper */}
        <div className="hidden md:flex items-center space-x-3 max-w-2xl flex-1 justify-center">
          {steps.map((s, index) => {
            const isCompleted = step > s.number;
            const isActive = step === s.number;

            return (
              <React.Fragment key={s.number}>
                <div
                  onClick={() => {
                    if (s.number === 1) setStep(1);
                    else if (s.number === 2 && selectedTemplate) setStep(2);
                    else if (s.number === 3 && selectedTemplate && inputPptName && sourcePdfName) setStep(3);
                    else if (s.number === 4 && selectedTemplate && inputPptName && sourcePdfName) setStep(4);
                    else if (s.number === 5 && selectedTemplate && inputPptName && sourcePdfName) setStep(5);
                  }}
                  className="flex items-center space-x-1.5 cursor-pointer select-none group transition-all"
                >
                  <div
                    className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold border transition-all ${
                      isCompleted
                        ? 'bg-[var(--color-success)] border-[var(--color-success)] text-white'
                        : isActive
                        ? 'bg-[var(--color-navy)] border-[var(--color-navy)] text-white ring-2 ring-slate-900/5'
                        : 'bg-white border-[var(--color-border)] text-[var(--color-muted)] group-hover:border-[var(--color-navy)] group-hover:text-[var(--color-navy)]'
                    }`}
                  >
                    {isCompleted ? <ShieldCheck className="w-3 h-3" /> : s.number}
                  </div>
                  <span
                    className={`text-[10px] font-bold tracking-tight uppercase ${
                      isActive
                        ? 'text-[var(--color-navy)]'
                        : 'text-[var(--color-muted)] group-hover:text-[var(--color-navy)]'
                    }`}
                  >
                    {s.label}
                  </span>
                </div>
                {index < steps.length - 1 && (
                  <div className={`w-8 h-[1px] ${isCompleted ? 'bg-[var(--color-success)]' : 'bg-[var(--color-border)]'}`} />
                )}
              </React.Fragment>
            );
          })}
        </div>

        {/* Autosave Chip */}
        <div className="flex items-center space-x-1.5 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded text-emerald-800 text-[10px] font-semibold flex-shrink-0">
          <div className="w-1.2 h-1.2 rounded-full bg-emerald-500 animate-pulse"></div>
          <span>Connected</span>
        </div>
      </header>

      {/* Main Workspace */}
      <main className="flex-1 p-4 md:p-5 max-w-full px-6 md:px-8 w-full mx-auto flex flex-col justify-start pt-6 md:pt-8">
        {step === 1 && <Step1Template />}
        {step === 2 && <Step2Upload />}
        {step === 3 && <Step3Figures />}
        {step === 4 && <Step4Mapping />}
        {step === 5 && <Step5Export />}
      </main>

      {/* Footer Navigation */}
      {step <= 2 && (
        <footer className="bg-white border-t border-[var(--color-border)] py-4 px-6 sticky bottom-0 z-50 shadow-md">
          <div className="max-w-4xl mx-auto flex items-center justify-between">
            <div className="text-xs font-medium text-[var(--color-muted)]">
              Step {step} of 5
            </div>
  
            <div className="flex items-center space-x-3">
              <button
                onClick={handleBack}
                disabled={step === 1}
                className="px-4 py-2 hover:bg-[var(--color-cream)] disabled:opacity-40 text-[var(--color-navy)] border border-[var(--color-border)] rounded-[var(--radius-custom)] font-semibold text-xs transition-all flex items-center space-x-1.5 cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                <span>Back</span>
              </button>
  
              {step < 5 && (
                <button
                  onClick={handleNext}
                  disabled={isNextDisabled()}
                  className="px-4 py-2 bg-[var(--color-navy)] hover:bg-[var(--color-navy-light)] disabled:bg-neutral-300 text-white rounded-[var(--radius-custom)] font-semibold text-xs transition-all flex items-center space-x-1.5 cursor-pointer shadow-sm"
                >
                  <span>Next</span>
                  <ArrowRight className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          </div>
        </footer>
      )}

      {/* Rich Notifications Toaster */}
      <Toaster position="top-right" expand={false} richColors />
    </div>
  );
}

export default App;
