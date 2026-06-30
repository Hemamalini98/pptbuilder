import React from 'react';
import { useStore } from './store';
import logo from './assets/logo.png';
import { Step1Template } from './components/Step1Template';
import { Step2Upload } from './components/Step2Upload';
import { Step3Figures } from './components/Step3Figures';
import { Step4Mapping } from './components/Step4Mapping';
import { Step5Export } from './components/Step5Export';
import { Toaster } from 'sonner';
import { ShieldCheck } from 'lucide-react';

const steps = [
  { number: 1, label: 'Template Master' },
  { number: 2, label: 'Upload Content' },
  { number: 3, label: 'PDF Figures' },
  { number: 4, label: 'Review Presentation' },
  { number: 5, label: 'Export Deck' },
];

function App() {
  const { step, setStep, selectedTemplate, inputPptName, sourcePdfName } = useStore();

  return (
    <div className="min-h-screen bg-[var(--color-cream)] text-[var(--color-navy)] flex flex-col font-sans">
      {/* Merged Header & Stepper Bar */}
      <header className="bg-white border-b border-[var(--color-border)] py-2.5 px-6 flex items-center justify-between sticky top-0 z-50 shadow-sm gap-4">
        {/* Logo and Brand */}
        <div className="flex items-center space-x-3 flex-shrink-0">
          <img src={logo} alt="S4Carlisle Logo" className="h-9 object-contain" />
          <div className="border-l border-zinc-200 pl-3">
            <h1 className="text-base font-black tracking-wider text-[var(--color-navy)] uppercase">SlideFormatter</h1>
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



      {/* Rich Notifications Toaster */}
      <Toaster position="top-right" expand={false} richColors />
    </div>
  );
}

export default App;
