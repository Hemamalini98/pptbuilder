import { create } from 'zustand';

export interface ShapeStyle {
  font_name?: string | null;
  font_size_pt?: number | null;
  bold?: boolean | null;
  italic?: boolean | null;
  color_rgb?: string | null;
  alignment?: string | null;
}

export interface ShapeData {
  index: number;
  shapeName: string;
  shapeType: string;
  text?: string;
  position: {
    x_pt: number;
    y_pt: number;
  };
  size: {
    width_pt: number;
    height_pt: number;
  };
  placeholder?: {
    type: string;
    idx: number;
  };
  style?: ShapeStyle;
  imageUrl?: string; // If an image has been placed/exists
  textBody?: {
    bodyProperties?: Record<string, any>;
    paragraphs?: any[];
  };
  fill?: string;
  border?: {
    color?: string;
  };
}

export interface SlideData {
  index: number;
  slide_id: number;
  backgroundColor?: string;
  shapes: ShapeData[];
}

export interface StylesData {
  slide_width_pt: number;
  slide_height_pt: number;
  slideLayouts: any[];
  slides: SlideData[];
  theme?: any;
}

export interface Figure {
  id: string;
  name: string; // e.g. Figure2.3
  url: string; // url to render
  page: number; // 1-indexed
  filename: string; // unique identifier in backend uploads
  caption?: string; // extracted figure caption
  credit?: string;  // extracted figure credit
  mappedTo?: {
    slideIndex: number;
    shapeIndex: number;
  } | null;
}

export interface PdfCaption {
  id: string;
  page: number; // 1-indexed
  label: string; // e.g. "Figure 1.1" or "Table 1"
  text: string;  // full caption text
  credit?: string; // credit text
}

interface StoredTemplate {
  name: string;
  filename: string;
}

interface DeckforgeState {
  // Navigation
  step: number;
  setStep: (step: number) => void;

  // Step 1: Template Upload
  savedTemplates: StoredTemplate[];
  selectedTemplate: StoredTemplate | null;
  templateStyles: StylesData | null;
  templateLoading: boolean;
  customerName: string;
  projectName: string;
  setCustomerName: (name: string) => void;
  setProjectName: (name: string) => void;
  customers: string[];
  fetchCustomers: () => Promise<void>;
  
  // Step 2: Source Uploads
  inputPptName: string | null;
  sourcePdfName: string | null;
  sourcePdfPages: number;
  isConverting: boolean;
  conversionProgress: number;

  // Step 3: PDF Figure Extraction
  pdfUrl: string | null;
  currentPdfPage: number; // 0-indexed
  figures: Figure[];
  pdfCaptions: PdfCaption[];
  addFigure: (figure: Omit<Figure, 'id' | 'name'>) => void;
  renameFigure: (id: string, newName: string) => void;
  updateFigureCaption: (id: string, caption: string, credit?: string) => void;
  deleteFigure: (id: string) => void;

  // Step 4: Review & Mapping
  slides: SlideData[] | null;
  currentSlideIndex: number;
  setCurrentSlideIndex: (idx: number) => void;
  focusedShapeIndex: number | null;
  setFocusedShapeIndex: (idx: number | null) => void;
  placeFigureOnShape: (slideIndex: number, shapeIndex: number, figureId: string) => Promise<void>;
  placeFigureAtCoordinates: (slideIndex: number, figureId: string, x_pt: number, y_pt: number, w_pt: number, h_pt: number) => Promise<void>;
  removeFigureFromShape: (slideIndex: number, shapeIndex: number) => Promise<void>;

  // API Methods
  fetchTemplates: () => Promise<void>;
  uploadTemplateFile: (file: File) => Promise<void>;
  selectTemplate: (filename: string) => Promise<void>;
  uploadInputPptFile: (file: File) => Promise<void>;
  uploadPdfFile: (file: File) => Promise<void>;
  convertDeck: (targetStep?: number) => Promise<void>;
  resetSession: () => Promise<void>;
}

export const useStore = create<DeckforgeState>((set, get) => ({
  step: 1,
  setStep: (step) => set({ step }),

  // Step 1
  savedTemplates: [],
  selectedTemplate: null,
  templateStyles: null,
  templateLoading: false,
  customerName: '',
  projectName: '',
  setCustomerName: (customerName) => set({ customerName }),
  setProjectName: (projectName) => set({ projectName }),
  customers: [],

  // Step 2
  inputPptName: null,
  sourcePdfName: null,
  sourcePdfPages: 0,
  isConverting: false,
  conversionProgress: 0,

  // Step 3
  pdfUrl: null,
  currentPdfPage: 0,
  figures: [],
  pdfCaptions: [],

  // Step 4
  slides: null,
  currentSlideIndex: 0,
  focusedShapeIndex: null,

  fetchTemplates: async () => {
    try {
      const res = await fetch('/api/templates');
      const data = await res.json();
      if (data.ok) {
        set({ savedTemplates: data.templates });
      }
    } catch (err) {
      console.error('Failed to fetch templates:', err);
    }
  },

  fetchCustomers: async () => {
    try {
      const res = await fetch('/api/customers');
      const data = await res.json();
      if (data.ok) {
        set({ customers: data.customers });
      }
    } catch (err) {
      console.error('Failed to fetch customers:', err);
    }
  },

  uploadTemplateFile: async (file: File) => {
    set({ templateLoading: true });
    try {
      const formData = new FormData();
      formData.append('file', file);
      const res = await fetch('/api/upload-template', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json();
      if (data.ok) {
        set({
          selectedTemplate: { name: file.name.replace('.pptx', ''), filename: data.filename },
          templateStyles: data.styles,
        });
        await get().fetchTemplates();
      } else {
        throw new Error(data.detail || 'Upload failed');
      }
    } catch (err) {
      console.error('Failed to upload template:', err);
      throw err;
    } finally {
      set({ templateLoading: false });
    }
  },

  selectTemplate: async (filename: string) => {
    set({ templateLoading: true });
    try {
      const res = await fetch('/api/select-template', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename }),
      });
      const data = await res.json();
      if (data.ok) {
        set({
          selectedTemplate: { name: filename.replace('.pptx', ''), filename },
          templateStyles: data.styles,
        });
      }
    } catch (err) {
      console.error('Failed to select template:', err);
    } finally {
      set({ templateLoading: false });
    }
  },

  uploadInputPptFile: async (file: File) => {
    try {
      const formData = new FormData();
      formData.append('file', file);
      const res = await fetch('/api/upload-ppt', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json();
      if (data.ok) {
        set({ 
          inputPptName: file.name,
          slides: data.slidesInfo?.slides || null
        });
      }
    } catch (err) {
      console.error('Failed to upload input PPT:', err);
      throw err;
    }
  },

  uploadPdfFile: async (file: File) => {
    try {
      const formData = new FormData();
      formData.append('file', file);
      const res = await fetch('/api/upload-pdf', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json();
      if (data.ok) {
        set({
          sourcePdfName: file.name,
          sourcePdfPages: data.pageCount,
          pdfUrl: `/api/pdf/file`,
          currentPdfPage: 0,
          pdfCaptions: data.captions || [],
        });
      }
    } catch (err) {
      console.error('Failed to upload PDF:', err);
      throw err;
    }
  },

  convertDeck: async (targetStep = 3) => {
    set({ isConverting: true, conversionProgress: 10 });
    const interval = setInterval(() => {
      set((state) => ({
        conversionProgress: Math.min(state.conversionProgress + 15, 90),
      }));
    }, 300);

    try {
      const figuresPayload = get().figures.map((f) => ({
        name: f.name,
        filename: f.filename,
        caption: f.caption,
        credit: f.credit,
      }));

      const res = await fetch('/api/process-ppt', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ figures: figuresPayload }),
      });
      const data = await res.json();
      if (data.ok) {
        clearInterval(interval);

        // Sync figure mapping state for figures auto-inserted by Python
        const updatedFigures = get().figures.map((fig) => {
          const mapping = (data.autoInserted || []).find((m: any) => m.filename === fig.filename);
          if (mapping) {
            return {
              ...fig,
              mappedTo: {
                slideIndex: mapping.slideIndex,
                shapeIndex: mapping.shapeIndex,
              }
            };
          }
          return fig;
        });

        set({
          conversionProgress: 100,
          slides: data.slidesInfo.slides,
          figures: updatedFigures,
        });
        setTimeout(() => {
          set({ isConverting: false, step: targetStep });
        }, 500);
      } else {
        throw new Error(data.detail || 'Conversion failed');
      }
    } catch (err) {
      clearInterval(interval);
      set({ isConverting: false, conversionProgress: 0 });
      console.error('Failed to convert deck:', err);
      throw err;
    }
  },

  addFigure: (figData) => {
    set((state) => {
      // Calculate figure count for this specific page to assign an index
      const pageFigures = state.figures.filter((f) => f.page === figData.page);
      const index = pageFigures.length + 1;
      const id = `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
      const name = `Figure${figData.page}.${index}`;

      return {
        figures: [...state.figures, { ...figData, id, name, mappedTo: null }],
      };
    });
  },

  renameFigure: (id, newName) => {
    set((state) => ({
      figures: state.figures.map((f) => (f.id === id ? { ...f, name: newName } : f)),
    }));
  },

  updateFigureCaption: (id, caption, credit) => {
    set((state) => ({
      figures: state.figures.map((f) => (f.id === id ? { ...f, caption, credit } : f)),
    }));
  },

  deleteFigure: (id) => {
    set((state) => ({
      figures: state.figures.filter((f) => f.id !== id),
    }));
  },

  placeFigureOnShape: async (slideIndex, shapeIndex, figureId) => {
    const figure = get().figures.find((f) => f.id === figureId);
    const slides = get().slides;
    if (!figure || !slides) return;

    const targetSlide = slides[slideIndex];
    const targetShape = targetSlide.shapes[shapeIndex];

    try {
      const formData = new FormData();
      formData.append('slide_index', String(slideIndex));
      formData.append('image_name', figure.filename);
      formData.append('x_pt', String(targetShape.position.x_pt));
      formData.append('y_pt', String(targetShape.position.y_pt));
      formData.append('w_pt', String(targetShape.size.width_pt));
      formData.append('h_pt', String(targetShape.size.height_pt));
      if (figure.caption) {
        formData.append('caption', figure.caption);
      }

      const res = await fetch('/api/add-image', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json();
      if (data.ok) {
        set((state) => ({
          slides: data.slidesInfo.slides,
          figures: state.figures.map((f) =>
            f.id === figureId
              ? { ...f, mappedTo: { slideIndex, shapeIndex } }
              : f.mappedTo?.slideIndex === slideIndex && f.mappedTo?.shapeIndex === shapeIndex
              ? { ...f, mappedTo: null }
              : f
          ),
        }));
      }
    } catch (err) {
      console.error('Failed to map figure to shape:', err);
    }
  },

  placeFigureAtCoordinates: async (slideIndex, figureId, x_pt, y_pt, w_pt, h_pt) => {
    const figure = get().figures.find((f) => f.id === figureId);
    if (!figure) return;

    try {
      const formData = new FormData();
      formData.append('slide_index', String(slideIndex));
      formData.append('image_name', figure.filename);
      formData.append('x_pt', String(x_pt));
      formData.append('y_pt', String(y_pt));
      formData.append('w_pt', String(w_pt));
      formData.append('h_pt', String(h_pt));
      if (figure.caption) {
        formData.append('caption', figure.caption);
      }

      const res = await fetch('/api/add-image', {
        method: 'POST',
        body: formData,
      });
      const data = await res.json();
      if (data.ok) {
        set((state) => ({
          slides: data.slidesInfo.slides,
          figures: state.figures.map((f) =>
            f.id === figureId
              ? { ...f, mappedTo: { slideIndex, shapeIndex: data.slidesInfo.slides[slideIndex].shapes.length - 1 } }
              : f
          ),
        }));
      }
    } catch (err) {
      console.error('Failed to place figure at coordinates:', err);
    }
  },

  removeFigureFromShape: async (slideIndex, shapeIndex) => {
    // Currently, there's no explicit endpoint to delete a shape image,
    // but we can update our local mapping so the frontend removes it from view
    // and resets the figure status.
    set((state) => ({
      figures: state.figures.map((f) =>
        f.mappedTo?.slideIndex === slideIndex && f.mappedTo?.shapeIndex === shapeIndex
          ? { ...f, mappedTo: null }
          : f
      ),
      slides: state.slides
        ? state.slides.map((slide, sIdx) =>
            sIdx === slideIndex
              ? {
                  ...slide,
                  shapes: slide.shapes.map((sh, shIdx) =>
                    shIdx === shapeIndex ? { ...sh, imageUrl: undefined } : sh
                  ),
                }
              : slide
          )
        : null,
    }));
  },

  setCurrentSlideIndex: (currentSlideIndex) => set({ currentSlideIndex }),

  setFocusedShapeIndex: (focusedShapeIndex) => set({ focusedShapeIndex }),

  resetSession: async () => {
    // Clear uploaded files on the server first
    try {
      await fetch('/api/reset', { method: 'POST' });
    } catch (_) {}
    // Then reset all local frontend state
    set({
      step: 1,
      selectedTemplate: null,
      templateStyles: null,
      inputPptName: null,
      sourcePdfName: null,
      sourcePdfPages: 0,
      isConverting: false,
      conversionProgress: 0,
      pdfUrl: null,
      currentPdfPage: 0,
      figures: [],
      pdfCaptions: [],
      slides: null,
      currentSlideIndex: 0,
      focusedShapeIndex: null,
      customerName: '',
      projectName: '',
    });
  },
}));
