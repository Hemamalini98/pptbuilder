const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

const app = express();
const PORT = 8001;

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const UPLOAD_DIR = path.resolve(__dirname, 'uploads');
const TEMPLATES_DIR = path.join(UPLOAD_DIR, 'templates');
if (!fs.existsSync(UPLOAD_DIR)) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
}
if (!fs.existsSync(TEMPLATES_DIR)) {
  fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
}

// In-memory state tracking
const state = {
  template_pptx: null,
  template_style_json: null,
  content_pptx: null,
  styled_pptx: null,
  pdf_path: null
};

// Set up multer for file storage
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
  }
});
const upload = multer({ storage });

// Helper to execute python commands safely
function runPythonCommand(cmd) {
  return new Promise((resolve, reject) => {
    exec(`python3 -c "${cmd.replace(/"/g, '\\"')}"`, (error, stdout, stderr) => {
      if (error) {
        console.error('Python Error:', stderr || error.message);
        reject(stderr || error.message);
      } else {
        resolve(stdout.trim());
      }
    });
  });
}

// 1. Template Upload
app.post('/api/upload-template', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ detail: 'No file uploaded' });
    const originalName = req.file.originalname;
    const targetPath = path.join(TEMPLATES_DIR, originalName);
    fs.copyFileSync(req.file.path, targetPath);
    state.template_pptx = targetPath;

    // Run main.py style extractor inline
    const pythonCode = `import json; from main import extract_template; print(json.dumps(extract_template('${targetPath}')))`;
    const output = await runPythonCommand(pythonCode);
    const styles = JSON.parse(output);

    const styleJsonName = path.basename(originalName, path.extname(originalName)) + '_styles.json';
    const styleJsonPath = path.join(TEMPLATES_DIR, styleJsonName);
    fs.writeFileSync(styleJsonPath, JSON.stringify(styles, null, 2));
    state.template_style_json = styleJsonPath;

    res.json({ ok: true, styles, filename: originalName });
  } catch (err) {
    res.status(500).json({ detail: err.toString() });
  }
});

// GET Templates list
app.get('/api/templates', (req, res) => {
  try {
    const templates = [];
    if (fs.existsSync(TEMPLATES_DIR)) {
      const files = fs.readdirSync(TEMPLATES_DIR);
      files.forEach(f => {
        if (f.endsWith('.pptx')) {
          templates.push({
            name: path.basename(f, '.pptx'),
            filename: f
          });
        }
      });
    }
    res.json({ ok: true, templates });
  } catch (err) {
    res.status(500).json({ detail: err.toString() });
  }
});

// POST Select Template
app.post('/api/select-template', (req, res) => {
  const { filename } = req.body;
  if (!filename) return res.status(400).json({ detail: 'Missing filename' });

  const targetPath = path.join(TEMPLATES_DIR, filename);
  const styleJsonName = path.basename(filename, '.pptx') + '_styles.json';
  const styleJsonPath = path.join(TEMPLATES_DIR, styleJsonName);

  if (!fs.existsSync(targetPath) || !fs.existsSync(styleJsonPath)) {
    return res.status(404).json({ detail: 'Template or styles not found' });
  }

  state.template_pptx = targetPath;
  state.template_style_json = styleJsonPath;

  const styles = JSON.parse(fs.readFileSync(styleJsonPath, 'utf8'));
  res.json({ ok: true, filename, styles });
});


// 2. Content PPT Upload
app.post('/api/upload-ppt', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ detail: 'No file uploaded' });
    const targetPath = path.join(UPLOAD_DIR, 'uploaded_content.pptx');
    fs.copyFileSync(req.file.path, targetPath);
    state.content_pptx = targetPath;

    res.json({ ok: true, filename: req.file.originalname });
  } catch (err) {
    res.status(500).json({ detail: err.toString() });
  }
});

// 3. PDF Upload
app.post('/api/upload-pdf', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ detail: 'No file uploaded' });
    const targetPath = path.join(UPLOAD_DIR, 'uploaded_document.pdf');
    fs.copyFileSync(req.file.path, targetPath);
    state.pdf_path = targetPath;

    // Get page count using PyMuPDF (fitz)
    const pythonCode = `import fitz; doc = fitz.open('${targetPath}'); print(len(doc))`;
    const pageCount = parseInt(await runPythonCommand(pythonCode), 10);

    res.json({ ok: true, filename: req.file.originalname, pageCount });
  } catch (err) {
    res.status(500).json({ detail: err.toString() });
  }
});

// 4. Process/Style PPT
app.post('/api/process-ppt', async (req, res) => {
  if (!state.content_pptx || !state.template_style_json) {
    return res.status(400).json({ detail: 'Missing template style configuration or input content PPTX' });
  }
  try {
    const outputPath = path.join(UPLOAD_DIR, 'styled_output.pptx');

    // Execute convert.py styling logic
    const convertCode = `from convert import convert; convert('${state.content_pptx}', '${state.template_style_json}', '${outputPath}', apply_geometry=True)`;
    await runPythonCommand(convertCode);
    state.styled_pptx = outputPath;

    // Extract slides information and pictures
    const extractCode = `import json; from main import extract_template; print(json.dumps(extract_template('${outputPath}')))`;
    const slidesOutput = await runPythonCommand(extractCode);
    const slidesInfo = JSON.parse(slidesOutput);

    // Scan for and extract images inside the PPTX slide shapes so the canvas can render them
    const scanImagesCode = `import os; from pptx import Presentation; prs = Presentation('${outputPath}'); res = {};
for s_idx, slide in enumerate(prs.slides):
  for sh_idx, shape in enumerate(slide.shapes):
    if hasattr(shape, 'shape_type') and int(shape.shape_type) == 13:
      try:
        img = shape.image
        ext = img.ext
        fname = f'slide_{s_idx}_shape_{sh_idx}.{ext}'
        with open(os.path.join('${UPLOAD_DIR}', fname), 'wb') as f:
          f.write(img.blob)
        res[f'{s_idx}_{sh_idx}'] = f'/api/media/{fname}'
      except:
        pass
import json; print(json.dumps(res))`;

    const imagesOutput = await runPythonCommand(scanImagesCode);
    const extractedImages = JSON.parse(imagesOutput);

    // Map extracted images to slide shape URLs
    slidesInfo.slides.forEach((slide, sIdx) => {
      slide.shapes.forEach((shape, shIdx) => {
        const key = `${sIdx}_${shIdx}`;
        if (extractedImages[key]) {
          shape.imageUrl = extractedImages[key];
        }
      });
    });

    res.json({ ok: true, slidesInfo });
  } catch (err) {
    res.status(500).json({ detail: err.toString() });
  }
});

// 5. PDF info
app.get('/api/pdf/info', async (req, res) => {
  if (!state.pdf_path) return res.status(404).json({ detail: 'No PDF uploaded' });
  try {
    const pythonCode = `import fitz; doc = fitz.open('${state.pdf_path}'); print(len(doc))`;
    const count = parseInt(await runPythonCommand(pythonCode), 10);
    res.json({ filename: path.basename(state.pdf_path), count });
  } catch (err) {
    res.status(500).json({ detail: err.toString() });
  }
});

// 6. Serve PDF file
app.get('/api/pdf/file', (req, res) => {
  if (!state.pdf_path) return res.status(404).json({ detail: 'No PDF uploaded' });
  res.sendFile(state.pdf_path);
});

// 7. Crop extraction
app.post('/api/extract', async (req, res) => {
  if (!state.pdf_path) return res.status(404).json({ detail: 'No PDF uploaded' });
  const { page, scale, x0, y0, x1, y1 } = req.body;
  try {
    const idx = fs.readdirSync(UPLOAD_DIR).filter(f => f.startsWith('crop_')).length + 1;
    const filename = `crop_p${parseInt(page, 10) + 1}_${String(idx).padStart(3, '0')}.png`;
    const filepath = path.join(UPLOAD_DIR, filename);

    // Crop using PyMuPDF (fitz)
    const pythonCode = `import fitz; doc = fitz.open('${state.pdf_path}'); p = doc[${page}];
clip = fitz.Rect(min(${x0}, ${x1})/${scale}, min(${y0}, ${y1})/${scale}, max(${x0}, ${x1})/${scale}, max(${y0}, ${y1})/${scale});
pix = p.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), clip=clip, alpha=False);
pix.save('${filepath}')`;

    await runPythonCommand(pythonCode);
    res.json({ filename, url: `/api/media/${filename}`, page: parseInt(page, 10) + 1 });
  } catch (err) {
    res.status(500).json({ detail: err.toString() });
  }
});

// 8. Place dropped image
app.post('/api/add-image', async (req, res) => {
  const targetPptx = state.styled_pptx || state.content_pptx;
  if (!targetPptx) return res.status(400).json({ detail: 'No presentation available to modify' });

  const { slide_index, image_name, x_pt, y_pt, w_pt, h_pt } = req.body;
  const imagePath = path.join(UPLOAD_DIR, image_name);
  if (!fs.existsSync(imagePath)) return res.status(404).json({ detail: `Image ${image_name} not found` });

  try {
    const pythonCode = `import json; from pptx import Presentation; from pptx.util import Pt; from main import extract_template;
prs = Presentation('${targetPptx}');
slide = prs.slides[${slide_index}];
slide.shapes.add_picture('${imagePath}', Pt(${x_pt}), Pt(${y_pt}), Pt(${w_pt}), Pt(${h_pt}));
prs.save('${targetPptx}');
print(json.dumps(extract_template('${targetPptx}')))`;

    const slidesOutput = await runPythonCommand(pythonCode);
    const slidesInfo = JSON.parse(slidesOutput);

    // Re-scan and expose image paths
    const scanImagesCode = `import os; from pptx import Presentation; prs = Presentation('${targetPptx}'); res = {};
for s_idx, slide in enumerate(prs.slides):
  for sh_idx, shape in enumerate(slide.shapes):
    if hasattr(shape, 'shape_type') and int(shape.shape_type) == 13:
      try:
        img = shape.image
        ext = img.ext
        fname = f'slide_{s_idx}_shape_{sh_idx}.{ext}'
        with open(os.path.join('${UPLOAD_DIR}', fname), 'wb') as f:
          f.write(img.blob)
        res[f'{s_idx}_{sh_idx}'] = f'/api/media/{fname}'
      except:
        pass
import json; print(json.dumps(res))`;

    const imagesOutput = await runPythonCommand(scanImagesCode);
    const extractedImages = JSON.parse(imagesOutput);

    // Map extracted images to slide shape URLs
    slidesInfo.slides.forEach((slide, sIdx) => {
      slide.shapes.forEach((shape, shIdx) => {
        const key = `${sIdx}_${shIdx}`;
        if (extractedImages[key]) {
          shape.imageUrl = extractedImages[key];
        }
      });
    });

    res.json({ ok: true, slidesInfo });
  } catch (err) {
    res.status(500).json({ detail: err.toString() });
  }
});

// 9. Download output
app.get('/api/download', (req, res) => {
  const targetPptx = state.styled_pptx || state.content_pptx;
  if (!targetPptx || !fs.existsSync(targetPptx)) {
    return res.status(404).json({ detail: 'No output PPTX available to download' });
  }
  res.download(targetPptx, 'styled_presentation.pptx');
});

// Media static server
app.use('/api/media', express.static(UPLOAD_DIR));

// React production assets if available
const frontendDist = path.resolve(__dirname, 'frontend', 'dist');
if (fs.existsSync(frontendDist)) {
  app.use(express.static(frontendDist));
  app.use((req, res, next) => {
    if (!req.path.startsWith('/api')) {
      res.sendFile(path.join(frontendDist, 'index.html'));
    } else {
      next();
    }
  });
}

app.listen(PORT, () => {
  console.log(`Express PPT Builder API running on http://localhost:${PORT}`);
});
