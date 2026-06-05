# Quick Start Guide

### Step 1: Install Dependencies (One Time Setup)

```bash
pip install -r requirements.txt
```

### Step 2: Run the Application

```bash
streamlit run app.py
```

### Step 3: Upload Your Tender Files

1. Open the web page (usually http://localhost:8501)
2. In the sidebar, click "Browse files"
3. Select ALL your tender documents (PDF, DOCX, TXT)
4. Click "🔍 Process Documents"
5. Wait for processing (2-5 minutes for large files)

### Step 4: Review Extracted Content

- Click on keywords in the left sidebar to see extracted content
- Check location map and climate information
- Review nearby infrastructure

### Step 5: Generate PPT

1. Click "📊 Generate PowerPoint Presentation"
2. Wait for generation (creates 60-75 slides)
3. Click "⬇️ Download Presentation"
4. Open the PPT file - it's ready for your director!

## Key Features You Need to Know

✅ **Semantic Detection**: System understands "Scope of Work" = "Nature of Work"  
✅ **OCR Support**: Works with scanned PDFs and images  
✅ **No Hardcoding**: All content comes from YOUR documents  
✅ **60-75 Slides**: Professional presentation ready for director  
✅ **Short Summaries**: Easy to read bullet points  

## Important Notes

- **First Time**: May need to configure OCR paths (see README.md)
- **Large Files**: 200-300 page documents take 5-10 minutes to process
- **Multiple Files**: Upload all tender files at once - they'll be processed together
- **Dynamic**: Each tender may have different headings - system adapts automatically

## Troubleshooting

**If OCR doesn't work:**
- Check if Tesseract is installed
- Update path in `backend/ocr_utils.py`

**If location not found:**
- Check document for location text
- System will still work, just won't show map

**If PPT generation is slow:**
- Normal for large documents
- Be patient - it's processing everything

---

**You're all set! The system will extract everything from your tender documents and create a professional presentation automatically.**

