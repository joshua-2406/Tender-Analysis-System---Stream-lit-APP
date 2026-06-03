AI & NLP Details

## Overview
This project uses NLP-based techniques for intelligent 
document processing without relying on paid LLM APIs, 
making it lightweight, cost-free, and production-ready.

---

## 🧠 NLP Techniques Used

### 1. Semantic Heading Detection
- **What it does:** Identifies headings in tender documents 
  even when they use different words
- **How it works:** Uses semantic similarity mapping to match 
  variations like:
  - "Scope of Work" = "Nature of Work" = "Work Description"
  - "Technical Specifications" = "Technical Requirements"
  - "Payment Terms" = "Commercial Terms" = "Financial Terms"
- **Library:** `sentence-transformers` / keyword similarity mapping
- **Why no LLM:** Rule-based semantic mapping is faster, 
  free, and accurate enough for structured tender documents

---

### 2. Keyword & Location Extraction
- **What it does:** Extracts project location automatically 
  from document text
- **How it works:** NLP-based named entity recognition (NER) 
  to identify city, state, district names
- **Library:** `spaCy` / regex-based NLP patterns
- **Output:** Feeds into geocoding API to plot on map

---

### 3. Content Summarization
- **What it does:** Converts long paragraphs into short 
  bullet points (max 100 chars)
- **How it works:** Extractive summarization — picks the 
  most relevant sentences from each section
- **Library:** Rule-based extraction (no LLM needed)
- **Output:** Used directly in PowerPoint slides

---

### 4. OCR Text Recognition
- **What it does:** Extracts text from scanned PDFs and images
- **How it works:** Converts PDF pages to images, then runs 
  OCR to extract text
- **Library:** `Tesseract OCR` + `pdf2image`
- **Accuracy:** Industry-standard OCR engine used by Google

Full AI/NLP Tech Stack

| Component               | Technology              | Purpose                        |
|-------------------------|-------------------------|--------------------------------|
| Semantic Matching       | Sentence Transformers   | Heading variation detection    |
| Named Entity Recognition| spaCy / regex           | Location extraction            |
| OCR                     | Tesseract OCR           | Scanned document processing    |
| Text Extraction         | pdfplumber              | PDF text extraction            |
| Extractive Summarization| Rule-based NLP          | Bullet point generation        |
| Geocoding               | Nominatim API           | Location to coordinates        |
| Infrastructure Detection| Overpass API            | Nearby hotels, airports        |
| Climate Data            | Open-Meteo API          | Weather information            |
