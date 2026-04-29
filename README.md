# DocuFlow - Professional Document Converter

DocuFlow is a sleek, modern, and high-performance document conversion application. It allows users to seamlessly transform various document formats into pristine PDFs or editable Word files with zero formatting loss.

## 🚀 Features

- **Multi-Format Support:** Convert Word (.doc, .docx), Excel (.xls, .xlsx), Images (.jpg, .png), Text (.txt), and PDF (.pdf).
- **Dual Conversion Modes:**
    - **PDF Converter:** Generate high-fidelity PDFs from documents and images.
    - **Word Converter:** Transform PDFs and other formats back into editable Word documents.
- **Smart Image Merging:** Upload multiple images to automatically combine them into a single PDF or Word document.
- **Premium Dark UI:** A beautiful, responsive, and minimalist design with modern animations.
- **Fast & Reliable:** Uses background worker threads and MS Office automation for pixel-perfect results.

## 🛠️ Technology Stack

- **Backend:** [FastAPI](https://fastapi.tiangolo.com/) (Python 3.10+)
- **Frontend:** Vanilla HTML5, CSS3, and JavaScript
- **Office Automation:** [comtypes](https://pythonhosted.org/comtypes/) (for high-fidelity MS Office conversion)
- **Document Processing:** 
    - [pdf2docx](https://github.com/dothinking/pdf2docx) for PDF to Word
    - [python-docx](https://python-docx.readthedocs.io/) for Word generation
    - [Pillow](https://python-pillow.org/) for image processing
- **Icons:** FontAwesome 6

## 💻 Requirements

- Windows OS (Required for MS Office COM automation)
- Microsoft Word & Excel installed
- Python 3.10+

## ⚙️ Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/rkaviramachandiran/Docu-Flow.git
   ```
2. Install dependencies:
   ```bash
   pip install fastapi uvicorn comtypes Pillow python-docx pdf2docx reportlab
   ```
3. Run the application:
   ```bash
   python main.py
   ```
   *The app will be available at `http://127.0.0.1:8000`*

## 🎨 Design Philosophy

DocuFlow was designed with a focus on **Visual Excellence** and **Minimalism**. Every interaction is smoothed with micro-animations and a curated dark-slate color palette to provide a premium user experience.

---
Built with ❤️ using Antigravity AI.
