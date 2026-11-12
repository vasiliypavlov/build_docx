# build_docx: Academic Markdown to Word Compiler

**build_docx** is a highly customizable Python tool designed to compile Markdown (`.md`) files into publication-ready Microsoft Word (`.docx`) documents. 

Unlike standard converters, **build_docx** is tailored specifically for academic and scientific publishing. It handles complex formatting, native Word equations, academic table styling, and strict typographic rules automatically.

## ✨ Key Features

*   **Native Word Formulas (OMML):** Converts LaTeX math (`$inline$` and `$$block$$`) directly into native Microsoft Word equations, editable in the Word formula editor.
*   **Academic Table Styling (Clean Grid / Booktabs):** Automatically calculates optimal column widths, prevents text wrapping inside cells, aligns numbers to the right, and applies professional borders.
*   **Native Word Footnotes:** Bypasses `python-docx` limitations by injecting native `footnotes.xml` into the document package. Supports both endnotes and true footnotes.
*   **Live Table of Contents:** Generates a native Word TOC field that automatically updates page numbers (press `F9` in Word to refresh).
*   **Syntax Highlighting:** Uses `Pygments` to add beautiful, code-block highlighting in the Appendix, while respecting `plaintext` for terminal outputs.
*   **Smart Typography:** Automatically handles non-breaking spaces, prevents "widows and orphans" in headings, and controls hyphenation.
*   **YAML-Driven:** All logic and styling are separated into two configuration files (`config.yaml` and `theme.yaml`), making it easy to change fonts, margins, and rules without touching the Python code.
*   **Embedded Metadata:** Injects Title, Author, Keywords, and License directly into the DOCX file properties for digital libraries (arXiv, Zenodo, etc.).
*   **Bilingual Support:** Perfectly handles mixed Russian/English texts (with automatic language detection for hyphenation and typography).

## 🚀 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/vasiliypavlov/build_docx.git
   cd build_docx
   ```

2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. (Optional) For rendering Mermaid diagrams, install Node.js and Mermaid CLI:
   ```bash
   npm install -g @mermaid-js/mermaid-cli
   ```