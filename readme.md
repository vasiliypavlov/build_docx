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
*   **Multi-Project Support:** You can place `config.yaml` and `theme.yaml` directly inside your input text folder for isolated project settings.
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

## 📖 Quick Start

1. Place your Markdown files in a folder (e.g., `./text`).
2. Rename `config.example.yaml` to `config.yaml` and fill in your project details.
3. Run the compiler:

```bash
python build_docx.py -i ./text -o My_Academic_Paper.docx
```

To automatically generate an appendix with your source code (using Pygments highlighting):

```bash
python build_docx.py -i ./text -o My_Academic_Paper.docx --code ./src
```

## ⚙️ Configuration & Multi-Project Support

The project uses two YAML files for full control over the output:
*   `config.yaml` (Logic & Content: title, author, paths, feature toggles).
*   `theme.yaml` (Visual Styling: fonts, colors, margins, table padding).

### Working with Multiple Articles
If you are working on several papers, you don't need to constantly edit the global `config.yaml` and `theme.yaml`. You can place **local config files directly inside the input folder** (the one specified with `-i`).

**Priority logic for config files:**
1.  Path explicitly specified in CLI (`-c` or `-t`).
2.  Files located inside the input directory (`-i`).
3.  Global files in the root directory of the project (fallback).

**Recommended Project Structure:**
```text
build_docx/
├── build_docx.py
├── config.yaml         # Global fallback configs
├── theme.yaml          # Global fallback styles
├── article_1/
│   ├── config.yaml     # Settings specific to article_1
│   ├── theme.yaml      # Styles specific to article_1
│   ├── text.md         # Markdown content
│   └── code/           # Source code for appendix
└── article_2/
    └── text.md         # Uses global configs
```
By using this structure, you can simply run `python build_docx.py -i ./article_1` and the compiler will automatically pick up the local configurations!

## 📝 Supported Markdown Syntax

*   **Headings:** `# H1`, `## H2`, `### H3`
*   **Text Formatting:** `**Bold**`, `*Italic*`, `~~Strikethrough~~`, `` `Inline Code` ``
*   **Lists:** Standard bullet points and numbered lists (automatically styled, supports hanging indents for bibliographies).
*   **Math:** `$E=mc^2$` (inline) and `$$\int x dx$$` (block).
*   **Tables:** Standard Markdown tables (converted to smart, academic-style tables).
*   **Callouts:** `> [!NOTE]`, `> [!WARNING]` (creates colored background blocks).
*   **Footnotes:** `Text with a footnote[^1]`
*   **Code Blocks:** ` ```python ... ``` ` (rendered as code blocks; ` ```text ` is rendered as plaintext with borders).

## 🛡️ License

This project is licensed under the MIT License. Feel free to use, modify, and distribute it for your own academic and commercial publications.