# ppt-xtract

A simple Python script to extract all text from a PowerPoint (`.pptx`) file, preserving the visual layout order and line breaks. It exports to `.docx`, `.md`, or `.rtf`, including all slide text, speaker notes, and comments.

## Features

-   **Preserve Order of Text Flow**: Extracts text from shapes in their visual top-to-bottom order.
-   **Preserves Formatting**: Keeps all line breaks from within text boxes.
-   **Comprehensive Extraction**: Pulls text from slides, speaker notes, and all comments.
-   **Multiple Output Formats**: Exports to DOCX (default), Markdown, and RTF.

## Installation

The script requires Python 3.

### Recommended (Highest Quality Output)

This method uses Pandoc for the best conversion results.

1.  **Install Pandoc:**
    -   **macOS (with Homebrew):** `brew install pandoc`
    -   **Windows/Linux:** See the [official Pandoc installation guide](https://pandoc.org/installing.html).

2.  **Install Python Libraries:**
    ```bash
    pip install python-pptx lxml pypandoc mdutils python-docx PyRTF3
    ```

### Basic (No Pandoc Required)

This method uses pure Python libraries for conversion.

```bash
pip install python-pptx lxml mdutils python-docx PyRTF3
```

## Usage

Run the script from your terminal.

```bash
python ppt-xtract.py [options] <input_file.pptx> [output_format]
```

### Arguments

-   `input_file.pptx`: The path to your PowerPoint file.
-   `output_format`: The desired output format (`docx`, `md`, or `rtf`). **Defaults to `docx` if omitted.**
-   `--no-comments`: Excludes comments from the output.
-   `--output-lib`: Force a specific conversion library (`auto`, `pandoc`, or `native`).
-   `--wrap-text WIDTH`: For Markdown output, wrap text at `WIDTH` characters. `0` disables wrapping (default).

### Examples

**1. Convert to DOCX (Default)**
```bash
python ppt-xtract.py "My Presentation.pptx"
```
> This will create `My Presentation.docx`.

**2. Convert to Markdown without comments**
```bash
python ppt-xtract.py "My Presentation.pptx" md --no-comments
```

**3. Convert to RTF using only native Python libraries**
```bash
python ppt-xtract.py "My Presentation.pptx" rtf --output-lib native
```

**4. Convert to Markdown with text wrapped at 80 characters**
```bash
python ppt-xtract.py "My Presentation.pptx" md --wrap-text 80
```

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.
