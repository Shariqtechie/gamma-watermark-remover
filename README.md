<div align="center">

# PPTX Tools

**A free browser-based PPTX cleaner for removing unwanted repeated embedded layout images from PowerPoint files.**

[![Live Demo](https://img.shields.io/badge/Live%20Demo-Open%20Tool-8b5cf6?style=for-the-badge)](https://pptx-tools.pages.dev/)
[![Runs Locally](https://img.shields.io/badge/Runs-Locally%20in%20Browser-10b981?style=for-the-badge)](https://pptx-tools.pages.dev/)
[![No Upload](https://img.shields.io/badge/Files-Not%20Uploaded-3b82f6?style=for-the-badge)](https://pptx-tools.pages.dev/)

[**Use the live tool**](https://pptx-tools.pages.dev/)

</div>

---

## What is PPTX Tools?

PPTX Tools is a small client-side utility that cleans unwanted repeated embedded images from `.pptx` files.

Some exported presentations include extra images inside slides, slide layouts, or slide masters. Removing those manually can be annoying, especially when the same asset is repeated across many layouts.

This tool scans the internal PPTX structure and removes matching repeated layout assets while trying to preserve the actual slide design and full-slide backgrounds.

---

## Privacy

Your PPTX file never leaves your device.

The tool runs fully inside your browser. There is no upload server, no account, and no file storage.

---

## Features

- Works directly in the browser
- No file upload
- Cleans `.pptx` files locally
- Detects repeated image hashes
- Scans:
  - slides
  - slide layouts
  - slide masters
- Skips full-slide background-like images
- Downloads a cleaned copy automatically
- Free and open source

---

## Live Demo

https://pptx-tools.pages.dev/

---

## How it works

A `.pptx` file is basically a zip archive containing XML files and media assets.

PPTX Tools:

1. Opens the `.pptx` file locally using JSZip.
2. Reads slide, layout, and master relationship files.
3. Maps image references to their XML traces.
4. Detects repeated embedded image assets using file hashes.
5. Filters out full-slide background-like images.
6. Removes matching unwanted image references.
7. Rebuilds the `.pptx` file and downloads the cleaned version.

---

## When it is useful

This tool can help when a presentation contains:

- repeated embedded layout images
- unwanted exported branding assets
- small repeated images inside slide layouts
- assets that are difficult to remove manually from PowerPoint

---

## Limitations

This tool may not work for every PPTX file.

It currently focuses on repeated embedded image assets. It may not remove:

- text-based marks
- shapes drawn directly in XML
- watermarks baked into full-slide background images
- locked/protected presentation content
- unusual PowerPoint export structures

Always check the cleaned file before using it.

---

## Tech Stack

- HTML
- CSS
- JavaScript
- JSZip
- FileSaver.js
- Cloudflare Pages

---

## Run locally

Clone the repository and open `index.html` in your browser.

No build step is required.

---

## Contributing

Bug reports and test cases are welcome.

If the tool fails on a specific PPTX structure, open an issue with details about:

- what was expected
- what happened
- whether the background changed
- whether the unwanted image remained

Please avoid uploading private or sensitive presentations publicly.

---

<div align="center">

Made by [Shariqtechie](https://github.com/Shariqtechie)

</div>
