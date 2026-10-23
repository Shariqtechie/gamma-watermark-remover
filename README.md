<div align="center">

# 🧼 Gamma Watermark Remover

### Remove the "Made with Gamma" watermark from your exported PPTX in seconds.
**100% in your browser. No upload. No signup. No BS.**

[![Try it Live](https://img.shields.io/badge/🚀%20Try%20It%20Live-Free-8b5cf6?style=for-the-badge)](https://ungamma.pages.dev/)
[![Runs Locally](https://img.shields.io/badge/Runs-In%20Your%20Browser-10b981?style=for-the-badge)](https://ungamma.pages.dev/)
[![No Upload](https://img.shields.io/badge/Files-Never%20Uploaded-3b82f6?style=for-the-badge)](https://ungamma.pages.dev/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)](LICENSE)

### [👉 **Open the tool: ungamma.pages.dev**](https://ungamma.pages.dev/)

<img src="site/assets/img/og-image.png" alt="Ungamma: your slides, minus the Gamma badge" width="720">

<!-- 🎬 ADD YOUR DEMO GIF HERE. This is the single biggest star-driver. -->
<!-- ![Demo](assets/demo.gif) -->

**If this saved you time, a ⭐ helps other people find it. Thank you!**

</div>

---

## ⚡ The 10-second version

1. Export your deck from Gamma as **.pptx**
2. Drop it into the tool
3. Download the clean copy

That's it. Your file is processed **locally in your browser tab** and never touches a server.

---

## 🤔 Why does this exist?

Exporting a presentation from Gamma can leave branding images baked into your slides, layouts, and masters. Deleting them by hand in PowerPoint means opening every layout and master and hunting for the same image again and again.

This tool does that cleanup automatically.

---

## ✨ Features

| | |
|---|---|
| 🔒 **Private by design** | Runs 100% client-side. Your file never leaves your device. |
| ⚡ **Instant** | No queue, no waiting on a server. |
| 🆓 **Free forever** | No account, no paywall, no limits. |
| 🧠 **Smart detection** | Finds repeated embedded images by hash across slides, layouts, and masters. |
| 🎨 **Keeps your design** | Skips full-slide background-like images so your deck still looks right. |
| 📦 **Zero install** | Open the site and go. Or clone the repo and open `index.html`. |
| 🔓 **Open source** | Read every line of the code. Nothing hidden. |

---

## 🔐 Privacy

Presentations are often confidential, so this was built to never see your file.

- ❌ No upload server
- ❌ No account or login
- ❌ No file storage
- ❌ No tracking of your file contents
- ✅ Everything happens in your browser with JavaScript

Don't trust it? Disconnect from the internet after the page loads and it still works. Or read the source.

---

## 🛠️ How it works

A `.pptx` is just a ZIP of XML files and media assets. The tool:

1. Opens the file locally using [JSZip](https://stuk.github.io/jszip/)
2. Reads the relationship files for slides, layouts, and masters
3. Maps image references to where they're used
4. Detects repeated embedded images using file hashes
5. Filters out full-slide background-like images
6. Removes the matching unwanted image references
7. Rebuilds the `.pptx` and downloads the cleaned copy

---

## ❓ FAQ

<details>
<summary><b>How do I remove the Gamma watermark from a PowerPoint?</b></summary>

Export your deck from Gamma as PPTX, open the tool, drop the file in, and download the cleaned version. No software to install.
</details>

<details>
<summary><b>Is it safe? Do you upload my presentation?</b></summary>

No. The file is processed entirely in your browser and is never sent anywhere.
</details>

<details>
<summary><b>Does it work on Mac, Windows, Linux, and mobile?</b></summary>

It works in any modern browser (Chrome, Edge, Firefox, Safari). Very large decks on phones may be slower.
</details>

<details>
<summary><b>Will it break my slide design?</b></summary>

It's built to skip full-slide background-like images and only remove repeated embedded assets. Still, always open the cleaned file and check it before presenting. Your original file is never modified.
</details>

<details>
<summary><b>It didn't remove the watermark. Why?</b></summary>

See [Limitations](#️-limitations) below, then [open an issue](../../issues) with details. Real-world failing files are the best way to improve the tool.
</details>

---

## ⚠️ Limitations

This tool is built **only** to remove the **"Made with Gamma" watermark** from PPTX files exported from Gamma. It is not a general-purpose watermark remover, and it won't work on decks from other tools.

If the watermark is still there after cleaning, [open an issue](../../issues) and tell us what happened.

Always open the cleaned file and check it before presenting.

---

## 💻 Run locally

```bash
git clone https://github.com/Shariqtechie/gamma-watermark-remover.git
cd gamma-watermark-remover
# open index.html in your browser, that's it
```

No build step. No dependencies to install.

---

## 🧰 Tech stack

HTML · CSS · Vanilla JavaScript · [JSZip](https://stuk.github.io/jszip/) · [FileSaver.js](https://github.com/eligrey/FileSaver.js) · Cloudflare Pages

---

## 🤝 Contributing

PRs, bug reports, and test cases are very welcome.

If the tool fails on a deck, open an issue and include:

- What you expected
- What actually happened
- Whether the background changed
- Whether the unwanted image remained

> 🔒 Please don't attach private or sensitive presentations to public issues.

---

## 🗺️ Roadmap

All of these will run 100% client-side, with no backend and no uploads.

- [ ] **Badge gallery:** preview every suspected watermark image and click the ones to remove, so a logo you want to keep is never deleted
- [ ] **Batch cleaning:** drop 50 decks at once and get one `Cleaned_Presentations.zip` back
- [ ] **Metadata scrubber:** clear author, company, and edit-time info from `docProps/core.xml` and `docProps/app.xml`
- [ ] **Presenter notes remover:** one toggle to strip everything in `ppt/notesSlides/`
- [ ] **Embedded font stripper:** drop `ppt/fonts/*.fntdata` and shrink bloated files by 10 to 50 MB
- [ ] **Gamma PDF exports:** support cleaning PDFs exported from Gamma, using `pdf-lib`
- [ ] **Local PDF export:** render the cleaned deck to PDF right in the browser
- [ ] **Before/after preview** and a report of exactly what was removed
- [ ] **Dark mode / theme inverter:** swap `<a:clrScheme>` colors in `ppt/theme/theme1.xml`

Got an idea or a deck that fails? [Open an issue](../../issues) and let's talk.

---

## ⭐ Star History

<a href="https://star-history.com/#Shariqtechie/gamma-watermark-remover&Date">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/svg?repos=Shariqtechie/gamma-watermark-remover&type=Date&theme=dark" />
    <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/svg?repos=Shariqtechie/gamma-watermark-remover&type=Date" />
    <img alt="Star History Chart" src="https://api.star-history.com/svg?repos=Shariqtechie/gamma-watermark-remover&type=Date" />
  </picture>
</a>

---

## 📣 Share it

Know someone fighting a watermark on their deck? Send them the link:

**https://ungamma.pages.dev/**

---

## ⚖️ Disclaimer

This is an independent open-source project and is **not affiliated with, endorsed by, or sponsored by Gamma**. "Gamma" is a trademark of its respective owner. Use this tool only on presentations you own or have the right to modify, and respect the terms of service of any platform you use.

---

<div align="center">

### Made with ☕ by [Shariqtechie](https://github.com/Shariqtechie)

**If this helped you, drop a ⭐. It takes one second and means a lot.**

</div>
