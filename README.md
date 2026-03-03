# PPTX Opener

A lightweight, 100% client-side tool for viewing PowerPoint (`.pptx`) presentations directly in your browser. 

Unlike many online converters, **PPTX Opener processes everything locally**. Your presentation files never leave your computer, ensuring absolute privacy and security—all without the need for a backend or cloud service.

## Features

- **100% Local Processing:** Files are parsed and rendered directly via the browser using Javascript.
- **Full Presentation Support:** Accurately renders slides including text styling, shapes, standard objects, and embedded images.
- **Drag & Drop Interface:** Drop any `.pptx` file onto the page to immediately begin rendering. 
- **Presenter View:** Includes a full-screen mode, navigation controls, and a mini-thumbnail strip to jump between slides. 
- **Standalone:** No build steps required. No bloated dependencies. Just open `index.html` in your browser and you're good to go.

## How It Works

Under the hood, a `.pptx` file is just a zip archive (OOXML) containing XML configurations and media files. This tool leverages the following setup to reconstruct the presentation:

1. **JSZip** reads the raw `.pptx` file and unpacks its internal structures and embedded media (like images) on the fly. 
2. A custom-patched version of **pptx2html** parses the XML metadata to map positions, dimensions, formatting, and colors.
3. The slides are then natively generated as HTML, CSS, and inline base64 images that look nearly identical to how they appear in a traditional presentation software.

*(Note: We use a patched version of `pptx2html` to fix critical parsing issues related to missing theme colors and un-styled table structures, which normally cause the stock library to crash on real-world files.)*

## Getting Started

Because the app is entirely client-side, getting started is extremely straightforward:

1. Clone or download this repository.
2. Open `index.html` in any modern web browser (Chrome, Firefox, Edge, Safari).
3. Drag and drop a `.pptx` file into the box on the screen.

## Project Structure

- `index.html` - The main UI structure and fallback loader.
- `styles.css` - UI layout, dark-mode styling, and logic to cleanly scale up generated slides.
- `app.js` - Handles the application lifecycle, including drag/drop mechanics, extracting image streams from JSZip, calculating screen dimensions, and slide navigation.
- `pptx2html.patched.js` - The customized presentation-rendering engine.

## Troubleshooting

If a very complex slideshow fails to render structurally (due to deeply nested SmartArt or unsupported custom plugins from the presenter), the tool will seamlessly drop into a **Fallback Mode**. Fallback mode parses and returns all raw text from the presentation slide-by-slide so you can still access the core information.

## License

This project relies on open-source libraries:
- [JSZip](https://stuk.github.io/jszip/)
- [jQuery](https://jquery.com/)
- [pptx2html](https://github.com/g21589/PPTX2HTML) (Patched)

Feel free to fork, modify, or embed this viewer into your own internal apps!