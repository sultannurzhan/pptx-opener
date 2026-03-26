/* ─────────────────────────────────────────────────────────────────
  Nexuslides — app.js
   Libraries:
     • pptx2html  (global: pptx2html)  renders slides into DOM
     • JSZip                           fallback text extraction
─────────────────────────────────────────────────────────────────── */

(function () {
  "use strict";

  /* ── JSZip v2 → v3 compatibility patch ─────────────────────────
     pptx2html@0.3.4 uses the old synchronous JSZip v2 API:
       zip.file(name).asArrayBuffer()   (removed in JSZip v3)
     We patch JSZip.loadAsync to pre-fetch every entry, attach
     the old v2 methods, AND add a fuzzy file matcher in case
     pptx2html's path concatenation doesn't match the zip keys.
  ─────────────────────────────────────────────────────────────── */
  (function patchJSZipV2Compat() {
    if (typeof JSZip === "undefined") return;
    var _origLoad = JSZip.loadAsync.bind(JSZip);
    JSZip.loadAsync = function (data, opts) {
      return _origLoad(data, opts).then(function (zip) {
        var names = Object.keys(zip.files).filter(function (n) {
          return !zip.files[n].dir;
        });

        // Pre-fetch every file so we can expose sync-like methods
        return Promise.all(names.map(function (name) {
          return zip.files[name].async("arraybuffer").then(function (ab) {
            var entry = zip.files[name];
            entry.asArrayBuffer = function () {
              // Return a fresh slice to avoid contamination
              return ab.slice(0);
            };
            entry.asUint8Array = function () { return new Uint8Array(ab); };
            entry.asBinary = function () {
              var s = "";
              var u8 = new Uint8Array(ab);
              for (var i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
              return s;
            };
            entry.asText = function () {
              return new TextDecoder("utf-8").decode(new Uint8Array(ab));
            };
          });
        })).then(function () {
          // Override zip.file(name) to be fuzzy
          var _origFile = zip.file.bind(zip);
          zip.file = function(name) {
            if (!name) return null;
            // 1. Exact match
            var exact = _origFile(name);
            if (exact) return exact;

            // 2. Fuzzy match (ignore case and leading slashes)
            var cleanName = name.toLowerCase().replace(/\\/g, '/');
            try { cleanName = decodeURIComponent(cleanName); } catch(e) {}
            if (cleanName.startsWith('/')) cleanName = cleanName.substring(1);
            if (cleanName.startsWith('ppt/')) cleanName = cleanName.substring(4); // Remove ppt/ to match just media/...
            if (cleanName.startsWith('../')) cleanName = cleanName.substring(3);

            var keys = Object.keys(zip.files);
            for (var i = 0; i < keys.length; i++) {
              var lowKey = keys[i].toLowerCase();
              try { lowKey = decodeURIComponent(lowKey); } catch(e) {}
              if (lowKey.endsWith(cleanName)) {
                return zip.files[keys[i]];
              }
            }

            // 3. Fallback dummy to prevent PPTX engine crash
            console.warn("Fuzzy match failed for image/resource:", name);
            return {
              asArrayBuffer: function() { return new ArrayBuffer(0); },
              asUint8Array: function() { return new Uint8Array(0); },
              asText: function() { return ""; },
              asBinary: function() { return ""; }
            };
          };

          return zip;
        });
      });
    };
  })();

  /* ── DOM refs ─────────────────────────────────────────────── */
  const dropzone       = document.getElementById("dropzone");
  const fileInput      = document.getElementById("fileInput");
  const loading        = document.getElementById("loading");
  const viewer         = document.getElementById("viewer");
  const slideContainer = document.getElementById("slideContainer");
  const thumbnailStrip = document.getElementById("thumbnailStrip");
  const slideCounter   = document.getElementById("slideCounter");
  const btnPrev        = document.getElementById("btnPrev");
  const btnNext        = document.getElementById("btnNext");
  const btnClose       = document.getElementById("btnClose");
  const btnFullscreen  = document.getElementById("btnFullscreen");

  /* ── State ─────────────────────────────────────────────────── */
  let totalSlides  = 0;
  let currentIndex = 0;
  let isFullscreen = false;

  /* ── Drag-and-drop setup ───────────────────────────────────── */
  // Guard: the <label for="fileInput"> already opens the picker natively.
  // Without this check the dropzone click also fires, opening it twice.
  dropzone.addEventListener("click", (e) => {
    if (e.target.closest("label")) return;
    fileInput.click();
  });
  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("over");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("over"));
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("over");
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
  });

  /* ── Toolbar ───────────────────────────────────────────────── */
  btnPrev.addEventListener("click", () => goTo(currentIndex - 1));
  btnNext.addEventListener("click", () => goTo(currentIndex + 1));
  btnClose.addEventListener("click", reset);
  btnFullscreen.addEventListener("click", toggleFullscreen);

  document.addEventListener("keydown", (e) => {
    if (viewer.classList.contains("hidden")) return;
    if (e.key === "ArrowRight" || e.key === "ArrowDown") goTo(currentIndex + 1);
    if (e.key === "ArrowLeft"  || e.key === "ArrowUp")   goTo(currentIndex - 1);
    if (e.key === "Escape" && isFullscreen) toggleFullscreen();
    if (e.key === "f" || e.key === "F") toggleFullscreen();
  });

  /* ── File handling ─────────────────────────────────────────── */
  function handleFile(file) {
    if (!file.name.toLowerCase().endsWith(".pptx")) {
      showError("Please open a .pptx file.");
      return;
    }
    dropzone.classList.add("hidden");
    loading.classList.remove("hidden");
    viewer.classList.add("hidden");
    slideContainer.innerHTML = "";
    thumbnailStrip.innerHTML = "";
    totalSlides = 0;
    currentIndex = 0;

    const reader = new FileReader();
    reader.onload = (e) => startRender(e.target.result);
    reader.onerror = () => showError("Could not read the file.");
    reader.readAsArrayBuffer(file);
  }

  /* ── Attempt pptx2html, fall back to JSZip ─────────────────── */
  function startRender(arrayBuffer) {
    if (typeof pptx2html !== "undefined") {
      // Render pptx2html into an OFFSCREEN container that has real dimensions.
      // This is critical: pptx2html's internal resize() reads offsetWidth to
      // compute its scale transform. If the element is display:none (scale=0)
      // everything is invisible — including images.
      const offscreen = document.createElement("div");
      offscreen.style.cssText = [
        "position:fixed",
        "left:-99999px",
        "top:0",
        "width:960px",
        "visibility:hidden",
        "pointer-events:none",
      ].join(";");
      document.body.appendChild(offscreen);

      let resolved = false;
      const timeout = setTimeout(() => {
        if (!resolved) {
          console.warn("pptx2html timed out, using fallback");
          if (document.body.contains(offscreen)) document.body.removeChild(offscreen);
          fallbackRender(arrayBuffer);
        }
      }, 25000);

      pptx2html(arrayBuffer, offscreen, null)
        .then(() => {
          resolved = true;
          clearTimeout(timeout);
          onPptx2HtmlDone(offscreen, arrayBuffer);
          if (document.body.contains(offscreen)) document.body.removeChild(offscreen);
        })
        .catch((err) => {
          resolved = true;
          clearTimeout(timeout);
          console.warn("pptx2html error, using fallback:", err);
          if (document.body.contains(offscreen)) document.body.removeChild(offscreen);
          fallbackRender(arrayBuffer);
        });
    } else {
      fallbackRender(arrayBuffer);
    }
  }

  /* ── After pptx2html renders all slides ────────────────────── */
  function onPptx2HtmlDone(offscreen, arrayBuffer) {
    // pptx2html puts a single div.pptx-wrapper inside our container.
    // That wrapper holds <style> elements and one <section> per slide.
    const wrapper = offscreen.querySelector(".pptx-wrapper");
    if (!wrapper) {
      console.warn("pptx-wrapper not found, using fallback");
      fallbackRender(arrayBuffer);
      return;
    }

    // Collect the CSS pptx2html injected (needed so section content renders)
    const styleHtml = Array.from(wrapper.querySelectorAll("style"))
      .map(s => s.outerHTML).join("");

    const sections = Array.from(wrapper.querySelectorAll("section"));
    totalSlides = sections.length;

    if (!totalSlides) {
      console.warn("No sections found, using fallback");
      fallbackRender(arrayBuffer);
      return;
    }

    const DISPLAY_W = 800;
    const THUMB_W   = 110;
    const THUMB_H   = 70;

    sections.forEach((section, i) => {
      // Read real dimensions from the offscreen layout
      const natW = section.offsetWidth  || 960;
      const natH = section.offsetHeight || 540;
      const scale = DISPLAY_W / natW;

      const sectionHTML = section.outerHTML;

      // ── Main slide wrapper ────────────────────────────────────
      const wrap = document.createElement("div");
      wrap.dataset.slideIndex = i;
      wrap.style.cssText = [
        "width:" + DISPLAY_W + "px",
        "height:" + Math.round(natH * scale) + "px",
        "overflow:hidden",
        "display:none",
        "background:#fff",
      ].join(";");

      const inner = document.createElement("div");
      inner.style.cssText = [
        "width:" + natW + "px",
        "transform:scale(" + scale + ")",
        "transform-origin:top left",
      ].join(";");
      inner.innerHTML = styleHtml + sectionHTML;
      wrap.appendChild(inner);
      slideContainer.appendChild(wrap);

      // ── Thumbnail ─────────────────────────────────────────────
      const tScale = THUMB_W / natW;

      const outer = document.createElement("div");
      outer.className = "thumb";
      outer.style.cssText = [
        "width:" + THUMB_W + "px",
        "height:" + THUMB_H + "px",
        "overflow:hidden",
        "position:relative",
        "flex-shrink:0",
        "cursor:pointer",
      ].join(";");

      const tInner = document.createElement("div");
      tInner.style.cssText = [
        "width:" + natW + "px",
        "transform:scale(" + tScale + ")",
        "transform-origin:top left",
        "pointer-events:none",
      ].join(";");
      tInner.innerHTML = styleHtml + sectionHTML;
      outer.appendChild(tInner);

      const label = document.createElement("div");
      label.className = "thumb-label";
      label.textContent = i + 1;
      outer.appendChild(label);

      outer.addEventListener("click", () => goTo(i));
      thumbnailStrip.appendChild(outer);
    });

    loading.classList.add("hidden");
    viewer.classList.remove("hidden");
    goTo(0);
  }

  /* ── Navigate to a specific slide index ────────────────────── */
  function goTo(index) {
    if (!totalSlides) return;
    index = Math.max(0, Math.min(totalSlides - 1, index));
    currentIndex = index;

    Array.from(slideContainer.children).forEach((wrap, i) => {
      wrap.style.display = i === index ? "block" : "none";
    });

    slideCounter.textContent = (index + 1) + " / " + totalSlides;
    btnPrev.disabled = index === 0;
    btnNext.disabled = index === totalSlides - 1;

    Array.from(thumbnailStrip.children).forEach((t, i) => {
      t.classList.toggle("active", i === index);
    });

    const activeThumb = thumbnailStrip.children[index];
    if (activeThumb) {
      activeThumb.scrollIntoView({ behavior: "smooth", inline: "nearest", block: "nearest" });
    }
  }

  /* ── Fullscreen ────────────────────────────────────────────── */
  function toggleFullscreen() {
    isFullscreen = !isFullscreen;
    viewer.classList.toggle("fullscreen-mode", isFullscreen);
    btnFullscreen.innerHTML = isFullscreen ? "&#9633;" : "&#9974;";
    btnFullscreen.title = isFullscreen ? "Exit Fullscreen (F / Esc)" : "Fullscreen (F)";
  }

  /* ── Reset to initial state ────────────────────────────────── */
  function reset() {
    totalSlides = 0;
    currentIndex = 0;
    isFullscreen = false;
    viewer.classList.add("hidden");
    viewer.classList.remove("fullscreen-mode");
    dropzone.classList.remove("hidden");
    loading.classList.add("hidden");
    slideContainer.innerHTML = "";
    thumbnailStrip.innerHTML = "";
    fileInput.value = "";
  }

  /* ── Fallback: JSZip + raw XML → plain HTML ─────────────────── */
  async function fallbackRender(arrayBuffer) {
    try {
      const zip = await JSZip.loadAsync(arrayBuffer);
      const slideFiles = Object.keys(zip.files)
        .filter(function (n) { return /^ppt\/slides\/slide[0-9]+\.xml$/.test(n); })
        .sort(function (a, b) {
          return parseInt(a.match(/(\d+)/)[1]) - parseInt(b.match(/(\d+)/)[1]);
        });

      if (!slideFiles.length) {
        showError("Could not read any slides from this file.");
        return;
      }

      for (var i = 0; i < slideFiles.length; i++) {
        var xml  = await zip.files[slideFiles[i]].async("text");
        var num  = (slideFiles[i].match(/(\d+)/) || ["", i + 1])[1];

        var wrap = document.createElement("div");
        wrap.dataset.slideIndex = i;
        wrap.style.cssText = "width:800px;overflow:hidden;display:none;";
        wrap.appendChild(buildSlideNode(xml, num, 800));
        slideContainer.appendChild(wrap);

        var outer = document.createElement("div");
        outer.className = "thumb";
        outer.style.cssText = "width:110px;height:70px;overflow:hidden;position:relative;flex-shrink:0;cursor:pointer;";
        var tContent = buildSlideNode(xml, num, 110);
        tContent.style.pointerEvents = "none";
        outer.appendChild(tContent);
        var lbl = document.createElement("div");
        lbl.className = "thumb-label";
        lbl.textContent = i + 1;
        outer.appendChild(lbl);
        (function (idx) {
          outer.addEventListener("click", function () { goTo(idx); });
        })(i);
        thumbnailStrip.appendChild(outer);
      }

      totalSlides = slideFiles.length;
      loading.classList.add("hidden");
      viewer.classList.remove("hidden");
      goTo(0);
    } catch (err) {
      showError("Failed to parse the file: " + err.message);
    }
  }

  /* ── Build a rendered slide node from PPTX XML ──────────────── */
  function buildSlideNode(xmlString, slideNum, targetWidth) {
    var parser = new DOMParser();
    var doc = parser.parseFromString(xmlString, "application/xml");

    function byLocalName(node, localName) {
      return Array.prototype.slice.call(node.getElementsByTagNameNS("*", localName));
    }

    function firstByLocalName(node, localName) {
      var all = node.getElementsByTagNameNS("*", localName);
      return all.length ? all[0] : null;
    }

    var paragraphs = [];
    byLocalName(doc, "p").forEach(function (p) {
      var runs = [];
      byLocalName(p, "r").forEach(function (r) {
        var t = firstByLocalName(r, "t");
        if (!t) return;
        var rPr = firstByLocalName(r, "rPr");
        runs.push({
          text: t.textContent,
          bold: rPr && (rPr.getAttribute("b") === "1" || rPr.getAttribute("b") === "true"),
          sz:   rPr ? Math.max(10, Math.min(parseInt(rPr.getAttribute("sz") || "1800") / 100, 72)) : 18,
        });
      });
      if (runs.some(function (r) { return r.text.trim(); })) paragraphs.push(runs);
    });

    var NATURAL_W = 800;
    var NATURAL_H = 450;
    var scale = targetWidth / NATURAL_W;

    var slide = document.createElement("div");
    slide.style.cssText = [
      "width:" + NATURAL_W + "px",
      "min-height:" + NATURAL_H + "px",
      "background:#fff",
      "padding:48px 64px 60px",
      "font-family:Calibri,Arial,sans-serif",
      "color:#222",
      "position:relative",
      "line-height:1.4",
      "box-sizing:border-box",
      "transform:scale(" + scale + ")",
      "transform-origin:top left",
    ].join(";");

    var badge = document.createElement("div");
    badge.style.cssText = "position:absolute;bottom:12px;right:18px;font-size:11px;color:#aaa;";
    badge.textContent = "Slide " + slideNum;
    slide.appendChild(badge);

    if (!paragraphs.length) {
      var empty = document.createElement("p");
      empty.style.cssText = "color:#aaa;font-style:italic;";
      empty.textContent = "No text content on this slide.";
      slide.appendChild(empty);
      return slide;
    }

    paragraphs.forEach(function (runs) {
      var maxSz = Math.max.apply(null, runs.map(function (r) { return r.sz; }));
      var p = document.createElement("p");
      p.style.margin = "0 0 " + (maxSz > 24 ? 14 : 6) + "px";

      runs.forEach(function (r) {
        if (!r.text) return;
        var span = document.createElement("span");
        span.textContent = r.text;
        span.style.fontSize = r.sz + "px";
        if (r.bold) span.style.fontWeight = "bold";
        p.appendChild(span);
      });
      slide.appendChild(p);
    });

    return slide;
  }

  /* ── Error helper ──────────────────────────────────────────── */
  function showError(msg) {
    loading.classList.add("hidden");
    dropzone.classList.remove("hidden");

    var prev = document.getElementById("pptx-error");
    if (prev) prev.remove();

    var err = document.createElement("div");
    err.id = "pptx-error";
    err.style.cssText = "color:#ff4466;margin-top:1rem;font-size:0.9rem;text-align:center;";
    err.textContent = msg;
    dropzone.insertAdjacentElement("afterend", err);
    setTimeout(function () { err.remove(); }, 6000);
  }

})();
