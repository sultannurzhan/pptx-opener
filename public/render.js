import { state } from './state.js';
import { slideContainer, thumbnailStrip, loading, viewer, dropzone } from './dom.js';
import { updateSlideScale, goTo, showError } from './ui.js';

export function handleFile(file) {
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    showError("Please open a .pptx file.");
    return;
  }
  dropzone.classList.add("hidden");
  loading.classList.remove("hidden");
  viewer.classList.add("hidden");
  slideContainer.innerHTML = "";
  thumbnailStrip.innerHTML = "";
  state.totalSlides = 0;
  state.currentIndex = 0;

  const reader = new FileReader();
  reader.onload = (e) => startRender(e.target.result);
  reader.onerror = () => showError("Could not read the file.");
  reader.readAsArrayBuffer(file);
}

function startRender(arrayBuffer) {
  if (typeof pptx2html !== "undefined") {
    // Render pptx2html into an OFFSCREEN container that has real dimensions.
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

function onPptx2HtmlDone(offscreen, arrayBuffer) {
  const wrapper = offscreen.querySelector(".pptx-wrapper");
  if (!wrapper) {
    console.warn("pptx-wrapper not found, using fallback");
    fallbackRender(arrayBuffer);
    return;
  }

  const styleHtml = Array.from(wrapper.querySelectorAll("style"))
    .map(s => s.outerHTML).join("");

  const sections = Array.from(wrapper.querySelectorAll("section"));
  state.totalSlides = sections.length;

  if (!state.totalSlides) {
    console.warn("No sections found, using fallback");
    fallbackRender(arrayBuffer);
    return;
  }

  const DISPLAY_W = 800;
  const THUMB_W = 110;
  const THUMB_H = 70;

  sections.forEach((section, i) => {
    const natW = section.offsetWidth || 960;
    const natH = section.offsetHeight || 540;
    const scale = DISPLAY_W / natW;

    const sectionHTML = section.outerHTML;

    const wrap = document.createElement("div");
    wrap.className = "slide-wrap";
    wrap.dataset.slideIndex = i;
    wrap.dataset.natW = natW;
    wrap.dataset.natH = natH;
    wrap.style.cssText = [
      "width:100%",
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
  setTimeout(updateSlideScale, 0);
}

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
      var xml = await zip.files[slideFiles[i]].async("text");
      var num = (slideFiles[i].match(/(\d+)/) || ["", i + 1])[1];

      var wrap = document.createElement("div");
      wrap.className = "slide-wrap";
      wrap.dataset.slideIndex = i;
      wrap.dataset.natW = 800;
      wrap.dataset.natH = 450;
      wrap.style.cssText = "width:100%;overflow:hidden;display:none;background:#fff;";
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

    state.totalSlides = slideFiles.length;
    loading.classList.add("hidden");
    viewer.classList.remove("hidden");
    goTo(0);
    setTimeout(updateSlideScale, 0);
  } catch (err) {
    showError("Failed to parse the file: " + err.message);
  }
}

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
        sz: rPr ? Math.max(10, Math.min(parseInt(rPr.getAttribute("sz") || "1800") / 100, 72)) : 18,
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
