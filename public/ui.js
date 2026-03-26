import {
  slideContainer, thumbnailStrip, slideCounter,
  btnPrev, btnNext, viewer, btnFullscreen, dropzone, loading, fileInput
} from './dom.js';
import { state } from './state.js';

export function updateSlideScale() {
  const cw = slideContainer.getBoundingClientRect().width;
  if (!cw) return;

  const wraps = slideContainer.querySelectorAll('.slide-wrap');
  wraps.forEach(wrap => {
    const natW = parseFloat(wrap.dataset.natW);
    const natH = parseFloat(wrap.dataset.natH);
    if (!natW) return;
    const scale = cw / natW;
    wrap.style.height = (natH * scale) + "px";
    const inner = wrap.firstElementChild;
    if (inner) {
      inner.style.transform = "scale(" + scale + ")";
    }
  });
}

export function goTo(index) {
  if (!state.totalSlides) return;
  index = Math.max(0, Math.min(state.totalSlides - 1, index));
  state.currentIndex = index;

  Array.from(slideContainer.children).forEach((wrap, i) => {
    wrap.style.display = i === index ? "block" : "none";
  });

  slideCounter.textContent = (index + 1) + " / " + state.totalSlides;
  btnPrev.disabled = index === 0;
  btnNext.disabled = index === state.totalSlides - 1;

  Array.from(thumbnailStrip.children).forEach((t, i) => {
    t.classList.toggle("active", i === index);
  });

  const activeThumb = thumbnailStrip.children[index];
  if (activeThumb) {
    activeThumb.scrollIntoView({ behavior: "smooth", inline: "nearest", block: "nearest" });
  }
}

export function toggleFullscreen() {
  state.isFullscreen = !state.isFullscreen;
  viewer.classList.toggle("fullscreen-mode", state.isFullscreen);
  btnFullscreen.innerHTML = state.isFullscreen ? "&#9633;" : "&#9974;";
  btnFullscreen.title = state.isFullscreen ? "Exit Fullscreen (F / Esc)" : "Fullscreen (F)";
}

export function reset() {
  state.totalSlides = 0;
  state.currentIndex = 0;
  state.isFullscreen = false;
  viewer.classList.add("hidden");
  viewer.classList.remove("fullscreen-mode");
  dropzone.classList.remove("hidden");
  loading.classList.add("hidden");
  slideContainer.innerHTML = "";
  thumbnailStrip.innerHTML = "";
  fileInput.value = "";
}

export function showError(msg) {
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
