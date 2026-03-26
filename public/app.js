import './jszip-patch.js';
import {
  dropzone, fileInput, btnPrev, btnNext, btnClose, btnFullscreen, viewer, slideContainer
} from './dom.js';
import { state } from './state.js';
import { handleFile } from './render.js';
import { updateSlideScale, goTo, reset, toggleFullscreen } from './ui.js';

/* ── Responsive slide scaling ───────────────────────────────── */
const slideResizer = new ResizeObserver(() => {
  updateSlideScale();
});
slideResizer.observe(slideContainer);

/* ── Drag-and-drop setup ───────────────────────────────────── */
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
btnPrev.addEventListener("click", () => goTo(state.currentIndex - 1));
btnNext.addEventListener("click", () => goTo(state.currentIndex + 1));
btnClose.addEventListener("click", reset);
btnFullscreen.addEventListener("click", toggleFullscreen);

document.addEventListener("keydown", (e) => {
  if (viewer.classList.contains("hidden")) return;
  if (e.key === "ArrowRight" || e.key === "ArrowDown") goTo(state.currentIndex + 1);
  if (e.key === "ArrowLeft" || e.key === "ArrowUp") goTo(state.currentIndex - 1);
  if (e.key === "Escape" && state.isFullscreen) toggleFullscreen();
  if (e.key === "f" || e.key === "F") toggleFullscreen();
});
