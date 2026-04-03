const navToggle = document.getElementById("navToggle");
const siteNav = document.getElementById("siteNav");

function closeMobileNav() {
  if (!navToggle || !siteNav) return;
  siteNav.classList.remove("is-open");
  navToggle.setAttribute("aria-expanded", "false");
}

if (navToggle && siteNav) {
  navToggle.addEventListener("click", () => {
    const isOpen = siteNav.classList.toggle("is-open");
    navToggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
  });

  document.addEventListener("click", (event) => {
    if (window.innerWidth > 860) return;
    const clickedInsideMenu = siteNav.contains(event.target) || navToggle.contains(event.target);
    if (!clickedInsideMenu) closeMobileNav();
  });

  window.addEventListener("resize", () => {
    if (window.innerWidth > 860) closeMobileNav();
  });
}

const currentPage = document.body?.dataset?.page;
if (currentPage) {
  document.querySelectorAll("[data-nav-page]").forEach((link) => {
    link.classList.toggle("active", link.dataset.navPage === currentPage);
  });
}

const yearNode = document.getElementById("year");
if (yearNode) {
  yearNode.textContent = String(new Date().getFullYear());
}
