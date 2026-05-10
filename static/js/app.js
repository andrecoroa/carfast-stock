document.addEventListener("DOMContentLoaded", () => {
  const shell = document.querySelector(".app-shell");
  const toggle = document.querySelector("[data-sidebar-toggle]");
  if (shell && toggle) {
    toggle.addEventListener("click", () => {
      const current = shell.dataset.sidebarState || "open";
      shell.dataset.sidebarState = current === "open" ? "closed" : "open";
      document.body.classList.toggle("sidebar-collapsed", shell.dataset.sidebarState === "closed");
    });
  }

  document.querySelectorAll("[data-filter-submit]").forEach((el) => {
    el.addEventListener("change", () => el.form && el.form.submit());
  });

  document.querySelectorAll("a.nav-item, a.hub-card, a.quick-action, a.quick-indicator").forEach((link) => {
    link.addEventListener("click", () => {
      if (link.target === "_blank" || link.href.startsWith("javascript:")) return;
      document.body.classList.add("is-navigating");
    });
  });
});
