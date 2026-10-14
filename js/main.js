document.addEventListener("DOMContentLoaded", () => {
  const savedView = sessionStorage.getItem("dashboard-view");
  const initialView = savedView && document.getElementById(`view-${savedView}`) ? savedView : "monev";

  // Sidebar nav switching
  document.querySelectorAll(".nav-item[data-nav]").forEach((item) => {
    item.addEventListener("click", () => setActiveView(item.dataset.nav));
  });

  function setActiveView(target) {
      document.querySelectorAll(".nav-item[data-nav]").forEach((n) => n.classList.remove("active"));
      document.querySelector(`.nav-item[data-nav="${target}"]`)?.classList.add("active");
      document.querySelectorAll(".view").forEach((v) => v.classList.remove("active"));
      document.getElementById(`view-${target}`)?.classList.add("active");
      sessionStorage.setItem("dashboard-view", target);
  }
  setActiveView(initialView);

  // Tab switching (delegated, works for both systems' tabbars)
  document.querySelectorAll(".tabbar").forEach((bar) => {
    bar.querySelectorAll(".tab-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        bar.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");
        const panels = bar.parentElement.querySelectorAll(".tab-panel");
        panels.forEach((p) => p.classList.remove("active"));
        document.getElementById(`tab-${btn.dataset.tab}`).classList.add("active");
      });
    });
  });

  // Theme switcher (Terang / Gelap)
  const currentTheme = localStorage.getItem("bpvp-dashboard-theme") || "light";
  applyTheme(currentTheme);

  function applyTheme(theme) {
    document.documentElement.setAttribute("data-theme", theme);
    try { localStorage.setItem("bpvp-dashboard-theme", theme); } catch(e) {}
    const lightBtn = document.getElementById("themeLightBtn");
    const darkBtn = document.getElementById("themeDarkBtn");
    if (lightBtn && darkBtn) {
      lightBtn.classList.toggle("active", theme === "light");
      darkBtn.classList.toggle("active", theme === "dark");
    }
  }

  document.getElementById("themeLightBtn")?.addEventListener("click", () => applyTheme("light"));
  document.getElementById("themeDarkBtn")?.addEventListener("click", () => applyTheme("dark"));
});

function updateSidebarStatus(text, sub) {
  const el = document.getElementById("sidebarStatus");
  if (!el) return;
  el.replaceChildren();
  const title = document.createElement("b");
  const dot = document.createElement("span");
  dot.className = "status-pulse-dot";
  title.append(dot, document.createTextNode(" " + text));
  el.append(title, document.createTextNode(sub || ""));
}
