// Helper fetch tipis untuk semua endpoint /api/*
const Api = {
  async postForm(url, formData) {
    const res = await fetchWithTimeout(url, { method: "POST", body: formData });
    const ct = res.headers.get("content-type") || "";
    if (!ct.includes("application/json")) {
      throw new Error("Server mengembalikan respons tak terduga.");
    }
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Terjadi kesalahan.");
    return data;
  },

  async postJson(url, body) {
    const res = await fetchWithTimeout(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const ct = res.headers.get("content-type") || "";
    if (ct.includes("application/json")) {
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Terjadi kesalahan.");
      return data;
    }
    if (!res.ok) throw new Error("Terjadi kesalahan saat memproses permintaan.");
    return res; // caller handles blob (e.g. xlsx/pdf download)
  },

  async downloadBlobFrom(response, fallbackName) {
    const blob = await response.blob();
    const disp = response.headers.get("content-disposition") || "";
    const match = disp.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : fallbackName;
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    return blob;
  },
};

async function fetchWithTimeout(url, options = {}, timeoutMs = 120000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("Permintaan terlalu lama. Periksa koneksi lalu coba lagi.");
    }
    throw new Error("Tidak dapat terhubung ke server. Periksa koneksi lalu coba lagi.");
  } finally {
    clearTimeout(timeout);
  }
}

// Util umum
const Util = {
  escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  },
  fmtScore(v) {
    if (v === null || v === undefined || Number.isNaN(v)) return "–";
    return Number(v).toFixed(1);
  },
  fmtDelta(v) {
    if (v === null || v === undefined || Number.isNaN(v)) return "–";
    const n = Number(v);
    return (n >= 0 ? "+" : "") + n.toFixed(1);
  },
  arrayToCsv(rows) {
    if (!rows.length) return "";
    const headers = Object.keys(rows[0]);
    const esc = (v) => {
      if (v === null || v === undefined) return "";
      const s = String(v).replace(/"/g, '""');
      return /[",\n]/.test(s) ? `"${s}"` : s;
    };
    const lines = [headers.join(",")];
    for (const r of rows) lines.push(headers.map((h) => esc(r[h])).join(","));
    return lines.join("\n");
  },
  downloadText(text, filename, mime = "text/csv;charset=utf-8") {
    const blob = new Blob(["\uFEFF" + text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  },
  setBusy(btn, busy, label) {
    if (!btn) return;
    if (busy) {
      btn.dataset.label = btn.dataset.label || btn.textContent;
      btn.disabled = true;
      btn.innerHTML = `<span class="spinner"></span> ${label || "Memproses..."}`;
    } else {
      btn.disabled = false;
      btn.textContent = btn.dataset.label || label || btn.textContent;
      delete btn.dataset.label;
    }
  },
  notice(container, type, message) {
    if (!container) return;
    container.innerHTML = `<div class="notice notice-${type}">${message}</div>`;
  },
  clear(container) {
    if (container) container.innerHTML = "";
  },
};
