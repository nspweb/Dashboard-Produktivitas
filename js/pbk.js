(function () {
  const Q_COLS = Array.from({ length: 10 }, (_, i) => `q${i + 1}`);
  const Q_SHORT = Object.fromEntries(Q_COLS.map((q, i) => [q, `A${i + 1}`]));

  const state = {
    preRows: null,
    postRows: null,
    preFile: null,
    postFile: null,
    filters: { program: "Semua program", kabkota: "Semua kabupaten/kota", gender: "Semua jenis kelamin", search: "" },
    reportBlob: null,
    reportFilename: "laporan.xlsx",
  };

  let chartBar, chartRadar, chartModalRadar;

  const $ = (id) => document.getElementById(id);

  // ---------------------------------------------------------------- upload
  function wireDropzone(dzId, inputId, chipId, key) {
    const dz = $(dzId), input = $(inputId);
    ["dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("dragover"); }));
    ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("dragover"); }));
    dz.addEventListener("drop", (e) => {
      if (e.dataTransfer.files.length) { input.files = e.dataTransfer.files; input.dispatchEvent(new Event("change")); }
    });
    input.addEventListener("change", () => {
      const file = input.files[0];
      state[key] = file || null;
      if (file && (!file.name.toLowerCase().endsWith(".xlsx") || file.size > 15 * 1024 * 1024)) {
        state[key] = null;
        input.value = "";
        $(chipId).innerHTML = "";
        Util.notice($("pbkMsg"), "warn", "Pilih file .xlsx dengan ukuran maksimum 15 MB.");
        $("pbkProcessBtn").disabled = !(state.preFile || state.postFile);
        return;
      }
      if (!file) { $(chipId).innerHTML = ""; }
      else {
        $(chipId).innerHTML = `<div class="file-chip"><span class="dot"></span>
          <div><div class="name">${Util.escapeHtml(file.name)}</div><div class="meta">${(file.size / 1024).toFixed(0)} KB</div></div></div>`;
      }
      $("pbkProcessBtn").disabled = !(state.preFile || state.postFile);
    });
  }
  wireDropzone("pbkPreDropzone", "pbkPreFileInput", "pbkPreFileChip", "preFile");
  wireDropzone("pbkPostDropzone", "pbkPostFileInput", "pbkPostFileChip", "postFile");

  $("pbkProcessBtn").addEventListener("click", async () => {
    const btn = $("pbkProcessBtn");
    Util.setBusy(btn, true, "Memproses...");
    Util.clear($("pbkMsg"));
    try {
      const fd = new FormData();
      if (state.preFile) fd.append("pre_file", state.preFile);
      if (state.postFile) fd.append("post_file", state.postFile);
      const data = await Api.postForm("/api/pbk/process", fd);
      state.preRows = data.pre || null;
      state.postRows = data.post || null;

      $("pbkUploadCard").style.display = "none";
      $("pbkDashboard").style.display = "block";
      updateSidebarStatus("Data Pre/Post-Test siap",
        `${state.preRows ? state.preRows.length : 0} Pre-Test · ${state.postRows ? state.postRows.length : 0} Post-Test`);

      renderFilterOptions();
      renderAll();
    } catch (e) {
      Util.notice($("pbkMsg"), "error", Util.escapeHtml(e.message));
    } finally {
      Util.setBusy(btn, false, "Proses Data");
    }
  });

  // ---------------------------------------------------------------- filters
  function uniqueVals(rows, key) {
    const set = new Set();
    (rows || []).forEach((r) => { if (r[key]) set.add(String(r[key])); });
    return Array.from(set).sort();
  }

  function renderFilterOptions() {
    const programs = new Set([...uniqueVals(state.preRows, "program"), ...uniqueVals(state.postRows, "program")]);
    $("pbkFilterProgram").innerHTML = ['Semua program', ...programs].map((p) => `<option>${Util.escapeHtml(p)}</option>`).join("");
    $("pbkFilterKabkota").innerHTML = ['Semua kabupaten/kota', ...uniqueVals(state.preRows, "kabkota")].map((p) => `<option>${Util.escapeHtml(p)}</option>`).join("");
    $("pbkFilterGender").innerHTML = ['Semua jenis kelamin', ...uniqueVals(state.preRows, "gender")].map((p) => `<option>${Util.escapeHtml(p)}</option>`).join("");

    $("pbkRepProgram").innerHTML = [...programs].map((p) => `<option>${Util.escapeHtml(p)}</option>`).join("");

    ["pbkFilterProgram", "pbkFilterKabkota", "pbkFilterGender"].forEach((id) => {
      $(id).addEventListener("change", () => { readFilters(); renderAll(); });
    });
    $("pbkFilterSearch").addEventListener("input", () => { readFilters(); renderAll(); });
  }

  function readFilters() {
    state.filters.program = $("pbkFilterProgram").value;
    state.filters.kabkota = $("pbkFilterKabkota").value;
    state.filters.gender = $("pbkFilterGender").value;
    state.filters.search = $("pbkFilterSearch").value.trim().toLowerCase();
  }

  function applyFilters(rows) {
    if (!rows) return null;
    let d = rows;
    const f = state.filters;
    if (f.program !== "Semua program") d = d.filter((r) => r.program === f.program);
    if (f.kabkota !== "Semua kabupaten/kota") d = d.filter((r) => r.kabkota === f.kabkota);
    if (f.gender !== "Semua jenis kelamin") d = d.filter((r) => r.gender === f.gender);
    if (f.search) {
      d = d.filter((r) =>
        (r.nama || "").toLowerCase().includes(f.search) ||
        (r.nik_clean || "").toLowerCase().includes(f.search)
      );
    }
    return d;
  }

  // ---------------------------------------------------------------- merge
  function mergeAll(preF, postF) {
    const map = new Map();
    (preF || []).forEach((r, i) => {
      const key = r.nik_clean || `__no_nik_pre_${i}`;
      map.set(key, { nik_clean: key, nama: r.nama, program: r.program, kabkota: r.kabkota,
        gender: r.gender, telp: r.telp, tgl_mulai_fmt: r.tgl_mulai_fmt, tgl_selesai_fmt: r.tgl_selesai_fmt,
        rata_rata_skor_pre: r.rata_rata_skor, rata_rata_skor_post: null, _pre: r, _post: null });
    });
    (postF || []).forEach((r, i) => {
      const key = r.nik_clean || `__no_nik_post_${i}`;
      const existing = map.get(key);
      if (existing) {
        existing.rata_rata_skor_post = r.rata_rata_skor;
        existing._post = r;
        existing.nama = existing.nama || r.nama;
        existing.program = existing.program || r.program;
      } else {
        map.set(key, { nik_clean: key, nama: r.nama, program: r.program, kabkota: null,
          gender: null, telp: null, tgl_mulai_fmt: null, tgl_selesai_fmt: null,
          rata_rata_skor_pre: null, rata_rata_skor_post: r.rata_rata_skor, _pre: null, _post: r });
      }
    });
    const merged = Array.from(map.values());
    merged.forEach((m) => {
      m.status = (m.rata_rata_skor_pre !== null && m.rata_rata_skor_post !== null) ? "Lengkap"
        : (m.rata_rata_skor_pre !== null ? "Hanya Pre" : "Hanya Post");
      m.peningkatan = (m.rata_rata_skor_pre !== null && m.rata_rata_skor_post !== null)
        ? m.rata_rata_skor_post - m.rata_rata_skor_pre : null;
    });
    return merged;
  }

  // ---------------------------------------------------------------- render
  function renderAll() {
    const preF = applyFilters(state.preRows);
    const postF = applyFilters(state.postRows);
    const merged = mergeAll(preF, postF);
    const matched = merged.filter((m) => m.status === "Lengkap");

    renderKpi(preF, postF, merged);
    renderCharts(preF, postF);
    renderInsights(matched);
    renderTable(merged);
  }

  function avgOf(rows, key) {
    if (!rows || !rows.length) return null;
    const vals = rows.map((r) => Number(r[key])).filter((v) => !Number.isNaN(v) && v !== null);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  }

  function renderKpi(preF, postF, merged) {
    const avgPre = avgOf(preF, "rata_rata_skor");
    const avgPost = avgOf(postF, "rata_rata_skor");
    const kenaikan = (avgPre !== null && avgPost !== null) ? avgPost - avgPre : null;
    const relatif = (kenaikan !== null && avgPre) ? (kenaikan / avgPre) * 100 : null;

    $("pbkKpiRow").innerHTML = `
      <div class="kpi-card"><div class="kpi-value num">${merged.length}</div><div class="kpi-label">Total peserta terfilter</div></div>
      <div class="kpi-card"><div class="kpi-value num">${avgPre !== null ? avgPre.toFixed(1) + "/5" : "–"}</div><div class="kpi-label">Rata-rata skor Pre-Test</div></div>
      <div class="kpi-card"><div class="kpi-value num">${avgPost !== null ? avgPost.toFixed(1) + "/5" : "–"}</div><div class="kpi-label">Rata-rata skor Post-Test</div></div>
      <div class="kpi-card accent"><div class="kpi-value num">${Util.fmtDelta(kenaikan)}</div><div class="kpi-label">Kenaikan rata-rata skor</div></div>
      <div class="kpi-card"><div class="kpi-value num">${relatif !== null ? (relatif >= 0 ? "+" : "") + relatif.toFixed(1) + "%" : "–"}</div><div class="kpi-label">Perubahan relatif</div></div>
    `;
  }

  function renderCharts(preF, postF) {
    const labels = Q_COLS.map((q) => Q_SHORT[q]);
    const preMeans = Q_COLS.map((q) => avgOf(preF, q) || 0);
    const postMeans = Q_COLS.map((q) => avgOf(postF, q) || 0);

    if (chartBar) chartBar.destroy();
    chartBar = new Chart($("pbkChartBar"), {
      type: "bar",
      data: { labels, datasets: [
        { label: "Pre-Test", data: preMeans, backgroundColor: "#9AA5BD" },
        { label: "Post-Test", data: postMeans, backgroundColor: "#2f6f5e" },
      ] },
      options: { scales: { y: { min: 0, max: 5 } }, plugins: { legend: { position: "top" } } },
    });

    if (chartRadar) chartRadar.destroy();
    chartRadar = new Chart($("pbkChartRadar"), {
      type: "radar",
      data: { labels, datasets: [
        { label: "Pre-Test", data: preMeans, borderColor: "#9AA5BD", backgroundColor: "rgba(154,165,189,0.35)" },
        { label: "Post-Test", data: postMeans, borderColor: "#2f6f5e", backgroundColor: "rgba(47,111,94,0.3)" },
      ] },
      options: { scales: { r: { min: 0, max: 5 } } },
    });
  }

  function renderInsights(matched) {
    if (!matched.length) {
      $("pbkInsightUp").innerHTML = `<p class="card-sub">Belum ada data yang cocok antara Pre-Test dan Post-Test.</p>`;
      $("pbkInsightDown").innerHTML = `<p class="card-sub">Belum ada data yang cocok antara Pre-Test dan Post-Test.</p>`;
      return;
    }
    const up = [...matched].sort((a, b) => b.peningkatan - a.peningkatan).slice(0, 5);
    const down = [...matched].sort((a, b) => a.peningkatan - b.peningkatan).slice(0, 5);
    const itemHtml = (r, badgeClassFn) => `
      <div class="insight-item">
        <div><div class="insight-name">${Util.escapeHtml(r.nama || "–")}</div><div class="insight-sub">${Util.escapeHtml(r.program || "–")}</div></div>
        <div class="${badgeClassFn(r)}">${Util.fmtDelta(r.peningkatan)}</div>
      </div>`;
    $("pbkInsightUp").innerHTML = up.map((r) => itemHtml(r, () => "badge-pos")).join("");
    $("pbkInsightDown").innerHTML = down.map((r) => itemHtml(r, (r) => r.peningkatan < 0 ? "badge-neg" : "badge-pos")).join("");
  }

  function renderTable(merged) {
    const table = $("pbkParticipantTable");
    if (!merged.length) { table.innerHTML = `<tbody><tr><td class="empty-state">Belum ada data peserta untuk filter yang dipilih.</td></tr></tbody>`; return; }
    const sorted = [...merged].sort((a, b) => String(a.nama || "").localeCompare(String(b.nama || "")));
    const show = sorted.slice(0, 300);
    table.innerHTML = `
      <thead><tr><th>Nama</th><th>NIK</th><th>Program</th><th>Kab/Kota</th><th>Rata Pre</th><th>Rata Post</th><th>Perubahan</th><th>Status</th></tr></thead>
      <tbody>${show.map((r, i) => `
        <tr data-idx="${i}">
          <td>${Util.escapeHtml(r.nama)}</td>
          <td>${(r.nik_clean || "").startsWith("__no_nik") ? "–" : Util.escapeHtml(r.nik_clean)}</td>
          <td>${Util.escapeHtml(r.program)}</td>
          <td>${Util.escapeHtml(r.kabkota || "–")}</td>
          <td class="num">${Util.fmtScore(r.rata_rata_skor_pre)}</td>
          <td class="num">${Util.fmtScore(r.rata_rata_skor_post)}</td>
          <td class="num">${Util.fmtDelta(r.peningkatan)}</td>
          <td>${Util.escapeHtml(r.status)}</td>
        </tr>`).join("")}</tbody>
    `;
    table.querySelectorAll("tbody tr").forEach((tr) => {
      tr.addEventListener("click", () => openDetail(show[Number(tr.dataset.idx)]));
    });
  }

  // ---------------------------------------------------------------- detail modal
  function openDetail(r) {
    $("pbkModalName").textContent = r.nama || "–";
    $("pbkModalSub").textContent = `${r.program || "–"} · ${r.kabkota || "–"}`;
    const periode = (r._pre && r._pre.tgl_mulai_fmt && r._pre.tgl_selesai_fmt)
      ? `${r._pre.tgl_mulai_fmt} s.d. ${r._pre.tgl_selesai_fmt}` : "–";
    $("pbkModalDetails").innerHTML = `
      <div class="detail-item"><div class="k">NIK</div><div class="v">${(r.nik_clean || "").startsWith("__no_nik") ? "–" : Util.escapeHtml(r.nik_clean)}</div></div>
      <div class="detail-item"><div class="k">Jenis kelamin</div><div class="v">${Util.escapeHtml(r.gender || "–")}</div></div>
      <div class="detail-item"><div class="k">No. WhatsApp</div><div class="v">${Util.escapeHtml(r.telp || "–")}</div></div>
      <div class="detail-item"><div class="k">Periode pelatihan</div><div class="v">${Util.escapeHtml(periode)}</div></div>
      <div class="detail-item"><div class="k">Rata-rata Pre-Test</div><div class="v">${Util.fmtScore(r.rata_rata_skor_pre)}</div></div>
      <div class="detail-item"><div class="k">Rata-rata Post-Test</div><div class="v">${Util.fmtScore(r.rata_rata_skor_post)}</div></div>
    `;
    const labels = Q_COLS.map((q) => Q_SHORT[q]);
    const preInd = Q_COLS.map((q) => (r._pre ? Number(r._pre[q]) : null) || 0);
    const postInd = Q_COLS.map((q) => (r._post ? Number(r._post[q]) : null) || 0);
    if (chartModalRadar) chartModalRadar.destroy();
    chartModalRadar = new Chart($("pbkModalRadar"), {
      type: "radar",
      data: { labels, datasets: [
        { label: "Pre-Test", data: preInd, borderColor: "#9AA5BD", backgroundColor: "rgba(154,165,189,0.35)" },
        { label: "Post-Test", data: postInd, borderColor: "#2f6f5e", backgroundColor: "rgba(47,111,94,0.3)" },
      ] },
      options: { scales: { r: { min: 0, max: 5 } } },
    });
    $("pbkModal").classList.add("active");
  }
  $("pbkModalClose").addEventListener("click", () => $("pbkModal").classList.remove("active"));
  $("pbkModal").addEventListener("click", (e) => { if (e.target === $("pbkModal")) $("pbkModal").classList.remove("active"); });

  // ---------------------------------------------------------------- download data
  $("pbkDownloadPre").addEventListener("click", () => {
    const rows = applyFilters(state.preRows);
    if (!rows || !rows.length) return;
    Util.downloadText(Util.arrayToCsv(rows), "Data_PreTest_Pelatihan.csv");
  });
  $("pbkDownloadPost").addEventListener("click", () => {
    const rows = applyFilters(state.postRows);
    if (!rows || !rows.length) return;
    Util.downloadText(Util.arrayToCsv(rows), "Data_PostTest_Pelatihan.csv");
  });

  // ---------------------------------------------------------------- laporan resmi
  function refreshReportPreview() {
    const kompetensi = $("pbkRepKompetensi").value;
    const program = $("pbkRepProgram").value;
    $("pbkRepKompetensiPreview").textContent = "Pratinjau: " + formatTrainingTitle(kompetensi);
    $("pbkRepProgramPreview").textContent = program ? formatProgramTitle(program) : "–";
    const start = $("pbkRepTglMulai").value, end = $("pbkRepTglSelesai").value;
    if (start && end) $("pbkRepDatePreview").textContent = "Pratinjau tanggal: " + formatDateRangeId(start, end);
  }
  ["pbkRepKompetensi", "pbkRepProgram", "pbkRepTglMulai", "pbkRepTglSelesai"].forEach((id) => {
    $(id).addEventListener("input", refreshReportPreview);
    $(id).addEventListener("change", refreshReportPreview);
  });

  const HARI_ID = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  const BULAN_ID = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  function formatDateRangeId(startStr, endStr) {
    const s = new Date(startStr + "T00:00:00"), e = new Date(endStr + "T00:00:00");
    const fmt = (d) => `${HARI_ID[d.getDay()]} ${d.getDate()} ${BULAN_ID[d.getMonth() + 1]}`;
    return `TGL ${fmt(s)} S.D ${fmt(e)} ${e.getFullYear()}`;
  }
  function formatTrainingTitle(extra) {
    const PREFIX = "PELATIHAN BERBASIS KOMPETENSI KEJURUAN";
    let name = (extra || "").trim().toUpperCase().replace(/\s+/g, " ");
    if (!name) return PREFIX;
    name = name.replace(new RegExp("^" + PREFIX + "\\s*"), "");
    return `${PREFIX} ${name}`.trim();
  }
  function formatProgramTitle(program) {
    const PREFIX = "PROGRAM PELATIHAN DASAR";
    let name = (program || "").trim().toUpperCase().replace(/\s+/g, " ");
    if (!name) return PREFIX;
    name = name.replace(new RegExp("^" + PREFIX + "\\s*"), "");
    name = name.replace(/\s*\([^)]*\)/g, "").trim();
    if (!name.includes("PAKET")) {
      const m = name.match(/^(.*\S)\s+(\d+)$/);
      if (m) name = `${m[1]} PAKET ${m[2]}`;
    }
    return `${PREFIX} ${name}`.trim();
  }

  $("pbkGenerateReportBtn").addEventListener("click", async () => {
    const btn = $("pbkGenerateReportBtn");
    Util.clear($("pbkReportMsg"));

    const jenis = $("pbkRepJenis").value;
    const program = $("pbkRepProgram").value;
    const kompetensi = $("pbkRepKompetensi").value.trim();
    const tglMulai = $("pbkRepTglMulai").value;
    const tglSelesai = $("pbkRepTglSelesai").value;

    if (!program) { Util.notice($("pbkReportMsg"), "warn", "Belum ada program pelatihan pada data."); return; }
    if (!kompetensi) { Util.notice($("pbkReportMsg"), "warn", "Isi nama kompetensi terlebih dahulu."); return; }
    if (!tglMulai || !tglSelesai) { Util.notice($("pbkReportMsg"), "warn", "Isi tanggal mulai dan berakhir pelatihan."); return; }
    if (tglSelesai < tglMulai) { Util.notice($("pbkReportMsg"), "error", "Tanggal berakhir tidak boleh sebelum tanggal mulai."); return; }

    const source = jenis === "Pre-Test" ? state.preRows : state.postRows;
    const rows = (source || []).filter((r) => r.program === program);
    if (!rows.length) { Util.notice($("pbkReportMsg"), "error", "Tidak ada responden pada program ini."); return; }

    Util.setBusy(btn, true, "Membuat laporan...");
    try {
      const res = await Api.postJson("/api/pbk/report", {
        rows, jenis_test: jenis, training_title_input: kompetensi, program_name: program,
        tanggal_mulai: tglMulai, tanggal_selesai: tglSelesai,
      });
      const disp = res.headers.get("content-disposition") || "";
      const match = disp.match(/filename="?([^"]+)"?/);
      state.reportFilename = match ? match[1] : "Laporan.xlsx";
      state.reportBlob = await res.blob();
      $("pbkReportDownloads").style.display = "flex";
      Util.notice($("pbkReportMsg"), "success", `Laporan siap — ${rows.length} peserta dimasukkan ke tabel.`);
    } catch (e) {
      Util.notice($("pbkReportMsg"), "error", Util.escapeHtml(e.message));
    } finally {
      Util.setBusy(btn, false, "Buat Laporan");
    }
  });

  $("pbkDownloadReportXlsx").addEventListener("click", () => {
    if (!state.reportBlob) return;
    const url = URL.createObjectURL(state.reportBlob);
    const a = document.createElement("a");
    a.href = url; a.download = state.reportFilename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  });

  $("pbkConvertPdfBtn").addEventListener("click", async () => {
    if (!state.reportBlob) return;
    const btn = $("pbkConvertPdfBtn");
    Util.setBusy(btn, true, "Mengonversi ke PDF...");
    try {
      const base64 = await blobToBase64(state.reportBlob);
      const res = await Api.postJson("/api/convert_pdf", { xlsx_base64: base64, filename: state.reportFilename });
      await Api.downloadBlobFrom(res, state.reportFilename.replace(/\.xlsx$/, ".pdf"));
    } catch (e) {
      Util.notice($("pbkReportMsg"), "error", Util.escapeHtml(e.message));
    } finally {
      Util.setBusy(btn, false, "Konversi & Unduh PDF");
    }
  });

  function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result.split(",")[1]);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }
})();
