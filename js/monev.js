(function () {
  const state = {
    rows: [],
    programCol: null,
    timestampCol: null,
    infoCol: null,
    instructorCol: null,
    scoreColMap: {},
    commentColMap: {},
    selectedPrograms: new Set(),
    sheets: null,
    file: null,
    reportBlob: null,
    reportFilename: "laporan.xlsx",
  };

  let chartCount, chartAvg, chartSrc;

  const $ = (id) => document.getElementById(id);

  // ---------------------------------------------------------------- upload
  const dz = $("monevDropzone");
  const fileInput = $("monevFileInput");

  ["dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("dragover"); }));
  ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("dragover"); }));
  dz.addEventListener("drop", (e) => {
    if (e.dataTransfer.files.length) {
      fileInput.files = e.dataTransfer.files;
      fileInput.dispatchEvent(new Event("change"));
    }
  });

  fileInput.addEventListener("change", () => {
    const file = fileInput.files[0];
    Util.clear($("monevMsg"));
    $("monevSheetChoice").innerHTML = "";
    if (!file) { $("monevProcessBtn").disabled = true; return; }
    if (!file.name.toLowerCase().endsWith(".xlsx") || file.size > 15 * 1024 * 1024) {
      state.file = null;
      fileInput.value = "";
      $("monevProcessBtn").disabled = true;
      Util.notice($("monevMsg"), "warn", "Pilih file .xlsx dengan ukuran maksimum 15 MB.");
      return;
    }
    state.file = file;
    $("monevFileChip").innerHTML = `
      <div class="file-chip">
        <span class="dot"></span>
        <div><div class="name">${Util.escapeHtml(file.name)}</div>
        <div class="meta">${(file.size / 1024).toFixed(0)} KB</div></div>
      </div>`;
    $("monevProcessBtn").disabled = false;
  });

  $("monevProcessBtn").onclick = () => processFile();

  async function processFile(sheetName) {
    const btn = $("monevProcessBtn");
    Util.setBusy(btn, true, "Memproses...");
    Util.clear($("monevMsg"));
    try {
      const fd = new FormData();
      fd.append("file", state.file);
      if (sheetName) fd.append("sheet_name", sheetName);

      const data = await Api.postForm("/api/monev/process", fd);

      if (data.needs_sheet_choice) {
        $("monevSheetChoice").innerHTML = `
          <div class="field" style="max-width:360px; margin-top:12px;">
            <label class="field-label">Pilih sheet yang berisi data responden</label>
            <select id="monevSheetSelect">
              ${data.sheets.map((s) => `<option value="${Util.escapeHtml(s)}">${Util.escapeHtml(s)}</option>`).join("")}
            </select>
          </div>`;
        Util.setBusy(btn, false);
          btn.onclick = () => processFile($("monevSheetSelect").value);
        return;
      }

      state.rows = data.rows;
      state.programCol = data.program_col;
      state.timestampCol = data.timestamp_col;
      state.infoCol = data.info_col;
      state.instructorCol = data.instructor_col;
      state.scoreColMap = data.score_col_map;
      state.commentColMap = data.comment_col_map;
      state.selectedPrograms = new Set(getPrograms());

      $("monevUploadCard").style.display = "none";
      $("monevDashboard").style.display = "block";
      updateSidebarStatus("Data Monev siap", `Sheet "${data.sheet_choice}" · ${data.total_rows} responden`);
      renderFilters();
      renderAll();
    } catch (e) {
      Util.notice($("monevMsg"), "error", Util.escapeHtml(e.message));
    } finally {
      Util.setBusy(btn, false);
    }
  }

  // ---------------------------------------------------------------- filters
  function getPrograms() {
    const set = new Set();
    state.rows.forEach((r) => { const v = r[state.programCol]; if (v) set.add(String(v)); });
    return Array.from(set).sort();
  }

  function renderFilters() {
    const programs = getPrograms();
    $("monevProgramFilter").innerHTML = programs.map((p) => `
      <span class="chip-toggle on" data-program="${Util.escapeHtml(p)}">${Util.escapeHtml(p)}</span>
    `).join("");
    $("monevProgramFilter").querySelectorAll(".chip-toggle").forEach((chip) => {
      chip.addEventListener("click", () => {
        const p = chip.dataset.program;
        if (state.selectedPrograms.has(p)) { state.selectedPrograms.delete(p); chip.classList.remove("on"); }
        else { state.selectedPrograms.add(p); chip.classList.add("on"); }
        renderAll();
      });
    });

    // Populate program selects used elsewhere
    const opts = programs.map((p) => `<option value="${Util.escapeHtml(p)}">${Util.escapeHtml(p)}</option>`).join("");
    $("monevRecapProgram").innerHTML = opts;
    $("monevReportProgram").innerHTML = opts;
    updateRecapInfo();
    updateReportInfo();
  }

  function filteredRows() {
    return state.rows.filter((r) => state.selectedPrograms.has(String(r[state.programCol])));
  }

  // ---------------------------------------------------------------- render all
  function renderAll() {
    const rows = filteredRows();
    renderKpi(rows);
    renderTable(rows);
    renderCharts(rows);
    renderComments(rows);
  }

  function renderKpi(rows) {
    const scoreCols = Object.values(state.scoreColMap);
    let overallAvg = null;
    if (scoreCols.length) {
      let sum = 0, count = 0;
      rows.forEach((r) => scoreCols.forEach((c) => { const v = Number(r[c]); if (!Number.isNaN(v) && r[c] !== null) { sum += v; count++; } }));
      overallAvg = count ? sum / count : null;
    }
    $("monevKpiRow").innerHTML = `
      <div class="kpi-card"><div class="kpi-value num">${rows.length}</div><div class="kpi-label">Jumlah Responden (terfilter)</div></div>
      <div class="kpi-card"><div class="kpi-value num">${state.selectedPrograms.size}</div><div class="kpi-label">Jumlah Program Terpilih</div></div>
      <div class="kpi-card accent"><div class="kpi-value num">${overallAvg !== null ? overallAvg.toFixed(2) + " / 5" : "–"}</div><div class="kpi-label">Rata-rata Skor Keseluruhan</div></div>
    `;
  }

  function renderTable(rows) {
    const table = $("monevDataTable");
    if (!rows.length) { table.innerHTML = `<tbody><tr><td class="empty-state">Tidak ada data untuk filter ini.</td></tr></tbody>`; return; }
    const cols = Object.keys(rows[0]);
    const show = rows.slice(0, 300);
    table.innerHTML = `
      <thead><tr>${cols.map((c) => `<th>${Util.escapeHtml(c)}</th>`).join("")}</tr></thead>
      <tbody>${show.map((r) => `<tr>${cols.map((c) => `<td>${Util.escapeHtml(r[c])}</td>`).join("")}</tr>`).join("")}</tbody>
    `;
  }

  $("monevDownloadCsv").addEventListener("click", () => {
    const rows = filteredRows();
    if (!rows.length) return;
    Util.downloadText(Util.arrayToCsv(rows), "evaluasi_pelatihan_terfilter.csv");
  });

  $("monevDownloadXlsx").addEventListener("click", async () => {
    const rows = filteredRows();
    if (!rows.length) return;
    // Kirim sebagai CSV dari client -- unduhan "data mentah" cukup dalam bentuk CSV
    // yang bisa dibuka Excel; menghindari perlu endpoint tambahan hanya untuk ini.
    Util.downloadText(Util.arrayToCsv(rows), "evaluasi_pelatihan_terfilter.csv");
  });

  // ---------------------------------------------------------------- charts
  function renderCharts(rows) {
    const programCounts = {};
    rows.forEach((r) => { const p = String(r[state.programCol]); programCounts[p] = (programCounts[p] || 0) + 1; });
    const programLabels = Object.keys(programCounts);

    if (chartCount) chartCount.destroy();
    chartCount = new Chart($("monevChartCount"), {
      type: "bar",
      data: { labels: programLabels, datasets: [{ label: "Jumlah Responden", data: programLabels.map((p) => programCounts[p]), backgroundColor: "#2A3D63" }] },
      options: { indexAxis: "y", plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true } } },
    });

    const scoreLabels = Object.keys(state.scoreColMap);
    const scoreAverages = scoreLabels.map((label) => {
      const col = state.scoreColMap[label];
      const vals = rows.map((r) => Number(r[col])).filter((v) => !Number.isNaN(v));
      return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    });
    if (chartAvg) chartAvg.destroy();
    chartAvg = new Chart($("monevChartAvg"), {
      type: "bar",
      data: { labels: scoreLabels, datasets: [{ label: "Rata-rata Skor", data: scoreAverages, backgroundColor: "#A9782E" }] },
      options: { indexAxis: "y", plugins: { legend: { display: false } }, scales: { x: { min: 0, max: 5 } } },
    });

    if (state.infoCol) {
      $("monevChartSrcCard").style.display = "block";
      const srcCounts = {};
      rows.forEach((r) => { const v = String(r[state.infoCol] || "–"); srcCounts[v] = (srcCounts[v] || 0) + 1; });
      if (chartSrc) chartSrc.destroy();
      chartSrc = new Chart($("monevChartSrc"), {
        type: "doughnut",
        data: { labels: Object.keys(srcCounts), datasets: [{ data: Object.values(srcCounts), backgroundColor: ["#1B2A4A", "#A9782E", "#7C8AB0", "#C9D0DC", "#2A3D63", "#D8C08A"] }] },
      });
    }
  }

  // ---------------------------------------------------------------- comments
  function renderComments(rows) {
    const labels = Object.keys(state.commentColMap);
    const select = $("monevCommentSelect");
    if (!labels.length) {
      select.innerHTML = "";
      $("monevCommentList").innerHTML = `<div class="notice notice-info">Tidak ada kolom komentar/keluhan yang terdeteksi.</div>`;
      return;
    }
    select.innerHTML = labels.map((l) => `<option value="${Util.escapeHtml(l)}">${Util.escapeHtml(l)}</option>`).join("");
    select.onchange = () => renderCommentList(rows);
    renderCommentList(rows);
  }

  function renderCommentList(rows) {
    const label = $("monevCommentSelect").value;
    const col = state.commentColMap[label];
    const grouped = {};
    rows.forEach((r) => {
      const val = r[col];
      if (!val || String(val).trim() === "") return;
      const p = String(r[state.programCol]);
      (grouped[p] = grouped[p] || []).push(val);
    });
    const programs = Object.keys(grouped);
    if (!programs.length) { $("monevCommentList").innerHTML = `<div class="notice notice-info">Tidak ada komentar untuk kombinasi filter ini.</div>`; return; }
    $("monevCommentList").innerHTML = programs.map((p) => `
      <details class="comment-group">
        <summary>${Util.escapeHtml(p)} (${grouped[p].length} komentar)</summary>
        <ul>${grouped[p].map((c) => `<li>${Util.escapeHtml(c)}</li>`).join("")}</ul>
      </details>
    `).join("");
  }

  // ---------------------------------------------------------------- rekap komentar
  $("monevRecapProgram").addEventListener("change", updateRecapInfo);
  function updateRecapInfo() {
    const program = $("monevRecapProgram").value;
    const rows = state.rows.filter((r) => String(r[state.programCol]) === program);
    Util.clear($("monevRecapInfo"));
    if (!Object.keys(state.commentColMap).length) {
      Util.notice($("monevRecapInfo"), "warn", "Tidak ada kolom komentar yang terdeteksi pada sheet ini.");
    } else {
      $("monevRecapInfo").innerHTML = `<p class="card-sub">Jumlah Responden (= jumlah baris rekap): <b>${rows.length}</b></p>`;
    }
  }

  $("monevRecapBtn").addEventListener("click", async () => {
    const btn = $("monevRecapBtn");
    const program = $("monevRecapProgram").value;
    const rows = state.rows.filter((r) => String(r[state.programCol]) === program);
    if (!rows.length) { Util.notice($("monevRecapInfo"), "warn", "Tidak ada responden pada program ini."); return; }
    Util.setBusy(btn, true, "Membuat rekap...");
    try {
      const res = await Api.postJson("/api/monev/comment_recap", {
        rows, info_col: state.infoCol, comment_col_map: state.commentColMap, program_name: program,
      });
      await Api.downloadBlobFrom(res, "Komentar_Peserta.xlsx");
    } catch (e) {
      Util.notice($("monevRecapInfo"), "error", Util.escapeHtml(e.message));
    } finally {
      Util.setBusy(btn, false, "Unduh Rekap Komentar (.xlsx)");
    }
  });

  // ---------------------------------------------------------------- laporan resmi
  const MONEV_DAYS = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  const MONEV_MONTHS = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  function formatMonevDateRange(start, end) {
    const startDate = new Date(`${start}T00:00:00`);
    const endDate = new Date(`${end}T00:00:00`);
    const formatDate = (date) => `${MONEV_DAYS[date.getDay()]} ${date.getDate()} ${MONEV_MONTHS[date.getMonth() + 1]}`;
    return `TGL ${formatDate(startDate)} S.D ${formatDate(endDate)} ${endDate.getFullYear()}`;
  }

  function refreshMonevDatePreview() {
    const start = $("monevDateStart").value;
    const end = $("monevDateEnd").value;
    const preview = $("monevDatePreview");
    if (!start || !end) { preview.textContent = ""; return; }
    preview.textContent = end < start ? "Tanggal selesai harus sama atau setelah tanggal mulai." : `Pratinjau tanggal: ${formatMonevDateRange(start, end)}`;
  }

  $("monevTrainingTitle").addEventListener("input", () => {
    $("monevTrainingTitle").dataset.autoFilled = "0";
  });
  $("monevDateStart").addEventListener("change", () => {
    $("monevDateEnd").min = $("monevDateStart").value;
    refreshMonevDatePreview();
  });
  $("monevDateEnd").addEventListener("change", refreshMonevDatePreview);
  $("monevReportProgram").addEventListener("change", updateReportInfo);
  function updateReportInfo() {
    const program = $("monevReportProgram").value;
    const rows = state.rows.filter((r) => String(r[state.programCol]) === program);
    const matched = Object.keys(state.scoreColMap).length;

    // Otomatis isi Nama Pelatihan dengan program terpilih jika belum diisi atau saat ganti program
    if (program && (!$("monevTrainingTitle").value || $("monevTrainingTitle").dataset.autoFilled === "1")) {
      $("monevTrainingTitle").value = program;
      $("monevTrainingTitle").dataset.autoFilled = "1";
    }

    if (state.instructorCol && rows.length) {
      const instructors = rows
        .map((r) => r[state.instructorCol])
        .filter((v) => v !== null && v !== undefined && String(v).trim() !== "" && isNaN(Number(String(v).trim())));
      if (instructors.length) {
        const freq = {};
        instructors.forEach((n) => {
          const name = String(n).trim();
          freq[name] = (freq[name] || 0) + 1;
        });
        const topInstructor = Object.keys(freq).sort((a, b) => freq[b] - freq[a])[0];
        if (topInstructor) {
          $("monevInstructorName").value = topInstructor;
        }
      }
    }

    Util.clear($("monevReportInfo"));
    let html = `<div class="kpi-row" style="margin-bottom:12px;">
      <div class="kpi-card"><div class="kpi-value num">${rows.length}</div><div class="kpi-label">Jumlah Peserta</div></div>
      <div class="kpi-card"><div class="kpi-value num">${matched}/17</div><div class="kpi-label">Pertanyaan Terisi</div></div>
    </div>`;
    if (rows.length > 16) html += `<div class="notice notice-info">Jumlah peserta (${rows.length}) lebih dari 16, kolom cadangan template akan otomatis dipakai (kapasitas maksimum 24 peserta).</div>`;
    if (rows.length > 24) html += `<div class="notice notice-warn">Jumlah peserta (${rows.length}) melebihi kapasitas maksimum template (24). Hanya 24 peserta pertama yang akan dimasukkan.</div>`;
    $("monevReportInfo").innerHTML = html;
  }

  $("monevGenerateReportBtn").addEventListener("click", async () => {
    const btn = $("monevGenerateReportBtn");
    const program = $("monevReportProgram").value;
    const rows = state.rows.filter((r) => String(r[state.programCol]) === program);
    const training_title = $("monevTrainingTitle").value.trim();
    const instructor_name = $("monevInstructorName").value.trim();
    const dateStart = $("monevDateStart").value;
    const dateEnd = $("monevDateEnd").value;
    const date_text = dateStart && dateEnd ? formatMonevDateRange(dateStart, dateEnd) : "";

    Util.clear($("monevReportMsg"));
    if (!rows.length) { Util.notice($("monevReportMsg"), "error", "Tidak ada responden pada program ini."); return; }
    if (!training_title || !instructor_name || !date_text) {
      Util.notice($("monevReportMsg"), "warn", "Lengkapi nama pelatihan, nama instruktur, tanggal mulai, dan tanggal selesai.");
      return;
    }
    if (dateEnd < dateStart) {
      Util.notice($("monevReportMsg"), "error", "Tanggal selesai tidak boleh sebelum tanggal mulai.");
      return;
    }

    Util.setBusy(btn, true, "Membuat laporan...");
    try {
      const res = await Api.postJson("/api/monev/report", {
        rows, score_col_map: state.scoreColMap, program_name: program,
        training_title, date_text, instructor_name,
      });
      const disp = res.headers.get("content-disposition") || "";
      const match = disp.match(/filename="?([^"]+)"?/);
      state.reportFilename = match ? match[1] : "Laporan_Evaluasi.xlsx";
      state.reportBlob = await res.blob();
      $("monevReportDownloads").style.display = "flex";
      Util.notice($("monevReportMsg"), "success", "Laporan berhasil dibuat.");
    } catch (e) {
      Util.notice($("monevReportMsg"), "error", Util.escapeHtml(e.message));
    } finally {
      Util.setBusy(btn, false, "Buat Laporan Resmi (.xlsx)");
    }
  });

  $("monevDownloadReportXlsx").addEventListener("click", () => {
    if (!state.reportBlob) return;
    const url = URL.createObjectURL(state.reportBlob);
    const a = document.createElement("a");
    a.href = url; a.download = state.reportFilename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  });

  $("monevConvertPdfBtn").addEventListener("click", async () => {
    if (!state.reportBlob) return;
    const btn = $("monevConvertPdfBtn");
    Util.setBusy(btn, true, "Mengonversi ke PDF...");
    try {
      const base64 = await blobToBase64(state.reportBlob);
      const res = await Api.postJson("/api/convert_pdf", { xlsx_base64: base64, filename: state.reportFilename });
      await Api.downloadBlobFrom(res, state.reportFilename.replace(/\.xlsx$/, ".pdf"));
    } catch (e) {
      Util.notice($("monevReportMsg"), "error", Util.escapeHtml(e.message));
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
