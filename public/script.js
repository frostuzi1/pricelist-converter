(function () {
  /** Matches Excel template: DESCRIPTION | PRICE | DEAL | DISC % */
  const COLUMNS = ["DESCRIPTION", "PRICE", "DEAL", "DISC %"];

  function apiUrl(path) {
    const raw = (
      document.querySelector('meta[name="api-base"]')?.getAttribute("content") || ""
    ).trim();
    const p = path.startsWith("/") ? path : `/${path}`;
    if (!raw) return p;
    const base = raw.replace(/\/$/, "");
    if (base.startsWith("http://") || base.startsWith("https://")) {
      return base + p;
    }
    return base + p;
  }

  const LEGACY_KEY_MAP = {
    DESCRIPTION: ["DESCRIPTION", "Description"],
    PRICE: ["PRICE", "Price"],
    DEAL: ["DEAL", "Deal"],
    "DISC %": ["DISC %", "DISC%", "Discount", "Disc %", "Disc"],
  };

  const fileInput = document.getElementById("file-input");
  const statusEl = document.getElementById("status");
  const previewContainer = document.getElementById("preview-container");
  const previewImage = document.getElementById("preview-image");
  const resultContainer = document.getElementById("result-container");
  const exportToolbar = document.getElementById("export-toolbar");
  const exportButton = document.getElementById("export-button");
  const resultModal = document.getElementById("result-modal");
  const modalTitle = document.getElementById("modal-title");
  const modalMessage = document.getElementById("modal-message");
  const modalClose = document.getElementById("modal-close");
  const modalPanel = resultModal.querySelector(".modal-overlay__panel");

  let previewObjectUrl = null;
  let lastExportRows = null;
  let lastExportBaseName = "invoice";

  function revokePreviewUrl() {
    if (previewObjectUrl) {
      URL.revokeObjectURL(previewObjectUrl);
      previewObjectUrl = null;
    }
  }

  const previewLabelEl = document.querySelector(".preview-label");

  function setPreview(files) {
    revokePreviewUrl();
    const list = Array.isArray(files) ? files : files ? [files] : [];
    const first = list.find((f) => f.type.startsWith("image/"));
    if (!first) {
      previewContainer.hidden = true;
      previewImage.removeAttribute("src");
      previewImage.alt = "";
      if (previewLabelEl) previewLabelEl.textContent = "Uploaded image";
      return;
    }
    const imageFiles = list.filter((f) => f.type.startsWith("image/"));
    previewObjectUrl = URL.createObjectURL(first);
    previewImage.src = previewObjectUrl;
    previewImage.alt = first.name ? `Preview: ${first.name}` : "Uploaded invoice preview";
    previewContainer.hidden = false;
    if (previewLabelEl) {
      previewLabelEl.textContent =
        imageFiles.length > 1
          ? `Uploaded images (${imageFiles.length})`
          : "Uploaded image";
    }
  }

  function setStatus(message, isError) {
    statusEl.hidden = false;
    statusEl.textContent = message;
    statusEl.classList.toggle("status--error", Boolean(isError));
    statusEl.classList.toggle("status--loading", !isError && message);
  }

  function showResultModal(title, message) {
    modalTitle.textContent = title;
    modalMessage.textContent = message;
    resultModal.hidden = false;
    document.body.style.overflow = "hidden";
    modalClose.focus();
  }

  function hideResultModal() {
    resultModal.hidden = true;
    document.body.style.overflow = "";
  }

  resultModal.addEventListener("click", (e) => {
    if (e.target === resultModal) hideResultModal();
  });
  if (modalPanel) {
    modalPanel.addEventListener("click", (e) => e.stopPropagation());
  }
  modalClose.addEventListener("click", hideResultModal);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !resultModal.hidden) {
      hideResultModal();
    }
  });

  function clearResult() {
    resultContainer.innerHTML = "";
    lastExportRows = null;
    exportToolbar.hidden = true;
  }

  function pickField(row, canonicalKey) {
    const keys = LEGACY_KEY_MAP[canonicalKey];
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (Object.prototype.hasOwnProperty.call(row, k)) return row[k];
    }
    return null;
  }

  /** Remove leading row indices like "1. ", "2) " (space required after . or ) so "1.5mg" stays). */
  function stripLeadingRowIndexFromDescription(val) {
    if (val === null || val === undefined) return val;
    const s = String(val).trim();
    if (!s) return val;
    const stripped = s.replace(/^\d{1,4}[\.\)]\s+/u, "").trim();
    return stripped || s;
  }

  function normalizeRow(row) {
    const out = {};
    COLUMNS.forEach((key) => {
      let v = pickField(row, key);
      if (key === "DESCRIPTION") {
        v = stripLeadingRowIndexFromDescription(v);
      }
      out[key] = v;
    });
    return out;
  }

  function normalizeRows(rows) {
    return rows.map(normalizeRow);
  }

  function dedupeByDescription(rows) {
    const seen = new Set();
    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const raw = row["DESCRIPTION"];
      const empty =
        raw === null || raw === undefined || String(raw).trim() === "";
      if (empty) {
        out.push(row);
        continue;
      }
      const key = String(raw).trim().toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(row);
    }
    return out;
  }

  function sortAlphabeticallyByDescription(rows) {
    return [...rows].sort((a, b) => {
      const da = (a["DESCRIPTION"] ?? "").toString().trim();
      const db = (b["DESCRIPTION"] ?? "").toString().trim();
      const emptyA = da === "";
      const emptyB = db === "";
      if (emptyA && emptyB) return 0;
      if (emptyA) return 1;
      if (emptyB) return -1;
      return da.localeCompare(db, undefined, { sensitivity: "base" });
    });
  }

  function processRows(rows) {
    return sortAlphabeticallyByDescription(dedupeByDescription(rows));
  }

  function isBlankRow(row) {
    return COLUMNS.every((key) => {
      const v = row[key];
      return v === null || v === undefined || String(v).trim() === "";
    });
  }

  function formatPriceExcel(val) {
    if (val === null || val === undefined || val === "") return "";
    const raw = String(val)
      .replace(/,/g, "")
      .replace(/^\s*\$\s*/, "")
      .trim();
    const n = parseFloat(raw);
    if (Number.isNaN(n)) return String(val).trim();
    return n.toLocaleString("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  function formatDiscExcel(val) {
    if (val === null || val === undefined || val === "") return "";
    const s = String(val).trim();
    if (s === "") return "";
    const num = parseFloat(s.replace(/%/g, "").replace(/,/g, "").trim());
    if (!Number.isNaN(num)) return `${num}%`;
    return s;
  }

  function cellForDisplay(key, val) {
    if (val === null || val === undefined || String(val).trim() === "") return "";
    if (key === "PRICE") return formatPriceExcel(val);
    if (key === "DISC %") return formatDiscExcel(val);
    return String(val);
  }

  async function exportToXlsx() {
    if (!lastExportRows || lastExportRows.length === 0) return;

    exportButton.disabled = true;
    try {
      const res = await fetch(apiUrl("/export-xlsx"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rows: lastExportRows,
          baseName: lastExportBaseName,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Export failed (${res.status})`);
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const safeBase =
        lastExportBaseName.replace(/[^\w\-]+/g, "_").slice(0, 80) || "invoice";
      a.download = `${safeBase}-extracted.xlsx`;
      a.rel = "noopener";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      setStatus(e.message || "Export failed", true);
      statusEl.hidden = false;
    } finally {
      exportButton.disabled = false;
    }
  }

  function buildTable(rows) {
    const table = document.createElement("table");
    table.className = "data-table";

    const colgroup = document.createElement("colgroup");
    ["col-desc", "col-price", "col-deal", "col-disc"].forEach((cls) => {
      const col = document.createElement("col");
      col.className = cls;
      colgroup.appendChild(col);
    });
    table.appendChild(colgroup);

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    COLUMNS.forEach((h) => {
      const th = document.createElement("th");
      th.scope = "col";
      th.textContent = h;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.forEach((row) => {
      const tr = document.createElement("tr");
      if (isBlankRow(row)) {
        tr.className = "data-table__blank-row";
      }
      COLUMNS.forEach((key) => {
        const td = document.createElement("td");
        td.textContent = cellForDisplay(key, row[key]);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    return table;
  }

  exportButton.addEventListener("click", () => {
    exportToXlsx();
  });

  fileInput.addEventListener("change", async () => {
    const files = fileInput.files
      ? Array.from(fileInput.files).filter((f) => f.type.startsWith("image/"))
      : [];
    if (files.length === 0) return;

    clearResult();
    setPreview(files);
    setStatus(
      files.length === 1
        ? "Analyzing invoice…"
        : `Analyzing image 1 of ${files.length}…`,
      false
    );

    const allRows = [];

    try {
      for (let i = 0; i < files.length; i++) {
        if (files.length > 1) {
          setStatus(`Analyzing image ${i + 1} of ${files.length}…`, false);
        }

        const formData = new FormData();
        formData.append("image", files[i]);

        const res = await fetch(apiUrl("/analyze-invoice"), {
          method: "POST",
          body: formData,
        });

        const data = await res.json().catch(() => ({}));

        if (!res.ok) {
          const errText = data.error || `Request failed (${res.status})`;
          statusEl.hidden = true;
          showResultModal(
            "Analysis failed",
            files.length > 1
              ? `Could not analyze “${files[i].name}”: ${errText}`
              : errText
          );
          return;
        }

        const rows = data.rows;
        if (Array.isArray(rows) && rows.length > 0) {
          allRows.push(...rows);
        }
      }

      if (allRows.length === 0) {
        statusEl.hidden = true;
        resultContainer.innerHTML =
          '<p class="empty-message">No product rows were extracted.</p>';
        showResultModal(
          "No rows found",
          files.length > 1
            ? "No product rows could be extracted from any of the selected images. Try clearer scans or different files."
            : "The invoice was analyzed, but no product rows could be extracted from the image. Try a clearer scan or a different file."
        );
        return;
      }

      const normalized = processRows(normalizeRows(allRows));

      setStatus("", false);
      statusEl.hidden = true;
      lastExportRows = normalized;
      lastExportBaseName =
        files.length === 1
          ? (files[0].name && files[0].name.replace(/\.[^.]+$/, "")) || "invoice"
          : `invoice-${files.length}-images`;
      resultContainer.appendChild(buildTable(normalized));
      exportToolbar.hidden = false;

      const count = normalized.filter((r) => !isBlankRow(r)).length;
      const rowLabel = count === 1 ? "row" : "rows";
      const sourceHint =
        files.length > 1
          ? ` from ${files.length} images`
          : "";
      showResultModal(
        "Analysis complete",
        `Extracted ${count} product ${rowLabel}${sourceHint}. Review the table below or download the Excel file.`
      );
    } catch (e) {
      const msg = e.message || "Network error";
      statusEl.hidden = true;
      showResultModal("Analysis failed", msg);
    } finally {
      fileInput.value = "";
    }
  });
})();
