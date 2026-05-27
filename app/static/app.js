(function () {
  const drop = document.getElementById("drop");
  const fileInput = document.getElementById("file");
  const browseBtn = document.getElementById("browse");
  const status = document.getElementById("status");
  const statusText = document.getElementById("statusText");
  const result = document.getElementById("result");
  const errorBox = document.getElementById("error");
  const errorText = document.getElementById("errorText");
  const bankEl = document.getElementById("bank");
  const rowsEl = document.getElementById("rows");
  const extractedEl = document.getElementById("extracted");
  const downloadEl = document.getElementById("download");
  const againBtn = document.getElementById("again");
  const errorAgainBtn = document.getElementById("errorAgain");
  const bankSelect = document.getElementById("bankSelect");
  const hangingSelect = document.getElementById("hangingSelect");
  const dateMergedSelect = document.getElementById("dateMergedSelect");
  const columnMapper = document.getElementById("columnMapper");
  const previewTable = document.getElementById("previewTable");
  const mapperSort = document.getElementById("mapperSort");
  const mapperCancel = document.getElementById("mapperCancel");
  const mapperError = document.getElementById("mapperError");

  const ROLES = [
    { value: "ignore",     label: "Ignore" },
    { value: "date",       label: "Date" },
    { value: "remarks",    label: "Narration" },
    { value: "withdrawal", label: "Debit / Withdrawal" },
    { value: "deposit",    label: "Credit / Deposit" },
    { value: "balance",    label: "Balance" },
  ];
  const REQUIRED_ROLES = ["date", "remarks", "withdrawal", "deposit"];
  const ROLE_FROM_SUGGESTED = {
    col_date: "date",
    col_remarks: "remarks",
    col_withdrawal: "withdrawal",
    col_deposit: "deposit",
    col_balance: "balance",
  };

  let lastFile = null;
  // Persisted across renders so the user's edits survive validation re-renders.
  let colRoles = [];
  let dataStartRow = 0;

  function show(el) { el.classList.remove("hidden"); }
  function hide(el) { el.classList.add("hidden"); }

  function reset() {
    hide(status);
    hide(result);
    hide(errorBox);
    hide(columnMapper);
    show(drop);
    fileInput.value = "";
    lastFile = null;
    colRoles = [];
    dataStartRow = 0;
    previewTable.innerHTML = "";
    hide(mapperError);
  }

  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      fileInput.click();
    }
  });
  browseBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  ["dragenter", "dragover"].forEach((ev) => {
    drop.addEventListener(ev, (e) => {
      e.preventDefault();
      e.stopPropagation();
      drop.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((ev) => {
    drop.addEventListener(ev, (e) => {
      e.preventDefault();
      e.stopPropagation();
      drop.classList.remove("dragover");
    });
  });

  drop.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) startPreview(file);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files && fileInput.files[0]) startPreview(fileInput.files[0]);
  });

  againBtn.addEventListener("click", reset);
  errorAgainBtn.addEventListener("click", reset);
  mapperCancel.addEventListener("click", reset);
  mapperSort.addEventListener("click", submitSort);

  async function startPreview(file) {
    lastFile = file;
    hide(drop);
    hide(result);
    hide(errorBox);
    hide(columnMapper);
    statusText.textContent = `Reading ${file.name}…`;
    show(status);

    const fd = new FormData();
    fd.append("file", file);
    fd.append("bank", (bankSelect && bankSelect.value) || "auto");
    fd.append("date_merged", (dateMergedSelect && dateMergedSelect.value) || "no");
    fd.append("hanging", (hangingSelect && hangingSelect.value) || "no");

    try {
      const res = await fetch("/api/preview", { method: "POST", body: fd });
      if (!res.ok) {
        let msg = `HTTP ${res.status}`;
        try {
          const j = await res.json();
          if (j.error) msg = j.error;
        } catch (_) {}
        throw new Error(msg);
      }
      const j = await res.json();
      if (!j.rows || !j.rows.length) {
        throw new Error("No rows could be read from this file.");
      }
      renderColumnMapper(j);
    } catch (err) {
      hide(status);
      errorText.textContent = err && err.message ? err.message : String(err);
      show(errorBox);
    }
  }

  function renderColumnMapper(preview) {
    const nCols = preview.n_cols || (preview.rows[0] ? preview.rows[0].length : 0);
    const suggested = preview.suggested_columns || {};

    // Seed role assignments from server suggestions.
    colRoles = new Array(nCols).fill("ignore");
    Object.keys(ROLE_FROM_SUGGESTED).forEach((key) => {
      const idx = suggested[key];
      if (idx != null && idx >= 0 && idx < nCols) {
        colRoles[idx] = ROLE_FROM_SUGGESTED[key];
      }
    });

    // Default data start: one row after the suggested header, or row 0 if none.
    const sh = preview.suggested_header_row;
    dataStartRow = (sh != null && sh >= 0 && sh + 1 < preview.rows.length)
      ? sh + 1
      : 0;

    drawTable(preview.rows, nCols);

    let countEl = document.getElementById("previewRowCount");
    if (!countEl) {
      countEl = document.createElement("p");
      countEl.id = "previewRowCount";
      countEl.className = "preview-row-count";
      previewTable.parentNode.insertBefore(countEl, previewTable);
    }
    countEl.textContent = `${preview.rows.length} rows in file`;

    hide(status);
    show(columnMapper);
  }

  function drawTable(rows, nCols) {
    previewTable.innerHTML = "";

    // Role-select header row
    const thead = document.createElement("thead");
    const trRoles = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "corner";
    corner.textContent = "Data starts";
    trRoles.appendChild(corner);
    for (let c = 0; c < nCols; c++) {
      const th = document.createElement("th");
      const sel = document.createElement("select");
      sel.dataset.col = String(c);
      ROLES.forEach((r) => {
        const opt = document.createElement("option");
        opt.value = r.value;
        opt.textContent = r.label;
        if (colRoles[c] === r.value) opt.selected = true;
        sel.appendChild(opt);
      });
      sel.addEventListener("change", () => {
        colRoles[c] = sel.value;
        // If user picks a unique role that another column already has, clear the other.
        if (sel.value !== "ignore" && sel.value !== "balance") {
          for (let i = 0; i < colRoles.length; i++) {
            if (i !== c && colRoles[i] === sel.value) {
              colRoles[i] = "ignore";
            }
          }
          drawTable(rows, nCols);
        }
        // Balance is also single-assignment but cheap to enforce identically.
        if (sel.value === "balance") {
          for (let i = 0; i < colRoles.length; i++) {
            if (i !== c && colRoles[i] === "balance") colRoles[i] = "ignore";
          }
          drawTable(rows, nCols);
        }
        hide(mapperError);
      });
      th.appendChild(sel);
      trRoles.appendChild(th);
    }
    thead.appendChild(trRoles);
    previewTable.appendChild(thead);

    // Body: one row per preview row, with a "data starts here" radio on the left.
    const tbody = document.createElement("tbody");
    rows.forEach((row, rIdx) => {
      const tr = document.createElement("tr");
      if (rIdx < dataStartRow) tr.classList.add("preamble");
      const tdRadio = document.createElement("td");
      tdRadio.className = "radio-cell";
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = "dataStart";
      radio.value = String(rIdx);
      radio.checked = (rIdx === dataStartRow);
      radio.title = "Transaction data starts at this row";
      radio.addEventListener("change", () => {
        dataStartRow = rIdx;
        drawTable(rows, nCols);
      });
      tdRadio.appendChild(radio);
      tr.appendChild(tdRadio);
      for (let c = 0; c < nCols; c++) {
        const td = document.createElement("td");
        td.textContent = row[c] != null ? String(row[c]) : "";
        if (colRoles[c] && colRoles[c] !== "ignore") {
          td.classList.add("col-" + colRoles[c]);
        }
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    previewTable.appendChild(tbody);
  }

  function validateMapping() {
    const counts = {};
    colRoles.forEach((r) => { counts[r] = (counts[r] || 0) + 1; });
    const missing = REQUIRED_ROLES.filter((r) => !counts[r]);
    if (missing.length) {
      const pretty = missing.map((r) => {
        const found = ROLES.find((x) => x.value === r);
        return found ? found.label : r;
      });
      return `Please assign a column for: ${pretty.join(", ")}.`;
    }
    return null;
  }

  function columnIndexFor(role) {
    return colRoles.indexOf(role);
  }

  async function submitSort() {
    const err = validateMapping();
    if (err) {
      mapperError.textContent = err;
      show(mapperError);
      return;
    }
    hide(mapperError);

    statusText.textContent = `Sorting ${lastFile.name}…`;
    hide(columnMapper);
    show(status);

    const fd = new FormData();
    fd.append("file", lastFile);
    fd.append("bank", (bankSelect && bankSelect.value) || "auto");
    fd.append("date_merged", (dateMergedSelect && dateMergedSelect.value) || "no");
    fd.append("hanging", (hangingSelect && hangingSelect.value) || "no");
    fd.append("data_start_row", String(dataStartRow));
    fd.append("col_date", String(columnIndexFor("date")));
    fd.append("col_remarks", String(columnIndexFor("remarks")));
    fd.append("col_withdrawal", String(columnIndexFor("withdrawal")));
    fd.append("col_deposit", String(columnIndexFor("deposit")));
    const bal = columnIndexFor("balance");
    if (bal >= 0) fd.append("col_balance", String(bal));

    try {
      const res = await fetch("/api/sort", { method: "POST", body: fd });
      if (!res.ok) {
        let msg = `HTTP ${res.status}`;
        try {
          const j = await res.json();
          if (j.error) msg = j.error;
        } catch (_) {}
        throw new Error(msg);
      }

      const blob = await res.blob();
      const cd = res.headers.get("Content-Disposition") || "";
      const m = cd.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
      const filename = (m && decodeURIComponent(m[1])) || "sorted.xlsx";

      const url = URL.createObjectURL(blob);
      downloadEl.href = url;
      downloadEl.download = filename;

      bankEl.textContent = (res.headers.get("X-Bank-Profile") || "?").toUpperCase();
      rowsEl.textContent = res.headers.get("X-Row-Count") || "?";
      extractedEl.textContent = res.headers.get("X-Extracted") || "?";

      hide(status);
      show(result);
    } catch (err) {
      hide(status);
      errorText.textContent = err && err.message ? err.message : String(err);
      show(errorBox);
    }
  }
})();
