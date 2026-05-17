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

  function show(el) { el.classList.remove("hidden"); }
  function hide(el) { el.classList.add("hidden"); }

  function reset() {
    hide(status);
    hide(result);
    hide(errorBox);
    show(drop);
    fileInput.value = "";
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
    if (file) submit(file);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files && fileInput.files[0]) submit(fileInput.files[0]);
  });

  againBtn.addEventListener("click", reset);
  errorAgainBtn.addEventListener("click", reset);

  async function submit(file) {
    hide(drop);
    hide(result);
    hide(errorBox);
    statusText.textContent = `Processing ${file.name}…`;
    show(status);

    const fd = new FormData();
    fd.append("file", file);

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
