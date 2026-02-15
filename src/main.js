const { invoke } = window.__TAURI__.core;
const { open, save } = window.__TAURI__.dialog;

let selectedFilePath = null;
let selectedOutputPath = null;

document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("file-upload");
  if (fileInput) {
    fileInput.style.display = "none";
  }

  const selectFileBtn = document.getElementById("select-file-btn");
  const selectOutputBtn = document.getElementById("select-output-btn");
  const splitBtn = document.getElementById("split-btn");
  const selectedFileEl = document.getElementById("selected-file");
  const selectedOutputEl = document.getElementById("selected-output");
  const resultOutput = document.getElementById("result-output");

  function refreshRunButtonState() {
    const canRun = Boolean(selectedFilePath) && Boolean(selectedOutputPath);
    if (splitBtn) {
      splitBtn.disabled = !canRun;
    }
  }

  function getPickedPath(value) {
    if (!value) return null;
    return Array.isArray(value) ? value[0] : value;
  }

  if (selectFileBtn) {
    selectFileBtn.addEventListener("click", async () => {
      const picked = await open({
        multiple: false,
        filters: [{ name: "Excel Files", extensions: ["xlsx"] }],
      });
      const path = getPickedPath(picked);
      selectedFilePath = path || null;
      if (selectedFileEl) {
        selectedFileEl.textContent = selectedFilePath || "No file selected";
      }
      refreshRunButtonState();
    });
  }

  if (selectOutputBtn) {
    selectOutputBtn.addEventListener("click", async () => {
      const picked = await save({
        defaultPath: "orders_by_sheet.xlsx",
        filters: [{ name: "Excel Files", extensions: ["xlsx"] }],
      });
      const path = getPickedPath(picked);
      selectedOutputPath = path || null;
      if (selectedOutputEl) {
        selectedOutputEl.textContent = selectedOutputPath || "No output file selected";
      }
      refreshRunButtonState();
    });
  }

  if (splitBtn) {
    splitBtn.addEventListener("click", async () => {
      if (!selectedFilePath || !selectedOutputPath) {
        alert("Select input file and output file first.");
        return;
      }

      splitBtn.disabled = true;
      if (resultOutput) {
        resultOutput.textContent = "Processing...";
      }

      try {
        const payload = {
          file_path: selectedFilePath,
          filePath: selectedFilePath,
          output_path: selectedOutputPath,
          outputPath: selectedOutputPath,
        };
        const result = await invoke("split_xlsx_by_order", payload);
        if (resultOutput) {
          resultOutput.textContent = [
            `Status: ${result.status}`,
            `Output file: ${result.output_path}`,
            `Sheets created: ${result.sheets_created}`,
            `Rows exported: ${result.rows_exported}`,
            `Rows skipped (invalid G): ${result.skipped_invalid_rows}`,
          ].join("\n");
        }
      } catch (error) {
        if (resultOutput) {
          resultOutput.textContent = `Error: ${JSON.stringify(error)}`;
        }
      } finally {
        refreshRunButtonState();
      }
    });
  }
});
