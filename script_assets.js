(async () => {
  /*
    1. TEST WITH A FEW SERIALS FIRST.
    2. Replace this list with your complete list once confirmed.
  */
  const serials = `
1HLI10SE0086
1RPE4634VK
`
    .split(/\r?\n/)
    .map(x => x.trim())
    .filter(Boolean);

  const SERIAL_ID =
    "ctl00_ContentPlaceHolder1_FilterPanel_i0_i0_txtFilterSerialNo";

  const SERIAL_STATE_ID =
    "ctl00_ContentPlaceHolder1_FilterPanel_i0_i0_txtFilterSerialNo_ClientState";

  const APPLY_ID =
    "ctl00_ContentPlaceHolder1_FilterPanel_i0_i0_btnFilterApply_input";

  const FILTER_CONFIRM_ID =
    "ContentPlaceHolder1_hidAssetGridFilter";

  const EXPORT_URL =
    "/download.ashx?type=xlsx&id=asset&function=material";

  const SEARCH_DELAY_MS = 800;
  const EXPORT_DELAY_MS = 1200;

  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));

  async function loadSheetJS() {
    if (window.XLSX) return window.XLSX;

    await new Promise((resolve, reject) => {
      const script = document.createElement("script");

      script.src =
        "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

      script.onload = resolve;
      script.onerror = () =>
        reject(
          new Error(
            "Could not load the Excel processing library. The work network may be blocking the SheetJS site."
          )
        );

      document.head.appendChild(script);
    });

    if (!window.XLSX) {
      throw new Error("Excel library did not load correctly.");
    }

    return window.XLSX;
  }

  const XLSX = await loadSheetJS();

  const form = document.querySelector("form");
  const serialInput = document.getElementById(SERIAL_ID);
  const serialState = document.getElementById(SERIAL_STATE_ID);
  const applyButton = document.getElementById(APPLY_ID);

  if (!form || !serialInput || !serialState || !applyButton) {
    throw new Error(
      "Could not find the AssetTagz Serial Number or Apply controls."
    );
  }

  const frame = document.createElement("iframe");
  frame.name = "assettagzAutomationFrame";
  frame.style.display = "none";
  document.body.appendChild(frame);

  const combinedRows = [];
  const summaryRows = [];

  function updateAspNetState(responseDocument) {
    for (const source of responseDocument.querySelectorAll(
      'input[type="hidden"][name]'
    )) {
      const currentPageInput = document.getElementsByName(source.name)[0];

      if (currentPageInput) {
        currentPageInput.value = source.value;
      }
    }
  }

  async function applySerial(serial) {
    serialInput.value = serial;

    serialInput.dispatchEvent(
      new Event("input", { bubbles: true })
    );

    serialInput.dispatchEvent(
      new Event("change", { bubbles: true })
    );

    serialState.value = JSON.stringify({
      enabled: true,
      emptyMessage: "",
      validationText: serial,
      valueAsString: serial,
      lastSetTextBoxValue: serial
    });

    const originalTarget = form.target;

    await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        form.target = originalTarget;
        reject(new Error(`Search timed out for ${serial}`));
      }, 30000);

      frame.onload = () => {
        clearTimeout(timeout);
        form.target = originalTarget;
        resolve();
      };

      form.target = frame.name;
      applyButton.click();
    });

    const responseDocument = frame.contentDocument;

    if (!responseDocument) {
      throw new Error(`Could not read the search response for ${serial}`);
    }

    updateAspNetState(responseDocument);

    const confirmedSerial =
      responseDocument.getElementById(SERIAL_ID)?.value?.trim();

    const appliedFilter =
      responseDocument.getElementById(FILTER_CONFIRM_ID)?.value || "";

    if (confirmedSerial !== serial || !appliedFilter.includes(serial)) {
      throw new Error(
        `AssetTagz did not confirm that ${serial} was applied as the filter.`
      );
    }
  }

  async function getExportWorkbook(serial) {
    const response = await fetch(EXPORT_URL, {
      credentials: "same-origin",
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(
        `Export failed for ${serial}: HTTP ${response.status}`
      );
    }

    const buffer = await response.arrayBuffer();
    const firstBytes = new Uint8Array(buffer.slice(0, 2));

    // XLSX files normally start with PK because they are ZIP files.
    if (firstBytes[0] !== 0x50 || firstBytes[1] !== 0x4b) {
      throw new Error(
        `The export response for ${serial} was not an Excel file.`
      );
    }

    return XLSX.read(buffer, {
      type: "array",
      cellDates: true
    });
  }

  function appendWorkbookRows(workbook, serial) {
    const sheetName = workbook.SheetNames[0];

    if (!sheetName) {
      return 0;
    }

    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false
    });

    for (const row of rows) {
      combinedRows.push({
        "Requested Serial Number": serial,
        ...row
      });
    }

    return rows.length;
  }

  console.clear();
  console.log(
    `Starting combined export for ${serials.length} serial numbers...`
  );

  for (let index = 0; index < serials.length; index++) {
    const serial = serials[index];

    console.log(
      `Processing ${index + 1}/${serials.length}: ${serial}`
    );

    try {
      await applySerial(serial);
      await wait(SEARCH_DELAY_MS);

      const workbook = await getExportWorkbook(serial);
      const rowCount = appendWorkbookRows(workbook, serial);

      summaryRows.push({
        "Requested Serial Number": serial,
        Status: rowCount ? "Exported" : "No rows returned",
        "Rows Added": rowCount,
        Error: ""
      });

      console.log(`✓ ${serial}: ${rowCount} row(s) added.`);
    } catch (error) {
      console.error(`✗ ${serial}:`, error);

      summaryRows.push({
        "Requested Serial Number": serial,
        Status: "Failed",
        "Rows Added": 0,
        Error: error.message || String(error)
      });
    }

    await wait(EXPORT_DELAY_MS);
  }

  if (!combinedRows.length) {
    throw new Error(
      "No data rows were collected. Check the Run Summary output in the Console."
    );
  }

  const allHeaders = [];

  for (const row of combinedRows) {
    for (const key of Object.keys(row)) {
      if (!allHeaders.includes(key)) {
        allHeaders.push(key);
      }
    }
  }

  const normalisedRows = combinedRows.map(row => {
    const result = {};

    for (const header of allHeaders) {
      result[header] = row[header] ?? "";
    }

    return result;
  });

  const finalWorkbook = XLSX.utils.book_new();

  const dataSheet = XLSX.utils.json_to_sheet(
    normalisedRows,
    { header: allHeaders }
  );

  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);

  dataSheet["!freeze"] = { xSplit: 0, ySplit: 1 };
  summarySheet["!freeze"] = { xSplit: 0, ySplit: 1 };

  XLSX.utils.book_append_sheet(
    finalWorkbook,
    dataSheet,
    "Combined Export"
  );

  XLSX.utils.book_append_sheet(
    finalWorkbook,
    summarySheet,
    "Run Summary"
  );

  const date = new Date().toISOString().slice(0, 10);

  XLSX.writeFile(
    finalWorkbook,
    `AssetTagz_Combined_Export_${date}.xlsx`
  );

  console.log("Finished.");
  console.table(summaryRows);
})();
