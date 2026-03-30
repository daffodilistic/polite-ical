const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const sheetControls = document.getElementById("sheetControls");
const mappingControls = document.getElementById("mappingControls");
const actions = document.getElementById("actions");
const convertButton = document.getElementById("convertButton");
const downloadArea = document.getElementById("downloadArea");
const resultMessage = document.getElementById("resultMessage");
const downloadLink = document.getElementById("downloadLink");
const previewArea = document.getElementById("previewArea");
const previewText = document.getElementById("previewText");
const logArea = document.getElementById("logArea");
const logOutput = document.getElementById("logOutput");

const summarySelect = document.getElementById("summarySelect");
const descriptionSelect = document.getElementById("descriptionSelect");
const startDateSelect = document.getElementById("startDateSelect");
const startTimeSelect = document.getElementById("startTimeSelect");
const endDateSelect = document.getElementById("endDateSelect");
const endTimeSelect = document.getElementById("endTimeSelect");
const locationSelect = document.getElementById("locationSelect");

let workbook = null;
let rows = [];
let headers = [];

function show(element) {
  element.classList.remove("hidden");
}

function hide(element) {
  element.classList.add("hidden");
}

function log(message) {
  logOutput.textContent = message;
  show(logArea);
}

function clearLog() {
  logOutput.textContent = "";
  hide(logArea);
}

function arrayBufferToUint8Array(buffer) {
  return new Uint8Array(buffer);
}

function normalizeText(text) {
  return String(text || "").trim();
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value)) {
    return value;
  }

  const trimmed = normalizeText(value);
  if (!trimmed) {
    return null;
  }

  const date = new Date(trimmed);
  if (!Number.isNaN(date.getTime())) {
    return date;
  }

  const excelDate = Number(trimmed);
  if (!Number.isNaN(excelDate)) {
    const parsed = XLSX.SSF.parse_date_code(excelDate, { date1904: false });
    if (parsed && parsed.y) {
      return new Date(
        parsed.y,
        parsed.m - 1,
        parsed.d,
        parsed.H || 0,
        parsed.M || 0,
        parsed.S || 0,
      );
    }
  }

  return null;
}

function toIcsDateTime(value) {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    return null;
  }
  const utc = new Date(value.getTime() - value.getTimezoneOffset() * 60000);
  const iso = utc.toISOString().replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z");
  return iso;
}

function foldIcsLine(line) {
  return line.replace(/(.{1,75})(?=.)/g, "$1\r\n ");
}

function buildIcs(items) {
  const nowStamp = toIcsDateTime(new Date());
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Polite ICS Converter//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
  ];

  items.forEach((item, index) => {
    const uid = item.uid || `polite-ics-${Date.now()}-${index}`;
    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${uid}`);
    lines.push(`DTSTAMP:${nowStamp}`);
    lines.push(`SUMMARY:${item.summary}`);
    if (item.description) {
      lines.push(`DESCRIPTION:${item.description}`);
    }
    if (item.location) {
      lines.push(`LOCATION:${item.location}`);
    }
    lines.push(`DTSTART:${item.start}`);
    lines.push(`DTEND:${item.end}`);
    lines.push("END:VEVENT");
  });

  lines.push("END:VCALENDAR");
  return lines.map(foldIcsLine).join("\r\n");
}

function fillSelect(select, options) {
  select.innerHTML = "";
  const emptyOption = document.createElement("option");
  emptyOption.value = "";
  emptyOption.textContent = "(none)";
  select.appendChild(emptyOption);

  options.forEach((option) => {
    const opt = document.createElement("option");
    opt.value = option;
    opt.textContent = option;
    select.appendChild(opt);
  });
}

function setMappingDefaults() {
  const guess = (pattern) => {
    const regex = new RegExp(pattern, "i");
    return headers.find((value) => regex.test(value));
  };

  summarySelect.value = guess("summary|title|event|subject|name") || "";
  descriptionSelect.value = guess("description|details|notes|comment") || "";
  locationSelect.value = guess("location|venue|room|place") || "";
  startDateSelect.value = guess("start.*date|date|day|attendance|log.*date") ||
    "";
  startTimeSelect.value = guess("start.*time|time.*in|in time|time") || "";
  endDateSelect.value = guess("end.*date|date|day") || startDateSelect.value ||
    "";
  endTimeSelect.value = guess("end.*time|time.*out|out time|finish|leave") ||
    "";
}

function buildRows(sheetName) {
  if (!workbook) {
    rows = [];
    return;
  }

  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
  });
  const headerRow = data[0] || [];
  headers = headerRow.map(normalizeText).filter((value) => value !== "");

  const bodyRows = data.slice(1).map((row) => {
    return headers.reduce((acc, header, index) => {
      acc[header] = normalizeText(row[index]);
      return acc;
    }, {});
  });

  rows = bodyRows;
  fillSelect(summarySelect, headers);
  fillSelect(descriptionSelect, headers);
  fillSelect(locationSelect, headers);
  fillSelect(startDateSelect, headers);
  fillSelect(startTimeSelect, ["", ...headers]);
  fillSelect(endDateSelect, headers);
  fillSelect(endTimeSelect, ["", ...headers]);
  setMappingDefaults();
}

function renderPreview(sheetName) {
  if (!rows.length) {
    previewText.textContent = "No rows found in the selected sheet.";
    show(previewArea);
    return;
  }

  const sampleRows = rows.slice(0, 5);
  const preview = sampleRows.map((row, rowIndex) => {
    const values = headers.map((header) => `${header}: ${row[header] || ""}`);
    return `Row ${rowIndex + 1}\n${values.join("\n")}`;
  }).join("\n\n");

  previewText.textContent = preview;
  show(previewArea);
}

function createEvents() {
  const startDateKey = startDateSelect.value;
  const startTimeKey = startTimeSelect.value;
  const endDateKey = endDateSelect.value || startDateKey;
  const endTimeKey = endTimeSelect.value;
  const summaryKey = summarySelect.value;
  const descriptionKey = descriptionSelect.value;
  const locationKey = locationSelect.value;

  if (!startDateKey) {
    throw new Error("Start date column is required.");
  }

  const items = rows.reduce((events, row, index) => {
    const summary = row[summaryKey] || row[descriptionKey] || "Calendar Event";
    const description = row[descriptionKey]
      ? row[descriptionKey]
      : Object.entries(row)
        .map(([key, value]) => `${key}: ${value}`)
        .filter((line) => line.trim() !== "")
        .join("\n");

    const startDateValue = row[startDateKey];
    const startTimeValue = row[startTimeKey];
    const endDateValue = row[endDateKey];
    const endTimeValue = row[endTimeKey];
    const location = row[locationKey] || "";

    const startDate = parseDateValue(startDateValue);
    if (!startDate) {
      return events;
    }

    let start = startDate;
    if (startTimeKey && startTimeValue) {
      const parsedTime = parseDateValue(`${startDateValue} ${startTimeValue}`);
      if (parsedTime) {
        start = parsedTime;
      }
    }

    let end = null;
    if (endTimeValue || endDateValue) {
      const endDate = parseDateValue(endDateValue || startDateValue);
      if (endDate) {
        end = endDate;
        if (endTimeValue) {
          const parsedEnd = parseDateValue(
            `${endDateValue || startDateValue} ${endTimeValue}`,
          );
          if (parsedEnd) {
            end = parsedEnd;
          }
        }
      }
    }

    if (!end) {
      end = new Date(start.getTime() + 60 * 60 * 1000);
    }

    const startText = toIcsDateTime(start);
    const endText = toIcsDateTime(end);
    if (!startText || !endText) {
      return events;
    }

    events.push({
      uid: `polite-${Date.now()}-${index}`,
      summary: summary.replace(/\r?\n/g, " "),
      description: description.replace(/\r?\n/g, "\\n"),
      location: location.replace(/\r?\n/g, " "),
      start: startText,
      end: endText,
    });
    return events;
  }, []);

  return items;
}

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  clearLog();
  hide(downloadArea);
  hide(previewArea);
  hide(mappingControls);
  hide(actions);

  try {
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data, { type: "array" });
    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      sheetSelect.appendChild(option);
    });
    show(sheetControls);
    show(mappingControls);
    show(actions);
    buildRows(workbook.SheetNames[0]);
    renderPreview(workbook.SheetNames[0]);
  } catch (error) {
    log(`Unable to parse XLSX file: ${error.message}`);
  }
});

sheetSelect.addEventListener("change", (event) => {
  const sheetName = event.target.value;
  buildRows(sheetName);
  renderPreview(sheetName);
  hide(downloadArea);
});

convertButton.addEventListener("click", () => {
  try {
    const items = createEvents();
    if (!items.length) {
      log("No valid events could be created from the selected rows.");
      return;
    }

    const icsContent = buildIcs(items);
    const blob = new Blob([icsContent], {
      type: "text/calendar;charset=utf-8",
    });
    const url = URL.createObjectURL(blob);
    downloadLink.href = url;
    resultMessage.textContent = `Created ${items.length} event${
      items.length === 1 ? "" : "s"
    }.`;
    show(downloadArea);
    clearLog();
  } catch (error) {
    log(error.message);
  }
});
