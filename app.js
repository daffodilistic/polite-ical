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

const summaryInput = document.getElementById("summaryInput");
const summaryPreview = document.getElementById("summaryPreview");
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

// Naive implementation, but input is assumed to be machine-generated
function parseDateTimeValue(dateString) {
  // console.log("Parsing date string: ", dateString);

  const dateParts = dateString.split(" ");
  const eventDate = dateParts[0].split("/");
  const eventTime = dateParts[1].split(":");

  // console.log("eventDate", eventDate)
  // console.log("eventTime", eventTime)

  parsedDateTime = new Date(
    Number(eventDate[2]), // Year
    Number(eventDate[1]) - 1, // Month (0-based)
    Number(eventDate[0]), // Day
    Number(eventTime[0] || 0), // Hours
    Number(eventTime[1] || 0), // Minutes
    0, // Seconds
  );

  // console.log("Parsed DateTime", parsedDateTime);

  return parsedDateTime;
}

// NOTE we are assuming all inputs to be SG time (GMT+8)
function toIcsDateTime(value) {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
    return null;
  }
  const iso = value
    .toISOString()
    .replace(/[-:]/g, "") // Remove delimiters/markers
    .replace(/\.\d+Z$/, "Z"); // Remove milliseconds
  return iso;
}

function foldIcsLine(line) {
  return line.replace(/(.{1,75})(?=.*)/g, "$1\r\n");
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
    lines.push(`SUMMARY:${item.summary} - Session ${item.session}`); // Used as event title in GCal
    if (item.description) {
      let description = `DESCRIPTION:${item.description}`;
      let formatted_description = [];

      const size = 75;
      // Remove first 75 characters
      formatted_description.push(description.slice(0, size));
      // Check for extras
      if (description.length > size) {
        let start_index = size;
        do {
          output = " " + description.slice(start_index, start_index + size - 1);
          formatted_description.push(output);
          start_index += size - 1;
        } while (start_index < description.length);
      }

      // console.log(description);
      // console.log(formatted_description);

      lines.push(formatted_description.join(""));
    }
    // Default value for now (lat;long)
    // NOTE does not work in Gcal
    // lines.push(`GEO:1.3097757;103.7775495`);
    if (item.location) {
      lines.push(`LOCATION:Singapore Polytechnic`);
    } else {
      lines.push(`LOCATION:Singapore Polytechnic`);
    }
    lines.push(`DTSTART:${item.start}`);
    lines.push(`DTEND:${item.end}`);
    lines.push("END:VEVENT");
  });

  lines.push("END:VCALENDAR");
  return lines.map(foldIcsLine).join("");
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

  summaryInput.value = "{Module code} Learning to Read [{Session code}/{Total sessions}]";
  descriptionSelect.value = guess("description|details|notes|comment") || "";
  locationSelect.value = guess("location|venue|room|place|facility") || "";
  startDateSelect.value = guess("start.*time|time.*in|in time|time") || "";
  startTimeSelect.value = guess("start.*time|time.*in|in time|time") || "";
  endDateSelect.value =
    guess("end.*time|time.*out|out time|finish|leave") || "";
  endTimeSelect.value =
    guess("end.*time|time.*out|out time|finish|leave") || "";
  updateSummaryPreview();
}

function renderSummaryTemplate(template, row, totals = {}) {
  if (!template) {
    return "";
  }

  const value = template.trim();
  if (headers.includes(value)) {
    return row[value] == null ? "" : String(row[value]);
  }

  return value.replace(/\{([^}]+)\}/g, (match, token) => {
    const key = token.trim();
    if (/^total\s+sessions$/i.test(key)) {
      return String(totals.totalSessions || 0);
    }
    return row[key] == null ? "" : String(row[key]);
  });
}

function updateSummaryPreview() {
  if (!rows.length) {
    summaryPreview.textContent = "";
    return;
  }

  const value = summaryInput.value.trim();
  const isHeader = headers.includes(value);
  const sampleSummary = renderSummaryTemplate(value, rows[0] || {}, {
    totalSessions: rows.length,
  });
  const isTemplate = /\{[^}]+\}/.test(value);

  if (value && isHeader) {
    summaryPreview.innerHTML = `Using column "${value}" → <b>${sampleSummary || "(blank)"}</b>`;
  } else if (value && isTemplate) {
    summaryPreview.innerHTML = `Template preview → <b>${sampleSummary || "(blank)"}</b>`;
  } else if (value) {
    summaryPreview.innerHTML = `Using fixed summary text: <b>"${value}"</b>`;
  } else {
    summaryPreview.textContent =
      "Enter a summary column name or template to preview event titles.";
  }
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

  // console.log("Raw XLS data:");
  // console.log(rows);

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
  const preview = sampleRows
    .map((row, rowIndex) => {
      const values = headers.map((header) => `${header}: ${row[header] || ""}`);
      return `Row ${rowIndex + 1}\n${values.join("\n")}`;
    })
    .join("\n\n");

  previewText.textContent = preview;
  show(previewArea);
}

function createEvents() {
  // console.log("[createEvents]", rows);

  const startDateKey = startDateSelect.value;
  const startTimeKey = startTimeSelect.value;
  const endDateKey = endDateSelect.value || startDateKey;
  const endTimeKey = endTimeSelect.value;
  const summaryInputValue = summaryInput.value.trim();
  const descriptionKey = descriptionSelect.value;
  const locationKey = locationSelect.value;

  if (!startDateKey) {
    throw new Error("Start date column is required.");
  }

  const items = rows.reduce((parsedEvents, row, index) => {
    // console.log("Parsing current row: ", row);

    if (row["Module code"] == false) return parsedEvents;

    const summaryFromTemplate = renderSummaryTemplate(summaryInputValue, row, {
      totalSessions: rows.length,
    });
    const summary = summaryFromTemplate || row[descriptionKey] || "Calendar Event";
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

    const startDate = parseDateTimeValue(startDateValue);
    // if (!startDate) {
    //   console.log("Invalid start date!", row);
    //   return row;
    // }

    let start = startDate;
    if (startTimeKey && startTimeValue) {
      const parsedTime = parseDateTimeValue(
        `${startDateValue} ${startTimeValue}`,
      );
      if (parsedTime) {
        start = parsedTime;
      }
    }

    let end = null;
    if (endTimeValue || endDateValue) {
      const endDate = parseDateTimeValue(endDateValue || startDateValue);
      if (endDate) {
        end = endDate;
        if (endTimeValue) {
          const parsedEnd = parseDateTimeValue(
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
      console.log("Invalid datetime conversions!", parsedEvents);
    }

    parsedEvents.push({
      uid: `polite-${Date.now()}-${index}`,
      summary: summary.replace(/\r?\n/g, " "),
      session: row["Session code"],
      description: description.replace(/\r?\n/g, "\\n"),
      location: location.replace(/\r?\n/g, " "),
      start: startText,
      end: endText,
    });
    return parsedEvents;
  }, []);

  // console.log("Parsed rows", items);
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
    updateSummaryPreview();
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

summaryInput.addEventListener("input", updateSummaryPreview);

convertButton.addEventListener("click", () => {
  try {
    const items = createEvents();

    // console.log("[convertButton.addEventListener]", items);

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
    console.error(error);
    log(error.message);
  }
});
