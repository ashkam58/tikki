const state = {
  area: "",
  visited: "",
  date: "",
  shift: "",
  entries: [],
  editingIndex: null,
};

const refs = {
  areaInput: document.getElementById("areaInput"),
  visitedInput: document.getElementById("visitedInput"),
  dateInput: document.getElementById("dateInput"),
  shiftInput: document.getElementById("shiftInput"),
  locationInput: document.getElementById("locationInput"),
  observationInput: document.getElementById("observationInput"),
  suggestedInput: document.getElementById("suggestedInput"),
  informInput: document.getElementById("informInput"),
  pictureInput: document.getElementById("pictureInput"),
  captionInput: document.getElementById("captionInput"),
  addRowBtn: document.getElementById("addRowBtn"),
  clearRowBtn: document.getElementById("clearRowBtn"),
  downloadBtn: document.getElementById("downloadBtn"),
  resetReportBtn: document.getElementById("resetReportBtn"),
  entrySummary: document.getElementById("entrySummary"),
  entryList: document.getElementById("entryList"),
  reportBody: document.getElementById("reportBody"),
  previewArea: document.getElementById("previewArea"),
  previewVisited: document.getElementById("previewVisited"),
  previewDate: document.getElementById("previewDate"),
  previewShift: document.getElementById("previewShift"),
  entryForm: document.getElementById("entry-form"),
  // Bulk upload elements - COMMENTED OUT (elements are commented in HTML)
  /*
  bulkDataInput: document.getElementById("bulkDataInput"),
  bulkFormatSelect: document.getElementById("bulkFormatSelect"),
  customSeparatorInput: document.getElementById("customSeparatorInput"),
  customSeparatorLabel: document.getElementById("customSeparatorLabel"),
  previewBulkBtn: document.getElementById("previewBulkBtn"),
  addBulkBtn: document.getElementById("addBulkBtn"),
  clearBulkBtn: document.getElementById("clearBulkBtn"),
  bulkPreview: document.getElementById("bulkPreview"),
  bulkPreviewTable: document.getElementById("bulkPreviewTable"),
  confirmBulkBtn: document.getElementById("confirmBulkBtn"),
  cancelBulkBtn: document.getElementById("cancelBulkBtn"),
  */
  // Column-specific bulk upload elements
  columnSelect: document.getElementById("columnSelect"),
  columnModeSelect: document.getElementById("columnModeSelect"),
  columnDataInput: document.getElementById("columnDataInput"),
  previewColumnBtn: document.getElementById("previewColumnBtn"),
  applyColumnBtn: document.getElementById("applyColumnBtn"),
  clearColumnBtn: document.getElementById("clearColumnBtn"),
  columnPreview: document.getElementById("columnPreview"),
  columnPreviewTable: document.getElementById("columnPreviewTable"),
  confirmColumnBtn: document.getElementById("confirmColumnBtn"),
  cancelColumnBtn: document.getElementById("cancelColumnBtn"),
  fillTestDataBtn: document.getElementById("fillTestDataBtn"),
};

const DOCX_HTML_CHUNK_PATH = "word/chunk.xhtml";

init();

function init() {
  hydrateDefaultDate();
  attachMetaListeners();
  refs.addRowBtn.addEventListener("click", handleAddRow);
  refs.clearRowBtn.addEventListener("click", handleClearEntryForm);
  refs.downloadBtn.addEventListener("click", handleDownloadPdf);
  refs.resetReportBtn.addEventListener("click", handleResetReport);
  if (refs.fillTestDataBtn) {
    refs.fillTestDataBtn.addEventListener("click", handleFillTestData);
  }
  // DOCX button (may not exist in older HTML until added)
  refs.downloadDocxBtn = document.getElementById("downloadDocxBtn");
  if (refs.downloadDocxBtn) {
    refs.downloadDocxBtn.addEventListener("click", handleDownloadDocx);
  }
  
  // Bulk upload event listeners - COMMENTED OUT (elements are commented in HTML)
  /*
  refs.bulkFormatSelect.addEventListener("change", handleFormatChange);
  refs.previewBulkBtn.addEventListener("click", handlePreviewBulk);
  refs.addBulkBtn.addEventListener("click", handleAddBulk);
  refs.clearBulkBtn.addEventListener("click", handleClearBulk);
  refs.confirmBulkBtn.addEventListener("click", handleConfirmBulk);
  refs.cancelBulkBtn.addEventListener("click", handleCancelBulk);
  */
  
  // Column-specific bulk upload event listeners
  refs.columnSelect.addEventListener("change", handleColumnSelectChange);
  refs.previewColumnBtn.addEventListener("click", handlePreviewColumn);
  refs.applyColumnBtn.addEventListener("click", handleApplyColumn);
  refs.clearColumnBtn.addEventListener("click", handleClearColumn);
  refs.confirmColumnBtn.addEventListener("click", handleConfirmColumn);
  refs.cancelColumnBtn.addEventListener("click", handleCancelColumn);
  
  renderHeader();
  renderEntries();
  renderEntrySummary();
  initThemeSelector();
}

// Theme selector: apply and persist
function initThemeSelector() {
  const themeSelect = document.getElementById('themeSelect');
  const report = document.getElementById('report');
  if (!themeSelect || !report) return;

  // Restore saved theme
  const saved = localStorage.getItem('tikki_report_theme') || 'default';
  themeSelect.value = saved;
  applyTheme(saved);

  themeSelect.addEventListener('change', (e) => {
    const t = e.target.value || 'default';
    applyTheme(t);
    localStorage.setItem('tikki_report_theme', t);
  });
}

function applyTheme(themeName) {
  const report = document.getElementById('report');
  if (!report) return;
  // remove any theme-... classes we know about
  const themes = ['default','modern','minimal','compact','corporate','slate','green','mono'];
  themes.forEach(t => report.classList.remove(`theme-${t}`));
  const cls = `theme-${themeName || 'default'}`;
  report.classList.add(cls);
}

function hydrateDefaultDate() {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");
  const iso = `${yyyy}-${mm}-${dd}`;
  refs.dateInput.value = iso;
  state.date = iso;
}

function attachMetaListeners() {
  refs.areaInput.addEventListener("input", (event) => {
    state.area = event.target.value.trim();
    renderHeader();
  });

  refs.visitedInput.addEventListener("input", (event) => {
    state.visited = event.target.value.trim();
    renderHeader();
  });

  refs.dateInput.addEventListener("input", (event) => {
    state.date = event.target.value;
    renderHeader();
  });

  refs.shiftInput.addEventListener("input", (event) => {
    state.shift = event.target.value.trim();
    renderHeader();
  });
}

function renderHeader() {
  refs.previewArea.textContent = state.area || "\u00A0";
  refs.previewVisited.textContent = state.visited || "\u00A0";
  refs.previewDate.textContent = formatDateForDisplay(state.date) || "\u00A0";
  refs.previewShift.textContent = state.shift ? `Shift- ${state.shift}` : "\u00A0";
}

function formatDateForDisplay(isoDate) {
  if (!isoDate) return "";
  const [year, month, day] = isoDate.split("-");
  if (year && month && day) {
    const formattedDay = String(parseInt(day, 10));
    const formattedMonth = String(parseInt(month, 10));
    return `${formattedDay}-${formattedMonth}-${year}`;
  }
  return isoDate;
}

async function handleAddRow() {
  const baseEntry = collectEntryInputs();
  if (!baseEntry) {
    return;
  }

  try {
    const nextEntry = await buildEntryWithImage(baseEntry);
    if (state.editingIndex !== null) {
      state.entries[state.editingIndex] = {
        ...state.entries[state.editingIndex],
        ...nextEntry,
      };
      state.editingIndex = null;
    } else {
      state.entries.push({
        ...nextEntry,
        id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`,
      });
    }
  } catch (error) {
    console.error(error);
    alert("Something went wrong while reading the image. Please try again.");
    return;
  }

  refs.entryForm.reset();
  refs.captionInput.value = "";
  clearFileInput(refs.pictureInput);
  updateAddButtonLabel();
  renderEntries();
  renderEntrySummary();
}

function collectEntryInputs() {
  const location = refs.locationInput.value.trim();
  const observation = refs.observationInput.value.trim();
  const suggested = refs.suggestedInput.value.trim();
  const inform = refs.informInput.value.trim();
  const caption = refs.captionInput.value.trim();

  if (!location || !observation || !suggested) {
    alert("Location, Observation, and Suggested Action are required.");
    return null;
  }

  return { location, observation, suggested, inform, caption };
}

async function buildEntryWithImage(entry) {
  const file = refs.pictureInput.files[0];
  const imageData = file
    ? await readFileAsDataURL(file)
    : state.editingIndex !== null
    ? state.entries[state.editingIndex].imageData || null
    : null;

  return {
    ...entry,
    imageData,
  };
}

function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Unable to read file"));
    reader.readAsDataURL(file);
  });
}

function clearFileInput(input) {
  input.value = "";
}

function renderEntries() {
  refs.reportBody.innerHTML = "";

  state.entries.forEach((entry, index) => {
    const row = document.createElement("tr");

    const snoCell = document.createElement("td");
    snoCell.className = "sno-cell";
    snoCell.dataset.label = "S.NO";
    snoCell.textContent = `${index + 1}.`;

    const controls = createRowControls(index);
    if (controls) {
      snoCell.appendChild(controls);
    }

    const pictureCell = document.createElement("td");
    pictureCell.className = "picture-cell";
    pictureCell.dataset.label = "Picture";
    const pictureWrapper = document.createElement("div");
    pictureWrapper.className = "picture-wrapper";

    if (entry.imageData || entry.picture) {
      const image = document.createElement("img");
      image.src = entry.imageData || entry.picture;
      image.alt = `Row ${index + 1} picture`;
      pictureWrapper.appendChild(image);
    } else {
      pictureWrapper.textContent = "No image";
    }

    pictureCell.appendChild(pictureWrapper);

    // Add upload button for each picture cell
    const uploadContainer = document.createElement("div");
    uploadContainer.className = "picture-upload-container";
    
    const uploadBtn = document.createElement("button");
    uploadBtn.type = "button";
    uploadBtn.className = "picture-upload-btn";
    uploadBtn.textContent = entry.imageData || entry.picture ? "Change" : "Upload";
    uploadBtn.onclick = () => handlePictureUpload(index);
    
    uploadContainer.appendChild(uploadBtn);
    pictureCell.appendChild(uploadContainer);

    if (entry.caption) {
      const caption = document.createElement("div");
      caption.className = "picture-caption";
      caption.textContent = entry.caption;
      pictureCell.appendChild(caption);
    }

    const locationCell = document.createElement("td");
    locationCell.className = "location-cell text-cell";
    locationCell.dataset.label = "Location";
    locationCell.textContent = entry.location;

    const observationCell = document.createElement("td");
    observationCell.className = "observation-cell text-cell";
    observationCell.dataset.label = "Observation";
    observationCell.textContent = entry.observation;

    const suggestedCell = document.createElement("td");
    suggestedCell.className = "suggested-cell text-cell";
    suggestedCell.dataset.label = "Suggested Action";
    suggestedCell.textContent = entry.suggested;

    const informCell = document.createElement("td");
    informCell.className = "inform-cell text-cell";
    informCell.dataset.label = "Inform To";
    informCell.textContent = entry.inform || "";

    row.appendChild(snoCell);
    row.appendChild(pictureCell);
    row.appendChild(locationCell);
    row.appendChild(observationCell);
    row.appendChild(suggestedCell);
    row.appendChild(informCell);

    refs.reportBody.appendChild(row);
  });
}

function createRowControls(index) {
  if (state.entries.length === 0) return null;

  const wrapper = document.createElement("div");
  wrapper.className = "non-print";

  const editBtn = document.createElement("button");
  editBtn.type = "button";
  editBtn.className = "edit";
  editBtn.textContent = "Edit";
  editBtn.addEventListener("click", () => startEditEntry(index));

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "remove";
  removeBtn.textContent = "Delete";
  removeBtn.addEventListener("click", () => removeEntry(index));

  wrapper.appendChild(editBtn);
  wrapper.appendChild(removeBtn);

  return wrapper;
}

function renderEntrySummary() {
  if (!state.entries.length) {
    refs.entrySummary.hidden = true;
    refs.entryList.innerHTML = "";
    return;
  }

  refs.entrySummary.hidden = false;
  refs.entryList.innerHTML = "";

  state.entries.forEach((entry, index) => {
    const item = document.createElement("li");
    const summaryText = document.createElement("span");
    summaryText.textContent = `${index + 1}. ${entry.location.slice(0, 60)}${entry.location.length > 60 ? "..." : ""}`;

    const actions = document.createElement("span");
    actions.className = "row-actions";

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "edit";
    editBtn.textContent = "Edit";
    editBtn.addEventListener("click", () => startEditEntry(index));

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "remove";
    removeBtn.textContent = "Delete";
    removeBtn.addEventListener("click", () => removeEntry(index));

    actions.appendChild(editBtn);
    actions.appendChild(removeBtn);

    item.appendChild(summaryText);
    item.appendChild(actions);

    refs.entryList.appendChild(item);
  });
}

function startEditEntry(index) {
  const entry = state.entries[index];
  refs.locationInput.value = entry.location;
  refs.observationInput.value = entry.observation;
  refs.suggestedInput.value = entry.suggested;
  refs.informInput.value = entry.inform || "";
  refs.captionInput.value = entry.caption || "";
  clearFileInput(refs.pictureInput);
  state.editingIndex = index;
  updateAddButtonLabel();
  refs.locationInput.focus();
}

function removeEntry(index) {
  state.entries.splice(index, 1);
  if (state.editingIndex === index) {
    state.editingIndex = null;
    refs.entryForm.reset();
    updateAddButtonLabel();
  }
  renderEntries();
  renderEntrySummary();
}

function updateAddButtonLabel() {
  refs.addRowBtn.textContent = state.editingIndex !== null ? "Update Row" : "Add Row";
}

function handleClearEntryForm() {
  refs.entryForm.reset();
  state.editingIndex = null;
  updateAddButtonLabel();
}

function handleResetReport() {
  const confirmed = window.confirm("Clear all report data?");
  if (!confirmed) return;

  state.area = "";
  state.date = "";
  state.shift = "";
  state.entries = [];
  state.editingIndex = null;

  refs.entryForm.reset();
  refs.areaInput.value = "";
  refs.shiftInput.value = "";
  hydrateDefaultDate();
  renderHeader();
  renderEntries();
  renderEntrySummary();
  updateAddButtonLabel();
}

function handleFillTestData() {
  const confirmed = window.confirm("Fill the report with sample test data? This will replace existing entries.");
  if (!confirmed) return;

  // Clear existing entries
  state.entries = [];

  // Sample inspection data
  const testData = [
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-1`,
      location: "LMV Crossing-01",
      observation: "Haul road lights malfunctioning during night shift",
      suggested: "Repair or replace the lights immediately",
      inform: "Maintenance Team",
      caption: "Traffic light inspection",
      imageData: "https://picsum.photos/seed/1/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-2`,
      location: "HMV Parking Area",
      observation: "Diesel bowser parked in non-ready position",
      suggested: "Ensure proper ready-to-go parking procedures",
      inform: "Safety Officer",
      caption: "Diesel bowser positioning",
      imageData: "https://picsum.photos/seed/2/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-3`,
      location: "Quarry Site A",
      observation: "Safety equipment missing for workforce",
      suggested: "Install proper PPE and safety gear",
      inform: "Safety Department",
      caption: "PPE audit",
      imageData: "https://picsum.photos/seed/3/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-4`,
      location: "Site B Loading Area",
      observation: "Dust accumulation excessive during operations",
      suggested: "Implement dust suppression measures",
      inform: "Environmental Team",
      caption: "Dust at loading area",
      imageData: "https://picsum.photos/seed/4/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-5`,
      location: "Ghato Tand Public Crossing",
      observation: "Public crossing obstructed by equipment",
      suggested: "Create clear pedestrian path",
      inform: "Community Relations",
      caption: "Obstruction photo",
      imageData: "https://picsum.photos/seed/5/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-6`,
      location: "Kedla Excavation Edge",
      observation: "Excavation edge unsecured and dangerous",
      suggested: "Install safety barricades immediately",
      inform: "Site Supervisor",
      caption: "Unsecured edge",
      imageData: "https://picsum.photos/seed/6/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-7`,
      location: "Crusher Area Floor",
      observation: "Loose electrical cables on floor",
      suggested: "Secure cables and reroute properly",
      inform: "Electrical Team",
      caption: "Cables near crusher",
      imageData: "https://picsum.photos/seed/7/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-8`,
      location: "Weighbridge Station",
      observation: "Weighbridge not calibrated for 3 months",
      suggested: "Schedule immediate calibration",
      inform: "Quality Control",
      caption: "Calibration sticker missing",
      imageData: "https://picsum.photos/seed/8/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-9`,
      location: "Main Gatehouse",
      observation: "Unauthorized entry observed during shift change",
      suggested: "Improve gate control procedures",
      inform: "Security Department",
      caption: "Entry point security",
      imageData: "https://picsum.photos/seed/9/400/300"
    },
    {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-10`,
      location: "Material Stockyard",
      observation: "Material spillage near storage stacks",
      suggested: "Clear spill and retrain handling staff",
      inform: "Operations Manager",
      caption: "Spillage near stack",
      imageData: "https://picsum.photos/seed/10/400/300"
    }
  ];

  // Fill header with sample data if empty
  if (!state.area) {
    state.area = "QUARRY-AB";
    refs.areaInput.value = state.area;
  }
  if (!state.visited) {
    state.visited = "Rakesh Singh";
    refs.visitedInput.value = state.visited;
  }
  if (!state.shift) {
    state.shift = "B";
    refs.shiftInput.value = state.shift;
  }

  // Add test entries to state
  state.entries = testData;

  // Re-render everything
  renderHeader();
  renderEntries();
  renderEntrySummary();
  
  // Show success message
  alert(`Successfully filled ${testData.length} test entries!`);
}

// Bulk Upload Functions
let bulkData = [];

function handleFormatChange() {
  const format = refs.bulkFormatSelect.value;
  if (format === "custom") {
    refs.customSeparatorLabel.style.display = "block";
  } else {
    refs.customSeparatorLabel.style.display = "none";
  }
}

function handlePreviewBulk() {
  const rawData = refs.bulkDataInput.value.trim();
  if (!rawData) {
    showBulkMessage("Please enter some data to preview.", "error");
    return;
  }

  try {
    bulkData = parseBulkData(rawData);
    if (bulkData.length === 0) {
      showBulkMessage("No valid data found. Please check your format.", "error");
      return;
    }

    renderBulkPreview(bulkData);
    refs.bulkPreview.hidden = false;
    refs.addBulkBtn.disabled = false;
    showBulkMessage(`${bulkData.length} rows parsed successfully!`, "success");
  } catch (error) {
    showBulkMessage(`Error parsing data: ${error.message}`, "error");
  }
}

function parseBulkData(rawData) {
  const format = refs.bulkFormatSelect.value;
  let separator = ",";
  
  if (format === "tab") {
    separator = "\t";
  } else if (format === "custom") {
    separator = refs.customSeparatorInput.value || ",";
  }

  const lines = rawData.split("\n").filter(line => line.trim());
  const parsedData = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const columns = line.split(separator).map(col => col.trim());
    
    if (columns.length < 2) {
      throw new Error(`Line ${i + 1}: At least Location and Observation are required`);
    }

    const entry = {
      location: columns[0] || "",
      observation: columns[1] || "",
      suggested: columns[2] || "",
      inform: columns[3] || "",
      pictureUrl: columns[4] || "",  // Picture URL or file path
      caption: columns[5] || "",     // Picture caption
      picture: null // Will be processed if pictureUrl is provided
    };

    // If picture URL is provided, we'll handle it as a URL-based image
    if (entry.pictureUrl) {
      // Validate if it's a valid URL or file path
      try {
        if (entry.pictureUrl.startsWith('http') || entry.pictureUrl.startsWith('data:')) {
          entry.picture = entry.pictureUrl;
        } else {
          // For local file paths, we'll keep the reference but note it won't work in browser
          entry.picture = entry.pictureUrl;
        }
      } catch (error) {
        console.warn(`Line ${i + 1}: Invalid picture URL/path: ${entry.pictureUrl}`);
      }
    }

    parsedData.push(entry);
  }

  return parsedData;
}

function renderBulkPreview(data) {
  const table = document.createElement("table");
  table.className = "bulk-preview-table";

  // Create header
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  ["S.NO", "PICTURE", "LOCATION", "OBSERVATION", "SUGGESTED ACTION", "INFORM TO"].forEach(header => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Create body
  const tbody = document.createElement("tbody");
  data.forEach((entry, index) => {
    const row = document.createElement("tr");
    
    // S.NO (auto-generated)
    const snoCell = document.createElement("td");
    snoCell.textContent = (state.entries.length + index + 1).toString();
    snoCell.style.fontWeight = "bold";
    snoCell.style.textAlign = "center";
    row.appendChild(snoCell);

    // Picture column
    const pictureCell = document.createElement("td");
    if (entry.picture || entry.pictureUrl) {
      if (entry.picture && (entry.picture.startsWith('http') || entry.picture.startsWith('data:'))) {
        // Show a small preview if it's a valid URL
        const img = document.createElement("img");
        img.src = entry.picture;
        img.style.maxWidth = "60px";
        img.style.maxHeight = "40px";
        img.style.objectFit = "cover";
        img.onerror = () => {
          img.style.display = "none";
          pictureCell.textContent = "📷 " + (entry.caption || "Invalid URL");
        };
        pictureCell.appendChild(img);
        if (entry.caption) {
          const caption = document.createElement("div");
          caption.textContent = entry.caption;
          caption.style.fontSize = "0.7rem";
          caption.style.color = "#666";
          pictureCell.appendChild(caption);
        }
      } else {
        // Show file path/URL as text
        pictureCell.innerHTML = `📷 ${entry.pictureUrl || entry.picture}<br><small>${entry.caption || ""}</small>`;
        pictureCell.style.fontSize = "0.8rem";
      }
    } else {
      pictureCell.textContent = "—";
    }
    pictureCell.style.textAlign = "center";
    row.appendChild(pictureCell);

    // Data cells
    ["location", "observation", "suggested", "inform"].forEach(field => {
      const cell = document.createElement("td");
      cell.textContent = entry[field] || "—";
      if (field === "observation") {
        cell.style.color = "#d32f2f";
        cell.style.fontWeight = "600";
      }
      if (field === "suggested") {
        cell.style.color = "#2e7d32";
        cell.style.fontWeight = "600";
      }
      row.appendChild(cell);
    });

    tbody.appendChild(row);
  });
  table.appendChild(tbody);

  refs.bulkPreviewTable.innerHTML = "";
  refs.bulkPreviewTable.appendChild(table);
}

function handleAddBulk() {
  if (bulkData.length === 0) {
    showBulkMessage("No data to add. Please preview first.", "error");
    return;
  }
  renderBulkPreview(bulkData);
  refs.bulkPreview.hidden = false;
}

function handleConfirmBulk() {
  if (bulkData.length === 0) return;

  // Add all entries to the state
  bulkData.forEach(entry => {
    const newEntry = {
      id: Date.now() + Math.random(), // Simple unique ID
      location: entry.location,
      observation: entry.observation,
      suggested: entry.suggested,
      inform: entry.inform,
      caption: entry.caption,
      picture: entry.picture // This could be a URL or null
    };
    
    state.entries.push(newEntry);
  });

  // Clear bulk form
  handleClearBulk();
  
  // Re-render the report
  renderEntries();
  renderEntrySummary();
  
  showBulkMessage(`Successfully added ${bulkData.length} rows to the report!`, "success");
  
  // Hide preview after a delay
  setTimeout(() => {
    refs.bulkPreview.hidden = true;
  }, 2000);
}

function handleCancelBulk() {
  refs.bulkPreview.hidden = true;
  bulkData = [];
}

function handleClearBulk() {
  refs.bulkDataInput.value = "";
  refs.customSeparatorInput.value = "";
  refs.bulkPreview.hidden = true;
  refs.addBulkBtn.disabled = true;
  bulkData = [];
  clearBulkMessages();
}

function showBulkMessage(message, type) {
  clearBulkMessages();
  
  const messageDiv = document.createElement("div");
  messageDiv.className = type === "error" ? "bulk-error" : "bulk-success";
  messageDiv.textContent = message;
  messageDiv.id = "bulkMessage";
  
  refs.bulkPreviewTable.parentNode.insertBefore(messageDiv, refs.bulkPreviewTable);
}

function clearBulkMessages() {
  const existingMessage = document.getElementById("bulkMessage");
  if (existingMessage) {
    existingMessage.remove();
  }
}

// Column-Specific Bulk Upload Functions
let columnUpdateData = [];

function handleColumnSelectChange() {
  const selectedColumn = refs.columnSelect.value;
  const mode = refs.columnModeSelect.value;
  
  if (selectedColumn) {
    updateColumnPlaceholder(selectedColumn, mode);
  }
}

function updateColumnPlaceholder(column, mode) {
  const examples = {
    location: "LMV Crossing-01\nHMV Parking Area\nQuarry Site A\nSafety Station B",
    observation: "Safety equipment missing\nDust accumulation excessive\nTraffic lights malfunctioning\nEquipment in wrong position",
    suggested: "Install proper safety gear\nImplement dust control measures\nRepair or replace lights\nReposition equipment properly",
    inform: "Safety Department\nMaintenance Team\nElectrical Team\nOperations Manager",
    pictureUrl: "https://example.com/image1.jpg\nhttps://example.com/image2.jpg\nhttp://site.com/photo3.png",
    caption: "Safety equipment audit\nDust control inspection\nTraffic light maintenance\nEquipment positioning"
  };
  
  const columnNames = {
    location: "Location",
    observation: "Observation", 
    suggested: "Suggested Action",
    inform: "Inform To",
    pictureUrl: "Picture URL",
    caption: "Picture Caption"
  };
  
  const modeText = mode === "existing" ? "existing rows" : "new rows";
  const placeholder = `Enter ${columnNames[column]} values (one per line) to fill ${modeText}:\n\n${examples[column] || "Value 1\nValue 2\nValue 3"}`;
  
  refs.columnDataInput.placeholder = placeholder;
}

function handlePreviewColumn() {
  const column = refs.columnSelect.value;
  const mode = refs.columnModeSelect.value;
  const rawData = refs.columnDataInput.value.trim();
  
  if (!column) {
    showColumnMessage("Please select a column to update.", "error");
    return;
  }
  
  if (!rawData) {
    showColumnMessage("Please enter data for the selected column.", "error");
    return;
  }
  
  try {
    const values = rawData.split("\n").map(line => line.trim()).filter(line => line);
    
    if (values.length === 0) {
      showColumnMessage("No valid data found. Please check your input.", "error");
      return;
    }
    
    columnUpdateData = prepareColumnUpdate(column, values, mode);
    renderColumnPreview(columnUpdateData, column, mode);
    refs.columnPreview.hidden = false;
    refs.applyColumnBtn.disabled = false;
    
    showColumnMessage(`Preview ready: ${values.length} values for ${column} column`, "success");
  } catch (error) {
    showColumnMessage(`Error preparing column update: ${error.message}`, "error");
  }
}

function prepareColumnUpdate(column, values, mode) {
  const updateData = [];
  
  if (mode === "existing") {
    // Update existing rows
    values.forEach((value, index) => {
      if (index < state.entries.length) {
        const existingEntry = { ...state.entries[index] };
        const originalValue = existingEntry[column] || "";
        existingEntry[column] = value;
        updateData.push({
          ...existingEntry,
          rowIndex: index,
          isUpdate: true,
          originalValue: originalValue,
          newValue: value
        });
      } else {
        // Create new entry if we have more values than existing rows
        const newEntry = {
          id: Date.now() + Math.random() + index,
          location: column === "location" ? value : "",
          observation: column === "observation" ? value : "",
          suggested: column === "suggested" ? value : "",
          inform: column === "inform" ? value : "",
          pictureUrl: column === "pictureUrl" ? value : "",
          caption: column === "caption" ? value : "",
          picture: column === "pictureUrl" ? value : null,
          rowIndex: state.entries.length + (index - state.entries.length),
          isNewRow: true,
          newValue: value
        };
        updateData.push(newEntry);
      }
    });
  } else {
    // Create new rows
    values.forEach((value, index) => {
      const newEntry = {
        id: Date.now() + Math.random() + index,
        location: column === "location" ? value : "",
        observation: column === "observation" ? value : "",
        suggested: column === "suggested" ? value : "",
        inform: column === "inform" ? value : "",
        pictureUrl: column === "pictureUrl" ? value : "",
        caption: column === "caption" ? value : "",
        picture: column === "pictureUrl" ? value : null,
        rowIndex: state.entries.length + index,
        isNewRow: true,
        newValue: value
      };
      updateData.push(newEntry);
    });
  }
  
  return updateData;
}

function renderColumnPreview(data, column, mode) {
  const table = document.createElement("table");
  table.className = "column-preview-table";

  // Create header
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  ["S.NO", "LOCATION", "OBSERVATION", "SUGGESTED ACTION", "INFORM TO", "STATUS"].forEach(header => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Create body
  const tbody = document.createElement("tbody");
  data.forEach((entry, index) => {
    const row = document.createElement("tr");
    
    if (entry.isNewRow) {
      row.className = "column-new-row";
    } else if (entry.isUpdate) {
      row.className = "column-updated";
    }
    
    // S.NO
    const snoCell = document.createElement("td");
    snoCell.textContent = (entry.rowIndex + 1).toString();
    snoCell.style.fontWeight = "bold";
    snoCell.style.textAlign = "center";
    row.appendChild(snoCell);

    // Data cells
    ["location", "observation", "suggested", "inform"].forEach(field => {
      const cell = document.createElement("td");
      cell.textContent = entry[field] || "—";
      
      if (field === column) {
        cell.className = "column-highlight";
      }
      
      if (field === "observation") {
        cell.style.color = "#d32f2f";
      }
      if (field === "suggested") {
        cell.style.color = "#2e7d32";
      }
      row.appendChild(cell);
    });
    
    // Status
    const statusCell = document.createElement("td");
    if (entry.isNewRow) {
      statusCell.textContent = "New Row";
      statusCell.style.color = "#f59e0b";
      statusCell.style.fontWeight = "600";
    } else if (entry.isUpdate) {
      statusCell.textContent = `Updated (was: "${entry.originalValue || "empty"}")`;
      statusCell.style.color = "#16a34a";
      statusCell.style.fontWeight = "600";
    }
    row.appendChild(statusCell);

    tbody.appendChild(row);
  });
  table.appendChild(tbody);

  refs.columnPreviewTable.innerHTML = "";
  refs.columnPreviewTable.appendChild(table);
}

function handleApplyColumn() {
  if (columnUpdateData.length === 0) {
    showColumnMessage("No column data to apply. Please preview first.", "error");
    return;
  }
  refs.columnPreview.hidden = false;
}

function handleConfirmColumn() {
  if (columnUpdateData.length === 0) return;

  // Apply the column updates
  columnUpdateData.forEach(entry => {
    if (entry.isNewRow) {
      // Add new entry
      const newEntry = {
        id: entry.id,
        location: entry.location,
        observation: entry.observation,
        suggested: entry.suggested,
        inform: entry.inform,
        caption: entry.caption,
        picture: entry.picture
      };
      state.entries.push(newEntry);
    } else if (entry.isUpdate) {
      // Update existing entry
      if (state.entries[entry.rowIndex]) {
        const field = Object.keys(entry).find(key => 
          entry[key] === entry.newValue && 
          ["location", "observation", "suggested", "inform", "pictureUrl", "caption"].includes(key)
        );
        if (field) {
          if (field === "pictureUrl") {
            state.entries[entry.rowIndex].picture = entry.newValue;
          } else {
            state.entries[entry.rowIndex][field] = entry.newValue;
          }
        }
      }
    }
  });

  // Clear column form
  handleClearColumn();
  
  // Re-render the report
  renderEntries();
  renderEntrySummary();
  
  const column = refs.columnSelect.value;
  showColumnMessage(`Successfully updated ${columnUpdateData.length} rows in ${column} column!`, "success");
  
  // Hide preview after a delay
  setTimeout(() => {
    refs.columnPreview.hidden = true;
  }, 2000);
}

function handleCancelColumn() {
  refs.columnPreview.hidden = true;
  columnUpdateData = [];
}

function handleClearColumn() {
  refs.columnSelect.value = "";
  refs.columnDataInput.value = "";
  refs.columnDataInput.placeholder = "Select a column first...";
  refs.columnPreview.hidden = true;
  refs.applyColumnBtn.disabled = true;
  columnUpdateData = [];
  clearColumnMessages();
}

function showColumnMessage(message, type) {
  clearColumnMessages();
  
  const messageDiv = document.createElement("div");
  messageDiv.className = type === "error" ? "bulk-error" : "bulk-success";
  messageDiv.textContent = message;
  messageDiv.id = "columnMessage";
  
  refs.columnPreviewTable.parentNode.insertBefore(messageDiv, refs.columnPreviewTable);
}

function clearColumnMessages() {
  const existingMessage = document.getElementById("columnMessage");
  if (existingMessage) {
    existingMessage.remove();
  }
}

// Picture Upload Function for Individual Cells
function handlePictureUpload(rowIndex) {
  // Create a file input element
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = "image/*";
  fileInput.style.display = "none";
  
  fileInput.addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (!file) return;
    
    try {
      // Convert file to base64 for storage
      const imageData = await convertFileToBase64(file);
      
      // Update the entry
      if (state.entries[rowIndex]) {
        state.entries[rowIndex].imageData = imageData;
        state.entries[rowIndex].picture = imageData;
        
        // Optionally prompt for caption
        const caption = prompt("Enter a caption for this image (optional):");
        if (caption !== null) {
          state.entries[rowIndex].caption = caption.trim();
        }
        
        // Re-render the entries to show the new image
        renderEntries();
      }
    } catch (error) {
      alert("Error uploading image: " + error.message);
    }
    
    // Clean up
    document.body.removeChild(fileInput);
  });
  
  // Add to body and trigger click
  document.body.appendChild(fileInput);
  fileInput.click();
}

// Helper function to convert file to base64
function convertFileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = error => reject(error);
    reader.readAsDataURL(file);
  });
}

// ---- Export / Print Helpers ----
function handleDownloadPdf() {
  window.print();
}

function handleDownloadDocx() {
  try {
    const title = getDocxTitle();
    const html = buildDocxHtml(title);
    const blob = createDocxFromHtml(html, { orientation: 'landscape', title });
    const filename = `${title || 'InspectionReport'}.docx`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename.replace(/[^a-z0-9._-]/gi, '_');
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error('DOCX export failed', err);
    alert('DOCX export failed: ' + (err && err.message ? err.message : 'unknown'));
  }
}

function getDocxTitle() {
  const area = state.area || 'Inspection Report';
  const date = state.date ? state.date.replace(/[^0-9-]/g, '') : '';
  return date ? `${area}-${date}` : area;
}

function buildDocxHtml(title) {
  const safeTitle = escapeHtml(title || 'Inspection Report');
  const headerRows = [
    { label: 'Inspected Area', value: state.area || '' },
    { label: 'Visited with', value: state.visited || '' },
    { label: 'Date', value: formatDateForDisplay(state.date) || '' },
    { label: 'Shift', value: state.shift ? `Shift- ${state.shift}` : '' },
  ];

  const entryRows = state.entries.map((entry, index) => {
    const picture = entry.imageData || entry.picture || '';
    const caption = entry.caption ? `<div class="picture-caption">${escapeHtml(entry.caption)}</div>` : '';
    const pictureHtml = picture
      ? `<div class="picture-wrapper"><img src="${escapeHtml(picture)}" alt="Row ${index + 1} picture" /></div>${caption}`
      : `<div class="picture-wrapper no-image">No image</div>${caption}`;

    return `
      <tr>
        <td class="sno-cell"><span class="sno-index">${index + 1}.</span></td>
        <td class="picture-cell">${pictureHtml}</td>
        <td>${formatMultiline(entry.location)}</td>
        <td class="observation-cell">${formatMultiline(entry.observation)}</td>
        <td class="suggested-cell">${formatMultiline(entry.suggested)}</td>
        <td>${formatMultiline(entry.inform || '')}</td>
      </tr>
    `;
  }).join('');

  const headerCellsHtml = headerRows.map((row) => `
    <td>
      <div class="label">${escapeHtml(row.label)}</div>
      <div class="value">${escapeHtml(row.value)}</div>
    </td>
  `).join('');

  const tableHtml = entryRows || `
    <tr>
      <td class="empty" colspan="6">No observations added.</td>
    </tr>
  `;

  const styles = `
    <style>
      body { font-family: "Segoe UI", Arial, sans-serif; margin: 0; padding: 24px; color: #111827; background: #ffffff; }
      .report-header-table { width: 100%; border-collapse: collapse; border: 1px solid #2c5aa0; }
      .report-header-table td { background: #4a7bc8; color: #ffffff; padding: 12px 16px; border-right: 1px solid rgba(255,255,255,0.35); vertical-align: top; }
      .report-header-table td:last-child { border-right: none; }
      .report-header-table .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; font-weight: 600; margin-bottom: 4px; }
      .report-header-table .value { font-size: 13px; font-weight: 500; }
      table.report-table { width: 100%; border-collapse: collapse; margin-top: 16px; table-layout: fixed; font-size: 12px; }
      table.report-table thead th { background: #4a7bc8; color: #ffffff; font-weight: 600; letter-spacing: 0.04em; padding: 10px 8px; border: 1px solid #2c5aa0; }
      table.report-table td { border: 1px solid #2c5aa0; padding: 10px 8px; vertical-align: top; }
      table.report-table td.observation-cell { background: #ffebee; color: #b71c1c; }
      table.report-table td.suggested-cell { background: #e8f5e8; color: #1b5e20; }
      table.report-table td.empty { text-align: center; font-style: italic; color: #6b7280; padding: 30px 12px; }
      .picture-wrapper { width: 100%; min-height: 220px; border: 1px solid #b5bfd3; border-radius: 8px; background: #0f1c38; display: flex; align-items: center; justify-content: center; overflow: hidden; color: #f3f4f6; }
      .picture-wrapper img { width: 100%; height: 100%; object-fit: cover; display: block; }
      .picture-wrapper.no-image { font-size: 12px; letter-spacing: 0.05em; text-transform: uppercase; background: #111827; color: #e5e7eb; }
      .picture-caption { font-size: 11px; color: #4b5563; margin-top: 6px; text-align: center; }
      .sno-cell { font-weight: 700; text-align: center; }
      .sno-index { font-weight: 700; display: inline-block; }
    </style>
  `;

  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <title>${safeTitle}</title>
    ${styles}
  </head>
  <body>
    <table class="report-header-table">
      <tbody>
        <tr>
          ${headerCellsHtml}
        </tr>
      </tbody>
    </table>
    <table class="report-table">
      <thead>
        <tr>
          <th>S.NO</th>
          <th>PICTURE</th>
          <th>LOCATION</th>
          <th>OBSERVATION</th>
          <th>SUGGESTED ACTION</th>
          <th>INFORM TO</th>
        </tr>
      </thead>
      <tbody>
        ${tableHtml}
      </tbody>
    </table>
  </body>
</html>`;
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeXml(str) {
  return escapeHtml(str);
}

function formatMultiline(value) {
  return escapeHtml(value || '').replace(/\r?\n/g, '<br />');
}

function createDocxFromHtml(html, options = {}) {
  const encoder = new TextEncoder();
  const orientation = options.orientation === 'landscape' ? 'landscape' : 'portrait';
  const title = options.title || 'Inspection Report';
  const now = new Date();

  const entries = [
    { path: '[Content_Types].xml', data: encoder.encode(buildContentTypes()) },
    { path: '_rels/.rels', data: encoder.encode(buildRootRels()) },
    { path: 'docProps/core.xml', data: encoder.encode(buildCoreXml(title, now)) },
    { path: 'docProps/app.xml', data: encoder.encode(buildAppXml(title)) },
    { path: 'word/_rels/document.xml.rels', data: encoder.encode(buildDocumentRels()) },
    { path: 'word/document.xml', data: encoder.encode(buildDocumentXml(orientation)) },
    { path: 'word/styles.xml', data: encoder.encode(buildStylesXml()) },
    { path: DOCX_HTML_CHUNK_PATH, data: encoder.encode(html) },
  ];

  const zipBytes = createDocxZip(entries, now);
  return new Blob([zipBytes], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

function buildContentTypes() {
  const chunkPartName = `/${DOCX_HTML_CHUNK_PATH}`;
  return `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="${chunkPartName}" ContentType="application/xhtml+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

function buildRootRels() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function buildDocumentRels() {
  const chunkTarget = DOCX_HTML_CHUNK_PATH.includes('/')
    ? DOCX_HTML_CHUNK_PATH.slice(DOCX_HTML_CHUNK_PATH.lastIndexOf('/') + 1)
    : DOCX_HTML_CHUNK_PATH;
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/afChunk" Target="${chunkTarget}"/>
</Relationships>`;
}

function buildDocumentXml(orientation) {
  const isLandscape = orientation === 'landscape';
  const pgWidth = isLandscape ? 16838 : 11906;
  const pgHeight = isLandscape ? 11906 : 16838;
  const orientAttr = isLandscape ? ' w:orient="landscape"' : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:altChunk r:id="rId2"/>
    <w:sectPr>
      <w:pgSz w:w="${pgWidth}" w:h="${pgHeight}"${orientAttr}/>
      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="160"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>`;
}

function buildCoreXml(title, date) {
  const escapedTitle = escapeXml(title);
  const created = date.toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escapedTitle}</dc:title>
  <dc:subject>Inspection Report</dc:subject>
  <dc:creator>Inspection Report Builder</dc:creator>
  <cp:lastModifiedBy>Inspection Report Builder</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>
</cp:coreProperties>`;
}

function buildAppXml(title) {
  const escapedTitle = escapeXml(title);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Inspection Report Builder</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>${escapedTitle}</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
</Properties>`;
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let j = 0; j < 8; j += 1) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let crc = 0 ^ -1;
  for (let i = 0; i < buf.length; i += 1) {
    crc = (crc >>> 8) ^ CRC32_TABLE[(crc ^ buf[i]) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

function createDocxZip(entries, date) {
  const encoder = new TextEncoder();
  const fileChunks = [];
  const centralChunks = [];
  let offset = 0;

  const { dosDate, dosTime } = getDosDateTime(date);

  entries.forEach((entry) => {
    const nameBytes = encoder.encode(entry.path);
    const data = entry.data;
    const crc = crc32(data);
    const compressedSize = data.length;
    const uncompressedSize = data.length;

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, dosTime, true);
    localView.setUint16(12, dosDate, true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, compressedSize, true);
    localView.setUint32(22, uncompressedSize, true);
    localView.setUint16(26, nameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);

    fileChunks.push(localHeader, data);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 0x0314, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, dosTime, true);
    centralView.setUint16(14, dosDate, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, compressedSize, true);
    centralView.setUint32(24, uncompressedSize, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);

    centralChunks.push(centralHeader);

    offset += localHeader.length + data.length;
  });

  const centralSize = centralChunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const endRecord = new Uint8Array(22);
  const endView = new DataView(endRecord.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, offset, true);
  endView.setUint16(20, 0, true);

  const allChunks = [...fileChunks, ...centralChunks, endRecord];
  return concatUint8Arrays(allChunks);
}

function getDosDateTime(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = Math.floor(date.getSeconds() / 2);

  const dosDate = ((year - 1980) << 9) | (month << 5) | day;
  const dosTime = (hours << 11) | (minutes << 5) | seconds;

  return { dosDate, dosTime };
}

function concatUint8Arrays(chunks) {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  chunks.forEach((chunk) => {
    result.set(chunk, offset);
    offset += chunk.length;
  });
  return result;
}
