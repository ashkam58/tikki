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
};

init();

function init() {
  hydrateDefaultDate();
  attachMetaListeners();
  refs.addRowBtn.addEventListener("click", handleAddRow);
  refs.clearRowBtn.addEventListener("click", handleClearEntryForm);
  refs.downloadBtn.addEventListener("click", () => window.print());
  refs.resetReportBtn.addEventListener("click", handleResetReport);
  
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
    snoCell.textContent = `${index + 1}.`;

    const controls = createRowControls(index);
    if (controls) {
      snoCell.appendChild(controls);
    }

    const pictureCell = document.createElement("td");
    pictureCell.className = "picture-cell";
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
    locationCell.textContent = entry.location;

    const observationCell = document.createElement("td");
    observationCell.className = "observation-cell text-cell";
    observationCell.textContent = entry.observation;

    const suggestedCell = document.createElement("td");
    suggestedCell.className = "suggested-cell text-cell";
    suggestedCell.textContent = entry.suggested;

    const informCell = document.createElement("td");
    informCell.className = "inform-cell text-cell";
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
