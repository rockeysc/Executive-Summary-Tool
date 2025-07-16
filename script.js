// Table templates for different types of data
const tableTemplates = {
  "performance-trends": {
    headers: [
      "",
      "5/15 - 5/21 (Prior Wednesday and Prior 7 Days)",
      "5/22 - 5/28 (Most Recent Wednesday and Prior 7 Days)",
    ],
    defaultRows: [
      ["IMPRESSIONS", "", ""],
      ["GROSS REVENUE", "", ""],
      ["GROSS CPM", "", ""],
      ["OBSERVATIONS", "", ""],
    ],
    specialFormatting: {
      observationsRow: 3, // Index of the observations row
      mergeObservations: true, // Merge the observations cell across both data columns
      editableHeaders: [1, 2], // Indices of headers that should be editable week pickers
    },
  },
  "ab-test-updates": {
    headers: ["", "", "", "", "", ""],
    defaultRows: [
      ["PUBLISHER", "", "TIER", ""],
      ["COMPETITOR", "", "% SPLIT", ""],
      ["STAGE", "", "CSM/OBS", ""],
      ["STAKEHOLDERS", "", "", ""],
      ["START DATE", "", "END DATE", "", "LAST COMM. DATE", ""],
      ["TEST STATUS", "", "", ""],
    ],
    specialFormatting: {
      customLayout: true,
      statusRow: 5, // Index of the test status row for special formatting
    },
  },
  "newly-onboarded-publishers": {
    headers: ["", "", "", ""],
    defaultRows: [
      ["PUBLISHER", "", "LIVE DATE", ""],
      ["PROJECTED REVENUE", "", "OBS", ""],
      ["DEMAND STATUS", "", "", ""],
    ],
    specialFormatting: {
      customLayout: true,
      demandStatusRow: 2, // Index of the demand status row for special formatting
    },
  },
  "in-progress-publishers": {
    headers: ["", "", "", ""],
    defaultRows: [
      ["PUBLISHER", "", "EXP. LIVE DATE", ""],
      ["STAGE", "", "OBS", ""],
      ["PROJECTED REVENUE", "", "LAST COMM. DATE", ""],
      ["STATUS", "", "", ""],
    ],
    specialFormatting: {
      customLayout: true,
      statusRow: 3, // Index of the status row for special formatting
    },
  },
  "nps-data": {
    headers: ["", "MTD", "QTD", "YTD", ""],
    defaultRows: [
      ["NETWORK", "", "", "", ""],
      ["T1", "", "", "", ""],
    ],
    specialFormatting: {
      customLayout: true,
      npsTable: true,
    },
  },
  "churned-data": {
    headers: [
      "PUBLISHER",
      "SITE(S)",
      "AVG. MONTHLY REV.",
      "DARK DATE & REASON",
      "",
    ],
    defaultRows: [["", "", "", "", ""]],
    specialFormatting: {
      customLayout: true,
      churnedTable: true,
    },
  },
  "notice-data": {
    headers: ["PUBLISHER", "AVG. MONTHLY REV.", "NOTICE DATE & REASON"],
    defaultRows: [["", "", ""]],
    specialFormatting: {
      customLayout: true,
      noticeTable: true,
    },
  },
  "publisher-updates": {
    headers: ["", "", "", "", "", ""],
    defaultRows: [
      ["PUBLISHER", "", "TIER", "", "AVG. MONTHLY REV.", ""],
      ["STAGE", "", "", "CSM", "", ""],
      ["STAKEHOLDERS", "", "", "LAST COMM. DATE", "", ""],
      ["ISSUE(S)/ UPDATE(S)", "", "", "", "", ""],
      ["NEXT STEPS", "", "", "", "", ""],
    ],
    specialFormatting: {
      customLayout: true,
      publisherUpdatesTable: true,
    },
  },
};

let tableCounter = 0;

// Function to check if a cell should be a date input
function isDateField(labelText, cellIndex, rowIndex, tableType) {
  const dateKeywords = [
    "DATE",
    "LIVE DATE",
    "EXP. LIVE DATE",
    "START DATE",
    "END DATE",
    "LAST COMM. DATE",
  ];

  // Check if the label contains date keywords
  if (dateKeywords.some((keyword) => labelText.includes(keyword))) {
    return true;
  }

  // Specific checks for different table types
  if (tableType === "ab-test-updates" && rowIndex === 4) {
    // START DATE row: positions 1, 3, 5 are date fields
    return cellIndex === 1 || cellIndex === 3 || cellIndex === 5;
  }

  if (
    tableType === "newly-onboarded-publishers" &&
    rowIndex === 0 &&
    cellIndex === 3
  ) {
    // LIVE DATE field
    return true;
  }

  if (tableType === "in-progress-publishers") {
    if (
      (rowIndex === 0 && cellIndex === 3) ||
      (rowIndex === 2 && cellIndex === 3)
    ) {
      // EXP. LIVE DATE and LAST COMM. DATE fields
      return true;
    }
  }

  if (tableType === "publisher-updates" && rowIndex === 2 && cellIndex === 4) {
    // LAST COMM. DATE field
    return true;
  }

  return false;
}

// Function to create date input HTML
function createDateInput(value = "", className = "") {
  const dateValue = value ? formatDateForInput(value) : "";
  const displayValue = value || "mm/dd/yyyy";
  const uniqueId = `date-${Math.random().toString(36).substring(2, 11)}`;

  return `
    <div class="date-field-wrapper ${className}">
      <div class="date-display" onclick="document.getElementById('${uniqueId}').showPicker()">
        <span class="date-text">${displayValue}</span>
        <span class="material-icons date-icon">calendar_today</span>
      </div>
      <input type="date" id="${uniqueId}" class="date-input-hidden" value="${dateValue}" onchange="handleDateChange(this)">
    </div>
  `;
}

// Function to format date for input (convert from m/d/y to yyyy-mm-dd)
function formatDateForInput(dateStr) {
  if (!dateStr || dateStr.trim() === "") return "";

  // Try to parse various date formats
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) return "";

  return date.toISOString().split("T")[0];
}

// Function to format date for display (convert from yyyy-mm-dd to m/d/y)
function formatDateForDisplay(dateStr) {
  if (!dateStr) return "";

  const date = new Date(dateStr);
  if (isNaN(date.getTime())) return dateStr;

  return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
}

// Handle date input changes
function handleDateChange(input) {
  const displayValue = formatDateForDisplay(input.value);
  // Store the formatted value in a data attribute for export
  input.setAttribute("data-display-value", displayValue);

  // Update the display text in the wrapper
  const wrapper = input.closest(".date-field-wrapper");
  if (wrapper) {
    const dateText = wrapper.querySelector(".date-text");
    if (dateText) {
      dateText.textContent = displayValue || "mm/dd/yyyy";
    }
  }

  // Trigger auto-save when date changes
  triggerAutoSave();
}

// Function to create week picker HTML
function createWeekPicker(value = "", className = "") {
  const dateValue = value ? extractDateFromWeekRange(value) : "";
  return `<input type="date" class="week-picker ${className}" value="${dateValue}" onchange="handleWeekChange(this)" title="Select a Wednesday to set the week range">`;
}

// Function to extract date from week range (e.g., "5/15 - 5/21" -> "2024-05-15")
function extractDateFromWeekRange(weekRange) {
  if (!weekRange || weekRange.trim() === "") return "";

  // Extract the start date from the range (e.g., "5/15" from "5/15 - 5/21")
  const startDateStr = weekRange.split(" - ")[0];
  if (!startDateStr) return "";

  // Parse the date (assuming current year if not specified)
  const currentYear = new Date().getFullYear();
  const date = new Date(`${startDateStr}/${currentYear}`);

  if (isNaN(date.getTime())) return "";

  return date.toISOString().split("T")[0];
}

// Function to format week range from selected date
function formatWeekRange(selectedDate) {
  if (!selectedDate) return "";

  const date = new Date(selectedDate);
  if (isNaN(date.getTime())) return "";

  // Find the Wednesday of the week containing the selected date
  const dayOfWeek = date.getDay(); // 0 = Sunday, 3 = Wednesday
  const daysToWednesday = (3 - dayOfWeek + 7) % 7; // Days to add to get to Wednesday

  const wednesday = new Date(date);
  wednesday.setDate(date.getDate() + daysToWednesday);

  // Calculate the date range (Wednesday to Tuesday, 7 days total)
  const startDate = new Date(wednesday);
  const endDate = new Date(wednesday);
  endDate.setDate(wednesday.getDate() + 6); // 6 days after Wednesday

  // Format as m/d
  const startFormatted = `${startDate.getMonth() + 1}/${startDate.getDate()}`;
  const endFormatted = `${endDate.getMonth() + 1}/${endDate.getDate()}`;

  return `${startFormatted} - ${endFormatted}`;
}

// Handle week picker changes
function handleWeekChange(input) {
  const weekRange = formatWeekRange(input.value);
  // Update the display value
  input.setAttribute("data-display-value", weekRange);

  // Update the header if this is a header cell
  const headerCell = input.closest("th");
  if (headerCell && headerCell.classList.contains("editable-header")) {
    // Find the text content and update it while preserving the description and week picker
    const currentText = headerCell.textContent;
    const description = currentText.includes("Prior Wednesday")
      ? " (Prior Wednesday and Prior 7 Days)"
      : currentText.includes("Most Recent")
      ? " (Most Recent Wednesday and Prior 7 Days)"
      : "";

    // Recreate the header with updated week range
    headerCell.innerHTML = `${weekRange}${description}<br>${createWeekPicker(
      weekRange,
      "header-week-picker"
    )}`;
  }

  // Trigger auto-save when week range changes
  triggerAutoSave();
}

// Function to add a new table
function addTable(subsectionId, tableType) {
  tableCounter++;
  const container = document.getElementById(`${subsectionId}-tables`);
  const template = tableTemplates[tableType];

  if (!container) {
    console.error(
      "No container found for subsection:",
      subsectionId,
      "- looking for element with ID:",
      `${subsectionId}-tables`
    );
    return;
  }

  if (!template) {
    console.error("No template found for:", tableType);
    return;
  }

  const tableWrapper = document.createElement("div");
  tableWrapper.className = "table-wrapper";
  tableWrapper.id = `table-${tableCounter}`;

  // Generate table rows with special formatting
  let rowsHTML = "";
  if (tableType === "performance-trends" && template.specialFormatting) {
    // Generate div-based structure for performance trends
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (
          index === template.specialFormatting.observationsRow &&
          template.specialFormatting.mergeObservations
        ) {
          // Special formatting for observations row - span across both data columns
          return `<div class="performance-row observations-row">
          <div class="metric-label">${row[0]}</div>
          <div contenteditable="true" class="observations-cell">${row[1]}</div>
        </div>`;
        } else {
          return `<div class="performance-row">
            <div class="metric-label">${row[0]}</div>
            <div contenteditable="true" class="metric-data">${row[1]}</div>
            <div contenteditable="true" class="metric-data">${row[2]}</div>
          </div>`;
        }
      })
      .join("");
  } else if (
    tableType === "ab-test-updates" &&
    template.specialFormatting &&
    template.specialFormatting.customLayout
  ) {
    // Generate div-based structure for A/B test updates
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (index === 4) {
          // START DATE row with 3 date inputs
          return `<div class="ab-test-row dates-row">
            <div class="ab-label">${row[0]}</div>
            <div class="ab-date-input">${createDateInput(row[1])}</div>
            <div class="ab-label">${row[2]}</div>
            <div class="ab-date-input">${createDateInput(row[3])}</div>
            <div class="ab-label">${row[4]}</div>
            <div class="ab-date-input">${createDateInput(row[5])}</div>
          </div>`;
        } else if (index === template.specialFormatting.statusRow) {
          // TEST STATUS row spans multiple columns
          return `<div class="ab-test-row status-row">
            <div class="ab-label">${row[0]}</div>
            <div contenteditable="true" class="test-status-cell">${row[1]}</div>
          </div>`;
        } else if (index === 3) {
          // STAKEHOLDERS row spans all remaining columns
          return `<div class="ab-test-row stakeholders-row">
            <div class="ab-label">${row[0]}</div>
            <div contenteditable="true" class="stakeholders-cell">${row[1]}</div>
          </div>`;
        } else if (index === 0) {
          // PUBLISHER row with TIER spanning wider
          return `<div class="ab-test-row regular-row publisher-row">
            <div class="ab-label">${row[0]}</div>
            <div contenteditable="true" class="ab-data">${row[1]}</div>
            <div class="ab-label">${row[2]}</div>
            <div contenteditable="true" class="ab-data-tier-wide">${row[3]}</div>
          </div>`;
        } else {
          // Regular rows with 4 columns
          return `<div class="ab-test-row regular-row">
            <div class="ab-label">${row[0]}</div>
            <div contenteditable="true" class="ab-data">${row[1]}</div>
            <div class="ab-label">${row[2]}</div>
            <div contenteditable="true" class="ab-data-wide">${row[3]}</div>
          </div>`;
        }
      })
      .join("");
  } else if (
    tableType === "newly-onboarded-publishers" &&
    template.specialFormatting &&
    template.specialFormatting.customLayout
  ) {
    // Generate div-based structure for newly onboarded publishers
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (index === template.specialFormatting.demandStatusRow) {
          // DEMAND STATUS row spans multiple columns
          return `<div class="newly-onboarded-row demand-status-row">
            <div class="publisher-label">${row[0]}</div>
            <div contenteditable="true" class="demand-status-cell">${row[1]}</div>
          </div>`;
        } else {
          // Regular rows with 4 columns
          return `<div class="newly-onboarded-row regular-row">
            <div class="publisher-label">${row[0]}</div>
            <div contenteditable="true" class="publisher-data">${row[1]}</div>
            <div class="publisher-label">${row[2]}</div>
            <div class="publisher-data-input">${
              isDateField(row[2], 3, index, tableType)
                ? createDateInput(row[3])
                : `<span contenteditable="true">${row[3]}</span>`
            }</div>
          </div>`;
        }
      })
      .join("");
  } else if (tableType === "in-progress-publishers") {
    // Generate div-based structure for in-progress publishers
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (index === 3) {
          // STATUS row (index 3)
          // STATUS row spans multiple columns
          return `<div class="in-progress-row status-row">
            <div class="progress-label">${row[0] || ""}</div>
            <div contenteditable="true" class="progress-status-cell">${
              row[1] || ""
            }</div>
          </div>`;
        } else {
          // Regular rows with 4 columns
          return `<div class="in-progress-row regular-row">
            <div class="progress-label">${row[0] || ""}</div>
            <div contenteditable="true" class="progress-data">${
              row[1] || ""
            }</div>
            <div class="progress-label">${row[2] || ""}</div>
            <div class="progress-data-input">${
              isDateField(row[2], 3, index, tableType)
                ? createDateInput(row[3] || "")
                : `<span contenteditable="true">${row[3] || ""}</span>`
            }</div>
          </div>`;
        }
      })
      .join("");
  } else if (
    tableType === "nps-data" &&
    template.specialFormatting &&
    template.specialFormatting.npsTable
  ) {
    // Generate div-based structure for NPS table
    rowsHTML = template.defaultRows
      .map((row, index) => {
        // For NPS table, all initial rows are default rows (NETWORK and T1)
        return `<div class="nps-row nps-default-row">
          <div class="nps-label">${row[0]}</div>
          <div contenteditable="true" class="nps-score">${row[1]}</div>
          <div contenteditable="true" class="nps-score">${row[2]}</div>
          <div contenteditable="true" class="nps-score">${row[3]}</div>
          <div class="row-actions">
            <button class="add-row-btn" onclick="addRowToTable(this)" title="Add row"><span class="material-icons">add</span></button>
          </div>
        </div>`;
      })
      .join("");
  } else if (
    tableType === "churned-data" &&
    template.specialFormatting &&
    template.specialFormatting.churnedTable
  ) {
    // Generate simple div-based structure for churned data table (all 4 columns + actions)
    rowsHTML = template.defaultRows
      .map((row, index) => {
        // First row gets + button, others get - button (if there are multiple rows)
        const isFirstRow = index === 0;
        const actionButton = isFirstRow
          ? `<button class="add-row-btn" onclick="addRowToTable(this)" title="Add row"><span class="material-icons">add</span></button>`
          : `<button class="remove-row-btn" onclick="removeRowFromTable(this)" title="Remove row"><span class="material-icons">remove</span></button>`;

        return `<div class="churned-row ${
          isFirstRow ? "churned-default-row" : "churned-added-row"
        }">
          <div contenteditable="true" class="churned-publisher">${row[0]}</div>
          <div contenteditable="true" class="churned-site">${row[1]}</div>
          <div contenteditable="true" class="churned-revenue">${row[2]}</div>
          <div contenteditable="true" class="churned-reason">${row[3]}</div>
          <div class="row-actions">${actionButton}</div>
        </div>`;
      })
      .join("");
  } else if (
    tableType === "notice-data" &&
    template.specialFormatting &&
    template.specialFormatting.noticeTable
  ) {
    // Generate div-based structure for notice data table
    rowsHTML = template.defaultRows
      .map((row, index) => {
        return `<div class="notice-row">
          <div contenteditable="true" class="notice-publisher">${row[0]}</div>
          <div contenteditable="true" class="notice-revenue">${row[1]}</div>
          <div contenteditable="true" class="notice-reason">${row[2]}</div>
        </div>`;
      })
      .join("");
  } else if (
    tableType === "publisher-updates" &&
    template.specialFormatting &&
    template.specialFormatting.publisherUpdatesTable
  ) {
    // Generate div-based structure for publisher updates table
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (index === 0) {
          // First row: PUBLISHER | name | TIER | 1 | AVG. MONTHLY REV. | 20,000
          return `<div class="updates-row updates-row-1">
            <div class="updates-label">${row[0]}</div>
            <div contenteditable="true" class="updates-data">${row[1]}</div>
            <div class="updates-label">${row[2]}</div>
            <div contenteditable="true" class="updates-data">${row[3]}</div>
            <div class="updates-label">${row[4]}</div>
            <div contenteditable="true" class="updates-data">${row[5]}</div>
          </div>`;
        } else if (index === 1) {
          // Second row: STAGE | Onboarding or Ongoing | | | CSM | Kristen
          return `<div class="updates-row updates-row-2">
            <div class="updates-label">${row[0]}</div>
            <div contenteditable="true" class="updates-data updates-data-span3">${row[1]}</div>
            <div class="updates-label">${row[3]}</div>
            <div contenteditable="true" class="updates-data">${row[4]}</div>
          </div>`;
        } else if (index === 2) {
          // Third row: STAKEHOLDERS | list of those involved | | | LAST COMM. DATE | date
          return `<div class="updates-row updates-row-3">
            <div class="updates-label">${row[0]}</div>
            <div contenteditable="true" class="updates-data updates-data-span3">${
              row[1]
            }</div>
            <div class="updates-label">${row[3]}</div>
            <div class="updates-data">${createDateInput(row[4])}</div>
          </div>`;
        } else if (index === 3) {
          // Fourth row: ISSUE(S)/ UPDATE(S) spanning full width
          return `<div class="updates-row updates-row-4">
            <div class="updates-label">${row[0]}</div>
            <div contenteditable="true" class="updates-issues updates-data-span5">${row[1]}</div>
          </div>`;
        } else if (index === 4) {
          // Fifth row: NEXT STEPS spanning full width
          return `<div class="updates-row updates-row-5">
            <div class="updates-label">${row[0]}</div>
            <div contenteditable="true" class="updates-steps updates-data-span5">${row[1]}</div>
          </div>`;
        }
      })
      .join("");
  } else {
    rowsHTML = template.defaultRows
      .map(
        (row) =>
          `<tr>${row
            .map((cell) => `<td contenteditable="true">${cell}</td>`)
            .join("")}</tr>`
      )
      .join("");
  }

  let tableHTML = "";

  if (tableType === "performance-trends") {
    // Generate div-based structure for performance trends
    tableHTML = `
        <div class="performance-trends-table">
            <div class="performance-header">
                ${template.headers
                  .map((header, index) => {
                    // Check if this header should be editable (week picker)
                    if (
                      template.specialFormatting &&
                      template.specialFormatting.editableHeaders &&
                      template.specialFormatting.editableHeaders.includes(index)
                    ) {
                      const weekRange = header.split(" (")[0]; // Extract just the date range
                      const description = header.includes("Prior Wednesday")
                        ? " (Prior Wednesday and Prior 7 Days)"
                        : header.includes("Most Recent")
                        ? " (Most Recent Wednesday and Prior 7 Days)"
                        : "";
                      return `<div class="performance-header-cell editable-header">${weekRange}${description}<br>${createWeekPicker(
                        weekRange,
                        "header-week-picker"
                      )}</div>`;
                    } else {
                      return `<div class="performance-header-cell">${header}</div>`;
                    }
                  })
                  .join("")}
            </div>
            <div class="performance-body">
                ${rowsHTML}
            </div>
        </div>`;
  } else if (tableType === "ab-test-updates") {
    // Generate div-based structure for A/B test updates
    tableHTML = `
        <div class="ab-test-table">
            <div class="ab-test-body">
                ${rowsHTML}
            </div>
        </div>`;
  } else if (tableType === "newly-onboarded-publishers") {
    // Generate div-based structure for newly onboarded publishers
    tableHTML = `
        <div class="newly-onboarded-table">
            <div class="newly-onboarded-body">
                ${rowsHTML}
            </div>
        </div>`;
  } else if (tableType === "in-progress-publishers") {
    // Generate div-based structure for in-progress publishers
    tableHTML = `
        <div class="in-progress-table">
            <div class="in-progress-body">
                ${rowsHTML}
            </div>
        </div>`;
  } else if (tableType === "nps-data") {
    // Generate div-based structure for NPS data
    tableHTML = `
        <div class="nps-table">
            <div class="nps-header">
                ${template.headers
                  .map(
                    (header) => `<div class="nps-header-cell">${header}</div>`
                  )
                  .join("")}
            </div>
            <div class="nps-body">
                ${rowsHTML}
            </div>
        </div>`;
  } else if (tableType === "churned-data") {
    // Generate simple div-based structure for churned data (just PUBLISHER for now)
    tableHTML = `
        <div class="churned-table">
            <div class="churned-header">
                ${template.headers
                  .map(
                    (header) =>
                      `<div class="churned-header-cell">${header}</div>`
                  )
                  .join("")}
            </div>
            <div class="churned-body">
                ${rowsHTML}
            </div>
        </div>`;
  } else if (tableType === "notice-data") {
    // Generate div-based structure for notice data
    tableHTML = `
        <div class="notice-table">
            <div class="notice-header">
                ${template.headers
                  .map(
                    (header) =>
                      `<div class="notice-header-cell">${header}</div>`
                  )
                  .join("")}
            </div>
            <div class="notice-body">
                ${rowsHTML}
            </div>
        </div>`;
  } else if (tableType === "publisher-updates") {
    // Generate div-based structure for publisher updates
    tableHTML = `
        <div class="publisher-updates-table">
            <div class="publisher-updates-body">
                ${rowsHTML}
            </div>
        </div>`;
  } else {
    // Keep existing table structure for other table types
    tableHTML = `
        <table class="data-table">
            <thead>
                <tr>
                    ${template.headers
                      .map((header, index) => {
                        // Check if this header should be editable (week picker)
                        if (
                          template.specialFormatting &&
                          template.specialFormatting.editableHeaders &&
                          template.specialFormatting.editableHeaders.includes(
                            index
                          )
                        ) {
                          const weekRange = header.split(" (")[0]; // Extract just the date range
                          const description = header.includes("Prior Wednesday")
                            ? " (Prior Wednesday and Prior 7 Days)"
                            : header.includes("Most Recent")
                            ? " (Most Recent Wednesday and Prior 7 Days)"
                            : "";
                          return `<th class="editable-header">${weekRange}${description}<br>${createWeekPicker(
                            weekRange,
                            "header-week-picker"
                          )}</th>`;
                        } else {
                          return `<th>${header}</th>`;
                        }
                      })
                      .join("")}
                </tr>
            </thead>
            <tbody>
                ${rowsHTML}
            </tbody>
        </table>`;
  }

  // Complete the table HTML with mobile cards and actions
  const completeTableHTML = `
        ${tableHTML}

        <!-- Mobile card layout for add-new-content -->
        ${generateEditableMobileCards(template, tableType, tableCounter)}

        ${
          subsectionId !== "wow-performance" &&
          subsectionId !== "nps" &&
          subsectionId !== "recently-churned"
            ? `<div class="table-actions">
            <button class="remove-table-btn" onclick="removeTable('table-${tableCounter}')"><span class="material-icons">delete</span> Remove Table</button>
        </div>`
            : ""
        }
    `;

  tableWrapper.innerHTML = completeTableHTML;
  container.appendChild(tableWrapper);

  // Add event listeners for the new table
  addTableEventListeners(tableWrapper);

  // Auto-save when a new table is added
  triggerAutoSave();
}

// Function to remove a table
async function removeTable(tableId) {
  const table = document.getElementById(tableId);
  if (table) {
    const confirmed = await showCustomConfirm(
      "Are you sure you want to remove this table?",
      "Remove Table"
    );
    if (confirmed) {
      table.remove();
      // Auto-save when a table is removed
      triggerAutoSave();
    }
  }
}

// Function to add a new row to a table
function addRow(tableId) {
  const tableWrapper = document.getElementById(tableId);
  const table = tableWrapper.querySelector(".data-table tbody");
  const headerCount = tableWrapper.querySelector(".data-table thead tr")
    .children.length;

  const newRow = document.createElement("tr");
  // Add data cells
  for (let i = 0; i < headerCount; i++) {
    const cell = document.createElement("td");

    // Check if this should be a date field
    const tableType = tableWrapper.id
      .replace("table-", "")
      .split("-")
      .slice(0, -1)
      .join("-");
    const rowIndex = table.rows.length; // Current row index

    if (isDateField("", i, rowIndex, tableType)) {
      cell.innerHTML = createDateInput("");
    } else {
      cell.contentEditable = true;
      cell.textContent = "";
    }
    newRow.appendChild(cell);
  }

  table.appendChild(newRow);
  addRowEventListeners(newRow);
}

// Function to remove a row
async function removeRow(button) {
  const row = button.closest("tr");
  const table = row.closest("tbody");

  // Don't allow removing the last row
  if (table.children.length > 1) {
    const confirmed = await showCustomConfirm(
      "Are you sure you want to remove this row?",
      "Remove Row"
    );
    if (confirmed) {
      row.remove();
    }
  } else {
    await showCustomAlert(
      "Cannot remove the last row. Tables must have at least one row.",
      "Cannot Remove Row",
      "warning"
    );
  }
}

// Function to toggle section visibility
function toggleSection(sectionId) {
  const section = document.getElementById(sectionId);
  const content = section.querySelector(".section-content");
  const button = section.querySelector(".collapse-btn");

  if (content.classList.contains("collapsed")) {
    content.classList.remove("collapsed");
    button.textContent = "−";
  } else {
    content.classList.add("collapsed");
    button.textContent = "+";
  }
}

// Function to generate mobile-friendly card layout for editable tables in add-new-content
function generateEditableMobileCards(template, tableType, tableId) {
  if (
    !template.headers ||
    !template.defaultRows ||
    template.defaultRows.length === 0
  ) {
    return "";
  }

  let mobileHTML = "";

  // Handle different table layouts based on structure
  if (tableType === "performance-trends") {
    // Special handling for WEEK-OVER-WEEK PERFORMANCE TRENDS
    mobileHTML = generateEditablePerformanceTrendsCard(template, tableId);
  } else if (tableType === "ab-test-updates") {
    // Special handling for A/B test updates
    mobileHTML = generateEditableABTestCard(template, tableId);
  } else if (tableType === "newly-onboarded-publishers") {
    // Special handling for newly onboarded publishers
    mobileHTML = generateEditableNewlyOnboardedCard(template, tableId);
  } else if (tableType === "in-progress-publishers") {
    // Special handling for in-progress publishers
    mobileHTML = generateEditableInProgressCard(template, tableId);
  } else if (tableType === "publisher-updates") {
    // Special handling for publisher updates
    mobileHTML = generateEditablePublisherUpdatesCard(template, tableId);
  } else if (template.headers.length > 2) {
    // For tables with multiple columns, create cards for each row
    template.defaultRows.forEach((row, rowIndex) => {
      mobileHTML += generateEditableRowCard(template, row, rowIndex, tableId);
    });
  } else {
    // For simple two-column tables, create a single card
    mobileHTML += generateEditableSimpleCard(template, tableId);
  }

  return mobileHTML;
}

// Generate editable performance trends card for add-new-content
function generateEditablePerformanceTrendsCard(template, tableId) {
  let mobileHTML = `<div class="mobile-table-card">`;

  // Add date headers as a special header section
  mobileHTML += `<div class="mobile-table-header">Performance Trends</div>`;
  mobileHTML += `<div class="mobile-date-headers">`;
  mobileHTML += `<div class="mobile-date-header prior">${template.headers[1]}</div>`;
  mobileHTML += `<div class="mobile-date-header recent">${template.headers[2]}</div>`;
  mobileHTML += `</div>`;

  mobileHTML += `<div class="mobile-table-content">`;

  // Process each metric row
  template.defaultRows.forEach((row, index) => {
    if (row[0] && row[0].trim() !== "") {
      if (
        index === template.specialFormatting?.observationsRow &&
        template.specialFormatting?.mergeObservations
      ) {
        // Special handling for observations row
        mobileHTML += `<div class="observations-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
        mobileHTML += `<div class="mobile-row-data observations mobile-editable" contenteditable="true">${
          row[1] || ""
        }</div>`;
        mobileHTML += `</div>`;
      } else {
        // Regular metric row with two values
        mobileHTML += `<div class="mobile-metric-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
        mobileHTML += `<div class="mobile-metric-values">`;
        mobileHTML += `<div class="mobile-metric-value">`;
        mobileHTML += `<div class="metric-period">Prior</div>`;
        mobileHTML += `<div class="metric-data mobile-editable" contenteditable="true">${
          row[1] || ""
        }</div>`;
        mobileHTML += `</div>`;
        mobileHTML += `<div class="mobile-metric-value">`;
        mobileHTML += `<div class="metric-period">Recent</div>`;
        mobileHTML += `<div class="metric-data mobile-editable" contenteditable="true">${
          row[2] || ""
        }</div>`;
        mobileHTML += `</div>`;
        mobileHTML += `</div>`;
        mobileHTML += `</div>`;
      }
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate editable A/B test card for add-new-content
function generateEditableABTestCard(template, tableId) {
  let mobileHTML = `<div class="mobile-table-card">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Process the A/B test table structure
  template.defaultRows.forEach((row, index) => {
    if (index === 4) {
      // START DATE row with date inputs
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable">${createDateInput(
        row[1]
      )}</div>`;
      mobileHTML += `</div>`;
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[2]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable">${createDateInput(
        row[3]
      )}</div>`;
      mobileHTML += `</div>`;
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[4]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable">${createDateInput(
        row[5]
      )}</div>`;
      mobileHTML += `</div>`;
    } else if (index === template.specialFormatting?.statusRow) {
      // TEST STATUS row
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data status mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
    } else if (index === 3) {
      // STAKEHOLDERS row
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data stakeholders mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
    } else {
      // Regular rows
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
      if (row[2]) {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[2]}</div>`;
        mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
          row[3] || ""
        }</div>`;
        mobileHTML += `</div>`;
      }
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate simple editable card for basic tables
function generateEditableSimpleCard(template, tableId) {
  let cardHTML = `<div class="mobile-table-card">`;
  cardHTML += `<div class="mobile-table-content">`;

  template.defaultRows.forEach((row) => {
    if (row.length >= 2 && row[0] && row[1]) {
      cardHTML += `<div class="mobile-row">`;
      cardHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      cardHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${row[1]}</div>`;
      cardHTML += `</div>`;
    }
  });

  cardHTML += `</div>`;
  cardHTML += `</div>`;

  return cardHTML;
}

// Generate editable row card for multi-column tables
function generateEditableRowCard(template, row, rowIndex, tableId) {
  let cardHTML = `<div class="mobile-table-card">`;

  // Use first cell as card header if it's a label
  if (row[0]) {
    cardHTML += `<div class="mobile-table-header">${row[0]}</div>`;
  }

  cardHTML += `<div class="mobile-table-content">`;

  // Add remaining cells as data rows
  for (let i = 1; i < row.length; i++) {
    if (template.headers[i] && row[i] !== undefined) {
      cardHTML += `<div class="mobile-row">`;
      cardHTML += `<div class="mobile-row-label">${template.headers[i]}</div>`;
      cardHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${row[i]}</div>`;
      cardHTML += `</div>`;
    }
  }

  cardHTML += `</div>`;
  cardHTML += `</div>`;

  return cardHTML;
}

// Generate editable newly onboarded publishers card for add-new-content
function generateEditableNewlyOnboardedCard(template, tableId) {
  let mobileHTML = `<div class="mobile-table-card">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Process the newly onboarded publishers table structure
  template.defaultRows.forEach((row, index) => {
    if (index === template.specialFormatting?.demandStatusRow) {
      // DEMAND STATUS row spans multiple columns
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data status mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
    } else {
      // Regular rows with 4 columns
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      if (row[2] === "LIVE DATE") {
        mobileHTML += `<div class="mobile-row-data mobile-editable">${createDateInput(
          row[1] || ""
        )}</div>`;
      } else {
        mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
          row[1] || ""
        }</div>`;
      }
      mobileHTML += `</div>`;

      if (row[2]) {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[2]}</div>`;
        mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
          row[3] || ""
        }</div>`;
        mobileHTML += `</div>`;
      }
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate editable in-progress publishers card for add-new-content
function generateEditableInProgressCard(template, tableId) {
  let mobileHTML = `<div class="mobile-table-card">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Process the in-progress publishers table structure
  template.defaultRows.forEach((row, index) => {
    if (index === template.specialFormatting?.statusRow) {
      // STATUS row spans multiple columns
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data status mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
    } else {
      // Regular rows with 4 columns
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      if (row[2] === "EXP. LIVE DATE" || row[2] === "LAST COMM. DATE") {
        mobileHTML += `<div class="mobile-row-data mobile-editable">${createDateInput(
          row[1] || ""
        )}</div>`;
      } else {
        mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
          row[1] || ""
        }</div>`;
      }
      mobileHTML += `</div>`;

      if (row[2]) {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[2]}</div>`;
        if (row[2] === "EXP. LIVE DATE" || row[2] === "LAST COMM. DATE") {
          mobileHTML += `<div class="mobile-row-data mobile-editable">${createDateInput(
            row[3] || ""
          )}</div>`;
        } else {
          mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
            row[3] || ""
          }</div>`;
        }
        mobileHTML += `</div>`;
      }
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate editable publisher updates card for add-new-content
function generateEditablePublisherUpdatesCard(template, tableId) {
  let mobileHTML = `<div class="mobile-table-card">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Process the publisher updates table structure
  template.defaultRows.forEach((row, index) => {
    if (index === 0) {
      // First row: PUBLISHER | name | TIER | 1 | AVG. MONTHLY REV. | 20,000
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[2]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
        row[3] || ""
      }</div>`;
      mobileHTML += `</div>`;
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[4]}</div>`;
      mobileHTML += `<div class="mobile-row-data revenue mobile-editable" contenteditable="true">${
        row[5] || ""
      }</div>`;
      mobileHTML += `</div>`;
    } else if (index === 1) {
      // Second row: STAGE | Onboarding or Ongoing | | | CSM | Kristen
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[3]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable" contenteditable="true">${
        row[4] || ""
      }</div>`;
      mobileHTML += `</div>`;
    } else if (index === 2) {
      // Third row: STAKEHOLDERS | list of those involved | | | LAST COMM. DATE | date
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data stakeholders mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[3]}</div>`;
      mobileHTML += `<div class="mobile-row-data mobile-editable">${createDateInput(
        row[4] || ""
      )}</div>`;
      mobileHTML += `</div>`;
    } else if (index === 3) {
      // Fourth row: ISSUE(S)/ UPDATE(S) spanning full width
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data observations mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
    } else if (index === 4) {
      // Fifth row: NEXT STEPS spanning full width
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data observations mobile-editable" contenteditable="true">${
        row[1] || ""
      }</div>`;
      mobileHTML += `</div>`;
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Function to add event listeners to table elements
function addTableEventListeners(tableWrapper) {
  const cells = tableWrapper.querySelectorAll('td[contenteditable="true"]');
  cells.forEach((cell) => {
    addCellEventListeners(cell);
  });

  // Add event listeners to mobile card editable elements
  const mobileEditables = tableWrapper.querySelectorAll(".mobile-editable");
  mobileEditables.forEach((element) => {
    addCellEventListeners(element);
  });

  // Make mobile row data areas clickable to focus on editable content
  const mobileRowData = tableWrapper.querySelectorAll(".mobile-row-data");
  mobileRowData.forEach((rowData) => {
    rowData.addEventListener("click", (e) => {
      // Find the editable element within this row data
      const editableElement = rowData.querySelector(
        ".mobile-editable, [contenteditable='true']"
      );
      if (editableElement && editableElement !== e.target) {
        editableElement.focus();
        // Place cursor at the end of the content
        if (editableElement.textContent) {
          const range = document.createRange();
          const selection = window.getSelection();
          range.selectNodeContents(editableElement);
          range.collapse(false);
          selection.removeAllRanges();
          selection.addRange(range);
        }
      }
    });
  });

  // Make metric data areas clickable
  const metricData = tableWrapper.querySelectorAll(".metric-data");
  metricData.forEach((metric) => {
    metric.addEventListener("click", (e) => {
      if (metric.hasAttribute("contenteditable") && metric !== e.target) {
        metric.focus();
        // Place cursor at the end of the content
        if (metric.textContent) {
          const range = document.createRange();
          const selection = window.getSelection();
          range.selectNodeContents(metric);
          range.collapse(false);
          selection.removeAllRanges();
          selection.addRange(range);
        }
      }
    });
  });

  // Add rich text functionality to specific fields in this table
  addRichTextFunctionality();
}

// Function to add event listeners to row elements
function addRowEventListeners(row) {
  const cells = row.querySelectorAll('td[contenteditable="true"]');
  cells.forEach((cell) => {
    addCellEventListeners(cell);
  });
}

// Auto-save configuration
let autoSaveTimeout = null;
let lastSaveTime = 0;
const AUTO_SAVE_DELAY = 1000; // 1 second delay after user stops typing
const MIN_SAVE_INTERVAL = 2000; // Minimum 2 seconds between saves

// Auto-save status indicator functions
function showAutoSaveStatus(status, message) {
  const statusElement = document.getElementById("auto-save-status");
  if (!statusElement) return;

  const iconElement = statusElement.querySelector(".material-icons");
  const textElement = statusElement.querySelector(".status-text");

  // Remove existing status classes
  statusElement.classList.remove("saving", "saved");

  // Update content based on status
  switch (status) {
    case "saving":
      statusElement.classList.add("saving");
      iconElement.textContent = "cloud_upload";
      textElement.textContent = message || "Saving...";
      break;
    case "saved":
      statusElement.classList.add("saved");
      iconElement.textContent = "cloud_done";
      textElement.textContent = message || "Auto-saved";
      break;
    default:
      iconElement.textContent = "cloud_done";
      textElement.textContent = message || "Auto-saved";
  }

  // Show the status indicator
  statusElement.classList.add("show");

  // Hide after 3 seconds for saved status
  if (status === "saved") {
    setTimeout(() => {
      statusElement.classList.remove("show");
    }, 3000);
  }
}

// Enhanced auto-save function
function triggerAutoSave() {
  const now = Date.now();

  // Clear existing timeout
  if (autoSaveTimeout) {
    clearTimeout(autoSaveTimeout);
  }

  // Show saving status immediately
  showAutoSaveStatus("saving");

  // Set new timeout for auto-save
  autoSaveTimeout = setTimeout(() => {
    // Check if enough time has passed since last save
    if (now - lastSaveTime >= MIN_SAVE_INTERVAL) {
      saveData(false); // Silent save
      lastSaveTime = Date.now();
      showAutoSaveStatus("saved");
      console.log("Auto-saved at:", new Date().toLocaleTimeString());
    } else {
      // If we can't save yet, just hide the saving indicator
      const statusElement = document.getElementById("auto-save-status");
      if (statusElement) {
        statusElement.classList.remove("show");
      }
    }
  }, AUTO_SAVE_DELAY);
}

// Function to add event listeners to individual cells
function addCellEventListeners(cell) {
  // Handle Enter key to move to next row
  cell.addEventListener("keydown", function (e) {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      const currentRow = this.parentElement;
      const currentCellIndex = Array.from(currentRow.children).indexOf(this);
      const nextRow = currentRow.nextElementSibling;

      if (nextRow) {
        const nextCell = nextRow.children[currentCellIndex];
        if (nextCell) {
          nextCell.focus();
        }
      }
    }

    // Handle Tab key to move to next cell
    if (e.key === "Tab") {
      e.preventDefault();
      const currentRow = this.parentElement;
      const currentCellIndex = Array.from(currentRow.children).indexOf(this);
      const nextCell = currentRow.children[currentCellIndex + 1];

      if (nextCell) {
        nextCell.focus();
      } else {
        // Move to first cell of next row
        const nextRow = currentRow.nextElementSibling;
        if (nextRow) {
          nextRow.children[0].focus();
        }
      }
    }
  });

  // Enhanced auto-save functionality
  cell.addEventListener("input", function () {
    // Trigger auto-save on input (typing)
    triggerAutoSave();
  });

  cell.addEventListener("blur", function () {
    // Immediate save when field loses focus
    triggerAutoSave();
  });

  cell.addEventListener("paste", function () {
    // Auto-save after paste operations
    setTimeout(() => triggerAutoSave(), 100);
  });
}

// Function to save data
async function saveData(showAlert = true) {
  const data = extractAllData();
  localStorage.setItem("executiveSummaryData", JSON.stringify(data));

  console.log("Data saved to localStorage:", data);
  console.log("localStorage size:", JSON.stringify(data).length, "characters");

  if (showAlert) {
    // Automatically save as a named report with auto-generated title
    const saved = await saveReportWithAutoTitle();
    if (saved) {
      await showCustomAlert("Report saved successfully!", "Success", "success");
    }
  }
}

// Function to extract all data from tables
function extractAllData() {
  const data = {};
  const sections = document.querySelectorAll(".main-section");

  sections.forEach((section) => {
    const sectionId = section.id;

    // Skip view-only sections (those with IDs starting with "view-")
    if (sectionId.startsWith("view-")) {
      console.log(`Skipping view-only section: ${sectionId}`);
      return;
    }

    const sectionTitle =
      section.querySelector(".section-header h2").textContent;
    data[sectionId] = {
      title: sectionTitle,
      subsections: {},
    };

    const subsections = section.querySelectorAll(".subsection");
    subsections.forEach((subsection) => {
      const subsectionId = subsection.id;

      // Skip view-only subsections (those with IDs starting with "view-")
      if (subsectionId.startsWith("view-")) {
        console.log(`Skipping view-only subsection: ${subsectionId}`);
        return;
      }

      const subsectionTitle = subsection.querySelector(
        ".subsection-header h3"
      ).textContent;
      // Look for traditional tables and div-based tables
      const tables = subsection.querySelectorAll(".data-table");
      const performanceTables = subsection.querySelectorAll(
        ".performance-trends-table"
      );
      const abTestTables = subsection.querySelectorAll(".ab-test-table");
      const newlyOnboardedTables = subsection.querySelectorAll(
        ".newly-onboarded-table"
      );
      const inProgressTables =
        subsection.querySelectorAll(".in-progress-table");
      const npsTables = subsection.querySelectorAll(".nps-table");
      const churnedTables = subsection.querySelectorAll(".churned-table");
      const noticeTables = subsection.querySelectorAll(".notice-table");
      const publisherUpdatesTables = subsection.querySelectorAll(
        ".publisher-updates-table"
      );

      data[sectionId].subsections[subsectionId] = {
        title: subsectionTitle,
        tables: [],
      };

      // Helper function to check if a table has meaningful content
      function hasTableContent(tableElement, tableType = "traditional") {
        if (tableType === "traditional") {
          const rows = tableElement.querySelectorAll("tbody tr");
          return (
            rows.length > 0 &&
            Array.from(rows).some((row) => {
              const cells = row.querySelectorAll("td");
              return Array.from(cells).some(
                (cell) => cell.textContent.trim() !== ""
              );
            })
          );
        } else {
          // For div-based tables, check for rows with content
          const rows = tableElement.querySelectorAll("[class*='-row']");
          return (
            rows.length > 0 &&
            Array.from(rows).some((row) => {
              return row.textContent.trim() !== "";
            })
          );
        }
      }

      // Process traditional tables (only if they have content)
      tables.forEach((table) => {
        // Skip empty tables
        if (!hasTableContent(table, "traditional")) {
          return;
        }

        const tableData = {
          headers: [],
          rows: [],
        };

        // Check if this is a publisher updates table
        const isPublisherUpdatesTable = subsectionTitle.includes(
          "PUBLISHER ISSUES/UPDATES"
        );

        // Extract headers
        const headers = table.querySelectorAll("thead th");
        headers.forEach((header) => {
          // Check if this header has a week picker
          const weekPicker = header.querySelector(".week-picker");
          if (weekPicker) {
            // Extract the week range and description
            const displayValue = weekPicker.getAttribute("data-display-value");
            const headerText = header.textContent;
            const description = headerText.includes("Prior Wednesday")
              ? " (Prior Wednesday and Prior 7 Days)"
              : headerText.includes("Most Recent")
              ? " (Most Recent Wednesday and Prior 7 Days)"
              : "";
            tableData.headers.push(
              (displayValue || formatWeekRange(weekPicker.value)) + description
            );
          } else {
            tableData.headers.push(header.textContent);
          }
        });

        // Extract rows
        const rows = table.querySelectorAll("tbody tr");
        rows.forEach((row) => {
          const rowData = [];
          const cells = row.querySelectorAll("td");

          if (isPublisherUpdatesTable) {
            // Special handling for publisher updates table
            // We need to preserve the logical 6-column structure
            cells.forEach((cell, index) => {
              const colspan = parseInt(cell.getAttribute("colspan") || "1");

              // Extract cell content
              let cellContent = "";
              const dateWrapper = cell.querySelector(".date-field-wrapper");
              if (dateWrapper) {
                const hiddenInput =
                  dateWrapper.querySelector(".date-input-hidden");
                if (hiddenInput) {
                  const displayValue =
                    hiddenInput.getAttribute("data-display-value");
                  cellContent =
                    displayValue || formatDateForDisplay(hiddenInput.value);
                } else {
                  cellContent = cell.textContent.trim();
                }
              } else {
                const dateInput = cell.querySelector(".date-input");
                if (dateInput) {
                  const displayValue =
                    dateInput.getAttribute("data-display-value");
                  cellContent =
                    displayValue || formatDateForDisplay(dateInput.value);
                } else {
                  cellContent = cell.textContent.trim();
                }
              }

              // Add the cell content
              rowData.push(cellContent);

              // For publisher updates table, add empty strings to maintain 6-column structure
              for (let i = 1; i < colspan; i++) {
                rowData.push("");
              }
            });

            // Ensure we always have 6 columns for publisher updates table
            while (rowData.length < 6) {
              rowData.push("");
            }
          } else {
            // Standard handling for other tables
            cells.forEach((cell, index) => {
              const colspan = parseInt(cell.getAttribute("colspan") || "1");

              // Extract cell content
              let cellContent = "";
              const dateWrapper = cell.querySelector(".date-field-wrapper");
              if (dateWrapper) {
                const hiddenInput =
                  dateWrapper.querySelector(".date-input-hidden");
                if (hiddenInput) {
                  const displayValue =
                    hiddenInput.getAttribute("data-display-value");
                  cellContent =
                    displayValue || formatDateForDisplay(hiddenInput.value);
                } else {
                  cellContent = cell.textContent.trim();
                }
              } else {
                const dateInput = cell.querySelector(".date-input");
                if (dateInput) {
                  const displayValue =
                    dateInput.getAttribute("data-display-value");
                  cellContent =
                    displayValue || formatDateForDisplay(dateInput.value);
                } else {
                  cellContent = cell.textContent.trim();
                }
              }

              // Add the cell content
              rowData.push(cellContent);

              // If colspan > 1, add empty strings for the additional logical columns
              for (let i = 1; i < colspan; i++) {
                rowData.push("");
              }
            });
          }

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });

      // Process div-based performance trends tables (only if they have content)
      performanceTables.forEach((performanceTable) => {
        // Skip empty tables
        if (!hasTableContent(performanceTable, "div")) {
          return;
        }

        const tableData = {
          headers: [],
          rows: [],
        };

        // Extract headers from performance trends table
        const headers = performanceTable.querySelectorAll(
          ".performance-header-cell"
        );
        headers.forEach((header) => {
          // Check if this header has a week picker
          const weekPicker = header.querySelector(".week-picker");
          if (weekPicker) {
            // Extract the week range and description
            const displayValue = weekPicker.getAttribute("data-display-value");
            const headerText = header.textContent;
            const description = headerText.includes("Prior Wednesday")
              ? " (Prior Wednesday and Prior 7 Days)"
              : headerText.includes("Most Recent")
              ? " (Most Recent Wednesday and Prior 7 Days)"
              : "";
            tableData.headers.push(
              (displayValue || formatWeekRange(weekPicker.value)) + description
            );
          } else {
            tableData.headers.push(header.textContent);
          }
        });

        // Extract rows from performance trends table
        const rows = performanceTable.querySelectorAll(".performance-row");
        rows.forEach((row) => {
          const rowData = [];

          if (row.classList.contains("observations-row")) {
            // Handle observations row with merged cells
            const label = row.querySelector(".metric-label");
            const observationsCell = row.querySelector(".observations-cell");

            rowData.push(label ? label.textContent.trim() : "");
            rowData.push(
              observationsCell ? observationsCell.textContent.trim() : ""
            );
            // Add empty string for the third column since observations spans both data columns
            rowData.push("");
          } else {
            // Handle regular metric rows
            const label = row.querySelector(".metric-label");
            const dataCells = row.querySelectorAll(".metric-data");

            rowData.push(label ? label.textContent.trim() : "");
            dataCells.forEach((cell) => {
              rowData.push(cell.textContent.trim());
            });
          }

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });

      // Process div-based A/B test tables (only if they have content)
      abTestTables.forEach((abTestTable) => {
        // Skip empty tables
        if (!hasTableContent(abTestTable, "div")) {
          return;
        }

        const tableData = {
          headers: ["", "", "", "", "", ""], // A/B test tables have 6 logical columns
          rows: [],
        };

        // Extract rows from A/B test table
        const rows = abTestTable.querySelectorAll(".ab-test-row");
        rows.forEach((row) => {
          const rowData = [];

          if (row.classList.contains("dates-row")) {
            // Handle dates row with 6 columns
            const labels = row.querySelectorAll(".ab-label");
            const dateInputs = row.querySelectorAll(".ab-date-input");

            // Interleave labels and date inputs
            for (let i = 0; i < 3; i++) {
              if (labels[i]) {
                rowData.push(labels[i].textContent.trim());
              } else {
                rowData.push("");
              }

              if (dateInputs[i]) {
                // Extract date value from date input
                const dateInput = dateInputs[i].querySelector(
                  ".date-input, .date-input-hidden"
                );
                if (dateInput) {
                  const displayValue =
                    dateInput.getAttribute("data-display-value");
                  rowData.push(displayValue || dateInput.value || "");
                } else {
                  rowData.push(dateInputs[i].textContent.trim());
                }
              } else {
                rowData.push("");
              }
            }
          } else if (
            row.classList.contains("stakeholders-row") ||
            row.classList.contains("status-row")
          ) {
            // Handle stakeholders and status rows (span multiple columns)
            const label = row.querySelector(".ab-label");
            const cell = row.querySelector(
              ".stakeholders-cell, .test-status-cell"
            );

            rowData.push(label ? label.textContent.trim() : "");
            rowData.push(cell ? cell.textContent.trim() : "");
            // Add empty strings for remaining columns
            rowData.push("", "", "", "");
          } else {
            // Handle regular rows with 4 logical columns
            const allCells = row.querySelectorAll(
              ".ab-test-cell, .ab-label, .ab-data, .ab-data-wide, .ab-data-tier-wide"
            );

            // For regular rows, we expect: label1, data1, label2, data2
            if (allCells.length >= 4) {
              rowData.push(allCells[0] ? allCells[0].textContent.trim() : "");
              rowData.push(allCells[1] ? allCells[1].textContent.trim() : "");
              rowData.push(allCells[2] ? allCells[2].textContent.trim() : "");
              rowData.push(allCells[3] ? allCells[3].textContent.trim() : "");
            } else {
              // Fallback: try specific selectors
              const label1 = row.querySelector(".ab-label:first-child");
              const data1 = row.querySelector(".ab-data:first-child, .ab-data");
              const label2 = row.querySelector(".ab-label:last-child");
              const data2 = row.querySelector(
                ".ab-data-wide, .ab-data-tier-wide"
              );

              rowData.push(label1 ? label1.textContent.trim() : "");
              rowData.push(data1 ? data1.textContent.trim() : "");
              rowData.push(label2 ? label2.textContent.trim() : "");
              rowData.push(data2 ? data2.textContent.trim() : "");
            }
            // Add empty strings for remaining columns to maintain 6-column structure
            rowData.push("", "");
          }

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });

      // Process div-based newly onboarded tables (only if they have content)
      newlyOnboardedTables.forEach((newlyOnboardedTable) => {
        // Skip empty tables
        if (!hasTableContent(newlyOnboardedTable, "div")) {
          return;
        }

        const tableData = {
          headers: ["", "", "", ""], // Newly onboarded tables have 4 logical columns
          rows: [],
        };

        // Extract rows from newly onboarded table
        const rows = newlyOnboardedTable.querySelectorAll(
          ".newly-onboarded-row"
        );
        rows.forEach((row) => {
          const rowData = [];

          if (row.classList.contains("demand-status-row")) {
            // Handle demand status row (spans multiple columns)
            const label = row.querySelector(".publisher-label");
            const cell = row.querySelector(".demand-status-cell");

            rowData.push(label ? label.textContent.trim() : "");
            rowData.push(cell ? cell.textContent.trim() : "");
            // Add empty strings for remaining columns
            rowData.push("", "");
          } else {
            // Handle regular rows with 4 columns
            const label1 = row.querySelector(".publisher-label:first-child");
            const data1 = row.querySelector(".publisher-data");
            const label2 = row.querySelector(".publisher-label:last-child");
            const dataInput = row.querySelector(".publisher-data-input");

            rowData.push(label1 ? label1.textContent.trim() : "");
            rowData.push(data1 ? data1.textContent.trim() : "");
            rowData.push(label2 ? label2.textContent.trim() : "");

            // Handle date input or regular span
            if (dataInput) {
              const dateInput = dataInput.querySelector(
                ".date-input, .date-input-hidden"
              );
              if (dateInput) {
                const displayValue =
                  dateInput.getAttribute("data-display-value");
                rowData.push(displayValue || dateInput.value || "");
              } else {
                const span = dataInput.querySelector(
                  'span[contenteditable="true"]'
                );
                rowData.push(
                  span ? span.textContent.trim() : dataInput.textContent.trim()
                );
              }
            } else {
              rowData.push("");
            }
          }

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });

      // Process div-based in-progress tables (only if they have content)
      inProgressTables.forEach((inProgressTable) => {
        // Skip empty tables
        if (!hasTableContent(inProgressTable, "div")) {
          return;
        }

        const tableData = {
          headers: ["", "", "", ""], // In-progress tables have 4 logical columns
          rows: [],
        };

        // Extract rows from in-progress table
        const rows = inProgressTable.querySelectorAll(".in-progress-row");
        rows.forEach((row) => {
          const rowData = [];

          if (row.classList.contains("status-row")) {
            // Handle status row (spans multiple columns)
            const label = row.querySelector(".progress-label");
            const cell = row.querySelector(".progress-status-cell");

            rowData.push(label ? label.textContent.trim() : "");
            rowData.push(cell ? cell.textContent.trim() : "");
            // Add empty strings for remaining columns
            rowData.push("", "");
          } else {
            // Handle regular rows with 4 columns
            const label1 = row.querySelector(".progress-label:first-child");
            const data1 = row.querySelector(".progress-data");
            const label2 = row.querySelector(".progress-label:last-child");
            const dataInput = row.querySelector(".progress-data-input");

            rowData.push(label1 ? label1.textContent.trim() : "");
            rowData.push(data1 ? data1.textContent.trim() : "");
            rowData.push(label2 ? label2.textContent.trim() : "");

            // Handle date input or regular span
            if (dataInput) {
              const dateInput = dataInput.querySelector(
                ".date-input, .date-input-hidden"
              );
              if (dateInput) {
                const displayValue =
                  dateInput.getAttribute("data-display-value");
                rowData.push(displayValue || dateInput.value || "");
              } else {
                const span = dataInput.querySelector(
                  'span[contenteditable="true"]'
                );
                rowData.push(
                  span ? span.textContent.trim() : dataInput.textContent.trim()
                );
              }
            } else {
              rowData.push("");
            }
          }

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });

      // Process div-based NPS tables (only if they have content)
      npsTables.forEach((npsTable) => {
        // Skip empty tables
        if (!hasTableContent(npsTable, "div")) {
          return;
        }

        const tableData = {
          headers: ["", "MTD", "QTD", "YTD", ""], // NPS tables have 5 columns
          rows: [],
        };

        // Extract rows from NPS table
        const rows = npsTable.querySelectorAll(".nps-row");
        rows.forEach((row) => {
          const rowData = [];

          const label = row.querySelector(".nps-label");
          const scores = row.querySelectorAll(".nps-score");

          rowData.push(label ? label.textContent.trim() : "");
          scores.forEach((score) => {
            rowData.push(score ? score.textContent.trim() : "");
          });
          // Add empty string for actions column
          rowData.push("");

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });

      // Process div-based churned tables (only if they have content)
      churnedTables.forEach((churnedTable) => {
        // Skip empty tables
        if (!hasTableContent(churnedTable, "div")) {
          return;
        }

        const tableData = {
          headers: [
            "PUBLISHER",
            "SITE(S)",
            "AVG. MONTHLY REV.",
            "DARK DATE & REASON",
            "",
          ], // Churned tables have 5 columns (4 data + 1 actions)
          rows: [],
        };

        // Extract rows from churned table
        const rows = churnedTable.querySelectorAll(".churned-row");
        rows.forEach((row) => {
          const rowData = [];

          const publisher = row.querySelector(".churned-publisher");
          const site = row.querySelector(".churned-site");
          const revenue = row.querySelector(".churned-revenue");
          const reason = row.querySelector(".churned-reason");

          rowData.push(publisher ? publisher.textContent.trim() : "");
          rowData.push(site ? site.textContent.trim() : "");
          rowData.push(revenue ? revenue.textContent.trim() : "");
          rowData.push(reason ? reason.textContent.trim() : "");
          rowData.push(""); // Empty column for actions

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });

      // Process div-based notice tables (only if they have content)
      noticeTables.forEach((noticeTable) => {
        // Skip empty tables
        if (!hasTableContent(noticeTable, "div")) {
          return;
        }

        const tableData = {
          headers: ["PUBLISHER", "AVG. MONTHLY REV.", "NOTICE DATE & REASON"], // Notice tables have 3 columns
          rows: [],
        };

        // Extract rows from notice table
        const rows = noticeTable.querySelectorAll(".notice-row");
        rows.forEach((row) => {
          const rowData = [];

          const publisher = row.querySelector(".notice-publisher");
          const revenue = row.querySelector(".notice-revenue");
          const reason = row.querySelector(".notice-reason");

          rowData.push(publisher ? publisher.textContent.trim() : "");
          rowData.push(revenue ? revenue.textContent.trim() : "");
          rowData.push(reason ? reason.textContent.trim() : "");

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });

      // Process div-based publisher updates tables (only if they have content)
      publisherUpdatesTables.forEach((publisherUpdatesTable) => {
        // Skip empty tables
        if (!hasTableContent(publisherUpdatesTable, "div")) {
          return;
        }

        const tableData = {
          headers: ["", "", "", "", "", ""], // Publisher updates tables have 6 logical columns
          rows: [],
        };

        // Extract rows from publisher updates table
        const rows = publisherUpdatesTable.querySelectorAll(".updates-row");
        rows.forEach((row, rowIndex) => {
          const rowData = ["", "", "", "", "", ""]; // Initialize 6-column structure

          if (rowIndex === 0) {
            // Row 1: PUBLISHER | name | TIER | 1 | AVG. MONTHLY REV. | 20,000
            const cells = row.querySelectorAll(".updates-label, .updates-data");
            cells.forEach((cell, cellIndex) => {
              if (cellIndex < 6) {
                rowData[cellIndex] = cell.textContent.trim();
              }
            });
          } else if (rowIndex === 1) {
            // Row 2: STAGE | Onboarding (spans 3) | CSM | Kristen
            const label1 = row.querySelector(".updates-label");
            const data1 = row.querySelector(".updates-data");
            const label2 = row.querySelector(".updates-label:last-of-type");
            const data2 = row.querySelector(".updates-data:last-of-type");

            rowData[0] = label1 ? label1.textContent.trim() : "";
            rowData[1] = data1 ? data1.textContent.trim() : "";
            rowData[3] = label2 ? label2.textContent.trim() : "";
            rowData[4] = data2 ? data2.textContent.trim() : "";
          } else if (rowIndex === 2) {
            // Row 3: STAKEHOLDERS | list (spans 3) | LAST COMM. DATE | date
            const label1 = row.querySelector(".updates-label");
            const data1 = row.querySelector(".updates-data");
            const label2 = row.querySelector(".updates-label:last-of-type");
            const dateCell = row.querySelector(".updates-data:last-of-type");

            rowData[0] = label1 ? label1.textContent.trim() : "";
            rowData[1] = data1 ? data1.textContent.trim() : "";
            rowData[3] = label2 ? label2.textContent.trim() : "";

            // Handle date input
            if (dateCell) {
              const dateInput = dateCell.querySelector(
                ".date-input, .date-input-hidden"
              );
              if (dateInput) {
                const displayValue =
                  dateInput.getAttribute("data-display-value");
                rowData[4] = displayValue || dateInput.value || "";
              } else {
                rowData[4] = dateCell.textContent.trim();
              }
            }
          } else if (rowIndex === 3) {
            // Row 4: ISSUE(S)/UPDATE(S) | content (spans 5)
            const label = row.querySelector(".updates-label");
            const issues = row.querySelector(".updates-issues");

            rowData[0] = label ? label.textContent.trim() : "";
            rowData[1] = issues ? issues.textContent.trim() : "";
          } else if (rowIndex === 4) {
            // Row 5: NEXT STEPS | content (spans 5)
            const label = row.querySelector(".updates-label");
            const steps = row.querySelector(".updates-steps");

            rowData[0] = label ? label.textContent.trim() : "";
            rowData[1] = steps ? steps.textContent.trim() : "";
          }

          tableData.rows.push(rowData);
        });

        data[sectionId].subsections[subsectionId].tables.push(tableData);
      });
    });
  });

  return data;
}

// Function to clear all tables without showing confirmation
function clearAllTablesOnly() {
  const tableWrappers = document.querySelectorAll(".table-wrapper");
  console.log(`Clearing ${tableWrappers.length} existing table(s)`);
  tableWrappers.forEach((wrapper) => {
    wrapper.remove();
  });
}

// Function to add default tables to ensure the app is always usable
function addDefaultTables() {
  console.log("Adding default tables...");

  const tablesToAdd = [
    { subsection: "wow-performance", type: "performance-trends" },
    { subsection: "ab-test", type: "ab-test-updates" },
    { subsection: "newly-onboarded", type: "newly-onboarded-publishers" },
    { subsection: "in-progress", type: "in-progress-publishers" },
    { subsection: "nps", type: "nps-data" },
    { subsection: "recently-churned", type: "churned-data" },
    { subsection: "gave-notice", type: "notice-data" },
    { subsection: "publisher-issues", type: "publisher-updates" },
  ];

  tablesToAdd.forEach(({ subsection, type }) => {
    try {
      addTable(subsection, type);
    } catch (error) {
      console.error(`Failed to add table ${subsection}:`, error);
    }
  });

  console.log("Default tables addition completed");

  // Verify tables were actually created
  setTimeout(() => {
    const containers = [
      "wow-performance-tables",
      "ab-test-tables",
      "newly-onboarded-tables",
      "in-progress-tables",
      "nps-tables",
      "recently-churned-tables",
      "gave-notice-tables",
      "publisher-issues-tables",
    ];

    console.log("=== TABLE VERIFICATION ===");
    containers.forEach((containerId) => {
      const container = document.getElementById(containerId);
      const tableCount = container ? container.children.length : 0;
      console.log(`${containerId}: ${tableCount} tables`);
      if (tableCount === 0) {
        console.warn(`⚠️ No tables found in ${containerId}!`);
      }
    });
    console.log("=== END VERIFICATION ===");
  }, 200);
}

// Function to ensure all sections have at least one table
function ensureAllSectionsHaveTables() {
  console.log("Ensuring all sections have tables...");

  const requiredTables = [
    {
      subsection: "wow-performance",
      type: "performance-trends",
      containerId: "wow-performance-tables",
    },
    {
      subsection: "ab-test",
      type: "ab-test-updates",
      containerId: "ab-test-tables",
    },
    {
      subsection: "newly-onboarded",
      type: "newly-onboarded-publishers",
      containerId: "newly-onboarded-tables",
    },
    {
      subsection: "in-progress",
      type: "in-progress-publishers",
      containerId: "in-progress-tables",
    },
    { subsection: "nps", type: "nps-data", containerId: "nps-tables" },
    {
      subsection: "recently-churned",
      type: "churned-data",
      containerId: "recently-churned-tables",
    },
    {
      subsection: "gave-notice",
      type: "notice-data",
      containerId: "gave-notice-tables",
    },
    {
      subsection: "publisher-issues",
      type: "publisher-updates",
      containerId: "publisher-issues-tables",
    },
  ];

  let tablesAdded = 0;

  requiredTables.forEach(({ subsection, type, containerId }) => {
    const container = document.getElementById(containerId);
    if (container && container.children.length === 0) {
      console.log(`Adding missing table for ${subsection}`);
      try {
        addTable(subsection, type);
        tablesAdded++;
      } catch (error) {
        console.error(`Failed to add missing table for ${subsection}:`, error);
      }
    }
  });

  if (tablesAdded > 0) {
    console.log(`Added ${tablesAdded} missing tables`);
  } else {
    console.log("All sections already have tables");
  }
}

// Function to restore tables for a specific subsection
function restoreTablesForSubsection(subsectionId, tables) {
  // Map subsection IDs to their table types
  const subsectionToTableType = {
    "wow-performance": "performance-trends",
    "ab-test": "ab-test-updates",
    "newly-onboarded": "newly-onboarded-publishers",
    "in-progress": "in-progress-publishers",
    nps: "nps-data",
    "recently-churned": "churned-data",
    "gave-notice": "notice-data",
    "publisher-issues": "publisher-updates",
  };

  const tableType = subsectionToTableType[subsectionId];
  if (!tableType) {
    console.warn(`Unknown subsection ID: ${subsectionId}`);
    return;
  }

  // For NPS and recently-churned sections, only create one table and merge all data
  if (subsectionId === "nps" || subsectionId === "recently-churned") {
    // Only create one table for these sections
    addTable(subsectionId, tableType);

    // Get the newly created table
    const container = document.getElementById(`${subsectionId}-tables`);
    const tableWrappers = container.querySelectorAll(".table-wrapper");
    const newTableWrapper = tableWrappers[tableWrappers.length - 1];

    if (newTableWrapper && tables.length > 0) {
      // Use the first table's data (they should all be the same structure anyway)
      populateTableWithData(newTableWrapper, tables[0], tableType);
    }
  } else {
    // For other sections, create multiple tables as before
    tables.forEach((tableData, index) => {
      // Add a new table
      addTable(subsectionId, tableType);

      // Get the newly created table (it will be the last one)
      const container = document.getElementById(`${subsectionId}-tables`);
      const tableWrappers = container.querySelectorAll(".table-wrapper");
      const newTableWrapper = tableWrappers[tableWrappers.length - 1];

      if (newTableWrapper) {
        // Populate the table with saved data
        populateTableWithData(newTableWrapper, tableData, tableType);
      }
    });
  }
}

// Function to populate a table with saved data
function populateTableWithData(tableWrapper, tableData, tableType) {
  // Check if this is a performance trends table (div-based) or traditional table
  const performanceTable = tableWrapper.querySelector(
    ".performance-trends-table"
  );
  const table = tableWrapper.querySelector(".data-table");

  if (performanceTable && tableType === "performance-trends") {
    // Handle div-based performance trends table
    populatePerformanceTrendsWithData(performanceTable, tableData);
    return;
  }

  const abTestTable = tableWrapper.querySelector(".ab-test-table");
  if (abTestTable && tableType === "ab-test-updates") {
    // Handle div-based A/B test table
    populateABTestWithData(abTestTable, tableData);
    return;
  }

  const newlyOnboardedTable = tableWrapper.querySelector(
    ".newly-onboarded-table"
  );
  if (newlyOnboardedTable && tableType === "newly-onboarded-publishers") {
    // Handle div-based newly onboarded table
    populateNewlyOnboardedWithData(newlyOnboardedTable, tableData);
    return;
  }

  const inProgressTable = tableWrapper.querySelector(".in-progress-table");
  if (inProgressTable && tableType === "in-progress-publishers") {
    // Handle div-based in-progress table
    populateInProgressWithData(inProgressTable, tableData);
    return;
  }

  const npsTable = tableWrapper.querySelector(".nps-table");
  if (npsTable && tableType === "nps-data") {
    // Handle div-based NPS table
    populateNPSWithData(npsTable, tableData);
    return;
  }

  const churnedTable = tableWrapper.querySelector(".churned-table");
  if (churnedTable && tableType === "churned-data") {
    // Handle div-based churned table
    populateChurnedWithData(churnedTable, tableData);
    return;
  }

  const noticeTable = tableWrapper.querySelector(".notice-table");
  if (noticeTable && tableType === "notice-data") {
    // Handle div-based notice table
    populateNoticeWithData(noticeTable, tableData);
    return;
  }

  const publisherUpdatesTable = tableWrapper.querySelector(
    ".publisher-updates-table"
  );
  if (publisherUpdatesTable && tableType === "publisher-updates") {
    // Handle div-based publisher updates table
    populatePublisherUpdatesWithData(publisherUpdatesTable, tableData);
    return;
  }

  if (!table || !tableData) return;

  // Populate headers if they have editable content (like week pickers)
  const headerCells = table.querySelectorAll("thead th");
  if (tableData.headers && headerCells.length > 0) {
    headerCells.forEach((headerCell, index) => {
      if (
        index < tableData.headers.length &&
        headerCell.classList.contains("editable-header")
      ) {
        // This is an editable header (week picker)
        const headerText = tableData.headers[index];
        const weekRange = headerText.split(" (")[0]; // Extract just the date range
        const description = headerText.includes("Prior Wednesday")
          ? " (Prior Wednesday and Prior 7 Days)"
          : headerText.includes("Most Recent")
          ? " (Most Recent Wednesday and Prior 7 Days)"
          : "";

        headerCell.innerHTML = `${weekRange}${description}<br>${createWeekPicker(
          weekRange,
          "header-week-picker"
        )}`;
      }
    });
  }

  // Populate table body
  const tbody = table.querySelector("tbody");
  if (!tbody || !tableData.rows) return;

  const rows = tbody.querySelectorAll("tr");

  // If we need more rows than exist, add them
  while (rows.length < tableData.rows.length) {
    addRowToTable(tableWrapper, tableType);
  }

  // Get updated rows after potentially adding new ones
  const updatedRows = tbody.querySelectorAll("tr");

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < updatedRows.length) {
      const row = updatedRows[rowIndex];
      const cells = row.querySelectorAll("td");

      cells.forEach((cell, cellIndex) => {
        if (cellIndex < rowData.length && rowData[cellIndex] !== undefined) {
          const cellValue = rowData[cellIndex];

          // Check if this cell contains a date input
          const dateWrapper = cell.querySelector(".date-field-wrapper");
          if (dateWrapper) {
            const hiddenInput = dateWrapper.querySelector(".date-input-hidden");
            const dateText = dateWrapper.querySelector(".date-text");
            if (hiddenInput && dateText && cellValue) {
              // Convert display format back to input format
              const inputValue = formatDateForInput(cellValue);
              hiddenInput.value = inputValue;
              hiddenInput.setAttribute("data-display-value", cellValue);
              dateText.textContent = cellValue;
            }
          } else if (
            cell.contentEditable === "true" ||
            cell.querySelector('[contenteditable="true"]')
          ) {
            // Regular editable cell
            const editableElement =
              cell.contentEditable === "true"
                ? cell
                : cell.querySelector('[contenteditable="true"]');
            if (editableElement) {
              editableElement.textContent = cellValue;
            }
          }
        }
      });
    }
  });

  // Also populate mobile cards if they exist
  // populateMobileCardsWithData(tableWrapper, tableData, tableType);
  // TODO: Implement mobile card population if needed
}

// Function to populate div-based performance trends table with saved data
function populatePerformanceTrendsWithData(performanceTable, tableData) {
  if (!performanceTable || !tableData) return;

  // Populate headers if they have editable content (like week pickers)
  const headerCells = performanceTable.querySelectorAll(
    ".performance-header-cell"
  );
  if (tableData.headers && headerCells.length > 0) {
    headerCells.forEach((headerCell, index) => {
      if (
        index < tableData.headers.length &&
        headerCell.classList.contains("editable-header")
      ) {
        // This is an editable header (week picker)
        const headerText = tableData.headers[index];
        const weekRange = headerText.split(" (")[0]; // Extract just the date range
        const description = headerText.includes("Prior Wednesday")
          ? " (Prior Wednesday and Prior 7 Days)"
          : headerText.includes("Most Recent")
          ? " (Most Recent Wednesday and Prior 7 Days)"
          : "";

        headerCell.innerHTML = `${weekRange}${description}<br>${createWeekPicker(
          weekRange,
          "header-week-picker"
        )}`;
      }
    });
  }

  // Populate performance rows
  const rows = performanceTable.querySelectorAll(".performance-row");
  if (!tableData.rows) return;

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      if (row.classList.contains("observations-row")) {
        // Handle observations row
        const label = row.querySelector(".metric-label");
        const observationsCell = row.querySelector(".observations-cell");

        if (label && rowData[0] !== undefined) {
          label.textContent = rowData[0];
        }
        if (observationsCell && rowData[1] !== undefined) {
          observationsCell.textContent = rowData[1];
        }
      } else {
        // Handle regular metric rows
        const label = row.querySelector(".metric-label");
        const dataCells = row.querySelectorAll(".metric-data");

        if (label && rowData[0] !== undefined) {
          label.textContent = rowData[0];
        }

        dataCells.forEach((cell, cellIndex) => {
          const dataIndex = cellIndex + 1; // Skip the label column
          if (dataIndex < rowData.length && rowData[dataIndex] !== undefined) {
            cell.textContent = rowData[dataIndex];
          }
        });
      }
    }
  });
}

// Function to populate div-based A/B test table with saved data
function populateABTestWithData(abTestTable, tableData) {
  if (!abTestTable || !tableData) return;

  // Populate A/B test rows
  const rows = abTestTable.querySelectorAll(".ab-test-row");
  if (!tableData.rows) return;

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      if (row.classList.contains("dates-row")) {
        // Handle dates row with 6 columns (3 labels + 3 date inputs)
        const labels = row.querySelectorAll(".ab-label");
        const dateInputs = row.querySelectorAll(".ab-date-input");

        // Populate labels and date inputs alternately
        for (let i = 0; i < 3; i++) {
          const labelIndex = i * 2;
          const dateIndex = labelIndex + 1;

          if (labels[i] && rowData[labelIndex] !== undefined) {
            labels[i].textContent = rowData[labelIndex];
          }

          if (dateInputs[i] && rowData[dateIndex] !== undefined) {
            // Handle date input population
            const dateInput = dateInputs[i].querySelector(
              ".date-input, .date-input-hidden"
            );
            if (dateInput && rowData[dateIndex]) {
              const inputValue = formatDateForInput(rowData[dateIndex]);
              dateInput.value = inputValue;
              dateInput.setAttribute("data-display-value", rowData[dateIndex]);

              const dateText = dateInputs[i].querySelector(".date-text");
              if (dateText) {
                dateText.textContent = rowData[dateIndex];
              }
            }
          }
        }
      } else if (
        row.classList.contains("stakeholders-row") ||
        row.classList.contains("status-row")
      ) {
        // Handle stakeholders and status rows
        const label = row.querySelector(".ab-label");
        const cell = row.querySelector(".stakeholders-cell, .test-status-cell");

        if (label && rowData[0] !== undefined) {
          label.textContent = rowData[0];
        }
        if (cell && rowData[1] !== undefined) {
          cell.textContent = rowData[1];
        }
      } else {
        // Handle regular rows with 4 logical columns
        const label1 = row.querySelector(".ab-label:first-child");
        const data1 = row.querySelector(".ab-data:first-child");
        const label2 = row.querySelector(".ab-label:last-child");
        const data2 = row.querySelector(".ab-data-wide");

        if (label1 && rowData[0] !== undefined) {
          label1.textContent = rowData[0];
        }
        if (data1 && rowData[1] !== undefined) {
          data1.textContent = rowData[1];
        }
        if (label2 && rowData[2] !== undefined) {
          label2.textContent = rowData[2];
        }
        if (data2 && rowData[3] !== undefined) {
          data2.textContent = rowData[3];
        }
      }
    }
  });
}

// Function to populate div-based newly onboarded table with saved data
function populateNewlyOnboardedWithData(newlyOnboardedTable, tableData) {
  if (!newlyOnboardedTable || !tableData) return;

  // Populate newly onboarded rows
  const rows = newlyOnboardedTable.querySelectorAll(".newly-onboarded-row");
  if (!tableData.rows) return;

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      if (row.classList.contains("demand-status-row")) {
        // Handle demand status row
        const label = row.querySelector(".publisher-label");
        const cell = row.querySelector(".demand-status-cell");

        if (label && rowData[0] !== undefined) {
          label.textContent = rowData[0];
        }
        if (cell && rowData[1] !== undefined) {
          cell.textContent = rowData[1];
        }
      } else {
        // Handle regular rows with 4 columns
        const label1 = row.querySelector(".publisher-label:first-child");
        const data1 = row.querySelector(".publisher-data");
        const label2 = row.querySelector(".publisher-label:last-child");
        const dataInput = row.querySelector(".publisher-data-input");

        if (label1 && rowData[0] !== undefined) {
          label1.textContent = rowData[0];
        }
        if (data1 && rowData[1] !== undefined) {
          data1.textContent = rowData[1];
        }
        if (label2 && rowData[2] !== undefined) {
          label2.textContent = rowData[2];
        }

        if (dataInput && rowData[3] !== undefined) {
          // Handle date input or regular span
          const dateInput = dataInput.querySelector(
            ".date-input, .date-input-hidden"
          );
          if (dateInput && rowData[3]) {
            const inputValue = formatDateForInput(rowData[3]);
            dateInput.value = inputValue;
            dateInput.setAttribute("data-display-value", rowData[3]);

            const dateText = dataInput.querySelector(".date-text");
            if (dateText) {
              dateText.textContent = rowData[3];
            }
          } else {
            const span = dataInput.querySelector(
              'span[contenteditable="true"]'
            );
            if (span) {
              span.textContent = rowData[3];
            }
          }
        }
      }
    }
  });
}

// Function to populate div-based in-progress table with saved data
function populateInProgressWithData(inProgressTable, tableData) {
  if (!inProgressTable || !tableData) return;

  // Populate in-progress rows
  const rows = inProgressTable.querySelectorAll(".in-progress-row");
  if (!tableData.rows) return;

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      if (row.classList.contains("status-row")) {
        // Handle status row
        const label = row.querySelector(".progress-label");
        const cell = row.querySelector(".progress-status-cell");

        if (label && rowData[0] !== undefined) {
          label.textContent = rowData[0];
        }
        if (cell && rowData[1] !== undefined) {
          cell.textContent = rowData[1];
        }
      } else {
        // Handle regular rows with 4 columns
        const label1 = row.querySelector(".progress-label:first-child");
        const data1 = row.querySelector(".progress-data");
        const label2 = row.querySelector(".progress-label:last-child");
        const dataInput = row.querySelector(".progress-data-input");

        if (label1 && rowData[0] !== undefined) {
          label1.textContent = rowData[0];
        }
        if (data1 && rowData[1] !== undefined) {
          data1.textContent = rowData[1];
        }
        if (label2 && rowData[2] !== undefined) {
          label2.textContent = rowData[2];
        }

        if (dataInput && rowData[3] !== undefined) {
          // Handle date input or regular span
          const dateInput = dataInput.querySelector(
            ".date-input, .date-input-hidden"
          );
          if (dateInput && rowData[3]) {
            const inputValue = formatDateForInput(rowData[3]);
            dateInput.value = inputValue;
            dateInput.setAttribute("data-display-value", rowData[3]);

            const dateText = dataInput.querySelector(".date-text");
            if (dateText) {
              dateText.textContent = rowData[3];
            }
          } else {
            const span = dataInput.querySelector(
              'span[contenteditable="true"]'
            );
            if (span) {
              span.textContent = rowData[3];
            }
          }
        }
      }
    }
  });
}

// Function to populate div-based NPS table with saved data
function populateNPSWithData(npsTable, tableData) {
  if (!npsTable || !tableData) return;

  // Populate NPS rows
  const rows = npsTable.querySelectorAll(".nps-row");
  if (!tableData.rows) return;

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      const label = row.querySelector(".nps-label");
      const scores = row.querySelectorAll(".nps-score");

      if (label && rowData[0] !== undefined) {
        label.textContent = rowData[0];
      }

      scores.forEach((score, scoreIndex) => {
        if (rowData[scoreIndex + 1] !== undefined) {
          score.textContent = rowData[scoreIndex + 1];
        }
      });
    }
  });
}

// Function to populate div-based churned table with saved data
function populateChurnedWithData(churnedTable, tableData) {
  if (!churnedTable || !tableData) return;

  // Populate churned rows
  const rows = churnedTable.querySelectorAll(".churned-row");
  if (!tableData.rows) return;

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      const publisher = row.querySelector(".churned-publisher");
      const site = row.querySelector(".churned-site");
      const revenue = row.querySelector(".churned-revenue");
      const reason = row.querySelector(".churned-reason");

      if (publisher && rowData[0] !== undefined) {
        publisher.textContent = rowData[0];
      }
      if (site && rowData[1] !== undefined) {
        site.textContent = rowData[1];
      }
      if (revenue && rowData[2] !== undefined) {
        revenue.textContent = rowData[2];
      }
      if (reason && rowData[3] !== undefined) {
        reason.textContent = rowData[3];
      }
    }
  });
}

// Function to populate div-based notice table with saved data
function populateNoticeWithData(noticeTable, tableData) {
  if (!noticeTable || !tableData) return;

  // Populate notice rows
  const rows = noticeTable.querySelectorAll(".notice-row");
  if (!tableData.rows) return;

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      const publisher = row.querySelector(".notice-publisher");
      const revenue = row.querySelector(".notice-revenue");
      const reason = row.querySelector(".notice-reason");

      if (publisher && rowData[0] !== undefined) {
        publisher.textContent = rowData[0];
      }
      if (revenue && rowData[1] !== undefined) {
        revenue.textContent = rowData[1];
      }
      if (reason && rowData[2] !== undefined) {
        reason.textContent = rowData[2];
      }
    }
  });
}

// Function to populate div-based NPS table with saved data
function populateNPSWithData(npsTable, tableData) {
  if (!npsTable || !tableData) return;

  // Populate NPS rows
  const rows = npsTable.querySelectorAll(".nps-row");
  if (!tableData.rows) return;

  // Populate each row with data
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      const label = row.querySelector(".nps-label");
      const scores = row.querySelectorAll(".nps-score");

      if (label && rowData[0] !== undefined) {
        label.textContent = rowData[0];
      }

      scores.forEach((score, scoreIndex) => {
        if (rowData[scoreIndex + 1] !== undefined) {
          score.textContent = rowData[scoreIndex + 1];
        }
      });
    }
  });
}

// Function to populate div-based publisher updates table with saved data
function populatePublisherUpdatesWithData(publisherUpdatesTable, tableData) {
  if (!publisherUpdatesTable || !tableData) return;

  // Populate publisher updates rows
  const rows = publisherUpdatesTable.querySelectorAll(".updates-row");
  if (!tableData.rows) return;

  // Populate each row with data based on the complex structure
  tableData.rows.forEach((rowData, rowIndex) => {
    if (rowIndex < rows.length) {
      const row = rows[rowIndex];

      if (rowIndex === 0) {
        // Row 1: PUBLISHER | name | TIER | 1 | AVG. MONTHLY REV. | 20,000
        const cells = row.querySelectorAll(".updates-label, .updates-data");
        cells.forEach((cell, cellIndex) => {
          if (cellIndex < 6 && rowData[cellIndex] !== undefined) {
            if (cell.classList.contains("updates-data")) {
              cell.textContent = rowData[cellIndex];
            }
          }
        });
      } else if (rowIndex === 1) {
        // Row 2: STAGE | Onboarding (spans 3) | CSM | Kristen
        const data1 = row.querySelector(".updates-data");
        const data2 = row.querySelector(".updates-data:last-of-type");

        if (data1 && rowData[1] !== undefined) {
          data1.textContent = rowData[1];
        }
        if (data2 && rowData[4] !== undefined) {
          data2.textContent = rowData[4];
        }
      } else if (rowIndex === 2) {
        // Row 3: STAKEHOLDERS | list (spans 3) | LAST COMM. DATE | date
        const data1 = row.querySelector(".updates-data");
        const dateCell = row.querySelector(".updates-data:last-of-type");

        if (data1 && rowData[1] !== undefined) {
          data1.textContent = rowData[1];
        }

        // Handle date input
        if (dateCell && rowData[4] !== undefined) {
          const dateInput = dateCell.querySelector(
            ".date-input, .date-input-hidden"
          );
          if (dateInput) {
            dateInput.value = rowData[4];
            dateInput.setAttribute("data-display-value", rowData[4]);
            // Update display if it's a visible date input
            if (dateInput.classList.contains("date-input")) {
              dateInput.setAttribute("placeholder", rowData[4] || "mm/dd/yyyy");
            }
          } else {
            dateCell.textContent = rowData[4];
          }
        }
      } else if (rowIndex === 3) {
        // Row 4: ISSUE(S)/UPDATE(S) | content (spans 5)
        const issues = row.querySelector(".updates-issues");
        if (issues && rowData[1] !== undefined) {
          issues.textContent = rowData[1];
        }
      } else if (rowIndex === 4) {
        // Row 5: NEXT STEPS | content (spans 5)
        const steps = row.querySelector(".updates-steps");
        if (steps && rowData[1] !== undefined) {
          steps.textContent = rowData[1];
        }
      }
    }
  });
}

// Function to add a row to an existing table
function addRowToTable(tableWrapper, tableType) {
  const table = tableWrapper.querySelector(".data-table tbody");
  const headerCount = tableWrapper.querySelector(".data-table thead tr")
    .children.length;

  const newRow = document.createElement("tr");
  // Add data cells
  for (let i = 0; i < headerCount; i++) {
    const cell = document.createElement("td");

    // Check if this should be a date field
    const rowIndex = table.rows.length; // Current row index

    if (isDateField("", i, rowIndex, tableType)) {
      cell.innerHTML = createDateInput("");
    } else {
      cell.contentEditable = true;
      cell.textContent = "";
    }
    newRow.appendChild(cell);
  }

  table.appendChild(newRow);
  addRowEventListeners(newRow);
}

// Test data generation functions
function getRandomDate(start, end) {
  const date = new Date(
    start.getTime() + Math.random() * (end.getTime() - start.getTime())
  );
  return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
}

function getRandomNumber(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function getRandomRevenue() {
  return `$${getRandomNumber(5000, 50000).toLocaleString()}`;
}

function getRandomPublisher() {
  const publishers = [
    "AdTech Media",
    "Digital Publishers Inc",
    "Content Network Pro",
    "Media Solutions LLC",
    "Premium Publishers",
    "Global Ad Network",
    "Elite Media Group",
    "Strategic Publishers",
  ];
  return publishers[Math.floor(Math.random() * publishers.length)];
}

function getRandomCompetitor() {
  const competitors = [
    "Google AdSense",
    "Amazon Publisher Services",
    "Mediavine",
    "AdThrive",
    "Ezoic",
    "MonetizeMore",
    "Sovrn",
    "Media.net",
  ];
  return competitors[Math.floor(Math.random() * competitors.length)];
}

function getRandomStage() {
  const stages = ["Testing", "Live", "Paused", "Optimization", "Analysis"];
  return stages[Math.floor(Math.random() * stages.length)];
}

function getRandomTestStatus() {
  const statuses = [
    "A/B test running successfully with 50/50 traffic split",
    "Test paused for optimization",
    "Collecting baseline data",
    "Ready to launch test next week",
    "Test completed, analyzing results",
  ];
  return statuses[Math.floor(Math.random() * statuses.length)];
}

function getRandomStakeholders() {
  const stakeholders = [
    "John Smith (PM), Sarah Johnson (Dev), Mike Chen (Analytics)",
    "Lisa Wang (Lead), Tom Brown (QA), Alex Rivera (Design)",
    "Emma Davis (Manager), Chris Lee (Engineer), Jordan Taylor (Data)",
  ];
  return stakeholders[Math.floor(Math.random() * stakeholders.length)];
}

function getRandomObservations() {
  const observations = [
    "Revenue trending upward with strong mobile performance",
    "Desktop traffic showing consistent growth",
    "Mobile optimization needed for better conversion",
    "Strong performance across all demographics",
    "Seasonal trends affecting overall metrics",
  ];
  return observations[Math.floor(Math.random() * observations.length)];
}

function getRandomIssues() {
  const issues = [
    "Working on mobile optimization to improve user experience and ad viewability",
    "Investigating discrepancy in reporting between platforms",
    "Implementing new ad formats to increase revenue potential",
    "Addressing page speed concerns affecting ad performance",
  ];
  return issues[Math.floor(Math.random() * issues.length)];
}

function getRandomNextSteps() {
  const steps = [
    "Schedule weekly check-in meetings and implement new tracking system",
    "Launch A/B test for new ad placements next week",
    "Complete technical integration and begin testing phase",
    "Finalize contract terms and prepare for go-live",
  ];
  return steps[Math.floor(Math.random() * steps.length)];
}

function getRandomCSMName() {
  const csmNames = [
    "Kristen",
    "Sarah Johnson",
    "Mike Chen",
    "Lisa Wang",
    "Tom Brown",
    "Alex Rivera",
    "Emma Davis",
    "Chris Lee",
    "Jordan Taylor",
    "Ashley Miller",
    "David Wilson",
    "Rachel Green",
    "Kevin Martinez",
    "Nicole Thompson",
    "Ryan Anderson",
  ];
  return csmNames[Math.floor(Math.random() * csmNames.length)];
}

// Fill all fields with test data
async function fillTestData() {
  const confirmed = await showCustomConfirm(
    "This will fill all fields with test data. Continue?",
    "Fill Test Data"
  );
  if (!confirmed) {
    return;
  }

  // Fill all contenteditable elements
  const editableElements = document.querySelectorAll(
    '[contenteditable="true"]'
  );
  editableElements.forEach((element) => {
    if (
      element.classList.contains("metric-label") ||
      element.classList.contains("ab-label") ||
      element.classList.contains("publisher-label") ||
      element.classList.contains("progress-label") ||
      element.classList.contains("nps-label") ||
      element.classList.contains("updates-label")
    ) {
      // Skip labels - they should not be filled
      return;
    }

    // Skip main section headers (YIELD, ONBOARDING, CUSTOMER SUCCESS)
    if (element.tagName === "H2" && element.closest(".section-header")) {
      return;
    }

    // Skip subsection headers (WEEK-OVER-WEEK PERFORMANCE TRENDS, A/B TEST UPDATES, etc.)
    if (element.tagName === "H3" && element.closest(".subsection-header")) {
      return;
    }

    // Skip main page title
    if (element.tagName === "H1") {
      return;
    }

    if (element.classList.contains("observations-cell")) {
      element.textContent = getRandomObservations();
    } else if (element.classList.contains("test-status-cell")) {
      element.textContent = getRandomTestStatus();
    } else if (element.classList.contains("stakeholders-cell")) {
      element.textContent = getRandomStakeholders();
    } else if (element.classList.contains("updates-issues")) {
      element.textContent = getRandomIssues();
    } else if (element.classList.contains("updates-steps")) {
      element.textContent = getRandomNextSteps();
    } else if (element.classList.contains("nps-score")) {
      element.textContent = getRandomNumber(6, 10);
    } else if (
      element.classList.contains("churned-publisher") ||
      element.classList.contains("notice-publisher")
    ) {
      element.textContent = getRandomPublisher();
    } else if (
      element.classList.contains("churned-revenue") ||
      element.classList.contains("notice-revenue")
    ) {
      element.textContent = getRandomRevenue();
    } else if (
      element.classList.contains("churned-reason") ||
      element.classList.contains("notice-reason")
    ) {
      element.textContent = "Contract ended";
    } else if (
      element.textContent.includes("REVENUE") ||
      element.textContent.includes("RPM")
    ) {
      element.textContent = getRandomRevenue();
    } else if (element.textContent.includes("IMPRESSIONS")) {
      element.textContent = `${getRandomNumber(100, 999)}K`;
    } else {
      // Default content for other fields
      const cellText =
        element.closest("td")?.previousElementSibling?.textContent || "";
      if (cellText.includes("PUBLISHER")) {
        element.textContent = getRandomPublisher();
      } else if (cellText.includes("COMPETITOR")) {
        element.textContent = getRandomCompetitor();
      } else if (cellText.includes("STAGE")) {
        element.textContent = getRandomStage();
      } else if (cellText.includes("CSM")) {
        element.textContent = getRandomCSMName();
      } else {
        element.textContent = getRandomRevenue();
      }
    }
  });

  // Fill date fields
  const dateWrappers = document.querySelectorAll(".date-field-wrapper");
  dateWrappers.forEach((wrapper) => {
    const hiddenInput = wrapper.querySelector(".date-input-hidden");
    const dateText = wrapper.querySelector(".date-text");
    if (hiddenInput && dateText) {
      const randomDate = getRandomDate(
        new Date(2024, 0, 1),
        new Date(2024, 11, 31)
      );
      const inputDate = formatDateForInput(randomDate);
      hiddenInput.value = inputDate;
      hiddenInput.setAttribute("data-display-value", randomDate);
      dateText.textContent = randomDate;
    }
  });

  // Fill week pickers
  const weekPickers = document.querySelectorAll(".week-picker");
  weekPickers.forEach((picker) => {
    const randomDate = new Date(2024, 4, 15 + Math.floor(Math.random() * 14)); // May 2024
    picker.value = randomDate.toISOString().split("T")[0];
    handleWeekChange(picker);
  });

  // Special handling for publisher updates table
  fillPublisherUpdatesData();

  await showCustomAlert(
    "Test data has been filled in all fields!",
    "Success",
    "success"
  );
}

// Special function to fill publisher updates table data
function fillPublisherUpdatesData() {
  // Find the publisher updates table specifically
  const publisherUpdatesSection = Array.from(
    document.querySelectorAll("h3")
  ).find((h3) => h3.textContent.includes("PUBLISHER ISSUES/UPDATES"));

  if (!publisherUpdatesSection) return;

  let table = publisherUpdatesSection
    .closest(".table-container")
    ?.querySelector("table");

  // Try alternative ways to find the table
  if (!table) {
    const subsection = publisherUpdatesSection.closest(".subsection");
    const altTable = subsection?.querySelector("table");
    if (altTable) {
      table = altTable;
    } else {
      return;
    }
  }

  // Get all rows in the table
  const rows = table.querySelectorAll("tbody tr");

  rows.forEach((row, rowIndex) => {
    const cells = row.querySelectorAll("td");

    if (rowIndex === 0) {
      // Row 0: PUBLISHER | value | TIER | value | AVG. MONTHLY REV. | value
      if (cells[1]) cells[1].textContent = getRandomPublisher();
      if (cells[3]) cells[3].textContent = getRandomNumber(1, 3);
      if (cells[5]) cells[5].textContent = getRandomRevenue();
    } else if (rowIndex === 1) {
      // Row 1: STAGE | value (colspan=3) | CSM | value
      if (cells[1]) cells[1].textContent = getRandomStage();
      if (cells[3]) cells[3].textContent = getRandomCSMName();
    } else if (rowIndex === 2) {
      // Row 2: STAKEHOLDERS | value (colspan=3) | LAST COMM. DATE | value
      if (cells[1]) cells[1].textContent = getRandomStakeholders();
      // Note: LAST COMM. DATE is handled by date picker, skip it
    } else if (rowIndex === 3) {
      // Row 3: ISSUE(S)/UPDATE(S) | value (colspan=5)
      if (cells[1]) cells[1].textContent = getRandomIssues();
    } else if (rowIndex === 4) {
      // Row 4: NEXT STEPS | value (colspan=5)
      if (cells[1]) cells[1].textContent = getRandomNextSteps();
    }
  });
}

// Clear all data
async function clearAllData() {
  const confirmed = await showCustomConfirm(
    "This will clear ALL data from the page. This cannot be undone. Continue?",
    "Clear All Data"
  );
  if (!confirmed) {
    return;
  }

  // Clear all contenteditable elements (except labels and main headers)
  const editableElements = document.querySelectorAll(
    '[contenteditable="true"]'
  );
  editableElements.forEach((element) => {
    if (
      !element.classList.contains("metric-label") &&
      !element.classList.contains("ab-label") &&
      !element.classList.contains("publisher-label") &&
      !element.classList.contains("progress-label") &&
      !element.classList.contains("nps-label") &&
      !element.classList.contains("updates-label") &&
      // Don't clear main section headers
      !(element.tagName === "H2" && element.closest(".section-header")) &&
      // Don't clear subsection headers
      !(element.tagName === "H3" && element.closest(".subsection-header")) &&
      // Don't clear main page title
      !(element.tagName === "H1")
    ) {
      element.textContent = "";
    }
  });

  // Clear date fields
  const dateWrappers = document.querySelectorAll(".date-field-wrapper");
  dateWrappers.forEach((wrapper) => {
    const hiddenInput = wrapper.querySelector(".date-input-hidden");
    const dateText = wrapper.querySelector(".date-text");
    if (hiddenInput && dateText) {
      hiddenInput.value = "";
      hiddenInput.removeAttribute("data-display-value");
      dateText.textContent = "mm/dd/yyyy";
    }
  });

  // Clear week pickers and reset headers
  const weekPickers = document.querySelectorAll(".week-picker");
  weekPickers.forEach((picker) => {
    picker.value = "";
    picker.removeAttribute("data-display-value");

    // Reset header text to default
    const headerCell = picker.closest("th");
    if (headerCell) {
      const currentText = headerCell.textContent;
      if (currentText.includes("Prior Wednesday")) {
        headerCell.innerHTML =
          "5/15 - 5/21 (Prior Wednesday and Prior 7 Days)<br>" +
          createWeekPicker("5/15 - 5/21", "header-week-picker");
      } else if (currentText.includes("Most Recent")) {
        headerCell.innerHTML =
          "5/22 - 5/28 (Most Recent Wednesday and Prior 7 Days)<br>" +
          createWeekPicker("5/22 - 5/28", "header-week-picker");
      }
    }
  });

  // Clear localStorage
  localStorage.removeItem("executiveSummaryData");

  await showCustomAlert("All data has been cleared!", "Success", "success");
}

// Flag to prevent multiple initializations
let isInitialized = false;

// Initialize the application
document.addEventListener("DOMContentLoaded", function () {
  // Prevent multiple initializations
  if (isInitialized) {
    console.log("Application already initialized, skipping...");
    return;
  }
  isInitialized = true;

  console.log("Initializing application...");

  // Clean up any corrupted localStorage data first
  cleanupSavedReports();

  // For now, let's completely clear localStorage to start fresh (commented out for testing)
  // console.log("Clearing all localStorage data to start fresh...");
  // localStorage.removeItem("executiveSummaryData");
  // localStorage.removeItem("savedReports");

  // Add event listeners to buttons
  document
    .getElementById("save-btn")
    .addEventListener("click", () => saveData());
  document
    .getElementById("test-btn")
    .addEventListener("click", () => fillTestData());
  document
    .getElementById("clear-btn")
    .addEventListener("click", () => clearAllData());

  // Clear any existing tables first to prevent duplicates
  clearAllTablesOnly();
  console.log("Cleared existing tables");

  // Initialize tab navigation - set Reader as default
  console.log(
    "Initializing Executive Summary Tool with Reader section as default..."
  );

  // Hide header actions since Reader is the default tab
  const headerActions = document.querySelector(".header-actions");
  if (headerActions) {
    headerActions.style.display = "none";
  }

  // Load saved reports on startup and open most recent report
  if (document.getElementById("view-only-content")) {
    console.log("Reader section found, loading saved reports...");
    loadSavedReports();

    // Automatically open the most recently saved report
    setTimeout(() => {
      openMostRecentReport();
    }, 100);
  }

  // Always ensure tables are available - first try to load saved data, then ensure defaults exist
  setTimeout(() => {
    const savedData = localStorage.getItem("executiveSummaryData");

    if (savedData) {
      console.log("Found saved data, loading...");
      try {
        const data = JSON.parse(savedData);
        console.log("Loading saved data:", data);

        let totalTablesRestored = 0;

        // Restore data to each section (skip view-only sections)
        Object.keys(data).forEach((sectionId) => {
          // Skip view-only sections
          if (sectionId.startsWith("view-")) {
            console.log(
              `Skipping view-only section during restore: ${sectionId}`
            );
            return;
          }

          const section = data[sectionId];
          if (section.subsections) {
            Object.keys(section.subsections).forEach((subsectionId) => {
              // Skip view-only subsections
              if (subsectionId.startsWith("view-")) {
                console.log(
                  `Skipping view-only subsection during restore: ${subsectionId}`
                );
                return;
              }

              const subsection = section.subsections[subsectionId];
              if (subsection.tables && subsection.tables.length > 0) {
                console.log(
                  `Restoring ${subsection.tables.length} table(s) for ${subsectionId}`
                );
                // Restore tables for this subsection
                restoreTablesForSubsection(subsectionId, subsection.tables);
                totalTablesRestored += subsection.tables.length;
              }
            });
          }
        });

        console.log(
          `Data loaded successfully - restored ${totalTablesRestored} tables total`
        );

        // Always ensure we have default tables if none were restored
        if (totalTablesRestored === 0) {
          console.log(
            "No tables were restored from saved data, adding default tables..."
          );
          addDefaultTables();
        }
      } catch (error) {
        console.error("Error loading saved data:", error);
        console.log("Loading failed, adding default tables...");
        addDefaultTables();
      }
    } else {
      console.log("No saved data found, adding default tables...");
      addDefaultTables();
    }

    // Additional safety check - ensure all sections have at least one table
    setTimeout(() => {
      ensureAllSectionsHaveTables();
    }, 200);
  }, 100);

  // Auto-save when page becomes hidden (user switches tabs or minimizes)
  document.addEventListener("visibilitychange", function () {
    if (document.hidden) {
      // Force immediate save when page becomes hidden
      if (autoSaveTimeout) {
        clearTimeout(autoSaveTimeout);
      }
      saveData(false);
      lastSaveTime = Date.now();
      showAutoSaveStatus("saved", "Saved on tab switch");
      console.log(
        "Auto-saved on page visibility change at:",
        new Date().toLocaleTimeString()
      );
    }
  });

  // Auto-save before page unload (user closes tab/window or navigates away)
  window.addEventListener("beforeunload", function (e) {
    // Force immediate save before unload
    if (autoSaveTimeout) {
      clearTimeout(autoSaveTimeout);
    }
    saveData(false);
    console.log(
      "Auto-saved before page unload at:",
      new Date().toLocaleTimeString()
    );
  });

  // Periodic auto-save every 30 seconds as a backup
  setInterval(() => {
    const now = Date.now();
    if (now - lastSaveTime >= 30000) {
      // 30 seconds
      saveData(false);
      lastSaveTime = now;
      showAutoSaveStatus("saved", "Periodic save");
      console.log("Periodic auto-save at:", new Date().toLocaleTimeString());
    }
  }, 30000);

  // Initialize rich text editor and functionality after everything is loaded
  setTimeout(() => {
    initializeRichTextEditor();
    addRichTextFunctionality();

    // Test function to verify rich text functionality (remove in production)
    // testRichTextFunctionality();
  }, 500);
});

// Function to generate mobile-friendly card layout for tables
function generateMobileTableCards(table) {
  if (!table.headers || !table.rows || table.rows.length === 0) {
    return "";
  }

  let mobileHTML = "";

  // Handle different table layouts based on structure
  if (isPerformanceTrendsTable(table)) {
    // Special handling for WEEK-OVER-WEEK PERFORMANCE TRENDS
    mobileHTML = generatePerformanceTrendsCard(table);
  } else if (hasComplexLayout(table)) {
    // For complex layouts (A/B tests, publisher updates, etc.), create smart cards
    mobileHTML = generateComplexLayoutCards(table);
  } else if (table.headers.length > 2) {
    // For tables with multiple columns, create cards for each row
    table.rows.forEach((row) => {
      mobileHTML += generateRowCard(table, row);
    });
  } else {
    // For simple two-column tables, create a single card
    mobileHTML += generateSimpleCard(table);
  }

  return mobileHTML;
}

// Helper function to detect performance trends table
function isPerformanceTrendsTable(table) {
  // Check if this is a performance trends table by looking for date headers
  return (
    table.headers.length === 3 &&
    table.headers[0] === "" &&
    table.headers[1] &&
    table.headers[1].includes("Prior Wednesday") &&
    table.headers[2] &&
    table.headers[2].includes("Most Recent Wednesday")
  );
}

// Generate special card for performance trends table
function generatePerformanceTrendsCard(table) {
  let mobileHTML = `<div class="mobile-table-card">`;

  // Add date headers as a special header section
  mobileHTML += `<div class="mobile-table-header">Performance Trends</div>`;
  mobileHTML += `<div class="mobile-date-headers">`;
  mobileHTML += `<div class="mobile-date-header prior">${table.headers[1]}</div>`;
  mobileHTML += `<div class="mobile-date-header recent">${table.headers[2]}</div>`;
  mobileHTML += `</div>`;

  mobileHTML += `<div class="mobile-table-content">`;

  // Process each metric row
  table.rows.forEach((row) => {
    if (row[0] && row[0].trim() !== "") {
      const metric = row[0];
      const priorValue = row[1] || "";
      const recentValue = row[2] || "";

      // Special handling for observations row
      if (metric.toUpperCase().includes("OBSERVATION")) {
        mobileHTML += `
          <div class="mobile-row observations-row">
            <div class="mobile-row-label">${metric}</div>
            <div class="mobile-row-data observations">${
              priorValue || recentValue
            }</div>
          </div>
        `;
      } else {
        // Regular metric with two values
        mobileHTML += `
          <div class="mobile-metric-row">
            <div class="mobile-row-label">${metric}</div>
            <div class="mobile-metric-values">
              <div class="mobile-metric-value">
                <span class="metric-period">Prior</span>
                <span class="metric-data">${priorValue}</span>
              </div>
              <div class="mobile-metric-value">
                <span class="metric-period">Recent</span>
                <span class="metric-data">${recentValue}</span>
              </div>
            </div>
          </div>
        `;
      }
    }
  });

  mobileHTML += `</div></div>`;
  return mobileHTML;
}

// Helper function to detect complex layouts
function hasComplexLayout(table) {
  // Check if table has mostly empty headers (indicates complex layout)
  const emptyHeaders = table.headers.filter(
    (h) => !h || h.trim() === ""
  ).length;
  const totalHeaders = table.headers.length;

  // If more than half the headers are empty, it's likely a complex layout
  return emptyHeaders > totalHeaders / 2;
}

// Generate cards for complex layouts (A/B tests, publisher updates, etc.)
function generateComplexLayoutCards(table) {
  console.log("generateComplexLayoutCards called with table:", table);
  let mobileHTML = "";

  // Check for specific table types that need special handling
  if (isABTestTable(table)) {
    console.log("Detected A/B test table");
    return generateABTestCard(table);
  } else if (isNewlyOnboardedTable(table)) {
    console.log("Detected newly onboarded table");
    return generateNewlyOnboardedCard(table);
  } else if (isInProgressTable(table)) {
    console.log("Detected in-progress table");
    return generateInProgressCard(table);
  } else if (isPublisherUpdatesTable(table)) {
    console.log("Detected publisher updates table");
    return generatePublisherUpdatesCard(table);
  }

  console.log("No special table type detected, using default generation");

  table.rows.forEach((row, rowIndex) => {
    if (!row || row.every((cell) => !cell || cell.trim() === "")) {
      return; // Skip empty rows
    }

    mobileHTML += `<div class="mobile-table-card">`;

    // Try to find a meaningful header from the first non-empty cell
    const headerCell = row.find((cell) => cell && cell.trim() !== "");
    const isLabelRow = headerCell && headerCell.toString().trim() !== "";

    if (isLabelRow && headerCell) {
      mobileHTML += `<div class="mobile-table-header">${headerCell}</div>`;
    }

    mobileHTML += `<div class="mobile-table-content">`;

    // Process cells in pairs or meaningful groups
    for (let i = 0; i < row.length; i++) {
      const cell = row[i];
      if (!cell || cell.trim() === "") continue;

      // Skip if this cell was used as header
      if (isLabelRow && cell === headerCell) continue;

      // Look for label-value pairs
      if (i < row.length - 1) {
        const nextCell = row[i + 1];
        const isCurrentCellLabel = isLikelyLabel(cell);

        if (isCurrentCellLabel && nextCell && nextCell.trim() !== "") {
          // Found a label-value pair
          const dataType = getDataType(nextCell, cell);
          mobileHTML += `
            <div class="mobile-row">
              <div class="mobile-row-label">${cell}</div>
              <div class="mobile-row-data ${dataType}">${nextCell}</div>
            </div>
          `;
          i++; // Skip the next cell since we used it as value
          continue;
        }
      }

      // If not part of a pair, show as standalone data
      if (!isLikelyLabel(cell)) {
        const dataType = getDataType(cell, "");
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-data ${dataType}">${cell}</div>
          </div>
        `;
      }
    }

    mobileHTML += `</div></div>`;
  });

  return mobileHTML;
}

// Helper function to detect A/B test tables
function isABTestTable(table) {
  // Check if this looks like an A/B test table structure
  return (
    table.headers.length === 6 &&
    table.headers.every((h) => h === "") &&
    table.rows.some((row) =>
      row.some(
        (cell) =>
          cell &&
          (cell.includes("COMPETITOR") ||
            cell.includes("TEST STATUS") ||
            cell.includes("% SPLIT"))
      )
    ) &&
    // Make sure it's NOT a publisher updates table by checking it doesn't have ISSUE or NEXT STEPS
    !table.rows.some((row) =>
      row.some(
        (cell) =>
          cell && (cell.includes("ISSUE") || cell.includes("NEXT STEPS"))
      )
    )
  );
}

// Helper function to detect newly onboarded publishers tables
function isNewlyOnboardedTable(table) {
  return (
    table.headers.length === 4 &&
    table.headers.every((h) => h === "") &&
    table.rows.some((row) =>
      row.some(
        (cell) =>
          cell && (cell.includes("LIVE DATE") || cell.includes("DEMAND STATUS"))
      )
    ) &&
    // Make sure it's NOT an in-progress table by checking it doesn't have EXP. LIVE DATE
    !table.rows.some((row) =>
      row.some((cell) => cell && cell.includes("EXP. LIVE DATE"))
    )
  );
}

// Helper function to detect in-progress publishers tables
function isInProgressTable(table) {
  console.log("Checking if in-progress table:", table);
  const result =
    table.headers.length === 4 &&
    table.headers.every((h) => h === "") &&
    table.rows.some((row) =>
      row.some(
        (cell) =>
          cell && (cell.includes("EXP. LIVE DATE") || cell.includes("STAGE"))
      )
    );
  console.log("isInProgressTable result:", result);
  return result;
}

// Helper function to detect publisher updates tables
function isPublisherUpdatesTable(table) {
  return (
    table.headers.length === 6 &&
    table.headers.every((h) => h === "") &&
    table.rows.some((row) =>
      row.some(
        (cell) =>
          cell &&
          (cell.includes("STAKEHOLDERS") ||
            cell.includes("ISSUE") ||
            cell.includes("NEXT STEPS"))
      )
    )
  );
}

// Generate special card for A/B test tables
function generateABTestCard(table) {
  let mobileHTML = `<div class="mobile-table-card">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Process the A/B test table structure
  // Expected structure: PUBLISHER | value | TIER | value | empty | empty
  //                    COMPETITOR | value | % SPLIT | value | empty | empty
  //                    STAGE | value | CSM/OBS | value | empty | empty
  //                    STAKEHOLDERS | value | value | value | empty | empty
  //                    START DATE | value | END DATE | value | LAST COMM. DATE | value
  //                    TEST STATUS | value spanning multiple columns

  table.rows.forEach((row) => {
    if (!row || row.every((cell) => !cell || cell.trim() === "")) {
      return;
    }

    // Handle different row types based on content
    if (row[0] && row[0].includes("PUBLISHER")) {
      // Publisher and Tier row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">PUBLISHER</div>
          <div class="mobile-row-data">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("TIER")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">TIER</div>
            <div class="mobile-row-data">${row[3] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("COMPETITOR")) {
      // Competitor and % Split row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">COMPETITOR</div>
          <div class="mobile-row-data">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("SPLIT")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">% SPLIT</div>
            <div class="mobile-row-data revenue">${row[3] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("STAGE")) {
      // Stage and CSM/OBS row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">STAGE</div>
          <div class="mobile-row-data status">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("CSM")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">CSM/OBS</div>
            <div class="mobile-row-data revenue">${row[3] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("STAKEHOLDERS")) {
      // Stakeholders row - combine all non-empty cells
      const stakeholderData = row
        .slice(1)
        .filter((cell) => cell && cell.trim() !== "")
        .join(", ");
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">STAKEHOLDERS</div>
          <div class="mobile-row-data stakeholders">${stakeholderData}</div>
        </div>
      `;
    } else if (row[0] && row[0].includes("START DATE")) {
      // Date row with three dates
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">START DATE</div>
          <div class="mobile-row-data date">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("END DATE")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">END DATE</div>
            <div class="mobile-row-data date">${row[3] || ""}</div>
          </div>
        `;
      }
      if (row[4] && row[4].includes("LAST COMM")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">LAST COMM. DATE</div>
            <div class="mobile-row-data date">${row[5] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("TEST STATUS")) {
      // Test status row - spans multiple columns
      const statusData = row
        .slice(1)
        .filter((cell) => cell && cell.trim() !== "")
        .join(" ");
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">TEST STATUS</div>
          <div class="mobile-row-data status">${statusData}</div>
        </div>
      `;
    }
  });

  mobileHTML += `</div></div>`;
  return mobileHTML;
}

// Generate special card for newly onboarded publishers
function generateNewlyOnboardedCard(table) {
  let mobileHTML = `<div class="mobile-table-card">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Expected structure: PUBLISHER | value | LIVE DATE | value
  //                    PROJECTED REVENUE | value | OBS | value
  //                    DEMAND STATUS | value spanning multiple columns

  table.rows.forEach((row) => {
    if (!row || row.every((cell) => !cell || cell.trim() === "")) {
      return;
    }

    if (row[0] && row[0].includes("PUBLISHER")) {
      // Publisher and Live Date row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">PUBLISHER</div>
          <div class="mobile-row-data">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("LIVE DATE")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">LIVE DATE</div>
            <div class="mobile-row-data date">${row[3] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("PROJECTED REVENUE")) {
      // Projected Revenue and OBS row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">PROJECTED REVENUE</div>
          <div class="mobile-row-data revenue">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("OBS")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">OBS</div>
            <div class="mobile-row-data">${row[3] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("DEMAND STATUS")) {
      // Demand status row - spans multiple columns
      const statusData = row
        .slice(1)
        .filter((cell) => cell && cell.trim() !== "")
        .join(" ");
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">DEMAND STATUS</div>
          <div class="mobile-row-data status">${statusData}</div>
        </div>
      `;
    }
  });

  mobileHTML += `</div></div>`;
  return mobileHTML;
}

// Generate special card for in-progress publishers - DYNAMIC APPROACH LIKE A/B TEST UPDATES
function generateInProgressCard(table) {
  console.log("generateInProgressCard called with table:", table);

  let mobileHTML = `<div class="mobile-table-card">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Process the in-progress publishers table structure
  // Expected structure: PUBLISHER | value | EXP. LIVE DATE | value
  //                    STAGE | value | OBS | value
  //                    PROJECTED REVENUE | value | LAST COMM. DATE | value
  //                    STATUS | value spanning multiple columns

  table.rows.forEach((row) => {
    if (!row || row.every((cell) => !cell || cell.trim() === "")) {
      return;
    }

    // Handle different row types based on content
    if (row[0] && row[0].includes("PUBLISHER")) {
      // Publisher and Exp. Live Date row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">PUBLISHER</div>
          <div class="mobile-row-data">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("EXP. LIVE DATE")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">EXP. LIVE DATE</div>
            <div class="mobile-row-data date">${row[3] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("STAGE")) {
      // Stage and OBS row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">STAGE</div>
          <div class="mobile-row-data status">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("OBS")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">OBS</div>
            <div class="mobile-row-data">${row[3] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("PROJECTED REVENUE")) {
      // Projected Revenue and Last Comm. Date row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">PROJECTED REVENUE</div>
          <div class="mobile-row-data revenue">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("LAST COMM")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">LAST COMM. DATE</div>
            <div class="mobile-row-data date">${row[3] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("STATUS")) {
      // Status row - spans multiple columns
      const statusData = row
        .slice(1)
        .filter((cell) => cell && cell.trim() !== "")
        .join(" ");
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">STATUS</div>
          <div class="mobile-row-data status">${statusData}</div>
        </div>
      `;
    }
  });

  mobileHTML += `</div></div>`;
  return mobileHTML;
}

// Generate special card for publisher updates/issues - DYNAMIC APPROACH LIKE A/B TEST UPDATES
function generatePublisherUpdatesCard(table) {
  console.log("generatePublisherUpdatesCard called with table:", table);

  let mobileHTML = `<div class="mobile-table-card">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Process the publisher updates table structure
  // Expected structure: PUBLISHER | value | TIER | value | AVG. MONTHLY REV. | value
  //                    STAGE | value (colspan=3) | CSM | value
  //                    STAKEHOLDERS | value (colspan=3) | LAST COMM. DATE | value
  //                    ISSUE(S)/ UPDATE(S) | value spanning multiple columns
  //                    NEXT STEPS | value spanning multiple columns

  table.rows.forEach((row) => {
    if (!row || row.every((cell) => !cell || cell.trim() === "")) {
      return;
    }

    // Handle different row types based on content
    if (row[0] && row[0].includes("PUBLISHER")) {
      // Publisher, Tier, and Avg. Monthly Rev. row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">PUBLISHER</div>
          <div class="mobile-row-data">${row[1] || ""}</div>
        </div>
      `;
      if (row[2] && row[2].includes("TIER")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">TIER</div>
            <div class="mobile-row-data">${row[3] || ""}</div>
          </div>
        `;
      }
      if (row[4] && row[4].includes("AVG. MONTHLY REV")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">AVG. MONTHLY REV.</div>
            <div class="mobile-row-data revenue">${row[5] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("STAGE")) {
      // Stage and CSM row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">STAGE</div>
          <div class="mobile-row-data status">${row[1] || ""}</div>
        </div>
      `;
      if (row[3] && row[3].includes("CSM")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">CSM</div>
            <div class="mobile-row-data stakeholders">${row[4] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("STAKEHOLDERS")) {
      // Stakeholders and Last Comm. Date row
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">STAKEHOLDERS</div>
          <div class="mobile-row-data stakeholders">${row[1] || ""}</div>
        </div>
      `;
      if (row[3] && row[3].includes("LAST COMM")) {
        mobileHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">LAST COMM. DATE</div>
            <div class="mobile-row-data date">${row[4] || ""}</div>
          </div>
        `;
      }
    } else if (row[0] && row[0].includes("ISSUE")) {
      // Issues/Updates row - spans multiple columns
      const issuesData = row
        .slice(1)
        .filter((cell) => cell && cell.trim() !== "")
        .join(" ");
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">ISSUE(S)/UPDATE(S)</div>
          <div class="mobile-row-data observations">${issuesData}</div>
        </div>
      `;
    } else if (row[0] && row[0].includes("NEXT STEPS")) {
      // Next Steps row - spans multiple columns
      const stepsData = row
        .slice(1)
        .filter((cell) => cell && cell.trim() !== "")
        .join(" ");
      mobileHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">NEXT STEPS</div>
          <div class="mobile-row-data observations">${stepsData}</div>
        </div>
      `;
    }
  });

  mobileHTML += `</div></div>`;
  return mobileHTML;
}

// Generate card for a single row in multi-column tables
function generateRowCard(table, row) {
  let cardHTML = `<div class="mobile-table-card">`;

  // Use first cell as card header if it's a label
  const isFirstCellLabel = row[0] && row[0].toString().trim() !== "";
  if (isFirstCellLabel && row[0]) {
    cardHTML += `<div class="mobile-table-header">${row[0]}</div>`;
    cardHTML += `<div class="mobile-table-content">`;

    // Add remaining cells as data rows
    for (let i = 1; i < row.length; i++) {
      if (row[i] && row[i].trim() !== "") {
        const headerLabel = table.headers[i] || `Column ${i + 1}`;
        const dataType = getDataType(row[i], headerLabel);
        cardHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">${headerLabel}</div>
            <div class="mobile-row-data ${dataType}">${row[i]}</div>
          </div>
        `;
      }
    }
  } else {
    // If first cell is not a label, treat all cells as data
    cardHTML += `<div class="mobile-table-content">`;
    for (let i = 0; i < row.length; i++) {
      if (row[i] && row[i].trim() !== "") {
        const headerLabel = table.headers[i] || `Column ${i + 1}`;
        const dataType = getDataType(row[i], headerLabel);
        cardHTML += `
          <div class="mobile-row">
            <div class="mobile-row-label">${headerLabel}</div>
            <div class="mobile-row-data ${dataType}">${row[i]}</div>
          </div>
        `;
      }
    }
  }

  cardHTML += `</div></div>`;
  return cardHTML;
}

// Generate simple card for two-column tables
function generateSimpleCard(table) {
  let cardHTML = `<div class="mobile-table-card">`;
  cardHTML += `<div class="mobile-table-content">`;

  table.rows.forEach((row) => {
    if (row.length >= 2 && row[0] && row[1]) {
      const dataType = getDataType(row[1], row[0]);
      cardHTML += `
        <div class="mobile-row">
          <div class="mobile-row-label">${row[0]}</div>
          <div class="mobile-row-data ${dataType}">${row[1]}</div>
        </div>
      `;
    }
  });

  cardHTML += `</div></div>`;
  return cardHTML;
}

// Helper function to detect if a cell is likely a label
function isLikelyLabel(cell) {
  if (!cell || typeof cell !== "string") return false;

  const cellUpper = cell.toUpperCase();
  const labelKeywords = [
    "PUBLISHER",
    "TIER",
    "STAGE",
    "STATUS",
    "DATE",
    "REVENUE",
    "CSM",
    "OBS",
    "STAKEHOLDERS",
    "COMPETITOR",
    "SPLIT",
    "START",
    "END",
    "LAST",
    "COMM",
    "TEST",
    "PROJECTED",
    "LIVE",
    "DEMAND",
    "EXP",
    "ISSUE",
    "UPDATE",
    "NEXT",
    "STEPS",
    "NETWORK",
    "MTD",
    "QTD",
    "YTD",
    "SITE",
    "AVG",
    "MONTHLY",
    "DARK",
    "REASON",
    "NOTICE",
  ];

  return (
    labelKeywords.some((keyword) => cellUpper.includes(keyword)) ||
    cell === cell.toUpperCase() || // All caps suggests label
    cell.endsWith(":")
  ); // Ends with colon suggests label
}

// Helper function to determine data type for styling
function getDataType(cellContent, headerOrLabel) {
  const content = cellContent.toString().toLowerCase();
  const label = headerOrLabel.toString().toLowerCase();

  if (label.includes("observation") || content.includes("observation")) {
    return "observations";
  }
  if (
    label.includes("revenue") ||
    label.includes("cpm") ||
    content.includes("$")
  ) {
    return "revenue";
  }
  if (
    label.includes("date") ||
    label.includes("week") ||
    content.match(/\d{1,2}\/\d{1,2}/)
  ) {
    return "date";
  }
  if (
    label.includes("status") ||
    label.includes("stage") ||
    content.includes("onboarding") ||
    content.includes("ongoing") ||
    content.includes("ended")
  ) {
    return "status";
  }
  if (
    label.includes("stakeholder") ||
    label.includes("csm") ||
    content.includes("kristen") ||
    content.includes("team")
  ) {
    return "stakeholders";
  }

  return "";
}

// Auto-save every 30 seconds (silent)
setInterval(() => saveData(false), 30000);

// Tab Navigation Functions
function switchTab(tabName) {
  // Remove active class from all tabs and content
  document
    .querySelectorAll(".tab-btn")
    .forEach((btn) => btn.classList.remove("active"));
  document
    .querySelectorAll(".tab-content")
    .forEach((content) => content.classList.remove("active"));

  // Add active class to selected tab and content
  document.getElementById(`${tabName}-tab`).classList.add("active");
  document.getElementById(`${tabName}-content`).classList.add("active");

  // Hide/show header actions based on tab
  const headerActions = document.querySelector(".header-actions");
  if (tabName === "view-only") {
    headerActions.style.display = "none";
    loadSavedReports();
  } else {
    headerActions.style.display = "flex";
  }
}

// Saved Reports Management Functions
function saveReportWithAutoTitle() {
  return new Promise((resolve) => {
    // Auto-generate title based on current date and time
    const now = new Date();
    const dateStr = now.toLocaleDateString("en-US", {
      month: "2-digit",
      day: "2-digit",
      year: "numeric",
    });
    const timeStr = now.toLocaleTimeString("en-US", {
      hour: "2-digit",
      minute: "2-digit",
      hour12: true,
    });
    const title = `Executive Summary - ${dateStr} ${timeStr}`;

    const data = extractAllData();

    // Clean up any view-only sections that might have been saved previously
    const cleanedData = {};
    Object.keys(data).forEach((sectionId) => {
      if (!sectionId.startsWith("view-")) {
        cleanedData[sectionId] = {
          ...data[sectionId],
          subsections: {},
        };

        // Also clean subsections
        if (data[sectionId].subsections) {
          Object.keys(data[sectionId].subsections).forEach((subsectionId) => {
            if (!subsectionId.startsWith("view-")) {
              cleanedData[sectionId].subsections[subsectionId] =
                data[sectionId].subsections[subsectionId];
            }
          });
        }
      }
    });

    const savedReports = JSON.parse(
      localStorage.getItem("savedReports") || "[]"
    );

    const newReport = {
      id: Date.now().toString(),
      title: title,
      data: cleanedData,
      savedDate: new Date().toISOString(),
      formattedDate:
        new Date().toLocaleDateString() + " " + new Date().toLocaleTimeString(),
    };

    // Add to beginning of array (most recent first)
    savedReports.unshift(newReport);

    // Keep only last 50 reports to prevent localStorage bloat
    if (savedReports.length > 50) {
      savedReports.splice(50);
    }

    localStorage.setItem("savedReports", JSON.stringify(savedReports));

    console.log("Report saved with auto-generated title:", title);
    resolve(true);
  });
}

// Function to clean up corrupted saved reports
function cleanupSavedReports() {
  const savedReports = JSON.parse(localStorage.getItem("savedReports") || "[]");
  let cleaned = false;

  const cleanedReports = savedReports.map((report) => {
    const cleanedData = {};
    let reportCleaned = false;

    Object.keys(report.data).forEach((sectionId) => {
      if (!sectionId.startsWith("view-")) {
        cleanedData[sectionId] = {
          ...report.data[sectionId],
          subsections: {},
        };

        // Also clean subsections
        if (report.data[sectionId].subsections) {
          Object.keys(report.data[sectionId].subsections).forEach(
            (subsectionId) => {
              if (!subsectionId.startsWith("view-")) {
                cleanedData[sectionId].subsections[subsectionId] =
                  report.data[sectionId].subsections[subsectionId];
              } else {
                reportCleaned = true;
              }
            }
          );
        }
      } else {
        reportCleaned = true;
      }
    });

    if (reportCleaned) {
      cleaned = true;
      console.log(`Cleaned report: ${report.title}`);
    }

    return {
      ...report,
      data: cleanedData,
    };
  });

  if (cleaned) {
    localStorage.setItem("savedReports", JSON.stringify(cleanedReports));
    console.log("Cleaned up corrupted saved reports");
  }
}

function loadSavedReports() {
  // Clean up any corrupted reports first
  cleanupSavedReports();

  const savedReports = JSON.parse(localStorage.getItem("savedReports") || "[]");
  const listContainer = document.getElementById("saved-reports-list");

  if (savedReports.length === 0) {
    listContainer.innerHTML = `
      <div class="no-reports-message">
        <div class="no-reports-icon">
          <span class="material-icons">description</span>
        </div>
        <h3>No Saved Reports</h3>
        <p>Save reports from the ADD NEW REPORT section to view them here.</p>
      </div>
    `;
    return;
  }

  listContainer.innerHTML = savedReports
    .map(
      (report) => `
    <div class="saved-report-item" onclick="viewSavedReport('${report.id}')">
      <div class="saved-report-info">
        <div class="saved-report-title">${report.title}</div>
        <div class="saved-report-date">Saved: ${report.formattedDate}</div>
      </div>
      <div class="saved-report-actions" onclick="event.stopPropagation()">
        <button class="btn btn-danger" onclick="deleteSavedReport('${report.id}')" title="Delete Report">
          <span class="material-icons">delete</span>
        </button>
      </div>
    </div>
  `
    )
    .join("");
}

function refreshSavedReports() {
  loadSavedReports();
  showCustomAlert("Reports list refreshed!", "Success", "success");
}

function deleteSavedReport(reportId) {
  showCustomConfirm(
    "Are you sure you want to delete this report? This cannot be undone.",
    "Delete Report"
  ).then((confirmed) => {
    if (confirmed) {
      const savedReports = JSON.parse(
        localStorage.getItem("savedReports") || "[]"
      );
      const updatedReports = savedReports.filter(
        (report) => report.id !== reportId
      );
      localStorage.setItem("savedReports", JSON.stringify(updatedReports));
      loadSavedReports();
      showCustomAlert("Report deleted successfully!", "Success", "success");
    }
  });
}

function viewSavedReport(reportId) {
  const savedReports = JSON.parse(localStorage.getItem("savedReports") || "[]");
  const report = savedReports.find((r) => r.id === reportId);

  if (!report) {
    showCustomAlert("Report not found!", "Error", "error");
    return;
  }

  // Hide saved reports list and show report viewer
  document.querySelector(".saved-reports-container").style.display = "none";
  document.getElementById("report-viewer").style.display = "block";

  // Update the report viewer header with the report title
  const reportViewerHeader = document.querySelector(".report-viewer-header");
  reportViewerHeader.innerHTML = `
    <button class="btn btn-secondary" onclick="backToReportsList()">← Back to Reports</button>
    <div class="report-title-header">
      <h2>${report.title}</h2>
      <div class="report-date">Saved: ${report.formattedDate}</div>
    </div>
  `;

  // Render the report content
  renderViewOnlyReport(report.data);
}

function showSavedReportsList() {
  document.querySelector(".saved-reports-container").style.display = "block";
  document.getElementById("report-viewer").style.display = "none";
}

function backToReportsList() {
  showSavedReportsList();
}

function openMostRecentReport() {
  const savedReports = JSON.parse(localStorage.getItem("savedReports") || "[]");

  if (savedReports.length > 0) {
    // Get the most recent report (first in array since they're stored with newest first)
    const mostRecentReport = savedReports[0];
    console.log("Opening most recent report:", mostRecentReport.title);
    viewSavedReport(mostRecentReport.id);
  } else {
    console.log("No saved reports found to open automatically");
  }
}

// Helper function to check if table data has meaningful content
function hasTableDataContent(tableData) {
  if (!tableData || !tableData.rows || tableData.rows.length === 0) {
    return false;
  }

  // Check if any row has meaningful content
  return tableData.rows.some((row) => {
    if (!row || row.length === 0) return false;
    // Check if any cell in the row has non-empty content
    return row.some((cell) => cell && cell.trim() !== "");
  });
}

// Function to render a saved report in view-only mode
function renderViewOnlyReport(reportData) {
  const reportContent = document.getElementById("viewed-report-content");
  reportContent.innerHTML = "";
  reportContent.className = "viewed-report-content view-only";

  // Create sections based on the saved data structure
  Object.keys(reportData).forEach((sectionId) => {
    const sectionData = reportData[sectionId];

    // Create main section
    const sectionElement = document.createElement("section");
    sectionElement.className = "main-section";
    sectionElement.id = `view-${sectionId}`;

    // Section header
    const sectionHeader = document.createElement("div");
    sectionHeader.className = "section-header";
    sectionHeader.innerHTML = `<h2>${
      sectionData.title || sectionId.toUpperCase()
    }</h2>`;
    sectionElement.appendChild(sectionHeader);

    // Section content
    const sectionContent = document.createElement("div");
    sectionContent.className = "section-content";

    // Process subsections
    let hasAnySubsections = false;
    if (sectionData.subsections) {
      Object.keys(sectionData.subsections).forEach((subsectionId) => {
        const subsectionData = sectionData.subsections[subsectionId];

        // Create subsection
        const subsectionElement = document.createElement("div");
        subsectionElement.className = "subsection";
        subsectionElement.id = `view-${subsectionId}`;

        // Subsection header
        const subsectionHeader = document.createElement("div");
        subsectionHeader.className = "subsection-header";
        subsectionHeader.innerHTML = `<h3>${
          subsectionData.title || subsectionId.toUpperCase().replace("-", " ")
        }</h3>`;
        subsectionElement.appendChild(subsectionHeader);

        // Tables container
        const tablesContainer = document.createElement("div");
        tablesContainer.className = "tables-container";
        tablesContainer.id = `view-${subsectionId}-tables`;

        // Process tables (only if they have meaningful content)
        let hasAnyTables = false;

        if (subsectionData.tables && subsectionData.tables.length > 0) {
          subsectionData.tables.forEach((tableData, tableIndex) => {
            // Check if table has meaningful content before rendering
            if (hasTableDataContent(tableData)) {
              const tableWrapper = createViewOnlyTable(
                subsectionId,
                tableData,
                tableIndex
              );
              tablesContainer.appendChild(tableWrapper);
              hasAnyTables = true;
            }
          });
        }

        // Only add the subsection if it has tables with content
        if (hasAnyTables) {
          subsectionElement.appendChild(tablesContainer);
          sectionContent.appendChild(subsectionElement);
          hasAnySubsections = true;
        }
      });
    }

    // Only add the section if it has subsections with content
    if (hasAnySubsections) {
      sectionElement.appendChild(sectionContent);
      reportContent.appendChild(sectionElement);
    }
  });

  // Add rich text functionality to the rendered content
  setTimeout(() => {
    addRichTextFunctionality();
  }, 100);
}

// Function to create view-only table from saved data
function createViewOnlyTable(subsectionId, tableData, tableIndex) {
  const tableWrapper = document.createElement("div");
  tableWrapper.className = "table-wrapper view-only-table";
  tableWrapper.id = `view-table-${subsectionId}-${tableIndex}`;

  // Determine table type based on subsection
  const subsectionToTableType = {
    "wow-performance": "performance-trends",
    "ab-test": "ab-test-updates",
    "newly-onboarded": "newly-onboarded-publishers",
    "in-progress": "in-progress-publishers",
    nps: "nps-data",
    "recently-churned": "churned-data",
    "gave-notice": "notice-data",
    "publisher-issues": "publisher-updates",
  };

  const tableType = subsectionToTableType[subsectionId];

  let tableHTML = "";

  if (tableType === "performance-trends") {
    tableHTML = createViewOnlyPerformanceTrendsTable(tableData);
  } else if (tableType === "ab-test-updates") {
    tableHTML = createViewOnlyABTestTable(tableData);
  } else if (tableType === "newly-onboarded-publishers") {
    tableHTML = createViewOnlyNewlyOnboardedTable(tableData);
  } else if (tableType === "in-progress-publishers") {
    tableHTML = createViewOnlyInProgressTable(tableData);
  } else if (tableType === "publisher-updates") {
    tableHTML = createViewOnlyPublisherUpdatesTable(tableData);
  } else {
    // Create standard div-based table for other types
    tableHTML = createViewOnlyStandardTable(tableData, tableType);
  }

  // Add mobile cards for view-only mode
  const mobileCards = generateViewOnlyMobileCards(tableData, tableType);

  tableWrapper.innerHTML = tableHTML + mobileCards;

  return tableWrapper;
}

// Function to create view-only performance trends table
function createViewOnlyPerformanceTrendsTable(tableData) {
  if (!tableData.headers || !tableData.rows) return "";

  let tableHTML = `
    <div class="performance-trends-table view-only">
      <div class="performance-header">
        ${tableData.headers
          .map(
            (header) => `
          <div class="performance-header-cell">${header || ""}</div>
        `
          )
          .join("")}
      </div>
      <div class="performance-body">
  `;

  tableData.rows.forEach((row, rowIndex) => {
    if (row && row.length > 0) {
      if (
        rowIndex === 3 &&
        row[0] &&
        row[0].toUpperCase().includes("OBSERVATIONS")
      ) {
        // Special handling for observations row
        tableHTML += `
          <div class="performance-row observations-row">
            <div class="performance-cell metric-label">${row[0] || ""}</div>
            <div class="performance-cell observations-cell" colspan="2">${
              row[1] || ""
            }</div>
          </div>
        `;
      } else {
        // Regular metric row
        tableHTML += `
          <div class="performance-row">
            ${row
              .map(
                (cell, cellIndex) => `
              <div class="performance-cell ${
                cellIndex === 0 ? "metric-label" : "metric-value"
              }">${cell || ""}</div>
            `
              )
              .join("")}
          </div>
        `;
      }
    }
  });

  tableHTML += `
      </div>
    </div>
  `;

  return tableHTML;
}

// Function to create view-only A/B test table
function createViewOnlyABTestTable(tableData) {
  if (!tableData.headers || !tableData.rows) return "";

  // Ensure A/B test data has the complete structure with TIER, % SPLIT, and CSM/OBS
  const enhancedTableData = ensureABTestStructure(tableData);

  let tableHTML = `
    <div class="ab-test-table view-only">
      <div class="ab-test-header">
        ${enhancedTableData.headers
          .map(
            (header) => `
          <div class="ab-test-header-cell">${header || ""}</div>
        `
          )
          .join("")}
      </div>
      <div class="ab-test-body">
  `;

  enhancedTableData.rows.forEach((row) => {
    if (row && row.length > 0) {
      const rowLabel = row[0] ? row[0].trim().toUpperCase() : "";

      // Handle special rows that should span multiple columns
      if (rowLabel === "STAKEHOLDERS" || rowLabel === "TEST STATUS") {
        tableHTML += `
          <div class="ab-test-row ${rowLabel
            .toLowerCase()
            .replace(" ", "-")}-row">
            <div class="ab-test-cell ab-label">${row[0]}</div>
            <div class="ab-test-cell data-cell spanning-cell">${
              row[1] || ""
            }</div>
          </div>
        `;
      } else {
        // Regular rows - only include cells with content, no empty divs
        if (rowLabel === "PUBLISHER") {
          // PUBLISHER row with wider TIER field
          tableHTML += `
            <div class="ab-test-row publisher-row view-only">
              <div class="ab-test-cell ab-label">${row[0] || ""}</div>
              <div class="ab-test-cell data-cell">${row[1] || ""}</div>
              <div class="ab-test-cell ab-label">${row[2] || ""}</div>
              <div class="ab-test-cell data-cell tier-cell-wide">${
                row[3] || ""
              }</div>
            </div>
          `;
        } else {
          // Other regular rows (COMPETITOR, STAGE) - only include the 4 meaningful cells
          tableHTML += `
            <div class="ab-test-row view-only">
              <div class="ab-test-cell ab-label">${row[0] || ""}</div>
              <div class="ab-test-cell data-cell">${row[1] || ""}</div>
              <div class="ab-test-cell ab-label">${row[2] || ""}</div>
              <div class="ab-test-cell data-cell">${row[3] || ""}</div>
            </div>
          `;
        }
      }
    }
  });

  tableHTML += `
      </div>
    </div>
  `;

  return tableHTML;
}

// Function to create view-only newly onboarded table
function createViewOnlyNewlyOnboardedTable(tableData) {
  if (!tableData.headers || !tableData.rows) return "";

  // Ensure newly onboarded data has the complete structure
  const enhancedTableData = ensureNewlyOnboardedStructure(tableData);

  let tableHTML = `
    <div class="newly-onboarded-table view-only">
      <div class="newly-onboarded-body">
  `;

  enhancedTableData.rows.forEach((row) => {
    if (row && row.length > 0) {
      const rowLabel = row[0] ? row[0].trim().toUpperCase() : "";

      // Handle special rows that should span multiple columns
      if (rowLabel === "DEMAND STATUS") {
        tableHTML += `
          <div class="newly-onboarded-row demand-status-row view-only">
            <div class="newly-onboarded-cell publisher-label">${
              row[0] || ""
            }</div>
            <div class="newly-onboarded-cell data-cell spanning-cell">${
              row[1] || ""
            }</div>
          </div>
        `;
      } else {
        // Regular rows - show all 4 cells with proper labels
        // Expected structure: PUBLISHER | data | LIVE DATE | data
        //                    PROJECTED REVENUE | data | OBS | data
        tableHTML += `
          <div class="newly-onboarded-row view-only">
            <div class="newly-onboarded-cell publisher-label">${
              row[0] || ""
            }</div>
            <div class="newly-onboarded-cell data-cell">${row[1] || ""}</div>
            <div class="newly-onboarded-cell publisher-label">${
              row[2] || ""
            }</div>
            <div class="newly-onboarded-cell data-cell">${row[3] || ""}</div>
          </div>
        `;
      }
    }
  });

  tableHTML += `
      </div>
    </div>
  `;

  return tableHTML;
}

// Function to create view-only in-progress table
function createViewOnlyInProgressTable(tableData) {
  if (!tableData.headers || !tableData.rows) return "";

  // Ensure in-progress data has the complete structure
  const enhancedTableData = ensureInProgressStructure(tableData);

  let tableHTML = `
    <div class="in-progress-table view-only">
      <div class="in-progress-body">
  `;

  enhancedTableData.rows.forEach((row) => {
    if (row && row.length > 0) {
      const rowLabel = row[0] ? row[0].trim().toUpperCase() : "";

      // Handle special rows that should span multiple columns
      if (rowLabel === "STATUS") {
        tableHTML += `
          <div class="in-progress-row status-row view-only">
            <div class="in-progress-cell publisher-label">${row[0] || ""}</div>
            <div class="in-progress-cell data-cell spanning-cell">${
              row[1] || ""
            }</div>
          </div>
        `;
      } else {
        // Regular rows - show all 4 cells with proper labels
        // Expected structure: PUBLISHER | data | EXP. LIVE DATE | data
        //                    STAGE | data | OBS | data
        //                    PROJECTED REVENUE | data | LAST COMM. DATE | data
        tableHTML += `
          <div class="in-progress-row view-only">
            <div class="in-progress-cell publisher-label">${row[0] || ""}</div>
            <div class="in-progress-cell data-cell">${row[1] || ""}</div>
            <div class="in-progress-cell publisher-label">${row[2] || ""}</div>
            <div class="in-progress-cell data-cell">${row[3] || ""}</div>
          </div>
        `;
      }
    }
  });

  tableHTML += `
      </div>
    </div>
  `;

  return tableHTML;
}

// Function to create view-only publisher updates table
function createViewOnlyPublisherUpdatesTable(tableData) {
  if (!tableData.headers || !tableData.rows) return "";

  // Ensure publisher updates data has the complete structure
  const enhancedTableData = ensurePublisherUpdatesStructure(tableData);

  let tableHTML = `
    <div class="publisher-updates-table view-only">
      <div class="publisher-updates-body">
  `;

  enhancedTableData.rows.forEach((row, index) => {
    if (row && row.length > 0) {
      const rowLabel = row[0] ? row[0].trim().toUpperCase() : "";

      // Handle special rows that should span multiple columns
      if (rowLabel === "ISSUE(S)/ UPDATE(S)" || rowLabel === "NEXT STEPS") {
        tableHTML += `
          <div class="publisher-updates-row spanning-row view-only">
            <div class="publisher-updates-cell publisher-label">${
              row[0] || ""
            }</div>
            <div class="publisher-updates-cell data-cell spanning-cell">${
              row[1] || ""
            }</div>
          </div>
        `;
      } else {
        // Regular rows - show all cells with proper labels
        // Expected structure: PUBLISHER | data | TIER | data | AVG. MONTHLY REV. | data
        //                    STAGE | data (spans 3) | CSM | data
        //                    STAKEHOLDERS | data (spans 3) | LAST COMM. DATE | data
        if (rowLabel === "STAGE") {
          // STAGE row: STAGE | data | CSM | data
          tableHTML += `
            <div class="publisher-updates-row view-only">
              <div class="publisher-updates-cell publisher-label">STAGE</div>
              <div class="publisher-updates-cell data-cell">${
                row[1] || ""
              }</div>
              <div class="publisher-updates-cell publisher-label">CSM</div>
              <div class="publisher-updates-cell data-cell">${
                row[4] || ""
              }</div>
            </div>
          `;
        } else if (rowLabel === "STAKEHOLDERS") {
          // STAKEHOLDERS row: STAKEHOLDERS | data | LAST COMM. DATE | data
          tableHTML += `
            <div class="publisher-updates-row view-only">
              <div class="publisher-updates-cell publisher-label">STAKEHOLDERS</div>
              <div class="publisher-updates-cell data-cell">${
                row[1] || ""
              }</div>
              <div class="publisher-updates-cell publisher-label">LAST COMM. DATE</div>
              <div class="publisher-updates-cell data-cell">${
                row[4] || ""
              }</div>
            </div>
          `;
        } else {
          // Regular 6-column rows (PUBLISHER row)
          // Structure: PUBLISHER | data | TIER | data | AVG. MONTHLY REV. | data
          tableHTML += `
            <div class="publisher-updates-row view-only">
              <div class="publisher-updates-cell publisher-label">${
                row[0] || ""
              }</div>
              <div class="publisher-updates-cell data-cell">${
                row[1] || ""
              }</div>
              <div class="publisher-updates-cell publisher-label">${
                row[2] || ""
              }</div>
              <div class="publisher-updates-cell data-cell">${
                row[3] || ""
              }</div>
              <div class="publisher-updates-cell publisher-label">${
                row[4] || ""
              }</div>
              <div class="publisher-updates-cell data-cell">${
                row[5] || ""
              }</div>
            </div>
          `;
        }
      }
    }
  });

  tableHTML += `
      </div>
    </div>
  `;

  return tableHTML;
}

// Function to ensure A/B test data has the complete structure
function ensureABTestStructure(tableData) {
  // Create a copy of the table data to avoid modifying the original
  const enhancedData = {
    headers: [...(tableData.headers || [])],
    rows: [],
  };

  // Ensure headers array has 6 elements
  while (enhancedData.headers.length < 6) {
    enhancedData.headers.push("");
  }

  // Note: Expected A/B test structure is defined inline below

  // Process existing rows and merge with expected structure
  const existingData = {};

  // Extract existing data from saved rows
  if (tableData.rows) {
    tableData.rows.forEach((row) => {
      if (row && row.length > 0 && row[0]) {
        const label = row[0].trim().toUpperCase();
        if (label === "PUBLISHER") {
          existingData.PUBLISHER = row[1] || "";
          // Check for TIER in position 2 or 3
          if (
            row.length > 2 &&
            row[2] &&
            row[2].trim().toUpperCase() === "TIER"
          ) {
            existingData.TIER = row[3] || "";
          } else if (row.length > 3) {
            existingData.TIER = row[3] || "";
          }
        } else if (label === "COMPETITOR") {
          existingData.COMPETITOR = row[1] || "";
          // Check for % SPLIT in position 2 or 3
          if (
            row.length > 2 &&
            row[2] &&
            row[2].trim().toUpperCase() === "% SPLIT"
          ) {
            existingData["% SPLIT"] = row[3] || "";
          } else if (row.length > 3) {
            existingData["% SPLIT"] = row[3] || "";
          }
        } else if (label === "STAGE") {
          existingData.STAGE = row[1] || "";
          // Check for CSM/OBS in position 2 or 3
          if (
            row.length > 2 &&
            row[2] &&
            row[2].trim().toUpperCase() === "CSM/OBS"
          ) {
            existingData["CSM/OBS"] = row[3] || "";
          } else if (row.length > 3) {
            existingData["CSM/OBS"] = row[3] || "";
          }
        } else if (label === "STAKEHOLDERS") {
          existingData.STAKEHOLDERS = row[1] || "";
        } else if (label === "START DATE") {
          existingData["START DATE"] = row[1] || "";
          // Look for END DATE and LAST COMM. DATE in the same row
          for (let i = 2; i < row.length; i += 2) {
            if (row[i] && row[i].trim().toUpperCase() === "END DATE") {
              existingData["END DATE"] = row[i + 1] || "";
            } else if (
              row[i] &&
              row[i].trim().toUpperCase() === "LAST COMM. DATE"
            ) {
              existingData["LAST COMM. DATE"] = row[i + 1] || "";
            }
          }
        } else if (label === "TEST STATUS") {
          existingData["TEST STATUS"] = row[1] || "";
        } else if (label === "TIER") {
          // Handle case where TIER is a separate row
          existingData.TIER = row[1] || "";
        } else if (label === "% SPLIT") {
          // Handle case where % SPLIT is a separate row
          existingData["% SPLIT"] = row[1] || "";
        } else if (label === "CSM/OBS") {
          // Handle case where CSM/OBS is a separate row
          existingData["CSM/OBS"] = row[1] || "";
        } else if (label === "END DATE") {
          // Handle case where END DATE is a separate row
          existingData["END DATE"] = row[1] || "";
        } else if (label === "LAST COMM. DATE") {
          // Handle case where LAST COMM. DATE is a separate row
          existingData["LAST COMM. DATE"] = row[1] || "";
        }
      }
    });
  }

  // Build enhanced rows with complete structure
  enhancedData.rows = [
    [
      "PUBLISHER",
      existingData.PUBLISHER || "",
      "TIER",
      existingData.TIER || "",
      "",
      "",
    ],
    [
      "COMPETITOR",
      existingData.COMPETITOR || "",
      "% SPLIT",
      existingData["% SPLIT"] || "",
      "",
      "",
    ],
    [
      "STAGE",
      existingData.STAGE || "",
      "CSM/OBS",
      existingData["CSM/OBS"] || "",
      "",
      "",
    ],
    ["STAKEHOLDERS", existingData.STAKEHOLDERS || "", "", "", "", ""],
    [
      "START DATE",
      existingData["START DATE"] || "",
      "END DATE",
      existingData["END DATE"] || "",
      "LAST COMM. DATE",
      existingData["LAST COMM. DATE"] || "",
    ],
    ["TEST STATUS", existingData["TEST STATUS"] || "", "", "", "", ""],
  ];

  return enhancedData;
}

// Function to ensure newly onboarded data has the complete structure
function ensureNewlyOnboardedStructure(tableData) {
  // Create a copy of the table data to avoid modifying the original
  const enhancedData = {
    headers: [...(tableData.headers || [])],
    rows: [],
  };

  // Ensure headers array has 4 elements
  while (enhancedData.headers.length < 4) {
    enhancedData.headers.push("");
  }

  // Process existing rows and merge with expected structure
  const existingData = {};

  // Extract existing data from saved rows
  if (tableData.rows) {
    tableData.rows.forEach((row) => {
      if (row && row.length > 0 && row[0]) {
        const label = row[0].trim().toUpperCase();
        if (label === "PUBLISHER") {
          existingData.PUBLISHER = row[1] || "";
          // Check for LIVE DATE in position 2 or 3
          if (
            row.length > 2 &&
            row[2] &&
            row[2].trim().toUpperCase() === "LIVE DATE"
          ) {
            existingData["LIVE DATE"] = row[3] || "";
          } else if (row.length > 3) {
            existingData["LIVE DATE"] = row[3] || "";
          }
        } else if (label === "PROJECTED REVENUE") {
          existingData["PROJECTED REVENUE"] = row[1] || "";
          // Check for OBS in position 2 or 3
          if (
            row.length > 2 &&
            row[2] &&
            row[2].trim().toUpperCase() === "OBS"
          ) {
            existingData.OBS = row[3] || "";
          } else if (row.length > 3) {
            existingData.OBS = row[3] || "";
          }
        } else if (label === "DEMAND STATUS") {
          existingData["DEMAND STATUS"] = row[1] || "";
        } else if (label === "LIVE DATE") {
          // Handle case where LIVE DATE is a separate row
          existingData["LIVE DATE"] = row[1] || "";
        } else if (label === "OBS") {
          // Handle case where OBS is a separate row
          existingData.OBS = row[1] || "";
        }
      }
    });
  }

  // Build enhanced rows with complete structure
  enhancedData.rows = [
    [
      "PUBLISHER",
      existingData.PUBLISHER || "",
      "LIVE DATE",
      existingData["LIVE DATE"] || "",
    ],
    [
      "PROJECTED REVENUE",
      existingData["PROJECTED REVENUE"] || "",
      "OBS",
      existingData.OBS || "",
    ],
    ["DEMAND STATUS", existingData["DEMAND STATUS"] || "", "", ""],
  ];

  return enhancedData;
}

// Function to ensure in-progress data has the complete structure
function ensureInProgressStructure(tableData) {
  // Create a copy of the table data to avoid modifying the original
  const enhancedData = {
    headers: [...(tableData.headers || [])],
    rows: [],
  };

  // Ensure headers array has 4 elements
  while (enhancedData.headers.length < 4) {
    enhancedData.headers.push("");
  }

  // Process existing rows and merge with expected structure
  const existingData = {};

  // Extract existing data from saved rows
  if (tableData.rows) {
    tableData.rows.forEach((row) => {
      if (row && row.length > 0 && row[0]) {
        const label = row[0].trim().toUpperCase();
        if (label === "PUBLISHER") {
          existingData.PUBLISHER = row[1] || "";
          // Check for EXP. LIVE DATE in position 2 or 3
          if (
            row.length > 2 &&
            row[2] &&
            row[2].trim().toUpperCase() === "EXP. LIVE DATE"
          ) {
            existingData["EXP. LIVE DATE"] = row[3] || "";
          } else if (row.length > 3) {
            existingData["EXP. LIVE DATE"] = row[3] || "";
          }
        } else if (label === "STAGE") {
          existingData.STAGE = row[1] || "";
          // Check for OBS in position 2 or 3
          if (
            row.length > 2 &&
            row[2] &&
            row[2].trim().toUpperCase() === "OBS"
          ) {
            existingData.OBS = row[3] || "";
          } else if (row.length > 3) {
            existingData.OBS = row[3] || "";
          }
        } else if (label === "PROJECTED REVENUE") {
          existingData["PROJECTED REVENUE"] = row[1] || "";
          // Check for LAST COMM. DATE in position 2 or 3
          if (
            row.length > 2 &&
            row[2] &&
            row[2].trim().toUpperCase() === "LAST COMM. DATE"
          ) {
            existingData["LAST COMM. DATE"] = row[3] || "";
          } else if (row.length > 3) {
            existingData["LAST COMM. DATE"] = row[3] || "";
          }
        } else if (label === "STATUS") {
          existingData.STATUS = row[1] || "";
        } else if (label === "EXP. LIVE DATE") {
          // Handle case where EXP. LIVE DATE is a separate row
          existingData["EXP. LIVE DATE"] = row[1] || "";
        } else if (label === "OBS") {
          // Handle case where OBS is a separate row
          existingData.OBS = row[1] || "";
        } else if (label === "LAST COMM. DATE") {
          // Handle case where LAST COMM. DATE is a separate row
          existingData["LAST COMM. DATE"] = row[1] || "";
        }
      }
    });
  }

  // Build enhanced rows with complete structure
  enhancedData.rows = [
    [
      "PUBLISHER",
      existingData.PUBLISHER || "",
      "EXP. LIVE DATE",
      existingData["EXP. LIVE DATE"] || "",
    ],
    ["STAGE", existingData.STAGE || "", "OBS", existingData.OBS || ""],
    [
      "PROJECTED REVENUE",
      existingData["PROJECTED REVENUE"] || "",
      "LAST COMM. DATE",
      existingData["LAST COMM. DATE"] || "",
    ],
    ["STATUS", existingData.STATUS || "", "", ""],
  ];

  return enhancedData;
}

// Function to ensure publisher updates data has the complete structure
function ensurePublisherUpdatesStructure(tableData) {
  // If the data already has the correct structure, just return it
  if (tableData.rows && tableData.rows.length > 0) {
    return tableData;
  }

  // Start with the template structure
  const template = tableTemplates["publisher-updates"];

  // Create enhanced data starting with template
  const enhancedData = {
    headers: [...template.headers],
    rows: JSON.parse(JSON.stringify(template.defaultRows)), // Deep copy
  };

  // Process existing rows and populate the template structure
  const existingData = {};

  // Extract existing data from saved rows
  if (tableData.rows) {
    tableData.rows.forEach((row) => {
      if (row && row.length > 0 && row[0]) {
        const label = row[0].trim().toUpperCase();
        if (label === "PUBLISHER") {
          existingData.PUBLISHER = row[1] || "";
          // Check for TIER in any position after PUBLISHER data
          for (let i = 2; i < row.length; i++) {
            if (
              row[i] &&
              row[i].trim().toUpperCase() === "TIER" &&
              i + 1 < row.length
            ) {
              existingData.TIER = row[i + 1] || "";
              break;
            }
          }
          // Check for AVG. MONTHLY REV. in any position
          for (let i = 2; i < row.length; i++) {
            if (
              row[i] &&
              row[i].trim().toUpperCase() === "AVG. MONTHLY REV." &&
              i + 1 < row.length
            ) {
              existingData["AVG. MONTHLY REV."] = row[i + 1] || "";
              break;
            }
          }
        } else if (label === "STAGE") {
          existingData.STAGE = row[1] || "";
          // Check for CSM in any position after STAGE data
          for (let i = 2; i < row.length; i++) {
            if (
              row[i] &&
              row[i].trim().toUpperCase() === "CSM" &&
              i + 1 < row.length
            ) {
              existingData.CSM = row[i + 1] || "";
              break;
            }
          }
        } else if (label === "STAKEHOLDERS") {
          existingData.STAKEHOLDERS = row[1] || "";
          // Check for LAST COMM. DATE in any position after STAKEHOLDERS data
          for (let i = 2; i < row.length; i++) {
            if (
              row[i] &&
              row[i].trim().toUpperCase() === "LAST COMM. DATE" &&
              i + 1 < row.length
            ) {
              existingData["LAST COMM. DATE"] = row[i + 1] || "";
              break;
            }
          }
        } else if (
          label === "ISSUE(S)/ UPDATE(S)" ||
          label === "ISSUE(S)/UPDATE(S)"
        ) {
          existingData["ISSUE(S)/ UPDATE(S)"] = row[1] || "";
        } else if (label === "NEXT STEPS") {
          existingData["NEXT STEPS"] = row[1] || "";
        } else if (label === "TIER") {
          // Handle case where TIER is a separate row
          existingData.TIER = row[1] || "";
        } else if (label === "AVG. MONTHLY REV.") {
          // Handle case where AVG. MONTHLY REV. is a separate row
          existingData["AVG. MONTHLY REV."] = row[1] || "";
        } else if (label === "CSM") {
          // Handle case where CSM is a separate row
          existingData.CSM = row[1] || "";
        } else if (label === "LAST COMM. DATE") {
          // Handle case where LAST COMM. DATE is a separate row
          existingData["LAST COMM. DATE"] = row[1] || "";
        }
      }
    });
  }

  // Populate the template structure with extracted data
  // Row 0: PUBLISHER row
  if (enhancedData.rows[0]) {
    // Position 1: PUBLISHER data
    enhancedData.rows[0][1] =
      existingData.PUBLISHER || enhancedData.rows[0][1] || "";
    // Position 2: TIER label (keep as "TIER")
    enhancedData.rows[0][2] = "TIER";
    // Position 3: TIER data
    enhancedData.rows[0][3] =
      existingData.TIER || enhancedData.rows[0][3] || "";
    // Position 4: AVG. MONTHLY REV. label (keep as is)
    enhancedData.rows[0][4] = "AVG. MONTHLY REV.";
    // Position 5: AVG. MONTHLY REV. data
    enhancedData.rows[0][5] =
      existingData["AVG. MONTHLY REV."] || enhancedData.rows[0][5] || "";
  }

  // Row 1: STAGE row
  if (enhancedData.rows[1]) {
    // Position 0: STAGE label (keep as "STAGE")
    enhancedData.rows[1][0] = "STAGE";
    // Position 1: STAGE data
    enhancedData.rows[1][1] =
      existingData.STAGE || enhancedData.rows[1][1] || "";
    // Position 3: CSM label (keep as "CSM")
    enhancedData.rows[1][3] = "CSM";
    // Position 4: CSM data
    enhancedData.rows[1][4] = existingData.CSM || enhancedData.rows[1][4] || "";
  }

  // Row 2: STAKEHOLDERS row
  if (enhancedData.rows[2]) {
    // Position 0: STAKEHOLDERS label (keep as "STAKEHOLDERS")
    enhancedData.rows[2][0] = "STAKEHOLDERS";
    // Position 1: STAKEHOLDERS data
    enhancedData.rows[2][1] =
      existingData.STAKEHOLDERS || enhancedData.rows[2][1] || "";
    // Position 3: LAST COMM. DATE label (keep as "LAST COMM. DATE")
    enhancedData.rows[2][3] = "LAST COMM. DATE";
    // Position 4: LAST COMM. DATE data
    enhancedData.rows[2][4] =
      existingData["LAST COMM. DATE"] || enhancedData.rows[2][4] || "";
  }

  // Row 3: ISSUE(S)/ UPDATE(S) row
  if (enhancedData.rows[3]) {
    // Position 0: ISSUE(S)/ UPDATE(S) label (keep as is)
    enhancedData.rows[3][0] = "ISSUE(S)/ UPDATE(S)";
    // Position 1: ISSUE(S)/ UPDATE(S) data
    enhancedData.rows[3][1] =
      existingData["ISSUE(S)/ UPDATE(S)"] || enhancedData.rows[3][1] || "";
  }

  // Row 4: NEXT STEPS row
  if (enhancedData.rows[4]) {
    // Position 0: NEXT STEPS label (keep as is)
    enhancedData.rows[4][0] = "NEXT STEPS";
    // Position 1: NEXT STEPS data
    enhancedData.rows[4][1] =
      existingData["NEXT STEPS"] || enhancedData.rows[4][1] || "";
  }

  return enhancedData;
}

// Function to create view-only standard table
function createViewOnlyStandardTable(tableData, tableType) {
  if (!tableData.headers || !tableData.rows) return "";

  let tableHTML = `
    <div class="data-table-div view-only ${tableType}-table">
      <div class="table-header">
        ${tableData.headers
          .map(
            (header) => `
          <div class="header-cell">${header || ""}</div>
        `
          )
          .join("")}
      </div>
      <div class="table-body">
  `;

  tableData.rows.forEach((row) => {
    if (row && row.length > 0) {
      tableHTML += `
        <div class="table-row">
          ${row
            .map(
              (cell, cellIndex) => `
            <div class="table-cell ${
              cellIndex === 0 ? "label-cell" : "data-cell"
            }">${cell || ""}</div>
          `
            )
            .join("")}
        </div>
      `;
    }
  });

  tableHTML += `
      </div>
    </div>
  `;

  return tableHTML;
}

// Function to generate view-only mobile cards
function generateViewOnlyMobileCards(tableData, tableType) {
  if (!tableData.headers || !tableData.rows) return "";

  let mobileHTML = "";

  if (tableType === "performance-trends") {
    mobileHTML = generateViewOnlyPerformanceTrendsCard(tableData);
  } else if (tableType === "ab-test-updates") {
    mobileHTML = generateViewOnlyABTestCard(tableData);
  } else if (tableType === "newly-onboarded-publishers") {
    mobileHTML = generateViewOnlyNewlyOnboardedCard(tableData);
  } else if (tableType === "in-progress-publishers") {
    mobileHTML = generateViewOnlyInProgressCard(tableData);
  } else if (tableType === "publisher-updates") {
    mobileHTML = generateViewOnlyPublisherUpdatesCard(tableData);
  } else if (tableType === "nps-data") {
    mobileHTML = generateViewOnlyNPSCard(tableData);
  } else if (tableType === "churned-data" || tableType === "notice-data") {
    mobileHTML = generateViewOnlyChurnedCard(tableData);
  } else {
    // Default card layout for other table types
    mobileHTML = generateViewOnlyDefaultCard(tableData);
  }

  return mobileHTML;
}

// Generate view-only performance trends card
function generateViewOnlyPerformanceTrendsCard(tableData) {
  let mobileHTML = `<div class="mobile-table-card view-only">`;

  // Add date headers
  mobileHTML += `<div class="mobile-table-header">Performance Trends</div>`;
  mobileHTML += `<div class="mobile-date-headers">`;
  mobileHTML += `<div class="mobile-date-header prior">${
    tableData.headers[1] || ""
  }</div>`;
  mobileHTML += `<div class="mobile-date-header recent">${
    tableData.headers[2] || ""
  }</div>`;
  mobileHTML += `</div>`;

  mobileHTML += `<div class="mobile-table-content">`;

  // Process each metric row
  tableData.rows.forEach((row, index) => {
    if (row[0] && row[0].trim() !== "") {
      if (index === 3 && row[0].toUpperCase().includes("OBSERVATIONS")) {
        // Special handling for observations row
        mobileHTML += `<div class="observations-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
        mobileHTML += `<div class="mobile-row-data observations">${
          row[1] || ""
        }</div>`;
        mobileHTML += `</div>`;
      } else {
        // Regular metric row with two values
        mobileHTML += `<div class="mobile-metric-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
        mobileHTML += `<div class="mobile-metric-values">`;
        mobileHTML += `<div class="mobile-metric-value prior">${
          row[1] || ""
        }</div>`;
        mobileHTML += `<div class="mobile-metric-value recent">${
          row[2] || ""
        }</div>`;
        mobileHTML += `</div>`;
        mobileHTML += `</div>`;
      }
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate view-only A/B test card
function generateViewOnlyABTestCard(tableData) {
  let mobileHTML = `<div class="mobile-table-card view-only">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Ensure A/B test data has the complete structure with TIER, % SPLIT, and CSM/OBS
  const enhancedTableData = ensureABTestStructure(tableData);

  // Process A/B test data structure with proper 6-column layout
  // Expected structure: PUBLISHER | value | TIER | value | |
  //                    COMPETITOR | value | % SPLIT | value | |
  //                    STAGE | value | CSM/OBS | value | |
  //                    STAKEHOLDERS | value | | | |
  //                    START DATE | value | END DATE | value | LAST COMM. DATE | value
  //                    TEST STATUS | value | | | |

  enhancedTableData.rows.forEach((row) => {
    if (row && row.length > 0 && row[0] && row[0].trim() !== "") {
      // First column pair (positions 0,1)
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data">${row[1] || ""}</div>`;
      mobileHTML += `</div>`;

      // Second column pair (positions 2,3) - TIER, % SPLIT, CSM/OBS
      if (row[2] && row[2].trim() !== "") {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[2]}</div>`;
        mobileHTML += `<div class="mobile-row-data">${row[3] || ""}</div>`;
        mobileHTML += `</div>`;
      }

      // Third column pair (positions 4,5) - for START DATE row: LAST COMM. DATE
      if (row[4] && row[4].trim() !== "") {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[4]}</div>`;
        mobileHTML += `<div class="mobile-row-data">${row[5] || ""}</div>`;
        mobileHTML += `</div>`;
      }
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate view-only default card for simple tables
function generateViewOnlyDefaultCard(tableData) {
  let mobileHTML = `<div class="mobile-table-card view-only">`;
  mobileHTML += `<div class="mobile-table-content">`;

  tableData.rows.forEach((row) => {
    if (row && row.length >= 2 && row[0] && row[0].trim() !== "") {
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data">${row[1] || ""}</div>`;
      mobileHTML += `</div>`;
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate view-only newly onboarded card
function generateViewOnlyNewlyOnboardedCard(tableData) {
  let mobileHTML = `<div class="mobile-table-card view-only">`;
  mobileHTML += `<div class="mobile-table-content">`;

  tableData.rows.forEach((row) => {
    if (row && row.length > 0 && row[0] && row[0].trim() !== "") {
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data">${row[1] || ""}</div>`;
      mobileHTML += `</div>`;

      if (row[2] && row[2].trim() !== "") {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[2]}</div>`;
        mobileHTML += `<div class="mobile-row-data">${row[3] || ""}</div>`;
        mobileHTML += `</div>`;
      }
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate view-only in-progress card
function generateViewOnlyInProgressCard(tableData) {
  let mobileHTML = `<div class="mobile-table-card view-only">`;
  mobileHTML += `<div class="mobile-table-content">`;

  tableData.rows.forEach((row) => {
    if (row && row.length > 0 && row[0] && row[0].trim() !== "") {
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-row-data">${row[1] || ""}</div>`;
      mobileHTML += `</div>`;

      if (row[2] && row[2].trim() !== "") {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">${row[2]}</div>`;
        mobileHTML += `<div class="mobile-row-data">${row[3] || ""}</div>`;
        mobileHTML += `</div>`;
      }
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate view-only publisher updates card
function generateViewOnlyPublisherUpdatesCard(tableData) {
  let mobileHTML = `<div class="mobile-table-card view-only">`;
  mobileHTML += `<div class="mobile-table-content">`;

  // Expected fields: PUBLISHER, TIER, AVG. MONTHLY REV, STAGE, CSM, STAKEHOLDERS, LAST COMM. DATE, ISSUE(S)/UPDATE(S), NEXT STEPS
  const fieldLabels = [
    "PUBLISHER",
    "TIER",
    "AVG. MONTHLY REV",
    "STAGE",
    "CSM",
    "STAKEHOLDERS",
    "LAST COMM. DATE",
    "ISSUE(S)/UPDATE(S)",
    "NEXT STEPS",
  ];

  tableData.rows.forEach((row) => {
    if (row && row.length > 0) {
      row.forEach((cell, index) => {
        if (cell && cell.trim() !== "" && fieldLabels[index]) {
          mobileHTML += `<div class="mobile-row">`;
          mobileHTML += `<div class="mobile-row-label">${fieldLabels[index]}</div>`;
          mobileHTML += `<div class="mobile-row-data">${cell}</div>`;
          mobileHTML += `</div>`;
        }
      });
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate view-only NPS card
function generateViewOnlyNPSCard(tableData) {
  let mobileHTML = `<div class="mobile-table-card view-only">`;
  mobileHTML += `<div class="mobile-table-header">Net Promoter Score (NPS)</div>`;
  mobileHTML += `<div class="mobile-table-content">`;

  tableData.rows.forEach((row) => {
    if (row && row.length > 0 && row[0] && row[0].trim() !== "") {
      mobileHTML += `<div class="mobile-nps-row">`;
      mobileHTML += `<div class="mobile-row-label">${row[0]}</div>`;
      mobileHTML += `<div class="mobile-nps-scores">`;
      mobileHTML += `<div class="mobile-nps-score"><span class="score-label">MTD:</span> ${
        row[1] || ""
      }</div>`;
      mobileHTML += `<div class="mobile-nps-score"><span class="score-label">QTD:</span> ${
        row[2] || ""
      }</div>`;
      mobileHTML += `<div class="mobile-nps-score"><span class="score-label">YTD:</span> ${
        row[3] || ""
      }</div>`;
      mobileHTML += `</div>`;
      mobileHTML += `</div>`;
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Generate view-only churned card
function generateViewOnlyChurnedCard(tableData) {
  let mobileHTML = `<div class="mobile-table-card view-only">`;
  mobileHTML += `<div class="mobile-table-content">`;

  tableData.rows.forEach((row) => {
    if (row && row.length > 0 && row[0] && row[0].trim() !== "") {
      mobileHTML += `<div class="mobile-churned-item">`;
      mobileHTML += `<div class="mobile-row">`;
      mobileHTML += `<div class="mobile-row-label">PUBLISHER</div>`;
      mobileHTML += `<div class="mobile-row-data">${row[0] || ""}</div>`;
      mobileHTML += `</div>`;

      if (row[1]) {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">SITE(S)</div>`;
        mobileHTML += `<div class="mobile-row-data">${row[1]}</div>`;
        mobileHTML += `</div>`;
      }

      if (row[2]) {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">AVG. MONTHLY REV.</div>`;
        mobileHTML += `<div class="mobile-row-data">${row[2]}</div>`;
        mobileHTML += `</div>`;
      }

      if (row[3]) {
        mobileHTML += `<div class="mobile-row">`;
        mobileHTML += `<div class="mobile-row-label">DARK DATE & REASON</div>`;
        mobileHTML += `<div class="mobile-row-data">${row[3]}</div>`;
        mobileHTML += `</div>`;
      }

      mobileHTML += `</div>`;
    }
  });

  mobileHTML += `</div>`;
  mobileHTML += `</div>`;

  return mobileHTML;
}

// Custom Modal Functions
function showCustomAlert(message, title = "Information", icon = "info") {
  return new Promise((resolve) => {
    const modal = document.getElementById("info-modal");
    const titleElement = document.getElementById("info-modal-title");
    const messageElement = document.getElementById("info-modal-message");
    const iconElement = document.getElementById("info-modal-icon");
    const okButton = document.getElementById("info-modal-ok");

    titleElement.textContent = title;
    messageElement.textContent = message;

    // Set icon based on type
    iconElement.className = "material-icons";
    switch (icon) {
      case "success":
        iconElement.textContent = "check_circle";
        iconElement.classList.remove("error", "warning");
        break;
      case "error":
        iconElement.textContent = "error";
        iconElement.classList.add("error");
        break;
      case "warning":
        iconElement.textContent = "warning";
        iconElement.classList.add("warning");
        break;
      default:
        iconElement.textContent = "info";
        iconElement.classList.remove("error", "warning");
    }

    modal.classList.add("show");

    const handleOk = () => {
      modal.classList.remove("show");
      okButton.removeEventListener("click", handleOk);
      resolve(true);
    };

    okButton.addEventListener("click", handleOk);

    // Close on overlay click
    modal.addEventListener("click", (e) => {
      if (e.target === modal) {
        handleOk();
      }
    });

    // Close on Escape key
    const handleEscape = (e) => {
      if (e.key === "Escape") {
        handleOk();
        document.removeEventListener("keydown", handleEscape);
      }
    };
    document.addEventListener("keydown", handleEscape);
  });
}

function showCustomConfirm(message, title = "Confirmation") {
  return new Promise((resolve) => {
    const modal = document.getElementById("custom-modal");
    const titleElement = document.getElementById("modal-title");
    const messageElement = document.getElementById("modal-message");
    const confirmButton = document.getElementById("modal-confirm");
    const cancelButton = document.getElementById("modal-cancel");

    titleElement.textContent = title;
    messageElement.textContent = message;

    modal.classList.add("show");

    const handleConfirm = () => {
      modal.classList.remove("show");
      confirmButton.removeEventListener("click", handleConfirm);
      cancelButton.removeEventListener("click", handleCancel);
      resolve(true);
    };

    const handleCancel = () => {
      modal.classList.remove("show");
      confirmButton.removeEventListener("click", handleConfirm);
      cancelButton.removeEventListener("click", handleCancel);
      resolve(false);
    };

    confirmButton.addEventListener("click", handleConfirm);
    cancelButton.addEventListener("click", handleCancel);

    // Close on overlay click (acts as cancel)
    modal.addEventListener("click", (e) => {
      if (e.target === modal) {
        handleCancel();
      }
    });

    // Close on Escape key (acts as cancel)
    const handleEscape = (e) => {
      if (e.key === "Escape") {
        handleCancel();
        document.removeEventListener("keydown", handleEscape);
      }
    };
    document.addEventListener("keydown", handleEscape);
  });
}

// Function to add a row to NPS or churned data tables
function addRowToTable(button) {
  console.log("addRowToTable called");
  const tableWrapper = button.closest(".table-wrapper");

  // Check if this is a div-based table
  const npsTable = tableWrapper.querySelector(".nps-table");
  const churnedTable = tableWrapper.querySelector(".churned-table");

  if (npsTable) {
    // Handle div-based NPS table
    addRowToNPSDiv(button, npsTable);
    return;
  }

  if (churnedTable) {
    // Handle div-based churned table
    addRowToChurnedDiv(button, churnedTable);
    return;
  }

  // Handle traditional HTML tables
  const table = button.closest("table");
  const tbody = table.querySelector("tbody");
  const tableId = tableWrapper.id;

  // Determine table type based on classes
  const isNpsTable = table.querySelector(".nps-label") !== null;
  const isChurnedTable = table.querySelector(".churned-publisher") !== null;

  console.log("isNpsTable:", isNpsTable);

  // Get the row that contains the clicked button
  const clickedRow = button.closest("tr");
  console.log("clickedRow:", clickedRow);

  // Create new row
  const newRow = document.createElement("tr");

  if (isNpsTable) {
    // Get the label from the clicked row to determine if it's NETWORK or T1
    const clickedLabel = clickedRow
      .querySelector(".nps-label")
      .textContent.trim();

    console.log("clickedLabel:", clickedLabel);

    newRow.className = "nps-added-row";
    newRow.innerHTML = `
      <td class="nps-label">${clickedLabel}</td>
      <td contenteditable="true" class="nps-score"></td>
      <td contenteditable="true" class="nps-score"></td>
      <td contenteditable="true" class="nps-score"></td>
      <td class="row-actions">
        <button class="remove-row-btn" onclick="removeRowFromTable(this)" title="Remove row"><span class="material-icons">remove</span></button>
      </td>
    `;

    // Insert the new row directly after the clicked row
    clickedRow.insertAdjacentElement("afterend", newRow);
    console.log("New row inserted");
  } else if (isChurnedTable) {
    newRow.className = "churned-added-row";
    newRow.innerHTML = `
      <td contenteditable="true" class="churned-publisher"></td>
      <td contenteditable="true" class="churned-site"></td>
      <td contenteditable="true" class="churned-revenue"></td>
      <td contenteditable="true" class="churned-reason"></td>
      <td class="row-actions">
        <button class="remove-row-btn" onclick="removeRowFromTable(this)" title="Remove row"><span class="material-icons">remove</span></button>
      </td>
    `;

    // For churned table, add at the end
    tbody.appendChild(newRow);
    console.log("New churned row added");
  }

  // Add event listeners to new editable cells
  const editableCells = newRow.querySelectorAll('td[contenteditable="true"]');
  editableCells.forEach((cell) => {
    addCellEventListeners(cell);
  });

  // Trigger auto-save
  triggerAutoSave();
}

// Function to add a row to div-based NPS table
function addRowToNPSDiv(button, npsTable) {
  console.log("addRowToNPSDiv called");

  // Get the row that contains the clicked button
  const clickedRow = button.closest(".nps-row");
  console.log("clickedRow:", clickedRow);

  // Get the label from the clicked row to determine if it's NETWORK or T1
  const clickedLabel = clickedRow
    .querySelector(".nps-label")
    .textContent.trim();
  console.log("clickedLabel:", clickedLabel);

  // Create new div row
  const newRow = document.createElement("div");
  newRow.className = "nps-row nps-added-row";
  newRow.innerHTML = `
    <div class="nps-label">${clickedLabel}</div>
    <div contenteditable="true" class="nps-score"></div>
    <div contenteditable="true" class="nps-score"></div>
    <div contenteditable="true" class="nps-score"></div>
    <div class="row-actions">
      <button class="remove-row-btn" onclick="removeRowFromNPSDiv(this)" title="Remove row"><span class="material-icons">remove</span></button>
    </div>
  `;

  // Insert the new row directly after the clicked row
  clickedRow.insertAdjacentElement("afterend", newRow);
  console.log("New NPS div row inserted");

  // Add event listeners to new editable cells
  const editableCells = newRow.querySelectorAll('div[contenteditable="true"]');
  editableCells.forEach((cell) => {
    addCellEventListeners(cell);
  });

  // Trigger auto-save
  triggerAutoSave();
}

// Function to add a row to div-based churned table
function addRowToChurnedDiv(button, churnedTable) {
  console.log("addRowToChurnedDiv called");

  // Get the churned body to append the new row
  const churnedBody = churnedTable.querySelector(".churned-body");

  // Create new div row
  const newRow = document.createElement("div");
  newRow.className = "churned-row churned-added-row";
  newRow.innerHTML = `
    <div contenteditable="true" class="churned-publisher"></div>
    <div contenteditable="true" class="churned-site"></div>
    <div contenteditable="true" class="churned-revenue"></div>
    <div contenteditable="true" class="churned-reason"></div>
    <div class="row-actions">
      <button class="remove-row-btn" onclick="removeRowFromChurnedDiv(this)" title="Remove row"><span class="material-icons">remove</span></button>
    </div>
  `;

  // Add the new row at the end
  churnedBody.appendChild(newRow);
  console.log("New churned div row added");

  // Add event listeners to new editable cells
  const editableCells = newRow.querySelectorAll('div[contenteditable="true"]');
  editableCells.forEach((cell) => {
    addCellEventListeners(cell);
  });

  // Trigger auto-save
  triggerAutoSave();
}

// Function to remove a row from div-based NPS table
async function removeRowFromNPSDiv(button) {
  console.log("removeRowFromNPSDiv called");

  const row = button.closest(".nps-row");

  // Don't allow removing default rows (NETWORK and T1)
  if (row.classList.contains("nps-default-row")) {
    await showCustomAlert(
      "Cannot remove default rows (NETWORK and T1). Only added rows can be removed.",
      "Cannot Remove Row",
      "warning"
    );
    return;
  }

  // Remove the row
  row.remove();
  console.log("NPS div row removed");

  // Trigger auto-save
  triggerAutoSave();
}

// Function to remove a row from div-based churned table
async function removeRowFromChurnedDiv(button) {
  console.log("removeRowFromChurnedDiv called");

  const row = button.closest(".churned-row");

  // Don't allow removing default rows
  if (row.classList.contains("churned-default-row")) {
    await showCustomAlert(
      "Cannot remove the default row. Only added rows can be removed.",
      "Cannot Remove Row",
      "warning"
    );
    return;
  }

  // Remove the row
  row.remove();
  console.log("Churned div row removed");

  // Trigger auto-save
  triggerAutoSave();
}

// Function to remove a row from NPS or churned data tables
async function removeRowFromTable(button) {
  const table = button.closest("table");
  const tbody = table.querySelector("tbody");
  const row = button.closest("tr");

  // Determine table type
  const isNpsTable = table.querySelector(".nps-label") !== null;
  const isChurnedTable = table.querySelector(".churned-publisher") !== null;

  // For NPS tables, don't allow removing default rows (NETWORK and T1)
  if (isNpsTable && row.classList.contains("nps-default-row")) {
    await showCustomAlert(
      "Cannot remove default rows (NETWORK and T1). Only added rows can be removed.",
      "Cannot Remove Row",
      "warning"
    );
    return;
  }

  // For churned tables, don't allow removing the default row
  if (isChurnedTable && row.classList.contains("churned-default-row")) {
    await showCustomAlert(
      "Cannot remove the default row. Only added rows can be removed.",
      "Cannot Remove Row",
      "warning"
    );
    return;
  }

  // Don't allow removing if only default rows remain
  if (isNpsTable) {
    const addedRows = tbody.querySelectorAll(".nps-added-row");
    if (addedRows.length <= 1) {
      await showCustomAlert(
        "Cannot remove the last added row. At least the default rows must remain.",
        "Cannot Remove Row",
        "warning"
      );
      return;
    }
  }

  if (isChurnedTable) {
    const addedRows = tbody.querySelectorAll(".churned-added-row");
    if (addedRows.length <= 1) {
      await showCustomAlert(
        "Cannot remove the last added row. At least one row must remain.",
        "Cannot Remove Row",
        "warning"
      );
      return;
    }
  }

  // For other tables, don't allow removing if only one row remains
  if (!isNpsTable && !isChurnedTable && tbody.children.length <= 1) {
    await showCustomAlert(
      "Cannot remove the last row. Tables must have at least one row.",
      "Cannot Remove Row",
      "warning"
    );
    return;
  }

  const confirmed = await showCustomConfirm(
    "Are you sure you want to remove this row?",
    "Remove Row"
  );

  if (confirmed) {
    // Remove the row
    row.remove();

    // Re-add plus/minus buttons to appropriate rows
    const rows = tbody.querySelectorAll("tr");
    rows.forEach((row, index) => {
      const actionCell = row.querySelector(".row-actions");

      if (isNpsTable) {
        // For NPS tables: default rows (NETWORK and T1) get plus buttons, added rows get minus buttons
        const isDefaultRow = row.classList.contains("nps-default-row");
        if (isDefaultRow) {
          actionCell.innerHTML = `<button class="add-row-btn" onclick="addRowToTable(this)" title="Add row"><span class="material-icons">add</span></button>`;
        } else {
          actionCell.innerHTML = `<button class="remove-row-btn" onclick="removeRowFromTable(this)" title="Remove row"><span class="material-icons">remove</span></button>`;
        }
      } else {
        // For other tables: first row gets plus, last row gets minus
        if (index === 0) {
          actionCell.innerHTML = `<button class="add-row-btn" onclick="addRowToTable(this)" title="Add row"><span class="material-icons">add</span></button>`;
        } else if (index === rows.length - 1) {
          actionCell.innerHTML = `<button class="remove-row-btn" onclick="removeRowFromTable(this)" title="Remove row"><span class="material-icons">remove</span></button>`;
        } else {
          actionCell.innerHTML = "";
        }
      }
    });

    // Trigger auto-save
    triggerAutoSave();
  }
}

// Rich Text Editor functionality
let richTextEditor = null;
let currentRichTextField = null;

// Initialize rich text editor (will be called from main initialization)
function initializeRichTextEditor() {
  try {
    richTextEditor = new Quill("#rich-text-editor", {
      theme: "snow",
      modules: {
        toolbar: [
          [{ header: [1, 2, 3, false] }],
          ["bold", "italic", "underline", "strike"],
          [{ list: "ordered" }, { list: "bullet" }],
          [{ indent: "-1" }, { indent: "+1" }],
          [{ color: [] }, { background: [] }],
          [{ align: [] }],
          ["link"],
          ["clean"],
        ],
      },
      placeholder: "Enter your content here...",
    });
  } catch (error) {
    console.error("Error initializing Quill editor:", error);
  }
}

// Function to open rich text editor
function openRichTextEditor(element, fieldName) {
  currentRichTextField = element;

  // Set modal title
  document.getElementById(
    "rich-text-modal-title"
  ).textContent = `Edit ${fieldName}`;

  // Get current content (strip HTML tags for plain text, or keep HTML for rich content)
  const currentContent = element.innerHTML || "";

  // Set content in editor
  richTextEditor.root.innerHTML = currentContent;

  // Show modal
  document.getElementById("rich-text-modal").style.display = "flex";

  // Focus editor
  setTimeout(() => {
    richTextEditor.focus();
  }, 100);
}

// Function to close rich text editor
function closeRichTextEditor() {
  document.getElementById("rich-text-modal").style.display = "none";
  currentRichTextField = null;
}

// Function to save rich text content
function saveRichTextContent() {
  if (currentRichTextField && richTextEditor) {
    // Get HTML content from editor
    const content = richTextEditor.root.innerHTML;

    // Update the field with new content
    currentRichTextField.innerHTML = content;

    // Trigger auto-save
    triggerAutoSave();

    // Close modal
    closeRichTextEditor();
  }
}

// Function to add rich text functionality to specific fields
function addRichTextFunctionality() {
  // Define the selectors for fields that should have rich text editing
  const richTextSelectors = [
    ".observations-cell",
    ".test-status-cell",
    ".demand-status-cell",
    ".progress-status-cell",
    ".updates-issues",
    ".updates-steps",
    ".mobile-row-data.observations",
    ".mobile-row-data.status",
  ];

  // Define field names for modal titles
  const fieldNames = {
    "observations-cell": "Observations",
    "test-status-cell": "Test Status",
    "demand-status-cell": "Demand Status",
    "progress-status-cell": "Status",
    "updates-issues": "Issues/Updates",
    "updates-steps": "Next Steps",
    "mobile-row-data observations": "Content",
    "mobile-row-data status": "Status",
  };

  richTextSelectors.forEach((selector) => {
    const elements = document.querySelectorAll(selector);
    elements.forEach((element) => {
      // Skip elements that are in view-only mode
      if (
        element.closest(".view-only") ||
        element.closest(".viewed-report-content")
      ) {
        return;
      }

      // Add rich text field class and click handler
      element.classList.add("rich-text-field");

      // Add double-click event listener for rich text editor
      element.addEventListener("dblclick", function (e) {
        e.preventDefault();
        e.stopPropagation();

        // Determine field name from class or context
        let fieldName = "Content";

        // Check desktop field classes first
        for (const [className, name] of Object.entries(fieldNames)) {
          if (element.classList.contains(className)) {
            fieldName = name;
            break;
          }
        }

        // For mobile fields, check the label in the same row
        if (element.classList.contains("mobile-row-data")) {
          const row = element.closest(".mobile-row");
          if (row) {
            const label = row.querySelector(".mobile-row-label");
            if (label) {
              const labelText = label.textContent.trim().toUpperCase();
              if (labelText.includes("OBSERVATION")) {
                fieldName = "Observations";
              } else if (labelText.includes("TEST STATUS")) {
                fieldName = "Test Status";
              } else if (labelText.includes("DEMAND STATUS")) {
                fieldName = "Demand Status";
              } else if (labelText === "STATUS") {
                fieldName = "Status";
              } else if (
                labelText.includes("ISSUE") ||
                labelText.includes("UPDATE")
              ) {
                fieldName = "Issues/Updates";
              } else if (labelText.includes("NEXT STEPS")) {
                fieldName = "Next Steps";
              }
            }
          }
        }

        openRichTextEditor(element, fieldName);
      });

      // Add placeholder text if empty
      if (!element.innerHTML.trim()) {
        element.innerHTML =
          '<span style="color: #999; font-style: italic;">Double-click for rich text editor...</span>';
      }
    });
  });
}
