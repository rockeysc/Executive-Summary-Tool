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
    headers: ["", "MTD", "QTD", "YTD"],
    defaultRows: [
      ["NETWORK", "", "", ""],
      ["T1", "", "", ""],
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
    ],
    defaultRows: [["", "", "", ""]],
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
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (
          index === template.specialFormatting.observationsRow &&
          template.specialFormatting.mergeObservations
        ) {
          // Special formatting for observations row - merge last two columns
          return `<tr>
          <td class="metric-label">${row[0]}</td>
          <td contenteditable="true" colspan="2" class="observations-cell">${row[1]}</td>
        </tr>`;
        } else {
          return `<tr>${row
            .map((cell, cellIndex) => {
              if (cellIndex === 0) {
                return `<td class="metric-label">${cell}</td>`;
              } else {
                return `<td contenteditable="true">${cell}</td>`;
              }
            })
            .join("")}</tr>`;
        }
      })
      .join("");
  } else if (
    tableType === "ab-test-updates" &&
    template.specialFormatting &&
    template.specialFormatting.customLayout
  ) {
    // Special formatting for A/B test updates
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (index === 4) {
          // START DATE row with 3 columns - use date inputs
          return `<tr>
            <td class="ab-label">${row[0]}</td>
            <td>${createDateInput(row[1])}</td>
            <td class="ab-label">${row[2]}</td>
            <td>${createDateInput(row[3])}</td>
            <td class="ab-label">${row[4]}</td>
            <td>${createDateInput(row[5])}</td>
          </tr>`;
        } else if (index === template.specialFormatting.statusRow) {
          // TEST STATUS row spans multiple columns
          return `<tr>
            <td class="ab-label">${row[0]}</td>
            <td contenteditable="true" colspan="5" class="test-status-cell">${row[1]}</td>
          </tr>`;
        } else if (index === 3) {
          // STAKEHOLDERS row spans all remaining columns
          return `<tr>
            <td class="ab-label">${row[0]}</td>
            <td contenteditable="true" colspan="5" class="stakeholders-cell">${row[1]}</td>
          </tr>`;
        } else {
          // Regular rows with 4 columns
          return `<tr>
            <td class="ab-label">${row[0]}</td>
            <td contenteditable="true">${row[1]}</td>
            <td class="ab-label">${row[2]}</td>
            <td contenteditable="true" colspan="3">${row[3]}</td>
          </tr>`;
        }
      })
      .join("");
  } else if (
    tableType === "newly-onboarded-publishers" &&
    template.specialFormatting &&
    template.specialFormatting.customLayout
  ) {
    // Special formatting for newly onboarded publishers
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (index === template.specialFormatting.demandStatusRow) {
          // DEMAND STATUS row spans multiple columns
          return `<tr>
            <td class="publisher-label">${row[0]}</td>
            <td contenteditable="true" colspan="3" class="demand-status-cell">${row[1]}</td>
          </tr>`;
        } else {
          // Regular rows with 4 columns
          return `<tr>
            <td class="publisher-label">${row[0]}</td>
            <td contenteditable="true">${row[1]}</td>
            <td class="publisher-label">${row[2]}</td>
            <td>${
              isDateField(row[2], 3, index, tableType)
                ? createDateInput(row[3])
                : `<span contenteditable="true">${row[3]}</span>`
            }</td>
          </tr>`;
        }
      })
      .join("");
  } else if (
    tableType === "in-progress-publishers" &&
    template.specialFormatting &&
    template.specialFormatting.customLayout
  ) {
    // Special formatting for in-progress publishers
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (index === template.specialFormatting.statusRow) {
          // STATUS row spans multiple columns
          return `<tr>
            <td class="progress-label">${row[0]}</td>
            <td contenteditable="true" colspan="3" class="progress-status-cell">${row[1]}</td>
          </tr>`;
        } else {
          // Regular rows with 4 columns
          return `<tr>
            <td class="progress-label">${row[0]}</td>
            <td contenteditable="true">${row[1]}</td>
            <td class="progress-label">${row[2]}</td>
            <td>${
              isDateField(row[2], 3, index, tableType)
                ? createDateInput(row[3])
                : `<span contenteditable="true">${row[3]}</span>`
            }</td>
          </tr>`;
        }
      })
      .join("");
  } else if (
    tableType === "nps-data" &&
    template.specialFormatting &&
    template.specialFormatting.npsTable
  ) {
    // Special formatting for NPS table
    rowsHTML = template.defaultRows
      .map((row, index) => {
        return `<tr>
          <td class="nps-label">${row[0]}</td>
          <td contenteditable="true" class="nps-score">${row[1]}</td>
          <td contenteditable="true" class="nps-score">${row[2]}</td>
          <td contenteditable="true" class="nps-score">${row[3]}</td>
        </tr>`;
      })
      .join("");
  } else if (
    tableType === "churned-data" &&
    template.specialFormatting &&
    template.specialFormatting.churnedTable
  ) {
    // Special formatting for churned data table
    rowsHTML = template.defaultRows
      .map((row, index) => {
        return `<tr>
          <td contenteditable="true" class="churned-publisher">${row[0]}</td>
          <td contenteditable="true" class="churned-site">${row[1]}</td>
          <td contenteditable="true" class="churned-revenue">${row[2]}</td>
          <td contenteditable="true" class="churned-reason">${row[3]}</td>
        </tr>`;
      })
      .join("");
  } else if (
    tableType === "notice-data" &&
    template.specialFormatting &&
    template.specialFormatting.noticeTable
  ) {
    // Special formatting for notice data table
    rowsHTML = template.defaultRows
      .map((row, index) => {
        return `<tr>
          <td contenteditable="true" class="notice-publisher">${row[0]}</td>
          <td contenteditable="true" class="notice-revenue">${row[1]}</td>
          <td contenteditable="true" class="notice-reason">${row[2]}</td>
        </tr>`;
      })
      .join("");
  } else if (
    tableType === "publisher-updates" &&
    template.specialFormatting &&
    template.specialFormatting.publisherUpdatesTable
  ) {
    // Special formatting for publisher updates table
    rowsHTML = template.defaultRows
      .map((row, index) => {
        if (index === 0) {
          // First row: PUBLISHER | name | TIER | 1 | AVG. MONTHLY REV. | 20,000
          return `<tr>
            <td class="updates-label">${row[0]}</td>
            <td contenteditable="true" class="updates-data">${row[1]}</td>
            <td class="updates-label">${row[2]}</td>
            <td contenteditable="true" class="updates-data">${row[3]}</td>
            <td class="updates-label">${row[4]}</td>
            <td contenteditable="true" class="updates-data">${row[5]}</td>
          </tr>`;
        } else if (index === 1) {
          // Second row: STAGE | Onboarding or Ongoing | | | CSM | Kristen
          return `<tr>
            <td class="updates-label">${row[0]}</td>
            <td contenteditable="true" class="updates-data" colspan="3">${row[1]}</td>
            <td class="updates-label">${row[3]}</td>
            <td contenteditable="true" class="updates-data">${row[4]}</td>
          </tr>`;
        } else if (index === 2) {
          // Third row: STAKEHOLDERS | list of those involved | | | LAST COMM. DATE | date
          return `<tr>
            <td class="updates-label">${row[0]}</td>
            <td contenteditable="true" class="updates-data" colspan="3">${
              row[1]
            }</td>
            <td class="updates-label">${row[3]}</td>
            <td class="updates-data">${createDateInput(row[4])}</td>
          </tr>`;
        } else if (index === 3) {
          // Fourth row: ISSUE(S)/ UPDATE(S) spanning full width
          return `<tr>
            <td class="updates-label">${row[0]}</td>
            <td contenteditable="true" class="updates-issues" colspan="5">${row[1]}</td>
          </tr>`;
        } else if (index === 4) {
          // Fifth row: NEXT STEPS spanning full width
          return `<tr>
            <td class="updates-label">${row[0]}</td>
            <td contenteditable="true" class="updates-steps" colspan="5">${row[1]}</td>
          </tr>`;
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

  let tableHTML = `
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
        </table>

        <!-- Mobile card layout for add-new-content -->
        ${generateEditableMobileCards(template, tableType, tableCounter)}

        <div class="table-actions">
            <button class="remove-table-btn" onclick="removeTable('table-${tableCounter}')"><span class="material-icons">delete</span> Remove Table</button>
        </div>
    `;

  tableWrapper.innerHTML = tableHTML;
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

  // Don't show auto-save status in view-only mode
  const viewOnlyContent = document.getElementById("view-only-content");
  if (viewOnlyContent && viewOnlyContent.classList.contains("active")) {
    return;
  }

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
  // Don't trigger auto-save in view-only mode
  const viewOnlyContent = document.getElementById("view-only-content");
  if (viewOnlyContent && viewOnlyContent.classList.contains("active")) {
    return;
  }

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

  // Also save as a historical report only when manually saving
  if (showAlert) {
    saveAsHistoricalReport(data);
  }

  console.log("Data saved to localStorage:", data);
  console.log("localStorage size:", JSON.stringify(data).length, "characters");

  if (showAlert) {
    await showCustomAlert("Report saved successfully!", "Success", "success");
  }
}

// Function to save report as historical/view-only report
function saveAsHistoricalReport(data) {
  const savedReports = JSON.parse(localStorage.getItem("savedReports") || "[]");

  const reportData = {
    id: Date.now().toString(),
    title: `Executive Summary - ${new Date().toLocaleDateString()}`,
    date: new Date().toISOString(),
    data: data,
    summary: generateReportSummary(data),
  };

  savedReports.unshift(reportData); // Add to beginning of array

  // Keep only the last 50 reports to prevent storage overflow
  if (savedReports.length > 50) {
    savedReports.splice(50);
  }

  localStorage.setItem("savedReports", JSON.stringify(savedReports));
}

// Function to generate a summary of the report
function generateReportSummary(data) {
  let summary = "";
  let tableCount = 0;

  Object.keys(data).forEach((sectionKey) => {
    const section = data[sectionKey];
    if (section.subsections) {
      Object.keys(section.subsections).forEach((subsectionKey) => {
        const subsection = section.subsections[subsectionKey];
        if (subsection.tables && subsection.tables.length > 0) {
          tableCount += subsection.tables.length;
        }
      });
    }
  });

  summary = `Contains ${tableCount} table${
    tableCount !== 1 ? "s" : ""
  } across ${Object.keys(data).length} section${
    Object.keys(data).length !== 1 ? "s" : ""
  }`;

  return summary;
}

// Function to extract all data from tables
function extractAllData() {
  const data = {};
  const sections = document.querySelectorAll(".main-section");

  sections.forEach((section) => {
    const sectionId = section.id;
    const sectionTitle =
      section.querySelector(".section-header h2").textContent;
    data[sectionId] = {
      title: sectionTitle,
      subsections: {},
    };

    const subsections = section.querySelectorAll(".subsection");
    subsections.forEach((subsection) => {
      const subsectionId = subsection.id;
      const subsectionTitle = subsection.querySelector(
        ".subsection-header h3"
      ).textContent;
      const tables = subsection.querySelectorAll(".data-table");

      data[sectionId].subsections[subsectionId] = {
        title: subsectionTitle,
        tables: [],
      };

      tables.forEach((table) => {
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
    });
  });

  return data;
}

// Function to load data and restore it to the interface
function loadData() {
  const savedData = localStorage.getItem("executiveSummaryData");
  if (savedData) {
    try {
      const data = JSON.parse(savedData);
      console.log("Loading saved data:", data);

      // Clear existing tables first to prevent duplicates
      clearAllTablesOnly();
      console.log("Cleared tables before loading saved data");

      // Small delay to ensure DOM is ready
      setTimeout(() => {
        let totalTablesRestored = 0;

        // Restore data to each section
        Object.keys(data).forEach((sectionId) => {
          const section = data[sectionId];
          if (section.subsections) {
            Object.keys(section.subsections).forEach((subsectionId) => {
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
      }, 50);
    } catch (error) {
      console.error("Error loading saved data:", error);
      // If loading fails, clear tables to prevent duplicates
      clearAllTablesOnly();
    }
  }
}

// Function to clear all tables without showing confirmation
function clearAllTablesOnly() {
  const tableWrappers = document.querySelectorAll(".table-wrapper");
  console.log(`Clearing ${tableWrappers.length} existing table(s)`);
  tableWrappers.forEach((wrapper) => {
    wrapper.remove();
  });
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

  // Create and populate each table
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

// Function to populate a table with saved data
function populateTableWithData(tableWrapper, tableData, tableType) {
  const table = tableWrapper.querySelector(".data-table");
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

  // Check if we have saved data first
  const savedData = localStorage.getItem("executiveSummaryData");

  if (savedData) {
    console.log("Found saved data, loading...");
    // Load saved data if available
    setTimeout(() => {
      loadData();
    }, 100);
  } else {
    console.log("No saved data found, adding default tables...");
    // Add initial tables for demonstration only if no saved data exists
    setTimeout(() => {
      addTable("wow-performance", "performance-trends");
      addTable("ab-test", "ab-test-updates");
      addTable("newly-onboarded", "newly-onboarded-publishers");
      addTable("in-progress", "in-progress-publishers");
      addTable("nps", "nps-data");
      addTable("recently-churned", "churned-data");
      addTable("gave-notice", "notice-data");
      addTable("publisher-issues", "publisher-updates");
      console.log("Default tables added");
    }, 100);
  }

  // Auto-save when page becomes hidden (user switches tabs or minimizes)
  document.addEventListener("visibilitychange", function () {
    if (document.hidden) {
      // Don't auto-save in view-only mode
      const viewOnlyContent = document.getElementById("view-only-content");
      if (viewOnlyContent && viewOnlyContent.classList.contains("active")) {
        return;
      }

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
    // Don't auto-save in view-only mode
    const viewOnlyContent = document.getElementById("view-only-content");
    if (viewOnlyContent && viewOnlyContent.classList.contains("active")) {
      return;
    }

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
    // Don't auto-save in view-only mode
    const viewOnlyContent = document.getElementById("view-only-content");
    if (viewOnlyContent && viewOnlyContent.classList.contains("active")) {
      return;
    }

    const now = Date.now();
    if (now - lastSaveTime >= 30000) {
      // 30 seconds
      saveData(false);
      lastSaveTime = now;
      showAutoSaveStatus("saved", "Periodic save");
      console.log("Periodic auto-save at:", new Date().toLocaleTimeString());
    }
  }, 30000);
});

// Tab switching functionality
function switchTab(tabName) {
  // Hide all tab contents
  document.querySelectorAll(".tab-content").forEach((content) => {
    content.classList.remove("active");
  });

  // Remove active class from all tab buttons
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.remove("active");
  });

  // Show selected tab content
  document.getElementById(`${tabName}-content`).classList.add("active");
  document.getElementById(`${tabName}-tab`).classList.add("active");

  // Handle button visibility and auto-save status based on tab
  const saveBtn = document.getElementById("save-btn");
  const testBtn = document.getElementById("test-btn");
  const clearBtn = document.getElementById("clear-btn");
  const autoSaveStatus = document.getElementById("auto-save-status");

  if (tabName === "view-only") {
    // Hide buttons and auto-save status in view-only mode
    if (saveBtn) saveBtn.style.display = "none";
    if (testBtn) testBtn.style.display = "none";
    if (clearBtn) clearBtn.style.display = "none";
    if (autoSaveStatus) autoSaveStatus.style.display = "none";

    // Load saved reports
    loadSavedReports();
  } else {
    // Show buttons and auto-save status in add-new mode
    if (saveBtn) saveBtn.style.display = "inline-block";
    if (testBtn) testBtn.style.display = "inline-block";
    if (clearBtn) clearBtn.style.display = "inline-block";
    if (autoSaveStatus) autoSaveStatus.style.display = "flex";
  }
}

// Function to load and display saved reports
function loadSavedReports() {
  const savedReports = JSON.parse(localStorage.getItem("savedReports") || "[]");
  const reportsList = document.getElementById("saved-reports-list");

  if (savedReports.length === 0) {
    reportsList.innerHTML = `
      <div class="empty-reports-message">
        <p>No saved reports yet. Create and save a report from the "Add New Report" tab to see it here.</p>
      </div>
    `;
    return;
  }

  reportsList.innerHTML = savedReports
    .map(
      (report) => `
    <div class="saved-report-card" onclick="viewReport('${report.id}')">
      <button class="delete-report-btn" onclick="event.stopPropagation(); deleteReport('${
        report.id
      }')" title="Delete Report">
        <span class="material-icons">close</span>
      </button>
      <h3>${report.title}</h3>
      <div class="report-date">Saved on ${new Date(
        report.date
      ).toLocaleDateString()} at ${new Date(
        report.date
      ).toLocaleTimeString()}</div>
      <div class="report-summary">${report.summary}</div>
    </div>
  `
    )
    .join("");
}

// Function to view a specific report
async function viewReport(reportId) {
  const savedReports = JSON.parse(localStorage.getItem("savedReports") || "[]");
  const report = savedReports.find((r) => r.id === reportId);

  if (!report) {
    await showCustomAlert("Report not found!", "Error", "error");
    return;
  }

  // Hide reports list and show report viewer
  document.querySelector(".saved-reports-container").style.display = "none";
  document.getElementById("report-viewer").style.display = "block";

  // Render the report content
  renderReportContent(report);
}

// Function to check if a cell should be styled as a label in view-only mode
function isViewOnlyLabelCell(cellContent, cellIndex, row) {
  // List of label texts that should have grey background
  const labelTexts = [
    "TIER",
    "% SPLIT",
    "CSM/OBS",
    "LAST COMM. DATE",
    "START DATE",
    "END DATE",
    "LIVE DATE",
    "EXP. LIVE DATE",
    "PUBLISHER",
    "COMPETITOR",
    "STAGE",
    "STAKEHOLDERS",
    "TEST STATUS",
    "PROJECTED REVENUE",
    "OBS",
    "DEMAND STATUS",
    "STATUS",
    "CSM",
    "AVG. MONTHLY REV.",
    "ISSUE(S)/ UPDATE(S)",
    "NEXT STEPS",
    "IMPRESSIONS",
    "GROSS REVENUE",
    "GROSS CPM",
    "OBSERVATIONS",
  ];

  // Check if the cell content matches any label text
  const trimmedContent = cellContent.toString().trim().toUpperCase();
  if (labelTexts.includes(trimmedContent)) {
    return true;
  }

  // For first column cells, check if they should be treated as data rather than labels
  if (cellIndex === 0 && trimmedContent !== "") {
    // Check if this is a known label text first
    if (labelTexts.includes(trimmedContent)) {
      return true; // This is definitely a label
    }

    // Special handling for CUSTOMER SUCCESS tables where first column contains publisher names
    // Check if any cell in the row contains revenue data ($ or numbers) or "ended"/"reason"
    const hasRevenueData = row.some((cell) => {
      const cellStr = cell.toString();
      return cellStr.includes("$") || /\$?\d{1,3}(,\d{3})+/.test(cellStr);
    });

    const hasStatusData = row.some((cell) => {
      const cellStr = cell.toString().toLowerCase();
      return (
        cellStr.includes("ended") ||
        cellStr.includes("reason") ||
        cellStr.includes("contract")
      );
    });

    // If this row contains revenue or status data, the first column is publisher name (data), not a label
    if (hasRevenueData || hasStatusData) {
      return false; // This is data, not a label
    }

    return true; // Default: first column is a label
  }

  return false;
}

// Function to render report content in view-only mode
function renderReportContent(report) {
  const viewerContent = document.getElementById("viewed-report-content");
  let html = "";

  Object.keys(report.data).forEach((sectionKey) => {
    const section = report.data[sectionKey];
    html += `
      <section class="main-section">
        <div class="section-header">
          <h2>${section.title}</h2>
        </div>
        <div class="section-content">
    `;

    if (section.subsections) {
      Object.keys(section.subsections).forEach((subsectionKey) => {
        const subsection = section.subsections[subsectionKey];
        html += `
          <div class="subsection">
            <div class="subsection-header">
              <h3>${subsection.title}</h3>
            </div>
            <div class="tables-container">
        `;

        if (subsection.tables && subsection.tables.length > 0) {
          subsection.tables.forEach((table) => {
            html += `
              <div class="table-wrapper">
                <!-- Traditional table for desktop -->
                <table class="data-table">
                  <thead>
                    <tr>
                      ${table.headers
                        .map((header) => `<th>${header}</th>`)
                        .join("")}
                    </tr>
                  </thead>
                  <tbody>
                    ${table.rows
                      .map(
                        (row) => `
                      <tr>
                        ${row
                          .map((cell, cellIndex) => {
                            // Check if this cell should be a label based on content
                            const isLabelCell = isViewOnlyLabelCell(
                              cell,
                              cellIndex,
                              row
                            );
                            const className = isLabelCell
                              ? ' class="view-only-label"'
                              : "";
                            return `<td${className}>${cell}</td>`;
                          })
                          .join("")}
                      </tr>
                    `
                      )
                      .join("")}
                  </tbody>
                </table>

                <!-- Mobile card layout -->
                ${generateMobileTableCards(table)}
              </div>
            `;
          });
        }

        html += `
            </div>
          </div>
        `;
      });
    }

    html += `
        </div>
      </section>
    `;
  });

  viewerContent.innerHTML = html;
}

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
    const isLabelRow = headerCell && isViewOnlyLabelCell(headerCell, 0, row);

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
  const isFirstCellLabel = isViewOnlyLabelCell(row[0], 0, row);
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

// Function to go back to reports list
function backToReportsList() {
  document.querySelector(".saved-reports-container").style.display = "block";
  document.getElementById("report-viewer").style.display = "none";
}

// Function to delete a report
async function deleteReport(reportId) {
  const confirmed = await showCustomConfirm(
    "Are you sure you want to delete this report? This action cannot be undone.",
    "Delete Report"
  );
  if (!confirmed) {
    return;
  }

  const savedReports = JSON.parse(localStorage.getItem("savedReports") || "[]");
  const filteredReports = savedReports.filter((r) => r.id !== reportId);

  localStorage.setItem("savedReports", JSON.stringify(filteredReports));
  loadSavedReports(); // Refresh the list
}

// Auto-save every 30 seconds (silent)
setInterval(() => saveData(false), 30000);

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
