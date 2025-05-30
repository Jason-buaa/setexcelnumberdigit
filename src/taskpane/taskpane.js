/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    // Add real-time validation for the range input
    const rangeInput = document.getElementById("range-input");
    const validationMessage = document.getElementById("range-validation");

    rangeInput.addEventListener("input", () => {
      const input = rangeInput.value;

      // Regular expression to validate Excel range (e.g., A1, A1:B10, etc.)
      const rangeRegex = /^[A-Z]+[1-9][0-9]*(:[A-Z]+[1-9][0-9]*)?$/;

      if (input === "" || rangeRegex.test(input)) {
        validationMessage.textContent = "Valid range";
        validationMessage.style.color = "green";
      } else {
        validationMessage.textContent = "Invalid range";
        validationMessage.style.color = "red";
      }
    });

    // Update the displayed decimal places value when the slider is moved
    const slider = document.getElementById("decimal-slider");
    const decimalValue = document.getElementById("decimal-value");

    slider.addEventListener("input", () => {
      decimalValue.textContent = slider.value;
    });
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeInput = document.getElementById("range-input").value.trim();
      const decimalPlaces = document.getElementById("decimal-slider").value; // Get decimal places from slider
      let range;

      if (rangeInput === "") {
        // Use the entire used range if no input is provided
        range = sheet.getUsedRange();
      } else {
        // Use the range specified in the input
        range = sheet.getRange(rangeInput);
      }

      // Load the values and number formats of the range
      range.load(["values", "numberFormat"]);
      await context.sync();

      const values = range.values;
      const numericFormat = `0.${"0".repeat(decimalPlaces)}`; // Define the number format dynamically

      // Create a new matrix for number formats
      const updatedNumberFormat = values.map(row =>
        row.map(cell => (typeof cell === "number" ? numericFormat : null))
      );

      // Apply the updated number formats to the specified range
      range.numberFormat = updatedNumberFormat;

      await context.sync();
      console.log(`Updated numeric cells to ${decimalPlaces} decimal places in the specified range.`);
    });
  } catch (error) {
    console.error(error);
  }
}
