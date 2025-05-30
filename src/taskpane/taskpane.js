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
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();

      // Load the values and number formats of the used range
      usedRange.load(["values", "numberFormat"]);

      await context.sync();

      // Iterate through the cells and update number format for numeric cells
      const values = usedRange.values;
      const numberFormat = usedRange.numberFormat;

      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          if (typeof values[i][j] === "number") {
            // Set the number format to display 2 decimal places
            numberFormat[i][j] = "0.00";
          }
        }
      }

      // Apply the updated number formats
      usedRange.numberFormat = numberFormat;

      await context.sync();
      console.log("Updated all numeric cells to 2 decimal places.");
    });
  } catch (error) {
    console.error(error);
  }
}
