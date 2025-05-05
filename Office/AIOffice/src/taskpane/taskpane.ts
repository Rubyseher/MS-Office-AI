/* global Excel console */

var selectionEventResult;

export async function insertText(text: string) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export const registerSelectionChangeHandler = async (setSelectedRange) => {
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      selectionEventResult = worksheet.onSelectionChanged.add(async (_e) => {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load(["values", "address"]);
          await context.sync();
          setSelectedRange(range);
        });
      });
      await context.sync();
      console.log("Event handler successfully registered for onSelectionChanged event.");
    });
  } catch (error) {
    console.error("Error registering selection change handler:", error);
  }
};

export const removeSelectionChangeHandler = async () => {
  try {
    if (selectionEventResult) {
      await Excel.run(selectionEventResult.context, async (context) => {
        selectionEventResult.remove();
        await context.sync();
        console.log("Event handler successfully removed.");
      });
    }
  } catch (error) {
    console.error("Error removing selection change handler:", error);
  }
};