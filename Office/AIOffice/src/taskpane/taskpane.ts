/* global Excel console */

var selectionEventResult;

export async function insertText(insertRange: string, content: any[][]): Promise<void> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Parse and adjust insertRange
    let [sheetName, rangePart] = insertRange.split("!");
    let [startCell, endCell] = rangePart.split(":");
    const startRow = parseInt(startCell.match(/\d+/)[0], 10);
    const startCol = startCell.charCodeAt(0) - 65; // Assuming columns are A-Z

    // Calculate dimensions of content
    const contentRows = content.length;
    const contentCols = content[0]?.length || 0;

    // Automatically calculate endCell if missing or incorrect
    if (!endCell || endCell === startCell) {
      const endRow = startRow + contentRows - 1;
      const endCol = startCol + contentCols - 1;
      endCell = `${String.fromCharCode(65 + endCol)}${endRow}`;
      insertRange = `${sheetName}!${startCell}:${endCell}`;
    }

    const range = sheet.getRange(insertRange).load(["address", "rowCount", "columnCount"]);
    await context.sync();

    // Trim dimensions if they don't match
    const rowsToUse = Math.min(contentRows, range.rowCount);
    const colsToUse = Math.min(contentCols, range.columnCount);
    const trimmedContent = content
      .slice(contentRows - rowsToUse)
      .map((row) => row.slice(contentCols - colsToUse));

    // Insert trimmed content
    range.values = trimmedContent;
    range.format.autofitColumns();
    await context.sync();
  });
}

export const registerSelectionChangeHandler = async (setSelectedRange) => {
  const onSelectionChanged = () =>
    Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "address"]);
      await context.sync();
      setSelectedRange(range);
    });
  try {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      await onSelectionChanged();
      selectionEventResult = worksheet.onSelectionChanged.add(
        async (_e) => await onSelectionChanged()
      );
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