function debugRubyGeneration() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Leads");
  const row = sheet.getActiveCell().getRow();
  
  // Don't process header row
  if (row === 1) {
    console.log("Error: Please select a data row, not the header");
    SpreadsheetApp.getUi().alert("Please select a data row (not the header row)");
    return;
  }

  console.log("Debug Info:");
  console.log("Current row:", row);
  console.log("Column AA (TYPE):", sheet.getRange(row, 27).getDisplayValue());
  console.log("Column T (Length):", sheet.getRange(row, 20).getValue());
  console.log("Column U (Width):", sheet.getRange(row, 21).getValue());
  console.log("Column V (Front Bar):", sheet.getRange(row, 22).getValue());
  console.log("Column X (Wing Height):", sheet.getRange(row, 24).getValue());
  console.log("Column Y (Num Wings):", sheet.getRange(row, 25).getValue());
  console.log("Column AB (Fabric):", sheet.getRange(row, 28).getDisplayValue());

  // Clear S to see what the handler writes
  sheet.getRange(row, 19).clearContent();

  // Simulate an edit event
  const e = {
    source: SpreadsheetApp.getActive(),
    range: sheet.getRange(row, 27)
  };

handleEditAwningRuby_(e); 

  const result = sheet.getRange(row, 19).getValue();
  console.log("Column S after handler:", result ? "Has content" : "Empty");
  
  if (!result) {
    console.log("No output generated. Check that column AA contains 'Lean-to' or 'Sloped L'");
  }
}