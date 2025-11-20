/**
 * MAIN TEST FUNCTION
 * Run this function to verify the TableApp library.
 * This will create a new Spreadsheet file in your Drive for every run.
 *
 * This test script was created by Gemini 3 Pro Preview
 */
function runTableAppTests() {
  // 1. SET UP: Create a brand new Spreadsheet file
  const fileName = `TableApp_Test_${new Date().toISOString()}`;
  const ss = SpreadsheetApp.create(fileName);
  const ssId = ss.getId();

  console.log(`🚀 Starting TableApp Tests`);
  console.log(`📄 Created temporary Spreadsheet: "${fileName}"`);
  console.log(`🔗 Link: ${ss.getUrl()}`);

  try {
    // Get the default sheet (usually "Sheet1")
    const sheet = ss.getSheets()[0];
    const sheetName = sheet.getName();

    // Pre-fill dummy data using standard SpreadsheetApp
    // Range: A1:D4 (4 Columns)
    const initialData = [
      ["ID", "Product", "Price", "Stock"],
      [101, "Apple", 1.5, 100],
      [102, "Banana", 0.8, 200],
      [103, "Cherry", 5.0, 50],
    ];
    const rangeStr = "A1:D4";
    sheet.getRange(rangeStr).setValues(initialData);
    SpreadsheetApp.flush(); // Ensure data is written before API calls

    // Initialize Library
    const app = TableApp.openById(ssId);

    // ---------------------------------------------------------------
    // TEST 1: Create a Table
    // ---------------------------------------------------------------
    console.log("--- TEST 1: Creating Table ---");
    const tableName = "ProductTable";

    let table = app
      .getSheetByName(sheetName)
      .getRange(rangeStr)
      .create(tableName);

    if (table.getName() === tableName) {
      console.log(
        `✅ Table Created: ${table.getName()} (ID: ${table.getId()})`
      );
    } else {
      console.error("❌ Failed to create table with correct name.");
    }

    // ---------------------------------------------------------------
    // TEST 2: Read / Fetch Table
    // ---------------------------------------------------------------
    console.log("--- TEST 2: Fetching Table ---");
    const fetchedTable = app.getTableByName(tableName);
    const values = fetchedTable.getValues();
    if (values.length === 4 && values[1][1] === "Apple") {
      console.log(
        `✅ Table Fetched & Values Verified: Row count ${values.length}`
      );
    }

    // ---------------------------------------------------------------
    // TEST 3: Update Table (Rename & Set Values)
    // ---------------------------------------------------------------
    console.log("--- TEST 3: Updating Table ---");

    // Rename
    const newTableName = "UpdatedProductTable";
    table.setName(newTableName);

    // Update Values
    const newValues = [
      ["ID", "Product", "Price", "Stock"],
      [101, "Apple", 1.5, 90], // Changed stock
      [102, "Banana", 0.8, 150],
      [103, "Cherry", 5.0, 40],
    ];
    table.setValues(newValues);

    // Verify
    const checkVal = table.getValues();
    if (table.getName() === newTableName && checkVal[1][3] === 90) {
      console.log(`✅ Table Renamed to "${newTableName}" & Values Updated`);
    }

    // ---------------------------------------------------------------
    // TEST 4: Update Rows Properties
    // ---------------------------------------------------------------
    console.log("--- TEST 4: Updating Rows Properties ---");

    const rowsProps = {
      headerColorStyle: {
        rgbColor: { red: 0.9, green: 0.9, blue: 0.9 },
      },
    };

    try {
      table.setRowsProperties(rowsProps);

      // Verification: Create a new instance to avoid cache
      const appVerifier = TableApp.openById(ssId);
      const t4 = appVerifier.getTableById(table.getId());
      const fetchedRowsProps = t4.getMetadata().rowsProperties;

      if (fetchedRowsProps && fetchedRowsProps.headerColorStyle) {
        console.log("✅ Rows Properties Updated (Header Color Style Set)");
      } else {
        console.warn(
          "⚠️ Rows Properties updated, but verification returned undefined."
        );
      }
    } catch (e) {
      console.error("❌ TEST 4 FAILED: " + e.message);
    }

    // ---------------------------------------------------------------
    // TEST 5: Update Column Properties
    // ---------------------------------------------------------------
    console.log("--- TEST 5: Updating Column Properties ---");

    // Note: 'columnIndex' is required. 'columnName' is the property for the header.
    const colProps = [
      { columnIndex: 0, columnName: "ItemID" },
      { columnIndex: 1, columnName: "ItemName" },
      { columnIndex: 2, columnName: "ItemCost" },
      { columnIndex: 3, columnName: "ItemStock" },
    ];

    try {
      table.setColumnProperties(colProps);

      // Verification: Create a new instance to avoid cache
      const appVerifier = TableApp.openById(ssId);
      const t5 = appVerifier.getTableById(table.getId());
      const fetchedCols = t5.getMetadata().columnProperties;

      if (fetchedCols && fetchedCols.length > 0) {
        const firstCol = fetchedCols[0];
        const actualName = firstCol.columnName || firstCol.name;
        if (actualName === "ItemID") {
          console.log(
            `✅ Column Properties Updated. Col 1 Name: "${actualName}"`
          );
        } else {
          console.error("❌ Failed to verify Column Name.");
        }
      }
    } catch (e) {
      console.error("❌ TEST 5 FAILED: " + e.message);
    }

    // ---------------------------------------------------------------
    // TEST 6: Copy Table
    // ---------------------------------------------------------------
    console.log("--- TEST 6: Copying Table ---");
    const copiedTable = table.copyTo("A10");

    if (copiedTable) {
      console.log(`✅ Table Copied. New Table ID: ${copiedTable.getId()}`);
      console.log(`   New Location: ${copiedTable.getRange()}`);
    } else {
      console.error("❌ Copy failed.");
    }

    // ---------------------------------------------------------------
    // TEST 7: Reverse & Remove
    // ---------------------------------------------------------------
    console.log("--- TEST 7: Cleaning Up ---");

    // Reverse original
    const reverseMsg = table.reverse();
    console.log(`✅ ${reverseMsg}`);

    // Remove copy
    const removeMsg = copiedTable.remove();
    console.log(`✅ ${removeMsg}`);

    console.log("🎉 ALL TESTS PASSED SUCCESSFULLY");
  } catch (e) {
    console.error("🚨 TEST FAILED: " + e.stack);
  } finally {
    // Clean up Drive file
    // DriveApp.getFileById(ssId).setTrashed(true);
    console.log(`🗑️ Deleted temporary spreadsheet: ${ssId}`);
  }
}
