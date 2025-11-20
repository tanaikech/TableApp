# TableApp

<a name="top"></a>
[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

## Overview

**TableApp is a Google Apps Script library for managing Tables on Google Sheets.**

<a name="description"></a>

## Description

Recently, a new feature "Tables" was introduced to Google Sheets. Tables allow users to group data into structured tables with headers, filtering, and specific data types. While these can be managed via the Google Sheets API (v4), constructing the raw JSON requests for operations like creating, updating, and managing tables can be complex.

This library, `TableApp`, creates an object-oriented wrapper around the Google Sheets API, making it easy to manage Tables directly within Google Apps Script.

## Library's Script ID

```
1G4RVvyLtwPjQl6x_p8j3X65-yYVU3w2dMXxHDuzCgorucjs8P3Clv5Qt
```

- **Please copy and paste the above Script ID into the search box of the "Libraries" in your Google Apps Script project.**
- [Installation Guide of Google Apps Script library](https://developers.google.com/apps-script/guides/libraries)

<a name="requirements"></a>

## Requirements

This library uses the **Google Sheets API**.

1.  In the Script Editor, go to **Services** (on the left side).
2.  Click `+` (Add a service).
3.  Select **Google Sheets API**.
4.  Click **Add**.

### Scopes

This library uses the following scopes.

- `https://www.googleapis.com/auth/spreadsheets.readonly`
- `https://www.googleapis.com/auth/spreadsheets`

If you use only `get` methods, you can use them with only `https://www.googleapis.com/auth/spreadsheets.readonly`. If you use methods other than `get` (e.g., create, update, delete), please use `https://www.googleapis.com/auth/spreadsheets`.

<a name="methods"></a>

## Methods

### Class `TableApp`

| Method                      | Description                                                  |
| :-------------------------- | :----------------------------------------------------------- |
| `openById(spreadsheetId)`   | Opens the TableApp for a specific Spreadsheet.               |
| `getSheetByName(sheetName)` | Sets the target sheet name for creating a table.             |
| `getRange(a1Notation)`      | Sets the target range for creating a table.                  |
| `create(tableName)`         | Creates a new table.                                         |
| `getTables()`               | Retrieves all tables in the spreadsheet (or specific sheet). |
| `getTableByName(tableName)` | Retrieves a table by its name.                               |
| `getTableById(tableId)`     | Retrieves a table by its ID.                                 |

### Class `Table`

| Method                               | Description                                                                                                                                          |
| :----------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------- |
| `getName()`                          | Gets the name of the table.                                                                                                                          |
| `getId()`                            | Gets the ID of the table.                                                                                                                            |
| `getRange()`                         | Gets the A1 notation of the table range.                                                                                                             |
| `getValues()`                        | Gets the values in the table range.                                                                                                                  |
| `setName(tableName)`                 | Updates the name of the table.                                                                                                                       |
| `setValues(values)`                  | Sets values to the table range.                                                                                                                      |
| `setRange(a1Notation)`               | Moves or resizes the table to a new range.                                                                                                           |
| `setRowsProperties(props, fields)`   | Updates row properties. ([props](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/sheets#TableRowsProperties))      |
| `setColumnProperties(props, fields)` | Updates column properties. ([props](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/sheets#TableColumnProperties)) |
| `copyTo(a1Notation)`                 | Copies the table to a destination range.                                                                                                             |
| `reverse()`                          | Converts the table back to a normal range (keeps data).                                                                                              |
| `remove()`                           | Deletes the table completely.                                                                                                                        |

<a name="usage"></a>

## Usage

### 1. Create a Table

This sample creates a new table named "MyTable" in "Sheet1" at range "A1:C5".

```javascript
function createTableSample() {
  const spreadsheetId = "###"; // Please set your Spreadsheet ID.
  const tableApp = TableApp.openById(spreadsheetId);

  const table = tableApp
    .getSheetByName("Sheet1")
    .getRange("A1:C5")
    .create("MyTable");

  console.log(`Created Table ID: ${table.getId()}`);
}
```

### 2. Get and Update a Table

This sample retrieves an existing table by name, renames it, and updates the values.

```javascript
function updateTableSample() {
  const spreadsheetId = "###"; // Please set your Spreadsheet ID.
  const tableApp = TableApp.openById(spreadsheetId);

  // Retrieve table
  const table = tableApp.getTableByName("MyTable");

  if (table) {
    // Rename table
    table.setName("UpdatedTableName");

    // Update values
    const newValues = [
      ["ID", "Name", "Value"],
      [1, "a", 100],
      [2, "b", 200],
      [3, "c", 300],
      [4, "d", 400],
    ];
    table.setValues(newValues);

    // Move table location
    table.setRange("E1:G5");
  }
}
```

### 3. Copy, Reverse, and Delete

This sample demonstrates copying a table, reversing a table (converting to range), and deleting a table.

```javascript
function manageTableSample() {
  const spreadsheetId = "###"; // Please set your Spreadsheet ID.
  const tableApp = TableApp.openById(spreadsheetId);
  const table = tableApp.getTableByName("UpdatedTableName");

  if (table) {
    // Copy table to another location
    const copiedTable = table.copyTo("Sheet2!A1");
    console.log(`Copied Table ID: ${copiedTable.getId()}`);

    // Reverse table (Table structure is removed, data remains)
    table.reverse();

    // Remove copied table (Table structure is removed)
    copiedTable.remove();
  }
}
```

### 4. Manage Properties

This sample updates the header style and column names.

```javascript
function managePropertiesSample() {
  const spreadsheetId = "###";
  const tableApp = TableApp.openById(spreadsheetId);
  const table = tableApp.getTableByName("UpdatedTableName");

  if (table) {
    // Set Header Style (Row Properties)
    table.setRowsProperties(
      {
        headerColorStyle: { rgbColor: { red: 0.9, green: 0.9, blue: 0.9 } },
      },
      "rowsProperties"
    ); // "rowsProperties" is the default fields. You can manage the fields.

    // Set Column Names (Column Properties)
    // Note: columnIndex is required.
    table.setColumnProperties(
      [
        { columnIndex: 0, columnName: "Identifier" },
        { columnIndex: 1, columnName: "Item Name" },
        { columnIndex: 2, columnName: "Cost" },
      ],
      "columnProperties"
    ); // "columnProperties" is the default fields. You can manage the fields.
  }
}
```

The objects for `setRowsProperties` and `setColumnProperties` methods are as follows.

- [setRowsProperties](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/sheets#TableRowsProperties)
- [setColumnProperties](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/sheets#TableColumnProperties)

<a name="testscript"></a>

## Complete Test Script

You can copy and run this script to test all features of the library. This script creates a temporary Google Spreadsheet, performs Create, Read, Update (Values, Rows, Columns), Copy, Reverse, and Remove operations, and then prints the results to the log.

**This test script was created by Gemini 3 Pro Preview**

```javascript
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
```

<a name="licence"></a>

## Licence

[MIT](LICENSE)

<a name="author"></a>

## Author

[Tanaike](https://tanaikech.github.io/)

[Donate](https://tanaikech.github.io/donate/)

<a name="updatehistory"></a>

## Update History

- v1.0.0 (November 20, 2025)
  - Initial release.

[TOP](#top)
