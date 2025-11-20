/**
 * TableApp: TableApp is a Google Apps Script library for managing Tables on Google Sheets.
 * (required) Sheets API
 * Author: Kanshi Tanaike
 * https://github.com/tanaikech/TableApp
 *
 * Updated on 20251120 11:45
 * version 1.0.0
 */

/**
 * Opens the TableApp for a specific Spreadsheet.
 *
 * @param {string} spreadsheetId The Spreadsheet ID.
 * @return {TableApp} The TableApp instance.
 */
function openById(spreadsheetId) {
  return new TableApp(spreadsheetId);
}

/**
 * Class for managing Tables within a Google Spreadsheet.
 * Acts as the main entry point for creating and retrieving tables.
 */
class TableApp {
  /**
   * @param {string} spreadsheetId The Spreadsheet ID.
   */
  constructor(spreadsheetId) {
    /** @private @type {string} */
    this.spreadsheetId = spreadsheetId;
    /** @private @type {string|null} */
    this.sheetName = null;
    /** @private @type {string|null} */
    this.a1Notation = null;
    /** @private @type {Object|null} */
    this.cachedTables = null;
  }

  /**
   * Sets the target sheet by name for subsequent operations (like create).
   *
   * @param {string} sheetName The name of the sheet.
   * @return {TableApp} This instance for chaining.
   * @throws {Error} If spreadsheet ID is missing.
   */
  getSheetByName(sheetName) {
    if (!this.spreadsheetId) throw new Error("No Spreadsheet ID defined.");
    this.sheetName = sheetName;
    return this;
  }

  /**
   * Sets the target range using A1 notation for creating a table.
   *
   * @param {string} a1Notation The range in A1 notation (e.g., "Sheet1!A1:B5" or "A1:B5").
   * @return {TableApp} This instance for chaining.
   * @throws {Error} If spreadsheet ID is missing.
   */
  getRange(a1Notation) {
    if (!this.spreadsheetId) throw new Error("No Spreadsheet ID defined.");

    const parsed = parseA1Notation_(a1Notation);
    // If the A1 notation includes a sheet name, override the stored sheetName
    if (parsed && parsed.sheetName) {
      this.sheetName = parsed.sheetName;
    }

    this.a1Notation = a1Notation;
    return this;
  }

  /**
   * Creates a new table in the defined range/sheet.
   *
   * @param {string} tableName The name of the new table.
   * @return {Table} The created Table instance.
   * @throws {Error} If the range resolution fails.
   */
  create(tableName) {
    const { gridRange, sheetName, sheetId } = this._resolveGridRange();

    const requests = [
      {
        addTable: {
          table: {
            name: tableName,
            range: gridRange,
          },
        },
      },
    ];

    const response = batchUpdate_(this.spreadsheetId, requests);
    // @ts-ignore
    const newTableObj = response.replies[0].addTable.table;

    // Invalidate cache since a new table exists
    this.cachedTables = null;

    return new Table({
      spreadsheetId: this.spreadsheetId,
      sheetName: sheetName,
      sheetId: sheetId,
      table: newTableObj,
    });
  }

  /**
   * Retrieves all tables.
   * If a sheet is selected via getSheetByName, returns tables for that sheet.
   * Otherwise returns all tables organized by sheet name.
   *
   * @return {Object|Array<Table>} An object of tables {SheetName: {sheetId, tables: []}} or array of Table instances if sheet is filtered.
   */
  getTables() {
    const data = this._getOrFetchTables();
    return this.sheetName
      ? data.tablesBySheetNames[this.sheetName]?.tables || []
      : data.tablesBySheetNames;
  }

  /**
   * Gets a specific table by its name.
   *
   * @param {string} tableName The name of the table.
   * @return {Table|null} The Table instance or null if not found.
   */
  getTableByName(tableName) {
    const data = this._getOrFetchTables();
    return data.tablesByTableNames[tableName] || null;
  }

  /**
   * Gets a specific table by its ID.
   *
   * @param {string} tableId The ID of the table.
   * @return {Table|null} The Table instance or null if not found.
   */
  getTableById(tableId) {
    const data = this._getOrFetchTables();
    return data.tablesByTableIds[tableId] || null;
  }

  /**
   * Internal method to fetch tables with caching.
   *
   * @private
   * @return {Object} The structure returned by fetchAllTables_
   */
  _getOrFetchTables() {
    if (!this.cachedTables) {
      this.cachedTables = fetchAllTables_(this.spreadsheetId);
    }
    return this.cachedTables;
  }

  /**
   * Internal method to resolve A1 notation to GridRange and Sheet properties.
   *
   * @private
   * @return {{gridRange: Object, sheetName: string, sheetId: number}}
   */
  _resolveGridRange() {
    const notation = this.a1Notation || "A1";
    const parsed = parseA1Notation_(notation);

    // Determine target sheet name
    let targetSheetName =
      parsed && parsed.sheetName ? parsed.sheetName : this.sheetName;

    // Fetch sheet metadata
    const sheetsData = sget_(
      this.spreadsheetId,
      "sheets(properties(sheetId,title))"
    );
    let targetSheetId = null;

    if (!targetSheetName) {
      // Default to first sheet if no specific sheet is defined
      // @ts-ignore
      targetSheetName = sheetsData.sheets[0].properties.title;
      // @ts-ignore
      targetSheetId = sheetsData.sheets[0].properties.sheetId;
    } else {
      // @ts-ignore
      const found = sheetsData.sheets.find(
        (s) => s.properties.title === targetSheetName
      );
      if (!found)
        throw new Error(`Sheet with name "${targetSheetName}" not found.`);
      targetSheetId = found.properties.sheetId;
    }

    // Convert to GridRange
    const gridRange = convA1NotationToGridRange_(notation, targetSheetId);

    return { gridRange, sheetName: targetSheetName, sheetId: targetSheetId };
  }
}

/**
 * Class representing a single Table in Google Sheets.
 */
class Table {
  /**
   * @param {Object} obj Configuration object.
   * @param {string} obj.spreadsheetId
   * @param {string} obj.sheetName
   * @param {number} obj.sheetId
   * @param {Object} obj.table The table object from Sheets API.
   */
  constructor(obj) {
    /** @private */
    this.spreadsheetId = obj.spreadsheetId;
    /** @private */
    this.sheetName = obj.sheetName;
    /** @private */
    this.sheetId = obj.sheetId;
    /** @private */
    this.table = obj.table;
    /** @private */
    this.rangeAsA1Notation = convGridRangeToA1Notation_(
      this.table.range,
      this.sheetName
    );
  }

  /**
   * Gets the name of the table.
   *
   * @return {string} The table name.
   */
  getName() {
    return this.table.name;
  }

  /**
   * Gets the Table ID.
   *
   * @return {string} The table ID.
   */
  getId() {
    return this.table.tableId;
  }

  /**
   * Gets the raw metadata object of the table.
   *
   * @return {Object} The raw table object.
   */
  getMetadata() {
    return this.table;
  }

  /**
   * Gets the table range in A1 notation.
   *
   * @return {string} The A1 notation of the range (e.g., "'Sheet1'!A1:C10").
   */
  getRange() {
    // Refresh range notation in case metadata was updated externally
    return convGridRangeToA1Notation_(this.table.range, this.sheetName);
  }

  /**
   * Gets the values in the table range.
   *
   * @return {Array<Array<any>>} 2D array of values.
   */
  getValues() {
    return valuesGet_(this.spreadsheetId, this.rangeAsA1Notation);
  }

  /**
   * Updates the table name.
   *
   * @param {string} tableName New name.
   * @return {Table} This instance.
   * @throws {Error} If tableName is invalid.
   */
  setName(tableName) {
    if (!tableName || typeof tableName !== "string")
      throw new Error("Invalid table name.");
    const requests = [
      {
        updateTable: {
          fields: "name",
          table: { name: tableName, tableId: this.table.tableId },
        },
      },
    ];
    this._updateTable(requests);
    this.table.name = tableName; // Update local state
    return this;
  }

  /**
   * Sets values into the table range.
   *
   * @param {Array<Array<any>>} values 2D array of values.
   * @return {Object} The updated range object from the API response.
   * @throws {Error} If values are not a 2D array.
   */
  setValues(values) {
    if (
      !Array.isArray(values) ||
      values.length === 0 ||
      !Array.isArray(values[0])
    ) {
      throw new Error("Invalid values. Must be a 2D array.");
    }
    return valuesUpdate_(this.spreadsheetId, values, this.rangeAsA1Notation);
  }

  /**
   * Moves or resizes the table to a new range.
   *
   * @param {string} a1Notation New range in A1 notation.
   * @return {Table} This instance.
   * @throws {Error} If A1 notation is invalid.
   */
  setRange(a1Notation) {
    if (!a1Notation || typeof a1Notation !== "string")
      throw new Error("Invalid A1 Notation.");
    const gridRange = convA1NotationToGridRange_(a1Notation, this.sheetId);

    const requests = [
      {
        updateTable: {
          fields: "range",
          table: { range: gridRange, tableId: this.table.tableId },
        },
      },
    ];

    this._updateTable(requests);

    // Update local state
    this.table.range = gridRange;
    this.rangeAsA1Notation = convGridRangeToA1Notation_(
      gridRange,
      this.sheetName
    );
    return this;
  }

  /**
   * Updates row properties (e.g. headers).
   *
   * @param {Object} rowsProperties API object for row properties.
   * @param {string} [fields="rowsProperties"] The fields to update.
   * @return {Table} This instance.
   * @throws {Error} If rowsProperties is invalid.
   */
  setRowsProperties(rowsProperties, fields = "rowsProperties") {
    if (!rowsProperties || typeof rowsProperties !== "object")
      throw new Error("Invalid object.");
    const requests = [
      {
        updateTable: {
          fields,
          table: { rowsProperties, tableId: this.table.tableId },
        },
      },
    ];
    return this._updateTable(requests);
  }

  /**
   * Updates column properties.
   *
   * @param {Object} columnProperties API object for column properties.
   * @param {string} [fields="columnProperties"] The fields to update.
   * @return {Table} This instance.
   * @throws {Error} If columnProperties is invalid.
   */
  setColumnProperties(columnProperties, fields = "columnProperties") {
    if (!columnProperties || typeof columnProperties !== "object")
      throw new Error("Invalid object.");
    const requests = [
      {
        updateTable: {
          fields,
          table: { columnProperties, tableId: this.table.tableId },
        },
      },
    ];
    return this._updateTable(requests);
  }

  /**
   * Deletes the table structure from the sheet.
   * The cell data remains, but it is no longer a "Table" entity.
   *
   * @return {string} Status message.
   */
  remove() {
    const requests = [{ deleteTable: { tableId: this.table.tableId } }];
    try {
      batchUpdate_(this.spreadsheetId, requests);
    } catch (e) {
      throw new Error(`Failed to remove table: ${e.message}`);
    }
    return `${this.table.name} (Table ID: ${this.table.tableId}) was successfully deleted.`;
  }

  /**
   * Reverses the table (converts it back to normal cells) and explicitly restores cell data.
   * This ensures that any table-specific metadata is removed while "baking" the
   * values and formats into the cells.
   *
   * @return {string} Status message.
   * @throws {Error} If the sheet data cannot be retrieved.
   */
  reverse() {
    // Fields explicitly requested to not be changed
    const fieldsToFetch =
      "sheets(properties(sheetId),data(rowData(values(userEnteredValue,textFormatRuns,chipRuns,userEnteredFormat,effectiveValue,hyperlink,note,dataValidation))))";

    const obj = sget_(this.spreadsheetId, fieldsToFetch, [
      this.rangeAsA1Notation,
    ]);
    // @ts-ignore
    const f = obj.sheets.find(
      ({ properties: { sheetId } }) => sheetId == this.sheetId
    );

    if (f) {
      this.remove();
      const rows = f.data[0].rowData;
      // Fields explicitly requested to not be changed
      const requests = [
        { updateCells: { rows, range: this.table.range, fields: "*" } },
      ];
      batchUpdate_(this.spreadsheetId, requests);
    } else {
      throw new Error("Sheet not found during reverse operation.");
    }
    return `${this.table.name} (Table ID: ${this.table.tableId}) was successfully reversed.`;
  }

  /**
   * Copies the table to a new destination.
   *
   * @param {string} a1Notation Destination range in A1 notation.
   * @return {Table} A new Table instance representing the copy.
   * @throws {Error} If invalid notation or destination sheet not found.
   */
  copyTo(a1Notation) {
    if (!a1Notation || typeof a1Notation !== "string")
      throw new Error("Invalid A1 Notation.");

    const parsed = parseA1Notation_(a1Notation);

    // Resolve Destination Sheet ID
    const sheetsData = sget_(
      this.spreadsheetId,
      "sheets(properties(sheetId,title))"
    );
    let destSheetId, destSheetName;

    if (parsed.sheetName) {
      // @ts-ignore
      const s = sheetsData.sheets.find(
        ({ properties: { title } }) => title === parsed.sheetName
      );
      if (s) {
        destSheetId = s.properties.sheetId;
        destSheetName = s.properties.title;
      } else {
        throw new Error(`Destination sheet "${parsed.sheetName}" not found.`);
      }
    } else {
      // If no sheet provided, assume same sheet as current table
      destSheetId = this.sheetId;
      destSheetName = this.sheetName;
    }

    const destination = convA1NotationToGridRange_(parsed.range, destSheetId);
    const requests = [{ copyPaste: { source: this.table.range, destination } }];

    batchUpdate_(this.spreadsheetId, requests);

    // Fetch the new table to return an instance of it
    const tablesData = fetchAllTables_(this.spreadsheetId);

    // Find the table that matches the destination range
    const copiedTableObj = tablesData.tablesBySheetNames[
      destSheetName
    ]?.tables?.find((t) => {
      const tr = t.getMetadata().range;
      return (
        (tr.sheetId || 0) === (destSheetId || 0) &&
        tr.startRowIndex === destination.startRowIndex &&
        tr.startColumnIndex === destination.startColumnIndex
      );
    });

    if (!copiedTableObj) {
      throw new Error(
        "Table copied, but could not retrieve the new table instance."
      );
    }

    return copiedTableObj;
  }

  /**
   * Helper to execute update requests.
   *
   * @private
   * @param {Array<Object>} requests The batch update requests.
   * @return {Table} This instance.
   */
  _updateTable(requests) {
    try {
      batchUpdate_(this.spreadsheetId, requests);
    } catch (e) {
      throw new Error(`Table update failed: ${e.message}`);
    }
    return this;
  }
}

/* -------------------------------------------------------------------------- */
/*                               PRIVATE HELPERS                              */
/* -------------------------------------------------------------------------- */

/**
 * Wrapper for Sheets.Spreadsheets.get
 *
 * @private
 * @param {string} spreadsheetId
 * @param {string} [fields="*"]
 * @param {Array<string>} [ranges=[]]
 * @return {Object} API response
 */
function sget_(spreadsheetId, fields = "*", ranges = []) {
  return Sheets.Spreadsheets.get(spreadsheetId, { fields, ranges });
}

/**
 * Wrapper for Sheets.Spreadsheets.batchUpdate
 *
 * @private
 * @param {string} spreadsheetId
 * @param {Array<Object>} requests
 * @return {Object} API response
 */
function batchUpdate_(spreadsheetId, requests) {
  return Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
}

/**
 * Wrapper for Sheets.Spreadsheets.Values.get
 *
 * @private
 * @param {string} spreadsheetId
 * @param {string} range
 * @return {Array<Array<any>>} Values
 */
function valuesGet_(spreadsheetId, range) {
  const res = Sheets.Spreadsheets.Values.get(spreadsheetId, range, {
    valueRenderOption: "FORMATTED_VALUE",
  });
  return res.values;
}

/**
 * Wrapper for Sheets.Spreadsheets.Values.update
 *
 * @private
 * @param {string} spreadsheetId
 * @param {Array<Array<any>>} values
 * @param {string} range
 * @return {Object} Updated range info
 */
function valuesUpdate_(spreadsheetId, values, range) {
  const res = Sheets.Spreadsheets.Values.update(
    { values },
    spreadsheetId,
    range,
    { valueInputOption: "USER_ENTERED" }
  );
  return res.updatedRange;
}

/**
 * Fetch and categorize all tables in the spreadsheet.
 *
 * @private
 * @param {string} spreadsheetId
 * @return {{tablesBySheetNames: Object, tablesByTableNames: Object, tablesByTableIds: Object}} Structured table data
 */
function fetchAllTables_(spreadsheetId) {
  const res = sget_(spreadsheetId, "sheets(properties(sheetId,title),tables)");

  const result = {
    tablesBySheetNames: {},
    tablesByTableNames: {},
    tablesByTableIds: {},
  };

  // @ts-ignore
  if (!res.sheets) return result;

  // @ts-ignore
  res.sheets.forEach((sheet) => {
    const title = sheet.properties.title;
    const sheetId = sheet.properties.sheetId;
    const apiTables = sheet.tables;

    if (apiTables && apiTables.length > 0) {
      // Create Table instances
      const tableInstances = apiTables.map(
        (t) =>
          new Table({
            spreadsheetId,
            sheetName: title,
            sheetId: sheetId,
            table: t,
          })
      );

      // Organize by Sheet Name
      result.tablesBySheetNames[title] = {
        sheetId,
        tables: tableInstances,
      };

      // Organize by Table Name and ID
      tableInstances.forEach((inst) => {
        const name = inst.getName();
        if (name) result.tablesByTableNames[name] = inst;
        result.tablesByTableIds[inst.getId()] = inst;
      });
    }
  });

  return result;
}

/**
 * Parse A1Notation to separate sheet name and range.
 *
 * @private
 * @param {string} a1Notation
 * @return {{sheetName: string|null, range: string}|null}
 */
function parseA1Notation_(a1Notation) {
  if (!a1Notation) return null;
  const regex = /(?:(?:'([^']*)'|([^!]+))!)?(.+)/;
  const match = a1Notation.match(regex);
  if (match) {
    const sheetName = match[1] || match[2] || null;
    const range = match[3];
    return { sheetName, range };
  }
  return null;
}

/**
 * Converts a column letter to an index (0-based).
 *
 * @private
 * @param {string} letter
 * @return {number}
 */
function columnLetterToIndex_(letter) {
  if (!letter || typeof letter !== "string")
    throw new Error("Column letter must be a string.");
  letter = letter.toUpperCase();
  return [...letter].reduce(
    (c, e, i, a) =>
      (c += (e.charCodeAt(0) - 64) * Math.pow(26, a.length - i - 1)),
    -1
  );
}

/**
 * Converts a column index (0-based) to a letter.
 *
 * @private
 * @param {number} index
 * @return {string}
 */
function columnIndexToLetter_(index) {
  if (index === null || isNaN(index))
    throw new Error("Index must be a number (0-based).");
  let a;
  return (a = Math.floor(index / 26)) >= 0
    ? columnIndexToLetter_(a - 1) + String.fromCharCode(65 + (index % 26))
    : "";
}

/**
 * Converts A1Notation to GridRange.
 *
 * @private
 * @param {string} a1Notation
 * @param {number} sheetId
 * @return {Object} GridRange
 */
function convA1NotationToGridRange_(a1Notation, sheetId = 0) {
  if (
    !a1Notation ||
    typeof a1Notation !== "string" ||
    sheetId === null ||
    isNaN(sheetId)
  ) {
    throw new Error("Invalid inputs for A1 to GridRange conversion.");
  }

  const parsed = parseA1Notation_(a1Notation);
  const rangePart = parsed ? parsed.range : a1Notation;

  const { col, row } = rangePart
    .toUpperCase()
    .split(":")
    .reduce(
      (o, part) => {
        const r1 = part.match(/[A-Z]+/);
        const r2 = part.match(/[0-9]+/);
        o.col.push(r1 ? columnLetterToIndex_(r1[0]) : null);
        o.row.push(r2 ? Number(r2[0]) : null);
        return o;
      },
      { col: [], row: [] }
    );

  // Sort to handle reverse ranges (e.g. B2:A1)
  col.sort((a, b) => a - b);
  row.sort((a, b) => a - b);

  const startCol = col[0];
  const endCol = col[1] !== undefined ? col[1] : startCol;

  const startRow = row[0];
  const endRow = row[1] !== undefined ? row[1] : startRow;

  const obj = {
    sheetId,
  };

  if (startRow !== null) obj.startRowIndex = startRow - 1;
  if (endRow !== null) obj.endRowIndex = endRow;
  if (startCol !== null) obj.startColumnIndex = startCol;
  if (endCol !== null) obj.endColumnIndex = endCol + 1;

  // Defaults for infinite ranges (e.g. "A:A")
  if (!obj.hasOwnProperty("startRowIndex") && obj.hasOwnProperty("endRowIndex"))
    obj.startRowIndex = 0;
  if (
    !obj.hasOwnProperty("startColumnIndex") &&
    obj.hasOwnProperty("endColumnIndex")
  )
    obj.startColumnIndex = 0;

  return obj;
}

/**
 * Converts GridRange to A1Notation.
 *
 * @private
 * @param {Object} gridrange
 * @param {string} sheetName
 * @return {string}
 */
function convGridRangeToA1Notation_(gridrange, sheetName = "") {
  if (!gridrange) throw new Error("GridRange object is missing.");

  const startCol = gridrange.hasOwnProperty("startColumnIndex")
    ? columnIndexToLetter_(gridrange.startColumnIndex)
    : "A";

  const startRow = gridrange.hasOwnProperty("startRowIndex")
    ? gridrange.startRowIndex + 1
    : "";

  const endCol = gridrange.hasOwnProperty("endColumnIndex")
    ? columnIndexToLetter_(gridrange.endColumnIndex - 1)
    : "";

  const endRow = gridrange.hasOwnProperty("endRowIndex")
    ? gridrange.endRowIndex
    : "";

  const startStr = `${startCol}${startRow}`;
  const endStr = `${endCol}${endRow}`;

  const rangeStr = startStr === endStr ? startStr : `${startStr}:${endStr}`;

  return sheetName ? `'${sheetName}'!${rangeStr}` : rangeStr;
}
