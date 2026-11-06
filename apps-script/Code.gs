/**
 * GENERATED — do not edit.
 *
 * Built from src/sheets/server.ts by scripts/build-sheets.js.
 * Edit the TypeScript and rebuild; changes here are lost.
 */
"use strict";
var __cocoServer = (() => {
  var __defProp = Object.defineProperty;
  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  var __export = (target, all) => {
    for (var name in all)
      __defProp(target, name, { get: all[name], enumerable: true });
  };
  var __copyProps = (to, from, except, desc) => {
    if (from && typeof from === "object" || typeof from === "function") {
      for (let key of __getOwnPropNames(from))
        if (!__hasOwnProp.call(to, key) && key !== except)
          __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
    }
    return to;
  };
  var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

  // src/sheets/server.ts
  var server_exports = {};
  __export(server_exports, {
    cocoCall: () => cocoCall,
    cocoShowSidebar: () => cocoShowSidebar,
    onInstall: () => onInstall,
    onOpen: () => onOpen
  });

  // src/refresh/types.ts
  function str(v) {
    return typeof v === "string" ? v : "";
  }
  function optStr(v) {
    return typeof v === "string" && v.length > 0 ? v : void 0;
  }
  function toConnection(id, raw) {
    if (!id) return void 0;
    if (typeof raw !== "object" || raw === null) return void 0;
    const r = raw;
    const sql = str(r.sql).trim();
    const tableName = str(r.tableName).trim();
    if (!sql || !tableName) return void 0;
    return {
      id,
      sql,
      tableName,
      tableId: optStr(r.tableId),
      sheet: str(r.sheet),
      lastRefreshed: optStr(r.lastRefreshed)
    };
  }

  // src/refresh/make-room.ts
  function makeRoomSpan(tableName, others, shift) {
    const { axis, at, delta } = shift;
    const across = axis === "rows" ? "cols" : "rows";
    const movesWholeFrom = delta > 0 ? at : at - delta;
    let [lo, hi] = shift.span;
    for (let widened = true; widened; ) {
      widened = false;
      for (const box of others) {
        const [start, end] = box[axis];
        const [from, until] = box[across];
        if (until <= lo || from >= hi || end <= at) continue;
        if (start < movesWholeFrom) throw new Error(splitRefusal(tableName, axis, delta, box.name));
        if (from < lo || until > hi) {
          lo = Math.min(lo, from);
          hi = Math.max(hi, until);
          widened = true;
        }
      }
    }
    return [lo, hi];
  }
  function splitRefusal(tableName, axis, delta, neighbour) {
    const n = Math.abs(delta);
    const change = axis === "rows" ? `${delta > 0 ? "grew" : "shrank"} by ${n} row${n === 1 ? "" : "s"}` : `${delta > 0 ? "gained" : "lost"} ${n} column${n === 1 ? "" : "s"}`;
    return `'${tableName}' ${change}, and ${delta > 0 ? "making room" : "closing the gap"} would split '${neighbour}' beside it. Move one of the tables, or refresh a different result.`;
  }

  // src/utils/error-codes.ts
  var ERROR_CODES = {
    // Addresses, sheets, and reads.
    /** An A1 address (cell, range, or sheet!range) failed to parse. */
    BAD_A1: "E_BAD_A1",
    /** A tool named a sheet no worksheet matches. */
    SHEET_NOT_FOUND: "E_SHEET_NOT_FOUND",
    /** A SQL result's footprint runs past the bottom or right edge of the grid. */
    GRID_OVERFLOW: "E_GRID_OVERFLOW",
    /** A read asked for more cells, columns, or characters than one call returns. */
    READ_LIMIT: "E_READ_LIMIT",
    /** The host cannot do what the read asked (e.g. inspect formatting). */
    UNSUPPORTED: "E_UNSUPPORTED",
    /** No citation with that id is in the store. */
    UNKNOWN_CITATION: "E_UNKNOWN_CITATION",
    /** The model called a tool this pane does not register. */
    UNKNOWN_TOOL: "E_UNKNOWN_TOOL",
    // Tool arguments.
    /** A sheet name (insert_sheet, or a rename / copy target) violated Excel's rules. */
    BAD_NAME: "E_BAD_NAME",
    /** A model-supplied table name is inside the `CocoTable_` namespace. */
    RESERVED_TABLE_NAME: "E_RESERVED_TABLE_NAME",
    /** A model-supplied table name Excel itself would refuse. */
    BAD_TABLE_NAME: "E_BAD_TABLE_NAME",
    /** apply_sort_filter column_index is >= the range's column count. */
    COLUMN_OUT_OF_RANGE: "E_COLUMN_OUT_OF_RANGE",
    /** manage_rows_cols received neither `indices` nor `range`. */
    MISSING_TARGET: "E_MISSING_TARGET",
    /** manage_rows_cols index is out of grid bounds, NaN, or non-integer. */
    BAD_INDEX: "E_BAD_INDEX",
    /** manage_rows_cols action=insert received a non-positive count. */
    BAD_COUNT: "E_BAD_COUNT",
    /** manage_rows_cols action=resize was called without size_pixels. */
    MISSING_SIZE: "E_MISSING_SIZE",
    /** manage_rows_cols action=resize received a malformed size_pixels. */
    BAD_SIZE: "E_BAD_SIZE",
    /** A conditional-format `format` object is not one Excel accepts. */
    CF_FORMAT: "E_CF_FORMAT",
    /** A conditional format on a table targets something other than its data body. */
    CF_TARGET: "E_CF_TARGET",
    /** A filter's range overlaps a table it does not cover header to footer. */
    FILTER_TABLE_RANGE: "E_FILTER_TABLE_RANGE",
    // Tables and refresh.
    /** update_table `action` is not update, move, or delete. */
    TABLE_ACTION: "E_TABLE_ACTION",
    /** update_table mixed options that must be sent separately. */
    TABLE_UPDATE_OPTIONS: "E_TABLE_UPDATE_OPTIONS",
    /** A table move has no usable destination. */
    TABLE_DESTINATION: "E_TABLE_DESTINATION",
    /** The table moved or changed identity after its move/delete preview. */
    TABLE_STALE: "E_TABLE_STALE",
    /** A query edit named a table with no saved query behind it. */
    TABLE_NOT_CONNECTED: "E_TABLE_NOT_CONNECTED",
    /** A query edit ran but produced nothing that could be staged. */
    TABLE_QUERY_UPDATE: "E_TABLE_QUERY_UPDATE",
    /** Replacement SQL needs exactly one live table and a nonempty query. */
    REFRESH_QUERY_TARGET: "E_REFRESH_QUERY_TARGET",
    /** Query editing is not available in this pane. */
    REFRESH_UNAVAILABLE: "E_REFRESH_UNAVAILABLE",
    /** The saved query or table changed after the refresh preview. */
    REFRESH_STALE: "E_REFRESH_STALE",
    /** The cells landed but the revised query could not be saved. */
    REFRESH_QUERY_SAVE: "E_REFRESH_QUERY_SAVE",
    // PivotTables.
    /** A requested pivot field is not among the hierarchies Excel enumerated. */
    PIVOT_MISSING_FIELD: "E_PIVOT_MISSING_FIELD",
    /** The pivot name is already taken; PivotTable names are workbook-scoped. */
    PIVOT_NAME_TAKEN: "E_PIVOT_NAME_TAKEN",
    /** The pivot was created but Excel enumerated no hierarchies for it. */
    PIVOT_EMPTY: "E_PIVOT_EMPTY",
    // compose_preview bundles.
    /** `changes` was missing, empty, or not an array. */
    EMPTY_CHANGES: "E_EMPTY_CHANGES",
    /** A change item was missing or had an unrecognized `kind`. */
    BAD_KIND: "E_BAD_KIND",
    /** A change item failed per-kind validation; the message names the index and field. */
    BAD_CHANGE: "E_BAD_CHANGE",
    // Accept and the write into the sheet.
    /** No preview with that id is known. */
    UNKNOWN_PREVIEW: "E_UNKNOWN_PREVIEW",
    /** The preview was already accepted, rejected, or expired. */
    PREVIEW_NOT_PENDING: "E_PREVIEW_NOT_PENDING",
    /** An accepted change index is out of range. */
    INVALID_CHANGE_INDEX: "E_INVALID_CHANGE_INDEX",
    /** A subset was requested of a bundle that cannot be split (`allowsPartialAccept`). */
    INDIVISIBLE_PREVIEW: "E_INDIVISIBLE_PREVIEW",
    /** The user pressed Stop on an Accept. The turn still resumes, so this tells the
     *  model not to run the write again. */
    APPLY_STOPPED: "E_APPLY_STOPPED",
    /** Accept's write threw. Office's own error rides on `cause`. */
    APPLY_FAILED: "E_APPLY_FAILED",
    /** A streamed SQL write stopped partway, on either host. `incompleteWriteSentence`
     *  keys off it to say what landed. */
    SQL_WRITE_INCOMPLETE: "E_SQL_WRITE_INCOMPLETE",
    // Google Sheets.
    /** The `PreviewChange` kind has no Sheets implementation yet. */
    SHEETS_UNSUPPORTED_KIND: "E_SHEETS_UNSUPPORTED_KIND",
    /** The kind is supported but this variant of it is not (e.g. data bars). */
    SHEETS_UNSUPPORTED_FEATURE: "E_SHEETS_UNSUPPORTED_FEATURE",
    /** A change names a sheet that does not exist, and no earlier change creates it. */
    SHEETS_UNKNOWN_SHEET: "E_SHEETS_UNKNOWN_SHEET",
    /** A change payload was internally inconsistent (e.g. no target given). */
    SHEETS_BAD_CHANGE: "E_SHEETS_BAD_CHANGE",
    /** The pasted `Code.gs` predates this kind: the pane is newer than the script. */
    SHEETS_SCRIPT_OUTDATED: "E_SHEETS_SCRIPT_OUTDATED",
    /** The pane called a bridge function this `Code.gs` does not have. */
    SHEETS_NO_SERVER_FUNCTION: "E_SHEETS_NO_SERVER_FUNCTION",
    /** Apps Script never acknowledged a request that may still have completed. */
    SHEETS_ACK_UNKNOWN: "E_SHEETS_ACK_UNKNOWN",
    /** A read or write into the citation log. Shared by reads and writes, so the
     *  model hears one reason for both. */
    PROTECTED_SHEET: "E_PROTECTED_SHEET",
    // Server passthrough calls with no code of their own.
    /** Title generation returned no text. */
    TITLE_EMPTY: "E_TITLE_EMPTY",
    /** Follow-up suggestion generation failed. */
    FOLLOWUPS_FAILED: "E_FOLLOWUPS_FAILED"
  };

  // src/utils/a1.ts
  var CELL = /^([A-Za-z]{1,3})(\d+)$/;
  function colLetterToIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) {
      n = n * 26 + (ch.charCodeAt(0) - 64);
    }
    return n - 1;
  }
  function colIndexToLetter(index) {
    let n = index + 1;
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }
  function parseCell(cell) {
    const m = CELL.exec(cell.trim());
    if (!m) throw new Error(`${ERROR_CODES.BAD_A1}: Invalid A1 cell: ${cell}`);
    return { col: colLetterToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
  }
  function parseA1(address) {
    const [sheetPart, rangePart] = splitSheetAndRange(address);
    const [startRaw, endRaw] = rangePart.includes(":") ? rangePart.split(":") : [rangePart, rangePart];
    const start = parseCell(startRaw);
    const end = parseCell(endRaw);
    return {
      sheet: sheetPart,
      startCol: Math.min(start.col, end.col),
      startRow: Math.min(start.row, end.row),
      endCol: Math.max(start.col, end.col),
      endRow: Math.max(start.row, end.row)
    };
  }
  function formatA1(range, sheet) {
    const a = `${colIndexToLetter(range.startCol)}${range.startRow + 1}`;
    const b = `${colIndexToLetter(range.endCol)}${range.endRow + 1}`;
    const body2 = a === b ? a : `${a}:${b}`;
    if (!sheet) return body2;
    return `${quoteSheet(sheet)}!${body2}`;
  }
  function rangeShape(range) {
    return {
      rows: range.endRow - range.startRow + 1,
      cols: range.endCol - range.startCol + 1
    };
  }
  function rangesOverlap(a, b) {
    const x = parseA1(a);
    const y = parseA1(b);
    return x.startRow <= y.endRow && y.startRow <= x.endRow && x.startCol <= y.endCol && y.startCol <= x.endCol;
  }
  function extendAlongAxisA1(sourceBody, targetBody) {
    const source = parseA1(sourceBody);
    const target = parseA1(targetBody);
    const sameCols = source.startCol === target.startCol && source.endCol === target.endCol;
    const sameRows = source.startRow === target.startRow && source.endRow === target.endRow;
    if (!sameCols && !sameRows) {
      throw new Error(
        `fill destination ${targetBody} does not extend ${sourceBody} along one axis`
      );
    }
    const contiguous = sameCols ? target.startRow <= source.endRow + 1 && source.startRow <= target.endRow + 1 : target.startCol <= source.endCol + 1 && source.startCol <= target.endCol + 1;
    if (!contiguous) {
      throw new Error(
        `fill destination ${targetBody} is not adjacent to ${sourceBody} \u2014 AutoFill would overwrite the cells between them`
      );
    }
    return formatA1({
      startRow: Math.min(source.startRow, target.startRow),
      startCol: Math.min(source.startCol, target.startCol),
      endRow: Math.max(source.endRow, target.endRow),
      endCol: Math.max(source.endCol, target.endCol)
    });
  }
  function splitSheetAndRange(address) {
    const trimmed = address.trim();
    if (trimmed.startsWith("'")) {
      let i = 1;
      let name = "";
      while (i < trimmed.length) {
        const ch = trimmed[i];
        if (ch === "'") {
          if (trimmed[i + 1] === "'") {
            name += "'";
            i += 2;
            continue;
          }
          if (trimmed[i + 1] !== "!") {
            throw new Error(`${ERROR_CODES.BAD_A1}: Expected '!' after quoted sheet in: ${address}`);
          }
          return [name, trimmed.slice(i + 2)];
        }
        name += ch;
        i++;
      }
      throw new Error(`${ERROR_CODES.BAD_A1}: Unterminated quoted sheet in: ${address}`);
    }
    const bang = trimmed.indexOf("!");
    if (bang === -1) return [void 0, trimmed];
    return [trimmed.slice(0, bang), trimmed.slice(bang + 1)];
  }
  function quoteSheet(sheet) {
    if (/[\s'!]/.test(sheet)) return `'${sheet.replace(/'/g, "''")}'`;
    return sheet;
  }
  function splitAddresses(body2) {
    return body2.split(",").map((s) => s.trim()).filter((s) => s.length > 0);
  }
  function firstAxisIndex(addr, axis) {
    var _a, _b;
    const head = (_b = (_a = addr.split(":")[0]) == null ? void 0 : _a.trim()) != null ? _b : "";
    if (axis === "row") {
      const m2 = /(\d+)/.exec(head);
      if (!m2) return -1;
      const n = parseInt(m2[1], 10);
      return n > 0 ? n - 1 : -1;
    }
    const m = /^\$?([A-Za-z]{1,3})/.exec(head);
    if (!m) return -1;
    return colLetterToIndex(m[1]);
  }
  var HTML_ENTITY = /&(?:amp|lt|gt|quot|apos|#(\d+)|#[xX]([0-9a-fA-F]+));/g;
  function decodeSheetName(name) {
    if (!name.includes("&")) return name;
    return name.replace(HTML_ENTITY, (match, dec, hex) => {
      if (dec !== void 0) return String.fromCodePoint(Number(dec));
      if (hex !== void 0) return String.fromCodePoint(parseInt(hex, 16));
      if (match === "&amp;") return "&";
      if (match === "&lt;") return "<";
      if (match === "&gt;") return ">";
      if (match === "&quot;") return '"';
      return "'";
    });
  }

  // src/citations/layout.ts
  var CITATION_SHEET_NAME = "_CORTEX_CITATIONS";
  var CITATION_HEADERS = ["id", "kind", "sheet", "address", "createdAt", "label", "json"];
  var CITATION_SHEET_COLUMN = 2;

  // src/api/types.ts
  function importRowCount(result) {
    var _a;
    return (_a = result.writeRows) != null ? _a : result.totalRows;
  }

  // src/host/sheets/tables.ts
  var MAX_USER_NAMED = 50;
  function properties() {
    const service = globalThis.PropertiesService;
    return service.getDocumentProperties();
  }
  function nativeApi() {
    const api = globalThis.Sheets;
    if (!api) throw new Error(`${ERROR_CODES.SHEETS_SCRIPT_OUTDATED}: enable the Sheets v4 advanced service and update Coco's Apps Script manifest.`);
    return api;
  }
  function tableRange(range) {
    var _a, _b;
    return formatA1({ startRow: (_a = range.startRowIndex) != null ? _a : 0, endRow: range.endRowIndex - 1, startCol: (_b = range.startColumnIndex) != null ? _b : 0, endCol: range.endColumnIndex - 1 });
  }
  function gridRange(sheetId, address) {
    const range = parseA1(address);
    return { sheetId, startRowIndex: range.startRow, endRowIndex: range.endRow + 1, startColumnIndex: range.startCol, endColumnIndex: range.endCol + 1 };
  }
  var NAMED_TABLE_PREFIX = "named:";
  function activeSpreadsheetForTables() {
    var _a, _b, _c;
    const g = globalThis;
    return (_c = (_b = (_a = g.SpreadsheetApp) == null ? void 0 : _a.getActiveSpreadsheet) == null ? void 0 : _b.call(_a)) != null ? _c : null;
  }
  function safeReadConnections() {
    try {
      return readConnections();
    } catch {
      return [];
    }
  }
  function listTables(spreadsheetId) {
    var _a, _b, _c, _d, _e;
    const doc = nativeApi().Spreadsheets.get(spreadsheetId, { fields: "sheets(properties(sheetId,title),tables)" });
    const native = ((_a = doc.sheets) != null ? _a : []).filter((sheet) => sheet.properties.title.toLowerCase() !== CITATION_SHEET_NAME.toLowerCase()).flatMap(
      (sheet) => {
        var _a2;
        return ((_a2 = sheet.tables) != null ? _a2 : []).map((table) => ({ ...table, sheet: sheet.properties.title, address: tableRange(table.range) }));
      }
    );
    const active = activeSpreadsheetForTables();
    const namedRanges = (_c = (_b = active == null ? void 0 : active.getNamedRanges) == null ? void 0 : _b.call(active)) != null ? _c : [];
    if (namedRanges.length === 0) return native;
    const connections = safeReadConnections();
    const connectedIds = new Set(connections.map((c) => c.tableId));
    const nativeNames = new Set(native.map((t) => t.name.toLowerCase()));
    const namedTables = [];
    const userTables = [];
    for (const nr of namedRanges) {
      try {
        const name = nr.getName();
        const namedId = `${NAMED_TABLE_PREFIX}${name}`;
        const ours = /^CocoTable_/i.test(name) || connectedIds.has(namedId);
        if (!ours && userTables.length >= MAX_USER_NAMED) continue;
        const r = nr.getRange();
        const sheet = r.getSheet();
        const sheetName = sheet.getName();
        if (sheetName.toLowerCase() === CITATION_SHEET_NAME.toLowerCase()) continue;
        const sheetId = sheet.getSheetId();
        const startRow = r.getRow() - 1;
        const startCol = r.getColumn() - 1;
        const numRows = r.getNumRows();
        const numCols = r.getNumColumns();
        const range = { sheetId, startRowIndex: startRow, endRowIndex: startRow + numRows, startColumnIndex: startCol, endColumnIndex: startCol + numCols };
        const address = tableRange(range);
        const converted = native.find((t) => t.range.sheetId === sheetId && t.address === address);
        if (converted && ours) {
          for (const conn of connections) {
            if (conn.tableId === namedId) {
              try {
                saveConnection({ ...conn, tableId: converted.tableId, tableName: converted.name, sheet: converted.sheet });
              } catch {
              }
            }
          }
          try {
            nr.remove();
          } catch {
          }
          continue;
        }
        if (nativeNames.has(name.toLowerCase())) continue;
        const headers = numRows > 0 && numCols > 0 ? (_d = sheet.getRange(r.getRow(), r.getColumn(), 1, numCols).getDisplayValues()[0]) != null ? _d : [] : [];
        const maxRows = ours ? void 0 : (_e = sheet.getMaxRows) == null ? void 0 : _e.call(sheet);
        (ours ? namedTables : userTables).push({
          ...ours ? {} : { userNamed: true },
          ...maxRows !== void 0 && range.endRowIndex >= maxRows ? { reachesEnd: true } : {},
          tableId: namedId,
          name,
          sheet: sheetName,
          address,
          range,
          columnProperties: headers.map((columnName, columnIndex) => ({ columnIndex, columnName }))
        });
      } catch {
      }
    }
    return [...native, ...namedTables, ...userTables];
  }
  function removeNamedTable(name) {
    var _a, _b, _c;
    const namedRanges = (_c = (_b = (_a = activeSpreadsheetForTables()) == null ? void 0 : _a.getNamedRanges) == null ? void 0 : _b.call(_a)) != null ? _c : [];
    for (const nr of namedRanges) {
      if (nr.getName().toLowerCase() === name.toLowerCase()) nr.remove();
    }
  }
  function tableBatch(spreadsheetId, requests) {
    globalThis.SpreadsheetApp.flush();
    return nativeApi().Spreadsheets.batchUpdate({ requests }, spreadsheetId);
  }
  function findTable(spreadsheetId, key) {
    const table = listTables(spreadsheetId).find((table2) => table2.tableId === key || table2.name.toLowerCase() === key.toLowerCase());
    if (!table) throw new Error(`${ERROR_CODES.TABLE_STALE}: table '${key}' no longer exists.`);
    return table;
  }
  function checkTableWrite(spreadsheetId, guard) {
    const table = listTables(spreadsheetId).find((table2) => table2.tableId === guard.tableId);
    if (!table || table.sheet !== guard.sheet || table.address !== guard.range) {
      throw new Error(`${ERROR_CODES.TABLE_STALE}: table moved or resized during the write. Inspect the sheet before retrying.`);
    }
    if (guard.connectionId) {
      const saved = readConnections().find((connection) => connection.id === guard.connectionId);
      if ((saved == null ? void 0 : saved.sql) !== guard.expectedSql || (saved == null ? void 0 : saved.tableId) !== guard.tableId) {
        throw new Error(`${ERROR_CODES.TABLE_STALE}: the saved query changed during the write.`);
      }
    }
  }
  function readConnections() {
    return Object.entries(properties().getProperties()).flatMap(([key, value]) => {
      if (!key.startsWith("coco.refresh.")) return [];
      const id = key.slice("coco.refresh.".length);
      try {
        const connection = toConnection(id, JSON.parse(value));
        return connection ? [connection] : [];
      } catch {
        return [];
      }
    });
  }
  function saveConnection(connection, expectedSql) {
    var _a;
    const valid = toConnection(connection.id, connection);
    if (!(valid == null ? void 0 : valid.tableId)) throw new Error(`${ERROR_CODES.TABLE_NOT_CONNECTED}: a saved query needs a native table id.`);
    if (expectedSql !== void 0 && ((_a = readConnections().find((saved) => saved.id === valid.id)) == null ? void 0 : _a.sql) !== expectedSql) {
      throw new Error(`${ERROR_CODES.TABLE_STALE}: the saved query changed during the write.`);
    }
    properties().setProperty(`coco.refresh.${valid.id}`, JSON.stringify(valid));
  }
  function removeConnection(id) {
    properties().deleteProperty(`coco.refresh.${id}`);
  }
  function assertTableSpace(tables, sheetId, change) {
    const upgrading = (table) => !change.columns && !table.userNamed && table.tableId.startsWith(NAMED_TABLE_PREFIX) && table.range.sheetId === sheetId && table.address === change.range;
    if (tables.some((table) => !upgrading(table) && (table.name.toLowerCase() === change.tableName.toLowerCase() || !table.userNamed && table.range.sheetId === sheetId && rangesOverlap(table.address, change.range)))) {
      throw new Error(`${ERROR_CODES.TABLE_DESTINATION}: table name or range overlaps an existing table.`);
    }
  }
  var NEUTRAL_HEADER = { rgbColor: { red: 241 / 255, green: 243 / 255, blue: 244 / 255 } };
  var PALE_BLUE = { rgbColor: { red: 248 / 255, green: 249 / 255, blue: 250 / 255 } };
  var WHITE = { rgbColor: { red: 1, green: 1, blue: 1 } };
  function createNativeTable(spreadsheetId, sheetId, change) {
    var _a, _b, _c, _d;
    const existing = listTables(spreadsheetId);
    assertTableSpace(existing, sheetId, change);
    const upgraded = existing.find(
      (t) => !t.userNamed && t.tableId.startsWith(NAMED_TABLE_PREFIX) && t.range.sheetId === sheetId && t.address === change.range
    );
    const range = gridRange(sheetId, change.range);
    const header = { ...range, endRowIndex: ((_a = range.startRowIndex) != null ? _a : 0) + 1 };
    const sameName = upgraded !== void 0 && upgraded.name.toLowerCase() === change.tableName.toLowerCase();
    if (sameName) removeNamedTable(upgraded.name);
    const restore = () => {
      var _a2;
      if (!sameName) return;
      try {
        const doc = activeSpreadsheetForTables();
        const sheet = (_a2 = doc == null ? void 0 : doc.getSheetByName) == null ? void 0 : _a2.call(doc, upgraded.sheet);
        if ((doc == null ? void 0 : doc.setNamedRange) && sheet) doc.setNamedRange(upgraded.name, sheet.getRange(upgraded.address));
      } catch {
      }
    };
    let reply;
    try {
      reply = tableBatch(spreadsheetId, [
        { addTable: { table: {
          name: change.tableName,
          range,
          rowsProperties: { headerColorStyle: NEUTRAL_HEADER, firstBandColorStyle: WHITE, secondBandColorStyle: PALE_BLUE }
        } } },
        // addTable.columnProperties misplaces headers on offset ranges. Write the
        // header after creating the table, with explicit grid coordinates instead.
        ...change.columns ? [{ updateCells: {
          range: header,
          rows: [{ values: change.columns.map((stringValue) => ({ userEnteredValue: { stringValue } })) }],
          fields: "userEnteredValue"
        } }] : [],
        { repeatCell: {
          range: header,
          cell: { userEnteredFormat: { textFormat: { foregroundColorStyle: { rgbColor: { red: 0, green: 0, blue: 0 } } } } },
          fields: "userEnteredFormat.textFormat.foregroundColorStyle"
        } }
      ]);
    } catch (error) {
      restore();
      throw error;
    }
    const table = (_d = (_c = (_b = reply.replies) == null ? void 0 : _b[0]) == null ? void 0 : _c.addTable) == null ? void 0 : _d.table;
    if (!table) {
      restore();
      throw new Error(`${ERROR_CODES.APPLY_FAILED}: the Sheets API did not return the created table. Inspect the sheet before retrying.`);
    }
    if (upgraded) {
      if (!sameName) removeNamedTable(upgraded.name);
      for (const conn of safeReadConnections()) {
        if (conn.tableId === upgraded.tableId) {
          saveConnection({ ...conn, tableId: table.tableId, tableName: change.tableName, sheet: change.sheet });
        }
      }
    }
    if (change.connection) saveConnection({ ...change.connection, tableId: table.tableId, tableName: table.name, sheet: change.sheet });
    return table;
  }
  function tableVariant(change) {
    var _a;
    if (change.kind === "create_table") return change.hasHeaders ? void 0 : "Google Sheets tables require a header row; include headers in the range.";
    if (((_a = change.lifecycle) == null ? void 0 : _a.action) === "move") return "Moving native tables is not available on Google Sheets; move it in Sheets instead.";
    if ([change.totalRow, change.totalAggregations, change.headerRow, change.tableStyle, change.filter].some((value) => value !== void 0)) {
      return "Google Sheets table edits support banded rows and sort, not totals, header visibility, Excel styles or filters.";
    }
    return void 0;
  }
  function resizeTable(spreadsheetId, change, beforeBatch) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _i, _j, _k, _l;
    const saved = readConnections().find((connection) => connection.id === change.connection.id);
    if (!saved || saved.sql !== change.expectedSql || saved.tableId !== change.connection.tableId) {
      throw new Error(`${ERROR_CODES.TABLE_STALE}: the saved query changed; request a new preview.`);
    }
    const all = listTables(spreadsheetId);
    const activeId = (_b = (_a = readConnections().find((connection) => connection.id === saved.id)) == null ? void 0 : _a.tableId) != null ? _b : saved.tableId;
    const table = all.find((table2) => table2.tableId === activeId);
    if (!table) throw new Error(`${ERROR_CODES.TABLE_STALE}: the connected table is gone.`);
    if ((_c = table.rowsProperties) == null ? void 0 : _c.footerColorStyle) throw new Error("Remove the native table footer before refreshing; nothing was written.");
    const range = table.range;
    const startRow = (_d = range.startRowIndex) != null ? _d : 0;
    const startCol = (_e = range.startColumnIndex) != null ? _e : 0;
    const newRows = importRowCount(change.sqlResult);
    const endRow = startRow + Math.max(1, newRows) + 1;
    const endCol = startCol + change.columns.length;
    const target = { sheetId: range.sheetId, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: startCol, endColumnIndex: endCol };
    const address = tableRange(target);
    beforeBatch == null ? void 0 : beforeBatch(table.sheet, address);
    const requests = [];
    const sheetId = range.sheetId;
    const userBoxes = /* @__PURE__ */ new Set();
    let others = all.filter((other) => other.range.sheetId === sheetId && other.tableId !== table.tableId).map((other) => {
      var _a2, _b2;
      if (other.userNamed) userBoxes.add(other.name);
      return {
        name: other.name,
        rows: [(_a2 = other.range.startRowIndex) != null ? _a2 : 0, other.range.endRowIndex],
        cols: [(_b2 = other.range.startColumnIndex) != null ? _b2 : 0, other.range.endColumnIndex]
      };
    });
    const tearable = (axis, span) => others.filter((other) => {
      const [from, until] = other[axis === "rows" ? "cols" : "rows"];
      return !(userBoxes.has(other.name) && from >= span[0] && until <= span[1]);
    });
    const columnDelta = endCol - range.endColumnIndex;
    if (columnDelta) {
      const at = columnDelta > 0 ? range.endColumnIndex : endCol;
      const span = [startRow, range.endRowIndex];
      const [top, bottom] = makeRoomSpan(table.name, tearable("cols", span), { axis: "cols", at, delta: columnDelta, span });
      const rows = { sheetId, startRowIndex: top, endRowIndex: bottom };
      requests.push(columnDelta > 0 ? { insertRange: { range: { ...rows, startColumnIndex: at, endColumnIndex: at + columnDelta }, shiftDimension: "COLUMNS" } } : { deleteRange: { range: { ...rows, startColumnIndex: endCol, endColumnIndex: range.endColumnIndex }, shiftDimension: "COLUMNS" } });
      others = others.map((other) => other.cols[0] >= at && other.rows[0] >= top && other.rows[1] <= bottom ? { ...other, cols: [other.cols[0] + columnDelta, other.cols[1] + columnDelta] } : other);
    }
    const rowDelta = endRow - range.endRowIndex;
    if (rowDelta) {
      const at = rowDelta > 0 ? range.endRowIndex : endRow;
      const span = [startCol, endCol];
      const [left, right] = makeRoomSpan(table.name, tearable("rows", span), { axis: "rows", at, delta: rowDelta, span });
      const columns = { sheetId, startColumnIndex: left, endColumnIndex: right };
      requests.push(rowDelta > 0 ? { insertRange: { range: { ...columns, startRowIndex: at, endRowIndex: at + rowDelta }, shiftDimension: "ROWS" } } : { deleteRange: { range: { ...columns, startRowIndex: endRow, endRowIndex: range.endRowIndex }, shiftDimension: "ROWS" } });
    }
    if (table.tableId.startsWith(NAMED_TABLE_PREFIX)) {
      if (requests.length > 0) tableBatch(spreadsheetId, requests);
      const doc = activeSpreadsheetForTables();
      const sheet = (_f = doc == null ? void 0 : doc.getSheetByName) == null ? void 0 : _f.call(doc, table.sheet);
      if ((doc == null ? void 0 : doc.setNamedRange) && sheet) {
        const maxRows = (_h = (_g = sheet.getMaxRows) == null ? void 0 : _g.call(sheet)) != null ? _h : endRow;
        const maxCols = (_j = (_i = sheet.getMaxColumns) == null ? void 0 : _i.call(sheet)) != null ? _j : endCol;
        if (endRow > maxRows) (_k = sheet.insertRowsAfter) == null ? void 0 : _k.call(sheet, maxRows, endRow - maxRows);
        if (endCol > maxCols) (_l = sheet.insertColumnsAfter) == null ? void 0 : _l.call(sheet, maxCols, endCol - maxCols);
        doc.setNamedRange(table.name, sheet.getRange(address));
      }
      return { ...table, range: target, address };
    }
    if (requests.length > 0) {
      requests.push({ updateTable: { table: { tableId: table.tableId, range: target }, fields: "range" } });
      tableBatch(spreadsheetId, requests);
    }
    return { ...table, range: target, address };
  }

  // src/preview/types.ts
  function pivotMeasureFields(change) {
    if (change.valueFields && change.valueFields.length > 0) return change.valueFields;
    return change.values.map((field) => ({ field }));
  }
  function fillDestinationA1(change) {
    const sourceBody = splitSheetAndRange(change.sourceRange)[1];
    if (typeof change.targetRangeOrCount === "string") {
      return extendAlongAxisA1(
        sourceBody,
        splitSheetAndRange(change.targetRangeOrCount)[1]
      );
    }
    const src = parseA1(sourceBody);
    const n = change.targetRangeOrCount;
    let { startRow, startCol, endRow, endCol } = src;
    switch (change.direction) {
      case "down":
        endRow = src.endRow + n;
        break;
      case "up":
        startRow = Math.max(0, src.startRow - n);
        break;
      case "right":
        endCol = src.endCol + n;
        break;
      case "left":
        startCol = Math.max(0, src.startCol - n);
        break;
    }
    return formatA1({ startRow, startCol, endRow, endCol });
  }
  function landingSheet(changes) {
    var _a;
    for (let i = changes.length - 1; i >= 0; i--) {
      const c = changes[i];
      if (c.kind === "refresh_table" || c.kind === "manage_sheet") continue;
      const sheet = c.kind === "create_pivot" ? c.destinationSheet : c.kind === "update_cells" ? (_a = c.edits[c.edits.length - 1]) == null ? void 0 : _a.sheet : "sheet" in c ? c.sheet : void 0;
      if (sheet === void 0) continue;
      return sheet && sheet.toLowerCase() !== CITATION_SHEET_NAME.toLowerCase() ? decodeSheetName(sheet) : void 0;
    }
    return void 0;
  }

  // src/preview/cf-colors.ts
  var HIGHLIGHT_HEX = {
    red: "#FEE2E2",
    green: "#DCFCE7",
    yellow: "#FEF3C7",
    blue: "#DBEAFE"
  };
  var COLOR_SCALE_PALETTES = {
    "red-yellow-green": ["#F87171", "#FDE68A", "#4ADE80"],
    "green-yellow-red": ["#4ADE80", "#FDE68A", "#F87171"],
    "blue-white-red": ["#60A5FA", "#FFFFFF", "#F87171"],
    "red-white-blue": ["#F87171", "#FFFFFF", "#60A5FA"],
    "white-green": ["#FFFFFF", "#BBF7D0", "#15803D"],
    "white-red": ["#FFFFFF", "#FECACA", "#B91C1C"]
  };
  function cellValueOperand(v) {
    if (typeof v === "number") return void 0;
    const s = v.trim();
    const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (iso) return `DATE(${Number(iso[1])},${Number(iso[2])},${Number(iso[3])})`;
    const asIs = s !== "" && Number.isFinite(Number(s.replace(/%$/, ""))) || /^[="]/.test(s) || /^(TRUE|FALSE)$/i.test(s);
    return asIs ? void 0 : `"${s.replace(/"/g, '""')}"`;
  }

  // src/host/sheets/coerce.ts
  var TranslateError = class extends Error {
    constructor(code, message) {
      super(`${code}: ${message}`);
      this.code = code;
      this.name = "TranslateError";
    }
  };
  function wireSafe(values, displayValues) {
    if (!values.some((row) => row.some((v) => v instanceof Date))) return values;
    const shown = displayValues();
    return values.map((row, r) => row.map((v, c) => v instanceof Date ? shown[r][c] : v));
  }
  var SHEETS_EPOCH_DAY = "1899-12-30";
  function isPercentFormat(format) {
    if (!format) return false;
    return format.replace(/"[^"]*"/g, "").replace(/\\./g, "").includes("%");
  }
  function modelValues(values, numberFormats, displayValues, formatDate) {
    let shown = null;
    return values.map(
      (row, r) => row.map((v, c) => {
        var _a, _b;
        if (v instanceof Date) {
          const day = formatDate(v, "yyyy-MM-dd");
          const time = formatDate(v, "HH:mm:ss");
          if (day === SHEETS_EPOCH_DAY) return time;
          return time === "00:00:00" ? day : `${day}T${time}`;
        }
        if (typeof v === "number" && isPercentFormat((_a = numberFormats[r]) == null ? void 0 : _a[c])) {
          shown != null ? shown : shown = displayValues();
          return ((_b = shown[r]) == null ? void 0 : _b[c]) || v;
        }
        return v;
      })
    );
  }

  // src/host/sheets/spreadsheet-app-driver.ts
  var SPREADSHEET_APP_KINDS = [
    "create_table",
    "update_table",
    "write_range",
    "update_cells",
    "apply_formula",
    "clear_range",
    "insert_sheet",
    "manage_sheet",
    "manage_rows_cols",
    "apply_sort_filter",
    "set_cell_format",
    "create_chart",
    "create_pivot",
    "apply_conditional_format",
    "apply_data_validation",
    "find_replace",
    "fill",
    "create_named_range",
    "update_chart",
    "delete_chart",
    "update_pivot"
  ];
  function driverSupports(kind) {
    return SPREADSHEET_APP_KINDS.includes(kind);
  }
  function applyViaSpreadsheetApp(ss2, changes, opts) {
    var _a;
    const createdTables = /* @__PURE__ */ new Set();
    for (const c of changes) {
      if (!driverSupports(c.kind)) {
        throw new TranslateError(
          ERROR_CODES.SHEETS_UNSUPPORTED_KIND,
          `${c.kind} has no SpreadsheetApp implementation; the registry should have hidden this tool`
        );
      }
      const intoLog = citationLogUse(c);
      if (intoLog) throw new TranslateError(ERROR_CODES.PROTECTED_SHEET, intoLog);
      const unsupported = unsupportedVariant(c);
      if (unsupported) throw new TranslateError(ERROR_CODES.SHEETS_UNSUPPORTED_FEATURE, unsupported);
      if (isCitationLog(opts.activeSheet) && (usesActiveSheet(c) || sourceSheetsOf(c).includes(null))) {
        throw new TranslateError(
          ERROR_CODES.PROTECTED_SHEET,
          `the active sheet is ${CITATION_SHEET_NAME}, Coco's citation log; switch to the sheet to change`
        );
      }
      if (c.kind === "insert_sheet" && ss2.getSheetByName(decodeSheetName(c.sheet))) {
        throw new TranslateError(
          ERROR_CODES.SHEETS_BAD_CHANGE,
          `a sheet named ${JSON.stringify(c.sheet)} already exists`
        );
      }
      if (c.kind === "manage_sheet" && (c.action === "rename" || c.action === "copy") && c.newName) {
        const taken = ss2.getSheetByName(decodeSheetName(c.newName));
        if (taken && !(c.action === "rename" && taken === ss2.getSheetByName(decodeSheetName(c.sheet)))) {
          throw new TranslateError(ERROR_CODES.SHEETS_BAD_CHANGE, `a sheet named ${JSON.stringify(c.newName)} already exists`);
        }
      }
      if (c.kind === "update_chart" || c.kind === "delete_chart") findChart(ss2, c.sheet, c.chartName, opts.activeSheet);
      if (c.kind === "update_pivot") findPivot(ss2, c.pivotName);
      if (c.kind === "create_table" && !createdTables.has(c.tableName.toLowerCase())) {
        const sheet = c.sheet ? ss2.getSheetByName(decodeSheetName(c.sheet)) : ss2.getSheetByName(opts.activeSheet);
        assertTableSpace(listTables(ss2.getId()), sheet == null ? void 0 : sheet.getSheetId(), c);
      }
      if (c.kind === "create_table") createdTables.add(c.tableName.toLowerCase());
      if (c.kind === "update_table" && !createdTables.has(c.tableName.toLowerCase())) tableUpdate(ss2, c);
    }
    const view = { sheet: ss2.getActiveSheet(), moved: false };
    const produced = producedSheets(changes, opts.activeSheet);
    for (const [name, position] of sheetsToEnsure(changes)) {
      if (produced.has(name) || ss2.getSheetByName(name)) continue;
      if (position === void 0) ss2.insertSheet(name);
      else ss2.insertSheet(name, position);
      view.moved = true;
    }
    for (const [name, need] of neededExtents(changes, opts.activeSheet)) {
      const sheet = requireSheet(ss2, (_a = produced.get(name)) != null ? _a : name);
      const rows = sheet.getMaxRows();
      const cols = sheet.getMaxColumns();
      if (need.rows > rows) sheet.insertRowsAfter(rows, need.rows - rows);
      if (need.cols > cols) sheet.insertColumnsAfter(cols, need.cols - cols);
    }
    let cellsWritten = 0;
    for (const c of changes) {
      cellsWritten += applyOne(ss2, c, opts, view);
    }
    const landing = opts.land === false ? void 0 : landingSheet(changes);
    const target = landing ? ss2.getSheetByName(landing) : null;
    const show = target && !target.isSheetHidden() ? target : view.moved ? view.sheet : null;
    try {
      if (show && !show.isSheetHidden()) show.activate();
    } catch {
    }
    return { cellsWritten };
  }
  function moveSheet(ss2, sheet, position, view) {
    const hidden = sheet.isSheetHidden();
    sheet.activate();
    ss2.moveActiveSheet(position + 1);
    if (hidden) sheet.hideSheet();
    view.moved = true;
  }
  function requireSheet(ss2, name) {
    const s = ss2.getSheetByName(name);
    if (!s) {
      throw new TranslateError(
        ERROR_CODES.SHEETS_UNKNOWN_SHEET,
        `sheet ${JSON.stringify(name)} does not exist`
      );
    }
    return s;
  }
  function sheetOf(ss2, name, active) {
    return requireSheet(ss2, name ? decodeSheetName(name) : active);
  }
  function body(a1) {
    return splitSheetAndRange(a1)[1];
  }
  function tableUpdate(ss2, c) {
    var _a, _b, _c, _d, _e;
    const table = findTable(ss2.getId(), (_b = (_a = c.lifecycle) == null ? void 0 : _a.tableId) != null ? _b : c.tableName);
    const sheet = requireSheet(ss2, table.sheet);
    if (c.lifecycle && (table.address !== c.lifecycle.range || table.sheet !== c.lifecycle.sheet || table.name !== c.tableName)) {
      throw new Error(`${ERROR_CODES.TABLE_STALE}: table changed after preview; request a new preview.`);
    }
    if (c.bandedRows !== void 0 && table.tableId.startsWith(NAMED_TABLE_PREFIX)) {
      throw new Error("Row banding requires a native Google Sheets table; use create_table or set_cell_format.");
    }
    if (((_c = c.sort) == null ? void 0 : _c.length) && table.userNamed) {
      throw new Error(`Sorting '${table.name}' needs a table; it is a named range. Convert it with create_table first.`);
    }
    const columns = (_d = table.columnProperties) != null ? _d : [];
    const sort = (_e = c.sort) == null ? void 0 : _e.map((field) => {
      var _a2, _b2;
      const column = columns.find((column2) => column2.columnName === field.column);
      if (!column) throw new Error(`Unknown table column '${field.column}'`);
      return { column: ((_a2 = table.range.startColumnIndex) != null ? _a2 : 0) + ((_b2 = column.columnIndex) != null ? _b2 : 0) + 1, ascending: field.direction === "asc" };
    });
    return { table, sheet, sort };
  }
  function applyOne(ss2, c, opts, view) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _i, _j, _k, _l, _m, _n, _o;
    switch (c.kind) {
      case "create_table": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        if (c.columns) {
          const at = parseA1(body(c.range));
          sheet.getRange(at.startRow + 1, at.startCol + 1, 1, c.columns.length).setValues([c.columns.map(asLiteral)]);
          ss2.setNamedRange(c.tableName, sheet.getRange(body(c.range)));
          return 0;
        }
        createNativeTable(ss2.getId(), sheet.getSheetId(), { ...c, sheet: sheet.getName() });
        return 0;
      }
      case "update_table": {
        const id = ss2.getId();
        const { table, sheet, sort } = tableUpdate(ss2, c);
        if (c.lifecycle) {
          if (table.tableId.startsWith(NAMED_TABLE_PREFIX)) removeNamedTable(table.name);
          else tableBatch(id, [{ deleteTable: { tableId: table.tableId } }]);
          const range = sheet.getRange(table.address);
          (_a = range.getBandings) == null ? void 0 : _a.call(range).forEach((b) => b.remove());
          range.clear();
          for (const connection of readConnections()) if (connection.tableId === table.tableId) removeConnection(connection.id);
          return 0;
        }
        if (c.bandedRows !== void 0 && !table.tableId.startsWith(NAMED_TABLE_PREFIX)) tableBatch(id, [{ updateTable: {
          table: { tableId: table.tableId, rowsProperties: { secondBandColorStyle: c.bandedRows ? PALE_BLUE : WHITE } },
          fields: "rowsProperties.secondBandColorStyle"
        } }]);
        if (sort == null ? void 0 : sort.length) {
          const range = parseA1(table.address);
          if (range.endRow > range.startRow) sheet.getRange(formatA1({ ...range, startRow: range.startRow + 1 })).sort(sort);
        }
        return 0;
      }
      case "write_range": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const at = parseA1(body(c.range));
        const shape = rangeShape(at);
        const values = rectangular(c.after, shape.rows, shape.cols);
        const formulas = c.formulasAfter ? rectangular(c.formulasAfter, shape.rows, shape.cols) : null;
        formulas == null ? void 0 : formulas.forEach(
          (row, r) => row.forEach((f, col) => {
            var _a2;
            if (f != null && f !== "" && !isFormula(f) && String((_a2 = values[r][col]) != null ? _a2 : "") !== f) values[r][col] = f;
          })
        );
        sheet.getRange(body(c.range)).setValues(values.map((row) => row.map(asLiteral)));
        for (const run of formulas ? formulaRuns(formulas) : []) {
          sheet.getRange(at.startRow + run.row + 1, at.startCol + run.col + 1, run.formulas.length, 1).setFormulas(run.formulas.map((f) => [f]));
        }
        return shape.rows * shape.cols;
      }
      case "update_cells": {
        for (const e of c.edits) {
          const sheet = sheetOf(ss2, e.sheet, opts.activeSheet);
          const range = sheet.getRange(body(e.cell));
          const f = e.formulaAfter;
          if (f != null && isFormula(f)) range.setFormula(f.trim());
          else range.setValue(asLiteral(f != null && f !== "" ? f : e.after));
        }
        return c.edits.length;
      }
      case "apply_formula": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const shape = rangeShape(parseA1(c.range));
        sheet.getRange(body(c.range)).setFormula(withEquals(c.formulaAfter));
        return shape.rows * shape.cols;
      }
      case "clear_range": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const range = sheet.getRange(body(c.range));
        switch (c.applyTo) {
          case "contents":
            range.clearContent();
            break;
          case "formats":
            (_b = range.getBandings) == null ? void 0 : _b.call(range).forEach((b) => b.remove());
            range.clearFormat();
            range.setNumberFormat("General");
            break;
          case "all":
            (_c = range.getBandings) == null ? void 0 : _c.call(range).forEach((b) => b.remove());
            range.clear();
            range.clearNote();
            break;
        }
        return 0;
      }
      case "insert_sheet":
        return 0;
      case "manage_sheet": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        switch (c.action) {
          case "rename":
            if (!c.newName) {
              throw new TranslateError(ERROR_CODES.SHEETS_BAD_CHANGE, "rename requires newName");
            }
            sheet.setName(decodeSheetName(c.newName));
            break;
          case "delete":
            ss2.deleteSheet(sheet);
            break;
          case "hide":
            sheet.hideSheet();
            break;
          case "unhide":
            sheet.showSheet();
            break;
          case "freeze_panes":
            if (c.freezeRows !== void 0) sheet.setFrozenRows(c.freezeRows);
            if (c.freezeCols !== void 0) sheet.setFrozenColumns(c.freezeCols);
            break;
          case "copy": {
            const copy = sheet.copyTo(ss2);
            view.moved = true;
            if (c.newName) copy.setName(decodeSheetName(c.newName));
            if (typeof c.position === "number") moveSheet(ss2, copy, c.position, view);
            break;
          }
          case "reorder":
            if (typeof c.position === "number") moveSheet(ss2, sheet, c.position, view);
            break;
        }
        return 0;
      }
      case "manage_rows_cols": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const isRow = c.target === "row";
        if (c.action === "insert" && !(c.indices && c.indices.length > 0) && c.range) {
          for (const run of rangeRuns(c).sort((a, b) => b.start - a.start)) {
            const n = run.end - run.start + Math.max(1, (_d = c.count) != null ? _d : 1);
            if (isRow) sheet.insertRows(run.start + 1, n);
            else sheet.insertColumns(run.start + 1, n);
          }
          return 0;
        }
        const indices = resolveIndices(c);
        if (c.action === "insert") {
          for (const i of [...indices].sort((a, b) => b - a)) {
            if (isRow) sheet.insertRows(i + 1, (_e = c.count) != null ? _e : 1);
            else sheet.insertColumns(i + 1, (_f = c.count) != null ? _f : 1);
          }
          return 0;
        }
        const px = c.sizePixels;
        if (c.action === "resize" && px === void 0) {
          throw new TranslateError(ERROR_CODES.SHEETS_BAD_CHANGE, "resize requires sizePixels");
        }
        for (const { start, count } of contiguousBlocks(indices).reverse()) {
          const pos = start + 1;
          switch (c.action) {
            case "delete":
              if (isRow) sheet.deleteRows(pos, count);
              else sheet.deleteColumns(pos, count);
              break;
            case "hide":
              if (isRow) sheet.hideRows(pos, count);
              else sheet.hideColumns(pos, count);
              break;
            case "unhide":
              if (isRow) sheet.showRows(pos, count);
              else sheet.showColumns(pos, count);
              break;
            case "resize":
              if (isRow) sheet.setRowHeights(pos, count, px);
              else sheet.setColumnWidths(pos, count, px);
              break;
            case "autofit":
              if (isRow) sheet.autoResizeRows(pos, count);
              else sheet.autoResizeColumns(pos, count);
              break;
          }
        }
        return 0;
      }
      case "apply_sort_filter": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const sort = (_g = c.sort) != null ? _g : [];
        const a1 = parseA1(c.range);
        const hasHeaders = c.hasHeaders !== false;
        const startRow = a1.startRow + (hasHeaders ? 1 : 0) + 1;
        const numRows = a1.endRow + 1 - startRow + 1;
        if (numRows <= 0) {
          throw new TranslateError(
            ERROR_CODES.SHEETS_BAD_CHANGE,
            `sort range ${c.range} has no data rows below its header`
          );
        }
        const target = sheet.getRange(
          startRow,
          a1.startCol + 1,
          numRows,
          a1.endCol - a1.startCol + 1
        );
        target.sort(
          sort.map((s) => ({
            column: a1.startCol + s.columnIndex + 1,
            ascending: s.direction !== "desc"
          }))
        );
        return 0;
      }
      case "create_chart": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const src = splitSheetRef(c.sourceRange);
        const dataSheet = src.sheet ? sheetOf(ss2, src.sheet, opts.activeSheet) : sheet;
        const anchorCell2 = parseA1((_h = c.anchor) != null ? _h : "A1");
        const data = dataSheet.getRange(src.body);
        const builder = sheet.newChart().addRange(data).setNumHeaders(headerRows(data)).setChartType(chartTypeFor(c.chartType)).setPosition(anchorCell2.startRow + 1, anchorCell2.startCol + 1, 0, 0);
        setChartDisplay(builder, c, c.chartType === "bar");
        if (c.chartType === "donut") builder.setOption("pieHole", 0.4);
        sheet.insertChart(builder.build());
        return 0;
      }
      case "create_pivot": {
        const src = splitSheetRef(c.sourceRange);
        const dataSheet = sheetOf(ss2, src.sheet, opts.activeSheet);
        const destSheet = sheetOf(ss2, c.destinationSheet, opts.activeSheet);
        const sourceRange = dataSheet.getRange(src.body);
        const headerRow = parseA1(src.body);
        const headers = dataSheet.getRange(headerRow.startRow + 1, headerRow.startCol + 1, 1, headerRow.endCol - headerRow.startCol + 1).getDisplayValues()[0];
        const indexOf = (field) => pivotColumn(headers, field, c.sourceRange, headerRow.startCol + 1);
        const rows = c.rows.map(indexOf);
        const cols = c.columns.map(indexOf);
        const values = pivotMeasureFields(c).map((v) => [indexOf(v.field), summarizeFunction(v.agg)]);
        const pivot = destSheet.getRange(body(c.destination)).createPivotTable(sourceRange);
        for (const i of rows) pivot.addRowGroup(i);
        for (const i of cols) pivot.addColumnGroup(i);
        for (const [i, fn] of values) pivot.addPivotValue(i, fn);
        return 0;
      }
      case "set_cell_format": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const range = sheet.getRange(body(c.range));
        const f = c.font;
        if ((f == null ? void 0 : f.bold) !== void 0) range.setFontWeight(f.bold ? "bold" : "normal");
        if ((f == null ? void 0 : f.italic) !== void 0) range.setFontStyle(f.italic ? "italic" : "normal");
        if ((f == null ? void 0 : f.underline) !== void 0) range.setFontLine(f.underline ? "underline" : "none");
        if ((f == null ? void 0 : f.size) !== void 0) range.setFontSize(f.size);
        if ((f == null ? void 0 : f.name) !== void 0) range.setFontFamily(f.name);
        if ((f == null ? void 0 : f.color) !== void 0) range.setFontColor(normaliseHex(f.color));
        if (c.fill) range.setBackground(normaliseHex(c.fill.color));
        if (c.numberFormat !== void 0) range.setNumberFormat(c.numberFormat.trim());
        const a = c.alignment;
        if ((a == null ? void 0 : a.horizontal) !== void 0) range.setHorizontalAlignment(a.horizontal);
        if ((a == null ? void 0 : a.vertical) !== void 0) range.setVerticalAlignment(a.vertical);
        if ((a == null ? void 0 : a.wrapText) !== void 0) range.setWrap(a.wrapText);
        if (c.border) {
          const style = (_i = c.border.style) != null ? _i : "thin";
          if (style === "none") range.setBorder(false, false, false, false, null, null, null, null);
          else {
            const color = c.border.color ? normaliseHex(c.border.color) : null;
            range.setBorder(true, true, true, true, null, null, color, borderStyle(style));
          }
        }
        return 0;
      }
      case "apply_conditional_format": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const a1 = body(c.range);
        const rules = sheet.getConditionalFormatRules();
        rules.push(conditionalRule(sheet.getRange(a1), a1, c.rule));
        sheet.setConditionalFormatRules(rules);
        return 0;
      }
      case "apply_data_validation": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const a1 = body(c.range);
        sheet.getRange(a1).setDataValidation(validationRule(ss2, c, a1, opts.activeSheet));
        return 0;
      }
      case "find_replace": {
        const finders = c.scopeRange ? [sheetOf(ss2, c.scopeSheet, opts.activeSheet).getRange(body(c.scopeRange)).createTextFinder(c.find)] : c.scopeSheet ? [sheetOf(ss2, c.scopeSheet, opts.activeSheet).createTextFinder(c.find)] : ss2.getSheets().filter((sh) => !isCitationLog(sh.getName())).map((sh) => sh.createTextFinder(c.find));
        let replaced = 0;
        for (const f of finders) {
          replaced += f.matchCase((_j = c.matchCase) != null ? _j : false).matchEntireCell((_k = c.wholeWord) != null ? _k : false).matchFormulaText((_l = c.matchFormulas) != null ? _l : false).replaceAllWith(c.replace);
        }
        return replaced;
      }
      case "fill": {
        const sheet = sheetOf(ss2, c.sheet, opts.activeSheet);
        const destA1 = fillDestinationA1(c);
        const src = parseA1(splitSheetRef(c.sourceRange).body);
        const dest = parseA1(destA1);
        const values = sheet.getRange(formatA1(src)).getValues();
        const down = c.direction === "down" || c.direction === "up";
        const lines = down ? src.endCol - src.startCol + 1 : src.endRow - src.startRow + 1;
        const along = down ? src.endRow - src.startRow + 1 : src.endCol - src.startCol + 1;
        const series = Array.from(
          { length: lines },
          (_, j) => {
            var _a2, _b2;
            return fillSeries(c.mode === "series", along === 1 && typeof (down ? (_a2 = values[0]) == null ? void 0 : _a2[j] : (_b2 = values[j]) == null ? void 0 : _b2[0]) === "number");
          }
        );
        if (series.every((s) => s === series[0])) {
          sheet.getRange(formatA1(src)).autoFill(sheet.getRange(destA1), autoFillSeries(series[0]));
        } else {
          series.forEach((s, j) => {
            const line = (r) => down ? { ...r, startCol: src.startCol + j, endCol: src.startCol + j } : { ...r, startRow: src.startRow + j, endRow: src.startRow + j };
            sheet.getRange(formatA1(line(src))).autoFill(sheet.getRange(formatA1(line(dest))), autoFillSeries(s));
          });
        }
        const { rows, cols } = rangeShape(dest);
        return rows * cols;
      }
      case "create_named_range": {
        const ref = namedRangeRef(c);
        const sheet = sheetOf(ss2, ref.sheet, opts.activeSheet);
        ss2.setNamedRange(c.name, sheet.getRange(ref.body));
        return 0;
      }
      case "update_chart": {
        const { sheet, chart } = findChart(ss2, c.sheet, c.chartName, opts.activeSheet);
        const b = chart.modify();
        setChartDisplay(b, c, c.chartType != null ? c.chartType === "bar" : b.getChartType() === chartTypeFor("bar"));
        if (c.sourceRange != null) {
          const src = splitSheetRef(c.sourceRange);
          const dataSheet = src.sheet ? sheetOf(ss2, src.sheet, opts.activeSheet) : sheet;
          const data = dataSheet.getRange(src.body);
          b.clearRanges().addRange(data).setNumHeaders(headerRows(data));
        }
        if (c.chartType != null) {
          b.setChartType(chartTypeFor(c.chartType));
          if (c.chartType === "donut" || c.chartType === "pie") b.setOption("pieHole", c.chartType === "donut" ? 0.4 : 0);
        }
        if (c.anchor != null) {
          const at = parseA1(body(c.anchor));
          b.setPosition(at.startRow + 1, at.startCol + 1, 0, 0);
        }
        sheet.updateChart(b.build());
        return 0;
      }
      case "delete_chart": {
        const { sheet, chart } = findChart(ss2, c.sheet, c.chartName, opts.activeSheet);
        sheet.removeChart(chart);
        return 0;
      }
      case "update_pivot": {
        const pivot = findPivot(ss2, c.pivotName);
        const source = pivot.getSourceDataRange();
        const headers = source.offset(0, 0, 1).getDisplayValues()[0];
        const where = `${quoteSheet(source.getSheet().getName())}!${source.getA1Notation()}`;
        const column = (field) => pivotColumn(headers, field, where, source.getColumn());
        const rows = (_m = c.rowFields) == null ? void 0 : _m.map(column);
        const cols = (_n = c.columnFields) == null ? void 0 : _n.map(column);
        const values = (_o = c.valueFields) == null ? void 0 : _o.map((v) => [column(v.field), summarizeFunction(v.agg)]);
        if (rows) {
          removeEach(() => pivot.getRowGroups());
          for (const i of rows) pivot.addRowGroup(i);
        }
        if (cols) {
          removeEach(() => pivot.getColumnGroups());
          for (const i of cols) pivot.addColumnGroup(i);
        }
        if (values) {
          removeEach(() => pivot.getPivotValues());
          for (const [i, fn] of values) pivot.addPivotValue(i, fn);
        }
        return 0;
      }
      default: {
        const unhandled = c.kind;
        throw new TranslateError(ERROR_CODES.SHEETS_UNSUPPORTED_KIND, `${unhandled} reached applyOne without support`);
      }
    }
  }
  var NETWORK_FUNCTION = /\b(?:IMPORT(?:DATA|XML|HTML|FEED|RANGE)|IMAGE)\s*\(/i;
  function citationLogUse(c) {
    if (namedSheetsOf(c).some(isCitationLog)) return `${CITATION_SHEET_NAME} is Coco's citation log, not a sheet to write to`;
    if (sourceSheetsOf(c).some((n) => n !== null && isCitationLog(n))) {
      return `${CITATION_SHEET_NAME} is Coco's citation log, not a source to read from`;
    }
    return void 0;
  }
  function sourceSheetsOf(c) {
    const of = (range, fallback) => {
      var _a;
      if (!range) return [];
      const named = (_a = splitSheetRef(range.trim().replace(/^=/, "")).sheet) != null ? _a : fallback;
      return [named ? decodeSheetName(named) : null];
    };
    if (c.kind === "create_pivot") return of(c.sourceRange, void 0);
    if (c.kind === "create_chart" || c.kind === "update_chart") return of(c.sourceRange, c.sheet);
    if (c.kind === "apply_data_validation" && c.rule.type === "range") return of(c.rule.sourceRange, c.sheet);
    return [];
  }
  function unsupportedVariant(c) {
    var _a, _b, _c, _d, _e;
    if (c.kind === "create_table" || c.kind === "update_table") return tableVariant(c);
    if (c.kind === "apply_conditional_format" && c.tableName) return "Native table conditional formatting is not available on Sheets.";
    if (c.kind === "apply_conditional_format" && c.rule.type === "custom_formula" && c.rule.format) return "Exact conditional font/fill formatting is not available on Sheets; use highlight.";
    if ((c.kind === "apply_conditional_format" || c.kind === "apply_data_validation") && NETWORK_FUNCTION.test(JSON.stringify(c.rule))) {
      return "a rule's formula may not call IMPORTDATA, IMPORTXML, IMPORTHTML, IMPORTFEED, IMPORTRANGE or IMAGE";
    }
    switch (c.kind) {
      case "clear_range":
        return c.applyTo === "contents" || c.applyTo === "formats" || c.applyTo === "all" ? void 0 : `clear_range applyTo=${c.applyTo} is not expressible on Sheets`;
      case "apply_sort_filter":
        if (c.filter) return "apply_sort_filter with a filter is not implemented on Sheets";
        return c.sort && c.sort.length > 0 ? void 0 : "apply_sort_filter needs a sort on Sheets";
      case "set_cell_format":
        return c.hyperlink ? "Sheets has no cell hyperlink property; setting one would overwrite the value" : notHexColour([(_a = c.font) == null ? void 0 : _a.color, (_b = c.fill) == null ? void 0 : _b.color, (_c = c.border) == null ? void 0 : _c.color]);
      case "apply_conditional_format":
        if (c.rule.type === "data_bar") return "data bars are not available in Google Sheets conditional formatting";
        if (c.rule.type === "icon_set") return "icon sets are not available in Google Sheets conditional formatting";
        return void 0;
      case "apply_data_validation":
        return validationProblem(c.rule);
      case "create_chart":
      case "update_chart":
        if (c.kind === "update_chart" && c.newName != null) return "a Sheets chart has no name to rename; charts are found by id";
        if (c.dataLabels != null) return "chart data labels are not implemented on Sheets";
        return notHexColour((_d = c.seriesColors) != null ? _d : []);
      case "create_pivot":
      case "update_pivot":
        for (const v of (_e = c.valueFields) != null ? _e : []) {
          if (summarizeName(v.agg) === void 0) return `pivot aggregation ${JSON.stringify(v.agg)} is not available on Sheets`;
        }
        return void 0;
      case "create_named_range":
        if (c.scope === "sheet") return "Sheets names are spreadsheet-wide; a sheet-scoped name is not available";
        return namedRangeRef(c) ? void 0 : "a named range on Sheets must refer to cells; named constants and formulas are Excel-only";
      default:
        return void 0;
    }
  }
  function isCitationLog(name) {
    return name.toLowerCase() === CITATION_SHEET_NAME.toLowerCase();
  }
  function namedSheetsOf(c) {
    var _a;
    const names = [];
    if (c.kind === "update_cells") names.push(...c.edits.map((e) => e.sheet));
    else if (c.kind === "create_pivot") names.push(c.destinationSheet);
    else if (c.kind === "find_replace") names.push(c.scopeSheet);
    else if (c.kind === "create_named_range") names.push((_a = namedRangeRef(c)) == null ? void 0 : _a.sheet);
    else if (c.kind === "manage_sheet") names.push(c.sheet, c.newName);
    else names.push(c.sheet);
    return names.filter((n) => !!n).map(decodeSheetName);
  }
  function producedSheets(changes, activeSheet) {
    const out = /* @__PURE__ */ new Map();
    for (const c of changes) {
      if (c.kind === "manage_sheet" && (c.action === "rename" || c.action === "copy") && c.newName) {
        out.set(decodeSheetName(c.newName), decodeSheetName(c.sheet) || activeSheet);
      }
    }
    return out;
  }
  function usesActiveSheet(c) {
    if (c.kind === "update_cells") return c.edits.some((e) => !e.sheet);
    if (c.kind === "find_replace") return !c.scopeSheet && !!c.scopeRange;
    return namedSheetsOf(c).length === 0;
  }
  function sheetsToEnsure(changes) {
    var _a;
    const out = /* @__PURE__ */ new Map();
    const add = (n, position) => {
      if (!n) return;
      const name = decodeSheetName(n);
      if (position !== void 0 || !out.has(name)) out.set(name, position != null ? position : out.get(name));
    };
    for (const c of changes) {
      if (c.kind === "update_cells") for (const e of c.edits) add(e.sheet);
      else if (c.kind === "insert_sheet") add(c.sheet, (_a = c.position) != null ? _a : void 0);
      else if (c.kind === "create_pivot") add(c.destinationSheet);
      else add(c.sheet);
    }
    return out;
  }
  function neededExtents(changes, active) {
    const need = /* @__PURE__ */ new Map();
    const note = (sheetName, a1) => {
      var _a;
      const name = sheetName ? decodeSheetName(sheetName) : active;
      if (!name) return;
      let parsed;
      try {
        parsed = parseA1(a1);
      } catch {
        return;
      }
      const cur = (_a = need.get(name)) != null ? _a : { rows: 0, cols: 0 };
      need.set(name, {
        rows: Math.max(cur.rows, parsed.endRow + 1),
        cols: Math.max(cur.cols, parsed.endCol + 1)
      });
    };
    for (const c of changes) {
      if (c.kind === "update_cells") for (const e of c.edits) note(e.sheet, e.cell);
      else if (c.kind === "create_pivot") note(c.destinationSheet, body(c.destination));
      else if (c.kind === "fill") note(c.sheet, fillDestinationA1(c));
      else {
        const r = c.range;
        if (r) note(c.sheet, r);
      }
    }
    return need;
  }
  function resolveIndices(c) {
    if (c.indices && c.indices.length > 0) return [...new Set(c.indices)].sort((a, b) => a - b);
    if (c.range) {
      const out = /* @__PURE__ */ new Set();
      for (const run of rangeRuns(c)) for (let i = run.start; i <= run.end; i++) out.add(i);
      return [...out].sort((a, b) => a - b);
    }
    throw new TranslateError(
      ERROR_CODES.SHEETS_BAD_CHANGE,
      "manage_rows_cols requires indices or range"
    );
  }
  function contiguousBlocks(sorted) {
    const out = [];
    for (const i of sorted) {
      const last = out[out.length - 1];
      if (last && last.start + last.count === i) last.count++;
      else out.push({ start: i, count: 1 });
    }
    return out;
  }
  function rangeRuns(c) {
    var _a;
    return splitAddresses((_a = c.range) != null ? _a : "").map((piece) => {
      const [head, tail = head] = body(piece).split(":");
      const start = firstAxisIndex(head, c.target);
      const end = firstAxisIndex(tail, c.target);
      if (start < 0 || end < start) {
        throw new TranslateError(ERROR_CODES.SHEETS_BAD_CHANGE, `no ${c.target}s in range ${piece}`);
      }
      return { start, end };
    });
  }
  function rectangular(grid, rows, cols) {
    var _a;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const v = (_a = grid == null ? void 0 : grid[r]) == null ? void 0 : _a[c];
        row.push(v === void 0 ? null : v);
      }
      out.push(row);
    }
    return out;
  }
  function asLiteral(v) {
    return typeof v === "string" && /^[=+]/.test(v) ? `'${v}` : v;
  }
  function withEquals(formula) {
    const t = formula.trim();
    return t.startsWith("=") ? t : `=${t}`;
  }
  function isFormula(entry) {
    return entry.trim().startsWith("=");
  }
  function formulaRuns(formulas) {
    var _a, _b;
    const runs = [];
    const cols = (_b = (_a = formulas[0]) == null ? void 0 : _a.length) != null ? _b : 0;
    for (let col = 0; col < cols; col++) {
      let run = null;
      for (let r = 0; r < formulas.length; r++) {
        const f = formulas[r][col];
        if (f == null || !isFormula(f)) {
          run = null;
          continue;
        }
        if (!run) runs.push(run = { row: r, col, formulas: [] });
        run.formulas.push(f.trim());
      }
    }
    return runs;
  }
  function hexColour(hex) {
    const raw = hex.trim().replace(/^#/, "");
    const expanded = raw.length === 3 ? raw.replace(/./g, (c) => c + c) : raw;
    return /^[0-9a-fA-F]{6}$/.test(expanded) ? `#${expanded.toLowerCase()}` : void 0;
  }
  function notHexColour(colours) {
    const bad = colours.find((c) => c !== void 0 && hexColour(c) === void 0);
    return bad === void 0 ? void 0 : `colours are #RRGGBB hex on Sheets, got ${JSON.stringify(bad)}`;
  }
  function setChartDisplay(b, d, barred) {
    if (d.title != null) b.setOption("title", d.title);
    if (d.axisTitles) {
      if (d.axisTitles.category != null) b.setOption(barred ? "vAxis.title" : "hAxis.title", d.axisTitles.category);
      if (d.axisTitles.value != null) b.setOption(barred ? "hAxis.title" : "vAxis.title", d.axisTitles.value);
    }
    if (d.legend) {
      if (d.legend.visible === false) b.setOption("legend.position", "none");
      else if (d.legend.position != null) b.setOption("legend.position", d.legend.position);
      else if (d.legend.visible === true) b.setOption("legend.position", "right");
    }
    if (d.seriesColors && d.seriesColors.length > 0) {
      b.setOption("series", Object.fromEntries(d.seriesColors.map((color, i) => [i, { color: normaliseHex(color) }])));
    }
    if (typeof d.widthPixels === "number") b.setOption("width", d.widthPixels);
    if (typeof d.heightPixels === "number") b.setOption("height", d.heightPixels);
  }
  function normaliseHex(hex) {
    const h = hexColour(hex);
    if (!h) throw new TranslateError(ERROR_CODES.SHEETS_BAD_CHANGE, `not a hex colour: ${JSON.stringify(hex)}`);
    return h;
  }
  function splitSheetRef(a1) {
    const [sheet, rest] = splitSheetAndRange(a1);
    return { sheet, body: rest };
  }
  var SHEETS_CHART_TYPE = {
    column: "COLUMN",
    bar: "BAR",
    line: "LINE",
    pie: "PIE",
    donut: "PIE",
    scatter: "SCATTER",
    area: "AREA"
  };
  function chartTypeFor(kind) {
    var _a;
    const g = globalThis;
    const types = (_a = g.Charts) == null ? void 0 : _a.ChartType;
    if (!types) throw new TranslateError(ERROR_CODES.SHEETS_UNSUPPORTED_FEATURE, "Charts service unavailable");
    const name = SHEETS_CHART_TYPE[kind];
    const type = name ? types[name] : void 0;
    if (type === void 0) {
      throw new TranslateError(ERROR_CODES.SHEETS_BAD_CHANGE, `unsupported chart type "${kind}"`);
    }
    return type;
  }
  function borderStyle(weight) {
    return enumValue("BorderStyle", { thin: "SOLID", medium: "SOLID_MEDIUM", thick: "SOLID_THICK" }[weight]);
  }
  function appStatics() {
    const app = globalThis.SpreadsheetApp;
    if (!(app == null ? void 0 : app.newConditionalFormatRule)) {
      throw new TranslateError(ERROR_CODES.SHEETS_UNSUPPORTED_FEATURE, "SpreadsheetApp builders unavailable");
    }
    return app;
  }
  function enumValue(group, name) {
    var _a;
    const app = globalThis.SpreadsheetApp;
    const v = (_a = app == null ? void 0 : app[group]) == null ? void 0 : _a[name];
    if (v === void 0) {
      throw new TranslateError(ERROR_CODES.SHEETS_UNSUPPORTED_FEATURE, `SpreadsheetApp.${group}.${name} unavailable`);
    }
    return v;
  }
  function autoFillSeries(name) {
    return enumValue("AutoFillSeries", name);
  }
  function fillSeries(series, loneNumber) {
    return series !== loneNumber ? "DEFAULT_SERIES" : "ALTERNATE_SERIES";
  }
  function anchorCell(a1) {
    const r = parseA1(a1);
    return formatA1({ startRow: r.startRow, startCol: r.startCol, endRow: r.startRow, endCol: r.startCol });
  }
  function absolute(a1) {
    const r = parseA1(a1);
    const cell = (row, col) => `$${colIndexToLetter(col)}$${row + 1}`;
    return `${cell(r.startRow, r.startCol)}:${cell(r.endRow, r.endCol)}`;
  }
  function asNumber(v) {
    if (typeof v === "number") return Number.isFinite(v) ? v : null;
    if (typeof v !== "string" || v.trim() === "") return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  }
  function operand(v) {
    var _a;
    return typeof v === "number" ? String(v) : (_a = cellValueOperand(v)) != null ? _a : `(${v.trim().replace(/^=/, "")})`;
  }
  var CELL_VALUE_SYMBOL = { gt: ">", lt: "<", gte: ">=", lte: "<=", eq: "=", neq: "<>" };
  function conditionalRule(range, a1, rule) {
    const b = appStatics().newConditionalFormatRule().setRanges([range]);
    const cell = anchorCell(a1);
    const all = absolute(a1);
    const highlighted = (x, color) => x.setBackground(HIGHLIGHT_HEX[color]).build();
    switch (rule.type) {
      case "color_scale": {
        const [low, mid, high] = COLOR_SCALE_PALETTES[rule.palette];
        return b.setGradientMinpoint(low).setGradientMidpointWithValue(mid, enumValue("InterpolationType", "PERCENTILE"), "50").setGradientMaxpoint(high).build();
      }
      case "cell_value": {
        const lo = asNumber(rule.value);
        const hi = asNumber(rule.value2);
        if (lo !== null && (hi !== null || rule.operator !== "between" && rule.operator !== "not_between")) {
          const byOperator = {
            gt: () => b.whenNumberGreaterThan(lo),
            lt: () => b.whenNumberLessThan(lo),
            gte: () => b.whenNumberGreaterThanOrEqualTo(lo),
            lte: () => b.whenNumberLessThanOrEqualTo(lo),
            eq: () => b.whenNumberEqualTo(lo),
            neq: () => b.whenNumberNotEqualTo(lo),
            between: () => b.whenNumberBetween(lo, hi),
            not_between: () => b.whenNumberNotBetween(lo, hi)
          };
          return highlighted(byOperator[rule.operator](), rule.highlight);
        }
        const x = operand(rule.value);
        if (rule.operator === "between" || rule.operator === "not_between") {
          if (rule.value2 === void 0) {
            throw new TranslateError(ERROR_CODES.SHEETS_BAD_CHANGE, `${rule.operator} needs value2`);
          }
          const y = operand(rule.value2);
          const f = rule.operator === "between" ? `=AND(${cell}>=${x},${cell}<=${y})` : `=OR(${cell}<${x},${cell}>${y})`;
          return highlighted(b.whenFormulaSatisfied(f), rule.highlight);
        }
        return highlighted(b.whenFormulaSatisfied(`=${cell}${CELL_VALUE_SYMBOL[rule.operator]}${x}`), rule.highlight);
      }
      case "top_bottom": {
        const n = rule.rank;
        const test = rule.asPercent ? rule.direction === "top" ? `${cell}>=PERCENTILE(${all},${1 - n / 100})` : `${cell}<=PERCENTILE(${all},${n / 100})` : rule.direction === "top" ? `${cell}>=LARGE(${all},${n})` : `${cell}<=SMALL(${all},${n})`;
        return highlighted(b.whenFormulaSatisfied(`=AND(ISNUMBER(${cell}),${test})`), rule.highlight);
      }
      case "text": {
        const byOperator = {
          contains: () => b.whenTextContains(rule.text),
          does_not_contain: () => b.whenTextDoesNotContain(rule.text),
          begins_with: () => b.whenTextStartsWith(rule.text),
          ends_with: () => b.whenTextEndsWith(rule.text)
        };
        return highlighted(byOperator[rule.operator](), rule.highlight);
      }
      case "preset_criteria": {
        const byCriterion = {
          blanks: () => b.whenCellEmpty(),
          non_blanks: () => b.whenCellNotEmpty(),
          errors: () => b.whenFormulaSatisfied(`=ISERROR(${cell})`),
          non_errors: () => b.whenFormulaSatisfied(`=NOT(ISERROR(${cell}))`),
          unique: () => b.whenFormulaSatisfied(`=AND(${cell}<>"",COUNTIF(${all},${cell})=1)`),
          duplicate: () => b.whenFormulaSatisfied(`=AND(${cell}<>"",COUNTIF(${all},${cell})>1)`),
          above_average: () => b.whenFormulaSatisfied(`=AND(ISNUMBER(${cell}),${cell}>AVERAGE(${all}))`),
          below_average: () => b.whenFormulaSatisfied(`=AND(ISNUMBER(${cell}),${cell}<AVERAGE(${all}))`)
        };
        return highlighted(byCriterion[rule.criterion](), rule.highlight);
      }
      case "custom_formula":
        return highlighted(b.whenFormulaSatisfied(withEquals(rule.formula)), rule.highlight);
      case "data_bar":
      case "icon_set":
        throw new TranslateError(ERROR_CODES.SHEETS_UNSUPPORTED_FEATURE, `${rule.type} is not available on Sheets`);
    }
  }
  function dateOperand(v) {
    if (typeof v === "number") return String(v);
    const t = v.trim();
    const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
    if (iso) return `DATE(${Number(iso[1])},${Number(iso[2])},${Number(iso[3])})`;
    return t.startsWith("=") ? `(${t.slice(1)})` : null;
  }
  var VALIDATION_TYPES = {
    list: true,
    range: true,
    custom_formula: true,
    decimal: true,
    integer: true,
    date: true,
    text_length: true
  };
  var VALIDATION_OPERATORS = {
    between: true,
    notBetween: true,
    equalTo: true,
    notEqualTo: true,
    greaterThan: true,
    lessThan: true,
    greaterThanOrEqual: true,
    lessThanOrEqual: true
  };
  var own = (o, key) => Object.prototype.hasOwnProperty.call(o, key);
  function validationProblem(rule) {
    var _a;
    if (!own(VALIDATION_TYPES, rule.type)) return `validation type ${JSON.stringify(rule.type)} is not available`;
    const compares = ["decimal", "integer", "date", "text_length"].includes(rule.type);
    if (compares && rule.operator !== void 0 && !own(VALIDATION_OPERATORS, rule.operator)) {
      return `validation operator ${JSON.stringify(rule.operator)} is not one of ${Object.keys(VALIDATION_OPERATORS).join(", ")}`;
    }
    const two = rule.operator === void 0 || rule.operator === "between" || rule.operator === "notBetween";
    if (["decimal", "integer", "date", "text_length"].includes(rule.type) && two && rule.max === void 0) {
      return `a ${rule.type} validation ${(_a = rule.operator) != null ? _a : "between"} needs both min and max`;
    }
    if (rule.type === "date") {
      for (const v of [rule.min, rule.max]) {
        if (v !== void 0 && dateOperand(v) === null) {
          return `date bound ${JSON.stringify(v)} must be yyyy-mm-dd or a formula`;
        }
      }
    }
    return void 0;
  }
  function comparison(x, rule, bound) {
    var _a, _b, _c, _d;
    const two = ((_a = rule.operator) != null ? _a : "between") === "between" || rule.operator === "notBetween";
    const lo = bound((_c = two ? rule.min : (_b = rule.min) != null ? _b : rule.max) != null ? _c : 0);
    const hi = rule.max === void 0 ? "" : bound(rule.max);
    switch ((_d = rule.operator) != null ? _d : "between") {
      case "between":
        return `AND(${x}>=${lo},${x}<=${hi})`;
      case "notBetween":
        return `OR(${x}<${lo},${x}>${hi})`;
      case "equalTo":
        return `${x}=${lo}`;
      case "notEqualTo":
        return `${x}<>${lo}`;
      case "greaterThan":
        return `${x}>${lo}`;
      case "lessThan":
        return `${x}<${lo}`;
      case "greaterThanOrEqual":
        return `${x}>=${lo}`;
      case "lessThanOrEqual":
        return `${x}<=${lo}`;
    }
  }
  function validationRule(ss2, c, a1, active) {
    var _a, _b, _c, _d, _e, _f, _g, _h;
    const b = appStatics().newDataValidation();
    const rule = c.rule;
    const cell = anchorCell(a1);
    switch (rule.type) {
      case "list":
        b.requireValueInList(((_a = rule.values) != null ? _a : []).map(String), true);
        break;
      case "range": {
        const src = splitSheetRef(((_b = rule.sourceRange) != null ? _b : "").trim().replace(/^=/, ""));
        const sheet = sheetOf(ss2, (_c = src.sheet) != null ? _c : c.sheet, active);
        b.requireValueInRange(sheet.getRange(src.body.replace(/\$/g, "")), true);
        break;
      }
      case "custom_formula":
        b.requireFormulaSatisfied(withEquals((_d = rule.formula) != null ? _d : "TRUE"));
        break;
      case "decimal": {
        const op = (_e = rule.operator) != null ? _e : "between";
        const two = op === "between" || op === "notBetween";
        const lo = asNumber((_g = two ? rule.min : (_f = rule.min) != null ? _f : rule.max) != null ? _g : 0);
        const hi = asNumber(rule.max);
        if (lo !== null && (!two || hi !== null)) {
          const byOperator = {
            between: () => b.requireNumberBetween(lo, hi),
            notBetween: () => b.requireNumberNotBetween(lo, hi),
            equalTo: () => b.requireNumberEqualTo(lo),
            notEqualTo: () => b.requireNumberNotEqualTo(lo),
            greaterThan: () => b.requireNumberGreaterThan(lo),
            lessThan: () => b.requireNumberLessThan(lo),
            greaterThanOrEqual: () => b.requireNumberGreaterThanOrEqualTo(lo),
            lessThanOrEqual: () => b.requireNumberLessThanOrEqualTo(lo)
          };
          byOperator[op]();
        } else {
          b.requireFormulaSatisfied(`=AND(ISNUMBER(${cell}),${comparison(cell, rule, operand)})`);
        }
        break;
      }
      case "integer":
        b.requireFormulaSatisfied(`=AND(ISNUMBER(${cell}),INT(${cell})=${cell},${comparison(cell, rule, operand)})`);
        break;
      case "text_length":
        b.requireFormulaSatisfied(`=${comparison(`LEN(${cell})`, rule, operand)}`);
        break;
      case "date":
        b.requireFormulaSatisfied(
          `=AND(ISDATE(${cell}),${comparison(cell, rule, (v) => dateOperand(v))})`
        );
        break;
      default: {
        const unhandled = rule.type;
        throw new TranslateError(ERROR_CODES.SHEETS_UNSUPPORTED_FEATURE, `validation type ${String(unhandled)}`);
      }
    }
    b.setAllowInvalid(false);
    const help = (_h = c.promptMessage) != null ? _h : c.errorMessage;
    if (help) b.setHelpText(help);
    return b.build();
  }
  function namedRangeRef(c) {
    var _a;
    const ref = splitSheetRef(c.refersTo.trim().replace(/^=/, ""));
    const cells = ref.body.replace(/\$/g, "");
    try {
      parseA1(cells);
    } catch {
      return null;
    }
    return { sheet: (_a = ref.sheet) != null ? _a : c.scopeSheet, body: cells };
  }
  function headerRows(data) {
    const [first = [], second = []] = data.offset(0, 0, 2).getValues();
    return first.some((v, i) => typeof v === "string" && v.trim() !== "" && typeof second[i] === "number") ? 1 : 0;
  }
  function chartHandle(id) {
    return `Chart ${id != null ? id : "?"}`;
  }
  function chartTitle(chart) {
    const t = chart.getOptions().get("title");
    return typeof t === "string" ? t : "";
  }
  function findChart(ss2, sheetName, name, active) {
    const id = /^(?:Chart\s+)?(\d+)$/i.exec(name.trim());
    if (id) {
      for (const sheet2 of ss2.getSheets()) {
        const chart = sheet2.getCharts().find((ch) => ch.getChartId() === Number(id[1]));
        if (chart) return { sheet: sheet2, chart };
      }
    }
    const sheet = sheetOf(ss2, sheetName, active);
    const charts = sheet.getCharts();
    const titled = charts.filter((ch) => chartTitle(ch) === name.trim());
    if (titled.length === 1) return { sheet, chart: titled[0] };
    const known = charts.map((ch) => `${chartHandle(ch.getChartId())}${chartTitle(ch) ? ` ("${chartTitle(ch)}")` : ""}`);
    throw new TranslateError(
      ERROR_CODES.SHEETS_BAD_CHANGE,
      titled.length > 1 ? `${titled.length} charts on ${sheet.getName()} are titled ${JSON.stringify(name)}; name one by its id: ${known.join(", ")}` : `no chart ${JSON.stringify(name)} on ${sheet.getName()}. Charts there: ${known.join(", ") || "none"}`
    );
  }
  function pivotAnchor(p) {
    const cell = p.getAnchorCell();
    return `${quoteSheet(cell.getSheet().getName())}!${cell.getA1Notation()}`;
  }
  function removeEach(list) {
    for (let parts = list(), left = parts.length; parts.length > 0 && left > 0; parts = list(), left--) parts[0].remove();
  }
  function findPivot(ss2, name) {
    const pivots = ss2.getSheets().flatMap((sh) => sh.getPivotTables());
    const want = name.trim().replace(/\$/g, "");
    const ref = splitSheetRef(want);
    const hits = pivots.filter((p) => {
      const cell = p.getAnchorCell();
      if (cell.getA1Notation() !== ref.body.toUpperCase()) return false;
      return !ref.sheet || cell.getSheet().getName() === decodeSheetName(ref.sheet);
    });
    if (hits.length === 1) return hits[0];
    if (hits.length === 0 && pivots.length === 1 && !/^[A-Z]{1,3}[0-9]+$/i.test(ref.body)) return pivots[0];
    throw new TranslateError(
      ERROR_CODES.SHEETS_BAD_CHANGE,
      `no single pivot at ${JSON.stringify(name)}; a Sheets pivot is named by its anchor cell. Pivots: ${pivots.map(pivotAnchor).join(", ") || "none"}`
    );
  }
  function pivotColumn(headers, field, where, firstColumn) {
    const trimmed = headers.map((h) => h.trim());
    const i = trimmed.findIndex((h) => h.toLowerCase() === field.trim().toLowerCase());
    if (i < 0) {
      throw new TranslateError(
        ERROR_CODES.SHEETS_BAD_CHANGE,
        `pivot field "${field}" is not a header in ${where}. Headers: ${trimmed.join(", ")}`
      );
    }
    return firstColumn + i;
  }
  function summarizeName(agg) {
    const names = {
      sum: "SUM",
      count: "COUNTA",
      countnumbers: "COUNT",
      average: "AVERAGE",
      max: "MAX",
      min: "MIN",
      product: "PRODUCT",
      standarddeviation: "STDEV",
      standarddeviationp: "STDEVP",
      variance: "VAR",
      variancep: "VARP"
    };
    const key = (agg != null ? agg : "sum").trim().toLowerCase().replace(/[\s_]/g, "");
    return own(names, key) ? names[key] : void 0;
  }
  function summarizeFunction(agg) {
    const name = summarizeName(agg);
    if (!name) throw new TranslateError(ERROR_CODES.SHEETS_UNSUPPORTED_FEATURE, `pivot aggregation ${JSON.stringify(agg)} is not available on Sheets`);
    return enumValue("PivotTableSummarizeFunction", name);
  }

  // src/host/sheets/citation-sheet.ts
  var COLS = CITATION_HEADERS.length;
  var NEW_SHEET_ROWS = 1e3;
  function findCitationLog(ss2) {
    const sheet = ss2.getSheetByName(CITATION_SHEET_NAME);
    if (!sheet) return { sheet, nextRow: 2, maxRows: NEW_SHEET_ROWS };
    return { sheet, nextRow: Math.max(sheet.getLastRow() + 1, 2), maxRows: sheet.getMaxRows() };
  }
  function appendCitationRows(ss2, cursor, rows, activeSheet, restore) {
    if (rows.length === 0) return 0;
    const filled = rows.map((r) => {
      const row = r.slice(0, COLS);
      if (!row[CITATION_SHEET_COLUMN]) row[CITATION_SHEET_COLUMN] = activeSheet;
      return row.map(asText);
    });
    let sheet = cursor.sheet;
    if (!sheet) {
      sheet = ss2.insertSheet(CITATION_SHEET_NAME);
      sheet.getRange(1, 1, 1, COLS).setValues([[...CITATION_HEADERS]]);
      sheet.hideSheet();
      try {
        if (restore && !restore.isSheetHidden()) restore.activate();
      } catch {
      }
    }
    const lastRow = cursor.nextRow + filled.length - 1;
    if (lastRow > cursor.maxRows) sheet.insertRowsAfter(cursor.maxRows, lastRow - cursor.maxRows);
    sheet.getRange(cursor.nextRow, 1, filled.length, COLS).setValues(filled);
    return filled.length;
  }
  function asText(v) {
    return typeof v === "string" && v !== "" ? `'${v}` : v;
  }
  function readCitationRows(ss2) {
    const { sheet, nextRow } = findCitationLog(ss2);
    if (!sheet || nextRow <= 2) return [];
    const range = sheet.getRange(2, 1, nextRow - 2, COLS);
    return wireSafe(range.getValues(), () => range.getDisplayValues());
  }

  // src/host/sheets/server-bridge.ts
  function tableSummaries() {
    try {
      const doc = activeSpreadsheet();
      return listTables(doc.getId()).map((table) => {
        var _a;
        return {
          name: table.name,
          sheet: table.sheet,
          range: table.address,
          columnNames: ((_a = table.columnProperties) != null ? _a : []).map((column) => column.columnName),
          ...table.tableId.startsWith(NAMED_TABLE_PREFIX) ? { namedRange: true } : {}
        };
      });
    } catch {
      return void 0;
    }
  }
  function cocoBridge_tables() {
    try {
      const doc = activeSpreadsheet();
      return ok(listTables(doc.getId()));
    } catch (error) {
      return fail(error);
    }
  }
  function cocoBridge_connections(action, value, expectedSql) {
    try {
      if (action === "upsert" && value && typeof value === "object") saveConnection(value, expectedSql != null ? expectedSql : void 0);
      else if (action === "remove" && typeof value === "string") removeConnection(value);
      else if (action !== "list") throw new Error("Invalid connection operation");
      return ok(readConnections());
    } catch (error) {
      return fail(error);
    }
  }
  var SHEETS_MAX_CELLS = 2e7;
  var CITATION_LOG_RESERVE_CELLS = 1e3;
  function measureCapacity(doc, sheet, address, trim) {
    var _a;
    const { name, target, a1 } = userRange(doc, sheet, address);
    const range = parseA1(a1);
    const allocated = doc.getSheets().reduce((sum, item) => sum + item.getMaxRows() * item.getMaxColumns(), 0);
    const columns = target.getMaxColumns();
    const others = allocated - target.getMaxRows() * columns;
    const rows = Math.max(target.getMaxRows(), range.endRow + 1);
    const keepColumns = Math.max(target.getLastColumn(), range.endCol + 1);
    const emptyColumns = Math.max(0, columns - keepColumns);
    if (trim && emptyColumns > 0) (_a = doc.getSheetByName(name)) == null ? void 0 : _a.deleteColumns(keepColumns + 1, emptyColumns);
    return {
      requiredCells: others + rows * Math.max(columns, range.endCol + 1),
      trimmedCells: others + rows * keepColumns,
      emptyColumns
    };
  }
  function cocoBridge_refreshTable(change, values, citationRows, lastRefreshed) {
    let table;
    try {
      const doc = activeSpreadsheet();
      table = resizeTable(doc.getId(), change, (sheet, address) => {
        const capacity = measureCapacity(doc, sheet, address);
        if (capacity.requiredCells + CITATION_LOG_RESERVE_CELLS > SHEETS_MAX_CELLS) {
          if (capacity.trimmedCells + CITATION_LOG_RESERVE_CELLS > SHEETS_MAX_CELLS) {
            throw new Error(`This import needs ${capacity.trimmedCells.toLocaleString("en-US")} cells across the spreadsheet, more than Google Sheets allows (${SHEETS_MAX_CELLS.toLocaleString("en-US")}). Import fewer rows or columns, or into another spreadsheet. Nothing was written.`);
          }
          measureCapacity(doc, sheet, address, true);
        }
      });
    } catch (error) {
      return fail(error);
    }
    const at = parseA1(table.address);
    const rows = citationRows.map(
      (row) => Array.isArray(row) && row.length >= 4 ? [row[0], row[1], table.sheet, table.address, ...row.slice(4)] : row
    );
    const written = cocoBridge_applyChanges(
      [{ kind: "write_range", sheet: table.sheet, range: formatA1({ ...at, endRow: at.startRow + values.length - 1 }), before: [], after: values, citations: [] }],
      rows
    );
    if (!written.ok) return ok({ table, writeError: written.error });
    if (lastRefreshed) {
      try {
        saveConnection({ ...change.connection, tableId: table.tableId, tableName: table.name, sheet: table.sheet, lastRefreshed }, change.expectedSql);
      } catch (error) {
        return ok({ table, ...written.value, saveError: describeError(error) });
      }
      return ok({ table, ...written.value, saved: true });
    }
    return ok({ table, ...written.value });
  }
  function ss() {
    const g = globalThis;
    if (!g.SpreadsheetApp) throw new Error("SpreadsheetApp is unavailable");
    return g.SpreadsheetApp;
  }
  function activeSpreadsheet() {
    const doc = ss().getActiveSpreadsheet();
    if (!doc) {
      throw new Error(
        "no active spreadsheet \u2014 the bridge must run from a container-bound script"
      );
    }
    return doc;
  }
  function utilities() {
    const g = globalThis;
    if (!g.Utilities) throw new Error("Utilities is unavailable");
    return g.Utilities;
  }
  function dateFormatter(doc) {
    const zone = doc.getSpreadsheetTimeZone();
    const u = utilities();
    return (d, pattern) => u.formatDate(d, zone, pattern);
  }
  function refuseCitationLog(name) {
    if (isCitationLog(name)) {
      throw new Error(`${ERROR_CODES.PROTECTED_SHEET}: '${name}' is Coco's citation log, not the user's data`);
    }
  }
  function userSheet(doc, sheet) {
    const name = sheet ? decodeSheetName(sheet) : activeSheetName();
    refuseCitationLog(name);
    const target = doc.getSheetByName(name);
    if (!target) throw new Error(`no sheet named ${name}`);
    return { name, target };
  }
  function userRange(doc, sheet, a1) {
    const [prefix, body2] = splitSheetAndRange(a1);
    return { ...userSheet(doc, prefix != null ? prefix : sheet), a1: body2 };
  }
  function userSheets(doc) {
    return doc.getSheets().filter((s) => !isCitationLog(s.getName()));
  }
  function readDoc() {
    return activeSpreadsheet();
  }
  function activeSheetName() {
    const doc = readDoc();
    const sheet = doc.getActiveSheet();
    if (sheet) return sheet.getName();
    const first = doc.getSheets()[0];
    if (!first) throw new Error("spreadsheet has no sheets");
    return first.getName();
  }
  function logOf(doc) {
    return doc;
  }
  function describeError(err) {
    return err instanceof Error ? err.message : String(err);
  }
  function ok(value) {
    return { ok: true, value };
  }
  function fail(err) {
    return { ok: false, error: describeError(err) };
  }
  function cocoBridge_applyChanges(changes, citationRows, landOrTable = true, maybeTable) {
    try {
      const land = typeof landOrTable === "boolean" ? landOrTable : true;
      const table = typeof landOrTable === "object" && landOrTable !== null ? landOrTable : maybeTable;
      const doc = activeSpreadsheet();
      if (table) checkTableWrite(doc.getId(), table);
      const active = activeSheetName();
      const rows = citationRows != null ? citationRows : [];
      const log = logOf(doc);
      let cursor = null;
      let citationsError;
      if (rows.length > 0) {
        try {
          cursor = findCitationLog(log);
        } catch (err) {
          citationsError = describeError(err);
        }
      }
      const res = applyViaSpreadsheetApp(doc, changes, { activeSheet: active, land });
      if (cursor) {
        try {
          appendCitationRows(log, cursor, rows, active, log.getActiveSheet());
        } catch (err) {
          citationsError = describeError(err);
        }
      }
      ss().flush();
      return ok({
        cellsWritten: res.cellsWritten,
        activeSheet: active,
        ...citationsError ? { citationsError } : {}
      });
    } catch (err) {
      return fail(err);
    }
  }
  function cocoBridge_importCapacity(sheet, address, trim) {
    try {
      return ok(measureCapacity(activeSpreadsheet(), sheet, address, trim));
    } catch (error) {
      return fail(error);
    }
  }
  function cocoBridge_readRange(sheet, a1) {
    try {
      const { target, a1: body2 } = userRange(readDoc(), sheet, a1);
      const range = target.getRange(body2);
      const values = wireSafe(range.getValues(), () => range.getDisplayValues());
      return ok({ values });
    } catch (err) {
      return fail(err);
    }
  }
  function cocoBridge_readCells(targets) {
    try {
      const doc = readDoc();
      const values = [];
      const formulas = [];
      for (const t of targets) {
        const { target, a1 } = userRange(doc, t.sheet, t.cell);
        const cell = target.getRange(a1);
        values.push(wireSafe([[cell.getValue()]], () => [[cell.getDisplayValue()]])[0][0]);
        formulas.push(cell.getFormula());
      }
      return ok({ values, formulas });
    } catch (err) {
      return fail(err);
    }
  }
  function cocoBridge_readForModel(sheet, a1, formulas) {
    try {
      const doc = readDoc();
      const { name, target, a1: body2 } = userRange(doc, sheet, a1);
      const want = parseA1(body2);
      const shape = rangeShape(want);
      const rows = Math.min(want.endRow + 1, target.getMaxRows()) - want.startRow;
      const cols = Math.min(want.endCol + 1, target.getMaxColumns()) - want.startCol;
      const range = rows > 0 && cols > 0 ? target.getRange(want.startRow + 1, want.startCol + 1, rows, cols) : null;
      const padded = (grid, blank) => Array.from({ length: shape.rows }, (_, r) => Array.from({ length: shape.cols }, (_2, c) => {
        var _a, _b;
        return (_b = (_a = grid[r]) == null ? void 0 : _a[c]) != null ? _b : blank;
      }));
      const merged = range ? range.getMergedRanges().map((r) => `${quoteSheet(name)}!${r.getA1Notation()}`) : [];
      const values = range ? modelValues(range.getValues(), range.getNumberFormats(), () => range.getDisplayValues(), dateFormatter(doc)) : [];
      return ok({
        sheet: name,
        values: padded(values, ""),
        formulas: formulas ? padded(range ? range.getFormulas() : [], "") : void 0,
        rowCount: shape.rows,
        columnCount: shape.cols,
        mergedAreas: merged.length > 0 ? merged : void 0
      });
    } catch (err) {
      return fail(err);
    }
  }
  function cocoBridge_readSelection() {
    var _a, _b;
    try {
      const doc = readDoc();
      const format = dateFormatter(doc);
      const areaOf = (r) => {
        const sheet = r.getSheet().getName();
        refuseCitationLog(sheet);
        return {
          sheet,
          address: `${quoteSheet(sheet)}!${r.getA1Notation()}`,
          values: modelValues(r.getValues(), r.getNumberFormats(), () => r.getDisplayValues(), format),
          rowCount: r.getNumRows(),
          columnCount: r.getNumColumns()
        };
      };
      const ranges = (_b = (_a = doc.getActiveRangeList()) == null ? void 0 : _a.getRanges()) != null ? _b : [];
      if (ranges.length > 0) {
        const areas = ranges.map(areaOf);
        return ok({ address: areas.map((a) => a.address).join(","), areas, fallback: false });
      }
      const cell = doc.getCurrentCell();
      if (!cell) throw new Error("nothing is selected");
      const area = areaOf(cell);
      return ok({ address: area.address, areas: [area], fallback: true });
    } catch (err) {
      return fail(err);
    }
  }
  function cocoBridge_select(sheet, a1) {
    try {
      const { target, a1: body2 } = userRange(readDoc(), sheet, a1);
      if (body2.includes(",")) target.getRangeList(body2.split(",").map((a) => a.trim())).activate();
      else target.getRange(body2).activate();
      return ok(null);
    } catch (err) {
      return fail(err);
    }
  }
  function cocoBridge_listCharts(sheet) {
    try {
      const doc = readDoc();
      let sheets;
      if (sheet) {
        const name = decodeSheetName(sheet);
        refuseCitationLog(name);
        const found = doc.getSheetByName(name);
        if (!found) {
          throw new Error(
            `${ERROR_CODES.SHEET_NOT_FOUND}: no worksheet named '${sheet}'. Use a sheet name from the outline, or omit \`sheet\` to list charts workbook-wide.`
          );
        }
        sheets = [found];
      } else {
        sheets = userSheets(doc);
      }
      return ok(
        sheets.flatMap(
          (s) => s.getCharts().map((ch) => {
            const title = ch.getOptions().get("title");
            return { sheet: s.getName(), name: chartHandle(ch.getChartId()), title: typeof title === "string" && title ? title : void 0 };
          })
        )
      );
    } catch (err) {
      return fail(err);
    }
  }
  function pivotsOf(sheet) {
    try {
      const name = sheet.getName();
      return sheet.getPivotTables().map((p) => {
        const src = p.getSourceDataRange();
        return {
          sheet: name,
          anchor: `${quoteSheet(name)}!${p.getAnchorCell().getA1Notation()}`,
          source: `${quoteSheet(src.getSheet().getName())}!${src.getA1Notation()}`
        };
      });
    } catch {
      return [];
    }
  }
  function cocoBridge_lastRow(sheet) {
    try {
      const { name, target } = userSheet(readDoc(), sheet);
      return ok({ sheet: name, lastRow: target.getLastRow() });
    } catch (err) {
      return fail(err);
    }
  }
  function cocoBridge_locale() {
    try {
      return ok({ locale: readDoc().getSpreadsheetLocale() });
    } catch (err) {
      return fail(err);
    }
  }
  function selectionAreasOf(doc) {
    var _a, _b;
    try {
      const areas = (_b = (_a = doc.getActiveRangeList()) == null ? void 0 : _a.getRanges()) != null ? _b : [];
      return areas.length === 0 || isCitationLog(areas[0].getSheet().getName()) ? [] : areas;
    } catch {
      return [];
    }
  }
  function selectionScopeOf(doc) {
    const areas = selectionAreasOf(doc);
    if (areas.length === 0) return {};
    return {
      selectionA1: areas.map((r) => `${quoteSheet(r.getSheet().getName())}!${r.getA1Notation()}`).join(",")
    };
  }
  function selectionOf(doc) {
    try {
      const areas = selectionAreasOf(doc);
      if (areas.length === 0) return {};
      const selectionA1 = areas.map((r) => `${quoteSheet(r.getSheet().getName())}!${r.getA1Notation()}`).join(",");
      const first = areas[0];
      const corner = first.getSheet().getRange(first.getRow(), first.getColumn(), Math.min(5, first.getNumRows()), Math.min(5, first.getNumColumns()));
      const values = wireSafe(corner.getValues(), () => corner.getDisplayValues());
      const preview = JSON.stringify(values).slice(0, 2048);
      const extra = areas.length - 1;
      return {
        selectionA1,
        selectionValuesPreview: extra > 0 ? `${preview} (+${extra} more area${extra === 1 ? "" : "s"})` : preview
      };
    } catch {
      return {};
    }
  }
  function cocoBridge_readCitations() {
    try {
      return ok(readCitationRows(logOf(activeSpreadsheet())));
    } catch (err) {
      return fail(err);
    }
  }
  function cocoBridge_outline(scopeOnly) {
    try {
      const doc = readDoc();
      if (scopeOnly) {
        let connections;
        let nativeTables;
        try {
          connections = readConnections();
          if (connections.length > 0) {
            nativeTables = listTables(activeSpreadsheet().getId());
            connections = readConnections();
          }
        } catch {
        }
        return ok({
          activeSheet: activeSheetName(),
          ...selectionScopeOf(doc),
          sheets: [],
          ...connections ? { connections } : {},
          ...nativeTables ? { nativeTables } : {}
        });
      }
      const sheets = userSheets(doc);
      const pivots = sheets.flatMap(pivotsOf);
      const tables = tableSummaries();
      return ok({
        ...pivots.length > 0 ? { pivots } : {},
        ...tables ? { tables } : {},
        activeSheet: activeSheetName(),
        ...selectionOf(doc),
        sheets: sheets.map((s) => {
          var _a;
          const rows = s.getLastRow();
          const cols = s.getLastColumn();
          const firstRow = rows > 0 && cols > 0 ? (_a = s.getRange(`A1:${colIndexToLetter(Math.min(cols, 20) - 1)}1`).getDisplayValues()[0]) != null ? _a : [] : [];
          return { name: s.getName(), rows, cols, firstRow, hidden: s.isSheetHidden(), gid: s.getSheetId() };
        })
      });
    } catch (err) {
      return fail(err);
    }
  }

  // src/sheets/server.ts
  var BRIDGE = {
    cocoBridge_tables,
    cocoBridge_connections,
    cocoBridge_refreshTable,
    cocoBridge_importCapacity,
    cocoBridge_applyChanges,
    cocoBridge_readRange,
    cocoBridge_readCells,
    cocoBridge_readCitations,
    cocoBridge_outline,
    cocoBridge_readForModel,
    cocoBridge_readSelection,
    cocoBridge_select,
    cocoBridge_locale,
    cocoBridge_listCharts,
    cocoBridge_lastRow
  };
  function cocoCall(fn, args) {
    if (typeof fn !== "string" || !Object.prototype.hasOwnProperty.call(BRIDGE, fn)) {
      return { ok: false, error: `${ERROR_CODES.SHEETS_NO_SERVER_FUNCTION}: Coco's server has no ${String(fn)}` };
    }
    return BRIDGE[fn](...Array.isArray(args) ? args : []);
  }
  function spreadsheetApp() {
    const g = globalThis;
    if (!g.SpreadsheetApp) throw new Error("SpreadsheetApp is unavailable");
    return g.SpreadsheetApp;
  }
  function ui() {
    return spreadsheetApp().getUi();
  }
  function htmlService() {
    const g = globalThis;
    if (!g.HtmlService) throw new Error("HtmlService is unavailable");
    return g.HtmlService;
  }
  function onOpen() {
    ui().createAddonMenu().addItem("Open Coco", "cocoShowSidebar").addToUi();
  }
  function onInstall() {
    onOpen();
  }
  function cocoShowSidebar() {
    var _a, _b;
    const id = ((_b = (_a = spreadsheetApp().getActiveSpreadsheet()) == null ? void 0 : _a.getId()) != null ? _b : "").replace(/[^\w-]/g, "");
    const html = '<!DOCTYPE html>\n<html lang="en" data-pane="https://ai.snowflake.com/static/sheets/releases/prod/pane.html">\n  <head>\n    <meta charset="utf-8">\n    <base target="_top">\n    <style>html,body{margin:0;height:100%;overflow:hidden}iframe{display:block;border:0;width:100%;height:100%}</style>\n  </head>\n  <body>\n    <script>"use strict";(()=>{var C="coco:frame-ready",_="coco:frame-hello",E="coco:frame-call",T="coco:frame-result";var F=7*6e4;function m(e){let{self:i,frame:t,paneOrigin:n,info:o,invoke:s}=e,u=o.functions;i.addEventListener("message",l=>{let d=t.contentWindow;if(!d||l.source!==d||l.origin!==n)return;let r=l.data;if(r?.type===C){d.postMessage({type:_,...o},n);return}if(r?.type!==E)return;let c=a=>d.postMessage({type:T,id:r.id,...a},n);if(typeof r.fn!="string"||!u.includes(r.fn)){c({ok:!1,error:`the shell does not relay ${String(r.fn)}`});return}s(r.fn,Array.isArray(r.args)?r.args:[]).then(a=>c({ok:!0,raw:a}),a=>c({ok:!1,error:a instanceof Error?a.message:String(a)}))})}var h="cocoCall",p={tables:"cocoBridge_tables",connections:"cocoBridge_connections",refreshTable:"cocoBridge_refreshTable",importCapacity:"cocoBridge_importCapacity",applyChanges:"cocoBridge_applyChanges",readRange:"cocoBridge_readRange",readCells:"cocoBridge_readCells",readCitations:"cocoBridge_readCitations",outline:"cocoBridge_outline",readForModel:"cocoBridge_readForModel",readSelection:"cocoBridge_readSelection",select:"cocoBridge_select",locale:"cocoBridge_locale",listCharts:"cocoBridge_listCharts",lastRow:"cocoBridge_lastRow"};function y(){return globalThis.google?.script?.run}function f(e,i){return new Promise((t,n)=>{let o=y();if(!o){n(new Error("google.script.run is unavailable in this frame"));return}let s=o.withSuccessHandler(t).withFailureHandler(l=>n(new Error(`google.script.run ${e} failed: ${A(l)}`))),u=s[e];if(typeof u!="function"){n(new Error(`server function ${e} is not declared in this spreadsheet\'s Code.gs`));return}u.apply(s,i)})}function A(e){if(e instanceof Error)return e.message;if(typeof e=="string")return e;try{return JSON.stringify(e)}catch{return String(e)}}var le={rgbColor:{red:241/255,green:243/255,blue:244/255}},$={rgbColor:{red:248/255,green:249/255,blue:250/255}};var R=["create_table","update_table","write_range","update_cells","apply_formula","clear_range","insert_sheet","manage_sheet","manage_rows_cols","apply_sort_filter","set_cell_format","create_chart","create_pivot","apply_conditional_format","apply_data_validation","find_replace","fill","create_named_range","update_chart","delete_chart","update_pivot"];function P(){let e=document.documentElement.getAttribute("data-pane")??"",i=document.documentElement.getAttribute("data-file"),t=document.createElement("iframe");t.src=i?`${e}#file=${i}`:e,t.title="Snowflake CoCo";let n=new URL(e).hostname==="localhost";t.allow=`clipboard-read; clipboard-write${n?"; loopback-network; local-network-access":""}`,document.body.appendChild(t),m({self:window,frame:t,paneOrigin:new URL(e).origin,info:{functions:Object.values(p),kinds:[...R,"refresh_table"],build:"bb3eb5b"},invoke:(o,s)=>f(h,[o,s])})}P();})();\n</script>\n  </body>\n</html>\n'.replace("<html ", `<html data-file="${id}" `);
    ui().showSidebar(htmlService().createHtmlOutput(html).setTitle("Snowflake CoCo"));
  }
  return __toCommonJS(server_exports);
})();


function onOpen() {
  return __cocoServer.onOpen.apply(null, arguments);
}

function onInstall() {
  return __cocoServer.onInstall.apply(null, arguments);
}

function cocoShowSidebar() {
  return __cocoServer.cocoShowSidebar.apply(null, arguments);
}

function cocoCall() {
  return __cocoServer.cocoCall.apply(null, arguments);
}