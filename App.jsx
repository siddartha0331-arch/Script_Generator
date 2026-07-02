import React, { useState, useCallback, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  Database,
  Download,
  Check,
  ChevronDown,
  FileSpreadsheet,
  AlertCircle,
  Copy,
  CheckCircle2,
  Table2,
  Trash2,
  Pencil,
} from "lucide-react";

// ---------------------------------------------------------------------------
// HELPERS — the actual "engine". Everything else in this file is UI wiring.
// ---------------------------------------------------------------------------

// FIX (change #6): earlier this matched ANY column with "date" in the name
// (DUE_DATE, EFFECTIVE_DATE, etc.) which was wrong — it should only fire for
// CREATED_DATE / MODIFIED_DATE. We normalize the header (strip spaces/
// underscores, uppercase) so "Created_Date", "CREATEDDATE" and "created date"
// all still match, but something like "DUE_DATE" correctly does NOT.
const isAutoTimestampColumn = (header) => {
  const normalized = header.toUpperCase().replace(/[^A-Z]/g, "");
  return (
    (normalized.includes("CREATED") && normalized.includes("DATE")) ||
    (normalized.includes("MODIFIED") && normalized.includes("DATE"))
  );
};

// SQL escapes a literal single quote by doubling it: O'Brien -> 'O''Brien'.
// This is NOT the same rule as JS/Python string escaping — worth remembering.
const escapeSqlString = (value) => String(value).replace(/'/g, "''");

const formatValue = (header, rawValue) => {
  if (isAutoTimestampColumn(header)) return "CURRENT_TIMESTAMP";
  if (rawValue === undefined || rawValue === null || rawValue === "")
    return "NULL";
  if (typeof rawValue === "number") return String(rawValue);
  if (typeof rawValue === "boolean") return rawValue ? "1" : "0";
  return `'${escapeSqlString(rawValue)}'`;
};

// keyColumns = which columns go in the WHERE clause of IF NOT EXISTS.
// This is independent from which columns get inserted — you can check
// uniqueness on a column that isn't even part of the INSERT list.
const buildWhereClause = (row, keyColumns) =>
  keyColumns.map((col) => `${col}=${formatValue(col, row[col])}`).join(" AND ");

// insertHeaders = which columns actually appear in the INSERT column/value
// list. This is the "include in script" checkbox set (change #3).
const buildRowScript = (tableName, insertHeaders, row, keyColumns) => {
  const whereClause = buildWhereClause(row, keyColumns);
  const columnList = insertHeaders.map((h) => `[${h}]`).join(",");
  const valueList = insertHeaders.map((h) => formatValue(h, row[h])).join(",");

  return (
    `IF NOT EXISTS (SELECT 1 FROM dbo.${tableName} WHERE ${whereClause})\n` +
    ` BEGIN \n` +
    ` INSERT dbo.${tableName} (${columnList})\n` +
    ` VALUES (${valueList})\n` +
    ` END \n` 
    
  );
};

const buildTableScript = (sheet, tableName, keyColumns, insertColumns, identityInsert) => {
  const insertHeaders = sheet.headers.filter((h) => insertColumns.has(h));
  const parts = [];

  if (identityInsert) parts.push(`SET IDENTITY_INSERT ${tableName} ON\n`);
  sheet.rows.forEach((row) => {
    parts.push(buildRowScript(tableName, insertHeaders, row, keyColumns));
  });
  if (identityInsert) parts.push(`SET IDENTITY_INSERT ${tableName} OFF `);

  return parts.join("\n");
};

const downloadTextFile = (filename, text) => {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

// ---------------------------------------------------------------------------
// SMALL REUSABLE PIECE: multi-select dropdown (checkbox list inside a
// popover). Used for picking the WHERE/key columns (change #1).
// ---------------------------------------------------------------------------
function MultiSelectDropdown({ label, options, selected, onToggle }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const summary =
    selected.size === 0
      ? "Select column(s)..."
      : Array.from(selected).join(" : ");

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button
        onClick={() => setOpen((o) => !o)}
        style={{
          width: "100%",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "9px 12px",
          borderRadius: 8,
          border: `1px solid ${selected.size > 0 ? "#1E8FE0" : "#25324A"}`,
          background: "#0B1220",
          color: selected.size > 0 ? "#EAF4FF" : "#7C8AA5",
          fontSize: 12.5,
          cursor: "pointer",
          textAlign: "left",
        }}
      >
        <span
          style={{
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap",
            fontFamily: selected.size > 0 ? "monospace" : "inherit",
          }}
        >
          {summary}
        </span>
        <ChevronDown size={14} style={{ flexShrink: 0, marginLeft: 8 }} />
      </button>

      {open && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 6px)",
            left: 0,
            right: 0,
            zIndex: 20,
            background: "#111C33",
            border: "1px solid #25324A",
            borderRadius: 10,
            padding: 8,
            maxHeight: 220,
            overflowY: "auto",
            boxShadow: "0 12px 28px rgba(0,0,0,0.4)",
          }}
        >
          {options.map((opt) => {
            const checked = selected.has(opt);
            return (
              <label
                key={opt}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  padding: "7px 8px",
                  borderRadius: 6,
                  fontSize: 12.5,
                  cursor: "pointer",
                  background: checked ? "rgba(30,143,224,0.12)" : "transparent",
                }}
              >
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => onToggle(opt)}
                  style={{ accentColor: "#1E8FE0" }}
                />
                {opt}
              </label>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// MAIN COMPONENT
// ---------------------------------------------------------------------------

export default function SqlScriptGenerator() {
  const [sheets, setSheets] = useState([]); // [{id, originalName, sourceFile, headers, rows}]
  const [fileNames, setFileNames] = useState([]); // multiple files can be loaded now
  const [activeId, setActiveId] = useState(null);
  const [configs, setConfigs] = useState({}); // { [id]: {tableName, keyColumns:Set, insertColumns:Set, identityInsert, generated} }
  const [error, setError] = useState("");
  const [copiedId, setCopiedId] = useState("");
  const [isDragging, setIsDragging] = useState(false);

  // Every parsed file gets a unique batch number so sheet IDs never collide,
  // even if two files share a name or a sheet name (e.g. two workbooks that
  // each have a "Sheet1"). A plain incrementing ref survives across renders
  // without itself causing a re-render the way state would.
  const fileBatchRef = useRef(0);

  const activeSheet = sheets.find((s) => s.id === activeId);
  const activeConfig = activeSheet ? configs[activeSheet.id] : null;

  // The white strip you saw is the page's own <html>/<body> background
  // showing through wherever our component doesn't reach (e.g. below a
  // short upload box). Painting them the same color as our container
  // closes that gap no matter how tall the surrounding page is.
  useEffect(() => {
    const prevBodyBg = document.body.style.background;
    const prevHtmlBg = document.documentElement.style.background;
    document.body.style.background = "#0B1220";
    document.body.style.margin = "0";
    document.documentElement.style.background = "#0B1220";
    document.documentElement.style.height = "100%";
    document.body.style.minHeight = "100vh";
    return () => {
      document.body.style.background = prevBodyBg;
      document.documentElement.style.background = prevHtmlBg;
    };
  }, []);

  // Parses ONE file's ArrayBuffer into sheet objects + their default configs.
  // This is a pure function — it returns data instead of calling setState —
  // which is what lets handleFiles() below process several files and merge
  // them all into state with a single update, instead of one state update
  // per file (which would cause each file's update to briefly overwrite the
  // previous one when they happen close together).
  const parseOneFile = (arrayBuffer, name) => {
    const batch = fileBatchRef.current++;
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const parsedSheets = workbook.SheetNames.map((sheetName, i) => {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
      const headers =
        rows.length > 0
          ? Object.keys(rows[0])
          : XLSX.utils.sheet_to_json(ws, { header: 1 })[0] || [];
      return {
        id: `f${batch}-s${i}-${sheetName}`,
        originalName: sheetName,
        sourceFile: name,
        headers,
        rows,
      };
    }).filter((s) => s.headers.length > 0);

    const newConfigs = {};
    parsedSheets.forEach((s) => {
      newConfigs[s.id] = {
        tableName: s.originalName,
        keyColumns: new Set(),
        insertColumns: new Set(s.headers),
        identityInsert: false,
        generated: null,
      };
    });

    return { parsedSheets, newConfigs };
  };

  // Reads a File as an ArrayBuffer, wrapped in a Promise so we can use
  // await instead of nesting callbacks — much easier to read once you're
  // processing a whole list of files one after another.
  const readFileAsArrayBuffer = (file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = () => reject(new Error(`Couldn't read ${file.name}`));
      reader.readAsArrayBuffer(file);
    });

  // Accepts a FileList (from <input multiple> or a multi-file drag-drop),
  // validates and parses every file, and appends everything onto whatever
  // is already loaded — so uploading a second file doesn't wipe out the
  // first one's sheets.
  const handleFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList || []);
    if (files.length === 0) return;

    const valid = files.filter((f) => /\.(xlsx|xls)$/i.test(f.name));
    const invalidCount = files.length - valid.length;

    if (valid.length === 0) {
      setError("Please upload .xlsx or .xls files.");
      return;
    }

    const allNewSheets = [];
    let mergedConfigs = {};
    const failedFiles = [];

    for (const file of valid) {
      try {
        const buffer = await readFileAsArrayBuffer(file);
        const { parsedSheets, newConfigs } = parseOneFile(buffer, file.name);
        if (parsedSheets.length === 0) {
          failedFiles.push(`${file.name} (no usable sheets — check row 1 has headers)`);
          continue;
        }
        allNewSheets.push(...parsedSheets);
        mergedConfigs = { ...mergedConfigs, ...newConfigs };
      } catch (e) {
        failedFiles.push(file.name);
      }
    }

    if (allNewSheets.length > 0) {
      setSheets((prev) => [...prev, ...allNewSheets]);
      setConfigs((prev) => ({ ...prev, ...mergedConfigs }));
      setFileNames((prev) => [...prev, ...valid.map((f) => f.name)]);
      setActiveId((prev) => prev ?? allNewSheets[0].id);
    }

    if (invalidCount > 0 || failedFiles.length > 0) {
      const parts = [];
      if (invalidCount > 0) parts.push(`${invalidCount} file(s) skipped — not .xlsx/.xls`);
      if (failedFiles.length > 0) parts.push(`Couldn't read: ${failedFiles.join(", ")}`);
      setError(parts.join(". "));
    } else {
      setError("");
    }
  }, []);

  const updateConfig = (id, patch) => {
    setConfigs((prev) => ({ ...prev, [id]: { ...prev[id], ...patch, generated: null } }));
  };

  const toggleKeyColumn = (id, col) => {
    setConfigs((prev) => {
      const set = new Set(prev[id].keyColumns);
      set.has(col) ? set.delete(col) : set.add(col);
      return { ...prev, [id]: { ...prev[id], keyColumns: set, generated: null } };
    });
  };

  const toggleInsertColumn = (id, col) => {
    setConfigs((prev) => {
      const set = new Set(prev[id].insertColumns);
      set.has(col) ? set.delete(col) : set.add(col);
      return { ...prev, [id]: { ...prev[id], insertColumns: set, generated: null } };
    });
  };

  const deleteSheet = (id) => {
    setSheets((prev) => {
      const next = prev.filter((s) => s.id !== id);
      setActiveId((prevActive) =>
        prevActive === id ? (next.length > 0 ? next[0].id : null) : prevActive
      );
      return next;
    });
    setConfigs((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
  };

  const generateForActiveSheet = () => {
    if (!activeSheet) return;
    const cfg = configs[activeSheet.id];
    for (const col of cfg.keyColumns) {
      const hasEmpty = activeSheet.rows.some(
        (row) => row[col] === null || row[col] === undefined || String(row[col]).trim() === ""
      );
      if (hasEmpty) {
        setError(`Error: The column "${col}" selected in your WHERE clause contains empty values. Please clean your data.`);
        return; 
      }
    }
    if (!cfg.tableName.trim()) {
      setError("Table name can't be empty.");
      return;
    }
    if (cfg.keyColumns.size === 0) {
      setError(
        `Pick at least one key column for "${cfg.tableName}" — it's used in the duplicate check (IF NOT EXISTS).`
      );
      return;
    }
    if (cfg.insertColumns.size === 0) {
      setError(`Select at least one column to include in the INSERT for "${cfg.tableName}".`);
      return;
    }
    setError("");
    const script = buildTableScript(
      activeSheet,
      cfg.tableName.trim(),
      Array.from(cfg.keyColumns),
      cfg.insertColumns,
      cfg.identityInsert
    );
    updateConfig(activeSheet.id, {}); // no-op patch keeps generated below intact
    setConfigs((prev) => ({
      ...prev,
      [activeSheet.id]: { ...prev[activeSheet.id], generated: script },
    }));
  };

  const generatedSheets = useMemo(
    () => sheets.filter((s) => configs[s.id]?.generated),
    [sheets, configs]
  );

  const combinedScript = useMemo(() => {
    return generatedSheets
      .map(
        (s) =>
          `-- =====================================================\n-- TABLE: ${configs[s.id].tableName}\n-- =====================================================\n\n${configs[s.id].generated}`
      )
      .join("\n\n");
  }, [generatedSheets, configs]);

  const copyToClipboard = async (text, id) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedId(id);
      setTimeout(() => setCopiedId(""), 1500);
    } catch (e) {
      /* clipboard can fail in sandboxed contexts; export still works */
    }
  };

  const reset = () => {
    setSheets([]);
    setConfigs({});
    setFileNames([]);
    setError("");
    setActiveId(null);
  };

  const panelStyle = {
    background: "#0F1830",
    border: "1px solid #1B2740",
    borderRadius: 14,
    padding: 16,
  };

  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans', -apple-system, BlinkMacSystemFont, sans-serif",
        background: "#0B1220",
        minHeight: "100vh",
        width: "100%",
        boxSizing: "border-box",
        color: "#E5EAF2",
        padding: "28px 20px",
      }}
    >
      <div style={{ maxWidth: 1180, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 6 }}>
          <div
            style={{
              width: 40,
              height: 40,
              borderRadius: 10,
              background: "linear-gradient(135deg, #1E8FE0, #0B5FA5)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              flexShrink: 0,
            }}
          >
            <Database size={20} color="#EAF4FF" />
          </div>
          <div>
            <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, letterSpacing: "-0.01em" }}>
              SQL Script Generator
            </h1>
            <p style={{ fontSize: 13, color: "#7C8AA5", margin: 0 }}>
              Sheet name suggests the table name — you can rename it. Row 1 becomes the columns.
            </p>
          </div>
        </div>

        {/* Upload zone */}
        {sheets.length === 0 && (
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setIsDragging(false);
              handleFiles(e.dataTransfer.files);
            }}
            style={{
              marginTop: 24,
              border: `2px dashed ${isDragging ? "#1E8FE0" : "#25324A"}`,
              borderRadius: 14,
              padding: "56px 24px",
              textAlign: "center",
              background: isDragging ? "rgba(30,143,224,0.06)" : "#0F1830",
              transition: "border-color 120ms ease, background 120ms ease",
            }}
          >
            <Upload size={30} color="#4A5A78" style={{ marginBottom: 12 }} />
            <p style={{ margin: "0 0 4px", fontSize: 15, fontWeight: 600 }}>
              Drop one or more .xlsx files here, or click to browse
            </p>
            <p style={{ margin: "0 0 18px", fontSize: 13, color: "#7C8AA5" }}>
              Every sheet, across every file, becomes a table. First row must be column headers.
            </p>
            <label
              style={{
                display: "inline-block",
                padding: "10px 20px",
                borderRadius: 8,
                background: "#1E8FE0",
                color: "#fff",
                fontSize: 13,
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              Choose file(s)
              <input
                type="file"
                accept=".xlsx,.xls"
                multiple
                onChange={(e) => handleFiles(e.target.files)}
                style={{ display: "none" }}
              />
            </label>
          </div>
        )}

        {error && (
          <div
            style={{
              marginTop: 16,
              padding: "10px 14px",
              borderRadius: 8,
              background: "rgba(224,63,63,0.1)",
              border: "1px solid rgba(224,63,63,0.35)",
              display: "flex",
              gap: 8,
              alignItems: "flex-start",
              fontSize: 13,
              color: "#FF9B9B",
            }}
          >
            <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>{error}</span>
          </div>
        )}

        {/* 3-column working area */}
        {sheets.length > 0 && (
          <div style={{ marginTop: 22 }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: 14,
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#7C8AA5", overflow: "hidden" }}>
                <FileSpreadsheet size={15} style={{ flexShrink: 0 }} />
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {fileNames.join(", ")}
                </span>
                <span style={{ flexShrink: 0 }}>·</span>
                <span style={{ flexShrink: 0 }}>{sheets.length} sheet{sheets.length > 1 ? "s" : ""}</span>
              </div>
              <div style={{ display: "flex", gap: 14, flexShrink: 0 }}>
                <label
                  style={{
                    fontSize: 12,
                    color: "#4FB3FF",
                    cursor: "pointer",
                    textDecoration: "underline",
                  }}
                >
                  Add more files
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    multiple
                    onChange={(e) => handleFiles(e.target.files)}
                    style={{ display: "none" }}
                  />
                </label>
                <button
                  onClick={reset}
                  style={{
                    fontSize: 12,
                    color: "#7C8AA5",
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    textDecoration: "underline",
                  }}
                >
                  Start over
                </button>
              </div>
            </div>

            <div style={{ display: "flex", gap: 16, alignItems: "flex-start", flexWrap: "wrap" }}>
              {/* LEFT: sheet list, grouped by source file once there's more than one */}
              <div style={{ ...panelStyle, width: 240, flexShrink: 0 }}>
                <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#7C8AA5", margin: "0 0 10px" }}>
                  Sheets
                </p>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {fileNames.map((fn, fileIdx) => {
                    const sheetsForFile = sheets.filter((s) => s.sourceFile === fn);
                    if (sheetsForFile.length === 0) return null;
                    return (
                      <div key={fn}>
                        {fileNames.length > 1 && (
                          <p
                            style={{
                              fontSize: 10.5,
                              color: "#4A5A78",
                              margin: fileIdx === 0 ? "0 0 5px" : "12px 0 5px",
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                            }}
                            title={fn}
                          >
                            {fn}
                          </p>
                        )}
                        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                          {sheetsForFile.map((s) => {
                            const cfg = configs[s.id];
                            const isActive = s.id === activeId;
                            return (
                              <div
                                key={s.id}
                                onClick={() => setActiveId(s.id)}
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  justifyContent: "space-between",
                                  padding: "8px 10px",
                                  borderRadius: 8,
                                  border: isActive ? "1px solid #1E8FE0" : "1px solid #1B2740",
                                  background: isActive ? "rgba(30,143,224,0.12)" : "transparent",
                                  cursor: "pointer",
                                }}
                              >
                                <div style={{ display: "flex", alignItems: "center", gap: 6, overflow: "hidden" }}>
                                  <Table2 size={13} style={{ flexShrink: 0 }} />
                                  <span
                                    style={{
                                      fontSize: 12.5,
                                      fontWeight: isActive ? 700 : 500,
                                      color: isActive ? "#EAF4FF" : "#B7C1D6",
                                      overflow: "hidden",
                                      textOverflow: "ellipsis",
                                      whiteSpace: "nowrap",
                                    }}
                                  >
                                    {cfg?.tableName || s.originalName}
                                  </span>
                                  {cfg?.generated && <CheckCircle2 size={12} color="#3BC48B" style={{ flexShrink: 0 }} />}
                                </div>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    deleteSheet(s.id);
                                  }}
                                  title="Remove this table"
                                  style={{
                                    background: "none",
                                    border: "none",
                                    color: "#5A6B8A",
                                    cursor: "pointer",
                                    padding: 2,
                                    flexShrink: 0,
                                  }}
                                >
                                  <Trash2 size={13} />
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 16, flex: "1 1 380px", minWidth: 300 }}>
              {/* MIDDLE: configuration */}
              <div style={{ ...panelStyle, width: "100%" }}>
                {!activeSheet ? (
                  <p style={{ fontSize: 13, color: "#7C8AA5" }}>No sheet selected.</p>
                ) : (
                  <>
                    <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#7C8AA5", margin: "0 0 12px" }}>
                      Configure
                    </p>

                    {/* Editable table name — change #2 */}
                    <label style={{ fontSize: 12, color: "#7C8AA5", display: "block", marginBottom: 6 }}>
                      Table name
                    </label>
                    <div style={{ position: "relative", marginBottom: 16 }}>
                      <Pencil size={12} style={{ position: "absolute", right: 10, top: 11, color: "#5A6B8A" }} />
                      <input
                        type="text"
                        value={activeConfig.tableName}
                        onChange={(e) => updateConfig(activeSheet.id, { tableName: e.target.value })}
                        style={{
                          width: "100%",
                          padding: "9px 30px 9px 12px",
                          borderRadius: 8,
                          border: "1px solid #25324A",
                          background: "#0B1220",
                          color: "#EAF4FF",
                          fontSize: 13,
                          fontFamily: "monospace",
                          boxSizing: "border-box",
                        }}
                      />
                    </div>

                    {/* Key column dropdown — change #1 */}
                    <label style={{ fontSize: 12, color: "#7C8AA5", display: "block", marginBottom: 6 }}>
                      IF NOT EXISTS - Unique Key
                    </label>
                    <div style={{ marginBottom: 16 }}>
                      <MultiSelectDropdown
                        options={activeSheet.headers}
                        selected={activeConfig.keyColumns}
                        onToggle={(col) => toggleKeyColumn(activeSheet.id, col)}
                      />
                    </div>

                    {/* Insert column checkboxes — change #3 */}
                    <label style={{ fontSize: 12, color: "#7C8AA5", display: "block", marginBottom: 6 }}>
                      Columns to include in the INSERT
                    </label>
                    <div
                      style={{
                        display: "grid",
                        gridTemplateColumns: "repeat(auto-fill, minmax(130px, 1fr))",
                        gap: 6,
                        marginBottom: 16,
                        maxHeight: 190,
                        overflowY: "auto",
                        paddingRight: 2,
                      }}
                    >
                      {activeSheet.headers.map((h) => {
                        const checked = activeConfig.insertColumns.has(h);
                        return (
                          <label
                            key={h}
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 6,
                              padding: "6px 8px",
                              borderRadius: 6,
                              border: `1px solid ${checked ? "#1E8FE0" : "#25324A"}`,
                              background: checked ? "rgba(30,143,224,0.08)" : "transparent",
                              fontSize: 11.5,
                              cursor: "pointer",
                            }}
                          >
                            <input
                              type="checkbox"
                              checked={checked}
                              onChange={() => toggleInsertColumn(activeSheet.id, h)}
                              style={{ accentColor: "#1E8FE0" }}
                            />
                            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                              {h}
                            </span>
                          </label>
                        );
                      })}
                    </div>

                    {/* Identity insert */}
                    <label
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 8,
                        fontSize: 13,
                        marginBottom: 18,
                        cursor: "pointer",
                        color: "#D6DCEA",
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={activeConfig.identityInsert}
                        onChange={() =>
                          updateConfig(activeSheet.id, { identityInsert: !activeConfig.identityInsert })
                        }
                        style={{ accentColor: "#1E8FE0" }}
                      />
                      Wrap with SET IDENTITY_INSERT ON / OFF
                    </label>

                    <button
                      onClick={generateForActiveSheet}
                      style={{
                        width: "100%",
                        padding: "10px 18px",
                        borderRadius: 8,
                        background: "#1E8FE0",
                        color: "#fff",
                        fontSize: 13,
                        fontWeight: 700,
                        border: "none",
                        cursor: "pointer",
                      }}
                    >
                      Generate SQL for this table
                    </button>
                  </>
                )}
              </div>

              {/* RIGHT: output + export */}
              <div style={{ ...panelStyle, width: "100%"}}>
                <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: "#7C8AA5", margin: "0 0 12px" }}>
                  Output
                </p>

                {activeConfig?.generated ? (
                  <>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                      <span style={{ fontSize: 12, color: "#7C8AA5", fontFamily: "monospace" }}>
                        {activeConfig.tableName}.sql
                      </span>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button
                          onClick={() => copyToClipboard(activeConfig.generated, activeSheet.id)}
                          style={{
                            fontSize: 12,
                            display: "flex",
                            alignItems: "center",
                            gap: 4,
                            padding: "5px 10px",
                            borderRadius: 6,
                            border: "1px solid #25324A",
                            background: "transparent",
                            color: "#B7C1D6",
                            cursor: "pointer",
                          }}
                        >
                          {copiedId === activeSheet.id ? <Check size={12} /> : <Copy size={12} />}
                          {copiedId === activeSheet.id ? "Copied" : "Copy"}
                        </button>
                        <button
                          onClick={() =>
                            downloadTextFile(`${activeConfig.tableName}.sql`, activeConfig.generated)
                          }
                          style={{
                            fontSize: 12,
                            display: "flex",
                            alignItems: "center",
                            gap: 4,
                            padding: "5px 10px",
                            borderRadius: 6,
                            border: "1px solid #25324A",
                            background: "transparent",
                            color: "#B7C1D6",
                            cursor: "pointer",
                          }}
                        >
                          <Download size={12} /> Export
                        </button>
                      </div>
                    </div>
                    <pre
                      style={{
                        background: "#080D1A",
                        border: "1px solid #1B2740",
                        borderRadius: 10,
                        padding: "14px 14px 20px",
                        fontSize: 11.5,
                        lineHeight: 1.6,
                        color: "#9FE6B8",
                        overflow: "auto",
                        height: 320,
                        minHeight: 140,
                        maxHeight: "70vh",
                        resize: "vertical",
                        fontFamily: "'JetBrains Mono', 'Fira Code', monospace",
                        margin: 0,
                        whiteSpace: "pre",
                      }}
                    >
                      {activeConfig.generated}
                    </pre>
                    <p style={{ fontSize: 10.5, color: "#4A5A78", margin: "6px 0 0", textAlign: "right" }}>
                      Drag the bottom-right corner to resize ↘
                    </p>
                  </>
                ) : (
                  <p style={{ fontSize: 12.5, color: "#5A6B8A" }}>
                    Configure the table on the left, then generate to see the script here.
                  </p>
                )}

                {/* Per-table export list once more than one has been generated */}
                {generatedSheets.length > 0 && (
                  <div style={{ marginTop: 20, borderTop: "1px solid #1B2740", paddingTop: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                      <span style={{ fontSize: 12, color: "#7C8AA5" }}>
                        Generated tables ({generatedSheets.length})
                      </span>
                      <button
                        onClick={() => downloadTextFile("all_tables.sql", combinedScript)}
                        style={{
                          fontSize: 12,
                          display: "flex",
                          alignItems: "center",
                          gap: 5,
                          padding: "6px 12px",
                          borderRadius: 7,
                          background: "#3BC48B",
                          color: "#062615",
                          fontWeight: 700,
                          border: "none",
                          cursor: "pointer",
                        }}
                      >
                        <Download size={12} /> Export all
                      </button>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {generatedSheets.map((s) => (
                        <div
                          key={s.id}
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                            padding: "7px 10px",
                            borderRadius: 7,
                            background: "#0B1220",
                            border: "1px solid #1B2740",
                          }}
                        >
                          <span style={{ fontSize: 12, fontFamily: "monospace", color: "#B7C1D6" }}>
                            {configs[s.id].tableName}.sql
                          </span>
                          <button
                            onClick={() =>
                              downloadTextFile(`${configs[s.id].tableName}.sql`, configs[s.id].generated)
                            }
                            style={{
                              fontSize: 11.5,
                              display: "flex",
                              alignItems: "center",
                              gap: 4,
                              padding: "4px 9px",
                              borderRadius: 6,
                              border: "1px solid #25324A",
                              background: "transparent",
                              color: "#B7C1D6",
                              cursor: "pointer",
                            }}
                          >
                            <Download size={11} /> Export
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
    </div>
  );
}
