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
// THEME — every color used below comes from this one object. Change #2
// (light SharePoint-style palette) lives entirely here: swap this object
// and the whole app re-themes, because nothing below hardcodes a color of
// its own. This is the standard way to avoid a "find every hex code in the
// file" hunt every time a client asks for a color tweak.
// ---------------------------------------------------------------------------
const theme = {
  pageBg: "#F3F5F9",
  cardBg: "#FFFFFF",
  cardBorder: "#E3E7EF",
  cardShadow: "0 1px 2px rgba(16,24,40,0.04), 0 1px 3px rgba(16,24,40,0.06)",
  textPrimary: "#1B2437",
  textSecondary: "#6B7488",
  textMuted: "#9AA3B5",
  brandBlue: "#0B5FA5",
  brandBlueDark: "#084A82",
  brandBlueLight: "#EAF3FC",
  brandBlueBorder: "#BFDBF7",
  accentOrange: "#F2994A",
  accentOrangeDark: "#D9822B",
  success: "#1E9E6A",
  successBg: "#E9F8F1",
  danger: "#D64545",
  dangerBg: "#FDEEEE",
  dangerBorder: "#F6C6C6",
  codeBg: "#F6F8FB",
  codeBorder: "#E1E6EF",
  codeText: "#1F2A44",
};

// ---------------------------------------------------------------------------
// HELPERS — the actual "engine". Everything else in this file is UI wiring.
// ---------------------------------------------------------------------------

const isAutoTimestampColumn = (header) => {
  const normalized = header.toUpperCase().replace(/[^A-Z]/g, "");
  return (
    (normalized.includes("CREATED") && normalized.includes("DATE")) ||
    (normalized.includes("MODIFIED") && normalized.includes("DATE"))
  );
};

const escapeSqlString = (value) => String(value).replace(/'/g, "''");

const formatValue = (header, rawValue) => {
  if (isAutoTimestampColumn(header)) return "CURRENT_TIMESTAMP";
  if (rawValue === undefined || rawValue === null || rawValue === "")
    return "NULL";
  if (typeof rawValue === "number") return String(rawValue);
  if (typeof rawValue === "boolean") return rawValue ? "1" : "0";
  return `'${escapeSqlString(rawValue)}'`;
};

// columnOverrides = { originalHeader: "RENAMED_DB_COLUMN" }. Change #3 feeds
// this in — the WHERE clause needs the *renamed* column too, since that's
// the name that actually exists in the real database table.
const buildWhereClause = (row, keyColumns, columnOverrides = {}) =>
  keyColumns
    .map((col) => `${columnOverrides[col] || col}=${formatValue(col, row[col])}`)
    .join(" AND ");

// insertColumnPairs = [{ original, display }]. "original" is the Excel
// header (used to read the right value out of the row). "display" is
// whatever the user renamed it to for the SQL column list — defaults to
// "original" if they never touched it.
const buildRowScript = (tableName, insertColumnPairs, row, keyColumns, columnOverrides) => {
  const whereClause = buildWhereClause(row, keyColumns, columnOverrides);
  const columnList = insertColumnPairs.map((p) => `[${p.display}]`).join(",");
  const valueList = insertColumnPairs
    .map((p) => formatValue(p.original, row[p.original]))
    .join(",");

  return (
    `IF NOT EXISTS (SELECT 1 FROM dbo.${tableName} WHERE ${whereClause})\n` +
    ` BEGIN \n` +
    ` INSERT dbo.${tableName} (${columnList})\n` +
    ` VALUES (${valueList})\n` +
    ` END \n`
  );
};

const buildTableScript = (sheet, tableName, keyColumns, insertColumns, columnOverrides, identityInsert) => {
  const insertColumnPairs = sheet.headers
    .filter((h) => insertColumns.has(h))
    .map((h) => ({ original: h, display: (columnOverrides[h] || "").trim() || h }));
  const parts = [];

  if (identityInsert) parts.push(`SET IDENTITY_INSERT ${tableName} ON\n`);
  sheet.rows.forEach((row) => {
    parts.push(buildRowScript(tableName, insertColumnPairs, row, keyColumns, columnOverrides));
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
// popover). Used for picking the WHERE/key columns.
// ---------------------------------------------------------------------------
function MultiSelectDropdown({ options, selected, onToggle }) {
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
          border: `1px solid ${selected.size > 0 ? theme.brandBlue : theme.cardBorder}`,
          background: selected.size > 0 ? theme.brandBlueLight : "#FBFCFE",
          color: selected.size > 0 ? theme.brandBlueDark : theme.textSecondary,
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
            background: theme.cardBg,
            border: `1px solid ${theme.cardBorder}`,
            borderRadius: 10,
            padding: 8,
            maxHeight: 220,
            overflowY: "auto",
            boxShadow: "0 12px 28px rgba(16,24,40,0.14)",
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
                  color: theme.textPrimary,
                  background: checked ? theme.brandBlueLight : "transparent",
                }}
              >
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => onToggle(opt)}
                  style={{ accentColor: theme.brandBlue }}
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
  const [fileNames, setFileNames] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [configs, setConfigs] = useState({}); // { [id]: {tableName, keyColumns:Set, insertColumns:Set, columnNameOverrides:{}, identityInsert, generated} }
  const [error, setError] = useState("");
  const [copiedId, setCopiedId] = useState("");
  const [isDragging, setIsDragging] = useState(false);

  // Column-rename state (change #3). Only one column can be "in edit mode"
  // at a time, so a single string is enough instead of a Set.
  const [editingHeader, setEditingHeader] = useState(null);
  const [editValue, setEditValue] = useState("");

  const fileBatchRef = useRef(0);

  const activeSheet = sheets.find((s) => s.id === activeId);
  const activeConfig = activeSheet ? configs[activeSheet.id] : null;

  // Close any in-progress column rename whenever you switch sheets, so you
  // don't end up editing sheet A's column while looking at sheet B.
  useEffect(() => {
    setEditingHeader(null);
  }, [activeId]);

  useEffect(() => {
    const prevBodyBg = document.body.style.background;
    const prevHtmlBg = document.documentElement.style.background;
    document.body.style.background = theme.pageBg;
    document.body.style.margin = "0";
    document.documentElement.style.background = theme.pageBg;
    document.documentElement.style.height = "100%";
    document.body.style.minHeight = "100vh";
    return () => {
      document.body.style.background = prevBodyBg;
      document.documentElement.style.background = prevHtmlBg;
    };
  }, []);

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
        columnNameOverrides: {},
        identityInsert: false,
        generated: null,
      };
    });

    return { parsedSheets, newConfigs };
  };

  const readFileAsArrayBuffer = (file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = () => reject(new Error(`Couldn't read ${file.name}`));
      reader.readAsArrayBuffer(file);
    });

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

  // --- column rename handlers (change #3) ---
  const startEditingColumn = (header) => {
    setEditingHeader(header);
    setEditValue(activeConfig?.columnNameOverrides?.[header] || header);
  };

  const commitColumnRename = (header) => {
    const trimmed = editValue.trim();
    updateConfig(activeSheet.id, {
      columnNameOverrides: {
        ...(activeConfig.columnNameOverrides || {}),
        [header]: trimmed || header,
      },
    });
    setEditingHeader(null);
  };

  const cancelColumnRename = () => setEditingHeader(null);

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
      cfg.columnNameOverrides || {},
      cfg.identityInsert
    );
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
    background: theme.cardBg,
    border: `1px solid ${theme.cardBorder}`,
    borderRadius: 12,
    boxShadow: theme.cardShadow,
    padding: 16,
  };

  const outlineButtonStyle = {
    fontSize: 12,
    display: "flex",
    alignItems: "center",
    gap: 5,
    padding: "6px 11px",
    borderRadius: 7,
    border: `1px solid ${theme.brandBlueBorder}`,
    background: theme.brandBlueLight,
    color: theme.brandBlueDark,
    cursor: "pointer",
    fontWeight: 600,
  };

  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans', -apple-system, BlinkMacSystemFont, sans-serif",
        background: theme.pageBg,
        minHeight: "100vh",
        width: "100%",
        boxSizing: "border-box",
        color: theme.textPrimary,
        padding: "24px 28px",
      }}
    >
      <div style={{ maxWidth: 1440, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 6 }}>
          <div
            style={{
              width: 40,
              height: 40,
              borderRadius: 10,
              background: `linear-gradient(135deg, ${theme.brandBlue}, ${theme.brandBlueDark})`,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              flexShrink: 0,
            }}
          >
            <Database size={20} color="#FFFFFF" />
          </div>
          <div>
            <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, letterSpacing: "-0.01em", color: theme.textPrimary }}>
              SQL Script Generator
            </h1>
            <p style={{ fontSize: 13, color: theme.textSecondary, margin: 0 }}>
              Convert spreadsheet data to database-ready SQL in seconds.
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
              border: `2px dashed ${isDragging ? theme.brandBlue : theme.cardBorder}`,
              borderRadius: 14,
              padding: "56px 24px",
              textAlign: "center",
              background: isDragging ? theme.brandBlueLight : theme.cardBg,
              boxShadow: theme.cardShadow,
              transition: "border-color 120ms ease, background 120ms ease",
            }}
          >
            <Upload size={30} color={theme.brandBlue} style={{ marginBottom: 12 }} />
            <p style={{ margin: "0 0 4px", fontSize: 15, fontWeight: 600, color: theme.textPrimary }}>
              Drop one or more .xlsx files here, or click to browse
            </p>
            <p style={{ margin: "0 0 18px", fontSize: 13, color: theme.textSecondary }}>
              Every sheet, across every file, becomes a table. First row must be column headers.
            </p>
            <label
              style={{
                display: "inline-block",
                padding: "10px 22px",
                borderRadius: 8,
                background: theme.brandBlue,
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
              background: theme.dangerBg,
              border: `1px solid ${theme.dangerBorder}`,
              display: "flex",
              gap: 8,
              alignItems: "flex-start",
              fontSize: 13,
              color: theme.danger,
            }}
          >
            <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>{error}</span>
          </div>
        )}

        {/* 3-column working area: sheets | configure+preview | export */}
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
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: theme.textSecondary, overflow: "hidden" }}>
                <FileSpreadsheet size={15} style={{ flexShrink: 0 }} />
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {fileNames.join(", ")}
                </span>
                <span style={{ flexShrink: 0 }}>·</span>
                <span style={{ flexShrink: 0 }}>{sheets.length} sheet{sheets.length > 1 ? "s" : ""}</span>
              </div>
              <div style={{ display: "flex", gap: 14, flexShrink: 0 }}>
                <label style={{ fontSize: 12, color: theme.brandBlue, cursor: "pointer", fontWeight: 600 }}>
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
                    color: theme.textSecondary,
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

            <div style={{ display: "flex", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
              {/* LEFT: sheet list */}
              <div style={{ ...panelStyle, width: 230, flexShrink: 0 }}>
                <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.textMuted, margin: "0 0 10px", fontWeight: 700 }}>
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
                              color: theme.textMuted,
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
                                  borderLeft: isActive ? `3px solid ${theme.brandBlue}` : "3px solid transparent",
                                  border: isActive ? `1px solid ${theme.brandBlueBorder}` : "1px solid transparent",
                                  borderLeftWidth: 3,
                                  borderLeftColor: isActive ? theme.brandBlue : "transparent",
                                  background: isActive ? theme.brandBlueLight : "transparent",
                                  cursor: "pointer",
                                }}
                              >
                                <div style={{ display: "flex", alignItems: "center", gap: 6, overflow: "hidden" }}>
                                  <Table2 size={13} color={isActive ? theme.brandBlue : theme.textMuted} style={{ flexShrink: 0 }} />
                                  <span
                                    style={{
                                      fontSize: 12.5,
                                      fontWeight: isActive ? 700 : 500,
                                      color: isActive ? theme.brandBlueDark : theme.textPrimary,
                                      overflow: "hidden",
                                      textOverflow: "ellipsis",
                                      whiteSpace: "nowrap",
                                    }}
                                  >
                                    {cfg?.tableName || s.originalName}
                                  </span>
                                  {cfg?.generated && <CheckCircle2 size={12} color={theme.success} style={{ flexShrink: 0 }} />}
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
                                    color: theme.textMuted,
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

              {/* MIDDLE: configure (top) + preview (below) — change #1 */}
              <div style={{ display: "flex", flexDirection: "column", gap: 18, flex: "2 1 480px", minWidth: 340 }}>
                <div style={panelStyle}>
                  {!activeSheet ? (
                    <p style={{ fontSize: 13, color: theme.textSecondary }}>No sheet selected.</p>
                  ) : (
                    <>
                      <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.textMuted, margin: "0 0 12px", fontWeight: 700 }}>
                        Configure
                      </p>

                      {/* Editable table name */}
                      <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                        Table name
                      </label>
                      <div style={{ position: "relative", marginBottom: 16 }}>
                        <Pencil size={12} style={{ position: "absolute", right: 10, top: 11, color: theme.textMuted }} />
                        <input
                          type="text"
                          value={activeConfig.tableName}
                          onChange={(e) => updateConfig(activeSheet.id, { tableName: e.target.value })}
                          style={{
                            width: "100%",
                            padding: "9px 30px 9px 12px",
                            borderRadius: 8,
                            border: `1px solid ${theme.cardBorder}`,
                            background: "#FBFCFE",
                            color: theme.textPrimary,
                            fontSize: 13,
                            fontFamily: "monospace",
                            boxSizing: "border-box",
                          }}
                        />
                      </div>

                      {/* Key column dropdown */}
                      <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                        IF NOT EXISTS - Unique Key
                      </label>
                      <div style={{ marginBottom: 16 }}>
                        <MultiSelectDropdown
                          options={activeSheet.headers}
                          selected={activeConfig.keyColumns}
                          onToggle={(col) => toggleKeyColumn(activeSheet.id, col)}
                        />
                      </div>

                      {/* Insert column checkboxes, now with per-column rename — change #3 */}
                      <label style={{ fontSize: 12, color: theme.textSecondary, display: "block", marginBottom: 6 }}>
                        Columns to include in the INSERT
                        <span style={{ color: theme.textMuted, fontWeight: 400 }}> (pencil renames the DB column)</span>
                      </label>
                      <div
                        style={{
                          display: "grid",
                          gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))",
                          gap: 6,
                          marginBottom: 16,
                          maxHeight: 220,
                          overflowY: "auto",
                          paddingRight: 2,
                        }}
                      >
                        {activeSheet.headers.map((h) => {
                          const checked = activeConfig.insertColumns.has(h);
                          const overrideName = activeConfig.columnNameOverrides?.[h];
                          const isRenamed = overrideName && overrideName !== h;
                          const isEditing = editingHeader === h;
                          return (
                            <div
                              key={h}
                              style={{
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "space-between",
                                gap: 4,
                                padding: "6px 6px 6px 8px",
                                borderRadius: 6,
                                border: `1px solid ${checked ? theme.brandBlue : theme.cardBorder}`,
                                background: checked ? theme.brandBlueLight : "#FBFCFE",
                                fontSize: 11.5,
                              }}
                            >
                              <label
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  gap: 6,
                                  overflow: "hidden",
                                  cursor: "pointer",
                                  flex: 1,
                                  minWidth: 0,
                                }}
                              >
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => toggleInsertColumn(activeSheet.id, h)}
                                  style={{ accentColor: theme.brandBlue, flexShrink: 0 }}
                                />
                                {isEditing ? (
                                  <input
                                    autoFocus
                                    value={editValue}
                                    onClick={(e) => e.preventDefault()}
                                    onChange={(e) => setEditValue(e.target.value)}
                                    onBlur={() => commitColumnRename(h)}
                                    onKeyDown={(e) => {
                                      if (e.key === "Enter") commitColumnRename(h);
                                      if (e.key === "Escape") cancelColumnRename();
                                    }}
                                    style={{
                                      width: "100%",
                                      minWidth: 0,
                                      fontSize: 11.5,
                                      padding: "2px 4px",
                                      borderRadius: 4,
                                      border: `1px solid ${theme.brandBlue}`,
                                      background: "#fff",
                                      color: theme.textPrimary,
                                    }}
                                  />
                                ) : (
                                  <span
                                    title={h}
                                    style={{
                                      overflow: "hidden",
                                      textOverflow: "ellipsis",
                                      whiteSpace: "nowrap",
                                      color: isRenamed ? theme.brandBlueDark : theme.textPrimary,
                                      fontWeight: isRenamed ? 700 : 500,
                                    }}
                                  >
                                    {overrideName || h}
                                  </span>
                                )}
                              </label>
                              {!isEditing && (
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    startEditingColumn(h);
                                  }}
                                  title={`Rename column (Excel: ${h})`}
                                  style={{
                                    background: "none",
                                    border: "none",
                                    color: theme.textMuted,
                                    cursor: "pointer",
                                    padding: 2,
                                    flexShrink: 0,
                                    display: "flex",
                                  }}
                                >
                                  <Pencil size={11} />
                                </button>
                              )}
                            </div>
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
                          color: theme.textPrimary,
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={activeConfig.identityInsert}
                          onChange={() =>
                            updateConfig(activeSheet.id, { identityInsert: !activeConfig.identityInsert })
                          }
                          style={{ accentColor: theme.brandBlue }}
                        />
                        Wrap with SET IDENTITY_INSERT ON / OFF
                      </label>

                      <button
                        onClick={generateForActiveSheet}
                        style={{
                          width: "100%",
                          padding: "10px 18px",
                          borderRadius: 8,
                          background: theme.brandBlue,
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

                {/* Preview only — export controls now live in the right panel (change #1) */}
                <div style={panelStyle}>
                  <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.textMuted, margin: "0 0 12px", fontWeight: 700 }}>
                    Preview
                  </p>
                  {activeConfig?.generated ? (
                    <>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                        <span style={{ fontSize: 12, color: theme.textSecondary, fontFamily: "monospace" }}>
                          {activeConfig.tableName}.sql
                        </span>
                        <button
                          onClick={() => copyToClipboard(activeConfig.generated, activeSheet.id)}
                          style={outlineButtonStyle}
                        >
                          {copiedId === activeSheet.id ? <Check size={12} /> : <Copy size={12} />}
                          {copiedId === activeSheet.id ? "Copied" : "Copy"}
                        </button>
                      </div>
                      <pre
                        style={{
                          background: theme.codeBg,
                          border: `1px solid ${theme.codeBorder}`,
                          borderRadius: 10,
                          padding: "14px 14px 20px",
                          fontSize: 11.5,
                          lineHeight: 1.6,
                          color: theme.codeText,
                          overflow: "auto",
                          height: 340,
                          minHeight: 160,
                          maxHeight: "65vh",
                          resize: "vertical",
                          fontFamily: "'JetBrains Mono', 'Fira Code', monospace",
                          margin: 0,
                          whiteSpace: "pre",
                        }}
                      >
                        {activeConfig.generated}
                      </pre>
                      <p style={{ fontSize: 10.5, color: theme.textMuted, margin: "6px 0 0", textAlign: "right" }}>
                        Drag the bottom-right corner to resize ↘
                      </p>
                    </>
                  ) : (
                    <p style={{ fontSize: 12.5, color: theme.textMuted }}>
                      Configure the table above, then generate to see the script here.
                    </p>
                  )}
                </div>
              </div>

              {/* RIGHT: export — change #1 */}
              <div style={{ ...panelStyle, flex: "1 1 300px", minWidth: 260, maxWidth: 340 }}>
                <p style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: theme.textMuted, margin: "0 0 12px", fontWeight: 700 }}>
                  Export
                </p>

                {generatedSheets.length === 0 ? (
                  <p style={{ fontSize: 12.5, color: theme.textMuted }}>
                    Generate at least one table to enable export.
                  </p>
                ) : (
                  <>
                    <button
                      onClick={() => downloadTextFile("all_tables.sql", combinedScript)}
                      style={{
                        width: "100%",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        gap: 6,
                        padding: "10px 14px",
                        borderRadius: 8,
                        background: theme.accentOrange,
                        color: "#fff",
                        fontWeight: 700,
                        fontSize: 13,
                        border: "none",
                        cursor: "pointer",
                        marginBottom: 14,
                      }}
                    >
                      <Download size={14} /> Export all ({generatedSheets.length})
                    </button>

                    <p style={{ fontSize: 11, color: theme.textMuted, margin: "0 0 8px", fontWeight: 600 }}>
                      Individually
                    </p>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {generatedSheets.map((s) => (
                        <div
                          key={s.id}
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                            padding: "8px 10px",
                            borderRadius: 7,
                            background: "#FBFCFE",
                            border: `1px solid ${theme.cardBorder}`,
                          }}
                        >
                          <span
                            style={{
                              fontSize: 12,
                              fontFamily: "monospace",
                              color: theme.textPrimary,
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {configs[s.id].tableName}.sql
                          </span>
                          <button
                            onClick={() =>
                              downloadTextFile(`${configs[s.id].tableName}.sql`, configs[s.id].generated)
                            }
                            title="Export this table"
                            style={{
                              background: "none",
                              border: "none",
                              color: theme.brandBlue,
                              cursor: "pointer",
                              padding: 2,
                              flexShrink: 0,
                              display: "flex",
                            }}
                          >
                            <Download size={14} />
                          </button>
                        </div>
                      ))}
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
