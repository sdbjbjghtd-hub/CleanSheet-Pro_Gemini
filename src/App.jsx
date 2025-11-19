import React, { useState, useMemo, useEffect, useCallback } from 'react';
import ReactDOM from 'react-dom/client';
import * as XLSX from 'xlsx';

// ==========================================
// 1. UTILITIES & SERVICES
// Description: Pure functions and helpers independent of UI
// ==========================================

/**
 * Simple logger wrapper for consistency and easier debugging.
 */
const Logger = {
    info: (msg, data) => console.log(`[INFO] ${msg}`, data || ''),
    warn: (msg, data) => console.warn(`[WARN] ${msg}`, data || ''),
    error: (msg, err) => console.error(`[ERROR] ${msg}`, err || ''),
};

/**
 * DataService handles all heavy lifting: parsing, logic processing, and exporting.
 */
const DataService = {
    /**
     * Reads a raw File object (CSV/Excel) and returns a structured JSON object.
     * Uses SheetJS (XLSX) library.
     */
    parseExcel: (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = e.target.result;
                    const workbook = XLSX.read(data, { type: 'binary' });
                    const sheetName = workbook.SheetNames[0]; // Default to first sheet
                    const worksheet = workbook.Sheets[sheetName];
                    // Convert sheet to array of arrays (header: 1)
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    if (jsonData.length > 0) {
                        const headers = jsonData[0];
                        // Filter out completely empty rows to keep data clean
                        const rows = jsonData.slice(1).filter(r => r.some(c => c !== null && c !== undefined && c !== ''));
                        Logger.info('File parsed successfully', { rows: rows.length });
                        resolve({ headers, rows, originalFileName: file.name });
                    } else {
                        throw new Error("File appears to be empty");
                    }
                } catch (error) {
                    Logger.error('Parse failed', error);
                    reject(error);
                }
            };
            reader.readAsBinaryString(file);
        });
    },

    /**
     * core logic for comparing a cell value against a rule.
     * @param {any} cellValue - The value in the spreadsheet cell
     * @param {string} operator - 'equals', 'contains', 'greater_than', etc.
     * @param {string} ruleValue - The user-defined value to compare against
     */
    checkCondition: (cellValue, operator, ruleValue) => {
        if (cellValue === undefined || cellValue === null) cellValue = "";
        const valStr = String(cellValue);
        const valLower = valStr.toLowerCase();
        const ruleValLower = String(ruleValue || '').toLowerCase();

        switch (operator) {
            // Text Comparisons
            case 'equals': return valLower === ruleValLower;
            case 'does_not_equal': return valLower !== ruleValLower;
            case 'contains': return valLower.includes(ruleValLower);
            case 'does_not_contain': return !valLower.includes(ruleValLower);
            case 'starts_with': return valLower.startsWith(ruleValLower);
            case 'ends_with': return valLower.endsWith(ruleValLower);
            case 'matches_regex': try { return new RegExp(ruleValue, 'i').test(valStr); } catch(e) { return false; }

            // Numeric Comparisons
            case 'greater_than': return parseFloat(cellValue) > parseFloat(ruleValue);
            case 'less_than': return parseFloat(cellValue) < parseFloat(ruleValue);
            case 'is_number': return !isNaN(parseFloat(cellValue)) && isFinite(cellValue);

            // Type/Content Checks
            case 'is_empty': return valStr.trim() === "";
            case 'is_not_empty': return valStr.trim() !== "";
            case 'text_length_greater': return valStr.length > parseInt(ruleValue || 0);
            case 'text_length_less': return valStr.length < parseInt(ruleValue || 0);
            case 'is_email': return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(valStr);
            case 'is_url': return /^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$/.test(valStr);
            case 'is_date': return !isNaN(Date.parse(valStr));
            default: return false;
        }
    },

    /**
     * Iterates through rows and applies the user-defined rules (Filter/Transform).
     * Returns a NEW array of rows (does not mutate original).
     */
    applyRules: (data, rules) => {
        if (!data) return [];
        // Create a shallow copy of rows to work on
        let currentRows = data.rows.map(r => [...r]);

        rules.forEach(rule => {
            const ifColIdx = data.headers.indexOf(rule.column);
            if (ifColIdx === -1) return; // Skip if column not found

            // Handle Row-Level Actions (Remove/Keep)
            if (['remove', 'keep'].includes(rule.action)) {
                currentRows = currentRows.filter(row => {
                    const match = DataService.checkCondition(row[ifColIdx], rule.operator, rule.value);
                    return rule.action === 'remove' ? !match : match;
                });
            } else {
                // Handle Cell-Level Actions (Set, Append, Upper, etc.)
                const targetColIdx = data.headers.indexOf(rule.targetColumn || rule.column);
                if (targetColIdx !== -1) {
                    currentRows = currentRows.map(row => {
                        const match = DataService.checkCondition(row[ifColIdx], rule.operator, rule.value);
                        if (match) {
                            const newRow = [...row]; // Copy row before mutating
                            let currentVal = newRow[targetColIdx];
                            switch (rule.action) {
                                case 'set': newRow[targetColIdx] = rule.actionValue; break;
                                case 'clear': newRow[targetColIdx] = ""; break;
                                case 'append': newRow[targetColIdx] = String(currentVal || "") + (rule.actionValue || ""); break;
                                case 'prepend': newRow[targetColIdx] = (rule.actionValue || "") + String(currentVal || ""); break;
                                case 'upper': newRow[targetColIdx] = String(currentVal || "").toUpperCase(); break;
                                case 'lower': newRow[targetColIdx] = String(currentVal || "").toLowerCase(); break;
                            }
                            return newRow;
                        }
                        return row;
                    });
                }
            }
        });
        return currentRows;
    },

    /**
     * Generates simple frequency analysis for the "Visualize" tab.
     */
    calculateColumnStats: (rows, colIndex) => {
        const counts = {};
        rows.forEach(row => {
            const val = String(row[colIndex] || '(Empty)');
            counts[val] = (counts[val] || 0) + 1;
        });
        // Sort by frequency and take top 8
        const sorted = Object.entries(counts).sort((a,b) => b[1] - a[1]).slice(0, 8);
        const maxVal = Math.max(...sorted.map(s => s[1]));
        return { sorted, maxVal };
    },

    /**
     * Browser-side file generation and download trigger.
     */
    exportFile: (type, headers, data, fileName) => {
        Logger.info(`Exporting as ${type}`, { fileName });

        if (type === 'json') {
            // Convert array of arrays back to array of objects
            const jsonOutput = data.map(row => { const obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj; });
            const blob = new Blob([JSON.stringify(jsonOutput, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `${fileName}.json`; a.click();
        }
        else if (type === 'sql') {
            // Generate INSERT statements
            const tableName = "data";
            const stmts = data.map(row => {
                const values = row.map(v => typeof v === 'string' ? `'${v.replace(/'/g, "''")}'` : v).join(", ");
                return `INSERT INTO ${tableName} (${headers.map(h => `"${h}"`).join(", ")}) VALUES (${values});`;
            }).join("\n");
            const blob = new Blob([stmts], { type: 'text/plain' });
            const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `${fileName}.sql`; a.click();
        }
        else {
            // Use SheetJS for Excel/CSV
            const wsData = [headers, ...data];
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Data");
            XLSX.writeFile(wb, `${fileName}.${type}`);
        }
    }
};

// ==========================================
// 2. ICONS
// Description: SVG components for UI
// ==========================================
const Icons = {
    FileSpreadsheet: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M8 13h2"/><path d="M8 17h2"/><path d="M14 13h2"/><path d="M14 17h2"/></svg>,
    Upload: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/></svg>,
    RefreshCw: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/></svg>,
    Download: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>,
    Filter: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></svg>,
    Plus: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M5 12h14"/><path d="M12 5v14"/></svg>,
    X: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M18 6 6 18"/><path d="m6 6 12 12"/></svg>,
    AlertCircle: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" x2="12" y1="8" y2="12"/><line x1="12" x2="12.01" y1="16" y2="16"/></svg>,
    Save: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>,
    Scissors: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><line x1="20" y1="4" x2="8.12" y2="15.88"/><line x1="14.47" y1="14.48" x2="20" y2="20"/><line x1="8.12" y1="8.12" x2="12" y2="12"/></svg>,
    Type: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/></svg>,
    Search: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>,
    Copy: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>,
    BarChart2: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>,
    Table: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect width="18" height="18" x="3" y="3" rx="2" ry="2"/><line x1="3" x2="21" y1="9" y2="9"/><line x1="3" x2="21" y1="15" y2="15"/><line x1="12" x2="12" y1="3" y2="21"/></svg>,
    Database: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></svg>,
    FileJson: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M10 12a1 1 0 0 0-1 1v1a1 1 0 0 1-1 1 1 1 0 0 1 1 1v1a1 1 0 0 0 1 1"/></svg>,
    ArrowRight: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>,
    ListFilter: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 6h18"/><path d="M7 12h10"/><path d="M10 18h4"/></svg>,
    GitCompare: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="18" cy="18" r="3"/><circle cx="6" cy="6" r="3"/><path d="M13 6h3a2 2 0 0 1 2 2v7"/><path d="M11 18H8a2 2 0 0 1-2-2V9"/></svg>,
    ArrowLeftRight: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M8 3 4 7l4 4"/><path d="M4 7h16"/><path d="m16 21 4-4-4-4"/><path d="M20 17H4"/></svg>,
    CheckCircle2: () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>
};

// ==========================================
// 3. CUSTOM HOOKS (LOGIC LAYER)
// Description: Encapsulates complex state management
// ==========================================

/**
 * Hooks for managing the Main File (Cleaning/Visualizing tabs).
 * Handles parsing, rule application state, and local storage persistence.
 */
const useFileProcessor = () => {
    const [fileData, setFileData] = useState(null); // Raw file content
    const [rules, setRules] = useState([]); // Array of active filter rules
    const [loading, setLoading] = useState(false); // UI loading state

    // Load saved rules from previous session
    useEffect(() => {
        const sessionRules = localStorage.getItem('cleansheet_active_rules');
        if (sessionRules) {
            try {
                const parsed = JSON.parse(sessionRules);
                if (parsed.length > 0) setRules(parsed);
            } catch(e) { Logger.warn('Failed to load session rules'); }
        }
    }, []);

    // Save rules to session whenever they change
    useEffect(() => {
        localStorage.setItem('cleansheet_active_rules', JSON.stringify(rules));
    }, [rules]);

    const processFile = async (file) => {
        setLoading(true);
        try {
            const data = await DataService.parseExcel(file);
            setFileData(data);
        } catch (e) {
            alert("Error parsing file. Please check the format.");
        } finally {
            setLoading(false);
        }
    };

    // Derived State: processedRows is recalculated only when data or rules change
    const processedRows = useMemo(() => {
        return DataService.applyRules(fileData, rules);
    }, [fileData, rules]);

    // Helper functions for Rules
    const updateRule = (id, field, value) => setRules(rules.map(r => r.id === id ? { ...r, [field]: value } : r));
    const addRule = () => setRules([...rules, { id: Date.now(), column: fileData?.headers[0] || '', operator: 'equals', value: '', action: 'remove', targetColumn: fileData?.headers[0] || '', actionValue: '' }]);
    const deleteRule = (id) => setRules(rules.filter(r => r.id !== id));

    // Direct data mutation (for tools like Trim, Find/Replace, etc.)
    const mutateData = (newRows, newHeaders = null) => {
        setFileData(prev => ({...prev, headers: newHeaders || prev.headers, rows: newRows}));
    };

    return {
        fileData, setFileData, rules, setRules, loading, processedRows,
        processFile, updateRule, addRule, deleteRule, mutateData
    };
};

/**
 * Hooks for managing the Reconciliation (Comparison) Tab.
 * Handles two file inputs (FileA, FileB) and the comparison logic.
 */
const useReconciliation = () => {
    const [state, setState] = useState({ fileA: null, fileB: null, keyA: '', keyB: '', results: null });
    const [activeSubTab, setActiveSubTab] = useState('onlyA'); // Results view toggle

    const handleUpload = async (file, slot) => {
        try {
            const data = await DataService.parseExcel(file);
            setState(prev => ({
                ...prev,
                [slot === 'A' ? 'fileA' : 'fileB']: data,
                // Auto-select first header as the Match Key initially
                [slot === 'A' ? 'keyA' : 'keyB']: data.headers[0] || ''
            }));
        } catch (e) { alert("Error parsing reconciliation file."); }
    };

    const run = () => {
        const { fileA, fileB, keyA, keyB } = state;
        if (!fileA || !fileB || !keyA || !keyB) return;

        const idxA = fileA.headers.indexOf(keyA);
        const idxB = fileB.headers.indexOf(keyB);
        const getNormKey = (val) => String(val || '').trim().toLowerCase(); // Normalize keys

        // Create HashMaps for O(1) lookup time
        const mapA = new Map();
        fileA.rows.forEach(row => mapA.set(getNormKey(row[idxA]), row));
        const mapB = new Map();
        fileB.rows.forEach(row => mapB.set(getNormKey(row[idxB]), row));

        const matches = [];
        const onlyA = [];
        const onlyB = [];

        // Comparison Logic
        mapA.forEach((row, key) => { if (mapB.has(key)) matches.push(row); else onlyA.push(row); });
        mapB.forEach((row, key) => { if (!mapA.has(key)) onlyB.push(row); });

        setState(prev => ({ ...prev, results: { matches, onlyA, onlyB } }));
        // Auto-switch tab to where the data is
        setActiveSubTab(onlyA.length > 0 ? 'onlyA' : (onlyB.length > 0 ? 'onlyB' : 'matches'));
    };

    const reset = () => setState({ fileA: null, fileB: null, keyA: '', keyB: '', results: null });

    return { state, setState, activeSubTab, setActiveSubTab, handleUpload, run, reset };
};

// ==========================================
// 4. UI COMPONENTS
// Description: Functional React Components
// ==========================================

const Header = ({ fileData, activeTab, setActiveTab, onReset, onExport }) => (
    <header className="bg-white border-b border-slate-200 px-4 py-3 flex items-center justify-between shrink-0 z-20">
        <div className="flex items-center gap-3">
            <div className="bg-indigo-600 p-2 rounded text-white"><Icons.FileSpreadsheet /></div>
            <div>
                <h1 className="font-bold text-slate-800 leading-tight">CleanSheet Pro</h1>
                <p className="text-xs text-slate-500">{fileData ? fileData.originalFileName : 'No file loaded'}</p>
            </div>
        </div>
        <div className="flex items-center bg-slate-100 p-1 rounded-lg">
            {['data', 'visualize', 'reconcile'].map(tab => (
                <button key={tab} onClick={() => setActiveTab(tab)}
                    className={`px-4 py-1.5 rounded-md text-sm font-medium transition-all flex items-center gap-2 capitalize ${activeTab === tab ? 'bg-white text-indigo-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}>
                    {tab === 'data' && <Icons.Table />}
                    {tab === 'visualize' && <Icons.BarChart2 />}
                    {tab === 'reconcile' && <Icons.GitCompare />}
                    {tab}
                </button>
            ))}
        </div>
        <div className="flex gap-2">
            <button onClick={onReset} className="text-slate-500 hover:text-red-500 px-3 py-2 text-sm"><Icons.RefreshCw /></button>
            <div className="h-8 w-px bg-slate-200 mx-1"></div>
            {fileData && (
                <>
                    <button onClick={() => onExport('csv')} className="bg-slate-700 text-white px-3 py-2 rounded text-xs font-medium flex gap-1 items-center hover:bg-slate-800"><Icons.Table /> CSV</button>
                    <button onClick={() => onExport('xlsx')} className="bg-green-600 text-white px-3 py-2 rounded text-xs font-medium flex gap-1 items-center hover:bg-green-700"><Icons.FileSpreadsheet /> Excel</button>
                    <button onClick={() => onExport('json')} className="bg-orange-500 text-white px-3 py-2 rounded text-xs font-medium flex gap-1 items-center hover:bg-orange-600"><Icons.FileJson /> JSON</button>
                    <button onClick={() => onExport('sql')} className="bg-blue-500 text-white px-3 py-2 rounded text-xs font-medium flex gap-1 items-center hover:bg-blue-600"><Icons.Database /> SQL</button>
                </>
            )}
        </div>
    </header>
);

/**
 * Drag-and-drop upload area shown when no file is loaded.
 */
const DropZone = ({ onUpload, loading }) => (
    <div className="flex-1 flex flex-col items-center justify-center p-4">
        <div className="max-w-2xl w-full bg-white rounded-xl shadow-xl overflow-hidden">
            <div className="bg-indigo-600 p-10 text-center">
                <h1 className="text-4xl font-bold text-white mb-2 flex items-center justify-center gap-3"><Icons.FileSpreadsheet /> CleanSheet <span className="bg-indigo-500 text-xs px-2 py-1 rounded uppercase tracking-wider">Pro</span></h1>
                <p className="text-indigo-100 text-lg">The browser-based Swiss Army Knife for your data.</p>
            </div>
            <div className="p-12 border-4 border-dashed border-slate-200 hover:border-indigo-300 hover:bg-slate-50 transition-all cursor-pointer text-center group"
                onClick={() => document.getElementById('fileInput').click()}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => { e.preventDefault(); if (e.dataTransfer.files[0]) onUpload(e.dataTransfer.files[0]); }}>
                <input type="file" id="fileInput" className="hidden" onChange={(e) => e.target.files[0] && onUpload(e.target.files[0])} accept=".csv,.xlsx,.xls" />
                <div className="flex flex-col items-center gap-4">
                    <div className="w-24 h-24 bg-indigo-50 rounded-full flex items-center justify-center text-indigo-600 group-hover:scale-110 transition-transform"><Icons.Upload /></div>
                    <div><h3 className="text-2xl font-semibold text-slate-700">Drop spreadsheet here</h3><p className="text-slate-500 mt-2">.csv, .xlsx, .xls handled locally.</p></div>
                    {loading && <div className="text-indigo-600 animate-pulse font-medium">Processing...</div>}
                </div>
            </div>
        </div>
    </div>
);

/**
 * Left-hand sidebar containing Rules list and Saved Recipes.
 */
const Sidebar = ({ rules, headers, updateRule, deleteRule, addRule, savedRecipes, onSaveClick, onLoadRecipe, onDeleteRecipe }) => (
    <aside className="w-96 bg-white border-r border-slate-200 flex flex-col shrink-0 overflow-hidden">
        {/* Saved Recipes Section */}
        <div className="p-4 border-b border-slate-200 bg-slate-50">
            <div className="flex items-center justify-between mb-2">
                <h3 className="text-xs font-bold text-slate-500 uppercase">Saved Recipes</h3>
                <button onClick={onSaveClick} className="text-indigo-600 hover:text-indigo-700" title="Save Current Rules"><Icons.Save /></button>
            </div>
            {savedRecipes.length === 0 ? <p className="text-xs text-slate-400 italic">No saved recipes.</p> : (
                <div className="flex flex-wrap gap-2">
                    {savedRecipes.map(r => (
                        <span key={r.id} className="inline-flex items-center bg-white border border-slate-200 px-2 py-1 rounded text-xs text-slate-600 group">
                            <span onClick={() => onLoadRecipe(r)} className="cursor-pointer hover:text-indigo-600 mr-1">{r.name}</span>
                            <button onClick={() => onDeleteRecipe(r.id)} className="text-slate-300 hover:text-red-500"><Icons.X /></button>
                        </span>
                    ))}
                </div>
            )}
        </div>
        {/* Active Rules List */}
        <div className="flex-1 overflow-y-auto p-4 space-y-4 bg-slate-50/50 scrollbar-thin">
            <h3 className="text-xs font-bold text-slate-500 uppercase flex items-center gap-2"><Icons.ListFilter /> Filter Logic ({rules.length})</h3>
            {rules.map((rule, index) => (
                <RuleCard key={rule.id} rule={rule} index={index} headers={headers} updateRule={updateRule} deleteRule={deleteRule} />
            ))}
            <button onClick={addRule} className="w-full py-3 border-2 border-dashed border-slate-300 text-slate-400 rounded-xl text-sm font-medium hover:bg-white hover:border-indigo-400 hover:text-indigo-600 hover:shadow-sm flex justify-center items-center gap-2 transition-all"><Icons.Plus /> Add New Logic Rule</button>
        </div>
    </aside>
);

/**
 * Individual Card for defining a Logic Rule (IF X THEN Y).
 */
const RuleCard = ({ rule, index, headers, updateRule, deleteRule }) => (
    <div className="bg-white border border-slate-200 rounded-lg shadow-sm relative group overflow-hidden">
        <div className="bg-slate-50 border-b border-slate-100 px-3 py-2 flex items-center justify-between">
            <span className="text-xs font-bold text-slate-400 uppercase">Logic Rule #{index + 1}</span>
            <button onClick={() => deleteRule(rule.id)} className="text-slate-400 hover:text-red-500"><Icons.X /></button>
        </div>
        <div className="p-3 space-y-3">
            {/* Condition Section (IF) */}
            <div className="flex items-start gap-2">
                <div className="mt-1.5 bg-indigo-100 text-indigo-700 text-[10px] font-bold px-1.5 py-0.5 rounded">IF</div>
                <div className="flex-1 space-y-2">
                    <select className="w-full p-1.5 bg-white border border-slate-200 rounded text-sm text-slate-700 outline-none focus:border-indigo-500" value={rule.column} onChange={(e) => updateRule(rule.id, 'column', e.target.value)}>
                        {headers.map(h => <option key={h} value={h}>{h}</option>)}
                    </select>
                    <select className="w-full p-1.5 bg-white border border-slate-200 rounded text-sm text-slate-700 outline-none focus:border-indigo-500" value={rule.operator} onChange={(e) => updateRule(rule.id, 'operator', e.target.value)}>
                        <optgroup label="Basic Text"><option value="equals">equals</option><option value="does_not_equal">does not equal</option><option value="contains">contains</option><option value="does_not_contain">does not contain</option><option value="starts_with">starts with</option><option value="ends_with">ends with</option></optgroup>
                        <optgroup label="Content Checks"><option value="is_empty">is empty</option><option value="is_not_empty">is not empty</option><option value="is_email">is an email</option><option value="is_url">is a URL</option><option value="is_date">is a valid Date</option><option value="is_number">is a Number</option></optgroup>
                        <optgroup label="Numeric"><option value="greater_than">number is greater than</option><option value="less_than">number is less than</option></optgroup>
                    </select>
                    {!['is_empty', 'is_not_empty', 'is_email', 'is_url', 'is_date', 'is_number'].includes(rule.operator) && (
                        <input type="text" className="w-full p-1.5 bg-white border border-slate-200 rounded text-sm text-slate-700 outline-none focus:border-indigo-500 placeholder:text-slate-300" placeholder="Value to match..." value={rule.value} onChange={(e) => updateRule(rule.id, 'value', e.target.value)} />
                    )}
                </div>
            </div>
            {/* Action Section (THEN) */}
            <div className="flex items-start gap-2 border-t border-slate-100 pt-3">
                <div className="mt-1.5 bg-slate-200 text-slate-600 text-[10px] font-bold px-1.5 py-0.5 rounded">THEN</div>
                <div className="flex-1 space-y-2">
                    <select className={`w-full text-sm font-bold bg-transparent border border-slate-200 rounded p-1.5 outline-none focus:ring-1 focus:border-indigo-500 ${['remove', 'keep'].includes(rule.action) ? (rule.action === 'remove' ? 'text-red-600' : 'text-green-600') : 'text-indigo-600'}`} value={rule.action} onChange={(e) => updateRule(rule.id, 'action', e.target.value)}>
                        <optgroup label="Row Actions"><option value="remove">Remove Row</option><option value="keep">Keep Row</option></optgroup>
                        <optgroup label="Cell Actions"><option value="set">Set Value</option><option value="clear">Clear Cell</option><option value="append">Append Text</option><option value="prepend">Prepend Text</option><option value="upper">Uppercase</option><option value="lower">Lowercase</option></optgroup>
                    </select>
                    {!['remove', 'keep'].includes(rule.action) && (
                        <div className="animate-in fade-in slide-in-from-top-1 duration-200">
                            <label className="block text-[10px] font-bold text-slate-400 mb-1 uppercase">Target Column</label>
                            <select className="w-full p-1.5 bg-slate-50 border border-slate-200 rounded text-sm text-slate-700 outline-none focus:border-indigo-500" value={rule.targetColumn} onChange={(e) => updateRule(rule.id, 'targetColumn', e.target.value)}>
                                {headers.map(h => <option key={h} value={h}>{h}</option>)}
                            </select>
                        </div>
                    )}
                    {['set', 'append', 'prepend'].includes(rule.action) && (
                        <input type="text" className="w-full p-1.5 bg-slate-50 border border-slate-200 rounded text-sm text-slate-700 outline-none focus:border-indigo-500 placeholder:text-slate-300" placeholder={rule.action === 'set' ? "New Value..." : "Text to add..."} value={rule.actionValue || ''} onChange={(e) => updateRule(rule.id, 'actionValue', e.target.value)} />
                    )}
                </div>
            </div>
        </div>
    </div>
);

/**
 * Main Data Table view. Shows headers, rows, and the toolbar for mutations.
 */
const DataView = ({ fileData, processedRows, onToolClick }) => {
    const total = fileData.rows.length;
    const filtered = processedRows.length;

    return (
        <div className="flex-1 flex flex-col overflow-hidden relative">
            {/* Toolbar */}
            <div className="bg-white border-b border-slate-200 p-2 flex items-center gap-2 overflow-x-auto shrink-0">
                <span className="text-xs font-bold text-slate-400 uppercase mr-2">Tools:</span>
                <button onClick={() => onToolClick('trim')} className="px-3 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-xs font-medium text-slate-700 flex items-center gap-2 whitespace-nowrap"><Icons.Scissors /> Trim Space</button>
                <button onClick={() => onToolClick('findReplace')} className="px-3 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-xs font-medium text-slate-700 flex items-center gap-2 whitespace-nowrap"><Icons.Search /> Find & Replace</button>
                <button onClick={() => onToolClick('case')} className="px-3 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-xs font-medium text-slate-700 flex items-center gap-2 whitespace-nowrap"><Icons.Type /> Change Case</button>
                <button onClick={() => onToolClick('dedupe')} className="px-3 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-xs font-medium text-slate-700 flex items-center gap-2 whitespace-nowrap"><Icons.Copy /> Remove Dupes</button>
                <button onClick={() => onToolClick('split')} className="px-3 py-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded text-xs font-medium text-slate-700 flex items-center gap-2 whitespace-nowrap"><Icons.ArrowRight /> Split Column</button>
            </div>
            {/* Table Area */}
            <div className="flex-1 overflow-auto p-6 bg-slate-100">
                <div className="bg-white border border-slate-200 rounded-lg p-4 mb-4 flex items-center justify-between shadow-sm">
                    <div className="flex gap-8">
                        <div><span className="block text-xs font-bold text-slate-400 uppercase">Total Rows</span><span className="text-xl font-bold text-slate-700">{total.toLocaleString()}</span></div>
                        <div className="w-px bg-slate-200 h-full"></div>
                        <div><span className="block text-xs font-bold text-slate-400 uppercase">Filtered</span><span className="text-xl font-bold text-indigo-600">{filtered.toLocaleString()}</span></div>
                        <div className="w-px bg-slate-200 h-full"></div>
                        <div><span className="block text-xs font-bold text-slate-400 uppercase">Removed</span><span className="text-xl font-bold text-red-500">{(total - filtered).toLocaleString()}</span></div>
                    </div>
                </div>
                <div className="bg-white border border-slate-200 rounded-lg shadow-sm overflow-hidden">
                    <div className="overflow-x-auto">
                        <table className="min-w-full divide-y divide-slate-200">
                            <thead className="bg-slate-50">
                                <tr>
                                    <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 uppercase tracking-wider w-12">#</th>
                                    {fileData.headers.map((h, i) => <th key={i} className="px-4 py-3 text-left text-xs font-bold text-slate-700 uppercase tracking-wider whitespace-nowrap">{h}</th>)}
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-slate-100">
                                {/* Optimization: Only render first 100 rows for DOM performance */}
                                {processedRows.slice(0, 100).map((row, i) => (
                                    <tr key={i} className="hover:bg-indigo-50/50 transition-colors">
                                        <td className="px-4 py-2 whitespace-nowrap text-xs text-slate-400 font-mono">{i + 1}</td>
                                        {fileData.headers.map((h, colIndex) => (
                                            <td key={colIndex} className="px-4 py-2 whitespace-nowrap text-sm text-slate-600 max-w-xs truncate" title={row[colIndex]}>{row[colIndex] !== undefined ? String(row[colIndex]) : ''}</td>
                                        ))}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                    {filtered > 100 && <div className="p-2 text-center text-xs text-slate-400 bg-slate-50 border-t border-slate-100">Showing first 100 rows only</div>}
                </div>
            </div>
        </div>
    );
};

/**
 * Visualization View. Renders simple bar charts for column value frequency.
 */
const StatsView = ({ fileData, processedRows }) => (
    <div className="flex-1 overflow-auto p-8 bg-slate-50">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {fileData.headers.map(header => {
                const colIndex = fileData.headers.indexOf(header);
                const stats = DataService.calculateColumnStats(processedRows, colIndex);
                return (
                    <div key={header} className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                        <h3 className="font-bold text-slate-700 mb-4 flex items-center gap-2"><Icons.BarChart2 /> {header}</h3>
                        <div className="space-y-3">
                            {stats.sorted.map(([val, count], i) => (
                                <div key={i}>
                                    <div className="flex justify-between text-xs mb-1"><span className="text-slate-600 font-medium truncate max-w-[70%]">{val}</span><span className="text-slate-400">{count}</span></div>
                                    <div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden"><div className="bg-indigo-500 h-full rounded-full" style={{ width: `${(count / stats.maxVal) * 100}%` }}></div></div>
                                </div>
                            ))}
                        </div>
                    </div>
                );
            })}
        </div>
    </div>
);

/**
 * Reconciliation View. Handles two file upload zones and comparison results table.
 */
const ReconcileView = ({ rec, onExport }) => (
    <div className="flex-1 overflow-auto p-8 bg-slate-50 flex flex-col">
        <div className="max-w-6xl mx-auto w-full flex-1 flex flex-col">
            <div className="mb-8 text-center">
                <h2 className="text-2xl font-bold text-slate-800">Dataset Reconciliation</h2>
                <p className="text-slate-500">Compare two datasets to find matches and differences.</p>
            </div>
            {/* Upload Zones (Side by Side) */}
            <div className="grid grid-cols-2 gap-8 mb-8">
                {['A', 'B'].map(slot => {
                    const file = slot === 'A' ? rec.state.fileA : rec.state.fileB;
                    return (
                        <div key={slot} className="bg-white p-6 rounded-xl border-2 border-dashed border-slate-300 hover:border-indigo-400 transition-colors">
                            <h3 className="font-bold text-slate-700 mb-2">File {slot} (Source {slot === 'A' ? 1 : 2})</h3>
                            {!file ? (
                                <div className="h-32 bg-slate-50 rounded flex flex-col items-center justify-center cursor-pointer" onClick={() => document.getElementById(`recFile${slot}`).click()}>
                                    <input type="file" id={`recFile${slot}`} className="hidden" onChange={(e) => rec.handleUpload(e.target.files[0], slot)} />
                                    <div className="text-slate-400 mb-2"><Icons.Upload /></div>
                                    <span className="text-sm text-slate-500">Click to Upload CSV/Excel</span>
                                </div>
                            ) : (
                                <div className="space-y-4">
                                    <div className="flex items-center gap-2 text-green-600 bg-green-50 p-2 rounded text-sm font-medium"><Icons.CheckCircle2 /> {file.originalFileName}</div>
                                    <div>
                                        <label className="block text-xs font-bold text-slate-400 uppercase mb-1">Match Key (Unique ID)</label>
                                        <select className="w-full p-2 border rounded" value={slot === 'A' ? rec.state.keyA : rec.state.keyB} onChange={e => rec.setState({...rec.state, [slot === 'A' ? 'keyA' : 'keyB']: e.target.value})}>{file.headers.map(h => <option key={h} value={h}>{h}</option>)}</select>
                                    </div>
                                </div>
                            )}
                        </div>
                    );
                })}
            </div>
            {/* Action Button */}
            {rec.state.fileA && rec.state.fileB && !rec.state.results && (
                <div className="text-center mb-8">
                    <button onClick={rec.run} className="bg-indigo-600 text-white px-8 py-3 rounded-full font-bold shadow-lg hover:bg-indigo-700 transition-transform hover:scale-105 flex items-center gap-2 mx-auto"><Icons.ArrowLeftRight /> Run Comparison</button>
                </div>
            )}
            {/* Results Table */}
            {rec.state.results && (
                <div className="bg-white border border-slate-200 rounded-xl shadow-sm flex flex-col flex-1 overflow-hidden">
                    <div className="flex border-b border-slate-200">
                        <button onClick={() => rec.setActiveSubTab('onlyA')} className={`flex-1 py-4 text-sm font-bold border-b-2 transition-colors ${rec.activeSubTab === 'onlyA' ? 'border-indigo-600 text-indigo-600 bg-indigo-50' : 'border-transparent text-slate-500 hover:bg-slate-50'}`}>Only in File A ({rec.state.results.onlyA.length})</button>
                        <button onClick={() => rec.setActiveSubTab('onlyB')} className={`flex-1 py-4 text-sm font-bold border-b-2 transition-colors ${rec.activeSubTab === 'onlyB' ? 'border-indigo-600 text-indigo-600 bg-indigo-50' : 'border-transparent text-slate-500 hover:bg-slate-50'}`}>Only in File B ({rec.state.results.onlyB.length})</button>
                        <button onClick={() => rec.setActiveSubTab('matches')} className={`flex-1 py-4 text-sm font-bold border-b-2 transition-colors ${rec.activeSubTab === 'matches' ? 'border-green-600 text-green-600 bg-green-50' : 'border-transparent text-slate-500 hover:bg-slate-50'}`}>Matches ({rec.state.results.matches.length})</button>
                        {rec.state.results[rec.activeSubTab].length > 0 && (
                            <button onClick={() => onExport('csv', rec.state.results[rec.activeSubTab], rec.activeSubTab)} className="px-6 text-sm text-slate-500 hover:text-indigo-600 border-l border-slate-200"><Icons.Download /></button>
                        )}
                    </div>
                    <div className="flex-1 overflow-auto p-0">
                        <table className="min-w-full divide-y divide-slate-200">
                            <thead className="bg-slate-50 sticky top-0">
                                <tr>
                                    {(rec.activeSubTab === 'onlyB' ? rec.state.fileB.headers : rec.state.fileA.headers).map(h => <th key={h} className="px-6 py-3 text-left text-xs font-bold text-slate-500 uppercase tracking-wider whitespace-nowrap">{h}</th>)}
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-slate-200">
                                {rec.state.results[rec.activeSubTab].slice(0,100).map((row, i) => (
                                    <tr key={i}>{row.map((cell, c) => <td key={c} className="px-6 py-2 whitespace-nowrap text-sm text-slate-600">{cell}</td>)}</tr>
                                ))}
                                {rec.state.results[rec.activeSubTab].length === 0 && <tr><td colSpan="100%" className="p-10 text-center text-slate-400 italic">No records found in this category.</td></tr>}
                            </tbody>
                        </table>
                        {rec.state.results[rec.activeSubTab].length > 100 && <div className="p-4 text-center text-xs text-slate-400 bg-slate-50 border-t border-slate-200">Showing first 100 rows. Use Export to get full list.</div>}
                    </div>
                </div>
            )}
        </div>
    </div>
);

/**
 * Modal for Tool Actions (Find/Replace, Split, Case, Dedupe).
 */
const TransformModal = ({ mode, config, setConfig, onClose, onExecute, headers }) => (
    <div className="absolute top-0 left-0 w-full h-full bg-white/90 z-50 flex items-start justify-center pt-20 backdrop-blur-sm">
        <div className="bg-white shadow-2xl border border-slate-200 rounded-xl p-6 w-96 animate-in fade-in zoom-in duration-200">
            <div className="flex justify-between items-center mb-4">
                <h3 className="font-bold text-slate-800 text-lg capitalize">{mode === 'dedupe' ? 'Remove Duplicates' : mode.replace(/([A-Z])/g, ' $1').trim()}</h3>
                <button onClick={onClose}><Icons.X /></button>
            </div>
            <div className="space-y-4">
                {mode !== 'trim' && (
                    <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-1">Target Column</label>
                        <select className="w-full p-2 border rounded-lg text-sm" onChange={e => setConfig({...config, column: e.target.value})}>
                            <option value="">Select Column...</option>
                            {headers.map(h => <option key={h} value={h}>{h}</option>)}
                        </select>
                    </div>
                )}
                {mode === 'findReplace' && (
                    <>
                        <input type="text" placeholder="Find..." className="w-full p-2 border rounded-lg text-sm" onChange={e => setConfig({...config, find: e.target.value})} />
                        <input type="text" placeholder="Replace with..." className="w-full p-2 border rounded-lg text-sm" onChange={e => setConfig({...config, replace: e.target.value})} />
                        <button onClick={onExecute} className="w-full bg-indigo-600 text-white py-2 rounded-lg font-medium hover:bg-indigo-700">Execute</button>
                    </>
                )}
                {mode === 'case' && (
                    <>
                        <div className="flex gap-2">
                            <button onClick={() => setConfig({...config, type: 'upper'})} className={`flex-1 py-2 border rounded text-xs font-bold ${config.type === 'upper' ? 'bg-indigo-50 border-indigo-500 text-indigo-700' : 'hover:bg-slate-50'}`}>UPPER</button>
                            <button onClick={() => setConfig({...config, type: 'lower'})} className={`flex-1 py-2 border rounded text-xs font-bold ${config.type === 'lower' ? 'bg-indigo-50 border-indigo-500 text-indigo-700' : 'hover:bg-slate-50'}`}>lower</button>
                            <button onClick={() => setConfig({...config, type: 'title'})} className={`flex-1 py-2 border rounded text-xs font-bold ${config.type === 'title' ? 'bg-indigo-50 border-indigo-500 text-indigo-700' : 'hover:bg-slate-50'}`}>Title</button>
                        </div>
                        <button onClick={onExecute} disabled={!config.type} className="w-full bg-indigo-600 text-white py-2 rounded-lg font-medium hover:bg-indigo-700 disabled:opacity-50 mt-2">Convert</button>
                    </>
                )}
                {mode === 'split' && (
                    <>
                        <input type="text" placeholder="Delimiter (e.g. comma, space)" className="w-full p-2 border rounded-lg text-sm" onChange={e => setConfig({...config, delimiter: e.target.value})} />
                        <button onClick={onExecute} className="w-full bg-indigo-600 text-white py-2 rounded-lg font-medium hover:bg-indigo-700 mt-2">Split Column</button>
                    </>
                )}
                {mode === 'dedupe' && (
                    <>
                        <p className="text-xs text-slate-500">This will keep the first row found and remove subsequent rows with the same value in the selected column.</p>
                        <button onClick={onExecute} className="w-full bg-indigo-600 text-white py-2 rounded-lg font-medium hover:bg-indigo-700 mt-2">Remove Duplicates</button>
                    </>
                )}
            </div>
        </div>
    </div>
);

const SaveRecipeModal = ({ name, setName, onClose, onSave }) => (
    <div className="absolute top-0 left-0 w-full h-full bg-black/20 z-50 flex items-center justify-center backdrop-blur-sm">
        <div className="bg-white rounded-xl shadow-xl p-6 w-80">
            <h3 className="font-bold text-slate-800 mb-4">Save Recipe</h3>
            <input type="text" autoFocus placeholder="Recipe Name" className="w-full border border-slate-300 rounded-lg p-2 mb-4 text-sm outline-none focus:border-indigo-500" value={name} onChange={(e) => setName(e.target.value)} />
            <div className="flex gap-2">
                <button onClick={onClose} className="flex-1 py-2 text-sm font-medium text-slate-500 hover:bg-slate-50 rounded-lg">Cancel</button>
                <button onClick={onSave} className="flex-1 py-2 text-sm font-medium bg-indigo-600 text-white rounded-lg hover:bg-indigo-700">Save</button>
            </div>
        </div>
    </div>
);

// ==========================================
// 5. MAIN APP COMPONENT
// Description: Orchestration Layer
// ==========================================

const App = () => {
    // --- State Initialization ---
    const { fileData, setFileData, rules, setRules, loading, processedRows, processFile, updateRule, addRule, deleteRule, mutateData } = useFileProcessor();
    const rec = useReconciliation();

    // View State
    const [activeTab, setActiveTab] = useState('data');
    const [savedRecipes, setSavedRecipes] = useState([]);
    const [modals, setModals] = useState({ save: false, transform: null });
    const [transformConfig, setTransformConfig] = useState({});
    const [recipeName, setRecipeName] = useState('');

    // Load saved recipes from LocalStorage on mount
    useEffect(() => {
        const saved = localStorage.getItem('cleansheet_recipes');
        if (saved) setSavedRecipes(JSON.parse(saved));
    }, []);

    // --- Action Handlers ---

    const handleExport = (type, dataOverride = null, suffix = 'clean') => {
        if (activeTab === 'reconcile' && !dataOverride) return;
        const data = dataOverride || processedRows;
        // Determine headers based on export context
        let headers = fileData?.headers;
        if (dataOverride) {
             headers = rec.activeSubTab === 'onlyB' ? rec.state.fileB.headers : rec.state.fileA.headers;
        }
        const fname = dataOverride ? `reconcile_${suffix}` : `${fileData.originalFileName.split('.')[0]}_${suffix}`;
        DataService.exportFile(type, headers, data, fname);
    };

    /**
     * Handles Tool executions (Data Mutations) like Trim, Find/Replace.
     * These mutate the underlying data state, not just the filtered view.
     */
    const handleMutation = (type) => {
        const { column, find, replace, type: caseType, delimiter } = transformConfig;
        const colIdx = fileData.headers.indexOf(column);

        // Generic guard for column selection
        if (type !== 'trim' && colIdx === -1) return;

        let newRows = [...fileData.rows];
        let newHeaders = null;

        if (type === 'trim') {
            newRows = newRows.map(row => row.map(cell => typeof cell === 'string' ? cell.trim() : cell));
        } else if (type === 'findReplace') {
            newRows = newRows.map(row => {
                const r = [...row];
                if (typeof r[colIdx] === 'string') r[colIdx] = r[colIdx].replaceAll(find, replace);
                return r;
            });
        } else if (type === 'case') {
            newRows = newRows.map(row => {
                const r = [...row];
                const val = r[colIdx];
                if (typeof val === 'string') {
                    if (caseType === 'upper') r[colIdx] = val.toUpperCase();
                    if (caseType === 'lower') r[colIdx] = val.toLowerCase();
                    if (caseType === 'title') r[colIdx] = val.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
                }
                return r;
            });
        } else if (type === 'dedupe') {
            const seen = new Set();
            newRows = newRows.filter(row => {
                const val = row[colIdx];
                if (seen.has(val)) return false;
                seen.add(val);
                return true;
            });
        } else if (type === 'split') {
            const safeDelim = delimiter || ' ';
            let maxSplits = 0;
            // 1. Calculate how many new columns we need
            const splitData = newRows.map(row => {
                const val = String(row[colIdx] || '');
                if (!val) return [];
                const parts = val.split(safeDelim);
                if (parts.length > maxSplits) maxSplits = parts.length;
                return parts;
            });

            // 2. Apply split if valid
            if (maxSplits > 1) {
                newHeaders = [...fileData.headers];
                const newCols = [];
                for(let i = 1; i < maxSplits; i++) newCols.push(`${column}_${i+1}`);
                newHeaders.splice(colIdx + 1, 0, ...newCols);

                newRows = newRows.map((row, rIndex) => {
                    const r = [...row];
                    const parts = splitData[rIndex];
                    r[colIdx] = parts[0] || '';
                    const remainingParts = [];
                    for(let i = 1; i < maxSplits; i++) remainingParts.push(parts[i] || '');
                    r.splice(colIdx + 1, 0, ...remainingParts);
                    return r;
                });
            }
        }

        mutateData(newRows, newHeaders);
        setModals({ ...modals, transform: null });
        setTransformConfig({});
    };

    // Recipe Management
    const saveRecipe = () => {
        if (!recipeName) return;
        const newRecipe = { id: Date.now(), name: recipeName, rules };
        const updated = [...savedRecipes, newRecipe];
        setSavedRecipes(updated);
        localStorage.setItem('cleansheet_recipes', JSON.stringify(updated));
        setModals({ ...modals, save: false });
        setRecipeName('');
    };

    const deleteRecipe = (id) => {
        const updated = savedRecipes.filter(r => r.id !== id);
        setSavedRecipes(updated);
        localStorage.setItem('cleansheet_recipes', JSON.stringify(updated));
    };

    // --- View Routing Logic ---

    const renderContent = () => {
        if (activeTab === 'reconcile') {
            return <ReconcileView rec={rec} onExport={handleExport} />;
        }
        if (!fileData) {
            return <DropZone onUpload={processFile} loading={loading} />;
        }
        return (
            <div className="flex flex-1 overflow-hidden">
                <Sidebar
                    rules={rules}
                    headers={fileData.headers}
                    updateRule={updateRule}
                    addRule={addRule}
                    deleteRule={deleteRule}
                    savedRecipes={savedRecipes}
                    onSaveClick={() => setModals({ ...modals, save: true })}
                    onLoadRecipe={(r) => setRules(r.rules)}
                    onDeleteRecipe={deleteRecipe}
                />
                <main className="flex-1 flex flex-col overflow-hidden relative">
                    {activeTab === 'data' ? (
                        <DataView
                            fileData={fileData}
                            processedRows={processedRows}
                            onToolClick={(mode) => mode === 'trim' ? handleMutation('trim') : setModals({ ...modals, transform: mode })}
                        />
                    ) : (
                        <StatsView fileData={fileData} processedRows={processedRows} />
                    )}
                </main>
            </div>
        );
    };

    return (
        <div className="h-screen flex flex-col bg-slate-50 font-sans overflow-hidden">
            <Header
                fileData={fileData}
                activeTab={activeTab}
                setActiveTab={setActiveTab}
                onReset={activeTab === 'reconcile' ? rec.reset : () => setFileData(null)}
                onExport={handleExport}
            />

            <div className="flex flex-1 overflow-hidden">
                {renderContent()}
            </div>

            {/* Floating Modals */}
            {modals.transform && (
                <TransformModal
                    mode={modals.transform}
                    config={transformConfig}
                    setConfig={setTransformConfig}
                    onClose={() => { setModals({ ...modals, transform: null }); setTransformConfig({}); }}
                    onExecute={() => handleMutation(modals.transform)}
                    headers={fileData?.headers || []}
                />
            )}
            {modals.save && (
                <SaveRecipeModal
                    name={recipeName}
                    setName={setRecipeName}
                    onClose={() => setModals({ ...modals, save: false })}
                    onSave={saveRecipe}
                />
            )}
        </div>
    );
};

export default App;
