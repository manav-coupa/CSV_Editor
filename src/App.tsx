import React, { useState } from 'react';
import { DataGrid } from '@mui/x-data-grid';
// @ts-expect-error: No types for papaparse in this setup
import Papa, { ParseResult } from 'papaparse';
import * as XLSX from 'xlsx';
import { Button, Box, Typography, Dialog, DialogTitle, DialogContent, DialogActions, TextField, MenuItem, IconButton } from '@mui/material';
import EditIcon from '@mui/icons-material/Edit';
import { saveAs } from 'file-saver';

// Custom column type for DataGrid
interface MyColDef {
  field: string;
  headerName: string;
  width?: number;
  editable?: boolean;
  renderHeader?: (params: any) => React.ReactNode;
}

const getColumns = (data: any[], onEdit: (field: string) => void): MyColDef[] => {
  if (!data.length) return [];
  return Object.keys(data[0]).map((key) => ({
    field: key,
    headerName: key,
    width: 150,
    editable: true,
    renderHeader: (params: any) => (
      <Box sx={{ display: 'flex', alignItems: 'center' }}>
        <span>{key}</span>
        <IconButton size="small" onClick={() => onEdit(key)} aria-label={`Edit ${key}`} sx={{ ml: 1 }}>
          <EditIcon fontSize="small" />
        </IconButton>
      </Box>
    ),
  }));
};

const OPERATION_LIST = [
  { value: 'removeSpaces', label: 'Remove Spaces', example: '"A B C" → "ABC"' },
  { value: 'removeSpecial', label: 'Remove Special Characters', example: '"A!B@C# 123$%^" → "ABC 123" (keeps letters, numbers, spaces)' },
  { value: 'splitBy', label: 'Split by Character', example: 'Split by "@": "user@example.com" → "user" (original), "example.com" (new col)' },
  { value: 'custom', label: 'Custom Expression', example: 'e.g. value.toUpperCase() or value.replace(/\\d/g, "")' },
];

function App() {
  const [rows, setRows] = useState<any[]>([]);
  const [columns, setColumns] = useState<MyColDef[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [editCol, setEditCol] = useState<string | null>(null);
  const [operation, setOperation] = useState<string>('removeSpaces');
  const [splitChar, setSplitChar] = useState<string>('');
  const [customExpr, setCustomExpr] = useState<string>('');
  const [preview, setPreview] = useState<any[]>([]);
  const [opExample, setOpExample] = useState<string>('');

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    const ext = file.name.split('.').pop()?.toLowerCase();
    if (ext === 'csv') {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (result: ParseResult<any>) => {
          setRows(result.data as any[]);
          setColumns(getColumns(result.data as any[], handleEditColumn));
        },
      });
    } else if (ext === 'xlsx' || ext === 'xls') {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const bstr = evt.target?.result;
        if (typeof bstr === 'string' || bstr instanceof ArrayBuffer) {
          const wb = XLSX.read(bstr, { type: 'binary' });
          const wsname = wb.SheetNames[0];
          const ws = wb.Sheets[wsname];
          const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
          setRows(data as any[]);
          setColumns(getColumns(data as any[], handleEditColumn));
        }
      };
      reader.readAsBinaryString(file);
    } else {
      alert('Please upload a CSV or Excel file.');
    }
  };

  const handleEditColumn = (field: string) => {
    setEditCol(field);
    setOperation('removeSpaces');
    setSplitChar('');
    setCustomExpr('');
    setPreview([]);
    setOpExample(OPERATION_LIST[0].example);
  };

  const handleOperationChange = (op: string) => {
    setOperation(op);
    setOpExample(OPERATION_LIST.find(o => o.value === op)?.example || '');
  };

  const handleCloseDialog = () => {
    setEditCol(null);
    setPreview([]);
  };

  // Preview logic for column operation
  const handlePreview = () => {
    if (!editCol) return;
    let newRows = rows.map((row) => ({ ...row }));
    if (operation === 'removeSpaces') {
      newRows.forEach((row) => {
        row[editCol] = String(row[editCol] ?? '').replace(/\s+/g, '');
      });
    } else if (operation === 'removeSpecial') {
      newRows.forEach((row) => {
        row[editCol] = String(row[editCol] ?? '').replace(/[^a-zA-Z0-9 ]/g, '');
      });
    } else if (operation === 'splitBy' && splitChar) {
      newRows.forEach((row) => {
        const parts = String(row[editCol] ?? '').split(splitChar);
        row[editCol] = parts[0];
        row[`${editCol}_split`] = parts.slice(1).join(splitChar);
      });
      // Add new column if not present
      if (!columns.find(col => col.field === `${editCol}_split`)) {
        setColumns([...columns, {
          field: `${editCol}_split`,
          headerName: `${editCol} (split)`,
          width: 150,
        }]);
      }
    } else if (operation === 'custom' && customExpr) {
      try {
        // eslint-disable-next-line no-new-func
        const fn = new Function('value', `return (${customExpr});`);
        newRows.forEach((row) => {
          row[editCol] = fn(row[editCol]);
        });
      } catch (e) {
        alert('Invalid expression');
        return;
      }
    }
    setRows(newRows);
    setEditCol(null);
    setPreview([]);
  };

  const handleApply = () => {
    if (!editCol) return;
    let newRows = rows.map((row) => ({ ...row }));
    if (operation === 'removeSpaces') {
      newRows.forEach((row) => {
        row[editCol] = String(row[editCol] ?? '').replace(/\s+/g, '');
      });
    } else if (operation === 'removeSpecial') {
      newRows.forEach((row) => {
        row[editCol] = String(row[editCol] ?? '').replace(/[^a-zA-Z0-9 ]/g, '');
      });
    } else if (operation === 'splitBy' && splitChar) {
      newRows.forEach((row) => {
        const parts = String(row[editCol] ?? '').split(splitChar);
        row[editCol] = parts[0];
        row[`${editCol}_split`] = parts.slice(1).join(splitChar);
      });
      // Add new column if not present
      if (!columns.find(col => col.field === `${editCol}_split`)) {
        setColumns([...columns, {
          field: `${editCol}_split`,
          headerName: `${editCol} (split)`,
          width: 150,
        }]);
      }
    } else if (operation === 'custom' && customExpr) {
      try {
        // eslint-disable-next-line no-new-func
        const fn = new Function('value', `return (${customExpr});`);
        newRows.forEach((row) => {
          row[editCol] = fn(row[editCol]);
        });
      } catch (e) {
        alert('Invalid expression');
        return;
      }
    }
    setRows(newRows);
    setEditCol(null);
    setPreview([]);
  };

  // Download CSV
  const handleDownloadCSV = () => {
    if (!rows.length) return;
    const csv = Papa.unparse(rows);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    saveAs(blob, 'edited_data.csv');
  };

  // Download Excel
  const handleDownloadExcel = () => {
    if (!rows.length) return;
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    saveAs(blob, 'edited_data.xlsx');
  };

  return (
    <Box sx={{ p: 4 }}>
      <Typography variant="h4" gutterBottom>
        CSV/Excel Expression Editor
      </Typography>
      <Button variant="contained" component="label">
        Upload CSV/Excel
        <input type="file" accept=".csv,.xlsx,.xls" hidden onChange={handleFileUpload} />
      </Button>
      {fileName && (
        <Typography variant="subtitle1" sx={{ mt: 2 }}>
          File: {fileName}
        </Typography>
      )}
      <Box sx={{ mb: 2, display: 'flex', gap: 2 }}>
        {/* <Button variant="outlined" onClick={handleDownloadCSV} disabled={!rows.length}>
          Download CSV
        </Button> */}
        <Button variant="outlined" onClick={handleDownloadExcel} disabled={!rows.length}>
          Download Excel
        </Button>
      </Box>
      <Box sx={{ height: 500, width: '100%', mt: 4 }}>
        <DataGrid
          rows={rows.map((row, i) => ({ id: i, ...row }))}
          columns={columns}
          pageSizeOptions={[10, 25, 50]}
          initialState={{ pagination: { paginationModel: { pageSize: 10, page: 0 } } }}
          disableRowSelectionOnClick
        />
      </Box>
      <Dialog open={!!editCol} onClose={handleCloseDialog} maxWidth="sm" fullWidth>
        <DialogTitle>Edit Column: {editCol}</DialogTitle>
        <DialogContent>
          <TextField
            select
            label="Operation"
            value={operation}
            onChange={(e) => handleOperationChange(e.target.value)}
            fullWidth
            sx={{ mt: 2 }}
          >
            {OPERATION_LIST.map((op) => (
              <MenuItem key={op.value} value={op.value}>{op.label}</MenuItem>
            ))}
          </TextField>
          {opExample && (
            <Typography variant="body2" sx={{ mt: 1, color: 'text.secondary' }}>
              Example: {opExample}
            </Typography>
          )}
          {operation === 'splitBy' && (
            <TextField
              label="Split Character"
              value={splitChar}
              onChange={(e) => setSplitChar(e.target.value)}
              fullWidth
              sx={{ mt: 2 }}
            />
          )}
          {operation === 'custom' && (
            <TextField
              label="Custom Expression (e.g. value.replace(/\\s/g, ''))"
              value={customExpr}
              onChange={(e) => setCustomExpr(e.target.value)}
              fullWidth
              sx={{ mt: 2 }}
              helperText="Use 'value' as the variable for the cell value."
            />
          )}
          <Button variant="outlined" sx={{ mt: 2 }} onClick={handlePreview}>
            Preview
          </Button>
          {preview.length > 0 && (
            <Box sx={{ mt: 2, maxHeight: 200, overflow: 'auto', border: '1px solid #ccc', borderRadius: 1 }}>
              <pre style={{ margin: 0, padding: 8, fontSize: 12 }}>{JSON.stringify(preview.slice(0, 10), null, 2)}</pre>
              <Typography variant="caption" sx={{ ml: 1 }}>
                Showing first 10 rows
              </Typography>
            </Box>
          )}
        </DialogContent>
        <DialogActions>
          <Button onClick={handleCloseDialog}>Cancel</Button>
          <Button onClick={handleApply} variant="contained" color="primary">Apply</Button>
        </DialogActions>
      </Dialog>
    </Box>
  );
}

export default App;
