import { useEffect, useMemo, useState } from 'react';
import { Container, Typography, Paper, Stack, Button, TextField, MenuItem, Dialog, DialogTitle, DialogContent, DialogActions, FormControlLabel, Checkbox, Chip, IconButton, Select, InputLabel, FormControl, Autocomplete } from '@mui/material';
import { useNavigate } from 'react-router-dom';
import SaveIcon from '@mui/icons-material/Save';
import AddIcon from '@mui/icons-material/Add';
import { DataGrid, GridColDef } from '@mui/x-data-grid';
import * as XLSX from 'xlsx';
import api from '../api';
import { useToast } from '../components/ToastProvider';

interface OnProcPallet { _id?: string; poNumber: string; groupName: string; totalPallet: number; finishedPallet: number; transferredPallet?: number; status?: string; notes?: string; locked?: boolean; remainingPallet?: number; createdAt?: string }
interface OnProcBatch { _id: string; reference: string; poNumber: string; status: 'in-progress'|'partial-done'|'completed'; estFinishDate?: string; notes?: string; itemCount?: number; createdAt?: string }
interface ItemGroupRow { name: string; lineItem?: string }

export default function OnProcess() {
  const toast = useToast();
  const navigate = useNavigate();
  const todayYmd = useMemo(() => new Date().toISOString().slice(0, 10), []);
  const defaultBatchEstYmd = useMemo(() => {
    const d = new Date();
    d.setMonth(d.getMonth() + 2);
    return d.toISOString().slice(0, 10);
  }, []);
  const [items, setItems] = useState<OnProcPallet[]>([]);
  const [palletIdByGroup, setPalletIdByGroup] = useState<Record<string, string>>({});
  const [q, setQ] = useState('');
  const [file, setFile] = useState<File | null>(null);
  const [fileInputKey, setFileInputKey] = useState(0);
  const [result, setResult] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [batches, setBatches] = useState<OnProcBatch[]>([]);
  const [selectedBatch, setSelectedBatch] = useState<OnProcBatch | null>(null);
  const [bq, setBq] = useState<string>('');
  const [batchItems, setBatchItems] = useState<OnProcPallet[]>([]);
  const [batchStatus, setBatchStatus] = useState<'in-progress'|'partial-done'|'completed'>('in-progress');
  const [batchEst, setBatchEst] = useState<string>('');
  const [batchNotes, setBatchNotes] = useState<string>('');
  const [itemsSearch, setItemsSearch] = useState<string>('');
  const [selectedRows, setSelectedRows] = useState<string[]>([]);
  const [rowSelection, setRowSelection] = useState<string[]>([]);
  const [bulkStatus, setBulkStatus] = useState<'in_progress'|'partial'|'completed'|'cancelled'>('in_progress');
  const [bulkNotes, setBulkNotes] = useState<string>('');
  const [transferBulkOpen, setTransferBulkOpen] = useState(false);
  const [transferWarehouse, setTransferWarehouse] = useState('');
  const [transferMode, setTransferMode] = useState<'delivered'|'on_water'>('delivered');
  const [transferEDD, setTransferEDD] = useState<string>('');
  const [transferConfirmed, setTransferConfirmed] = useState(false);
  const [warehouses, setWarehouses] = useState<Array<{ _id: string; name: string; isPrimary?: boolean }>>([]);
  const [pendingDates, setPendingDates] = useState<Record<string, string>>({});
  const [dirty, setDirty] = useState(false);
  const [batchRefreshing, setBatchRefreshing] = useState(false);
  
  const [addOpen, setAddOpen] = useState(false);
  const [addGroupName, setAddGroupName] = useState('');
  const [addGroupOptions, setAddGroupOptions] = useState<Array<{ _id: string; name: string }>>([]);
  const [addGroupSelected, setAddGroupSelected] = useState<{ _id: string; name: string } | null>(null);
  const [addTotal, setAddTotal] = useState<string>('');
  const addGroupSelectable = useMemo(() => {
    const existing = new Set((batchItems || []).map(b => String(b.groupName || '').toLowerCase()));
    return (addGroupOptions || []).filter(o => !existing.has(String(o?.name || '').toLowerCase()));
  }, [addGroupOptions, batchItems]);
  const eligibleSelected = useMemo(() => {
    const ids = new Set(rowSelection.map(String));
    return batchItems.filter(it => {
      const id = String((it as any)._id || `${selectedBatch?._id}:${it.groupName}`);
      const remaining = Math.max(0, (Number(it.totalPallet || 0)) - ((Number((it as any).transferredPallet || 0)) + (Number(it.finishedPallet || 0))));
      const effectiveLocked = Boolean((it as any).locked) && remaining === 0;
      return ids.has(id) && !effectiveLocked && (Number(it.finishedPallet) > 0);
    });
  }, [rowSelection, batchItems, selectedBatch]);
  useEffect(() => {
    if (!transferBulkOpen) return;
    (async()=>{
      try {
        const { data } = await api.get<Array<{ _id: string; name: string; isPrimary?: boolean }>>('/warehouses');
        setWarehouses(Array.isArray(data) ? data : []);
      } catch {}
    })();
  }, [transferBulkOpen]);

  useEffect(() => {
    if (!transferBulkOpen) return;
    const list = Array.isArray(warehouses) ? warehouses : [];
    if (!list.length) return;

    const primary = list.find((w) => Boolean((w as any)?.isPrimary)) || null;
    const second = list.find((w) => !Boolean((w as any)?.isPrimary)) || null;
    const next = transferMode === 'on_water' ? (primary?._id || '') : (second?._id || '');
    if (next && String(next) !== String(transferWarehouse || '')) {
      setTransferWarehouse(String(next));
    }
  }, [transferBulkOpen, transferMode, warehouses, transferWarehouse]);

  useEffect(() => {
    (async () => {
      try {
        const { data } = await api.get<ItemGroupRow[]>('/item-groups');
        const groups = Array.isArray(data) ? data : [];
        const map: Record<string, string> = {};
        for (const g of groups) {
          const name = String((g as any)?.name || '').trim();
          if (!name) continue;
          map[name] = String((g as any)?.lineItem || '').trim();
        }
        setPalletIdByGroup(map);
      } catch {
        setPalletIdByGroup({});
      }
    })();
  }, []);

  // Stable columns for DataGrid to prevent width reset on re-render
  const itemColumns = useMemo<GridColDef[]>(() => ([
    { field: 'poNumber', headerName: 'PO #', width: 120 },
    { field: 'palletId', headerName: 'Pallet ID', width: 140 },
    { field: 'groupName', headerName: 'Pallet Description', flex: 2, minWidth: 220 },
    { field: 'totalPallet', headerName: 'Total Pallet', type: 'number', width: 140, editable: true, cellClassName: 'cell-editable-total' },
    { field: 'finishedPallet', headerName: 'Finished', type: 'number', width: 120, editable: true, cellClassName: 'cell-editable' },
    { field: 'transferredPallet', headerName: 'Transferred', type: 'number', width: 130 },
    { field: 'remainingPallet', headerName: 'Remaining', type: 'number', width: 120 },
    { field: 'status', headerName: 'Status', type: 'singleSelect', valueOptions: ['in_progress','partial','completed','cancelled'], width: 150, editable: true, renderCell: (params:any) => {
        const v = String(params.value || 'in_progress');
        const color = v === 'completed' ? 'success' : (v === 'partial' ? 'warning' : (v === 'cancelled' ? 'default' : 'info'));
        return <Chip size="small" label={v} color={color as any} variant={color==='info' ? 'outlined' : 'filled'} />;
      }
    },
  ]), [dirty]);

  // Stable derived rows with search filter, preserve id as string
  const itemRows = useMemo(() => (
    batchItems
      .filter(r => {
        const t = itemsSearch.trim().toLowerCase();
        if (!t) return true;
        const gn = String(r.groupName || '').toLowerCase();
        const pid = String(palletIdByGroup[String(r.groupName || '').trim()] || '').toLowerCase();
        return gn.includes(t) || pid.includes(t);
      })
      .map(r => ({
        id: String((r as any)._id || `${selectedBatch?._id}:${r.groupName}`),
        poNumber: r.poNumber,
        palletId: String(palletIdByGroup[String(r.groupName || '').trim()] || ''),
        groupName: r.groupName,
        totalPallet: r.totalPallet,
        finishedPallet: r.finishedPallet ?? 0,
        transferredPallet: (r as any).transferredPallet ?? 0,
        remainingPallet: Math.max(0, (r.totalPallet || 0) - (((r as any).transferredPallet || 0) + (r.finishedPallet || 0))),
        status: r.status || 'in_progress',
        locked: Boolean((r as any).locked) && (Math.max(0, (r.totalPallet || 0) - (((r as any).transferredPallet || 0) + (r.finishedPallet || 0)))) === 0,
      }))
  ), [batchItems, itemsSearch, selectedBatch, palletIdByGroup]);

  const load = async () => {
    setLoading(true);
    try {
      const { data } = await api.get<OnProcPallet[]>('/on-process');
      setItems(data || []);
    } finally { setLoading(false); }
  };
  useEffect(() => { load(); }, []);

  const loadBatches = async () => {
    try {
      const { data } = await api.get<OnProcBatch[]>('/on-process/batches');
      setBatches(data || []);
      // if a batch is selected, refresh it
      if (selectedBatch) {
        const updated = (data || []).find((b: any) => b._id === selectedBatch._id) || null;
        if (updated) setSelectedBatch(updated);
      }
    } catch {}
  };

  const refreshBatchesAndResetImport = async () => {
    setResult(null);
    setFile(null);
    setFileInputKey((k) => k + 1);
    await loadBatches();
  };

  const refreshSelectedBatch = async () => {
    if (!selectedBatch?._id) return;
    setBatchRefreshing(true);
    try {
      await loadBatches();
      await loadBatchItems(selectedBatch._id);
    } finally {
      setBatchRefreshing(false);
    }
  };
  const loadBatchItems = async (batchId: string) => {
    try {
      const { data } = await api.get<OnProcPallet[]>(`/on-process/batches/${batchId}/pallets`);
      setBatchItems(data || []);
      const rows = Array.isArray(data) ? data : [];
      const allDoneOrCancelled = rows.length > 0 && rows.every((r:any)=> ['completed','cancelled'].includes(String(r.status || '')));
      const allZero = rows.length > 0 && rows.every((r:any)=> Number(r.finishedPallet||0) === 0 && Number((r as any).transferredPallet||0) === 0);
      if (allDoneOrCancelled) {
        setBatchStatus('completed');
        const today = new Date();
        const ymd = today.toISOString().slice(0,10);
        setBatchEst(ymd);
      } else if (allZero) {
        setBatchStatus('in-progress');
      } else {
        setBatchStatus('partial-done');
      }
    } catch {}
  };
  useEffect(() => { loadBatches(); }, []);

  const batchCols = useMemo<GridColDef[]>(() => ([
    { field: 'reference', headerName: 'Reference', flex: 1, minWidth: 140 },
    { field: 'poNumber', headerName: 'PO #', width: 120 },
    { field: 'pallets', headerName: 'Pallets', width: 150, sortable: false, filterable: false, renderCell: (params:any) => {
      const b = params?.row?.__raw as OnProcBatch;
      return (
        <Button size="small" variant="outlined" onClick={()=> exportBatchXlsx(b)}>Export Pallet</Button>
      );
    } },
    { field: 'status', headerName: 'Status', width: 150, renderCell: (params:any) => {
      const st = String(params.value || 'in-progress');
      const color = st === 'completed' ? 'success' : (st === 'partial-done' ? 'warning' : 'info');
      return <Chip size="small" label={st} color={color as any} variant={color==='info' ? 'outlined' : 'filled'} />;
    } },
    { field: 'estFinishDate', headerName: 'Estimated Date Finish', width: 180, renderCell: (params:any) => {
      const v = params?.row?.estFinishDate;
      const s = v ? String(v).substring(0,10) : '';
      return <span>{s}</span>;
    } },
    { field: 'notes', headerName: 'Notes/Remarks', flex: 1.2, minWidth: 220 },
    { field: 'actions', headerName: 'Actions', width: 120, sortable: false, filterable: false, renderCell: (params:any) => {
      const b = params?.row?.__raw as OnProcBatch;
      return (
        <Stack direction="row" spacing={1}>
          <Button size="small" variant="outlined" onClick={()=> selectBatch(b)}>Edit</Button>
        </Stack>
      );
    } },
  ]), []);

  const batchRows = useMemo(() => {
    const t = bq.trim().toLowerCase();
    const list = (batches || []).map((b) => ({
      id: b._id,
      reference: b.reference,
      poNumber: b.poNumber,
      pallets: '',
      status: b._id === selectedBatch?._id ? batchStatus : b.status,
      estFinishDate: b._id === selectedBatch?._id ? (batchEst || '') : (b.estFinishDate ? new Date(b.estFinishDate as any).toISOString().slice(0, 10) : ''),
      notes: b.notes || '',
      __raw: b,
    }));
    if (!t) return list;
    return list.filter(r =>
      String(r.poNumber||'').toLowerCase().includes(t) ||
      String(r.reference||'').toLowerCase().includes(t) ||
      String(r.estFinishDate||'').toLowerCase().includes(t) ||
      String(r.notes||'').toLowerCase().includes(t)
    );
  }, [batches, bq, selectedBatch, batchStatus, batchEst]);
  const exportBatchXlsx = async (b: OnProcBatch) => {
    try {
      const { data } = await api.get(`/on-process/batches/${b._id}/export`, { responseType: 'blob' as any });
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${b.reference}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e:any) {
      toast.error(e?.response?.data?.message || 'Export failed');
    }
  };
  const selectBatch = (b: OnProcBatch) => {
    setSelectedBatch(b);
    setBatchStatus(b.status);
    setBatchEst(b.estFinishDate ? new Date(b.estFinishDate as any).toISOString().slice(0, 10) : defaultBatchEstYmd);
    setBatchNotes(b.notes || '');
    loadBatchItems(b._id);
  };

  const downloadTemplate = () => {
    const header = ['PO #','Pallet Description','Total Pallet'];
    const ws = XLSX.utils.aoa_to_sheet([header]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'on_process_pallet_template.xlsx');
  };

  const importFile = async () => {
    if (!file) { toast.error('Please select a .xlsx file'); return; }
    const fd = new FormData();
    fd.append('file', file);
    try {
      const { data } = await api.post('/on-process/pallets/import', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      setResult(data);
      const created = data.created ?? 0; const issues = data.errorCount ?? 0;
      if (issues) toast.warning(`Import completed with issues. Created: ${created}, Issues: ${issues}. See details below.`);
      else toast.success(`Import completed. Created: ${created}.`);
      setFile(null);
      if (selectedBatch && !String(batchEst || '').trim()) {
        setBatchEst(defaultBatchEstYmd);
      }
      await load();
      await loadBatches();
      if (selectedBatch) await loadBatchItems(selectedBatch._id);
    } catch (e: any) {
      const msg = e?.response?.data?.message || e?.message || 'Import failed';
      toast.error(msg);
    }
  };

  const filtered = items.filter(it => {
    const t = q.trim().toLowerCase();
    if (!t) return true;
    return (
      it.poNumber.toLowerCase().includes(t) ||
      it.groupName.toLowerCase().includes(t)
    );
  });

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>ON-PROCESS</Typography>

      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs:'column', md:'row' }} spacing={2} alignItems="center">
          <input key={fileInputKey} type="file" accept=".xlsx" onChange={(e)=> setFile(e.target.files?.[0] || null)} />
          <Button variant="contained" onClick={importFile} disabled={!file}>Import (.xlsx)</Button>
          <Button variant="outlined" onClick={downloadTemplate}>Download Template</Button>
          <Button variant="outlined" onClick={refreshBatchesAndResetImport}>Refresh</Button>
        </Stack>
        {result && (
          <Paper variant="outlined" sx={{ p:2, mt:2, bgcolor:'#fafafa' }}>
            <Typography variant="subtitle2" gutterBottom>Import Summary</Typography>
            <Typography variant="body2">Created: <b>{result.created ?? 0}</b></Typography>
            <Typography variant="body2">Skipped: <b>{result.skipped ?? 0}</b></Typography>
            {Boolean(result.errorCount) && (
              <>
                <Typography variant="body2" sx={{ mt:1 }}>Issues found: <b>{result.errorCount}</b></Typography>
                <table style={{ width:'100%', borderCollapse:'collapse', marginTop: 8 }}>
                  <thead>
                    <tr>
                      <th align="left">Row</th>
                      <th align="left">PO #</th>
                      <th align="left">Pallet Description</th>
                      <th align="left">Reason</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(result.errors || []).slice(0,20).map((er:any, i:number)=> (
                      <tr key={i} style={{ borderTop:'1px solid #eee' }}>
                        <td>{er.rowNum ?? '-'}</td>
                        <td>{er.poNumber ?? '-'}</td>
                        <td>{er.groupName ?? '-'}</td>
                        <td>{Array.isArray(er.errors) ? er.errors.join(', ') : String(er.errors || '')}</td>
                      </tr>
                    ))}
                    {(result.errors || []).length > 20 && (
                      <tr><td colSpan={4} style={{ color:'#666', paddingTop: 6 }}>Showing first 20 of {result.errors.length} issues…</td></tr>
                    )}
                  </tbody>
                </table>
              </>
            )}
          </Paper>
        )}
      </Paper>

      <Paper sx={{ p:2, mb:2 }}>
        <Typography variant="h6" gutterBottom>Batches</Typography>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems="center" sx={{ mb: 1 }}>
          <TextField size="small" label="Search (PO#/Reference/Date/Notes)" value={bq} onChange={(e)=> setBq(e.target.value)} sx={{ minWidth: 280, flex: 1 }} />
        </Stack>
        <div style={{ height: 360, width: '100%' }}>
          <DataGrid
            rows={batchRows}
            columns={batchCols}
            disableRowSelectionOnClick
            onRowDoubleClick={(params: any) => {
              const raw = (params as any)?.row?.__raw;
              if (raw) selectBatch(raw);
            }}
            density="compact"
            pageSizeOptions={[5,10,20,50,100]}
          />
        </div>
        <Dialog
          open={Boolean(selectedBatch)}
          onClose={()=> {
            setTransferBulkOpen(false);
            setTransferConfirmed(false);
            setAddOpen(false);
            setSelectedBatch(null);
          }}
          maxWidth="xl"
          fullWidth
        >
          <DialogTitle>Batch {selectedBatch?.reference || ''} — Items</DialogTitle>
          <DialogContent sx={{ pt: 2 }}>
            <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} sx={{ mb: 2 }}>
              <TextField fullWidth label="Notes/Remarks" value={batchNotes} onChange={(e)=> setBatchNotes(e.target.value)} />
              <Button variant="outlined" onClick={refreshSelectedBatch} disabled={batchRefreshing} sx={{ minWidth: 120 }}>
                Refresh
              </Button>
            </Stack>
            <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} justifyContent="space-between" sx={{ mb: 1 }}>
              <TextField size="small" label="Search pallet description" value={itemsSearch} onChange={(e)=> setItemsSearch(e.target.value)} sx={{ minWidth: 260, flex: 1 }} />
              <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mt: { xs: 1, sm: 0 } }}>
                <TextField size="small" select label="Set status (PO)" value={batchStatus} onChange={(e)=> setBatchStatus(e.target.value as any)} sx={{ minWidth: 200 }}>
                  <MenuItem value="in-progress">in-progress</MenuItem>
                  <MenuItem value="partial-done">partial-done</MenuItem>
                  <MenuItem value="completed">completed</MenuItem>
                </TextField>
                <TextField
                  required
                  size="small"
                  type="date"
                  label="Estimated Date Finish"
                  InputLabelProps={{ shrink: true }}
                  helperText="Recommended to input EDD for tracking"
                  FormHelperTextProps={{ sx: { color: 'error.main' } }}
                  inputProps={{ min: todayYmd }}
                  value={batchEst}
                  onChange={(e)=> {
                    const v = String(e.target.value || '');
                    if (v && v < todayYmd) {
                      toast.error('Estimated Date Finish cannot be earlier than today');
                      return;
                    }
                    setBatchEst(v);
                  }}
                  sx={{ minWidth: 200 }}
                />
              </Stack>
            </Stack>
            <div style={{ height: 520, width: '100%' }}>
              <DataGrid
                rows={itemRows}
                columns={itemColumns}
                getRowId={(row: any) => row.id}
                getRowClassName={(params:any)=> {
                  const st = String(params?.row?.status || '');
                  const remaining = Number(params?.row?.remainingPallet || 0);
                  const locked = Boolean(params?.row?.locked) && remaining === 0;
                  return `${st === 'completed' ? 'row-completed' : ''} ${st === 'cancelled' ? 'row-cancelled' : ''} ${locked ? 'row-locked' : ''}`.trim();
                }}
                checkboxSelection
                density="compact"
                disableRowSelectionOnClick
                isRowSelectable={(params:any)=> {
                  const remaining = Number((params.row as any)?.remainingPallet || 0);
                  const locked = Boolean((params.row as any)?.locked) && remaining === 0;
                  const finished = Number((params.row as any)?.finishedPallet || 0);
                  return !locked && finished > 0;
                }}
                isCellEditable={(params:any)=>{
                  const editableFields = new Set(['totalPallet','finishedPallet','status']);
                  const remaining = Number((params.row as any)?.remainingPallet || 0);
                  const locked = Boolean((params.row as any)?.locked) && remaining === 0;
                  const st = String((params.row as any)?.status || '');

                  if (params.field === 'totalPallet') {
                    return st !== 'cancelled';
                  }

                  if (locked) return false;
                  if (st === 'cancelled') return false;

                  return editableFields.has(params.field);
                }}
                onRowSelectionModelChange={(m: any)=> {
                  let ids: string[] = [];
                  if (Array.isArray(m)) {
                    ids = (m as any[]).map((v)=> String(v));
                  } else if (m && typeof m === 'object') {
                    const maybe = (m as any).ids ?? (m as any).selection ?? m;
                    if (Array.isArray(maybe)) ids = maybe.map((v:any)=> String(v));
                    else if (maybe && typeof (maybe as any)[Symbol.iterator] === 'function') ids = Array.from(maybe as any, (v:any)=> String(v));
                  }
                  setRowSelection(ids);
                  setSelectedRows(ids);
                }}
                processRowUpdate={(newRow:any)=>{
                  const id = String(newRow.id);
                  const fieldVals = newRow;
                  setBatchItems(prev => {
                    const next = prev.map(it => {
                      const iid = String((it as any)._id || `${selectedBatch?._id}:${it.groupName}`);
                      if (iid !== id) return it;
                      const wasLocked = Boolean((it as any)?.locked);
                      const prevStatus = String((it as any)?.status || 'in_progress');
                      const prevTotal = Number((it as any)?.totalPallet || 0);
                      let total = Number(fieldVals.totalPallet);
                      if (!Number.isFinite(total)) total = it.totalPallet || 0;
                      let finished = Number(fieldVals.finishedPallet);
                      if (!Number.isFinite(finished) || finished < 0) finished = it.finishedPallet || 0;
                      const transferred = Number((it as any).transferredPallet || 0);
                      let status = String(fieldVals.status || it.status || 'in_progress');
                      if (!['in_progress','partial','completed','cancelled'].includes(status)) status = it.status || 'in_progress';
                      if (status === 'cancelled') {
                        if (transferred > 0) {
                          status = it.status || 'in_progress';
                        } else {
                          total = 0;
                          finished = 0;
                        }
                      }

                      if (wasLocked || prevStatus === 'completed') {
                        total = Math.max(prevTotal, total || 0);
                      }

                      if (status !== 'cancelled') {
                        total = Math.max(1, transferred, total || 0);
                      } else {
                        total = Math.max(total, transferred);
                      }
                      const maxFinish = Math.max(0, total - transferred);
                      finished = Math.min(finished, maxFinish);
                      const remaining = Math.max(0, total - (transferred + finished));
                      if (status !== 'cancelled') {
                        status = remaining === 0 ? 'completed' : (finished > 0 ? 'partial' : 'in_progress');
                      }

                      const nextLocked = wasLocked && total > prevTotal ? false : (it as any).locked;
                      return { ...it, totalPallet: total, finishedPallet: finished, status, locked: nextLocked } as any;
                    });
                    const allDoneOrCancelled = next.length > 0 && next.every((r:any)=> ['completed','cancelled'].includes(String(r.status || '')));
                    const allZero = next.length > 0 && next.every((r:any)=> Number(r.finishedPallet||0) === 0 && Number((r as any).transferredPallet||0) === 0);
                    if (allDoneOrCancelled) {
                      setBatchStatus('completed');
                      const today = new Date();
                      const ymd = today.toISOString().slice(0,10);
                      setBatchEst(ymd);
                    } else if (allZero) {
                      setBatchStatus('in-progress');
                    } else {
                      setBatchStatus('partial-done');
                    }
                    setDirty(true);
                    return next;
                  });
                  return newRow;
                }}
                pageSizeOptions={[10,20,50,100]}
                sx={{
                  '& .row-completed': {
                    backgroundColor: '#e8f5e9',
                  },
                  '& .row-locked': {
                    opacity: 0.6,
                  },
                  '& .row-cancelled': {
                    backgroundColor: '#ffebee',
                  },
                  '& .cell-editable': {
                    backgroundColor: '#fff8e1',
                    boxShadow: 'inset 0 0 0 1px #ffe082',
                  },
                  '& .cell-editable-total': {
                    backgroundColor: '#e3f2fd',
                    boxShadow: 'inset 0 0 0 1px #90caf9',
                  }
                }}
              />
            </div>
          </DialogContent>
          <DialogActions>
            <Stack direction="row" spacing={2} sx={{ width: '100%', justifyContent: 'flex-start', pl: 1 }}>
              <IconButton color="primary" onClick={async()=> {
                setAddOpen(true);
                setAddGroupName('');
                setAddGroupSelected(null);
                setAddTotal('');
                try {
                  const { data } = await api.get<Array<{ _id: string; name: string }>>('/item-groups');
                  setAddGroupOptions(Array.isArray(data) ? data : []);
                } catch {}
              }} aria-label="Add Pallet Description" title="Add Pallet Description">
                <AddIcon />
              </IconButton>
              <Button variant="contained" onClick={async()=> {
                try {
                  if (batchEst && String(batchEst) < todayYmd) {
                    toast.error('Estimated Date Finish cannot be earlier than today');
                    return;
                  }
                  if (!selectedBatch?._id) return;
                  await api.patch(`/on-process/batches/${selectedBatch._id}`, { status: batchStatus, estFinishDate: batchEst || null, notes: batchNotes });
                  await api.patch(`/on-process/batches/${selectedBatch._id}/pallets`, { pallets: batchItems.map(b => ({ groupName: b.groupName, totalPallet: b.totalPallet, finishedPallet: b.finishedPallet, status: b.status })) });
                  toast.success('Changes saved');
                  await loadBatches();
                  await loadBatchItems(selectedBatch._id);
                  setDirty(false);
                } catch (e:any) { toast.error(e?.response?.data?.message || 'Save failed'); }
              }}>SAVE CHANGES</Button>
              <Button variant="outlined" disabled={dirty || (rowSelection.length === 0)} onClick={()=> {
                if (eligibleSelected.length === 0) { toast.warning('No selected rows have finished pallets to transfer.'); return; }
                setTransferBulkOpen(true);
              }}>
                Transfer Selected
              </Button>
              <Button variant="outlined" onClick={()=> setSelectedBatch(null)}>CLOSE/CANCEL</Button>
            </Stack>
          </DialogActions>

          <Dialog open={addOpen} onClose={()=> setAddOpen(false)} maxWidth="xs" fullWidth>
              <DialogTitle>Add Pallet Description</DialogTitle>
              <DialogContent sx={{ pt: 2 }}>
                <Autocomplete
                  options={addGroupSelectable}
                  getOptionLabel={(o)=> o?.name || ''}
                  value={addGroupSelected}
                  onChange={(_, v)=> { setAddGroupSelected(v); setAddGroupName(v?.name || ''); }}
                  filterOptions={(options, state) => {
                    const q = String(state?.inputValue || '').trim().toLowerCase();
                    if (!q) return options;
                    return (Array.isArray(options) ? options : []).filter((o: any) => {
                      const name = String(o?.name || '').trim().toLowerCase();
                      const pid = String(palletIdByGroup[String(o?.name || '').trim()] || '').trim().toLowerCase();
                      return (name && name.includes(q)) || (pid && pid.includes(q));
                    });
                  }}
                  renderInput={(params)=> (
                    <TextField {...params} label="Search Pallet Description" placeholder="Type to search" sx={{ mb: 2 }} />
                  )}
                />
                <TextField fullWidth type="number" label="Total Pallet" value={addTotal} onChange={(e)=> setAddTotal(e.target.value)} />
              </DialogContent>
              <DialogActions>
                <Button onClick={()=> setAddOpen(false)}>Cancel</Button>
                <Button variant="contained" disabled={!addGroupSelected || !(Number(addTotal) > 0)} onClick={async()=>{
                  try {
                    const batchId = selectedBatch?._id;
                    if (!batchId) return;
                    await api.post(`/on-process/batches/${batchId}/pallets`, { groupName: addGroupSelected?.name, totalPallet: Number(addTotal) });
                    toast.success('Pallet description added');
                    setAddOpen(false);
                    setAddGroupName(''); setAddGroupSelected(null); setAddTotal('');
                    await loadBatchItems(batchId);
                  } catch (e:any) { toast.error(e?.response?.data?.message || 'Failed to add'); }
                }}>Add</Button>
              </DialogActions>
            </Dialog>

            <Dialog open={transferBulkOpen} onClose={()=> setTransferBulkOpen(false)} maxWidth="sm" fullWidth>
              <DialogTitle>Transfer selected finished pallets</DialogTitle>
              <DialogContent sx={{ pt: 2 }}>
                <Typography variant="body2" sx={{ mb: 0.5 }}>
                  {eligibleSelected.length} row(s) selected with finished pallets. Choose a warehouse and transfer mode to move them into inventory or shipment.
                </Typography>
                <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 2 }}>
                  {(rowSelection.length - eligibleSelected.length) > 0 ? `${rowSelection.length - eligibleSelected.length} selected row(s) are ineligible (finished = 0 or locked) and will be skipped.` : 'All selected rows are eligible.'}
                </Typography>
                <FormControl fullWidth sx={{ mb: 2 }}>
                  <InputLabel id="transfer-mode-label-bulk">Transfer as</InputLabel>
                  <Select labelId="transfer-mode-label-bulk" value={transferMode} label="Transfer as" onChange={(e)=> setTransferMode(e.target.value as any)}>
                    <MenuItem value="delivered">Delivered (receive today)</MenuItem>
                    <MenuItem value="on_water">On-Water (ask EDD, create shipment)</MenuItem>
                  </Select>
                </FormControl>
                <TextField fullWidth select disabled label="Warehouse (required)" value={transferWarehouse} onChange={(e)=> setTransferWarehouse(e.target.value)} sx={{ mb: 2 }}>
                  {warehouses.map(w => (
                    <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>
                  ))}
                </TextField>
                {transferMode === 'on_water' && (
                  <TextField
                    fullWidth
                    type="date"
                    label="EDD (required for On-Water)"
                    InputLabelProps={{ shrink: true }}
                    inputProps={{ min: todayYmd }}
                    value={transferEDD}
                    onChange={(e)=> {
                      const v = String(e.target.value || '');
                      if (v && v < todayYmd) {
                        toast.error('EDD cannot be earlier than today');
                        return;
                      }
                      setTransferEDD(v);
                    }}
                    sx={{ mb: 2 }}
                  />
                )}
                <Typography variant="body2" sx={{ mb: 1 }}>
                  Reference used: <b>PO - {selectedBatch?.poNumber || ''}</b>
                </Typography>
                <FormControlLabel control={<Checkbox checked={transferConfirmed} onChange={(e)=> setTransferConfirmed(e.target.checked)} />} label="I confirm this action is correct and cannot be undone." />
              </DialogContent>
              <DialogActions>
                <Button onClick={()=> setTransferBulkOpen(false)}>Cancel</Button>
                <Button variant="contained" disabled={!transferWarehouse.trim() || !transferConfirmed || (transferMode==='on_water' && !transferEDD) || eligibleSelected.length===0} onClick={async()=> {
                  try {
                    if (transferMode === 'on_water' && transferEDD && String(transferEDD) < todayYmd) {
                      toast.error('EDD cannot be earlier than today');
                      return;
                    }
                    const batchId = selectedBatch?._id;
                    if (!batchId) return;
                    const transferItems = eligibleSelected.map(b => ({ groupName: b.groupName, pallets: b.finishedPallet || 0 }));
                    await api.post(`/on-process/batches/${batchId}/pallets/transfer`, { mode: transferMode, warehouseId: transferWarehouse, estDeliveryDate: transferMode==='on_water' ? transferEDD : undefined, items: transferItems });
                    toast.success('Transfer created');
                    setTransferBulkOpen(false);
                    setTransferWarehouse('');
                    setTransferMode('delivered');
                    setTransferEDD('');
                    setTransferConfirmed(false);
                    await loadBatches();
                    await loadBatchItems(batchId);
                    if (transferMode === 'on_water') {
                      setSelectedBatch(null);
                      navigate('/ship');
                    }
                  } catch (e:any) { toast.error(e?.response?.data?.message || 'Transfer failed'); }
                }}>Confirm & Transfer</Button>
              </DialogActions>
            </Dialog>
        </Dialog>
      </Paper>
    </Container>
  );
}
