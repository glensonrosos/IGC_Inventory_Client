import React, { useEffect, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, Dialog, DialogTitle, DialogContent, DialogActions, IconButton, Menu, MenuItem } from '@mui/material';
import * as XLSX from 'xlsx';
import MoreVertIcon from '@mui/icons-material/MoreVert';
import { DataGrid, GridColDef, GridRenderCellParams } from '@mui/x-data-grid';
import api from '../api';
import { useToast } from '../components/ToastProvider';

interface Warehouse { _id: string; name: string; address?: string }

export default function Warehouses() {
  const [rows, setRows] = useState<Warehouse[]>([]);
  const [loading, setLoading] = useState(false);
  const [name, setName] = useState('');
  const [address, setAddress] = useState('');
  const [q, setQ] = useState('');
  const toast = useToast();
  const [page, setPage] = useState(0);
  const [pageSize, setPageSize] = useState(10);
  const [rowCount, setRowCount] = useState(0);
  const [editOpen, setEditOpen] = useState(false);
  const [editId, setEditId] = useState<string>('');
  const [editName, setEditName] = useState('');
  const [editAddress, setEditAddress] = useState('');
  const [exportingId, setExportingId] = useState<string>('');
  const [exportingAll, setExportingAll] = useState<boolean>(false);

  const load = async () => {
    setLoading(true);
    try {
      const { data } = await api.get<Warehouse[]>('/warehouses', { params: { q } });
      setRows(data);
    } finally {
      setLoading(false);
    }
  };

  const exportAllStocks = async () => {
    setExportingAll(true);
    try {
      // Fetch warehouses and pallet group overview
      const allWsResp = await api.get<Warehouse[]>('/warehouses');
      const allWarehouses: Warehouse[] = allWsResp.data || [];
      if (!allWarehouses.length) { setExportingAll(false); return; }
      const { data: overview } = await api.get('/pallet-inventory/groups');
      const groups = Array.isArray(overview) ? overview : [];
      // Build header: Pallet Group, Total Qty, then each warehouse name
      const whNames = allWarehouses.map(w => w.name);
      const header = ['Pallet Description', 'Total Pallet', ...whNames];
      // Build data rows
      const whIndex: Record<string, number> = {};
      allWarehouses.forEach((w, idx) => { whIndex[String(w._id)] = idx; });
      const rowsAoa: any[] = [];
      for (const g of groups) {
        const per: any[] = new Array(allWarehouses.length).fill(0);
        const perWarehouse = Array.isArray(g.perWarehouse) ? g.perWarehouse : [];
        for (const p of perWarehouse) {
          const i = whIndex[String(p.warehouseId)] ?? -1;
          if (i >= 0) per[i] = Number(p.pallets || 0);
        }
        const total = per.reduce((a,b)=> a + (Number(b)||0), 0);
        rowsAoa.push([g.groupName || '', total, ...per]);
      }
      // Create workbook
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([header, ...rowsAoa]);
      XLSX.utils.book_append_sheet(wb, ws, 'Pallet Inventory');
      const d = new Date();
      const pad = (n: number) => n.toString().padStart(2,'0');
      const ts = `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
      XLSX.writeFile(wb, `all-warehouses-pallet-inventory-(${ts}).xlsx`);
    } finally {
      setExportingAll(false);
    }
  };

  useEffect(() => { load(); }, []);

  const addWarehouse = async () => {
    if (!name.trim()) { toast.error('Name is required'); return; }
    try {
      await api.post('/warehouses', { name: name.trim(), address });
      toast.success('Warehouse added');
      setName(''); setAddress('');
      await load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || e?.message || 'Failed to add warehouse';
      toast.error(msg);
    }
  };

  const exportStock = async (id: string, name: string) => {
    setExportingId(id);
    try {
      // Get pallet inventory overview and filter by this warehouse
      const { data: overview } = await api.get('/pallet-inventory/groups');
      const groups = Array.isArray(overview) ? overview : [];
      const header = ['Pallet Description','Total Pallet'];
      const rows = groups.map((g: any) => {
        const perWarehouse = Array.isArray(g.perWarehouse) ? g.perWarehouse : [];
        const rec = perWarehouse.find((p: any) => String(p.warehouseId) === String(id));
        const qty = Number(rec?.pallets || 0);
        return [g.groupName || '', qty];
      }).filter(r => r[1] > 0);
      const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Pallet Inventory');
      const d = new Date();
      const pad = (n: number) => n.toString().padStart(2,'0');
      const ts = `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
      const safeName = (name || 'Warehouse').replace(/[/\\:*?"<>|]/g, '-');
      const fname = `warehouse-pallet-inventory-(${safeName})-(${ts}).xlsx`;
      XLSX.writeFile(wb, fname);
    } catch (e:any) {
      toast.error(e?.response?.data?.message || 'Export failed');
    } finally {
      setExportingId('');
    }
  };

  function ActionsMenuCell({ row, disabled, onEdit, onDelete, onExport }: { row: any, disabled?: boolean, onEdit: ()=>void, onDelete: ()=>void, onExport: ()=>void }) {
    const [anchorEl, setAnchorEl] = useState<null | HTMLElement>(null);
    const open = Boolean(anchorEl);
    const handleOpen = (e: React.MouseEvent<HTMLElement>) => setAnchorEl(e.currentTarget);
    const handleClose = () => setAnchorEl(null);
    const wrap = (fn: ()=>void) => () => { fn(); handleClose(); };
    return (
      <>
        <IconButton size="small" onClick={handleOpen} aria-label="Actions"><MoreVertIcon fontSize="small" /></IconButton>
        <Menu anchorEl={anchorEl} open={open} onClose={handleClose} anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }} transformOrigin={{ vertical: 'top', horizontal: 'right' }}>
          <MenuItem onClick={wrap(onEdit)}>Edit</MenuItem>
          <MenuItem onClick={wrap(onDelete)} disabled={!!disabled}>Delete</MenuItem>
          <MenuItem onClick={wrap(onExport)} disabled={!!disabled}>Export Stock</MenuItem>
        </Menu>
      </>
    );
  }

  const columns: GridColDef[] = [
    { field: 'name', headerName: 'Name', flex: 1, minWidth: 180 },
    { field: 'address', headerName: 'Address', flex: 2, minWidth: 240 },
    { field: 'actions', headerName: 'Actions', width: 80, sortable: false, filterable: false, renderCell: (params: GridRenderCellParams) => {
      const r = params.row as any;
      return (
        <ActionsMenuCell
          row={r}
          disabled={exportingId === r._id}
          onEdit={()=>{ setEditId(r._id); setEditName(r.name || ''); setEditAddress(r.address || ''); setEditOpen(true); }}
          onDelete={async ()=>{ if (!window.confirm('Delete this warehouse?')) return; try { await api.delete(`/warehouses/${r._id}`); toast.success('Deleted'); load(); } catch(e:any){ toast.error(e?.response?.data?.message || 'Delete failed'); } }}
          onExport={()=>{ exportStock(r._id, r.name || ''); }}
        />
      );
    } }
  ];

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Warehouses</Typography>

      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs: 'column', md: 'row' }} spacing={2} alignItems="center">
          <TextField required label="Name" value={name} onChange={e=>setName(e.target.value)} sx={{ minWidth: 220 }} />
          <TextField label="Address" value={address} onChange={e=>setAddress(e.target.value)} sx={{ minWidth: 300, flex: 1 }} />
          <Button variant="contained" onClick={addWarehouse} disabled={!name.trim()}>Add Warehouse</Button>
          <Stack direction="row" spacing={1} sx={{ ml: 'auto' }}>
            <Button variant="outlined" onClick={exportAllStocks} disabled={exportingAll || !rows.length}>{exportingAll ? 'Exporting…' : 'Export All Stocks'}</Button>
          </Stack>
        </Stack>
      </Paper>

      <Paper sx={{ p:2 }}>
        <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} alignItems="center" sx={{ mb: 1 }}>
          <TextField size="small" label="Search" value={q} onChange={(e)=>setQ(e.target.value)} onKeyDown={(e)=>{ if (e.key==='Enter') load(); }} />
          <Button variant="outlined" onClick={load}>Search</Button>
        </Stack>
        <div style={{ height: 500, width: '100%' }}>
          <DataGrid
            rows={rows.map(r=>({ id: r._id, ...r }))}
            columns={columns}
            loading={loading}
            disableRowSelectionOnClick
            initialState={{ pagination: { paginationModel: { pageSize: 10, page: 0 } } }}
            pageSizeOptions={[5,10,50]}
          />
        </div>
      </Paper>
      <EditDialog open={editOpen} onClose={()=>setEditOpen(false)} id={editId} name={editName} address={editAddress} onSaved={load} />
    </Container>
  );
}

function EditDialog({ open, onClose, id, name, address, onSaved }: { open: boolean; onClose: ()=>void; id: string; name: string; address: string; onSaved: ()=>void }) {
  const [n, setN] = useState(name);
  const [a, setA] = useState(address);
  useEffect(()=>{ setN(name); setA(address); }, [name, address, open]);
  const save = async () => {
    try { await api.put(`/warehouses/${id}`, { name: n, address: a }); onSaved(); onClose(); } catch(e:any){}
  };
  return (
    <Dialog open={open} onClose={onClose} fullWidth maxWidth="sm">
      <DialogTitle>Edit Warehouse</DialogTitle>
      <DialogContent>
        <Stack spacing={2} sx={{ mt:1 }}>
          <TextField label="Name" value={n} onChange={e=>setN(e.target.value)} required />
          <TextField label="Address" value={a} onChange={e=>setA(e.target.value)} />
        </Stack>
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose}>Cancel</Button>
        <Button variant="contained" onClick={save} disabled={!n.trim()}>Save</Button>
      </DialogActions>
    </Dialog>
  );
}
