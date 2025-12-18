import { useEffect, useMemo, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, IconButton, MenuItem, Chip } from '@mui/material';
import * as XLSX from 'xlsx';
import { DataGrid, GridColDef, GridRenderCellParams } from '@mui/x-data-grid';
import api from '../api';
import { useToast } from '../components/ToastProvider';
import UpdateIcon from '@mui/icons-material/Update';
import LocalShippingIcon from '@mui/icons-material/LocalShipping';
import LockIcon from '@mui/icons-material/Lock';

interface ShipmentItem { itemCode: string; qtyPieces: number; packSize?: number }
interface Shipment {
  _id: string;
  kind: 'import' | 'transfer';
  status: 'on_water' | 'delivered';
  reference?: string;
  warehouseId: string;
  sourceWarehouseId?: string;
  estDeliveryDate?: string;
  items: ShipmentItem[];
  createdAt: string;
}

export default function Ship() {
  const [rows, setRows] = useState<Shipment[]>([]);
  const [loading, setLoading] = useState(false);
  const [filter, setFilter] = useState<'on_water' | 'delivered' | ''>('');
  const [eddEdit, setEddEdit] = useState<Record<string, string>>({});
  const toast = useToast();
  const [warehouses, setWarehouses] = useState<any[]>([]);
  const [page, setPage] = useState(0);
  const [pageSize, setPageSize] = useState(10);
  const [rowCount, setRowCount] = useState(0);
  const [q, setQ] = useState('');
  const [exporting, setExporting] = useState(false);
  const [exportingRowId, setExportingRowId] = useState<string>('');
  const nameOf = useMemo(()=>{
    const map: Record<string,string> = {};
    for (const w of warehouses) map[String((w as any)._id)] = (w as any).name;
    return (id?: any) => {
      if (!id) return '-';
      const key = typeof id === 'string' ? id : (id?._id ? String(id._id) : String(id));
      return map[key] || '-';
    };
  }, [warehouses]);

  

  const load = async () => {
    setLoading(true);
    try {
      const params: any = { page, pageSize };
      if (filter) params.status = filter;
      if (q.trim()) params.q = q.trim();
      const { data } = await api.get('/shipments', { params });
      setRows((data?.items || []) as Shipment[]);
      setRowCount(Number(data?.total || 0));
    } catch (e:any) {
      setRows([]);
      setRowCount(0);
      toast.error(e?.response?.data?.message || 'Failed to load shipments');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, [filter, page, pageSize]);

  useEffect(() => {
    const loadWarehouses = async () => {
      try { const { data } = await api.get('/warehouses'); setWarehouses(data || []); } catch {}
    };
    loadWarehouses();
  }, []);

  // When warehouses list changes, force the grid rows to re-render so nameOf is applied
  useEffect(() => {
    setRows(prev => prev.slice());
  }, [warehouses]);

  // Global export removed: CSV export option has been deprecated.

  const deliver = async (id: string) => {
    if (!window.confirm('Mark this shipment as Delivered?')) return;
    try {
      const { data } = await api.post(`/shipments/${id}/deliver`);
      toast.success('Shipment delivered');
      try {
        const items = (data?.shipment?.items || []) as any[];
        const itemCodes = Array.from(new Set(items.map((i:any)=> i.itemCode).filter(Boolean)));
        window.dispatchEvent(new CustomEvent('shipment-delivered', { detail: { itemCodes } }));
      } catch {}
      load();
    } catch (e: any) {
      toast.error(e?.response?.data?.message || 'Failed to deliver');
    }
  };

  const updateEDD = async (id: string) => {
    try {
      const date = eddEdit[id];
      await api.put(`/shipments/${id}/edd`, { estDeliveryDate: date });
      toast.success('EDD updated');
      load();
    } catch (e: any) {
      toast.error(e?.response?.data?.message || 'Failed to update EDD');
    }
  };

  const columns: GridColDef[] = [
    { field: 'reference', headerName: 'Reference', width: 220, renderCell: (params: GridRenderCellParams) => {
      const r: any = params?.row || {};
      const ref = r.reference || r._id || '-';
      // For transfer created from Transfer page, display PO formatting
      if (r.kind === 'transfer' && r.reference) {
        return <span>{`PO - ${r.reference}`}</span>;
      }
      // For import created from On-Process, keep PO formatting
      const text = r.kind === 'import' && r.reference ? `PO - ${r.reference}` : ref;
      return <span>{text}</span>;
    } },
    { field: 'kind', headerName: 'Kind', width: 110, renderCell: (params: GridRenderCellParams) => {
      const r: any = params?.row || {};
      // If import came from On-Process flow, label as 'on-process'
      const fromOnProcess = r.kind === 'import' && typeof r.notes === 'string' && r.notes.toLowerCase().includes('on-process');
      return <span>{fromOnProcess ? 'on-process' : r.kind}</span>;
    } },
    { field: 'status', headerName: 'Status', width: 120, renderCell: (params: GridRenderCellParams) => {
      const r: any = params?.row || {};
      const status = String(r.status || '').toLowerCase();
      let color: any = 'default';
      if (status === 'on_water') color = 'warning';
      else if (status === 'delivered') color = 'success';
      else if (status === 'transferred') color = 'primary';
      return <Chip size="small" color={color} label={r.status || '-'} />;
    } },
    { field: 'warehouseId', headerName: 'To Warehouse', width: 200, renderCell:(params: GridRenderCellParams)=> {
      const w: any = (params?.row as any)?.warehouseId;
      const text = !w ? 'NA' : (typeof w === 'object' && w?.name) ? w.name : nameOf(w);
      return <span>{text}</span>;
    } },
    { field: 'sourceWarehouseId', headerName: 'From Warehouse', width: 200, renderCell:(params: GridRenderCellParams)=> {
      const w: any = (params?.row as any)?.sourceWarehouseId;
      const text = !w ? 'NA' : (typeof w === 'object' && w?.name) ? w.name : nameOf(w);
      return <span>{text}</span>;
    } },
    { field: 'items', headerName: 'Pallet', flex: 1, minWidth: 180, renderCell: (params: GridRenderCellParams) => {
      const row: any = params?.row || {};
      const list: any[] = Array.isArray(row?.items) ? row.items : [];
      const exportItems = async () => {
        try {
          setExportingRowId(row._id || '');
          // Build worksheet: metadata + list of pallets
          const toName = (typeof row.warehouseId === 'object' && row.warehouseId?.name) ? row.warehouseId.name : nameOf(row.warehouseId);
          const fromName = (typeof row.sourceWarehouseId === 'object' && row.sourceWarehouseId?.name) ? row.sourceWarehouseId.name : nameOf(row.sourceWarehouseId);
          const edd = row?.estDeliveryDate ? String(row.estDeliveryDate).substring(0,10) : '';
          const isFromOnProcess = row.kind === 'import' && typeof row.notes === 'string' && row.notes.toLowerCase().includes('on-process');
          const exportKind = isFromOnProcess ? 'on-process' : (row.kind || '');
          const metaRows = [
            ['Reference', row.reference || row._id || ''],
            ['Kind', exportKind],
            ['Status', row.status || ''],
            ['To Warehouse', toName || ''],
            ['From Warehouse', fromName || ''],
            ['Estimated Date', edd],
          ];
          let header: string[] = ['Pallet Description','Pallet Qty'];
          let body: any[] = [];
          // If this shipment came from On-Process page, derive from On-Process pallets by PO# (reference)
          if (isFromOnProcess) {
            try {
              const text: string = String(row.notes || '');
              const re = /pallet-group:([^;|]+);\s*pallets:(\d+)/g;
              let m: RegExpExecArray | null;
              while ((m = re.exec(text)) !== null) {
                const groupName = (m[1] || '').trim();
                const qty = Number(m[2] || 0);
                if (groupName && Number.isFinite(qty) && qty > 0) body.push([groupName, qty]);
              }
            } catch {}
          }
          // If this is a pallet transfer shipment, parse pallet-group segments from notes
          const isPalletTransfer = row.kind === 'transfer' && typeof row.notes === 'string' && row.notes.toLowerCase().includes('pallet-group:');
          if (!body.length && isPalletTransfer) {
            try {
              const text: string = String(row.notes || '');
              const re = /pallet-group:([^;|]+);\s*pallets:(\d+)/g;
              let m: RegExpExecArray | null;
              while ((m = re.exec(text)) !== null) {
                const groupName = (m[1] || '').trim();
                const qty = Number(m[2] || 0);
                if (groupName && Number.isFinite(qty) && qty > 0) body.push([groupName, qty]);
              }
            } catch {}
          }
          // Fallback: use shipment items metadata if On-Process lookup did not populate
          if (!body.length) {
            const codes = Array.from(new Set(list.map((i:any)=> i.itemCode).filter(Boolean)));
            const details: Record<string, any> = {};
            await Promise.all(codes.map(async (code) => {
              try {
                const { data } = await api.get(`/items/${encodeURIComponent(code)}`);
                details[code] = data || {};
              } catch {}
            }));
            body = list.map(it => {
              const meta = details[it.itemCode] || {};
              const palletGroup = meta.itemGroup || '';
              const palletQty = it.qtyPieces ?? '';
              return [palletGroup, palletQty];
            });
          }
          const ws = XLSX.utils.aoa_to_sheet([...metaRows, [], header, ...body]);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, 'Pallets');
          const pad = (n:number)=> n.toString().padStart(2,'0');
          const d = new Date();
          const ts = `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
          const fname = `shipment-pallets-${row.reference || row._id}-(${row.status || ''})-(${toName})-(${ts}).xlsx`;
          XLSX.writeFile(wb, fname);
        } finally {
          setExportingRowId('');
        }
      };
      return (
        <Stack direction="row" spacing={1} alignItems="center">
          <Button size="small" variant="outlined" onClick={exportItems} disabled={exportingRowId === (row._id||'')}>EXPORT PALLET</Button>
        </Stack>
      );
    } },
    { field: 'estDeliveryDate', headerName: 'Estimated Date Delivery', width: 300, renderCell: (params: GridRenderCellParams) => {
      const row: any = params?.row || {};
      const id = row?._id as string;
      if (!id) return null;
      const current = row?.estDeliveryDate ? String(row.estDeliveryDate).substring(0,10) : '';
      const val = eddEdit[id] ?? current;
      const delivered = String(row?.status) === 'delivered';
      const transferred = String(row?.status) === 'transferred';
      if (delivered) {
        return (
          <Stack direction="row" spacing={1} alignItems="center">
            <TextField type="date" size="small" value={current} InputLabelProps={{ shrink: true }} InputProps={{ readOnly: true }} disabled />
            <Chip size="small" color="success" icon={<LockIcon fontSize="small" />} label="Delivered" />
          </Stack>
        );
      }
      if (transferred) {
        return (
          <Stack direction="row" spacing={1} alignItems="center">
            <TextField type="date" size="small" value={current} InputLabelProps={{ shrink: true }} InputProps={{ readOnly: true }} disabled />
            <Chip size="small" color="primary" label="Transfered" />
          </Stack>
        );
      }
      return (
        <Stack direction="row" spacing={1} alignItems="center">
          <TextField type="date" size="small" value={val} onChange={(e)=> setEddEdit(prev=>({ ...prev, [id]: e.target.value }))} InputLabelProps={{ shrink: true }} />
          <Button size="small" variant="outlined" startIcon={<UpdateIcon fontSize="small" />} onClick={()=>{ if (!val) { toast.error('EDD is required'); return; } updateEDD(id); }} disabled={!val}>Update</Button>
        </Stack>
      );
    }},
    { field: 'actions', headerName: 'Actions', width: 160, renderCell: (params: GridRenderCellParams) => {
      const row: any = params?.row || {};
      const edd = row?.estDeliveryDate ? String(row.estDeliveryDate).substring(0,10) : '';
      return (
        <Stack direction="row" spacing={1}>
          {row?.status === 'on_water' && row?._id ? (
            <Button size="small" startIcon={<LocalShippingIcon />} disabled={!edd} onClick={()=>{
              if (!edd) { toast.error('Set EDD first'); return; }
              const d = new Date(`${edd}T00:00:00`);
              const now = new Date();
              const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
              const eddDateOnly = new Date(d.getFullYear(), d.getMonth(), d.getDate());
              if (eddDateOnly.getTime() > startOfToday.getTime()) {
                toast.error('EDD cannot be in the future');
                return;
              }
              const ok = window.confirm(`Mark this shipment as Delivered with delivery date ${edd}?`);
              if (!ok) return;
              deliver(row._id);
            }}>Deliver</Button>
          ) : null}
        </Stack>
      );
    }}
  ];

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Shipments</Typography>
      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems="center">
          <TextField select size="small" label="Status" value={filter} onChange={(e)=> setFilter(e.target.value as any)} sx={{ minWidth: 180 }}>
            <MenuItem value="">All</MenuItem>
            <MenuItem value="on_water">On Water</MenuItem>
            <MenuItem value="delivered">Delivered</MenuItem>
          </TextField>
          <TextField size="small" label="Search (Ref/Item Code)" value={q} onChange={(e)=> setQ(e.target.value)} onKeyDown={(e)=>{ if (e.key==='Enter') { setPage(0); load(); } }} sx={{ minWidth: 260 }} />
          <Button variant="outlined" onClick={()=>{ setPage(0); load(); }}>Search</Button>
        </Stack>
      </Paper>
      <Paper sx={{ p:1 }}>
        <div style={{ height: 560, width: '100%' }}>
          <DataGrid
            rows={rows.map(r=>({ id: r._id, ...r }))}
            columns={columns}
            loading={loading}
            disableRowSelectionOnClick
            paginationMode="server"
            rowCount={rowCount}
            paginationModel={{ page, pageSize }}
            onPaginationModelChange={(m)=>{ setPage(m.page); setPageSize(m.pageSize); }}
            pageSizeOptions={[5,10,50]}
          />
        </div>
      </Paper>
    </Container>
  );
}
