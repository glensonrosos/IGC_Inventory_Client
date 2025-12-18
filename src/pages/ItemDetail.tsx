import { useEffect, useMemo, useState } from 'react';
import { useParams } from 'react-router-dom';
import { Container, Typography, Paper, Stack, TextField, Button, MenuItem } from '@mui/material';
import { DataGrid, GridColDef } from '@mui/x-data-grid';
import api from '../api';

interface Item {
  _id: string;
  itemCode: string;
  itemGroup: string;
  description: string;
  color: string;
  totalQty: number;
  packSize?: number;
  packsOnHand?: number;
  palletsOnHand?: number;
  lowStockThreshold: number;
}

interface MovementItem { itemCode: string; qtyPieces: number; packSize: number; palletId?: string }
interface Txn { _id: string; type: string; reference: string; items: MovementItem[]; createdAt: string; notes?: string }

export default function ItemDetail() {
  const { itemCode = '' } = useParams();
  const [item, setItem] = useState<Item | null>(null);
  const [threshold, setThreshold] = useState<number>(0);
  const [txns, setTxns] = useState<Txn[]>([]);
  const [saving, setSaving] = useState(false);
  const [warehouses, setWarehouses] = useState<any[]>([]);
  const [warehouseId, setWarehouseId] = useState<string>('');
  const [warehouseQty, setWarehouseQty] = useState<number>(0);
  const [onWaterQty, setOnWaterQty] = useState<number>(0);
  const [warehouseByRef, setWarehouseByRef] = useState<Record<string,string>>({});

  const nameOf = useMemo(()=>{
    const map: Record<string,string> = {};
    for (const w of warehouses) map[String(w._id)] = w.name;
    return (id?: any) => {
      if (!id) return '-';
      const key = typeof id === 'string' ? id : (id?._id ? String(id._id) : String(id));
      return map[key] || '-';
    };
  }, [warehouses]);

  const load = async () => {
    const { data } = await api.get(`/items/${itemCode}`);
    setItem(data);
    setThreshold(data.lowStockThreshold || 0);
    const hist = await api.get('/transactions', { params: { itemCode, limit: 200 } });
    setTxns(hist.data.items || []);
    // load warehouses and default selection
    const ws = await api.get('/warehouses');
    setWarehouses(ws.data || []);
    const firstId = ws.data?.[0]?._id || '';
    setWarehouseId((prev)=> prev || firstId);
    // on-water total for this item across shipments (use server pagination API)
    let totalOnWater = 0;
    let page = 0;
    const pageSize = 200;
    // fetch pages until all matching shipments are summed (limit to 25 pages as safety)
    for (let tries = 0; tries < 25; tries++) {
      const { data } = await api.get('/shipments', { params: { status: 'on_water', q: itemCode, page, pageSize } });
      const items = Array.isArray(data?.items) ? data.items : [];
      for (const s of items) {
        const arr = Array.isArray(s?.items) ? s.items : [];
        totalOnWater += arr.filter((it: any) => it.itemCode === itemCode).reduce((ss: number, it: any) => ss + (Number(it.qtyPieces) || 0), 0);
      }
      const total = Number(data?.total || 0);
      const fetched = (page + 1) * pageSize;
      if (fetched >= total || items.length === 0) break;
      page += 1;
    }
    setOnWaterQty(totalOnWater);
  };

  const save = async () => {
    if (!item) return;
    setSaving(true);
    try {
      await api.put(`/items/${item.itemCode}`, { lowStockThreshold: Number(threshold) });
      await load();
    } finally {
      setSaving(false);
    }
  };

  useEffect(() => { if (itemCode) load(); }, [itemCode]);

  useEffect(() => {
    const handler = (e: any) => {
      const codes: string[] = e?.detail?.itemCodes || [];
      if (codes.includes(itemCode)) {
        load();
      }
    };
    window.addEventListener('shipment-delivered', handler as any);
    return () => window.removeEventListener('shipment-delivered', handler as any);
  }, [itemCode]);

  useEffect(() => {
    const loadWarehouseQty = async () => {
      if (!warehouseId || !itemCode) { setWarehouseQty(0); return; }
      const { data } = await api.get('/warehouse-stock', { params: { warehouseId } });
      const rec = (data || []).find((r: any) => r.itemCode === itemCode);
      setWarehouseQty(rec?.qtyPieces || 0);
    };
    loadWarehouseQty();
  }, [warehouseId, itemCode]);

  // Enrich warehouse names for transactions that came from On-Water -> Delivered
  useEffect(() => {
    const enrich = async () => {
      const refs: string[] = Array.from(new Set((txns||[])
        .filter((t: Txn) => (t.notes||'').toLowerCase().includes('on-water'))
        .map((t: Txn) => String(t.reference))
        .filter((v: string) => v && v.length > 0)));
      const missing = refs.filter((r: string) => !(r in warehouseByRef));
      if (!missing.length) return;
      const updates: Record<string,string> = {};
      for (const ref of missing) {
        try {
          const { data } = await api.get('/shipments', { params: { q: ref, page: 0, pageSize: 5 } });
          const list = Array.isArray(data?.items) ? data.items : [];
          const hit = list.find((s:any) => s.reference === ref);
          if (hit) {
            const w = hit.warehouseId;
            const name = (typeof w === 'object' && w?.name) ? w.name : nameOf(w);
            updates[ref] = name || '-';
          }
        } catch {}
      }
      if (Object.keys(updates).length) setWarehouseByRef(prev => ({ ...prev, ...updates }));
    };
    enrich();
  }, [txns, warehouseByRef, nameOf]);

  if (!item) return <Container sx={{ mt: 4 }}><Typography>Loading...</Typography></Container>;

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Item: {item.itemCode}</Typography>

      <Paper sx={{ p:2, mb:2 }}>
        <Typography variant="h6" gutterBottom>Info</Typography>
        <Stack spacing={2}>
          <TextField label="Item Group" value={item.itemGroup} InputProps={{ readOnly: true }} />
          <TextField label="Description" value={item.description} InputProps={{ readOnly: true }} />
          <TextField label="Color" value={item.color} InputProps={{ readOnly: true }} />
          <Stack direction={{ xs: 'column', md: 'row' }} spacing={2}>
            <TextField type="number" label="Total Qty" value={item.totalQty} InputProps={{ readOnly: true }} />
            <TextField type="number" label="Low Stock Threshold (pieces)" value={threshold} InputProps={{ readOnly: true }} />
          </Stack>
          <Stack direction={{ xs: 'column', md: 'row' }} spacing={2}>
            <TextField select label="Warehouse" value={warehouseId} onChange={(e)=>setWarehouseId(e.target.value)} sx={{ minWidth: 240 }}>
              {warehouses.map((w:any)=> (<MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>))}
            </TextField>
            <TextField type="number" label="Selected Warehouse Qty" value={warehouseQty} InputProps={{ readOnly: true }} />
            <TextField type="number" label="On-Water Qty (all)" value={onWaterQty} InputProps={{ readOnly: true }} />
          </Stack>
        </Stack>
      </Paper>

      <Paper sx={{ p:2 }}>
        <Typography variant="h6" gutterBottom>Recent Transactions</Typography>
        <div style={{ height: 520, width: '100%' }}>
          <DataGrid
            rows={txns.map(t=>{
              const qty = (t.items||[]).filter(i=>i.itemCode===item.itemCode).reduce((s,i)=>s+(i.qtyPieces||0),0);
              const notesLc = (t.notes||'').toLowerCase();
              const isTransfer = (t.reference||'').startsWith('TRANS -') || notesLc.includes('transfer');
              const isImportDelivered = !isTransfer && ((t.type||'').toUpperCase()==='IN') && notesLc.includes('delivered');
              const refDisp = isTransfer
                ? (t.reference || '').replace(/^TRANS -\s*/,'')
                : (isImportDelivered && t.reference ? `PO - ${t.reference}` : (t.reference || ''));
              let normNotes = t.notes || '';
              if (notesLc === 'import|delivered') normNotes = 'on-water|Delivered';
              if (notesLc === 'transfer|on-water') normNotes = 'transfer | on-water';
              if (notesLc === 'on-water|transfered') normNotes = 'on-water | Transfered';
              let wh = '-';
              const w = (t as any).warehouseId;
              if (w) {
                wh = (typeof w === 'object' && (w as any).name) ? (w as any).name : nameOf(w);
              } else {
                wh = warehouseByRef[t.reference] || '-';
              }
              return ({
                id: t._id,
                date: new Date(t.createdAt).toLocaleString(),
                type: t.type,
                reference: refDisp,
                warehouse: wh,
                qty,
                notes: normNotes
              });
            })}
            columns={[
              { field: 'date', headerName: 'Date', flex: 1.2, minWidth: 180 },
              { field: 'type', headerName: 'Type', width: 120 },
              { field: 'reference', headerName: 'Reference', width: 180 },
              { field: 'warehouse', headerName: 'Warehouse', width: 180 },
              { field: 'qty', headerName: 'Qty', width: 120 },
              { field: 'notes', headerName: 'Notes', flex: 1, minWidth: 220 },
            ] as GridColDef[]}
            disableRowSelectionOnClick
            initialState={{ pagination: { paginationModel: { pageSize: 10, page: 0 } } }}
            pageSizeOptions={[5,10,50]}
          />
        </div>
      </Paper>
    </Container>
  );
}
