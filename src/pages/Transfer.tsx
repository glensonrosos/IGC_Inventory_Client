import { useEffect, useMemo, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, MenuItem } from '@mui/material';
import { useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import api from '../api';
import { useToast } from '../components/ToastProvider';
 

interface Warehouse { _id: string; name: string }
interface TransferPallet { groupName: string; pallets: number }

export default function Transfer() {
  const [warehouses, setWarehouses] = useState<Warehouse[]>([]);
  const [sourceWarehouseId, setSourceWarehouseId] = useState('');
  const [warehouseId, setWarehouseId] = useState('');
  const [items, setItems] = useState<TransferPallet[]>([]);
  const defaultEddYmd = useMemo(() => {
    const d = new Date();
    d.setMonth(d.getMonth() + 3);
    return d.toISOString().slice(0, 10);
  }, []);
  const [edd, setEdd] = useState(defaultEddYmd);
  const [poNumber, setPoNumber] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const toast = useToast();
  const navigate = useNavigate();
  const [available, setAvailable] = useState<Record<string, number>>({});
  const [file, setFile] = useState<File | null>(null);
  const [fileInputKey, setFileInputKey] = useState(0);
  const todayYmd = useMemo(() => new Date().toISOString().slice(0, 10), []);

  const loadWarehouses = async () => {
    try {
      const { data } = await api.get('/warehouses');
      setWarehouses(data || []);
    } catch {
      setWarehouses([]);
    }
  };

  useEffect(() => {
    const list = Array.isArray(warehouses) ? warehouses : [];
    if (!list.length) return;

    const primary = list.find((w: any) => Boolean((w as any)?.isPrimary));
    const second = list.find((w: any) => !Boolean((w as any)?.isPrimary));

    if (!sourceWarehouseId && second?._id) setSourceWarehouseId(String(second._id));
    if (!warehouseId && primary?._id) setWarehouseId(String(primary._id));
  }, [warehouses, sourceWarehouseId, warehouseId]);

  const loadStock = async (srcId: string) => {
    if (!srcId) { setAvailable({}); return; }
    try {
      const { data } = await api.get('/pallet-inventory/groups');
      const map: Record<string, number> = {};
      (Array.isArray(data) ? data : []).forEach((g: any) => {
        const groupName = String(g.groupName || '').trim();
        if (!groupName) return;
        const per = Array.isArray(g.perWarehouse) ? g.perWarehouse : [];
        const rec = per.find((p: any) => String(p.warehouseId) === String(srcId));
        map[groupName] = Number(rec?.pallets || 0);
      });
      setAvailable(map);
    } catch {
      setAvailable({});
    }
  };

  const resetImport = () => {
    setItems([]);
    setFile(null);
    setFileInputKey((k) => k + 1);
  };

  const refresh = async () => {
    setRefreshing(true);
    try {
      resetImport();
      await loadWarehouses();
      await loadStock(sourceWarehouseId);
    } finally {
      setRefreshing(false);
    }
  };

  useEffect(() => {
    refresh();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    loadStock(sourceWarehouseId);
  }, [sourceWarehouseId]);

  // Items are provided exclusively by .xlsx import

  const submit = async () => {
    if (!sourceWarehouseId || !warehouseId) { toast.error('Select both source and destination warehouses'); return; }
    if (sourceWarehouseId === warehouseId) { toast.error('Source and destination must be different'); return; }
    if (!poNumber.trim()) { toast.error('PO# is required'); return; }
    if (edd && String(edd) < todayYmd) { toast.error('Estimated Delivery cannot be earlier than today'); return; }
    const valid = items.filter(i => i.groupName && Number.isFinite(i.pallets) && i.pallets > 0);
    if (!valid.length) { toast.error('Import at least one valid pallet row (Pallet Description, Total Pallet)'); return; }
    // client-side availability check
    const insufficient = valid.filter(i => (available[i.groupName] || 0) < i.pallets);
    if (insufficient.length) { toast.error(`Insufficient pallet stock for: ${insufficient.map(i=>i.groupName).join(', ')}`); return; }
    setSubmitting(true);
    try {
      await api.post('/shipments/transfer-pallet', {
        sourceWarehouseId,
        warehouseId,
        pallets: valid.map(v => ({ groupName: v.groupName, pallets: v.pallets })),
        estDeliveryDate: edd || undefined,
        reference: poNumber.trim(),
      });
      toast.success('Transfer created and items moved to on-water');
      setItems([]);
      setPoNumber(''); setEdd(defaultEddYmd); setFile(null);
      navigate('/ship');
    } catch (e: any) {
      toast.error(e?.response?.data?.message || 'Failed to create transfer');
    } finally {
      setSubmitting(false);
    }
  };

  const importXlsx = async () => {
    if (!file) { toast.error('Select a .xlsx file'); return; }
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!rawRows.length) { toast.error('Empty worksheet'); return; }
      const expectedHeader = ['Pallet Description', 'Total Pallet'];
      const normHeader = (h: any) => String(h || '').trim().toLowerCase();
      const receivedHeader = Array.isArray(rawRows[0]) ? rawRows[0].map(normHeader) : [];
      const expectedHeaderNorm = expectedHeader.map(normHeader);
      const headerMatches = receivedHeader.length === expectedHeaderNorm.length
        && expectedHeaderNorm.every((h, i) => receivedHeader[i] === h);
      if (!headerMatches) {
        toast.error('Invalid template. Column headers must match the template exactly.');
        return;
      }

      const rows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' });
      const parsed: TransferPallet[] = [];
      for (const r of rows) {
        const groupName = String(r['Pallet Description'] ?? r['pallet description'] ?? r['Pallet Group'] ?? r['pallet group'] ?? r.groupName ?? r.GroupName ?? '').trim();
        const pallets = Number(r['Total Pallet'] ?? r['total pallet'] ?? r.pallets ?? r.Pallets ?? 0);
        if (!groupName || !Number.isFinite(pallets) || pallets <= 0) continue;
        parsed.push({ groupName, pallets });
      }
      if (!parsed.length) { toast.error('No valid rows found (need Pallet Description, Total Pallet)'); return; }
      setItems(parsed);
      toast.success(`Loaded ${parsed.length} items from file`);
    } catch (e:any) {
      toast.error('Failed to parse .xlsx');
    }
  };

  const downloadTemplate = () => {
    const header = ['Pallet Description','Total Pallet'];
    const ws = XLSX.utils.aoa_to_sheet([header, ['Inverted Planters Mixed Smooth and VA (OW / MB / C / DAB)', 1]]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'transfer_pallet_template.xlsx');
  };

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Transfer Pallet</Typography>
      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb:2 }}>
          <TextField select label="From Warehouse" size="small" sx={{ minWidth: 220 }} value={sourceWarehouseId} onChange={(e)=>setSourceWarehouseId(e.target.value)}>
            {warehouses.map(w => <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>)}
          </TextField>
          <TextField select label="To Warehouse" size="small" sx={{ minWidth: 220 }} value={warehouseId} onChange={(e)=>setWarehouseId(e.target.value)}>
            {warehouses.map(w => <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>)}
          </TextField>
          <TextField
            type="date"
            label="Estimated Delivery"
            size="small"
            sx={{ minWidth: 200 }}
            value={edd}
            inputProps={{ min: todayYmd }}
            onChange={(e)=> {
              const v = String(e.target.value || '');
              if (v && v < todayYmd) {
                toast.error('Estimated Delivery cannot be earlier than today');
                return;
              }
              setEdd(v);
            }}
            InputLabelProps={{ shrink: true }}
          />
          <Button variant="outlined" onClick={refresh} disabled={refreshing} sx={{ minWidth: 120 }}>
            Refresh
          </Button>
        </Stack>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb:2 }}>
          <TextField label="PO #" size="small" value={poNumber} onChange={(e)=>setPoNumber(e.target.value)} fullWidth />
        </Stack>
        <Stack direction={{ xs:'column', md:'row' }} spacing={2} alignItems="center" sx={{ mb:2 }}>
          <input key={fileInputKey} type="file" accept=".xlsx" onChange={(e)=> setFile(e.target.files?.[0] || null)} />
          <Button variant="outlined" onClick={importXlsx} disabled={!file}>Import .xlsx</Button>
          <Button variant="text" onClick={downloadTemplate}>Download Template</Button>
        </Stack>
        <Paper variant="outlined" sx={{ p:2 }}>
          <Typography variant="subtitle2" gutterBottom>Imported Pallets</Typography>
          <table style={{ width:'100%', borderCollapse:'collapse' }}>
            <thead>
              <tr>
                <th align="left">Pallet Description</th>
                <th align="right">Total Pallet</th>
                <th align="right">Available (source)</th>
                <th align="left">Status</th>
              </tr>
            </thead>
            <tbody>
              {items.map((it, idx) => {
                const avail = available[it.groupName] || 0;
                const ok = Number(it.pallets) > 0 && avail >= Number(it.pallets);
                return (
                  <tr key={idx} style={{ borderTop:'1px solid #eee', background: ok ? undefined : '#fff4f4' }}>
                    <td>{it.groupName}</td>
                    <td align="right">{it.pallets}</td>
                    <td align="right">{avail}</td>
                    <td style={{ color: ok ? '#2e7d32' : '#c62828' }}>{ok ? 'OK' : 'Insufficient'}</td>
                  </tr>
                );
              })}
              {!items.length && (
                <tr><td colSpan={4} style={{ color:'#666', padding:'8px 0' }}>No pallets loaded. Import a .xlsx file.</td></tr>
              )}
            </tbody>
          </table>
        </Paper>
        <Stack direction="row" spacing={1} sx={{ mt:2 }}>
          <Button variant="outlined" onClick={()=>{ setItems([]); setFile(null); }}>Clear</Button>
          <Button variant="contained" onClick={submit} disabled={submitting || !items.length || !poNumber.trim() || items.some(i=> (available[i.groupName]||0) < Number(i.pallets) || Number(i.pallets)<=0)}>Create Transfer</Button>
        </Stack>
        <Typography variant="body2" sx={{ display: 'block', mt: 1, color: 'error.main' }}>
          Note: After you create a transfer, you can view the transfer request in the Ship page.
        </Typography>
      </Paper>
      
    </Container>
  );
}
