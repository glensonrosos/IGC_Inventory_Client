import { useEffect, useState } from 'react';
import { Dialog, DialogTitle, DialogContent, DialogActions, Button, Typography, Box, LinearProgress, MenuItem, TextField, Stack } from '@mui/material';
import * as XLSX from 'xlsx';
import api from '../api';
import { useToast } from './ToastProvider';

type Props = {
  open: boolean;
  onClose: () => void;
  onImported: () => void;
};

export default function ImportModal({ open, onClose, onImported }: Props) {
  const [file, setFile] = useState<File | null>(null);
  const [preview, setPreview] = useState<any>(null);
  const [errors, setErrors] = useState<any[]>([]);
  const [dups, setDups] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [committing, setCommitting] = useState(false);
  const [status, setStatus] = useState<'Delivered' | 'On-Water'>('Delivered');
  const [warehouses, setWarehouses] = useState<any[]>([]);
  const [warehouseId, setWarehouseId] = useState<string>('');
  const [edd, setEdd] = useState<string>(''); // yyyy-mm-dd
  const toast = useToast();

  useEffect(() => {
    if (!open) return;
    const loadWarehouses = async () => {
      try {
        const { data } = await api.get('/warehouses');
        setWarehouses(data || []);
      } catch (e:any) {
        // silent
      }
    };
    loadWarehouses();
  }, [open]);

  const downloadTemplate = () => {
    const header = ['PO #','Item Code','Total Qty','Pack Size'];
    const ws = XLSX.utils.aoa_to_sheet([header]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'import_stock_template.xlsx');
  };

  const handleFileChange = (e: any) => {
    setFile(e.target.files?.[0] || null);
    setPreview(null);
    setErrors([]);
  };

  const doPreview = async () => {
    if (!file) return;
    setLoading(true);
    setErrors([]);
    try {
      const form = new FormData();
      form.append('file', file);
      const { data } = await api.post('/items/import/preview', form, { headers: { 'Content-Type': 'multipart/form-data' } });
      setPreview(data);
      setErrors(data.errors || []);
      setDups(data.duplicates || []);
      toast.success('Preview parsed successfully');
    } catch (e: any) {
      setErrors([{ rowNum: '-', errors: [e?.response?.data?.message || 'Preview failed'] }]);
      toast.error(e?.response?.data?.message || 'Preview failed');
    } finally {
      setLoading(false);
    }
  };

  const doCommit = async () => {
    if (!file) return;
    if (!warehouseId) { toast.error('Please select a Warehouse'); return; }
    if (status === 'On-Water' && !edd) { toast.error('Please select Estimated Delivery Date'); return; }
    setCommitting(true);
    try {
      const form = new FormData();
      form.append('file', file);
      const params = new URLSearchParams();
      params.set('type','stock_in');
      params.set('status', status);
      params.set('warehouseId', warehouseId);
      if (status === 'On-Water' && edd) params.set('estDeliveryDate', edd);
      await api.post(`/items/import?${params.toString()}`, form, { headers: { 'Content-Type': 'multipart/form-data' } });
      onImported();
      onClose();
      setFile(null);
      setPreview(null);
      setErrors([]);
      setDups([]);
      setWarehouseId('');
      setEdd('');
      setStatus('Delivered');
      toast.success('Import committed');
    } catch (e: any) {
      setErrors([{ rowNum: '-', errors: [e?.response?.data?.message || 'Import failed'] }]);
      toast.error(e?.response?.data?.message || 'Import failed');
    } finally {
      setCommitting(false);
    }
  };

  return (
    <Dialog open={open} onClose={onClose} fullWidth maxWidth="md">
      <DialogTitle>Import Stock (Excel .xlsx)</DialogTitle>
      <DialogContent>
        <Box sx={{ my: 2, display:'flex', gap:1, alignItems:'center' }}>
          <input type="file" accept=".xlsx" onChange={handleFileChange} />
          <Button size="small" variant="text" onClick={downloadTemplate}>Download Template</Button>
          {loading && <LinearProgress sx={{ mt: 2 }} />}
        </Box>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb:2 }}>
          <TextField select label="Status" size="small" sx={{ minWidth: 160 }} value={status} onChange={(e)=>setStatus(e.target.value as any)}>
            <MenuItem value="Delivered">Delivered</MenuItem>
            <MenuItem value="On-Water">On-Water</MenuItem>
          </TextField>
          <TextField select label="Warehouse" size="small" sx={{ minWidth: 220 }} value={warehouseId} onChange={(e)=>setWarehouseId(e.target.value)} error={!warehouseId} helperText={!warehouseId ? 'Required' : ''}>
            {warehouses.map((w:any)=> (
              <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>
            ))}
          </TextField>
          <TextField type="date" label="Estimated Delivery" size="small" sx={{ minWidth: 200 }} value={edd} onChange={(e)=>setEdd(e.target.value)} InputLabelProps={{ shrink: true }} disabled={status !== 'On-Water'} error={status==='On-Water' && !edd} helperText={status==='On-Water' && !edd ? 'Required for On-Water' : ''} />
        </Stack>
        <Box sx={{ display:'flex', gap:1, mb:2 }}>
          <Button variant="outlined" disabled={!file || loading} onClick={doPreview}>Preview</Button>
          <Button variant="contained" disabled={!file || !preview || errors.length > 0 || committing || !warehouseId || (status==='On-Water' && !edd)} onClick={doCommit}>Commit (Stock IN)</Button>
        </Box>
        {preview && (
          <Typography variant="body2" sx={{ mb: 1 }}>
            Parsed rows: {preview.totalRows} | Errors: {preview.errorCount} | Duplicates: {preview.duplicateCount || 0}
          </Typography>
        )}
        {dups.length > 0 && (
          <Box sx={{ p:1, bgcolor:'#fff8e1', border:'1px solid #ffe0b2', borderRadius:1, mb:1 }}>
            <Typography variant="subtitle2">Duplicate rows by PO# + Item Code</Typography>
            <ul>
              {dups.slice(0,10).map((d: any, i: number) => (
                <li key={i}><Typography variant="caption">{d.poNumber} - {d.itemCode}</Typography></li>
              ))}
            </ul>
            {dups.length > 10 && <Typography variant="caption">...and {dups.length - 10} more</Typography>}
          </Box>
        )}
        {errors.length > 0 && (
          <Box sx={{ p:1, bgcolor:'#fff4f4', border:'1px solid #f5c2c7', borderRadius:1 }}>
            <Typography variant="subtitle2" color="error">Validation Errors</Typography>
            <ul>
              {errors.map((er: any, idx: number) => (
                <li key={idx}><Typography variant="caption">Row {er.rowNum}: {(er.errors||[]).join(', ')}</Typography></li>
              ))}
            </ul>
          </Box>
        )}
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose}>Close</Button>
      </DialogActions>
    </Dialog>
  );
}
