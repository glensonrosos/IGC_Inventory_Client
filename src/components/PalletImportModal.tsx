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

export default function PalletImportModal({ open, onClose, onImported }: Props) {
  const [file, setFile] = useState<File | null>(null);
  const [fileInputKey, setFileInputKey] = useState(0);
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
    const header = ['PO #','Pallet Description','Total Pallet'];
    const ws = XLSX.utils.aoa_to_sheet([header]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'pallet_import_template.xlsx');
  };

  const resetImport = () => {
    setFile(null);
    setFileInputKey((k) => k + 1);
    setPreview(null);
    setErrors([]);
    setDups([]);
    setWarehouseId('');
    setEdd('');
    setStatus('Delivered');
  };

  const handleFileChange = (e: any) => {
    setFile(e.target.files?.[0] || null);
    setPreview(null);
    setErrors([]);
    setDups([]);
  };

  const doPreview = async () => {
    if (!file) return;
    if (!warehouseId) { toast.error('Please select a Warehouse'); return; }
    setLoading(true);
    setErrors([]);
    try {
      const form = new FormData();
      form.append('file', file);
      const params = new URLSearchParams();
      params.set('warehouseId', warehouseId);
      const { data } = await api.post(`/pallet-inventory/import/preview?${params.toString()}`, form, { headers: { 'Content-Type': 'multipart/form-data' } });
      setPreview(data);
      setErrors(data.errors || []);
      // duplicates already counted by controller; keep placeholder if we later want to list
      setDups([]);
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
      params.set('status', status);
      params.set('warehouseId', warehouseId);
      if (status === 'On-Water' && edd) params.set('estDeliveryDate', edd);
      const { data } = await api.post(`/pallet-inventory/import?${params.toString()}`, form, { headers: { 'Content-Type': 'multipart/form-data' } });
      if (data?.errorCount) {
        setErrors(data.errors || []);
        toast.error('Some rows failed. Successful rows were committed.');
      } else {
        toast.success('Import committed');
      }
      onImported();
      onClose();
      setFile(null);
      setPreview(null);
      setErrors([]);
      setDups([]);
      setWarehouseId('');
      setEdd('');
      setStatus('Delivered');
    } catch (e: any) {
      setErrors([{ rowNum: '-', errors: [e?.response?.data?.message || 'Import failed'] }]);
      toast.error(e?.response?.data?.message || 'Import failed');
    } finally {
      setCommitting(false);
    }
  };

  return (
    <Dialog open={open} onClose={onClose} fullWidth maxWidth="md">
      <DialogTitle>IMPORT EXISTING PALLET INVENTORY ONLY</DialogTitle>
      <DialogContent>
        <Box sx={{ my: 2, display:'flex', gap:1, alignItems:'center' }}>
          <input key={fileInputKey} type="file" accept=".xlsx" onChange={handleFileChange} />
          <Button size="small" variant="text" onClick={downloadTemplate}>Download Template</Button>
          <Button size="small" variant="outlined" onClick={resetImport} disabled={loading || committing}>Refresh</Button>
          {loading && <LinearProgress sx={{ mt: 2 }} />}
        </Box>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb:2 }}>
          <TextField select label="Status" size="small" sx={{ minWidth: 160 }} value={status} onChange={(e)=>setStatus(e.target.value as any)}>
            <MenuItem value="Delivered">Delivered</MenuItem>
          </TextField>
          <TextField select label="Warehouse" size="small" sx={{ minWidth: 220 }} value={warehouseId} onChange={(e)=>setWarehouseId(e.target.value)} error={!warehouseId} helperText={!warehouseId ? 'Required' : ''}>
            {warehouses.map((w:any)=> (
              <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>
            ))}
          </TextField>
        </Stack>
        <Box sx={{ display:'flex', gap:1, mb:2 }}>
          <Button variant="outlined" disabled={!file || loading} onClick={doPreview}>Preview</Button>
          <Button variant="contained" disabled={!file || !preview || errors.length > 0 || committing || !warehouseId} onClick={doCommit}>Commit</Button>
        </Box>
        {preview && (
          <Typography variant="body2" sx={{ mb: 1 }}>
            Parsed rows: {preview.totalRows} | Errors: {preview.errorCount} | Duplicates: {preview.duplicateCount || 0}
          </Typography>
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
