import { useEffect, useState } from 'react';
import { Dialog, DialogTitle, DialogContent, DialogActions, Button, TextField, Stack, IconButton, Typography, Box } from '@mui/material';
import DeleteIcon from '@mui/icons-material/Delete';
import api from '../api';

type Composition = { itemCode: string; packSize: number; packs: number; pieces?: number };

type Props = {
  palletId: string | null;
  open: boolean;
  onClose: () => void;
  onSaved: () => void;
};

export default function PalletCompositionModal({ palletId, open, onClose, onSaved }: Props) {
  const [rows, setRows] = useState<Composition[]>([]);
  const [warehouseLocation, setLocation] = useState('');
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);

  const load = async () => {
    if (!palletId) return;
    setLoading(true);
    try {
      const { data } = await api.get('/pallets', { params: { q: palletId } });
      const found = (data || []).find((p: any) => p.palletId === palletId);
      if (found) {
        setRows((found.composition || []).map((c: any) => ({ itemCode: c.itemCode || '', packSize: c.packSize || 0, packs: c.packs || 0, pieces: c.pieces })));
        setLocation(found.warehouseLocation || '');
      }
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { if (open && palletId) load(); }, [open, palletId]);

  const addRow = () => setRows([...rows, { itemCode: '', packSize: 0, packs: 0 }]);
  const delRow = (i: number) => setRows(rows.filter((_, idx) => idx !== i));
  const upd = (i: number, patch: Partial<Composition>) => {
    const copy = [...rows];
    copy[i] = { ...copy[i], ...patch } as Composition;
    setRows(copy);
  };

  const save = async () => {
    if (!palletId) return;
    setSaving(true);
    try {
      await api.put(`/pallets/${palletId}/composition`, { composition: rows, warehouseLocation });
      onSaved();
      onClose();
    } finally {
      setSaving(false);
    }
  };

  return (
    <Dialog open={open} onClose={onClose} fullWidth maxWidth="md">
      <DialogTitle>Edit Pallet Composition</DialogTitle>
      <DialogContent>
        <Stack spacing={2} sx={{ mt: 1 }}>
          <TextField label="Warehouse Location" value={warehouseLocation} onChange={e=>setLocation(e.target.value)} />
          <Box>
            <Typography variant="subtitle1" gutterBottom>Items on Pallet</Typography>
            <Stack spacing={1}>
              {rows.map((r, i) => (
                <Stack key={i} direction={{ xs: 'column', md: 'row' }} spacing={1} alignItems="center">
                  <TextField label="Item Code" value={r.itemCode} onChange={e=>upd(i,{ itemCode: e.target.value })} sx={{ minWidth: 160 }} />
                  <TextField type="number" label="Pack Size" value={r.packSize} onChange={e=>upd(i,{ packSize: Number(e.target.value) })} sx={{ width: 140 }} />
                  <TextField type="number" label="Packs" value={r.packs} onChange={e=>upd(i,{ packs: Number(e.target.value) })} sx={{ width: 120 }} />
                  <TextField type="number" label="Pieces (optional)" value={r.pieces || ''} onChange={e=>upd(i,{ pieces: Number(e.target.value) })} sx={{ width: 160 }} />
                  <IconButton color="error" onClick={()=>delRow(i)}><DeleteIcon /></IconButton>
                </Stack>
              ))}
              <Button onClick={addRow}>Add Row</Button>
            </Stack>
          </Box>
        </Stack>
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose}>Close</Button>
        <Button variant="contained" onClick={save} disabled={saving || loading}>Save</Button>
      </DialogActions>
    </Dialog>
  );
}
