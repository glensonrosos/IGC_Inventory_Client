import { useEffect, useMemo, useState } from 'react';
import { Dialog, DialogTitle, DialogContent, Stack, TextField, FormGroup, FormControlLabel, Checkbox, Button } from '@mui/material';
import api from '../api';

interface Item { _id: string; itemCode: string; itemGroup: string; description: string; color: string; totalQty: number; enabled?: boolean }

export default function ItemSearchModal({ open, onClose, onSelect }: { open: boolean; onClose: () => void; onSelect: (item: Item) => void }) {
  const [items, setItems] = useState<Item[]>([]);
  const [loading, setLoading] = useState(false);
  const [search, setSearch] = useState('');
  const [fields, setFields] = useState<string[]>(['itemCode','description','color']);
  const [includeDisabled, setIncludeDisabled] = useState(false);

  useEffect(() => {
    if (!open) return;
    const load = async () => {
      setLoading(true);
      try {
        const { data } = await api.get<Item[]>('/items', { params: includeDisabled ? { includeDisabled: 1 } : {} });
        setItems(data);
      } finally {
        setLoading(false);
      }
    };
    load();
  }, [open, includeDisabled]);

  const activeFields = fields.length ? fields : ['itemCode','description','color'];
  const list = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return items;
    return items.filter((it) => activeFields.some((f) => String((it as any)[f] ?? '').toLowerCase().includes(q)));
  }, [items, search, fields]);

  return (
    <Dialog open={open} onClose={onClose} maxWidth="md" fullWidth>
      <DialogTitle>Search Item</DialogTitle>
      <DialogContent>
        <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} alignItems="center" sx={{ my: 1 }}>
          <TextField label="Search" placeholder="Item Code / Description / Color" value={search} onChange={(e)=>setSearch(e.target.value)} size="small" sx={{ minWidth: 260 }} />
          <FormGroup row>
            {['itemCode','description','color'].map((f) => (
              <FormControlLabel
                key={f}
                control={<Checkbox size="small" checked={fields.includes(f)} onChange={(e) => {
                  setFields((prev) => e.target.checked ? [...prev, f] : prev.filter(x => x !== f));
                }} />}
                label={f === 'itemCode' ? 'Item Code' : f.charAt(0).toUpperCase() + f.slice(1)}
              />
            ))}
            <FormControlLabel control={<Checkbox size="small" checked={includeDisabled} onChange={(e)=>setIncludeDisabled(e.target.checked)} />} label="Include disabled" />
          </FormGroup>
        </Stack>
        <table style={{ width:'100%', borderCollapse:'collapse' }}>
          <thead>
            <tr>
              <th align="left">Item Code</th>
              <th align="left">Description</th>
              <th align="left">Color</th>
              <th align="left">Enabled</th>
              <th align="right">Total Qty</th>
              <th align="left">Action</th>
            </tr>
          </thead>
          <tbody>
            {list.map((it) => (
              <tr key={it._id} style={{ borderTop:'1px solid #eee' }}>
                <td>{it.itemCode}</td>
                <td>{it.description}</td>
                <td>{it.color}</td>
                <td>{(it.enabled ?? true) ? 'Yes' : 'No'}</td>
                <td align="right">{it.totalQty}</td>
                <td><Button size="small" variant="contained" onClick={() => { onSelect(it); onClose(); }}>Select</Button></td>
              </tr>
            ))}
            {!list.length && (
              <tr><td colSpan={6} style={{ padding:'12px 0', color:'#666' }}>{loading ? 'Loading...' : 'No matching items'}</td></tr>
            )}
          </tbody>
        </table>
      </DialogContent>
    </Dialog>
  );
}
