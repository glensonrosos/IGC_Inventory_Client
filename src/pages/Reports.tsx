import { useEffect, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, Checkbox, FormControlLabel } from '@mui/material';
import api from '../api';

interface ReportItem {
  itemCode: string;
  itemGroup: string;
  description: string;
  color: string;
  totalQty: number;
  packSize: number;
  packsOnHand: number;
  palletsOnHand: number;
  lowStock: boolean;
}

export default function Reports() {
  const [group, setGroup] = useState('');
  const [color, setColor] = useState('');
  const [q, setQ] = useState('');
  const [lowStock, setLowStock] = useState(false);
  const [items, setItems] = useState<ReportItem[]>([]);
  const [count, setCount] = useState(0);

  const load = async () => {
    const { data } = await api.post('/reports/inventory', { group: group || undefined, color: color || undefined, q: q || undefined, lowStock });
    setItems(data.items || []);
    setCount(data.count || 0);
  };

  const exportFile = (format: 'xlsx'|'csv') => {
    const params = new URLSearchParams();
    if (group) params.set('group', group);
    if (color) params.set('color', color);
    if (q) params.set('q', q);
    if (lowStock) params.set('lowStock', 'true');
    params.set('format', format);
    // open in new tab to trigger browser download
    window.open(`/api/reports/export?${params.toString()}`, '_blank');
  };

  useEffect(()=>{ load(); },[]);

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Reports</Typography>
      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs: 'column', md: 'row' }} spacing={2} alignItems="center">
          <TextField label="Group" value={group} onChange={e=>setGroup(e.target.value)} />
          <TextField label="Color" value={color} onChange={e=>setColor(e.target.value)} />
          <TextField label="Search" value={q} onChange={e=>setQ(e.target.value)} />
          <FormControlLabel control={<Checkbox checked={lowStock} onChange={e=>setLowStock(e.target.checked)} />} label="Low Stock only" />
          <Button variant="contained" onClick={load}>Run</Button>
          <Button variant="outlined" onClick={()=>exportFile('xlsx')}>Export Excel</Button>
        </Stack>
      </Paper>

      <Paper sx={{ p:2 }}>
        <Typography variant="subtitle1" gutterBottom>Rows: {count}</Typography>
        <table style={{ width:'100%', borderCollapse:'collapse' }}>
          <thead>
            <tr>
              <th align="left">Item Code</th>
              <th align="left">Group</th>
              <th align="left">Description</th>
              <th align="left">Color</th>
              <th align="right">Total Qty</th>
              <th align="right">Pack Size</th>
              <th align="right">Packs</th>
              <th align="right">Pallets</th>
              <th align="left">Low</th>
            </tr>
          </thead>
          <tbody>
            {items.map((it, i) => (
              <tr key={i} style={{ borderTop:'1px solid #eee', background: it.lowStock ? '#fff4f4' : undefined }}>
                <td>{it.itemCode}</td>
                <td>{it.itemGroup}</td>
                <td>{it.description}</td>
                <td>{it.color}</td>
                <td align="right">{it.totalQty}</td>
                <td align="right">{it.packSize}</td>
                <td align="right">{it.packsOnHand}</td>
                <td align="right">{it.palletsOnHand}</td>
                <td>{it.lowStock ? 'Yes' : ''}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Paper>
    </Container>
  );
}
