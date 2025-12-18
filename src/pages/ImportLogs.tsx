import { useEffect, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, MenuItem } from '@mui/material';
import api from '../api';

interface LogItem {
  _id: string;
  createdAt: string;
  type: 'stock_in'|'initial'|'orders';
  fileName?: string;
  poNumber?: string;
  itemCode?: string;
  totalQty?: number;
  packSize?: number;
}

export default function ImportLogs() {
  const [items, setItems] = useState<LogItem[]>([]);
  const [page, setPage] = useState(1);
  const [pages, setPages] = useState(1);

  const [type, setType] = useState('');
  const [poNumber, setPo] = useState('');
  const [itemCode, setItem] = useState('');
  const [startDate, setStart] = useState('');
  const [endDate, setEnd] = useState('');

  const load = async (p = 1) => {
    const { data } = await api.get('/import-logs', { params: {
      page: p, limit: 20,
      type: type || undefined,
      poNumber: poNumber || undefined,
      itemCode: itemCode || undefined,
      startDate: startDate || undefined,
      endDate: endDate || undefined
    }});
    setItems(data.items || []);
    setPage(data.page || 1);
    setPages(data.pages || 1);
  };

  useEffect(()=>{ load(1); },[]);

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Import Logs</Typography>
      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs: 'column', md: 'row' }} spacing={2} alignItems="center">
          <TextField select label="Type" value={type} onChange={e=>setType(e.target.value)} sx={{ minWidth: 160 }}>
            <MenuItem value="">All</MenuItem>
            <MenuItem value="stock_in">stock_in</MenuItem>
            <MenuItem value="initial">initial</MenuItem>
            <MenuItem value="orders">orders</MenuItem>
          </TextField>
          <TextField label="PO #" value={poNumber} onChange={e=>setPo(e.target.value)} />
          <TextField label="Item Code" value={itemCode} onChange={e=>setItem(e.target.value)} />
          <TextField type="date" label="Start" InputLabelProps={{ shrink: true }} value={startDate} onChange={e=>setStart(e.target.value)} />
          <TextField type="date" label="End" InputLabelProps={{ shrink: true }} value={endDate} onChange={e=>setEnd(e.target.value)} />
          <Button variant="contained" onClick={()=>load(1)}>Apply</Button>
        </Stack>
      </Paper>

      <Paper sx={{ p:2 }}>
        <table style={{ width:'100%', borderCollapse:'collapse' }}>
          <thead>
            <tr>
              <th align="left">Date</th>
              <th align="left">Type</th>
              <th align="left">File</th>
              <th align="left">PO #</th>
              <th align="left">Item Code</th>
              <th align="right">Qty</th>
              <th align="right">Pack Size</th>
            </tr>
          </thead>
          <tbody>
            {items.map(it => (
              <tr key={it._id} style={{ borderTop:'1px solid #eee' }}>
                <td>{new Date(it.createdAt).toLocaleString()}</td>
                <td>{it.type}</td>
                <td>{it.fileName || ''}</td>
                <td>{it.poNumber || ''}</td>
                <td>{it.itemCode || ''}</td>
                <td align="right">{it.totalQty ?? ''}</td>
                <td align="right">{it.packSize ?? ''}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <Stack direction="row" spacing={1} sx={{ mt:2 }}>
          <Button disabled={page<=1} onClick={()=>load(page-1)}>Prev</Button>
          <Typography variant="body2" sx={{ alignSelf:'center' }}>Page {page} / {pages}</Typography>
          <Button disabled={page>=pages} onClick={()=>load(page+1)}>Next</Button>
        </Stack>
      </Paper>
    </Container>
  );
}
