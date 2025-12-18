import { useEffect, useState } from 'react';
import { Container, Typography, Paper, TextField, MenuItem, Stack, Button } from '@mui/material';
import api from '../api';

interface MovementItem { itemCode: string; qtyPieces: number; packSize: number; palletId?: string }
interface Txn { _id: string; type: string; reference: string; items: MovementItem[]; createdAt: string; notes?: string }

export default function Transactions() {
  const [items, setItems] = useState<Txn[]>([]);
  const [type, setType] = useState('');
  const [itemCode, setItemCode] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');

  const load = async (page = 1) => {
    const { data } = await api.get('/transactions', { params: { page, limit: 20, type: type || undefined, itemCode: itemCode || undefined, startDate: startDate || undefined, endDate: endDate || undefined } });
    setItems(data.items || []);
  };

  useEffect(()=>{ load(); },[]);

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Transactions</Typography>
      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs: 'column', md: 'row' }} spacing={2} alignItems="center">
          <TextField select label="Type" value={type} onChange={e=>setType(e.target.value)} sx={{ minWidth: 160 }}>
            <MenuItem value="">All</MenuItem>
            <MenuItem value="IN">IN</MenuItem>
            <MenuItem value="OUT">OUT</MenuItem>
            <MenuItem value="ADJUSTMENT">ADJUSTMENT</MenuItem>
            <MenuItem value="ALLOCATE">ALLOCATE</MenuItem>
            <MenuItem value="RECEIPT">RECEIPT</MenuItem>
          </TextField>
          <TextField label="Item Code" value={itemCode} onChange={e=>setItemCode(e.target.value)} />
          <TextField type="date" label="Start" InputLabelProps={{ shrink: true }} value={startDate} onChange={e=>setStartDate(e.target.value)} />
          <TextField type="date" label="End" InputLabelProps={{ shrink: true }} value={endDate} onChange={e=>setEndDate(e.target.value)} />
          <Button variant="contained" onClick={()=>load()}>Apply</Button>
        </Stack>
      </Paper>
      <Paper sx={{ p:2 }}>
        <table style={{ width:'100%', borderCollapse:'collapse' }}>
          <thead>
            <tr>
              <th align="left">Date</th>
              <th align="left">Type</th>
              <th align="left">Reference</th>
              <th align="left">Items</th>
              <th align="left">Notes</th>
            </tr>
          </thead>
          <tbody>
            {items.map(t => (
              <tr key={t._id} style={{ borderTop:'1px solid #eee' }}>
                <td>{new Date(t.createdAt).toLocaleString()}</td>
                <td>{t.type}</td>
                <td>{t.reference}</td>
                <td>{t.items?.map(i => `${i.itemCode}(${i.qtyPieces})`).join(', ')}</td>
                <td>{t.notes}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Paper>
    </Container>
  );
}
