import { useEffect, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button } from '@mui/material';
import api from '../api';

export default function Settings() {
  const [packsPerPallet, setPPP] = useState<number>(50);
  const [saving, setSaving] = useState(false);

  const load = async () => {
    const { data } = await api.get('/settings');
    setPPP(Number(data.packsPerPallet) || 50);
  };

  const save = async () => {
    setSaving(true);
    try {
      await api.put('/settings', { packsPerPallet });
      await load();
    } finally {
      setSaving(false);
    }
  };

  useEffect(()=>{ load(); },[]);

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Settings</Typography>
      <Paper sx={{ p:2 }}>
        <Stack direction={{ xs: 'column', md: 'row' }} spacing={2} alignItems="center">
          <TextField type="number" label="Default Packs per Pallet" value={packsPerPallet} onChange={e=>setPPP(Number(e.target.value))} />
          <Button variant="contained" onClick={save} disabled={saving}>Save</Button>
        </Stack>
      </Paper>
    </Container>
  );
}
