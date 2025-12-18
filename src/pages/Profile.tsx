import { useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button } from '@mui/material';
import api from '../api';
import { useToast } from '../components/ToastProvider';

export default function Profile() {
  const [oldPassword, setOld] = useState('');
  const [newPassword, setNew] = useState('');
  const [confirm, setConfirm] = useState('');
  const [saving, setSaving] = useState(false);
  const toast = useToast();

  const submit = async () => {
    if (!oldPassword || !newPassword) { toast.error('Fill all fields'); return; }
    if (newPassword !== confirm) { toast.error('Passwords do not match'); return; }
    setSaving(true);
    try {
      await api.post('/auth/change-password', { oldPassword, newPassword });
      toast.success('Password updated');
      setOld(''); setNew(''); setConfirm('');
    } catch (e: any) {
      toast.error(e?.response?.data?.message || 'Failed to update password');
    } finally {
      setSaving(false);
    }
  };

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Profile</Typography>
      <Paper sx={{ p:2 }}>
        <Typography variant="h6" gutterBottom>Change Password</Typography>
        <Stack direction={{ xs:'column', md:'row' }} spacing={2} alignItems="center">
          <TextField label="Old Password" type="password" value={oldPassword} onChange={(e)=>setOld(e.target.value)} />
          <TextField label="New Password" type="password" value={newPassword} onChange={(e)=>setNew(e.target.value)} />
          <TextField label="Confirm New Password" type="password" value={confirm} onChange={(e)=>setConfirm(e.target.value)} />
          <Button variant="contained" onClick={submit} disabled={saving}>Update</Button>
        </Stack>
      </Paper>
    </Container>
  );
}
