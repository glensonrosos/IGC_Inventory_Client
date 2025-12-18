import { useState } from 'react';
import { Container, TextField, Button, Typography, Box, Paper } from '@mui/material';
import api from '../api';
import { useNavigate } from 'react-router-dom';

export default function Login() {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const nav = useNavigate();

  const onSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    try {
      const { data } = await api.post('/auth/login', { username, password });
      localStorage.setItem('token', data.token);
      nav('/');
    } catch (err: any) {
      setError(err?.response?.data?.message || 'Login failed');
    }
  };

  return (
    <Box sx={{
      minHeight: '100vh',
      backgroundImage: 'url(https://images.unsplash.com/photo-1459664018906-085c36f472af?q=80&w=1600&auto=format&fit=crop)',
      backgroundSize: 'cover',
      backgroundPosition: 'center',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      p: 2
    }}>
      <Container maxWidth="sm">
        <Box sx={{ bgcolor: 'rgba(255,255,255,0.85)', borderRadius: 2, p: { xs: 2, sm: 3 } }}>
          <Paper elevation={6} sx={{ p: { xs: 3, sm: 4 }, borderRadius: 2 }}>
            <Typography variant="h5" gutterBottom align="center">Sign in to IGC Inventory</Typography>
            <Box component="form" onSubmit={onSubmit}>
              <TextField fullWidth margin="normal" label="Username" value={username} onChange={e=>setUsername(e.target.value)} />
              <TextField fullWidth margin="normal" type="password" label="Password" value={password} onChange={e=>setPassword(e.target.value)} />
              {error && <Typography color="error" variant="body2" sx={{ mt: 1 }}>{error}</Typography>}
              <Button fullWidth type="submit" variant="contained" sx={{ mt: 2 }}>Sign In</Button>
              <Typography variant="caption" display="block" align="center" sx={{ mt: 2, color: 'text.secondary' }}>
                © Glenson_Encode
              </Typography>
            </Box>
          </Paper>
        </Box>
      </Container>
    </Box>
  );
}
