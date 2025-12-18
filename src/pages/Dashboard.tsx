import { useEffect, useState } from 'react';
import { Container, Typography, Grid, Paper, Button } from '@mui/material';
import { useNavigate } from 'react-router-dom';

export default function Dashboard() {
  const [lowCount] = useState(0);
  const nav = useNavigate();

  useEffect(()=>{},[]);

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Dashboard</Typography>
      <Grid container spacing={2}>
        <Grid item xs={12} md={4}>
          <Paper sx={{ p: 2 }}>
            <Typography variant="h6">Low Stock</Typography>
            <Typography variant="h3">{lowCount}</Typography>
            <Button onClick={()=>nav('/inventory')} sx={{ mt: 1 }}>View Inventory</Button>
          </Paper>
        </Grid>
      </Grid>
    </Container>
  );
}
