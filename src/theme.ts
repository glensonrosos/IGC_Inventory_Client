import { createTheme } from '@mui/material/styles';

const theme = createTheme({
  palette: {
    mode: 'light',
    primary: { main: '#0d9488' }, // teal
    secondary: { main: '#2563eb' }, // blue
    background: { default: '#f8fafc', paper: '#ffffff' }
  },
  shape: { borderRadius: 10 },
  typography: {
    fontFamily: ['Inter', 'Segoe UI', 'Roboto', 'Helvetica', 'Arial', 'sans-serif'].join(','),
    h4: { fontWeight: 700 },
    h6: { fontWeight: 600 }
  },
  components: {
    MuiAppBar: { styleOverrides: { root: { boxShadow: 'none', borderBottom: '1px solid #e5e7eb' } } },
    MuiPaper: { styleOverrides: { root: { borderRadius: 12 } } },
  }
});

export default theme;
