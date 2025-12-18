import { AppBar, Toolbar, Typography, Button, Stack, Badge, IconButton, Menu, MenuItem, Avatar } from '@mui/material';
import { Link, useLocation, useNavigate } from 'react-router-dom';
import NotificationsIcon from '@mui/icons-material/Notifications';
import { useEffect, useState } from 'react';
import api from '../api';

const LinkButton = ({ to, label }: { to: string; label: string }) => {
  const loc = useLocation();
  const active = loc.pathname === to;
  return (
    <Button component={Link} to={to} color={active ? 'inherit' : 'secondary'} sx={{ color: '#fff', opacity: active ? 1 : 0.85 }}>
      {label}
    </Button>
  );
};

export default function NavBar() {
  const nav = useNavigate();
  const [due, setDue] = useState(0);
  const [onProcessDue, setOnProcessDue] = useState(0);
  useEffect(() => {
    let alive = true;

    const fetchCounts = async () => {
      if (!alive) return;
      try {
        const { data } = await api.get('/shipments/due-today');
        setDue(Number(data?.count || 0));
      } catch {
        setDue(0);
      }
      try {
        const { data } = await api.get('/on-process/batches/due-today');
        setOnProcessDue(Number(data?.count || 0));
      } catch {
        setOnProcessDue(0);
      }
    };

    fetchCounts();
    const interval = setInterval(fetchCounts, 15000);

    const onFocus = () => { fetchCounts(); };
    const onVisibility = () => {
      if (document.visibilityState === 'visible') fetchCounts();
    };

    window.addEventListener('focus', onFocus);
    document.addEventListener('visibilitychange', onVisibility);

    return () => {
      alive = false;
      clearInterval(interval);
      window.removeEventListener('focus', onFocus);
      document.removeEventListener('visibilitychange', onVisibility);
    };
  }, []);
  const logout = () => { localStorage.removeItem('token'); nav('/login'); };
  const userInitial = (() => {
    try {
      const t = localStorage.getItem('token') || '';
      const payload = t.split('.')[1];
      if (!payload) return 'U';
      const json = JSON.parse(atob(payload));
      const name = json?.name || json?.username || json?.email || '';
      const ch = String(name).trim().charAt(0).toUpperCase();
      return ch || 'U';
    } catch { return 'U'; }
  })();
  const isAdmin = (() => {
    try {
      const t = localStorage.getItem('token') || '';
      const payload = t.split('.')[1];
      if (!payload) return false;
      const json = JSON.parse(atob(payload));
      return String(json?.role || '') === 'admin';
    } catch { return false; }
  })();
  const [userMenuEl, setUserMenuEl] = useState<null | HTMLElement>(null);
  const userMenuOpen = Boolean(userMenuEl);
  return (
    <AppBar position="static">
      <Toolbar sx={{ display: 'flex', justifyContent: 'space-between' }}>
        <Typography variant="h6">IGC Inventory</Typography>
        <Stack direction="row" spacing={1} alignItems="center">
          {/* Reordered links */}
          <LinkButton to="/pallets" label="PALLETS SUMMARY" />
          <LinkButton to="/inventory" label="Inventory" />
          <Badge color="error" badgeContent={onProcessDue} max={99} overlap="circular">
            <Button component={Link} to="/on-process" color={useLocation().pathname === '/on-process' ? 'inherit' : 'secondary'} sx={{ color: '#fff', opacity: useLocation().pathname === '/on-process' ? 1 : 0.85 }}>On-Process</Button>
          </Badge>
          <LinkButton to="/transfer" label="Transfer" />
          {/* Ship link with notifications badge */}
          <Badge color="error" badgeContent={due} max={99} overlap="circular">
            <Button component={Link} to="/ship" color={useLocation().pathname === '/ship' ? 'inherit' : 'secondary'} sx={{ color: '#fff', opacity: useLocation().pathname === '/ship' ? 1 : 0.85 }}>Ship</Button>
          </Badge>
          <LinkButton to="/orders" label="Orders" />
          <Typography component="span" sx={{ color: 'rgba(255,255,255,0.7)', mx: 2, userSelect: 'none' }}>|</Typography>
          <LinkButton to="/warehouses" label="Warehouses" />
          <LinkButton to="/item-registry" label="Pallet Registry" />

          {/* User menu (avatar) */}
          <IconButton onClick={(e)=> setUserMenuEl(e.currentTarget)} sx={{ color: '#fff' }}>
            <Avatar sx={{ width: 28, height: 28, bgcolor: '#0ea5e9', fontSize: 14 }}>{userInitial}</Avatar>
          </IconButton>
          <Menu anchorEl={userMenuEl} open={userMenuOpen} onClose={()=> setUserMenuEl(null)} anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }} transformOrigin={{ vertical: 'top', horizontal: 'right' }}>
            <MenuItem onClick={()=> { setUserMenuEl(null); nav('/profile'); }}>Profile</MenuItem>
            {isAdmin && (
              <MenuItem onClick={()=> { setUserMenuEl(null); nav('/users'); }}>Users Management</MenuItem>
            )}
            <MenuItem onClick={()=> { setUserMenuEl(null); logout(); }}>Logout</MenuItem>
          </Menu>
        </Stack>
      </Toolbar>
    </AppBar>
  );
}
