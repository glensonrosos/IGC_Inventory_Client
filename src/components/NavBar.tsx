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
  const loc = useLocation();
  const [due, setDue] = useState(0);
  const [onProcessDue, setOnProcessDue] = useState(0);
  const [ordersDueToday, setOrdersDueToday] = useState(0);
  const [ordersDeliveredDue, setOrdersDeliveredDue] = useState(0);
  useEffect(() => {
    let alive = true;

    const pad2 = (n: number) => String(n).padStart(2, '0');
    const toLocalYmd = (v?: any) => {
      if (!v) return '';
      // Accept Date, ISO string, or YYYY-MM-DD. Always normalize to local date.
      const d = v instanceof Date ? v : new Date(String(v));
      if (Number.isNaN(d.getTime())) {
        const s = String(v);
        const slice = s.slice(0, 10);
        return /^\d{4}-\d{2}-\d{2}$/.test(slice) ? slice : '';
      }
      return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
    };
    const todayYmd = () => toLocalYmd(new Date());

    const normalizeStatus = (v: any) => {
      const s = String(v || '').trim().toLowerCase();
      if (!s) return '';
      if (s === 'created') return 'processing';
      if (s === 'create') return 'processing';
      if (s === 'backorder') return 'processing';
      if (s === 'fulfilled') return 'completed';
      if (s === 'cancelled') return 'canceled';
      if (s === 'cancel') return 'canceled';
      return s;
    };

    const toYmd = (v: any) => toLocalYmd(v);

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
      try {
        const { data } = await api.get<any[]>('/orders/unfulfilled');
        const list = Array.isArray(data) ? data : [];
        const today = todayYmd();
        const readyToShip = list.filter((o: any) => {
          const st = normalizeStatus(o?.status);
          return st === 'ready_to_ship';
        }).length;
        const deliveredDue = list.filter((o: any) => {
          const st = normalizeStatus(o?.status);
          if (st !== 'shipped') return false;
          const ymd = toYmd(o?.estDeliveredDate || o?.estDelivered);
          return ymd && ymd <= today;
        }).length;
        setOrdersDueToday(readyToShip);
        setOrdersDeliveredDue(deliveredDue);
      } catch {
        setOrdersDueToday(0);
        setOrdersDeliveredDue(0);
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
        <Typography
          variant="h6"
          component={Link}
          to="/pallets"
          onClick={(e) => {
            if (loc.pathname === '/pallets') {
              e.preventDefault();
              window.location.reload();
            }
          }}
          sx={{ color: '#fff', textDecoration: 'none', cursor: 'pointer' }}
        >
          IGC Inventory
        </Typography>
        <Stack direction="row" spacing={1} alignItems="center">
          {/* Orders and Early Buy first */}
          <Badge color="error" badgeContent={(ordersDueToday || 0) + (ordersDeliveredDue || 0)} max={99} overlap="circular">
            <Button component={Link} to="/orders" color={loc.pathname === '/orders' ? 'inherit' : 'secondary'} sx={{ color: '#fff', opacity: loc.pathname === '/orders' ? 1 : 0.85 }}>Orders</Button>
          </Badge>
          <LinkButton to="/early-buy" label="Early Buy" />
          <Typography component="span" sx={{ color: 'rgba(255,255,255,0.7)', mx: 2, userSelect: 'none' }}>|</Typography>
          {/* Remaining links */}
          <LinkButton to="/pallets" label="PALLETS SUMMARY" />
          <LinkButton to="/inventory" label="Inventory" />
          <Badge color="error" badgeContent={onProcessDue} max={99} overlap="circular">
            <Button component={Link} to="/on-process" color={loc.pathname === '/on-process' ? 'inherit' : 'secondary'} sx={{ color: '#fff', opacity: loc.pathname === '/on-process' ? 1 : 0.85 }}>On-Process</Button>
          </Badge>
          <LinkButton to="/transfer" label="Transfer" />
          {/* Ship link with notifications badge */}
          <Badge color="error" badgeContent={due} max={99} overlap="circular">
            <Button component={Link} to="/ship" color={loc.pathname === '/ship' ? 'inherit' : 'secondary'} sx={{ color: '#fff', opacity: loc.pathname === '/ship' ? 1 : 0.85 }}>Ship</Button>
          </Badge>
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
