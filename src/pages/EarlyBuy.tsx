import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, Dialog, DialogTitle, DialogContent, DialogActions, MenuItem, Box, IconButton, Chip } from '@mui/material';
import { DataGrid, GridColDef, GridToolbar } from '@mui/x-data-grid';
import DeleteIcon from '@mui/icons-material/Delete';
import DownloadIcon from '@mui/icons-material/Download';
import OpenInNewIcon from '@mui/icons-material/OpenInNew';
import * as XLSX from 'xlsx';
import api from '../api';
import { useToast } from '../components/ToastProvider';

// This page is a lightweight clone of Orders for "Early Buy" workflows.
// It does NOT affect inventory or reservations; all data is saved locally (localStorage).

type EarlyOrder = {
  id: string; // EORD-0001 pattern
  status: 'processing' | 'ready_to_ship' | 'shipped' | 'completed' | 'canceled';
  warehouseId: string; // fixed MPG (display-only)
  createdAt: string; // YYYY-MM-DD
  estFulfillment: string; // YYYY-MM-DD
  estDelivered: string; // YYYY-MM-DD
  customerEmail: string;
  customerName: string;
  customerPhone: string;
  shippingAddress: string;
  originalPrice?: string;
  shippingPercent?: string;
  discountPercent?: string;
  notes?: string;
  lines: Array<{ groupName: string; lineItem: string; palletName: string; qty: number }>;
};

const STATUS_OPTIONS: Array<{ label: string; value: EarlyOrder['status'] }> = [
  { label: 'PROCESSING', value: 'processing' },
  { label: 'READY TO SHIP', value: 'ready_to_ship' },
  { label: 'SHIPPED', value: 'shipped' },
  { label: 'COMPLETED', value: 'completed' },
  { label: 'CANCELED', value: 'canceled' },
];

// Backend is the source of truth; IDs are generated server-side.

export default function EarlyBuy() {
  const toast = useToast();
  const [orders, setOrders] = useState<EarlyOrder[]>([]);
  const [open, setOpen] = useState(false);
  const [editingOrder, setEditingOrder] = useState<EarlyOrder | null>(null);

  // Form state
  const [status, setStatus] = useState<EarlyOrder['status']>('processing');
  const [customerEmail, setCustomerEmail] = useState('');
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [createdAt, setCreatedAt] = useState(() => new Date().toISOString().slice(0,10));
  const [estFulfillment, setEstFulfillment] = useState('');
  const [estDelivered, setEstDelivered] = useState('');
  const [shippingAddress, setShippingAddress] = useState('');

  const isEditable = useMemo(() => {
    // Editable when creating (no editingOrder) OR when status is processing/ready_to_ship
    return !editingOrder || status === 'processing' || status === 'ready_to_ship';
  }, [editingOrder, status]);
  const [originalPrice, setOriginalPrice] = useState('');
  const [shippingPercent, setShippingPercent] = useState('');
  const [discountPercent, setDiscountPercent] = useState('');
  const [notes, setNotes] = useState('');

  const computedFinalPrice = useMemo(() => {
    const op = Number(originalPrice);
    const sp = Number(shippingPercent);
    const dp = Number(discountPercent);
    if (!Number.isFinite(op)) return '';
    const disc = Number.isFinite(dp) ? Math.min(100, Math.max(0, dp)) : 0;
    const ship = Number.isFinite(sp) ? Math.min(100, Math.max(0, sp)) : 0;
    const out = op * (1 - disc / 100) * (1 + ship / 100);
    return Number.isFinite(out) ? out.toFixed(2) : '';
  }, [originalPrice, shippingPercent, discountPercent]);

  const isValidEmail = (email: string) => {
    const s = String(email || '').trim();
    if (!s) return false;
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
  };

  // Picker
  const [pickerOpen, setPickerOpen] = useState(false);
  const [pickerLoading, setPickerLoading] = useState(false);
  const [pickerRows, setPickerRows] = useState<any[]>([]);
  const [pickerQ, setPickerQ] = useState('');
  const [pickerQDebounced, setPickerQDebounced] = useState('');
  useEffect(() => {
    const h = setTimeout(() => setPickerQDebounced(String(pickerQ || '')), 250);
    return () => clearTimeout(h);
  }, [pickerQ]);
  const [pickerWarehouses, setPickerWarehouses] = useState<any[]>([]);
  const [selectedGroups, setSelectedGroups] = useState<string[]>([]);

  // Lines added
  const [lines, setLines] = useState<Array<{ groupName: string; lineItem: string; palletName: string; qty: number }>>([]);

  // Orders table filtering
  const [tableQ, setTableQ] = useState('');
  const [tableStatus, setTableStatus] = useState<'all' | EarlyOrder['status']>('all');

  // MUI DataGrid compatibility shim (v5 vs v6 selection APIs)
  const DG: any = DataGrid as any;

  useEffect(() => {
    if (!pickerOpen) return;
    // reset selection each time dialog opens
    setSelectedGroups([]);
    let canceled = false;
    (async () => {
      try {
        setPickerLoading(true);
        // Ensure we have a warehouse to query against
        let whs = pickerWarehouses;
        if (!Array.isArray(whs) || whs.length === 0) {
          try {
            const { data: ws } = await api.get('/warehouses');
            whs = Array.isArray(ws) ? ws : [];
            if (!canceled) setPickerWarehouses(whs);
          } catch {
            whs = [];
          }
        }
        const wid = String((whs[0]?._id) || '').trim();
        if (!wid) { if (!canceled) setPickerRows([]); return; }
        // Fetch full dataset once; filter client-side for Pallet Name/ID/Description and EDDs
        const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: wid } });
        if (!canceled) setPickerRows(Array.isArray(data?.rows) ? data.rows : []);
      } catch {
        if (!canceled) setPickerRows([]);
      } finally {
        if (!canceled) setPickerLoading(false);
      }
    })();
    return () => { canceled = true; };
  }, [pickerOpen]);

  const pickerRowsFiltered = useMemo(() => {
    const q = String(pickerQDebounced || '').trim().toLowerCase();
    const rows = (Array.isArray(pickerRows) ? pickerRows : [])
      // Filter out rows lacking identifiers to avoid unstable selection/IDs
      .filter((r: any) => String(r?.lineItem || r?.groupName || r?.palletName || '').trim().length > 0);
    if (!q) return rows;
    return rows.filter((r: any) => {
      const gid = String(r?.lineItem || '').trim().toLowerCase();
      const gnameLower = String(r?.groupName || '').trim().toLowerCase();
      const pname = String(r?.palletName || '').trim().toLowerCase();
      return gid.includes(q) || gnameLower.includes(q) || pname.includes(q);
    });
  }, [pickerRows, pickerQDebounced]);

  const addSelectedToLines = () => {
    const rows = Array.isArray(pickerRows) ? pickerRows : [];
    const set = new Set(selectedGroups.map((g) => String(g)));
    const added: typeof lines = [];
    for (const r of rows) {
      const gName = String(r?.groupName || '');
      const lItem = String(r?.lineItem || '');
      const composedId = `${lItem.trim()}::${gName.trim()}`;
      // Accept selection by composed ID, Pallet ID, or Group Name (for compatibility)
      if (!set.has(composedId) && !set.has(lItem) && !set.has(gName)) continue;
      added.push({ groupName: gName, lineItem: lItem, palletName: String(r?.palletName || ''), qty: 0 });
    }
    const existing = new Map(lines.map(l => [l.groupName.toLowerCase(), l]));
    const merged: typeof lines = [...lines];
    for (const a of added) {
      if (!existing.has(a.groupName.toLowerCase())) merged.push(a);
    }
    setLines(merged);
    setPickerOpen(false);
    setSelectedGroups([]);
  };

  // Helpers to match Orders page presentation
  const normalizeStatus = (v: any) => {
    const s = String(v || '').trim().toLowerCase();
    if (s === 'ready_to_ship') return 'ready_to_ship';
    if (s === 'shipped') return 'shipped';
    if (s === 'completed') return 'completed';
    if (s === 'canceled' || s === 'cancelled' || s === 'cancel') return 'canceled';
    return 'processing';
  };
  const fmtDate = (v: any) => {
    const s = String(v || '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return '-';
    const d = new Date(`${s}T00:00:00`);
    return d && !Number.isNaN(d.getTime()) ? d.toLocaleDateString() : '-';
  };
  const todayYmd = new Date().toISOString().slice(0,10);

  const ordersColumns: GridColDef[] = useMemo(() => [
    { field: 'id', headerName: 'Order #', width: 140 },
    {
      field: 'status',
      headerName: 'Status',
      width: 140,
      renderCell: (p: any) => {
        const row = (p?.row as any) || {};
        const s = normalizeStatus(row?.status);
        const label = s === 'ready_to_ship' ? 'READY TO SHIP' : (s ? String(s).toUpperCase() : '-');
        const color: any = s === 'completed'
          ? 'success'
          : s === 'canceled'
            ? 'error'
            : s === 'processing'
              ? 'warning'
              : s === 'ready_to_ship'
                ? 'info'
                : 'default';
        const variant: any = (s === 'processing') ? 'outlined' : 'filled';
        return <Chip size="small" label={label} color={color} variant={variant} />;
      },
    },
    { field: 'customerEmail', headerName: 'Customer Email', width: 220 },
    { field: 'customerName', headerName: 'Customer Name', width: 180 },
    {
      field: 'createdAt',
      headerName: 'Order Created',
      width: 120,
      renderCell: (p: any) => fmtDate((p?.row as any)?.createdAt),
    },
    {
      field: 'estFulfillment',
      headerName: 'Estimated Shipdate for Customer',
      width: 160,
      renderCell: (p: any) => fmtDate((p?.row as any)?.estFulfillment),
    },
    {
      field: 'estDelivered',
      headerName: 'Estimated Arrival Date',
      width: 160,
      renderCell: (p: any) => {
        const row = (p?.row as any) || {};
        const st = normalizeStatus(row?.status);
        const ymd = String(row?.estDelivered || '').slice(0, 10);
        const isDue = st === 'shipped' && ymd && ymd <= todayYmd;
        const label = fmtDate(row?.estDelivered);
        return (
          <Box component="span" sx={isDue ? { px: 1, py: 0.25, borderRadius: 1, bgcolor: 'rgba(245, 124, 0, 0.18)', fontWeight: 600 } : undefined}>
            {label}
          </Box>
        );
      },
    },
    {
      field: 'shippingPercent',
      headerName: 'Shipping Charges (%)',
      width: 120,
      renderCell: (p: any) => {
        const v = (p?.row as any)?.shippingPercent;
        if (v === null || v === undefined || v === '') return '-';
        const n = Number(v);
        return Number.isFinite(n) ? `${n}%` : '-';
      },
    },
    {
      field: 'discountPercent',
      headerName: 'Discount (%)',
      width: 110,
      renderCell: (p: any) => {
        const v = (p?.row as any)?.discountPercent;
        if (v === null || v === undefined || v === '') return '-';
        const n = Number(v);
        return Number.isFinite(n) ? `${n}%` : '-';
      },
    },
    {
      field: 'finalPrice',
      headerName: 'Final Price',
      width: 110,
      renderCell: (p: any) => {
        const row = (p?.row as any) || {};
        const op = Number(row?.originalPrice);
        const sp = Number(row?.shippingPercent);
        const dp = Number(row?.discountPercent);
        if (Number.isFinite(op) && Number.isFinite(sp) && Number.isFinite(dp)) {
          const disc = Math.min(100, Math.max(0, dp));
          const ship = Math.min(100, Math.max(0, sp));
          const out = op * (1 - disc / 100) * (1 + ship / 100);
          return Number.isFinite(out) ? out.toFixed(2) : '-';
        }
        return '-';
      },
    },
    {
      field: 'linesCount',
      headerName: 'Lines',
      width: 80,
      align: 'right',
      headerAlign: 'right',
      renderCell: (p: any) => {
        const r: any = p?.row || {};
        const pre = (r as any).linesCount;
        if (typeof pre === 'number') return pre;
        const arr = Array.isArray(r.lines) ? r.lines : [];
        return arr.length;
      },
    },
    {
      field: 'qtyTotal',
      headerName: 'Qty',
      width: 80,
      align: 'right',
      headerAlign: 'right',
      renderCell: (p: any) => {
        const r: any = p?.row || {};
        const pre = (r as any).qtyTotal;
        if (typeof pre === 'number') return pre;
        const arr = Array.isArray(r.lines) ? r.lines : [];
        return arr.reduce((sum: number, l: any) => sum + Math.max(0, Math.floor(Number(l?.qty || 0))), 0);
      },
    },
    {
      field: 'actions',
      headerName: 'Actions',
      width: 120,
      sortable: false,
      filterable: false,
      renderCell: (p: any) => {
        const row = (p?.row as any) || {};
        return (
          <Stack direction="row" spacing={0.5} alignItems="center">
            <IconButton size="small" color="primary" onClick={() => {
              const r = orders.find(o => o.id === row.id);
              if (!r) return;
              setEditingOrder(r);
              setStatus(r.status);
              setCustomerEmail(r.customerEmail);
              setCustomerName(r.customerName);
              setCustomerPhone(r.customerPhone);
              setShippingAddress(r.shippingAddress);
              setCreatedAt(r.createdAt);
              setEstFulfillment(r.estFulfillment);
              setEstDelivered(r.estDelivered);
              setOriginalPrice(String(r.originalPrice||''));
              setShippingPercent(String(r.shippingPercent||''));
              setDiscountPercent(String(r.discountPercent||''));
              setNotes(String(r.notes||''));
              setLines(Array.isArray(r.lines) ? r.lines.map((l: any) => ({ ...l })) : []);
              setOpen(true);
            }}>
              <OpenInNewIcon fontSize="inherit" />
            </IconButton>
            <IconButton size="small" onClick={() => exportOrderXlsx(row)}>
              <DownloadIcon fontSize="inherit" />
            </IconButton>
          </Stack>
        );
      },
    },
  ], [todayYmd]);

  const filteredOrders = useMemo(() => {
    const q = String(tableQ || '').trim().toLowerCase();
    const st = tableStatus;
    return (orders || []).filter((o) => {
      const statusOk = st === 'all' ? true : normalizeStatus(o.status) === st;
      if (!q) return statusOk;
      const hay = `${o.id} ${o.customerEmail} ${o.customerName}`.toLowerCase();
      return statusOk && hay.includes(q);
    });
  }, [orders, tableQ, tableStatus]);

  const linesColumns: GridColDef[] = useMemo(() => [
    { field: 'palletName', headerName: 'Pallet Name', width: 220 },
    { field: 'groupName', headerName: 'Pallet Description', width: 260 },
    { field: 'lineItem', headerName: 'Pallet ID', width: 160 },
    {
      field: 'qty',
      headerName: 'Order Qty',
      width: 120,
      align: 'right',
      headerAlign: 'right',
      renderCell: (p: any) => {
        const idx = Math.max(0, Number(p?.id) - 1);
        const val = lines[idx]?.qty ?? 0;
        return (
          <TextField
            size="small"
            type="number"
            value={val}
            onChange={(e) => {
              const v = Math.max(0, Math.floor(Number(e.target.value) || 0));
              const next = [...lines];
              if (next[idx]) next[idx] = { ...next[idx], qty: v };
              setLines(next);
            }}
            inputProps={{ inputMode: 'numeric', pattern: '[0-9]*', min: 0 }}
            sx={{ '& input': { textAlign: 'right' } }}
          />
        );
      },
    },
    {
      field: 'actions',
      headerName: '',
      width: 60,
      sortable: false,
      filterable: false,
      align: 'center',
      renderCell: (p: any) => {
        const idx = Math.max(0, Number(p?.id) - 1);
        return (
          <IconButton size="small" color="error" onClick={() => {
            const next = [...lines];
            next.splice(idx, 1);
            setLines(next);
          }}>
            <DeleteIcon fontSize="small" />
          </IconButton>
        );
      },
    },
  ], [lines]);

  const refreshOrders = useCallback(async () => {
    try {
      const { data } = await api.get('/early-buy');
      const normalizeStatus = (v: any): EarlyOrder['status'] => {
        const s = String(v || '').trim().toLowerCase();
        if (s === 'ready_to_ship') return 'ready_to_ship';
        if (s === 'shipped') return 'shipped';
        if (s === 'completed') return 'completed';
        if (s === 'canceled' || s === 'cancelled' || s === 'cancel') return 'canceled';
        return 'processing';
      };
      const mapped: EarlyOrder[] = Array.isArray(data) ? data.map((d: any): EarlyOrder => ({
        id: String(d?.id || ''),
        status: normalizeStatus(d?.status),
        warehouseId: String(d?.warehouseId || ''),
        createdAt: String(d?.createdAtYmd || d?.createdAt || '').slice(0,10),
        estFulfillment: String(d?.estFulfillment || ''),
        estDelivered: String(d?.estDelivered || ''),
        customerEmail: String(d?.customerEmail || ''),
        customerName: String(d?.customerName || ''),
        customerPhone: String(d?.customerPhone || ''),
        shippingAddress: String(d?.shippingAddress || ''),
        originalPrice: String(d?.originalPrice || ''),
        shippingPercent: String(d?.shippingPercent || ''),
        discountPercent: String(d?.discountPercent || ''),
        notes: String(d?.notes || ''),
        lines: (() => {
          const raw = Array.isArray(d?.lines)
            ? d.lines
            : Array.isArray(d?.orderLines)
              ? d.orderLines
              : Array.isArray(d?.items)
                ? d.items
                : Array.isArray(d?.order?.lines)
                  ? d.order.lines
                  : Array.isArray(d?.payload?.lines)
                    ? d.payload.lines
                    : [];
          return raw.map((l: any) => ({
            groupName: String(l?.groupName || l?.description || ''),
            lineItem: String(l?.lineItem || l?.palletId || l?.sku || ''),
            palletName: String(l?.palletName || l?.name || ''),
            qty: Math.max(0, Math.floor(Number(
              l?.qty ?? l?.quantity ?? l?.orderedQty ?? l?.orderQty ?? l?.qtyOrdered ?? l?.qty_ordered ?? l?.quantityOrdered ?? 0
            )))
          }));
        })(),
        // Precompute counts for robustness in grid rendering
        ...( (() => {
          const arr = Array.isArray(d?.lines)
            ? d.lines
            : Array.isArray(d?.orderLines)
              ? d.orderLines
              : Array.isArray(d?.items)
                ? d.items
                : Array.isArray(d?.order?.lines)
                  ? d.order.lines
                  : Array.isArray(d?.payload?.lines)
                    ? d.payload.lines
                    : [];
          const qty = arr.reduce((s: number, l: any) => s + Math.max(0, Math.floor(Number(
            l?.qty ?? l?.quantity ?? l?.orderedQty ?? l?.orderQty ?? l?.qtyOrdered ?? l?.qty_ordered ?? l?.quantityOrdered ?? 0
          ))), 0);
          return { linesCount: arr.length, qtyTotal: qty } as any;
        })() ),
      })) : [];
      setOrders(mapped);
      try {
        const sample = mapped.slice(0, 3).map(o => ({
          id: o.id,
          lines: Array.isArray(o.lines) ? o.lines.length : 0,
          qty: Array.isArray(o.lines) ? o.lines.reduce((s, l) => s + Number(l.qty||0), 0) : 0,
        }));
        console.debug('[EarlyBuy] mapped orders sample', sample);
      } catch {}
    } catch {
      setOrders([]);
    }
  }, []);

  useEffect(() => { refreshOrders(); }, [refreshOrders]);

  const saveNewOrder = async () => {
    // Validate
    const errs: string[] = [];
    const today = new Date().toISOString().slice(0,10);
    const isYmd = (s: string) => /^\d{4}-\d{2}-\d{2}$/.test(String(s||''));
    if (!customerEmail.trim()) errs.push('Customer Email is required');
    if (!customerName.trim()) errs.push('Customer Name is required');
    if (customerEmail && !isValidEmail(customerEmail)) errs.push('Customer Email is invalid');
    if (!customerPhone.trim()) errs.push('Phone Number is required');
    if (!shippingAddress.trim()) errs.push('Shipping Address is required');
    if (!isYmd(createdAt)) errs.push('Created Order Date is invalid');
    if (createdAt > today) errs.push('Created Order Date cannot be in the future');
    // Estimated ship date required
    if (!estFulfillment) errs.push('Estimated ShipDate for Customer is required');
    if (estFulfillment && !isYmd(estFulfillment)) errs.push('Estimated ShipDate for Customer is invalid');
    if (estFulfillment && estFulfillment < createdAt) errs.push('Estimated ShipDate must be >= Created Order Date');
    // If status is SHIPPED, estimated arrival date required
    if (status === 'shipped' && !estDelivered) errs.push('Estimated Arrival Date is required when status is SHIPPED');
    if (estDelivered && !isYmd(estDelivered)) errs.push('Estimated Arrival Date is invalid');
    if (estDelivered && estFulfillment && estDelivered < estFulfillment) errs.push('Estimated Arrival Date must be >= Estimated ShipDate');

    const rows = Array.isArray(lines) ? lines : [];
    if (rows.length === 0) errs.push('Please add at least one pallet');
    const anyQty = rows.some(l => Number(l.qty) > 0);
    if (!anyQty) errs.push('Please add at least one pallet with quantity > 0');
    const anyInvalid = rows.some(l => !Number.isFinite(Number(l.qty)) || Number(l.qty) <= 0);
    if (anyInvalid) errs.push('Each pallet quantity must be > 0');

    if (errs.length) { toast.error(errs[0]); return; }

    try {
      const payload = {
        status,
        createdAt,
        estFulfillment,
        estDelivered,
        customerEmail,
        customerName,
        customerPhone,
        shippingAddress,
        originalPrice,
        shippingPercent,
        discountPercent,
        notes,
        lines: lines.map(l => ({ ...l, qty: Math.max(0, Math.floor(Number(l.qty)||0)) })),
      };
      let data;
      if (editingOrder && editingOrder.id) {
        // Update existing order
        ({ data } = await api.put(`/early-buy/${encodeURIComponent(editingOrder.id)}`, payload));
        toast.success(`Early Buy order ${String(data?.id || editingOrder.id)} updated`);
      } else {
        // Create new order
        ({ data } = await api.post('/early-buy', payload));
        toast.success(`Early Buy order ${String(data?.id || '')} saved`);
      }
      await refreshOrders();
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to save Early Buy order';
      toast.error(msg);
      return;
    }
    setOpen(false);
    setEditingOrder(null);
    // reset form
    setStatus('processing');
    setCustomerEmail(''); setCustomerName(''); setCustomerPhone(''); setShippingAddress('');
    setCreatedAt(new Date().toISOString().slice(0,10)); setEstFulfillment(''); setEstDelivered('');
    setOriginalPrice(''); setShippingPercent(''); setDiscountPercent(''); setNotes('');
    setLines([]);
  };

  const exportOrderXlsx = (row: EarlyOrder) => {
    const rows: any[] = [];
    const addKV = (k: string, v: any) => rows.push([k, v ?? '']);
    addKV('Order #', row.id);
    addKV('Status', String(row.status || '').toUpperCase());
    addKV('Customer Email', row.customerEmail);
    addKV('Customer Name', row.customerName);
    addKV('Phone Number', row.customerPhone);
    addKV('Created Order Date', row.createdAt);
    addKV('Estimated Shipdate for Customer', row.estFulfillment);
    addKV('Estimated Arrival Date', row.estDelivered);
    addKV('Original Price', row.originalPrice);
    addKV('Shipping Charges (%)', row.shippingPercent);
    addKV('Discount (%)', row.discountPercent);
    addKV('Shipping Address', row.shippingAddress);
    addKV('Remarks/Notes', row.notes);

    rows.push([]);
    rows.push(['Pallets to Order']);
    rows.push(['Pallet ID', 'Pallet Name', 'Pallet Description', 'Qty Ordered']);
    for (const l of (row.lines||[])) {
      rows.push([l.lineItem, l.palletName, l.groupName, Number(l.qty||0)]);
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Early Buy Order');
    XLSX.writeFile(wb, `${row.id}.xlsx`);
  };

  return (
    <Container maxWidth="lg" sx={{ mt: 2, mb: 4 }}>
      <Typography variant="h5" sx={{ mb: 2, fontWeight: 700 }}>EARLY BUY</Typography>
      <Paper sx={{ p: 2, mb: 2 }}>
        <Stack direction="row" spacing={1} alignItems="center">
          <Button variant="contained" onClick={() => setOpen(true)}>Add Order</Button>
        </Stack>
      </Paper>

      <Paper sx={{ p: 2 }}>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} sx={{ mb: 1 }}>
          <Box sx={{ flex: 1 }} />
          <TextField size="small" label="Search" value={tableQ} onChange={(e)=> setTableQ(e.target.value)} sx={{ minWidth: 240 }} />
          <TextField size="small" select label="Status" value={tableStatus} onChange={(e)=> setTableStatus(e.target.value as any)} sx={{ width: 220 }}>
            <MenuItem value="all">All</MenuItem>
            <MenuItem value="processing">PROCESSING</MenuItem>
            <MenuItem value="ready_to_ship">READY TO SHIP</MenuItem>
            <MenuItem value="shipped">SHIPPED</MenuItem>
            <MenuItem value="completed">COMPLETED</MenuItem>
            <MenuItem value="canceled">CANCELED</MenuItem>
          </TextField>
        </Stack>
        <div style={{ height: 460, width: '100%' }}>
          <DataGrid
            rows={(filteredOrders||[]).map(o => {
              const arr = Array.isArray(o.lines) ? o.lines : [];
              const lc = typeof (o as any).linesCount === 'number' ? (o as any).linesCount : arr.length;
              const qt = typeof (o as any).qtyTotal === 'number' ? (o as any).qtyTotal : arr.reduce((s, l: any) => s + Math.max(0, Math.floor(Number(l?.qty || 0))), 0);
              return { ...o, id: o.id, linesCount: lc, qtyTotal: qt } as any;
            })}
            columns={ordersColumns}
            columnHeaderHeight={56}
            disableRowSelectionOnClick
            density="compact"
            slots={{ toolbar: GridToolbar }}
            slotProps={{ toolbar: { showQuickFilter: true, quickFilterProps: { debounceMs: 250 } } as any }}
            sx={{
              '& .MuiDataGrid-columnHeaderTitle': {
                whiteSpace: 'normal',
                lineHeight: 1.1,
                textOverflow: 'clip',
              },
              '& .MuiDataGrid-columnHeader': {
                whiteSpace: 'normal',
                alignItems: 'center',
              },
            }}
            onRowDoubleClick={(p) => {
              const row = orders.find(o => o.id === p.row.id);
              if (!row) return;
              // populate form for view-only
              setEditingOrder(row);
              setStatus(row.status);
              setCustomerEmail(row.customerEmail);
              setCustomerName(row.customerName);
              setCustomerPhone(row.customerPhone);
              setShippingAddress(row.shippingAddress);
              setCreatedAt(row.createdAt);
              setEstFulfillment(row.estFulfillment);
              setEstDelivered(row.estDelivered);
              setOriginalPrice(String(row.originalPrice||''));
              setShippingPercent(String(row.shippingPercent||''));
              setDiscountPercent(String(row.discountPercent||''));
              setNotes(String(row.notes||''));
              setLines(Array.isArray(row.lines) ? row.lines.map(l => ({ ...l })) : []);
              setOpen(true);
            }}
          />
        </div>
      </Paper>

      <Dialog open={open} onClose={() => { setOpen(false); setEditingOrder(null); }} fullWidth maxWidth="lg">
        <DialogTitle>{editingOrder ? `View Early Buy Order ${editingOrder.id}` : 'New Early Buy Order'}</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <Stack direction={{ xs:'column', sm:'row' }} spacing={2}>
              <TextField size="small" label="Warehouse" value="MPG" disabled fullWidth />
              {/* Status is always editable */}
              <TextField size="small" select label="Status" value={status} onChange={(e)=> setStatus(e.target.value as any)} fullWidth>
                {STATUS_OPTIONS.map(op => (<MenuItem key={op.value} value={op.value}>{op.label}</MenuItem>))}
              </TextField>
            </Stack>
            <Stack direction={{ xs:'column', sm:'row' }} spacing={2}>
              <TextField
                size="small"
                required
                label="Customer Email"
                value={customerEmail}
                onChange={(e)=> setCustomerEmail(e.target.value)}
                fullWidth
                disabled={!isEditable}
                error={Boolean(customerEmail) && !isValidEmail(customerEmail)}
                helperText={Boolean(customerEmail) && !isValidEmail(customerEmail) ? 'Enter a valid email address' : ''}
              />
              <TextField size="small" required label="Customer Name" value={customerName} onChange={(e)=> setCustomerName(e.target.value)} fullWidth disabled={!isEditable} />
            </Stack>
            <Stack direction={{ xs:'column', sm:'row' }} spacing={2}>
              <TextField size="small" required label="Phone Number" value={customerPhone} onChange={(e)=> setCustomerPhone(e.target.value)} fullWidth disabled={!isEditable} />
            </Stack>
            <TextField size="small" required label="Shipping Address" value={shippingAddress} onChange={(e)=> setShippingAddress(e.target.value)} fullWidth multiline minRows={2} disabled={!isEditable} />
            <Stack direction={{ xs:'column', sm:'row' }} spacing={2}>
              <TextField size="small" required type="date" label="Created Order Date" value={createdAt} onChange={(e)=> setCreatedAt(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isEditable} />
              <TextField size="small" required type="date" label="Estimated ShipDate for Customer" value={estFulfillment} onChange={(e)=> setEstFulfillment(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isEditable} />
              <TextField size="small" type="date" label="Estimated Arrival Date" value={estDelivered} onChange={(e)=> setEstDelivered(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isEditable} />
            </Stack>
            <Stack direction={{ xs:'column', md:'row' }} spacing={2}>
              <TextField size="small" label="Original Price" value={originalPrice} onChange={(e)=> setOriginalPrice(e.target.value)} fullWidth disabled={!isEditable} />
              <TextField size="small" label="Shipping Charges (%)" value={shippingPercent} onChange={(e)=> setShippingPercent(e.target.value)} fullWidth disabled={!isEditable} />
              <TextField size="small" label="Discount (%)" value={discountPercent} onChange={(e)=> setDiscountPercent(e.target.value)} fullWidth disabled={!isEditable} />
              <TextField size="small" label="Final Price" value={computedFinalPrice} fullWidth disabled />
            </Stack>
            <TextField size="small" label="Remarks/Notes" value={notes} onChange={(e)=> setNotes(e.target.value)} fullWidth multiline minRows={3} disabled={!isEditable} />

            <Stack direction="row" spacing={1} alignItems="center">
              <Button variant="outlined" onClick={()=> setPickerOpen(true)} disabled={!isEditable}>Add to List</Button>
              <Box sx={{ flex: 1 }} />
            </Stack>

            <div style={{ height: 300, width: '100%' }}>
              <DataGrid
                rows={lines.map((l, idx) => ({ id: idx+1, ...l }))}
                columns={linesColumns.map(c => c.field === 'actions' ? { ...c, renderCell: (p:any) => (!isEditable ? null : (c as any).renderCell(p)) } : (c.field === 'qty' ? { ...c, renderCell: (p:any) => {
                  if (!isEditable) return <span>{Number(lines[Math.max(0, Number(p?.id)-1)]?.qty || 0)}</span>;
                  return (c as any).renderCell(p);
                } } : c))}
                disableRowSelectionOnClick
                density="compact"
              />
            </div>
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={saveNewOrder}>Save</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={pickerOpen} onClose={()=> setPickerOpen(false)} fullWidth maxWidth="xl">
        <DialogTitle>Select Pallets</DialogTitle>
        <DialogContent>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} sx={{ mb: 1, mt: 1 }}>
            <TextField size="small" label="Search Pallet ID / Pallet Description / Pallet Name" value={pickerQ} onChange={(e)=> setPickerQ(e.target.value)} sx={{ flex: 1, minWidth: 260 }} />
          </Stack>
          <div style={{ height: 460, width: '100%' }}>
            <DG
              rows={Array.isArray(pickerRowsFiltered) ? pickerRowsFiltered : []}
              getRowId={(r: any) => {
                const l = String(r?.lineItem || '').trim();
                if (l) return l;
                const g = String(r?.groupName || '').trim();
                if (g) return g;
                const p = String(r?.palletName || '').trim();
                if (p) return p;
                const base = String(r?._id || r?.id || '').trim();
                return base || Math.random().toString(36).slice(2);
              }}
              columns={([
                { field: 'palletName', headerName: 'Pallet Name', flex: 1, minWidth: 200, sortable: true },
                { field: 'groupName', headerName: 'Pallet Description', flex: 1.2, minWidth: 260, sortable: true },
                { field: 'lineItem', headerName: 'Pallet ID', width: 160, sortable: true },
              ]) as GridColDef[]}
              checkboxSelection
              onRowSelectionModelChange={(sel: any, _details?: any) => {
                let arr: string[] = [];
                if (Array.isArray(sel)) {
                  arr = sel.map((v:any)=>String(v));
                } else if (sel && Array.isArray(sel.selectionModel)) {
                  arr = sel.selectionModel.map((v:any)=>String(v));
                } else if (sel && sel?.ids) {
                  const ids = sel.ids;
                  if (Array.isArray(ids)) arr = ids.map((v:any)=>String(v));
                  else if (ids instanceof Set) arr = Array.from(ids).map((v:any)=>String(v));
                }
                // TEMP: debug selection payload
                try { console.debug('[EarlyBuy picker] selection payload:', sel, 'normalized:', arr); } catch {}
                setSelectedGroups(arr);
              }}
              disableRowSelectionOnClick
              density="compact"
              loading={pickerLoading}
              pagination
              pageSizeOptions={[10, 25, 50, 100]}
              slots={{ toolbar: GridToolbar }}
              slotProps={{ toolbar: { showQuickFilter: true, quickFilterProps: { debounceMs: 250 } } as any }}
            />
          </div>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setPickerOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={addSelectedToLines} disabled={selectedGroups.length === 0}>Add</Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
