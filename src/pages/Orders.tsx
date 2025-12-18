import { useEffect, useMemo, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, IconButton, Dialog, DialogTitle, DialogContent, DialogActions, LinearProgress, MenuItem, Box, Chip } from '@mui/material';
import Autocomplete from '@mui/material/Autocomplete';
import DeleteIcon from '@mui/icons-material/Delete';
import { DataGrid, GridColDef } from '@mui/x-data-grid';
import api from '../api';
import { useToast } from '../components/ToastProvider';

type Warehouse = { _id: string; name: string };
type OrdersRow = {
  id: string;
  rawId: string;
  orderNumber: string;
  type: 'import' | 'manual';
  status: string;
  warehouseId: string;
  warehouseName?: string;
  createdAt: string;
  dateCreated?: string;
  refDate: string;
  lineCount: number;
  totalQty: number;
  email: string;
  lines: Array<{ groupName?: string; lineItem?: string; qty?: number }>;
  customerName?: string;
  customerPhone?: string;
  shippingAddress?: string;
  createdAtOrder?: string;
  estFulfillmentDate?: string;
  source?: string;
};

export default function Orders() {
  const todayYmd = useMemo(() => new Date().toISOString().slice(0, 10), []);
  const [warehouses, setWarehouses] = useState<Warehouse[]>([]);
  const [ordersLoading, setOrdersLoading] = useState(false);
  const [ordersRows, setOrdersRows] = useState<OrdersRow[]>([]);
  const [ordersQ, setOrdersQ] = useState('');
  const [ordersStatusFilter, setOrdersStatusFilter] = useState<'all' | 'create' | 'backorder' | 'fulfilled' | 'cancel'>('all');
  const [detailsOpen, setDetailsOpen] = useState(false);
  const [detailsRow, setDetailsRow] = useState<OrdersRow | null>(null);
  const [detailsStatus, setDetailsStatus] = useState('');
  const [detailsCustomerName, setDetailsCustomerName] = useState('');
  const [detailsCustomerEmail, setDetailsCustomerEmail] = useState('');
  const [detailsCustomerPhone, setDetailsCustomerPhone] = useState('');
  const [detailsCreatedAtOrder, setDetailsCreatedAtOrder] = useState('');
  const [detailsEstFulfillment, setDetailsEstFulfillment] = useState('');
  const [detailsShippingAddress, setDetailsShippingAddress] = useState('');
  const [csvOpen, setCsvOpen] = useState(false);
  const [csvWarehouseId, setCsvWarehouseId] = useState('');
  const [csvFile, setCsvFile] = useState<File | null>(null);
  const [csvPreview, setCsvPreview] = useState<any>(null);
  const [csvErrors, setCsvErrors] = useState<any[]>([]);
  const [csvLoading, setCsvLoading] = useState(false);
  const [csvCommitting, setCsvCommitting] = useState(false);

  const [manualOpen, setManualOpen] = useState(false);
  const [manualWarehouseId, setManualWarehouseId] = useState('');
  const [manualStatus, setManualStatus] = useState<'create' | 'backorder'>('create');
  const [manualCustomerEmail, setManualCustomerEmail] = useState('');
  const [manualCustomerName, setManualCustomerName] = useState('');
  const [manualCustomerPhone, setManualCustomerPhone] = useState('');
  const [manualCreatedAt, setManualCreatedAt] = useState('');
  const [manualEstFulfillment, setManualEstFulfillment] = useState('');
  const [manualShippingAddress, setManualShippingAddress] = useState('');
  const [manualLineMode, setManualLineMode] = useState<'pallet_group' | 'line_item'>('pallet_group');
  const [manualPalletGroupOptions, setManualPalletGroupOptions] = useState<string[]>([]);
  const [manualLineItemOptions, setManualLineItemOptions] = useState<string[]>([]);
  const [manualLines, setManualLines] = useState<Array<{ lineItem: string; qty: string }>>([{ lineItem: '', qty: '' }]);
  const [manualAvailable, setManualAvailable] = useState<Record<string, number>>({});
  const [manualValidationErrors, setManualValidationErrors] = useState<string[]>([]);

  const fixedWarehouseId = useMemo(() => {
    const wh = Array.isArray(warehouses) ? warehouses : [];
    const preferred = wh.find((w) => String(w?.name || '').trim().toLowerCase() === 'mpg planters');
    return String((preferred || wh[0] || ({} as any))._id || '');
  }, [warehouses]);

  const isValidEmail = (email: string) => {
    const s = String(email || '').trim();
    if (!s) return false;
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
  };

  const normalizeManualLine = (v: any) => String(v || '').trim();
  const getActiveManualOptions = () => (manualLineMode === 'pallet_group' ? manualPalletGroupOptions : manualLineItemOptions);
  const getManualRowError = (idx: number) => {
    const current = normalizeManualLine(manualLines[idx]?.lineItem);
    if (!current) return '';
    const matches = manualLines.filter((ln, i) => i !== idx && normalizeManualLine(ln?.lineItem).toLowerCase() === current.toLowerCase());
    return matches.length ? 'Duplicate pallet description / pallet id' : '';
  };

  useEffect(() => {
    const loadAvail = async () => {
      if (!manualWarehouseId) { setManualAvailable({}); return; }
      try {
        const { data } = await api.get('/pallet-inventory/groups');
        const map: Record<string, number> = {};
        (Array.isArray(data) ? data : []).forEach((g: any) => {
          const groupName = String(g.groupName || '').trim();
          if (!groupName) return;
          const per = Array.isArray(g.perWarehouse) ? g.perWarehouse : [];
          const rec = per.find((p: any) => String(p.warehouseId) === String(manualWarehouseId));
          map[groupName] = Number(rec?.pallets || 0);
        });
        setManualAvailable(map);
      } catch {
        setManualAvailable({});
      }
    };
    loadAvail();
  }, [manualWarehouseId]);
  const getManualRowOptions = (idx: number) => {
    const current = normalizeManualLine(manualLines[idx]?.lineItem);
    const selected = new Set(
      manualLines
        .filter((_, i) => i !== idx)
        .map((ln) => normalizeManualLine(ln?.lineItem).toLowerCase())
        .filter((v) => v)
    );
    return (getActiveManualOptions() || []).filter((o) => {
      const name = normalizeManualLine(o);
      if (!name) return false;
      if (current && name.toLowerCase() === current.toLowerCase()) return true;
      return !selected.has(name.toLowerCase());
    });
  };

  const toast = useToast();

  const normalizeStatus = (v: any) => {
    const s = String(v || '').trim().toLowerCase();
    if (!s) return '';
    if (s === 'created') return 'create';
    if (s === 'cancelled') return 'cancel';
    return s;
  };

  const openDetails = (row: OrdersRow) => {
    setDetailsRow(row);
    setDetailsStatus(normalizeStatus(row.status || ''));
    setDetailsCustomerName(String(row.customerName || ''));
    setDetailsCustomerEmail(String(row.email || ''));
    setDetailsCustomerPhone(String(row.customerPhone || ''));
    setDetailsShippingAddress(String(row.shippingAddress || ''));
    setDetailsEstFulfillment(row.estFulfillmentDate ? String(row.estFulfillmentDate).slice(0, 10) : '');
    setDetailsCreatedAtOrder(row.createdAtOrder ? String(row.createdAtOrder).slice(0, 10) : '');
    setDetailsOpen(true);
  };

  const saveDetailsStatus = async () => {
    if (!detailsRow) return;
    try {
      const next = normalizeStatus(detailsStatus);
      const prev = normalizeStatus(detailsRow.status || '');

      if ((next === 'create' || next === 'backorder') && prev === 'cancel') {
        const msg = next === 'backorder'
          ? 'Changing status to backorder will deduct inventory immediately and can result in negative pallet stock. Proceed?'
          : 'Changing status to create will deduct inventory immediately. Proceed?';
        const ok = window.confirm(msg);
        if (!ok) return;
      }

      if (next === 'cancel' && prev !== 'cancel') {
        const ok = window.confirm('Changing status to cancel will restore inventory previously deducted for this order. Proceed?');
        if (!ok) return;
      }

      const rawId = encodeURIComponent(detailsRow.rawId);
      if (detailsRow.id.startsWith('manual:')) {
        // If transitioning to fulfilled, update details first (otherwise the server locks details updates).
        if (next === 'fulfilled' && prev !== 'fulfilled') {
          await api.put(`/orders/unfulfilled/${rawId}`, {
            customerName: detailsCustomerName,
            customerEmail: detailsCustomerEmail,
            customerPhone: detailsCustomerPhone,
            estFulfillmentDate: detailsEstFulfillment || null,
            shippingAddress: detailsShippingAddress,
          });
          await api.put(`/orders/unfulfilled/${rawId}/status`, { status: next });
        } else {
          await api.put(`/orders/unfulfilled/${rawId}/status`, { status: next });
          await api.put(`/orders/unfulfilled/${rawId}`, {
            customerName: detailsCustomerName,
            customerEmail: detailsCustomerEmail,
            customerPhone: detailsCustomerPhone,
            estFulfillmentDate: detailsEstFulfillment || null,
            shippingAddress: detailsShippingAddress,
          });
        }
      } else {
        // If transitioning to fulfilled, update details first (otherwise the server locks details updates).
        if (next === 'fulfilled' && prev !== 'fulfilled') {
          await api.put(`/orders/fulfilled/imports/${rawId}`, {
            fulfilledAt: detailsEstFulfillment || null,
          });
          await api.put(`/orders/fulfilled/imports/${rawId}/status`, { status: next });
        } else {
          await api.put(`/orders/fulfilled/imports/${rawId}/status`, { status: next });
          await api.put(`/orders/fulfilled/imports/${rawId}`, {
            fulfilledAt: detailsEstFulfillment || null,
          });
        }
      }

      toast.success('Saved');
      setDetailsOpen(false);
      await loadOrders();
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to save';
      toast.error(msg);
    }
  };

  useEffect(() => {
    if (!fixedWarehouseId) return;
    if (csvOpen) setCsvWarehouseId(fixedWarehouseId);
    if (manualOpen) setManualWarehouseId(fixedWarehouseId);
  }, [fixedWarehouseId, csvOpen, manualOpen]);

  const loadOrders = async (warehousesSnapshot?: Warehouse[]) => {
    setOrdersLoading(true);
    try {
      const [uRes, fRes] = await Promise.all([
        api.get<any[]>('/orders/unfulfilled'),
        api.get<any[]>('/orders/fulfilled/imports'),
      ]);
      const unfulfilled = Array.isArray(uRes?.data) ? uRes.data : [];
      const fulfilled = Array.isArray(fRes?.data) ? fRes.data : [];

      const whList = Array.isArray(warehousesSnapshot) ? warehousesSnapshot : warehouses;
      const whNameById = new Map((whList || []).map((w) => [String(w._id), String(w.name)]));

      const mappedUnfulfilled: OrdersRow[] = unfulfilled.map((o: any) => ({
        id: `manual:${String(o?._id || o?.orderNumber || Math.random())}`,
        rawId: String(o?._id || ''),
        orderNumber: String(o?.orderNumber || ''),
        type: 'manual',
        status: normalizeStatus(o?.status || 'create') || 'create',
        warehouseId: normalizeId(o?.warehouseId),
        warehouseName: whNameById.get(normalizeId(o?.warehouseId)) || '',
        createdAt: normalizeDateValue(o?.createdAt),
        dateCreated: normalizeDateValue(o?.createdAtOrder || o?.createdAt),
        refDate: String(o?.createdAtOrder || o?.createdAt || ''),
        lineCount: Array.isArray(o?.lines) ? o.lines.length : 0,
        totalQty: Array.isArray(o?.lines) ? o.lines.reduce((s: number, l: any) => s + (Number(l?.qty) || 0), 0) : 0,
        email: String(o?.customerEmail || ''),
        lines: Array.isArray(o?.lines) ? o.lines : [],
        customerName: String(o?.customerName || ''),
        customerPhone: String(o?.customerPhone || ''),
        shippingAddress: String(o?.shippingAddress || ''),
        createdAtOrder: normalizeDateValue(o?.createdAtOrder),
        estFulfillmentDate: normalizeDateValue(o?.estFulfillmentDate),
        source: 'manual',
      }));

      const mappedFulfilled: OrdersRow[] = fulfilled.map((o: any) => ({
        id: `import:${String(o?._id || o?.orderNumber || Math.random())}`,
        rawId: String(o?._id || ''),
        orderNumber: String(o?.orderNumber || ''),
        type: String(o?.source || '') === 'csv' ? 'import' : 'manual',
        status: normalizeStatus(o?.status || 'create') || 'create',
        warehouseId: normalizeId(o?.warehouseId),
        warehouseName: whNameById.get(normalizeId(o?.warehouseId)) || '',
        createdAt: normalizeDateValue(o?.createdAt),
        dateCreated: normalizeDateValue(o?.createdAtOrder || o?.createdAt),
        refDate: String(o?.createdAtOrder || o?.createdAt || ''),
        lineCount: Array.isArray(o?.lines) ? o.lines.length : 0,
        totalQty: Array.isArray(o?.lines) ? o.lines.reduce((s: number, l: any) => s + (Number(l?.qty) || 0), 0) : 0,
        email: String(o?.email || ''),
        lines: Array.isArray(o?.lines) ? o.lines : [],
        customerName: String(o?.billingName || o?.shippingName || ''),
        customerPhone: String(o?.billingPhone || o?.shippingPhone || ''),
        shippingAddress: String(o?.shippingStreet || o?.shippingAddress1 || ''),
        createdAtOrder: normalizeDateValue(o?.createdAtOrder),
        estFulfillmentDate: normalizeDateValue(o?.fulfilledAt),
        source: String(o?.source || ''),
      }));

      const merged = [...mappedUnfulfilled, ...mappedFulfilled].sort((a, b) => {
        const ad = new Date(a.createdAt || a.refDate || 0).getTime();
        const bd = new Date(b.createdAt || b.refDate || 0).getTime();
        return bd - ad;
      });
      setOrdersRows(merged);
    } catch {
      setOrdersRows([]);
    } finally {
      setOrdersLoading(false);
    }
  };

  useEffect(() => {
    const load = async () => {
      try {
        const { data } = await api.get<Warehouse[]>('/warehouses');
        const list = Array.isArray(data) ? data : [];
        setWarehouses(list);
        await loadOrders(list);
      } catch {
        setWarehouses([]);
        await loadOrders([]);
      }
    };
    load();
  }, []);

  const normalizeId = (v: any) => {
    if (!v) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'object') {
      const maybe = (v as any)._id || (v as any).id;
      if (maybe) return String(maybe);
    }
    return String(v);
  };

  const normalizeDateValue = (v: any) => {
    if (!v) return '';
    if (v instanceof Date) return v.toISOString();
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return new Date(v).toISOString();
    if (typeof v === 'object') {
      const d = (v as any).$date || (v as any).date || (v as any).value;
      if (d) return String(d);
      // if it's a mongoose date-like object
      if (typeof (v as any).toISOString === 'function') {
        try { return (v as any).toISOString(); } catch {}
      }
    }
    return '';
  };

  const fmtDate = (v: any) => {
    const s = normalizeDateValue(v);
    if (!s) return '-';
    const d = new Date(s);
    return d && !Number.isNaN(d.getTime()) ? d.toLocaleDateString() : '-';

  };

  const ordersColumns: GridColDef[] = useMemo(() => [
    { field: 'orderNumber', headerName: 'Order #', flex: 1, minWidth: 170 },
    { field: 'type', headerName: 'Type', width: 120 },
    {
      field: 'status',
      headerName: 'Status',
      width: 140,
      renderCell: (p: any) => {
        const s = normalizeStatus((p?.row as any)?.status || '');
        const label = s || '-';
        const color: any = s === 'fulfilled' ? 'success' : s === 'cancel' ? 'error' : s === 'backorder' ? 'warning' : 'default';
        const variant: any = s === 'create' ? 'outlined' : 'filled';
        return <Chip size="small" label={label} color={color} variant={variant} />;
      },
    },
    {
      field: 'dateCreated',
      headerName: 'Order Created',
      width: 150,
      sortable: true,
      renderCell: (p: any) => fmtDate((p?.row as any)?.dateCreated),
    },
    {
      field: 'estFulfillmentDate',
      headerName: 'Estimated Shipdate for Customer',
      width: 210,
      renderCell: (p: any) => fmtDate((p?.row as any)?.estFulfillmentDate),
    },
    {
      field: 'warehouseName',
      headerName: 'Warehouse',
      flex: 1,
      minWidth: 160,
      renderCell: (p: any) => String((p?.row as any)?.warehouseName || '-')
    },
    { field: 'email', headerName: 'Email', flex: 1, minWidth: 190 },
    { field: 'lineCount', headerName: 'Lines', width: 90, type: 'number' },
    { field: 'totalQty', headerName: 'Qty', width: 90, type: 'number' },
    {
      field: 'actions',
      headerName: 'Actions',
      width: 110,
      sortable: false,
      filterable: false,
      renderCell: (params: any) => (
        <Button size="small" variant="outlined" onClick={() => openDetails(params.row as OrdersRow)}>View</Button>
      )
    },
  ], [warehouses]);

  const filteredOrdersRows = useMemo(() => {
    const t = ordersQ.trim().toLowerCase();
    return ordersRows.filter((r) => {
      if (ordersStatusFilter !== 'all' && normalizeStatus(r.status || '') !== ordersStatusFilter) return false;
      if (!t) return true;
      const hay = [
        r.orderNumber,
        r.type,
        r.status,
        r.email,
        r.warehouseName || '',
      ].join(' ').toLowerCase();
      return hay.includes(t);
    });
  }, [ordersQ, ordersRows, ordersStatusFilter]);

  const openCsv = () => {
    setCsvOpen(true);
    setCsvWarehouseId(fixedWarehouseId);
    setCsvFile(null);
    setCsvPreview(null);
    setCsvErrors([]);
  };

  const openManual = async () => {
    setManualOpen(true);
    setManualWarehouseId(fixedWarehouseId);
    setManualStatus('create');
    setManualCustomerEmail('');
    setManualCustomerName('');
    setManualCustomerPhone('');
    setManualCreatedAt(new Date().toISOString().slice(0, 10));
    setManualEstFulfillment('');
    setManualShippingAddress('');
    setManualLineMode('pallet_group');
    setManualLines([{ lineItem: '', qty: '' }]);
    setManualAvailable({});
    setManualValidationErrors([]);
    try {
      const { data } = await api.get<any[]>('/item-groups');
      const palletGroups = (Array.isArray(data) ? data : [])
        .map((g: any) => String(g?.name || '').trim())
        .filter((v: string) => v);
      const lineItems = (Array.isArray(data) ? data : [])
        .map((g: any) => String(g?.lineItem || '').trim())
        .filter((v: string) => v);

      setManualPalletGroupOptions(Array.from(new Set(palletGroups)).sort((a, b) => a.localeCompare(b)));
      setManualLineItemOptions(Array.from(new Set(lineItems)).sort((a, b) => a.localeCompare(b)));
    } catch {
      setManualPalletGroupOptions([]);
      setManualLineItemOptions([]);
    }
  };

  const onCsvFileChange = (e: any) => {
    setCsvFile(e.target.files?.[0] || null);
    setCsvPreview(null);
    setCsvErrors([]);
  };

  const doCsvPreview = async () => {
    if (!csvFile) return;
    if (!csvWarehouseId) { toast.error('Please select a Warehouse'); return; }
    setCsvLoading(true);
    setCsvErrors([]);
    try {
      const form = new FormData();
      form.append('file', csvFile);
      const params = new URLSearchParams();
      params.set('warehouseId', csvWarehouseId);
      const { data } = await api.post(`/orders/fulfilled/preview?${params.toString()}`, form, { headers: { 'Content-Type': 'multipart/form-data' } });
      setCsvPreview(data);
      setCsvErrors(data.errors || []);
      toast.success('Preview parsed successfully');
    } catch (e: any) {
      const payload = e?.response?.data;
      setCsvPreview(null);
      setCsvErrors(payload?.errors || [{ rowNum: '-', errors: [payload?.message || 'Preview failed'] }]);
      toast.error(payload?.message || 'Preview failed');
    } finally {
      setCsvLoading(false);
    }
  };

  const doCsvCommit = async () => {
    if (!csvFile) return;
    if (!csvWarehouseId) { toast.error('Please select a Warehouse'); return; }
    setCsvCommitting(true);
    try {
      const form = new FormData();
      form.append('file', csvFile);
      const params = new URLSearchParams();
      params.set('warehouseId', csvWarehouseId);
      const { data } = await api.post(`/orders/fulfilled?${params.toString()}`, form, { headers: { 'Content-Type': 'multipart/form-data' } });
      toast.success(`Imported ${data?.committedOrders || 0} order(s)`);
      setCsvOpen(false);
      await loadOrders();
    } catch (e:any) {
      const payload = e?.response?.data;
      setCsvErrors(payload?.errors || [{ rowNum: '-', errors: [payload?.message || 'Import failed'] }]);
      toast.error(payload?.message || 'Import failed');
    } finally {
      setCsvCommitting(false);
    }
  };

  const submitManual = async () => {
    const errs: string[] = [];
    setManualValidationErrors([]);
    if (!manualWarehouseId) {
      errs.push('Warehouse is required');
    }
    if (!isValidEmail(manualCustomerEmail)) {
      errs.push('Customer Email is required and must be a valid email');
    }
    if (!String(manualCustomerName || '').trim()) {
      errs.push('Customer Name is required');
    }
    if (!String(manualCustomerPhone || '').trim()) {
      errs.push('Phone Number is required');
    }
    if (!manualCreatedAt) {
      errs.push('Created Order Date is required');
    } else if (manualCreatedAt > todayYmd) {
      errs.push('Created Order Date cannot be an advance date');
    }
    if (!manualEstFulfillment) {
      errs.push('Estimated Shipdate for Customer is required');
    } else if (manualEstFulfillment < todayYmd) {
      errs.push('Estimated Shipdate for Customer must be today or later');
    }
    if (!String(manualShippingAddress || '').trim()) {
      errs.push('Shipping Address is required');
    }

    const parsed = manualLines
      .map((l) => ({ lineItem: String(l.lineItem || '').trim(), qty: Number(l.qty) }))
      .filter((l) => l.lineItem);
    if (!parsed.length) {
      errs.push('At least 1 Pallet ID is required');
    }
    for (const l of parsed) {
      if (!Number.isFinite(l.qty) || l.qty <= 0) {
        errs.push('Qty must be > 0');
        break;
      }
    }
    const seen = new Set<string>();
    for (const l of parsed) {
      const k = l.lineItem.toLowerCase();
      if (seen.has(k)) {
        errs.push(`Duplicate Pallet ID: ${l.lineItem}`);
        break;
      }
      seen.add(k);
    }

    if (errs.length) {
      setManualValidationErrors(errs);
      toast.error(errs[0]);
      return;
    }

    // Client-side availability check (server also enforces)
    // For backorder, allow negative stock, so skip this check.
    if (manualStatus !== 'backorder') {
      try {
        const byLineItem = new Map<string, string>();
        const { data } = await api.get<any[]>('/item-groups');
        (Array.isArray(data) ? data : []).forEach((g: any) => {
          const groupName = String(g?.name || '').trim();
          const li = String(g?.lineItem || '').trim();
          if (li && groupName) byLineItem.set(li.toLowerCase(), groupName);
          if (groupName) byLineItem.set(groupName.toLowerCase(), groupName);
        });
        const needs = new Map<string, number>();
        for (const l of parsed) {
          const group = byLineItem.get(String(l.lineItem || '').toLowerCase()) || '';
          if (!group) continue;
          needs.set(group, (needs.get(group) || 0) + (Number(l.qty) || 0));
        }
        const insufficient: string[] = [];
        for (const [groupName, qty] of needs.entries()) {
          const avail = Number(manualAvailable[groupName] || 0);
          if (avail < qty) insufficient.push(`${groupName} (need ${qty}, available ${avail})`);
        }
        if (insufficient.length) {
          setManualValidationErrors(insufficient.map((s) => `Insufficient stock: ${s}`));
          toast.error(`Insufficient stock in selected warehouse for: ${insufficient.join(', ')}`);
          return;
        }
      } catch {
        // If item-groups fetch fails, rely on server validation
      }
    }
    try {
      await api.post('/orders/unfulfilled', {
        warehouseId: manualWarehouseId,
        status: manualStatus,
        customerEmail: manualCustomerEmail.trim(),
        customerName: manualCustomerName.trim(),
        customerPhone: manualCustomerPhone.trim(),
        createdAtOrder: manualCreatedAt || undefined,
        estFulfillmentDate: manualEstFulfillment || undefined,
        shippingAddress: manualShippingAddress.trim(),
        lines: parsed.map((l) => ({ search: l.lineItem, qty: l.qty })),
      });
      toast.success('Order created');
      setManualOpen(false);
      await loadOrders();
    } catch (e:any) {
      const msg = e?.response?.data?.message || 'Failed to save';
      setManualValidationErrors([msg]);
      toast.error(msg);
    }
  };

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Orders</Typography>

      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} alignItems={{ xs: 'stretch', sm: 'center' }}>
          <Button variant="contained" onClick={openCsv}>Import Order CSV from MPGWholeSale</Button>
          <Button variant="outlined" onClick={openManual}>Add Order</Button>
          <Button variant="text" onClick={() => loadOrders()} disabled={ordersLoading}>Refresh List</Button>
          <Box sx={{ flex: 1 }} />
          <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} alignItems={{ xs: 'stretch', sm: 'center' }}>
            <TextField
              size="small"
              label="Search"
              value={ordersQ}
              onChange={(e)=>setOrdersQ(e.target.value)}
              sx={{ minWidth: 260 }}
            />
            <TextField
              select
              size="small"
              label="Status"
              value={ordersStatusFilter}
              onChange={(e)=>setOrdersStatusFilter(e.target.value as any)}
              sx={{ minWidth: 180 }}
            >
              <MenuItem value="all">All</MenuItem>
              <MenuItem value="create">create</MenuItem>
              <MenuItem value="backorder">backorder</MenuItem>
              <MenuItem value="fulfilled">fulfilled</MenuItem>
              <MenuItem value="cancel">cancel</MenuItem>
            </TextField>
          </Stack>
        </Stack>
      </Paper>

      <Paper sx={{ height: 520, width: '100%', p:1, mb: 2 }}>
        <DataGrid
          rows={filteredOrdersRows}
          columns={ordersColumns}
          loading={ordersLoading}
          disableRowSelectionOnClick
          density="compact"
          initialState={{ pagination: { paginationModel: { pageSize: 20, page: 0 } } }}
          pageSizeOptions={[10,20,50,100]}
          onRowDoubleClick={(p:any)=> openDetails(p.row as OrdersRow)}
        />
      </Paper>

      <Dialog open={detailsOpen} onClose={()=>setDetailsOpen(false)} fullWidth maxWidth="lg">
        <DialogTitle>Order Details</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          {detailsRow ? (
            <Stack spacing={2} sx={{ mt: 1 }}>
              <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2}>
                <TextField label="Order #" size="small" value={detailsRow.orderNumber} InputProps={{ readOnly: true }} fullWidth />
                <TextField label="Type" size="small" value={detailsRow.type} InputProps={{ readOnly: true }} fullWidth />
              </Stack>

              <TextField
                select
                label="Status"
                size="small"
                value={detailsStatus}
                onChange={(e)=>setDetailsStatus(e.target.value)}
                disabled={normalizeStatus(detailsRow.status || '') === 'fulfilled'}
                helperText={normalizeStatus(detailsRow.status || '') === 'fulfilled' ? 'Fulfilled orders are locked' : ''}
              >
                <MenuItem value="create">create</MenuItem>
                <MenuItem value="backorder">backorder</MenuItem>
                <MenuItem value="fulfilled">fulfilled</MenuItem>
                <MenuItem value="cancel">cancel</MenuItem>
              </TextField>

              <TextField
                label="Warehouse"
                size="small"
                value={warehouses.find(w => String(w._id) === String(detailsRow.warehouseId))?.name || '-'}
                InputProps={{ readOnly: true }}
              />

              <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2}>
                <TextField
                  label="Customer Name"
                  size="small"
                  fullWidth
                  value={detailsCustomerName}
                  onChange={(e)=>setDetailsCustomerName(e.target.value)}
                  disabled={normalizeStatus(detailsRow.status || '') === 'fulfilled' || detailsRow.type === 'import'}
                />
                <TextField
                  label="Customer Email"
                  size="small"
                  fullWidth
                  value={detailsCustomerEmail}
                  onChange={(e)=>setDetailsCustomerEmail(e.target.value)}
                  disabled={normalizeStatus(detailsRow.status || '') === 'fulfilled' || detailsRow.type === 'import'}
                />
              </Stack>

              <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2}>
                <TextField
                  label="Phone Number"
                  size="small"
                  fullWidth
                  value={detailsCustomerPhone}
                  onChange={(e)=>setDetailsCustomerPhone(e.target.value)}
                  disabled={normalizeStatus(detailsRow.status || '') === 'fulfilled' || detailsRow.type === 'import'}
                />
                <TextField
                  type="date"
                  label="Order Created"
                  InputLabelProps={{ shrink: true }}
                  size="small"
                  fullWidth
                  value={detailsCreatedAtOrder}
                  onChange={(e)=>setDetailsCreatedAtOrder(e.target.value)}
                  disabled={true}
                  helperText={detailsRow.id.startsWith('manual:') ? 'Manual orders use Created Order Date from creation' : 'Imported orders keep Order Created from CSV'}
                />
              </Stack>

              <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2}>
                <TextField
                  type="date"
                  label="Estimated Shipdate for Customer"
                  InputLabelProps={{ shrink: true }}
                  size="small"
                  fullWidth
                  value={detailsEstFulfillment}
                  onChange={(e)=>setDetailsEstFulfillment(e.target.value)}
                  disabled={normalizeStatus(detailsRow.status || '') === 'fulfilled'}
                />
                <TextField
                  label="Total Qty"
                  size="small"
                  fullWidth
                  value={String(detailsRow.totalQty)}
                  InputProps={{ readOnly: true }}
                />
              </Stack>

              <TextField
                label="Shipping Address"
                size="small"
                fullWidth
                multiline
                minRows={2}
                value={detailsShippingAddress}
                onChange={(e)=>setDetailsShippingAddress(e.target.value)}
                disabled={normalizeStatus(detailsRow.status || '') === 'fulfilled' || detailsRow.type === 'import'}
              />

              <Box>
                <Typography variant="subtitle2" sx={{ mb: 1 }}>Pallet IDs</Typography>
                <div style={{ height: 260, width: '100%' }}>
                  <DataGrid
                    rows={(detailsRow.lines || []).map((l: any, i: number) => ({
                      id: i,
                      groupName: String(l?.groupName || ''),
                      lineItem: String(l?.lineItem || ''),
                      qty: Number(l?.qty || 0),
                    }))}
                    columns={([ 
                      { field: 'groupName', headerName: 'Pallet Description', flex: 1, minWidth: 220 },
                      { field: 'lineItem', headerName: 'Pallet ID', flex: 1, minWidth: 220 },
                      { field: 'qty', headerName: 'Qty', width: 100, type: 'number' },
                    ]) as GridColDef[]}
                    disableRowSelectionOnClick
                    density="compact"
                    hideFooter
                  />
                </div>
              </Box>
            </Stack>
          ) : null}
        </DialogContent>
        <DialogActions>
          <Button onClick={()=>setDetailsOpen(false)}>Close</Button>
          <Button variant="contained" onClick={saveDetailsStatus} disabled={!detailsRow || !detailsStatus || normalizeStatus(detailsRow.status || '') === 'fulfilled'}>Save</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={csvOpen} onClose={()=>setCsvOpen(false)} fullWidth maxWidth="md">
        <DialogTitle>Import Orders (MPGWholeSale site) (.csv)</DialogTitle>
        <DialogContent>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb: 2, mt: 1 }}>
            <TextField select disabled label="Warehouse" size="small" sx={{ minWidth: 220 }} value={csvWarehouseId} onChange={(e)=>setCsvWarehouseId(e.target.value)} error={!csvWarehouseId} helperText={!csvWarehouseId ? 'Required' : ''}>
              {warehouses.map((w)=> (
                <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>
              ))}
            </TextField>
            <Box sx={{ display:'flex', alignItems:'center', gap:1, flex: 1 }}>
              <input type="file" accept=".csv" onChange={onCsvFileChange} />
            </Box>
          </Stack>
          {(csvLoading || csvCommitting) && <LinearProgress sx={{ mb: 2 }} />}
          <Box sx={{ display:'flex', gap:1, mb:2 }}>
            <Button variant="outlined" disabled={!csvFile || csvLoading} onClick={doCsvPreview}>Preview</Button>
            <Button variant="contained" disabled={!csvFile || !csvPreview || csvErrors.length > 0 || csvCommitting || !csvWarehouseId} onClick={doCsvCommit}>Commit</Button>
          </Box>
          {csvPreview && (
            <Typography variant="body2" sx={{ mb: 1 }}>
              Parsed rows: {csvPreview.totalRows} | Orders: {csvPreview.orderCount} | Errors: {csvPreview.errorCount}
            </Typography>
          )}
          {csvErrors.length > 0 && (
            <Box sx={{ p:1, bgcolor:'#fff4f4', border:'1px solid #f5c2c7', borderRadius:1 }}>
              <Typography variant="subtitle2" color="error">Validation Errors</Typography>
              <ul>
                {csvErrors.map((er: any, idx: number) => (
                  <li key={idx}><Typography variant="caption">Row {er.rowNum}: {(er.errors||[]).join(', ')}</Typography></li>
                ))}
              </ul>
            </Box>
          )}
        </DialogContent>
        <DialogActions>
          <Button onClick={()=>setCsvOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={manualOpen} onClose={()=>setManualOpen(false)} fullWidth maxWidth="md">
        <DialogTitle>Add Order</DialogTitle>
        <DialogContent>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb: 2, mt: 1 }}>
            <TextField select disabled label="Warehouse" size="small" sx={{ minWidth: 220 }} value={manualWarehouseId} onChange={(e)=>setManualWarehouseId(e.target.value)} error={!manualWarehouseId} helperText={!manualWarehouseId ? 'Required' : ''}>
              {warehouses.map((w)=> (
                <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>
              ))}
            </TextField>
            <TextField
              select
              label="Status"
              size="small"
              sx={{ minWidth: 180 }}
              value={manualStatus}
              onChange={(e)=>setManualStatus(e.target.value as any)}
              helperText={manualStatus === 'backorder' ? 'Backorder allows negative pallet stock' : 'Create requires sufficient stock'}
            >
              <MenuItem value="create">create</MenuItem>
              <MenuItem value="backorder">backorder</MenuItem>
            </TextField>
            <TextField label="Customer Email" size="small" value={manualCustomerEmail} onChange={(e)=>setManualCustomerEmail(e.target.value)} sx={{ minWidth: 200 }} error={!isValidEmail(manualCustomerEmail)} helperText={!isValidEmail(manualCustomerEmail) ? 'Required (valid email)' : ''} />
            <TextField label="Customer Name" size="small" value={manualCustomerName} onChange={(e)=>setManualCustomerName(e.target.value)} sx={{ minWidth: 200 }} error={!String(manualCustomerName||'').trim()} helperText={!String(manualCustomerName||'').trim() ? 'Required' : ''} />
          </Stack>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb: 2 }}>
            <TextField label="Phone Number" size="small" value={manualCustomerPhone} onChange={(e)=>setManualCustomerPhone(e.target.value)} sx={{ minWidth: 220 }} error={!String(manualCustomerPhone||'').trim()} helperText={!String(manualCustomerPhone||'').trim() ? 'Required' : ''} />
          </Stack>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb: 2 }}>
            <TextField type="date" label="Created Order Date" InputLabelProps={{ shrink: true }} size="small" value={manualCreatedAt} onChange={(e)=>setManualCreatedAt(e.target.value)} sx={{ minWidth: 220 }} inputProps={{ max: todayYmd }} error={!manualCreatedAt || manualCreatedAt > todayYmd} helperText={!manualCreatedAt ? 'Required' : (manualCreatedAt > todayYmd ? 'Cannot be advance date' : '')} />
            <TextField type="date" label="Estimated Shipdate for Customer" InputLabelProps={{ shrink: true }} size="small" value={manualEstFulfillment} onChange={(e)=>setManualEstFulfillment(e.target.value)} sx={{ minWidth: 220 }} inputProps={{ min: todayYmd }} error={!manualEstFulfillment || manualEstFulfillment < todayYmd} helperText={!manualEstFulfillment ? 'Required' : (manualEstFulfillment < todayYmd ? 'Must be today or later' : '')} />
          </Stack>
          <TextField
            fullWidth
            label="Shipping Address"
            size="small"
            value={manualShippingAddress}
            onChange={(e)=>setManualShippingAddress(e.target.value)}
            sx={{ mb: 2 }}
            error={!String(manualShippingAddress||'').trim()}
            helperText={!String(manualShippingAddress||'').trim() ? 'Required' : ''}
          />

          <Typography variant="subtitle1" sx={{ mb: 1 }}>Pallet IDs</Typography>

          {manualValidationErrors.length > 0 && (
            <Box sx={{ p:1, mb: 2, bgcolor:'#fff4f4', border:'1px solid #f5c2c7', borderRadius:1 }}>
              <Typography variant="subtitle2" color="error">Validation Errors</Typography>
              <ul style={{ marginTop: 6, marginBottom: 0 }}>
                {manualValidationErrors.map((m, idx) => (
                  <li key={idx}><Typography variant="caption">{m}</Typography></li>
                ))}
              </ul>
            </Box>
          )}

          <TextField
            select
            size="small"
            label="Add By"
            value={manualLineMode}
            onChange={(e)=> {
              const next = e.target.value as any;
              setManualLineMode(next);
              setManualLines((prev)=> prev.map((p)=> ({ ...p, lineItem: '' })));
            }}
            sx={{ mb: 1, minWidth: 240 }}
          >
            <MenuItem value="pallet_group">Option 1 - Pallet Description</MenuItem>
            <MenuItem value="line_item">Option 2 - Pallet IDs</MenuItem>
          </TextField>

          <Stack spacing={1}>
            {manualLines.map((ln, idx) => (
              <Stack
                key={idx}
                direction={{ xs:'column', sm:'row' }}
                spacing={2}
                alignItems={{ xs:'stretch', sm:'center' }}
                sx={{ p: 0.5, borderRadius: 1, bgcolor: getManualRowError(idx) ? '#fff4f4' : 'transparent' }}
              >
                <Autocomplete
                  freeSolo
                  options={getManualRowOptions(idx)}
                  value={ln.lineItem}
                  disabled={!manualWarehouseId}
                  onChange={(_, v)=> {
                    setManualValidationErrors([]);
                    setManualLines((prev)=> prev.map((p, i)=> i===idx ? { ...p, lineItem: String(v || '') } : p));
                  }}
                  onInputChange={(_, v)=> {
                    setManualValidationErrors([]);
                    setManualLines((prev)=> prev.map((p, i)=> i===idx ? { ...p, lineItem: String(v || '') } : p));
                  }}
                  renderInput={(params)=> (
                    <TextField
                      {...params}
                      label={manualLineMode === 'pallet_group' ? 'Pallet Description' : 'Pallet ID'}
                      size="small"
                      placeholder={!manualWarehouseId ? 'Select Warehouse first' : undefined}
                      error={Boolean(getManualRowError(idx))}
                      helperText={getManualRowError(idx) || ''}
                    />
                  )}
                  sx={{ flex: 1, minWidth: 260 }}
                />
                <TextField type="number" label="Qty" size="small" value={ln.qty} onChange={(e)=> { setManualValidationErrors([]); setManualLines((prev)=> prev.map((p, i)=> i===idx ? { ...p, qty: e.target.value } : p)); }} sx={{ minWidth: 120 }} />
                <IconButton color="error" onClick={()=> setManualLines((prev)=> prev.filter((_, i)=> i!==idx))} disabled={manualLines.length===1}><DeleteIcon /></IconButton>
              </Stack>
            ))}
            <Button onClick={()=> setManualLines((prev)=> [...prev, { lineItem: '', qty: '' }])}>Add Line</Button>
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=>setManualOpen(false)}>Cancel</Button>
          <Button
            variant="contained"
            onClick={submitManual}
            disabled={
              !manualWarehouseId ||
              !isValidEmail(manualCustomerEmail) ||
              !String(manualCustomerName||'').trim() ||
              !String(manualCustomerPhone||'').trim() ||
              !manualCreatedAt ||
              manualCreatedAt > todayYmd ||
              !manualEstFulfillment ||
              manualEstFulfillment < todayYmd ||
              !String(manualShippingAddress||'').trim() ||
              (manualLines.filter((l)=> String(l.lineItem||'').trim()).length === 0)
            }
          >
            Save
          </Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
