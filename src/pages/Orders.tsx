import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, IconButton, Dialog, DialogTitle, DialogContent, DialogActions, LinearProgress, MenuItem, Box, Chip, Accordion, AccordionSummary, AccordionDetails } from '@mui/material';
import Autocomplete from '@mui/material/Autocomplete';
import DeleteIcon from '@mui/icons-material/Delete';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import { DataGrid, GridColDef, GridToolbar } from '@mui/x-data-grid';
import * as XLSX from 'xlsx';
import api from '../api';
import { useToast } from '../components/ToastProvider';

type Warehouse = { _id: string; name: string };

type Allocation = { groupName?: string; qty?: number; source?: string; warehouseId?: string };

type OrderStatus = 'processing' | 'shipped' | 'completed' | 'canceled';
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
  allocations?: Allocation[];
  customerName?: string;
  customerPhone?: string;
  shippingAddress?: string;
  createdAtOrder?: string;
  estFulfillmentDate?: string;
  source?: string;
};

export default function Orders() {
  const todayYmd = useMemo(() => new Date().toISOString().slice(0, 10), []);
  const lastAutoSavedShipdateRef = useRef<{ orderId: string; ymd: string; at: number }>({ orderId: '', ymd: '', at: 0 });
  const [warehouses, setWarehouses] = useState<Warehouse[]>([]);
  const [ordersLoading, setOrdersLoading] = useState(false);
  const [ordersRows, setOrdersRows] = useState<OrdersRow[]>([]);
  const [ordersQ, setOrdersQ] = useState('');
  const [ordersStatusFilter, setOrdersStatusFilter] = useState<'all' | OrderStatus>('all');
  const [csvOpen, setCsvOpen] = useState(false);
  const [csvWarehouseId, setCsvWarehouseId] = useState('');
  const [csvFile, setCsvFile] = useState<File | null>(null);
  const [csvPreview, setCsvPreview] = useState<any>(null);
  const [csvErrors, setCsvErrors] = useState<any[]>([]);
  const [csvLoading, setCsvLoading] = useState(false);
  const [csvCommitting, setCsvCommitting] = useState(false);

  const [manualOpen, setManualOpen] = useState(false);
  const [manualMode, setManualMode] = useState<'create' | 'edit'>('create');
  const [manualEditRow, setManualEditRow] = useState<OrdersRow | null>(null);
  const [manualAllocations, setManualAllocations] = useState<Allocation[]>([]);
  const [manualReservedBreakdown, setManualReservedBreakdown] = useState<any[]>([]);
  const [manualWarehouseId, setManualWarehouseId] = useState('');
  const [manualStatus, setManualStatus] = useState<OrderStatus>('processing');
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

  const [manualPickerLoading, setManualPickerLoading] = useState(false);
  const [manualPickerQ, setManualPickerQ] = useState('');
  const [manualPickerWarehouses, setManualPickerWarehouses] = useState<Array<{ _id: string; name: string }>>([]);
  const [manualPickerRows, setManualPickerRows] = useState<any[]>([]);
  const [manualOrderQtyByGroup, setManualOrderQtyByGroup] = useState<Record<string, string>>({});
  const [manualShipdateTouched, setManualShipdateTouched] = useState(false);
  const [manualPickOpen, setManualPickOpen] = useState(false);
  const [manualPickSelected, setManualPickSelected] = useState<any>({ type: 'include', ids: new Set() });
  const [manualOrderGroups, setManualOrderGroups] = useState<string[]>([]);

  const [onWaterOpen, setOnWaterOpen] = useState(false);
  const [onWaterLoading, setOnWaterLoading] = useState(false);
  const [onWaterGroupName, setOnWaterGroupName] = useState('');
  const [onWaterRows, setOnWaterRows] = useState<any[]>([]);

  const [onProcessOpen, setOnProcessOpen] = useState(false);
  const [onProcessLoading, setOnProcessLoading] = useState(false);
  const [onProcessGroupName, setOnProcessGroupName] = useState('');
  const [onProcessRows, setOnProcessRows] = useState<any[]>([]);

  const [viewOrderableOpen, setViewOrderableOpen] = useState(false);
  const [viewOrderableWarehouseId, setViewOrderableWarehouseId] = useState('');
  const [viewOrderableQ, setViewOrderableQ] = useState('');
  const [viewOrderableRows, setViewOrderableRows] = useState<any[]>([]);
  const [viewOrderableWarehouses, setViewOrderableWarehouses] = useState<any[]>([]);
  const [viewOrderableLoading, setViewOrderableLoading] = useState(false);
  const [viewOrderableExporting, setViewOrderableExporting] = useState(false);

  const viewOrderableFilteredRows = useMemo(() => {
    const q = String(viewOrderableQ || '').trim().toLowerCase();
    const rows = Array.isArray(viewOrderableRows) ? viewOrderableRows : [];
    if (!q) return rows;
    return rows.filter((r: any) => {
      const hay = `${String(r?.lineItem || '')} ${String(r?.groupName || '')}`.toLowerCase();
      return hay.includes(q);
    });
  }, [viewOrderableQ, viewOrderableRows]);

  const openOnWaterDetails = async ({ warehouseId, groupName }: { warehouseId: string; groupName: string }) => {
    const wid = String(warehouseId || '').trim();
    const g = String(groupName || '').trim();
    if (!wid || !g) return;
    setOnWaterOpen(true);
    setOnWaterGroupName(g);
    setOnWaterRows([]);
    setOnWaterLoading(true);
    try {
      const { data } = await api.get('/orders/pallet-picker/on-water', { params: { warehouseId: wid, groupName: g } });
      setOnWaterRows(Array.isArray(data?.rows) ? data.rows : []);
    } catch {
      setOnWaterRows([]);
    } finally {
      setOnWaterLoading(false);
    }
  };

  const exportAllOrderablePalletsXlsx = async () => {
    const wid = String(viewOrderableWarehouseId || fixedWarehouseId || '').trim();
    if (!wid) return;
    setViewOrderableExporting(true);
    try {
      const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: wid } });
      const rows = Array.isArray(data?.rows) ? data.rows : [];
      const whs = Array.isArray(data?.warehouses) ? data.warehouses : [];
      const second = whs.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
      const secondId = second ? String(second._id) : '';
      const secondName = second ? String(second.name || '').trim() : '';
      const thisWhName =
        String((Array.isArray(warehouses) ? warehouses : []).find((w: any) => String(w?._id || '') === wid)?.name || '').trim();

      const exportRows = rows.map((r: any) => {
        const primary = Number(r?.selectedWarehouseAvailable ?? 0);
        const onWater = Number(r?.onWaterPallets ?? 0);
        const onProcess = Number(r?.onProcessPallets ?? 0);
        let secondQty = 0;
        if (secondId) {
          const per = r?.perWarehouse || {};
          secondQty = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
        }
        const maxOrder =
          (Number.isFinite(primary) ? primary : 0) +
          (Number.isFinite(onWater) ? onWater : 0) +
          (Number.isFinite(onProcess) ? onProcess : 0) +
          (Number.isFinite(secondQty) ? secondQty : 0);

        const out: any = {
          'Pallet ID': String(r?.lineItem || ''),
          'Pallet Description': String(r?.groupName || ''),
          [`THIS - ${thisWhName || 'Warehouse'}`]: Math.max(0, Math.floor(Number.isFinite(primary) ? primary : 0)),
          'On-Water': Math.max(0, Math.floor(Number.isFinite(onWater) ? onWater : 0)),
          'On-Process': Math.max(0, Math.floor(Number.isFinite(onProcess) ? onProcess : 0)),
        };
        if (secondId) out[secondName || '2nd Warehouse'] = Math.max(0, Math.floor(Number.isFinite(secondQty) ? secondQty : 0));
        out['Max Order'] = Math.max(0, Math.floor(Number.isFinite(maxOrder) ? maxOrder : 0));
        return out;
      });

      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Orderable Pallets');
      const dt = new Date();
      const pad = (n: number) => String(n).padStart(2, '0');
      const stamp = `${dt.getFullYear()}${pad(dt.getMonth() + 1)}${pad(dt.getDate())}_${pad(dt.getHours())}${pad(dt.getMinutes())}${pad(dt.getSeconds())}`;
      const filename = `orderable_pallets_${stamp}.xlsx`;
      XLSX.writeFile(wb, filename);
    } catch {
      toast.error('Failed to export XLSX');
    } finally {
      setViewOrderableExporting(false);
    }
  };

  const openManualEdit = async (row: OrdersRow) => {
    const toYmd = (v: any) => {
      const s = String(v || '');
      const slice = s.slice(0, 10);
      return /^\d{4}-\d{2}-\d{2}$/.test(slice) ? slice : '';
    };

    setManualMode('edit');
    setManualEditRow(row);
    setManualOpen(true);
    setManualWarehouseId(String(row?.warehouseId || fixedWarehouseId));
    setManualStatus((normalizeStatus(row?.status || '') as any) || 'processing');
    setManualCustomerEmail(String(row?.email || '').trim());
    setManualCustomerName(String(row?.customerName || '').trim());
    setManualCustomerPhone(String(row?.customerPhone || '').trim());
    setManualCreatedAt(toYmd(row?.createdAtOrder || row?.dateCreated || row?.createdAt));
    setManualEstFulfillment(toYmd(row?.estFulfillmentDate));
    setManualShipdateTouched(false);
    setManualShippingAddress(String(row?.shippingAddress || '').trim());
    setManualLineMode('pallet_group');
    setManualLines([{ lineItem: '', qty: '' }]);
    setManualAvailable({});
    setManualValidationErrors([]);
    setManualPickerQ('');
    setManualPickerRows([]);
    setManualPickerWarehouses([]);
    setManualPickOpen(false);
    setManualPickSelected({ type: 'include', ids: new Set() });
    setManualAllocations(Array.isArray((row as any)?.allocations) ? ((row as any).allocations as any) : []);

    const baseLines = Array.isArray(row?.lines) ? row.lines : [];
    const groups: string[] = [];
    const qtyBy: Record<string, string> = {};
    for (const l of baseLines) {
      const g = String(l?.groupName || l?.lineItem || '').trim();
      if (!g) continue;
      groups.push(g);
      const q = Math.floor(Number(l?.qty || 0));
      qtyBy[g] = q > 0 ? String(q) : '';
    }
    setManualOrderGroups(Array.from(new Set(groups)));
    setManualOrderQtyByGroup(qtyBy);

    try {
      const wid = String(row?.warehouseId || fixedWarehouseId || '').trim();
      if (wid) {
        setManualPickerLoading(true);
        const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: wid } });
        setManualPickerRows(Array.isArray(data?.rows) ? data.rows : []);
        setManualPickerWarehouses(Array.isArray(data?.warehouses) ? data.warehouses : []);
      }
    } catch {
      setManualPickerRows([]);
      setManualPickerWarehouses([]);
    } finally {
      setManualPickerLoading(false);
    }

    try {
      const rawId = String(row?.rawId || '').trim();
      if (rawId) {
        const { data } = await api.get(`/orders/unfulfilled/${rawId}`);
        setManualAllocations(Array.isArray((data as any)?.allocations) ? (data as any).allocations : []);
      }
    } catch {
      setManualAllocations(Array.isArray((row as any)?.allocations) ? ((row as any).allocations as any) : []);
    }

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

  const openOnProcessDetails = async ({ groupName }: { groupName: string }) => {
    const g = String(groupName || '').trim();
    if (!g) return;
    setOnProcessOpen(true);
    setOnProcessGroupName(g);
    setOnProcessRows([]);
    setOnProcessLoading(true);
    try {
      const { data } = await api.get('/orders/pallet-picker/on-process', { params: { groupName: g } });
      setOnProcessRows(Array.isArray(data?.rows) ? data.rows : []);
    } catch {
      setOnProcessRows([]);
    } finally {
      setOnProcessLoading(false);
    }
  };

  const fixedWarehouseId = useMemo(() => {
    const wh = Array.isArray(warehouses) ? warehouses : [];
    return String(wh[0]?._id || '');
  }, [warehouses]);

  const openViewOrderableItems = async () => {
    const wid = String(viewOrderableWarehouseId || fixedWarehouseId || '').trim();
    if (!wid) return;
    setViewOrderableOpen(true);
    setViewOrderableLoading(true);
    try {
      const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: wid, q: String(viewOrderableQ || '').trim() || undefined } });
      setViewOrderableRows(Array.isArray(data?.rows) ? data.rows : []);
      setViewOrderableWarehouses(Array.isArray(data?.warehouses) ? data.warehouses : []);
    } catch {
      setViewOrderableRows([]);
      setViewOrderableWarehouses([]);
    } finally {
      setViewOrderableLoading(false);
    }
  };

  useEffect(() => {
    if (!viewOrderableWarehouseId && fixedWarehouseId) setViewOrderableWarehouseId(String(fixedWarehouseId));
  }, [fixedWarehouseId, viewOrderableWarehouseId]);

  const manualPrevStatus = useMemo(() => {
    if (!manualEditRow) return '';
    return (normalizeStatus(manualEditRow?.status || '') as any) || '';
  }, [manualEditRow]);

  const manualIsCanceled = manualMode === 'edit' && manualPrevStatus === 'canceled';
  const manualIsCompleted = manualMode === 'edit' && manualPrevStatus === 'completed';
  const manualIsLocked = manualIsCanceled || manualIsCompleted;

  const manualWarehouseName = useMemo(() => {
    const wh = Array.isArray(warehouses) ? warehouses : [];
    const found = wh.find((w) => String(w?._id || '') === String(manualWarehouseId));
    return String(found?.name || '').trim();
  }, [warehouses, manualWarehouseId]);

  const secondWarehouse = useMemo(() => {
    const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
    const wid = String(manualWarehouseId || '').trim();
    const found = list.find((w) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
    return found ? { _id: String(found._id), name: String(found.name || '').trim() } : null;
  }, [manualPickerWarehouses, manualWarehouseId]);

  const manualPickerRowsFiltered = useMemo(() => {
    const q = String(manualPickerQ || '').trim().toLowerCase();
    const rows = Array.isArray(manualPickerRows) ? manualPickerRows : [];
    if (!q) return rows;
    return rows.filter((r: any) => {
      const gid = String(r?.lineItem || '').toLowerCase();
      const gname = String(r?.groupName || '').toLowerCase();
      return gid.includes(q) || gname.includes(q);
    });
  }, [manualPickerRows, manualPickerQ]);

  const manualPickerRowByGroup = useMemo(() => {
    const m = new Map<string, any>();
    for (const r of (Array.isArray(manualPickerRows) ? manualPickerRows : [])) {
      const groupKey = String(r?.groupName || '').trim();
      const lineKey = String(r?.lineItem || '').trim();
      if (groupKey && !m.has(groupKey)) m.set(groupKey, r);
      if (lineKey && !m.has(lineKey)) m.set(lineKey, r);
    }
    return m;
  }, [manualPickerRows]);

  const manualOrderRows = useMemo(() => {
    return (manualOrderGroups || [])
      .map((g) => {
        const r: any = manualPickerRowByGroup.get(String(g)) || {};
        return { id: String(g), ...r, groupName: String(g) };
      })
      .filter((r) => String(r.groupName || '').trim());
  }, [manualOrderGroups, manualPickerRowByGroup]);

  const manualAllocationRows = useMemo(() => {
    const allocs = Array.isArray(manualAllocations) ? manualAllocations : [];
    const byGroup = new Map<string, { groupName: string; primary: number; onWater: number; onProcess: number; second: number }>();
    for (const a of allocs) {
      const g = String(a?.groupName || '').trim();
      const src = String(a?.source || '').trim().toLowerCase();
      const qty = Math.floor(Number(a?.qty || 0));
      if (!g || !Number.isFinite(qty) || qty <= 0) continue;
      if (!byGroup.has(g)) byGroup.set(g, { groupName: g, primary: 0, onWater: 0, onProcess: 0, second: 0 });
      const rec = byGroup.get(g)!;
      if (src === 'primary') rec.primary += qty;
      else if (src === 'on_water') rec.onWater += qty;
      else if (src === 'on_process') rec.onProcess += qty;
      else if (src === 'second') rec.second += qty;
    }
    return Array.from(byGroup.values())
      .map((r) => ({ id: r.groupName, ...r }))
      .sort((a, b) => String(a.groupName).localeCompare(String(b.groupName)));
  }, [manualAllocations]);

  const manualReservedRows = useMemo(() => {
    const rows = Array.isArray(manualReservedBreakdown) ? manualReservedBreakdown : [];
    if (rows.length) {
      return rows
        .map((r: any) => ({
          id: String(r?.id || r?.groupName || ''),
          groupName: String(r?.groupName || ''),
          primary: Math.floor(Number(r?.primary || 0)),
          onWater: Math.floor(Number(r?.onWater || 0)),
          second: Math.floor(Number(r?.second || 0)),
          onProcess: Math.floor(Number(r?.onProcess || 0)),
        }))
        .filter((r: any) => String(r.groupName || '').trim())
        .sort((a: any, b: any) => String(a.groupName).localeCompare(String(b.groupName)));
    }
    return manualAllocationRows;
  }, [manualReservedBreakdown, manualAllocationRows]);

  const manualAllocationWarehouseLabels = useMemo(() => {
    const all = [...(Array.isArray(warehouses) ? warehouses : []), ...(Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [])] as any[];
    const nameById = new Map<string, string>();
    for (const w of all) {
      const id = String(w?._id || '').trim();
      const nm = String(w?.name || '').trim();
      if (id && nm && !nameById.has(id)) nameById.set(id, nm);
    }

    const primaryWarehouseId = String(manualWarehouseId || '').trim();
    const primaryName =
      nameById.get(primaryWarehouseId) ||
      String((manualEditRow as any)?.warehouseName || '').trim() ||
      '';
    const allocs = Array.isArray(manualAllocations) ? manualAllocations : [];
    const reserved = Array.isArray(manualReservedBreakdown) ? manualReservedBreakdown : [];

    const secondAlloc = allocs.find((a) => String(a?.source || '').toLowerCase() === 'second') as any;
    const secondWarehouseId = String(secondAlloc?.warehouseId?._id || secondAlloc?.warehouseId || '').trim();
    const secondNameFromAlloc = String(secondAlloc?.warehouseId?.name || '').trim();

    const fallbackSecondName =
      all
        .map((w) => ({ id: String(w?._id || '').trim(), name: String(w?.name || '').trim() }))
        .find((w) => w.id && w.id !== primaryWarehouseId && w.name)?.name ||
      '';

    const secondName =
      secondNameFromAlloc ||
      (secondWarehouseId ? (nameById.get(secondWarehouseId) || '') : '') ||
      fallbackSecondName;

    return {
      primaryLabel: primaryName || 'Primary',
      secondLabel: secondName || '2nd',
    };
  }, [manualAllocations, manualEditRow, manualPickerWarehouses, manualReservedBreakdown, manualWarehouseId, warehouses]);

  const isValidEmail = (email: string) => {
    const s = String(email || '').trim();
    if (!s) return false;
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
  };

  const sanitizeIntText = (v: any) => {
    const s = String(v ?? '');
    // allow empty while typing
    if (!s) return '';
    // strip everything except digits
    const only = s.replace(/[^0-9]/g, '');
    return only;
  };

  const normalizeIntText = (v: any) => {
    const s = sanitizeIntText(v);
    if (!s) return '';
    const n = Math.floor(Number(s));
    if (!Number.isFinite(n) || n <= 0) return '';
    return String(n);
  };

  const normalizeManualLine = (v: any) => String(v || '').trim();

  const getActiveManualOptions = () => (manualLineMode === 'pallet_group' ? manualPalletGroupOptions : manualLineItemOptions);

  const suggestShipdateForSelection = useCallback(() => {
    const today = todayYmd;
    const rowsByGroup = new Map((manualPickerRows || []).map((r: any) => [String(r.groupName || ''), r]));
    const toYmd = (s: any) => {
      const v = String(s || '');
      const slice = v.slice(0, 10);
      return /^\d{4}-\d{2}-\d{2}$/.test(slice) ? slice : '';
    };

    const addMonthsYmd = (ymd: string, months: number) => {
      const s = String(ymd || '').trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return '';
      const dt = new Date(`${s}T00:00:00`);
      if (Number.isNaN(dt.getTime())) return '';
      const next = new Date(dt);
      next.setMonth(next.getMonth() + Math.floor(Number(months || 0)));
      const pad = (n: number) => String(n).padStart(2, '0');
      return `${next.getFullYear()}-${pad(next.getMonth() + 1)}-${pad(next.getDate())}`;
    };

    const pickedGroups = Object.entries(manualOrderQtyByGroup || {})
      .filter(([, qty]) => Number(qty || 0) > 0)
      .map(([g]) => String(g || '').trim())
      .filter((v) => v);

    let best = '';
    for (const g of pickedGroups) {
      const r: any = rowsByGroup.get(g) || {};

      const need = Math.floor(Number((manualOrderQtyByGroup || {})?.[g] || 0));
      if (!Number.isFinite(need) || need <= 0) continue;

      let remaining = need;
      const useDates: string[] = [];

      const primaryAvail = Math.max(0, Math.floor(Number(r?.selectedWarehouseAvailable ?? 0)));
      const takePrimary = Math.min(primaryAvail, remaining);
      if (takePrimary > 0) {
        useDates.push(today);
        remaining -= takePrimary;
      }

      const onWaterAvail = Math.max(0, Math.floor(Number(r?.onWaterPallets ?? 0)));
      const takeOnWater = Math.min(onWaterAvail, remaining);
      if (takeOnWater > 0) {
        const onWaterReady = toYmd(r?.onWaterEdd);
        if (onWaterReady) useDates.push(onWaterReady);
        remaining -= takeOnWater;
      }

      if (remaining > 0 && secondWarehouse?._id) {
        const per = r?.perWarehouse || {};
        const wid = String(secondWarehouse._id);
        const secondAvail = Math.max(0, Math.floor(Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0)));
        const takeSecond = Math.min(secondAvail, remaining);
        if (takeSecond > 0) {
          const cutoff = addMonthsYmd(today, 3);
          if (cutoff) useDates.push(cutoff);
          remaining -= takeSecond;
        }
      }

      if (remaining > 0) {
        const onProcessAvail = Math.max(0, Math.floor(Number(r?.onProcessPallets ?? 0)));
        const takeOnProcess = Math.min(onProcessAvail, remaining);
        if (takeOnProcess > 0) {
          const base = toYmd(r?.onProcessEdd);
          const onProcessReady = base ? addMonthsYmd(base, 3) : '';
          if (onProcessReady) useDates.push(onProcessReady);
          remaining -= takeOnProcess;
        }
      }

      for (const d of useDates) {
        if (!d) continue;
        if (!best || d > best) best = d;
      }
    }
    return best || today;
  }, [manualPickerRows, manualOrderQtyByGroup, secondWarehouse, todayYmd]);

  const suggestShipdateForProcessingAllocations = useCallback((opts?: { rows?: any[]; allocations?: any[] }) => {
    const today = todayYmd;
    const pickerRows = Array.isArray(opts?.rows) ? opts!.rows : (manualPickerRows || []);
    const allocs = Array.isArray(opts?.allocations) ? opts!.allocations : (Array.isArray(manualAllocations) ? manualAllocations : []);

    const rowsByGroup = new Map((pickerRows || []).map((r: any) => [String(r.groupName || ''), r]));
    const toYmd = (s: any) => {
      const v = String(s || '');
      const slice = v.slice(0, 10);
      return /^\d{4}-\d{2}-\d{2}$/.test(slice) ? slice : '';
    };
    const addMonthsYmd = (ymd: string, months: number) => {
      const s = String(ymd || '').trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return '';
      const dt = new Date(`${s}T00:00:00`);
      if (Number.isNaN(dt.getTime())) return '';
      const next = new Date(dt);
      next.setMonth(next.getMonth() + Math.floor(Number(months || 0)));
      const pad = (n: number) => String(n).padStart(2, '0');
      return `${next.getFullYear()}-${pad(next.getMonth() + 1)}-${pad(next.getDate())}`;
    };

    let best = '';
    let hasSecond = false;

    for (const a of allocs) {
      const g = String(a?.groupName || '').trim();
      const src = String(a?.source || '').trim().toLowerCase();
      const qty = Math.floor(Number(a?.qty || 0));
      if (!g || !Number.isFinite(qty) || qty <= 0) continue;

      if (src === 'second') {
        hasSecond = true;
        continue;
      }

      const r: any = rowsByGroup.get(g) || {};
      if (src === 'on_water') {
        const edd = toYmd(r?.onWaterEdd);
        if (edd && (!best || edd > best)) best = edd;
        continue;
      }
      if (src === 'on_process') {
        const edd = toYmd(r?.onProcessEdd);
        const ready = edd ? addMonthsYmd(edd, 3) : '';
        if (ready && (!best || ready > best)) best = ready;
        continue;
      }
    }

    let out = best || today;
    if (hasSecond) {
      const cutoff = addMonthsYmd(today, 3);
      if (cutoff) out = cutoff;
    }
    return out || today;
  }, [manualAllocations, manualPickerRows, todayYmd]);

  const refreshManualAvailable = useCallback(async (warehouseId: string) => {
    const wid = String(warehouseId || '').trim();
    if (!wid) {
      setManualAvailable({});
      return;
    }
    try {
      const { data } = await api.get('/pallet-inventory/groups');
      const map: Record<string, number> = {};
      (Array.isArray(data) ? data : []).forEach((g: any) => {
        const groupName = String(g.groupName || '').trim();
        if (!groupName) return;
        const per = Array.isArray(g.perWarehouse) ? g.perWarehouse : [];
        const rec = per.find((p: any) => String(p.warehouseId) === String(wid));
        map[groupName] = Number(rec?.pallets || 0);
      });
      setManualAvailable(map);
    } catch {
      setManualAvailable({});
    }
  }, []);

  const refreshManualPicker = useCallback(
    async (warehouseId: string) => {
      const wid = String(warehouseId || '').trim();
      if (!wid) {
        setManualPickerRows([]);
        setManualPickerWarehouses([]);
        return { rows: [] as any[], warehouses: [] as any[] };
      }
      try {
        const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: wid } });
        const rows = Array.isArray(data?.rows) ? data.rows : [];
        const warehouses = Array.isArray(data?.warehouses) ? data.warehouses : [];
        setManualPickerRows(rows);
        setManualPickerWarehouses(warehouses);
        return { rows, warehouses };
      } catch {
        setManualPickerRows([]);
        setManualPickerWarehouses([]);
        return { rows: [] as any[], warehouses: [] as any[] };
      }
    },
    []
  );

  const refreshManualAllocations = useCallback(async (rawId: string) => {
    const id = String(rawId || '').trim();
    if (!id) {
      setManualAllocations([]);
      setManualReservedBreakdown([]);
      return [] as any[];
    }
    try {
      const { data } = await api.get(`/orders/unfulfilled/${id}`, { params: { _ts: Date.now() } });
      const nextAllocs = Array.isArray((data as any)?.allocations) ? (data as any).allocations : [];
      const nextReserved = Array.isArray((data as any)?.reservedBreakdown) ? (data as any).reservedBreakdown : [];
      setManualAllocations(nextAllocs);
      setManualReservedBreakdown(nextReserved);
      setManualEditRow((prev) => {
        if (!prev) return prev;
        return { ...(prev as any), allocations: nextAllocs } as any;
      });
      return nextAllocs;
    } catch {
      // keep existing allocations if fetch fails
      return Array.isArray(manualAllocations) ? manualAllocations : [];
    }
  }, [manualAllocations]);

  const getRowSupplySource = (row: any) => {
    const r = row || {};
    return String(r?.queueSource || r?.source || r?.queue || '').trim();
  };

  useEffect(() => {
    if (!manualOpen) return;
    if (manualMode === 'edit') return;
    if (manualShipdateTouched) return;
    const next = suggestShipdateForSelection();
    if (next && next !== manualEstFulfillment) {
      setManualEstFulfillment(next);
    }
  }, [manualOpen, manualMode, manualShipdateTouched, manualOrderQtyByGroup, suggestShipdateForSelection, manualEstFulfillment]);

  useEffect(() => {
    const loadAvail = async () => {
      if (!manualWarehouseId) { setManualAvailable({}); return; }
      await refreshManualAvailable(manualWarehouseId);
    };
    loadAvail();
  }, [manualWarehouseId, refreshManualAvailable]);

  useEffect(() => {
    const isProcessing = manualOpen && manualMode === 'edit' && manualPrevStatus === 'processing' && !manualIsLocked;
    if (!isProcessing) return;
    const wid = String(manualWarehouseId || '').trim();
    const rawId = String((manualEditRow as any)?.rawId || '').trim();
    if (!wid || !rawId) return;

    let stopped = false;
    const tick = async () => {
      if (stopped) return;
      const picker = await refreshManualPicker(wid);
      await refreshManualAvailable(wid);
      const allocs = await refreshManualAllocations(rawId);
      if (!manualShipdateTouched) {
        const next = suggestShipdateForProcessingAllocations({ rows: picker?.rows, allocations: allocs });
        if (next && next !== manualEstFulfillment) {
          setManualEstFulfillment(next);
          setManualEditRow((prev) => {
            if (!prev) return prev;
            const prevRawId = String((prev as any)?.rawId || '').trim();
            if (prevRawId && prevRawId !== rawId) return prev;
            return { ...(prev as any), estFulfillmentDate: next } as any;
          });
          setOrdersRows((prev) => {
            const rows = Array.isArray(prev) ? prev : [];
            return rows.map((r) => {
              const rRawId = String((r as any)?.rawId || '').trim();
              if (rRawId && rRawId === rawId) return { ...(r as any), estFulfillmentDate: next } as any;
              return r;
            });
          });

          // Auto-persist the suggested shipdate so the user doesn't need to click Save.
          // Only for PROCESSING orders and only if the user has not manually touched the shipdate.
          try {
            const last = lastAutoSavedShipdateRef.current || { orderId: '', ymd: '', at: 0 };
            const now = Date.now();
            const tooSoon = now - Number(last.at || 0) < 5000;
            const sameOrder = String(last.orderId || '') === String(rawId || '');
            const sameYmd = String(last.ymd || '') === String(next || '');
            if (!(sameOrder && sameYmd) && !tooSoon) {
              lastAutoSavedShipdateRef.current = { orderId: String(rawId || ''), ymd: String(next || ''), at: now };
              await api.put(`/orders/unfulfilled/${rawId}`, { estFulfillmentDate: next });
            }
          } catch {
            // ignore
          }
        }
      }
    };

    tick();
    const id = window.setInterval(tick, 15000);
    const onFocus = () => { tick(); };
    const onVisibility = () => {
      if (document.visibilityState === 'visible') tick();
    };
    window.addEventListener('focus', onFocus);
    document.addEventListener('visibilitychange', onVisibility);
    return () => {
      stopped = true;
      window.clearInterval(id);
      window.removeEventListener('focus', onFocus);
      document.removeEventListener('visibilitychange', onVisibility);
    };
  }, [manualOpen, manualMode, manualPrevStatus, manualIsLocked, manualWarehouseId, manualEditRow, manualShipdateTouched, manualEstFulfillment, refreshManualPicker, refreshManualAvailable, refreshManualAllocations, suggestShipdateForProcessingAllocations]);
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

  function normalizeStatus(v: any) {
    const s = String(v || '').trim().toLowerCase();
    if (!s) return '';
    if (s === 'created') return 'processing';
    if (s === 'create') return 'processing';
    if (s === 'backorder') return 'processing';
    if (s === 'fulfilled') return 'completed';
    if (s === 'cancelled') return 'canceled';
    if (s === 'cancel') return 'canceled';
    return s;
  }

  const openDetailedEdit = async (row: OrdersRow) => {
    if (!row) return;
    if (!String(row.id || '').startsWith('manual:')) {
      toast.error('Admin disabled editing for imported orders. Please contact the admin.');
      return;
    }
    await openManualEdit(row);
  };

  useEffect(() => {
    if (!fixedWarehouseId) return;
    if (csvOpen) setCsvWarehouseId(fixedWarehouseId);
    if (manualOpen) setManualWarehouseId(fixedWarehouseId);
  }, [fixedWarehouseId, csvOpen, manualOpen]);

  async function loadOrders(warehousesSnapshot?: Warehouse[]) {
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
        status: normalizeStatus(o?.status || 'processing') || 'processing',
        warehouseId: normalizeId(o?.warehouseId),
        warehouseName: whNameById.get(normalizeId(o?.warehouseId)) || '',
        createdAt: normalizeDateValue(o?.createdAt),
        dateCreated: normalizeDateValue(o?.createdAtOrder || o?.createdAt),
        refDate: String(o?.createdAtOrder || o?.createdAt || ''),
        lineCount: Array.isArray(o?.lines) ? o.lines.length : 0,
        totalQty: Array.isArray(o?.lines) ? o.lines.reduce((s: number, l: any) => s + (Number(l?.qty) || 0), 0) : 0,
        email: String(o?.customerEmail || ''),
        lines: Array.isArray(o?.lines) ? o.lines : [],
        allocations: Array.isArray(o?.allocations) ? o.allocations : [],
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
        status: normalizeStatus(o?.status || 'completed') || 'completed',
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
  }

  useEffect(() => {
    const handler = async () => {
      try {
        await loadOrders();
      } catch {
        // ignore
      }

      const isProcessing = manualOpen && manualMode === 'edit' && manualPrevStatus === 'processing' && !manualIsLocked;
      const rawId = String((manualEditRow as any)?.rawId || '').trim();
      if (isProcessing && rawId) {
        try {
          await refreshManualAllocations(rawId);
        } catch {
          // ignore
        }
      }
    };

    window.addEventListener('shipments-changed', handler as any);
    return () => {
      window.removeEventListener('shipments-changed', handler as any);
    };
  }, [manualOpen, manualMode, manualPrevStatus, manualIsLocked, manualEditRow, refreshManualAllocations]);

  useEffect(() => {
    let stopped = false;
    const tick = async () => {
      if (stopped) return;
      if (ordersLoading) return;
      try {
        await loadOrders();
      } catch {
        // ignore
      }
    };

    const id = window.setInterval(tick, 20000);
    const onFocus = () => { tick(); };
    const onVisibility = () => {
      if (document.visibilityState === 'visible') tick();
    };
    window.addEventListener('focus', onFocus);
    document.addEventListener('visibilitychange', onVisibility);
    return () => {
      stopped = true;
      window.clearInterval(id);
      window.removeEventListener('focus', onFocus);
      document.removeEventListener('visibilitychange', onVisibility);
    };
  }, [ordersLoading]);

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
    {
      field: 'status',
      headerName: 'Status',
      width: 140,
      renderCell: (p: any) => {
        const s = normalizeStatus((p?.row as any)?.status || '');
        const label = s ? String(s).toUpperCase() : '-';
        const color: any = s === 'completed' ? 'success' : s === 'canceled' ? 'error' : s === 'processing' ? 'warning' : 'default';
        const variant: any = s === 'processing' ? 'outlined' : 'filled';
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
    { field: 'customerName', headerName: 'Customer Name', flex: 1, minWidth: 190, renderCell: (p: any) => String((p?.row as any)?.customerName || '-') },
    { field: 'lineCount', headerName: 'Lines', width: 90, type: 'number' },
    { field: 'totalQty', headerName: 'Qty', width: 90, type: 'number' },
    {
      field: 'actions',
      headerName: 'Actions',
      width: 120,
      sortable: false,
      filterable: false,
      renderCell: (params: any) => (
        <Button size="small" variant="outlined" onClick={() => openDetailedEdit(params.row as OrdersRow)}>View</Button>
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
        r.status,
        r.email,
        (r as any).customerName || '',
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
    setManualMode('create');
    setManualEditRow(null);
    setManualOpen(true);
    setManualWarehouseId(fixedWarehouseId);
    setManualStatus('processing');
    setManualCustomerEmail('');
    setManualCustomerName('');
    setManualCustomerPhone('');
    setManualCreatedAt(new Date().toISOString().slice(0, 10));
    setManualEstFulfillment('');
    setManualShipdateTouched(false);
    setManualShippingAddress('');
    setManualLineMode('pallet_group');
    setManualLines([{ lineItem: '', qty: '' }]);
    setManualAvailable({});
    setManualValidationErrors([]);
    setManualPickerQ('');
    setManualPickerRows([]);
    setManualPickerWarehouses([]);
    setManualOrderQtyByGroup({});
    setManualPickOpen(false);
    setManualPickSelected({ type: 'include', ids: new Set() });
    setManualOrderGroups([]);
    try {
      if (fixedWarehouseId) {
        setManualPickerLoading(true);
        const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: fixedWarehouseId } });
        setManualPickerRows(Array.isArray(data?.rows) ? data.rows : []);
        setManualPickerWarehouses(Array.isArray(data?.warehouses) ? data.warehouses : []);
      }
    } catch {
      setManualPickerRows([]);
      setManualPickerWarehouses([]);
    } finally {
      setManualPickerLoading(false);
    }
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

    const allowed = new Set((manualOrderGroups || []).map((g) => String(g).trim().toLowerCase()).filter((v) => v));
    const parsed = Object.entries(manualOrderQtyByGroup || {})
      .map(([groupName, qtyStr]) => ({
        groupName: String(groupName || '').trim(),
        qty: Math.floor(Number(qtyStr || 0)),
      }))
      .filter((l) => l.groupName && allowed.has(l.groupName.toLowerCase()) && Number.isFinite(l.qty) && l.qty > 0);
    if (!parsed.length) {
      errs.push('At least 1 Pallet ID is required');
    }
    const seen = new Set<string>();
    for (const l of parsed) {
      const k = l.groupName.toLowerCase();
      if (seen.has(k)) {
        errs.push(`Duplicate Pallet Description: ${l.groupName}`);
        break;
      }
      seen.add(k);
    }

    if (manualIsCanceled) {
      errs.push('Canceled orders cannot be edited');
    }

    if (manualIsCompleted) {
      errs.push('Completed orders cannot be edited');
    }

    if (errs.length) {
      setManualValidationErrors(errs);
      toast.error(errs[0]);
      return;
    }

    // Client-side availability check (server also enforces)
    // For backorder, allow negative stock, so skip this check.
    try {
      if (manualMode === 'edit' && manualEditRow) {
        if (manualPrevStatus !== 'processing' && manualStatus === 'processing') {
          throw new Error('Changing status back to PROCESSING is not allowed because inventory has already been deducted.');
        }

        if (manualPrevStatus === 'processing' && manualStatus === 'shipped') {
          const ok = window.confirm(
            'Changing the status to SHIPPED cannot be changed back to PROCESSING since the inventory will be deducted. Continue?'
          );
          if (!ok) return;
        }

        if (manualPrevStatus === 'shipped' && manualStatus === 'completed') {
          const ok = window.confirm(
            'Changing the status to COMPLETED will lock this order and cannot be undone. Are you sure you want to proceed?'
          );
          if (!ok) return;
        }

        if (manualStatus === 'canceled' && manualPrevStatus !== 'canceled') {
          const ok = window.confirm('Are you sure you want to cancel this order? Canceling order cannot be undone.');
          if (!ok) return;
        }

        // Status changes are handled by a separate endpoint.
        await api.put(`/orders/unfulfilled/${manualEditRow.rawId}/status`, {
          status: manualStatus,
        });
        if (manualStatus === 'processing') {
          await api.put(`/orders/unfulfilled/${manualEditRow.rawId}`, {
            customerEmail: manualCustomerEmail.trim(),
            customerName: manualCustomerName.trim(),
            customerPhone: manualCustomerPhone.trim(),
            estFulfillmentDate: manualEstFulfillment || undefined,
            shippingAddress: manualShippingAddress.trim(),
            lines: parsed.map((l) => ({ search: l.groupName, qty: l.qty })),
          });
        }
        toast.success('Order updated');
        setManualOpen(false);
        await loadOrders();
        return;
      }

      await api.post('/orders/unfulfilled', {
        warehouseId: manualWarehouseId,
        status: 'processing',
        customerEmail: manualCustomerEmail.trim(),
        customerName: manualCustomerName.trim(),
        customerPhone: manualCustomerPhone.trim(),
        createdAtOrder: manualCreatedAt || undefined,
        estFulfillmentDate: manualEstFulfillment || undefined,
        shippingAddress: manualShippingAddress.trim(),
        lines: parsed.map((l) => ({ search: l.groupName, qty: l.qty })),
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
          <Button variant="outlined" onClick={openManual}>Add Order</Button>
          <Button
            variant="outlined"
            onClick={async () => {
              setViewOrderableQ('');
              await openViewOrderableItems();
            }}
            disabled={!fixedWarehouseId}
          >
            VIEW ORDERABLE PALLETS
          </Button>
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
              <MenuItem value="processing">PROCESSING</MenuItem>
              <MenuItem value="shipped">SHIPPED</MenuItem>
              <MenuItem value="completed">COMPLETED</MenuItem>
              <MenuItem value="canceled">CANCELED</MenuItem>
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
          onRowDoubleClick={(p:any)=> openDetailedEdit(p.row as OrdersRow)}
        />
      </Paper>

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

      <Dialog open={manualOpen} onClose={()=>setManualOpen(false)} fullWidth maxWidth="xl">
        <DialogTitle>{manualMode === 'edit' ? `Edit Order${manualEditRow?.orderNumber ? ` - ${manualEditRow.orderNumber}` : ''}` : 'Add Order'}</DialogTitle>
        <DialogContent>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb: 2, mt: 1 }}>
            <TextField select disabled label="Warehouse" size="small" sx={{ minWidth: 220 }} value={manualWarehouseId} onChange={(e)=>setManualWarehouseId(e.target.value)} error={!manualWarehouseId} helperText={!manualWarehouseId ? 'Required' : ''}>
              {warehouses.map((w)=> (
                <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>
              ))}
            </TextField>
            {manualMode === 'edit' ? (
              <TextField
                select
                label="Status"
                size="small"
                sx={{ minWidth: 180 }}
                value={manualStatus}
                onChange={(e)=> setManualStatus(e.target.value as any)}
                disabled={manualIsLocked}
              >
                <MenuItem value="processing" disabled={manualPrevStatus === 'shipped'}>PROCESSING</MenuItem>
                <MenuItem value="shipped">SHIPPED</MenuItem>
                <MenuItem value="completed">COMPLETED</MenuItem>
                <MenuItem value="canceled" disabled={manualPrevStatus === 'shipped'}>CANCELED</MenuItem>
              </TextField>
            ) : null}
            <TextField disabled={manualIsLocked} label="Customer Email" size="small" value={manualCustomerEmail} onChange={(e)=>setManualCustomerEmail(e.target.value)} sx={{ minWidth: 200 }} error={!isValidEmail(manualCustomerEmail)} helperText={!isValidEmail(manualCustomerEmail) ? 'Required (valid email)' : ''} />
            <TextField disabled={manualIsLocked} label="Customer Name" size="small" value={manualCustomerName} onChange={(e)=>setManualCustomerName(e.target.value)} sx={{ minWidth: 200 }} error={!String(manualCustomerName||'').trim()} helperText={!String(manualCustomerName||'').trim() ? 'Required' : ''} />
          </Stack>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb: 2 }}>
            <TextField disabled={manualIsLocked} label="Phone Number" size="small" value={manualCustomerPhone} onChange={(e)=>setManualCustomerPhone(e.target.value)} sx={{ minWidth: 220 }} error={!String(manualCustomerPhone||'').trim()} helperText={!String(manualCustomerPhone||'').trim() ? 'Required' : ''} />
          </Stack>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} sx={{ mb: 2 }}>
            <TextField type="date" label="Created Order Date" InputLabelProps={{ shrink: true }} size="small" value={manualCreatedAt} onChange={(e)=>setManualCreatedAt(e.target.value)} sx={{ minWidth: 220 }} inputProps={{ max: todayYmd }} error={!manualCreatedAt || manualCreatedAt > todayYmd} helperText={!manualCreatedAt ? 'Required' : (manualCreatedAt > todayYmd ? 'Cannot be advance date' : '')} disabled={manualMode === 'edit'} />
            <TextField disabled={manualIsLocked} type="date" label="Estimated Shipdate for Customer" InputLabelProps={{ shrink: true }} size="small" value={manualEstFulfillment} onChange={(e)=>{ setManualShipdateTouched(true); setManualEstFulfillment(e.target.value); }} sx={{ minWidth: 220 }} inputProps={{ min: todayYmd }} error={!manualEstFulfillment || manualEstFulfillment < todayYmd} helperText={!manualEstFulfillment ? 'Required' : (manualEstFulfillment < todayYmd ? 'Must be today or later' : '')} />
            <Button
              variant="outlined"
              disabled={!manualOpen || !manualWarehouseId || manualIsLocked}
              onClick={() => {
                const isProcessing = manualMode === 'edit' && manualPrevStatus === 'processing';
                const next = isProcessing ? suggestShipdateForProcessingAllocations() : suggestShipdateForSelection();
                setManualShipdateTouched(false);
                setManualEstFulfillment(next);
              }}
              sx={{ minWidth: 170 }}
            >
              Reset to suggested
            </Button>
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
            disabled={manualIsLocked}
          />

          {manualMode === 'edit' ? (
            <Box sx={{ mb: 2 }}>
              <Accordion
                defaultExpanded={false}
                sx={{ border: '2px solid', borderColor: 'primary.main', borderRadius: 1, overflow: 'hidden' }}
              >
                <AccordionSummary
                  expandIcon={<ExpandMoreIcon />}
                  sx={{ bgcolor: 'rgba(25, 118, 210, 0.06)' }}
                >
                  <Typography variant="subtitle1">Reserved Stock for This Order (Hierarchy)</Typography>
                </AccordionSummary>
                <AccordionDetails>
                  {manualReservedRows.length ? (
                    <div style={{ height: 220, width: '100%' }}>
                      <DataGrid
                        rows={manualReservedRows}
                        columns={([
                          { field: 'groupName', headerName: 'Pallet Description', flex: 1, minWidth: 240 },
                          { field: 'primary', headerName: manualAllocationWarehouseLabels.primaryLabel, width: 150, type: 'number', align: 'right', headerAlign: 'right' },
                          { field: 'onWater', headerName: 'On-Water', width: 120, type: 'number', align: 'right', headerAlign: 'right' },
                          { field: 'second', headerName: manualAllocationWarehouseLabels.secondLabel, width: 150, type: 'number', align: 'right', headerAlign: 'right' },
                          { field: 'onProcess', headerName: 'On-Process', width: 120, type: 'number', align: 'right', headerAlign: 'right' },
                        ]) as GridColDef[]}
                        disableRowSelectionOnClick
                        density="compact"
                        hideFooter
                      />
                    </div>
                  ) : (
                    <Typography variant="body2" color="text.secondary">
                      No reservation breakdown available yet.
                    </Typography>
                  )}
                </AccordionDetails>
              </Accordion>
            </Box>
          ) : null}

          <Typography variant="subtitle1" sx={{ mb: 1 }}>Pallets to Order</Typography>

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

          <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} sx={{ mb: 1 }}>
            <Button
              variant="outlined"
              disabled={!manualWarehouseId || manualIsCanceled}
              onClick={() => {
                setManualPickSelected({ type: 'include', ids: new Set((manualOrderGroups || []).map((x)=> String(x))) });
                setManualPickOpen(true);
              }}
            >
              Add to list
            </Button>
            <Box sx={{ flex: 1 }} />
          </Stack>

          {manualOrderGroups.length === 0 ? (
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              No pallets added yet. Click "Add to list" to select pallet(s).
            </Typography>
          ) : null}

          <div style={{ height: 320, width: '100%' }}>
            <DataGrid
              rows={manualOrderRows}
              columns={(() => {
                const cols: any[] = [
                  { field: 'lineItem', headerName: 'Pallet ID', width: 140, renderCell: (p: any) => String(p?.row?.lineItem || '-') },
                  { field: 'groupName', headerName: 'Pallet Description', flex: 1, minWidth: 220, renderCell: (p: any) => String(p?.row?.groupName || '-') },
                  { field: 'selectedWarehouseAvailable', headerName: `THIS - ${manualWarehouseName || 'Warehouse'}`, width: 170, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => String(p?.row?.selectedWarehouseAvailable ?? 0) },
                  {
                    field: 'onWaterPallets',
                    headerName: 'On-Water',
                    width: 110,
                    type: 'number',
                    align: 'right',
                    headerAlign: 'right',
                    renderCell: (p: any) => {
                      const qty = Number(p?.row?.onWaterPallets ?? 0);
                      const groupName = String(p?.row?.groupName || '').trim();
                      const wid = String(manualWarehouseId || '').trim();
                      if (!qty) return '0';
                      return (
                        <Button
                          variant="text"
                          size="small"
                          onClick={() => openOnWaterDetails({ warehouseId: wid, groupName })}
                          sx={{ minWidth: 0, p: 0, textDecoration: 'underline', fontSize: 16, fontWeight: 700 }}
                        >
                          {qty}
                        </Button>
                      );
                    },
                  },
                ];

                if (secondWarehouse?._id) {
                  cols.push({
                    field: 'secondWarehouseAvailable',
                    headerName: `${secondWarehouse.name || 'Warehouse'}`,
                    width: 170,
                    type: 'number',
                    align: 'right',
                    headerAlign: 'right',
                    sortable: false,
                    filterable: false,
                    renderCell: (p: any) => {
                      const per = (p?.row as any)?.perWarehouse || {};
                      const wid = String(secondWarehouse._id);
                      const v = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
                      return String(Number.isFinite(v) ? v : 0);
                    },
                  });
                }

                cols.push({
                  field: 'onProcessPallets',
                  headerName: 'On-Process',
                  width: 120,
                  type: 'number',
                  align: 'right',
                  headerAlign: 'right',
                  renderCell: (p: any) => {
                    const qty = Number(p?.row?.onProcessPallets ?? 0);
                    const groupName = String(p?.row?.groupName || '').trim();
                    if (!qty) return '0';
                    return (
                      <Button
                        variant="text"
                        size="small"
                        onClick={() => openOnProcessDetails({ groupName })}
                        sx={{ minWidth: 0, p: 0, textDecoration: 'underline', fontSize: 16, fontWeight: 700 }}
                      >
                        {qty}
                      </Button>
                    );
                  },
                });

                cols.push({
                  field: 'maxQtyOrder',
                  headerName: 'Max Qty Order',
                  width: 140,
                  type: 'number',
                  align: 'right',
                  headerAlign: 'right',
                  sortable: false,
                  filterable: false,
                  valueGetter: (...args: any[]) => {
                    // MUI DataGrid has had multiple valueGetter signatures across versions:
                    // - (params) => any
                    // - (value, row) => any
                    const maybeParams = args?.[0];
                    const maybeRow = args?.[1];
                    const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});

                    // Max Qty Order is the total orderable quantity across tiers.
                    // It should reflect the displayed tier quantities (Primary + On-Water + 2nd Warehouse + On-Process).
                    const primary = Number(row?.selectedWarehouseAvailable ?? 0);
                    const onWater = Number(row?.onWaterPallets ?? 0);
                    const onProcess = Number(row?.onProcessPallets ?? 0);
                    let second = 0;
                    if (secondWarehouse?._id) {
                      const per = row?.perWarehouse || {};
                      const wid = String(secondWarehouse._id);
                      second = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
                    }
                    const total =
                      (Number.isFinite(primary) ? primary : 0) +
                      (Number.isFinite(onWater) ? onWater : 0) +
                      (Number.isFinite(second) ? second : 0) +
                      (Number.isFinite(onProcess) ? onProcess : 0);
                    return Math.max(0, Math.floor(total));
                  },
                  renderCell: (p: any) => {
                    const v = Number((p && typeof p === 'object' && 'value' in p) ? (p as any).value : 0);
                    return String(Number.isFinite(v) ? v : 0);
                  },
                });

                cols.push({
                  field: 'orderQty',
                  headerName: 'Order Qty',
                  width: 120,
                  sortable: false,
                  filterable: false,
                  renderCell: (p: any) => {
                    const groupName = String(p?.row?.groupName || '');
                    const val = manualOrderQtyByGroup[groupName] ?? '';
                    return (
                      <TextField
                        type="number"
                        size="small"
                        value={val}
                        disabled={manualIsLocked}
                        onChange={(e)=>{
                          const raw = sanitizeIntText(e.target.value);
                          setManualValidationErrors([]);
                          setManualOrderQtyByGroup((prev)=> ({ ...prev, [groupName]: raw }));
                        }}
                        onBlur={(e)=>{
                          const raw = normalizeIntText(e.target.value);
                          setManualOrderQtyByGroup((prev)=> ({ ...prev, [groupName]: raw }));
                        }}
                        onKeyDown={(e)=>{
                          const k = (e as any).key;
                          if (k === 'e' || k === 'E' || k === '.' || k === '-' || k === '+' ) {
                            e.preventDefault();
                          }
                        }}
                        inputProps={{ inputMode: 'numeric', pattern: '[0-9]*', min: 0, step: 1 }}
                        sx={{ width: 100 }}
                      />
                    );
                  }
                });
                cols.push({
                  field: 'remove',
                  headerName: 'Remove',
                  width: 110,
                  sortable: false,
                  filterable: false,
                  renderCell: (p: any) => {
                    const groupName = String(p?.row?.groupName || '');
                    return (
                      <IconButton
                        size="small"
                        disabled={manualIsLocked}
                        onClick={() => {
                          setManualOrderGroups((prev)=> prev.filter((x)=> String(x) !== groupName));
                          setManualOrderQtyByGroup((prev)=> {
                            const next = { ...prev };
                            delete (next as any)[groupName];
                            return next;
                          });
                        }}
                      >
                        <DeleteIcon fontSize="small" />
                      </IconButton>
                    );
                  }
                });
                return cols;
              })()}
              disableRowSelectionOnClick
              density="compact"
              slots={{ toolbar: GridToolbar }}
              slotProps={{ toolbar: { showQuickFilter: true, quickFilterProps: { debounceMs: 250 } } as any }}
              pagination
              pageSizeOptions={[5, 10, 20, 50]}
              initialState={{ pagination: { paginationModel: { page: 0, pageSize: 10 } } }}
            />
          </div>

          <Dialog open={manualPickOpen} onClose={()=> setManualPickOpen(false)} fullWidth maxWidth="xl">
            <DialogTitle>Select Pallets</DialogTitle>
            <DialogContent>
              <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} sx={{ mb: 1, mt: 1 }}>
                <TextField
                  size="small"
                  label="Search Pallet ID / Pallet Description"
                  value={manualPickerQ}
                  onChange={(e)=>setManualPickerQ(e.target.value)}
                  sx={{ flex: 1, minWidth: 260 }}
                />
                <Button variant="outlined" onClick={async ()=>{
                  try {
                    if (!manualWarehouseId) return;
                    setManualPickerLoading(true);
                    const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: manualWarehouseId, q: manualPickerQ.trim() || undefined } });
                    setManualPickerRows(Array.isArray(data?.rows) ? data.rows : []);
                    setManualPickerWarehouses(Array.isArray(data?.warehouses) ? data.warehouses : []);
                  } catch {
                    setManualPickerRows([]);
                    setManualPickerWarehouses([]);
                  } finally {
                    setManualPickerLoading(false);
                  }
                }} disabled={!manualWarehouseId || manualPickerLoading}>Refresh</Button>
              </Stack>
              <div style={{ height: 520, width: '100%' }}>
                <DataGrid
                  rows={(Array.isArray(manualPickerRows) ? manualPickerRows : []).map((r: any) => ({ id: String(r.groupName || r.lineItem || ''), ...r }))}
                  columns={(() => {
                    const cols: any[] = [
                      { field: 'lineItem', headerName: 'Pallet ID', width: 140, renderCell: (p: any) => String(p?.row?.lineItem || '-') },
                      { field: 'groupName', headerName: 'Pallet Description', flex: 1, minWidth: 220, renderCell: (p: any) => String(p?.row?.groupName || '-') },
                      { field: 'selectedWarehouseAvailable', headerName: `THIS - ${manualWarehouseName || 'Warehouse'}`, width: 170, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => String(p?.row?.selectedWarehouseAvailable ?? 0) },
                      {
                        field: 'onWaterPallets',
                        headerName: 'On-Water',
                        width: 120,
                        type: 'number',
                        align: 'right',
                        headerAlign: 'right',
                        renderCell: (p: any) => {
                          const qty = Number(p?.row?.onWaterPallets ?? 0);
                          const groupName = String(p?.row?.groupName || '').trim();
                          const wid = String(manualWarehouseId || '').trim();
                          if (!qty) return '0';
                          return (
                            <Button
                              variant="text"
                              size="small"
                              onClick={() => openOnWaterDetails({ warehouseId: wid, groupName })}
                              sx={{ minWidth: 0, p: 0, textDecoration: 'underline', fontSize: 16, fontWeight: 700 }}
                            >
                              {qty}
                            </Button>
                          );
                        },
                      },
                    ];

                    const wid = String(manualWarehouseId || '').trim();
                    const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
                    const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
                    const secondId = second ? String(second._id) : '';

                    if (secondId && second) {
                      cols.push({
                        field: `wh_${secondId}`,
                        headerName: second.name,
                        width: 120,
                        type: 'number',
                        align: 'right',
                        headerAlign: 'right',
                        renderCell: (p: any) => {
                          const per = (p?.row as any)?.perWarehouse || {};
                          return String(Number(per?.[String(secondId)] ?? 0));
                        },
                      });
                    }

                    for (const w of (manualPickerWarehouses || [])) {
                      if (String(w?._id || '') === String(manualWarehouseId)) continue;
                      if (secondId && String(w?._id || '') === String(secondId)) continue;
                      cols.push({
                        field: `wh_${w._id}`,
                        headerName: w.name,
                        width: 120,
                        type: 'number',
                        align: 'right',
                        headerAlign: 'right',
                        renderCell: (p: any) => {
                          const per = (p?.row as any)?.perWarehouse || {};
                          return String(Number(per?.[String(w._id)] ?? 0));
                        },
                      });
                    }

                    cols.push({
                      field: 'onProcessPallets',
                      headerName: 'On-Process',
                      width: 120,
                      type: 'number',
                      align: 'right',
                      headerAlign: 'right',
                      renderCell: (p: any) => {
                        const qty = Number(p?.row?.onProcessPallets ?? 0);
                        const groupName = String(p?.row?.groupName || '').trim();
                        if (!qty) return '0';
                        return (
                          <Button
                            variant="text"
                            size="small"
                            onClick={() => openOnProcessDetails({ groupName })}
                            sx={{ minWidth: 0, p: 0, textDecoration: 'underline', fontSize: 16, fontWeight: 700 }}
                          >
                            {qty}
                          </Button>
                        );
                      },
                    });

                    cols.push({
                      field: 'maxOrder',
                      headerName: 'Max Order',
                      width: 140,
                      type: 'number',
                      align: 'right',
                      headerAlign: 'right',
                      sortable: true,
                      valueGetter: (...args: any[]) => {
                        const maybeParams = args?.[0];
                        const maybeRow = args?.[1];
                        const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
                        const primary = Number(row?.selectedWarehouseAvailable ?? 0);
                        const onWater = Number(row?.onWaterPallets ?? 0);
                        const onProcess = Number(row?.onProcessPallets ?? 0);
                        let secondQty = 0;
                        if (secondId) {
                          const per = row?.perWarehouse || {};
                          secondQty = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                        }
                        const total =
                          (Number.isFinite(primary) ? primary : 0) +
                          (Number.isFinite(onWater) ? onWater : 0) +
                          (Number.isFinite(onProcess) ? onProcess : 0) +
                          (Number.isFinite(secondQty) ? secondQty : 0);
                        return Math.max(0, Math.floor(total));
                      },
                      renderCell: (p: any) => {
                        const v = Number((p && typeof p === 'object' && 'value' in p) ? (p as any).value : 0);
                        return String(Number.isFinite(v) ? v : 0);
                      },
                    });
                    return cols;
                  })()}
                  loading={manualPickerLoading}
                  getRowClassName={(params: any) => {
                    const row = (params as any)?.row || {};
                    const primary = Number(row?.selectedWarehouseAvailable ?? 0);
                    const onWater = Number(row?.onWaterPallets ?? 0);
                    const onProcess = Number(row?.onProcessPallets ?? 0);
                    const wid = String(manualWarehouseId || '').trim();
                    const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
                    const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
                    const secondId = second ? String(second._id) : '';
                    let secondQty = 0;
                    if (secondId) {
                      const per = row?.perWarehouse || {};
                      secondQty = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                    }
                    const total =
                      (Number.isFinite(primary) ? primary : 0) +
                      (Number.isFinite(onWater) ? onWater : 0) +
                      (Number.isFinite(onProcess) ? onProcess : 0) +
                      (Number.isFinite(secondQty) ? secondQty : 0);
                    const maxOrder = Math.max(0, Math.floor(total));
                    return maxOrder <= 0 ? 'row-maxorder-zero' : '';
                  }}
                  sx={{
                    '& .row-maxorder-zero': {
                      bgcolor: 'rgba(211, 47, 47, 0.08)',
                      '&:hover': { bgcolor: 'rgba(211, 47, 47, 0.12)' },
                    },
                  }}
                  checkboxSelection
                  isRowSelectable={(params: any) => {
                    const row = (params as any)?.row || {};
                    const primary = Number(row?.selectedWarehouseAvailable ?? 0);
                    const onWater = Number(row?.onWaterPallets ?? 0);
                    const onProcess = Number(row?.onProcessPallets ?? 0);
                    const wid = String(manualWarehouseId || '').trim();
                    const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
                    const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
                    const secondId = second ? String(second._id) : '';
                    let secondQty = 0;
                    if (secondId) {
                      const per = row?.perWarehouse || {};
                      secondQty = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                    }
                    const total =
                      (Number.isFinite(primary) ? primary : 0) +
                      (Number.isFinite(onWater) ? onWater : 0) +
                      (Number.isFinite(onProcess) ? onProcess : 0) +
                      (Number.isFinite(secondQty) ? secondQty : 0);
                    const maxOrder = Math.max(0, Math.floor(total));
                    return maxOrder > 0;
                  }}
                  disableRowSelectionOnClick
                  density="compact"
                  slots={{ toolbar: GridToolbar }}
                  slotProps={{ toolbar: { showQuickFilter: true, quickFilterProps: { debounceMs: 250 } } as any }}
                  onRowSelectionModelChange={(m: any)=> {
                    const idsArr = m?.ids ? Array.from(m.ids) : [];
                    const idsStr = idsArr.map((x: any) => String(x));
                    const byId = new Map(
                      (Array.isArray(manualPickerRows) ? manualPickerRows : [])
                        .map((r: any) => ({ id: String(r?.groupName || r?.lineItem || ''), row: r }))
                        .filter((x: any) => x.id)
                        .map((x: any) => [x.id, x.row])
                    );

                    const wid = String(manualWarehouseId || '').trim();
                    const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
                    const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
                    const secondId = second ? String(second._id) : '';

                    const allowed = new Set<string>();
                    for (const id of idsStr) {
                      const row = byId.get(id) || {};
                      const primary = Number(row?.selectedWarehouseAvailable ?? 0);
                      const onWater = Number(row?.onWaterPallets ?? 0);
                      const onProcess = Number(row?.onProcessPallets ?? 0);
                      let secondQty = 0;
                      if (secondId) {
                        const per = row?.perWarehouse || {};
                        secondQty = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                      }
                      const total =
                        (Number.isFinite(primary) ? primary : 0) +
                        (Number.isFinite(onWater) ? onWater : 0) +
                        (Number.isFinite(onProcess) ? onProcess : 0) +
                        (Number.isFinite(secondQty) ? secondQty : 0);
                      const maxOrder = Math.max(0, Math.floor(total));
                      if (maxOrder > 0) allowed.add(id);
                    }

                    setManualPickSelected({ type: 'include', ids: allowed });
                  }}
                  rowSelectionModel={manualPickSelected}
                  pagination
                  pageSizeOptions={[10, 20, 50, 100]}
                  initialState={{ pagination: { paginationModel: { page: 0, pageSize: 20 } } }}
                />
              </div>
            </DialogContent>
            <DialogActions>
              <Button onClick={()=> setManualPickOpen(false)}>Cancel</Button>
              <Button variant="contained" onClick={()=>{
                const idsArr = manualPickSelected?.ids ? Array.from(manualPickSelected.ids) : [];
                const selected = new Set(idsArr.map((x)=> String(x)));
                const wid = String(manualWarehouseId || '').trim();
                const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
                const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
                const secondId = second ? String(second._id) : '';
                const picked = (manualPickerRows || [])
                  .map((r: any) => ({
                    id: String(r?.groupName || r?.lineItem || ''),
                    row: r,
                  }))
                  .filter((x: any) => x.id && selected.has(x.id))
                  .filter((x: any) => {
                    const row = x.row || {};
                    const primary = Number(row?.selectedWarehouseAvailable ?? 0);
                    const onWater = Number(row?.onWaterPallets ?? 0);
                    const onProcess = Number(row?.onProcessPallets ?? 0);
                    let secondQty = 0;
                    if (secondId) {
                      const per = row?.perWarehouse || {};
                      secondQty = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                    }
                    const total =
                      (Number.isFinite(primary) ? primary : 0) +
                      (Number.isFinite(onWater) ? onWater : 0) +
                      (Number.isFinite(onProcess) ? onProcess : 0) +
                      (Number.isFinite(secondQty) ? secondQty : 0);
                    const maxOrder = Math.max(0, Math.floor(total));
                    return maxOrder > 0;
                  })
                  .map((x: any) => x.id);
                setManualOrderGroups((prev)=> {
                  const base = Array.isArray(prev) ? prev : [];
                  const set = new Set(base.map((x)=> String(x)));
                  for (const id of picked) set.add(String(id));
                  return Array.from(set);
                });
                setManualPickSelected({ type: 'include', ids: new Set() });
                setManualPickOpen(false);
              }} disabled={!(manualPickSelected?.ids && manualPickSelected.ids.size > 0)}>Add to Order</Button>
            </DialogActions>
          </Dialog>
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
              manualIsLocked ||
              (manualOrderGroups.length === 0) ||
              (Object.entries(manualOrderQtyByGroup || {}).filter(([g, q]) => manualOrderGroups.includes(g) && Number(q) > 0).length === 0)
            }
          >
            Save
          </Button>
        </DialogActions>
      </Dialog>

      <Dialog open={viewOrderableOpen} onClose={()=> setViewOrderableOpen(false)} fullWidth maxWidth="xl">
        <DialogTitle>VIEW ORDERABLE PALLETS</DialogTitle>
        <DialogContent>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} sx={{ mb: 1, mt: 1 }}>
            <TextField
              size="small"
              label="Search Pallet ID / Pallet Description"
              value={viewOrderableQ}
              onChange={(e)=>setViewOrderableQ(e.target.value)}
              sx={{ flex: 1, minWidth: 260 }}
            />
            <Button
              variant="outlined"
              onClick={exportAllOrderablePalletsXlsx}
              disabled={viewOrderableLoading || viewOrderableExporting || !viewOrderableWarehouseId}
            >
              Export XLSX
            </Button>
            <Button
              variant="outlined"
              onClick={openViewOrderableItems}
              disabled={!viewOrderableWarehouseId || viewOrderableLoading}
            >
              Refresh
            </Button>
          </Stack>

          <div style={{ height: 520, width: '100%' }}>
            <DataGrid
              rows={(Array.isArray(viewOrderableFilteredRows) ? viewOrderableFilteredRows : []).map((r: any) => ({ id: String(r.groupName || r.lineItem || ''), ...r }))}
              columns={(() => {
                const selectedWarehouseName =
                  String((Array.isArray(warehouses) ? warehouses : []).find((w: any) => String(w?._id || '') === String(viewOrderableWarehouseId))?.name || '').trim();
                const wid = String(viewOrderableWarehouseId || '').trim();
                const list = Array.isArray(viewOrderableWarehouses) ? viewOrderableWarehouses : [];
                const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
                const secondId = second ? String(second._id) : '';
                const secondName = second ? String(second.name || '').trim() : '';

                const cols: any[] = [
                  { field: 'lineItem', headerName: 'Pallet ID', width: 140, renderCell: (p: any) => String(p?.row?.lineItem || '-') },
                  { field: 'groupName', headerName: 'Pallet Description', flex: 1, minWidth: 220, renderCell: (p: any) => String(p?.row?.groupName || '-') },
                  { field: 'selectedWarehouseAvailable', headerName: `THIS - ${selectedWarehouseName || 'Warehouse'}`, width: 170, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => String(p?.row?.selectedWarehouseAvailable ?? 0) },
                  { field: 'onWaterPallets', headerName: 'On-Water', width: 120, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => String(p?.row?.onWaterPallets ?? 0) },
                ];

                if (secondId) {
                  cols.push({
                    field: 'secondWarehouseAvailable',
                    headerName: secondName || '2nd Warehouse',
                    width: 150,
                    type: 'number',
                    align: 'right',
                    headerAlign: 'right',
                    sortable: true,
                    filterable: false,
                    valueGetter: (...args: any[]) => {
                      const maybeParams = args?.[0];
                      const maybeRow = args?.[1];
                      const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
                      const per = row?.perWarehouse || {};
                      const v = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                      return Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
                    },
                    renderCell: (p: any) => {
                      const per = (p?.row as any)?.perWarehouse || {};
                      const v = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                      return String(Number.isFinite(v) ? v : 0);
                    },
                  });
                }

                cols.push({
                  field: 'onProcessPallets',
                  headerName: 'On-Process',
                  width: 120,
                  type: 'number',
                  align: 'right',
                  headerAlign: 'right',
                  renderCell: (p: any) => String(p?.row?.onProcessPallets ?? 0),
                });

                cols.push({
                  field: 'maxOrder',
                  headerName: 'Max Order',
                  width: 140,
                  type: 'number',
                  align: 'right',
                  headerAlign: 'right',
                  sortable: true,
                  filterable: false,
                  valueGetter: (...args: any[]) => {
                    const maybeParams = args?.[0];
                    const maybeRow = args?.[1];
                    const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
                    const primary = Number(row?.selectedWarehouseAvailable ?? 0);
                    const onWater = Number(row?.onWaterPallets ?? 0);
                    const onProcess = Number(row?.onProcessPallets ?? 0);
                    let secondQty = 0;
                    if (secondId) {
                      const per = row?.perWarehouse || {};
                      secondQty = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                    }
                    const total =
                      (Number.isFinite(primary) ? primary : 0) +
                      (Number.isFinite(onWater) ? onWater : 0) +
                      (Number.isFinite(onProcess) ? onProcess : 0) +
                      (Number.isFinite(secondQty) ? secondQty : 0);
                    return Math.max(0, Math.floor(total));
                  },
                  renderCell: (p: any) => {
                    const v = Number((p && typeof p === 'object' && 'value' in p) ? (p as any).value : 0);
                    return String(Number.isFinite(v) ? v : 0);
                  },
                });

                return cols;
              })()}
              loading={viewOrderableLoading}
              getRowClassName={(params: any) => {
                const row = (params as any)?.row || {};
                const primary = Number(row?.selectedWarehouseAvailable ?? 0);
                const onWater = Number(row?.onWaterPallets ?? 0);
                const onProcess = Number(row?.onProcessPallets ?? 0);
                const list = Array.isArray(viewOrderableWarehouses) ? viewOrderableWarehouses : [];
                const wid = String(viewOrderableWarehouseId || '').trim();
                const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
                const secondId = second ? String(second._id) : '';
                let secondQty = 0;
                if (secondId) {
                  const per = row?.perWarehouse || {};
                  secondQty = Number((per && typeof per === 'object') ? (per[secondId] ?? per[String(secondId)] ?? 0) : 0);
                }
                const total =
                  (Number.isFinite(primary) ? primary : 0) +
                  (Number.isFinite(onWater) ? onWater : 0) +
                  (Number.isFinite(onProcess) ? onProcess : 0) +
                  (Number.isFinite(secondQty) ? secondQty : 0);
                const maxOrder = Math.max(0, Math.floor(total));
                return maxOrder <= 0 ? 'row-maxorder-zero' : '';
              }}
              sx={{
                '& .row-maxorder-zero': {
                  bgcolor: 'rgba(211, 47, 47, 0.08)',
                  '&:hover': { bgcolor: 'rgba(211, 47, 47, 0.12)' },
                },
              }}
              disableRowSelectionOnClick
              density="compact"
              slots={{ toolbar: GridToolbar }}
              slotProps={{ toolbar: { showQuickFilter: true, quickFilterProps: { debounceMs: 250 } } as any }}
              pagination
              pageSizeOptions={[10, 20, 50, 100]}
              initialState={{ pagination: { paginationModel: { page: 0, pageSize: 20 } } }}
            />
          </div>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setViewOrderableOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={onWaterOpen} onClose={()=> setOnWaterOpen(false)} fullWidth maxWidth="md">
        <DialogTitle>{`On-Water - ${onWaterGroupName || ''}`}</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          {onWaterLoading ? <LinearProgress sx={{ mb: 2 }} /> : null}
          <div style={{ height: 420, width: '100%' }}>
            <DataGrid
              rows={(onWaterRows || []).map((r: any, idx: number) => ({ id: r?.id || `${idx}`, ...r }))}
              columns={([
                { field: 'reference', headerName: 'Reference', flex: 1, minWidth: 180 },
                { field: 'edd', headerName: 'EDD', width: 140 },
                { field: 'qty', headerName: 'QTY', width: 120, type: 'number', align: 'right', headerAlign: 'right' },
              ]) as GridColDef[]}
              disableRowSelectionOnClick
              density="compact"
              pagination
              pageSizeOptions={[5, 10, 20]}
              initialState={{ pagination: { paginationModel: { page: 0, pageSize: 10 } } }}
            />
          </div>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setOnWaterOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={onProcessOpen} onClose={()=> setOnProcessOpen(false)} fullWidth maxWidth="md">
        <DialogTitle>{`On-Process - ${onProcessGroupName || ''}`}</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          {onProcessLoading ? <LinearProgress sx={{ mb: 2 }} /> : null}
          <div style={{ height: 420, width: '100%' }}>
            <DataGrid
              rows={(onProcessRows || []).map((r: any, idx: number) => ({ id: r?.id || `${idx}`, ...r }))}
              columns={([
                { field: 'reference', headerName: 'Reference', flex: 1, minWidth: 180 },
                { field: 'edd', headerName: 'EDD', width: 140 },
                { field: 'qty', headerName: 'QTY', width: 120, type: 'number', align: 'right', headerAlign: 'right' },
              ]) as GridColDef[]}
              disableRowSelectionOnClick
              density="compact"
              pagination
              pageSizeOptions={[5, 10, 20]}
              initialState={{ pagination: { paginationModel: { page: 0, pageSize: 10 } } }}
            />
          </div>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setOnProcessOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
