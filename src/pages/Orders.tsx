import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, IconButton, Dialog, DialogTitle, DialogContent, DialogActions, LinearProgress, MenuItem, Box, Chip, Accordion, AccordionSummary, AccordionDetails, Tooltip } from '@mui/material';
import Autocomplete from '@mui/material/Autocomplete';
import DeleteIcon from '@mui/icons-material/Delete';
import ContentCopyIcon from '@mui/icons-material/ContentCopy';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import HelpOutlineIcon from '@mui/icons-material/HelpOutline';
import XLSX from 'xlsx-js-style';
import { DataGrid, GridColDef, GridToolbar } from '@mui/x-data-grid';
import api from '../api';
import { useToast } from '../components/ToastProvider';
import { formatDateTimeUS } from '../utils/datetime';
import pdfMake from 'pdfmake/build/pdfmake';
import pdfFonts from 'pdfmake/build/vfs_fonts';
// Support different export shapes between CJS/ESM bundles
try {
  const anyFonts: any = pdfFonts as any;
  (pdfMake as any).vfs = anyFonts?.pdfMake?.vfs || anyFonts?.vfs || (pdfMake as any).vfs || {};
} catch {}

type Warehouse = { _id: string; name: string };

type Allocation = { groupName?: string; qty?: number; source?: string; warehouseId?: string };

type OrderStatus = 'processing' | 'ready_to_ship' | 'shipped' | 'completed' | 'canceled';
type OrdersRow = {
  id: string;
  rawId: string;
  orderNumber: string;
  type: 'import' | 'manual';
  status: string;
  warehouseId: string;
  warehouseName?: string;
  createdAt: string;
  updatedAt?: string;
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
  originalPrice?: any;
  shippingPercent?: any;
  discountPercent?: any;
  finalPrice?: any;
  estFulfillmentDate?: string;
  estDeliveredDate?: string;
  notes?: string;
  source?: string;
};

export default function Orders() {
  const todayYmd = useMemo(() => new Date().toISOString().slice(0, 10), []);
  const lastAutoSavedShipdateRef = useRef<{ orderId: string; ymd: string; at: number }>({ orderId: '', ymd: '', at: 0 });
  const lastAutoSuggestedShipdateRef = useRef<{ ymd: string; at: number }>({ ymd: '', at: 0 });
  const shipmentsRebalanceTimerRef = useRef<any>(null);
  const manualEditRefreshInFlightRef = useRef(false);
  const manualEditRefreshLastAtRef = useRef(0);
  const toast = useToast();
  const [groupPriceByName, setGroupPriceByName] = useState<Record<string, number>>({});
  const [warehouses, setWarehouses] = useState<Warehouse[]>([]);
  const [ordersRows, setOrdersRows] = useState<OrdersRow[]>([]);
  const ordersRowsRef = useRef<OrdersRow[]>([]);
  const [ordersLoading, setOrdersLoading] = useState(false);
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
  const [manualLastUpdatedAt, setManualLastUpdatedAt] = useState('');
  const [manualLastUpdatedBy, setManualLastUpdatedBy] = useState('');
  const [manualEstFulfillment, setManualEstFulfillment] = useState('');
  const [manualEstDelivered, setManualEstDelivered] = useState('');
  const [manualShippingAddress, setManualShippingAddress] = useState('');
  const [manualNotes, setManualNotes] = useState('');
  const [manualOriginalPrice, setManualOriginalPrice] = useState('');
  const [manualShippingPercent, setManualShippingPercent] = useState('');
  const [manualDiscountPercent, setManualDiscountPercent] = useState('');
  const [manualLineMode, setManualLineMode] = useState<'pallet_group' | 'line_item'>('pallet_group');
  const [manualUnitPriceByRow, setManualUnitPriceByRow] = useState<Record<string, number>>({});
  const [manualPalletGroupOptions, setManualPalletGroupOptions] = useState<string[]>([]);
  const [manualLineItemOptions, setManualLineItemOptions] = useState<string[]>([]);
  const [manualLines, setManualLines] = useState<Array<{ lineItem: string; qty: string }>>([{ lineItem: '', qty: '' }]);
  const [manualAvailable, setManualAvailable] = useState<Record<string, number>>({});
  const [manualValidationErrors, setManualValidationErrors] = useState<string[]>([]);
  const [manualRecalcTick, setManualRecalcTick] = useState(0);

  
  // Preserve original Pallet ID mapping from the order at the moment the modal opens
  const manualBaseLineItemByGroupRef = useRef<Map<string, string>>(new Map());

  // Utility to clear focus before closing dialogs to avoid aria-hidden + focus conflicts
  const blurActive = useCallback(() => {
    try {
      const el = (document?.activeElement as any) as HTMLElement | null;
      if (el && typeof el.blur === 'function') el.blur();
    } catch {}
  }, []);

  // Refs to late-defined helpers to avoid TDZ in early effects
  const refreshPickerRef = useRef<null | ((wid: string) => Promise<any>)>(null);
  const refreshAllocRef = useRef<null | ((rawId: string) => Promise<any>)>(null);

  // Auto-refresh Orders list at a modest cadence, without forcing rebalances each time
  useEffect(() => {
    let stopped = false;
    let ticking = false;
    const tick = async () => {
      if (stopped) return;
      if (manualOpen) return; // don't interrupt when add/edit dialog is open
      if (ticking) return;
      ticking = true;
      try { await loadOrders(); } catch {}
      ticking = false;
    };
    const onFocus = () => { void tick(); };
    const interval = setInterval(tick, 120000); // every 120s
    window.addEventListener('focus', onFocus);
    const kick = setTimeout(tick, 600);
    return () => {
      stopped = true;
      clearInterval(interval);
      clearTimeout(kick);
      window.removeEventListener('focus', onFocus);
    };
  }, [manualOpen]);

  // While viewing/editing an order, refresh the Reserved Stock (Hierarchy) periodically without spamming rebalance
  useEffect(() => {
    if (!manualOpen) return;
    const rawId = String((manualEditRow as any)?.rawId || '').trim();
    const wid = String(manualWarehouseId || '').trim();
    let stopped = false;
    let ticking = false;
    const tick = async () => {
      if (stopped) return;
      if (ticking) return;
      ticking = true;
      try { if (wid) await refreshPickerRef.current?.(wid); } catch {}
      try { if (rawId) await refreshAllocRef.current?.(rawId); } catch {}
      ticking = false;
    };
    const onFocus = () => { void tick(); };
    const interval = setInterval(tick, 45000); // every 45s while dialog is open
    window.addEventListener('focus', onFocus);
    const kick = setTimeout(tick, 300);
    return () => {
      stopped = true;
      clearInterval(interval);
      clearTimeout(kick);
      window.removeEventListener('focus', onFocus);
    };
  }, [manualOpen, manualEditRow, manualWarehouseId]);

  // After helpers are declared, assign them to refs for early effects to use safely
  useEffect(() => {
    // @ts-ignore
    refreshPickerRef.current = typeof refreshManualPicker === 'function' ? refreshManualPicker : null;
    // @ts-ignore
    refreshAllocRef.current = typeof refreshManualAllocations === 'function' ? refreshManualAllocations : null;
  });

  const keyToGroupName = (key: string) => String(key || '').split('||')[0];

  const makeRowKey = (groupName: string) => {
    const g = String(groupName || '').trim();
    const uid = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    return `${g}||${uid}`;
  };

  // Snapshots captured at Add time (declare before any usage below)
  const [manualUnitPriceByGroup, setManualUnitPriceByGroup] = useState<Record<string, number>>({});
  const [manualItemsByGroup, setManualItemsByGroup] = useState<Record<string, any[]>>({});

  const openOrderableGroupItems = useCallback(async ({ groupName }: { groupName: string }) => {
    const g = String(groupName || '').trim();
    if (!g) return;
    setOrderableGroupOpen(true);
    setOrderableGroupName(g);
    setOrderableGroupItems([]);
    setOrderableGroupLoading(true);
    try {
      const gLower = g.toLowerCase();
      const snap = (manualItemsByGroup as any)?.[gLower];
      if (Array.isArray(snap) && snap.length) {
        setOrderableGroupItems(snap);
      } else {
        // Fallback: check currently loaded order lines for persisted snapshot (itemsAtAdd)
        let fromOrder: any[] | null = null;
        try {
          const lines = Array.isArray((manualEditRow as any)?.lines) ? (manualEditRow as any).lines : [];
          const match = lines.find((l: any) => String(l?.groupName || l?.lineItem || '').trim().toLowerCase() === gLower);
          if (match && Array.isArray((match as any)?.itemsAtAdd) && (match as any).itemsAtAdd.length) {
            fromOrder = (match as any).itemsAtAdd as any[];
          }
        } catch {}
        if (Array.isArray(fromOrder) && fromOrder.length) {
          setOrderableGroupItems(fromOrder);
        } else {
          // Fallback: try localStorage cache per order
          let fromCache: any[] | null = null;
          try {
            const orderKey = String((manualEditRow as any)?.rawId || (manualEditRow as any)?.id || '').trim();
            if (orderKey) {
              const cacheKey = `orderLineSnapshots:${orderKey}`;
              const cached = localStorage.getItem(cacheKey);
              if (cached) {
                const parsed = JSON.parse(cached);
                const maybe = parsed && parsed.items ? parsed.items[gLower] || parsed.items[g] : null;
                if (Array.isArray(maybe) && maybe.length) fromCache = maybe as any[];
              }
            }
          } catch {}
          if (Array.isArray(fromCache) && fromCache.length) {
            setOrderableGroupItems(fromCache);
          } else {
            const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
            const list = Array.isArray((data as any)?.items) ? (data as any).items : [];
            setOrderableGroupItems(list);
            // LEGACY BACKFILL: If no snapshot exists anywhere, capture the current list as a snapshot for this order/group
            try {
              const gLower = g.toLowerCase();
              const hasSnap = Array.isArray((manualItemsByGroup as any)?.[gLower]) && (manualItemsByGroup as any)[gLower].length;
              if (!hasSnap && list && list.length) {
                setManualItemsByGroup((prev)=> ({ ...(prev||{}), [gLower]: list }));
                const orderKey = String((manualEditRow as any)?.rawId || (manualEditRow as any)?.id || '').trim();
                if (orderKey) {
                  const cacheKey = `orderLineSnapshots:${orderKey}`;
                  const cached = localStorage.getItem(cacheKey);
                  const parsed = cached ? JSON.parse(cached) : {};
                  parsed.items = { ...(parsed.items || {}), [gLower]: list };
                  localStorage.setItem(cacheKey, JSON.stringify(parsed));
                }
              }
            } catch {}
          }
        }
      }
    } catch {
      setOrderableGroupItems([]);
    } finally {
      setOrderableGroupLoading(false);
    }
  }, [manualItemsByGroup, manualEditRow]);

  const [manualPickerLoading, setManualPickerLoading] = useState(false);
  const [manualPickerQ, setManualPickerQ] = useState('');
  const [manualPickerEddFrom, setManualPickerEddFrom] = useState('');
  const [manualPickerEddTo, setManualPickerEddTo] = useState('');
  const [manualPickerWarehouses, setManualPickerWarehouses] = useState<Array<{ _id: string; name: string }>>([]);
  const [manualPickerRows, setManualPickerRows] = useState<any[]>([]);
  const [manualOrderQtyByGroup, setManualOrderQtyByGroup] = useState<Record<string, string>>({});
  const [manualDiscountByGroup, setManualDiscountByGroup] = useState<Record<string, string>>({});
  const [manualShipdateTouched, setManualShipdateTouched] = useState(false);
  const [manualStatusTouched, setManualStatusTouched] = useState(false);
  const [manualPickOpen, setManualPickOpen] = useState(false);
  const [manualPickSelected, setManualPickSelected] = useState<any>({ type: 'include', ids: new Set() });
  const [manualOrderGroups, setManualOrderGroups] = useState<string[]>([]);
  const [manualBaseTiersByGroup, setManualBaseTiersByGroup] = useState<Record<string, { primary: number; onWater: number; second: number; onProcess: number; total: number }>>({});
  

  const [onWaterOpen, setOnWaterOpen] = useState(false);
  const [onWaterLoading, setOnWaterLoading] = useState(false);
  const [onWaterGroupName, setOnWaterGroupName] = useState('');

  // Debounced auto-fetch for pallet picker when typing search while dialog is open
  useEffect(() => {
    if (!manualPickOpen) return;
    const wid = String(manualWarehouseId || '').trim();
    if (!wid) return;
    const handle = setTimeout(async () => {
      try {
        setManualPickerLoading(true);
        const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: wid } });
        setManualPickerRows(Array.isArray((data as any)?.rows) ? (data as any).rows : []);
        setManualPickerWarehouses(Array.isArray((data as any)?.warehouses) ? (data as any).warehouses : []);
      } catch {
        setManualPickerRows([]);
        setManualPickerWarehouses([]);
      } finally {
        setManualPickerLoading(false);
      }
    }, 350);
    return () => clearTimeout(handle);
  }, [manualPickOpen, manualWarehouseId]);
  const [onWaterRows, setOnWaterRows] = useState<any[]>([]);

  const [onProcessOpen, setOnProcessOpen] = useState(false);
  const [onProcessLoading, setOnProcessLoading] = useState(false);
  const [onProcessGroupName, setOnProcessGroupName] = useState('');
  const [onProcessRows, setOnProcessRows] = useState<any[]>([]);

  const [viewOrderableOpen, setViewOrderableOpen] = useState(false);
  const [viewOrderableWarehouseId, setViewOrderableWarehouseId] = useState('');
  const [viewOrderableQ, setViewOrderableQ] = useState('');
  const [viewOrderableEddFrom, setViewOrderableEddFrom] = useState('');
  const [viewOrderableEddTo, setViewOrderableEddTo] = useState('');
  const [viewOrderableRows, setViewOrderableRows] = useState<any[]>([]);
  const [viewOrderableWarehouses, setViewOrderableWarehouses] = useState<any[]>([]);
  const [viewOrderableLoading, setViewOrderableLoading] = useState(false);
  const [viewOrderableExporting, setViewOrderableExporting] = useState(false);
  const [palletNameByGroup, setPalletNameByGroup] = useState<Record<string, string>>({});
  const [palletDescByGroup, setPalletDescByGroup] = useState<Record<string, string>>({});

  const [orderableGroupOpen, setOrderableGroupOpen] = useState(false);
  const [orderableGroupLoading, setOrderableGroupLoading] = useState(false);
  const [orderableGroupName, setOrderableGroupName] = useState('');
  const [orderableGroupItems, setOrderableGroupItems] = useState<any[]>([]);

  const [reportOpen, setReportOpen] = useState(false);
  const [reportFrom, setReportFrom] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    return d.toISOString().slice(0, 10);
  });
  const [reportTo, setReportTo] = useState(() => new Date().toISOString().slice(0, 10));
  const [reportExporting, setReportExporting] = useState(false);

  const viewOrderableFilteredRows = useMemo(() => {
    const q = String(viewOrderableQ || '').trim().toLowerCase();
    const rows = Array.isArray(viewOrderableRows) ? viewOrderableRows : [];

    const bySearch = !q
      ? rows
      : rows.filter((r: any) => {
          const gname = String(r?.groupName || '');
          const gLower = gname.toLowerCase();
          const pLower = String(palletNameByGroup?.[gLower] || '').toLowerCase();
          const dLower = String(palletDescByGroup?.[gLower] || '').toLowerCase();
          const hay = `${String(r?.lineItem || '')} ${gLower} ${pLower} ${dLower}`.toLowerCase();
          return hay.includes(q);
        });

    // Hide rows where total availability across tiers is zero
    const wid = String(viewOrderableWarehouseId || '').trim();
    const list = Array.isArray(viewOrderableWarehouses) ? viewOrderableWarehouses : [];
    const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
    const secondId = second ? String(second._id) : '';

    const from = String(viewOrderableEddFrom || '').trim();
    const to = String(viewOrderableEddTo || '').trim();
    const inRange = (ymd: string) => {
      const s = String(ymd || '').trim();
      if (!s) return false;
      if (from && s < from) return false;
      if (to && s > to) return false;
      return true;
    };
    const matchesEddRange = (row: any) => {
      if (!from && !to) return true;
      const ships = Array.isArray(row?.onWaterShipments) ? row.onWaterShipments : [];
      for (const x of ships) if (inRange(String(x?.edd || ''))) return true;
      const batches = Array.isArray(row?.onProcessBatches) ? row.onProcessBatches : [];
      for (const b of batches) if (inRange(String(b?.edd || ''))) return true;
      return false;
    };

    return bySearch.filter((row: any) => {
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
      if (!(Math.max(0, Math.floor(total)) > 0)) return false;
      return matchesEddRange(row);
    });
  }, [viewOrderableQ, viewOrderableRows, viewOrderableWarehouseId, viewOrderableWarehouses, viewOrderableEddFrom, viewOrderableEddTo, palletNameByGroup, palletDescByGroup]);

  useEffect(() => {
    let stopped = false;
    const loadPrices = async () => {
      try {
        // Mirror Pallet Registry logic: compute pallet price as sum of item prices per pallet group
        const [groupsRes, itemsRes] = await Promise.all([
          api.get<any[]>('/item-groups'),
          api.get<any[]>('/items', { params: { includeDisabled: 1 } }),
        ]);

        const groups = Array.isArray(groupsRes.data) ? groupsRes.data : [];
        const items = Array.isArray(itemsRes.data) ? itemsRes.data : [];

        const activeSet = new Set(
          groups
            .filter((g: any) => (g as any).active !== false)
            .map((g: any) => String(g?.name || '').trim())
            .filter((v: string) => v)
        );

        const map: Record<string, number> = {};
        for (const it of items as any[]) {
          const groupName = String((it as any)?.itemGroup || '').trim();
          if (!groupName || !activeSet.has(groupName)) continue;
          const key = groupName.toLowerCase();
          const prev = map[key] || 0;
          const pack = Number((it as any).packSize ?? 0) || 0;
          const price = Number((it as any).price ?? 0) || 0;
          map[key] = prev + (pack * price);
        }

        if (!stopped) setGroupPriceByName(map);
      } catch {
        if (!stopped) setGroupPriceByName({});
      }
    };
    loadPrices();
    return () => {
      stopped = true;
    };
  }, []);

  const openOnWaterDetails = useCallback(async ({ warehouseId, groupName }: { warehouseId: string; groupName: string }) => {
    const g = String(groupName || '').trim();
    const w = String(warehouseId || '').trim();
    if (!w || !g) return;
    setOnWaterOpen(true);
    setOnWaterGroupName(g);
    setOnWaterRows([]);
    setOnWaterLoading(true);
    try {
      const { data } = await api.get('/orders/pallet-picker/on-water', { params: { warehouseId: w, groupName: g } });
      setOnWaterRows(Array.isArray(data?.rows) ? data.rows : []);
    } catch {
      setOnWaterRows([]);
    } finally {
      setOnWaterLoading(false);
    }
  }, []);

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

      // Collect distinct EDD dates across all rows for On-Water and On-Process
      const waterEddsSet = new Set<string>();
      const processEddsSet = new Set<string>();
      for (const r of rows) {
        for (const s of (Array.isArray(r?.onWaterShipments) ? r.onWaterShipments : [])) {
          const d = String(s?.edd || '').trim();
          if (d) waterEddsSet.add(d);
        }
        for (const b of (Array.isArray(r?.onProcessBatches) ? r.onProcessBatches : [])) {
          const d = String(b?.edd || '').trim();
          if (d) processEddsSet.add(d);
        }
      }
      const waterEdds = Array.from(waterEddsSet).sort((a, b) => a.localeCompare(b));
      const processEdds = Array.from(processEddsSet).sort((a, b) => a.localeCompare(b));

      const exportRows = rows.map((r: any) => {
        const primary = Number(r?.selectedWarehouseAvailable ?? 0);
        const onWater = Number(r?.onWaterPallets ?? 0);
        const onProcess = Number(r?.onProcessPallets ?? 0);
        const groupName = String(r?.groupName || '').trim();
        const gLower = groupName.toLowerCase();
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
          'Pallet Name': String(r?.palletName || ''),
          'Pallet Description': String(palletDescByGroup?.[gLower] || ''),
          'Pallet ID': String(r?.lineItem || ''),
          'MPG': Math.max(0, Math.floor(Number.isFinite(primary) ? primary : 0)),
        };

        if (secondId) out['PEBA'] = Math.max(0, Math.floor(Number.isFinite(secondQty) ? secondQty : 0));

        // Per-EDD On-Water columns
        const shipMap = new Map<string, number>();
        for (const s of (Array.isArray(r?.onWaterShipments) ? r.onWaterShipments : [])) {
          const d = String(s?.edd || '').trim();
          const qty = Math.max(0, Math.floor(Number(s?.qty || 0)));
          if (d && qty > 0) shipMap.set(d, (shipMap.get(d) || 0) + qty);
        }
        for (const d of waterEdds) out[`On-Water EDD ${d}`] = Math.max(0, Math.floor(shipMap.get(d) || 0));

        // Per-EDD On-Process columns
        const procMap = new Map<string, number>();
        for (const b of (Array.isArray(r?.onProcessBatches) ? r.onProcessBatches : [])) {
          const d = String(b?.edd || '').trim();
          const qty = Math.max(0, Math.floor(Number(b?.qty || 0)));
          if (d && qty > 0) procMap.set(d, (procMap.get(d) || 0) + qty);
        }
        for (const d of processEdds) out[`On-Process EDD ${d}`] = Math.max(0, Math.floor(procMap.get(d) || 0));

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

  const exportPalletSalesReportXlsx = useCallback(async () => {
    const from = String(reportFrom || '').slice(0, 10);
    const to = String(reportTo || '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
      toast.error('Please select a valid date range');
      return;
    }

    setReportExporting(true);
    try {
      const { data } = await api.get('/reports/pallet-sales', {
        params: {
          from,
          to,
          top: 20,
        },
      });

      const pad = (n: number) => String(n).padStart(2, '0');
      const dt = new Date();
      const stamp = `${dt.getFullYear()}${pad(dt.getMonth() + 1)}${pad(dt.getDate())}_${pad(dt.getHours())}${pad(dt.getMinutes())}${pad(dt.getSeconds())}`;
      const safe = (v: any) => String(v ?? '').replace(/[/\\:*?"<>|]/g, '-');
      const filename = `pallet_sales_report_${safe(from)}_${safe(to)}_${stamp}.xlsx`;

      const rows: any[] = [];
      rows.push(['Pallet Sales Report']);
      rows.push([`Date Range: ${from} to ${to}`]);
      rows.push([`Exported At: ${formatDateTimeUS(dt)}`]);
      rows.push([]);

      rows.push(['Top Selling Pallets']);
      rows.push(['Pallet ID', 'Pallet Name', 'Pallet Description', 'Pallets Sold']);
      const topSelling = Array.isArray(data?.topSelling) ? data.topSelling : [];
      for (const r of topSelling) {
        const gLower = String(r?.groupName || '').trim().toLowerCase();
        rows.push([
          String(r?.palletId || ''),
          String(palletNameByGroup?.[gLower] || ''),
          String(palletDescByGroup?.[gLower] || ''),
          Number(r?.soldPallets || 0),
        ]);
      }
      if (!topSelling.length) {
        rows.push(['', '', '', 0]);
      }

      rows.push([]);
      rows.push(['Non-Performing Pallets']);
      rows.push(['Pallet ID', 'Pallet Name', 'Pallet Description', 'Pallets Sold', 'Reason']);
      const nonPerforming = Array.isArray(data?.nonPerforming) ? data.nonPerforming : [];
      for (const r of nonPerforming) {
        const gLower = String(r?.groupName || '').trim().toLowerCase();
        rows.push([
          String(r?.palletId || ''),
          String(palletNameByGroup?.[gLower] || ''),
          String(palletDescByGroup?.[gLower] || ''),
          Number(r?.soldPallets || 0),
          String(r?.reason || ''),
        ]);
      }
      if (!nonPerforming.length) {
        rows.push(['', '', '', 0, '']);
      }

      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Pallet Sales');
      XLSX.writeFile(wb, filename);
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to export report';
      toast.error(msg);
    } finally {
      setReportExporting(false);
    }
  }, [reportFrom, reportTo, toast, palletNameByGroup, palletDescByGroup]);

  const openManualEdit = async (row: OrdersRow) => {
    const toYmd = (v: any) => {
      const s = String(v || '');
      const slice = s.slice(0, 10);
      return /^\d{4}-\d{2}-\d{2}$/.test(slice) ? slice : '';
    };
    const toLocalDateTime = (v: any) => formatDateTimeUS(v);

    setManualMode('edit');
    setManualEditRow(row);
    setManualOpen(true);
    setManualWarehouseId(String(row?.warehouseId || fixedWarehouseId));
    setManualStatus((normalizeStatus(row?.status || '') as any) || 'processing');
    setManualStatusTouched(false);
    setManualCustomerEmail(String(row?.email || '').trim());
    setManualCustomerName(String(row?.customerName || '').trim());
    setManualCustomerPhone(String(row?.customerPhone || '').trim());
    setManualCreatedAt(toYmd(row?.createdAtOrder || row?.dateCreated || row?.createdAt));
    setManualLastUpdatedAt(toLocalDateTime((row as any)?.updatedAt || row?.createdAt || ''));
    setManualLastUpdatedBy(String((row as any)?.lastUpdatedBy || '').trim());
    setManualEstFulfillment(toYmd(row?.estFulfillmentDate));
    setManualEstDelivered(toYmd((row as any)?.estDeliveredDate));
    setManualShipdateTouched(false);
    setManualShippingAddress(String(row?.shippingAddress || '').trim());
    setManualNotes(String((row as any)?.notes || '').trim());
    setManualOriginalPrice(String((row as any)?.originalPrice ?? '').trim());
    setManualShippingPercent(String((row as any)?.shippingPercent ?? '').trim());
    setManualDiscountPercent(String((row as any)?.discountPercent ?? '').trim());
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
    setManualReservedBreakdown([]);

    const baseLines = Array.isArray(row?.lines) ? row.lines : [];
    const groups: string[] = [];
    const qtyBy: Record<string, string> = {};
    const discountBy: Record<string, string> = {};
    const baseTiersMap: Record<string, { primary: number; onWater: number; second: number; onProcess: number; total: number }> = {};
    const unitByRow: Record<string, number> = {};
    const unitByGroupSnap: Record<string, number> = {};
    const itemsByGroupSnap: Record<string, any[]> = {};
    const baseMap = new Map<string, string>();
    for (const l of baseLines) {
      const g = String(l?.groupName || l?.lineItem || '').trim();
      if (!g) continue;

      // Create a unique row key per line so duplicates are preserved as separate rows
      const rowKey = makeRowKey(g);
      groups.push(rowKey);

      const q = Math.floor(Number(l?.qty || 0));
      qtyBy[rowKey] = q > 0 ? String(q) : '';

      const gLower = String(l?.groupName || g || '').trim().toLowerCase();
      const id = String(l?.lineItem || '').trim();
      if (gLower && id && !baseMap.has(gLower)) baseMap.set(gLower, id);

      const discRaw = Number((l as any)?.discountPercent ?? (l as any)?.discount ?? 0);
      if (Number.isFinite(discRaw)) {
        const sanitized = Math.max(0, Math.min(100, Math.floor(discRaw)));
        discountBy[rowKey] = String(sanitized);
      }
      const unitRaw = Number((l as any)?.unitPrice);
      if (Number.isFinite(unitRaw) && unitRaw >= 0) {
        unitByRow[rowKey] = unitRaw;
      }
      // Hydrate per-group snapshots if present or derivable
      const gLowerSnap = String(l?.groupName || g || '').trim().toLowerCase();
      if (Number.isFinite(unitRaw) && unitRaw >= 0 && !(gLowerSnap in unitByGroupSnap)) {
        unitByGroupSnap[gLowerSnap] = unitRaw;
      }
      const itemsAtAdd = Array.isArray((l as any)?.itemsAtAdd) ? (l as any).itemsAtAdd : null;
      if (itemsAtAdd && itemsAtAdd.length && !(gLowerSnap in itemsByGroupSnap)) {
        itemsByGroupSnap[gLowerSnap] = itemsAtAdd as any[];
      }
      const bt = (l as any)?.baseTiers;
      if (bt && typeof bt === 'object') {
        const primary = Math.max(0, Math.floor(Number((bt as any)?.primary || 0)));
        const onWater = Math.max(0, Math.floor(Number((bt as any)?.onWater || 0)));
        const second = Math.max(0, Math.floor(Number((bt as any)?.second || 0)));
        const onProcess = Math.max(0, Math.floor(Number((bt as any)?.onProcess || 0)));
        const total = Math.max(0, Math.floor(Number((bt as any)?.total || (primary + onWater + second + onProcess))));
        const gLowerKey = String((l as any)?.groupName || '').trim().toLowerCase();
        if (gLowerKey && !baseTiersMap[gLowerKey]) baseTiersMap[gLowerKey] = { primary, onWater, second, onProcess, total };
      }
    }
    setManualOrderGroups(groups);
    setManualOrderQtyByGroup(qtyBy);
    setManualDiscountByGroup(discountBy);
    setManualBaseTiersByGroup(baseTiersMap);
    setManualUnitPriceByRow(unitByRow);
    if (Object.keys(unitByGroupSnap).length) setManualUnitPriceByGroup((prev)=> ({ ...(prev||{}), ...unitByGroupSnap }));
    if (Object.keys(itemsByGroupSnap).length) setManualItemsByGroup((prev)=> ({ ...(prev||{}), ...itemsByGroupSnap }));
    manualBaseLineItemByGroupRef.current = baseMap;
    // Kick a recompute so Available Qty/rows reflect hydrated quantities immediately
    setManualRecalcTick((t) => t + 1);

    // Load previously saved snapshots for this order (client-side cache)
    try {
      const orderKey = String((row as any)?.rawId || (row as any)?.id || '').trim();
      if (orderKey) {
        const cacheKey = `orderLineSnapshots:${orderKey}`;
        const cached = localStorage.getItem(cacheKey);
        if (cached) {
          const parsed = JSON.parse(cached);
          if (parsed && typeof parsed === 'object') {
            const u = (parsed as any).unit; const it = (parsed as any).items;
            if (u && typeof u === 'object') setManualUnitPriceByGroup((prev)=> ({ ...(prev||{}), ...(u as any) }));
            if (it && typeof it === 'object') setManualItemsByGroup((prev)=> ({ ...(prev||{}), ...(it as any) }));
          }
        }
      }
    } catch {}

    try {
      const wid = String(row?.warehouseId || fixedWarehouseId || '').trim();
      if (wid) {
        setManualPickerLoading(true);
        const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: wid } });
        setManualPickerRows(Array.isArray(data?.rows) ? data.rows : []);
        setManualPickerWarehouses(Array.isArray(data?.warehouses) ? data.warehouses : []);
        // Ensure grid recomputes once source availability data arrives
        setManualRecalcTick((t) => t + 1);
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
        // Fetch latest allocations and metadata without forcing server-side rebalances
        try { await new Promise((r) => setTimeout(r, 150)); } catch {}
        const { allocations, reserved } = await refreshManualAllocations(rawId);
        setManualAllocations(Array.isArray(allocations) ? allocations : []);
        setManualReservedBreakdown(Array.isArray(reserved) ? reserved : []);
        try {
          const { data } = await api.get(`/orders/unfulfilled/${rawId}`);
          setManualLastUpdatedAt(toLocalDateTime((data as any)?.updatedAt || (data as any)?.createdAt || ''));
          setManualLastUpdatedBy(String((data as any)?.lastUpdatedBy || '').trim());
          const nextStatus = (normalizeStatus((data as any)?.status || '') as any) || 'processing';
          if (!manualStatusTouched) setManualStatus(nextStatus);
          setManualEstFulfillment(toYmd((data as any)?.estFulfillmentDate));
        } catch {}
      }
    } catch {
      setManualAllocations(Array.isArray((row as any)?.allocations) ? ((row as any).allocations as any) : []);
      setManualReservedBreakdown([]);
      setManualLastUpdatedBy(String((row as any)?.lastUpdatedBy || '').trim());
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

  const openOnProcessDetails = useCallback(async ({ groupName }: { groupName: string }) => {
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
  }, []);

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

  // Load Pallet Name mapping early and cache it for fast startup
  useEffect(() => {
    let canceled = false;
    // Warm from cache for instant availability after restart
    try {
      const cached = localStorage.getItem('palletNameByGroup');
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          setPalletNameByGroup(parsed as Record<string, string>);
        }
      }
    } catch {}

    try {
      const cached = localStorage.getItem('palletDescByGroup');
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          setPalletDescByGroup(parsed as Record<string, string>);
        }
      }
    } catch {}

    (async () => {
      try {
        const { data } = await api.get<any[]>('/item-groups');
        if (canceled) return;
        const map: Record<string, string> = {};
        const descMap: Record<string, string> = {};
        for (const g of (Array.isArray(data) ? data : [])) {
          const name = String((g as any)?.name || '').trim();
          if (!name) continue;
          map[name.toLowerCase()] = String((g as any)?.palletName || '').trim();
          descMap[name.toLowerCase()] = String((g as any)?.palletDescription || '').trim();
        }
        setPalletNameByGroup(map);
        try { localStorage.setItem('palletNameByGroup', JSON.stringify(map)); } catch {}
        setPalletDescByGroup(descMap);
        try { localStorage.setItem('palletDescByGroup', JSON.stringify(descMap)); } catch {}
      } catch {
        if (!canceled) setPalletNameByGroup((m) => (m && Object.keys(m).length ? m : {}));
        if (!canceled) setPalletDescByGroup((m) => (m && Object.keys(m).length ? m : {}));
      }
    })();

    return () => { canceled = true; };
  }, []);

  const manualPrevStatus = useMemo(() => {
    if (!manualEditRow) return '';
    return (normalizeStatus(manualEditRow?.status || '') as any) || '';
  }, [manualEditRow]);

  const manualIsCanceled = manualMode === 'edit' && manualPrevStatus === 'canceled';
  const manualIsCompleted = manualMode === 'edit' && manualPrevStatus === 'completed';
  const manualIsShipped = manualMode === 'edit' && manualPrevStatus === 'shipped';
  const manualCanPostAction = manualMode === 'edit' && (manualPrevStatus === 'completed' || manualPrevStatus === 'shipped');
  const manualIsLocked = manualIsCanceled || manualIsCompleted;
  const manualFieldsLocked = manualIsLocked || manualIsShipped;

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

  // Attach palletName to rows once so filtering/rendering doesn't repeatedly map
  const manualPickerRowsWithName = useMemo(() => {
    const rows = Array.isArray(manualPickerRows) ? manualPickerRows : [];
    return rows.map((r: any) => {
      const gLower = String(r?.groupName || '').trim().toLowerCase();
      const palletName = String(r?.palletName || palletNameByGroup?.[gLower] || '');
      return { ...r, palletName };
    });
  }, [manualPickerRows, palletNameByGroup]);

  // Debounce the manual picker search input to avoid heavy re-renders while typing
  const [manualPickerQDebounced, setManualPickerQDebounced] = useState('');
  useEffect(() => {
    const h = setTimeout(() => setManualPickerQDebounced(String(manualPickerQ || '')), 150);
    return () => clearTimeout(h);
  }, [manualPickerQ]);

  const manualPickerRowsFiltered = useMemo(() => {
    const q = String(manualPickerQDebounced || '').trim().toLowerCase();
    const rows = Array.isArray(manualPickerRowsWithName) ? manualPickerRowsWithName : [];
    const wid = String(manualWarehouseId || '').trim();
    // Determine a single second warehouse like the grid does
    const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
    const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
    const secondId = second ? String(second._id) : '';

    const bySearch = !q
      ? rows
      : rows.filter((r: any) => {
          const gid = String(r?.lineItem || '').trim().toLowerCase();
          const gnameLower = String(r?.groupName || '').trim().toLowerCase();
          const pname = String(r?.palletName || '').trim().toLowerCase();
          const pdesc = String(palletDescByGroup?.[gnameLower] || '').trim().toLowerCase();
          return gid.includes(q) || gnameLower.includes(q) || pname.includes(q) || pdesc.includes(q);
        });

    // EDD date range filter: Only include rows having any On-Water or On-Process EDD within [from, to]
    const from = String(manualPickerEddFrom || '').trim();
    const to = String(manualPickerEddTo || '').trim();
    const inRange = (ymd: string) => {
      const s = String(ymd || '').trim();
      if (!s) return false;
      if (from && s < from) return false;
      if (to && s > to) return false;
      return true;
    };
    const matchesEddRange = (row: any) => {
      if (!from && !to) return true;
      const ships = Array.isArray(row?.onWaterShipments) ? row.onWaterShipments : [];
      for (const x of ships) if (inRange(String(x?.edd || ''))) return true;
      const batches = Array.isArray(row?.onProcessBatches) ? row.onProcessBatches : [];
      for (const b of batches) if (inRange(String(b?.edd || ''))) return true;
      return false;
    };

    // Hide rows with zero availability across tiers (Primary + On-Water + 2nd + On-Process)
    return bySearch.filter((row: any) => {
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
      if (!(Math.max(0, Math.floor(total)) > 0)) return false;
      return matchesEddRange(row);
    });
  }, [manualPickerRowsWithName, manualPickerQDebounced, manualWarehouseId, manualPickerWarehouses, manualPickerEddFrom, manualPickerEddTo, palletDescByGroup]);

  // Ensure Pallet Name map is available quickly when opening the manual picker dialog
  useEffect(() => {
    const needsMap = !palletNameByGroup || Object.keys(palletNameByGroup).length === 0;
    if (!manualPickOpen || !needsMap) return;
    let canceled = false;
    (async () => {
      try {
        const { data } = await api.get<any[]>('/item-groups');
        if (canceled) return;
        const map: Record<string, string> = {};
        for (const g of (Array.isArray(data) ? data : [])) {
          const name = String((g as any)?.name || '').trim().toLowerCase();
          if (!name) continue;
          map[name] = String((g as any)?.palletName || '').trim();
        }
        setPalletNameByGroup(map);
      } catch {
        if (!canceled) setPalletNameByGroup((m) => (m && Object.keys(m).length ? m : {}));
      }
    })();
    return () => { canceled = true; };
  }, [manualPickOpen, palletNameByGroup]);

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

  // Memoize EDD lists for the manual picker to avoid recomputation on each keystroke while searching
  const manualPickerEddLists = useMemo(() => {
    const from = String(manualPickerEddFrom || '').trim();
    const to = String(manualPickerEddTo || '').trim();
    const rows = Array.isArray(manualPickerRows) ? manualPickerRows : [];
    const waterEddsSet = new Set<string>();
    const processEddsSet = new Set<string>();
    for (const r of rows) {
      for (const s of (Array.isArray((r as any)?.onWaterShipments) ? (r as any).onWaterShipments : [])) {
        const edd = String((s as any)?.edd || '').trim();
        if (edd && (!from || edd >= from) && (!to || edd <= to)) waterEddsSet.add(edd);
      }
      for (const b of (Array.isArray((r as any)?.onProcessBatches) ? (r as any).onProcessBatches : [])) {
        const edd = String((b as any)?.edd || '').trim();
        if (edd && (!from || edd >= from) && (!to || edd <= to)) processEddsSet.add(edd);
      }
    }
    const waterEdds = Array.from(waterEddsSet).sort((a, b) => a.localeCompare(b));
    const processEdds = Array.from(processEddsSet).sort((a, b) => a.localeCompare(b));
    return { waterEdds, processEdds };
  }, [manualPickerRows, manualPickerEddFrom, manualPickerEddTo]);

  const manualPickerColumns = useMemo(() => {
    const cleanedPrimaryName = String(manualWarehouseName || 'Warehouse').replace(/^THIS\s*-\s*/i, '');
    const cols: any[] = [
      { field: 'palletName', headerName: 'Pallet Name', flex: 1, minWidth: 200, renderCell: (p: any) => {
        const fromRow = String((p?.row as any)?.palletName || '').trim();
        const gRaw = String(p?.row?.groupName || '').trim();
        const gLower = gRaw.toLowerCase();
        const isAdded = (manualOrderGroups || []).some((k) => keyToGroupName(String(k || '')).trim().toLowerCase() === gLower);
        const base = fromRow || String(palletNameByGroup?.[gLower] || '');
        return isAdded && base ? `${base} — Added` : (base || '');
      } },
      { field: 'palletDescription', headerName: 'Pallet Description', flex: 1, minWidth: 220, renderCell: (p: any) => {
        const g = String(p?.row?.groupName || '').trim().toLowerCase();
        return String(palletDescByGroup?.[g] || '');
      } },
      { field: 'lineItem', headerName: 'Pallet ID', width: 140, renderCell: (p: any) => String(p?.row?.lineItem || '-') },
      { field: 'palletPrice', headerName: 'Pallet Price', width: 120, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => {
        const gLower = String(p?.row?.groupName || '').trim().toLowerCase();
        const val = Number(groupPriceByName?.[gLower]);
        if (!Number.isFinite(val)) return '';
        return val.toFixed(2);
      } },
      { field: 'selectedWarehouseAvailable', headerName: `${cleanedPrimaryName}`, width: 90, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => String(p?.row?.selectedWarehouseAvailable ?? 0) },
    ];

    // On-Water EDD columns
    for (const edd of manualPickerEddLists.waterEdds) {
      const header = `On-Water ${(() => { const [y,m,d] = String(edd).split('-'); return `${m}/${d}/${y}`; })()}`;
      cols.push({
        field: `ow_${edd}`,
        headerName: header,
        width: 100,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
        sortable: true,
        filterable: false,
        valueGetter: (...args: any[]) => {
          const maybeParams = args?.[0];
          const maybeRow = args?.[1];
          const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
          const list = Array.isArray((row as any)?.onWaterShipments) ? (row as any).onWaterShipments : [];
          const hit = list.find((x: any) => String(x?.edd || '') === edd);
          const v = Number(hit?.qty ?? 0);
          return Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
        },
        renderCell: (p: any) => {
          const list = Array.isArray(p?.row?.onWaterShipments) ? p.row.onWaterShipments : [];
          const hit = list.find((x: any) => String(x?.edd || '') === edd);
          const v = Number(hit?.qty ?? 0);
          const qty = Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
          return String(qty);
        },
      });
    }

    // Second warehouse after On-Water columns
    const wid = String(manualWarehouseId || '').trim();
    const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
    const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
    const secondId = second ? String(second._id) : '';
    const secondName = second ? String((second as any).name || '') : '';
    if (secondId && second) {
      cols.push({
        field: 'secondWarehouseAvailable',
        headerName: `${secondName}`,
        width: 90,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
        sortable: true,
        filterable: false,
        valueGetter: (...args: any[]) => {
          const maybeParams = args?.[0];
          const maybeRow = args?.[1];
          const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
          const per = (row as any)?.perWarehouse || {};
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

    // On-Process EDD columns
    for (const edd of manualPickerEddLists.processEdds) {
      const header = `On-Process ${(() => { const [y,m,d] = String(edd).split('-'); return `${m}/${d}/${y}`; })()}`;
      cols.push({
        field: `op_${edd}`,
        headerName: header,
        width: 100,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
        sortable: true,
        filterable: false,
        valueGetter: (...args: any[]) => {
          const maybeParams = args?.[0];
          const maybeRow = args?.[1];
          const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
          const list = Array.isArray((row as any)?.onProcessBatches) ? (row as any).onProcessBatches : [];
          const hit = list.find((x: any) => String(x?.edd || '') === edd);
          const v = Number(hit?.qty ?? 0);
          return Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
        },
        renderCell: (p: any) => {
          const list = Array.isArray(p?.row?.onProcessBatches) ? p.row.onProcessBatches : [];
          const hit = list.find((x: any) => String(x?.edd || '') === edd);
          const v = Number(hit?.qty ?? 0);
          const qty = Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
          return String(qty);
        },
      });
    }

    cols.push({
      field: 'maxOrder',
      headerName: 'Max Order',
      width: 90,
      type: 'number',
      align: 'right',
      headerAlign: 'right',
      sortable: true,
      valueGetter: (...args: any[]) => {
        const maybeParams = args?.[0];
        const maybeRow = args?.[1];
        const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
        const primary = Number((row as any)?.selectedWarehouseAvailable ?? 0);
        const onWater = Number((row as any)?.onWaterPallets ?? 0);
        const onProcess = Number((row as any)?.onProcessPallets ?? 0);
        const list = Array.isArray(manualPickerWarehouses) ? manualPickerWarehouses : [];
        const wid = String(manualWarehouseId || '').trim();
        const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
        const secondId = second ? String(second._id) : '';
        let secondQty = 0;
        if (secondId) {
          const per = (row as any)?.perWarehouse || {};
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
        const row = (p && typeof p === 'object' && 'row' in p) ? (p as any).row : {};
        if (row && (row as any).isDuplicate) return '-';
        const primary = Number((row as any)?.selectedWarehouseAvailable ?? 0);
        const onWater = Number((row as any)?.onWaterPallets ?? 0);
        const onProcess = Number((row as any)?.onProcessPallets ?? 0);
        let second = 0;
        if (secondWarehouse?._id) {
          const per = (row as any)?.perWarehouse || {};
          const wid = String(secondWarehouse._id);
          second = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
        }
        let baseTotal =
          (Number.isFinite(primary) ? primary : 0) +
          (Number.isFinite(onWater) ? onWater : 0) +
          (Number.isFinite(second) ? second : 0) +
          (Number.isFinite(onProcess) ? onProcess : 0);
        const gLower = String((row as any)?.groupName || '').trim().toLowerCase();
        let otherQty = 0;
        for (const r of (Array.isArray(manualOrderRows) ? manualOrderRows : []) as any[]) {
          const rg = String((r as any)?.groupName || '').trim().toLowerCase();
          const rid = String((r as any)?.id || '');
          if (rg === gLower) {
            const qRaw = (manualOrderQtyByGroup as any)?.[rid] ?? '';
            const q = Math.max(0, Math.floor(Number(qRaw || 0)) || 0);
            otherQty += q;
          }
        }
        const remaining = Math.max(0, Math.floor(baseTotal - otherQty));
        return String(remaining);
      },
    });

    return cols;
  }, [manualWarehouseName, manualPickerEddLists, manualPickerWarehouses, manualWarehouseId, palletNameByGroup, palletDescByGroup, groupPriceByName, manualOrderGroups]);

  // Map the original order's lines to quickly lookup Pallet ID by Pallet Description
  const manualLineItemByGroup = useMemo(() => {
    const m = new Map<string, string>();
    const lines = Array.isArray((manualEditRow as any)?.lines) ? ((manualEditRow as any).lines as any[]) : [];
    for (const l of lines) {
      const g = String(l?.groupName || '').trim().toLowerCase();
      const id = String(l?.lineItem || '').trim();
      if (g && id && !m.has(g)) m.set(g, id);
    }
    // fallback to base mapping captured on open if lines are missing after save
    if (m.size === 0 && manualBaseLineItemByGroupRef.current?.size) {
      return new Map(manualBaseLineItemByGroupRef.current);
    }
    return m;
  }, [manualEditRow]);

  const manualOrderRows = useMemo(() => {
    const keys = Array.isArray(manualOrderGroups) ? manualOrderGroups : [];
    const seenIndexByGroup = new Map<string, number>();

    const rows = keys
      .map((key) => {
        const keyStr = String(key || '');
        const groupNameRaw = keyToGroupName(keyStr);
        const groupNameTrimmed = groupNameRaw.trim();
        const groupKey = groupNameTrimmed.toLowerCase();
        const baseRow: any = manualPickerRowByGroup.get(groupNameTrimmed) || {};

        const persistedUnitGroup = Number((manualUnitPriceByGroup as any)?.[groupNameTrimmed] ?? (manualUnitPriceByGroup as any)?.[groupKey]);
        const persistedUnit = Number((manualUnitPriceByRow as any)?.[keyStr]);
        const priceFromMap = Number(groupPriceByName?.[groupKey]);
        const fallbackPrice = Number(baseRow?.palletPrice ?? baseRow?.price ?? 0);
        const basePrice = Number.isFinite(persistedUnitGroup)
          ? persistedUnitGroup
          : (Number.isFinite(persistedUnit)
            ? persistedUnit
            : (Number.isFinite(priceFromMap) ? priceFromMap : (Number.isFinite(fallbackPrice) ? fallbackPrice : 0)));

        const discountRaw = (manualDiscountByGroup as any)?.[keyStr] ?? '';
        let discountPct = Number(discountRaw || 0);
        if (!Number.isFinite(discountPct)) discountPct = 0;
        discountPct = Math.max(0, Math.min(100, discountPct));

        const discountedPrice = Number.isFinite(basePrice)
          ? basePrice * (1 - discountPct / 100)
          : 0;

        const qtyRaw = (manualOrderQtyByGroup as any)?.[keyStr] ?? '';
        const qtySanitized = Math.max(0, Math.floor(Number(qtyRaw || 0)) || 0);

        const subTotal = Number.isFinite(discountedPrice) ? discountedPrice * qtySanitized : 0;

        const gLower = groupNameTrimmed.toLowerCase();
        const prevIndex = seenIndexByGroup.get(gLower) || 0;
        const instanceIndex = prevIndex + 1;
        seenIndexByGroup.set(gLower, instanceIndex);
        const isDuplicate = instanceIndex > 1;

        return {
          id: keyStr,
          ...baseRow,
          groupName: groupNameRaw,
          unitPrice: Number.isFinite(basePrice) ? basePrice : 0,
          discountedPrice: Number.isFinite(discountedPrice) ? discountedPrice : 0,
          discountPercent: discountPct,
          subTotal: Number.isFinite(subTotal) ? subTotal : 0,
          isDuplicate,
          instanceIndex,
        };
      })
      .filter((r) => String(r.groupName || '').trim());

    // Sort by pallet (group) then by instance index so duplicates are grouped
    rows.sort((a: any, b: any) => {
      const ag = String(a.groupName || '').trim().toLowerCase();
      const bg = String(b.groupName || '').trim().toLowerCase();
      if (ag < bg) return 1 * -1; // keep existing order but group logically
      if (ag > bg) return 1;
      return (a.instanceIndex || 0) - (b.instanceIndex || 0);
    });

    return rows;
  }, [manualOrderGroups, manualPickerRowByGroup, manualDiscountByGroup, manualOrderQtyByGroup, groupPriceByName, manualUnitPriceByRow, manualUnitPriceByGroup]);

  // Sum of all row subtotals for the order pricing summary
  const manualOrderSubTotal = useMemo(() => {
    const rows = Array.isArray(manualOrderRows) ? manualOrderRows : [];
    let sum = 0;
    for (const r of rows as any[]) {
      const v = Number((r as any)?.subTotal ?? 0);
      if (Number.isFinite(v)) sum += v;
    }
    return sum;
  }, [manualOrderRows]);

  const manualOrderColumns = useMemo(() => {
    const cols: any[] = [
      { field: 'palletName', headerName: 'Pallet Name', flex: 1, minWidth: 200, renderCell: (p: any) => {
        const g = String(p?.row?.groupName || '').trim().toLowerCase();
        const name = String(palletNameByGroup?.[g] || '');
        const isDup = Boolean((p?.row as any)?.isDuplicate);
        if (!isDup) return name;
        return `${name} (copy)`;
      } },
      { field: 'palletDescription', headerName: 'Pallet Description', flex: 1, minWidth: 220, renderCell: (p: any) => {
        const g = String(p?.row?.groupName || '').trim().toLowerCase();
        return String(palletDescByGroup?.[g] || '');
      } },
      { field: 'lineItem', headerName: 'Pallet ID', width: 140, renderCell: (p: any) => String(p?.row?.lineItem || '-') },
      {
        field: 'palletPrice',
        headerName: 'Pallet Price',
        width: 110,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
        renderCell: (p: any) => {
          const rowUnit = Number((p?.row as any)?.unitPrice);
          if (Number.isFinite(rowUnit) && rowUnit >= 0) return rowUnit.toFixed(2);
          const gLower = String(p?.row?.groupName || '').trim().toLowerCase();
          const v = Number(groupPriceByName?.[gLower]);
          if (!Number.isFinite(v)) return '';
          return v.toFixed(2);
        },
      },
      {
        field: 'discountPercent',
        headerName: 'Discount (%)',
        width: 120,
        sortable: false,
        filterable: false,
        renderCell: (p: any) => {
          const rowKey = String(p?.row?.id || '');
          const val = manualDiscountByGroup[rowKey] ?? '';
          return (
            <TextField
              type="number"
              size="small"
              value={val}
              disabled={manualFieldsLocked}
              onChange={(e) => {
                const raw = String(e.target.value || '').replace(/[^0-9.\-+eE]/g, '');
                let n = Number(raw || 0);
                if (!Number.isFinite(n)) n = 0;
                n = Math.max(0, Math.min(100, Math.floor(n)));
                setManualDiscountByGroup((prev) => ({ ...prev, [rowKey]: String(n) }));
              }}
              onBlur={(e) => {
                const raw = String(e.target.value || '').trim();
                let n = Number(raw || 0);
                if (!Number.isFinite(n)) n = 0;
                n = Math.max(0, Math.min(100, Math.floor(n)));
                setManualDiscountByGroup((prev) => ({ ...prev, [rowKey]: String(n) }));
              }}
              onKeyDown={(e) => {
                const k = (e as any).key;
                if (k === 'e' || k === 'E' || k === '.' || k === '-' || k === '+' ) {
                  e.preventDefault();
                }
              }}
              inputProps={{ inputMode: 'numeric', pattern: '[0-9]*', min: 0, max: 100, step: 1 }}
              sx={{ width: 90 }}
            />
          );
        },
      },
      {
        field: 'discountedPrice',
        headerName: 'Discounted Price',
        width: 130,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
        renderCell: (p: any) => {
          const v = Number((p && typeof p === 'object' && 'row' in p) ? (p as any).row?.discountedPrice ?? (p as any).value : 0);
          if (!Number.isFinite(v) || v <= 0) return '-';
          return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        },
        valueGetter: (params: any) => {
          const row = (params && typeof params === 'object' && 'row' in params) ? (params as any).row : params;
          const v = Number(row?.discountedPrice ?? 0);
          return Number.isFinite(v) ? v : 0;
        },
      },
      {
        field: 'subTotal',
        headerName: 'Sub Total',
        width: 140,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
        valueGetter: (params: any) => {
          const row = (params && typeof params === 'object' && 'row' in params) ? (params as any).row : params;
          const v = Number(row?.subTotal ?? 0);
          return Number.isFinite(v) ? v : 0;
        },
        renderCell: (p: any) => {
          const v = Number((p && typeof p === 'object' && 'row' in p) ? (p as any).row?.subTotal ?? (p as any).value : 0);
          if (!Number.isFinite(v) || !v) return '-';
          return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        },
      },
    ];

    cols.push({
      field: 'maxQtyOrder',
      headerName: 'Available Qty',
      width: 120,
      type: 'number',
      align: 'right',
      headerAlign: 'right',
      sortable: false,
      filterable: false,
      headerClassName: 'availableQty--header',
      cellClassName: 'availableQty--cell',
      valueGetter: (...args: any[]) => {
        const maybeParams = args?.[0];
        const maybeRow = args?.[1];
        const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
        if (row && (row as any).isDuplicate) return null;
        // Use live picker totals (global post-reservation availability); do NOT subtract this order's quantities
        const primary = Number(row?.selectedWarehouseAvailable ?? 0);
        const onWater = Number(row?.onWaterPallets ?? 0);
        const onProcess = Number(row?.onProcessPallets ?? 0);
        let second = 0;
        if (secondWarehouse?._id) {
          const per = row?.perWarehouse || {};
          const wid = String(secondWarehouse._id);
          second = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
        }
        const liveTotal =
          (Number.isFinite(primary) ? primary : 0) +
          (Number.isFinite(onWater) ? onWater : 0) +
          (Number.isFinite(second) ? second : 0) +
          (Number.isFinite(onProcess) ? onProcess : 0);
        const groupNameTrimmed = String((row as any)?.groupName || '').trim();
        if (liveTotal > 0 || manualPickerRowByGroup.has(groupNameTrimmed)) {
          return Math.max(0, Math.floor(liveTotal));
        }
        // Fallback to saved snapshot only when live picker not ready yet (Edit mode)
        if (manualMode === 'edit') {
          const gLower = String(groupNameTrimmed).toLowerCase();
          const snapTotal = Number((manualBaseTiersByGroup as any)?.[gLower]?.total ?? NaN);
          if (Number.isFinite(snapTotal) && snapTotal >= 0) return Math.floor(snapTotal);
        }
        return null;
      },
      renderCell: (p: any) => {
        const row = (p && typeof p === 'object' && 'row' in p) ? (p as any).row : {};
        if (row && (row as any).isDuplicate) return '-';
        const primary = Number((row as any)?.selectedWarehouseAvailable ?? 0);
        const onWater = Number((row as any)?.onWaterPallets ?? 0);
        const onProcess = Number((row as any)?.onProcessPallets ?? 0);
        let second = 0;
        if (secondWarehouse?._id) {
          const per = (row as any)?.perWarehouse || {};
          const wid = String(secondWarehouse._id);
          second = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
        }
        const liveTotal =
          (Number.isFinite(primary) ? primary : 0) +
          (Number.isFinite(onWater) ? onWater : 0) +
          (Number.isFinite(second) ? second : 0) +
          (Number.isFinite(onProcess) ? onProcess : 0);
        const groupNameTrimmed = String((row as any)?.groupName || '').trim();
        if (liveTotal > 0 || manualPickerRowByGroup.has(groupNameTrimmed)) {
          const tipPrimary = Math.max(0, Math.floor(Number(primary || 0)));
          const tipOnWater = Math.max(0, Math.floor(Number(onWater || 0)));
          const tipSecond = Math.max(0, Math.floor(Number(second || 0)));
          const tipOnProcess = Math.max(0, Math.floor(Number(onProcess || 0)));
          const title = `Available (MPG ${tipPrimary} + On-Water ${tipOnWater} + PEBA ${tipSecond} + On-Process ${tipOnProcess}) = Total ${Math.max(0, Math.floor(liveTotal))}`;
          return (
            <Tooltip title={title} placement="top" arrow>
              <span>{String(Math.max(0, Math.floor(liveTotal)))}</span>
            </Tooltip>
          );
        }
        if (manualMode === 'edit') {
          const gLower = String(groupNameTrimmed).toLowerCase();
          const bt = (manualBaseTiersByGroup as any)?.[gLower];
          const snapTotal = Number(bt?.total ?? NaN);
          if (Number.isFinite(snapTotal) && snapTotal >= 0) {
            const tipPrimary = Math.max(0, Math.floor(Number(bt?.primary || 0)));
            const tipOnWater = Math.max(0, Math.floor(Number(bt?.onWater || 0)));
            const tipSecond = Math.max(0, Math.floor(Number(bt?.second || 0)));
            const tipOnProcess = Math.max(0, Math.floor(Number(bt?.onProcess || 0)));
            const title = `Original Base (MPG ${tipPrimary} + On-Water ${tipOnWater} + PEBA ${tipSecond} + On-Process ${tipOnProcess}) = Total ${Math.floor(snapTotal)}`;
            return (
              <Tooltip title={title} placement="top" arrow>
                <span>{String(Math.floor(snapTotal))}</span>
              </Tooltip>
            );
          }
        }
        return '-';
      },
    });

    cols.push({
      field: 'orderQty',
      headerName: 'Order Qty',
      width: 120,
      sortable: false,
      filterable: false,
      renderCell: (p: any) => {
        const rowKey = String(p?.row?.id || '');
        const groupName = String(p?.row?.groupName || '');
        const val = manualOrderQtyByGroup[rowKey] ?? '';
        const primary = Number(p?.row?.selectedWarehouseAvailable ?? 0);
        const onWater = Number(p?.row?.onWaterPallets ?? 0);
        const onProcess = Number(p?.row?.onProcessPallets ?? 0);
        let second = 0;
        if (secondWarehouse?._id) {
          const per = (p?.row as any)?.perWarehouse || {};
          const wid = String(secondWarehouse._id);
          second = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
        }
        const baseMax = Math.max(0, Math.floor(
          (Number.isFinite(primary) ? primary : 0) +
          (Number.isFinite(onWater) ? onWater : 0) +
          (Number.isFinite(second) ? second : 0) +
          (Number.isFinite(onProcess) ? onProcess : 0)
        ));
        // Remaining available for this row after subtracting quantities of other duplicate rows with the same pallet
        let otherQty = 0;
        for (const r of (Array.isArray(manualOrderRows) ? manualOrderRows : []) as any[]) {
          const rid = String((r as any)?.id || '');
          const rg = String((r as any)?.groupName || '').trim().toLowerCase();
          if (rid !== rowKey && rg === String(groupName).trim().toLowerCase()) {
            const qRaw = (manualOrderQtyByGroup as any)?.[rid] ?? '';
            const q = Math.max(0, Math.floor(Number(qRaw || 0)) || 0);
            otherQty += q;
          }
        }
        const reserved = manualReservedByGroup.get(groupName)?.total || 0;
        const remaining = Math.max(0, baseMax - otherQty);
        const maxAllowed = Math.max(0, Math.floor(remaining + Math.max(0, reserved)));
        return (
          <TextField
            type="number"
            size="small"
            value={val}
            disabled={manualFieldsLocked}
            onChange={(e)=>{
              const raw = sanitizeIntText(e.target.value);
              setManualValidationErrors([]);
              let n = Math.max(1, Math.floor(Number(raw || 0)));
              if (Number.isFinite(maxAllowed)) n = Math.min(n, maxAllowed);
              setManualOrderQtyByGroup((prev)=> ({ ...prev, [rowKey]: String(n) }));
              setManualRecalcTick((t)=> t + 1);
            }}
            onBlur={(e)=>{
              const raw = normalizeIntText(e.target.value);
              let n = Math.max(1, Math.floor(Number(raw || 0)));
              if (Number.isFinite(maxAllowed)) n = Math.min(n, maxAllowed);
              setManualOrderQtyByGroup((prev)=> ({ ...prev, [rowKey]: String(n) }));
              setManualRecalcTick((t)=> t + 1);
            }}
            onKeyDown={(e)=>{
              const k = (e as any).key;
              if (k === 'e' || k === 'E' || k === '.' || k === '-' || k === '+' ) {
                e.preventDefault();
              }
            }}
            inputProps={{ inputMode: 'numeric', pattern: '[0-9]*', min: 1, max: maxAllowed, step: 1 }}
            error={Number(val || 0) > maxAllowed}
            helperText={Number(val || 0) > maxAllowed ? `Max ${maxAllowed}` : undefined}
            sx={{ width: 100 }}
          />
        );
      }
    });

    cols.push({
      field: 'remove',
      headerName: 'Action',
      width: 140,
      sortable: false,
      filterable: false,
      renderCell: (p: any) => {
        const rowKey = String(p?.row?.id || '');
        const groupName = String(p?.row?.groupName || '');
        return (
          <Box sx={{ display: 'flex', gap: 0.5 }}>
            <IconButton
              size="small"
              disabled={(function(){
                if (manualFieldsLocked) return true;
                const primary = Number((p?.row as any)?.selectedWarehouseAvailable ?? 0);
                const onWater = Number((p?.row as any)?.onWaterPallets ?? 0);
                const onProcess = Number((p?.row as any)?.onProcessPallets ?? 0);
                let second = 0;
                if (secondWarehouse?._id) {
                  const per = (p?.row as any)?.perWarehouse || {};
                  const wid = String(secondWarehouse._id);
                  second = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
                }
                // Prefer live picker totals; if not ready, fall back to snapshot in edit mode
                const liveTotal = (
                  (Number.isFinite(primary) ? primary : 0) +
                  (Number.isFinite(onWater) ? onWater : 0) +
                  (Number.isFinite(second) ? second : 0) +
                  (Number.isFinite(onProcess) ? onProcess : 0)
                );
                const groupNameTrimmed = String(groupName || '').trim();
                let baseTotal = 0;
                if (liveTotal > 0 || manualPickerRowByGroup.has(groupNameTrimmed)) {
                  baseTotal = Math.max(0, Math.floor(liveTotal));
                } else if (manualMode === 'edit') {
                  const gLower = groupNameTrimmed.toLowerCase();
                  const snapTotal = Number((manualBaseTiersByGroup as any)?.[gLower]?.total ?? NaN);
                  if (Number.isFinite(snapTotal) && snapTotal >= 0) baseTotal = Math.floor(snapTotal);
                }
                let otherQty = 0;
                const gLower = String(groupName).trim().toLowerCase();
                for (const r of (Array.isArray(manualOrderRows) ? manualOrderRows : []) as any[]) {
                  const rid = String((r as any)?.id || '');
                  const rg = String((r as any)?.groupName || '').trim().toLowerCase();
                  if (rg === gLower && rid !== rowKey) {
                    const qRaw = (manualOrderQtyByGroup as any)?.[rid] ?? '';
                    const q = Math.max(0, Math.floor(Number(qRaw || 0)) || 0);
                    otherQty += q;
                  }
                }
                const reserved = Math.max(0, Number(manualReservedByGroup.get(groupName)?.total || 0));
                const remaining = Math.max(0, baseTotal - otherQty);
                const capacity = Math.max(0, remaining + reserved);
                return capacity <= 0;
              })()}
              onClick={() => {
                // Seed snapshot unit price for this group if missing
                try {
                  const gTrim = String(groupName || '').trim();
                  const gLower = gTrim.toLowerCase();
                  const rowUnit = Number((p?.row as any)?.unitPrice);
                  if (Number.isFinite(rowUnit)) {
                    setManualUnitPriceByGroup((prev)=> (prev && (gLower in prev || gTrim in prev)) ? prev : ({ ...(prev||{}), [gLower]: rowUnit } as any));
                  }
                } catch {}
                const key = makeRowKey(groupName);
                setManualOrderGroups((prev) => ([...(Array.isArray(prev) ? prev : []), key]));
                setManualOrderQtyByGroup((prev) => ({ ...(prev || {}), [key]: '1' }));
                setManualDiscountByGroup((prev) => ({ ...(prev || {}), [key]: '0' }));
                setManualRecalcTick((t)=> t + 1);
              }}
            >
              <ContentCopyIcon fontSize="small" />
            </IconButton>
            <IconButton
              size="small"
              disabled={manualFieldsLocked}
              onClick={() => {
                // Duplicate: ensure snapshot price exists for this group
                try {
                  const gTrim = String(groupName || '').trim();
                  const gLower = gTrim.toLowerCase();
                  const rowUnit = Number((p?.row as any)?.unitPrice);
                  if (Number.isFinite(rowUnit)) {
                    setManualUnitPriceByGroup((prev)=> (prev && (gLower in prev || gTrim in prev)) ? prev : ({ ...(prev||{}), [gLower]: rowUnit } as any));
                  }
                } catch {}
                setManualOrderGroups((prev)=> (Array.isArray(prev) ? prev.filter((x)=> String(x) !== rowKey) : []));
                setManualOrderQtyByGroup((prev)=> {
                  const next = { ...(prev || {}) } as any;
                  delete next[rowKey];
                  return next;
                });
                setManualDiscountByGroup((prev) => {
                  const next = { ...(prev || {}) } as any;
                  delete next[rowKey];
                  return next;
                });
                // If this was the last instance of the group, clear snapshots and update cache
                try {
                  const gTrim = String(groupName || '').trim();
                  const gLower = gTrim.toLowerCase();
                  const remaining = (Array.isArray(manualOrderGroups) ? manualOrderGroups : [])
                    .filter((k) => String(k) !== rowKey && keyToGroupName(String(k)).trim().toLowerCase() === gLower).length;
                  if (remaining <= 0) {
                    setManualUnitPriceByGroup((prev) => { const next = { ...(prev||{}) } as any; delete next[gLower]; delete next[gTrim]; return next; });
                    setManualItemsByGroup((prev) => { const next = { ...(prev||{}) } as any; delete next[gLower]; delete next[gTrim]; return next; });
                    const orderKey = String((manualEditRow as any)?.rawId || (manualEditRow as any)?.id || '').trim();
                    if (orderKey) {
                      const cacheKey = `orderLineSnapshots:${orderKey}`;
                      try {
                        const cached = localStorage.getItem(cacheKey);
                        if (cached) {
                          const parsed = JSON.parse(cached) || {};
                          if (parsed && typeof parsed === 'object') {
                            if (parsed.unit && typeof parsed.unit === 'object') { delete parsed.unit[gLower]; delete parsed.unit[gTrim]; }
                            if (parsed.items && typeof parsed.items === 'object') { delete parsed.items[gLower]; delete parsed.items[gTrim]; }
                            localStorage.setItem(cacheKey, JSON.stringify(parsed));
                          }
                        }
                      } catch {}
                    }
                  }
                } catch {}
                setManualRecalcTick((t)=> t + 1);
              }}
            >
              <DeleteIcon fontSize="small" />
            </IconButton>
          </Box>
        );
      }
    });

    // ... existing code ...

    return cols;
  }, [
    manualWarehouseName,
    manualWarehouseId,
    secondWarehouse,
    manualMode,
    manualReservedBreakdown,
    manualOrderQtyByGroup,
    manualDiscountByGroup,
    manualFieldsLocked,
    manualRecalcTick,
    openOnWaterDetails,
    openOnProcessDetails,
    groupPriceByName,
    manualOrderRows,
    manualBaseTiersByGroup,
  ]);

  const manualAllocationRows = useMemo(() => {
    const allocs = Array.isArray(manualAllocations) ? manualAllocations : [];
    const byGroup = new Map<string, { groupName: string; primary: number; onWater: number; onProcess: number; second: number }>();

    for (const a of allocs) {
      const g = String((a as any)?.groupName || '').trim();
      const src = String((a as any)?.source || '').trim().toLowerCase();
      const qty = Math.floor(Number((a as any)?.qty || 0));
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

  // Quick lookup of reserved-by-this-order per group
  const manualReservedByGroup = useMemo(() => {
    const map = new Map<string, { primary: number; onWater: number; onProcess: number; second: number; total: number }>();
    const rows = Array.isArray(manualAllocations) ? manualAllocations : [];
    for (const a of rows) {
      const g = String((a as any)?.groupName || '').trim();
      const src = String((a as any)?.source || '').trim().toLowerCase();
      const qty = Math.floor(Number((a as any)?.qty || 0));
      if (!g || !Number.isFinite(qty) || qty <= 0) continue;
      const cur = map.get(g) || { primary: 0, onWater: 0, onProcess: 0, second: 0, total: 0 };
      if (src === 'primary') cur.primary += qty;
      else if (src === 'on_water') cur.onWater += qty;
      else if (src === 'on_process') cur.onProcess += qty;
      else if (src === 'second') cur.second += qty;
      cur.total = cur.primary + cur.onWater + cur.onProcess + cur.second;
      map.set(g, cur);
    }
    return map;
  }, [manualAllocations]);

  // Validation: detect any over-ordered qty beyond total available across tiers
  const manualHasOverOrder = useMemo(() => {
    const qtyBy = manualOrderQtyByGroup || {};
    const byGroup = new Map<string, any>();
    for (const r of (Array.isArray(manualPickerRows) ? manualPickerRows : [])) {
      const g = String((r as any)?.groupName || '').trim();
      if (!g) continue;
      byGroup.set(g, r);
    }
    for (const [gName, qtyStr] of Object.entries(qtyBy)) {
      const g = String(gName || '').trim();
      const need = Math.max(0, Math.floor(Number(qtyStr || 0)));
      if (!g || !Number.isFinite(need) || need <= 0) continue;
      const r: any = byGroup.get(g) || {};
      const primary = Number(r?.selectedWarehouseAvailable ?? 0);
      const onWater = Number(r?.onWaterPallets ?? 0);
      const onProcess = Number(r?.onProcessPallets ?? 0);
      let second = 0;
      if (secondWarehouse?._id) {
        const per = r?.perWarehouse || {};
        const wid = String(secondWarehouse._id);
        second = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
      }
      const baseMax = Math.max(0, Math.floor(
        (Number.isFinite(primary) ? primary : 0) +
        (Number.isFinite(onWater) ? onWater : 0) +
        (Number.isFinite(second) ? second : 0) +
        (Number.isFinite(onProcess) ? onProcess : 0)
      ));
      const reserved = manualReservedByGroup.get(g)?.total || 0;
      const maxAllowed = baseMax + Math.max(0, reserved);
      if (need > maxAllowed) return true;
    }
    return false;
  }, [manualOrderQtyByGroup, manualPickerRows, secondWarehouse, manualReservedByGroup]);

  

  const manualReservedRows = useMemo(() => {
    const rows = Array.isArray(manualReservedBreakdown) ? manualReservedBreakdown : [];
    if (rows.length) {
      return rows
        .map((r: any) => {
          const groupName = String(r?.groupName || '');
          const picker = manualPickerRowByGroup.get(groupName) || {};
          const gLower = String(groupName || '').trim().toLowerCase();
          const hit = (Array.isArray(manualPickerRows) ? manualPickerRows : []).find(
            (p: any) => String(p?.groupName || '').trim().toLowerCase() === gLower
          );
          const fromOrder = (Array.isArray(manualOrderRows) ? manualOrderRows : []).find(
            (x: any) => String(x?.groupName || '').trim().toLowerCase() === gLower
          );
          const fromBase = manualLineItemByGroup.get(gLower) || manualBaseLineItemByGroupRef.current.get(gLower) || '';
          const palletId = String(picker?.lineItem || hit?.lineItem || fromOrder?.lineItem || fromBase || '');
          return {
            id: String(r?.id || r?.groupName || ''),
            palletId,
            groupName,
            primary: Math.floor(Number(r?.primary || 0)),
            onWater: Math.floor(Number(r?.onWater || 0)),
            second: Math.floor(Number(r?.second || 0)),
            onProcess: Math.floor(Number(r?.onProcess || 0)),
          };
        })
        .filter((r: any) => String(r.groupName || '').trim())
        .sort((a: any, b: any) => String(a.groupName).localeCompare(String(b.groupName)));
    }
    return manualAllocationRows;
  }, [manualReservedBreakdown, manualAllocationRows, manualPickerRowByGroup, manualPickerRows, manualOrderRows, manualLineItemByGroup]);

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

  const exportManualOrderXlsx = useCallback(async () => {
    if (manualMode !== 'edit' || !manualEditRow) return;
    const orderNumber = String(manualEditRow?.orderNumber || manualEditRow?.rawId || manualEditRow?.id || '').trim();
    const sStatus = normalizeStatus(manualStatus || '');

    const sub = Number(manualOrderSubTotal);
    const sp = Number(manualShippingPercent);
    const ship = Number.isFinite(sp) ? Math.min(100, Math.max(0, sp)) : 0;
    const grand = Number.isFinite(sub) ? sub * (1 + ship / 100) : NaN;

    const pad = (n: number) => n.toString().padStart(2, '0');
    const d = new Date();
    const ts = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;

    const safe = (v: any) => String(v ?? '').replace(/[/\\:*?"<>|]/g, '-');
    const fname = `order-details-${safe(orderNumber || 'order')}-(${safe(sStatus || 'status')})-(${ts}).xlsx`;

    const rows: any[] = [];

    // Header block (A1..E1 and A2..E2)
    // A1: ORDER NUMBER value, B1: 'To:', C1: [CUSTOMER NAME], D1: 'Phone:', E1: [PHONE NUMBER]
    rows.push([orderNumber, 'To:', manualCustomerName, 'Phone:', manualCustomerPhone]);
    // A2: empty, B2: 'Address:', C2: [SHIPPING ADDRESS], D2: 'Email:', E2: [EMAIL ADDRESS]
    rows.push(['', 'Address:', manualShippingAddress, 'Email:', manualCustomerEmail]);

    rows.push([]);
    rows.push(['Status:', (sStatus || '').toUpperCase()]);
    rows.push(['Estimated Shipdate for Customer:', manualEstFulfillment]);
    rows.push(['Estimated Arrival Date:', manualEstDelivered]);

    // Helper: resolve pallet metadata and items for a group
    const pickerRows = Array.isArray(manualPickerRows) ? manualPickerRows : [];
    const getPalletMeta = (g: string) => {
      const byMap = manualPickerRowByGroup?.get ? manualPickerRowByGroup.get(g) : null;
      const gLower = g.toLowerCase();
      const hit = pickerRows.find((p: any) => String(p?.groupName || '').trim().toLowerCase() === gLower);
      const palletId = String(byMap?.lineItem || hit?.lineItem || '').trim();
      const palletName = String(byMap?.palletName || hit?.palletName || palletNameByGroup?.[gLower] || '').trim();
      return { palletId, palletName };
    };

    const itemsCache = new Map<string, any[]>();
    const getGroupItems = async (g: string) => {
      const k = g.toLowerCase();
      if (itemsCache.has(k)) return itemsCache.get(k)!;
      // Prefer snapshots: in-memory -> saved order lines -> localStorage -> API
      const snap = (manualItemsByGroup as any)?.[k];
      if (Array.isArray(snap) && snap.length) { itemsCache.set(k, snap); return snap; }
      try {
        const lines = Array.isArray((manualEditRow as any)?.lines) ? (manualEditRow as any).lines : [];
        const hit = lines.find((l: any) => String(l?.groupName || l?.lineItem || '').trim().toLowerCase() === k);
        if (hit && Array.isArray((hit as any)?.itemsAtAdd) && (hit as any).itemsAtAdd.length) {
          const list = (hit as any).itemsAtAdd as any[];
          itemsCache.set(k, list);
          return list;
        }
      } catch {}
      try {
        const orderKey = String((manualEditRow as any)?.rawId || (manualEditRow as any)?.id || '').trim();
        if (orderKey) {
          const cacheKey = `orderLineSnapshots:${orderKey}`;
          const cached = localStorage.getItem(cacheKey);
          if (cached) {
            const parsed = JSON.parse(cached);
            const list = parsed && parsed.items ? (parsed.items[k] || parsed.items[g]) : null;
            if (Array.isArray(list) && list.length) { itemsCache.set(k, list); return list; }
          }
        }
      } catch {}
      try {
        const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
        const list = Array.isArray((data as any)?.items) ? (data as any).items : [];
        itemsCache.set(k, list);
        return list;
      } catch {
        itemsCache.set(k, []);
        return [] as any[];
      }
    };

    // Per-line pallet blocks (do not merge duplicates so per-line discounts/qty are preserved)
    const rowsList = Array.isArray(manualOrderRows) ? manualOrderRows : [];
    const printedGroups = new Set<string>();
    let prevGroup: string | null = null;
    for (const r of rowsList as any[]) {
      const g = String(r?.groupName || '').trim();
      if (!g) continue;
      if (prevGroup && prevGroup !== g) {
        // separator between different pallet groups
        rows.push([]);
      }
      const rowId = String((r as any)?.id || '').trim();
      const qtyText = (manualOrderQtyByGroup as any)?.[rowId] ?? (manualOrderQtyByGroup as any)?.[g] ?? (r as any)?.orderQty ?? (r as any)?.qty;
      const qtyNum = Number(qtyText);
      const qty = Number.isFinite(qtyNum) ? Math.floor(qtyNum) : NaN;
      const discount = Math.min(100, Math.max(0, Number((r as any)?.discountPercent ?? 0)));
      const gLower = g.toLowerCase();
      const unitSnap = (manualUnitPriceByGroup as any)?.[gLower];
      const unit = Number.isFinite(Number(unitSnap)) ? Number(unitSnap) : Number((r as any)?.unitPrice ?? (groupPriceByName as any)?.[gLower] ?? 0);
      const discounted = Number.isFinite(unit) ? unit * (1 - discount / 100) : NaN;
      const lineSub = (Number.isFinite(discounted) && Number.isFinite(qty)) ? discounted * qty : NaN;

      const { palletId, palletName } = getPalletMeta(g);
      const palletDesc = String(palletDescByGroup?.[g.toLowerCase()] || g || '').trim();
      if (!printedGroups.has(g)) {
        // First occurrence of this group: print meta and items list once
        rows.push([palletName, palletDesc, palletId, '', '', '']);
        const items = await getGroupItems(g);
        if (items.length) {
          for (const it of items) {
            const code = String((it as any)?.itemCode || '');
            const desc = String((it as any)?.description || '');
            const upc = String((it as any)?.upc || '');
            const pack = (it as any)?.packSize;
            const price = Number((it as any)?.price);
            rows.push([
              '',
              code,
              desc,
              upc,
              Number.isFinite(Number(pack)) ? Number(pack) : '',
              Number.isFinite(price) ? price : '',
            ]);
          }
        } else {
          // No items available; still output a placeholder row with pallet meta
          rows.push([palletName, palletDesc, palletId, '', '', '']);
        }
        printedGroups.add(g);
      }

      // Summary for this pallet line in one row per spec
      rows.push([
        `PALLET PRICE: ${Number.isFinite(unit) ? `$${unit.toFixed(2)}` : ''}`,
        `DISCOUNT: ${Number.isFinite(Number(discount)) ? `${discount}%` : ''}`,
        `DISCOUNTED PRICE: ${Number.isFinite(discounted) ? `$${discounted.toFixed(2)}` : ''}`,
        `QUANTITY: ${Number.isFinite(qty) ? qty : ''}`,
        'SUB TOTAL',
        Number.isFinite(lineSub) ? lineSub : '',
      ]);
      prevGroup = g;
    }

    rows.push([]);
    rows.push(['SUB TOTAL', Number.isFinite(sub) ? sub : '']);
    rows.push(['SHIPPING', Number.isFinite(sp) ? `${sp}%` : '']);
    rows.push(['GRAND TOTAL', Number.isFinite(grand) ? grand : '']);

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // Column widths for readability
    (ws as any)['!cols'] = [
      { wch: 22 }, // A
      { wch: 18 }, // B
      { wch: 28 }, // C
      { wch: 16 }, // D
      { wch: 14 }, // E
      { wch: 16 }, // F
    ];

    const enc = XLSX.utils.encode_cell;
    const get = (r: number, c: number) => (ws as any)[enc({ r, c })];
    const ensure = (r: number, c: number) => {
      const addr = enc({ r, c });
      (ws as any)[addr] = (ws as any)[addr] || { t: 's', v: '' };
      return (ws as any)[addr];
    };
    const style = (r: number, c: number, s: any) => {
      const cell = ensure(r, c);
      cell.s = { ...(cell.s || {}), ...(s || {}) };
    };
    const fill = (rgb: string) => ({ fill: { patternType: 'solid', fgColor: { rgb } } });
    const bold = { font: { bold: true } };
    const border = {
      border: {
        top: { style: 'thin', color: { rgb: 'FFCCCCCC' } },
        bottom: { style: 'thin', color: { rgb: 'FFCCCCCC' } },
        left: { style: 'thin', color: { rgb: 'FFCCCCCC' } },
        right: { style: 'thin', color: { rgb: 'FFCCCCCC' } },
      },
    } as any;

    // Header rows (0-based indexing): rows 0 and 1
    for (let c = 0; c <= 4; c++) {
      style(0, c, { ...bold });
      style(1, c, {});
    }
    // Make order number a bit larger and underlined bottom
    style(0, 0, { font: { bold: true, sz: 12 }, border: { bottom: { style: 'thin', color: { rgb: 'FFAAAAAA' } } } });
    // Emphasize labels in row 0 col 1 and 3, row 1 col 1 and 3
    style(0, 1, { ...bold });
    style(0, 3, { ...bold });
    style(1, 1, { ...bold });
    style(1, 3, { ...bold });
    // Apply same light-blue highlight as totals to ORDER# (A1), To (B1), Address (B2), Phone (D1), Email (D2)
    const blue = fill('FFDDEBF7');
    style(0, 0, { ...blue, ...border }); // ORDER# value cell highlighted
    style(0, 1, { ...blue, ...border }); // To:
    style(1, 1, { ...blue, ...border }); // Address:
    style(0, 3, { ...blue, ...border }); // Phone:
    style(1, 3, { ...blue, ...border }); // Email:

    // Status and date labels (rows 3..5 in column A)
    style(3, 0, { ...bold });
    style(4, 0, { ...bold });
    style(5, 0, { ...bold });

    // Detect and style pallet meta rows and summary rows; also zebra-strip item rows and group borders
    const rowCount = rows.length;
    let groupStart: number | null = null;
    let inItems = false;
    let zebra = false;
    for (let r = 0; r < rowCount; r++) {
      const a = get(r, 0)?.v;
      const b = get(r, 1)?.v;
      const c = get(r, 2)?.v;
      const d = get(r, 3)?.v;
      const e = get(r, 4)?.v;
      const f = get(r, 5)?.v;
      const aStr = typeof a === 'string' ? a : '';

      // Pallet meta row heuristic: A,B,C non-empty and D,E,F empty
      if (a && b && c && !d && !e && !f) {
        // Close previous group box if any
        if (groupStart !== null && groupStart < r) {
          for (let rr = groupStart; rr < r; rr++) {
            for (let cc = 0; cc <= 5; cc++) style(rr, cc, { ...border });
          }
        }
        groupStart = r;
        inItems = true;
        zebra = false;
        for (let col = 0; col <= 5; col++) style(r, col, { ...fill('FFE2EFDA'), ...border }); // light green meta
        continue;
      }

      // Summary row starts with 'PALLET PRICE:'
      if (aStr && aStr.toUpperCase().startsWith('PALLET PRICE')) {
        for (let col = 0; col <= 5; col++) style(r, col, { ...fill('FFF2F2F2'), ...border }); // light grey
        style(r, 0, { ...bold });
        style(r, 1, { ...bold });
        style(r, 2, { ...bold });
        style(r, 3, { ...bold });
        style(r, 4, { ...bold, ...blue }); // highlight per-pallet SUB TOTAL label
        style(r, 5, { ...blue, alignment: { horizontal: 'right' }, numFmt: '"$"#,##0.00' }); // highlight value
        // Close current group box at summary row (inclusive)
        if (groupStart !== null) {
          for (let rr = groupStart; rr <= r; rr++) {
            for (let cc = 0; cc <= 5; cc++) style(rr, cc, { ...border });
          }
          groupStart = null;
        }
        inItems = false;
        continue;
      }

      // Item rows: A empty and any of B-F non-empty
      const isItem = !a && (b || c || d || e || f);
      if (inItems && isItem) {
        zebra = !zebra;
        if (zebra) {
          for (let col = 0; col <= 5; col++) style(r, col, { ...fill('FFF9F9F9') }); // very light stripe
        }
      }
    }
    // Ensure final group box closes if file ended without explicit summary (edge case)
    if (groupStart !== null) {
      for (let rr = groupStart; rr < rowCount; rr++) {
        for (let cc = 0; cc <= 5; cc++) style(rr, cc, { ...border });
      }
    }

    // Totals section: find labeled rows (SUB TOTAL, SHIPPING, GRAND TOTAL)
    let totalsStart = -1;
    let totalsEnd = -1;
    for (let r = rowCount - 1; r >= 0 && r >= rowCount - 10; r--) {
      const a = get(r, 0)?.v;
      const aStr = typeof a === 'string' ? a.toUpperCase() : '';
      if (aStr === 'GRAND TOTAL' && totalsEnd === -1) totalsEnd = r;
      if (aStr === 'SUB TOTAL') totalsStart = r;
      if (aStr === 'SUB TOTAL' || aStr === 'SHIPPING' || aStr === 'GRAND TOTAL') {
        style(r, 0, { ...bold, ...fill('FFDDEBF7'), ...border }); // light blue label
        style(r, 1, { ...border });
      }
    }
    if (totalsStart !== -1 && totalsEnd !== -1 && totalsEnd >= totalsStart) {
      // Apply thin border around the entire totals block A..F
      for (let r = totalsStart; r <= totalsEnd; r++) {
        for (let c = 0; c <= 5; c++) {
          style(r, c, { ...border });
        }
      }
      // Emphasize GRAND TOTAL value (row totalsEnd assumed to be GRAND TOTAL or below it)
      style(totalsEnd, 1, { font: { bold: true, sz: 12 } });
      // Apply currency numFmt to SUB TOTAL and GRAND TOTAL values in column B
      for (let r = totalsStart; r <= totalsEnd; r++) {
        const a = get(r, 0)?.v;
        const aStr = typeof a === 'string' ? a.toUpperCase() : '';
        if (aStr === 'SUB TOTAL' || aStr === 'GRAND TOTAL') {
          style(r, 1, { alignment: { horizontal: 'right' }, numFmt: '"$"#,##0.00' });
        }
      }
    }

    // Align numeric columns and set basic number formats
    for (let r = 0; r < rowCount; r++) {
      // D (UPC column) keep general; E (PACK SIZE) right; F (ITEM PRICE/SUBTOTAL) currency-like
      ensure(r, 4); ensure(r, 5);
      style(r, 4, { alignment: { horizontal: 'right' } });
      style(r, 5, { alignment: { horizontal: 'right' }, numFmt: '"$"#,##0.00' });
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Order Details');
    XLSX.writeFile(wb, fname);
  }, [
    manualMode,
    manualEditRow,
    manualStatus,
    manualCustomerEmail,
    manualCustomerName,
    manualCustomerPhone,
    manualShippingAddress,
    manualEstFulfillment,
    manualEstDelivered,
    manualOrderSubTotal,
    manualShippingPercent,
    manualOrderRows,
    manualOrderQtyByGroup,
    manualPickerRows,
    manualPickerRowByGroup,
    palletNameByGroup,
    groupPriceByName,
    manualItemsByGroup,
    manualUnitPriceByGroup,
  ]);

  const isValidEmail = (email: string) => {
    const s = String(email || '').trim();
    if (!s) return false;
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
  };

  const exportManualOrderPdf = useCallback(async () => {
    if (manualMode !== 'edit' || !manualEditRow) return;
    const orderNumber = String(manualEditRow?.orderNumber || manualEditRow?.rawId || manualEditRow?.id || '').trim();
    const sStatus = String(normalizeStatus(manualStatus || '') || '').toUpperCase();

    const sub = Number(manualOrderSubTotal);
    const sp = Number(manualShippingPercent);
    const ship = Number.isFinite(sp) ? Math.min(100, Math.max(0, sp)) : 0;
    const grand = Number.isFinite(sub) ? sub * (1 + ship / 100) : NaN;

    const blue = '#DDEBF7';
    const green = '#E2EFDA';
    const gray = '#F2F2F2';

    const pickerRows = Array.isArray(manualPickerRows) ? manualPickerRows : [];
    const getPalletMeta = (g: string) => {
      const byMap = (manualPickerRowByGroup as any)?.get ? (manualPickerRowByGroup as any).get(g) : null;
      const gLower = g.toLowerCase();
      const hit = pickerRows.find((p: any) => String(p?.groupName || '').trim().toLowerCase() === gLower);
      const palletId = String(byMap?.lineItem || hit?.lineItem || '').trim();
      const palletName = String(byMap?.palletName || hit?.palletName || (palletNameByGroup as any)?.[gLower] || '').trim();
      const palletDesc = String((palletDescByGroup as any)?.[gLower] || g || '').trim();
      return { palletId, palletName, palletDesc };
    };
    const itemsCache = new Map<string, any[]>();
    const getGroupItems = async (g: string) => {
      const k = g.toLowerCase();
      if (itemsCache.has(k)) return itemsCache.get(k)!;
      // Prefer snapshots: in-memory -> saved order lines -> localStorage -> API
      const snap = (manualItemsByGroup as any)?.[k];
      if (Array.isArray(snap) && snap.length) { itemsCache.set(k, snap); return snap; }
      try {
        const lines = Array.isArray((manualEditRow as any)?.lines) ? (manualEditRow as any).lines : [];
        const hit = lines.find((l: any) => String(l?.groupName || l?.lineItem || '').trim().toLowerCase() === k);
        if (hit && Array.isArray((hit as any)?.itemsAtAdd) && (hit as any).itemsAtAdd.length) {
          const list = (hit as any).itemsAtAdd as any[];
          itemsCache.set(k, list);
          return list;
        }
      } catch {}
      try {
        const orderKey = String((manualEditRow as any)?.rawId || (manualEditRow as any)?.id || '').trim();
        if (orderKey) {
          const cacheKey = `orderLineSnapshots:${orderKey}`;
          const cached = localStorage.getItem(cacheKey);
          if (cached) {
            const parsed = JSON.parse(cached);
            const list = parsed && parsed.items ? (parsed.items[k] || parsed.items[g]) : null;
            if (Array.isArray(list) && list.length) { itemsCache.set(k, list); return list; }
          }
        }
      } catch {}
      try {
        const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
        const list = Array.isArray((data as any)?.items) ? (data as any).items : [];
        itemsCache.set(k, list);
        return list;
      } catch {
        itemsCache.set(k, []);
        return [] as any[];
      }
    };

    const content: any[] = [];
    content.push({
      table: {
        widths: [160, 40, '*', 55, 120],
        body: [
          [
            { text: orderNumber, fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: 'To:', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: manualCustomerName || '', border: [true, true, true, true] },
            { text: 'Phone:', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: manualCustomerPhone || '', border: [true, true, true, true] },
          ],
          [
            { text: '', border: [true, false, true, true] },
            { text: 'Address:', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: manualShippingAddress || '', border: [true, true, true, true] },
            { text: 'Email:', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: manualCustomerEmail || '', border: [true, true, true, true] },
          ],
        ],
      },
      layout: 'noHorizontalLines',
      margin: [0, 0, 0, 8],
    });

    content.push({ text: 'Status:', bold: true, margin: [0, 2, 0, 0] });
    content.push({ text: sStatus, margin: [0, 0, 0, 2] });
    content.push({ text: 'Estimated Shipdate for Customer:', bold: true });
    content.push({ text: manualEstFulfillment || '', margin: [0, 0, 0, 2] });
    content.push({ text: 'Estimated Arrival Date:', bold: true });
    content.push({ text: manualEstDelivered || '', margin: [0, 0, 0, 8] });

    const rowsList = Array.isArray(manualOrderRows) ? manualOrderRows : [];
    const printedGroups = new Set<string>();
    let prevGroup: string | null = null;
    for (const r of rowsList as any[]) {
      const g = String(r?.groupName || '').trim();
      if (!g) continue;
      if (prevGroup && prevGroup !== g) content.push({ text: '\n' });

      const rowId = String((r as any)?.id || '').trim();
      const qtyText = (manualOrderQtyByGroup as any)?.[rowId] ?? (manualOrderQtyByGroup as any)?.[g] ?? (r as any)?.orderQty ?? (r as any)?.qty;
      const qtyNum = Number(qtyText);
      const qty = Number.isFinite(qtyNum) ? Math.floor(qtyNum) : NaN;
      const discount = Math.min(100, Math.max(0, Number((r as any)?.discountPercent ?? 0)));
      const gLower = g.toLowerCase();
      const unitSnap = (manualUnitPriceByGroup as any)?.[gLower];
      const unit = Number.isFinite(Number(unitSnap)) ? Number(unitSnap) : Number((r as any)?.unitPrice ?? (groupPriceByName as any)?.[gLower] ?? 0);
      const discounted = Number.isFinite(unit) ? unit * (1 - discount / 100) : NaN;
      const lineSub = (Number.isFinite(discounted) && Number.isFinite(qty)) ? discounted * qty : NaN;

      const { palletId, palletName, palletDesc } = getPalletMeta(g);
      if (!printedGroups.has(g)) {
        content.push({
          table: {
            widths: ['*', '*', 120, 60, 60, 80],
            body: [[
              { text: palletName, fillColor: green, border: [true, true, true, true] },
              { text: palletDesc, fillColor: green, border: [true, true, true, true] },
              { text: palletId, fillColor: green, border: [true, true, true, true] },
              { text: '', fillColor: green, border: [true, true, true, true] },
              { text: '', fillColor: green, border: [true, true, true, true] },
              { text: '', fillColor: green, border: [true, true, true, true] },
            ]],
          },
          layout: 'noHorizontalLines',
        });

        const items = await getGroupItems(g);
        if (items.length) {
          const itemRows = items.map((it: any) => [
            { text: '', border: [true, true, true, true] },
            { text: String(it?.itemCode || ''), border: [true, true, true, true] },
            { text: String(it?.description || ''), border: [true, true, true, true] },
            { text: String(it?.upc || ''), alignment: 'right', border: [true, true, true, true] },
            { text: (Number.isFinite(Number(it?.packSize)) ? String(Number(it?.packSize)) : ''), alignment: 'right', border: [true, true, true, true] },
            { text: (Number.isFinite(Number(it?.price)) ? `$${Number(it?.price).toFixed(2)}` : ''), alignment: 'right', border: [true, true, true, true] },
          ]);
          content.push({ table: { widths: ['*', 90, '*', 70, 60, 70], body: itemRows }, layout: 'noHorizontalLines' });
        }
        printedGroups.add(g);
      }

      content.push({
        table: {
          widths: ['*', 90, 150, 90, 90, 80],
          body: [[
            { text: `PALLET PRICE: ${Number.isFinite(unit) ? `$${unit.toFixed(2)}` : ''}`, fillColor: gray, bold: true, border: [true, true, true, true] },
            { text: `DISCOUNT: ${Number.isFinite(Number(discount)) ? `${discount}%` : ''}`, fillColor: gray, bold: true, border: [true, true, true, true] },
            { text: `DISCOUNTED PRICE: ${Number.isFinite(discounted) ? `$${discounted.toFixed(2)}` : ''}`, fillColor: gray, bold: true, border: [true, true, true, true] },
            { text: `QUANTITY: ${Number.isFinite(qty) ? qty : ''}`, fillColor: gray, bold: true, border: [true, true, true, true] },
            { text: 'SUB TOTAL', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: Number.isFinite(lineSub) ? `$${lineSub.toFixed(2)}` : '', alignment: 'right', fillColor: blue, border: [true, true, true, true] },
          ]],
        },
        layout: 'noHorizontalLines',
        margin: [0, 2, 0, 4],
      });

      prevGroup = g;
    }

    content.push({
      table: {
        widths: [120, 120],
        body: [
          [ { text: 'SUB TOTAL', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: Number.isFinite(sub) ? `$${sub.toFixed(2)}` : '', alignment: 'right', border: [true, true, true, true] } ],
          [ { text: 'SHIPPING', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: Number.isFinite(sp) ? `${sp}%` : '', alignment: 'right', border: [true, true, true, true] } ],
          [ { text: 'GRAND TOTAL', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: Number.isFinite(grand) ? `$${grand.toFixed(2)}` : '', alignment: 'right', bold: true, border: [true, true, true, true] } ],
        ],
      },
      layout: 'noHorizontalLines',
      margin: [0, 6, 0, 0],
    });

    const docDefinition = {
      pageOrientation: 'landscape',
      pageMargins: [20, 20, 20, 20],
      content,
      defaultStyle: { fontSize: 9 },
    } as any;

    const safe = (v: any) => String(v ?? '').replace(/[\/\\:*?"<>|]/g, '-');
    const fname = `order-details-${safe(orderNumber || 'order')}-(${safe(sStatus || 'status')}).pdf`;
    (pdfMake as any).createPdf(docDefinition).download(fname);
  }, [
    manualMode,
    manualEditRow,
    manualStatus,
    manualCustomerName,
    manualCustomerPhone,
    manualShippingAddress,
    manualCustomerEmail,
    manualEstFulfillment,
    manualEstDelivered,
    manualOrderSubTotal,
    manualShippingPercent,
    manualOrderRows,
    manualOrderQtyByGroup,
    manualPickerRows,
    manualPickerRowByGroup,
    palletNameByGroup,
    palletDescByGroup,
    groupPriceByName,
    manualItemsByGroup,
    manualUnitPriceByGroup,
  ]);

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

  const normalizePriceText = (v: any) => {
    const s = String(v ?? '').trim();
    if (!s) return '';
    const cleaned = s.replace(/[^0-9.]/g, '');
    if (!cleaned) return '';
    const n = Number(cleaned);
    if (!Number.isFinite(n) || n < 0) return '';
    return String(n);
  };

  const normalizeDiscountText = (v: any) => {
    const s = String(v ?? '').trim();
    if (!s) return '';
    const cleaned = s.replace(/[^0-9.]/g, '');
    if (!cleaned) return '';
    const n = Number(cleaned);
    if (!Number.isFinite(n)) return '';
    const clamped = Math.min(100, Math.max(0, n));
    return String(clamped);
  };

  const manualFinalPrice = useMemo(() => {
    const sub = Number(manualOrderSubTotal);
    if (!Number.isFinite(sub)) return '';
    const sp = Number(manualShippingPercent);
    const ship = Number.isFinite(sp) ? Math.min(100, Math.max(0, sp)) : 0;
    const out = sub * (1 + ship / 100);
    if (!Number.isFinite(out)) return '';
    return out.toFixed(2);
  }, [manualOrderSubTotal, manualShippingPercent]);

  const getActiveManualOptions = () => (manualLineMode === 'pallet_group' ? manualPalletGroupOptions : manualLineItemOptions);

  const suggestShipdateForSelection = useCallback(() => {
    const today = todayYmd;
    const rowsByGroup = new Map((manualPickerRows || []).map((r: any) => [String(r.groupName || ''), r]));
    const toYmd = (s: any) => {
      const v = String(s || '');
      const slice = v.slice(0, 10);
      return /^\d{4}-\d{2}-\d{2}$/.test(slice) ? slice : '';
    };

    const onWaterCompletionYmd = (row: any, neededQty: number) => {
      let remaining = Math.max(0, Math.floor(Number(neededQty || 0)));
      if (!Number.isFinite(remaining) || remaining <= 0) return '';

      const ships = Array.isArray(row?.onWaterShipments) ? row.onWaterShipments : [];
      const normalized = ships
        .map((x: any) => ({ edd: toYmd(x?.edd), qty: Math.floor(Number(x?.qty || 0)) }))
        .filter((x: any) => x.edd && Number.isFinite(x.qty) && x.qty > 0)
        .sort((a: any, b: any) => String(a.edd).localeCompare(String(b.edd)));
      for (const s of normalized) {
        if (remaining <= 0) break;
        remaining -= Math.min(remaining, Math.max(0, s.qty));
        if (remaining <= 0) return s.edd;
      }

      // Backward compatibility fallback: if we don't have per-shipment rows, use the aggregated EDD.
      return toYmd(row?.onWaterEdd);
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
        const onWaterReady = onWaterCompletionYmd(r, takeOnWater);
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

    const onWaterCompletionYmd = (row: any, neededQty: number) => {
      let remaining = Math.max(0, Math.floor(Number(neededQty || 0)));
      if (!Number.isFinite(remaining) || remaining <= 0) return '';

      const ships = Array.isArray(row?.onWaterShipments) ? row.onWaterShipments : [];
      const normalized = ships
        .map((x: any) => ({ edd: toYmd(x?.edd), qty: Math.floor(Number(x?.qty || 0)) }))
        .filter((x: any) => x.edd && Number.isFinite(x.qty) && x.qty > 0)
        .sort((a: any, b: any) => String(a.edd).localeCompare(String(b.edd)));
      for (const s of normalized) {
        if (remaining <= 0) break;
        remaining -= Math.min(remaining, Math.max(0, s.qty));
        if (remaining <= 0) return s.edd;
      }
      return toYmd(row?.onWaterEdd);
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
        const edd = onWaterCompletionYmd(r, qty);
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

  const suggestShipdateForReservedBreakdown = useCallback((opts?: { rows?: any[]; reserved?: any[] }) => {
    const today = todayYmd;
    const pickerRows = Array.isArray(opts?.rows) ? opts!.rows : (manualPickerRows || []);
    const reserved = Array.isArray(opts?.reserved) ? opts!.reserved : (Array.isArray(manualReservedBreakdown) ? manualReservedBreakdown : []);

    const rowsByGroup = new Map((pickerRows || []).map((r: any) => [String(r.groupName || ''), r]));
    const toYmd = (s: any) => {
      const v = String(s || '');
      const slice = v.slice(0, 10);
      return /^\d{4}-\d{2}-\d{2}$/.test(slice) ? slice : '';
    };

    const onWaterCompletionYmd = (row: any, neededQty: number) => {
      let remaining = Math.max(0, Math.floor(Number(neededQty || 0)));
      if (!Number.isFinite(remaining) || remaining <= 0) return '';

      const ships = Array.isArray(row?.onWaterShipments) ? row.onWaterShipments : [];
      const normalized = ships
        .map((x: any) => ({ edd: toYmd(x?.edd), qty: Math.floor(Number(x?.qty || 0)) }))
        .filter((x: any) => x.edd && Number.isFinite(x.qty) && x.qty > 0)
        .sort((a: any, b: any) => String(a.edd).localeCompare(String(b.edd)));
      for (const s of normalized) {
        if (remaining <= 0) break;
        remaining -= Math.min(remaining, Math.max(0, s.qty));
        if (remaining <= 0) return s.edd;
      }
      return toYmd(row?.onWaterEdd);
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

    for (const rr of reserved) {
      const g = String((rr as any)?.groupName || (rr as any)?.id || '').trim();
      if (!g) continue;
      const onWaterQty = Math.floor(Number((rr as any)?.onWater || 0));
      const onProcessQty = Math.floor(Number((rr as any)?.onProcess || 0));
      const secondQty = Math.floor(Number((rr as any)?.second || 0));

      if (Number.isFinite(secondQty) && secondQty > 0) {
        hasSecond = true;
      }

      const r: any = rowsByGroup.get(g) || {};
      if (Number.isFinite(onWaterQty) && onWaterQty > 0) {
        const edd = onWaterCompletionYmd(r, onWaterQty);
        if (edd && (!best || edd > best)) best = edd;
      }
      if (Number.isFinite(onProcessQty) && onProcessQty > 0) {
        const base = toYmd(r?.onProcessEdd);
        const ready = base ? addMonthsYmd(base, 3) : '';
        if (ready && (!best || ready > best)) best = ready;
      }
    }

    let out = best || today;
    if (hasSecond) {
      const cutoff = addMonthsYmd(today, 3);
      if (cutoff) {
        // When both second-warehouse and on-process/on-water are involved,
        // shipdate should be the later of the two dates.
        out = !out || cutoff > out ? cutoff : out;
      }
    }
    return out || today;
  }, [manualPickerRows, manualReservedBreakdown, todayYmd]);

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
      return { allocations: [] as any[], reserved: [] as any[] };
    }
    try {
      const { data } = await api.get(`/orders/unfulfilled/${id}`, { params: { _ts: Date.now() } });
      const nextAllocs = Array.isArray((data as any)?.allocations) ? (data as any).allocations : [];
      const nextReserved = Array.isArray((data as any)?.reservedBreakdown) ? (data as any).reservedBreakdown : [];
      setManualAllocations(nextAllocs);
      setManualReservedBreakdown(nextReserved);
      setManualRecalcTick((t) => t + 1);
      setManualEditRow((prev) => {
        if (!prev) return prev;
        return { ...(prev as any), allocations: nextAllocs } as any;
      });
      // Also sync server-driven status and estFulfillmentDate to the modal and main list
      try {
        const nextStatus = normalizeStatus((data as any)?.status || '') as any;
        const nextShip = String((data as any)?.estFulfillmentDate || '').slice(0, 10);
        if (nextStatus && !manualStatusTouched) setManualStatus(nextStatus as any);
        setManualEstFulfillment(nextShip || '');
        if (!manualOpen) {
          setOrdersRows((prev) => {
            const rows = Array.isArray(prev) ? prev : [];
            return rows.map((r) => {
              const rRawId = String((r as any)?.rawId || '').trim();
              if (rRawId && rRawId === id) {
                const safeShip = (nextShip || (r as any).estFulfillmentDate);
                const safeStatus = (nextStatus || (r as any).status);
                return { ...(r as any), status: safeStatus, estFulfillmentDate: safeShip } as any;
              }
              return r;
            });
          });
        }
      } catch {}
      return { allocations: nextAllocs, reserved: nextReserved };
    } catch {
      // keep existing allocations if fetch fails
      return { allocations: Array.isArray(manualAllocations) ? manualAllocations : [], reserved: [] as any[] };
    }
  }, [manualAllocations, manualReservedBreakdown, manualStatusTouched, manualOpen]);

  // Return/Damage feature removed

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
    const isEditing = manualOpen && manualMode === 'edit' && !manualIsLocked;
    if (!isEditing) return;
    if (manualShipdateTouched) return;

    const rawId = String((manualEditRow as any)?.rawId || '').trim();

    const reserved = Array.isArray(manualReservedBreakdown) ? manualReservedBreakdown : [];
    // Build per-group needs by summing quantities across rows that share the same pallet group
    const needByGroup = new Map<string, number>();
    for (const r of (Array.isArray(manualOrderRows) ? manualOrderRows : []) as any[]) {
      const g = String((r as any)?.groupName || '').trim();
      if (!g) continue;
      const rowId = String((r as any)?.id || '');
      const qtyRaw = (manualOrderQtyByGroup as any)?.[rowId] ?? '';
      const q = Math.max(0, Math.floor(Number(qtyRaw || 0)) || 0);
      if (q > 0) needByGroup.set(g, (needByGroup.get(g) || 0) + q);
    }

    // Keep the server-computed completion date if updated quantities are still fully covered by PRIMARY only.
    const stillFullyPrimary = Array.from(needByGroup.entries()).every(([groupName, need]) => {
      const rr: any = reserved.find((r: any) => String(r?.groupName || '').trim() === String(groupName)) || {};
      const primary = Math.max(0, Math.floor(Number(rr?.primary || 0)));
      const onWater = Math.max(0, Math.floor(Number(rr?.onWater || 0)));
      const second = Math.max(0, Math.floor(Number(rr?.second || 0)));
      const onProcess = Math.max(0, Math.floor(Number(rr?.onProcess || 0)));
      const other = onWater + second + onProcess;
      return primary >= need && other <= 0;
    });

    if (stillFullyPrimary) return;

    // Edit mode shipdate must come from the reserved breakdown (system-managed) only.
    // Do NOT fall back to selection-based logic because reserved breakdown can briefly be empty during refresh,
    // which causes flip-flopping between two dates.
    if (!reserved.length) return;
    const next = suggestShipdateForReservedBreakdown({ rows: manualPickerRows, reserved });
    // Do not override server-computed status here. The server rebalance determines READY TO SHIP vs PROCESSING.

    if (next && next !== manualEstFulfillment) {
      const last = lastAutoSuggestedShipdateRef.current || { ymd: '', at: 0 };
      const now = Date.now();
      const tooSoon = now - Number(last.at || 0) < 1500;
      const same = String(last.ymd || '') === String(next || '');
      if (tooSoon && same) return;
      lastAutoSuggestedShipdateRef.current = { ymd: String(next || ''), at: now };

      setManualEstFulfillment(next);
      setManualEditRow((prev) => {
        if (!prev) return prev;
        const rawId = String((prev as any)?.rawId || '').trim();
        if (!rawId) return prev;
        return { ...(prev as any), estFulfillmentDate: next } as any;
      });
      if (!manualOpen) {
        setOrdersRows((prev) => {
          const rows = Array.isArray(prev) ? prev : [];
          if (!rawId) return rows;
          return rows.map((r) => {
            const rRawId = String((r as any)?.rawId || '').trim();
            if (rRawId && rRawId === rawId) return { ...(r as any), estFulfillmentDate: next } as any;
            return r;
          });
        });
      }
    }
  }, [manualOpen, manualMode, manualIsLocked, manualShipdateTouched, manualOrderQtyByGroup, manualReservedBreakdown, manualPickerRows, suggestShipdateForReservedBreakdown, suggestShipdateForSelection, manualEstFulfillment, manualStatus, (manualEditRow as any)?.rawId]);

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
      if (manualEditRefreshInFlightRef.current) return;
      const now = Date.now();
      const lastAt = Number(manualEditRefreshLastAtRef.current || 0);
      // throttle to avoid accidental rapid re-runs due to re-render/effect churn
      if (now - lastAt < 2500) return;
      manualEditRefreshLastAtRef.current = now;
      manualEditRefreshInFlightRef.current = true;
      const picker = await refreshManualPicker(wid);
      await refreshManualAvailable(wid);
      await refreshManualAllocations(rawId);
      manualEditRefreshInFlightRef.current = false;
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
  }, [manualOpen, manualMode, manualPrevStatus, manualIsLocked, manualWarehouseId, (manualEditRow as any)?.rawId, refreshManualPicker, refreshManualAvailable, refreshManualAllocations]);
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
      if (name.toLowerCase() === current.toLowerCase()) return true;
      return !selected.has(name.toLowerCase());
    });
  };

  function normalizeStatus(v: any) {
    const s = String(v || '').trim().toLowerCase();
    if (s === 'create') return 'processing';
    if (s === 'created') return 'processing';
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

      const mappedUnfulfilled: OrdersRow[] = unfulfilled.map((o: any) => {
        const estShip = normalizeDateValue(o?.estFulfillmentDate);
        const st = normalizeStatus(o?.status || 'processing') as any;
        const safeStatus = st === 'ready_to_ship' && !String(estShip || '').trim() ? 'processing' : st;
        return {
          id: `manual:${String(o?._id || o?.orderNumber || Math.random())}`,
          rawId: String(o?._id || ''),
          orderNumber: String(o?.orderNumber || ''),
          type: 'manual',
          status: safeStatus,
          warehouseId: normalizeId(o?.warehouseId),
          warehouseName: whNameById.get(normalizeId(o?.warehouseId)) || '',
          createdAt: normalizeDateValue(o?.createdAt),
          updatedAt: normalizeDateValue(o?.updatedAt),
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
          originalPrice: o?.originalPrice,
          shippingPercent: o?.shippingPercent,
          discountPercent: o?.discountPercent,
          finalPrice: o?.finalPrice,
          estFulfillmentDate: estShip,
          estDeliveredDate: normalizeDateValue(o?.estDeliveredDate),
          notes: String(o?.notes || ''),
          source: 'manual',
        } as any;
      });

      const mappedFulfilled: OrdersRow[] = fulfilled.map((o: any) => ({
        id: `import:${String(o?._id || o?.orderNumber || Math.random())}`,
        rawId: String(o?._id || ''),
        orderNumber: String(o?.orderNumber || ''),
        type: String(o?.source || '') === 'csv' ? 'import' : 'manual',
        status: normalizeStatus(o?.status || 'completed') || 'completed',
        warehouseId: normalizeId(o?.warehouseId),
        warehouseName: whNameById.get(normalizeId(o?.warehouseId)) || '',
        createdAt: normalizeDateValue(o?.createdAt),
        updatedAt: normalizeDateValue(o?.updatedAt),
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
      ordersRowsRef.current = merged;
    } catch {
      setOrdersRows([]);
      ordersRowsRef.current = [];
    } finally {
      setOrdersLoading(false);
    }
  }

  const syncSomeOrdersFromServer = useCallback(async (max: number = 15) => {
    try {
      const list = Array.isArray(ordersRowsRef.current) ? ordersRowsRef.current : [];
      // Prioritize rows that appear as DEFICIT in the grid (processing with empty shipdate)
      const deficitLike = list.filter((r) => String(r.id || '').startsWith('manual:') && normalizeStatus(r.status) === 'processing' && !String(r?.estFulfillmentDate || '').trim());
      const others = list.filter((r) => String(r.id || '').startsWith('manual:') && !(normalizeStatus(r.status) === 'processing' && !String(r?.estFulfillmentDate || '').trim()));
      const ordered = [...deficitLike, ...others];
      const candidates = ordered.slice(0, Math.max(1, Math.min(max, 50)));
      if (!candidates.length) return;
      const updates = await Promise.all(
        candidates.map(async (r) => {
          const rawId = String((r as any)?.rawId || '').trim();
          if (!rawId) return null as any;
          try {
            const { data } = await api.get(`/orders/unfulfilled/${rawId}`, { params: { _ts: Date.now() } });
            const nextStatus = String((data as any)?.status || '').trim();
            const nextShip = String((data as any)?.estFulfillmentDate || '').slice(0, 10);
            return { rawId, nextStatus, nextShip };
          } catch {
            return null as any;
          }
        })
      );
      const valid = (updates || []).filter(Boolean) as Array<{ rawId: string; nextStatus: string; nextShip: string }>
      if (!valid.length) return;
      setOrdersRows((prev) => {
        const rows = Array.isArray(prev) ? prev : [];
        const byId = new Map(valid.map((u) => [u.rawId, u]));
        const next = rows.map((r) => {
          const rawId = String((r as any)?.rawId || '').trim();
          const upd = rawId ? byId.get(rawId) : undefined;
          if (!upd) return r;
          let ns = normalizeStatus(upd.nextStatus || (r as any).status);
          const ship = upd.nextShip || (r as any).estFulfillmentDate;
          // Guard: do not show READY TO SHIP without a shipdate; treat as processing until date exists.
          if (ns === 'ready_to_ship' && !String(ship || '').trim()) ns = 'processing' as any;
          return { ...(r as any), status: ns, estFulfillmentDate: ship } as any;
        });
        ordersRowsRef.current = next;
        return next;
      });
    } catch {}
  }, []);

  useEffect(() => {
    const handler = async () => {
      if (shipmentsRebalanceTimerRef.current) {
        clearTimeout(shipmentsRebalanceTimerRef.current);
      }
      shipmentsRebalanceTimerRef.current = setTimeout(async () => {
        try {
          await api.post('/orders/unfulfilled/rebalance-processing', {});
        } catch {
          // ignore
        }

        try {
          await loadOrders();
        } catch {
          // ignore
        }

        const curStatus = String(manualStatus || '').trim().toLowerCase();
        const isRecalcAllowed = manualOpen && manualMode === 'edit' && !manualIsLocked && (curStatus === 'processing' || curStatus === 'ready_to_ship');
        const rawId = String((manualEditRow as any)?.rawId || '').trim();
        const wid = String(manualWarehouseId || '').trim();
        if (isRecalcAllowed && rawId) {
          try {
            const picker = wid ? await refreshManualPicker(wid) : { rows: [] as any[], warehouses: [] as any[] };
            if (wid) {
              try {
                await refreshManualAvailable(wid);
              } catch {
                // ignore
              }
            }
            const { reserved } = await refreshManualAllocations(rawId);

            if (!manualShipdateTouched) {
              const next = suggestShipdateForReservedBreakdown({ rows: picker?.rows, reserved });
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
          } catch {
            // ignore
          }
        }
      }, 500);
    };

    window.addEventListener('shipments-changed', handler as any);
    return () => {
      if (shipmentsRebalanceTimerRef.current) {
        clearTimeout(shipmentsRebalanceTimerRef.current);
        shipmentsRebalanceTimerRef.current = null;
      }
      window.removeEventListener('shipments-changed', handler as any);
    };
  }, [manualOpen, manualMode, manualStatus, manualIsLocked, manualEditRow, manualWarehouseId, manualShipdateTouched, manualEstFulfillment, refreshManualPicker, refreshManualAvailable, refreshManualAllocations, suggestShipdateForReservedBreakdown]);

  useEffect(() => {
    let stopped = false;
    const tick = async () => {
      if (stopped) return;
      if (ordersLoading) return;
      // Pause background polling while modal is open to keep typing snappy
      if (manualOpen) return;
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
  }, [ordersLoading, manualOpen]);

  // Listen for cross-page signals that inventory tiers changed (e.g., On-Process updates/transfers)
  // and refresh the main Orders list without requiring the user to open the order modal.
  useEffect(() => {
    let timer: any = null;
    let trailing: any = null;
    const handler = async () => {
      // While modal is open (especially in create mode), skip heavy refresh to avoid input lag
      if (manualOpen) return;
      if (timer) clearTimeout(timer);
      timer = setTimeout(async () => {
        try {
          try { await api.post('/orders/unfulfilled/rebalance-processing', {}); } catch {}
          try { await new Promise((r) => setTimeout(r, 180)); } catch {}
          await loadOrders();
          await syncSomeOrdersFromServer(50);
          if (trailing) clearTimeout(trailing);
          trailing = setTimeout(async () => {
            try { await api.post('/orders/unfulfilled/rebalance-processing', {}); } catch {}
            try { await new Promise((r) => setTimeout(r, 250)); } catch {}
            await loadOrders();
            await syncSomeOrdersFromServer(50);
          }, 1200);
        } catch {}
      }, 400);
    };
    window.addEventListener('orders-changed', handler as any);
    window.addEventListener('shipments-changed', handler as any);
    return () => {
      if (timer) clearTimeout(timer);
      if (trailing) clearTimeout(trailing);
      window.removeEventListener('orders-changed', handler as any);
      window.removeEventListener('shipments-changed', handler as any);
    };
  }, [ordersLoading, syncSomeOrdersFromServer, manualOpen]);

  // When the order modal is open, also refresh the reserved breakdown immediately
  // after external change signals so the Deficit column recomputes without
  // requiring the user to close/reopen the modal.
  useEffect(() => {
    if (!manualOpen || manualMode !== 'edit') return;
    let timer: any = null;
    let trailing: any = null;
    const handler = async () => {
      const rawId = String((manualEditRow as any)?.rawId || '').trim();
      if (!rawId) return;
      if (timer) clearTimeout(timer);
      timer = setTimeout(async () => {
        try {
          // Proactively ask the server to rebalance before fetching allocations,
          // so reservedBreakdown reflects the latest On-Process cancellations/deductions.
          try { await api.post('/orders/unfulfilled/rebalance-processing', {}); } catch {}
          const wid = String(manualWarehouseId || '').trim();
          if (wid) {
            try { await refreshManualPicker(wid); } catch {}
            try { await refreshManualAvailable(wid); } catch {}
          }
          await refreshManualAllocations(rawId);
          // Schedule a trailing refresh shortly after to catch any server-side
          // rebalancing that completes just after the first fetch.
          if (trailing) clearTimeout(trailing);
          trailing = setTimeout(async () => {
            try { await api.post('/orders/unfulfilled/rebalance-processing', {}); } catch {}
            try { await refreshManualAllocations(rawId); } catch {}
          }, 900);
        } catch {}
      }, 300);
    };
    window.addEventListener('orders-changed', handler as any);
    window.addEventListener('shipments-changed', handler as any);
    return () => {
      if (timer) clearTimeout(timer);
      if (trailing) clearTimeout(trailing);
      window.removeEventListener('orders-changed', handler as any);
      window.removeEventListener('shipments-changed', handler as any);
    };
  }, [manualOpen, manualMode, (manualEditRow as any)?.rawId, manualWarehouseId, refreshManualPicker, refreshManualAvailable, refreshManualAllocations]);

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
    { field: 'orderNumber', headerName: 'Order #', flex: 1, minWidth: 110 },
    {
      field: 'status',
      headerName: 'Status',
      width: 140,
      renderCell: (p: any) => {
        const row = (p?.row as any) || {};
        const s = normalizeStatus(row?.status || '');
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
    {
      field: 'dateCreated',
      headerName: 'Order Created',
      width: 120,
      sortable: true,
      renderCell: (p: any) => fmtDate((p?.row as any)?.dateCreated),
    },
    {
      field: 'estFulfillmentDate',
      headerName: 'Estimated Shipdate for Customer',
      width: 120,
      renderCell: (p: any) => fmtDate((p?.row as any)?.estFulfillmentDate),
    },
    {
      field: 'estDeliveredDate',
      headerName: 'Estimated Arrival Date',
      width: 120,
      renderCell: (p: any) => {
        const row = (p?.row as any) || {};
        const st = normalizeStatus(row?.status || '');
        const ymd = String(row?.estDeliveredDate || '').slice(0, 10);
        const isDue = st === 'shipped' && ymd && ymd <= todayYmd;
        const label = fmtDate(row?.estDeliveredDate);
        return (
          <Box
            component="span"
            sx={
              isDue
                ? {
                    px: 1,
                    py: 0.25,
                    borderRadius: 1,
                    bgcolor: 'rgba(245, 124, 0, 0.18)',
                    fontWeight: 600,
                  }
                : undefined
            }
          >
            {label}
          </Box>
        );
      },
    },
    {
      field: 'shippingPercent',
      headerName: 'Shipping Charges (%)',
      width: 100,
      type: 'number',
      renderCell: (p: any) => {
        const v = (p?.row as any)?.shippingPercent;
        if (v === null || v === undefined || v === '') return '-';
        const n = Number(v);
        return Number.isFinite(n) ? `${n}%` : '-';
      },
    },
    {
      field: 'finalPrice',
      headerName: 'Grand Total',
      width: 110,
      type: 'number',
      renderCell: (p: any) => {
        const row = (p?.row as any) || {};
        const op = Number(row?.originalPrice);
        const sp = Number(row?.shippingPercent);
        if (Number.isFinite(op)) {
          const ship = Number.isFinite(sp) ? Math.min(100, Math.max(0, sp)) : 0;
          const calc = op * (1 + ship / 100);
          if (Number.isFinite(calc)) {
            return calc.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
          }
        }
        return '-';
      },
    },
    {
      field: 'warehouseName',
      headerName: 'Warehouse',
      flex: 1,
      minWidth: 100,
      renderCell: (p: any) => String((p?.row as any)?.warehouseName || '-')
    },
    { field: 'email', headerName: 'Email', flex: 1, minWidth: 190 },
    { field: 'customerName', headerName: 'Customer Name', flex: 1, minWidth: 190, renderCell: (p: any) => String((p?.row as any)?.customerName || '-') },
    { field: 'lineCount', headerName: 'Lines', width: 80, type: 'number' },
    { field: 'totalQty', headerName: 'Qty', width: 80, type: 'number' },
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
    setManualStatusTouched(false);
    setManualCustomerEmail('');
    setManualCustomerName('');
    setManualCustomerPhone('');
    setManualCreatedAt(new Date().toISOString().slice(0, 10));
    setManualEstFulfillment('');
    setManualEstDelivered('');
    setManualShipdateTouched(false);
    setManualShippingAddress('');
    setManualNotes('');
    setManualOriginalPrice('');
    setManualShippingPercent('');
    setManualDiscountPercent('');
    setManualLineMode('pallet_group');
    setManualLines([{ lineItem: '', qty: '' }]);
    setManualAvailable({});
    setManualValidationErrors([]);
    setManualPickerQ('');
    setManualPickerRows([]);
    setManualPickerWarehouses([]);
    setManualOrderQtyByGroup({});
    setManualDiscountByGroup({});
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
    // Only require shipdate during create. In edit mode, server auto-computes; allow saving without shipdate
    if (manualMode !== 'edit' && !manualEstFulfillment) {
      errs.push('Estimated Shipdate for Customer is required');
    }
    if (manualMode === 'edit' && manualStatus === 'shipped' && !manualEstDelivered) {
      errs.push('Estimated Arrival Date is required when status is SHIPPED');
    }
    if (manualMode === 'edit' && manualStatus === 'shipped' && manualEstDelivered && manualEstFulfillment && manualEstDelivered < manualEstFulfillment) {
      errs.push('Estimated Arrival Date cannot be earlier than Estimated Shipdate for Customer');
    }
    if (manualMode === 'edit' && manualStatus === 'shipped' && !String(manualShippingPercent || '').trim()) {
      errs.push('Shipping Charges (%) is required when status is SHIPPED');
    }
    // Order-level Discount (%) removed; only per-line discounts remain
    const sp = Number(manualShippingPercent);
    if (manualShippingPercent && (!Number.isFinite(sp) || sp < 0 || sp > 100)) {
      errs.push('Shipping Charges (%) must be between 0 and 100');
    }
    if (!String(manualShippingAddress || '').trim()) {
      errs.push('Shipping Address is required');
    }

    // manualOrderQtyByGroup is keyed by row ID; derive pallet groupName via keyToGroupName
    const allowed = new Set(
      (manualOrderGroups || [])
        .map((k) => keyToGroupName(String(k || '')).trim().toLowerCase())
        .filter((v) => v)
    );
    const parsed = Object.entries(manualOrderQtyByGroup || {})
      .map(([rowKey, qtyStr]) => {
        const groupName = keyToGroupName(String(rowKey || ''));
        const gTrim = String(groupName || '').trim();
        const gLower = gTrim.toLowerCase();
        const persistedUnit = Number((manualUnitPriceByRow as any)?.[String(rowKey || '')]);
        const latestFromMap = Number(groupPriceByName?.[gLower]);
        const unitPrice = Number.isFinite(persistedUnit)
          ? persistedUnit
          : (Number.isFinite(latestFromMap) ? latestFromMap : 0);
        return {
          rowKey: String(rowKey || ''),
          groupName: gTrim,
          qty: Math.floor(Number(qtyStr || 0)),
          discountPercent: (() => {
            const raw = (manualDiscountByGroup as any)?.[rowKey] ?? '';
            const num = Number(raw || 0);
            if (!Number.isFinite(num)) return 0;
            return Math.max(0, Math.min(100, Math.floor(num)));
          })(),
          unitPrice,
        };
      })
      .filter((l) => l.groupName && allowed.has(l.groupName.toLowerCase()) && Number.isFinite(l.qty) && l.qty > 0);
    if (!parsed.length) {
      errs.push('At least 1 Pallet ID is required');
    }

    // No manual Original Price validation; Sub Total is computed from rows
    // Validate over-ordered quantities against total availability across tiers per pallet.
    // Preflight: refresh latest availability/allocations to ensure checks use up-to-date data
    try {
      const wid = String(manualWarehouseId || '').trim();
      const rawId = String((manualEditRow as any)?.rawId || '').trim();
      if (wid) { try { await refreshManualPicker(wid); } catch {} try { await refreshManualAvailable(wid); } catch {} }
      if (rawId) { try { await refreshManualAllocations(rawId); } catch {} }
      // small settle delay for UI state to apply
      try { await new Promise((r)=> setTimeout(r, 120)); } catch {}
    } catch {}
    // Sum quantities across all rows (main + duplicates) per pallet group, then compare
    // against a single availability pool for that pallet.
    if (parsed.length) {
      const byGroup = new Map<string, any>();
      for (const r of (Array.isArray(manualOrderRows) ? manualOrderRows : [])) {
        const g = String((r as any)?.groupName || '').trim();
        if (!g) continue;
        // last one wins, but all instances share the same availability numbers
        byGroup.set(g, r);
      }

      const needByGroup = new Map<string, number>();
      for (const l of parsed) {
        const gName = String(l.groupName || '').trim();
        if (!gName) continue;
        const prev = needByGroup.get(gName) || 0;
        needByGroup.set(gName, prev + Number(l.qty || 0));
      }

      const overErrors: string[] = [];
      for (const [gNameRaw, totalNeedRaw] of needByGroup.entries()) {
        const gName = String(gNameRaw || '').trim();
        const totalNeed = Math.floor(Number(totalNeedRaw || 0));
        if (!gName || !Number.isFinite(totalNeed) || totalNeed <= 0) continue;

        const r: any = byGroup.get(gName) || {};
        const primary = Number(r?.selectedWarehouseAvailable ?? 0);
        const onWater = Number(r?.onWaterPallets ?? 0);
        const onProcess = Number(r?.onProcessPallets ?? 0);
        let second = 0;
        if (secondWarehouse?._id) {
          const per = r?.perWarehouse || {};
          const wid = String(secondWarehouse._id);
          second = Number((per && typeof per === 'object') ? (per[wid] ?? per[String(wid)] ?? 0) : 0);
        }

        // If we have no availability info for this pallet, skip client-side check
        const hasAnyAvailabilityInfo =
          Number.isFinite(primary) ||
          Number.isFinite(onWater) ||
          Number.isFinite(onProcess) ||
          Number.isFinite(second);
        if (!hasAnyAvailabilityInfo) continue;

        const baseMax = Math.max(0, Math.floor(
          (Number.isFinite(primary) ? primary : 0) +
          (Number.isFinite(onWater) ? onWater : 0) +
          (Number.isFinite(second) ? second : 0) +
          (Number.isFinite(onProcess) ? onProcess : 0)
        ));
        const reserved = manualReservedByGroup.get(gName)?.total || 0;
        const maxAllowed = baseMax + Math.max(0, reserved);

        if (totalNeed > maxAllowed) {
          const overBy = Math.max(0, totalNeed - maxAllowed);
          const gLower = gName.toLowerCase();
          const palletId = String(r?.lineItem || manualBaseLineItemByGroupRef.current.get(gLower) || '') || '';
          const palletName = String(palletNameByGroup?.[gLower] || '');
          const labelParts = [
            palletName ? `Name: ${palletName}` : '',
            `Desc: ${gName}`,
            palletId ? `ID: ${palletId}` : '',
          ].filter(Boolean);
          overErrors.push(`Over-ordered — ${labelParts.join(' | ')} — Ordered ${totalNeed}, Max ${maxAllowed} (over by ${overBy})`);
        }
      }
      if (overErrors.length) {
        setManualValidationErrors([...(errs || []), ...overErrors]);
        toast.error(overErrors[0]);
        return;
      }
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

    const originalPriceRounded = Number(Number(manualOrderSubTotal).toFixed(2));

    // Client-side availability check (server also enforces)
    // For backorder, allow negative stock, so skip this check.
    try {
      if (manualMode === 'edit' && manualEditRow) {
        // Only block switching back to PROCESSING if inventory has already been deducted (SHIPPED/DELIVERED/COMPLETED).
        // READY TO SHIP is not inventory-deducting and can revert to PROCESSING.
        if ((manualPrevStatus === 'shipped' || manualPrevStatus === 'delivered' || manualPrevStatus === 'completed') && manualStatus === 'processing') {
          throw new Error('Changing status back to PROCESSING is not allowed because inventory has already been deducted.');
        }

        if (manualPrevStatus === 'processing' && manualStatus === 'shipped') {
          const ok = window.confirm(
            'This action cannot be undone. Are you sure you want to ship this order now?'
          );
          if (!ok) return;
        }

        if (manualPrevStatus === 'ready_to_ship' && manualStatus === 'shipped') {
          const ok = window.confirm(
            'This action cannot be undone. Are you sure you want to ship this order now?'
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

        if (manualStatus === 'processing' || manualStatus === 'ready_to_ship') {
          // Let the details endpoint recompute status based on reservations after edits.
          await api.put(`/orders/unfulfilled/${manualEditRow.rawId}`, {
            customerEmail: manualCustomerEmail.trim(),
            customerName: manualCustomerName.trim(),
            customerPhone: manualCustomerPhone.trim(),
            originalPrice: originalPriceRounded,
            shippingPercent: manualShippingPercent ? Number(manualShippingPercent) : undefined,
            // order-level discount removed; per-line discounts are used
            estFulfillmentDate: manualEstFulfillment || undefined,
            estDeliveredDate: manualEstDelivered || undefined,
            shippingAddress: manualShippingAddress.trim(),
            notes: manualNotes.trim(),
            lines: parsed.map((l) => {
              const g = String(l.groupName || '').trim();
              const gLower = g.toLowerCase();
              const snapItems = (manualItemsByGroup as any)?.[gLower];
              const snapUnit = (manualUnitPriceByGroup as any)?.[gLower];
              return {
                search: g,
                qty: l.qty,
                discountPercent: l.discountPercent,
                unitPrice: l.unitPrice,
                unitPriceAtAdd: Number.isFinite(Number(snapUnit)) ? Number(snapUnit) : undefined,
                itemsAtAdd: Array.isArray(snapItems) ? snapItems : undefined,
              } as any;
            }),
          });
          // Persist snapshots to local cache for this order in case backend strips custom fields
          try {
            const orderKey = String((manualEditRow as any)?.rawId || (manualEditRow as any)?.id || '').trim();
            if (orderKey) {
              const cacheKey = `orderLineSnapshots:${orderKey}`;
              const cached = localStorage.getItem(cacheKey);
              const parsed = cached ? JSON.parse(cached) : {};
              parsed.unit = { ...(parsed.unit || {}), ...(manualUnitPriceByGroup || {}) };
              parsed.items = { ...(parsed.items || {}), ...(manualItemsByGroup || {}) };
              localStorage.setItem(cacheKey, JSON.stringify(parsed));
            }
          } catch {}
          try { window.dispatchEvent(new Event('orders-changed')); } catch {}
        } else {
          // Status changes are handled by a separate endpoint.
          // Only persist line edits first if shipping from READY TO SHIP so allocations/inventory deduction use latest quantities.
          if (manualPrevStatus === 'ready_to_ship' && manualStatus === 'shipped') {
            const { data: updated } = await api.put(`/orders/unfulfilled/${manualEditRow.rawId}`, {
              customerEmail: manualCustomerEmail.trim(),
              customerName: manualCustomerName.trim(),
              customerPhone: manualCustomerPhone.trim(),
              originalPrice: originalPriceRounded,
              shippingPercent: manualShippingPercent ? Number(manualShippingPercent) : undefined,
              // order-level discount removed; per-line discounts are used
              estFulfillmentDate: manualEstFulfillment || undefined,
              estDeliveredDate: manualEstDelivered || undefined,
              shippingAddress: manualShippingAddress.trim(),
              notes: manualNotes.trim(),
              lines: parsed.map((l) => {
                const g = String(l.groupName || '').trim();
                const gLower = g.toLowerCase();
                const snapItems = (manualItemsByGroup as any)?.[gLower];
                const snapUnit = (manualUnitPriceByGroup as any)?.[gLower];
                return {
                  search: g,
                  qty: l.qty,
                  discountPercent: l.discountPercent,
                  unitPrice: l.unitPrice,
                  unitPriceAtAdd: Number.isFinite(Number(snapUnit)) ? Number(snapUnit) : undefined,
                  itemsAtAdd: Array.isArray(snapItems) ? snapItems : undefined,
                } as any;
              }),
            });
            // Persist snapshots to local cache for this order in case backend strips custom fields
            try {
              const orderKey = String((manualEditRow as any)?.rawId || (manualEditRow as any)?.id || '').trim();
              if (orderKey) {
                const cacheKey = `orderLineSnapshots:${orderKey}`;
                const cached = localStorage.getItem(cacheKey);
                const parsed = cached ? JSON.parse(cached) : {};
                parsed.unit = { ...(parsed.unit || {}), ...(manualUnitPriceByGroup || {}) };
                parsed.items = { ...(parsed.items || {}), ...(manualItemsByGroup || {}) };
                localStorage.setItem(cacheKey, JSON.stringify(parsed));
              }
            } catch {}
            try { window.dispatchEvent(new Event('orders-changed')); } catch {}

            const recomputedStatus = normalizeStatus(updated?.status || '');
            if (recomputedStatus !== 'ready_to_ship') {
              const ok = window.confirm(
                'Some pallets are no longer fully available in the primary warehouse. This order cannot be shipped yet and will be set to PROCESSING. Do you want to proceed?'
              );
              if (!ok) {
                setManualStatus('ready_to_ship');
                return;
              }
              setManualStatus('processing');
              toast.info('Order updated and set to PROCESSING.');
              setManualOpen(false);
              await loadOrders();
              return;
            }
          }
          await api.put(`/orders/unfulfilled/${manualEditRow.rawId}/status`, {
            status: manualStatus,
            estDeliveredDate: manualStatus === 'shipped' ? (manualEstDelivered || undefined) : undefined,
          });
          // If canceling an order, immediately rebalance reservations so PROCESSING/READY TO SHIP orders can claim freed stock
          if (manualStatus === 'canceled') {
            try {
              const wid = String(manualWarehouseId || (manualEditRow as any)?.warehouseId || '').trim();
              // Prefer explicit group names from the order lines to scope the rebalance
              const groupNames = Array.isArray((manualEditRow as any)?.lines)
                ? ((manualEditRow as any).lines as any[])
                    .map((l) => String(l?.groupName || '').trim())
                    .filter((v) => v)
                : [];
              const { data: rebalanceResult } = await api.post('/orders/unfulfilled/rebalance-processing', {
                warehouseId: wid || undefined,
                groupNames: groupNames.length ? groupNames : undefined,
              });
              try { if (rebalanceResult && typeof rebalanceResult.updated === 'number') toast.info(`Rebalanced ${rebalanceResult.updated} orders`); } catch {}
              // Close dialog and reload orders so other orders (e.g., PROCESSING/READY TO SHIP) immediately reflect reclaimed stock
              setManualOpen(false);
              await loadOrders();
            } catch {}
          }
          try { window.dispatchEvent(new Event('orders-changed')); } catch {}
        }
        // Immediately recompute shipdate (client-side) to reflect new allocations/reservations
        try {
          const rawId = String(manualEditRow.rawId || '').trim();
          const wid = String(manualWarehouseId || '').trim();
          if (!manualShipdateTouched && rawId && wid) {
            const picker = await refreshManualPicker(wid);
            const { reserved } = await refreshManualAllocations(rawId);
            const next = suggestShipdateForReservedBreakdown({ rows: picker?.rows, reserved });
            if (next) {
              setOrdersRows((prev) => {
                const rows = Array.isArray(prev) ? prev : [];
                return rows.map((r) => {
                  const rRawId = String((r as any)?.rawId || '').trim();
                  if (rRawId && rRawId === rawId) return { ...(r as any), estFulfillmentDate: next } as any;
                  return r;
                });
              });
              // Persist the recomputed date to the server to keep in sync
              try { await api.put(`/orders/unfulfilled/${rawId}`, { estFulfillmentDate: next }); } catch {}
            }
          }
        } catch {}
        toast.success('Order updated');
        try { window.dispatchEvent(new CustomEvent('shipments-changed', { detail: { kind: 'order_updated', orderId: String(manualEditRow.rawId || '') } })); } catch {}
        setManualOpen(false);
        await loadOrders();
        return;
      }

      // Validate availability against backend before creating
      try {
        const wid = String(manualWarehouseId || '').trim();
        if (wid) {
          const picker = await refreshManualPicker(wid);
          const pickerRows = Array.isArray((picker as any)?.rows) ? (picker as any).rows : [];
          const byGroupCurrent = new Map<string, any>();
          for (const r of pickerRows as any[]) {
            const g = String((r as any)?.groupName || '').trim();
            if (g) byGroupCurrent.set(g, r);
          }
          const needByGroup = new Map<string, number>();
          for (const l of parsed) {
            const gName = String(l.groupName || '').trim();
            if (!gName) continue;
            needByGroup.set(gName, (needByGroup.get(gName) || 0) + Number(l.qty || 0));
          }
          const overErrors: string[] = [];
          for (const [gNameRaw, totalNeedRaw] of needByGroup.entries()) {
            const gName = String(gNameRaw || '').trim();
            const totalNeed = Math.floor(Number(totalNeedRaw || 0));
            const r: any = byGroupCurrent.get(gName) || {};
            const primary = Number(r?.selectedWarehouseAvailable ?? 0);
            const onWater = Number(r?.onWaterPallets ?? 0);
            const onProcess = Number(r?.onProcessPallets ?? 0);
            let second = 0;
            if (secondWarehouse?._id) {
              const per = r?.perWarehouse || {};
              const wid2 = String(secondWarehouse._id);
              second = Number((per && typeof per === 'object') ? (per[wid2] ?? per[String(wid2)] ?? 0) : 0);
            }
            const baseMax = Math.max(0, Math.floor(
              (Number.isFinite(primary) ? primary : 0) +
              (Number.isFinite(onWater) ? onWater : 0) +
              (Number.isFinite(second) ? second : 0) +
              (Number.isFinite(onProcess) ? onProcess : 0)
            ));
            if (totalNeed > baseMax) {
              const overBy = Math.max(0, totalNeed - baseMax);
              const gLower = gName.toLowerCase();
              const palletId = String(r?.lineItem || manualBaseLineItemByGroupRef.current.get(gLower) || '') || '';
              const palletName = String((palletNameByGroup as any)?.[gLower] || '');
              const labelParts = [palletName ? `Name: ${palletName}` : '', `Desc: ${gName}`, palletId ? `ID: ${palletId}` : ''].filter(Boolean);
              overErrors.push(`Over-ordered — ${labelParts.join(' | ')} — Ordered ${totalNeed}, Max ${baseMax} (over by ${overBy})`);
            }
          }
          if (overErrors.length) {
            setManualValidationErrors([...(errs || []), ...overErrors]);
            toast.error(overErrors[0]);
            return;
          }
        }
      } catch {}

      const { data: created } = await api.post('/orders/unfulfilled', {
        warehouseId: manualWarehouseId,
        status: 'processing',
        customerEmail: manualCustomerEmail.trim(),
        customerName: manualCustomerName.trim(),
        customerPhone: manualCustomerPhone.trim(),
        createdAtOrder: manualCreatedAt || undefined,
        originalPrice: originalPriceRounded,
        shippingPercent: manualShippingPercent ? Number(manualShippingPercent) : undefined,
        // order-level discount removed; per-line discounts are used
        estFulfillmentDate: manualEstFulfillment || undefined,
        estDeliveredDate: manualEstDelivered || undefined,
        shippingAddress: manualShippingAddress.trim(),
        notes: manualNotes.trim(),
        lines: parsed.map((l) => {
          const g = String(l.groupName || '').trim();
          const gLower = g.toLowerCase();
          const snapItems = (manualItemsByGroup as any)?.[gLower];
          const snapUnit = (manualUnitPriceByGroup as any)?.[gLower];
          return {
            search: g,
            qty: l.qty,
            discountPercent: l.discountPercent,
            unitPrice: l.unitPrice,
            unitPriceAtAdd: Number.isFinite(Number(snapUnit)) ? Number(snapUnit) : undefined,
            itemsAtAdd: Array.isArray(snapItems) ? snapItems : undefined,
          } as any;
        }),
      });
      // Persist snapshots to local cache for this new order (backend may strip custom fields)
      try {
        const orderKey = String((created as any)?.rawId || (created as any)?._id || (created as any)?.id || '').trim();
        if (orderKey) {
          const cacheKey = `orderLineSnapshots:${orderKey}`;
          const cached = localStorage.getItem(cacheKey);
          const parsedCache = cached ? JSON.parse(cached) : {};
          parsedCache.unit = { ...(parsedCache.unit || {}), ...(manualUnitPriceByGroup || {}) };
          parsedCache.items = { ...(parsedCache.items || {}), ...(manualItemsByGroup || {}) };
          localStorage.setItem(cacheKey, JSON.stringify(parsedCache));
        }
      } catch {}
      toast.success('Order created');
      try {
        window.dispatchEvent(new CustomEvent('shipments-changed', { detail: { kind: 'order_created' } }));
      } catch {}
      setManualOpen(false);
      await loadOrders();
    } catch (e:any) {
      const payload = e?.response?.data;
      const code = String(payload?.code || '').trim();
      if (code === 'NO_STOCKS' && Array.isArray(payload?.noStocks)) {
        const lines: string[] = ['No Stocks available for:'];
        for (const it of payload.noStocks) {
          const pid = String(it?.lineItem || '-').trim() || '-';
          const g = String(it?.groupName || '-').trim() || '-';
          const avail = Math.max(0, Math.floor(Number(it?.available || 0)));
          const req = Math.max(0, Math.floor(Number(it?.required || 0)));
          lines.push(`${pid} - ${g} - ${avail} available${req ? ` (required ${req})` : ''}`);
        }
        setManualValidationErrors(lines);
        toast.error('No Stocks available');
        return;
      }

      const msg = payload?.message || 'Failed to save';
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
          <Button variant="outlined" onClick={() => setReportOpen(true)}>Report</Button>
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
              <MenuItem value="ready_to_ship">READY TO SHIP</MenuItem>
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
          sx={{
            '& .MuiDataGrid-columnHeaderTitle': {
              whiteSpace: 'normal',
              lineHeight: 1.2,
            },
            '& .MuiDataGrid-row': {
              alignItems: 'center',
            },
          }}
          columnHeaderHeight={90}
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

      <Dialog open={reportOpen} onClose={() => setReportOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Pallet Sales Report</DialogTitle>
        <DialogContent>
          <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} sx={{ mt: 1 }}>
            <TextField
              type="date"
              size="small"
              label="From"
              value={reportFrom}
              onChange={(e) => setReportFrom(e.target.value)}
              InputLabelProps={{ shrink: true }}
              fullWidth
            />
            <TextField
              type="date"
              size="small"
              label="To"
              value={reportTo}
              onChange={(e) => setReportTo(e.target.value)}
              InputLabelProps={{ shrink: true }}
              fullWidth
            />
          </Stack>
          {reportExporting ? <LinearProgress sx={{ mt: 2 }} /> : null}
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setReportOpen(false)} disabled={reportExporting}>Close</Button>
          <Button variant="outlined" onClick={exportPalletSalesReportXlsx} disabled={reportExporting}>Export .xlsx</Button>
        </DialogActions>
      </Dialog>


      <Dialog open={manualOpen} onClose={()=>{ setManualOpen(false); try { window.dispatchEvent(new Event('orders-changed')); } catch {} try { setTimeout(async()=>{ try { await api.post('/orders/unfulfilled/rebalance-processing', {}); } catch {} try { await loadOrders(); } catch {} }, 400); } catch {} try { setTimeout(async()=>{ try { await api.post('/orders/unfulfilled/rebalance-processing', {}); } catch {} try { await loadOrders(); } catch {} }, 1100); } catch {} }} fullWidth maxWidth="xl">
        <DialogTitle>{manualMode === 'edit' ? `Edit Order${manualEditRow?.orderNumber ? ` - ${manualEditRow.orderNumber}` : ''}` : 'Add Order'}</DialogTitle>
        <DialogContent>
          <Typography variant="subtitle2" sx={{ mt: 1, mb: 1 }}>Customer</Typography>
          <Box
            sx={{
              display: 'grid',
              gridTemplateColumns: { xs: '1fr', md: 'repeat(4, minmax(0, 1fr))' },
              gap: 2,
              mb: 2,
            }}
          >
            <TextField select disabled label="Warehouse" size="small" value={manualWarehouseId} onChange={(e)=>setManualWarehouseId(e.target.value)} error={!manualWarehouseId} helperText={!manualWarehouseId ? 'Required' : ''}>
              {warehouses.map((w)=> (
                <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>
              ))}
            </TextField>
            {manualMode === 'edit' ? (
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
                <TextField
                  select
                  label="Status"
                  size="small"
                  value={manualStatus}
                  onChange={(e)=> { setManualStatusTouched(true); setManualStatus(e.target.value as any); }}
                  disabled={manualIsLocked}
                  sx={{ flex: 1, minWidth: 200 }}
                  SelectProps={{
                    renderValue: (value: any) => {
                      const v = String(value || '').toLowerCase();
                      return v === 'ready_to_ship' ? 'READY TO SHIP' : (v ? v.toUpperCase() : '');
                    }
                  }}
                >
                  <MenuItem value="processing" disabled>PROCESSING</MenuItem>
                  <MenuItem value="ready_to_ship" disabled>READY TO SHIP</MenuItem>
                  <MenuItem
                    value="shipped"
                    disabled={
                      manualStatus === 'processing' ||
                      manualPrevStatus !== 'ready_to_ship'
                    }
                  >
                    SHIPPED
                  </MenuItem>
                  <MenuItem
                    value="completed"
                    disabled={
                      manualStatus === 'processing' ||
                      manualPrevStatus !== 'shipped'
                    }
                  >
                    COMPLETED
                  </MenuItem>
                  <MenuItem value="canceled" disabled={manualPrevStatus === 'completed' || manualPrevStatus === 'canceled'}>CANCELED</MenuItem>
                </TextField>
                <Tooltip
                  arrow
                  placement="top-start"
                  title={(
                    <Box sx={{ p: 0.5 }}>
                      <Typography variant="body2" sx={{ fontWeight: 600 }}>
                        Order Status - Meaning
                      </Typography>
                      <Typography variant="body2" sx={{ mt: 0.5 }}>
                        PROCESSING (auto set): pallets are not fully available in MPG (Primary WH) yet. Inventory is not deducted.
                      </Typography>
                      <Typography variant="body2" sx={{ mt: 0.5 }}>
                        READY TO SHIP (auto set): all pallets are available in MPG (Primary WH). Inventory is not deducted yet.
                      </Typography>
                      <Typography variant="body2" sx={{ mt: 0.5 }}>
                        SHIPPED: pallets are shipped and inventory is deducted. "Estimated Arrival Date" is required.
                      </Typography>
                      <Typography variant="body2" sx={{ mt: 0.5 }}>
                        COMPLETED: customer received the order and it is fully fulfilled.
                      </Typography>
                      <Typography variant="body2" sx={{ mt: 0.5 }}>
                        CANCELED: order is canceled and reserved quantity is returned back to inventory.
                      </Typography>
                    </Box>
                  )}
                >
                  <IconButton size="small" aria-label="Order status help">
                    <HelpOutlineIcon fontSize="inherit" />
                  </IconButton>
                </Tooltip>
              </Box>
            ) : null}
            <TextField disabled={manualFieldsLocked} label="Customer Email" size="small" value={manualCustomerEmail} onChange={(e)=>setManualCustomerEmail(e.target.value)} error={!isValidEmail(manualCustomerEmail)} helperText={!isValidEmail(manualCustomerEmail) ? 'Required (valid email)' : ''} />
            <TextField disabled={manualFieldsLocked} label="Customer Name" size="small" value={manualCustomerName} onChange={(e)=>setManualCustomerName(e.target.value)} error={!String(manualCustomerName||'').trim()} helperText={!String(manualCustomerName||'').trim() ? 'Required' : ''} />
            <TextField disabled={manualFieldsLocked} label="Phone Number" size="small" value={manualCustomerPhone} onChange={(e)=>setManualCustomerPhone(e.target.value)} error={!String(manualCustomerPhone||'').trim()} helperText={!String(manualCustomerPhone||'').trim() ? 'Required' : ''} />
          </Box>

          <Typography variant="subtitle2" sx={{ mb: 1 }}>Dates</Typography>
          <Box
            sx={{
              display: 'grid',
              gridTemplateColumns: { xs: '1fr', md: 'repeat(3, minmax(0, 1fr))' },
              gap: 2,
              mb: 2,
            }}
          >
            <TextField type="date" label="Created Order Date" InputLabelProps={{ shrink: true }} size="small" value={manualCreatedAt} onChange={(e)=>setManualCreatedAt(e.target.value)} inputProps={{ max: todayYmd }} error={!manualCreatedAt || manualCreatedAt > todayYmd} helperText={!manualCreatedAt ? 'Required' : (manualCreatedAt > todayYmd ? 'Cannot be advance date' : '')} disabled={manualMode === 'edit'} />
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
              <TextField
                disabled
                type="date"
                label="Estimated Shipdate for Customer"
                InputLabelProps={{ shrink: true }}
                size="small"
                value={manualEstFulfillment}
                onChange={()=>{}}
                error={!manualEstFulfillment}
                helperText={!manualEstFulfillment ? 'Required' : ''}
                fullWidth
              />
              <Tooltip
                arrow
                placement="top-start"
                title={(
                  <Box sx={{ p: 0.5 }}>
                    <Typography variant="body2" sx={{ fontWeight: 600 }}>
                      Estimated Shipdate for Customer - Computation
                    </Typography>
                    <Typography variant="body2" sx={{ mt: 0.5 }}>
                      We check where your ordered quantity will come from, then the shipdate is based on the slowest required source.
                    </Typography>
                    <Typography variant="body2" sx={{ mt: 0.5 }}>
                      1) MPG (Primary WH): if covered, ready today
                      <br />
                      2) On-Water: use the earliest EDD that can cover the remaining qty
                      <br />
                      3) PEBA (2nd WH): today + 3 months (transfer/handling time)
                      <br />
                      4) On-Process: EDD + 3 months
                    </Typography>
                  </Box>
                )}
              >
                <IconButton size="small" aria-label="Estimated shipdate computation help" sx={{ mt: 0.5 }}>
                  <HelpOutlineIcon fontSize="inherit" />
                </IconButton>
              </Tooltip>
            </Box>
            <TextField
              type="date"
              label="Estimated Arrival Date"
              InputLabelProps={{ shrink: true }}
              size="small"
              value={manualEstDelivered}
              onChange={(e)=> setManualEstDelivered(String(e.target.value || ''))}
              inputProps={{ min: manualEstFulfillment || undefined }}
              disabled={manualIsLocked || !(manualStatus === 'shipped' || manualPrevStatus === 'shipped')}
            />
          </Box>

          <Typography variant="subtitle2" sx={{ mb: 1 }}>Pricing</Typography>
          <Box
            sx={{
              display: 'grid',
              gridTemplateColumns: { xs: '1fr', md: 'repeat(3, minmax(0, 1fr))' },
              gap: 2,
              mb: 2,
            }}
          >
            <TextField
              label="Sub Total"
              size="small"
              value={Number.isFinite(Number(manualOrderSubTotal)) ? Number(manualOrderSubTotal).toFixed(2) : ''}
              disabled
            />
            <TextField
              label="Shipping Charges (%)"
              size="small"
              value={manualShippingPercent}
              onChange={(e)=> setManualShippingPercent(normalizeDiscountText(e.target.value))}
              disabled={manualFieldsLocked}
            />
            <TextField
              label="Grand Total"
              size="small"
              value={manualFinalPrice}
              disabled
            />
          </Box>

          {manualMode === 'edit' && manualLastUpdatedAt ? (
            <Typography variant="caption" color="error" sx={{ display: 'block', mb: 2 }}>
              Last Updated: {manualLastUpdatedAt}{manualLastUpdatedBy ? ` (by ${manualLastUpdatedBy})` : ''}
            </Typography>
          ) : null}
          <TextField
            fullWidth
            label="Shipping Address"
            size="small"
            value={manualShippingAddress}
            onChange={(e)=>setManualShippingAddress(e.target.value)}
            sx={{ mb: 2 }}
            error={!String(manualShippingAddress||'').trim()}
            helperText={!String(manualShippingAddress||'').trim() ? 'Required' : ''}
            disabled={manualFieldsLocked}
          />

          <TextField
            fullWidth
            label="Remarks/Notes"
            size="small"
            value={manualNotes}
            onChange={(e)=> setManualNotes(e.target.value)}
            sx={{ mb: 2 }}
            multiline
            minRows={3}
            disabled={manualFieldsLocked}
          />

          {null}

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
                  <Box sx={{ mb: 1 }}>
                    <Typography variant="body2" color="text.secondary">
                      Reserved stock breakdown for this order.
                    </Typography>
                  </Box>
                  {manualReservedRows.length ? (
                    <div style={{ height: 220, width: '100%' }}>
                      <DataGrid
                        rows={manualReservedRows}
                        columns={([
                          { field: 'palletName', headerName: 'Pallet Name', flex: 1, minWidth: 200, renderCell: (p: any) => {
                            const row: any = p?.row || {};
                            const gLower = String(row?.groupName || '').trim().toLowerCase();
                            return String(palletNameByGroup?.[gLower] || '');
                          } },
                          { field: 'palletDescription', headerName: 'Pallet Description', flex: 1, minWidth: 240, renderCell: (p: any) => {
                            const row: any = p?.row || {};
                            const gLower = String(row?.groupName || '').trim().toLowerCase();
                            return String(palletDescByGroup?.[gLower] || '');
                          } },
                          { field: 'palletId', headerName: 'Pallet ID', width: 140, renderCell: (p: any) => {
                            const row: any = p?.row || {};
                            const groupName = String(row?.groupName || '');
                            const gLower = groupName.trim().toLowerCase();
                            const picker = manualPickerRowByGroup.get(groupName) || {};
                            const hit = (Array.isArray(manualPickerRows) ? manualPickerRows : []).find(
                              (x: any) => String(x?.groupName || '').trim().toLowerCase() === gLower
                            );
                            const fromOrder = (Array.isArray(manualOrderRows) ? manualOrderRows : []).find(
                              (x: any) => String(x?.groupName || '').trim().toLowerCase() === gLower
                            );
                            const fromBase = manualLineItemByGroup.get(gLower) || manualBaseLineItemByGroupRef.current.get(gLower) || '';
                            const id = String(picker?.lineItem || hit?.lineItem || fromOrder?.lineItem || fromBase || row?.palletId || '');
                            return id || '-';
                          } },
                          { field: 'primary', headerName: manualAllocationWarehouseLabels.primaryLabel, width: 100, type: 'number', align: 'right', headerAlign: 'right' },
                          {
                            field: 'onWater',
                            headerName: 'On-Water',
                            width: 100,
                            type: 'number',
                            align: 'right',
                            headerAlign: 'right',
                            renderCell: (p: any) => {
                              const qty = Number((p?.row as any)?.onWater ?? 0);
                              if (!qty) return '0';
                              const groupName = String((p?.row as any)?.groupName || '').trim();
                              const wid = String(manualWarehouseId || '').trim();
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
                          { field: 'second', headerName: manualAllocationWarehouseLabels.secondLabel, width: 100, type: 'number', align: 'right', headerAlign: 'right' },
                          {
                            field: 'onProcess',
                            headerName: 'On-Process',
                            width: 100,
                            type: 'number',
                            align: 'right',
                            headerAlign: 'right',
                            renderCell: (p: any) => {
                              const qty = Number((p?.row as any)?.onProcess ?? 0);
                              if (!qty) return '0';
                              const groupName = String((p?.row as any)?.groupName || '').trim();
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
                          },
                        ]) as GridColDef[]}
                        sx={{
                          '& .MuiDataGrid-cell': {
                            whiteSpace: 'normal',
                            lineHeight: 2,
                            alignItems: 'center',
                          },
                          '& .MuiDataGrid-columnHeaderTitle': {
                            whiteSpace: 'normal',
                            lineHeight: 1.2,
                          },
                          '& .MuiDataGrid-row': {
                            alignItems: 'center',
                          },
                        }}
                        columnHeaderHeight={90}
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
              disabled={!manualWarehouseId || manualFieldsLocked}
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
              columns={manualOrderColumns as any}
              onRowDoubleClick={(p: any) => {
                const g = String(p?.row?.groupName || '').trim();
                if (g) openOrderableGroupItems({ groupName: g });
              }}
              getRowClassName={(params: any) => {
                const row = (params && typeof params === 'object') ? (params as any).row : {};
                return row && (row as any).isDuplicate ? 'duplicate-pallet-row' : 'original-pallet-row';
              }}
              sx={{
                '& .duplicate-pallet-row': {
                  backgroundColor: '#FAF6E9',
                },
                '& .original-pallet-row .MuiDataGrid-cell': {
                  fontWeight: 700,
                },
              }}
               columnHeaderHeight={90}
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
                  label="Search Pallet ID / Pallet Description / Pallet Name"
                  value={manualPickerQ}
                  onChange={(e)=>setManualPickerQ(e.target.value)}
                  sx={{ flex: 1, minWidth: 260 }}
                />
                <TextField
                  size="small"
                  label="EDD From"
                  type="date"
                  value={manualPickerEddFrom}
                  onChange={(e)=> setManualPickerEddFrom(e.target.value)}
                  InputLabelProps={{ shrink: true }}
                />
                <TextField
                  size="small"
                  label="EDD To"
                  type="date"
                  value={manualPickerEddTo}
                  onChange={(e)=> setManualPickerEddTo(e.target.value)}
                  InputLabelProps={{ shrink: true }}
                />
                <Button variant="outlined" onClick={async ()=>{
                  try {
                    if (!manualWarehouseId) return;
                    setManualPickerLoading(true);
                    const { data } = await api.get('/orders/pallet-picker', { params: { warehouseId: manualWarehouseId } });
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
                  rows={(Array.isArray(manualPickerRowsFiltered) ? manualPickerRowsFiltered : []).map((r: any) => ({ id: String(r.groupName || r.lineItem || ''), ...r }))}
                  columns={manualPickerColumns}
                  loading={manualPickerLoading}
                  columnHeaderHeight={90}
                  rowHeight={55}
                  onRowDoubleClick={(p: any) => {
                    const g = String(p?.row?.groupName || '').trim();
                    if (g) openOrderableGroupItems({ groupName: g });
                  }}
                  getRowClassName={(params: any) => {
                    const row = (params as any)?.row || {};
                    // If already added to order, mark distinctly
                    const isAdded = (manualOrderGroups || []).some((k) => keyToGroupName(String(k || '')).trim().toLowerCase() === String(row?.groupName || '').trim().toLowerCase());
                    if (isAdded) return 'row-already-added';
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
                    '& .row-already-added': {
                      bgcolor: 'rgba(25, 118, 210, 0.08)',
                      '&:hover': { bgcolor: 'rgba(25, 118, 210, 0.12)' },
                      color: 'rgba(0,0,0,0.6)'
                    },
                    '& .MuiDataGrid-columnHeaderTitle': {
                      whiteSpace: 'normal',
                      lineHeight: 1.1,
                    },
                    '& .MuiDataGrid-columnHeader': {
                      whiteSpace: 'normal',
                    },
                    '& .MuiDataGrid-cell': {
                      whiteSpace: 'normal',
                      lineHeight: 1.2,
                      display: 'flex',
                      alignItems: 'center',
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
                    // Disallow selecting if already added in the order list
                    const alreadyAdded = (manualOrderGroups || []).some((k) => keyToGroupName(String(k || '')).trim().toLowerCase() === String(row?.groupName || '').trim().toLowerCase());
                    if (alreadyAdded) return false;
                    // Allow selecting even when availability is 0
                    const maxOrder = Math.max(0, Math.floor(total));
                    return maxOrder >= 0;
                  }}
                  
                  disableRowSelectionOnClick
                  density="compact"
                  slots={{ toolbar: GridToolbar }}
                  slotProps={{ toolbar: { showQuickFilter: true, quickFilterProps: { debounceMs: 250 } } as any }}
                  onRowSelectionModelChange={(m: any)=> {
                    const idsArr = m?.ids ? Array.from(m.ids) : [];
                    const idsStr = idsArr.map((x: any) => String(x));
                    // Accept all selected ids; allow deficits
                    setManualPickSelected({ type: 'include', ids: new Set(idsStr) });
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
              <Button variant="contained" onClick={async ()=>{
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
                  .map((x: any) => x.id);
                // Determine which are newly added (not previously in order)
                const prevGroups = new Set((Array.isArray(manualOrderGroups) ? manualOrderGroups : []).map((g)=> String(g)));
                const newGroups = picked.filter((g)=> !prevGroups.has(String(g)));

                // Snapshot unit price and items for newly added groups
                if (newGroups.length) {
                  // Build snapshot maps
                  const unitUpdates: Record<string, number> = {};
                  const itemsUpdates: Record<string, any[]> = {};
                  // Unit price snapshot from computed groupPriceByName or row price fallback
                  for (const id of newGroups) {
                    const g = String(id);
                    const gLower = g.toLowerCase();
                    const row = (manualPickerRows || []).find((r: any)=> String(r?.groupName||'').trim().toLowerCase() === gLower) || {};
                    const mapUnit = Number((groupPriceByName as any)?.[gLower]);
                    const fallback = Number((row as any)?.palletPrice ?? (row as any)?.price ?? 0);
                    const snapUnit = Number.isFinite(mapUnit) ? mapUnit : (Number.isFinite(fallback) ? fallback : NaN);
                    if (Number.isFinite(snapUnit)) unitUpdates[g] = Number(snapUnit);
                  }
                  try {
                    await Promise.all(newGroups.map(async (g) => {
                      try {
                        const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
                        const list = Array.isArray((data as any)?.items) ? (data as any).items : [];
                        itemsUpdates[String(g).toLowerCase()] = list;
                      } catch {
                        itemsUpdates[String(g).toLowerCase()] = [];
                      }
                    }));
                  } catch {}
                  // Commit snapshots without overwriting existing ones
                  setManualUnitPriceByGroup((prev)=> ({ ...(prev||{}), ...unitUpdates }));
                  setManualItemsByGroup((prev)=> ({ ...(prev||{}), ...itemsUpdates }));
                  // Persist to localStorage cache for this order (edit mode only)
                  try {
                    const orderKey = String((manualEditRow as any)?.rawId || (manualEditRow as any)?.id || '').trim();
                    if (orderKey) {
                      const cacheKey = `orderLineSnapshots:${orderKey}`;
                      const cached = localStorage.getItem(cacheKey);
                      const parsed = cached ? JSON.parse(cached) : {};
                      if (parsed && typeof parsed === 'object') {
                        parsed.unit = { ...(parsed.unit || {}), ...unitUpdates };
                        parsed.items = { ...(parsed.items || {}), ...itemsUpdates };
                        localStorage.setItem(cacheKey, JSON.stringify(parsed));
                      } else {
                        localStorage.setItem(cacheKey, JSON.stringify({ unit: unitUpdates, items: itemsUpdates }));
                      }
                    }
                  } catch {}
                }

                setManualOrderGroups((prev)=> {
                  const base = Array.isArray(prev) ? prev : [];
                  const set = new Set(base.map((x)=> String(x)));
                  for (const id of picked) set.add(String(id));
                  return Array.from(set);
                });
                // Initialize default Order Qty to '1' for any newly added groups (do not overwrite existing values)
                setManualOrderQtyByGroup((prev)=>{
                  const next = { ...(prev || {}) } as Record<string, string>;
                  for (const id of picked) {
                    const key = String(id);
                    if (!(key in next) || String(next[key] ?? '').trim() === '') {
                      next[key] = '1';
                    }
                  }
                  return next;
                });
                // Initialize default Discount (%) to '0' for any newly added groups (do not overwrite existing values)
                setManualDiscountByGroup((prev)=>{
                  const next = { ...(prev || {}) } as Record<string, string>;
                  for (const id of picked) {
                    const key = String(id);
                    if (!(key in next) || String(next[key] ?? '').trim() === '') {
                      next[key] = '0';
                    }
                  }
                  return next;
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
            variant="outlined"
            onClick={exportManualOrderXlsx}
            disabled={manualMode !== 'edit' || !manualEditRow}
          >
            Export .xlsx
          </Button>
          <Button
            variant="outlined"
            onClick={exportManualOrderPdf}
            disabled={manualMode !== 'edit' || !manualEditRow}
          >
            Export .pdf
          </Button>
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
              // Only require shipdate when creating a new order; in edit mode allow saving without shipdate
              ((manualMode !== 'edit') && (!manualEstFulfillment)) ||
              !String(manualShippingAddress||'').trim() ||
              manualIsLocked ||
              ((manualMode !== 'edit') && (manualOrderGroups.length === 0)) ||
              ((manualMode !== 'edit') && (Object.entries(manualOrderQtyByGroup || {}).filter(([g, q]) => manualOrderGroups.includes(g) && Number(q) > 0).length === 0))
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
              label="Search Pallet ID / Pallet Description / Pallet Name"
              value={viewOrderableQ}
              onChange={(e)=>setViewOrderableQ(e.target.value)}
              sx={{ flex: 1, minWidth: 260 }}
            />
            <TextField
              size="small"
              label="EDD From"
              type="date"
              value={viewOrderableEddFrom}
              onChange={(e)=> setViewOrderableEddFrom(e.target.value)}
              InputLabelProps={{ shrink: true }}
            />
            <TextField
              size="small"
              label="EDD To"
              type="date"
              value={viewOrderableEddTo}
              onChange={(e)=> setViewOrderableEddTo(e.target.value)}
              InputLabelProps={{ shrink: true }}
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
                const cleanedPrimaryName = selectedWarehouseName.replace(/^THIS\s*-\s*/i, '');
                const wid = String(viewOrderableWarehouseId || '').trim();
                const list = Array.isArray(viewOrderableWarehouses) ? viewOrderableWarehouses : [];
                const second = list.find((w: any) => String(w?._id || '').trim() && String(w?._id || '').trim() !== wid) || null;
                const secondId = second ? String(second._id) : '';
                const secondName = second ? String(second.name || '').trim() : '';
                const from = String(viewOrderableEddFrom || '').trim();
                const to = String(viewOrderableEddTo || '').trim();

                const allRows = Array.isArray(viewOrderableFilteredRows) ? viewOrderableFilteredRows : [];
                const waterEddsSet = new Set<string>();
                const processEddsSet = new Set<string>();
                for (const r of allRows) {
                  for (const s of (Array.isArray(r?.onWaterShipments) ? r.onWaterShipments : [])) {
                    const edd = String(s?.edd || '').trim();
                    if (edd && (!from || edd >= from) && (!to || edd <= to)) waterEddsSet.add(edd);
                  }
                  for (const b of (Array.isArray(r?.onProcessBatches) ? r.onProcessBatches : [])) {
                    const edd = String(b?.edd || '').trim();
                    if (edd && (!from || edd >= from) && (!to || edd <= to)) processEddsSet.add(edd);
                  }
                }
                const waterEdds = Array.from(waterEddsSet).sort((a, b) => a.localeCompare(b));
                const processEdds = Array.from(processEddsSet).sort((a, b) => a.localeCompare(b));

                const cols: any[] = [
                  { field: 'palletName', headerName: 'Pallet Name', flex: 1, minWidth: 200, renderCell: (p: any) => {
                    const g = String(p?.row?.groupName || '').trim().toLowerCase();
                    return String(palletNameByGroup[g] || '');
                  } },
                  { field: 'palletDescription', headerName: 'Pallet Description', flex: 1, minWidth: 220, renderCell: (p: any) => {
                    const g = String(p?.row?.groupName || '').trim().toLowerCase();
                    return String(palletDescByGroup?.[g] || '');
                  } },
                  { field: 'lineItem', headerName: 'Pallet ID', width: 140, renderCell: (p: any) => String(p?.row?.lineItem || '-') },
                  { field: 'selectedWarehouseAvailable', headerName: `${cleanedPrimaryName || 'Warehouse'}`, width: 80, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => String(p?.row?.selectedWarehouseAvailable ?? 0) },
                ];

                for (const edd of waterEdds) {
                  const header = `On-Water ${(() => { const [y,m,d] = String(edd).split('-'); return `${m}/${d}/${y}`; })()}`;
                  cols.push({
                    field: `ow_${edd}`,
                    headerName: header,
                    width: 110,
                    type: 'number',
                    align: 'right',
                    headerAlign: 'right',
                    sortable: true,
                    filterable: false,
                    valueGetter: (...args: any[]) => {
                      const maybeParams = args?.[0];
                      const maybeRow = args?.[1];
                      const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
                      const list = Array.isArray(row?.onWaterShipments) ? row.onWaterShipments : [];
                      const hit = list.find((x: any) => String(x?.edd || '') === edd);
                      const v = Number(hit?.qty ?? 0);
                      return Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
                    },
                    renderCell: (p: any) => {
                      const list = Array.isArray(p?.row?.onWaterShipments) ? p.row.onWaterShipments : [];
                      const hit = list.find((x: any) => String(x?.edd || '') === edd);
                      const v = Number(hit?.qty ?? 0);
                      const qty = Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
                      return String(qty);
                    },
                  });
                }

                if (secondId) {
                  cols.push({
                    field: 'secondWarehouseAvailable',
                    headerName: `${secondName || 'Warehouse'}`,
                    width: 80,
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

                for (const edd of processEdds) {
                  const header = `On-Process ${(() => { const [y,m,d] = String(edd).split('-'); return `${m}/${d}/${y}`; })()}`;
                  cols.push({
                    field: `op_${edd}`,
                    headerName: header,
                    width: 110,
                    type: 'number',
                    align: 'right',
                    headerAlign: 'right',
                    sortable: true,
                    filterable: false,
                    valueGetter: (...args: any[]) => {
                      const maybeParams = args?.[0];
                      const maybeRow = args?.[1];
                      const row = (maybeRow && typeof maybeRow === 'object') ? maybeRow : (maybeParams?.row || {});
                      const list = Array.isArray(row?.onProcessBatches) ? row.onProcessBatches : [];
                      const hit = list.find((x: any) => String(x?.edd || '') === edd);
                      const v = Number(hit?.qty ?? 0);
                      return Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
                    },
                    renderCell: (p: any) => {
                      const list = Array.isArray(p?.row?.onProcessBatches) ? p.row.onProcessBatches : [];
                      const hit = list.find((x: any) => String(x?.edd || '') === edd);
                      const v = Number(hit?.qty ?? 0);
                      const qty = Math.max(0, Math.floor(Number.isFinite(v) ? v : 0));
                      return String(qty);
                    },
                  });
                }

                cols.push({
                  field: 'maxOrder',
                  headerName: 'Max Order',
                  width: 80,
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
              rowHeight={55}
              columnHeaderHeight={90}
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
                '& .MuiDataGrid-columnHeaderTitle': {
                  whiteSpace: 'normal',
                  lineHeight: 1.1,
                },
                '& .MuiDataGrid-columnHeader': {
                  whiteSpace: 'normal',
                },
                '& .MuiDataGrid-cell': {
                  whiteSpace: 'normal',
                  lineHeight: 1.2,
                  display: 'flex',
                  alignItems: 'center',
                },
              }}
              disableRowSelectionOnClick
              density="compact"
              slots={{ toolbar: GridToolbar }}
              slotProps={{ toolbar: { showQuickFilter: true, quickFilterProps: { debounceMs: 250 } } as any }}
              pagination
              pageSizeOptions={[10, 20, 50, 100]}
              initialState={{ pagination: { paginationModel: { page: 0, pageSize: 20 } } }}
              onRowDoubleClick={(p: any) => {
                const g = String(p?.row?.groupName || '').trim();
                if (g) openOrderableGroupItems({ groupName: g });
              }}
            />
          </div>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setViewOrderableOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={orderableGroupOpen} onClose={()=> setOrderableGroupOpen(false)} fullWidth maxWidth="md">
        <DialogTitle>{`Pallet Items - ${orderableGroupName || ''}`}</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          {orderableGroupLoading ? <LinearProgress sx={{ mb: 2 }} /> : null}
          <div style={{ height: 420, width: '100%' }}>
            <DataGrid
              rows={(Array.isArray(orderableGroupItems) ? orderableGroupItems : []).map((r: any, idx: number) => ({ id: String(r?.itemCode || idx), ...r }))}
              columns={([
                { field: 'itemCode', headerName: 'Item Code', width: 160 },
                { field: 'description', headerName: 'Description', flex: 1, minWidth: 220 },
                { field: 'upc', headerName: 'UPC', flex: 1, minWidth: 160 },
                { field: 'color', headerName: 'Color', width: 140 },
                { field: 'packSize', headerName: 'Pack Size', width: 110, type: 'number', align: 'right', headerAlign: 'right' },
                { field: 'price', headerName: 'Item Price', width: 120, type: 'number', align: 'right', headerAlign: 'right', valueGetter: (p: any) => {
                  const v = Number((p?.row as any)?.price ?? 0);
                  return Number.isFinite(v) ? v : 0;
                }, renderCell: (p: any) => {
                  const v = Number((p?.row as any)?.price ?? 0);
                  if (!Number.isFinite(v)) return '';
                  return v.toFixed(2);
                }},
              ]) as GridColDef[]}
              disableRowSelectionOnClick
              density="compact"
              pagination
              pageSizeOptions={[5, 10, 20, 50]}
              initialState={{ pagination: { paginationModel: { page: 0, pageSize: 10 } } }}
            />
          </div>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setOrderableGroupOpen(false)}>Close</Button>
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
