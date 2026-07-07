import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, Dialog, DialogTitle, DialogContent, DialogActions, MenuItem, Box, IconButton, Chip, LinearProgress } from '@mui/material';
import { DataGrid, GridColDef, GridToolbar } from '@mui/x-data-grid';
import DeleteIcon from '@mui/icons-material/Delete';
import DownloadIcon from '@mui/icons-material/Download';
import OpenInNewIcon from '@mui/icons-material/OpenInNew';
import ContentCopyIcon from '@mui/icons-material/ContentCopy';
import XLSX from 'xlsx-js-style';
import ExcelJS from 'exceljs';
import pdfMake from 'pdfmake/build/pdfmake';
import pdfFonts from 'pdfmake/build/vfs_fonts';
try {
  const anyFonts: any = pdfFonts as any;
  (pdfMake as any).vfs = anyFonts?.pdfMake?.vfs || anyFonts?.vfs || (pdfMake as any).vfs || {};
} catch {}
import api from '../api';
import { useToast } from '../components/ToastProvider';

// This page is a lightweight clone of Orders for "Early Buy" workflows.
// It does NOT affect inventory or reservations; all data is saved locally (localStorage).

type EarlyOrder = {
  id: string; // EORD-0001 pattern
  status: 'processing' | 'ready_to_ship' | 'shipped' | 'completed' | 'canceled';
  warehouseId: string; // fixed MPG (display-only)
  createdAt: string; // YYYY-MM-DD
  containerArrival?: string; // YYYY-MM-DD
  estFulfillment: string; // YYYY-MM-DD
  estDelivered: string; // YYYY-MM-DD
  updatedAt?: string;
  updatedBy?: string;
  customerEmail: string;
  customerName: string;
  customerPhone: string;
  shippingAddress: string;
  companyName?: string;
  customerNumber?: string;
  salesRepresentative?: string;
  originalPrice?: string;
  shippingPercent?: string;
  discountPercent?: string;
  paymentTerms?: string;
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
  // Preflight session check on navigation to this page
  useEffect(() => {
    try {
      const token = localStorage.getItem('token');
      if (!token) {
        const path = typeof window !== 'undefined' ? (window.location.pathname || '/') : '/';
        const search = typeof window !== 'undefined' ? (window.location.search || '') : '';
        const next = encodeURIComponent(`${path}${search}`);
        if (typeof window !== 'undefined') window.location.assign(`/login?next=${next}`);
      }
    } catch {}
  }, []);
  const toast = useToast();
  const [orders, setOrders] = useState<EarlyOrder[]>([]);
  const [open, setOpen] = useState(false);
  const [editingOrder, setEditingOrder] = useState<EarlyOrder | null>(null);

  const [palletDescByGroup, setPalletDescByGroup] = useState<Record<string, string>>({});
  const [groupPriceByName, setGroupPriceByName] = useState<Record<string, number>>({});

  // Form state
  const [status, setStatus] = useState<EarlyOrder['status']>('processing');
  const [customerEmail, setCustomerEmail] = useState('');
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [customerNumber, setCustomerNumber] = useState('');
  const [salesRepresentative, setSalesRepresentative] = useState('');
  const salesRepInputRef = useRef<HTMLInputElement | null>(null);
  const [companyName, setCompanyName] = useState('');
  const [createdAt, setCreatedAt] = useState(() => new Date().toISOString().slice(0,10));
  const [containerArrival, setContainerArrival] = useState('');
  const [estFulfillment, setEstFulfillment] = useState('');
  const [estDelivered, setEstDelivered] = useState('');
  const [requestedShip, setRequestedShip] = useState('');
  const [shippingAddress, setShippingAddress] = useState('');
  const [lastUpdatedAt, setLastUpdatedAt] = useState('');
  const [lastUpdatedBy, setLastUpdatedBy] = useState('');
  const [paymentTerms, setPaymentTerms] = useState('Net 60');
  const [paymentStatus, setPaymentStatus] = useState('Not Paid');

  // Customer Browser (uncontrolled search with ref-based debounce for snappy typing)
  const [customerPickOpen, setCustomerPickOpen] = useState(false);
  const [customerRows, setCustomerRows] = useState<any[]>([]);
  const [customerLoading, setCustomerLoading] = useState(false);
  const customerPickDebounceRef = useRef<any>(null);
  const [customerPickQ, setCustomerPickQ] = useState('');
  const loadCustomers = useCallback(async ()=>{
    setCustomerLoading(true);
    try {
      const { data } = await api.get('/customers');
      setCustomerRows(Array.isArray(data) ? data : []);
    } catch { setCustomerRows([]); } finally { setCustomerLoading(false); }
  },[]);

  const isEditable = useMemo(() => {
    // Editable when creating (no editingOrder) OR when status is processing/ready_to_ship
    return !editingOrder || status === 'processing' || status === 'ready_to_ship';
  }, [editingOrder, status]);

  const isArrivalDateEditable = useMemo(() => {
    // Allow editing arrival date even when status is shipped
    return !editingOrder || status === 'processing' || status === 'ready_to_ship' || status === 'shipped';
  }, [editingOrder, status]);

  const isPaymentEditable = useMemo(() => {
    // Allow editing payment fields when shipped as well
    return !editingOrder || status === 'processing' || status === 'ready_to_ship' || status === 'shipped';
  }, [editingOrder, status]);

  const isStatusEditable = useMemo(() => {
    // Once an existing order is COMPLETED, it becomes view-only
    return !editingOrder || (status !== 'completed' && status !== 'canceled');
  }, [editingOrder, status]);

  const isContainerArrivalEditable = useMemo(() => {
    return status === 'processing' && isEditable;
  }, [isEditable, status]);

  // Auto-promote form status: if currently PROCESSING and Container Arrival is today or earlier, set to READY TO SHIP
  useEffect(() => {
    try {
      if (status !== 'processing') return;
      const now = Date.now();
      if (Number(suppressPromoteUntilRef.current || 0) > now) return; // user is adjusting; don't auto-promote yet
      const today = new Date().toISOString().slice(0,10);
      const ca = String(containerArrival || '').slice(0,10);
      if (ca && ca <= today) setStatus('ready_to_ship');
    } catch {}
  }, [status, containerArrival]);

  const [originalPrice, setOriginalPrice] = useState('');
  const [shippingPercent, setShippingPercent] = useState('');
  const [discountPercent, setDiscountPercent] = useState('');
  const [notes, setNotes] = useState('');

  // Uncontrolled inputs (lines grid): refs and error markers
  const qtyInputRefs = useRef<Record<number, HTMLInputElement | null>>({});
  const discountInputRefs = useRef<Record<number, HTMLInputElement | null>>({});
  const [invalidQtyRows, setInvalidQtyRows] = useState<Set<number>>(new Set());
  const [invalidDiscountRows, setInvalidDiscountRows] = useState<Set<number>>(new Set());

  const numberFmt2 = useMemo(() => new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }), []);

  // Report (Early Buy only)
  const [ebReportOpen, setEbReportOpen] = useState(false);
  const [ebReportFrom, setEbReportFrom] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    return d.toISOString().slice(0, 10);
  });
  const [ebReportTo, setEbReportTo] = useState(() => new Date().toISOString().slice(0, 10));
  const [ebReportExporting, setEbReportExporting] = useState(false);

  const exportEarlyBuyReportXlsx = useCallback(async () => {
    const from = String(ebReportFrom || '').slice(0, 10);
    const to = String(ebReportTo || '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
      toast?.error?.('Please select a valid date range');
      return;
    }
    try {
      setEbReportExporting(true);
      // Filter in-memory Early Buy orders by createdAt
      const inRange = (d: string) => {
        const s = String(d || '').slice(0, 10);
        if (!s) return false;
        return (!from || s >= from) && (!to || s <= to);
      };
      const list = (Array.isArray(orders) ? orders : [])
        .filter((o) => inRange(o.createdAt))
        .filter((o) => String(o?.status || '').toLowerCase() !== 'canceled');
      // Aggregate pallets ordered by group (and capture palletId and palletName where available)
      type Agg = { groupKey: string; palletId: string; palletName: string; orderedPallets: number };
      const map = new Map<string, Agg>();
      for (const o of list) {
        const linesArr = Array.isArray(o.lines) ? o.lines : [];
        for (const l of linesArr) {
          const gLower = String(l?.groupName || '').trim().toLowerCase();
          if (!gLower) continue;
          const prev = map.get(gLower) || { groupKey: gLower, palletId: String(l?.lineItem || ''), palletName: String(l?.palletName || ''), orderedPallets: 0 };
          prev.orderedPallets += Math.max(0, Math.floor(Number((l as any)?.qty || 0)));
          if (!prev.palletId) prev.palletId = String(l?.lineItem || '');
          if (!prev.palletName) prev.palletName = String(l?.palletName || '');
          map.set(gLower, prev);
        }
      }
      const agg = Array.from(map.values()).sort((a, b) => b.orderedPallets - a.orderedPallets);

      const rows: any[] = [];
      rows.push(['Early Buy Pallet Orders Report']);
      rows.push([`Date Range: ${from} to ${to}`]);
      try { rows.push([`Exported At: ${new Date().toLocaleString()}`]); } catch { rows.push(['']); }
      rows.push([]);
      rows.push(['Top Ordered Pallets']);
      rows.push(['Pallet ID', 'Pallet Name', 'Pallet Description', 'Pallets Ordered']);
      const top = agg.slice(0, 20);
      for (const r of top) {
        rows.push([
          r.palletId,
          r.palletName,
          String(palletDescByGroup?.[r.groupKey] || ''),
          Number(r.orderedPallets || 0),
        ]);
      }
      if (!top.length) rows.push(['', '', '', 0]);

      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'EarlyBuy');
      const safe = (v: any) => String(v ?? '').replace(/[\/\\:*?"<>|]/g, '-');
      XLSX.writeFile(wb, `early_buy_pallet_orders_${safe(from)}_${safe(to)}.xlsx`);
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to export report';
      toast?.error?.(msg);
    } finally {
      setEbReportExporting(false);
    }
  }, [orders, ebReportFrom, ebReportTo, palletDescByGroup, toast]);

  // When user explicitly sets status to PROCESSING, suppress auto-promote briefly to allow edits
  const suppressPromoteUntilRef = useRef<number>(0);

  const handleStatusChange = useCallback((nextRaw: any) => {
    const extracted = (nextRaw && typeof nextRaw === 'object' && 'target' in nextRaw)
      ? (nextRaw as any)?.target?.value
      : nextRaw;
    const next = String(extracted || '').trim().toLowerCase() as EarlyOrder['status'];
    if (!['processing', 'ready_to_ship', 'shipped', 'completed', 'canceled'].includes(String(next))) return;
    if (next === 'processing') {
      try { suppressPromoteUntilRef.current = Date.now() + 15000; } catch {}
    }
    if (next === 'shipped') {
      const sp = Number(shippingPercent);
      if (!String(shippingPercent || '').trim() || !Number.isFinite(sp) || sp < 0 || sp > 100) {
        toast?.error?.('Shipping Charges (%) is required before setting status to SHIPPED');
        return;
      }
    }
    if (next === 'completed') {
      if (!String(estDelivered || '').trim()) {
        toast?.error?.('Estimated Arrival Date is required before setting status to COMPLETED');
        return;
      }
      const sp = Number(shippingPercent);
      if (!String(shippingPercent || '').trim() || !Number.isFinite(sp) || sp < 0 || sp > 100) {
        toast?.error?.('Shipping Charges (%) is required before setting status to COMPLETED');
        return;
      }

      const ok = window.confirm(
        'Are you sure you want to change the status to COMPLETED?\n\nThis cannot be undone (Disabled from Editing) and you can no longer edit the order.'
      );
      if (!ok) return;
    }
    if (next === 'canceled') {
      const ok = window.confirm(
        'Are you sure you want to change the status to CANCELED?\n\nThis cannot be undone and you can no longer edit the order.'
      );
      if (!ok) return;
    }
    setStatus(next);
  }, [discountPercent, estDelivered, originalPrice, shippingPercent, toast]);

  // moved below lines state

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
  const pickerQDebounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [pickerWarehouses, setPickerWarehouses] = useState<any[]>([]);
  const [selectedGroups, setSelectedGroups] = useState<string[]>([]);

  // Lines added (with snapshots captured at Add time)
  const [lines, setLines] = useState<Array<{
    groupName: string;
    lineItem: string;
    palletName: string;
    qty: number;
    discount?: number;
    unitPriceAtAdd?: number; // snapshot of pallet price when added
    itemsAtAdd?: any[];      // snapshot of items under the pallet when added
  }>>([]);

  // Computed totals based on lines
  const computedSubTotal = useMemo(() => {
    let sum = 0;
    for (let i = 0; i < lines.length; i++) {
      const g = String(lines[i]?.groupName || '').trim().toLowerCase();
      const snap = Number(lines[i]?.unitPriceAtAdd);
      const unit = Number.isFinite(snap) ? snap : Number(groupPriceByName?.[g]);
      if (!Number.isFinite(unit)) continue;
      const disc = Math.max(0, Math.min(100, Number(lines[i]?.discount ?? 0)));
      const qty = Math.max(0, Math.floor(Number(lines[i]?.qty || 0)));
      const discounted = unit * (1 - disc / 100);
      const subtotal = discounted * qty;
      if (Number.isFinite(subtotal)) sum += subtotal;
    }
    return Number.isFinite(sum) ? sum.toFixed(2) : '';
  }, [lines, groupPriceByName]);

  const computedGrandTotal = useMemo(() => {
    const base = Number(computedSubTotal);
    const sp = Number(shippingPercent);
    if (!Number.isFinite(base)) return '';
    const ship = Number.isFinite(sp) ? Math.min(100, Math.max(0, sp)) : 0;
    const out = base * (1 + ship / 100);
    return Number.isFinite(out) ? out.toFixed(2) : '';
  }, [computedSubTotal, shippingPercent]);

  // Orders table filtering
  const [tableQ, setTableQ] = useState('');
  const [tableQDebounced, setTableQDebounced] = useState('');
  const tableQDebounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [tableStatus, setTableStatus] = useState<'all' | EarlyOrder['status']>('all');

  // MUI DataGrid compatibility shim (v5 vs v6 selection APIs)
  const DG: any = DataGrid as any;

  const [palletItemsOpen, setPalletItemsOpen] = useState(false);
  const [palletItemsLoading, setPalletItemsLoading] = useState(false);
  const [palletItemsGroupName, setPalletItemsGroupName] = useState('');
  const [palletItemsRows, setPalletItemsRows] = useState<any[]>([]);

  const safe = (v: any) => String(v ?? '').replace(/[\/\\:*?"<>|]/g, '-');

  const exportEarlyBuyXlsx = useCallback(async () => {
    const orderNumber = String(editingOrder?.id || 'EORD');
    const sStatus = String(status || '').toUpperCase();

    const sub = Number(computedSubTotal);
    const sp = Number(shippingPercent);
    const ship = Number.isFinite(sp) ? Math.min(100, Math.max(0, sp)) : 0;
    const subRounded = Number.isFinite(sub) ? Number(Number(sub).toFixed(2)) : NaN;
    const grand = Number.isFinite(subRounded) ? Number((subRounded * (1 + ship / 100)).toFixed(2)) : NaN;

    const pad = (n: number) => n.toString().padStart(2, '0');
    const d = new Date();
    const ts = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
    const fname = `early-buy-${safe(orderNumber)}-(${safe(sStatus)})-(${ts}).xlsx`;

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Order Details');
    ws.properties.defaultRowHeight = 18;
    ws.columns = [ { width: 28 }, { width: 28 }, { width: 34 }, { width: 20 }, { width: 16 }, { width: 18 } ];

    const bold = { bold: true } as const;
    const lightGreen = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } } as const;
    const lightBlue = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } } as const;
    const lightGrey = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } } as const;
    const darkerGrey = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } } as const;
    const right = { horizontal: 'right' } as const;
    const thinBorder = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } } as const;

    const fetchBase64WithExt = async (url: string) => {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`Failed to fetch ${url}`);
      const ctype = String(res.headers.get('content-type') || '').toLowerCase();
      if (!ctype.includes('image/')) throw new Error(`Not an image: ${url} (${ctype})`);
      const buf = await res.arrayBuffer();
      let binary = ''; const bytes = new Uint8Array(buf);
      for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
      const base64 = btoa(binary);
      let ext: 'png' | 'jpeg' | 'jpg' = 'png';
      if (ctype.includes('jpeg') || ctype.includes('jpg')) ext = 'jpeg';
      else if (ctype.includes('png')) ext = 'png';
      return { base64, ext } as const;
    };
    const fetchImageBase64 = async (paths: string[]) => { let last: any=null; for (const p of paths){ try{return await fetchBase64WithExt(p);}catch(e){last=e;} } throw last||new Error('No image'); };

    // Header with logo + contact (compact A1:A2 like Orders)
    ws.mergeCells('A1:A2');
    try { ws.getColumn(1).width = 26; } catch {}
    try { ws.getRow(1).height = 18; ws.getRow(2).height = 17; } catch {}
    try { (ws.getCell('A1') as any).border = thinBorder as any; } catch {}
    try {
      const logoImg = await fetchImageBase64(['/logo.png','/company-logo.png','/company-logo.jpg','/logo.jpg']);
      const imgId = wb.addImage({ base64: logoImg.base64, extension: logoImg.ext });
      ws.addImage(imgId, { tl: { col: 0, row: 0 }, ext: { width: 221, height: 47 } });
    } catch {}
    ws.getCell('B1').value = 'Phone:'; ws.getCell('B1').font = bold; (ws.getCell('B1') as any).fill = lightGreen; ws.getCell('C1').value = '+1 470-812-0762';
    ws.getCell('D1').value = 'Contact Email:'; ws.getCell('D1').font = bold; (ws.getCell('D1') as any).fill = lightGreen; ws.getCell('E1').value = 'brandon@mpgwholesale.com';
    ws.getCell('B2').value = 'Location:'; ws.getCell('B2').font = bold; (ws.getCell('B2') as any).fill = lightGreen; ws.getCell('C2').value = '145 Industrial Drive, Thomson, GA 30824';
    ws.getCell('D2').value = 'Website:'; ws.getCell('D2').font = bold; (ws.getCell('D2') as any).fill = lightGreen; ws.getCell('E2').value = 'www.mpgwholesale.com';
    ;(['B1','C1','D1','E1','B2','C2','D2','E2'] as const).forEach((addr)=>{ (ws.getCell(addr) as any).border = thinBorder as any; });

    // Pre-export validation: ensure key fields and lines are present
    try {
      const missing: string[] = [];
      if (!String(customerName || '').trim()) missing.push('Customer Name');
      // Company Name and Customer Account allowed to be empty per request
      if (!String(shippingAddress || '').trim()) missing.push('Shipping Address');
      if (!Array.isArray(lines) || lines.length === 0) missing.push('Order Lines');
      if (missing.length) { toast.error(`Missing required data: ${missing.join(', ')}`); return; }

      const badLines = (Array.isArray(lines) ? lines : []).filter((l: any) => {
        const q = Number(l?.qty);
        return !l?.groupName || !Number.isFinite(q) || q <= 0;
      });
      if (badLines.length) { toast.error('Some line items are incomplete (group or quantity). Please review the order lines.'); return; }
    } catch {}

    // Order info (shifted up to row 4-5)
    ws.getCell('A4').value = orderNumber; ws.getCell('A4').font = bold; (ws.getCell('A4') as any).fill = lightBlue; (ws.getCell('A4') as any).border = thinBorder as any;
    ws.getCell('B4').value = 'To:'; ws.getCell('B4').font = bold; (ws.getCell('B4') as any).fill = lightBlue; ws.getCell('C4').value = String(customerName || '');
    ws.getCell('D4').value = 'Phone:'; ws.getCell('D4').font = bold; (ws.getCell('D4') as any).fill = lightBlue; ws.getCell('E4').value = String(customerPhone || '');
    ws.getCell('B5').value = 'Address:'; ws.getCell('B5').font = bold; (ws.getCell('B5') as any).fill = lightBlue; ws.getCell('C5').value = String(shippingAddress || '');
    ws.getCell('D5').value = 'Email:'; ws.getCell('D5').font = bold; (ws.getCell('D5') as any).fill = lightBlue; ws.getCell('E5').value = String(customerEmail || '');
    // Add Company Name and Customer Account
    ws.getCell('B6').value = 'Company Name:'; ws.getCell('B6').font = bold; (ws.getCell('B6') as any).fill = lightBlue;
    ws.getCell('C6').value = String(companyName || '');
    ws.getCell('D6').value = 'Customer Account:'; ws.getCell('D6').font = bold; (ws.getCell('D6') as any).fill = lightBlue;
    ws.getCell('E6').value = String(customerNumber || '');
    ;(['B4','C4','D4','E4','B5','C5','D5','E5','B6','D6'] as const).forEach((addr)=>{ (ws.getCell(addr) as any).border = thinBorder as any; });
    ws.getCell('C6').border = thinBorder as any; ws.getCell('E6').border = thinBorder as any;

    // Removed Status and Estimated date rows per request

    // Grouped pallets
    let currentRow = 11;
    const encodeMoney = (n: any) => (Number.isFinite(Number(n)) ? Number(n) : '');
    const groups = new Map<string, any[]>();
    for (const r of (Array.isArray(lines) ? lines : [])) {
      const g = String(r?.groupName || '').trim(); if (!g) continue; const k = g.toLowerCase(); if (!groups.has(k)) groups.set(k, []); groups.get(k)!.push(r);
    }
    const itemsCache = new Map<string, any[]>();
    const getGroupItems = async (g: string) => {
      const k = g.toLowerCase(); if (itemsCache.has(k)) return itemsCache.get(k)!;
      const snap = (Array.isArray(lines) ? lines : []).find(l => String(l?.groupName||'').trim().toLowerCase()===k && Array.isArray((l as any)?.itemsAtAdd) && (l as any).itemsAtAdd.length);
      if (snap) { const list = (snap as any).itemsAtAdd as any[]; itemsCache.set(k, list); return list; }
      // Fallback to localStorage snapshot captured on add/save
      try {
        const orderId = String(editingOrder?.id || '').trim();
        if (orderId) {
          const cacheKey = `earlyBuyLineSnapshots:${orderId}`;
          const cached = localStorage.getItem(cacheKey);
          if (cached) {
            const parsed = JSON.parse(cached);
            const list = parsed && parsed.items ? (parsed.items[k] || parsed.items[g]) : null;
            if (Array.isArray(list) && list.length) { itemsCache.set(k, list); return list; }
          }
        }
      } catch {}
      try { const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`); const list = Array.isArray((data as any)?.items) ? (data as any).items : []; itemsCache.set(k, list); return list; } catch { itemsCache.set(k, []); return []; }
    };
    const getPalletMeta = (g: string) => {
      const gl = g.toLowerCase(); const hit = (Array.isArray(lines) ? lines : []).find(l=> String(l?.groupName||'').trim().toLowerCase()===gl);
      const palletId = String((hit as any)?.lineItem || '').trim(); const palletName = String((hit as any)?.palletName || g).trim(); return { palletId, palletName };
    };
    for (const [key, groupRows] of Array.from(groups.entries())) {
      const g = String(groupRows[0]?.groupName || '').trim();
      const { palletId, palletName } = getPalletMeta(g);
      const palletDesc = String(palletDescByGroup?.[key] || g || '').trim();
      ws.getRow(currentRow).values = [palletName, palletDesc, `Pallet ID: ${palletId || '-'}`, 'UPC', 'Qty', 'Price'];
      for (let c = 1; c <= 6; c++) { const cell = ws.getRow(currentRow).getCell(c); cell.font = bold; cell.fill = lightGreen; cell.border = thinBorder as any; }
      const items = await getGroupItems(g);
      for (const it of items) {
        currentRow += 1;
        ws.getRow(currentRow).values = ['', String((it as any)?.itemCode || ''), String((it as any)?.description || (it as any)?.name || ''), String((it as any)?.upc || ''), String((it as any)?.packSize || ''), encodeMoney((it as any)?.price)];
        for (let c = 1; c <= 6; c++) { const cell = ws.getRow(currentRow).getCell(c); cell.fill = lightGrey; cell.border = thinBorder as any; }
        ws.getRow(currentRow).getCell(6).numFmt = '"$"#,##0.00'; (ws.getRow(currentRow).getCell(6) as any).alignment = right;
      }
      // Compute fallback unit from items if snapshots are missing
      let fallbackUnitFromItems = 0;
      try {
        for (const it of (Array.isArray(items) ? items : [])) {
          const pack = Number((it as any)?.packSize ?? 0) || 0;
          const price = Number((it as any)?.price ?? 0) || 0;
          fallbackUnitFromItems += pack * price;
        }
      } catch {}

      for (const r of groupRows) {
        const unitBase = (() => {
          const snapAtAdd = Number((r as any)?.unitPriceAtAdd);
          if (Number.isFinite(snapAtAdd)) return snapAtAdd;
          const gp = Number((groupPriceByName as any)?.[key]);
          if (Number.isFinite(gp)) return gp;
          return Number.isFinite(fallbackUnitFromItems) && fallbackUnitFromItems > 0 ? fallbackUnitFromItems : NaN;
        })();
        const disc = Math.min(100, Math.max(0, Number((r as any)?.discount ?? (r as any)?.discountPercent ?? 0)));
        const unitDiscounted = Number.isFinite(unitBase) ? unitBase * (1 - disc/100) : NaN;
        const qty = Math.max(0, Math.floor(Number((r as any)?.qty || 0)));
        const lineSub = (Number.isFinite(unitDiscounted) && Number.isFinite(qty)) ? unitDiscounted * qty : NaN;
        currentRow += 1;
        ws.getRow(currentRow).values = [
          `PALLET PRICE: ${Number.isFinite(unitBase) ? `$${unitBase.toFixed(2)}` : ''}`,
          `DISCOUNT: ${Number.isFinite(disc) ? `${disc}%` : ''}`,
          (disc === 0 ? '' : `DISCOUNTED PRICE: ${Number.isFinite(unitDiscounted) ? `$${unitDiscounted.toFixed(2)}` : ''}`),
          `QUANTITY: ${Number.isFinite(qty) ? qty : ''}`,
          'SUB TOTAL',
          encodeMoney(lineSub),
        ];
        for (let c = 1; c <= 6; c++) { const cell = ws.getRow(currentRow).getCell(c); cell.font = bold; cell.fill = darkerGrey; cell.border = thinBorder as any; }
        ws.getRow(currentRow).getCell(6).numFmt = '"$"#,##0.00'; (ws.getRow(currentRow).getCell(6) as any).alignment = right;
      }
      currentRow += 2;
    }

    ws.getCell(`A${currentRow}`).value = 'SUB TOTAL'; ws.getCell(`A${currentRow}`).font = bold; (ws.getCell(`A${currentRow}`) as any).fill = lightBlue; (ws.getCell(`A${currentRow}`) as any).border = thinBorder as any; ws.getCell(`B${currentRow}`).value = Number.isFinite(subRounded) ? subRounded : ''; ws.getCell(`B${currentRow}`).numFmt = '"$"#,##0.00'; ws.getCell(`B${currentRow}`).font = bold; (ws.getCell(`B${currentRow}`) as any).border = thinBorder as any;
    currentRow += 1;
    const shipAmt = (Number.isFinite(subRounded) && Number.isFinite(sp)) ? Number((subRounded * (sp/100)).toFixed(2)) : NaN;
    const shipLabel = Number.isFinite(sp) ? `SHIPPING (${sp}%)` : 'SHIPPING';
    ws.getCell(`A${currentRow}`).value = shipLabel; ws.getCell(`A${currentRow}`).font = bold; (ws.getCell(`A${currentRow}`) as any).fill = lightBlue; (ws.getCell(`A${currentRow}`) as any).border = thinBorder as any;
    ws.getCell(`B${currentRow}`).value = Number.isFinite(shipAmt) ? shipAmt : ''; ws.getCell(`B${currentRow}`).numFmt = '"$"#,##0.00'; ws.getCell(`B${currentRow}`).font = bold; (ws.getCell(`B${currentRow}`) as any).border = thinBorder as any;
    currentRow += 1;
    ws.getCell(`A${currentRow}`).value = 'GRAND TOTAL'; ws.getCell(`A${currentRow}`).font = bold; (ws.getCell(`A${currentRow}`) as any).fill = lightBlue; (ws.getCell(`A${currentRow}`) as any).border = thinBorder as any; ws.getCell(`B${currentRow}`).value = Number.isFinite(grand) ? grand : ''; ws.getCell(`B${currentRow}`).numFmt = '"$"#,##0.00'; ws.getCell(`B${currentRow}`).font = bold; (ws.getCell(`B${currentRow}`) as any).border = thinBorder as any;
    currentRow += 1;
    ws.getCell(`A${currentRow}`).value = 'PAYMENT TERMS'; ws.getCell(`A${currentRow}`).font = bold; (ws.getCell(`A${currentRow}`) as any).fill = lightBlue; (ws.getCell(`A${currentRow}`) as any).border = thinBorder as any; ws.getCell(`B${currentRow}`).value = String(paymentTerms || ''); ws.getCell(`B${currentRow}`).font = bold; (ws.getCell(`B${currentRow}`) as any).border = thinBorder as any;
    const totalsEndRow = currentRow;

    // Order Authorization block (left-aligned), similar to Orders export
    currentRow += 2;
    const authStartRow = currentRow;
    ws.mergeCells(`D${currentRow}:F${currentRow}`);
    ws.getCell(`D${currentRow}`).value = 'Order Authorization';
    ws.getCell(`D${currentRow}`).font = { bold: true } as any;
    currentRow += 1;
    ws.mergeCells(`D${currentRow}:F${currentRow}`);
    ws.getCell(`D${currentRow}`).value = 'By signing below, I authorize this order and agree to the pricing, shipping charges, and payment terms outlined in this Proforma Invoice.';
    (ws.getCell(`D${currentRow}`) as any).alignment = { wrapText: true } as any;
    currentRow += 1; ws.mergeCells(`D${currentRow}:F${currentRow}`); ws.getCell(`D${currentRow}`).value = 'Authorized Signature: ______________________________________';
    currentRow += 1; ws.mergeCells(`D${currentRow}:F${currentRow}`); ws.getCell(`D${currentRow}`).value = 'Printed Name: _____________________________________________';
    currentRow += 1; ws.mergeCells(`D${currentRow}:F${currentRow}`); ws.getCell(`D${currentRow}`).value = 'Date: _____/_____/_______';
    const authEndRow = currentRow;

    // Center align most cells, then left-align authorization block
    try {
      ws.eachRow({ includeEmpty: true }, (row: any) => {
        row.eachCell({ includeEmpty: true }, (cell: any) => {
          const prev = (cell.alignment || {}) as any;
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: Boolean(prev.wrapText) } as any;
        });
      });
    } catch {}
    try {
      for (let r = authStartRow; r <= authEndRow; r++) {
        (ws.getCell(`D${r}`) as any).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true } as any;
        (ws.getCell(`E${r}`) as any).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true } as any;
        (ws.getCell(`F${r}`) as any).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true } as any;
      }
    } catch {}

    // Social section: place below authorization block (no extra signature line)
    const signatureRow = authEndRow;
    currentRow = signatureRow + 7; ws.getCell(`E${currentRow}`).value = 'Follow us on social media !'; (ws.getCell(`E${currentRow}`) as any).alignment = { horizontal: 'left' } as any;
    try {
      const socialImg = await fetchImageBase64(['/socialmedia-qrcode.png']); const idSocial = wb.addImage({ base64: socialImg.base64, extension: socialImg.ext }); const imgRow = currentRow + 1; ws.addImage(idSocial, { tl: { col: 4.2, row: imgRow }, ext: { width: 130, height: 60 } });
      try { const bigLogo = await fetchImageBase64(['/logo.png','/company-logo.png','/company-logo.jpg','/logo.jpg']); const idBigLogo = wb.addImage({ base64: bigLogo.base64, extension: bigLogo.ext }); ws.addImage(idBigLogo, { tl: { col: 1, row: imgRow - 1 }, ext: { width: 546, height: 96 } }); } catch {}
    } catch {}

    // Center all text across the sheet
    try {
      ws.eachRow({ includeEmpty: false }, (row: any) => {
        row.eachCell({ includeEmpty: false }, (cell: any) => {
          const prev = cell.alignment || {};
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: Boolean((prev as any).wrapText) } as any;
        });
      });
    } catch {}

    const buffer = await wb.xlsx.writeBuffer(); const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = fname; a.click(); setTimeout(() => URL.revokeObjectURL(url), 4000);
  }, [editingOrder, status, customerName, customerPhone, shippingAddress, customerEmail, estFulfillment, estDelivered, lines, groupPriceByName, palletDescByGroup, computedSubTotal, shippingPercent, paymentTerms]);

  const exportEarlyBuyPdf = useCallback(async () => {
    const orderNo = safe(String(editingOrder?.id || 'EORD'));
    const orderNumber = String(editingOrder?.id || '').trim();
    const number = (n: any) => {
      const v = Number(n);
      return Number.isFinite(v) ? numberFmt2.format(v) : '';
    };

    // Brief toast while we hydrate details needed for the export
    try { toast.info('Preparing export…'); } catch {}

    // Group duplicate lines so pallet header/items render once, with multiple summary rows
    const groupLinesMap = new Map<string, { display: string; palletName: string; lineItem: string; lines: Array<{ discount: number; qty: number }> }>();
    for (const l of lines) {
      const key = String(l?.groupName || '').trim().toLowerCase();
      if (!key) continue;
      const display = String(l?.groupName || '');
      const palletName = String(l?.palletName || '');
      const lineItem = String(l?.lineItem || '');
      const discount = Math.max(0, Math.min(100, Number(l?.discount ?? 0)));
      const qty = Math.max(0, Math.floor(Number(l?.qty || 0)));
      if (!groupLinesMap.has(key)) groupLinesMap.set(key, { display, palletName, lineItem, lines: [] });
      groupLinesMap.get(key)!.lines.push({ discount, qty });
    }

    const uniqueDisplays = Array.from(new Set(Array.from(groupLinesMap.values()).map(g => g.display)));
    const itemsByGroup: Record<string, any[]> = {};
    await Promise.all(uniqueDisplays.map(async (g) => {
      const gLower = String(g || '').toLowerCase();
      // Prefer snapshot from current lines if available
      const snap = (Array.isArray(lines) ? lines : []).find(l => String(l?.groupName||'').trim().toLowerCase() === gLower && Array.isArray(l?.itemsAtAdd) && l.itemsAtAdd.length);
      if (snap) {
        itemsByGroup[gLower] = snap.itemsAtAdd as any[];
        return;
      }
      // Try fetch with a quick retry window
      const fetchWithRetry = async () => {
        for (let i = 0; i < 2; i++) {
          try {
            const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
            itemsByGroup[gLower] = Array.isArray((data as any)?.items) ? (data as any).items : [];
            return;
          } catch {
            itemsByGroup[gLower] = [];
          }
          if (i === 0) { try { await new Promise(res => setTimeout(res, 150)); } catch {} }
        }
      };
      await fetchWithRetry();
    }));

    // Validate: for each group, ensure we have at least pallet meta (line item/name) or nonzero quantity
    try {
      let invalid = false;
      for (const [key, meta] of Array.from(groupLinesMap.entries())) {
        const hasQty = Array.isArray(meta.lines) && meta.lines.some(x => Number(x.qty) > 0);
        const hasMeta = Boolean(String(meta?.lineItem || meta?.palletName || '').trim());
        if (!hasQty && !hasMeta) { invalid = true; break; }
      }
      if (invalid) {
        try { toast.error('Some pallet details are still loading. Please try export again in a moment.'); } catch {}
        return;
      }
    } catch {}

    // Colors and widths matching Orders PDF
    const blue = '#DDEBF7';
    const green = '#CCFFCC';
    const lightGrey = '#F2F2F2';
    const darkGrey = '#BFBFBF';
    const palletColWidths: any[] = [160, 120, 210, 100, 50, 70];

    // Helper to fetch image as data URI for pdfMake
    const fetchBase64 = async (url: string) => {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`Failed to fetch ${url}`);
      const ctype = String(res.headers.get('content-type') || '').toLowerCase();
      const buf = await res.arrayBuffer();
      let binary = '';
      const bytes = new Uint8Array(buf);
      for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
      const prefix = ctype.includes('jpeg') || ctype.includes('jpg') ? 'data:image/jpeg;base64,' : 'data:image/png;base64,';
      return `${prefix}${btoa(binary)}`;
    };

    const content: any[] = [];
    // Header with logo and company info (fallback without logo if fetch fails)
    try {
      const logo = await fetchBase64('/company-logo.png');
      const companyPhone = '+1 470-812-0762';
      const companyEmail = 'brandon@mpgwholesale.com';
      const companyLocation = '145 Industrial Drive, Thomson, GA 30824';
      const companyWebsite = 'www.mpgwholesale.com';
      content.push({
        table: {
          widths: [130, 100, 260, 100, 150],
          body: [
            [
              { image: logo, height:35, width:135, rowSpan: 2, border: [true, true, true, true], alignment: 'center' },
              { text: 'Phone:', fillColor: green, bold: true, border: [true, true, true, true] },
              { text: companyPhone, border: [true, true, true, true] },
              { text: 'Contact Email:', fillColor: green, bold: true, border: [true, true, true, true] },
              { text: companyEmail, border: [true, true, true, true] },
            ],
            [
              {},
              { text: 'Location:', fillColor: green, bold: true, border: [true, true, true, true] },
              { text: companyLocation, border: [true, true, true, true] },
              { text: 'Website:', fillColor: green, bold: true, border: [true, true, true, true] },
              { text: companyWebsite, border: [true, true, true, true] },
            ],
             [
              { text: '', colSpan: 5, margin: [-4, -4], border: [false, false, false, false] }, {}, {}, {}, {}
            ],
            [
              { text: orderNumber, fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: 'To:', fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: customerName || '', border: [true, true, true, true] },
              { text: 'Phone:', fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: customerPhone || '', border: [true, true, true, true] },
            ],
            [
              {},
              { text: 'Address:', fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: shippingAddress || '', border: [true, true, true, true] },
              { text: 'Email:', fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: customerEmail || '', border: [true, true, true, true] },
            ],
            [
              {},
              { text: 'Company Name:', fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: companyName || '', border: [true, true, true, true] },
              { text: 'Customer Account:', fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: customerNumber || '', border: [true, true, true, true] },
            ],
          ],
        },
        layout: 'noHorizontalLines',
        margin: [0, 0, 0, 8],
      });
    } catch {
      // Fallback header without logo
      content.push({
        table: {
          widths: [160, 40, '*', 55, 120],
          body: [
            [ { text: orderNumber, fillColor: blue, bold: true, border: [true, true, true, true] }, { text: 'To:', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: customerName || '', border: [true, true, true, true] }, { text: 'Phone:', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: customerPhone || '', border: [true, true, true, true] } ],
            [ { text: '', border: [true, false, true, true] }, { text: 'Address:', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: shippingAddress || '', border: [true, true, true, true] }, { text: 'Email:', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: customerEmail || '', border: [true, true, true, true] } ],
            [ { text: '', border: [true, false, true, true] }, { text: 'Company Name:', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: companyName || '', border: [true, true, true, true] }, { text: 'Customer Account:', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: customerNumber || '', border: [true, true, true, true] } ],
          ],
        },
        layout: 'noHorizontalLines',
        margin: [0, 0, 0, 8],
      });
    }

    // Removed Status and Estimated date rows per request

    // Per-pallet sections
    for (const [gKey, gMeta] of groupLinesMap.entries()) {
      const unitSnap = Number((lines.find(l => String(l?.groupName || '').trim().toLowerCase() === gKey) as any)?.unitPriceAtAdd);
      const unit0 = Number.isFinite(unitSnap) ? unitSnap : Number(groupPriceByName?.[gKey] ?? 0);

      content.push({
        table: {
          widths: palletColWidths,
          body: [[
            { text: gMeta.palletName || gMeta.display, fillColor: green, bold: true, border: [true, true, true, true] },
            { text: String(palletDescByGroup?.[gKey] || gMeta.display || ''), fillColor: green, bold: true, border: [true, true, true, true] },
            { text: `Pallet ID: ${gMeta.lineItem || '-'}`, fillColor: green, bold: true, border: [true, true, true, true] },
            { text: 'UPC', fillColor: green, bold: true, border: [true, true, true, true] },
            { text: 'Qty', fillColor: green, bold: true, border: [true, true, true, true] },
            { text: 'Price', fillColor: green, bold: true, border: [true, true, true, true] },
          ]],
        },
        layout: 'noHorizontalLines',
      });

      const items = itemsByGroup[gKey] || [];
      if (items.length) {
        const itemRows = items.map((it: any) => ([
          { text: '', fillColor: lightGrey, border: [true, true, true, true] },
          { text: String(it?.itemCode || ''), fillColor: lightGrey, border: [true, true, true, true] },
          { text: String(it?.description || ''), fillColor: lightGrey, border: [true, true, true, true] },
          { text: String(it?.upc || ''), alignment: 'center', fillColor: lightGrey, border: [true, true, true, true] },
          { text: (Number.isFinite(Number(it?.packSize)) ? String(Number(it?.packSize)) : ''), alignment: 'center', fillColor: lightGrey, border: [true, true, true, true] },
          { text: (Number.isFinite(Number(it?.price)) ? `$${Number(it?.price).toFixed(2)}` : ''), alignment: 'center', fillColor: lightGrey, border: [true, true, true, true] },
        ]));
        content.push({ table: { widths: palletColWidths, body: itemRows }, layout: 'noHorizontalLines' });
      }

      for (const ln of gMeta.lines) {
        const unitBase = unit0;
        const disc = Math.min(100, Math.max(0, Number(ln.discount)));
        const unitDisc = Number.isFinite(unitBase) ? unitBase * (1 - disc / 100) : NaN;
        const qtyEach = Math.max(0, Math.floor(Number(ln.qty || 0)));
        const lineSubEach = (Number.isFinite(unitDisc) && Number.isFinite(qtyEach)) ? unitDisc * qtyEach : NaN;
        content.push({
          table: {
            widths: palletColWidths,
            body: [[
              { text: `PALLET PRICE: ${Number.isFinite(unitBase) ? `$${unitBase.toFixed(2)}` : ''}`, fillColor: darkGrey, bold: true, border: [true, true, true, true] },
              { text: `DISCOUNT: ${Number.isFinite(disc) ? `${disc}%` : ''}`, fillColor: darkGrey, bold: true, border: [true, true, true, true] },
              { text: (disc === 0 ? '' : `DISCOUNTED PRICE: ${Number.isFinite(unitDisc) ? `$${unitDisc.toFixed(2)}` : ''}`), fillColor: darkGrey, bold: true, border: [true, true, true, true] },
              { text: `QUANTITY: ${Number.isFinite(qtyEach) ? qtyEach : ''}`, fillColor: darkGrey, bold: true, border: [true, true, true, true] },
              { text: 'SUB TOTAL', fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: Number.isFinite(lineSubEach) ? `$${lineSubEach.toFixed(2)}` : '', alignment: 'center', fillColor: blue, border: [true, true, true, true] },
            ]],
          },
          layout: 'noHorizontalLines',
          margin: [0, 2, 0, 4],
        });
      }
      content.push({ text: ' ', margin: [0, 2, 0, 2] });
    }

    // Totals + Authorization (two-column section to match Orders PDF)
    const subN = Number(computedSubTotal);
    const subRounded = Number.isFinite(subN) ? Number(Number(subN).toFixed(2)) : NaN;
    const sp = Number(shippingPercent);
    const ship = Number.isFinite(sp) ? sp : NaN;
    const shipAmt = (Number.isFinite(subRounded) && Number.isFinite(sp)) ? Number((subRounded * (sp / 100)).toFixed(2)) : NaN;
    const shipLabel = Number.isFinite(sp) ? `SHIPPING (${sp}%)` : 'SHIPPING';
    const grandN = Number(computedGrandTotal);
    const grandRounded = Number.isFinite(grandN) ? Number(Number(grandN).toFixed(2)) : NaN;
    content.push({
      table: {
        widths: ['*', '*'],
        body: [[
          {
            table: {
              widths: [120, 140],
              body: [
                [ { text: 'SUB TOTAL', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: Number.isFinite(subRounded) ? `$${subRounded.toFixed(2)}` : '', alignment: 'center', bold: true, border: [true, true, true, true] } ],
                [ { text: shipLabel, fillColor: blue, bold: true, border: [true, true, true, true] }, { text: Number.isFinite(shipAmt) ? `$${shipAmt.toFixed(2)}` : '', alignment: 'center', bold: true, border: [true, true, true, true] } ],
                [ { text: 'GRAND TOTAL', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: Number.isFinite(grandRounded) ? `$${grandRounded.toFixed(2)}` : '', alignment: 'center', bold: true, border: [true, true, true, true] } ],
                [ { text: 'PAYMENT TERMS', fillColor: blue, bold: true, border: [true, true, true, true] }, { text: String(paymentTerms || ''), alignment: 'center', bold: true, border: [true, true, true, true] } ],
              ],
            },
            layout: 'noHorizontalLines',
            border: [false, false, false, false],
            margin: [0, 6, 12, 0],
          },
          {
            stack: [
              { text: 'Order Authorization', bold: true, fontSize: 14, alignment: 'left', margin: [0, 0, 0, 8] },
              { text: 'By signing below, I authorize this order and agree to the pricing, shipping charges, and payment terms outlined in this Proforma Invoice.', bold: true, fontSize: 11, alignment: 'left', margin: [0, 0, 0, 12] },
              { text: 'Authorized Signature: ______________________________________', bold: true, fontSize: 11, alignment: 'left', margin: [0, 0, 0, 10] },
              { text: 'Printed Name: _____________________________________________', bold: true, fontSize: 11, alignment: 'left', margin: [0, 0, 0, 10] },
              { text: 'Date: _____/_____/_______', bold: true, fontSize: 11, alignment: 'left' },
            ],
            border: [false, false, false, false],
          }
        ]],
      },
      layout: 'noBorders',
      margin: [0, 6, 0, 0],
    });

    // Spacer to social section
    content.push({ text: ' ', margin: [0, 70, 0, 0] });
    try {
      const social = await fetchBase64('/socialmedia-qrcode.png');
      const bottomLogo = await fetchBase64('/company-logo.png');
      content.push({
        table: {
          widths: ['*','auto','*','*','auto','*'],
          body: [[
            { text: '', border: [false, false, false, false] },
            { image: bottomLogo, fit: [410, 72], alignment: 'left', border: [false, false, false, false] },
            { text: '', border: [false, false, false, false] },
            { text: '', border: [false, false, false, false] },
            { stack: [ { text: 'Follow us on social media !', alignment: 'center', margin: [0, 0, 0, 4] }, { image: social, fit: [120, 60], alignment: 'center' } ], border: [false, false, false, false] },
            { text: '', border: [false, false, false, false] },
          ]],
        },
        layout: 'noHorizontalLines',
      });
    } catch {}

    const docDefinition = { pageOrientation: 'landscape', pageMargins: [20,20,20,20], content, defaultStyle: { fontSize: 9, alignment: 'center' } } as any;
    const fname = `early-buy-${orderNo}.pdf`;
    (pdfMake as any).createPdf(docDefinition).download(fname);
  }, [editingOrder, lines, customerName, customerPhone, shippingAddress, customerEmail, companyName, customerNumber, groupPriceByName, palletDescByGroup, computedSubTotal, shippingPercent, computedGrandTotal, numberFmt2, paymentTerms]);

  useEffect(() => {
    try {
      const cached = localStorage.getItem('palletDescByGroup');
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
          setPalletDescByGroup(parsed as Record<string, string>);
        }
      }
    } catch {}

    let canceled = false;
    (async () => {
      try {
        const { data } = await api.get<any[]>('/item-groups');
        if (canceled) return;
        const descMap: Record<string, string> = {};
        for (const g of (Array.isArray(data) ? data : [])) {
          const name = String((g as any)?.name || '').trim().toLowerCase();
          if (!name) continue;
          descMap[name] = String((g as any)?.palletDescription || '').trim();
        }
        setPalletDescByGroup(descMap);
        try { localStorage.setItem('palletDescByGroup', JSON.stringify(descMap)); } catch {}
      } catch {
        if (!canceled) setPalletDescByGroup((m) => (m && Object.keys(m).length ? m : {}));
      }
    })();
    return () => { canceled = true; };
  }, []);

  // Compute Pallet Price per group from Items (mirrors Orders page logic)
  useEffect(() => {
    let canceled = false;
    (async () => {
      try {
        const [groupsRes, itemsRes] = await Promise.all([
          api.get<any[]>('/item-groups'),
          api.get<any[]>('/items', { params: { includeDisabled: 1 } }),
        ]);
        if (canceled) return;
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
        setGroupPriceByName(map);
      } catch {
        if (!canceled) setGroupPriceByName({});
      }
    })();
    return () => { canceled = true; };
  }, []);

  const openPalletItems = useCallback(async ({ groupName }: { groupName: string }) => {
    const g = String(groupName || '').trim();
    if (!g) return;
    setPalletItemsOpen(true);
    setPalletItemsGroupName(g);
    setPalletItemsRows([]);
    setPalletItemsLoading(true);
    try {
      // Prefer snapshot from current lines if available
      const snap = (Array.isArray(lines) ? lines : []).find((l)=> String(l?.groupName||'').trim().toLowerCase() === g.toLowerCase() && Array.isArray(l.itemsAtAdd) && l.itemsAtAdd.length);
      if (snap) {
        setPalletItemsRows(snap.itemsAtAdd as any[]);
      } else {
        // Fallback to local cache by order id
        let fromCache: any[] | null = null;
        try {
          const orderId = String(editingOrder?.id || '').trim();
          if (orderId) {
            const cacheKey = `earlyBuyLineSnapshots:${orderId}`;
            const cached = localStorage.getItem(cacheKey);
            if (cached) {
              const parsed = JSON.parse(cached);
              const items = parsed && parsed.items ? parsed.items[g.toLowerCase()] || parsed.items[g] : null;
              if (Array.isArray(items) && items.length) fromCache = items as any[];
            }
          }
        } catch {}
        if (Array.isArray(fromCache) && fromCache.length) {
          setPalletItemsRows(fromCache);
        } else {
          const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
          setPalletItemsRows(Array.isArray((data as any)?.items) ? (data as any).items : []);
        }
      }
    } catch {
      setPalletItemsRows([]);
    } finally {
      setPalletItemsLoading(false);
    }
  }, [lines]);

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

  // Clear pending search debounce when dialog closes
  useEffect(() => {
    if (!pickerOpen && pickerQDebounceRef.current) {
      clearTimeout(pickerQDebounceRef.current);
      pickerQDebounceRef.current = null;
    }
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
      const pdesc = String(palletDescByGroup?.[gnameLower] || '').trim().toLowerCase();
      return gid.includes(q) || gnameLower.includes(q) || pname.includes(q) || pdesc.includes(q);
    });
  }, [pickerRows, pickerQDebounced, palletDescByGroup]);

  const addSelectedToLines = async () => {
    const rows = Array.isArray(pickerRows) ? pickerRows : [];
    const sel = new Set(selectedGroups.map((g) => String(g)));
    const chosen: Array<{ groupName: string; lineItem: string; palletName: string }> = [];
    for (const r of rows) {
      const gName = String(r?.groupName || '');
      const lItem = String(r?.lineItem || '');
      const composedId = `${lItem.trim()}::${gName.trim()}`;
      if (!sel.has(composedId) && !sel.has(lItem) && !sel.has(gName)) continue;
      chosen.push({ groupName: gName, lineItem: lItem, palletName: String(r?.palletName || '') });
    }
    if (!chosen.length) { setPickerOpen(false); setSelectedGroups([]); return; }
    const existing = new Set(lines.map(l => String(l.groupName).toLowerCase()));
    const uniqueNew = chosen.filter(a => !existing.has(a.groupName.toLowerCase()));
    const enriched = await Promise.all(uniqueNew.map(async (a) => {
      const key = a.groupName.toLowerCase();
      // Snapshot current unit price
      const unit = Number(groupPriceByName?.[key]);
      // Snapshot items list
      let itemsSnap: any[] = [];
      try {
        const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(a.groupName)}`);
        itemsSnap = Array.isArray((data as any)?.items) ? (data as any).items : [];
      } catch {}
      // Fallback unit price from items if group map not ready
      let fallbackUnit = 0;
      try {
        for (const it of (Array.isArray(itemsSnap) ? itemsSnap : [])) {
          const pack = Number((it as any)?.packSize ?? 0) || 0;
          const price = Number((it as any)?.price ?? 0) || 0;
          fallbackUnit += pack * price;
        }
      } catch {}
      const unitAtAdd = Number.isFinite(unit) ? unit : (Number.isFinite(fallbackUnit) && fallbackUnit > 0 ? fallbackUnit : undefined);
      // Default Order Qty to 1 for newly added pallets
      return { ...a, qty: 1, unitPriceAtAdd: unitAtAdd, itemsAtAdd: itemsSnap } as any;
    }));
    setLines(prev => [...prev, ...enriched]);
    try {
      const orderId = String(editingOrder?.id || '').trim();
      if (orderId) {
        const cacheKey = `earlyBuyLineSnapshots:${orderId}`;
        const cached = localStorage.getItem(cacheKey);
        const parsed = cached ? JSON.parse(cached) : {};
        if (parsed && typeof parsed === 'object') {
          parsed.unit = parsed.unit && typeof parsed.unit === 'object' ? parsed.unit : {};
          parsed.items = parsed.items && typeof parsed.items === 'object' ? parsed.items : {};
          for (const e of enriched) {
            const k = String(e?.groupName || '').trim().toLowerCase();
            if (!k) continue;
            if (Number.isFinite(Number(e?.unitPriceAtAdd))) parsed.unit[k] = Number(e.unitPriceAtAdd);
            if (Array.isArray(e?.itemsAtAdd)) parsed.items[k] = e.itemsAtAdd;
          }
          localStorage.setItem(cacheKey, JSON.stringify(parsed));
        } else {
          const unit: Record<string, number> = {};
          const items: Record<string, any[]> = {};
          for (const e of enriched) {
            const k = String(e?.groupName || '').trim().toLowerCase();
            if (!k) continue;
            if (Number.isFinite(Number(e?.unitPriceAtAdd))) unit[k] = Number(e.unitPriceAtAdd);
            if (Array.isArray(e?.itemsAtAdd)) items[k] = e.itemsAtAdd;
          }
          localStorage.setItem(cacheKey, JSON.stringify({ unit, items }));
        }
      }
    } catch {}
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
    { field: 'companyName', headerName: 'Company Name', width: 200 },
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
      field: 'finalPrice',
      headerName: 'Grand Total',
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
              try { setSalesRepresentative(String((r as any)?.salesRepresentative || '')); } catch {}
              try { setCustomerNumber(String((r as any)?.customerNumber || '')); } catch {}
              try { setCompanyName(String((r as any)?.companyName || '')); } catch {}
              setCreatedAt(r.createdAt);
              setEstFulfillment(r.estFulfillment);
              setEstDelivered(r.estDelivered);
              setOriginalPrice(String(r.originalPrice||''));
              setShippingPercent(String(r.shippingPercent||''));
              setDiscountPercent(String(r.discountPercent||''));
              setNotes(String(r.notes||''));
              // Build lines with persisted snapshots preferred
              try {
                let base = Array.isArray(r.lines) ? r.lines.map((l: any) => ({ ...l })) : [];
                const orderId = String(r?.id || '').trim();
                if (orderId) {
                  const cacheKey = `earlyBuyLineSnapshots:${orderId}`;
                  const cached = localStorage.getItem(cacheKey);
                  if (cached) {
                    const parsed = JSON.parse(cached);
                    if (parsed && typeof parsed === 'object') {
                      const unit = (parsed as any).unit || {};
                      const items = (parsed as any).items || {};
                      base = base.map((l: any) => {
                        const k = String(l?.groupName || '').trim().toLowerCase();
                        const u = Number((unit || {})[k]);
                        const it = (items || {})[k];
                        return {
                          ...l,
                          ...(Number.isFinite(u) ? { unitPriceAtAdd: u } : {}),
                          ...(Array.isArray(it) && it.length ? { itemsAtAdd: it } : {}),
                        } as any;
                      });
                    }
                  }
                }
                setLines(base);
              } catch { setLines(Array.isArray(r.lines) ? r.lines.map((l: any) => ({ ...l })) : []); }
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
    const q = String(tableQDebounced || '').trim().toLowerCase();
    const st = tableStatus;
    return (orders || []).filter((o) => {
      const statusOk = st === 'all' ? true : normalizeStatus(o.status) === st;
      if (!q) return statusOk;
      const hay = `${o.id} ${o.customerEmail} ${o.customerName} ${o.companyName || ''}`.toLowerCase();
      return statusOk && hay.includes(q);
    });
  }, [orders, tableQDebounced, tableStatus]);

  const linesColumns: GridColDef[] = useMemo(() => [
    { field: 'palletName', headerName: 'Pallet Name', width: 220 },
    { field: 'groupName', headerName: 'Pallet Description', width: 260, renderCell: (p: any) => {
      const row: any = p?.row || {};
      const gLower = String(row?.groupName || '').trim().toLowerCase();
      return String(palletDescByGroup?.[gLower] || row?.groupName || '');
    } },
    { field: 'lineItem', headerName: 'Pallet ID', width: 160 },
    { field: 'palletPrice', headerName: 'Pallet Price', width: 110, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => {
      const row: any = p?.row || {};
      const gLower = String(row?.groupName || '').trim().toLowerCase();
      const snap = Number(row?.unitPriceAtAdd);
      const val = Number.isFinite(snap) ? snap : Number(groupPriceByName?.[gLower]);
      if (!Number.isFinite(val)) return '';
      return numberFmt2.format(val);
    } },
    { field: 'discount', headerName: 'Discount (%)', width: 120, align: 'right', headerAlign: 'right', renderCell: (p: any) => {
      const row: any = p?.row || {};
      const v = Math.max(0, Math.min(100, Number(row?.discount ?? 0)));
      return (
        <TextField
          size="small"
          type="number"
          defaultValue={v}
          onFocus={(e)=> { try { (e.target as HTMLInputElement).select(); } catch {} }}
          onBlur={(e) => {
            const n = Math.max(0, Math.min(100, Math.floor(Number(e.target.value) || 0)));
            const next = [...lines];
            const i = Math.max(0, Number(p?.id) - 1);
            if (next[i]) next[i] = { ...next[i], discount: n } as any;
            setLines(next);
          }}
          inputProps={{ inputMode: 'numeric', pattern: '[0-9]*', min: 0, max: 100 }}
          inputRef={(el)=> { const i = Math.max(0, Number(p?.id) - 1); discountInputRefs.current[i] = el; }}
          error={invalidDiscountRows.has(Math.max(0, Number(p?.id) - 1))}
          helperText={invalidDiscountRows.has(Math.max(0, Number(p?.id) - 1)) ? '0-100 required' : undefined}
          sx={{ '& input': { textAlign: 'right' } }}
        />
      );
    } },
    { field: 'discountedPrice', headerName: 'Discounted Price', width: 140, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => {
      const row: any = p?.row || {};
      const gLower = String(row?.groupName || '').trim().toLowerCase();
      const snap = Number(row?.unitPriceAtAdd);
      const unit = Number.isFinite(snap) ? snap : Number(groupPriceByName?.[gLower]);
      if (!Number.isFinite(unit)) return '';
      const disc = Math.max(0, Math.min(100, Number(row?.discount ?? 0)));
      const discounted = unit * (1 - disc / 100);
      return Number.isFinite(discounted) ? numberFmt2.format(discounted) : '';
    } },
    {
      field: 'qty',
      headerName: 'Order Qty',
      width: 120,
      align: 'right',
      headerAlign: 'right',
      renderCell: (p: any) => {
        const row: any = p?.row || {};
        const i = Math.max(0, Number(p?.id) - 1);
        const val = Number(row?.qty ?? 0);
        return (
          <TextField
            size="small"
            type="number"
            defaultValue={val}
            onFocus={(e)=> { try { (e.target as HTMLInputElement).select(); } catch {} }}
            onBlur={(e) => {
              const v = Math.max(0, Math.floor(Number(e.target.value) || 0));
              const next = [...lines];
              if (next[i]) next[i] = { ...next[i], qty: v };
              setLines(next);
            }}
            inputProps={{ inputMode: 'numeric', pattern: '[0-9]*', min: 0 }}
            inputRef={(el)=> { qtyInputRefs.current[i] = el; }}
            error={invalidQtyRows.has(i)}
            helperText={invalidQtyRows.has(i) ? 'Qty must be > 0' : undefined}
            sx={{ '& input': { textAlign: 'right' } }}
          />
        );
      },
    },
    { field: 'subTotal', headerName: 'Sub Total', width: 120, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => {
      const row: any = p?.row || {};
      const gLower = String(row?.groupName || '').trim().toLowerCase();
      const snap = Number(row?.unitPriceAtAdd);
      const unit = Number.isFinite(snap) ? snap : Number(groupPriceByName?.[gLower]);
      if (!Number.isFinite(unit)) return '';
      const disc = Math.max(0, Math.min(100, Number(row?.discount ?? 0)));
      const qty = Math.max(0, Math.floor(Number(row?.qty || 0)));
      const discounted = unit * (1 - disc / 100);
      const subtotal = discounted * qty;
      return Number.isFinite(subtotal) ? numberFmt2.format(subtotal) : '';
    } },
    {
      field: 'actions',
      headerName: '',
      width: 110,
      sortable: false,
      filterable: false,
      align: 'center',
      renderCell: (p: any) => {
        const idx = Math.max(0, Number(p?.id) - 1);
        return (
          <Stack direction="row" spacing={0.5} alignItems="center">
            <IconButton size="small" onClick={() => {
              const src = lines[idx];
              if (!src) return;
              const dup = { ...src, qty: 0 } as any;
              const next = [...lines];
              next.splice(idx + 1, 0, dup);
              setLines(next);
            }} title="Duplicate">
              <ContentCopyIcon fontSize="small" />
            </IconButton>
            <IconButton size="small" color="error" onClick={() => {
              const next = [...lines];
              next.splice(idx, 1);
              // If this was the last instance of the group, clear snapshots from localStorage cache
              try {
                const removed = lines[idx];
                const gTrim = String(removed?.groupName || '').trim();
                const gLower = gTrim.toLowerCase();
                const remaining = next.filter(l => String(l?.groupName || '').trim().toLowerCase() === gLower).length;
                if (remaining <= 0) {
                  const orderId = String(editingOrder?.id || '').trim();
                  if (orderId) {
                    const cacheKey = `earlyBuyLineSnapshots:${orderId}`;
                    const cached = localStorage.getItem(cacheKey);
                    if (cached) {
                      const parsed = JSON.parse(cached) || {};
                      if (parsed && typeof parsed === 'object') {
                        if (parsed.unit && typeof parsed.unit === 'object') { delete parsed.unit[gLower]; delete parsed.unit[gTrim]; }
                        if (parsed.items && typeof parsed.items === 'object') { delete parsed.items[gLower]; delete parsed.items[gTrim]; }
                        localStorage.setItem(cacheKey, JSON.stringify(parsed));
                      }
                    }
                  }
                }
              } catch {}
              setLines(next);
            }} title="Remove">
              <DeleteIcon fontSize="small" />
            </IconButton>
          </Stack>
        );
      },
    },
  ], [lines, palletDescByGroup, groupPriceByName, numberFmt2]);

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
        containerArrival: String(d?.containerArrival || d?.containerArrivalDate || d?.containerArrivalYmd || ''),
        estFulfillment: String(d?.estFulfillment || ''),
        estDelivered: String(d?.estDelivered || ''),
        updatedAt: String(d?.updatedAt || ''),
        updatedBy: String(d?.updatedBy || ''),
        customerEmail: String(d?.customerEmail || ''),
        customerName: String(d?.customerName || ''),
        customerPhone: String(d?.customerPhone || ''),
        shippingAddress: String(d?.shippingAddress || ''),
        companyName: String(d?.companyName || ''),
        customerNumber: String(d?.customerNumber || ''),
        salesRepresentative: String(d?.salesRepresentative || ''),
        originalPrice: String(d?.originalPrice || ''),
        shippingPercent: String(d?.shippingPercent || ''),
        discountPercent: String(d?.discountPercent || ''),
        paymentTerms: String(d?.paymentTerms || ''),
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
            ))),
            discount: (() => {
              const v = Number(l?.discount ?? l?.discountPercent ?? 0);
              if (!Number.isFinite(v)) return 0;
              return Math.max(0, Math.min(100, Math.floor(v)));
            })(),
            unitPriceAtAdd: (Number.isFinite(Number(l?.unitPriceAtAdd)) ? Number(l?.unitPriceAtAdd) : undefined),
            itemsAtAdd: (Array.isArray(l?.itemsAtAdd) ? l.itemsAtAdd : undefined),
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

  // Auto-promote: if status is PROCESSING and containerArrival is today or earlier, set to READY TO SHIP
  const promotingRef = useRef<Set<string>>(new Set());
  useEffect(() => {
    const today = new Date().toISOString().slice(0,10);
    const run = async () => {
      let didUpdate = false;
      for (const o of orders) {
        const st = String(o?.status || '').toLowerCase();
        const arr = String((o as any)?.containerArrival || '').slice(0,10);
        if (st === 'processing' && arr && arr <= today && !promotingRef.current.has(o.id)) {
          promotingRef.current.add(o.id);
          try {
            const payload = {
              status: 'ready_to_ship',
              createdAt: o.createdAt,
              containerArrival: arr,
              estFulfillment: o.estFulfillment,
              estDelivered: o.estDelivered,
              customerEmail: o.customerEmail,
              customerName: o.customerName,
              customerPhone: o.customerPhone,
              shippingAddress: o.shippingAddress,
              originalPrice: o.originalPrice,
              shippingPercent: o.shippingPercent,
              discountPercent: o.discountPercent,
              notes: o.notes,
              lines: Array.isArray(o.lines) ? o.lines.map(l => ({
                groupName: String((l as any)?.groupName || ''),
                lineItem: String((l as any)?.lineItem || ''),
                palletName: String((l as any)?.palletName || ''),
                qty: Math.max(0, Math.floor(Number((l as any)?.qty || 0))),
                discountPercent: (() => {
                  const v = Number((l as any)?.discount ?? (l as any)?.discountPercent ?? 0);
                  return Number.isFinite(v) ? Math.max(0, Math.min(100, Math.floor(v))) : 0;
                })(),
              })) : [],
            } as any;
            await api.put(`/early-buy/${encodeURIComponent(o.id)}`, payload);
            didUpdate = true;
          } catch {}
        }
      }
      if (didUpdate) {
        try { await refreshOrders(); } catch {}
      }
    };
    if (Array.isArray(orders) && orders.length) run();
  }, [orders, refreshOrders]);

  // Load pallet prices by group (exactly mirror Orders page logic):
  // - Build active item-group set
  // - Sum packSize * price for items per itemGroup (case-insensitive)
  useEffect(() => {
    let stopped = false;
    const loadPrices = async () => {
      try {
        const [groupsRes, itemsRes] = await Promise.all([
          api.get<any[]>('/item-groups'),
          api.get<any[]>('/items', { params: { includeDisabled: 1 } }),
        ]);

        const groups = Array.isArray(groupsRes?.data) ? groupsRes.data : [];
        const items = Array.isArray(itemsRes?.data) ? itemsRes.data : [];

        const activeSet = new Set(
          groups
            .filter((g: any) => (g as any).active !== false)
            .map((g: any) => String((g as any)?.name || '').trim())
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
    return () => { stopped = true; };
  }, []);

  const saveNewOrder = async () => {
    // Sync latest values from uncontrolled inputs (in case Save is clicked while a cell is focused)
    const syncedLines = (() => {
      const base = Array.isArray(lines) ? [...lines] : [];
      for (let i = 0; i < base.length; i++) {
        const qEl = qtyInputRefs.current[i];
        if (qEl && typeof qEl.value === 'string') {
          const v = Math.max(0, Math.floor(Number(qEl.value) || 0));
          base[i] = { ...base[i], qty: v } as any;
        }
        const dEl = discountInputRefs.current[i];
        if (dEl && typeof dEl.value === 'string') {
          const v = Math.max(0, Math.min(100, Math.floor(Number(dEl.value) || 0)));
          base[i] = { ...base[i], discount: v } as any;
        }
      }
      return base;
    })();

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
    // Container Arrival required (processing only)
    if (status === 'processing' && !String(containerArrival || '').trim()) errs.push('Container Arrival is required when status is PROCESSING');
    if (containerArrival && !isYmd(containerArrival)) errs.push('Container Arrival is invalid');
    if (containerArrival && createdAt && containerArrival < createdAt) errs.push('Container Arrival must be >= Created Order Date');
    if (containerArrival && estFulfillment && containerArrival > estFulfillment) errs.push('Container Arrival must be <= Estimated ShipDate for Customer');
    // If status is SHIPPED, estimated arrival date required
    if (status === 'shipped' && !estDelivered) errs.push('Estimated Arrival Date is required when status is SHIPPED');
    if (estDelivered && !isYmd(estDelivered)) errs.push('Estimated Arrival Date is invalid');
    if (estDelivered && estFulfillment && estDelivered < estFulfillment) errs.push('Estimated Arrival Date must be >= Estimated ShipDate');

    // If status is SHIPPED, shipping charges are required
    if (status === 'shipped') {
      const sp = Number(shippingPercent);
      if (!String(shippingPercent || '').trim() || !Number.isFinite(sp) || sp < 0 || sp > 100) errs.push('Shipping Charges (%) is required when status is SHIPPED');
    }

    // If status is COMPLETED, shipping charges are required
    if (status === 'completed') {
      const sp = Number(shippingPercent);
      if (!String(shippingPercent || '').trim() || !Number.isFinite(sp) || sp < 0 || sp > 100) errs.push('Shipping Charges (%) is required when status is COMPLETED');
    }

    const rows = Array.isArray(syncedLines) ? syncedLines : [];
    if (rows.length === 0) errs.push('Please add at least one pallet');
    const anyQty = rows.some(l => Number(l.qty) > 0);
    if (!anyQty) errs.push('Please add at least one pallet with quantity > 0');
    const invalidQtyIdx: number[] = [];
    const invalidDiscIdx: number[] = [];
    rows.forEach((l, idx) => {
      const q = Number(l.qty);
      const d = Number((l as any)?.discount ?? 0);
      if (!Number.isFinite(q) || q <= 0) invalidQtyIdx.push(idx);
      if (!Number.isFinite(d) || d < 0 || d > 100) invalidDiscIdx.push(idx);
    });
    if (invalidQtyIdx.length) errs.push('Each pallet quantity must be > 0');
    if (invalidDiscIdx.length) errs.push('Discount must be between 0 and 100');
    if (!String(paymentTerms || '').trim()) errs.push('Payment Terms is required');
    if (requestedShip && createdAt && requestedShip <= createdAt) errs.push('Requested Ship Date must be > Created Order Date');

    if (errs.length) {
      // mark invalid rows
      setInvalidQtyRows(new Set(invalidQtyIdx));
      setInvalidDiscountRows(new Set(invalidDiscIdx));
      toast.error(errs[0]);
      return;
    }

    // clear error markers and persist synced values before saving
    setInvalidQtyRows(new Set());
    setInvalidDiscountRows(new Set());
    if (rows !== lines) setLines(rows);

    // Read latest Sales Rep from input ref (uncontrolled)
    const salesRepValue = (salesRepInputRef.current && typeof salesRepInputRef.current.value === 'string')
      ? salesRepInputRef.current.value
      : salesRepresentative;

    try {
      const normalizedLines = rows.map(l => ({
        ...l,
        qty: Math.max(0, Math.floor(Number(l.qty) || 0)),
        discount: (() => {
          const v = Number((l as any)?.discount ?? 0);
          if (!Number.isFinite(v)) return 0;
          return Math.max(0, Math.min(100, Math.floor(v)));
        })(),
        discountPercent: (() => {
          const v = Number((l as any)?.discount ?? 0);
          if (!Number.isFinite(v)) return 0;
          return Math.max(0, Math.min(100, Math.floor(v)));
        })(),
      }));

      const payload = {
        status,
        createdAt,
        containerArrival,
        estFulfillment,
        estDelivered,
        requestedShipDate: requestedShip || undefined,
        customerEmail,
        customerName,
        customerPhone,
        shippingAddress,
        salesRepresentative: salesRepValue,
        customerNumber,
        companyName,
        originalPrice: computedSubTotal,
        shippingPercent,
        paymentTerms,
        discountPercent: '',
        notes,
        // multiple aliases for backend compatibility
        // include snapshots so server may echo them back on subsequent fetches
        lines: normalizedLines,
        items: normalizedLines,
        orderLines: normalizedLines,
        order: { lines: normalizedLines },
        payload: { lines: normalizedLines },
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
      const doc = data as any;
      setLastUpdatedAt(String(doc?.updatedAt || ''));
      setLastUpdatedBy(String(doc?.updatedBy || ''));
      // Persist snapshots locally keyed by saved order id
      try {
        const orderId = String((doc?.id || (editingOrder && editingOrder.id) || '')).trim();
        if (orderId) {
          const cacheKey = `earlyBuyLineSnapshots:${orderId}`;
          const unit: Record<string, number> = {};
          const items: Record<string, any[]> = {};
          for (const l of (Array.isArray(rows) ? rows : [])) {
            const k = String((l as any)?.groupName || '').trim().toLowerCase();
            if (!k) continue;
            const u = Number((l as any)?.unitPriceAtAdd);
            if (Number.isFinite(u)) unit[k] = u;
            const it = (l as any)?.itemsAtAdd;
            if (Array.isArray(it) && it.length) items[k] = it as any[];
          }
          localStorage.setItem(cacheKey, JSON.stringify({ unit, items }));
        }
      } catch {}
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
    setSalesRepresentative(''); setCustomerNumber(''); setCompanyName('');
    setCreatedAt(new Date().toISOString().slice(0,10)); setContainerArrival(''); setEstFulfillment(''); setEstDelivered('');
    setOriginalPrice(''); setShippingPercent(''); setDiscountPercent(''); setNotes('');
    setLastUpdatedAt('');
    setLastUpdatedBy('');
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
      const gLower = String(l.groupName || '').trim().toLowerCase();
      const desc = String(palletDescByGroup?.[gLower] || l.groupName || '');
      rows.push([l.lineItem, l.palletName, desc, Number(l.qty||0)]);
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Early Buy Order');
    XLSX.writeFile(wb, `${row.id}.xlsx`);
  };

  const resetEarlyBuyForm = useCallback(() => {
    setEditingOrder(null);
    setStatus('processing');
    setCustomerEmail('');
    setCustomerName('');
    setCustomerPhone('');
    setShippingAddress('');
    setSalesRepresentative('');
    try { if (salesRepInputRef.current) salesRepInputRef.current.value = ''; } catch {}
    setCustomerNumber('');
    setCompanyName('');
    setCreatedAt(new Date().toISOString().slice(0,10));
    setEstFulfillment('');
    setEstDelivered('');
    setRequestedShip('');
    setOriginalPrice('');
    setShippingPercent('');
    setDiscountPercent('');
    setNotes('');
    setLastUpdatedAt('');
    setLastUpdatedBy('');
    setLines([]);
    setPaymentTerms('Net 60');
    setPaymentStatus('Not Paid');
  }, []);

  return (
    <Container maxWidth="lg" sx={{ mt: 2, mb: 4 }}>
      <Typography variant="h5" sx={{ mb: 2, fontWeight: 700 }}>EARLY BUY</Typography>
      <Paper sx={{ p: 2, mb: 2 }}>
        <Stack direction="row" spacing={1} alignItems="center">
          <Button variant="contained" onClick={() => { resetEarlyBuyForm(); setOpen(true); }}>Add Order</Button>
          <Button variant="outlined" onClick={() => setEbReportOpen(true)}>Report</Button>
          <Button variant="text" onClick={() => refreshOrders()}>Refresh List</Button>
        </Stack>
      </Paper>

      <Paper sx={{ p: 2 }}>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} sx={{ mb: 1 }}>
          <Box sx={{ flex: 1 }} />
          <TextField
            size="small"
            label="Search"
            defaultValue={tableQ}
            onChange={(e)=> {
              const val = String(e.target.value || '');
              if (tableQDebounceRef.current) clearTimeout(tableQDebounceRef.current);
              tableQDebounceRef.current = setTimeout(() => {
                setTableQDebounced(val);
              }, 500);
            }}
            onBlur={(e)=> {
              const val = String(e.target.value || '');
              if (tableQDebounceRef.current) { clearTimeout(tableQDebounceRef.current); tableQDebounceRef.current = null; }
              setTableQ(val);
              setTableQDebounced(val);
            }}
            sx={{ minWidth: 240 }}
          />
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
            onRowDoubleClick={(p: any) => {
              const row = p?.row as any;
              if (!row) return;
              // populate form for view-only
              setEditingOrder(row);
              setStatus(row.status);
              setCustomerEmail(row.customerEmail);
              setCustomerName(row.customerName);
              setCustomerPhone(row.customerPhone);
              setShippingAddress(row.shippingAddress);
              try { setSalesRepresentative(String((row as any)?.salesRepresentative || '')); } catch {}
              try { setCustomerNumber(String((row as any)?.customerNumber || '')); } catch {}
              try { setCompanyName(String((row as any)?.companyName || '')); } catch {}
              setCreatedAt(row.createdAt);
              setContainerArrival(String((row as any).containerArrival || ''));
              setEstFulfillment(row.estFulfillment);
              setEstDelivered(row.estDelivered);
              setRequestedShip(String((row as any)?.requestedShipDate || ''));
              setOriginalPrice(row.originalPrice || '');
              setShippingPercent(row.shippingPercent || '');
              setDiscountPercent(row.discountPercent || '');
              setPaymentTerms(String((row as any)?.paymentTerms || 'Net 60'));
              setPaymentStatus(String((row as any)?.paymentStatus || 'Not Paid'));
              setNotes(row.notes || '');
              setLastUpdatedAt(String(row.updatedAt || ''));
              setLastUpdatedBy(String(row.updatedBy || ''));
              // Build lines with persisted snapshots preferred
              try {
                let base = Array.isArray(row.lines) ? row.lines.map((l: any) => ({ ...l })) : [];
                const orderId = String((row?.id || '')).trim();
                if (orderId) {
                  const cacheKey = `earlyBuyLineSnapshots:${orderId}`;
                  const cached = localStorage.getItem(cacheKey);
                  if (cached) {
                    const parsed = JSON.parse(cached);
                    if (parsed && typeof parsed === 'object') {
                      const unit = (parsed as any).unit || {};
                      const items = (parsed as any).items || {};
                      base = base.map((l: any) => {
                        const k = String(l?.groupName || '').trim().toLowerCase();
                        const u = Number((unit || {})[k]);
                        const it = (items || {})[k];
                        return {
                          ...l,
                          ...(Number.isFinite(u) ? { unitPriceAtAdd: u } : {}),
                          ...(Array.isArray(it) && it.length ? { itemsAtAdd: it } : {}),
                        } as any;
                      });
                    }
                  }
                }
                setLines(base);
              } catch { setLines(Array.isArray(row.lines) ? row.lines : []); }
              setOpen(true);
            }}
          />
        </div>
      </Paper>

      <Dialog
        open={open}
        onClose={() => { setOpen(false); resetEarlyBuyForm(); }}
        fullWidth
        maxWidth="xl"
        PaperProps={{ sx: { width: '90%', maxWidth: '90%' } }}
      >
        <DialogTitle>{editingOrder ? `Edit Order - ${editingOrder.id}` : 'New Early Buy Order'}</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <Box sx={{
              display: 'grid',
              gridTemplateColumns: { xs: '1fr', md: '1fr 1fr 1fr' },
              gap: 2,
              alignItems: 'center',
            }}>
              <TextField size="small" label="Warehouse" value="MPG" disabled fullWidth />
              {/* Status is always editable */}
              <TextField size="small" select label="Status" value={status} onChange={(e)=> handleStatusChange(e.target.value)} fullWidth disabled={!isStatusEditable}>
                {STATUS_OPTIONS.map(op => (<MenuItem key={op.value} value={op.value}>{op.label}</MenuItem>))}
              </TextField>
              <TextField
                label="Sales Representative"
                size="small"
                defaultValue={salesRepresentative}
                inputRef={(el)=> { salesRepInputRef.current = el; }}
                fullWidth
              />
            </Box>
            <Typography variant="subtitle2">Customer</Typography>
            <Box sx={{
              display: 'grid',
              gridTemplateColumns: { xs: '1fr', md: '1fr 1fr 1fr' },
              gap: 2,
              alignItems: 'center',
            }}>
              <Button
                variant="contained"
                size="small"
                onClick={()=>{ setCustomerPickOpen(true); loadCustomers(); }}
                sx={{ height: 40, width: { xs: '100%', sm: 'auto' } }}
              >
                Browse Customer
              </Button>
              <TextField
                size="small"
                required
                label="Customer Email"
                value={customerEmail}
                onChange={(e)=> setCustomerEmail(e.target.value)}
                fullWidth
                disabled
                error={Boolean(customerEmail) && !isValidEmail(customerEmail)}
                helperText={Boolean(customerEmail) && !isValidEmail(customerEmail) ? 'Enter a valid email address' : ''}
              />
              <TextField size="small" required label="Customer Name" value={customerName} onChange={(e)=> setCustomerName(e.target.value)} fullWidth disabled />
            </Box>
            <Box sx={{
              display: 'grid',
              gridTemplateColumns: { xs: '1fr', md: '1fr 1fr 1fr' },
              gap: 2,
              alignItems: 'center',
            }}>
              <TextField
                size="small"
                required
                label="Phone Number"
                value={customerPhone}
                onChange={(e)=> setCustomerPhone(e.target.value)}
                disabled
                fullWidth
              />
              <TextField disabled label="Customer Number" size="small" value={customerNumber} fullWidth />
              <TextField disabled label="Company Name" size="small" value={companyName} fullWidth />
            </Box>
            <TextField size="small" required label="Shipping Address" value={shippingAddress} onChange={(e)=> setShippingAddress(e.target.value)} fullWidth multiline minRows={2} disabled />
            <Typography variant="subtitle2">Dates</Typography>
            {(() => {
              const isYmd = (s: string) => /^\d{4}-\d{2}-\d{2}$/.test(String(s||''));
              const cA = createdAt && isYmd(createdAt) ? createdAt : '';
              const cR = containerArrival && isYmd(containerArrival) ? containerArrival : '';
              const eS = estFulfillment && isYmd(estFulfillment) ? estFulfillment : '';
              const eD = estDelivered && isYmd(estDelivered) ? estDelivered : '';
              const plusOne = (ymd?: string) => {
                const s = String(ymd || '').trim();
                if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return '';
                const d = new Date(`${s}T00:00:00`);
                if (Number.isNaN(d.getTime())) return '';
                const n = new Date(d); n.setDate(n.getDate() + 1);
                const mm = `${n.getMonth()+1}`.padStart(2,'0');
                const dd = `${n.getDate()}`.padStart(2,'0');
                return `${n.getFullYear()}-${mm}-${dd}`;
              };
              const rS = requestedShip && isYmd(requestedShip) ? requestedShip : '';
              const errReqShipLtCreated = Boolean(rS && cA && rS <= cA);
              const errCreatedFuture = Boolean(cA && cA > todayYmd);
              const errShipLtCreated = Boolean(eS && cA && eS < cA);
              const errContainerGtShip = Boolean(cR && eS && cR > eS);
              const errContainerLtCreated = Boolean(cR && cA && cR < cA);
              const errArrLtShip = Boolean(eD && eS && eD < eS);
              return (
                <Stack direction={{ xs:'column', sm:'row' }} spacing={2}>
                  <TextField
                    size="small" required type="date" label="Created Order Date" defaultValue={createdAt}
                    onBlur={(e)=> setCreatedAt(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isEditable}
                    inputProps={{ max: todayYmd }}
                    error={Boolean(cA) && errCreatedFuture}
                    helperText={Boolean(cA) && errCreatedFuture ? 'Created Order Date cannot be in the future' : ''}
                  />
                  <TextField
                    size="small" type="date" label="Requested Ship Date" defaultValue={requestedShip}
                    onBlur={(e)=> setRequestedShip(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isEditable}
                    inputProps={{ min: plusOne(cA) || undefined }}
                    error={Boolean(rS && cA) && errReqShipLtCreated}
                    helperText={Boolean(rS && cA) && errReqShipLtCreated ? 'Requested Ship Date must be > Created Order Date' : ''}
                  />
                  <TextField
                    size="small" required type="date" label="Container Arrival" defaultValue={containerArrival}
                    onBlur={(e)=> setContainerArrival(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isContainerArrivalEditable}
                    inputProps={{ min: cA || undefined, max: eS || undefined }}
                    error={(Boolean(cR && eS) && errContainerGtShip) || (Boolean(cR && cA) && errContainerLtCreated)}
                    helperText={
                      (Boolean(cR && eS) && errContainerGtShip)
                        ? 'Container Arrival must be ≤ Estimated Shipdate for Customer'
                        : ((Boolean(cR && cA) && errContainerLtCreated) ? 'Container Arrival must be ≥ Created Order Date' : '')
                    }
                  />
                  <TextField
                    size="small" required type="date" label="Estimated ShipDate for Customer" defaultValue={estFulfillment}
                    onBlur={(e)=> setEstFulfillment(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isEditable}
                    inputProps={{ min: cA || undefined }}
                    error={Boolean(eS && cA) && errShipLtCreated}
                    helperText={Boolean(eS && cA) && errShipLtCreated ? 'Estimated Shipdate must be ≥ Created Order Date' : ''}
                  />
                  <TextField
                    size="small" type="date" label="Estimated Arrival Date" defaultValue={estDelivered}
                    onBlur={(e)=> setEstDelivered(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isArrivalDateEditable}
                    inputProps={{ min: eS || undefined }}
                    error={Boolean(eD && eS) && errArrLtShip}
                    helperText={Boolean(eD && eS) && errArrLtShip ? 'Estimated Arrival Date must be ≥ Estimated Shipdate' : ''}
                  />
                </Stack>
              );
            })()}
            <Typography variant="subtitle2">Pricing</Typography>
            <Stack direction={{ xs:'column', md:'row' }} spacing={2}>
              <TextField size="small" label="Sub Total" value={computedSubTotal} fullWidth disabled />
              <TextField
                size="small"
                label="Shipping Charges (%)"
                defaultValue={shippingPercent}
                onBlur={(e)=> {
                  const raw = String(e.target.value || '').trim();
                  const n = Number(raw);
                  if (!Number.isFinite(n)) { setShippingPercent(''); (e.target as HTMLInputElement).value = ''; return; }
                  const clamped = Math.max(0, Math.min(100, n));
                  const out = String(clamped);
                  setShippingPercent(out);
                  try { (e.target as HTMLInputElement).value = out; } catch {}
                }}
                inputProps={{ inputMode: 'decimal', step: '0.01', min: 0, max: 100, pattern: '^[0-9]*\.?[0-9]*$' }}
                fullWidth
                disabled={!isEditable}
              />
              <TextField size="small" label="Grand Total" value={computedGrandTotal} fullWidth disabled />
            </Stack>
            <Typography variant="subtitle2">Payment</Typography>
            <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap' }}>
              <TextField
                size="small"
                label="Payment Terms"
                select
                value={paymentTerms}
                onChange={(e)=> setPaymentTerms(e.target.value)}
                required
                disabled={!isPaymentEditable}
                error={!String(paymentTerms||'').trim()}
                helperText={!String(paymentTerms||'').trim() ? 'Required' : ''}
                sx={{ width: { xs: '100%', sm: 300 } }}
              >
                <MenuItem value={"Pre-Pay"}>Pre-Pay</MenuItem>
                <MenuItem value={"Net 30"}>Net 30</MenuItem>
                <MenuItem value={"Net 60"}>Net 60</MenuItem>
              </TextField>
              <TextField
                size="small"
                label="Payment Status"
                select
                value={paymentStatus}
                onChange={(e)=> setPaymentStatus(e.target.value)}
                disabled={!isPaymentEditable}
                sx={{ width: { xs: '100%', sm: 260 } }}
              >
                <MenuItem value={"Paid"}>Paid</MenuItem>
                <MenuItem value={"Not Paid"}>Not Paid</MenuItem>
                <MenuItem value={"10% Paid"}>10% Paid</MenuItem>
              </TextField>
            </Box>
            <Typography variant="subtitle2">Pallets to Order</Typography>
            <TextField size="small" label="Remarks/Notes" defaultValue={notes} onBlur={(e)=> setNotes(e.target.value)} fullWidth multiline minRows={3} disabled={!isEditable} />
            <Typography variant="caption" color="error" sx={{ mt: -1 }}>
              {(() => {
                const at = String(lastUpdatedAt || '').trim();
                const by = String(lastUpdatedBy || '').trim();
                if (!at && !by) return '';
                let labelAt = at;
                try {
                  labelAt = at ? new Date(at).toLocaleString() : '';
                } catch {}
                const who = by ? ` (by ${by})` : '';
                return `Last Updated: ${labelAt || '-'}${who}`;
              })()}
            </Typography>

            <Stack direction="row" spacing={1} alignItems="center">
              <Button variant="outlined" onClick={()=> setPickerOpen(true)} disabled={!isEditable}>Add to List</Button>
              <Box sx={{ flex: 1 }} />
            </Stack>

            <div style={{ height: 300, width: '100%' }}>
              <DataGrid
                rows={lines.map((l, idx) => ({ id: idx+1, ...l }))}
                columns={linesColumns.map(c =>
                  c.field === 'actions'
                    ? { ...c, renderCell: (p:any) => (!isEditable ? null : (c as any).renderCell(p)) }
                    : c.field === 'qty'
                      ? {
                          ...c,
                          renderCell: (p:any) => {
                            if (!isEditable) return <span>{Number(lines[Math.max(0, Number(p?.id)-1)]?.qty || 0)}</span>;
                            return (c as any).renderCell(p);
                          },
                        }
                      : c.field === 'discount'
                        ? {
                            ...c,
                            renderCell: (p:any) => {
                              if (!isEditable) {
                                const i = Math.max(0, Number(p?.id) - 1);
                                const v = Math.max(0, Math.min(100, Number(lines[i]?.discount ?? 0)));
                                return <span>{v}</span>;
                              }
                              return (c as any).renderCell(p);
                            },
                          }
                        : c
                )}
                onCellDoubleClick={(p: any) => {
                  const field = String(p?.field || '').trim();
                  if (field === 'discount' || field === 'qty' || field === 'actions') return;
                  const g = String(p?.row?.groupName || '').trim();
                  if (g) openPalletItems({ groupName: g });
                }}
                getRowClassName={(params: any) => {
                  const row = (params && typeof params === 'object') ? (params as any).row : {};
                  const id = Number((params && (params as any).id) || 0);
                  const idx = Math.max(0, Math.floor(id) - 1);
                  const g = String((row as any)?.groupName || '').trim().toLowerCase();
                  if (!g) return '';
                  let firstIndex = -1;
                  for (let i = 0; i < lines.length; i++) {
                    const gi = String(lines[i]?.groupName || '').trim().toLowerCase();
                    if (gi && gi === g) { firstIndex = i; break; }
                  }
                  if (firstIndex === -1) return '';
                  return idx > firstIndex ? 'duplicate-pallet-row' : 'original-pallet-row';
                }}
                disableRowSelectionOnClick
                density="compact"
                sx={{
                  '& .duplicate-pallet-row': {
                    backgroundColor: '#FAF6E9',
                  },
                  '& .original-pallet-row .MuiDataGrid-cell': {
                    fontWeight: 700,
                  },
                }}
              />
            </div>
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> { setOpen(false); resetEarlyBuyForm(); }}>Cancel</Button>
          <Button variant="outlined" onClick={exportEarlyBuyXlsx} disabled={!lines.length}>Export .xlsx</Button>
          <Button variant="outlined" onClick={exportEarlyBuyPdf} disabled={!lines.length}>Export .pdf</Button>
          <Button variant="contained" onClick={saveNewOrder}>Save</Button>
        </DialogActions>
      </Dialog>

      {/* Customer Picker Dialog */}
      <Dialog open={customerPickOpen} onClose={()=> setCustomerPickOpen(false)} fullWidth maxWidth="md">
        <DialogTitle>Browse Customer</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <TextField
              size="small"
              label="Search"
              placeholder="Email, name, phone, number, company, address"
              defaultValue={customerPickQ}
              onChange={(e)=>{
                if (customerPickDebounceRef.current) try { clearTimeout(customerPickDebounceRef.current); } catch {}
                const val = e.target.value;
                customerPickDebounceRef.current = setTimeout(()=> setCustomerPickQ(val), 300);
              }}
            />
            <Paper sx={{ height: 420, width: '100%', p: 1 }}>
              <DataGrid
                rows={(Array.isArray(customerRows) ? customerRows : [])
                  .filter((r:any)=>{
                    const s = customerPickQ.trim().toLowerCase();
                    if (!s) return true;
                    const hay = [r.email||'', r.name||'', r.phone||'', r.accountNumber||'', r.companyName||'', r.shippingAddress||''].join(' ').toLowerCase();
                    return hay.includes(s);
                  })
                  .map((r:any)=> ({ ...r, id: r._id }))}
                columns={[
                  {
                    field: 'actions',
                    headerName: 'Actions',
                    width: 120,
                    sortable: false,
                    filterable: false,
                    renderCell: (params: any) => (
                      <Button size="small" variant="outlined" onClick={()=>{
                        const r = params.row;
                        setCustomerEmail(String(r.email||''));
                        setCustomerName(String(r.name||''));
                        setCustomerPhone(String(r.phone||''));
                        setCustomerNumber(String(r.accountNumber||''));
                        if (r && String(r.salesRepresentative || '').trim()) {
                          setSalesRepresentative(String(r.salesRepresentative));
                        }
                        setCompanyName(String(r.companyName||''));
                        setShippingAddress(String(r.shippingAddress||''));
                        setCustomerPickOpen(false);
                      }}>Use</Button>
                    ),
                  },
                  { field: 'email', headerName: 'Customer Email', flex: 1.4, minWidth: 200 },
                  { field: 'name', headerName: 'Customer Name', flex: 1.2, minWidth: 160 },
                  { field: 'phone', headerName: 'Phone Number', width: 150 },
                  { field: 'accountNumber', headerName: 'Customer Number', width: 160 },
                  { field: 'companyName', headerName: 'Company Name', width: 180 },
                  { field: 'shippingAddress', headerName: 'Shipping Address', flex: 1.6, minWidth: 240 },
                ]}
                loading={customerLoading}
                disableRowSelectionOnClick
                initialState={{ pagination: { paginationModel: { pageSize: 5, page: 0 } } }}
                pageSizeOptions={[5, 10, 20, 50]}
                density="compact"
              />
            </Paper>
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setCustomerPickOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={ebReportOpen} onClose={() => setEbReportOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Early Buy - Pallet Orders Report</DialogTitle>
        <DialogContent>
          <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} sx={{ mt: 1 }}>
            <TextField
              type="date"
              size="small"
              label="From"
              value={ebReportFrom}
              onChange={(e) => setEbReportFrom(e.target.value)}
              InputLabelProps={{ shrink: true }}
              fullWidth
            />
            <TextField
              type="date"
              size="small"
              label="To"
              value={ebReportTo}
              onChange={(e) => setEbReportTo(e.target.value)}
              InputLabelProps={{ shrink: true }}
              fullWidth
            />
          </Stack>
          {ebReportExporting ? <LinearProgress sx={{ mt: 2 }} /> : null}
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setEbReportOpen(false)} disabled={ebReportExporting}>Close</Button>
          <Button variant="outlined" onClick={exportEarlyBuyReportXlsx} disabled={ebReportExporting}>Export .xlsx</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={pickerOpen} onClose={()=> setPickerOpen(false)} fullWidth maxWidth="xl">
        <DialogTitle>Select Pallets</DialogTitle>
        <DialogContent>
          <Stack direction={{ xs:'column', sm:'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }} sx={{ mb: 1, mt: 1 }}>
            <TextField
              size="small"
              label="Search Pallet ID / Pallet Description / Pallet Name"
              defaultValue={pickerQ}
              onChange={(e)=> {
                const val = String(e.target.value || '');
                if (pickerQDebounceRef.current) clearTimeout(pickerQDebounceRef.current);
                pickerQDebounceRef.current = setTimeout(() => {
                  setPickerQDebounced(val);
                }, 500);
              }}
              onBlur={(e)=> {
                const val = String(e.target.value || '');
                if (pickerQDebounceRef.current) { clearTimeout(pickerQDebounceRef.current); pickerQDebounceRef.current = null; }
                setPickerQ(val);
                setPickerQDebounced(val);
              }}
              sx={{ flex: 1, minWidth: 260 }}
            />
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
                { field: 'groupName', headerName: 'Pallet Description', flex: 1.2, minWidth: 260, sortable: true, renderCell: (p: any) => {
                  const row: any = p?.row || {};
                  const gLower = String(row?.groupName || '').trim().toLowerCase();
                  return String(palletDescByGroup?.[gLower] || row?.groupName || '');
                } },
                { field: 'lineItem', headerName: 'Pallet ID', width: 160, sortable: true },
                { field: 'palletPrice', headerName: 'Pallet Price', width: 110, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => {
                  const gLower = String(p?.row?.groupName || '').trim().toLowerCase();
                  const val = Number(groupPriceByName?.[gLower]);
                  if (!Number.isFinite(val)) return '';
                  return numberFmt2.format(val);
                } },
              ]) as GridColDef[]}
              checkboxSelection
              isRowSelectable={(p: any) => {
                const g = String((p && (p.id ? (pickerRowsFiltered.find((r:any)=> String(r?.lineItem||r?.groupName||r?.palletName||'').trim()===String(p.id)) || (p as any).row) : (p as any)?.row))?.groupName || '').trim().toLowerCase();
                if (!g) return true;
                const exists = (Array.isArray(lines) ? lines : []).some((l) => String(l?.groupName||'').trim().toLowerCase() === g);
                return !exists;
              }}
              onRowDoubleClick={(p: any) => {
                const g = String(p?.row?.groupName || '').trim();
                if (g) openPalletItems({ groupName: g });
              }}
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

      <Dialog open={palletItemsOpen} onClose={()=> setPalletItemsOpen(false)} fullWidth maxWidth="md">
        <DialogTitle>{`Pallet Items - ${palletItemsGroupName || ''}`}</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          {palletItemsLoading ? <LinearProgress sx={{ mb: 2 }} /> : null}
          <div style={{ height: 420, width: '100%' }}>
            <DataGrid
              rows={(Array.isArray(palletItemsRows) ? palletItemsRows : []).map((r: any, idx: number) => ({ id: String(r?.itemCode || idx), ...r }))}
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
          <Button onClick={()=> setPalletItemsOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
