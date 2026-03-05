import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, Dialog, DialogTitle, DialogContent, DialogActions, MenuItem, Box, IconButton, Chip, LinearProgress } from '@mui/material';
import { DataGrid, GridColDef, GridToolbar } from '@mui/x-data-grid';
import DeleteIcon from '@mui/icons-material/Delete';
import DownloadIcon from '@mui/icons-material/Download';
import OpenInNewIcon from '@mui/icons-material/OpenInNew';
import ContentCopyIcon from '@mui/icons-material/ContentCopy';
import XLSX from 'xlsx-js-style';
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

  const [palletDescByGroup, setPalletDescByGroup] = useState<Record<string, string>>({});
  const [groupPriceByName, setGroupPriceByName] = useState<Record<string, number>>({});

  // Form state
  const [status, setStatus] = useState<EarlyOrder['status']>('processing');
  const [customerEmail, setCustomerEmail] = useState('');
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [createdAt, setCreatedAt] = useState(() => new Date().toISOString().slice(0,10));
  const [containerArrival, setContainerArrival] = useState('');
  const [estFulfillment, setEstFulfillment] = useState('');
  const [estDelivered, setEstDelivered] = useState('');
  const [shippingAddress, setShippingAddress] = useState('');
  const [lastUpdatedAt, setLastUpdatedAt] = useState('');
  const [lastUpdatedBy, setLastUpdatedBy] = useState('');

  const isEditable = useMemo(() => {
    // Editable when creating (no editingOrder) OR when status is processing/ready_to_ship
    return !editingOrder || status === 'processing' || status === 'ready_to_ship';
  }, [editingOrder, status]);

  const isArrivalDateEditable = useMemo(() => {
    // Allow editing arrival date even when status is shipped
    return !editingOrder || status === 'processing' || status === 'ready_to_ship' || status === 'shipped';
  }, [editingOrder, status]);

  const isStatusEditable = useMemo(() => {
    // Once an existing order is COMPLETED, it becomes view-only
    return !editingOrder || (status !== 'completed' && status !== 'canceled');
  }, [editingOrder, status]);

  const isContainerArrivalEditable = useMemo(() => {
    return status === 'processing' && isEditable;
  }, [isEditable, status]);

  const [originalPrice, setOriginalPrice] = useState('');
  const [shippingPercent, setShippingPercent] = useState('');
  const [discountPercent, setDiscountPercent] = useState('');
  const [notes, setNotes] = useState('');

  const numberFmt2 = useMemo(() => new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }), []);

  const handleStatusChange = useCallback((nextRaw: any) => {
    const extracted = (nextRaw && typeof nextRaw === 'object' && 'target' in nextRaw)
      ? (nextRaw as any)?.target?.value
      : nextRaw;
    const next = String(extracted || '').trim().toLowerCase() as EarlyOrder['status'];
    if (!['processing', 'ready_to_ship', 'shipped', 'completed', 'canceled'].includes(String(next))) return;
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
        'Are you sure you want to change the status to COMPLETED?\n\nThis cannot be undone and you can no longer edit the order.'
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
  useEffect(() => {
    const h = setTimeout(() => setPickerQDebounced(String(pickerQ || '')), 250);
    return () => clearTimeout(h);
  }, [pickerQ]);
  const [pickerWarehouses, setPickerWarehouses] = useState<any[]>([]);
  const [selectedGroups, setSelectedGroups] = useState<string[]>([]);

  // Lines added
  const [lines, setLines] = useState<Array<{ groupName: string; lineItem: string; palletName: string; qty: number; discount?: number }>>([]);

  // Computed totals based on lines
  const computedSubTotal = useMemo(() => {
    let sum = 0;
    for (let i = 0; i < lines.length; i++) {
      const g = String(lines[i]?.groupName || '').trim().toLowerCase();
      const unit = Number(groupPriceByName?.[g]);
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
  const [tableStatus, setTableStatus] = useState<'all' | EarlyOrder['status']>('all');

  // MUI DataGrid compatibility shim (v5 vs v6 selection APIs)
  const DG: any = DataGrid as any;

  const [palletItemsOpen, setPalletItemsOpen] = useState(false);
  const [palletItemsLoading, setPalletItemsLoading] = useState(false);
  const [palletItemsGroupName, setPalletItemsGroupName] = useState('');
  const [palletItemsRows, setPalletItemsRows] = useState<any[]>([]);

  const safe = (v: any) => String(v ?? '').replace(/[\/\\:*?"<>|]/g, '-');

  const exportEarlyBuyXlsx = useCallback(async () => {
    const orderNo = safe(String(editingOrder?.id || 'EORD'));
    const wb = XLSX.utils.book_new();

    const rows: any[][] = [];
    rows.push([String(orderNo || ''), 'To:', String(customerName || ''), 'Phone:', String(customerPhone || '')]);
    rows.push(['', 'Address:', String(shippingAddress || ''), 'Email:', String(customerEmail || '')]);
    rows.push([]);
    rows.push(['Status:', String(status || '').toUpperCase()]);
    rows.push(['Estimated Shipdate for Customer:', String(estFulfillment || '')]);
    rows.push(['Estimated Arrival Date:', String(estDelivered || '')]);
    rows.push([]);

    // Group lines by pallet group so we don't duplicate header/items for duplicates
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
      try {
        const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
        itemsByGroup[g.toLowerCase()] = Array.isArray((data as any)?.items) ? (data as any).items : [];
      } catch { itemsByGroup[g.toLowerCase()] = []; }
    }));

    for (const [gKey, gMeta] of groupLinesMap.entries()) {
      const unit = Number(groupPriceByName?.[gKey]);

      rows.push([String(gMeta.palletName || gMeta.display), String(palletDescByGroup?.[gKey] || gMeta.display || ''), String(gMeta.lineItem || '')]);
      const items = itemsByGroup[gKey] || [];
      for (const it of items) {
        rows.push([
          '',
          String((it as any)?.itemCode || ''),
          String((it as any)?.description || ''),
          String((it as any)?.upc || ''),
          Number((it as any)?.packSize ?? 0) || '',
          Number.isFinite(Number((it as any)?.price ?? 0)) ? Number(Number((it as any)?.price ?? 0).toFixed(2)) : '',
        ]);
      }
      // One summary line per occurrence in the order for this group
      for (const ln of gMeta.lines) {
        const discounted = Number.isFinite(unit) ? unit * (1 - ln.discount/100) : NaN;
        const subtotal = Number.isFinite(discounted) ? discounted * ln.qty : NaN;
        rows.push([
          `PALLET PRICE: ${Number.isFinite(unit) ? `$${Number(unit.toFixed(2))}` : ''}`,
          `DISCOUNT: ${Number.isFinite(ln.discount) ? `${ln.discount}%` : ''}`,
          `DISCOUNTED PRICE: ${Number.isFinite(discounted) ? `$${Number(discounted.toFixed(2))}` : ''}`,
          `QUANTITY: ${ln.qty || ''}`,
          'SUB TOTAL',
          Number.isFinite(subtotal) ? Number(subtotal.toFixed(2)) : '',
        ]);
      }
      rows.push([]);
    }

    rows.push(['SUB TOTAL', Number.isFinite(Number(computedSubTotal)) ? Number(Number(computedSubTotal).toFixed(2)) : '']);
    rows.push(['SHIPPING', (shippingPercent !== undefined && shippingPercent !== null && String(shippingPercent) !== '') ? `${shippingPercent}%` : '']);
    rows.push(['GRAND TOTAL', Number.isFinite(Number(computedGrandTotal)) ? Number(Number(computedGrandTotal).toFixed(2)) : '']);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    // Column widths for readability
    (ws as any)['!cols'] = [
      { wch: 28 }, // A
      { wch: 22 }, // B
      { wch: 30 }, // C
      { wch: 16 }, // D
      { wch: 14 }, // E
      { wch: 16 }, // F
    ];

    // Styling helpers and constants
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

    // Header rows styling (rows 0 and 1)
    const blue = fill('FFDDEBF7');
    for (let c = 0; c <= 4; c++) style(0, c, { ...bold });
    style(0, 0, { ...blue, ...border, font: { bold: true, sz: 12 } }); // ORDER #
    style(0, 1, { ...blue, ...border, ...bold }); // To:
    style(0, 3, { ...blue, ...border, ...bold }); // Phone:
    style(1, 1, { ...blue, ...border, ...bold }); // Address:
    style(1, 2, { ...bold }); // Shipping Address value bold
    style(1, 3, { ...blue, ...border, ...bold }); // Email:
    style(1, 4, { ...bold }); // Email value bold

    // Status and dates labels bold (rows 3..5 col A)
    style(3, 0, { ...bold });
    style(4, 0, { ...bold });
    style(5, 0, { ...bold });

    // Walk rows to style group headers, item rows, and summary rows
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

      // Pallet meta row: A,B,C have values; D,E,F empty
      if (a && b && c && !d && !e && !f) {
        // close previous group box
        if (groupStart !== null && groupStart < r) {
          for (let rr = groupStart; rr < r; rr++) {
            for (let cc = 0; cc <= 5; cc++) style(rr, cc, { ...border });
          }
        }
        groupStart = r;
        inItems = true;
        zebra = false;
        for (let col = 0; col <= 5; col++) style(r, col, { ...fill('FFE2EFDA'), ...border }); // light green
        continue;
      }

      // Summary row starts with 'PALLET PRICE:'
      if (aStr && aStr.toUpperCase().startsWith('PALLET PRICE')) {
        for (let col = 0; col <= 5; col++) style(r, col, { ...fill('FFF2F2F2'), ...border }); // gray
        style(r, 0, { ...bold });
        style(r, 1, { ...bold });
        style(r, 2, { ...bold });
        style(r, 3, { ...bold });
        style(r, 4, { ...bold, ...blue });
        style(r, 5, { ...blue, alignment: { horizontal: 'right' }, numFmt: '"$"#,##0.00' });
        // close the current group box
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
          for (let col = 0; col <= 5; col++) style(r, col, { ...fill('FFF9F9F9') });
        }
      }
    }
    // Ensure last open group gets bordered
    if (groupStart !== null) {
      for (let rr = groupStart; rr < rowCount; rr++) {
        for (let cc = 0; cc <= 5; cc++) style(rr, cc, { ...border });
      }
    }

    // Totals block styling: find SUB TOTAL..GRAND TOTAL
    let totalsStart = -1;
    let totalsEnd = -1;
    for (let r = rowCount - 1; r >= 0 && r >= rowCount - 12; r--) {
      const a = get(r, 0)?.v;
      const aStr = typeof a === 'string' ? a.toUpperCase() : '';
      if (aStr === 'GRAND TOTAL' && totalsEnd === -1) totalsEnd = r;
      if (aStr === 'SUB TOTAL') totalsStart = r;
      if (aStr === 'SUB TOTAL' || aStr === 'SHIPPING' || aStr === 'GRAND TOTAL') {
        style(r, 0, { ...bold, ...blue, ...border });
        style(r, 1, { ...border });
      }
    }
    if (totalsStart !== -1 && totalsEnd !== -1 && totalsEnd >= totalsStart) {
      for (let r = totalsStart; r <= totalsEnd; r++) {
        for (let c = 0; c <= 5; c++) style(r, c, { ...border });
      }
      // Right-align and currency format for monetary rows
      for (let r = totalsStart; r <= totalsEnd; r++) {
        const a = get(r, 0)?.v;
        const aStr = typeof a === 'string' ? a.toUpperCase() : '';
        if (aStr === 'SUB TOTAL' || aStr === 'GRAND TOTAL') style(r, 1, { alignment: { horizontal: 'right' }, numFmt: '"$"#,##0.00' });
      }
      // Emphasize GRAND TOTAL value
      style(totalsEnd, 1, { font: { bold: true, sz: 12 } });
    }

    // Align numeric columns (E: pack size, F: price/subtotal)
    for (let r = 0; r < rowCount; r++) {
      ensure(r, 4); ensure(r, 5);
      style(r, 4, { alignment: { horizontal: 'right' } });
      style(r, 5, { alignment: { horizontal: 'right' }, numFmt: '"$"#,##0.00' });
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Order Details');
    const fname = `early-buy-${orderNo}.xlsx`;
    XLSX.writeFile(wb, fname);
  }, [editingOrder, customerName, customerPhone, shippingAddress, customerEmail, status, estFulfillment, estDelivered, lines, groupPriceByName, palletDescByGroup, computedSubTotal, shippingPercent, computedGrandTotal]);

  const exportEarlyBuyPdf = useCallback(async () => {
    const orderNo = safe(String(editingOrder?.id || 'EORD'));
    const number = (n: any) => {
      const v = Number(n);
      return Number.isFinite(v) ? numberFmt2.format(v) : '';
    };

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
      try {
        const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
        itemsByGroup[g.toLowerCase()] = Array.isArray((data as any)?.items) ? (data as any).items : [];
      } catch { itemsByGroup[g.toLowerCase()] = []; }
    }));

    const content: any[] = [];

    // Header block (two rows) with blue label fill and borders
    const blue = '#DDEBF7';
    const green = '#E2EFDA';
    const gray = '#F2F2F2';
    content.push({
      table: {
        widths: [160, 40, '*', 55, 120],
        body: [
          [
            { text: String(orderNo || ''), fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: 'To:', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: String(customerName || ''), border: [true, true, true, true] },
            { text: 'Phone:', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: String(customerPhone || ''), border: [true, true, true, true] },
          ],
          [
            { text: '', border: [true, false, true, true] },
            { text: 'Address:', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: String(shippingAddress || ''), bold: true, border: [true, true, true, true] },
            { text: 'Email:', fillColor: blue, bold: true, border: [true, true, true, true] },
            { text: String(customerEmail || ''), bold: true, border: [true, true, true, true] },
          ],
        ],
      },
      layout: 'noHorizontalLines',
      margin: [0,0,0,8],
    });

    // Status and dates (match Orders formatting)
    const sStatus = String(status || '').toUpperCase();
    content.push({ text: 'Status:', bold: true, margin: [0, 2, 0, 0] });
    content.push({ text: sStatus, margin: [0, 0, 0, 2] });
    content.push({ text: 'Estimated Shipdate for Customer:', bold: true });
    content.push({ text: String(estFulfillment || ''), margin: [0, 0, 0, 2] });
    content.push({ text: 'Estimated Arrival Date:', bold: true });
    content.push({ text: String(estDelivered || ''), margin: [0, 0, 0, 8] });

    // Per-pallet sections (header/items once, multiple summary rows)
    for (const [gKey, gMeta] of groupLinesMap.entries()) {
      const unit = Number(groupPriceByName?.[gKey]);

      // Group header band (A..F)
      content.push({
        table: {
          widths: ['*', '*', 120, 60, 60, 80],
          body: [[
            { text: String(gMeta.palletName || gMeta.display), fillColor: green, bold: true, border: [true, true, true, true] },
            { text: String(palletDescByGroup?.[gKey] || gMeta.display || ''), fillColor: green, border: [true, true, true, true] },
            { text: String(gMeta.lineItem || ''), fillColor: green, border: [true, true, true, true] },
            { text: '', fillColor: green, border: [true, true, true, true] },
            { text: '', fillColor: green, border: [true, true, true, true] },
            { text: '', fillColor: green, border: [true, true, true, true] },
          ]],
        },
        layout: 'noHorizontalLines',
      });

      // Item table (no header row), bordered
      const items = itemsByGroup[gKey] || [];
      if (items.length) {
        const itemRows = items.map((it: any) => ([
          { text: '', border: [true, true, true, true] },
          { text: String(it?.itemCode || ''), border: [true, true, true, true] },
          { text: String(it?.description || ''), border: [true, true, true, true] },
          { text: String(it?.upc || ''), alignment: 'right', border: [true, true, true, true] },
          { text: (Number.isFinite(Number(it?.packSize)) ? String(Number(it?.packSize)) : ''), alignment: 'right', border: [true, true, true, true] },
          { text: (Number.isFinite(Number(it?.price)) ? `$${Number(it?.price).toFixed(2)}` : ''), alignment: 'right', border: [true, true, true, true] },
        ]));
        content.push({ table: { widths: ['*', 90, '*', 70, 60, 70], body: itemRows }, layout: 'noHorizontalLines' });
      }

      // One summary row per occurrence
      for (const ln of gMeta.lines) {
        const discounted = Number.isFinite(unit) ? unit * (1 - ln.discount/100) : NaN;
        const subtotal = Number.isFinite(discounted) ? discounted * ln.qty : NaN;
        content.push({
          table: {
            widths: ['*', 90, 150, 90, 90, 80],
            body: [[
              { text: `PALLET PRICE: ${Number.isFinite(unit) ? `$${unit.toFixed(2)}` : ''}`, fillColor: gray, bold: true, border: [true, true, true, true] },
              { text: `DISCOUNT: ${Number.isFinite(ln.discount) ? `${ln.discount}%` : ''}`, fillColor: gray, bold: true, border: [true, true, true, true] },
              { text: `DISCOUNTED PRICE: ${Number.isFinite(discounted) ? `$${discounted.toFixed(2)}` : ''}`, fillColor: gray, bold: true, border: [true, true, true, true] },
              { text: `QUANTITY: ${ln.qty || ''}`, fillColor: gray, bold: true, border: [true, true, true, true] },
              { text: 'SUB TOTAL', fillColor: blue, bold: true, border: [true, true, true, true] },
              { text: Number.isFinite(subtotal) ? `$${subtotal.toFixed(2)}` : '', alignment: 'right', fillColor: blue, bold: true, border: [true, true, true, true] },
            ]],
          },
          layout: 'noHorizontalLines',
          margin: [0, 2, 0, 6],
        });
      }
    }

    // Footer totals (blue fill + borders and bold like Orders PDF)
    const subText = Number.isFinite(Number(computedSubTotal)) ? `$${Number(computedSubTotal).toFixed(2)}` : '';
    const grandText = Number.isFinite(Number(computedGrandTotal)) ? `$${Number(computedGrandTotal).toFixed(2)}` : '';
    const shipText = `${String(shippingPercent||'')}%`;
    content.push({
      table: {
        widths: [120, 120],
        body: [
          [ { text: 'SUB TOTAL', fillColor: blue, bold: true, border: [true,true,true,true] }, { text: subText, alignment: 'right', border: [true,true,true,true] } ],
          [ { text: 'SHIPPING', fillColor: blue, bold: true, border: [true,true,true,true] }, { text: shipText, alignment: 'right', border: [true,true,true,true] } ],
          [ { text: 'GRAND TOTAL', fillColor: blue, bold: true, border: [true,true,true,true] }, { text: grandText, alignment: 'right', bold: true, border: [true,true,true,true] } ],
        ],
      },
      layout: 'noHorizontalLines',
      margin: [0, 4, 0, 0],
    });

    const docDefinition = { pageOrientation: 'landscape', pageMargins: [20,20,20,20], content, defaultStyle: { fontSize: 9 } } as any;
    const fname = `early-buy-${orderNo}.pdf`;
    (pdfMake as any).createPdf(docDefinition).download(fname);
  }, [editingOrder, lines, customerName, customerPhone, shippingAddress, customerEmail, status, estFulfillment, estDelivered, groupPriceByName, palletDescByGroup, computedSubTotal, shippingPercent, computedGrandTotal, numberFmt2]);

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

  const openPalletItems = useCallback(async ({ groupName }: { groupName: string }) => {
    const g = String(groupName || '').trim();
    if (!g) return;
    setPalletItemsOpen(true);
    setPalletItemsGroupName(g);
    setPalletItemsRows([]);
    setPalletItemsLoading(true);
    try {
      const { data } = await api.get(`/pallet-inventory/groups/${encodeURIComponent(g)}`);
      setPalletItemsRows(Array.isArray((data as any)?.items) ? (data as any).items : []);
    } catch {
      setPalletItemsRows([]);
    } finally {
      setPalletItemsLoading(false);
    }
  }, []);

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
      const pdesc = String(palletDescByGroup?.[gnameLower] || '').trim().toLowerCase();
      return gid.includes(q) || gnameLower.includes(q) || pname.includes(q) || pdesc.includes(q);
    });
  }, [pickerRows, pickerQDebounced, palletDescByGroup]);

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
    { field: 'groupName', headerName: 'Pallet Description', width: 260, renderCell: (p: any) => {
      const row: any = p?.row || {};
      const gLower = String(row?.groupName || '').trim().toLowerCase();
      return String(palletDescByGroup?.[gLower] || row?.groupName || '');
    } },
    { field: 'lineItem', headerName: 'Pallet ID', width: 160 },
    { field: 'palletPrice', headerName: 'Pallet Price', width: 110, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => {
      const gLower = String(p?.row?.groupName || '').trim().toLowerCase();
      const val = Number(groupPriceByName?.[gLower]);
      if (!Number.isFinite(val)) return '';
      return numberFmt2.format(val);
    } },
    { field: 'discount', headerName: 'Discount (%)', width: 120, align: 'right', headerAlign: 'right', renderCell: (p: any) => {
      const idx = Math.max(0, Number(p?.id) - 1);
      const v = Math.max(0, Math.min(100, Number(lines[idx]?.discount ?? 0)));
      return (
        <TextField
          size="small"
          type="number"
          value={v}
          onChange={(e) => {
            const n = Math.max(0, Math.min(100, Math.floor(Number(e.target.value) || 0)));
            const next = [...lines];
            if (next[idx]) next[idx] = { ...next[idx], discount: n } as any;
            setLines(next);
          }}
          inputProps={{ inputMode: 'numeric', pattern: '[0-9]*', min: 0, max: 100 }}
          sx={{ '& input': { textAlign: 'right' } }}
        />
      );
    } },
    { field: 'discountedPrice', headerName: 'Discounted Price', width: 140, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => {
      const gLower = String(p?.row?.groupName || '').trim().toLowerCase();
      const unit = Number(groupPriceByName?.[gLower]);
      if (!Number.isFinite(unit)) return '';
      const idx = Math.max(0, Number(p?.id) - 1);
      const disc = Math.max(0, Math.min(100, Number(lines[idx]?.discount ?? 0)));
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
    { field: 'subTotal', headerName: 'Sub Total', width: 120, type: 'number', align: 'right', headerAlign: 'right', renderCell: (p: any) => {
      const gLower = String(p?.row?.groupName || '').trim().toLowerCase();
      const unit = Number(groupPriceByName?.[gLower]);
      if (!Number.isFinite(unit)) return '';
      const idx = Math.max(0, Number(p?.id) - 1);
      const disc = Math.max(0, Math.min(100, Number(lines[idx]?.discount ?? 0)));
      const qty = Math.max(0, Math.floor(Number(lines[idx]?.qty || 0)));
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
              setLines(next);
            }} title="Remove">
              <DeleteIcon fontSize="small" />
            </IconButton>
          </Stack>
        );
      },
    },
  ], [lines, palletDescByGroup, groupPriceByName]);

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
            ))),
            discount: (() => {
              const v = Number(l?.discount ?? l?.discountPercent ?? 0);
              if (!Number.isFinite(v)) return 0;
              return Math.max(0, Math.min(100, Math.floor(v)));
            })(),
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

    const rows = Array.isArray(lines) ? lines : [];
    if (rows.length === 0) errs.push('Please add at least one pallet');
    const anyQty = rows.some(l => Number(l.qty) > 0);
    if (!anyQty) errs.push('Please add at least one pallet with quantity > 0');
    const anyInvalid = rows.some(l => !Number.isFinite(Number(l.qty)) || Number(l.qty) <= 0);
    if (anyInvalid) errs.push('Each pallet quantity must be > 0');

    if (errs.length) { toast.error(errs[0]); return; }

    try {
      const normalizedLines = lines.map(l => ({
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
        customerEmail,
        customerName,
        customerPhone,
        shippingAddress,
        originalPrice: computedSubTotal,
        shippingPercent,
        discountPercent: '',
        notes,
        // multiple aliases for backend compatibility
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
    setCreatedAt(new Date().toISOString().slice(0,10));
    setEstFulfillment('');
    setEstDelivered('');
    setOriginalPrice('');
    setShippingPercent('');
    setDiscountPercent('');
    setNotes('');
    setLastUpdatedAt('');
    setLastUpdatedBy('');
    setLines([]);
  }, []);

  return (
    <Container maxWidth="lg" sx={{ mt: 2, mb: 4 }}>
      <Typography variant="h5" sx={{ mb: 2, fontWeight: 700 }}>EARLY BUY</Typography>
      <Paper sx={{ p: 2, mb: 2 }}>
        <Stack direction="row" spacing={1} alignItems="center">
          <Button variant="contained" onClick={() => { resetEarlyBuyForm(); setOpen(true); }}>Add Order</Button>
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
              setCreatedAt(row.createdAt);
              setContainerArrival(String((row as any).containerArrival || ''));
              setEstFulfillment(row.estFulfillment);
              setEstDelivered(row.estDelivered);
              setOriginalPrice(row.originalPrice || '');
              setShippingPercent(row.shippingPercent || '');
              setDiscountPercent(row.discountPercent || '');
              setNotes(row.notes || '');
              setLastUpdatedAt(String(row.updatedAt || ''));
              setLastUpdatedBy(String(row.updatedBy || ''));
              setLines(Array.isArray(row.lines) ? row.lines : []);
              setOpen(true);
            }}
          />
        </div>
      </Paper>

      <Dialog open={open} onClose={() => { setOpen(false); resetEarlyBuyForm(); }} fullWidth maxWidth="lg">
        <DialogTitle>{editingOrder ? `Edit Order - ${editingOrder.id}` : 'New Early Buy Order'}</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <Stack direction={{ xs:'column', sm:'row' }} spacing={2}>
              <TextField size="small" label="Warehouse" value="MPG" disabled fullWidth />
              {/* Status is always editable */}
              <TextField size="small" select label="Status" value={status} onChange={(e)=> handleStatusChange(e.target.value)} fullWidth disabled={!isStatusEditable}>
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
              <TextField size="small" required type="date" label="Container Arrival" value={containerArrival} onChange={(e)=> setContainerArrival(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isContainerArrivalEditable} />
              <TextField size="small" required type="date" label="Estimated ShipDate for Customer" value={estFulfillment} onChange={(e)=> setEstFulfillment(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isEditable} />
              <TextField size="small" type="date" label="Estimated Arrival Date" value={estDelivered} onChange={(e)=> setEstDelivered(e.target.value)} InputLabelProps={{ shrink: true }} fullWidth disabled={!isArrivalDateEditable} />
            </Stack>
            <Stack direction={{ xs:'column', md:'row' }} spacing={2}>
              <TextField size="small" label="Sub Total" value={computedSubTotal} fullWidth disabled />
              <TextField size="small" label="Shipping Charges (%)" value={shippingPercent} onChange={(e)=> setShippingPercent(e.target.value)} fullWidth disabled={!isEditable} />
              <TextField size="small" label="Grand Total" value={computedGrandTotal} fullWidth disabled />
            </Stack>
            <TextField size="small" label="Remarks/Notes" value={notes} onChange={(e)=> setNotes(e.target.value)} fullWidth multiline minRows={3} disabled={!isEditable} />
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
                columns={linesColumns.map(c => c.field === 'actions' ? { ...c, renderCell: (p:any) => (!isEditable ? null : (c as any).renderCell(p)) } : (c.field === 'qty' ? { ...c, renderCell: (p:any) => {
                  if (!isEditable) return <span>{Number(lines[Math.max(0, Number(p?.id)-1)]?.qty || 0)}</span>;
                  return (c as any).renderCell(p);
                } } : c))}
                onRowDoubleClick={(p: any) => {
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
