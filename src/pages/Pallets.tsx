import { useCallback, useEffect, useMemo, useState } from 'react';
import { Container, Typography, Paper, Stack, Button, TextField, Dialog, DialogTitle, DialogContent, DialogActions, LinearProgress } from '@mui/material';
import { DataGrid, GridColDef } from '@mui/x-data-grid';
import * as XLSX from 'xlsx';
import api from '../api';

type SummaryWarehouse = { _id: string; name: string };
type SummaryRow = {
  itemGroup: string;
  warehouses: Record<string, number>;
  onProcessQty: number;
  onWaterQty: number;
};
type SummaryResponse = { warehouses: SummaryWarehouse[]; rows: SummaryRow[] };

type ItemGroupRow = { name: string; lineItem?: string };

export default function Pallets() {
  const [loading, setLoading] = useState(false);
  const [warehouses, setWarehouses] = useState<SummaryWarehouse[]>([]);
  const [rows, setRows] = useState<SummaryRow[]>([]);
  const [palletIdByGroup, setPalletIdByGroup] = useState<Record<string, string>>({});
  const [q, setQ] = useState('');

  const [onWaterOpen, setOnWaterOpen] = useState(false);
  const [onWaterLoading, setOnWaterLoading] = useState(false);
  const [onWaterGroupName, setOnWaterGroupName] = useState('');
  const [onWaterRows, setOnWaterRows] = useState<any[]>([]);

  const [onProcessOpen, setOnProcessOpen] = useState(false);
  const [onProcessLoading, setOnProcessLoading] = useState(false);
  const [onProcessGroupName, setOnProcessGroupName] = useState('');
  const [onProcessRows, setOnProcessRows] = useState<any[]>([]);

  const load = async () => {
    setLoading(true);
    try {
      const [summaryResp, groupsResp] = await Promise.all([
        api.get<SummaryResponse>('/reports/pallet-summary-by-group'),
        api.get<ItemGroupRow[]>('/item-groups'),
      ]);
      const data = summaryResp?.data as any;
      setWarehouses(Array.isArray(data?.warehouses) ? data.warehouses : []);
      setRows(Array.isArray(data?.rows) ? data.rows : []);

      const groups = Array.isArray(groupsResp?.data) ? groupsResp.data : [];
      const map: Record<string, string> = {};
      for (const g of groups) {
        const name = String((g as any)?.name || '').trim();
        if (!name) continue;
        map[name] = String((g as any)?.lineItem || '').trim();
      }
      setPalletIdByGroup(map);
    } catch {
      setWarehouses([]);
      setRows([]);
      setPalletIdByGroup({});
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

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
      setOnWaterRows(Array.isArray((data as any)?.rows) ? (data as any).rows : []);
    } catch {
      setOnWaterRows([]);
    } finally {
      setOnWaterLoading(false);
    }
  }, []);

  const openOnProcessDetails = useCallback(async ({ groupName }: { groupName: string }) => {
    const g = String(groupName || '').trim();
    if (!g) return;
    setOnProcessOpen(true);
    setOnProcessGroupName(g);
    setOnProcessRows([]);
    setOnProcessLoading(true);
    try {
      const { data } = await api.get('/orders/pallet-picker/on-process', { params: { groupName: g } });
      setOnProcessRows(Array.isArray((data as any)?.rows) ? (data as any).rows : []);
    } catch {
      setOnProcessRows([]);
    } finally {
      setOnProcessLoading(false);
    }
  }, []);

  const filteredRows = useMemo(() => {
    const t = q.trim().toLowerCase();
    if (!t) return rows;
    return rows.filter((r) => {
      const groupName = String(r.itemGroup || '').trim();
      const name = groupName.toLowerCase();
      const pid = String(palletIdByGroup[groupName] || '').trim().toLowerCase();
      return (name && name.includes(t)) || (pid && pid.includes(t));
    });
  }, [rows, q, palletIdByGroup]);

  const columns = useMemo<GridColDef[]>(() => {
    const primaryWh = warehouses.find((w) => Boolean((w as any)?.isPrimary)) || null;
    const secondWh = warehouses.find((w) => !Boolean((w as any)?.isPrimary)) || null;
    const primaryCols: GridColDef[] = primaryWh
      ? [
          {
            field: `wh_${primaryWh._id}`,
            headerName: `${primaryWh.name}`,
            width: 170,
            type: 'number',
          },
        ]
      : [];
    const secondCols: GridColDef[] = secondWh
      ? [
          {
            field: `wh_${secondWh._id}`,
            headerName: `${secondWh.name}`,
            width: 170,
            type: 'number',
          },
        ]
      : [];
    const otherWhCols: GridColDef[] = warehouses
      .filter((w) => String(w?._id) !== String(primaryWh?._id || '') && String(w?._id) !== String(secondWh?._id || ''))
      .map((w) => ({
        field: `wh_${w._id}`,
        headerName: `${w.name}`,
        width: 170,
        type: 'number',
      }));
    return [
      { field: 'palletId', headerName: 'Pallet ID', width: 160 },
      { field: 'itemGroup', headerName: 'Pallet Description', flex: 1, minWidth: 220 },
      ...primaryCols,
      {
        field: 'onWaterQty',
        headerName: 'On-Water',
        width: 140,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
        renderCell: (p: any) => {
          const qty = Number((p?.row as any)?.onWaterQty ?? 0);
          if (!qty) return '0';
          const groupName = String((p?.row as any)?.itemGroup || '').trim();
          const wid = String((primaryWh as any)?._id || '').trim();
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
      ...secondCols,
      {
        field: 'onProcessQty',
        headerName: 'On-Process',
        width: 150,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
        renderCell: (p: any) => {
          const qty = Number((p?.row as any)?.onProcessQty ?? 0);
          if (!qty) return '0';
          const groupName = String((p?.row as any)?.itemGroup || '').trim();
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
      ...otherWhCols,
      { field: 'totalQty', headerName: 'Total Qty', width: 140, type: 'number' },
    ];
  }, [warehouses, openOnProcessDetails, openOnWaterDetails]);

  const gridRows = useMemo(() => (
    filteredRows.map((r) => {
      const wh: Record<string, number> = r.warehouses || {};
      const flat: Record<string, any> = {};
      for (const w of warehouses) {
        flat[`wh_${w._id}`] = Number(wh[String(w._id)] || 0);
      }
      const warehousesTotal = warehouses.reduce((sum, w) => sum + Number(wh[String(w._id)] || 0), 0);
      const totalQty = warehousesTotal + Number(r.onProcessQty || 0) + Number(r.onWaterQty || 0);
      const groupName = String(r.itemGroup || '');
      return {
        id: r.itemGroup,
        palletId: String(palletIdByGroup[groupName] || ''),
        itemGroup: r.itemGroup,
        warehouses: wh,
        onProcessQty: Number(r.onProcessQty || 0),
        onWaterQty: Number(r.onWaterQty || 0),
        totalQty,
        ...flat,
      };
    })
  ), [filteredRows, warehouses, palletIdByGroup]);

  const exportExcel = () => {
    const primaryWh = warehouses.find((w) => Boolean((w as any)?.isPrimary)) || null;
    const secondWh = warehouses.find((w) => !Boolean((w as any)?.isPrimary)) || null;
    const otherWh = warehouses.filter((w) => String(w?._id) !== String(primaryWh?._id || '') && String(w?._id) !== String(secondWh?._id || ''));
    const header = [
      'Pallet ID',
      'Pallet Description',
      ...(primaryWh ? [`${primaryWh.name}`] : []),
      'On-Water',
      ...(secondWh ? [`${secondWh.name}`] : []),
      'On-Process',
      ...otherWh.map(w => `${w.name}`),
      'Total Qty',
    ];
    const aoa = gridRows.map((r:any) => [
      String(r.palletId || ''),
      r.itemGroup,
      ...(primaryWh ? [Number(r.warehouses?.[String(primaryWh._id)] || 0)] : []),
      Number(r.onWaterQty || 0),
      ...(secondWh ? [Number(r.warehouses?.[String(secondWh._id)] || 0)] : []),
      Number(r.onProcessQty || 0),
      ...otherWh.map(w => Number(r.warehouses?.[String(w._id)] || 0)),
      Number(r.totalQty || 0),
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...aoa]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pallets');
    const pad = (n:number)=> n.toString().padStart(2,'0');
    const d = new Date();
    const ts = `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
    XLSX.writeFile(wb, `Pallet_Quantity_Summary_${ts}.xlsx`);
  };

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Pallets Summary</Typography>
      <Paper sx={{ p:2 }}>
        <Stack direction={{ xs: 'column', sm: 'row' }} spacing={1} alignItems="center" justifyContent="space-between" sx={{ mb: 2 }}>
          <TextField size="small" label="Search Pallet ID or Description" value={q} onChange={(e)=> setQ(e.target.value)} sx={{ minWidth: 260, flex: 1 }} />
          <Stack direction="row" spacing={1}>
            <Button variant="outlined" onClick={load} disabled={loading}>Refresh</Button>
            <Button variant="contained" onClick={exportExcel} disabled={!gridRows.length}>Export List</Button>
          </Stack>
        </Stack>
        <div style={{ height: 560, width: '100%' }}>
          <DataGrid
            rows={gridRows}
            columns={columns}
            loading={loading}
            disableRowSelectionOnClick
            density="compact"
            pageSizeOptions={[10,20,50,100]}
          />
        </div>
      </Paper>

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
