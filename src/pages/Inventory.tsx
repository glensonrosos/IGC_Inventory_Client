import { useEffect, useMemo, useState } from 'react';
import { Container, Typography, Paper, Button, Stack, TextField, Dialog, DialogTitle, DialogContent, DialogActions, Chip, IconButton, MenuItem } from '@mui/material';
import Autocomplete from '@mui/material/Autocomplete';
import { DataGrid, GridColDef } from '@mui/x-data-grid';
import api from '../api';
import PalletImportModal from '../components/PalletImportModal';
import { useToast } from '../components/ToastProvider';
import AddIcon from '@mui/icons-material/Add';
import DeleteIcon from '@mui/icons-material/Delete';

type Warehouse = { _id: string; name: string };
type GroupOverview = { groupName: string; perWarehouse: { warehouseId: string; pallets: number }[]; totalPallets: number };
type GroupDetails = {
  groupName: string;
  perWarehouse: { warehouseId: string; pallets: number }[];
  items: Array<{ itemCode: string; description: string; color: string; packSize: number; totalPallets: number; perWarehouse: Record<string, { pallets: number; qty: number }>; totalQty: number }>;
  recentTransactions: any[];
};

type ItemGroupOption = { _id: string; name: string };

export default function Inventory() {
  const [warehouses, setWarehouses] = useState<Warehouse[]>([]);
  const [groups, setGroups] = useState<GroupOverview[]>([]);
  const [search, setSearch] = useState('');
  const [importOpen, setImportOpen] = useState(false);
  const [lossOpen, setLossOpen] = useState(false);
  const [lossWarehouse, setLossWarehouse] = useState('');
  const [lossReference, setLossReference] = useState('');
  const [groupOptions, setGroupOptions] = useState<ItemGroupOption[]>([]);
  const [lossItems, setLossItems] = useState<Array<{ group: ItemGroupOption | null; qty: string }>>([
    { group: null, qty: '' },
  ]);
  const [loading, setLoading] = useState(false);
  const [detailsOpen, setDetailsOpen] = useState(false);
  const [selectedGroup, setSelectedGroup] = useState<string>('');
  const [details, setDetails] = useState<GroupDetails | null>(null);
  const toast = useToast();

  const loadWarehouses = async () => {
    try {
      const { data } = await api.get<Warehouse[]>('/warehouses');
      setWarehouses(data || []);
    } catch (e:any) {
      // silent
    }
  };

  const getLossGroupName = (g: any) => String(g?.name || '').trim();
  const getLossRowError = (idx: number) => {
    const current = getLossGroupName(lossItems[idx]?.group);
    if (!current) return '';
    const matches = lossItems.filter((it, i) => i !== idx && getLossGroupName(it?.group).toLowerCase() === current.toLowerCase());
    return matches.length ? 'Duplicate pallet description' : '';
  };

  const getLossRowOptions = (idx: number) => {
    const current = getLossGroupName(lossItems[idx]?.group);
    const selected = new Set(
      lossItems
        .filter((_, i) => i !== idx)
        .map((it) => getLossGroupName(it?.group).toLowerCase())
        .filter((v) => v)
    );
    return (groupOptions || []).filter((o: any) => {
      const name = String(o?.name || '').trim();
      if (!name) return false;
      if (current && name.toLowerCase() === current.toLowerCase()) return true;
      return !selected.has(name.toLowerCase());
    });
  };

  const loadGroups = async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (search.trim()) params.set('search', search.trim());
      const { data } = await api.get<GroupOverview[]>(`/pallet-inventory/groups?${params.toString()}`);
      setGroups(data || []);
    } catch (e:any) {
      toast.error(e?.response?.data?.message || 'Failed to load');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { loadWarehouses(); }, []);
  useEffect(() => { loadGroups(); }, [search]);

  const columns: GridColDef[] = useMemo(() => {
    const cols: GridColDef[] = [
      { field: 'groupName', headerName: 'Pallet Description', flex: 2, minWidth: 240 },
      { field: 'totalPallets', headerName: 'Total Pallets', type: 'number', align: 'right', headerAlign: 'right', width: 140 },
    ];
    for (const w of warehouses) {
      cols.push({
        field: `wh_${w._id}`,
        headerName: w.name,
        width: 140,
        type: 'number',
        align: 'right',
        headerAlign: 'right',
      });
    }
    cols.push({
      field: 'action', headerName: 'Action', width: 120, sortable: false, filterable: false,
      renderCell: (params) => (
        <Button size="small" onClick={() => onView(params.row.groupName)}>View</Button>
      )
    });
    return cols;
  }, [warehouses]);

  const rows = useMemo(() => {
    return groups.map((g) => {
      const row: any = { id: g.groupName, groupName: g.groupName, totalPallets: g.totalPallets || 0 };
      for (const w of warehouses) {
        const found = g.perWarehouse.find(p => String(p.warehouseId) === String(w._id));
        row[`wh_${w._id}`] = found ? found.pallets : 0;
      }
      return row;
    });
  }, [groups, warehouses]);

  const onView = async (groupName: string) => {
    setSelectedGroup(groupName);
    setDetailsOpen(true);
    setDetails(null);
    try {
      const { data } = await api.get<GroupDetails>(`/pallet-inventory/groups/${encodeURIComponent(groupName)}`);
      setDetails(data);
    } catch (e:any) {
      toast.error(e?.response?.data?.message || 'Failed to load details');
    }
  };

  const openLoss = async () => {
    setLossOpen(true);
    setLossWarehouse('');
    setLossReference('');
    setLossItems([{ group: null, qty: '' }]);
    try {
      const { data } = await api.get<ItemGroupOption[]>('/item-groups');
      setGroupOptions(Array.isArray(data) ? data : []);
    } catch {
      setGroupOptions([]);
    }
  };

  const submitLoss = async () => {
    if (!lossWarehouse) {
      toast.error('Warehouse is required');
      return;
    }
    const normalized = lossItems
      .map((it) => ({ groupName: it.group?.name || '', qty: Number(it.qty) }))
      .filter((it) => it.groupName);
    if (normalized.length === 0) {
      toast.error('At least one Pallet Description is required');
      return;
    }
    for (const it of normalized) {
      if (!Number.isFinite(it.qty) || it.qty <= 0) {
        toast.error('Qty must be > 0');
        return;
      }
    }
    const seen = new Set<string>();
    for (const it of normalized) {
      const key = it.groupName.toLowerCase();
      if (seen.has(key)) {
        toast.error(`Duplicate Pallet Description: ${it.groupName}`);
        return;
      }
      seen.add(key);
    }

    try {
      await api.post('/pallet-inventory/adjustments', {
        warehouseId: lossWarehouse,
        reference: lossReference.trim() || undefined,
        items: normalized,
      });
      toast.success('Loss recorded');
      setLossOpen(false);
      await loadGroups();
      if (detailsOpen && selectedGroup) {
        await onView(selectedGroup);
      }
    } catch (e:any) {
      toast.error(e?.response?.data?.message || 'Failed to record loss');
    }
  };

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Pallet Inventory</Typography>
      <Paper sx={{ p:2, mb:2 }}>
        <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} alignItems="center" justifyContent="space-between">
          <TextField
            label="Search Pallet Description / Item / Color"
            placeholder="Type to search..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            sx={{ minWidth: 320, flex: 1 }}
          />
          <Stack direction="row" spacing={1}>
            <Button variant="outlined" onClick={loadGroups}>Refresh</Button>
            <Button variant="outlined" onClick={openLoss}>Record Loss</Button>
            <Button variant="contained" onClick={()=>setImportOpen(true)}>Import Existing Inventory</Button>
          </Stack>
        </Stack>
      </Paper>
      <Paper sx={{ height: 600, width: '100%', p:1 }}>
        <DataGrid
          rows={rows}
          columns={columns}
          loading={loading}
          disableRowSelectionOnClick
          initialState={{ pagination: { paginationModel: { pageSize: 20, page: 0 } } }}
          pageSizeOptions={[10,20,50,100]}
          density="compact"
        />
      </Paper>
      <PalletImportModal open={importOpen} onClose={()=>setImportOpen(false)} onImported={loadGroups} />

      <Dialog open={lossOpen} onClose={()=>setLossOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Record Loss (Inventory Adjustment)</DialogTitle>
        <DialogContent sx={{ pt: 2 }}>
          <TextField
            select
            fullWidth
            label="Warehouse (required)"
            value={lossWarehouse}
            onChange={(e)=> setLossWarehouse(e.target.value)}
            sx={{ mb: 2 }}
          >
            <MenuItem value="">Select warehouse</MenuItem>
            {warehouses.map((w) => (
              <MenuItem key={w._id} value={w._id}>{w.name}</MenuItem>
            ))}
          </TextField>
          <TextField
            fullWidth
            label="Reference (optional)"
            value={lossReference}
            onChange={(e)=> setLossReference(e.target.value)}
            sx={{ mb: 2 }}
          />

          <Stack spacing={1}>
            {lossItems.map((row, idx) => (
              <Stack
                key={idx}
                direction="row"
                spacing={1}
                alignItems="center"
                sx={{ p: 0.5, borderRadius: 1, bgcolor: getLossRowError(idx) ? '#fff4f4' : 'transparent' }}
              >
                <Autocomplete
                  options={getLossRowOptions(idx)}
                  getOptionLabel={(o)=> o?.name || ''}
                  value={row.group}
                  onChange={(_, v)=> {
                    setLossItems((prev)=> prev.map((p, i)=> i === idx ? { ...p, group: v } : p));
                  }}
                  renderInput={(params)=> (
                    <TextField
                      {...params}
                      label="Pallet Description"
                      size="small"
                      error={Boolean(getLossRowError(idx))}
                      helperText={getLossRowError(idx) || ''}
                    />
                  )}
                  sx={{ flex: 1 }}
                />
                <TextField
                  size="small"
                  type="number"
                  label="Qty"
                  value={row.qty}
                  onChange={(e)=> {
                    const v = e.target.value;
                    setLossItems((prev)=> prev.map((p, i)=> i === idx ? { ...p, qty: v } : p));
                  }}
                  sx={{ width: 120 }}
                />
                <IconButton
                  aria-label="Remove"
                  disabled={lossItems.length === 1}
                  onClick={()=> setLossItems((prev)=> prev.filter((_, i)=> i !== idx))}
                >
                  <DeleteIcon />
                </IconButton>
              </Stack>
            ))}
            <Button
              variant="outlined"
              startIcon={<AddIcon />}
              onClick={()=> setLossItems((prev)=> [...prev, { group: null, qty: '' }])}
              sx={{ alignSelf: 'flex-start' }}
            >
              Add Pallet Description
            </Button>
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={()=>setLossOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={submitLoss}>Save</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={detailsOpen} onClose={()=>setDetailsOpen(false)} fullWidth maxWidth="lg">
        <DialogTitle>{selectedGroup || 'Group'} — Details</DialogTitle>
        <DialogContent>
          {details ? (
            <>
              <Stack direction="row" spacing={1} sx={{ mb: 1, flexWrap: 'wrap' }}>
                <Chip
                  color="primary"
                  size="small"
                  label={`Total Pallets: ${details.perWarehouse.reduce((s, p) => s + (p.pallets || 0), 0)}`}
                />
                {warehouses.map((w) => {
                  const pallets = details.perWarehouse.find(p=>String(p.warehouseId)===String(w._id))?.pallets || 0;
                  return (
                    <Chip key={w._id} size="small" label={`${w.name}: ${pallets}`} />
                  );
                })}
              </Stack>
              <Typography variant="subtitle1" sx={{ mb:1 }}>Items per Pallet</Typography>
              <div style={{ height: 360, width: '100%' }}>
                <DataGrid
                  rows={details.items.map((it:any)=>{
                    const r: any = { id: it.itemCode, ...it };
                    return r;
                  })}
                  columns={([
                    { field: 'itemCode', headerName: 'Item Code', flex: 1, minWidth: 140 },
                    { field: 'description', headerName: 'Description', flex: 3, minWidth: 320 },
                    { field: 'color', headerName: 'Color', flex: 1, minWidth: 100 },
                    { field: 'packSize', headerName: 'Pack Size', width: 110, type: 'number', align: 'right', headerAlign: 'right' },
                  ]) as GridColDef[]}
                  disableRowSelectionOnClick
                  density="compact"
                />
              </div>
              <Typography variant="subtitle1" sx={{ mt:2, mb:1 }}>Recent Transactions</Typography>
              <div style={{ height: 300, width: '100%' }}>
                <DataGrid
                  rows={details.recentTransactions.map((t:any, i:number) => {
                    let statusLabel = t.status;
                    if (t.status === 'On-Water' && t.committedBy === 'transfer') {
                      statusLabel = 'on_transfer | on_water';
                    }
                    if (t.status === 'Adjustment') {
                      const rsn = String(t.reason || '').toLowerCase();
                      if (rsn === 'loss') statusLabel = 'adjustment | loss';
                      else if (rsn === 'order_fulfilled') statusLabel = 'adjustment | order_fulfilled';
                      else statusLabel = 'adjustment';
                    }
                    if (t.status === 'Delivered') {
                      if (t.committedBy === 'existing_inventory') statusLabel = 'existing_inventory | Delivered';
                      else if (t.committedBy === 'transfer') statusLabel = 'transfered | Delivered';
                      else statusLabel = t.wasOnWater ? 'On-Water | Delivered' : 'on_process | Delivered';
                    }
                    return {
                      id: i,
                      date: new Date(t.createdAt).toLocaleString(),
                      poNumber: t.poNumber,
                      warehouse: warehouses.find(w => String(w._id) === String(t.warehouseId))?.name || '-',
                      type: (Number(t.palletsDelta) || 0) >= 0 ? 'IN' : 'OUT',
                      status: statusLabel,
                      pallets: t.palletsDelta,
                    };
                  })}
                  columns={([
                    { field: 'date', headerName: 'Date/Time', flex: 1.5, minWidth: 80 },
                    { field: 'poNumber', headerName: 'PO #', width: 100 },
                    { field: 'warehouse', headerName: 'Warehouse', flex: 1, minWidth: 140 },
                    { field: 'type', headerName: 'Type', width: 100 },
                    { field: 'status', headerName: 'Status', width: 220 },
                    { field: 'pallets', headerName: 'Pallets', type: 'number', align: 'right', headerAlign: 'right', width: 110 },
                  ]) as GridColDef[]}
                  disableRowSelectionOnClick
                  density="compact"
                />
              </div>
            </>
          ) : (
            <Typography variant="body2">Loading...</Typography>
          )}
        </DialogContent>
        <DialogActions>
          <Button onClick={()=>setDetailsOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
