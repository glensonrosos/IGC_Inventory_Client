import { useEffect, useState, useRef } from 'react';
import { Container, Typography, Paper, Stack, TextField, Button, MenuItem, Dialog, DialogTitle, DialogContent, DialogActions, IconButton, Tooltip, Autocomplete, Menu, FormControlLabel, Checkbox, Chip, Divider, Grid } from '@mui/material';
import MoreVertIcon from '@mui/icons-material/MoreVert';
import { DataGrid, GridColDef, GridRenderCellParams, GridRowSelectionModel } from '@mui/x-data-grid';
import api from '../api';
import { useToast } from '../components/ToastProvider';
import * as XLSX from 'xlsx';

interface ItemGroup { _id: string; name: string }
interface Item { _id: string; itemCode: string; itemGroup: string; description: string; color: string; packSize?: number; enabled?: boolean }

export default function ItemRegistry() {
  const [groups, setGroups] = useState<ItemGroup[]>([]);
  const [items, setItems] = useState<Item[]>([]);
  const [loading, setLoading] = useState(false);
  const toast = useToast();

  const [renameLineItemOpen, setRenameLineItemOpen] = useState(false);
  const [renameLineItemGroupId, setRenameLineItemGroupId] = useState('');
  const [renameLineItemGroupName, setRenameLineItemGroupName] = useState('');
  const [renameLineItemValue, setRenameLineItemValue] = useState('');

  const [groupSelection, setGroupSelection] = useState<GridRowSelectionModel>({ type: 'include', ids: new Set() } as any);

  // Group form
  const [groupName, setGroupName] = useState('');
  const [groupFile, setGroupFile] = useState<File | null>(null);
  const [groupImportResult, setGroupImportResult] = useState<any>(null);
  const groupFileInputRef = useRef<HTMLInputElement | null>(null);

  // Item form
  const [itemCode, setItemCode] = useState('');
  const [itemGroup, setItemGroup] = useState('');
  const [description, setDescription] = useState('');
  const [color, setColor] = useState('');
  const [packSize, setPackSize] = useState<number | ''>('');
  const [enabled, setEnabled] = useState<boolean>(true);
  // Group detail (per-pallet-group items management)
  const [selectedGroup, setSelectedGroup] = useState<string>('');
  const [gItemCode, setGItemCode] = useState('');
  const [gDesc, setGDesc] = useState('');
  const [gColor, setGColor] = useState('');
  const [gPack, setGPack] = useState<number>(0);
  // Import
  const [file, setFile] = useState<File | null>(null);
  const [importResult, setImportResult] = useState<any>(null);
  // Search
  const [groupsQ, setGroupsQ] = useState('');
  const [itemsQ, setItemsQ] = useState('');
  const [enabledFilter, setEnabledFilter] = useState<'all'|'enabled'|'disabled'>('all');
  // Edit dialog
  const [editOpen, setEditOpen] = useState(false);
  const [editItemCode, setEditItemCode] = useState('');
  const [editGroup, setEditGroup] = useState('');
  const [editDesc, setEditDesc] = useState('');
  const [editColor, setEditColor] = useState('');
  const [editPack, setEditPack] = useState<number | ''>('');
  const [editEnabled, setEditEnabled] = useState<boolean>(true);

  const openEdit = (code: string) => {
    const it = items.find(i => i.itemCode === code);
    if (!it) return;
    setEditItemCode(it.itemCode);
    setEditGroup(it.itemGroup || '');
    setEditDesc(it.description || '');
    setEditColor(it.color || '');
    setEditPack(typeof (it as any).packSize === 'number' ? (it as any).packSize : '');
    setEditEnabled(typeof it.enabled === 'boolean' ? it.enabled : true);
    setEditOpen(true);
  };

  const manageGroup = (row: any) => {
    if (!row?.active) return;
    setGroupSelection({ type: 'include', ids: new Set([String(row.id)]) } as any);
    setSelectedGroup(row.name);
    setGItemCode('');
    setGDesc('');
    setGColor('');
    setGPack(0);
  };

  const openRenameLineItem = (row: any) => {
    setRenameLineItemGroupId(String(row?.id || ''));
    setRenameLineItemGroupName(String(row?.name || ''));
    setRenameLineItemValue(String(row?.lineItem || ''));
    setRenameLineItemOpen(true);
  };

  const saveRenameLineItem = async () => {
    if (!renameLineItemGroupId) return;
    try {
      await api.put(`/item-groups/${encodeURIComponent(renameLineItemGroupId)}`, { lineItem: renameLineItemValue });
      setGroups(prev => prev.map((g: any) => String(g._id) === String(renameLineItemGroupId) ? { ...g, lineItem: renameLineItemValue } : g));
      toast.success('Pallet ID updated');
      setRenameLineItemOpen(false);
      load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || e?.message || 'Failed to update Pallet ID';
      toast.error(msg);
    }
  };

  function GroupActionsMenu({ row }: { row: any }) {
    const [anchorEl, setAnchorEl] = useState<null | HTMLElement>(null);
    const open = Boolean(anchorEl);
    const handleOpen = (e: React.MouseEvent<HTMLElement>) => setAnchorEl(e.currentTarget);
    const handleClose = () => setAnchorEl(null);
    const onManage = () => { manageGroup(row); handleClose(); };
    const onRenameLineItem = () => { openRenameLineItem(row); handleClose(); };
    const onToggle = async () => { await toggleGroupActive(row.id, !row.active); handleClose(); };
    const onDelete = async () => { await deleteGroup(row.id, row.name); handleClose(); };
    return (
      <>
        <IconButton size="small" onClick={handleOpen} aria-label="Actions"><MoreVertIcon fontSize="small" /></IconButton>
        <Menu anchorEl={anchorEl} open={open} onClose={handleClose} anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }} transformOrigin={{ vertical: 'top', horizontal: 'right' }}>
          <MenuItem onClick={onManage} disabled={!row.active}>Manage Items</MenuItem>
          <MenuItem onClick={onRenameLineItem}>Rename Pallet ID</MenuItem>
          <MenuItem onClick={onToggle}>{row.active ? 'Deactivate' : 'Activate'}</MenuItem>
          <MenuItem onClick={onDelete} disabled={row.itemCount > 0}>Delete</MenuItem>
        </Menu>
      </>
    );
  }

  const deleteItem = async (code: string, group?: string) => {
    if (!window.confirm(`Remove item "${code}"${group?` from ${group}`:''}? This cannot be undone.`)) return;
    try {
      const qs = group ? `?group=${encodeURIComponent(group)}` : '';
      await api.delete(`/items/${encodeURIComponent(code)}${qs}`);
      toast.success('Item removed');
      setItems(prev => prev.filter(it => {
        if (!group) return it.itemCode !== code;
        return !(it.itemCode === code && (it.itemGroup || '') === group);
      }));
      // refresh in background
      load();
    } catch (e:any) {
      const msg = e?.response?.data?.message || e?.message || 'Failed to delete item';
      toast.error(msg);
    }
  };

  const exportGroupsExcel = () => {
    // Export only active groups and enabled items, sorted
    const header = ['Pallet Description','Pallet ID','Item Code','Item Description','Color','Pack Size'];
    const lineItemByGroup = new Map(groups.map((g:any)=> [g.name, (g as any).lineItem || '']));
    const activeSet = new Set(groups.filter(g => (g as any).active !== false).map(g => g.name));
    const sizeOrder = (d: string) => {
      const m = (d || '').match(/\b(XXL|XL|L|M|S)\b/i);
      const v = m ? m[1].toUpperCase() : '';
      return v === 'S' ? 1 : v === 'M' ? 2 : v === 'L' ? 3 : v === 'XL' ? 4 : v === 'XXL' ? 5 : 0;
    };
    const baseDesc = (d: string) => (d || '').replace(/\b(XXL|XL|L|M|S)\b/gi, '').trim();
    const rows = items
      .filter(it => (it.enabled ?? true) && activeSet.has(it.itemGroup || ''))
      .sort((a, b) => {
        const ga = (a.itemGroup || '').localeCompare(b.itemGroup || '');
        if (ga !== 0) return ga;
        const ca = (a.color || '').localeCompare(b.color || '');
        if (ca !== 0) return ca;
        const ba = baseDesc(a.description || '');
        const bb = baseDesc(b.description || '');
        const bd = ba.localeCompare(bb);
        if (bd !== 0) return bd;
        const sd = sizeOrder(a.description || '') - sizeOrder(b.description || '');
        if (sd !== 0) return sd;
        return 0;
      })
      .map(it => [it.itemGroup||'', lineItemByGroup.get(it.itemGroup||'') || '', it.itemCode, it.description||'', it.color||'', (it as any).packSize ?? 0]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Groups + Items');
    XLSX.writeFile(wb, 'pallet_groups.xlsx');
  };

  const saveEdit = async () => {
    if (!editItemCode) return;
    if (!editGroup.trim() || !editDesc.trim() || !editColor.trim()) { toast.error('All fields are required'); return; }
    const ps = packSizeFrom(editPack);
    if (ps === null) { toast.error('Pack Size must be a positive number'); return; }
    if (!window.confirm('Proceed to update this item?')) return;
    await api.put(`/items/${editItemCode}`, { itemGroup: editGroup, description: editDesc, color: editColor, packSize: ps, enabled: editEnabled });
    // Optimistic local update so the row reflects immediately
    setItems(prev => prev.map(it => it.itemCode === editItemCode ? { ...it, itemGroup: editGroup, description: editDesc, color: editColor, packSize: ps, enabled: editEnabled } as any : it));
    toast.success('Item updated');
    setEditOpen(false);
    // Refresh in background to keep in sync
    load();
  };


  const exportItemsExcel = () => {
    const header = ['Pallet Description','Pallet ID','Item Code','Item Description','Color','Pack Size'];
    const lineItemByGroup = new Map(groups.map((g:any)=> [g.name, (g as any).lineItem || '']));
    const activeSet = new Set(groups.filter(g => (g as any).active !== false).map(g => g.name));
    const sizeOrder = (d: string) => {
      const m = (d || '').match(/\b(XXL|XL|L|M|S)\b/i);
      const v = m ? m[1].toUpperCase() : '';
      return v === 'S' ? 1 : v === 'M' ? 2 : v === 'L' ? 3 : v === 'XL' ? 4 : v === 'XXL' ? 5 : 0;
    };
    const baseDesc = (d: string) => (d || '').replace(/\b(XXL|XL|L|M|S)\b/gi, '').trim();
    const rows = items
      .filter(it => (it.enabled ?? true) && activeSet.has(it.itemGroup || ''))
      .sort((a, b) => {
        const ga = (a.itemGroup || '').localeCompare(b.itemGroup || '');
        if (ga !== 0) return ga;
        const ba = baseDesc(a.description || '');
        const bb = baseDesc(b.description || '');
        const bd = ba.localeCompare(bb);
        if (bd !== 0) return bd;
        const sd = sizeOrder(a.description || '') - sizeOrder(b.description || '');
        if (sd !== 0) return sd;
        return (a.color || '').localeCompare(b.color || '');
      })
      .map(it => [it.itemGroup||'', lineItemByGroup.get(it.itemGroup||'') || '', it.itemCode, it.description||'', it.color||'', (it as any).packSize ?? 0]);
    const data = [header, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Items');
    XLSX.writeFile(wb, 'items_master.xlsx');
  };

  const downloadTemplate = () => {
    const header = ['Pallet Description','Pallet ID','Item Code','Item Description','Color','Pack Size'];
    const example = [
      // Examples: last duplicate row per (Pallet Group + Item Code) overwrites values; identical duplicates are skipped
      ['Inverted Planters Smooth (OW)','Tall + Short Rounded Bottom Planters Pallet - Volcanic Ash Brown/Cement','PC2014AFBR-OW','Smooth Finish Inverted Planter S - Oyster White - Fiber Finish','Oyster White',8],
      ['Inverted Planters Mixed Smooth and VA (OW / MB / C / DAB)','32" Beaded Commercial Planter - Cement','MPC2032D-DAB','Volcanic Ash Texture Inverted Planter S - Dark Antique Bronze','Dark Antique Bronze',2]
    ];
    const ws = XLSX.utils.aoa_to_sheet([header, ...example]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'items_template.xlsx');
  };

  const load = async () => {
    setLoading(true);
    try {
      const [g, it] = await Promise.all([
        api.get<ItemGroup[]>('/item-groups'),
        api.get<Item[]>('/items', { params: { includeDisabled: 1 } })
      ]);
      setGroups(g.data);
      setItems(it.data);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  // Press Enter on Item Code to auto-fill description & color if code exists (pack stays 0)
  const handleGItemCodeKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const code = gItemCode.trim();
    if (!code) return;
    const existing = items.find(it => it.itemCode.toLowerCase() === code.toLowerCase());
    if (existing) {
      setGDesc(existing.description || '');
      setGColor(existing.color || '');
      setGPack(0);
      toast.info('Existing item found. Description and Color auto-filled.');
    } else {
      setGDesc('');
      setGColor('');
      setGPack(0);
      toast.warning('Item code not registered yet.');
    }
  };

  const deleteGroup = async (id: string, name: string) => {
    if (!window.confirm(`Delete pallet description "${name}"? This cannot be undone.`)) return;
    try {
      await api.delete(`/item-groups/${id}`);
      toast.success('Pallet Description deleted');
      await load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || e?.message || 'Failed to delete group';
      toast.error(msg);
    }
  };

  const toggleGroupActive = async (id: string, nextActive: boolean) => {
    try {
      await api.put(`/item-groups/${id}`, { active: nextActive });
      setGroups(prev => prev.map(g => g._id === id ? { ...g, active: nextActive } as any : g));
      toast.success(`Pallet Description ${nextActive ? 'activated' : 'deactivated'}`);
    } catch (e:any) {
      const msg = e?.response?.data?.message || e?.message || 'Failed to update group';
      toast.error(msg);
    }
  };

  const createGroup = async () => {
    if (!groupName.trim()) { toast.error('Pallet Description is required'); return; }
    try {
      await api.post('/item-groups', { name: groupName.trim() });
      toast.success('Pallet Description added');
      setGroupName('');
      await load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || e?.message || 'Failed to add group';
      toast.error(msg);
    }
  };

  const downloadGroupTemplate = () => {
    // Provide a template that includes items per pallet group, with Line Item column
    const header = ['Pallet Description','Pallet ID','Item Code','Item Description','Color','Pack Size'];
    const example = [
      ['Inverted Planters Smooth (OW)','Tall + Short Rounded Bottom Planters Pallet - Volcanic Ash Brown/Cement','PC2014AFBR-OW','Smooth Finish Inverted Planter S - Oyster White - Fiber Finish','Oyster White',8],
      ['Inverted Planters Mixed Smooth and VA (OW / MB / C / DAB)','32" Beaded Commercial Planter - Cement','PC2014AFBR-OW','Smooth Finish Inverted Planter S - Oyster White - Fiber Finish','Oyster White',2]
    ];
    const ws = XLSX.utils.aoa_to_sheet([header, ...example]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'pallet_groups_template.xlsx');
  };

  const importGroups = async () => {
    if (!groupFile) { toast.error('Please select a .xlsx file'); return; }
    const fd = new FormData();
    fd.append('file', groupFile);
    try {
      const { data } = await api.post('/item-groups/import', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      setGroupImportResult(data);
      const created = data.created ?? 0;
      const issues = data.errorCount ?? 0;
      if (issues) {
        toast.warning(`Import completed with issues. Created: ${created}, Issues: ${issues}. See details below.`);
      } else {
        toast.success(`Import completed. Created: ${created}.`);
      }
      setGroupFile(null);
      await load();
    } catch (e: any) {
      const res = e?.response?.data;
      const msg = res?.message || e?.message || 'Import failed';
      const issues = Number(res?.errorCount) || 0;
      // Surface server-provided detailed errors in the UI
      if (res) setGroupImportResult(res);
      else setGroupImportResult({ errorCount: 1, errors: [{ rowNum: '-', name: '-', errors: [msg] }] });
      if (issues > 0) toast.error(`Import rejected: ${issues} issue(s). See details below.`);
      else toast.error(msg);
    }
  };

  const packSizeFrom = (v: number | '' ): number | null => {
    const n = v === '' ? 0 : Number(v);
    if (!Number.isFinite(n) || n < 0) return null;
    return n;
  };

  const createItem = async () => {
    if (!itemCode.trim()) { toast.error('Item Code is required'); return; }
    if (!itemGroup.trim()) { toast.error('Pallet Description is required'); return; }
    if (!description.trim()) { toast.error('Item Description is required'); return; }
    if (!color.trim()) { toast.error('Color is required'); return; }
    const ps = packSizeFrom(packSize);
    if (ps === null) { toast.error('Pack Size must be a non-negative number'); return; }
    try {
      await api.post('/items', { itemCode: itemCode.trim(), itemGroup, description, color, packSize: ps, enabled });
      toast.success('Item created');
      setItemCode(''); setItemGroup(''); setDescription(''); setColor(''); setPackSize(''); setEnabled(true);
      await load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || e?.message || 'Failed to create item';
      toast.error(msg);
    }
  };

  const importItems = async () => {
    if (!file) { toast.error('Please select a .xlsx file'); return; }
    const fd = new FormData();
    fd.append('file', file);
    try {
      const { data } = await api.post('/items/registry-import', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      setImportResult(data);
      toast.success(`Import completed. Created: ${data.created}, Updated: ${data.updated}, Skipped: ${data.skipped}`);
      await load();
    } catch (e: any) {
      const res = e?.response?.data;
      const msg = res?.message || e?.message || 'Import failed';
      // Try to summarize common cause: missing or unregistered groups
      let extra = '';
      if (res?.errors && Array.isArray(res.errors)) {
        const groupErrors = res.errors.filter((er: any) => (er.errors||[]).some((s: string) => s.toLowerCase().includes('group')));
        if (groupErrors.length) {
          const firstFew = groupErrors.slice(0, 5).map((er: any) => `Row ${er.rowNum} (${er.itemCode||''})`).join(', ');
          extra = ` Possible cause: Some Item Groups are not registered. Affected: ${groupErrors.length}. ${firstFew}${groupErrors.length>5?'...':''}`;
        }
      }
      toast.error(msg + extra);
      setImportResult(res || { message: msg });
    }
  };

  function ActionsMenuCell({ row, onEdit }: { row: any, onEdit: ()=>void }) {
    const [anchorEl, setAnchorEl] = useState<null | HTMLElement>(null);
    const open = Boolean(anchorEl);
    const handleOpen = (e: React.MouseEvent<HTMLElement>) => setAnchorEl(e.currentTarget);
    const handleClose = () => setAnchorEl(null);
    const wrap = (fn: ()=>void) => () => { fn(); handleClose(); };
    return (
      <>
        <IconButton size="small" onClick={handleOpen} aria-label="Actions"><MoreVertIcon fontSize="small" /></IconButton>
        <Menu anchorEl={anchorEl} open={open} onClose={handleClose} anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }} transformOrigin={{ vertical: 'top', horizontal: 'right' }}>
          <MenuItem onClick={wrap(onEdit)}>Edit</MenuItem>
          <MenuItem onClick={wrap(()=> deleteItem(row.itemCode))}>Delete</MenuItem>
        </Menu>
      </>
    );
  }

  // Auto-populate item fields when entering an existing item code
  useEffect(() => {
    const code = itemCode.trim();
    if (!code) return;
    const existing = items.find(it => it.itemCode.toLowerCase() === code.toLowerCase());
    if (existing) {
      if (!description) setDescription(existing.description || '');
      if (!color) setColor(existing.color || '');
      if (packSize === '') setPackSize((existing as any).packSize ?? 0);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [itemCode]);

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Pallet Registry</Typography>

      <Dialog open={renameLineItemOpen} onClose={()=> setRenameLineItemOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Rename Pallet ID</DialogTitle>
        <DialogContent>
          <Typography variant="body2" sx={{ mb: 1 }}>Pallet Description: <b>{renameLineItemGroupName || '-'}</b></Typography>
          <TextField fullWidth label="Pallet ID" value={renameLineItemValue} onChange={(e)=> setRenameLineItemValue(e.target.value)} autoFocus />
        </DialogContent>
        <DialogActions>
          <Button onClick={()=> setRenameLineItemOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={saveRenameLineItem}>Save</Button>
        </DialogActions>
      </Dialog>

      <Paper sx={{ p:2, mb:2 }}>
        <Typography variant="h6" gutterBottom>Pallet Descriptions</Typography>
        <Grid container spacing={2} alignItems="center" sx={{ mb: 2 }}>
          <Grid item xs={12} md={6}>
            <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} alignItems={{ xs:'stretch', sm:'center' }}>
              <TextField fullWidth label="Pallet Description" value={groupName} onChange={e=>setGroupName(e.target.value)} />
              <Button variant="contained" onClick={createGroup} disabled={!groupName.trim()}>Add Pallet Description</Button>
            </Stack>
          </Grid>
          <Grid item xs={12} md={6}>
            <Stack direction="row" spacing={2} alignItems="center" justifyContent={{ xs:'flex-start', md:'flex-end' }}>
              <input ref={groupFileInputRef} hidden type="file" accept=".xlsx" onChange={(e)=>setGroupFile(e.target.files?.[0] || null)} data-component-name="ItemGroups" />
              <Button variant="outlined" onClick={()=> groupFileInputRef.current?.click()}>Select File</Button>
              <Button variant="contained" onClick={importGroups} disabled={!groupFile}>Import (.xlsx)</Button>
              <Button variant="text" onClick={downloadGroupTemplate}>Groups Template</Button>
              <Button variant="outlined" onClick={exportGroupsExcel}>Groups Export</Button>
            </Stack>
            {groupFile && (
              <Typography variant="caption" sx={{ display:'block', mt: 0.5 }}>Selected: {groupFile.name}</Typography>
            )}
          </Grid>
        </Grid>
        <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2} alignItems="center" justifyContent="space-between" sx={{ mb: 1 }}>
          <Typography variant="subtitle2">Pallet Descriptions: {groups.length}</Typography>
          <TextField size="small" label="Search pallet descriptions" value={groupsQ} onChange={(e)=>setGroupsQ(e.target.value)} sx={{ minWidth: 240 }} />
        </Stack>
        <div style={{ height: 420, width: '100%' }}>
          <DataGrid
            rows={groups
              .filter(g=>g.name.toLowerCase().includes(groupsQ.trim().toLowerCase()))
              .map(g=>({ id: g._id, name: g.name, lineItem: (g as any).lineItem || '', active: (g as any).active !== false, itemCount: items.filter(it => (it.itemGroup||'') === g.name).length }))}
            columns={([
              { field: 'lineItem', headerName: 'Pallet ID', flex: 1, minWidth: 220 },
              { field: 'name', headerName: 'Pallet Description', flex: 1, minWidth: 220 },
              { field: 'active', headerName: 'Active', width: 80, renderCell: (p: GridRenderCellParams) => (
                <Chip size="small" label={p.value ? 'Yes' : 'No'} color={p.value ? 'success' : 'default'} />
              ) },
              { field: 'itemCount', headerName: 'Items', width: 90 },
              { field: 'actions', headerName: 'Actions', width: 120, sortable: false, filterable: false, renderCell: (params: GridRenderCellParams) => {
                const r = params.row as any;
                return (<GroupActionsMenu row={r} />);
              }}
            ]) as GridColDef[]}
            loading={loading}
            disableRowSelectionOnClick
            density="compact"
            rowSelectionModel={groupSelection}
            onRowSelectionModelChange={(m: any) => {
              setGroupSelection(m);
            }}
            onRowDoubleClick={(params: any) => {
              const r = params?.row;
              if (!r) return;
              setGroupSelection({ type: 'include', ids: new Set([String(r.id)]) } as any);
              manageGroup(r);
            }}
            initialState={{ pagination: { paginationModel: { pageSize: 10, page: 0 } } }}
            pageSizeOptions={[10,20,50]}
            sx={{
              '& .MuiDataGrid-row.Mui-selected': {
                backgroundColor: 'rgba(25, 118, 210, 0.14)',
              },
              '& .MuiDataGrid-row.Mui-selected:hover': {
                backgroundColor: 'rgba(25, 118, 210, 0.20)',
              },
            }}
          />
        </div>
        {Boolean(selectedGroup) && (
          <Paper variant="outlined" sx={{ p:2, mt:2 }}>
            <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1 }}>
              <Typography variant="h6">Pallet Description</Typography>
              <Chip label={selectedGroup} color="success" size="medium" />
              <Typography variant="h6">— Items</Typography>
              {(() => { const gi = groups.find(g=>g.name===selectedGroup) as any; const li = gi?.lineItem || ''; return li ? <Chip label={li} size="small" sx={{ ml: 1 }} /> : null; })()}
            </Stack>
            <Grid container spacing={2} alignItems="center" sx={{ mb: 2 }}>
              <Grid item xs={12} md={3} sx={{mt:3}}>
                <TextField fullWidth required label="Item Code" value={gItemCode} onChange={e=>setGItemCode(e.target.value)} onKeyDown={handleGItemCodeKeyDown} helperText="Press Enter to check and auto-fill" />
              </Grid>
              <Grid item xs={12} md={5}>
                <TextField fullWidth required label="Item Description" value={gDesc} onChange={e=>setGDesc(e.target.value)} />
              </Grid>
              <Grid item xs={12} md={2}>
                <TextField fullWidth required label="Color" value={gColor} onChange={e=>setGColor(e.target.value)} />
              </Grid>
              <Grid item xs={12} md={2}>
                <TextField fullWidth required type="number" label="Pack Size" value={gPack} onChange={e=>setGPack(e.target.value === '' ? 0 : Number(e.target.value))} />
              </Grid>
              <Grid item xs={12}>
                <Stack direction={{ xs:'column', sm:'row' }} spacing={2} justifyContent={{ xs:'stretch', sm:'flex-end' }} alignItems={{ xs:'stretch', sm:'center' }}>
                  <Button variant="contained" onClick={async()=>{
                    if (!gItemCode.trim() || !gDesc.trim() || !gColor.trim() || !Number.isFinite(Number(gPack)) || Number(gPack) < 0) { toast.error('Complete all fields'); return; }
                    try {
                      await api.post('/items', { itemCode: gItemCode.trim(), itemGroup: selectedGroup, description: gDesc.trim(), color: gColor.trim(), packSize: Number(gPack), enabled: true });
                      toast.success('Item added');
                      setGItemCode(''); setGDesc(''); setGColor(''); setGPack(0);
                      await load();
                    } catch (e:any) {
                      const msg = e?.response?.data?.message || e?.message || 'Failed to add item';
                      toast.error(msg);
                    }
                  }}>Add Item</Button>
                  <Button variant="text" onClick={()=> setSelectedGroup('')}>Close</Button>
                </Stack>
              </Grid>
            </Grid>
            <div style={{ height: 360, width: '100%' }}>
              <DataGrid
                rows={items.filter(it=> (it.itemGroup||'') === selectedGroup).map(it=>({ id: it._id, itemCode: it.itemCode, description: it.description, color: it.color, packSize: (it as any).packSize ?? 0 }))}
                columns={([
                  { field: 'itemCode', headerName: 'Item Code', flex: 1, minWidth: 160 },
                  { field: 'description', headerName: 'Description', flex: 2, minWidth: 240 },
                  { field: 'color', headerName: 'Color', width: 140 },
                  { field: 'packSize', headerName: 'Pack Size', type: 'number', width: 140 },
                  { field: 'actions', headerName: 'Actions', width: 120, sortable: false, filterable: false, renderCell: (p: GridRenderCellParams) => {
                    const r = p.row as any;
                    return (
                      <Button size="small" color="error" variant="outlined" onClick={()=> deleteItem(r.itemCode, selectedGroup)}>Delete</Button>
                    );
                  } }
                ]) as GridColDef[]}
                loading={loading}
                disableRowSelectionOnClick
                density="compact"
                initialState={{ pagination: { paginationModel: { pageSize: 10, page: 0 } } }}
                pageSizeOptions={[10,20,50]}
              />
            </div>
          </Paper>
        )}
        {groupImportResult && (
          <Paper variant="outlined" sx={{ p:2, mt:2, bgcolor:'#fafafa' }}>
            <Typography variant="subtitle2" gutterBottom>Group Import Summary</Typography>
            <Typography variant="body2">Created: <b>{groupImportResult.created ?? 0}</b></Typography>
            <Typography variant="body2">Skipped: <b>{groupImportResult.skipped ?? 0}</b></Typography>
            {Boolean(groupImportResult.errorCount) && (
              <>
                <Typography variant="body2" sx={{ mt: 1 }}>Issues found: <b>{groupImportResult.errorCount}</b></Typography>
                {/* Categorized quick summary */}
                {(() => {
                  const errs = Array.isArray(groupImportResult.errors) ? groupImportResult.errors : [];
                  const dup = errs.filter((er: any) => (er.errors||[]).some((t: string) => String(t).toLowerCase().includes('duplicate')));
                  const locked = errs.filter((er: any) => (er.errors||[]).some((t: string) => String(t).toLowerCase().includes('item count > 0')));
                  const exist = errs.filter((er: any) => (er.errors||[]).some((t: string) => String(t).toLowerCase().includes('already exists')));
                  return (
                    <div style={{ marginTop: 8, marginBottom: 8 }}>
                      <Typography variant="caption" display="block">Summary:</Typography>
                      {!!dup.length && <Typography variant="caption" display="block">• Duplicates in file: <b>{dup.length}</b> (e.g., {dup.slice(0,3).map((d:any)=>d.name).filter(Boolean).join(', ')}{dup.length>3?'…':''})</Typography>}
                      {!!locked.length && <Typography variant="caption" display="block">• Existing groups with Item Count &gt; 0: <b>{locked.length}</b> (cannot update)</Typography>}
                      {!!exist.length && <Typography variant="caption" display="block">• Already existing groups: <b>{exist.length}</b> (skipped)</Typography>}
                    </div>
                  );
                })()}
                <table style={{ width:'100%', borderCollapse:'collapse', marginTop: 8 }}>
                  <thead>
                    <tr>
                      <th align="left">Row</th>
                      <th align="left">Group Name</th>
                      <th align="left">Reason</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(groupImportResult.errors || []).slice(0, 20).map((er: any, i: number) => (
                      <tr key={i} style={{ borderTop:'1px solid #eee' }}>
                        <td>{er.rowNum ?? '-'}</td>
                        <td>{er.name ?? '-'}</td>
                        <td>{Array.isArray(er.errors) ? er.errors.join(', ') : String(er.errors || '')}</td>
                      </tr>
                    ))}
                    {(groupImportResult.errors || []).length > 20 && (
                      <tr><td colSpan={3} style={{ color:'#666', paddingTop: 6 }}>Showing first 20 of {groupImportResult.errors.length} issues…</td></tr>
                    )}
                  </tbody>
                </table>
              </>
            )}
          </Paper>
        )}
      </Paper>
    
    </Container>
  );
}
