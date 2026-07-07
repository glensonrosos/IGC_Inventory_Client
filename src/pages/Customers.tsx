import { useEffect, useMemo, useState } from 'react';
import { Container, Typography, Paper, Stack, Button, Dialog, DialogTitle, DialogContent, DialogActions, TextField, Chip, Switch, FormControlLabel } from '@mui/material';
import { DataGrid, GridColDef, GridRenderCellParams } from '@mui/x-data-grid';
import api from '../api';
import { useToast } from '../components/ToastProvider';

export type CustomerRow = {
  _id: string;
  email: string;
  name: string;
  phone?: string;
  accountNumber?: string;
  companyName?: string;
  shippingAddress: string;
  enabled?: boolean;
  createdAt?: string;
};

export default function Customers() {
  const toast = useToast();
  const [loading, setLoading] = useState(false);
  const [rows, setRows] = useState<CustomerRow[]>([]);
  const [qInput, setQInput] = useState('');
  const [qDebounced, setQDebounced] = useState('');

  const [createOpen, setCreateOpen] = useState(false);
  const [newEmail, setNewEmail] = useState('');
  const [newName, setNewName] = useState('');
  const [newPhone, setNewPhone] = useState('');
  const [newAccountNumber, setNewAccountNumber] = useState('');
  const [newCompanyName, setNewCompanyName] = useState('');
  const [newShippingAddress, setNewShippingAddress] = useState('');
  const [newEnabled, setNewEnabled] = useState(true);
  const [createError, setCreateError] = useState('');

  const [editOpen, setEditOpen] = useState(false);
  const [editCustomer, setEditCustomer] = useState<CustomerRow | null>(null);
  const [editEmail, setEditEmail] = useState('');
  const [editName, setEditName] = useState('');
  const [editPhone, setEditPhone] = useState('');
  const [editAccountNumber, setEditAccountNumber] = useState('');
  const [editCompanyName, setEditCompanyName] = useState('');
  const [editShippingAddress, setEditShippingAddress] = useState('');
  const [editError, setEditError] = useState('');

  const load = async () => {
    setLoading(true);
    try {
      const { data } = await api.get<CustomerRow[]>('/customers');
      setRows(Array.isArray(data) ? data : []);
    } catch (e: any) {
      toast.error(e?.response?.data?.message || 'Failed to load customers');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const openCreate = () => {
    setCreateError('');
    setNewEmail('');
    setNewName('');
    setNewPhone('');
    setNewAccountNumber('');
    setNewCompanyName('');
    setNewShippingAddress('');
    setNewEnabled(true);
    setCreateOpen(true);
  };

  const createCustomer = async () => {
    if (!newEmail.trim() || !newName.trim() || !newShippingAddress.trim()) {
      toast.error('Customer Email, Customer Name, and Shipping Address are required');
      return;
    }
    // Uniqueness checks (case-insensitive)
    const emailKey = newEmail.trim().toLowerCase();
    const nameKey = newName.trim().toLowerCase();
    const emailExists = rows.some(r => String(r.email || '').trim().toLowerCase() === emailKey);
    const nameExists = rows.some(r => String(r.name || '').trim().toLowerCase() === nameKey);
    if (emailExists || nameExists) {
      const parts: string[] = [];
      if (emailExists) parts.push('Customer Email');
      if (nameExists) parts.push('Customer Name');
      const msg = `${parts.join(' and ')} already exist${parts.length > 1 ? '' : 's'}.`;
      setCreateError(msg);
      toast.error(msg);
      return;
    }
    try {
      setCreateError('');
      await api.post('/customers', {
        email: newEmail.trim(),
        name: newName.trim(),
        phone: newPhone.trim(),
        accountNumber: newAccountNumber.trim(),
        companyName: newCompanyName.trim(),
        shippingAddress: newShippingAddress.trim(),
        enabled: newEnabled,
      });
      toast.success('Customer created');
      setCreateOpen(false);
      load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to create customer';
      setCreateError(msg);
      toast.error(msg);
    }
  };

  const openEdit = (c: CustomerRow) => {
    setEditCustomer(c);
    setEditEmail(c.email || '');
    setEditName(c.name || '');
    setEditPhone(c.phone || '');
    setEditAccountNumber(c.accountNumber || '');
    setEditCompanyName(c.companyName || '');
    setEditShippingAddress(c.shippingAddress || '');
    setEditError('');
    setEditOpen(true);
  };

  const saveEdit = async () => {
    if (!editCustomer) return;
    if (!editEmail.trim() || !editName.trim() || !editShippingAddress.trim()) {
      toast.error('Customer Email, Customer Name, and Shipping Address are required');
      return;
    }
    try {
      setEditError('');
      await api.patch(`/customers/${encodeURIComponent(editCustomer._id)}`, {
        email: editEmail.trim(),
        name: editName.trim(),
        phone: editPhone.trim(),
        accountNumber: editAccountNumber.trim(),
        companyName: editCompanyName.trim(),
        shippingAddress: editShippingAddress.trim(),
      });
      toast.success('Customer updated');
      setEditOpen(false);
      setEditCustomer(null);
      load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to update customer';
      setEditError(msg);
      toast.error(msg);
    }
  };

  const toggleEnabled = async (c: CustomerRow) => {
    const next = !(c.enabled !== false);
    try {
      await api.patch(`/customers/${encodeURIComponent(c._id)}/status`, { enabled: next });
      toast.success(next ? 'Customer enabled' : 'Customer disabled');
      load();
    } catch (e: any) {
      toast.error(e?.response?.data?.message || 'Failed to update status');
    }
  };

  const columns: GridColDef[] = useMemo(() => {
    return [
      { field: 'email', headerName: 'Customer Email', flex: 1.4, minWidth: 200 },
      { field: 'name', headerName: 'Customer Name', flex: 1.2, minWidth: 180 },
      { field: 'phone', headerName: 'Phone Number', width: 150 },
      { field: 'accountNumber', headerName: 'Customer Number', width: 160 },
      { field: 'companyName', headerName: 'Company Name', width: 180 },
      { field: 'shippingAddress', headerName: 'Shipping Address', flex: 1.6, minWidth: 240 },
      {
        field: 'enabled',
        headerName: 'Status',
        width: 130,
        renderCell: (params: GridRenderCellParams) => {
          const enabled = params.row.enabled !== false;
          return <Chip size="small" color={enabled ? 'success' : 'default'} label={enabled ? 'Enabled' : 'Disabled'} />;
        },
      },
      {
        field: 'createdAt',
        headerName: 'Created',
        width: 180,
        renderCell: (params: GridRenderCellParams) => {
          const v = (params.row as any)?.createdAt;
          return <span>{v ? new Date(v).toLocaleString() : ''}</span>;
        },
      },
      {
        field: 'actions',
        headerName: 'Actions',
        width: 320,
        sortable: false,
        filterable: false,
        renderCell: (params: GridRenderCellParams) => {
          const c: CustomerRow = params.row;
          return (
            <Stack direction="row" spacing={1}>
              <Button size="small" variant="outlined" onClick={() => openEdit(c)}>Edit</Button>
              <Button size="small" variant="outlined" onClick={() => toggleEnabled(c)}>{c.enabled !== false ? 'Disable' : 'Enable'}</Button>
            </Stack>
          );
        },
      },
    ];
  }, []);

  useEffect(() => {
    const t = setTimeout(() => setQDebounced(qInput), 500);
    return () => clearTimeout(t);
  }, [qInput]);

  const filteredRows = useMemo(() => {
    const s = qDebounced.trim().toLowerCase();
    if (!s) return rows;
    return rows.filter((r) => {
      const hay = [
        r.email || '',
        r.name || '',
        r.phone || '',
        r.accountNumber || '',
        r.companyName || '',
        r.shippingAddress || '',
      ].join(' ').toLowerCase();
      return hay.includes(s);
    });
  }, [rows, qDebounced]);

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Customer Management</Typography>
      <Paper sx={{ p: 2, mb: 2 }}>
        <Stack direction={{ xs:'column', sm:'row' }} spacing={1} alignItems={{ xs:'stretch', sm:'center' }}>
          <Button variant="contained" onClick={openCreate}>Add Customer</Button>
          <Button variant="outlined" onClick={load} disabled={loading}>Refresh</Button>
          <TextField
            size="small"
            label="Search"
            placeholder="Email, name, phone, number, company, address"
            value={qInput}
            onChange={(e)=> setQInput(e.target.value)}
            sx={{ ml: { sm: 1 }, width: { xs: '100%', sm: 380 } }}
          />
        </Stack>
      </Paper>
      <Paper sx={{ height: 600, width: '100%', p: 1 }}>
        <DataGrid
          rows={filteredRows.map(r => ({ ...r, id: r._id }))}
          columns={columns}
          loading={loading}
          disableRowSelectionOnClick
          initialState={{ pagination: { paginationModel: { pageSize: 20, page: 0 } } }}
          pageSizeOptions={[10, 20, 50, 100]}
          density="compact"
        />
      </Paper>

      <Dialog open={createOpen} onClose={() => setCreateOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Add Customer</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <TextField label="Customer Email" defaultValue={newEmail} onBlur={(e) => setNewEmail(e.target.value)} required />
            <TextField label="Customer Name" defaultValue={newName} onBlur={(e) => setNewName(e.target.value)} required />
            <TextField label="Phone Number" defaultValue={newPhone} onBlur={(e) => setNewPhone(e.target.value)} />
            <TextField label="Customer Number" defaultValue={newAccountNumber} onBlur={(e) => setNewAccountNumber(e.target.value)} />
            <TextField label="Company Name" defaultValue={newCompanyName} onBlur={(e) => setNewCompanyName(e.target.value)} />
            <TextField label="Shipping Address" defaultValue={newShippingAddress} onBlur={(e) => setNewShippingAddress(e.target.value)} required multiline rows={2} />
            <FormControlLabel control={<Switch checked={newEnabled} onChange={(e)=> setNewEnabled(e.target.checked)} />} label="Enabled" />
            {createError ? (
              <Typography variant="body2" sx={{ color: '#c62828' }}>{createError}</Typography>
            ) : null}
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setCreateOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={createCustomer}>Create</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={editOpen} onClose={() => setEditOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Edit Customer</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <TextField label="Customer Email" defaultValue={editEmail} onBlur={(e) => setEditEmail(e.target.value)} required />
            <TextField label="Customer Name" defaultValue={editName} onBlur={(e) => setEditName(e.target.value)} required />
            <TextField label="Phone Number" defaultValue={editPhone} onBlur={(e) => setEditPhone(e.target.value)} />
            <TextField label="Customer Number" defaultValue={editAccountNumber} onBlur={(e) => setEditAccountNumber(e.target.value)} />
            <TextField label="Company Name" defaultValue={editCompanyName} onBlur={(e) => setEditCompanyName(e.target.value)} />
            <TextField label="Shipping Address" defaultValue={editShippingAddress} onBlur={(e) => setEditShippingAddress(e.target.value)} required multiline rows={2} />
            {editError ? (
              <Typography variant="body2" sx={{ color: '#c62828' }}>{editError}</Typography>
            ) : null}
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setEditOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={saveEdit}>Save</Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
