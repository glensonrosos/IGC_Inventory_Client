import { useEffect, useMemo, useState } from 'react';
import { Container, Typography, Paper, Stack, Button, Dialog, DialogTitle, DialogContent, DialogActions, TextField, MenuItem, Chip } from '@mui/material';
import { DataGrid, GridColDef, GridRenderCellParams } from '@mui/x-data-grid';
import api from '../api';
import { useToast } from '../components/ToastProvider';

type UserRow = {
  _id: string;
  username: string;
  role: 'admin' | 'user' | string;
  enabled?: boolean;
  createdAt?: string;
};

export default function Users() {
  const toast = useToast();
  const [loading, setLoading] = useState(false);
  const [rows, setRows] = useState<UserRow[]>([]);

  const [createOpen, setCreateOpen] = useState(false);
  const [newUsername, setNewUsername] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [newRole, setNewRole] = useState<'admin' | 'user'>('user');
  const [createError, setCreateError] = useState('');

  const [resetOpen, setResetOpen] = useState(false);
  const [resetUser, setResetUser] = useState<UserRow | null>(null);
  const [resetPassword, setResetPassword] = useState('');
  const [resetError, setResetError] = useState('');

  const [roleOpen, setRoleOpen] = useState(false);
  const [roleUser, setRoleUser] = useState<UserRow | null>(null);
  const [roleValue, setRoleValue] = useState<'admin' | 'user'>('user');
  const [roleError, setRoleError] = useState('');

  const load = async () => {
    setLoading(true);
    try {
      const { data } = await api.get<UserRow[]>('/users');
      setRows(Array.isArray(data) ? data : []);
    } catch (e: any) {
      toast.error(e?.response?.data?.message || 'Failed to load users');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const createUser = async () => {
    if (!newUsername.trim() || !newPassword.trim()) {
      toast.error('username and password required');
      return;
    }
    try {
      setCreateError('');
      await api.post('/users', { username: newUsername.trim(), password: newPassword, role: newRole });
      toast.success('User created');
      setCreateOpen(false);
      setNewUsername('');
      setNewPassword('');
      setNewRole('user');
      load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to create user';
      setCreateError(msg);
      toast.error(msg);
    }
  };

  const openRole = (u: UserRow) => {
    setRoleUser(u);
    setRoleValue((u.role === 'admin' ? 'admin' : 'user') as any);
    setRoleError('');
    setRoleOpen(true);
  };

  const saveRole = async () => {
    if (!roleUser) return;
    try {
      setRoleError('');
      await api.patch(`/users/${encodeURIComponent(roleUser._id)}/role`, { role: roleValue });
      toast.success('Role updated');
      setRoleOpen(false);
      setRoleUser(null);
      load();
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to update role';
      setRoleError(msg);
      toast.error(msg);
    }
  };

  const toggleEnabled = async (u: UserRow) => {
    const next = !(u.enabled !== false);
    try {
      await api.patch(`/users/${encodeURIComponent(u._id)}/status`, { enabled: next });
      toast.success(next ? 'User enabled' : 'User disabled');
      load();
    } catch (e: any) {
      toast.error(e?.response?.data?.message || 'Failed to update status');
    }
  };

  const openReset = (u: UserRow) => {
    setResetUser(u);
    setResetPassword('');
    setResetError('');
    setResetOpen(true);
  };

  const resetPwd = async () => {
    if (!resetUser) return;
    if (!resetPassword.trim()) {
      toast.error('new password required');
      return;
    }
    try {
      setResetError('');
      await api.post(`/users/${encodeURIComponent(resetUser._id)}/reset-password`, { newPassword: resetPassword });
      toast.success('Password reset');
      setResetOpen(false);
      setResetUser(null);
      setResetPassword('');
    } catch (e: any) {
      const msg = e?.response?.data?.message || 'Failed to reset password';
      setResetError(msg);
      toast.error(msg);
    }
  };

  const columns: GridColDef[] = useMemo(() => {
    const cols: GridColDef[] = [
      { field: 'username', headerName: 'Username', flex: 1.2, minWidth: 180 },
      { field: 'role', headerName: 'Role', width: 120 },
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
        width: 360,
        sortable: false,
        filterable: false,
        renderCell: (params: GridRenderCellParams) => {
          const u: UserRow = params.row;
          return (
            <Stack direction="row" spacing={1}>
              <Button size="small" variant="outlined" onClick={() => openRole(u)}>Change Role</Button>
              <Button size="small" variant="outlined" onClick={() => openReset(u)}>Reset Password</Button>
              <Button size="small" variant="outlined" onClick={() => toggleEnabled(u)}>{u.enabled !== false ? 'Disable' : 'Enable'}</Button>
            </Stack>
          );
        },
      },
    ];
    return cols;
  }, []);

  return (
    <Container sx={{ mt: 4 }}>
      <Typography variant="h4" gutterBottom>Users Management</Typography>
      <Paper sx={{ p: 2, mb: 2 }}>
        <Stack direction="row" spacing={1}>
          <Button variant="contained" onClick={() => setCreateOpen(true)}>Add User</Button>
          <Button variant="outlined" onClick={load} disabled={loading}>Refresh</Button>
        </Stack>
      </Paper>
      <Paper sx={{ height: 600, width: '100%', p: 1 }}>
        <DataGrid
          rows={rows.map(r => ({ ...r, id: r._id }))}
          columns={columns}
          loading={loading}
          disableRowSelectionOnClick
          initialState={{ pagination: { paginationModel: { pageSize: 20, page: 0 } } }}
          pageSizeOptions={[10, 20, 50, 100]}
          density="compact"
        />
      </Paper>

      <Dialog open={createOpen} onClose={() => setCreateOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Add User</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <TextField label="Username" value={newUsername} onChange={(e) => setNewUsername(e.target.value)} />
            <TextField label="Password" type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} />
            <TextField select label="Role" value={newRole} onChange={(e) => setNewRole(e.target.value as any)}>
              <MenuItem value="admin">admin</MenuItem>
              <MenuItem value="user">user</MenuItem>
            </TextField>
            {createError ? (
              <Typography variant="body2" sx={{ color: '#c62828' }}>{createError}</Typography>
            ) : null}
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setCreateOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={createUser}>Create</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={resetOpen} onClose={() => setResetOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Reset Password</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <TextField label="Username" value={resetUser?.username || ''} disabled />
            <TextField label="New Password" type="password" value={resetPassword} onChange={(e) => setResetPassword(e.target.value)} />
            {resetError ? (
              <Typography variant="body2" sx={{ color: '#c62828' }}>{resetError}</Typography>
            ) : null}
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setResetOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={resetPwd}>Reset</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={roleOpen} onClose={() => setRoleOpen(false)} fullWidth maxWidth="sm">
        <DialogTitle>Change Role</DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <TextField label="Username" value={roleUser?.username || ''} disabled />
            <TextField select label="Role" value={roleValue} onChange={(e) => setRoleValue(e.target.value as any)}>
              <MenuItem value="admin">admin</MenuItem>
              <MenuItem value="user">user</MenuItem>
            </TextField>
            {roleError ? (
              <Typography variant="body2" sx={{ color: '#c62828' }}>{roleError}</Typography>
            ) : null}
          </Stack>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setRoleOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={saveRole}>Save</Button>
        </DialogActions>
      </Dialog>
    </Container>
  );
}
