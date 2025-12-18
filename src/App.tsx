import React, { useEffect } from 'react';
import { Routes, Route, Navigate, useNavigate, useLocation } from 'react-router-dom';
import Login from './pages/Login';
import Inventory from './pages/Inventory';
import Pallets from './pages/Pallets';
import Orders from './pages/Orders';
import ItemDetail from './pages/ItemDetail';
import NavBar from './components/NavBar';
import ItemRegistry from './pages/ItemRegistry';
import Warehouses from './pages/Warehouses';
import Ship from './pages/Ship';
import Transfer from './pages/Transfer';
import Profile from './pages/Profile';
import OnProcess from './pages/OnProcess';
import Users from './pages/Users';

const RequireAuth = ({ children }: { children: React.ReactElement }) => {
  const token = localStorage.getItem('token');
  const nav = useNavigate();
  useEffect(() => { if (!token) nav('/login'); }, [token, nav]);
  return token ? children : <></>;
};

export default function App() {
  const location = useLocation();
  const showNav = location.pathname !== '/login';
  return (
    <>
    {showNav && <NavBar />}
    <Routes>
      <Route path="/login" element={<Login />} />
      <Route path="/" element={<RequireAuth><Navigate to="/pallets" /></RequireAuth>} />
      <Route path="/inventory" element={<RequireAuth><Inventory /></RequireAuth>} />
      <Route path="/item-registry" element={<RequireAuth><ItemRegistry /></RequireAuth>} />
      <Route path="/items/:itemCode" element={<RequireAuth><ItemDetail /></RequireAuth>} />
      <Route path="/pallets" element={<RequireAuth><Pallets /></RequireAuth>} />
      <Route path="/orders" element={<RequireAuth><Orders /></RequireAuth>} />
      <Route path="/warehouses" element={<RequireAuth><Warehouses /></RequireAuth>} />
      <Route path="/on-process" element={<RequireAuth><OnProcess /></RequireAuth>} />
      <Route path="/ship" element={<RequireAuth><Ship /></RequireAuth>} />
      <Route path="/transfer" element={<RequireAuth><Transfer /></RequireAuth>} />
      <Route path="/users" element={<RequireAuth><Users /></RequireAuth>} />
      <Route path="/profile" element={<RequireAuth><Profile /></RequireAuth>} />
      <Route path="*" element={<Navigate to="/pallets" />} />
    </Routes>
    </>
  );
}
