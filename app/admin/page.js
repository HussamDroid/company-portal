'use client';

import { useState, useEffect } from 'react';
import { supabase } from '../../utils/supabase';

export default function DedicatedAdminPortal() {
  // --- AUTHENTICATION STATES ---
  const [isAdminLoggedIn, setIsAdminLoggedIn] = useState(false);
  const [adminUser, setAdminUser] = useState('');
  const [adminPass, setAdminPass] = useState('');
  const [authError, setAuthError] = useState(null);

  // --- STAFF REGISTRY MANAGEMENT STATES ---
  const [staffRegistry, setStaffRegistry] = useState([]);
  const [regUser, setRegUser] = useState('');
  const [regPass, setRegPass] = useState('');
  const [regRole, setRegRole] = useState('Operator');
  const [provisioning, setProvisioning] = useState(false);

  // --- FETCH REGISTRY UPON SECURE ACCESS ---
  async function fetchStaffRegistry() {
    const { data, error } = await supabase
      .from('user_registry')
      .select('*')
      .order('created_at', { ascending: false });
    if (!error) setStaffRegistry(data || []);
  }

  useEffect(() => {
    if (isAdminLoggedIn) {
      fetchStaffRegistry();
    }
  }, [isAdminLoggedIn]);

  // --- ADMIN AUTH ROUTINE ---
  const handleAdminLogin = (e) => {
    e.preventDefault();
    setAuthError(null);

    if (adminUser.trim() === 'admin' && adminPass.trim() === 'Blackrose2026!') {
      setIsAdminLoggedIn(true);
    } else {
      setAuthError('Wrong Credentials.');
    }
  };

  const handleAdminLogout = () => {
    setIsAdminLoggedIn(false);
    setAdminUser('');
    setAdminPass('');
  };

  // --- STAFF WRITE OPERATION ---
  const handleCreateEmployee = async (e) => {
    e.preventDefault();
    if (!regUser.trim() || !regPass.trim()) {
      alert('Credentials parameters cannot remain completely unmapped.');
      return;
    }

    setProvisioning(true);
    try {
      const { error } = await supabase
        .from('user_registry')
        .insert([{ username: regUser.trim(), password: regPass.trim(), role: regRole }]);

      if (error) {
        if (error.code === '23505') alert('This username identifier is already registered.');
        else throw error;
      } else {
        setRegUser('');
        setRegPass('');
        fetchStaffRegistry();
      }
    } catch (err) {
      alert('Internal Server Error appending registry profile.');
    } finally {
      setProvisioning(false);
    }
  };

  const handleDeleteEmployee = async (id, name) => {
    if (!confirm(`Revoke operational tokens and permanently terminate "${name}"?`)) return;
    const { error } = await supabase.from('user_registry').delete().eq('id', id);
    if (!error) fetchStaffRegistry();
  };

  return (
    <div className="min-h-screen bg-slate-900 font-sans flex flex-col text-slate-100">
      
      {/* ADMIN CONTROL TOWER HEADER */}
      <header className="w-full bg-slate-950 border-b border-slate-800 h-16 flex items-center justify-between px-6 sticky top-0 z-50">
        <div className="flex items-center gap-3">
          
          <div>
            <span className="font-black text-white tracking-tight block text-xs uppercase">Black-Rose</span>
          </div>
        </div>

        {isAdminLoggedIn && (
          <button
            onClick={handleAdminLogout}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 transition-all font-bold text-xs uppercase tracking-wider rounded-xl cursor-pointer"
          >
            Exit Admin Control
          </button>
        )}
      </header>

      {/* CORE CANVAS WORKSPACE */}
      <main className="flex-grow flex items-center justify-center p-6">
        {!isAdminLoggedIn ? (
          /* PRIVATE ACCESS FORMS GATEWAY */
          <div className="w-full max-w-md bg-slate-950 p-8 rounded-2xl shadow-2xl border border-slate-800">
            <div className="text-center mb-6">
              <span className="text-4xl">🛡️</span>
              <h2 className="text-xl font-black uppercase tracking-wider text-white mt-3">Admin Panel</h2>
            </div>

            {authError && (
              <div className="mb-4 p-3 bg-red-950 border border-red-800 text-red-400 rounded-lg text-xs font-bold uppercase text-center tracking-wider">
                🚨 {authError}
              </div>
            )}

            <form onSubmit={handleAdminLogin} className="space-y-4">
              <div>
                <label className="block text-xxs font-black text-slate-500 uppercase tracking-widest mb-1">Username</label>
                <input
                  type="text"
                  required
                  value={adminUser}
                  onChange={(e) => setAdminUser(e.target.value)}
                  placeholder="admin"
                  className="w-full px-4 py-3 border border-slate-800 rounded-xl bg-slate-900 text-sm text-white outline-none focus:border-red-600 transition-all"
                />
              </div>
              <div>
                <label className="block text-xxs font-black text-slate-500 uppercase tracking-widest mb-1">Password</label>
                <input
                  type="password"
                  required
                  value={adminPass}
                  onChange={(e) => setAdminPass(e.target.value)}
                  placeholder="•••••••••••••"
                  className="w-full px-4 py-3 border border-slate-800 rounded-xl bg-slate-900 text-sm text-white outline-none focus:border-red-600 transition-all"
                />
              </div>
              <button
                type="submit"
                className="w-full py-3 bg-red-600 hover:bg-red-700 text-white text-xs font-black uppercase tracking-widest rounded-xl transition-all shadow-md cursor-pointer"
              >
                Login
              </button>
            </form>
          </div>
        ) : (
          /* EXCLUSIVE ACCOUNT PROVISIONING CONSOLE PANELS */
          <div className="w-full max-w-5xl bg-slate-950 border border-slate-800 rounded-3xl p-6 md:p-8 shadow-2xl grid grid-cols-1 md:grid-cols-3 gap-8 animate-fadeIn">
            
            {/* SUB-PANEL A: INGESTION CREATION SLOTS */}
            <div className="space-y-4">
              <div>
                <h2 className="text-md font-black text-white uppercase tracking-wider">New Staff Registration</h2>
                <p className="text-xxs text-slate-400 font-medium">Issue new employees the login credentials.</p>
              </div>

              <form onSubmit={handleCreateEmployee} className="space-y-3 pt-2">
                <div>
                  <label className="block text-xxs font-bold text-slate-400 uppercase tracking-wider mb-1">Assign Username</label>
                  <input
                    type="text"
                    required
                    placeholder=""
                    value={regUser}
                    onChange={(e) => setRegUser(e.target.value)}
                    className="w-full px-3 py-2 border border-slate-800 rounded-lg text-xs bg-slate-900 text-white outline-none focus:border-blue-500"
                  />
                </div>
                <div>
                  <label className="block text-xxs font-bold text-slate-400 uppercase tracking-wider mb-1">Assign Password</label>
                  <input
                    type="text"
                    required
                    placeholder="•••••••••••••"
                    value={regPass}
                    onChange={(e) => setRegPass(e.target.value)}
                    className="w-full px-3 py-2 border border-slate-800 rounded-lg text-xs bg-slate-900 text-white outline-none focus:border-blue-500"
                  />
                </div>
                <div>
                  <label className="block text-xxs font-bold text-slate-400 uppercase tracking-wider mb-1">Assign Access</label>
                  <select
                    value={regRole}
                    onChange={(e) => setRegRole(e.target.value)}
                    className="w-full px-3 py-2 border border-slate-800 rounded-lg text-xs bg-slate-900 text-slate-300 font-bold outline-none focus:border-blue-500"
                  >
                    <option value="Operator">(Floor Operations Dashboard View)</option>
                    <option value="Manager">(Full Catalog Management View)</option>
                  </select>
                </div>
                <button
                  type="submit"
                  disabled={provisioning}
                  className="w-full py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-black text-xs uppercase tracking-wider rounded-lg transition-all disabled:opacity-40 cursor-pointer"
                >
                  {provisioning ? 'Writing Registry...' : 'Authorize and Register Staff Account'}
                </button>
              </form>
            </div>

            {/* SUB-PANEL B: LIVE STAFF MATRIX LISTINGS */}
            <div className="md:col-span-2 border-t md:border-t-0 md:border-l border-slate-800 md:pl-8">
              <h3 className="text-xs font-black uppercase tracking-widest text-slate-400 mb-4">Current Staff Workspace Registry ({staffRegistry.length})</h3>
              
              <div className="border border-slate-800 rounded-xl overflow-hidden bg-slate-900/40">
                <div className="max-h-96 overflow-y-auto divide-y divide-slate-800">
                  {staffRegistry.length === 0 ? (
                    <div className="p-8 text-center text-xs text-slate-500 font-medium">No external staff users logged in database directory.</div>
                  ) : (
                    staffRegistry.map((staff) => (
                      <div key={staff.id} className="p-4 flex items-center justify-between text-xs hover:bg-slate-900/60 transition-colors">
                        <div className="space-y-1">
                          <div className="flex items-center gap-2">
                            <span className="font-bold text-white font-mono text-sm">{staff.username}</span>
                            <span className={`px-2 py-0.5 rounded text-xxs font-black uppercase border ${
                              staff.role === 'Manager' ? 'bg-blue-950/60 text-blue-400 border-blue-900' : 'bg-amber-950/60 text-amber-400 border-amber-900'
                            }`}>{staff.role}</span>
                          </div>
                          <div className="text-slate-500 font-mono text-xxs">Token Code: {staff.password}</div>
                        </div>
                        <button
                          onClick={() => handleDeleteEmployee(staff.id, staff.username)}
                          className="px-3 py-1 text-red-400 hover:text-white hover:bg-red-600/20 border border-red-900 hover:border-red-600 rounded-lg font-bold tracking-wider uppercase text-xxs transition-all cursor-pointer"
                        >
                          Revoke Access
                        </button>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </div>

          </div>
        )}
      </main>
    </div>
  );
}