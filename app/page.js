'use client';

import { useState, useEffect } from 'react';
import { supabase } from '../utils/supabase';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

// THEME NOTE: UI accent updated to exact BlackRose white + maroon palette.
// Primary accent uses rgba(138, 21, 56, 0.85) with white surfaces.
// Semantic status colors are still used where helpful.

export default function IntegratedOperationsPortal() {
  // --- AUTHENTICATION STATES ---
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [authRole, setAuthRole] = useState(''); 
  const [loginUser, setLoginUser] = useState('');
  const [loginPass, setLoginPass] = useState('');
  const [authError, setAuthError] = useState(null);

  // --- ACCESS CONTROL RBAC REGISTRY STATE ---
  const [customRoles, setCustomRoles] = useState([
    {
      roleName: 'Photographer',
      permissions: { canUploadAssets: true, canModifyDataSheets: false, canReviewArrays: false, canSuperviseStaff: false }
    },
    {
      roleName: 'Content Editor',
      permissions: { canUploadAssets: false, canModifyDataSheets: true, canReviewArrays: false, canSuperviseStaff: false }
    }
  ]);
  const [showRoleModal, setShowRoleModal] = useState(false);
  const [newRoleForm, setNewRoleForm] = useState({
    roleName: '',
    canUploadAssets: false,
    canModifyDataSheets: false,
    canReviewArrays: false,
    canSuperviseStaff: false
  });

  // --- INTEGRATED NEW STAFF REGISTRATION FORM STATES ---
  const [regUsername, setRegUsername] = useState('');
  const [regPassword, setRegPassword] = useState('');
  const [regRole, setRegRole] = useState('Operator');

  // --- DATA FLOW STATES ---
  const [products, setProducts] = useState([]);
  const [manifestHistory, setManifestHistory] = useState([]);
  const [userRegistry, setUserRegistry] = useState([]);
  const [stores, setStores] = useState([]);
  const [selectedStoreId, setSelectedStoreId] = useState('ALL');
  const [showStoreCreate, setShowStoreCreate] = useState(false);
  const [newStoreName, setNewStoreName] = useState('');
  const [storeImageUploadingId, setStoreImageUploadingId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [imageUploadingProductId, setImageUploadingProductId] = useState(null);
  const [realtimeStatus, setRealtimeStatus] = useState('Connecting...');

  // --- TASK BOARD STATES ---
  const [tasks, setTasks] = useState([]);
  const [taskSaving, setTaskSaving] = useState(false);
  const [editingTaskId, setEditingTaskId] = useState(null);
  const [taskError, setTaskError] = useState(null);
  const [taskForm, setTaskForm] = useState({
    title: '',
    description: '',
    priority: 'Normal',
    assigned_role: 'All',
    due_at: ''
  });

  // --- UTILITY UI STATES ---
  const [fullViewImage, setFullViewImage] = useState(null);
  const [dashboardClock, setDashboardClock] = useState(new Date());

  // --- APPLICATION NAVIGATION SYSTEM STATE ---
  const [activeTab, setActiveTab] = useState('home');
  const [isSidePanelCollapsed, setIsSidePanelCollapsed] = useState(false);
  const [isStoresPanelOpen, setIsStoresPanelOpen] = useState(false);
  const [selectedHistoryScope, setSelectedHistoryScope] = useState(null); 
  const [selectedOperatorStats, setSelectedOperatorStats] = useState(null); 

  // --- AD-HOC TOOL STATES ---
  const [showAdHocModal, setShowAdHocModal] = useState(false);
  const [adHocForm, setAdHocForm] = useState({
    sku: '',
    product_name: '',
    category: 'General',
    warehouse: 'Black Rose Trading',
    stock_quantity: 0
  });

  // --- FILTER & MODIFICATION STATES ---
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('All');
  const [assetFolderSearchQuery, setAssetFolderSearchQuery] = useState('');
  const [assetFolderStoreId, setAssetFolderStoreId] = useState('ALL');
  const [editingId, setEditingId] = useState(null);
  
  // --- MODAL WORKSPACE STATES ---
  const [selectedProduct, setSelectedProduct] = useState(null); 
  const [managerPreview, setManagerPreview] = useState(null);   
  const [isRejecting, setIsRejecting] = useState(false);        
  const [rejectNote, setRejectNote] = useState('');             

  const [editForm, setEditForm] = useState({
    sku: '',
    product_name: '',
    category: '',
    warehouse: '',
    stock_quantity: 0,
    platforms: '',
    status: ''
  });

  // --- SYSTEM PERMISSION RESOLUTION UTILITY ---
  const checkPermission = (action) => {
    if (authRole === 'Admin' || authRole === 'Manager') return true;
    if (authRole === 'Operator' && action === 'view_workspace') return true;
    
    const targetRole = customRoles.find(r => r.roleName.toLowerCase() === authRole.toLowerCase());
    if (!targetRole) return false;
    
    if (action === 'upload_assets') return targetRole.permissions.canUploadAssets;
    if (action === 'modify_sheets') return targetRole.permissions.canModifyDataSheets;
    if (action === 'review_arrays') return targetRole.permissions.canReviewArrays;
    if (action === 'supervise_staff') return targetRole.permissions.canSuperviseStaff;
    return false;
  };

  // --- SESSION PERSISTENCE ---
  useEffect(() => {
    const savedUser = localStorage.getItem('blackrose_user');
    const savedRole = localStorage.getItem('blackrose_role');
    const savedCustomRoles = localStorage.getItem('blackrose_custom_roles');
    
    if (savedCustomRoles) {
      setCustomRoles(JSON.parse(savedCustomRoles));
    }
    if (savedUser && savedRole) {
      setLoginUser(savedUser);
      setAuthRole(savedRole);
      setIsLoggedIn(true);
      setActiveTab('home');
    }
  }, []);

  // --- SUBTLE DASHBOARD DATE/TIME ---
  useEffect(() => {
    const timer = setInterval(() => setDashboardClock(new Date()), 30000);
    return () => clearInterval(timer);
  }, []);

  // --- DATABASE DATA SYNCHRONIZERS ---
  async function fetchProducts() {
    const [prodsRes, historyRes, usersRes, tasksRes, storesRes] = await Promise.all([
      supabase
        .from('products')
        .select('*')
        .order('created_at', { ascending: false }),
      supabase
        .from('manifest_history')
        .select('*')
        .order('created_at', { ascending: false }),
      supabase
        .from('user_registry')
        .select('id, username, role, created_at'),
      supabase
        .from('task_board')
        .select('*')
        .order('created_at', { ascending: false }),
      supabase
        .from('stores')
        .select('*')
        .order('name', { ascending: true })
    ]);

    if (!prodsRes.error) {
      setProducts(prodsRes.data || []);
      
      if (selectedProduct) {
        const updatedModalTarget = prodsRes.data.find(p => p.id === selectedProduct.id);
        if (updatedModalTarget) setSelectedProduct(updatedModalTarget);
      }
      if (managerPreview) {
        const updatedPreviewTarget = prodsRes.data.find(p => p.id === managerPreview.id);
        if (updatedPreviewTarget) setManagerPreview(updatedPreviewTarget);
      }
    }

    if (!historyRes.error && historyRes.data) {
      setManifestHistory(mergeRemoteAndLocalManifestHistory(historyRes.data));
    } else {
      console.warn('manifest_history fetch failed. Using local Excel archive backup only:', historyRes.error);
      setManifestHistory(mergeRemoteAndLocalManifestHistory([]));
    }

    if (!usersRes.error) {
      setUserRegistry(usersRes.data || []);
    }

    if (!tasksRes.error) {
      setTasks(tasksRes.data || []);
      setTaskError(null);
    } else {
      console.warn('task_board fetch failed. Run the task_board SQL setup if this is the first install:', tasksRes.error);
      setTaskError(tasksRes.error.message || 'Task board table is not ready.');
      setTasks([]);
    }

    if (!storesRes.error) {
      setStores(storesRes.data || []);
    } else {
      console.warn('stores fetch failed. Run the stores SQL/policies if this is the first install:', storesRes.error);
      setStores([]);
    }

    setLoading(false);
  }

  const getArray = (val) => {
    if (!val) return [];
    if (Array.isArray(val)) return val;
    return [val];
  };

  const normalizeStoreId = (value) => {
    if (value === 'ALL' || value === undefined || value === null || value === '') return 'ALL';
    return Number(value);
  };

  const getSelectedStore = () => {
    if (selectedStoreId === 'ALL') return null;
    return stores.find(store => Number(store.id) === Number(selectedStoreId)) || null;
  };

  const getStoreNameById = (storeId) => {
    if (!storeId) return 'Unassigned Store';
    return stores.find(store => Number(store.id) === Number(storeId))?.name || `Store #${storeId}`;
  };

  const isStoreScoped = selectedStoreId !== 'ALL';
  const selectedStore = getSelectedStore();

  const isProductInSelectedStore = (product) => {
    if (!isStoreScoped) return true;
    return Number(product.store_id || 0) === Number(selectedStoreId);
  };

  const handleSelectStore = (storeId, targetTab = 'dashboard') => {
    const normalizedStoreId = normalizeStoreId(storeId);
    setSelectedStoreId(normalizedStoreId);
    setAssetFolderStoreId(normalizedStoreId);
    setSelectedHistoryScope(null);
    setSearchQuery('');
    setStatusFilter('All');
    setActiveTab(targetTab);
  };

  const handleCreateStore = async (e) => {
    e.preventDefault();
    const cleanName = newStoreName.trim();
    if (!cleanName) return;

    try {
      const { data, error } = await supabase
        .from('stores')
        .insert([{ name: cleanName }])
        .select('*')
        .single();

      if (error) throw error;

      setStores(prev => {
        const merged = [data, ...prev.filter(store => Number(store.id) !== Number(data.id))];
        return merged.sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')));
      });
      setSelectedStoreId(data.id);
      setNewStoreName('');
      setShowStoreCreate(false);
      setActiveTab('home');
    } catch (err) {
      alert('Store creation failed: ' + (err.message || 'Unknown error'));
    }
  };

  const handleStoreImageUpload = async (storeId, e) => {
    const file = e.target.files?.[0];
    if (!file || !storeId) return;

    if (!file.type?.startsWith('image/')) {
      alert('Please upload a valid image file for the store card.');
      e.target.value = null;
      return;
    }

    setStoreImageUploadingId(storeId);

    try {
      const fileExt = file.name.split('.').pop() || 'jpg';
      const safeStoreId = String(storeId).replace(/[^a-zA-Z0-9_-]/g, '_');
      const storagePath = `stores/${safeStoreId}-${Date.now()}.${fileExt}`;

      const { error: uploadError } = await supabase.storage
        .from('product-assets')
        .upload(storagePath, file, { upsert: true });

      if (uploadError) throw uploadError;

      const { data: urlData } = supabase
        .storage
        .from('product-assets')
        .getPublicUrl(storagePath);

      const imageUrl = urlData?.publicUrl;
      if (!imageUrl) throw new Error('Store image URL could not be generated.');

      const { data: updatedStore, error: updateError } = await supabase
        .from('stores')
        .update({ image_url: imageUrl })
        .eq('id', storeId)
        .select('*')
        .single();

      if (updateError) throw updateError;

      setStores(prev =>
        prev
          .map(store => Number(store.id) === Number(storeId) ? updatedStore : store)
          .sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')))
      );
    } catch (err) {
      alert('Store image upload failed: ' + (err.message || 'Unknown error'));
    } finally {
      setStoreImageUploadingId(null);
      e.target.value = null;
    }
  };

  const handleDeleteStore = async (storeId, storeName) => {
    if (!(authRole === 'Admin' || authRole === 'Manager')) return;
    if (!storeId) return;

    const linkedProducts = products.filter(p => Number(p.store_id || 0) === Number(storeId));
    const linkedHistory = manifestHistory.filter(h => Number(h.store_id || 0) === Number(storeId));

    const warningMessage = linkedProducts.length > 0 || linkedHistory.length > 0
      ? `Delete store "${storeName}"?\n\nThis store has ${linkedProducts.length} linked product(s) and ${linkedHistory.length} uploaded sheet archive(s).\n\nDeleting the store will also remove those store products and sheet history from the portal. This action cannot be undone.`
      : `Delete store "${storeName}"?\n\nThis action cannot be undone.`;

    if (!window.confirm(warningMessage)) return;

    setUploading(true);
    try {
      const sheetNamesToClear = Array.from(new Set(linkedHistory.map(h => h.filename).filter(Boolean)));

      const { error: productsDeleteError } = await supabase
        .from('products')
        .delete()
        .eq('store_id', storeId);

      if (productsDeleteError) throw productsDeleteError;

      const { error: historyDeleteError } = await supabase
        .from('manifest_history')
        .delete()
        .eq('store_id', storeId);

      if (historyDeleteError) throw historyDeleteError;

      const { error: storeDeleteError } = await supabase
        .from('stores')
        .delete()
        .eq('id', storeId);

      if (storeDeleteError) throw storeDeleteError;

      if (typeof window !== 'undefined') {
        sheetNamesToClear.forEach(filename => {
          localStorage.removeItem(getLocalManifestKey(filename, storeId));
        });
      }

      setStores(prev => prev.filter(store => Number(store.id) !== Number(storeId)));
      setProducts(prev => prev.filter(product => Number(product.store_id || 0) !== Number(storeId)));
      setManifestHistory(prev => prev.filter(history => Number(history.store_id || 0) !== Number(storeId)));

      if (Number(selectedStoreId) === Number(storeId)) {
        setSelectedStoreId('ALL');
        setAssetFolderStoreId('ALL');
        setSelectedHistoryScope(null);
        setActiveTab('home');
      }

      alert(`Store "${storeName}" was deleted successfully.`);
    } catch (err) {
      alert('Store deletion failed: ' + (err.message || 'Unknown error'));
    } finally {
      setUploading(false);
    }
  };

  const applyProductPatchLocally = (productId, patch) => {
    if (!productId || !patch) return;

    setProducts(prev => {
      const exists = prev.some(p => p.id === productId);
      if (!exists) return [patch, ...prev];
      return prev.map(p => p.id === productId ? { ...p, ...patch } : p);
    });

    setSelectedProduct(prev => prev && prev.id === productId ? { ...prev, ...patch } : prev);
    setManagerPreview(prev => prev && prev.id === productId ? { ...prev, ...patch } : prev);
  };


  // --- EXCEL ARCHIVE + STATUS HELPERS ---
  // Keep these columns separate from OpenCart's own "Status (Enabled=1/Disabled=0)" column.
  const IMAGE_STATUS_COLUMN = 'Image_Workflow_Status';
  const SYSTEM_EXPORT_COLUMNS = [
    IMAGE_STATUS_COLUMN,
    'System_Processed_By',
    'System_Raw_Assets',
    'System_Edited_Assets'
  ];

  // Local backup is used only as a safety net. The main archive should still be saved in Supabase manifest_history.
  const LOCAL_MANIFEST_PREFIX = 'blackrose_manifest_archive_v2_';

  const getLocalManifestKey = (filename, storeId = null) => `${LOCAL_MANIFEST_PREFIX}${encodeURIComponent(storeId || 'general')}__${encodeURIComponent(filename)}`;

  const saveManifestArchiveLocally = (record) => {
    if (typeof window === 'undefined' || !record?.filename) return;

    try {
      localStorage.setItem(
        getLocalManifestKey(record.filename, record.store_id),
        JSON.stringify({
          ...record,
          id: `local-${record.store_id || 'general'}-${record.filename}`,
          local_only: true
        })
      );
    } catch (err) {
      console.warn('Local manifest backup failed:', err);
    }
  };

  const loadLocalManifestArchiveByName = (filename, storeId = selectedStoreId !== 'ALL' ? selectedStoreId : null) => {
    if (typeof window === 'undefined' || !filename) return null;

    try {
      const exactSaved = localStorage.getItem(getLocalManifestKey(filename, storeId));
      if (exactSaved) return JSON.parse(exactSaved);

      // Backward compatibility for old local archives saved before stores existed.
      const legacySaved = localStorage.getItem(`${LOCAL_MANIFEST_PREFIX}${encodeURIComponent(filename)}`);
      return legacySaved ? JSON.parse(legacySaved) : null;
    } catch (err) {
      console.warn('Local manifest read failed:', err);
      return null;
    }
  };

  const loadAllLocalManifestArchives = () => {
    if (typeof window === 'undefined') return [];

    try {
      return Object.keys(localStorage)
        .filter(key => key.startsWith(LOCAL_MANIFEST_PREFIX))
        .map(key => JSON.parse(localStorage.getItem(key)))
        .filter(record => record?.filename && record?.raw_payload);
    } catch (err) {
      console.warn('Local manifest list failed:', err);
      return [];
    }
  };

  const mergeRemoteAndLocalManifestHistory = (remoteRows = []) => {
    const merged = new Map();

    loadAllLocalManifestArchives().forEach(record => {
      merged.set(`${record.store_id || 'general'}__${record.filename}`, record);
    });

    remoteRows.forEach(record => {
      if (record?.filename) {
        merged.set(`${record.store_id || 'general'}__${record.filename}`, record);
      }
    });

    return Array.from(merged.values()).sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));
  };

  const normalizeHeaderForSearch = (value) =>
    String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  const buildSheetArchive = (worksheet) => {
    // header: 1 = array-of-arrays. defval:'' preserves blank cells.
    // This is what prevents blank columns from being stripped.
    const matrix = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
      raw: false,
      blankrows: false
    });

    if (!matrix || matrix.length === 0) {
      return { headers: [], uniqueHeaders: [], rows: [], objects: [] };
    }

    const headers = (matrix[0] || []).map((header, index) => {
      const cleanHeader = String(header || '').trim();
      return cleanHeader || `Column ${index + 1}`;
    });

    // Objects cannot have duplicate keys. Excel can. So use unique keys only for processing.
    // The original headers are still saved separately for export.
    const seen = {};
    const uniqueHeaders = headers.map((header, index) => {
      const base = header || `Column ${index + 1}`;
      seen[base] = (seen[base] || 0) + 1;
      return seen[base] === 1 ? base : `${base}__${seen[base]}`;
    });

    const rows = matrix
      .slice(1)
      .filter(row => row.some(cell => String(cell || '').trim() !== ''))
      .map(row => headers.map((_, index) => row[index] ?? ''));

    const objects = rows.map(row => {
      const obj = {};
      uniqueHeaders.forEach((header, index) => {
        obj[header] = row[index] ?? '';
      });
      return obj;
    });

    return { headers, uniqueHeaders, rows, objects };
  };

  const archivePayloadToMatrix = (payload) => {
    if (!payload) return null;

    // New format: shape-preserving array archive.
    if (payload.type === 'aoa-v1' && Array.isArray(payload.headers) && Array.isArray(payload.rows)) {
      return [
        payload.headers,
        ...payload.rows.map(row => payload.headers.map((_, index) => row[index] ?? ''))
      ];
    }

    // Backward compatibility for old history records saved as array-of-objects.
    if (Array.isArray(payload)) {
      const headers = [];
      payload.forEach(row => {
        Object.keys(row || {}).forEach(key => {
          if (!headers.includes(key)) headers.push(key);
        });
      });
      return [headers, ...payload.map(row => headers.map(header => row?.[header] ?? ''))];
    }

    return null;
  };

  const findColumnIndexByKeywords = (headers, keywords) => {
    return headers.findIndex(header => {
      const cleanHeader = normalizeHeaderForSearch(header);
      return keywords.some(keyword => cleanHeader.includes(keyword));
    });
  };

  const addLiveStatusColumnsToMatrix = (matrix, filename) => {
    if (!matrix || matrix.length === 0) return [];

    const scopedStoreId = selectedStoreId !== 'ALL' ? Number(selectedStoreId) : null;
    const headers = matrix[0] || [];
    const skuIndex = findColumnIndexByKeywords(headers, ['sku', 'itemcode', 'barcode']);

    return [
      [...headers, ...SYSTEM_EXPORT_COLUMNS],
      ...matrix.slice(1).map(row => {
        const skuVal = skuIndex >= 0 ? String(row[skuIndex] || '').trim() : '';
        const activeMatch = products.find(
          p => String(p.sku).trim() === skuVal &&
            p.sheet_reference === filename &&
            (!scopedStoreId || Number(p.store_id || 0) === scopedStoreId)
        );

        return [
          ...headers.map((_, index) => row[index] ?? ''),
          activeMatch ? activeMatch.status : 'Missing',
          activeMatch ? (activeMatch.processed_by || 'Unassigned') : '',
          activeMatch ? getArray(activeMatch.raw_image_url).length : 0,
          activeMatch ? getArray(activeMatch.edited_image_url).length : 0
        ];
      })
    ];
  };

  const calculateImageWorkflowStatus = (rawImages, editedImages, currentStatus = 'Missing') => {
    if (currentStatus === 'Rejected') return 'Rejected';
    if (getArray(editedImages).length > 0) return 'Completed';
    if (getArray(rawImages).length > 0) return 'Processing';
    return 'Missing';
  };

  // --- WEBSOCKET REAL-TIME SYNC ---
  useEffect(() => {
    if (!isLoggedIn) return;

    fetchProducts();

    const productsChannel = supabase
      .channel('live-portal-stream')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'products' }, (payload) => {
        if (payload.eventType === 'DELETE') {
          const deletedId = payload.old?.id;
          if (deletedId) {
            setProducts(prev => prev.filter(p => p.id !== deletedId));
            setSelectedProduct(prev => prev && prev.id === deletedId ? null : prev);
            setManagerPreview(prev => prev && prev.id === deletedId ? null : prev);
          }
          return;
        }

        if (payload.new?.id) {
          applyProductPatchLocally(payload.new.id, payload.new);
        }
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'manifest_history' }, () => {
        fetchProducts();
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'task_board' }, () => {
        fetchProducts();
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'stores' }, () => {
        fetchProducts();
      })
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') setRealtimeStatus('Live Connected');
        else setRealtimeStatus('Disconnected');
      });

    return () => {
      supabase.removeChannel(productsChannel);
    };
  }, [isLoggedIn]);

  // --- LOGIN ROUTINES ---
  const handlePortalLogin = async (e) => {
    e.preventDefault();
    setAuthError(null);

    const cleanUser = loginUser.trim();
    const cleanPass = loginPass.trim();

    if (cleanUser === 'admin' && cleanPass === 'admin123') {
      setAuthRole('Admin');
      setIsLoggedIn(true);
      setActiveTab('home');
      localStorage.setItem('blackrose_user', 'System Administrator');
      localStorage.setItem('blackrose_role', 'Admin');
      return;
    }

    try {
      const { data, error } = await supabase
        .from('user_registry')
        .select('*')
        .eq('username', cleanUser)
        .eq('password', cleanPass)
        .single();

      if (error || !data) {
        setAuthError('Invalid system username or access code.');
      } else {
        setAuthRole(data.role);
        setIsLoggedIn(true);
        localStorage.setItem('blackrose_user', data.username);
        localStorage.setItem('blackrose_role', data.role);
        setActiveTab('home');
      }
    } catch (err) {
      setAuthError('Network communication timeout.');
    }
  };

  const handlePortalLogout = () => {
    setIsLoggedIn(false);
    setAuthRole('');
    setLoginUser('');
    setLoginPass('');
    setEditingId(null);
    setSelectedProduct(null);
    setSelectedHistoryScope(null);
    setSelectedOperatorStats(null);
    closeManagerPreview();
    setSelectedStoreId('ALL');
    setActiveTab('home');
    localStorage.removeItem('blackrose_user');
    localStorage.removeItem('blackrose_role');
  };

  // --- SHEET PROCESSING INGESTION ENGINE ---
  const handleExcelUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setUploading(true);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const arrayBuffer = evt.target.result;
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const worksheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[worksheetName];

        const archive = buildSheetArchive(worksheet);
        const rawRows = archive.objects;

        if (rawRows.length === 0) {
          alert("⚠️ Upload Aborted: The selected file does not contain any readable data rows.");
          setUploading(false);
          return;
        }

        const extractFieldByKeywords = (row, keywords) => {
          const rowKeys = Object.keys(row);
          const foundKey = rowKeys.find(key => {
            const cleanKey = normalizeHeaderForSearch(key);
            return keywords.some(keyword => cleanKey.includes(keyword));
          });
          return foundKey ? row[foundKey] : undefined;
        };

        const currentTimestamp = new Date().toISOString();
        const activeStoreIdForUpload = selectedStoreId !== 'ALL' ? Number(selectedStoreId) : null;
        const activeStoreNameForUpload = activeStoreIdForUpload ? getStoreNameById(activeStoreIdForUpload) : null;

        if ((authRole === 'Admin' || authRole === 'Manager') && !activeStoreIdForUpload) {
          alert('Select a store from the Stores dropdown before uploading a product Excel sheet.');
          setUploading(false);
          return;
        }

        const isContentEditorSync = authRole === 'Content Editor' || (authRole !== 'Admin' && authRole !== 'Manager' && checkPermission('modify_sheets'));

        if (isContentEditorSync) {
          if (products.length === 0) {
            alert(
              "⚠️ Sheet Sync Failed!\n\n" +
              "Reason: You are using an Editor profile, which updates existing catalog data. Because the database was cleared, there are no products to match.\n\n" +
              "Fix Action: Log out and re-login as an Admin or Manager to do a completely fresh structural sheet import first!"
            );
            setUploading(false);
            return;
          }

          let fieldsPatchedCount = 0;
          for (const row of rawRows) {
            const rowSku = extractFieldByKeywords(row, ['sku', 'itemcode', 'barcode']);
            if (!rowSku) continue;

            const cleanSkuStr = String(rowSku).trim();
            const existingMatch = products.find(p => String(p.sku).trim() === cleanSkuStr && (!activeStoreIdForUpload || Number(p.store_id || 0) === Number(activeStoreIdForUpload)));

            if (existingMatch) {
              const patchedPayload = {};
              
              const rowName = extractFieldByKeywords(row, ['productname', 'name', 'itemname', 'nomenclature', 'title']);
              const rowCategory = extractFieldByKeywords(row, ['category', 'type', 'group', 'department']);
              const rowWarehouse = extractFieldByKeywords(row, ['warehouse', 'location', 'store', 'hub']);
              const rowStock = extractFieldByKeywords(row, ['stock', 'quantity', 'qty', 'count', 'available']);

              if (rowName && !existingMatch.product_name) patchedPayload.product_name = String(rowName).trim();
              if (rowCategory && (!existingMatch.category || existingMatch.category === 'General')) patchedPayload.category = String(rowCategory).trim();
              if (rowWarehouse && (!existingMatch.warehouse || existingMatch.warehouse === 'Black Rose Trading')) patchedPayload.warehouse = String(rowWarehouse).trim();
              if (rowStock !== undefined) patchedPayload.stock_quantity = parseInt(String(rowStock).replace(/[^0-9]/g, ''), 10) || 0;

              if (Object.keys(patchedPayload).length > 0) {
                await supabase.from('products').update(patchedPayload).eq('id', existingMatch.id);
                fieldsPatchedCount++;
              }
            }
          }
          alert(`🎉 Sync Complete! Patched structural details across ${fieldsPatchedCount} matching rows smoothly.`);
          fetchProducts();
          return;
        }

        const sanitizedRaw = rawRows.map(row => {
          const matchedSku = extractFieldByKeywords(row, ['sku', 'itemcode', 'barcode']);
          const matchedName = extractFieldByKeywords(row, ['productname', 'name', 'itemname', 'nomenclature', 'title']);
          const matchedCategory = extractFieldByKeywords(row, ['category', 'type', 'group', 'department']);
          const matchedWarehouse = extractFieldByKeywords(row, ['warehouse', 'location', 'store', 'hub']);
          const matchedPlatforms = extractFieldByKeywords(row, ['platform', 'outlet', 'channel', 'ecommerce']);
          const matchedStock = extractFieldByKeywords(row, ['stock', 'quantity', 'qty', 'count', 'available']);

          return {
            sku: matchedSku ? String(matchedSku).trim() : '',
            product_name: matchedName ? String(matchedName).trim() : '',
            category: matchedCategory ? String(matchedCategory).trim() : 'General',
            warehouse: matchedWarehouse ? String(matchedWarehouse).trim() : 'Black Rose Trading',
            platforms: matchedPlatforms ? String(matchedPlatforms).split(',').map(p => p.trim()) : ['E-commerce'],
            stock_quantity: matchedStock ? parseInt(String(matchedStock).replace(/[^0-9]/g, ''), 10) || 0 : 0,
            // Excel does not need to contain this. The system creates it automatically.
            status: 'Missing',
            raw_image_url: [],
            edited_image_url: [],
            processed_by: null,
            rejection_note: null,
            sheet_reference: file.name,
            store_id: activeStoreIdForUpload
          };
        }).filter(p => p.sku && p.product_name);

        if (sanitizedRaw.length === 0) {
          alert("❌ Import Cancelled: The sheet contains no valid rows with both a recognizable SKU/Item Code and Product Name.");
          setUploading(false);
          return;
        }

        const uniqueProductsMap = new Map();
        sanitizedRaw.forEach(item => {
          uniqueProductsMap.set(item.sku, item);
        });
        const sanitized = Array.from(uniqueProductsMap.values());

        // Purge old matches for the same file name inside the selected store only.
        if (activeStoreIdForUpload) {
          await supabase.from('products').delete().eq('sheet_reference', file.name).eq('store_id', activeStoreIdForUpload);
          await supabase.from('manifest_history').delete().eq('filename', file.name).eq('store_id', activeStoreIdForUpload);
        } else {
          await supabase.from('products').delete().eq('sheet_reference', file.name);
          await supabase.from('manifest_history').delete().eq('filename', file.name);
        }

        // Store the original worksheet shape, not a stripped object list.
        const originalPayload = {
          type: 'aoa-v1',
          sheet_name: worksheetName,
          headers: archive.headers,
          rows: archive.rows
        };

        const editedPayload = {
          type: 'aoa-v1',
          sheet_name: worksheetName,
          headers: [...archive.headers, ...SYSTEM_EXPORT_COLUMNS],
          rows: archive.rows.map(row => [
            ...archive.headers.map((_, index) => row[index] ?? ''),
            'Missing',
            '',
            0,
            0
          ])
        };

        const newHistoryRecord = {
          filename: file.name,
          raw_payload: originalPayload,
          edited_payload: editedPayload,
          archive_type: activeStoreNameForUpload ? `Excel Import - ${activeStoreNameForUpload}` : 'Excel Import',
          store_id: activeStoreIdForUpload,
          archived_by: loginUser || 'System Manager Account',
          created_at: currentTimestamp
        };

        // Save a local browser backup first, then try permanent Supabase history.
        // This prevents Live Matrix from falling back to only SKU/Product columns if Supabase history is blocked.
        saveManifestArchiveLocally(newHistoryRecord);

        let historySavedToDatabase = true;
        const { error: historyError } = await supabase.from('manifest_history').insert([newHistoryRecord]);
        if (historyError) {
          historySavedToDatabase = false;
          console.warn('Supabase manifest_history save failed. Local archive backup was saved:', historyError);
        }

        const { error: productUpsertError } = await supabase
          .from('products')
          .upsert(sanitized, { onConflict: activeStoreIdForUpload ? 'store_id,sku' : 'sku' });

        if (productUpsertError) throw productUpsertError;

        setManifestHistory(prev => [newHistoryRecord, ...prev.filter(h => h.filename !== file.name)]);
        setSelectedHistoryScope(file.name);
        
        alert(
          `🚀 Success! Processed "${file.name}" for ${activeStoreNameForUpload || 'General Catalog'} with [${sanitized.length}] items. New rows were marked as Missing automatically.` +
          (historySavedToDatabase
            ? ''
            : '\n\n⚠️ Important: Supabase history save was blocked, so the original Excel archive was saved in this browser only. Run the manifest_history SQL policy fix so history is permanent.')
        );
        await fetchProducts(); 
      } catch (err) {
        alert(`❌ Database Ingestion Failure!\n\nDetails: ${err.message || 'Unknown conflict error.'}`);
      } finally {
        setUploading(false);
        e.target.value = null;
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // --- AD-HOC MANUAL ENTRY HANDLER ---
  const handleAdHocSubmit = async (e) => {
    e.preventDefault();
    if (!adHocForm.sku || !adHocForm.product_name) return;

    setUploading(true);
    try {
      const activeStoreIdForAdHoc = selectedStoreId !== 'ALL' ? Number(selectedStoreId) : null;
      if (!activeStoreIdForAdHoc) {
        alert('Select a store before adding an ad-hoc product.');
        setUploading(false);
        return;
      }

      const payload = {
        sku: adHocForm.sku.trim(),
        product_name: adHocForm.product_name.trim(),
        category: adHocForm.category.trim() || 'General',
        warehouse: adHocForm.warehouse.trim() || 'Black Rose Trading',
        stock_quantity: parseInt(adHocForm.stock_quantity, 10) || 0,
        status: 'Missing',
        platforms: ['E-commerce'],
        sheet_reference: 'Ad-Hoc Manual Entry',
        store_id: activeStoreIdForAdHoc
      };

      const { error } = await supabase.from('products').upsert([payload], { onConflict: 'store_id,sku' });
      if (error) throw error;

      alert(`✅ Ad-Hoc Product [${payload.sku}] added successfully.`);
      setShowAdHocModal(false);
      setAdHocForm({ sku: '', product_name: '', category: 'General', warehouse: 'Black Rose Trading', stock_quantity: 0 });
      fetchProducts();
    } catch (err) {
      alert("Ad-Hoc Entry Failed: " + err.message);
    } finally {
      setUploading(false);
    }
  };

  const handleDashboardPurgeSheetByName = async (filename) => {
    const scopedStoreId = selectedStoreId !== 'ALL' ? Number(selectedStoreId) : null;
    const storeLabel = scopedStoreId ? getStoreNameById(scopedStoreId) : 'all stores';
    if (!window.confirm(`Are you sure you want to completely purge "${filename}" from ${storeLabel}?`)) return;
    
    setUploading(true);
    try {
      let productDelete = supabase.from('products').delete().eq('sheet_reference', filename);
      let historyDelete = supabase.from('manifest_history').delete().eq('filename', filename);

      if (scopedStoreId) {
        productDelete = productDelete.eq('store_id', scopedStoreId);
        historyDelete = historyDelete.eq('store_id', scopedStoreId);
      }

      await productDelete;
      historyDelete.then(() => {});
      if (typeof window !== 'undefined') localStorage.removeItem(getLocalManifestKey(filename, scopedStoreId));
      
      setManifestHistory(prev => prev.filter(h => !(h.filename === filename && (!scopedStoreId || Number(h.store_id || 0) === scopedStoreId))));
      if (selectedHistoryScope === filename) setSelectedHistoryScope(null);
      
      alert(`Purged spreadsheet file "${filename}" from ${storeLabel}.`);
      await fetchProducts();
    } catch (err) {
      alert("Purge failed: " + err.message);
    } finally {
      setUploading(false);
    }
  };

  // --- EXCEL SHEET DOWNLOAD HANDLER (UN-STRIPPED COLUMNS) ---
  const getRawPayloadForFilename = async (filename) => {
    const scopedStoreId = selectedStoreId !== 'ALL' ? Number(selectedStoreId) : null;

    // 1) Current React state
    const localRecord = manifestHistory.find(h => h.filename === filename && h.raw_payload && (!scopedStoreId || Number(h.store_id || 0) === scopedStoreId));
    if (localRecord?.raw_payload) return localRecord.raw_payload;

    // 2) Browser backup
    const localBackup = loadLocalManifestArchiveByName(filename, scopedStoreId);
    if (localBackup?.raw_payload) return localBackup.raw_payload;

    // 3) Supabase permanent archive. Use limit(1), not maybeSingle(), because duplicate old rows can break maybeSingle().
    try {
      let query = supabase
        .from('manifest_history')
        .select('raw_payload')
        .eq('filename', filename)
        .order('created_at', { ascending: false })
        .limit(1);

      if (scopedStoreId) query = query.eq('store_id', scopedStoreId);

      const { data, error } = await query;

      if (error) {
        console.warn('manifest_history raw_payload lookup failed:', error);
        return null;
      }

      return data?.[0]?.raw_payload || null;
    } catch (err) {
      console.warn('manifest_history raw_payload lookup crashed:', err);
      return null;
    }
  };

  const buildLiveMatrixForFilename = async (filename) => {
    const rawPayload = await getRawPayloadForFilename(filename);
    const rawMatrix = archivePayloadToMatrix(rawPayload);

    if (rawMatrix) {
      return addLiveStatusColumnsToMatrix(rawMatrix, filename);
    }

    // Fallback only. If you see this fallback, the original archive was not readable.
    return [
      ['SKU', 'Product', IMAGE_STATUS_COLUMN, 'Processed_By'],
      ...products.filter(p => p.sheet_reference === filename && isProductInSelectedStore(p)).map(p => [
        p.sku,
        p.product_name,
        p.status,
        p.processed_by || 'Unassigned'
      ])
    ];
  };

  const handleDownloadDashboardManifestByName = async (filename, mode) => {
    // Handling Ad-Hoc Entries
    if (filename === 'Ad-Hoc Manual Entry') {
      const adHocMatrix = [
        ['SKU', 'Product Name', 'Category', 'Warehouse Location', 'Stock', IMAGE_STATUS_COLUMN],
        ...products.filter(p => p.sheet_reference === filename && isProductInSelectedStore(p)).map(p => [
          p.sku,
          p.product_name,
          p.category,
          p.warehouse,
          p.stock_quantity,
          p.status
        ])
      ];
      handleDownloadSheetExport(`Live-AdHoc_Entries`, adHocMatrix);
      return;
    }

    if (mode === 'live') {
      const liveMatrix = await buildLiveMatrixForFilename(filename);
      if (liveMatrix.length <= 1) { alert('No data available to export.'); return; }
      handleDownloadSheetExport(`Live-${filename}`, liveMatrix);
      return;
    }

    const rawPayload = await getRawPayloadForFilename(filename);
    const rawMatrix = archivePayloadToMatrix(rawPayload);

    if (!rawMatrix) {
      alert('Original upload archive is still missing. Re-upload the original Excel file, or fix Supabase manifest_history permissions.');
      return;
    }

    handleDownloadSheetExport(`Original-${filename}`, rawMatrix);
  };

  const handleDownloadMissingByName = async (filename) => {
    const liveMatrix = await buildLiveMatrixForFilename(filename);

    if (!liveMatrix || liveMatrix.length <= 1) {
      alert('No data available to export.');
      return;
    }

    const headers = liveMatrix[0];
    const statusIndex = headers.findIndex(header => String(header || '').trim() === IMAGE_STATUS_COLUMN);

    if (statusIndex < 0) {
      alert('Missing status column could not be found in the live export matrix.');
      return;
    }

    const missingRows = liveMatrix.slice(1).filter(row => String(row[statusIndex] || '').trim() === 'Missing');

    if (missingRows.length === 0) {
      alert('No Missing products found for this sheet.');
      return;
    }

    handleDownloadSheetExport(`Missing-${filename}`, [headers, ...missingRows]);
  };

  const handleDownloadSheetExport = (filename, dataset) => {
    const ws = Array.isArray(dataset?.[0])
      ? XLSX.utils.aoa_to_sheet(dataset)
      : XLSX.utils.json_to_sheet(dataset);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet Logs Data');
    XLSX.writeFile(wb, filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`);
  };

  // --- BULK MEDIA ZIP DOWNLOADER ---
  const handleBulkDownloadAllAssets = async () => {
    const activeStoreProducts = products.filter(p => selectedStoreId === 'ALL' || Number(p.store_id || 0) === Number(selectedStoreId));

    if (!activeStoreProducts || activeStoreProducts.length === 0) {
      alert('No products available for the selected store scope.');
      return;
    }

    setUploading(true);
    try {
      const zip = new JSZip();
      let count = 0;
      for (const p of activeStoreProducts) {
        const edits = getArray(p.edited_image_url);
        for (let i = 0; i < edits.length; i++) {
          const url = edits[i];
          try {
            const response = await fetch(url);
            const blob = await response.blob();
            const ext = url.split('.').pop().split('?')[0] || 'jpg';
            zip.file(`${getStoreNameById(p.store_id)}/${p.sku || 'UNKNOWN'}/EDITED/asset_${i + 1}.${ext}`, blob);
            count++;
          } catch(err) { console.warn('Failed fetching', url); }
        }
      }
      if (count === 0) {
        alert('No edited assets available to download for the selected store scope.');
        setUploading(false);
        return;
      }
      const content = await zip.generateAsync({type:'blob'});
      const objectUrl = window.URL.createObjectURL(content);
      const a = document.createElement('a');
      a.href = objectUrl;
      const storeLabel = selectedStoreId === 'ALL' ? 'All-Stores' : getStoreNameById(selectedStoreId).replace(/[^a-zA-Z0-9_-]/g, '_');
      a.download = `BlackRose_${storeLabel}_Edited_Assets_${new Date().toISOString().split('T')[0]}.zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(objectUrl);
    } catch(err) {
      alert('ZIP Download error: ' + err.message);
    }
    setUploading(false);
  };

  const handleBulkDownloadAssetDirectory = async (targetProducts, assetMode = 'all') => {
    const productsForDownload = (targetProducts || []).filter(Boolean);

    if (productsForDownload.length === 0) {
      alert('No asset folders match the selected search/store filters.');
      return;
    }

    const includeRaw = assetMode === 'all' || assetMode === 'raw';
    const includeEdited = assetMode === 'all' || assetMode === 'edited';

    setUploading(true);
    try {
      const zip = new JSZip();
      let count = 0;

      for (const product of productsForDownload) {
        const cleanSku = String(product.sku || 'UNKNOWN').replace(/[^a-zA-Z0-9_-]/g, '_');
        const cleanStoreName = getStoreNameById(product.store_id).replace(/[^a-zA-Z0-9_-]/g, '_');
        const groups = [];

        if (includeRaw) groups.push({ label: 'RAW', urls: getArray(product.raw_image_url).filter(Boolean) });
        if (includeEdited) groups.push({ label: 'EDITED', urls: getArray(product.edited_image_url).filter(Boolean) });

        for (const group of groups) {
          for (let i = 0; i < group.urls.length; i++) {
            const url = group.urls[i];
            try {
              const response = await fetch(url);
              if (!response.ok) throw new Error(`HTTP ${response.status}`);
              const blob = await response.blob();
              const cleanPath = String(url).split('?')[0];
              const guessedExt = cleanPath.includes('.') ? cleanPath.split('.').pop() : 'jpg';
              const safeExt = String(guessedExt || 'jpg').replace(/[^a-zA-Z0-9]/g, '') || 'jpg';

              zip.file(`${cleanStoreName}/${cleanSku}/${group.label}/${group.label.toLowerCase()}_${i + 1}.${safeExt}`, blob);
              count++;
            } catch (err) {
              console.warn(`Failed downloading ${group.label} asset for ${product.sku}`, url, err);
            }
          }
        }
      }

      if (count === 0) {
        alert('No downloadable images found for the selected filters.');
        return;
      }

      const content = await zip.generateAsync({ type: 'blob' });
      const objectUrl = window.URL.createObjectURL(content);
      const a = document.createElement('a');
      a.href = objectUrl;
      const storeLabel = assetFolderStoreId === 'ALL' ? 'All-Stores' : getStoreNameById(assetFolderStoreId).replace(/[^a-zA-Z0-9_-]/g, '_');
      const modeLabel = assetMode === 'all' ? 'All-Images' : assetMode === 'raw' ? 'Raw-Images' : 'Edited-Images';
      a.download = `BlackRose_${storeLabel}_${modeLabel}_${new Date().toISOString().split('T')[0]}.zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(objectUrl);
    } catch (err) {
      alert('Bulk asset directory download failed: ' + (err.message || 'Unknown error'));
    } finally {
      setUploading(false);
    }
  };

  const handleDownloadProductAssets = async (product, assetType) => {
    if (!product) return;

    const isRawDownload = assetType === 'raw';
    const urls = getArray(isRawDownload ? product.raw_image_url : product.edited_image_url).filter(Boolean);
    const folderLabel = isRawDownload ? 'RAW' : 'EDITED';
    const cleanSku = String(product.sku || 'UNKNOWN').replace(/[^a-zA-Z0-9_-]/g, '_');

    if (urls.length === 0) {
      alert(`No ${folderLabel} images available for SKU ${product.sku || 'UNKNOWN'}.`);
      return;
    }

    setUploading(true);
    try {
      const zip = new JSZip();
      let downloadedCount = 0;

      for (let i = 0; i < urls.length; i++) {
        const url = urls[i];
        try {
          const response = await fetch(url);
          if (!response.ok) throw new Error(`HTTP ${response.status}`);

          const blob = await response.blob();
          const cleanPath = String(url).split('?')[0];
          const guessedExt = cleanPath.includes('.') ? cleanPath.split('.').pop() : 'jpg';
          const safeExt = String(guessedExt || 'jpg').replace(/[^a-zA-Z0-9]/g, '') || 'jpg';

          zip.file(`${cleanSku}/${folderLabel}/${folderLabel.toLowerCase()}_${i + 1}.${safeExt}`, blob);
          downloadedCount++;
        } catch (err) {
          console.warn(`Failed downloading ${folderLabel} asset`, url, err);
        }
      }

      if (downloadedCount === 0) {
        alert(`Could not download any ${folderLabel} images. Check if the storage URLs are public.`);
        return;
      }

      const content = await zip.generateAsync({ type: 'blob' });
      const objectUrl = window.URL.createObjectURL(content);
      const a = document.createElement('a');
      a.href = objectUrl;
      a.download = `${cleanSku}_${folderLabel}_Images.zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(objectUrl);
    } catch (err) {
      alert(`${folderLabel} download failed: ${err.message || 'Unknown error'}`);
    } finally {
      setUploading(false);
    }
  };

  const handleDownloadSingleAsset = async (url, suggestedName = 'image.jpg') => {
    if (!url) return;

    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      const blob = await response.blob();
      const objectUrl = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = objectUrl;
      a.download = suggestedName.replace(/[^a-zA-Z0-9._-]/g, '_') || 'image.jpg';
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(objectUrl);
    } catch (err) {
      console.warn('Single image download failed, opening image instead:', err);
      window.open(url, '_blank', 'noopener,noreferrer');
    }
  };

  const handleCardInteraction = (prod) => {
    setSelectedProduct(prod);
  };

  // --- ACCOUNT USER REGISTRATION METHODS ---
  const handleRegisterStaffAccount = async (e) => {
    e.preventDefault();
    if (!regUsername.trim() || !regPassword.trim()) return;

    try {
      const { error } = await supabase
        .from('user_registry')
        .insert([{
          username: regUsername.trim(),
          password: regPassword.trim(),
          role: regRole
        }]);

      if (error) throw error;

      alert(`Successfully registered account for "${regUsername.trim()}"`);
      setRegUsername('');
      setRegPassword('');
      fetchProducts();
    } catch (err) {
      alert("Registration failed: " + err.message);
    }
  };

  const handleRevokeStaffAccess = async (userId, userHandle) => {
    if (!window.confirm(`Are you sure you want to revoke privileges for user "${userHandle}"?`)) return;

    try {
      const { error } = await supabase
        .from('user_registry')
        .delete()
        .eq('id', userId);

      if (error) throw error;
      alert(`Privileges revoked for user "${userHandle}".`);
      fetchProducts();
    } catch (err) {
      alert("Failed to drop account context: " + err.message);
    }
  };

  // --- CUSTOM SECURITY ROLE CONFIGURATIONS ---
  const handleCreateCustomRole = (e) => {
    e.preventDefault();
    if (!newRoleForm.roleName.trim()) return;

    const formattedRole = {
      roleName: newRoleForm.roleName.trim(),
      permissions: {
        canUploadAssets: newRoleForm.canUploadAssets,
        canModifyDataSheets: newRoleForm.canModifyDataSheets,
        canReviewArrays: newRoleForm.canReviewArrays,
        canSuperviseStaff: newRoleForm.canSuperviseStaff
      }
    };

    const updatedRolesList = [...customRoles, formattedRole];
    setCustomRoles(updatedRolesList);
    localStorage.setItem('blackrose_custom_roles', JSON.stringify(updatedRolesList));
    
    setNewRoleForm({
      roleName: '',
      canUploadAssets: false,
      canModifyDataSheets: false,
      canReviewArrays: false,
      canSuperviseStaff: false
    });
    setShowRoleModal(false);
    alert(`Custom profile "${formattedRole.roleName}" deployed.`);
  };

  const handleDeleteCustomRole = (roleNameToDrop) => {
    const filtered = customRoles.filter(r => r.roleName !== roleNameToDrop);
    setCustomRoles(filtered);
    localStorage.setItem('blackrose_custom_roles', JSON.stringify(filtered));
  };

  // --- METRICS CALCULATION UTILITY ---
  const formatDurationFromSeconds = (seconds) => {
    const safeSeconds = Math.max(0, Math.floor(Number(seconds) || 0));

    if (safeSeconds <= 0) return 'No tracked time';

    const hours = Math.floor(safeSeconds / 3600);
    const minutes = Math.floor((safeSeconds % 3600) / 60);
    const remainingSeconds = safeSeconds % 60;

    if (hours > 0) return `${hours}h ${minutes}m ${remainingSeconds}s`;
    if (minutes > 0) return `${minutes}m ${remainingSeconds}s`;
    return `${remainingSeconds}s`;
  };

  const formatDateTime = (value) => {
    if (!value) return '—';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return '—';
    return date.toLocaleString();
  };

  const compileUserPerformanceMetrics = (operatorName) => {
    const currentTimestamp = new Date();
    const oneWeekAgo = new Date(currentTimestamp.getTime() - 7 * 24 * 60 * 60 * 1000);
    const oneMonthAgo = new Date(currentTimestamp.getFullYear(), currentTimestamp.getMonth(), 1);

    const getProductDate = (product) => new Date(product.updated_at || product.created_at);
    const getTrackedSeconds = (product) => Math.max(0, Number(product.total_time_spent) || 0);

    const matchUserProducts = products.filter(p => p.processed_by === operatorName);

    const completedProducts = matchUserProducts.filter(p => p.status === 'Completed');
    const modifiedProducts = matchUserProducts.filter(p => p.status === 'Modified');
    const completionLikeProducts = [...completedProducts, ...modifiedProducts];
    const rejectedProducts = matchUserProducts.filter(p => p.status === 'Rejected');
    const processingProducts = matchUserProducts.filter(p => p.status === 'Processing');
    const missingProducts = matchUserProducts.filter(p => p.status === 'Missing');

    const completedThisWeek = completedProducts.filter(p => getProductDate(p) >= oneWeekAgo).length;
    const completedThisMonth = completedProducts.filter(p => getProductDate(p) >= oneMonthAgo).length;

    const modifiedThisWeek = modifiedProducts.filter(p => getProductDate(p) >= oneWeekAgo).length;
    const modifiedThisMonth = modifiedProducts.filter(p => getProductDate(p) >= oneMonthAgo).length;

    const rejectedThisWeek = rejectedProducts.filter(p => getProductDate(p) >= oneWeekAgo).length;
    const rejectedThisMonth = rejectedProducts.filter(p => getProductDate(p) >= oneMonthAgo).length;

    const completedTimedProducts = completionLikeProducts.filter(p => getTrackedSeconds(p) > 0);
    const rejectedTimedProducts = rejectedProducts.filter(p => getTrackedSeconds(p) > 0);

    const completedTotalSeconds = completedTimedProducts.reduce((sum, p) => sum + getTrackedSeconds(p), 0);
    const rejectedTotalSeconds = rejectedTimedProducts.reduce((sum, p) => sum + getTrackedSeconds(p), 0);

    const avgCompletedSeconds = completedTimedProducts.length > 0
      ? Math.round(completedTotalSeconds / completedTimedProducts.length)
      : 0;

    const avgRejectedSeconds = rejectedTimedProducts.length > 0
      ? Math.round(rejectedTotalSeconds / rejectedTimedProducts.length)
      : 0;

    const fastestCompletedSeconds = completedTimedProducts.length > 0
      ? Math.min(...completedTimedProducts.map(getTrackedSeconds))
      : 0;

    const slowestCompletedSeconds = completedTimedProducts.length > 0
      ? Math.max(...completedTimedProducts.map(getTrackedSeconds))
      : 0;

    const doneProductLedger = matchUserProducts
      .filter(p => ['Completed', 'Modified', 'Rejected'].includes(p.status))
      .sort((a, b) => new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at))
      .map(p => ({
        id: p.id,
        sku: p.sku || '—',
        product_name: p.product_name || 'Unnamed Product',
        status: p.status || 'Missing',
        updated_at: p.updated_at || p.created_at,
        time_spent_seconds: getTrackedSeconds(p),
        time_spent_label: formatDurationFromSeconds(getTrackedSeconds(p)),
        raw_count: getArray(p.raw_image_url).length,
        edited_count: getArray(p.edited_image_url).length,
        rejection_note: p.rejection_note || ''
      }));

    return {
      weekCount: completedThisWeek,
      monthCount: completedThisMonth,
      rejectedWeekCount: rejectedThisWeek,
      rejectedMonthCount: rejectedThisMonth,
      totalClaimed: matchUserProducts.length,
      totalMissing: missingProducts.length,
      totalProcessing: processingProducts.length,
      totalCompleted: completedProducts.length,
      totalModified: modifiedProducts.length,
      totalRejected: rejectedProducts.length,
      modifiedWeekCount: modifiedThisWeek,
      modifiedMonthCount: modifiedThisMonth,
      avgCompletedTime: formatDurationFromSeconds(avgCompletedSeconds),
      avgRejectedTime: formatDurationFromSeconds(avgRejectedSeconds),
      fastestCompletedTime: formatDurationFromSeconds(fastestCompletedSeconds),
      slowestCompletedTime: formatDurationFromSeconds(slowestCompletedSeconds),
      doneProductLedger
    };
  };

  // --- BULK MEDIA ZIP EXTRACTOR ---
  const handleBulkZipUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (selectedStoreId === 'ALL') {
      alert('Select a specific store before uploading a ZIP asset folder. This prevents images from being attached to the wrong store.');
      e.target.value = null;
      return;
    }

    setUploading(true);
    const zip = new JSZip();

    const normalizeSku = (value) => {
      return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
    };

    const cleanZipName = (filename) => {
      return String(filename || '')
        .replace(/\.zip$/i, '')
        .trim();
    };

    const isImageFile = (filename) => {
      return /\.(jpg|jpeg|png|webp|gif|avif)$/i.test(String(filename || ''));
    };

    const safeFileName = (filename) => {
      return String(filename || 'image.jpg')
        .replace(/[^a-zA-Z0-9._-]/g, '_');
    };

    try {
      const contents = await zip.loadAsync(file);
      const zipNameAsSku = cleanZipName(file.name);

      // Fetch all products once. This allows exact SKU match first,
      // then a safe prefix match for ZIP names with extra barcode/color/size text.
      const { data: productCatalog, error: catalogError } = await supabase
        .from('products')
        .select('id, sku, raw_image_url, edited_image_url, status, processed_by, store_id');

      if (catalogError) throw catalogError;

      const zipScopedStoreId = selectedStoreId !== 'ALL' ? Number(selectedStoreId) : null;
      const catalog = (productCatalog || []).filter(p => !zipScopedStoreId || Number(p.store_id || 0) === zipScopedStoreId);

      const findMatchingProduct = (skuCandidates) => {
        const cleanedCandidates = Array.from(
          new Set(
            skuCandidates
              .filter(Boolean)
              .map(s => String(s).trim())
          )
        );

        // 1. Exact SKU match first.
        for (const candidate of cleanedCandidates) {
          const exact = catalog.find(p => normalizeSku(p.sku) === normalizeSku(candidate));
          if (exact) return exact;
        }

        // 2. Safe prefix match.
        // Example:
        // DB SKU: BR10-04-4
        // ZIP name: BR10-04-430025713577-BL-42cm
        for (const candidate of cleanedCandidates) {
          const candidateNorm = normalizeSku(candidate);

          const prefixMatches = catalog.filter(p => {
            const productSkuNorm = normalizeSku(p.sku);
            if (!productSkuNorm || !candidateNorm) return false;

            return (
              candidateNorm.startsWith(productSkuNorm) ||
              productSkuNorm.startsWith(candidateNorm)
            );
          });

          // Only use prefix match if exactly one product matches.
          // This prevents attaching images to the wrong SKU.
          if (prefixMatches.length === 1) return prefixMatches[0];
        }

        return null;
      };

      const productUpdates = {};
      const skippedFiles = [];
      const unmatchedSkuCandidates = new Set();
      let uploadedFileCount = 0;

      for (const [relativePath, zipEntry] of Object.entries(contents.files)) {
        if (zipEntry.dir) continue;

        const parts = relativePath.split('/').filter(Boolean);
        const originalFileName = parts[parts.length - 1];

        if (!isImageFile(originalFileName)) {
          skippedFiles.push(`${relativePath} → skipped, not an image`);
          continue;
        }

        const folderIndex = parts.findIndex(part => {
          const upper = String(part).toUpperCase();
          return upper.includes('RAW') || upper.includes('EDIT');
        });

        if (folderIndex === -1) {
          skippedFiles.push(`${relativePath} → skipped, RAW/EDITED folder not found`);
          continue;
        }

        const folderLabel = String(parts[folderIndex]).toUpperCase();
        const assetType = folderLabel.includes('RAW') ? 'RAW' : 'EDIT';

        // Supported ZIP structures:
        // 1) SKU/RAW/image.jpg
        // 2) SKU/EDITED/image.jpg
        // 3) RAW/image.jpg, where ZIP filename is the SKU
        // 4) EDITED/image.jpg, where ZIP filename is the SKU
        const skuCandidates = [];

        if (folderIndex > 0) {
          skuCandidates.push(parts[folderIndex - 1]); // usually SKU folder
          skuCandidates.push(parts[0]);               // root folder fallback
        }

        skuCandidates.push(zipNameAsSku);             // ZIP filename fallback

        const matchedProduct = findMatchingProduct(skuCandidates);

        if (!matchedProduct) {
          unmatchedSkuCandidates.add(skuCandidates.join(' OR '));
          skippedFiles.push(`${relativePath} → skipped, no matching SKU found`);
          continue;
        }

        const fileData = await zipEntry.async('blob');
        const uniqueName = `${Date.now()}-${Math.random().toString(36).slice(2)}-${safeFileName(originalFileName)}`;
        const storagePath = `${matchedProduct.sku}/${assetType}/${uniqueName}`;

        const { error: uploadError } = await supabase.storage
          .from('product-assets')
          .upload(storagePath, fileData);

        if (uploadError) {
          skippedFiles.push(`${relativePath} → storage upload failed: ${uploadError.message}`);
          console.error(`Failed to upload ${storagePath}`, uploadError);
          continue;
        }

        const { data: urlData } = supabase
          .storage
          .from('product-assets')
          .getPublicUrl(storagePath);

        if (!productUpdates[matchedProduct.id]) {
          productUpdates[matchedProduct.id] = {
            product: matchedProduct,
            raw: [],
            edit: []
          };
        }

        if (assetType === 'RAW') {
          productUpdates[matchedProduct.id].raw.push(urlData.publicUrl);
        } else {
          productUpdates[matchedProduct.id].edit.push(urlData.publicUrl);
        }

        uploadedFileCount++;
      }

      let successCount = 0;

      for (const update of Object.values(productUpdates)) {
        const currentRaw = getArray(update.product.raw_image_url);
        const currentEdit = getArray(update.product.edited_image_url);

        const nextRaw = [...currentRaw, ...update.raw];
        const nextEdit = [...currentEdit, ...update.edit];

        const updatePayload = {
          raw_image_url: nextRaw,
          edited_image_url: nextEdit
        };

        // If product was Missing and now has assets, move it to Processing.
        // Do not auto-complete because staff may still need to upload more edited images.
        if (update.product.status === 'Missing' && (nextRaw.length > 0 || nextEdit.length > 0)) {
          updatePayload.status = 'Processing';
        }

        const { error: updateError } = await supabase
          .from('products')
          .update(updatePayload)
          .eq('id', update.product.id);

        if (updateError) {
          console.error(`Failed to update product ${update.product.sku}`, updateError);
          continue;
        }

        successCount++;
      }

      let message = `ZIP Asset upload completed for ${successCount} production item(s).\nUploaded image files: ${uploadedFileCount}`;

      if (skippedFiles.length > 0) {
        message += `\nSkipped files: ${skippedFiles.length}`;
        console.warn('ZIP skipped files:', skippedFiles);
      }

      if (unmatchedSkuCandidates.size > 0) {
        message += `\n\nSome SKU names did not match products. Check browser console for details.`;
        console.warn('Unmatched SKU candidates:', Array.from(unmatchedSkuCandidates));
      }

      alert(message);

      await fetchProducts();
    } catch (err) {
      alert('ZIP upload failed: ' + (err.message || 'Unknown error'));
    } finally {
      setUploading(false);
      e.target.value = null;
    }
  };

  // --- MANUAL EDIT WORKSPACE ACTIONS ---
  const startEditing = (prod) => {
    setEditingId(prod.id);
    setEditForm({
      sku: prod.sku || '',
      product_name: prod.product_name || '',
      category: prod.category || '',
      warehouse: prod.warehouse || '',
      stock_quantity: prod.stock_quantity || 0,
      platforms: prod.platforms ? prod.platforms.join(', ') : '',
      status: prod.status || 'Missing'
    });
  };

  const handleSaveEdit = async (id) => {
    try {
      const updatedFields = {
        ...editForm,
        stock_quantity: parseInt(editForm.stock_quantity, 10) || 0,
        platforms: editForm.platforms.split(',').map(p => p.trim())
      };
      if (editForm.status === 'Missing') {
        updatedFields.processed_by = null;
        updatedFields.rejection_note = null;
      }

      const { error } = await supabase.from('products').update(updatedFields).eq('id', id);
      if (!error) {
        setEditingId(null);
        fetchProducts();
      }
    } catch (err) {
      alert('Update rejected by server constraints.');
    }
  };

  const handleOperatorStatusChange = async (id, targetStatus) => {
    const nowIso = new Date().toISOString();

    const { data: liveData, error: readError } = await supabase
      .from('products')
      .select('id, sku, status, processed_by, claimed_at, total_time_spent, rejection_note')
      .eq('id', id)
      .single();

    if (readError || !liveData) {
      alert('Could not read the latest product status. Please try again.');
      return;
    }

    const isAdminOrManager = authRole === 'Admin' || authRole === 'Manager';
    const isOwnedBySomeoneElse = liveData.processed_by && liveData.processed_by !== loginUser;

    if (!isAdminOrManager && isOwnedBySomeoneElse) {
      alert(`Action Denied: ${liveData.processed_by} has already claimed this item.`);
      return;
    }

    if (liveData.status === 'Rejected' && targetStatus !== 'Processing' && !isAdminOrManager) {
      alert('Rejected items must be moved back to Processing before they can be completed.');
      return;
    }

    const previousTrackedSeconds = Number(liveData.total_time_spent) || 0;
    const startedAt = liveData.claimed_at ? new Date(liveData.claimed_at).getTime() : null;
    const elapsedSeconds = startedAt
      ? Math.max(0, Math.floor((Date.now() - startedAt) / 1000))
      : 0;

    const updatePayload = {
      status: targetStatus,
      updated_at: nowIso
    };

    if (targetStatus === 'Processing') {
      const isRejectedReopen = liveData.status === 'Rejected';
      updatePayload.processed_by = liveData.processed_by || loginUser || 'System Operator';
      updatePayload.claimed_at = nowIso;

      // Keep the manager rejection note visible while the employee is fixing the product.
      // That note is also used to detect that this is a resubmission and should land in Modified.
      if (!isRejectedReopen) {
        updatePayload.rejection_note = null;
      }
    }

    if (targetStatus === 'Missing') {
      updatePayload.processed_by = null;
      updatePayload.claimed_at = null;
      updatePayload.total_time_spent = 0;
      updatePayload.rejection_note = null;
    }

    if (targetStatus === 'Completed') {
      const isModifiedResubmission = Boolean(liveData.rejection_note);
      updatePayload.status = isModifiedResubmission ? 'Modified' : 'Completed';
      updatePayload.processed_by = liveData.processed_by || loginUser || 'System Operator';
      updatePayload.claimed_at = null;
      updatePayload.total_time_spent = previousTrackedSeconds + elapsedSeconds;

      // Keep the old rejection note on Modified rows so the manager can see what was fixed.
      if (!isModifiedResubmission) {
        updatePayload.rejection_note = null;
      }
    }

    const { data: updatedProduct, error } = await supabase
      .from('products')
      .update(updatePayload)
      .eq('id', id)
      .select('*')
      .single();

    if (error) {
      alert('Status update failed: ' + error.message);
      return;
    }

    if (updatedProduct) {
      applyProductPatchLocally(id, updatedProduct);
    }
  };

  const handleRejectProduct = async (id) => {
    if (!rejectNote.trim()) {
      alert("Please enter a clear reason for rejecting these assets.");
      return;
    }

    const nowIso = new Date().toISOString();

    const { data: currentTargetData, error: readError } = await supabase
      .from('products')
      .select('sku, processed_by, claimed_at, total_time_spent')
      .eq('id', id)
      .single();

    if (readError || !currentTargetData) {
      alert('Could not read product before rejection.');
      return;
    }

    const originalOperator = currentTargetData?.processed_by || null;
    const previousTrackedSeconds = Number(currentTargetData.total_time_spent) || 0;
    const startedAt = currentTargetData.claimed_at ? new Date(currentTargetData.claimed_at).getTime() : null;
    const elapsedSeconds = startedAt
      ? Math.max(0, Math.floor((Date.now() - startedAt) / 1000))
      : 0;

    const { data: updatedProduct, error } = await supabase
      .from('products')
      .update({
        status: 'Rejected',
        rejection_note: rejectNote.trim(),
        processed_by: originalOperator,
        total_time_spent: previousTrackedSeconds + elapsedSeconds,
        claimed_at: null,
        updated_at: nowIso
      })
      .eq('id', id)
      .select('*')
      .single();

    if (!error) {
      if (updatedProduct) {
        applyProductPatchLocally(id, updatedProduct);
      }
      closeManagerPreview();
      fetchProducts();
    } else {
      alert("Rejection failed: " + error.message);
    }
  };

  const closeManagerPreview = () => {
    setManagerPreview(null);
    setIsRejecting(false);
    setRejectNote('');
  };

  const handleImageUpload = async (id, fieldType, e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    setImageUploadingProductId(id);

    try {
      const { data: liveProduct, error: readError } = await supabase
        .from('products')
        .select('*')
        .eq('id', id)
        .single();

      if (readError || !liveProduct) {
        throw new Error('Could not read the latest product record.');
      }

      const isAdminOrManager = authRole === 'Admin' || authRole === 'Manager';
      const isClaimedByCurrentEmployee = liveProduct.status === 'Processing' && liveProduct.processed_by === loginUser;

      if (!isAdminOrManager && !isClaimedByCurrentEmployee) {
        alert('Upload locked. First change the product status to Processing so it is claimed under your name.');
        return;
      }

      const dbField = fieldType === 'raw' ? 'raw_image_url' : 'edited_image_url';
      const folderName = fieldType === 'raw' ? 'RAW' : 'EDITED';
      const cleanSku = String(liveProduct.sku || id).replace(/[^a-zA-Z0-9_-]/g, '_');
      const uploadedUrls = [];

      for (let index = 0; index < files.length; index++) {
        const file = files[index];
        const originalName = file.name || `image-${index + 1}`;
        const fileExt = originalName.includes('.') ? originalName.split('.').pop() : 'jpg';
        const safeBaseName = originalName
          .replace(/\.[^/.]+$/, '')
          .replace(/[^a-zA-Z0-9_-]/g, '_')
          .slice(0, 60) || `image_${index + 1}`;

        const storagePath = `${cleanSku}/${folderName}/${Date.now()}-${index + 1}-${safeBaseName}.${fileExt}`;

        const { error: uploadError } = await supabase.storage
          .from('product-assets')
          .upload(storagePath, file, {
            cacheControl: '3600',
            upsert: false
          });

        if (uploadError) {
          throw new Error(`Upload failed for ${originalName}: ${uploadError.message}`);
        }

        const { data: publicData } = supabase.storage
          .from('product-assets')
          .getPublicUrl(storagePath);

        if (publicData?.publicUrl) uploadedUrls.push(publicData.publicUrl);
      }

      const nextArray = [...getArray(liveProduct[dbField]), ...uploadedUrls];

      const updatePayload = {
        [dbField]: nextArray,
        // Keep it Processing while the employee is uploading multiple batches.
        // Employee can manually mark Completed after all RAW/EDITED images are uploaded.
        status: 'Processing',
        processed_by: liveProduct.processed_by || loginUser || 'System Operator',
        updated_at: new Date().toISOString()
      };

      const { data: updatedProduct, error: dbError } = await supabase
        .from('products')
        .update(updatePayload)
        .eq('id', id)
        .select('*')
        .single();

      if (dbError) {
        throw new Error('Database save failure: ' + dbError.message);
      }

      if (updatedProduct) {
        applyProductPatchLocally(id, updatedProduct);
      }
    } catch (err) {
      alert(err.message || 'Image upload failed.');
    } finally {
      setImageUploadingProductId(null);
      if (e?.target) e.target.value = '';
    }
  };

  const handleRemoveIndividualImage = async (id, fieldType, urlToRemove) => {
    if(!window.confirm("Permanently drop this media file?")) return;
    
    const { data: existingProd } = await supabase.from('products').select('raw_image_url, edited_image_url').eq('id', id).single();
    const dbField = fieldType === 'raw' ? 'raw_image_url' : 'edited_image_url';
    
    const currentArray = getArray(existingProd[dbField]);
    const filteredArray = currentArray.filter(url => url !== urlToRemove);

    await supabase.from('products').update({ [dbField]: filteredArray }).eq('id', id);
    fetchProducts();
  };

  // --- TASK BOARD MANAGEMENT METHODS ---
  const isTaskManager = authRole === 'Admin' || authRole === 'Manager' || checkPermission('supervise_staff');

  const normalizeTaskTarget = (value) => String(value || 'All').trim().toLowerCase();

  const visibleTasks = tasks.filter(task => {
    const target = normalizeTaskTarget(task.assigned_role);
    if (target === 'all') return true;
    if (target === normalizeTaskTarget(authRole)) return true;
    if (target === normalizeTaskTarget(loginUser)) return true;
    return false;
  });

  const openTaskCount = visibleTasks.filter(task => task.status !== 'Done' && task.status !== 'Archived').length;
  const sidePanelTasks = visibleTasks
    .filter(task => task.status !== 'Archived')
    .slice(0, 3);

  const resetTaskForm = () => {
    setTaskForm({ title: '', description: '', priority: 'Normal', assigned_role: 'All', due_at: '' });
    setEditingTaskId(null);
  };

  const handleTaskSubmit = async (e) => {
    e.preventDefault();
    if (!isTaskManager) return;

    const cleanTitle = taskForm.title.trim();
    if (!cleanTitle) {
      alert('Task title is required.');
      return;
    }

    setTaskSaving(true);
    try {
      const nowIso = new Date().toISOString();
      const payload = {
        title: cleanTitle,
        description: taskForm.description.trim(),
        priority: taskForm.priority || 'Normal',
        assigned_role: taskForm.assigned_role || 'All',
        due_at: taskForm.due_at || null,
        updated_by: loginUser || 'System Manager',
        updated_at: nowIso
      };

      if (editingTaskId) {
        const { data, error } = await supabase
          .from('task_board')
          .update(payload)
          .eq('id', editingTaskId)
          .select('*')
          .single();

        if (error) throw error;
        if (data) setTasks(prev => prev.map(task => task.id === data.id ? data : task));
      } else {
        const { data, error } = await supabase
          .from('task_board')
          .insert([{ ...payload, created_by: loginUser || 'System Manager', status: 'Open' }])
          .select('*')
          .single();

        if (error) throw error;
        if (data) setTasks(prev => [data, ...prev]);
      }

      resetTaskForm();
      setTaskError(null);
    } catch (err) {
      alert('Task save failed: ' + (err.message || 'Unknown error'));
      setTaskError(err.message || 'Task save failed.');
    } finally {
      setTaskSaving(false);
    }
  };

  const handleStartEditTask = (task) => {
    setEditingTaskId(task.id);
    setTaskForm({
      title: task.title || '',
      description: task.description || '',
      priority: task.priority || 'Normal',
      assigned_role: task.assigned_role || 'All',
      due_at: task.due_at ? String(task.due_at).slice(0, 16) : ''
    });
    setActiveTab('task_board');
  };

  const handleDeleteTask = async (taskId) => {
    if (!isTaskManager) return;
    if (!window.confirm('Delete this task permanently?')) return;

    try {
      const { error } = await supabase.from('task_board').delete().eq('id', taskId);
      if (error) throw error;
      setTasks(prev => prev.filter(task => task.id !== taskId));
      if (editingTaskId === taskId) resetTaskForm();
    } catch (err) {
      alert('Task delete failed: ' + (err.message || 'Unknown error'));
    }
  };

  const handleUpdateTaskStatus = async (taskId, nextStatus) => {
    if (!taskId || !nextStatus) return;

    try {
      const { data, error } = await supabase
        .from('task_board')
        .update({ status: nextStatus, updated_by: loginUser || 'System User', updated_at: new Date().toISOString() })
        .eq('id', taskId)
        .select('*')
        .single();

      if (error) throw error;
      if (data) setTasks(prev => prev.map(task => task.id === data.id ? data : task));
    } catch (err) {
      alert('Task status update failed: ' + (err.message || 'Unknown error'));
    }
  };

  const getPriorityClass = (priority) => {
    if (priority === 'Urgent') return 'bg-red-100 text-red-700 border-red-200';
    if (priority === 'High') return 'bg-orange-100 text-orange-700 border-orange-200';
    if (priority === 'Low') return 'bg-slate-100 text-slate-600 border-slate-200';
    return 'bg-[rgba(138,21,56,0.10)] text-[#8a1538] border-[rgba(138,21,56,0.28)]';
  };

  const getTaskStatusClass = (status) => {
    if (status === 'Done') return 'bg-green-100 text-green-700 border-green-200';
    if (status === 'In Progress') return 'bg-[rgba(138,21,56,0.10)] text-[#8a1538] border-[rgba(138,21,56,0.28)]';
    if (status === 'Archived') return 'bg-gray-100 text-gray-500 border-gray-200';
    return 'bg-amber-100 text-amber-700 border-amber-200';
  };

  const formatDisplayDateTime = (value) => {
    if (!value) return 'No deadline';
    try {
      return new Date(value).toLocaleString([], {
        year: 'numeric',
        month: 'short',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch (err) {
      return String(value);
    }
  };

  // --- FILTER ROW MATRICES ---
  const filteredProducts = products.filter((prod) => {
    const matchesStatus = statusFilter === 'All' || prod.status === statusFilter;
    const matchesStore = isProductInSelectedStore(prod);
    const matchesSheetContext = !selectedHistoryScope || prod.sheet_reference === selectedHistoryScope;
    
    const matchesOperatorBound = authRole === 'Admin' || authRole === 'Manager' || 
      authRole === 'Content Editor' || checkPermission('modify_sheets') ||
      (!prod.processed_by || prod.processed_by === loginUser);

    const searchLower = searchQuery.toLowerCase().trim();
    return matchesStatus && matchesStore && matchesSheetContext && matchesOperatorBound && (
      searchLower === '' ||
      prod.product_name?.toLowerCase().includes(searchLower) ||
      prod.sku?.toLowerCase().includes(searchLower) ||
      prod.warehouse?.toLowerCase().includes(searchLower)
    );
  });

  const scopedProducts = products.filter(p => isProductInSelectedStore(p));

  const assetDirectoryProducts = products.filter((prod) => {
    const raws = getArray(prod.raw_image_url);
    const edits = getArray(prod.edited_image_url);
    const hasAssets = raws.length > 0 || edits.length > 0;
    const matchesAssetStore = assetFolderStoreId === 'ALL' || Number(prod.store_id || 0) === Number(assetFolderStoreId);
    const searchLower = assetFolderSearchQuery.toLowerCase().trim();
    const matchesAssetSearch =
      searchLower === '' ||
      prod.sku?.toLowerCase().includes(searchLower) ||
      prod.product_name?.toLowerCase().includes(searchLower) ||
      getStoreNameById(prod.store_id).toLowerCase().includes(searchLower);

    return hasAssets && matchesAssetStore && matchesAssetSearch;
  });

  const assetDirectoryRawCount = assetDirectoryProducts.reduce((sum, prod) => sum + getArray(prod.raw_image_url).filter(Boolean).length, 0);
  const assetDirectoryEditedCount = assetDirectoryProducts.reduce((sum, prod) => sum + getArray(prod.edited_image_url).filter(Boolean).length, 0);

  const metrics = {
    total: scopedProducts.filter(p => !selectedHistoryScope || p.sheet_reference === selectedHistoryScope).length,
    missing: scopedProducts.filter(p => p.status === 'Missing' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    processing: scopedProducts.filter(p => p.status === 'Processing' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    completed: scopedProducts.filter(p => p.status === 'Completed' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    modified: scopedProducts.filter(p => p.status === 'Modified' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    rejected: scopedProducts.filter(p => p.status === 'Rejected' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    lowStock: scopedProducts.filter(p => p.stock_quantity < 5 && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length
  };

  const productUploadedSheets = Array.from(
    new Set(
      scopedProducts
        .map(p => p.sheet_reference)
        .filter(s => s && s !== 'Ad-Hoc Manual Entry')
    )
  );

  const historyUploadedSheets = Array.from(
    new Set(
      manifestHistory
        .filter(h => !isStoreScoped || Number(h.store_id || 0) === Number(selectedStoreId))
        .map(h => h.filename)
        .filter(Boolean)
    )
  );

  const uniqueUploadedSheets = Array.from(
    new Set([...productUploadedSheets, ...historyUploadedSheets])
  );

  const originalUploadedSheets = historyUploadedSheets;

  const hasOriginalArchive = (filename) => {
    return manifestHistory.some(h => h.filename === filename && h.raw_payload && (!isStoreScoped || Number(h.store_id || 0) === Number(selectedStoreId))) || !!loadLocalManifestArchiveByName(filename, isStoreScoped ? selectedStoreId : null)?.raw_payload;
  };

  const storeCards = stores.map(store => {
    const storeProducts = products.filter(p => Number(p.store_id || 0) === Number(store.id));
    return {
      ...store,
      total: storeProducts.length,
      missing: storeProducts.filter(p => p.status === 'Missing').length,
      processing: storeProducts.filter(p => p.status === 'Processing').length,
      completed: storeProducts.filter(p => p.status === 'Completed').length,
      modified: storeProducts.filter(p => p.status === 'Modified').length,
      rejected: storeProducts.filter(p => p.status === 'Rejected').length
    };
  });

  const unassignedProductsCount = products.filter(p => !p.store_id).length;

  const canViewMyPerformance = isLoggedIn && authRole !== 'Admin' && authRole !== 'Manager';
  const selfPerformanceStats = canViewMyPerformance ? compileUserPerformanceMetrics(loginUser) : null;

  return (
    <div className="min-h-screen bg-gray-50 font-sans flex flex-col selection:bg-[rgba(138,21,56,0.85)] selection:text-white">
      
      {/* GLOBAL NAVIGATION NAVBAR */}
      <header className="w-full bg-white border-b border-gray-200 sticky top-0 z-50 shrink-0">
        <div className="w-full max-w-none px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="h-12 w-12 shrink-0 flex items-center justify-center overflow-hidden">
              <img 
                src="https://xjzyhzmqibcnvdtrtdcs.supabase.co/storage/v1/object/public/product-assets/logo.jpeg" 
                alt="Black-Rose Logo" className="max-h-12 max-w-12 w-auto h-auto object-contain" 
              />
            </div>
            <div>
              <span className="font-black text-gray-950 tracking-tight block text-sm uppercase">Black-Rose</span>
              <span className="text-xxs text-gray-400 font-bold uppercase tracking-widest block -mt-0.5">Management Hub</span>
            </div>
          </div>
          <nav className="flex items-center gap-4">
            {isLoggedIn ? (
              <>
                <span className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-xl text-xs font-bold uppercase tracking-wider border ${
                  authRole === 'Admin' ? 'bg-[rgba(138,21,56,0.06)] text-[#8a1538] border-[rgba(138,21,56,0.28)] shadow-xs' :
                  authRole === 'Manager' ? 'bg-[rgba(138,21,56,0.06)] text-[#8a1538] border-[rgba(138,21,56,0.28)]' : 'bg-amber-50 text-amber-700 border-amber-200'
                }`}>
                   Signature: <span className="underline font-black">{loginUser} [{authRole}]</span>
                </span>
                <button onClick={handlePortalLogout} className="px-4 py-1.5 bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 font-bold text-xs uppercase tracking-wider rounded-xl cursor-pointer transition-all">
                  Logout
                </button>
              </>
            ) : (
              <span className="px-3 py-1.5 bg-gray-50 border border-gray-200 text-gray-400 rounded-xl text-xs font-bold uppercase tracking-wider">
                System Key Required
              </span>
            )}
          </nav>
        </div>
      </header>

      {/* COMPONENT CONTENT BODY PANEL WRAPPER */}
      <div className="flex-grow flex flex-row min-h-0 w-full">
        {isLoggedIn && activeTab !== 'home' && (
          <aside className={`border-r border-gray-200 bg-white flex flex-col justify-between shrink-0 transition-all duration-300 ${isSidePanelCollapsed ? 'w-20 p-3' : 'w-64 p-4'}`}>
            <div className="space-y-4">
              <div className={`flex items-center ${isSidePanelCollapsed ? 'justify-center' : 'justify-between'} border-b border-gray-100 pb-4`}>
                <button
                  onClick={() => setActiveTab('home')}
                  className={`flex items-center ${isSidePanelCollapsed ? 'justify-center' : 'gap-3'} min-w-0`}
                  title="Go to Home"
                >
                  <span className={`${isSidePanelCollapsed ? 'w-12 h-12' : 'w-14 h-14'} rounded-2xl bg-white border border-gray-100 shadow-xs flex items-center justify-center overflow-hidden shrink-0`}>
                    <img
                      src="https://xjzyhzmqibcnvdtrtdcs.supabase.co/storage/v1/object/public/product-assets/logo.jpeg"
                      alt="Black-Rose Logo"
                      className={`${isSidePanelCollapsed ? 'max-h-11 max-w-11' : 'max-h-[52px] max-w-[52px]'} object-contain`}
                    />
                  </span>
                  {!isSidePanelCollapsed && (
                    <span className="min-w-0 text-left">
                      <span className="block text-xs font-black uppercase tracking-tight text-gray-950">Black-Rose</span>
                      <span className="block text-[10px] font-black uppercase tracking-widest text-gray-400 truncate">Management Hub</span>
                    </span>
                  )}
                </button>

                {!isSidePanelCollapsed && (
                  <button
                    type="button"
                    onClick={() => setIsSidePanelCollapsed(true)}
                    className="w-8 h-8 rounded-xl border border-gray-200 bg-gray-50 hover:bg-gray-100 text-gray-500 font-black text-xs transition-all"
                    title="Collapse side panel"
                  >
                    ◀
                  </button>
                )}
              </div>

              {isSidePanelCollapsed ? (
                <div className="space-y-2">
                  <button onClick={() => setIsSidePanelCollapsed(false)} className="w-full h-10 rounded-xl border border-gray-200 bg-gray-50 hover:bg-gray-100 text-gray-600 font-black text-xs" title="Expand side panel">▶</button>
                  <button onClick={() => setActiveTab('home')} className="w-full h-10 rounded-xl bg-white hover:bg-gray-100 text-sm border border-gray-100" title="Home">🏠</button>
                  <button onClick={() => setActiveTab('dashboard')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'dashboard' ? 'bg-[rgba(138,21,56,0.85)] text-white border-[rgba(138,21,56,0.85)]' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="Products Dashboard">📊</button>
                  <button onClick={() => setActiveTab('task_board')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'task_board' ? 'bg-indigo-700 text-white border-indigo-700' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="Task Board">🧾</button>
                  {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('supervise_staff')) && (
                    <button onClick={() => setActiveTab('operators')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'operators' ? 'bg-teal-700 text-white border-teal-700' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="Active Operators">👥</button>
                  )}
                  {canViewMyPerformance && selfPerformanceStats && (
                    <button onClick={() => setActiveTab('my_performance')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'my_performance' ? 'bg-amber-600 text-white border-amber-600' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="My Performance">📈</button>
                  )}
                  {authRole === 'Admin' && (
                    <button onClick={() => setActiveTab('admin_panel')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'admin_panel' ? 'bg-[rgba(138,21,56,0.85)] text-white border-[rgba(138,21,56,0.85)]' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="Admin Panel">👑</button>
                  )}
                  <div className="border-t border-gray-100 pt-2 space-y-2">
                    <button onClick={() => setActiveTab('history_old')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'history_old' ? 'bg-amber-600 text-white border-amber-600' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="Original Sheets">📂</button>
                    <button onClick={() => setActiveTab('history_new')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'history_new' ? 'bg-[rgba(138,21,56,0.85)] text-white border-[rgba(138,21,56,0.85)]' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="Live Sheets">📈</button>
                    <button onClick={() => setActiveTab('images')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'images' ? 'bg-[rgba(138,21,56,0.85)] text-white border-[rgba(138,21,56,0.85)]' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="SKU Assets">🖼️</button>
                  </div>
                </div>
              ) : (
                <div className="space-y-4">
                  <div>
                    <h3 className="text-xxs font-black tracking-widest uppercase text-gray-400 px-2 mb-2">Navigation</h3>
                <div className="space-y-1">
                  <button
                    onClick={() => setActiveTab('home')}
                    className="w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all text-gray-700 hover:bg-gray-100"
                  >
                    🏠 Home
                  </button>

                  <button 
                    onClick={() => setActiveTab('dashboard')} 
                    className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'dashboard' ? 'bg-[rgba(138,21,56,0.85)] text-white shadow-xs' : 'text-gray-700 hover:bg-gray-100'}`}
                  >
                    📊 Products Dashboard
                  </button>

                  <button 
                    onClick={() => setActiveTab('task_board')} 
                    className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'task_board' ? 'bg-indigo-700 text-white shadow-xs' : 'text-gray-700 hover:bg-gray-100'}`}
                  >
                    🧾 Task Board {openTaskCount > 0 ? `(${openTaskCount})` : ''}
                  </button>

                  {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('supervise_staff')) && (
                    <button 
                      onClick={() => setActiveTab('operators')} 
                      className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'operators' ? 'bg-teal-700 text-white shadow-xs' : 'text-gray-700 hover:bg-gray-100'}`}
                    >
                      👥 Active Operators
                    </button>
                  )}

                  {canViewMyPerformance && selfPerformanceStats && (
                    <button 
                      onClick={() => setActiveTab('my_performance')} 
                      className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'my_performance' ? 'bg-amber-600 text-white shadow-xs' : 'text-gray-700 hover:bg-gray-100'}`}
                    >
                      📈 My Performance
                    </button>
                  )}

                  {authRole === 'Admin' && (
                    <button 
                      onClick={() => setActiveTab('admin_panel')} 
                      className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'admin_panel' ? 'bg-[rgba(138,21,56,0.85)] text-white shadow-md' : 'text-[#8a1538] hover:bg-[rgba(138,21,56,0.06)]'}`}
                    >
                      👑 Admin Panel
                    </button>
                  )}
                </div>
              </div>

              <div className="border-t pt-4">
                <div className="flex items-center justify-between px-2 mb-2">
                  <button
                    type="button"
                    onClick={() => setIsStoresPanelOpen(prev => !prev)}
                    className="flex items-center gap-2 text-xxs font-black tracking-widest uppercase text-gray-400 hover:text-gray-700 transition-colors"
                    title={isStoresPanelOpen ? 'Collapse Stores' : 'Expand Stores'}
                  >
                    <span className="text-[10px]">{isStoresPanelOpen ? '▾' : '▸'}</span>
                    Stores
                    <span className="text-[10px] text-gray-300">({stores.length})</span>
                  </button>
                  {(authRole === 'Admin' || authRole === 'Manager') && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        setIsStoresPanelOpen(true);
                        setShowStoreCreate(prev => !prev);
                      }}
                      className="text-[10px] font-black text-[rgba(138,21,56,0.85)] hover:text-[rgba(138,21,56,0.95)]"
                      title="Add store"
                    >
                      + Add
                    </button>
                  )}
                </div>

                {isStoresPanelOpen && (authRole === 'Admin' || authRole === 'Manager') && showStoreCreate && (
                  <form onSubmit={handleCreateStore} className="mb-2 p-2 bg-[rgba(138,21,56,0.06)] border border-[rgba(138,21,56,0.18)] rounded-xl space-y-2">
                    <input
                      type="text"
                      value={newStoreName}
                      onChange={(e) => setNewStoreName(e.target.value)}
                      placeholder="Store name..."
                      className="w-full px-2 py-2 text-xs border rounded-lg bg-white text-gray-900 outline-none"
                    />
                    <button type="submit" className="w-full py-2 bg-[rgba(138,21,56,0.85)] text-white text-[10px] font-black uppercase rounded-lg">
                      Save Store
                    </button>
                  </form>
                )}

                {isStoresPanelOpen && (
                  <div className="space-y-1 max-h-56 overflow-y-auto pr-1">
                  <button
                    onClick={() => handleSelectStore('ALL', 'dashboard')}
                    className={`w-full text-left px-3 py-2 text-xs font-bold uppercase tracking-wide rounded-lg transition-all ${selectedStoreId === 'ALL' ? 'bg-gray-900 text-white' : 'text-gray-600 hover:bg-gray-100'}`}
                  >
                    🏬 All Stores
                    <span className="float-right text-[10px] opacity-70">{products.length}</span>
                  </button>

                  {stores.length === 0 ? (
                    <div className="text-[10px] font-bold text-gray-400 uppercase bg-gray-50 border border-gray-100 rounded-lg p-2 text-center">
                      No stores yet
                    </div>
                  ) : (
                    stores.map(store => {
                      const count = products.filter(p => Number(p.store_id || 0) === Number(store.id)).length;
                      const isActiveStore = Number(selectedStoreId) === Number(store.id);
                      return (
                        <div
                          key={store.id}
                          className={`rounded-lg transition-all border ${isActiveStore ? 'bg-emerald-700 text-white border-emerald-700 shadow-xs' : 'bg-white text-gray-700 border-gray-100 hover:bg-emerald-50'}`}
                        >
                          <button
                            onClick={() => handleSelectStore(store.id, 'dashboard')}
                            className="w-full text-left px-3 py-2 text-xs font-bold rounded-lg transition-all"
                          >
                            <span className="inline-flex items-center gap-2 min-w-0 max-w-[150px] align-bottom">
                              {store.image_url ? (
                                <img src={store.image_url} alt={store.name} className="w-5 h-5 rounded-md object-cover border border-white/40 shrink-0" />
                              ) : (
                                <span className="w-5 h-5 rounded-md bg-emerald-100 text-emerald-700 flex items-center justify-center text-[10px] shrink-0">🏪</span>
                              )}
                              <span className="truncate">{store.name}</span>
                            </span>
                            <span className="float-right text-[10px] opacity-80">{count}</span>
                          </button>

                          {(authRole === 'Admin' || authRole === 'Manager') && (
                            <div className="mx-2 mb-2 grid grid-cols-2 gap-1.5">
                              <label
                                className={`px-2 py-1 rounded-md text-center text-[9px] font-black uppercase cursor-pointer border transition-all ${isActiveStore ? 'bg-white/10 text-white border-white/20 hover:bg-white/20' : 'bg-gray-50 text-gray-500 border-gray-100 hover:bg-gray-100'}`}
                                title="Upload or replace store card image"
                              >
                                {storeImageUploadingId === store.id ? 'Uploading...' : store.image_url ? 'Image' : 'Upload'}
                                <input
                                  type="file"
                                  accept="image/*"
                                  onChange={(e) => handleStoreImageUpload(store.id, e)}
                                  className="hidden"
                                />
                              </label>
                              <button
                                type="button"
                                onClick={() => handleDeleteStore(store.id, store.name)}
                                disabled={uploading}
                                className={`px-2 py-1 rounded-md text-center text-[9px] font-black uppercase border transition-all ${isActiveStore ? 'bg-red-500/20 text-white border-red-200/30 hover:bg-red-500/30' : 'bg-red-50 text-red-600 border-red-100 hover:bg-red-100'} ${uploading ? 'opacity-50 cursor-not-allowed' : 'cursor-pointer'}`}
                                title="Delete this store and its linked products/history"
                              >
                                Delete
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })
                  )}
                  </div>
                )}
              </div>

              <div className="border-t pt-4">
                <h3 className="text-xxs font-black tracking-widest uppercase text-gray-400 px-2 mb-2">Files & Media</h3>
                <div className="space-y-1">
                  <button 
                    onClick={() => setActiveTab('history_old')} 
                    className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'history_old' ? 'bg-amber-600 text-white' : 'text-gray-700 hover:bg-gray-100'}`}
                  >
                    📂 Original Sheets
                  </button>
                  <button 
                    onClick={() => setActiveTab('history_new')} 
                    className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'history_new' ? 'bg-[rgba(138,21,56,0.85)] text-white' : 'text-gray-700 hover:bg-gray-100'}`}
                  >
                    📈 Live Sheets
                  </button>
                  <button 
                    onClick={() => setActiveTab('images')} 
                    className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'images' ? 'bg-[rgba(138,21,56,0.85)] text-white' : 'text-gray-700 hover:bg-gray-100'}`}
                  >
                    🖼️ SKU Assets
                  </button>
                </div>
              </div>
                </div>
              )}
            </div>

            <div className="p-2 border-t text-center">
              {isSidePanelCollapsed ? (
                <span className="text-[10px] font-mono text-gray-400">v4.2</span>
              ) : (
                <span className="text-[10px] font-mono text-gray-400">v4.2 Wide Layout</span>
              )}
            </div>
          </aside>
        )}

        {/* MAIN APPLICATION VIEWPORT CANVAS */}
        <main className="flex-grow overflow-y-auto p-4 sm:p-6 lg:p-8 w-full">
          <div className="w-full max-w-none">
            
            {!isLoggedIn ? (
              <div className="min-h-[70vh] flex items-center justify-center p-4">
                <div className="w-full max-w-md bg-white p-8 rounded-2xl shadow-xl border border-gray-200/60">
                  <div className="text-center mb-8">
                    <span className="text-4xl">🔑</span>
                    <h2 className="text-2xl font-black text-gray-900 tracking-tight mt-3 uppercase">Staff Entry Gateway</h2>
                    <p className="text-xs text-gray-400 font-semibold mt-1">Provide credentials to clear secure channel ports</p>
                  </div>
                  
                  {authError && (
                    <div className="mb-4 p-3 bg-red-50 border border-red-200 text-red-700 rounded-lg text-xs font-bold text-center">
                      🚨 {authError}
                    </div>
                  )}

                  <form onSubmit={handlePortalLogin} className="space-y-4">
                    <div>
                      <label className="block text-xxs font-black text-gray-400 uppercase tracking-widest mb-1">Username / ID</label>
                      <input type="text" required value={loginUser} onChange={(e) => setLoginUser(e.target.value)} placeholder="Enter username..." className="w-full px-4 py-3 border border-gray-200 rounded-xl bg-gray-50 text-sm text-gray-900 outline-none" />
                    </div>
                    <div>
                      <label className="block text-xxs font-black text-gray-400 uppercase tracking-widest mb-1">Access Passcode</label>
                      <input type="password" required value={loginPass} onChange={(e) => setLoginPass(e.target.value)} placeholder="•••••••••••••" className="w-full px-4 py-3 border border-gray-200 rounded-xl bg-gray-50 text-sm text-gray-900 outline-none" />
                    </div>
                    <button type="submit" className="w-full py-3 bg-[rgba(138,21,56,0.85)] hover:bg-[rgba(138,21,56,0.95)] text-white text-xs font-black uppercase tracking-widest rounded-xl transition-all shadow-md cursor-pointer">
                      Authenticate Session
                    </button>
                  </form>
                </div>
              </div>
            ) : (
              <div className="animate-fadeIn">

                {/* HOME LANDING PAGE */}
                {activeTab === 'home' && (
                  <div className="space-y-8">
                    <div className="bg-gradient-to-r from-[rgba(138,21,56,1)] via-[rgba(138,21,56,0.85)] to-emerald-900 p-8 rounded-3xl border border-[rgba(138,21,56,0.95)] text-white shadow-xl relative overflow-hidden">
                      <div className="absolute right-0 bottom-0 translate-x-1/4 translate-y-1/4 w-96 h-96 bg-emerald-500/10 rounded-full blur-3xl"></div>
                      <div className="relative z-10 flex flex-col lg:flex-row lg:items-end lg:justify-between gap-6">
                        <div>
                          <span className="text-xxs uppercase tracking-widest font-black text-emerald-300 bg-emerald-900/40 border border-emerald-700/50 px-2.5 py-1 rounded-md">
                            {authRole === 'Operator' ? 'Employee Home' : authRole === 'Admin' ? 'Admin Home' : 'Manager Home'}
                          </span>
                          <h1 className="text-3xl font-black mt-3 tracking-tight">Stores & Task Command Center</h1>
                          <p className="text-sm text-slate-300 font-medium mt-1">
                            {authRole === 'Operator'
                              ? 'Select a store or task to start work. The side panel appears after you open a workspace.'
                              : 'Review store workload, post tasks, and open the workspace you need. The side panel appears after you leave Home.'}
                          </p>
                        </div>
                        <div className="flex flex-col sm:items-end gap-3">
                          <div className="text-right text-base font-bold text-slate-200">
                            {dashboardClock.toLocaleDateString([], { weekday: 'short', year: 'numeric', month: 'short', day: '2-digit' })} • {dashboardClock.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                          </div>
                          <div className="flex flex-wrap gap-2 justify-start sm:justify-end">
                            <button onClick={() => setActiveTab('task_board')} className="px-4 py-2 bg-white/10 hover:bg-white/20 border border-white/20 rounded-xl text-xs font-black uppercase tracking-wider">
                              Open Tasks
                            </button>
                            {(authRole === 'Admin' || authRole === 'Manager') && (
                              <button onClick={() => setActiveTab('operators')} className="px-4 py-2 bg-white/10 hover:bg-white/20 border border-white/20 rounded-xl text-xs font-black uppercase tracking-wider">
                                Operators
                              </button>
                            )}
                            {authRole === 'Admin' && (
                              <button onClick={() => setActiveTab('admin_panel')} className="px-4 py-2 bg-[#8a1538] hover:bg-[#8a1538] border border-[rgba(138,21,56,0.50)] rounded-xl text-xs font-black uppercase tracking-wider">
                                Admin Panel
                              </button>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 xl:grid-cols-4 gap-6 w-full">
                      <div className="xl:col-span-3 bg-white border border-gray-200 rounded-2xl p-6 shadow-xs">
                        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-5">
                          <div>
                            <h2 className="text-lg font-black text-gray-900 uppercase">Stores Overview</h2>
                            <p className="text-xs text-gray-400 font-bold mt-1">
                              {authRole === 'Operator' ? 'Pick the store you want to work on.' : 'Pick a store to open its dashboard or add a new store.'}
                            </p>
                          </div>
                          <div className="flex flex-wrap gap-2">
                            {unassignedProductsCount > 0 && (
                              <button onClick={() => handleSelectStore('ALL', 'dashboard')} className="px-3 py-2 rounded-xl bg-gray-100 text-gray-700 text-xs font-black uppercase">
                                View All ({products.length})
                              </button>
                            )}
                            {(authRole === 'Admin' || authRole === 'Manager') && (
                              <button onClick={() => setShowStoreCreate(prev => !prev)} className="px-3 py-2 rounded-xl bg-[rgba(138,21,56,0.85)] text-white text-xs font-black uppercase">
                                + Add Store
                              </button>
                            )}
                          </div>
                        </div>

                        {(authRole === 'Admin' || authRole === 'Manager') && showStoreCreate && (
                          <form onSubmit={handleCreateStore} className="mb-5 p-4 bg-[rgba(138,21,56,0.06)] border border-[rgba(138,21,56,0.18)] rounded-2xl grid grid-cols-1 sm:grid-cols-4 gap-3">
                            <input
                              type="text"
                              value={newStoreName}
                              onChange={(e) => setNewStoreName(e.target.value)}
                              placeholder="New store name..."
                              className="sm:col-span-3 w-full px-3 py-2 text-sm border rounded-xl bg-white text-gray-900 outline-none"
                            />
                            <button type="submit" className="py-2 bg-[rgba(138,21,56,0.85)] text-white text-xs font-black uppercase rounded-xl">
                              Save Store
                            </button>
                          </form>
                        )}

                        {storeCards.length === 0 ? (
                          <div className="text-center py-16 bg-gray-50 border border-dashed rounded-2xl">
                            <div className="text-4xl mb-3">🏪</div>
                            <div className="text-sm font-black text-gray-700 uppercase">No stores added yet</div>
                            <p className="text-xs text-gray-400 font-semibold mt-1">
                              {(authRole === 'Admin' || authRole === 'Manager') ? 'Add a store above, then upload its Excel sheet from the store dashboard.' : 'Ask your manager to create a store and upload its Excel sheet.'}
                            </p>
                          </div>
                        ) : (
                          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-4 gap-5">
                            {storeCards.map(store => (
                              <div
                                key={store.id}
                                className="group overflow-hidden rounded-2xl border border-gray-200 bg-white hover:border-emerald-400 hover:shadow-md transition-all"
                              >
                                <button
                                  type="button"
                                  onClick={() => handleSelectStore(store.id, 'dashboard')}
                                  className="w-full text-left"
                                >
                                  <div className="relative h-32 bg-gradient-to-br from-emerald-900 via-[rgba(138,21,56,0.95)] to-[rgba(138,21,56,1)] overflow-hidden">
                                    {store.image_url ? (
                                      <img
                                        src={store.image_url}
                                        alt={store.name}
                                        className="w-full h-full object-cover transition-transform duration-300 group-hover:scale-105"
                                      />
                                    ) : (
                                      <div className="w-full h-full flex items-center justify-center text-5xl text-white/80">🏪</div>
                                    )}
                                    <div className="absolute inset-0 bg-gradient-to-t from-black/60 via-black/10 to-transparent"></div>
                                    <div className="absolute left-4 bottom-3 right-4 flex items-end justify-between gap-3">
                                      <div className="min-w-0">
                                        <div className="text-lg font-black text-white truncate">{store.name}</div>
                                        <div className="text-[10px] font-black text-emerald-100 uppercase tracking-wider">Total Products</div>
                                      </div>
                                      <div className="text-3xl font-black text-white shrink-0">{store.total}</div>
                                    </div>
                                  </div>

                                  <div className="p-4">
                                    <div className="grid grid-cols-3 gap-2 text-center">
                                      <div className="bg-amber-50 border border-amber-200 rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-amber-700 font-black">Missing</span>
                                        <b className="block text-xl font-black text-amber-800 leading-tight">{store.missing}</b>
                                      </div>
                                      <div className="bg-[rgba(138,21,56,0.06)] border border-[rgba(138,21,56,0.28)] rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-[#8a1538] font-black">Process</span>
                                        <b className="block text-xl font-black text-[rgba(138,21,56,0.95)] leading-tight">{store.processing}</b>
                                      </div>
                                      <div className="bg-red-50 border border-red-200 rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-red-700 font-black">Reject</span>
                                        <b className="block text-xl font-black text-red-800 leading-tight">{store.rejected}</b>
                                      </div>
                                    </div>

                                    <div className="grid grid-cols-2 gap-2 mt-2 text-center">
                                      <div className="bg-green-50 border border-green-200 rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-green-700 font-black">Done</span>
                                        <b className="block text-lg font-black text-green-800 leading-tight">{store.completed}</b>
                                      </div>
                                      <div className="bg-[rgba(138,21,56,0.06)] border border-[rgba(138,21,56,0.28)] rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-[#8a1538] font-black">Modified</span>
                                        <b className="block text-lg font-black text-[#8a1538] leading-tight">{store.modified}</b>
                                      </div>
                                    </div>
                                  </div>
                                </button>

                                {(authRole === 'Admin' || authRole === 'Manager') && (
                                  <div className="px-4 pb-4 grid grid-cols-2 gap-2">
                                    <label className="px-3 py-2 rounded-xl text-center text-[10px] font-black uppercase cursor-pointer border bg-gray-50 text-gray-600 border-gray-100 hover:bg-gray-100">
                                      {storeImageUploadingId === store.id ? 'Uploading...' : store.image_url ? 'Replace Image' : 'Upload Image'}
                                      <input
                                        type="file"
                                        accept="image/*"
                                        onChange={(e) => handleStoreImageUpload(store.id, e)}
                                        className="hidden"
                                      />
                                    </label>
                                    <button
                                      type="button"
                                      onClick={() => handleDeleteStore(store.id, store.name)}
                                      disabled={uploading}
                                      className={`px-3 py-2 rounded-xl text-center text-[10px] font-black uppercase border bg-red-50 text-red-600 border-red-100 hover:bg-red-100 ${uploading ? 'opacity-50 cursor-not-allowed' : 'cursor-pointer'}`}
                                    >
                                      Delete Store
                                    </button>
                                  </div>
                                )}
                              </div>
                            ))}
                          </div>
                        )}
                      </div>

                      <div className="space-y-6">
                        {(authRole === 'Admin' || authRole === 'Manager') && (
                          <div className="bg-white border border-gray-200 rounded-2xl p-6 shadow-xs">
                            <div className="flex items-center justify-between mb-4">
                              <div>
                                <h2 className="text-lg font-black text-gray-900 uppercase">Post New Task</h2>
                                <p className="text-xs text-gray-400 font-bold mt-1">Create a task without opening the full task board.</p>
                              </div>
                            </div>
                            <form onSubmit={handleTaskSubmit} className="space-y-3">
                              <input
                                type="text"
                                value={taskForm.title}
                                onChange={(e) => setTaskForm({ ...taskForm, title: e.target.value })}
                                placeholder="Task title..."
                                className="w-full px-3 py-2 text-sm border rounded-xl bg-gray-50 text-gray-900 outline-none"
                              />
                              <textarea
                                value={taskForm.description}
                                onChange={(e) => setTaskForm({ ...taskForm, description: e.target.value })}
                                placeholder="Task details..."
                                rows="3"
                                className="w-full px-3 py-2 text-sm border rounded-xl bg-gray-50 text-gray-900 outline-none resize-none"
                              />
                              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                                <select
                                  value={taskForm.priority}
                                  onChange={(e) => setTaskForm({ ...taskForm, priority: e.target.value })}
                                  className="px-3 py-2 text-xs border rounded-xl bg-white text-gray-900 font-bold"
                                >
                                  <option value="Low">Low</option>
                                  <option value="Normal">Normal</option>
                                  <option value="High">High</option>
                                  <option value="Urgent">Urgent</option>
                                </select>
                                <select
                                  value={taskForm.assigned_role}
                                  onChange={(e) => setTaskForm({ ...taskForm, assigned_role: e.target.value })}
                                  className="px-3 py-2 text-xs border rounded-xl bg-white text-gray-900 font-bold"
                                >
                                  <option value="All">All Staff</option>
                                  <option value="Operator">Operators</option>
                                  <option value="Photographer">Photographers</option>
                                  <option value="Content Editor">Content Editors</option>
                                  {userRegistry.map(user => (
                                    <option key={user.id} value={user.username}>{user.username}</option>
                                  ))}
                                </select>
                                <input
                                  type="datetime-local"
                                  value={taskForm.due_at}
                                  onChange={(e) => setTaskForm({ ...taskForm, due_at: e.target.value })}
                                  className="px-3 py-2 text-xs border rounded-xl bg-white text-gray-900 font-bold"
                                />
                              </div>
                              <div className="flex gap-2">
                                <button type="submit" disabled={taskSaving} className="flex-1 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-black uppercase">
                                  {taskSaving ? 'Saving...' : editingTaskId ? 'Update Task' : 'Post Task'}
                                </button>
                                {editingTaskId && (
                                  <button type="button" onClick={resetTaskForm} className="px-4 py-2.5 bg-gray-100 text-gray-700 rounded-xl text-xs font-black uppercase">
                                    Cancel
                                  </button>
                                )}
                              </div>
                            </form>
                          </div>
                        )}

                        <div className="bg-white border border-gray-200 rounded-2xl p-6 shadow-xs">
                          <div className="flex items-center justify-between mb-5">
                            <div>
                              <h2 className="text-lg font-black text-gray-900 uppercase">Task Board</h2>
                              <p className="text-xs text-gray-400 font-bold mt-1">{authRole === 'Operator' ? 'Your active instructions.' : 'Latest active instructions.'}</p>
                            </div>
                            <button onClick={() => setActiveTab('task_board')} className="px-3 py-2 bg-indigo-600 text-white rounded-xl text-xs font-black uppercase">Open Large</button>
                          </div>
                          <div className="space-y-3 max-h-[430px] overflow-y-auto pr-1">
                            {visibleTasks.filter(task => task.status !== 'Done' && task.status !== 'Archived').length === 0 ? (
                              <div className="text-center py-10 bg-gray-50 border rounded-xl text-xs font-bold text-gray-400 uppercase">No active tasks</div>
                            ) : (
                              visibleTasks.filter(task => task.status !== 'Done' && task.status !== 'Archived').slice(0, 8).map(task => (
                                <button key={task.id} onClick={() => setActiveTab('task_board')} className="w-full text-left p-4 rounded-xl border border-gray-200 bg-gray-50 hover:bg-white hover:border-indigo-300 transition-all">
                                  <div className="flex items-center justify-between gap-2">
                                    <span className="text-sm font-black text-gray-900 truncate">{task.title}</span>
                                    <span className={`text-[9px] px-2 py-1 rounded border font-black uppercase ${getPriorityClass(task.priority)}`}>{task.priority}</span>
                                  </div>
                                  <p className="text-xs text-gray-500 font-semibold mt-2 line-clamp-2">{task.description || 'No description provided.'}</p>
                                  <div className="text-[10px] text-gray-400 font-bold mt-2">Due: {formatDisplayDateTime(task.due_at)}</div>
                                </button>
                              ))
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                )}
                
                {/* ADMIN Tab Content */}
                {activeTab === 'admin_panel' && authRole === 'Admin' && (
                  <div className="space-y-8">
                    <div className="bg-gradient-to-r from-[rgba(138,21,56,0.85)] via-[rgba(138,21,56,1)] to-[rgba(138,21,56,0.85)] p-8 rounded-3xl border border-[rgba(138,21,56,0.95)] text-white shadow-xl relative overflow-hidden">
                      <div className="absolute right-0 bottom-0 translate-x-1/4 translate-y-1/4 w-96 h-96 bg-[#8a1538]/10 rounded-full blur-3xl"></div>
                      <div className="relative z-10 flex flex-col md:flex-row md:items-center md:justify-between gap-6">
                        <div>
                          <span className="text-xxs uppercase tracking-widest font-black text-[#8a1538] bg-[#8a1538]/40 border border-[#8a1538]/50 px-2.5 py-1 rounded-md">Control Center Settings</span>
                          <h1 className="text-3xl font-black mt-3 tracking-tight">System Administration Hub</h1>
                          <p className="text-sm text-slate-300 font-medium mt-1">Deploy custom permission rules, add new team accounts, and supervise global workflows seamlessly.</p>
                        </div>
                        <button 
                          onClick={() => setShowRoleModal(true)}
                          className="shrink-0 bg-[rgba(138,21,56,0.85)] hover:bg-[rgba(138,21,56,0.85)] text-white font-black text-xs uppercase tracking-wider px-6 py-3.5 rounded-xl shadow-lg transition-all border border-[#8a1538] cursor-pointer"
                        >
                          ⚡ Create Custom Role Matrix
                        </button>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                      <div className="bg-white border border-gray-200 rounded-2xl p-6 shadow-xs flex flex-col justify-between">
                        <div>
                          <h3 className="text-sm font-black uppercase tracking-wider text-gray-900 mb-1">New Staff Registration</h3>
                          <p className="text-xxs font-semibold text-gray-400 mb-4">Issue new employees their logged portal access tokens.</p>
                          
                          <form onSubmit={handleRegisterStaffAccount} className="space-y-4">
                            <div>
                              <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Assign Username</label>
                              <input 
                                type="text" required value={regUsername} onChange={(e) => setRegUsername(e.target.value)} placeholder="e.g. hussam_rose" 
                                className="w-full px-3 py-2 border rounded-xl text-xs bg-gray-50 outline-none text-gray-900 font-medium focus:bg-white"
                              />
                            </div>
                            <div>
                              <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Assign Password</label>
                              <input 
                                type="password" required value={regPassword} onChange={(e) => setRegPassword(e.target.value)} placeholder="••••••••" 
                                className="w-full px-3 py-2 border rounded-xl text-xs bg-gray-50 outline-none text-gray-900 font-medium focus:bg-white"
                              />
                            </div>
                            <div>
                              <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Assign Access Profile</label>
                              <select value={regRole} onChange={(e) => setRegRole(e.target.value)} className="w-full px-3 py-2 border rounded-xl text-xs bg-white text-gray-800 outline-none font-bold uppercase">
                                <option value="Operator">Operator (Floor Operations)</option>
                                <option value="Manager">Manager (Supervision Scope)</option>
                                <option value="Photographer">Photographer (Bulk RAW Stream)</option>
                                <option value="Content Editor">Content Editor (Differential Workbook Match)</option>
                                {customRoles.map((cr, idx) => (
                                  <option key={idx} value={cr.roleName}>{cr.roleName}</option>
                                ))}
                              </select>
                            </div>
                            <button type="submit" className="w-full py-2.5 bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] text-white text-xs font-black uppercase tracking-wider rounded-xl cursor-pointer">
                              Authorize & Register Staff Account
                            </button>
                          </form>
                        </div>
                      </div>

                      <div className="lg:col-span-2 bg-white border border-gray-200 rounded-2xl p-6 shadow-xs">
                        <h3 className="text-sm font-black uppercase tracking-wider text-gray-900 mb-1">Current Staff Workspace Registry ({userRegistry.length})</h3>
                        <p className="text-xxs font-semibold text-gray-400 mb-4">Revoke running access tokens inside running servers.</p>
                        
                        <div className="overflow-y-auto max-h-[340px] border rounded-xl divide-y">
                          {userRegistry.map(user => (
                            <div key={user.id} className="p-4 flex items-center justify-between hover:bg-gray-50/50 transition-colors">
                              <div className="flex items-center gap-3">
                                <div className="w-8 h-8 rounded-full bg-[rgba(138,21,56,0.85)] text-white flex items-center justify-center font-bold text-xs">👤</div>
                                <div>
                                  <div className="text-xs font-bold text-gray-900">{user.username}</div>
                                  <div className="text-[10px] text-gray-400 mt-0.5">Created: {new Date(user.created_at).toLocaleDateString()}</div>
                                </div>
                              </div>
                              <div className="flex items-center gap-4">
                                <span className="text-[9px] font-black tracking-widest uppercase bg-[rgba(138,21,56,0.06)] text-[#8a1538] border border-[rgba(138,21,56,0.18)] px-2.5 py-1 rounded">
                                  {user.role}
                                </span>
                                <button 
                                  onClick={() => handleRevokeStaffAccess(user.id, user.username)}
                                  className="text-xxs font-black text-red-600 border border-red-100 hover:bg-red-50 px-3 py-1 rounded-lg uppercase tracking-wider cursor-pointer"
                                >
                                  Revoke Access
                                </button>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>

                    {/* Permissions Matrix */}
                    <div className="bg-white border border-gray-200 rounded-2xl p-6 shadow-xs">
                      <div className="mb-4">
                        <h2 className="text-sm font-black text-gray-900 uppercase tracking-tight">Access Permissions Matrix Logs</h2>
                        <p className="text-xs text-gray-400 font-medium">RBAC Security parameters configured inside system components.</p>
                      </div>

                      <div className="overflow-x-auto border rounded-xl">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="bg-gray-50 border-b text-xxs uppercase tracking-wider font-black text-gray-400">
                              <th className="p-4">Assigned Department / Group</th>
                              <th className="p-4 text-center">Upload Assets Array</th>
                              <th className="p-4 text-center">Modify Layout Sheets</th>
                              <th className="p-4 text-center">Review Validation Lists</th>
                              <th className="p-4 text-center">Supervise Operations</th>
                              <th className="p-4 text-center">Configuration Actions</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y text-xs font-semibold text-gray-700">
                            <tr className="bg-slate-50/50">
                              <td className="p-4 font-bold text-gray-900">⚙️ System Manager (Root Core)</td>
                              <td className="p-4 text-center text-emerald-600 font-black">✔️ ENABLED</td>
                              <td className="p-4 text-center text-emerald-600 font-black">✔️ ENABLED</td>
                              <td className="p-4 text-center text-emerald-600 font-black">✔️ ENABLED</td>
                              <td className="p-4 text-center text-emerald-600 font-black">✔️ ENABLED</td>
                              <td className="p-4 text-center tracking-widest uppercase text-xxs font-black text-gray-400">CORE PROTECTED</td>
                            </tr>
                            <tr className="bg-slate-50/50">
                              <td className="p-4 font-bold text-gray-900">👤 Production Operator</td>
                              <td className="p-4 text-center text-emerald-600 font-black">✔️ ENABLED</td>
                              <td className="p-4 text-center text-red-500 font-black">❌ DISABLED</td>
                              <td className="p-4 text-center text-red-500 font-black">❌ DISABLED</td>
                              <td className="p-4 text-center text-red-500 font-black">❌ DISABLED</td>
                              <td className="p-4 text-center tracking-widest uppercase text-xxs font-black text-gray-400">CORE PROTECTED</td>
                            </tr>
                            {customRoles.map((role, i) => (
                              <tr key={i} className="hover:bg-gray-50/40 transition-colors">
                                <td className="p-4 font-black text-[rgba(138,21,56,0.85)]">🎨 {role.roleName}</td>
                                <td className="p-4 text-center font-bold">
                                  <span className={`px-2 py-0.5 rounded text-[10px] ${role.permissions.canUploadAssets ? 'bg-green-50 text-green-700 border border-green-200' : 'bg-red-50 text-red-600'}`}>
                                    {role.permissions.canUploadAssets ? 'TRUE' : 'FALSE'}
                                  </span>
                                </td>
                                <td className="p-4 text-center font-bold">
                                  <span className={`px-2 py-0.5 rounded text-[10px] ${role.permissions.canModifyDataSheets ? 'bg-green-50 text-green-700 border border-green-200' : 'bg-red-50 text-red-600'}`}>
                                    {role.permissions.canModifyDataSheets ? 'TRUE' : 'FALSE'}
                                  </span>
                                </td>
                                <td className="p-4 text-center font-bold">
                                  <span className={`px-2 py-0.5 rounded text-[10px] ${role.permissions.canReviewArrays ? 'bg-green-50 text-green-700 border border-green-200' : 'bg-red-50 text-red-600'}`}>
                                    {role.permissions.canReviewArrays ? 'TRUE' : 'FALSE'}
                                  </span>
                                </td>
                                <td className="p-4 text-center font-bold">
                                  <span className={`px-2 py-0.5 rounded text-[10px] ${role.permissions.canSuperviseStaff ? 'bg-green-50 text-green-700 border border-green-200' : 'bg-red-50 text-red-600'}`}>
                                    {role.permissions.canSuperviseStaff ? 'TRUE' : 'FALSE'}
                                  </span>
                                </td>
                                <td className="p-4 text-center">
                                  <button 
                                    onClick={() => handleDeleteCustomRole(role.roleName)}
                                    className="px-2.5 py-1 text-red-600 border border-red-100 hover:bg-red-50 rounded uppercase tracking-wider text-xxs font-black transition-colors cursor-pointer"
                                  >
                                    Drop Role
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}

                {/* TASK BOARD LARGE VIEW */}
                {activeTab === 'task_board' && (
                  <div className="space-y-8">
                    <div className="bg-gradient-to-r from-[rgba(138,21,56,0.85)] via-[rgba(138,21,56,1)] to-[rgba(138,21,56,0.85)] p-8 rounded-3xl border border-[#8a1538] text-white shadow-xl relative overflow-hidden">
                      <div className="relative z-10 flex flex-col md:flex-row md:items-center md:justify-between gap-6">
                        <div>
                          <span className="text-xxs uppercase tracking-widest font-black text-indigo-300 bg-indigo-900/40 border border-indigo-800/50 px-2.5 py-1 rounded-md">Team Dispatch Board</span>
                          <h1 className="text-3xl font-black mt-3 tracking-tight">Task Board</h1>
                          <p className="text-sm text-slate-300 font-medium mt-1">Managers can post assignments. Employees can view the work queue from the side panel or this large view.</p>
                        </div>
                        <div className="text-right">
                          <div className="text-3xl font-black">{openTaskCount}</div>
                          <div className="text-[10px] font-black uppercase tracking-widest text-indigo-200">Active Tasks</div>
                        </div>
                      </div>
                    </div>

                    {taskError && (
                      <div className="p-4 bg-red-50 border border-red-200 text-red-700 rounded-xl text-xs font-bold">
                        Task board warning: {taskError}. Run the task_board SQL setup if this is the first install.
                      </div>
                    )}

                    {isTaskManager && (
                      <div className="bg-white border border-gray-200 rounded-2xl p-6 shadow-xs">
                        <h2 className="text-sm font-black text-gray-900 uppercase tracking-wider mb-1">
                          {editingTaskId ? 'Edit Task' : 'Post New Task'}
                        </h2>
                        <p className="text-xs text-gray-400 font-semibold mb-5">Tasks posted here are visible to employees immediately after save.</p>
                        <form onSubmit={handleTaskSubmit} className="grid grid-cols-1 lg:grid-cols-12 gap-4">
                          <div className="lg:col-span-4">
                            <label className="block text-[10px] font-black uppercase text-gray-400 mb-1">Task Title *</label>
                            <input
                              type="text"
                              value={taskForm.title}
                              onChange={(e) => setTaskForm({ ...taskForm, title: e.target.value })}
                              placeholder="e.g. Recheck rejected product images"
                              className="w-full px-3 py-2 border rounded-xl text-xs bg-gray-50 text-gray-900 outline-none"
                            />
                          </div>
                          <div className="lg:col-span-2">
                            <label className="block text-[10px] font-black uppercase text-gray-400 mb-1">Priority</label>
                            <select
                              value={taskForm.priority}
                              onChange={(e) => setTaskForm({ ...taskForm, priority: e.target.value })}
                              className="w-full px-3 py-2 border rounded-xl text-xs bg-white text-gray-900 outline-none font-bold"
                            >
                              <option value="Low">Low</option>
                              <option value="Normal">Normal</option>
                              <option value="High">High</option>
                              <option value="Urgent">Urgent</option>
                            </select>
                          </div>
                          <div className="lg:col-span-2">
                            <label className="block text-[10px] font-black uppercase text-gray-400 mb-1">Visible To</label>
                            <select
                              value={taskForm.assigned_role}
                              onChange={(e) => setTaskForm({ ...taskForm, assigned_role: e.target.value })}
                              className="w-full px-3 py-2 border rounded-xl text-xs bg-white text-gray-900 outline-none font-bold"
                            >
                              <option value="All">All Staff</option>
                              <option value="Operator">Operators</option>
                              <option value="Photographer">Photographers</option>
                              <option value="Content Editor">Content Editors</option>
                              <option value="Manager">Managers</option>
                              {userRegistry.map(user => (
                                <option key={user.id} value={user.username}>{user.username}</option>
                              ))}
                            </select>
                          </div>
                          <div className="lg:col-span-2">
                            <label className="block text-[10px] font-black uppercase text-gray-400 mb-1">Due Date/Time</label>
                            <input
                              type="datetime-local"
                              value={taskForm.due_at}
                              onChange={(e) => setTaskForm({ ...taskForm, due_at: e.target.value })}
                              className="w-full px-3 py-2 border rounded-xl text-xs bg-gray-50 text-gray-900 outline-none"
                            />
                          </div>
                          <div className="lg:col-span-2 flex items-end gap-2">
                            <button
                              type="submit"
                              disabled={taskSaving}
                              className="flex-1 px-4 py-2 bg-indigo-700 hover:bg-indigo-800 text-white rounded-xl text-xs font-black uppercase tracking-wider disabled:opacity-60"
                            >
                              {taskSaving ? 'Saving...' : editingTaskId ? 'Update' : 'Post'}
                            </button>
                            {editingTaskId && (
                              <button type="button" onClick={resetTaskForm} className="px-4 py-2 bg-gray-100 text-gray-700 rounded-xl text-xs font-black uppercase">
                                Cancel
                              </button>
                            )}
                          </div>
                          <div className="lg:col-span-12">
                            <label className="block text-[10px] font-black uppercase text-gray-400 mb-1">Task Details</label>
                            <textarea
                              rows="3"
                              value={taskForm.description}
                              onChange={(e) => setTaskForm({ ...taskForm, description: e.target.value })}
                              placeholder="Write clear instructions for the employee team..."
                              className="w-full px-3 py-2 border rounded-xl text-xs bg-gray-50 text-gray-900 outline-none"
                            />
                          </div>
                        </form>
                      </div>
                    )}

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                      {visibleTasks.length === 0 ? (
                        <div className="lg:col-span-2 text-center text-xs text-gray-400 font-bold uppercase py-16 bg-white border rounded-2xl">
                          No tasks available for your role.
                        </div>
                      ) : (
                        visibleTasks.map(task => (
                          <div key={task.id} className="bg-white border border-gray-200 rounded-2xl p-5 shadow-xs hover:shadow-sm transition-all">
                            <div className="flex items-start justify-between gap-4 mb-3">
                              <div>
                                <h3 className="text-base font-black text-gray-900">{task.title}</h3>
                                <div className="text-[10px] font-bold text-gray-400 mt-1 uppercase tracking-wider">
                                  Posted by {task.created_by || 'Manager'} • {formatDisplayDateTime(task.created_at)}
                                </div>
                              </div>
                              <div className="flex flex-col items-end gap-1">
                                <span className={`px-2.5 py-1 rounded-lg border text-[10px] font-black uppercase ${getPriorityClass(task.priority)}`}>{task.priority || 'Normal'}</span>
                                <span className={`px-2.5 py-1 rounded-lg border text-[10px] font-black uppercase ${getTaskStatusClass(task.status)}`}>{task.status || 'Open'}</span>
                              </div>
                            </div>
                            {task.description && <p className="text-sm text-gray-600 font-medium whitespace-pre-wrap leading-relaxed mb-4">{task.description}</p>}
                            <div className="grid grid-cols-2 gap-3 text-[11px] font-bold text-gray-500 mb-4">
                              <div className="bg-gray-50 border rounded-xl p-3"><span className="block text-gray-400 uppercase text-[9px]">Visible To</span>{task.assigned_role || 'All'}</div>
                              <div className="bg-gray-50 border rounded-xl p-3"><span className="block text-gray-400 uppercase text-[9px]">Due</span>{formatDisplayDateTime(task.due_at)}</div>
                            </div>
                            <div className="flex flex-wrap gap-2">
                              <select
                                value={task.status || 'Open'}
                                onChange={(e) => handleUpdateTaskStatus(task.id, e.target.value)}
                                className="px-3 py-2 border rounded-xl text-xs font-bold bg-white text-gray-800"
                              >
                                <option value="Open">Open</option>
                                <option value="In Progress">In Progress</option>
                                <option value="Done">Done</option>
                                {isTaskManager && <option value="Archived">Archived</option>}
                              </select>
                              {isTaskManager && (
                                <>
                                  <button onClick={() => handleStartEditTask(task)} className="px-3 py-2 bg-[rgba(138,21,56,0.06)] text-[#8a1538] border border-[rgba(138,21,56,0.18)] rounded-xl text-xs font-black uppercase">Edit</button>
                                  <button onClick={() => handleDeleteTask(task.id)} className="px-3 py-2 bg-red-50 text-red-700 border border-red-100 rounded-xl text-xs font-black uppercase">Delete</button>
                                </>
                              )}
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                )}

                {/* DASHBOARD CORE HUB */}
                {activeTab === 'dashboard' && (
                  <>
                    <div className="mb-8 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                      <div>
                        <h1 className="text-3xl font-bold text-gray-900 tracking-tight">
                          {isStoreScoped ? `${getStoreNameById(selectedStoreId)} Store Dashboard` : 'BlackRose Operations Dashboard'}
                        </h1>
                        <p className="text-xs text-gray-400 mt-1">
                          {selectedHistoryScope ? `Showing isolated rows for workbook context: [ ${selectedHistoryScope} ]` : isStoreScoped ? 'Displaying products for the selected store only.' : "Displaying merged inventory catalog datasets."}
                        </p>
                        <p className="text-[15px] text-gray-500 font-semibold mt-1">
                          {dashboardClock.toLocaleDateString([], { weekday: 'short', year: 'numeric', month: 'short', day: '2-digit' })} • {dashboardClock.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                        </p>
                      </div>
                      <div className="flex gap-2">
                        {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('modify_sheets')) && (
                          <button 
                            onClick={() => setShowAdHocModal(true)}
                            className="bg-[#8a1538] text-white font-black text-xs px-4 py-2 rounded-xl shadow hover:bg-[rgba(138,21,56,0.95)] transition uppercase"
                          >
                            + Add Ad-Hoc Entry
                          </button>
                        )}
                        {selectedHistoryScope && (
                          <button 
                            onClick={() => setSelectedHistoryScope(null)}
                            className="bg-red-50 text-red-600 border border-red-200 font-bold text-xs px-3 py-1.5 rounded-xl hover:bg-red-100 cursor-pointer"
                          >
                            ✕ Clear File Filter
                          </button>
                        )}
                        <span className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-bold border tracking-wider uppercase bg-white ${realtimeStatus === 'Live Connected' ? 'text-green-700 border-green-200' : 'text-red-700 border-red-200'}`}>
                          <span className={`w-2 h-2 rounded-full ${realtimeStatus === 'Live Connected' ? 'bg-green-500 animate-pulse' : 'bg-green-500'}`}></span>
                          {realtimeStatus}
                        </span>
                      </div>
                    </div>

                    {/* DYNAMIC ACTIVE SPREADSHEETS ROW DIRECT CONTROLS */}
                    {(authRole === 'Admin' || authRole === 'Manager') && uniqueUploadedSheets.length > 0 && (
                      <div className="bg-white p-5 rounded-2xl border border-gray-200 shadow-xs mb-8">
                        <h3 className="text-xs font-black uppercase text-gray-400 tracking-wider mb-3">🛠️ Active Spreadsheet Direct Controls</h3>
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          {uniqueUploadedSheets.map((sheetName, sIdx) => (
                            <div key={sIdx} className="p-3 border border-gray-100 rounded-xl bg-gray-50/50 flex flex-col sm:flex-row sm:items-center justify-between gap-3 hover:border-[rgba(138,21,56,0.28)] transition-colors">
                              <div className="truncate">
                                <span className="text-xs font-bold text-gray-800 block truncate">📄 {sheetName}</span>
                                <span className="text-[10px] text-gray-400 block font-medium mt-0.5">Linked products: {products.filter(p => p.sheet_reference === sheetName).length} units</span>
                              </div>
                              <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                                {hasOriginalArchive(sheetName) ? (
                                  <button 
                                    onClick={() => handleDownloadDashboardManifestByName(sheetName, 'original')}
                                    className="px-2.5 py-1.5 bg-white hover:bg-gray-100 text-gray-700 border rounded-lg text-[10px] font-black uppercase tracking-wider transition-colors cursor-pointer shadow-sm"
                                  >
                                    Original Sheet
                                  </button>
                                ) : (
                                  <button 
                                    disabled
                                    title="Original Excel archive is missing. Re-upload this file to create the original archive."
                                    className="px-2.5 py-1.5 bg-gray-100 text-gray-400 border rounded-lg text-[10px] font-black uppercase tracking-wider cursor-not-allowed shadow-sm"
                                  >
                                    Original Missing
                                  </button>
                                )}
                                <button 
                                  onClick={() => handleDownloadDashboardManifestByName(sheetName, 'live')}
                                  className="px-2.5 py-1.5 bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] text-white border border-[#8a1538] rounded-lg text-[10px] font-black uppercase tracking-wider transition-colors cursor-pointer shadow-sm"
                                >
                                  Live Matrix
                                </button>
                                <button 
                                  onClick={() => handleDownloadMissingByName(sheetName)}
                                  className="px-2.5 py-1.5 bg-amber-500 hover:bg-amber-600 text-white border border-amber-600 rounded-lg text-[10px] font-black uppercase tracking-wider transition-colors cursor-pointer shadow-sm"
                                >
                                  Download Missing
                                </button>
                                <button 
                                  onClick={() => handleDashboardPurgeSheetByName(sheetName)}
                                  className="px-2.5 py-1.5 bg-red-50 hover:bg-red-100 text-red-600 border border-red-100 rounded-lg text-[10px] font-black uppercase tracking-wider transition-colors cursor-pointer"
                                >
                                  🗑️ Delete
                                </button>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}

                    <div className="grid grid-cols-2 lg:grid-cols-7 gap-4 mb-8">
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs">
                        <div className="text-xs font-bold text-gray-400 uppercase tracking-wider">Catalog Scope</div>
                        <div className="text-2xl font-black text-gray-900 mt-1">{metrics.total}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-amber-500">
                        <div className="text-xs font-bold text-amber-600 uppercase tracking-wider">Missing</div>
                        <div className="text-2xl font-black text-amber-700 mt-1">{metrics.missing}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-blue-500">
                        <div className="text-xs font-bold text-[rgba(138,21,56,0.85)] uppercase tracking-wider">In Progress</div>
                        <div className="text-2xl font-black text-[#8a1538] mt-1">{metrics.processing}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-green-500">
                        <div className="text-xs font-bold text-green-600 uppercase tracking-wider">Completed</div>
                        <div className="text-2xl font-black text-green-700 mt-1">{metrics.completed}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-purple-500">
                        <div className="text-xs font-bold text-[rgba(138,21,56,0.85)] uppercase tracking-wider">Modified</div>
                        <div className="text-2xl font-black text-[#8a1538] mt-1">{metrics.modified}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-red-500">
                        <div className="text-xs font-bold text-red-600 uppercase tracking-wider">Rejected</div>
                        <div className="text-2xl font-black text-red-700 mt-1">{metrics.rejected}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-orange-500">
                        <div className="text-xs font-bold text-orange-600 uppercase tracking-wider">Low Stock</div>
                        <div className="text-2xl font-black text-orange-700 mt-1">{metrics.lowStock}</div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
                      {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('modify_sheets')) && (
                        <div className="lg:col-span-2 bg-white p-6 rounded-xl shadow-sm border border-gray-200 flex flex-col justify-between">
                          <div>
                            <div className="flex items-center justify-between mb-3">
                              <h2 className="text-sm font-bold uppercase tracking-wider text-gray-500">Import Master Excel Sheets</h2>
                            </div>
                            <div className="border-2 border-dashed border-gray-300 rounded-lg p-7 text-center bg-gray-50/50 hover:bg-gray-50 transition-all relative">
                              <div className="space-y-2">
                                <div className="text-3xl">📊</div>
                                <div className="text-sm text-gray-600 font-medium">
                                  {uploading ? "Analyzing updates..." : isStoreScoped ? `Upload sheet for ${getStoreNameById(selectedStoreId)}` : "Select a store first, then upload its sheet"}
                                </div>
                                <div className="text-[11px] text-gray-400 font-bold">
                                  Current upload target: {isStoreScoped ? getStoreNameById(selectedStoreId) : 'No store selected'}
                                </div>
                              </div>
                              <input type="file" accept=".xlsx, .xls, .csv" onChange={handleExcelUpload} disabled={uploading} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                            </div>
                          </div>
                        </div>
                      )}

                      {(authRole === 'Admin' || authRole === 'Manager' || authRole === 'Photographer' || checkPermission('upload_assets')) && (
                        <div className="bg-white p-6 rounded-xl shadow-sm border border-[rgba(138,21,56,0.28)] flex flex-col justify-between bg-gradient-to-br from-blue-50/40 via-white to-white">
                          <div>
                            <h3 className="text-sm font-black text-[rgba(138,21,56,0.85)] uppercase tracking-wide flex items-center gap-2">
                              <span>🗂️</span> Bulk Image Asset Upload
                            </h3>
                            <p className="text-xs text-[#8a1538] font-medium mt-1">Supported ZIP formats: <strong>[SKU]/RAW/img.jpg</strong>, <strong>[SKU]/EDITED/img.jpg</strong>, or a ZIP named <strong>[SKU].zip</strong> containing <strong>RAW/img.jpg</strong>.</p>
                            <p className={`text-[11px] font-bold mt-2 ${isStoreScoped ? 'text-emerald-700' : 'text-red-600'}`}>
                              Upload target: {isStoreScoped ? getStoreNameById(selectedStoreId) : 'Select one store first'}
                            </p>
                          </div>
                          <div className="mt-4 flex flex-col gap-2">
                            <label className={`w-full py-3 rounded-xl text-xs font-bold uppercase tracking-wider shadow-xs transition-colors block text-center ${isStoreScoped && !uploading ? 'bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] text-white cursor-pointer' : 'bg-gray-200 text-gray-500 cursor-not-allowed'}`}>
                              {uploading ? "Extracting ZIP tree..." : isStoreScoped ? "Upload Assets ZIP" : "Select Store To Upload ZIP"}
                              <input type="file" accept=".zip" onChange={handleBulkZipUpload} disabled={uploading || !isStoreScoped} className="hidden" />
                            </label>
                            
                            {/* RESTORED BULK ASSETS EXPORT BUTTON */}
                            {(authRole === 'Admin' || authRole === 'Manager') && (
                              <button 
                                onClick={handleBulkDownloadAllAssets} disabled={uploading}
                                className="w-full bg-white hover:bg-gray-50 text-[rgba(138,21,56,0.95)] border border-[rgba(138,21,56,0.28)] py-3 rounded-xl text-xs font-bold uppercase tracking-wider cursor-pointer shadow-xs transition-colors"
                              >
                                {uploading ? "Processing..." : "⬇️ Download All DB Images (ZIP)"}
                              </button>
                            )}
                          </div>
                        </div>
                      )}
                    </div>

                    <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-200 mb-6 flex flex-col md:flex-row items-center gap-4 justify-between">
                      <input
                        type="text" placeholder="Filter operational catalog items..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)}
                        className="w-full md:max-w-md px-4 py-2 border border-gray-300 rounded-lg text-sm bg-gray-50 outline-none text-gray-900"
                      />
                      <div className="flex flex-wrap items-center gap-2">
                        {['All', 'Missing', 'Processing', 'Completed', 'Modified', 'Rejected'].map((status) => (
                          <button
                            key={status} onClick={() => setStatusFilter(status)}
                            className={`px-3 py-1.5 rounded-lg text-xs font-bold uppercase transition-all border ${
                              statusFilter === status ? 'bg-[rgba(138,21,56,0.85)] text-white border-[rgba(138,21,56,0.85)]' : 'bg-gray-50 text-gray-600 border-gray-200 hover:bg-gray-100'
                            }`}
                          >
                            {status}
                          </button>
                        ))}
                      </div>
                    </div>

                    {authRole === 'Operator' ? (
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-6">
                        {filteredProducts.map((prod) => {
                          const isProcessing = prod.status === 'Processing';
                          const isRejected = prod.status === 'Rejected';
                          const statusColors = prod.status === 'Completed' ? 'bg-green-100 text-green-700 border-green-200' : prod.status === 'Modified' ? 'bg-[rgba(138,21,56,0.10)] text-[#8a1538] border-[rgba(138,21,56,0.28)]' : prod.status === 'Rejected' ? 'bg-red-100 text-red-700 border-red-200' : prod.status === 'Processing' ? 'bg-[rgba(138,21,56,0.10)] text-[#8a1538] border-[rgba(138,21,56,0.28)]' : 'bg-amber-100 text-amber-700 border-amber-200';
                          return (
                            <div 
                              key={prod.id} onClick={() => handleCardInteraction(prod)}
                              className={`bg-white rounded-2xl border shadow-sm p-6 flex flex-col h-full transition-all hover:shadow-md hover:-translate-y-0.5 cursor-pointer group ${isRejected ? 'border-red-300 bg-red-50/20' : isProcessing ? 'border-[rgba(138,21,56,0.28)] bg-[rgba(138,21,56,0.06)]/10' : 'border-gray-200'}`}
                            >
                              <div className="flex-grow">
                                <div className="flex items-center justify-between mb-4">
                                  <span className={`text-[10px] font-bold uppercase px-2.5 py-1 rounded-md border ${statusColors}`}>{prod.status || 'Missing'}</span>
                               </div>
                                <h3 className="font-bold text-base text-gray-900 line-clamp-2 mb-4 group-hover:text-[rgba(138,21,56,0.85)]">{prod.product_name}</h3>
                                <div className="grid grid-cols-2 gap-4">
                                  <div>
                                    <p className="text-[10px] text-gray-400 font-bold uppercase mb-1">SKU</p>
                                    <p className="text-xs font-mono font-medium text-gray-600 bg-gray-50 px-2 py-0.5 rounded border inline-block">{prod.sku}</p>
                                  </div>
                                  <div>
                                    <p className="text-[10px] text-gray-400 font-bold uppercase mb-1">Location</p>
                                    <p className="text-xs font-medium text-gray-600 truncate">{prod.warehouse}</p>
                                  </div>
                                </div>

                                <div className="mt-4 grid grid-cols-2 gap-2" onClick={(e) => e.stopPropagation()}>
                                  <button
                                    disabled={getArray(prod.raw_image_url).length === 0 || uploading}
                                    onClick={() => handleDownloadProductAssets(prod, 'raw')}
                                    className={`px-3 py-2 rounded-lg text-[10px] font-black uppercase tracking-wider border ${getArray(prod.raw_image_url).length > 0 && !uploading ? 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50 cursor-pointer' : 'bg-gray-100 text-gray-400 border-gray-200 cursor-not-allowed'}`}
                                  >
                                    ⬇ RAW ({getArray(prod.raw_image_url).length})
                                  </button>
                                  <button
                                    disabled={getArray(prod.edited_image_url).length === 0 || uploading}
                                    onClick={() => handleDownloadProductAssets(prod, 'edited')}
                                    className={`px-3 py-2 rounded-lg text-[10px] font-black uppercase tracking-wider border ${getArray(prod.edited_image_url).length > 0 && !uploading ? 'bg-[rgba(138,21,56,0.06)] text-[#8a1538] border-[rgba(138,21,56,0.28)] hover:bg-[rgba(138,21,56,0.10)] cursor-pointer' : 'bg-gray-100 text-gray-400 border-gray-200 cursor-not-allowed'}`}
                                  >
                                    ⬇ EDITED ({getArray(prod.edited_image_url).length})
                                  </button>
                                </div>

                                {prod.status === 'Rejected' && (
                                  <div className="mt-4 p-3 rounded-lg bg-red-50 border border-red-200">
                                    <div className="text-[10px] font-black uppercase tracking-wider text-red-600 mb-1">
                                      Manager Note
                                    </div>
                                    <p className="text-xs font-semibold text-red-800 line-clamp-3 whitespace-pre-wrap">
                                      {prod.rejection_note || 'No rejection note was provided.'}
                                    </p>
                                  </div>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="bg-gray-100 border-b border-gray-200 text-xs uppercase text-gray-500 font-bold">
                                <th className="p-4">SKU</th>
                                <th className="p-4">Product Name</th>
                                <th className="p-4">Warehouse</th>
                                <th className="p-4 text-center">Stock</th>
                                <th className="p-4">Employee</th>
                                <th className="p-4 text-center">Images</th>
                                <th className="p-4">Status</th>
                                <th className="p-4 text-center">Actions</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-200 text-sm text-gray-700">
                              {filteredProducts.map((prod) => {
                                const isEditing = prod.id === editingId;
                                return (
                                  <tr key={prod.id} className={isEditing ? 'bg-[rgba(138,21,56,0.06)]/60' : 'hover:bg-gray-50/20'}>
                                    <td className="p-4 font-mono font-medium">{prod.sku}</td>
                                    <td className="p-4">
                                      {isEditing ? <input type="text" value={editForm.product_name} onChange={(e) => setEditForm({...editForm, product_name: e.target.value})} className="border rounded px-1.5 py-0.5 text-xs text-gray-900" /> : <span className="font-bold text-gray-900">{prod.product_name}</span>}
                                    </td>
                                    <td className="p-4 text-xs font-semibold text-gray-500">{prod.warehouse}</td>
                                    <td className="p-4 text-center font-bold">
                                      {isEditing ? <input type="number" value={editForm.stock_quantity} onChange={(e) => setEditForm({...editForm, stock_quantity: e.target.value})} className="border rounded px-1 text-xs w-14 text-gray-900" /> : prod.stock_quantity}
                                    </td>
                                    <td className="p-4 text-xs italic">{prod.processed_by || '—'}</td>
                                    <td className="p-4 text-center">
                                      <button onClick={() => setManagerPreview(prod)} className="text-[10px] font-bold px-4 py-2 bg-slate-100 rounded-lg border border-gray-200 cursor-pointer">Review</button>
                                    </td>
                                    <td className="p-4">
                                      <span className="px-3 py-1 rounded-full border text-xs font-bold">{prod.status}</span>
                                    </td>
                                    <td className="p-4 text-center">
                                      {isEditing ? (
                                        <div className="space-x-1 flex justify-center">
                                          <button onClick={() => handleSaveEdit(prod.id)} className="px-3 py-1 bg-green-600 text-white rounded text-xs cursor-pointer">Save</button>
                                          <button onClick={() => setEditingId(null)} className="px-3 py-1 bg-gray-200 text-gray-600 rounded text-xs cursor-pointer">Exit</button>
                                        </div>
                                      ) : (
                                        <div className="space-x-1 flex justify-center">
                                          {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('modify_sheets')) && (
                                            <button onClick={() => startEditing(prod)} className="px-3 py-1 border text-xs rounded text-gray-600 bg-white cursor-pointer hover:bg-gray-50">Edit</button>
                                          )}
                                        </div>
                                      )}
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </>
                )}

                {/* ACTIVE REGISTERED DIRECTORY TAB */}
                {activeTab === 'operators' && (authRole === 'Admin' || authRole === 'Manager' || checkPermission('supervise_staff')) && (
                  <div className="bg-white rounded-xl p-6 border border-gray-200 shadow-xs">
                    <h2 className="text-xl font-bold text-gray-900 mb-6">Registered Systems Staff Directory</h2>
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                      {userRegistry.filter(u => u.role === 'Operator').map(user => {
                        const stats = compileUserPerformanceMetrics(user.username);
                        return (
                          <div key={user.id} onClick={() => setSelectedOperatorStats({ username: user.username, performance: stats })} className="p-5 border border-gray-200 rounded-xl bg-gray-50 hover:bg-white hover:border-teal-500 transition-all cursor-pointer group">
                            <span className="font-bold text-sm text-gray-900 block">👤 {user.username}</span>
                            <div className="mt-4 text-xxs font-medium text-gray-400 uppercase tracking-wider">
                              Claimed: <span className="font-black text-gray-900 block text-xs">{stats.totalClaimed} units</span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* MY PERFORMANCE ANALYTICS */}
                {activeTab === 'my_performance' && canViewMyPerformance && selfPerformanceStats && (
                  <div className="space-y-8">
                    <div className="bg-gradient-to-r from-amber-700 via-orange-700 to-[rgba(138,21,56,0.85)] p-8 rounded-3xl border border-amber-800 text-white shadow-xl relative overflow-hidden">
                      <div className="absolute right-0 bottom-0 translate-x-1/4 translate-y-1/4 w-96 h-96 bg-white/10 rounded-full blur-3xl"></div>
                      <div className="relative z-10">
                        <span className="text-xxs uppercase tracking-widest font-black text-amber-100 bg-white/10 border border-white/20 px-2.5 py-1 rounded-md">Employee Performance</span>
                        <h1 className="text-3xl font-black mt-3 tracking-tight">My Performance Dashboard</h1>
                        <p className="text-sm text-amber-50 font-medium mt-1">Your completed work, rejection history, image counts, and time tracking summary.</p>
                      </div>
                    </div>

                    <div className="grid grid-cols-2 md:grid-cols-6 gap-4">
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Total Claimed</span>
                        <div className="text-xl font-black text-slate-800 mt-1">{selfPerformanceStats.totalClaimed}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Missing</span>
                        <div className="text-xl font-black text-amber-600 mt-1">{selfPerformanceStats.totalMissing}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Processing</span>
                        <div className="text-xl font-black text-[rgba(138,21,56,0.85)] mt-1">{selfPerformanceStats.totalProcessing}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Completed</span>
                        <div className="text-xl font-black text-green-600 mt-1">{selfPerformanceStats.totalCompleted}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Modified</span>
                        <div className="text-xl font-black text-[rgba(138,21,56,0.85)] mt-1">{selfPerformanceStats.totalModified}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Rejected</span>
                        <div className="text-xl font-black text-red-600 mt-1">{selfPerformanceStats.totalRejected}</div>
                      </div>
                    </div>

                    <div className="grid grid-cols-2 md:grid-cols-6 gap-4">
                      <div className="border p-4 rounded-xl bg-green-50 border-green-100 shadow-xs">
                        <span className="text-[9px] font-bold text-green-700 uppercase block">Completed This Week</span>
                        <div className="text-xl font-black text-green-700 mt-1">{selfPerformanceStats.weekCount}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-green-50 border-green-100 shadow-xs">
                        <span className="text-[9px] font-bold text-green-700 uppercase block">Completed This Month</span>
                        <div className="text-xl font-black text-green-700 mt-1">{selfPerformanceStats.monthCount}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-[rgba(138,21,56,0.06)] border-[rgba(138,21,56,0.18)] shadow-xs">
                        <span className="text-[9px] font-bold text-[#8a1538] uppercase block">Modified This Week</span>
                        <div className="text-xl font-black text-[#8a1538] mt-1">{selfPerformanceStats.modifiedWeekCount}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-red-50 border-red-100 shadow-xs">
                        <span className="text-[9px] font-bold text-red-700 uppercase block">Rejected This Week</span>
                        <div className="text-xl font-black text-red-700 mt-1">{selfPerformanceStats.rejectedWeekCount}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-[rgba(138,21,56,0.06)] border-[rgba(138,21,56,0.18)] shadow-xs">
                        <span className="text-[9px] font-bold text-[#8a1538] uppercase block">Modified This Month</span>
                        <div className="text-xl font-black text-[#8a1538] mt-1">{selfPerformanceStats.modifiedMonthCount}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-red-50 border-red-100 shadow-xs">
                        <span className="text-[9px] font-bold text-red-700 uppercase block">Rejected This Month</span>
                        <div className="text-xl font-black text-red-700 mt-1">{selfPerformanceStats.rejectedMonthCount}</div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-500 uppercase block">Average Time To Complete</span>
                        <div className="text-lg font-black text-slate-900 mt-1">{selfPerformanceStats.avgCompletedTime}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-red-50 border-red-100 shadow-xs">
                        <span className="text-[9px] font-bold text-red-700 uppercase block">Average Time Before Rejection</span>
                        <div className="text-lg font-black text-red-800 mt-1">{selfPerformanceStats.avgRejectedTime}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Fastest Completed Product</span>
                        <div className="text-lg font-black text-emerald-700 mt-1">{selfPerformanceStats.fastestCompletedTime}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Slowest Completed Product</span>
                        <div className="text-lg font-black text-orange-700 mt-1">{selfPerformanceStats.slowestCompletedTime}</div>
                      </div>
                    </div>

                    <div className="bg-white rounded-xl p-6 border border-gray-200 shadow-xs">
                      <div className="flex flex-col sm:flex-row sm:items-end sm:justify-between gap-3 mb-5">
                        <div>
                          <h2 className="text-xl font-black text-gray-900 uppercase tracking-tight">Done Product Time Log</h2>
                          <p className="text-xs text-gray-400 font-semibold mt-1">Completed, modified, and rejected products assigned to you, with the tracked time spent on each item.</p>
                        </div>
                        <span className="text-[10px] font-black uppercase tracking-wider bg-slate-100 text-slate-600 px-3 py-1 rounded-full border">
                          {selfPerformanceStats.doneProductLedger.length} done item(s)
                        </span>
                      </div>

                      {selfPerformanceStats.doneProductLedger.length === 0 ? (
                        <div className="text-center text-xs text-gray-400 font-bold uppercase py-12 border rounded-xl bg-gray-50">
                          No completed, modified, or rejected products found yet.
                        </div>
                      ) : (
                        <div className="overflow-x-auto border rounded-xl">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="bg-gray-100 border-b border-gray-200 text-xs uppercase text-gray-500 font-bold">
                                <th className="p-4">SKU</th>
                                <th className="p-4">Product</th>
                                <th className="p-4">Status</th>
                                <th className="p-4 text-center">Images</th>
                                <th className="p-4">Time Taken</th>
                                <th className="p-4">Finished / Updated</th>
                                <th className="p-4">Manager Note</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-200 text-sm text-gray-700">
                              {selfPerformanceStats.doneProductLedger.map((item) => (
                                <tr key={item.id} className="hover:bg-gray-50/60">
                                  <td className="p-4 font-mono font-bold text-xs text-gray-700">{item.sku}</td>
                                  <td className="p-4">
                                    <span className="font-bold text-gray-900 line-clamp-2">{item.product_name}</span>
                                  </td>
                                  <td className="p-4">
                                    <span className={`px-3 py-1 rounded-full border text-[10px] font-black uppercase ${
                                      item.status === 'Completed'
                                        ? 'bg-green-50 text-green-700 border-green-200'
                                        : 'bg-red-50 text-red-700 border-red-200'
                                    }`}>
                                      {item.status}
                                    </span>
                                  </td>
                                  <td className="p-4 text-center text-xs font-black text-gray-600">
                                    RAW {item.raw_count} / EDITED {item.edited_count}
                                  </td>
                                  <td className="p-4 text-sm font-black text-slate-900">{item.time_spent_label}</td>
                                  <td className="p-4 text-xs text-gray-500 font-semibold">{formatDateTime(item.updated_at)}</td>
                                  <td className="p-4 text-xs text-red-700 font-semibold max-w-xs">
                                    {item.status === 'Rejected' ? (item.rejection_note || 'No note') : '—'}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}

                      <div className="mt-4 p-3 rounded-xl bg-amber-50 border border-amber-200 text-amber-800 text-[11px] font-bold leading-relaxed">
                        Note: time tracking starts when you change a product to Processing and stops when it becomes Completed or Rejected. Older products may show “No tracked time”.
                      </div>
                    </div>
                  </div>
                )}

                {/* WORKBOOK MANAGEMENT HISTORY MODULES */}
                {activeTab === 'history_old' && (
                  <div className="bg-white rounded-xl p-6 border border-gray-200 shadow-xs">
                    <h2 className="text-xl font-bold tracking-tight text-gray-900 mb-6">Original Excel Sheets (OLD Archives)</h2>
                    {originalUploadedSheets.length === 0 ? (
                      <div className="text-center text-xs text-gray-400 font-bold uppercase py-12">
                        No original Excel archives found. Re-upload the Excel file to create an original archive.
                      </div>
                    ) : (
                      <div className="space-y-4">
                        {originalUploadedSheets.map((sheetName) => (
                          <div key={sheetName} className="p-4 border rounded-xl flex items-center justify-between bg-gray-50">
                            <h4 className="text-sm font-bold text-gray-800">{sheetName}</h4>
                            <button onClick={() => handleDownloadDashboardManifestByName(sheetName, 'original')} className="px-3 py-1.5 bg-amber-50 text-amber-700 font-bold text-xs uppercase rounded-lg cursor-pointer hover:bg-amber-100">Download Original</button>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}

                {activeTab === 'history_new' && (
                  <div className="bg-white rounded-xl p-6 border border-gray-200 shadow-xs">
                    <h2 className="text-xl font-bold tracking-tight text-gray-900 mb-6">Live Dynamic Manifests (NEW System)</h2>
                    {uniqueUploadedSheets.length === 0 ? <div className="text-center text-xs text-gray-400 font-bold uppercase py-12">No Manifests Available.</div> : (
                      <div className="space-y-4">
                        {uniqueUploadedSheets.map((sheetName) => (
                          <div key={sheetName} className="p-4 border rounded-xl flex items-center justify-between bg-slate-50/50">
                            <h4 className="text-sm font-bold text-slate-800">{sheetName}</h4>
                            <div className="flex items-center gap-2">
                              <button onClick={() => handleDownloadDashboardManifestByName(sheetName, 'live')} className="px-3 py-1.5 bg-[rgba(138,21,56,0.85)] text-white font-bold text-xs uppercase rounded-lg cursor-pointer hover:bg-[#8a1538]">Download Edited</button>
                              <button onClick={() => handleDownloadMissingByName(sheetName)} className="px-3 py-1.5 bg-amber-500 text-white font-bold text-xs uppercase rounded-lg cursor-pointer hover:bg-amber-600">Download Missing</button>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                )}

                {/* MEDIA FOLDER DISPATCH OVERVIEW */}
                {activeTab === 'images' && (
                  <div className="bg-white rounded-xl p-6 border border-gray-200 shadow-xs space-y-6">
                    <div className="flex flex-col lg:flex-row lg:items-start lg:justify-between gap-4">
                      <div>
                        <h2 className="text-xl font-bold text-gray-900">Media File Directory System</h2>
                        <p className="text-xs text-gray-400 font-semibold mt-1">
                          Search, filter by store, and bulk-download product RAW / EDITED image folders.
                        </p>
                      </div>

                      <div className="flex flex-wrap items-center gap-2">
                        <button
                          onClick={() => handleBulkDownloadAssetDirectory(assetDirectoryProducts, 'all')}
                          disabled={uploading || assetDirectoryProducts.length === 0 || (assetDirectoryRawCount + assetDirectoryEditedCount) === 0}
                          className="px-3 py-2 bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                        >
                          ⬇️ Bulk All
                        </button>
                        <button
                          onClick={() => handleBulkDownloadAssetDirectory(assetDirectoryProducts, 'raw')}
                          disabled={uploading || assetDirectoryRawCount === 0}
                          className="px-3 py-2 bg-[rgba(138,21,56,0.95)] hover:bg-[rgba(138,21,56,0.85)] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                        >
                          RAW ZIP
                        </button>
                        <button
                          onClick={() => handleBulkDownloadAssetDirectory(assetDirectoryProducts, 'edited')}
                          disabled={uploading || assetDirectoryEditedCount === 0}
                          className="px-3 py-2 bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                        >
                          EDITED ZIP
                        </button>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-4 gap-3 bg-gray-50 border border-gray-200 rounded-2xl p-4">
                      <div className="lg:col-span-2">
                        <label className="block text-[10px] font-black uppercase tracking-widest text-gray-400 mb-1">Search SKU / Product / Store</label>
                        <input
                          type="text"
                          value={assetFolderSearchQuery}
                          onChange={(e) => setAssetFolderSearchQuery(e.target.value)}
                          placeholder="Search asset folders..."
                          className="w-full px-3 py-2 border border-gray-200 rounded-xl bg-white text-sm text-gray-900 outline-none"
                        />
                      </div>

                      <div>
                        <label className="block text-[10px] font-black uppercase tracking-widest text-gray-400 mb-1">Store</label>
                        <select
                          value={assetFolderStoreId}
                          onChange={(e) => setAssetFolderStoreId(e.target.value)}
                          className="w-full px-3 py-2 border border-gray-200 rounded-xl bg-white text-sm text-gray-900 outline-none font-bold"
                        >
                          <option value="ALL">All Stores</option>
                          {stores.map(store => (
                            <option key={store.id} value={store.id}>{store.name}</option>
                          ))}
                        </select>
                      </div>

                      <div className="flex items-end">
                        <div className="w-full grid grid-cols-3 gap-2 text-center">
                          <div className="bg-white border rounded-xl p-2">
                            <div className="text-[9px] font-black text-gray-400 uppercase">Folders</div>
                            <div className="text-base font-black text-gray-900">{assetDirectoryProducts.length}</div>
                          </div>
                          <div className="bg-white border rounded-xl p-2">
                            <div className="text-[9px] font-black text-gray-400 uppercase">RAW</div>
                            <div className="text-base font-black text-slate-800">{assetDirectoryRawCount}</div>
                          </div>
                          <div className="bg-white border rounded-xl p-2">
                            <div className="text-[9px] font-black text-gray-400 uppercase">Edited</div>
                            <div className="text-base font-black text-[#8a1538]">{assetDirectoryEditedCount}</div>
                          </div>
                        </div>
                      </div>
                    </div>

                    {assetDirectoryProducts.length === 0 ? (
                      <div className="text-center text-xs text-gray-400 font-bold uppercase py-12 bg-gray-50 rounded-2xl border border-dashed">
                        No image folders found for the selected search/store filter.
                      </div>
                    ) : (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        {assetDirectoryProducts.map((prod) => {
                          const raws = getArray(prod.raw_image_url).filter(Boolean);
                          const edits = getArray(prod.edited_image_url).filter(Boolean);
                          return (
                            <div key={prod.id} className="p-5 border rounded-2xl bg-white shadow-xs hover:border-[rgba(138,21,56,0.28)] transition-colors">
                              <div className="flex items-start justify-between gap-3 mb-4">
                                <div>
                                  <span className="text-xs font-mono font-black text-[#8a1538] bg-[rgba(138,21,56,0.06)] px-2.5 py-1 rounded-md block w-fit">📁 SKU: {prod.sku || 'UNKNOWN'}</span>
                                  <div className="text-[11px] font-bold text-gray-400 mt-2">{prod.product_name || 'Unnamed Product'}</div>
                                  <div className="text-[10px] font-black uppercase tracking-wider text-emerald-700 mt-1">{getStoreNameById(prod.store_id)}</div>
                                </div>
                                <span className="text-[10px] font-black px-2 py-1 rounded-lg bg-gray-100 text-gray-600 uppercase">{prod.status || 'Missing'}</span>
                              </div>

                              <div className="space-y-4">
                                <div className="flex items-center justify-between gap-3 border rounded-xl p-3 bg-slate-50/70">
                                  <div>
                                    <h5 className="text-[10px] font-black tracking-wider uppercase text-gray-500">/RAW/ ({raws.length})</h5>
                                    <p className="text-[10px] text-gray-400 font-semibold mt-1">Original product images</p>
                                  </div>
                                  <button
                                    onClick={() => handleDownloadProductAssets(prod, 'raw')}
                                    disabled={raws.length === 0 || uploading}
                                    className="px-3 py-2 bg-[rgba(138,21,56,0.95)] hover:bg-[rgba(138,21,56,0.85)] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-lg text-[10px] font-black uppercase"
                                  >
                                    Download RAW
                                  </button>
                                </div>

                                <div className="flex items-center justify-between gap-3 border rounded-xl p-3 bg-[rgba(138,21,56,0.06)]/70">
                                  <div>
                                    <h5 className="text-[10px] font-black tracking-wider uppercase text-[rgba(138,21,56,0.85)]">/EDITED/ ({edits.length})</h5>
                                    <p className="text-[10px] text-[#8a1538] font-semibold mt-1">Final edited images</p>
                                  </div>
                                  <button
                                    onClick={() => handleDownloadProductAssets(prod, 'edited')}
                                    disabled={edits.length === 0 || uploading}
                                    className="px-3 py-2 bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-lg text-[10px] font-black uppercase"
                                  >
                                    Download Edited
                                  </button>
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}

              </div>
            )}

          </div>
        </main>
      </div>

      {/* AD-HOC MANUAL ENTRY MODAL */}
      {showAdHocModal && (
        <div className="fixed inset-0 bg-[rgba(138,21,56,1)]/60 backdrop-blur-md flex items-center justify-center p-4 z-[70] animate-fadeIn">
          <div className="bg-white border rounded-3xl w-full max-w-lg shadow-2xl p-8 overflow-hidden">
            <div className="flex items-center justify-between border-b pb-4 mb-6">
              <h3 className="font-black text-lg text-gray-900 uppercase">Create Ad-Hoc Item</h3>
              <button onClick={() => setShowAdHocModal(false)} className="text-gray-400 text-xl cursor-pointer hover:text-gray-900">✕</button>
            </div>
            <form onSubmit={handleAdHocSubmit} className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xxs font-black uppercase text-gray-400 tracking-widest mb-1.5">SKU / Barcode *</label>
                  <input type="text" required value={adHocForm.sku} onChange={(e) => setAdHocForm({...adHocForm, sku: e.target.value})} className="w-full px-4 py-2 text-sm border rounded-xl bg-gray-50 text-gray-900" />
                </div>
                <div>
                  <label className="block text-xxs font-black uppercase text-gray-400 tracking-widest mb-1.5">Stock Level</label>
                  <input type="number" value={adHocForm.stock_quantity} onChange={(e) => setAdHocForm({...adHocForm, stock_quantity: e.target.value})} className="w-full px-4 py-2 text-sm border rounded-xl bg-gray-50 text-gray-900" />
                </div>
              </div>
              <div>
                <label className="block text-xxs font-black uppercase text-gray-400 tracking-widest mb-1.5">Product Name *</label>
                <input type="text" required value={adHocForm.product_name} onChange={(e) => setAdHocForm({...adHocForm, product_name: e.target.value})} className="w-full px-4 py-2 text-sm border rounded-xl bg-gray-50 text-gray-900" />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xxs font-black uppercase text-gray-400 tracking-widest mb-1.5">Category</label>
                  <input type="text" value={adHocForm.category} onChange={(e) => setAdHocForm({...adHocForm, category: e.target.value})} className="w-full px-4 py-2 text-sm border rounded-xl bg-gray-50 text-gray-900" />
                </div>
                <div>
                  <label className="block text-xxs font-black uppercase text-gray-400 tracking-widest mb-1.5">Warehouse</label>
                  <input type="text" value={adHocForm.warehouse} onChange={(e) => setAdHocForm({...adHocForm, warehouse: e.target.value})} className="w-full px-4 py-2 text-sm border rounded-xl bg-gray-50 text-gray-900" />
                </div>
              </div>
              <div className="flex gap-3 pt-4">
                <button type="button" onClick={() => setShowAdHocModal(false)} className="w-1/3 py-3 border rounded-xl text-xs font-bold uppercase text-gray-600 hover:bg-gray-50">Cancel</button>
                <button type="submit" disabled={uploading} className="w-2/3 py-3 bg-[#8a1538] hover:bg-[rgba(138,21,56,0.95)] text-white text-xs uppercase font-black rounded-xl cursor-pointer">
                  {uploading ? "Saving..." : "Add to Live Dashboard"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* CUSTOM DEPARTMENT MODAL */}
      {showRoleModal && (
        <div className="fixed inset-0 bg-[rgba(138,21,56,1)]/60 backdrop-blur-md flex items-center justify-center p-4 z-50 animate-fadeIn">
          <div className="bg-white border rounded-3xl w-full max-w-md shadow-2xl p-6 overflow-hidden">
            <div className="flex items-center justify-between border-b pb-3 mb-5">
              <h3 className="font-black text-base text-gray-900 uppercase">Deploy Custom Role Perms</h3>
              <button onClick={() => setShowRoleModal(false)} className="text-gray-400 text-lg cursor-pointer">✕</button>
            </div>
            <form onSubmit={handleCreateCustomRole} className="space-y-5">
              <div>
                <label className="block text-xxs font-black uppercase text-gray-400 tracking-widest mb-1.5">Bespoke Role Name</label>
                <input 
                  type="text" required value={newRoleForm.roleName} onChange={(e) => setNewRoleForm({...newRoleForm, roleName: e.target.value})} placeholder="e.g. Lead Editor..." 
                  className="w-full px-4 py-2.5 text-sm border rounded-xl bg-gray-50 text-gray-900"
                />
              </div>
              <div className="space-y-3 bg-slate-50 p-4 rounded-xl border">
                <label className="flex items-center justify-between cursor-pointer"><span className="text-xs text-gray-700">Bulk Upload Media</span><input type="checkbox" checked={newRoleForm.canUploadAssets} onChange={(e) => setNewRoleForm({...newRoleForm, canUploadAssets: e.target.checked})} /></label>
                <label className="flex items-center justify-between cursor-pointer"><span className="text-xs text-gray-700">Modify Layout Sheets</span><input type="checkbox" checked={newRoleForm.canModifyDataSheets} onChange={(e) => setNewRoleForm({...newRoleForm, canModifyDataSheets: e.target.checked})} /></label>
              </div>
              <div className="flex gap-3 pt-2">
                <button type="button" onClick={() => setShowRoleModal(false)} className="w-1/3 py-2.5 border rounded-xl text-xs font-bold text-gray-600">Cancel</button>
                <button type="submit" className="w-2/3 py-2.5 bg-[#8a1538] text-white text-xs uppercase font-black rounded-xl cursor-pointer">Deploy Role</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* SELECTED STAFF DETAILS MODAL */}
      {selectedOperatorStats && (
        <div className="fixed inset-0 bg-[rgba(138,21,56,0.85)]/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fadeIn">
          <div className="bg-white rounded-2xl w-full max-w-3xl shadow-2xl p-6 border border-gray-100 max-h-[92vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b pb-3 mb-4">
              <h3 className="font-black text-lg text-gray-900 uppercase tracking-tight">📈 Staff Operations Ledger</h3>
              <button onClick={() => setSelectedOperatorStats(null)} className="text-gray-400 hover:text-gray-600 font-bold text-lg cursor-pointer">✕</button>
            </div>

            <div className="space-y-5 text-xs font-semibold text-gray-500">
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200">
                <span className="text-[10px] font-black tracking-wider text-gray-400 uppercase block">Operator Target Profile ID</span>
                <div className="text-base font-black text-slate-800 mt-1">👤 {selectedOperatorStats.username}</div>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-6 gap-4">
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Total Claimed</span>
                  <div className="text-xl font-black text-slate-800 mt-1">{selectedOperatorStats.performance.totalClaimed}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Missing</span>
                  <div className="text-xl font-black text-amber-600 mt-1">{selectedOperatorStats.performance.totalMissing}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Processing</span>
                  <div className="text-xl font-black text-[rgba(138,21,56,0.85)] mt-1">{selectedOperatorStats.performance.totalProcessing}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Completed</span>
                  <div className="text-xl font-black text-green-600 mt-1">{selectedOperatorStats.performance.totalCompleted}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Modified</span>
                  <div className="text-xl font-black text-[rgba(138,21,56,0.85)] mt-1">{selectedOperatorStats.performance.totalModified}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Rejected</span>
                  <div className="text-xl font-black text-red-600 mt-1">{selectedOperatorStats.performance.totalRejected}</div>
                </div>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-6 gap-4">
                <div className="border p-4 rounded-xl bg-green-50 border-green-100">
                  <span className="text-[9px] font-bold text-green-700 uppercase block">Completed This Week</span>
                  <div className="text-xl font-black text-green-700 mt-1">{selectedOperatorStats.performance.weekCount}</div>
                </div>
                <div className="border p-4 rounded-xl bg-green-50 border-green-100">
                  <span className="text-[9px] font-bold text-green-700 uppercase block">Completed This Month</span>
                  <div className="text-xl font-black text-green-700 mt-1">{selectedOperatorStats.performance.monthCount}</div>
                </div>
                <div className="border p-4 rounded-xl bg-[rgba(138,21,56,0.06)] border-[rgba(138,21,56,0.18)]">
                  <span className="text-[9px] font-bold text-[#8a1538] uppercase block">Modified This Week</span>
                  <div className="text-xl font-black text-[#8a1538] mt-1">{selectedOperatorStats.performance.modifiedWeekCount}</div>
                </div>
                <div className="border p-4 rounded-xl bg-red-50 border-red-100">
                  <span className="text-[9px] font-bold text-red-700 uppercase block">Rejected This Week</span>
                  <div className="text-xl font-black text-red-700 mt-1">{selectedOperatorStats.performance.rejectedWeekCount}</div>
                </div>
                <div className="border p-4 rounded-xl bg-[rgba(138,21,56,0.06)] border-[rgba(138,21,56,0.18)]">
                  <span className="text-[9px] font-bold text-[#8a1538] uppercase block">Modified This Month</span>
                  <div className="text-xl font-black text-[#8a1538] mt-1">{selectedOperatorStats.performance.modifiedMonthCount}</div>
                </div>
                <div className="border p-4 rounded-xl bg-red-50 border-red-100">
                  <span className="text-[9px] font-bold text-red-700 uppercase block">Rejected This Month</span>
                  <div className="text-xl font-black text-red-700 mt-1">{selectedOperatorStats.performance.rejectedMonthCount}</div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="border p-4 rounded-xl bg-slate-50">
                  <span className="text-[9px] font-bold text-gray-500 uppercase block">Avg Time To Complete One Product</span>
                  <div className="text-lg font-black text-slate-900 mt-1">{selectedOperatorStats.performance.avgCompletedTime}</div>
                </div>
                <div className="border p-4 rounded-xl bg-red-50 border-red-100">
                  <span className="text-[9px] font-bold text-red-700 uppercase block">Avg Time Before Rejection</span>
                  <div className="text-lg font-black text-red-800 mt-1">{selectedOperatorStats.performance.avgRejectedTime}</div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="border p-4 rounded-xl bg-white">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Fastest Completed Product</span>
                  <div className="text-lg font-black text-emerald-700 mt-1">{selectedOperatorStats.performance.fastestCompletedTime}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Slowest Completed Product</span>
                  <div className="text-lg font-black text-orange-700 mt-1">{selectedOperatorStats.performance.slowestCompletedTime}</div>
                </div>
              </div>

              <div className="p-3 rounded-xl bg-amber-50 border border-amber-200 text-amber-800 text-[11px] font-bold leading-relaxed">
                Note: time tracking starts when the employee changes status to Processing. Older products completed before this update may show “No tracked time” because their timer was not recorded.
              </div>
            </div>

            <button onClick={() => setSelectedOperatorStats(null)} className="w-full mt-6 py-2.5 bg-[rgba(138,21,56,0.85)] text-white font-bold text-xs uppercase tracking-wider rounded-xl cursor-pointer hover:bg-[rgba(138,21,56,0.95)]">Dismiss Ledger View</button>
          </div>
        </div>
      )}

      {/* OPERATOR WORKSPACE MODAL */}
      {selectedProduct && (() => {
        const rawAssets = getArray(selectedProduct.raw_image_url);
        const editAssets = getArray(selectedProduct.edited_image_url);
        const canUploadProductImages =
          authRole === 'Admin' ||
          authRole === 'Manager' ||
          (selectedProduct.status === 'Processing' && selectedProduct.processed_by === loginUser);
        const isUploadingImagesForThisProduct = imageUploadingProductId === selectedProduct.id;
        return (
          <div className="fixed inset-0 bg-[rgba(138,21,56,0.85)]/60 backdrop-blur-xs flex items-center justify-center p-4 z-[60] animate-fadeIn">
            <div className="bg-white rounded-2xl w-full max-w-4xl shadow-2xl overflow-hidden flex flex-col max-h-[95vh]">
              <div className="p-6 border-b flex justify-between bg-gray-50">
                <h2 className="text-lg font-bold text-gray-900">Workspace • {selectedProduct.product_name}</h2>
                <button onClick={() => setSelectedProduct(null)} className="text-gray-400 text-xl font-bold cursor-pointer hover:text-gray-900">✕</button>
              </div>
              <div className="p-6 overflow-y-auto space-y-6 flex-grow">
                <div className="grid grid-cols-2 gap-4 bg-gray-50 p-4 rounded-xl border">
                  <div><label className="block text-xxs font-black text-gray-400 uppercase">Stock Level</label><div className="text-sm font-bold text-gray-900">{selectedProduct.stock_quantity} Units</div></div>
                  <div><label className="block text-xxs font-black text-gray-400 uppercase mb-1">Operational Status</label>
                    <select 
                      value={selectedProduct.status || 'Missing'} 
                      onChange={(e) => handleOperatorStatusChange(selectedProduct.id, e.target.value)}
                      className={`w-full text-xs font-bold uppercase rounded-lg p-2 border shadow-sm focus:outline-none focus:ring-2 focus:ring-[#8a1538] cursor-pointer ${
                        selectedProduct.status === 'Rejected'
                          ? 'bg-red-50 border-red-200 text-red-700'
                          : 'bg-white border-gray-300 text-gray-900'
                      }`}
                    >
                      <option value="Missing" disabled={selectedProduct.status === 'Rejected'}>Missing</option>
                      <option value="Processing">Processing</option>
                      <option value="Completed" disabled={selectedProduct.status === 'Rejected'}>Completed</option>
                      <option value="Modified" disabled>Modified</option>
                      <option value="Rejected" disabled>Rejected</option>
                    </select>
                  </div>
                </div>
                {selectedProduct.status === 'Rejected' && (
                  <div className="p-4 rounded-xl border border-red-200 bg-red-50">
                    <div className="text-xs font-black text-red-700 uppercase tracking-wider mb-2">
                      Rejected by Manager
                    </div>
                    <p className="text-sm font-semibold text-red-900 whitespace-pre-wrap">
                      {selectedProduct.rejection_note || 'No rejection note was provided.'}
                    </p>
                    <p className="text-[11px] text-red-600 font-bold mt-3">
                      Change the status to <span className="underline">Processing</span> to reopen this item, upload corrected images, and then mark it Completed again.
                    </p>
                  </div>
                )}

                {!canUploadProductImages && (
                  <div className={`p-4 rounded-xl border text-xs font-bold ${selectedProduct.status === 'Rejected' ? 'border-red-200 bg-red-50 text-red-800' : 'border-amber-200 bg-amber-50 text-amber-800'}`}>
                    🔒 Image upload is locked. Change Operational Status to <span className="underline">Processing</span> first. That claims/reopens this product under your employee name and unlocks RAW/EDITED uploads.
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                  <div>
                    <div className="flex items-center justify-between gap-3 mb-3 border-b pb-2">
                      <h3 className="text-xs font-black text-gray-400 uppercase">📸 RAW IMAGES ({rawAssets.length})</h3>
                      <button
                        type="button"
                        disabled={rawAssets.length === 0 || uploading}
                        onClick={() => handleDownloadProductAssets(selectedProduct, 'raw')}
                        className={`px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider border ${rawAssets.length > 0 && !uploading ? 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50 cursor-pointer' : 'bg-gray-100 text-gray-400 border-gray-200 cursor-not-allowed'}`}
                      >
                        ⬇ Download RAW
                      </button>
                    </div>
                    <div className="grid grid-cols-3 gap-2">
                      {rawAssets.map((url, i) => (<img src={url} className="w-full h-20 object-cover border rounded" key={url} />))}
                      
                      {canUploadProductImages ? (
                        <label className={`w-full h-20 border-2 border-dashed rounded flex flex-col items-center justify-center font-bold text-center transition-colors ${isUploadingImagesForThisProduct ? 'cursor-not-allowed bg-gray-100 text-gray-400' : 'cursor-pointer text-gray-400 hover:bg-gray-50 hover:text-gray-600'}`}>
                          <span className="text-xl">{isUploadingImagesForThisProduct ? '⏳' : '+'}</span>
                          <span className="text-[9px] uppercase tracking-wider">{isUploadingImagesForThisProduct ? 'Uploading' : 'Multi Upload'}</span>
                          <input
                            type="file"
                            accept="image/*"
                            multiple
                            disabled={isUploadingImagesForThisProduct}
                            onChange={(e) => handleImageUpload(selectedProduct.id, 'raw', e)}
                            className="hidden"
                          />
                        </label>
                      ) : (
                        <div className="w-full h-20 border-2 border-dashed rounded flex flex-col items-center justify-center text-center text-gray-400 bg-gray-50 text-[10px] font-black uppercase tracking-wider">
                          🔒 Set Processing First
                        </div>
                      )}
                    </div>
                    <p className="text-[10px] text-gray-400 font-semibold mt-2">You can select 5, 10, or more RAW images at once.</p>
                  </div>
                  <div>
                    <div className="flex items-center justify-between gap-3 mb-3 border-b pb-2">
                      <h3 className="text-xs font-black text-gray-400 uppercase">✨ EDITED IMAGES ({editAssets.length})</h3>
                      <button
                        type="button"
                        disabled={editAssets.length === 0 || uploading}
                        onClick={() => handleDownloadProductAssets(selectedProduct, 'edited')}
                        className={`px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider border ${editAssets.length > 0 && !uploading ? 'bg-[rgba(138,21,56,0.06)] text-[#8a1538] border-[rgba(138,21,56,0.28)] hover:bg-[rgba(138,21,56,0.10)] cursor-pointer' : 'bg-gray-100 text-gray-400 border-gray-200 cursor-not-allowed'}`}
                      >
                        ⬇ Download EDITED
                      </button>
                    </div>
                    <div className="grid grid-cols-3 gap-2">
                      {editAssets.map((url, i) => (<img src={url} className="w-full h-20 object-cover border rounded" key={url} />))}
                      
                      {canUploadProductImages ? (
                        <label className={`w-full h-20 border-2 border-dashed rounded flex flex-col items-center justify-center font-bold text-center transition-colors ${isUploadingImagesForThisProduct ? 'cursor-not-allowed bg-gray-100 text-gray-400' : 'cursor-pointer text-gray-400 hover:bg-gray-50 hover:text-gray-600'}`}>
                          <span className="text-xl">{isUploadingImagesForThisProduct ? '⏳' : '+'}</span>
                          <span className="text-[9px] uppercase tracking-wider">{isUploadingImagesForThisProduct ? 'Uploading' : 'Multi Upload'}</span>
                          <input
                            type="file"
                            accept="image/*"
                            multiple
                            disabled={isUploadingImagesForThisProduct}
                            onChange={(e) => handleImageUpload(selectedProduct.id, 'edited', e)}
                            className="hidden"
                          />
                        </label>
                      ) : (
                        <div className="w-full h-20 border-2 border-dashed rounded flex flex-col items-center justify-center text-center text-gray-400 bg-gray-50 text-[10px] font-black uppercase tracking-wider">
                          🔒 Set Processing First
                        </div>
                      )}
                    </div>
                    <p className="text-[10px] text-gray-400 font-semibold mt-2">You can select 5, 10, or more EDITED images at once.</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {/* MANAGER PREVIEW MODAL */}
      {managerPreview && (() => {
        const rawAssets = getArray(managerPreview.raw_image_url);
        const editAssets = getArray(managerPreview.edited_image_url);
        const cleanSku = String(managerPreview.sku || 'UNKNOWN').replace(/[^a-zA-Z0-9_-]/g, '_');
        const renderManagerAssetCard = (url, typeLabel, index) => {
          const filename = `${cleanSku}_${typeLabel.toLowerCase()}_${index + 1}.jpg`;
          return (
            <div key={`${typeLabel}-${url}-${index}`} className="relative group border rounded-xl overflow-hidden bg-gray-50">
              <button
                type="button"
                onClick={() => setFullViewImage({ url, label: `${typeLabel} Image ${index + 1}`, sku: managerPreview.sku })}
                className="block w-full cursor-zoom-in"
                title="Open full view"
              >
                <img src={url} className="w-full h-40 object-cover" alt={`${typeLabel} image ${index + 1}`} />
              </button>
              <div className="p-2 flex gap-2 bg-white border-t">
                <button
                  type="button"
                  onClick={() => setFullViewImage({ url, label: `${typeLabel} Image ${index + 1}`, sku: managerPreview.sku })}
                  className="flex-1 px-2 py-1.5 bg-[rgba(138,21,56,0.85)] text-white rounded-lg text-[10px] font-black uppercase"
                >
                  Full View
                </button>
                <button
                  type="button"
                  onClick={() => handleDownloadSingleAsset(url, filename)}
                  className="flex-1 px-2 py-1.5 bg-[rgba(138,21,56,0.06)] text-[#8a1538] border border-[rgba(138,21,56,0.18)] rounded-lg text-[10px] font-black uppercase"
                >
                  Download
                </button>
              </div>
            </div>
          );
        };
        return (
          <div className="fixed inset-0 bg-[rgba(138,21,56,0.85)]/60 backdrop-blur-sm flex items-center justify-center p-4 z-[80]">
            <div className="bg-white rounded-2xl w-full max-w-6xl p-8 shadow-2xl max-h-[95vh] overflow-y-auto flex flex-col">
              <div className="flex flex-col md:flex-row md:justify-between md:items-start gap-4 mb-6 border-b pb-4">
                <div>
                  <h2 className="text-2xl font-black text-gray-900">{managerPreview.product_name}</h2>
                  <p className="text-xs text-gray-400 font-bold mt-1">SKU: {managerPreview.sku} • Status: {managerPreview.status}</p>
                </div>
                <div className="flex flex-wrap gap-2 items-center">
                  <button
                    onClick={() => handleDownloadProductAssets(managerPreview, 'raw')}
                    disabled={rawAssets.length === 0 || uploading}
                    className="px-3 py-2 bg-amber-50 text-amber-700 border border-amber-100 rounded-xl text-[10px] font-black uppercase disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    ⬇ RAW ZIP
                  </button>
                  <button
                    onClick={() => handleDownloadProductAssets(managerPreview, 'edited')}
                    disabled={editAssets.length === 0 || uploading}
                    className="px-3 py-2 bg-[rgba(138,21,56,0.06)] text-[#8a1538] border border-[rgba(138,21,56,0.18)] rounded-xl text-[10px] font-black uppercase disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    ⬇ EDITED ZIP
                  </button>
                  <button onClick={closeManagerPreview} className="text-gray-400 font-bold text-2xl hover:text-gray-900 cursor-pointer px-2">✕</button>
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-6">
                <div>
                  <div className="flex items-center justify-between mb-3">
                    <h3 className="text-xs font-black text-gray-400 uppercase">RAW Images ({rawAssets.length})</h3>
                    <button onClick={() => handleDownloadProductAssets(managerPreview, 'raw')} disabled={rawAssets.length === 0 || uploading} className="text-[10px] font-black uppercase text-amber-700 disabled:text-gray-300">Download All</button>
                  </div>
                  {rawAssets.length === 0 ? (
                    <div className="h-40 rounded-xl border-2 border-dashed flex items-center justify-center text-xs font-bold text-gray-400 uppercase">No RAW Images</div>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">{rawAssets.map((url, i) => renderManagerAssetCard(url, 'RAW', i))}</div>
                  )}
                </div>
                <div>
                  <div className="flex items-center justify-between mb-3">
                    <h3 className="text-xs font-black text-gray-400 uppercase">EDITED Images ({editAssets.length})</h3>
                    <button onClick={() => handleDownloadProductAssets(managerPreview, 'edited')} disabled={editAssets.length === 0 || uploading} className="text-[10px] font-black uppercase text-[#8a1538] disabled:text-gray-300">Download All</button>
                  </div>
                  {editAssets.length === 0 ? (
                    <div className="h-40 rounded-xl border-2 border-dashed flex items-center justify-center text-xs font-bold text-gray-400 uppercase">No EDITED Images</div>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">{editAssets.map((url, i) => renderManagerAssetCard(url, 'EDITED', i))}</div>
                  )}
                </div>
              </div>
              {!isRejecting ? (
                <div className="flex gap-4"><button onClick={() => setIsRejecting(true)} className="w-1/3 py-3 bg-red-50 text-red-600 hover:bg-red-100 font-bold rounded-xl text-sm border uppercase cursor-pointer transition-colors">Reject</button><button onClick={closeManagerPreview} className="w-2/3 py-3 bg-gray-100 hover:bg-gray-200 font-bold rounded-xl text-sm text-gray-600 uppercase cursor-pointer transition-colors">Close</button></div>
              ) : (
                <div className="p-5 bg-red-50 border border-red-200 rounded-xl">
                  <textarea className="w-full p-3 border rounded-lg text-sm mb-4 bg-white text-gray-900 focus:outline-none focus:ring-2 focus:ring-red-500" rows="3" placeholder="Provide rejection reasons..." value={rejectNote} onChange={(e) => setRejectNote(e.target.value)}></textarea>
                  <div className="flex gap-3"><button onClick={() => handleRejectProduct(managerPreview.id)} className="px-6 py-2 bg-red-600 hover:bg-red-700 text-white font-bold rounded-lg text-xs uppercase cursor-pointer">Confirm Reject</button><button onClick={() => setIsRejecting(false)} className="px-6 py-2 bg-white text-gray-700 font-bold rounded-lg border text-xs uppercase cursor-pointer hover:bg-gray-50">Cancel</button></div>
                </div>
              )}
            </div>
          </div>
        );
      })()}

      {/* FULL IMAGE VIEWER MODAL */}
      {fullViewImage && (
        <div className="fixed inset-0 bg-black/85 backdrop-blur-sm flex items-center justify-center p-4 z-[100]">
          <div className="bg-white rounded-2xl w-full max-w-6xl max-h-[95vh] overflow-hidden shadow-2xl flex flex-col">
            <div className="p-4 border-b flex items-center justify-between gap-4">
              <div>
                <h3 className="text-sm font-black text-gray-900 uppercase">{fullViewImage.label}</h3>
                <p className="text-[10px] font-bold text-gray-400">SKU: {fullViewImage.sku || 'UNKNOWN'}</p>
              </div>
              <div className="flex gap-2">
                <button onClick={() => handleDownloadSingleAsset(fullViewImage.url, `${String(fullViewImage.sku || 'image').replace(/[^a-zA-Z0-9_-]/g, '_')}_${String(fullViewImage.label || 'image').replace(/[^a-zA-Z0-9_-]/g, '_')}.jpg`)} className="px-3 py-2 bg-[rgba(138,21,56,0.85)] text-white rounded-xl text-xs font-black uppercase">Download</button>
                <button onClick={() => setFullViewImage(null)} className="px-3 py-2 bg-gray-100 text-gray-700 rounded-xl text-xs font-black uppercase">Close</button>
              </div>
            </div>
            <div className="p-4 overflow-auto bg-[rgba(138,21,56,1)] flex items-center justify-center">
              <img src={fullViewImage.url} alt={fullViewImage.label || 'Full image'} className="max-w-full max-h-[78vh] object-contain rounded-lg" />
            </div>
          </div>
        </div>
      )}

    </div>
  );
}