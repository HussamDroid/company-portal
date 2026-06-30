'use client';

import { useState, useEffect } from 'react';
import { supabase } from '../utils/supabase';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

// THEME NOTE: UI accent updated to exact BlackRose white + maroon palette.
// Primary accent uses rgba(138, 21, 56, 0.85) with white surfaces.
// Semantic status colors are still used where helpful.


const PORTAL_PERMISSION_FEATURES = [
  { key: 'dashboard', label: 'Dashboard Access', category: 'Core Portal' },
  { key: 'stores_manage', label: 'Create/Edit/Delete Stores', category: 'Stores & Catalog' },
  { key: 'excel_upload', label: 'Upload Excel Sheets', category: 'Stores & Catalog' },
  { key: 'product_edit', label: 'Edit Product Details', category: 'Stores & Catalog' },
  { key: 'ad_hoc', label: 'Add Ad-Hoc Products', category: 'Stores & Catalog' },
  { key: 'bulk_images', label: 'Bulk Image Upload', category: 'Media & Assets' },
  { key: 'product_images', label: 'Product Image Upload', category: 'Media & Assets' },
  { key: 'asset_downloads', label: 'SKU Asset Downloads', category: 'Media & Assets' },
  { key: 'selected_downloads', label: 'Selected Asset Downloads', category: 'Media & Assets' },
  { key: 'claim_status', label: 'Claim / Change Status', category: 'Product Workflow' },
  { key: 'submit_review', label: 'Submit Under Review', category: 'Product Workflow' },
  { key: 'review_images', label: 'Review Modal', category: 'Product Workflow' },
  { key: 'compare_images', label: 'Compare RAW/EDITED', category: 'Product Workflow' },
  { key: 'reject', label: 'Reject Products', category: 'Product Workflow' },
  { key: 'approve', label: 'Ready To Upload Approval', category: 'Product Workflow' },
  { key: 'status_override', label: 'Status Override', category: 'Product Workflow' },
  { key: 'remove_edited', label: 'Remove Edited Images', category: 'Product Workflow' },
  { key: 'sheet_exports', label: 'Original / Live Sheet Exports', category: 'Sheets & Exports' },
  { key: 'tasks_view', label: 'View Tasks', category: 'Tasks & Messaging' },
  { key: 'tasks_reply', label: 'Reply To Tasks', category: 'Tasks & Messaging' },
  { key: 'tasks_manage', label: 'Post/Edit/Delete Tasks', category: 'Tasks & Messaging' },
  { key: 'staff_register', label: 'Register / Edit Staff', category: 'Team & Admin' },
  { key: 'google_view', label: 'Open Google Sheets', category: 'Shared Links' },
  { key: 'google_manage', label: 'Manage Google Sheet Link', category: 'Shared Links' },
  { key: 'performance', label: 'Performance View', category: 'Team & Admin' }
];

const createPortalPermissionSet = (defaultValue = false) =>
  PORTAL_PERMISSION_FEATURES.reduce((acc, feature) => {
    acc[feature.key] = defaultValue;
    return acc;
  }, {});

const createEmptyRoleForm = () => ({
  roleName: '',
  permissions: createPortalPermissionSet(false)
});

const LEGACY_PERMISSION_GROUPS = {
  canUploadAssets: ['bulk_images', 'product_images', 'asset_downloads', 'selected_downloads'],
  canModifyDataSheets: ['excel_upload', 'product_edit', 'sheet_exports', 'ad_hoc'],
  canReviewArrays: ['review_images', 'compare_images', 'reject', 'approve'],
  canSuperviseStaff: ['stores_manage', 'tasks_manage', 'staff_register', 'google_manage', 'status_override', 'performance']
};

const CORE_ROLE_PERMISSION_KEYS = {
  Operator: [
    'dashboard', 'product_images', 'claim_status', 'submit_review', 'remove_edited',
    'asset_downloads', 'selected_downloads', 'tasks_view', 'tasks_reply', 'google_view', 'performance'
  ],
  Photographer: ['dashboard', 'bulk_images', 'asset_downloads', 'selected_downloads', 'tasks_view', 'tasks_reply', 'google_view'],
  'Content Editor': ['dashboard', 'excel_upload', 'product_edit', 'sheet_exports', 'ad_hoc', 'tasks_view', 'tasks_reply', 'google_view']
};

const ACTION_PERMISSION_ALIASES = {
  view_workspace: ['dashboard'],
  upload_assets: ['bulk_images', 'product_images', 'asset_downloads', 'selected_downloads'],
  modify_sheets: ['excel_upload', 'product_edit', 'sheet_exports', 'ad_hoc'],
  review_arrays: ['review_images', 'compare_images', 'reject', 'approve'],
  supervise_staff: ['stores_manage', 'tasks_manage', 'staff_register', 'google_manage', 'status_override', 'performance']
};

const hasOwnPermissionKey = (permissions, key) => Object.prototype.hasOwnProperty.call(permissions || {}, key);

const roleHasGranularPermissions = (permissions = {}) =>
  PORTAL_PERMISSION_FEATURES.some(feature => hasOwnPermissionKey(permissions, feature.key));

const normalizeRolePermissions = (permissions = {}) => {
  const normalized = createPortalPermissionSet(false);
  PORTAL_PERMISSION_FEATURES.forEach(feature => {
    if (hasOwnPermissionKey(permissions, feature.key)) {
      normalized[feature.key] = Boolean(permissions[feature.key]);
    }
  });
  return normalized;
};

const expandLegacyPermissions = (permissions = {}) => {
  const expanded = createPortalPermissionSet(false);

  Object.entries(LEGACY_PERMISSION_GROUPS).forEach(([legacyKey, featureKeys]) => {
    if (permissions?.[legacyKey]) {
      featureKeys.forEach(featureKey => {
        expanded[featureKey] = true;
      });
    }
  });

  // Old custom groups were always given these basic read permissions by the matrix.
  expanded.dashboard = true;
  expanded.tasks_view = true;
  expanded.tasks_reply = true;
  expanded.google_view = true;

  return expanded;
};

const roleAllowsPortalFeature = (roleName, customRoles = [], featureKey) => {
  if (!featureKey) return false;
  if (roleName === 'Admin' || roleName === 'Manager') return true;

  const coreFeatures = CORE_ROLE_PERMISSION_KEYS[roleName];
  if (coreFeatures) return coreFeatures.includes(featureKey);

  const customRole = customRoles.find(role => String(role.roleName || '').toLowerCase() === String(roleName || '').toLowerCase());
  if (!customRole) return false;

  const permissions = customRole.permissions || {};
  const resolvedPermissions = roleHasGranularPermissions(permissions)
    ? normalizeRolePermissions(permissions)
    : expandLegacyPermissions(permissions);

  return Boolean(resolvedPermissions[featureKey]);
};

const getLegacyFlagsFromFeaturePermissions = (permissions = {}) => ({
  canUploadAssets: LEGACY_PERMISSION_GROUPS.canUploadAssets.some(featureKey => Boolean(permissions[featureKey])),
  canModifyDataSheets: LEGACY_PERMISSION_GROUPS.canModifyDataSheets.some(featureKey => Boolean(permissions[featureKey])),
  canReviewArrays: LEGACY_PERMISSION_GROUPS.canReviewArrays.some(featureKey => Boolean(permissions[featureKey])),
  canSuperviseStaff: LEGACY_PERMISSION_GROUPS.canSuperviseStaff.some(featureKey => Boolean(permissions[featureKey]))
});

export default function IntegratedOperationsPortal() {
  // --- AUTHENTICATION STATES ---
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [authRole, setAuthRole] = useState(''); 
  const [loginUser, setLoginUser] = useState('');
  const [loginDisplayName, setLoginDisplayName] = useState('');
  const [loginPass, setLoginPass] = useState('');
  const [showLoginPassword, setShowLoginPassword] = useState(false);
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
  const [newRoleForm, setNewRoleForm] = useState(() => createEmptyRoleForm());

  // --- INTEGRATED NEW STAFF REGISTRATION FORM STATES ---
  const [regFullName, setRegFullName] = useState('');
  const [regUsername, setRegUsername] = useState('');
  const [regPassword, setRegPassword] = useState('');
  const [showRegPassword, setShowRegPassword] = useState(false);
  const [regRole, setRegRole] = useState('Operator');
  const [showPermissions, setShowPermissions] = useState(false);
  const [editingStaffId, setEditingStaffId] = useState(null);
  const [staffEditForm, setStaffEditForm] = useState({
    full_name: '',
    username: '',
    password: '',
    role: 'Operator'
  });
  const [showStaffEditPassword, setShowStaffEditPassword] = useState(false);

  // --- DATA FLOW STATES ---
  const [products, setProducts] = useState([]);
  const [manifestHistory, setManifestHistory] = useState([]);
  const [userRegistry, setUserRegistry] = useState([]);
  const [stores, setStores] = useState([]);
  const [selectedStoreId, setSelectedStoreId] = useState('ALL');
  const [showStoreCreate, setShowStoreCreate] = useState(false);
  const [newStoreName, setNewStoreName] = useState('');
  const [storeImageUploadingId, setStoreImageUploadingId] = useState(null);
  const [storeImageRemovingId, setStoreImageRemovingId] = useState(null);
  const [editingStore, setEditingStore] = useState(null);
  const [storeEditForm, setStoreEditForm] = useState({ name: '' });
  const [storeEditSaving, setStoreEditSaving] = useState(false);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [imageUploadingProductId, setImageUploadingProductId] = useState(null);
  const [realtimeStatus, setRealtimeStatus] = useState('Connecting...');

  // --- TASK BOARD STATES ---
  const [tasks, setTasks] = useState([]);
  const [taskSaving, setTaskSaving] = useState(false);
  const [editingTaskId, setEditingTaskId] = useState(null);
  const [taskError, setTaskError] = useState(null);
  const [taskReplies, setTaskReplies] = useState([]);
  const [taskReplyDrafts, setTaskReplyDrafts] = useState({});
  const [taskReplySavingId, setTaskReplySavingId] = useState(null);
  const [taskReplyError, setTaskReplyError] = useState(null);
  const [expandedTaskReplyIds, setExpandedTaskReplyIds] = useState({});
  const [taskSearchQuery, setTaskSearchQuery] = useState('');
  const [taskStatusFilter, setTaskStatusFilter] = useState('All');
  const [taskPriorityFilter, setTaskPriorityFilter] = useState('All');
  const [taskTargetFilter, setTaskTargetFilter] = useState('All');
  const [taskForm, setTaskForm] = useState({
    title: '',
    description: '',
    priority: 'Normal',
    assigned_role: 'All',
    due_at: ''
  });

  // --- SHARED GOOGLE SHEETS LINK STATE ---
  const [googleSheetLink, setGoogleSheetLink] = useState('');
  const [googleSheetDraft, setGoogleSheetDraft] = useState('');
  const [showGoogleSheetModal, setShowGoogleSheetModal] = useState(false);
  const [googleSheetSaving, setGoogleSheetSaving] = useState(false);

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
  const [selectedAssetProductIds, setSelectedAssetProductIds] = useState([]);
  const [editingId, setEditingId] = useState(null);
  
  // --- MODAL WORKSPACE STATES ---
  const [selectedProduct, setSelectedProduct] = useState(null); 
  const [managerPreview, setManagerPreview] = useState(null);   
  const [isRejecting, setIsRejecting] = useState(false);        
  const [rejectNote, setRejectNote] = useState('');             
  const [showManagerCompare, setShowManagerCompare] = useState(false);
  const [fullCompareIndex, setFullCompareIndex] = useState(null);

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

    const featureKeys = ACTION_PERMISSION_ALIASES[action] || [action];
    return featureKeys.some(featureKey => roleAllowsPortalFeature(authRole, customRoles, featureKey));
  };

  // --- SESSION PERSISTENCE ---
  useEffect(() => {
    const savedUser = localStorage.getItem('blackrose_user');
    const savedRole = localStorage.getItem('blackrose_role');
    const savedDisplayName = localStorage.getItem('blackrose_display_name');
    const savedCustomRoles = localStorage.getItem('blackrose_custom_roles');
    
    if (savedCustomRoles) {
      setCustomRoles(JSON.parse(savedCustomRoles));
    }
    if (savedUser && savedRole) {
      setLoginUser(savedUser);
      setLoginDisplayName(savedDisplayName || savedUser);
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

  // --- FULL IMAGE VIEWER KEYBOARD NAVIGATION ---
  useEffect(() => {
    if (!fullViewImage) return;

    const handleKeyNavigation = (event) => {
      if (event.key === 'Escape') {
        setFullViewImage(null);
        return;
      }

      if (event.key === 'ArrowRight') {
        event.preventDefault();
        moveFullViewImage(1);
      }

      if (event.key === 'ArrowLeft') {
        event.preventDefault();
        moveFullViewImage(-1);
      }
    };

    window.addEventListener('keydown', handleKeyNavigation);
    return () => window.removeEventListener('keydown', handleKeyNavigation);
  }, [fullViewImage]);

  // --- FULLSCREEN RAW/EDITED COMPARE KEYBOARD NAVIGATION ---
  useEffect(() => {
    if (fullCompareIndex === null || !managerPreview) return;

    const rawCount = getArray(managerPreview.raw_image_url).length;
    const editedCount = getArray(managerPreview.edited_image_url).length;
    const pairCount = Math.max(rawCount, editedCount);
    if (pairCount <= 0) return;

    const handleCompareKeyNavigation = (event) => {
      if (event.key === 'Escape') {
        setFullCompareIndex(null);
        return;
      }

      if (event.key === 'ArrowRight') {
        event.preventDefault();
        setFullCompareIndex(prev => ((Number.isInteger(prev) ? prev : 0) + 1 + pairCount) % pairCount);
      }

      if (event.key === 'ArrowLeft') {
        event.preventDefault();
        setFullCompareIndex(prev => ((Number.isInteger(prev) ? prev : 0) - 1 + pairCount) % pairCount);
      }
    };

    window.addEventListener('keydown', handleCompareKeyNavigation);
    return () => window.removeEventListener('keydown', handleCompareKeyNavigation);
  }, [fullCompareIndex, managerPreview]);

  // --- DATABASE DATA SYNCHRONIZERS ---
  async function fetchProducts() {
    const [prodsRes, historyRes, usersRes, tasksRes, taskRepliesRes, storesRes, settingsRes] = await Promise.all([
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
        .select('*'),
      supabase
        .from('task_board')
        .select('*')
        .order('created_at', { ascending: false }),
      supabase
        .from('task_replies')
        .select('*')
        .order('created_at', { ascending: true }),
      supabase
        .from('stores')
        .select('*')
        .order('name', { ascending: true }),
      supabase
        .from('portal_settings')
        .select('*')
        .eq('setting_key', 'google_sheet_link')
        .maybeSingle()
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

    if (!taskRepliesRes.error) {
      setTaskReplies(taskRepliesRes.data || []);
      setTaskReplyError(null);
    } else {
      console.warn('task_replies fetch failed. Check task_replies permissions or SQL setup:', taskRepliesRes.error);
      setTaskReplyError(taskRepliesRes.error.message || 'Task replies table is not ready.');
      setTaskReplies([]);
    }

    if (!storesRes.error) {
      setStores(storesRes.data || []);
    } else {
      console.warn('stores fetch failed. Run the stores SQL/policies if this is the first install:', storesRes.error);
      setStores([]);
    }

    if (!settingsRes.error) {
      const savedGoogleSheetLink = settingsRes.data?.setting_value || '';
      setGoogleSheetLink(savedGoogleSheetLink);
      setGoogleSheetDraft(savedGoogleSheetLink);
    } else {
      console.warn('portal_settings fetch failed. Run the v25 SQL setup to enable Google Sheets link sharing:', settingsRes.error);
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

  const sortStoresByName = (storeRows = []) =>
    [...storeRows].sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')));

  const applyStorePatchLocally = (storeId, patch) => {
    if (!storeId || !patch) return;

    setStores(prev => {
      const exists = prev.some(store => Number(store.id) === Number(storeId));
      const nextStores = exists
        ? prev.map(store => Number(store.id) === Number(storeId) ? { ...store, ...patch } : store)
        : [...prev, patch];
      return sortStoresByName(nextStores);
    });

    setEditingStore(prev =>
      prev && Number(prev.id) === Number(storeId) ? { ...prev, ...patch } : prev
    );
  };

  const openStoreEditModal = (store) => {
    if (!(authRole === 'Admin' || authRole === 'Manager') || !store?.id) return;

    const latestStore = stores.find(item => Number(item.id) === Number(store.id)) || store;
    setEditingStore(latestStore);
    setStoreEditForm({ name: latestStore.name || '' });
  };

  const closeStoreEditModal = () => {
    setEditingStore(null);
    setStoreEditForm({ name: '' });
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


  const normalizeGoogleSheetUrl = (value) => {
    const clean = String(value || '').trim();
    if (!clean) return '';
    if (clean.startsWith('http://') || clean.startsWith('https://')) return clean;
    return `https://${clean}`;
  };

  const openGoogleSheetLink = () => {
    const safeUrl = normalizeGoogleSheetUrl(googleSheetLink);

    if (safeUrl) {
      window.open(safeUrl, '_blank', 'noopener,noreferrer');
      return;
    }

    if (authRole === 'Admin' || authRole === 'Manager') {
      setGoogleSheetDraft('');
      setShowGoogleSheetModal(true);
      return;
    }

    alert('Google Sheet link has not been added yet.');
  };

  const handleSaveGoogleSheetLink = async (e) => {
    e.preventDefault();
    if (!(authRole === 'Admin' || authRole === 'Manager')) return;

    const normalizedUrl = normalizeGoogleSheetUrl(googleSheetDraft);
    if (!normalizedUrl) {
      alert('Paste a valid Google Sheets link before saving.');
      return;
    }

    setGoogleSheetSaving(true);
    try {
      const { data, error } = await supabase
        .from('portal_settings')
        .upsert({
          setting_key: 'google_sheet_link',
          setting_value: normalizedUrl,
          updated_by: loginDisplayName || loginUser || 'System User',
          updated_at: new Date().toISOString()
        }, { onConflict: 'setting_key' })
        .select('*')
        .single();

      if (error) throw error;

      setGoogleSheetLink(data?.setting_value || normalizedUrl);
      setGoogleSheetDraft(data?.setting_value || normalizedUrl);
      setShowGoogleSheetModal(false);
      alert('Google Sheets link saved successfully.');
    } catch (err) {
      alert('Google Sheets link save failed: ' + (err.message || 'Unknown error. Run the v25 SQL setup first.'));
    } finally {
      setGoogleSheetSaving(false);
    }
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
        return sortStoresByName(merged);
      });
      setSelectedStoreId(data.id);
      setNewStoreName('');
      setShowStoreCreate(false);
      setActiveTab('home');
    } catch (err) {
      alert('Store creation failed: ' + (err.message || 'Unknown error'));
    }
  };

  const handleUpdateStoreName = async (e) => {
    e.preventDefault();
    if (!(authRole === 'Admin' || authRole === 'Manager')) return;
    if (!editingStore?.id) return;

    const cleanName = storeEditForm.name.trim();
    if (!cleanName) {
      alert('Store name is required.');
      return;
    }

    setStoreEditSaving(true);
    try {
      const { data, error } = await supabase
        .from('stores')
        .update({ name: cleanName })
        .eq('id', editingStore.id)
        .select('*')
        .single();

      if (error) throw error;

      const updatedStore = data || { ...editingStore, name: cleanName };
      applyStorePatchLocally(editingStore.id, updatedStore);
      setEditingStore(updatedStore);
      setStoreEditForm({ name: updatedStore.name || cleanName });
      alert('Store name updated successfully.');
    } catch (err) {
      alert('Store name update failed: ' + (err.message || 'Unknown error'));
    } finally {
      setStoreEditSaving(false);
    }
  };

  const handleRemoveStoreImage = async (storeId) => {
    if (!(authRole === 'Admin' || authRole === 'Manager')) return;
    if (!storeId) return;

    const targetStore = stores.find(store => Number(store.id) === Number(storeId)) || editingStore;
    const previousImageUrl = targetStore?.image_url || '';

    if (!previousImageUrl) {
      alert('This store does not currently have an image to remove.');
      return;
    }

    if (!window.confirm(`Remove the image from store "${targetStore?.name || 'this store'}"?`)) return;

    setStoreImageRemovingId(storeId);
    try {
      const { data, error } = await supabase
        .from('stores')
        .update({ image_url: null })
        .eq('id', storeId)
        .select('*')
        .single();

      if (error) throw error;

      const updatedStore = data || { ...targetStore, image_url: null };
      applyStorePatchLocally(storeId, updatedStore);

      const storagePath = extractStoragePathFromPublicUrl(previousImageUrl);
      if (storagePath) {
        const { error: storageError } = await supabase.storage
          .from('product-assets')
          .remove([storagePath]);

        if (storageError) {
          console.warn('Store image was removed from the store record, but storage file deletion failed:', storageError);
        }
      }
    } catch (err) {
      alert('Store image removal failed: ' + (err.message || 'Unknown error'));
    } finally {
      setStoreImageRemovingId(null);
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

    const previousStore = stores.find(store => Number(store.id) === Number(storeId)) || editingStore;
    const previousImageUrl = previousStore?.image_url || '';

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

      applyStorePatchLocally(storeId, updatedStore);

      if (previousImageUrl && previousImageUrl !== imageUrl) {
        const previousStoragePath = extractStoragePathFromPublicUrl(previousImageUrl);
        if (previousStoragePath) {
          const { error: storageCleanupError } = await supabase.storage
            .from('product-assets')
            .remove([previousStoragePath]);

          if (storageCleanupError) {
            console.warn('Store image was replaced, but old storage file cleanup failed:', storageCleanupError);
          }
        }
      }
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
      setEditingStore(prev => prev && Number(prev.id) === Number(storeId) ? null : prev);

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
    if (currentStatus === 'Ready to Upload') return 'Ready to Upload';
    if (currentStatus === 'Modified') return 'Modified';
    if (getArray(editedImages).length > 0) return 'Under Review';
    if (getArray(rawImages).length > 0) return 'Processing';
    return 'Missing';
  };

  // --- WORKFLOW STATUS DISPLAY HELPERS ---
  // Database values are kept stable where possible. Some labels are business-facing aliases:
  // Missing = Ready to work, Processing = In Progress, Completed legacy rows = Under Review.
  const WORKFLOW_STATUS_FILTERS = [
    { value: 'All', label: 'All' },
    { value: 'Missing', label: 'Ready to work' },
    { value: 'Processing', label: 'In Progress' },
    { value: 'Under Review', label: 'Under Review' },
    { value: 'Rejected', label: 'Rejected' },
    { value: 'Ready to Upload', label: 'Ready to Upload' },
    { value: 'Modified', label: 'Modified' }
  ];

  const getStatusLabel = (status) => {
    const normalized = status || 'Missing';
    if (normalized === 'Missing') return 'Ready to work';
    if (normalized === 'Processing') return 'In Progress';
    if (normalized === 'Completed') return 'Under Review'; // legacy compatibility
    return normalized;
  };

  const isUnderReviewStatus = (status) => status === 'Under Review' || status === 'Completed';
  const isReadyToUploadStatus = (status) => status === 'Ready to Upload';

  const getStatusBadgeClass = (status) => {
    if (status === 'Ready to Upload') return 'bg-emerald-100 text-emerald-700 border-emerald-200';
    if (isUnderReviewStatus(status)) return 'bg-blue-100 text-blue-700 border-blue-200';
    if (status === 'Modified') return 'bg-[rgba(138,21,56,0.10)] text-[#8a1538] border-[rgba(138,21,56,0.28)]';
    if (status === 'Rejected') return 'bg-red-100 text-red-700 border-red-200';
    if (status === 'Processing') return 'bg-[rgba(138,21,56,0.10)] text-[#8a1538] border-[rgba(138,21,56,0.28)]';
    return 'bg-amber-100 text-amber-700 border-amber-200';
  };

  const moveFullViewImage = (direction) => {
    setFullViewImage(prev => {
      const gallery = Array.isArray(prev?.images) ? prev.images : [];
      if (!prev || gallery.length <= 1) return prev;

      const currentIndex = Number.isInteger(prev.index) ? prev.index : gallery.findIndex(item => item.url === prev.url);
      const safeIndex = currentIndex >= 0 ? currentIndex : 0;
      const nextIndex = (safeIndex + direction + gallery.length) % gallery.length;
      return { ...gallery[nextIndex], images: gallery, index: nextIndex };
    });
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
      .on('postgres_changes', { event: '*', schema: 'public', table: 'task_replies' }, (payload) => {
        if (payload.eventType === 'DELETE') {
          const deletedReplyId = payload.old?.id;
          if (deletedReplyId) {
            setTaskReplies(prev => prev.filter(reply => reply.id !== deletedReplyId));
          }
          return;
        }

        if (payload.new?.id) {
          setTaskReplies(prev => {
            const exists = prev.some(reply => reply.id === payload.new.id);
            const nextReplies = exists
              ? prev.map(reply => reply.id === payload.new.id ? payload.new : reply)
              : [...prev, payload.new];

            return nextReplies.sort((a, b) => new Date(a.created_at || 0) - new Date(b.created_at || 0));
          });
        }
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
    const cleanUserLower = cleanUser.toLowerCase();

    if (cleanUserLower === 'admin' && cleanPass === 'admin123') {
      setAuthRole('Admin');
      setLoginUser('admin');
      setLoginDisplayName('System Administrator');
      setIsLoggedIn(true);
      setActiveTab('home');
      localStorage.setItem('blackrose_user', 'admin');
      localStorage.setItem('blackrose_display_name', 'System Administrator');
      localStorage.setItem('blackrose_role', 'Admin');
      return;
    }

    try {
      const { data, error } = await supabase
        .from('user_registry')
        .select('*')
        .ilike('username', cleanUser)
        .eq('password', cleanPass)
        .limit(1);

      const matchedUser = data?.[0];

      if (error || !matchedUser) {
        setAuthError('Invalid system username or access code.');
      } else {
        const displayName = matchedUser.full_name || matchedUser.name || matchedUser.username;
        setAuthRole(matchedUser.role);
        setLoginUser(matchedUser.username);
        setLoginDisplayName(displayName);
        setIsLoggedIn(true);
        localStorage.setItem('blackrose_user', matchedUser.username);
        localStorage.setItem('blackrose_display_name', displayName);
        localStorage.setItem('blackrose_role', matchedUser.role);
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
    setLoginDisplayName('');
    setLoginPass('');
    setEditingId(null);
    setSelectedProduct(null);
    setSelectedHistoryScope(null);
    setSelectedOperatorStats(null);
    closeManagerPreview();
    setSelectedStoreId('ALL');
    setActiveTab('home');
    localStorage.removeItem('blackrose_user');
    localStorage.removeItem('blackrose_display_name');
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

        const isContentEditorSync = authRole === 'Content Editor' || (authRole !== 'Admin' && authRole !== 'Manager' && checkPermission('excel_upload'));

        if (isContentEditorSync) {
          if (products.length === 0) {
            alert(
              "⚠️ Sheet Sync Failed!\n\n" +
              "Reason: You are using an Editor profile, which updates existing catalog data. Because the database was cleared, there are no products to match.\n\n" +
              "Fix Action: Log out with an authorized account to do a completely fresh structural sheet import first!"
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
          archived_by: loginUser || 'System Account',
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
          `🚀 Success! Processed "${file.name}" for ${activeStoreNameForUpload || 'General Catalog'} with [${sanitized.length}] items. New rows were marked as Ready to work automatically.` +
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
      alert('No Ready to work products found for this sheet.');
      return;
    }

    handleDownloadSheetExport(`Ready-To-Work-${filename}`, [headers, ...missingRows]);
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
    if (!regFullName.trim() || !regUsername.trim() || !regPassword.trim()) {
      alert('Full name, username, and password are required.');
      return;
    }

    try {
      const normalizedUsername = regUsername.trim().toLowerCase();

      const { error } = await supabase
        .from('user_registry')
        .insert([{
          full_name: regFullName.trim(),
          username: normalizedUsername,
          password: regPassword.trim(),
          role: regRole
        }]);

      if (error) throw error;

      alert(`Successfully registered account for "${regFullName.trim()}" with username "${normalizedUsername}"`);
      setRegFullName('');
      setRegUsername('');
      setRegPassword('');
      fetchProducts();
    } catch (err) {
      alert("Registration failed: " + err.message + "\n\nIf this mentions full_name or portal_settings, run the v25 SQL setup first.");
    }
  };


  const openStaffEditPanel = (staffUser) => {
    if (!staffUser?.id) return;

    setEditingStaffId(staffUser.id);
    setStaffEditForm({
      full_name: staffUser.full_name || '',
      username: staffUser.username || '',
      password: staffUser.password || '',
      role: staffUser.role || 'Operator'
    });
    setShowStaffEditPassword(false);
  };

  const cancelStaffEditPanel = () => {
    setEditingStaffId(null);
    setStaffEditForm({
      full_name: '',
      username: '',
      password: '',
      role: 'Operator'
    });
    setShowStaffEditPassword(false);
  };

  const handleUpdateStaffAccount = async (e) => {
    e.preventDefault();

    if (!editingStaffId) return;

    const cleanFullName = staffEditForm.full_name.trim();
    const cleanUsername = staffEditForm.username.trim().toLowerCase();
    const cleanPassword = staffEditForm.password.trim();
    const cleanRole = staffEditForm.role || 'Operator';

    if (!cleanFullName || !cleanUsername || !cleanPassword) {
      alert('Full name, username, and password are required before saving changes.');
      return;
    }

    try {
      const { data, error } = await supabase
        .from('user_registry')
        .update({
          full_name: cleanFullName,
          username: cleanUsername,
          password: cleanPassword,
          role: cleanRole
        })
        .eq('id', editingStaffId)
        .select('*')
        .single();

      if (error) throw error;

      if (data) {
        setUserRegistry(prev => prev.map(user => user.id === data.id ? data : user));
      }

      const updatedCurrentSession = cleanUsername === String(loginUser || '').toLowerCase();
      if (updatedCurrentSession) {
        setLoginUser(cleanUsername);
        setLoginDisplayName(cleanFullName);
        setAuthRole(cleanRole);
        localStorage.setItem('blackrose_user', cleanUsername);
        localStorage.setItem('blackrose_display_name', cleanFullName);
        localStorage.setItem('blackrose_role', cleanRole);
      }

      cancelStaffEditPanel();
      alert('Staff account updated successfully. If that staff member is currently logged in on another device, ask them to log out and log back in so the new access level applies.');
    } catch (err) {
      alert('Staff update failed: ' + (err.message || 'Unknown error'));
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

    const cleanRoleName = newRoleForm.roleName.trim();
    if (!cleanRoleName) return;

    const selectedPermissions = normalizeRolePermissions(newRoleForm.permissions || {});
    const legacyFlags = getLegacyFlagsFromFeaturePermissions(selectedPermissions);

    const formattedRole = {
      roleName: cleanRoleName,
      permissions: {
        ...selectedPermissions,
        ...legacyFlags
      }
    };

    const updatedRolesList = [
      ...customRoles.filter(role => String(role.roleName || '').toLowerCase() !== cleanRoleName.toLowerCase()),
      formattedRole
    ];

    setCustomRoles(updatedRolesList);
    localStorage.setItem('blackrose_custom_roles', JSON.stringify(updatedRolesList));
    setNewRoleForm(createEmptyRoleForm());
    setShowRoleModal(false);
    alert(`Custom permission group "${formattedRole.roleName}" deployed.`);
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

    const completedProducts = matchUserProducts.filter(p => isUnderReviewStatus(p.status));
    const modifiedProducts = matchUserProducts.filter(p => p.status === 'Modified');
    const readyToUploadProducts = matchUserProducts.filter(p => p.status === 'Ready to Upload');
    const completionLikeProducts = [...completedProducts, ...modifiedProducts, ...readyToUploadProducts];
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
      .filter(p => ['Completed', 'Under Review', 'Ready to Upload', 'Modified', 'Rejected'].includes(p.status))
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
      totalReadyToUpload: readyToUploadProducts.length,
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
  const uploadBulkAssetEntries = async (entries, sourceLabel = 'Bulk asset upload') => {
    const zipScopedStoreId = selectedStoreId !== 'ALL' ? Number(selectedStoreId) : null;

    if (!zipScopedStoreId) {
      throw new Error('Select a specific store before uploading image assets. This prevents files from being attached to the wrong store.');
    }

    const normalizeSku = (value) => {
      return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
    };

    const isImageFile = (filename) => {
      return /\.(jpg|jpeg|png|webp|gif|avif)$/i.test(String(filename || ''));
    };

    const safeFileName = (filename) => {
      return String(filename || 'image.jpg')
        .replace(/[^a-zA-Z0-9._-]/g, '_');
    };

    const cleanZipName = (filename) => {
      return String(filename || '')
        .replace(/\.zip$/i, '')
        .trim();
    };

    const { data: productCatalog, error: catalogError } = await supabase
      .from('products')
      .select('*');

    if (catalogError) throw catalogError;

    const catalog = (productCatalog || []).filter(
      p => Number(p.store_id || 0) === zipScopedStoreId
    );

    // Build a safe lookup list for every product.
    // It starts with the product SKU, then adds Barcode/EAN/UPC aliases from the original Excel archive.
    // This avoids forcing photographers to name folders by SKU only. They can name them by SKU or Barcode.
    const productAliasMap = new Map();

    const addAliasToProduct = (productId, value) => {
      const cleanValue = String(value || '').trim();
      if (!productId || !cleanValue) return;
      if (!productAliasMap.has(productId)) productAliasMap.set(productId, new Set());
      productAliasMap.get(productId).add(cleanValue);
    };

    catalog.forEach(product => {
      addAliasToProduct(product.id, product.sku);
      addAliasToProduct(product.id, product.barcode);
      getArray(product.barcodes).forEach(code => addAliasToProduct(product.id, code));
      getArray(product.barcode_aliases).forEach(code => addAliasToProduct(product.id, code));
    });

    const SKU_COLUMN_KEYWORDS = ['sku', 'itemcode', 'itemno', 'itemnumber', 'productcode', 'productid', 'model'];
    const BARCODE_COLUMN_KEYWORDS = ['barcode', 'barcodeno', 'barcodenumber', 'ean', 'upc', 'gtin', 'isbn'];

    const findCatalogProductBySku = (skuValue) => {
      const cleanSku = normalizeSku(skuValue);
      if (!cleanSku) return null;
      return catalog.find(product => normalizeSku(product.sku) === cleanSku) || null;
    };

    const scopedExcelArchives = manifestHistory.filter(record =>
      record?.raw_payload && Number(record.store_id || 0) === zipScopedStoreId
    );

    scopedExcelArchives.forEach(record => {
      const matrix = archivePayloadToMatrix(record.raw_payload);
      if (!matrix || matrix.length < 2) return;

      const headers = matrix[0] || [];
      const skuIndex = findColumnIndexByKeywords(headers, SKU_COLUMN_KEYWORDS);
      const barcodeIndex = findColumnIndexByKeywords(headers, BARCODE_COLUMN_KEYWORDS);

      if (skuIndex < 0 || barcodeIndex < 0) return;

      matrix.slice(1).forEach(row => {
        const rowSku = row?.[skuIndex];
        const rowBarcode = row?.[barcodeIndex];
        const matchedProduct = findCatalogProductBySku(rowSku);
        if (matchedProduct) addAliasToProduct(matchedProduct.id, rowBarcode);
      });
    });

    const getProductLookupValues = (product) => Array.from(productAliasMap.get(product.id) || new Set([product.sku])).filter(Boolean);

    const findMatchingProduct = (skuCandidates) => {
      const cleanedCandidates = Array.from(
        new Set(
          skuCandidates
            .filter(Boolean)
            .map(s => String(s).trim())
        )
      );

      for (const candidate of cleanedCandidates) {
        const candidateNorm = normalizeSku(candidate);
        if (!candidateNorm) continue;

        const exactMatches = catalog.filter(product =>
          getProductLookupValues(product).some(alias => normalizeSku(alias) === candidateNorm)
        );

        // Exact SKU or exact Barcode match is safest. Use it only if it points to one product.
        if (exactMatches.length === 1) return exactMatches[0];
      }

      for (const candidate of cleanedCandidates) {
        const candidateNorm = normalizeSku(candidate);
        if (!candidateNorm) continue;

        const prefixMatches = catalog.filter(product =>
          getProductLookupValues(product).some(alias => {
            const aliasNorm = normalizeSku(alias);
            if (!aliasNorm) return false;

            return (
              candidateNorm.startsWith(aliasNorm) ||
              aliasNorm.startsWith(candidateNorm)
            );
          })
        );

        // Only use prefix match when it is safe and unique.
        if (prefixMatches.length === 1) return prefixMatches[0];
      }

      return null;
    };

    const productUpdates = {};
    const skippedFiles = [];
    const unmatchedSkuCandidates = new Set();
    let uploadedFileCount = 0;

    for (const entry of entries) {
      const relativePath = entry.relativePath || entry.name || '';
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
      const assetType = folderLabel.includes('RAW') ? 'RAW' : 'EDITED';

      const skuCandidates = [];

      if (folderIndex > 0) {
        skuCandidates.push(parts[folderIndex - 1]); // Best match: SKU/Barcode folder directly before RAW/EDITED.
        skuCandidates.push(parts[0]);               // Fallback: root folder.
      }

      const fileNameWithoutExtension = String(originalFileName || '').replace(/\.[^/.]+$/, '');
      if (fileNameWithoutExtension) {
        skuCandidates.push(fileNameWithoutExtension); // Fallback: image file named by SKU/Barcode.
      }

      if (entry.defaultSkuCandidate) {
        skuCandidates.push(entry.defaultSkuCandidate);
        skuCandidates.push(cleanZipName(entry.defaultSkuCandidate));
      }

      const matchedProduct = findMatchingProduct(skuCandidates);

      if (!matchedProduct) {
        unmatchedSkuCandidates.add(skuCandidates.join(' OR ') || relativePath);
        skippedFiles.push(`${relativePath} → skipped, no matching SKU/Barcode found`);
        continue;
      }

      const fileData = await entry.getBlob();
      const uniqueName = `${Date.now()}-${Math.random().toString(36).slice(2)}-${safeFileName(originalFileName)}`;
      const cleanSku = String(matchedProduct.sku || 'UNKNOWN').replace(/[^a-zA-Z0-9_-]/g, '_');
      const storagePath = `stores/${zipScopedStoreId}/${cleanSku}/${assetType}/${uniqueName}`;

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
        edited_image_url: nextEdit,
        updated_at: new Date().toISOString()
      };

      // Important:
      // Bulk uploads are usually done by Photographer/Admin.
      // They must NOT claim the product and must NOT change Ready to work → In Progress.
      // Only an employee/operator changing the status manually should start tracking.
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

    return {
      successCount,
      uploadedFileCount,
      skippedFiles,
      unmatchedSkuCandidates: Array.from(unmatchedSkuCandidates),
      sourceLabel
    };
  };

  // --- BULK MEDIA ZIP EXTRACTOR ---
  const handleBulkZipUpload = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    if (selectedStoreId === 'ALL') {
      alert('Select a specific store before uploading ZIP asset folders. This prevents images from being attached to the wrong store.');
      e.target.value = null;
      return;
    }

    const zipFiles = files.filter(file => /\.zip$/i.test(file.name || ''));

    if (zipFiles.length === 0) {
      alert('Please select one or more .zip files.');
      e.target.value = null;
      return;
    }

    setUploading(true);

    try {
      const allEntries = [];

      for (const file of zipFiles) {
        const zip = new JSZip();
        const contents = await zip.loadAsync(file);
        const zipNameAsSku = String(file.name || '').replace(/\.zip$/i, '').trim();

        Object.entries(contents.files).forEach(([relativePath, zipEntry]) => {
          if (zipEntry.dir) return;

          allEntries.push({
            relativePath,
            defaultSkuCandidate: zipNameAsSku,
            sourceName: file.name,
            getBlob: () => zipEntry.async('blob')
          });
        });
      }

      const result = await uploadBulkAssetEntries(allEntries, `${zipFiles.length} ZIP file(s)`);

      let message = `Bulk asset upload completed for ${result.successCount} product item(s).\nUploaded image files: ${result.uploadedFileCount}`;

      if (zipFiles.length > 1) {
        message += `\nZIP files processed: ${zipFiles.length}`;
      }

      if (result.skippedFiles.length > 0) {
        message += `\nSkipped files: ${result.skippedFiles.length}`;
        console.warn('Bulk ZIP skipped files:', result.skippedFiles);
      }

      if (result.unmatchedSkuCandidates.length > 0) {
        message += `\n\nSome SKU/Barcode names did not match products. Check browser console for details.`;
        console.warn('Bulk ZIP unmatched SKU candidates:', result.unmatchedSkuCandidates);
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

  // --- BULK MEDIA FOLDER EXTRACTOR ---
  const handleBulkFolderUpload = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    if (selectedStoreId === 'ALL') {
      alert('Select a specific store before uploading an asset folder. This prevents images from being attached to the wrong store.');
      e.target.value = null;
      return;
    }

    setUploading(true);

    try {
      const entries = files.map(file => ({
        relativePath: file.webkitRelativePath || file.name,
        defaultSkuCandidate: '',
        sourceName: file.webkitRelativePath || file.name,
        getBlob: () => Promise.resolve(file)
      }));

      const result = await uploadBulkAssetEntries(entries, 'folder upload');

      let message = `Folder asset upload completed for ${result.successCount} product item(s).\nUploaded image files: ${result.uploadedFileCount}`;

      if (result.skippedFiles.length > 0) {
        message += `\nSkipped files: ${result.skippedFiles.length}`;
        console.warn('Folder upload skipped files:', result.skippedFiles);
      }

      if (result.unmatchedSkuCandidates.length > 0) {
        message += `\n\nSome SKU/Barcode names did not match products. Check browser console for details.`;
        console.warn('Folder upload unmatched SKU candidates:', result.unmatchedSkuCandidates);
      }

      alert(message);
      await fetchProducts();
    } catch (err) {
      alert('Folder upload failed: ' + (err.message || 'Unknown error'));
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

    if (liveData.status === 'Ready to Upload' && !isAdminOrManager) {
      alert('This product is already marked Ready to Upload. Only authorized users can change it now.');
      return;
    }

    if (!isAdminOrManager && isOwnedBySomeoneElse) {
      alert(`Action Denied: ${liveData.processed_by} has already claimed this item.`);
      return;
    }

    if (liveData.status === 'Rejected' && targetStatus !== 'Processing' && !isAdminOrManager) {
      alert('Rejected items must be moved back to In Progress before they can be submitted for review.');
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
      updatePayload.processed_by = liveData.processed_by || loginUser || 'System User';
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

    if (targetStatus === 'Under Review') {
      const isModifiedResubmission = Boolean(liveData.rejection_note);
      updatePayload.status = isModifiedResubmission ? 'Modified' : 'Under Review';
      updatePayload.processed_by = liveData.processed_by || loginUser || 'System User';
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

  const handleReadyToUploadProduct = async (id) => {
    if (!(authRole === 'Admin' || authRole === 'Manager')) return;

    if (!window.confirm('Mark this product as Ready to Upload? This confirms the approval step.')) return;

    const nowIso = new Date().toISOString();

    const { data: currentTargetData, error: readError } = await supabase
      .from('products')
      .select('sku, processed_by, claimed_at, total_time_spent')
      .eq('id', id)
      .single();

    if (readError || !currentTargetData) {
      alert('Could not read product before approval.');
      return;
    }

    const previousTrackedSeconds = Number(currentTargetData.total_time_spent) || 0;
    const startedAt = currentTargetData.claimed_at ? new Date(currentTargetData.claimed_at).getTime() : null;
    const elapsedSeconds = startedAt
      ? Math.max(0, Math.floor((Date.now() - startedAt) / 1000))
      : 0;

    const { data: updatedProduct, error } = await supabase
      .from('products')
      .update({
        status: 'Ready to Upload',
        rejection_note: null,
        processed_by: currentTargetData?.processed_by || null,
        total_time_spent: previousTrackedSeconds + elapsedSeconds,
        claimed_at: null,
        updated_at: nowIso
      })
      .eq('id', id)
      .select('*')
      .single();

    if (error) {
      alert('Approval failed: ' + error.message);
      return;
    }

    if (updatedProduct) {
      applyProductPatchLocally(id, updatedProduct);
    }

    closeManagerPreview();
    fetchProducts();
  };

  const closeManagerPreview = () => {
    setManagerPreview(null);
    setIsRejecting(false);
    setRejectNote('');
    setShowManagerCompare(false);
    setFullCompareIndex(null);
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
        alert('Upload locked. First change the product status to In Progress so it is claimed under your name.');
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
        updated_at: new Date().toISOString()
      };

      // Important:
      // Image upload should not claim the product and should not change status.
      // Status/time tracking must start only when an employee manually changes Ready to work → In Progress.
      // If the product is already In Progress, it naturally stays Processing because we do not overwrite it here.

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

  const extractStoragePathFromPublicUrl = (url) => {
    if (!url) return null;

    try {
      const parsed = new URL(url);
      const marker = '/storage/v1/object/public/product-assets/';
      const markerIndex = parsed.pathname.indexOf(marker);
      if (markerIndex === -1) return null;
      return decodeURIComponent(parsed.pathname.slice(markerIndex + marker.length));
    } catch (err) {
      return null;
    }
  };

  const handleRemoveIndividualImage = async (id, fieldType, urlToRemove) => {
    const dbField = fieldType === 'raw' ? 'raw_image_url' : 'edited_image_url';
    const folderLabel = fieldType === 'raw' ? 'RAW' : 'EDITED';

    try {
      const { data: existingProd, error: readError } = await supabase
        .from('products')
        .select('id, sku, status, processed_by, raw_image_url, edited_image_url')
        .eq('id', id)
        .single();

      if (readError || !existingProd) {
        alert('Could not read the latest product record before removing the image.');
        return;
      }

      const isAdminOrManager = authRole === 'Admin' || authRole === 'Manager';
      const isReadyToUploadLocked = existingProd.status === 'Ready to Upload';
      const isAssignedEmployee = existingProd.processed_by === loginUser;
      const canEmployeeRemoveEdited =
        fieldType === 'edited' &&
        isAssignedEmployee &&
        ['Processing', 'Rejected'].includes(existingProd.status);

      if (isReadyToUploadLocked && !isAdminOrManager) {
        alert('This product is already Ready to Upload. Only authorized users can remove images now.');
        return;
      }

      if (!isAdminOrManager && !canEmployeeRemoveEdited) {
        alert('You can only remove EDITED images from products assigned to you while they are In Progress or Rejected.');
        return;
      }

      if (!window.confirm(`Remove this ${folderLabel} image from SKU ${existingProd.sku || 'UNKNOWN'}?`)) return;

      const currentArray = getArray(existingProd[dbField]);
      const filteredArray = currentArray.filter(url => url !== urlToRemove);

      const { data: updatedProduct, error: updateError } = await supabase
        .from('products')
        .update({ [dbField]: filteredArray, updated_at: new Date().toISOString() })
        .eq('id', id)
        .select('*')
        .single();

      if (updateError) {
        alert('Image removal failed: ' + updateError.message);
        return;
      }

      const storagePath = extractStoragePathFromPublicUrl(urlToRemove);
      if (storagePath) {
        const { error: storageError } = await supabase.storage
          .from('product-assets')
          .remove([storagePath]);

        if (storageError) {
          console.warn('Image was removed from the product record, but storage file deletion failed:', storageError);
        }
      }

      if (updatedProduct) {
        applyProductPatchLocally(id, updatedProduct);
      }
    } catch (err) {
      alert('Image removal failed: ' + (err.message || 'Unknown error'));
    }
  };

  // --- TASK BOARD MANAGEMENT METHODS ---
  const isTaskManager = authRole === 'Admin' || authRole === 'Manager' || checkPermission('tasks_manage');
  const canViewTasks = isTaskManager || checkPermission('tasks_view');
  const canReplyToTasks = isTaskManager || checkPermission('tasks_reply');

  // Staff only see tasks meant for them. Managers/Admins/supervisors see every task so
  // specific employee assignments do not disappear from the manager portal.
  const normalizeTaskTarget = (value) => String(value || 'All').trim().toLowerCase();

  const isTaskVisibleToCurrentUser = (task) => {
    if (!canViewTasks) return false;
    if (isTaskManager) return true;

    const target = normalizeTaskTarget(task.assigned_role);
    if (target === 'all') return true;
    if (target === normalizeTaskTarget(authRole)) return true;
    if (target === normalizeTaskTarget(loginUser)) return true;
    if (target === normalizeTaskTarget(loginDisplayName)) return true;
    return false;
  };

  const visibleTasks = tasks.filter(isTaskVisibleToCurrentUser);

  const taskStatusFilterOptions = ['All', 'Open', 'In Progress', 'Done', 'Archived'];
  const taskPriorityFilterOptions = ['All', 'Urgent', 'High', 'Normal', 'Low'];
  const taskTargetFilterOptions = Array.from(new Set([
    'All',
    'Operator',
    'Photographer',
    'Content Editor',
    'Manager',
    'Admin',
    ...tasks.map(task => task.assigned_role || 'All'),
    ...userRegistry.map(user => user.username).filter(Boolean)
  ])).filter(Boolean);

  const filteredVisibleTasks = visibleTasks.filter(task => {
    const normalizedTaskTarget = normalizeTaskTarget(task.assigned_role || 'All');
    const searchLower = taskSearchQuery.toLowerCase().trim();

    const matchesStatus = taskStatusFilter === 'All' || (task.status || 'Open') === taskStatusFilter;
    const matchesPriority = taskPriorityFilter === 'All' || (task.priority || 'Normal') === taskPriorityFilter;
    const matchesTarget = taskTargetFilter === 'All' || normalizedTaskTarget === normalizeTaskTarget(taskTargetFilter);
    const matchesSearch = searchLower === '' || [
      task.title,
      task.description,
      task.created_by,
      task.updated_by,
      task.assigned_role,
      task.priority,
      task.status
    ].some(value => String(value || '').toLowerCase().includes(searchLower));

    return matchesStatus && matchesPriority && matchesTarget && matchesSearch;
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
        updated_by: loginUser || 'System User',
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
          .insert([{ ...payload, created_by: loginUser || 'System User', status: 'Open' }])
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
    if (!window.confirm('Delete this task permanently? This will also remove all replies attached to it.')) return;

    try {
      const { error: repliesDeleteError } = await supabase
        .from('task_replies')
        .delete()
        .eq('task_id', taskId);

      if (repliesDeleteError) throw repliesDeleteError;

      const { error } = await supabase.from('task_board').delete().eq('id', taskId);
      if (error) throw error;
      setTasks(prev => prev.filter(task => task.id !== taskId));
      setTaskReplies(prev => prev.filter(reply => reply.task_id !== taskId));
      setTaskReplyDrafts(prev => {
        const nextDrafts = { ...prev };
        delete nextDrafts[taskId];
        return nextDrafts;
      });
      setExpandedTaskReplyIds(prev => {
        const nextExpanded = { ...prev };
        delete nextExpanded[taskId];
        return nextExpanded;
      });
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

  const createClientUuid = () => {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
      return crypto.randomUUID();
    }

    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (char) => {
      const rand = Math.floor(Math.random() * 16);
      const value = char === 'x' ? rand : (rand & 0x3) | 0x8;
      return value.toString(16);
    });
  };

  const isUuidString = (value) => {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || ''));
  };

  const getCurrentUserRegistryUuid = () => {
    const matchedUser = userRegistry.find(user =>
      normalizeTaskTarget(user.username) === normalizeTaskTarget(loginUser) ||
      normalizeTaskTarget(user.full_name) === normalizeTaskTarget(loginDisplayName)
    );

    return isUuidString(matchedUser?.id) ? matchedUser.id : null;
  };

  const getTaskRepliesForTask = (taskId) => {
    return taskReplies
      .filter(reply => reply.task_id === taskId)
      .sort((a, b) => new Date(a.created_at || 0) - new Date(b.created_at || 0));
  };

  const getTaskReplyCount = (taskId) => getTaskRepliesForTask(taskId).length;

  const setTaskReplyDraft = (taskId, value) => {
    setTaskReplyDrafts(prev => ({ ...prev, [taskId]: value }));
  };

  const toggleTaskReplies = (taskId) => {
    setExpandedTaskReplyIds(prev => ({ ...prev, [taskId]: !(prev[taskId] ?? true) }));
  };

  const upsertTaskReplyLocally = (replyRow) => {
    if (!replyRow?.id) return;

    setTaskReplies(prev => {
      const exists = prev.some(reply => reply.id === replyRow.id);
      const nextReplies = exists
        ? prev.map(reply => reply.id === replyRow.id ? replyRow : reply)
        : [...prev, replyRow];

      return nextReplies.sort((a, b) => new Date(a.created_at || 0) - new Date(b.created_at || 0));
    });
  };

  const handleTaskReplySubmit = async (taskId) => {
    if (!taskId) return;

    if (!canReplyToTasks) {
      alert('You do not have permission to reply to tasks.');
      return;
    }

    const cleanMessage = String(taskReplyDrafts[taskId] || '').trim();
    if (!cleanMessage) return;

    setTaskReplySavingId(taskId);
    try {
      const nowIso = new Date().toISOString();
      const payload = {
        id: createClientUuid(),
        task_id: taskId,
        user_name: loginDisplayName || loginUser || 'System User',
        user_id: getCurrentUserRegistryUuid(),
        message: cleanMessage,
        created_at: nowIso
      };

      const { data, error } = await supabase
        .from('task_replies')
        .insert([payload])
        .select('*')
        .single();

      if (error) throw error;

      upsertTaskReplyLocally(data || payload);
      setTaskReplyDrafts(prev => ({ ...prev, [taskId]: '' }));
      setExpandedTaskReplyIds(prev => ({ ...prev, [taskId]: true }));
      setTaskReplyError(null);
    } catch (err) {
      alert('Reply send failed: ' + (err.message || 'Unknown error'));
      setTaskReplyError(err.message || 'Reply send failed.');
    } finally {
      setTaskReplySavingId(null);
    }
  };

  const TASK_REPLY_DELETE_WINDOW_MINUTES = 15;
  const TASK_REPLY_DELETE_WINDOW_MS = TASK_REPLY_DELETE_WINDOW_MINUTES * 60 * 1000;

  const isOwnTaskReply = (reply) => {
    const replyName = normalizeTaskTarget(reply?.user_name);
    return replyName === normalizeTaskTarget(loginUser) || replyName === normalizeTaskTarget(loginDisplayName);
  };

  const getTaskReplyDeleteInfo = (reply, forceNowMs = null) => {
    const createdAtMs = new Date(reply?.created_at || '').getTime();
    const hasValidCreatedAt = !Number.isNaN(createdAtMs);
    const nowMs = typeof forceNowMs === 'number'
      ? forceNowMs
      : dashboardClock instanceof Date
        ? dashboardClock.getTime()
        : Date.now();

    const ageMs = hasValidCreatedAt ? Math.max(0, nowMs - createdAtMs) : TASK_REPLY_DELETE_WINDOW_MS + 1;
    const isExpired = !hasValidCreatedAt || ageMs > TASK_REPLY_DELETE_WINDOW_MS;
    const isOwnReplyMessage = isOwnTaskReply(reply);

    // Replies can only be deleted during this time window. Change
    // TASK_REPLY_DELETE_WINDOW_MINUTES above if you want 5, 10, 30, etc. minutes instead.
    const canDelete = (isTaskManager || isOwnReplyMessage) && !isExpired;
    const remainingMs = hasValidCreatedAt ? Math.max(0, TASK_REPLY_DELETE_WINDOW_MS - ageMs) : 0;
    const remainingMinutes = Math.max(0, Math.ceil(remainingMs / 60000));
    const deleteUntilLabel = hasValidCreatedAt
      ? new Date(createdAtMs + TASK_REPLY_DELETE_WINDOW_MS).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
      : 'the delete window';

    return {
      canDelete,
      isExpired,
      isOwnReplyMessage,
      remainingMinutes,
      deleteUntilLabel,
      reason: isExpired
        ? `Delete window expired. Replies can only be deleted for ${TASK_REPLY_DELETE_WINDOW_MINUTES} minutes after posting.`
        : 'Only the sender or an authorized manager can delete this reply during the delete window.'
    };
  };

  const handleDeleteTaskReply = async (reply) => {
    if (!reply?.id) return;

    const deleteInfo = getTaskReplyDeleteInfo(reply, Date.now());
    if (!deleteInfo.canDelete) {
      alert(deleteInfo.reason);
      return;
    }

    if (!window.confirm('Delete this reply?')) return;

    try {
      const { error } = await supabase
        .from('task_replies')
        .delete()
        .eq('id', reply.id);

      if (error) throw error;
      setTaskReplies(prev => prev.filter(existingReply => existingReply.id !== reply.id));
    } catch (err) {
      alert('Reply delete failed: ' + (err.message || 'Unknown error'));
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
    const matchesStatus = statusFilter === 'All' ||
      (statusFilter === 'Under Review' ? isUnderReviewStatus(prod.status) : prod.status === statusFilter);
    const matchesStore = isProductInSelectedStore(prod);
    const matchesSheetContext = !selectedHistoryScope || prod.sheet_reference === selectedHistoryScope;
    
    const matchesOperatorBound = authRole === 'Admin' || authRole === 'Manager' || 
      authRole === 'Content Editor' || checkPermission('excel_upload') || checkPermission('product_edit') || checkPermission('sheet_exports') || checkPermission('ad_hoc') ||
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

  const assetDirectoryProductIds = assetDirectoryProducts.map(prod => prod.id);
  const selectedAssetDirectoryProducts = products.filter(prod => selectedAssetProductIds.includes(prod.id) && (getArray(prod.raw_image_url).filter(Boolean).length > 0 || getArray(prod.edited_image_url).filter(Boolean).length > 0));
  const selectedAssetRawCount = selectedAssetDirectoryProducts.reduce((sum, prod) => sum + getArray(prod.raw_image_url).filter(Boolean).length, 0);
  const selectedAssetEditedCount = selectedAssetDirectoryProducts.reduce((sum, prod) => sum + getArray(prod.edited_image_url).filter(Boolean).length, 0);
  const allVisibleAssetProductsSelected = assetDirectoryProducts.length > 0 && assetDirectoryProductIds.every(id => selectedAssetProductIds.includes(id));

  const toggleAssetProductSelection = (productId) => {
    setSelectedAssetProductIds(prev =>
      prev.includes(productId)
        ? prev.filter(id => id !== productId)
        : [...prev, productId]
    );
  };

  const handleSelectAllVisibleAssetProducts = () => {
    setSelectedAssetProductIds(prev => {
      if (allVisibleAssetProductsSelected) {
        return prev.filter(id => !assetDirectoryProductIds.includes(id));
      }
      return Array.from(new Set([...prev, ...assetDirectoryProductIds]));
    });
  };

  const clearAssetProductSelection = () => setSelectedAssetProductIds([]);

  const metrics = {
    total: scopedProducts.filter(p => !selectedHistoryScope || p.sheet_reference === selectedHistoryScope).length,
    missing: scopedProducts.filter(p => p.status === 'Missing' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    processing: scopedProducts.filter(p => p.status === 'Processing' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    underReview: scopedProducts.filter(p => isUnderReviewStatus(p.status) && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    rejected: scopedProducts.filter(p => p.status === 'Rejected' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    readyToUpload: scopedProducts.filter(p => isReadyToUploadStatus(p.status) && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length,
    modified: scopedProducts.filter(p => p.status === 'Modified' && (!selectedHistoryScope || p.sheet_reference === selectedHistoryScope)).length
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
      underReview: storeProducts.filter(p => isUnderReviewStatus(p.status)).length,
      rejected: storeProducts.filter(p => p.status === 'Rejected').length,
      readyToUpload: storeProducts.filter(p => isReadyToUploadStatus(p.status)).length,
      modified: storeProducts.filter(p => p.status === 'Modified').length
    };
  });

  const unassignedProductsCount = products.filter(p => !p.store_id).length;

  const canViewMyPerformance = isLoggedIn && authRole !== 'Admin' && authRole !== 'Manager';
  const selfPerformanceStats = canViewMyPerformance ? compileUserPerformanceMetrics(loginUser) : null;

  const currentStaffName = loginDisplayName || loginUser || 'Staff Member';

  const getUserDisplayName = (username) => {
    if (!username) return '';
    const match = userRegistry.find(user => user.username === username || user.full_name === username);
    return match?.full_name || match?.username || username;
  };

  const formatTaskTargetLabel = (value) => {
    const target = String(value || 'All').trim();
    if (!target || target.toLowerCase() === 'all') return 'All Staff';
    if (target === 'Operator') return 'Workflow Staff';
    if (target === 'Manager') return 'Review Leads';
    if (target === 'Admin') return 'Full Access';
    if (target === 'Photographer') return 'Media Staff';
    if (target === 'Content Editor') return 'Sheet Staff';
    return getUserDisplayName(target);
  };

  const renderTaskRepliesPanel = (task) => {
    if (!task?.id) return null;

    const repliesForTask = getTaskRepliesForTask(task.id);
    const isExpanded = expandedTaskReplyIds[task.id] ?? true;
    const draftValue = taskReplyDrafts[task.id] || '';
    const isSavingReply = taskReplySavingId === task.id;

    return (
      <div className="mt-4 border-t border-gray-100 pt-4">
        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 mb-3">
          <button
            type="button"
            onClick={() => toggleTaskReplies(task.id)}
            className="w-fit px-3 py-1.5 rounded-xl bg-indigo-50 text-indigo-700 border border-indigo-100 text-[10px] font-black uppercase tracking-wider hover:bg-indigo-100 transition-colors"
          >
            💬 Replies / Questions ({repliesForTask.length}) {isExpanded ? '▲' : '▼'}
          </button>
          <span className="text-[10px] font-bold uppercase tracking-wider text-gray-400">Internal task chat</span>
        </div>

        {isExpanded && (
          <div className="space-y-3">
            <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
              {repliesForTask.length === 0 ? (
                <div className="p-4 rounded-xl border border-dashed border-gray-200 bg-gray-50 text-center text-[11px] font-bold text-gray-400 uppercase">
                  No replies yet. Ask a question or send an update below.
                </div>
              ) : (
                repliesForTask.map((reply) => {
                  const isOwnReplyMessage = isOwnTaskReply(reply);
                  const deleteInfo = getTaskReplyDeleteInfo(reply);
                  const showDeleteControl = isTaskManager || isOwnReplyMessage;
                  return (
                    <div
                      key={reply.id}
                      className={`p-3 rounded-2xl border ${isOwnReplyMessage ? 'bg-[rgba(138,21,56,0.06)] border-[rgba(138,21,56,0.18)]' : 'bg-gray-50 border-gray-200'}`}
                    >
                      <div className="flex items-start justify-between gap-3 mb-1">
                        <div className="min-w-0">
                          <span className={`text-[10px] font-black uppercase tracking-wider ${isOwnReplyMessage ? 'text-[#8a1538]' : 'text-gray-700'}`}>
                            {getUserDisplayName(reply.user_name) || reply.user_name || 'Staff Member'}
                          </span>
                          {isOwnReplyMessage && (
                            <span className="ml-2 text-[9px] font-black uppercase tracking-wider text-[#8a1538] bg-white/70 border border-[rgba(138,21,56,0.18)] px-1.5 py-0.5 rounded">
                              You
                            </span>
                          )}
                          <div className="text-[9px] font-bold text-gray-400 mt-0.5">
                            {formatDisplayDateTime(reply.created_at)}
                            {showDeleteControl && !deleteInfo.isExpired && (
                              <span className="ml-2 text-[#8a1538]">Delete available for {deleteInfo.remainingMinutes}m</span>
                            )}
                            {showDeleteControl && deleteInfo.isExpired && (
                              <span className="ml-2 text-gray-400">Delete locked after {TASK_REPLY_DELETE_WINDOW_MINUTES}m</span>
                            )}
                          </div>
                        </div>
                        {showDeleteControl && (
                          <button
                            type="button"
                            onClick={() => handleDeleteTaskReply(reply)}
                            disabled={!deleteInfo.canDelete}
                            title={deleteInfo.canDelete ? `Delete reply before ${deleteInfo.deleteUntilLabel}` : deleteInfo.reason}
                            className={`text-[9px] font-black uppercase tracking-wider rounded-lg px-2 py-1 border ${deleteInfo.canDelete ? 'text-red-500 hover:text-red-700 bg-white border-red-100 hover:bg-red-50' : 'text-gray-400 bg-gray-100 border-gray-200 cursor-not-allowed'}`}
                          >
                            {deleteInfo.canDelete ? `Delete ${deleteInfo.remainingMinutes}m` : 'Locked'}
                          </button>
                        )}
                      </div>
                      <p className="text-xs sm:text-sm text-gray-700 font-medium leading-relaxed whitespace-pre-wrap">
                        {reply.message}
                      </p>
                    </div>
                  );
                })
              )}
            </div>

            {canReplyToTasks ? (
              <form
                onSubmit={(event) => {
                  event.preventDefault();
                  handleTaskReplySubmit(task.id);
                }}
                className="bg-gray-50 border border-gray-200 rounded-2xl p-3 space-y-2"
              >
                <textarea
                  value={draftValue}
                  onChange={(event) => setTaskReplyDraft(task.id, event.target.value)}
                  placeholder="Reply to this task, ask for clarification, or send a progress update..."
                  rows="2"
                  className="w-full px-3 py-2 border border-gray-200 rounded-xl bg-white text-xs text-gray-900 outline-none resize-none focus:border-indigo-300"
                />
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                  <span className="text-[10px] font-bold text-gray-400">
                    Posting as {currentStaffName}
                  </span>
                  <button
                    type="submit"
                    disabled={isSavingReply || !draftValue.trim()}
                    className="px-4 py-2 bg-indigo-700 hover:bg-indigo-800 disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                  >
                    {isSavingReply ? 'Sending...' : 'Send Reply'}
                  </button>
                </div>
              </form>
            ) : (
              <div className="bg-gray-50 border border-gray-200 rounded-2xl p-3 text-[11px] font-bold text-gray-400 uppercase text-center">
                Reply permission is not enabled for your access group.
              </div>
            )}
          </div>
        )}
      </div>
    );
  };

  const permissionFeatures = PORTAL_PERMISSION_FEATURES;

  const permissionFeaturesByCategory = permissionFeatures.reduce((groups, feature) => {
    const category = feature.category || 'Other';
    if (!groups[category]) groups[category] = [];
    groups[category].push(feature);
    return groups;
  }, {});

  const selectedNewRolePermissionCount = Object.values(newRoleForm.permissions || {}).filter(Boolean).length;

  const setNewRolePermissionValue = (featureKey, checked) => {
    setNewRoleForm(prev => ({
      ...prev,
      permissions: {
        ...createPortalPermissionSet(false),
        ...(prev.permissions || {}),
        [featureKey]: checked
      }
    }));
  };

  const setAllNewRolePermissions = (checked) => {
    setNewRoleForm(prev => ({
      ...prev,
      permissions: createPortalPermissionSet(checked)
    }));
  };

  const hasMatrixPermission = (roleName, featureKey) => roleAllowsPortalFeature(roleName, customRoles, featureKey);

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
                   Signed in: <span className="underline font-black">{currentStaffName}</span>
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
                  {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('performance')) && (
                    <button onClick={() => setActiveTab('operators')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'operators' ? 'bg-teal-700 text-white border-teal-700' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="Staff Activity">👥</button>
                  )}
                  {canViewMyPerformance && selfPerformanceStats && (
                    <button onClick={() => setActiveTab('my_performance')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'my_performance' ? 'bg-amber-600 text-white border-amber-600' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="My Performance">📈</button>
                  )}
                  {authRole === 'Admin' && (
                    <button onClick={() => setActiveTab('admin_panel')} className={`w-full h-10 rounded-xl text-sm border ${activeTab === 'admin_panel' ? 'bg-[rgba(138,21,56,0.85)] text-white border-[rgba(138,21,56,0.85)]' : 'bg-white hover:bg-gray-100 border-gray-100'}`} title="Control Panel">👑</button>
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

                  {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('performance')) && (
                    <button 
                      onClick={() => setActiveTab('operators')} 
                      className={`w-full text-left px-3 py-2 text-xs font-black uppercase tracking-wide rounded-lg transition-all ${activeTab === 'operators' ? 'bg-teal-700 text-white shadow-xs' : 'text-gray-700 hover:bg-gray-100'}`}
                    >
                      👥 Staff Activity
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
                      👑 Control Panel
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
                              <button
                                type="button"
                                onClick={(event) => {
                                  event.stopPropagation();
                                  openStoreEditModal(store);
                                }}
                                className={`px-2 py-1 rounded-md text-center text-[9px] font-black uppercase border transition-all ${isActiveStore ? 'bg-white/10 text-white border-white/20 hover:bg-white/20' : 'bg-gray-50 text-gray-500 border-gray-100 hover:bg-gray-100'} cursor-pointer`}
                                title="Edit store name, image, or remove the image"
                              >
                                Edit
                              </button>
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
                <span className="text-[10px] font-mono text-gray-400">v4.5</span>
              ) : (
                <span className="text-[10px] font-mono text-gray-400">v4.5 Store Editing</span>
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
                      <div className="relative">
                        <input
                          type={showLoginPassword ? "text" : "password"}
                          required
                          value={loginPass}
                          onChange={(e) => setLoginPass(e.target.value)}
                          placeholder="•••••••••••••"
                          className="w-full px-4 py-3 pr-20 border border-gray-200 rounded-xl bg-gray-50 text-sm text-gray-900 outline-none"
                        />
                        <button
                          type="button"
                          onClick={() => setShowLoginPassword(prev => !prev)}
                          className="absolute right-3 top-1/2 -translate-y-1/2 text-[10px] font-black uppercase tracking-wider text-[#8a1538] hover:text-[rgba(138,21,56,0.95)]"
                        >
                          {showLoginPassword ? 'Hide' : 'Show'}
                        </button>
                      </div>
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
                            Welcome • {currentStaffName}
                          </span>
                          <h1 className="text-3xl font-black mt-3 tracking-tight">Stores & Task Command Center</h1>
                          <p className="text-sm text-slate-300 font-medium mt-1">
                            Black Rose: Redefining Management Excellence.
                          </p>
                        </div>
                        <div className="flex flex-col sm:items-end gap-3">
                          <div className="text-right text-base font-bold text-slate-200">
                            {dashboardClock.toLocaleDateString([], { weekday: 'short', year: 'numeric', month: 'short', day: '2-digit' })} • {dashboardClock.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                          </div>
                          <div className="flex flex-wrap gap-2 justify-start sm:justify-end">
                            <button onClick={openGoogleSheetLink} className="px-4 py-2 bg-white/10 hover:bg-white/20 border border-white/20 rounded-xl text-xs font-black uppercase tracking-wider">
                              Google Sheets
                            </button>
                            {(authRole === 'Admin' || authRole === 'Manager') && (
                              <button
                                onClick={() => { setGoogleSheetDraft(googleSheetLink || ''); setShowGoogleSheetModal(true); }}
                                className="px-4 py-2 bg-white/10 hover:bg-white/20 border border-white/20 rounded-xl text-xs font-black uppercase tracking-wider"
                              >
                                {googleSheetLink ? 'Edit Sheet Link' : 'Set Sheet Link'}
                              </button>
                            )}
                            <button onClick={() => setActiveTab('task_board')} className="px-4 py-2 bg-white/10 hover:bg-white/20 border border-white/20 rounded-xl text-xs font-black uppercase tracking-wider">
                              Open Tasks
                            </button>
                            {(authRole === 'Admin' || authRole === 'Manager') && (
                              <button onClick={() => setActiveTab('operators')} className="px-4 py-2 bg-white/10 hover:bg-white/20 border border-white/20 rounded-xl text-xs font-black uppercase tracking-wider">
                                Staff
                              </button>
                            )}
                            {authRole === 'Admin' && (
                              <button onClick={() => setActiveTab('admin_panel')} className="px-4 py-2 bg-[#8a1538] hover:bg-[#8a1538] border border-[rgba(138,21,56,0.50)] rounded-xl text-xs font-black uppercase tracking-wider">
                                Control Panel
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
                              Pick a store to open its dashboard or continue work.
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
                              {(authRole === 'Admin' || authRole === 'Manager') ? 'Add a store above, then upload its Excel sheet from the store dashboard.' : 'Ask an authorized user to create a store and upload its Excel sheet.'}
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
                                        <span className="block text-[9px] uppercase text-amber-700 font-black">Ready</span>
                                        <b className="block text-xl font-black text-amber-800 leading-tight">{store.missing}</b>
                                      </div>
                                      <div className="bg-[rgba(138,21,56,0.06)] border border-[rgba(138,21,56,0.28)] rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-[#8a1538] font-black">Progress</span>
                                        <b className="block text-xl font-black text-[rgba(138,21,56,0.95)] leading-tight">{store.processing}</b>
                                      </div>
                                      <div className="bg-red-50 border border-red-200 rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-red-700 font-black">Reject</span>
                                        <b className="block text-xl font-black text-red-800 leading-tight">{store.rejected}</b>
                                      </div>
                                    </div>

                                    <div className="grid grid-cols-3 gap-2 mt-2 text-center">
                                      <div className="bg-blue-50 border border-blue-200 rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-blue-700 font-black">Review</span>
                                        <b className="block text-lg font-black text-blue-800 leading-tight">{store.underReview}</b>
                                      </div>
                                      <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-2">
                                        <span className="block text-[9px] uppercase text-emerald-700 font-black">Upload</span>
                                        <b className="block text-lg font-black text-emerald-800 leading-tight">{store.readyToUpload}</b>
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
                                    <button
                                      type="button"
                                      onClick={(event) => {
                                        event.stopPropagation();
                                        openStoreEditModal(store);
                                      }}
                                      className="px-3 py-2 rounded-xl text-center text-[10px] font-black uppercase cursor-pointer border bg-[rgba(138,21,56,0.06)] text-[#8a1538] border-[rgba(138,21,56,0.18)] hover:bg-[rgba(138,21,56,0.10)]"
                                    >
                                      Edit Store
                                    </button>
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
                                  <option value="Operator">Workflow Staff</option>
                                  <option value="Photographer">Media Staff</option>
                                  <option value="Content Editor">Sheet Staff</option>
                                  {userRegistry.map(user => (
                                    <option key={user.id} value={user.username}>{user.full_name ? `${user.full_name} (@${user.username})` : user.username}</option>
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
                              <p className="text-xs text-gray-400 font-bold mt-1">Latest active instructions.</p>
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
                                  <div className="flex items-center justify-between gap-2 mt-2">
                                    <div className="text-[10px] text-gray-400 font-bold">Due: {formatDisplayDateTime(task.due_at)}</div>
                                    <span className="text-[10px] text-indigo-700 bg-indigo-50 border border-indigo-100 px-2 py-0.5 rounded-full font-black uppercase">
                                      💬 {getTaskReplyCount(task.id)}
                                    </span>
                                  </div>
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
                          <h1 className="text-3xl font-black mt-3 tracking-tight">System Control Hub</h1>
                          <p className="text-sm text-slate-300 font-medium mt-1">Deploy custom permission rules, add new team accounts, and supervise global workflows seamlessly.</p>
                        </div>
                        <button 
                          onClick={() => setShowRoleModal(true)}
                          className="shrink-0 bg-[rgba(138,21,56,0.85)] hover:bg-[rgba(138,21,56,0.85)] text-white font-black text-xs uppercase tracking-wider px-6 py-3.5 rounded-xl shadow-lg transition-all border border-[#8a1538] cursor-pointer"
                        >
                          ⚡ Create Permission Group
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
                              <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Staff Full Name *</label>
                              <input
                                type="text" required value={regFullName} onChange={(e) => setRegFullName(e.target.value)} placeholder="e.g. Hussam Ahmed"
                                className="w-full px-3 py-2 border rounded-xl text-xs bg-gray-50 outline-none text-gray-900 font-medium focus:bg-white"
                              />
                            </div>
                            <div>
                              <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Assign Username *</label>
                              <input 
                                type="text" required value={regUsername} onChange={(e) => setRegUsername(e.target.value.toLowerCase())} placeholder="e.g. hussam_rose" 
                                className="w-full px-3 py-2 border rounded-xl text-xs bg-gray-50 outline-none text-gray-900 font-medium focus:bg-white"
                              />
                            </div>
                            <div>
                              <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Assign Password</label>
                              <div className="relative">
                                <input 
                                  type={showRegPassword ? "text" : "password"}
                                  required
                                  value={regPassword}
                                  onChange={(e) => setRegPassword(e.target.value)}
                                  placeholder="••••••••" 
                                  className="w-full px-3 py-2 pr-16 border rounded-xl text-xs bg-gray-50 outline-none text-gray-900 font-medium focus:bg-white"
                                />
                                <button
                                  type="button"
                                  onClick={() => setShowRegPassword(prev => !prev)}
                                  className="absolute right-3 top-1/2 -translate-y-1/2 text-[9px] font-black uppercase text-[#8a1538]"
                                >
                                  {showRegPassword ? 'Hide' : 'Show'}
                                </button>
                              </div>
                            </div>
                            <div>
                              <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Assign Access Profile</label>
                              <select value={regRole} onChange={(e) => setRegRole(e.target.value)} className="w-full px-3 py-2 border rounded-xl text-xs bg-white text-gray-800 outline-none font-bold uppercase">
                                <option value="Operator">Workflow Staff</option>
                                <option value="Manager">Review Access</option>
                                <option value="Photographer">Media Staff</option>
                                <option value="Content Editor">Sheet Staff</option>
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
                          {userRegistry.map(user => {
                            const isEditingThisStaff = editingStaffId === user.id;
                            return (
                              <div key={user.id} className="p-4 hover:bg-gray-50/50 transition-colors">
                                {!isEditingThisStaff ? (
                                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                                    <div className="flex items-center gap-3 min-w-0">
                                      <div className="w-8 h-8 rounded-full bg-[rgba(138,21,56,0.85)] text-white flex items-center justify-center font-bold text-xs shrink-0">👤</div>
                                      <div className="min-w-0">
                                        <div className="text-xs font-bold text-gray-900 truncate">{user.full_name || user.username}</div>
                                        <div className="text-[10px] text-gray-400 mt-0.5 truncate">@{user.username} • Created: {new Date(user.created_at).toLocaleDateString()}</div>
                                      </div>
                                    </div>
                                    <div className="flex items-center gap-2 shrink-0">
                                      <button
                                        type="button"
                                        onClick={() => openStaffEditPanel(user)}
                                        className="text-xxs font-black text-[#8a1538] border border-[rgba(138,21,56,0.18)] hover:bg-[rgba(138,21,56,0.06)] px-3 py-1 rounded-lg uppercase tracking-wider cursor-pointer"
                                      >
                                        Edit
                                      </button>
                                      <button 
                                        onClick={() => handleRevokeStaffAccess(user.id, user.username)}
                                        className="text-xxs font-black text-red-600 border border-red-100 hover:bg-red-50 px-3 py-1 rounded-lg uppercase tracking-wider cursor-pointer"
                                      >
                                        Revoke Access
                                      </button>
                                    </div>
                                  </div>
                                ) : (
                                  <form onSubmit={handleUpdateStaffAccount} className="space-y-4 bg-[rgba(138,21,56,0.04)] border border-[rgba(138,21,56,0.16)] rounded-2xl p-4">
                                    <div className="flex items-center justify-between gap-3">
                                      <div>
                                        <h4 className="text-xs font-black text-gray-900 uppercase tracking-wider">Edit Staff Account</h4>
                                        <p className="text-[10px] text-gray-400 font-bold mt-0.5">Update name, username, password, or access profile.</p>
                                      </div>
                                      <button type="button" onClick={cancelStaffEditPanel} className="text-gray-400 hover:text-gray-900 text-lg font-black">✕</button>
                                    </div>

                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                                      <div>
                                        <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Full Name *</label>
                                        <input
                                          type="text"
                                          required
                                          value={staffEditForm.full_name}
                                          onChange={(e) => setStaffEditForm({ ...staffEditForm, full_name: e.target.value })}
                                          className="w-full px-3 py-2 border rounded-xl text-xs bg-white outline-none text-gray-900 font-medium"
                                        />
                                      </div>
                                      <div>
                                        <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Username *</label>
                                        <input
                                          type="text"
                                          required
                                          value={staffEditForm.username}
                                          onChange={(e) => setStaffEditForm({ ...staffEditForm, username: e.target.value.toLowerCase() })}
                                          className="w-full px-3 py-2 border rounded-xl text-xs bg-white outline-none text-gray-900 font-medium"
                                        />
                                      </div>
                                      <div>
                                        <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Password *</label>
                                        <div className="relative">
                                          <input
                                            type={showStaffEditPassword ? 'text' : 'password'}
                                            required
                                            value={staffEditForm.password}
                                            onChange={(e) => setStaffEditForm({ ...staffEditForm, password: e.target.value })}
                                            className="w-full px-3 py-2 pr-16 border rounded-xl text-xs bg-white outline-none text-gray-900 font-medium"
                                          />
                                          <button
                                            type="button"
                                            onClick={() => setShowStaffEditPassword(prev => !prev)}
                                            className="absolute right-3 top-1/2 -translate-y-1/2 text-[9px] font-black uppercase text-[#8a1538]"
                                          >
                                            {showStaffEditPassword ? 'Hide' : 'Show'}
                                          </button>
                                        </div>
                                      </div>
                                      <div>
                                        <label className="block text-[10px] font-bold uppercase text-gray-400 mb-1">Access Profile</label>
                                        <select
                                          value={staffEditForm.role}
                                          onChange={(e) => setStaffEditForm({ ...staffEditForm, role: e.target.value })}
                                          className="w-full px-3 py-2 border rounded-xl text-xs bg-white text-gray-800 outline-none font-bold uppercase"
                                        >
                                          <option value="Operator">Workflow Staff</option>
                                          <option value="Manager">Review Access</option>
                                          <option value="Photographer">Media Staff</option>
                                          <option value="Content Editor">Sheet Staff</option>
                                          {customRoles.map((cr, idx) => (
                                            <option key={idx} value={cr.roleName}>{cr.roleName}</option>
                                          ))}
                                        </select>
                                      </div>
                                    </div>

                                    <div className="flex flex-col sm:flex-row gap-2 pt-1">
                                      <button type="submit" className="flex-1 py-2.5 bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] text-white text-xs font-black uppercase tracking-wider rounded-xl cursor-pointer">
                                        Save Staff Changes
                                      </button>
                                      <button type="button" onClick={cancelStaffEditPanel} className="sm:w-32 py-2.5 bg-white border border-gray-200 text-gray-600 text-xs font-black uppercase tracking-wider rounded-xl cursor-pointer hover:bg-gray-50">
                                        Cancel
                                      </button>
                                    </div>
                                  </form>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </div>

                    {/* Permissions Matrix */}
                    <div className="bg-white border border-gray-200 rounded-2xl shadow-xs overflow-hidden">
                      <button
                        type="button"
                        onClick={() => setShowPermissions(prev => !prev)}
                        className="w-full p-6 flex items-center justify-between gap-4 text-left hover:bg-gray-50 transition-colors"
                      >
                        <div>
                          <h2 className="text-sm font-black text-gray-900 uppercase tracking-tight">Permissions</h2>
                          <p className="text-xs text-gray-400 font-medium mt-1">
                            Expanded access matrix for current portal features. Click to {showPermissions ? 'hide' : 'show'} details.
                          </p>
                        </div>
                        <span className="px-3 py-1.5 rounded-xl bg-[rgba(138,21,56,0.06)] text-[#8a1538] border border-[rgba(138,21,56,0.18)] text-xs font-black uppercase">
                          {showPermissions ? 'Hide ▲' : 'Show ▼'}
                        </span>
                      </button>

                      {showPermissions && (
                        <div className="border-t border-gray-100 p-6">
                          <div className="mb-4 p-4 bg-amber-50 border border-amber-200 rounded-2xl text-[11px] font-bold text-amber-800 leading-relaxed">
                            This matrix is a clear permission reference. Actual restrictions are still enforced inside the workflow buttons, status rules, and permission checks in the app.
                          </div>
                          <div className="overflow-x-auto border rounded-xl">
                            <table className="w-full text-left border-collapse min-w-[1500px]">
                              <thead>
                                <tr className="bg-gray-50 border-b text-xxs uppercase tracking-wider font-black text-gray-400">
                                  <th className="p-3 sticky left-0 bg-gray-50 z-10 min-w-[190px]">Access Group</th>
                                  {permissionFeatures.map(feature => (
                                    <th key={feature.key} className="p-3 text-center min-w-[105px]">{feature.label}</th>
                                  ))}
                                  <th className="p-3 text-center min-w-[120px]">Actions</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y text-xs font-semibold text-gray-700">
                                {permissionRoles.map((role) => {
                                  const isCustomRole = !['Admin', 'Manager', 'Operator', 'Photographer', 'Content Editor'].includes(role.roleName);
                                  return (
                                    <tr key={role.roleName} className="hover:bg-gray-50/40 transition-colors">
                                      <td className="p-3 font-black text-gray-900 sticky left-0 bg-white z-10 border-r border-gray-100">
                                        {role.label || role.roleName}
                                      </td>
                                      {permissionFeatures.map(feature => {
                                        const allowed = hasMatrixPermission(role.roleName, feature.key);
                                        return (
                                          <td key={`${role.roleName}-${feature.key}`} className="p-3 text-center">
                                            <span className={`inline-flex items-center justify-center w-8 h-8 rounded-full border text-[11px] font-black ${allowed ? 'bg-green-50 text-green-700 border-green-200' : 'bg-red-50 text-red-500 border-red-100'}`}>
                                              {allowed ? '✓' : '×'}
                                            </span>
                                          </td>
                                        );
                                      })}
                                      <td className="p-3 text-center">
                                        {isCustomRole ? (
                                          <button
                                            onClick={() => handleDeleteCustomRole(role.roleName)}
                                            className="px-2.5 py-1 text-red-600 border border-red-100 hover:bg-red-50 rounded uppercase tracking-wider text-xxs font-black transition-colors cursor-pointer"
                                          >
                                            Drop Group
                                          </button>
                                        ) : (
                                          <span className="tracking-widest uppercase text-xxs font-black text-gray-400">Core Protected</span>
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
                          <p className="text-sm text-slate-300 font-medium mt-1">Authorized users can post assignments. Staff can reply, ask questions, and send progress updates inside each task thread.</p>
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

                    {taskReplyError && (
                      <div className="p-4 bg-amber-50 border border-amber-200 text-amber-800 rounded-xl text-xs font-bold">
                        Task replies warning: {taskReplyError}. Check the task_replies table permissions if replies do not load or send.
                      </div>
                    )}

                    {isTaskManager && (
                      <div className="bg-white border border-gray-200 rounded-2xl p-6 shadow-xs">
                        <h2 className="text-sm font-black text-gray-900 uppercase tracking-wider mb-1">
                          {editingTaskId ? 'Edit Task' : 'Post New Task'}
                        </h2>
                        <p className="text-xs text-gray-400 font-semibold mb-5">Tasks posted here are visible immediately, and every visible employee can reply in the task chat thread.</p>
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
                              <option value="Operator">Workflow Staff</option>
                              <option value="Photographer">Media Staff</option>
                              <option value="Content Editor">Sheet Staff</option>
                              <option value="Manager">Review Leads</option>
                              {userRegistry.map(user => (
                                <option key={user.id} value={user.username}>{user.full_name ? `${user.full_name} (@${user.username})` : user.username}</option>
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

                    <div className="bg-white border border-gray-200 rounded-2xl p-5 shadow-xs">
                      <div className="flex flex-col lg:flex-row lg:items-end lg:justify-between gap-4">
                        <div>
                          <h2 className="text-sm font-black text-gray-900 uppercase tracking-wider">Task Filters</h2>
                          <p className="text-xs text-gray-400 font-semibold mt-1">
                            {isTaskManager
                              ? 'Managers see every task. Use filters to view All Staff, role-based, or individual employee tasks.'
                              : 'Filter your visible task queue by status, priority, or search.'}
                          </p>
                        </div>
                        <button
                          type="button"
                          onClick={() => {
                            setTaskSearchQuery('');
                            setTaskStatusFilter('All');
                            setTaskPriorityFilter('All');
                            setTaskTargetFilter('All');
                          }}
                          className="px-3 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl text-[10px] font-black uppercase tracking-wider"
                        >
                          Clear Filters
                        </button>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-4 gap-3 mt-4">
                        <div className="md:col-span-1">
                          <label className="block text-[10px] font-black uppercase tracking-widest text-gray-400 mb-1">Search</label>
                          <input
                            type="text"
                            value={taskSearchQuery}
                            onChange={(e) => setTaskSearchQuery(e.target.value)}
                            placeholder="Search title, details, sender..."
                            className="w-full px-3 py-2 border border-gray-200 rounded-xl bg-gray-50 text-xs text-gray-900 outline-none"
                          />
                        </div>

                        <div>
                          <label className="block text-[10px] font-black uppercase tracking-widest text-gray-400 mb-1">Status</label>
                          <select
                            value={taskStatusFilter}
                            onChange={(e) => setTaskStatusFilter(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-200 rounded-xl bg-white text-xs text-gray-900 font-bold outline-none"
                          >
                            {taskStatusFilterOptions.map(statusOption => (
                              <option key={statusOption} value={statusOption}>{statusOption === 'All' ? 'All Statuses' : statusOption}</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-[10px] font-black uppercase tracking-widest text-gray-400 mb-1">Priority</label>
                          <select
                            value={taskPriorityFilter}
                            onChange={(e) => setTaskPriorityFilter(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-200 rounded-xl bg-white text-xs text-gray-900 font-bold outline-none"
                          >
                            {taskPriorityFilterOptions.map(priorityOption => (
                              <option key={priorityOption} value={priorityOption}>{priorityOption === 'All' ? 'All Priorities' : priorityOption}</option>
                            ))}
                          </select>
                        </div>

                        <div>
                          <label className="block text-[10px] font-black uppercase tracking-widest text-gray-400 mb-1">Visible To</label>
                          <select
                            value={taskTargetFilter}
                            onChange={(e) => setTaskTargetFilter(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-200 rounded-xl bg-white text-xs text-gray-900 font-bold outline-none"
                          >
                            {taskTargetFilterOptions.map(targetOption => (
                              <option key={targetOption} value={targetOption}>{targetOption === 'All' ? 'All Targets' : formatTaskTargetLabel(targetOption)}</option>
                            ))}
                          </select>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mt-4 text-center">
                        <div className="bg-gray-50 border rounded-xl p-3">
                          <span className="block text-[9px] font-black uppercase text-gray-400">Visible</span>
                          <b className="block text-lg font-black text-gray-900">{visibleTasks.length}</b>
                        </div>
                        <div className="bg-amber-50 border border-amber-100 rounded-xl p-3">
                          <span className="block text-[9px] font-black uppercase text-amber-700">Open</span>
                          <b className="block text-lg font-black text-amber-800">{visibleTasks.filter(task => (task.status || 'Open') === 'Open').length}</b>
                        </div>
                        <div className="bg-[rgba(138,21,56,0.06)] border border-[rgba(138,21,56,0.18)] rounded-xl p-3">
                          <span className="block text-[9px] font-black uppercase text-[#8a1538]">Progress</span>
                          <b className="block text-lg font-black text-[#8a1538]">{visibleTasks.filter(task => task.status === 'In Progress').length}</b>
                        </div>
                        <div className="bg-green-50 border border-green-100 rounded-xl p-3">
                          <span className="block text-[9px] font-black uppercase text-green-700">Done</span>
                          <b className="block text-lg font-black text-green-800">{visibleTasks.filter(task => task.status === 'Done').length}</b>
                        </div>
                        <div className="bg-indigo-50 border border-indigo-100 rounded-xl p-3">
                          <span className="block text-[9px] font-black uppercase text-indigo-700">Showing</span>
                          <b className="block text-lg font-black text-indigo-800">{filteredVisibleTasks.length}</b>
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                      {visibleTasks.length === 0 ? (
                        <div className="lg:col-span-2 text-center text-xs text-gray-400 font-bold uppercase py-16 bg-white border rounded-2xl">
                          No tasks available for you.
                        </div>
                      ) : filteredVisibleTasks.length === 0 ? (
                        <div className="lg:col-span-2 text-center text-xs text-gray-400 font-bold uppercase py-16 bg-white border rounded-2xl">
                          No tasks match the selected filters.
                        </div>
                      ) : (
                        filteredVisibleTasks.map(task => (
                          <div key={task.id} className="bg-white border border-gray-200 rounded-2xl p-5 shadow-xs hover:shadow-sm transition-all">
                            <div className="flex items-start justify-between gap-4 mb-3">
                              <div>
                                <h3 className="text-base font-black text-gray-900">{task.title}</h3>
                                <div className="text-[10px] font-bold text-gray-400 mt-1 uppercase tracking-wider">
                                  Posted by {getUserDisplayName(task.created_by) || task.created_by || 'System User'} • {formatDisplayDateTime(task.created_at)}
                                </div>
                              </div>
                              <div className="flex flex-col items-end gap-1">
                                <span className={`px-2.5 py-1 rounded-lg border text-[10px] font-black uppercase ${getPriorityClass(task.priority)}`}>{task.priority || 'Normal'}</span>
                                <span className={`px-2.5 py-1 rounded-lg border text-[10px] font-black uppercase ${getTaskStatusClass(task.status)}`}>{task.status || 'Open'}</span>
                                <span className="px-2.5 py-1 rounded-lg border text-[10px] font-black uppercase bg-indigo-50 text-indigo-700 border-indigo-100">💬 {getTaskReplyCount(task.id)}</span>
                              </div>
                            </div>
                            {task.description && <p className="text-sm text-gray-600 font-medium whitespace-pre-wrap leading-relaxed mb-4">{task.description}</p>}
                            <div className="grid grid-cols-2 gap-3 text-[11px] font-bold text-gray-500 mb-4">
                              <div className="bg-gray-50 border rounded-xl p-3"><span className="block text-gray-400 uppercase text-[9px]">Visible To</span>{formatTaskTargetLabel(task.assigned_role)}</div>
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
                            {renderTaskRepliesPanel(task)}
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
                        {(authRole === 'Admin' || authRole === 'Manager') && isStoreScoped && selectedStore && (
                          <button
                            type="button"
                            onClick={() => openStoreEditModal(selectedStore)}
                            className="bg-[rgba(138,21,56,0.06)] text-[#8a1538] border border-[rgba(138,21,56,0.18)] font-black text-xs px-4 py-2 rounded-xl hover:bg-[rgba(138,21,56,0.10)] transition uppercase"
                          >
                            Edit Store
                          </button>
                        )}
                        {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('ad_hoc')) && (
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
                                  Download Ready to Work
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
                        <div className="text-xs font-bold text-gray-400 uppercase tracking-wider">Total Products</div>
                        <div className="text-2xl font-black text-gray-900 mt-1">{metrics.total}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-amber-500">
                        <div className="text-xs font-bold text-amber-600 uppercase tracking-wider">Ready to work</div>
                        <div className="text-2xl font-black text-amber-700 mt-1">{metrics.missing}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-blue-500">
                        <div className="text-xs font-bold text-[rgba(138,21,56,0.85)] uppercase tracking-wider">In Progress</div>
                        <div className="text-2xl font-black text-[#8a1538] mt-1">{metrics.processing}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-sky-500">
                        <div className="text-xs font-bold text-blue-600 uppercase tracking-wider">Under Review</div>
                        <div className="text-2xl font-black text-blue-700 mt-1">{metrics.underReview}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-red-500">
                        <div className="text-xs font-bold text-red-600 uppercase tracking-wider">Rejected</div>
                        <div className="text-2xl font-black text-red-700 mt-1">{metrics.rejected}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-emerald-500">
                        <div className="text-xs font-bold text-emerald-600 uppercase tracking-wider">Ready to Upload</div>
                        <div className="text-2xl font-black text-emerald-700 mt-1">{metrics.readyToUpload}</div>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-xs border-l-4 border-l-purple-500">
                        <div className="text-xs font-bold text-[rgba(138,21,56,0.85)] uppercase tracking-wider">Modified</div>
                        <div className="text-2xl font-black text-[#8a1538] mt-1">{metrics.modified}</div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
                      {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('excel_upload')) && (
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

                      {(authRole === 'Admin' || authRole === 'Manager' || authRole === 'Photographer' || checkPermission('bulk_images')) && (
                        <div className="bg-white p-6 rounded-xl shadow-sm border border-[rgba(138,21,56,0.28)] flex flex-col justify-between bg-gradient-to-br from-blue-50/40 via-white to-white">
                          <div>
                            <h3 className="text-sm font-black text-[rgba(138,21,56,0.85)] uppercase tracking-wide flex items-center gap-2">
                              <span>🗂️</span> Bulk Image Asset Upload
                            </h3>
                            <p className="text-xs text-[#8a1538] font-medium mt-1">Supported ZIP/folder formats: <strong>[SKU or Barcode]/RAW/img.jpg</strong>, <strong>[SKU or Barcode]/EDITED/img.jpg</strong>, or ZIP files named by SKU/Barcode containing <strong>RAW/img.jpg</strong>.</p>
                            <p className={`text-[11px] font-bold mt-2 ${isStoreScoped ? 'text-emerald-700' : 'text-red-600'}`}>
                              Upload target: {isStoreScoped ? getStoreNameById(selectedStoreId) : 'Select one store first'}
                            </p>
                          </div>
                          <div className="mt-4 flex flex-col gap-2">
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                              <label className={`w-full py-3 rounded-xl text-xs font-bold uppercase tracking-wider shadow-xs transition-colors block text-center ${isStoreScoped && !uploading ? 'bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] text-white cursor-pointer' : 'bg-gray-200 text-gray-500 cursor-not-allowed'}`}>
                                {uploading ? "Extracting assets..." : isStoreScoped ? "Upload ZIP File(s)" : "Select Store First"}
                                <input
                                  type="file"
                                  accept=".zip"
                                  multiple
                                  onChange={handleBulkZipUpload}
                                  disabled={uploading || !isStoreScoped}
                                  className="hidden"
                                />
                              </label>

                              <label className={`w-full py-3 rounded-xl text-xs font-bold uppercase tracking-wider shadow-xs transition-colors block text-center ${isStoreScoped && !uploading ? 'bg-white hover:bg-gray-50 text-[rgba(138,21,56,0.95)] border border-[rgba(138,21,56,0.28)] cursor-pointer' : 'bg-gray-200 text-gray-500 cursor-not-allowed'}`}>
                                {uploading ? "Reading folders..." : isStoreScoped ? "Upload Folder Tree" : "Select Store First"}
                                <input
                                  type="file"
                                  multiple
                                  webkitdirectory=""
                                  directory=""
                                  onChange={handleBulkFolderUpload}
                                  disabled={uploading || !isStoreScoped}
                                  className="hidden"
                                />
                              </label>
                            </div>
                            <p className="text-[10px] text-gray-500 font-bold leading-relaxed">
                              For full inventory upload, select one parent folder that contains many <strong>SKU/RAW</strong> and <strong>SKU/EDITED</strong> folders, or select multiple ZIP files at once.
                            </p>

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
                        {WORKFLOW_STATUS_FILTERS.map((statusOption) => (
                          <button
                            key={statusOption.value} onClick={() => setStatusFilter(statusOption.value)}
                            className={`px-3 py-1.5 rounded-lg text-xs font-bold uppercase transition-all border ${
                              statusFilter === statusOption.value ? 'bg-[rgba(138,21,56,0.85)] text-white border-[rgba(138,21,56,0.85)]' : 'bg-gray-50 text-gray-600 border-gray-200 hover:bg-gray-100'
                            }`}
                          >
                            {statusOption.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    {authRole === 'Operator' ? (
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-6">
                        {filteredProducts.map((prod) => {
                          const isProcessing = prod.status === 'Processing';
                          const isRejected = prod.status === 'Rejected';
                          const statusColors = getStatusBadgeClass(prod.status);
                          return (
                            <div 
                              key={prod.id} onClick={() => handleCardInteraction(prod)}
                              className={`bg-white rounded-2xl border shadow-sm p-6 flex flex-col h-full transition-all hover:shadow-md hover:-translate-y-0.5 cursor-pointer group ${isRejected ? 'border-red-300 bg-red-50/20' : isProcessing ? 'border-[rgba(138,21,56,0.28)] bg-[rgba(138,21,56,0.06)]/10' : 'border-gray-200'}`}
                            >
                              <div className="flex-grow">
                                <div className="flex items-center justify-between mb-4">
                                  <span className={`text-[10px] font-bold uppercase px-2.5 py-1 rounded-md border ${statusColors}`}>{getStatusLabel(prod.status)}</span>
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
                                      Review Note
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
                                      <span className={`px-3 py-1 rounded-full border text-xs font-bold ${getStatusBadgeClass(prod.status)}`}>{getStatusLabel(prod.status)}</span>
                                    </td>
                                    <td className="p-4 text-center">
                                      {isEditing ? (
                                        <div className="space-x-1 flex justify-center">
                                          <button onClick={() => handleSaveEdit(prod.id)} className="px-3 py-1 bg-green-600 text-white rounded text-xs cursor-pointer">Save</button>
                                          <button onClick={() => setEditingId(null)} className="px-3 py-1 bg-gray-200 text-gray-600 rounded text-xs cursor-pointer">Exit</button>
                                        </div>
                                      ) : (
                                        <div className="space-x-1 flex justify-center">
                                          {(authRole === 'Admin' || authRole === 'Manager' || checkPermission('product_edit')) && (
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
                {activeTab === 'operators' && (authRole === 'Admin' || authRole === 'Manager' || checkPermission('performance')) && (
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

                    <div className="grid grid-cols-2 md:grid-cols-7 gap-4">
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Total Claimed</span>
                        <div className="text-xl font-black text-slate-800 mt-1">{selfPerformanceStats.totalClaimed}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Ready to work</span>
                        <div className="text-xl font-black text-amber-600 mt-1">{selfPerformanceStats.totalMissing}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">In Progress</span>
                        <div className="text-xl font-black text-[rgba(138,21,56,0.85)] mt-1">{selfPerformanceStats.totalProcessing}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Under Review</span>
                        <div className="text-xl font-black text-blue-600 mt-1">{selfPerformanceStats.totalCompleted}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Ready to Upload</span>
                        <div className="text-xl font-black text-emerald-600 mt-1">{selfPerformanceStats.totalReadyToUpload}</div>
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
                        <span className="text-[9px] font-bold text-green-700 uppercase block">Under Review This Week</span>
                        <div className="text-xl font-black text-green-700 mt-1">{selfPerformanceStats.weekCount}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-green-50 border-green-100 shadow-xs">
                        <span className="text-[9px] font-bold text-green-700 uppercase block">Under Review This Month</span>
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
                        <span className="text-[9px] font-bold text-gray-500 uppercase block">Average Time To Under Review</span>
                        <div className="text-lg font-black text-slate-900 mt-1">{selfPerformanceStats.avgCompletedTime}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-red-50 border-red-100 shadow-xs">
                        <span className="text-[9px] font-bold text-red-700 uppercase block">Average Time Before Rejection</span>
                        <div className="text-lg font-black text-red-800 mt-1">{selfPerformanceStats.avgRejectedTime}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Fastest Under Review Product</span>
                        <div className="text-lg font-black text-emerald-700 mt-1">{selfPerformanceStats.fastestCompletedTime}</div>
                      </div>
                      <div className="border p-4 rounded-xl bg-white shadow-xs">
                        <span className="text-[9px] font-bold text-gray-400 uppercase block">Slowest Under Review Product</span>
                        <div className="text-lg font-black text-orange-700 mt-1">{selfPerformanceStats.slowestCompletedTime}</div>
                      </div>
                    </div>

                    <div className="bg-white rounded-xl p-6 border border-gray-200 shadow-xs">
                      <div className="flex flex-col sm:flex-row sm:items-end sm:justify-between gap-3 mb-5">
                        <div>
                          <h2 className="text-xl font-black text-gray-900 uppercase tracking-tight">Done Product Time Log</h2>
                          <p className="text-xs text-gray-400 font-semibold mt-1">Under review, ready to upload, modified, and rejected products assigned to you, with the tracked time spent on each item.</p>
                        </div>
                        <span className="text-[10px] font-black uppercase tracking-wider bg-slate-100 text-slate-600 px-3 py-1 rounded-full border">
                          {selfPerformanceStats.doneProductLedger.length} done item(s)
                        </span>
                      </div>

                      {selfPerformanceStats.doneProductLedger.length === 0 ? (
                        <div className="text-center text-xs text-gray-400 font-bold uppercase py-12 border rounded-xl bg-gray-50">
                          No under review, ready to upload, modified, or rejected products found yet.
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
                                <th className="p-4">Review Note</th>
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
                                    <span className={`px-3 py-1 rounded-full border text-[10px] font-black uppercase ${getStatusBadgeClass(item.status)}`}>
                                      {getStatusLabel(item.status)}
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
                        Note: time tracking starts when you change a product to In Progress and stops when it becomes Under Review or Rejected. Older products may show “No tracked time”.
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
                              <button onClick={() => handleDownloadMissingByName(sheetName)} className="px-3 py-1.5 bg-amber-500 text-white font-bold text-xs uppercase rounded-lg cursor-pointer hover:bg-amber-600">Download Ready to Work</button>
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

                      <div className="flex flex-col sm:items-end gap-2">
                        <div className="flex flex-wrap items-center gap-2 justify-start sm:justify-end">
                          <button
                            onClick={() => handleBulkDownloadAssetDirectory(assetDirectoryProducts, 'all')}
                            disabled={uploading || assetDirectoryProducts.length === 0 || (assetDirectoryRawCount + assetDirectoryEditedCount) === 0}
                            className="px-3 py-2 bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                          >
                            ⬇️ All Visible
                          </button>
                          <button
                            onClick={() => handleBulkDownloadAssetDirectory(assetDirectoryProducts, 'raw')}
                            disabled={uploading || assetDirectoryRawCount === 0}
                            className="px-3 py-2 bg-[rgba(138,21,56,0.95)] hover:bg-[rgba(138,21,56,0.85)] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                          >
                            Visible RAW
                          </button>
                          <button
                            onClick={() => handleBulkDownloadAssetDirectory(assetDirectoryProducts, 'edited')}
                            disabled={uploading || assetDirectoryEditedCount === 0}
                            className="px-3 py-2 bg-[rgba(138,21,56,0.85)] hover:bg-[#8a1538] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                          >
                            Visible EDITED
                          </button>
                        </div>

                        <div className="flex flex-wrap items-center gap-2 justify-start sm:justify-end">
                          <button
                            onClick={() => handleBulkDownloadAssetDirectory(selectedAssetDirectoryProducts, 'all')}
                            disabled={uploading || selectedAssetDirectoryProducts.length === 0 || (selectedAssetRawCount + selectedAssetEditedCount) === 0}
                            className="px-3 py-2 bg-gray-900 hover:bg-black disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                          >
                            Selected All ({selectedAssetDirectoryProducts.length})
                          </button>
                          <button
                            onClick={() => handleBulkDownloadAssetDirectory(selectedAssetDirectoryProducts, 'raw')}
                            disabled={uploading || selectedAssetRawCount === 0}
                            className="px-3 py-2 bg-slate-700 hover:bg-slate-800 disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                          >
                            Selected RAW
                          </button>
                          <button
                            onClick={() => handleBulkDownloadAssetDirectory(selectedAssetDirectoryProducts, 'edited')}
                            disabled={uploading || selectedAssetEditedCount === 0}
                            className="px-3 py-2 bg-[#8a1538] hover:bg-[rgba(138,21,56,0.95)] disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-xl text-[10px] font-black uppercase tracking-wider"
                          >
                            Selected EDITED
                          </button>
                        </div>
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

                    <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3 bg-white border border-gray-200 rounded-2xl p-4">
                      <div>
                        <div className="text-xs font-black text-gray-900 uppercase tracking-wider">Selected Asset Folders: {selectedAssetDirectoryProducts.length}</div>
                        <p className="text-[11px] text-gray-400 font-bold mt-1">Selected images available: RAW {selectedAssetRawCount} / EDITED {selectedAssetEditedCount}. Select products below, then download only those folders.</p>
                      </div>
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={handleSelectAllVisibleAssetProducts}
                          disabled={assetDirectoryProducts.length === 0}
                          className="px-3 py-2 bg-gray-100 hover:bg-gray-200 disabled:opacity-50 text-gray-700 rounded-xl text-[10px] font-black uppercase tracking-wider"
                        >
                          {allVisibleAssetProductsSelected ? 'Unselect Visible' : 'Select Visible'}
                        </button>
                        <button
                          type="button"
                          onClick={clearAssetProductSelection}
                          disabled={selectedAssetProductIds.length === 0}
                          className="px-3 py-2 bg-red-50 hover:bg-red-100 disabled:opacity-50 text-red-700 border border-red-100 rounded-xl text-[10px] font-black uppercase tracking-wider"
                        >
                          Clear Selection
                        </button>
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
                            <div key={prod.id} className={`p-5 border rounded-2xl bg-white shadow-xs hover:border-[rgba(138,21,56,0.28)] transition-colors ${selectedAssetProductIds.includes(prod.id) ? 'ring-2 ring-[rgba(138,21,56,0.35)] border-[rgba(138,21,56,0.45)]' : ''}`}>
                              <div className="flex items-start justify-between gap-3 mb-4">
                                <div className="flex items-start gap-3 min-w-0">
                                  <label className="mt-1 w-5 h-5 flex items-center justify-center cursor-pointer shrink-0" title="Select this product for bulk download">
                                    <input
                                      type="checkbox"
                                      checked={selectedAssetProductIds.includes(prod.id)}
                                      onChange={() => toggleAssetProductSelection(prod.id)}
                                      className="w-4 h-4 accent-[#8a1538] cursor-pointer"
                                    />
                                  </label>
                                  <div className="min-w-0">
                                    <span className="text-xs font-mono font-black text-[#8a1538] bg-[rgba(138,21,56,0.06)] px-2.5 py-1 rounded-md block w-fit">📁 SKU: {prod.sku || 'UNKNOWN'}</span>
                                    <div className="text-[11px] font-bold text-gray-400 mt-2 line-clamp-2">{prod.product_name || 'Unnamed Product'}</div>
                                    <div className="text-[10px] font-black uppercase tracking-wider text-emerald-700 mt-1">{getStoreNameById(prod.store_id)}</div>
                                  </div>
                                </div>
                                <span className="text-[10px] font-black px-2 py-1 rounded-lg bg-gray-100 text-gray-600 uppercase shrink-0">{getStatusLabel(prod.status)}</span>
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

      {/* STORE EDIT MODAL */}
      {editingStore && (authRole === 'Admin' || authRole === 'Manager') && (
        <div className="fixed inset-0 bg-[rgba(138,21,56,1)]/60 backdrop-blur-md flex items-center justify-center p-4 z-[74] animate-fadeIn">
          <div className="bg-white border rounded-3xl w-full max-w-2xl shadow-2xl overflow-hidden">
            <div className="flex items-start justify-between gap-4 border-b p-6 bg-gray-50">
              <div>
                <h3 className="font-black text-lg text-gray-900 uppercase">Edit Store</h3>
                <p className="text-xs font-semibold text-gray-400 mt-1">Change the store name, add or replace its image, or remove the current image.</p>
              </div>
              <button onClick={closeStoreEditModal} className="text-gray-400 text-xl cursor-pointer hover:text-gray-900">✕</button>
            </div>

            <div className="p-6 space-y-6">
              <div className="relative h-56 rounded-2xl overflow-hidden border bg-gradient-to-br from-emerald-900 via-[rgba(138,21,56,0.95)] to-[rgba(138,21,56,1)]">
                {editingStore.image_url ? (
                  <img src={editingStore.image_url} alt={editingStore.name || 'Store image'} className="w-full h-full object-cover" />
                ) : (
                  <div className="w-full h-full flex flex-col items-center justify-center text-white/85">
                    <div className="text-6xl mb-2">🏪</div>
                    <div className="text-xs font-black uppercase tracking-widest">No store image added</div>
                  </div>
                )}
                <div className="absolute inset-0 bg-gradient-to-t from-black/55 via-transparent to-transparent pointer-events-none"></div>
                <div className="absolute left-5 bottom-4 right-5">
                  <div className="text-2xl font-black text-white truncate">{editingStore.name || 'Unnamed Store'}</div>
                  <div className="text-[10px] font-black text-white/80 uppercase tracking-wider">Store ID: {editingStore.id}</div>
                </div>
              </div>

              <form onSubmit={handleUpdateStoreName} className="space-y-3">
                <div>
                  <label className="block text-[10px] font-black uppercase text-gray-400 tracking-widest mb-1.5">Store Name</label>
                  <input
                    type="text"
                    required
                    value={storeEditForm.name}
                    onChange={(e) => setStoreEditForm({ ...storeEditForm, name: e.target.value })}
                    placeholder="Store name..."
                    className="w-full px-4 py-3 text-sm border rounded-xl bg-gray-50 text-gray-900 outline-none focus:bg-white"
                  />
                </div>
                <button
                  type="submit"
                  disabled={storeEditSaving}
                  className="w-full py-3 bg-[#8a1538] hover:bg-[rgba(138,21,56,0.95)] text-white text-xs uppercase font-black rounded-xl cursor-pointer disabled:opacity-60 disabled:cursor-not-allowed"
                >
                  {storeEditSaving ? 'Saving Store Name...' : 'Save Store Name'}
                </button>
              </form>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 border-t pt-5">
                <label className={`w-full py-3 rounded-xl text-center text-xs font-black uppercase tracking-wider border transition-all ${storeImageUploadingId === editingStore.id ? 'bg-gray-100 text-gray-400 border-gray-200 cursor-not-allowed' : 'bg-[rgba(138,21,56,0.06)] text-[#8a1538] border-[rgba(138,21,56,0.18)] hover:bg-[rgba(138,21,56,0.10)] cursor-pointer'}`}>
                  {storeImageUploadingId === editingStore.id ? 'Uploading Image...' : editingStore.image_url ? 'Replace Store Image' : 'Add Store Image'}
                  <input
                    type="file"
                    accept="image/*"
                    disabled={storeImageUploadingId === editingStore.id}
                    onChange={(e) => handleStoreImageUpload(editingStore.id, e)}
                    className="hidden"
                  />
                </label>

                <button
                  type="button"
                  onClick={() => handleRemoveStoreImage(editingStore.id)}
                  disabled={!editingStore.image_url || storeImageRemovingId === editingStore.id || storeImageUploadingId === editingStore.id}
                  className="w-full py-3 rounded-xl text-center text-xs font-black uppercase tracking-wider border bg-red-50 text-red-600 border-red-100 hover:bg-red-100 disabled:bg-gray-100 disabled:text-gray-400 disabled:border-gray-200 disabled:cursor-not-allowed"
                >
                  {storeImageRemovingId === editingStore.id ? 'Removing Image...' : 'Remove Store Image'}
                </button>
              </div>

              <div className="p-4 rounded-2xl bg-amber-50 border border-amber-200 text-[11px] font-bold text-amber-800 leading-relaxed">
                Replacing an image updates the store card immediately and attempts to clean up the old storage file. Removing an image clears the store card image but keeps the store and all linked products.
              </div>

              <div className="flex flex-col sm:flex-row gap-3">
                <button
                  type="button"
                  onClick={closeStoreEditModal}
                  className="flex-1 py-3 border rounded-xl text-xs font-black uppercase text-gray-600 hover:bg-gray-50"
                >
                  Close
                </button>
                <button
                  type="button"
                  onClick={() => handleDeleteStore(editingStore.id, editingStore.name)}
                  disabled={uploading}
                  className="flex-1 py-3 bg-red-50 hover:bg-red-100 border border-red-100 text-red-700 rounded-xl text-xs font-black uppercase disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  Delete Store
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

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

      {/* GOOGLE SHEETS LINK MODAL */}
      {showGoogleSheetModal && (authRole === 'Admin' || authRole === 'Manager') && (
        <div className="fixed inset-0 bg-[rgba(138,21,56,1)]/60 backdrop-blur-md flex items-center justify-center p-4 z-[75] animate-fadeIn">
          <div className="bg-white border rounded-3xl w-full max-w-lg shadow-2xl p-8 overflow-hidden">
            <div className="flex items-center justify-between border-b pb-4 mb-6">
              <div>
                <h3 className="font-black text-lg text-gray-900 uppercase">Google Sheets Link</h3>
                <p className="text-xs font-semibold text-gray-400 mt-1">Only authorized users can save this. Staff can open it from Home.</p>
              </div>
              <button onClick={() => setShowGoogleSheetModal(false)} className="text-gray-400 text-xl cursor-pointer hover:text-gray-900">✕</button>
            </div>
            <form onSubmit={handleSaveGoogleSheetLink} className="space-y-4">
              <div>
                <label className="block text-xxs font-black uppercase text-gray-400 tracking-widest mb-1.5">Google Sheet URL</label>
                <input
                  type="url"
                  required
                  value={googleSheetDraft}
                  onChange={(e) => setGoogleSheetDraft(e.target.value)}
                  placeholder="https://docs.google.com/spreadsheets/d/..."
                  className="w-full px-4 py-3 text-sm border rounded-xl bg-gray-50 text-gray-900 outline-none"
                />
              </div>
              {googleSheetLink && (
                <button
                  type="button"
                  onClick={openGoogleSheetLink}
                  className="w-full py-2.5 bg-gray-50 hover:bg-gray-100 border border-gray-200 text-gray-700 text-xs font-black uppercase rounded-xl"
                >
                  Open Current Google Sheet
                </button>
              )}
              <div className="flex gap-3 pt-3">
                <button type="button" onClick={() => setShowGoogleSheetModal(false)} className="w-1/3 py-3 border rounded-xl text-xs font-bold uppercase text-gray-600 hover:bg-gray-50">Cancel</button>
                <button type="submit" disabled={googleSheetSaving} className="w-2/3 py-3 bg-[#8a1538] hover:bg-[rgba(138,21,56,0.95)] text-white text-xs uppercase font-black rounded-xl cursor-pointer disabled:opacity-60">
                  {googleSheetSaving ? 'Saving...' : 'Save Google Sheet Link'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* CUSTOM DEPARTMENT MODAL */}
      {showRoleModal && (
        <div className="fixed inset-0 bg-[rgba(138,21,56,1)]/60 backdrop-blur-md flex items-center justify-center p-4 z-50 animate-fadeIn">
          <div className="bg-white border rounded-3xl w-full max-w-5xl shadow-2xl p-6 overflow-hidden max-h-[92vh] flex flex-col">
            <div className="flex items-start justify-between gap-4 border-b pb-4 mb-5">
              <div>
                <h3 className="font-black text-base text-gray-900 uppercase">Deploy Custom Permissions</h3>
                <p className="text-xs font-semibold text-gray-400 mt-1">
                  Choose exactly which portal features this custom access group can use.
                </p>
              </div>
              <button onClick={() => setShowRoleModal(false)} className="text-gray-400 text-lg cursor-pointer hover:text-gray-900">✕</button>
            </div>
            <form onSubmit={handleCreateCustomRole} className="space-y-5 overflow-y-auto pr-1">
              <div>
                <label className="block text-xxs font-black uppercase text-gray-400 tracking-widest mb-1.5">Custom Group Name</label>
                <input
                  type="text"
                  required
                  value={newRoleForm.roleName}
                  onChange={(e) => setNewRoleForm({ ...newRoleForm, roleName: e.target.value })}
                  placeholder="e.g. Lead Editor, Store Auditor, Task Coordinator..."
                  className="w-full px-4 py-2.5 text-sm border rounded-xl bg-gray-50 text-gray-900 outline-none focus:bg-white focus:border-[rgba(138,21,56,0.45)]"
                />
              </div>

              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                  <div>
                    <h4 className="text-xs font-black uppercase tracking-wider text-gray-900">Feature Permissions</h4>
                    <p className="text-[10px] font-bold text-gray-400 mt-1">
                      Selected: {selectedNewRolePermissionCount} / {permissionFeatures.length}
                    </p>
                  </div>
                  <div className="flex gap-2">
                    <button
                      type="button"
                      onClick={() => setAllNewRolePermissions(true)}
                      className="px-3 py-1.5 bg-white hover:bg-gray-100 border border-gray-200 rounded-lg text-[10px] font-black uppercase text-gray-700"
                    >
                      Select All
                    </button>
                    <button
                      type="button"
                      onClick={() => setAllNewRolePermissions(false)}
                      className="px-3 py-1.5 bg-red-50 hover:bg-red-100 border border-red-100 rounded-lg text-[10px] font-black uppercase text-red-600"
                    >
                      Clear All
                    </button>
                  </div>
                </div>

                <div className="space-y-4 max-h-[52vh] overflow-y-auto pr-2">
                  {Object.entries(permissionFeaturesByCategory).map(([category, features]) => (
                    <div key={category} className="bg-white border border-gray-200 rounded-2xl p-4">
                      <div className="text-[10px] font-black uppercase tracking-widest text-[#8a1538] mb-3">
                        {category}
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
                        {features.map(feature => {
                          const checked = Boolean(newRoleForm.permissions?.[feature.key]);
                          return (
                            <label
                              key={feature.key}
                              className={`flex items-center justify-between gap-3 rounded-xl border px-3 py-2 cursor-pointer transition-all ${checked ? 'bg-[rgba(138,21,56,0.06)] border-[rgba(138,21,56,0.28)] text-[#8a1538]' : 'bg-gray-50 border-gray-100 text-gray-700 hover:bg-gray-100'}`}
                            >
                              <span className="text-[11px] font-black uppercase tracking-wide leading-tight">{feature.label}</span>
                              <input
                                type="checkbox"
                                checked={checked}
                                onChange={(e) => setNewRolePermissionValue(feature.key, e.target.checked)}
                                className="w-4 h-4 accent-[#8a1538] shrink-0"
                              />
                            </label>
                          );
                        })}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex gap-3 pt-2 border-t border-gray-100 sticky bottom-0 bg-white pb-1">
                <button type="button" onClick={() => setShowRoleModal(false)} className="w-1/3 py-2.5 border rounded-xl text-xs font-bold text-gray-600 hover:bg-gray-50">Cancel</button>
                <button type="submit" className="w-2/3 py-2.5 bg-[#8a1538] text-white text-xs uppercase font-black rounded-xl cursor-pointer hover:bg-[rgba(138,21,56,0.95)]">Deploy Group</button>
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
                <span className="text-[10px] font-black tracking-wider text-gray-400 uppercase block">Staff Profile</span>
                <div className="text-base font-black text-slate-800 mt-1">👤 {selectedOperatorStats.username}</div>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-7 gap-4">
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Total Claimed</span>
                  <div className="text-xl font-black text-slate-800 mt-1">{selectedOperatorStats.performance.totalClaimed}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Ready to work</span>
                  <div className="text-xl font-black text-amber-600 mt-1">{selectedOperatorStats.performance.totalMissing}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">In Progress</span>
                  <div className="text-xl font-black text-[rgba(138,21,56,0.85)] mt-1">{selectedOperatorStats.performance.totalProcessing}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Under Review</span>
                  <div className="text-xl font-black text-blue-600 mt-1">{selectedOperatorStats.performance.totalCompleted}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white shadow-xxs">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Ready to Upload</span>
                  <div className="text-xl font-black text-emerald-600 mt-1">{selectedOperatorStats.performance.totalReadyToUpload}</div>
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
                  <span className="text-[9px] font-bold text-green-700 uppercase block">Under Review This Week</span>
                  <div className="text-xl font-black text-green-700 mt-1">{selectedOperatorStats.performance.weekCount}</div>
                </div>
                <div className="border p-4 rounded-xl bg-green-50 border-green-100">
                  <span className="text-[9px] font-bold text-green-700 uppercase block">Under Review This Month</span>
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
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Fastest Under Review Product</span>
                  <div className="text-lg font-black text-emerald-700 mt-1">{selectedOperatorStats.performance.fastestCompletedTime}</div>
                </div>
                <div className="border p-4 rounded-xl bg-white">
                  <span className="text-[9px] font-bold text-gray-400 uppercase block">Slowest Under Review Product</span>
                  <div className="text-lg font-black text-orange-700 mt-1">{selectedOperatorStats.performance.slowestCompletedTime}</div>
                </div>
              </div>

              <div className="p-3 rounded-xl bg-amber-50 border border-amber-200 text-amber-800 text-[11px] font-bold leading-relaxed">
                Note: time tracking starts when the employee changes status to In Progress. Older products completed before this update may show “No tracked time” because their timer was not recorded.
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
        const isAdminOrManager = authRole === 'Admin' || authRole === 'Manager';
        const isReadyToUploadLockedForEmployee = selectedProduct.status === 'Ready to Upload' && !isAdminOrManager;
        const canUploadProductImages =
          isAdminOrManager ||
          (!isReadyToUploadLockedForEmployee && selectedProduct.status === 'Processing' && selectedProduct.processed_by === loginUser);
        const canRemoveEditedImages =
          isAdminOrManager ||
          (!isReadyToUploadLockedForEmployee && selectedProduct.processed_by === loginUser && ['Processing', 'Rejected'].includes(selectedProduct.status));
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
                      value={selectedProduct.status === 'Completed' ? 'Under Review' : (selectedProduct.status || 'Missing')} 
                      onChange={(e) => handleOperatorStatusChange(selectedProduct.id, e.target.value)}
                      disabled={isReadyToUploadLockedForEmployee}
                      className={`w-full text-xs font-bold uppercase rounded-lg p-2 border shadow-sm focus:outline-none focus:ring-2 focus:ring-[#8a1538] ${
                        isReadyToUploadLockedForEmployee
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-700 cursor-not-allowed'
                          : selectedProduct.status === 'Rejected'
                            ? 'bg-red-50 border-red-200 text-red-700 cursor-pointer'
                            : 'bg-white border-gray-300 text-gray-900 cursor-pointer'
                      }`}
                    >
                      <option value="Missing" disabled={selectedProduct.status === 'Rejected'}>Ready to work</option>
                      <option value="Processing">In Progress</option>
                      <option value="Under Review" disabled={selectedProduct.status === 'Rejected'}>Under Review</option>
                      <option value="Rejected" disabled>Rejected</option>
                      <option value="Ready to Upload" disabled={!isAdminOrManager}>Ready to Upload</option>
                      <option value="Modified" disabled>Modified</option>
                    </select>
                  </div>
                </div>
                {isReadyToUploadLockedForEmployee && (
                  <div className="p-4 rounded-xl border border-emerald-200 bg-emerald-50 text-emerald-800 text-xs font-bold">
                    ✅ This product is marked <span className="underline">Ready to Upload</span>. It is now locked for employees. Only authorized users can change its status or remove images.
                  </div>
                )}
                {selectedProduct.status === 'Rejected' && (
                  <div className="p-4 rounded-xl border border-red-200 bg-red-50">
                    <div className="text-xs font-black text-red-700 uppercase tracking-wider mb-2">
                      Rejected after review
                    </div>
                    <p className="text-sm font-semibold text-red-900 whitespace-pre-wrap">
                      {selectedProduct.rejection_note || 'No rejection note was provided.'}
                    </p>
                    <p className="text-[11px] text-red-600 font-bold mt-3">
                      Change the status to <span className="underline">In Progress</span> to reopen this item, upload corrected images, and then submit it for Under Review again.
                    </p>
                  </div>
                )}

                {!canUploadProductImages && !isReadyToUploadLockedForEmployee && (
                  <div className={`p-4 rounded-xl border text-xs font-bold ${selectedProduct.status === 'Rejected' ? 'border-red-200 bg-red-50 text-red-800' : 'border-amber-200 bg-amber-50 text-amber-800'}`}>
                    🔒 Image upload is locked. Change Operational Status to <span className="underline">In Progress</span> first. That claims/reopens this product under your employee name and unlocks RAW/EDITED uploads.
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
                          🔒 Set In Progress First
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
                      {editAssets.map((url, i) => (
                        <div key={url} className="relative group rounded border overflow-hidden bg-gray-50">
                          <img src={url} className="w-full h-20 object-cover" alt={`Edited ${i + 1}`} />
                          <div className="absolute left-1 top-1 bg-black/60 text-white text-[8px] font-black px-1.5 py-0.5 rounded uppercase">
                            #{i + 1}
                          </div>
                          {canRemoveEditedImages && (
                            <button
                              type="button"
                              onClick={(event) => {
                                event.stopPropagation();
                                handleRemoveIndividualImage(selectedProduct.id, 'edited', url);
                              }}
                              disabled={imageUploadingProductId === selectedProduct.id}
                              className="absolute right-1 top-1 bg-red-600 hover:bg-red-700 disabled:bg-gray-400 text-white text-[8px] font-black px-1.5 py-0.5 rounded uppercase shadow"
                              title="Remove this edited image"
                            >
                              Remove
                            </button>
                          )}
                        </div>
                      ))}
                      
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
                          🔒 Set In Progress First
                        </div>
                      )}
                    </div>
                    <p className="text-[10px] text-gray-400 font-semibold mt-2">
                      You can select 5, 10, or more EDITED images at once.
                      {canRemoveEditedImages ? ' Use Remove on any edited image added by mistake or rejected after review.' : ''}
                    </p>
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
        const managerGalleryImages = [
          ...rawAssets.map((url, index) => ({ url, label: `RAW Image ${index + 1}`, sku: managerPreview.sku, typeLabel: 'RAW', groupIndex: index })),
          ...editAssets.map((url, index) => ({ url, label: `EDITED Image ${index + 1}`, sku: managerPreview.sku, typeLabel: 'EDITED', groupIndex: index }))
        ];
        const comparisonRows = Array.from({ length: Math.max(rawAssets.length, editAssets.length) }, (_, index) => ({
          raw: rawAssets[index],
          edited: editAssets[index],
          index
        }));
        const openManagerGalleryAt = (url) => {
          const index = managerGalleryImages.findIndex(item => item.url === url);
          const safeIndex = index >= 0 ? index : 0;
          if (managerGalleryImages[safeIndex]) {
            setFullViewImage({ ...managerGalleryImages[safeIndex], images: managerGalleryImages, index: safeIndex });
          }
        };
        const openFullCompareAt = (index = 0) => {
          if (comparisonRows.length === 0) return;
          setFullCompareIndex(Math.min(Math.max(index, 0), comparisonRows.length - 1));
        };
        const renderCompareImage = (url, typeLabel, index) => {
          if (!url) {
            return (
              <div className="h-52 rounded-xl border-2 border-dashed border-gray-200 bg-gray-50 flex items-center justify-center text-[11px] font-black uppercase tracking-wider text-gray-400">
                No {typeLabel} image #{index + 1}
              </div>
            );
          }

          return (
            <button
              type="button"
              onClick={() => openManagerGalleryAt(url)}
              className="relative h-52 rounded-xl overflow-hidden border bg-gray-50 group cursor-zoom-in"
              title={`Open ${typeLabel} image ${index + 1}`}
            >
              <img src={url} alt={`${typeLabel} comparison image ${index + 1}`} className="w-full h-full object-contain bg-white" />
              <span className={`absolute left-3 top-3 px-2 py-1 rounded-lg text-[10px] font-black uppercase border ${typeLabel === 'RAW' ? 'bg-amber-50 text-amber-700 border-amber-200' : 'bg-[rgba(138,21,56,0.06)] text-[#8a1538] border-[rgba(138,21,56,0.18)]'}`}>
                {typeLabel} #{index + 1}
              </span>
              <span className="absolute inset-x-0 bottom-0 bg-black/55 text-white text-[10px] font-black uppercase tracking-wider py-2 opacity-0 group-hover:opacity-100 transition-opacity">
                Click to open full view
              </span>
            </button>
          );
        };
        const renderManagerAssetCard = (url, typeLabel, index) => {
          const filename = `${cleanSku}_${typeLabel.toLowerCase()}_${index + 1}.jpg`;
          return (
            <div key={`${typeLabel}-${url}-${index}`} className="relative group border rounded-xl overflow-hidden bg-gray-50">
              <button
                type="button"
                onClick={() => openManagerGalleryAt(url)}
                className="block w-full cursor-zoom-in"
                title="Open full view"
              >
                <img src={url} className="w-full h-40 object-cover" alt={`${typeLabel} image ${index + 1}`} />
              </button>
              <div className="p-2 flex gap-2 bg-white border-t">
                <button
                  type="button"
                  onClick={() => openManagerGalleryAt(url)}
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
            <div className="bg-white rounded-2xl w-full max-w-7xl p-8 shadow-2xl max-h-[95vh] overflow-y-auto flex flex-col">
              <div className="flex flex-col md:flex-row md:justify-between md:items-start gap-4 mb-6 border-b pb-4">
                <div>
                  <h2 className="text-2xl font-black text-gray-900">{managerPreview.product_name}</h2>
                  <p className="text-xs text-gray-400 font-bold mt-1">SKU: {managerPreview.sku} • Status: {getStatusLabel(managerPreview.status)}</p>
                  <p className="text-[11px] text-gray-400 font-semibold mt-1">Tip: open any image, then use ← / → keyboard keys or the side arrows to move between images.</p>
                  {(authRole === 'Admin' || authRole === 'Manager') && (
                    <div className="mt-3 flex flex-col sm:flex-row sm:items-center gap-2">
                      <span className="text-[10px] font-black uppercase tracking-widest text-gray-400">Status Override</span>
                      <select
                        value={managerPreview.status === 'Completed' ? 'Under Review' : (managerPreview.status || 'Missing')}
                        onChange={(e) => handleOperatorStatusChange(managerPreview.id, e.target.value)}
                        className="w-full sm:w-56 px-3 py-2 border border-gray-200 rounded-xl bg-white text-gray-900 text-[11px] font-black uppercase outline-none"
                      >
                        <option value="Missing">Ready to work</option>
                        <option value="Processing">In Progress</option>
                        <option value="Under Review">Under Review</option>
                        <option value="Rejected" disabled>Rejected - use Reject button</option>
                        <option value="Ready to Upload">Ready to Upload</option>
                        <option value="Modified" disabled>Modified</option>
                      </select>
                    </div>
                  )}
                </div>
                <div className="flex flex-wrap gap-2 items-center">
                  <button
                    type="button"
                    onClick={() => setShowManagerCompare(prev => !prev)}
                    disabled={comparisonRows.length === 0}
                    className={`px-3 py-2 rounded-xl text-[10px] font-black uppercase border transition-all ${showManagerCompare ? 'bg-[#8a1538] text-white border-[#8a1538]' : 'bg-white text-[#8a1538] border-[rgba(138,21,56,0.28)] hover:bg-[rgba(138,21,56,0.06)]'} disabled:opacity-50 disabled:cursor-not-allowed`}
                  >
                    {showManagerCompare ? 'Hide Compare' : 'Compare'}
                  </button>
                  {showManagerCompare && comparisonRows.length > 0 && (
                    <button
                      type="button"
                      onClick={() => openFullCompareAt(0)}
                      className="px-3 py-2 bg-gray-900 text-white border border-gray-900 rounded-xl text-[10px] font-black uppercase hover:bg-black"
                    >
                      Full Screen Compare
                    </button>
                  )}
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

              {showManagerCompare && (
                <div className="mb-8 p-5 rounded-2xl border border-gray-200 bg-gray-50">
                  <div className="flex flex-col sm:flex-row sm:items-end sm:justify-between gap-3 mb-4">
                    <div>
                      <h3 className="text-sm font-black text-gray-900 uppercase tracking-tight">RAW vs EDITED Comparison</h3>
                      <p className="text-xs text-gray-500 font-semibold mt-1">Compare matching RAW and EDITED images side-by-side before approving or rejecting.</p>
                    </div>
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-[10px] font-black uppercase tracking-wider bg-white border border-gray-200 text-gray-500 px-3 py-1 rounded-full">
                        {rawAssets.length} RAW / {editAssets.length} EDITED
                      </span>
                      <button
                        type="button"
                        onClick={() => openFullCompareAt(0)}
                        disabled={comparisonRows.length === 0}
                        className="px-3 py-1.5 bg-gray-900 hover:bg-black disabled:bg-gray-200 disabled:text-gray-400 disabled:cursor-not-allowed text-white rounded-lg text-[10px] font-black uppercase"
                      >
                        Open Full Screen
                      </button>
                    </div>
                  </div>

                  {comparisonRows.length === 0 ? (
                    <div className="h-40 rounded-xl border-2 border-dashed flex items-center justify-center text-xs font-bold text-gray-400 uppercase bg-white">
                      No images available for comparison
                    </div>
                  ) : (
                    <div className="space-y-4 max-h-[620px] overflow-y-auto pr-1">
                      {comparisonRows.map((row) => (
                        <div key={`compare-${row.index}`} className="grid grid-cols-1 lg:grid-cols-2 gap-4 bg-white border rounded-2xl p-4">
                          <div>
                            <div className="flex items-center justify-between mb-2">
                              <div className="text-[10px] font-black uppercase tracking-wider text-amber-700">RAW Image {row.index + 1}</div>
                              <button type="button" onClick={() => openFullCompareAt(row.index)} className="text-[9px] font-black uppercase text-gray-500 hover:text-gray-900">Full Screen Pair</button>
                            </div>
                            {renderCompareImage(row.raw, 'RAW', row.index)}
                          </div>
                          <div>
                            <div className="flex items-center justify-between mb-2">
                              <div className="text-[10px] font-black uppercase tracking-wider text-[#8a1538]">EDITED Image {row.index + 1}</div>
                              <button type="button" onClick={() => openFullCompareAt(row.index)} className="text-[9px] font-black uppercase text-gray-500 hover:text-gray-900">Full Screen Pair</button>
                            </div>
                            {renderCompareImage(row.edited, 'EDITED', row.index)}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

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
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <button onClick={() => setIsRejecting(true)} className="py-3 bg-red-50 text-red-600 hover:bg-red-100 font-bold rounded-xl text-sm border uppercase cursor-pointer transition-colors">Reject</button>
                  <button onClick={() => handleReadyToUploadProduct(managerPreview.id)} className="py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-sm border border-emerald-700 uppercase cursor-pointer transition-colors">Ready to Upload</button>
                  <button onClick={closeManagerPreview} className="py-3 bg-gray-100 hover:bg-gray-200 font-bold rounded-xl text-sm text-gray-600 uppercase cursor-pointer transition-colors">Close</button>
                </div>
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

      {/* FULLSCREEN RAW / EDITED COMPARE MODAL */}
      {managerPreview && fullCompareIndex !== null && (() => {
        const rawAssets = getArray(managerPreview.raw_image_url);
        const editAssets = getArray(managerPreview.edited_image_url);
        const pairCount = Math.max(rawAssets.length, editAssets.length);
        if (pairCount === 0) return null;

        const safeIndex = Math.min(Math.max(Number.isInteger(fullCompareIndex) ? fullCompareIndex : 0, 0), pairCount - 1);
        const rawUrl = rawAssets[safeIndex];
        const editedUrl = editAssets[safeIndex];
        const cleanSku = String(managerPreview.sku || 'UNKNOWN').replace(/[^a-zA-Z0-9_-]/g, '_');
        const moveComparePair = (direction) => {
          setFullCompareIndex(prev => ((Number.isInteger(prev) ? prev : safeIndex) + direction + pairCount) % pairCount);
        };
        const renderLargeCompareImage = (url, typeLabel) => {
          if (!url) {
            return (
              <div className="min-h-[65vh] rounded-2xl border-2 border-dashed border-gray-300 bg-gray-50 flex items-center justify-center text-sm font-black uppercase tracking-wider text-gray-400">
                No {typeLabel} image for pair {safeIndex + 1}
              </div>
            );
          }

          return (
            <div className="relative min-h-[65vh] rounded-2xl overflow-hidden border bg-white flex items-center justify-center">
              <img src={url} alt={`${typeLabel} full compare ${safeIndex + 1}`} className="max-w-full max-h-[72vh] object-contain" />
              <span className={`absolute left-4 top-4 px-3 py-1.5 rounded-xl text-[11px] font-black uppercase border ${typeLabel === 'RAW' ? 'bg-amber-50 text-amber-700 border-amber-200' : 'bg-[rgba(138,21,56,0.06)] text-[#8a1538] border-[rgba(138,21,56,0.18)]'}`}>
                {typeLabel} #{safeIndex + 1}
              </span>
              <button
                type="button"
                onClick={() => handleDownloadSingleAsset(url, `${cleanSku}_${typeLabel.toLowerCase()}_${safeIndex + 1}.jpg`)}
                className="absolute right-4 top-4 px-3 py-1.5 bg-white/90 hover:bg-white text-gray-800 border border-gray-200 rounded-xl text-[10px] font-black uppercase shadow-sm"
              >
                Download
              </button>
            </div>
          );
        };

        return (
          <div className="fixed inset-0 bg-black/90 backdrop-blur-sm flex items-center justify-center p-4 z-[110]">
            <div className="bg-white rounded-2xl w-full max-w-[96vw] max-h-[96vh] overflow-hidden shadow-2xl flex flex-col">
              <div className="p-4 border-b flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
                <div>
                  <h3 className="text-sm font-black text-gray-900 uppercase">Full Screen Compare • {managerPreview.product_name}</h3>
                  <p className="text-[10px] font-bold text-gray-400">
                    SKU: {managerPreview.sku || 'UNKNOWN'} • Pair {safeIndex + 1} of {pairCount} • Use ← / → keys or side buttons
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  {pairCount > 1 && (
                    <>
                      <button onClick={() => moveComparePair(-1)} className="px-3 py-2 bg-gray-100 text-gray-700 rounded-xl text-xs font-black uppercase">← Previous Pair</button>
                      <button onClick={() => moveComparePair(1)} className="px-3 py-2 bg-gray-100 text-gray-700 rounded-xl text-xs font-black uppercase">Next Pair →</button>
                    </>
                  )}
                  <button onClick={() => setFullCompareIndex(null)} className="px-3 py-2 bg-[rgba(138,21,56,0.85)] text-white rounded-xl text-xs font-black uppercase">Close Compare</button>
                </div>
              </div>

              <div className="relative p-4 overflow-auto bg-gray-100">
                {pairCount > 1 && (
                  <button
                    type="button"
                    onClick={() => moveComparePair(-1)}
                    className="absolute left-6 top-1/2 -translate-y-1/2 z-10 w-12 h-12 rounded-full bg-white/95 hover:bg-white text-gray-900 text-3xl font-black shadow-lg flex items-center justify-center"
                    title="Previous pair"
                  >
                    ‹
                  </button>
                )}

                <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
                  <div>
                    <div className="text-[10px] font-black uppercase tracking-wider text-amber-700 mb-2">RAW Image {safeIndex + 1}</div>
                    {renderLargeCompareImage(rawUrl, 'RAW')}
                  </div>
                  <div>
                    <div className="text-[10px] font-black uppercase tracking-wider text-[#8a1538] mb-2">EDITED Image {safeIndex + 1}</div>
                    {renderLargeCompareImage(editedUrl, 'EDITED')}
                  </div>
                </div>

                {pairCount > 1 && (
                  <button
                    type="button"
                    onClick={() => moveComparePair(1)}
                    className="absolute right-6 top-1/2 -translate-y-1/2 z-10 w-12 h-12 rounded-full bg-white/95 hover:bg-white text-gray-900 text-3xl font-black shadow-lg flex items-center justify-center"
                    title="Next pair"
                  >
                    ›
                  </button>
                )}
              </div>
            </div>
          </div>
        );
      })()}

      {/* FULL IMAGE VIEWER MODAL */}
      {fullViewImage && (() => {
        const gallery = Array.isArray(fullViewImage.images) ? fullViewImage.images : [];
        const hasGalleryNavigation = gallery.length > 1;
        const currentGalleryIndex = Number.isInteger(fullViewImage.index) ? fullViewImage.index : gallery.findIndex(item => item.url === fullViewImage.url);
        const displayIndex = currentGalleryIndex >= 0 ? currentGalleryIndex + 1 : 1;
        return (
          <div className="fixed inset-0 bg-black/85 backdrop-blur-sm flex items-center justify-center p-4 z-[100]">
            <div className="bg-white rounded-2xl w-full max-w-7xl max-h-[95vh] overflow-hidden shadow-2xl flex flex-col">
              <div className="p-4 border-b flex items-center justify-between gap-4">
                <div>
                  <h3 className="text-sm font-black text-gray-900 uppercase">{fullViewImage.label}</h3>
                  <p className="text-[10px] font-bold text-gray-400">
                    SKU: {fullViewImage.sku || 'UNKNOWN'}
                    {hasGalleryNavigation ? ` • Image ${displayIndex} of ${gallery.length}` : ''}
                  </p>
                  {hasGalleryNavigation && (
                    <p className="text-[10px] font-bold text-gray-400 mt-0.5">Use keyboard ← / → or the side buttons to move between images.</p>
                  )}
                </div>
                <div className="flex gap-2">
                  {hasGalleryNavigation && (
                    <>
                      <button onClick={() => moveFullViewImage(-1)} className="px-3 py-2 bg-gray-100 text-gray-700 rounded-xl text-xs font-black uppercase">← Prev</button>
                      <button onClick={() => moveFullViewImage(1)} className="px-3 py-2 bg-gray-100 text-gray-700 rounded-xl text-xs font-black uppercase">Next →</button>
                    </>
                  )}
                  <button onClick={() => handleDownloadSingleAsset(fullViewImage.url, `${String(fullViewImage.sku || 'image').replace(/[^a-zA-Z0-9_-]/g, '_')}_${String(fullViewImage.label || 'image').replace(/[^a-zA-Z0-9_-]/g, '_')}.jpg`)} className="px-3 py-2 bg-[rgba(138,21,56,0.85)] text-white rounded-xl text-xs font-black uppercase">Download</button>
                  <button onClick={() => setFullViewImage(null)} className="px-3 py-2 bg-gray-100 text-gray-700 rounded-xl text-xs font-black uppercase">Close</button>
                </div>
              </div>
              <div className="relative p-4 overflow-auto bg-[rgba(138,21,56,1)] flex items-center justify-center min-h-[70vh]">
                {hasGalleryNavigation && (
                  <button
                    type="button"
                    onClick={() => moveFullViewImage(-1)}
                    className="absolute left-4 top-1/2 -translate-y-1/2 z-10 w-12 h-12 rounded-full bg-white/90 hover:bg-white text-gray-900 text-3xl font-black shadow-lg flex items-center justify-center"
                    title="Previous image"
                  >
                    ‹
                  </button>
                )}
                <img src={fullViewImage.url} alt={fullViewImage.label || 'Full image'} className="max-w-full max-h-[78vh] object-contain rounded-lg" />
                {hasGalleryNavigation && (
                  <button
                    type="button"
                    onClick={() => moveFullViewImage(1)}
                    className="absolute right-4 top-1/2 -translate-y-1/2 z-10 w-12 h-12 rounded-full bg-white/90 hover:bg-white text-gray-900 text-3xl font-black shadow-lg flex items-center justify-center"
                    title="Next image"
                  >
                    ›
                  </button>
                )}
              </div>
            </div>
          </div>
        );
      })()}


    </div>
  );
}