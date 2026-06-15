// db.js — The Juicy Burger
// Google Sheets backend via Apps Script Web App.
// ---------------------------------------------------------------
// SETUP: Paste your Apps Script Web App URL below.
// ---------------------------------------------------------------

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyfdpm8iUpdBuTb5xPswBM-l2DBxNS7w8k2E3fb5SAsVNaGnMix42us8DjSab1jBRCR/exec"; // ← Paste your Web App URL here

// ---------------------------------------------------------------
// Tax rates (Ontario, Canada)
// ---------------------------------------------------------------
const GST_RATE = 0.05;
const PST_RATE = 0.08;

function calcTotals(subtotal) {
  const gst = subtotal * GST_RATE;
  const pst = subtotal * PST_RATE;
  return { subtotal, gst, pst, total: subtotal + gst + pst };
}

function formatCAD(amount) {
  return '$' + Number(amount).toFixed(2);
}

function generateOrderId() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `ORD-${pad(now.getHours())}${pad(now.getMinutes())}-${Math.floor(Math.random() * 900 + 100)}`;
}

function showToast(msg, type = '') {
  let container = document.getElementById('toastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toastContainer';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = msg;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 3500);
}

// ---------------------------------------------------------------
// DB — thin wrapper around the Apps Script Web App
// Falls back to localStorage if SCRIPT_URL is not set.
// ---------------------------------------------------------------
const DB = {
  _ready: !!SCRIPT_URL,

  // ---- Internal fetch helpers --------------------------------
  async _get(params) {
    if (!this._ready) return null;
    const url = SCRIPT_URL + '?' + new URLSearchParams(params).toString();
    const res = await fetch(url);
    const json = await res.json();
    if (json.error) throw new Error(json.error);
    return json;
  },

  async _post(body) {
    if (!this._ready) return null;
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    const json = await res.json();
    if (json.error) throw new Error(json.error);
    return json;
  },

  // ---- Public API (mirrors the old Firebase DB wrapper) ------

  // Save/replace a full order or menu item by path
  async set(path, value) {
    if (!this._ready) { this._lsSet(path, value); this._emit(path, value); return; }
    const parts = path.split('/');
    if (parts[0] === 'menu')   return this._post({ action: 'setMenuItem', item: value });
    if (parts[0] === 'orders') return this._post({ action: 'setOrder',    order: value });
  },

  // Merge fields into an existing order
  async update(path, fields) {
    if (!this._ready) {
      const existing = this._lsGet(path) || {};
      const merged = { ...existing, ...fields };
      this._lsSet(path, merged);
      this._emit(path, merged);
      return;
    }
    const parts = path.split('/');
    if (parts[0] === 'orders') return this._post({ action: 'updateOrder', id: parts[1], fields });
    if (parts[0] === 'menu')   return this._post({ action: 'setMenuItem', item: { ...fields, id: parts[1] } });
  },

  // One-time fetch
  async get(path) {
    if (!this._ready) return this._lsGet(path);
    const parts = path.split('/');
    if (parts[0] === 'menu')   return this._get({ action: 'getMenu' });
    if (parts[0] === 'orders') return this._get({ action: 'getOrders' });
    return null;
  },

  // Polling listener — calls callback(data) immediately and every 5s
  on(path, callback) {
    const fetchAndCall = async () => {
      try {
        const data = await this.get(path);
        callback(data || {});
      } catch(e) {
        console.warn('DB.on poll error:', e.message);
        // Fall back to localStorage
        callback(this._lsGetAll(path) || this._lsGet(path) || {});
      }
    };

    fetchAndCall(); // immediate

    // Poll every 5 seconds for live updates across terminals
    const interval = setInterval(fetchAndCall, 5000);

    // Store so callers could clear if needed (not required for basic use)
    this._intervals = this._intervals || [];
    this._intervals.push(interval);
  },

  async remove(path) {
    if (!this._ready) {
      const prefix = 'tjb_' + path.replace(/\//g, '_');
      Object.keys(localStorage)
        .filter(k => k === prefix || k.startsWith(prefix + '_'))
        .forEach(k => localStorage.removeItem(k));
      this._emit(path, null);
      return;
    }
    const parts = path.split('/');
    if (parts[0] === 'menu')   return this._post({ action: 'deleteMenuItem', id: parts[1] });
    if (parts[0] === 'orders') return this._post({ action: 'deleteOrder',    id: parts[1] });
  },

  // ---- localStorage fallback (demo / no URL set) -------------
  _lsKey(path)  { return 'tjb_' + path.replace(/\//g, '_'); },
  _lsSet(path, value) { localStorage.setItem(this._lsKey(path), JSON.stringify(value)); },
  _lsGet(path)  {
    const raw = localStorage.getItem(this._lsKey(path));
    try { return raw ? JSON.parse(raw) : null; } catch { return null; }
  },
  _lsGetAll(path) {
    const prefix = this._lsKey(path) + '_';
    const result = {};
    Object.keys(localStorage).forEach(k => {
      if (k.startsWith(prefix)) {
        try {
          const val = JSON.parse(localStorage.getItem(k));
          const id  = val._key || k.slice(prefix.length);
          result[id] = val;
        } catch {}
      }
    });
    return Object.keys(result).length ? result : null;
  },
  _emit(path, value) {
    window.dispatchEvent(new CustomEvent('localdb_change', { detail: { path, value } }));
  }
};

// ---------------------------------------------------------------
// Seed default menu on first load if sheet is empty
// ---------------------------------------------------------------
async function seedMenuIfEmpty() {
  if (!DB._ready) {
    // localStorage fallback — seed if empty
    const menu = DB._lsGet('menu');
    if (menu && Object.keys(menu).length) return;
    // minimal seed via localStorage
    return;
  }
  try {
    await DB._post({ action: 'seedMenu' });
  } catch(e) {
    console.warn('Seed failed:', e.message);
  }
}

// ---------------------------------------------------------------
// No-op — kept so pages that call initFirebase() don't break
// ---------------------------------------------------------------
function initFirebase() {
  if (!SCRIPT_URL) {
    console.warn('No SCRIPT_URL set in db.js — running in localStorage demo mode (single tab only).');
  }
}
