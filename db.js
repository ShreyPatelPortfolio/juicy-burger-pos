// db.js — The Juicy Burger
// All requests use GET to avoid CORS issues with Apps Script.
// ---------------------------------------------------------------
// SETUP: Paste your Apps Script Web App URL below.
// ---------------------------------------------------------------

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyfdpm8iUpdBuTb5xPswBM-l2DBxNS7w8k2E3fb5SAsVNaGnMix42us8DjSab1jBRCR/exec";

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
// DB — all writes and reads go through GET params
// Falls back to localStorage if SCRIPT_URL is not set.
// ---------------------------------------------------------------
const DB = {
  _ready: !!SCRIPT_URL,

  // ---- Single GET helper — all traffic goes here -------------
  async _get(params) {
    if (!this._ready) return null;
    const url = SCRIPT_URL + '?' + new URLSearchParams(params).toString();
    const res  = await fetch(url);
    if (!res.ok) throw new Error('Network error ' + res.status);
    const json = await res.json();
    if (json && json.error) throw new Error(json.error);
    return json;
  },

  // ---- Public API --------------------------------------------

  // Save/replace a menu item or order
  async set(path, value) {
    if (!this._ready) { this._lsSet(path, value); this._emit(path, value); return; }
    const parts = path.split('/');
    if (parts[0] === 'menu') {
      return this._get({ action: 'setMenuItem', item: JSON.stringify(value) });
    }
    if (parts[0] === 'orders') {
      return this._get({ action: 'setOrder', order: JSON.stringify(value) });
    }
  },

  // Merge fields into an existing record
  async update(path, fields) {
    if (!this._ready) {
      const existing = this._lsGet(path) || {};
      const merged   = { ...existing, ...fields };
      this._lsSet(path, merged);
      this._emit(path, merged);
      return;
    }
    const parts = path.split('/');
    if (parts[0] === 'orders') {
      // updateOrder merges only the fields you pass — perfect for status changes
      return this._get({ action: 'updateOrder', id: parts[1], fields: JSON.stringify(fields) });
    }
    if (parts[0] === 'menu') {
      // For menu item updates (e.g. toggle available) pass the full merged object
      return this._get({ action: 'setMenuItem', item: JSON.stringify({ id: parts[1], ...fields }) });
    }
  },

  // One-time fetch
  async get(path) {
    if (!this._ready) return this._lsGet(path);
    const parts = path.split('/');
    if (parts[0] === 'menu')   return this._get({ action: 'getMenu' });
    if (parts[0] === 'orders') return this._get({ action: 'getOrders' });
    return null;
  },

  // Polling listener — fires immediately then every 5 s
  on(path, callback) {
    const poll = async () => {
      try {
        const data = await this.get(path);
        callback(data || {});
      } catch(e) {
        console.warn('DB.on poll error:', e.message);
        // Fall back to localStorage snapshot
        callback(this._lsGetAll(path) || this._lsGet(path) || {});
      }
    };

    poll(); // immediate first call

    const interval = setInterval(poll, 5000);
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
    if (parts[0] === 'menu')   return this._get({ action: 'deleteMenuItem', id: parts[1] });
    if (parts[0] === 'orders') return this._get({ action: 'deleteOrder',    id: parts[1] });
  },

  // ---- localStorage fallback ---------------------------------
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
  if (!DB._ready) return;
  try {
    await DB._get({ action: 'seedMenu' });
  } catch(e) {
    console.warn('Seed skipped:', e.message);
  }
}

// ---------------------------------------------------------------
// No-op — kept so pages that call initFirebase() don't break
// ---------------------------------------------------------------
function initFirebase() {
  if (!SCRIPT_URL) {
    console.warn('No SCRIPT_URL — running in localStorage demo mode (single tab only).');
  }
}
