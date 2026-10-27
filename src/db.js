// Drop-in replacement for window.storage → localStorage, now backed by a
// real cloud store (Supabase) for cross-device sync — localStorage stays as
// the fast local cache + offline write queue. Same async get/set/delete
// interface as before, so every existing call site (store.js's `save`
// object, loadAll(), loadTemplates()) works completely unchanged.
//
//  - get(key): always reads localStorage — instant, works offline. By the
//    time anything calls this, initSync() below has already pulled the
//    latest remote snapshot into localStorage (see App.jsx boot sequence).
//  - set(key, value)/delete(key): write localStorage immediately, then
//    best-effort push to Supabase. A failed push (offline, or signed out)
//    is queued and retried on reconnect — no write is ever lost, it just
//    syncs late.
//  - onRemoteUpdate(cb): fires cb(key, value) when ANOTHER device's write
//    arrives via Supabase Realtime (this tab's own writes are filtered out
//    via a per-tab origin id, so a save doesn't boomerang back to itself).
//  - initSync(): call once after sign-in, BEFORE store.js's loadAll() —
//    pulls this user's full remote snapshot into localStorage (remote wins,
//    it's the cross-device source of truth) and opens the realtime feed.
//  - pushAllLocal(): one-time manual action — re-uploads every `data:*` key
//    already on THIS device. Used once, from the device holding the real
//    accumulated business data, to seed the cloud table for the first time.
//  - signIn/signOut/getSession/onAuthChange: thin wraps over supabase.auth,
//    consumed only by the login screen in App.jsx.
//
// If VITE_SUPABASE_URL/VITE_SUPABASE_ANON_KEY aren't configured, `supabase`
// is null and every cloud operation here becomes a no-op — the app still
// works, purely local, same as before this change (useful for a dev
// checkout with no Supabase project set up yet).

import { supabase } from './supabaseClient.js';

const TABLE = 'kv_store';
// Supabase Storage bucket for catalogue photos (Feature #2). Must be created
// once in the Supabase dashboard — see the setup note alongside TeamModule's
// equivalent one-time doc. Public-read, authenticated-upload RLS, same spirit
// as kv_store's row-level policy.
const PHOTO_BUCKET = 'item-photos';
const QUEUE_KEY = '__pendingSync';
// Per-tab id so a realtime echo of this tab's own write can be told apart
// from a genuine change made elsewhere.
const ORIGIN_ID = (crypto.randomUUID ? crypto.randomUUID() : String(Math.random()) + Date.now());

function readQueue() {
  try { return JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]'); } catch { return []; }
}
function writeQueue(q) {
  try { localStorage.setItem(QUEUE_KEY, JSON.stringify(q)); } catch {}
}
function enqueue(op) {
  const q = readQueue().filter(o => o.key !== op.key); // de-dup — only the latest write per key matters
  q.push(op);
  writeQueue(q);
}

async function currentUserId() {
  if (!supabase) return null;
  const { data } = await supabase.auth.getSession();
  return data?.session?.user?.id || null;
}

async function pushOne(key, value, isDelete) {
  if (!supabase) return false;
  const uid = await currentUserId();
  if (!uid) return false;
  try {
    if (isDelete) {
      const { error } = await supabase.from(TABLE).delete().eq('user_id', uid).eq('key', key);
      if (error) throw error;
    } else {
      const { error } = await supabase.from(TABLE).upsert({
        user_id: uid, key, value, updated_at: new Date().toISOString(), origin_id: ORIGIN_ID,
      });
      if (error) throw error;
    }
    return true;
  } catch (e) {
    console.warn('[db] cloud push failed, queued for retry:', key, e?.message || e);
    return false;
  }
}

export async function flushQueue() {
  const q = readQueue();
  if (!q.length) return;
  const remaining = [];
  for (const op of q) {
    const ok = await pushOne(op.key, op.value, op.isDelete);
    if (!ok) remaining.push(op);
  }
  writeQueue(remaining);
}

if (typeof window !== 'undefined') {
  window.addEventListener('online', () => { flushQueue(); });
  // Belt-and-suspenders sweep — some mobile browsers don't fire 'online'
  // reliably. Cheap no-op once the queue is empty.
  setInterval(() => { flushQueue(); }, 30000);
}

const remoteUpdateListeners = new Set();
let realtimeChannel = null;
function subscribeRealtime(uid) {
  if (!supabase || realtimeChannel) return;
  realtimeChannel = supabase
    .channel('kv_store_changes')
    .on('postgres_changes', { event: '*', schema: 'public', table: TABLE, filter: `user_id=eq.${uid}` }, (payload) => {
      const row = payload.new && Object.keys(payload.new).length ? payload.new : payload.old;
      if (!row?.key) return;
      if (row.origin_id === ORIGIN_ID) return; // our own write echoing back
      if (payload.eventType === 'DELETE') {
        try { localStorage.removeItem(row.key); } catch {}
      } else {
        try { localStorage.setItem(row.key, JSON.stringify(row.value)); } catch {}
      }
      remoteUpdateListeners.forEach(fn => fn(row.key, row.value));
    })
    .subscribe();
}

export const DB = {
  get: async (key) => {
    try {
      const v = localStorage.getItem(key);
      return v ? JSON.parse(v) : null;
    } catch { return null; }
  },
  set: async (key, value) => {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (e) {
      console.error('DB.set failed (local)', key, e);
      return false;
    }
    const ok = await pushOne(key, value, false);
    if (!ok) enqueue({ key, value, isDelete: false });
    return true;
  },
  delete: async (key) => {
    try { localStorage.removeItem(key); } catch { return false; }
    const ok = await pushOne(key, null, true);
    if (!ok) enqueue({ key, value: null, isDelete: true });
    return true;
  },

  // ── Auth (used only by the login screen) ───────────────────────────────
  signIn: async (email, password) => {
    if (!supabase) throw new Error('Cloud sync is not configured for this build.');
    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) throw error;
    return data.session;
  },
  signOut: async () => {
    if (!supabase) return;
    await supabase.auth.signOut();
  },
  getSession: async () => {
    if (!supabase) return null;
    const { data } = await supabase.auth.getSession();
    return data?.session || null;
  },
  onAuthChange: (cb) => {
    if (!supabase) return () => {};
    const { data } = supabase.auth.onAuthStateChange((_event, session) => cb(session));
    return () => data?.subscription?.unsubscribe();
  },

  // ── Cross-device sync ───────────────────────────────────────────────────
  initSync: async () => {
    if (!supabase) return;
    const uid = await currentUserId();
    if (!uid) return;
    try {
      const { data, error } = await supabase.from(TABLE).select('key,value').eq('user_id', uid);
      if (error) throw error;
      for (const row of (data || [])) {
        try { localStorage.setItem(row.key, JSON.stringify(row.value)); } catch {}
      }
    } catch (e) {
      console.warn('[db] initSync pull failed — continuing with whatever is cached locally', e?.message || e);
    }
    await flushQueue();
    subscribeRealtime(uid);
  },

  onRemoteUpdate: (cb) => { remoteUpdateListeners.add(cb); return () => remoteUpdateListeners.delete(cb); },

  // One-time manual action (see the "Upload local data to cloud" button in
  // Settings) — re-pushes every `data:*` key held in THIS device's
  // localStorage up to the cloud. Run once, from the device that already
  // has the real accumulated business data.
  pushAllLocal: async () => {
    const uid = await currentUserId();
    if (!uid) throw new Error('Sign in first.');
    let n = 0;
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (!k || !k.startsWith('data:')) continue;
      let v;
      try { v = JSON.parse(localStorage.getItem(k)); } catch { continue; }
      if (await pushOne(k, v, false)) n++;
    }
    return n;
  },

  // ── Catalogue photos (Feature #2) ───────────────────────────────────────
  // Uploads to the `item-photos` Storage bucket and returns its public URL
  // (stored in item.photos[] by the caller, same pattern as any other
  // state field — no new sync mechanism needed since the URL itself is
  // just a string that rides along with the item through the normal
  // kv_store sync path).
  uploadPhoto: async (file, itemId) => {
    if (!supabase) throw new Error('Cloud sync is not configured for this build.');
    const uid = await currentUserId();
    if (!uid) throw new Error('Sign in first.');
    const ext = (file.name && file.name.includes('.')) ? file.name.split('.').pop().toLowerCase() : 'jpg';
    const path = `${uid}/${itemId || 'misc'}/${Date.now()}_${Math.random().toString(36).slice(2, 8)}.${ext}`;
    const { error } = await supabase.storage.from(PHOTO_BUCKET).upload(path, file, { cacheControl: '3600', upsert: false });
    if (error) throw error;
    const { data } = supabase.storage.from(PHOTO_BUCKET).getPublicUrl(path);
    return data.publicUrl;
  },
  deletePhoto: async (url) => {
    if (!supabase || !url) return;
    try {
      const marker = `/${PHOTO_BUCKET}/`;
      const idx = url.indexOf(marker);
      if (idx === -1) return;
      const path = url.slice(idx + marker.length);
      await supabase.storage.from(PHOTO_BUCKET).remove([path]);
    } catch (e) {
      console.warn('[db] deletePhoto failed', e?.message || e);
    }
  },

  // ── Sound notifications (new Sales Order / new Dispatch) ───────────────
  // Deliberately NOT real Web Push — that needs a server (Edge Function +
  // VAPID keys + a Database Webhook), real infra to stand up and maintain.
  // Instead this rides the cross-device sync that already exists: every
  // device's Supabase Realtime subscription (see onRemoteUpdate below)
  // already receives every order-array write from every other signed-in
  // device, so store.js just diffs old-vs-new and calls DB.notify() —
  // which shows a browser Notification (with the OS's own notification
  // sound) via the Service Worker already registered in main.jsx. Zero
  // Supabase setup, zero secrets. The one real trade-off: this only fires
  // while the tab is open somewhere on that device (foreground or a
  // backgrounded tab) — a fully closed browser gets nothing, since that
  // needs true push and there's no way around it.
  notificationsSupported: () => typeof window !== 'undefined' && 'serviceWorker' in navigator && typeof Notification !== 'undefined',

  // 'unsupported' | 'default' (not yet asked) | 'granted' | 'denied'
  getNotificationPermission: () => {
    if (!DB.notificationsSupported()) return 'unsupported';
    return Notification.permission;
  },

  isNotificationsMuted: () => { try { return localStorage.getItem('__notifyMuted') === '1'; } catch { return false; } },
  setNotificationsMuted: (muted) => { try { localStorage.setItem('__notifyMuted', muted ? '1' : '0'); } catch {} },

  enableNotifications: async () => {
    if (!DB.notificationsSupported()) throw new Error('Notifications aren\'t supported in this browser.');
    const perm = await Notification.requestPermission();
    if (perm !== 'granted') throw new Error('Notification permission was not granted.');
    DB.setNotificationsMuted(false);
    return true;
  },

  // There's no browser API to revoke a permission already granted — only
  // the user can do that, from their own browser/site settings. "Disable"
  // here is a local mute flag DB.notify() checks, so it still has a real
  // effect without sending the user hunting through OS settings.
  disableNotifications: () => { DB.setNotificationsMuted(true); },

  // Core notify — called from store.js whenever a new order/dispatch is
  // detected (either locally, right after the action, or from another
  // device's Realtime update). Silently no-ops if unsupported/not
  // permitted/muted, so every call site can fire-and-forget this.
  notify: (title, body) => {
    if (!DB.notificationsSupported()) return;
    if (Notification.permission !== 'granted') return;
    if (DB.isNotificationsMuted()) return;
    // tag is unique per call (not just per-body) so a second notification
    // whose text happens to match an earlier still-visible one (e.g. the
    // same order's "New Sales Order" then "Dispatched" alerts, same
    // party/item/total) doesn't silently coalesce into it — a repeated tag
    // replaces the existing notification without re-alerting (no sound) by
    // default, which would look exactly like "nothing happened".
    navigator.serviceWorker.ready.then(reg => reg.showNotification(title, {
      body, icon: 'icon-192-v2.png', badge: 'icon-192-v2.png', vibrate: [200, 100, 200], tag: `${title}-${Date.now()}`,
    })).catch(e => console.warn('[notify] showNotification failed', e?.message || e));
  },

  testNotification: async () => {
    if (!DB.notificationsSupported()) throw new Error('Notifications aren\'t supported in this browser.');
    if (Notification.permission !== 'granted') throw new Error('Enable notifications first.');
    const reg = await navigator.serviceWorker.ready;
    await reg.showNotification('Upkaran ERP', {
      body: 'Test alert — if you heard/saw this, sound notifications are working on this device.',
      icon: 'icon-192-v2.png',
      vibrate: [200, 100, 200],
      tag: `test-${Date.now()}`,
    });
  },
};
