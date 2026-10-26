// Upkaran ERP notification service worker.
//
// This app is otherwise a plain client-rendered SPA with no offline/caching
// strategy (deliberately — it's a single source of truth synced through
// Supabase, caching pages would just invite stale-data bugs). This worker
// exists for exactly one job: let page JS call
// `registration.showNotification(...)` (see src/db.js's DB.notify()) so a
// "new Sales Order" / "order Dispatched" alert shows up with the OS's own
// notification sound even when this tab isn't focused.
//
// Deliberately NOT using the Push API — no `push` event handler here. That
// would need a real server (Edge Function + VAPID keys + a Database
// Webhook) to send messages from, which is more infra than this feature
// needs. Instead, every call to DB.notify() is triggered directly by page
// JS — either right after a local action, or from store.js's diff on a
// Realtime update from another signed-in device — and this worker's only
// job is to render it. The one real trade-off of that approach: it only
// works while this tab is open somewhere on the device (foreground or a
// backgrounded tab), not with the browser fully closed.
//
// No fetch/caching handlers on purpose — every request still goes straight
// to the network, same as before this file existed.

self.addEventListener('install', () => { self.skipWaiting(); });
self.addEventListener('activate', (event) => { event.waitUntil(self.clients.claim()); });

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clients) => {
      for (const c of clients) {
        if ('focus' in c) return c.focus();
      }
      if (self.clients.openWindow) return self.clients.openWindow('./');
    })
  );
});
