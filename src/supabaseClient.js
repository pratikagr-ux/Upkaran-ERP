// Single Supabase client instance, shared by db.js (data sync) and the
// login screen in App.jsx (auth). VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY
// are the project's public "anon" key — meant to be embedded in a client
// bundle (this app ships to public GitHub Pages). Row-Level Security on the
// kv_store table, not this key, is what actually protects the data; see
// the SQL in the sync-architecture plan for the policy that enforces it.
//
// Set these in a local .env.local (gitignored) for `npm run dev`, and as
// repo secrets injected into the build step for the GitHub Pages deploy —
// see .github/workflows/deploy.yml.
import { createClient } from '@supabase/supabase-js';

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // Don't throw — let the app still boot in pure-local mode (useful for a
  // dev checkout with no Supabase project configured yet). db.js checks
  // `supabase` for null before trying to use it.
  console.warn('[supabaseClient] VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY not set — cloud sync disabled, running local-only.');
}

export const supabase = (url && anonKey) ? createClient(url, anonKey) : null;
