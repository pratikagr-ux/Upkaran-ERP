// Drop-in replacement for window.storage → localStorage
// Same async interface so all existing code works unchanged

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
      return true;
    } catch (e) {
      console.error('DB.set failed', key, e);
      return false;
    }
  },
  delete: async (key) => {
    try { localStorage.removeItem(key); return true; }
    catch { return false; }
  }
};
