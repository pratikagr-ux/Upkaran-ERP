import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.jsx';

// Show any crash directly on screen
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(e) { return { error: e }; }
  render() {
    if (this.state.error) {
      return (
        <div style={{padding:40,fontFamily:'monospace',color:'#dc2626',background:'#fef2f2',minHeight:'100vh'}}>
          <h2>⚠ App Error</h2>
          <p style={{marginTop:12,fontSize:14}}>{this.state.error.message}</p>
          <pre style={{marginTop:16,fontSize:12,whiteSpace:'pre-wrap',color:'#374151'}}>
            {this.state.error.stack}
          </pre>
        </div>
      );
    }
    return this.props.children;
  }
}

createRoot(document.getElementById('root')).render(
  <ErrorBoundary><App /></ErrorBoundary>
);

// Registers public/sw.js so sound notifications (new sale order / new
// dispatch — see src/db.js's DB.notify()) can show up with sound even
// when this tab isn't focused. import.meta.env.BASE_URL resolves to '/' in
// dev and '/Upkaran-ERP/' in the GitHub Pages build (see vite.config.js),
// so this registers correctly in both. No-op (silently) on browsers
// without Service Worker support (older Safari, etc.) — the rest of the
// app works identically either way, this is purely additive.
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register(`${import.meta.env.BASE_URL}sw.js`).catch(() => {});
  });
}
