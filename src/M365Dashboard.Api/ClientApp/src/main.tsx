import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { initializeMsal } from './services/authConfig';
import './styles/index.css';

// Initialize MSAL before rendering
initializeMsal().then(() => {
  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
});
