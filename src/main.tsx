import React from 'react';
import ReactDOM from 'react-dom/client';
import { ThemeProvider } from './context/ThemeContext';
import { ToastProvider } from './context/ToastContext';
import { SessionFilesProvider } from './context/SessionFilesContext';
import { WorkspaceProvider } from './context/WorkspaceContext';
import App from './App';
import './styles/index.css';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <ThemeProvider>
      <ToastProvider>
        <SessionFilesProvider>
          <WorkspaceProvider>
            <App />
          </WorkspaceProvider>
        </SessionFilesProvider>
      </ToastProvider>
    </ThemeProvider>
  </React.StrictMode>,
);
