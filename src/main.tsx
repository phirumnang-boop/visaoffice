import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import { WorkspaceSettingsProvider } from './context/WorkspaceSettingsContext';
import './index.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <WorkspaceSettingsProvider>
      <App />
    </WorkspaceSettingsProvider>
  </StrictMode>,
);
