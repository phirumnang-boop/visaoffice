import React, { createContext, useContext, useState, useEffect } from 'react';

export type WorkspaceScale = 'fit' | '100%' | '90%' | '85%' | '80%' | '75%';

interface WorkspaceSettingsContextType {
  scaleMode: WorkspaceScale;
  setScaleMode: (scale: WorkspaceScale) => void;
  isSidebarCollapsed: boolean;
  setIsSidebarCollapsed: (collapsed: boolean | ((prev: boolean) => boolean)) => void;
  toggleSidebarCollapse: () => void;
  containerWidthMode: 'standard' | 'full';
  setContainerWidthMode: (mode: 'standard' | 'full') => void;
}

const WorkspaceSettingsContext = createContext<WorkspaceSettingsContextType | undefined>(undefined);

const STORAGE_KEY_SCALE = 'app_workspace_scale_mode';
const STORAGE_KEY_SIDEBAR = 'app_sidebar_collapsed';
const STORAGE_KEY_WIDTH = 'app_workspace_width_mode';

export const WorkspaceSettingsProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [scaleMode, setScaleModeState] = useState<WorkspaceScale>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_SCALE);
      if (saved && ['fit', '100%', '90%', '85%', '80%', '75%'].includes(saved)) {
        return saved as WorkspaceScale;
      }
    } catch {}
    return '100%';
  });

  const [isSidebarCollapsed, setIsSidebarCollapsedState] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_SIDEBAR);
      if (saved !== null) {
        return saved === 'true';
      }
    } catch {}
    return false;
  });

  const [containerWidthMode, setContainerWidthModeState] = useState<'standard' | 'full'>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_WIDTH);
      if (saved === 'standard' || saved === 'full') {
        return saved;
      }
    } catch {}
    return 'standard';
  });

  const setScaleMode = (mode: WorkspaceScale) => {
    setScaleModeState(mode);
    try {
      localStorage.setItem(STORAGE_KEY_SCALE, mode);
    } catch {}
  };

  const setIsSidebarCollapsed = (value: boolean | ((prev: boolean) => boolean)) => {
    setIsSidebarCollapsedState((prev) => {
      const next = typeof value === 'function' ? value(prev) : value;
      try {
        localStorage.setItem(STORAGE_KEY_SIDEBAR, String(next));
      } catch {}
      return next;
    });
  };

  const toggleSidebarCollapse = () => {
    setIsSidebarCollapsed((prev) => !prev);
  };

  const setContainerWidthMode = (mode: 'standard' | 'full') => {
    setContainerWidthModeState(mode);
    try {
      localStorage.setItem(STORAGE_KEY_WIDTH, mode);
    } catch {}
  };

  return (
    <WorkspaceSettingsContext.Provider
      value={{
        scaleMode,
        setScaleMode,
        isSidebarCollapsed,
        setIsSidebarCollapsed,
        toggleSidebarCollapse,
        containerWidthMode,
        setContainerWidthMode,
      }}
    >
      {children}
    </WorkspaceSettingsContext.Provider>
  );
};

export const useWorkspaceSettings = (): WorkspaceSettingsContextType => {
  const context = useContext(WorkspaceSettingsContext);
  if (!context) {
    throw new Error('useWorkspaceSettings must be used within a WorkspaceSettingsProvider');
  }
  return context;
};
