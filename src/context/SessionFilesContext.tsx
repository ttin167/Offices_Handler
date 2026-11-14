import React, { createContext, useCallback, useContext, useState } from 'react';
import type { SessionFile } from '../types/workspace';
import {
  buildDownloadUrl,
  buildPreviewUrl,
  extractBaseFilename,
  getExtension,
  purgeSessionTraces,
} from '../services/api';

interface AddSessionFileInput {
  filePath: string;
  filename?: string;
  size?: number;
  previewUrl?: string;
  downloadUrl?: string;
  sourceLabel: string;
}

interface SessionFilesContextValue {
  sessionFiles: SessionFile[];
  addSessionFile: (input: AddSessionFileInput) => SessionFile;
  removeSessionFile: (id: string) => void;
  clearShelfOnly: () => void;
  purgeAllTraces: (purgeHistory?: boolean) => Promise<{ removedFiles: number; mbFreed: number }>;
  isPurging: boolean;
}

const SessionFilesContext = createContext<SessionFilesContextValue | null>(null);

export function SessionFilesProvider({ children }: { children: React.ReactNode }) {
  const [sessionFiles, setSessionFiles] = useState<SessionFile[]>([]);
  const [isPurging, setIsPurging] = useState(false);

  const addSessionFile = useCallback((input: AddSessionFileInput): SessionFile => {
    const filename = input.filename || extractBaseFilename(input.filePath);
    const ext = getExtension(filename);
    const newItem: SessionFile = {
      id: `sf_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      filename,
      filePath: input.filePath,
      ext,
      size: input.size,
      previewUrl: input.previewUrl || buildPreviewUrl(filename),
      downloadUrl: input.downloadUrl || buildDownloadUrl(filename),
      sourceLabel: input.sourceLabel,
      addedAt: Date.now(),
    };

    setSessionFiles((prev) => {
      const filtered = prev.filter((item) => item.filePath !== input.filePath);
      return [newItem, ...filtered].slice(0, 15);
    });

    return newItem;
  }, []);

  const removeSessionFile = useCallback((id: string) => {
    setSessionFiles((prev) => prev.filter((item) => item.id !== id));
  }, []);

  const clearShelfOnly = useCallback(() => {
    setSessionFiles([]);
  }, []);

  const purgeAllTraces = useCallback(async (purgeHistory = false) => {
    setIsPurging(true);
    try {
      const res = await purgeSessionTraces(purgeHistory);
      setSessionFiles([]);
      return {
        removedFiles: res.removed_files || 0,
        mbFreed: res.mb_freed || 0,
      };
    } finally {
      setIsPurging(false);
    }
  }, []);

  return (
    <SessionFilesContext.Provider
      value={{
        sessionFiles,
        addSessionFile,
        removeSessionFile,
        clearShelfOnly,
        purgeAllTraces,
        isPurging,
      }}
    >
      {children}
    </SessionFilesContext.Provider>
  );
}

export function useSessionFiles(): SessionFilesContextValue {
  const ctx = useContext(SessionFilesContext);
  if (!ctx) throw new Error('useSessionFiles must be used inside SessionFilesProvider');
  return ctx;
}
