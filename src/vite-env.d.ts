/// <reference types="vite/client" />

interface DesktopFile {
  path: string;
  name: string;
  extension: string;
  data: Uint8Array;
}

interface Window {
  desktop?: {
    platform: string;
    openFile(): Promise<DesktopFile | null>;
    readPath(path: string): Promise<DesktopFile>;
    saveFile(request: { path?: string; saveAs?: boolean; extension: string; suggestedName: string; data: Uint8Array }): Promise<{ path: string; name: string; extension: string } | null>;
    print(): Promise<{ success: boolean; failureReason?: string }>;
    checkForUpdates(): Promise<void>;
    installUpdate(): Promise<void>;
    onCommand(callback: (command: string, payload?: unknown) => void): () => void;
  };
}

type UpdateState = "checking" | "available" | "downloading" | "downloaded" | "none" | "error" | "dev";
interface UpdateStatus { state: UpdateState; version?: string; percent?: number; message?: string }
