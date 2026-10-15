export type StoredPhoto = { id: string; name: string; file: File; originalName: string; originalFile: File };
export type StoredStation = { id: string; name: string; notes: string; photos: StoredPhoto[] };
export type StoredProject = { id: string; name: string; activeStationId: string; stations: StoredStation[]; updatedAt: number };
export type UploadJob = {
  id: string;
  projectId: string;
  projectName: string;
  status: "queued" | "uploading" | "retrying" | "complete" | "needs-sign-in";
  attempts: number;
  nextAttemptAt: number;
  progress: string;
  error: string;
  folderId?: string;
  folderUrl?: string;
  shareUrl?: string;
  warnings: string[];
  completedFiles: string[];
  updatedAt: number;
};

const DB = "fieldnote-projects";
const PROJECT_STORE = "projects";
const UPLOAD_STORE = "uploads";

function database() {
  return new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DB, 2);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(PROJECT_STORE)) request.result.createObjectStore(PROJECT_STORE, { keyPath: "id" });
      if (!request.result.objectStoreNames.contains(UPLOAD_STORE)) request.result.createObjectStore(UPLOAD_STORE, { keyPath: "id" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function listProjects(): Promise<StoredProject[]> {
  const db = await database();
  return new Promise((resolve, reject) => {
    const request = db.transaction(PROJECT_STORE).objectStore(PROJECT_STORE).getAll();
    request.onsuccess = () => resolve((request.result as StoredProject[]).sort((a, b) => b.updatedAt - a.updatedAt));
    request.onerror = () => reject(request.error);
  });
}

export async function getProject(projectId: string): Promise<StoredProject | undefined> {
  const db = await database();
  return new Promise((resolve, reject) => {
    const request = db.transaction(PROJECT_STORE).objectStore(PROJECT_STORE).get(projectId);
    request.onsuccess = () => resolve(request.result as StoredProject | undefined);
    request.onerror = () => reject(request.error);
  });
}

export async function saveProject(project: StoredProject) {
  const db = await database();
  return new Promise<void>((resolve, reject) => {
    const request = db.transaction(PROJECT_STORE, "readwrite").objectStore(PROJECT_STORE).put(project);
    request.onsuccess = () => resolve(); request.onerror = () => reject(request.error);
  });
}

export async function deleteProject(projectId: string) {
  const db = await database();
  return new Promise<void>((resolve, reject) => {
    const request = db.transaction(PROJECT_STORE, "readwrite").objectStore(PROJECT_STORE).delete(projectId);
    request.onsuccess = () => resolve(); request.onerror = () => reject(request.error);
  });
}

export async function listUploadJobs(): Promise<UploadJob[]> {
  const db = await database();
  return new Promise((resolve, reject) => {
    const request = db.transaction(UPLOAD_STORE).objectStore(UPLOAD_STORE).getAll();
    request.onsuccess = () => resolve((request.result as UploadJob[]).sort((a, b) => b.updatedAt - a.updatedAt));
    request.onerror = () => reject(request.error);
  });
}

export async function saveUploadJob(job: UploadJob) {
  const db = await database();
  return new Promise<void>((resolve, reject) => {
    const request = db.transaction(UPLOAD_STORE, "readwrite").objectStore(UPLOAD_STORE).put(job);
    request.onsuccess = () => resolve(); request.onerror = () => reject(request.error);
  });
}
