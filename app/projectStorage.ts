export type StoredPhoto = { id: string; name: string; file: File; originalName: string; originalFile: File };
export type StoredStation = { id: string; name: string; notes: string; photos: StoredPhoto[] };
export type StoredProject = { id: string; name: string; activeStationId: string; stations: StoredStation[]; updatedAt: number };

const DB = "fieldnote-projects";
const STORE = "projects";

function database() {
  return new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DB, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE, { keyPath: "id" });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function listProjects(): Promise<StoredProject[]> {
  const db = await database();
  return new Promise((resolve, reject) => {
    const request = db.transaction(STORE).objectStore(STORE).getAll();
    request.onsuccess = () => resolve((request.result as StoredProject[]).sort((a, b) => b.updatedAt - a.updatedAt));
    request.onerror = () => reject(request.error);
  });
}

export async function saveProject(project: StoredProject) {
  const db = await database();
  return new Promise<void>((resolve, reject) => {
    const request = db.transaction(STORE, "readwrite").objectStore(STORE).put(project);
    request.onsuccess = () => resolve(); request.onerror = () => reject(request.error);
  });
}

export async function deleteProject(projectId: string) {
  const db = await database();
  return new Promise<void>((resolve, reject) => {
    const request = db.transaction(STORE, "readwrite").objectStore(STORE).delete(projectId);
    request.onsuccess = () => resolve(); request.onerror = () => reject(request.error);
  });
}
