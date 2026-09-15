import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";

const clientId = "ced34bd9-a501-4fad-8ba4-7bb5db371c82";
const tenantId = "513f9529-ec84-415b-a0b6-e97302b1524f";
const scopes = ["User.Read", "Files.ReadWrite"];
let clientPromise: Promise<PublicClientApplication> | null = null;

function cleanName(value: string, fallback: string) {
  const cleaned = value.replace(/["*:<>?\\/|]/g, "-").replace(/[. ]+$/g, "").trim();
  return cleaned || fallback;
}

function extension(name: string, fallback: string) {
  const match = name.match(/\.[a-zA-Z0-9]{1,8}$/);
  return match?.[0].toLowerCase() || fallback;
}

async function client() {
  if (!clientPromise) {
    clientPromise = (async () => {
      const app = new PublicClientApplication({
        auth: {
          clientId,
          authority: `https://login.microsoftonline.com/${tenantId}`,
          redirectUri: window.location.origin,
        },
        cache: { cacheLocation: "localStorage" },
      });
      await app.initialize();
      return app;
    })();
  }
  return clientPromise;
}

export async function currentMicrosoftAccount() {
  return (await client()).getAllAccounts()[0] ?? null;
}

export async function signInMicrosoft() {
  const app = await client();
  const existing = app.getAllAccounts()[0];
  if (existing) return existing;
  const response = await app.loginPopup({ scopes, prompt: "select_account" });
  return response.account;
}

async function accessToken(account: AccountInfo) {
  const app = await client();
  try {
    return (await app.acquireTokenSilent({ account, scopes })).accessToken;
  } catch {
    return (await app.acquireTokenPopup({ account, scopes })).accessToken;
  }
}

async function graph(token: string, path: string, init: RequestInit) {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, ...(init.headers ?? {}) },
  });
  if (!response.ok) throw new Error((await response.json().catch(() => null))?.error?.message ?? `OneDrive returned ${response.status}`);
  return response.json();
}

async function graphResponse(token: string, path: string, init: RequestInit = {}) {
  return fetch(`https://graph.microsoft.com/v1.0${path}`, { ...init, headers: { Authorization: `Bearer ${token}`, ...(init.headers ?? {}) } });
}

async function ensureFolder(token: string, path: string, parentPath: string, name: string) {
  const existing = await graphResponse(token, path);
  if (existing.ok) return existing.json();
  if (existing.status !== 404) throw new Error((await existing.json().catch(() => null))?.error?.message ?? "OneDrive folder could not be opened.");
  return createFolder(token, parentPath, name);
}

async function createFolder(token: string, parentPath: string, name: string) {
  return graph(token, parentPath, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "rename" }),
  });
}

export type UploadStation = { name: string; notes: string; photos: Array<{ name: string; file: File; originalName: string; originalFile: File }> };
export type CloudProjectSummary = { id: string; name: string; updatedAt: number; stationCount: number; photoCount: number; folderId: string };
type SyncProject = { id: string; name: string; activeStationId: string; updatedAt: number; stations: Array<{ id: string; name: string; notes: string; photos: Array<{ id: string; name: string; file: File; originalName: string; originalFile: File }> }> };

const savedRootName = "Fieldnote Saved Projects";
const fieldnoteUrl = "https://fieldnote-site-survey.nick-templeton-work.chatgpt.site";

async function uploadProjectShortcut(token: string, folderId: string, projectId: string) {
  const shortcut = `[InternetShortcut]\r\nURL=${fieldnoteUrl}/?project=${encodeURIComponent(projectId)}\r\n`;
  await graph(token, `/me/drive/items/${folderId}:/${encodeURIComponent("Open in Fieldnote.url")}:/content`, { method: "PUT", headers: { "Content-Type": "application/internet-shortcut" }, body: shortcut });
}

export async function syncProjectToOneDrive(account: AccountInfo, project: SyncProject) {
  const token = await accessToken(account);
  const root = await ensureFolder(token, `/me/drive/root:/${encodeURIComponent(savedRootName)}`, "/me/drive/root/children", savedRootName);
  const folderName = `Project-${project.id}`;
  const folder = await ensureFolder(token, `/me/drive/items/${root.id}:/${folderName}`, `/me/drive/items/${root.id}/children`, folderName);
  const filesFolder = await ensureFolder(token, `/me/drive/items/${folder.id}:/Files`, `/me/drive/items/${folder.id}/children`, "Files");
  const manifest = { id: project.id, name: project.name, activeStationId: project.activeStationId, updatedAt: project.updatedAt, stations: [] as Array<Record<string, unknown>> };
  for (const station of project.stations) {
    const savedPhotos = [] as Array<Record<string, unknown>>;
    for (const photo of station.photos) {
      const originalFileName = `${photo.id}-original${extension(photo.originalName, ".jpg")}`;
      await graph(token, `/me/drive/items/${filesFolder.id}:/${originalFileName}:/content`, { method: "PUT", headers: { "Content-Type": photo.originalFile.type || "application/octet-stream" }, body: photo.originalFile });
      let annotatedFileName: string | null = null;
      if (photo.file !== photo.originalFile) {
        annotatedFileName = `${photo.id}-annotated.jpg`;
        await graph(token, `/me/drive/items/${filesFolder.id}:/${annotatedFileName}:/content`, { method: "PUT", headers: { "Content-Type": photo.file.type || "image/jpeg" }, body: photo.file });
      }
      savedPhotos.push({ id: photo.id, name: photo.name, originalName: photo.originalName, originalFileName, annotatedFileName });
    }
    manifest.stations.push({ id: station.id, name: station.name, notes: station.notes, photos: savedPhotos });
  }
  await graph(token, `/me/drive/items/${folder.id}:/project.json:/content`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(manifest) });
  await uploadProjectShortcut(token, folder.id, project.id);
}

export async function listOneDriveProjects(account: AccountInfo): Promise<CloudProjectSummary[]> {
  const token = await accessToken(account);
  const response = await graphResponse(token, `/me/drive/root:/${encodeURIComponent(savedRootName)}:/children`);
  if (response.status === 404) return [];
  if (!response.ok) throw new Error("Saved OneDrive projects could not be loaded.");
  const children = (await response.json()).value as Array<{ id: string; folder?: unknown }>;
  const results = await Promise.all(children.filter((item) => item.folder).map(async (item) => {
    const manifestResponse = await graphResponse(token, `/me/drive/items/${item.id}:/project.json:/content`);
    if (!manifestResponse.ok) return null;
    const data = await manifestResponse.json();
    return { id: data.id, name: data.name, updatedAt: data.updatedAt, stationCount: data.stations.length, photoCount: data.stations.reduce((sum: number, station: { photos: unknown[] }) => sum + station.photos.length, 0), folderId: item.id } as CloudProjectSummary;
  }));
  return results.filter((item): item is CloudProjectSummary => Boolean(item)).sort((a, b) => b.updatedAt - a.updatedAt);
}

export async function loadOneDriveProject(account: AccountInfo, summary: CloudProjectSummary): Promise<SyncProject> {
  const token = await accessToken(account);
  const manifestResponse = await graphResponse(token, `/me/drive/items/${summary.folderId}:/project.json:/content`);
  if (!manifestResponse.ok) throw new Error("This saved project could not be opened.");
  const data = await manifestResponse.json();
  for (const station of data.stations) for (const photo of station.photos) {
    const originalResponse = await graphResponse(token, `/me/drive/items/${summary.folderId}:/Files/${photo.originalFileName}:/content`);
    if (!originalResponse.ok) throw new Error("An original project photo could not be downloaded.");
    photo.originalFile = new File([await originalResponse.blob()], photo.originalName);
    if (photo.annotatedFileName) {
      const annotatedResponse = await graphResponse(token, `/me/drive/items/${summary.folderId}:/Files/${photo.annotatedFileName}:/content`);
      if (!annotatedResponse.ok) throw new Error("An annotated project photo could not be downloaded.");
      photo.file = new File([await annotatedResponse.blob()], photo.name, { type: "image/jpeg" });
    } else photo.file = photo.originalFile;
    delete photo.originalFileName; delete photo.annotatedFileName;
  }
  return data as SyncProject;
}

export async function uploadSurveyToOneDrive(account: AccountInfo, surveyName: string, stations: UploadStation[], onProgress: (message: string) => void, projectId?: string) {
  const token = await accessToken(account);
  const root = await createFolder(token, "/me/drive/root/children", cleanName(surveyName, "Site Survey"));
  for (let index = 0; index < stations.length; index++) {
    const station = stations[index];
    const stationName = cleanName(station.name, `Station ${index + 1}`);
    onProgress(`Creating ${stationName}…`);
    const folder = await createFolder(token, `/me/drive/items/${root.id}/children`, stationName);
    const originals = await createFolder(token, `/me/drive/items/${folder.id}/children`, "Original Photos");
    const annotatedPhotos = station.photos.filter((photo) => photo.file !== photo.originalFile);
    const annotated = annotatedPhotos.length ? await createFolder(token, `/me/drive/items/${folder.id}/children`, "Annotated Photos") : null;
    for (let photoIndex = 0; photoIndex < station.photos.length; photoIndex++) {
      const photo = station.photos[photoIndex];
      const sequence = String(photoIndex + 1).padStart(2, "0");
      const originalName = `${stationName} ${sequence}${extension(photo.originalName, ".jpg")}`;
      const annotatedName = `${stationName} ${sequence} - Annotated.jpg`;
      onProgress(`Uploading ${stationName}: photo ${photoIndex + 1} of ${station.photos.length}`);
      await graph(token, `/me/drive/items/${originals.id}:/${encodeURIComponent(cleanName(originalName, `Photo ${sequence}.jpg`))}:/content`, {
        method: "PUT", headers: { "Content-Type": photo.originalFile.type || "application/octet-stream" }, body: photo.originalFile,
      });
      if (annotated && photo.file !== photo.originalFile) await graph(token, `/me/drive/items/${annotated.id}:/${encodeURIComponent(cleanName(annotatedName, `Photo ${sequence} - Annotated.jpg`))}:/content`, {
        method: "PUT", headers: { "Content-Type": photo.file.type || "image/jpeg" }, body: photo.file,
      });
    }
    const notes = station.notes.trim() || "No field notes were entered for this station.";
    await graph(token, `/me/drive/items/${folder.id}:/Notes.txt:/content`, {
      method: "PUT", headers: { "Content-Type": "text/plain; charset=utf-8" }, body: notes,
    });
  }
  onProgress("Creating sharing link…");
  if (projectId) await uploadProjectShortcut(token, root.id, projectId);
  const share = await graph(token, `/me/drive/items/${root.id}/createLink`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ type: "view", scope: "organization" }),
  });
  return { folderUrl: root.webUrl as string, shareUrl: share.link.webUrl as string };
}
