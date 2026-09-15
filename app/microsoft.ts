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

async function createFolder(token: string, parentPath: string, name: string) {
  return graph(token, parentPath, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "rename" }),
  });
}

export type UploadStation = { name: string; notes: string; photos: Array<{ name: string; file: File; originalName: string; originalFile: File }> };

export async function uploadSurveyToOneDrive(account: AccountInfo, surveyName: string, stations: UploadStation[], onProgress: (message: string) => void) {
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
  const share = await graph(token, `/me/drive/items/${root.id}/createLink`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ type: "view", scope: "organization" }),
  });
  return { folderUrl: root.webUrl as string, shareUrl: share.link.webUrl as string };
}
