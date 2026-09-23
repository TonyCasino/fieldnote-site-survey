"use client";
import "./walkthrough.css";
import "./microsoft.css";

import { useEffect, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import { Camera, Check, ChevronLeft, Cloud, FolderOpen, ImagePlus, MapPin, PenLine, Plus, Share2, Trash2, Upload } from "lucide-react";
import { currentMicrosoftAccount, listOneDriveProjects, loadOneDriveProject, signInMicrosoft, syncProjectToOneDrive, uploadSurveyToOneDrive, type CloudProjectSummary } from "./microsoft";
import PhotoAnnotator from "./PhotoAnnotator";
import { deleteProject, listProjects, saveProject, type StoredProject } from "./projectStorage";

type Photo = { id: string; name: string; url: string; file: File; originalName: string; originalFile: File };
type Station = { id: string; name: string; notes: string; photos: Photo[] };
const id = () => Math.random().toString(36).slice(2, 9);

export default function Home() {
  const [surveyName, setSurveyName] = useState("");
  const [projectId, setProjectId] = useState("");
  const [savedProjects, setSavedProjects] = useState<StoredProject[]>([]);
  const [cloudProjects, setCloudProjects] = useState<CloudProjectSummary[]>([]);
  const [cloudStatus, setCloudStatus] = useState("");
  const [localStatus, setLocalStatus] = useState("Loading saved projects…");
  const [draftName, setDraftName] = useState("");
  const [creating, setCreating] = useState(false);
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [active, setActive] = useState("");
  const [stations, setStations] = useState<Station[]>([]);
  const [completed, setCompleted] = useState(false);
  const [uploadStatus, setUploadStatus] = useState("");
  const [uploadError, setUploadError] = useState("");
  const [uploaded, setUploaded] = useState<{ folderUrl: string; shareUrl: string; warnings: string[] } | null>(null);
  const [editingPhoto, setEditingPhoto] = useState<Photo | null>(null);
  const photoInput = useRef<HTMLInputElement>(null);
  const openedLink = useRef(false);
  const completingUpload = useRef(false);
  const skipRestoredProjectSync = useRef(false);
  const current = stations.find((station) => station.id === active);
  const totalPhotos = stations.reduce((total, station) => total + station.photos.length, 0);
  const connected = Boolean(account);

  useEffect(() => {
    navigator.storage?.persist?.().catch(() => undefined);
    currentMicrosoftAccount().then((value) => { setAccount(value); if (value) listOneDriveProjects(value).then(setCloudProjects).catch(() => undefined); }).catch(() => undefined);
    listProjects().then((projects) => { setSavedProjects(projects); setLocalStatus("Saved on this device"); }).catch(() => setLocalStatus("Device save unavailable"));
  }, []);
  useEffect(() => {
    const linkedId = new URLSearchParams(window.location.search).get("project");
    if (!account || !linkedId || openedLink.current) return;
    openedLink.current = true; setCloudStatus("Opening project from OneDrive…");
    (async () => {
      try {
        const projects = await listOneDriveProjects(account), summary = projects.find((project) => project.id === linkedId);
        if (!summary) throw new Error("This Fieldnote project was not found in your OneDrive.");
        const loaded = await loadOneDriveProject(account, summary);
        const restored = loaded.stations.map((station) => ({ ...station, photos: station.photos.map((photo) => normalizePhoto(photo)) }));
        skipRestoredProjectSync.current = true; setCloudProjects(projects); setProjectId(loaded.id); setSurveyName(loaded.name); setStations(restored); setActive(loaded.activeStationId || restored[0]?.id || ""); setCloudStatus("Saved to OneDrive");
        window.history.replaceState({}, "", window.location.pathname);
      } catch (error) { setCloudStatus(error instanceof Error ? error.message : "The linked project could not be opened."); }
    })();
  }, [account]);
  useEffect(() => {
    if (!projectId || !surveyName || !stations.length) return;
    setLocalStatus("Saving on this device…");
    const timer = window.setTimeout(async () => {
      await saveProject({ id: projectId, name: surveyName, activeStationId: active, stations: stations.map((station) => ({ ...station, photos: station.photos.map(({ url: _url, ...photo }) => photo) })), updatedAt: Date.now() });
      setSavedProjects(await listProjects());
      setLocalStatus(`Saved at ${new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`);
    }, 350);
    return () => window.clearTimeout(timer);
  }, [projectId, surveyName, active, stations]);
  useEffect(() => {
    if (!account || !projectId || !surveyName || !stations.length) return;
    if (skipRestoredProjectSync.current) { skipRestoredProjectSync.current = false; return; }
    const timer = window.setTimeout(async () => {
      if (completingUpload.current) return;
      setCloudStatus("Syncing to OneDrive…");
      try {
        await syncProjectToOneDrive(account, { id: projectId, name: surveyName, activeStationId: active, stations, updatedAt: Date.now() });
        setCloudStatus("Saved to OneDrive");
      } catch { setCloudStatus("OneDrive sync paused"); }
    }, 20000);
    return () => window.clearTimeout(timer);
  }, [account, projectId, surveyName, active, stations]);

  const connectMicrosoft = async () => {
    setUploadError("");
    try { const signedIn = await signInMicrosoft(); setAccount(signedIn); setCloudProjects(await listOneDriveProjects(signedIn)); }
    catch (error) { setUploadError(error instanceof Error ? error.message : "Microsoft sign-in did not finish."); }
  };

  const uploadSurvey = async () => {
    if (!account) { await connectMicrosoft(); return; }
    completingUpload.current = true;
    setUploadError(""); setUploadStatus("Preparing your OneDrive folders…");
    try { setUploaded(await uploadSurveyToOneDrive(account, surveyName, stations, setUploadStatus, projectId)); setUploadStatus(""); }
    catch (error) { setUploadStatus(""); setUploadError(`${error instanceof Error ? error.message : "The OneDrive upload could not be completed."} Your walkthrough is still saved and safe. You can close this message and try again.`); }
    finally { completingUpload.current = false; }
  };

  const createSurvey = () => {
    if (!draftName.trim()) return;
    const first = { id: id(), name: "", notes: "", photos: [] };
    setProjectId(id()); setSurveyName(draftName.trim()); setStations([first]); setActive(first.id); setCreating(false); setDraftName("");
  };
  const openProject = (project: StoredProject) => {
    const restored = project.stations.map((station) => ({ ...station, photos: station.photos.map((photo) => normalizePhoto(photo)) }));
    skipRestoredProjectSync.current = true; setProjectId(project.id); setSurveyName(project.name); setStations(restored); setActive(project.activeStationId || restored[0]?.id || ""); setUploaded(null);
  };
  const openCloudProject = async (project: CloudProjectSummary) => {
    if (!account) return;
    setCloudStatus("Opening from OneDrive…");
    try {
      const loaded = await loadOneDriveProject(account, project);
      const restored = loaded.stations.map((station) => ({ ...station, photos: station.photos.map((photo) => normalizePhoto(photo)) }));
      skipRestoredProjectSync.current = true; setProjectId(loaded.id); setSurveyName(loaded.name); setStations(restored); setActive(loaded.activeStationId || restored[0]?.id || ""); setCloudStatus("Saved to OneDrive");
    } catch (error) { setCloudStatus(error instanceof Error ? error.message : "The project could not be opened."); }
  };
  const removeProject = async (project: StoredProject) => { await deleteProject(project.id); setSavedProjects(await listProjects()); };
  const saveNow = async () => {
    if (!projectId || !surveyName) return;
    setLocalStatus("Saving on this device…");
    const saved = { id: projectId, name: surveyName, activeStationId: active, stations: stations.map((station) => ({ ...station, photos: station.photos.map(({ url: _url, ...photo }) => photo) })), updatedAt: Date.now() };
    await saveProject(saved); setSavedProjects(await listProjects()); setLocalStatus(`Saved at ${new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`);
    if (account) {
      setCloudStatus("Syncing to OneDrive…");
      try { await syncProjectToOneDrive(account, saved); setCloudStatus("Saved to OneDrive"); }
      catch { setCloudStatus("Saved on device; OneDrive sync paused"); }
    }
  };
  const nextStation = () => {
    const station = { id: id(), name: "", notes: "", photos: [] };
    setStations([...stations, station]); setActive(station.id);
  };
  const addPhotos = (files: FileList | null) => {
    if (!files || !active) return;
    const photos = Array.from(files).filter((file) => file.type.startsWith("image/")).map((file) => ({ id: id(), name: file.name, url: "", file, originalName: file.name, originalFile: file }));
    setStations(stations.map((station) => station.id === active ? { ...station, photos: [...station.photos, ...photos] } : station));
  };
  const setNotes = (notes: string) => setStations(stations.map((station) => station.id === active ? { ...station, notes } : station));
  const setStationName = (name: string) => setStations(stations.map((station) => station.id === active ? { ...station, name } : station));
  const saveAnnotatedPhoto = (file: File, url: string) => {
    if (!editingPhoto) return;
    URL.revokeObjectURL(editingPhoto.url);
    URL.revokeObjectURL(url);
    setStations(stations.map((station) => ({ ...station, photos: station.photos.map((photo) => photo.id === editingPhoto.id ? { ...photo, name: file.name, file, url: "" } : photo) })));
    setEditingPhoto(null);
  };
  const editPhoto = (photo: Photo, stationName: string, photoIndex: number) => setEditingPhoto({ ...photo, name: numberedPhotoName(stationName, photoIndex, false), url: URL.createObjectURL(photo.file) });
  const closeAnnotator = () => { if (editingPhoto?.url) URL.revokeObjectURL(editingPhoto.url); setEditingPhoto(null); };

  if (!surveyName) return <div className="start">
    <header className="start-header"><Brand/><button className="connect" onClick={connectMicrosoft}><Cloud size={17}/>{connected ? account?.username : "Sign in with Microsoft"}</button></header>
    <main><div className="eyebrow">SITE SURVEY TOOL</div><h1>Document the site<br/><em>while you’re there.</em></h1><p className="lead">Capture photos by station, add field notes, then send a clean, organized survey to OneDrive.</p>
      {creating ? <form className="create-box" onSubmit={(event) => { event.preventDefault(); createSurvey(); }}><label htmlFor="site">What site are you surveying?</label><input autoFocus id="site" value={draftName} onChange={(event) => setDraftName(event.target.value)} placeholder="e.g. Northside Warehouse"/><div><button type="button" onClick={() => setCreating(false)}>Cancel</button><button className="primary" type="submit">Create survey <Plus size={17}/></button></div></form> : <button className="primary large" onClick={() => setCreating(true)}><Plus size={18}/>Start a site survey</button>}
      {savedProjects.length > 0 && <section className="saved-projects"><div><b>Saved projects</b><span>Reopen a survey and keep working.</span></div>{savedProjects.map((project) => <article key={project.id}><button className="saved-open" onClick={() => openProject(project)}><FolderOpen size={20}/><span><b>{project.name}</b><small>{project.stations.length} stations · {project.stations.reduce((sum, station) => sum + station.photos.length, 0)} photos · {new Date(project.updatedAt).toLocaleDateString()}</small></span></button><button className="saved-delete" aria-label={`Delete ${project.name}`} onClick={() => removeProject(project)}><Trash2 size={17}/></button></article>)}</section>}
      {connected && cloudProjects.some((project) => !savedProjects.some((local) => local.id === project.id)) && <section className="saved-projects cloud-projects"><div><b>From OneDrive</b><span>Projects saved on your other devices.</span></div>{cloudProjects.filter((project) => !savedProjects.some((local) => local.id === project.id)).map((project) => <article key={project.id}><button className="saved-open" onClick={() => openCloudProject(project)}><Cloud size={20}/><span><b>{project.name}</b><small>{project.stationCount} stations · {project.photoCount} photos · {new Date(project.updatedAt).toLocaleDateString()}</small></span></button></article>)}</section>}
      {cloudStatus && !surveyName && <p className="cloud-home-status">{cloudStatus}</p>}
      <div className="promise"><span><Check size={17}/>Works on your phone</span><span><Check size={17}/>Original photos preserved</span><span><Check size={17}/>Ready for OneDrive</span></div>
    </main><footer><span>Built for the walk-through.</span><span>iPhone · Android · Desktop</span></footer>
  </div>;

  return <div className="app-shell">
    <aside className="rail"><Brand/><button className="new-survey" onClick={() => { setProjectId(""); setSurveyName(""); setStations([]); }}><Plus size={18}/>New survey</button><div className="nav-title">WALK-THROUGH</div><button className="nav-item selected"><FolderOpen size={18}/>{surveyName}</button><div className="station-list">{stations.map((station, index) => <button key={station.id} className={station.id === active ? "station active" : "station"} onClick={() => setActive(station.id)}><span>{String(index + 1).padStart(2, "0")}</span>{station.name || "Name this station"}<small>{station.photos.length}</small></button>)}</div><div className="onedrive"><Cloud size={20}/><div><b>{connected ? "Connected to OneDrive" : "Offline workspace"}</b><p>{connected ? "Ready to upload" : "Connect when you’re ready"}</p></div></div></aside>
    <main className="workspace"><header><div><button className="back" onClick={() => setSurveyName("")}><ChevronLeft size={18}/>Surveys</button><h1>{surveyName}</h1><small className="local-save-status">{localStatus}</small></div><div className="header-actions"><span className={connected ? "sync good" : "sync"}><i/>{connected ? cloudStatus || "Saved to OneDrive" : localStatus}</span><button className="save-now" onClick={saveNow}><Check size={17}/>Save now</button><button className="outline" onClick={connectMicrosoft}>{connected ? <Check size={17}/> : <Cloud size={17}/>}{connected ? "Connected" : "Connect OneDrive"}</button><button className="primary" onClick={() => setCompleted(true)}><Share2 size={17}/>Complete & share</button></div></header>
      <section className="station-area"><label className="mobile-station-picker">View station<select value={active} onChange={(event) => setActive(event.target.value)}>{stations.map((station, index) => <option key={station.id} value={station.id}>{index + 1}. {station.name || "Unnamed station"} ({station.photos.length} photos)</option>)}</select></label><div className="station-title"><div><div className="eyebrow">STATION {String(stations.findIndex((station) => station.id === active) + 1).padStart(2, "0")}</div><label className="station-name-label" htmlFor="station-name">Where are you?</label><input key={active} autoFocus id="station-name" className="station-name" value={current?.name ?? ""} onChange={(event) => setStationName(event.target.value)} placeholder="Type the station name"/></div><span>{current?.photos.length ?? 0} photos</span></div>
      <div className="capture-card"><div><Camera size={26}/><div><h3>Capture the condition</h3><p>Use your camera or add photos from this device. Originals stay with this station.</p></div></div><button className="primary" onClick={() => photoInput.current?.click()}><Camera size={18}/>Add photos</button><input ref={photoInput} type="file" accept="image/*" capture="environment" multiple onChange={(event) => addPhotos(event.target.files)}/></div>
      <div className="photo-grid">{current?.photos.map((photo, photoIndex) => { const displayName = numberedPhotoName(current.name, photoIndex, photo.file !== photo.originalFile); return <article className="photo" key={photo.id}><button className="photo-preview" onClick={() => editPhoto(photo, current.name, photoIndex)}><PhotoImage file={photo.file} alt={displayName}/><span><PenLine size={15}/>Annotate</span></button><div><span>{displayName}</span><button aria-label={`Annotate ${displayName}`} title="Add measurements" onClick={() => editPhoto(photo, current.name, photoIndex)}><PenLine size={16}/></button></div></article>; })}{!current?.photos.length && <div className="photo-empty"><ImagePlus size={27}/><span>Photos added here will be organized in<br/><b>{surveyName} / {current?.name}</b></span></div>}</div>
      <div className="notes"><label htmlFor="notes">Field notes <small>optional</small></label><textarea id="notes" value={current?.notes ?? ""} onChange={(event) => setNotes(event.target.value)} placeholder="Add observations, measurements, or follow-up items…"/><span>Notes are saved as you type.</span></div><button className="next-station" onClick={nextStation}><span><Check size={18}/></span><div><b>Next station</b><small>Save this station and move on</small></div><Plus size={20}/></button></section>
    </main>
    <aside className="station-tools"><h3>Walk-through progress</h3><p>Name the station, add every photo you need, then use Next station.</p><div className="progress-number"><b>{stations.length}</b><span>stations visited</span></div><div className="tool-note"><Upload size={18}/><span><b>Upload queue</b><br/>{totalPhotos ? `${totalPhotos} photo${totalPhotos === 1 ? "" : "s"} ready to upload` : "No photos yet"}</span></div></aside>
    {editingPhoto && <PhotoAnnotator photoUrl={editingPhoto.url} photoName={editingPhoto.name} onCancel={closeAnnotator} onSave={saveAnnotatedPhoto}/>}
    {completed && <div className="completion-backdrop" role="dialog" aria-modal="true" aria-labelledby="complete-title"><section className="completion"><button className="completion-close" onClick={() => setCompleted(false)}>×</button><div className="complete-check"><Check size={25}/></div><div className="eyebrow">SURVEY READY</div><h2 id="complete-title">Your OneDrive folders</h2><p>Completing this survey creates this exact structure. Every station keeps its own pictures and field notes together.</p><div className="folder-tree"><b><FolderOpen size={18}/>{surveyName}</b>{stations.map((station, index) => <div key={station.id}><strong><FolderOpen size={17}/>{station.name.trim() || `Station ${index + 1}`}</strong>{station.photos.length > 0 && <span><FolderOpen size={15}/>Original Photos ({station.photos.length})</span>}{station.photos.some((photo) => photo.file !== photo.originalFile) && <span><FolderOpen size={15}/>Annotated Photos ({station.photos.filter((photo) => photo.file !== photo.originalFile).length})</span>}<span><PenLine size={15}/>Notes.txt</span></div>)}</div>{uploaded ? <><div className={uploaded.warnings.length ? "upload-success upload-partial" : "upload-success"}><Check size={20}/><b>{uploaded.warnings.length ? "Upload finished with skipped photos" : "Survey uploaded"}</b><a href={uploaded.folderUrl} target="_blank" rel="noreferrer">Open folder in OneDrive</a><button onClick={() => navigator.clipboard.writeText(uploaded.shareUrl)}><Share2 size={16}/>Copy team sharing link</button></div>{uploaded.warnings.length > 0 && <div className="upload-warnings" role="alert"><b>{uploaded.warnings.length} photo{uploaded.warnings.length === 1 ? "" : "s"} could not be read or uploaded</b><p>The rest of the survey was uploaded. Keep this project on your phone and try again after freeing storage or reconnecting.</p><ul>{uploaded.warnings.map((warning) => <li key={warning}>{warning}</li>)}</ul><button className="primary" onClick={() => setUploaded(null)}>Retry skipped photos</button></div>}</> : <button className="primary complete-action" disabled={Boolean(uploadStatus)} onClick={uploadSurvey}><Cloud size={18}/>{uploadStatus || (connected ? "Upload survey to OneDrive" : "Sign in with Microsoft to upload")}</button>}{uploadError && <p className="upload-error" role="alert">{uploadError}</p>}<small className="complete-help">{connected ? `Connected as ${account?.username}` : "Microsoft will ask you to approve access to files in your OneDrive."}</small></section></div>}
  </div>;
}
function Brand() { return <a className="brand" href="#"><span><MapPin size={21}/></span>fieldnote<i>.</i></a>; }
function numberedPhotoName(stationName: string, index: number, annotated: boolean) { return `${stationName.trim() || "Station"} ${String(index + 1).padStart(2, "0")}${annotated ? " - Annotated" : ""}.jpg`; }
function normalizePhoto(photo: any): Photo { const file = photo.file as File; const originalFile = (photo.originalFile as File | undefined) ?? file; return { ...photo, file, originalFile, originalName: photo.originalName || file.name || photo.name || "Photo.jpg", name: photo.name || file.name || "Photo.jpg", url: "" }; }
function PhotoImage({ file, alt }: { file: File; alt: string }) {
  const [src, setSrc] = useState("");
  useEffect(() => {
    if (!(file instanceof Blob)) return;
    const next = URL.createObjectURL(file);
    setSrc(next);
    return () => URL.revokeObjectURL(next);
  }, [file]);
  return src ? <img src={src} alt={alt}/> : <span className="photo-unavailable">Photo unavailable</span>;
}



