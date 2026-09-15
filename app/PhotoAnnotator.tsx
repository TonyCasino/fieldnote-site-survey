"use client";

import { useEffect, useRef, useState } from "react";
import { Check, RotateCcw, X } from "lucide-react";
import "./annotator.css";

type Line = { x1: number; y1: number; x2: number; y2: number; label: string };

export default function PhotoAnnotator({ photoUrl, photoName, onCancel, onSave }: { photoUrl: string; photoName: string; onCancel: () => void; onSave: (file: File, url: string) => void }) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const imageRef = useRef<HTMLImageElement | null>(null);
  const startRef = useRef<{ x: number; y: number } | null>(null);
  const draftRef = useRef<Line | null>(null);
  const [lines, setLines] = useState<Line[]>([]);
  const [label, setLabel] = useState("");
  const [drawing, setDrawing] = useState(false);

  const paint = (items = lines, draft = draftRef.current) => {
    const canvas = canvasRef.current, image = imageRef.current;
    if (!canvas || !image) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height); ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
    [...items, ...(draft ? [draft] : [])].forEach((line) => drawMeasurement(ctx, line, canvas.width));
  };

  useEffect(() => {
    const image = new Image();
    image.onload = () => {
      const canvas = canvasRef.current;
      if (!canvas) return;
      const scale = Math.min(1, 2400 / Math.max(image.naturalWidth, image.naturalHeight));
      canvas.width = Math.round(image.naturalWidth * scale); canvas.height = Math.round(image.naturalHeight * scale);
      imageRef.current = image; paint([], null);
    };
    image.src = photoUrl;
  }, [photoUrl]);
  useEffect(() => { paint(lines); }, [lines]);

  const point = (event: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = event.currentTarget, rect = canvas.getBoundingClientRect();
    return { x: (event.clientX - rect.left) * canvas.width / rect.width, y: (event.clientY - rect.top) * canvas.height / rect.height };
  };
  const start = (event: React.PointerEvent<HTMLCanvasElement>) => {
    if (!label.trim()) return;
    event.currentTarget.setPointerCapture(event.pointerId);
    const p = point(event); startRef.current = p;
    draftRef.current = { x1: p.x, y1: p.y, x2: p.x, y2: p.y, label: label.trim() }; setDrawing(true);
  };
  const move = (event: React.PointerEvent<HTMLCanvasElement>) => {
    if (!startRef.current || !draftRef.current) return;
    const p = point(event); draftRef.current = { ...draftRef.current, x2: p.x, y2: p.y }; paint(lines, draftRef.current);
  };
  const finish = (event: React.PointerEvent<HTMLCanvasElement>) => {
    if (!startRef.current || !draftRef.current) return;
    const p = point(event), next = { ...draftRef.current, x2: p.x, y2: p.y };
    if (Math.hypot(next.x2 - next.x1, next.y2 - next.y1) > 12) setLines((value) => [...value, next]);
    draftRef.current = null; startRef.current = null; setDrawing(false); setLabel("");
  };
  const save = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    paint(lines, null);
    canvas.toBlob((blob) => {
      if (!blob) return;
      const base = photoName.replace(/\.[^.]+$/, ""), file = new File([blob], `${base}-annotated.jpg`, { type: "image/jpeg", lastModified: Date.now() });
      onSave(file, URL.createObjectURL(file));
    }, "image/jpeg", .92);
  };

  return <div className="annotator-backdrop" role="dialog" aria-modal="true" aria-label={`Annotate ${photoName}`}><section className="annotator">
    <header><div><b>Add measurements</b><span>Type a measurement, then drag a line across the photo.</span></div><button aria-label="Close annotation editor" onClick={onCancel}><X size={20}/></button></header>
    <div className="annotation-stage"><canvas ref={canvasRef} onPointerDown={start} onPointerMove={move} onPointerUp={finish} onPointerCancel={finish}/></div>
    <div className="annotation-controls"><label><span>Measurement label</span><input autoFocus value={label} onChange={(event) => setLabel(event.target.value)} placeholder='e.g. 42 in or 8 ft 6 in'/></label><p className={label.trim() ? "ready" : ""}>{drawing ? "Keep dragging to place the line" : label.trim() ? "Now drag from one end to the other" : "Enter a measurement to begin"}</p><div><button disabled={!lines.length} onClick={() => setLines((value) => value.slice(0, -1))}><RotateCcw size={17}/>Undo</button><button className="primary" disabled={!lines.length} onClick={save}><Check size={18}/>Save annotated photo</button></div></div>
  </section></div>;
}

function drawMeasurement(ctx: CanvasRenderingContext2D, line: Line, width: number) {
  const scale = Math.max(1, width / 1100), head = 13 * scale, angle = Math.atan2(line.y2 - line.y1, line.x2 - line.x1);
  ctx.save(); ctx.strokeStyle = "#ffcf33"; ctx.fillStyle = "#ffcf33"; ctx.lineWidth = 6 * scale; ctx.lineCap = "round"; ctx.shadowColor = "rgba(0,0,0,.65)"; ctx.shadowBlur = 3 * scale;
  ctx.beginPath(); ctx.moveTo(line.x1, line.y1); ctx.lineTo(line.x2, line.y2); ctx.stroke();
  for (const [x, y, a] of [[line.x1, line.y1, angle], [line.x2, line.y2, angle + Math.PI]] as const) { ctx.beginPath(); ctx.moveTo(x, y); ctx.lineTo(x + Math.cos(a + .55) * head, y + Math.sin(a + .55) * head); ctx.moveTo(x, y); ctx.lineTo(x + Math.cos(a - .55) * head, y + Math.sin(a - .55) * head); ctx.stroke(); }
  const midX = (line.x1 + line.x2) / 2, midY = (line.y1 + line.y2) / 2, fontSize = 28 * scale;
  ctx.font = `700 ${fontSize}px Arial`; ctx.textAlign = "center"; ctx.textBaseline = "middle"; const pad = 11 * scale, metrics = ctx.measureText(line.label);
  ctx.shadowBlur = 0; ctx.fillStyle = "rgba(13,35,40,.88)"; ctx.fillRect(midX - metrics.width / 2 - pad, midY - fontSize / 2 - pad / 2, metrics.width + pad * 2, fontSize + pad); ctx.fillStyle = "#fff"; ctx.fillText(line.label, midX, midY); ctx.restore();
}
