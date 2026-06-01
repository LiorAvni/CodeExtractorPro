import React, { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { saveAs } from 'file-saver';
import { PDFDocument } from 'pdf-lib';
import { ArrowDown, ArrowLeft, ArrowUp, FileText, HelpCircle, Merge, Moon, Plus, Sun, Trash2, UploadCloud, X } from 'lucide-react';
import './styles.css';
import { Analytics } from '@vercel/analytics/react';
import posthog from 'posthog-js';

const POSTHOG_KEY = import.meta.env.VITE_POSTHOG_PROJECT_TOKEN;
const POSTHOG_HOST = import.meta.env.VITE_POSTHOG_HOST || 'https://us.i.posthog.com';

if (POSTHOG_KEY && typeof window !== 'undefined') {
  posthog.init(POSTHOG_KEY, {
    api_host: POSTHOG_HOST,
    defaults: '2025-05-24',
    person_profiles: 'identified_only',
  });
}

const PDF_MIME = 'application/pdf';

function getInitialTheme() {
  if (typeof window === 'undefined') return 'light';
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}
function fileRow() { return { id: crypto.randomUUID(), file: null }; }
function delayFrame() { return new Promise(resolve => requestAnimationFrame(resolve)); }
function mergeKind(file) {
  const name = file?.name?.toLowerCase() || '';
  if (name.endsWith('.pdf')) return 'pdf';
  return '';
}
function safeMergedName(file) {
  return (file?.name || 'merged-files.pdf').replace(/\.pdf$/i, '').replace(/[^a-z0-9._-]+/gi, '_') || 'merged-files';
}
function validateFiles(files) {
  if (files.length < 2) return { ok: false, message: 'Choose at least 2 PDF files.' };
  if (files.some(file => mergeKind(file) !== 'pdf')) {
    return { ok: false, message: 'Only .pdf files are supported.' };
  }
  return { ok: true };
}
async function mergePdfFiles(files) {
  const merged = await PDFDocument.create();
  for (const file of files) {
    const source = await PDFDocument.load(await file.arrayBuffer(), { ignoreEncryption: true });
    const pages = await merged.copyPages(source, source.getPageIndices());
    pages.forEach(page => merged.addPage(page));
    await delayFrame();
  }
  const bytes = await merged.save({ useObjectStreams: true });
  return new Blob([bytes], { type: PDF_MIME });
}


function GuidePanel({ open, onClose }) {
  return <aside className={`guide-panel ${open ? 'open' : ''}`} aria-hidden={!open}>
    <div className="guide-head"><strong>PDF Merger Guide</strong><button onClick={onClose}><X size={16}/></button></div>
    <div className="guide-content">
      <h3>How to use the PDF Merger</h3>
      <p>Use this page to join several PDF files into one PDF, fully inside your browser.</p>
      <ol>
        <li>Drop or pick one PDF in each row.</li>
        <li>Use the arrow buttons to decide which PDF comes first, second, and so on.</li>
        <li>Use “Add another file” for more PDFs.</li>
        <li>Use the trash button to clear a required row or remove an extra row.</li>
        <li>Click “Merge and download” to save the final PDF.</li>
      </ol>
      <div className="guide-he" dir="rtl">
        <h3>מדריך בעברית</h3>
        <p>העמוד הזה מחבר כמה קבצי PDF לקובץ PDF אחד, ישירות בדפדפן.</p>
        <ol>
          <li>גרור או בחר קובץ PDF אחד בכל שורה.</li>
          <li>השתמש בחיצים כדי לבחור את סדר הקבצים.</li>
          <li>לחץ על “Add another file” כדי להוסיף עוד PDF.</li>
          <li>כפתור הפח מנקה שורה חובה או מוחק שורה נוספת.</li>
          <li>לחץ על “Merge and download” כדי להוריד את הקובץ המאוחד.</li>
        </ol>
      </div>
      <span className="owner-watermark" aria-hidden="true">Lior Avni 15/05/2026</span>
    </div>
  </aside>;
}

function UploadSlot({ row, index, total, onFile, onMove, onDelete }) {
  const inputRef = useRef(null);
  const setFromFiles = fileList => {
    const file = Array.from(fileList || [])[0];
    if (!file) return;
    onFile(row.id, file);
    if (inputRef.current) inputRef.current.value = '';
  };
  return <div className="merge-row">
    <div className="merge-index">{index + 1}</div>
    <label className="docx-drop" onDrop={e => { e.preventDefault(); setFromFiles(e.dataTransfer.files); }} onDragOver={e => e.preventDefault()}>
      <UploadCloud size={26}/>
      <div>
        <strong>{row.file ? row.file.name : 'Drop a PDF file here or click to choose'}</strong>
        <span>{row.file ? `${(row.file.size / 1024 / 1024).toFixed(2)} MB` : 'Max 1 file in this slot'}</span>
      </div>
      <input ref={inputRef} type="file" accept=".pdf,application/pdf" onChange={e => setFromFiles(e.target.files)} />
    </label>
    <div className="row-tools">
      <button onClick={() => onMove(index, -1)} disabled={index === 0} title="Move up"><ArrowUp size={16}/></button>
      <button onClick={() => onMove(index, 1)} disabled={index === total - 1} title="Move down"><ArrowDown size={16}/></button>
      <button className="delete-row" onClick={() => onDelete(row.id)} title={total <= 2 ? 'Clear file' : 'Delete row'}><Trash2 size={16}/></button>
    </div>
  </div>;
}

function MergeApp() {
  const [theme, setTheme] = useState(getInitialTheme);
  const [rows, setRows] = useState([fileRow(), fileRow()]);
  const [message, setMessage] = useState('Choose at least 2 PDF files. The merge happens fully in your browser.');
  const [busy, setBusy] = useState(false);
  const [guideOpen, setGuideOpen] = useState(false);
  React.useEffect(() => { document.documentElement.dataset.theme = theme; }, [theme]);
  const toggleTheme = () => setTheme(prev => prev === 'dark' ? 'light' : 'dark');
  const setFile = (id, file) => {
    if (!mergeKind(file)) { setMessage('Please choose a .pdf file.'); return; }
    setRows(prev => prev.map(row => row.id === id ? { ...row, file } : row));
    setMessage('Use the arrow buttons to choose the merge order.');
  };
  const moveRow = (index, direction) => {
    setRows(prev => {
      const nextIndex = index + direction;
      if (nextIndex < 0 || nextIndex >= prev.length) return prev;
      const next = [...prev];
      [next[index], next[nextIndex]] = [next[nextIndex], next[index]];
      return next;
    });
  };
  const deleteRow = id => setRows(prev => prev.length <= 2 ? prev.map(row => row.id === id ? { ...row, file: null } : row) : prev.filter(row => row.id !== id));
  const readyFiles = rows.map(row => row.file).filter(Boolean);
  const validation = validateFiles(readyFiles);
  const canMerge = validation.ok && !busy;
  const doMerge = async () => {
    if (!validation.ok) { setMessage(validation.message); return; }
    setBusy(true);
    setMessage('Merging PDF files...');
    try {
      const blob = await mergePdfFiles(readyFiles);
      saveAs(blob.type ? blob : new Blob([blob], { type: PDF_MIME }), `${safeMergedName(readyFiles[0])}_merged.pdf`);
      setMessage(`Merged ${readyFiles.length} PDF files successfully.`);
    } catch (e) {
      console.error(e);
      setMessage(e?.message || 'Could not merge the selected files. Very complex or encrypted files may not be mergeable in the browser.');
    } finally {
      setBusy(false);
    }
  };
  return <>
  <GuidePanel open={guideOpen} onClose={() => setGuideOpen(false)} />
  <main className={`merge-page ${guideOpen ? 'with-guide' : ''}`}>
    <div className="top-actions">
      <button className="theme-toggle" onClick={toggleTheme} aria-label={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}>
        {theme === 'dark' ? <Sun size={17}/> : <Moon size={17}/>}
        {theme === 'dark' ? 'Light' : 'Dark'}
      </button>
    </div>
    <section className="merge-hero">
      <div>
        <p className="eyebrow">CodeExtractor Pro</p>
        <h1>PDF Merger</h1>
        <div className="subtitle-line"><button className="guide-button" onClick={() => setGuideOpen(prev => !prev)} aria-expanded={guideOpen}><HelpCircle size={16}/>How To Use</button><p>Upload several PDF files, choose their order, and download one merged PDF.</p></div>
      </div>
      <a className="back-link" href="/" target="_blank" rel="noreferrer"><ArrowLeft size={17}/> CodeExtractor</a>
    </section>
    <section className="merge-card">
      <div className="merge-list">
        {rows.map((row, index) => <UploadSlot key={row.id} row={row} index={index} total={rows.length} onFile={setFile} onMove={moveRow} onDelete={deleteRow} />)}
      </div>
      <div className="merge-footer">
        <div className="merge-footer-left">
          <button onClick={() => setRows(prev => [...prev, fileRow()])}><Plus size={17}/> Add another file</button>
          <button className="merge-primary" onClick={doMerge} disabled={!canMerge}><Merge size={17}/>{busy ? 'Merging...' : 'Merge and download'}</button>
        </div>
        <p className="merge-message"><FileText size={15}/> {message}</p>
      </div>
    </section>
  </main>
  </>;
}

createRoot(document.getElementById('merge-root')).render(
  <React.StrictMode>
    <MergeApp />
    <Analytics />
  </React.StrictMode>
);
