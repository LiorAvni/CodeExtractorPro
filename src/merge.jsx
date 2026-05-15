import React, { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { saveAs } from 'file-saver';
import { PDFDocument } from 'pdf-lib';
import { ArrowDown, ArrowLeft, ArrowUp, FileText, Merge, Moon, Plus, Sun, Trash2, UploadCloud } from 'lucide-react';
import './styles.css';

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
  return <main className="merge-page">
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
        <p>Upload several PDF files, choose their order, and download one merged PDFe.</p>
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
  </main>;
}

createRoot(document.getElementById('root')).render(<MergeApp />);
