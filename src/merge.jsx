import React, { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { ArrowDown, ArrowLeft, ArrowUp, FileText, Merge, Moon, Plus, Sun, Trash2, UploadCloud } from 'lucide-react';
import './styles.css';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function getInitialTheme() {
  if (typeof window === 'undefined') return 'light';
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}
function fileRow() { return { id: crypto.randomUUID(), file: null }; }
function safeDocxName(name) { return (name || 'merged-documents').replace(/\.docx$/i, '').replace(/[^a-z0-9._-]+/gi, '_') || 'merged-documents'; }
function isDocx(file) { return file && file.name.toLowerCase().endsWith('.docx'); }

function splitDocumentXml(xml) {
  const match = xml.match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/);
  if (!match) throw new Error('Could not find document body.');
  const bodyInner = match[1];
  const sectMatch = bodyInner.match(/<w:sectPr[\s\S]*?<\/w:sectPr>\s*$/);
  const sectPr = sectMatch ? sectMatch[0] : '';
  const bodyWithoutSectPr = sectPr ? bodyInner.slice(0, bodyInner.length - sectPr.length) : bodyInner;
  return { bodyInner: bodyWithoutSectPr.trim(), sectPr };
}
function pageBreakParagraph() {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}
async function mergeDocxFiles(files) {
  const loaded = [];
  for (const file of files) {
    const zip = await JSZip.loadAsync(file);
    const docFile = zip.file('word/document.xml');
    if (!docFile) throw new Error(`${file.name} does not look like a valid DOCX file.`);
    const xml = await docFile.async('string');
    loaded.push({ file, zip, xml, ...splitDocumentXml(xml) });
  }

  const base = loaded[0];
  const mergedBody = loaded.map((doc, index) => `${index > 0 ? pageBreakParagraph() : ''}${doc.bodyInner}`).join('');
  const newDocumentXml = base.xml.replace(/<w:body[^>]*>[\s\S]*?<\/w:body>/, `<w:body>${mergedBody}${base.sectPr}</w:body>`);
  base.zip.file('word/document.xml', newDocumentXml);
  return base.zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME });
}

function UploadSlot({ row, index, total, onFile, onMove, onDelete }) {
  const inputRef = useRef(null);
  const setFromFiles = fileList => {
    const file = Array.from(fileList || [])[0];
    if (!file) return;
    onFile(row.id, file);
  };
  return <div className="merge-row">
    <div className="merge-index">{index + 1}</div>
    <label className="docx-drop" onDrop={e => { e.preventDefault(); setFromFiles(e.dataTransfer.files); }} onDragOver={e => e.preventDefault()}>
      <UploadCloud size={26}/>
      <div>
        <strong>{row.file ? row.file.name : 'Drop a DOCX file here or click to choose'}</strong>
        <span>{row.file ? `${(row.file.size / 1024 / 1024).toFixed(2)} MB` : 'Max 1 DOCX file in this slot'}</span>
      </div>
      <input ref={inputRef} type="file" accept=".docx" onChange={e => setFromFiles(e.target.files)} />
    </label>
    <div className="row-tools">
      <button onClick={() => onMove(index, -1)} disabled={index === 0} title="Move up"><ArrowUp size={16}/></button>
      <button onClick={() => onMove(index, 1)} disabled={index === total - 1} title="Move down"><ArrowDown size={16}/></button>
      <button className="delete-row" onClick={() => onDelete(row.id)} disabled={total <= 2} title="Delete row"><Trash2 size={16}/></button>
    </div>
  </div>;
}

function MergeApp() {
  const [theme, setTheme] = useState(getInitialTheme);
  const [rows, setRows] = useState([fileRow(), fileRow()]);
  const [message, setMessage] = useState('Choose at least 2 DOCX files. The merge happens fully in your browser.');
  const [busy, setBusy] = useState(false);
  React.useEffect(() => { document.documentElement.dataset.theme = theme; }, [theme]);
  const toggleTheme = () => setTheme(prev => prev === 'dark' ? 'light' : 'dark');
  const setFile = (id, file) => {
    if (!isDocx(file)) { setMessage('Please choose a .docx file.'); return; }
    setRows(prev => prev.map(row => row.id === id ? { ...row, file } : row));
    setMessage('Ready. Use the arrow buttons to choose the merge order.');
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
  const deleteRow = id => setRows(prev => prev.length <= 2 ? prev : prev.filter(row => row.id !== id));
  const readyFiles = rows.map(row => row.file).filter(Boolean);
  const canMerge = readyFiles.length >= 2 && !busy;
  const doMerge = async () => {
    if (!canMerge) return;
    setBusy(true);
    setMessage('Merging DOCX files...');
    try {
      const blob = await mergeDocxFiles(readyFiles);
      saveAs(blob, `${safeDocxName(readyFiles[0].name)}_merged.docx`);
      setMessage(`Merged ${readyFiles.length} DOCX files successfully.`);
    } catch (e) {
      console.error(e);
      setMessage(e?.message || 'Could not merge the selected DOCX files.');
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
        <h1>DOCX Merger</h1>
        <p>Upload several Word documents, choose their order, and download one merged DOCX file.</p>
      </div>
      <a className="back-link" href="/" target="_blank" rel="noreferrer"><ArrowLeft size={17}/> CodeExtractor</a>
    </section>
    <section className="merge-card">
      <div className="merge-list">
        {rows.map((row, index) => <UploadSlot key={row.id} row={row} index={index} total={rows.length} onFile={setFile} onMove={moveRow} onDelete={deleteRow} />)}
      </div>
      <div className="merge-footer">
        <div className="merge-footer-left">
          <button onClick={() => setRows(prev => [...prev, fileRow()])}><Plus size={17}/> Add another DOCX</button>
          <button className="merge-primary" onClick={doMerge} disabled={!canMerge}><Merge size={17}/>{busy ? 'Merging...' : 'Merge and download'}</button>
        </div>
        <p className="merge-message"><FileText size={15}/> {message}</p>
      </div>
    </section>
  </main>;
}

createRoot(document.getElementById('root')).render(<MergeApp />);
