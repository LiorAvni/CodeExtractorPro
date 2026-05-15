import React, { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { Document, Packer, PageBreak, PageOrientation, Paragraph, TextRun, AlignmentType } from 'docx';
import { PDFDocument } from 'pdf-lib';
import { ArrowDown, ArrowLeft, ArrowUp, FileText, Merge, Moon, Plus, Sun, Trash2, UploadCloud } from 'lucide-react';
import './styles.css';

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const PDF_MIME = 'application/pdf';

function getInitialTheme() {
  if (typeof window === 'undefined') return 'light';
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}
function fileRow() { return { id: crypto.randomUUID(), file: null }; }
function delayFrame() { return new Promise(resolve => requestAnimationFrame(resolve)); }
function mergeKind(file) {
  const name = file?.name?.toLowerCase() || '';
  if (name.endsWith('.docx')) return 'docx';
  if (name.endsWith('.pdf')) return 'pdf';
  return '';
}
function safeMergedName(file, kind) {
  return (file?.name || `merged-files.${kind}`).replace(/\.(docx|pdf)$/i, '').replace(/[^a-z0-9._-]+/gi, '_') || 'merged-files';
}
function validateFiles(files) {
  if (files.length < 2) return { ok: false, message: 'Choose at least 2 DOCX files or at least 2 PDF files.' };
  const kinds = files.map(mergeKind);
  if (kinds.some(k => !k)) return { ok: false, message: 'Only .docx and .pdf files are supported.' };
  if (!kinds.every(k => k === kinds[0])) return { ok: false, message: 'Do not mix DOCX and PDF files. Merge DOCX with DOCX, or PDF with PDF.' };
  return { ok: true, kind: kinds[0] };
}
function textRuns(text, options = {}) {
  const value = text || ' ';
  return [new TextRun({ text: value, size: 14, font: 'Consolas', color: '000000', ...options })];
}
function collectParagraphText(paragraphNode) {
  let out = '';
  const walk = node => {
    for (const child of node.childNodes || []) {
      const name = child.localName;
      if (name === 't') out += child.textContent || '';
      else if (name === 'tab') out += '    ';
      else if (name === 'br') out += '\n';
      else walk(child);
    }
  };
  walk(paragraphNode);
  return out;
}
async function extractDocxParagraphs(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const docFile = zip.file('word/document.xml');
  if (!docFile) throw new Error(`${file.name} does not look like a valid DOCX file.`);
  const xmlText = await docFile.async('string');
  const xml = new DOMParser().parseFromString(xmlText, 'application/xml');
  if (xml.getElementsByTagName('parsererror').length) throw new Error(`Could not read ${file.name}.`);
  const body = Array.from(xml.getElementsByTagNameNS('*', 'body'))[0];
  if (!body) throw new Error(`Could not find document body in ${file.name}.`);
  const paragraphs = Array.from(body.getElementsByTagNameNS('*', 'p'));
  return paragraphs.flatMap(p => {
    const text = collectParagraphText(p).replace(/\u00a0/g, ' ');
    const split = text.split(/\r\n|\n|\r/);
    return split.length ? split : [''];
  });
}
async function mergeDocxFiles(files) {
  const children = [];
  for (let i = 0; i < files.length; i++) {
    if (i > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
    const paragraphs = await extractDocxParagraphs(files[i]);
    for (const line of paragraphs) {
      children.push(new Paragraph({ alignment: AlignmentType.LEFT, bidirectional: false, children: textRuns(line) }));
    }
    await delayFrame();
  }
  const doc = new Document({
    sections: [{
      properties: { page: { size: { orientation: PageOrientation.LANDSCAPE }, margin: { top: 360, right: 360, bottom: 360, left: 360 } } },
      children
    }]
  });
  return Packer.toBlob(doc);
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
        <strong>{row.file ? row.file.name : 'Drop a DOCX/PDF file here or click to choose'}</strong>
        <span>{row.file ? `${(row.file.size / 1024 / 1024).toFixed(2)} MB` : 'Max 1 file in this slot'}</span>
      </div>
      <input ref={inputRef} type="file" accept=".docx,.pdf" onChange={e => setFromFiles(e.target.files)} />
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
  const [message, setMessage] = useState('Choose at least 2 DOCX files or at least 2 PDF files. The merge happens fully in your browser.');
  const [busy, setBusy] = useState(false);
  React.useEffect(() => { document.documentElement.dataset.theme = theme; }, [theme]);
  const toggleTheme = () => setTheme(prev => prev === 'dark' ? 'light' : 'dark');
  const setFile = (id, file) => {
    if (!mergeKind(file)) { setMessage('Please choose a .docx or .pdf file.'); return; }
    setRows(prev => prev.map(row => row.id === id ? { ...row, file } : row));
    setMessage('Ready. Use the arrow buttons to choose the merge order. DOCX files cannot be mixed with PDF files.');
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
  const validation = validateFiles(readyFiles);
  const canMerge = validation.ok && !busy;
  const doMerge = async () => {
    if (!validation.ok) { setMessage(validation.message); return; }
    setBusy(true);
    setMessage(`Merging ${validation.kind.toUpperCase()} files...`);
    try {
      const blob = validation.kind === 'pdf' ? await mergePdfFiles(readyFiles) : await mergeDocxFiles(readyFiles);
      const extension = validation.kind === 'pdf' ? 'pdf' : 'docx';
      const mime = validation.kind === 'pdf' ? PDF_MIME : DOCX_MIME;
      saveAs(blob.type ? blob : new Blob([blob], { type: mime }), `${safeMergedName(readyFiles[0], extension)}_merged.${extension}`);
      setMessage(`Merged ${readyFiles.length} ${validation.kind.toUpperCase()} files successfully.`);
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
        <h1>DOCX/PDF Merger</h1>
        <p>Upload several Word documents or several PDF files, choose their order, and download one merged file.</p>
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
