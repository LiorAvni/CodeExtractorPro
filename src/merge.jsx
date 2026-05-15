import React, { useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
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
function xmlEscape(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
function buildDocxContentTypes(fileCount) {
  const chunks = Array.from({ length: fileCount }, (_, index) =>
    `<Override PartName="/word/afchunks/file${index + 1}.docx" ContentType="${DOCX_MIME}"/>`
  ).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="docx" ContentType="${DOCX_MIME}"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  ${chunks}
</Types>`;
}
function buildDocxPackageRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}
function buildDocxDocumentRels(fileCount) {
  const rels = Array.from({ length: fileCount }, (_, index) =>
    `<Relationship Id="afchunk${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunks/file${index + 1}.docx"/>`
  ).join('\n  ');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rels}
</Relationships>`;
}
function buildAltChunkDocument(files) {
  const body = files.map((file, index) => {
    const pageBreak = index === 0 ? '' : '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    const name = xmlEscape(file.name);
    return `${pageBreak}
<w:p>
  <w:pPr><w:spacing w:after="120"/></w:pPr>
  <w:r><w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="666666"/></w:rPr><w:t>${name}</w:t></w:r>
</w:p>
<w:altChunk r:id="afchunk${index + 1}"/>`;
  }).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${body}
    <w:sectPr>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
      <w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}
async function mergeDocxFiles(files) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', buildDocxContentTypes(files.length));
  zip.folder('_rels').file('.rels', buildDocxPackageRels());
  zip.folder('word').file('document.xml', buildAltChunkDocument(files));
  zip.folder('word').folder('_rels').file('document.xml.rels', buildDocxDocumentRels(files.length));
  const chunkFolder = zip.folder('word').folder('afchunks');

  for (let i = 0; i < files.length; i++) {
    const arrayBuffer = await files[i].arrayBuffer();
    chunkFolder.file(`file${i + 1}.docx`, arrayBuffer, { binary: true });
    await delayFrame();
  }

  return zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'STORE' });
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
      <button className="delete-row" onClick={() => onDelete(row.id)} title={total <= 2 ? 'Clear file' : 'Delete row'}><Trash2 size={16}/></button>
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
  const deleteRow = id => setRows(prev => prev.length <= 2 ? prev.map(row => row.id === id ? { ...row, file: null } : row) : prev.filter(row => row.id !== id));
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
      setMessage(validation.kind === 'docx' ? `Merged ${readyFiles.length} DOCX files successfully. Open in Microsoft Word to see the original formatting, colors, and images preserved.` : `Merged ${readyFiles.length} PDF files successfully.`);
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
