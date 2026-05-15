import React, { useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import hljs from 'highlight.js/lib/core';
import csharp from 'highlight.js/lib/languages/csharp';
import javascript from 'highlight.js/lib/languages/javascript';
import css from 'highlight.js/lib/languages/css';
import xml from 'highlight.js/lib/languages/xml';
import cpp from 'highlight.js/lib/languages/cpp';
import makefile from 'highlight.js/lib/languages/makefile';
import json from 'highlight.js/lib/languages/json';
import bash from 'highlight.js/lib/languages/bash';
import plaintext from 'highlight.js/lib/languages/plaintext';
import { Document, Packer, Paragraph, TextRun, PageOrientation, AlignmentType } from 'docx';
import { Archive, ChevronDown, ChevronRight, Copy, Download, FileCode2, FileText, Folder, FolderOpen, Moon, Sun, Trash2, UploadCloud, X } from 'lucide-react';
import './styles.css';

hljs.registerLanguage('csharp', csharp);
hljs.registerLanguage('javascript', javascript);
hljs.registerLanguage('css', css);
hljs.registerLanguage('xml', xml);
hljs.registerLanguage('cpp', cpp);
hljs.registerLanguage('makefile', makefile);
hljs.registerLanguage('json', json);
hljs.registerLanguage('bash', bash);
hljs.registerLanguage('plaintext', plaintext);

const OUTPUT_EXTENSIONS = new Set([
  '.cs', '.xaml', '.c', '.h', '.cpp', '.hpp', '.cc', '.cxx', '.m', '.mm',
  '.as', '.css', '.js', '.jsx', '.ts', '.tsx', '.html', '.htm', '.config',
  '.razor', '.cshtml', '.json', '.xml', '.sql', '.md', '.txt', '.yml', '.yaml',
  '.scss', '.sass', '.less', '.php', '.java', '.py', '.rb', '.go', '.rs', '.swift',
  '.kt', '.kts', '.vb', '.fs', '.fsx', '.sh', '.bat', '.ps1', '.dockerfile', '.env'
]);

const CONTEXT_EXTENSIONS = new Set(['.csproj', '.vbproj', '.fsproj', '.sln']);
const SPECIAL_FILENAMES = new Set(['makefile', 'dockerfile', 'app.config', 'web.config', '.editorconfig']);
const IGNORE_PARTS = new Set(['node_modules', '.git', '.vs', 'bin', 'obj', 'dist', 'build', '.next', '.vercel', 'coverage', '.idea', '.vscode']);
const MAX_DOCX_PARAGRAPHS = 24000;

const TYPE_COLORS = {
  keyword: '0000FF', built_in: '2B91AF', type: '2B91AF', literal: '0000FF', number: '098658',
  string: 'A31515', regexp: 'A31515', symbol: 'A31515', comment: '008000', quote: '008000',
  doctag: '808080', meta: '808080', title: '795E26', section: '795E26', name: '800000',
  attr: 'FF0000', attribute: 'FF0000', variable: '001080', params: '001080', subst: '001080',
  tag: '800000', selector_tag: '800000', selector_id: '800000', selector_class: '800000',
  property: 'FF0000', function: '795E26', class: '2B91AF', default: '000000'
};

function normalizePath(path) { return path.replace(/\\/g, '/').replace(/^\/+/, ''); }
function getFileName(path) { return normalizePath(path).split('/').filter(Boolean).pop() || ''; }
function getExt(path) {
  const name = getFileName(path).toLowerCase();
  if (name === 'makefile' || name.endsWith('.mk')) return '.makefile';
  if (name === 'dockerfile') return '.dockerfile';
  const dot = name.lastIndexOf('.');
  return dot >= 0 ? name.slice(dot) : '';
}
function isIgnoredPath(path) { return normalizePath(path).split('/').some(part => IGNORE_PARTS.has(part)); }
function shouldOutput(path) {
  const name = getFileName(path).toLowerCase();
  const ext = getExt(path);
  return SPECIAL_FILENAMES.has(name) || OUTPUT_EXTENSIONS.has(ext) || ext === '.makefile' || ext === '.dockerfile';
}
function shouldReadAsContext(path) { return CONTEXT_EXTENSIONS.has(getExt(path)); }
function depthOf(path) { return normalizePath(path).split('/').filter(Boolean).length; }
function safeBaseName(name) { return name.replace(/\.zip$/i, '').replace(/[^a-z0-9._-]+/gi, '_') || 'code-extract'; }
function indent(level) { return '  '.repeat(level); }
function stripBom(text) { return text.replace(/^\uFEFF/, ''); }
function detectLanguage(path) {
  const ext = getExt(path);
  const name = getFileName(path).toLowerCase();
  if (['.cs', '.razor', '.cshtml'].includes(ext)) return 'csharp';
  if (['.xaml', '.xml', '.config', '.html', '.htm', '.svg'].includes(ext)) return 'xml';
  if (['.js', '.jsx', '.ts', '.tsx'].includes(ext)) return 'javascript';
  if (['.css', '.scss', '.sass', '.less'].includes(ext)) return 'css';
  if (['.c', '.h', '.cpp', '.hpp', '.cc', '.cxx'].includes(ext)) return 'cpp';
  if (ext === '.json') return 'json';
  if (ext === '.makefile' || name.endsWith('.mk')) return 'makefile';
  if (['.sh', '.bat', '.ps1'].includes(ext)) return 'bash';
  return 'plaintext';
}
function xmlTag(text, tag) {
  const match = text.match(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, 'i'));
  return match ? match[1].trim() : '';
}
function detectProjectType(path, content) {
  const ext = getExt(path);
  const lower = content.toLowerCase();
  if (ext === '.sln') return 'solution';
  const outputType = xmlTag(content, 'OutputType').toLowerCase();
  const targetFramework = xmlTag(content, 'TargetFramework') || xmlTag(content, 'TargetFrameworks') || xmlTag(content, 'TargetFrameworkVersion');
  const useWpf = xmlTag(content, 'UseWPF').toLowerCase() === 'true' || lower.includes('<usewpf>true</usewpf>');
  const useMaui = xmlTag(content, 'UseMaui').toLowerCase() === 'true' || lower.includes('<usemaui>true</usemaui>');
  const useWinForms = xmlTag(content, 'UseWindowsForms').toLowerCase() === 'true';
  const isBlazor = lower.includes('microsoft.net.sdk.blazorwebassembly') || lower.includes('microsoft.net.sdk.web') || lower.includes('<razorcompileonbuild>') || lower.includes('blazor');
  const isWcf = lower.includes('system.servicemodel') || lower.includes('wcflibrary') || lower.includes('wcf service');
  const isFramework = /net4\d|v4\./i.test(targetFramework) || lower.includes('<targetframeworkversion>v4.');
  if (useMaui) return '.NET MAUI app';
  if (isBlazor && outputType === 'exe') return 'Blazor Web App';
  if (isBlazor) return 'Blazor project';
  if (useWpf && isFramework) return 'WPF App (.NET Framework)';
  if (useWpf) return 'WPF App';
  if (useWinForms) return isFramework ? 'Windows Forms App (.NET Framework)' : 'Windows Forms App';
  if (isWcf) return 'WCF Service Library';
  if (outputType === 'exe' || outputType === 'winexe') return isFramework ? '.NET Framework app' : '.NET app';
  if (ext === '.csproj' || ext === '.vbproj' || ext === '.fsproj') return isFramework ? 'Class Library (.NET Framework)' : '.NET class library';
  return 'project';
}
function decodeHtmlEntities(s) {
  const map = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&#x27;': "'", '&#39;': "'" };
  return s.replace(/&(amp|lt|gt|quot|#x27|#39);/g, m => map[m] || m);
}
function highlightedTokens(code, language) {
  let html;
  try { html = hljs.highlight(code, { language, ignoreIllegals: true }).value; }
  catch { html = hljs.highlight(code, { language: 'plaintext', ignoreIllegals: true }).value; }
  const tokens = [];
  const regex = /<span class="hljs-([^"]+)">([\s\S]*?)<\/span>|([^<]+)/g;
  let match;
  while ((match = regex.exec(html))) {
    if (match[2] != null) {
      const cls = match[1].split(' ')[0].replace(/-/g, '_');
      tokens.push({ text: decodeHtmlEntities(match[2].replace(/<[^>]+>/g, '')), color: TYPE_COLORS[cls] || TYPE_COLORS.default });
    } else if (match[3] != null) {
      tokens.push({ text: decodeHtmlEntities(match[3]), color: TYPE_COLORS.default });
    }
  }
  return tokens.length ? tokens : [{ text: code, color: TYPE_COLORS.default }];
}
function tokensToLines(tokens) {
  const lines = [[]];
  for (const token of tokens) {
    const parts = token.text.split(/(\r\n|\n|\r)/);
    for (const part of parts) {
      if (part === '\n' || part === '\r' || part === '\r\n') lines.push([]);
      else if (part) lines[lines.length - 1].push({ text: part, color: token.color });
    }
  }
  return lines;
}
function makeDocxRunsFromLine(lineTokens) {
  if (!lineTokens.length) return [new TextRun({ text: ' ', size: 14, font: 'Consolas' })];
  return lineTokens.map(t => new TextRun({ text: t.text.replace(/\t/g, '    '), color: t.color, size: 14, font: 'Consolas' }));
}
function newNode(name, type) {
  return { name, type, children: {}, sortedChildren: [], file: null, label: '', id: crypto.randomUUID() };
}
function insertTreePath(root, filePath, fileObj) {
  const parts = normalizePath(filePath).split('/').filter(Boolean);
  let node = root;
  parts.forEach((part, index) => {
    const isFile = index === parts.length - 1;
    if (!node.children[part]) node.children[part] = newNode(part, isFile ? 'file' : 'folder');
    if (isFile) node.children[part].file = fileObj;
    node = node.children[part];
  });
}
function sortTree(node) {
  const sorted = Object.values(node.children || {}).sort((a, b) => {
    const aFolderish = a.type === 'folder';
    const bFolderish = b.type === 'folder';
    if (aFolderish !== bFolderish) return aFolderish ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
  node.sortedChildren = sorted;
  sorted.forEach(sortTree);
  return node;
}
function applyProjectLabels(tree, projectMap) {
  function walk(node, currentPath = '') {
    const full = currentPath ? `${currentPath}/${node.name}` : node.name;
    if (node.type === 'folder') {
      const directProject = Object.entries(projectMap).find(([p]) => normalizePath(p).startsWith(`${full}/`) && depthOf(normalizePath(p)) === depthOf(full) + 1);
      if (directProject) node.label = directProject[1];
      Object.values(node.children || {}).forEach(child => walk(child, full));
    }
  }
  Object.values(tree.children).forEach(child => walk(child, ''));
}
function collectFilePaths(node) {
  const paths = [];
  if ((node.type === 'file') && node.file) paths.push(node.file.path);
  node.sortedChildren?.forEach(child => paths.push(...collectFilePaths(child)));
  return paths;
}
function getSelectionState(node, selectedPaths) {
  const paths = collectFilePaths(node);
  if (!paths.length) return 'none';
  const selectedCount = paths.filter(path => selectedPaths.has(path)).length;
  if (selectedCount === 0) return 'none';
  if (selectedCount === paths.length) return 'all';
  return 'partial';
}
function buildSelectedOutput(tree, rootName, selectedPaths) {
  const lines = [`${rootName} (${tree.solutionLabel || 'zip project'}):`];
  const filesForDocx = [];
  function walk(node, level, path) {
    const hasSelectedDescendant = collectFilePaths(node).some(p => selectedPaths.has(p));
    if (!hasSelectedDescendant) return;
    const isFolderish = node.type === 'folder';
    const suffix = isFolderish ? ` (${node.label || ('folder')}):` : ':';
    lines.push(`${indent(level)}${node.name}${suffix}`);
    if ((node.type === 'file') && node.file && selectedPaths.has(node.file.path)) {
      lines.push(`${indent(level + 1)}~~~`);
      lines.push(node.file.content.replace(/\t/g, '    '));
      lines.push(`${indent(level + 1)}~~~`);
      filesForDocx.push({ ...node.file, displayPath: path });
    }
    node.sortedChildren?.forEach(child => walk(child, level + 1, `${path}/${child.name}`));
  }
  tree.sortedChildren.forEach(child => walk(child, 1, child.name));
  return { text: lines.join('\n'), filesForDocx };
}
function countSelected(result) {
  return result.allFilePaths.filter(path => result.selectedPaths.has(path)).length;
}
function downloadBlob(text, filename, type = 'text/plain;charset=utf-8') { saveAs(new Blob([text], { type }), filename); }
async function copyText(text) { await navigator.clipboard.writeText(text); }
async function downloadDocx(result, output) {
  const paragraphs = [];
  const para = (children) => new Paragraph({ alignment: AlignmentType.LEFT, bidirectional: false, children });
  paragraphs.push(para([new TextRun({ text: `${result.rootName} (${result.tree.solutionLabel || 'zip project'}):`, bold: true, size: 14, font: 'Consolas' })]));
  let count = 1;
  function addPlain(line, bold = false) {
    if (count++ > MAX_DOCX_PARAGRAPHS) return;
    paragraphs.push(para([new TextRun({ text: line || ' ', bold, size: 14, font: 'Consolas', color: '000000' })]));
  }
  function walk(node, level, path) {
    if (count > MAX_DOCX_PARAGRAPHS) return;
    const hasSelectedDescendant = collectFilePaths(node).some(p => result.selectedPaths.has(p));
    if (!hasSelectedDescendant) return;
    const isFolderish = node.type === 'folder';
    addPlain(`${indent(level)}${node.name}${isFolderish ? ` (${node.label || ('folder')}):` : ':'}`, isFolderish);
    if ((node.type === 'file') && node.file && result.selectedPaths.has(node.file.path)) {
      addPlain(`${indent(level + 1)}~~~`);
      const prefix = indent(level + 1);
      const codeLines = tokensToLines(highlightedTokens(node.file.content, node.file.language));
      for (const line of codeLines) {
        if (count++ > MAX_DOCX_PARAGRAPHS) break;
        paragraphs.push(para([new TextRun({ text: prefix, size: 14, font: 'Consolas' }), ...makeDocxRunsFromLine(line)]));
      }
      addPlain(`${indent(level + 1)}~~~`);
    }
    node.sortedChildren?.forEach(child => walk(child, level + 1, `${path}/${child.name}`));
  }
  result.tree.sortedChildren.forEach(child => walk(child, 1, child.name));
  if (count > MAX_DOCX_PARAGRAPHS) addPlain('[DOCX truncated to keep Word stable. Use TXT for the full extraction.]');
  const doc = new Document({
    sections: [{
      properties: { page: { size: { orientation: PageOrientation.LANDSCAPE }, margin: { top: 360, right: 360, bottom: 360, left: 360 } } },
      children: paragraphs
    }]
  });
  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${safeBaseName(result.name)}.docx`);
}
async function parseZip(file) {
  const zip = await JSZip.loadAsync(file);
  const rootName = safeBaseName(file.name);
  const tree = newNode(rootName, 'folder');
  tree.solutionLabel = 'zip project';
  const projectMap = {};
  const outputFiles = [];
  const entries = Object.values(zip.files).filter(e => !e.dir && !isIgnoredPath(e.name));

  for (const entry of entries) {
    const path = normalizePath(entry.name);
    if (shouldReadAsContext(path)) {
      const content = stripBom(await entry.async('string'));
      const type = detectProjectType(path, content);
      if (getExt(path) === '.sln') tree.solutionLabel = 'solution';
      else projectMap[path] = type;
    }
  }
  for (const entry of entries) {
    const path = normalizePath(entry.name);
    if (!shouldOutput(path) || shouldReadAsContext(path)) continue;
    const content = stripBom(await entry.async('string'));
    const fileObj = { path, name: getFileName(path), content, language: detectLanguage(path), ext: getExt(path) };
    outputFiles.push(fileObj);
    insertTreePath(tree, path, fileObj);
  }
  applyProjectLabels(tree, projectMap);
  sortTree(tree);
  const allFilePaths = outputFiles.map(f => f.path);
  const selectedPaths = new Set(allFilePaths);
  const output = buildSelectedOutput(tree, rootName, selectedPaths);
  return { id: crypto.randomUUID(), name: file.name, rootName, tree, fileCount: outputFiles.length, allFilePaths, selectedPaths, outputText: output.text, filesForDocx: output.filesForDocx, createdAt: new Date().toLocaleString() };
}
function TriStateBox({ state, onClick }) {
  return <button
    className={`checkbox ${state}`}
    onClick={e => { e.stopPropagation(); onClick(); }}
    aria-label="Toggle selection"
    title="Toggle selection"
  >{state === 'all' ? '✓' : state === 'partial' ? '—' : ''}</button>;
}
function TreeNode({ node, depth = 0, selectedPaths, setSelectedPaths }) {
  const [open, setOpen] = useState(true);
  const isExpandable = (node.type === 'folder') && node.sortedChildren?.length > 0;
  const isFolderish = node.type === 'folder';
  const selectionState = getSelectionState(node, selectedPaths);
  const toggleSelected = () => {
    const paths = collectFilePaths(node);
    setSelectedPaths(prev => {
      const next = new Set(prev);
      const shouldSelect = getSelectionState(node, prev) !== 'all';
      paths.forEach(path => shouldSelect ? next.add(path) : next.delete(path));
      return next;
    });
  };
  return <div>
    <div className="tree-row" style={{ paddingLeft: 8 + depth * 16 }} title={node.name} onClick={() => isExpandable && setOpen(!open)}>
      <span className="chevron">{isExpandable ? (open ? <ChevronDown size={14}/> : <ChevronRight size={14}/>) : <span className="chevron-spacer" />}</span>
      <TriStateBox state={selectionState} onClick={toggleSelected} />
      <span className="node-icon">{isFolderish ? (open ? <FolderOpen size={15}/> : <Folder size={15}/>) : <FileCode2 size={15}/>}</span>
      <span className="node-name">{node.name}</span>{node.label && <em>{node.label}</em>}
    </div>
    {isExpandable && open && node.sortedChildren?.map(child => <TreeNode key={child.name} node={child} depth={depth + 1} selectedPaths={selectedPaths} setSelectedPaths={setSelectedPaths} />)}
  </div>;
}
function ResultBlock({ result, onDelete, onSelectionChange }) {
  const [leftWidth, setLeftWidth] = useState(320);
  const [copied, setCopied] = useState(false);
  const dragging = useRef(false);
  const selectedCount = countSelected(result);
  const selectedOutput = useMemo(() => buildSelectedOutput(result.tree, result.rootName, result.selectedPaths), [result.tree, result.rootName, result.selectedPaths]);
  const setSelectedPaths = updater => onSelectionChange(result.id, updater);
  const startDrag = e => { dragging.current = true; e.preventDefault(); };
  React.useEffect(() => {
    const move = e => { if (dragging.current) setLeftWidth(Math.min(620, Math.max(210, e.clientX - 28))); };
    const up = () => { dragging.current = false; };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
    return () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
  }, []);
  const selectAll = () => onSelectionChange(result.id, () => new Set(result.allFilePaths));
  return <section className="result-card">
    <header className="result-header">
      <div className="title-wrap"><Archive size={20}/><div><h2>{result.name}</h2><p>{selectedCount} out of {result.fileCount} files selected • {result.createdAt}</p></div></div>
      <div className="actions">
        <button onClick={async () => { await copyText(selectedOutput.text); setCopied(true); setTimeout(() => setCopied(false), 1400); }}><Copy size={16}/>{copied ? 'Copied' : 'Copy text'}</button>
        <button onClick={() => downloadBlob(selectedOutput.text, `${safeBaseName(result.name)}.txt`)}><FileText size={16}/>TXT</button>
        <button onClick={() => downloadDocx(result, selectedOutput)}><Download size={16}/>DOCX</button>
        <button className="danger" onClick={() => onDelete(result.id)}><Trash2 size={16}/>Delete</button>
      </div>
    </header>
    <div className="workspace">
      <aside className="explorer" style={{ width: leftWidth }}>
        <div className="explorer-title"><span>File Explorer</span><button onClick={selectAll}>Select All</button></div>
        <TreeNode node={result.tree} selectedPaths={result.selectedPaths} setSelectedPaths={setSelectedPaths} />
      </aside>
      <div className="resizer" onMouseDown={startDrag} title="Drag to resize" />
      <pre className="output"><code>{selectedOutput.text}</code></pre>
    </div>
  </section>;
}
function getInitialTheme() {
  if (typeof window === 'undefined') return 'light';
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function App() {
  const [results, setResults] = useState([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [theme, setTheme] = useState(getInitialTheme);
  const inputRef = useRef(null);
  const totalFiles = useMemo(() => results.reduce((sum, r) => sum + r.fileCount, 0), [results]);
  React.useEffect(() => {
    document.documentElement.dataset.theme = theme;
  }, [theme]);
  async function handleFiles(fileList) {
    const zips = Array.from(fileList || []).filter(f => f.name.toLowerCase().endsWith('.zip'));
    if (!zips.length) { setError('Please choose one or more .zip files.'); return; }
    setBusy(true); setError('');
    try {
      const parsed = [];
      for (const zip of zips) parsed.push(await parseZip(zip));
      setResults(prev => [...parsed, ...prev]);
    } catch (e) {
      console.error(e);
      setError('Could not parse one of the ZIP files. Make sure it is a valid project ZIP and not password-protected.');
    } finally { setBusy(false); }
  }
  const updateSelection = (id, updater) => {
    setResults(prev => prev.map(item => {
      if (item.id !== id) return item;
      const selectedPaths = typeof updater === 'function' ? updater(item.selectedPaths) : updater;
      return { ...item, selectedPaths };
    }));
  };
  const toggleTheme = () => setTheme(prev => prev === 'dark' ? 'light' : 'dark');
  return <main>
    <button className="theme-toggle" onClick={toggleTheme} aria-label={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}>
      {theme === 'dark' ? <Sun size={17}/> : <Moon size={17}/>}
      {theme === 'dark' ? 'Light' : 'Dark'}
    </button>
    <section className="hero">
      <div><p className="eyebrow">CodeExtractor Pro</p><h1>Turn code project ZIPs into clean text.</h1><p className="subtitle">Multiple ZIPs, solution-style structure, selectable files, TXT export, and DOCX export with syntax-colored code at 7pt.</p></div>
      <div className="stats"><strong>{results.length}</strong><span>ZIPs loaded</span><strong>{totalFiles}</strong><span>files extracted</span></div>
    </section>
    <section className="dropzone" onDrop={e => { e.preventDefault(); handleFiles(e.dataTransfer.files); }} onDragOver={e => e.preventDefault()} onClick={() => inputRef.current?.click()}>
      <UploadCloud size={42}/><h2>{busy ? 'Extracting…' : 'Drop ZIP files here or click to choose'}</h2><p>Runs fully in your browser. No project files are uploaded to a server.</p>
      <input ref={inputRef} type="file" accept=".zip" multiple onChange={e => handleFiles(e.target.files)} />
    </section>
    {error && <div className="error"><X size={16}/>{error}</div>}
    <section className="results">{results.map(r => <ResultBlock key={r.id} result={r} onDelete={id => setResults(prev => prev.filter(x => x.id !== id))} onSelectionChange={updateSelection} />)}</section>
  </main>;
}

createRoot(document.getElementById('root')).render(<App />);
