import React, { useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { createExtractorFromData } from 'node-unrar-js/esm';
import unrarWasmUrl from 'node-unrar-js/esm/js/unrar.wasm?url';
import { saveAs } from 'file-saver';
import { PDFDocument } from 'pdf-lib';
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
import { Archive, ArrowDown, ArrowUp, ChevronDown, ChevronRight, Copy, Download, FileCode2, FileText, Folder, FolderOpen, HelpCircle, Merge, Moon, Plus, Settings, Sun, Trash2, UploadCloud, X } from 'lucide-react';
import './styles.css';
import { Analytics } from '@vercel/analytics/react';

hljs.registerLanguage('csharp', csharp);
hljs.registerLanguage('javascript', javascript);
hljs.registerLanguage('css', css);
hljs.registerLanguage('xml', xml);
hljs.registerLanguage('cpp', cpp);
hljs.registerLanguage('makefile', makefile);
hljs.registerLanguage('json', json);
hljs.registerLanguage('bash', bash);
hljs.registerLanguage('plaintext', plaintext);

const EXTENSION_OPTIONS = [
  { ext: '.cs', label: 'C# (.cs)', defaultOn: true },
  { ext: '.xaml', label: 'XAML (.xaml)', defaultOn: true },
  { ext: '.razor', label: 'Razor (.razor)', defaultOn: true },
  { ext: '.html', label: 'HTML (.html)', defaultOn: true },
  { ext: '.config', label: 'Config (.config)', defaultOn: true },
  { ext: '.c', label: 'C (.c)' },
  { ext: '.h', label: 'C/C++ Header (.h)' },
  { ext: '.cpp', label: 'C++ (.cpp)' },
  { ext: '.hpp', label: 'C++ Header (.hpp)' },
  { ext: '.js', label: 'JavaScript (.js)' },
  { ext: '.jsx', label: 'React JS (.jsx)' },
  { ext: '.ts', label: 'TypeScript (.ts)' },
  { ext: '.tsx', label: 'React TS (.tsx)' },
  { ext: '.css', label: 'CSS (.css)' },
  { ext: '.scss', label: 'SCSS (.scss)' },
  { ext: '.json', label: 'JSON (.json)' },
  { ext: '.xml', label: 'XML (.xml)' },
  { ext: '.sql', label: 'SQL (.sql)' },
  { ext: '.md', label: 'Markdown (.md)' },
  { ext: '.txt', label: 'Text (.txt)' },
  { ext: '.yml', label: 'YAML (.yml)' },
  { ext: '.yaml', label: 'YAML (.yaml)' },
  { ext: '.py', label: 'Python (.py)' },
  { ext: '.java', label: 'Java (.java)' },
  { ext: '.php', label: 'PHP (.php)' },
  { ext: '.rb', label: 'Ruby (.rb)' },
  { ext: '.go', label: 'Go (.go)' },
  { ext: '.rs', label: 'Rust (.rs)' },
  { ext: '.swift', label: 'Swift (.swift)' },
  { ext: '.kt', label: 'Kotlin (.kt)' },
  { ext: '.vb', label: 'VB.NET (.vb)' },
  { ext: '.fs', label: 'F# (.fs)' },
  { ext: '.sh', label: 'Shell (.sh)' },
  { ext: '.bat', label: 'Batch (.bat)' },
  { ext: '.ps1', label: 'PowerShell (.ps1)' },
  { ext: '.makefile', label: 'Makefile' },
  { ext: '.dockerfile', label: 'Dockerfile' },
  { ext: '.env', label: 'Environment (.env)' }
];
const DEFAULT_EXTENSION_SET = EXTENSION_OPTIONS.filter(o => o.defaultOn).map(o => o.ext);
const ALL_EXTENSION_SET = new Set(EXTENSION_OPTIONS.map(o => o.ext));

const DOCX_CODE_FONT_SIZE = 10; // docx uses half-points, so 10 = 5pt
const DOCX_TITLE_FONT_SIZE = 30; // 30 = 15pt

const DEFAULT_SETTINGS = {
  docxOrientation: 'portrait',
  docxSavePages: false,
  selectedExtensions: DEFAULT_EXTENSION_SET
};
const DEFAULT_MERGE_SETTINGS = {
  fileNameMode: 'first',
  customFileName: 'merged-files'
};
const SETTINGS_KEY = 'codeExtractorPro.settings.v4';
const MERGE_SETTINGS_KEY = 'codeExtractorPro.mergeSettings.v1';
const THEME_KEY = 'codeExtractorPro.theme.v1';
const ACTIVE_PAGE_KEY = 'codeExtractorPro.activePage.v1';
const DB_NAME = 'CodeExtractorProState';
const DB_VERSION = 1;
const STATE_STORE = 'state';
const STATE_KEY = 'main';

const CONTEXT_EXTENSIONS = new Set(['.csproj', '.vbproj', '.fsproj', '.sln']);
const SPECIAL_FILENAMES = new Map([
  ['makefile', '.makefile'],
  ['dockerfile', '.dockerfile'],
  ['app.config', '.config'],
  ['web.config', '.config'],
  ['.editorconfig', '.config']
]);
const IGNORE_PARTS = new Set(['node_modules', '.git', '.vs', 'bin', 'obj', 'dist', 'build', '.next', '.vercel', 'coverage', '.idea', '.vscode']);
const LARGE_DOCX_COLOR_CHAR_LIMIT = 3_000_000;

const TYPE_COLORS = {
  keyword: '0000FF', built_in: '2B91AF', type: '2B91AF', literal: '0000FF', number: '098658',
  string: 'A31515', regexp: 'A31515', symbol: 'A31515', comment: '008000', quote: '008000',
  doctag: '808080', meta: '808080', title: '795E26', section: '795E26', name: '800000',
  attr: 'FF0000', attribute: 'FF0000', variable: '001080', params: '001080', subst: '001080',
  tag: '800000', selector_tag: '800000', selector_id: '800000', selector_class: '800000',
  property: 'FF0000', function: '795E26', class: '2B91AF', default: '000000'
};

function normalizeSettings(raw) {
  const selectedExtensions = Array.isArray(raw?.selectedExtensions)
    ? raw.selectedExtensions.filter(ext => ALL_EXTENSION_SET.has(ext))
    : DEFAULT_SETTINGS.selectedExtensions;
  return {
    docxOrientation: raw?.docxOrientation === 'landscape' ? 'landscape' : 'portrait',
    docxSavePages: Boolean(raw?.docxSavePages),
    selectedExtensions: selectedExtensions.length ? selectedExtensions : DEFAULT_SETTINGS.selectedExtensions
  };
}
function loadSettings() {
  try { return normalizeSettings(JSON.parse(localStorage.getItem(SETTINGS_KEY) || 'null')); }
  catch { return DEFAULT_SETTINGS; }
}
function saveSettings(settings) {
  try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(normalizeSettings(settings))); } catch {}
}

function normalizeMergeSettings(raw) {
  const fileNameMode = raw?.fileNameMode === 'custom' ? 'custom' : 'first';
  const customFileName = String(raw?.customFileName || DEFAULT_MERGE_SETTINGS.customFileName).trim() || DEFAULT_MERGE_SETTINGS.customFileName;
  return { fileNameMode, customFileName };
}
function loadMergeSettings() {
  try { return normalizeMergeSettings(JSON.parse(localStorage.getItem(MERGE_SETTINGS_KEY) || 'null')); }
  catch { return DEFAULT_MERGE_SETTINGS; }
}
function saveMergeSettings(settings) {
  try { localStorage.setItem(MERGE_SETTINGS_KEY, JSON.stringify(normalizeMergeSettings(settings))); } catch {}
}
function loadTheme() {
  try {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored === 'dark' || stored === 'light') return stored;
  } catch {}
  return getSystemTheme();
}
function saveTheme(theme) {
  try { localStorage.setItem(THEME_KEY, theme); } catch {}
}
function loadActivePage() {
  if (typeof window === 'undefined') return 'extractor';
  if (window.location.hash === '#merger') return 'merger';
  if (window.location.hash === '#extractor') return 'extractor';
  try { return localStorage.getItem(ACTIVE_PAGE_KEY) === 'merger' ? 'merger' : 'extractor'; }
  catch { return 'extractor'; }
}
function saveActivePage(page) {
  try { localStorage.setItem(ACTIVE_PAGE_KEY, page); } catch {}
}
function openStateDb() {
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) { reject(new Error('IndexedDB unavailable')); return; }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => req.result.createObjectStore(STATE_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function saveAppState(state) {
  try {
    const db = await openStateDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STATE_STORE, 'readwrite');
      tx.objectStore(STATE_STORE).put(state, STATE_KEY);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
    db.close();
  } catch {}
}
async function loadAppState() {
  try {
    const db = await openStateDb();
    const value = await new Promise((resolve, reject) => {
      const tx = db.transaction(STATE_STORE, 'readonly');
      const req = tx.objectStore(STATE_STORE).get(STATE_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
    db.close();
    return value;
  } catch { return null; }
}
function serializeTree(node) {
  const children = {};
  for (const [key, child] of Object.entries(node.children || {})) children[key] = serializeTree(child);
  const copy = { ...node, children };
  delete copy.sortedChildren;
  return copy;
}
function deserializeResult(item) {
  const tree = sortTree(item.tree);
  return {
    ...item,
    tree,
    selectedPaths: new Set(item.selectedPaths || []),
    openNodeIds: new Set(item.openNodeIds || collectExpandableNodeIds(tree)),
    collapsed: Boolean(item.collapsed),
    leftWidth: item.leftWidth || 320
  };
}
function serializeResult(item) {
  return {
    ...item,
    tree: serializeTree(item.tree),
    selectedPaths: Array.from(item.selectedPaths || []),
    openNodeIds: Array.from(item.openNodeIds || []),
    collapsed: Boolean(item.collapsed),
    leftWidth: item.leftWidth || 320
  };
}

function normalizePath(path) { return path.replace(/\\/g, '/').replace(/^\/+/, ''); }
function getFileName(path) { return normalizePath(path).split('/').filter(Boolean).pop() || ''; }
function getExt(path) {
  const name = getFileName(path).toLowerCase();
  if (name === 'makefile' || name.endsWith('.mk')) return '.makefile';
  if (name === 'dockerfile') return '.dockerfile';
  const dot = name.lastIndexOf('.');
  return dot >= 0 ? name.slice(dot) : '';
}
function isIgnoredPath(path) {
  const parts = normalizePath(path).split('/');
  return parts.some(part => IGNORE_PARTS.has(part));
}
function shouldOutput(path, selectedExtensions = DEFAULT_EXTENSION_SET) {
  const selected = new Set(selectedExtensions);
  const name = getFileName(path).toLowerCase();
  const specialExt = SPECIAL_FILENAMES.get(name);
  const ext = specialExt || getExt(path);
  return ALL_EXTENSION_SET.has(ext) && selected.has(ext);
}
function shouldReadAsContext(path) { return CONTEXT_EXTENSIONS.has(getExt(path)); }
function depthOf(path) { return normalizePath(path).split('/').filter(Boolean).length; }
function safeBaseName(name) {
  return name.replace(/\.(zip|rar)$/i, '').replace(/[^a-z0-9._-]+/gi, '_') || 'code-extract';
}
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
  if (!lineTokens.length) return [new TextRun({ text: ' ', size: DOCX_CODE_FONT_SIZE, font: 'Consolas' })];
  return lineTokens.map(t => new TextRun({ text: t.text.replace(/\t/g, '    '), color: t.color, size: DOCX_CODE_FONT_SIZE, font: 'Consolas' }));
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
function downloadBlob(text, filename, type = 'text/plain;charset=utf-8') {
  saveAs(new Blob([text], { type }), filename);
}
async function copyText(text) {
  const blob = new Blob([text], { type: 'text/plain' });
  try {
    if (navigator.clipboard?.write && window.ClipboardItem) {
      await navigator.clipboard.write([new ClipboardItem({ 'text/plain': blob })]);
      return true;
    }
  } catch {}
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {}
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.setAttribute('readonly', '');
  ta.style.position = 'fixed';
  ta.style.left = '-9999px';
  ta.style.top = '0';
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  try { return document.execCommand('copy'); }
  finally { document.body.removeChild(ta); }
}
async function downloadDocx(result, output, settings = DEFAULT_SETTINGS) {
  const paragraphs = [];
  const para = (children, extra = {}) => new Paragraph({ alignment: AlignmentType.LEFT, bidirectional: false, children, ...extra });
  const useColor = output.text.length <= LARGE_DOCX_COLOR_CHAR_LIMIT;
  const savePages = settings.docxSavePages;
  const orientation = settings.docxOrientation === 'landscape' ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT;
  let firstFileBlock = true;
  function addPlain(line, bold = false, extra = {}) {
    paragraphs.push(para([new TextRun({ text: line || ' ', bold, size: DOCX_CODE_FONT_SIZE, font: 'Consolas', color: '000000' })], extra));
  }
  function addCodeLine(prefix, lineTokensOrText) {
    if (Array.isArray(lineTokensOrText)) {
      paragraphs.push(para([new TextRun({ text: prefix, size: DOCX_CODE_FONT_SIZE, font: 'Consolas' }), ...makeDocxRunsFromLine(lineTokensOrText)]));
    } else {
      paragraphs.push(para([new TextRun({ text: prefix + (lineTokensOrText || ' '), size: DOCX_CODE_FONT_SIZE, font: 'Consolas', color: '000000' })]));
    }
  }
  function addFileBlock(file, headers) {
    const pageBreakBefore = !savePages && !firstFileBlock;
    headers.forEach((header, index) => addPlain(header.text, header.bold, index === 0 ? { pageBreakBefore, keepNext: true } : { keepNext: true }));
    addPlain(`${indent(headers.length + 1)}~~~`, false, { keepNext: true });
    const prefix = indent(headers.length + 1);
    if (useColor) {
      const codeLines = tokensToLines(highlightedTokens(file.content, file.language));
      for (const line of codeLines) addCodeLine(prefix, line);
    } else {
      for (const line of file.content.replace(/\t/g, '    ').split(/\r\n|\n|\r/)) addCodeLine(prefix, line);
    }
    addPlain(`${indent(headers.length + 1)}~~~`);
    firstFileBlock = false;
  }
  function walk(node, level, headers) {
    const hasSelectedDescendant = collectFilePaths(node).some(p => result.selectedPaths.has(p));
    if (!hasSelectedDescendant) return;
    const isFolderish = node.type === 'folder';
    const title = `${indent(level)}${node.name}${isFolderish ? ` (${node.label || 'folder'}):` : ':'}`;
    const nextHeaders = [...headers, { text: title, bold: isFolderish }];
    if ((node.type === 'file') && node.file && result.selectedPaths.has(node.file.path)) {
      addFileBlock(node.file, nextHeaders);
      return;
    }
    node.sortedChildren?.forEach(child => walk(child, level + 1, nextHeaders));
  }
  paragraphs.push(para([
  new TextRun({
    text: safeBaseName(result.name),
    bold: true,
    size: DOCX_TITLE_FONT_SIZE,
    font: 'Consolas',
    color: '000000'
  })
], {
  alignment: AlignmentType.CENTER,
  spacing: { after: 240 }
}));

const rootHeader = { text: `${result.rootName} (${result.tree.solutionLabel || 'zip project'}):`, bold: true };
result.tree.sortedChildren.forEach(child => walk(child, 1, [rootHeader]));
if (!paragraphs.length) addPlain(`${result.rootName} (${result.tree.solutionLabel || 'zip project'}):`, true);
  if (!useColor) addPlain('[Large DOCX safety mode: syntax colors were skipped to keep the export stable.]');
  const doc = new Document({
    title: safeBaseName(result.name),
    sections: [{
      properties: { page: { size: { orientation }, margin: { top: 360, right: 360, bottom: 360, left: 360 } } },
      children: paragraphs
    }]
  });
  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${safeBaseName(result.name)}.docx`);
}

function buildArchiveResult(file, files, settings = DEFAULT_SETTINGS) {
  const rootName = safeBaseName(file.name);
  const archiveKind = file.name.toLowerCase().endsWith('.rar') ? 'rar project' : 'zip project';

  const tree = newNode(rootName, 'folder');
  tree.solutionLabel = archiveKind;

  const projectMap = {};
  const outputFiles = [];

  for (const item of files) {
    const path = normalizePath(item.path);
    if (shouldReadAsContext(path)) {
      const type = detectProjectType(path, item.content);
      if (getExt(path) === '.sln') tree.solutionLabel = 'solution';
      else projectMap[path] = type;
    }
  }

  for (const item of files) {
    const path = normalizePath(item.path);
    if (!shouldOutput(path, settings.selectedExtensions) || shouldReadAsContext(path)) continue;

    const fileObj = {
      path,
      name: getFileName(path),
      content: item.content,
      language: detectLanguage(path),
      ext: getExt(path)
    };

    outputFiles.push(fileObj);
    insertTreePath(tree, path, fileObj);
  }

  applyProjectLabels(tree, projectMap);
  sortTree(tree);

  const allFilePaths = outputFiles.map(f => f.path);
  const selectedPaths = new Set(allFilePaths);
  const output = buildSelectedOutput(tree, rootName, selectedPaths);

  return {
    id: crypto.randomUUID(),
    name: file.name,
    rootName,
    tree,
    fileCount: outputFiles.length,
    allFilePaths,
    selectedPaths,
    openNodeIds: new Set(collectExpandableNodeIds(tree)),
    collapsed: false,
    leftWidth: 320,
    createdAt: new Date().toLocaleString()
  };
}

async function parseZip(file, settings = DEFAULT_SETTINGS) {
  const zip = await JSZip.loadAsync(file);

  const entries = Object.values(zip.files).filter(e => {
    const path = normalizePath(e.name);
    return !e.dir && !isIgnoredPath(path, settings) && (shouldOutput(path, settings.selectedExtensions) || shouldReadAsContext(path));
  });

  const files = [];
  for (const entry of entries) {
    const path = normalizePath(entry.name);
    const content = stripBom(await entry.async('string'));
    files.push({ path, content });
  }

  return buildArchiveResult(file, files, settings);
}

async function parseRar(file, settings = DEFAULT_SETTINGS) {
  const data = await file.arrayBuffer();
  const wasmBinary = await fetch(unrarWasmUrl).then(response => response.arrayBuffer());

  const extractor = await createExtractorFromData({ data, wasmBinary });

  const list = extractor.getFileList();
  const headers = [...list.fileHeaders];

  const wantedNames = headers
    .filter(header => {
      const path = normalizePath(header.name);
      return !header.flags.directory &&
        !isIgnoredPath(path, settings) &&
        (shouldOutput(path, settings.selectedExtensions) || shouldReadAsContext(path));
    })
    .map(header => header.name);

  const extracted = extractor.extract({ files: wantedNames });
  const extractedFiles = [...extracted.files];

  const decoder = new TextDecoder('utf-8');
  const files = extractedFiles
    .filter(item => item.extraction && !item.fileHeader.flags.directory)
    .map(item => ({
      path: normalizePath(item.fileHeader.name),
      content: stripBom(decoder.decode(item.extraction))
    }));

  return buildArchiveResult(file, files, settings);
}

async function parseArchive(file, settings = DEFAULT_SETTINGS) {
  const name = file.name.toLowerCase();

  if (name.endsWith('.zip')) return parseZip(file, settings);
  if (name.endsWith('.rar')) return parseRar(file, settings);

  throw new Error('Unsupported archive type');
}
function TriStateBox({ state, onClick }) {
  return <button
    className={`checkbox ${state}`}
    onClick={e => { e.stopPropagation(); onClick(); }}
    aria-label="Toggle selection"
    title="Toggle selection"
  >{state === 'all' ? '✓' : state === 'partial' ? '—' : ''}</button>;
}
function collectExpandableNodeIds(node) {
  const ids = [];
  const isExpandable = (node.type === 'folder') && node.sortedChildren?.length > 0;
  if (isExpandable) ids.push(node.id);
  node.sortedChildren?.forEach(child => ids.push(...collectExpandableNodeIds(child)));
  return ids;
}
function TreeNode({ node, depth = 0, selectedPaths, setSelectedPaths, openNodes, setOpenNodes }) {
  const isExpandable = (node.type === 'folder') && node.sortedChildren?.length > 0;
  const isFolderish = node.type === 'folder';
  const open = isExpandable && openNodes.has(node.id);
  const selectionState = getSelectionState(node, selectedPaths);
  const toggleOpen = () => {
    if (!isExpandable) return;
    setOpenNodes(prev => {
      const next = new Set(prev);
      if (next.has(node.id)) next.delete(node.id);
      else next.add(node.id);
      return next;
    });
  };
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
    <div className="tree-row" style={{ paddingLeft: 8 + depth * 16 }} title={node.name} onClick={toggleOpen}>
      <span className="chevron">{isExpandable ? (open ? <ChevronDown size={DOCX_CODE_FONT_SIZE}/> : <ChevronRight size={DOCX_CODE_FONT_SIZE}/>) : <span className="chevron-spacer" />}</span>
      <TriStateBox state={selectionState} onClick={toggleSelected} />
      <span className="node-icon">{isFolderish ? (open ? <FolderOpen size={15}/> : <Folder size={15}/>) : <FileCode2 size={15}/>}</span>
      <span className="node-name">{node.name}</span>{node.label && <em>{node.label}</em>}
    </div>
    {isExpandable && open && node.sortedChildren?.map(child => <TreeNode key={child.name} node={child} depth={depth + 1} selectedPaths={selectedPaths} setSelectedPaths={setSelectedPaths} openNodes={openNodes} setOpenNodes={setOpenNodes} />)}
  </div>;
}
function ResultBlock({ result, settings, onDelete, onSelectionChange, onMetaChange }) {
  const [copied, setCopied] = useState(false);
  const leftWidth = result.leftWidth || 320;
  const collapsed = Boolean(result.collapsed);
  const openNodes = result.openNodeIds || new Set(collectExpandableNodeIds(result.tree));
  const dragging = useRef(false);
  const selectedCount = countSelected(result);
  const selectedOutput = useMemo(() => buildSelectedOutput(result.tree, result.rootName, result.selectedPaths), [result.tree, result.rootName, result.selectedPaths]);
  const setSelectedPaths = updater => onSelectionChange(result.id, updater);
  const setOpenNodes = updater => onMetaChange(result.id, item => ({ ...item, openNodeIds: typeof updater === 'function' ? updater(item.openNodeIds || new Set()) : updater }));
  const setCollapsed = updater => onMetaChange(result.id, item => ({ ...item, collapsed: typeof updater === 'function' ? updater(Boolean(item.collapsed)) : updater }));
  const setLeftWidth = value => onMetaChange(result.id, item => ({ ...item, leftWidth: value }));
  const startDrag = e => { dragging.current = true; e.preventDefault(); };
  React.useEffect(() => {
    const move = e => { if (dragging.current) setLeftWidth(Math.min(620, Math.max(210, e.clientX - 28))); };
    const up = () => { dragging.current = false; };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
    return () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
  }, []);
  return <section className={`result-card ${collapsed ? 'is-collapsed' : ''}`}>
    <header className="result-header">
      <div className="title-wrap">
        <button className="zip-collapse-button" onClick={() => setCollapsed(prev => !prev)} aria-expanded={!collapsed} title={collapsed ? 'Open ZIP block' : 'Close ZIP block'}>
          <Archive size={20}/>
        </button>
        <div><h2>{result.name}</h2><p>{selectedCount} out of {result.fileCount} files selected • {result.createdAt}</p></div>
      </div>
      <div className="actions">
        <button onClick={async () => { await copyText(selectedOutput.text); setCopied(true); setTimeout(() => setCopied(false), 1400); }}><Copy size={16}/>{copied ? 'Copied' : 'Copy text'}</button>
        <button onClick={() => downloadBlob(selectedOutput.text, `${safeBaseName(result.name)}.txt`)}><FileText size={16}/>TXT</button>
        <button onClick={() => downloadDocx(result, selectedOutput, settings)}><Download size={16}/>DOCX</button>
        <button className="danger" onClick={() => onDelete(result.id)}><Trash2 size={16}/>Delete</button>
      </div>
    </header>
    {!collapsed && <div className="workspace">
      <aside className="explorer" style={{ width: leftWidth }}>
        <div className="explorer-title"><span>File Explorer</span></div>
        <TreeNode node={result.tree} selectedPaths={result.selectedPaths} setSelectedPaths={setSelectedPaths} openNodes={openNodes} setOpenNodes={setOpenNodes} />
      </aside>
      <div className="resizer" onMouseDown={startDrag} title="Drag to resize" />
      <pre className="output"><code>{selectedOutput.text}</code></pre>
    </div>}
  </section>;
}

function GuidePanel({ open, onClose, kind }) {
  const isMerge = kind === 'merge';
  return <aside className={`guide-panel ${open ? 'open' : ''}`} aria-hidden={!open}>
    <div className="guide-head"><strong>{isMerge ? 'PDF Merger Guide' : 'CodeExtractor Guide'}</strong><button onClick={onClose}><X size={16}/></button></div>
    {isMerge ? <div className="guide-content">
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
    </div> : <div className="guide-content">
      <h3>How to use CodeExtractor Pro</h3>
      <p>Use this page to turn project ZIP files into clean text for AI prompts, reviews, and documentation.</p>
      <ol>
        <li>Drop one or more project ZIP/RAR files into the upload box, or click the box to pick them.</li>
        <li>Each ZIP opens in its own block with a Solution Explorer on the left and generated text on the right.</li>
        <li>Select or unselect files and folders with the blue checkboxes. The output updates automatically.</li>
        <li>Drag the divider to resize the Solution Explorer.</li>
        <li>Use Copy text, TXT, or DOCX to export the selected code.</li>
        <li>Open Settings to choose default DOCX layout and which file types are extracted from future ZIP/RAR uploads.</li>
      </ol>
      <div className="guide-he" dir="rtl">
        <h3>מדריך בעברית</h3>
        <p>העמוד הזה הופך קבצי ZIP של פרויקטים לטקסט נקי ומסודר לשימוש עם AI, בדיקות ותיעוד.</p>
        <ol>
          <li>גרור קובץ ZIP/RAR אחד או יותר לתיבת ההעלאה, או לחץ עליה כדי לבחור מהמחשב.</li>
          <li>כל ZIP נפתח בבלוק משלו עם Solution Explorer בצד שמאל וטקסט שנוצר בצד ימין.</li>
          <li>בחר או בטל בחירה של קבצים ותיקיות בעזרת תיבות הסימון הכחולות.</li>
          <li>גרור את הקו המפריד כדי לשנות את רוחב ה־Solution Explorer.</li>
          <li>השתמש ב־Copy text, TXT או DOCX כדי לייצא את הקוד שנבחר.</li>
          <li>פתח Settings כדי לבחור הגדרות DOCX וסוגי קבצים שיופקו מה־ZIP/RAR בהעלאות הבאות.</li>
        </ol>
      </div>
      <span className="owner-watermark" aria-hidden="true">Lior Avni 15/05/2026</span>
    </div>}
  </aside>;
}
function SettingsModal({ page, settings, mergeSettings, onChange, onMergeChange, onClose, onResetPage }) {
  const isMerger = page === 'merger';
  const selected = new Set(settings.selectedExtensions);
  const update = patch => onChange(normalizeSettings({ ...settings, ...patch }));
  const updateMerge = patch => onMergeChange(normalizeMergeSettings({ ...mergeSettings, ...patch }));
  const toggleExt = ext => {
    const next = new Set(selected);
    if (next.has(ext)) next.delete(ext); else next.add(ext);
    update({ selectedExtensions: Array.from(next) });
  };
  return <div className="modal-backdrop" onMouseDown={onClose}>
    <section className="settings-modal" onMouseDown={e => e.stopPropagation()}>
      <div className="settings-head"><div><p className="eyebrow">CodeExtractor Pro</p><h2>{isMerger ? 'PDF Merger Settings' : 'CodeExtractor Settings'}</h2></div><button onClick={onClose}><X size={18}/></button></div>
      {isMerger ? <>
        <div className="settings-section">
          <h3>Merged PDF file name</h3>
          <label className="setting-row"><span>Download name</span><select value={mergeSettings.fileNameMode} onChange={e => updateMerge({ fileNameMode: e.target.value })}><option value="first">Use the first PDF name + _merged</option><option value="custom">Use a custom name</option></select></label>
          <label className="setting-row setting-row-stacked"><span>Custom file name</span><input type="text" value={mergeSettings.customFileName} onChange={e => updateMerge({ customFileName: e.target.value })} placeholder="merged-files" disabled={mergeSettings.fileNameMode !== 'custom'} /></label>
          <p className="settings-note">PDF files, their order, and this setting are saved in this browser and restored after reload.</p>
        </div>
      </> : <>
        <div className="settings-section">
          <h3>DOCX export</h3>
          <label className="setting-row"><span>Page orientation</span><select value={settings.docxOrientation} onChange={e => update({ docxOrientation: e.target.value })}><option value="portrait">Vertical / Portrait</option><option value="landscape">Horizontal / Landscape</option></select></label>
          <label className="setting-check"><input type="checkbox" checked={settings.docxSavePages} onChange={e => update({ docxSavePages: e.target.checked })}/><span>Save pages: do not start every code file on a new page</span></label>
        </div>
        <div className="settings-section">
          <div className="settings-section-title"><h3>ZIP/RAR file types and folders</h3><button onClick={() => update({ selectedExtensions: DEFAULT_EXTENSION_SET })}>Reset defaults</button><button onClick={() => update({ selectedExtensions: Array.from(ALL_EXTENSION_SET) })}>Select all</button></div>
          <p className="settings-note">These choices apply to the next ZIP/RAR files you upload.</p>
          <div className="extension-grid">
            {EXTENSION_OPTIONS.map(option => <label key={option.ext} className="extension-pill"><input type="checkbox" checked={selected.has(option.ext)} onChange={() => toggleExt(option.ext)}/><span>{option.label}</span></label>)}
          </div>
        </div>
      </>}
      <div className="settings-reset-zone">
        <button className="settings-reset-button" onClick={onResetPage}>Reset this page to the beginning</button>
        <p>{isMerger ? 'Clears all selected PDF files, order, and PDF Merger settings.' : 'Clears all uploaded projects, selections, explorer state, and CodeExtractor settings.'}</p>
      </div>
    </section>
  </div>;
}

const PDF_MIME = 'application/pdf';
function fileRow(file = null) { return { id: crypto.randomUUID(), file }; }
function ensureTwoRows(rows) {
  const next = Array.isArray(rows) ? rows.filter(Boolean) : [];
  while (next.length < 2) next.push(fileRow());
  return next;
}
function delayFrame() { return new Promise(resolve => requestAnimationFrame(resolve)); }
function mergeKind(file) {
  const name = file?.name?.toLowerCase() || '';
  if (name.endsWith('.pdf')) return 'pdf';
  return '';
}
function safeMergedNameFromText(text) {
  return String(text || 'merged-files').replace(/\.pdf$/i, '').replace(/[^a-z0-9._-]+/gi, '_') || 'merged-files';
}
function safeMergedName(file) {
  return safeMergedNameFromText(file?.name || 'merged-files.pdf');
}
function validateFiles(files) {
  if (files.length < 2) return { ok: false, message: 'Choose at least 2 PDF files.' };
  if (files.some(file => mergeKind(file) !== 'pdf')) return { ok: false, message: 'Only .pdf files are supported.' };
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
function mergedDownloadName(files, mergeSettings) {
  const normalized = normalizeMergeSettings(mergeSettings);
  const base = normalized.fileNameMode === 'custom' ? safeMergedNameFromText(normalized.customFileName) : `${safeMergedName(files[0])}_merged`;
  return `${base}.pdf`;
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
function PdfMergerPage({ rows, setRows, mergeSettings, busy, setBusy, message, setMessage, guideOpen, setGuideOpen }) {
  const setFile = (id, file) => {
    if (!mergeKind(file)) { setMessage('Please choose a .pdf file.'); return; }
    setRows(prev => ensureTwoRows(prev.map(row => row.id === id ? { ...row, file } : row)));
    setMessage('Use the arrow buttons to choose the merge order. Your selected PDFs are saved for this browser.');
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
  const deleteRow = id => setRows(prev => ensureTwoRows(prev.length <= 2 ? prev.map(row => row.id === id ? { ...row, file: null } : row) : prev.filter(row => row.id !== id)));
  const readyFiles = rows.map(row => row.file).filter(Boolean);
  const validation = validateFiles(readyFiles);
  const canMerge = validation.ok && !busy;
  const doMerge = async () => {
    if (!validation.ok) { setMessage(validation.message); return; }
    setBusy(true);
    setMessage('Merging PDF files...');
    try {
      const blob = await mergePdfFiles(readyFiles);
      saveAs(blob.type ? blob : new Blob([blob], { type: PDF_MIME }), mergedDownloadName(readyFiles, mergeSettings));
      setMessage(`Merged ${readyFiles.length} PDF files successfully.`);
    } catch (e) {
      console.error(e);
      setMessage(e?.message || 'Could not merge the selected files. Very complex or encrypted files may not be mergeable in the browser.');
    } finally {
      setBusy(false);
    }
  };
  return <main className={`merge-page ${guideOpen ? 'with-guide' : ''}`}>
    <section className="merge-hero">
      <div>
        <p className="eyebrow">CodeExtractor Pro</p>
        <h1>PDF Merger</h1>
        <div className="subtitle-line"><button className="guide-button" onClick={() => setGuideOpen(prev => !prev)} aria-expanded={guideOpen}><HelpCircle size={16}/>How To Use</button><p>Upload several PDF files, choose their order, and download one merged PDF.</p></div>
      </div>
      <div className="stats"><strong>{readyFiles.length}</strong><span>PDF files ready</span><strong>{rows.length}</strong><span>merge slots</span></div>
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

function getSystemTheme() {
  if (typeof window === 'undefined') return 'light';
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function App() {
  const [activePage, setActivePageState] = useState(loadActivePage);
  const [results, setResults] = useState([]);
  const [mergeRows, setMergeRows] = useState(() => ensureTwoRows([]));
  const [busy, setBusy] = useState(false);
  const [mergeBusy, setMergeBusy] = useState(false);
  const [error, setError] = useState('');
  const [mergeMessage, setMergeMessage] = useState('Choose at least 2 PDF files. The merge happens fully in your browser.');
  const [theme, setTheme] = useState(loadTheme);
  const [settings, setSettings] = useState(loadSettings);
  const [mergeSettings, setMergeSettings] = useState(loadMergeSettings);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [guideOpen, setGuideOpen] = useState(false);
  const [loadedSavedState, setLoadedSavedState] = useState(false);
  const inputRef = useRef(null);
  const totalFiles = useMemo(() => results.reduce((sum, r) => sum + r.fileCount, 0), [results]);

  const setActivePage = page => {
    setGuideOpen(false);
    setActivePageState(page);
    saveActivePage(page);
    if (typeof window !== 'undefined') {
      const hash = page === 'merger' ? '#merger' : '#extractor';
      if (window.location.hash !== hash) window.history.replaceState(null, '', hash);
    }
  };

  React.useEffect(() => {
    const onHashChange = () => setActivePageState(window.location.hash === '#merger' ? 'merger' : 'extractor');
    window.addEventListener('hashchange', onHashChange);
    return () => window.removeEventListener('hashchange', onHashChange);
  }, []);
  React.useEffect(() => {
    document.documentElement.dataset.theme = theme;
    saveTheme(theme);
  }, [theme]);
  React.useEffect(() => { saveSettings(settings); }, [settings]);
  React.useEffect(() => { saveMergeSettings(mergeSettings); }, [mergeSettings]);
  React.useEffect(() => {
    let alive = true;
    loadAppState().then(saved => {
      if (!alive) return;
      if (saved?.results) setResults(saved.results.map(deserializeResult));
      if (saved?.mergeRows) setMergeRows(ensureTwoRows(saved.mergeRows));
      setLoadedSavedState(true);
    });
    return () => { alive = false; };
  }, []);
  React.useEffect(() => {
    if (!loadedSavedState) return;
    saveAppState({
      results: results.map(serializeResult),
      mergeRows: ensureTwoRows(mergeRows).map(row => ({ id: row.id, file: row.file || null }))
    });
  }, [results, mergeRows, loadedSavedState]);

  async function handleFiles(fileList) {
    const archives = Array.from(fileList || []).filter(f => {
      const name = f.name.toLowerCase();
      return name.endsWith('.zip') || name.endsWith('.rar');
    });
    if (!archives.length) {
      setError('Please choose one or more .zip or .rar files.');
      return;
    }
    setBusy(true);
    setError('');
    try {
      const parsed = [];
      for (const archive of archives) parsed.push(await parseArchive(archive, settings));
      setResults(prev => [...parsed, ...prev]);
    } catch (e) {
      console.error(e);
      setError('Could not parse one of the archive files. Make sure it is a valid .zip/.rar project archive and not password-protected.');
    } finally {
      setBusy(false);
    }
  }
  const updateSelection = (id, updater) => {
    setResults(prev => prev.map(item => {
      if (item.id !== id) return item;
      const selectedPaths = typeof updater === 'function' ? updater(item.selectedPaths) : updater;
      return { ...item, selectedPaths };
    }));
  };
  const updateResultMeta = (id, updater) => setResults(prev => prev.map(item => item.id === id ? updater(item) : item));
  const toggleTheme = () => setTheme(prev => prev === 'dark' ? 'light' : 'dark');
  const resetExtractorPage = () => {
    setResults([]);
    setSettings(DEFAULT_SETTINGS);
    setError('');
    setSettingsOpen(false);
  };
  const resetMergerPage = () => {
    setMergeRows(ensureTwoRows([]));
    setMergeSettings(DEFAULT_MERGE_SETTINGS);
    setMergeMessage('Choose at least 2 PDF files. The merge happens fully in your browser.');
    setSettingsOpen(false);
  };
  const resetCurrentPage = () => {
    if (activePage === 'merger') resetMergerPage();
    else resetExtractorPage();
  };

  return <>
    <GuidePanel open={guideOpen} onClose={() => setGuideOpen(false)} kind={activePage === 'merger' ? 'merge' : 'extractor'} />
    {settingsOpen && <SettingsModal page={activePage} settings={settings} mergeSettings={mergeSettings} onChange={setSettings} onMergeChange={setMergeSettings} onClose={() => setSettingsOpen(false)} onResetPage={resetCurrentPage} />}
    <div className="top-actions">
      <button className="merge-nav-button" onClick={() => setActivePage(activePage === 'merger' ? 'extractor' : 'merger')} title={activePage === 'merger' ? 'Go to CodeExtractor Pro' : 'Go to PDF Merger'}>
        {activePage === 'merger' ? <FileCode2 size={17}/> : <Merge size={17}/>} {activePage === 'merger' ? 'CodeExtractor' : 'PDF Merger'}
      </button>
      <button className="theme-toggle" onClick={() => setSettingsOpen(true)}><Settings size={17}/>Settings</button>
      <button className="theme-toggle" onClick={toggleTheme} aria-label={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}>
        {theme === 'dark' ? <Sun size={17}/> : <Moon size={17}/>}
        {theme === 'dark' ? 'Light' : 'Dark'}
      </button>
    </div>
    {activePage === 'merger' ? <PdfMergerPage rows={mergeRows} setRows={setMergeRows} mergeSettings={mergeSettings} busy={mergeBusy} setBusy={setMergeBusy} message={mergeMessage} setMessage={setMergeMessage} guideOpen={guideOpen} setGuideOpen={setGuideOpen} /> : <main className={guideOpen ? 'with-guide' : ''}>
      <section className="hero">
        <div><p className="eyebrow">CodeExtractor Pro</p><h1>Turn code project ZIPs and RARs into clean text.</h1><div className="subtitle-line"><button className="guide-button" onClick={() => setGuideOpen(prev => !prev)} aria-expanded={guideOpen}><HelpCircle size={16}/>How To Use</button><p className="subtitle">Multiple ZIPs and RARs, solution-style structure, selectable files, TXT export, and DOCX export with syntax-colored code at 5pt.</p></div></div>
        <div className="stats"><strong>{results.length}</strong><span>archives loaded</span><strong>{totalFiles}</strong><span>files extracted</span></div>
      </section>
      <section className="dropzone" onDrop={e => { e.preventDefault(); handleFiles(e.dataTransfer.files); }} onDragOver={e => e.preventDefault()} onClick={() => inputRef.current?.click()}>
        <UploadCloud size={42}/><h2>{busy ? 'Extracting…' : 'Drop ZIP/RAR files here or click to choose'}</h2><p>Runs fully in your browser. No project files are uploaded to a server.</p>
        <input ref={inputRef} type="file" accept=".zip,.rar,application/zip,application/x-rar-compressed" multiple onChange={e => handleFiles(e.target.files)} />
      </section>
      {error && <div className="error"><X size={16}/>{error}</div>}
      <section className="results">{results.map(r => <ResultBlock key={r.id} result={r} settings={settings} onDelete={id => setResults(prev => prev.filter(x => x.id !== id))} onSelectionChange={updateSelection} onMetaChange={updateResultMeta} />)}</section>
    </main>}
  </>;
}

createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
    <Analytics />
  </React.StrictMode>
);
