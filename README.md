# CodeExtractor Pro

A Vercel-ready React/Vite website that turns one or more project `.zip` files into organized AI-ready text output.

## Features

- Drag/drop or click to upload multiple ZIP files.
- Each ZIP gets its own result block.
- Left-side Solution Explorer with draggable resize handle.
- Right-side complete text extraction.
- Copy, download `.txt`, download `.docx`, and delete per ZIP.
- `.sln` and `.csproj`/`.vbproj`/`.fsproj` are parsed for project-structure context only. Their contents are not printed in the output.
- DOCX export uses 7pt Consolas and syntax-colored code runs generated in the browser.
- Runs locally in the browser; project files are not uploaded to a backend.

## Supported output files

Includes common code/config files such as `.cs`, `.xaml`, `.c`, `.h`, Makefile, `.as`, `.css`, `.js`, `.html`, `.config`, `.razor`, `.json`, `.xml`, `.ts`, `.tsx`, `.jsx`, `.cpp`, `.hpp`, `.sql`, `.md`, `.yml`, `.yaml`, and more.

## How to run locally

```bash
npm install
npm run dev
```

## How to deploy on Vercel

1. Upload this folder to GitHub.
2. Import the repository in Vercel.
3. Framework preset: Vite.
4. Build command: `npm run build`.
5. Output directory: `dist`.

## Notes about DOCX colors

A ZIP contains plain text, not Visual Studio formatting. The app recreates syntax coloring by highlighting source text with `highlight.js`, then writing colored Word text runs using `docx`. This gives a Visual-Studio-like colored DOCX export without needing a server.
