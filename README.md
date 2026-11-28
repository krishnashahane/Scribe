# SCRIBE

A collaborative word processor built around one rule: **never lose or silently change a user's document.**
It opens straight into a blank page, no sign-in required.
SCRIBE opens `.docx` files, lets you edit them on real pages, and writes them back
with everything it doesn't understand preserved byte-for-byte.

**What you can do:** write on paginated pages with styles (a styles gallery with Modify / New style), lists,
tables, images, headers and footers, sections, comments, tracked changes and real-time collaboration; insert
footnotes, endnotes, charts, shapes, text boxes, freehand drawings and equations; edit the text of text boxes and diagrams (SmartArt), the
data of charts and custom XML parts that come in imported files; cite sources and generate a bibliography
(APA, MLA, Chicago); run a mail merge from a CSV file; check grammar and style on your own device; and export
to DOCX, PDF, HTML or plain text.

## Quick start

Live: https://scribe-pi-indol.vercel.app

Requirements: Node.js 22+, npm.

```bash
npm install
npm run dev          # web on http://localhost:5173, API on :8787 (embedded PGlite database in ./data)
```

Production build served locally:

```bash
npm run build
npm start            # http://127.0.0.1:8787
```

Open the address and start typing. Each browser gets a private local identity automatically and SCRIBE
reopens your last document. Accounts are optional (File → Save to an account): use one to share documents,
open them on other devices, or keep them if browser data is cleared. Upgrading keeps every document.

### Configuration

Copy `.env.example` to `.env.local`. Every variable is server-side only (none use a `VITE_` prefix), and on
Vercel they're set in the project settings, never in code.

| Variable | Purpose |
|---|---|
| `DATABASE_URL` | PostgreSQL (Neon on Vercel). Unset locally → embedded PGlite in `./data/pglite` |
| `ANTHROPIC_API_KEY` | Enables AI assistance with Claude (default model `claude-opus-5-5`) |
| `SCRIBE_AI_PROVIDER`, `SCRIBE_AI_MODEL`, `SCRIBE_AI_BASE_URL`, `SCRIBE_AI_API_KEY` | Alternative / OpenAI-compatible AI providers |
| `ALLOWED_ORIGINS` | Extra origins allowed to call the API (default: same origin only) |
| `PORT`, `HOST`, `SCRIBE_DB_DIR` | Local server only |

### Deploying to Vercel

The client is a static Vite build; the API is one Vercel Function (`api/index.js`, Express) backed by Postgres.
`vercel.json` sets the build, the `/api/*` rewrite, an app fallback so links written as paths (`/documents`,
`/doc/<id>`) open the right screen, and security headers for every response.

```bash
vercel link --project scribe
vercel integration add neon      # provisions Postgres and sets DATABASE_URL in the project
vercel env add ANTHROPIC_API_KEY production   # optional
vercel --prod
```

The database schema is created and migrated automatically on first request.

## Tests

```bash
npm run lint              # ESLint (typescript-eslint, react-hooks, no-console)
npm run typecheck         # TypeScript
npm test                  # 66 unit/integration tests: DOCX import/export, security, server, collaboration
npm run test:e2e          # 11 browser tests (Playwright + Chromium) against the production build
LIVE_URL=https://… npx playwright test -c playwright.live.config.ts   # same suite against a deployment
npm run fixtures          # regenerate the DOCX corpus in tests/fixtures
npm run verify:python-docx  # open every exported file with python-docx (pip install python-docx)
./scripts/verify-libreoffice.sh  # open every exported file in LibreOffice; compare pages and text with the originals
```

## Architecture

```
src/model        document model: types, Tiptap/ProseMirror schema extensions, templates, migrations
src/docx         OOXML engine: safe package reader, styles/numbering, importer, exporter, validator
src/layout       pagination (measured layout → page-break decorations, line-level paragraph splitting)
src/editor       find/replace, track-changes (suggesting) mode, section bookkeeping
src/formats      HTML (sanitised) and TXT import/export
src/collab       Yjs session: IndexedDB local copy + WebSocket sync
src/ui           React interface (ribbon, page canvas, panels, dialogs)
server/          Express API: Postgres storage, auth + RBAC, zod validation, HTTP Yjs sync, chunked uploads, AI
api/             Vercel Function entry
tests/           unit, integration, e2e, DOCX fixture corpus
```

**Document model.** The body is ProseMirror JSON with stable block ids (`nid`). Paragraph and run formatting is
stored in OOXML units (twips, half-points) so values round-trip exactly. Pages, sections, headers/footers,
comments, notes, styles and metadata live in `DocMeta`. Documents carry a schema version and pass through
`migrateDocument()` on load.

**DOCX round trip.** Import parses the package (zip-bomb, path-traversal and DOCTYPE/entity guards), resolves
styles, themes and numbering, and maps OOXML to the model. Anything SCRIBE cannot edit (text boxes, shapes, charts,
equations, SmartArt, content-control TOCs, unknown elements, unknown property children) is kept **verbatim** in
opaque nodes or `*Extra` attributes and reported in the compatibility report. Export starts from the **original
package**, replaces only the story parts SCRIBE edits (document, headers, footers, comments when changed), adds
new numbering/styles/media, and keeps every other part untouched. Every export runs through `validatePackage()`
(well-formed XML, relationships, content types, schema child order, numbering and comment references) and is
refused if it fails. Exports are always new files; the uploaded original is never modified.

**Persistence and recovery.** The live document is a Yjs CRDT. Edits are saved to IndexedDB on the device
immediately (crash/offline recovery) and synced over HTTP to the server (push on change, pull every ~1.2 s),
which appends them to Postgres and periodically compacts them. HTTP sync works on serverless hosts that don't
support WebSockets. Images and original DOCX files are stored in Postgres in chunks that fit serverless request
limits; images are served from `/api/media/<id>`. JSON snapshots are saved every 4 seconds, automatic versions every
10 minutes, named versions on demand (`⌘/Ctrl+S`).

**Security.** zod validation on every request body and route parameter; server-side sanitising of stored
content (unsafe link/image URLs and unknown node types removed); authorization on every route (owner / editor /
viewer; viewers' sync writes are refused); rate limits stored in the database (auth, guest creation, uploads, AI,
general API) so they hold across serverless instances; parameterised SQL only; strict same-origin CORS; CSRF header
check; CSP, HSTS, X-Frame-Options, X-Content-Type-Options, Referrer-Policy, COOP, Permissions-Policy; HttpOnly
SameSite=Strict Secure session cookies, replaced on every sign-in and when a guest upgrades to an account; expired
sessions, stale presence and abandoned uploads cleaned up automatically; share invitations rate-limited; scrypt password hashes with constant-time comparison; no-login guest
identities can't be signed into with a password; uploads verified by magic bytes; generic error messages with no
stack traces; no source maps in production; XML entity and zip-bomb defences in the DOCX reader.

## DOCX compatibility matrix

Tested with a corpus generated by the independent `docx` library plus hand-injected OOXML and malformed files
(`tests/fixtures`). "Round-trip" means import → export → re-import produces an identical model, stable over
5 cycles.

| Feature | Status |
|---|---|
| Paragraphs, runs, bold/italic/underline/strike, colour, size, font, highlight, shading, sub/superscript, caps | Edit + round-trip |
| Styles (paragraph/character), theme fonts, docDefaults, inheritance | Rendered + round-trip; unknown style props preserved |
| Alignment, indentation (left/right/first/hanging), spacing before/after, line spacing, RTL paragraphs | Edit + round-trip |
| Bulleted, numbered, multi-level lists (incl. start overrides) | Edit + round-trip |
| Tables: merges (row/column), nested tables, shading, widths, header rows, borders | Edit + round-trip |
| Inline images (PNG/JPEG/GIF/BMP/WebP) | Edit + round-trip |
| Floating images | Anchor/wrap XML preserved; previewed floated left/right |
| Hyperlinks (external and to bookmarks), bookmarks | Edit + round-trip (bookmark ranges collapse to a point; reported) |
| Headers/footers (default + first page), PAGE/NUMPAGES fields | Edit + round-trip |
| Even-page headers/footers | Preserved; not shown |
| Sections, section breaks, page size, margins, orientation per section | Edit + round-trip |
| Section columns (1–3) | Set from Layout → Columns; shown in columns while editing, exported and printed in columns |
| Page breaks, column breaks | Edit + round-trip |
| Comments | Edit + round-trip; threaded replies flattened if comments are edited (reported) |
| Tracked insertions/deletions | Shown, accept/reject, suggesting mode, round-trip |
| Footnotes/endnotes | Insert (⌘/Ctrl+Alt+F, ⌘/Ctrl+Alt+D), edit, delete; numbered automatically; footnotes at the bottom of their page, endnotes after the text; untouched notes kept byte-for-byte |
| Fields (TOC, REF, PAGEREF, dates…) | Preserved with cached result |
| Merge fields (MERGEFIELD) | Insert, preview with CSV data, merge to a new document or DOCX; round-trip |
| Citations (CITATION fields), bibliography, sources part | Insert / edit sources, APA 7 / MLA 9 / Chicago author-date; round-trip |
| Equations | Edit as linear text with live preview (fractions, scripts, roots, sums, integrals, brackets, symbols); other structures (matrices, accents…) editable piece by piece; round-trip |
| Text boxes | Text editable (all copies kept in sync); size, position and styling preserved |
| Charts (column, bar, line, area, pie, doughnut) | Insert new charts; data, series, categories and title editable; imported formatting preserved; edited data stored in the chart |
| SmartArt | Node text editable in the data part and the cached drawing; layout preserved |
| Shapes (rectangle, rounded rectangle, oval, triangle, diamond, arrow, star, line) | Insert + edit (fill, outline, size, text); exported as native shapes; round-trip |
| Freehand drawing | Inserted as a PNG picture |
| Custom XML parts | Editable as XML (well-formedness checked); only edited parts rewritten |
| Paragraph styles | Gallery, Modify Style (font, size, colour, bold/italic/underline, alignment, spacing), New Style; written to the style sheet with unmanaged properties kept |
| Other drawings, combination charts, OLE objects, macros | Preserved verbatim, shown as placeholders, reported; macros are never run |
| Content controls | Unwrapped (text kept, control properties dropped; reported) |
| Legacy `.doc` | Rejected with a clear message |

## Known limitations

- Exports are verified with SCRIBE's structural validator, python-docx and LibreOffice 26.8 (round trips give
  identical text and page counts; charts, diagrams, shapes, equations, citations and merged documents render). No
  other desktop editor was available for testing.
- Grammar checking is rule-based (repeated words, spacing and punctuation, capitalisation, a/an, commonly confused
  words, wordy phrases, very long sentences); it is deliberately conservative and English-only. Spelling uses the
  browser's spell checker.
- Editing the data of an imported chart detaches the spreadsheet copy embedded in the original file. Charts created
  in SCRIBE store their data inside the chart.
- Equations edited as linear text are rebuilt in the default math font; run-level formatting inside an edited
  equation is not kept.
- Tables break between rows across pages (header rows are not repeated). A single row taller than a page overflows
  into the page gap; pictures taller than a page are shown scaled to the page (the file keeps their real size).
- In multi-column sections, a paragraph, list or table moves to the next column whole instead of splitting between
  columns, and columns are not balanced at the end of a section (exported files flow and balance as usual).
- Editing a note keeps its paragraph and character formatting; rich formatting inside an edited note is simplified.
- PDF export uses the browser print engine with the first section's page size for all pages.
- Fields in the body show their cached value (PAGE/NUMPAGES are live in headers and footers).
- Track changes record edits inside a paragraph; structural edits across paragraphs aren't tracked.
- Collaboration uses HTTP polling (~1.2 s) rather than WebSockets, because Vercel Functions don't hold sockets.
- Images are limited to 4 MB each (one serverless response); original DOCX files up to 100 MB are stored in chunks.
- Media and originals live in Postgres; on Neon's free tier the database is limited to 0.5 GB.
