# hope-Office

AI-powered assistant for Microsoft Word, Excel, and PowerPoint.

## Status

Early development. The project is being built from scratch. See [`SPEC.md`](./SPEC.md) for the architecture and roadmap.

## Goals

- Local-first — runs entirely in the user's browser as an Office task pane
- BYO API key — keys are stored on the user's machine, traffic goes only to the provider they pick
- Multi-provider — Anthropic, OpenAI, and any OpenAI-compatible endpoint (Ollama, LM Studio, …)
- Multi-host — single add-in works in Word, Excel, and PowerPoint
- Open source under MIT

## Development

Prerequisites: Node 20+, Microsoft 365 (Word/Excel/PowerPoint) for sideload testing.

```bash
# Install deps
npm install

# One-time: install the HTTPS dev cert (Office Add-ins require HTTPS)
npx office-addin-dev-certs install

# Run the dev server (HTTPS on https://localhost:3721)
npm run dev

# In another terminal: sideload the add-in into Word
npm run sideload

# Type-check + production build
npm run build

# Production build: rewrites `manifest.xml`'s `SourceLocation`
# to point at GH Pages (or whatever SOURCE_URL is set to), then
# bakes the relative-asset base `/hopeOffice/` into the bundle.
SOURCE_URL=https://your.domain.example npm run build:prod

# Deploy the production build to the gh-pages branch. The
# `predeploy` script runs `build:prod` first.
npm run deploy

# Install the production add-in (Windows, per-user, no admin)
pwsh installer/install.ps1 -Production

# Run unit tests
npm test
```

After `npm run dev`, open Word → `Insert` → `My Add-ins` → `Manage My Add-ins` → `Upload My Add-in` → pick `manifest.xml`.

The task pane lives at `https://localhost:3721/src/taskpane/index.html` in dev, or at the URL baked into `dist/manifest.xml` (default: `https://hope-young.github.io/hopeOffice/...`) in production. The Custom Tab "hope-Office" should appear in the ribbon with three groups: **Chat**, **History**, **Settings**.

### Production deploy via GitHub Pages

The deploy target is a GH Pages project page, so the asset base is
`/hopeOffice/` (set automatically by `vite.config.ts` in production
mode). To switch to a different host (Cloudflare Pages, Vercel, an
internal nginx), set `SOURCE_URL` at build time and adjust the
`base` option accordingly.

## Layout

```
src/
├── core/        pure logic, no React, no Office
└── taskpane/    React app, DOM-only
manifest.xml     dev (localhost:3721)
SPEC.md         architecture + roadmap
SPEC_DETAILS.md implementation details (types, protocols, algorithms)
```

## License

MIT — see [`LICENSE`](./LICENSE).
