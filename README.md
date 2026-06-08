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

# Run unit tests
npm test
```

After `npm run dev`, open Word → `Insert` → `My Add-ins` → `Manage My Add-ins` → `Upload My Add-in` → pick `manifest.xml`.

The task pane lives at `https://localhost:3721/src/taskpane/index.html`. The Custom Tab "hope-Office" should appear in the ribbon with three groups: **Chat**, **History**, **Settings**.

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
