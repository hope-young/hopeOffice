import { defineConfig, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { homedir } from 'node:os'

const __dirname = dirname(fileURLToPath(import.meta.url))

/**
 * Locate the dev cert produced by `npx office-addin-dev-certs install`.
 * It lives under the user profile and the filename is stable.
 */
function devCertPaths() {
  const root = join(homedir(), '.office-addin-dev-certs')
  return {
    key: join(root, 'localhost.key'),
    cert: join(root, 'localhost.crt'),
  }
}

function loadDevCert() {
  const { key, cert } = devCertPaths()
  if (!existsSync(key) || !existsSync(cert)) return undefined
  return { key: readFileSync(key), cert: readFileSync(cert) }
}

export default defineConfig(({ mode }) => {
  // Ensure dev certs exist on first run; if not, instruct the user to install.
  // We don't auto-install because cert trust is a user-level decision.
  const https = loadDevCert()
  if (mode === 'development' && !https) {
    console.warn(
      '[hope-office] HTTPS dev cert not found at ~/.office-addin-dev-certs/.\n' +
        '            Run `npx office-addin-dev-certs install` once, then retry `npm run dev`.\n' +
        '            Office Add-ins will refuse to load an HTTP or self-signed page.',
    )
  }

  return {
    plugins: [
      react(),
      tailwindcss(),
      ...(mode === 'production' ? [manifestRewritePlugin()] : []),
    ],

    server: {
      port: 3721,
      strictPort: true,
      host: 'localhost',
      // `https: undefined` lets Vite generate its own self-signed cert when
      // ours is missing — useful for `vite build` smoke tests in CI, even
      // though Office itself will reject it on sideload.
      https,
      hmr: https
        ? { protocol: 'wss', host: 'localhost', port: 3721 }
        : undefined,
    },

    preview: {
      port: 3721,
      strictPort: true,
    },

    build: {
      outDir: 'dist',
      sourcemap: true,
      target: 'es2022',
      rollupOptions: {
        // Two HTML entries, both referenced from runtime code:
        //  - taskpane/index.html  — the manifest's SourceLocation
        //  - taskpane/executor/iframe.html — the sandbox iframe; loaded
        //    at runtime by src/taskpane/executor/sandbox.ts as
        //    /src/taskpane/executor/iframe.html
        // Build output preserves the same relative paths under dist/.
        input: {
          taskpane: resolve(__dirname, 'src/taskpane/index.html'),
          'taskpane/executor/iframe': resolve(
            __dirname,
            'src/taskpane/executor/iframe.html',
          ),
        },
      },
    },

    resolve: {
      alias: {
        '@core': new URL('./src/core', import.meta.url).pathname,
        '@taskpane': new URL('./src/taskpane', import.meta.url).pathname,
      },
    },

    // Office Add-ins have CSP and no Node globals; we use import.meta.env in the
    // browser code and never read process.env there.
    define: {
      __DEV__: JSON.stringify(mode !== 'production'),
    },

    // GH Pages project-page base path. Asset URLs in dist/taskpane-*.js
    // and dist/assets/* will be prefixed with `/hopeOffice/`. The dev
    // server keeps `/` because Vite's HMR client expects relative URLs.
    base: mode === 'production' ? '/hopeOffice/' : '/',
  }
})

/**
 * Production manifest rewriter.
 *
 * The committed `manifest.xml` at the project root uses
 * `https://localhost:3721/...` everywhere — that URL is what the
 * dev server serves. For a production build we copy the manifest
 * into `dist/manifest.xml` and rewrite every reference to point
 * at the GH Pages project page. The installer ships the rewritten
 * copy.
 *
 * We rewrite the WHOLE string `https://localhost:3721` rather than
 * per-attribute so we don't have to track Office's evolving list
 * of URL-bearing elements (IconUrl, AppDomain, SourceLocation,
 * bt:Url, bt:Image, …). The dev URL is unique enough that a
 * global swap is safe.
 */
function manifestRewritePlugin(): Plugin {
  const PROD_BASE = process.env.SOURCE_URL ??
    'https://hope-young.github.io/hopeOffice'
  return {
    name: 'hope-office:rewrite-manifest',
    apply: 'build',
    closeBundle() {
      const src = resolve(__dirname, 'manifest.xml')
      const dst = resolve(__dirname, 'dist/manifest.xml')
      const xml = readFileSync(src, 'utf-8')
      const rewritten = xml.replaceAll(
        'https://localhost:3721',
        PROD_BASE,
      )
      writeFileSync(dst, rewritten, 'utf-8')
      console.warn(`[hope-office] wrote ${dst} (base: ${PROD_BASE})`)
    },
  }
}
