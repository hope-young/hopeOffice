import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { existsSync, readFileSync } from 'node:fs'
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
    plugins: [react(), tailwindcss()],

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
        // The Office manifest points to src/taskpane/index.html, so that is
        // the real entry. We keep the project root as the Vite root and
        // build into dist/ preserving the same relative path.
        input: {
          taskpane: resolve(__dirname, 'src/taskpane/index.html'),
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
  }
})
