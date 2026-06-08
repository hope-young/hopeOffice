// One-off probe to see what vite returns for the icon URL.
import https from 'node:https'

const urls = [
  '/assets/icon-32.png',
  '/assets/icon-16.png',
  '/src/taskpane/index.html',
  '/',
]

for (const u of urls) {
  await new Promise((resolve) => {
    https
      .get(
        { host: 'localhost', port: 3721, path: u, rejectUnauthorized: false },
        (res) => {
          let body = ''
          res.on('data', (c) => (body += c))
          res.on('end', () => {
            console.log(
              u,
              '->',
              res.statusCode,
              '[' + (res.headers['content-type'] || '?') + ']',
              body.length + 'B',
              body.slice(0, 80).replace(/\n/g, '\\n'),
            )
            resolve()
          })
        },
      )
      .on('error', (e) => {
        console.log(u, 'ERR', e.message)
        resolve()
      })
  })
}
