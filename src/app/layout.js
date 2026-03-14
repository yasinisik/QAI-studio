import './globals.css'

export const metadata = {
  title: 'QA Studio',
  description: 'AI destekli test mühendisliği platformu',
  manifest: '/manifest.json',
}

export default function RootLayout({ children }) {
  return (
    <html lang="tr">
      <head>
        <meta name="theme-color" content="#0a1020" />
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1" />
        <meta name="apple-mobile-web-app-capable" content="yes" />
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
        <meta name="apple-mobile-web-app-title" content="QA Studio" />
        <link rel="manifest" href="/manifest.json" />
      </head>
      <body style={{ margin: 0, background: '#060a12', color: '#d4e0f0' }}>
        {children}
        <script dangerouslySetInnerHTML={{ __html: `
          if ('serviceWorker' in navigator) {
            window.addEventListener('load', () => {
              navigator.serviceWorker.register('/sw.js')
            })
          }
        `}} />
      </body>
    </html>
  )
}
