// Service Worker der digitalen Visitenkarte
// Bei jeder Änderung an den Dateien die Versionsnummer erhöhen, damit Handys die neue Fassung laden.
const CACHE = 'visitenkarte-v7';

const ASSETS = [
    './',
    'index.html',
    'manifest.webmanifest',
    'mmbbs_logo.svg',
    'nm.jpg',
    'nm_qr.png',
    'nm.vcf',
    'icons/favicon-32.png',
    'icons/favicon-48.png',
    'icons/icon-192.png',
    'icons/icon-512.png',
    'icons/apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE)
            // einzeln hinzufügen: fehlt eine Datei, bricht nicht die ganze Installation ab
            .then((cache) => Promise.allSettled(ASSETS.map((a) => cache.add(a))))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys()
            .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
            .then(() => self.clients.claim())
    );
});

// Cache zuerst (schnell, offline-fähig), im Hintergrund wird die neue Fassung nachgeladen
self.addEventListener('fetch', (event) => {
    const req = event.request;
    if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;

    event.respondWith(
        caches.match(req, { ignoreSearch: true }).then((cached) => {
            const network = fetch(req)
                .then((res) => {
                    if (res && res.ok) {
                        const copy = res.clone();
                        caches.open(CACHE).then((c) => c.put(req, copy));
                    }
                    return res;
                })
                .catch(() => cached || (req.mode === 'navigate' ? caches.match('index.html') : undefined));
            return cached || network;
        })
    );
});
