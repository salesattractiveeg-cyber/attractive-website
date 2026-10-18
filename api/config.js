// GET /api/config -> public, non-secret runtime settings for the browser.
// OAuth client/app IDs are public identifiers; private provider secrets are never returned here.
import { readStoreRaw } from './_store.js';

export default function handler(req, res) {
    if (req.method !== 'GET') { res.setHeader('Allow', ['GET']); return res.status(405).end('Method not allowed'); }
    res.setHeader('Cache-Control', 'public, max-age=0, must-revalidate');
    res.setHeader('Vercel-CDN-Cache-Control', 'public, s-maxage=60, stale-while-revalidate=60');
    return readStoreRaw().catch(() => ({ data: {} })).then(({ data }) => {
        const fromStore = data && data.settings && typeof data.settings.googleClientId === 'string' ? data.settings.googleClientId.trim() : '';
        const fromEnv = String(process.env.GOOGLE_CLIENT_ID || '').trim();
        const valid = (id) => /^[A-Za-z0-9_-]+\.apps\.googleusercontent\.com$/.test(id);
        const facebookAppId = String(process.env.FACEBOOK_APP_ID || '').trim();
        return res.status(200).json({ googleClientId: valid(fromStore) ? fromStore : (valid(fromEnv) ? fromEnv : ''), facebookAppId: /^\d{5,24}$/.test(facebookAppId) ? facebookAppId : '' });
    });
}
