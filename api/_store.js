// Reads the public store content blob ("attractive-store-data.json") and merges it with the built-in defaults,
// exactly the way the browser does (a section the admin never saved falls back to the defaults).
// Files starting with "_" are NOT exposed as routes by Vercel.
import { head, put } from '@vercel/blob';
import DEFAULTS from './_defaults.js';

export const STORE_PATH = 'attractive-store-data.json';
const CLEAR_PRODUCTS_MIGRATION_KEY = 'empty_store_catalog_v1';

// -> { data, etag }   (data = {} when nothing was saved yet)
export async function readStoreRaw(attempt = 0) {
    const info = await head(STORE_PATH).catch(() => null);
    if (!info) {
        const initial = { products: [], _catalogMigrations: { [CLEAR_PRODUCTS_MIGRATION_KEY]: true } };
        try {
            const created = await put(STORE_PATH, JSON.stringify(initial), {
                access: 'public', addRandomSuffix: false, allowOverwrite: false,
                contentType: 'application/json', cacheControlMaxAge: 60
            });
            return { data: initial, etag: created.etag || null };
        } catch (error) {
            // A concurrent first write won the create; re-read it before applying the migration.
            if (attempt < 4 && /already exists|exists|precondition/i.test(String(error && error.message))) return readStoreRaw(attempt + 1);
            throw error;
        }
    }
    const r = await fetch(info.url + (info.url.includes('?') ? '&' : '?') + 't=' + Date.now(), { cache: 'no-store' });
    if (!r.ok) throw new Error('store read failed ' + r.status);
    const data = (await r.json()) || {};
    const migrations = data._catalogMigrations && typeof data._catalogMigrations === 'object' ? data._catalogMigrations : {};
    if (!migrations[CLEAR_PRODUCTS_MIGRATION_KEY]) {
        const migrated = Object.assign({}, data, {
            products: [],
            _catalogMigrations: Object.assign({}, migrations, { [CLEAR_PRODUCTS_MIGRATION_KEY]: true })
        });
        try {
            const saved = await put(STORE_PATH, JSON.stringify(migrated), {
                access: 'public', addRandomSuffix: false, allowOverwrite: true,
                contentType: 'application/json', cacheControlMaxAge: 60,
                ...(info.etag ? { ifMatch: info.etag } : {})
            });
            return { data: migrated, etag: saved.etag || info.etag };
        } catch (error) {
            // Never report a deletion as complete unless the Blob write succeeded.
            // If another save won the ETag race, reload and apply the migration to that version.
            if (attempt < 4 && (error && error.name === 'BlobPreconditionFailedError' || /precondition/i.test(String(error && error.message)))) return readStoreRaw(attempt + 1);
            throw error;
        }
    }
    return { data, etag: info.etag };
}

// -> { products, connectors, cables, shipping, coupons, settings }  (never null, never contains adminPassHash)
export async function getCatalog() {
    let stored = {};
    try { stored = (await readStoreRaw()).data || {}; } catch (e) { stored = {}; }
    const list = (k) => (Array.isArray(stored[k]) ? stored[k] : DEFAULTS[k]);
    const settings = Object.assign({}, DEFAULTS.settings, stored.settings && typeof stored.settings === 'object' ? stored.settings : {});
    delete settings.adminPassHash;
    return {
        products: list('products'),
        connectors: list('connectors'),
        cables: list('cables'),
        shipping: list('shipping').map(z => Object.assign({ fee: 50, enabled: true }, z)),
        coupons: list('coupons'),
        settings
    };
}
