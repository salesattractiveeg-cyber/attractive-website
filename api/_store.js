// Reads the public store content blob ("attractive-store-data.json") and merges it with the built-in defaults,
// exactly the way the browser does (a section the admin never saved falls back to the defaults).
// Files starting with "_" are NOT exposed as routes by Vercel.
import { head, put } from '@vercel/blob';
import DEFAULTS from './_defaults.js';

export const STORE_PATH = 'attractive-store-data.json';
const QUOTE_MIGRATION_KEY = 'S03436_quote_products_v1';

// -> { data, etag }   (data = {} when nothing was saved yet)
export async function readStoreRaw() {
    const info = await head(STORE_PATH).catch(() => null);
    if (!info) return { data: {}, etag: null };
    const r = await fetch(info.url + (info.url.includes('?') ? '&' : '?') + 't=' + Date.now(), { cache: 'no-store' });
    if (!r.ok) throw new Error('store read failed ' + r.status);
    const data = (await r.json()) || {};
    const migrations = data._catalogMigrations && typeof data._catalogMigrations === 'object' ? data._catalogMigrations : {};
    const quoteDefaults = DEFAULTS.products.filter(product => product.quoteRef === 'S03436');
    if (!migrations[QUOTE_MIGRATION_KEY] && quoteDefaults.length) {
        const products = Array.isArray(data.products) ? data.products : DEFAULTS.products;
        const existingIds = new Set(products.map(product => String(product.id)));
        const missingQuoteProducts = quoteDefaults.filter(product => !existingIds.has(String(product.id)));
        const migrated = Object.assign({}, data, {
            products: [...products, ...missingQuoteProducts],
            _catalogMigrations: Object.assign({}, migrations, { [QUOTE_MIGRATION_KEY]: true })
        });
        try {
            const saved = await put(STORE_PATH, JSON.stringify(migrated), {
                access: 'public', addRandomSuffix: false, allowOverwrite: true,
                contentType: 'application/json', cacheControlMaxAge: 60,
                ...(info.etag ? { ifMatch: info.etag } : {})
            });
            return { data: migrated, etag: saved.etag || info.etag };
        } catch (error) {
            // A concurrent catalog save may win this migration attempt. Return
            // the merged view for this request; the next read retries safely.
            return { data: migrated, etag: info.etag };
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
