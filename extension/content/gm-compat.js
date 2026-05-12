/*
 * gm-compat.js
 * -----------------------------------------------------------------
 * Tiny compatibility shim that lets the original Tampermonkey
 * userscript (`content/main.js`) run as a Chrome/Edge MV3 content
 * script without any further code changes to the GM_* call sites.
 *
 * It exposes:
 *   - GM_getValue / GM_setValue   → backed by chrome.storage.local
 *                                   (with a synchronous in-memory mirror
 *                                   so the existing sync getters keep
 *                                   working). The first read is awaited
 *                                   via window.__FEG_INIT__ before main.js
 *                                   touches any persisted setting.
 *   - GM_setClipboard             → navigator.clipboard.writeText
 *   - GM_registerMenuCommand      → no-op (the toolbar popup replaces it)
 *   - unsafeWindow                → window (content scripts can call
 *                                   window.showOpenFilePicker directly)
 * -----------------------------------------------------------------
 */
(function () {
    'use strict';

    // Keys we want to mirror synchronously. Keep this list in sync with
    // anything that flows through Storage / GM_getValue in main.js.
    const CACHED_KEYS = ['fegx.template.mode'];
    const cache = Object.create(null);

    const storage = (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local)
        ? chrome.storage.local
        : null;

    function readAll() {
        if (!storage) return Promise.resolve();
        return new Promise((resolve) => {
            try {
                storage.get(CACHED_KEYS, (items) => {
                    if (items && typeof items === 'object') {
                        for (const k of CACHED_KEYS) {
                            if (Object.prototype.hasOwnProperty.call(items, k)) {
                                cache[k] = items[k];
                            }
                        }
                    }
                    resolve();
                });
            } catch (_) {
                resolve();
            }
        });
    }

    // Expose the init promise so main.js can await it before reading.
    window.__FEG_INIT__ = readAll();

    // React to changes made elsewhere (e.g. another tab) so the cache
    // stays fresh for as long as the page is open.
    if (storage && chrome.storage.onChanged) {
        chrome.storage.onChanged.addListener((changes, areaName) => {
            if (areaName !== 'local') return;
            for (const k of Object.keys(changes)) {
                if (CACHED_KEYS.indexOf(k) === -1) continue;
                const c = changes[k];
                if (c && Object.prototype.hasOwnProperty.call(c, 'newValue')) {
                    cache[k] = c.newValue;
                }
            }
        });
    }

    window.GM_getValue = function (key, defaultValue) {
        const v = cache[key];
        return v === undefined ? defaultValue : v;
    };

    window.GM_setValue = function (key, value) {
        cache[key] = value;
        if (storage) {
            try { storage.set({ [key]: value }); } catch (_) { /* ignore */ }
        }
    };

    window.GM_setClipboard = function (text /*, type */) {
        const s = (text == null) ? '' : String(text);
        if (navigator.clipboard && navigator.clipboard.writeText) {
            // Best-effort; failures (e.g. document not focused) are
            // surfaced by the calling code's try/catch.
            return navigator.clipboard.writeText(s);
        }
        // Fallback for very old browsers — should not happen on a
        // Chromium-based extension target, but kept for robustness.
        try {
            const ta = document.createElement('textarea');
            ta.value = s;
            ta.setAttribute('readonly', '');
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
        } catch (_) { /* ignore */ }
        return undefined;
    };

    // The original userscript registered GM menu commands; in the
    // extension we expose them through the toolbar popup instead, so
    // this becomes a no-op.
    window.GM_registerMenuCommand = function () { /* handled by popup */ };

    // Content scripts in MV3 already have access to page-world DOM
    // APIs through `window`. There is no privileged unsafeWindow.
    if (typeof window.unsafeWindow === 'undefined') {
        try { window.unsafeWindow = window; } catch (_) { /* ignore */ }
    }
})();
