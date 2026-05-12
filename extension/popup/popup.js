/*
 * popup.js
 * -----------------------------------------------------------------
 * Forwards toolbar-popup button clicks to the active tab's content
 * script via chrome.runtime messaging.
 *
 * The content script lives at content/main.js and exposes a
 * COMMAND_HANDLERS map behind a `feg.command` message type.
 * -----------------------------------------------------------------
 */
(function () {
    'use strict';

    const HOST_MATCH = /^https:\/\/portal\.ul\.com\/Project\/Details\//i;

    const statusEl = document.getElementById('feg-pop-status');
    const buttons  = Array.from(document.querySelectorAll('button[data-cmd]'));

    function setStatus(text, kind) {
        statusEl.textContent = text || '';
        statusEl.classList.remove('ok', 'err');
        if (kind) statusEl.classList.add(kind);
    }

    function setBusy(busy) {
        for (const b of buttons) b.disabled = !!busy;
    }

    async function getActiveTab() {
        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        return tabs && tabs[0];
    }

    async function sendCommand(command) {
        setStatus('');
        const tab = await getActiveTab();
        if (!tab || !tab.id) {
            setStatus('No active tab.', 'err');
            return;
        }
        if (!tab.url || !HOST_MATCH.test(tab.url)) {
            setStatus('Open a Project Details page first.', 'err');
            return;
        }
        setBusy(true);
        try {
            const res = await chrome.tabs.sendMessage(tab.id, {
                type: 'feg.command',
                command,
            });
            if (res && res.ok === false) {
                setStatus(res.error || 'Command failed.', 'err');
            } else {
                setStatus('Done.', 'ok');
            }
        } catch (e) {
            // Most common cause: the content script hasn't loaded yet
            // (e.g. user opened the popup before the page finished
            // loading, or the page doesn't match the manifest pattern).
            const msg = (e && e.message) ? e.message : String(e);
            if (/Receiving end does not exist/i.test(msg)) {
                setStatus('Reload the page and try again.', 'err');
            } else {
                setStatus(msg, 'err');
            }
        } finally {
            setBusy(false);
        }
    }

    for (const btn of buttons) {
        btn.addEventListener('click', () => {
            const cmd = btn.getAttribute('data-cmd');
            if (cmd) sendCommand(cmd);
        });
    }
})();
