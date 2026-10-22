(() => {
    const storageKey = 'attractive-theme-mode';
    const readTheme = () => {
        try { return localStorage.getItem(storageKey) === 'dark' ? 'dark' : 'light'; }
        catch (_) { return 'light'; }
    };
    const applyTheme = mode => {
        const theme = mode === 'dark' ? 'dark' : 'light';
        document.documentElement.dataset.theme = theme;
        try { localStorage.setItem(storageKey, theme); } catch (_) {}
        const isEnglish = document.documentElement.lang.toLowerCase().startsWith('en');
        document.querySelectorAll('.site-theme-toggle').forEach(button => {
            const dark = theme === 'dark';
            button.setAttribute('aria-pressed', String(dark));
            const label = dark
                ? (isEnglish ? 'Switch to light mode' : 'تفعيل الوضع الفاتح')
                : (isEnglish ? 'Switch to dark mode' : 'تفعيل الوضع الداكن');
            button.setAttribute('aria-label', label);
            button.setAttribute('title', label);
            button.innerHTML = `<span class="material-symbols-outlined" aria-hidden="true">${dark ? 'light_mode' : 'dark_mode'}</span>`;
        });
    };

    applyTheme(readTheme());
    const makeButton = () => {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'site-theme-toggle';
        button.addEventListener('click', () => applyTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'));
        return button;
    };

    const install = () => {
        const desktopActions = document.querySelector('nav.apple-nav .glass-nav-content > div:last-child');
        const mobileActions = document.querySelector('nav.apple-nav .mobile-nav-controls');
        const adminActions = document.querySelector('#adminApp > header > div:last-child');
        const authGate = document.getElementById('authGate');
        if (desktopActions) desktopActions.appendChild(makeButton());
        if (mobileActions) mobileActions.appendChild(makeButton());
        if (adminActions) {
            adminActions.appendChild(makeButton());
            if (authGate) {
                const gateButton = makeButton();
                gateButton.style.cssText = 'position:absolute;inset-inline-end:16px;top:16px;z-index:2';
                authGate.appendChild(gateButton);
            }
        }
        if (!desktopActions && !mobileActions && !adminActions) {
            const button = makeButton();
            button.style.cssText = 'position:fixed;inset-inline-end:20px;top:20px;z-index:9999;box-shadow:0 8px 24px rgba(0,0,0,.14)';
            document.body.appendChild(button);
        }
        applyTheme(readTheme());
    };
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true });
    else install();
})();
