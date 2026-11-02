/**
 * MSL Checker - Modern Interactive Script
 */

document.addEventListener("DOMContentLoaded", async () => {
    // 1. 動的要素とテーマの初期化
    initTheme();
    initViewTabs();
    initReleaseFilter();
    initFaqSearch();
    initScrollSpy();

    // 2. Material Symbols / Webフォントのロード完了を待機 (FOUT防止)
    await waitForResourcesReady();

    // 3. 準備が完了してから画面をフェードイン表示
    document.body.classList.remove("is-loading");

    // 4. 最新リリースの情報を非同期取得 (完了時にダウンロードボタンを更新)
    initReleasesData();
});

/**
 * Webフォント (Material Symbols & Noto Sans JP) の読み込み完了を待つ。
 * ネットワーク遅延やオフライン時でも画面がブロックされないよう最大1200msでフォールバック。
 */
function waitForResourcesReady() {
    const fontPromise = document.fonts ? document.fonts.ready : Promise.resolve();
    const timeoutPromise = new Promise(resolve => setTimeout(resolve, 1200));
    return Promise.race([fontPromise, timeoutPromise]);
}

/* --------------------------------------------------------------------------
   1. Theme Management (Light / Dark Mode)
   -------------------------------------------------------------------------- */
function initTheme() {
    const themeToggle = document.getElementById("theme-toggle");
    const currentTheme = localStorage.getItem("msl_theme") ||
        (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");

    document.documentElement.setAttribute("data-theme", currentTheme);
    updateThemeIcon(currentTheme);

    if (themeToggle) {
        themeToggle.addEventListener("click", () => {
            const activeTheme = document.documentElement.getAttribute("data-theme");
            const newTheme = activeTheme === "dark" ? "light" : "dark";
            document.documentElement.setAttribute("data-theme", newTheme);
            localStorage.setItem("msl_theme", newTheme);
            updateThemeIcon(newTheme);
        });
    }
}

function updateThemeIcon(theme) {
    const iconSpan = document.getElementById("theme-icon");
    if (iconSpan) {
        iconSpan.textContent = theme === "dark" ? "light_mode" : "dark_mode";
    }
}

/* --------------------------------------------------------------------------
   2. Main View Tabs (Manual vs. Release History)
   -------------------------------------------------------------------------- */
function initViewTabs() {
    const tabBtns = document.querySelectorAll(".tab-btn");
    const panels = document.querySelectorAll(".view-panel");

    function switchTab(targetId, updateHash = true) {
        tabBtns.forEach(btn => {
            const isTarget = btn.getAttribute("data-target") === targetId;
            btn.classList.toggle("active", isTarget);
            btn.setAttribute("aria-selected", isTarget ? "true" : "false");
        });

        panels.forEach(panel => {
            const isTarget = panel.id === targetId;
            panel.classList.toggle("active", isTarget);
        });

        if (updateHash) {
            history.replaceState(null, "", `#${targetId}`);
        }
    }

    tabBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            const targetId = btn.getAttribute("data-target");
            switchTab(targetId);
        });
    });

    // Check URL hash on page load
    const hash = window.location.hash.replace("#", "");
    if (hash === "releases" || hash === "manual") {
        switchTab(hash, false);
    } else if (hash && document.getElementById(hash)) {
        // If hash refers to a section in manual, ensure manual is active
        switchTab("manual", false);
    }
}

/* --------------------------------------------------------------------------
   3. Dynamic GitHub Release Assets Fetching
   -------------------------------------------------------------------------- */
async function initReleasesData() {
    const REPO = "sam-hat-86/MSL_Checker";
    const API_URL = `https://api.github.com/repos/${REPO}/releases/latest`;
    const downloadBtn = document.getElementById("main-download-btn");
    const versionBadge = document.getElementById("latest-version-badge");

    try {
        const res = await fetch(API_URL);
        if (!res.ok) throw new Error(`HTTP error ${res.status}`);
        const data = await res.json();

        const tagName = data.tag_name || "v17.0";
        if (versionBadge) {
            versionBadge.textContent = `最新: ${tagName}`;
        }

        // Search for primary release exe (excluding _dev)
        const exeAsset = data.assets && data.assets.find(a =>
            a.name.endsWith(".exe") && !a.name.includes("_dev")
        );

        if (downloadBtn && exeAsset) {
            downloadBtn.href = exeAsset.browser_download_url;
            downloadBtn.innerHTML = `<span class="material-symbols-outlined">bolt</span> ${exeAsset.name} をダウンロード`;
        } else if (downloadBtn) {
            downloadBtn.href = data.html_url || `https://github.com/${REPO}/releases/latest`;
        }
    } catch (e) {
        console.warn("Could not fetch latest release info from GitHub:", e);
        if (versionBadge) versionBadge.textContent = "最新: v17.0";
        if (downloadBtn) {
            downloadBtn.href = `https://github.com/${REPO}/releases/latest`;
            downloadBtn.innerHTML = `<span class="material-symbols-outlined">bolt</span> 最新版 MSL Checker をダウンロード`;
        }
    }
}

/* --------------------------------------------------------------------------
   4. Release Notes Filters & Live Search
   -------------------------------------------------------------------------- */
function initReleaseFilter() {
    const pillBtns = document.querySelectorAll(".filter-pills .pill-btn");
    const searchInput = document.getElementById("release-search");
    const releaseCards = document.querySelectorAll(".timeline-item");

    function applyFilters() {
        const activePill = document.querySelector(".filter-pills .pill-btn.active");
        const category = activePill ? activePill.getAttribute("data-filter") : "all";
        const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

        releaseCards.forEach(card => {
            const cardGen = card.getAttribute("data-generation");
            const textContent = card.textContent.toLowerCase();

            const matchesCategory = (category === "all" || cardGen === category);
            const matchesQuery = !query || textContent.includes(query);

            if (matchesCategory && matchesQuery) {
                card.style.display = "";
            } else {
                card.style.display = "none";
            }
        });
    }

    pillBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            pillBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            applyFilters();
        });
    });

    if (searchInput) {
        searchInput.addEventListener("input", applyFilters);
    }
}

/* --------------------------------------------------------------------------
   5. FAQ Interactive Search
   -------------------------------------------------------------------------- */
function initFaqSearch() {
    const faqInput = document.getElementById("faq-search");
    const faqItems = document.querySelectorAll(".accordion-item");

    if (!faqInput) return;

    faqInput.addEventListener("input", () => {
        const query = faqInput.value.trim().toLowerCase();

        faqItems.forEach(item => {
            const text = item.textContent.toLowerCase();
            if (!query || text.includes(query)) {
                item.style.display = "";
                if (query) item.open = true; // Auto open when searching
            } else {
                item.style.display = "none";
            }
        });
    });
}

/* --------------------------------------------------------------------------
   6. ScrollSpy for Manual TOC
   -------------------------------------------------------------------------- */
function initScrollSpy() {
    const sections = document.querySelectorAll(".manual-content .content-card[id]");
    const tocLinks = document.querySelectorAll(".sidebar-toc .toc-list a");

    if (sections.length === 0 || tocLinks.length === 0) return;

    const observer = new IntersectionObserver(entries => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const id = entry.target.id;
                tocLinks.forEach(link => {
                    const href = link.getAttribute("href").replace("#", "");
                    link.classList.toggle("active", href === id);
                });
            }
        });
    }, { rootMargin: "-20% 0px -70% 0px" });

    sections.forEach(sec => observer.observe(sec));
}
