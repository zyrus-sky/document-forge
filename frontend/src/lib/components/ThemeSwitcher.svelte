<script lang="ts">
    import { onMount } from "svelte";

    // --- Theme Definitions ---
    const themes: Record<
        string,
        {
            label: string;
            swatch: string;
            primary: Record<string, string>;
            surface: Record<string, string>;
            body: string;
            text: string;
            glow: string;
        }
    > = {
        crimson: {
            label: "Crimson",
            swatch: "#f43f5e",
            glow: "rgba(244, 63, 94, 0.04)",
            primary: {
                "50": "#fff1f2",
                "100": "#ffe4e6",
                "200": "#fecdd3",
                "300": "#fda4af",
                "400": "#fb7185",
                "500": "#f43f5e",
                "600": "#e11d48",
                "700": "#be123c",
                "800": "#9f1239",
                "900": "#881337",
                "950": "#4c0519",
            },
            surface: {
                "50": "#f8fafc",
                "100": "#f1f5f9",
                "200": "#e2e8f0",
                "300": "#cbd5e1",
                "400": "#94a3b8",
                "500": "#64748b",
                "600": "#475569",
                "700": "#334155",
                "800": "#1e293b",
                "900": "#0f172a",
                "950": "#020617",
            },
            body: "#020617",
            text: "#f8fafc",
        },
        neon: {
            label: "Neon",
            swatch: "#06b6d4",
            glow: "rgba(6, 182, 212, 0.05)",
            primary: {
                "50": "#ecfeff",
                "100": "#cffafe",
                "200": "#a5f3fc",
                "300": "#67e8f9",
                "400": "#22d3ee",
                "500": "#06b6d4",
                "600": "#0891b2",
                "700": "#0e7490",
                "800": "#155e75",
                "900": "#164e63",
                "950": "#083344",
            },
            surface: {
                "50": "#f9fafb",
                "100": "#f3f4f6",
                "200": "#e5e7eb",
                "300": "#d1d5db",
                "400": "#9ca3af",
                "500": "#6b7280",
                "600": "#4b5563",
                "700": "#374151",
                "800": "#1f2937",
                "900": "#111827",
                "950": "#030712",
            },
            body: "#030712",
            text: "#f9fafb",
        },
        ocean: {
            label: "Ocean",
            swatch: "#3b82f6",
            glow: "rgba(59, 130, 246, 0.05)",
            primary: {
                "50": "#eff6ff",
                "100": "#dbeafe",
                "200": "#bfdbfe",
                "300": "#93c5fd",
                "400": "#60a5fa",
                "500": "#3b82f6",
                "600": "#2563eb",
                "700": "#1d4ed8",
                "800": "#1e40af",
                "900": "#1e3a8a",
                "950": "#172554",
            },
            surface: {
                "50": "#f0f4f8",
                "100": "#d9e2ec",
                "200": "#bcccdc",
                "300": "#9fb3c8",
                "400": "#829ab1",
                "500": "#627d98",
                "600": "#486581",
                "700": "#334e68",
                "800": "#243b53",
                "900": "#102a43",
                "950": "#0c1929",
            },
            body: "#0c1929",
            text: "#f0f4f8",
        },
        emerald: {
            label: "Emerald",
            swatch: "#10b981",
            glow: "rgba(16, 185, 129, 0.05)",
            primary: {
                "50": "#ecfdf5",
                "100": "#d1fae5",
                "200": "#a7f3d0",
                "300": "#6ee7b7",
                "400": "#34d399",
                "500": "#10b981",
                "600": "#059669",
                "700": "#047857",
                "800": "#065f46",
                "900": "#064e3b",
                "950": "#022c22",
            },
            surface: {
                "50": "#f0fdf4",
                "100": "#dcfce7",
                "200": "#c2d9cd",
                "300": "#a3b8b0",
                "400": "#7a9485",
                "500": "#5a7568",
                "600": "#3d5a4a",
                "700": "#264334",
                "800": "#162e22",
                "900": "#0d1f16",
                "950": "#0a1a15",
            },
            body: "#0a1a15",
            text: "#f0fdf4",
        },
        amethyst: {
            label: "Amethyst",
            swatch: "#8b5cf6",
            glow: "rgba(139, 92, 246, 0.05)",
            primary: {
                "50": "#f5f3ff",
                "100": "#ede9fe",
                "200": "#ddd6fe",
                "300": "#c4b5fd",
                "400": "#a78bfa",
                "500": "#8b5cf6",
                "600": "#7c3aed",
                "700": "#6d28d9",
                "800": "#5b21b6",
                "900": "#4c1d95",
                "950": "#2e1065",
            },
            surface: {
                "50": "#faf5ff",
                "100": "#f3e8ff",
                "200": "#d8c5e6",
                "300": "#b8a0c9",
                "400": "#8b6fa8",
                "500": "#6b5080",
                "600": "#4e3a60",
                "700": "#362848",
                "800": "#221a33",
                "900": "#150f22",
                "950": "#0d0b1a",
            },
            body: "#0d0b1a",
            text: "#faf5ff",
        },
    };

    let currentTheme = $state("crimson");
    let isOpen = $state(false);

    function applyTheme(name: string) {
        const theme = themes[name];
        if (!theme) return;

        const root = document.documentElement;
        for (const [shade, value] of Object.entries(theme.primary)) {
            root.style.setProperty(`--color-primary-${shade}`, value);
        }
        for (const [shade, value] of Object.entries(theme.surface)) {
            root.style.setProperty(`--color-surface-${shade}`, value);
        }
        root.style.setProperty("--color-body", theme.body);
        root.style.setProperty("--color-text", theme.text);

        // Update background gradients
        document.body.style.backgroundColor = theme.body;
        document.body.style.backgroundImage = `
            radial-gradient(circle at 15% 50%, ${theme.glow}, transparent 50%),
            radial-gradient(circle at 85% 30%, ${theme.glow}, transparent 50%)
        `;
        document.body.style.color = theme.text;

        currentTheme = name;
        localStorage.setItem("docforge_theme", name);
        isOpen = false;
    }

    onMount(() => {
        const saved = localStorage.getItem("docforge_theme");
        if (saved && themes[saved]) {
            applyTheme(saved);
        }
    });

    function handleClickOutside(event: MouseEvent) {
        const target = event.target as HTMLElement;
        if (!target.closest(".theme-switcher")) {
            isOpen = false;
        }
    }
</script>

<svelte:window on:click={handleClickOutside} />

<div class="theme-switcher fixed top-4 right-4 z-50">
    <!-- Toggle Button -->
    <button
        class="group relative w-10 h-10 rounded-xl bg-surface-900/80 backdrop-blur-xl border border-surface-700/50
               hover:border-primary-500/50 transition-all duration-300 flex items-center justify-center
               shadow-lg hover:shadow-primary-500/20"
        on:click|stopPropagation={() => (isOpen = !isOpen)}
        title="Switch Theme"
    >
        <div
            class="w-5 h-5 rounded-full transition-all duration-300 group-hover:scale-110"
            style="background: {themes[currentTheme]
                .swatch}; box-shadow: 0 0 12px {themes[currentTheme].swatch}40;"
        ></div>
    </button>

    <!-- Dropdown -->
    {#if isOpen}
        <div
            class="absolute top-12 right-0 bg-surface-900/95 backdrop-blur-2xl border border-surface-700/50
                   rounded-2xl p-2 shadow-2xl min-w-[180px]
                   animate-in fade-in slide-in-from-top-2 duration-200"
        >
            <div
                class="text-[10px] uppercase tracking-[0.2em] text-surface-500 font-bold px-3 py-1.5 mb-1"
            >
                Theme
            </div>

            {#each Object.entries(themes) as [key, theme]}
                <button
                    class="w-full flex items-center gap-3 px-3 py-2 rounded-xl text-sm font-medium
                           transition-all duration-200
                           {currentTheme === key
                        ? 'bg-surface-800 text-white'
                        : 'text-surface-300 hover:bg-surface-800/60 hover:text-white'}"
                    on:click|stopPropagation={() => applyTheme(key)}
                >
                    <div
                        class="w-4 h-4 rounded-full shrink-0 transition-transform duration-200
                               {currentTheme === key
                            ? 'scale-125 ring-2 ring-white/30'
                            : ''}"
                        style="background: {theme.swatch}; box-shadow: 0 0 8px {theme.swatch}50;"
                    ></div>
                    <span>{theme.label}</span>
                    {#if currentTheme === key}
                        <span class="ml-auto text-xs text-primary-400">✓</span>
                    {/if}
                </button>
            {/each}
        </div>
    {/if}
</div>
