<script lang="ts">
    import Generator from "$lib/components/Generator.svelte";
    import Extractor from "$lib/components/Extractor.svelte";
    import ThemeSwitcher from "$lib/components/ThemeSwitcher.svelte";
    import Particles from "$lib/components/Particles.svelte";

    let activeTab: "generator" | "extractor" = "generator";
</script>

<svelte:head>
    <title>Document Forge - The Ultimate Tool</title>
</svelte:head>

<svelte:window on:dragover|preventDefault on:drop|preventDefault />

<Particles />
<ThemeSwitcher />

<div
    class="min-h-screen bg-surface-950 font-body p-4 md:p-8 flex flex-col items-center justify-start selection:bg-primary-500/30 relative z-10"
>
    <!-- Tab Switcher -->
    <div
        class="glass-pill flex p-1.5 w-full max-w-md mx-auto mb-8 relative z-20"
    >
        <button
            class={`flex-1 py-3 px-6 rounded-full text-sm font-extrabold tracking-widest font-display transition-all duration-500 ease-out ${activeTab === "generator" ? "bg-gradient-to-r from-primary-500 to-primary-600 text-white shadow-[0_0_20px_rgba(244,63,94,0.3)]" : "text-surface-400 hover:text-surface-200"}`}
            on:click={() => (activeTab = "generator")}
        >
            Template Engine
        </button>
        <button
            class={`flex-1 py-3 px-6 rounded-full text-sm font-extrabold tracking-widest font-display transition-all duration-500 ease-out ${activeTab === "extractor" ? "bg-gradient-to-r from-purple-500 to-indigo-500 text-white shadow-[0_0_20px_rgba(168,85,247,0.3)]" : "text-surface-400 hover:text-surface-200"}`}
            on:click={() => (activeTab = "extractor")}
        >
            Data Extractor
        </button>
    </div>

    <!-- Active View -->
    <div class="w-full transition-all duration-500">
        {#if activeTab === "generator"}
            <Generator />
        {:else}
            <Extractor />
        {/if}
    </div>
</div>
