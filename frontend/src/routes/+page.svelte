<script lang="ts">
    import {
        FileCsv,
        FileDoc,
        UploadSimple,
        ArrowRight,
        Table,
        TextT,
        CheckCircle,
        Spinner,
        DownloadSimple,
        Faders,
    } from "phosphor-svelte";
    import { slide, scale, fade } from "svelte/transition";
    import { onMount } from "svelte";

    // State
    let step = 1;
    let csvFile: File | null = null;
    let docxFile: File | null = null;
    let isUploading = false;
    let sessionId = "";
    let errorMsg = "";

    // Metadata
    let csvHeaders: string[] = [];
    let docxPlaceholders: string[] = [];
    let totalRows = 0;

    // Mapping State
    type MappingType = "csv_column" | "custom_text" | "combined";
    interface MappingConfig {
        type: MappingType;
        value: string;
        customText: string;
        prefix: string;
        suffix: string;
    }
    let mappings: Record<string, MappingConfig> = {};

    // Generation State
    let options = { docx: true, pdf: true };
    let isGenerating = false;
    let generationComplete = false;
    let downloadUrl = "";

    // FastAPI Root
    const API_URL = "http://localhost:8000/api";

    function autoDetectMappings() {
        docxPlaceholders.forEach((ph) => {
            const cleanPh = ph
                .replace("#", "")
                .toLowerCase()
                .replace(/\s/g, "");
            let match = "";
            for (const header of csvHeaders) {
                const cleanHeader = header.toLowerCase().replace(/\s/g, "");
                if (
                    cleanHeader === cleanPh ||
                    (header === "P_ ADDRESS" && ph === "#P_ ADDRESS")
                ) {
                    match = header;
                    break;
                }
            }
            mappings[ph] = {
                type: match ? "csv_column" : "custom_text",
                value: match,
                customText: "",
                prefix: "",
                suffix: "",
            };
        });
    }

    async function handleUpload() {
        if (!csvFile || !docxFile) return;
        isUploading = true;
        errorMsg = "";

        const formData = new FormData();
        formData.append("csv_file", csvFile);
        formData.append("template_file", docxFile);

        try {
            // 1. Upload
            const upRes = await fetch(`${API_URL}/upload`, {
                method: "POST",
                body: formData,
            });
            if (!upRes.ok) throw new Error("Upload failed");
            const upData = await upRes.json();
            sessionId = upData.session_id;

            // 2. Get Metadata
            const metaRes = await fetch(
                `${API_URL}/metadata?session_id=${sessionId}`,
            );
            if (!metaRes.ok) throw new Error("Metadata extraction failed");
            const metaData = await metaRes.json();

            csvHeaders = metaData.csv_headers;
            docxPlaceholders = metaData.docx_placeholders;
            totalRows = metaData.total_rows;

            autoDetectMappings();
            step = 2;
        } catch (err: any) {
            errorMsg = err.message;
        } finally {
            isUploading = false;
        }
    }

    async function handleGenerate() {
        isGenerating = true;
        errorMsg = "";

        // Clean mapping payload for API
        const cleanMapping: Record<
            string,
            { type: string; value: string; prefix?: string; suffix?: string }
        > = {};
        for (const [key, conf] of Object.entries(mappings)) {
            cleanMapping[key] = {
                type: conf.type,
                value:
                    conf.type === "custom_text" ? conf.customText : conf.value,
                prefix: conf.prefix,
                suffix: conf.suffix,
            };
        }

        try {
            const response = await fetch(`${API_URL}/generate`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    session_id: sessionId,
                    mapping: cleanMapping,
                    generate_docx: options.docx,
                    generate_pdf: options.pdf,
                }),
            });

            if (!response.ok) throw new Error("Generation failed");

            // Download Blob
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            downloadUrl = url;
            generationComplete = true;
            step = 3;
        } catch (err: any) {
            errorMsg = err.message;
        } finally {
            isGenerating = false;
        }
    }

    // File Handlers
    function onCsvDrop(e: any) {
        csvFile = e.dataTransfer?.files[0] || e.target?.files[0];
    }
    function onDocxDrop(e: any) {
        docxFile = e.dataTransfer?.files[0] || e.target?.files[0];
    }
</script>

<svelte:window on:dragover|preventDefault on:drop|preventDefault />

<div
    class="min-h-screen bg-surface-950 font-sans p-4 md:p-8 flex items-center justify-center selection:bg-primary-500/30"
>
    <!-- Main Card container mapping Skeleton Labs variables directly via utility classes -->
    <div
        class="w-full max-w-5xl bg-surface-900 border border-surface-800 rounded-[2rem] shadow-2xl overflow-hidden flex flex-col relative"
    >
        <!-- Header -->
        <header
            class="p-6 md:px-10 border-b border-surface-800/80 bg-surface-900/50 backdrop-blur-md flex items-center justify-between z-10 sticky top-0"
        >
            <div class="flex items-center gap-4">
                <div
                    class="p-3 bg-gradient-to-br from-primary-500 to-primary-700 rounded-2xl shadow-[0_0_20px_rgba(244,63,94,0.3)] border border-primary-400/20"
                >
                    <Faders class="w-6 h-6 text-white" weight="duotone" />
                </div>
                <div>
                    <h1 class="text-2xl font-bold text-surface-50">
                        Document Forge
                    </h1>
                    <p
                        class="text-xs font-semibold tracking-wider text-primary-400/80 uppercase mt-0.5"
                    >
                        Advanced Template Engine
                    </p>
                </div>
            </div>

            <div class="flex items-center gap-2">
                {#each [1, 2, 3] as s}
                    <div class="flex items-center">
                        <div
                            class={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold transition-all duration-500 ${step >= s ? "bg-primary-500/20 text-primary-400 ring-1 ring-primary-500/50 shadow-[0_0_15px_rgba(244,63,94,0.2)]" : "bg-surface-800 text-surface-500"}`}
                        >
                            {step > s ? "✓" : s}
                        </div>
                        {#if s < 3}
                            <div
                                class={`w-8 h-0.5 mx-2 transition-colors duration-500 ${step > s ? "bg-primary-500/50" : "bg-surface-800"}`}
                            ></div>
                        {/if}
                    </div>
                {/each}
            </div>
        </header>

        <main class="flex-1 p-6 md:p-10 relative overflow-y-auto max-h-[75vh]">
            {#if errorMsg}
                <div
                    transition:slide
                    class="mb-8 p-4 bg-red-500/10 border border-red-500/30 rounded-xl flex gap-4 text-red-500 items-start"
                >
                    <p class="text-sm font-medium mt-0.5">{errorMsg}</p>
                </div>
            {/if}

            <!-- STEP 1: UPLOAD -->
            {#if step === 1}
                <div
                    in:fade={{ duration: 400, delay: 100 }}
                    out:fade={{ duration: 200 }}
                    class="space-y-8"
                >
                    <div class="text-center max-w-2xl mx-auto space-y-3 mb-12">
                        <h2
                            class="text-4xl font-extrabold tracking-tight text-surface-50"
                        >
                            Load your Data
                        </h2>
                        <p class="text-surface-400 text-lg">
                            Drop your raw spreadsheet and the master Word
                            template to instantly begin generating perfectly
                            formatted documents.
                        </p>
                    </div>

                    <div class="grid md:grid-cols-2 gap-8">
                        <!-- CSV Drop -->
                        <div class="relative group">
                            <div
                                class={`absolute inset-0 bg-gradient-to-br from-primary-500/10 to-transparent rounded-[2rem] blur-xl transition-opacity duration-500 ${csvFile ? "opacity-100" : "opacity-0 group-hover:opacity-50"}`}
                            ></div>
                            <label
                                class={`relative flex flex-col items-center justify-center h-72 border-2 border-dashed rounded-[2rem] cursor-pointer transition-all duration-300 ${csvFile ? "border-primary-500/50 bg-primary-500/5" : "border-surface-700 hover:border-surface-500 hover:bg-surface-800/50"}`}
                            >
                                <input
                                    type="file"
                                    accept=".csv"
                                    class="hidden"
                                    on:change={onCsvDrop}
                                />
                                <FileCsv
                                    weight={csvFile ? "fill" : "duotone"}
                                    class={`w-20 h-20 mb-6 transition-all duration-300 ${csvFile ? "text-primary-500 scale-110 drop-shadow-[0_0_15px_rgba(244,63,94,0.4)]" : "text-surface-600 group-hover:text-primary-400/80"}`}
                                />
                                <h3
                                    class="text-xl items-center font-bold text-surface-200 mb-2 truncate px-8 max-w-full"
                                >
                                    {csvFile
                                        ? csvFile.name
                                        : "Select CSV Dataset"}
                                </h3>
                                <p class="text-surface-500 text-sm">
                                    {csvFile
                                        ? "Ready for processing"
                                        : "Drag & drop or browse"}
                                </p>

                                {#if csvFile}
                                    <div
                                        class="absolute top-6 right-6 text-primary-500 bg-surface-900 rounded-full border border-surface-800 shadow-xl"
                                        in:scale
                                    >
                                        <CheckCircle
                                            weight="fill"
                                            class="w-8 h-8"
                                        />
                                    </div>
                                {/if}
                            </label>
                        </div>

                        <!-- Template Drop -->
                        <div class="relative group">
                            <div
                                class={`absolute inset-0 bg-gradient-to-br from-blue-500/10 to-transparent rounded-[2rem] blur-xl transition-opacity duration-500 ${docxFile ? "opacity-100" : "opacity-0 group-hover:opacity-50"}`}
                            ></div>
                            <label
                                class={`relative flex flex-col items-center justify-center h-72 border-2 border-dashed rounded-[2rem] cursor-pointer transition-all duration-300 ${docxFile ? "border-blue-500/50 bg-blue-500/5" : "border-surface-700 hover:border-surface-500 hover:bg-surface-800/50"}`}
                            >
                                <input
                                    type="file"
                                    accept=".docx"
                                    class="hidden"
                                    on:change={onDocxDrop}
                                />
                                <FileDoc
                                    weight={docxFile ? "fill" : "duotone"}
                                    class={`w-20 h-20 mb-6 transition-all duration-300 ${docxFile ? "text-blue-500 scale-110 drop-shadow-[0_0_15px_rgba(59,130,246,0.4)]" : "text-surface-600 group-hover:text-blue-400/80"}`}
                                />
                                <h3
                                    class="text-xl items-center font-bold text-surface-200 mb-2 truncate px-8 max-w-full"
                                >
                                    {docxFile
                                        ? docxFile.name
                                        : "Select DOCX Template"}
                                </h3>
                                <p class="text-surface-500 text-sm">
                                    {docxFile
                                        ? "Ready for mapping"
                                        : "Drag & drop or browse"}
                                </p>

                                {#if docxFile}
                                    <div
                                        class="absolute top-6 right-6 text-blue-500 bg-surface-900 rounded-full border border-surface-800 shadow-xl"
                                        in:scale
                                    >
                                        <CheckCircle
                                            weight="fill"
                                            class="w-8 h-8"
                                        />
                                    </div>
                                {/if}
                            </label>
                        </div>
                    </div>

                    <div class="flex justify-end pt-8">
                        <button
                            disabled={!csvFile || !docxFile || isUploading}
                            class="group relative inline-flex items-center gap-3 px-8 py-4 bg-surface-50 text-surface-950 hover:bg-white font-bold rounded-2xl overflow-hidden disabled:opacity-50 disabled:cursor-not-allowed transition-all hover:shadow-[0_0_30px_rgba(255,255,255,0.2)]"
                            on:click={handleUpload}
                        >
                            {#if isUploading}
                                <Spinner class="w-5 h-5 animate-spin" />
                                <span>Analyzing...</span>
                            {:else}
                                <span>Proceed to Mapping</span>
                                <ArrowRight
                                    weight="bold"
                                    class="w-5 h-5 group-hover:translate-x-1 transition-transform"
                                />
                            {/if}
                        </button>
                    </div>
                </div>
            {/if}

            <!-- STEP 2: MAPPING (The Mindblowing UI) -->
            {#if step === 2}
                <div
                    in:fade={{ duration: 400, delay: 100 }}
                    out:fade={{ duration: 200 }}
                    class="flex flex-col h-full"
                >
                    <div class="flex items-end justify-between mb-8">
                        <div>
                            <h2
                                class="text-3xl font-extrabold text-surface-50 mb-2"
                            >
                                Configure Mappings
                            </h2>
                            <p class="text-surface-400">
                                Map <strong class="text-primary-400"
                                    >{docxPlaceholders.length} tags</strong
                                > to data columns or custom static text.
                            </p>
                        </div>
                        <div
                            class="px-4 py-2 bg-surface-800/50 rounded-lg border border-surface-700/50 flex items-center gap-3"
                        >
                            <div
                                class="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"
                            ></div>
                            <span class="text-sm font-medium text-surface-300"
                                >Targeting {totalRows} Rows</span
                            >
                        </div>
                    </div>

                    <div
                        class="grid lg:grid-cols-2 gap-4 auto-rows-max overflow-y-auto pb-8 pr-2 custom-scrollbar"
                    >
                        {#each docxPlaceholders as ph}
                            {@const map = mappings[ph]}

                            <div
                                class={`p-5 rounded-2xl border transition-all duration-300 ${map.value || map.customText ? "border-primary-500/30 bg-primary-500/5" : "border-surface-700 bg-surface-800/30"} flex flex-col gap-4`}
                            >
                                <div class="flex items-center justify-between">
                                    <div
                                        class="font-mono text-sm px-3 py-1 bg-surface-950 text-primary-400 rounded-lg shadow-inner border border-surface-800 truncate"
                                        title={ph}
                                    >
                                        {ph}
                                    </div>

                                    <!-- Toggle: CSV vs Custom vs Combined -->
                                    <div
                                        class="flex p-1 bg-surface-950 rounded-lg border border-surface-800/80"
                                    >
                                        <button
                                            class={`p-1.5 rounded-md transition-all ${map.type === "csv_column" ? "bg-surface-800 text-white shadow-sm" : "text-surface-500 hover:text-surface-300"}`}
                                            on:click={() =>
                                                (mappings[ph].type =
                                                    "csv_column")}
                                            title="Map solely to Data Column"
                                        >
                                            <Table
                                                weight="duotone"
                                                class="w-4 h-4"
                                            />
                                        </button>
                                        <button
                                            class={`p-1.5 rounded-md transition-all ${map.type === "combined" ? "bg-surface-800 text-purple-400 shadow-sm" : "text-surface-500 hover:text-surface-300"}`}
                                            on:click={() =>
                                                (mappings[ph].type =
                                                    "combined")}
                                            title="Combine Static Text with Data Column"
                                        >
                                            <div
                                                class="font-[Inter] text-[10px] items-center flex font-bold w-4 h-4 justify-center leading-none tracking-tighter"
                                            >
                                                C+T
                                            </div>
                                        </button>
                                        <button
                                            class={`p-1.5 rounded-md transition-all ${map.type === "custom_text" ? "bg-surface-800 text-white shadow-sm" : "text-surface-500 hover:text-surface-300"}`}
                                            on:click={() =>
                                                (mappings[ph].type =
                                                    "custom_text")}
                                            title="Replace entirely with Static Text"
                                        >
                                            <TextT
                                                weight="duotone"
                                                class="w-4 h-4"
                                            />
                                        </button>
                                    </div>
                                </div>

                                {#if map.type === "csv_column"}
                                    <div
                                        class="relative"
                                        in:slide={{ duration: 200 }}
                                    >
                                        <select
                                            bind:value={mappings[ph].value}
                                            class="w-full appearance-none bg-surface-950 border border-surface-700 focus:border-primary-500 text-surface-200 text-sm font-medium rounded-xl px-4 py-3 pr-10 outline-none transition-colors shadow-inner"
                                        >
                                            <option value=""
                                                >-- Ignore Tag --</option
                                            >
                                            {#each csvHeaders as header}
                                                <option value={header}
                                                    >{header}</option
                                                >
                                            {/each}
                                        </select>
                                        <div
                                            class="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-surface-500"
                                        >
                                            <ArrowRight
                                                weight="bold"
                                                class="w-4 h-4 rotate-90"
                                            />
                                        </div>
                                    </div>
                                {:else if map.type === "combined"}
                                    <div
                                        class="flex flex-col sm:flex-row gap-2 relative"
                                        in:slide={{ duration: 200 }}
                                    >
                                        <input
                                            type="text"
                                            bind:value={mappings[ph].prefix}
                                            placeholder="Prefix (e.g. ID: )"
                                            class="sm:w-1/4 bg-surface-950 border border-purple-500/30 focus:border-purple-500 focus:ring-1 focus:ring-purple-500 text-white text-sm font-medium rounded-xl px-3 py-3 outline-none transition-all shadow-inner placeholder:text-surface-600"
                                        />
                                        <div class="relative sm:flex-1">
                                            <select
                                                bind:value={mappings[ph].value}
                                                class="w-full appearance-none bg-surface-950 border border-purple-500/50 focus:border-purple-500 text-surface-200 text-sm font-medium rounded-xl px-3 py-3 pr-10 outline-none transition-colors shadow-inner"
                                            >
                                                <option value=""
                                                    >- Column -</option
                                                >
                                                {#each csvHeaders as header}
                                                    <option value={header}
                                                        >{header}</option
                                                    >
                                                {/each}
                                            </select>
                                            <div
                                                class="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-purple-400"
                                            >
                                                <Table
                                                    weight="bold"
                                                    class="w-4 h-4"
                                                />
                                            </div>
                                        </div>
                                        <input
                                            type="text"
                                            bind:value={mappings[ph].suffix}
                                            placeholder="Suffix (e.g.  items)"
                                            class="sm:w-1/4 bg-surface-950 border border-purple-500/30 focus:border-purple-500 focus:ring-1 focus:ring-purple-500 text-white text-sm font-medium rounded-xl px-3 py-3 outline-none transition-all shadow-inner placeholder:text-surface-600"
                                        />
                                    </div>
                                {:else}
                                    <div
                                        class="relative"
                                        in:slide={{ duration: 200 }}
                                    >
                                        <input
                                            type="text"
                                            bind:value={mappings[ph].customText}
                                            placeholder="Type custom text to inject..."
                                            class="w-full bg-surface-950 border border-primary-500/50 focus:border-primary-500 focus:ring-1 focus:ring-primary-500 text-white text-sm font-medium rounded-xl px-4 py-3 outline-none transition-all shadow-inner placeholder:text-surface-600"
                                        />
                                    </div>
                                {/if}
                            </div>
                        {/each}
                    </div>

                    <!-- Options Footer -->
                    <div
                        class="mt-auto pt-6 border-t border-surface-800/80 mt-6 grid md:grid-cols-2 gap-6 items-center"
                    >
                        <div class="flex gap-4">
                            <!-- Toggle Card: DOCX -->
                            <button
                                class={`flex-1 flex flex-col items-center justify-center p-4 rounded-2xl border-2 transition-all ${options.docx ? "bg-blue-500/10 border-blue-500 text-blue-400 shadow-[0_0_20px_rgba(59,130,246,0.15)]" : "border-surface-700 bg-surface-800 text-surface-500 hover:border-surface-600"}`}
                                on:click={() => (options.docx = !options.docx)}
                            >
                                <span class="font-bold">.DOCX</span>
                            </button>

                            <!-- Toggle Card: PDF -->
                            <button
                                class={`flex-1 flex flex-col items-center justify-center p-4 rounded-2xl border-2 transition-all ${options.pdf ? "bg-rose-500/10 border-rose-500 text-rose-400 shadow-[0_0_20px_rgba(244,63,94,0.15)]" : "border-surface-700 bg-surface-800 text-surface-500 hover:border-surface-600"}`}
                                on:click={() => (options.pdf = !options.pdf)}
                            >
                                <span class="font-bold">.PDF</span>
                            </button>
                        </div>

                        <button
                            disabled={Object.keys(mappings).length === 0 ||
                                isGenerating}
                            class="group relative flex items-center justify-center gap-3 w-full py-4 bg-primary-500 text-white font-bold rounded-2xl overflow-hidden shadow-[0_0_20px_rgba(244,63,94,0.3)] hover:shadow-[0_0_40px_rgba(244,63,94,0.5)] transition-all hover:bg-primary-400 disabled:opacity-50 disabled:shadow-none"
                            on:click={handleGenerate}
                        >
                            {#if isGenerating}
                                <Spinner class="w-5 h-5 animate-spin" />
                                <span>Generating {totalRows} Files...</span>
                            {:else}
                                <span>Forge Documents</span>
                                <ArrowRight weight="bold" class="w-5 h-5" />
                            {/if}
                        </button>
                    </div>
                </div>
            {/if}

            <!-- STEP 3: DONE -->
            {#if step === 3}
                <div
                    in:fade={{ duration: 500 }}
                    class="flex flex-col items-center justify-center h-full text-center max-w-lg mx-auto py-12"
                >
                    <div
                        class="w-32 h-32 mb-8 relative flex flex-col items-center justify-center"
                    >
                        <div
                            class="absolute inset-0 bg-emerald-500/20 rounded-full blur-2xl animate-pulse"
                        ></div>
                        <div
                            class="relative bg-surface-900 border border-emerald-500/30 rounded-full p-6 shadow-2xl"
                        >
                            <CheckCircle
                                weight="fill"
                                class="w-16 h-16 text-emerald-500"
                            />
                        </div>
                    </div>

                    <h2 class="text-4xl font-extrabold text-white mb-4">
                        Forge Complete!
                    </h2>
                    <p class="text-surface-400 text-lg mb-10 leading-relaxed">
                        Successfully generated and packaged your formatted
                        documents into a unified archive.
                    </p>

                    <a
                        href={downloadUrl}
                        download="DocumentForge_Output.zip"
                        class="flex border-2 border-emerald-500/50 shadow-[0_0_30px_rgba(16,185,129,0.2)] hover:border-emerald-400 hover:shadow-[0_0_40px_rgba(16,185,129,0.4)] transition-all items-center gap-3 px-10 py-5 bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20 font-bold rounded-2xl text-lg w-full justify-center group"
                    >
                        <DownloadSimple
                            weight="bold"
                            class="w-6 h-6 group-hover:-translate-y-1 transition-transform"
                        />
                        <span>Download ZIP Archive</span>
                    </a>

                    <button
                        class="mt-8 text-surface-500 font-medium hover:text-white transition-colors"
                        on:click={() => window.location.reload()}
                    >
                        Start Another Batch
                    </button>
                </div>
            {/if}
        </main>
    </div>
</div>

<style>
    .custom-scrollbar::-webkit-scrollbar {
        width: 6px;
    }
    .custom-scrollbar::-webkit-scrollbar-track {
        background: transparent;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
        background: #334155;
        border-radius: 10px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {
        background: #475569;
    }
</style>
