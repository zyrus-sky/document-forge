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
        GearSix,
        CaretDown,
    } from "phosphor-svelte";
    import { slide, scale, fade } from "svelte/transition";
    import { onMount } from "svelte";

    // State
    let step = 1;
    let dataFile: File | null = null;
    let docxFile: File | null = null;
    let isUploading = false;
    let sessionId = "";
    let errorMsg = "";

    // Metadata
    let previewRows: Record<string, string>[] = [];
    let rawHeaders: string[] = [];
    let isEditingData = false;
    let isSavingData = false;
    let findText = "";
    let replaceText = "";
    let progressCurrent = 0;
    let progressTotal = 0;
    let progressMessage = "";
    let wsConnection: WebSocket | null = null;

    function replaceAllData() {
        if (!findText) return;
        previewRows = previewRows.map((row) => {
            const newRow = { ...row };
            for (const key of Object.keys(newRow)) {
                if (typeof newRow[key] === "string") {
                    newRow[key] = newRow[key].replaceAll(findText, replaceText);
                }
            }
            return newRow;
        });
    }
    let dataHeaders: string[] = [];
    let docxPlaceholders: string[] = [];
    let totalRows = 0;

    // Mapping State
    type MappingType = "data_column" | "custom_text" | "combined";
    interface MappingConfig {
        type: MappingType;
        value: string;
        customText: string;
        prefix: string;
        suffix: string;
        fallback?: string;
        modifier?: string;
        fromMemory?: boolean;
    }
    let mappings: Record<string, MappingConfig> = {};

    // Generation State
    let options = {
        docx: true,
        pdf: true,
        removeEmptyPages: true,
        mergeOutput: false,
    };
    let isGenerating = false;
    let generationComplete = false;

    // Document Settings State
    let showDocSettings = false;
    let docSettings = {
        pageSize: "default",
        customWidth: 8.5,
        customHeight: 11,
        fontName: "",
        fontSize: 0,
    };
    let templatePageSize = "letter";
    let templatePageWidth = 8.5;
    let templatePageHeight = 11;

    // Multi-Row Batching State
    let rowsPerDoc = 1;
    let placeholderCounts: Record<string, number> = {};

    $: estimatedDocs =
        rowsPerDoc > 0 ? Math.ceil(totalRows / rowsPerDoc) : totalRows;

    const COMMON_FONTS = [
        "",
        "Arial",
        "Times New Roman",
        "Calibri",
        "Cambria",
        "Georgia",
        "Verdana",
        "Tahoma",
        "Trebuchet MS",
        "Courier New",
        "Comic Sans MS",
        "Impact",
        "Garamond",
        "Book Antiqua",
        "Palatino Linotype",
    ];

    const PAGE_SIZE_OPTIONS = [
        { value: "default", label: "Default (from Template)" },
        { value: "letter", label: 'Letter (8.5" × 11")' },
        { value: "legal", label: 'Legal (8.5" × 14")' },
        { value: "a4", label: 'A4 (8.27" × 11.69")' },
        { value: "a3", label: 'A3 (11.69" × 16.54")' },
        { value: "a5", label: 'A5 (5.83" × 8.27")' },
        { value: "custom", label: "Custom" },
    ];
    let downloadUrl = "";

    // Live Preview State
    $: previewOutput = computePreview(mappings, previewRows);

    function computePreview(maps: Record<string, MappingConfig>, rows: any[]) {
        if (!rows || rows.length === 0) return {};
        const firstRow = rows[0];
        const result: Record<string, string> = {};

        for (const [ph, conf] of Object.entries(maps)) {
            let val = "";
            if (conf.type === "custom_text") {
                val = conf.customText;
            } else if (conf.type === "data_column" && conf.value) {
                val = firstRow[conf.value] || "";
            }

            // Apply modifiers and conditionals
            if (conf.prefix) val = conf.prefix + val;
            if (conf.suffix) val = val + conf.suffix;

            // Inline function basic support: uppercase/lowercase/capitalize
            // Normally handled by Python backend, but let's preview it live
            // Assume format: "modifier(column_name)" => we simulate it via a new modifier property,
            // or just parse the 'customText' for advanced conditional rules.
            // For MVP Phase 3: Add uppercase/lowercase radio or dropdown in the UI?
            // Lets just use simple JS transforms if requested.
            if (conf.type === "data_column") {
                if (conf.modifier === "uppercase") val = val.toUpperCase();
                if (conf.modifier === "lowercase") val = val.toLowerCase();
                // Conditional: fallback text if empty
                if (!val && conf.fallback) val = conf.fallback;
            }

            result[ph] = val;
        }
        return result;
    }

    // FastAPI Root
    const API_URL = "http://localhost:8000/api";

    // Auto-save mappings when they change
    $: {
        if (Object.keys(mappings).length > 0) {
            try {
                localStorage.setItem(
                    "docforge_mappings",
                    JSON.stringify(mappings),
                );
            } catch (e) {}
        }
    }

    // Utility for Map All
    function levenshteinDistance(a: string, b: string): number {
        const matrix = [];
        let i, j;
        if (a.length === 0) return b.length;
        if (b.length === 0) return a.length;
        for (i = 0; i <= b.length; i++) matrix[i] = [i];
        for (j = 0; j <= a.length; j++) matrix[0][j] = j;
        for (i = 1; i <= b.length; i++) {
            for (j = 1; j <= a.length; j++) {
                if (b.charAt(i - 1) === a.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1,
                        Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1),
                    );
                }
            }
        }
        return matrix[b.length][a.length];
    }

    function exportProfile() {
        const dataStr =
            "data:text/json;charset=utf-8," +
            encodeURIComponent(JSON.stringify(mappings, null, 2));
        const anchor = document.createElement("a");
        anchor.setAttribute("href", dataStr);
        anchor.setAttribute("download", "mapping_profile.json");
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
    }

    function importProfile(e: any) {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const loaded = JSON.parse(event.target?.result as string);
                if (loaded) {
                    mappings = loaded;
                    // Auto-save immediately
                    localStorage.setItem(
                        "docforge_mappings",
                        JSON.stringify(mappings),
                    );
                }
            } catch (err) {
                console.error("Invalid profile JSON");
                errorMsg = "Failed to parse mapping profile.";
            }
        };
        reader.readAsText(file);
    }

    function smartMapAll() {
        if (!dataHeaders.length || !docxPlaceholders.length) return;

        docxPlaceholders.forEach((ph) => {
            const cleanPh = ph
                .replace("#", "")
                .toLowerCase()
                .replace(/\s/g, "");
            let bestMatch = "";
            let minDistance = Infinity;

            for (const header of dataHeaders) {
                const cleanHeader = header.toLowerCase().replace(/\s/g, "");
                // Exact match takes precedence
                if (
                    cleanHeader === cleanPh ||
                    (header === "P_ ADDRESS" && ph === "#P_ ADDRESS")
                ) {
                    bestMatch = header;
                    minDistance = 0;
                    break;
                }

                // Fuzzy match
                const dist = levenshteinDistance(cleanPh, cleanHeader);
                // Threshold of 3 character difference for a mapping hint
                if (dist < minDistance && dist <= 3) {
                    minDistance = dist;
                    bestMatch = header;
                }
            }

            if (bestMatch) {
                mappings[ph] = {
                    type: "data_column",
                    value: bestMatch,
                    customText: "",
                    prefix: "",
                    suffix: "",
                };
            }
        });

        // Trigger reactivity
        mappings = { ...mappings };
    }

    // Reactive auto-save to global memory whenever mappings change
    $: if (Object.keys(mappings).length > 0) {
        try {
            let memory: Record<string, any> = {};
            const stored = localStorage.getItem("docforge_global_memory");
            if (stored) memory = JSON.parse(stored);

            // Update memory with current mappings
            for (const [ph, config] of Object.entries(mappings)) {
                if (config.value || config.customText) {
                    memory[ph] = config;
                }
            }
            localStorage.setItem(
                "docforge_global_memory",
                JSON.stringify(memory),
            );
        } catch (e) {}
    }

    function autoDetectMappings() {
        let memory: Record<string, MappingConfig> = {};
        try {
            const stored = localStorage.getItem("docforge_global_memory");
            if (stored) memory = JSON.parse(stored);
        } catch (e) {}

        docxPlaceholders.forEach((ph) => {
            // Check Global Memory first
            if (memory[ph]) {
                const memConfig = memory[ph];
                // For data_column or combined, verify the header actually exists in this CSV
                if (
                    (memConfig.type === "data_column" ||
                        memConfig.type === "combined") &&
                    dataHeaders.includes(memConfig.value)
                ) {
                    mappings[ph] = { ...memConfig, fromMemory: true };
                    return;
                } else if (
                    memConfig.type === "custom_text" &&
                    memConfig.customText
                ) {
                    mappings[ph] = { ...memConfig, fromMemory: true };
                    return;
                }
            }

            // Fallback: Levenshtein / Exact String Match
            const cleanPh = ph
                .replace("#", "")
                .toLowerCase()
                .replace(/\s/g, "");
            let bestMatch = "";

            for (const header of dataHeaders) {
                const cleanHeader = header.toLowerCase().replace(/\s/g, "");
                if (
                    cleanHeader === cleanPh ||
                    (header === "P_ ADDRESS" && ph === "#P_ ADDRESS")
                ) {
                    bestMatch = header;
                    break;
                }
            }

            mappings[ph] = {
                type: bestMatch ? "data_column" : "custom_text",
                value: bestMatch,
                customText: "",
                prefix: "",
                suffix: "",
            };
        });
    }

    async function handleUpload() {
        if (!dataFile || !docxFile) return;
        isUploading = true;
        errorMsg = "";

        const formData = new FormData();
        formData.append("data_file", dataFile);
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

            dataHeaders = metaData.csv_headers;
            docxPlaceholders = metaData.docx_placeholders;
            totalRows = metaData.total_rows;
            previewRows = metaData.preview_rows;
            rawHeaders = metaData.raw_headers;

            // Multi-row & template page info
            placeholderCounts = metaData.placeholder_counts || {};
            rowsPerDoc = metaData.rows_per_doc || 1;
            templatePageSize = metaData.template_page_size || "letter";
            templatePageWidth = metaData.template_page_width || 8.5;
            templatePageHeight = metaData.template_page_height || 11;

            autoDetectMappings();
            step = 2;
        } catch (err: any) {
            errorMsg = err.message;
        } finally {
            isUploading = false;
        }
    }

    async function saveEditedData() {
        isSavingData = true;
        try {
            const response = await fetch(`${API_URL}/update_data`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    session_id: sessionId,
                    rows: previewRows,
                    headers: rawHeaders,
                }),
            });
            if (!response.ok) throw new Error("Failed to save data");
            const resData = await response.json();
            totalRows = resData.total_rows;
            isEditingData = false;
        } catch (err: any) {
            errorMsg = err.message;
        } finally {
            isSavingData = false;
        }
    }

    async function handleGenerate() {
        isGenerating = true;
        errorMsg = "";
        progressCurrent = 0;
        progressTotal = totalRows;
        progressMessage = "Starting generation...";

        // Connect WebSocket
        const wsUrl = API_URL.replace("http://", "ws://").replace("/api", "");
        wsConnection = new WebSocket(`${wsUrl}/ws/progress/${sessionId}`);
        wsConnection.onmessage = (event) => {
            try {
                const data = JSON.parse(event.data);
                if (data.type === "progress") {
                    progressCurrent = data.current;
                    if (data.total) progressTotal = data.total;
                    if (data.message) progressMessage = data.message;
                }
            } catch (e) {}
        };

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
                    remove_empty_pages: options.removeEmptyPages,
                    merge_output: options.mergeOutput,
                    rows_per_doc: rowsPerDoc,
                    doc_settings: {
                        page_size: docSettings.pageSize,
                        page_width:
                            docSettings.pageSize === "custom"
                                ? docSettings.customWidth
                                : null,
                        page_height:
                            docSettings.pageSize === "custom"
                                ? docSettings.customHeight
                                : null,
                        font_name: docSettings.fontName || null,
                        font_size: docSettings.fontSize || null,
                    },
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
            if (wsConnection) {
                wsConnection.close();
                wsConnection = null;
            }
            isGenerating = false;
        }
    }

    // File Handlers
    function onDataDrop(e: any) {
        dataFile = e.dataTransfer?.files[0] || e.target?.files[0];
    }
    function onDocxDrop(e: any) {
        docxFile = e.dataTransfer?.files[0] || e.target?.files[0];
    }
</script>

<div class="w-full max-w-5xl mx-auto glass-card flex flex-col relative">
    <!-- Header -->
    <header
        class="p-6 md:px-10 border-b border-white/5 glass-material flex items-center justify-between z-10 sticky top-0"
    >
        <div class="flex items-center gap-4">
            <div
                class="glass-icon p-3 bg-gradient-to-br from-primary-500/30 to-primary-700/30 shadow-[0_0_20px_rgba(244,63,94,0.3)]"
            >
                <Faders class="w-6 h-6 text-white" weight="duotone" />
            </div>
            <div>
                <h1 class="text-2xl font-bold text-surface-50 font-display">
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
                        class="text-4xl font-extrabold tracking-tight text-surface-50 font-display"
                    >
                        Load your Data
                    </h2>
                    <p class="text-surface-400 text-lg">
                        Drop your raw spreadsheet and the master Word template
                        to instantly begin generating perfectly formatted
                        documents.
                    </p>
                </div>

                <div class="grid md:grid-cols-2 gap-8">
                    <!-- Data Drop -->
                    <div class="relative group">
                        <div
                            class={`absolute inset-0 bg-gradient-to-br from-primary-500/10 to-transparent rounded-[2rem] blur-xl transition-opacity duration-500 ${dataFile ? "opacity-100" : "opacity-0 group-hover:opacity-50"}`}
                        ></div>
                        <label
                            class={`glass-surface relative flex flex-col items-center justify-center h-72 border-2 border-dashed cursor-pointer transition-all duration-300 ${dataFile ? "border-primary-500/50 !bg-primary-500/5" : "border-surface-700 hover:border-surface-500"}`}
                        >
                            <input
                                type="file"
                                accept=".csv,.xlsx,.json"
                                class="hidden"
                                on:change={onDataDrop}
                            />
                            <div
                                class="glass-icon glass-icon-lg mb-6 animate-float"
                            >
                                <FileCsv
                                    weight={dataFile ? "fill" : "duotone"}
                                    class={`w-10 h-10 transition-all duration-300 ${dataFile ? "text-primary-500 drop-shadow-[0_0_15px_rgba(244,63,94,0.4)]" : "text-surface-400 group-hover:text-primary-400/80"}`}
                                />
                            </div>
                            <h3
                                class="text-xl items-center font-bold text-surface-200 mb-2 truncate px-8 max-w-full font-heading"
                            >
                                {dataFile
                                    ? dataFile.name
                                    : "Select Dataset File"}
                            </h3>
                            <p class="text-surface-500 text-sm">
                                {dataFile
                                    ? "Ready for processing"
                                    : "Drag & drop or browse"}
                            </p>

                            {#if dataFile}
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
                            class={`glass-surface relative flex flex-col items-center justify-center h-72 border-2 border-dashed cursor-pointer transition-all duration-300 ${docxFile ? "border-blue-500/50 !bg-blue-500/5" : "border-surface-700 hover:border-surface-500"}`}
                        >
                            <input
                                type="file"
                                accept=".docx"
                                class="hidden"
                                on:change={onDocxDrop}
                            />
                            <div
                                class="glass-icon glass-icon-lg mb-6 animate-float"
                                style="animation-delay: 0.5s"
                            >
                                <FileDoc
                                    weight={docxFile ? "fill" : "duotone"}
                                    class={`w-10 h-10 transition-all duration-300 ${docxFile ? "text-blue-500 drop-shadow-[0_0_15px_rgba(59,130,246,0.4)]" : "text-surface-400 group-hover:text-blue-400/80"}`}
                                />
                            </div>
                            <h3
                                class="text-xl items-center font-bold text-surface-200 mb-2 truncate px-8 max-w-full font-heading"
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
                        disabled={!dataFile || !docxFile || isUploading}
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

                    <div class="flex items-center gap-3">
                        <button
                            class="px-4 py-2 bg-purple-500/20 text-purple-400 hover:bg-purple-500/30 rounded-lg border border-purple-500/30 text-sm font-bold transition-all flex items-center gap-2"
                            on:click={smartMapAll}
                        >
                            <TextT weight="fill" class="w-4 h-4" />
                            Smart Map All
                        </button>
                        <button
                            class="px-4 py-2 bg-blue-500/20 text-blue-400 hover:bg-blue-500/30 rounded-lg border border-blue-500/30 text-sm font-bold transition-all flex items-center gap-2"
                            on:click={() => (isEditingData = !isEditingData)}
                        >
                            <Table weight="fill" class="w-4 h-4" />
                            {isEditingData ? "Close Editor" : "Edit Live Data"}
                        </button>
                        <div
                            class="px-4 py-2 bg-surface-800/50 rounded-lg border border-surface-700/50 flex items-center gap-2"
                        >
                            <div
                                class="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"
                            ></div>
                            <span class="text-sm font-medium text-surface-300"
                                >Targeting {totalRows} Rows</span
                            >
                        </div>
                        {#if rowsPerDoc > 1}
                            <div
                                class="px-4 py-2 bg-amber-500/10 rounded-lg border border-amber-500/30 flex items-center gap-2"
                            >
                                <span class="text-sm font-bold text-amber-400"
                                    >⚡ {rowsPerDoc} rows/doc → {estimatedDocs} docs</span
                                >
                            </div>
                        {/if}
                    </div>
                </div>

                {#if isEditingData}
                    <div
                        class="mb-8 p-6 bg-surface-900 border border-blue-500/30 rounded-2xl shadow-xl shadow-blue-500/10 overflow-hidden flex flex-col"
                        in:slide={{ duration: 300 }}
                    >
                        <div class="flex justify-between items-center mb-4">
                            <h3 class="font-bold text-white text-lg">
                                Live Data Editor
                            </h3>
                            <button
                                class="bg-blue-500 text-white px-4 py-2 rounded-lg font-bold text-sm shadow-[0_0_15px_rgba(59,130,246,0.5)] hover:bg-blue-400 transition-colors flex items-center gap-2 disabled:opacity-50"
                                on:click={saveEditedData}
                                disabled={isSavingData}
                            >
                                {#if isSavingData}<Spinner
                                        class="w-4 h-4 animate-spin"
                                    />{/if}
                                Save Changes
                            </button>
                        </div>

                        <!-- Find & Replace Toolbar -->
                        <div
                            class="flex flex-col sm:flex-row gap-3 mb-4 bg-surface-950/50 p-3 rounded-xl border border-surface-700/50"
                        >
                            <input
                                type="text"
                                bind:value={findText}
                                placeholder="Find text..."
                                class="flex-1 bg-surface-900 border border-surface-700 focus:border-blue-500 rounded-lg px-3 py-2 text-sm text-white outline-none"
                            />
                            <input
                                type="text"
                                bind:value={replaceText}
                                placeholder="Replace with..."
                                class="flex-1 bg-surface-900 border border-surface-700 focus:border-blue-500 rounded-lg px-3 py-2 text-sm text-white outline-none"
                            />
                            <button
                                class="bg-surface-800 hover:bg-surface-700 text-surface-200 px-4 py-2 border border-surface-600 rounded-lg font-bold text-sm transition-colors"
                                on:click={replaceAllData}
                            >
                                Replace All
                            </button>
                        </div>

                        <div
                            class="overflow-x-auto overflow-y-auto max-h-[400px] border border-surface-700 rounded-xl custom-scrollbar"
                        >
                            <table
                                class="w-full text-left border-collapse text-sm"
                            >
                                <thead>
                                    <tr>
                                        {#each rawHeaders as head}
                                            <th
                                                class="p-3 bg-surface-800/80 border-b border-surface-700 font-bold text-surface-300 whitespace-nowrap sticky top-0 z-10 backdrop-blur-md"
                                                >{head}</th
                                            >
                                        {/each}
                                    </tr>
                                </thead>
                                <tbody>
                                    {#each previewRows as row, i}
                                        <tr
                                            class="hover:bg-surface-800/50 transition-colors"
                                        >
                                            {#each rawHeaders as head}
                                                <td
                                                    class={`p-0 border-b border-surface-700/50 bg-surface-950/50 focus-within:bg-blue-900/20 transition-colors ${!row[head] || row[head].trim() === "" ? "ring-2 ring-inset ring-red-500/50 bg-red-500/10" : ""}`}
                                                >
                                                    <input
                                                        type="text"
                                                        bind:value={row[head]}
                                                        class="w-full bg-transparent p-3 outline-none focus:ring-2 focus:ring-inset focus:ring-blue-500 text-surface-200"
                                                    />
                                                </td>
                                            {/each}
                                        </tr>
                                    {/each}
                                </tbody>
                            </table>
                        </div>
                        <p
                            class="text-xs text-surface-500 mt-3 font-medium text-center"
                        >
                            Showing early preview subset of targeted rows. Edit
                            cells to update the dataset prior to generation.
                        </p>
                    </div>
                {/if}

                <div class="grid lg:grid-cols-3 gap-6 h-[500px]">
                    <!-- Left: Mappings -->
                    <div
                        class="lg:col-span-2 overflow-y-auto pr-2 custom-scrollbar bg-surface-950/20 rounded-2xl border border-surface-800/50 p-4"
                    >
                        <div class="grid lg:grid-cols-2 gap-4 auto-rows-max">
                            {#each docxPlaceholders as ph}
                                {@const map = mappings[ph]}

                                <div
                                    class={`p-5 rounded-2xl border transition-all duration-300 ${map.value || map.customText ? "border-primary-500/30 bg-primary-500/5" : "border-surface-700 bg-surface-800/30"} flex flex-col gap-4`}
                                >
                                    <div
                                        class="flex items-center justify-between"
                                    >
                                        <div
                                            class="font-mono text-sm px-3 py-1 bg-surface-950 text-primary-400 rounded-lg shadow-inner border border-surface-800 truncate flex items-center gap-2"
                                            title={ph}
                                        >
                                            <span>{ph}</span>
                                            {#if map.fromMemory}
                                                <span
                                                    class="text-[10px] bg-purple-500/20 text-purple-300 px-1.5 py-0.5 rounded border border-purple-500/30 font-sans tracking-wide"
                                                    title="Restored from Global Memory"
                                                    >✨ Auto-Saved</span
                                                >
                                            {/if}
                                        </div>

                                        <!-- Toggle: CSV vs Custom vs Combined -->
                                        <div
                                            class="flex p-1 bg-surface-950 rounded-lg border border-surface-800/80"
                                        >
                                            <button
                                                class={`p-1.5 rounded-md transition-all ${map.type === "data_column" ? "bg-surface-800 text-white shadow-sm" : "text-surface-500 hover:text-surface-300"}`}
                                                on:click={() =>
                                                    (mappings[ph].type =
                                                        "data_column")}
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

                                    {#if map.type === "data_column"}
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
                                                {#each dataHeaders as header}
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
                                                    bind:value={
                                                        mappings[ph].value
                                                    }
                                                    class="w-full appearance-none bg-surface-950 border border-purple-500/50 focus:border-purple-500 text-surface-200 text-sm font-medium rounded-xl px-3 py-3 pr-10 outline-none transition-colors shadow-inner"
                                                >
                                                    <option value=""
                                                        >- Column -</option
                                                    >
                                                    {#each dataHeaders as header}
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
                                                bind:value={
                                                    mappings[ph].customText
                                                }
                                                placeholder="Type custom text to inject..."
                                                class="w-full bg-surface-950 border border-primary-500/50 focus:border-primary-500 focus:ring-1 focus:ring-primary-500 text-white text-sm font-medium rounded-xl px-4 py-3 outline-none transition-all shadow-inner placeholder:text-surface-600"
                                            />
                                        </div>
                                    {/if}
                                </div>
                            {/each}
                        </div>
                    </div>

                    <!-- Right: Live Preview Pane -->
                    <div
                        class="lg:col-span-1 bg-surface-900 border border-surface-800 rounded-2xl p-5 flex flex-col shadow-inner overflow-hidden"
                    >
                        <div
                            class="flex items-center gap-3 mb-4 sticky top-0 bg-surface-900 pb-2 border-b border-surface-800 z-10"
                        >
                            <div
                                class="w-2 h-2 rounded-full bg-blue-500 animate-pulse"
                            ></div>
                            <h3 class="font-bold text-surface-200">
                                Live Preview
                            </h3>
                            <span
                                class="text-xs text-surface-500 ml-auto font-medium bg-surface-950 px-2 py-1 rounded-md"
                                >Row 1 Output</span
                            >
                        </div>
                        <div
                            class="flex-1 overflow-y-auto custom-scrollbar space-y-4 pr-2"
                        >
                            {#each Object.entries(previewOutput) as [ph, val]}
                                <div
                                    class="bg-surface-950 rounded-lg p-3 border border-surface-800/50 hover:border-surface-700 transition-colors"
                                >
                                    <div
                                        class="text-xs font-mono text-purple-400 mb-1 truncate"
                                    >
                                        {ph}
                                    </div>
                                    <div
                                        class="text-sm text-surface-50 break-words font-medium"
                                    >
                                        {#if val === ""}
                                            <span
                                                class="text-surface-600 italic"
                                                >Empty string</span
                                            >
                                        {:else}
                                            {val}
                                        {/if}
                                    </div>
                                </div>
                            {/each}
                        </div>
                    </div>
                </div>

                <!-- Options Footer -->
                <div
                    class="mt-auto pt-6 border-t border-surface-800/80 mt-6 flex flex-col gap-6"
                >
                    <div class="grid md:grid-cols-2 gap-6 items-center">
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

                            <!-- Document Settings Toggle -->
                            <button
                                class={`flex-1 flex flex-col items-center justify-center p-4 rounded-2xl border-2 transition-all ${showDocSettings ? "bg-emerald-500/10 border-emerald-500 text-emerald-400 shadow-[0_0_20px_rgba(16,185,129,0.15)]" : "border-surface-700 bg-surface-800 text-surface-500 hover:border-surface-600"}`}
                                on:click={() =>
                                    (showDocSettings = !showDocSettings)}
                            >
                                <span class="font-bold flex items-center gap-2">
                                    <GearSix weight="fill" class="w-4 h-4" />
                                    Settings
                                </span>
                            </button>
                        </div>

                        <!-- Post-Processing Options -->
                        <div class="flex gap-4">
                            <!-- Toggle: Remove Empty Pages -->
                            <button
                                class={`flex-1 flex flex-col items-center justify-center p-4 rounded-2xl border-2 transition-all ${options.removeEmptyPages ? "bg-amber-500/10 border-amber-500 text-amber-400 shadow-[0_0_20px_rgba(245,158,11,0.15)]" : "border-surface-700 bg-surface-800 text-surface-500 hover:border-surface-600"}`}
                                on:click={() =>
                                    (options.removeEmptyPages =
                                        !options.removeEmptyPages)}
                            >
                                <span class="font-bold text-sm"
                                    >🧹 Strip Empty Pages</span
                                >
                            </button>

                            <!-- Toggle: Merge All -->
                            <button
                                class={`flex-1 flex flex-col items-center justify-center p-4 rounded-2xl border-2 transition-all ${options.mergeOutput ? "bg-violet-500/10 border-violet-500 text-violet-400 shadow-[0_0_20px_rgba(139,92,246,0.15)]" : "border-surface-700 bg-surface-800 text-surface-500 hover:border-surface-600"}`}
                                on:click={() =>
                                    (options.mergeOutput =
                                        !options.mergeOutput)}
                            >
                                <span class="font-bold text-sm"
                                    >📎 Merge into One File</span
                                >
                            </button>
                        </div>
                    </div>

                    <!-- Document Settings Panel (Collapsible) -->
                    {#if showDocSettings}
                        <div
                            class="p-5 bg-surface-900 border border-emerald-500/30 rounded-2xl shadow-xl shadow-emerald-500/5"
                            transition:slide={{ duration: 250 }}
                        >
                            <h3
                                class="font-bold text-white text-sm uppercase tracking-wider mb-4 flex items-center gap-2"
                            >
                                <GearSix
                                    weight="duotone"
                                    class="w-4 h-4 text-emerald-400"
                                />
                                Document Settings
                            </h3>

                            <div class="grid sm:grid-cols-3 gap-4">
                                <!-- Page Size -->
                                <div class="flex flex-col gap-2">
                                    <label
                                        class="text-xs font-bold text-surface-400 uppercase tracking-wider"
                                    >
                                        Page Size
                                    </label>
                                    <div class="relative">
                                        <select
                                            bind:value={docSettings.pageSize}
                                            class="w-full appearance-none bg-surface-950 border border-surface-700 focus:border-emerald-500 text-surface-200 text-sm font-medium rounded-xl px-4 py-3 pr-10 outline-none transition-colors shadow-inner"
                                        >
                                            {#each PAGE_SIZE_OPTIONS as opt}
                                                <option value={opt.value}
                                                    >{opt.label}</option
                                                >
                                            {/each}
                                        </select>
                                        <div
                                            class="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-surface-500"
                                        >
                                            <CaretDown
                                                weight="bold"
                                                class="w-4 h-4"
                                            />
                                        </div>
                                    </div>
                                    {#if docSettings.pageSize === "default"}
                                        <span
                                            class="text-[11px] text-surface-500 font-medium"
                                        >
                                            Template: {templatePageSize.toUpperCase()}
                                            ({templatePageWidth}" × {templatePageHeight}")
                                        </span>
                                    {/if}
                                    {#if docSettings.pageSize === "custom"}
                                        <div class="flex gap-2 mt-1">
                                            <input
                                                type="number"
                                                step="0.01"
                                                bind:value={
                                                    docSettings.customWidth
                                                }
                                                placeholder="Width (in)"
                                                class="flex-1 bg-surface-950 border border-surface-700 focus:border-emerald-500 text-white text-sm rounded-lg px-3 py-2 outline-none"
                                            />
                                            <span
                                                class="text-surface-500 flex items-center text-sm"
                                                >×</span
                                            >
                                            <input
                                                type="number"
                                                step="0.01"
                                                bind:value={
                                                    docSettings.customHeight
                                                }
                                                placeholder="Height (in)"
                                                class="flex-1 bg-surface-950 border border-surface-700 focus:border-emerald-500 text-white text-sm rounded-lg px-3 py-2 outline-none"
                                            />
                                        </div>
                                    {/if}
                                </div>

                                <!-- Font Name -->
                                <div class="flex flex-col gap-2">
                                    <label
                                        class="text-xs font-bold text-surface-400 uppercase tracking-wider"
                                    >
                                        Font
                                    </label>
                                    <div class="relative">
                                        <select
                                            bind:value={docSettings.fontName}
                                            class="w-full appearance-none bg-surface-950 border border-surface-700 focus:border-emerald-500 text-surface-200 text-sm font-medium rounded-xl px-4 py-3 pr-10 outline-none transition-colors shadow-inner"
                                        >
                                            <option value=""
                                                >Keep Template Font</option
                                            >
                                            {#each COMMON_FONTS.slice(1) as font}
                                                <option value={font}
                                                    >{font}</option
                                                >
                                            {/each}
                                        </select>
                                        <div
                                            class="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-surface-500"
                                        >
                                            <CaretDown
                                                weight="bold"
                                                class="w-4 h-4"
                                            />
                                        </div>
                                    </div>
                                </div>

                                <!-- Font Size -->
                                <div class="flex flex-col gap-2">
                                    <label
                                        class="text-xs font-bold text-surface-400 uppercase tracking-wider"
                                    >
                                        Font Size (pt)
                                    </label>
                                    <input
                                        type="number"
                                        min="0"
                                        max="72"
                                        bind:value={docSettings.fontSize}
                                        placeholder="0 = keep template"
                                        class="w-full bg-surface-950 border border-surface-700 focus:border-emerald-500 text-white text-sm font-medium rounded-xl px-4 py-3 outline-none transition-colors shadow-inner placeholder:text-surface-600"
                                    />
                                    <span
                                        class="text-[11px] text-surface-500 font-medium"
                                    >
                                        0 = keep template size
                                    </span>
                                </div>
                            </div>
                        </div>
                    {/if}

                    {#if isGenerating}
                        <div
                            class="relative w-full bg-surface-800 rounded-2xl overflow-hidden shadow-[0_0_20px_rgba(244,63,94,0.15)] flex flex-col p-3 border border-primary-500/30 gap-2"
                        >
                            <!-- Progress Bar Background/Fill -->
                            <div
                                class="w-full h-3 bg-surface-950 rounded-full overflow-hidden shrink-0"
                            >
                                <div
                                    class="h-full bg-primary-500 transition-all duration-300 ease-out flex items-center justify-center relative shadow-[0_0_10px_rgba(244,63,94,0.8)]"
                                    style={`width: ${progressTotal > 0 ? Math.min(100, (progressCurrent / progressTotal) * 100) : 0}%`}
                                >
                                    <!-- Shine effect -->
                                    <div
                                        class="absolute inset-0 bg-white/20 w-1/2 skew-x-12 animate-slide"
                                    ></div>
                                </div>
                            </div>
                            <div
                                class="flex justify-between items-center text-xs font-semibold px-1"
                            >
                                <span
                                    class="text-surface-300 flex items-center gap-2"
                                >
                                    <Spinner class="w-3 h-3 animate-spin" />
                                    {progressMessage}
                                </span>
                                <span class="text-primary-400"
                                    >{progressCurrent} / {progressTotal}</span
                                >
                            </div>
                        </div>
                    {:else}
                        <button
                            disabled={Object.keys(mappings).length === 0}
                            class="group relative flex items-center justify-center gap-3 w-full py-4 bg-primary-500 text-white font-bold rounded-2xl overflow-hidden shadow-[0_0_20px_rgba(244,63,94,0.3)] hover:shadow-[0_0_40px_rgba(244,63,94,0.5)] transition-all hover:bg-primary-400 disabled:opacity-50 disabled:shadow-none"
                            on:click={handleGenerate}
                        >
                            <span
                                >Forge {rowsPerDoc > 1
                                    ? `${estimatedDocs} Documents`
                                    : "Documents"}</span
                            >
                            <ArrowRight
                                weight="bold"
                                class="w-5 h-5 group-hover:translate-x-1 transition-transform"
                            />
                        </button>
                    {/if}
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
                    Successfully generated and packaged your formatted documents
                    into a unified archive.
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
