<script lang="ts">
    import {
        FilePdf,
        UploadSimple,
        Table,
        TextT,
        Spinner,
        ArrowRight,
        XCircle,
    } from "phosphor-svelte";
    import { slide, scale, fade } from "svelte/transition";

    let step = 1;
    let pdfFile: File | null = null;
    let isUploading = false;
    let processingOption: "allText" | "tablesOnly" = "allText";
    let outputFormat: "excel" | "csv" = "excel";
    let errorMsg = "";

    const API_URL = "http://localhost:8000/api";

    function handlePdfDrop(event: DragEvent) {
        event.preventDefault();
        const files = event.dataTransfer?.files;
        if (files && files.length > 0) processPdfSelection(files[0]);
    }

    function handlePdfSelect(event: Event) {
        const target = event.target as HTMLInputElement;
        if (target.files && target.files.length > 0)
            processPdfSelection(target.files[0]);
    }

    function processPdfSelection(file: File) {
        if (!file.name.toLowerCase().endsWith(".pdf")) {
            errorMsg = "Please upload a valid PDF file.";
            return;
        }
        pdfFile = file;
        errorMsg = "";
        step = 2; // Move to option selection immediately
    }

    function clearFile() {
        pdfFile = null;
        step = 1;
        errorMsg = "";
    }

    async function handleExtract() {
        if (!pdfFile) return;

        isUploading = true;
        errorMsg = "";

        const formData = new FormData();
        formData.append("pdfFile", pdfFile);
        formData.append("processingOption", processingOption);
        formData.append("outputFormat", outputFormat);

        try {
            const response = await fetch(`${API_URL}/converter/extract`, {
                method: "POST",
                body: formData,
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || "Extraction failed");
            }

            // Trigger file download
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;

            let ext = outputFormat === "csv" ? "csv" : "xlsx";
            if (processingOption === "tablesOnly" && outputFormat === "csv") {
                ext = "zip"; // Multiple tables as CSVs are zipped
            }
            a.download = `Extracted_${processingOption}_${new Date().getTime()}.${ext}`;

            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);

            clearFile(); // Reset after success
        } catch (err: any) {
            errorMsg =
                err.message || "Could not connect to the extraction server.";
        } finally {
            isUploading = false;
        }
    }
</script>

<div class="flex flex-col gap-6 w-full max-w-4xl mx-auto pb-20">
    <header class="text-center mb-4 mt-8">
        <div class="glass-icon glass-icon-lg inline-flex mb-4">
            <Table
                weight="duotone"
                class="w-8 h-8 text-purple-400 relative z-10"
            />
        </div>
        <h2
            class="text-4xl font-extrabold tracking-tight text-white mb-2 font-display"
        >
            PDF <span
                class="text-transparent bg-clip-text bg-gradient-to-r from-purple-400 to-rose-400"
                >Data Extractor</span
            >
        </h2>
        <p class="text-surface-400 max-w-lg mx-auto font-medium">
            Instantly convert complex PDF documents and tables into structured
            Excel spreadsheets.
        </p>
    </header>

    {#if errorMsg}
        <div
            transition:slide
            class="bg-red-500/10 border border-red-500/30 text-red-400 p-4 rounded-xl flex items-center justify-between text-sm font-bold shadow-lg shadow-red-500/5"
        >
            <div class="flex items-center gap-3">
                <XCircle weight="fill" class="w-5 h-5 text-red-500" />
                {errorMsg}
            </div>
            <button
                class="hover:bg-red-500/20 p-1.5 rounded-lg transition-colors"
                on:click={() => (errorMsg = "")}
            >
                <XCircle weight="bold" class="w-4 h-4" />
            </button>
        </div>
    {/if}

    <!-- STEP 1: UPLOAD -->
    <div
        class={`transition-all duration-500 ${step === 1 ? "opacity-100" : "opacity-60 scale-[0.98]"}`}
    >
        <div class="glass-card p-8 relative group">
            <div class="flex items-center gap-4 mb-6 relative z-10">
                <div
                    class={`w-8 h-8 rounded-full flex items-center justify-center font-bold text-sm transition-all duration-300 ${step >= 1 ? "bg-primary-500 text-white shadow-[0_0_15px_rgba(244,63,94,0.4)]" : "bg-surface-800 text-surface-400"}`}
                >
                    1
                </div>
                <h3 class="text-xl font-bold text-white">Upload Target PDF</h3>
            </div>

            {#if !pdfFile}
                <div
                    in:fade
                    role="button"
                    tabindex="0"
                    on:dragover|preventDefault
                    on:drop={handlePdfDrop}
                    class="glass-surface relative w-full h-48 border-2 border-dashed border-surface-700 flex flex-col items-center justify-center transition-all hover:border-primary-500/50 cursor-pointer focus:outline-none focus:ring-2 ring-primary-500/50"
                >
                    <input
                        type="file"
                        id="pdfUpload"
                        class="hidden"
                        accept=".pdf"
                        on:change={handlePdfSelect}
                    />
                    <label
                        for="pdfUpload"
                        class="absolute inset-0 cursor-pointer"
                    ></label>
                    <div class="glass-icon mb-4">
                        <UploadSimple
                            weight="duotone"
                            class="w-7 h-7 text-primary-400"
                        />
                    </div>
                    <p class="font-bold text-white mb-1">
                        Drag & Drop PDF Document
                    </p>
                    <p class="text-xs text-surface-500 font-medium">
                        or click here to browse
                    </p>
                </div>
            {:else}
                <div
                    in:scale={{ duration: 300, start: 0.95 }}
                    class="bg-surface-950 border border-surface-700 p-5 rounded-2xl flex items-center justify-between shadow-inner"
                >
                    <div class="flex items-center gap-4">
                        <div
                            class="p-3 bg-red-500/10 rounded-xl border border-red-500/20"
                        >
                            <FilePdf
                                weight="duotone"
                                class="w-6 h-6 text-red-500"
                            />
                        </div>
                        <div>
                            <p class="font-bold text-white mb-0.5">
                                {pdfFile.name}
                            </p>
                            <p
                                class="text-xs text-surface-500 font-medium tracking-wide"
                            >
                                {(pdfFile.size / 1024 / 1024).toFixed(2)} MB PDF
                                Document
                            </p>
                        </div>
                    </div>
                    <button
                        class="text-surface-500 hover:text-red-400 p-2 hover:bg-surface-800 rounded-lg transition-colors"
                        on:click={clearFile}
                        title="Remove File"
                    >
                        <XCircle weight="bold" class="w-5 h-5" />
                    </button>
                </div>
            {/if}
        </div>
    </div>

    <!-- STEP 2: PROCESSING OPTIONS -->
    {#if step >= 2}
        <div
            in:slide={{ duration: 400, delay: 100 }}
            class={`transition-all duration-500`}
        >
            <div class="glass-card p-8 relative overflow-hidden flex flex-col">
                <div class="flex items-center gap-4 mb-6 relative z-10">
                    <div
                        class="w-8 h-8 rounded-full flex items-center justify-center font-bold text-sm bg-purple-500 text-white shadow-[0_0_15px_rgba(168,85,247,0.4)]"
                    >
                        2
                    </div>
                    <h3 class="text-xl font-bold text-white">
                        Extraction Options
                    </h3>
                </div>

                <div class="grid sm:grid-cols-2 gap-4 flex-1">
                    <!-- Option 1: All Text -->
                    <button
                        class={`glass-surface flex items-start gap-4 p-5 border-2 transition-all text-left ${processingOption === "allText" ? "!border-purple-500/50 text-purple-400" : "border-transparent hover:border-white/10 text-surface-400"}`}
                        on:click={() => (processingOption = "allText")}
                    >
                        <div
                            class={`p-2 rounded-lg mt-0.5 ${processingOption === "allText" ? "bg-purple-500/20" : "bg-surface-800"}`}
                        >
                            <TextT
                                weight={processingOption === "allText"
                                    ? "fill"
                                    : "duotone"}
                                class="w-6 h-6"
                            />
                        </div>
                        <div>
                            <h4
                                class={`font-bold text-lg mb-1 transition-colors ${processingOption === "allText" ? "text-white" : "text-surface-300"}`}
                            >
                                Full Extraction
                            </h4>
                            <p
                                class="text-sm opacity-80 leading-relaxed font-medium"
                            >
                                Extracts tables and unstructured text. Attempts
                                to align text as columns alongside tables.
                            </p>
                        </div>
                    </button>

                    <!-- Option 2: Tables Only -->
                    <button
                        class={`glass-surface flex items-start gap-4 p-5 border-2 transition-all text-left ${processingOption === "tablesOnly" ? "!border-blue-500/50 text-blue-400" : "border-transparent hover:border-white/10 text-surface-400"}`}
                        on:click={() => (processingOption = "tablesOnly")}
                    >
                        <div
                            class={`p-2 rounded-lg mt-0.5 ${processingOption === "tablesOnly" ? "bg-blue-500/20" : "bg-surface-800"}`}
                        >
                            <Table
                                weight={processingOption === "tablesOnly"
                                    ? "fill"
                                    : "duotone"}
                                class="w-6 h-6"
                            />
                        </div>
                        <div>
                            <h4
                                class={`font-bold text-lg mb-1 transition-colors ${processingOption === "tablesOnly" ? "text-white" : "text-surface-300"}`}
                            >
                                Tables Only (Strict)
                            </h4>
                            <p
                                class="text-sm opacity-80 leading-relaxed font-medium"
                            >
                                Uses lattice and stream analysis to find strict
                                grid structures. Ignores loose paragraphs.
                            </p>
                        </div>
                    </button>
                </div>

                <!-- Format Toggle -->
                <div class="mt-6">
                    <h4
                        class="text-sm font-bold text-surface-400 mb-3 uppercase tracking-wider"
                    >
                        Export Format
                    </h4>
                    <div class="flex gap-3">
                        <button
                            class={`glass-surface flex-1 py-3 px-4 border-2 font-bold transition-all flex items-center justify-center gap-2 ${outputFormat === "excel" ? "!border-green-500/50 text-green-400" : "border-transparent text-surface-400 hover:border-white/10"}`}
                            on:click={() => (outputFormat = "excel")}
                        >
                            <Table weight="fill" class="w-5 h-5" />
                            Excel (.xlsx)
                        </button>
                        <button
                            class={`glass-surface flex-1 py-3 px-4 border-2 font-bold transition-all flex items-center justify-center gap-2 ${outputFormat === "csv" ? "!border-orange-500/50 text-orange-400" : "border-transparent text-surface-400 hover:border-white/10"}`}
                            on:click={() => (outputFormat = "csv")}
                        >
                            <TextT weight="fill" class="w-5 h-5" />
                            CSV (.csv)
                        </button>
                    </div>
                </div>

                <div class="mt-8 pt-6 border-t border-surface-800/80">
                    <button
                        disabled={isUploading}
                        class={`group relative flex items-center justify-center gap-3 w-full py-4 text-white font-bold rounded-2xl overflow-hidden shadow-lg transition-all disabled:opacity-50 disabled:shadow-none bg-gradient-to-r ${processingOption === "allText" ? "from-purple-600 to-primary-600 shadow-purple-500/20 hover:shadow-purple-500/40" : "from-blue-600 to-indigo-600 shadow-blue-500/20 hover:shadow-blue-500/40"}`}
                        on:click={handleExtract}
                    >
                        {#if isUploading}
                            <Spinner class="w-5 h-5 animate-spin" />
                            <span>Processing PDF Engine...</span>
                        {:else}
                            <span
                                >Extract to {outputFormat === "excel"
                                    ? "Microsoft Excel"
                                    : "CSV"}</span
                            >
                            <ArrowRight
                                weight="bold"
                                class="w-5 h-5 group-hover:translate-x-1 transition-transform"
                            />
                        {/if}
                    </button>
                </div>
            </div>
        </div>
    {/if}
</div>
