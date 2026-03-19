document.addEventListener('DOMContentLoaded', () => {
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const fileNameDisplay = document.getElementById('fileName');
    const removeFileBtn = document.getElementById('removeFile');
    const processBtn = document.getElementById('processBtn');

    const progressArea = document.getElementById('progressArea');
    const funMessage = document.getElementById('funMessage');
    const progressBar = document.getElementById('progressBar');

    let selectedFile = null;

    // Fun messages to display while "processing" to make it look like a 50-person team effort
    const funnyMessages = [
        "Waking up the hamsters in the server room...",
        "Applying quantum entanglement to your PPTX...",
        "Hunting down filthy watermarks...",
        "Executing Operation: Clean Slate...",
        "Bribing MS PowerPoint with virtual coffee...",
        "Nuking embedded images from orbit...",
        "Scrubbing slides with digital soap...",
        "Deleting the evidence...",
        "Sprinkling fairy dust on XML schemas...",
        "Reticulating splines...",
        "Zipping it all back together...",
        "Baking the final presentation..."
    ];

    // --- DOM Event Listeners ---

    // Click on dropzone opens native file picker
    dropzone.addEventListener('click', () => fileInput.click());

    // File input change
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });

    // Drag and Drop Events
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
    });

    // Remove selected file
    removeFileBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        selectedFile = null;
        fileInput.value = '';
        fileInfo.classList.remove('active');
        fileInfo.style.display = ''; // Clear any inline non-flex styles
        dropzone.style.display = 'block';
        processBtn.textContent = 'Clean Presentation';
        processBtn.disabled = true;
        progressArea.classList.remove('active');
    });

    // Process button click
    processBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        // UI transitions
        dropzone.style.display = 'none';
        fileInfo.style.display = 'none';
        processBtn.style.display = 'none';
        progressArea.classList.add('active');

        await startFakeProgressAndProcess();
    });

    // --- Core Functions ---

    function handleFileSelect(file) {
        if (!file.name.toLowerCase().endsWith('.pptx')) {
            alert("Whoops! We only eat .pptx files here. Try again!");
            return;
        }
        selectedFile = file;
        fileNameDisplay.textContent = file.name;
        dropzone.style.display = 'none';
        fileInfo.classList.add('active');
        processBtn.disabled = false;
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    async function startFakeProgressAndProcess() {
        // We will run the fake progress bar concurrently with the actual processing.
        // The actual zip processing is usually instant, but we want to show off the UI.

        let progress = 0;
        let isProcessingDone = false;
        let pResult = null;
        let pError = null;

        // Start actual processing in background
        processPPTX().then(blob => {
            pResult = blob;
            isProcessingDone = true;
        }).catch(err => {
            pError = err;
            isProcessingDone = true;
        });

        // UI "Fake" complex loading sequences
        for (let i = 0; i < funnyMessages.length; i++) {
            if (isProcessingDone && i > 3 && progress >= 80) {
                // Wait for the bar to at least show something before jumping to 100%
                break;
            }

            funMessage.textContent = funnyMessages[i];

            // Random chunk of progress
            let targetProgress = Math.min(progress + (Math.random() * 15 + 5), 95);

            // Increment progress smoothly over UI ticks
            while (progress < targetProgress) {
                progress += 1.5;
                if (progress > targetProgress) progress = targetProgress;
                progressBar.style.width = `${progress}%`;
                await sleep(30);
            }

            await sleep(400 + Math.random() * 600); // Wait on the message text
        }

        // Wait for actual work to finish if it somehow took longer than our fake animation
        while (!isProcessingDone) {
            funMessage.textContent = "Almost there, convincing the bits to align...";
            await sleep(500);
        }

        if (pError) {
            funMessage.textContent = "Mission Failed: " + pError.message;
            funMessage.style.color = "var(--error-color)";
            progressBar.style.background = "var(--error-color)";
            processBtn.style.display = 'block';
            processBtn.textContent = 'Retry';
            processBtn.disabled = false;
            return;
        }

        if (pResult === 'no_watermarks') {
            funMessage.textContent = "Clean! No watermarks detected! 🎉";
            progressBar.style.width = `100%`;
            progressBar.style.background = "var(--success-color)";

            // Reset UI after a few seconds
            setTimeout(() => {
                progressArea.classList.remove('active');
                fileInfo.style.display = 'flex';
                processBtn.style.display = 'block';
                processBtn.textContent = 'Clean Presentation';
                progressBar.style.width = '0%';
                progressBar.style.background = '';
                funMessage.style.color = '';
            }, 3000);
            return;
        }

        // Processing success
        funMessage.textContent = "Done! Downloading masterpiece... 🚀";
        funMessage.style.color = "var(--success-color)";
        progressBar.style.width = `100%`;
        progressBar.style.background = "var(--success-color)";

        // Download file
        saveAs(pResult, selectedFile.name.replace('.pptx', '-cleaned.pptx'));

        // Reset UI eventually
        setTimeout(() => {
            progressArea.classList.remove('active');
            dropzone.style.display = 'block';
            fileInput.value = '';
            selectedFile = null;
            processBtn.style.display = 'block';
            processBtn.disabled = true;
            processBtn.textContent = 'Clean Presentation';
            progressBar.style.width = '0%';
            progressBar.style.background = '';
            funMessage.style.color = '';
        }, 3000);
    }

    // --- Original Core Logic Overhauled for robustness ---
    async function processPPTX() {
        const zip = await JSZip.loadAsync(await selectedFile.arrayBuffer());

        const mediaFiles = Object.keys(zip.files)
            .filter(n => /^ppt\/media\//i.test(n) && /\.(png|jpe?g|gif|webp|svg)$/i.test(n));

        const relFiles = Object.keys(zip.files).filter(n => n.endsWith('.xml.rels'));

        const usageCount = {};

        for (const relPath of relFiles) {
            const txt = await zip.files[relPath].async('text');
            const relXml = new DOMParser().parseFromString(txt, 'application/xml');
            const rels = relXml.getElementsByTagName('Relationship');

            for (let i = 0; i < rels.length; i++) {
                const target = rels[i].getAttribute('Target') || '';
                for (const file of mediaFiles) {
                    const name = file.split('/').pop();
                    if (target.includes(name)) {
                        usageCount[name] = (usageCount[name] || 0) + 1;
                    }
                }
            }
        }

        const candidates = [];

        for (const path of mediaFiles) {
            const arr = await zip.files[path].async('uint8array');
            const size = arr.length;
            const name = path.split('/').pop();

            if (size >= 500 && size <= 100000 && (usageCount[name] || 0) >= 3) {
                candidates.push({ path, name });
            }
        }

        if (candidates.length === 0) {
            return 'no_watermarks';
        }

        function removeAllImageRefs(xmlDoc, relId) {
            let removed = 0;
            const all = xmlDoc.getElementsByTagName('*');

            for (let i = 0; i < all.length; i++) {
                const el = all[i];
                const tag = (el.localName || '').toLowerCase();

                if (tag === 'blip') {
                    const embed = el.getAttribute('r:embed') || el.getAttribute('embed');
                    if (embed === relId) {
                        let parent = el;
                        while (parent && parent.parentNode) {
                            const pTag = (parent.localName || '').toLowerCase();
                            if (['pic', 'sp', 'graphicframe', 'bg'].includes(pTag)) break;
                            parent = parent.parentNode;
                        }
                        if (parent && parent.parentNode) {
                            parent.parentNode.removeChild(parent);
                            removed++;
                        }
                    }
                }
            }

            return removed;
        }

        function nukeBrokenImageContainers(xmlDoc) {
            let removed = 0;
            const all = xmlDoc.getElementsByTagName('*');

            for (let i = 0; i < all.length; i++) {
                const el = all[i];
                const tag = (el.localName || '').toLowerCase();

                if (tag === 'pic') {
                    const blips = el.getElementsByTagName('*');
                    let hasValidImage = false;

                    for (let j = 0; j < blips.length; j++) {
                        const b = blips[j];
                        if ((b.localName || '').toLowerCase() === 'blip') {
                            const embed = b.getAttribute('r:embed');
                            if (embed && embed.trim() !== '') {
                                hasValidImage = true;
                                break;
                            }
                        }
                    }

                    if (!hasValidImage) {
                        if (el.parentNode) {
                            el.parentNode.removeChild(el);
                            removed++;
                        }
                    }
                }
            }

            return removed;
        }

        for (const c of candidates) {
            for (const relPath of relFiles) {
                const txt = await zip.files[relPath].async('text');
                const relXml = new DOMParser().parseFromString(txt, 'application/xml');
                const rels = relXml.getElementsByTagName('Relationship');

                for (let i = rels.length - 1; i >= 0; i--) {
                    const r = rels[i];
                    const target = r.getAttribute('Target') || '';

                    if (target.includes(c.name)) {
                        const relId = r.getAttribute('Id');

                        const parentXmlPath = relPath
                            .replace('/_rels', '')
                            .replace('.rels', '');

                        if (zip.files[parentXmlPath]) {
                            const xmlTxt = await zip.files[parentXmlPath].async('text');
                            const xmlDoc = new DOMParser().parseFromString(xmlTxt, 'application/xml');

                            removeAllImageRefs(xmlDoc, relId);
                            nukeBrokenImageContainers(xmlDoc);

                            zip.file(parentXmlPath, new XMLSerializer().serializeToString(xmlDoc));
                        }

                        r.parentNode.removeChild(r);
                    }
                }

                zip.file(relPath, new XMLSerializer().serializeToString(relXml));
            }

            zip.remove(c.path);
        }

        return await zip.generateAsync({ type: 'blob' });
    }
});
