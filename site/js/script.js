document.addEventListener('DOMContentLoaded', () => {
    const $ = id => document.getElementById(id);
    const stage = $('stage'), slide = $('slide'), fileInput = $('fileInput');
    const panels = [...document.querySelectorAll('.panel')];
    let selectedFile = null, resultBlob = null, resultName = '';

    const messages = [
        'Reading PPTX package...',
        'Mapping slides, layouts and masters...',
        'Fingerprinting embedded images...',
        'Hunting the Gamma badge...',
        'Erasing safe traces...',
        'Rebuilding your deck...'
    ];

    const sleep = ms => new Promise(r => setTimeout(r, ms));
    const fmt = n => n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';

    function setState(s) {
        stage.dataset.state = s;
        panels.forEach(p => { p.hidden = !p.dataset.show.split(' ').includes(s); });
    }

    function reset() {
        selectedFile = null; resultBlob = null; fileInput.value = '';
        setState('idle');
    }

    function pick(file) {
        if (!file) return;
        if (!/\.pptx$/i.test(file.name)) {
            $('errText').textContent = 'That is not a .pptx file. Export your Gamma deck as PowerPoint first.';
            setState('error');
            return;
        }
        selectedFile = file;
        $('fileName').textContent = file.name;
        $('fileSize').textContent = fmt(file.size);
        setState('ready');
    }

    slide.addEventListener('click', () => { if (stage.dataset.state === 'idle') fileInput.click(); });
    slide.addEventListener('keydown', e => {
        if ((e.key === 'Enter' || e.key === ' ') && stage.dataset.state === 'idle') { e.preventDefault(); fileInput.click(); }
    });
    $('browseBtn').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', e => pick(e.target.files[0]));
    $('clearBtn').addEventListener('click', reset);
    $('anotherBtn').addEventListener('click', reset);
    $('emptyBtn').addEventListener('click', reset);
    $('errBtn').addEventListener('click', reset);
    $('againDl').addEventListener('click', () => { if (resultBlob) saveAs(resultBlob, resultName); });

    // Accept a drop anywhere on the page so a missed target never opens the file in the browser.
    window.addEventListener('dragover', e => { e.preventDefault(); stage.classList.add('drag'); });
    window.addEventListener('dragleave', e => { if (!e.relatedTarget) stage.classList.remove('drag'); });
    window.addEventListener('drop', e => {
        e.preventDefault();
        stage.classList.remove('drag');
        if (['idle', 'ready', 'empty', 'error', 'done'].includes(stage.dataset.state) && e.dataTransfer.files.length) pick(e.dataTransfer.files[0]);
    });

    $('goBtn').addEventListener('click', async () => {
        if (!selectedFile) return;
        setState('scanning');
        const bar = $('barFill'), text = $('statusText');
        const started = Date.now();
        let done = false, result = null, error = null;

        processPPTX().then(r => { result = r; }).catch(e => { error = e; }).finally(() => { done = true; });

        let i = 0, pct = 4;
        while (!done || Date.now() - started < 2400) {
            text.textContent = messages[Math.min(i, messages.length - 1)];
            pct = Math.min(94, pct + 7);
            bar.style.width = pct + '%';
            i++;
            await sleep(400);
        }

        if (error) {
            console.error(error);
            $('errText').textContent = 'Could not read this file: ' + error.message;
            setState('error');
            return;
        }
        if (result.status === 'no_watermarks') { setState('empty'); return; }

        bar.style.width = '100%';
        resultBlob = result.blob;
        resultName = selectedFile.name.replace(/\.pptx$/i, '-cleaned.pptx');
        $('stats').innerHTML = [
            result.removedImages + ' image' + (result.removedImages === 1 ? '' : 's') + ' removed',
            result.removedShapes + ' trace' + (result.removedShapes === 1 ? '' : 's') + ' cleaned',
            fmt(selectedFile.size) + ' to ' + fmt(result.blob.size)
        ].map(t => '<span>' + t + '</span>').join('');
        setState('done');
        saveAs(resultBlob, resultName);
    });

    // Live GitHub star count (only a public API call, your file is never involved).
    fetch('https://api.github.com/repos/Shariqtechie/gamma-watermark-remover')
        .then(r => r.ok ? r.json() : null)
        .then(d => {
            if (!d || typeof d.stargazers_count !== 'number') return;
            document.querySelectorAll('[data-stars]').forEach(el => { el.textContent = d.stargazers_count.toLocaleString(); });
        })
        .catch(() => {});

    setState('idle');

    function localName(node) {
        return (node.localName || node.nodeName || '').toLowerCase();
    }

    function getAttr(el, names) {
        for (const name of names) {
            const value = el.getAttribute(name);
            if (value !== null && value !== '') return value;
        }
        return null;
    }

    function parentXmlPathFromRels(relPath) {
        return relPath.replace('/_rels/', '/').replace(/\.rels$/i, '');
    }

    function normalizeTarget(relPath, target) {
        if (!target) return '';

        if (target.startsWith('/')) {
            return target.slice(1).replace(/\\/g, '/');
        }

        let base = relPath.split('/');
        base.pop();

        if (base[base.length - 1] === '_rels') {
            base.pop();
        }

        const parts = [...base, ...target.replace(/\\/g, '/').split('/')];
        const clean = [];

        for (const part of parts) {
            if (!part || part === '.') continue;
            if (part === '..') clean.pop();
            else clean.push(part);
        }

        return clean.join('/');
    }

    async function sha1Hex(uint8) {
        const digest = await crypto.subtle.digest('SHA-1', uint8);
        return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, '0')).join('');
    }

    function findImageContainersByRelId(xmlDoc, relId) {
        const containers = [];
        const all = Array.from(xmlDoc.getElementsByTagName('*'));

        for (const el of all) {
            const tag = localName(el);
            if (tag !== 'blip' && tag !== 'svgblip') continue;

            const embed = getAttr(el, ['r:embed', 'embed']);
            const link = getAttr(el, ['r:link', 'link']);

            if (embed !== relId && link !== relId) continue;

            let parent = el;
            while (parent && parent.parentNode) {
                const pTag = localName(parent);
                if (['pic', 'sp', 'graphicframe', 'bg', 'grpSp'.toLowerCase()].includes(pTag)) break;
                parent = parent.parentNode;
            }

            if (parent && parent.parentNode && !containers.includes(parent)) {
                containers.push(parent);
            }
        }

        return containers;
    }

    function getContainerInfo(container) {
        const info = {
            x: null,
            y: null,
            cx: null,
            cy: null,
            hasHyperlink: false,
            shapeName: '',
            shapeId: ''
        };

        const all = Array.from(container.getElementsByTagName('*'));

        for (const el of all) {
            const tag = localName(el);

            if (tag === 'off') {
                info.x = Number(el.getAttribute('x'));
                info.y = Number(el.getAttribute('y'));
            }

            if (tag === 'ext') {
                const cx = Number(el.getAttribute('cx'));
                const cy = Number(el.getAttribute('cy'));
                if (Number.isFinite(cx) && Number.isFinite(cy)) {
                    info.cx = cx;
                    info.cy = cy;
                    break;
                }
            }
        }

        for (const el of all) {
            const tag = localName(el);
            if (tag === 'hlinkclick' || tag === 'hlinkhover') {
                info.hasHyperlink = true;
            }
            if (tag === 'cnvpr') {
                info.shapeName = el.getAttribute('name') || '';
                info.shapeId = el.getAttribute('id') || '';
            }
        }

        return info;
    }

    function isFullSlideLike(trace) {
        const slideW = 14630400;
        const slideH = 8229600;

        if (!Number.isFinite(trace.cx) || !Number.isFinite(trace.cy)) return false;

        return trace.x === 0 && trace.y === 0 && trace.cx >= slideW * 0.85 && trace.cy >= slideH * 0.85;
    }

    function isBottomRightLike(trace) {
        const slideW = 14630400;
        const slideH = 8229600;

        if (!Number.isFinite(trace.x) || !Number.isFinite(trace.y)) return false;

        return trace.x > slideW * 0.65 && trace.y > slideH * 0.70;
    }

    function isWatermarkLikeTrace(trace, parentXmlPath) {
        if (isFullSlideLike(trace)) return false;

        const inTemplate = /ppt\/(slideLayouts|slideMasters)\//i.test(parentXmlPath);
        const smallShape = Number.isFinite(trace.cx) && Number.isFinite(trace.cy)
            ? trace.cx < 4000000 && trace.cy < 1200000
            : true;

        return (
            trace.hasHyperlink ||
            isBottomRightLike(trace) ||
            (inTemplate && smallShape)
        );
    }

    function removeContainers(xmlDoc, relIds) {
        let removed = 0;

        for (const relId of relIds) {
            const containers = findImageContainersByRelId(xmlDoc, relId);
            for (const container of containers) {
                if (container.parentNode) {
                    container.parentNode.removeChild(container);
                    removed++;
                }
            }
        }

        return removed;
    }

    function removeRelationshipNodes(relXml, relIds) {
        let removed = 0;
        const rels = Array.from(relXml.getElementsByTagName('Relationship'));

        for (const rel of rels) {
            const id = rel.getAttribute('Id');
            if (relIds.has(id) && rel.parentNode) {
                rel.parentNode.removeChild(rel);
                removed++;
            }
        }

        return removed;
    }

    async function processPPTX() {
        const inputBuffer = await selectedFile.arrayBuffer();
        const zip = await JSZip.loadAsync(inputBuffer);

        const zipPaths = Object.keys(zip.files);
        const mediaFiles = zipPaths.filter(path =>
            /^ppt\/media\//i.test(path) && /\.(png|jpe?g|gif|webp|svg|emf|wmf)$/i.test(path)
        );
        const relFiles = zipPaths.filter(path => /\.xml\.rels$/i.test(path));

        const mediaMap = new Map();
        const mediaByHash = new Map();

        for (const path of mediaFiles) {
            const bytes = await zip.files[path].async('uint8array');
            const hash = await sha1Hex(bytes);
            const item = {
                path,
                name: path.split('/').pop(),
                size: bytes.length,
                hash,
                refs: []
            };

            mediaMap.set(path, item);
            if (!mediaByHash.has(hash)) mediaByHash.set(hash, []);
            mediaByHash.get(hash).push(item);
        }

        const relationships = [];

        for (const relPath of relFiles) {
            const relText = await zip.files[relPath].async('text');
            const relXml = new DOMParser().parseFromString(relText, 'application/xml');
            const rels = Array.from(relXml.getElementsByTagName('Relationship'));
            const parentXmlPath = parentXmlPathFromRels(relPath);
            const parentXmlExists = Boolean(zip.files[parentXmlPath]);
            let parentXmlDoc = null;

            if (parentXmlExists) {
                const parentXmlText = await zip.files[parentXmlPath].async('text');
                parentXmlDoc = new DOMParser().parseFromString(parentXmlText, 'application/xml');
            }

            for (const rel of rels) {
                const id = rel.getAttribute('Id');
                const target = rel.getAttribute('Target') || '';
                const normalizedTarget = normalizeTarget(relPath, target);
                const media = mediaMap.get(normalizedTarget);

                if (!media) continue;

                const containers = parentXmlDoc ? findImageContainersByRelId(parentXmlDoc, id) : [];
                const traces = containers.map(container => getContainerInfo(container));

                const ref = {
                    relPath,
                    parentXmlPath,
                    id,
                    target,
                    mediaPath: normalizedTarget,
                    traces
                };

                media.refs.push(ref);
                relationships.push(ref);
            }
        }

        const candidateMediaPaths = new Set();
        const candidateHashes = [];

        for (const [hash, items] of mediaByHash.entries()) {
            const totalRefs = items.reduce((sum, item) => sum + item.refs.length, 0);
            const totalSizeOk = items.every(item => item.size >= 500 && item.size <= 350000);
            const appearsRepeatedByHash = items.length >= 3 || totalRefs >= 3;

            if (!totalSizeOk || !appearsRepeatedByHash) continue;

            let safeWatermarkTraceCount = 0;
            let fullSlideTraceCount = 0;
            let templateRefCount = 0;

            for (const item of items) {
                for (const ref of item.refs) {
                    if (/ppt\/(slideLayouts|slideMasters)\//i.test(ref.parentXmlPath)) {
                        templateRefCount++;
                    }

                    if (ref.traces.length === 0 && /\.svg$/i.test(item.path)) {
                        safeWatermarkTraceCount++;
                    }

                    for (const trace of ref.traces) {
                        if (isFullSlideLike(trace)) fullSlideTraceCount++;
                        if (isWatermarkLikeTrace(trace, ref.parentXmlPath)) safeWatermarkTraceCount++;
                    }
                }
            }

            const isCandidate = safeWatermarkTraceCount > 0 && fullSlideTraceCount === 0 && templateRefCount > 0;

            if (isCandidate) {
                candidateHashes.push(hash);
                for (const item of items) candidateMediaPaths.add(item.path);
            }
        }

        // Fallback for exports that use unique files but identical bottom-right layout placement.
        if (candidateMediaPaths.size === 0) {
            for (const media of mediaMap.values()) {
                if (media.size < 500 || media.size > 350000) continue;

                for (const ref of media.refs) {
                    if (!/ppt\/(slideLayouts|slideMasters)\//i.test(ref.parentXmlPath)) continue;
                    if (ref.traces.some(trace => isWatermarkLikeTrace(trace, ref.parentXmlPath))) {
                        candidateMediaPaths.add(media.path);
                    }
                }
            }
        }

        if (candidateMediaPaths.size === 0) {
            console.info('No safe watermark candidates found.', { mediaMap, relationships });
            return { status: 'no_watermarks' };
        }

        console.info('Watermark candidate image paths:', Array.from(candidateMediaPaths));
        console.info('Watermark candidate hashes:', candidateHashes);

        const relIdsByRelFile = new Map();
        const relIdsByParentXml = new Map();

        for (const ref of relationships) {
            if (!candidateMediaPaths.has(ref.mediaPath)) continue;

            if (!relIdsByRelFile.has(ref.relPath)) relIdsByRelFile.set(ref.relPath, new Set());
            relIdsByRelFile.get(ref.relPath).add(ref.id);

            if (!relIdsByParentXml.has(ref.parentXmlPath)) relIdsByParentXml.set(ref.parentXmlPath, new Set());
            relIdsByParentXml.get(ref.parentXmlPath).add(ref.id);
        }

        let removedShapes = 0;
        let removedRelationships = 0;

        for (const [parentXmlPath, relIds] of relIdsByParentXml.entries()) {
            if (!zip.files[parentXmlPath]) continue;

            const xmlText = await zip.files[parentXmlPath].async('text');
            const xmlDoc = new DOMParser().parseFromString(xmlText, 'application/xml');
            removedShapes += removeContainers(xmlDoc, relIds);
            zip.file(parentXmlPath, new XMLSerializer().serializeToString(xmlDoc));
        }

        for (const [relPath, relIds] of relIdsByRelFile.entries()) {
            if (!zip.files[relPath]) continue;

            const relText = await zip.files[relPath].async('text');
            const relXml = new DOMParser().parseFromString(relText, 'application/xml');
            removedRelationships += removeRelationshipNodes(relXml, relIds);
            zip.file(relPath, new XMLSerializer().serializeToString(relXml));
        }

        for (const mediaPath of candidateMediaPaths) {
            if (zip.files[mediaPath]) {
                zip.remove(mediaPath);
            }
        }

        const blob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });

        return {
            status: 'cleaned',
            blob,
            removedImages: candidateMediaPaths.size,
            removedShapes,
            removedRelationships
        };
    }
});
