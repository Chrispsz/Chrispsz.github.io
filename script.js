'use strict';

// Garante que o código principal só execute após o DOM e os scripts externos estarem prontos.
document.addEventListener('DOMContentLoaded', () => {
    // Verificação de segurança: se a biblioteca externa não carregou, para a execução.
    if (typeof XLSX === 'undefined') {
        document.body.innerHTML = '<div style="text-align: center; padding: 2rem; font-family: sans-serif; color: #dc2626;"><h1>Erro Crítico</h1><p>Não foi possível carregar a biblioteca de planilhas (XLSX). Por favor, verifique sua conexão com a internet, desative bloqueadores de anúncio e tente recarregar a página.</p></div>';
        return;
    }
    
    // Se a biblioteca carregou, a aplicação começa.
    const state = { workbook: null, fileName: null, settings: { targetColumn: 'F', colorMap: {}, unirPares: [], aplicarMascara: true, dedupeResults: true, sortResults: false }, rawResultsText: '', activeModalTrigger: null, lastSettingsBeforeReset: null, colorToAdd: null, currentSheetData: [], undoStack: [] };
    const DOM = {
        doc: document.documentElement,
        uploadArea: document.getElementById('upload-area'),
        fileInput: document.getElementById('file-input'),
        sheetSelectorArea: document.getElementById('sheet-selector-area'),
        sheetSelector: document.getElementById('sheet-selector'),
        resultsArea: document.getElementById('results-area'),
        resultsTitle: document.getElementById('results-title'),
        generatorOutput: document.getElementById('generator-output'),
        diagnosticoContainer: document.getElementById('diagnostico-container'),
        clearBtn: document.getElementById('clear-btn'),
        copyBtn: document.getElementById('copy-btn'),
        copyCsvBtn: document.getElementById('copy-csv-btn'),
        downloadBtn: document.getElementById('download-btn'),
        downloadSheetBtn: document.getElementById('download-sheet-btn'),
        settingsModal: document.getElementById('settings-modal'),
        settingsBtn: document.getElementById('settings-btn'),
        targetColumnInput: document.getElementById('target-column-input'),
        targetColumnSelect: document.getElementById('target-column-select'),
        dedupeResultsCheck: document.getElementById('dedupe-results-check'),
        sortResultsCheck: document.getElementById('sort-results-check'),
        colorRulesList: document.getElementById('color-rules-list'),
        ruleSearchInput: document.getElementById('rule-search-input'),
        resetSettingsBtn: document.getElementById('reset-settings-btn'),
        exportSettingsBtn: document.getElementById('export-settings-btn'),
        importSettingsBtn: document.getElementById('import-settings-btn'),
        importSettingsInput: document.getElementById('import-settings-input'),
        themeToggleBtn: document.getElementById('theme-toggle-btn'),
        liveFeedback: document.getElementById('live-feedback'),
        toastContainer: document.getElementById('toast-container'),
        addRuleFormContainer: document.getElementById('add-rule-form-container'),
        templates: { rule: document.getElementById('rule-template'), addRuleForm: document.getElementById('add-rule-form-template') },
        unirTelefonesModal: document.getElementById('unir-telefones-modal'),
        paresContainer: document.getElementById('pares-container'),
        addParBtn: document.getElementById('add-par-btn'),
        confirmarUnirBtn: document.getElementById('confirmar-unir-btn'),
        aplicarMascaraCheck: document.getElementById('aplicar-mascara-br-check'),
        resetParesBtn: document.getElementById('reset-pares-btn'),
        undoBtn: document.getElementById('undo-btn'),
    };

    function esc(str = '') { return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
    
    function manageFocus(modal, open) {
        const focusableElements = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';
        const focusableContent = modal.querySelectorAll(focusableElements);
        if (focusableContent.length === 0) return;
        const firstFocusable = focusableContent[0];
        const lastFocusable = focusableContent[focusableContent.length - 1];

        const trapFocus = (e) => {
            let isTabPressed = e.key === 'Tab' || e.keyCode === 9;
            if (!isTabPressed) return;

            if (e.shiftKey) { 
                if (document.activeElement === firstFocusable) {
                    lastFocusable.focus();
                    e.preventDefault();
                }
            } else {
                if (document.activeElement === lastFocusable) {
                    firstFocusable.focus();
                    e.preventDefault();
                }
            }
        };

        if (open) {
            modal.addEventListener('keydown', trapFocus);
        } else {
            modal.removeEventListener('keydown', trapFocus);
        }
    }

    const ui = {
        setLoading(isLoading) {
            const content = DOM.uploadArea.querySelector('#upload-content');
            if (content) content.classList.toggle('hidden', isLoading);
            const existingSpinner = DOM.uploadArea.querySelector('.loading-spinner');
            if (isLoading && !existingSpinner) {
                const spinner = document.createElement('div');
                spinner.className = 'loading-spinner';
                DOM.uploadArea.appendChild(spinner);
            } else if (!isLoading && existingSpinner) {
                existingSpinner.remove();
            }
        },
        renderResults({ fileName, html, raw, hasContent, skippedRows, unmappedColors }) {
            document.title = `${fileName.substring(0, fileName.lastIndexOf('.')) || fileName} - Planilista`;
            state.rawResultsText = raw;
            DOM.generatorOutput.innerHTML = html || this.getEmptyStateHTML('no_results');
            let diagnosticoHTML = '';
            if (unmappedColors.size > 0) {
                let unmappedItemsHTML = '';
                unmappedColors.forEach(color => {
                    unmappedItemsHTML += `<div class="unmapped-color-item"><div class="unmapped-color-info"><div class="color-swatch-small" style="background-color:${color}"></div><span>${color}</span></div><button class="btn btn-small" data-configure-color="${color}">Configurar</button></div>`;
                });
                diagnosticoHTML += `<div class="diagnostico-aviso"><strong>Ação necessária:</strong> Encontramos ${unmappedColors.size} cor(es) não configurada(s).<br>${unmappedItemsHTML}</div>`;
            }
            if (skippedRows.length > 0) diagnosticoHTML += `<div class="diagnostico-aviso"><strong>Aviso:</strong> ${skippedRows.length} linha(s) ignorada(s) por não conterem números na coluna alvo.</div>`;
            DOM.diagnosticoContainer.innerHTML = diagnosticoHTML;
            DOM.resultsArea.classList.add('visible');
            const hasResults = hasContent || raw.length > 0;
            DOM.copyBtn.disabled = !hasResults;
            DOM.copyCsvBtn.disabled = !hasResults;
            DOM.downloadBtn.disabled = !hasResults;
            DOM.downloadSheetBtn.disabled = false;
            DOM.undoBtn.disabled = state.undoStack.length === 0;
            this.toggleSettingsAlert(unmappedColors.size > 0);
        },
        renderColorRules(colorMap) {
            const fragment = document.createDocumentFragment();
            if (Object.keys(colorMap).length === 0) {
                DOM.colorRulesList.innerHTML = this.getEmptyStateHTML('no_rules');
                return;
            }
            for (const color in colorMap) {
                const ruleDiv = DOM.templates.rule.content.cloneNode(true).firstElementChild;
                ruleDiv.dataset.color = color;
                ruleDiv.querySelector('.color-swatch-large').style.backgroundColor = color;
                ruleDiv.querySelector('input[type="text"]').value = colorMap[color];
                fragment.appendChild(ruleDiv);
            }
            DOM.colorRulesList.innerHTML = '';
            DOM.colorRulesList.appendChild(fragment);
        },
        getEmptyStateHTML(stateKey) {
            const states = {
                initial: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"></path></svg><div><strong>Aguardando um arquivo.</strong><p>Arraste uma planilha para começar.</p></div>`,
                no_results: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg><div><strong>Tudo certo!</strong><p>Nenhuma pendência foi encontrada na sua planilha.</p></div>`,
                no_rules: `<div style="padding:1rem; text-align:center;"><strong>Nenhuma regra configurada.</strong><p>Cores não mapeadas aparecerão na tela principal para configuração.</p></div>`
            };
            return `<div class="placeholder">${states[stateKey] || states.initial}</div>`;
        },
        resetUI() {
            Object.assign(state, { workbook: null, fileName: null, rawResultsText: '', currentSheetData: [], undoStack: [] });
            document.title = 'Planilista - Crie listas a partir de planilhas';
            DOM.fileInput.value = '';
            ['results-area', 'sheet-selector-area'].forEach(id => document.getElementById(id).classList.remove('visible'));
            DOM.generatorOutput.innerHTML = this.getEmptyStateHTML('initial');
            DOM.copyBtn.disabled = true;
            DOM.copyCsvBtn.disabled = true;
            DOM.downloadBtn.disabled = true;
            DOM.downloadSheetBtn.disabled = true;
            DOM.undoBtn.disabled = true;
            DOM.diagnosticoContainer.innerHTML = '';
            this.toggleSettingsAlert(false);
        },
        toggleModal(modal, open) {
            if (!modal) return;
            manageFocus(modal, open);
            if (open) {
                state.activeModalTrigger = document.activeElement;
                modal.classList.add('open');
                const firstFocusable = modal.querySelector('input, button, select');
                if (firstFocusable) firstFocusable.focus();
            } else {
                modal.classList.remove('open');
                if (state.activeModalTrigger) state.activeModalTrigger.focus();
                if (modal.id === 'settings-modal') { DOM.addRuleFormContainer.innerHTML = ''; state.colorToAdd = null; }
            }
            document.body.classList.toggle('modal-open', open);
        },
        showToast(message, type = 'info', options = {}) {
            const toast = document.createElement('div');
            toast.className = `toast ${type}`;
            toast.textContent = message;
            toast.addEventListener('click', () => toast.remove());
            if (options.actionText) {
                const btn = document.createElement('button');
                btn.className = 'toast-action';
                btn.textContent = options.actionText;
                btn.onclick = (e) => { e.stopPropagation(); options.actionCallback(); toast.remove(); };
                toast.appendChild(btn);
            }
            DOM.toastContainer.appendChild(toast);
            setTimeout(() => toast.remove(), 4000);
        },
        toggleTheme(theme) {
            const newTheme = theme || (DOM.doc.classList.contains('light-theme') ? 'dark' : 'light');
            DOM.doc.classList.remove('light-theme', 'dark-theme');
            DOM.doc.classList.add(`${newTheme}-theme`);
            localStorage.setItem('planilistaTheme', newTheme);
        },
        filterRules: debounce((query) => {
            const normalizedQuery = query.toLowerCase().trim();
            DOM.colorRulesList.querySelectorAll('.rule-item').forEach(item => {
                const desc = item.querySelector('input[type="text"]').value.toLowerCase();
                const color = item.dataset.color.toLowerCase();
                item.classList.toggle('hidden', !(desc.includes(normalizedQuery) || color.includes(normalizedQuery)));
            });
        }, 200),
        showSuccessOnButton(button) {
            button.classList.add('is-success');
            button.disabled = true;
            setTimeout(() => { button.classList.remove('is-success'); button.disabled = false; }, 2000);
        },
        prepareAddRuleForm(color) {
            state.colorToAdd = color;
            const formTemplate = DOM.templates.addRuleForm.content.cloneNode(true);
            DOM.addRuleFormContainer.innerHTML = '';
            DOM.addRuleFormContainer.appendChild(formTemplate);
            document.getElementById('rule-color-display-swatch').style.backgroundColor = color;
            document.getElementById('rule-color-display-hex').textContent = color;
            const descInput = document.getElementById('new-desc-input');
            if (descInput) descInput.focus();
        },
        toggleSettingsAlert(show) {
            DOM.settingsBtn.classList.toggle('has-alert', show);
        },
        criarTemplatePar(headers = []) {
            const div = document.createElement('div');
            div.className = 'par-item';
            const optionsHTML = headers.map((h, i) => `<option value="${i}">${esc(h) || `Coluna ${data.toColumnName(i + 1)}`}</option>`).join('');
            div.innerHTML = `<select class="ddd-col-select">${optionsHTML}</select><select class="tel-col-select">${optionsHTML}</select><button class="btn btn-icon btn-delete" data-delete-par aria-label="Excluir par" title="Excluir par"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#dc2626" stroke-width="2"><path d="M18 6L6 18M6 6l12 12"/></svg></button>`;
            return div;
        }
    };

    const data = {
        loadSettings() {
            try {
                const saved = localStorage.getItem('planilistaSettings');
                if (saved) {
                    const parsed = JSON.parse(saved);
                    Object.assign(state.settings, { dedupeResults: true, sortResults: false }, parsed);
                }
            } catch { state.settings = { targetColumn: 'F', colorMap: {}, unirPares: [], aplicarMascara: true, dedupeResults: true, sortResults: false }; }
            DOM.targetColumnInput.value = state.settings.targetColumn;
            DOM.aplicarMascaraCheck.checked = state.settings.aplicarMascara;
            DOM.dedupeResultsCheck.checked = state.settings.dedupeResults;
            DOM.sortResultsCheck.checked = state.settings.sortResults;
            ui.renderColorRules(state.settings.colorMap);
        },
        saveSettings(newSettings) {
            Object.assign(state.settings, newSettings);
            localStorage.setItem('planilistaSettings', JSON.stringify(state.settings));
        },
        processSheet(sheet, {colorMap, targetColumn}) {
            const colIndex = this.columnLetterToIndex(targetColumn);
            if (colIndex === -1) { ui.showToast(`Coluna "${targetColumn}" inválida.`, 'error'); return null; }
            const json = state.currentSheetData;
            const results = {}, skippedRows = [], unmappedColors = new Set();
            for (let i = 1; i < json.length; i++) {
                const lancamento = json[i][colIndex];
                if (lancamento === undefined || lancamento === null || `${lancamento}`.trim() === '') continue;
                if (isNaN(Number(String(lancamento).replace(/\D/g,'')))) { skippedRows.push(`Linha ${i + 1}`); continue; }
                const cellAddress = XLSX.utils.encode_cell({ c: colIndex, r: i });
                const cell = sheet[cellAddress];
                let corHex = null;
                const fill = cell?.s?.fill;
                const colorObj = fill?.fgColor || fill?.bgColor;
                if (colorObj?.rgb) corHex = '#' + colorObj.rgb.slice(-6).toUpperCase();

                const desc = corHex ? colorMap[corHex] : undefined;
                if (desc !== undefined) {
                    if (!results[desc]) results[desc] = [];
                    results[desc].push(lancamento);
                } else if (corHex) {
                    unmappedColors.add(corHex);
                }
            }
            return { ...this.formatResults(state.fileName, results), skippedRows, unmappedColors };
        },
        formatResults(fileName, resultsData) {
            const baseName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
            let raw = `*${esc(baseName.toUpperCase())}*\n\n`;
            let html = '';
            let hasContent = false;
            const categories = Object.keys(resultsData).filter(cat => cat.toUpperCase() !== 'OK');
            if (categories.length > 0) {
                hasContent = true;
                categories.forEach(category => {
                    let nums = resultsData[category].map(String);
                    if (state.settings.dedupeResults) nums = [...new Set(nums)];
                    if (state.settings.sortResults) nums.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

                    raw += `${esc(category)}:\n${nums.map(esc).join('\n')}\n\n`;
                    html += `<div class="result-category"><h4>${esc(category)} (${nums.length})</h4><ul class="result-list">${nums.map(n => `<li>${esc(n)}</li>`).join('')}</ul></div>`;
                });
            }
            return { html, raw: raw.trim(), hasContent };
        },
        formatarTelefoneBR(ddd, tel) {
            const d = String(ddd).replace(/\D/g, ''); let n = String(tel).replace(/\D/g, '');
            if (n.length === 9) return `(${d}) ${n.slice(0, 5)}-${n.slice(5)}`;
            if (n.length === 8) return `(${d}) ${n.slice(0, 4)}-${n.slice(4)}`;
            return `(${d}) ${n}`;
        },
        toColumnName(num) { let s = '', t; while (num > 0) { t = (num - 1) % 26; s = String.fromCharCode(65 + t) + s; num = (num - t - 1) / 26 | 0; } return s || undefined; },
        columnLetterToIndex: (letter) => letter.toUpperCase().split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0) - 1,
    };

    const events = {
        async handleFileSelect(file) {
            if (!file) return;
            ui.resetUI();
            ui.setLoading(true);
            try {
                const buffer = await file.arrayBuffer();
                state.workbook = XLSX.read(buffer, { type: 'array', cellStyles: true });
                state.fileName = file.name;
                const sheetNames = state.workbook.SheetNames;
                const fragment = document.createDocumentFragment();
                sheetNames.forEach(name => {
                    const option = document.createElement('option');
                    option.value = name; option.textContent = name;
                    fragment.appendChild(option);
                });
                DOM.sheetSelector.innerHTML = ''; DOM.sheetSelector.appendChild(fragment);
                DOM.sheetSelectorArea.classList.toggle('visible', sheetNames.length > 1);
                this.reprocessCurrentSheet();
            } catch (error) { console.error("Erro ao processar arquivo:", error); ui.showToast(`Ocorreu um erro: ${error.message}`, 'error');
            } finally { ui.setLoading(false); }
        },
        reprocessCurrentSheet: debounce(() => {
            if (!state.workbook) return;
            const sheetName = DOM.sheetSelector.value;
            const sheet = state.workbook.Sheets[sheetName];
            state.currentSheetData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
            events.populateTargetColumnSelect();
            events.refreshDisplay();
        }, 300),
        refreshDisplay() {
            if (!state.workbook) return;
            const sheetName = DOM.sheetSelector.value;
            const sheet = state.workbook.Sheets[sheetName];
            const processedData = data.processSheet(sheet, state.settings);
            if (processedData) ui.renderResults({ fileName: state.fileName, ...processedData });
        },
        populateTargetColumnSelect() {
            const headers = (state.currentSheetData && state.currentSheetData[0]) ? state.currentSheetData[0] : [];
            const frag = document.createDocumentFragment();
            DOM.targetColumnSelect.innerHTML = '';
            headers.forEach((h, i) => {
                const opt = document.createElement('option');
                const letter = data.toColumnName(i + 1);
                opt.value = i;
                opt.textContent = `${letter} — ${esc(h) || `Coluna ${letter}`}`;
                frag.appendChild(opt);
            });
            DOM.targetColumnSelect.appendChild(frag);
            const colIdx = data.columnLetterToIndex(state.settings.targetColumn);
            if (colIdx >= 0) DOM.targetColumnSelect.value = String(colIdx);
        },
        handleSettingsChange() {
            const newSettings = { ...state.settings, targetColumn: DOM.targetColumnInput.value.trim().toUpperCase() || 'F' };
            data.saveSettings(newSettings);
            this.refreshDisplay();
        },
        handleAddRule(desc) {
            if (!state.colorToAdd || !desc) return;
            if (state.settings.colorMap[state.colorToAdd]) { ui.showToast('Esta cor já está configurada.', 'error'); return; }
            const newColorMap = { ...state.settings.colorMap, [state.colorToAdd]: desc };
            data.saveSettings({ ...state.settings, colorMap: newColorMap });
            ui.renderColorRules(newColorMap); this.refreshDisplay();
            DOM.addRuleFormContainer.innerHTML = ''; state.colorToAdd = null;
            ui.showToast('Nova regra adicionada.', 'success');
        },
        handleRuleListEvents(e) {
            const ruleItem = e.target.closest('.rule-item');
            if (!ruleItem) return;
            const color = ruleItem.dataset.color;
            const newColorMap = { ...state.settings.colorMap };
            if (e.target.closest('.btn-delete')) { delete newColorMap[color];
            } else if (e.target.matches('input[type="text"]')) {
                e.target.onchange = () => {
                    const newDesc = e.target.value.trim();
                    if (newDesc) {
                        newColorMap[color] = newDesc;
                        data.saveSettings({ ...state.settings, colorMap: newColorMap });
                        this.refreshDisplay();
                        ui.showToast(`Regra para ${color} atualizada.`, 'info');
                    } else { e.target.value = state.settings.colorMap[color]; }
                }; return;
            }
            data.saveSettings({ ...state.settings, colorMap: newColorMap });
            ui.renderColorRules(newColorMap); this.refreshDisplay();
        },
        handleResetSettings() {
            state.lastSettingsBeforeReset = { ...state.settings };
            const newSettings = { targetColumn: 'F', colorMap: {}, unirPares: [], aplicarMascara: true, dedupeResults: true, sortResults: false };
            data.saveSettings(newSettings);
            data.loadSettings();
            this.refreshDisplay();
            ui.showToast('Configurações resetadas.', 'info', { actionText: 'Desfazer', actionCallback: () => {
                data.saveSettings(state.lastSettingsBeforeReset); data.loadSettings(); this.refreshDisplay(); ui.showToast('Ação desfeita.', 'success'); }
            });
        },
        async handleCopyClick(e) {
            if (!navigator.clipboard) { this.fallbackCopyText(e.currentTarget); return; }
            try {
                await navigator.clipboard.writeText(state.rawResultsText);
                ui.showToast('Resultados copiados!', 'success'); ui.showSuccessOnButton(e.currentTarget);
            } catch (err) { console.error('Falha ao usar a API Clipboard:', err); this.fallbackCopyText(e.currentTarget); }
        },
        fallbackCopyText(button) {
            const textArea = document.createElement('textarea');
            textArea.value = state.rawResultsText; textArea.style.position = 'fixed'; textArea.style.opacity = '0';
            document.body.appendChild(textArea); textArea.focus(); textArea.select();
            try {
                if (document.execCommand('copy')) { ui.showToast('Resultados copiados!', 'success'); ui.showSuccessOnButton(button);
                } else { ui.showToast('Falha ao copiar.', 'error'); }
            } catch (err) { console.error('Falha no método alternativo:', err); ui.showToast('Falha ao copiar.', 'error'); }
            document.body.removeChild(textArea);
        },
        handleDownloadClick() {
            if (!state.rawResultsText) return;
            const blob = new Blob([state.rawResultsText], { type: 'text/plain;charset=utf-8' });
            const url = URL.createObjectURL(blob); const a = document.createElement('a');
            a.href = url; a.download = `${(state.fileName || 'resultados').replace(/\.[^/.]+$/, "")}.txt`;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            URL.revokeObjectURL(url); ui.showToast('Download iniciado.', 'info');
        },
        handleDownloadSheet() {
            if (state.currentSheetData.length === 0) { ui.showToast('Nenhuma planilha carregada para baixar.', 'error'); return; }
            try {
                const newSheet = XLSX.utils.aoa_to_sheet(state.currentSheetData);
                const headerStyle = { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: "FFFFFFFF" } }, alignment: { horizontal: 'center', vertical: 'center' }, fill: { fgColor: { rgb: "FF4F4F4F" } } };
                const bodyStyle = { font: { name: 'Calibri', sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' } };
                const colWidths = []; const range = XLSX.utils.decode_range(newSheet['!ref']);
                for (let R = range.s.r; R <= range.e.r; ++R) {
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                        const cell = newSheet[cellAddress]; if (!cell) continue;
                        cell.s = (R === 0) ? headerStyle : bodyStyle; cell.t = 's'; 
                        const cellValue = String(cell.v || ''); colWidths[C] = Math.max(colWidths[C] || 10, cellValue.length + 2);
                    }
                }
                newSheet['!cols'] = colWidths.map(w => ({ wch: w }));
                newSheet['!autofilter'] = { ref: newSheet['!ref'] }; newSheet['!freeze'] = { ySplit: 1, state: 'frozen' };
                const newWorkbook = XLSX.utils.book_new();
                const sheetName = DOM.sheetSelector.value || 'Dados Modificados';
                XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
                const baseName = (state.fileName || 'planilha').replace(/\.[^/.]+$/, "");
                XLSX.writeFile(newWorkbook, `${baseName}_modificado.xlsx`);
                ui.showToast('Download da planilha formatada iniciado.', 'success');
            } catch (error) { console.error("Erro ao gerar a planilha:", error); ui.showToast(`Ocorreu um erro ao gerar a planilha: ${error.message}`, 'error'); }
        },
        handleUnirTelefones() {
            if (state.currentSheetData.length === 0) { ui.showToast('Nenhuma planilha carregada.', 'error'); return; }
            const pares = []; const parItems = DOM.paresContainer.querySelectorAll('.par-item');
            for (const item of parItems) {
                const dddIndex = parseInt(item.querySelector('.ddd-col-select').value, 10);
                const telIndex = parseInt(item.querySelector('.tel-col-select').value, 10);
                if (!isNaN(dddIndex) && !isNaN(telIndex)) { pares.push({ dddIndex, telIndex }); }
            }
            if (pares.length === 0) { ui.showToast('Nenhum par de colunas válido.', 'info'); return; }
            state.undoStack.push(JSON.parse(JSON.stringify(state.currentSheetData)));
            if(state.undoStack.length > 5) state.undoStack.shift();
            const aplicarMascara = DOM.aplicarMascaraCheck.checked; let modificacoes = 0;
            for (let i = 1; i < state.currentSheetData.length; i++) {
                const row = state.currentSheetData[i];
                pares.forEach(({ dddIndex, telIndex }) => {
                    if (row[dddIndex] !== undefined && row[telIndex] !== undefined) {
                        const ddd = String(row[dddIndex] || '').trim(); const tel = String(row[telIndex] || '').trim();
                        const telNumeros = tel.replace(/\D/g, '');
                        if (ddd && tel && !telNumeros.startsWith(ddd)) {
                            row[telIndex] = aplicarMascara ? data.formatarTelefoneBR(ddd, tel) : ddd + tel; modificacoes++;
                        }
                    }
                });
            }
            this.refreshDisplay(); data.saveSettings({ unirPares: pares, aplicarMascara });
            ui.showToast(`${modificacoes} células de telefone foram atualizadas.`, 'success');
            ui.toggleModal(DOM.unirTelefonesModal, false);
        },
        handleUndo() { if (state.undoStack.length === 0) return; state.currentSheetData = state.undoStack.pop(); this.refreshDisplay(); ui.showToast('Última união desfeita.', 'info'); },
        exportSettings() {
            const json = JSON.stringify(state.settings, null, 2);
            const blob = new Blob([json], { type: 'application/json' });
            const url = URL.createObjectURL(blob); const a = document.createElement('a');
            a.href = url; a.download = 'planilista_settings.json';
            document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
        },
        importSettings(file) {
            const reader = new FileReader();
            reader.onload = () => {
                try {
                    const parsed = JSON.parse(reader.result);
                    data.saveSettings(parsed); data.loadSettings();
                    this.refreshDisplay(); ui.showToast('Configurações importadas.', 'success');
                } catch (e) { ui.showToast('Arquivo inválido de configurações.', 'error'); }
            };
            reader.readAsText(file);
        },
        copyAsCSV() {
            const parser = new DOMParser();
            const doc = parser.parseFromString(DOM.generatorOutput.innerHTML, 'text/html');
            const rows = [];
            doc.querySelectorAll('.result-category').forEach(cat => {
                const title = cat.querySelector('h4')?.textContent?.replace(/\sKATEX_INLINE_OPEN\d+KATEX_INLINE_CLOSE$/, '').trim() || '';
                cat.querySelectorAll('li').forEach(li => rows.push([title, li.textContent]));
            });
            const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
            navigator.clipboard.writeText(csv).then(() => ui.showToast('CSV copiado!', 'success'));
        },
        handleKeyboardShortcuts(e) {
            if (e.key === 'Escape') { const openModal = document.querySelector('.modal.open'); if (openModal) ui.toggleModal(openModal, false); }
            if ((e.ctrlKey || e.metaKey) && e.key === ',') { e.preventDefault(); ui.toggleModal(DOM.settingsModal, true); }
            if ((e.ctrlKey || e.metaKey) && e.key === 'z') { e.preventDefault(); this.handleUndo(); }
            if (e.key.toLowerCase() === 't' && document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'TEXTAREA') { e.preventDefault(); ui.toggleTheme(); }
        },
        init() {
            data.loadSettings(); ui.resetUI();
            DOM.uploadArea.addEventListener('click', () => DOM.fileInput.click());
            DOM.fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files[0]));
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eName => DOM.uploadArea.addEventListener(eName, (e) => { e.preventDefault(); e.stopPropagation(); DOM.uploadArea.classList.toggle('drag-over', e.type === 'dragenter' || e.type === 'dragover'); if (e.type === 'drop') this.handleFileSelect(e.dataTransfer.files[0]); }));
            DOM.sheetSelector.addEventListener('change', this.reprocessCurrentSheet.bind(this));
            DOM.clearBtn.addEventListener('click', ui.resetUI.bind(ui));
            DOM.copyBtn.addEventListener('click', (e) => this.handleCopyClick(e));
            DOM.copyCsvBtn.addEventListener('click', this.copyAsCSV.bind(this));
            DOM.downloadBtn.addEventListener('click', this.handleDownloadClick);
            DOM.downloadSheetBtn.addEventListener('click', this.handleDownloadSheet.bind(this));
            DOM.themeToggleBtn.addEventListener('click', () => ui.toggleTheme());
            DOM.undoBtn.addEventListener('click', this.handleUndo.bind(this));
            DOM.targetColumnInput.addEventListener('input', debounce(this.handleSettingsChange.bind(this), 500));
            DOM.targetColumnSelect.addEventListener('change', () => {
                const idx = parseInt(DOM.targetColumnSelect.value, 10);
                DOM.targetColumnInput.value = data.toColumnName(idx + 1);
                this.handleSettingsChange();
            });
            DOM.dedupeResultsCheck.addEventListener('change', () => { data.saveSettings({ dedupeResults: DOM.dedupeResultsCheck.checked }); this.refreshDisplay(); });
            DOM.sortResultsCheck.addEventListener('change', () => { data.saveSettings({ sortResults: DOM.sortResultsCheck.checked }); this.refreshDisplay(); });
            DOM.colorRulesList.addEventListener('click', this.handleRuleListEvents.bind(this));
            DOM.ruleSearchInput.addEventListener('input', (e) => ui.filterRules(e.target.value));
            DOM.resetSettingsBtn.addEventListener('click', this.handleResetSettings.bind(this));
            DOM.exportSettingsBtn.addEventListener('click', this.exportSettings.bind(this));
            DOM.importSettingsBtn.addEventListener('click', () => DOM.importSettingsInput.click());
            DOM.importSettingsInput.addEventListener('change', (e) => this.importSettings(e.target.files[0]));
            DOM.diagnosticoContainer.addEventListener('click', (e) => { if (e.target.dataset.configureColor) { const color = e.target.dataset.configureColor; ui.toggleModal(DOM.settingsModal, true); ui.prepareAddRuleForm(color); } });
            DOM.settingsModal.addEventListener('keydown', (e) => { if (e.target.id === 'new-desc-input' && e.key === 'Enter') { e.preventDefault(); this.handleAddRule(e.target.value.trim()); } });
            DOM.settingsModal.addEventListener('change', (e) => { if (e.target.id === 'new-desc-input') { this.handleAddRule(e.target.value.trim()); } });
            document.addEventListener('click', (e) => { 
                const trigger = e.target.closest('[data-modal-trigger]'); const close = e.target.closest('[data-modal-close]');
                if (trigger) { const modalId = trigger.dataset.modalTrigger; if (modalId === 'unir-telefones-modal') this.setupUnirTelefonesModal(); ui.toggleModal(document.getElementById(modalId), true); }
                if (close) ui.toggleModal(document.getElementById(close.dataset.modalClose), false);
                if (e.target.matches('.modal')) ui.toggleModal(e.target, false);
            });
            document.addEventListener('keydown', this.handleKeyboardShortcuts.bind(this));
            DOM.addParBtn.addEventListener('click', () => { const headers = (state.currentSheetData && state.currentSheetData[0]) ? state.currentSheetData[0] : []; DOM.paresContainer.appendChild(ui.criarTemplatePar(headers)); });
            DOM.paresContainer.addEventListener('click', (e) => { const deleteBtn = e.target.closest('[data-delete-par]'); if (deleteBtn) { deleteBtn.closest('.par-item').remove(); } });
            DOM.confirmarUnirBtn.addEventListener('click', this.handleUnirTelefones.bind(this));
            DOM.resetParesBtn.addEventListener('click', () => {
                const newSettings = { ...state.settings, unirPares: [], aplicarMascara: true };
                data.saveSettings(newSettings); state.settings = newSettings;
                this.setupUnirTelefonesModal(); ui.showToast('Configurações de pares resetadas.', 'info');
            });
            const savedTheme = localStorage.getItem('planilistaTheme');
            if (savedTheme) ui.toggleTheme(savedTheme); else if (window.matchMedia?.('(prefers-color-scheme: dark)').matches) ui.toggleTheme('dark');
        },
        setupUnirTelefonesModal() {
            DOM.paresContainer.innerHTML = '';
            const headers = (state.currentSheetData && state.currentSheetData.length > 0) ? state.currentSheetData[0] : [];
            const pares = state.settings.unirPares || [];
            if (pares.length > 0) {
                pares.forEach(par => {
                    const parElement = ui.criarTemplatePar(headers);
                    parElement.querySelector('.ddd-col-select').value = par.dddIndex;
                    parElement.querySelector('.tel-col-select').value = par.telIndex;
                    DOM.paresContainer.appendChild(parElement);
                });
            } else if (headers.length > 0) {
                DOM.paresContainer.appendChild(ui.criarTemplatePar(headers));
            }
        }
    };
    function debounce(func, delay) { let timeout; return function(...args) { clearTimeout(timeout); timeout = setTimeout(() => func.apply(this, args), delay); }; }
    
    events.init();
});