/* Planilista - script.js (Versão Final com Download) */
'use strict';

const state = { workbook: null, fileName: null, settings: { targetColumn: 'F', colorMap: {} }, rawResultsText: '', activeModalTrigger: null, lastSettingsBeforeReset: null, colorToAdd: null };
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
    downloadBtn: document.getElementById('download-btn'),
    settingsModal: document.getElementById('settings-modal'),
    targetColumnInput: document.getElementById('target-column-input'),
    colorRulesList: document.getElementById('color-rules-list'),
    ruleSearchInput: document.getElementById('rule-search-input'),
    resetSettingsBtn: document.getElementById('reset-settings-btn'),
    themeToggleBtn: document.getElementById('theme-toggle-btn'),
    liveFeedback: document.getElementById('live-feedback'),
    toastContainer: document.getElementById('toast-container'),
    addRuleFormContainer: document.getElementById('add-rule-form-container'),
    templates: { rule: document.getElementById('rule-template'), addRuleForm: document.getElementById('add-rule-form-template') },
};

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
        DOM.copyBtn.disabled = !hasContent;
        DOM.downloadBtn.disabled = !hasContent;
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
            initial: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"></path></svg><div><strong>Aguardando um ficheiro.</strong><p>Arraste uma planilha para começar.</p></div>`,
            no_results: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg><div><strong>Tudo certo!</strong><p>Nenhuma pendência foi encontrada no seu ficheiro.</p></div>`,
            no_rules: `<div style="padding:1rem; text-align:center;"><strong>Nenhuma regra configurada.</strong><p>Cores não mapeadas aparecerão na tela principal para configuração.</p></div>`
        };
        return `<div class="placeholder">${states[stateKey] || states.initial}</div>`;
    },
    resetUI() {
        Object.assign(state, { workbook: null, fileName: null, rawResultsText: '' });
        document.title = 'Planilista - Carregue uma planilha';
        DOM.fileInput.value = '';
        ['results-area', 'sheet-selector-area'].forEach(id => document.getElementById(id).classList.remove('visible'));
        DOM.generatorOutput.innerHTML = this.getEmptyStateHTML('initial');
        DOM.copyBtn.disabled = true;
        DOM.downloadBtn.disabled = true;
        DOM.diagnosticoContainer.innerHTML = '';
    },
    toggleModal(modal, open) {
        if (open) {
            state.activeModalTrigger = document.activeElement;
            modal.classList.add('open');
            const firstFocusable = modal.querySelector('input, button');
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
        if (options.actionText) {
            const btn = document.createElement('button');
            btn.className = 'toast-action';
            btn.textContent = options.actionText;
            btn.onclick = () => { options.actionCallback(); toast.remove(); };
            toast.appendChild(btn);
        }
        DOM.toastContainer.appendChild(toast);
        setTimeout(() => toast.remove(), 4000);
    },
    toggleTheme(theme) {
        const newTheme = theme || (DOM.doc.classList.contains('light-theme') ? 'dark' : 'light');
        DOM.doc.className = `${newTheme}-theme`;
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
        setTimeout(() => {
            button.classList.remove('is-success');
            button.disabled = false;
        }, 2000);
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
};

const data = {
    loadSettings() {
        try {
            const saved = localStorage.getItem('planilistaSettings');
            state.settings = saved ? JSON.parse(saved) : { targetColumn: 'F', colorMap: {} };
        } catch {
            state.settings = { targetColumn: 'F', colorMap: {} };
        }
        DOM.targetColumnInput.value = state.settings.targetColumn;
        ui.renderColorRules(state.settings.colorMap);
    },
    saveSettings(newSettings) {
        state.settings = newSettings;
        localStorage.setItem('planilistaSettings', JSON.stringify(state.settings));
    },
    processSheet(sheet, {colorMap, targetColumn}) {
        const colIndex = this.columnLetterToIndex(targetColumn);
        if (colIndex === -1) { ui.showToast(`Coluna "${targetColumn}" inválida.`, 'error'); return null; }
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const results = {}, skippedRows = [], unmappedColors = new Set();
        for (let i = 1; i < json.length; i++) {
            const lancamento = json[i][colIndex];
            if (lancamento === undefined || lancamento === null || `${lancamento}`.trim() === '') continue;
            if (isNaN(Number(lancamento))) { skippedRows.push(`Linha ${i + 1}`); continue; }
            const cell = sheet[XLSX.utils.encode_cell({ c: colIndex, r: i })];
            let corHex = cell?.s?.fgColor?.rgb ? '#' + cell.s.fgColor.rgb.slice(-6).toUpperCase() : null;
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
        let raw = `*${baseName.toUpperCase()}*\n\n`;
        let html = '';
        let hasContent = false;
        const categories = Object.keys(resultsData).filter(cat => cat.toUpperCase() !== 'OK');
        if (categories.length > 0) {
            hasContent = true;
            categories.forEach(category => {
                const nums = resultsData[category];
                raw += `${category}:\n${nums.join('\n')}\n\n`;
                html += `<div class="result-category"><h4>${category}</h4><ul class="result-list">${nums.map(n => `<li>${n}</li>`).join('')}</ul></div>`;
            });
        }
        return { html, raw: raw.trim(), hasContent };
    },
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
                option.value = name;
                option.textContent = name;
                fragment.appendChild(option);
            });
            DOM.sheetSelector.innerHTML = '';
            DOM.sheetSelector.appendChild(fragment);
            DOM.sheetSelectorArea.classList.toggle('visible', sheetNames.length > 1);
            this.reprocessCurrentSheet();
        } catch (error) {
            console.error("Erro ao processar ficheiro:", error);
            ui.showToast(`Ocorreu um erro: ${error.message}`, 'error');
        } finally {
            ui.setLoading(false);
        }
    },
    reprocessCurrentSheet: debounce(() => {
        if (!state.workbook) return;
        const sheet = state.workbook.Sheets[DOM.sheetSelector.value];
        const processedData = data.processSheet(sheet, state.settings);
        if (processedData) ui.renderResults({ fileName: state.fileName, ...processedData });
    }, 300),
    handleSettingsChange() {
        const newSettings = { ...state.settings, targetColumn: DOM.targetColumnInput.value.trim().toUpperCase() || 'F' };
        data.saveSettings(newSettings);
        this.reprocessCurrentSheet();
    },
    handleAddRule(desc) {
        if (!state.colorToAdd || !desc) return;
        if (state.settings.colorMap[state.colorToAdd]) { ui.showToast('Esta cor já está configurada.', 'error'); return; }
        const newColorMap = { ...state.settings.colorMap, [state.colorToAdd]: desc };
        data.saveSettings({ ...state.settings, colorMap: newColorMap });
        ui.renderColorRules(newColorMap);
        this.reprocessCurrentSheet();
        DOM.addRuleFormContainer.innerHTML = '';
        state.colorToAdd = null;
        ui.showToast('Nova regra adicionada.', 'success');
    },
    handleRuleListEvents(e) {
        const ruleItem = e.target.closest('.rule-item');
        if (!ruleItem) return;
        const color = ruleItem.dataset.color;
        const newColorMap = { ...state.settings.colorMap };
        if (e.target.closest('.btn-delete')) {
            delete newColorMap[color];
        } else if (e.target.matches('input[type="text"]')) {
            e.target.onchange = () => {
                const newDesc = e.target.value.trim();
                if (newDesc) {
                    newColorMap[color] = newDesc;
                    data.saveSettings({ ...state.settings, colorMap: newColorMap });
                    this.reprocessCurrentSheet();
                    ui.showToast(`Regra para ${color} atualizada.`, 'info');
                } else { e.target.value = state.settings.colorMap[color]; }
            };
            return;
        }
        data.saveSettings({ ...state.settings, colorMap: newColorMap });
        ui.renderColorRules(newColorMap);
        this.reprocessCurrentSheet();
    },
    handleResetSettings() {
        state.lastSettingsBeforeReset = { ...state.settings };
        const newSettings = { targetColumn: 'F', colorMap: {} };
        data.saveSettings(newSettings);
        ui.renderColorRules(newSettings.colorMap);
        this.reprocessCurrentSheet();
        ui.showToast('Configurações resetadas.', 'info', {
            actionText: 'Desfazer',
            actionCallback: () => {
                data.saveSettings(state.lastSettingsBeforeReset);
                ui.renderColorRules(state.lastSettingsBeforeReset.colorMap);
                this.reprocessCurrentSheet();
                ui.showToast('Ação desfeita.', 'success');
            }
        });
    },
    async handleCopyClick(e) {
        if (!navigator.clipboard) {
            this.fallbackCopyText(e.currentTarget);
            return;
        }
        try {
            await navigator.clipboard.writeText(state.rawResultsText);
            ui.showToast('Resultados copiados!', 'success');
            ui.showSuccessOnButton(e.currentTarget);
        } catch (err) {
            console.error('Falha ao usar a API Clipboard (esperado em http/file):', err);
            this.fallbackCopyText(e.currentTarget);
        }
    },
    fallbackCopyText(button) {
        const textArea = document.createElement('textarea');
        textArea.value = state.rawResultsText;
        textArea.style.position = 'fixed';
        textArea.style.opacity = '0';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        try {
            if (document.execCommand('copy')) {
                ui.showToast('Resultados copiados!', 'success');
                ui.showSuccessOnButton(button);
            } else {
                ui.showToast('Falha ao copiar.', 'error');
            }
        } catch (err) {
            console.error('Falha no método alternativo de cópia:', err);
            ui.showToast('Falha ao copiar.', 'error');
        }
        document.body.removeChild(textArea);
    },
    handleDownloadClick() {
        if (!state.rawResultsText) return;
        const blob = new Blob([state.rawResultsText], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${(state.fileName || 'resultados').replace(/\.[^/.]+$/, "")}.txt`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        ui.showToast('Download iniciado.', 'info');
    },
    handleKeyboardShortcuts(e) {
        if (e.key === 'Escape') { const openModal = document.querySelector('.modal.open'); if (openModal) ui.toggleModal(openModal, false); }
        if ((e.ctrlKey || e.metaKey) && e.key === ',') { e.preventDefault(); ui.toggleModal(DOM.settingsModal, true); }
        if (e.key.toLowerCase() === 't' && document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'TEXTAREA') { e.preventDefault(); ui.toggleTheme(); }
    },
    init() {
        data.loadSettings();
        ui.resetUI();
        DOM.uploadArea.addEventListener('click', () => DOM.fileInput.click());
        DOM.fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files[0]));
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eName => DOM.uploadArea.addEventListener(eName, (e) => { e.preventDefault(); e.stopPropagation(); DOM.uploadArea.classList.toggle('drag-over', e.type === 'dragenter' || e.type === 'dragover'); if (e.type === 'drop') this.handleFileSelect(e.dataTransfer.files[0]); }));
        DOM.sheetSelector.addEventListener('change', this.reprocessCurrentSheet.bind(this));
        DOM.clearBtn.addEventListener('click', ui.resetUI.bind(ui));
        DOM.copyBtn.addEventListener('click', (e) => this.handleCopyClick(e));
        DOM.downloadBtn.addEventListener('click', this.handleDownloadClick);
        DOM.themeToggleBtn.addEventListener('click', () => ui.toggleTheme());
        DOM.targetColumnInput.addEventListener('input', debounce(this.handleSettingsChange.bind(this), 500));
        DOM.colorRulesList.addEventListener('click', this.handleRuleListEvents.bind(this));
        DOM.ruleSearchInput.addEventListener('input', (e) => ui.filterRules(e.target.value));
        DOM.resetSettingsBtn.addEventListener('click', this.handleResetSettings.bind(this));
        DOM.diagnosticoContainer.addEventListener('click', (e) => { if (e.target.dataset.configureColor) { const color = e.target.dataset.configureColor; ui.toggleModal(DOM.settingsModal, true); ui.prepareAddRuleForm(color); } });
        DOM.settingsModal.addEventListener('keydown', (e) => { if (e.target.id === 'new-desc-input' && e.key === 'Enter') { e.preventDefault(); this.handleAddRule(e.target.value.trim()); } });
        DOM.settingsModal.addEventListener('change', (e) => { if (e.target.id === 'new-desc-input') { this.handleAddRule(e.target.value.trim()); } });
        document.addEventListener('click', (e) => { if (e.target.closest('[data-modal-trigger]')) ui.toggleModal(document.getElementById(e.target.closest('[data-modal-trigger]').dataset.modalTrigger), true); if (e.target.closest('[data-modal-close]')) ui.toggleModal(document.getElementById(e.target.closest('[data-modal-close]').dataset.modalClose), false); });
        document.addEventListener('keydown', this.handleKeyboardShortcuts);
        const savedTheme = localStorage.getItem('planilistaTheme');
        if (savedTheme) ui.toggleTheme(savedTheme); else if (window.matchMedia?.('(prefers-color-scheme: dark)').matches) ui.toggleTheme('dark');
    }
};

function debounce(func, delay) { let timeout; return function(...args) { clearTimeout(timeout); timeout = setTimeout(() => func.apply(this, args), delay); }; }
events.init();
    </script>
</body>
</html>
