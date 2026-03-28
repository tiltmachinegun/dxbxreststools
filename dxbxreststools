// ==UserScript==
// @name         DxBx rests toolkit
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Корректировка объёмов в актах списания (ручная, из XML), сбор марок, удаление бутылок по XML, сбор остатков ЕГАИС
// @author       t.me/tiltmachinegun
// @match        https://dxbx.ru/fe/*
// @match        https://dxbx.ru/index*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_log
// @connect      dxbx.ru
// ==/UserScript==
(function () {
    'use strict';

    let legalPersonId = null;
    let initializationInProgress = false;

    const URLS = {
        bottleSearch: 'https://dxbx.ru/app/egaisbottle/search',
        bottleEdit: 'https://dxbx.ru/app/edit/egaisbottle/',
        bottleLink: 'https://dxbx.ru/index#app/edit/egaisbottle/',
        writeoffActData: 'https://dxbx.ru/api/front/egais/writeoffdocuments/acts/'
    };

    const SELECTORS = {
        actsMarkWrapper: '.documents-writeoffstyled__MarkItemWrapper-sc-e4s6pf-2',
        tableRowLevel0: '.ant-table-row-level-0',
        expandedRow: '.ant-table-expanded-row',
        tableBody: '.ant-table-tbody'
    };

    const JSON_HEADERS = {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    };

    const debugLog = (msg, data = null) => {
        const m = data ? `[ACTS-RESTS] ${msg} ${JSON.stringify(data)}` : `[ACTS-RESTS] ${msg}`;
        console.log(m);
        GM_log(m);
    };

    const notify = (message, type = 'info') => {
        const colors = { success: '#52c41a', error: '#f5222d', warning: '#faad14', info: '#1890ff' };
        const el = document.createElement('div');
        el.style.cssText = `position:fixed;top:20px;right:20px;background:${colors[type]};color:#fff;padding:12px 16px;border-radius:4px;z-index:10000;font-size:14px;box-shadow:0 4px 12px rgba(0,0,0,0.15);max-width:500px;word-break:break-word;`;
        el.textContent = message;
        document.body.appendChild(el);
        setTimeout(() => el.remove(), 4000);
    };

    function createEl(tag, className = '', text = '') {
        const el = document.createElement(tag);
        if (className) el.className = className;
        if (text) el.textContent = text;
        return el;
    }

    const getJSON = (url, { headers = {}, timeout = 15000 } = {}) => new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
            method: 'GET', url, headers: { ...headers }, timeout,
            onload: r => r.status === 200 ? resolve(JSON.parse(r.responseText)) : reject(new Error(`HTTP ${r.status}`)),
            onerror: reject,
            ontimeout: () => reject(new Error('Таймаут запроса'))
        });
    });

    const postJSON = (url, payload, { headers = {}, timeout = 15000 } = {}) => new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
            method: 'POST', url, headers: { ...JSON_HEADERS, ...headers }, data: JSON.stringify(payload), timeout,
            onload: r => {
                if (r.status === 200) {
                    const text = r.responseText?.trim();
                    if (!text) return resolve({});
                    try { resolve(JSON.parse(text)); } catch (e) { resolve({}); }
                } else {
                    reject(new Error(`HTTP ${r.status}`));
                }
            },
            onerror: reject, ontimeout: () => reject(new Error('Таймаут запроса'))
        });
    });

    const putJSON = (url, payload, { headers = {}, timeout = 15000 } = {}) => new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
            method: 'PUT', url, headers: { ...JSON_HEADERS, ...headers }, data: JSON.stringify(payload), timeout,
            onload: r => r.status >= 200 && r.status < 300 ? resolve(r.responseText?.trim() ? JSON.parse(r.responseText) : {}) : reject(new Error(`HTTP ${r.status}`)),
            onerror: reject, ontimeout: () => reject(new Error('Таймаут запроса'))
        });
    });

    const BOTTLE_COLUMNS = [
        { data: 'legalPerson', name: '', searchable: true, orderable: true, search: { value: '', regex: false } },
        { data: 'egaisActItem', name: '', searchable: true, orderable: true, search: { value: '', regex: false } },
        { data: 'shortMarkCode', name: '', searchable: true, orderable: true, search: { value: '', regex: false } },
        { data: 'restsItem', name: '', searchable: true, orderable: true, search: { value: '', regex: false } },
        { data: 'egaisNomenclatureInfo', name: '', searchable: true, orderable: true, search: { value: '', regex: false } },
        { data: 'egaisVolume', name: '', searchable: true, orderable: true, search: { value: '', regex: false } },
        { data: 'egaisVolumeUpdateDate', name: '', searchable: true, orderable: true, search: { value: '', regex: false } },
        { data: 'active', name: '', searchable: true, orderable: true, search: { value: '', regex: false } },
        { data: 'availableVolume', name: '', searchable: true, orderable: false, search: { value: '', regex: false } }
    ];

    const bottleCriterion = (legalPerson, attr, value, oper = 'EQUALS') => ({
        attr: 'legalPerson', value: legalPerson, oper: 'EQUALS',
        clauses: [{ oper: 'AND', criterion: { attr, value, oper, clauses: [] } }]
    });

    const buildBottleSearchPayload = ({ start = 0, length = 200, criterion, order = [{ column: 0, dir: 'asc' }] }) => ({
        draw: 1, columns: BOTTLE_COLUMNS, order, start, length, search: { value: '', regex: false },
        model: 'egaisbottle', searchFormName: 'egaisbottle.default', simpleCrit: { crits: [criterion] }
    });

    const bottleSearch = ({ criterion, start = 0, length = 200, headers = {}, order }) =>
        postJSON(URLS.bottleSearch, buildBottleSearchPayload({ start, length, criterion, order }), { headers });

    function getBottleDetails(bottleId) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url: URLS.bottleEdit + bottleId,
                onload: (response) => {
                    if (response.status !== 200) return reject(new Error(`HTTP error: ${response.status}`));
                    try {
                        const doc = new DOMParser().parseFromString(response.responseText, 'text/html');
                        const volumeInput = doc.querySelector('input[name="availableVolume"]');
                        const markInfoInput = doc.querySelector('input[name="markInfo"]');
                        const egaisVolumeInput = doc.querySelector('#id_egaisVolume');
                        const egaisVolumeDateInput = doc.querySelector('#id_egaisVolumeUpdateDate');

                        const result = {
                            volume: volumeInput?.value || null,
                            markInfo: markInfoInput?.value || null,
                            egaisVolume: null,
                            egaisVolumeUpdateDate: null
                        };

                        if (egaisVolumeInput?.value?.trim()) {
                            const v = egaisVolumeInput.value.trim();
                            result.egaisVolume = (!isNaN(v) && v !== '') ? v : null;
                        }
                        if (egaisVolumeDateInput?.value?.trim()) {
                            result.egaisVolumeUpdateDate = egaisVolumeDateInput.value.trim() || null;
                        }

                        resolve(result);
                    } catch (error) { reject(new Error(`Ошибка парсинга: ${error.message}`)); }
                },
                onerror: (error) => reject(new Error(`Ошибка сети: ${error}`))
            });
        });
    }

    const isActsPage = () => {
        const href = window.location.href;
        const hash = window.location.hash;
        return href.includes('/fe/egais/documents-writeoff/acts/') || hash.includes('fe/egais/documents-writeoff/acts/');
    };

    const isRestsPage = () => window.location.href.includes('https://dxbx.ru/fe/egais/rests');

    const currentLegalPerson = () => legalPersonId || GM_getValue('legalPersonId');

    function interceptXHR() {
        if (window.XMLHttpRequest.prototype._arToolkitIntercepted) return;
        const origOpen = XMLHttpRequest.prototype.open;
        const origSend = XMLHttpRequest.prototype.send;

        XMLHttpRequest.prototype.open = function (method, url, ...rest) {
            this._arUrl = url;
            return origOpen.call(this, method, url, ...rest);
        };

        XMLHttpRequest.prototype.send = function (body) {
            try {
                if (this._arUrl) {
                    const restsMatch = this._arUrl.match(/\/api\/front\/egais\/rests\/legalpersons\/(\d+)\/strong/);
                    if (restsMatch?.[1]) {
                        legalPersonId = restsMatch[1];
                        GM_setValue('legalPersonId', legalPersonId);
                    }
                }
            } catch (e) { }
            return origSend.call(this, body);
        };

        window.XMLHttpRequest.prototype._arToolkitIntercepted = true;
    }

    let actsProcessedRows = new Set();
    let actsLastActId = null;
    let actsLegalPersonId = null;

    function getActIdFromUrl() {
        const match = (window.location.href + window.location.hash).match(/\/acts\/(\d+)/);
        return match ? match[1] : null;
    }

    function insertActsBottleInfo(expandedRow, cssClass, content) {
        const markContainer = expandedRow.querySelector(SELECTORS.actsMarkWrapper);
        if (!markContainer) return;
        markContainer.parentNode.querySelectorAll('.bottle-info-container, .bottle-info-not-found, .bottle-info-error')
            .forEach(el => el.remove());
        const info = createEl('div', cssClass);
        info.style.cssText = 'margin-top:10px;padding:8px;border-radius:4px;font-size:12px;line-height:1.4;';
        if (cssClass.includes('not-found') || cssClass.includes('error')) {
            info.style.background = '#fff2f0';
            info.style.border = '1px solid #ffccc7';
            info.style.color = '#a8071a';
            info.style.fontWeight = '500';
        } else {
            info.style.background = '#f8f9fa';
            info.style.border = '1px solid #e9ecef';
        }
        if (typeof content === 'string') {
            info.textContent = content;
        } else {
            info.append(...content);
        }
        markContainer.parentNode.insertBefore(info, markContainer.nextSibling);
    }

    function renderActsBottleInfo(expandedRow, bottleId, details) {
        const link = createEl('a');
        link.href = URLS.bottleLink + bottleId;
        link.target = '_blank';
        link.textContent = '📦 Открыть бутылку';
        link.style.cssText = 'margin-right:12px;color:#1890ff;text-decoration:underline;cursor:pointer;font-weight:500;';

        const volumeText = details.egaisVolume ? `${details.egaisVolume} мл` : 'нет объема из ЕГАИС';
        const volumeSpan = createEl('span', `bottle-volume ${details.egaisVolume ? 'has-volume' : 'no-volume'}`);
        volumeSpan.textContent = `Объем: ${volumeText}`;
        volumeSpan.style.cssText = `margin-right:12px;font-weight:500;color:${details.egaisVolume ? '#52c41a' : '#ff4d4f'};`;

        const dateSpan = createEl('span', 'bottle-date');
        dateSpan.textContent = `Обновлено: ${details.egaisVolumeUpdateDate || 'дата неизвестна'}`;
        dateSpan.style.cssText = 'color:#666;font-size:11px;';

        insertActsBottleInfo(expandedRow, 'bottle-info-container', [link, volumeSpan, dateSpan]);
    }

    async function processActsRow(row, rowKey) {
        actsProcessedRows.add(rowKey);
        const expandedRow = row.nextElementSibling;
        if (!expandedRow?.classList?.contains('ant-table-expanded-row')) return;

        const markEl = expandedRow.querySelector(SELECTORS.actsMarkWrapper);
        if (!markEl) return;

        const markInfo = markEl.textContent.trim();
        debugLog('Acts: processing mark', { rowKey, mark: markInfo.substring(0, 50) + '...' });

        try {
            const criterion = bottleCriterion(actsLegalPersonId, 'markInfo', markInfo);
            const result = await bottleSearch({ criterion });
            const bottle = result?.data?.[0];

            if (bottle?.DT_RowId) {
                const details = await getBottleDetails(bottle.DT_RowId);
                renderActsBottleInfo(expandedRow, bottle.DT_RowId, details);
            } else {
                insertActsBottleInfo(expandedRow, 'bottle-info-not-found', '❌ Бутылка не найдена в системе');
            }
        } catch (error) {
            debugLog('Acts: error processing row', { rowKey, error: error.message });
            insertActsBottleInfo(expandedRow, 'bottle-info-error', '❌ Ошибка при получении данных о бутылке');
        }
    }

    function checkActsExpandedRows() {
        document.querySelectorAll(SELECTORS.tableRowLevel0).forEach(row => {
            const rowKey = row.getAttribute('data-row-key');
            if (!rowKey || actsProcessedRows.has(rowKey)) return;
            const next = row.nextElementSibling;
            if (next?.classList?.contains('ant-table-expanded-row') && next.querySelector(SELECTORS.actsMarkWrapper)) {
                processActsRow(row, rowKey);
            }
        });
    }

    function setReactInputValue(input, value) {
        const nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        nativeSetter.call(input, value);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }

    function addActButtons() {
        const tableTitle = document.querySelector('.ant-table-title');
        if (!tableTitle || tableTitle.querySelector('.volume-correction-btn')) return;

        const row = tableTitle.querySelector('.ant-row');
        if (!row) return;

        const container = document.createElement('div');
        container.className = 'act-buttons-container';
        container.style.cssText = 'display:flex;align-items:center;gap:6px;margin-left:auto;padding-left:12px;flex-wrap:wrap;';

        const btnStyle = 'ant-btn ant-btn-default button__DxBxButton-sc-gxuid3-0';

        function addBtn(cls, bgColor, label, handler) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = `${btnStyle} ${cls}`;
            btn.style.cssText = `background:${bgColor};color:#fff;border-color:${bgColor};font-size:12px;padding:4px 10px;`;
            btn.innerHTML = `<span>${label}</span>`;
            btn.addEventListener('click', handler);
            container.appendChild(btn);
        }

        addBtn('volume-correction-btn', '#1890ff', 'Скорректировать объем', correctActVolumes);
        addBtn('volume-correction-xml-btn', '#722ed1', 'Корректировка из XML', correctVolumesFromXml);
        addBtn('collect-marks-btn', '#52c41a', 'Собрать марки для проверки', collectMarksForCheck);
        addBtn('delete-bottles-xml-btn', '#f5222d', 'Удалить бутылки по XML', deleteBottlesFromXml);

        row.style.flexWrap = 'nowrap';
        row.style.alignItems = 'center';
        row.appendChild(container);
    }

    async function correctActVolumes() {
        const btn = document.querySelector('.volume-correction-btn');
        if (!btn || btn.disabled) return;
        btn.disabled = true;
        btn.style.opacity = '0.6';
        btn.querySelector('span').textContent = 'Корректируем...';

        let corrected = 0, skipped = 0, errors = 0;

        try {
            const actId = getActIdFromUrl();
            if (!actId) { notify('Не удалось определить ID акта из URL', 'error'); return; }

            if (!actsLegalPersonId) {
                const actData = await getJSON(URLS.writeoffActData + actId);
                actsLegalPersonId = actData.legalPerson.id;
            }
            if (!actsLegalPersonId) { notify('Не удалось определить юр. лицо', 'error'); return; }

            const rows = document.querySelectorAll(SELECTORS.tableRowLevel0);
            const total = rows.length;
            let processed = 0;

            for (const row of rows) {
                processed++;
                btn.querySelector('span').textContent = `Корректировка... (${processed}/${total})`;

                const itemId = row.getAttribute('data-row-key');
                if (!itemId) { skipped++; continue; }

                const expandedRow = row.nextElementSibling;
                if (!expandedRow?.classList?.contains('ant-table-expanded-row')) { skipped++; continue; }

                const markEl = expandedRow.querySelector(SELECTORS.actsMarkWrapper);
                if (!markEl) { skipped++; continue; }
                const markInfo = markEl.textContent.trim();
                if (!markInfo) { skipped++; continue; }

                const input = row.querySelector('.ant-input-number-input');
                if (!input) { skipped++; continue; }
                const currentVolume = parseInt(input.value, 10);
                if (isNaN(currentVolume) || currentVolume <= 0) { skipped++; continue; }

                try {
                    let egaisVolume = null;
                    const volumeSpan = expandedRow.querySelector('.bottle-volume.has-volume');
                    if (volumeSpan) {
                        const m = volumeSpan.textContent.match(/(\d+)/);
                        if (m) egaisVolume = parseInt(m[1], 10);
                    }

                    if (egaisVolume === null) {
                        const criterion = bottleCriterion(actsLegalPersonId, 'markInfo', markInfo);
                        const result = await bottleSearch({ criterion });
                        const bottle = result?.data?.[0];
                        if (bottle?.DT_RowId) {
                            const details = await getBottleDetails(bottle.DT_RowId);
                            if (details.egaisVolume) egaisVolume = parseInt(details.egaisVolume, 10);
                        }
                    }

                    if (egaisVolume !== null && !isNaN(egaisVolume) && egaisVolume > 0 && egaisVolume < currentVolume) {
                        await putJSON(`https://dxbx.ru/api/front/egais/writeoffdocuments/acts/${actId}/items/${itemId}/volumetric`, { volume: egaisVolume });
                        setReactInputValue(input, egaisVolume);
                        corrected++;
                        debugLog('Acts volume corrected', { itemId, from: currentVolume, to: egaisVolume });
                    } else {
                        skipped++;
                    }
                } catch (e) {
                    debugLog('Acts volume correction error', { itemId, error: e.message });
                    errors++;
                }
            }
        } catch (e) {
            debugLog('Acts volume correction fatal error', e.message);
            notify('Ошибка при корректировке: ' + e.message, 'error');
        } finally {
            btn.disabled = false;
            btn.style.opacity = '1';
            btn.querySelector('span').textContent = 'Скорректировать объем';
            notify(`Корректировка завершена: изменено ${corrected}, пропущено ${skipped}${errors ? `, ошибок ${errors}` : ''}`, corrected > 0 ? 'success' : 'info');
        }
    }

    function parseWriteOffXml(xmlText) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, 'text/xml');
        const results = [];
        const allElements = xmlDoc.getElementsByTagName('*');
        const positions = [];
        for (const el of allElements) {
            if (el.localName === 'Position') positions.push(el);
        }
        for (const pos of positions) {
            let mark = null;
            let volume = null;
            for (const el of pos.getElementsByTagName('*')) {
                if (el.localName === 'amc') mark = el.textContent.trim();
                if (el.localName === 'volume') volume = parseInt(el.textContent.trim(), 10);
            }
            if (mark) results.push({ mark, volume: isNaN(volume) ? null : volume });
        }
        return results;
    }

    function filePickerPromise(accept = '.xml') {
        return new Promise(resolve => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = accept;
            input.addEventListener('change', () => resolve(input.files[0] || null));
            input.click();
        });
    }

    async function getActItems(actId) {
        return getJSON(`https://dxbx.ru/api/front/egais/writeoffdocuments/acts/${actId}/items/volumetric`);
    }

    async function correctVolumesFromXml() {
        const file = await filePickerPromise('.xml');
        if (!file) return;

        const btn = document.querySelector('.volume-correction-xml-btn');
        if (btn) { btn.disabled = true; btn.style.opacity = '0.6'; btn.querySelector('span').textContent = 'Обработка XML...'; }

        try {
            const xmlText = await file.text();
            const xmlData = parseWriteOffXml(xmlText);
            if (xmlData.length === 0) { notify('XML файл не содержит позиций с марками', 'warning'); return; }

            const markVolumeMap = new Map();
            xmlData.forEach(item => { if (item.mark && item.volume) markVolumeMap.set(item.mark, item.volume); });

            const actId = getActIdFromUrl();
            if (!actId) { notify('Не удалось определить ID акта', 'error'); return; }

            if (btn) btn.querySelector('span').textContent = 'Загрузка позиций...';
            const actItems = await getActItems(actId);

            let corrected = 0, skipped = 0, errors = 0;
            const total = actItems.length;

            for (let i = 0; i < actItems.length; i++) {
                const item = actItems[i];
                if (btn) btn.querySelector('span').textContent = `Корректировка из XML... (${i + 1}/${total})`;

                if (!item.mark) { skipped++; continue; }

                const newVolume = markVolumeMap.get(item.mark);
                if (!newVolume) { skipped++; continue; }
                if (newVolume === item.volume) { skipped++; continue; }

                try {
                    await putJSON(`https://dxbx.ru/api/front/egais/writeoffdocuments/acts/${actId}/items/${item.id}/volumetric`, { volume: newVolume });
                    corrected++;
                    debugLog('XML volume corrected', { itemId: item.id, from: item.volume, to: newVolume });
                } catch (e) {
                    errors++;
                    debugLog('XML volume correction error', { itemId: item.id, error: e.message });
                }
            }

            notify(`Корректировка из XML: изменено ${corrected}, пропущено ${skipped}${errors ? `, ошибок ${errors}` : ''}`, corrected > 0 ? 'success' : 'info');
            if (corrected > 0) setTimeout(() => window.location.reload(), 1500);
        } catch (e) {
            notify('Ошибка при обработке XML: ' + e.message, 'error');
        } finally {
            if (btn) { btn.disabled = false; btn.style.opacity = '1'; btn.querySelector('span').textContent = 'Корректировка из XML'; }
        }
    }

    async function collectMarksForCheck() {
        const btnEl = document.querySelector('.collect-marks-btn');
        if (btnEl) { btnEl.disabled = true; btnEl.style.opacity = '0.6'; btnEl.querySelector('span').textContent = 'Загрузка...'; }

        try {
            const actId = getActIdFromUrl();
            if (!actId) { notify('Не удалось определить ID акта', 'error'); return; }

            const actItems = await getActItems(actId);
            const marks = actItems.map(item => item.mark).filter(Boolean);

            if (marks.length === 0) { notify('Не найдено марок в акте.', 'warning'); return; }

            const blob = new Blob([marks.join('\n')], { type: 'text/plain;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `marks_act_${actId}_${new Date().toISOString().slice(0, 10)}.txt`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            notify(`Собрано ${marks.length} марок`, 'success');
        } catch (e) {
            notify('Ошибка при сборе марок: ' + e.message, 'error');
        } finally {
            if (btnEl) { btnEl.disabled = false; btnEl.style.opacity = '1'; btnEl.querySelector('span').textContent = 'Собрать марки для проверки'; }
        }
    }

    function deleteActItem(actId, itemId) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'DELETE',
                url: `https://dxbx.ru/api/front/egais/writeoffdocuments/acts/${actId}/items/${itemId}`,
                headers: JSON_HEADERS,
                timeout: 15000,
                onload: r => {
                    if (r.status >= 200 && r.status < 300) return resolve({ success: true });
                    try {
                        const body = JSON.parse(r.responseText);
                        resolve({ success: false, code: body.code, status: r.status });
                    } catch (_) {
                        reject(new Error(`HTTP ${r.status}`));
                    }
                },
                onerror: reject,
                ontimeout: () => reject(new Error('Таймаут запроса'))
            });
        });
    }

    function formatDateDDMMYYYY(date) {
        const dd = String(date.getDate()).padStart(2, '0');
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        return `${dd}.${mm}.${date.getFullYear()}`;
    }

    async function changeActDate(actId, newDateStr) {
        const actData = await getJSON(URLS.writeoffActData + actId);
        await putJSON(URLS.writeoffActData + actId, {
            typeWriteOff: actData.typeWriteOff,
            comment: actData.comment || '',
            status: actData.status,
            date: newDateStr + ' 00:00'
        });
    }

    async function deleteActItemWithDateRetry(actId, itemId) {
        let result = await deleteActItem(actId, itemId);
        if (result.success) return { success: true };

        if (result.code === 'WRONG_ACT_DATE') {
            const today = formatDateDDMMYYYY(new Date());
            debugLog('WRONG_ACT_DATE → changing date to today', today);
            await changeActDate(actId, today);

            result = await deleteActItem(actId, itemId);
            if (result.success) return { success: true, dateChanged: today };

            if (result.code === 'WRONG_ACT_DATE') {
                const tomorrow = new Date();
                tomorrow.setDate(tomorrow.getDate() + 1);
                const tomorrowStr = formatDateDDMMYYYY(tomorrow);
                debugLog('WRONG_ACT_DATE again → changing date to tomorrow', tomorrowStr);
                await changeActDate(actId, tomorrowStr);

                result = await deleteActItem(actId, itemId);
                if (result.success) return { success: true, dateChanged: tomorrowStr };
            }
        }
        return { success: false, code: result.code };
    }

    async function deleteBottlesFromXml() {
        const file = await filePickerPromise('.xml');
        if (!file) return;

        const btn = document.querySelector('.delete-bottles-xml-btn');
        if (btn) { btn.disabled = true; btn.style.opacity = '0.6'; btn.querySelector('span').textContent = 'Обработка...'; }

        try {
            const xmlText = await file.text();
            const xmlData = parseWriteOffXml(xmlText);
            const xmlMarks = new Set(xmlData.map(item => item.mark).filter(Boolean));

            if (xmlMarks.size === 0) { notify('XML файл не содержит марок', 'warning'); return; }

            const actId = getActIdFromUrl();
            if (!actId) { notify('Не удалось определить ID акта', 'error'); return; }

            if (btn) btn.querySelector('span').textContent = 'Загрузка позиций...';
            const actItems = await getActItems(actId);

            const toDelete = actItems.filter(item => item.mark && !xmlMarks.has(item.mark));

            if (toDelete.length === 0) {
                notify('Все бутылки в акте присутствуют в XML. Удалять нечего.', 'info');
                return;
            }

            if (!confirm(`В акте ${actItems.length} позиций, в XML ${xmlMarks.size} марок.\nБудет удалено ${toDelete.length} позиций, которых нет в XML. Продолжить?`)) return;

            let deleted = 0, errors = 0;
            for (let i = 0; i < toDelete.length; i++) {
                const item = toDelete[i];
                if (btn) btn.querySelector('span').textContent = `Удаление... (${i + 1}/${toDelete.length})`;
                try {
                    const result = await deleteActItemWithDateRetry(actId, item.id);
                    if (result.success) {
                        deleted++;
                        debugLog('Deleted item', { itemId: item.id, mark: item.mark.substring(0, 50) });
                    } else {
                        errors++;
                        debugLog('Failed to delete', { itemId: item.id, code: result.code });
                    }
                } catch (e) {
                    errors++;
                    debugLog('Delete error', { itemId: item.id, error: e.message });
                }
            }

            notify(`Удаление завершено: удалено ${deleted}${errors ? `, ошибок ${errors}` : ''}`, deleted > 0 ? 'success' : 'info');
            if (deleted > 0) setTimeout(() => window.location.reload(), 1500);
        } catch (e) {
            notify('Ошибка: ' + e.message, 'error');
        } finally {
            if (btn) { btn.disabled = false; btn.style.opacity = '1'; btn.querySelector('span').textContent = 'Удалить бутылки по XML'; }
        }
    }

    function waitForElement(selector, maxAttempts = 10, interval = 1000) {
        return new Promise(resolve => {
            let attempts = 0;
            const check = () => {
                if (document.querySelector(selector) || ++attempts >= maxAttempts) resolve();
                else setTimeout(check, interval);
            };
            check();
        });
    }

    async function initActsPage() {
        if (initializationInProgress) return;
        initializationInProgress = true;
        try {
            const actId = getActIdFromUrl();
            if (!actId) { debugLog('Acts: actId not found in URL'); return; }

            if (actsLastActId !== actId) {
                debugLog('Acts: new act', { prev: actsLastActId, current: actId });
                actsProcessedRows.clear();
                const actData = await getJSON(URLS.writeoffActData + actId);
                actsLegalPersonId = actData.legalPerson.id;
                debugLog('Acts: legalPersonId', actsLegalPersonId);
                actsLastActId = actId;
            }

            await waitForElement(SELECTORS.tableRowLevel0);
            addActButtons();
            checkActsExpandedRows();
        } catch (error) {
            debugLog('Acts: init error', error.message);
        } finally {
            initializationInProgress = false;
        }
    }

    let _actsObserverActive = false;

    function initActsModule() {
        debugLog('Acts module: starting');
        if (!_actsObserverActive) {
            _actsObserverActive = true;

            const targetNode = document.querySelector('.ant-layout-content') || document.body;
            const observer = new MutationObserver(mutations => {
                if (!isActsPage()) return;
                const hasNewRows = mutations.some(m =>
                    m.type === 'childList' && [...m.addedNodes].some(node =>
                        node.nodeType === 1 && (
                            node.classList?.contains('ant-table-expanded-row') ||
                            node.querySelector?.(SELECTORS.expandedRow) ||
                            node.querySelector?.(SELECTORS.actsMarkWrapper)
                        )
                    )
                );
                if (hasNewRows && !initializationInProgress) {
                    clearTimeout(window._arCheckTimeout);
                    window._arCheckTimeout = setTimeout(checkActsExpandedRows, 500);
                }
            });
            observer.observe(targetNode, { childList: true, subtree: true });

            setInterval(() => {
                if (isActsPage() && !initializationInProgress) {
                    addActButtons();
                    checkActsExpandedRows();
                }
            }, 3000);
        }

        initializationInProgress = false;
        setTimeout(initActsPage, 2000);
    }

    const RESTS_PAGE_SIZE = 100;
    const RESTS_ITEMS_PAGE_SIZE = 5000;
    const RESTS_BASE_URL = 'https://dxbx.ru/api/front/egais/rests/legalpersons';

    function restsApiGet(url) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET', url,
                headers: JSON_HEADERS,
                responseType: 'json',
                timeout: 30000,
                onload: r => {
                    if (r.status >= 200 && r.status < 300) {
                        const data = typeof r.response === 'string' ? JSON.parse(r.response) : r.response;
                        resolve(data);
                    } else reject(new Error(`HTTP ${r.status}`));
                },
                onerror: () => reject(new Error('Network error')),
                ontimeout: () => reject(new Error('Timeout'))
            });
        });
    }

    function restsApiPost(url, body) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'POST', url,
                headers: JSON_HEADERS,
                data: JSON.stringify(body),
                responseType: 'json',
                timeout: 30000,
                onload: r => {
                    if (r.status >= 200 && r.status < 300) {
                        const data = typeof r.response === 'string' ? JSON.parse(r.response) : r.response;
                        resolve(data);
                    } else reject(new Error(`HTTP ${r.status}`));
                },
                onerror: () => reject(new Error('Network error')),
                ontimeout: () => reject(new Error('Timeout'))
            });
        });
    }

    function createRestsProgressUI() {
        const overlay = document.createElement('div');
        overlay.id = 'rests-collect-overlay';
        Object.assign(overlay.style, {
            position: 'fixed', top: '0', left: '0', width: '100%', height: '100%',
            background: 'rgba(0,0,0,0.45)', zIndex: '999999',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
        });
        const box = document.createElement('div');
        Object.assign(box.style, {
            background: '#fff', borderRadius: '12px', padding: '28px 36px',
            minWidth: '380px', maxWidth: '520px', boxShadow: '0 8px 32px rgba(0,0,0,0.25)',
            fontFamily: 'sans-serif', textAlign: 'center',
        });
        const title = document.createElement('div');
        title.textContent = 'Сбор остатков ЕГАИС';
        Object.assign(title.style, { fontSize: '18px', fontWeight: '700', marginBottom: '16px' });
        const statusText = document.createElement('div');
        Object.assign(statusText.style, { fontSize: '14px', marginBottom: '12px', color: '#333' });
        const barOuter = document.createElement('div');
        Object.assign(barOuter.style, {
            width: '100%', height: '22px', background: '#e0e0e0', borderRadius: '11px', overflow: 'hidden',
        });
        const barInner = document.createElement('div');
        Object.assign(barInner.style, {
            width: '0%', height: '100%', background: 'linear-gradient(90deg,#4caf50,#81c784)',
            borderRadius: '11px', transition: 'width 0.3s',
        });
        const pctText = document.createElement('div');
        Object.assign(pctText.style, { fontSize: '13px', marginTop: '8px', color: '#666' });

        barOuter.appendChild(barInner);
        box.append(title, statusText, barOuter, pctText);
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        return {
            setStatus(msg) { statusText.textContent = msg; },
            setProgress(current, total) {
                const pct = total > 0 ? Math.round((current / total) * 100) : 0;
                barInner.style.width = pct + '%';
                pctText.textContent = `${current} / ${total}  (${pct}%)`;
            },
            destroy() { overlay.remove(); },
        };
    }

    async function collectAllRests() {
        const lpId = legalPersonId || currentLegalPerson();
        if (!lpId) {
            notify('Не удалось определить юр. лицо. Откройте страницу остатков.', 'error');
            return;
        }

        const ui = createRestsProgressUI();
        try {
            ui.setStatus('Загрузка первой страницы...');
            ui.setProgress(0, 1);

            const firstPage = await restsApiPost(`${RESTS_BASE_URL}/${lpId}/strong`, { pagination: { page: 0, size: RESTS_PAGE_SIZE }, sort: [] });
            const total = firstPage.total || 0;
            const totalPages = Math.ceil(total / RESTS_PAGE_SIZE);

            let allRecords = firstPage.records || [];
            ui.setStatus(`Загрузка страниц алкокодов (${totalPages})...`);
            ui.setProgress(1, totalPages);

            for (let page = 1; page < totalPages; page++) {
                const data = await restsApiPost(`${RESTS_BASE_URL}/${lpId}/strong`, { pagination: { page, size: RESTS_PAGE_SIZE }, sort: [] });
                allRecords = allRecords.concat(data.records || []);
                ui.setProgress(page + 1, totalPages);
            }

            const alcoCodes = allRecords.map(r => r.alcoCode);
            debugLog(`Rests: собрано алкокодов: ${alcoCodes.length}`);

            ui.setStatus(`Загрузка марок по ${alcoCodes.length} алкокодам (параллельно)...`);
            ui.setProgress(0, alcoCodes.length);

            let completedCount = 0;
            const BATCH_SIZE = 20;
            const allLines = [];

            for (let i = 0; i < alcoCodes.length; i += BATCH_SIZE) {
                const batch = alcoCodes.slice(i, i + BATCH_SIZE);
                const results = await Promise.allSettled(
                    batch.map(code =>
                        restsApiGet(`${RESTS_BASE_URL}/${lpId}/strong/alcocode/${code}/items?page=0&size=${RESTS_ITEMS_PAGE_SIZE}`)
                    )
                );
                for (const result of results) {
                    if (result.status === 'fulfilled') {
                        const items = result.value?.items || [];
                        for (const item of items) {
                            if (item.markInfo && item.formB) {
                                allLines.push(item.markInfo);
                                allLines.push(item.formB);
                            }
                        }
                    }
                }
                completedCount += batch.length;
                ui.setProgress(Math.min(completedCount, alcoCodes.length), alcoCodes.length);
            }

            if (allLines.length === 0) {
                notify('Не найдено ни одной марки.', 'warning');
            } else {
                const blob = new Blob([allLines.join('\n')], { type: 'text/plain;charset=utf-8' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `rests_${lpId}_${new Date().toISOString().slice(0, 10)}.txt`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                notify(`Готово! Собрано ${allLines.length / 2} марок.`, 'success');
            }
        } catch (err) {
            debugLog('Rests collect error', err.message);
            notify('Ошибка при сборе остатков: ' + err.message, 'error');
        } finally {
            ui.destroy();
        }
    }

    function addRestsCollectButton() {
        if (!isRestsPage()) return;
        if (document.querySelector('.rests-collect-btn')) return;

        const allBtns = document.querySelectorAll('.ant-btn.ant-btn-default');
        let writeOffZeroCol = null;
        for (const b of allBtns) {
            if (b.querySelector('span')?.textContent?.trim() === 'Списать в ноль') {
                writeOffZeroCol = b.closest('.ant-col');
                break;
            }
        }
        if (!writeOffZeroCol) return;

        const col = document.createElement('div');
        col.style.cssText = 'padding-left:8px;padding-right:8px;';
        col.className = 'ant-col';
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'ant-btn ant-btn-default button__DxBxButton-sc-gxuid3-0 rests-collect-btn';
        btn.style.cssText = 'background:#52c41a;color:#fff;border-color:#52c41a;';
        btn.innerHTML = '<span>Собрать остатки</span>';
        btn.addEventListener('click', collectAllRests);
        col.appendChild(btn);

        writeOffZeroCol.parentNode.insertBefore(col, writeOffZeroCol);
    }

    let _restsObserverActive = false;

    function initRestsModule() {
        if (_restsObserverActive) return;
        _restsObserverActive = true;

        const targetNode = document.querySelector('.ant-layout-content') || document.body;
        const observer = new MutationObserver(() => {
            if (isRestsPage()) addRestsCollectButton();
        });
        observer.observe(targetNode, { childList: true, subtree: true });

        setInterval(() => { if (isRestsPage()) addRestsCollectButton(); }, 5000);
        addRestsCollectButton();
    }

    let _spaLastHref = window.location.href;

    function onSPANavigation() {
        const href = window.location.href;
        if (href === _spaLastHref) return;
        _spaLastHref = href;
        debugLog('SPA navigation detected', href);

        if (isActsPage()) initActsModule();
        if (isRestsPage()) initRestsModule();
    }

    (function setupSPAWatcher() {
        const origPush = history.pushState;
        const origReplace = history.replaceState;
        history.pushState = function (...args) {
            origPush.apply(this, args);
            setTimeout(onSPANavigation, 200);
        };
        history.replaceState = function (...args) {
            origReplace.apply(this, args);
            setTimeout(onSPANavigation, 200);
        };
        window.addEventListener('popstate', () => setTimeout(onSPANavigation, 200));
    })();

    let lastActsHash = '';

    const handleActsHashChange = () => {
        if (isActsPage()) {
            lastActsHash = window.location.hash;
            initActsModule();
        }
    };

    function init() {
        debugLog('Script loaded, initializing...');
        interceptXHR();

        if (isActsPage()) {
            lastActsHash = window.location.hash;
            initActsModule();
            window.addEventListener('hashchange', handleActsHashChange);
        }

        if (isRestsPage()) {
            initRestsModule();
        }

        window.addEventListener('hashchange', () => {
            if (isActsPage()) handleActsHashChange();
            if (isRestsPage()) initRestsModule();
        });

        if (!isActsPage() && !isRestsPage()) {
            const navObserver = new MutationObserver(() => {
                if (isActsPage() || isRestsPage()) onSPANavigation();
            });
            navObserver.observe(document.body, { childList: true, subtree: true });
        }
    }

    document.readyState === 'loading' ? document.addEventListener('DOMContentLoaded', init) : init();
})();
