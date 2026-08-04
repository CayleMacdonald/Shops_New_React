(function () {
    function clearEBFilters() {
        const startInput = document.getElementById('ebStartDate');
        const endInput = document.getElementById('ebEndDate');
        const cscSelect = document.getElementById('ebCSC');
        const ccpSelect = document.getElementById('ebCCP');
        if (startInput) startInput.value = '';
        if (endInput) endInput.value = '';
        if (cscSelect) cscSelect.value = '';
        if (ccpSelect) ccpSelect.value = '';
    }

    function stripEBEmojiLabels() {
        const replacements = new Map([
            ['🔎 Filters', 'Filters'],
            ['📅 Start Date', 'Start Date'],
            ['📅 End Date', 'End Date'],
            ['👤 CSC (Sales Rep)', 'CSC (Sales Rep)'],
            ['📞 CCP (Contact Point)', 'CCP (Contact Point)'],
            ['📊 KPI Summary', 'KPI Summary'],
            ['📈 Charts', 'Charts'],
            ['📈 New vs Reactivated per Month', 'New vs Reactivated per Month'],
            ['📞 Performance by Contact Point (CCP)', 'Performance by Contact Point (CCP)']
        ]);

        document.querySelectorAll('#ebusinessPage summary span, #ebusinessPage label, #ebusinessPage h3').forEach(node => {
            const text = (node.textContent || '').trim();
            if (replacements.has(text)) node.textContent = replacements.get(text);
        });
    }

    function overrideImportHelpers() {
        window.normImportHeader = function (h) {
            return String(h == null ? '' : h)
                .replace(/\u00A0/g, ' ')
                .replace(/[%()]/g, '')
                .replace(/[^a-zA-Z0-9]+/g, ' ')
                .replace(/\s+/g, ' ')
                .trim()
                .toLowerCase();
        };

        window.pickMappedColumn = function (headerMap, aliases) {
            for (const alias of aliases) {
                const key = window.normImportHeader(alias);
                if (Object.prototype.hasOwnProperty.call(headerMap, key)) return headerMap[key];
            }
            return '';
        };

        window.parseImportNumber = function (value) {
            if (value == null || value === '') return 0;
            if (typeof value === 'number' && Number.isFinite(value)) return value;
            const normalized = String(value)
                .replace(/\u00A0/g, ' ')
                .replace(/\s+/g, '')
                .replace(/[R$,]/g, '')
                .replace(/\(([^)]+)\)/, '-$1');
            const num = parseFloat(normalized);
            return Number.isFinite(num) ? num : 0;
        };

        window.parseImportDate = function (value) {
            if (value == null || value === '') return '';
            if (value instanceof Date && !isNaN(value.getTime())) {
                const y = value.getFullYear();
                const m = String(value.getMonth() + 1).padStart(2, '0');
                const d = String(value.getDate()).padStart(2, '0');
                return y + '/' + m + '/' + d;
            }
            if (typeof value === 'number' && window.XLSX && XLSX.SSF && XLSX.SSF.parse_date_code) {
                const dc = XLSX.SSF.parse_date_code(value);
                if (dc && dc.y && dc.m && dc.d) {
                    return dc.y + '/' + String(dc.m).padStart(2, '0') + '/' + String(dc.d).padStart(2, '0');
                }
            }
            const raw = String(value).trim();
            const compact = raw.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ');
            const dmy = compact.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
            if (dmy) {
                const d = Number(dmy[1]);
                const m = Number(dmy[2]);
                let y = Number(dmy[3]);
                if (y < 100) y += 2000;
                if (d >= 1 && d <= 31 && m >= 1 && m <= 12) {
                    return y + '/' + String(m).padStart(2, '0') + '/' + String(d).padStart(2, '0');
                }
            }
            const normalized = compact.replace(/-/g, '/');
            const parsed = new Date(normalized);
            if (!isNaN(parsed.getTime())) {
                const y = parsed.getFullYear();
                const m = String(parsed.getMonth() + 1).padStart(2, '0');
                const d = String(parsed.getDate()).padStart(2, '0');
                return y + '/' + m + '/' + d;
            }
            return normalized;
        };
    }

    function overrideImportHandler() {
        window.handleEBImportFile = async function (file) {
            const status = document.getElementById('ebImportStatus');
            const setStatus = (msg, isError) => {
                if (!status) return;
                status.textContent = msg;
                status.style.color = isError ? '#c0392b' : '#2f6f3e';
            };

            try {
                setStatus('Reading ' + file.name + '...', false);
                const buf = await file.arrayBuffer();
                const wb = XLSX.read(buf, { type: 'array', cellDates: true });
                const firstSheet = wb.SheetNames[0];
                if (!firstSheet) throw new Error('No sheets found in workbook.');

                const ws = wb.Sheets[firstSheet];
                const matrixRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
                if (!matrixRows.length) throw new Error('No rows found in import sheet.');

                let headerRowIndex = -1;
                let headerMap = {};
                for (let i = 0; i < Math.min(20, matrixRows.length); i++) {
                    const row = Array.isArray(matrixRows[i]) ? matrixRows[i] : [];
                    const candidateMap = {};
                    row.forEach((cell, index) => {
                        const normalized = window.normImportHeader(cell);
                        if (normalized) candidateMap[normalized] = index;
                    });
                    const hasDay = candidateMap['calendar day'] != null || candidateMap['date'] != null || candidateMap['transaction date'] != null || candidateMap['order date'] != null;
                    const hasCustomer = candidateMap['customer'] != null || candidateMap['customer number'] != null || candidateMap['customer no'] != null || candidateMap['customer code'] != null || candidateMap['business code'] != null;
                    const hasName = candidateMap['cust name'] != null || candidateMap['customer name'] != null;
                    if (hasDay && hasCustomer && hasName) {
                        headerRowIndex = i;
                        headerMap = candidateMap;
                        break;
                    }
                }

                if (headerRowIndex === -1) {
                    throw new Error('Could not find a header row with Date, Customer, and Customer Name columns.');
                }

                const rows = matrixRows.slice(headerRowIndex + 1).map(row => {
                    const data = {};
                    Object.keys(headerMap).forEach(key => {
                        data[key] = Array.isArray(row) ? row[headerMap[key]] : '';
                    });
                    return data;
                }).filter(row => Object.values(row).some(value => String(value == null ? '' : value).trim() !== ''));

                if (!rows.length) throw new Error('No data rows found below the header row.');

                const col = {
                    day: window.pickMappedColumn(headerMap, ['Calendar day', 'Date', 'Sales day', 'Transaction date', 'Order date']),
                    month: window.pickMappedColumn(headerMap, ['Month', 'Period', 'Month-Year']),
                    csc: window.pickMappedColumn(headerMap, ['CSC', 'Sales Rep', 'Representative']),
                    ccp: window.pickMappedColumn(headerMap, ['CCP', 'Contact point', 'Channel']),
                    customer: window.pickMappedColumn(headerMap, ['Customer', 'Customer number', 'Customer no', 'Customer code', 'Business code']),
                    customerName: window.pickMappedColumn(headerMap, ['Cust Name', 'Customer name']),
                    newReact: window.pickMappedColumn(headerMap, ['New / React', 'New/React', 'Status', 'Client type']),
                    turnover: window.pickMappedColumn(headerMap, ['Turnover', 'Total revenue', 'Revenue', 'Turnover in ZAR', 'Sales value']),
                    grossProfit: window.pickMappedColumn(headerMap, ['GrossProfitPFEP', 'Gross profit', 'TotalGrossProfit', 'Gross Profit PFEP']),
                    saves: window.pickMappedColumn(headerMap, ['Saves', 'Save', 'Is save', 'Is saved', 'Saved']),
                    isNew: window.pickMappedColumn(headerMap, ['IsNew', 'Is New', 'New']),
                    isReactivated: window.pickMappedColumn(headerMap, ['IsReactivated', 'Is Reactivated', 'Reactivated'])
                };

                if (col.day === '' || col.customer === '' || col.customerName === '') {
                    throw new Error('Required columns missing. Need Date, Customer/Customer Number, and Customer Name.');
                }

                const importedTransactions = rows.map(r => {
                    const day = window.parseImportDate(r[col.day]);
                    const row = {
                        'Calendar day': day,
                        Month: window.normalizeMonthLabel(col.month !== '' ? r[col.month] : '', day),
                        CSC: col.csc !== '' ? String(r[col.csc] || '').trim() : '',
                        CCP: col.ccp !== '' ? String(r[col.ccp] || 'General').trim() : 'General',
                        Customer: r[col.customer],
                        'Cust Name': col.customerName !== '' ? String(r[col.customerName] || '').trim() : '',
                        'New / React': window.normalizeNewReactValue(col.newReact !== '' ? r[col.newReact] : '', {
                            IsNew: col.isNew !== '' ? r[col.isNew] : '',
                            IsReactivated: col.isReactivated !== '' ? r[col.isReactivated] : ''
                        }),
                        Turnover: window.parseImportNumber(col.turnover !== '' ? r[col.turnover] : 0),
                        GrossProfitPFEP: window.parseImportNumber(col.grossProfit !== '' ? r[col.grossProfit] : 0),
                        Saves: window.normalizeYesNo(col.saves !== '' ? r[col.saves] : '')
                    };
                    row.GrossProfitPFEP = Number(row.GrossProfitPFEP.toFixed(2));
                    row.Turnover = Number(row.Turnover.toFixed(2));
                    return row;
                }).filter(r => r['Calendar day'] && String(r.Customer == null ? '' : r.Customer).trim() !== '' && r['Cust Name']);

                if (!importedTransactions.length) throw new Error('No valid transaction rows found after parsing.');

                const mergedTransactions = window.mergeTransactions(window.ebRawTransactions, importedTransactions);
                window.ebRawTransactions = mergedTransactions;
                window.ebCustomerData = window.buildCustomersFromTransactions(mergedTransactions);
                window.ebVisibleCustomerCount = 50;
                window.initEBusinessFilters();
                clearEBFilters();
                window.showPage('ebusiness');
                window.updateEBusinessDashboard();
                window.persistEBusinessData();

                setStatus('Imported ' + importedTransactions.length.toLocaleString() + ' rows from ' + file.name + '. Total dataset now ' + window.ebRawTransactions.length.toLocaleString() + ' transactions and ' + window.ebCustomerData.length.toLocaleString() + ' customers.', false);
            } catch (err) {
                setStatus('Import failed: ' + (err.message || 'Unknown error'), true);
            }
        };
    }

    document.addEventListener('DOMContentLoaded', function () {
        stripEBEmojiLabels();
        overrideImportHelpers();
        overrideImportHandler();
    });
})();
