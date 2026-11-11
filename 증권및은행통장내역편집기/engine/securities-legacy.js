// 자동 생성 파일 — 직접 고치지 마세요. 원본: reference/stockbalance/index.html (SHA256 dda8762dc5d3a30d76b147b4a7f8d53fd1cabbf3f08fc182f9e4c95225efc236)
// 생성: node tools/extract-securities-engine.js  (2026-11-10T08:49:56.980Z)
// 포함: DataTransformer 기본 정의 + 재정의 최종본 36개, BrokerAutoMapper, buildAutoClassifyConfig
(function (root) {
    const UIRenderer = {
        showModal(title, message) { (root.LegacySecuritiesNotices = root.LegacySecuritiesNotices || []).push({ title, message }); console.warn('[legacy notice]', title, message); },
        hideModal() {}
    };

    const DataTransformer = {
        validatePhase5Config(config) {
            const requiredFields = [
                { key: 'dateCol', label: '일자' },
                { key: 'amountCol', label: '거래금액' },
                { key: 'balanceCol', label: '잔액' },
                { key: 'tradeTypeCol', label: '거래종류' },
                { key: 'jukyoCol', label: '적요' }
            ];
            const missing = requiredFields.filter(field => config[field.key] === undefined).map(field => field.label);
            if (missing.length > 0) {
                return { isValid: false, message: `${missing.join(', ')} 열을 먼저 선택해주세요.` };
            }
            return { isValid: true };
        },
        _toNumber(value) {
            if (value === null || value === undefined || String(value).trim() === '') return null;
            const num = parseFloat(String(value).replace(/,/g, ''));
            return Number.isNaN(num) ? null : num;
        },
        _getRowValues(row, config) {
            const getValue = (key) => (config[key] === undefined || row[config[key]] === null || row[config[key]] === undefined) ? "" : row[config[key]];
            return {
                tradeType: String(getValue('tradeTypeCol') || '').trim(),
                amount: this._toNumber(getValue('amountCol')) || 0,
                fee: this._toNumber(getValue('feeCol')) || 0,
                tax: this._toNumber(getValue('taxCol')) || 0,
                balance: this._toNumber(getValue('balanceCol'))
            };
        },
        _buildCashflowCandidates(rowInfo) {
            const { amount, fee, tax } = rowInfo;
            return [
                { mode: 'amountOnly', direction: 'deposit', delta: amount },
                { mode: 'amountOnly', direction: 'withdrawal', delta: -amount },
                { mode: 'feeIncluded', direction: 'deposit', delta: amount - fee - tax },
                { mode: 'feeIncluded', direction: 'withdrawal', delta: -(amount + fee + tax) }
            ];
        },
        _inferDirectionFromText(text) {
            const normalized = String(text || '').toLowerCase();
            const depositKeywords = ['매도', '입금', '배당', '환급', '입고', '이입'];
            const withdrawalKeywords = ['매수', '출금', '출고', '수수료', '세금', '제출', '이체'];
            if (depositKeywords.some(keyword => normalized.includes(keyword))) return 'deposit';
            if (withdrawalKeywords.some(keyword => normalized.includes(keyword))) return 'withdrawal';
            return null;
        },
        suggestPhase4Removals(data, config) {
            const requiredFields = [
                { key: 'tradeTypeCol', label: '거래종류' },
                { key: 'amountCol', label: '거래금액' },
                { key: 'balanceCol', label: '예수금잔액' }
            ];
            const missing = requiredFields.filter(field => config[field.key] === undefined).map(field => field.label);
            if (missing.length > 0) {
                return { result: null, error: `${missing.join(', ')} 열을 먼저 선택해주세요.` };
            }
            const body = data.slice(1);
            if (body.length === 0) {
                return { result: { valuesToRemove: [], suggestedRemovals: [] }, error: null };
            }

            const keywordWeights = [
                { keyword: '수수료', weight: 3.5 },
                { keyword: '세금', weight: 3.5 },
                { keyword: '제세', weight: 3.5 },
                { keyword: '출금', weight: 2.5 },
                { keyword: '이체', weight: 2.5 },
                { keyword: '대체', weight: 2.2 },
                { keyword: '정산', weight: 2.2 },
                { keyword: '환전', weight: 2.0 },
                { keyword: '입금', weight: 1.3 },
                { keyword: '매수', weight: 0.8 },
                { keyword: '매도', weight: 0.8 },
                { keyword: '출고', weight: 0.4 },
                { keyword: '입고', weight: 0.4 }
            ];
            const stats = {};
            let prevBalance = config.baseAmount || 0;
            let observedRows = 0;

            body.forEach(row => {
                const tradeType = String(row[config.tradeTypeCol] || '').trim();
                const amount = Math.abs(this._toNumber(row[config.amountCol]) || 0);
                const balance = this._toNumber(row[config.balanceCol]);
                if (!tradeType || balance === null) {
                    if (balance !== null) prevBalance = balance;
                    return;
                }
                const observedDelta = balance - prevBalance;
                prevBalance = balance;
                observedRows += 1;
                stats[tradeType] = stats[tradeType] || { keepError: 0, removeError: 0, rows: 0, keywordWeight: 1 };
                stats[tradeType].rows += 1;
                stats[tradeType].keepError += Math.min(Math.abs(observedDelta - amount), Math.abs(observedDelta + amount));
                stats[tradeType].removeError += Math.abs(observedDelta);
            });

            const suggestedRemovals = Object.entries(stats).filter(([tradeType, tradeStats]) => {
                const normalized = tradeType.toLowerCase();
                const weight = keywordWeights.reduce((maxWeight, entry) => normalized.includes(entry.keyword) ? Math.max(maxWeight, entry.weight) : maxWeight, 1);
                tradeStats.keywordWeight = weight;
                const adjustedRemoveError = tradeStats.removeError / weight;
                if (tradeStats.rows === 0) return false;
                if (tradeStats.keepError === 0 && adjustedRemoveError > 0) return false;
                return adjustedRemoveError + 1 < tradeStats.keepError;
            }).sort((a, b) => {
                const scoreA = (a[1].keepError - (a[1].removeError / a[1].keywordWeight));
                const scoreB = (b[1].keepError - (b[1].removeError / b[1].keywordWeight));
                return scoreB - scoreA;
            }).map(([tradeType]) => tradeType);

            if (suggestedRemovals.length === 0 && observedRows > 0) {
                UIRenderer.showModal('알림', '잔액 흐름을 기준으로 제거 후보를 추론했지만 확신이 높은 항목을 찾지 못했습니다. 수동으로 확인해주세요.', [{ text: '확인', className: 'bg-indigo-600 text-white', onClick: () => UIRenderer.hideModal() }]);
            } else if (suggestedRemovals.length > 0) {
                UIRenderer.showModal('알림', `${suggestedRemovals.length}개 거래종류를 제거 추천으로 반영했습니다. 추천 결과는 드래그해서 자유롭게 수정할 수 있습니다.`, [{ text: '확인', className: 'bg-indigo-600 text-white', onClick: () => UIRenderer.hideModal() }]);
            }

            return {
                result: {
                    valuesToRemove: suggestedRemovals,
                    suggestedRemovals
                },
                error: null
            };
        },
        suggestPhase5Removals(data, config) {
            const requiredFields = [
                { key: 'dateCol', label: '일자' },
                { key: 'tradeTypeCol', label: '거래종류' },
                { key: 'amountCol', label: '거래금액' },
                { key: 'balanceCol', label: '예수금잔액' },
                { key: 'jukyoCol', label: '적요' }
            ];
            const missing = requiredFields.filter(field => config[field.key] === undefined).map(field => field.label);
            if (missing.length > 0) {
                return { result: null, error: `${missing.join(', ')} 열을 먼저 선택해주세요.` };
            }

            const body = data.slice(1);
            if (body.length === 0) {
                return { result: { valuesToRemove: config.valuesToRemove || [], suggestedRemovals: [] }, error: null };
            }

            const grouped = {};
            let prevBalance = config.baseAmount || 0;
            const feeTolerance = 20;

            body.forEach(row => {
                const tradeType = String(row[config.tradeTypeCol] || '').trim();
                const balance = this._toNumber(row[config.balanceCol]);
                const amount = Math.abs(this._toNumber(row[config.amountCol]) || 0);
                const fee = Math.abs(this._toNumber(config.feeCol !== undefined ? row[config.feeCol] : null) || 0);
                const tax = Math.abs(this._toNumber(config.taxCol !== undefined ? row[config.taxCol] : null) || 0);
                if (!tradeType || balance === null) {
                    if (balance !== null) prevBalance = balance;
                    return;
                }
                const observedDelta = balance - prevBalance;
                prevBalance = balance;
                const dateKey = config.dateCol !== undefined ? String(row[config.dateCol] || '') : '';
                const memoKey = config.jukyoCol !== undefined ? String(row[config.jukyoCol] || '') : '';
                const groupKey = `${tradeType}__${dateKey}__${memoKey}`;
                grouped[groupKey] = grouped[groupKey] || { tradeType, observedDelta: 0, amount: 0, fee: 0, tax: 0, count: 0 };
                grouped[groupKey].observedDelta += observedDelta;
                grouped[groupKey].amount += amount;
                grouped[groupKey].fee += fee;
                grouped[groupKey].tax += tax;
                grouped[groupKey].count += 1;
            });

            const typeScores = {};
            Object.values(grouped).forEach(group => {
                const candidates = [
                    group.amount,
                    -group.amount,
                    group.amount - group.fee - group.tax,
                    -(group.amount + group.fee + group.tax),
                    group.amount + group.fee,
                    -(group.amount + group.fee),
                    group.amount - group.tax,
                    -(group.amount + group.tax)
                ];
                const keepError = Math.min(...candidates.map(delta => {
                    const diff = Math.abs(group.observedDelta - delta);
                    return diff <= feeTolerance ? diff * 0.5 : diff;
                }));
                const removeError = Math.abs(group.observedDelta);
                typeScores[group.tradeType] = typeScores[group.tradeType] || { keepError: 0, removeError: 0, count: 0 };
                typeScores[group.tradeType].keepError += keepError;
                typeScores[group.tradeType].removeError += removeError;
                typeScores[group.tradeType].count += 1;
            });

            const recommended = Object.entries(typeScores)
                .filter(([, score]) => score.count > 0 && score.removeError + 1 < score.keepError)
                .sort((a, b) => (b[1].keepError - b[1].removeError) - (a[1].keepError - a[1].removeError))
                .map(([tradeType]) => tradeType);

            const merged = [...new Set([...(config.valuesToRemove || []), ...recommended])];
            if (recommended.length > 0) {
                UIRenderer.showModal('알림', `${recommended.length}개 항목을 잔액 흐름 기준으로 다시 제거 후보에 반영했습니다. 필요하면 직접 수정해주세요.`, [{ text: '확인', className: 'bg-indigo-600 text-white', onClick: () => UIRenderer.hideModal() }]);
            } else {
                UIRenderer.showModal('알림', '최종 매핑 기준으로 제거 후보를 다시 검토했지만 새로 확정할 항목은 찾지 못했습니다. 기존 후보를 확인해주세요.', [{ text: '확인', className: 'bg-indigo-600 text-white', onClick: () => UIRenderer.hideModal() }]);
            }

            return { result: { valuesToRemove: merged, suggestedRemovals: recommended }, error: null };
        },
        _parseDate(value) {
            if (value === null || value === undefined || String(value).trim() === '') return null;
            
            // Handle Excel's numeric date format
            if (typeof value === 'number' && value > 10000) {
                // Adjust for timezone offset by creating date in UTC
                const utc_days  = Math.floor(value - 25569);
                const utc_value = utc_days * 86400;                                        
                const date_info = new Date(utc_value * 1000);
                return new Date(date_info.getUTCFullYear(), date_info.getUTCMonth(), date_info.getUTCDate());
            }
            
            // Convert to string and take only the date part (before any space)
            const dateStr = String(value).trim().split(' ')[0];
            
            // Try parsing YYYYMMDD
            if (/^\d{8}$/.test(dateStr)) {
                const year = dateStr.substring(0, 4);
                const month = dateStr.substring(4, 6);
                const day = dateStr.substring(6, 8);
                const d2 = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                if (!isNaN(d2.getTime())) return d2;
            }

            // Try parsing various other formats (e.g., YYYY-MM-DD, YYYY.MM.DD)
            const d = new Date(dateStr.replace(/\./g, '-').replace(/\//g, '-'));
            if (!isNaN(d.getTime())) return d;
            
            return null;
        },

        _formatDate(dateObj) {
            if (!dateObj || isNaN(dateObj.getTime())) return "";
            const year = dateObj.getFullYear();
            const month = String(dateObj.getMonth() + 1).padStart(2, '0');
            const day = String(dateObj.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        },

        autoSortByDate(data) { 
            if (data.length < 2) return data; 
            const header = data[0]; 
            const body = data.slice(1); 
            const dateKeywords = ['일자', '거래일', 'date']; 
            let dateColIndex = header.findIndex(h => dateKeywords.some(k => String(h).toLowerCase().includes(k))); 
            if (dateColIndex === -1) { 
                console.log("날짜 열을 찾을 수 없어 정렬을 건너뜁니다."); 
                return data; 
            } 
            
            const parseDate = (value) => this._parseDate(value); 

            let firstDate, lastDate; 
            for(let i=0; i<body.length; i++) { if(body[i][dateColIndex]) { firstDate = parseDate(body[i][dateColIndex]); break; } } 
            for(let i=body.length-1; i>=0; i--) { if(body[i][dateColIndex]) { lastDate = parseDate(body[i][dateColIndex]); break; } } 
            
            let sortedBody = body; 
            if (firstDate && lastDate && firstDate > lastDate) { 
                console.log("날짜가 역순으로 확인되어 전체 데이터를 뒤집습니다."); 
                sortedBody = [...body].reverse(); 
            } 
            
            console.log("날짜 자동 정렬이 완료되었습니다."); 
            return [header, ...sortedBody]; 
        },
        transformPhase1(data, headerRows) { const numRowsPerGroup = headerRows.length; const repeatingHeadersData = headerRows.map(rowIndex => data[rowIndex] || []); const newHeaders = repeatingHeadersData.flat().map(cell => cell || ''); const outputData = [newHeaders]; const dataStartIndex = Math.max(...headerRows) + 1; for (let i = dataStartIndex; i < data.length; i += numRowsPerGroup) { const group = data.slice(i, i + numRowsPerGroup); if (group.every(row => !row || row.every(cell => cell === null || String(cell).trim() === ''))) continue; const newRow = group.flat(); const finalRow = newHeaders.map((_, j) => newRow[j] ?? null); outputData.push(finalRow); } return outputData; },
        transformPhase2(data, splitCol, groups) { if (splitCol === null) return [{ name: '단일계좌', data: data }]; const header = data[0]; const body = data.slice(1); const groupedData = groups.map(group => ({ name: group.name, data: [header, ...body.filter(row => group.values.includes(String(row[splitCol])))] })); const allGroupedValues = groups.flatMap(g => g.values); const ungroupedRows = body.filter(row => !allGroupedValues.includes(String(row[splitCol]))); if(ungroupedRows.length > 0) { groupedData.push({ name: '기타', data: [header, ...ungroupedRows] }); } return groupedData.filter(g => g.data.length > 1); },
        transformPhase3(data, config) { const header = data[0]; const body = data.slice(1); const newHeader = [...header, "결과 생성"]; const newBody = body.map(row => { const getValue = (key, isNumeric = false) => { const colIndex = config[key]; if (colIndex === undefined || row[colIndex] === null || row[colIndex] === undefined) return ""; const value = String(row[colIndex]).trim().replace(/\u00A0/g, ""); if (isNumeric && (value === "" || parseFloat(value.replace(/,/g, '')) === 0)) return ""; return value; }; const jukyoParts = [ getValue('tradeTypeCol'), getValue('stockNameCol'), getValue('stockQtyCol', true) ? `${getValue('stockQtyCol', true)}주` : '', getValue('unitPriceCol', true) ? `@${getValue('unitPriceCol', true)}` : '', getValue('stockBalanceCol', true) ? `잔고${getValue('stockBalanceCol', true)}주` : '', getValue('additionalCol1'), getValue('additionalCol2') ]; return [...row, jukyoParts.filter(Boolean).join(' / ')]; }); return [newHeader, ...newBody]; },
        transformPhase4(data, config) { const header = data[0]; const body = data.slice(1); if (!config || config.tradeTypeCol === undefined) return { kept: data, removed: [header] }; const valuesToRemove = config.valuesToRemove || []; const keptRows = body.filter(row => !valuesToRemove.includes(String(row[config.tradeTypeCol]))); const removedRows = body.filter(row => valuesToRemove.includes(String(row[config.tradeTypeCol]))); return { kept: [header, ...keptRows], removed: [header, ...removedRows] }; },
        classifyByBalance(data, config) {
            const validation = this.validatePhase5Config(config);
            if (!validation.isValid) {
                return { result: null, error: validation.message };
            }
            const body = data.slice(1);
            if (body.length === 0) {
                return { result: { depositTypes: [], withdrawalTypes: [], inferredFeeMode: 'separate' }, error: null };
            }

            const rows = body.map(row => this._getRowValues(row, config));
            const tolerance = 1;
            const typeStats = {};
            let prevBalance = config.baseAmount || 0;

            rows.forEach(rowInfo => {
                const tradeType = rowInfo.tradeType;
                if (!tradeType || rowInfo.balance === null) {
                    if (rowInfo.balance !== null) prevBalance = rowInfo.balance;
                    return;
                }

                const observedDelta = rowInfo.balance - prevBalance;
                prevBalance = rowInfo.balance;
                const candidates = this._buildCashflowCandidates(rowInfo).map(candidate => {
                    let score = -Math.abs(observedDelta - candidate.delta);
                    if (Math.abs(observedDelta - candidate.delta) <= tolerance) score += 5;
                    const textDirection = this._inferDirectionFromText(tradeType);
                    if (textDirection && textDirection === candidate.direction) score += 2;
                    if ((rowInfo.fee > 0 || rowInfo.tax > 0) && candidate.mode === 'feeIncluded') score += 1.5;
                    if (rowInfo.amount > 0 && observedDelta !== 0 && Math.sign(observedDelta) === Math.sign(candidate.delta)) score += 1;
                    return { ...candidate, score };
                }).sort((a, b) => b.score - a.score);

                const best = candidates[0];
                typeStats[tradeType] = typeStats[tradeType] || {
                    deposit: 0,
                    withdrawal: 0,
                    amountOnly: 0,
                    feeIncluded: 0,
                    matchedRows: 0,
                    unresolvedRows: 0
                };

                if (Math.abs(observedDelta - best.delta) <= tolerance || best.score > -2) {
                    typeStats[tradeType][best.direction] += 1;
                    typeStats[tradeType][best.mode] += 1;
                    typeStats[tradeType].matchedRows += 1;
                } else {
                    typeStats[tradeType].unresolvedRows += 1;
                }
            });

            const depositTypes = [];
            const withdrawalTypes = [];
            let feeIncludedHits = 0;
            let amountOnlyHits = 0;

            Object.entries(typeStats).forEach(([tradeType, stats]) => {
                if (stats.deposit > stats.withdrawal) depositTypes.push(tradeType);
                else if (stats.withdrawal > stats.deposit) withdrawalTypes.push(tradeType);
                feeIncludedHits += stats.feeIncluded;
                amountOnlyHits += stats.amountOnly;
            });

            const inferredFeeMode = feeIncludedHits > amountOnlyHits ? 'included' : 'separate';
            const classifiedCount = depositTypes.length + withdrawalTypes.length;

            if (classifiedCount === 0) {
                UIRenderer.showModal('알림', '잔액 변동을 기준으로 입출금을 추정했지만 확실한 규칙을 찾지 못했습니다. 불확실한 항목만 직접 확인해주세요.', [{ text: '확인', className: 'bg-indigo-600 text-white', onClick: () => UIRenderer.hideModal() }]);
            } else {
                const feeModeText = inferredFeeMode === 'included' ? '수수료 포함형' : '수수료 별도형';
                UIRenderer.showModal('알림', `${depositTypes.length}개의 입금 종류와 ${withdrawalTypes.length}개의 출금 종류가 자동으로 분류되었습니다. 추정 정산 방식: ${feeModeText}`, [{ text: '확인', className: 'bg-indigo-600 text-white', onClick: () => UIRenderer.hideModal() }]);
            }

            return { result: { depositTypes, withdrawalTypes, inferredFeeMode }, error: null };
        },
        verifyPhase5(data, config) { const body = data.slice(1); if (body.length === 0) return { isCorrect: true }; const finalTransformedData = this.transformPhase5(data, config); if (finalTransformedData.length <= 1) return { isCorrect: true }; const lastTransformedRow = finalTransformedData[finalTransformedData.length - 1]; const finalCalculatedBalance = lastTransformedRow[3]; const lastOriginalRow = body[body.length - 1]; const getOriginalValue = (key) => (config[key] === undefined || lastOriginalRow[config[key]] === null) ? "" : lastOriginalRow[config[key]]; const getOriginalNumeric = (key) => parseFloat(String(getOriginalValue(key)).replace(/,/g, '')) || 0; const finalOriginalBalance = getOriginalNumeric('balanceCol'); if (finalOriginalBalance !== 0 && Math.abs(finalCalculatedBalance - finalOriginalBalance) > 1) { return { isCorrect: false, calculated: finalCalculatedBalance, original: finalOriginalBalance }; } return { isCorrect: true }; },
        transformPhase5(data, config) {
            const body = data.slice(1); 
            const newSheet = [['일자', '입금', '출금', '잔액', '(참고용)잔액', '적요']]; 
            let runningBalance = config.baseAmount || 0; 
            body.forEach(row => { 
                const getValue = (key) => (config[key] === undefined || row[config[key]] === null) ? "" : row[config[key]]; 
                const getNumeric = (key) => parseFloat(String(getValue(key)).replace(/,/g, '')) || 0; 
                
                const tradeType = getValue('tradeTypeCol'); 
                const rawDate = getValue('dateCol');
                const dateObj = this._parseDate(rawDate);
                const date = this._formatDate(dateObj);
                const jukyo = getValue('jukyoCol'); 
                const originalBalance = getValue('balanceCol'); 
                const amount = getNumeric('amountCol'); 
                const fee = getNumeric('feeCol'); 
                const tax = getNumeric('taxCol'); 
                let deposit = 0, withdrawal = 0; 
                if (config.depositTypes?.includes(tradeType)) { deposit = amount; } 
                else if (config.withdrawalTypes?.includes(tradeType)) { withdrawal = amount; } 
                
                if (deposit > 0) { runningBalance += deposit; newSheet.push([date, deposit, 0, runningBalance, originalBalance, jukyo]); } 
                else if (withdrawal > 0) { runningBalance -= withdrawal; newSheet.push([date, 0, withdrawal, runningBalance, originalBalance, jukyo]); } 
                
                if (fee > 0) { runningBalance -= fee; newSheet.push([date, 0, fee, runningBalance, originalBalance, `(수수료)${jukyo}`]); } 
                if (tax > 0) { runningBalance -= tax; newSheet.push([date, 0, tax, runningBalance, originalBalance, `(제세금)${jukyo}`]); } 
            }); 
            return newSheet; 
        },
        _getExtraCostDetails(row, config, headerRow = []) {
            const keys = ['extraCostCol1', 'extraCostCol2', 'extraCostCol3'];
            return keys.map((key, index) => {
                const colIndex = config[key];
                if (colIndex === undefined || row[colIndex] === null || row[colIndex] === undefined || String(row[colIndex]).trim() === '') return null;
                const amount = Math.abs(this._toNumber(row[colIndex]) || 0);
                if (amount <= 0) return null;
                const rawLabel = headerRow[colIndex] === undefined || headerRow[colIndex] === null ? '' : String(headerRow[colIndex]).trim();
                return {
                    key,
                    amount,
                    label: rawLabel || `기타비용${index + 1}`
                };
            }).filter(Boolean);
        },
        _getRowValues(row, config) {
            const getValue = (key) => (config[key] === undefined || row[config[key]] === null || row[config[key]] === undefined) ? "" : row[config[key]];
            const extraCosts = this._getExtraCostDetails(row, config);
            const extraCost = extraCosts.reduce((sum, item) => sum + item.amount, 0);
            return {
                tradeType: String(getValue('tradeTypeCol') || '').trim(),
                date: String(getValue('dateCol') || '').trim(),
                memo: String(getValue('jukyoCol') || '').trim(),
                amount: this._toNumber(getValue('amountCol')) || 0,
                fee: this._toNumber(getValue('feeCol')) || 0,
                tax: this._toNumber(getValue('taxCol')) || 0,
                extraCost,
                extraCosts,
                balance: this._toNumber(getValue('balanceCol'))
            };
        },
        _buildCashflowCandidates(rowInfo) {
            const { amount, fee, tax, extraCost } = rowInfo;
            const totalCost = fee + tax + extraCost;
            return [
                { mode: 'amountOnly', direction: 'deposit', delta: amount },
                { mode: 'amountOnly', direction: 'withdrawal', delta: -amount },
                { mode: 'feeIncluded', direction: 'deposit', delta: amount - totalCost },
                { mode: 'feeIncluded', direction: 'withdrawal', delta: -(amount + totalCost) }
            ];
        },
        suggestPhase4Removals(data, config) {
            return this.suggestPhase5Removals(data, config);
        },
        suggestPhase5Removals(data, config) {
            const requiredFields = [
                { key: 'dateCol', label: '일자' },
                { key: 'tradeTypeCol', label: '거래종류' },
                { key: 'amountCol', label: '거래금액' },
                { key: 'balanceCol', label: '잔액' },
                { key: 'jukyoCol', label: '적요' }
            ];
            const missing = requiredFields.filter(field => config[field.key] === undefined).map(field => field.label);
            if (missing.length > 0) {
                return { result: null, error: `${missing.join(', ')} 열을 먼저 선택해주세요.` };
            }
            const body = data.slice(1);
            if (body.length === 0) {
                return { result: { valuesToRemove: config.valuesToRemove || [], suggestedRemovals: [] }, error: null };
            }

            const grouped = {};
            let prevBalance = config.baseAmount || 0;
            const feeTolerance = 20;

            body.forEach(row => {
                const rowInfo = this._getRowValues(row, config);
                if (!rowInfo.tradeType || rowInfo.balance === null) {
                    if (rowInfo.balance !== null) prevBalance = rowInfo.balance;
                    return;
                }
                const observedDelta = rowInfo.balance - prevBalance;
                prevBalance = rowInfo.balance;
                const groupKey = `${rowInfo.tradeType}__${rowInfo.date}__${rowInfo.memo}`;
                grouped[groupKey] = grouped[groupKey] || { tradeType: rowInfo.tradeType, observedDelta: 0, amount: 0, fee: 0, tax: 0, extraCost: 0, count: 0 };
                grouped[groupKey].observedDelta += observedDelta;
                grouped[groupKey].amount += Math.abs(rowInfo.amount);
                grouped[groupKey].fee += Math.abs(rowInfo.fee);
                grouped[groupKey].tax += Math.abs(rowInfo.tax);
                grouped[groupKey].extraCost += Math.abs(rowInfo.extraCost);
                grouped[groupKey].count += 1;
            });

            const keywordWeights = [
                { keyword: '수수료', weight: 3.4 },
                { keyword: '세금', weight: 3.4 },
                { keyword: '제세금', weight: 3.4 },
                { keyword: '기타수수료', weight: 3.2 },
                { keyword: '기타제세금', weight: 3.2 },
                { keyword: '이체', weight: 2.5 },
                { keyword: '대체', weight: 2.5 },
                { keyword: '정산', weight: 2.3 },
                { keyword: '환전', weight: 2.0 },
                { keyword: '출고', weight: 1.4 }
            ];

            const typeScores = {};
            Object.values(grouped).forEach(group => {
                const totalCost = group.fee + group.tax + group.extraCost;
                const candidates = [
                    group.amount,
                    -group.amount,
                    group.amount - totalCost,
                    -(group.amount + totalCost),
                    group.amount + totalCost,
                    -Math.max(0, group.amount - totalCost)
                ];
                const keepError = Math.min(...candidates.map(delta => {
                    const diff = Math.abs(group.observedDelta - delta);
                    return diff <= feeTolerance ? diff * 0.5 : diff;
                }));
                const removeError = Math.abs(group.observedDelta);
                const lowered = String(group.tradeType || '').toLowerCase();
                const weight = keywordWeights.reduce((max, entry) => lowered.includes(entry.keyword) ? Math.max(max, entry.weight) : max, 1);
                typeScores[group.tradeType] = typeScores[group.tradeType] || { keepError: 0, removeError: 0, count: 0, weight: 1, rawRemoveError: 0 };
                typeScores[group.tradeType].keepError += keepError;
                typeScores[group.tradeType].removeError += removeError / weight;
                typeScores[group.tradeType].rawRemoveError += removeError;
                typeScores[group.tradeType].count += 1;
                typeScores[group.tradeType].weight = Math.max(typeScores[group.tradeType].weight, weight);
            });

            const recommended = Object.entries(typeScores)
                .filter(([, score]) => {
                    if (score.count <= 0) return false;
                    if (score.removeError + 1 < score.keepError) return true;
                    if (score.weight >= 3 && score.rawRemoveError <= score.keepError + (feeTolerance * score.count)) return true;
                    return false;
                })
                .sort((a, b) => (b[1].keepError - b[1].removeError) - (a[1].keepError - a[1].removeError))
                .map(([tradeType]) => tradeType);

            const merged = [...new Set([...(config.valuesToRemove || []), ...recommended])];
            return { result: { valuesToRemove: merged, suggestedRemovals: recommended }, error: null };
        },
        autoDistributePhase5(data, config) {
            const removal = this.suggestPhase5Removals(data, config);
            if (removal.error) return { result: null, error: removal.error };

            const removalConfig = {
                ...config,
                valuesToRemove: removal.result.valuesToRemove
            };
            const filtered = this.transformPhase4(data, removalConfig).kept;
            const classified = this.classifyByBalance(filtered, removalConfig);
            if (classified.error) return { result: null, error: classified.error };

            return {
                result: {
                    valuesToRemove: removal.result.valuesToRemove,
                    suggestedRemovals: removal.result.suggestedRemovals,
                    depositTypes: classified.result.depositTypes,
                    withdrawalTypes: classified.result.withdrawalTypes,
                    inferredFeeMode: classified.result.inferredFeeMode
                },
                error: null
            };
        },
        classifyByBalance(data, config) {
            const validation = this.validatePhase5Config(config);
            if (!validation.isValid) {
                return { result: null, error: validation.message };
            }
            const body = data.slice(1);
            if (body.length === 0) {
                return { result: { depositTypes: [], withdrawalTypes: [], inferredFeeMode: 'separate' }, error: null };
            }

            const rows = body.map(row => this._getRowValues(row, config));
            const tolerance = 1;
            const typeStats = {};
            let prevBalance = config.baseAmount || 0;

            rows.forEach(rowInfo => {
                const tradeType = rowInfo.tradeType;
                if (!tradeType || rowInfo.balance === null) {
                    if (rowInfo.balance !== null) prevBalance = rowInfo.balance;
                    return;
                }

                const observedDelta = rowInfo.balance - prevBalance;
                prevBalance = rowInfo.balance;
                const candidates = this._buildCashflowCandidates(rowInfo).map(candidate => {
                    let score = -Math.abs(observedDelta - candidate.delta);
                    if (Math.abs(observedDelta - candidate.delta) <= tolerance) score += 5;
                    const textDirection = this._inferDirectionFromText(tradeType);
                    if (textDirection && textDirection === candidate.direction) score += 2;
                    if ((rowInfo.fee > 0 || rowInfo.tax > 0 || rowInfo.extraCost > 0) && candidate.mode === 'feeIncluded') score += 1.5;
                    if (rowInfo.amount > 0 && observedDelta !== 0 && Math.sign(observedDelta) === Math.sign(candidate.delta)) score += 1;
                    return { ...candidate, score };
                }).sort((a, b) => b.score - a.score);

                const best = candidates[0];
                typeStats[tradeType] = typeStats[tradeType] || {
                    deposit: 0,
                    withdrawal: 0,
                    amountOnly: 0,
                    feeIncluded: 0,
                    matchedRows: 0,
                    unresolvedRows: 0
                };

                if (Math.abs(observedDelta - best.delta) <= tolerance || best.score > -2) {
                    typeStats[tradeType][best.direction] += 1;
                    typeStats[tradeType][best.mode] += 1;
                    typeStats[tradeType].matchedRows += 1;
                } else {
                    typeStats[tradeType].unresolvedRows += 1;
                }
            });

            const depositTypes = [];
            const withdrawalTypes = [];
            let feeIncludedHits = 0;
            let amountOnlyHits = 0;

            Object.entries(typeStats).forEach(([tradeType, stats]) => {
                if (stats.deposit > stats.withdrawal) depositTypes.push(tradeType);
                else if (stats.withdrawal > stats.deposit) withdrawalTypes.push(tradeType);
                feeIncludedHits += stats.feeIncluded;
                amountOnlyHits += stats.amountOnly;
            });

            return {
                result: {
                    depositTypes,
                    withdrawalTypes,
                    inferredFeeMode: feeIncludedHits > amountOnlyHits ? 'included' : 'separate'
                },
                error: null
            };
        },
        transformPhase5(data, config) {
            const header = data[0] || [];
            const body = data.slice(1);
            const newSheet = [['일자', '입금', '출금', '잔액', '(참고용)원본잔액', '적요']];
            let runningBalance = config.baseAmount || 0;
            const splitCostsSeparately = config.inferredFeeMode !== 'included';
            body.forEach(row => {
                const getValue = (key) => (config[key] === undefined || row[config[key]] === null || row[config[key]] === undefined) ? "" : row[config[key]];
                const getNumeric = (key) => parseFloat(String(getValue(key)).replace(/,/g, '')) || 0;
                const tradeType = getValue('tradeTypeCol');
                const rawDate = getValue('dateCol');
                const dateObj = this._parseDate(rawDate);
                const date = this._formatDate(dateObj);
                const jukyo = getValue('jukyoCol');
                const originalBalance = getValue('balanceCol');
                const amount = getNumeric('amountCol');
                const fee = Math.abs(getNumeric('feeCol'));
                const tax = Math.abs(getNumeric('taxCol'));
                const extraCosts = this._getExtraCostDetails(row, config, header);
                let deposit = 0;
                let withdrawal = 0;

                if (config.depositTypes?.includes(tradeType)) deposit = amount;
                else if (config.withdrawalTypes?.includes(tradeType)) withdrawal = amount;

                if (deposit > 0) {
                    runningBalance += deposit;
                    newSheet.push([date, deposit, 0, runningBalance, originalBalance, jukyo]);
                } else if (withdrawal > 0) {
                    runningBalance -= withdrawal;
                    newSheet.push([date, 0, withdrawal, runningBalance, originalBalance, jukyo]);
                }

                if (splitCostsSeparately && fee > 0) {
                    runningBalance -= fee;
                    newSheet.push([date, 0, fee, runningBalance, originalBalance, `(수수료) ${jukyo}`]);
                }
                if (splitCostsSeparately && tax > 0) {
                    runningBalance -= tax;
                    newSheet.push([date, 0, tax, runningBalance, originalBalance, `(제세금) ${jukyo}`]);
                }
                if (splitCostsSeparately) extraCosts.forEach(item => {
                    runningBalance -= item.amount;
                    const labeledMemo = [ `(${item.label})`, jukyo ].filter(Boolean).join(' ');
                    newSheet.push([date, 0, item.amount, runningBalance, originalBalance, labeledMemo]);
                });
            });
            return newSheet;
        },
        generateMultiSheetDownloadLink(sheetDataArray, filename) {
            const workbook = XLSX.utils.book_new();
            sheetDataArray.forEach(sheetInfo => {
                if(sheetInfo.data && sheetInfo.data.length > 0) {
                    const worksheet = XLSX.utils.aoa_to_sheet(sheetInfo.data);
                    this.applyNumericFormatToWorksheet(worksheet);
                    const safeSheetName = sheetInfo.name.replace(/[/\\?*:[\]]/g, '').substring(0, 31);
                    XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName);
                }
            });
            if (workbook.SheetNames.length > 0) {
                XLSX.writeFile(workbook, filename);
            } else {
                 UIRenderer.showModal('알림', `엑셀 파일로 만들 데이터가 없습니다.`, [{ text: '확인', className: 'bg-indigo-600 text-white', onClick: () => UIRenderer.hideModal() }]);
            }
        },
        applyNumericFormatToWorksheet(worksheet) {
            Object.entries(worksheet || {}).forEach(([cellAddress, cell]) => {
                if (cellAddress.startsWith('!')) return;
                if (!cell || cell.t !== 'n') return;
                cell.z = '#,##0.########';
                if (cell.w) delete cell.w;
            });
        }
    };

    DataTransformer._getPhase5FinalMismatch = function(data, config) {
        const verification = this.verifyPhase5(data, config);
        if (verification.isCorrect) return 0;
        return Math.abs((verification.calculated || 0) - (verification.original || 0));
    };

    DataTransformer._summarizeTypeResiduals = function(data, config, removedTypes = []) {
        const removed = new Set(removedTypes || []);
        const body = this._getPhase5OrderedBody(data, config);
        const stats = {};
        let prevBalance = config.baseAmount || 0;
        const removalKeywords = ['수수료', '세금', '제세금', '기타수수료', '기타제세금', '이체', '대체', '정산', '환전', '출고'];

        body.forEach(row => {
            const rowInfo = this._getRowValues(row, config);
            if (rowInfo.balance === null) return;
            const observedDelta = rowInfo.balance - prevBalance;
            prevBalance = rowInfo.balance;
            const tradeType = rowInfo.tradeType;
            if (!tradeType || removed.has(tradeType)) return;

            const candidates = this._buildCashflowCandidates(rowInfo);
            const keepError = Math.min(...candidates.map(candidate => Math.abs(observedDelta - candidate.delta)));
            const removeError = Math.abs(observedDelta);
            const lowered = String(tradeType || '').toLowerCase();
            const keywordWeight = removalKeywords.reduce((max, keyword) => lowered.includes(keyword) ? Math.max(max, 2.5) : max, 1);

            stats[tradeType] = stats[tradeType] || {
                rows: 0,
                keepError: 0,
                removeError: 0,
                absDelta: 0,
                netDelta: 0,
                keywordWeight: 1
            };
            stats[tradeType].rows += 1;
            stats[tradeType].keepError += keepError;
            stats[tradeType].removeError += removeError;
            stats[tradeType].absDelta += Math.abs(observedDelta);
            stats[tradeType].netDelta += observedDelta;
            stats[tradeType].keywordWeight = Math.max(stats[tradeType].keywordWeight, keywordWeight);
        });

        return stats;
    };

    DataTransformer.autoDistributePhase5 = function(data, config) {
        const validation = this.validatePhase5Config(config);
        if (!validation.isValid) {
            return { result: null, error: validation.message };
        }

        const sortValues = (values) => [...new Set(values)].sort();
        const allTypes = config.tradeTypeCol === undefined
            ? []
            : [...new Set(data.slice(1).map(row => this._normalizeTradeTypeValue(row[config.tradeTypeCol])).filter(Boolean))];

        let removalSet = new Set(config.valuesToRemove || []);
        let depositSet = new Set(config.depositTypes || []);
        let withdrawalSet = new Set(config.withdrawalTypes || []);
        let inferredFeeMode = config.inferredFeeMode || 'separate';
        const seen = new Set();
        let iterations = 0;

        while (iterations < 5) {
            iterations += 1;
            const stateKey = JSON.stringify({
                remove: sortValues([...removalSet]),
                deposit: sortValues([...depositSet]),
                withdrawal: sortValues([...withdrawalSet]),
                inferredFeeMode
            });
            if (seen.has(stateKey)) break;
            seen.add(stateKey);

            const removalResult = this.suggestPhase5Removals(data, {
                ...config,
                valuesToRemove: sortValues([...removalSet])
            });
            if (removalResult.error) return { result: null, error: removalResult.error };
            removalResult.result.valuesToRemove.forEach(value => removalSet.add(value));

            const filteredAfterRemoval = this.transformPhase4(data, {
                ...config,
                valuesToRemove: sortValues([...removalSet])
            }).kept;
            const classified = this.classifyByBalance(filteredAfterRemoval, {
                ...config,
                valuesToRemove: sortValues([...removalSet])
            });
            if (classified.error) return { result: null, error: classified.error };

            depositSet = new Set(classified.result.depositTypes || []);
            withdrawalSet = new Set(classified.result.withdrawalTypes || []);
            inferredFeeMode = classified.result.inferredFeeMode || inferredFeeMode;
            removalSet.forEach(value => {
                depositSet.delete(value);
                withdrawalSet.delete(value);
            });

            const leftovers = allTypes.filter(value => !removalSet.has(value) && !depositSet.has(value) && !withdrawalSet.has(value));
            if (leftovers.length > 0) {
                const residuals = this._summarizeTypeResiduals(data, { ...config, inferredFeeMode }, sortValues([...removalSet]));
                leftovers.forEach(value => {
                    const stat = residuals[value];
                    if (!stat) return;
                    const removalPreferred = stat.removeError + 1 < stat.keepError;
                    const lowImpact = stat.absDelta <= (stat.rows * 10);
                    const keywordPreferred = stat.keywordWeight >= 2 && stat.removeError <= stat.keepError + (20 * stat.rows);
                    if (removalPreferred || lowImpact || keywordPreferred) {
                        removalSet.add(value);
                    }
                });
            }
        }

        let improved = true;
        while (improved) {
            improved = false;
            let finalConfig = {
                ...config,
                valuesToRemove: sortValues([...removalSet]),
                depositTypes: sortValues([...depositSet].filter(value => !removalSet.has(value))),
                withdrawalTypes: sortValues([...withdrawalSet].filter(value => !removalSet.has(value))),
                inferredFeeMode
            };
            let filtered = this.transformPhase4(data, finalConfig).kept;
            let baselineMismatch = this._getPhase5FinalMismatch(filtered, finalConfig);
            const leftovers = allTypes.filter(value => !removalSet.has(value) && !depositSet.has(value) && !withdrawalSet.has(value));

            for (const value of leftovers) {
                const trialRemovalSet = new Set(removalSet);
                trialRemovalSet.add(value);
                const trialRemovalConfig = {
                    ...config,
                    valuesToRemove: sortValues([...trialRemovalSet])
                };
                const trialFiltered = this.transformPhase4(data, trialRemovalConfig).kept;
                const trialClassified = this.classifyByBalance(trialFiltered, trialRemovalConfig);
                if (trialClassified.error) continue;

                const trialConfig = {
                    ...trialRemovalConfig,
                    depositTypes: sortValues((trialClassified.result.depositTypes || []).filter(type => !trialRemovalSet.has(type))),
                    withdrawalTypes: sortValues((trialClassified.result.withdrawalTypes || []).filter(type => !trialRemovalSet.has(type))),
                    inferredFeeMode: trialClassified.result.inferredFeeMode || inferredFeeMode
                };
                const trialMismatch = this._getPhase5FinalMismatch(trialFiltered, trialConfig);
                if (trialMismatch + 1 < baselineMismatch) {
                    removalSet = trialRemovalSet;
                    depositSet = new Set(trialConfig.depositTypes || []);
                    withdrawalSet = new Set(trialConfig.withdrawalTypes || []);
                    inferredFeeMode = trialConfig.inferredFeeMode || inferredFeeMode;
                    improved = true;
                    break;
                }
            }
        }

        const finalValuesToRemove = sortValues([...removalSet]);
        const finalDepositTypes = sortValues([...depositSet].filter(value => !removalSet.has(value)));
        const finalWithdrawalTypes = sortValues([...withdrawalSet].filter(value => !removalSet.has(value)));
        const finalConfig = {
            ...config,
            valuesToRemove: finalValuesToRemove,
            depositTypes: finalDepositTypes,
            withdrawalTypes: finalWithdrawalTypes,
            inferredFeeMode
        };
        const finalFiltered = this.transformPhase4(data, finalConfig).kept;
        const verification = this.verifyPhase5(finalFiltered, finalConfig);

        return {
            result: {
                valuesToRemove: finalValuesToRemove,
                suggestedRemovals: finalValuesToRemove,
                depositTypes: finalDepositTypes,
                withdrawalTypes: finalWithdrawalTypes,
                inferredFeeMode,
                verification,
                iterations
            },
            error: null
        };
    };

    DataTransformer._normalizeTradeTypeValue = function(value) {
        if (value === null || value === undefined) return '';
        return String(value).trim();
    };

    DataTransformer.transformPhase4 = function(data, config) {
        const header = data[0] || [];
        const body = data.slice(1);
        if (!config || config.tradeTypeCol === undefined) return { kept: data, removed: [header] };

        const removalSet = new Set((config.valuesToRemove || []).map(value => this._normalizeTradeTypeValue(value)).filter(Boolean));
        const removeAmountlessDerivedRows = config.specialRules?.deriveAmountFromTradeTypeText && config.amountCol !== undefined;
        const keptRows = [];
        const removedRows = [];

        body.forEach(row => {
            const tradeType = this._normalizeTradeTypeValue(row[config.tradeTypeCol]);
            const derivedAmount = removeAmountlessDerivedRows ? this._toNumber(row[config.amountCol]) : null;
            if (removeAmountlessDerivedRows && tradeType && (derivedAmount === null || Math.abs(derivedAmount) === 0)) removedRows.push(row);
            else if (tradeType && removalSet.has(tradeType)) removedRows.push(row);
            else keptRows.push(row);
        });

        return {
            kept: [header, ...keptRows],
            removed: [header, ...removedRows]
        };
    };

    DataTransformer._getPhase5OrderedBody = function(data, config = {}) {
        const body = data.slice(1);
        const orderMode = config.orderMode || 'original';
        if (orderMode === 'reverse') return [...body].reverse();
        return [...body];
    };

    DataTransformer._getPhase5OrderedData = function(data, config = {}) {
        if (!Array.isArray(data) || data.length === 0) return data;
        return [data[0], ...this._getPhase5OrderedBody(data, config)];
    };

    DataTransformer._buildPhase5Projection = function(data, config) {
        const header = data[0] || [];
        const body = this._getPhase5OrderedBody(data, config);
        const sheet = [['\uC77C\uC790', '\uC785\uAE08', '\uCD9C\uAE08', '\uC794\uC561', '(\uCC38\uACE0)\uC6D0\uBCF8\uC794\uC561', '\uC801\uC694']];
        const outputMeta = [];
        const depositSet = new Set((config.depositTypes || []).map(value => this._normalizeTradeTypeValue(value)).filter(Boolean));
        const withdrawalSet = new Set((config.withdrawalTypes || []).map(value => this._normalizeTradeTypeValue(value)).filter(Boolean));
        const splitCostsSeparately = config.inferredFeeMode !== 'included';
        let runningBalance = config.baseAmount || 0;

        const pushRow = (date, deposit, withdrawal, originalBalance, memo, sourceRowIndex, dateKey) => {
            sheet.push([date, deposit, withdrawal, runningBalance, originalBalance, memo]);
            outputMeta.push({
                sourceRowIndex,
                date,
                dateKey,
                sheetRowIndex: sheet.length - 1
            });
        };

        const checkpoints = body.map((row, rowIndex) => {
            const getValue = (key) => (config[key] === undefined || row[config[key]] === null || row[config[key]] === undefined) ? "" : row[config[key]];
            const getNumeric = (key) => parseFloat(String(getValue(key)).replace(/,/g, '')) || 0;
            const tradeType = this._normalizeTradeTypeValue(getValue('tradeTypeCol'));
            const rawDate = getValue('dateCol');
            const dateObj = this._parseDate(rawDate);
            const date = this._formatDate(dateObj);
            const dateKey = this._phase5DateKey(rawDate || date);
            const jukyo = getValue('jukyoCol');
            const originalBalanceText = getValue('balanceCol');
            const originalBalance = config.balanceCol === undefined ? null : this._toNumber(originalBalanceText);
            const amount = getNumeric('amountCol');
            const fee = Math.abs(getNumeric('feeCol'));
            const tax = Math.abs(getNumeric('taxCol'));
            const extraCosts = this._getExtraCostDetails(row, config, header);
            const extraCostTotal = extraCosts.reduce((sum, item) => sum + item.amount, 0);
            const rowInfo = {
                tradeType,
                amount,
                fee,
                tax,
                extraCost: extraCostTotal,
                balance: originalBalance
            };
            const beforeRowCount = sheet.length - 1;

            if (depositSet.has(tradeType) && amount > 0) {
                runningBalance += amount;
                pushRow(date, amount, 0, originalBalanceText, jukyo, rowIndex, dateKey);
            } else if (withdrawalSet.has(tradeType) && amount > 0) {
                runningBalance -= amount;
                pushRow(date, 0, amount, originalBalanceText, jukyo, rowIndex, dateKey);
            }

            if (splitCostsSeparately && fee > 0) {
                runningBalance -= fee;
                pushRow(date, 0, fee, originalBalanceText, `(\uC218\uC218\uB8CC) ${jukyo}`.trim(), rowIndex, dateKey);
            }
            if (splitCostsSeparately && tax > 0) {
                runningBalance -= tax;
                pushRow(date, 0, tax, originalBalanceText, `(\uC81C\uC138\uAE08) ${jukyo}`.trim(), rowIndex, dateKey);
            }
            if (splitCostsSeparately) {
                extraCosts.forEach(item => {
                    runningBalance -= item.amount;
                    const labeledMemo = [`(${item.label})`, jukyo].filter(Boolean).join(' ');
                    pushRow(date, 0, item.amount, originalBalanceText, labeledMemo, rowIndex, dateKey);
                });
            }

            const generatedRowCount = (sheet.length - 1) - beforeRowCount;
            const checkpointEnabled = originalBalance !== null && !this._phase5IsPlaceholderBalance(rowInfo);
            return {
                sourceRowIndex: rowIndex,
                date,
                dateKey,
                tradeType,
                originalBalance,
                calculatedBalance: runningBalance,
                diff: checkpointEnabled ? (originalBalance - runningBalance) : 0,
                isCheckpoint: checkpointEnabled,
                generatedRowCount,
                lastSheetRowIndex: generatedRowCount > 0 ? (sheet.length - 1) : beforeRowCount
            };
        });

        return {
            sheet,
            checkpoints,
            outputMeta,
            finalCalculatedBalance: runningBalance
        };
    };

    DataTransformer._getPhase5CostProfile = function(data, config) {
        const explicitCostColumns = ['feeCol', 'taxCol', 'extraCostCol1', 'extraCostCol2', 'extraCostCol3']
            .some(key => config[key] !== undefined);
        const body = this._getPhase5OrderedBody(data, config);
        let nonZeroCostRows = 0;
        let totalCost = 0;

        body.forEach(row => {
            const rowInfo = this._getRowValues(row, config);
            const rowCost = Math.abs(rowInfo.fee || 0) + Math.abs(rowInfo.tax || 0) + Math.abs(rowInfo.extraCost || 0);
            if (rowCost > 0) {
                nonZeroCostRows += 1;
                totalCost += rowCost;
            }
        });

        return {
            explicitCostColumns,
            nonZeroCostRows,
            totalCost,
            hasCostActivity: nonZeroCostRows > 0
        };
    };

    DataTransformer._summarizePhase5Projection = function(projection) {
        const checkpoints = (projection?.checkpoints || []).filter(item => item.isCheckpoint);
        const mismatches = checkpoints.filter(item => Math.abs(item.diff) > 1);
        const totalCheckpointMismatch = mismatches.reduce((sum, item) => sum + Math.abs(item.diff), 0);
        const worstMismatch = mismatches.reduce((best, item) => {
            if (!best || Math.abs(item.diff) > Math.abs(best.diff)) return item;
            return best;
        }, null);
        const firstMismatch = mismatches[0] || null;
        const finalCheckpoint = checkpoints[checkpoints.length - 1] || null;

        return {
            checkpoints,
            mismatches,
            checkpointCount: checkpoints.length,
            mismatchCount: mismatches.length,
            totalCheckpointMismatch,
            worstMismatch,
            worstDiffAbs: worstMismatch ? Math.abs(worstMismatch.diff) : 0,
            firstMismatch,
            finalCalculated: finalCheckpoint ? finalCheckpoint.calculatedBalance : (projection?.finalCalculatedBalance || 0),
            finalOriginal: finalCheckpoint ? finalCheckpoint.originalBalance : null,
            finalDiff: finalCheckpoint ? finalCheckpoint.diff : 0,
            finalDiffAbs: finalCheckpoint ? Math.abs(finalCheckpoint.diff) : 0,
            isCorrect: mismatches.length === 0
        };
    };

    DataTransformer._summarizePhase5MismatchDates = function(projection) {
        const checkpoints = (projection?.checkpoints || []).filter(item => item.isCheckpoint);
        const dailyCheckpoints = [];
        const byDate = new Map();

        checkpoints.forEach(checkpoint => {
            if (!checkpoint.date) return;
            byDate.set(checkpoint.date, checkpoint);
        });

        checkpoints.forEach(checkpoint => {
            if (!checkpoint.date) return;
            if (byDate.get(checkpoint.date) === checkpoint) {
                dailyCheckpoints.push(checkpoint);
            }
        });

        const dailyMismatches = dailyCheckpoints
            .map(checkpoint => ({
                ...checkpoint,
                diffAbs: Math.abs(checkpoint.diff)
            }))
            .filter(checkpoint => checkpoint.diffAbs > 1);

        return {
            dailyCheckpoints,
            dailyMismatchCount: dailyMismatches.length,
            dailyMismatches,
            firstDailyMismatch: dailyMismatches[0] || null,
            totalDailyMismatch: dailyMismatches.reduce((sum, item) => sum + item.diffAbs, 0)
        };
    };

    DataTransformer.transformPhase5 = function(data, config) {
        const projection = this._buildPhase5Projection(data, config);
        const summary = this._summarizePhase5Projection(projection);
        const dailySummary = this._summarizePhase5MismatchDates(projection);
        const baseSheet = projection.sheet.map(row => [...row]);
        const adjustmentMode = config.balanceAdjustmentMode;

        if (!adjustmentMode || baseSheet.length <= 1 || summary.isCorrect) {
            return baseSheet;
        }

        if (adjustmentMode === 'previous_day') {
            const diff = summary.finalDiff;
            if (Math.abs(diff) <= 1) return baseSheet;

            const firstDateText = baseSheet[1][0] || '';
            const firstDateObj = this._parseDate(firstDateText);
            let adjustmentDate = firstDateText;
            if (firstDateObj && !isNaN(firstDateObj.getTime())) {
                const prevDay = new Date(firstDateObj);
                prevDay.setDate(prevDay.getDate() - 1);
                adjustmentDate = this._formatDate(prevDay);
            }

            const adjustmentBalance = (config.baseAmount || 0) + diff;
            const adjustmentRow = diff >= 0
                ? [adjustmentDate, diff, 0, adjustmentBalance, '', '(\uC794\uC561\uBCF4\uC815) \uC99D\uAD8C\uC0AC \uC794\uC561 \uB9DE\uCDA4']
                : [adjustmentDate, 0, Math.abs(diff), adjustmentBalance, '', '(\uC794\uC561\uBCF4\uC815) \uC99D\uAD8C\uC0AC \uC794\uC561 \uB9DE\uCDA4'];

            const adjustedRows = baseSheet.slice(1).map(row => [row[0], row[1], row[2], row[3] + diff, row[4], row[5]]);
            return [baseSheet[0], adjustmentRow, ...adjustedRows];
        }

        if (adjustmentMode === 'mismatch_dates') {
            if ((dailySummary.dailyCheckpoints || []).length === 0) return baseSheet;

            let adjustedSheet = [baseSheet[0], ...baseSheet.slice(1).map(row => [...row])];
            let appliedOffset = 0;

            dailySummary.dailyCheckpoints.forEach(checkpoint => {
                const effectiveDiff = checkpoint.originalBalance - (checkpoint.calculatedBalance + appliedOffset);
                if (Math.abs(effectiveDiff) <= 1) return;

                let insertAfterIndex = 0;
                for (let i = 1; i < adjustedSheet.length; i += 1) {
                    if (String(adjustedSheet[i][0] || '').trim() === String(checkpoint.date || '').trim()) {
                        insertAfterIndex = i;
                    }
                }
                if (insertAfterIndex === 0) {
                    insertAfterIndex = adjustedSheet.length - 1;
                }

                const anchorBalance = insertAfterIndex >= 1 ? (adjustedSheet[insertAfterIndex][3] || 0) : (config.baseAmount || 0);
                const adjustmentBalance = anchorBalance + effectiveDiff;
                const adjustmentRow = effectiveDiff >= 0
                    ? [checkpoint.date, effectiveDiff, 0, adjustmentBalance, '', '(\uC794\uC561\uBCF4\uC815) \uD574\uB2F9\uC77C \uB9D0\uD589 \uC794\uC561 \uB9DE\uCDA4']
                    : [checkpoint.date, 0, Math.abs(effectiveDiff), adjustmentBalance, '', '(\uC794\uC561\uBCF4\uC815) \uD574\uB2F9\uC77C \uB9D0\uD589 \uC794\uC561 \uB9DE\uCDA4'];

                adjustedSheet.splice(insertAfterIndex + 1, 0, adjustmentRow);
                for (let i = insertAfterIndex + 2; i < adjustedSheet.length; i += 1) {
                    adjustedSheet[i][3] = (adjustedSheet[i][3] || 0) + effectiveDiff;
                }
                appliedOffset += effectiveDiff;
            });

            return adjustedSheet;
        }

        if (adjustmentMode === 'mismatch_date_end') {
            const mismatch = summary.firstMismatch || summary.worstMismatch;
            if (!mismatch || Math.abs(mismatch.diff) <= 1) return baseSheet;

            const targetDate = mismatch.date || baseSheet[baseSheet.length - 1]?.[0] || '';
            let insertAfterIndex = 0;
            for (let i = 1; i < baseSheet.length; i += 1) {
                if (String(baseSheet[i][0] || '').trim() === String(targetDate || '').trim()) {
                    insertAfterIndex = i;
                }
            }
            if (insertAfterIndex === 0) {
                insertAfterIndex = Math.max(1, mismatch.lastSheetRowIndex || (baseSheet.length - 1));
            }

            const diff = mismatch.diff;
            const anchorBalance = insertAfterIndex >= 1 ? (baseSheet[insertAfterIndex][3] || 0) : (config.baseAmount || 0);
            const adjustmentBalance = anchorBalance + diff;
            const adjustmentRow = diff >= 0
                ? [targetDate, diff, 0, adjustmentBalance, '', '(\uC794\uC561\uBCF4\uC815) \uD574\uB2F9\uC77C \uC794\uC561 \uB9DE\uCDA4']
                : [targetDate, 0, Math.abs(diff), adjustmentBalance, '', '(\uC794\uC561\uBCF4\uC815) \uD574\uB2F9\uC77C \uC794\uC561 \uB9DE\uCDA4'];

            const beforeRows = baseSheet.slice(1, insertAfterIndex + 1).map(row => [...row]);
            const afterRows = baseSheet.slice(insertAfterIndex + 1).map(row => [row[0], row[1], row[2], row[3] + diff, row[4], row[5]]);
            return [baseSheet[0], ...beforeRows, adjustmentRow, ...afterRows];
        }

        return baseSheet;
    };

    DataTransformer.getPhase5BalanceAdjustmentInfo = function(data, config) {
        const body = this._getPhase5OrderedBody(data, config);
        if (body.length === 0) return null;

        const projection = this._buildPhase5Projection(data, {
            ...config,
            balanceAdjustmentMode: null
        });
        const summary = this._summarizePhase5Projection(projection);
        const dailySummary = this._summarizePhase5MismatchDates(projection);
        if (summary.isCorrect || summary.checkpointCount === 0) return null;

        const firstRawDate = config.dateCol === undefined
            ? ''
            : body.find(row => row[config.dateCol] !== null && row[config.dateCol] !== undefined && String(row[config.dateCol]).trim() !== '')?.[config.dateCol];
        const firstDateObj = this._parseDate(firstRawDate);
        let previousDayAdjustmentDate = String(firstRawDate ?? '').trim();
        if (firstDateObj && !isNaN(firstDateObj.getTime())) {
            const prevDay = new Date(firstDateObj);
            prevDay.setDate(prevDay.getDate() - 1);
            previousDayAdjustmentDate = this._formatDate(prevDay);
        }

        const mismatch = summary.firstMismatch || summary.worstMismatch;

        return {
            diff: summary.finalDiff,
            finalCalculated: summary.finalCalculated,
            finalOriginal: summary.finalOriginal,
            checkpointMismatch: summary.totalCheckpointMismatch,
            mismatchCount: summary.mismatchCount,
            dailyMismatchCount: dailySummary.dailyMismatchCount,
            adjustmentDate: previousDayAdjustmentDate,
            previousDayAdjustmentDate,
            mismatchDate: mismatch?.date || '',
            mismatchDateDiff: mismatch?.diff || 0,
            mismatchDates: dailySummary.dailyMismatches || [],
            worstDiffAbs: summary.worstDiffAbs
        };
    };

    DataTransformer.verifyPhase5 = function(data, config) {
        const body = this._getPhase5OrderedBody(data, config);
        if (body.length === 0) {
            return {
                isCorrect: true,
                calculated: config.baseAmount || 0,
                original: config.baseAmount || 0,
                checkpointCount: 0,
                mismatchCount: 0,
                totalCheckpointMismatch: 0,
                finalDiffAbs: 0,
                worstDiffAbs: 0
            };
        }

        const projection = this._buildPhase5Projection(data, {
            ...config,
            balanceAdjustmentMode: null
        });
        const summary = this._summarizePhase5Projection(projection);

        return {
            isCorrect: summary.isCorrect,
            calculated: summary.finalCalculated,
            original: summary.finalOriginal,
            diff: summary.finalDiff,
            checkpointCount: summary.checkpointCount,
            mismatchCount: summary.mismatchCount,
            totalCheckpointMismatch: summary.totalCheckpointMismatch,
            worstDiffAbs: summary.worstDiffAbs,
            finalDiffAbs: summary.finalDiffAbs,
            firstMismatchDate: summary.firstMismatch?.date || '',
            firstMismatchDiff: summary.firstMismatch?.diff || 0,
            mismatches: summary.mismatches
        };
    };

    DataTransformer._phase5MirrorPatterns = [
        {
            name: 'sell_summary_to_detail',
            detailKeywords: ['\uC8FC\uC2DD\uB9E4\uB3C4\uCD9C\uACE0', '\uB9E4\uB3C4\uCD9C\uACE0'],
            summaryKeywords: ['\uC8FC\uC2DD\uB9E4\uB3C4\uC785\uAE08', '\uB9E4\uB3C4\uC785\uAE08'],
            detailCategory: 'deposit',
            summaryCategory: 'remove'
        },
        {
            name: 'buy_summary_to_detail',
            detailKeywords: ['\uC8FC\uC2DD\uB9E4\uC218\uC785\uACE0', '\uB9E4\uC218\uC785\uACE0'],
            summaryKeywords: ['\uC8FC\uC2DD\uB9E4\uC218\uCD9C\uAE08', '\uB9E4\uC218\uCD9C\uAE08'],
            detailCategory: 'withdrawal',
            summaryCategory: 'remove'
        }
    ];

    DataTransformer._phase5MatchesKeywords = function(text, keywords = []) {
        const normalized = this._normalizeTradeTypeValue(text).toLowerCase();
        return keywords.some(keyword => normalized.includes(String(keyword || '').toLowerCase()));
    };

    DataTransformer._phase5DateKey = function(value) {
        const parsed = this._parseDate(value);
        if (parsed && !isNaN(parsed.getTime())) return this._formatDate(parsed);
        return String(value ?? '').trim();
    };

    DataTransformer._detectPhase5MirrorOverrides = function(data, config) {
        const body = this._getPhase5OrderedBody(data, config);
        const rows = body.map((row, index) => {
            const rowInfo = this._getRowValues(row, config);
            const tradeType = this._normalizeTradeTypeValue(rowInfo.tradeType);
            const dateKey = this._phase5DateKey(rowInfo.date);
            const balance = rowInfo.balance;
            const placeholderBalance = balance === null || balance === 0;
            return {
                index,
                dateKey,
                tradeType,
                amount: Math.abs(rowInfo.amount || 0),
                fee: Math.abs(rowInfo.fee || 0),
                tax: Math.abs(rowInfo.tax || 0),
                extraCost: Math.abs(rowInfo.extraCost || 0),
                balance,
                placeholderBalance
            };
        }).filter(row => row.tradeType);

        const byDate = {};
        rows.forEach(row => {
            byDate[row.dateKey] = byDate[row.dateKey] || [];
            byDate[row.dateKey].push(row);
        });

        const forcedAssignments = {};
        const matches = [];

        Object.entries(byDate).forEach(([dateKey, dateRows]) => {
            const orderedRows = [...dateRows].sort((a, b) => a.index - b.index);

            (this._phase5MirrorPatterns || []).forEach(pattern => {
                orderedRows.forEach((summaryRow, summaryPos) => {
                    if (!this._phase5MatchesKeywords(summaryRow.tradeType, pattern.summaryKeywords)) return;
                    if (summaryRow.amount <= 0) return;

                    const detailBlock = [];
                    for (let cursor = summaryPos - 1; cursor >= 0; cursor -= 1) {
                        const candidate = orderedRows[cursor];
                        if (this._phase5MatchesKeywords(candidate.tradeType, pattern.summaryKeywords)) break;

                        if (this._phase5MatchesKeywords(candidate.tradeType, pattern.detailKeywords)) {
                            detailBlock.unshift(candidate);
                            continue;
                        }

                        if (detailBlock.length === 0) continue;
                        break;
                    }

                    if (detailBlock.length === 0) return;

                    const detailAmount = detailBlock.reduce((sum, row) => sum + row.amount, 0);
                    const detailCosts = detailBlock.reduce((sum, row) => sum + row.fee + row.tax + row.extraCost, 0);
                    const grossTolerance = Math.max(1, detailBlock.length);
                    const netTolerance = Math.max(10, detailBlock.length * 5);
                    const grossMatched = Math.abs(detailAmount - summaryRow.amount) <= grossTolerance;
                    const netMatched = Math.abs((detailAmount - detailCosts) - summaryRow.amount) <= netTolerance;
                    const detailLooksPlaceholder = detailBlock.every(row => row.placeholderBalance);
                    const summaryCarriesRealBalance = !summaryRow.placeholderBalance;
                    const blockLooksLikeSettlement = detailBlock.length >= 2 || detailLooksPlaceholder;

                    if (!(grossMatched || netMatched)) return;
                    if (!blockLooksLikeSettlement) return;
                    if (!(detailLooksPlaceholder || summaryCarriesRealBalance)) return;

                    detailBlock.forEach(row => { forcedAssignments[row.tradeType] = pattern.detailCategory; });
                    forcedAssignments[summaryRow.tradeType] = pattern.summaryCategory;

                    matches.push({
                        dateKey,
                        detailTypes: [...new Set(detailBlock.map(row => row.tradeType))],
                        summaryTypes: [summaryRow.tradeType],
                        summaryAmount: summaryRow.amount,
                        detailCount: detailBlock.length,
                        amountMatchedBy: grossMatched ? 'gross' : 'net',
                        pattern: pattern.name
                    });
                });
            });
        });

        return { forcedAssignments, matches };
    };

    DataTransformer._phase5IsPlaceholderBalance = function(row) {
        if (!row) return true;
        if (row.balance === null || row.balance === undefined) return true;
        const numericMovement = Math.abs(row.amount || 0) + Math.abs(row.fee || 0) + Math.abs(row.tax || 0) + Math.abs(row.extraCost || 0);
        if (row.balance === 0 && numericMovement > 0) return true;
        return false;
    };

    DataTransformer._detectPhase5ZeroImpactOverrides = function(data, config) {
        const body = this._getPhase5OrderedBody(data, config);
        const rows = body.map((row, seq) => {
            const rowInfo = this._getRowValues(row, config);
            return {
                seq,
                tradeType: this._normalizeTradeTypeValue(rowInfo.tradeType),
                dateKey: this._phase5DateKey(rowInfo.date),
                amount: Math.abs(rowInfo.amount || 0),
                fee: Math.abs(rowInfo.fee || 0),
                tax: Math.abs(rowInfo.tax || 0),
                extraCost: Math.abs(rowInfo.extraCost || 0),
                balance: rowInfo.balance
            };
        }).filter(row => row.tradeType);

        const blocks = [];
        let currentBlock = null;

        rows.forEach(row => {
            const sameBlock = currentBlock
                && currentBlock.tradeType === row.tradeType
                && currentBlock.dateKey === row.dateKey
                && currentBlock.lastSeq === row.seq - 1;

            if (!sameBlock) {
                currentBlock = {
                    tradeType: row.tradeType,
                    dateKey: row.dateKey,
                    startSeq: row.seq,
                    endSeq: row.seq,
                    lastSeq: row.seq,
                    rows: [row]
                };
                blocks.push(currentBlock);
                return;
            }

            currentBlock.rows.push(row);
            currentBlock.endSeq = row.seq;
            currentBlock.lastSeq = row.seq;
        });

        const findPrevReliable = (startSeq) => {
            for (let i = startSeq - 1; i >= 0; i -= 1) {
                if (!this._phase5IsPlaceholderBalance(rows[i])) return rows[i];
            }
            return null;
        };

        const findNextReliable = (endSeq) => {
            for (let i = endSeq + 1; i < rows.length; i += 1) {
                if (!this._phase5IsPlaceholderBalance(rows[i])) return rows[i];
            }
            return null;
        };

        const typeStats = {};

        blocks.forEach(block => {
            const insideReliable = block.rows.filter(row => !this._phase5IsPlaceholderBalance(row));
            const prevReliable = findPrevReliable(block.startSeq);
            const nextReliable = findNextReliable(block.endSeq);
            let impact = 'unknown';

            if (insideReliable.length > 0 && prevReliable) {
                const delta = insideReliable[insideReliable.length - 1].balance - prevReliable.balance;
                if (Math.abs(delta) <= 1) impact = 'zero';
                else if (delta > 1) impact = 'positive';
                else impact = 'negative';
            } else if (prevReliable && nextReliable) {
                const delta = nextReliable.balance - prevReliable.balance;
                if (Math.abs(delta) <= 1) impact = 'zero';
            }

            typeStats[block.tradeType] = typeStats[block.tradeType] || {
                zeroBlocks: 0,
                positiveBlocks: 0,
                negativeBlocks: 0,
                unknownBlocks: 0,
                blocks: []
            };

            if (impact === 'zero') typeStats[block.tradeType].zeroBlocks += 1;
            else if (impact === 'positive') typeStats[block.tradeType].positiveBlocks += 1;
            else if (impact === 'negative') typeStats[block.tradeType].negativeBlocks += 1;
            else typeStats[block.tradeType].unknownBlocks += 1;

            typeStats[block.tradeType].blocks.push({
                dateKey: block.dateKey,
                impact,
                rowCount: block.rows.length
            });
        });

        const forcedAssignments = {};
        const matches = [];

        Object.entries(typeStats).forEach(([tradeType, stat]) => {
            const hasDirectionalImpact = stat.positiveBlocks > 0 || stat.negativeBlocks > 0;
            if (stat.zeroBlocks > 0 && !hasDirectionalImpact) {
                forcedAssignments[tradeType] = 'remove';
                matches.push({
                    tradeType,
                    zeroBlocks: stat.zeroBlocks,
                    unknownBlocks: stat.unknownBlocks
                });
            }
        });

        return { forcedAssignments, matches };
    };

    DataTransformer._phase5RemovalKeywords = [
        '\uC218\uC218\uB8CC',
        '\uC138\uAE08',
        '\uC81C\uC138\uAE08',
        '\uAE30\uD0C0\uC218\uC218\uB8CC',
        '\uAE30\uD0C0\uC81C\uC138\uAE08',
        '\uC774\uCCB4',
        '\uB300\uCCB4',
        '\uC815\uC0B0',
        '\uD658\uC804',
        '\uCD9C\uACE0',
        '\uC5F0\uCCB4\uB8CC'
    ];

    DataTransformer._phase5BlockDepositKeywords = [
        '\uC218\uC218\uB8CC',
        '\uC138\uAE08',
        '\uC81C\uC138\uAE08',
        '\uAE30\uD0C0\uC218\uC218\uB8CC',
        '\uAE30\uD0C0\uC81C\uC138\uAE08',
        '\uCD9C\uAE08',
        '\uCD9C\uACE0',
        '\uC5F0\uCCB4\uB8CC',
        '\uB300\uCD9C\uC774\uC790'
    ];

    DataTransformer._phase5ShouldBlockDepositByText = function(text) {
        const normalized = this._normalizeTradeTypeValue(text).toLowerCase();
        return (this._phase5BlockDepositKeywords || []).some(keyword => normalized.includes(String(keyword).toLowerCase()));
    };

    DataTransformer._phase5ScoreCandidate = function(candidate, rowInfo, observedDelta, textDirection, preferredFeeMode = null) {
        const difference = Math.abs(observedDelta - candidate.delta);
        let score = -difference;

        if (difference <= 1) score += 7;
        else if (difference <= 10) score += 2;

        if (textDirection && textDirection === candidate.direction) score += 2.5;
        else if (textDirection && textDirection !== candidate.direction) score -= 1.5;

        if (candidate.direction === 'deposit' && observedDelta > 0) score += 0.75;
        if (candidate.direction === 'withdrawal' && observedDelta < 0) score += 0.75;

        const totalCost = Math.abs(rowInfo.fee || 0) + Math.abs(rowInfo.tax || 0) + Math.abs(rowInfo.extraCost || 0);
        if (totalCost > 0 && candidate.mode === 'feeIncluded') score += 1.5;
        if (preferredFeeMode === 'included' && candidate.mode === 'feeIncluded') score += 1;
        if (preferredFeeMode === 'separate' && candidate.mode === 'amountOnly') score += 1;
        if (totalCost > 0 && candidate.mode === 'feeIncluded') {
            if (difference <= 1) score += 4;
            else if (difference <= 10) score += 1.5;
        }

        if (rowInfo.amount > 0 && observedDelta !== 0 && Math.sign(observedDelta) === Math.sign(candidate.delta)) {
            score += 1;
        }

        return { ...candidate, score, error: difference };
    };

    DataTransformer._phase5EvaluateRow = function(rowInfo, observedDelta, category, preferredFeeMode = null, options = {}) {
        const tradeType = rowInfo.tradeType || '';
        const textDirection = this._inferDirectionFromText(tradeType);
        const totalCost = Math.abs(rowInfo.fee || 0) + Math.abs(rowInfo.tax || 0) + Math.abs(rowInfo.extraCost || 0);
        const totalMovement = Math.abs(rowInfo.amount || 0) + totalCost;
        const directionalEvaluations = this._buildCashflowCandidates(rowInfo)
            .map(candidate => this._phase5ScoreCandidate(candidate, rowInfo, observedDelta, textDirection, preferredFeeMode));
        const bestDirectionalEvaluation = directionalEvaluations
            .filter(candidate => candidate.direction === 'deposit' || candidate.direction === 'withdrawal')
            .sort((a, b) => {
                if (a.error !== b.error) return a.error - b.error;
                return b.score - a.score;
            })[0] || null;

        if (category === 'deposit' && this._phase5ShouldBlockDepositByText(tradeType)) {
            return {
                category,
                score: Number.NEGATIVE_INFINITY,
                error: Number.POSITIVE_INFINITY,
                delta: 0,
                mode: preferredFeeMode === 'included' ? 'feeIncluded' : 'amountOnly',
                direction: 'deposit',
                blockedByKeyword: true
            };
        }

        if (category === 'remove') {
            const lowered = String(tradeType).toLowerCase();
            const keywordWeight = (this._phase5RemovalKeywords || []).reduce((max, keyword) => {
                return lowered.includes(String(keyword).toLowerCase()) ? Math.max(max, 2.5) : max;
            }, 1);

            let score = -Math.abs(observedDelta);
            if (Math.abs(observedDelta) <= 1) score += 7;
            else if (Math.abs(observedDelta) <= 10) score += 2;
            score += (keywordWeight - 1) * 1.5;
            if (textDirection) score -= 0.5;
            if (!options.removeLocked) {
                if (Math.abs(observedDelta) > 10 && totalMovement > 0) score -= 12;
                if (bestDirectionalEvaluation) {
                    if (bestDirectionalEvaluation.error <= Math.max(1, totalMovement * 0.02)) score -= 45;
                    else if (bestDirectionalEvaluation.error <= Math.max(20, Math.abs(observedDelta) * 0.2)) score -= 20;
                    if (totalCost > 0 && bestDirectionalEvaluation.mode === 'feeIncluded' && bestDirectionalEvaluation.error <= 1) score -= 15;
                }
            }

            return {
                category,
                score,
                error: Math.abs(observedDelta),
                delta: 0,
                mode: 'removed',
                keywordWeight,
                removeLocked: !!options.removeLocked
            };
        }

        const best = directionalEvaluations
            .filter(candidate => candidate.direction === category)
            .sort((a, b) => b.score - a.score)[0];

        return best || {
            category,
            score: Number.NEGATIVE_INFINITY,
            error: Number.POSITIVE_INFINITY,
            delta: 0,
            mode: preferredFeeMode === 'included' ? 'feeIncluded' : 'amountOnly',
            direction: category
        };
    };

    DataTransformer._collectPhase5TypeStats = function(data, config, preferredFeeMode = null) {
        const body = this._getPhase5OrderedBody(data, config);
        const stats = {};
        const lockedRemovalSet = new Set((config.valuesToRemove || []).map(value => this._normalizeTradeTypeValue(value)).filter(Boolean));
        let prevBalance = config.baseAmount || 0;

        body.forEach(row => {
            const rowInfo = this._getRowValues(row, config);
            if (rowInfo.balance === null) return;

            const observedDelta = rowInfo.balance - prevBalance;
            prevBalance = rowInfo.balance;

            if (!rowInfo.tradeType) return;

            const evaluationOptions = {
                removeLocked: lockedRemovalSet.has(this._normalizeTradeTypeValue(rowInfo.tradeType))
            };
            const depositEval = this._phase5EvaluateRow(rowInfo, observedDelta, 'deposit', preferredFeeMode, evaluationOptions);
            const withdrawalEval = this._phase5EvaluateRow(rowInfo, observedDelta, 'withdrawal', preferredFeeMode, evaluationOptions);
            const removeEval = this._phase5EvaluateRow(rowInfo, observedDelta, 'remove', preferredFeeMode, evaluationOptions);

            stats[rowInfo.tradeType] = stats[rowInfo.tradeType] || {
                rows: 0,
                absObserved: 0,
                depositScore: 0,
                withdrawalScore: 0,
                removeScore: 0,
                depositError: 0,
                withdrawalError: 0,
                removeError: 0,
                depositAmountOnlyHits: 0,
                depositFeeIncludedHits: 0,
                withdrawalAmountOnlyHits: 0,
                withdrawalFeeIncludedHits: 0,
                nearZeroRows: 0,
                keywordWeight: 1,
                rowBestCounts: { remove: 0, deposit: 0, withdrawal: 0 },
                rowBestGapSum: { remove: 0, deposit: 0, withdrawal: 0 }
            };

            const stat = stats[rowInfo.tradeType];
            const rankedRowEvaluations = [
                { category: 'deposit', ...depositEval },
                { category: 'withdrawal', ...withdrawalEval },
                { category: 'remove', ...removeEval }
            ].sort((a, b) => {
                if (b.score !== a.score) return b.score - a.score;
                return a.error - b.error;
            });
            const bestRowEvaluation = rankedRowEvaluations[0];
            const secondRowEvaluation = rankedRowEvaluations[1];
            const rowGap = Math.max(0, (bestRowEvaluation?.score ?? 0) - (secondRowEvaluation?.score ?? 0));

            stat.rows += 1;
            stat.absObserved += Math.abs(observedDelta);
            stat.depositScore += depositEval.score;
            stat.withdrawalScore += withdrawalEval.score;
            stat.removeScore += removeEval.score;
            stat.depositError += depositEval.error;
            stat.withdrawalError += withdrawalEval.error;
            stat.removeError += removeEval.error;
            if (depositEval.mode === 'feeIncluded') stat.depositFeeIncludedHits += 1;
            else stat.depositAmountOnlyHits += 1;
            if (withdrawalEval.mode === 'feeIncluded') stat.withdrawalFeeIncludedHits += 1;
            else stat.withdrawalAmountOnlyHits += 1;
            if (Math.abs(observedDelta) <= 1) stat.nearZeroRows += 1;
            stat.keywordWeight = Math.max(stat.keywordWeight, removeEval.keywordWeight || 1);
            if (bestRowEvaluation?.category) {
                stat.rowBestCounts[bestRowEvaluation.category] += 1;
                stat.rowBestGapSum[bestRowEvaluation.category] += rowGap;
            }
        });

        return stats;
    };

    DataTransformer._phase5GetRowPriority = function(stat) {
        if (!stat) return null;
        const depositVotes = stat.rowBestCounts?.deposit || 0;
        const withdrawalVotes = stat.rowBestCounts?.withdrawal || 0;
        const removeVotes = stat.rowBestCounts?.remove || 0;
        const category = depositVotes >= withdrawalVotes ? 'deposit' : 'withdrawal';
        const categoryVotes = category === 'deposit' ? depositVotes : withdrawalVotes;
        const oppositeVotes = category === 'deposit' ? withdrawalVotes : depositVotes;
        const avgGap = categoryVotes > 0 ? ((stat.rowBestGapSum?.[category] || 0) / categoryVotes) : 0;
        const strong =
            categoryVotes > 0 &&
            categoryVotes >= Math.max(1, removeVotes) &&
            categoryVotes >= oppositeVotes + 1 &&
            avgGap >= 3;

        if (!strong) return null;

        return {
            category,
            votes: categoryVotes,
            removeVotes,
            oppositeVotes,
            avgGap,
            unanimousDirectional: categoryVotes === stat.rows && removeVotes === 0
        };
    };

    DataTransformer._clonePhase5BucketState = function(state) {
        return {
            remove: new Set(state.remove || []),
            deposit: new Set(state.deposit || []),
            withdrawal: new Set(state.withdrawal || [])
        };
    };

    DataTransformer._assignPhase5Type = function(state, tradeType, category) {
        state.remove.delete(tradeType);
        state.deposit.delete(tradeType);
        state.withdrawal.delete(tradeType);
        if (category === 'remove') state.remove.add(tradeType);
        if (category === 'deposit') state.deposit.add(tradeType);
        if (category === 'withdrawal') state.withdrawal.add(tradeType);
        return state;
    };

    DataTransformer._phase5BuildLockedState = function(allTypes, config) {
        const allowedTypes = new Set((allTypes || []).map(value => this._normalizeTradeTypeValue(value)).filter(Boolean));
        const state = {
            remove: new Set(),
            deposit: new Set(),
            withdrawal: new Set()
        };
        const assignments = {};
        const applyLockedValues = (values, category) => {
            (values || [])
                .map(value => this._normalizeTradeTypeValue(value))
                .filter(Boolean)
                .forEach(value => {
                    if (!allowedTypes.has(value)) return;
                    assignments[value] = category;
                    this._assignPhase5Type(state, value, category);
                });
        };

        applyLockedValues(config.depositTypes, 'deposit');
        applyLockedValues(config.withdrawalTypes, 'withdrawal');
        applyLockedValues(config.valuesToRemove, 'remove');

        return { state, assignments };
    };

    DataTransformer._phase5StateToConfig = function(baseConfig, state, inferredFeeMode) {
        const sortValues = values => [...new Set(values)].sort();
        const removalValues = sortValues([...state.remove]);
        return {
            ...baseConfig,
            valuesToRemove: removalValues,
            suggestedRemovals: removalValues,
            depositTypes: sortValues([...state.deposit].filter(value => !state.remove.has(value))),
            withdrawalTypes: sortValues([...state.withdrawal].filter(value => !state.remove.has(value))),
            inferredFeeMode
        };
    };

    DataTransformer._evaluatePhase5Distribution = function(data, baseConfig, state, preferredFeeMode = null) {
        const feeModes = preferredFeeMode ? [preferredFeeMode] : ['separate', 'included'];
        const costProfile = this._getPhase5CostProfile(data, baseConfig);
        const groupedPathCount = baseConfig.pathGroupCount || 1;
        const lockedRemovalSet = new Set((baseConfig.valuesToRemove || []).map(value => this._normalizeTradeTypeValue(value)).filter(Boolean));
        let bestEvaluation = null;

        feeModes.forEach(feeMode => {
            const config = this._phase5StateToConfig(baseConfig, state, feeMode);
            const body = this._getPhase5OrderedBody(data, config);
            let prevBalance = config.baseAmount || 0;
            let rowMismatch = 0;
            let unassignedRows = 0;
            let amountOnlyHits = 0;
            let feeIncludedHits = 0;

            body.forEach(row => {
                const rowInfo = this._getRowValues(row, config);
                if (rowInfo.balance === null) return;

                const observedDelta = rowInfo.balance - prevBalance;
                prevBalance = rowInfo.balance;

                if (!rowInfo.tradeType) return;
                const evaluationOptions = {
                    removeLocked: lockedRemovalSet.has(this._normalizeTradeTypeValue(rowInfo.tradeType))
                };

                let category = null;
                if (state.remove.has(rowInfo.tradeType)) category = 'remove';
                else if (state.deposit.has(rowInfo.tradeType)) category = 'deposit';
                else if (state.withdrawal.has(rowInfo.tradeType)) category = 'withdrawal';

                if (!category) {
                    const bestFallback = ['remove', 'deposit', 'withdrawal']
                        .map(candidate => this._phase5EvaluateRow(rowInfo, observedDelta, candidate, feeMode, evaluationOptions))
                        .sort((a, b) => b.score - a.score)[0];
                    rowMismatch += bestFallback.error + 25;
                    unassignedRows += 1;
                    if (bestFallback.mode === 'feeIncluded') feeIncludedHits += 1;
                    if (bestFallback.mode === 'amountOnly') amountOnlyHits += 1;
                    return;
                }

                const evaluation = this._phase5EvaluateRow(rowInfo, observedDelta, category, feeMode, evaluationOptions);
                rowMismatch += evaluation.error;
                if (evaluation.mode === 'feeIncluded') feeIncludedHits += 1;
                if (evaluation.mode === 'amountOnly') amountOnlyHits += 1;
            });

            const filtered = this.transformPhase4(data, config).kept;
            const verification = this.verifyPhase5(filtered, {
                ...config,
                balanceAdjustmentMode: null
            });
            const checkpointMismatch = verification.totalCheckpointMismatch || 0;
            const mismatchCount = verification.mismatchCount || 0;
            const worstCheckpointMismatch = verification.worstDiffAbs || 0;
            const finalMismatch = verification.finalDiffAbs || 0;
            const totalScore =
                rowMismatch +
                (checkpointMismatch * 12) +
                (worstCheckpointMismatch * 6) +
                (finalMismatch * 3) +
                (mismatchCount * 250) +
                (unassignedRows * 50);
            const includedPenalty = (feeMode === 'included' && groupedPathCount > 1 && costProfile.hasCostActivity)
                ? Math.max(120, costProfile.nonZeroCostRows * 40)
                : 0;

            const evaluation = {
                config,
                verification,
                checkpointMismatch,
                mismatchCount,
                worstCheckpointMismatch,
                finalMismatch,
                rowMismatch,
                unassignedRows,
                totalScore: totalScore + includedPenalty,
                inferredFeeMode: feeMode,
                amountOnlyHits,
                feeIncludedHits,
                includedPenalty,
                costProfile
            };

            if (!bestEvaluation || evaluation.totalScore < bestEvaluation.totalScore) {
                bestEvaluation = evaluation;
            }
        });

        return bestEvaluation;
    };

    DataTransformer._seedPhase5Buckets = function(allTypes, typeStats, config) {
        const allowedTypes = new Set(allTypes || []);
        const seeded = {
            remove: new Set((config.valuesToRemove || []).filter(value => allowedTypes.has(value))),
            deposit: new Set((config.depositTypes || []).filter(value => allowedTypes.has(value))),
            withdrawal: new Set((config.withdrawalTypes || []).filter(value => allowedTypes.has(value)))
        };

        allTypes.forEach(tradeType => {
            if (seeded.remove.has(tradeType) || seeded.deposit.has(tradeType) || seeded.withdrawal.has(tradeType)) {
                return;
            }

            const stat = typeStats[tradeType];
            if (!stat) return;
            const rowPriority = this._phase5GetRowPriority(stat);
            if (rowPriority) {
                this._assignPhase5Type(seeded, tradeType, rowPriority.category);
                return;
            }

            const ranked = [
                { category: 'remove', score: stat.removeScore, error: stat.removeError },
                { category: 'deposit', score: stat.depositScore, error: stat.depositError },
                { category: 'withdrawal', score: stat.withdrawalScore, error: stat.withdrawalError }
            ].sort((a, b) => b.score - a.score);

            let winner = ranked[0].category;
            const bestNonRemove = ranked.find(item => item.category !== 'remove');
            const removeShouldWait =
                winner === 'remove' &&
                bestNonRemove &&
                stat.removeError > (bestNonRemove.error + Math.max(5, stat.rows * 2)) &&
                stat.absObserved > (stat.rows * 20) &&
                stat.nearZeroRows < stat.rows &&
                stat.keywordWeight < 2.3;

            if (removeShouldWait) {
                winner = bestNonRemove.category;
            }

            this._assignPhase5Type(seeded, tradeType, winner);
        });

        return seeded;
    };

    DataTransformer._seedPhase5BucketsByError = function(allTypes, typeStats, config) {
        const allowedTypes = new Set(allTypes || []);
        const seeded = {
            remove: new Set((config.valuesToRemove || []).filter(value => allowedTypes.has(value))),
            deposit: new Set((config.depositTypes || []).filter(value => allowedTypes.has(value))),
            withdrawal: new Set((config.withdrawalTypes || []).filter(value => allowedTypes.has(value)))
        };

        allTypes.forEach(tradeType => {
            if (seeded.remove.has(tradeType) || seeded.deposit.has(tradeType) || seeded.withdrawal.has(tradeType)) {
                return;
            }

            const stat = typeStats[tradeType];
            if (!stat) return;
            const rowPriority = this._phase5GetRowPriority(stat);
            if (rowPriority) {
                this._assignPhase5Type(seeded, tradeType, rowPriority.category);
                return;
            }

            const ranked = [
                { category: 'remove', error: stat.removeError, score: stat.removeScore },
                { category: 'deposit', error: stat.depositError, score: stat.depositScore },
                { category: 'withdrawal', error: stat.withdrawalError, score: stat.withdrawalScore }
            ].sort((a, b) => {
                if (a.error !== b.error) return a.error - b.error;
                return b.score - a.score;
            });

            let winner = ranked[0].category;
            if (winner === 'remove') {
                const bestDirectional = ranked.find(item => item.category !== 'remove');
                if (
                    bestDirectional &&
                    stat.absObserved > Math.max(10, stat.rows * 5) &&
                    ranked[0].error >= (bestDirectional.error - Math.max(1, stat.rows))
                ) {
                    winner = bestDirectional.category;
                }
            }

            this._assignPhase5Type(seeded, tradeType, winner);
        });

        return seeded;
    };

    DataTransformer._optimizePhase5Scenario = function(data, config, initialState, typeStats, forcedAssignments = {}, preferredFeeMode = null, lockedAssignments = {}) {
        let state = this._clonePhase5BucketState(initialState);
        const combinedForcedAssignments = {
            ...(forcedAssignments || {}),
            ...(lockedAssignments || {})
        };

        Object.entries(combinedForcedAssignments).forEach(([tradeType, category]) => {
            this._assignPhase5Type(state, tradeType, category);
        });

        let evaluation = this._evaluatePhase5Distribution(data, config, state, preferredFeeMode);
        let iterations = 0;
        let changed = true;
        const allTypes = Object.keys(typeStats || {});

        while (changed && iterations < 10) {
            iterations += 1;
            changed = false;

            const orderedTypes = [...allTypes].sort((a, b) => {
                const statA = typeStats[a] || {};
                const statB = typeStats[b] || {};
                const gapA = Math.abs(Math.max(statA.depositScore || 0, statA.withdrawalScore || 0, statA.removeScore || 0) - Math.min(statA.depositScore || 0, statA.withdrawalScore || 0, statA.removeScore || 0));
                const gapB = Math.abs(Math.max(statB.depositScore || 0, statB.withdrawalScore || 0, statB.removeScore || 0) - Math.min(statB.depositScore || 0, statB.withdrawalScore || 0, statB.removeScore || 0));
                return gapA - gapB;
            });

            orderedTypes.forEach(tradeType => {
                if (combinedForcedAssignments[tradeType]) return;
                const stat = typeStats[tradeType] || {};
                const rowPriority = this._phase5GetRowPriority(stat);

                const currentCategory = state.remove.has(tradeType)
                    ? 'remove'
                    : state.deposit.has(tradeType)
                        ? 'deposit'
                        : state.withdrawal.has(tradeType)
                            ? 'withdrawal'
                            : 'remove';

                if (rowPriority && currentCategory === 'remove') {
                    const priorityState = this._clonePhase5BucketState(state);
                    this._assignPhase5Type(priorityState, tradeType, rowPriority.category);
                    const priorityEvaluation = this._evaluatePhase5Distribution(data, config, priorityState, preferredFeeMode);
                    state = priorityState;
                    evaluation = priorityEvaluation;
                    changed = true;
                    return;
                }

                let bestCategory = currentCategory;
                let bestEvaluation = evaluation;

                ['remove', 'deposit', 'withdrawal'].forEach(candidateCategory => {
                    if (candidateCategory === currentCategory) return;
                    if (rowPriority && candidateCategory === 'remove') return;

                    const trialState = this._clonePhase5BucketState(state);
                    this._assignPhase5Type(trialState, tradeType, candidateCategory);
                    const trialEvaluation = this._evaluatePhase5Distribution(data, config, trialState, preferredFeeMode);

                    if (trialEvaluation.totalScore + 0.5 < bestEvaluation.totalScore) {
                        bestEvaluation = trialEvaluation;
                        bestCategory = candidateCategory;
                    }
                });

                if (bestCategory !== currentCategory) {
                    this._assignPhase5Type(state, tradeType, bestCategory);
                    evaluation = bestEvaluation;
                    changed = true;
                }
            });
        }

        return {
            state,
            evaluation,
            iterations
        };
    };

    DataTransformer.autoDistributePhase5 = function(data, config) {
        const validation = this.validatePhase5Config(config);
        if (!validation.isValid) {
            return { result: null, error: validation.message };
        }

        const allTypes = config.tradeTypeCol === undefined
            ? []
            : [...new Set(this._getPhase5OrderedBody(data, config).map(row => this._normalizeTradeTypeValue(row[config.tradeTypeCol])).filter(Boolean))];

        if (allTypes.length === 0) {
            return {
                result: {
                    valuesToRemove: [],
                    suggestedRemovals: [],
                    depositTypes: [],
                    withdrawalTypes: [],
                    inferredFeeMode: config.inferredFeeMode || 'separate',
                    verification: { isCorrect: true },
                    iterations: 0
                },
                error: null
            };
        }

        const preferredFeeMode = null;
        const typeStats = this._collectPhase5TypeStats(data, config, preferredFeeMode);
        const lockedState = this._phase5BuildLockedState(allTypes, config);
        const zeroImpactOverrides = this._detectPhase5ZeroImpactOverrides(data, config);
        const mirrorOverrides = this._detectPhase5MirrorOverrides(data, config);
        const forcedAssignments = {
            ...(zeroImpactOverrides.forcedAssignments || {}),
            ...(mirrorOverrides.forcedAssignments || {})
        };
        const scenarioSeeds = [
            this._seedPhase5Buckets(allTypes, typeStats, config),
            this._seedPhase5BucketsByError(allTypes, typeStats, config),
            {
                remove: new Set((config.valuesToRemove || []).filter(value => allTypes.includes(value))),
                deposit: new Set((config.depositTypes || []).filter(value => allTypes.includes(value))),
                withdrawal: new Set((config.withdrawalTypes || []).filter(value => allTypes.includes(value)))
            }
        ];

        const seenSignatures = new Set();
        const optimizedScenarios = scenarioSeeds.map(seedState => {
            const signature = JSON.stringify({
                remove: [...(seedState.remove || [])].sort(),
                deposit: [...(seedState.deposit || [])].sort(),
                withdrawal: [...(seedState.withdrawal || [])].sort()
            });
            if (seenSignatures.has(signature)) return null;
            seenSignatures.add(signature);
            return this._optimizePhase5Scenario(data, config, seedState, typeStats, forcedAssignments, preferredFeeMode, lockedState.assignments);
        }).filter(Boolean);

        let bestScenario = optimizedScenarios.sort((a, b) => a.evaluation.totalScore - b.evaluation.totalScore)[0];
        let state = bestScenario.state;
        let evaluation = bestScenario.evaluation;
        let iterations = optimizedScenarios.reduce((max, scenario) => Math.max(max, scenario.iterations), 0);
        const scenarioCount = optimizedScenarios.length;
        const costProfile = this._getPhase5CostProfile(data, config);
        const forceSeparateCosts = config.specialRules?.forceSeparateCosts === true;
        const shouldForceSeparateCosts =
            evaluation.inferredFeeMode === 'included' &&
            (forceSeparateCosts || costProfile.hasCostActivity);

        if (shouldForceSeparateCosts) {
            evaluation = this._evaluatePhase5Distribution(data, config, state, 'separate');
        }

        let finalConfig = this._phase5StateToConfig(config, state, (shouldForceSeparateCosts ? 'separate' : evaluation.inferredFeeMode) || config.inferredFeeMode || 'separate');
        const assignedTypes = new Set([
            ...finalConfig.valuesToRemove,
            ...finalConfig.depositTypes,
            ...finalConfig.withdrawalTypes
        ].map(value => this._normalizeTradeTypeValue(value)).filter(Boolean));
        const leftoverTypes = allTypes.filter(value => !assignedTypes.has(this._normalizeTradeTypeValue(value)));

        if (leftoverTypes.length > 0 && evaluation.verification?.isCorrect) {
            finalConfig = {
                ...finalConfig,
                valuesToRemove: [...new Set([...finalConfig.valuesToRemove, ...leftoverTypes.map(value => this._normalizeTradeTypeValue(value))])].sort(),
                suggestedRemovals: [...new Set([...finalConfig.valuesToRemove, ...leftoverTypes.map(value => this._normalizeTradeTypeValue(value))])].sort()
            };
            evaluation = this._evaluatePhase5Distribution(data, config, {
                remove: new Set(finalConfig.valuesToRemove),
                deposit: new Set(finalConfig.depositTypes),
                withdrawal: new Set(finalConfig.withdrawalTypes)
            }, finalConfig.inferredFeeMode);
        }

        return {
            result: {
                valuesToRemove: finalConfig.valuesToRemove,
                suggestedRemovals: finalConfig.valuesToRemove,
                depositTypes: finalConfig.depositTypes,
                withdrawalTypes: finalConfig.withdrawalTypes,
                inferredFeeMode: finalConfig.inferredFeeMode,
                verification: evaluation.verification,
                iterations,
                scenarioCount,
                rowMismatch: evaluation.rowMismatch,
                checkpointMismatch: evaluation.checkpointMismatch,
                mismatchCount: evaluation.mismatchCount,
                finalMismatch: evaluation.finalMismatch,
                forcedSeparateCosts: shouldForceSeparateCosts,
                mirrorMatches: mirrorOverrides.matches || [],
                zeroImpactMatches: zeroImpactOverrides.matches || []
            },
            error: null
        };
    };

    const BrokerAutoMapper = {
        requiredKeys: ['dateCol', 'tradeTypeCol', 'amountCol', 'balanceCol'],
        presets: [
            { id: 'kyobo', name: '교보증권', aliases: ['교보증권'], profiles: [{ columns: { dateCol: ['거래일자'], tradeTypeCol: ['적요명'], amountCol: ['거래금액'], balanceCol: ['예수금잔고'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명(거래상대명)'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], stockBalanceCol: ['유가증권잔고'] }, depositTypes: ['정기예탁금이용료입금', '펀드환매예약매도'], removalTypes: ['공모주입고', '타사대체출고'] }] },
            { id: 'daol', name: '다올투자증권', aliases: ['다올투자증권'], profiles: [] },
            { id: 'daishin', name: '대신증권', aliases: ['대신증권'], profiles: [
                {
                    columns: { dateCol: ['거래일자'], tradeTypeCol: ['적요명'], amountCol: ['거래금액'], balanceCol: ['예수금잔고'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가잔고'] },
                    depositTypes: ['개별상품대체입금', '예탁금이용료', '일괄매도'],
                    withdrawalTypes: ['개별상품대체출금', '계좌대체출고', '현금매수', '이체출금'],
                    removalTypes: ['공모주', '타사대체출고', '계좌대체입고', '타사대체입고', '현금매도'],
                    specialRules: {
                        ignoreTradeTypeContains: ['일괄매수'],
                        amountOverrideByTradeType: [{ match: '일괄매도', sourceHeaders: ['이자'] }],
                        forceSeparateCosts: true
                    }
                },
                {
                    columns: { dateCol: ['거래일자'], tradeTypeCol: ['적요명'], amountCol: ['거래금액'], balanceCol: ['예수금잔액'], feeCol: ['세금/수수료'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['채권단가'], stockBalanceCol: ['잔고수량'] },
                    depositTypes: ['개별상품대체입금', '예탁금이용료', '일괄매도'],
                    withdrawalTypes: ['개별상품대체출금', '계좌대체출고', '현금매수', '이체출금'],
                    removalTypes: ['공모주', '타사대체출고', '계좌대체입고', '타사대체입고', '현금매도'],
                    specialRules: {
                        ignoreTradeTypeContains: ['일괄매수'],
                        amountOverrideByTradeType: [{ match: '일괄매도', sourceHeaders: ['이자'] }],
                        forceSeparateCosts: true
                    }
                }
            ] },
            { id: 'db', name: 'DB금융투자', aliases: ['DB금융투자', '디비금융투자'], profiles: [
                { columns: { dateCol: ['일자'], tradeTypeCol: ['거래구분'], amountCol: ['입출금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['세금'], stockNameCol: ['종목명'], stockQtyCol: ['거래좌수'], unitPriceCol: ['기준가'], stockBalanceCol: ['잔고좌수'] }, depositTypes: ['예탁금이용료입금', '결산분배금입금'], withdrawalTypes: ['재투자'], removalTypes: [] },
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금잔'], feeCol: ['수수료'], taxCol: ['세금'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가금잔'] }, depositTypes: ['예탁금이용료입금', '결산분배금입금'], withdrawalTypes: ['재투자'], removalTypes: [] }
            ] },
            { id: 'ds', name: 'DS투자증권', aliases: ['DS투자증권', '디에스투자증권'], profiles: [] },
            { id: 'leading', name: '리딩투자증권', aliases: ['리딩투자증권'], profiles: [] },
            { id: 'meritz', name: '메리츠증권', aliases: ['메리츠증권'], profiles: [{ columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래적요'], amountCol: ['거래금액'], balanceCol: ['예수금잔고'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], stockBalanceCol: ['유가증권잔고'] }, removalTypes: ['공모주입고', '타사대체출고'] }] },
            { id: 'mirae', name: '미래에셋증권', aliases: ['미래에셋증권'], profiles: [{ columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['제세금합'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['주식매도출고', '예탁금이용료입금'], withdrawalTypes: ['은행이체대체송금'], removalTypes: ['주식매도입금', '공모주입고', '이체입고'] }] },
            { id: 'bookook', name: '부국증권', aliases: ['부국증권'], profiles: [] },
            { id: 'bnk', name: 'BNK투자증권', aliases: ['BNK투자증권', '비엔케이투자증권'], profiles: [{ columns: { dateCol: ['일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금잔'], feeCol: ['수수료'], taxCol: ['거래세'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['단가'], stockBalanceCol: ['유가금잔'] }, removalTypes: [] }] },
            { id: 'samsung', name: '삼성증권', aliases: ['삼성증권'], profiles: [
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래명'], amountCol: ['거래금액'], balanceCol: ['현금잔액'], feeCol: ['수수료/Fee'], taxCol: ['제세금/대출이자'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['잔고수량/펀드평가금액'] }, depositTypes: ['이용료입금'], removalTypes: ['청약입고', '타사출고'] },
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래명'], amountCol: ['거래금액'], balanceCol: ['현금잔액'], feeCol: ['수수료/Fee'], taxCol: ['제세금/대출이자'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['잔고수량'] }, depositTypes: ['이용료입금'], removalTypes: ['청약입고', '타사출고'] }
            ] },
            { id: 'sangsangin', name: '상상인증권', aliases: ['상상인증권'], profiles: [] },
            { id: 'shinyoung', name: '신영증권', aliases: ['신영증권'], profiles: [{ columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래적요'], amountCol: ['거래금액'], balanceCol: ['RP+예수금'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], stockBalanceCol: ['잔량'] }, removalTypes: ['타사이체출고'] }] },
            { id: 'shinhan', name: '신한투자증권', aliases: ['신한투자증권'], profiles: [
                { columns: { dateCol: ['일자'], tradeTypeCol: ['구분'], amountCol: ['거래대금'], balanceCol: ['최종금액'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['가격'] }, depositTypes: ['예탁금이용료', '장내_매도', '코스닥_매도', '환전입금', '배당금'], withdrawalTypes: ['장내_매수', '코스닥_매수', '은행이체출금'], removalTypes: ['공모주입고', '타사대체출고'] },
                { columns: { dateCol: ['일자'], tradeTypeCol: ['구분'], amountCol: ['거래대금'], balanceCol: ['최종금액'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명(상대처)'], stockQtyCol: ['수량'], unitPriceCol: ['가격'] }, depositTypes: ['예탁금이용료', '장내_매도', '코스닥_매도', '환전입금', '배당금'], withdrawalTypes: ['장내_매수', '코스닥_매수', '은행이체출금'], removalTypes: ['공모주입고', '타사대체출고'] }
            ] },
            { id: 'si', name: 'SI증권', aliases: ['SI증권', '에스아이증권'], profiles: [] },
            { id: 'sk', name: 'SK증권', aliases: ['SK증권', '에스케이증권'], profiles: [{ columns: { dateCol: ['거래일'], tradeTypeCol: ['거래종류'], amountCol: ['세전금액'], balanceCol: ['예수금잔고'], feeCol: ['수수료'], taxCol: ['거래세+농특세'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], stockBalanceCol: ['유가증권잔고'] }, depositTypes: ['예탁금이용료(원화)', '사채원금'], withdrawalTypes: ['이체-출금'], removalTypes: [] }] },
            { id: 'nh', name: 'NH투자증권', aliases: ['NH투자증권', '엔에이치투자증권'], profiles: [
                { columns: { dateCol: ['실거래일자'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['잔고금액'], feeCol: ['수수료'], taxCol: ['세금'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['가격'], stockBalanceCol: ['잔고수량'] }, depositTypes: ['은행이체입금', 'KOSDAQ매도', '예탁금이용료'], withdrawalTypes: ['IPO청약수수료출금', '공모청약출금', '업무수수료출금'], removalTypes: ['공모주입고', '타사대체출고'] },
                { columns: { dateCol: ['실거래일자'], tradeTypeCol: ['상세내용'], amountCol: ['거래금액'], balanceCol: ['잔고금액'], feeCol: ['수수료'], taxCol: ['세금'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], stockBalanceCol: ['잔고'] }, depositTypes: ['은행이체입금', 'KOSDAQ매도', '예탁금이용료'], withdrawalTypes: ['IPO청약수수료출금', '공모청약출금', '업무수수료출금'], removalTypes: ['공모주입고', '타사대체출고'] },
                { columns: { dateCol: ['실거래일자'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['잔고금액'] }, depositTypes: ['은행이체입금', 'KOSDAQ매도', '예탁금이용료'], withdrawalTypes: ['IPO청약수수료출금', '공모청약출금', '업무수수료출금'], removalTypes: ['공모주입고', '타사대체출고'], minMatchedKeys: 4 }
            ] },
            { id: 'ls', name: 'LS증권', aliases: ['LS증권', '엘에스증권', '이베스트투자증권'], profiles: [{ columns: { dateCol: ['결제일자'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['예수금잔고'], feeCol: ['수수료'], taxCol: ['제세금합'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['KOSDAQ보통매도_KRX', '예탁금이용료입금'], withdrawalTypes: ['은행이체출금'], removalTypes: ['타사대체입고'] }] },
            { id: 'woori', name: '우리투자증권', aliases: ['우리투자증권', '한국포스증권', '우리종합금융'], profiles: [{ columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['펀드명'], stockQtyCol: ['거래좌수'], unitPriceCol: ['기준가'], stockBalanceCol: ['잔고좌수'] }, removalTypes: [] }] },
            { id: 'yuanta', name: '유안타증권', aliases: ['유안타증권'], profiles: [{ columns: { dateCol: ['거래일'], tradeTypeCol: ['적요'], amountCol: ['거래금액'], balanceCol: ['예수금 잔액'], feeCol: ['수수료'], taxCol: ['제세금(거래세)'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['잔 고'] }, removalTypes: [] }] },
            { id: 'eugene', name: '유진투자증권', aliases: ['유진투자증권'], profiles: [
                { columns: { dateCol: ['일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['잔액'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['매매가격'], stockBalanceCol: ['잔고'] }, depositTypes: ['공모불입(현금)', '입금', '정기예탁금이용료'], withdrawalTypes: ['공모주 추가납입', '청약수수료 징수'], removalTypes: ['출고', '공모주', '공모불입(현금)'] },
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['예수금잔고'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], stockBalanceCol: ['유가증권잔고'] }, depositTypes: ['공모불입(현금)', '입금', '정기예탁금이용료'], withdrawalTypes: ['공모주 추가납입', '청약수수료 징수'], removalTypes: ['출고', '공모주', '공모불입(현금)'] }
            ] },
            { id: 'yuhwa', name: '유화증권', aliases: ['유화증권'], profiles: [] },
            { id: 'im', name: 'iM증권', aliases: ['iM증권', '아이엠증권', '하이투자증권'], profiles: [] },
            { id: 'ibk', name: 'IBK투자증권', aliases: ['IBK투자증권', '아이비케이투자증권'], profiles: [{ columns: { dateCol: ['결제일자'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['예수금잔고'], feeCol: ['거래수수료'], taxCol: ['제세금합'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['예탁금이용료입금'], removalTypes: ['공모주입고', '타사대체출고'] }] },
            { id: 'kakaopay', name: '카카오페이증권', aliases: ['카카오페이증권'], profiles: [] },
            { id: 'kb', name: 'KB증권', aliases: ['KB증권', '케이비증권'], profiles: [
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['거래세 등'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['은행이체 입금'], withdrawalTypes: ['기관납입 출금', '타사대체수수료출금'], removalTypes: ['공모주 입고', '타사대체 출고'] },
                {
                    columns: { dateCol: ['거래일'], tradeTypeCol: ['내용'], balanceCol: ['예수금잔액(원)'], stockNameCol: ['종목명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'], kbDepositAmountCol: ['입금/입고/매도', '입금 / 입고 / 매도'], kbWithdrawalAmountCol: ['출금/출고/매수', '출금 / 출고 / 매수'] },
                    depositTypes: ['은행이체 입금'],
                    withdrawalTypes: ['기관납입 출금', '타사대체수수료출금'],
                    removalTypes: ['공모주 입고', '타사대체 출고'],
                    specialRules: {
                        deriveAmountFromTradeTypeText: true,
                        classificationKeywords: {
                            deposit: ['입금', '입고', '매도'],
                            withdrawal: ['출금', '출고', '매수']
                        },
                        unmatchedClassification: 'remove',
                        notice: 'KB증권 간편 포맷으로 인식되었습니다. 거래금액 열이 없어도 진행되며, 입금/입고/매도 또는 출금/출고/매수 컬럼의 금액만 거래금액으로 사용합니다. 잔액 차이로 금액을 자동 추정하지 않습니다.'
                    }
                }
            ] },
            { id: 'kr', name: 'KR투자증권', aliases: ['KR투자증권', '케이알투자증권'], profiles: [] },
            { id: 'cape', name: '케이프투자증권', aliases: ['케이프투자증권'], profiles: [] },
            { id: 'koreaasset', name: '코리아에셋투자증권', aliases: ['코리아에셋투자증권'], profiles: [] },
            { id: 'kiwoom', name: '키움증권', aliases: ['키움증권'], profiles: [
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['거래세/농특세'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['예탁금이용료(이자)입금', '자금부입금법인청약', '예탁금이용료(이자)', '배당금입금', '수익분배금입금', '이체입금(연계은행)', '장내매도', '세금외화매도'], withdrawalTypes: ['공모불입출금', '수요예측청약수수료출금', '이체출금', '장내매수', '예탁금이용료이자세금출금', 'KOSDAQ매수'], removalTypes: ['공모주입고', '타사대체출고'] },
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['거래세/농특세'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['예탁금이용료(이자)입금', '자금부입금법인청약', '예탁금이용료(이자)', '배당금입금', '수익분배금입금', '이체입금(연계은행)', '장내매도', '세금외화매도'], withdrawalTypes: ['공모불입출금', '수요예측청약수수료출금', '이체출금', '장내매수', '예탁금이용료이자세금출금', 'KOSDAQ매수'], removalTypes: ['공모주입고', '타사대체출고'] },
                { columns: { dateCol: ['일자'], tradeTypeCol: ['거래구분'], amountCol: ['입출금액'], balanceCol: ['예수금잔고'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명'], stockQtyCol: ['거래좌수'], unitPriceCol: ['기준가'], stockBalanceCol: ['잔고좌수'] }, depositTypes: ['예탁금이용료(이자)입금', '자금부입금법인청약', '예탁금이용료(이자)', '배당금입금', '수익분배금입금', '이체입금(연계은행)', '장내매도', '세금외화매도'], withdrawalTypes: ['공모불입출금', '수요예측청약수수료출금', '이체출금', '장내매수', '예탁금이용료이자세금출금', 'KOSDAQ매수'], removalTypes: ['공모주입고', '타사대체출고'] },
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['거래세/농특세'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['예탁금이용료(이자)입금', '자금부입금법인청약', '예탁금이용료(이자)', '배당금입금', '수익분배금입금', '이체입금(연계은행)', '장내매도', '세금외화매도'], withdrawalTypes: ['공모불입출금', '수요예측청약수수료출금', '이체출금', '장내매수', '예탁금이용료이자세금출금', 'KOSDAQ매수'], removalTypes: ['공모주입고', '타사대체출고'] }
            ] },
            { id: 'toss', name: '토스증권', aliases: ['토스증권'], profiles: [] },
            { id: 'hana', name: '하나증권', aliases: ['하나증권'], profiles: [{ columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금잔액'], feeCol: ['수수료'], taxCol: ['거래/소득세'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가증권금잔'] }, depositTypes: ['예탁금이용료'], removalTypes: ['타사대체출고', '공모주입고'] }] },
            { id: 'sc', name: '한국스탠다드차타드증권', aliases: ['한국스탠다드차타드증권'], profiles: [] },
            { id: 'koreainvest', name: '한국투자증권', aliases: ['한국투자증권'], profiles: [
                { columns: { dateCol: ['거래일'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['거래세'], extraCostCol1: ['세금'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['HTS코스닥주식매도', '타사이체입금', '예탁금이용료', '펀드현금배당'], withdrawalTypes: ['HTS타사이체출금', 'WTS추납대체청약', '기관일괄대체청약납입', '출고수수료'], removalTypes: ['공모주입고', '타사이체출고'] },
                { columns: { dateCol: ['거래일'], tradeTypeCol: ['거래종류'], amountCol: ['거래금액'], balanceCol: ['잔액'], feeCol: ['수수료'], taxCol: ['거래세'], extraCostCol1: ['세금'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가'], stockBalanceCol: ['유가잔고'] }, depositTypes: ['HTS코스닥주식매도', '타사이체입금', '예탁금이용료', '펀드현금배당'], withdrawalTypes: ['HTS타사이체출금', 'WTS추납대체청약', '기관일괄대체청약납입', '출고수수료'], removalTypes: ['공모주입고', '타사이체출고'] }
            ] },
            { id: 'hanyang', name: '한양증권', aliases: ['한양증권'], profiles: [{ columns: { dateCol: ['일자'], tradeTypeCol: ['거래구분'], amountCol: ['입출구분'], balanceCol: ['예수금잔'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명'], stockQtyCol: ['거래좌수'], unitPriceCol: ['기준가'], stockBalanceCol: ['잔고좌수'] }, depositTypes: ['결산분배금입금'], withdrawalTypes: ['은행이체출금'], removalTypes: [] }] },
            { id: 'hanhwa', name: '한화투자증권', aliases: ['한화투자증권'], profiles: [
                { columns: { dateCol: ['거래일'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금'], feeCol: ['수수료'], taxCol: ['제세금합'], stockNameCol: ['종목(상품)명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'] }, depositTypes: ['[입금] 예탁금이용료입금'], removalTypes: [] },
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['원화예수금'], feeCol: ['수수료'], taxCol: ['제세금합'], stockNameCol: ['종목(상품)명'], stockQtyCol: ['수량'], unitPriceCol: ['단가'] }, depositTypes: ['[입금] 예탁금이용료입금'], removalTypes: [] }
            ] },
            { id: 'hyundai', name: '현대차증권', aliases: ['현대차증권'], profiles: [
                { columns: { dateCol: ['일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금잔(원)'], feeCol: ['수수료'], taxCol: ['세금합'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['단가'], stockBalanceCol: ['유가금잔'] }, withdrawalTypes: ['매수'], removalTypes: [] },
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['거래구분'], amountCol: ['거래금액'], balanceCol: ['예수금잔'], feeCol: ['수수료'], taxCol: ['세금합'], stockNameCol: ['종목명'], stockQtyCol: ['거래좌수'], unitPriceCol: ['기준가'], stockBalanceCol: ['잔고좌수'] }, withdrawalTypes: ['매수'], removalTypes: [] }
            ] },
            { id: 'heungkuk', name: '흥국증권', aliases: ['흥국증권'], profiles: [] }
        ],
        foreignPresets: [
            { id: 'kyobo', name: '교보증권', aliases: ['교보증권'], profiles: [] },
            { id: 'daol', name: '다올투자증권', aliases: ['다올투자증권'], profiles: [] },
            { id: 'daishin', name: '대신증권', aliases: ['대신증권'], profiles: [] },
            { id: 'db', name: 'DB금융투자', aliases: ['DB금융투자', '디비금융투자'], profiles: [] },
            { id: 'ds', name: 'DS투자증권', aliases: ['DS투자증권', '디에스투자증권'], profiles: [] },
            { id: 'leading', name: '리딩투자증권', aliases: ['리딩투자증권'], profiles: [] },
            { id: 'meritz', name: '메리츠증권', aliases: ['메리츠증권'], profiles: [] },
            { id: 'mirae', name: '미래에셋증권', aliases: ['미래에셋증권'], profiles: [] },
            { id: 'bookook', name: '부국증권', aliases: ['부국증권'], profiles: [] },
            { id: 'bnk', name: 'BNK투자증권', aliases: ['BNK투자증권', '비엔케이투자증권'], profiles: [] },
            { id: 'samsung', name: '삼성증권', aliases: ['삼성증권'], profiles: [] },
            { id: 'sangsangin', name: '상상인증권', aliases: ['상상인증권'], profiles: [] },
            { id: 'shinyoung', name: '신영증권', aliases: ['신영증권'], profiles: [] },
            { id: 'shinhan', name: '신한투자증권', aliases: ['신한투자증권'], profiles: [
                { columns: { dateCol: ['일자'], tradeTypeCol: ['구분'], amountCol: ['거래대금'], balanceCol: ['최종금액'], feeCol: ['수수료'], taxCol: ['제세금'], stockNameCol: ['종목명(상대처)'], stockQtyCol: ['수량'], unitPriceCol: ['가격'] }, withdrawalTypes: ['환전출금'] }
            ] },
            { id: 'si', name: 'SI증권', aliases: ['SI증권', '에스아이증권'], profiles: [] },
            { id: 'sk', name: 'SK증권', aliases: ['SK증권', '에스케이증권'], profiles: [] },
            { id: 'nh', name: 'NH투자증권', aliases: ['NH투자증권', '엔에이치투자증권'], profiles: [] },
            { id: 'ls', name: 'LS증권', aliases: ['LS증권', '엘에스증권', '이베스트투자증권'], profiles: [] },
            { id: 'woori', name: '우리투자증권', aliases: ['우리투자증권', '한국포스증권', '우리종합금융'], profiles: [] },
            { id: 'yuanta', name: '유안타증권', aliases: ['유안타증권'], profiles: [] },
            { id: 'eugene', name: '유진투자증권', aliases: ['유진투자증권'], profiles: [] },
            { id: 'yuhwa', name: '유화증권', aliases: ['유화증권'], profiles: [] },
            { id: 'im', name: 'iM증권', aliases: ['iM증권', '아이엠증권', '하이투자증권'], profiles: [] },
            { id: 'ibk', name: 'IBK투자증권', aliases: ['IBK투자증권', '아이비케이투자증권'], profiles: [] },
            { id: 'kakaopay', name: '카카오페이증권', aliases: ['카카오페이증권'], profiles: [] },
            { id: 'kb', name: 'KB증권', aliases: ['KB증권', '케이비증권'], profiles: [] },
            { id: 'kr', name: 'KR투자증권', aliases: ['KR투자증권', '케이알투자증권'], profiles: [] },
            { id: 'cape', name: '케이프투자증권', aliases: ['케이프투자증권'], profiles: [] },
            { id: 'koreaasset', name: '코리아에셋투자증권', aliases: ['코리아에셋투자증권'], profiles: [] },
            { id: 'kiwoom', name: '키움증권', aliases: ['키움증권'], profiles: [
                { columns: { dateCol: ['거래일자'], tradeTypeCol: ['적요명'], amountCol: ['거래금액(외)'], balanceCol: ['외화예수금잔고'], feeCol: ['수수료(외)'], taxCol: ['외국납부세액'], stockNameCol: ['종목명'], stockQtyCol: ['거래수량'], unitPriceCol: ['거래단가/환율', '‘거래단가/환율'], stockBalanceCol: ['유가금잔'] }, depositTypes: ['매도', '배당금(외화)입금', '외화납부세환급(외화)입금', '외화예탁금이용료(이자)입금'], withdrawalTypes: ['매수', '외국납부세출금(외화)', '외화매도'] }
            ] },
            { id: 'toss', name: '토스증권', aliases: ['토스증권'], profiles: [] },
            { id: 'hana', name: '하나증권', aliases: ['하나증권'], profiles: [] },
            { id: 'sc', name: '한국스탠다드차타드증권', aliases: ['한국스탠다드차타드증권'], profiles: [] },
            { id: 'koreainvest', name: '한국투자증권', aliases: ['한국투자증권'], profiles: [] },
            { id: 'hanyang', name: '한양증권', aliases: ['한양증권'], profiles: [] },
            { id: 'hanhwa', name: '한화투자증권', aliases: ['한화투자증권'], profiles: [] },
            { id: 'hyundai', name: '현대차증권', aliases: ['현대차증권'], profiles: [] },
            { id: 'heungkuk', name: '흥국증권', aliases: ['흥국증권'], profiles: [] }
        ],
        normalizeText(value) {
            return String(value ?? '').trim().replace(/\u00a0/g, ' ').replace(/\s+/g, '');
        },
        normalizeTradeType(value) {
            return String(value ?? '').trim().replace(/\u00a0/g, ' ').replace(/\s+/g, '');
        },
        findHeaderIndex(headerRow, candidates = []) {
            const normalizedCandidates = candidates.map(candidate => this.normalizeText(candidate)).filter(Boolean);
            if (normalizedCandidates.length === 0) return -1;
            return (headerRow || []).findIndex(header => normalizedCandidates.includes(this.normalizeText(header)));
        },
        detect(data, sourceName = '', presets = this.presets) {
            const headerRow = data?.[0] || [];
            if (headerRow.length === 0) return { detected: false, name: '', brokerId: '', mappedColumns: {}, autoMappedKeys: [], removalTypes: [], depositTypes: [], withdrawalTypes: [], specialRules: null };
            const sourceNormalized = this.normalizeText(sourceName);
            const candidates = [];
            presets.forEach(broker => {
                (broker.profiles || []).forEach(profile => {
                    const mappedColumns = {};
                    let requiredCount = 0;
                    let score = 0;
                    Object.entries(profile.columns || {}).forEach(([key, headers]) => {
                        const headerIndex = this.findHeaderIndex(headerRow, headers);
                        if (headerIndex >= 0) {
                            mappedColumns[key] = headerIndex;
                            if (this.requiredKeys.includes(key)) {
                                requiredCount += 1;
                                score += 2;
                            } else {
                                score += 1;
                            }
                        }
                    });
                    const aliasMatched = (broker.aliases || []).some(alias => sourceNormalized.includes(this.normalizeText(alias)));
                    if (aliasMatched) score += 2.5;
                    candidates.push({ brokerId: broker.id, name: broker.name, mappedColumns, autoMappedKeys: Object.keys(mappedColumns), removalTypes: [...(profile.removalTypes || [])], depositTypes: [...(profile.depositTypes || [])], withdrawalTypes: [...(profile.withdrawalTypes || [])], specialRules: profile.specialRules || null, minMatchedKeys: profile.minMatchedKeys || 5, requiredCount, score, aliasMatched });
                });
            });
            candidates.sort((a, b) => {
                if (b.score !== a.score) return b.score - a.score;
                if (b.requiredCount !== a.requiredCount) return b.requiredCount - a.requiredCount;
                if (b.autoMappedKeys.length !== a.autoMappedKeys.length) return b.autoMappedKeys.length - a.autoMappedKeys.length;
                return (a.minMatchedKeys || 5) - (b.minMatchedKeys || 5);
            });
            const best = candidates[0];
            const second = candidates[1];
            const tied = second && best && second.score === best.score && second.requiredCount === best.requiredCount && second.autoMappedKeys.length === best.autoMappedKeys.length && second.brokerId !== best.brokerId;
            if (!best || best.requiredCount < 3 || best.autoMappedKeys.length < (best.minMatchedKeys || 5) || tied) return { detected: false, name: '', brokerId: '', mappedColumns: {}, autoMappedKeys: [], removalTypes: [], depositTypes: [], withdrawalTypes: [], specialRules: null };
            return { detected: true, name: best.name, brokerId: best.brokerId, mappedColumns: best.mappedColumns, autoMappedKeys: best.autoMappedKeys, removalTypes: best.removalTypes, depositTypes: best.depositTypes || [], withdrawalTypes: best.withdrawalTypes || [], specialRules: best.specialRules || null, score: best.score, requiredCount: best.requiredCount, aliasMatched: best.aliasMatched };
        },
        preprocessData(data, config = {}) {
            if (!Array.isArray(data) || data.length === 0) return data;
            if (config.detectedBrokerId === 'daishin') return this.preprocessDaishinData(data, config);
            if (config.specialRules?.deriveAmountFromTradeTypeText) return this.preprocessAmountFromTradeTypeText(data, config);
            return data.map(row => Array.isArray(row) ? [...row] : row);
        },
        toNumber(value) {
            if (value === null || value === undefined || String(value).trim() === '') return null;
            const num = parseFloat(String(value).replace(/,/g, ''));
            return Number.isFinite(num) ? num : null;
        },
                preprocessAmountFromTradeTypeText(data, config = {}) {
                    const [header = [], ...body] = data;
                    const tradeTypeIndex = config.tradeTypeCol;
                    const amountIndex = config.amountCol ?? header.length;
                    const depositAmountIndex = config.kbDepositAmountCol;
                    const withdrawalAmountIndex = config.kbWithdrawalAmountCol;
                    if (tradeTypeIndex === undefined) return data.map(row => Array.isArray(row) ? [...row] : row);

                    const keywordRules = config.specialRules?.classificationKeywords || {};
                    const depositKeywords = (keywordRules.deposit || []).map(value => this.normalizeTradeType(value)).filter(Boolean);
                    const withdrawalKeywords = (keywordRules.withdrawal || []).map(value => this.normalizeTradeType(value)).filter(Boolean);
            const removalTypes = (config.presetRemovalTypes || []).map(value => this.normalizeTradeType(value)).filter(Boolean);
            const hasClassificationKeyword = value => {
                const normalizedValue = this.normalizeTradeType(value);
                if (removalTypes.some(removalValue => normalizedValue === removalValue || normalizedValue.includes(removalValue))) return false;
                return depositKeywords.some(keyword => normalizedValue.includes(keyword)) ||
                    withdrawalKeywords.some(keyword => normalizedValue.includes(keyword));
            };
                    const getAmountFromColumn = (row, columnIndex) => {
                        if (columnIndex === undefined) return '';
                        const amount = this.toNumber(row?.[columnIndex]);
                        if (amount === null || Math.abs(amount) === 0) return '';
                        return Math.abs(amount);
                    };
                    const resolveTradeAmount = row => {
                        const depositAmount = getAmountFromColumn(row, depositAmountIndex);
                        if (depositAmount !== '') return depositAmount;
                        return getAmountFromColumn(row, withdrawalAmountIndex);
                    };

                    const clonedHeader = Array.isArray(header) ? [...header] : header;
                    clonedHeader[amountIndex] = clonedHeader[amountIndex] || '거래금액(KB)';
                    const nextBody = body.map(row => {
                        const nextRow = Array.isArray(row) ? [...row] : row;
                        const tradeType = nextRow?.[tradeTypeIndex];
                        nextRow[amountIndex] = hasClassificationKeyword(tradeType) ? resolveTradeAmount(nextRow) : '';
                        return nextRow;
                    });

                    return [clonedHeader, ...nextBody];
                },
        preprocessDaishinData(data, config = {}) {
            const [header = [], ...body] = data;
            const tradeTypeIndex = config.tradeTypeCol;
            const amountIndex = config.amountCol;
            if (tradeTypeIndex === undefined) return data.map(row => Array.isArray(row) ? [...row] : row);
            const rules = config.specialRules || {};
            const ignoredTradeTypes = (rules.ignoreTradeTypeContains || []).map(value => this.normalizeTradeType(value)).filter(Boolean);
            const amountOverrides = rules.amountOverrideByTradeType || [];
            const clonedHeader = Array.isArray(header) ? [...header] : header;
            const nextBody = body.reduce((rows, row) => {
                const nextRow = Array.isArray(row) ? [...row] : row;
                const tradeType = this.normalizeTradeType(nextRow?.[tradeTypeIndex]);
                if (tradeType && ignoredTradeTypes.some(keyword => tradeType.includes(keyword))) return rows;
                const matchedOverride = amountOverrides.find(rule => tradeType && tradeType.includes(this.normalizeTradeType(rule.match)));
                if (matchedOverride && amountIndex !== undefined) {
                    const overrideIndex = this.findHeaderIndex(clonedHeader, matchedOverride.sourceHeaders || []);
                    if (overrideIndex >= 0) {
                        const overrideValue = nextRow[overrideIndex];
                        if (overrideValue !== undefined && overrideValue !== null && String(overrideValue).trim() !== '') {
                            nextRow[amountIndex] = overrideValue;
                        }
                    }
                }
                rows.push(nextRow);
                return rows;
            }, []);
            return [clonedHeader, ...nextBody];
        },
        getPresetRemovalValues(data, config) {
            const tradeTypeIndex = config?.tradeTypeCol;
            const presetRemovalTypes = config?.presetRemovalTypes || [];
            if (tradeTypeIndex === undefined || presetRemovalTypes.length === 0) return [];
            const normalizedRemovalTypes = presetRemovalTypes.map(value => this.normalizeTradeType(value)).filter(Boolean);
            const actualTradeTypes = [...new Set((data || []).slice(1).map(row => row?.[tradeTypeIndex]).filter(Boolean).map(value => String(value).trim()))];
            return actualTradeTypes.filter(actualValue => {
                const normalizedActual = this.normalizeTradeType(actualValue);
                return normalizedRemovalTypes.some(removalValue => normalizedActual === removalValue || (config?.specialRules?.classificationKeywords && normalizedActual.includes(removalValue)));
            });
        },
        getPresetClassificationValues(data, config, presetKey) {
            const tradeTypeIndex = config?.tradeTypeCol;
            const presetTypes = config?.[presetKey] || [];
            if (tradeTypeIndex === undefined || presetTypes.length === 0) return [];
            const normalizedPresetTypes = presetTypes.map(value => this.normalizeTradeType(value)).filter(Boolean);
            const actualTradeTypes = [...new Set((data || []).slice(1).map(row => row?.[tradeTypeIndex]).filter(Boolean).map(value => String(value).trim()))];
            return actualTradeTypes.filter(actualValue => {
                const normalizedActual = this.normalizeTradeType(actualValue);
                return normalizedPresetTypes.some(presetValue => normalizedActual === presetValue);
            });
        },
        getKeywordClassificationValues(data, config = {}) {
            const tradeTypeIndex = config?.tradeTypeCol;
            const rules = config?.specialRules?.classificationKeywords;
            if (tradeTypeIndex === undefined || !rules) return { deposit: [], withdrawal: [], remove: [] };

            const depositKeywords = (rules.deposit || []).map(value => this.normalizeTradeType(value)).filter(Boolean);
            const withdrawalKeywords = (rules.withdrawal || []).map(value => this.normalizeTradeType(value)).filter(Boolean);
            const removeUnmatched = config?.specialRules?.unmatchedClassification === 'remove';
            const shouldRemoveAmountlessMatched = config?.specialRules?.deriveAmountFromTradeTypeText && config?.amountCol !== undefined;
            const actualTradeTypes = [...new Set((data || []).slice(1).map(row => row?.[tradeTypeIndex]).filter(Boolean).map(value => String(value).trim()))];
            const result = { deposit: [], withdrawal: [], remove: [] };
            const hasResolvedAmount = actualValue => {
                if (!shouldRemoveAmountlessMatched) return true;
                return (data || []).slice(1).some(row => {
                    if (String(row?.[tradeTypeIndex] ?? '').trim() !== actualValue) return false;
                    const amount = this.toNumber(row?.[config.amountCol]);
                    return amount !== null && Math.abs(amount) > 0;
                });
            };

            actualTradeTypes.forEach(actualValue => {
                const normalizedActual = this.normalizeTradeType(actualValue);
                const isDeposit = depositKeywords.some(keyword => normalizedActual.includes(keyword));
                const isWithdrawal = withdrawalKeywords.some(keyword => normalizedActual.includes(keyword));

                if ((isDeposit || isWithdrawal) && !hasResolvedAmount(actualValue)) result.remove.push(actualValue);
                else if (isDeposit && !isWithdrawal) result.deposit.push(actualValue);
                else if (isWithdrawal && !isDeposit) result.withdrawal.push(actualValue);
                else if (removeUnmatched) result.remove.push(actualValue);
            });

            return result;
        },
        buildPhase3Config(dataSet, sourceName = '') {
            const marketType = dataSet?.marketType || null;
            const presets = marketType === 'foreign' ? this.foreignPresets : this.presets;
            const detected = this.detect(dataSet?.data || dataSet, `${sourceName} ${dataSet?.name || ''}`.trim(), presets);
            const config = { ...detected.mappedColumns, detectedBrokerName: detected.name, detectedBrokerId: detected.brokerId, brokerDetected: detected.detected, autoMappedKeys: detected.autoMappedKeys || [], presetRemovalTypes: detected.removalTypes || [], presetDepositTypes: detected.depositTypes || [], presetWithdrawalTypes: detected.withdrawalTypes || [], specialRules: detected.specialRules || null, marketType, isComplete: false };
            if (config.specialRules?.deriveAmountFromTradeTypeText && config.amountCol === undefined) {
                const headerRow = (dataSet?.data || dataSet)?.[0] || [];
                const existingAmountIndex = this.findHeaderIndex(headerRow, ['거래금액(KB)', '거래금액(내용)']);
                config.amountCol = existingAmountIndex >= 0 ? existingAmountIndex : headerRow.length;
                config.autoMappedKeys = [...new Set([...(config.autoMappedKeys || []), 'amountCol'])];
            }
            return config;
        }
    };

    const buildAutoClassifyConfig = function(rawData, config = {}) {
        const sortValues = values => [...new Set((values || []).filter(Boolean))].sort();
        const presetRemovalValues = BrokerAutoMapper.getPresetRemovalValues(rawData, config);
        const presetDepositValues = BrokerAutoMapper.getPresetClassificationValues(rawData, config, 'presetDepositTypes');
        const presetWithdrawalValues = BrokerAutoMapper.getPresetClassificationValues(rawData, config, 'presetWithdrawalTypes');
        const keywordClassificationValues = BrokerAutoMapper.getKeywordClassificationValues(rawData, config);
        const removalSet = new Set(sortValues([...(config.valuesToRemove || []), ...presetRemovalValues, ...keywordClassificationValues.remove]));
        const depositSet = new Set(sortValues([...(config.depositTypes || []), ...presetDepositValues, ...keywordClassificationValues.deposit]));
        const withdrawalSet = new Set(sortValues([...(config.withdrawalTypes || []), ...presetWithdrawalValues, ...keywordClassificationValues.withdrawal]));
        removalSet.forEach(value => {
            depositSet.delete(value);
            withdrawalSet.delete(value);
        });
        depositSet.forEach(value => withdrawalSet.delete(value));
        return {
            ...config,
            valuesToRemove: sortValues([...removalSet]),
            depositTypes: sortValues([...depositSet]),
            withdrawalTypes: sortValues([...withdrawalSet])
        };
    };

    const api = { DataTransformer, BrokerAutoMapper, buildAutoClassifyConfig, sourceSha256: 'dda8762dc5d3a30d76b147b4a7f8d53fd1cabbf3f08fc182f9e4c95225efc236' };
    root.LegacySecurities = api;
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
