// DOM Elements
const sheetsContainer = document.getElementById('sheets-container');
const addSheetBtn = document.getElementById('add-sheet-btn');
const downloadBtn = document.getElementById('download-btn');
let elementCounter = 0; // Universal counter for unique IDs

// --- TEMPLATE DATA ---
const SSA_TEMPLATE = {
    sheetName: 'SS-A',
    tables: [
        {
            tableName: 'Land Cost',
            columns: [
                { name: 'S.No' },
                { name: 'Particulars' },
                { name: 'Sq-Ft' },
                { name: 'Rs Lakhs/Sq-Ft' },
                { name: 'Rs Lakhs', formula: '[Sq-Ft] * [Rs Lakhs/Sq-Ft]' }
            ]
        }
    ]
};

const SSB_TEMPLATE = {
    sheetName: 'SS-B',
    tables: [
        {
            tableName: 'Cost Of Building',
            columns: [
                { name: 'S.No' },
                { name: 'Particulars' },
                { name: 'Sq-Ft' },
                { name: 'Rs Lakhs/Sq-Ft' },
                { name: 'Rs Lakhs', formula: '[Sq-Ft] * [Rs Lakhs/Sq-Ft]' }
            ]
        }
    ]
};

const SSC_TEMPLATE = {
    sheetName: 'SS-C',
    tables: [
        {
            tableName: 'Plant & Equipment',
            columns: [
                { name: 'S.No' },
                { name: 'Particulars' },
                { name: 'Rs Lakhs/Nos' }
            ]
        }
    ]
};

const SSC1_TEMPLATE = {
    sheetName: 'SS-C1',
    tables: [
        {
            tableName: 'Furniture & Fittings',
            columns: [
                { name: 'S.No' },
                { name: 'Particulars' },
                { name: 'Rs Lakhs/Nos' }
            ]
        }
    ]
};

const SSC2_TEMPLATE = {
    sheetName: 'SS-C2',
    tables: [
        {
            tableName: 'Electrical Fittings',
            columns: [
                { name: 'S.No' },
                { name: 'Particulars' },
                { name: 'Rs Lakhs/Nos' }
            ]
        },
        {
            tableName: 'Other Fixed Assets',
            columns: [
                { name: 'S.No' },
                { name: 'Particulars' },
                { name: 'Rs Lakhs/Nos' }
            ]
        }
    ]
};

const SSD_TEMPLATE = {
    sheetName: 'SS-D',
    tables: [
        {
            tableName: 'Pre-Operative Expenses',
            columns: [
                { name: 'S.No' },
                { name: 'Particulars' },
                { name: 'Rs Lakhs/Nos' }
            ]
        }
    ]
};

const SSD1_TEMPLATE = {
    sheetName: 'SS-D1',
    tables: [
        {
            tableName: 'Interest During Construction Period',
            columns: [
                { name: 'Month' },
                { name: '% of Loan withdrawn during construction period' },
                { name: 'Loan amount withdrawn' },
                { name: 'Cumulative Amount Outstanding' },
                { name: 'Interest' }
            ]
        }
    ]
};

// --- Core UI Functions (Refactored for Reliability) ---

const addSheet = (sheetName = `Sheet ${++elementCounter}`) => {
    const sheetId = `sheet-${elementCounter}`;
    const sheetHTML = `
        <div id="${sheetId}" class="p-4 border-2 border-indigo-200 rounded-lg sheet-block">
            <div class="flex justify-between items-center mb-4">
                <div class="flex items-center gap-4 w-1/2">
                     <label class="text-xl font-bold text-indigo-800">Sheet Name:</label>
                     <input type="text" value="${sheetName}" class="sheet-name-input flex-grow px-3 py-2 bg-white border border-slate-300 rounded-md shadow-sm">
                </div>
                <button onclick="removeElement('${sheetId}')" class="bg-red-500 text-white font-bold py-1 px-3 rounded-lg hover:bg-red-600">&times; Remove Sheet</button>
            </div>
            <div class="tables-area space-y-6 pl-6 border-l-4 border-slate-200"></div>
            <button onclick="orchestrateAddTable(this)" class="mt-4 ml-6 bg-sky-600 text-white font-semibold py-1 px-3 rounded-lg hover:bg-sky-700 text-sm">+ Add Table to this Sheet</button>
        </div>`;
    sheetsContainer.insertAdjacentHTML('beforeend', sheetHTML);
    return document.getElementById(sheetId);
};

const addTable = (sheetElement, tableName = 'New Table') => {
    const tablesArea = sheetElement.querySelector('.tables-area');
    const tableId = `table-${++elementCounter}`;
    
    const tableHTML = `
        <div id="${tableId}" class="p-4 border-2 border-slate-300 rounded-lg table-block">
            <div class="flex justify-between items-center mb-4">
                <div class="flex items-center gap-3 w-2/3">
                    <label class="font-bold text-slate-800 text-lg">Table Name:</label>
                    <input type="text" value="${tableName}" class="table-name-input flex-grow px-2 py-1 bg-white border border-slate-300 rounded-md shadow-sm">
                </div>
                <button onclick="removeElement('${tableId}')" class="bg-red-500 text-white font-bold py-1 px-2 rounded-lg hover:bg-red-600 text-sm">&times; Remove Table</button>
            </div>
            <div>
                <div class="flex justify-between items-center mb-1">
                    <label class="font-semibold text-gray-700">Column Names:</label>
                    <button onclick="addColumn(this)" class="bg-gray-600 text-white font-semibold py-1 px-3 rounded-lg hover:bg-gray-700 text-xs">+ Add Column</button>
                </div>
                <div class="column-headers-container flex flex-col sm:flex-row gap-2 mt-1"></div>
            </div>
            <div class="mt-2 overflow-x-auto">
                <table class="w-full border-collapse text-sm">
                    <tbody></tbody>
                    <tfoot class="bg-slate-200 font-bold"></tfoot>
                </table>
            </div>
            <button onclick="orchestrateAddRow(this)" class="mt-3 bg-blue-600 text-white font-semibold py-1 px-3 rounded-lg hover:bg-blue-700 text-sm">+ Add Row</button>
        </div>`;

    tablesArea.insertAdjacentHTML('beforeend', tableHTML);
    return document.getElementById(tableId);
};

const addRow = (tableBody, isRemovable = true) => {
    const tableBlock = tableBody.closest('.table-block');
    const numCols = tableBlock.querySelectorAll('.column-name-input').length;
    const rowId = `row-${++elementCounter}`;
    let cells = '';

    for (let i = 0; i < numCols; i++) {
        cells += `<td class="p-1"><input type="text" class="data-input w-full px-2 py-1 border border-slate-300 rounded-md"></td>`;
    }

    const removeButtonHTML = isRemovable ? `<button onclick="removeElement('${rowId}')" class="text-red-500 hover:text-red-700 font-bold px-2 rounded-full">&ndash;</button>` : ``;
    const rowHTML = `<tr id="${rowId}" class="bg-white">${cells}<td class="p-1 text-center w-12">${removeButtonHTML}</td></tr>`;
    tableBody.insertAdjacentHTML('beforeend', rowHTML);
    
    const newRow = document.getElementById(rowId);
    newRow.querySelectorAll('.data-input').forEach(input => {
        input.addEventListener('input', () => {
            updateCalculatedColumns(newRow);
            updateTableTotals(tableBlock);
        });
    });
    
    updateSerialNumbers(tableBlock);
    updateTableTotals(tableBlock);
};

const addColumn = (button, colName = '', formula = '') => {
    const tableBlock = button.closest('.table-block');
    const headersContainer = tableBlock.querySelector('.column-headers-container');
    const placeholder = colName || `Column ${headersContainer.children.length + 1}`;

    const headerHTML = `
        <div class="flex-1 flex items-center gap-1 p-1 bg-slate-100 rounded-t-md border border-b-0">
            <input type="text" placeholder="${placeholder}" value="${colName}" class="column-name-input w-full px-2 py-1 border rounded" data-formula="${formula}">
            <button onclick="setFormula(this)" class="text-blue-600 font-bold hover:text-blue-800" title="Set Formula">(fx)</button>
            <button onclick="removeColumn(this)" class="text-red-500 font-bold hover:text-red-700 px-1" title="Remove Column">(x)</button>
        </div>`;
    headersContainer.insertAdjacentHTML('beforeend', headerHTML);

    const newHeaderInput = headersContainer.lastElementChild.querySelector('.column-name-input');
    newHeaderInput.addEventListener('change', () => updateSerialNumbers(tableBlock));

    tableBlock.querySelectorAll('tbody tr').forEach(row => {
        row.querySelector('td:last-child').insertAdjacentHTML('beforebegin', `<td class="p-1"><input type="text" class="data-input w-full px-2 py-1 border border-slate-300 rounded-md"></td>`);
    });
    
    if (formula) {
        const columnIndex = Array.from(headersContainer.children).length - 1;
        tableBlock.querySelectorAll('tbody tr').forEach(row => {
            const cellInput = row.querySelectorAll('.data-input')[columnIndex];
            if (cellInput) {
                cellInput.readOnly = true;
                cellInput.classList.add('bg-slate-100');
            }
        });
    }

    updateSerialNumbers(tableBlock);
    updateTableTotals(tableBlock);
};

const removeColumn = (button) => {
    const headerDiv = button.closest('.flex-1');
    const headersContainer = headerDiv.parentElement;
    const tableBlock = headersContainer.closest('.table-block');

    if (headersContainer.children.length <= 1) {
        alert('Error: Each table must have at least one column.');
        return;
    }

    const colIndex = Array.from(headersContainer.children).indexOf(headerDiv);
    headerDiv.remove();

    tableBlock.querySelectorAll('tbody tr, tfoot tr').forEach(row => {
        const cellToRemove = row.children[colIndex];
        if(cellToRemove) cellToRemove.remove();
    });

    updateSerialNumbers(tableBlock);
    updateTableTotals(tableBlock);
};

const orchestrateAddTable = (button) => {
    const sheetElement = button.closest('.sheet-block');
    const tableElement = addTable(sheetElement, 'New Table');
    const addColBtn = tableElement.querySelector('button[onclick^="addColumn"]');
    addColumn(addColBtn, 'Column 1');
    const tableBody = tableElement.querySelector('tbody');
    addRow(tableBody);
};

const orchestrateAddRow = (button) => {
    const tableBody = button.closest('.table-block').querySelector('tbody');
    addRow(tableBody);
};

window.removeElement = (elementId) => {
    const element = document.getElementById(elementId);
    if (!element) return;
    
    const tableBlock = element.closest('.table-block');

    if (element.classList.contains('sheet-block')) {
        if (document.querySelectorAll('.sheet-block').length <= 1) {
            alert('Error: You must have at least one sheet.');
            return;
        }
    }
    if (element.classList.contains('table-block')) {
        const parentSheet = element.closest('.sheet-block');
        if (parentSheet && parentSheet.querySelectorAll('.table-block').length <= 1) {
            alert('Error: Each sheet must have at least one table.');
            return;
        }
    }
    
    const isRow = element.tagName === 'TR';
    element.remove();
    
    if (tableBlock) {
         if (isRow) {
            updateSerialNumbers(tableBlock);
         }
         updateTableTotals(tableBlock);
    }
};

const updateSerialNumbers = (tableBlock) => {
    const headers = Array.from(tableBlock.querySelectorAll('.column-name-input'));
    const snoIndex = headers.findIndex(h => (h.value || h.placeholder) === 'S.No');

    headers.forEach((h, i) => {
        const isSnoCol = (h.value || h.placeholder) === 'S.No';
        tableBlock.querySelectorAll('tbody tr').forEach(row => {
            const cellInput = row.querySelectorAll('.data-input')[i];
            if (cellInput) {
                cellInput.readOnly = isSnoCol;
                cellInput.classList.toggle('bg-slate-200', isSnoCol);
                cellInput.classList.toggle('text-gray-600', isSnoCol);
            }
        });
    });

    if (snoIndex !== -1) {
        const rows = tableBlock.querySelectorAll('tbody tr');
        rows.forEach((row, rowIndex) => {
            const cellInput = row.querySelectorAll('.data-input')[snoIndex];
            if (cellInput) {
                cellInput.value = rowIndex + 1;
            }
        });
    }
};

window.setFormula = (button) => {
    const headerInput = button.parentElement.querySelector('.column-name-input');
    const currentFormula = headerInput.dataset.formula || '';
    const newFormula = prompt('Enter formula (e.g., [Sq-Ft] * [Price/Sq-Ft]).\nLeave blank to clear.', currentFormula);

    if (newFormula !== null) {
        headerInput.dataset.formula = newFormula;
        const tableBlock = headerInput.closest('.table-block');
        tableBlock.querySelectorAll('tbody tr').forEach(row => {
            updateCalculatedColumns(row);
        });
    }
};

const updateCalculatedColumns = (rowElement) => {
    const tableBlock = rowElement.closest('.table-block');
    const headers = Array.from(tableBlock.querySelectorAll('.column-name-input'));
    const rowInputs = Array.from(rowElement.querySelectorAll('.data-input'));

    headers.forEach((header, index) => {
        const formula = header.dataset.formula;
        if (formula) {
            try {
                let expression = formula;
                headers.forEach((h, i) => {
                    const colName = h.value || h.placeholder;
                    const value = parseFloat(rowInputs[i].value) || 0;
                    expression = expression.replace(new RegExp(`\\[${escapeRegExp(colName)}\\]`, 'g'), value);
                });
                
                const result = Function(`'use strict'; return (${expression})`)();
                rowInputs[index].value = isFinite(result) ? result.toFixed(2) : '';
            } catch (e) {
                rowInputs[index].value = '#ERROR!';
            }
        }
    });
    updateTableTotals(tableBlock);
};

const updateTableTotals = (tableBlock) => {
    const tfoot = tableBlock.querySelector('tfoot');
    const numCols = tableBlock.querySelectorAll('.column-name-input').length;
    if (numCols === 0) {
        tfoot.innerHTML = '';
        return;
    }
    const allRows = tableBlock.querySelectorAll('tbody tr');
    let totals = Array(numCols).fill(0);
    
    let isNumeric = Array(numCols).fill(true);
    if (allRows.length === 0) {
         isNumeric.fill(false);
    } else {
        for (let i = 0; i < numCols; i++) {
            for (const row of allRows) {
                const input = row.querySelectorAll('.data-input')[i];
                const value = input.value.trim();
                if (value === '' || isNaN(value)) {
                    isNumeric[i] = false;
                    break;
                }
            }
        }
    }

    allRows.forEach(row => {
        row.querySelectorAll('.data-input').forEach((input, colIndex) => {
            if (isNumeric[colIndex]) {
                totals[colIndex] += parseFloat(input.value) || 0;
            }
        });
    });

    let footerHTML = `<tr>`;
    for (let i = 0; i < numCols; i++) {
        let content = '';
        if (i === 0) {
            content = 'Total';
        } else if (isNumeric[i]) {
            content = totals[i].toFixed(2);
        }
        footerHTML += `<td class="p-2 ${isNumeric[i] ? 'total-value-cell' : ''}">${content}</td>`;
    }
    footerHTML += `<td class="p-2"></td></tr>`;
    tfoot.innerHTML = footerHTML;
};

const escapeRegExp = (string) => {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

const loadInitialTemplates = () => {
    sheetsContainer.innerHTML = '';
    elementCounter = 0;
    
    [SSA_TEMPLATE, SSB_TEMPLATE, SSC_TEMPLATE, SSC1_TEMPLATE, SSC2_TEMPLATE, SSD_TEMPLATE, SSD1_TEMPLATE].forEach(template => {
        const sheetElement = addSheet(template.sheetName);
        template.tables.forEach(tableData => {
            const tableElement = addTable(sheetElement, tableData.tableName);
            const addColBtn = tableElement.querySelector('button[onclick^="addColumn"]');
            tableData.columns.forEach(colData => {
                addColumn(addColBtn, colData.name, colData.formula);
            });
            const tableBody = tableElement.querySelector('tbody');
            addRow(tableBody, false);
            updateTableTotals(tableElement);
        });
    });
};

// Event Listeners
addSheetBtn.addEventListener('click', () => {
    const sheetElement = addSheet();
    orchestrateAddTable(sheetElement.querySelector('button[onclick^="orchestrateAddTable"]'));
});

downloadBtn.addEventListener('click', async () => {
    downloadBtn.textContent = 'Generating...';
    downloadBtn.disabled = true;
    try {
        const workbook = new ExcelJS.Workbook();
        const createSummary = document.getElementById('summary-page-checkbox').checked;
        const summaryData = [];
        let summarySheet = null;

        if (createSummary) {
            summarySheet = workbook.addWorksheet('Summary');
        }

        document.querySelectorAll('.sheet-block').forEach(sheetBlock => {
            const sheetName = sheetBlock.querySelector('.sheet-name-input').value || 'Untitled Sheet';
            const worksheet = workbook.addWorksheet(sheetName);
            let currentRowNum = 1;

            sheetBlock.querySelectorAll('.table-block').forEach((tableBlock, tableIndex) => {
                const tableName = tableBlock.querySelector('.table-name-input').value || `Table ${tableIndex + 1}`;
                const headerInputs = tableBlock.querySelectorAll('.column-name-input');
                const headers = Array.from(headerInputs).map(input => ({ name: input.value || input.placeholder, formula: input.dataset.formula }));
                
                const titleCell = worksheet.getCell(`A${currentRowNum}`);
                titleCell.value = tableName;
                titleCell.font = { name: 'Calibri', size: 16, bold: true };
                currentRowNum += 2;
                
                const tableStartRow = currentRowNum;
                const rowElements = tableBlock.querySelectorAll('tbody > tr');
                const headerNames = headers.map(h => h.name);

                const dataRows = Array.from(rowElements).map(rowEl => {
                    return Array.from(rowEl.querySelectorAll('.data-input')).map(input => {
                        const val = input.value;
                        // For calculated columns, ensure we get the calculated value
                        if (input.readOnly && input.classList.contains('bg-slate-100')) {
                            return isNaN(val) || val.trim() === '' ? 0 : Number(val);
                        }
                        return isNaN(val) || val.trim() === '' ? val : Number(val);
                    });
                });

                const numCols = headers.length;
                let isNumeric = Array(numCols).fill(true);
                 dataRows.forEach(row => {
                    row.forEach((cell, colIndex) => {
                        if (typeof cell !== 'number') {
                            isNumeric[colIndex] = false;
                        }
                    });
                });

                let lastNumericIndex = -1;
                for (let i = isNumeric.length - 1; i >= 0; i--) { if (isNumeric[i]) { lastNumericIndex = i; break; } }
                
                const excelColumns = headers.map((h, index) => {
                    const colDef = { name: h.name, filterButton: true };
                    if (index === lastNumericIndex && index > 0) {
                        colDef.totalsRowFunction = 'sum';
                    }
                    return colDef;
                });

                // Add headers first
                const headerRow = headers.map(h => h.name);
                const headerRowObj = worksheet.getRow(currentRowNum);
                headerRowObj.values = headerRow;
                
                // Style the header row with blue background and white text
                headerRowObj.eachCell((cell, colNumber) => {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF4472C4' } // Blue color
                    };
                    cell.font = {
                        bold: true,
                        color: { argb: 'FFFFFFFF' } // White text
                    };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FF000000' } },
                        left: { style: 'thin', color: { argb: 'FF000000' } },
                        bottom: { style: 'thin', color: { argb: 'FF000000' } },
                        right: { style: 'thin', color: { argb: 'FF000000' } }
                    };
                });
                currentRowNum++;
                
                // Add data rows
                dataRows.forEach((rowData, rowIndex) => {
                    const excelRow = [];
                    headers.forEach((header, colIndex) => {
                        excelRow.push(rowData[colIndex]);
                    });
                    const dataRowObj = worksheet.getRow(currentRowNum);
                    dataRowObj.values = excelRow;
                    
                    // Style data rows with light background
                    dataRowObj.eachCell((cell, colNumber) => {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFF2F2F2' } // Light gray background
                        };
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'FF000000' } },
                            left: { style: 'thin', color: { argb: 'FF000000' } },
                            bottom: { style: 'thin', color: { argb: 'FF000000' } },
                            right: { style: 'thin', color: { argb: 'FF000000' } }
                        };
                    });
                    currentRowNum++;
                });
                
                // Add totals row
                const totalsRow = [];
                headers.forEach((header, colIndex) => {
                    if (colIndex === 0) {
                        totalsRow.push('Total');
                    } else if (isNumeric[colIndex]) {
                        const total = dataRows.reduce((sum, row) => sum + (parseFloat(row[colIndex]) || 0), 0);
                        totalsRow.push(total.toFixed(2));
                    } else {
                        totalsRow.push('');
                    }
                });
                const totalsRowObj = worksheet.getRow(currentRowNum);
                totalsRowObj.values = totalsRow;
                
                // Style the totals row with blue background and white text
                totalsRowObj.eachCell((cell, colNumber) => {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF4472C4' } // Blue color
                    };
                    cell.font = {
                        bold: true,
                        color: { argb: 'FFFFFFFF' } // White text
                    };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FF000000' } },
                        left: { style: 'thin', color: { argb: 'FF000000' } },
                        bottom: { style: 'thin', color: { argb: 'FF000000' } },
                        right: { style: 'thin', color: { argb: 'FF000000' } }
                    };
                });
                currentRowNum++;
                
                if (createSummary) {
                    // Find the last numeric column's total
                    const totalCells = tableBlock.querySelectorAll('.total-value-cell');
                    let totalValue = 0;
                    if (totalCells.length > 0) {
                        const lastTotalCell = totalCells[totalCells.length - 1];
                        totalValue = parseFloat(lastTotalCell.textContent) || 0;
                    }
                    const sheetName = sheetBlock.querySelector('.sheet-name-input').value || 'Untitled Sheet';
                    summaryData.push([`${sheetName} - ${tableName}`, totalValue]);
                }

                worksheet.columns.forEach(column => { column.width = 20; });
                currentRowNum += rowElements.length + 4;
            });
        });

        if (createSummary && summarySheet) {
            summarySheet.addTable({
                name: 'OverallSummary', ref: 'A1', headerRow: true,
                style: { theme: 'TableStyleMedium2', showRowStripes: true },
                columns: [
                    { name: 'Table Name', filterButton: true },
                    { name: 'Total Value', filterButton: true, totalsRowFunction: 'sum' }
                ],
                rows: summaryData
            });

            summarySheet.getColumn(2).numFmt = '#,##0.00';
            summarySheet.getColumn(1).width = 35;
            summarySheet.getColumn(2).width = 20;
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const a = document.createElement('a');
        a.href = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        a.download = 'my-data-file.xlsx';
        a.click();
        URL.revokeObjectURL(a.href);
    } catch (error) {
        console.error('Failed to generate Excel file:', error);
        alert('An error occurred. Check console for details.');
    } finally {
        downloadBtn.textContent = '💾 Download Excel File';
        downloadBtn.disabled = false;
    }
});

// --- INITIAL LOAD ---
loadInitialTemplates(); 