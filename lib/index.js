export const append = (sheetA, sheetB) => {
    var _a, _b, _c, _d, _e, _f, _g;
    const lastRowA = Number((_b = (_a = sheetA['!ref']) === null || _a === void 0 ? void 0 : _a.split(':')[1].match(/\d+/)) === null || _b === void 0 ? void 0 : _b[0]);
    const lastRowB = Number((_d = (_c = sheetB['!ref']) === null || _c === void 0 ? void 0 : _c.split(':')[1].match(/\d+/)) === null || _d === void 0 ? void 0 : _d[0]);
    const colsB = [];
    Object.entries(sheetB).forEach(([key]) => {
        const col = key.match(/^[A-Z]+/);
        if (col) {
            colsB.push(col[0]);
        }
    });
    for (let i = lastRowA + 1; i <= lastRowB + lastRowA; ++i) {
        colsB.forEach((col) => {
            // eslint-disable-next-line no-param-reassign
            sheetA[`${col}${i}`] = sheetB[`${col}${i - lastRowA}`];
        });
    }
    const rangeA = (_e = sheetA['!ref']) === null || _e === void 0 ? void 0 : _e.split(':')[0];
    const rangeB = `${(_g = (_f = sheetB['!ref']) === null || _f === void 0 ? void 0 : _f.split(':')[1].match(/^[A-Z]+/)) === null || _g === void 0 ? void 0 : _g[0]}${lastRowA + lastRowB}`;
    // eslint-disable-next-line no-param-reassign
    sheetA['!ref'] = `${rangeA}:${rangeB}`;
};
export const formulaCellByValue = (sheet, formula, valueToStyle) => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            if (value.v === valueToStyle) {
                // eslint-disable-next-line no-param-reassign
                value.f = formula;
            }
        }
    });
};
export const formulaCol = (sheet, formula, col) => {
    const cols = [];
    Object.entries(sheet).forEach(([key, value]) => {
        const keyCol = key.match(/^[A-Z]+/);
        if (keyCol) {
            cols.push(keyCol[0]);
            if (key !== '!ref' && keyCol) {
                if (keyCol[0] === cols[col]) {
                    // eslint-disable-next-line no-param-reassign
                    value.f = formula;
                }
            }
        }
    });
};
export const formulaRow = (sheet, formula, row) => {
    Object.entries(sheet).forEach(([key, value]) => {
        const keyRow = key.match(/\d+/);
        if (key !== '!ref' && keyRow) {
            if (keyRow[0] === String(row)) {
                // eslint-disable-next-line no-param-reassign
                value.f = formula;
            }
        }
    });
};
export const formulaAll = (sheet, format) => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            // eslint-disable-next-line no-param-reassign
            value.z = format;
        }
    });
};
export const formatCellByValue = (sheet, format, valueToStyle) => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            if (value.v === valueToStyle) {
                // eslint-disable-next-line no-param-reassign
                value.z = format;
            }
        }
    });
};
export const formatCol = (sheet, format, col) => {
    const cols = [];
    Object.entries(sheet).forEach(([key, value]) => {
        const keyCol = key.match(/^[A-Z]+/);
        if (keyCol) {
            cols.push(keyCol[0]);
            if (key !== '!ref' && keyCol) {
                if (keyCol[0] === cols[col]) {
                    // eslint-disable-next-line no-param-reassign
                    value.z = format;
                }
            }
        }
    });
};
export const formatRow = (sheet, format, row) => {
    Object.entries(sheet).forEach(([key, value]) => {
        const keyRow = key.match(/\d+/);
        if (key !== '!ref' && keyRow) {
            if (keyRow[0] === String(row)) {
                // eslint-disable-next-line no-param-reassign
                value.z = format;
            }
        }
    });
};
export const formatAll = (sheet, format) => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            // eslint-disable-next-line no-param-reassign
            value.z = format;
        }
    });
};
export const styleCellByValue = (sheet, style, valueToStyle) => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            if (value.v === valueToStyle) {
                // eslint-disable-next-line no-param-reassign
                value.s = Object.assign(Object.assign({}, value.s), style);
            }
        }
    });
};
export const styleCol = (sheet, style, col) => {
    Object.entries(sheet).forEach(([key, value]) => {
        const keyCol = key.match(/^[A-Z]+/);
        if (key !== '!ref' && keyCol) {
            if (keyCol[0] === col) {
                // eslint-disable-next-line no-param-reassign
                value.s = Object.assign(Object.assign({}, value.s), style);
            }
        }
    });
};
export const styleRow = (sheet, style, row) => {
    Object.entries(sheet).forEach(([key, value]) => {
        const keyRow = key.match(/\d+/);
        if (key !== '!ref' && keyRow) {
            if (keyRow[0] === String(row)) {
                // eslint-disable-next-line no-param-reassign
                value.s = Object.assign(Object.assign({}, value.s), style);
            }
        }
    });
};
export const styleAll = (sheet, style) => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            // eslint-disable-next-line no-param-reassign
            value.s = Object.assign(Object.assign({}, value.s), style);
        }
    });
};
export const removeCol = (sheet, col) => {
    // remove values matching col and shift all succeeding cols to the left
    // cols are identified by the alphabetical part of the cell key
    // examples: A1, B2, C3, AA1, AB2, AC3
    Object.entries(sheet).forEach(([key, value]) => {
        const keyCol = key.match(/^[A-Z]+/);
        if (key !== '!ref' && keyCol) {
            const colIndex = keyCol[0].charCodeAt(0) - 65;
            if (colIndex > col.charCodeAt(0) - 65) {
                // eslint-disable-next-line no-param-reassign
                sheet[`${String.fromCharCode(colIndex + 64)}${key.match(/\d+/)}`] = value;
                // eslint-disable-next-line no-param-reassign
                delete sheet[key];
            }
        }
    });
};
export const removeRow = (sheet, row) => {
    // remove values matching row and shift all succeeding rows up
    // rows are identified by the numerical part of the cell key
    // examples: A1, B2, C3, AA1, AB2, AC3
    Object.entries(sheet).forEach(([key, value]) => {
        const keyRow = key.match(/\d+/);
        if (key !== '!ref' && keyRow) {
            if (Number(keyRow[0]) > row) {
                // eslint-disable-next-line no-param-reassign
                sheet[`${key.match(/^[A-Z]+/)}${Number(keyRow[0]) - 1}`] =
                    value;
                // eslint-disable-next-line no-param-reassign
                delete sheet[key];
            }
        }
    });
};
export const insertRow = (sheet, rows, rowNumber) => {
    // insert rows (plural) at rowNumber and shift all succeeding rows down by the length of rows
    // rows are identified by the numerical part of the cell key
    // examples: A1, B2, C3, AA1, AB2, AC3
    const rowSize = Math.max(...Object.keys(rows).map((key) => {
        const keyRow = key.match(/\d+/);
        return keyRow ? Number(keyRow[0]) : 0;
    }));
    const tempSheet = {};
    // clone rows greater than or equal to rowNumber to tempSheet
    Object.entries(sheet).forEach(([key, value]) => {
        const keyRow = key.match(/\d+/);
        if (key !== '!ref' && keyRow) {
            if (Number(keyRow[0]) >= rowNumber) {
                // eslint-disable-next-line no-param-reassign
                // advance number portion of key by rowSize
                const newKey = `${key.match(/^[A-Z]+/)}${Number(keyRow[0]) + rowSize}`;
                console.log({ key, newKey });
                tempSheet[newKey] = value;
            }
        }
    });
    // copy tempSheet back to sheet
    Object.entries(tempSheet).forEach(([key, value]) => {
        // eslint-disable-next-line no-param-reassign
        sheet[key] = value;
    });
    // remove rows at rowNumber and up to rowNumber + rowSize
    Object.entries(sheet).forEach(([key]) => {
        const keyRow = key.match(/\d+/);
        if (key !== '!ref' && keyRow) {
            if (Number(keyRow[0]) >= rowNumber) {
                if (Number(keyRow[0]) < rowNumber + rowSize) {
                    // eslint-disable-next-line no-param-reassign
                    delete sheet[key];
                }
            }
        }
    });
    // copy rows to sheet at entry point rowNumber
    Object.entries(rows).forEach(([key, value]) => {
        const insertKey = `${key.match(/^[A-Z]+/)}${rowNumber + Number(key.match(/\d+/)) - 1}`;
        // eslint-disable-next-line no-param-reassign
        sheet[insertKey] = value;
    });
};
export const styleRoot = (ws) => {
    // eslint-disable-next-line no-param-reassign
    ws['!sheetFormat'] = {
        col: {
            width: 20,
        },
    };
};
