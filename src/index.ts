import SheetJS, { Style, WorkSheet } from '@sheet/core';

export const append = (
    sheetA: SheetJS.WorkSheet,
    sheetB: SheetJS.WorkSheet
): void => {
    const lastRowA = Number(sheetA['!ref']?.split(':')[1].match(/\d+/)?.[0]);
    const lastRowB = Number(sheetB['!ref']?.split(':')[1].match(/\d+/)?.[0]);

    const colsB: string[] = [];

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

    const rangeA = sheetA['!ref']?.split(':')[0];
    const rangeB = `${sheetB['!ref']?.split(':')[1].match(/^[A-Z]+/)?.[0]}${
        lastRowA + lastRowB
    }`;

    // eslint-disable-next-line no-param-reassign
    sheetA['!ref'] = `${rangeA}:${rangeB}`;
};

export const formulaCellByValue = (
    sheet: SheetJS.WorkSheet,
    formula: string,
    valueToStyle: string
): void => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            if (value.v === valueToStyle) {
                // eslint-disable-next-line no-param-reassign
                value.f = formula;
            }
        }
    });
};

export const formulaCol = (
    sheet: SheetJS.WorkSheet,
    formula: string,
    col: number
): void => {
    const cols: string[] = [];
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

export const formulaRow = (
    sheet: SheetJS.WorkSheet,
    formula: string,
    row: number
): void => {
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

export const formulaAll = (sheet: SheetJS.WorkSheet, format: string): void => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            // eslint-disable-next-line no-param-reassign
            value.z = format;
        }
    });
};

export const formatCellByValue = (
    sheet: SheetJS.WorkSheet,
    format: string,
    valueToStyle: string
): void => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            if (value.v === valueToStyle) {
                // eslint-disable-next-line no-param-reassign
                value.z = format;
            }
        }
    });
};

export const formatCol = (
    sheet: SheetJS.WorkSheet,
    format: string,
    col: number
): void => {
    const cols: string[] = [];
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

export const formatRow = (
    sheet: SheetJS.WorkSheet,
    format: string,
    row: number
): void => {
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

export const formatAll = (sheet: SheetJS.WorkSheet, format: string): void => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            // eslint-disable-next-line no-param-reassign
            value.z = format;
        }
    });
};

export const styleCellByValue = (
    sheet: SheetJS.WorkSheet,
    style: Style,
    valueToStyle: string
): void => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            if (value.v === valueToStyle) {
                // eslint-disable-next-line no-param-reassign
                value.s = {
                    ...value.s,
                    ...style,
                };
            }
        }
    });
};

export const styleCol = (
    sheet: SheetJS.WorkSheet,
    style: Style,
    col: string
): void => {
    Object.entries(sheet).forEach(([key, value]) => {
        const keyCol = key.match(/^[A-Z]+/);
        if (key !== '!ref' && keyCol) {
            if (keyCol[0] === col) {
                // eslint-disable-next-line no-param-reassign
                value.s = {
                    ...value.s,
                    ...style,
                };
            }
        }
    });
};

export const styleRow = (
    sheet: SheetJS.WorkSheet,
    style: Style,
    row: number
): void => {
    Object.entries(sheet).forEach(([key, value]) => {
        const keyRow = key.match(/\d+/);
        if (key !== '!ref' && keyRow) {
            if (keyRow[0] === String(row)) {
                // eslint-disable-next-line no-param-reassign
                value.s = {
                    ...value.s,
                    ...style,
                };
            }
        }
    });
};

export const styleAll = (sheet: SheetJS.WorkSheet, style: Style): void => {
    Object.entries(sheet).forEach(([key, value]) => {
        if (key !== '!ref') {
            // eslint-disable-next-line no-param-reassign
            value.s = {
                ...value.s,
                ...style,
            };
        }
    });
};

export const removeCol = (sheet: SheetJS.WorkSheet, col: string): void => {
    // remove values matching col and shift all succeeding cols to the left
    // cols are identified by the alphabetical part of the cell key
    // examples: A1, B2, C3, AA1, AB2, AC3
    Object.entries(sheet).forEach(([key, value]) => {
        const keyCol = key.match(/^[A-Z]+/);
        if (key !== '!ref' && keyCol) {
            const colIndex = keyCol[0].charCodeAt(0) - 65;
            if (colIndex > col.charCodeAt(0) - 65) {
                // eslint-disable-next-line no-param-reassign
                sheet[
                    `${String.fromCharCode(colIndex + 64)}${key.match(/\d+/)}`
                ] = value;
                // eslint-disable-next-line no-param-reassign
                delete sheet[key];
            }
        }
    });
};

export const removeRow = (sheet: SheetJS.WorkSheet, row: number): void => {
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

export const insertRow = (
    sheet: SheetJS.WorkSheet,
    rows: SheetJS.WorkSheet,
    rowNumber: number
): void => {
    // insert rows (plural) at rowNumber and shift all succeeding rows down by the length of rows
    // rows are identified by the numerical part of the cell key
    // examples: A1, B2, C3, AA1, AB2, AC3

    const rowSize = Math.max(
        ...Object.keys(rows).map((key) => {
            const keyRow = key.match(/\d+/);
            return keyRow ? Number(keyRow[0]) : 0;
        })
    );

    const tempSheet: SheetJS.WorkSheet = {};

    // clone rows greater than or equal to rowNumber to tempSheet
    Object.entries(sheet).forEach(([key, value]) => {
        const keyRow = key.match(/\d+/);
        if (key !== '!ref' && keyRow) {
            if (Number(keyRow[0]) >= rowNumber) {
                // eslint-disable-next-line no-param-reassign
                // advance number portion of key by rowSize
                const newKey = `${key.match(/^[A-Z]+/)}${
                    Number(keyRow[0]) + rowSize
                }`;
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
        const insertKey = `${key.match(/^[A-Z]+/)}${
            rowNumber + Number(key.match(/\d+/)) - 1
        }`;

        // eslint-disable-next-line no-param-reassign
        sheet[insertKey] = value;
    });
};

export const styleRoot = (ws: WorkSheet): void => {
    // eslint-disable-next-line no-param-reassign
    ws['!sheetFormat'] = {
        col: {
            width: 20,
        },
    };
};
