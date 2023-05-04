/*
    MIT License

    Copyright (c) 2023 Lukas Buchs, netas.ch

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
 */

import { Csv2Json } from './Csv2Json.js';
import { Utils } from '../util/Utils.js';

export class CsvProcessing {
    #rows;
    #columnTypes = [];
    #columns = [];
    #defaultFormatCodes = {
        text: '@',
        date: 'dd.mm.yyyy',
        datetime: 'dd.mm.yyyy h:mm',
        number: '0',
        float: '0.0'
    };
    #defaultColumnWidths = {
        text: 0,
        date: 12,
        datetime: 16,
        number: 0,
        float: 0
    };

    constructor(csvString, csvSeparator, formatCodes) {
        this.#rows = Csv2Json(csvString, {separator: csvSeparator ? csvSeparator : '', returnArray: true});

        if (formatCodes) {
            for (const fc in formatCodes) {
                this.#defaultFormatCodes[fc] = formatCodes[fc];
            }
        }

        for (let i = 0; i < this.#getColumnCount(); i++) {
            const t = this.#getColumnType(i, 1);

            // column width
            let cW = this.#defaultColumnWidths[t] ? this.#defaultColumnWidths[t] : 0;
            if (t === 'text') {
                cW = this.#getColumnWidth(i, 1);
            }

            this.#columns.push({
                name: this.#rows && this.#rows[0] && this.#rows[0][i] ? this.#rows[0][i] : 'Column ' + Utils.numericToAlphaColumn(i+1),
                rowKey: i,
                type: t, // skip row 0 as it contains the header
                width: cW,
                columnType: this.#getColumnTypeObj(t)
            });
        }

        // convert
        this.#convertRowData();
    }

    getCsvData() {
        return {
            columnTypes: this.#columnTypes,
            columns: this.#columns,
            rows: this.#rows
        };
    }


    // PRIVATE


    /**
     * gets the numer of columns
     * @returns {Number}
     */
    #getColumnCount() {
        let cnt = 0;
        for (let i = 0; i < this.#rows.length; i++) {
            cnt = Math.max(cnt, this.#rows[i].length);
        }
        return cnt;
    }

    /**
     * get the type of a column.
     *
     * @param {Number} columnIndex
     * @param {Number} from
     * @returns {String} text | date | datetime | integer | float1 | float2 | float[...x]
     */
    #getColumnType(columnIndex, from=0) {
        let type = '', floatLen=0;

        for (let i = from; i < this.#rows.length; i++) {
            const v = this.#rows[i][columnIndex];

            if (typeof v !== 'undefined' && v !== '') {
                let t = 'text';

                if (Utils.stringIsInteger(v)) {
                    t = 'number';

                } else if (Utils.stringIsFloat(v)) {
                    t = 'float';

                    const mt = v.match(/(?<=\.)[0-9]+$/);
                    if (mt) {
                        floatLen = Math.max(floatLen, mt[0].length);
                    }

                } else if (Utils.stringIsDateTime(v)) {
                    t = 'datetime';

                } else if (Utils.stringIsDate(v)) {
                    t = 'date';
                }

                if ((type === 'float' && t === 'number') || (type === 'number' && t === 'float')) {
                    type = 'float';

                // different types? use text
                } else if (type !== '' && type !== t) {
                    return 'text';

                // letters? then its text.
                } else if (t === 'text' && t.match(/[a-zA-ZöüäÖÜÄéèàÉÈÀÂâ]{2}/)) {
                    return 'text';

                // check next row
                } else {
                    type = t;
                }
            }
        }

        if (type === 'float') {
            return 'float' + floatLen;
        }

        if (type !== '') {
            return type;
        }

        return 'text';
    }

    #getColumnWidth(columnIndex, from=0, maxWidth=120) {
        let width = 10, floatLen=0;

        for (let i = from; i < this.#rows.length; i++) {
            const v = this.#rows[i][columnIndex];

            if (v && typeof v === 'string') {
                let wx = Utils.excelStringWidth(v);
                width = Math.max(width, wx);

                if (width > maxWidth) {
                    return maxWidth;
                }
            }
        }
        return Math.ceil(width);
    }

    #convertRowData() {
        for (let i = 0; i < this.#rows.length; i++) {
            for (let y = 0; y < this.#columns.length; y++) {
                const col = this.#columns[y];
                if (typeof this.#rows[i][y] === 'undefined' || this.#rows[i][y] === '') {
                    this.#rows[i][y] = i===0 ? col.name : null;

                } else if (i>0) {
                    if (col.type === 'number') {
                        this.#rows[i][y] = parseInt(this.#rows[i][y]);

                    } else if (col.type.substring(0,5) === 'float') {
                        this.#rows[i][y] = parseFloat(this.#rows[i][y]);

                    } else if (col.type === 'date' || col.type === 'datetime') {
                        this.#rows[i][y] = Utils.parseDate(this.#rows[i][y]);
                    }
                }
            }
        }
    }

    #getColumnTypeObj(typeStr, formatCode='') {
        if (typeStr === 'text') {
             return null;
        }

        if (!formatCode) {
            if (this.#defaultFormatCodes[typeStr]) {
                formatCode = this.#defaultFormatCodes[typeStr];
            } else if (typeStr.substring(0,5) === 'float') {
                formatCode = this.#defaultFormatCodes.float;
                let floatLen = parseInt(typeStr.substring(5));
                formatCode += ''.padEnd(floatLen-1, '0');
            }
        }

        // search for existing
        for (const columnType of this.#columnTypes) {
            if (columnType.type === typeStr && columnType.formatCode === formatCode) {
                return columnType;
            }
        }

        // create a new one
        const nt = {
            type: typeStr,
            formatCode: formatCode,
            applyNumberFormat: 1,
            numFmtId: null
        };
        this.#columnTypes.push(nt);

        return nt;
    }



}

