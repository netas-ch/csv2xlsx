/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-18
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
        number: '#,##0',
        float: '#,##0.0'
    };

    constructor(csvString, formatCodes) {
        this.#rows = Csv2Json(csvString, {separator: ';', returnArray: true});

        if (formatCodes) {
            for (const fc in formatCodes) {
                this.#defaultFormatCodes[fc] = formatCodes[fc];
            }
        }

        for (let i = 0; i < this.#getColumnCount(); i++) {
            const t = this.#getColumnType(i, 1);
            this.#columns.push({
                name: this.#rows && this.#rows[0] && this.#rows[0][i] ? this.#rows[0][i] : 'Column ' + (i+1),
                rowKey: i,
                type: t, // skip row 0 as it contains the header
                width: 0,
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

    #convertRowData() {
        for (let i = 1; i < this.#rows.length; i++) {
            for (let y = 0; y < this.#rows[i].length; y++) {
                const col = this.#columns[y];
                if (col.type === 'number') {
                    this.#rows[i][y] = parseInt(this.#rows[i][y]);

                } else if (col.type.substring(0,5) === 'float') {
                    this.#rows[i][y] = parseFloat(this.#rows[i][y]);

                } else if (col.type === 'date') {
                    this.#rows[i][y] = Utils.parseDate(this.#rows[i][y]);
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

