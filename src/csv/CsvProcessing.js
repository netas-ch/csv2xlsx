/*
        Copyright (c) 2024 Lukas Buchs, netas.ch

    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.
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
        datetime: 'dd.mm.yyyy hh:mm',
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
                cW = this.#getColumnWidth(i);
            }

            this.#columns.push({
                name: this.#getValidColumnName(this.#rows && this.#rows[0] && this.#rows[0][i] ? this.#rows[0][i] : null, i+1),
                rowKey: i,
                type: t, // skip row 0 as it contains the header
                width: cW,
                columnType: this.#getColumnTypeObj(t)
            });
        }

        // make column names unique
        this.#uniqueColumnNames();

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

    #getValidColumnName(str, colNr) {
        if (str) {

            // line breaks are not allowed in column headers
            str = str.replace(/(\w)\-[\r\n]+(\w)/g, '$1$2').replace(/[\r\n\t]/g, ' ').trim();
            if (str) {
                return str;
            }
        }

        // return default
        return 'Column ' + Utils.numericToAlphaColumn(colNr);
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

                // Compare with other rows
                // -----------------------

                if ((type === 'float' && t === 'number') || (type === 'number' && t === 'float')) {
                    type = 'float';

                // different types? use text
                } else if (type !== '' && type !== t) {
                    return 'text';

                // letters? then its text.
                } else if (t === 'text' && v.match(/[a-zA-ZöüäÖÜÄéèàÉÈÀÂâ]{2}/)) {
                    return 'text';

                // Excel will automatically convert numbers to Scientific Notation if longer than 15 digits.
                // If you need to enter long numeric strings, but don't want them converted, then format the cells in question as Text.
                } else if (t === 'number' && v.length > 15) {
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
                let wx = Utils.excelStringWidth(v, i===0);
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

                // first row contains column names
                if (i === 0) {
                    this.#rows[i][y] = col.name;

                } else if (typeof this.#rows[i][y] === 'undefined' || this.#rows[i][y] === '') {
                    this.#rows[i][y] = null;

                } else {
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

    /**
     * check that every column name is unique
     */
    #uniqueColumnNames() {
        const nameArray = [];
        for (const column of this.#columns) {
            let suffix = '', cntr = 1;

            while (nameArray.indexOf((column.name + suffix).toLowerCase()) !== -1) {
                cntr++;
                suffix = ' ' + cntr;
            }

            nameArray.push((column.name + suffix).toLowerCase());

            if (suffix) {
                column.name += suffix;
            }
        }
    }

}

