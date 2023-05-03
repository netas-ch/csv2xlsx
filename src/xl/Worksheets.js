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
import {XmlBuilder} from '../xml/XmlBuilder.js';
import {Utils} from '../util/Utils.js';

export class Worksheets {

    /**
     * Workbook.xml
     * @param {Array} columnTypes
     * @param {Array} columns
     * @param {Array} rows
     * @returns {String}
     */
    static sheet(columnTypes, columns, rows) {
        const xmlnsX14ac = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac';
        const xml = new XmlBuilder('worksheet', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        xml.setAttribute('root', 'Ignorable', 'x14ac', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        const tableDimension = 'A1:' + Utils.numericToAlphaColumn(columns.length) + rows.length.toString();

        xml.createAppend('root', 'dimension', null, { ref: tableDimension });

        // sheetFormatPr
        let sheetFormatPr = xml.createAppend('root', 'sheetFormatPr', null, { baseColWidth: '10', defaultRowHeight: '14.5' });
        xml.setAttribute(sheetFormatPr, 'x14ac:dyDescent', '0.35', xmlnsX14ac);

        // special column widths
        let cols = null;
        for (const column of columns) {
            if (column.width && column.width !== 10) {

                if (cols === null) {
                    cols = xml.createAppend('root', 'cols');
                }

                xml.createAppend(cols, 'col', null, { min: columns.indexOf(column)+1, max: columns.indexOf(column)+1, width: column.width, customWidth: 1 });
            }
        }

        // rows
        const sheetData = xml.createAppend('root', 'sheetData');

        for (const row of rows) {
            const rowNr = rows.indexOf(row) + 1;
            let rowEl = xml.createAppend(sheetData, 'row', null, {r: rowNr.toString(), spans: '1:' + columns.length.toString() });
            xml.setAttribute(rowEl, 'x14ac:dyDescent', '0.35', xmlnsX14ac);

            // columns
            for (const column of columns) {
                const colNr = columns.indexOf(column) + 1, attributes = {}, val = Worksheets.#getValueAsStr(row[column.rowKey]);

                // attributes
                attributes.r = Utils.numericToAlphaColumn(colNr) + rowNr.toString();

                // field type
                if (rowNr === 1 || !column.type || column.type === 'text') {

                    if (val !== '') {
                        attributes.t = 'str';
                    }

                } else if (column.type) {
                    attributes.s = columnTypes.indexOf(column.columnType) + 1;
                }

                // column
                const cEl = xml.createAppend(rowEl, 'c', null, attributes);

                // column value
                if (val !== '') {

                    // first row as string as its the header
                    if (rowNr === 1 || !column.type || column.type === 'text') {

                        // inline string
                        //xml.createAppend(xml.createAppend(cEl, 'is'), 't', null, null, val);
                        xml.createAppend(cEl, 'v', null, null, val);

                    } else {

                        // numeric value
                        xml.createAppend(cEl, 'v', null, null, val);
                    }
                }
            }
        }

        // tableParts
        const tableParts = xml.createAppend('root', 'tableParts', null, {count: '1'});
        const tablePart = xml.createAppend(tableParts, 'tablePart');
        xml.setAttribute(tablePart, 'id', 'rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        return xml.getXml();
    }

    /**
     * \xl\worksheets\_rels\sheet1.xml.rels file
     * @returns {String}
     */
    static rels() {
        const xml = new XmlBuilder('Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships');
        xml.createAppend('root', 'Relationship', null, {Id:'rId1', Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table', Target:'../tables/table1.xml'});
        return xml.getXml();
    }


    /**
     * return a value as a string
     * @param {Mixed} rawValue
     * @returns {String}
     */
    static #getValueAsStr(rawValue) {
        if (rawValue === '' || rawValue === null || typeof rawValue === 'undefined') {
            return '';
        }
        if (typeof rawValue === 'string') {
            return rawValue;
        }
        if (typeof rawValue === 'number') {
            if (isNaN(rawValue)) {
                return '';
            } else {
                return rawValue.toString();
            }
        }
        if (rawValue instanceof Date) {
            return Utils.dateToExcelTimestamp(rawValue).toString();
        }
        if (rawValue && rawValue.toString) {
            return rawValue.toString();
        }

        return '';
    }

}