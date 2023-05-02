/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-12
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