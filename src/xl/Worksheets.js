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
        xml.setAttribute(sheetFormatPr, 'dyDescent', '0.35', xmlnsX14ac);

        // special column widths
        const cols = xml.createAppend('root', 'cols');
        for (const column of columns) {
            if (column.width && column.width !== 10) {
                xml.createAppend(cols, 'col', null, { min: columns.indexOf(column)+1, max: columns.indexOf(column)+1, width: column.width, customWidth: 1 });
            }
        }

        // rows
        let sheetData = xml.createAppend('root', 'sheetData');

        for (const row of rows) {
            const rowNr = rows.indexOf(row) + 1;
            let rowEl = xml.createAppend(sheetData, 'row', null, {r: rowNr.toString(), spans: '1:' + columns.length.toString() });
            xml.setAttribute(rowEl, 'dyDescent', '0.35', xmlnsX14ac);

            // columns
            for (const column of columns) {
                const colNr = columns.indexOf(column) + 1, attributes = {};

                // attributes
                attributes.r = Utils.numericToAlphaColumn(colNr) + rowNr.toString();

                // field type
                if (column.type === 'text') {
                    attributes.t = 'inlineStr';
                } else {
                    attributes.s = columnTypes.indexOf(column.columnType) + 1;
                }

                // column.columnType
                const cEl = xml.createAppend(rowEl, 'c', null, attributes);

                if (!column.type || column.type === 'text') {

                    // inline string
                    xml.createAppend(xml.createAppend(cEl, 'is'), 't', row[column.rowKey]);

                } else {

                    // numeric value
                    xml.createAppend(cEl, 'v', row[column.rowKey]);
                }
            }
        }



        return xml.getXml();
    }

    /**
     * \_rels\.rels file
     * @param {Array|null} sheets
     * @returns {String}
     */
    static rels(sheets=null) {
        const xml = new XmlBuilder('Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships');

        xml.createAppend('root', 'Relationship', null, {Id:'rId1', Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table', Target:'../tables/table1.xml'});


        return xml.getXml();
    }

}