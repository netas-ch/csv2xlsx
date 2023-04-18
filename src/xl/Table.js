/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-12
 */
import {XmlBuilder} from '../xml/XmlBuilder.js';
import {Utils} from '../util/Utils.js';

export class Table {

    /**
     * table.xml
     * @param {Array} columnTypes
     * @param {Array} columns
     * @param {Array} rows
     * @returns {String}
     */
    static table(columnTypes, columns, rows, tableId=1) {
        const xmlnsX14ac = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac';
        const xml = new XmlBuilder('table', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        const tableDimension = 'A1:' + Utils.numericToAlphaColumn(columns.length) + rows.length.toString();

        xml.setAttribute('root', 'id', tableId.toString());
        xml.setAttribute('root', 'name', 'Tabelle' + tableId);
        xml.setAttribute('root', 'displayName', 'Tabelle' + tableId);
        xml.setAttribute('root', 'ref', tableDimension);

        xml.createAppend('root', 'autoFilter', null, {ref: tableDimension});
        const tableColumns = xml.createAppend('root', 'tableColumns', null, {count: columns.length });

        let id = 0;
        for (const column of columns) {
            id++;

            let attributes = {
                id: id.toString(),
                name: column.name
            };

            if (column.columnType) {
                if (columnTypes.indexOf(column.columnType) !== -1) {
              //      attributes.dataDxfId = columnTypes.indexOf(column.columnType);
                }
                if (column.columnType.totalsRowFunction) {
                    attributes.totalsRowFunction = column.columnType.totalsRowFunction;
                }
                if (column.columnType.totalsRowLabel) {
                    attributes.totalsRowLabel = column.columnType.totalsRowLabel;
                }
            }

            xml.createAppend(tableColumns, 'tableColumn', null, attributes);
        }

        // tableStyleInfo
        xml.createAppend('root', 'tableStyleInfo', null, {
            name: "TableStyleLight1",
            showFirstColumn: "0",
            showLastColumn: "0",
            showRowStripes: "1",
            showColumnStripes: "0"
        });

        return xml.getXml();
    }

}