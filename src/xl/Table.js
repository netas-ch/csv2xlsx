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