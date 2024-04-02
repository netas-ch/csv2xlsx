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

export class Workbook {

    /**
     * Workbook.xml
     * @param {Array|null} sheets
     * @returns {String}
     */
    static workbook(sheets=null) {
        const xml = new XmlBuilder('workbook', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        xml.createAppend('root', 'workbookPr', null, {codeName: 'ThisWorkbook'});
        const sheetsEl = xml.createAppend('root', 'sheets');

        if (sheets === null) {
            sheets = ['Sheet 1'];
        }

        for (let i=0; i<sheets.length; i++) {
            const se = xml.createAppend(sheetsEl, 'sheet', null, {name: sheets[i], sheetId: (i+1).toString()});
            xml.setAttribute(se, 'r:id', 'rId' + (i+1), 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        }

        return xml.getXml();
    }

    /**
     * \xl\_rels\workbook.xml.rels file
     * @param {Array|null} sheets
     * @returns {String}
     */
    static rels(sheets=null) {
        const xml = new XmlBuilder('Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships');

        if (sheets === null) {
            sheets = ['Sheet 1'];
        }

        for (let i=0; i<sheets.length; i++) {
            xml.createAppend('root', 'Relationship', null, {Id:'rId' + (i+1), Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target:'worksheets/sheet' + (i+1) + '.xml'});
        }
        xml.createAppend('root', 'Relationship', null, {Id:'rId' + (sheets.length+1), Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', Target:'theme/theme1.xml'});
        xml.createAppend('root', 'Relationship', null, {Id:'rId' + (sheets.length+2), Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', Target:'styles.xml'});

        return xml.getXml();
    }

}