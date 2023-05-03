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