/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-12
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