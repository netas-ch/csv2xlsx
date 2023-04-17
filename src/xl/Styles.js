/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-12
 */
import {XmlBuilder} from '../xml/XmlBuilder.js';

export class Styles {

    /**
     * Workbook.xml
     * @param {Array} columnTypes
     * @returns {String}
     */
    static styles(columnTypes) {
        const xmlnsX14ac = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac';
        const xml = new XmlBuilder('styleSheet', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        xml.setAttribute('root', 'Ignorable', 'x14ac', 'http://schemas.openxmlformats.org/markup-compatibility/2006');

        const numFmts = xml.createAppend('root', 'numFmts', null, {count: columnTypes.length.toString() });
        let nextId = 256;
        for (const columnType of columnTypes) {
            if (!columnType.numFmtId) {
                columnType.numFmtId = nextId;
                nextId++;
            }
            xml.createAppend(numFmts, 'numFmt', null, {numFmtId: columnType.numFmtId.toString(), formatCode: columnType.formatCode ? columnType.formatCode : '@' });
        }

        // Fonts (fix Calibri)
        const fonts = xml.createAppend('root', 'fonts', null, {count: '1' });
        xml.setAttribute(fonts,'x14ac:knownFonts', '1', xmlnsX14ac);

        const font = xml.createAppend(fonts, 'font');
        xml.createAppend(font, 'sz', null, {val: '11'});
        xml.createAppend(font, 'color', null, {theme: '1'});
        xml.createAppend(font, 'name', null, {val: 'Calibri'});
        xml.createAppend(font, 'family', null, {val: '2'});
        xml.createAppend(font, 'scheme', null, {val: 'minor'});

        // fills
        const fills = xml.createAppend('root', 'fills', null, {count: '2' });
        xml.createAppend(xml.createAppend(fills, 'fill'), 'patternFill', null, {patternType: 'none' });
        xml.createAppend(xml.createAppend(fills, 'fill'), 'patternFill', null, {patternType: 'gray125' });

        // borders
        const border = xml.createAppend(xml.createAppend('root', 'borders', null, {count: '1' }), 'border');
        xml.createAppend(border, 'left');
        xml.createAppend(border, 'right');
        xml.createAppend(border, 'top');
        xml.createAppend(border, 'bottom');
        xml.createAppend(border, 'diagonal');

        // cellStyleXfs
        xml.createAppend(xml.createAppend('root', 'cellStyleXfs', null, {count: '1' }), 'xf', null, { numFmtId: '0', fontId: '0', fillId: '0', borderId: '0' });

        const cellXfs = xml.createAppend('root', 'cellXfs', null, {count: columnTypes.length.toString() });
        for (const columnType of columnTypes) {
            xml.createAppend(cellXfs, 'xf', null, {numFmtId: columnType.numFmtId.toString(), fontId: "0", fillId: "0", borderId: "0", xfId: "0", applyNumberFormat: columnType.applyNumberFormat ? '1' : '0' });
        }

        // cellStyles
        xml.createAppend(xml.createAppend('root', 'cellStyles', null, {count: '1' }), 'cellStyle', null, { name: "Standard", xfId: "0", builtinId: "0" });

        const dxfs = xml.createAppend('root', 'dxfs', null, {count: columnTypes.length.toString() });
        for (const columnType of columnTypes) {
            xml.createAppend(xml.createAppend(dxfs, 'dxf'), 'numFmt', null, {numFmtId: columnType.numFmtId.toString(), formatCode: columnType.formatCode ? columnType.formatCode : '@' });
        }

        // tableStyles
        xml.createAppend('root', 'tableStyles', null, { count: "0", defaultTableStyle: "TableStyleMedium2", defaultPivotStyle: "PivotStyleLight16" });


        return xml.getXml();
    }

    /**
     * \_rels\.rels file
     * @param {Array|null} sheets
     * @returns {String}
     */
    static rels(sheets=null) {
        const xml = new XmlBuilder('Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships');

        if (sheets === null) {
            sheets = ['Sheet 1'];
        }

        for (let i=0; i<sheets.length; i++) {
            xml.createAppend('root', 'Relationship', null, {Id:'rId_sheet' + (i+1), Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target:'worksheets/sheet' + (i+1) + '.xml'});
        }
        xml.createAppend('root', 'Relationship', null, {Id:'rId_theme1', Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', Target:'theme/theme1.xml'});
        xml.createAppend('root', 'Relationship', null, {Id:'rId_style1', Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', Target:'styles.xml'});

        return xml.getXml();
    }

}