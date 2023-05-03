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

export class Styles {

    /**
     * /xl/styles.xml
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

        const cellXfs = xml.createAppend('root', 'cellXfs', null, {count: (columnTypes.length+1).toString() });
        xml.createAppend(cellXfs, 'xf', null, {numFmtId: '0', fontId: "0", fillId: "0", borderId: "0", xfId: "0" });

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

}