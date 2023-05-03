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

export class DocProps {

    /**
     * \docProps\app.xml file
     * @param {Array|null} sheets
     * @param {String} company
     * @returns {unresolved}
     */
    static app(sheets=null, company=null) {

        if (sheets === null) {
            sheets = ['Sheet 1'];
        }

        const xml = new XmlBuilder('Properties', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
        const vtNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

        xml.createAppend('root', 'Application', null, null, 'CSV2XLSX');

        const headingPairs = xml.createAppend('root', 'HeadingPairs');
        let vector = xml.createAppend(headingPairs, 'vt:vector', vtNamespace, {size: '2', baseType: 'variant'});

        let variant = xml.createAppend(vector, 'vt:variant', vtNamespace);
        xml.createAppend(variant, 'vt:lpstr', vtNamespace, null, 'Worksheets');

        variant = xml.createAppend(vector, 'vt:variant', vtNamespace);
        xml.createAppend(variant, 'vt:i4', vtNamespace, null, sheets.length.toString());

        // TitlesOfParts
        const titlesOfParts = xml.createAppend('root', 'TitlesOfParts');
        vector = xml.createAppend(titlesOfParts, 'vt:vector', vtNamespace, {size: sheets.length.toString(), baseType: 'lpstr'});

        for (const sheetName of sheets) {
            xml.createAppend(vector, 'vt:lpstr', vtNamespace, null, sheetName);
        }

        // company
        if (company) {
            xml.createAppend('root', 'Company', null, null, company);
        }

        return xml.getXml();
    }

    /**
     * \docProps\core.xml file
     * @param {String} title
     * @param {String} subject
     * @param {String} creator
     * @param {String} lastModifiedBy
     * @param {Date|null} created
     * @param {Date|null} modified
     * @returns {unresolved}
     */
    static core(title='My Document', subject='', creator='', lastModifiedBy='', created=null, modified=null) {
        const xmlns_cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            xmlns_dc='http://purl.org/dc/elements/1.1/',
            xmlns_dcterms='http://purl.org/dc/terms/',
            xmlns_dcmitype='http://purl.org/dc/dcmitype/',
            xmlns_xsi='http://www.w3.org/2001/XMLSchema-instance';


        const xml = new XmlBuilder('cp:coreProperties', xmlns_cp);

        xml.createAppend('root', 'dc:title', xmlns_dc, null, title ? title : 'My Document');
        xml.createAppend('root', 'dc:subject', xmlns_dc, null, subject ? subject : '');
        xml.createAppend('root', 'dc:creator', xmlns_dc, null, creator ? creator : '');
        xml.createAppend('root', 'cp:lastModifiedBy', xmlns_cp, null, lastModifiedBy ? lastModifiedBy : (creator ? creator : ''));

        if (!created || !(created instanceof Date)) {
            created = new Date();
        }
        const cr = xml.createAppend('root', 'dcterms:created', xmlns_dcterms, null, created.toISOString());
        xml.setAttribute(cr, 'xsi:type', 'dcterms:W3CDTF', xmlns_xsi);

        if (!modified || !(modified instanceof Date)) {
            modified = new Date();
        }

        const md = xml.createAppend('root', 'dcterms:modified', xmlns_dcterms, null, modified.toISOString());
        xml.setAttribute(md, 'xsi:type', 'dcterms:W3CDTF', xmlns_xsi);

        return xml.getXml();
    }

}