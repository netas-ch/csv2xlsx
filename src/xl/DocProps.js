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