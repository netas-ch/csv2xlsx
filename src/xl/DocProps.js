/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-12
 */
import {XmlBuilder} from '../xml/XmlBuilder.js';

export class DocProps {

    /**
     * \docProps\app.xml file
     * @param {Array|null} sheets
     * @returns {unresolved}
     */
    static app(sheets=null) {

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


        const xml = new XmlBuilder('coreProperties', xmlns_cp);

        xml.createAppend('root', 'title', xmlns_dc, null, title ? title : 'My Document');
        xml.createAppend('root', 'subject', xmlns_dc, null, subject ? subject : '');
        xml.createAppend('root', 'creator', xmlns_dc, null, creator ? creator : '');
        xml.createAppend('root', 'lastModifiedBy', xmlns_cp, null, lastModifiedBy ? lastModifiedBy : (creator ? creator : ''));

        if (!created) {
            created = new Date();
        }
        const cr = xml.createAppend('root', 'created', xmlns_dcterms, null, created.toISOString());
        xml.setAttribute(cr, 'xsi:type', 'dcterms:W3CDTF', xmlns_xsi);

        if (!modified) {
            modified = new Date();
        }
        const md = xml.createAppend('root', 'modified', xmlns_dcterms, null, modified.toISOString());
        xml.setAttribute(md, 'xsi:type', 'dcterms:W3CDTF', xmlns_xsi);

        return xml.getXml();
    }

}