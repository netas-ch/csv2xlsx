/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-12
 */
import {XmlBuilder} from '../xml/XmlBuilder.js';

export class ContentTypes {

    /**
     * [Content_Types].xml
     * @param {Array|null} sheets
     * @returns {String}
     */
    static contentTypes(sheets=null) {
        const xml = new XmlBuilder('Types', 'http://schemas.openxmlformats.org/package/2006/content-types');

        xml.createAppend('root', 'Default', null, {Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml'});
        xml.createAppend('root', 'Default', null, {Extension: 'xml', ContentType: 'application/xml'});

        const overrides = [
            { PartName:"/xl/workbook.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" },
           // { PartName:"/xl/worksheets/sheet1.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" },
           // { PartName:"/xl/worksheets/sheet2.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" },
           // { PartName:"/xl/worksheets/sheet3.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" },
            { PartName:"/xl/theme/theme1.xml", ContentType: "application/vnd.openxmlformats-officedocument.theme+xml" },
            { PartName:"/xl/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" },
            { PartName:"/xl/sharedStrings.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" },
           // { PartName:"/xl/tables/table1.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" },
            { PartName:"/docProps/core.xml", ContentType: "application/vnd.openxmlformats-package.core-properties+xml" },
            { PartName:"/docProps/app.xml", ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml" }
        ];

        if (sheets === null) {
            sheets = ['Sheet 1'];
        }
        for (let i=0; i<sheets.length; i++) {
            overrides.push({ PartName:"/xl/worksheets/sheet" + (i+1) + '.xml', ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" });
            overrides.push({ PartName:"/xl/tables/table" + (i+1) + '.xml', ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" });
        }
        for (let override of overrides) {
            xml.createAppend('root', 'Override', null, override);
        }

        return xml.getXml();
    }

    /**
     * \_rels\.rels file
     * @returns {String}
     */
    static rels() {
        const xml = new XmlBuilder('Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships');

        xml.createAppend('root', 'Relationship', null, {Id:'rId1', Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', Target:'docProps/app.xml'});
        xml.createAppend('root', 'Relationship', null, {Id:'rId2', Type:'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target:'docProps/core.xml'});
        xml.createAppend('root', 'Relationship', null, {Id:'rId3', Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', Target:'xl/workbook.xml'});

        return xml.getXml();
    }

}