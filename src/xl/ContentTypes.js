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
            //{ PartName:"/xl/sharedStrings.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" },
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