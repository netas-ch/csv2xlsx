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
import {ContentTypes} from './ContentTypes.js';
import {DocProps} from './DocProps.js';
import {Styles} from './Styles.js';
import {Table} from './Table.js';
import {Theme} from './Theme.js';
import {Workbook} from './Workbook.js';
import {Worksheets} from './Worksheets.js';

import {NullZipArchive} from '../zip/NullZipArchive.js';

export class Spreadsheet {
    #zip;

    constructor(filename, csv, metadata) {
        const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        this.#zip = new NullZipArchive(filename, false, mime);

        // build the zip

        // [Content_Types].xml
        this.#zip.addFileFromString('[Content_Types].xml', ContentTypes.contentTypes());

        // /_rels/.rels
        this.#zip.addFileFromString('_rels/.rels', ContentTypes.rels());

        // /docProps/app.xml
        this.#zip.addFileFromString('docProps/app.xml', DocProps.app(null, metadata.company ? metadata.company : null));

        // /docProps/core.xml
        this.#zip.addFileFromString('docProps/core.xml', DocProps.core(
                metadata.title ? metadata.title : 'My Document',
                metadata.subject ? metadata.subject : '',
                metadata.creator ? metadata.creator : 'unknown',
                metadata.lastModifiedBy ? metadata.lastModifiedBy : '',
                metadata.created ? metadata.created : null,     // date of creation
                metadata.modified ? metadata.modified : null
            ));

        // styles
        this.#zip.addFileFromString('xl/styles.xml', Styles.styles(csv.columnTypes));

        // workbook
        this.#zip.addFileFromString('xl/workbook.xml', Workbook.workbook());
        this.#zip.addFileFromString('xl/_rels/workbook.xml.rels', Workbook.rels());

        // theme
        this.#zip.addFileFromString('xl/theme/theme1.xml', Theme.theme());

        // tables
        this.#zip.addFileFromString('xl/tables/table1.xml', Table.table(csv.columnTypes, csv.columns, csv.rows));

        // worksheets
        this.#zip.addFileFromString('xl/worksheets/sheet1.xml', Worksheets.sheet(csv.columnTypes, csv.columns, csv.rows));
        this.#zip.addFileFromString('xl/worksheets/_rels/sheet1.xml.rels', Worksheets.rels());
    }

    /**
     * get the excel xlsx file as URL
     * @returns {unresolved}
     */
    getXlsx(asLink=false) {
        if (asLink) {
            return this.#zip.createDownloadLink();
        }
        return this.#zip.createDownloadUrl();
    }


}