/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-18
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
        this.#zip.addFileFromString('docProps/app.xml', DocProps.app());

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