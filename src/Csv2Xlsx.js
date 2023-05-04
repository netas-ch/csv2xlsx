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

export class Csv2Xlsx {

    /**
     * Converts a CSV to a xlsx Table
     * @param {String} csvUrl URL to fetch the csv from
     * @param {String} filename
     * @param {Object} metaData
     *                  > .title
     *                  > .subject
     *                  > .creator
     *                  > .company
     *                  > .lastModifiedBy
     *                  > .created          (Date object, default to now)
     *                  > .modified         (Date object, default to now)
     * @param {String} charset of CSV, default to Windows-1252
     * @param {Function} updateFn you can set a callback function to receive status informations.
     * @param {Bool} returnAsLink false, to return a Blob href  instead of a <a> element
     * @param {String} csvSeparator csv separator, null for auto
     * @param {Object} formatCodes excel format codes. defaults to
     *                  > .date: 'dd.mm.yyyy'
     *                  > .datetime: 'dd.mm.yyyy h:mm'
     *                  > .number: '#,##0'
     *                  > .float: '#,##0.0'
     *
     * @returns {Promise} which resolves to {String|Element} Blob URL as String or Element depending on {returnAsLink}, default to String
     */
    static async convertCsv(csvUrl, filename=null, metaData=null, charset=null, updateFn=null, returnAsLink=true, csvSeparator=null, formatCodes=null) {
        try {

            // default csv excel charset is Windows-1252
            if (!charset) {
                charset = 'Windows-1252';
            }

            if (updateFn) {
                updateFn({state:'download'});
            }

            let rep = await fetch(csvUrl, { cache: 'no-store' });
            if (!rep.ok) {
                if (updateFn) {
                    updateFn({state:'fail', msg: 'download failed: ' + rep.statusText});
                }
                return;
            }

            if (updateFn) {
                updateFn({state:'processing'});
            }

            // get filename from response header
            if (!filename && rep.headers) {
                // Content-Disposition: attachment; filename="super.csv"
                let disp = rep.headers.get('Content-Disposition');
                if (disp) {
                    const fnM = disp.match(/filename=\"([^\"]+)\"/);
                    filename = fnM[1];
                    if (filename.substr(filename.length-4).toLowerCase() === '.csv') {
                        filename = filename.substr(0, filename.length-4);
                    }
                }
            }

            // get filename from csv url
            if (!filename) {
                const csM = csvUrl.match(/\w+\.csv$/i);
                if (csM) {
                    filename = csM[0].substr(0, csM[0].length-4);
                }
            }

            // default filename
            if (!filename) {
                filename = 'document';
            }

            // add xlsx extension
            if (filename.substr(filename.length-5).toLowerCase() !== '.xlsx') {
                filename += '.xlsx';
            }

            // metaData object
            if (typeof metaData !== 'object') {
                metaData = {};
            }

            // default title = filename
            if (typeof metaData.title !== 'string') {
                metaData.title = filename.substr(0, filename.length-5).replace('_', ' ');
            }

            const buf = await rep.arrayBuffer();
            const decoder = new TextDecoder(charset);
            const rawCsv = decoder.decode(buf);

            // convert csv to json
            const csvp = await import('./csv/CsvProcessing.js');
            const cp = new csvp.CsvProcessing(rawCsv, csvSeparator, formatCodes);
            const csvData = cp.getCsvData();

            // create the spreadsheet
            const spst = await import('./xl/Spreadsheet.js');
            const sp = new spst.Spreadsheet(filename, csvData, metaData);
            const convData = sp.getXlsx(returnAsLink);

            if (updateFn && returnAsLink) {
                updateFn({state:'finished', aElement: convData });

            } else if (updateFn && !returnAsLink) {
                updateFn({state:'finished', blobHref: convData });
            }

            return convData;

        } catch (e) {
            if (updateFn) {
                updateFn({state:'fail', msg: e.message ? e.message : e});
            }
        }
    }
}