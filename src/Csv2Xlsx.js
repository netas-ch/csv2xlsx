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
                    let repObj = null;
                    if (rep.headers.get('Content-Type') === 'application/json') {
                        repObj = await rep.json();
                    }

                    updateFn({state:'fail', msg: 'download failed: ' + rep.statusText, code: rep.status, responseJson: repObj});
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

            if (rep.headers) {
                const ctpe = rep.headers.get('Content-Type');
                if (ctpe && !ctpe.match(/text\/csv/i) && !ctpe.match(/application\/octet-stream/i)) {
                    throw new Error('invalid content type');
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