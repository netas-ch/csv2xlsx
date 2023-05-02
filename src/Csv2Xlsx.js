

class Csv2Xlsx {

    /**
     * Converts a CSV to a xlsx Table
     * @param {String} csvUrl URL to Fetch the csv
     * @param {String} charset of CSV, default to Windows-1252
     * @param {Function} updateFn
     * @param {Bool} returnAsLink true, to return a <a> element instead of a Blob href
     * @returns {String|Element}
     */
    static async convertCsv(csvUrl, charset=null, updateFn=null, returnAsLink=false) {
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

            const buf = await rep.arrayBuffer();
            const decoder = new TextDecoder(charset);
            const rawCsv = decoder.decode(buf);

            // convert csv to json
            const csvp = await import('./csv/CsvProcessing.js');
            const cp = new csvp.CsvProcessing(rawCsv);
            const csvData = cp.getCsvData();

            // make a zip
            const spst = await import('./xl/Spreadsheet.js');
            const sp = new spst.Spreadsheet('filename.xlsx', csvData, {});

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