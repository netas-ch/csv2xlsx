

class Csv2Xlsx {

    static async convertCsv(csvUrl, charset, updateFn) {
        try {

            // default excel charset is Windows-1252
            if (!charset) {
                charset = 'Windows-1252';
            }

            updateFn({state:'download'});

            let rep = await fetch(csvUrl, { cache: 'no-cache' });
            if (!rep.ok) {
                updateFn({state:'fail', msg: 'download failed: ' + rep.statusText});
                return;
            }

            updateFn({state:'processing'});

            const buf = await rep.arrayBuffer();
            const decoder = new TextDecoder(charset);
            const rawCsv = decoder.decode(buf);

            // convert csv to json
            const csvp = await import('./csv/CsvProcessing.js');
            const cp = new csvp.CsvProcessing(rawCsv);
            const csvData = cp.getCsvData();

            // make a zip

            console.log(csvData);
            const spst = await import('./xl/Spreadsheet.js');
            const sp = new spst.Spreadsheet('filename.xlsx', csvData, {});

            const src = sp.getXlsx(true);

            updateFn({state:'finished', src:src});

            return src;

        } catch (e) {
            console.log(e);
            updateFn({state:'fail', msg: e.message ? e.message : e});
        }
    }
}