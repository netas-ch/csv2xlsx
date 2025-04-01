# csv2xlsx
JavaScript csv to spreadsheet (xlsx) converter.
Converts columns automatically to the correct column type and formats the data as a table, which allows simple filtering of data.
For numerical values and floats, a sum row is added, for percentage values the average value.

## parameter documentation

    Csv2Xlsx.convertCsv(csvUrl, filename, metaData, charset, updateFn, returnAsLink, csvSeparator, formatCodes)
        @param {String} csvUrl              URL to fetch the csv from
        @param {String} filename            of the generated file
        @param {Object} metaData            metaData of the generated xlsx
                         > .title
                         > .subject
                         > .creator
                         > .company
                         > .lastModifiedBy
                         > .created         (Date object, default to now)
                         > .modified        (Date object, default to now)

        @param {String} charset             charset of CSV, default to Windows-1252
        @param {Function} updateFn          you can set a callback function to receive status informations.
        @param {Bool} returnAsLink          false, to return a Blob url as string instead of a <a> element
        @param {String} csvSeparator        csv separator, null for auto
        @param {Object} formatCodes         excel format codes. defaults to
                         > .date: 'dd.mm.yyyy'
                         > .datetime: 'dd.mm.yyyy hh:mm'
                         > .number: '0'
                         > .float: '0.0'
                         > .percentage: '0%'

        @return {Promise} which resolves to {String|Element} Blob URL or Element depending on {returnAsLink}, default to String

## example code
```javascript
async function downloadAsExcel() {
    const cvrt = await import('src/Csv2Xlsx.js');

    // Meta-Data of the xlsx
    const metaData = {
        title: 'My Demo Spreadsheet',
        subject: 'A Demo for CSV to XLSX',
        creator: 'John Doe',
        company: 'Super Doe Ltd.',
        lastModifiedBy: 'John Doe',
        created: new Date(2020, 0, 1, 11, 21),
        modified: null // now
    };

    // download the csv and convert it to a excel file
    const aTag = await cvrt.Csv2Xlsx.convertCsv('my/demo/data.csv', 'mynewfilename.xlsx', metaData);

    // return value is a <a> element, add it to the DOM to start the download
    document.body.appendChild(aTag);

    // start the download automatically
    aTag.click();
}
```
## browser support
the latest versions of
* Firefox (Gecko)
* Chrome/Edge (Blink)
* Safari (Webkit)

## demo
check out [_demo/demo.html](https://raw.githack.com/netas-ch/csv2xlsx/main/_demo/demo.html) for a simple demo

## license
Copyright © 2024 Lukas Buchs, [netas.ch](https://netas.ch), [Apache License, Version 2.0](./LICENSE).

### used third party code
 * https://github.com/Neovici/nullxlsx Copyright © 2020 Neovici, Apache License 2.0
 * https://github.com/FlatFilers/csvjson-csv2json Copyright © 2019 Martin Drapeau, MIT License