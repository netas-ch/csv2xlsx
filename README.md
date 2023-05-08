# csv2xlsx
JavaScript csv to spreadsheet (xlsx) converter.
Converts columns automatically to the correct column type and formats the data as a table, which allows simple filtering of data.

## usage

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
                         > .datetime: 'dd.mm.yyyy h:mm'
                         > .number: '0'
                         > .float: '0.0'

        @return {Promise} which resolves to {String|Element} Blob URL or Element depending on {returnAsLink}, default to String

## demo
check out [https://raw.githack.com/netas-ch/csv2xlsx/main/_demo/demo.html](_demo/demo.html) for a simple demo

## license
Copyright © 2023 Lukas Buchs, netas.ch, MIT licensed.

### used third party code
 * https://github.com/Neovici/nullxlsx Copyright © 2020 Neovici, Apache License 2.0
 * https://github.com/FlatFilers/csvjson-csv2json Copyright © 2019 Martin Drapeau, MIT License