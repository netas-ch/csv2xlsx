/*
 * Copyright © 2023 Netas Ltd., Switzerland.
 * All rights reserved.
 * @author  Lukas Buchs, lukas.buchs@netas.ch
 * @date    2023-04-13
 */

export class Utils {

    // Excel Formate
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformats?view=openxml-2.8.1

    /**
     * Convert a String to a Number
     * Note that this function treats column letters as 1-based indices, so 'A' corresponds to 1, 'B' to 2, and so on.
     * @param {String} column
     * @returns {Number}
     */
    static alphaToNumericColumn(column) {
        let numericColumn = 0;
        const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

        for (let i = 0; i < column.length; i++) {
            const char = column.charAt(i).toUpperCase();
            const charIndex = alphabet.indexOf(char) + 1;
            numericColumn = numericColumn * 26 + charIndex;
        }

        return numericColumn;
    }

    /**
     * Convert a number to a string
     * Note that this function treats column letters as 1-based indices, so 'A' corresponds to 1, 'B' to 2, and so on.
     * @param {Number} column
     * @returns {String}
     */
    static numericToAlphaColumn(column) {
        let alphaColumn = '';
        const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

        while (column > 0) {
            const remainder = (column - 1) % 26;
            alphaColumn = alphabet.charAt(remainder) + alphaColumn;
            column = Math.floor((column - remainder) / 26);
        }

        return alphaColumn;
    }


    /**
     * Convert a Date to a Excel Timestamp
     * @param {Date} date
     * @returns {Number}
     */
    static dateToExcelTimestamp(date) {
        const epochStart = new Date('1899-12-30T00:00:00'); // Excel epoch starts from Dec 30, 1899
        const millisecondsPerDay = 24 * 60 * 60 * 1000; // Number of milliseconds in a day
        return (date.getTime() - epochStart.getTime()) / millisecondsPerDay;
        // Add 1 to account for Excel's leap year bug ?
    }

    /**
     * Check if a string value is a date and time
     * @param {String} str
     * @returns {Boolean}
     */
    static stringIsDateTime(str) {
        let mtch = str.match(/^(\d{4})(?:\-|\/)([0-1][0-9])(?:\-|\/)([0-3][0-9])(?:(?:T| )([0-2][0-9])\:([0-5][0-9])\:([0-5][0-9](?:\.\d+)?))?$/);
        if (mtch && mtch[4]) {
            return true;
        }

        mtch = str.match(/^([0-3][0-9])\.([0-1][0-9])\.(\d{4})(?: ([0-2][0-9])\:([0-5][0-9])(?:\:([0-5][0-9](?:\.\d+)?))?)?$/);
        if (mtch && mtch[4]) {
            return true;
        }

        return false;
    }

    /**
     * Check if a string value is a date
     * @param {String} str
     * @returns {Boolean}
     */
    static stringIsDate(str) {
        let mtch = str.match(/^(\d{4})(?:\-|\/)([0-1][0-9])(?:\-|\/)([0-3][0-9])(?:(?:T| )([0-2][0-9])\:([0-5][0-9])\:([0-5][0-9](?:\.\d+)?))?$/);
        if (mtch) {
            return true;
        }

        mtch = str.match(/^([0-3][0-9])\.([0-1][0-9])\.(\d{4})(?: ([0-2][0-9])\:([0-5][0-9])(?:\:([0-5][0-9](?:\.\d+)?))?)?$/);
        if (mtch) {
            return true;
        }

        return false;
    }

    /**
     * Check if a string value is a integer
     * @param {String} str
     * @returns {Boolean}
     */
    static stringIsInteger(str) {
        return !isNaN(parseInt(str)) && parseInt(str).toString() === str;
    }

    /**
     * Check if a string value is a float
     * @param {String} str
     * @returns {Boolean}
     */
    static stringIsFloat(str) {
        return !!(!Utils.stringIsInteger(str) && !isNaN(parseFloat(str)) && str.match(/^\d+(?:\.\d+)?$/));
    }

    /**
     * Convert a String to a Date
     * @param {String} dateStr
     * @returns {null|Date}
     */
    static parseDate(dateStr) {
        let mtch = dateStr.match(/^(\d{4})(?:\-|\/)([0-1][0-9])(?:\-|\/)([0-3][0-9])(?:(?:T| )([0-2][0-9])\:([0-5][0-9])\:([0-5][0-9](?:\.\d+)?))?$/);
        if (mtch) {
            return new Date(
                    parseInt(mtch[1]),
                    parseInt(mtch[2])-1,
                    parseInt(mtch[3]),
                    parseInt(mtch[4] ? mtch[4] : 0),
                    parseInt(mtch[5] ? mtch[5] : 0),
                    parseInt(mtch[6] ? mtch[6] : 0)
                );
        }

        mtch = dateStr.match(/^([0-3][0-9])\.([0-1][0-9])\.(\d{4})(?: ([0-2][0-9])\:([0-5][0-9])(?:\:([0-5][0-9](?:\.\d+)?))?)?$/);
        if (mtch) {
            return new Date(
                    parseInt(mtch[3]),
                    parseInt(mtch[2])-1,
                    parseInt(mtch[1]),
                    parseInt(mtch[4] ? mtch[4] : 0),
                    parseInt(mtch[5] ? mtch[5] : 0),
                    parseInt(mtch[6] ? mtch[6] : 0)
                );
        }

        if (!isNaN(Date.parse(dateStr))) {
            return new Date(dateStr);
        }

        return null;
    }

}