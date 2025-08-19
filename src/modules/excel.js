const { exit } = require("process");

const COLOR_INPUT = '#C6E0B4';

var LISTS = null;

module.exports = {

    initialize: async function (lists) {
        
        LISTS = lists;

        await Excel.run(async (context) => {

            var sheet = await getSheet(context, 'Summary');

            // G2 must be set to a value since it is referenced in formulas on the BOM's which are created before the Summary
            // BOM formulas referencing Summary!$G$2 won't work if this isn't done

            sheet.getRange('G2').values = [[0]];

            await context.sync();
        });
    },
    addData: async function (sheetName, sheetData, options) {
    
        await Excel.run(async (context) => {

            var sheet = await getSheet(context, sheetName);

            if (options.bomId) sheetData[sheet.id.toString()] = {bomId: options.bomId};

            var rangeString = getRangeString({
                rows: options.data.length,
                columns: options.data[0].length
            });

            sheet.getRange(rangeString).values = options.data;

            if (options.ranges) {

                for (var i = 0; i < options.ranges.length; i++) {

                    var rangeOptions = options.ranges[i];

                    for (var j = 0; j < rangeOptions.range.length; j++) {
    
                        var range = sheet.getRange(rangeOptions.range[j]);

                        var rangeProperties = getRangeProperties(rangeOptions.range[j]);

                        var formulas = [];
                        var numberFormats = [];

                        for (var r = 0; r < rangeProperties.rows; r++) {

                            var formula = [];
                            var numberFormat = [];

                            for (var c = 0; c < rangeProperties.columns; c++) {

                                if (rangeOptions.formula) {

                                    formula.push('=' + rangeOptions.formula.replaceAll('?', (r + rangeProperties.firstRow)));
                                }

                                if (rangeOptions.numberFormat) {

                                    numberFormat.push(rangeOptions.numberFormat);
                                }
                            }

                            if (formula.length) formulas.push(formula);
                            if (numberFormat.length) numberFormats.push(numberFormat);
                        }

                        if (formulas.length) range.formulas = formulas;
                        if (numberFormats.length) range.numberFormat = numberFormats;

                        if (rangeOptions.color) {
                            
                            if (rangeOptions.color == 'white')
                                range.format.fill.clear();
                            else
                                range.format.fill.color = rangeOptions.color;
                        }

                        if (rangeOptions.horizontalAlignment) range.format.horizontalAlignment = rangeOptions.horizontalAlignment;
                        if (rangeOptions.bold) range.format.font.bold = true;
                        if (rangeOptions.groupByRows) {range.group(Excel.GroupOption.byRows); range.hideGroupDetails();}
                        if (rangeOptions.groupByColumns) range.group(Excel.GroupOption.byColumns);
                        if (rangeOptions.hideRows) range.rowHidden = rangeOptions.hideRows;
                        if (rangeOptions.hideColumns) range.columnHidden = rangeOptions.hideColumns;
                        if (rangeOptions.dataValidationRule) range.dataValidation.rule = rangeOptions.dataValidationRule;
                        if (rangeOptions.columnWidth) range.format.columnWidth = rangeOptions.columnWidth;
                    }
                }
            }

            if (options.autofitColumns) {

                var rangeString = getRangeString({
                    columns: options.autofitColumns
                });
                
                sheet.getRange(rangeString).format.autofitColumns();
            }

            await sheet.onChanged.add(handleWorksheetChange);

            await context.sync();
        });
    },
    reset: async function () {

        await Excel.run(async (context) => {

            let sheetsToDelete = [];

            let sheets = context.workbook.worksheets;
    
            sheets.load("items");

            await context.sync();

            for (let i = 0; i < sheets.items.length; i++) {

                let sheet = sheets.items[i];

                let range = sheet.getRange('A1');
                    
                range.load("values");

                await context.sync();

                if (range.values[0] == 'Quote' || range.values[0] == 'Quantity') sheetsToDelete.push(sheet) // Only act on quote worksheets
            }

            if (sheets.items.length == sheetsToDelete.length) sheets.add().activate(); // There must be at least one worksheet. Add new worksheet to let Excel name it

            for (let i = 0; i < sheetsToDelete.length; i++) {

                sheetsToDelete[i].delete();
            }

            await context.sync();
        });
    }
};

async function clear(context, sheetName = null) {

    var sheet = null;

    if (sheetName)
        sheet = context.workbook.worksheets.getItem(sheetName);
    else
        sheet = context.workbook.worksheets.getActiveWorksheet();

    var range = sheet.getRange('1:100000'); // getRange() is acceptable for everything but ungroup method

    range.ungroup(Excel.GroupOption.byRows);

    var range = sheet.getRange('A:Z'); // getRange() is acceptable for everything but ungroup method

    range.ungroup(Excel.GroupOption.byColumns);

    range.clear();

    sheet.getRange('A1').select();

    range.format.useStandardWidth = true;

    await context.sync();
}

async function getSheet(context, sheetName = null) {
    
    var sheet = null;

    if (sheetName) {

        sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);

        sheet.load("isNullObject");

        await context.sync();

        if (sheet.isNullObject) {

            sheet = context.workbook.worksheets.getActiveWorksheet();

            var range = sheet.getUsedRangeOrNullObject(true);

            range.load("isNullObject");

            await context.sync();

            if (range.isNullObject) {

                sheet.name = sheetName;
            }
            else {

                sheet = context.workbook.worksheets.add(sheetName);
            }

            await context.sync();
        }
        else {

            await clear(context, sheetName);
        }
    }
    else {

        sheet = context.workbook.worksheets.getActiveWorksheet();

        await clear(context);
    }

    if (sheetName == 'Summary') sheet.position = 0;

    return sheet;
}

async function handleWorksheetChange(eventArgs) {

    await Excel.run(async (context) => {

        if (eventArgs.changeType === Excel.DataChangeType.rowInserted) {

            var rows = eventArgs.address.split(':'); // This creates a two element array of the first and last rows inserted in string format

            var rowFirst = parseInt(rows[0]);
            var rowLast = parseInt(rows[1]);
            
            var sheet = context.workbook.worksheets.getActiveWorksheet();
        
            var range = sheet.getUsedRange();

            range.load("values");

            await context.sync();

            var isSummary = (range.values[0][0] == "Quote");

            if (isSummary) { // Is Summary

                if (!isItem(range.values, isSummary, rowFirst)) return;
            }
            else if (range.values[0][0] == "Quantity") { // Is BOM

                if (!isItem(range.values, isSummary, rowFirst)) {
                    
                    let isExp = isExpense(range.values, rowFirst);
                    if (!isExp) return;
                }
            }
            else {

                return;
            }

            for (var i = rowFirst; i <= rowLast; i++) {

                insertRow(sheet, i, {
                    isSummary: isSummary,
                    isExpense: isExp
                });
            }
        }

        await context.sync();
    });
}

async function insertRow(sheet, row, options) {

    if (options.isSummary) {

        sheet.getRange('A' + row + ':G' + row).values = [[1,'','',0,0,0,0]];

        sheet.getRange('A' + row + ':D' + row).format.fill.color = COLOR_INPUT;
    }
    else {

        if (options.isExpense) {
            
            sheet.getRange('A' + row + ':I' + row).values = [[1,'','',0,0,0,0,'','No']];

            sheet.getRange('A' + row).format.fill.color = COLOR_INPUT;
            sheet.getRange('C' + row + ':D' + row).format.fill.color = COLOR_INPUT;
            sheet.getRange('H' + row + ':I' + row).format.fill.color = COLOR_INPUT;

            sheet.getRange('C' + row).dataValidation.rule = {
                list: {
                    inCellDropDown: true,
                    source: LISTS.accounts.join(',')
                }
            }

            sheet.getRange('I' + row).dataValidation.rule = {
                list: {
                    inCellDropDown: true,
                    source: 'Yes,No'
                }
            }
        }
        else {

            sheet.getRange('A' + row + ':J' + row).values = [[1,'','',0,0,0,0,'','No','Ea']];

            sheet.getRange('A' + row + ':D' + row).format.fill.color = COLOR_INPUT;
            sheet.getRange('H' + row + ':M' + row).format.fill.color = COLOR_INPUT;

            sheet.getRange('I' + row).dataValidation.rule = {
                list: {
                    inCellDropDown: true,
                    source: 'Yes,No'
                }
            }

            sheet.getRange('J' + row).dataValidation.rule = {
                list: {
                    inCellDropDown: true,
                    source: LISTS.units.filter(unit => unit.type != 3).map(unit => unit.names).join(',') // Filter out labor units
                }
            }
        }
    }

    var range = sheet.getRange('A' + row);

    range.format.horizontalAlignment = 'center';
    range.format.font.bold = false;

    sheet.getRange('D' + row + ':G' + row).numberFormat = '$#,###.00';

    sheet.getRange('E' + row + ':G' + row).formulas = [[('=A?*D?').replaceAll('?', row),('=D?*(1+IF(I?="Yes",-1,1)*IF(ISNUMBER(H?),H?,Summary!$G$2))').replaceAll('?', row),('=A?*F?').replaceAll('?', row)]];
}

function isItem(values, isSummary, row) {

    for (var i = row; i > 0; i--) {

        if (!isNaN(values[i-1][0])) continue;

        if (values[i-1][0] != 'Items') return false;

        break;
    }

    for (var i = row; i < values.length; i++) {

        if (!isNaN(values[i-1][0])) continue;

        if (values[i-1][0] != (isSummary ? 'Total' : 'Labor')) return false;

        break;
    }

    return true;
}

function isExpense(values, row) {

    for (var i = row; i > 0; i--) {

        if (!isNaN(values[i-1][0])) continue;

        if (values[i-1][0] != 'Expenses') return false;

        break;
    }

    for (var i = row; i < values.length; i++) {

        if (!isNaN(values[i-1][0])) continue;

        if (values[i-1][0] != 'Total') return false;

        break;
    }

    return true;
}

function getRangeString(options) {

    var firstColumn = (options.firstColumn ? options.firstColumn : 'A');
    var firstRow = (options.firstRow ? options.firstRow : 1);

    return (firstColumn + (options.rows ? firstRow : '') + ':' + String.fromCharCode(firstColumn.charCodeAt(0) + options.columns - 1) + (options.rows ? firstRow + options.rows - 1 : ''));
}

function getRangeProperties(rangeString) {

    // Using this method is much faster then syncing to get the rowIndex, rowCount &columnCount properties

    var ranges = rangeString.split(':');

    var cp1 = getCellProperties(ranges[0]);

    var props = {
        rows: 1,
        columns: 1,
        firstRow: (cp1.row ? cp1.row : 1)
    };

    if (ranges.length == 1) return props;

    cp2 = getCellProperties(ranges[1]);

    if (cp1.row && cp2.row) props.rows = cp2.row - cp1.row + 1;
    if (cp1.column && cp2.column) props.columns = cp2.column - cp1.column + 1;

    return props;
}

function getCellProperties(rangeString) {

    var props = {};

    for (var i = 0; i < rangeString.length; i++) {

        var code = rangeString.charCodeAt(i);

        if (code >= 65 && code <= 90) { // column

            props.column = (i * 26) + (code - 64);
        }
        else { // row

            props.row = parseInt(rangeString.substring(i));

            break;
        }
    }

    return props;
}
