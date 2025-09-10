module.exports = {

    getSheet: async function (context, name = null) {
    
        var sheet = null;

        if (name) {

            sheet = context.workbook.worksheets.getItemOrNullObject(name);

            sheet.load("isNullObject");

            await context.sync();

            if (sheet.isNullObject) {

                sheet = context.workbook.worksheets.getActiveWorksheet();

                var range = sheet.getUsedRangeOrNullObject(true);

                range.load("isNullObject");

                await context.sync();

                if (range.isNullObject) {

                    sheet.name = name;
                }
                else {

                    sheet = context.workbook.worksheets.add(name);
                }

                await context.sync();
            }
            else {

                await clear(context, name);
            }
        }
        else {

            sheet = context.workbook.worksheets.getActiveWorksheet();

            await clear(context);
        }

        if (name == 'Summary') sheet.position = 0;

        return sheet;
    },
    setSheet: async function (context, sheet, options) {

        if (options.data) {

            var rangeString = getRangeString({
                rows: options.data.length,
                columns: options.data[0].length
            });

            sheet.getRange(rangeString).values = options.data;
        }

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

                            if (rangeOptions.formula) formula.push('=' + rangeOptions.formula.replaceAll('?', (r + rangeProperties.firstRow)));
                            if (rangeOptions.numberFormat) numberFormat.push(rangeOptions.numberFormat);
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

                    if (rangeOptions.values) range.values = rangeOptions.values;
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

        await context.sync();
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
/*
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
*/
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
