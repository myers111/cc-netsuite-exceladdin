
var objExcel = null;
var objGroupbyRows = null;

module.exports = {

    initialize: async function (options) {

        objExcel = options.excel;
        objGroupbyRows = options.groupByRows;
    },
    addData: async function (sheetName, options) {
    
        await objExcel.run(async (context) => {

            var sheet = await getSheet(context, sheetName);

            var rangeString = getRangeString({
                rows: options.data.length,
                columns: options.data[0].length
            });

            sheet.getRange(rangeString).values = options.data;

            if (options.ranges) {

                for (var i = 0; i < options.ranges.length; i++) {

                    var rangeOptions = options.ranges[i];

                    for (var j = 0; j < rangeOptions.range.length; j++) {
    
                        var range = sheet.getRange(rangeOptions.range[j]).load(['rowIndex','rowCount','columnCount']);

                        await context.sync();

                        setRange(range, rangeOptions);
                    }
                }
            }

            var rangeString = getRangeString({
                columns: options.data[0].length
            });

            sheet.getRange(rangeString).format.autofitColumns();

            await context.sync();
        });
    }
};

async function clearData(context, sheetName = null) {

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

                await context.sync();
            }
            else {

                sheet = context.workbook.worksheets.add(sheetName);
            }
        }
        else {

            await clearData(context, sheetName);
        }
    }
    else {

        sheet = context.workbook.worksheets.getActiveWorksheet();

        await clearData(context);
    }

    return sheet;
}

function getRangeString(options) {

    var firstColumn = (options.firstColumn ? options.firstColumn : 'A');
    var firstRow = (options.firstRow ? options.firstRow : 1);

    return (firstColumn + (options.rows ? firstRow : '') + ':' + String.fromCharCode(firstColumn.charCodeAt(0) + options.columns - 1) + (options.rows ? firstRow + options.rows - 1 : ''));
}

function setRange(range, options) {

    var formulas = [];
    var numberFormats = [];

    for (var i = 0; i < range.rowCount; i++) {

        var formula = [];
        var numberFormat = [];

        for (var j = 0; j < range.columnCount; j++) {

            if (options.formula) {

                options.formula = options.formula.replaceAll('?', (i + range.rowIndex + 1));

                formula.push('=' + options.formula);
            }

            if (options.numberFormat) {

                numberFormat.push(options.numberFormat);
            }
        }

        if (formula.length) formulas.push(formula);
        if (numberFormat.length) numberFormats.push(numberFormat);
    }

    if (formulas.length) range.formulas = formulas;
    if (numberFormats.length) range.numberFormat = numberFormats;

    if (options.color) {
        
        if (options.color == 'white')
            range.format.fill.clear();
        else
            range.format.fill.color = options.color;
    }

    if (options.horizontalAlignment) range.format.horizontalAlignment = options.horizontalAlignment;
    if (options.bold) range.format.font.bold = true;
    if (options.groupByRows) range.group(Excel.GroupOption.byRows);
    if (options.groupByColumns) range.group(Excel.GroupOption.byColumns);
}