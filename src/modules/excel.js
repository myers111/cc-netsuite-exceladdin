
var objExcel = null;

module.exports = {

    initialize: async function (options) {

        objExcel = options.excel;

        Excel.run(async (context) => {

            context.workbook.worksheets.getActiveWorksheet().name = 'NetSuite';
    
            await context.sync();
        });
    },
    addData: async function (options) {
    
        await objExcel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            sheet.getRange().clear();

            var rangeString = getRangeString({
                rows: options.data.length,
                columns: options.data[0].length
            });

            sheet.getRange(rangeString).values = options.data;

            if (options.ranges) {

                for (var i = 0; i < options.ranges.length; i++) {

                    var range = options.ranges[i];

                    if (!range.rows) range.rows = options.data.length - (range.firstRow ? range.firstRow - 1 : 0);

                    if (range.cell) {

                        setRange(sheet, range);
                    }
                    else if (range.columns) {

                        for (var j = 0; j < range.columns.length; j++) {

                            var rng = JSON.parse(JSON.stringify(range));

                            rng.firstColumn = range.columns[j];
                            rng.columns = 1;

                            setRange(sheet, rng);
                        }
                    }
                    else {

                        range.columns = options.data[0].length;

                        setRange(sheet, range);
                    }
                }
            }

            var rangeString = getRangeString({
                columns: options.data[0].length
            });

            await context.sync();
        });
    },
    clearData: async function () {
    
        await objExcel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            sheet.getRange().clear();

            sheet.getRange().format.useStandardWidth = true;
            
            await context.sync();
        });
    },
};

function getRangeString(options) {

    var firstColumn = (options.firstColumn ? options.firstColumn : 'A');
    var firstRow = (options.firstRow ? options.firstRow : 1);

    return (firstColumn + (options.rows ? firstRow : '') + ':' + String.fromCharCode(firstColumn.charCodeAt(0) + options.columns - 1) + (options.rows ? firstRow + options.rows - 1 : ''));
}

function setRange(sheet, options) {

    var range = null;
    
    if (options.cell)
        range = sheet.getRange(options.cell);
    else
        range = sheet.getRange(getRangeString(options));

    if (options.formula) {

        var formulas = [];
    
        for (var i = 0; i < options.rows; i++) {
    
            formulas.push(['=' + options.formula.replaceAll('?', (i + options.firstRow))]);
        }

        range.formulas = formulas;
    }

    if (options.numberFormat) {

        var numberFormats = [];
    
        for (var i = 0; i < options.rows; i++) {
    
            numberFormats.push([options.numberFormat]);
        }

        range.numberFormat = numberFormats;
    }

    if (options.color) range.format.fill.color = options.color;

    if (options.bold) range.format.font.bold = true;
}
