
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

                    var rangeOptions = options.ranges[i];

                    if (!rangeOptions.rows) rangeOptions.rows = (rangeOptions.cells ? 1 : options.data.length - (rangeOptions.firstRow ? rangeOptions.firstRow - 1 : 0));

                    if (rangeOptions.rows == 0) continue;

                    if (rangeOptions.cells) {

                        for (var j = 0; j < rangeOptions.cells.length; j++) {
    
                            var range = sheet.getRange(rangeOptions.cells[j]);
                
                            setRange(range, rangeOptions);
                        }
                    }
                    else if (rangeOptions.columns) {

                        for (var j = 0; j < rangeOptions.columns.length; j++) {

                            var rng = JSON.parse(JSON.stringify(rangeOptions));

                            rng.firstColumn = rangeOptions.columns[j];
                            rng.columns = 1;

                            var range = sheet.getRange(getRangeString(rng));

                            setRange(range, rangeOptions);
                        }
                    }
                    else {

                        rangeOptions.columns = options.data[0].length;

                        var range = sheet.getRange(getRangeString(rangeOptions));

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
    },
    clearData: async function () {
    
        await objExcel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            var range = sheet.getRange();

            range.clear();

            sheet.getRange('A1').select();

            range.format.useStandardWidth = true;
            
            await context.sync();
        });
    },
};

function getRangeString(options) {

    var firstColumn = (options.firstColumn ? options.firstColumn : 'A');
    var firstRow = (options.firstRow ? options.firstRow : 1);

    return (firstColumn + (options.rows ? firstRow : '') + ':' + String.fromCharCode(firstColumn.charCodeAt(0) + options.columns - 1) + (options.rows ? firstRow + options.rows - 1 : ''));
}

function setRange(range, options) {

    if (options.formula) {

        var formulas = [];
    
        for (var i = 0; i < options.rows; i++) {
    
            if (options.cells) {

                formulas.push(['=' + options.formula]);
            }
            else {

                formulas.push(['=' + options.formula.replaceAll('?', (i + options.firstRow))]);
            }
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

    if (options.color) {
        
        if (options.color == 'white')
            range.format.fill.clear();
        else
            range.format.fill.color = options.color;
    }

    if (options.bold) range.format.font.bold = true;
}
