
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
                firstRow: 1,
                columns: options.data[0].length,
                rows: options.data.length
            });

            sheet.getRange(rangeString).values = options.data;

            if (options.data.length > 1) {

                if (options.formulas) {

                    for (var i = 0; i < options.formulas.length; i++) {
                
                        var formula = options.formulas[i];
    
                        var formulas = [];
    
                        for (var j = 0; j < options.data.length - 1; j++) {
        
                            formulas.push([formula.formula.replaceAll('?', (j + 2))]);
                        }
    
                        var rangeString = getRangeString({
                            firstColumn: formula.column,
                            firstRow: 2,
                            columns: 1,
                            rows: options.data.length - 1
                        });
    
                        sheet.getRange(rangeString).formulas = formulas;
                    }
                }
    
                if (options.sumColumns) {
    
                    for (var i = 0; i < options.sumColumns.length; i++) {
    
                        var column = options.sumColumns[i];
    
                        var rangeString = getRangeString({
                            firstColumn: column,
                            firstRow: options.data.length + 1,
                            columns: 1,
                            rows: 1
                        });
    
                        var range = sheet.getRange(rangeString);
                        
                        range.formulas = [['=SUM(' + column + '2:' + column + (options.data.length) + ')']];
                        range.format.font.bold = true;
                    }
                }
    
                if (options.formats) {
    
                    for (var i = 0; i < options.formats.length; i++) {
    
                        var format = options.formats[i];
    
                        var formats = [];
    
                        for (var j = 0; j < options.data.length - 1; j++) {
        
                            formats.push([format.format]);
                        }
    
                        for (var j = 0; j < format.columns.length; j++) {
        
                            var rangeString = getRangeString({
                                firstColumn: format.columns[j],
                                firstRow: 2,
                                columns: 1,
                                rows: options.data.length - 1
                            });
    
                            sheet.getRange(rangeString).numberFormat = formats;
                        }
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
