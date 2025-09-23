/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const api = require('../modules/api.js');
const excel = require('../modules/excel.js');

const NEW_ITEM = 3757;
const COLOR_INPUT = '#C6E0B4';
const LABEL_HEADER = ['Quantity','Item','Description','Cost','Ext. Cost','Quote','Ext. Quote'];
const LABEL_HEADER_EX = ['MU%','Discount','Units','Vendor','Manufacturer','MPN'];
const LABEL_LABOR = 'Labor';
const LABEL_ITEMS = 'Items';
const LABEL_EXPENSES = 'Expenses';
const LABEL_TOTAL = 'Total';

/* global console, document, Excel, Office */

var WORKSHEET = {};
var LISTS = null;

var summaryFormulas = null;

Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {

        document.getElementById("addBom").onclick = onAddBom;
        document.getElementById("createRevision").onclick = onCreateRevision;
        document.getElementById("reload").onclick = onReload;
        document.getElementById("save").onclick = onSave;

        // The this jQuery function must be used here since the functions being called use jQuery. Otherwise an error occurs.
            
        $(document).ready(function() {
        
            initList('customer');
            initList('project');
            initList('quote');
        });
    }
});

async function initList(id) {

    var selector = '#' + id + 'List';

    $(selector).focus(function() { loadList(id); });

    $(selector).on('change', function() { onChange(id); });
}

async function loadList(id) {

    var selector = '#' + id + 'List';
    var path = id + 's';

    if ($(selector + ' option').length) return;

    var sel = $(selector);

    var params = {path: path};

    var options = getOptions(id);

    if (options) params['options'] = options;

    var data = await api.get(params);

    if (data.length == 0) return;

    sel.append('<option value=""></option>');

    for (var i = 0; i < data.length; i++) {

        var d = data[i];

        sel.append('<option value="' + d.id + '"' + (d.selected ? ' selected' : '') + '>' + d.name + '</option>');
    }
}

async function emptyList(id) {

    $('#' + id + 'List').empty();
}

function getOptions(id) {

    var options = {};

    switch (id) {
        case 'quote':
            var projectId = $('#projectList').val();
            if (parseInt(projectId)) options['projectId'] = projectId;
        case 'project':
            var customerId = $('#customerList').val();
            if (parseInt(customerId)) options['customerId'] = customerId;
    }

    return (Object.keys(options).length == 0 ? null : options);
}

async function onChange(id) {

    switch (id) {
        case 'customer':
            emptyList("project");
        case 'project':
            emptyList("quote");
        case 'quote':
            await onQuote();
    }
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onQuote() {

    await reset();

    var quoteId = $('#quoteList').val();

    if (quoteId) {

        summaryFormulas = {
            labor: {},
            cost: {},
            quote: {}
        };

        var params = {
            path: 'quote',
            options: {
                id: quoteId
            }
        };

        var data = await api.get(params);
        
        LISTS = {
            defaultLabor: data.defaultLabor,
            units: data.units,
            expAccounts: data.expAccounts
        };

        await initialize();

        const promises = [];

        for (var i = 0; i < data.boms.length; i++) {

            promises.push(addBom({
                bom: data.boms[i],
                units: data.units,
                expAccounts: data.expAccounts
            }));
        }

        await Promise.all(promises);

        await addSummary(data);
    }

    document.getElementById("controls").style.display = (quoteId > 0 ? '' : 'none');
}

async function reset() {

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

async function initialize() {

    // G2 must be set to a value since it is referenced in formulas on the BOM's which are created before the Summary
    // BOM formulas referencing Summary!$G$2 won't work if this isn't done

    await Excel.run(async (context) => {

        var sheet = await excel.getSheet(context, 'Summary');

        await excel.setSheet(context, sheet, {
            ranges: [
                {
                    range: ['G2'],
                    values: [[0]]
                }
            ]
        });
    });
}

async function addSummary(data) {

    var dataArray = [
        ['Quote','','','','','',0],
        ['MU (Default)','','','','','',(data.defaultMU / 100)],
        ['GM','','','','','',0],
        ['MU','','','','Cost','Quote',0],
        [LABEL_ITEMS,'','','',0,0,0],
        [LABEL_LABOR,'','','',0,0,0],
        ['','','','','','',''],
        LABEL_HEADER
    ];

    for (var i = 0; i < dataArray.length; i++) {

        dataArray[i] = dataArray[i].concat(['','','']); // Add spaces for hidden columns
    }

    var dataRanges = [];

    var laborData = getLaborData({
        boms: data.boms,
        rowFirst: dataArray.length + 1,
        isSummary: true
    });

    dataArray = dataArray.concat(laborData.values);
    dataRanges = dataRanges.concat(laborData.ranges);

    var itemData = getItemData({
        items: data.items,
        rowFirst: dataArray.length + 1,
        isSummary: true
    });

    dataArray = dataArray.concat(itemData.values);
    dataRanges = dataRanges.concat(itemData.ranges);

    dataArray.push([LABEL_TOTAL,'','','',0,'',0]);

    dataArray[dataArray.length - 1] = dataArray[dataArray.length - 1].concat(['','','']); // Add spaces for hidden columns

    dataRanges = dataRanges.concat([
        {
            range: ['G2'],
            color: COLOR_INPUT
        },
        {
            range: ['G1'],
            formula: '$G$' + dataArray.length,
            numberFormat: '$#,###.00'
        },
        {
            range: ['E5:F6','E' + dataArray.length, 'G' + dataArray.length],
            numberFormat: '$#,###.00'
        },
        {
            range: ['A1:G8','A' + dataArray.length + ':G' + dataArray.length],
            bold: true
        },
        {
            range: ['G2:G6'],
            numberFormat: '#,###.00%'
        },
        {
            range: ['G3'],
            formula: '($G$1-$E$' + dataArray.length + ')/IF($G$1>0,$G$1,1)'
        },
        {
            range: ['G4'],
            formula: '($G$1-$E$' + dataArray.length + ')/IF($E$' + dataArray.length + '>0,$E$' + dataArray.length + ',1)'
        },
        {
            range: ['E5'],
            formula: 'E' + dataArray.length + '-E6'
        },
        {
            range: ['F5'],
            formula: 'G' + dataArray.length + '-F6'
        },
        {
            range: ['G5'],
            formula: ('IF($E$' + dataArray.length + '=0,0,$E$5/$E$' + dataArray.length + ')')
        },
        {
            range: ['E6'],
            formula: 'SUMIFS(E' + (laborData.rowFirst + 1) + ':E' + laborData.rowLast + ',D' + (laborData.rowFirst + 1) + ':D' + laborData.rowLast + ',"<>")'
        },
        {
            range: ['F6'],
            formula: 'SUMIFS(G' + (laborData.rowFirst + 1) + ':G' + laborData.rowLast + ',F' + (laborData.rowFirst + 1) + ':F' + laborData.rowLast + ',"<>")'
        },
        {
            range: ['G6'],
            formula: ('IF($E$' + dataArray.length + '=0,0,$E$6/$E$' + dataArray.length + ')')
        },
        {
            range:['H:J'],
            hideColumns: true
        }
    ]);

    dataRanges = dataRanges.concat(getTotalRowRanges(true, {
        row: dataArray.length,
        itemRowFirst: itemData.rowFirst,
        itemRowLast: itemData.rowLast
    }));

    await Excel.run(async (context) => {

        var sheet = await excel.getSheet(context, "Summary");

        await excel.setSheet(context, sheet, {
            values: dataArray,
            ranges: dataRanges,
            autofitColumns: 7
        });

        await sheet.onChanged.add(onWorksheetChange);

        await context.sync();
    });
}

async function addBom(data) {

    var dataArray = [
        LABEL_HEADER.concat(LABEL_HEADER_EX).concat(['']).concat(['','',''])
    ];

    var dataRanges = [];

    // Items

    var itemData = getItemData({
        items: data.bom.items,
        units: data.units,
        rowFirst: dataArray.length + 1,
        isSummary: false
    });

    dataArray = dataArray.concat(itemData.values);
    dataRanges = dataRanges.concat(itemData.ranges);

    // Labor

    var laborData = getLaborData({
        boms: [data.bom],
        rowFirst: dataArray.length + 1,
        isSummary: false
    });

    dataArray = dataArray.concat(laborData.values);
    dataRanges = dataRanges.concat(laborData.ranges);

    // Expenses

    var expenseData = getExpenseData({
        expenses: data.bom.expenses,
        rowFirst: dataArray.length + 1
    });

    dataArray = dataArray.concat(expenseData.values);
    dataRanges = dataRanges.concat(expenseData.ranges);

    // Totals

    dataArray.push([LABEL_TOTAL,'','','',0,'',0,'','','','','','']);

    dataArray[dataArray.length - 1] = dataArray[dataArray.length - 1].concat(['']); // Add space for group control
    dataArray[dataArray.length - 1] = dataArray[dataArray.length - 1].concat(['','','']); // Add spaces for hidden columns

    // Set summaryFormulas

    summaryFormulas.cost[data.bom.name] = "E" + dataArray.length;
    summaryFormulas.quote[data.bom.name] = "G" + dataArray.length;

    dataRanges = dataRanges.concat([
        {
            range: ['A1:M1','A' + dataArray.length + ':M' + dataArray.length],
            bold: true
        },
        {
            range:['H:M'],
            groupByColumns: true
        },
        {
            range:['O:Q'],
            hideColumns: true
        },
        {
            range:['N:N'],
            columnWidth: 15
        }
    ]);

    dataRanges = dataRanges.concat(getTotalRowRanges(false, {
        row: dataArray.length,
        itemRowFirst: itemData.rowFirst,
        itemRowLast: itemData.rowLast,
        laborRowFirst: laborData.rowFirst,
        laborRowLast: laborData.rowLast,
        expenseRowFirst: expenseData.rowFirst,
        expenseRowLast: expenseData.rowLast
    }));

    await Excel.run(async (context) => {

        var sheet = await excel.getSheet(context, data.bom.name);

        WORKSHEET[sheet.id.toString()] = {bomId: data.bom.id};

        await excel.setSheet(context, sheet, {
            values: dataArray,
            ranges: dataRanges,
            autofitColumns: 13
        });

        await sheet.onChanged.add(onWorksheetChange);

        await context.sync();
    });
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getItemData(data) {

    /* data = { items, rowFirst, isSummary } *******************************************************************************************************************/
    
    var itemData = {
        values: [
            [LABEL_ITEMS,'','','','','','']
        ],
        ranges: [
            {
                range: ['A' + data.rowFirst],
                bold: true
            }
        ],
        rowFirst: data.rowFirst,
        rowLast: 0
    };

    if (!data.isSummary) {

        itemData.values[0] = itemData.values[0].concat(['','','','','','']); // Add spaces for extra BOM columns
        itemData.values[0] = itemData.values[0].concat(['']); // Add space for group control
    }

    itemData.values[0] = itemData.values[0].concat(['','','']); // Add spaces for hidden columns

    // Set values

    for (var i = 0; i < data.items.length; i++) {
    
        var item = data.items[i];

        var quantity = (item.quantity ? parseInt(item.quantity) : 0);
        var price = (item.price ? parseFloat(item.price) : 0);

        itemData.values.push([
            quantity,
            (item.id == NEW_ITEM ? item.newItem : item.name),
            item.description,
            price,
            0,
            0,
            0
        ]);

        var itemDataValuesLength = itemData.values.length;

        var type = {
            isSummary: data.isSummary,
            isItem: true
        };

        if (data.isSummary) {

            if (item.bomId > 0) type['bomName'] = item.description;
        }
        else {

            const unit = data.units.find((obj) => obj.type === item.unitsType);

            if (unit) type['unitNames'] = unit.names;

            itemData.values[itemDataValuesLength - 1] = itemData.values[itemDataValuesLength - 1].concat([
                (item.markUp > 0 ? item.markUp : ''),
                (item.discount == 'T' ? 'Yes' : 'No'),
                item.units,
                (item.vendorId ? item.vendor : item.newVendor),
                item.manufacturer,
                item.mpn
            ]);
        }

        if (!data.isSummary) itemData.values[itemDataValuesLength - 1].push(''); // Add space for group control
        itemData.values[itemDataValuesLength - 1].push(item.key); // Add key for hidden column
        itemData.values[itemDataValuesLength - 1].push(item.id); // Add item ID for hidden column
        itemData.values[itemDataValuesLength - 1].push(data.isSummary ? item.bomId  : 0); // Add bom ID for hidden column

        itemData.ranges = itemData.ranges.concat(getRowRanges(data.rowFirst + 1 + i, type));
    }

    itemData.rowLast = itemData.rowFirst + data.items.length;

    return itemData;
}

function getLaborData(data) {

    /* data = { boms, rowFirst, isSummary } *******************************************************************************************************************/
    
    var laborData = {
        values: [
            [LABEL_LABOR,'','','','','','']
        ],
        ranges: [
            {
                range: ['A' + data.rowFirst],
                bold: true
            }
        ],
        rowFirst: data.rowFirst,
        rowLast: 0
    };

    if (!data.isSummary) {
        
        laborData.values[0] = laborData.values[0].concat(['']); // Add space for group control
        laborData.values[0] = laborData.values[0].concat(['','','','','','']); // Add spaces for extra columns
    }

    laborData.values[0] = laborData.values[0].concat(['','','']); // Add spaces for hidden columns

    var type = {
        isSummary: data.isSummary,
        isLabor: true
    };

    // Create labor object

    var objLabor = {};
    
    for (var i = 0; i < data.boms.length; i++) {

        for (var j = 0; j < data.boms[i].labor.length; j++) {

            var labor = data.boms[i].labor[j];

            if (!objLabor[labor.sgName]) objLabor[labor.sgName] = [];

            var index = objLabor[labor.sgName].findIndex(obj => obj.id == labor.id && obj.price == labor.price); // See if labor role already exists

            if (index < 0) objLabor[labor.sgName].push(labor);
        }
    }

    // Set values

    for (const key in objLabor) {

        if (!objLabor.hasOwnProperty(key)) continue;

        var groupArray = [0,key,'','',0,'',0];

        if (!data.isSummary) {
            
            groupArray = groupArray.concat(['','','','','','']); // Add spaces for extra columns
            groupArray = groupArray.concat(['']); // Add space for group control
        }

        groupArray = groupArray.concat(['','','']); // Add spaces for hidden columns

        laborData.values.push(groupArray);

        for (var i = 0; i < objLabor[key].length; i++) {

            var labor = objLabor[key][i];

            laborData.values.push([
                labor.quantity,
                '',
                labor.name,
                labor.price,
                0,
                0,
                0
            ]);

            var laborDataValuesLength = laborData.values.length;

            var row = laborData.rowFirst + laborDataValuesLength - 1;

            var idString = labor.id.toString();

            if (data.isSummary) {

                laborData.values[laborDataValuesLength - 1] = laborData.values[laborDataValuesLength - 1].concat(['','','']); // Add spaces for hidden columns

                type['summaryFormulas'] = summaryFormulas.labor[key][idString];
            }
            else {

                laborData.values[laborDataValuesLength - 1] = laborData.values[laborDataValuesLength - 1].concat([
                    (labor.markUp > 0 ? labor.markUp : ''),
                    (labor.discount == 'T' ? 'Yes' : 'No'),
                    '',
                    '',
                    '',
                    ''
                ]);

                laborData.values[laborDataValuesLength - 1].push(''); // Add space for group control
                laborData.values[laborDataValuesLength - 1].push(labor.key); // Add key to hidden column
                laborData.values[laborDataValuesLength - 1].push(labor.id); // Add item ID to hidden column
                laborData.values[laborDataValuesLength - 1].push(labor.sgId); // Add group ID to hidden column

                // Formulas
                
                if (!summaryFormulas.labor[key]) summaryFormulas.labor[key] = {};

                if (summaryFormulas.labor[key][idString]) {

                    summaryFormulas.labor[key][idString].qty += '+';
                    summaryFormulas.labor[key][idString].cost += '+';
                    summaryFormulas.labor[key][idString].quote += '+';
                }
                else {

                    summaryFormulas.labor[key][idString] = {
                        qty: '',
                        cost: '',
                        quote: ''
                    };
                }

                summaryFormulas.labor[key][idString].qty += ("'" + data.boms[0].name + "'!A" + row);
                summaryFormulas.labor[key][idString].cost += ("'" + data.boms[0].name + "'!D" + row);
                summaryFormulas.labor[key][idString].quote += ("'" + data.boms[0].name + "'!F" + row);
            }

            laborData.ranges = laborData.ranges.concat(getRowRanges(data.rowFirst + laborDataValuesLength - 1, type));
        }
    }

    laborData.rowLast = laborData.rowFirst + laborData.values.length - 1;

    // Set labor formulas & grouping

    type['rowSum'] = 0;

    for (var i = laborData.rowFirst + 1; i <= laborData.rowLast; i++) {

        if (laborData.values[i - laborData.rowFirst][3] == '') {

            if (type.rowSum != 0) laborData.ranges = laborData.ranges.concat(getGroupRowRanges(i, type));

            type.rowSum = i;
        }
    }

    if (type.rowSum != 0) laborData.ranges = laborData.ranges.concat(getGroupRowRanges(laborData.rowLast + 1, type));

    return laborData;
}

function getExpenseData(data) {

    /* data = { expenses, rowFirst } *******************************************************************************************************************/
    
    var expenseData = {
        values: [
            [LABEL_EXPENSES,'','','','','','','','','','','','']
        ],
        ranges: [
            {
                range: ['A' + data.rowFirst],
                bold: true
            }
        ],
        rowFirst: data.rowFirst,
        rowLast: 0
    };

    expenseData.values[0] = expenseData.values[0].concat(['']); // Add space for group control
    expenseData.values[0] = expenseData.values[0].concat(['','','']); // Add spaces for hidden columns

    // Set values

    for (var i = 0; i < data.expenses.length; i++) {
    
        var expense = data.expenses[i];

        expenseData.values.push([
            expense.quantity,
            '',
            expense.name,
            expense.price,
            0,
            0,
            0,
            (expense.markUp > 0 ? expense.markUp : ''),
            (expense.discount == 'T' ? 'Yes' : 'No'),
            '',
            '',
            '',
            '',
            '', // Add space for group control
            expense.key, // Add key for hidden column
            expense.accountId, // Add account ID for hidden column
            ''
        ]);   

        expenseData.ranges = expenseData.ranges.concat(getRowRanges(data.rowFirst + 1 + i, {
            isExpense: true
        }));
    }

    expenseData.rowLast = expenseData.rowFirst + data.expenses.length;

    return expenseData;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onWorksheetChange(eventArgs) {

    await Excel.run(async (context) => {

        if (eventArgs.changeType === Excel.DataChangeType.rowInserted) {

            var rows = eventArgs.address.split(':'); // This creates a two element array of the first and last rows inserted in string format

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            var range = sheet.getUsedRange();

            range.load("values");

            await context.sync();

            var rowFirst = parseInt(rows[0]);
            var rowLast = parseInt(rows[1]);
            
            var type = getType(range.values, {
                rowFirst: rowFirst,
                rowLast: rowLast
            });

            type['isInsert'] = true;

            var ranges = [];

            for (var i = rowFirst; i <= rowLast; i++) {

                if (type.isItem) {

                    if (type.isSummary) {

                        ranges = ranges.concat([{
                            range: ['A' + i + ':G' + i],
                            values: [[1,'','',0,0,0,0]]
                        }]);
                    }
                    else {

                        ranges = ranges.concat([{
                            range: ['A' + i + ':J' + i],
                            values: [[1,'','',0,0,0,0,'','No','Ea']]
                        }]);
                    }
                }
                else if (type.isExpense) {

                    ranges = ranges.concat([{
                        range: ['A' + i + ':I' + i],
                        values: [[1,'','',0,0,0,0,'','No']]
                    }]);
                }

                var rowRanges = getRowRanges(i, type);

                if (rowRanges.length) ranges = ranges.concat(rowRanges);
            }

            ranges = ranges.concat(getTotalRowRanges(type.isSummary));

            await excel.setSheet(context, sheet, {
                ranges: ranges
            });

            await context.sync();
        }
    });
}

function getType(values, options) {

    var type = {
        isSummary: (values[0][0] == "Quote"),
        isItem: false,
        isLabor: false,
        isExpense: false
    };

    if (type.isSummary) { // Is Summary

        type.isItem = isItem(values, {
            isSummary: true,
            rowFirst: options.rowFirst,
            rowLast: options.rowLast
        });
    }
    else if (values[0][0] == "Quantity") { // Is BOM

        type.isItem = isItem(values, {
            isSummary: false,
            rowFirst: options.rowFirst,
            rowLast: options.rowLast
        });
        
        if (!type.isItem) {

            type.isExpense = isExpense(values, {
                rowFirst: options.rowFirst,
                rowLast: options.rowLast
            });
        }
    }

    return  type;
}

function isItem(values, params) {

    for (var i = params.rowFirst - 1; i > 0; i--) {

        if (!isNaN(values[i-1][0])) continue;

        if (values[i-1][0] != 'Items') return false;

        break;
    }

    for (var i = params.rowLast + 1; i < values.length; i++) {

        if (!isNaN(values[i-1][0])) continue;

        if (values[i-1][0] != (params.isSummary ? 'Total' : 'Labor')) return false;

        break;
    }

    return true;
}

function isExpense(values, params) {

    for (var i = params.rowFirst - 1; i > 0; i--) {

        if (!isNaN(values[i-1][0])) continue;

        if (values[i-1][0] != 'Expenses') return false;

        break;
    }

    for (var i = params.rowLast + 1; i < values.length; i++) {

        if (!isNaN(values[i-1][0])) continue;

        if (values[i-1][0] != 'Total') return false;

        break;
    }

    return true;
}

function getRowRanges(row, type) {

    // Range formats common to all types (with summary conditions)

    var ranges = [
        {
            range: ['A' + row],
            color: (type.isSummary && type.isLabor ? '' : COLOR_INPUT),
            horizontalAlignment: 'center',
            bold: false
        },
        {
            range: ['D' + row + ':G' + row],
            numberFormat: '$#,###.00'
        },
        {
            range: ['E' + row],
            formula: 'A?*D?'
        },
        {
            range: ['F' + row],
            formula: (type.isSummary ? '' : 'D?*(1+IF(I?="Yes",-1,1)*IF(ISNUMBER(H?),H?,Summary!$G$2))')
        },
        {
            range: ['G' + row],
            formula: 'A?*F?'
        }
    ];

    // Range formats common to BOM's

    if (!type.isSummary) {

        ranges = ranges.concat([
            {
                range: ['D' + row],
                color: COLOR_INPUT
            },
            {
                range: ['H' + row + ':I' + row],
                color: COLOR_INPUT,
                horizontalAlignment: 'center'
            },
            {
                range: ['H' + row],
                numberFormat: '#,###.00%'
            },
            {
                range: ['I' + row],
                dataValidationRule: {
                    list: {
                        inCellDropDown: true,
                        source: "Yes,No"
                    }
                }
            }
        ]);
    }
    
    // Range formats specific to each section

    if (type.isItem) {

        ranges = ranges.concat([
            {
                range: ['B' + row], // Set column B to text format to preserve leading zero's in item names
                data: [[0]]
            }
        ]);

        if (type.isSummary) {

            if (type.bomName) {

                ranges = ranges.concat([
                    {
                        range: ['C' + row],
                        formula: 'TEXTAFTER(CELL("filename",\'' + type.bomName + '\'!A1),"]")'

                    },
                    {
                        range: ['D' + row],
                        formula: "'" + type.bomName + "'!" + summaryFormulas.cost[type.bomName]
                    },
                    {
                        range: ['F' + row],
                        formula: "'" + type.bomName + "'!" + summaryFormulas.quote[type.bomName]
                    }
                ]);
            }
            else {

                ranges = ranges.concat([
                    {
                        range: ['B' + row + ':D' + row],
                        color: COLOR_INPUT
                    }
                ]);
            }
        }
        else {

            if (type.isInsert) {

                ranges = ranges.concat([
                    {
                        range: ['B' + row + ':C' + row],
                        color: COLOR_INPUT
                    }
                ]);
            }

            ranges = ranges.concat([
                {
                    range: ['K' + row + ':M' + row],
                    color: COLOR_INPUT
                },
                {
                    range: ['J' + row],
                    color: COLOR_INPUT,
                    dataValidationRule: {
                        list: {
                            inCellDropDown: true,
                            source: (type.unitNames ? type.unitNames : LISTS.units.filter(unit => unit.type != 3).map(unit => unit.names).join(',')) // Filter out labor units
                        }
                    }
                },
            ]);
        }
    }
    else if (type.isLabor) {

        if (type.isSummary) {

            ranges = ranges.concat([
                {
                    range: ['A' + row],
                    formula: type.summaryFormulas.qty
                },
                {
                    range: ['D' + row],
                    formula: type.summaryFormulas.cost
                },
                {
                    range: ['F' + row],
                    formula: type.summaryFormulas.quote
                }
            ]);
        }
    }
    else if (type.isExpense) {

        if (type.isInsert) {

            ranges = ranges.concat([
                {
                    range: ['C' + row],
                    color: COLOR_INPUT
                },
                {
                    range: ['C' + row],
                    dataValidationRule: {
                        list: {
                            inCellDropDown: true,
                            source: LISTS.expAccounts.join(',')
                        }
                    }
                }
            ]);
        }
    }

    return ranges;
}

function getGroupRowRanges(row, type) {

    var ranges = [];

    if (type.isLabor) {

        if (type.rowSum) {

            ranges = ranges.concat([
                {
                    range: ['A' + type.rowSum],
                    formula: ('SUM(A' + (type.rowSum + 1) + ':A' + (row - 1) + ')'),
                    bold: true,
                    horizontalAlignment: 'center'
                },
                {
                    range: ['B' + type.rowSum],
                    bold: true
                },
                {
                    range: ['E' + type.rowSum],
                    formula: ('SUM(E' + (type.rowSum + 1) + ':E' + (row - 1) + ')'),
                    bold: true,
                    numberFormat: '$#,###.00'
                },
                {
                    range: ['G' + type.rowSum],
                    formula: ('SUM(G' + (type.rowSum + 1) + ':G' + (row - 1) + ')'),
                    bold: true,
                    numberFormat: '$#,###.00'
                },
                {
                    range: [(type.rowSum + 1) + ':' + (row - 1)],
                    groupByRows: true
                }
            ]);
        }
    }

    return ranges;
}

function getTotalRowRanges(isSummary, rows = null) {

    var ranges = [];

    if (isSummary) {

        if (!rows) {

            rows = {
                
            };
        }

        ranges = [
            {
                range: ['E' + rows.row],
                formula: 'SUM(E' + (rows.itemRowFirst + 1) + ':E' + rows.itemRowLast + ')'
            },
            {
                range: ['G' + rows.row],
                formula: 'SUM(G' + (rows.itemRowFirst + 1) + ':G' + rows.itemRowLast + ')'
            },
        ];
    }
    else  {

        if (!rows) {

            rows = {
                
            };
        }

        ranges = [
            {
                range: ['E' + rows.row],
                formula: 'SUM(E' + (rows.itemRowFirst + 1) + ':E' + rows.itemRowLast + ')+SUMIFS(E' + (rows.laborRowFirst + 1) + ':E' + rows.laborRowLast + ',D' + (rows.laborRowFirst + 1) + ':D' + rows.laborRowLast + ',"<>")+SUM(E' + (rows.expenseRowFirst + 1) + ':E' + rows.expenseRowLast + ')',
            },
            {
                range: ['G' + rows.row],
                formula: 'SUM(G' + (rows.itemRowFirst + 1) + ':G' + rows.itemRowLast + ')+SUMIFS(G' + (rows.laborRowFirst + 1) + ':G' + rows.laborRowLast + ',F' + (rows.laborRowFirst + 1) + ':F' + rows.laborRowLast + ',"<>")+SUM(G' + (rows.expenseRowFirst + 1) + ':G' + rows.expenseRowLast + ')',
            }
        ];
    }

    return ranges;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onAddBom() {

    addBom({
        bom: {
            id: 0,
            name: 'NEW BOM',
            items: [],
            expenses: [],
            labor: LISTS.defaultLabor
        },
        units: LISTS.units,
        expAccounts: LISTS.expAccounts
    });
}

async function onCreateRevision() {

    onSave();
}

async function onReload() {

    onQuote();
}

async function onSave() {

    var quoteId = $('#quoteList').val();

    if (!quoteId) return;

    try {

        await Excel.run(async (context) => {

            var data = {
                id: quoteId,
                defaultMU: 0,
                items: [],
                boms: []
            };

            var sheets = context.workbook.worksheets;
        
            sheets.load("items/name");

            await context.sync();

            for (const sheet of sheets.items) {

                var range = sheet.getUsedRange();

                range.load('values,formulas');

                await context.sync();

                if (sheet.name == 'Summary')
                    data.defaultMU = range.values[1][6];
                else if (!range.values.length)
                    return;
                else if (range.values[0][0] != 'Quantity')
                    return;
                else
                    data.boms.push({id: WORKSHEET[sheet.id.toString()].bomId, name: sheet.name, items: [], expenses: []});

                var section = '';
     
                for (var i = 0; i < range.values.length; i++) {
                
                    var values = range.values[i];
                    var formulas = range.formulas[i];

                    switch (values[0]) {
                        case LABEL_ITEMS:
                        case LABEL_LABOR:
                        case LABEL_EXPENSES:
                            if ((sheet.name == 'Summary' && i >= 9) || sheet.name != 'Summary') section = values[0];
                            continue;
                        case LABEL_TOTAL:
                            section = '';
                            continue;
                        default:
                            break;
                    }

                    if (section == LABEL_ITEMS) {

                        if (sheet.name == 'Summary') {

                            var item = {
                                quantity: values[0],
                                description: values[2]
                            };

                            if (formulas[3].substring(0,1) != '=') item['price'] = values[3];

                            data.items.push(item);
                        }
                        else if (values[0] > 0) {

                            data.boms[data.boms.length - 1].items.push({
                                key: values[14],
                                id: values[15],
                                bomId: values[16],
                                quantity: values[0],
                                name: values[1],
                                description: values[2],
                                price: values[3],
                                markUp: values[7],
                                discount: values[8],
                                units: values[9],
                                vendor: values[10],
                                manufacturer: values[11],
                                mpn: values[12]
                            });
                        }
                    }
                    else if (section == LABEL_LABOR) {

                        if (values[0] > 0 && values[1].length == 0) {

                            data.boms[data.boms.length - 1].items.push({
                                key: values[14],
                                id: values[15],
                                sgId: values[16],
                                quantity: values[0],
                                price: values[3],
                                markUp: values[7],
                                discount: values[8]
                            });
                        }
                    }
                    else if (section == LABEL_EXPENSES) {

                        if (values[0] > 0 && values[2].length > 0) {

                            data.boms[data.boms.length - 1].expenses.push({
                                key: values[14],
                                accountId:values[15],
                                account:values[2],
                                quantity: values[0],
                                price: values[3],
                                markUp: values[7],
                                discount: values[8]
                            });
                        }
                    }
                }
            }

            var params = {
                path: 'quote',
                options: {data: data}
            };

            await api.post(params);

            onReload();
        });
    }
    catch (error) {

        console.error(error);
    }
}
