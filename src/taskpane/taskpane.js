/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const api = require('../modules/api.js');
const excel = require('../modules/excel.js');

const NEW_ITEM = 3757;
const MILESTONE_PAYMENT = 34;
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
var BOM_NAME = '';

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

        await addSummary(data);

        const promises = [];

        for (var i = 0; i < data.boms.length; i++) {

            promises.push(addBom({
                bom: data.boms[i],
                units: data.units,
                expAccounts: data.expAccounts
            }));
        }

        await Promise.all(promises);
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

async function addSummary(data) {

    // The data is 7 columns (add 3 spaces for hidden columns)

    var values = [
        ['Quote','','','','','',0,'','',''],
        ['MU (Default)','','','','','',(data.defaultMU / 100),'','',''],
        ['GM','','','','','',0,'','',''],
        ['MU','','','','Cost','Quote',0,'','',''],
        [LABEL_ITEMS,'','','',0,0,0,'','',''],
        [LABEL_LABOR,'','','',0,0,0,'','',''],
        ['','','','','','','','','',''],
        LABEL_HEADER.concat(['','','']),
        [LABEL_LABOR,'','','','','','','','',''],
        [LABEL_ITEMS,'','','','','','','','',''],
        [LABEL_TOTAL,'','','',0,'',0,'','','']
    ];

    // Add sections in reverse in order to preserve the starting index location for insertion.

    // Items

    var itemValues = getItemValues({
        items: data.items,
        isSummary: true
    });

    if (itemValues.length) values.splice(10, 0, ...itemValues);

    // Labor

    var laborValues = getLaborValues({
        boms: data.boms,
        isSummary: true
    });

    if (laborValues.length) values.splice(9, 0, ...laborValues);

    // Sheet Info

    var sheetInfo = getSheetInfo(values);

    // Ranges

    var ranges = [];

    for (var i = 1; i <= values.length; i++) {

        ranges = ranges.concat(getRanges(i + 1, sheetInfo, false));
    }

    ranges = ranges.concat([
        {
            range: ['G2'],
            color: COLOR_INPUT
        },
        {
            range: ['G1'],
            formula: '$G$' + values.length,
            numberFormat: '$#,###.00'
        },
        {
            range: ['E5:F6','E' + values.length, 'G' + values.length],
            numberFormat: '$#,###.00'
        },
        {
            range: ['A1:G8','A' + values.length + ':G' + values.length,'A' + (sheetInfo.labor.first - 1),'A' + (sheetInfo.item.first - 1)],
            bold: true
        },
        {
            range: ['G2:G6'],
            numberFormat: '#,###.00%'
        },
        {
            range: ['G3'],
            formula: '($G$1-$E$' + values.length + ')/IF($G$1>0,$G$1,1)'
        },
        {
            range: ['G4'],
            formula: '($G$1-$E$' + values.length + ')/IF($E$' + values.length + '>0,$E$' + values.length + ',1)'
        },
        {
            range: ['E5'],
            formula: 'E' + values.length + '-E6'
        },
        {
            range: ['F5'],
            formula: 'G' + values.length + '-F6'
        },
        {
            range: ['G5'],
            formula: ('IF($E$' + values.length + '=0,0,$E$5/$E$' + values.length + ')')
        },
        {
            range: ['E6'],
            formula: 'SUMIFS(E' + (sheetInfo.labor.first + 1) + ':E' + sheetInfo.labor.last + ',D' + (sheetInfo.labor.first + 1) + ':D' + sheetInfo.labor.last + ',"<>")'
        },
        {
            range: ['F6'],
            formula: 'SUMIFS(G' + (sheetInfo.labor.first + 1) + ':G' + sheetInfo.labor.last + ',F' + (sheetInfo.labor.first + 1) + ':F' + sheetInfo.labor.last + ',"<>")'
        },
        {
            range: ['G6'],
            formula: ('IF($E$' + values.length + '=0,0,$E$6/$E$' + values.length + ')')
        },
        {
            range:['H:J'],
            hideColumns: true
        }
    ]);

    ranges = ranges.concat(getTotalRowRanges(values.length, sheetInfo));

    await Excel.run(async (context) => {

        var sheet = await excel.getEmptySheet(context, "Summary");

        await excel.setSheet(context, sheet, {
            values: values,
            ranges: ranges,
            autofitColumns: 7
        });

        await sheet.onChanged.add(onWorksheetChange);

        await context.sync();
    });
}

async function addBom(data) {

    // The data is 13 columns (add 1 space for group control and 3 spaces for hidden columns)

    var values = [
        LABEL_HEADER.concat(LABEL_HEADER_EX).concat(['','','','']),
        [LABEL_ITEMS,'','','','','','','','','','','','','','','',''],
        [LABEL_LABOR,'','','','','','','','','','','','','','','',''],
        [LABEL_EXPENSES,'','','','','','','','','','','','','','','',''],
        [LABEL_TOTAL,'','','',0,'',0,'','','','','','','','','','']
    ];

    // Add sections in reverse in order to preserve the starting index location for insertion.

    // Expenses

    var expenseValues = getExpenseValues({
        expenses: data.bom.expenses,
        isSummary: false
    });

    if (expenseValues.length) values.splice(4, 0, ...expenseValues);

    // Labor

    var laborValues = getLaborValues({
        boms: [data.bom],
        isSummary: false
    });

    if (laborValues.length) values.splice(3, 0, ...laborValues);

    // Items

    var itemValues = getItemValues({
        items: data.bom.items,
        units: data.units,
        isSummary: false
    });

    if (itemValues.length) values.splice(2, 0, ...itemValues);

    // Sheet Info

    var sheetInfo = getSheetInfo(values);

    // Ranges

    var ranges = [];

    for (var i = 1; i <= values.length; i++) {

        ranges = ranges.concat(getRanges(i + 1, sheetInfo, false));
    }

    ranges = ranges.concat([
        {
            range: ['A1:M1','A' + values.length + ':G' + values.length,'A' + (sheetInfo.item.first - 1),'A' + (sheetInfo.labor.first - 1),'A' + (sheetInfo.expense.first - 1)],
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

    ranges = ranges.concat(getTotalRowRanges(values.length, sheetInfo));

    await Excel.run(async (context) => {

        var sheet = await excel.getEmptySheet(context, data.bom.name);

        WORKSHEET[sheet.id.toString()] = {bomId: data.bom.id};

        await excel.setSheet(context, sheet, {
            values: values,
            ranges: ranges,
            autofitColumns: 13
        });

        await sheet.onChanged.add(onWorksheetChange);

        await context.sync();
    });

    if (data.bom.id == 0) { // This is a new BOM so add the line to the Summary

        await addBomItemToSummary({
            name: data.bom.name
        });
    }

    await addBomLinksToSummary(laborValues, {
        bomName: data.bom.name,
        rowLabor: (3 + itemValues.length + 1),
        rowTotal: values.length
    });
}

async function addBomItemToSummary(params) {tfvb54

    await Excel.run(async (context) => {

        var sheet = await excel.getSheet(context, 'Summary');

        var range = sheet.getRange('A:A').find("Total", { completeMatch: true }); 

        range.load("rowIndex"); // Zero based

        await context.sync();

        // Insert a row

        BOM_NAME = params.name;

        var row = range.rowIndex + 1;

        sheet.getRange(row + ":" + row).insert(Excel.InsertShiftDirection.down);

        await context.sync();
    });
}
async function addBomLinksToSummary(laborValues, params) {

    await Excel.run(async (context) => {

        var laborItems = [];

        for (var i = 0; i < laborValues.length; i++) {

            var laborValue = laborValues[i];

            if (laborValue[1] == '') { // Labor item

                laborItems.push({
                    itemId: laborValue[15],
                    sgId: laborValue[16], // Service Group
                    row: params.rowLabor + i
                });
            }
        }

        var sheet = await excel.getSheet(context, 'Summary');

        /* Read all the values in formulas on this sheet at once.
         * We can read through an array in JavaScript much faster than having to sync constantly.
        */

        var range = sheet.getUsedRange();

        range.load("formulas"); // This gets the values as well when there isn't a formula

        await context.sync();

        var ranges = [];

        var section = LABEL_LABOR; // Starts at row 11 (10 for 0 based)

        for (var i = 10; i < range.formulas.length - 1; i++) { // Start at first labor value and end at last item

            var formulas = range.formulas[i];

            if (section == LABEL_LABOR) {

                if (formulas[0] == LABEL_ITEMS) {

                    section = LABEL_ITEMS;
                }
                else if (formulas[1] == '') { // Labor item

                    const laborItem = laborItems.find((li) => li.itemId === formulas[8] && li.sgId === formulas[9]);

                    if (laborItem) {

                        ranges = ranges.concat([
                            {
                                range: ['A' + (i + 1)],
                                formula: (Number.isInteger(formulas[0]) ? '' : formulas[0].substring(1) + "+") + "'" + params.bomName + "'!A" + laborItem.row
                            },
                            {
                                range: ['F' + (i + 1)],
                                formula: (Number.isInteger(formulas[5]) ? '' : formulas[5].substring(1) + "+") + "'" + params.bomName + "'!F" + laborItem.row
                            }
                        ]);
                    }
                }
            }
            else { // LABEL_ITEMS

                ranges = ranges.concat([
                    {
                        range: ['D' + (i + 1)],
                        formula: "'" + params.bomName + "'!E" + params.rowTotal
                    },
                    {
                        range: ['F' + (i + 1)],
                        formula: "'" + params.bomName + "'!G" + params.rowTotal
                    },
                ]);
            }
        }

        await excel.setSheet(context, sheet, {
            ranges: ranges
        });
    });
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getItemValues(data) {

    var values = [];

    for (var i = 0; i < data.items.length; i++) {

        var item = data.items[i];

        var quantity = (item.quantity ? parseInt(item.quantity) : 0);
        var price = (item.price ? parseFloat(item.price) : 0);

        var line = [
            quantity,
            (item.id == NEW_ITEM ? item.newItem : item.name),
            item.description,
            price,
            0,
            0,
            0
        ];

        if (!data.isSummary) {

            const unit = data.units.find((obj) => obj.type === item.unitsType);

            //if (unit) sheetInfo['unitNames'] = unit.names;

            line = line.concat([
                (item.markUp > 0 ? item.markUp : ''),
                (item.discount == 'T' ? 'Yes' : 'No'),
                item.units,
                (item.vendorId ? item.vendor : item.newVendor),
                item.manufacturer,
                item.mpn,
                '' // Add space for group control
            ]);
        }

        line.push(item.key); // Add key for hidden column
        line.push(item.id); // Add item ID for hidden column
        line.push(data.isSummary ? item.bomId : 0); // Add bom ID for hidden column

        values.push(line);
    }

    return values;
}

function getLaborValues(data) {

    var values = [];

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

        // Labor group

        var line = [0,key,'','',0,'',0];

        if (!data.isSummary) {
            
            line = line.concat(['','','','','','']); // Add spaces for extra columns
            line = line.concat(['']); // Add space for group control
        }

        line = line.concat(['','','']); // Add spaces for hidden columns

        values.push(line);

        // Labor items

        for (var i = 0; i < objLabor[key].length; i++) {

            var labor = objLabor[key][i];

            line = [
                labor.quantity,
                '',
                labor.name,
                labor.price,
                0,
                0,
                0
            ];

            if (!data.isSummary) {

                line = line.concat([
                    (labor.markUp > 0 ? labor.markUp : ''),
                    (labor.discount == 'T' ? 'Yes' : 'No'),
                    '',
                    '',
                    '',
                    '',
                    '' // Add space for group control
                ]);
            }

            line.push(data.isSummary ? '' : labor.key); // Add key to hidden column
            line.push(labor.id); // Add item ID for hidden column
            line.push(labor.sgId); // Add group ID to hidden column

            values.push(line);
        }
    }

    return values;
}

function getExpenseValues(data) {

    var values = [];

    // Set values

    for (var i = 0; i < data.expenses.length; i++) {
    
        var expense = data.expenses[i];

        values.push([
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
    }

    return values;
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

            var sheetInfo = getSheetInfo(range.values);

            var ranges = [];

            for (var i = rowFirst; i <= rowLast; i++) {

                ranges = ranges.concat(getRanges(i, sheetInfo, true));

                if (i >= sheetInfo.item.first && i <= sheetInfo.item.last) {

                    if (sheetInfo.isSummary) {

                        if (BOM_NAME) {

                            ranges = ranges.concat([{
                                range: ['A' + i + ':J' + i],
                                values: [[1,'Milestone Payment',BOM_NAME,0,0,0,0,'',MILESTONE_PAYMENT,0]]
                            }]);

                            BOM_NAME = '';
                        }
                        else {

                            ranges = ranges.concat([{
                                range: ['A' + i + ':G' + i],
                                values: [[1,'','',0,0,0,0]]
                            }]);
                        }
                    }
                    else {

                        ranges = ranges.concat([{
                            range: ['A' + i + ':J' + i],
                            values: [[1,'','',0,0,0,0,'','No','Ea']]
                        }]);
                    }
                }
                else if (i >= sheetInfo.expense.first && i <= sheetInfo.expense.last) {

                    ranges = ranges.concat([{
                        range: ['A' + i + ':I' + i],
                        values: [[1,'','',0,0,0,0,'','No']]
                    }]);
                }
            }

            ranges = ranges.concat(getTotalRowRanges(range.values.length, sheetInfo));
console.log('00000000 - ' + JSON.stringify(ranges));
            await excel.setSheet(context, sheet, {
                ranges: ranges
            });
        }
    });
}

function getSheetInfo(values) {

    // All first and last elements are the actual Excel row numbers

    var info = {
        isSummary: (values[0][0] == 'Quote'),
        isBom: (values[0][0] == 'Quantity'),
        item: {
            first: 0,
            last: 0
        },
        labor: {
            first: 0,
            last: 0,
            group: {}
        },
        expense: {
            first: 0,
            last: 0
        }
    };

    var groupRow = '';

    for (var i = 0; i < values.length; i++) {
        
        var row = i + 1;

        if ((info.isSummary && row <= 8) || (info.isBom && row == 1)) continue; // Skip the headings

        switch (values[i][0]) {
            case 'Items':
                info.item.first = row + 1;
                if (info.isSummary) info.labor.last = row - 1;
                break;
            case 'Labor':
                info.labor.first = row + 1;
                if (info.isBom) info.item.last = row - 1;
                break;
            case 'Expenses':
                info.expense.first = row + 1;
                info.labor.last = row - 1;
                break;
            case 'Total':
                info[info.isSummary ? 'item' : 'expense'].last = row - 1;
                break;
            default:
                { // Group elements are named by the actual Excel row number

                    if (info.labor.first > 0 && info.labor.last == 0) {

                        if (values[i][1].length) { // This is a labor group

                            groupRow = row.toString();

                            info.labor.group[groupRow] = {first: row + 1, last: row + 1};
                        }
                        else {

                            info.labor.group[groupRow].last = row;
                        }
                    }
                }
                break;
        }
    }

    return info;
}

function getRanges(row, sheetInfo, isInsert) {

    var ranges = [];

    if (row >= sheetInfo.item.first && row <= sheetInfo.item.last) {
        
        ranges = ranges.concat(getItemSectionRanges(row, sheetInfo, isInsert));
    }
    else if (row >= sheetInfo.labor.first && row <= sheetInfo.labor.last) {

        if (sheetInfo.labor.group[row.toString()]) { // This is a labor group

            ranges = ranges.concat(getLaborSectionGroupRanges(row, sheetInfo));
        }
        else {

            ranges = ranges.concat(getLaborSectionRanges(row, sheetInfo));
        }
    }
    else if (row >= sheetInfo.expense.first && row <= sheetInfo.expense.last) {
        
        ranges = ranges.concat(getExpenseSectionRanges(row, sheetInfo, isInsert));
    }

    return  ranges;
}

function getDefaultSectionRanges(row, sheetInfo) {

    // Range formats common to all types

    var ranges = [
        {
            range: ['D' + row + ':G' + row],
            numberFormat: '$#,###.00'
        },
        {
            range: ['E' + row],
            formula: 'A?*D?'
        },
        {
            range: ['G' + row],
            formula: 'A?*F?'
        },
        {
            range: ['A' + row],
            horizontalAlignment: 'center'
        }
    ];

    // Range formats common to BOM's only

    if (sheetInfo.isBom) {

        ranges = ranges.concat([
            {
                range: ['F' + row],
                formula: 'D?*(1+IF(I?="Yes",-1,1)*IF(ISNUMBER(H?),H?,Summary!$G$2))'
            },
            {
                range: ['H' + row + ':I' + row],
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

    // Input range formats

    var inputRanges = [];

    if (row >= sheetInfo.item.first && row <= sheetInfo.item.last) {
        
        inputRanges.push('A' + row);
    }
    else if (row >= sheetInfo.labor.first && row <= sheetInfo.labor.last && sheetInfo.isBom && !sheetInfo.labor.group[row.toString()]) {

        inputRanges.push('A' + row);
    }
    else if (row >= sheetInfo.expense.first && row <= sheetInfo.expense.last) {
        
        inputRanges.push('A' + row);
    }

    var ranges = ranges.concat([
        {
            range: inputRanges,
            color: COLOR_INPUT
        }
    ]);

    return ranges;
}

function getItemSectionRanges(row, sheetInfo, isInsert) {

    var ranges = getDefaultSectionRanges(row, sheetInfo);

    ranges = ranges.concat([
        {
            range: ['B' + row], // Set column B to text format to preserve leading zero's in item names
            data: [[0]]
        }
    ]);

    if (sheetInfo.isSummary) {
/*
        if (sheetInfo.bomName) {

            ranges = ranges.concat([
                {
                    range: ['C' + row],
                    formula: 'TEXTAFTER(CELL("filename",\'' + sheetInfo.bomName + '\'!A1),"]")'

                },
                {
                    range: ['D' + row],
                    formula: "'" + sheetInfo.bomName + "'!" + summaryFormulas.cost[sheetInfo.bomName]
                },
                {
                    range: ['F' + row],
                    formula: "'" + sheetInfo.bomName + "'!" + summaryFormulas.quote[sheetInfo.bomName]
                },
                {
                    range: ['C' + row + ':D' + row],
                    color: ''
                }
            ]);
        }
*/    }
    else {

        if (isInsert) {

            ranges = ranges.concat([
                {
                    range: ['A' + row],
                    bold: false
                }
            ]);
        }

        ranges = ranges.concat([
            {
                range: ['J' + row],
                dataValidationRule: {
                    list: {
                    inCellDropDown: true,
                    source: (sheetInfo.unitNames ? sheetInfo.unitNames : LISTS.units.filter(unit => unit.type != 3).map(unit => unit.names).join(',')) // Filter out labor units
                    }
                }
            },
        ]);
    }

    return ranges;
}

function getLaborSectionGroupRanges(row, sheetInfo) {

    var ranges = [];

    var group = sheetInfo.labor.group[row.toString()];

    if (group) {

        ranges = ranges.concat([
            {
                range: ['A' + row],
                formula: ('SUM(A' + group.first + ':A' + group.last + ')'),
                bold: true,
                horizontalAlignment: 'center'
            },
            {
                range: ['B' + row],
                bold: true
            },
            {
                range: ['E' + row],
                formula: ('SUM(E' + group.first + ':E' + group.last + ')'),
                bold: true,
                numberFormat: '$#,###.00'
            },
            {
                range: ['G' + row],
                formula: ('SUM(G' + group.first + ':G' + group.last + ')'),
                bold: true,
                numberFormat: '$#,###.00'
            },
            {
                range: [group.first + ':' + group.last],
                groupByRows: true
            }
        ]);
    }

    return ranges;
}

function getLaborSectionRanges(row, sheetInfo) {

    var ranges = getDefaultSectionRanges(row, sheetInfo);

    if (sheetInfo.isSummary) {

        ranges = ranges.concat([
            {
                range: ['A' + row],
                bold: false
            },
        ]);
    }

    return ranges;
}

function getExpenseSectionRanges(row, sheetInfo, isInsert) {

    var ranges = getDefaultSectionRanges(row, sheetInfo);

    if (isInsert) {

        ranges = ranges.concat([
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

    return ranges;
}

function getTotalRowRanges(row, sheetInfo) {

    var E = [];
    var G = [];
    
    if (sheetInfo.item.first) {

        E.push('SUM(E' + sheetInfo.item.first + ':E' + sheetInfo.item.last + ')');
        G.push('SUM(G' + sheetInfo.item.first + ':G' + sheetInfo.item.last + ')');
    }
    
    if (sheetInfo.labor.first) {

        E.push('SUMIFS(E' + sheetInfo.labor.first + ':E' + sheetInfo.labor.last + ',D' + sheetInfo.labor.first + ':D' + sheetInfo.labor.last + ',"<>")');
        G.push('SUMIFS(G' + sheetInfo.labor.first + ':G' + sheetInfo.labor.last + ',F' + sheetInfo.labor.first + ':F' + sheetInfo.labor.last + ',"<>")');
    }
    
    if (!sheetInfo.isSummary && sheetInfo.expense.first) {

        E.push('SUM(E' + sheetInfo.expense.first + ':E' + sheetInfo.expense.last + ')');
        G.push('SUM(G' + sheetInfo.expense.first + ':G' + sheetInfo.expense.last + ')');
    }

    var ranges = [
        {
            range: ['E' + row],
            formula: (E.length ? E.join('+') : '0')
        },
        {
            range: ['G' + row],
            formula: (G.length ? G.join('+') : '0')
        }
    ];

    return ranges;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onAddBom() {

    // Activate Summary for insert since it calls the generic onWorksheetChange function

    await Excel.run(async (context) => {

        context.workbook.worksheets.getItem('Summary').activate();

        await context.sync();
    });

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

    var params = {
        path: 'quote-revision',
        options: {data: {quoteId: $('#quoteList').val()}}
    };

    var quote = await api.post(params);

    var sel = $('#quoteList');

    sel.append('<option value="' + quote.id + '" selected>' + quote.name + '</option>');

    onQuote();
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
                else {
                    
                    let bomId = WORKSHEET[sheet.id.toString()].bomId;
                    if (!bomId) bomId = 0;

                    data.boms.push({id: bomId, name: sheet.name, items: [], expenses: []});
                }

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
