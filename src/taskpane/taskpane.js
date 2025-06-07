/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const { getgid, exit } = require('process');
const api = require('../modules/api.js');
const excel = require('../modules/excel.js');

const NEW_ITEM = 3757;
//const COLOR_HEADER = '#E7E6E6';
const COLOR_INPUT = '#C6E0B4';
const LABEL_HEADER = ['Quantity','Item','Description','Cost','Ext. Cost','Quote','Ext. Quote'];
const LABEL_HEADER_EX = ['Units','Vendor','Manufacturer','MPN','MU%','Discount'];
const LABEL_LABOR = 'Labor';
const LABEL_ITEMS = 'Items';
const LABEL_EXPENSES = 'Expenses';

/* global console, document, Excel, Office */

Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {

        excel.initialize({
            excel: Excel,
            groupByRows: Excel.GroupOption.byRows
        });

        document.getElementById("reload").onclick = onReload;
        document.getElementById("save").onclick = onSave;

        initList('customer');
        initList('project');
        initList('quote');
        initList('revision', false);
    }
});

async function initList(id, onFocus = true) {

    var selector = '#' + id + 'List';

    if (onFocus) $(selector).focus(function() { loadList(id); });

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
        case 'revision':
            var quoteId = $('#quoteList').val();
            if (parseInt(quoteId)) options['quoteId'] = quoteId;
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
            emptyList("revision");
            await onQuote();
            break;
        case 'revision':
            await onRevision();
    }
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onQuote() {

    var quoteId = $('#quoteList').val();

    if (quoteId) {

        await loadList('revision');

        onRevision();
    }

    document.getElementById("revisionControls").style.display = (quoteId > 0 ? '' : 'none');
    document.getElementById("controls").style.display = (quoteId > 0 ? '' : 'none');
}

async function onRevision() {

    var params = {
        path: 'revision',
        options: {
            id: $('#revisionList').val(),
            quoteId: $('#quoteList').val()
        }
    };

    var data = await api.get(params);

    await addSummary(data);

    const promises = [];

    for (var i = 0; i < data.boms.length; i++) {

        promises.push(addBom({
            defaultMU: data.defaultMU,
            bom: data.boms[i]
        }));
    }

    await Promise.all(promises);

    //await setLinkedValues();
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

    var dataRanges = [
        {
            range: ['G2'],
            color: COLOR_INPUT
        }
    ];

    var itemData = getItemData({
        defaultMU: data.defaultMU,
        items: data.items,
        rowFirst: dataArray.length + 1,
        isSummary: true
    });

    dataArray = dataArray.concat(itemData.values);
    dataRanges = dataRanges.concat(itemData.ranges);

    var laborData = getLaborData({
        defaultMU: data.defaultMU,
        boms: data.boms,
        rowFirst: dataArray.length + 1,
        isSummary: true
    });

    dataArray = dataArray.concat(laborData.values);
    dataRanges = dataRanges.concat(laborData.ranges);

    dataArray.push(['Total','','','',0,'',0]);

    dataRanges = dataRanges.concat([
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
            formula: '($G$1-SUM($G' + itemData.rowFirst + ':$G' + itemData.rowLast + '))/IF($G$1>0,$G$1,1)'
        },
        {
            range: ['G4'],
            formula: '($G$1-SUM($G' + itemData.rowFirst + ':$G' + itemData.rowLast + '))/IF(SUM($G' + itemData.rowFirst + ':$G' + itemData.rowLast + ')>0,SUM($G' + itemData.rowFirst + ':$G' + itemData.rowLast + '),1)'
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
            formula: ('IF($G$' + dataArray.length + '=0,0,$E$5/$G$' + dataArray.length + ')')
        },
        {
            range: ['E6'],
            formula: 'SUMIFS(E' + laborData.rowFirst + ':E' + laborData.rowLast + ',D' + laborData.rowFirst + ':D' + laborData.rowLast + ',"<>")'
        },
        {
            range: ['F6'],
            formula: 'SUMIFS(G' + laborData.rowFirst + ':G' + laborData.rowLast + ',F' + laborData.rowFirst + ':F' + laborData.rowLast + ',"<>")'
        },
        {
            range: ['G6'],
            formula: ('IF($G$' + dataArray.length + '=0,0,$E$6/$G$' + dataArray.length + ')')
        },
        {
            range: ['E' + dataArray.length],
            formula: 'SUM(E' + itemData.rowFirst + ':E' + itemData.rowLast + ')+SUMIFS(E' + laborData.rowFirst + ':E' + laborData.rowLast + ',D' + laborData.rowFirst + ':D' + laborData.rowLast + ',"<>")'
        },
        {
            range: ['G' + dataArray.length],
            formula: 'SUM(G' + itemData.rowFirst + ':G' + itemData.rowLast + ')+SUMIFS(G' + laborData.rowFirst + ':G' + laborData.rowLast + ',F' + laborData.rowFirst + ':F' + laborData.rowLast + ',"<>")'
        },
    ]);

    await excel.addData("Summary", {
        data: dataArray,
        ranges: dataRanges
    });
}

async function addBom(data) {

    var dataArray = [
        LABEL_HEADER.concat(LABEL_HEADER_EX)
    ];

    var dataRanges = [];

    // Items

    var itemData = getItemData({
        defaultMU: data.defaultMU,
        items: data.bom.items,
        rowFirst: dataArray.length + 1,
        isSummary: false
    });

    itemData.values[0] = itemData.values[0].concat(['','','','','','']);

    for (var i = 0; i < data.bom.items.length; i++) {
    
        var item = data.bom.items[i];

        var markup = (item.markup ? parseFloat(item.markup) : 0);

        var index = i + 1;

        itemData.values[index] = itemData.values[index].concat([
            item.units,
            (item.vendorId ? item.vendor : item.newVendor),
            item.manufacturer,
            item.mpn,
            (markup > 0 ? markup : ''),
            (item.discount == 'T' ? 'Yes' : 'No')
        ]);
    }

    itemData.ranges = itemData.ranges.concat([
        {
            range: ['D' + itemData.rowFirst + ':D' + itemData.rowLast],
            color: COLOR_INPUT
        },
    ]);

    dataArray = dataArray.concat(itemData.values);
    dataRanges = dataRanges.concat(itemData.ranges);

    // Labor

    var laborData = getLaborData({
        defaultMU: data.defaultMU,
        boms: [data.bom],
        rowFirst: dataArray.length + 1,
        isSummary: false
    });

    for (var i = 0; i < laborData.values.length; i++) {
    
        laborData.values[i] = laborData.values[i].concat(['','','','','','']);
    }

    dataArray = dataArray.concat(laborData.values);
    dataRanges = dataRanges.concat(laborData.ranges);

    // Expenses
    
    var expenseData = getExpenseData({
        defaultMU: data.defaultMU,
        expenses: data.bom.expenses,
        rowFirst: dataArray.length + 1,
        isSummary: false
    });

    dataArray = dataArray.concat(expenseData.values);
    dataRanges = dataRanges.concat(expenseData.ranges);

    // Totals

    dataArray.push(['Total','','','',0,'',0,'','','','','','']);

    dataRanges = dataRanges.concat([
        {
            range: ['A1:M1','A' + dataArray.length + ':M' + dataArray.length],
            bold: true
        },
        {
            range: ['E' + dataArray.length],
            formula: 'SUM(E' + itemData.rowFirst + ':E' + itemData.rowLast + ')+SUMIFS(E' + laborData.rowFirst + ':E' + laborData.rowLast + ',D' + laborData.rowFirst + ':D' + laborData.rowLast + ',"<>")+SUM(E' + expenseData.rowFirst + ':E' + expenseData.rowLast + ')',
            numberFormat: '$#,###.00'
        },
        {
            range: ['G' + dataArray.length],
            formula: 'SUM(G' + itemData.rowFirst + ':G' + itemData.rowLast + ')+SUMIFS(G' + laborData.rowFirst + ':G' + laborData.rowLast + ',F' + laborData.rowFirst + ':F' + laborData.rowLast + ',"<>")+SUM(F' + expenseData.rowFirst + ':F' + expenseData.rowLast + ')',
            numberFormat: '$#,###.00'
        },
        {
            range:['H:M'],
            groupByColumns: true
        }
    ]);

    await excel.addData(data.bom.name, {
        data: dataArray,
        ranges: dataRanges
    });
}

async function setLinkedValues() {

    Excel.run(async (context) => {

        // Get BOM values

        var linkedValues = {};

        var sheets = context.workbook.worksheets;
    
        sheets.load("items/name");

        await context.sync();

        sheets.items.forEach(async function (sheet) {

            if (sheet.name != 'Summary') {

                var range = sheet.getRange();

                range.load('values');
                
                await context.sync();

                if (range.values.length && range.values[0][0] == 'Quantity') {
/*
                    var values = {
                        labor: {},
                        cost: range.values[range.values.length - 1][4],
                        quote: range.values[range.values.length - 1][6]
                    }; 
*//*
                    var labor = false;

                    for (var i = 0; i < range.values.length; i++) {

                        var v = range.values[i];

                        if (labor) {

                            if (v[0] == LABEL_EXPENSES) exit;

                            if (v[3] != '') {
                                
                                values.labor[v[1]][v[2]] = {
                                    qty: v[0],
                                    cost: v[3]
                                }
                            }
                        }
                        else if (v[0] == LABEL_LABOR) {
                            
                            labor = true;
                        }
                    }
*/
                    //linkedValues[sheet.name] = values;
                }
            }
        });
return;
        // Set Summary formulas

        var sheet = context.workbook.worksheets.getItem('Summary');

        var itemsRange = sheet.getRange('A:A').find(LABEL_ITEMS, {
            completeMatch: true, /* Match the whole cell value, not any part of the text. */
            matchCase: false /* Make the search case-insensitive. */,
            searchDirection: Excel.SearchDirection.forward
        });

        var laborRange = sheet.getRange('A:A').find(LABEL_LABOR, {
            completeMatch: true, /* Match the whole cell value, not any part of the text. */
            matchCase: false /* Make the search case-insensitive. */,
            searchDirection: Excel.SearchDirection.forward,
            after: itemsRange
        });

        var expRange = sheet.getRange('A:A').find(LABEL_EXPENSES, {
            completeMatch: true, /* Match the whole cell value, not any part of the text. */
            matchCase: false /* Make the search case-insensitive. */,
            searchDirection: Excel.SearchDirection.forward,
            after: laborRange
        });

        itemsRange.load('rowIndex');
        laborRange.load('rowIndex');
        expRange.load('rowIndex');
                
        context.sync();
return;
        var range = sheet.getRange('A' + (laborRange.rowIndex + 2) + ':D' + expRange.rowIndex - 1);

        range.load('values');
    });
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getItemData(data) {

    /* data = { defaultMU, rowFirst, items } *******************************************************************************************************************/
    
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
        rowFirst: data.rowFirst + 1,
        rowLast: 0
    };

    // Set values

    for (var i = 0; i < data.items.length; i++) {
    
        var item = data.items[i];

        var quantity = (item.quantity ? parseInt(item.quantity) : 0);
        var price = (item.price ? parseFloat(item.price) : 0);

        itemData.values.push([
            quantity,
            (item.id == NEW_ITEM ? item.newItem : item.name),
            (item.id == NEW_ITEM ? item.newDescription : item.description),
            price,
            0,
            0,
            0
        ]);

        if (!item.isBom) {

            itemData.ranges.push({
                range: ['B' + (itemData.rowFirst + i),'D' + (itemData.rowFirst + i)],
                color: COLOR_INPUT
            });
        }
    }

    itemData.rowLast = itemData.rowFirst + itemData.values.length - 2;

    // Set ranges

    itemData.ranges = itemData.ranges.concat([
        {
            range: ['A' + itemData.rowFirst + ':A' + itemData.rowLast,'M' + itemData.rowFirst + ':M' + itemData.rowLast],
            horizontalAlignment: 'center'
        },
        {
            range: ['A' + itemData.rowFirst + ':A' + itemData.rowLast,'C' + itemData.rowFirst + ':C' + itemData.rowLast],
            color: COLOR_INPUT
        },
        {
            range: ['D' + itemData.rowFirst + ':G' + itemData.rowLast],
            numberFormat: '$#,###.00'
        },
        {
            range: ['L' + itemData.rowFirst + ':L' + itemData.rowLast],
            numberFormat: '#,###.00%'
        },
        {
            range: ['E' + itemData.rowFirst + ':E' + itemData.rowLast],
            formula: 'A?*D?'
        },
        {
            range: ['G' + itemData.rowFirst + ':G' + itemData.rowLast],
            formula: 'A?*F?'
        }
    ]);

    if (!data.isSummary) {
        
        itemData.ranges = itemData.ranges.concat([
            {
                range: ['F' + itemData.rowFirst + ':F' + itemData.rowLast],
                formula: 'ROUND(D?*(1+IF(M?="Yes",-1,1)*IF(ISNUMBER(L?),L?,' + data.defaultMU + ')/100),0)'
            }
        ]);
    }

    return itemData;
}

function getLaborData(data) {

    /* data = { defaultMU, rowFirst, boms } *******************************************************************************************************************/
    
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
        rowFirst: data.rowFirst + 1,
        rowLast: 0
    };

    // Create labor object

    var objLabor = {};
    
    for (var i = 0; i < data.boms.length; i++) {

        for (var j = 0; j < data.boms[i].labor.length; j++) {

            var labor = data.boms[i].labor[j];

            if (!objLabor[labor.groupName]) objLabor[labor.groupName] = [];

            var index = objLabor[labor.groupName].findIndex(obj => obj.id == labor.id && obj.cost == labor.cost);

            if (index < 0) objLabor[labor.groupName].push(labor);
        }
    }

    // Set values

    for (const key in objLabor) {

        if (!objLabor.hasOwnProperty(key)) continue;

        laborData.values.push([0,key,'','',0,'',0]);

        for (var i = 0; i < objLabor[key].length; i++) {

            var labor = objLabor[key][i];

            laborData.values.push([labor.quantity,'',labor.name,labor.cost,0,0,0]);
        }
    }

    laborData.rowLast = laborData.rowFirst + laborData.values.length - 2;

    // Set ranges

    laborData.ranges = laborData.ranges.concat([
        {
            range: ['A' + laborData.rowFirst + ':A' + laborData.rowLast],
            horizontalAlignment: 'center'
        },
        {
            range: ['D' + laborData.rowFirst + ':G' + laborData.rowLast],
            numberFormat: '$#,###.00'
        }
    ]);

    // Set labor formulas & grouping

    var rngLaborGroup = [];

    var rowSum = 0;

    for (var i = laborData.rowFirst; i < laborData.rowLast; i++) {

        var index = i - laborData.rowFirst + 1;

        if (laborData.values[index][3] == '') {

            if (rowSum != 0) setLaborGroupRanges(rngLaborGroup, rowSum, i);

            rowSum = i;
        }
        else {

            if (!data.isSummary) {

                rngLaborGroup.push({
                    range: ['A' + i,'D' + i],
                    color: COLOR_INPUT
                });

                rngLaborGroup.push({
                    range: ['F' + i],
                    formula: 'ROUND(D?*(1+IF(M?="Yes",-1,1)*IF(ISNUMBER(L?),L?,' + data.defaultMU + ')/100),0)'
                });
            }

            rngLaborGroup.push({
                range: ['E' + i],
                formula: 'A' + i + '*D' + i
            });

            rngLaborGroup.push({
                range: ['G' + i],
                formula: 'A' + i + '*F' + i
            });
        }

    }

    if (rowSum != 0) setLaborGroupRanges(rngLaborGroup, rowSum, laborData.rowLast + 1);

    laborData.ranges = laborData.ranges.concat(rngLaborGroup);

    return laborData;
}

function setLaborGroupRanges(rngLabor, rowSum, i) {

    rngLabor.push({
        range:['A' + rowSum],
        formula: ('SUM(A' + (rowSum + 1) + ':A' + (i - 1) + ')'),
        bold: true
    });

    rngLabor.push({
        range:['B' + rowSum],
        bold: true
    });

    rngLabor.push({
        range:['E' + rowSum],
        formula: ('SUM(E' + (rowSum + 1) + ':E' + (i - 1) + ')'),
        bold: true
    });

    rngLabor.push({
        range:['G' + rowSum],
        formula: ('SUM(G' + (rowSum + 1) + ':G' + (i - 1) + ')'),
        bold: true
    });

    rngLabor.push({
        range:[(rowSum + 1) + ':' + (i - 1)],
        groupByRows: true
    });
}

function getExpenseData(data) {

    /* data = { defaultMU, rowFirst, expenses } *******************************************************************************************************************/
    
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
        rowFirst: data.rowFirst + 1,
        rowLast: 0
    };

    // Set values

    for (var i = 0; i < data.expenses.length; i++) {
    
        var expense = data.expenses[i];

        var quantity = (expense.quantity ? parseInt(expense.quantity) : 0);
        var price = (expense.price ? parseFloat(expense.price) : 0);
        var markup = (expense.markup ? parseFloat(expense.markup) : 0);
        var defaultMU = (data.defaultMU ? parseFloat(data.defaultMU) : 0);

        expenseData.values.push([
            quantity,
            '',
            expense.name,
            price,
            0,
            0,
            0,
            '',
            '',
            '',
            '',
            (markup > 0 ? markup : ''),
            (expense.discount == 'T' ? 'Yes' : 'No')
        ]);   
    }

    expenseData.rowLast = expenseData.rowFirst + expenseData.values.length - 2;

    // Set ranges

    expenseData.ranges = expenseData.ranges.concat([
        {
            range: ['A' + expenseData.rowFirst + ':A' + expenseData.rowLast,'M' + expenseData.rowFirst + ':M' + expenseData.rowLast],
            horizontalAlignment: 'center'
        },
        {
            range: ['A' + expenseData.rowFirst + ':A' + expenseData.rowLast,'D' + expenseData.rowFirst + ':D' + expenseData.rowLast],
            color: COLOR_INPUT
        },
        {
            range: ['D' + expenseData.rowFirst + ':G' + expenseData.rowLast],
            numberFormat: '$#,###.00'
        },
        {
            range: ['L' + expenseData.rowFirst + ':L' + expenseData.rowLast],
            numberFormat: '#,###.00%'
        },
        {
            range: ['E' + expenseData.rowFirst + ':E' + expenseData.rowLast],
            formula: 'A?*D?'
        },
        {
            range: ['G' + expenseData.rowFirst + ':G' + expenseData.rowLast],
            formula: 'A?*F?'
        }
    ]);

    if (!data.isSummary) {
        
        expenseData.ranges = expenseData.ranges.concat([
            {
                range: ['F' + expenseData.rowFirst + ':F' + expenseData.rowLast],
                formula: 'ROUND(D?*(1+IF(M?="Yes",-1,1)*IF(ISNUMBER(L?),L?,' + data.defaultMU + ')/100),0)'
            }
        ]);
    }

    return expenseData;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onReload() {


}

async function onSave() {

}
/*
async function saveSummary() {

    var quoteId = $('#quoteList').val();

    if (!quoteId) return;

    try {

        await Excel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            var range = sheet.getUsedRange();

            range.load("values");

            await context.sync();

            range.values.shift();

            var quote = {
                id: quoteId,
                items: []
            };

            for (var i = 0; i < range.values.length; i++) {
            
                var values = range.values[i];

                //if (!values[0].length) break;

                quote.items.push({
                    quantity: values[0],
                    description: values[1]
                });
            }

            var params = {
                path: 'quote-summary',
                options: {quote: quote}
            };

            await api.post(params);
        });
    }
    catch (error) {

        console.error(error);
    }
}

async function saveOverview() {

    var quoteId = $('#quoteList').val();

    if (!quoteId) return;

    try {

        await Excel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            var range = sheet.getRange('E2');

            range.load("values");

            await context.sync();

            var quote = {
                id: quoteId,
                defaultMU: range.values[0].values[0]
            };

            var params = {
                path: 'quote-overview',
                options: {quote: quote}
            };

            await api.post(params);
        });
    }
    catch (error) {

        console.error(error);
    }
}

async function saveItems() {

    var bomId = $('#bomList').val();

    try {

        await Excel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            var range = sheet.getUsedRange();

            range.load("values");

            await context.sync();

            range.values.shift();

            var bom = {
                id: bomId,
                items: []
            };

            for (var i = 0; i < range.values.length; i++) {
            
                var values = range.values[i];

                if (!values[0].length) break;

                bom.items.push({
                    name: values[0],
                    description: values[1],
                    quantity: values[2],
                    units: values[3],
                    price: values[4],
                    vendor: values[6],
                    manufacturer: values[7],
                    mpn: values[8],
                    markUp: values[9],
                    discount: values[10]
                });
            }

            var params = {
                path: 'bom-items',
                options: {bom: bom}
            };

            await api.post(params);
        });
    }
    catch (error) {

        console.error(error);
    }
}

async function saveExpenses() {

    var bomId = $('#bomList').val();

    try {

        await Excel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            var range = sheet.getUsedRange();

            range.load("values");

            await context.sync();

            range.values.shift();

            var bom = {
                id: bomId,
                expenses: []
            };

            for (var i = 0; i < range.values.length; i++) {
            
                var values = range.values[i];

                if (!values[0].length) break;

                bom.expenses.push({
                    quantity: values[1],
                    price: values[2],
                    markUp: values[4],
                    discount: values[5]
                });
            }

            var params = {
                path: 'bom-expenses',
                options: {bom: bom}
            };

            await api.post(params);
        });
    }
    catch (error) {

        console.error(error);
    }
}

async function saveLabor() {

    var bomId = $('#bomList').val();
    var laborId = $('#laborList').val();

    try {

        await Excel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            var range = sheet.getUsedRange();

            range.load("values");

            await context.sync();

            range.values.shift();

            var bom = {
                id: bomId,
                laborId: laborId,
                labor: []
            };

            for (var i = 0; i < range.values.length; i++) {
            
                var values = range.values[i];

                if (!values[0].length) break;

                bom.labor.push({
                    quantity: values[1],
                    cost: values[2],
                    sell: values[4]
                });
            }

            var params = {
                path: 'bom-labor',
                options: {bom: bom}
            };

            await api.post(params);
        });
    }
    catch (error) {

        console.error(error);
    }
}
*/