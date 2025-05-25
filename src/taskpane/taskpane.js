/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const { getgid } = require('process');
const api = require('../modules/api.js');
const excel = require('../modules/excel.js');

const NEW_ITEM = 3757;
//const COLOR_HEADER = '#E7E6E6';
const COLOR_INPUT = '#C6E0B4';
const LABEL_LABOR = 'Labor';
const LABEL_ITEMS = 'Items';

/* global console, document, Excel, Office */

Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {

        excel.initialize({
            excel: Excel
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

    await addSummary({
        defaultMU: data.defaultMU,
        labor: (data.boms.length ? data.boms[0].labor : []),
        items: data.items
    });

    for (var i = 0; i < data.boms.length; i++) {

        await addBom({
            defaultMU: data.defaultMU,
            bom: data.boms[i]
        });
    }
}

async function addSummary(data) {

    var summaryArray = getSummaryArray(data);

    // Set first & last labor row (starts at 1)
    
    var rowLaborFirst = 0;
    var rowLaborLast = 0;

    var index = summaryArray.findLastIndex(el => el[0] == LABEL_LABOR);

    if (index >= 0) {
    
        rowLaborFirst = index + 3;

        index = summaryArray.findIndex(el => el[0] == 'Total');
        
        rowLaborLast = index;
    }

    // Set first & last items row (starts at 1)
    
    var rowItemsFirst = 0;
    var rowItemsLast = 0;

    index = summaryArray.findLastIndex(el => el[0] == LABEL_ITEMS);

    if (index >= 0) {
    
        rowItemsFirst = index + 3;

        index = summaryArray.findLastIndex(el => el[0] == 'Total');
        
        rowItemsLast = index;
    }

    var rowLaborTotal = (rowLaborLast ? rowLaborLast + 1 : 0);
    var rowItemsTotal = rowItemsLast + 1;

    var ranges = [
        {
            range: ['D1'],
            formula: '$F$' + rowItemsTotal,
            numberFormat: '$#,###.00'
        },
        {
            range: ['D3'],
            formula: '($D$1-$D$' + rowItemsTotal + ')/IF($D$1>0,$D$1,1)'
        },
        {
            range: ['D4'],
            formula: '($D$1-$D$' + rowItemsTotal + ')/IF($D$' + rowItemsTotal + '>0,$D$' + rowItemsTotal + ',1)'
        },
        {
            range: ['B5'],
            formula: (rowLaborTotal ? '$D$' + rowLaborTotal : '=0')
        },
        {
            range: ['C5'],
            formula: 'ROUNDUP($D$1*$D$5,0)'
        },
        {
            range: ['D5'],
            formula: ('IF($D$' + rowItemsTotal + '=0,0,$B$5/$D$' + rowItemsTotal + ')')
        },
        {
            range: ['B6'],
            formula: ('$D$' + rowItemsTotal + '-$D$' + rowLaborTotal)
        },
        {
            range: ['C6'],
            formula: '$D$1-$C$5'
        },
        {
            range: ['D6'],
            formula: ('IF($D$' + rowItemsTotal + '=0,0,$B$6/$D$' + rowItemsTotal + ')')
        },
        {
            range: ['D7'],
            formula: ('IF($D$' + rowItemsTotal + '=0,0,$B$7/$D$' + rowItemsTotal + ')')
        },
        {
            range: ['F2'],
            formula: '$F$3-0.01'
        },
        {
            range: ['F3'],
            formula: '$F$4-0.01'
        },
        {
            range: ['F4'],
            formula: '$D$3-0.01'
        },
        {
            range: ['F5'],
            formula: '$D$3+0.01'
        },
        {
            range: ['F6'],
            formula: '$F$5+0.01'
        },
        {
            range: ['F7'],
            formula: '$F$6+0.01'
        },
        {
            range: ['G2'],
            formula: 'IF(1-F2=0,$D$1,$D$1/(1-F2))'
        },
        {
            range: ['G3'],
            formula: 'IF(1-F3=0,$D$1,$D$1/(1-F3))'
        },
        {
            range: ['G4'],
            formula: 'IF(1-F4=0,$D$1,$D$1/(1-F4))'
        },
        {
            range: ['G5'],
            formula: 'IF(1-F5=0,$D$1,$D$1/(1-F5))'
        },
        {
            range: ['G6'],
            formula: 'IF(1-F6=0,$D$1,$D$1/(1-F6))'
        },
        {
            range: ['G7'],
            formula: 'IF(1-F7=0,$D$1,$D$1/(1-F7))'
        },
        {
            range: ['D' + rowItemsFirst + ':D' + rowItemsLast],
            formula: 'A?*C?'
        },
        {
            range: ['F' + rowItemsFirst + ':F' + rowItemsLast],
            formula: 'A?*E?'
        },
        {
            range: ['D' + rowItemsTotal],
            formula: ('SUM(D' + rowItemsFirst + ':D' + rowItemsLast + ')'),
            numberFormat: '$#,###.00'
        },
        {
            range: ['F' + rowItemsTotal],
            formula: ('SUM(F' + rowItemsFirst + ':F' + rowItemsLast + ')'),
            numberFormat: '$#,###.00'
        },
        {
            range: ['D2:D7','F2:F7'],
            numberFormat: '#,###.00%'
        },
        {
            range: ['B4:C4','F1:G1','C10:D10','C' + (rowItemsFirst - 1) + ':F' + (rowItemsFirst - 1)],
            horizontalAlignment: 'right'
        },
        {
            range: ['A10:A' + rowLaborLast,'A' + (rowItemsFirst - 1) + ':A' + rowItemsLast],
            horizontalAlignment: 'center'
        },
        {
            range: ['D2','A' + rowItemsFirst + ':A' + rowItemsLast,'B' + rowItemsFirst + ':B' + rowItemsLast],
            color: COLOR_INPUT
        }
    ];

    var rngCurrencyFormat = {
        range: ['B5:C7','G2:G7','C' + rowItemsFirst + ':F' + rowItemsLast],
        numberFormat: '$#,###.00'
    };

    var rngBoldFormat = {
        range: ['A1:D7','F1:G1','A9:D10','A' + (rowItemsFirst - 2) + ':F' + (rowItemsFirst - 1),'A' + rowItemsTotal + ':F' + rowItemsTotal],
        bold: true
    };

    // Set labor format

    if (rowLaborFirst) {
        
        rngCurrencyFormat.range.push('C' + rowLaborFirst + ':D' + rowLaborLast);
        rngCurrencyFormat.range.push('D' + rowLaborTotal);
        
        rngBoldFormat.range.push('A' + (rowLaborFirst - 2) + ':D' + (rowLaborFirst - 1),'A' + rowLaborTotal + ':D' + rowLaborTotal);
        rngBoldFormat.range.push('A' + rowLaborTotal + ':D' + rowLaborTotal);
    }

    ranges.push(rngCurrencyFormat);
    ranges.push(rngBoldFormat);

    // Set labor formulas

    var rngFormula = [];

    var rowSum = 0;

    for (var i = rowLaborFirst; i < rowLaborLast; i++) {

        if (summaryArray[i -1][2] == '') {

            if (rowSum != 0) {

                rngFormula.push({
                    range:['A' + rowSum],
                    formula: ('SUM(A' + (rowSum + 1) + ':A' + (i - 1) + ')'),
                    bold: true
                });

                rngFormula.push({
                    range:['B' + rowSum],
                    bold: true
                });

                rngFormula.push({
                    range:['D' + rowSum],
                    formula: ('SUM(D' + (rowSum + 1) + ':D' + (i - 1) + ')'),
                    bold: true
                });
            }

            rowSum = i;
        }
        else {

            rngFormula.push({
                range:['D' + i],
                formula: 'A' + i + '*C' + i
            });
        }
    }

    await excel.addData("Summary", {
        data: summaryArray,
        ranges: ranges.concat(rngFormula)
    });
}

async function addBom(data) {

    await excel.addData(data.bom.name, {
        data: getBomArray(data.bom),
/*        ranges: [
            {
                firstRow: 3,
                rows: data.items.length,
                columns: ['F'],
                formula: 'C?*E?'
            },
            {
                firstRow: 3,
                rows: data.items.length,
                columns: ['L'],
                formula: 'ROUND(F?*(1+IF(K?="Yes",-1,1)*IF(ISNUMBER(J?),J?,' + data.defaultMU + ')/100),0)'
            },
            {
                firstRow: 3,
                rows: data.items.length,
                columns: ['E','F','J'],
                numberFormat: '0.00'
            },
            {
                firstRow: 3,
                rows: data.items.length,
                columns: ['C'],
                numberFormat: '0'
            },
            {
                color: 'lightgrey'
            },
            {
                firstRow: 3,
                rows: data.items.length,
                columns: ['A','B','C','E','J'],
                color: 'white'
            }
        ]
*/    });
}
/*
async function onBomExpenses() {

    var data = await getBomData("expenses");

    excel.addData({
        data: getExpenseArray(data.expenses),
        ranges: [
            {
                firstRow: 2,
                columns: ['D'],
                formula: 'B?*C?'
            },
            {
                firstRow: 2,
                columns: ['G'],
                formula: 'ROUND(D?*(1+IF(F?="Yes",-1,1)*IF(ISNUMBER(E?),E?,' + data.defaultMU + ')/100),0)'
            },
            {
                firstRow: 2,
                columns: ['C','D','G'],
                numberFormat: '0.00'
            },
            {
                firstRow: 2,
                columns: ['B'],
                numberFormat: '0'
            },
            {
                color: 'lightgrey'
            },
            {
                firstRow: 2,
                columns: ['B','C','E'],
                color: 'white'
            }
        ]
    });
}

async function onBomLabor() {

    var data = await getBomData("labor");

    excel.addData({
        data: getLaborArray(data.labor),
        ranges: [
            {
                firstRow: 2,
                columns: ['D'],
                formula: 'B?*C?'
            },
            {
                firstRow: 2,
                columns: ['F'],
                formula: 'ROUND(IF(ISNUMBER(E?),B?*E?,D?*(1+' + data.defaultMU + '/100)),0)'
            },
            {
                firstRow: 2,
                columns: ['C','D','E','F'],
                numberFormat: '0.00'
            },
            {
                firstRow: 2,
                columns: ['B'],
                numberFormat: '0'
            },
            {
                color: 'lightgrey'
            },
            {
                firstRow: 2,
                columns: ['B','C','E'],
                color: 'white'
            }
        ]

     });
}
*/
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getSummaryArray(data) {

    var labor = getLaborArray(data.labor, true);

    var dataArray = [
        ['Quote','','',0,'','GM','Sell'],
        ['MU (Default)','','',data.defaultMU,'',0,0],
        ['GM','','',0,'',0,0],
        ['MU','Cost','Sell',0,'',0,0],
        [LABEL_LABOR,0,0,0,'',0,0],
        [LABEL_ITEMS,0,0,0,'',0,0],
        ['Misc.',0,0,0,'',0,0],
        ['','','','','','',''],
    ];

    if (labor.length) {

        dataArray.push([LABEL_LABOR,'','','','','','']);
        dataArray.push(['Qty.','','Cost','Ext. Cost','','','']);

        dataArray = dataArray.concat(labor);

        dataArray.push(['Total','','',0,'','','']);
        dataArray.push(['','','','','','','']);
    }

    dataArray.push([LABEL_ITEMS,'','','','','','']);
    dataArray.push(['Qty.','Description','Cost','Ext. Cost','Quote','Ext. Quote','']);

    for (var i = 0; i < data.items.length; i++) {
    
        var item = data.items[i];

        dataArray.push([
            item.quantity,
            item.desc,
            item.cost,
            item.quantity * item.cost,
            item.quote,
            item.quantity * item.quote,
            ''
        ]);   
    }

    dataArray.push(['Total','','',0,'',0,'']);   

    return dataArray;
}

function getBomArray(data) {

    var dataArray = [
        ['Item','Description','Quantity','Units','Price','Amount','Vendor','Manufacturer','MPN','MU%','Discount','Quote'],
        ['Items','','','','','','','','','','','']
    ];

    for (var i = 0; i < data.items.length; i++) {
    
        var item = data.items[i];

        var quantity = (item.quantity ? parseInt(item.quantity) : 0);
        var price = (item.price ? parseFloat(item.price) : 0);
        var markup = (item.markup ? parseFloat(item.markup) : 0);
        var defaultMU = (data.defaultMU ? parseFloat(data.defaultMU) : 0);

        dataArray.push([
            (item.itemId == NEW_ITEM ? item.newItem : item.name),
            (item.itemId == NEW_ITEM ? item.newDescription : item.description),
            quantity,
            item.units,
            price,
            (quantity * price),
            (item.vendorId ? item.vendor : item.newVendor),
            item.manufacturer,
            item.mpn,
            (markup > 0 ? markup : ''),
            (item.discount == 'T' ? 'Yes' : 'No'),
            (quantity * price * (1 + (markup > 0 ? markup : defaultMU)))
        ]);
    }

    dataArray.push(['Labor','','','','','','','','','','','']);

    var labor = getLaborArray(data.labor);

    if (labor.length) dataArray = dataArray.concat(labor);

    dataArray.push(['Expenses','','','','','','','','','','','']);

    for (var i = 0; i < data.expenses.length; i++) {
    
        var expense = data.expenses[i];

        var quantity = (expense.quantity ? parseInt(expense.quantity) : 0);
        var price = (expense.price ? parseFloat(expense.price) : 0);
        var markup = (expense.markup ? parseFloat(expense.markup) : 0);
        var defaultMU = (data.defaultMU ? parseFloat(data.defaultMU) : 0);

        dataArray.push([
            '',
            expense.name,
            quantity,
            '',
            price,
            (quantity * price),
            '',
            '',
            '',
            (markup > 0 ? markup : ''),
            (expense.discount == 'T' ? 'Yes' : 'No'),
            (quantity * price * (1 + (markup > 0 ? markup : defaultMU)))
        ]);   

        dataArray.push(['Total','','','','',0,'','','','','',0]);
    }

    return dataArray;
}

function getLaborArray(data, summary = false) {

    var laborArray = [];

    var groupName = '';

    for (var i = 0; i < data.length; i++) {
    
        var labor = data[i];

        if (groupName != labor.groupName) {

            groupName = labor.groupName;

            if (summary)
                laborArray.push([0,labor.groupName,'',0,'','','']);
            else
                laborArray.push(['',labor.groupName,0,'',0,'','','','',0,0,0]);
        }

        if (summary)
            laborArray.push([labor.quantity,labor.name,labor.cost,0,'','','']);
        else
            laborArray.push(['',labor.name,labor.quantity,'',0,0,'','','',0,0,0]);
    }

    return laborArray;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onReload() {


}

async function onSave() {

}

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
