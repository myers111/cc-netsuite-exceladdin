/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const { getgid, exit } = require('process');
const api = require('../modules/api.js');
const excel = require('../modules/excel.js');

const NEW_ITEM = 3757;
const COLOR_INPUT = '#C6E0B4';
const LABEL_HEADER = ['Quantity','Item','Description','Cost','Ext. Cost','Quote','Ext. Quote'];
const LABEL_HEADER_EX = ['Units','Vendor','Manufacturer','MPN','MU%','Discount'];
const LABEL_LABOR = 'Labor';
const LABEL_ITEMS = 'Items';
const LABEL_EXPENSES = 'Expenses';
const LABEL_TOTAL = 'Total';

/* global console, document, Excel, Office */

var WORKSHEET = {};

var summaryFormulas = null;

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

    var id = $('#revisionList').val();

    if (!id) return;

    summaryFormulas = {
        labor: {},
        cost: {},
        quote: {}
    };

    var params = {
        path: 'revision',
        options: {
            id: id,
            quoteId: $('#quoteList').val()
        }
    };

    var data = await api.get(params);

    const promises = [];

    for (var i = 0; i < data.boms.length; i++) {

        promises.push(addBom({
            defaultMU: data.defaultMU,
            bom: data.boms[i]
        }));
    }

    await Promise.all(promises);

    await addSummary(data);
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

    var laborData = getLaborData({
        defaultMU: data.defaultMU,
        boms: data.boms,
        rowFirst: dataArray.length + 1,
        isSummary: true
    });

    dataArray = dataArray.concat(laborData.values);
    dataRanges = dataRanges.concat(laborData.ranges);

    var itemData = getItemData({
        defaultMU: data.defaultMU,
        items: data.items,
        rowFirst: dataArray.length + 1,
        isSummary: true
    });

    dataArray = dataArray.concat(itemData.values);
    dataRanges = dataRanges.concat(itemData.ranges);

    dataArray.push([LABEL_TOTAL,'','','',0,'',0]);

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
        {
            range:['H:H'],
            hideColumns: true
        }
    ]);

    await excel.addData("Summary", WORKSHEET, {
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
        bomName: data.bom.name,
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

    dataArray.push([LABEL_TOTAL,'','','',0,'',0,'','','','','','']);

    // Set summaryFormulas

    summaryFormulas.cost[data.bom.name] = "E" + dataArray.length;
    summaryFormulas.quote[data.bom.name] = "G" + dataArray.length;

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
            formula: 'SUM(G' + itemData.rowFirst + ':G' + itemData.rowLast + ')+SUMIFS(G' + laborData.rowFirst + ':G' + laborData.rowLast + ',F' + laborData.rowFirst + ':F' + laborData.rowLast + ',"<>")+SUM(G' + expenseData.rowFirst + ':G' + expenseData.rowLast + ')',
            numberFormat: '$#,###.00'
        },
        {
            range:['H:M'],
            groupByColumns: true
        },
        {
            range:['N:N'],
            hideColumns: true
        }
    ]);

    await excel.addData(data.bom.name, WORKSHEET, {
        data: dataArray,
        ranges: dataRanges,
        bomId: data.bom.id
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

        if (data.isSummary) {

            if (item.bomId > 0) {

                var bomName = item.description;

                itemData.ranges.push({
                    range: ['D' + (itemData.rowFirst + i)],
                    formula: "'" + bomName + "'!" + summaryFormulas.cost[bomName]
                });

                itemData.ranges.push({
                    range: ['F' + (itemData.rowFirst + i)],
                    formula: "'" + bomName + "'!" + summaryFormulas.quote[bomName]
                });
            }
            else {

                itemData.ranges.push({
                    range: ['B' + (itemData.rowFirst + i),'D' + (itemData.rowFirst + i)],
                    color: COLOR_INPUT
                });
            }
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
                formula: 'ROUND(D?*(1+IF(M?="Yes",-1,1)*IF(ISNUMBER(L?),L?,Summary!$G$2)),0)'
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

            var row = (laborData.rowFirst + laborData.values.length - 2);

            var idString = labor.id.toString();

            if (data.isSummary) {

                laborData.ranges.push({
                    range: ['A' + row],
                    formula: summaryFormulas.labor[key][idString].qty
                });

                laborData.ranges.push({
                    range: ['D' + row],
                    formula: summaryFormulas.labor[key][idString].cost
                });

                laborData.ranges.push({
                    range: ['F' + row],
                    formula: summaryFormulas.labor[key][idString].quote
                });
            }
            else {

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
                    formula: 'ROUND(D?*(1+IF(M?="Yes",-1,1)*IF(ISNUMBER(L?),L?,Summary!$G$2)),0)'
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

    onRevision();
}

async function onSave() {

    var revisionId = $('#revisionList').val();

    if (!revisionId) return;

    try {

        await Excel.run(async (context) => {

            var data = {
                id: revisionId,
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
                    data.boms.push({id: WORKSHEET[sheet.id.toString()].bomId, items: [], labor: [], expenses: []});

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
                        else {

                            data.boms[data.boms.length - 1].items.push({
                                quantity: values[0],
                                name: values[1],
                                description: values[2],
                                price: values[3],
                                units: values[7],
                                vendor: values[8],
                                manufacturer: values[9],
                                mpn: values[10],
                                markUp: values[11],
                                discount: values[12]
                            });
                        }
                    }
                    else if (section == LABEL_LABOR) {

                        data.boms[data.boms.length - 1].labor.push({
                            quantity: values[0],
                            group: values[1],
                            item: values[2],
                            price: values[3]
                        });
                    }
                    else if (section == LABEL_EXPENSES) {

                        data.boms[data.boms.length - 1].expenses.push({
                            quantity: values[1],
                            price: values[2],
                            markUp: values[4],
                            discount: values[5]
                        });
                    }
                }
            }

            var params = {
                path: 'revision',
                options: {data: data}
            };

            await api.post(params);
        });
    }
    catch (error) {

        console.error(error);
    }
}
