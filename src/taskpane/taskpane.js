/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const { getgid } = require('process');
const api = require('../modules/api.js');
const excel = require('../modules/excel.js');

const NEW_ITEM = 3757;
const HIGHLIGHT_CLASS = "highLight";

/* global console, document, Excel, Office */

Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {

        window.onload = onRefresh;

        document.getElementById("reload").onclick = onReload;
        document.getElementById("save").onclick = onSave;

        initList('customer');
        initList('project');
        initList('quote');
        initList('revision', false);

        initButton("summary");

        excel.initialize({
            excel: Excel
        });
    }
});

async function  onRefresh() {

    excel.clearData();
}

async function initList(id, onFocus = true) {

    var selector = '#' + id + 'List';

    if (onFocus) $(selector).focus(function() { loadList(id); });

    $(selector).on('change', function() { onChange(id); });
}

async function initButton(id) {

    var selector = '#' + id;

    $(selector).on('click', function() { onClick(id); });
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

    excel.clearData();

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

async function onClick(id) {

    var buttonId = getButton();

    if (id == buttonId) return;

    if (buttonId) document.getElementById(buttonId).className = "";

    switch (id) {
        case 'summary':
            onSummary();
            break;
        default:
            onBomLabor();
            break;
    }

    document.getElementById(id).className = HIGHLIGHT_CLASS;
}

function addButtons(boms) {

    removeButtons();

    for (var i = 0; i < boms.length; i++) {
                
        var bom = boms[i];

        $("#bomButtons").append('<button id="bom' + bom.id + '">' + bom.name + '</button>');

        initButton("bom" + bom.id);
    }
}

function removeButtons() {

    $("#bomButtons").empty();
}

function getButton() {

    return $("button." + HIGHLIGHT_CLASS).attr("id");
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onQuote() {

    var quoteId = $('#quoteList').val();

    if (quoteId) {

        await loadList('revision');

        onRevision();
    }
    else {

        removeButtons();
    }

    document.getElementById("revisionControls").style.display = (quoteId > 0 ? '' : 'none');
    document.getElementById("controls").style.display = (quoteId > 0 ? '' : 'none');
}

async function onRevision() {

    var params = {
        path: 'revision-boms',
        options: {
            id: $('#revisionList').val(),
            quoteId: $('#quoteList').val()
        }
    };

    var data = await api.get(params);

    addButtons(data);

    onClick("summary");
}

async function onSummary() {

    var params = {
        path: 'revision-summary',
        options: {
            id: $('#revisionList').val(),
            quoteId: $('#quoteList').val()
        }
    };

    var data = await api.get(params);

    excel.addData({
        data: getRevisionArray(data),
/*        ranges: [
            {
                firstRow: 2,
                columns: ['D'],
                formula: 'A?*C?'
            },
            {
                firstRow: 2,
                columns: ['F'],
                formula: 'A?*E?'
            },
            {
                firstRow: 2,
                columns: ['C','D','E','F'],
                numberFormat: '0.00'
            },
            {
                firstRow: 2,
                columns: ['A'],
                numberFormat: '0'
            },
            {
                color: 'lightgrey'
            },
            {
                firstRow: 2,
                columns: ['A','B'],
                color: 'white'
            }
        ]
*/    });
}

async function getBomData(id) {

    var params = {
        path: 'bom-' + id,
        options: {id: $('#bomList').val()}
    };

    if (id == 'labor') params.options['laborId'] = parseInt($('#laborList').val());

    return await api.get(params);
}

async function onBomItems() {

    var data = await getBomData("items");

    excel.addData({
        data: getItemArray(data.items),
        ranges: [
            {
                firstRow: 2,
                columns: ['F'],
                formula: 'C?*E?'
            },
            {
                firstRow: 2,
                columns: ['L'],
                formula: 'ROUND(F?*(1+IF(K?="Yes",-1,1)*IF(ISNUMBER(J?),J?,' + data.defaultMU + ')/100),0)'
            },
            {
                firstRow: 2,
                columns: ['E','F','J'],
                numberFormat: '0.00'
            },
            {
                firstRow: 2,
                columns: ['C'],
                numberFormat: '0'
            },
            {
                color: 'lightgrey'
            },
            {
                firstRow: 2,
                columns: ['A','B','C','E','J'],
                color: 'white'
            }
        ]
    });
}

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

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getRevisionArray(data) {

    var dataArray = [
        ['Quote','','',0,'','GM','Sell'],
        ['MU (Default)','','',data.defaultMU,'',0,0],
        ['GM','','',0,'',0,0],
        ['MU','Cost','Sell',0,'',0,0],
        ['Labor',0,0,0,'',0,0],
        ['Material',0,0,0,'',0,0],
        ['Misc.',0,0,0,'',0,0],
        ['','','','','',0,0],
        ['Hours','','','','',0,0],
        ['Qty.','','Cost','Ext. Cost','',0,0],
        ['Total','','',0,'','',''],
        ['','','','','','',''],
        ['Quantity','Description','Cost','Ext. Cost','Quote','Ext. Quote','']
    ];

    var hours = getHoursArray(data.hours);

    if (hours.length) dataArray.splice(10, 0, hours);

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

    dataArray.push(['Total','','',0,'','',0]);   

    return dataArray;
}

function getBomArray(data) {

    var dataArray = [['Item *','Description','Quantity *','Units','Price','Amount','Vendor','Manufacturer','MPN','MU%','Discount','Quote']];

    for (var i = 0; i < data.length; i++) {
    
        var item = data[i];

        var quantity = (item.quantity ? parseInt(item.quantity) : 0);
        var price = (item.price ? parseFloat(item.price) : 0);
        var markup = (item.markup ? parseFloat(item.markup) : 0);
        var defaultMU = (item.defaultMU ? parseFloat(item.defaultMU) : 0);

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

    return dataArray;
}

function getExpenseArray(data) {

    var dataArray = [['Account *','Quantity *','Price *','Amount','MU%','Discount','Quote']];

    for (var i = 0; i < data.length; i++) {
    
        var expense = data[i];

        var quantity = (expense.quantity ? parseInt(expense.quantity) : 0);
        var price = (expense.price ? parseFloat(expense.price) : 0);
        var markup = (expense.markup ? parseFloat(expense.markup) : 0);
        var defaultMU = (expense.defaultMU ? parseFloat(expense.defaultMU) : 0);

        dataArray.push([
            expense.name,
            quantity,
            price,
            (quantity * price),
            (markup > 0 ? markup : ''),
            (expense.discount == 'T' ? 'Yes' : 'No'),
            (quantity * price * (1 + (markup > 0 ? markup : defaultMU)))
        ]);   
    }

    return dataArray;
}

function getLaborArray(data) {

    data.sort((a, b) => (a.name > b.name ? 1 : (b.name > a.name ? -1 : 0) ));

    var dataArray = [['Item','Quantity','Cost','Ext. Cost','Sell Price','Quote']];

    for (var i = 0; i < data.length; i++) {
    
        var labor = data[i];

        var quantity = (labor.quantity ? parseInt(labor.quantity) : 0);
        var price = (labor.price ? parseFloat(labor.price) : 0);

        dataArray.push([
            labor.name,
            labor.quantity,
            labor.cost,
            (quantity * price),
            labor.sellPrice,
            labor.quote
        ]);   
    }

    return dataArray;
}

function getHoursArray(data, summary = true) {

    var hoursArray = [];

    for (var i = 0; i < data.length; i++) {
    
        var hours = data[i];

        var array = [hours.quantity,hours.desc];   

        if (!summary) {

            array.push('','');   
        }

        array.push(hours.cost,0);   

        if (summary) {

            array.push('','','');   
        }
        else {

            array.push(0,0,0,0,0,0);   
        }

        hoursArray.push(array);   
    }

    return hoursArray;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onReload() {

    var buttonId = getButton();

    switch (buttonId) {
        case 'summary':
            reloadSummary();
            break;
        case 'overview':
            reloadOverview();
            break;
        case 'items':
            reloadItems();
            break;
        case 'expenses':
            reloadExpenses();
            break;
        case 'labor':
            reloadLabor();
            break;
    }
}

async function onSave() {

    var buttonId = getButton();

    switch (buttonId) {
        case 'summary':
            saveSummary();
            break;
        case 'overview':
            saveOverview();
            break;
        case 'items':
            saveItems();
            break;
        case 'expenses':
            saveExpenses();
            break;
        case 'labor':
            saveLabor();
            break;
    }
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
