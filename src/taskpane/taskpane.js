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

        initList('customer');
        initList('project');
        initList('quote');
        initList('revision');
        initList('bom');
        initList('labor');

        initButton("summary");
        initButton("overview");
        initButton("items");
        initButton("expenses");
        initButton("labor");

        excel.initialize({
            excel: Excel
        });
    }
});

async function  onRefresh() {

    excel.clearData();
}

async function initList(id) {

    var selector = '#' + id + 'List';

    switch (id) {
        case 'customer':
        case 'project':
        case 'quote':
            $(selector).focus(function() { loadList(id); });
            break;
        default:
    }

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

    for (var i = 0; i < data.length; i++) {

        sel.append('<option value="' + data[i].id + '">' + data[i].name + '</option>');
    }

    switch (id) {
        case 'customer':
        case 'project':
        case 'quote':
        case 'bom':
            sel.prepend('<option value=""></option>');
            break;
        default:                
    }

    sel.prop("selectedIndex", 0).trigger('change');
}

async function emptyList(id) {

    $('#' + id + 'List').empty();
}

function getOptions(id) {

    var options = {};

    switch (id) {
        case 'labor':
            var bomId = $('#bomList').val();
            if (parseInt(bomId)) options['bomId'] = bomId;
        case 'bom':
            var revisionId = $('#revisionList').val();
            if (parseInt(revisionId)) options['revisionId'] = revisionId;
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

    var buttonId = getButton();

    switch (id) {
        case 'customer':
            emptyList("project");
        case 'project':
            emptyList("quote");
        case 'quote':
            {
                emptyList("revision");
                emptyList("bom");
                emptyList("labor");

                var quoteId = $('#quoteList').val();

                if (quoteId) {

                    await loadList('revision');

                    onClick('summary');
                }
                else {

                    excel.clearData();
                }
                document.getElementById("revisionControls").style.display = (quoteId > 0 ? '' : 'none');
                document.getElementById("bomButtons").style.display = 'none';
                document.getElementById("laborControls").style.display = 'none';
            }
            break;
        case 'revision':
            emptyList("bom");
            emptyList("labor");
            await loadList('bom');
            onRevisionSummary();
            break;
        case 'bom':
            emptyList("labor");
            var bomId = $('#bomList').val();
            document.getElementById("bomButtons").style.display = (bomId > 0 ? '' : 'none');
            onClick(bomId > 0 ? 'items' : 'summary');
            break;
        case 'labor':
            onBomLabor();
            break;
    }
}

async function onClick(id) {

    var buttonId = getButton();

    if (id == buttonId) return;

    if (buttonId) document.getElementById(buttonId).className = "";

    switch (id) {
        case 'summary':
            document.getElementById("bomButtons").style.display = 'none';
            onRevisionSummary();
            break;
        case 'overview':
            document.getElementById("bomButtons").style.display = 'none';
            onRevisionOverview();
            break;
        case 'items':
            onBomItems();
            break;
        case 'expenses':
            onBomExpenses();
            break;
        case 'labor':
            await loadList('labor');
            onBomLabor();
            break;
    }

    document.getElementById("laborControls").style.display = (id == 'labor' ? '' : 'none');

    document.getElementById(id).className = HIGHLIGHT_CLASS;
}

function getButton() {

    return $("button." + HIGHLIGHT_CLASS).attr("id");
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function getRevisionData(id) {

    var params = {
        path: 'revision-' + id,
        options: {
            id: $('#revisionList').val(),
            quoteId: $('#quoteList').val()
        }
    };

    return await api.get(params);
}

async function onRevisionSummary() {

    $('#bomList').val(0);

    var data = await getRevisionData("summary");

    excel.addData({
        data: getSummaryArray(data),
        ranges: [
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
    });
}

async function onRevisionOverview() {

    var data = await api.get({
        path: 'revision-overview',
        options: {
            id: $('#revisionList').val(),
            quoteId: $('#quoteList').val()
        }
    });
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
                formula: 'C?*E?'
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

function getSummaryArray(data) {

    var dataArray = [['Quantity','Description','Cost','Ext. Cost','Quote','Ext. Quote']];

    for (var i = 0; i < data.length; i++) {
    
        var item = data[i];

        dataArray.push([
            item.quantity,
            item.desc,
            item.cost,
            item.quantity * item.cost,
            item.quote,
            item.quantity * item.quote
        ]);   
    }

    return dataArray;
}

function getOverviewArray(data) {

    var dataArray = [['Quantity','Description','Cost','Ext. Cost','Quote','Ext. Quote']];

    for (var i = 0; i < data.length; i++) {
    
        var item = data[i];

        dataArray.push([
            item.quantity,
            item.desc,
            item.cost,
            item.quantity * item.cost,
            item.quote,
            item.quantity * item.quote
        ]);   
    }

    return dataArray;
}

function getItemArray(data) {

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

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function saveItems() {

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
                path: 'bom',
                options: {bom: bom}
            };

            await api.post(params);
        });
    }
    catch (error) {

        console.error(error);
    }
}
