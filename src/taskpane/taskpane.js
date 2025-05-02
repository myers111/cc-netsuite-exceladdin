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

        document.getElementById("summary").onclick = onRevisionSummary;
        document.getElementById("overview").onclick = onRevisionOverview;
        document.getElementById("items").onclick = onBomItems;
        document.getElementById("expenses").onclick = onBomExpenses;
        document.getElementById("labor").onclick = onBomLabor;
        //document.getElementById("reload").onclick = onReload;
        //document.getElementById("save").onclick = onSave;

        initList('customer');
        initList('project');
        initList('quote');
        initList('revision');
        initList('bom');

        excel.initialize({
            excel: Excel
        });
    }
});

async function  onRefresh() {

    excel.clearData();
}

async function initList(id) {

    var path = id + 's';
    var selector = '#' + id + 'List';

    switch (id) {
        case 'revision':
        case 'bom':
            break;
        default:
        {
            $(selector).focus(function() {
            
                loadList(selector, path);
            });
        }
    }

    $(selector).on('change', function() {

        switch (path) {
            case 'customers':
                emptyList("#projectList");
            case 'projects':
                emptyList("#quoteList");
            case 'quotes':
                emptyList("#revisionList");
                emptyList("#bomList");
                onQuoteChange();
                break;
            case 'revisions':
                emptyList("#bomList");
                onRevisionChange();
                break;
            case 'boms':
                onBomChange();
                break;
        } 
    });
}

async function loadList(selector, path) {

    if ($(selector + ' option').length) return;

    var sel = $(selector);

    var params = {path: path};

    var options = getOptions(path);

    if (options) params['options'] = options;

    var data = await api.get(params);

    if (data.length == 0) return;

    for (var i = 0; i < data.length; i++) {

        sel.append('<option value="' + data[i].id + '">' + data[i].name + '</option>');
    }

    if (path != 'revisions')  sel.prepend('<option value=""></option>');

    switch (path) {
        case 'revisions':
        case 'boms':
            sel.prop("selectedIndex", 0).trigger('change');
    }
}

async function emptyList(selector) {

    var sel = $(selector);

    sel.empty();
}

function getOptions(path) {

    var options = {};

    switch (path) {
        case 'boms':
            var revisionId = $('#revisionList').val();
            if (parseInt(revisionId)) options['revisionId'] = revisionId;
        case 'revisions':
            var quoteId = $('#quoteList').val();
            if (parseInt(quoteId)) options['quoteId'] = quoteId;
        case 'quotes':
            var projectId = $('#projectList').val();
            if (parseInt(projectId)) options['projectId'] = projectId;
        case 'projects':
            var customerId = $('#customerList').val();
            if (parseInt(customerId)) options['customerId'] = customerId;
    }

    return (Object.keys(options).length == 0 ? null : options);
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

async function onQuoteChange() {

    var quoteId = $('#quoteList').val();

    if (quoteId) {
        
        await loadList('#revisionList', 'revisions');
    }

    document.getElementById("revisionControls").style.display = (quoteId > 0 ? '' : 'none');
    document.getElementById("revisionButtons").style.display = (quoteId > 0 ? '' : 'none');

    document.getElementById("controls").style.display = (quoteId > 0 ? '' : 'none');
}

async function onRevisionChange() {

    onRevisionSummary();

    await loadList('#bomList', 'boms');
}

async function getRevisionData(buttonId) {

    setButton(buttonId);

    var params = {
        path: 'revision-' + buttonId,
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

async function onBomChange() {

    var bomOptionCount = $('#bomList option').length;

    document.getElementById("bomControls").style.display = (bomOptionCount > 1 ? '' : 'none');

    var bomId = $('#bomList').val();

    document.getElementById("bomButtons").style.display = (bomId > 0 ? '' : 'none');

    if (bomId > 0)
        onBomItems();
    else
        onRevisionSummary();
}

async function getBomData(buttonId) {

    setButton(buttonId);

    var params = {
        path: 'bom-' + buttonId,
        options: {id: $('#bomList').val()}
    };

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
     });
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getSummaryArray(items) {

    var dataArray = [['Quantity','Description','Cost','Ext. Cost','Quote','Ext. Quote']];

    for (var i = 0; i < items.length; i++) {
    
        var item = items[i];

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

function getOverviewArray(items) {

    var dataArray = [['Quantity','Description','Cost','Ext. Cost','Quote','Ext. Quote']];

    for (var i = 0; i < items.length; i++) {
    
        var item = items[i];

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

function getItemArray(items) {

    var dataArray = [['Item *','Description','Quantity *','Units','Price','Amount','Vendor','Manufacturer','MPN','MU%','Discount','Quote']];

    for (var i = 0; i < items.length; i++) {
    
        var item = items[i];

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

function getExpenseArray(expenses) {

    var dataArray = [['Account *','Quantity *','Price *','Amount','MU%','Discount','Quote']];

    for (var i = 0; i < expenses.length; i++) {
    
        var expense = expenses[i];

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

function getLaborArray(items) {

    var dataArray = [['Item *','Description','Quantity *','Units','Price','Amount','Vendor','Manufacturer','MPN','MU%','Discount','Quote']];

    for (var i = 0; i < items.length; i++) {
    
        var item = items[i];

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

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function setButton(id) {

    if (id == "summary" || id == "overview") document.getElementById("bomButtons").style.display = 'none';

    resetButtons();

    document.getElementById(id).className = HIGHLIGHT_CLASS;
}

function resetButtons() {

    document.getElementById("summary").className = "";
    document.getElementById("overview").className = "";
    document.getElementById("items").className = "";
    document.getElementById("expenses").className = "";
    document.getElementById("labor").className = "";
}

function getButton() {

    return $("#bomControls" + HIGHLIGHT_CLASS).attr("id");
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
