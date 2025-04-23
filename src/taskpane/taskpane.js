/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const api = require('../modules/api.js');
const excel = require('../modules/excel.js');

const NEWITEM = 3757;
const ITEM_HEADERS = [['Item *','Description','Quantity *','Units','Price','Amount','Vendor','Manufacturer','MPN','MU%','Discount','Quote']];
const EXPENSE_HEADERS = [['Account *','Quantity *','Price *','Amount','MU%','Discount','Quote']];
const LABOR_HEADERS = [['Item *','Description','Quantity *','Units','Price','Amount','Vendor','Manufacturer','MPN','MU%','Discount','Quote']];

const BOM = {
    items: 0,
    expenses: 1,
    labor: 2
};

var BOM_ID = 0;
var BOM_SECTION = 0;

/* global console, document, Excel, Office */

Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {

        window.onload = onRefresh;

        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";

        document.getElementById("items").onclick = onItems;
        document.getElementById("expenses").onclick = onExpenses;
        document.getElementById("labor").onclick = onLabor;

        document.getElementById("reload").onclick = reload;
        document.getElementById("save").onclick = save;

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
                onQuoteChange({
                    id: ($('#quoteList').val() ? $('#quoteList').val() : 0)
                });
                break;
            case 'revisions':
                emptyList("#bomList");
                onRevisionChange({
                    id: ($('#revisionList').val() ? $('#revisionList').val() : 0),
                    quoteId: ($('#quoteList').val() ? $('#quoteList').val() : 0)
                });
                break;
            case 'boms':
                onBomChange({
                    id: ($('#bomList').val() ? $('#bomList').val() : 0),
                    revId: ($('#revisionList').val() ? $('#revisionList').val() : 0),
                    //defaultMU: $('#defaultMU').val().replace('%','')
                });
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

    switch (path) {
        case 'revisions':
        case 'boms':
            sel.prop("selectedIndex", 0).trigger('change');
            break;
        default:
            sel.prepend('<option value=""></option>');
            break;
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
            var revId = $('#revisionList').val();
            if (parseInt(revId)) options['revId'] = revId;
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

async function onQuoteChange(options) {

    if (options.id) await loadList('#revisionList', 'revisions');

    document.getElementById("revData").style.display = (options.id > 0 ? '' : 'none');
}

async function onRevisionChange(options) {

    if (options.id) {
    
        var params = {path: 'revision', 'options': options};

        var data = await api.get(params);
    
        document.getElementById("revAmount").value = data.amount.toLocaleString();
        document.getElementById("revDefaultMU").value = (data.defaultMU * 100).toFixed(2);

        await loadList('#bomList', 'boms');
    }
}

async function onBomChange(options) {

    document.getElementById("bomData").style.display = (options.id > 0 ? '' : 'none');

    BOM_ID = options.id;

    if (BOM_ID > 0)
        onItems();
    else
        onSummary();
}

async function onSummary() {

}

async function onItems() {

    BOM_SECTION = BOM.items;

    document.getElementById("items").className = "highLight";
    document.getElementById("expenses").className = "";
    document.getElementById("labor").className = "";

    var params = {
        path: 'bom-items',
        options: {id: BOM_ID}
    };

    var data = await api.get(params);

    if (!data) return;

    var dataArray = ITEM_HEADERS;

    if (data.items.length) dataArray = dataArray.concat(getItemArray(data.items));

    excel.addData({
        data: dataArray,
        //sumColumns: ['F','K'],
        formulas: [
            {
                column: 'F',
                formula: '=C?*E?'
            },
            {
                column: 'L',
                formula: '=ROUND(F?*(1+IF(K?="Yes",-1,1)*IF(ISNUMBER(J?),J?,' + data.defaultMU + ')/100),0)'
            }
        ],
        formats: [
            {
                columns: ['E','F','J'],
                format: '0.00'
            },
            {
                columns: ['C'],
                format: '0'
            }
        ]
    });
}

async function onExpenses() {

    BOM_SECTION = BOM.expenses;

    document.getElementById("items").className = "";
    document.getElementById("expenses").className = "highLight";
    document.getElementById("labor").className = "";

    var params = {
        path: 'bom-expenses',
        options: {id: BOM_ID}
    };

    var data = await api.get(params);

    if (!data) return;

    var dataArray = EXPENSE_HEADERS;

    if (data.expenses.length) dataArray = dataArray.concat(getExpenseArray(data.expenses));

    excel.addData({
        data: dataArray,
        //sumColumns: ['F','K'],
        formulas: [
            {
                column: 'D',
                formula: '=C?*E?'
            },
            {
                column: 'G',
                formula: '=ROUND(D?*(1+IF(F?="Yes",-1,1)*IF(ISNUMBER(E?),E?,' + data.defaultMU + ')/100),0)'
            }
        ],
        formats: [
            {
                columns: ['C','D','G'],
                format: '0.00'
            },
            {
                columns: ['B'],
                format: '0'
            }
        ]
    });
}

async function onLabor() {

    BOM_SECTION = BOM.labor;

    document.getElementById("items").className = "";
    document.getElementById("expenses").className = "";
    document.getElementById("labor").className = "highLight";

    var params = {
        path: 'bom-labor',
        options: {id: BOM_ID}
    };

    var data = await api.get(params);

    if (!data) return;

    var dataArray = LABOR_HEADERS;

    if (data.labor.length) dataArray = dataArray.concat(getItemArray(data.labor));

    excel.addData({
        data: dataArray,
        //sumColumns: ['F','K'],
        formulas: [
            {
                column: 'F',
                formula: '=C?*E?'
            },
            {
                column: 'L',
                formula: '=ROUND(F?*(1+IF(K?="Yes",-1,1)*IF(ISNUMBER(J?),J?,' + data.defaultMU + ')/100),0)'
            }
        ],
        formats: [
            {
                columns: ['E','F','J'],
                format: '0.00'
            },
            {
                columns: ['C'],
                format: '0'
            }
        ]
    });
}

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

function getItemArray(items) {

    var itemArray = [];

    for (var i = 0; i < items.length; i++) {
    
        var item = items[i];

        var quantity = (item.quantity ? parseInt(item.quantity) : 0);
        var price = (item.price ? parseFloat(item.price) : 0);
        var markup = (item.markup ? parseFloat(item.markup) : 0);
        var defaultMU = (item.defaultMU ? parseFloat(item.defaultMU) : 0);

        itemArray.push([
            (item.itemId == NEWITEM ? item.newItem : item.name),
            (item.itemId == NEWITEM ? item.newDescription : item.description),
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

    return itemArray;
}

function getExpenseArray(expenses) {

    var expenseArray = [];

    for (var i = 0; i < expenses.length; i++) {
    
        var expense = expenses[i];

        var quantity = (expense.quantity ? parseInt(expense.quantity) : 0);
        var price = (expense.price ? parseFloat(expense.price) : 0);
        var markup = (expense.markup ? parseFloat(expense.markup) : 0);
        var defaultMU = (expense.defaultMU ? parseFloat(expense.defaultMU) : 0);

        expenseArray.push([
            expense.name,
            quantity,
            price,
            (quantity * price),
            (markup > 0 ? markup : ''),
            (expense.discount == 'T' ? 'Yes' : 'No'),
            (quantity * price * (1 + (markup > 0 ? markup : defaultMU)))
        ]);   
    }

    return expenseArray;
}

function getLaborArray(items) {

    var itemArray = [];

    for (var i = 0; i < items.length; i++) {
    
        var item = items[i];

        var quantity = (item.quantity ? parseInt(item.quantity) : 0);
        var price = (item.price ? parseFloat(item.price) : 0);
        var markup = (item.markup ? parseFloat(item.markup) : 0);
        var defaultMU = (item.defaultMU ? parseFloat(item.defaultMU) : 0);

        itemArray.push([
            (item.itemId == NEWITEM ? item.newItem : item.name),
            (item.itemId == NEWITEM ? item.newDescription : item.description),
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

    return itemArray;
}