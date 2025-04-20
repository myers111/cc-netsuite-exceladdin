/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const api = require('../modules/api.js');
const excel = require('../modules/excel.js');

const NEWITEM = 3757;
const HEADERS = [['Item *','Description','Quantity *','Units','Price','Amount','Vendor','Manufacturer','MPN','MU%','Discount','Quote']];

/* global console, document, Excel, Office */

var bomId = 0;

Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {

        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("reload").onclick = loadItems;
        document.getElementById("save").onclick = saveItems;

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

    var data = null;

    if (options.id) {
/*
        var params = {path: 'quote', 'options': options};

        var quote = await api.get(params);
    
        if (!quote) return;

        var revList = JSON.parse(quote.revList);

        var sel = $('#revisionList');

        for (var i = 0; i < revList.length; i++) {

            var rev = revList[i];

            var selected = (quote.revNo == -1 || quote.revNo == parseInt(rev.name));

            sel.append('<option value="' + rev.id + '"' + (selected ? ' selected' : '') + '>' + rev.name + '</option>');
        }

        if (quote.data) data = quote.data
*/
        await loadList('#revisionList', 'revisions');
    }
}

async function onRevisionChange(options) {

    var data = null;

    if (options.id) {
    
        //var params = {path: 'revision', 'options': options};

        //data = await api.get(params);
    
        await loadList('#bomList', 'boms');
    }
}
/*
async function loadRevisionData(data) {

    if (data) {

        document.getElementById("amount").value = data.amount;
        document.getElementById("defaultMU").value = data.defaultMU;
        document.getElementById("gm").value = data.gm;
        document.getElementById("mu").value = data.mu;
        document.getElementById("pctMaterial").value = data.pctMaterial;
        document.getElementById("pctLabor").value = data.pctLabor;
    }
    else {

        document.getElementById("amount").value = '';
        document.getElementById("defaultMU").value = '';
        document.getElementById("gm").value = '';
        document.getElementById("mu").value = '';
        document.getElementById("pctMaterial").value = '';
        document.getElementById("pctLabor").value = '';
    }
}
*/
async function onBomChange(options) {
return;
    try {

        if (options.id) {

            var params = {path: 'bom'};

            if (options) params['options'] = options;
        
            var data = await api.get(params);

            document.getElementById("quantity").value = data.quantity;
            document.getElementById("cost").value = data.cost;
            document.getElementById("extcost").value = data.extcost;
            document.getElementById("quote").value = data.quote;

            var items = [];

            for (var i = 0; i < data.items.length; i++) {
        
                var item = [];
        
                var row = data.items[i];
        
                item.push(row.quantity);
                item.push(row.item);
                item.push(row.desc);
                item.push(row.price);
                item.push(row.quantity*row.price);
                item.push(row.mu > 0 ? row.mu : '');
                item.push(row.quote);

                items.push(item);
            }
        
            excel.addData({
                sheetName: WS_ITEMS,
                data: items,
                sumColumns: ['E','G'],
                formulas: [
                    {
                        column: 'E',
                        formula: '=A?*D?'
                    },
                    {
                        column: 'G',
                        formula: '=ROUND(E?*(1+IF(ISNUMBER(F?),F?,' + options.defaultMU + ')/100),0)'
                    }
                ],
                formats: [
                    {
                        columns: ['D','E','F'],
                        format: '0.00'
                    },
                    {
                        columns: ['G'],
                        format: '0'
                    }
                ]
            });

            var expenses = [];

            for (var i = 0; i < data.expenses.length; i++) {
        
                var expense = [];
        
                var row = data.expenses[i];
        
                expense.push(row.quantity);
                expense.push(row.item);
                expense.push(row.desc);
                expense.push(row.price);
                expense.push(row.quantity*row.price);
                expense.push(row.quote);
        
                expenses.push(expense);
            }
        
            excel.addData({
                sheetName: WS_EXPENSES,
                data: expenses,
                sumColumns: ['E','F'],
                formulas: [
                    {
                        column: 'E',
                        formula: '=A?*D?'
                    }
                ],
                formats: [
                    {
                        columns: ['D','E'],
                        format: '0.00'
                    },
                    {
                        columns: ['F'],
                        format: '0'
                    }

                ]
            });
        }
        else {
            
            document.getElementById("quantity").value = '';
            document.getElementById("cost").value = '';
            document.getElementById("extcost").value = '';
            document.getElementById("quote").value = '';

            excel.clearData(WS_ITEMS);
            excel.clearData(WS_EXPENSES);
            excel.clearData(WS_LABOR);
        }
    }
    catch (error) {

        console.error(error);
    }
}

async function loadItems() {

    var params = {
        path: 'bom',
        options: {id: bomId}
    };

    var data = await api.get(params);

    if (!data) return;

    var dataArray = HEADERS;

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

async function addData(options) {

    try {

        await Excel.run(async (context) => {

            var sheet = context.workbook.worksheets.getActiveWorksheet();

            sheet.getRange().clear();

            var rangeString = getRangeString({
                firstRow: 1,
                columns: options.data[0].length,
                rows: options.data.length
            });

            sheet.getRange(rangeString).values = options.data;

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

            var rangeString = getRangeString({
                columns: options.data[0].length
            });

            sheet.getRange(rangeString).format.autofitColumns();

            await context.sync();
        });
    }
    catch (error) {

        console.error(error);
    }
}

function getRangeString(options) {

    var firstColumn = (options.firstColumn ? options.firstColumn : 'A');
    var firstRow = (options.firstRow ? options.firstRow : 1);

    return (firstColumn + (options.rows ? firstRow : '') + ':' + String.fromCharCode(firstColumn.charCodeAt(0) + options.columns - 1) + (options.rows ? firstRow + options.rows - 1 : ''));
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