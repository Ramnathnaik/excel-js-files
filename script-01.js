/* 
    This project is for Tableau Excel download
    Product developed by LTIMindtree
*/

'use strict';

window.onload = function () {
    //FUNCTION RUNS ON BUTTON CLICK
    document.getElementById("demo").onclick = () => {
        //TABLEAU EXTENSION API
        window.top.openReportPopup();
        window.top.Spinner.show();
        window.top.Mask.show();
        tableau.extensions.initializeAsync().then(function () {
            let dashboard = tableau.extensions.dashboardContent.dashboard;

            setTimeout(() => {
                //PROMISE
                Promise.all([processDashboard(dashboard)]).then((values) => {
                    if (values != 'error') {
                        console.log('Excel Export - V6');
                        // "FORCE DOWNLOAD" XLSX FILE
                        var today = new Date();
                        var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate();
                        var time = today.getHours() + "_" + today.getMinutes() + "_" + today.getSeconds();
                        // var dateTime = date + ' ' + time;

                        var excelFileName = dashboard.name;

                        window.top.Mask.hide();
                        window.top.Spinner.hide();
                        // XLSX.writeFile(workbook, excelFileName + ".xlsx");
                    }
                });
            }, 2000);




        });
    }

    document.getElementById('download').onclick = () => {
        // CREATE NEW EXCEL "FILE"
        var workbook = XLSX.utils.book_new();

        let worksheetArr = window.top.x;

        worksheetArr.forEach((worksheetInfo) => {
            worksheetInfo.name = worksheetInfo.name.length >= 31 ? worksheetInfo.name.substring(0, 30) : worksheetInfo.name;
            workbook.SheetNames.push(worksheetInfo.name);
            workbook.Sheets[worksheetInfo.name] = worksheetInfo.worksheet;
        });

        XLSX.writeFile(workbook, "Test" + ".xlsx");
    }
}

function fitToColumn(arrayOfArray) {
    // get maximum character of each column
    return arrayOfArray[0].map((a, i) => ({ wch: Math.max(...arrayOfArray.map(a2 => a2[i] ? a2[i].toString().length : 0)) }));
}

//find whether object with specific value is present in array of objects
function getIndex(arr, name) {
    const { length } = arr;
    const id = length + 1;
    return arr.findIndex(el => el.fieldName === name);
}

//find whether object with specific value is present in array of objects
function getIndexUsingStartsWith(arr, name) {
    const { length } = arr;
    const id = length + 1;
    return arr.findIndex(el => el.fieldName.startsWith(name));
}

//returns an array of elements with includes given name
function getIncludedArr(arr, name) {
    return arr.filter(x => x.fieldName.includes(name)).map(x => x.fieldName);
}

//returns an array by removing duplicate elements
function removeDuplicates(arr) {
    return arr.filter((item,
        index) => arr.indexOf(item) === index);
}

function processDashboard(dashboard) {
    //DECLARE REQUIRED OBJECTS FOR STYLEJS
    const DEF_Size14Vert = { font: { sz: 24 }, alignment: { vertical: 'center', horizontal: 'center' } };
    const DEF_FxSz14RgbVert = { font: { name: 'Calibri', sz: 11, color: { rgb: '000000' } }, alignment: { vertical: 'center', horizontal: 'center' } };
    let detailsWorksheet;

    return new Promise(async function (resolve, reject) {
        let arr = dashboard.worksheets;

        let worksheetArr = [];

        let checkCount = 0;

        let dashboardFilters = [];

        for (let object of dashboard.objects) {
            if (object.type === 'quick-filter')
                dashboardFilters.push(object.name);
        }

        console.log(dashboardFilters);

        let worksheetsToBeExtracted = arr.reduce((accumulator, obj) => {
            if (obj.name.includes('Report_Export_Details_D')) {
                return accumulator + 1;
            }
            return accumulator;
        }, 0);

        let sheetName = '';
        let reportHeader = '';
        let reportRefreshTime = '';
        let reportFooter = '';
        let user = '';
        let groupsParams = '';
        let setsParams = '';
        let p = '';
        let f = '';
        let result = [];
        let totalRowCount = 0;
        let columnLength = 0;
        let filtersCounter = 0;
        let parametersCounter = 0;
        let parameters = [];

        await dashboard.getParametersAsync().then(async function (rawParameters) {
            for (let rawParameter of rawParameters) {
                parameters.push({
                    'parameterName': rawParameter.name, 
                    'parameterValue': rawParameter.currentValue.formattedValue
                });
                parametersCounter++;
            }
        });

        await dashboard.worksheets.forEach(async function (worksheet, key, arr) {
            if (worksheet.name.includes('Report_Export_Details_D')) {
                detailsWorksheet = worksheet;
                await detailsWorksheet.getSummaryDataAsync().then(async function (mydata) {
                    let dashboardData = mydata.data;
                    let dashboardColumns = mydata.columns;

                    if (getIndexUsingStartsWith(dashboardColumns, 'Sheet name') != -1) {
                        sheetName = dashboardData[0][getIndexUsingStartsWith(dashboardColumns, 'Sheet name')].value;
                    }

                    if (getIndexUsingStartsWith(dashboardColumns, 'Report Header') != -1) {
                        reportHeader = dashboardData[0][getIndexUsingStartsWith(dashboardColumns, 'Report Header')].value;
                    }

                    if (getIndex(dashboardColumns, 'Report Refresh Time') != -1) {
                        reportRefreshTime = dashboardData[0][getIndex(dashboardColumns, 'Report Refresh Time')].value;
                    }

                    if (getIndexUsingStartsWith(dashboardColumns, 'Report Footer') != -1) {
                        reportFooter = dashboardData[0][getIndexUsingStartsWith(dashboardColumns, 'Report Footer')].value;
                    }

                    if (getIndex(dashboardColumns, 'User') != -1) {
                        user = dashboardData[0][getIndex(dashboardColumns, 'User')].value;
                    }

                    //let sheetOrder = dashboardData[0][getIndex(dashboardColumns, 'Sheet order')].value;


                    if (getIndex(dashboardColumns, 'Groups Parameter') != -1) {
                        groupsParams = dashboardData[0][getIndex(dashboardColumns, 'Groups Parameter')].value;
                    }

                    if (getIndex(dashboardColumns, 'Sets Parameter') != -1) {
                        setsParams = dashboardData[0][getIndex(dashboardColumns, 'Sets Parameter')].value;
                    }

                    let paramsArr = getIncludedArr(dashboardColumns, 'Param');
                    paramsArr.forEach(param => {
                        p += dashboardData[0][getIndex(dashboardColumns, param)].value + ';  ';
                    });

                    let filtersArr = getIncludedArr(dashboardColumns, 'Filter');
                    filtersArr.forEach(filter => {
                        f += dashboardData[0][getIndex(dashboardColumns, filter)].value + ';  ';
                    });

                    await dashboard.worksheets.forEach(async function (sheet) {
                        if (sheet.name === sheetName) {
                            let filters = [];

                            await sheet.getFiltersAsync().then(async function (mydata) {
                                let rawFilters = mydata;
                                console.log(sheet.name);
                                console.log(rawFilters);
                                if (rawFilters.length > 0) {
                                    for (let rawFilter of rawFilters) {
                                        if (dashboardFilters.includes(rawFilter.fieldName)) {
                                            let appliedValues = rawFilter.appliedValues || [];
                                            let rawValues = [];
                                            if (appliedValues.length > 0) {
                                                for (let appliedValue of appliedValues) {
                                                    rawValues.push(appliedValue.formattedValue);
                                                }
                                            }
                                            let tempObj = {
                                                'fieldName': rawFilter.fieldName,
                                                'filterValues': rawValues
                                            }
                                            filtersCounter++;
                                            filters.push(tempObj);
                                        }
                                    }
                                }
                            });
                            await sheet.getSummaryDataAsync().then(async function (d) {
                                if (checkCount == 0) {
                                    let sheetData = d;

                                    // console.log(sheet.name);
                                    // console.log(filters);

                                    checkCount++;
                                    // console.log(sheetData);
                                    columnLength = sheetData.columns.length;
                                    let columns = sheetData.columns;
                                    let slNoIndex = -1;
                                    let emptyColIndex = -1;

                                    /* Excel data type map */
                                    let definedExcelDataTypeMap = {
                                        'string': 's',
                                        'date': 'd',
                                        'int': 'n',
                                        'float': 'n',
                                        'date-time': 'd'
                                    };

                                    let columnDataTypeMap = {};

                                    /* Check whether column as Measure Names and Measure values field.
                                    If present, find the index */
                                    let measureNamesIndex = -1;
                                    let measureValuesIndex = -1;

                                    for (let i = 0; i < columnLength; i++) {
                                        let colEle = columns[i];
                                        if (colEle.fieldName === 'Measure Names') {
                                            measureNamesIndex = i;
                                        } else if (colEle.fieldName === 'Measure Values') {
                                            measureValuesIndex = i;
                                        }

                                        /* Get Sl_No index */
                                        if (colEle.fieldName === 'AGG(Sl_No)') {
                                            slNoIndex = i;
                                        }

                                        /* Get the empty column index */
                                        if (colEle.fieldName.trim() === "' '") {
                                            emptyColIndex = i;
                                        }

                                        /* Get the data type of each column and populate into map */
                                        columnDataTypeMap[i] = colEle.dataType;
                                    }

                                    /* If measure names are present, count how much measure names are present */
                                    let colData = sheetData.data;
                                    let measureNames = [];
                                    let mCount = 1;

                                    if (measureNamesIndex != -1) {
                                        // let mFlag = false;
                                        let mIndex = -1;
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];

                                            if (mIndex == -1) {
                                                for (let j = 0; j < arrEle.length; j++) {
                                                    if (measureNamesIndex != j || measureValuesIndex != j) {
                                                        mIndex = j;
                                                        break;
                                                    }
                                                }
                                            }
                                            if (mIndex != -1) {
                                                if (colData[i]?.[mIndex].value === colData[i + 1]?.[mIndex].value) {
                                                    // mCount++;
                                                    measureNames.push(colData[i][measureNamesIndex].formattedValue);
                                                    measureNames.push(colData[i + 1][measureNamesIndex].formattedValue);
                                                } else {
                                                    break;
                                                }
                                            }

                                        }

                                        if (measureNames.length == 0) {
                                            measureNames.push(colData[0][measureNamesIndex].formattedValue);
                                        }
                                    }

                                    measureNames = removeDuplicates(measureNames);
                                    mCount = measureNames.length;

                                    console.log(measureNames);
                                    console.log(mCount);

                                    let tt = [];
                                    let rr = [];
                                    let empt = [];

                                    let actualColumnLength = columnLength;
                                    columnLength = measureNames.length > 0 ? columnLength - 2 + mCount : columnLength;
                                    columnLength = slNoIndex == -1 ? columnLength : columnLength - 1;
                                    columnLength = emptyColIndex == -1 ? columnLength : columnLength - 1;

                                    for (let i = 0; i < columnLength; i++) {
                                        if (i == 0) {
                                            tt.push({ v: reportHeader, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 14, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'left', vertical: 'center' } } });
                                        } else {
                                            tt.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 22, name: 'Calibri', color: { rgb: 'f1f1f1' } } } });
                                        }
                                        if (i == 0) {
                                            rr.push({ v: `Report executed by ${user} ${reportRefreshTime}`, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'left' } } });
                                        } else {
                                            rr.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '538DD5' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { horizontal: 'right' } } });
                                        }
                                        empt.push(" ");
                                    }

                                    result.push(tt);
                                    result.push(empt);
                                    result.push(rr);

                                    if (filters.length > 0) {
                                        for (let filter of filters) {
                                            tt = [];
                                            for (let i = 0; i < columnLength; i++) {
                                                if (i == 0)
                                                    tt.push({ v: filter.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                                else if (i == 1)
                                                    tt.push({ v: filter.filterValues, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                                else
                                                    tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            }
                                            result.push(tt);
                                        }
                                    }

                                    if (parameters.length > 0) {
                                        for (let parameter of parameters) {
                                            tt = [];
                                            for (let i = 0; i < columnLength; i++) {
                                                if (i == 0)
                                                    tt.push({ v: parameter.parameterName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                                else if (i == 1)
                                                    tt.push({ v: parameter.parameterValue, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                                else
                                                    tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            }
                                            result.push(tt);
                                        }
                                    }

                                    if (p != '') {
                                        tt = [];
                                        for (let i = 0; i < columnLength; i++) {
                                            if (i == columnLength - 2) {
                                                tt.push({ v: p, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            } else {
                                                tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            }
                                        }
                                        result.push(tt);
                                    }

                                    if (f != '') {
                                        tt = [];
                                        for (let i = 0; i < columnLength; i++) {
                                            if (i == columnLength - 2) {
                                                tt.push({ v: f, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            } else {
                                                tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            }
                                        }
                                        result.push(tt);
                                    }

                                    if (groupsParams != '') {
                                        tt = [];
                                        for (let i = 0; i < columnLength; i++) {
                                            if (i == columnLength - 2) {
                                                tt.push({ v: groupsParams, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            } else {
                                                tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            }
                                        }
                                        result.push(tt);
                                    }

                                    if (setsParams != '') {
                                        tt = [];
                                        for (let i = 0; i < columnLength; i++) {
                                            if (i == columnLength - 2) {
                                                tt.push({ v: setsParams, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            } else {
                                                tt.push({ v: '', t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'right' } } });
                                            }
                                        }
                                        result.push(tt);
                                    }

                                    result.push(empt);
                                    result.push(empt);

                                    tt = [];
                                    if (measureNames.length > 0) {
                                        for (let i = 0; i < actualColumnLength; i++) {
                                            if ((i != measureNamesIndex) && (i != measureValuesIndex) && (i != slNoIndex) && (i != emptyColIndex)) {
                                                let colEle = columns[i];

                                                tt.push({ v: ((colEle.fieldName.startsWith('SUM(') || colEle.fieldName.startsWith('AGG(') || colEle.fieldName.startsWith('ATTR(')) && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(4, colEle.fieldName.length - 1) : (colEle.fieldName.startsWith('ATTR(') && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(5, colEle.fieldName.length - 1) : colEle.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left' } } });
                                            }
                                        }
                                        for (let i = 0; i < measureNames.length; i++) {
                                            tt.push({ v: measureNames[i], t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left' } } });
                                        }
                                    } else {
                                        for (let i = 0; i < columnLength; i++) {
                                            let colEle = columns[i];

                                            if ((i != slNoIndex) && (i != i != emptyColIndex)) {
                                                tt.push({ v: ((colEle.fieldName.startsWith('SUM(') || colEle.fieldName.startsWith('AGG(')) && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(4, colEle.fieldName.length - 1) : (colEle.fieldName.startsWith('ATTR(') && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(5, colEle.fieldName.length - 1) : colEle.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left' } } });
                                            }
                                        }
                                    }

                                    result.push(tt);

                                    if (measureNames.length > 0) {
                                        let lCount = mCount;
                                        let tempDict = {};
                                        let tempArr = [];
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];

                                            if (lCount != 0) {
                                                for (let j = 0; j < arrEle.length; j++) {
                                                    if ((j != measureNamesIndex) && (j != measureValuesIndex) && (j != slNoIndex) && (j != emptyColIndex) && (lCount == mCount)) {
                                                        tempArr.push({ v: arrEle[j].value == '%null%' ? '' : columnDataTypeMap[j] === 'date' || columnDataTypeMap[j] === 'date-time' ? arrEle[j].formattedValue.substring(0, arrEle[j].formattedValue.indexOf(" ") === -1 ? arrEle[j].formattedValue.length : arrEle[j].formattedValue.indexOf(" ")) : arrEle[j].value, t: arrEle[j].value == '%null%' ? 's' : definedExcelDataTypeMap?.[columnDataTypeMap[j]] ? definedExcelDataTypeMap?.[columnDataTypeMap[j]] : isNaN(arrEle[j].value) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(arrEle[j].value) ? { horizontal: 'left' } : { horizontal: 'right' } } });
                                                    }
                                                }
                                                tempDict[arrEle[measureNamesIndex].formattedValue] = arrEle[measureValuesIndex].value;
                                                lCount--;
                                            }

                                            if (lCount == 0) {
                                                for (let j = 0; j < measureNames.length; j++) {
                                                    let tempData = tempDict[measureNames[j]];
                                                    tempArr.push({ v: tempData == '%null%' ? '' : tempData, t: isNaN(tempData) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(tempData) ? { horizontal: 'left' } : { horizontal: 'right' } } });
                                                }

                                                result.push(tempArr);
                                                totalRowCount++;
                                                tempArr = [];
                                                tempDict = {};
                                                lCount = mCount;
                                            }

                                        }
                                    } else {
                                        for (let i = 0; i < colData.length; i++) {
                                            let arrEle = colData[i];
                                            let tempArr = [];
                                            for (let j = 0; j < arrEle.length; j++) {
                                                if ((j != slNoIndex) && (j != emptyColIndex)) {
                                                    tempArr.push({ v: arrEle[j].value == '%null%' ? '' : arrEle[j].value, t: arrEle[j].value == '%null%' ? 's' : definedExcelDataTypeMap?.[columnDataTypeMap[j]] ? definedExcelDataTypeMap?.[columnDataTypeMap[j]] : isNaN(arrEle[j].value) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(arrEle[j].value) ? { horizontal: 'left' } : { horizontal: 'right' } } });
                                                }
                                            }
                                            result.push(tempArr);
                                            totalRowCount++;
                                        }
                                    }
                                } else {
                                    let sheetData = d;

                                    checkCount++;
                                    // console.log(sheetData);
                                    let sheetColumnLength = sheetData.columns.length;
                                    let columns = sheetData.columns;

                                    /* Excel data type map */
                                    let definedExcelDataTypeMap = {
                                        'string': 's',
                                        'date': 'd',
                                        'int': 'n',
                                        'float': 'n',
                                        'date-time': 'd'
                                    };

                                    let columnDataTypeMap = {};

                                    for (let i = 0; i < sheetColumnLength; i++) {
                                        let colEle = columns[i];

                                        /* Get the data type of each column and populate into map */
                                        columnDataTypeMap[i] = colEle.dataType;
                                    }

                                    let colData = sheetData.data;

                                    let empt = [];
                                    let tt = [];

                                    for (let i = 0; i < sheetColumnLength; i++) {
                                        empt.push(" ");
                                    }

                                    result.push(empt);

                                    // for (let i = 0; i < sheetColumnLength; i++) {
                                    //     let colEle = columns[i];
                                    //     tt.push({ v: ((colEle.fieldName.startsWith('SUM(') || colEle.fieldName.startsWith('AGG(')) && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(4, colEle.fieldName.length - 1) : (colEle.fieldName.startsWith('ATTR(') && colEle.fieldName.endsWith(')')) ? colEle.fieldName.substring(5, colEle.fieldName.length - 1) : colEle.fieldName, t: 's', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, font: { sz: 11, name: 'Calibri', bold: true }, alignment: { horizontal: 'left' } } });
                                    // }

                                    // result.push(tt);

                                    for (let i = 0; i < colData.length; i++) {
                                        let arrEle = colData[i];
                                        let tempArr = [];
                                        for (let j = 0; j < arrEle.length; j++) {
                                            tempArr.push({ v: arrEle[j].value == '%null%' ? '' : arrEle[j].value, t: arrEle[j].value == '%null%' ? 's' : definedExcelDataTypeMap?.[columnDataTypeMap[j]] ? definedExcelDataTypeMap?.[columnDataTypeMap[j]] : isNaN(arrEle[j].value) ? 's' : 'n', s: { ...DEF_FxSz14RgbVert, border: { right: { style: 'thin', color: { rgb: '000000' } }, left: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } }, top: { style: 'thin', color: { rgb: '000000' } } }, alignment: isNaN(arrEle[j].value) ? { horizontal: 'left' } : { horizontal: 'right' } } });

                                        }
                                        result.push(tempArr);
                                        totalRowCount++;
                                    }

                                }

                                if (checkCount == worksheetsToBeExtracted) {
                                    let tt = [];
                                    let empt = [];
                                    for (let i = 0; i < columnLength; i++) {
                                        if (i == 0) {
                                            tt.push({ v: reportFooter, t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { vertical: 'center', horizontal: 'left' } } });
                                        } else {
                                            tt.push({ v: ' ', t: 's', s: { ...DEF_Size14Vert, fill: { fgColor: { rgb: '404040' } }, font: { sz: 11, name: 'Calibri', color: { rgb: 'f1f1f1' } }, alignment: { vertical: 'bottom', horizontal: 'center' } } });
                                        }
                                        empt.push(" ");
                                    }

                                    if (worksheetsToBeExtracted === 1) {
                                        result.push(empt);
                                    }
                                    result.push(empt);
                                    result.push(tt);

                                    //CREATE WORKSHEET(S) AND ADD IT TO EXCEL FILE
                                    let worksheet = XLSX.utils.aoa_to_sheet(result);

                                    let rowFooterMergeStart = 8 + totalRowCount;
                                    rowFooterMergeStart = groupsParams != '' ? rowFooterMergeStart + 1 : rowFooterMergeStart;
                                    rowFooterMergeStart = setsParams != '' ? rowFooterMergeStart + 1 : rowFooterMergeStart;
                                    rowFooterMergeStart = p != '' ? rowFooterMergeStart + 1 : rowFooterMergeStart;
                                    rowFooterMergeStart = f != '' ? rowFooterMergeStart + 1 : rowFooterMergeStart;
                                    rowFooterMergeStart = filtersCounter !== 0 ? rowFooterMergeStart + filtersCounter : rowFooterMergeStart;
                                    rowFooterMergeStart = parametersCounter !== 0 ? rowFooterMergeStart + parametersCounter : rowFooterMergeStart;

                                    worksheet['!cols'] = fitToColumn(result);
                                    worksheet['!rows'] = [{ 'hpt': 40 }];
                                    worksheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: columnLength - 1 } },
                                    { s: { r: rowFooterMergeStart, c: 0 }, e: { r: rowFooterMergeStart + 2, c: columnLength - 1 } }
                                    ];

                                    worksheet["!merges"].push({ s: { r: 2, c: 0 }, e: { r: 2, c: columnLength - 1 } });
                                    worksheet["!merges"] = p != '' ? [...worksheet["!merges"], { s: { r: 3, c: columnLength - 2 }, e: { r: 3, c: columnLength - 1 } }] : worksheet["!merges"];
                                    worksheet["!merges"] = f != '' ? [...worksheet["!merges"], { s: { r: 4, c: columnLength - 2 }, e: { r: 4, c: columnLength - 1 } }] : worksheet["!merges"];
                                    worksheet["!merges"] = groupsParams != '' ? [...worksheet["!merges"], { s: { r: 5, c: columnLength - 2 }, e: { r: 5, c: columnLength - 1 } }] : worksheet["!merges"];
                                    worksheet["!merges"] = setsParams != '' ? [...worksheet["!merges"], { s: { r: 6, c: columnLength - 2 }, e: { r: 6, c: columnLength - 1 } }] : worksheet["!merges"];

                                    let obj = {
                                        //index: sheetOrder,
                                        name: sheetName,
                                        worksheet: worksheet
                                    }

                                    //Stringfy the result
                                    if (!window.top.x) {
                                        window.top.x = [obj];
                                    } else {
                                        window.top.x = [obj, ...window.top.x];
                                    }
                                    resolve();
                                }
                            });
                        }
                    });
                });
            }
        });
    });
}