var onSusRender = async function(dtName, fieldOnList, fieldQuery, secondFieldonTable, secondList){

    await filterDataTable_LookupField(dtName, fieldOnList, fieldQuery, secondList, secondFieldonTable);

    fixTableCols(dtName);
}


// let arrayFunctions = [
//     getParameter("EnableTransmittal"),
//     getParameter("AllowMultiTrade"),
//     getParameter('Phase'),
//     getParameter('EnableWorkHours'),
//     getParameter('CheckAgainst_FileConvention'),
//     getParameter('CheckRevision'),
//     getGridMType(_web, _webUrl, 'LOD', false),
//     isMultiContractor() // _isMultiContracotr is already global variable there
// ];
// const params = await Promise.all(arrayFunctions);

 

