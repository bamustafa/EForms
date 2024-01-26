var _web, _webUrl, _siteUrl, _list, _layout, _module = '', _formType = '', _htLibraryUrl;
var _hot, _container, _data = [], _searchResultCount = 0, _Fields = [], _fieldsInternalName = [], _fieldsDisplayName = [], _fieldsType= [], _fieldSchema = [], _schemaInstance = [];

var _isNew = false, _isEdit = false, _isDisplay = false, _isMain = true, _isLead = false, _isPart = false;
var delayTime = 100, retryTime = 10, _timeout;

var _rootSite = '';
var _editableColumns = ['Considered', 'WhyNotConsidered'];
var _editableColumnsWidth = ['100px', '350px'];
var _colIndex = 0;

var defaultClassName = 'TransparentRow htMiddle';
var errorClassName = 'ErrorRow htMiddle';

let checkboxClicked = false;
var masterErrors = {};
var primaryField = 'Title';
var _currentTrade;

var onRender = async function (relativeLayoutPath, moduleName, formType){
  try{
    await extractValues(relativeLayoutPath, moduleName, formType);
    await getListFields();

    _data = await getData();
    _spComponentLoader.loadScript(_htLibraryUrl).then(renderHandsonTable);
    preloader("remove");
  }
  catch (e) {
    console.log(e);
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";
  }
}

//#region GENERAL
var getSoapRequest = async function(method, serviceUrl, isAsync, soapContent){
	var xhr = new XMLHttpRequest(); 
    xhr.open(method, serviceUrl, isAsync); 
    xhr.onreadystatechange = async function() 
    {
        if (xhr.readyState == 4) 
        {   
            try 
            {
                if (xhr.status == 200)
                {                
                    const obj = this.responseText;
                    var xmlDoc = $.parseXML(this.responseText),
                    xml = $(xmlDoc);
					
                    var value= xml.find("GLOBAL_PARAMResult");
                    if(value.length > 0){
                        text = value.text();
                        //_layout = value[0].children[0].textContent;
                        _rootSite = value[0].children[1].textContent;
                    }
                }            
            }
            catch(err) 
            {
                console.log(err + "\n" + text);             
            }
        }
    }
	xhr.setRequestHeader('Content-Type', 'text/xml');
    xhr.send(soapContent);

}

var extractValues = async function(relativeLayoutPath, moduleName, formType){
    localStorage.clear();

    if($('.text-muted').length > 0)
      $('.text-muted').remove();

    _web = pnp.sp.web;
    _module = moduleName;
    _formType = formType;
    _layout = relativeLayoutPath;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl; //https://db-sp.darbeirut.com/sites/QM-eAudit
    _siteUrl = new URL(_webUrl).origin;

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;

    if(_formType === 'New'){
        fd.clear();
        _isNew = true;
    }
    else if(_formType === 'Edit')
        _isEdit = true;
    else _isDisplay = true;

    var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GLOBAL_PARAM';
    var soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body>' +
                      '<GLOBAL_PARAM xmlns="http://tempuri.org/" />' +
                    '</soap:Body>' +
                  '</soap:Envelope>';
    await getSoapRequest('POST', serviceUrl, false, soapContent); // set _layout variable

    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

    await _spComponentLoader.loadScript(_layout + '/controls/jqwidgets/jqxcore.js').then(async () =>{
      var script = document.createElement("script"); // create a script DOM node
      script.src = _layout + "/plumsail/js/config/lodConfig.js"; // set its src to the provided URL
      document.head.appendChild(script);
      await renderControls();
    });

}

var renderControls = async function(){
  var isValid = false;
  var retry = 1;
  while (!isValid)
  {
    try{
      if(retry >= retryTime) break;
       await setFormHeaderTitle();
       preloader();
       await setCustomButtons();
       setButtonCustomToolTip('Submit', submitMesg);
       setButtonCustomToolTip('Cancel', cancelMesg);
       isValid = true;
    }
      catch{
        retry++;
        await delay(delayTime);
      }
  } 
}

var setCustomButtons = async function () {
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

    if(!_isDisplay)
     await setButtonActions("Accept", "Submit");
    await setButtonActions("ChromeClose", "Cancel");
}

const setButtonActions = async function(icon, text){
    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          click: async function() {
           if(text == "Close" || text == "Cancel")
               fd.close();
           if(text == "Submit"){
             debugger;
             await setItemsforInsertion();
             fd.close();
           }
       }
    });
}

function delay(time) {
    return new Promise(function(resolve) { 
        setTimeout(resolve, time)
    });
}

var ensureFunction = async function(funcName, ...params){
    var isValid = false;
    var retry = 1;
    while (!isValid)
    {
        try{
          if(retry >= retryTime) break;

          if(funcName === 'setListBox')
            setListBox(...params);
           isValid = true;
        }
        catch{
          retry++;
          await delay(delayTime);
        }
    }
}

function setErrorMessage(errMesg, textButton){

  function checkForErrors() {
    var errorElement = $('.errors[data-v-386d995a]');
    if (errorElement.length > 0) {
        clearInterval(intervalId);
        handleErrors(errorElement, errMesg, textButton);
    }

    $('button').hover(
      function () {
          // Mouseenter event handler
          $(this).css('color', 'white');
      },
      function () {
          // Mouseleave event handler
          $(this).css('color', 'black');
      }
  );
  }

  function handleErrors(element, errorMessage, buttonText) {
    if (errorMessage !== '') {
        element.css({
            'height': 'auto',
            'opacity': '1'
        });
        errorMessage = '<br/>' + errorMessage;
        var mesgElement = $('#customErrorId');
         if(mesgElement.length === 0)
          $('p.alert-heading').append(`<p id='customErrorId'>${errorMessage}</p>`);
         else mesgElement.html(errorMessage);
        $('span').filter(function () { return $(this).text() == buttonText; }).parent().css('color', '#737373').attr("disabled", "disabled");
    } else {
        element.css({
            'height': '0',
            'opacity': '0'
        });
        $('#customErrorId').remove();
        $('span').filter(function () { return $(this).text() == buttonText; }).parent().css('color', '#444').removeAttr('disabled');
    }
  }

  $('button.close').on('click', () => {
    $('div.alert').css({
      'height': '0',
      'opacity': '0'
    });
  });

  var intervalId = setInterval(checkForErrors, 100);
}

function addError(key, message) {
  if (!masterErrors.hasOwnProperty(key))
    masterErrors[key] = message;
}

function removeError(key) {
  if (masterErrors.hasOwnProperty(key))
   delete masterErrors[key];
}

function getErrors(){
  var mesg = '';

  for (var key in masterErrors) {
    mesg += masterErrors[key];
  }
  return mesg;
}
//#endregion

//#region GET FIELD AND BIND TO LISTBOX
var getListFields = async function(){  
    _colIndex = 0;
    await setFieldsSchema(true, false);
    await setFieldsSchema(false, true); 
}

var setFieldsSchema = async function(isEditableColumns, bindFieldsToControl){

  const excludedFieldNames = [
    '_ComplianceFlags',
    '_ComplianceTag',
    '_ComplianceTagUserId',
    '_ComplianceTagWrittenTime',
    '_CopySource',
    '_EditMenuTableEnd',
    '_EditMenuTableStart',
    '_EditMenuTableStart2',
    '_HasCopyDestinations',
    '_IsCurrentVersion',
    '_IsRecord',
    '_Level',
    '_ModerationComments',
    '_ModerationStatus',
    '_UIVersion',
    '_UIVersionString',
    '_VirusInfo',
    '_VirusStatus',
    '_VirusVendorID',
    'AccessPolicy',
    'AppAuthor',
    'AppEditor',
    'Attachments',
    'Author',
    'BaseName',
    'ComplianceAssetId',
    'ContentType',
    'ContentTypeId',
    'ContentVersion',
    'Created',
    'Created_x0020_Date',
    'DocIcon',
    'Edit',
    'Editor',
    'EncodedAbsUrl',
    'File_x0020_Type',
    'FileDirRef',
    'FileLeafRef',
    'FileRef',
    'FolderChildCount',
    'FSObjType',
    'GUID',
    'HTML_x0020_File_x0020_Type',
    'ID',
    'InstanceID',
    'ItemChildCount',
    'Last_x0020_Modified',
    'LinkFilename',
    'LinkFilename2',
    'LinkFilenameNoMenu',
    'LinkTitle',
    'LinkTitle2',
    'LinkTitleNoMenu',
    'MetaInfo',
    'Modified',
    'NoExecute',
    'Order',
    'OriginatorId',
    'owshiddenversion',
    'PermMask',
    'ProgId',
    'Restricted',
    'ScopeId',
    'SelectTitle',
    'ServerUrl',
    'SMLastModifiedDate',
    'SMTotalFileCount',
    'SMTotalFileStreamSize',
    'SMTotalSize',
    'SortBehavior',
    'SyncClientId',
    'UniqueId',
    'WorkflowInstanceID',
    'WorkflowVersion'
  ];

  await pnp.sp.web.lists.getByTitle(_list).fields.select("Title", "InternalName").orderBy("Title").get()
    .then(fields => {
      _Fields = fields.filter(field => {
          var internalName = field.InternalName;
          var displayName = field.Title;
          let fieldType = field['odata.type'];

           if(isEditableColumns){
            if(_editableColumns.includes(internalName))
              fetchFields(isEditableColumns, displayName, internalName, fieldType);
           }

           else {
              if(!excludedFieldNames.includes(internalName))
              {
                if(!_fieldsInternalName.includes(internalName) && !_editableColumns.includes(internalName)){
                  fetchFields(isEditableColumns, displayName, internalName, fieldType);
                }
                return !excludedFieldNames.includes(internalName);
              }
           }
         });
     })
    .then(() =>{
      if(bindFieldsToControl){ 
        _schemaInstance = JSON.parse(JSON.stringify(_fieldSchema));
        bindHTMLControls();
      }
    })
    .then(() =>{
      if(bindFieldsToControl) 
        ensureFunction('setListBox');
    })
    .catch(error => {
      console.error("Error", error);
    });
}

var fetchFields = async function(isEditableColumns, displayName, internalName, fieldType){
  var field = {title: displayName, data: internalName, index: _colIndex, className: 'htMiddle'};
  if(fieldType === 'SP.FieldMultiLineText' || fieldType === 'SP.FieldText' || fieldType === 'SP.FieldChoice'){
    fieldType = 'text';
    field.type = fieldType;
    field.renderer = 'html';
  }
  else if(fieldType === 'SP.Field'){
    fieldType = 'checkbox';
    field.type = fieldType;
  }

  if(isEditableColumns){
    var columnIndex = _editableColumns.indexOf(internalName);
    field.width = _editableColumnsWidth[columnIndex];
  }
  else {
    if(internalName === 'Electrical')
      field.width = '400px';
    else field.width = '150px';
  }

  if(_isDisplay)
  field.readOnly = true;

  // if(internalName.length > 10)
  //   field.width = '190px';
  // else field.width = '100px';

  _fieldSchema.push(field);
  _fieldsInternalName.push(internalName);
  _fieldsDisplayName.push(displayName);
  _fieldsType.push(field.type);
  _colIndex++;
}

var bindHTMLControls = async function(){
    if($('label.title-label').length === 0){
        var div1 = $('<div>').css({'float': 'left', 'margin-right': '20px'});
        var label1 = $('<label>').attr('for', 'listFields').addClass('title-label').text('List Fields');
        var listBoxA = $('<div>').attr('id', 'listBoxA');
        div1.append(label1, listBoxA);

        $('#contentId').append(div1);
    }
}

function setListBox(){

    //$("#jqxSplitter").jqxSplitter({ theme: 'summer', panels: [{ size: '100px' }] });
    //debugger;
    //$('#jqxSplitter').jqxSplitter({ width: '100%', panels: [{ size: '20%' }] });

    const listBox = $("#listBoxA");
    listBox.jqxListBox({ checkboxes: true, filterable: true, source: _fieldsDisplayName, width: 300, height: 680});

    listBox.jqxListBox('getItems').forEach(function (item, index) {
      listBox.jqxListBox('checkIndex', index);
    });

    var index = 0;
    _editableColumns.find(item => {
        listBox.jqxListBox('disableAt', index);
        index++;
    });

    listBox.on('checkChange', function (event) {
        var args = event.args;
        var itemTitle = args.label;
        //var itemIndex = args.item.index;
        if (args.checked) 
          updateSchema(itemTitle, 'add');
        else updateSchema(itemTitle, 'remove');
    });
}
//#endregion

//#region GET DATA AND BIND TO HANDSONTABLE WITH VALIDATION
var getData = async function(){ 
  var _itemArray = [];
  var _cols = _fieldsInternalName.join(',');

  await getUserTradeFromPMIS();//'ME';
  var _query = `substringof('${_currentTrade}', LeadDiscipline)`;

  var rootWeb;
  if(_isNew)
   rootWeb = new Web(_rootSite);
  else rootWeb = pnp.sp.web;

  await rootWeb.lists.getByTitle(_list).items.filter(_query)
  .select('Id,' + _cols)
  .getAll().then(async function(items) {
    _itemCount = items.length;
    if (_itemCount > 0) {
      for(var i = 0; i < _itemCount; i++){
         var item = items[i];
         var rowData  = {};

         for(var j = 0; j < _fieldsInternalName.length; j++){
            var _colname = _fieldsInternalName[j];
            var _value = item[_colname];
            rowData[_colname] = _value;
         }
         _itemArray.push(rowData);
         
      }
      console.log(_itemArray);
    }
});
return _itemArray;
}

const renderHandsonTable = (Handsontable) => {
     _container = document.getElementById('dt');
     _hot = new Handsontable(_container, {
          data: _data,
          columns: _fieldSchema,
          //width:'100%',
          height: '620',
          search: {
            searchResultClass: 'highlight-cell'
          },
          // filters: true,
          // filter_action_bar: true,
          //dropdownMenu: ['filter_by_condition', 'filter_by_value', 'filter_action_bar'],
          fixedColumnsLeft: _editableColumns.length,
          rowHeaders: true,
          colHeaders: true,
          manualColumnResize: true,
          manualRowMove: true,
          stretchH: 'all',
          licenseKey: htLicenseKey
     });
  setSearchField();

  if(!_isDisplay){
    setRowsReadOnly(); //WRITE ERRORS HERE
    setErrorMessage(getErrors(), 'Submit');
    hooks();
  }
}

var updateSchema = async function(fieldTitle, trans){
  if(trans === 'remove'){
    var columnIndexTofind = _schemaInstance.findIndex(column => column.title === fieldTitle);
    //updatedFieldSchema = [..._fieldSchema];
    _schemaInstance.splice(columnIndexTofind, 1);
  }
  else if(trans === 'add'){
    _fieldSchema.map(item =>{
      if(fieldTitle === item.title){
        _schemaInstance.push(item);
      }
    });

    _schemaInstance.sort(function (a, b) {
      return a.index - b.index;
    });
  }

  _hot.updateSettings({
    columns: _schemaInstance
  });
  setRowsReadOnly();
}

var setRowsReadOnly = async(value, rowIndex, colIndex) => {
  var defaultMetaObject = {
    readOnly: true,
    className: defaultClassName
  };

  var errorMetaObject = {
    readOnly: false,
    className: errorClassName
  };

  var considered = _editableColumns[0];
  
  if(rowIndex !== undefined && colIndex !== undefined){
    validateEditableColumns(value, rowIndex, colIndex, defaultMetaObject, errorMetaObject);
    setErrorMessage(getErrors(), 'Submit');  
  }
  
  else{
    const rowCount = _hot.countRows();
    for (var row = 0; row < rowCount; row++) {
      var columnIndex = 0;
        _hot.getSettings().columns.find(item => {
          if(item.title === considered){
            let isConsidered = _hot.getDataAtCell(row, columnIndex);
            validateEditableColumns(isConsidered, row, columnIndex, defaultMetaObject, errorMetaObject);
            columnIndex++;
          }
          else _hot.setCellMetaObject(row, columnIndex, defaultMetaObject);
          columnIndex++;
        });
    }   
  }
  _hot.render();
}

var validateEditableColumns = async function(isConsidered, rowIndex, columnIndex, defaultMetaObject, errorMetaObject){ 
  var verifyIndex = columnIndex + 1;
  var _mesg = ''
  var verifyValue = _hot.getDataAtCell(rowIndex, verifyIndex);
   if(!isConsidered){
      if (verifyValue === null || verifyValue === ''){
        _hot.setCellMetaObject(rowIndex, verifyIndex, errorMetaObject);
        //_hot.setCellMeta(rowIndex, verifyIndex, 'className', errorClassName);
        _mesg = `Why not Considered is required field at row ${rowIndex+1} <br/>`;
        addError(rowIndex, _mesg);
      }
      else{
        _hot.setCellMeta(rowIndex, verifyIndex, 'readOnly', false);
        removeError(rowIndex);
      }
   }
   else { 
    _hot.setCellMetaObject(rowIndex, verifyIndex, defaultMetaObject);
    //_hot.setCellMeta(rowIndex, verifyIndex, 'className', defaultClassName);
    removeError(rowIndex);
   }
}

 var hooks = async function(){
  _hot.addHook('afterChange', (changes, source) => {
    if (source === 'edit') {
        changes.forEach(([rowIndex, prop, oldValue, value]) => {
            const columnIndex = _hot.propToCol(prop);
            if (_hot.getCellMeta(rowIndex, columnIndex).type === 'checkbox') {
              setRowsReadOnly(value, rowIndex, columnIndex);
            }
            else if(_hot.getCellMeta(rowIndex, columnIndex).data === 'WhyNotConsidered'){
              var getText = _hot.getDataAtCell(rowIndex, columnIndex);
              if(getText !== undefined && getText !== null && getText !== ''){
                _hot.setCellMeta(rowIndex, columnIndex, 'className', defaultClassName);
                //_hot.setCellMetaObject(rowIndex, columnIndex, defaultMetaObject);
                removeError(rowIndex);
              }
              else {
                _hot.setCellMeta(rowIndex, columnIndex, 'className', errorClassName);
                //_hot.setCellMetaObject(rowIndex, columnIndex, errorMetaObject);
                addError(rowIndex, `Why not Considered is required field at row ${rowIndex+1} <br/>`);
              }
              _hot.render();
              var errMesg = getErrors();
              setErrorMessage(errMesg, 'Submit');
            }
        });
    }
  });
}

var setItemsforInsertion = async function(){
  var mentaInfo = _hot.getSettings().columns;

	var rows;
  if(_isNew)
    rows = _data;
  else rows = _hot.getData();

  var rowCount = rows.length;

  var itemsToInsert = [], objColumns = [], objTypes = [];

  for(var rowIndex = 0; rowIndex < rowCount; rowIndex++){
    var _objValue = { };
    var _columns = '', _colType = '';
    var doInsert = false;

    for(var colIndex = 0; colIndex < mentaInfo.length; colIndex++){
      var _value = rows[rowIndex][colIndex];
      var internalName = mentaInfo[colIndex].data;

      if(_value !== null && _value !== ''){
        _columns += internalName + ',';
        _colType += mentaInfo[colIndex].type + ',';
        _objValue[internalName] = _value;
        doInsert = true;
       }
    }

    if(doInsert){
      _columns = _columns.endsWith(',') ? _columns.slice(0, -1) : _columns;
      _colType = _colType.endsWith(',') ? _colType.slice(0, -1) : _colType;

      objColumns.push(_columns);
      objTypes.push(_colType);
      itemsToInsert.push(_objValue);
    }
  }

  if(itemsToInsert.length > 0)
   await insertItemsInBatches(itemsToInsert, objColumns, objTypes);
}

var insertItemsInBatches = async function(itemsToInsert, objColumns, objTypes) {
  const list = pnp.sp.web.lists.getByTitle(_list);
  const batch = pnp.sp.createBatch();
  
  var row = 0;
  var _columns, _colType;

  for (const item of itemsToInsert){
    var _query = `${primaryField} eq '${item[primaryField]}'`;

    _columns = objColumns[row];
    _colType = objTypes[row];

    const existingItems = await list.items
    .select("Id," + _columns)
    .filter(_query)
    .top(1)
    .get();

    if (existingItems.length > 0) {
      var _item = existingItems[0]; // _item is the oldItem values while objValue is the new one
      var _cols = _columns.split(',');
      var _type = _colType.split(',');
      var doUpdate = false;

      for(var i = 0; i < _cols.length; i++){
        var columnName = _cols[i];
        var oldValue = '';
        var currentValue = '';

        oldValue = _item[columnName];
        currentValue = item[columnName];

        if( oldValue != currentValue ){
          doUpdate = true;
          break;
        }
      }
      if(doUpdate)
       list.items.getById(existingItems[0].Id).inBatch(batch).update(item);
    } else {
       list.items.inBatch(batch).add(item);
    }
    row++;
  }
  await batch.execute();
}

var getUserTradeFromPMIS = async function() 
{   
    var currentUserMetaInfo = await pnp.sp.web.currentUser.get();
    var currentDisplayName = currentUserMetaInfo.Title;

    var projectCode = _webUrl.split('/')[4];
    if(_webUrl.includes('db-sp.')){
      projectCode = 'AN22063-0100D';
      currentDisplayName = 'Hisham Kassouf';
    }
  
    var SERVICE_URL = _siteUrl + '/AjaxService/XMLService.asmx?op=GetPMISTrade';

    var xhr = new XMLHttpRequest();
    var soapRequest = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
                    + '<soap:Body>'
                    + '<GetPMISTrade xmlns="http://dargroup.com/">'
                       + '<ProjectCode>' + projectCode + '</ProjectCode>'
                       + '<UserDisplayName>' + currentDisplayName + '</UserDisplayName>'
                    + '</GetPMISTrade>'
                    + '</soap:Body>'
                    + '</soap:Envelope>';
                
    xhr.open("POST", SERVICE_URL, false);
    xhr.onreadystatechange = async function() 
    {
        if (xhr.readyState == 4) 
        {   
            try 
            {
                if (xhr.status == 200)
                {                                                 
                    var xmlDoc = $.parseXML( xhr.responseText );
                    $xml = $( xmlDoc );
                    $value= $xml.find("GetPMISTradeResult");

                    _currentTrade = $value.text();
                    return _currentTrade;
                }
                else
                    console.log('Request failed with status code:', xhr.status);
            }
            catch(err) 
            {
                console.log(err + "\n" + text);             
            }
        }
    }           

    xhr.setRequestHeader('Content-Type', 'text/xml');
    xhr.send(soapRequest);
}
//#endregion

//#region HANDSONTABLE SEARCH
function onlyExactMatch_notused(queryStr, value) {
  if(queryStr !== '' && queryStr !== null && queryStr !== undefined)
    return String(value).toLowerCase().includes(queryStr.toLowerCase());
  else return false;
};

var setSearchField = async function(){
  var id = '#search_field';
     if ($(id).length === 0) {
       var input = $('<input>', {
         id: 'search_field',
         type: 'text',
         width: '300px',
         placeholder: 'Search any...'
       });

       $('#dt').before(input);
       input.after('<br><br>');
     }
     const searchField = document.querySelector('#search_field');

     searchField.addEventListener('keyup', function(event) {
      const colCount = _hot.countCols();
      var searchResult = event.target.value.toLowerCase();

      const search = _hot.getPlugin('search');
      const queryResult = search.query(searchResult);

      if(searchResult.trim().length === 1)
         return;

      var _cols = _hot.getSettings().columns;
      _hot.updateSettings({
          data: _data.filter((row, rowIndex) => {
            //debugger;
          if(rowIndex === 0)
          _searchResultCount = 0;
          
           let found = false;
            for (let col = 0; col < colCount; col++) {
              var colTitle = _cols[col].data;
              if(colTitle === 'attachment' || colTitle === 'Considered')
               continue;
              cellValue = row[colTitle];
              if(cellValue !== null){
                // if(searchResult === '' && cellValue.includes('</span>') ){
                //   cellValue = cellValue.replace(/<span class=['"]my-custom-search-result-class['"]>/g, '').replace(/<\/span>/g, '');
                //   _hot.setDataAtCell(rowIndex, col, cellValue);
                //   doRender = true;
                // }
                //else{
                  if(searchResult !== ''){
                    //cellValue = cellValue.replace(/<span class=['"]my-custom-search-result-class['"]>/g, '').replace(/<\/span>/g, '');;
                    if (String(cellValue).toLowerCase().includes(searchResult)) {
                        found =  true;
                        //rowRenderer(_hot, col, rowIndex, cellValue, searchResult);
                        _searchResultCount++;
                        break;
                    }
                  }
                  else found = true
                //}
              }
            }
            return found;
          })
      });

      var id = '#resultlbl';
      if(_searchResultCount > 0){
          var mesg = _searchResultCount + ' rows found';
          if(_searchResultCount === 1)
           mesg = _searchResultCount + ' row found';
          
          if(searchResult !== '' && searchResult !== null && searchResult !== undefined){
              if ($(id).length === 0) {
                  var label = $('<label>', {
                      id: 'resultlbl',
                      text: mesg
                  });
                  label.css('color', '#3CDBC0');
                  label.css('width', '350px'); 
                  
                  $('#search_field').after(label);
                  label.before('<br>');
              }
              else $('#resultlbl').text(mesg);
          }
          else{
              if ($(id).length > 0){
               const brElement = $(id).prev();
               brElement.remove();
                $(id).remove();
              }
          }
      }
      else {
          if ($(id).length > 0){
              const brElement = $(id).prev();
              brElement.remove();
              $(id).remove();
          }
      }
      _hot.render();
  });
  
}

// Custom cell renderer function with highlighting
function rowRenderer_notused(hotInstance, col, rowIndex, cellValue, searchResult){

  let highlightedCellValue = highlightSearchResult(cellValue, searchResult);
  // Set the highlighted value back to the cell
  hotInstance.setDataAtCell(rowIndex, col, highlightedCellValue);
  
  // for (let col = 0; col < colCount; col++) {
  //   var colTitle = _cols[col].data;
  //   if (colTitle === 'attachment' || colTitle === 'Considered') continue;

  //   let cellValue = hotInstance.getDataAtCell(rowIndex, col);
  // }
}

// Custom function to highlight search result in a cell value
function highlightSearchResult_notused(cellValue, searchResult) {
  // Implement your custom logic to highlight the search result in the cell value
  // For example, you can wrap the search result in <span> tags with a specific class for styling
  return cellValue.replace(new RegExp(searchResult, 'gi'), match => `<span class='my-custom-search-result-class'>${match}</span>`);
}
//#endregion


