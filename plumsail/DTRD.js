var _web, _webUrl, _siteUrl, _list, _layout, _module = '', _formType = '', _htLibraryUrl;
var _hot, _container, _data = [], _searchResultCount = 0, _Fields = [], _fieldsInternalName = [], _fieldsDisplayName = [], _fieldsType= [], _fieldSchema = [], _schemaInstance = [];

var _isNew = false, _isEdit = false, _isDisplay = false, _isMain = true, _isLead = false, _isPart = false, isChamp = false;
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
let _formStatus = '';
let screenHeight = screen.height - 500;

let pmisTrades = [], rootsiteTrades = [], submittedTrades = [], PendingChampTrades = [], missingTrades = [], statuses = [];
let isSentToChampion = false, isApproved = true, dimSubmit = false;

let DTRDC = 'DRDC', DTRDL = 'DRDL', _commonMembers;

var onRender = async function (relativeLayoutPath, moduleName, formType){
  try{ 
    const startTime = performance.now();
    $(fd.field('RejTrades').$parent.$el).hide();

    _layout = relativeLayoutPath;

    _spComponentLoader.loadScript( _layout + '/controls/preloader/jquery.dim-background.min.js').then(()=>{
      _spComponentLoader.loadScript( _layout + '/plumsail/js/preloader.js').then(()=>{
        preloader();
      })
    })
    .then(async ()=>{

      await extractValues(relativeLayoutPath, moduleName, formType);
    
      let members = await getSharePointGroupMembers(DTRDL);
      if(members.length === 0){
        alert('DRD is not applicable');
        fd.close();
      }
      await getListFields();
  
      const startTime1 = performance.now();
      _data = await getData();
      const endTime1 = performance.now();
      const elapsedTime1 = endTime1 - startTime1;
      console.log(`getData: ${elapsedTime1} milliseconds`);
  
      _spComponentLoader.loadScript(_htLibraryUrl).then(renderHandsonTable)
      .then(async ()=>{
        await setStatuses();
      })
      .then(async ()=>{
        await bindTrades();
        setRowsReadOnly(); //WRITE ERRORS HERE
      })
      .then(async ()=>{
        await setCustomButtons();
        setButtonCustomToolTip('Submit', submitMesg);
        setButtonCustomToolTip('Cancel', cancelMesg);
      });
  
      preloader("remove");
      const endTime = performance.now();
      const elapsedTime = endTime - startTime;
      console.log(`Execution onRender: ${elapsedTime} milliseconds`);
    })
  }
  catch (e) {
    console.log(e);
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";
  }
}

//#region GENERAL
var loadScripts = async function(){
  const libraryUrls = [
    _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
    _layout + '/plumsail/js/customMessages.js',
    _layout + '/plumsail/js/commonUtils.js',

    _layout + '/controls/jqwidgets/jqxcore.js',
    _layout + '/controls/jqwidgets/jqxbuttons.js',
    _layout + '/controls/jqwidgets/jqxscrollbar.js',
    _layout + '/controls/jqwidgets/jqxlistbox.js',
    _layout + '/controls/jqwidgets/jqxcheckbox.js',
    _layout + '/controls/jqwidgets/jqxsplitter.js'
  ];

  const cacheBusting = `?v=${Date.now()}`;
    libraryUrls.map(url => { 
        $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
    });
      
  const stylesheetUrls = [
       _layout + '/controls/tooltipster/tooltipster.css',
       _layout + '/controls/handsonTable/libs/handsontable.full.min.css',
       _layout + '/plumsail/css/CssStyle.css',

       _layout + '/plumsail/css/CssStyleRACI.css',
       _layout + '/controls/jqwidgets/styles/jqx.base.css',
       _layout + '/controls/jqwidgets/styles/jqx.summer.css'
      ];


  stylesheetUrls.map((item) => {
    var stylesheet = item;
    $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
  });
}

var getSoapResponse = async function(method, serviceUrl, isAsync, soapContent, getResultTag){
	var xhr = new XMLHttpRequest(); 
    xhr.open(method, serviceUrl, isAsync); 
    xhr.onreadystatechange = async function() 
    {
        if (xhr.readyState == 4) 
        {   
            try 
            {
                if (xhr.status == 200 && getResultTag !== '')
                {                
                    const obj = this.responseText;
                    var xmlDoc = $.parseXML(this.responseText),
                    xml = $(xmlDoc);
					
                    var value= xml.find(getResultTag);
                    if(value.length > 0){
                        text = value.text();
                        //_layout = value[0].children[0].textContent;
                        _rootSite = value[0].children[1].textContent;
                    }
                }
                else console.log(`status ${xhr.status} - ${xhr.statusText} `);          
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
  fd.toolbar.buttons[0].style = "display: none;";
  fd.toolbar.buttons[1].style = "display: none;";

    if($('.text-muted').length > 0)
      $('.text-muted').remove();

    _web = pnp.sp.web;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl; //https://db-sp.darbeirut.com/sites/QM-eAudit
    _siteUrl = new URL(_webUrl).origin;

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;

    if(_formType === 'New')
        _isNew = true;
    else if(_formType === 'Edit')
        _isEdit = true;
    else _isDisplay = true;

    var serviceUrl = _siteUrl + '/AjaxService/DarPSUtils.asmx?op=GLOBAL_PARAM';
    var soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body>' +
                      '<GLOBAL_PARAM xmlns="http://tempuri.org/" />' +
                    '</soap:Body>' +
                  '</soap:Envelope>';
    await getSoapResponse('POST', serviceUrl, false, soapContent, 'GLOBAL_PARAMResult'); // set _layout variable

    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

    await loadScripts();
    await setFormHeaderTitle();
}

var setCustomButtons = async function () {
  

    if(missingTrades.length === 0 && PendingChampTrades.length === 0){
      await setButtonActions("ChromeClose", "Cancel");
    }
    else{
      if(isChamp){
        $('button.btn-outline-primary').remove();
        if(isSentToChampion && isChamp){
          await setButtonActions("Accept", "Approve");
          await setButtonActions("ChromeClose", "Reject");
        }
        else {
          if(statuses[0] !== 'Completed')
            await setButtonActions("Accept", "Submit");
        }
        
        await setButtonActions("ChromeClose", "Cancel");
        return;
      }

      if(!_isDisplay)
      await setButtonActions("Accept", "Submit");
      await setButtonActions("ChromeClose", "Cancel");

      if(dimSubmit)
        $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#737373').attr("disabled", "disabled");
    }
}

const setButtonActions = async function(icon, text){
   let query = '';
    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          click: async function(){
           
           if(text == "Close" || text == "Cancel"){
               preloader();
               fd.close();
           }
           else if(text == "Submit"){
             preloader();
             setItemsforInsertion('Sent to Champion')
             .then(async () => {
                query = `<Where><Eq><FieldRef Name='Status' /><Value Type='Text'>Sent to Champion</Value></Eq></Where>`;
                let champEmails = '';
                if(_commonMembers !== undefined){
                  _commonMembers.map(user=>{
                    champEmails += user.Email + ',';
                  })
                }
                await sendEmail('TradeDTRD', query, 'TradeDTRD', champEmails);
                fd.close();
              })
              .catch(error => {
                console.error("Error during insertion:", error);
             });
           }

           else if(text == "Approve"){
            preloader();
            setItemsforInsertion('Completed')
              .then(async () => {
                missingTrades = missingTrades.filter(item => item !== _currentTrade);
                query = `<Where><Eq><FieldRef Name='Status' /><Value Type='Text'>Completed</Value></Eq></Where>`;
                if(missingTrades.length === 0){
                  await sendEmail('ApprovedDTRD', query, 'ApprovedDTRD');
                }
                fd.close();
              })
              .catch(error => {
                console.error("Error during Approve:", error);
              });
           }

           else if(text == "Reject"){
            //let rejTrades = fd.field('RejTrades').value;
            // if (rejTrades === null || rejTrades.length === 0){
            //    alert('select trades for rejection and proceed');
            //    return;
            // }
          
            preloader();
            setItemsforInsertion('Rejected')
              .then(async () => {
                //query = setQuery(rejTrades);
                 query = `<Where><Eq><FieldRef Name='LeadDiscipline' /><Value Type='Text'>${_currentTrade}</Value></Eq></Where>`;
                await sendEmail('RejectDTRD', query, 'RejectDTRD');
                fd.close();
              })
              .catch(error => {
                console.error("Error during Rejected:", error);
            });
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
    // var isValid = false;
    // var retry = 1;
    // while (!isValid)
    // {
    //     try{
    //       if(retry >= retryTime) break;

          if(funcName === 'setListBox')
            setListBox(...params);
           //isValid = true;
    //     }
    //     catch{
    //       retry++;
    //       await delay(delayTime);
    //     }
    // }
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

function setQuery_notUsed(rejTrades){
  let query = `<Where><Eq><FieldRef Name='LeadDiscipline' /><Value Type='Text'>${rejTrades[0]}</Value></Eq></Where>`;

  if(rejTrades.length === 2){
    query = `<Where><And>
               <Eq><FieldRef Name='LeadDiscipline' /><Value Type='Text'>${rejTrades[0]}</Value></Eq>
               <Eq><FieldRef Name='LeadDiscipline' /><Value Type='Text'>${rejTrades[1]}</Value></Eq>
             </And></Where>`;
  }
  else if(rejTrades.length > 2){
    query = '<Where>';

    for (let i = 0; i < rejTrades.length; i++){
      query += '<And>';
    }

    for (let i = 0; i < rejTrades.length; i++){
      query += "<Eq><FieldRef Name='LeadDiscipline' /><Value Type='Text'>" + rejTrades[i] + "</Value></Eq>";
      if(i === 1 || i > 2)
          query += '</And>';
    }
    query += '</Where>';
  }
  return query
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

  if(displayName !== 'Status')
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
  let ignoreChecking = false;
  
  // let groupName = await getAllowedTeam();
  // console.log(`groupName is ${groupName}`);
  
   await getUserTrade();//'ME';
   console.log(`groupName is ${_currentTrade}`);

  if(_currentTrade === undefined){
    let isExist = await IsUserInGroup(DTRDL);
    if(!isExist){
      alert('you are not allowed to view items');
      fd.close();
    }
  }
  else{
    let tradeMembers = await getSharePointGroupMembers(_currentTrade);
    let dtrdcMembers = await getSharePointGroupMembers(DTRDC);
    _commonMembers = compareArrays(dtrdcMembers, tradeMembers);

    if(_commonMembers.length === 0){
      alert('please select champion from pmis');
      fd.close();
    }
  }


 
   let currentStatus;
   let statusitems =  await _web.lists.getByTitle(_list).items.select('Status,LeadDiscipline').getAll();
   if(statusitems.length > 0){
    for(const item of statusitems){
      currentStatus = item.Status;
      if(currentStatus !== 'Completed'){
        isApproved = false;
        break;
      }
    }
  }
  else  isApproved = false;

  var _query = `LeadDiscipline eq '${_currentTrade}'`; //`substringof('${_currentTrade}', LeadDiscipline)`;
  if(isApproved){
    _query = `LeadDiscipline ne 'blabla'`
  }
  else if(isChamp && currentStatus !== 'Rejected')
   _query = `LeadDiscipline eq '${_currentTrade}' and Status eq 'Sent to Champion'`;


   let isItemsFound  = true;
  var rootWeb = _web;
  let items =  await rootWeb.lists.getByTitle(_list).items.filter(_query).getAll();
  if(items.length === 0){
      rootWeb = new Web(_rootSite);
      _fieldsInternalName = _fieldsInternalName.filter(col => col !== "Status");

      if(isChamp){
        ignoreChecking = true;
        isItemsFound = false;
      }
  }
  
  if(!ignoreChecking){
    var _cols = _fieldsInternalName.join(',');
    await rootWeb.lists.getByTitle(_list).items.filter(_query)
    .select('Id,' + _cols)
    .getAll().then(async function(items){
      _itemCount = items.length;
      if (_itemCount > 0) {
        for(var i = 0; i < _itemCount; i++){
          var item = items[i];
          var rowData  = {};

          for(var j = 0; j < _fieldsInternalName.length; j++){
              var _colname = _fieldsInternalName[j];
              var _value = item[_colname];
              rowData[_colname] = _value;

              // if(_colname === 'LeadDiscipline'){
              //   let tempStatus = item['Status'];
              //   if(tempStatus !== 'Rejected' && !submittedTrades.includes(_value))
              //     submittedTrades.push(_value);
              // }
              //else 
              if(_colname === 'Status'){
                if(!statuses.includes(_value))
                  statuses.push(_value);
              }
          }
          _itemArray.push(rowData);
        }
      }
    });
  }

 if(!isApproved && !isItemsFound){
  if(isChamp){
    alert('you login as Champ, not items yet found to send to Lead');
    fd.close();
  }
 }
 return _itemArray;
}

function compareArrays(arr1, arr2) {
  const differences = [];

  arr1.forEach(obj1 => {
    const matchingObj = arr2.find(obj2 => obj1.Title === obj2.Title);
    
    if (matchingObj !== undefined) {
      differences.push(obj1);
    }
  });

  return differences;
}

var getAllowedTeam_notUsed = async function(){
    let isAllowedUser = true;
    if(_currentTrade === undefined){
        isAllowedUser = await IsUserInGroup(DTRDC);
        if(isAllowedUser)
          isChamp = true;
    }
    else groupName = _currentTrade;

    if(!isAllowedUser){
      alert(dtrdPermission);
      fd.close();
    }
    return groupName
}
  
const renderHandsonTable = (Handsontable) => {

     _container = document.getElementById('dt');
     _hot = new Handsontable(_container, {
          data: _data,
          columns: _fieldSchema,
          width:'99%',
          height: screenHeight,
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

     let index = _hot.propToCol('Status'); //set data
        _hot.updateSettings({
          hiddenColumns: {
            columns: [index],
            indicators: false // Show the hidden columns indicators
          }
        });

     setSearchField();

  if(!_isDisplay){
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
          if(isChamp || dimSubmit){
            _hot.setCellMeta(row, columnIndex, 'readOnly', true);
          }
          else{
            if(item.data === 'Considered'){
              let isConsidered = _hot.getDataAtCell(row, columnIndex);
              validateEditableColumns(isConsidered, row, columnIndex, defaultMetaObject, errorMetaObject);
              columnIndex++;
            }
            else _hot.setCellMetaObject(row, columnIndex, defaultMetaObject);
          }
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

   if(isConsidered){
      if (verifyValue === null || verifyValue === ''){
        _hot.setCellMetaObject(rowIndex, verifyIndex, errorMetaObject);
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

var setItemsforInsertion = async function(nextStatus){
  var mentaInfo, rows;
    rows = _data;

  mentaInfo = _fieldSchema;
  var itemsToInsert = [], objColumns = [], objTypes = [];

  rows.map((columns, rowIndex) => {
    var _objValue = { };
    var _columns = '', _colType = '';
    var doInsert = false;

    var colIndex = 0;
    for (const column in columns) {
          var value = columns[column];
          var internalName = mentaInfo[colIndex].data;

          if(value !== null && value !== ''){
            _columns += internalName + ',';
            _colType += mentaInfo[colIndex].type + ',';

            if(nextStatus !== undefined && internalName === 'Status')
              _objValue[internalName] = nextStatus;
            else _objValue[internalName] = value;
            doInsert = true;
          }
      colIndex++;
    }
    if(doInsert){
      _columns = _columns.endsWith(',') ? _columns.slice(0, -1) : _columns;
      _colType = _colType.endsWith(',') ? _colType.slice(0, -1) : _colType;

      objColumns.push(_columns);
      objTypes.push(_colType);
      itemsToInsert.push(_objValue);
    }
  });

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

    const existingItems =  await list.items 
    .select("Id," + _columns)
    .filter(_query)
    .top(1)
    .get();

    if (existingItems.length > 0){
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
        list.items.getById(_item.Id).inBatch(batch).update(item);
    }
    else list.items.inBatch(batch).add(item);
    
     row++;
  }
  await batch.execute();
}

var getUserTrade = async function(){

  let [items, user] = await Promise.all([
     _web.lists.getByTitle('Trades').items.select('Title').orderBy('Title').getAll(),
     _web.currentUser.get()
  ]);

  let userGroups = await _web.siteUsers.getById(user.Id).groups.get();
  if(isChamp)
    userGroups = userGroups.filter(group=> group.Title !== DTRDC)
  
   for (let trade of items){
      let matchingGroup = userGroups.find(group => group.Title === trade.Title);
        if (matchingGroup) {
            _currentTrade = trade.Title;
            isChamp = await IsUserInGroup(DTRDC);
            break;
        }
    }
}

async function getSharePointGroupMembers(groupName) {
    const group = await _web.siteGroups.getByName(groupName);
    const members = await group.users.get();
    return members;
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
      setTimeout(()=>{ 
            _hot.render()}, 
            300);
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

//#region COMPARE TRADES FROM PMIS
const bindTrades = async function(){
  //if(isSentToChampion){
    
    let webMetaInfo = await _web.select("Title,Url").get();
    let projNo = webMetaInfo.Url.includes('db-sp') || webMetaInfo.Url.includes('Dtemp22') ? 'AN21146-0100R' : webMetaInfo.Title;
    console.log(`projNo = ${projNo}`);
    await getPMISProjectDepartments(projNo); //pmisTrades

    await getDistinctTrades_From_RootSite(true); //rootsiteTrades
    await getDistinctTrades_From_RootSite(false, 'Completed'); //submittedTrades
    await getDistinctTrades_From_RootSite(false, 'Sent to Champion'); //PendingChampTrades

    let commonTrades = pmisTrades.filter(trade => rootsiteTrades.includes(trade));
    missingTrades = commonTrades.filter(trade => !submittedTrades.includes(trade) && !PendingChampTrades.includes(trade));

     
     addkeyValTextCtlr('Submitted Trades:', submittedTrades, '120px', 'green');
     addkeyValTextCtlr('Pending Champion Approval:', PendingChampTrades, '90px', 'orange');
     addkeyValTextCtlr('Missing Trades:', missingTrades, '40px', 'red');
     //if(isChamp && submittedTrades.length > 0){
      //$(fd.field('RejTrades').$parent.$el).show();
      // fd.field('RejTrades').options = submittedTrades;
      // let rejTrades = $('.row').eq(3);
      // $('#dt').before(rejTrades);
     //}
  //}
}

var getPMISProjectDepartments = async function(ProjectNo) 
{ 
  var serviceUrl = _webUrl+ "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetProjectDepartments&ProjectNo="+encodeURIComponent(ProjectNo);
  let xmlDoc = await getSoapResponse1('POST', serviceUrl, true, '', 'GetProjectDepartmentsResult');

  const table1Nodes = xmlDoc.getElementsByTagName("Table1");
  for (let i = 0; i < table1Nodes.length; i++) {
    const dptAcronym = table1Nodes[i].getElementsByTagName("ABR")[0].textContent;
    if(!pmisTrades.includes(dptAcronym))
      pmisTrades.push(dptAcronym);
    //const departmentName = table1Nodes[i].getElementsByTagName("DepartmentName")[0].textContent;
  }
}

const addkeyValTextCtlr = async function(key, value, padding, color){
  let html = `<label style="padding-left: ${padding}; color: #4e778f; font-weight:bold;">
                <span class="overflow-hidden fd-title-wrap">${key}</span>
              </label>
                  
              <label class="fd-field-control col-sm" style='color: ${color}'><div class="fd-sp-field-text col-form-label">
                 ${value}
              </label>`;

  $('#search_field').after(html);
}

const getDistinctTrades_From_RootSite = async function(isRoot, status){

  let rootWeb = isRoot ? new Web(_rootSite) : _web;
  let filter = "LeadDiscipline ne '123'";

  if(!isRoot && status !== undefined){
    filter = `Status eq '${status}'`;
  }

  await rootWeb.lists.getByTitle(_list).items.filter(filter).select('LeadDiscipline').getAll()
  .then(async function(items) {
     
    for(let item of items){
       let discipline = item.LeadDiscipline;

       if(isRoot){
        if(!rootsiteTrades.includes(discipline))
          rootsiteTrades.push(discipline);
       }
       else{
         if(status === 'Completed'){
          if(!submittedTrades.includes(discipline))
             submittedTrades.push(discipline);
         }
         else{
            if(!PendingChampTrades.includes(discipline))
              PendingChampTrades.push(discipline);
         }
       }
    }
  });
}
//#endregion

var sendEmail = async function(emailName, query, notificationName, champEmails){
  let CurrentUser = await GetCurrentUser();
  let approvalCC = _currentTrade;

  if(emailName === 'TradeDTRD'){
    if(champEmails !== undefined && champEmails !== ''){
      if (champEmails.endsWith(','))
        champEmails = champEmails.slice(0, -1);
        approvalCC += '|' + champEmails;
    }
  }

  let serviceUrl = _siteUrl + '/AjaxService/DarPSUtils.asmx?op=SEND_EMAIL_TEMPLATE';
  let soapContent = '';
  soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                  <soap:Body>
                    <SEND_EMAIL_TEMPLATE xmlns="http://tempuri.org/">
                      <WebURL>${_webUrl}</WebURL>
                      <Email_Name>${emailName}</Email_Name>
                      <Query><![CDATA[${query}]]></Query>
                      <UserDisplayName>${CurrentUser.Title}</UserDisplayName>
                      <CurrentUserEmail>${CurrentUser.Email}</CurrentUserEmail>
                      <ApprovalCC>${approvalCC}</ApprovalCC>
                      <CheckPageRootFolder></CheckPageRootFolder>
                      <ModuleName>DTRD</ModuleName>
                      <Notification_Name></Notification_Name>
                    </SEND_EMAIL_TEMPLATE>
                  </soap:Body>
                </soap:Envelope>`
  let res = await getSoapResponse('POST', serviceUrl, false, soapContent, 'SEND_EMAIL_TEMPLATEResult');
}

const setStatuses = async function(){
  let status = statuses[0];

  statuses.map(status=>{
     if(status === 'Sent to Champion')
       isSentToChampion = true;
  })

  if(!isChamp){
    let mesg = '';
    if( status === 'Sent to Champion')
      mesg = 'items already sent to champ';
    else if(status === 'Completed')
      mesg = 'items already Completed';
  
      if(mesg !== ''){
        setErrorMessage(mesg, 'Submit');
        dimSubmit = true;
      }
  }
}