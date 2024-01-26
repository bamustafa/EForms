var _hot, _container;
var _data = [];
var batchSize = 30;

var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _htLibraryUrl;
var _list, _RLOD = 'RLOD', _SLOD = 'SLOD', _itemId, _lodRef = '', _status = '', _colArray = [], _requiredFields = [], _targetList, _filterField, _dataArray;
var _isSiteAdmin = false, _isMain = true, _isLead = false, _isPart = false, _isNew = false, _isEdit = false, _isMultiContracotr = false, _ignoreChange = false, _isAllowed = false,
    _updateTitle = false;
var delayTime = 100, retryTime = 10, _timeout;

var mtItem, _fncSchemas= {}, _masterErrors = {};
var defaultClassName = 'TransparentRow htMiddle';
var closedClassName = 'ClosedRow htMiddle';
var errorClassName = 'ErrorRow htMiddle';

var defaultMetaObject = {
    className: defaultClassName
};

var errorMetaObject = {
  className: errorClassName
};

var closeMetaObject = {
    readOnly: true,
    className: closedClassName
};

var _revNumStart, contractorGroupName;

var onRender = async function (relativeLayoutPath, moduleName, formType){

    await getGlobalParameters(relativeLayoutPath, moduleName, formType);

    if(_isMultiContracotr){
      _isAllowed = await IsUserInGroup(contractorGroupName);
      if(!_isSiteAdmin && !_isAllowed){
        alert('You are not allowed to submit LOD.');
        fd.close();
      }
    }

    await renderControls();

    await ensureFunction('getGridMType', _web, _webUrl, _module); //set mtItem
    var colsInternal = [], colsType = [];
    if(mtItem.colArray.length > 0){
        _colArray = mtItem.colArray;
        _targetList = mtItem.targetList;
        _filterField = mtItem.targetFilter;
        _revNumStart = mtItem.revNumStart;
        _updateTitle = mtItem.updateTitle;
    
        _colArray.map(item =>{
            colsInternal.push(item.data);
            colsType.push(item.type);

            if(item.allowEmpty === false)
            _requiredFields.push(item.data);
        });
    }

    if(_isNew){
      validateListFields(colsInternal, colsType);
      _lodRef = 'NA';
    }
    else _lodRef = fd.field("Title").value;

    if(!_isNew)
     _data = await _getData(_web, _RLOD, _filterField, _lodRef , colsInternal, colsType);
    
    if(_data.length < batchSize){
        var remainingLength = batchSize - _data.length;
        for (var i = 0; i < remainingLength; i++) {
          var rowData = { id: i + 1, value: 'Row ' + (i + 1) }
          _data.push(rowData);
        }
    }

    _spComponentLoader.loadScript(_htLibraryUrl).then(_setData);   
    $('.handsontable .htDimmed').addClass('ErrorMesg');
}

var _getData = async function(_web, _listname, _filterfld, _refNo, _colsInternalArray, _colsTypeArray){
    var _itemArray = [];
    var _dataArray = [];

    var _query = _filterfld + " eq '" + _refNo + "'";
    var filterArray = _colsInternalArray.filter(item => item !== 'Mesg' && item !== 'Status');
    const _cols = filterArray.join(',');

    var items = await _web.lists.getByTitle(_listname).items.filter(_query).select(_cols).getAll();
    
    if(items.length === 0)
     items = await _web.lists.getByTitle(_SLOD).items.filter(_query).select(_cols).getAll();
         
     if(items.length > 0){
        items.forEach(item => {
            for(var j = 0; j < _colsInternalArray.length; j++){
                var _type = _colsTypeArray[j];
                var _colname = _colsInternalArray[j];
                var _value = item[_colname];

                if(_colname === 'Status'){
                    var iconUrl =  _layout + '/Images/Submitted.png'; //_value;
                    _value  = "<img src='" + iconUrl + "' alt='na'></img>";
                }
                _itemArray[_colname] = _value;
            }
            _dataArray.push(_itemArray);
            _itemArray = [];
        });
     }
     return _dataArray;
}

const _setData = (Handsontable) => {

    var contextMenu = contextMenu = ['row_below', '---------', 'remove_row'];
    //Handsontable.validators.registerValidator('getCellValidator', getCellValidator);

     _container = document.getElementById('dt');
	 _hot = new Handsontable(_container, {
		data: _data,
        columns: _colArray,
        contextMenu: contextMenu,
        width:'100%',
        height: '600',
        filters: true,
        filter_action_bar: true,
        dropdownMenu: ['filter_by_condition', 'filter_by_value', 'filter_action_bar'],
        rowHeaders: true,
        colHeaders: true,
        //manualColumnResize: true,
        stretchH: 'all',
        licenseKey: htLicenseKey
	});

    performAfterChangeActions((rowIndex) => {
        // Code to be executed after afterChange is completed
        console.log('AfterChange is completed.');

        _hot.batch(() => {
            for (var key in _masterErrors) {
               var mesg = _masterErrors[key];
               setErrorMesg(key, mesg);
               if(mesg === '' || mesg === 'empty')
                removeError(key);
            }
        });
        
        if(Object.keys(_masterErrors).length > 0)
           $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#737373').attr("disabled", "disabled");
        else $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#444').removeAttr('disabled');

        fixButtonHover();
        _ignoreChange = false;
        console.log('done after update');
        remove_preloader();
    });
    
    if(!_isNew)
     setRowsReadOnly();
}

async function performAfterChangeActions(callback) {
    var isLoadingData = false;
    function addRowsInBatch(startIndex, endIndex) {
        var newData = []; 

        for (var i = startIndex; i < endIndex; i++) {
            var rowData = { id: i + 1, value: 'Row ' + (i + 1) }; 
            newData.push(rowData);
        }

        _data = _data.concat(newData); // Concatenate the new data with the existing data array

        // Update the settings to reflect the changes
        _hot.updateSettings({
            data: _data
        });
        isLoadingData = false; // Reset the flag variable
    }

    function handleScroll() {
     
        var scrollElement = _container.querySelector('.wtHolder');
        var scrollPosition = scrollElement.scrollTop;
        //console.log('scrollPosition = ' + scrollPosition);
  
        //var visibleHeight = _container.offsetHeight;
        var totalHeight = scrollElement.scrollHeight - _container.scrollHeight; // - 50;
  
        // Calculate the scroll position at the bottom of the table
        //var bottomScrollPosition = totalHeight - visibleHeight;
      
        // Check if the user has scrolled to the bottom of the table
         if (!isLoadingData && scrollPosition >= totalHeight) {
             isLoadingData = true;
  
          var startIndex = _hot.countRows();
          var endIndex = startIndex + batchSize;
          //if(!isContractor && !_isTeamLeader) 
           addRowsInBatch(startIndex, endIndex);
        }
    }

    var isFilenameExist = async function(deliverableType, filename){
      var isFound = false;
      var schemDelivType = _fncSchemas[deliverableType];
      
      if(schemDelivType.containsRev)
        filename = filename.substring(0, filename.lastIndexOf('-'));

      var _query = "FileName eq '" + filename + "'";

       const list = _web.lists.getByTitle(_RLOD);
       const items = await list.items.select("Id,LODRef").filter(_query).top(1).get();
       if (items.length > 0)
         return {
            isFound: true,
            LODRef: items[0].LODRef
         }
       else return {
        isFound: false
       }
    }

    _hot.addHook('afterScrollVertically', handleScroll);

    _hot.addHook('afterChange', async (changes, source) => {
        
        if(_ignoreChange || !changes.hasOwnProperty('length'))
           return;

         if(changes.length > 5)
          runPreloader();

            const latestColumns = {};
            for (const [value, column] of changes) {
                latestColumns[value] = column;
            }

          if (source === 'edit' || source === 'Autofill.fill' || source === 'CopyPaste.paste') {
            var filename = '';
            const promises = changes.map(async ([rowIndex, prop, oldValue, value]) => {
              const columnIndex = _hot.propToCol(prop);
              const item = _hot.getCellMeta(rowIndex, columnIndex);
              const fnColumnIndex = _hot.propToCol('FullFileName');

              if(prop === 'FullFileName')
                 filename = value;
              else filename = _hot.getDataAtCell(rowIndex, fnColumnIndex);
              
              if ( (filename !== null && filename !== '') && _requiredFields.includes(item.data)) {
                var mesg = validateFields(rowIndex, prop, value, item);
        
                if (mesg === ''){
                    if(latestColumns[rowIndex] === prop){
                        var fname = filename;
                        filename = '';

                        if(fname !== ''){
                            const delivColIndex = _hot.propToCol('DeliverableType');
                            var deliverableType = _hot.getDataAtCell(rowIndex, delivColIndex);
                            var result = await setNamingConvention(rowIndex, columnIndex, fname, deliverableType);
                            mesg = result.mesg;

                            if(mesg === '' && !_hot.getCellMeta(rowIndex, fnColumnIndex).readOnly){
                                var res = await isFilenameExist(deliverableType, fname);
                                if(res.isFound === true)
                                 mesg = `${fname} is submitted previously in ${res.LODRef}`;
                            }
                            
                            return{
                                rowIndex: rowIndex,
                                mesg: mesg
                            }

                       }
                    }
                }
                else {
                    return{
                        rowIndex: rowIndex,
                        mesg: mesg
                    }
                }
              }
              else{
                const msgColumnIndex = _hot.propToCol('Mesg');
                var mesgValue = _hot.getDataAtCell(rowIndex, msgColumnIndex);
                if(mesgValue !== null && mesgValue !== ''){
                    return{
                        rowIndex: rowIndex,
                        mesg: 'empty'
                    }
                }
              }
            });

        const results = await Promise.all(promises);
        results.forEach((result, index) => {
            if (result !== undefined && result !== ''){
                if(result.mesg !== '')
                  addError(result.rowIndex, result.mesg);
                else{
                    const mesgColumnIndex = _hot.propToCol('Mesg');
                    var mesgValue = _hot.getDataAtCell(result.rowIndex, mesgColumnIndex);
                    if(mesgValue !== null && mesgValue !== '')
                      addError(result.rowIndex, '');
                      else{
                        const statusColumnIndex = _hot.propToCol('Status');
                        var mesgValue = _hot.getDataAtCell(result.rowIndex, mesgColumnIndex);
                        if(mesgValue === null || mesgValue === '')
                           addError(result.rowIndex, '');
                      }
                }
            }
        });

        // Call the callback function after afterChange is completed
        _ignoreChange = true;
        if (typeof callback === 'function') {
            callback();
        }
     }
    });

     _hot.addHook('afterRemoveRow', function(index, amount, changes, source) {     
        changes.map(index =>{
            removeError(index);
        });   

        if(Object.keys(_masterErrors).length > 0)
          $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#737373').attr("disabled", "disabled");
        else $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#444').removeAttr('disabled');

        fixButtonHover();
    });
}

//#region GENERAL
var renderControls = async function(){
    var isValid = false;
    var retry = 1;
    while (!isValid)
    {
      try{
        if(retry >= retryTime) break;

         await setFormHeaderTitle();
         await setButtons();
        
         isValid = true;
      }
        catch{
          retry++;
          await delay(delayTime);
        }
    } 

    var fields = ['Title','Status'];
    HideFields(fields, true);
}

var setButtons = async function () {
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

    await setButtonActions("Accept", "Submit");
    await setButtonActions("ChromeClose", "Close");
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
            runPreloader();
            if(_isNew){
              _lodRef = 'LOD-' + await getCounter(_web, "LOD");
              fd.field('Title').value = _lodRef;
              fd.field('Status').value = 'Submitted';
            }
            await insertLODItems();
            fd.save();
           }
       }
    });
}

function delay(time) {
    return new Promise(function(resolve) { 
        setTimeout(resolve, time)
    });
}

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
                        _layout = value[0].children[0].textContent;
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

var getGlobalParameters = async function(relativeLayoutPath, moduleName, formType){
    localStorage.clear();  

    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;
    _layout = relativeLayoutPath;

    if(_formType === 'New'){
        fd.clear();
        _isNew = true;
    }
    else if(_formType === 'Edit'){
        _isEdit = true;
        _itemId = fd.itemId;
        _lodRef = fd.field('Title').value;
        _status = fd.field('Status').value;
        $(fd.field('Status').$parent.$el).hide();
    }

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;

    // var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GLOBAL_PARAM';
    // var soapContent;
    // soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    //                 '<soap:Body>' +
    //                   '<GLOBAL_PARAM xmlns="http://tempuri.org/" />' +
    //                 '</soap:Body>' +
    //               '</soap:Envelope>';
    // await getSoapRequest('POST', serviceUrl, false, soapContent);

    _errorImg = _layout + '/Images/Error.png';
    _submitImg = _layout + '/Images/Submitted.png';
    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

    var script = document.createElement("script");
    script.src = _layout + "/plumsail/js/config/lodConfig.js";
    document.head.appendChild(script);  

    var types = ['DWG','DOC','TRM','GEN'];
    types.forEach(async function(type) {
        var items = await _web.lists.getByTitle("FNC").items.select("Delimeter,Schema").filter(`Title eq '${type}'`).get();
        if(items.length > 0){

            var item = items[0];
            var delimeter = item.Delimeter;
    
            var schema = item.Schema;
            schema = schema.replace(/&nbsp;/g, '');
            schema = JSON.parse(schema);
    
            var scheamResult = await setFilenameText(schema, delimeter);
     
            var containsRev = false;
            schema.map(fld => {
                var fieldName = fld.InternalName.toLowerCase();
                if(fieldName === 'rev' || fieldName === 'revision'){
                  containsRev = true;
                  return true;
                }
            });

            if(scheamResult.filenameText !== ''){
                var majorItems = await _web.lists.getByTitle("MajorTypes").items.select("Category").filter(`Title eq '${type}'`).get();
                var category = '';
                if(majorItems.length > 0)
                  category = majorItems[0].Category;
                
                _fncSchemas[type] = {
                    delimeter: delimeter,
                    schemaFields: schema, // replace with actual values
                    filenameText: scheamResult.filenameText.slice(0, -1), // replace with actual values
                    category: category,
                    containsRev: containsRev
                };
            }
        }
    });

    await ensureFunction('isMultiContractor');
    fixButtonHover(); 
}

var ensureFunction = async function(funcName, ...params){
    var isValid = false;
    var retry = 1;
    while (!isValid)
    {
        try{
          if(retry >= retryTime) break;

          if(funcName === 'getGridMType')
            mtItem = await getGridMType(...params);
          else if(funcName === '_setData')
             _setData();
          else if(funcName === 'preloader')
             preloader();
          else if(funcName === 'isMultiContractor')
          contractorGroupName = await isMultiContractor();
           isValid = true;
        }
        catch{
          retry++;
          await delay(delayTime);
        }
    }
}

var  runPreloader = async function(){
	try {
		var targetControl = $('#ms-notdlgautosize').addClass('remove-position-preloadr');
		//var mainDimmerElement = 'pageLayout_5a558a10,pageLayoutDesktop_5a558a10'; //'div.ControlZone-control'//'div.fd-toolbar-primary-commands';
	    var webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
		
        var _layout = '/_layouts/15/PCW/General/EForms';
		var ImageUrl = webUrl + _layout + '/Images/Loading.gif';
		targetControl.dimBackground({
					darkness: 0.7
				}, function () {});

        // $("<img id='loader' src='" + ImageUrl + "' />")
        //     .css({
        //         "position": "absolute",
        //         "top": "300px",
        //         "left": "790px",
        //         "width": "100px",
        //         "height": "100px"
        //     }).insertAfter(targetControl);
	}
   catch(err) { console.log(err.message); }
}

function remove_preloader(){
    var targetControl = $('#ms-notdlgautosize')
    targetControl.undim();

    // if($('div.dimbackground-curtain').length > 0)
    //   $('div.dimbackground-curtain').remove();

    // if($('#loader').length > 0)
    //     $('#loader').remove();
}

var getCounter = async function(_web, key){
    var value = 1;
    var listname = 'Counter';

    await _web.lists
          .getByTitle(listname)
          .items
          .select("Id, Title, Counter")
          .filter("Title eq '" + key + "'")
          .get()
          .then(async function (items) {
            var _cols = { };
              if(items.length == 0){
                 _cols["Title"] = key;
                 _cols["Counter"] = value.toString();
                 value = String(value).padStart(4,'0');
                 var blabla = await _web.lists.getByTitle(listname).items.add(_cols);
                 return value;
              }

              else if(items.length > 0)
                    var _item = items[0];
                    value = parseInt(_item.Counter) + 1;
                    _cols["Counter"] = value.toString();
                    value = String(value).padStart(4,'0');
                    var blabla = await _web.lists.getByTitle(listname).items.getById(_item.Id).update(_cols);
                    return value;
              });
   return value;
}

function HideFields(fields, isHide){
	var field;
	for(let i = 0; i < fields.length; i++)
	{
		field = fd.field(fields[i]);
		if(isHide || isHide == undefined)
		  $(field.$parent.$el).hide();
		else $(field.$parent.$el).show();
	}
}

var setRowsReadOnly = async(value, rowIndex, colIndex) => {
      const rowCount = _hot.countRows();
      for (var row = 0; row < rowCount; row++) {
        var columnIndex = 0;

        const fnColumnIndex = _hot.propToCol('FullFileName');
        var filename = _hot.getDataAtCell(row, fnColumnIndex);

        if(filename !== undefined && filename !== null && filename !== ''){
          _hot.getSettings().columns.find(item => {
              if(_updateTitle && item.data === 'Title'){}
              else {
                //_hot.setCellMeta(row, columnIndex, 'readOnly', true);
                _hot.setCellMetaObject(row, columnIndex, closeMetaObject);
              }
              columnIndex++;
          });
        }
      }   
    _hot.render();
}

function fixButtonHover(){
    $('.btn-outline-primary').hover(
        function() {
            $(this).css('color', 'white');
        },
        function() {
            $(this).css('color', 'black');
        }
    );
}
//#endregion

//#region FNC CHECKER
var setNamingConvention = async function(rowIndex, columnIndex, filename, deliverableType){
    var typeMetaInfo = _fncSchemas[deliverableType];
    var delimeter = typeMetaInfo.delimeter;
    var schemaFields = typeMetaInfo.schemaFields;
    var filenameText = typeMetaInfo.filenameText;
    var checkRev = typeMetaInfo.containsRev;

    var mesg = '';
    const illegalChars = /['{}[\]\\;':,\/?!@#$%&*()+=]/;

    if (!filename.includes(delimeter))
        mesg = 'Unstructured Filename';
    else if(illegalChars.test(filename))
        mesg = 'Filename Contains illegal character';
    else mesg = await checkFileName(deliverableType, delimeter, schemaFields, filenameText, filename, true, checkRev);
    
    return {
        rowIndex: rowIndex,
        mesg: mesg
    }
}

var setFilenameText = async function(schema, delimeter){
    var filenameText = '', schemaFields = '';
    schema.filter(item => {
      var fieldName = item.InternalName;
    //   if(fieldName === 'Rev' || fieldName === 'Revision'){
    //     schemaFields += fieldName + delimeter;
    //     return;
    //   }
        
      if(fieldName !== 'properties'){
        schemaFields += fieldName + delimeter;
        filenameText += fieldName + delimeter;
  
         try {
          fd.field(fieldName).value;
          fd.field(fieldName).disabled = true;
        }
        catch(e){}
      }
     });
     return {
      filenameText: filenameText,
      schemaFields: schemaFields
    };
}

function addError(key, message) {
    //if (!_masterErrors.hasOwnProperty(key))
      _masterErrors[key] = message;
  }
  
function removeError(key) {
    if (_masterErrors.hasOwnProperty(key))
     delete _masterErrors[key];
}

var setErrorMesg = async function(rowIndex, mesg){
    const mesgColumnIndex = _hot.propToCol('Mesg');
    const statusColumnIndex = _hot.propToCol('Status');
    var iconUrl = `<img src='${_layout}/Images/Submitted.png' alt='na'></img>` ;

    if(mesg === 'empty'){
        mesg = '';
        iconUrl = '';
    }
    else{
        if(mesg !== ''){
                iconUrl = _layout + '/Images/Error.png';
                iconUrl = "<img src='" + iconUrl + "' alt='na'></img>";
        }    
    }
    _hot.setDataAtCell(rowIndex, mesgColumnIndex, mesg);
    _hot.setDataAtCell(rowIndex, statusColumnIndex, iconUrl);  
}

function validateFields(rowIndex, prop, value){
    var colIndex = 0;
    var mesg = '';
    
    _colArray.map(item =>{
        const columnIndex = _hot.propToCol(item.data);
        var value = _hot.getDataAtCell(rowIndex, colIndex);

        if(item.source !== undefined && item.source !== null){
            var isCorrect = false;
            if(value !== undefined && value !== null && value !== '' ){
                item.source.map(type =>{
                    if(type === value){
                        isCorrect = true;
                        return isCorrect;
                    }
                });
            }
            else isCorrect = true

            if(!isCorrect){
               if(mesg !== '') mesg += '<br/>';
               mesg += `select a valid ${item.title} from the drop down list`;
            }
        }

        if(item.allowEmpty === false){
            if(value === '' || value === null || value === undefined){
              if(mesg !== '') mesg += '<br/>';
               mesg += `${item.title} is required field`;
            }
        }

        if(item.length !== null && item.length !== undefined && value !== null && value !== undefined){
            if(value.length > parseInt(item.length)){
                if(mesg !== '') mesg += '<br/>';
                mesg += `${item.title} is too long, maximum allowed characters is ${item.length}`;
            }
        }
        colIndex++;
    });
    return mesg;
}

//#endregion

//#region VALIDATE LIST COLUMNS tempLOD/RLOD/SLOD
var getListFields = async function(_web, _listname){
    var fieldInfo = [];
    await  _web.lists
           .getByTitle(_listname)
           .fields
           .filter("ReadOnlyField eq false and Hidden eq false")
           .get()
           .then(function(result) {
              for (var i = 0; i < result.length; i++) {
                  fieldInfo.push(result[i].InternalName);
                  // fieldInfo += "Title: " + result[i].Title + "<br/>";
                  // fieldInfo += "Name:" + result[i].InternalName + "<br/>";
                  // fieldInfo += "ID:" + result[i].Id + "<br/><br/>";
              }
      }).catch(function(err) {
          alert(err);
      });
      return fieldInfo;
}

var createFields = async function(_web, _listname, _field, _type){
    try{
     _type = _type.toLowerCase();
     if(_type === 'text' || _type === 'dropdown')
        await _web.lists.getByTitle(_listname).fields.addText(_field);
     else if(_type === 'multi')
        await _web.lists.getByTitle(_listname).fields.addMultilineText(_field);
     else if(_type === 'calendar')
        await _web.lists.getByTitle(_listname).fields.addDateTime(_field);
        
        await _web.lists.getByTitle(_listname).defaultView.fields.add(_field);
    }
    catch (e) {
       console.log(e);
    }
}

var validateListFields = async function(colsInternal, colsType){
    var lists = [_RLOD, _SLOD];
    for (const currentList of lists) {
        var _listFields = await getListFields(_web, currentList);

        for(var i = 0; i < colsInternal.length; i++){
            var isFldExist = false;
            var _fld = colsInternal[i];
    
            for(var j = 0; j < _listFields.length; j++){
                var _lstFld = _listFields[j];
                if(_fld === _lstFld){
                    isFldExist = true;
                    break;
                }
            }
            if(!isFldExist){
             if(_fld !== 'Mesg')
                await createFields(_web, currentList, _fld, colsType[i]);
            }
        }
      }
}
//#endregion

//#region CRUD OPERATION FOR tempLOD/RLOD/SLOD

var insertLODItems= async function(){

  var mentaInfo = _hot.getSettings().columns;
 
  var rows = _hot.getData();
  var rowCount = rows.length;
  var acceptedRow = 0;
  var emptyRows = 0;
  var itemsToInsert = [];
  var objColumns = [];
  var objTypes = [];

  for(var row = 0; row < rowCount; row++){
    var _objValue = { };
    var doInsert = false;
    var cellValue = rows[row][0];

    if(cellValue!== null && cellValue !== ''){
      acceptedRow++;
      var _columns = '', _colType = '';
      for(var j = 0; j < mentaInfo.length; j++){
        var _value = rows[row][j];
        var internalName = mentaInfo[j].data;

        if( internalName === 'Status' || internalName === 'Mesg') continue;

         if(_value !== null && _value !== ''){
        
          _columns += internalName + ',';
          _colType += mentaInfo[j].type + ',';
          _objValue[internalName] = _value;
         }
      }
      doInsert = true;
    }
    else emptyRows++;

    if(emptyRows > 10)
      break;
    
    if(doInsert){
       var isExist = await isFileExistinArray(row, itemsToInsert, _objValue);
       if(!isExist){
         _columns = _columns.endsWith(',') ? _columns.slice(0, -1) : _columns;
         _colType = _colType.endsWith(',') ? _colType.slice(0, -1) : _colType;

         objColumns.push(_columns);
         objTypes.push(_colType);
         itemsToInsert.push(_objValue);
       }
    }
  }

  await insertItemsInBatches(itemsToInsert, objColumns, objTypes); 
}

var insertItemsInBatches = async function(itemsToInsert, objColumns, objTypes) {
    const list = _web.lists.getByTitle(_RLOD);
  
    const batch = pnp.sp.createBatch();
    
    var row = 0;
    var _columns, _colType;
  
    for (const item of itemsToInsert){
      _columns = objColumns[row];

      var deliverableType = item.DeliverableType;
      var schemDelivType = _fncSchemas[deliverableType];
      var delimeter = schemDelivType.delimeter;
      var schema = schemDelivType.schemaFields;
      var category = schemDelivType.category;

      var _query, filename = item.FullFileName;
      if(schemDelivType.containsRev)
         filename = filename.substring(0, filename.lastIndexOf('-'));
       _query = "FileName eq '" + filename + "'";
  
      const existingItems = await list.items
      .select("Id," + _columns)
      .filter(_query)
      .top(1)
      .get();
  
      if (existingItems.length > 0) {
        var _item = existingItems[0]; // _item is the oldItem values while objValue is the new one
        var _cols = _columns.split(',');
        //var _type = _colType.split(',');
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
      } 
      else {
        item['Category'] = category;
        item['FileName'] = filename;
        item['LODRef'] = _lodRef;
        await splitFileName(schema, delimeter, item)
        list.items.inBatch(batch).add(item);
      }
      row++;
    }
    await batch.execute();
}

var splitFileName = async function(schema, delimeter, item){
    var filename = item.FullFileName.split(delimeter);
    var colIndex = 0;
    var spanFieldValue;
    schema.map(fld =>{
       var field = fld.InternalName;
       if(fld.InternalName !== 'properties'){
            var value = '';

            if(fld.SpanField !== undefined){
              spanFieldValue = fld.SpanField;
              if(spanFieldValue === 2)
                value = filename[colIndex] + delimeter + filename[colIndex+1];
              else if(spanFieldValue === 3)
                value = filename[colIndex] + delimeter + filename[colIndex+1] + delimeter + filename[colIndex+2];
            }
            else if(fld.isOptionalField){
                var filenameLength = filename.length;
                var schemaLength = schema.length -1;

                if(spanFieldValue !== undefined)
                  filenameLength = filenameLength - (spanFieldValue -1);

                if( filenameLength === schemaLength)
                  value = filename[colIndex];
                else return;
            }
            else value = filename[colIndex];

            item[field] = value;

            if(field !== 'properties'){
                if(fld.SpanField !== undefined)
                  colIndex = colIndex + fld.SpanField;
                else colIndex++;
           }
       }
    });
}

var isFileExistinArray = async function(rowIndex, itemsToInsert, _objValue){
    const fnColumnIndex = _hot.propToCol('FullFileName');
    var filename = _hot.getDataAtCell(rowIndex, fnColumnIndex);
    var filenameExists = itemsToInsert.some(item => item.FullFileName === _objValue.FullFileName);
    return filenameExists;
}
//#endregion