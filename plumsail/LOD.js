var _hot, _container;
var _data = [];
const batchSize = 30;
let updateBatchSize = 10;

var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _phase = '', _htLibraryUrl;
var _list, _itemId, _lodRef = '', _status = '', _colArray = [], _requiredFields = [], _targetList, _filterField, _dataArray;
var _isSiteAdmin = false, _isMain = true, _isLead = false, _isPart = false, _isNew = false, _isEdit = false, _isMultiContracotr = false, _ignoreChange = false, _isAllowed = false,
    _updateTitle = false, _isDesign = false, _isMidpExist = false, _isTradeChecker = false, _isMIDPSubjectExist = false, _ignoreChecking = false, _isOriginatorChecker = false;
var delayTime = 100, retryTime = 10, _timeout, _revStartWithNeeded;

var _darTrade = '', _cdsTitle = '', _CheckRevision = 'yes';

var mtItem, _fncSchemas= {}, _masterErrors = {};
var defaultClassName = 'TransparentRow htMiddle';
var closedClassName = 'ClosedRow htMiddle';
var errorClassName = 'ErrorRow htMiddle';
var cancelClassName = 'CancelRow htMiddle';

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

var _revNumStart, contractorGroupName, _exact_Filename_Match, _closingStatus;
let _defaultStatusImgCol = 'TempStatus';
var _updateItems = [];
var _listDictionary = [];

let itemExistMesg = 'filename is already submitted'
let disableTable = false;
var onRender = async function (relativeLayoutPath, moduleName, formType){

    debugger
    _layout = relativeLayoutPath;
    await loadScripts();
    await getLODGlobalParameters(moduleName, formType);

    await renderControls();
    await validateBeforeLoad();

    await ensureFunction('getGridMType', _web, _webUrl, _module, _isDesign); //set mtItem
 
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

    if(_isDesign)
      addLegends();
}

var _getData = async function(_web, _listname, _filterfld, _refNo, _colsInternalArray, _colsTypeArray){
    var _itemArray = [];
    var _dataArray = [];

    var _query = _filterfld + " eq '" + _refNo + "' and Status ne 'Cancelled'";

    if(_isEdit && _isDesign)
        _query += ` and DarTrade eq '${fd.field('Trade').value}'`
     
    var filterArray = _colsInternalArray.filter(item => item !== 'Mesg' && item !== _defaultStatusImgCol);
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

                if(_colname === _defaultStatusImgCol){
                    var iconUrl =  _layout + '/Images/Submitted.png'; //_value;
                    _value  = "<img src='" + iconUrl + "' alt='na'></img>";
                }
                else if(_colname === 'Mesg')
                    _value  = `<p style='color: green'>${itemExistMesg}</img>`;
                _itemArray[_colname] = _value;
            }
            _dataArray.push(_itemArray);
            _itemArray = [];
        });
     }
     return _dataArray;
}

const _setData = (Handsontable) => {

    var contextMenu = ['row_below']; //, '---------', 'remove_row'];
    if(_isDesign){
        contextMenu = {
            callback(key, selection, clickEvent) {
              // Common callback for all options
              console.log(key, selection, clickEvent);
            },
            items: {
              row_below: {
                name: 'Click to add row below' // Set custom text for predefined option
              },
            //   remove_row: {
            //     name: "Remove Row", // The text for the menu item
            //     callback: async function (key, selection, clickEvent) {
            //       // Get the selected row index
            //       debugger;
            //       var selectedRow = selection[0].start.row;
            //       const msgColIndex = _hot.propToCol('Mesg');
            //       var _mesg = _hot.getDataAtCell(selectedRow, msgColIndex);
            //       if (!_mesg.includes('is submitted previously in'))
            //         await isItemFound(selectedRow);
    
            //        //Remove the row from the data
            //       _hot.alter("remove_row", selectedRow);
    
            //        //Render the changes
            //       _hot.render();
            //     },
            //   },
              cancel: {
                name: 'Cancel',
                 callback(key, selection, clickEvent) {
                  var startInex = selection[0].start.row;
                  var endInex = selection[0].end.row;
                  var rowData = _hot.getCellMeta(startInex, 1);

                  //setRowsReadOnly(startInex,true);
                  isItemFound(startInex, 'Cancelled');
                }
              },
              // uncancel: {
              //   name: 'Uncancel',
              //   callback(key, selection, clickEvent) {
              //     var startInex = selection[0].start.row;
              //     var endInex = selection[0].end.row;
              //     var setIndexes = [];
              //     setIndexes.push[startInex];

              //     var rowData = hot.getCellMeta(startInex, 1);
              //     setRowsReadOnly(startInex, false);
              //   }
              // },
            }
          }
    }
    //Handsontable.validators.registerValidator('getCellValidator', getCellValidator);

     _container = document.getElementById('dt');
	 _hot = new Handsontable(_container, {
		data: _data,
        columns: _colArray,
        contextMenu: contextMenu,
        width:'100%',
        height: '500',
        filters: true,
        filter_action_bar: true,
        dropdownMenu: ['filter_by_condition', 'filter_by_value', 'filter_action_bar'],
        rowHeaders: true,
        colHeaders: true,
        //manualColumnResize: true,
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

        if(disableTable){
          _hot.updateSettings({
            allowInsertRow: false, // Disable adding rows
            fillHandle: false, // disable drag
            readOnly: true    // Disable cell editing
            });
        }

    performAfterChangeActions((rowIndex, operation) => {

        if(operation === 'ContextMenu'){
            return rowIndex;
        }
        else{
            // Code to be executed after afterChange is completed
            console.log('AfterChange is completed.');

            _hot.batch(() => {
                var removedIndexes = [];
                for (var key in _masterErrors) {
                var mesg = _masterErrors[key];
                setErrorMesg(key, mesg, removedIndexes);
                    if(mesg === '' || mesg === 'empty'){
                        removeError(key);
                    }
                }
                if(removedIndexes.length > 0){
                    _hot.alter("remove_row", removedIndexes[0], removedIndexes.length);
                    //_hot.alter("remove_row",1);
                    removedIndexes = [];
                }
            });

            if(Object.keys(_masterErrors).length > 0)
            $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#737373').attr("disabled", "disabled");
            else $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#444').removeAttr('disabled');

            fixButtonHover();
            _ignoreChange = false;
            console.log('done after update');
            if($('#totalId').length > 0)
              $('#totalId').remove();
            remove_preloader();
            
            
            // setTimeout(() => {
            //     _hot.render();
            // }, 200);
        }
    });
    
   if(!_isNew)
     setRowsReadOnly();


     setTimeout(() => {
        _hot.render();
    }, 200);
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

    function customizeContextMenu(options) {
        if(!_isDesign)
         return options;
    
        var selectedCell = _hot.getSelectedLast();
        var rowIndex = selectedCell[0];
        // const fnColumnIndex = _hot.propToCol('FullFileName');
        // var filename = _hot.getDataAtCell(rowIndex, fnColumnIndex);

        let statusIndex = _hot.propToCol('Status');
        let status = _hot.getDataAtCell(rowIndex, statusIndex);

        
        //if (Status == "pending" || Status == "rejected")
        // let isFound = false;
        // const items = _web.lists.getByTitle(_RLOD).items.select("Id, Status").filter(`FullFileName eq '${filename}'`).top(1).get()
        // .then(items=> {
        //     isFound = true;
        //     status = items[0]['Status'];
        // })
        // .then(()=>{
            
        // });

        if(!_isNew){
            options.items.forEach((feature) => {
                if(feature.key === 'cancel'){
                    if(status !== null && status.toLowerCase() !== 'pending' && status.toLowerCase() !== 'rejected')
                      feature.disabled = true;
                    else feature.disabled = false;    
                }
            });
        }
        return options;
    }
    _hot.addHook('afterContextMenuDefaultOptions',  customizeContextMenu); //


    var isFilenameExist = async function(deliverableType, filename, subject, isExactMatch){
      var isFound = false;
      let CurrentRevisionNumber = 0;
      let fullfilename = filename;
      if(!_ignoreChecking){
        var schemDelivType = _fncSchemas[deliverableType];
        
        if(schemDelivType.containsRev){      
            CurrentRevisionNumber = filename.substring(filename.lastIndexOf('-') + 1); // part after the last '-'            
            filename = filename.substring(0, filename.lastIndexOf('-'));
        }
      }     

       var _query = isExactMatch ? `FullFileName eq '${fullfilename}'`: `FileName eq '${filename}'`;

       const list = _web.lists.getByTitle(_RLOD);
       var items = await list.items.select("Id, Revision, FullFileName, Status, Title, "+_filterField).filter(_query).top(1).get();
       if (items.length > 0)
       {
            let item = items[0];
            return {
                isFound: true,
                LODRef: item[_filterField],//.LODRef
                rlodItem: item
            }
        }
        else {      
            _query = `FileName eq '${filename}'`;
            items = await list.items.select("Id, Revision, FullFileName, Status, Title, "+_filterField).filter(_query).top(1).get();
            if (items.length > 0)
            {
                let item = items[0];
                return {
                    isFound: false,
                    ExistRev: item['Revision'],//.LODRef
                    RevisionNumber: CurrentRevisionNumber
                }
            }
            else{
                return { isFound: false }
            }
        }
    }
    

    _hot.addHook('afterScrollVertically', handleScroll);

    _hot.addHook('afterChange', async (changes, source) => {
        
        const startTime = performance.now();
        if(_ignoreChange || changes === null || !changes.hasOwnProperty('length')) //_ignoreChecking
           return;        

        const latestColumns = {};
        for (const [value, column] of changes) {
            latestColumns[value] = column;
        }

        var totalRows = Object.keys(latestColumns).length;
        if(totalRows > 5)
         {
            runPreloader();
         }

        if (source === 'edit' || source === 'Autofill.fill' || source === 'CopyPaste.paste'){
            var filename = '';
            const delivIndex = _hot.propToCol('DeliverableType');
            //let removeRow = true, currentIndex;
            const promises = changes.map(async ([rowIndex, prop, oldValue, value]) => {
                var mesg = '';
                const columnIndex = _hot.propToCol(prop);
                const item = _hot.getCellMeta(rowIndex, columnIndex);
                const fnColumnIndex = _hot.propToCol('FullFileName');

                if(prop === 'FullFileName')
                     filename = value;
                else filename = _hot.getDataAtCell(rowIndex, fnColumnIndex);

                    if(mesg == '' && filename !== null && filename.includes(' '))
                        mesg = 'Space is not allowed';

                if(_ignoreChecking){
                    if ( (filename !== null && filename !== '') && _requiredFields.includes(item.data)) {
                            mesg = validateFields(rowIndex, prop, value);
                            if(mesg === ''){
                                let result = await isFilenameExist('', filename);
                                if(result.isFound === true){
                                    if(result.rlodItem.Status === 'Cancelled')
                                      mesg = 'The filename has been canceled. To restore it, add the filename and submit as usual.';
                                }
                                return{
                                    rowIndex: rowIndex,
                                    mesg: mesg
                                }
                            }
                            else {
                                return{
                                    rowIndex: rowIndex,
                                    mesg: mesg
                                }
                            }
                    }
                }
                else{
                    if ( (filename !== null && filename !== '') && _requiredFields.includes(item.data)) {
                        mesg = validateFields(rowIndex, prop, value);

                        if (mesg === ''){
                            if(latestColumns[rowIndex] === prop){
                            
                                var fname = filename;
                                filename = '';

                                if(fname !== ''){                                 
                                    let subject = _hot.getDataAtCell(rowIndex, _hot.propToCol('Title'));
                                    const delivColIndex = _hot.propToCol('DeliverableType');
                                    var deliverableType = _hot.getDataAtCell(rowIndex, delivColIndex);
                                    var result = await setNamingConvention(rowIndex, columnIndex, fname, deliverableType);
                                    mesg = result.mesg;                          
                                    let res, isFound, isVisited = false;
                                    if(_exact_Filename_Match === 'yes' && mesg !== ''){
                                        if(mesg.toLowerCase().startsWith('revision number')){
                                            res = await isFilenameExist(deliverableType, fname, '', true);
                                            isFound = res.isFound;                                            
                                            let revStart = _revStartWithNeeded;                                   
                                            let CuRev = parseInt(res.RevisionNumber, 10); // parse as base 10 integer
                                            let ExRev = parseInt(res.ExistRev, 10);       // parse as base 10 integer
                                            
                                            if (revStart) {                                               
                                                CuRev = parseInt(res.RevisionNumber.replace(revStart, ''), 10);
                                                ExRev = parseInt(res.ExistRev.replace(revStart, ''), 10);
                                            }

                                            if(isFound === true)
                                                mesg = itemExistMesg;
                                            else{
                                                if(!isNaN(CuRev) && !isNaN(ExRev) && (CuRev - (ExRev + 1) !== 0))
                                                    mesg = `revision number must be ${ExRev + 1} instead of ${CuRev}`;
                                                else if(isNaN(CuRev) && isNaN(ExRev)){}
                                                else 
                                                    mesg = '';
                                            }
                                        isVisited = true;
                                        }
                                        else if(mesg.toLowerCase().startsWith('revision must'))
                                        mesg = mesg.substring(0, mesg.indexOf('<br>'));
                                    }

                                    if(mesg === '' && !_hot.getCellMeta(rowIndex, fnColumnIndex).readOnly){
                                        if(!isVisited){
                                            res = await isFilenameExist(deliverableType, fname, subject);
                                            isFound = res.isFound;
                                        }
                                        if(isFound === true){

                                        if(res.rlodItem.Status === 'Cancelled')
                                            mesg = 'filename is already Cancelled. Contact IT admin to re-open';

                                        else if(_exact_Filename_Match === 'yes'){
                                            
                                            mesg = await validateExactFileNameMatch(deliverableType, subject, fname, res.rlodItem);
                                            if(mesg === undefined || mesg === null)
                                            mesg = '';
                                        }
                                        else mesg = `${fname} is submitted previously in ${res.LODRef}`;
                                        }
                                    }
                                    // else if(!_isDesign && _updateTitle && prop === 'Title'){
                                    //     if(value.toLowerCase() !== oldValue.toLowerCase()){
                                    //         let tempfilename = fname;
                                    //         if(!_ignoreChecking){
                                    //             var schemDelivType = _fncSchemas[deliverableType];
                                                
                                    //             if(schemDelivType.containsRev)
                                    //               tempfilename = fname.substring(0, fname.lastIndexOf('-'));
                                    //         }
                                        
                                    //         var _query = "FileName eq '" + tempfilename + "'";
                                    //         const list = _web.lists.getByTitle(_RLOD);
                                    //         const items = await list.items.select("Id").filter(_query).top(1).get();

                                    //         if(items.length > 0){
                                    //             var _objValue = { };
                                    //             _objValue['Title'] = subject;
                                    //             _updateItems.push({
                                    //                 id: items[0].Id, 
                                    //                 metaInfo: _objValue
                                    //             });
                                    //         }
                                    //     }
                                    // }

                                    if(_isDesign && mesg == ''){
                                        if(_isMidpExist){
                                            mesg = await checkMIDPIfExist(deliverableType, fname, subject);
                                        }

                                        if(_isTradeChecker && mesg == ''){
                                            mesg = await checkTradeChecker(deliverableType, fname, subject);
                                        }

                                        if(_isOriginatorChecker && mesg == ''){
                                            mesg = await checkOriginatorChecker(deliverableType, fname, subject);
                                        }
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
                        //const msgColumnIndex = _hot.propToCol('Mesg');
                        //var mesgValue = _hot.getDataAtCell(rowIndex, msgColumnIndex);
                        //if(mesgValue !== null){ // && mesgValue !== ''){
                            return{
                                rowIndex: rowIndex,
                                mesg: ''
                            }
                        //}
                    }
                }
            });
        

        let results = await Promise.all(promises);
        results = results.filter((item)=>{
            return item !== undefined;
        });

        results.forEach((result, index) => {
            if (result !== undefined && result !== ''){
                if(result.mesg !== '')
                  addError(result.rowIndex, result.mesg);
                else{
                    const mesgColumnIndex = _hot.propToCol('Mesg');
                    var mesgValue = _hot.getDataAtCell(result.rowIndex, mesgColumnIndex);
                    if(mesgValue === null)
                        mesgValue = '';
                    if(mesgValue.includes(itemExistMesg)){}
                    else if(mesgValue !== null && mesgValue !== '')
                      addError(result.rowIndex, '');
                    else{
                        const statusColumnIndex = _hot.propToCol('Status');
                        var mesgValue = _hot.getDataAtCell(result.rowIndex, mesgColumnIndex);
                        if(mesgValue === null || mesgValue === '' || mesgValue.includes(itemExistMesg)){
                           addError(result.rowIndex, '');

                        //    let cellElement = _hot.getCell(result.rowIndex,2);
                        //     if(cellElement !== null && cellElement.className !== 'TransparentRow'){
                        //         cellElement.className = 'TransparentRow';
                        //     }
                        }
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
          const endTime = performance.now();
          const elapsedTime = endTime - startTime;
          console.log(`Execution time onRender: ${elapsedTime} milliseconds`);
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
var loadScripts = async function(){
    const libraryUrls = [
        //_layout + '/controls/detailedLoader/dist/resloader.js',
        _layout + '/plumsail/js/commonUtils.js',
        _layout + '/plumsail/js/customMessages.js',
        _layout + '/controls/preloader/jquery.dim-background.min.js',
        _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
        _layout + '/plumsail/js/preloader.js'
      ];

    const cacheBusting = `?v=${Date.now()}`;
      libraryUrls.map(url => { 
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
        });
        
    const stylesheetUrls = [
        _layout + '/controls/tooltipster/tooltipster.css',
        _layout + '/controls/handsonTable/libs/handsontable.full.min.css',
        _layout + '/plumsail/css/CssStyle.css'
        ];
  
    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
    });
}

var renderControls = async function(){
    var isValid = false;
    var retry = 1;
    //while (!isValid)
    //{
      //try{
       // if(retry >= retryTime) break;

         await setFormHeaderTitle();
         await setButtons();
         //isValid = true;
      //}
        // catch{
        //   retry++;
        //   await delay(delayTime);
        // }
    //} 

    var fields = ['Title','Status'];
    HideFields(fields, true);
    setToolTipMessages();
}

var setButtons = async function () {
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

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
            detailedLoader(null, null, true);
            if(_isNew){
              _lodRef = 'LOD-' + await get_Counter(_web, "LOD");
              fd.field('Title').value = _lodRef;
              fd.field('Status').value = 'Submitted';
            }
            else{
                if(_isDesign)
                  fd.field('Status').value = 'Submitted';
            }
             insertLODItems().then(() => {
                fd.save();
             }).catch(error => {
                console.error("Error inserting LOD items:", error);
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

var getLODGlobalParameters = async function(moduleName, formType){
    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;

    _phase = await getParameter('Phase');
    if(_phase.toLowerCase() === 'design'){
       _isDesign = true;
       updateBatchSize = 50;

        var Insert_Whatever  = await getParameter('Insert_Whatever');
        if(Insert_Whatever !== "" && Insert_Whatever.toLocaleLowerCase() === 'yes' )
          _ignoreChecking = true;

        _CheckRevision = await getParameter('CheckRevision');
        if(_CheckRevision === '') 
           _CheckRevision = 'yes';
    }

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

         if(_phase.toLowerCase() === 'design'){
        //     $(fd.field('Trade').$parent.$el).show();
             _darTrade = fd.field('Trade').value;
        //     $(fd.field('Trade').$parent.$el).hide();
  
        //     $(fd.field('CDSTitle').$parent.$el).show();
             _cdsTitle = fd.field('CDSTitle').value;
        //     $(fd.field('CDSTitle').$parent.$el).hide();
         }
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

    // var script = document.createElement("script");
    // script.src = _layout + "/plumsail/js/config/lodConfig.js";
    // document.head.appendChild(script);  

    var types = ['DWG','DOC','TRM','GEN', 'MOD'];
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
            var revStartWith;
            schema.map(fld => {
                var fieldName = fld.InternalName.toLowerCase();
                if(fieldName === 'rev' || fieldName === 'revision'){
                  containsRev = true;
                  revStartWith = fld.RevStartWith;
                  _revStartWithNeeded = revStartWith;
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
                    containsRev: containsRev,
                    revStartWith: revStartWith
                };
            }
        }
    });

    if(_isDesign){
        //_defaultStatusImgCol = 'TempStatus';
        _exact_Filename_Match = await getParameter("Exact_Filename_Match");
        _exact_Filename_Match = _exact_Filename_Match.toLowerCase();

        let TradeChecker = await getParameter('TradeChecker');
        if(TradeChecker.toLowerCase() === 'yes')
            _isTradeChecker = true; 
        
        let OriginatorChecker = await getParameter('OriginatorChecker');
        if(OriginatorChecker.toLowerCase() === 'yes')
            _isOriginatorChecker = true;   

        if(_exact_Filename_Match === 'yes'){
            let wfItems = await _web.lists.getByTitle("WorkflowSteps").items.select("ApprovedStatus").filter(`Title eq 'Client'`).get();
            if(wfItems.length === 0 || wfItems[0].ApprovedStatus === null)
            wfItems = await _web.lists.getByTitle("WorkflowSteps").items.select("ApprovedStatus").filter(`Title eq 'Area'`).get();

            if(wfItems.length > 0){
                var wfItem = wfItems[0];
                _closingStatus = (wfItem.ApprovedStatus !== null && wfItem.ApprovedStatus !== undefined) ? wfItem.ApprovedStatus : 'Reviewed';
            }
            else _closingStatus = 'Reviewed';
        }
        
        _web.lists.getByTitle(_MIDP).get().then(list => {
            _isMidpExist = true;
        })
        .then(()=>{
            return _web.lists.getByTitle(_MIDP).fields.getByInternalNameOrTitle('Subject').get().then(field => {
                _isMIDPSubjectExist = true;
            });
        })
        .catch(function(err) {
            //alert(err);
        });
    }

    await ensureFunction('isMultiContractor');
    fixButtonHover(); 
}

var validateBeforeLoad = async function(){
    // if(_isMultiContracotr){
    //     _isAllowed = await IsUserInGroup(contractorGroupName);
    //     if(!_isSiteAdmin && !_isAllowed){
    //       alert('You are not allowed to submit LOD.');
    //       fd.close();
    //     }
    // }
  
    if(_isDesign){
       
         let Reference = (fd.field('Reference') !== null && fd.field('Reference') !== undefined) ? fd.field('Reference').value : '';
         let Trade = (fd.field('Trade') !== null && fd.field('Trade') !== undefined) ? fd.field('Trade').value : '';
         let CDSTitle = (fd.field('CDSTitle') !== null && fd.field('CDSTitle') !== undefined) ? fd.field('CDSTitle').value : '';

         var mesg = '';
         if(Reference === '') 
           mesg = "Reference cant be empty Can't be Empty";
         else if(Trade === '') 
           mesg = "Trade cant be empty Can't be Empty";
         else if(CDSTitle === '') 
           mesg = "CDSTitle cant be empty Can't be Empty";

           if(mesg !== ''){
             alert(mesg);
             fd.close();
           }

           debugger;
           let filter = `Reference eq '${Reference}' and Trade eq '${Trade}'`
           let items = await _web.lists.getByTitle(_DesignTasks).items.select("WorkflowStatus").filter(filter).get();
           if(items.length > 0){
                let item = items[0];
                let workflowStatus = (item['WorkflowStatus'] !== null && item['WorkflowStatus'] !== undefined) ? item['WorkflowStatus'].toLowerCase() : '';
                    
                if (workflowStatus !== "pending")
                {
                    const mesg = `${Reference} for ${Trade} is already ${workflowStatus}. PM Rejection is required so you can proceed`
                    await setPSErrorMesg(mesg, false)

                    $('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().attr("disabled", "disabled");
                    disableTable = true
                   
                    //alert(mesg);
                    //fd.close();
                }
           }
    }
}

var checkMIDPIfExist = async function(deliverableType, filename, subject){
    let mesg = '';
    let typeMetaInfo = _fncSchemas[deliverableType];
    let containsRev = (typeMetaInfo.containsRev !== null && typeMetaInfo.containsRev !== undefined) ? typeMetaInfo.containsRev : false;
    let delimeter = typeMetaInfo.delimeter;
    if(containsRev)
     filename = filename.substring(0, filename.lastIndexOf(delimeter));

    let filter = `Title eq '${filename}' and Status eq 'Approved'`;
    let items = await _web.lists.getByTitle(_MIDP).items.filter(filter).get();
    if(items.length > 0){
      if(_isMIDPSubjectExist){
        var item = items[0];
        let midpSubject = (item.Subject !== null && item.Subject !== undefined) ? item.Subject : '';
        if(midpSubject.toLowerCase().trim() !== subject.toLowerCase().trim())
          mesg = 'The reference is registered in MIDP; however, the subject does not match';
      }
    }
    else mesg ='The specified reference is not registered in the MIDP';
    return mesg;
}

var checkTradeChecker = async function(deliverableType, filename, subject){
    let mesg = '';
    let tradeValue = '';
    let typeMetaInfo = _fncSchemas[deliverableType];    
    let delimeter = typeMetaInfo.delimeter; 
    let filenameText = typeMetaInfo.filenameText;     

    let partsArray = filenameText.split(delimeter);
    let FilenameArray = filename.split(delimeter);

    let tradeIndex = partsArray.map((part, index) => {
        return part === "Trade" ? index : -1;
    }).find(index => index !== -1);    
    
    if (tradeIndex !== -1 && tradeIndex !== undefined) {
        let TradeList = typeMetaInfo.schemaFields[tradeIndex+1]?.isList?.split('|')[0] || '';
        if(TradeList !== '') {
            tradeValue = FilenameArray[tradeIndex];          
            let isExist = await isTradeMatched(TradeList, tradeValue, 'darTrade');
            
            if(!isExist)
                mesg =`The specified Trade '${tradeValue}' does not match the folder's trade '${_darTrade}'`;
        }
    } 

    return mesg;
}

var checkOriginatorChecker = async function(deliverableType, filename, subject){
    let mesg = '';
    let tradeValue = '';
    let typeMetaInfo = _fncSchemas[deliverableType];    
    let delimeter = typeMetaInfo.delimeter; 
    let filenameText = typeMetaInfo.filenameText;     

    let partsArray = filenameText.split(delimeter);
    let FilenameArray = filename.split(delimeter);

    let OriginatorIndex = partsArray.map((part, index) => {
        return part === "Originator" ? index : -1;
    }).find(index => index !== -1); 
    
    let DrawingTypeIndex = partsArray.map((part, index) => {
        return (part === "DrawingType" || part === "DocumentType") ? index : -1;
    }).find(index => index !== -1);     
    
    if (OriginatorIndex !== -1 && OriginatorIndex !== undefined) {
        let TradeList = typeMetaInfo.schemaFields[OriginatorIndex+1]?.isList?.split('|')[0] || '';
        if(TradeList !== '') {

            tradeValue = FilenameArray[OriginatorIndex];
            
            if (DrawingTypeIndex !== -1 && DrawingTypeIndex !== undefined) {     
                DrawingTypeValue = FilenameArray[DrawingTypeIndex]; 
                if(DrawingTypeValue.toLowerCase() !== 'm3d')
                    if(tradeValue.toLowerCase() !== 'dar')
                        mesg =`The specified DrawingType '${DrawingTypeValue}' does not match the Originator '${tradeValue}'`; 
                else
                    return mesg;         
            }
            
            if(mesg === ''){
                let isExist = await isTradeMatched(TradeList, tradeValue, 'darOriginator');
                
                if(!isExist)
                    mesg =`The specified Originator '${tradeValue}' does not match the folder's Originator '${_darTrade}'`;
            }
        }
    } 

    return mesg;
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

        $("<img id='loader' src='" + ImageUrl + "' />")
            .css({
                "position": "absolute",
                "top": "300px",
                "left": "790px",
                "width": "100px",
                "height": "100px"
            }).insertAfter(targetControl);
	}
   catch(err) { console.log(err.message); }
}

function remove_preloader(){
    var targetControl = $('#ms-notdlgautosize')
    targetControl.undim();

    // if($('div.dimbackground-curtain').length > 0)
    //   $('div.dimbackground-curtain').remove();

     if($('#loader').length > 0)
         $('#loader').remove();
}

var get_Counter = async function(_web, key){
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

var setRowsReadOnly = async(rowIndex, setReadOnly) => {
    var metaObject = {};
    if(rowIndex !== null && rowIndex !== undefined){
        if(setReadOnly === true){
            metaObject['readOnly'] = true;
            metaObject['className'] = closedClassName;
        }
        else{
            metaObject['readOnly'] = false;
            metaObject['className'] = defaultClassName;
        }
        
        _hot.getSettings().columns.forEach((column, columnIndex) => {
            _hot.setCellMetaObject(rowIndex, columnIndex, metaObject);
        });
      }

      else{
        const rowCount = _hot.countRows();
        for (var row = 0; row < rowCount; row++){
            const fnColumnIndex = _hot.propToCol('FullFileName');
            var filename = _hot.getDataAtCell(row, fnColumnIndex);

            let statusIndex = _hot.propToCol('Status');
            let status = _hot.getDataAtCell(row, statusIndex);

            if(filename !== undefined && filename !== null && filename !== ''){
            _hot.getSettings().columns.forEach((item, columnIndex) => {
                    if(_updateTitle && item.data === 'Title'){
                        if(!_isDesign){
                            if(status !== null && (status.toLowerCase() === 'pending' || status.toLowerCase() === 'rejected'))
                              metaObject['readOnly'] = false;
                            else metaObject['readOnly'] = true;
                        }
                        else metaObject['readOnly'] = false;
                    }
                    else metaObject['readOnly'] = true;

                    if(_isDesign && status !== null && (status.toLowerCase() === 'pending' || status.toLowerCase() === 'rejected'))
                        metaObject['className'] = cancelClassName;
                    //else metaObject['className'] = closedClassName;
                    _hot.setCellMetaObject(row, columnIndex, metaObject);
              });
            }
        }
      }
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

function setToolTipMessages(){
    //setButtonToolTip('Save', saveMesg);
    setButtonCustomToolTip('Submit', submitMesg);
    setButtonCustomToolTip('Cancel', cancelMesg);
  
    //else addLegend();
    //adjustlblText('Comment', ' (Optional)', false);
        
      if($('p').find('small').length > 0)
      $('p').find('small').remove();
}
//#endregion

//#region FNC CHECKER
var setNamingConvention = async function(rowIndex, columnIndex, filename, deliverableType){
    var typeMetaInfo = _fncSchemas[deliverableType];
    var mesg = '';

    if(typeMetaInfo === undefined)
        mesg = `${deliverableType} is not defined in FNC`;
    else{
        var delimeter = typeMetaInfo.delimeter;
        var schemaFields = typeMetaInfo.schemaFields;
        var filenameText = typeMetaInfo.filenameText;
        var checkRev = typeMetaInfo.containsRev;

        const illegalChars = /['{}[\]\\;':,\/?!@#$%&*()+=]/;

        if (!filename.includes(delimeter))
            mesg = 'Unstructured Filename';
        else if(illegalChars.test(filename))
            mesg = 'Filename Contains illegal character';
        else mesg = await checkFileName(deliverableType, delimeter, schemaFields, filenameText, filename, true, checkRev);
    }
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

var setErrorMesg = async function(rowIndex, mesg, removedIndexes){
    const mesgColumnIndex = _hot.propToCol('Mesg');
    const statusColumnIndex = _hot.propToCol(_defaultStatusImgCol);

    var iconUrl = `<img src='${_submitImg}' alt='na'></img>` ;

    if(mesg === 'empty'){
        mesg = '';
        iconUrl = '';
    }
    else{
        if(mesg !== '')
            iconUrl = "<img src='" + _errorImg + "' alt='na'></img>";
    }
    

    let isRemove = false;
    let rowData;
    //if(mesg === ''){
        rowData = _hot.getDataAtRow(rowIndex).reverse();
        let containsInfo = false;
        for (let i = 2; i < rowData.length; i++) { //-2 to execlude last 2 columns (Mesg & Status)
            if(rowData[i]  !== undefined && rowData[i] !== null && rowData[i]  !== ''){
                containsInfo = true;
                break;
            }
        }
        if(!containsInfo) 
          isRemove = true;
        //console.log('rowData = ' + rowData);
    //}

    if(isRemove){
        mesg = '';
        iconUrl = ''

        let valueToAdd = parseInt(rowIndex);
        if (!removedIndexes.includes(valueToAdd)) {
            removedIndexes.push(valueToAdd);
        }
        removeError(valueToAdd);
    }
    else{
         var index = parseInt(rowIndex);
        _hot.setDataAtCell(index, mesgColumnIndex, mesg);
        _hot.setDataAtCell(index, statusColumnIndex, iconUrl);   
    }
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
             if(_fld !== 'Mesg' && _fld !== 'TempStatus')
                await createFields(_web, currentList, _fld, colsType[i]);
            }
        }
      }
}

var isItemFound = async function(selectedRow, operation, item){
    const fnColumnIndex = _hot.propToCol('FullFileName');
    var filename = '';
    if(item !== null && item !== undefined)
      filename = selectedRow;
    else filename = _hot.getDataAtCell(selectedRow, fnColumnIndex);

    _web.lists.getByTitle(_RLOD).items.select("Id").filter(`FullFileName eq '${filename}'`).top(1).get()
    .then(items => {
        if (items.length > 0){
            // Delete the item
            const itemId = items[0].Id;
            if(operation === 'Remove'){
            _web.lists.getByTitle(_RLOD).items.getById(itemId).delete()
                .then(() => {
                console.log("Item deleted successfully");
                })
                .catch(error => {
                console.error("Error deleting item:", error);
                });
            }
            else if(operation === 'Cancelled'){
                _web.lists.getByTitle(_RLOD).items.getById(itemId).update({Status: operation})
                .then(() => {
                    console.log("Item updated successfully.");
                    _hot.alter("remove_row", selectedRow);
                })
                .catch((error) => {
                    console.error("Error updating item:", error);
                });
            }
            else if(item !== null && item !== undefined){
                _web.lists.getByTitle(_RLOD).items.getById(itemId).update({item})
                .then(() => {
                    console.log("Extract item match updated successfully.");
                })
                .catch((error) => {
                    console.error("Error updating extract item match:", error);
                });
            }
        }
        else console.log("Item not found");
    })
    .catch(error => {
        console.error("Error retrieving item:", error);
    });
}

var isTradeMatched = async function(TradeList, tradeValue, columnName){   

    let isExist = false;      
    const camlFilter = `<View>                            
                            <Query>
                                <Where>	
                                    <And>								
                                        <Eq><FieldRef Name='Title'/><Value Type='Text'>${tradeValue}</Value></Eq>	
                                        <Eq><FieldRef Name='${columnName}'/><Value Type='MultiChoice'>${_darTrade}</Value></Eq>	
                                    </And>					
                                </Where>
                            </Query>
                        </View>`;

    const existingItems = await pnp.sp.web.lists.getByTitle(TradeList).getItemsByCAMLQuery({ ViewXml: camlFilter });
    if(existingItems.length > 0) 
        isExist = true;
    
    return isExist;
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

        if( internalName === _defaultStatusImgCol || internalName === 'Mesg') continue;

         if(_value !== null && _value !== ''){
        
          _columns += internalName + ',';
          _colType += mentaInfo[j].type + ',';

          if(internalName === 'FullFileName')
          _value = _value.toUpperCase();

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
    //Insert to RLOD
    const list = _web.lists.getByTitle(_RLOD);
    var batch = pnp.sp.createBatch(); // eg: 10 min
    
    let FolderName = '';//, listUri = '/projects/PILOT-S/Lists/RLOD', folderUri  = '/SBG';
    const listUri = `${fd.webUrl}/Lists/${_RLOD}`;
    let folderUri;
    if(_isMultiContracotr){
        var queryString = window.location.search;
        var urlParams = new URLSearchParams(queryString);
        FolderName = urlParams.get('rf');

        if(_isEdit){

            if(FolderName !== undefined && FolderName !== null && FolderName !== '' && FolderName.includes('/')){
                //IF NOT SITE ADMIN IT COMES HERE
                FolderName = FolderName.substring(FolderName.lastIndexOf('/')+1);
                folderUri = `/${FolderName}`;
            }
            else{
                    let folderUrl = urlParams.get('source');
                    urlParams = new URL(folderUrl);
                    const searchParams = urlParams.searchParams;
                    FolderName = searchParams.get('RootFolder');
                }
        }
        
        if(FolderName !== null && FolderName.includes('/')){
          FolderName = FolderName.split("/")[5];
          folderUri = `/${FolderName}`;
        }
    }
    
    var row = 0, rowBatchSize = 1;  
    var _columns, _colType;
    for (const item of itemsToInsert){
     
      _columns = objColumns[row];
      var deliverableType = item.DeliverableType;
      var schemDelivType, delimeter, schema, category;
      var _query, filename = item.FullFileName;
      
      if(!_ignoreChecking){

       schemDelivType = _fncSchemas[deliverableType];
       delimeter = schemDelivType.delimeter;
       schema = schemDelivType.schemaFields;
       category = schemDelivType.category;

        if(schemDelivType.containsRev && !_ignoreChecking)
            filename = filename.substring(0, filename.lastIndexOf('-'));
        _query = "FileName eq '" + filename + "'";
      }
      else _query = "FileName eq '" + filename + "'";
  
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

        if(_CheckRevision.toLocaleLowerCase() === 'no')
          update_MetaInfo_For_CheckRevision(_item, filename, delimeter, schemDelivType.containsRev);
        
        if(doUpdate){

            if(_isDesign && _exact_Filename_Match === 'yes'){

                item['DarTrade'] = _darTrade;
                item['CDSTitle'] = _cdsTitle;

                item['Category'] = category;
                item['FileName'] = filename;
                item[_filterField] = _lodRef;
                item['Status'] = 'Pending';

                if(!_ignoreChecking)
                    await splitFileName(schema, delimeter, item)
            }

            list.items.getById(_item.Id).inBatch(batch).update(item);
        }
      } 
      else {
        item['Category'] = category;
        item['FileName'] = filename;
        item[_filterField] = _lodRef;
        
        if(_isDesign){
			item['DarTrade'] = _darTrade;
			item['CDSTitle'] = _cdsTitle;
        }

        if(!_ignoreChecking)
         await splitFileName(schema, delimeter, item)

        list.items.inBatch(batch).add(item)
        .then((item) => {
            if(_isMultiContracotr){
                return _web
                    .getFileByServerRelativeUrl(`${listUri}/${item.data.Id}_.000`)
                    .moveTo(`${listUri}${folderUri}/${item.data.Id}_.000`);
            }
        });
      }
      

      if(rowBatchSize === updateBatchSize){
        row++;
        detailedLoader(itemsToInsert.length, row);
        rowBatchSize = 1;
        await batch.execute();
        batch = pnp.sp.createBatch();
        continue;
      }
      rowBatchSize++;
      row++;
    }

    for(const item of _updateItems){
        let Id = item.id;
        let metaInfo = item.metaInfo
        list.items.getById(Id).inBatch(batch).update(metaInfo);
         
        if(rowBatchSize === updateBatchSize){
            row++;
            detailedLoader(itemsToInsert.length, row);
            rowBatchSize = 1;

            await batch.execute();
            batch = pnp.sp.createBatch();
            continue;
        }
        rowBatchSize++;
        row++;
    }
    
    if( batch._index > -1){  
      let overallTotal = (itemsToInsert.length + _updateItems.length)
      detailedLoader(overallTotal, row);
      await batch.execute();
    }

    // await batch.execute()
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

var update_MetaInfo_For_CheckRevision = async function(item, filename, delimeter, containsRev){
    item["FullFileName"] = filename;

    if(containsRev)
      item["Revision"]  = filename.substring(filename.lastIndexOf(delimeter)+1);

    item["SubmittalRef"] = _lodRef;
    item["Status"] = "Pending";

    item["CDSLink"] = null;
    item["CDSTitle"] = "";
    item["CRSLink"] = null;
    item["CRSNumber"] = "";
    item["DWFLink"] = null;
    item["DWFSize"] = "";
    item["DWGLink"] = null;
    item["DWGSize"] = "";
    item["MarkupLink"] = null;
    item["PDFLink"] = null;
    item["PDFSize"] = "";
    item["Code"] = "";
    item["Reference"] = "";
    item["PendingFileName"] = "";
    item["PendingCDS"] = "";
    item["ReviewedLink"] = null;
    item["ReviewedDate"] = null;
    item["SentToPMCDate"] = null;
    item["SentToPMDate"] = null;
    item["SentToSiteDate"] = null;
}
//#endregion

function isAllDigits(str) {
    return /^\d+$/.test(str);
}

var validateExactFileNameMatch = async function(deliverableType, title,  filename, roldItem){
    debugger;
    let typeMetaInfo = _fncSchemas[deliverableType];
    let containsRev = (typeMetaInfo.containsRev !== null && typeMetaInfo.containsRev !== undefined) ? typeMetaInfo.containsRev : false;
    let revStartWith = (typeMetaInfo.revStartWith !== null && typeMetaInfo.revStartWith !== undefined) ? typeMetaInfo.revStartWith : '';
    let delimeter = typeMetaInfo.delimeter;
    
    let rlodLODRef = (roldItem.SubmittalRef !== null && roldItem.SubmittalRef !== undefined) ? roldItem.SubmittalRef : ''; 
    let rlodFullFilename = (roldItem.FullFileName !== null && roldItem.FullFileName !== undefined) ? roldItem.FullFileName : '';
   
    let filenameOriginalRev = '', fncRev, rlodRev;
    if(containsRev){
      fncRev = filename.substring(filename.lastIndexOf(delimeter)+1);
      filenameOriginalRev = fncRev;
      rlodRev = (roldItem.Revision !== null && roldItem.Revision !== undefined) ? roldItem.Revision : null;
   
      if(revStartWith !== ''){
       fncRev = fncRev.replace(revStartWith, '');
       rlodRev = rlodRev.replace(revStartWith, '');
      }
    }
   
    //if(_exact_Filename_Match === 'yes'){
      let checkLogic = false;
      if(fncRev != _revNumStart && fncRev != "A" && fncRev != "AA" && fncRev != 0)
        checkLogic = true;
      else if (fncRev != rlodRev)
        checkLogic = true;
     
    if (_lodRef !== rlodLODRef && rlodFullFilename === filename && containsRev)
      return `Filename was previously submitted under ${rlodLODRef}`;
    
     if(checkLogic && containsRev){
        let rlodStatus = (roldItem.Status !== null && roldItem.Status !== undefined) ? roldItem.Status : '';
        if(_closingStatus === rlodStatus){

            let result = fncRev.localeCompare(rlodRev, 'en', { sensitivity: 'base' });
            if (result > 0 || result < 0){
                if (isAllDigits(rlodRev) && isAllDigits(fncRev)) {
                  let targetRev = parseInt(rlodRev) + 1;
                  targetRev = targetRev.toString().padStart(rlodRev.length, '0');
                  if(targetRev !== fncRev) 
                    return `Filename should be submitted as Revision ${targetRev} instead of ${fncRev} as RLOD latest submitted revision is ${rlodRev}`;
                }
                else if(!isAllDigits(rlodRev) && !isAllDigits(fncRev)){
                    //let conRev = parseInt(rlodRev.charAt(0)); // Convert the first character of RLODRev to an integer
                    let rlodRevTemp = String.fromCharCode(rlodRev.charCodeAt(0) + 1) // Convert the incremented character code to a string
                    if(rlodRevTemp !== fncRev)
                      return `Filename should be submitted as Revision ${rlodRevTemp} instead of Revision ${fncRev} as RLOD Revision is ${rlodRev}`;
                }
            }
        }
        else if(rlodStatus === 'Cancelled')
          return "Contact PWS admin to un-cancel to proceed";
        else return "Previous Revision is Already Under Review";
      }

      //updateItems
      var _objValue = { };
      //_objValue['ID'] = ;
      _objValue['FullFileName'] = filename;
      _objValue['Revision'] = filenameOriginalRev;
      _objValue['SubmittalRef'] = _lodRef;
      _objValue['Status'] = 'Open';
      _objValue['Title'] = title;
      _updateItems.push({
              id: roldItem.Id, 
              metaInfo: _objValue
      });
        // isItemFound(filename, '', {
        //     FullFileName: filename,
        //     Revision: fncRev,
        //     SubmittalRef: _lodRef,
        //     Status: 'Open'
        // });
    //}
} 

var addLegends = async function(){
    let legend = document.createElement('legend');
    legend.innerHTML = "<div style='display:flex'>" +
                        "<div style='width: 17px;height: 15px;background-color:#efa2137a; border:1px solid; font-size: 10px; text-align: center;margin-top: 2px;'></div>" +
                        "<div style='font-size:14px;padding-left: 5px;'>Indicates that the record is allowed for cancellation</div>" +
                       "</div>";
                       
    let targetDiv = document.getElementById('dt');
    targetDiv.parentNode.insertBefore(legend, targetDiv);
}

var  detailedLoader = async function(total, index, isDim){
		var targetControl = $('#ms-notdlgautosize').addClass('remove-position-preloadr');
        
        if(isDim === true){
            targetControl.dimBackground({
                darkness: 0.4
            }, function () {}); 
        }
        else{
            let mesg = `Processing ${index} out of ${total}. Please dont close this page.`;
            if($('#totalId').length > 0)
              $('#totalId').text(mesg);
            else{
                $(`<span id='totalId'>${mesg}</span>`)
                .css({
                    "position": "fixed",
                    "top": "50%",
                    "left": "50%",
                    "transform": "translate(-50%, -50%)",
                    "background-color": "rgb(115 115 115 / 55%)",
                    "border-radius": "15px",
                    "color": "white",
                    "box-shadow": "2px 10px 10px rgba(0, 0, 0, 0.1)",
                    "transition": "opacity 0.3s ease-out",
                    "font-weight": "bold",
                    "padding": "10px",
                    "font-size": "20px",
                    "width": "auto"
                }).insertAfter(targetControl);
            }         
        }          
}