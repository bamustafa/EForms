var hot;
var container;
var batchSize = 20;
var data = [];

var errorClassName = 'ErrorRow';
var defaultClassName = 'TransparentRow';
var closedClassName = 'ClosedRow';
var cancelClassName = 'CancelRow';

var editedCells = null, columnProps = null;

var haveError = false, _isNew = false, isFilterSelected = false, _doHide = false, isResubmit = false;
var webURL = '', _targetListName = '', _filterfld;

var Reference = '', Trade = '', RejectionTrades = '';

var _colsArray, targetList, targetFilter;
var contextMenu = [];
var _hideColumn = '';

var buildGrid = async function() {
    //#region EXTRACT VALUES
    Reference = fd.field("Reference").value;
    
    if(_isPart){
      Trade = fd.field("Trade").value;
      isResubmit = fd.field("Resubmit").value;
      RejectionTrades = fd.field('RejectionTrades').value;
    }

    data = [];
    
    webURL = document.URL.substring(0, document.URL.indexOf('/PlumsailForms')).replace('SitePages','');
    var _web = pnp.sp.web;

    var _objArray = await getGridMajorType(_web, webURL, _module);

     _colsArray = _objArray.colArray;
     targetList = _objArray.targetList;
     targetFilter = _objArray.targetFilter;

    if(_formType === 'New')
        _isNew = true;
    else _isNew = false;
    //#endregion

    //#region GET GRID DATAT
    if(!_isNew){

      var _colsInternal = [], _colsType = [], _colsht = [];

      if(isContractor)
        _colsArray = _colsArray.filter(item => item.data === 'attachFiles' || item.readOnly !== true);
      else{
        if(_isTeamLeader)
          _colsArray = _colsArray.filter(item => (item.hide !== true || item.data === 'trade') && (item.owner == 'Dar' || item.owner == 'TL'));
        else _colsArray = _colsArray.filter(item => {
          if(isResubmit)
           return (item.isResubmit === true || (item.hide !== true && item.owner == 'Dar'));
           else return (item.hide !== true && item.owner == 'Dar')
        });
      }

      _colsArray.map(item =>{
          _colsInternal.push(item.InternalName);
          _colsType.push(item.type);
          _colsht.push(item.data);
      });

      data = await getData(_web, webURL, targetList, targetFilter, Reference, _colsInternal, _colsType, _colsht);
      if(data.length === 0){
        //hide Attach files column
        _hideColumn = 'attachFiles';
        _doHide = true;
      }

      if(!isContractor && !_isTeamLeader){
        if(data.length < batchSize){
          var remainingLength = batchSize - data.length;

          for (var i = 0; i < remainingLength; i++) {
            var rowData = { id: i + 1, value: 'Row ' + (i + 1) }
            data.push(rowData);
          }
        }
      }
    }
    else {
      for (var i = 0; i < batchSize; i++) {
        var rowData = { id: i + 1, value: 'Row ' + (i + 1) }
        data.push(rowData);
       }
    }
    //#endregion

    //#region CONTEXTMENU
    if(Status === 'Completed' || _isTeamLeader || _isMain){
      contextMenu = [];
    }
    else{
      if(_isNew){
        contextMenu = ['row_below', '---------', 'remove_row'];
      }
      else{
            if(isContractor)
            contextMenu = {}
            else{
              contextMenu = {
                callback(key, selection, clickEvent) {
                  // Common callback for all options
                  console.log(key, selection, clickEvent);
                },
                items: {
                  row_below: {
                    name: 'Click to add row below' // Set custom text for predefined option
                  },
                  remove_row: {
                    name: "Remove Row", // The text for the menu item
                    callback: async function (key, selection, clickEvent) {
                      // Get the selected row index
                      debugger;
                      var selectedRow = selection[0].start.row;
                      var isFound = await isItemFound(selectedRow);

                      if(isFound){
                        alert(htPreventDeleteMesg);
                        return;
                      }
        
                      // Remove the row from the data
                      hot.alter("remove_row", selectedRow);
        
                      // Render the changes
                      hot.render();
                    },
                  },
                  // cancel: {
                  //   name: 'Cancel',
                  //   callback(key, selection, clickEvent) {
                  //     var startInex = selection[0].start.row;
                  //     var endInex = selection[0].end.row;
                  //     var setIndexes = [];
                  //     setIndexes.push[startInex];

                  //     var rowData = hot.getCellMeta(startInex, 1);
                  //     setRowsReadOnly(startInex,true);
                  //   }
                  // },
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
      }
    }
    //#endregion

    if(!_isLead && !_isPart)
    {
      if(!isContractor || (isContractor && Status !== 'Issued to Contractor'))
        return;
    }

    if(!_isLead && !_isPart && isContractor){
      var id = '#notelbl';
      if ($(id).length === 0) {
        var label = $('<label>', {
          id: 'notelbl',
          text: contractorLabelMesg
        });
        label.css('color', 'red');
        label.css('font-weight', 'bold');
        label.css('width', '950px'); 
        $('#dt').before(label);
      }
    }

    container = document.getElementById("dt");
   
    _spComponentLoader.loadScript(_htLibraryUrl).then(setGridHandsonTable);

    return {
        element: hot,
        targetList: targetList,
        targetFilter: targetFilter
    };
}

const setGridHandsonTable = async (Handsontable) => {
    Handsontable.validators.registerValidator('getCellValidator', getCellValidator);

    var _height = '400';
    if(_isMain)
    _height = '550';
    hot = new Handsontable(container, {
        data,
        height: _height,
        width:'100%',
        rowHeights: 25,
        columns: _colsArray,
        dropdownMenu: false,
        contextMenu: contextMenu,
        //autoRowSize: true,

        // fixedColumnsLeft: 4,
        filters: true,
        filter_action_bar: true,
        rowHeaders: true,
        dropdownMenu: ['filter_by_condition', 'filter_by_value', 'filter_action_bar'],
        // manualColumnFreeze: true,
        manualColumnResize: true,
        stretchH: 'all',
        licenseKey: htLicenseKey
    });

    hideColumns(_hideColumn, _doHide);
    columnProps = await getTargetandColumnIndex();

    if(!_isNew)
      setRowsReadOnly(null, true, isContractor);

     hooks();

     addLegend('SLlbl', SLFLabelMesg, container);

     hot.render();
}

var getTargetandColumnIndex = async function(){

  cols = hot.getSettings().columns;
  var colsMandatoryForClosingIndexes = [];
  var colsMandatoryForClosingTitles = [];

  var targetColumnIndex, targetColumnTitle, closingMandatoryColumnIndex, closingMandatoryColumnTitle;
  
  for(var i = 0;i<cols.length;i++){
      if(cols[i]["TriggerValidation"] === true){
        targetColumnIndex = i;
        targetColumnTitle = cols[i]['data'];
      }
      if (cols[i].closingMandatoryColumn === true){
        closingMandatoryColumnIndex = i;
        closingMandatoryColumnTitle = cols[i]['data'];
        closedStatus = cols[i]['closedStatus'];
      }
      if (cols[i].isMandatoryForClosing === true){
       colsMandatoryForClosingIndexes.push(i);
       colsMandatoryForClosingTitles.push(cols[i]['data']);
      }
  }
  
  return {
      targetColumnIndex: targetColumnIndex,
      targetColumnTitle: targetColumnTitle,
      
      closingMandatoryColumnIndex: closingMandatoryColumnIndex,
      closingMandatoryColumnTitle: closingMandatoryColumnTitle,
      closedStatus: closedStatus,

      colsMandatoryForClosingIndexes: colsMandatoryForClosingIndexes,
      colsMandatoryForClosingTitles: colsMandatoryForClosingTitles
  } 
}

var hideColumns = async function(_column, doHide){
    var hiddenColumns = [];
    var colIndex = 0;
    if(!doHide){
      hiddenColumns = [];
    }
    else{
      hot.getSettings().columns.find(item => {
        if(_column !== ''){
          if(item.data === _column){
            hiddenColumns.push(colIndex);
          }
        }
        else{
        if(_isNew && (item.owner.toLowerCase() !== 'dar' || item.readOnly === true))
          hiddenColumns.push(colIndex);
        }
        colIndex++;
      });
    }

    hot.updateSettings({
      hiddenColumns: {
        columns: hiddenColumns,
        indicators: true // Show the hidden columns indicators
      }
    });
}

var hooks = async function(){
  var isLoadingData = false;
  function addRowsInBatch(startIndex, endIndex) {
      var newData = []; 

      for (var i = startIndex; i < endIndex; i++) {
          var rowData = { id: i + 1, value: 'Row ' + (i + 1) }; 
          newData.push(rowData);
      }

       data = data.concat(newData); // Concatenate the new data with the existing data array

      // Update the settings to reflect the changes
      hot.updateSettings({
          data: data
      });
      isLoadingData = false; // Reset the flag variable
  }

  function handleScroll() {
     
      var scrollElement = container.querySelector('.wtHolder');
      var scrollPosition = scrollElement.scrollTop;
      //console.log('scrollPosition = ' + scrollPosition);

      var visibleHeight = container.offsetHeight;
      var totalHeight = scrollElement.scrollHeight - container.scrollHeight; // - 50;

      // Calculate the scroll position at the bottom of the table
      var bottomScrollPosition = totalHeight - visibleHeight;
    
      // Check if the user has scrolled to the bottom of the table
      if (!isLoadingData && scrollPosition >= totalHeight) {
          isLoadingData = true;

        var startIndex = hot.countRows();
        var endIndex = startIndex + batchSize;
        if(!isContractor && !_isTeamLeader) 
         addRowsInBatch(startIndex, endIndex);
      }
  }

  hot.addHook('beforeChange',  function(changes, source) {
    if (source === 'edit' || source === 'CopyPaste.paste') {
      editedCells = changes;
      if(editedCells.length > 200){
        preloader();
        preloader("remove");
      }
    }
  });

  //afterGetColHeader is not used but i leave it as a sample code for later use if needed
  hot.addHook('afterGetColHeader_notUsed', (col, TH) => {
    // if (TH.textContent === 'Item No.' || TH.textContent === 'Incremental') {
    //   const button = TH.querySelector('.changeType');
    //   if (!button) {
    //       return;
    //   }
    //   button.parentElement.removeChild(button);
    // }
  });

  hot.addHook('afterScrollVertically', handleScroll);container

  hot.addHook('beforeValidate', (value, row, prop, source) => {

      let  targetColumnIndex = -1;
      
      targetRowIndex = row;
      hot.getSettings().columns.find(item => {
          if(targetColumnIndex === -1)
            targetColumnIndex = 0;
          else targetColumnIndex++;

          if(prop.toLowerCase() === item.data.toLowerCase()){
            return true;
          }
      });
      
      return {
          rowIndex: row,
          columnIndex: targetColumnIndex,
          columnTitle: prop,
          value: value
      }
  });

  hot.addHook('afterRemoveRow', function(index, amount) {
    var Indexes = [];
    for(var i =1; i <= amount; i++){
      index++
      Indexes.push(index);
    }
    setErrortoMasterError(Indexes,null,true);
  });

  hot.addHook('afterContextMenuDefaultOptions', (options) => {
    var selectedCell = hot.getSelectedLast();
    var rowIndex = selectedCell[0];
    var columnIndex = columnProps.closingMandatoryColumnIndex;
    var cellValue = hot.getDataAtCell(rowIndex, columnIndex);
  
    if(!_isNew){
  
      var itemIndex = -1;
      options.items.find(item => {
        itemIndex++;
        if(item.key === 'remove_row'){
          return true
        }
      });
  
      // if (cellValue !== 'Closed'){
  
      //   var isFound = await isItemFound(rowIndex);
      //   if(isFound){
      //     options.items[itemIndex].disabled = true;
      //     return options;
      //   }
      // }
        
      if (cellValue === 'Closed') 
      {
          options.items[itemIndex].disabled = true;
  
          // if(cellValue === 'Cancelled'){
          //   options.items[cancelIndex].disabled = true;
          //   options.items[unCancelIndex].disabled = false;
          // }
          // else if(cellValue === 'Open'){
          //   options.items[cancelIndex].disabled = false;
          //   options.items[unCancelIndex].disabled = true;
          // }
          // else{
          //   options.items[cancelIndex].disabled = true;
          //   options.items[unCancelIndex].disabled = true;
          // }
          return options;
        }
      }
  });

  hot.addHook('afterFilter', (arguments) => {
    if(arguments.length > 0)
      isFilterSelected = true;
     else isFilterSelected = false;
  });
}

var getData = async function(_web, webURL, _listname, _filterfld, _refNo, _colsInternal, _colsType, _colsht){ 
    var _itemArray = [];
    const _cols = _colsInternal.join(',');
    var _query = _filterfld + " eq '" + _refNo + "'";

    if(_isPart)
    {
       var trade = fd.field('Trade').value;
        if(!_isTeamLeader){
          if (trade) {
            _query += " and Trade eq '" + trade + "'";
          } else {
            _query += " and Trade eq null";
          }
        }
   }

    var fileUrl = webURL + _layout + '/Images/Icons/'
    var fileIcons = {
      pdf: fileUrl + 'Pdf.png',
      doc: fileUrl + 'Word.png',
      docx: fileUrl + 'Word.png',
      xls: fileUrl + 'Excel.gif',
      xlsx: fileUrl + 'Excel.gif',
      ppt: fileUrl + 'Ppt.png',
      pptx: fileUrl + 'Ppt.png',
      png: fileUrl + 'Png.png',
      gif: fileUrl + 'Png.png',
      jpeg: fileUrl + 'Jpeg.gif',
      jpg: fileUrl + 'Jpeg.gif',
      rar: fileUrl + 'Zip.gif',
      zip: fileUrl + 'Zip.gif',
      attach: fileUrl + 'attach.gif'
      // Add more file extensions and corresponding icons here
    };

    await _web.lists.getByTitle(_listname).items.filter(_query)
    .select('Id,' + _cols)
    .expand('AttachmentFiles') // Include attachments in the query
    //.expand(_expandColumns)
    .getAll().then(async function(items) {
      _itemCount = items.length;
      if (_itemCount > 0) {
        for(var i = 0; i < _itemCount; i++){
           var item = items[i];
           var rowData  = {};

           var _colStatus = '';
           for(var j = 0; j < _colsInternal.length; j++){
              var _type = _colsType[j];
              var _colname = _colsInternal[j];
              var htColumn = _colsht[j];
              var _value = item[_colname];

              if(_colname === 'Status')
                _colStatus = _value;

              if(_colname === 'Attachments'){
          
                _value = '';
                var attachments = item.AttachmentFiles;
                //var tblHeight = '20%';
                
                if (attachments.length > 0) {

                  //  if(attachments.length === 1)
                  //     tblHeight = '61%';
                  
                  _value = "<table align='center'>";
                  for (var k = 0; k < attachments.length; k++) {
                    var attachment = attachments[k];
                    var attachmentName = attachment.FileName;
                    var fileExtension = attachmentName.split('.').pop().toLowerCase();
                    var iconUrl = fileIcons[fileExtension] || fileUrl + 'icon-default.png';

                    var attachmentUrl = attachment.ServerRelativeUrl;
                    //var thumbnailUrl = attachmentUrl + "?&thumbnailMode=1";
                    _value += "<tr>" +
                                "<td style='border: none'><img src='" + iconUrl + "' alt='na'></img></td>  " +
                                "<td style='border: none'><a target='_blank' href='" +  attachmentUrl + "'>" + attachment.FileName + "</a></td>" +
                              "</tr>";
                    //_value += "<a target='_blank' href='" +  attachmentUrl + "'><img src='" + thumbnailUrl + "' alt='Attachment " + (k + 1) + "' width='100'></img></a><br/>";
                  }
                  _value += "</table>";
                }

                if(_isPart && Status === 'Completed'){}
                else if(_isPart && _colStatus !== 'Closed' && !_isTeamLeader){
                  var itemId = item.Id;
                  var link = webURL + 'SitePages/PlumsailForms/' + _listname.replace('Snag Items','SLFItems') + '/Item/EditForm.aspx?item=' + itemId;
                  //_value = "<a target='_blank' href='" +  link + "'>upload files</a>";
                  var iconUrl = fileIcons['attach'];
                  //_value += "<table align='left' height='" + tblHeight + "'>" +
                  _value += "<table align='left'>" +
                              "<tr>" +
                                "<td style='border: none; padding-top: 5px; vertical-align: bottom'> " + 
                                    "<img src='" + iconUrl + "' alt='na'></img>  " + 
                                    "<a href='#' onclick=\"openSmallWindow('" + link + "')\">Attach files</a>" + 
                                  "</td>" +
                              "</tr>" +
                            "</table>";
                }
                rowData[htColumn] = _value;
              }

              else if( _type.toLowerCase() == "date" && _value !== undefined && _value !== '' && _value !== null){

                  // const parts = _value.split("T")[0].split("-");
                  // var day = parseInt(parts[2]);
                  // var month = parseInt(parts[1]);
                  // var year = parseInt(parts[0]);

                  // _value = day + '/' + month + '/' + year;

                  // //var date = new Date(_value); //year, month, day);
                  // rowData[htColumn] = _value;

                  var date = new Date(_value);
                  var formattedDate = date.toLocaleDateString('en-GB'); // Format the date as "DD/MM/YYYY"
                  rowData[htColumn] = formattedDate;
              }
              else rowData[htColumn] = _value;
           }
           _itemArray.push(rowData);
           
        }
        console.log(_itemArray);
      }
  });
  return _itemArray;
}

function openSmallWindow(url) {
  var width = 1100; // Specify the desired width of the small window
  var height = 600; // Specify the desired height of the small window
  var left = (window.innerWidth - width) / 2;
  var top = (window.innerHeight - height) / 2;
  var options = 'width=' + width + ',height=' + height + ',top=' + top + ',left=' + left;
  window.open(url, '_blank', options);
}

var isItemFound = async function(ItemNo){
 ItemNo++;
 ItemNo = String(ItemNo).padStart(5, '0');
  var _query = targetFilter + " eq '" + Reference + "' and ItemNo eq '" + ItemNo + "'";
  if (Trade) {
    _query += " and Trade eq '" + Trade + "'";
  } 

  try {
    const items = await pnp.sp.web.lists
      .getByTitle(targetList)
      .items
      .select("Id")
      .filter(_query)
      .top(1)
      .get();

    return items.length > 0;
  } catch (error) {
    return false;
  }
}


// {  "title": "<b>Ready for checking</b>",
// "data":"ready for checking",
//   "type": "checkbox",
//   "width": 120,
// "validator": "preventEdit",
// "InternalName": "IsChecked",
// "owner": "Contractor",
// "hide": false
// },