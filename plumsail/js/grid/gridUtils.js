var masterErrors = {};
var triggerDescription = true;

var getGridMajorType = async function(_web, webURL, key){
    //var _itemArray = {};
    var _colArray = [];
    var targetList;
    var targetFilter;

    var result = "";
    await _web.lists
          .getByTitle("MajorTypes")
          .items
          .select("HandsonTblSchema,LODListName,LODFilterColumn")
          .filter("Title eq '" + key + "'")
          .get()
          .then(async function (items) {
              if(items.length > 0)
                //for (var i = 0; i < items.length; i++) {
                  var item = items[0];
                  var HandsonTblSchema = item.HandsonTblSchema;
                  targetList = item.LODListName;
                  targetFilter = item.LODFilterColumn;

                    var fetchUrl = webURL + HandsonTblSchema;
                    await fetch(fetchUrl)
                        .then(response => response.text())
                        .then(async data => {
                          _colArray = JSON.parse(data); 

                          //Manipulate the objects as needed
                          for (const obj of _colArray) {
                            if (obj.renderer === "incrementRenderer"){
                              obj.renderer = incrementRenderer;
                            }
                            else if (obj.source === "getDropDownListValues"){
                              obj.source = await getDropDownListValues(obj.listname, obj.listColumn);
                            }
                            else if (obj.validator === "validateDateRequired"){
                              obj.validator = validateDateRequired;
                            }
                            else if (obj.validator === "getDescriptionValidator_PLF"){
                              obj.validator = getDescriptionValidator_PLF;
                            }
                            else if (obj.validator === "preventEdit"){
                              obj.validator = preventEdit;
                            }
                          }
                    });
              });
   return {
    colArray: _colArray,
    targetList: targetList,
    targetFilter: targetFilter
  };
}

var getSLF_ReferenceFormat_MajorType = async function(_web, key){
  var CDSFormat;

  var result = "";
  await _web.lists
        .getByTitle("MajorTypes")
        .items
        .select("CDSFormat")
        .filter("Title eq '" + key + "'")
        .get()
        .then(async function (items) {
            if(items.length > 0)
              //for (var i = 0; i < items.length; i++) {
                var item = items[0];
                CDSFormat = item.CDSFormat;
            });
 return CDSFormat;
}

function formatStringToArray(_cols){
  var partArray = [];
  if(_cols.includes(","))
{
    _cols = _cols.split(',');
    for(var i = 0; i < _cols.length; i++)
    {
      partArray.push(_cols[i].trim());
    }
}
  else partArray.push(_cols);
  return partArray;
}

var getDropDownListValues = async function(_listname, _column){
  var _colArray = [];

  await pnp.sp.web.lists
        .getByTitle(_listname)
        .items
        .select(_column)
        .filter(_column + " ne null")
        .get()
        .then(async function (items) {
            if(items.length > 0)
              for (var i = 0; i < items.length; i++) {
                  _colArray.push(items[i][_column]);
              }
            });
 return _colArray;
}

var validateItems = async function(element){
	var targetColIndex;
	var rowsIndex = [];

	var cols = element.getColHeader();
	var rowCount = element.countRows();
	//var rows = element.getData();
	//var dataWithHeaders = [cols].concat(rows);

	targetColIndex = await getTargetIndex(cols)

	for (var row = 0; row < rowCount; row++) {
		var cellValue = element.getDataAtCell(row, targetColIndex);
		if(cellValue !== null && cellValue !== '')
		 rowsIndex.push(row);
	  }

	  return new Promise((resolve) => {
		element.validateRows(rowsIndex, (valid) => {
		  resolve(valid);
		 });
	  });
}

var insertPLFItems= async function(element, targetList, targetFilter, trade){
	//var cols = element.getColHeader();
  var currentUser = await pnp.sp.web.currentUser.get();
  currentUser = currentUser.Title;
  var targetColIndex = columnProps.targetColumnIndex; //  await getTargetIndex(cols);

  //var rowCount = element.countRows();

	//var mentaInfo = '';
  // var _colsArray = element.getSettings().columns;
  // if(isContractor)
  //    mentaInfo = _colsArray.filter(item => item.owner == 'Contractor');
  // else mentaInfo = _colsArray.filter(item => item.owner == 'Dar');
  var mentaInfo = element.getSettings().columns;
  Reference = fd.field("Reference").value;
  
	var rows = element.getData();
  var rowCount = rows.length;
  var acceptedRow = 0;
  var emptyRows = 0;
  var status = '';
  var itemsToInsert = [];
  var objColumns = [];
  var objTypes = [];

  var maxItemNo = 0;

  if(_module === 'SLF'){
    var filter = targetFilter + " eq '" + Reference + "'";

    const list = pnp.sp.web.lists.getByTitle(targetList);
    const existingItems = await list.items
    .select("Id")
    .filter(filter)
    .getAll();

    if (existingItems.length > 0) 
     maxItemNo = existingItems.length;
  }

  for(var row = 0; row < rowCount; row++){
    var _objValue = { };
    var doInsert = false;
    var cellValue = rows[row][targetColIndex];

    if(cellValue!== null && cellValue !== ''){
      acceptedRow++;
      var _columns = '', _colType = '';
      for(var j = 0; j < mentaInfo.length; j++){
        var _value = rows[row][j];
        var internalName = mentaInfo[j].InternalName;
        if( internalName === 'Attachments')
           continue;
        else if(internalName== 'Status')
          status = _value;
        else if(internalName === 'ClosedDate'){
         
          if(status === 'Closed' && (_value === '' || _value === null || _value === undefined)){
           _value = new Date();
           _objValue['ClosedBy'] = currentUser;
          }
           else if(_value !== null && _value !== ''){
             const [day, month, year] = _value.split("/");
              _value = new Date(year, month - 1, day);
              _objValue['ClosedBy'] = currentUser;
           }
           else _value = null; 
        }

        else if(internalName === 'IsChecked' && _value === true){
           _objValue['RectifiedBy'] = currentUser;
        }

        else if(_value !== null && _value !== '')
        {
          if(mentaInfo[j].type === 'date'){
              const [day, month, year] = _value.split("/");
              _value = new Date(year, month - 1, day);
          }
        }
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
      _objValue['Title'] = Reference;

      if(!_isMain && trade !== '' && trade !== null && trade !== undefined)
       _objValue['Trade'] = trade;
     
      var itemNo = String(acceptedRow).padStart(5, '0');
      _objValue['ItemNo'] = itemNo; 

      _columns = _columns.endsWith(',') ? _columns.slice(0, -1) : _columns;
      _colType = _colType.endsWith(',') ? _colType.slice(0, -1) : _colType;

      objColumns.push(_columns);
      objTypes.push(_colType);
      itemsToInsert.push(_objValue);
    }
  }

  await insertItemsInBatches(itemsToInsert, targetList, targetFilter, objColumns, objTypes, trade, maxItemNo); 
}

var insertItemsInBatches = async function(itemsToInsert, targetList, targetFilter, objColumns, objTypes, trade, maxItemNo) {
  const list = pnp.sp.web.lists.getByTitle(targetList);

  const batch = pnp.sp.createBatch();
  
  var row = 0;
  var _columns, _colType;

  for (const item of itemsToInsert){
    var _query = '';

    _columns = objColumns[row];
    _colType = objTypes[row];

    if(_module === 'SLF'){
      debugger;
      if(_isPart && _role === 'Inspectors'){
      _query = targetFilter + " eq '" + Reference + "' and ItemNo eq '" + item.ItemNo + "'";
        if (trade) {
          _query += " and Trade eq '" + trade + "'";
        } 
        else {
          _query += " and Trade eq null";
        }
      }
      else _query = targetFilter + " eq '" + item.Title + "' and MasterItemNo eq '" + item.ItemNo + "'";
    }
    else _query = targetFilter + " eq '" + item.Title + "'";

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

        if(_type[i].toLowerCase() === 'date'){
           if(oldValue !== undefined && oldValue !== '' && oldValue !== null){
              const oldDateParts = oldValue.split("T")[0].split("-");

              var day = parseInt(oldDateParts[2]) +1;
              var month = parseInt(oldDateParts[1]);
              var year = parseInt(oldDateParts[0])
              oldValue = day + '-' + month + '-' + year;

              //const [_day, _month, _year] = currentValue.split("/")
              const date = new Date(currentValue);
              const _day = date.getDate(); // Get the day (1-31)
              const _month = date.getMonth() + 1; // Get the month (0-11). Adding 1 because months are zero-based.
              const _year = date.getFullYear(); // Get the full year

              currentValue = _day + '-' + _month + '-' + _year;
           }
        }
          if( oldValue != currentValue ){
            doUpdate = true;
            break;
          }
      }
      if(doUpdate)
       list.items.getById(existingItems[0].Id).inBatch(batch).update(item);
    } else {
      debugger;
      maxItemNo++;
      var formatMaxItemNo = String(maxItemNo).padStart(5, '0');
      item['MasterItemNo'] = formatMaxItemNo;
      _doHide = true; // set true to pop up message on submit to ask for attachment
       list.items.inBatch(batch).add(item);
    }
    row++;
  }
  await batch.execute();
}

const getDescriptionValidator_PLF = async (valueArray, callback) => {
  var rowIndex = valueArray.rowIndex;
  debugger;
  //var len = editedCells;

   if(!triggerDescription){
    triggerDescription = true;
    callback(true);
    return;
   }

   
  await creatErrorlbl();
  var errors = [];
  var colsError = {};
  var _cols = hot.getSettings().columns;
  isVisited = true;
  
  var columnIndex = valueArray.columnIndex;
  //var columnTitle = valueArray.columnTitle;
  var columnValue = valueArray.value;
  var itemNo = rowIndex +1;


  var itemNo = rowIndex +1;
    
    var cssClassName = defaultClassName;
    var metaObject = {
      className: cssClassName,
    };

    var isDescEmpty = true;

    if(columnValue !== '' && columnValue !== null)
      isDescEmpty = false;
    

    var haveValues = false;

    for(var i = 0;i<_cols.length;i++){
      var col = _cols[i];
      var colTitle = col.data.toLowerCase();
      if(colTitle === 'attachFiles') continue;
      var colType = col.type;
      var mesg = '';

      var _val = '';
      if(_val === '' || _val === null)
       _val = hot.getDataAtCell(rowIndex, i);

       if(_val === '' || _val === null){
        editedCells.find(item => {
            if(rowIndex === item[0] && item[1] === colTitle){
              _val = item[3];
              return ;
            }
        });
       }



      if(_val !== '' && _val !== undefined && _val !== null && i !== columnIndex) 
       haveValues = true;

     if(!isContractor && col['owner'] !== undefined && col['owner'].toLowerCase() === 'contractor'){
      hot.setCellMeta(rowIndex, i, 'readOnly', true);
      continue;
     }

      if(col['allowEmpty'] == false || colTitle === 'closed date'){
       
        // editedCells.find(item => {
        //     if(rowIndex === item[0] && item[1] === colTitle){
        //       _val = item[3];
        //       return ;
        //     }
        // });
        
        if(!isDescEmpty){
          if( (_val === '' || _val === null) && colTitle !== 'closed date'){
            cssClassName = errorClassName ;
            metaObject['className'] = errorClassName;
            mesg = htRequiredFieldMesg;
           }
          else if(colTitle === 'closed date' || (_val !== '' && _val !== null) ){
            var objResult = await checkValueColumnType(rowIndex, colType, i, _val);
            cssClassName =  objResult.cssClassName;
            metaObject['className'] = objResult.cssClassName;
            mesg = objResult._mesg;
          }
        }
        colsError['itemNo'] = itemNo;
        colsError['field'] = '<p style="color: black; display: inline;">' + colTitle + '</p>'; 
        if(mesg === undefined) {
          mesg = '';
          metaObject['className'] = defaultClassName;
        }
        colsError['message'] = mesg;
        errors.push(colsError);
        colsError = {};

        //hot.setCellMeta(rowIndex, i, 'className', cssClassName);
        if(mesg === '')
         metaObject['className'] = defaultClassName;
        hot.setCellMetaObject(rowIndex, i, metaObject);
        console.log(colTitle + ' class = ' + metaObject.className);
      }
    }

    var message = ''
    if(isDescEmpty && haveValues){
      message = htRequiredFieldMesg;
      callback(false);
     }
     else {
      message = '';
      callback(true);
     }

     colsError['itemNo'] = itemNo;
     colsError['field'] = '<p style="color: black; display: inline;">description</p>'; 
     colsError['message'] = message;
     errors.push(colsError);

    var errorResult = setErrortoMasterError(itemNo, errors);
    creatErrorlbl(errorResult);
    // if(editedCells.length === 1)
    // hot.render();
    
}

const getCellValidator = async (valueArray, callback) => {
  // var cellLength = editedCells.length;
  // setCallBack(cellLength, callback); // to Visit the trigger once
  
  debugger;
  var errors = [];
  var colsError = {};
  var _mesg;
  var isValid = true;
  var cssClassName = defaultClassName;
  

  var rowIndex = valueArray.rowIndex;
  var columnIndex = valueArray.columnIndex;
  var columnTitle = valueArray.columnTitle;
  var columnValue = valueArray.value;
  var itemNo = rowIndex +1;
  await creatErrorlbl();

  // if(columnValue == '' || columnValue == undefined){
  //   callback(true);
  //   return;
  // }
 
  var _cols = hot.getSettings().columns;
  isVisited = true;

  var targetColumnValue = '';
        editedCells.find(item => {
            if(item[0] === rowIndex && item[1] === columnProps.targetColumnTitle.toLowerCase()){
              targetColumnValue = item[3];
              return true;
            }
        });
  if(targetColumnValue === '' || targetColumnValue === null)
    targetColumnValue = hot.getDataAtCell(rowIndex, columnProps.targetColumnIndex);

  if(targetColumnValue === '' || targetColumnValue === null) //Description
      callback(true);
  else {
    triggerDescription = false;
      // Check if the cell is required
      var isRequired = false;
      var allowEmpty = _cols[columnIndex]['allowEmpty'];
      var colType = _cols[columnIndex]['type'];

      if(allowEmpty === false)
        isRequired = true;

      if (isRequired && (columnValue === '' || columnValue === null)) {
        cssClassName = errorClassName;
        _mesg = htRequiredFieldMesg;
        isValid = false;
      } 
      else if(columnValue !== '' && columnValue !== null){
        var objResult = await checkValueColumnType(rowIndex, colType, columnIndex, columnValue);
        cssClassName = objResult.cssClassName;
        _mesg = objResult._mesg;
        isValid = objResult.isValid;
      }
      
      if(_mesg !== undefined){
        hot.setCellMeta(rowIndex, columnIndex, 'className', cssClassName);
        colsError['itemNo'] = itemNo;
        colsError['field'] = '<p style="color: black; display: inline;">' + columnTitle + '</p>';
        colsError['message'] = _mesg;
        errors.push(colsError);
        colsError = {};
      }


      var checkMandatoryFields = false;
      const closingMandatoryColumnIndex = columnProps.closingMandatoryColumnIndex;
      const closingMandatoryColumnTitle = columnProps.closingMandatoryColumnTitle;
      const colsMandatoryForClosingIndexes = columnProps.colsMandatoryForClosingIndexes;
      if(columnIndex == closingMandatoryColumnIndex) 
        checkMandatoryFields = true;
      else if(colsMandatoryForClosingIndexes !== undefined && colsMandatoryForClosingIndexes.length > 0){
      colsMandatoryForClosingIndexes.map((val) =>{
            if(val === columnIndex) {
              checkMandatoryFields = true;
              return;
            }
          })
      }
    
      if(checkMandatoryFields){

        if(columnProps.closingMandatoryColumnIndex === columnIndex && columnValue === 'Cancelled'){
          hot.setCellMeta(rowIndex, columnIndex, 'className', cancelClassName);
        }
        else {
            var error = await checkMandatoryColumns(itemNo, rowIndex, columnIndex, columnTitle, columnValue, checkMandatoryFields, closingMandatoryColumnIndex, closingMandatoryColumnTitle, colsMandatoryForClosingIndexes)
            if(error !== undefined && error.length > 0){
              for(var i =0; i < error.length; i++){
                colsError['itemNo'] = error[i].itemNo;
                colsError['field'] =  error[i].field;
                colsError['message'] = error[i].message;
                errors.push(colsError);
                colsError = {};
              }
            }
            else{
              hot.setCellMeta(rowIndex, columnIndex, 'className', cssClassName);
              colsError['itemNo'] = itemNo;
              colsError['field'] = '<p style="color: black; display: inline;">' + columnTitle + '</p>';
              colsError['message'] = '';
              errors.push(colsError);
              colsError = {};
            }
         }
       }

      var errorResult = setErrortoMasterError(itemNo, errors);
      creatErrorlbl(errorResult);
      // if(callback.length === 1)
      //   hot.render();
        
      callback(isValid);
   }
}

const preventEdit = async (valueArray, callback) => {
  var columnValue = valueArray.value;
   if(!isContractor && columnValue === true){
      var rowIndex = valueArray.rowIndex;
      var columnIndex = valueArray.columnIndex;
      hot.setDataAtCell(rowIndex, columnIndex, false);
      hot.setCellMeta(rowIndex, columnIndex, 'readOnly', true);
   }
   callback(true);
}

const checkMandatoryColumns = async(itemNo, rowIndex, columnIndex, columnTitle, columnValue, checkMandatoryFields, closingMandatoryColumnIndex, closingMandatoryColumnTitle, colsMandatoryForClosingIndexes) =>{
  var errors = [];
  var colsError = {};
  var _mesg = '';

  if(checkMandatoryFields){
    if(closingMandatoryColumnIndex !== '' && closingMandatoryColumnIndex !== undefined && colsMandatoryForClosingIndexes !== undefined){

      if(this._module === 'SLF'){
        for(var i = 0; i<colsMandatoryForClosingIndexes.length; i++){
          
          var colsingColumnIndex = colsMandatoryForClosingIndexes[i];
          var colsingColumnTitle = columnProps.colsMandatoryForClosingTitles[i];

          var cssClassName = defaultClassName;
          var metaObject = {
            readOnly: false,
            className: cssClassName,
          };

          // if(columnTitle.toLowerCase() === colsingColumnTitle.toLowerCase())
          //    return;

          var _closingVal = '';
          var statusVal = ''
          if(columnValue === 'Open' || columnValue === 'Closed'){
            statusVal = columnValue;

            editedCells.find(item => {
                if(rowIndex === item[0] && item[1] === colsingColumnTitle.toLowerCase()){
                  _closingVal = item[3];
                  return ;
                }
            });

            if(_closingVal === '' || _closingVal === null)
            _closingVal = hot.getDataAtCell(rowIndex, colsingColumnIndex);

            if(columnValue === 'Closed' && _closingVal !== '' && _closingVal !== null){
              //metaObject['readOnly'] = true;
               //hot.setCellMeta(rowIndex, colsingColumnIndex, 'readOnly', false);
            }
            else  if(columnValue === 'Open' && (_closingVal === '' || _closingVal === null)){}
            else{
              metaObject['className'] = errorClassName;
              _mesg = htremoveValueMesg;
            }
          }
          else{
            editedCells.find(item => {
              if(rowIndex === item[0] && item[1] === closingMandatoryColumnTitle.toLowerCase()){
                statusVal = item[3];
                return true;
              }
            });

            if(statusVal === '' || statusVal === null)
             statusVal = hot.getDataAtCell(rowIndex, closingMandatoryColumnIndex);
            if(statusVal !== '' && statusVal !== undefined && statusVal !== null)
            statusVal = statusVal
            _closingVal = columnValue;
          }
          
          if(statusVal === 'Closed' && (_closingVal === '' || _closingVal === null)){
            metaObject['className'] = errorClassName;
            _mesg = htRequiredFieldMesg;
          }
          else if(statusVal === 'Open' && (_closingVal !== '' && _closingVal !== null)){
            metaObject['className'] = errorClassName;
            _mesg = htremoveValueMesg;
          }

          hot.setCellMetaObject(rowIndex, colsingColumnIndex, metaObject);
          colsError['itemNo'] = itemNo;
          colsError['field'] = '<p style="color: black; display: inline;">' + hot.getSettings().columns[colsingColumnIndex]['data'] + '</p>';
          colsError['message'] = _mesg;
          errors.push(colsError);
          colsError = {};
        }
      }
    }
  }
  return errors;
}

const checkValueColumnType = async (rowIndex, colType, columnIndex, columnValue) =>{

  var cssClassName = defaultClassName;
  var _mesg = '';
  var isValid = true;
  if(colType === 'dropdown'){
    if (!hot.getSettings().columns[columnIndex].source.includes(columnValue)){

      if(columnProps.closingMandatoryColumnIndex === columnIndex && columnValue === 'Cancelled'){
        cssClassName = cancelClassName;
      }
      else{
        cssClassName = errorClassName;
        _mesg = htdropdownMesg;
        isValid = false;
      }
    }
  }
  else if(colType === 'date'){
    const dateRegex = /^\d{2}\/\d{2}\/\d{4}$/;
    if (dateRegex.test(columnValue)) {
      const parts = columnValue.split('/');
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // Month is zero-based in JavaScript Date object
      const year = parseInt(parts[2], 10);
      const date = new Date(year, month, day);
      if (date.getDate() === day && date.getMonth() === month && date.getFullYear() === year) {
        const closingMandatoryColumnIndex = columnProps.closingMandatoryColumnIndex;
        if(closingMandatoryColumnIndex !== '' && closingMandatoryColumnIndex !== undefined){

          // getValueBeofe
          var statusVal = '';
          var colsingColumnTitle = columnProps.colsMandatoryForClosingTitles[0];

          editedCells.find(item => {
            if(item[0] === rowIndex && item[1] === colsingColumnTitle.toLowerCase()){
              statusVal = item[3];
              return true;
            }
          });

          if(statusVal === '' || statusVal === null)
            statusVal = hot.getDataAtCell(rowIndex, closingMandatoryColumnIndex); 

          // if(statusVal === '' || statusVal === null)

          if(statusVal === 'Open' && columnIndex === columnProps.colsMandatoryForClosingIndexes[0]){
              cssClassName = errorClassName;
              _mesg = htremoveValueMesg;
              isValid = false;
          }    
        }
      }
      else{
        cssClassName = errorClassName;
        _mesg = 'date format is not correct';
        isValid = false;
      }
    }
    else{

      if(columnValue !== '' && columnValue !== null && columnValue !== undefined){
        cssClassName = errorClassName;
        _mesg = htremoveValueMesg;
        isValid = false;
      }

     else{
          var statusVal = hot.getDataAtCell(rowIndex, columnProps.closingMandatoryColumnIndex);
          if(statusVal === null){
            var colsingColumnTitle = columnProps.colsMandatoryForClosingTitles[0];
            editedCells.find(item => {
              if(item[0] === rowIndex && item[1] === colsingColumnTitle.toLowerCase()){
                statusVal = item[3];
                return true;
              }
            });
          }

            if(statusVal !== '' && statusVal !== null && statusVal !== undefined){
                if(isValid){
                  
                  if(statusVal === 'Closed' && (columnValue === '' || columnValue === null)){
                    cssClassName = errorClassName;
                    _mesg = htRequiredFieldMesg;
                    isValid = false;
                  }
                }
            }
    }
   }
  }

  return {
    cssClassName: cssClassName,
    _mesg: _mesg,
    isValid: isValid
  }
}

function setErrortoMasterError(itemNo, Error, isDelete){

  if(isDelete && itemNo.length > 0){
    var startIndex;

    itemNo.forEach(item => {
      if(startIndex === undefined)
        startIndex = item;

      if(masterErrors[item]) {
        delete masterErrors[item];
      }
    });

    //UPDATE ITEM NO's
    if(startIndex > 0){
      const updatedMasterErrors = Object.fromEntries(
        Object.entries(masterErrors).map(([key, value]) => {
          if (key > startIndex) {
            var result = [startIndex, value];
            startIndex++;
            return result;
          } 
          else return [key, value];
        })
      );
      //console.log('updatedMasterErrors ' + updatedMasterErrors);
      masterErrors = updatedMasterErrors;
    }

    creatErrorlbl(masterErrors);
    return;
  }


    Error.map(val => {
      var _itemNo = val.itemNo;

      if (!masterErrors.hasOwnProperty(_itemNo)) {
        masterErrors[_itemNo] = [];
      }

      var existingField = masterErrors[_itemNo].find(item => {
           return item.field.toLowerCase() === val.field.toLowerCase()
      });
      if (existingField) {
        // Update existing field message
        existingField.message = val.message;
      } else {
        // Add new field and message
        masterErrors[_itemNo].push({ field: val.field, message: val.message });
      }
    });

    Object.keys(masterErrors).forEach(itemNo => {
      masterErrors[itemNo] = masterErrors[itemNo].filter(item => {
        //var mesg = item.message;
        return item.message.trim() !== ''
      });
    
      if (masterErrors[itemNo].length === 0) {
        delete masterErrors[itemNo];
      }
    });
   //console.log(masterErrors);
  return masterErrors;
}

const creatErrorlbl = async(errArray) =>{

  var id = '#errorlbl';
  if ($(id).length === 0) {
    var label = $('<label>', {
      id: 'errorlbl',
      text: ''
    });
    label.css('color', 'red');
    label.css('width', '950px'); 
    $('#dt').after(label);
  }

  if(errArray === null || errArray === undefined)
    return;

  let itemNo = '', errorText = '';
  for (const [key, value] of Object.entries(masterErrors)) {
    errorText += `itemNo[${key}]: {`;
    value.forEach((item) => {
      errorText += `${item.field}: ${item.message}     `;
    });
    errorText += `}<br>`;
  }

  $(id).html(errorText);
}

var setRowsReadOnly = async(rowIndex, setReadOnly, isContractor) => {

  var setOpen = false;
  var metaObject = {
    readOnly: false,
    className: closedClassName
  };

  if(setReadOnly === true){
   metaObject['readOnly'] = true;
  }
  else setOpen = true;

  var statusColumnIndex = columnProps.closingMandatoryColumnIndex;
  var closedStatus = columnProps.closedStatus;

  if(rowIndex !== null && rowIndex !== undefined){
    if(setOpen) 
     metaObject['className'] = defaultClassName;
    else metaObject['className'] = cancelClassName;
    var status = hot.getDataAtCell(rowIndex, statusColumnIndex);
      
    var columnIndex = 0;
    hot.getSettings().columns.find(item => {
      hot.setCellMetaObject(rowIndex, columnIndex, metaObject);
      columnIndex++;
    });
    if(setOpen)
    hot.setDataAtCell(rowIndex, statusColumnIndex, 'Open');
    else hot.setDataAtCell(rowIndex, statusColumnIndex, 'Cancelled');
  }
  else{
    var cancelMetaObject = {
      readOnly: true,
      className: cancelClassName
    };
    for (var row = 0; row < hot.countRows(); row++) {

     if(_isPart && Status === 'Completed' && !_isTeamLeader){
      var columnIndex = 0;
      hot.getSettings().columns.find(item => {
        hot.setCellMeta(row, columnIndex, 'readOnly', true);
        columnIndex++;
      });
     }

      var status = hot.getDataAtCell(row, statusColumnIndex);
      var masterMetaObject = metaObject;

       if(status === 'Open'){
          var columnIndex = 0;
          hot.getSettings().columns.find(item => {
            if(isContractor && item.owner !== undefined && item.owner.toLowerCase() === 'dar')
              hot.setCellMeta(row, columnIndex, 'readOnly', true);
            else if(isContractor && item.owner.toLowerCase() === 'tl') 
              hot.setCellMeta(row, columnIndex, 'readOnly', true);
            else if(!isContractor && item.owner !== undefined && item.owner.toLowerCase() === 'contractor')
              hot.setCellMeta(row, columnIndex, 'readOnly', true);

            else if(_isPart){
              if(!_isTeamLeader && (statusColumnIndex === columnIndex || statusColumnIndex + 1 === columnIndex) ){}
              else if(_isTeamLeader || Status === 'Completed')
                hot.setCellMeta(row, columnIndex, 'readOnly', true);
              else if(isResubmit){
                if(item.isResubmit === true || item.hide === true && item.owner == 'Dar'){}
                else 
                {
                  if(RejectionTrades === '')
                   hot.setCellMeta(row, columnIndex, 'readOnly', true);
                }
              }
            }
            columnIndex++;
          });
       }

       if (status === closedStatus || status === 'Cancelled') {
        if(status === 'Cancelled')
         masterMetaObject = cancelMetaObject;

        var columnIndex = 0;
        hot.getSettings().columns.find(item => {
          if(item.InternalName !== 'Attachments'){
           hot.setCellMetaObject(row, columnIndex, masterMetaObject);
          }
          columnIndex++;
        });
      }
    }
  }
  hot.render();
}

const setCallBack_notUsed = async (cellLength, callback) =>{
  if(isVisited){
    callback(true);
    return;
    // if(cellLength > 1){
    //     visitIndex++
    //     if(visitIndex === cellLength){
    //       isVisited = false;
    //       visitIndex = 1;
    //     }
    //     callback(true);
    //     return;
    // }
  }
  else isVisited = true;
}

var addListItem_notUsed = async function(_listname, objValue, _columns, _colType, targetFilter, Reference, ItemNo){
  var _web = pnp.sp.web;
  var filter = '';

  if(_module === 'SLF')
   filter = targetFilter + " eq '" + Reference + "' and ItemNo eq '" + ItemNo + "'";
  else filter = targetFilter + " eq '" + Reference + "'";

  await _web.lists
  .getByTitle(_listname)
  .items
  .select("Id," + _columns)
  .filter(filter)
  .get()
  .then(async function (items) {
    var _cols = { };
      if(items.length == 0){
          await _web.lists.getByTitle(_listname).items.add(objValue);
          return;
      }

      else if(items.length > 0){
            var _item = items[0]; // _item is the oldItem values while objValue is the new one
            var _cols = _columns.split(',');
            var _type = _colType.split(',');
            var doUpdate = false;

            for(var i = 0; i < _cols.length; i++){
              var columnName = _cols[i];
              var oldValue = '';
              var currentValue = '';

              oldValue = _item[columnName];
              currentValue = objValue[columnName];

              if(_type[i].toLowerCase() === 'date'){
                  const oldDateParts = oldValue.split("T")[0].split("-");

                  var day = parseInt(oldDateParts[2]) +1;
                  var month = parseInt(oldDateParts[1]);
                  var year = parseInt(oldDateParts[0])
                  oldValue = day + '-' + month + '-' + year;
  
                  //const [_day, _month, _year] = currentValue.split("/")
                  const date = new Date(currentValue);
                  const _day = date.getDate(); // Get the day (1-31)
                  const _month = date.getMonth() + 1; // Get the month (0-11). Adding 1 because months are zero-based.
                  const _year = date.getFullYear(); // Get the full year

                  currentValue = _day + '-' + _month + '-' + _year;
              }

              if( oldValue != currentValue ){
                doUpdate = true;
                break;
              }
            }

            if(doUpdate)
              await _web.lists.getByTitle(_listname).items.getById(_item.Id).update(objValue);
            return;
          }
      });  
}

var validateSelectedRow_notUsed = async function(element, rowIndex){
	return new Promise((resolve) => {
		element.validateRows(rowIndex, (valid) => {
		  resolve(valid);
		});
	});
}

var setAttachFiles = async function(){
  var data = hot.getData();
  for (var rowIndex = 0; rowIndex < data.length; rowIndex++) {
    var row = data[rowIndex];
    var attachIndex =0;

    hot.getSettings().columns.find(item => {
      if(item.data === 'attachFiles')
        return attachIndex;
      attachIndex++;
    });

    hot.setCellMeta(rowIndex, i, 'readOnly', true);
  }


  
}