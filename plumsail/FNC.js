var _layout, _module = '', _formType = '', _acronym = '', _mainList = '';

var _Fields = [], _fieldsInernalName = [], _schema = [], selectedFieldsSchema = [], selectedFieldsInternalName = [];
var _features =      ['SpanField', 'isList', 'RevStartWith', 'isText', 'textLength', 'AllowedCharacters', 'isOptionalField'];
var _featuresType =  ['number',    'string', 'string',       'string', 'string',     'string',            'boolean'];
var _featuresWidth = ['10%',       '16%',    '12%',          '16%',    '12%',        '12%',               '12%'];

var _isNew = false, _isEdit = false;


var isSort = false, isTblCreated = false, isSpanExist_InSchema = false, isOptionalExist_InSchema = false, disableIsList = false, disableIsText = false;

var listBoxFields = [];
var validationlistBoxPromises = [];
var _utilsPath;

var onRender = async function (relativeLayoutPath, moduleName, formType){
    _module = moduleName;
    _formType = formType;
    _layout = relativeLayoutPath;
    
    _utilsPath = _layout + '/plumsail/js/utilities.js';

    var script = document.createElement("script"); // create a script DOM node
    script.src = _layout + "/plumsail/js/config/configFileRoutingFNC.js"; // set its src to the provided URL
    document.head.appendChild(script);
    await delay(500);

    if($('.text-muted').length > 0)
     $('.text-muted').remove();

    //preloader();
    if(_formType === 'New')
      _isNew = true;
    else if(_formType === 'Edit')
      _isEdit = true;

    if(_isNew)
    {
        fd.field('DeliverableType').clear();
        fd.field('Acronym').clear();
        fd.field('MainList').clear();
        fd.field('Schema').clear();

        var result = await getDeliverableTypes();
        if(result !== undefined && result.length > 0){
            var _query = '';
            var isFirstVisit = true;
            result.map(item =>{
                if(isFirstVisit){
                 _query = `Title ne '${item}'`;
                 isFirstVisit = false;
                }
                else _query += ` and Title ne '${item}'`;
            });

            fd.field('DeliverableType').filter = _query; //`Title ne 'DWG' and Title ne 'DOC'`;  
            fd.field('DeliverableType').refresh();
        }
    }
    else {
        var  schemaValue = fd.field("Schema").value.replace(/&nbsp;/g, '');
        _schema = JSON.parse(schemaValue);

        fd.field('DeliverableType').disabled = true;
        fd.field('Delimeter').disabled = true;

        selectedFieldsSchema = _schema.filter(item => {
            if(item.InternalName !== 'properties'){
                selectedFieldsInternalName.push(item.InternalName);
                if(item.isOptionalField === true)
                  isOptionalExist_InSchema = true;
                else if(item.SpanField !== undefined)
                  isSpanExist_InSchema = true;
                return true;
            }
         });

         let index = 0;
         for (let i = 0; i < selectedFieldsInternalName.length; i++) {
            disableIsList = false; 
            disableIsText = false;
            var value = selectedFieldsInternalName[i]
            await setHTML_Column_PerRow(index, value, 'Add');
            index++;
        }
    }

    $(fd.field('Schema').$parent.$el).hide();
    $(fd.field('Title').$parent.$el).hide();

    var lookupField = fd.field('DeliverableType');
    disableFields();

    await getDeliverableMetaInfo(lookupField);

    _spComponentLoader.loadScript(_htL_utilsPathibraryUrl).then(setButtons);
    //preloader("remove");
}

var getDeliverableMetaInfo = async function(lookupField){
    if(_isNew){
        lookupField.$on('change', async function(value) {
            await getTypeResult(value);
        });
    }
    else await getTypeResult(lookupField.value);
}

var getTypeResult = async function(value){
    if (value) {
        var result = await getMajorType(value.Id, true)
        if(result.length > 0){
            _fieldsInernalName = [];
            _Fields = [];

          _acronym = result[0].Title;
          _mainList = result[0].MainList;
          fd.field('Acronym').value = _acronym;
          fd.field('MainList').value = _mainList;
        
          await getListFields();
        }
    }
}

var getListFields = async function(){
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
    'WorkflowVersion',

    'Title',
    'Status',
    'PDFSize',
    'DWGSize',
    'DWFSize',
    'FullFileName',
    'FileName',
    'LeadTrade',
    'LeadStatus',
    'CDSTitle',
    'Code',
    'State',
    'CDSNumber',
    'CRSNumber',
    'GeneralStatus',
    'ReviewOffice',
    'TransmittalNo',
    'BIC',
    'SubmitToPMC',
    'GeneralState',
    'PendingCDS',
    'PendingFileName',
    'LODRef',

    'AttachFiles',
    'FinalResponse',
    'HandledBy',
    'Response',
    'Editedby',
    'Question',
    'Last_x0020_Edited_x0020_Date',
    'FullRef',
    'ORFI',
    'Indicator_KPIInternal',
    'LeadAction',
    'Remarks',
    'InspPurpose',
    'InspDesc',
    'Submittedby',
    'SubmittedBy',
    'EngComment',
    'InstDetails',
    'ResponsedBy',
    'RE',
    'ReasonofRejection',
    'Box_x0020_ID',
    'Reference',
    'PartTrades'
  ];

  await pnp.sp.web.lists.getByTitle(_mainList).fields.select("Title", "InternalName").orderBy("Title").get()
  .then(fields => {
    _Fields = fields.filter(field => {
        var internalName = field.InternalName;
        var fieldType = field['odata.type'];
        if(!excludedFieldNames.includes(internalName) && !internalName.includes('_Part') && fieldType !== 'SP.FieldDateTime' && fieldType !== 'SP.FieldUrl' 
            && fieldType !== 'SP.FieldUser' && fieldType !== 'SP.Field')
        {
          if(!_fieldsInernalName.includes(internalName))
          {
            if(_isEdit && selectedFieldsInternalName.includes(internalName))
                return false;
            
            _fieldsInernalName.push(internalName);
          }
          return !excludedFieldNames.includes(internalName);
        }
      });
      console.log(_Fields);
      console.log(_fieldsInernalName);
  })
  .then(() =>{
    bindHTMLControls();
  })
  .then(() =>{
    setListBox();
  })
  .catch(error => {
    console.error("Error", error);
  });
}

function setListBox(){
    var fieldName = '';
    var isFieldVisited = false;
    $("#listBoxA").jqxListBox({ allowDrop: true, allowDrag: true, filterable: true, source: _fieldsInernalName, width: 300, height: 240, 
        dragStart: function (item) { // item.index can be used
            //fieldName = item.label;
            //isFieldVisited = false;
        },
        renderer: function (index, label, value) {
            if (label == "Breve") {
                return "<span style='color: red;'>" + label + "</span>";
            }
            return label;
        }
    });

    const listBox = $("#listBoxB");
    listBox.jqxListBox({ allowDrop: true, allowDrag: true, filterable: true, width: 300, height: 240,
        dragEnd: function (dragItem, dropItem) {
            if(dropItem !== null){
                let position = event.clientX;
                if(position > 400)
                  isSort = true;
                else isSort = false;
            }

             if(dropItem === null){
               dropItem = dragItem;
               isSort = false;
             }
             var columnName = dragItem.label;
             setHTML_Column_PerRow(dragItem.index, columnName, 'remove', dropItem);
             //console.log('dragEnd[listBoxB]: ' + columnName + ', ' + dragItem.index);
            
        },
        renderer: function (index, label, value, operation) {
            var lisboxBLength = $('#listBoxContentlistBoxB div div').length;
             if(lisboxBLength > 0 && index === 0)
              listBoxFields = [];

              if(!listBoxFields.includes(label))
                listBoxFields.push(label);

                var _Row = $('#tblColumns').find('tr').filter(function() {
                    return $(this).find('td:first').attr('id') === label;
                });

                if(_Row.length === 0){
                //NOTE:EACH VALUE IT IS RENDERED TWICE UNTIL _Row.length BECAME > 0
                  if(!isFieldVisited){
                    isFieldVisited = true;
                    validationlistBoxPromises.push(setHTML_Column_PerRow(index, label, 'Add'));
                  }
                  else isFieldVisited = false;
                }
            return label;
        }
    });

    $(listBox).on('change', async function (event) {
        debugger;
        await Promise.all(validationlistBoxPromises);
        var lisboxBLength = $('#listBoxContentlistBoxB div div').length;
        if(listBoxFields.length == lisboxBLength){
           //REORDER TABLE ROWS
          if(listBoxFields.length > 0){
              var tbl = $('#tblColumns tbody');
              listBoxFields.map(item => {
                  var _Row = $('#tblColumns').find('tr').filter(function() {
                      return $(this).find('td:first').attr('id') === item;
                  });
                  var nextRow = $(_Row).next();

                  tbl.append(_Row, nextRow);
              });
          }
        }
    });

    if(_isEdit){
        var dataLength = selectedFieldsInternalName.length;
        for (let i = 0; i < dataLength; i++) {
            var value = selectedFieldsInternalName[i]
            listBox.jqxListBox('addItem', value);
        }
    }
}

var bindHTMLControls = async function(){

    //#region LISTBOX CONTROL
    if($('label.title-label').length === 0){
        var div1 = $('<div>').css({'float': 'left', 'margin-right': '20px'});
        var label1 = $('<label>').attr('for', 'listFields').addClass('title-label').text('List Fields');
        var listBoxA = $('<div>').attr('id', 'listBoxA');
        div1.append(label1, listBoxA);

        // Create the second div with ListBoxB
        var div2 = $('<div>').css('float', 'left');
        var label2 = $('<label>').attr('for', 'filenameFields').addClass('title-label').text('Filename Fields');
        var listBoxB = $('<div>').attr('id', 'listBoxB');
        div2.append(label2, listBoxB);

        // Append both divs to the body
        $('#contentId').append(div1, div2);
    }
    //#endregion
}

var setHTML_Column_PerRow = async function(columnIndex, columnName, operation, dropItem){

    var contentElement = $('div.col-sm-12').eq(1);
    var tblId = 'tblColumns';
    var tblElement = $('#' + tblId);

    if(operation === 'remove')
    {
       var colElement = $(`#${columnName}`);
       if(colElement.length > 0){
            var closestTr = $(colElement).closest("tr");
            var nextTr = $(closestTr).next();

            if (!isSort)
            {
              $(nextTr).remove();
              $(closestTr).remove();
            }
            else{
                // SWAP THE ITEM TO ITS POSITION
                tblElement.find('tr').each(function(index) {
                    if(dropItem !== null && index === dropItem.index)
                    {
                        var rowId = columnName + 'col' + columnIndex;
                        var valueRowId = 'ctlr' + rowId;

                        var rowToMoveId = $('#' + rowId);
                        var rowToMove = rowToMoveId[0].parentElement.parentElement;
                        var rowToMoveIndex = $(rowToMove).index();

                        var rowValueToMoveId = $('#' + valueRowId);
                        var rowValueToMove = rowValueToMoveId[0].parentElement.parentElement;

                        var targetRow = 1;
                        if(index > 0)
                          targetRow = (index*2) + 1;

                        var newPositionRow = $(tblElement).find('tr:nth-child(' + targetRow + ')');
                        var newPositionRowIndex = $(newPositionRow).index();
                        var newPositionControlsRow = $(tblElement).find('tr:nth-child(' + (targetRow+1) + ')');

                        if(rowToMoveIndex < newPositionRowIndex){
                            $(rowToMove).insertAfter(newPositionControlsRow);
                            $(rowValueToMove).insertAfter(rowToMove);
                        }
                        else{
                            $(rowToMove).insertBefore(newPositionRow);
                            $(rowValueToMove).insertAfter(rowToMove);
                        }
                        return;
                    }
                });
            }
       }

       if($(tblElement).find('tr').length === 0)
         $(tblElement).remove();
      return;
    }

    if(!isTblCreated){
        isTblCreated = true;
        var tbl = "<table id='" + tblId + "' width='100%' class='customtbl'>";

        var ticksRow = await renderFeaturesTicks(columnName);
        var controlsRow= await renderFeaturesControls(columnName);

        tbl += ticksRow + controlsRow;
        tbl += "</table>";
        contentElement.after(tbl);
    }
    else{
        //append row into html table
        var ticksRow = await renderFeaturesTicks(columnName);
        var controlsRow= await renderFeaturesControls(columnName);
        const newRow = ticksRow + controlsRow;
        $(tblElement).append(newRow);
    }

  //$('#contentId').append(tbl);
  //$(tbl).insertAfter("#contentId");
}

var renderFeaturesTicks = async function(columnName){
    
    var row = "<tr>";
        row += "<td id='" + columnName + "' rowspan='2' width='10%' class='borderStyle'><label class='lblcolumnyStyle'>" + columnName + "</label></td>";

    for(var i =0; i < _features.length; i++){
        var checkBoxId = columnName + 'col' + i; //columnIndex is the rowIndex
        var controlId = 'ctlr' + columnName + 'col' + i;
        var featureName = _features[i];

        if(featureName === 'isOptionalField')
          row += "<td rowspan='2' width='" + _featuresWidth[i] + "' class='borderStyle'>";
        else row += "<td width='" + _featuresWidth[i] + "' class='borderStyle'>";
                    
             row += "<input type='checkbox' id='" + checkBoxId + "' onclick='checkFeature(this)' name='" + checkBoxId + "' class='ckStyle' refId='" + controlId + "'";

             row += await setFeatureCheckBoxes(featureName, columnName);

             row += "<span for='" + checkBoxId + "' class='lblpropertyStyle'>" + featureName + "</span>" +
         "</td>" ;
    }
    row += "</tr>";
    return row;
}

var setFeatureCheckBoxes = async function(featureName, columnName){
    var row = '';
    var result = false;

    if(columnName === 'Revision' || columnName === 'Rev'){
       if(featureName === 'SpanField' || featureName === 'isList' || featureName === 'isText' || featureName === 'isOptionalField')
          row += ' disabled>';
       else{ 
        if(!_isNew){

            result = await isFeatureFound(columnName, featureName);
            if(result.isFound)
              row += ' checked';
        }
         row += '>';
      }
    }
    else if(featureName === 'RevStartWith' || featureName === 'SpanField'){
        if(!_isNew){
            result = await isFeatureFound(columnName, featureName);
            if(result.isFound) 
                row += ' checked>';
            else row += ' disabled>';
        }
        else row += ' disabled>';
    }
    else {
        if(!_isNew){
            result = await isFeatureFound(columnName, featureName);
            if(result.isFound){
              row += ' checked';
              if(featureName === 'isList')
                disableIsText = true;
              else if(featureName === 'isText')
                disableIsList = true;
            }
            else{
                 if(featureName === 'isOptionalField'){ 
                   if(isOptionalExist_InSchema === true)
                     row += ' disabled';
                     disableIsText = false;
                 }
                 else if(featureName === 'SpanField' && isSpanExist_InSchema === true)
                   row += ' disabled';
                else if(disableIsText === true || disableIsList === true){ //featureName === 'isText' && 
                   row += ' disabled';
                   //disableIsText = false;
                }
                else if(disableIsList && (featureName === 'textLength' || featureName === 'AllowedCharacters')){
                    row += ' disabled';
                }
                else if(result.disableFeature)
                    row += ' disabled';
            }
        }
         row += '>';
    }
    return row;
}

var renderFeaturesControls = async function(columnName){
    var row = "<tr>";
    for(var i =0; i < _features.length; i++){
        var controlId = 'ctlr' + columnName + 'col' + i;
        var columnType = _featuresType[i];
        var featureName = _features[i];

        if(columnType !== 'boolean'){
            row += "<td class='borderStyle'>";
             row += "<input type='" + columnType + "' id='" + controlId + "' style='width: 100% !important;";// display:none;'>";
            
                if(!_isNew){
                    var isFound = false;
                    var _controlValue = '';
                    _schema.map(item => {
                        if(item.InternalName === columnName && item.hasOwnProperty(featureName)){
                            isFound = true;
                            _controlValue = item[featureName];
                            return;
                        }
                    });

                    if(!isFound)
                    row += "display:none;'>";
                    else row += "' value='" + _controlValue + "'>";
                }
                else row += "display:none;'>";
            }
            row += "</td>";
     }
    
    row += "</tr>";
    return row;
}

var isCorrect = async function(){
    var tblElement = $('#tblColumns');
    var columnsLength = tblElement.find('tr').length / 2;
    var revIndex = 0;
    
    var colIndex = 0;
    tblElement.find('tr').each(function(index) {
        if(index % 2 === 0){
        colIndex++;
        var td = $(this).children();
        var colName = td[0].id;
        if(colName === 'Revision' || colName === 'Rev')
            revIndex = colIndex;
        }
    });

    if(revIndex !== columnsLength)
        return false
    else return true;
}

var isFeatureFound = async function(columnName, featureName){
    var isFound = false, disableFeature = false;
    _schema.map(item => {
        if(item.InternalName === columnName ){
            if(item.hasOwnProperty(featureName))
                isFound = true;
            
           else{
                if(featureName === 'isText' || featureName === 'isList'){
                    if(item.isText === undefined && item.isList === undefined){
                        if(item.textLength !== undefined || item.AllowedCharacters !== undefined)
                            disableFeature = true;
                    }

                    if(featureName === 'isText' && item.isList !== undefined)
                        disableFeature = true;
                    else if(featureName === 'isList' && item.isText !== undefined)
                        disableFeature = true;
                }
           }
           return;
        }
     });

     return {
        isFound: isFound,
        disableFeature: disableFeature
      };
}

//#region GENERAL FUNCTIONS
var setButtons = async function () {
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

    await customButtons("Accept", 'Submit', false, '', false, true, false, false, _module);
    await customButtons("ChromeClose", "Cancel", false);
}

var disableFields = async function(){
    fd.field('Acronym').disabled = true;
    fd.field('MainList').disabled = true;
    

    fd.field("DeliverableType").$on("change", function (value) {
        if(value !== null)
         fd.field('DeliverableType').disabled = true;
         fd.field('Delimeter').disabled = true;
    });
}

function delay(time) {
    return new Promise(function(resolve) { 
        setTimeout(resolve, time)
    });
}

var checkColumnsFeatureSelection = async function(isOptional, tickedFieldRefId, isTicked){
    var haveError = false;
    var _message = '';

    var validationPromises = [];
    
    $('#tblColumns').find('tr').each(function(index){
        if(index % 2 === 0)
        {
            var td = $(this).children();
            var columnName = td[0].id;

            if(isOptional === true){
              var tdColumnName = td[(td.length-1)];
              var inputControl = $(tdColumnName).find('input');
              const refId = '#' + inputControl.attr("refid");

               if(tickedFieldRefId !== refId){
                if(isTicked)
                 inputControl.prop('disabled', true);
                else if(columnName !== 'Revision' && columnName !== 'Rev')
                 inputControl.prop('disabled', false);
               }
            }
            else{
                var inputsInTds = $(td).find('input:checked');
                if(inputsInTds.length === 0){
                  _message += `${columnName}: at least one feature should be selected <br/><br/>`;  
                  haveError = true;      
                }
                if(inputsInTds.length === 1 && inputsInTds[0].nextSibling.innerText === 'isOptionalField'){
                    _message += `${columnName}: at least one feature should be selected <br/><br/>`;  
                    haveError = true;      
                }
                else {
                    inputsInTds.each(function(row, column){
                      let featureName = column.nextSibling.innerText;
                      const valueCtlrId = '#' + $(this).attr("refid");
                      var controlValue = $(valueCtlrId).val();

                      if( controlValue !== undefined)
                        controlValue = controlValue.trim().replace(' ', '');

                      if(controlValue === ''){
                        _message += `${columnName}: kindly mention the value for ${featureName} feature <br/><br/>`;  
                        haveError = true;    
                      }
                      else if(featureName === 'isList'){
                        if(!controlValue.includes('|')){
                            _message += `${columnName}: value should be set as Listname|Fieldname for ${featureName} feature <br/><br/>`;  
                            haveError = true;
                        }
                        else {
                            var splitValue = controlValue.split('|');
                            var listname = splitValue[0];
                            var fieldname = splitValue[1];

                            if(splitValue.length > 2 || listname === '' || fieldname === ''){
                                _message += `${columnName}: value should be set as Listname|Fieldname for ${featureName} feature <br/><br/>`;  
                                haveError = true;
                            }
                            else {
                                var result = isListValidation(listname, fieldname)
                                            .then(function (result) {
                                                if (!result.isListExist) {
                                                    _message += `${columnName}: listname ${listname} is not found for ${featureName} feature <br/><br/>`;
                                                    haveError = true;
                                                } else if (!result.isFieldExist) {
                                                    _message += `${columnName}: Fieldname ${fieldname} is not found for ${featureName} feature <br/><br/>`;
                                                    haveError = true;
                                                }
                                            })
                                            .catch(function (error) {
                                            });
                                 validationPromises.push(result);
                            }
                        }
                      }
                      else if(featureName === 'textLength'){
                        debugger;
                        var lengthArray = [];

                        if(controlValue.includes(','))
                            lengthArray = controlValue.split(',');
                         else lengthArray.push(controlValue);

                         for (var i = 0; i < lengthArray.length; i++) {
                            var number = lengthArray[i];//, 10); // Parse the substring as an integer
                            if (isNaN(number)) { // Check if the parsing was successful
                                _message += `${columnName}: textLength should be integers only <br/><br/>`;  
                                haveError = true;
                            }
                          }
                      }
                    });
                }
            }
        }
   });
   await Promise.all(validationPromises);

   if(haveError)
    set_FNC_ErrorMessage(_message, false);
   return haveError;
}

function checkFeature(element){
 
    let featureName = element.nextSibling.innerText;
    var columnName = element.parentNode.parentNode.children;
    columnName = columnName[0].id;

    var enableFeatures = ['SpanField']; //, 'textLength', 'AllowedCharacters'];
    var disableFeatures = ['isList'];

    if(featureName === 'isList')
    {
      enableFeatures = [];
      disableFeatures = ['isText', 'textLength', 'AllowedCharacters'];
    }

    else if(featureName === 'isText')
    {
      enableFeatures = ['SpanField'];
      disableFeatures = ['isList', 'textLength', 'AllowedCharacters'];
    }

    else if(featureName === 'textLength' || featureName === 'AllowedCharacters')
    {
      enableFeatures = [];
      disableFeatures = ['isList', 'isText'];
    }

    const refId = '#' + element.getAttribute("refid");
    if (element.checked === true){
      
      $(refId).show();

      if(featureName === 'isOptionalField')
        checkColumnsFeatureSelection(true, refId, true);
      else if( featureName === 'isText' || featureName === 'isList' || featureName === 'SpanField' || featureName === 'textLength' || featureName === 'AllowedCharacters'){

        if( featureName === 'SpanField' && !isSpanExist_InSchema)
           isSpanExist_InSchema = true;
        Enable_Disable_CheckBox(false, enableFeatures, disableFeatures, columnName);
      }
    }
    else {
        $(refId).hide();
        if(featureName === 'isOptionalField')
          checkColumnsFeatureSelection(true, refId, false);
        else if( featureName === 'isText' || featureName === 'isList' || featureName === 'SpanField' || featureName === 'textLength' || featureName === 'AllowedCharacters'){
            if( featureName === 'SpanField')
              isSpanExist_InSchema = false;
          Enable_Disable_CheckBox(true, enableFeatures, disableFeatures, columnName);
        }
    }
}

var Enable_Disable_CheckBox = async function(isDisable, enablefeatures, disableFeatures, columnName){

    var _Row = $('#tblColumns').find('tr').filter(function() {
        return $(this).find('td:first').attr('id') === columnName;
    });
    //var nextRow = _Row.next();

    var td = _Row.find('td:not(:first-child)');

    var listCtlr, textCtlr;
    td.each(function(row, column){
       let featureName = column.innerText;
       if(enablefeatures.includes(featureName)){
        var inputControl = $(this).find('input');
        if(inputControl.length > 0){
            if(featureName === 'SpanField')
            {
                if(!isDisable){
                    if(isSpanExist_InSchema){
                        if (inputControl.is(':checked'))
                            inputControl.prop('disabled', false);
                        else inputControl.prop('disabled', true);
                        disableSpanFields(true);
                    }
                    else inputControl.prop('disabled', false);
                }
                else{
                    var isTextTd = td.filter(function() {
                        return $(this).find("span").text() === "isText"; 
                    });
                   var isTextInput = isTextTd[0].children[0];
               
                    if(isTextInput.checked && !isSpanExist_InSchema){
                      inputControl.prop('disabled', false);
                      disableSpanFields(false);
                    }
                    else {
                        inputControl.prop('disabled', true);
                        
                    }
                }
            }

            else{
                inputControl.prop('disabled', isDisable);
                if(isDisable === true && inputControl[0].checked){
                    inputControl.prop('checked', false);
                    var valueCtlrId = '#' + inputControl.attr("refid");
                    $(valueCtlrId).hide();
                }
            }
        }
       }
       else if(disableFeatures.includes(featureName)){
            var inputControl = $(this).find('input');
            if(inputControl.length > 0){
                if(isDisable === true)// && !inputControl.prop('checked'))
                inputControl.prop('disabled', false);
                else inputControl.prop('disabled', true);
            }
       }
    });
}

var setFilename_Schema  =async function(){
    var _fieldArray = [];
    var rowData  = {};
    var containsRev = false, isPropertySet = false;

    $('#tblColumns').find('tr').each(function(index){

        if(!isPropertySet){
           //#region ADD SCHEMA PROPERTIES
           rowData['InternalName'] = 'properties';
           rowData['DeliverableType'] = fd.field("DeliverableType").value.LookupValue;
           rowData['Acronym'] = fd.field("Acronym").value;
           rowData['MainList'] = fd.field("MainList").value;
           rowData['Delimeter'] = fd.field("Delimeter").value;
           _fieldArray.push(rowData);
           isPropertySet = true;
           //#endregion
        }

        if(index % 2 === 0)
        {
            rowData  = {};
            var td = $(this).children();
            var firstVisit = false;
            var internalName = '';

            td.each(function(row, column){
              if(!firstVisit){
                internalName = column.id;
                if(internalName === 'Revision' || internalName === 'Rev' )
                 containsRev = true;
                rowData['InternalName'] = internalName;
                firstVisit = true;
              }
              else{
                let featureName = column.innerText;
                var inputControl = $(this).find('input');
                if(inputControl.length > 0 && !inputControl.prop('disabled')){
                    var valueCtlrId, controlValue;
                    
                    if(featureName === 'isOptionalField'){
                        valueCtlrId = '#' + inputControl[0].id;
                        controlValue = $(valueCtlrId).prop("checked");
                        if(controlValue)
                        rowData[featureName] = true;
                    }
                    else {
                        valueCtlrId = '#' + inputControl.attr("refid")
                        controlValue = $(valueCtlrId).val();

                        if(controlValue !== null && controlValue !== undefined && controlValue !== ''){
                           if( featureName === 'SpanField')
                             rowData[featureName] = parseInt(controlValue);
                           else rowData[featureName] = controlValue.trim();
                        }
                   }
                }
              }
            });
            _fieldArray.push(rowData);
        }
    });
    return {
           _fieldArray,
           containsRev
        };
}

const getDeliverableTypes = async function(){
	let result = [];
      await pnp.sp.web.lists
			.getByTitle("FNC")
			.items
			.select("Title")
			//.filter(`Title eq '${  key  }'`)
			.get()
			.then((items) => {
				if(items.length > 0)
                  items.map(item => {
                    result.push(item.Title);
                  });
				  
				});
	 return result;
}

var disableSpanFields = async function(isDiable){
    $('#tblColumns').find('tr').each(function(index){
         if(index % 2 === 0)
         {
            var td = $(this).children();
            var spanFieldTd = td.filter(function() {
                return $(this).find("span").text() === "SpanField";
            });
           var spanFieldInput = spanFieldTd[0].children[0];

           if(isDiable){
                if(!spanFieldInput.checked){
                    spanFieldInput.disabled = true;
                }
           }
           else{

            var isTextTd = td.filter(function() {
                return $(this).find("span").text() === "isText"; 
            });
            var isTextInput = isTextTd[0].children[0];
       
            if(isTextInput.checked){
             spanFieldInput.disabled = false;
            }
         }
     }
   });
}

var isListValidation = async function(listname, listFieldName){
    var isListExist = false, isFieldExist = false;
    try {
        // Check if the list exists
        await pnp.sp.web.lists.getByTitle(listname).get();
        isListExist = true;

        // Check if the field exists
        var fields = await pnp.sp.web.lists.getByTitle(listname).fields
            .filter("InternalName eq '" + listFieldName + "'")
            .get();

        if (fields.length > 0) 
            isFieldExist = true;
        
    } catch (error) {
        if (error.status === 404) {
            console.log("The list '" + listname + "' does not exist.");
        } else {
            console.log("An error occurred: " + error);
        }
    }

    return {
        isListExist: isListExist,
        isFieldExist: isFieldExist
    };
}
//#endregion


