var _layout, _module = '', _formType = '', _acronym = '', _mainList = '', _phase ='';

var _Fields = [], _schema = [], selectedFieldsSchema = [];

var _fieldsInernalName = {}, _displayNames = [];
var _features =      ['SpanField', 'isList', 'RevStartWith', 'isText', 'textLength', 'AllowedCharacters', 'isOptionalField'];
var _featuresType =  ['number',    'string', 'string',       'string', 'string',     'string',            'boolean'];
var _featuresWidth = ['10%',       '16%',    '12%',          '16%',    '12%',        '12%',               '12%'];

var _isNew = false, _isEdit = false, _isDesign = false, _isWithRev = false;
var _isMain = true, _isPart = false, _isLead = false;


var isTblCreated = false, isSpanExist_InSchema = false, isOptionalExist_InSchema = false, disableIsList = false, disableIsText = false, _isMultiContracotr = false;

var listBoxFields = [];
var validationlistBoxPromises = [];
var _fNCLists = [];
//var _utilsPath;

var onRender = async function (relativeLayoutPath, moduleName, formType){
    _module = moduleName;
    _formType = formType;
    _layout = relativeLayoutPath;

    $(fd.field('Delimeter').$parent.$el).hide();
    $(fd.field('Acronym').$parent.$el).hide();
    $(fd.field('MainList').$parent.$el).hide();
    $(fd.field('Title').$parent.$el).hide();
    $(fd.field('Schema').$parent.$el).hide();

    await loadScripts();
    await setFormHeaderTitle();

    //_utilsPath = _layout + '/plumsail/js/utilities.js';

    // var script = document.createElement("script"); // create a script DOM node
    // script.src = _layout + "/plumsail/js/config/configFileRoutingFNC.js"; // set its src to the provided URL
    // document.head.appendChild(script);
    // await delay(500);

    let paramValue = await getParameter('FNCLists');
    if(paramValue !== null && paramValue !== undefined && paramValue !== '')
      _fNCLists = paramValue.split(',');

    await isMultiContractor(); //define var _isMultiContracotr;

    _phase = await getParameter('Phase');
    if(_phase.toLowerCase() === 'design')
      _isDesign = true;

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
        let _query = '';
        if(result !== undefined && result.length > 0){
            // if(_isDesign){
            //     result.map(item =>{
            //         $('div.k-list-scroller ul').find('li').each(function(index){
            //             var element = $(this);
            //             var _value = $(element).text().trim();
                    
            //             if(_value === item){
            //               $(element).css("pointer-events", "none").css("opacity", "0.6");
            //               $(element).prop('disabled', true);
            //             }
            //         });
            //     });
            // }

            //else{
                var isFirstVisit = true;
                _query = "IncludeModuleInFNC eq '1' ";
                result.map(item =>{
                    if(isFirstVisit){
                    _query += `and FNCModuleName ne '${item}' and FNCModuleName ne null `; //
                    isFirstVisit = false;
                    }
                    else _query += ` and FNCModuleName ne '${item}'`;
                });
                fd.field('DeliverableType').ready().then(() => {
                    fd.field('DeliverableType').filter = _query; //`Title ne 'DWG' and Title ne 'DOC'`;  
                    fd.field('DeliverableType').refresh();
                });
           // }
        }
        else{ //if(!_isDesign)
             _query = `FNCModuleName ne null and IncludeModuleInFNC eq '1'`;
             fd.field('DeliverableType').ready().then(() => {
                fd.field('DeliverableType').filter = _query; //`Title ne 'DWG' and Title ne 'DOC'`;  
                fd.field('DeliverableType').refresh();
            });
        }
    }
    else {
        
        var  schemaValue = fd.field("Schema").value.replace(/&nbsp;/g, '');
        _schema = JSON.parse(schemaValue);

        fd.field('DeliverableType').disabled = true;
        fd.field('Delimeter').disabled = true;

        _schema.filter(item => {
            if(item.InternalName !== 'properties'){
                let fieldTitle = item.Title;
                if(fieldTitle === undefined || fieldTitle === null)
                    fieldTitle = '';
                
                selectedFieldsSchema.push({'internalName': item.InternalName, 'Title': fieldTitle});

                if(item.isOptionalField === true)
                  isOptionalExist_InSchema = true;
                else if(item.SpanField !== undefined)
                  isSpanExist_InSchema = true;
                return true;
            }
         });

         let index = 0;
         for (let i = 0; i < selectedFieldsSchema.length; i++) {
            disableIsList = false; 
            disableIsText = false;

            let displayName = selectedFieldsSchema[i].Title;
            let internalName = selectedFieldsSchema[i].internalName;
            
            await setHTML_Column_PerRow(index, displayName, internalName, 'Add');
            index++;
        }
    }

    var lookupField = fd.field('DeliverableType');
    disableFields();

    await getDeliverableMetaInfo(lookupField);
    setButtons();

    //_spComponentLoader.loadScript(_utilsPath).then(setButtons);
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
        var result = '';
        // if(_isDesign)
        //   result = await getMajorType(value, false);
        // else 
        result = await getMajorType(value.LookupId, true);

        if(result.length > 0){
            _fieldsInernalName = {};
            _Fields = [];

          _acronym = result[0].Title;

          if(_isDesign)
            _mainList = result[0].LODListName;
          else {
            _mainList = result[0].MainList;
            _isWithRev = Boolean(result[0].isWithRev);
          }

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
    'SubmitToPM',
    'SubmitToSite',
    'SubmitToCont',
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
    'PartTrades',
    'DeliverableType',
    'Category'
  ];

  //#region EXCLUDE EXTRA COLUMNS PER MODULE
   if(_mainList !== 'RLOD')
    excludedFieldNames.push('Discipline');
//    if(_mainList !== 'RLOD' && _mainList !== 'Material Inspection Request')
//     excludedFieldNames.push('SubDiscipline');

    if(_mainList === 'Material Inspection Request')
        excludedFieldNames.push('CountryofOrigin', 'Manufacturer', 'MATReferenceBOQ', 'MaterialSubmittalNo', 'MATReferenceRecDate', 'MATReferenceRecDate', 
                                'MATReferenceSpecs', 'MATReferenceTitle', 'Quantity', 'SerialNo', 'STORAGELocation', 'ModelNo');
    else if(_mainList === 'Inspection Request')
        excludedFieldNames.push('Room','Drawingno','FromStation','ToStation');
    else if(_mainList === 'Inspection Request')
        excludedFieldNames.push('Room','Drawingno');
    else if(_mainList === 'Material Submittal')
        excludedFieldNames.push('Address', 'Room','Availability','BaseType', 'BOQ', 'Country', 'Drawingref', 'MatDate', 'ArrivalTime', 'Agent', 'Manufacturer', 
                                'Standards', 'TotalDuration', 'Specs');
    else if(_isDesign && _mainList === 'RLOD')
       excludedFieldNames.push('Approved', 'Comments', 'by', 'CExtensions', 'CFilename', 'Checked', 'CRevision', 'CScale', 'DAR', 'RejectedComments', 'Release',
                               'SubmittalRef', 'CPart', 'DarTrade', 'Description');
    //#endregion

if(!_isMultiContracotr)
 excludedFieldNames.push('Contractor');
else excludedFieldNames.push('Sender');

  await pnp.sp.web.lists.getByTitle(_mainList).fields.select("Title", "InternalName").orderBy("Title").get()
  .then(fields => {
  _Fields = fields.filter(field => {
        var internalName = field.InternalName;
        let fieldTitle = field.Title;
        var fieldType = field['odata.type'];
        if(!excludedFieldNames.includes(internalName) && !internalName.includes('_Part') && fieldType !== 'SP.FieldDateTime' && fieldType !== 'SP.FieldUrl' 
            && fieldType !== 'SP.FieldUser' && fieldType !== 'SP.Field')
        {
          //if(!_fieldsInernalName.includes(internalName)){
            if (!_fieldsInernalName.hasOwnProperty(fieldTitle)) {
                if(_isEdit && selectedFieldsSchema.find((field)=> field.internalName === internalName))
                    return false;
                
                //_fieldsInernalName.push({"internalName": internalName, "displayName": field.Title});
                _fieldsInernalName[fieldTitle] = internalName;
                _displayNames.push(fieldTitle);
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

function setListBox(){
    var fieldName = '';
    var isFieldVisited = false;

    $("#listBoxA").jqxListBox({ allowDrop: true, allowDrag: true, filterable: true, source: _displayNames, width: 300, height: 240, 
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
            let columnName = dragItem.label;
            let isFound = false;
                
            setTimeout(() => {
                let spanElements = $('#listBoxContentlistBoxB div span.jqx-listitem-state-normal');
                console.log(spanElements);

                let targetRowIndex;
                spanElements.map((rowIndex, spanElement) => {
                    let currentColumn = spanElement.textContent;
                    if (columnName === currentColumn) {
                        // if(dropItem.index === undefined || dropItem.index === null){
                        //     debugger;
                        //     dropItem.index = rowIndex;
                        // }
                        targetRowIndex = rowIndex;
                        console.log(`map rowIndex: ${rowIndex}, drop Index: ${dropItem.index}, drag Index: ${dragItem.index}, columnName: ${columnName}, Transaction: swap`);
                        isFound = true;
                    }
                });

                // /dropItem.element.parentElement

                let operation = 'swap';
                if(!isFound || dropItem.element.parentElement.id.includes('listBoxA')){
                  operation = 'remove';
                  if(targetRowIndex === undefined){
                    targetRowIndex = dragItem.index;
                  }

                //   if(isFound){
                //     let itemId = `#${dragItem.element.parentElement.id}`;
                //     $("#listBoxA").jqxListBox('addItem', columnName);
                //     $(itemId).remove();
                //     //listBox.jqxListBox('removeAt', targetRowIndex);
                //   }
                }

                else if(targetRowIndex === undefined){
                    if(dropItem !== null)
                      targetRowIndex = dropItem.index;
                    else dropItem = dragItem;
                }
    
                let internalName = _fieldsInernalName[columnName];
                if(internalName === undefined){
                 let getItem = selectedFieldsSchema.find((field)=> field.Title === columnName)
                 internalName = getItem.internalName;
                }
                setHTML_Column_PerRow(targetRowIndex, columnName, internalName, operation, dropItem);
               
            }, 100);
        },
        renderer: function (index, label, value, operation) {
            var lisboxBLength = $('#listBoxContentlistBoxB div div').length;

            let internalName = _fieldsInernalName[label];
            if(internalName === undefined){
                let getItem = selectedFieldsSchema.find((field)=> field.Title === label)
                internalName = getItem.internalName;
            }

             if(lisboxBLength > 0 && index === 0)
              listBoxFields = [];

              if(!listBoxFields.includes(internalName))
                listBoxFields.push(internalName);

                var _Row = $('#tblColumns').find('tr').filter(function() {
                    return $(this).find('td:first').attr('id') === internalName;
                });

                if(_Row.length === 0){
                //NOTE:EACH VALUE IT IS RENDERED TWICE UNTIL _Row.length BECAME > 0
                  if(!isFieldVisited){
                    isFieldVisited = true;
                    console.log(`Index: ${index}, columnName: ${label}, Transaction: add`);
                    
                    validationlistBoxPromises.push(setHTML_Column_PerRow(index, label, internalName, 'Add'));
                  }
                  else isFieldVisited = false;
                }
            return label;
        }
    });

    $(listBox).on('change', async function (event) {        
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

                  var currentRowIndex = $('#tblColumns').find('tr').index(_Row);
                  var nextRow = $(_Row).next();
                  var nextRowIndex = $('#tblColumns').find('tr').index(nextRow);

                //   console.log({'currentRowIndex': currentRowIndex, 
                //                'currentRow': _Row,
                //                'nextRowIndex': nextRowIndex, 
                //                'nextRow': nextRow});
                  tbl.append(_Row, nextRow);
              });
          }
        }
    });

    if(_isEdit){
        var dataLength = selectedFieldsSchema.length;
        for (let i = 0; i < dataLength; i++) {
            var value = selectedFieldsSchema[i].Title;
            listBox.jqxListBox('addItem', value);
        }
    }
}

var setHTML_Column_PerRow = async function(columnIndex, columnName, internalName, operation, dropItem){

    var contentElement = $('div.col-sm-12').eq(1);
    var tblId = 'tblColumns';
    var tblElement = $('#' + tblId);
    var colElement = $(`#${internalName}`);

    if(colElement.length > 0){
        if(operation === 'remove'){
            var closestTr = $(colElement).closest("tr");
            var nextTr = $(closestTr).next();

            $(nextTr).remove();
            $(closestTr).remove();
        }
        else if(operation === 'swap'){       
            // SWAP THE ITEM TO ITS POSITION
            tblElement.find('tr').each(function(index) {
                if(columnIndex !== null && index === columnIndex) //dropItem.index)
                {
                    var rowId = internalName + 'col' + columnIndex;    //columnName + 'col' + columnIndex;
                    var valueRowId = 'ctlr' + rowId;

                    var rowToMoveId = $('#' + rowId);
                    var rowToMove = rowToMoveId[0].parentElement.parentElement;
                    var rowToMoveIndex = $(rowToMove).index();

                    var rowValueToMoveId = $('#' + valueRowId);
                    var rowValueToMove = rowValueToMoveId[0].parentElement.parentElement;

                    var targetRow = 1;
                    if(index > 0)
                        targetRow = (index*2) + 1;
                    console.log(`targetRowIndex: ${targetRow}, row Index: ${index}`);

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

        if(operation === 'remove' || operation === 'swap'){
            if($(tblElement).find('tr').length === 0){
                $(tblElement).remove();
                isTblCreated = false;
            }
            return;
        }
    }

    if(!isTblCreated){
        isTblCreated = true;
        var tbl = "<table id='" + tblId + "' width='100%' class='customtbl'>";

        var ticksRow = await renderFeaturesTicks(columnName, internalName);
        var controlsRow= await renderFeaturesControls(internalName);

        tbl += ticksRow + controlsRow;
        tbl += "</table>";
        contentElement.after(tbl);
    }
    else{
        //append row into html table
        var ticksRow = await renderFeaturesTicks(columnName, internalName);
        var controlsRow= await renderFeaturesControls(internalName);
        const newRow = ticksRow + controlsRow;
        $(tblElement).append(newRow);
    }

  //$('#contentId').append(tbl);
  //$(tbl).insertAfter("#contentId");
}

var renderFeaturesTicks = async function(columnName, internalName){
    
    var row = "<tr>";
        row += "<td id='" + internalName + "' rowspan='2' width='10%' class='borderStyle'><label class='lblcolumnyStyle'>" + columnName + "</label></td>";

    for(var i =0; i < _features.length; i++){
        var checkBoxId = internalName + 'col' + i; //columnIndex is the rowIndex
        var controlId = 'ctlr' + internalName + 'col' + i;
        var featureName = _features[i];

        if(featureName === 'isOptionalField')
          row += "<td rowspan='2' width='" + _featuresWidth[i] + "' class='borderStyle'>";
        else row += "<td width='" + _featuresWidth[i] + "' class='borderStyle'>";
                    
             row += "<input type='checkbox' id='" + checkBoxId + "' onclick='checkFeature(this)' name='" + checkBoxId + "' class='ckStyle' refId='" + controlId + "'";

             row += await setFeatureCheckBoxes(featureName, columnName, internalName);

             row += "<span for='" + checkBoxId + "' class='lblpropertyStyle'>" + featureName + "</span>" +
         "</td>" ;
    }
    row += "</tr>";
    return row;
}

var setFeatureCheckBoxes = async function(featureName, columnName, internalName){
    var row = '';
    var result = false;

    if(columnName === 'Revision' || columnName === 'Rev'){
       if(featureName === 'SpanField' || featureName === 'isList' || featureName === 'isText' || featureName === 'isOptionalField')
          row += ' disabled>';
       else{ 
        if(!_isNew){

            result = await isFeatureFound(internalName, featureName);
            if(result.isFound)
              row += ' checked';
        }
         row += '>';
      }
    }
    else if(featureName === 'RevStartWith' || featureName === 'SpanField'){
        if(!_isNew){
            result = await isFeatureFound(internalName, featureName);
            if(result.isFound) 
                row += ' checked>';
            else row += ' disabled>';
        }
        else row += ' disabled>';
    }
    else {
        if(!_isNew){
            result = await isFeatureFound(internalName, featureName);
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

var renderFeaturesControls = async function(internalName){
    var row = "<tr>";
    let padding = 'padding: 5px;';
    for(var i =0; i < _features.length; i++){
        var controlId = 'ctlr' + internalName + 'col' + i;
        var columnType = _featuresType[i];
        var featureName = _features[i];

        if(columnType !== 'boolean'){
            row += "<td class='borderStyle'>";

             let isDDL = false;
             if(featureName === 'isList' && _fNCLists.length > 0)
                isDDL = true;

             if(isDDL){
              row += `<select id='${controlId}' style='width: 100% !important; ${padding}; border-radius: 4px; border: solid 1px;`;
             }
             else row += "<input type='" + columnType + "' id='" + controlId + "' style='width: 100% !important;";// display:none;'>";
            
                if(!_isNew){
                    var isFound = false;
                    var _controlValue = '';
                    _schema.map(item => {
                        if(item.InternalName === internalName && item.hasOwnProperty(featureName)){
                            isFound = true;
                            _controlValue = item[featureName];
                            return;
                        }
                    });

                    if(!isFound)
                      row += "display:none;'>";
                    else {
                        // if(isDDL)
                        //     row += "'>";
                        //else 
                        row += "' value='" + _controlValue + "'>";
                    }
                }
                else row += "display:none;'>";


                if(isDDL){
                    if(_isNew)
                      row += `<option style='${padding}' value='' selected></option>`;
                    else _controlValue = _controlValue.split('|')[0];
                    
                    for (const element of _fNCLists) {
                        row += `<option value='${element}'`;
                         if(!_isNew && element === _controlValue)
                          row += ' selected';
                        row += `>${element}</option>`;
                    }
                    row += "</select>";
                }
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

var isFeatureFound = async function(internalName, featureName){
    var isFound = false, disableFeature = false;
    _schema.map(item => {
        if(item.InternalName === internalName ){
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
    var haveError = false, containsRev = false;
    var _message = '';

    var validationPromises = [];
    
    $('#tblColumns').find('tr').each(function(index){
        if(index % 2 === 0)
        {
            var td = $(this).children();
            var columnName = td[0].id;

            if(columnName === 'Revision' || columnName === 'Rev')
              containsRev = true;

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
                        if(_fNCLists.length > 0){

                        }
                        else{
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
                      }
                      else if(featureName === 'textLength'){
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
   else if(!_isDesign && _isWithRev && !containsRev){
    set_FNC_ErrorMessage('Revision is Required field at the end of the filename', false);
    haveError = true;
   }
   
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
        //    if(_isDesign)
        //     rowData['DeliverableType'] = fd.field("DeliverableType").value;
        //    else 
           rowData['DeliverableType'] = fd.field("DeliverableType").value.LookupValue;
           rowData['Acronym'] = fd.field("Acronym").value;
           rowData['MainList'] = fd.field("MainList").value;
           rowData['Delimeter'] = fd.field("Delimeter").value;
           _fieldArray.push(rowData);
           isPropertySet = true;
           //#endregion
        }

        if(index % 2 === 0){
            rowData  = {};
            var td = $(this).children();
            var firstVisit = false;
            var internalName = '';

            td.each(function(row, column){
              if(!firstVisit){
                //internalName = column.id;
                internalName = _fieldsInernalName[column.innerText];
                if(internalName === undefined){
                    let getItem = selectedFieldsSchema.find((field)=> field.Title === column.innerText);
                    internalName = getItem.internalName;
                }
                //gfhfghfh
                if(internalName === 'Revision' || internalName === 'Rev' )
                 containsRev = true;
                rowData['InternalName'] = internalName;
                rowData['Title'] = column.innerText;
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
                           else if( featureName === 'isList' && _fNCLists.length > 0){
                             let mappingColumn = 'Title';
                             if(controlValue === 'Discipline' && _phase.toLowerCase() !== 'design')
                               mappingColumn = 'Acronym';

                             rowData[featureName] = `${controlValue.trim()}|${mappingColumn}`;
                           }
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
			.select("DeliverableType/FNCModuleName")
            .expand("DeliverableType")
			//.filter("Title ne 'GEN'")
			.get()
			.then((items) => {
				if(items.length > 0)
                  items.map(item => {
                   if(item.DeliverableType !== undefined)
                    result.push(item.DeliverableType.FNCModuleName);
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

var loadScripts = async function(){
    const libraryUrls = [
        _layout + '/controls/preloader/jquery.dim-background.min.js',
        _layout + '/plumsail/js/customMessages.js',
        _layout + '/plumsail/js/commonUtils.js',
        _layout + '/plumsail/js/utilities.js',
        _layout + '/plumsail/js/preloader.js',
        
        
        // _layout + '/controls/jqwidgets/scripts/jquery-1.11.1.min.js',
        _layout + '/controls/jqwidgets/scripts/demos.js',
        _layout + '/controls/jqwidgets/jqxcore.js',
        _layout + '/controls/jqwidgets/jqxlistbox.js',
        _layout + '/controls/jqwidgets/jqxscrollbar.js',
        _layout + '/controls/jqwidgets/jqxbuttons.js',
        _layout + '/controls/jqwidgets/jqxexpander.js',
        _layout + '/controls/jqwidgets/jqxvalidator.js',
        _layout + '/controls/jqwidgets/jqxinput.js',
        _layout + '/controls/jqwidgets/jqxdragdrop.js'
    ];
  
    const cacheBusting = `?v=${Date.now()}`;
      libraryUrls.map(url => { 
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
        });
        
    const stylesheetUrls = [
        _layout + '/plumsail/css/CssStyle.css',
        _layout + '/plumsail/css/FNC.css',
        _layout + '/controls/jqwidgets/styles/jqx.base.css'
    ];
  
    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
    });
}
//#endregion


var getInternalNameColumn_notused = async function(displayName){
    let field = _fieldsInernalName.find(field => {
      return field.displayName === displayName
    });
    return field.internalName;
}