var _schemaAcronyms = [];

var getCurrentWebUrl = async function() {
    return pnp.sp.web.get().then((web) => {
      return web.Url;
    });
}

const GetCurrentUser = async function(){
	try {
        const user = await pnp.sp.web.currentUser.get();
        return user;
    } catch (error) {
        console.error("Error fetching current user:", error);
        throw error; // Re-throw the error if needed
    }
}

var getDesign_GridSchema = async function(_web, webURL, key){
    var _colArray = [];
    var targetList;
    var targetFilter;

    await _web.lists
          .getByTitle("SchemaConfig")
          .items
          .select("ListName,hansonTblSchema,DataQuery")
          .filter("Title eq '" + key + "'")
          .get()
          .then(async function (items) {
              if(items.length > 0)
                //for (var i = 0; i < items.length; i++) {
                  var item = items[0];
                  var HandsonTblSchema = item.HandsonTblSchema;
                  targetList = item.ListName;
                  targetFilter = item.DataQuery;

                    var fetchUrl = webURL + HandsonTblSchema;
                    await fetch(fetchUrl)
                        .then(response => response.text())
                        .then(async data => {
                          _colArray = JSON.parse(data);
                    });
              });
    return {
      colArray: _colArray,
      targetList: targetList,
      targetFilter: targetFilter
    };
}

function getFileExtension(filePath) {
    const lastDotIndex = filePath.lastIndexOf('.');

    if (lastDotIndex !== -1 && lastDotIndex < filePath.length - 1) {
     return filePath.substring(lastDotIndex + 1).toLowerCase();
  }
    return null;
}

var addLegend = async function (lblId, _text, parentElement, step, ignoreCss){
	if(lblId !== undefined){
		jQueryId = '#' + parentElement;
		if(step === 'same'){
		  $(jQueryId)
		  .html(_text)
		  .addClass('FormTitle');
		}

		else{
		  jQueryId = '#' + lblId;
		  if ($(jQueryId).length === 0) {
			var label = $('<label>', {
			  id: lblId,
			  html: _text
			});
			//label.css('color', '#3cdbc0');
			if(!ignoreCss){
				label.css('color', '#4e778f');
				label.css('font-weight', 'bold');
				label.css('font-size', '16px');
				label.css('width', '1350px');
			}
			//label.css('text-decoration', 'underline');
			//label.css('text-align', 'center');

			if(step === 'after')
			  $(parentElement).after(label);
			else $(parentElement).before(label);
		  }
		}
	  }

	else if( (_isMain &&  _formType === 'New') || (_isPart && _formType === 'Edit') ){
		var id = '#legendlbl';
		if ($(id).length === 0) {

		var textNote = `Click 'Save' to keep your work as a draft, or 'Submit' to officially submit your answer.`;

		if(_module === 'SLF' && _isTeamLeader)
			textNote = `Click 'Submit' to officially submit your answer.`;

		var titleElement = $('div.col-sm-12')[0]; //$("button:contains('Cancel')");
		var label = $('<label>', {
			id: 'legendlbl',
			html: textNote
		});
		label.css('color', 'red');
		label.css('font-weight', 'bold');
		label.css('font-size', '16px');
		label.css('width', '1250px');
		//label.css('text-align', 'center');
		$(titleElement).before(label);
		}
	}

	if($('p').find('small').length > 0)
		$('p').find('small').remove();
}

var setFormHeaderTitle = async function(){
	let moduleTitle = '';
	  if(_isPart)
       addLegend('partlblId', partTaskHeaderTitle, 'partTitle', 'same');
      else if(_isLead)
       addLegend('leadlblId', leadTaskHeaderTitle, 'leadTitle', 'same');
	  else if(_isMain){
		if(_module === 'SCR')
		  moduleTitle = scrHeaderTitle;
		else if(_module === 'IR')
		  moduleTitle  = irHeaderTitle;
		else if(_module === 'MIR')
		  moduleTitle  = mirHeaderTitle;
		else if(_module === 'MAT')
		  moduleTitle  = matHeaderTitle;
		else if(_module === 'SI')
		  moduleTitle  = siHeaderTitle;
		else if(_module === 'SLF')
		  moduleTitle  = slfHeaderTitle;
		else if(_module === 'AUR')
		  moduleTitle  = aurHeaderTitle;
		else if(_module === 'DCC')
		  moduleTitle  = dccHeaderTitle;
		else if(_module === 'LOD')
		  moduleTitle  = lodHeaderTitle;
		else if(_module === 'DTRD')
		  moduleTitle  = dtrdHeaderTitle;
		else if(_module === 'DPR')
		  moduleTitle  = dprHeaderTitle;
		else if(_module === 'FNC')
		  moduleTitle  = fncHeaderTitle;
		else if(_module === 'CKD')
		  moduleTitle  = ckdHeaderTitle;
		else if(_module === 'MAP')
		  moduleTitle  = mapHeaderTitle;
		else if(_module === 'AUS')
		  moduleTitle  = ausHeaderTitle;
		else if(_module === 'INS')
		  moduleTitle  = insHeaderTitle;
		else if(_module === 'PINT') //FOR PROJECT CENTER
			moduleTitle  = pintHeaderTitle;
		else if(_module === 'MTD') //FOR PROJECT CENTER
			moduleTitle  = mtdHeaderTitle;
     }
	 await addLegend('modulelblId', moduleTitle, 'moduleTitle', 'same');
}

function clearStoragedFields(fields, execludeField){

	for (const field in fields) {

		if(execludeField !== undefined && field === execludeField)
			continue;

		var fieldproperties = fd.field(field);
		if(fieldproperties._fieldCtx.schema !== undefined){
			var fieldDefaultVal = fieldproperties._fieldCtx.schema.DefaultValue;
			//var fieldType = fieldproperties._fieldCtx.schema.FieldType;

			if (fieldDefaultVal !== undefined && fieldDefaultVal !== null) {}
			else
				fd.field(field).clear();
		}
	}
}

var getParameter = async function(key){
	let result = "";
      await pnp.sp.web.lists
			.getByTitle("Parameters")
			.items
			.select("Title,Value")
			.filter(`Title eq '${  key  }'`)
			.get()
			.then((items) => {
				if(items.length > 0)
				  result = items[0].Value;
				});
	 return result;
}

var getCounter = async function(web, type, doUpdate){
	let listname = 'Counter';
	let query = `Title eq '${type}'`;

	return await web.lists.getByTitle(listname).items.select("Id, Title, Counter").filter(query).get()
	.then(async items=>{
		let counterValue = 1;
		if(items.length === 0){
			if(doUpdate){
				await web.lists.getByTitle(listname).items.add({
					Title: type,
					Counter: counterValue.toString()
				});
			}
		}
		else{
			let item = items[0];
			counterValue = parseInt(item.Counter) + 1;
			if(doUpdate){
				await pnp.sp.web.lists.getByTitle(listname).items.getById(item.Id).update({
					Counter: counterValue.toString()
				});
			}
		}
		return counterValue;
	})
}

var isMultiContractor = async function(){

	_isMultiContracotr = await getParameter("isMultiContracotr");
	if(_isMultiContracotr.toLowerCase() === 'yes'){
		_isMultiContracotr = true;
		let result = '';

		const queryString = window.location.search;
		var urlParams = new URLSearchParams(queryString);
		var folderUrl;
		if(_formType == "New")
		 folderUrl = urlParams.get('rf');
		else {
			folderUrl = urlParams.get('source');
			urlParams = new URL(folderUrl);
			const searchParams = urlParams.searchParams;
			folderUrl = searchParams.get('RootFolder');
		}

		if(folderUrl !== null){
			const pathSegments = folderUrl.split("/");
			var folderName = pathSegments[pathSegments.length - 1];

			await pnp.sp.web.lists
				.getByTitle("Contractor")
				.items
				.select("AllowedUsers/Title")
				.expand('AllowedUsers')
				.filter(`Title eq '${  folderName  }'`)
				.get()
				.then((items) => {
					if(items.length > 0)
						result = items[0]['AllowedUsers'];
						console.log('allowedUsers = ' + result);
						result = result[0].Title;
				});
	   }
	   return result;
	}
	else _isMultiContracotr = false;
}

var IsUserInGroup = async function(group){
  let IsPMUser = false
  try{
       await pnp.sp.web.currentUser.get()
           .then(async (user) =>{
        await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
               .then(async (groupsData) =>{
                  for (let i = 0; i < groupsData.length; i++) {
                     //alert("groupsData[i].Title = " + groupsData[i].Title + "      group = " + group);
                      //if(groupsData[i].Title == group.Title)
                      if(groupsData[i].Title == group)
                      {
                             IsPMUser = true;
                             return IsPMUser;
                      }
                  }
              });
           });
  }
  catch(e){alert(e);}
  return IsPMUser;
}

var getMajorType = async function(key, byId){
	  let result = '';
	  //let listname = 'SchemaConfig';
      var query = `Title eq '${  key  }'`;

     if(byId !== undefined && byId === true)
	    query = `ID eq ${  key  }`;

	 listname = 'MajorTypes';


      await pnp.sp.web.lists
			.getByTitle(listname)
			.items
			//.select("Title,Value")
			.filter(query)
			.get()
			.then((items) => {
				if(items.length > 0)
				  result = items;
				});
	 return result;
}

var getGridMType = async function(_web, webURL, key, isDesign){
    //var _itemArray = {};
    var _colArray = [];
    var targetList, targetFilter, revNumStart, updateTitle;
    let listname = 'MajorTypes';
	let columns = 'HandsonTblSchema,LODListName,LODFilterColumn,RevNumStart,UpdateTitle';
	if(_module === 'AUS')
	  columns = 'HandsonTblSchema,ListName,FilterColumns';

	// if(isDesign === true){
	//    listname = 'SchemaConfig';
	//    columns = 'HandsonTblSchema,ListName,LODFilterColumn,RevNumStart,UpdateTitle';
	// }

    var result = "";
    await _web.lists
          .getByTitle(listname)
          .items
          .select(columns)
          .filter("Title eq '" + key + "'")
          .get()
          .then(async function (items) {
              if(items.length > 0)
                //for (var i = 0; i < items.length; i++) {
                  var item = items[0];
                  var hansonTblSchema = item.HandsonTblSchema;

				  if(_module === 'AUS'){
					targetList = item.ListName;
					targetFilter = item.FilterColumns;
				  }
				  else{
					targetList = item.LODListName;
					targetFilter = item.LODFilterColumn;
					revNumStart = item.RevNumStart;
					updateTitle = item.UpdateTitle;

				  }

                    var fetchUrl = webURL + hansonTblSchema;
                    await fetch(fetchUrl)
                        .then(response => response.text())
                        .then(async data => {
                          _colArray = JSON.parse(data);

                           for (const obj of _colArray) {
                            if (obj.renderer === "customDropdownRenderer"){
                              obj.renderer = customDropdownRenderer;
                            }
                        if (obj.source === "getDropDownListValues"){
                              obj.source = await getDropDownListValues(obj.listname, obj.listColumn);
                            }
                        else if (obj.source === "getQMDropDownListValues"){
							obj.source = await getQMDropDownListValues(obj.listname, obj.listColumn);
						  }

                        //     else if (obj.validator === "validateDateRequired"){
                        //       obj.validator = validateDateRequired;
                        //     }
                        //     else if (obj.validator === "getDescriptionValidator_PLF"){
                        //       obj.validator = getDescriptionValidator_PLF;
                        //     }
                        //     else if (obj.validator === "preventEdit"){
                        //       obj.validator = preventEdit;
                        //     }
                           }
                    });
              });
   return {
    colArray: _colArray,
    targetList: targetList,
    targetFilter: targetFilter,
	revNumStart: revNumStart,
	updateTitle: updateTitle
  };
}

var getDropDownListValues = async function(_listname, _column){
	var _colArray = [];
    let filter = `${_column} ne null`

	await pnp.sp.web.lists
		  .getByTitle(_listname)
		  .items
		  .select(_column)
		  .filter(filter)
		  .get()
		  .then(async function (items) {
			  if(items.length > 0)
				for (var i = 0; i < items.length; i++) {
			         let val = items[i][_column];
					 if(!_colArray.includes(val))
					   _colArray.push(val);
				}
			  });
   return _colArray;
}

function fixTextArea (){
	$("textarea").each(function(index){
		var height = (this.scrollHeight + 5) + "px";
        $(this).css('height', height);
	});
}

function fixEnhancedRichText (){
	var isFound = false;

	var iframes = $('iframe.k-content');
    iframes.each(function(index) {
		var iframe = this;
		if(_module === 'AUR'){
            if(activeTabName === auditReportTab && index === 0)
                setIframeHeight(iframe);
			else if( (activeTabName === corrActionTab || _isAurClosed) && index > 0)
			   setIframeHeight(iframe);
		}
		else setIframeHeight(iframe);
    });

	if(isFound)
	clearInterval(_timeOut);
}

function setIframeHeight(iframe){
	if (iframe && iframe.contentWindow && iframe.contentWindow.document) {
		var iframeScrollHeight = iframe.contentWindow.document.body.scrollHeight;
		var totalHeight = (iframeScrollHeight + 20) + 'px' ;
		iframe.style.height = totalHeight;
	}
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

                    if(getParams){
					  var value= xml.find("GLOBAL_PARAMResult");
					  if(value.length > 0){
						text = value.text();
						_layout = value[0].children[0].textContent;
					  }
					}

                    else{
						if(_module === 'AUR' && activeTabName === auditReportTab){
							var value= xml.find("GET_AUR_SUMMARYResult");
							var text;

							if(value.length > 0){
							text = value.text();
							fd.field('Summary').value = text;
							_header = value[0].children[0].textContent;
							_main = value[0].children[1].textContent;
							_footer = value[0].children[2].textContent;
							}
						}
						else{
							var value= xml.find("GET_EMAIL_BODYResult");
							if(value.length > 0){
								if(_module === 'AUR' && activeTabName === corrActionTab)
								fd.field('CARSummary').value = value.text();
							}
						}
						fixEnhancedRichText();
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

var getGlobalParameters = async function(){
	preloader();
    var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GLOBAL_PARAM';
    var soapContent;
    soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body>' +
                      '<GLOBAL_PARAM xmlns="http://tempuri.org/" />' +
                    '</soap:Body>' +
                  '</soap:Envelope>';
     var result = await getSoapRequest('POST', serviceUrl, false, soapContent, true);
	 console.log(result);
}

var setButtonCustomToolTip = async function(_btnText, toolTipMessage, animationType){
	animationType = animationType === undefined ? 'fade' : animationType
    var btnElement = $('span').filter(function(){ return $(this).text() == _btnText; }).prev();
	if(btnElement.length === 0)
	  btnElement = $(`button:contains('${_btnText}')`);

    if(btnElement.length > 0){
	  if(btnElement.length > 1)
		btnElement = btnElement[1].parentElement;
      else btnElement = btnElement[0].parentElement;

      $(btnElement).attr('title', toolTipMessage,);

	  await _spComponentLoader.loadScript( _layout + '/controls/tooltipster/jquery.tooltipster.min.js').then(async () =>{
		$(btnElement).tooltipster({
			delay: 100,
			maxWidth: 350,
			speed: 500,
			interactive: true,
			animation: animationType, //fade, grow, swing, slide, fall
			trigger: 'hover'
		  });
	  });
    }
}

//#region FNC VALIDATION
var checkFileName = async function(acronym, delimeter, schema, filenameText, filename, isSingle, checkRev){

    var selectedSchema = '', selectedDelimeter = '', _mesg = '';
	var selectedType = [];

     _schemaAcronyms.map(item => {
        if(item.Acronym === acronym){
		   selectedType = item;
		   return true;
		}
     });

    if(selectedType.length === 0){
		selectedSchema = schema;
        selectedDelimeter = delimeter;
        var rowData  = {};
        rowData['Acronym'] = acronym;
        rowData['Delimeter'] = delimeter;
        rowData['Schema'] = schema;
        rowData['FNC'] = filenameText;
        _schemaAcronyms.push(rowData);
     }
     else{
        selectedDelimeter = selectedType['Delimeter'];
        selectedSchema = selectedType['Schema'];
        filenameText = selectedType['FNC'];
     }

      var schemaParts = filenameText.split(delimeter);

	  var schemaPartsLength = schemaParts.length;
      var filenameParts = filename.toUpperCase().split(delimeter);
      var fielnameLength = filenameParts.length;

	  var resultFields = await getSpan_Optional_Field(selectedSchema);
      var spanFieldValue = resultFields.spanField;
	  var isOptionalField = resultFields.isOptionalFieldValue;
	  var removeRevLength = resultFields.removeRevLength;
	  var ignoreOptionalField = true;

	  if(spanFieldValue !== undefined)
	   fielnameLength = fielnameLength-(spanFieldValue-1);

	  //if(removeRevLength)
	  schemaParts.map(item => {
		var fieldname = item.InternalName;
        if(!checkRev && (fieldname === 'Rev' || fieldname === 'Revision'))
		   schemaPartsLength -= 1;
     });


	  if(isOptionalField !== undefined)
	  {
		if(fielnameLength === schemaPartsLength)
		  ignoreOptionalField = false;
	  }

      if(schemaPartsLength !== fielnameLength){

        if(isOptionalField !== undefined && fielnameLength === (schemaPartsLength -1)){}
        else{
			_mesg += fncLengthMesg;

			if(spanFieldValue !== undefined)
			  _mesg += ' (Note: SpandField is Enabled) and delimeter is ' + selectedDelimeter;
		}
      }

	  if(_mesg === '')
		_mesg += await validateFileName(selectedSchema, selectedDelimeter, filenameParts, ignoreOptionalField, checkRev);

      return _mesg
    //var _features =      ['SpanField', 'isList', 'RevStartWith', 'isText', 'textLength', 'AllowedCharacters', 'isOptionalField'];
}

var getSpan_Optional_Field = async function(schema, checkRev){
	var spanFieldValue, isOptionalFieldValue, removeRevLength = false;;
	schema.map(item => {
		var fieldname = item.InternalName;

		if(item.SpanField !== undefined)
		 spanFieldValue = item.SpanField;

		if(item.isOptionalField === true)
		 isOptionalFieldValue = fieldname;

		if(!checkRev && (fieldname === 'Rev' || fieldname === 'Revision'))
		 removeRevLength = true;
	});

	return {
		spanField: spanFieldValue,
		isOptionalFieldValue: isOptionalFieldValue,
		removeRevLength: removeRevLength
	};
}

var isAllowedCharacters = async function(value, allowedCharacters){
	const filteredValue = value.split('').filter(char => {
		return !allowedCharacters.includes(char)
	});
	return filteredValue.length === 0;
}

var validateFileName = async function(schema, delimeter, filenameParts, ignoreOptionalField, checkRev){
	var position = 0;
	var errorMesg = '';
	for (const item of schema) {
		var fieldname = item.InternalName;
		var filenamePartValue = filenameParts[position];

		if((fieldname === 'Rev' || fieldname === 'Revision'))
		{
			if(!checkRev)
			  continue;

			if(item.RevStartWith !== undefined){
				if (!filenamePartValue.startsWith(item.RevStartWith))
				  errorMesg += `revision must start with ${item.RevStartWith}<br>`; //for ${fieldname} field
			}

			if(_module === 'LOD'){
				if(_revNumStart !== undefined){
					if(isAllDigits(_revNumStart))
					 _revNumStart = parseInt(_revNumStart);

					 if(_revNumStart != filenamePartValue.replace(item.RevStartWith,''))
						errorMesg += `revision number must be ${_revNumStart}<br>`;
				}
		    }
	   }

		if(ignoreOptionalField && item.isOptionalField){
		  continue;
		}

		if(item.SpanField !== undefined){
			var getfilenameValue = '';
			var spanFieldValue = item.SpanField;

			if(spanFieldValue === 2)
			getfilenameValue = filenameParts[position] + delimeter + filenameParts[position+1];
			else if(spanFieldValue === 3)
			  getfilenameValue = filenameParts[position] + delimeter + filenameParts[position+1] + delimeter + filenameParts[position+2];
			else if(spanFieldValue === 4)
				getfilenameValue = filenameParts[position] + delimeter + filenameParts[position+1] + delimeter + filenameParts[position+2] + delimeter + filenameParts[position+3];
			var isText = item.isText;
			var dictArray = [];

			if(isText.includes(',')){
			  dictArray = isText.split(',');

			  var isValid = false;
			  for (var i = 0; i < dictArray.length; i++) {
					var _value = dictArray[i]; // Parse the substring as an integer
					if (getfilenameValue === _value){
						isValid = true;
						break;
					}
			  }
			  if(!isValid)
				errorMesg += `${fieldname} must be a text field having ${isText.replace(',', ' Or ')}<br>`;
			}
			else{
				if(getfilenameValue !== isText){
					errorMesg += `text must be ${isText} for ${fieldname} (SpanField)<br>`;
					return errorMesg;
				}
			}
		}

		else if(item.isList !== undefined){
          var splitValue = item.isList.split('|');
		  var listname = splitValue[0];
		  var listField = splitValue[1];

          var descField = 'Title';
		  var descValue;
		  if(listname === 'SubDiscipline')
		    descField = 'Acronym';
		  var viewFields = listField + ',' + descField;

		  var isFound = false;
		  var itemId;
		  var listValuesMesg = '';
		  var items =  await pnp.sp.web.lists
			.getByTitle(listname)
			.items
			.select('Id,' + viewFields)
			//.filter(`Title eq '${  listValue  }'`)
			.getAll();

			if(items < 1){
				errorMesg += `${listname} list is Empty. Contact the admin.<br>`;
				position++;
				continue;
			}

			if(_module === 'LOD'){
				let listitem = localStorage.getItem(listname);
				if(listitem !== undefined && listitem !== null && listitem !== ''){

				}
				else{
					items.map(item => {
						var itemValue = item[listField];
						if (!_listDictionary.includes(itemValue))
						 _listDictionary.push(itemValue);

						if(itemValue === filenamePartValue){
						itemId = item.Id;
						descValue = item[descField];
						isFound = true;
						return;
						}
						else{
							if(rowValue < 4){
								listValuesMesg += itemValue + ',';
								rowValue++;
							}
						}
					});
					localStorage.setItem(listname, listname);
				}
				let item = _listDictionary.find(item => item === filenamePartValue);
				if(item !== undefined && item !== null && item !== '')
				   isFound = true;
			}

			if(items.length > 0){
				var rowValue = 0;
				items.map(item => {
					var itemValue = item[listField];
					if(itemValue === filenamePartValue){
					 itemId = item.Id;
					 descValue = item[descField];
					 isFound = true;
					 return;
					}
					else{
						if(rowValue < 4){
							listValuesMesg += itemValue + ',';
							rowValue++;
						}
					}
				});

				if(_isMultiContracotr && listname === 'Contractor'){
					var queryString = window.location.search;
					var urlParams = new URLSearchParams(queryString);
					var FolderName = urlParams.get('rf');
                    let ignoreCheck = false;
					if(_isEdit){
                        if(FolderName !== undefined && FolderName !== null && FolderName !== '' && FolderName.includes('/')){
							//IF NOT SITE ADMIN IT COMES HERE
							FolderName = FolderName.substring(FolderName.lastIndexOf('/')+1);
							if(filenamePartValue.toUpperCase() !== FolderName.toUpperCase())
                                errorMesg += `Contractor name must be ${FolderName}<br>`;
							ignoreCheck = true;;
						}
						else{
							folderUrl = urlParams.get('source');
							urlParams = new URL(folderUrl);
							const searchParams = urlParams.searchParams;
							FolderName = searchParams.get('RootFolder');
						}
					}

					if(!ignoreCheck){
						if(FolderName !== null){
							FolderName = FolderName.split("/")[5];
							if(filenamePartValue.toUpperCase() !== FolderName.toUpperCase())
							errorMesg += `Contractor name must be ${FolderName}<br>`;
						}
						else if(!isFound)
						errorMesg += `${fieldname} must be one of the following: ${listValuesMesg} etc..<br>`;
						else errorMesg += `Select a contractor name folder to submit your deliverables<br>`;
					}
				}
				 else{
					if(!isFound)
					  errorMesg += `${fieldname} must be one of the following: ${listValuesMesg} etc..<br>`;
					else{
						try{
							if(fieldname === 'Contractor')
							  fd.field(fieldname).value = descValue;
							else{
								var obj = {  LookupId: itemId,
											LookupValue: descValue
										};

							   if(_module !== 'LOD'){

									if(fieldname === 'Discipline_x003a_Acronym')
									  fieldname = 'Discipline';
									else if(fieldname === 'SubDiscipline_x003a_Title')
									  fieldname = 'SubDiscipline';

									fd.field(fieldname).value = obj;
									fd.field(fieldname).disabled = true;
							   }
							}
						}
						catch{}
					}
				}
			}
		}

		else if(item.isText !== undefined){
			var isTextValue = item.isText.toUpperCase();
			var lengthArray = [];

			if(isTextValue.includes(','))
				lengthArray = isTextValue.split(',').map(item => item.trim()).filter(item => item !== '');
			else lengthArray.push(isTextValue);

			var isValid = false;
			for (var i = 0; i < lengthArray.length; i++){
				var value = lengthArray[i]; // Parse the substring as an integer
				if (filenamePartValue.toUpperCase() === value){
					isValid = true;
					break;
				}
			}
			if(!isValid)
			   errorMesg += `text must be ${isTextValue.replace(',', ' Or ')} instead of ${filenamePartValue} for ${fieldname} <br>`;
		}

		else if(item.textLength !== undefined){

           if(item.RevStartWith !== undefined)
		     filenamePartValue = filenamePartValue.replace(item.RevStartWith, '');

			var textLength = item.textLength;
			var lengthArray = [];

			if(textLength.includes(','))
			  lengthArray = textLength.split(',');
			else lengthArray.push(textLength);

			var isValid = false;
			for (var i = 0; i < lengthArray.length; i++) {
				var number = parseInt(lengthArray[i], 10); // Parse the substring as an integer
				if (filenamePartValue.length === number){
					isValid = true;
					break;
				}
			}
			if(!isValid)
			 errorMesg += `${fieldname} must be a text field having ${textLength.replace(',', ' Or ')} characters<br>`;
            else{
			 var AllowedCharacters = item.AllowedCharacters;
				if(AllowedCharacters !== undefined){
					var isAllowed = await isAllowedCharacters(filenamePartValue, AllowedCharacters);
						if (!isAllowed)
						errorMesg += `${fieldname} must be a text field with characters from ${AllowedCharacters}<br>`;
				}
			}
		}

		else if(item.AllowedCharacters !== undefined){
			var AllowedCharacters = item.AllowedCharacters;
			var isAllowed = await isAllowedCharacters(filenamePartValue, AllowedCharacters);
			if (!isAllowed)
				errorMesg += `${fieldname} must be a text field with characters from ${AllowedCharacters}<br>`;
		}

		if(fieldname !== 'properties'){
		 if(item.SpanField !== undefined)
		  position = position + item.SpanField;
		 else position++;
		}
	}
	return errorMesg;
}
//#endregion


const _sendEmail = async function(ModuleName, emailName, query, ApprovalTradeCC, notificationName, rootFolder, currUser){
	let webUrl = _spPageContextInfo.siteAbsoluteUrl;
	let siteUrl = new URL(webUrl).origin;
    let CurrentUser;

	debugger;
	if(!isNullOrEmpty(currUser))
		CurrentUser = currUser
	else CurrentUser = await pnp.sp.web.currentUser.get(); //await GetCurrentUser();


    let serviceUrl = `${siteUrl}/AjaxService/DarPSUtils.asmx?op=SEND_EMAIL_TEMPLATE`;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                        <soap:Body>
                            <SEND_EMAIL_TEMPLATE xmlns="http://tempuri.org/">
                                <WebURL>${webUrl}</WebURL>
                                <Email_Name>${emailName}</Email_Name>
                                <Query><![CDATA[${query}]]></Query>
                                <UserDisplayName>${CurrentUser.Title}</UserDisplayName>
                                <CurrentUserEmail>${CurrentUser.Email}</CurrentUserEmail>
                                <ApprovalCC>${ApprovalTradeCC}</ApprovalCC>
                                <CheckPageRootFolder>${rootFolder}</CheckPageRootFolder>
                                <ModuleName>${ModuleName}</ModuleName>
                                <Notification_Name>${notificationName}</Notification_Name>
                            </SEND_EMAIL_TEMPLATE>
                        </soap:Body>
                    </soap:Envelope>`
    let res = await getSoapResponse('POST', serviceUrl, true, soapContent, 'SEND_EMAIL_TEMPLATEResult');
}

const _generateErrorEmail = async function (webUrl, username, displayName, errorMessage, errorStack) {

    const usernameHtml = username ? `<p><strong>Username:</strong> ${username}</p>` : '';
    const displayNameHtml = displayName ? `<p><strong>Display Name:</strong> ${displayName}</p>` : '';

    let errroBody = `
        <html>
        <head>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    background-color: #f4f4f4;
                    margin: 0;
                    padding: 20px;
                    font-size: 15px; /* Set default font size */
                }
                .container {
                    max-width: 600px;
                    margin: 0 auto;
                    background-color: #ffffff;
                    padding: 20px;
                    box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                }
                .header {
                    background-color: #ff6f61;
                    padding: 10px 20px;
                    color: #ffffff;
                    text-align: center;
                    font-size: 17px; /* Set font size for header */
                }
                .content {
                    margin: 20px 0;
                }
                .content p {
                    font-size: 13px; /* Set font size for content paragraphs */
                }
                .footer {
                    text-align: center;
                    color: #999999;
                    margin-top: 20px;
                    font-size: 11px; /* Set font size for footer */
                }
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>Error Notification</h1>
                </div>
                <div class="content">
                    <p>Dear Admin,</p>
                    <p>An error has occurred in the system. Below are the details:</p>
                    ${usernameHtml}
                    ${displayNameHtml}
					<p><strong>Site URL:</strong> ${webUrl}</p>
                    <p><strong>Error Message:</strong> ${errorMessage}</p>
                    <p><strong>Error Detail:</strong> ${errorStack}</p>
                    <p>Please investigate the issue as soon as possible.</p>
                </div>
                <div class="footer">
                    <p>This is an automated message. Please do not reply to this email.</p>
                </div>
            </div>
        </body>
        </html>
    `;

    errroBody = htmlEncode(errroBody);
    await _sendErrorEmail(errroBody);
}

const _sendErrorEmail = async function(Body){

	let webUrl = _spPageContextInfo.siteAbsoluteUrl;
	let siteUrl = new URL(webUrl).origin;

    let serviceUrl = `${siteUrl}/AjaxService/DarPSUtils.asmx?op=SEND_ERROR_EMAIL_CVGUIDELINE`;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                        <soap:Body>
                            <SEND_ERROR_EMAIL_CVGUIDELINE xmlns="http://tempuri.org/">
                                <WebURL>${webUrl}</WebURL>
                                <Body>${Body}</Body>
                            </SEND_ERROR_EMAIL_CVGUIDELINE>
                        </soap:Body>
                    </soap:Envelope>`
    let res = await getSoapResponse('POST', serviceUrl, true, soapContent, 'SEND_ERROR_EMAIL_CVGUIDELINEResult');
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
                        //_rootSite = value[0].children[1].textContent;
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
	if(soapContent !== '')
      xhr.send(soapContent);
	else xhr.send();
}

var getSoapResponse1 = async function(method, serviceUrl, isAsync, soapContent, getResultTag){
	return new Promise((resolve, reject) => {
        var xhr = new XMLHttpRequest();
        xhr.open(method, serviceUrl, isAsync);
        xhr.onreadystatechange = function() {
            if (xhr.readyState == 4) {
                try {
                    if (xhr.status == 200 && getResultTag !== '') {
                        const response = this.responseText;
                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(response, "text/xml");
                        resolve(xmlDoc); // Resolve the promise with the parsed XML document
                    } else {
                        reject(new Error('Failed to get valid response'));
                    }
                } catch (err) {
                    reject(err);
                }
            }
        };
        xhr.setRequestHeader('Content-Type', 'text/xml');
        if (soapContent !== '')
		  xhr.send(soapContent);
        else xhr.send();
    });
}

function setPSErrorMesg(errMesg, removeAlertHeadingText){


	function checkForErrors() {

	  var errorElement = $('.errors[data-v-386d995a]');
	  if (errorElement.length > 0) {
		  clearInterval(intervalId);
		  handleErrors(errorElement, errMesg);
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

	function handleErrors(element, errorMessage) {
	  if (errorMessage) {
		  element.css({
			  'height': 'auto',
			  'opacity': '1'
		  });
		  if(removeAlertHeadingText === true){
			$('p.alert-heading').text('').css({'display': 'none'});
			errorMessage = errorMessage;
		  }
		  else errorMessage = '<br/>' + errorMessage;
		  var mesgElement = $('#customErrorId');
		   if(mesgElement.length === 0){
			// if(module === 'PMG')
			// 	$('p.alert-heading').text('Current Status')
			$('p.alert-heading').css({'display': 'inline'}).append(`<p id='customErrorId'>${errorMessage}</p>`);
		   }
		   else mesgElement.html(errorMessage);
		  //$('span').filter(function () { return $(this).text() == buttonText; }).parent().css('color', '#737373').attr("disabled", "disabled");
	  } else {
		  element.css({
			  'height': '0',
			  'opacity': '0'
		  });
		  $('#customErrorId').remove();
		  $('.alert-errors[data-v-386d995a]').addClass('errors');
		  //$('span').filter(function () { return $(this).text() == buttonText; }).parent().css('color', '#444').removeAttr('disabled');
	  }
	}

	$('button.close').on('click', () => {
		$('div.alert').css({
			'height': '0',
			'opacity': '0'
		});

		$('.fd-grid, .container-fluid').css({'margin-top': '0px'});
	 });

	var intervalId = setInterval(checkForErrors, 100);
}

const chckRequiredFields = async function(){
    let mesg = '';
    for (const fieldname in fd.spForm.fields) {
        let field = fd.field(fieldname);
        let isRequired = field.required;
        let val = field.value;
        if(isRequired && (val === undefined || val === null || val === ''))
            mesg += `${fieldname} is required <br/>`

    }

    if(mesg !== ''){
        setPSErrorMesg(mesg);
        return false
    }
    else return true
}

function isNullOrEmpty(value) {
	return value === null || value === undefined || value === '' || value.length === 0;
}

function htmlEncode(str) {
    return str.replace(/[&<>"']/g, function(match) {
        return {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        }[match];
    });
}

function TextQueryEncode(str) {
    return str.replace(/[&<>"']/g, function(match) {
        return {
            '&': '%26'
        }[match];
    });
}

function clearLocalStorageItemsByField(fields) {
	fields.forEach(field => {
		let cachedFields = localStorage;
        for (let i = 0; i < cachedFields.length; i++) {
	        const key = localStorage.key(i);
	        if (key.includes(field)){
	        	localStorage.removeItem(key);
	        }
    	}
    });
}

function clearFormFields(fields){
	for(let i = 0; i < fields.length; i++){
        const field = fields[i];
		if(field !== 'InkSign'){

			var fieldproperties = fd.field(field);

			if(fieldproperties._fieldCtx !== undefined){
				var fieldDefaultVal = fieldproperties._fieldCtx.schema.DefaultValue;

				if (fieldDefaultVal !== undefined && fieldDefaultVal !== null) {}
				else
					fd.field(field).clear();
			}
			else fd.field(field).clear();
		}
    }
}

function disableRichTextField(fieldname){

	let elem = $(fd.field(fieldname).$el).find('.k-editor tr');

	elem.each(function(index, element){

	 if(index === 0)
		$(element).remove()

	 else if(index === 1){

		let iframe = $(element).find('iframe');

		if(iframe.length > 0){

			let content = iframe.contents();
			let divElement = content.find('div');

			var lblElement = $('<label>', {
			  for: 'inputField',
			}).html(divElement.html());

			if(divElement.length === 0){
				lblElement = $('<label>', {
					for: 'inputField',
				  }).html(content[0].activeElement.innerHTML);
			}

			lblElement.css({
				'padding-top': '6px',
				'padding-bottom': '6px',
				'padding-left': '12px',
				'background-color': '#e9ecef',
				'width': '100%',
				'border-radius': '4px'
			});

			let tblElement = iframe.parent().parent().parent().parent();
			tblElement.parent().append(lblElement);
			tblElement.remove();
		}
	   }
	})
}


function setPSHeaderMessage(newText) {

	document.querySelectorAll('.alert-heading').forEach(el => {
		el.style.display = 'block'; // Remove display: none
		if (newText && el.textContent.includes("Oops! There seem to be some errors below:")) {
            el.textContent = newText;
        }
    });

	document.querySelector('.fd-grid.container-fluid').style.setProperty('margin-top', '5px', 'important');
}