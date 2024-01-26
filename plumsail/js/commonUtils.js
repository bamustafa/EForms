var _schemaAcronyms = [];

var getCurrentWebUrl = async function() {
    return pnp.sp.web.get().then((web) => {
      return web.Url;
    });
}

var getDesign_GridSchema = async function(_web, webURL, key){
    var _colArray = [];
    var targetList;
    var targetFilter;
  
    await _web.lists
          .getByTitle("SchemaConfig")
          .items
          .select("ListName,HandsonTblSchema,DataQuery")
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

var addLegend = async function (lblId, _text, parentElement, step){
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
			label.css('color', '#4e778f');
			label.css('font-weight', 'bold');
			label.css('font-size', '16px');
			label.css('width', '1350px');
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
		await addLegend('modulelblId', moduleTitle, 'moduleTitle', 'same');
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
	  let result = "";
      var query = `Title eq '${  key  }'`;
     if(byId !== undefined && byId === true)
	   query = `ID eq ${  key  }`;

      await pnp.sp.web.lists
			.getByTitle("MajorTypes")
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

var getGridMType = async function(_web, webURL, key){
    //var _itemArray = {};
    var _colArray = [];
    var targetList, targetFilter, revNumStart, updateTitle;

    var result = "";
    await _web.lists
          .getByTitle("MajorTypes")
          .items
          .select("HandsonTblSchema,LODListName,LODFilterColumn,RevNumStart,UpdateTitle")
          .filter("Title eq '" + key + "'")
          .get()
          .then(async function (items) {
              if(items.length > 0)
                //for (var i = 0; i < items.length; i++) {
                  var item = items[0];
                  var HandsonTblSchema = item.HandsonTblSchema;
                  targetList = item.LODListName;
                  targetFilter = item.LODFilterColumn;
				  targetFilter = item.LODFilterColumn;
				  revNumStart = item.RevNumStart;
				  updateTitle = item.UpdateTitle;

                    var fetchUrl = webURL + HandsonTblSchema;
                    await fetch(fetchUrl)
                        .then(response => response.text())
                        .then(async data => {
                          _colArray = JSON.parse(data); 

                        //   for (const obj of _colArray) {
                        //     if (obj.renderer === "incrementRenderer"){
                        //       obj.renderer = incrementRenderer;
                        //     }
                        //     else if (obj.source === "getDropDownListValues"){
                        //       obj.source = await getDropDownListValues(obj.listname, obj.listColumn);
                        //     }
                        //     else if (obj.validator === "validateDateRequired"){
                        //       obj.validator = validateDateRequired;
                        //     }
                        //     else if (obj.validator === "getDescriptionValidator_PLF"){
                        //       obj.validator = getDescriptionValidator_PLF;
                        //     }
                        //     else if (obj.validator === "preventEdit"){
                        //       obj.validator = preventEdit;
                        //     }
                        //   }
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

var getSoapRequest = async function(method, serviceUrl, isAsync, soapContent, getParams){
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

var setButtonCustomToolTip = async function(_btnText, toolTipMessage){
  
    var btnElement = $('span').filter(function(){ return $(this).text() == _btnText; }).prev();
	if(btnElement.length === 0)
	  btnElement = $(`button:contains('${_btnText}')`);
	
    if(btnElement.length > 0){
	  if(btnElement.length > 1)
		btnElement = btnElement[1].parentElement;
      else btnElement = btnElement[0].parentElement;
	  
      $(btnElement).attr('title', toolTipMessage);

	  await _spComponentLoader.loadScript( _layout + '/controls/tooltipster/jquery.tooltipster.min.js').then(async () =>{
		$(btnElement).tooltipster({
			delay: 100,
			maxWidth: 350,
			speed: 500,
			interactive: true,
			animation: 'slide', //fade, grow, swing, slide, fall
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
				  errorMesg += `revision must start with ${item.RevStartWith} for ${fieldname} <br>`;
			}

			if(_module === 'LOD'){
				if(_revNumStart !== filenamePartValue.replace(item.RevStartWith,'')){
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
			.get();

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

					if(_isEdit){
						folderUrl = urlParams.get('source');
						urlParams = new URL(folderUrl);
						const searchParams = urlParams.searchParams;
						FolderName = searchParams.get('RootFolder');
					}
					 
					if(FolderName !== null){
						FolderName = FolderName.split("/")[5];
						if(filenamePartValue.toUpperCase() !== FolderName.toUpperCase())
						  errorMesg += `Contractor name must be ${FolderName}<br>`;
					}
					else if(!isFound)
					  errorMesg += `${fieldname} must be one of the following: ${listValuesMesg} etc..<br>`;
					else errorMesg += `Select a contractor name folder to submit your deliverables<br>`;
				 }
				 else{
					if(!isFound)
					  errorMesg += `${fieldname} must be one of the following: ${listValuesMesg} etc..<br>`;
					else{
						try{
							var obj = {  LookupId: itemId,
										 LookupValue: descValue
									  };
							fd.field(fieldname).value = obj;						 
						}
						catch{}
					}
				}
			}	
		}

		else if(item.isText !== undefined){
			var isTextValue = item.isText.toUpperCase();

			if(isTextValue !== filenamePartValue.toUpperCase())
			  errorMesg += `text must be ${isTextValue} instead of ${filenamePartValue} for ${fieldname} <br>`;
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
			 var isAllowed = await isAllowedCharacters(filenamePartValue, AllowedCharacters);
				if (!isAllowed)
				  errorMesg += `${fieldname} must be a text field with characters from ${AllowedCharacters}<br>`;
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