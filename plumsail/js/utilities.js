const currentUser = async function(){
	var LoginName = '';
	await pnp.sp.web.currentUser.get()
		         .then(async (user) =>{
					debugger;
					LoginName = user.LoginName;
	 });
	 return LoginName;
}
  
const customButtons = async function(icon, text, isAttach, Trans, isAttachmentMandatory, isSubmit, isOneFile, isPart, moduleName){
	  fd.toolbar.buttons.push({
	        icon: icon,
	        class: 'btn-outline-primary',
	        text: text,
	        click: async function() {
			 debugger;
			 if(text == "Close" || text == "Cancel")
			 {
				 
				 //debugger;
				 //preloader(true);			
				 fd.validators.length = 0;
				 fd.close();
			 }
             else if(moduleName == "TEST" && text == "Click Me"){
				var xhttp = new XMLHttpRequest();
				xhttp.onreadystatechange = function() {
					 if (this.readyState == 4 && this.status == 200) {
						 alert(this.responseText);
					 }
				};
				xhttp.open("GET", "http://db-sp.darbeirut.com:8080/api/v1/hw", true);
			
				//xhttp.open("GET", "https://jsonplaceholder.typicode.com/todos/1", true);
				//xhttp.open("GET", "http://db-sp16.darbeirut.com/api/v1/hw", true);
				xhttp.setRequestHeader("Content-type", "application/json");
				xhttp.send("Your JSON Data Here");
			 }
			 else if(moduleName == "TEST" && text == "Submit")
			 {
				//https://maps.googleapis.com/maps/api/js?key=YOUR_API_KEY&callback=initMap
				 fd.validators.length = 0;

				 const payload = {
					method: 'GET',
					headers: { "Accept": "application/json; odata=verbose" },
					credentials: 'same-origin'    // or credentials: 'include'  
				};

				 const userAction = async() => {
					const response = await fetch('http://db-sp.darbeirut.com/projects/PILOT-S/_api/web/lists',payload);//'http://db-sp.darbeirut.com:3000/api/v1/tours');
					const myJson = await response.json(); //extract JSON from the http response
					const len = await myJson.d.results.length;
					// do something with myJson
					alert(len);
				  }
				  var xx = await userAction();

				 //preloader(false);
				 //fd.close();
			 }
			 else if(text == "Compile & Close")
			 {
				 fd.validators.length = 0;
				 const params = {
					             ID: fd.itemId, 
								 ListId: fd.spFormCtx.ListAttributes.Id
							  };
				 const webUrl = `${window.location.protocol  }//${  window.location.host  }${_spPageContextInfo.siteServerRelativeUrl}`;
				 const Pageurl = `${webUrl + _layout  }/CompileForm.aspx?${  $.param(params)}`;
                 window.open(Pageurl,"_blank"); 
				 return;
			 }
			 else if(text == "Preview Form")
			 {
				debugger;
				 fd.validators.length = 0;
				 const params = {
					             ID: fd.itemId,
								 ListId: fd.spFormCtx.ListAttributes.Id
							    };
				 //const webUrl = `${window.location.protocol}/${window.location.host}${_spPageContextInfo.siteServerRelativeUrl}`;
				 const webUrl = _spPageContextInfo.siteAbsoluteUrl;
				 const Pageurl = `${webUrl + _layout}/PrintView.aspx?${$.param(params)}`;
                 window.open(Pageurl,"_blank"); 
				 return;
			 }
			 else if(text == "Print")
			 {
				fd.validators.length = 0;
				PrintCustom();
				 return;
			 }
			 else if(text == "Send to PMC")
			 {
				 $(fd.field('SentToPMC').$parent.$el).show();
				 fd.field('SentToPMC').value = true;
				 $(fd.field('SentToPMC').$parent.$el).hide();
				 
				 $(fd.field('BIC').$parent.$el).show();
				 fd.field('BIC').value = "Office";
				 $(fd.field('BIC').$parent.$el).hide();
				 fd.save();
				 return;
			 }
			 else if(text === "Assign")
			 {
				let leadTrade = "", partTrades;
				 partTrades = fd.field('Part').value;
				 $(fd.field('PartTrades').$parent.$el).show();
				 fd.field('PartTrades').value = partTrades;
				 $(fd.field('PartTrades').$parent.$el).hide();
				 
				 if(_isPMC || !_isAutoAssign) //!_isAutoAssign || 
				 {
					if(fd.field('Lead').value == null || fd.field('Lead').value == ""){
					  setErrorMessage(leadTradeRequiredFieldMesg, "Assign", false);
					  return false;
					}
					setErrorMessage(leadTradeRequiredFieldMesg, "Assign", true);

					 leadTrade = fd.field('Lead').value;
					 $(fd.field('LeadTrade').$parent.$el).show();
					 fd.field('LeadTrade').value = leadTrade;
					 $(fd.field('LeadTrade').$parent.$el).hide();
				 }
				 else{
					if(fd.field('Part').value === null || fd.field('Part').value === ""){
					  setErrorMessage(partTradeRequiredFieldMesg, "Assign", false);
					  return false;
					}
					setErrorMessage(partTradeRequiredFieldMesg, "Assign", true);
				 }
				 
				 try{
						$(fd.field('Assigned').$parent.$el).show();
						fd.field('Assigned').value = true;
						$(fd.field('Assigned').$parent.$el).hide();
						
						$(fd.field('AssignedDate').$parent.$el).show();
						fd.field('AssignedDate').value = new Date();
						$(fd.field('AssignedDate').$parent.$el).hide();
				 }
				 catch(err){}				
		         fd.save();
				 return;
			 }
			 else if(text == "Reject")
			 {
				var _rows = hTable.getData();
				var rowCount = _rows.length;
				var rejectedTrades = '';
				var isSelected = false;
				for(var row = 0; row < rowCount; row++){
					//linkIndex = 0, tradeIndex = 0, statusIndex = 0, remarkIndex = 0;
					var linkValue = _rows[row][linkIndex];
					var tradeValue = _rows[row][tradeIndex];
					//var statusValue = _rows[row][statusIndex];
					var remarkValue = _rows[row][remarkIndex];

                    if(linkValue && (remarkValue === undefined || remarkValue === null || remarkValue === '') ){
						alert(reject_SLF_TeamLeader_ValidationMesg + ' for ' + tradeValue);
						return;
					}
                    
					else if(linkValue){
					  isSelected = true;
					  rejectedTrades += `${tradeValue}|${remarkValue};`
					}
				}

				if(!isSelected){
					alert(rejectButton_SLF_TeamLeader_ValidationMesg);
					return;
				}
				
				$(fd.field('isRejected').$parent.$el).show();
				fd.field('isRejected').value = true;
				$(fd.field('isRejected').$parent.$el).hide();
				
				$(fd.field('RejectionTrades').$parent.$el).show();
				fd.field('RejectionTrades').value = rejectedTrades;
				$(fd.field('RejectionTrades').$parent.$el).hide();
				fd.save();
				return;
			 }
			 else if(moduleName == "FNC" && text == "Submit"){
                debugger;
				var haveErrors= await checkColumnsFeatureSelection();
				if(haveErrors)
				  return;

				//#region SET SCHEMA
				var _result = await setFilename_Schema();
				var multiLineString = '';
				var fieldsArray = _result._fieldArray;
				if(fieldsArray.length > 0){
					multiLineString = '[\n';
					for (var i = 0; i < fieldsArray.length; i++){
						var obj = fieldsArray[i];
						var fieldsLength = Object.keys(obj).length;

						multiLineString += '{\n';	
						var colIndex = 1;				
						for (var key in obj) {
							
							var _fiedlValue = obj[key];
							var valType = typeof _fiedlValue;

							if(valType !== 'boolean' && valType !== 'number'){
								if( key !== 'Title' && key !== 'DeliverableType' && key !== 'MainList')
							      _fiedlValue = _fiedlValue.trim().replace(' ', '');
								else _fiedlValue = _fiedlValue.trim();
							  multiLineString += '\"' + key + '\": \"' + _fiedlValue + '\"';
							}
							else multiLineString += '\"' + key + '\": ' + _fiedlValue + '';

							
							if( colIndex === fieldsLength)
							 multiLineString += '\n';
							else multiLineString += ',\n';
							colIndex++;
						}
						multiLineString += '},\n';
					}

					multiLineString = multiLineString.trim();
					if (multiLineString.endsWith(",")) 
					  multiLineString = multiLineString.slice(0, -1);
					multiLineString += '\n]';
					
					fd.field('Schema').value = multiLineString.replace(/&nbsp;/g, '');
			    }
				//#endregion

				var containsRev = _result.containsRev;

				if(containsRev){
				var isValidResult = await isCorrect();
					if(!isValidResult){
						set_FNC_ErrorMessage('Revision must be at the end of the filename', false);
						return;
					}
			    }

				if(multiLineString === ''){
					set_FNC_ErrorMessage('One field at least must be selected', false);
					return;
				}

				
				fd.field("Title").value = _acronym;
				fd.save();
			 }
			 else if(moduleName == "DPR" && text == "Submit"){
				const listname = 'PR Template';
				var colsInternal = ['Title', 'No','Hrs'];
				var gridControls = ['OnSiteStaffLabour', 'OnSitePlant'];
				var gridCategories = ['Staff', 'Equipment'];
				
				for(var i = 0; i < gridControls.length; i++){
				  var _gridsItems = [];
				  var gridControl = gridControls[i];
				  var gridCategory = gridCategories[i];
			
				  var data = fd.control(gridControl).widget.dataSource.data();
				   data.map(item =>{
					  var _gridItems = {};
					  colsInternal.map(field =>{
						_gridItems[field] = item[field];
					  });
					  _gridItems['TableName'] = gridCategory;
					  _gridItems['DeliverableType'] = moduleName;
					  _gridsItems.push(_gridItems);
				   });
				   await insertItemsInBulk(_gridsItems, listname, gridCategory, colsInternal);
				}
			 }
			 else if(moduleName == "DPR" && text == "acknowledge"){
				$(fd.field('Submit').$parent.$el).show();
				fd.field('Submit').value = true;
				$(fd.field('Submit').$parent.$el).hide();
			 }

             if(isAttach){
				 if(isAttachmentMandatory || (isSubmit && _isPart))
				 {
					fd.validators;
					fd.validators.push({
						error: "Please upload the softcopy",
						validate: function(value) {	
							if(isSubmit && _isPart && moduleName != 'SLF'){
								if (fd.field('Code').value === null || fd.field('Code').value === ""){
									this.error = "Code is required.";
									return false;
								}
							}

							else if(isOneFile)
							{
								if(fd.field('Attachments').value.length == 0)
								   return false;
								   
								if(fd.field('Attachments').value.length < 1)
								{
									this.error = "PDF File is Required to be attached on Submit";
									return false; 
								}
										 
								if(fd.field('Attachments').value.length > 1)
								{
									this.error = 'Only one pdf file is required, please compile as neccessary.';
									return false; 
								}

								 if(fd.field('Attachments').value.length > 0)
								 {
									   const valext = fd.field('Attachments').value[0];
									   const ext = fd.field('Attachments').value[0].name.split(".")[1];
									   //alert(ext);		  
									   if(ext.toString().toLowerCase() != 'pdf')
									   {
										  this.error = "PDF File is Required and dots are not allowed in the filename";
										  return false;
									   }
								 }	
							}
							
						return true;
						}
					});
				 }
				 else fd.validators.length = 0;
				
				 if(!fd.isValid)
				   $(fd.field('AttachFiles').$parent.$el).hide();
				 else
				 {			   
				    let NCountofATT = 0;
					let OCountofATT = 0;
					for(i = 0; i < fd.field('Attachments').value.length; i++) 
					{
						const Val = fd.field('Attachments').value[i].extension.toString();
						if(Val === "")
						{OCountofATT++;}
						else
						{
								IsNewAttachment = true;
								NCountofATT++;			
						}
					}

					if(moduleName !== "SLFI"){
						$(fd.field('AttachFiles').$parent.$el).show();
						fd.field('AttachFiles').value = `${Trans  },${  NCountofATT  },${  OCountofATT}`;
						$(fd.field('AttachFiles').$parent.$el).hide();	
					}

					if(isSubmit)
					{
						if(isPart)
						{
							if(moduleName != 'SLF'){
								$(fd.field('Status').$parent.$el).show();
								fd.field('Status').value = "Completed";
								$(fd.field('Status').$parent.$el).hide();
							}
						}
						else {
							if(moduleName !== "SLFI")
							 fd.field('Submit').value = true;
						}
					}
					
					 if(_module == 'SLF'){

						if(_isLead == false && _isPart == false){

							if(isContractor && data.length > 0){
                                var showError = true;
								if(editedCells !== null && editedCells.length > 0){
									for(var j =0; j < editedCells.length; j++){
										var cellValue = editedCells[j][3];
                                     if(cellValue !== '' && cellValue !== null && cellValue !== undefined){
										showError = false;
										break;
									 }
									}
								}
								if(showError){
								 alert(contractorRequiredFieldMesg);
								 $(fd.field('AttachFiles').$parent.$el).hide();
								 return;
								}
								await insertPLFItems(element, targetList, targetFilter, '');
							}
							else{
								var partTrades = fd.field('Part').value;
								if(partTrades[0] === null || partTrades[0] === ""){
									setErrorMessage(partTradeRequiredFieldMesg, text, false);
									$(fd.field('AttachFiles').$parent.$el).hide();
									return false;
								}

								$(fd.field('PartTrades').$parent.$el).show();
								fd.field('PartTrades').value = partTrades
								$(fd.field('PartTrades').$parent.$el).hide();

								$(fd.field('AssignedDate').$parent.$el).show();
								fd.field('AssignedDate').value = new Date();
								$(fd.field('AssignedDate').$parent.$el).hide();
							}
						}

						else if(_isPart && !_isTeamLeader){

							var errorLength = Object.keys(masterErrors).length;
							if (errorLength > 0){
								alert('fix errors in the table below and try again');
								$(fd.field('AttachFiles').$parent.$el).hide();
								return;
							}
							else if (isFilterSelected){
								alert('remove filter columns before submission');
								$(fd.field('AttachFiles').$parent.$el).hide();
								return;
							}
							else {
								if(!fd.isValid) return;
								var trade = fd.field('Trade').value;
								await insertPLFItems(element, targetList, targetFilter, trade);
								if(!_isTeamLeader && _doHide){
								 var confirmed = window.confirm("Do you want to attach the items ?");
								 if (confirmed) {
									$(fd.field('AttachFiles').$parent.$el).hide();
									//hideColumns('', false);
									//SET ATTACH FILES PER ROW
									location.reload();
									return;
								 }
								 else {
									if(text !== "Save")
										$(fd.field('Status').$parent.$el).show();
										fd.field('Status').value = "Completed";
										$(fd.field('Status').$parent.$el).hide();
								 }
								 
								}
							}
						}

						else if(_isPart && _isTeamLeader){
                          if(_haveOpenItems){
							alert(submit_SLF_TeamLeader_ValidationMesg);
							$(fd.field('AttachFiles').$parent.$el).hide();
							return;
						  }
						}

						// 	if(text == "Submit")
						// 	 fd.field('Status').value = 'Open';
						// 	if(_formType === 'New')
						// 	 setCounter(counterType, counter);

						$(fd.field('Status').$parent.$el).show();

						if(_isPart){
						 if(text !== "Save")
						  fd.field('Status').value = "Completed";
						}
						else if(isContractor && _isMain && Status === 'Issued to Contractor'){
						   fd.field('Status').value = "Assigned";

						   $(fd.field('Resubmit').$parent.$el).show();
						   fd.field('Resubmit').value = true;
						   $(fd.field('Resubmit').$parent.$el).hide();
						}
						$(fd.field('Status').$parent.$el).hide();
					 }
		            fd.save();
				 }
			 }
			 else{
				if(_module === 'AUR' && _isEdit){
				  $(fd.field('SummaryPlainText').$parent.$el).show();
				  fd.field('SummaryPlainText').value = fd.field('Summary').value;
				  $(fd.field('SummaryPlainText').$parent.$el).hide();

				  $(fd.field('CARSummaryPlainText').$parent.$el).show();
				  fd.field('CARSummaryPlainText').value = fd.field('CARSummary').value;
				  $(fd.field('CARSummaryPlainText').$parent.$el).hide();
				  
				  if(activeTabName === auditReportTab){
                    if(_isSentForReview)
					 fd.field('Code').value = '';

					$(fd.field('Submit').$parent.$el).show();
					fd.field('Submit').value = true;
					$(fd.field('Submit').$parent.$el).hide();
				  }
				}
				fd.save();
			 } 
	     }
      });
}

const isUserAllowed = async function(userName){
	let _isAllowed = false
	await pnp.sp.web.currentUser.get()
		 .then(async (user) =>{
			if(user.Title == userName){
				_isAllowed = true;
			}
		});	
		
		if(!_isAllowed)
		{
			const {userId} = _spPageContextInfo;
            const userGroups = [];
			await pnp.sp.web.siteUsers.getById(userId).groups.get()                               
			.then(async (groupsData) =>{                        
				for (let i = 0; i < groupsData.length; i++) {
					userGroups.push(groupsData[i].Title);
				}
				if(userGroups.indexOf('Owners') >= 0 || _spPageContextInfo.isSiteAdmin)
				  _isAllowed = true;
			});
		}

		return _isAllowed;	
}
 
function setErrorMessage(err, textButton, isRemove){
	if(!isRemove)
	{
		const errLength = $('div.alert').length;
		if(errLength == 0){
			const appendCtrl = "<div role='alert' class='alert alert-danger'>" +
								"<button type='button' data-dismiss='alert' aria-label='Close' class='close'>" +
								"<span aria-hidden='true'>×</span>" +
								"</button>" +
								"<h4 id='errId' class='alert-heading'>Please correct the errors below:</h4><br/><br/>" +
							"</div>";

			const masterErrorPosition = $('div.container-fluid').first();
			masterErrorPosition.prepend(appendCtrl);
		}
		
		var elem = $('#errId').find(`p:contains('${  err  }')`);
		if(elem.length == 0){
			$('#errId').append(`<p style='font-size: 14px; font-family: Arial; padding-top:10px'>${  err  }</p>`);
			if(textButton != ""){
				if(_module !== 'SLF')
				 $('span').filter(function(){ return $(this).text() == textButton; }).parent().attr("disabled", "disabled");
			}
		}
	}
	else {
        var pLen = $('#errId').find('p').length;
        if(pLen == 0){
		  $('div.alert').remove();
		  $('span').filter(function(){ return $(this).text() == textButton; }).parent().removeAttr('disabled');
		}
		else
		{
			var elem = $('#errId').find(`p:contains('${  err  }')`);
			if(elem.length > 0)
			{
				$(elem).remove();
				var pLen = $('#errId').find('p').length;
				if(pLen == 0){
				  $('div.alert').remove();
				  $('span').filter(function(){ return $(this).text() == textButton; }).parent().removeAttr('disabled');
				}
			}
		}
	}
}

function set_FNC_ErrorMessage(err, isRemove){
	if(!isRemove)
	{
		$('div.alert').remove();

		const appendCtrl = "<div role='alert' class='alert alert-danger'>" +
								"<div class='alert-body'>" + 
									"<button onclick='closePanel(this)' type='button' data-dismiss='alert' aria-label='Close' class='btn close'>" +
									"<i aria-hidden='true' class='ms-Icon ms-Icon--Cancel'></i>" +
									"</button>" +
								"<h4 id='errId' class='alert-heading'>Please correct the errors below:</h4>" +
								"</div>" +
								// "<div class='alert-expander'>" +
								// "<i aria-hidden='true' class='ms-Icon ms-Icon--DoubleChevronDown'></i>" + 
								// "</div>" +
							"</div>";

		const masterErrorPosition = $('div.container-fluid').first();
		masterErrorPosition.prepend(appendCtrl);
		
		var elem = $('#errId').find(`p:contains('${  err  }')`);
		if(elem.length == 0){
			$('#errId').after(`<p style='font-size: 14px; font-family: Arial; padding-top:10px'>${  err  }</p>`);
		}
		$('.ControlZone').css('overflow', 'auto');
	}
}

function closePanel(element){
	$('div.alert').remove();
	$('div.ControlZone').css('overflow', '');
}

const ApplyFieldChanges = async function(fields, disableControls, disableCustomButtons, defaultButtons, _status){
	 let isAllowed = false; //await IsUserInGroup("Owners");
	
     if(_isSiteAdmin)
	     isAllowed = true;
	 
	 for(let i = 0; i < fields.length; i++){
		 if(fields[i] == "Attachments")
		 {	
			if(_isLead && !isAllowed)
			 $(fd.field('Attachments').$parent.$el).hide();
		   else if(!isAllowed)
		   {		
			 $("div.k-upload-sync").removeClass("k-state-disabled");
			  if(taskStatus !== 'Open'){
			    $('div.k-upload-button').remove();
				$('button.k-upload-action').remove();
				$('.k-dropzone').remove();
			 }
		   }
		   else {
			   fd.field('Attachments').disabled = false;			   
			   //if(_status == "Completed" || _status == "Closed" || _status == "Issued to Contractor")
			   await customButtons("Save", "Save Attachment", true, "U");
		   }
		 }
		 
		else try { 
		
		if(fields[i] == "Status")
		{	   
		   if(isAllowed)
		   {
			   var GetStatus = fd.field("Status").value;
			   if(GetStatus == "Closed" || GetStatus == "Issued to Contractor" || GetStatus == "Completed")
			   {
				   $(fd.field('Status').$parent.$el).show();
				   fd.field('Status').disabled = false;
			   }
			   else
				   $(fd.field('Status').$parent.$el).hide();				   
			  
		   }
		   else
		    $(fd.field('Status').$parent.$el).hide();	   
		}		
		else
			fd.field(fields[i]).disabled = disableControls;	
			
		} 
		catch(e){alert(`fields[i] = ${  fields[i]  }<br/>${  e}`);}
	 }

	 if(disableCustomButtons)
	 {
	   	if(defaultButtons != null && defaultButtons != 'undefined')
			fd.toolbar.buttons[0].disabled = true;
		
		$('span').filter(function(){ return $(this).text() == 'Save'; }).parent().attr("disabled", "disabled");
		$('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().attr("disabled", "disabled");
	 }
}

const showHideTabs = async function(key, BIC, isPart, isLead){
	let _isAllowed = false;
	
    try
    {
        // get all groups the current user belongs to
        const {userId} = _spPageContextInfo;
        const userGroups = [];
        await pnp.sp.web.siteUsers.getById(userId).groups.get()                               
        .then(async (groupsData) =>{                        
			for (let i = 0; i < groupsData.length; i++) {
				userGroups.push(groupsData[i].Title);
			}
								
			let isAdmin = false;
			const Status = fd.field('Status').value;			
			
			if(userGroups.indexOf('Owners') >= 0 || _spPageContextInfo.isSiteAdmin) isAdmin = true;
											
			let isClosed = false;
			if(Status == 'Issued to Contractor' || Status == 'Completed') isClosed = true;
												
			let isDar = false;                                                     
			if(userGroups.indexOf('PMC') >= 0) isDar = true;  
			if(userGroups.indexOf('RE') >= 0) isDar = true; 
					
			if(isAdmin && isClosed)
				fd.field('Status').disabled = false;
			else if(isClosed)
				fd.field('Status').disabled = true;
			
			if(isPart || isLead){
				let listName, codes;
				const mType = await getMajorType(key);
				if(mType !=null && mType.length > 0){
					listName = mType[0].MatrixList;
					_AllowManualAssignment = mType[0].AllowManualAssignment;
					_isAutoAssign = mType[0].isAutoAssign;
				}
			else listName = await getParameter("Inspection-Matrix");
  

			if(isLead){
				if(_AllowManualAssignment){
					if( isClosed || !isDar )
					  fd.container('Tab1').tabs[1].disabled = true;
				}
			    else 
				{
					if(_isAutoAssign)
					  $("ul.nav,nav-tabs").remove();	
				}
				_isAllowed = true;
			}
			  
			else if(isPart)
			{
				if(_AllowManualAssignment)
				{           
					if(Status !== "Completed"){       
						if( (BIC == "Site" && userGroups.indexOf('RE') >= 0) || (BIC == "Office" && Status == "Send to PMC" && userGroups.indexOf('PMC') >= 0))
						{
						  fd.container('Tab1').tabs[1].disabled = false; 
						  _isAllowed = true;
						}
						else if(isPart && fd.field('Trade').value == "PMC" && userGroups.indexOf('PMC') >= 0){
							 fd.container('Tab1').tabs[1].disabled = false; 
							_isAllowed = true;
						} 
						else{
							fd.container('Tab1').tabs[1].disabled = true; 
							_isAllowed = false;
						}
					}
					else fd.container('Tab1').tabs[1].disabled = true
				}
				else {
					$("ul.nav,nav-tabs").remove();
					_isAllowed = true;
				}
			}
			$('span').filter(async function(){ return $(this).text() == 'Assign'; }).parent().prop("disabled", false);
			
				var orderColumn = 'Title';
				if(isPart && _isPMC)
					orderColumn = 'PMCTrades';
				
				const query = pnp.sp.web.lists.getByTitle(listName).items.select("Title,PMCTrades");
					await query.orderBy(orderColumn, true)
						.get()
						.then(async (items) => { 
								const TradeArray = [];
								const leadArray = [];
								for(let i = 0; i < items.length; i++)
								{
									let trade = '';
									if(isPart && _isPMC){
										trade = items[i].PMCTrades;
										if(trade != null){
											leadArray.push(trade);
											TradeArray.push(trade);
										}
									}

									else{
											trade = items[i].Title;
										if(isPart && _isPMC){}
										else TradeArray.push(trade);
										
										// if((key == "MAT" || key == "SCR") && trade != "PMC"){
											leadArray.push(trade);
										// }
										
										if(listName == "Trades")
										{
											const {Permission} = items[i];
											if(Permission != null && Permission.length > 0)
											{
												for(let j = 0; j < Permission.length; j++)
												{
													const group = Permission[j];
													const isUserValid = await IsUserInGroup(group);
													//console.log("isUserValid = " + isUserValid);
													if(isUserValid)
													{
														if(isPart && _isPMC){}
														else TradeArray.push(trade);
														break;
													}
												}
											}
										}
									}
								}
								await setTrades(TradeArray, leadArray, key, isLead);
						}); 	
			}   			
        });
		return _isAllowed;
    }
    catch(err){
        alert(err);
    }
}

var DisableFields = async function (filterColumn, filterValue, operator, fields, disableControls, disableCustomButtons, defaultButtons){
	 try
	 {
	   let _status = fd.field(filterColumn).value;
	   if(_status == "") {
		if(_module === 'SLF')
		 _status = "Pending";
		else _status = "Initiated";
	   }
	  
		   
	   if(operator == "eq")
	   {
			if(_status == filterValue || taskStatus === 'Completed')
			   await ApplyFieldChanges(fields, disableControls, disableCustomButtons, defaultButtons, _status);
	   }

	   else
	   if(_status != filterValue)
		    await ApplyFieldChanges(fields, disableControls, disableCustomButtons, defaultButtons, _status);
	 }
	catch(e){ console.log(e); }
}

function HideFields(fields, hideButtons, isHide){
	if(hideButtons)
	{
	  fd.toolbar.buttons[0].style = "display: none;";
      fd.toolbar.buttons[1].style = "display: none;";
	}
	
	var field;
	for(let i = 0; i < fields.length; i++)
	{
		field = fd.field(fields[i]);
		if(isHide || isHide == undefined)
		  $(field.$parent.$el).hide();
		else $(field.$parent.$el).show();
	}
}
  
function PrintCustom(){
	var contentData;
	if(_module === 'AUR'){
		if(activeTabName === auditReportTab){
			contentData = fd.field('Summary').value;
			var elements = $.parseHTML(contentData);

			var h1 = $(elements).filter('#header1');
			var h2 = $(elements).filter('#header2');
			if(h1.length > 0 && h2.length > 0){
			_header = '<table id="header1" style="border-collapse: collapse; color:black; margin:auto;" border="0" cellspacing="0" cellpadding="2" width="100%" align="center">' +
							h1[0].innerHTML + 
						'</table>' +
						'<table id="header2" style="border-collapse: collapse; color:black; margin:auto;" border="0" cellspacing="0" cellpadding="5" width="100%" align="center">' +
						h2[0].innerHTML + 
						'</table>';
			}

			var b1 = $(elements).filter('#body1');
			if(b1.length > 0){
			_main = b1[0].innerHTML;
			var partMainHTML = $(b1).nextAll('p');
			if(partMainHTML.length > 0){
					var combinedInnerHTML = "";
					partMainHTML.each(function() {
						combinedInnerHTML += '<p>' + $(this).html() + '</p>';
					});
					_main += combinedInnerHTML;
			}		    
			}

			var f1 = $(elements).filter('#footer1');
			if(f1.length > 0){
				_footer = '<table id="footer1" style="border-collapse: collapse; color: black; margin: auto; text-align: center;" border="0" cellspacing="0" cellpadding="5" width="100%">' +
							f1[0].innerHTML + 
						'</table>';
			}
			contentData = _header + _main + _footer;
		}
		else if(activeTabName === corrActionTab){
			contentData = fd.field('CARSummary').value;
		}
	}

	  const printWindow = window.open('','printWindow');
	  printWindow.document.open();
	  printWindow.document.write(contentData);
	  printWindow.document.close();
	  //printWindow.focus();

	  printWindow.print();
	  printWindow.close();
}

var setTrades = async function (array, leadArray, moduleName, isLead){
	fd.field('Part').widget.dataSource.data(array);
	 
	const partArray = [];
	let Trade = ""; let PartTrades = "";
	PartTrades = fd.field('PartTrades').value;

	if(isLead && _isAutoAssign){
		Trade = fd.field('Trade').value;
		if(Trade != null & Trade != ""){
			if(PartTrades != null & PartTrades != ""){}
				//PartTrades = `${Trade},${PartTrades}`;
			else PartTrades = Trade;
		}
	}
	
   if(PartTrades != null)
   {
	  if(PartTrades.includes(","))
	  {
        PartTrades = PartTrades.split(',');
		for(let i = 0; i < PartTrades.length; i++)
		{
			partArray.push(PartTrades[i]);
		}
	  }
	  else partArray.push(PartTrades);
      fd.field('Part').value = partArray;
   }
   
   if(_module != 'SLF' && !_isAutoAssign || _isPMC){
	   fd.field('Lead').widget.dataSource.data(leadArray);
	   const _leadTrade = fd.field('LeadTrade').value;
	   if(_leadTrade != null)
	   {
	     const leadObject = [];
	     leadObject.push(_leadTrade);
	     fd.field('Lead').value = leadObject;
	   }
   }
   else if(moduleName == "SLF" && _formType == "New"){
	//fd.field('LeadTrade').value = _spPageContextInfo.userDisplayName; 
    //fd.field('LeadTrade').disabled = true 
	debugger;
	var LoginName = await currentUser();
	fd.field('LeadInspector').value = LoginName; //.replace('i:0#.w|','');	 
	fd.field('LeadInspector').disabled = true;
   }
}

function setDefaultTrades(items){
   const array = [];
   array.push("");
   for(let i = 0; i < items.length; i++)
   {
     array.push(items[i].Title);
   }
   setTrades(array);
}

function checkTradeAssignment(columnName){
	fd.field(columnName).$on('change', (value) => {   
		if(value == "")
		{
			if( (fd.field('Lead').value == null || fd.field('Lead').value.length == 0) && (fd.field('Part').value == null || fd.field('Part').value.length == 0)) {
				fd.field(columnName).clear();
				$('span').filter(function(){ return $(this).text() == 'Send to PMC'; }).parent().removeAttr('disabled');
			  }
			return;
		} 
 
		const isLeadNull = isFieldNULL("Lead");
		const isPartNull = isFieldNULL("Part");

		  if( !isLeadNull ||  !isPartNull)
			  $('span').filter(function(){ return $(this).text() == 'Send to PMC'; }).parent().attr("disabled", "disabled");
		  else if(isLeadNull && isPartNull) {
			fd.field(columnName).clear();
			$('span').filter(function(){ return $(this).text() == 'Send to PMC'; }).parent().removeAttr('disabled');
		  }

		  if(!_isAutoAssign && fd.field('Lead').value == null)
			setErrorMessage(leadTradeRequiredFieldMesg, "Assign", false);
		  else setErrorMessage(leadTradeRequiredFieldMesg, "Assign", true);

		  if(fd.field('Lead').value != null && fd.field('Part').value != null)
		  {
			 const _lead = fd.field('Lead').value;
			 const _part = fd.field('Part').value;

				if(_part != null && _part.length > 0)
				{
					let isError = false;
					for(let i = 0; i < _part.length; i++)
					{
						if(_part[i] == _lead)
						{
							isError = true;
							break;
						}
					}
					
					if(isError)
						setErrorMessage(partError, "Assign", false);
					else setErrorMessage(partError, "Assign", true);
				}
			}
	 }); 
}

function isFieldNULL(columnName){
	let isNULL = true;
	try{
	  if (fd.field(columnName).value.length > 0)
	    isNULL = false
	}
	catch(err){}
	return isNULL;
}

fd.spBeforeSave(() =>
{
	preloader();
});


var setButtonToolTip = async function(_btnText, toolTipMessage){
  
    var btnElement = $('span').filter(function(){ return $(this).text() == _btnText; }).prev();
	if(btnElement.length === 0)
	  btnElement = $(`button:contains('${_btnText}')`);
	
    if(btnElement.length > 0){
	  if(btnElement.length > 1)
		btnElement = btnElement[1].parentElement;
      else btnElement = btnElement[0].parentElement;
	  
      $(btnElement).attr('title', toolTipMessage);

      $(btnElement).tooltipster({
        delay: 100,
        maxWidth: 350,
        speed: 500,
        interactive: true,
        animation: 'slide', //fade, grow, swing, slide, fall
        trigger: 'hover'
      });
    }
}

var adjustlblText = async function(text, appendedText, isRequired){
	var targetLabel = $('label').filter(function(){ return $(this).text().trim() == text; });
	var _color = 'gray'
	if(isRequired) _color = 'red'

	var additionalText = $('<span>', {
		text: appendedText,
		css: {
		  color: _color,
		  'font-size': '11px',
		  'text-decoration': 'underline'
		}
	});
	
	targetLabel.append(additionalText);
}

