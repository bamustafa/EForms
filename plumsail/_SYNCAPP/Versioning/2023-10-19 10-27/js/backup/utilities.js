// var script = document.createElement("script");  // create a script DOM node
// script.src = _layout + "/plumsail/js/configFileRouting.js";  // set its src to the provided URL
// document.head.appendChild(script);
 
var getParameter = async function(key){
	var result = "";
      await pnp.sp.web.lists
			.getByTitle("Parameters")
			.items
			.select("Title,Value")
			.filter("Title eq '" + key + "'")
			.get()
			.then(function (items) {
				if(items.length > 0)
				  result = items[0].Value;
				});
	 return result;
 }
 
var getMajorType = async function(key){
	  var result = "";
      await pnp.sp.web.lists
			.getByTitle("MajorTypes")
			.items
			//.select("Title,Value")
			.filter("Title eq '" + key + "'")
			.get()
			.then(function (items) {
				if(items.length > 0)
				  result = items;
				});
	 return result;
 }
 
var IsUserInGroup = async function(group){
	    var IsPMUser = false
		try{
			 await pnp.sp.web.currentUser.get()
		         .then(async function(user){
			  await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
					 .then(async function(groupsData){
						for (var i = 0; i < groupsData.length; i++) {
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
  
var customButtons = async function(icon, text, isAttach, Trans, isAttachmentMandatory, isSubmit, isOneFile, isPart, moduleName){
	  fd.toolbar.buttons.push({
	        icon: icon,
	        class: 'btn-outline-primary',
	        text: text,
	        click: function() {
			 if(text == "Close")
			 {
				 fd.validators.length = 0;
				 preloader(false);
				 fd.close();
			 }
			 else if(text == "Compile & Close")
			 {
				 fd.validators.length = 0;
				 var params = {
					             ID: fd.itemId, 
								 ListId: fd.spFormCtx.ListAttributes.Id
							  };
				 var webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
				 var Pageurl = webUrl + _layout + "/CompileForm.aspx?" + $.param(params);
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
			 else if(text == "Assign")
			 {
				var leadTrade = "", partTrades;
				
				 partTrades = fd.field('Part').value;
				 $(fd.field('PartTrades').$parent.$el).show();
				 fd.field('PartTrades').value = partTrades
				 $(fd.field('PartTrades').$parent.$el).hide();
				 
				 if(moduleName == "MAT" || moduleName == "SCR")
				 {
					var leadError = "Lead Trade is Required to be filled";
					if(fd.field('Lead').value == null || fd.field('Lead').value == ""){
					  setErrorMessage(leadError, "Assign", false);
					  return false;
					}
					else setErrorMessage(leadError, "Assign", true);

					 leadTrade = fd.field('Lead').value;
					 $(fd.field('LeadTrade').$parent.$el).show();
					 fd.field('LeadTrade').value = leadTrade;
					 $(fd.field('LeadTrade').$parent.$el).hide();
				 }
				 else if(moduleName == "IR"){
					var leadError = "Part Trade is Required to be filled";
					if(fd.field('Part').value == null || fd.field('Part').value == ""){
					  setErrorMessage(leadError, "Assign", false);
					  return false;
					}
					else setErrorMessage(leadError, "Assign", true);
				 }
				 
				 try{
						$(fd.field('Assigned').$parent.$el).show();
						fd.field('Assigned').value = true;
						$(fd.field('Assigned').$parent.$el).hide();
						
						$(fd.field('AssignedDate').$parent.$el).show();
						fd.field('AssignedDate').value = new Date();
						$(fd.field('AssignedDate').$parent.$el).hide();
				 }
				 catch{}				
		         fd.save();
				 return;
			 }

             if(isAttach){
				 if(isAttachmentMandatory)
				 {
					fd.validators;
					fd.validators.push({
						name: 'Check Attachment',
						error: "Please upload the softcopy",
						validate: function(value) {	
							if(isOneFile)
							{
								if(fd.field('Attachments').value.length < 1)
								{
									this.error = "PDF File is Required to be attached on Submit";
									return false; 
								}
										 
								else if(fd.field('Attachments').value.length > 1)
								{
									this.error = 'Only one pdf file is required, please compile as neccessary.';
									return false; 
								}

								 else if(fd.field('Attachments').value.length > 0)
								 {
									   var valext = fd.field('Attachments').value[0];
									   var ext = fd.field('Attachments').value[0].name.split(".")[1];
									   //alert(ext);		  
									   if(ext.toString().toLowerCase() != 'pdf')
									   {
										  this.error = "PDF File is Required and dots are not allowed in the filename";
										  return false;
									   }
								 }	
							}
							else
							{
								if(fd.field('Attachments').value.length == 0)
									return false; 	
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
				    var NCountofATT = 0;
					var OCountofATT = 0;
					for(i = 0; i < fd.field('Attachments').value.length; i++) 
					{
						var Val = fd.field('Attachments').value[i].extension.toString();
						if(Val === "")
						{OCountofATT++;}
						else
						{
								IsNewAttachment = true;
								NCountofATT++;			
						}
					}
					$(fd.field('AttachFiles').$parent.$el).show();
					fd.field('AttachFiles').value = Trans + "," + NCountofATT + "," + OCountofATT;	
					if(isSubmit)
					{
						if(isPart)
						{
							$(fd.field('Status').$parent.$el).show();
							fd.field('Status').value = "Completed";
							$(fd.field('Status').$parent.$el).hide();
						}
						else fd.field('Submit').value = true;
					}
					
		            fd.save();
				 }
			 }
	     }
      });
}

var isUserAllowed = async function(userName){
	var _isAllowed = false
	await pnp.sp.web.currentUser.get()
		 .then(async function(user){
			if(user.Title == userName){
				_isAllowed = true;
			}
		});	
		
		if(!_isAllowed)
		{
			var userId = _spPageContextInfo.userId;
            var userGroups = [];
			await pnp.sp.web.siteUsers.getById(userId).groups.get()                               
			.then(async function(groupsData){                        
				for (var i = 0; i < groupsData.length; i++) {
					userGroups.push(groupsData[i].Title);
				}
				if(userGroups.indexOf('Owners') >= 0)
				  _isAllowed = true;
			});
		}

		return _isAllowed;	
}
 

function setErrorMessage(err, textButton, isRemove){
	
	if(!isRemove)
	{
		var errLength = $('div.alert').length;
		if(errLength == 0){
			var appendCtrl = "<div role='alert' class='alert alert-danger'>" +
								"<button type='button' data-dismiss='alert' aria-label='Close' class='close'>" +
								"<span aria-hidden='true'>×</span>" +
								"</button>" +
								"<h4 id='errId' class='alert-heading'>Please correct the errors below:</h4>" +
							"</div>";

			var masterErrorPosition = $('div.container-fluid').first();
			masterErrorPosition.prepend(appendCtrl);
		}
		
		var elem = $('#errId').find("p:contains('" + err + "')");
		if(elem.length == 0){
			$('#errId').append("<p style='font-size: 14px; font-family: Arial; padding-top:10px'>" + err + "</p>");
			if(textButton != ""){
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
			var elem = $('#errId').find("p:contains('" + err + "')");
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

var ApplyFieldChanges = async function(fields, disableControls, disableCustomButtons, defaultButtons, _status){
	 for(var i = 0; i < fields.length; i++){
		 if(fields[i] == "Attachments")
		 {
		   var isAllowed = await IsUserInGroup("Owners");
		   if(!isAllowed)
		   {		   
			 $('div.k-upload-button').remove();
			 $('button.k-upload-action').remove();
		   }
		   else {
			   $("div.k-upload-sync").removeClass("k-state-disabled");
			   //if(_status == "Completed" || _status == "Closed" || _status == "Issued to Contractor")
			     await customButtons("Save", "Save Attachment", true, "U");
		   }
		 }
		 
		else try { fd.field(fields[i]).disabled = disableControls;} catch(e){alert("fields[i] = " + fields[i] + "<br/>" + e);}
	 }

	 if(disableCustomButtons)
	 {
	   	if(defaultButtons != null && defaultButtons != 'undefined')
			fd.toolbar.buttons[0].disabled = true;
		
		$('span').filter(function(){ return $(this).text() == 'Save'; }).parent().attr("disabled", "disabled");
		$('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().attr("disabled", "disabled");
	 }
}

var showHideTabs = async function(key, BIC, isPart, isLead){
	var _isAllowed = false;
    try
    {
        // get all groups the current user belongs to
        var userId = _spPageContextInfo.userId;
        var userGroups = [];
        await pnp.sp.web.siteUsers.getById(userId).groups.get()                               
        .then(async function(groupsData){                        
			for (var i = 0; i < groupsData.length; i++) {
				userGroups.push(groupsData[i].Title);
			}
								
			var isAdmin = false;
			var Status = fd.field('Status').value;
			
			if(userGroups.indexOf('Owners') >= 0) isAdmin = true;
											
			var isClosed = false;
			if(Status == 'Issued to Contractor' || Status == 'Completed') isClosed = true;
												
			var isDar = false;                                                     
			if(userGroups.indexOf('SiteTeam') >= 0) isDar = true;
			if(userGroups.indexOf('PMC') >= 0) isDar = true;  
					
			if(isAdmin && isClosed)
				fd.field('Status').disabled = false;
			else if(isClosed)
				fd.field('Status').disabled = true;
			
			if( (isClosed) || (!isDar && !isAdmin))
			{                                                                      
				fd.container('Tab1').tabs[1].disabled = true;     
                $('span').filter(async function(){ return $(this).text() == 'Assign'; }).parent().attr("disabled", "disabled");				
				return;
			}   
            else
			{
				if(Status == "Open" || Status == "Assigned" || Status == "Reassigned" || Status == "Send to PMC")
				{
					if(key == "MAT" || key == "SCR")
					{
						//var spGroup = "RE";                      
					    if( (BIC == "Site" && userGroups.indexOf('RE') >= 0) || (BIC == "Office" && Status == "Send to PMC" && userGroups.indexOf('PMC') >= 0))
						{
							fd.container('Tab1').tabs[1].disabled = false; 
					       _isAllowed = true;
						}
						else{
							if(isPart && fd.field('Trade').value == "PMC" && userGroups.indexOf('PMC') >= 0){
								fd.container('Tab1').tabs[1].disabled = false; 
							   _isAllowed = true;
							} 
							else{
								fd.container('Tab1').tabs[1].disabled = true; 
							   _isAllowed = false;
							}
						}
					}
					else
					{
					   //fd.container('Tab1').tabs[1].disabled = false; 
					   _isAllowed = true;
					}
				}
                //else fd.container('Tab1').tabs[1].disabled = true; 
                $('span').filter(async function(){ return $(this).text() == 'Assign'; }).parent().prop("disabled", false);
               
				if(isPart || isLead){
					var listName;
					var mType = await getMajorType(key);
					if(mType !=null && mType.length > 0)
						listName = mType[0].MatrixList;
					else listName = await getParameter("Inspection-Matrix");

					var query = pnp.sp.web.lists.getByTitle(listName).items.select("Title");
						await query.orderBy("Title", true)
							.get()
							.then(async function (items) { 
									var TradeArray = [];
									var leadArray = [];
									for(var i = 0; i < items.length; i++)
									{
										var trade = items[i].Title;
										if(isPart && trade == "PMC"){}
										else TradeArray.push(trade);
										
										if((key == "MAT" || key == "SCR") && trade != "PMC"){
											leadArray.push(trade);
										}
										
										if(listName == "Trades")
										{
											var Permission = items[i].Permission;
											if(Permission != null && Permission.length > 0)
											{
												for(var j = 0; j < Permission.length; j++)
												{
													var group = Permission[j];
													var isUserValid = await IsUserInGroup(group);
													//console.log("isUserValid = " + isUserValid);
													if(isUserValid)
													{
														if(isPart && trade == "PMC"){}
														else TradeArray.push(trade);
														break;
													}
												}
											}
										}
									}
									await setTrades(TradeArray, leadArray, key, isLead);
							}); 	
				}   
			}				
        });
		return _isAllowed;
    }
    catch(err){
        alert(err.message);
    }
}

function DisableFields (filterColumn, filterValue, operator, fields, disableControls, disableCustomButtons, defaultButtons){
	 try
	 {
	   var _status = fd.field(filterColumn).value;
	   if(_status == "") _status = "Initiated";
		   
	   if(operator == "eq")
	   {
			if(_status == filterValue)
			   ApplyFieldChanges(fields, disableControls, disableCustomButtons, defaultButtons, _status);
	   }

	   else
	   {
		  if(_status != filterValue)
		    ApplyFieldChanges(fields, disableControls, disableCustomButtons, defaultButtons, _status);
	   }
	 }
	catch(e){ console.log(e); }
}

function HideFields(fields, hideButtons, isHide){
	if(hideButtons)
	{
	  fd.toolbar.buttons[0].style = "display: none;";
      fd.toolbar.buttons[1].style = "display: none;";
	}
	
	for(var i = 0; i < fields.length; i++)
	{
		if(isHide || isHide == undefined)
		  $(fd.field(fields[i]).$parent.$el).hide();
		else $(fd.field(fields[i]).$parent.$el).show();
	}
}
  
function PrintCustom(){
	debugger;
	  var PageHtml = $('div.canvasWrapper_79a31d19').html().replace('inline-block','block');
	  var printpageHTML = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">\r\n <HTML><HEAD>\n' + document.getElementsByTagName('HEAD')[0].innerHTML + '</HEAD>'
					  '\n<BODY>\n' + PageHtml + '\n</BODY></HTML>';
	  console.log(printpageHTML);
	  
	  var printWindow = window.open('','printWindow');
	  printWindow.document.open();
	  printWindow.document.write(printpageHTML);
	  printWindow.document.close();
	  //printWindow.focus();

	  printWindow.print();
	  printWindow.close();
}

var setTrades = async function (array, leadArray, moduleName, isLead){
	fd.field('Part').widget.dataSource.data(array);
	 
	var partArray = [];
	var Trade = "", PartTrades = "";
	PartTrades = fd.field('PartTrades').value;

	if(isLead){
		Trade = fd.field('Trade').value;
		if(Trade != null & Trade != ""){
			if(PartTrades != null & PartTrades != "")
				PartTrades = Trade + "," + PartTrades;
			else PartTrades = Trade;
		}
	}
	
   if(PartTrades != null)
   {
	  if(PartTrades.includes(","))
	  {
        PartTrades = PartTrades.split(',');
		for(var i = 0; i < PartTrades.length; i++)
		{
			partArray.push(PartTrades[i]);
		}
	  }
	  else partArray.push(PartTrades);
      fd.field('Part').value = partArray;
   }
   
   if(moduleName == "MAT" || moduleName == "SCR"){
	   fd.field('Lead').widget.dataSource.data(leadArray);
	   var _leadTrade = fd.field('LeadTrade').value;
	   if(_leadTrade != null)
	   {
	     var leadObject = [];
	     leadObject.push(_leadTrade);
	     fd.field('Lead').value = leadObject;
	   }
   } 
}

function setDefaultTrades(items){
   var array = [];
   array.push("");
   for(var i = 0; i < items.length; i++)
   {
     array.push(items[i].Title);
   }
   setTrades(array);
}
 
function setOldRef(_status){
	 fd.field('ORFI').ready().then(function() {
		//fd.field('ORFI').filter = "Reference ne null";  
		fd.field('ORFI').filter = "Reference ne null  and Status eq '" + _status + "' and IsLatestRev eq '1'";  
		fd.field('ORFI').orderBy = { field: 'Reference', desc: false };
		fd.field('ORFI').refresh();
    });
}

function checkTradeAssignment(columnName){
	fd.field(columnName).$on('change', function(value) {   
		debugger;
		if(value == "")
		{
			if( (fd.field('Lead').value == null || fd.field('Lead').value.length == 0) && (fd.field('Part').value == null || fd.field('Part').value.length == 0)) {
				fd.field(columnName).clear();
				$('span').filter(function(){ return $(this).text() == 'Send to PMC'; }).parent().removeAttr('disabled');
			  }
			return;
		} 
 
		var isLeadNull = isFieldNULL("Lead");
		var isPartNull = isFieldNULL("Part");

		  if( !isLeadNull ||  !isPartNull)
			  $('span').filter(function(){ return $(this).text() == 'Send to PMC'; }).parent().attr("disabled", "disabled");
		  else if(isLeadNull && isPartNull) {
			fd.field(columnName).clear();
			$('span').filter(function(){ return $(this).text() == 'Send to PMC'; }).parent().removeAttr('disabled');
		  }

		  var leadError = "Lead Trade is Required to be filled";
		  if(fd.field('Lead').value == null)
			setErrorMessage(leadError, "Assign", false);
		  else setErrorMessage(leadError, "Assign", true);

		  if(fd.field('Lead').value != null && fd.field('Part').value != null)
		  {
			 var _lead = fd.field('Lead').value;
			 var _part = fd.field('Part').value;

				if(_part != null && _part.length > 0)
				{
					var isError = false;
					for(var i = 0; i < _part.length; i++)
					{
						if(_part[i] == _lead)
						{
							isError = true;
							break;
						}
					}
					var partError = "Can not assign same trade as lead and part.";
					if(isError)
						setErrorMessage(partError, "Assign", false);
					else setErrorMessage(partError, "Assign", true);
				}
			}
	 }); 
}

function isFieldNULL(columnName){
	var isNULL = true;
	try{
	  if (fd.field(columnName).value.length > 0)
	    isNULL = false
	}
	catch{}
	return isNULL;
}

fd.spBeforeSave(function ()
{
	preloader();
});
