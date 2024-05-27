var _layout = "/_layouts/15/PCW/General/EForms";

var _modulename = "", _formType = "";

var _OutProceed = 0, _Proceed = 0;

var _CORDirection = "";
var ProposalNumber = "";
var FinalVal = "";
var _isManualEntry = false;

var saveMesg = "Click 'Save' to store your progress and keep your work as a draft.";
var submitMesg = "Click 'Submit' to finalize and send officially.";
var cancelMesg = "Click 'Cancel' to discard changes and exit without saving.";

var onRender = async function (moduleName, formType, relativeLayoutPath){
	//localStorage.clear();	
	try {	

		if(relativeLayoutPath !== undefined && relativeLayoutPath !== null && relativeLayoutPath !== '')
		_layout = relativeLayoutPath;		

		await PreloaderScripts();		
		await loadScripts();
		fixTextArea();

		_modulename = moduleName;
		_formType = formType;
		if(moduleName == 'ALGCOR')
			await onALGCORRender(formType);

		if(formType == 'New')
			await setButtonToolTip('Save', saveMesg);
		await setButtonToolTip('Submit', submitMesg);			
		await setButtonToolTip('Cancel', cancelMesg);	
		
		preloader("remove");
	}
	catch (e) {
		alert(e);
		console.log(e);		
		preloader();
	}
}

var onALGCORRender = async function (formType){	

	if(formType == 'New'){
		await ALGCOR_newForm();       	
    }
    else if(formType == 'Edit'){
		await ALGCOR_editForm();        
    }    
    else if(formType == 'Display'){ 
		await ALGCOR_displayForm();     
    }
}

fd.spBeforeSave(function()
{     
	if(_modulename == "ALGCOR")
        AttachFiles();
	
	return fd._vue.$nextTick();
});

var ALGCOR_newForm = async function(){		

	//fd.clear();		
	
	fd.validators.push({
		name: 'Check Attachment',
		error: "Please upload the softcopy",
		validate: function(value) {		
			if(fd.field('Attachments').value.length == 0)
			{
				return false; 	
			}	
			return true;
		}
	});
	
	fd.validators.push({
		name: 'Check Folder',
		error: "Please select a project folder",
		validate: function(value) {				
			if(fd.field('ProjectNumberFolder').value == null || fd.field('ProjectNumberFolder').value == "")
			{			
				return false; 	
			}
			return true;
		}
	});	

	fd.toolbar.buttons[0].style = "display: none;";
	fd.toolbar.buttons[1].style = "display: none;";

	$(fd.field('Submit').$parent.$el).hide();

	$(fd.field('From').$parent.$el).hide();
	$(fd.field('To').$parent.$el).hide();

	fd.field('LookupFrom').required = true;
	fd.field('LookupTo').required = true;

	fd.field('FromFilter').required = true;
	fd.field('ToFilter').required = true;
		
	const items = await pnp.sp.web.lists.getByTitle("Categories").items.select("Category").orderBy("Category").getAll();
	const distinctValues = [...new Set(items.map((item) => item.Category))];
	fd.field('FromFilter').widget.setDataSource(distinctValues); 
	fd.field('ToFilter').widget.setDataSource(distinctValues);

	fd.field('FromFilter').$on('change', async function(value) {	
			await pnp.sp.web.lists.getByTitle("Categories").items.select("Title").filter("Category eq '" + value + "'").orderBy("Category").get().then(async function(items) {		
			await fd.field('LookupFrom').widget.setDataSource({data: items.map(function(i) { return i.Title })
			});
		});		
	});

	fd.field('ToFilter').$on('change', async function(value) {	
			await pnp.sp.web.lists.getByTitle("Categories").items.select("Title").filter("Category eq '" + value + "'").orderBy("Category").get().then(async function(items) {		
			await fd.field('LookupTo').widget.setDataSource({data: items.map(function(i) { return i.Title })
			}); 
		});
	});

	fd.field('LookupFrom').$on('change', function(value) {
		$(fd.field('From').$parent.$el).show();	
		fd.field('From').value = value;
		$(fd.field('From').$parent.$el).hide();	
	});

	fd.field('LookupTo').$on('change', function(value) {
		$(fd.field('To').$parent.$el).show();	
		fd.field('To').value = value;
		$(fd.field('To').$parent.$el).hide();
	});

	fd.toolbar.buttons.push({
			icon: 'Save',
			class: 'btn-outline-primary',
			text: 'Save',
			click: async function() {

			fd.validators.length = 0; 
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
				fd.field('AttachFiles').value = "A" + "," + NCountofATT + "," + OCountofATT;
				
				await updateCounter();
					
				fd.save();
			}
		}
	});
		
	fd.toolbar.buttons.push({
			icon: 'Accept',
			class: 'btn-outline-primary',
			text: 'Submit',
			click: async function() {
			
			if(!fd.isValid)
			{
			$(fd.field('AttachFiles').$parent.$el).hide();
			$(fd.field('Submit').$parent.$el).hide();
			}
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
				fd.field('AttachFiles').value = "A" + "," + NCountofATT + "," + OCountofATT;
				
				$(fd.field('Submit').$parent.$el).show();
				fd.field('Submit').value = true;
				
				await updateCounter();
				
				fd.save();
			}
		}
	});
		
	fd.toolbar.buttons.push({
		icon: 'ChromeClose',
		class: 'btn-outline-primary',
		text: 'Cancel',
		click: function() {
	
		fd.validators.length = 0; 			
		fd.close();
	}
	});

	fd.field('CORDirection').value = true;

	fd.field('ReservedRef').disabled = false;
	fd.field('ReservedRef').value = "";
	fd.field('ReservedRef').disabled = true;

	fd.field('ProjectNumberFolder').disabled = true;

	var queryString = window.location.search;
	var urlParams = new URLSearchParams(queryString);
	var rfVal = "";
	try{
		rfVal = urlParams.get('rf').split("/")[5];
	}
	catch(err)
	{
		alert("Please choose a Project folder First.");
		fd.close();
	}
		
	fd.field('ProjectNumberFolder').disabled = false;
	fd.field('ProjectNumberFolder').value = rfVal;
	fd.field('ProjectNumberFolder').disabled = true;

	$(fd.field('ProposalNumber').$parent.$el).hide();

	if(rfVal == "Proposals")
	{
		$(fd.field('ProposalNumber').$parent.$el).show();
		fd.field('ProposalNumber').required = true;
		ProposalNumber = fd.field('ProposalNumber').value;	
	}	

	const LetDate = fd.field('Date').value;
	const year = LetDate.getFullYear();
	FinalVal = year.toString().slice(-2);

	var Stage = "";

	fd.field('CORDirection').$on('change', async function(value) {
		CORDirectionhideOrShow();
	});
	CORDirectionhideOrShow();

	fd.field('Date').$on('change', async function(value) { 
		
		FinalVal = value.getFullYear().toString().slice(-2);
		
		if(_CORDirection === "Out")
			Stage = "A" + "-" + FinalVal;
		else
			Stage = "IN" + "-" + FinalVal;
				
		var camlF = "Title" + " eq '" + Stage + "'";
		
		await pnp.sp.web.lists.getByTitle('Counter').items.select("Title", "Counter").filter(camlF).get().then(function(items){
						
			if(fd.field('Stage').value != '')
			{ 			
				AutoReference(items, rfVal, ProposalNumber, fd.field('Stage').value, FinalVal);
			}
			else
			{
				fd.field('ReservedRef').disabled = false;				                    
				fd.field('ReservedRef').value = "";	
				fd.field('ReservedRef').disabled = true;
			}	    
			
		});
	});

	fd.field('Stage').$on('change', async function(value){	
		
		if(_CORDirection === "Out")
			Stage = "A" + "-" + FinalVal;
		else
			Stage = "IN" + "-" + FinalVal;
			
		var camlF = "Title" + " eq '" + Stage + "'";
		
		await pnp.sp.web.lists.getByTitle('Counter').items.select("Title", "Counter").filter(camlF).get().then(function(items){
			AutoReference(items, rfVal, ProposalNumber, value, FinalVal);
		});
	});

	if(rfVal == "Proposals")
	{
		fd.field('ProposalNumber').$on('change', async function(value) {
		
			if(_CORDirection === "Out")
				Stage = "A" + "-" + FinalVal;
			else
				Stage = "IN" + "-" + FinalVal;	
						
			var camlF = "Title" + " eq '" + Stage + "'";
			await pnp.sp.web.lists.getByTitle('Counter').items.select("Title", "Counter").filter(camlF).get().then(function(items){		
				AutoReference(items, rfVal, value, fd.field('Stage').value, FinalVal);	        
			});
		});
	}

	fd.field('RelatedLetters').ready().then(function() {
		fd.field('RelatedLetters').filter = "RefNo ne null";  
		fd.field('RelatedLetters').orderBy = { field: 'Title', desc: false };
		fd.field('RelatedLetters').refresh();
	});

	$(fd.field('AttachFiles').$parent.$el).hide();

	fd.field('Confidential').value = false;
	fd.field('Confidential').$on('change', function(value) {
		hideOrShow();
	});

	var previousValue = fd.field('ReservedRef').value ;
	fd.field('ReservedRef').$on('change', async function(value) {

		if (value !== previousValue && previousValue !== '') {
			await CheckDuplicateReference(value);
			previousValue = value;
		}
		else
			previousValue = value;				
	});	
}

var ALGCOR_editForm = async function(){

	fd.validators.push({
		name: 'Check Attachment',
		error: "Please upload the softcopy",
		validate: function(value) {		
			if(fd.field('Attachments').value.length == 0)
			{
				return false; 	
			}	
			return true;
		}
	});

	$(fd.field('Status').$parent.$el).show();
	var Status = fd.field('Status').value;
	$(fd.field('Status').$parent.$el).hide();

	fd.toolbar.buttons[0].style = "display: none;";
	fd.toolbar.buttons[1].style = "display: none;";
	$(fd.field('Submit').$parent.$el).hide();

	fd.toolbar.buttons.push({
				icon: 'Accept',
				class: 'btn-outline-primary',
				text: 'Submit',
				click: function() {
				
				if(!fd.isValid)
				{
				$(fd.field('AttachFiles').$parent.$el).hide();
				$(fd.field('Submit').$parent.$el).hide();
				}
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
					fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
					
					$(fd.field('Submit').$parent.$el).show();
					fd.field('Submit').value = true;
					fd.save();
				}
			}
		});
		
	fd.toolbar.buttons.push({
			icon: 'ChromeClose',
			class: 'btn-outline-primary',
			text: 'Cancel',
			click: function() {

			fd.validators.length = 0; 		 
			fd.close();
		}
	});

	$(fd.field('AttachFiles').$parent.$el).hide();

	await CustomListEditor(Status);
}

var ALGCOR_displayForm = async function(){
	$(fd.field('AttachFiles').$parent.$el).hide();
}

//#region Utilities
function AttachFiles()
{
    var NCountofATT = 0;
	var OCountofATT = 0;

	for(i = 0; i < fd.field('Attachments').value.length; i++) 
	{
		var Val = fd.field('Attachments').value[i].extension.toString();
		if(Val === "")
		{OCountofATT++;}
		else
		  NCountofATT++;			
	}
	$(fd.field('AttachFiles').$parent.$el).show();
    if(_formType == "New")
	  fd.field('AttachFiles').value = "A" + "," + NCountofATT + "," + OCountofATT;
    else fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;	
    $(fd.field('AttachFiles').$parent.$el).hide();
}

function SetAttachmentToReadOnly(){
	fd.field('Attachments').disabled = false;

	var spanATTDelElement = document.querySelector('.k-upload .k-upload-files .k-upload-status');
	if(spanATTDelElement !== null)
	{
		spanATTDelElement.style.display = 'none';
		
		var spanATTUpElement = document.querySelector('.k-upload .k-upload-button');
		spanATTUpElement.style.display = 'none';
		
		var spanATTZoneElement = document.querySelector('.k-dropzone');
		if(spanATTZoneElement !== null)
			spanATTZoneElement.style.display = 'none';
	}
	else
		DisableAttachment();
}

function delay(time) {
	return new Promise(function(resolve) { 
		setTimeout(resolve, time)
	});
}

function fixTextArea(){
	$("textarea").each(function(index){
		var height = (this.scrollHeight + 5) + "px";
        $(this).css('height', height);
	});
}

function hideOrShow() {

	if (fd.field('Confidential').value == true) 	
		fd.field('ToEmailUsers').required = true;			
	
	else 
		fd.field('ToEmailUsers').required = false;		
}

function CORDirectionhideOrShow() {

	if (fd.field('CORDirection').value == true) {
	
		$(fd.field('RefNo').$parent.$el).show();
        fd.field('RefNo').value = "";
		fd.field('RefNo').required = false;
		$(fd.field('RefNo').$parent.$el).hide();					
		$(fd.field('ReservedRef').$parent.$el).show();
		fd.field('ReservedRef').value = "";	
		fd.field('ReservedRef').disabled = true;
		
		_CORDirection = "Out";	
	} 
	else {
	
		$(fd.field('RefNo').$parent.$el).show();
        fd.field('RefNo').value = "";
		fd.field('RefNo').required = false;
		$(fd.field('RefNo').$parent.$el).hide();					
		$(fd.field('ReservedRef').$parent.$el).show();
		fd.field('ReservedRef').value = "";	
		fd.field('ReservedRef').disabled = true;
		
		_CORDirection = "In";																
	}
}

function WriteFromBoxToRef() {
	$(fd.field('RefNo').$parent.$el).show();
	fd.field('RefNo').value = fd.field('ReservedRef').value;  
	$(fd.field('RefNo').$parent.$el).hide();  
}

function AutoReference(items, rfVal, ProposalNumber, Stage, FinalVal) {

	fd.field('ReservedRef').disabled = false;
	var ReservedRef = "";
	
	var CountNumber = 1;
	
	if (items.length == 1)    
		CountNumber = items[0].Counter;
		
	CountNumber = String(CountNumber).padStart(4,'0');
	if(rfVal == "Proposals" && _CORDirection === "Out")
	{
		ProposalNumber = fd.field('ProposalNumber').value;
		ReservedRef = "PA" + ProposalNumber + "/" + Stage + "/" + "A" + CountNumber + "/" + FinalVal;	
		fd.field('ReservedRef').value = ReservedRef;
	}
	else {
		
		if(_CORDirection === "Out")
			ReservedRef = rfVal + "/" + Stage + "/" + "A" + CountNumber + "/" + FinalVal;
		else
			ReservedRef = rfVal + "/IN/" + "A" + CountNumber + "/" + FinalVal;
			
		fd.field('ReservedRef').value = ReservedRef;
	}	
		
	DisableReservedRef(FinalVal);				
}

async function CheckDuplicateReference(Reference){

	var isRecordExist = false;
	
	var camlF = "RefNo eq '" + Reference + "'";
	
	await pnp.sp.web.lists.getByTitle('Correspondences').items.select("Id").filter(camlF).get().then(function(items)
	{
		if(items.length > 0)
			isRecordExist = true;
		else
			isRecordExist = false;		 
	}); 
 
	if(isRecordExist)
	{		
		alert("Apologies, Reference you provided already exists.");
		$('span').filter(function(){ return $(this).text() == 'Save'; }).parent().attr("disabled", "disabled");
		$('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().attr("disabled", "disabled");		
	}
	else
	{	
		$('span').filter(function(){ return $(this).text() == 'Save'; }).parent().removeAttr('disabled');
		$('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().removeAttr('disabled');	
		
		$(fd.field('RefNo').$parent.$el).show();
		fd.field('RefNo').value = Reference;  
		$(fd.field('RefNo').$parent.$el).hide();
		
		_isManualEntry = true;
    }
}

async function updateCounter() {

	var value = 1;
	var Stage = "";

	if(_isManualEntry) {

		var Reference = "";		
		Reference = fd.field('ReservedRef').value; 

		var arr = Reference.split('/');
		var seq = arr[2].toString();	

		if (seq.indexOf('In') !== -1) {
			value = seq.replace('In', '');
		}
		else
			value = seq.replace('A', '');
	}

	if(_CORDirection === "Out")
		Stage = "A" + "-" + FinalVal;
	else
		Stage = "IN" + "-" + FinalVal;	
				
	var camlF = "Title" + " eq '" + Stage + "'";
	
	var listname = 'Counter';
	
	await pnp.sp.web.lists.getByTitle(listname).items.select("Id, Title, Counter").filter(camlF).get().then(async function(items){
		var _cols = { };
        if(items.length == 0){
             _cols["Title"] = Stage;
			 value = parseInt(value) + 1;
             _cols["Counter"] = value.toString();                 
             await pnp.sp.web.lists.getByTitle(listname).items.add(_cols);                 
        }
          else if(items.length > 0){

            var _item = items[0];

			if(_isManualEntry) 
				value = parseInt(value) + 1;
			else            
				value = parseInt(_item.Counter) + 1;

            _cols["Counter"] = value.toString();                   
            await pnp.sp.web.lists.getByTitle(listname).items.getById(_item.Id).update(_cols); 
		}                   
         
    });	
}

function DisableReservedRef(FinalVal){

	if(FinalVal !== null || FinalVal !== '')
	{
		const number = parseInt(FinalVal);
		
		if(number < 24)
		{
			fd.field('ReservedRef').disabled = false;
			fd.field('ReservedRef').$on('change', WriteFromBoxToRef());
		}
		else
		{
			fd.field('ReservedRef').disabled = true;
			WriteFromBoxToRef();
		}
	}
	else
		fd.field('ReservedRef').disabled = true;
}

function CustomListEditor(Status) {				

	if (Status == 'Closed') 
	{		
		fd.field('ResponseDate').disabled = true;	
		fd.field('Response').disabled = true;	
		fd.field('SentToSBG').disabled = true;
		fd.field('Attachments').disabled = true;
	}	
	else if(Status == 'Initiated')
	{
		$(fd.field('SentToSBG').$parent.$el).hide();
	}	
}
//#endregion

var loadScripts = async function(){
	const libraryUrls = [
		//_layout + '/controls/preloader/jquery.dim-background.min.js',
		_layout + "/plumsail/js/customMessages.js",
		_layout + '/controls/tooltipster/jquery.tooltipster.min.js',
		//_layout + '/plumsail/js/preloader.js',
		_layout + '/plumsail/js/utilities.js'
	];
  
	const cacheBusting = `?v=${Date.now()}`;
	  libraryUrls.map(url => { 
		  $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
		});
		
	const stylesheetUrls = [
		_layout + '/controls/tooltipster/tooltipster.css',
		_layout + '/plumsail/css/CssStyle.css',
		_layout + '/plumsail/css/CssStyleRACI.css'
	];
  
	stylesheetUrls.map((item) => {
	  var stylesheet = item;
	  $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
	});
}

var PreloaderScripts = async function(){
    await _spComponentLoader.loadScript(_layout + '/controls/preloader/jquery.dim-background.min.js');
    await _spComponentLoader.loadScript(_layout + '/plumsail/js/preloader.js');
    preloader();
}