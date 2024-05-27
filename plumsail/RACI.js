var _layout = "/_layouts/15/PCW/General/EForms";

var _modulename = "", _formType = "";
var countryURL = 'https://4ce.dar.com/sites/CountryCodes';

var _OutProceed = 0, _Proceed = 0;

var BASEURL = location.protocol + "//" + location.host;
var SERVICE_URL = BASEURL + "/AjaxService/XMLService.asmx?op=CallCountryCodesService";
var International_SERVICE_URL = BASEURL + "/AjaxService/XMLService.asmx?op=CallInternationalCountryCodesService";
var CodeArr = [];
const CodeObjectArray = [];
var PackageArr = [];

var saveMesg = "Click 'Save' to store your progress and keep your work as a draft.";
var submitMesg = "Click 'Submit' to finalize and send officially.";
var cancelMesg = "Click 'Cancel' to discard changes and exit without saving.";

var onRender = async function (moduleName, formType, relativeLayoutPath){

	try {	

		if(relativeLayoutPath !== undefined && relativeLayoutPath !== null && relativeLayoutPath !== '')
		_layout = relativeLayoutPath;

		await PreloaderScripts();		
		await loadScripts();
		
		fixTextArea();

		_modulename = moduleName;
		_formType = formType;
		if(_formType === 'New'){
			clearStoragedFields(fd.spForm.fields);
		}
		if(moduleName == 'CL')
			await onCLRender(formType);
		else if(moduleName == 'RR')
			await onRRender(formType);
		else if(moduleName == 'DCR')
			await onDCRender(formType);
		else if(moduleName == 'TQ')
			await onTQRender(formType);
		else if(moduleName === "DDSInitiate")
			await onDDSRender(formType);

		await setButtonToolTip('Save', saveMesg);
		await setButtonToolTip('Submit', submitMesg);
		await setButtonToolTip('Submit for Approval', submitMesg);
		await setButtonToolTip('Cancel', cancelMesg);

		preloader("remove");
	}
	catch (e) {
		alert(e);
		console.log(e);		
		preloader();
	}
}

var onCLRender = async function (formType){
    if(formType == 'Edit'){
        fd.field('WorkflowStatus').disabled = true; 
        $(fd.field('Reviewed').$parent.$el).hide();
        $(fd.field('AttachFiles').$parent.$el).hide();
        $(fd.field('ClaimAttachment').$parent.$el).hide();
        $(fd.field('ReviewedByPD').$parent.$el).hide();
        $(fd.field('ReviewedByAD').$parent.$el).hide();

        // addLegend();        
        
        var currentUser = await pnp.sp.web.currentUser.get();

        let ChangeLogPMFullPermission = "no";
        ChangeLogPMFullPermission = await getParameterRACI("ChangeLogPMFullPermission");
        
        if(ChangeLogPMFullPermission.toLowerCase() == "yes") 
            await showHideFields("yes", currentUser.Title);
        else				
			await showHideFields("no", currentUser.Title);
    }
    else if(formType == 'New'){

        $(fd.field('AttachFiles').$parent.$el).hide();	
        $(fd.field('Reviewed').$parent.$el).hide();
        $(fd.field('WorkflowStatus').$parent.$el).hide();
        
        fd.toolbar.buttons[0].style = "display: none;";
        fd.toolbar.buttons[1].icon = 'Cancel';
        fd.toolbar.buttons[1].text = "Cancel";

        // addLegend();
        
        // var currentUser = await pnp.sp.web.currentUser.get();
        // fd.field('RequestedBy').value = currentUser.Title;
        // fd.field('RequestedBy').disabled = true;
		var RedirectRule = "Dont Redirect";

		fd.spSaved(function(result) {			
			try
			{		
				if(RedirectRule === "Redirect")
				{
					var listId = fd.spFormCtx.ListAttributes.Id;
					var itemId = result.Id;						
					var webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
					
					result.RedirectUrl = webUrl + "/SitePages/PlumsailForms/Change%20Log/Item/EditForm.aspx?item=" + itemId;	
				}					
			}
			catch(e){alert(e);}								 
		 });
		 

        fd.toolbar.buttons.push({
                icon: 'Save',
                class: 'btn-outline-primary',
                text: 'Save',
                click: function() {                     
	
                $(fd.field('Reviewed').$parent.$el).show();
                fd.field('Reviewed').value = false;
                $(fd.field('Reviewed').$parent.$el).hide();						
				RedirectRule = "Redirect";
                fd.save();			 
            }
        });
    
        fd.toolbar.buttons.push({
                    icon: 'Accept',
                    class: 'btn-outline-primary',
                    text: 'Submit',
                    click: function() {				
                    
                    WorkflowStatus = "Submitted to PD";	
                    
                    $(fd.field('WorkflowStatus').$parent.$el).show();				
                    fd.field('WorkflowStatus').value = WorkflowStatus;
                    $(fd.field('WorkflowStatus').$parent.$el).hide();
                        
                    $(fd.field('Reviewed').$parent.$el).show();
                    fd.field('Reviewed').value = true;	
                    $(fd.field('Reviewed').$parent.$el).hide();	
                    
                    fd.save();
                }	     
        });		
    }
    if(formType == 'Display'){        
        showHideFieldsDisplay();
    }
}

var onRRender = async function (formType){
    try
	{
        if(formType == 'New'){    

         var prob = "<table cellspacing='0' cellpadding='0' width='100%' border='1' style='border-collapse:collapse; font-size:11px;'>" +
         "<tr><td width='10%' style='text-align:center; background-color:lightgrey; font-weight:bold'>Rating</td><td width='20%' style='background-color:lightgrey; font-weight:bold;text-align:center'>Description</td><td style='background-color:lightgrey; font-weight:bold'>Definition</td></tr>" +
         "<tr><td style='text-align:center'>5</td><td style='text-align:center'>Almost certain</td><td>>81% probability or<br />Could occur within days or<br />Event could occur several times during the life of the facility or activity</td></tr>" +
         "<tr><td style='text-align:center'>4</td><td style='text-align:center'>Likely</td><td>>61% to 80% probability or<br />Could occur within a couple of weeks or<br />Event has occurred several times per year in similar industries OR more than once per year in Dar’s experience</td></tr>" +
         "<tr><td style='text-align:center'>3</td><td style='text-align:center'>Possible</td><td>>21 to 60% probability or<br />Could occur within a month or<br />Event has occurred in Dar’s experience</td></tr>" +
         "<tr><td style='text-align:center'>2</td><td style='text-align:center'>Unlikely</td><td>>11 to 20% probability or<br />Could occur in a couple of months or<br />Event occurrence is possible but unlikely or has happened a few times in similar industries.</td></tr>" +
         "<tr><td style='text-align:center'>1</td><td style='text-align:center'>Rare</td><td>>0 to 10% probability or<br />Could occur in more than a few months or<br />Event is physically possible but has never been heard of in similar industries.</td></tr>" +
        "</table>";
        $('#prob').append(prob);

        var impact = "<table cellspacing='0' cellpadding='0' width='100%' border='1' style='border-collapse:collapse; font-size:11px;'>" +
                    "<tr><td width='10%' style='text-align:center; background-color:lightgrey; font-weight:bold'>Rating</td><td width='20%' style='background-color:lightgrey; font-weight:bold;text-align:center'>Description</td><td style='background-color:lightgrey; font-weight:bold'>Definition</td></tr>" +
                    
                    "<tr><td style='text-align:center'>5</td><td style='text-align:center'>Critical</td>" +
                        "<td>Impact on project schedule: Deviation 16% or<br />Impact on project costs (direct man hours): deviation over 60% or<br />Reputation: Risk covered on international TV and newspapers or<br />" +
                            "People: Multiple fatalities, Asset / Property or<br />Asset / Property / Revenue: Loss > $10M,<br />External Effect: Major and sustained  external pollution</td></tr>" +

                    "<tr><td style='text-align:center'>4</td><td style='text-align:center'>Major</td>" +
                        "<td>Impact on project schedule: Deviation 12% or<br />Impact on project costs (direct man hours): deviation between 40%-59% or<br />Reputation: Risk covered on regional TV and press or<br />" +
                            "People: Multiple LTIs OR one or more Permanent Disability OR 1 Fatality or<br />Asset / Property / Revenue: Loss between $1M – $10M or<br />External Effect :Significant pollution with reversible environmental consequences</td></tr>" +
                    
                        "<tr><td style='text-align:center'>3</td><td style='text-align:center'>Moderate</td>" +
                        "<td>Impact on project schedule: Deviation 8% or<br />Impact on project costs (direct man hours): deviation between 20%-39% or<br />Reputation: Risk covered on local TV and press or<br />" +
                            "People: Single LTI OR multiple RWDC or<br />Asset / Property / Revenue: Loss between $100K – $1M or<br />External Effect: Significant pollution having external impact</td></tr>" +
                    
                    "<tr><td style='text-align:center'>2</td><td style='text-align:center'>Minor</td>" +
                        "<td>Impact on project schedule: Deviation 4% or<br />Impact on project costs (direct man hours): deviation between 10%-19% or<br />Reputation: Local media interest or<br />" +
                            "People: MTC OR Single RWDC or<br />Asset / Property / Revenue: Loss between $10K - $100K or<br />External Effect: Moderate Spill, notifiable, without environmental consequence</td></tr>" +
                    
                    "<tr><td style='text-align:center'>1</td><td style='text-align:center'>Insignificant</td>" +
                        "<td>Impact on project schedule: Deviation 1% or<br />Impact on project costs (direct man hours): deviation < 10%  or<br />Reputation: No reaction<br />" +
                            "People : Minor injury with First Aid or<br />Asset / Property / Revenue: Loss < $10K or<br />External Effect: Minor Spill with no environmental Impact</td></tr>" +
                "</table>";
        $('#impact').append(impact);

		fd.toolbar.buttons[1].icon = 'Cancel';
        fd.toolbar.buttons[1].text = "Cancel";

		fd.toolbar.buttons[0].style = "display: none;";

        //$(fd.field('Approval_x0020_Status').$parent.$el).hide();
        fd.field('Approval_x0020_Status').value = "Pending";	
        fd.field('Approval_x0020_Status').disabled = true;
                
        await RR_IsUserInGroupNew('PM');
        //#endregion
        }
        else if(formType == 'Edit'){ 			
			await RR_editForm();
        }
	}
	catch(err)
	{
		alert(err.message);
	}
}

async function onDCRender(formType){
    if(formType == 'New'){
        await DCR_newForm(formType);
    }
    else if(formType == 'Edit'){
        await DCR_editForm(formType);
    }
}

async function onTQRender(formType){
	if(formType == 'New'){
        await TQ_newForm(formType);
    }
    else if(formType == 'Edit'){
        await TQ_editForm(formType);
    }
}

fd.spBeforeSave(function()
{
    if(_modulename == "DCR" && _formType == "New")
        new_beforeSaveDCR();
    else if(_modulename == "DCR" && _formType == "Edit")
        edit_beforeSaveDCR();
    else if(_modulename == "CL" && _formType == "New")
        AttachFiles();
    else if(_modulename == "CL" && _formType == "Edit")
        AttachFiles(); 
	else if(_modulename == "TQ" && _formType == "New")
        AttachFiles();  
	else if(_modulename == "TQ" && _formType == "Edit")
        AttachFiles();  
});

//#region CHANGE LOG FUNCTIONS
async function disableAllFields(GroupName)
{     
	fd.toolbar.buttons[0].style = "display: none;";
	fd.toolbar.buttons[1].text = "Cancel";
	fd.toolbar.buttons[1].icon = 'Cancel';
		
	fd.field('Title').disabled = true;
	fd.field('Description').disabled = true;
	fd.field('RequestedBy').disabled = true;
	fd.field('DateofRequest').disabled = true;
	fd.field('LetterofReference').disabled = true;

	fd.field('TimeImpact').disabled = true;
	fd.field('DeliverablesImpact').disabled = true;
	//fd.field('Attachments').disabled = true;
	SetAttachmentToReadOnly()

	fd.field('AR').disabled = true;
	fd.field('AO').disabled = true;
	fd.field('SB').disabled = true;
	fd.field('EC').disabled = true;

	fd.field('EL').disabled = true;
	fd.field('GE').disabled = true;
	fd.field('IT').disabled = true;
	fd.field('LAD').disabled = true;

	fd.field('ME').disabled = true;
	fd.field('PM').disabled = true;
	fd.field('PMC').disabled = true;
	fd.field('PUD').disabled = true;

	fd.field('QM').disabled = true;
	fd.field('WE').disabled = true;
	fd.field('TR').disabled = true;

	await pnp.sp.web.lists.getByTitle("Change Log").get().then(()=>{
		return pnp.sp.web.lists.getByTitle("Change Log").fields.getByInternalNameOrTitle('DSS').get().then(field => {
			fd.field('DSS').disabled = true;
		});
	})
	.catch(function(err) {		
	});	

	fd.field('Communications').disabled = true;
	fd.field('OutsideServices').disabled = true;
	fd.field('PrintingandDispatch').disabled = true;
	fd.field('TravelandAccommodation').disabled = true;
	fd.field('TotalODC').disabled = true;
	fd.field('Comments').disabled = true;
	
	fd.field('AreaApproval').disabled = true; 		
	fd.field('AreaOfficeRemarks').disabled = true; 	
	fd.field('AreaOfficeRemarks').disabled = true;
	$(fd.field('PDApprovalDate').$parent.$el).hide();
	$(fd.field('AreaFeedbackDate').$parent.$el).hide();
	
	fd.field('ClaimStatus').disabled = true;
	fd.field('ApprovedMM').disabled = true;
	
	fd.field('ProjectDirectorApproval').disabled = true;	
	fd.field('DirectorRemarks').disabled = true; 
	$(fd.field('PDApprovalDate').$parent.$el).hide();
	$(fd.field('AreaFeedbackDate').$parent.$el).hide();

	if(GroupName == "Area")
	{
		var ADSectionHideElement = document.querySelector('.ADSectionHide');
		ADSectionHideElement.style.display = 'none';
	}
}

async function disablePMFields()
{ 
    fd.field('Title').disabled = true;
	var elem = $("textarea")[0];
	$(elem).prop("readonly", true);
	// elem.style.height = "auto";
	// elem.style.height = (elem.scrollHeight) + "px";
	//fd.field('Description').disabled = true;
	fd.field('RequestedBy').disabled = true;
	fd.field('DateofRequest').disabled = true;
	fd.field('LetterofReference').disabled = true;

	// fd.field('TimeImpact').disabled = true;
	// fd.field('DeliverablesImpact').disabled = true;

	var elem1 = $("textarea")[1];
	$(elem1).prop("readonly", true);

	var elem2 = $("textarea")[2];
	$(elem2).prop("readonly", true);
	//fd.field('Attachments').disabled = true;	

	fd.field('AR').disabled = true;
	fd.field('AO').disabled = true;
	fd.field('SB').disabled = true;
	fd.field('EC').disabled = true;

	fd.field('EL').disabled = true;
	fd.field('GE').disabled = true;
	fd.field('IT').disabled = true;
	fd.field('LAD').disabled = true;

	fd.field('ME').disabled = true;
	fd.field('PM').disabled = true;
	fd.field('PMC').disabled = true;
	fd.field('PUD').disabled = true;

	fd.field('QM').disabled = true;
	fd.field('WE').disabled = true;
	fd.field('TR').disabled = true;
	
	await pnp.sp.web.lists.getByTitle("Change Log").get().then(()=>{
		return pnp.sp.web.lists.getByTitle("Change Log").fields.getByInternalNameOrTitle('DSS').get().then(field => {
			fd.field('DSS').disabled = true;
		});
	})
	.catch(function(err) {		
	});	

	fd.field('Communications').disabled = true;
	fd.field('OutsideServices').disabled = true;
	fd.field('PrintingandDispatch').disabled = true;
	fd.field('TravelandAccommodation').disabled = true;
	fd.field('TotalODC').disabled = true;
	fd.field('Comments').disabled = true;
}

async function SetDisabledForPM(ChangeLogPMFullPermission) 
{			
	if(ChangeLogPMFullPermission.toLowerCase() == "no") 
	{		
		$(fd.field('AreaApproval').$parent.$el).hide();	
		$(fd.field('AreaOfficeRemarks').$parent.$el).hide();			
		$(fd.field('PDApprovalDate').$parent.$el).hide();
		$(fd.field('AreaFeedbackDate').$parent.$el).hide();		
		$(fd.field('ClaimStatus').$parent.$el).hide();			
		$(fd.field('ApprovedMM').$parent.$el).hide();			
	
		$(fd.field('ProjectDirectorApproval').$parent.$el).hide();		
		$(fd.field('DirectorRemarks').$parent.$el).hide();
		$(fd.field('PDApprovalDate').$parent.$el).hide();
		$(fd.field('AreaFeedbackDate').$parent.$el).hide();	
		
		var spanPDElement = document.querySelector('.PDSection');
		spanPDElement.style.display = 'none';	
		
		fd.toolbar.buttons[0].style = "display: none;";
		fd.toolbar.buttons[1].text = "Cancel"; 
		fd.toolbar.buttons[1].icon = 'Cancel';
		
		fd.toolbar.buttons.push({
	        icon: 'Save',
	        class: 'btn-outline-primary',
	        text: 'Save',
	        click: function() {							
			
			WorkflowStatus = "Pending";
			fd.field('WorkflowStatus').value = WorkflowStatus;
			
			$(fd.field('Reviewed').$parent.$el).show();
			fd.field('Reviewed').value = false;
			$(fd.field('Reviewed').$parent.$el).hide();			
				
            fd.save();						 
	     }
    	});
		
		fd.toolbar.buttons.push({
	        icon: 'Accept',
	        class: 'btn-outline-primary',
	        text: 'Submit',
		        click: function() {							
				
				WorkflowStatus = "Submitted to PD";					
				fd.field('WorkflowStatus').value = WorkflowStatus;
				
				$(fd.field('Reviewed').$parent.$el).show();
				fd.field('Reviewed').value = true;	
				$(fd.field('Reviewed').$parent.$el).hide();					
					
	            fd.save();						 
		     }
    	}); 
	}		
}

async function SetDisabledForReturnToPM(ChangeLogPMFullPermission) 
{		
	disablePMFields();	
			
	if(ChangeLogPMFullPermission.toLowerCase() == "no") 
	{
		fd.field('AreaApproval').disabled = true; 		
		fd.field('AreaOfficeRemarks').disabled = true;		
		$(fd.field('PDApprovalDate').$parent.$el).hide();
		$(fd.field('AreaFeedbackDate').$parent.$el).hide();
		
		fd.field('ClaimStatus').disabled = true;
		fd.field('ApprovedMM').disabled = true;
		
		fd.field('ProjectDirectorApproval').disabled = true;	
		fd.field('DirectorRemarks').disabled = true; 
		$(fd.field('PDApprovalDate').$parent.$el).hide();
		$(fd.field('AreaFeedbackDate').$parent.$el).hide();
		fd.field('Attachments').disabled = false;
		
		fd.toolbar.buttons[0].style = "display: none;";
		fd.toolbar.buttons[1].text = "Cancel"; 
		fd.toolbar.buttons[1].icon = 'Cancel';
		
		fd.toolbar.buttons.push({
	        icon: 'Save',
	        class: 'btn-outline-primary',
	        text: 'Save',
	        click: function() {							
			
			WorkflowStatus = "Returned to PM";
			fd.field('WorkflowStatus').value = WorkflowStatus;
			fd.field('Reviewed').value = false;			
				
            fd.save();						 
	     }
    	});
		
		fd.toolbar.buttons.push({
	        icon: 'Accept',
	        class: 'btn-outline-primary',
	        text: 'Submit to Area',
		        click: function() {							
				
				WorkflowStatus = "Returned to Area";					
				fd.field('WorkflowStatus').value = WorkflowStatus;
				fd.field('Reviewed').value = true;				
				
				fd.validators.push({
				name: 'Check Attachment',
				error: "Please upload the Claim File",
					validate: function(value) {				
							
							if(fd.field('Attachments').value.length < 1)
							{
							    this.error = "Please ensure that the claim file is attached before submitting.";
								return false; 
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
										NCountofATT++;										
								}
								
								$(fd.field('AttachFiles').$parent.$el).show();
								fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
								$(fd.field('AttachFiles').$parent.$el).hide();
								
								$(fd.field('ClaimAttachment').$parent.$el).show();
								fd.field('ClaimAttachment').value = "Submitted";
								$(fd.field('ClaimAttachment').$parent.$el).hide();
								
								if(NCountofATT === 0)
									return false;									
							}					
						 	return true;
						}
					});		
					
	            fd.save();						 
		     }
    	});	
	}		
}

async function SetDisabledForPD(currentUser) 
{				
	fd.field('AreaApproval').disabled = true; 		
	fd.field('AreaOfficeRemarks').disabled = true; 
	$(fd.field('PDApprovalDate').$parent.$el).hide();
	$(fd.field('AreaFeedbackDate').$parent.$el).hide();
	
	fd.field('ClaimStatus').disabled = true;
	fd.field('ApprovedMM').disabled = true;
	disablePMFields();	
	SetAttachmentToReadOnly();	
	
	fd.toolbar.buttons[0].style = "display: none;";
	fd.toolbar.buttons[1].text = "Cancel";
	fd.toolbar.buttons[1].icon = 'Cancel';
	
	fd.toolbar.buttons.push({
	        icon: 'Save',
	        class: 'btn-outline-primary',
	        text: 'Save',
	        click: function() {							
			
			WorkflowStatus = "Submitted to PD";
			fd.field('WorkflowStatus').value = WorkflowStatus;
			fd.field('Reviewed').value = false;			
				
            fd.save();						 
	     }
    });
	
	fd.toolbar.buttons.push({
	        icon: 'Accept',
	        class: 'btn-outline-primary',
	        text: 'Submit',
	        click: function() {							
			
			if(fd.field('ProjectDirectorApproval').value === "Approved")
				WorkflowStatus = "Submitted to Area";
			else if(fd.field('ProjectDirectorApproval').value === "Returned to PM")
				WorkflowStatus = "Pending";				
			else
				WorkflowStatus = "Submitted";
				
			fd.field('WorkflowStatus').value = WorkflowStatus;
			fd.field('Reviewed').value = true;	
			
			//var currentUser = await pnp.sp.web.currentUser.get();		

			$(fd.field('ReviewedByPD').$parent.$el).show();
			fd.field('ReviewedByPD').value = currentUser;
			$(fd.field('ReviewedByPD').$parent.$el).hide();
			
			var date = new Date();
		    fd.field('PDApprovalDate').value = date;
			
			fd.field('ProjectDirectorApproval').addValidator({
	        name: 'PD Approval Status',
	        error: 'Project Director Approval is required field',
	        validate:function(value){
	            if(value == null || value == ""){
	                return false;
	            }
	            return true;
	           }
	        });		
				
            fd.save();						 
	     }
    });				
}

async function SetDisabledForAD(currentUser) 
{	    
	fd.field('ProjectDirectorApproval').disabled = true;	
	fd.field('DirectorRemarks').disabled = true; 
	$(fd.field('PDApprovalDate').$parent.$el).hide();
	$(fd.field('AreaFeedbackDate').$parent.$el).hide();
	$(fd.field('ClaimStatus').$parent.$el).hide();
	$(fd.field('ApprovedMM').$parent.$el).hide();
	
	var ADSectionHideElement = document.querySelector('.ADSectionHide');
	ADSectionHideElement.style.display = 'none';	

	disablePMFields();
	SetAttachmentToReadOnly();
	
	fd.toolbar.buttons[0].style = "display: none;";
	fd.toolbar.buttons[1].text = "Cancel"; 
	fd.toolbar.buttons[1].icon = 'Cancel';      
	
	fd.toolbar.buttons.push({
	        icon: 'Save',
	        class: 'btn-outline-primary',
	        text: 'Save',
	        click: function() {							
			
			WorkflowStatus = "Submitted to Area";
			fd.field('WorkflowStatus').value = WorkflowStatus;
			fd.field('Reviewed').value = false;			
				
            fd.save();						 
	     }
    });
	
	fd.toolbar.buttons.push({
	        icon: 'Accept',
	        class: 'btn-outline-primary',
	        text: 'Submit',
	        click: function() {		
			
			if(fd.field('AreaApproval').value === "Proceed and submit claim" || fd.field('AreaApproval').value === "Do not proceed, submit claim, and wait for client decision")
				WorkflowStatus = "Returned to PM";				
			else
				WorkflowStatus = "Submitted";
								
			fd.field('WorkflowStatus').value = WorkflowStatus;
			fd.field('Reviewed').value = true;
			
			//var currentUser = await pnp.sp.web.currentUser.get();		

			$(fd.field('ReviewedByAD').$parent.$el).show();
			fd.field('ReviewedByAD').value = currentUser;
			$(fd.field('ReviewedByAD').$parent.$el).hide();
			
			var date = new Date();
		    fd.field('AreaFeedbackDate').value = date;
			
			fd.field('AreaApproval').addValidator({
	        name: 'Area Feedback',
	        error: 'Area Feedback is required field',
	        validate:function(value){
	            if(value == null || value == ""){
	                return false;
	            }
	            return true;
	           }
	        });	
								
            fd.save();						 
	     }
    });			
}

async function SetDisabledForReturnedAD() 
{	    
	fd.field('ProjectDirectorApproval').disabled = true;	
	fd.field('DirectorRemarks').disabled = true; 
	$(fd.field('PDApprovalDate').$parent.$el).hide();
	$(fd.field('AreaFeedbackDate').$parent.$el).hide();
	fd.field('AreaApproval').disabled = true;
	
	var ADSectionHideElement = document.querySelector('.ADSectionHide');
	ADSectionHideElement.style.display = 'none';
	
	disablePMFields();	
	SetAttachmentToReadOnly();	
	
	fd.toolbar.buttons[0].style = "display: none;";
	fd.toolbar.buttons[1].text = "Cancel"; 
	fd.toolbar.buttons[1].icon = 'Cancel';
	
	fd.field('ClaimStatus').$on('change',ClaimStatushideOrShow);
	ClaimStatushideOrShow();      
	
	fd.toolbar.buttons.push({
	        icon: 'Save',
	        class: 'btn-outline-primary',
	        text: 'Save',
	        click: function() {							
			
			WorkflowStatus = "Returned to Area";
			fd.field('WorkflowStatus').value = WorkflowStatus;
			fd.field('Reviewed').value = false;			
				
            fd.save();						 
	     }
    });
	
	fd.toolbar.buttons.push({
	        icon: 'Accept',
	        class: 'btn-outline-primary',
	        text: 'Submit',
	        click: function() {	
			
			WorkflowStatus = "Submitted";									
			fd.field('WorkflowStatus').value = WorkflowStatus;
			fd.field('Reviewed').value = true;
			
			var date = new Date();
		    fd.field('AreaFeedbackDate').value = date;				
			
			fd.field('ClaimStatus').addValidator({
	        name: 'Claim Status',
	        error: 'Claim Status is required field',
	        validate:function(value){
	            if(value == null || value == ""){
	                return false;
	            }
	            return true;
	           }
	        });		
				
            fd.save();						 
	     }
    });			
}

async function showHideFields(ChangeLogPMFullPermission, currentUser) 
{
	try
	{        	
		// get all groups the current user belongs to

		// var ItemID = fd.itemId;
		// var wfStatus = await GetColumnValueByID('Change Log', ItemID, 'WorkflowStatus');		
		var wfStatus = fd.field('WorkflowStatus').value;		
		
        var userId = _spPageContextInfo.userId;
        var userGroups = [];
        await pnp.sp.web.siteUsers.getById(userId).groups.get()
        .then(async function(groupsData){
            for (var i = 0; i < groupsData.length; i++) {
                userGroups.push(groupsData[i].Title);
            }

            var isAdmin = false;
            var GroupName = "";
            if (userGroups.indexOf('Owners') >= 0) isAdmin = true;

            if (!isAdmin) { 			

                if (userGroups.indexOf('PM') >= 0 && wfStatus === "Pending") {
                    GroupName = "PM";
                    await SetDisabledForPM(ChangeLogPMFullPermission);
                }
                else if (userGroups.indexOf('PM') >= 0 && wfStatus === "Returned to PM") {
                    GroupName = "PM";
                    await SetDisabledForReturnToPM(ChangeLogPMFullPermission);
                }
                else if (userGroups.indexOf('PD') >= 0 && wfStatus === "Submitted to PD") {
                    GroupName = "PD";							
                    await SetDisabledForPD(currentUser);
                }
                else if (userGroups.indexOf('Area') >= 0 && wfStatus === "Submitted to Area") {
                    GroupName = "Area";							
                    await SetDisabledForAD(currentUser);
                }
                else if (userGroups.indexOf('AD') >= 0 && wfStatus === "Submitted to Area") {
                    GroupName = "Area";	
                    await SetDisabledForAD(currentUser);
                }
                else if (userGroups.indexOf('Area') >= 0 && wfStatus === "Returned to Area") {
                    GroupName = "Area";	
                    await SetDisabledForReturnedAD();
                }
                else if (userGroups.indexOf('AD') >= 0 && wfStatus === "Returned to Area") {
                    GroupName = "Area";	
                    await SetDisabledForReturnedAD();
                }
                else
                {                       
                    if (userGroups.indexOf('Area') >= 0 || userGroups.indexOf('AD') >= 0)
                        GroupName = "Area";

					await disableAllFields(GroupName);                      	
                }					
            }

        });		
	}
	catch(err)
	{
		alert(err.message);
	}
}

function showHideFieldsDisplay() 
{
	try
	{        	
		// get all groups the current user belongs to		
		var wfStatus = fd.field('WorkflowStatus').value;

        var userId = _spPageContextInfo.userId;
        var userGroups = [];
        pnp.sp.web.siteUsers.getById(userId).groups.get()
        .then(function(groupsData){
            for (var i = 0; i < groupsData.length; i++) {
                userGroups.push(groupsData[i].Title);
            }

            var isAdmin = false;
            var GroupName = "";
            if (userGroups.indexOf('Owners') >= 0) isAdmin = true;

            if (!isAdmin) { 

                if (userGroups.indexOf('Area') >= 0 || userGroups.indexOf('AD') >= 0)
                    GroupName = "Area";

                if(GroupName == "Area")
                {
                    var ADSectionHideElement = document.querySelector('.ADSectionHide');
                    ADSectionHideElement.style.display = 'none';
                } 
                
                if(wfStatus == "Submitted")                
                    fd.toolbar.buttons[0].style = "display: none;";             
                
                fd.toolbar.buttons[1].text = "Cancel";
            }
        });		
	}
	catch(err)
	{
		alert(err.message);
	}
}

function ClaimStatushideOrShow() {	

    if (fd.field('ClaimStatus').value === "Approved") {
		$(fd.field('ApprovedMM').$parent.$el).show();		
		fd.field('ApprovedMM').required = true;			
	}
	
	else {		
		fd.field('ApprovedMM').required = false;
		$(fd.field('ApprovedMM').$parent.$el).hide();							
	}	
} 

const getParameterRACI = async function(key){
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
//#endregion
 
//#region RISK REGISTER FUNCTIONS

async function RR_IsUserInGroupNew(group) 
{
    var IsPMUser = false
		try{
			await pnp.sp.web.currentUser.get()
	         .then(async function(user){
				await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
				 .then(function(groupsData){
					for (var i = 0; i < groupsData.length; i++) {
					
						if(groupsData[i].Title == group)
						{						
							IsPMUser = true;
					    }
					}			
					
                    var textVal;
					if(IsPMUser)
					  textVal = "Submit";
					else textVal = "Submit for Approval";

					fd.toolbar.buttons.push({
						icon: 'Save',
						class: 'btn-outline-primary',
						text: 'Save',
						click: function() { 					
						fd.save();			 
					    }
				    });
					
					if(fd.field('Approval_x0020_Status').value == 'Pending')
						fd.toolbar.buttons.push({
							icon: 'Accept',
							class: 'btn-outline-primary',
							text: textVal,
							click: function() {
								fd.field('Approval_x0020_Status').disabled = false;
							    if(IsPMUser)
							      fd.field('Approval_x0020_Status').value = "Approved";
								else fd.field('Approval_x0020_Status').value = 'Sent for Approval';
								fd.field('Approval_x0020_Status').disabled = true;
								setTimeout(function(){ fd.save(); }, 300);		
							}
						});											
				});
		     });
	    }
		catch(e){alert(e);}
		return IsPMUser;
}

async function RR_IsUserInGroup(group) 
{
    var IsPMUser = false
		try{
			await pnp.sp.web.currentUser.get()
	         .then(async function(user){
				await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
				 .then(function(groupsData){
					for (var i = 0; i < groupsData.length; i++) 
					{					
						if(groupsData[i].Title == group)
						{
							IsPMUser = true;
							
							if(fd.field('Approval_x0020_Status').value == 'Sent for Approval')
							{
								fd.toolbar.buttons[0].style = "display: none;";
								fd.toolbar.buttons.push({
									icon: 'Accept',
									class: 'btn-outline-primary',
									text: 'Approve',
									click: function() {
										fd.field('Approval_x0020_Status').value = 'Approved';
										setTimeout(function(){ fd.save(); }, 300);		
									}
								});
								fd.toolbar.buttons.push({
									icon: 'Reply',
									class: 'btn-outline-primary',
									text: 'Reject',
									click: function() {
										fd.field('Approval_x0020_Status').value = 'Rejected';
										setTimeout(function(){ fd.save(); }, 300);		
									}
								});
							}
							else if(fd.field('Approval_x0020_Status').value == 'Pending')
							{								
								fd.toolbar.buttons.push({
									icon: 'Accept',
									class: 'btn-outline-primary',
									text: "Submit",
									click: function() {
										fd.field('Approval_x0020_Status').disabled = false;
									    fd.field('Approval_x0020_Status').value = "Approved";										
										fd.field('Approval_x0020_Status').disabled = true;
										setTimeout(function(){ fd.save(); }, 300);		
									}
								});								
							}							   
					    }						
					}				
					
				  	var textVal;
					if(IsPMUser)
					  textVal = "Submit";
					else textVal = "Submit for Approval";					   
													
					if(fd.field('Status').value == "Open" && !IsPMUser)
					{
						if(fd.field('Approval_x0020_Status').value == 'Sent for Approval')
						{}
						else
						{
							fd.toolbar.buttons.push({
						    icon: 'Accept',
							class: 'btn-outline-primary',
							text: textVal,
							click: function() {
								fd.field('Approval_x0020_Status').disabled = false;
							    if(IsPMUser)
							      fd.field('Approval_x0020_Status').value = "Approved";
								else fd.field('Approval_x0020_Status').value = 'Sent for Approval';
								fd.field('Approval_x0020_Status').disabled = true;
								setTimeout(function(){ fd.save(); }, 300);		
									}
								});
						}
					}
					else if(fd.field('Status').value == "Open" && IsPMUser && fd.field('Approval_x0020_Status').value == 'Approved')
					{
						//fd.toolbar.buttons[0].style = "display: none;";
						
						fd.toolbar.buttons.push({
							icon: 'Accept',
							class: 'btn-outline-primary',
							text: "Submit & Close",
							click: function() {	
								 fd.field('Status').disabled = false;								
							     fd.field('Status').value = "Closed";	
								 fd.field('Status').disabled = true;								
								setTimeout(function(){ fd.save(); }, 300);		
									}
								});	
					}						
					else if(fd.field('Status').value == "Closed" && IsPMUser)
					{
						fd.toolbar.buttons.push({
							icon: 'Reply',
							class: 'btn-outline-primary',
							text: "Re-Open",
							click: function() {	
								 fd.field('Status').disabled = false;								
							     fd.field('Status').value = "Open";	
								 fd.field('Status').disabled = true;								
								setTimeout(function(){ fd.save(); }, 300);		
									}
						});	
					}				
				});
		     });
	    }
		catch(e){alert(e);}
		return IsPMUser;
}

async function RR_editForm(){
    var prob = "<table cellspacing='0' cellpadding='0' width='100%' border='1' style='border-collapse:collapse; font-size:11px;'>" +
		            		"<tr><td width='10%' style='text-align:center; background-color:lightgrey; font-weight:bold'>Rating</td><td width='20%' style='background-color:lightgrey; font-weight:bold;text-align:center'>Description</td><td style='background-color:lightgrey; font-weight:bold'>Definition</td></tr>" +
		            		"<tr><td style='text-align:center'>5</td><td style='text-align:center'>Almost certain</td><td>>81% probability or<br />Could occur within days or<br />Event could occur several times during the life of the facility or activity</td></tr>" +
		            		"<tr><td style='text-align:center'>4</td><td style='text-align:center'>Likely</td><td>>61% to 80% probability or<br />Could occur within a couple of weeks or<br />Event has occurred several times per year in similar industries OR more than once per year in Dar’s experience</td></tr>" +
		            		"<tr><td style='text-align:center'>3</td><td style='text-align:center'>Possible</td><td>>21 to 60% probability or<br />Could occur within a month or<br />Event has occurred in Dar’s experience</td></tr>" +
		            		"<tr><td style='text-align:center'>2</td><td style='text-align:center'>Unlikely</td><td>>11 to 20% probability or<br />Could occur in a couple of months or<br />Event occurrence is possible but unlikely or has happened a few times in similar industries.</td></tr>" +
		            		"<tr><td style='text-align:center'>1</td><td style='text-align:center'>Rare</td><td>>0 to 10% probability or<br />Could occur in more than a few months or<br />Event is physically possible but has never been heard of in similar industries.</td></tr>" +
						  "</table>";
			   $('#prob').append(prob);
	   
	   var impact = "<table cellspacing='0' cellpadding='0' width='100%' border='1' style='border-collapse:collapse; font-size:11px;'>" +
			            "<tr><td width='10%' style='text-align:center; background-color:lightgrey; font-weight:bold'>Rating</td><td width='20%' style='background-color:lightgrey; font-weight:bold;text-align:center'>Description</td><td style='background-color:lightgrey; font-weight:bold'>Definition</td></tr>" +
			            
			            "<tr><td style='text-align:center'>5</td><td style='text-align:center'>Critical</td>" +
			                "<td>Impact on project schedule: Deviation 16% or<br />Impact on project costs (direct man hours): deviation over 60% or<br />Reputation: Risk covered on international TV and newspapers or<br />" +
			                    "People: Multiple fatalities, Asset / Property or<br />Asset / Property / Revenue: Loss > $10M,<br />External Effect: Major and sustained  external pollution</td></tr>" +

			              "<tr><td style='text-align:center'>4</td><td style='text-align:center'>Major</td>" +
			                "<td>Impact on project schedule: Deviation 12% or<br />Impact on project costs (direct man hours): deviation between 40%-59% or<br />Reputation: Risk covered on regional TV and press or<br />" +
			                    "People: Multiple LTIs OR one or more Permanent Disability OR 1 Fatality or<br />Asset / Property / Revenue: Loss between $1M – $10M or<br />External Effect :Significant pollution with reversible environmental consequences</td></tr>" +
			            
			               "<tr><td style='text-align:center'>3</td><td style='text-align:center'>Moderate</td>" +
			                "<td>Impact on project schedule: Deviation 8% or<br />Impact on project costs (direct man hours): deviation between 20%-39% or<br />Reputation: Risk covered on local TV and press or<br />" +
			                    "People: Single LTI OR multiple RWDC or<br />Asset / Property / Revenue: Loss between $100K – $1M or<br />External Effect: Significant pollution having external impact</td></tr>" +
			            
			             "<tr><td style='text-align:center'>2</td><td style='text-align:center'>Minor</td>" +
			                "<td>Impact on project schedule: Deviation 4% or<br />Impact on project costs (direct man hours): deviation between 10%-19% or<br />Reputation: Local media interest or<br />" +
			                    "People: MTC OR Single RWDC or<br />Asset / Property / Revenue: Loss between $10K - $100K or<br />External Effect: Moderate Spill, notifiable, without environmental consequence</td></tr>" +
			            
			             "<tr><td style='text-align:center'>1</td><td style='text-align:center'>Insignificant</td>" +
			                "<td>Impact on project schedule: Deviation 1% or<br />Impact on project costs (direct man hours): deviation < 10%  or<br />Reputation: No reaction<br />" +
			                    "People : Minor injury with First Aid or<br />Asset / Property / Revenue: Loss < $10K or<br />External Effect: Minor Spill with no environmental Impact</td></tr>" +
					"</table>";
					
	   $('#impact').append(impact);	

	    fd.toolbar.buttons[1].icon = 'Cancel';
        fd.toolbar.buttons[1].text = "Cancel";
		
		// var ItemID = fd.itemId;
		// var status = await GetColumnValueByID('Risk Register', ItemID, 'Approval_x0020_Status');

		var status = fd.field('Approval_x0020_Status').value 
		fd.field('Approval_x0020_Status').disabled = true;
		fd.field('Status').disabled = true;			 

        if(status == "Sent for Approval")
		{
		   fd.toolbar.buttons[0].style = "display: none;";		   
		}
		else if(status == "Approved" || status == "Rejected")
		{
			fd.toolbar.buttons[0].icon = 'Save';

		   if(fd.field('Status').value == "Open") {				   
		   }
		   else { 
				   fd.toolbar.buttons[0].style = "display: none;";				
				   fd.field('Status').disabled = true;
				   
				   fd.field('Risk_x0020_Category').disabled = true;
				   fd.field('Approval_x0020_Status').disabled = true;
				   fd.field('RiskDescription').disabled = true;
				   fd.field('CauseofRisk').disabled = true;
				   fd.field('ProbabilityofRisk').disabled = true;
				   
				   fd.field('ImpactofRisk').disabled = true;
				   fd.field('Risk_x0020_Impact').disabled = true;
				   fd.field('RiskRating').disabled = true;
				   fd.field('RiskOwner').disabled = true; 
				   
				   fd.field('Show_x0020_in_x0020_Client_x0020').disabled = true;
				   fd.field('Risk_x0020_Description_x0020_for').disabled = true;
				   
				   fd.field('ResProbaOfRisk').disabled = true;
				   fd.field('ResImpecOfRisk').disabled = true;
				   fd.field('ResRiskRating').disabled = true;
				   fd.field('Residual_x0020_Impact').disabled = true;
				   fd.field('RiskStrategy_x002f_Response').disabled = true;
				   fd.field('Mitigation_x002f_EnhancementStra').disabled = true;

				   $(fd.field('Approval_x0020_Status').$parent.$el).show();
				   fd.field('Approval_x0020_Status').disabled = true;
			    }		   
		   
		   //return;
		}

        await RR_IsUserInGroup('PM');
}
//#endregion

//#region DESIGN CODES AND REGULATIONS FUNCTIONS

var GetAllLocalCodesbyCountry = async function(Country){		
	return new Promise(function(resolve, reject) {
        var xhr = new XMLHttpRequest();
        var soapRequest = "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"
            + "<soap:Body>"
            + "<CallCountryCodesService xmlns=\"http://dargroup.com/\">"
            + "<country>" + Country + "</country>"
            + "</CallCountryCodesService>"
            + "</soap:Body>"
            + "</soap:Envelope>";

        xhr.open("POST", SERVICE_URL, true);

        xhr.onreadystatechange = async function() {
            if (xhr.readyState == 4) {
                var text;
                try {
                    if (xhr.status == 200) {
                        var xmlDoc = $.parseXML(xhr.responseText);
                        $xml = $(xmlDoc);
                        $value = $xml.find("CallCountryCodesServiceResult");

                        text = $value.text();

                        const obj = await JSON.parse(text, async function(key, value) {
                            var _columnName = key;
                            var _value = value;

                            if (_columnName === 'Title') {
                                if (!CodeArr.includes(_value))
                                    CodeArr.push(_value);
                            }
                        });

                        var response = JSON.parse(text);
                        CodeObjectArray.push(response.d.results);

                        fd.field('DropDownCC').widget.setDataSource({
                            data: CodeArr
                        })
                        resolve(CodeArr);
                    } else {
                        console.log('Request failed with status code:', xhr.status);
                        reject('Request failed with status code:', xhr.status);
                    }
                } catch (err) {
                    console.log(err + "\n" + text);
                    reject(err + "\n" + text);
                }
            }
        }

        xhr.setRequestHeader('Content-Type', 'text/xml');
        xhr.send(soapRequest);
    });
}

var GetAllInternationalCodes = async function(){	
	return new Promise(function(resolve, reject) {
        var xhr = new XMLHttpRequest();
        var soapRequest = "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"
            + "<soap:Body>"
            + "<CallInternationalCountryCodesService xmlns=\"http://dargroup.com/\" />"
            + "</soap:Body>"
            + "</soap:Envelope>";

        xhr.open("POST", International_SERVICE_URL, true);

        xhr.onreadystatechange = async function() {
            if (xhr.readyState == 4) {
                var text;
                try {
                    if (xhr.status == 200) {
                        var xmlDoc = $.parseXML(xhr.responseText);
                        $xml = $(xmlDoc);
                        $value = $xml.find("CallInternationalCountryCodesServiceResult");

                        text = $value.text();

                        const obj = await JSON.parse(text, async function(key, value) {
                            var _columnName = key;
                            var _value = value;

                            if (_columnName === 'Title') {
                                if (!CodeArr.includes(_value))
                                    CodeArr.push(_value);
                            }
                        });

                        var response = JSON.parse(text);
                        CodeObjectArray.push(response.d.results);

                        fd.field('DropDownCC').widget.setDataSource({
                            data: CodeArr
                        });
                        resolve(CodeArr);
                    } else {
                        console.log('Request failed with status code:', xhr.status);
                        reject('Request failed with status code:', xhr.status);
                    }
                } catch (err) {
                    console.log(err + "\n" + text);
                    reject(err + "\n" + text);
                }
            }
        }

        xhr.setRequestHeader('Content-Type', 'text/xml');
        xhr.send(soapRequest);
    });
}

async function populateCountryCodes()
{ 
    $(fd.field('IsCodeAvailable').$parent.$el).hide(); 	
   
	var country = fd.field('Country').value; //"Lebanon";//
	if(country === "") return;
	if(country === "International") 
	{
		fd.field('DropDownCC').widget.dataSource.data(['Please wait while retreiving ...']);  
		$("ul.k-reset")
			.find("li")
			.first()
			.css("pointer-events", "none")
			.css("opacity", "0.6"); 
        CodeArr=[];
		delay(100);
	    await GetAllInternationalCodes();	
	}	
	else 
	{	
		fd.field('DropDownCC').widget.dataSource.data(['Please wait while retreiving ...']);  
		$("ul.k-reset")
			.find("li")
			.first()
			.css("pointer-events", "none")
			.css("opacity", "0.6"); 	
		CodeArr=[];
		delay(100);
	    await GetAllLocalCodesbyCountry(country);
	}
}

function GetCountryCodeDetails(code) {
	
	for (const items of CodeObjectArray) {
		for (const item of items) {
			if (item.Title === code.toString()) {

				$(fd.field('Description').$parent.$el).show();	
				fd.field('Description').disabled = false;	  	
 		  		fd.field('Description').value = item.Description
				fd.field('Description').disabled = true;

				$(fd.field('Hyperlink').$parent.$el).show();				
				var folderURL = item.Hyperlink;
				var SiteUrl = "https://4ce.dar.com";//window.location.protocol + "//" + window.location.host
				if(folderURL !== null)
				{
					fd.field('Hyperlink').value.url = SiteUrl + folderURL;//relativePath;
					fd.field('Hyperlink').value.description = 'go to folder';
					fd.field('Hyperlink').disabled = true;
				}
				else
				{
					fd.field('Hyperlink').value.url = "";//relativePath;
					fd.field('Hyperlink').value.description = '';
					$(fd.field('Hyperlink').$parent.$el).hide();
				}				
			}
		}
	}
}

function GetInternationalCodeDetails(code) {

	for (const items of CodeObjectArray) {
		for (const item of items) {
			if (item.Title === code.toString()) {

				$(fd.field('Description').$parent.$el).show();	
				fd.field('Description').disabled = false;	  	
 		  		fd.field('Description').value = item.Description
				if(code.toString() !== "Others")
				   fd.field('Description').disabled = true;
			    else
				   fd.field('Description').disabled = false;

				$(fd.field('Hyperlink').$parent.$el).show();				
				var folderURL = item.Hyperlink;
				var SiteUrl = "https://4ce.dar.com";//window.location.protocol + "//" + window.location.host
				if(folderURL !== null)
				{
					fd.field('Hyperlink').value.url = SiteUrl + folderURL;//relativePath;
					fd.field('Hyperlink').value.description = 'go to folder';
					fd.field('Hyperlink').disabled = true;
				}
				else
				{
					fd.field('Hyperlink').value.url = "";//relativePath;
					fd.field('Hyperlink').value.description = '';
					$(fd.field('Hyperlink').$parent.$el).hide();
				}				
			}
		}
	}	
}
//#endregion

//#region TENDER QUERY
var TQ_newForm = async function(){	

	fd.validators.push({
		name: 'MyCustomValidatorTenderCircularNo',
		error: "Group No. and ID No. for same Tenderer is already exists.",
		validate: function(value) {				
			if (_Proceed > 0)
				return false;
			return true;			
		}
	});

	$(fd.field('AttachFiles').$parent.$el).hide();
	fd.field('Answered').value = false;
	fd.field('SendToTM').value = false;
	$(fd.field('Answered').$parent.$el).hide();
	$(fd.field('SendToTM').$parent.$el).hide();
	fd.field('WorkflowStatus').disabled = true;
	$(fd.field('QueryType').$parent.$el).hide();
	$(fd.field('QueryLookupNumber').$parent.$el).hide();
	$(fd.field('ClientInternalNote').$parent.$el).hide();
	$(fd.field('DarInternalNotes').$parent.$el).hide();
	$(fd.field('Attachments').$parent.$el).hide();
	
	$(fd.field('EntryUniqueNumber').$parent.$el).hide();
	fd.field('BatchNoText').value = "";
	fd.field('SerialNoText').value = "";
	fd.field('EntryBatch').value = "";
	fd.field('SerialNumber').value = "";	
	
	const GroupName = await CheckifUserinTMorATM();	
	
	try{
		$(fd.field('PackageAcronym').$parent.$el).hide();
		$(fd.field('Package').$parent.$el).hide();

		fd.field('PackageDropDown').widget.dataSource.data(['Please wait while retreiving ...']);  
		$("ul.k-reset")
			.find("li")
			.first()
			.css("pointer-events", "none")
			.css("opacity", "0.6"); 	
		PackageArr=[];
		delay(100);
		var ProjectNo = _spPageContextInfo.siteAbsoluteUrl.split('/')[4]; 
	    await GetAllPckagesByProjectNumber(ProjectNo);	
		
		fd.field('PackageDropDown').$on('change', function(value)
		{	
			$(fd.field('Package').$parent.$el).show();	
			fd.field('Package').value = value;
			$(fd.field('Package').$parent.$el).hide();
		});
	} 
	catch{
		fd.field("Package").ready().then(function() {
			fd.field("Package").value = null;
			fd.field("Package").filter = "ListName eq 'Package'";
			 fd.field("Package").refresh();	      
		});
	}	
		
	fd.field("Contractor").ready().then(function() {
		fd.field("Contractor").value = null;
		fd.field("Contractor").filter = "ListName eq 'Tenderer'";
	 	fd.field("Contractor").refresh();	      
    });	

	var TradesLegend = "<table cellspacing='0' cellpadding='0' width='100%' border='1' style='border-collapse:collapse; font-size:11px;'>" +
         "<tr><td width='25%' style='background-color:lightgrey; font-weight:bold;text-align:center'>Trade</td><td width='25%' style='background-color:lightgrey; font-weight:bold'>Description</td><td width='25%' style='background-color:lightgrey; font-weight:bold;text-align:center'>Trade</td><td width='25%' style='background-color:lightgrey; font-weight:bold'>Description</td></tr>" +
         "<tr><td style='text-align:center'>AR</td><td>Architectural</td><td style='text-align:center'>ME</td><td>Mechanical</td></tr>" +
		 "<tr><td style='text-align:center'>Area</td><td>Area Office</td><td style='text-align:center'>PCS</td><td>Project Control Specialist</td></tr>" +
		 "<tr><td style='text-align:center'>CA</td><td>Contract Administrator</td><td style='text-align:center'>PM</td><td>Project Manager</td></tr>" +
		 "<tr><td style='text-align:center'>Client</td><td>Client</td><td style='text-align:center'>PMC</td><td>PMC General</td></tr>" +
		 "<tr><td style='text-align:center'>CP</td><td>Construction Specialist</td><td style='text-align:center'>PUD</td><td>Planning and Urban Design</td></tr>" +
		 "<tr><td style='text-align:center'>EC</td><td>Economics</td><td style='text-align:center'>QS</td><td>Quantity Surveyor</td></tr>" +
		 "<tr><td style='text-align:center'>EL</td><td>Electrical</td><td style='text-align:center'>SB</td><td>Structural</td></tr>" +
		 "<tr><td style='text-align:center'>GE</td><td>Geotechnical</td><td style='text-align:center'>TR</td><td>Transportation</td></tr>" +
		 "<tr><td style='text-align:center'>LAD</td><td>Landscape</td><td style='text-align:center'>WE</td><td>Water and Environment</td></tr>" +
         "</table>";
	$('#TradesLegend').append(TradesLegend);
	
	fd.toolbar.buttons[0].text = "Submit";
	fd.toolbar.buttons[1].text = "Cancel";

	fd.toolbar.buttons[0].style = 'background-color: #3CDBC0; color: white;';
    fd.toolbar.buttons[1].style = 'background-color: #c8c8c8; color: white;';	
		
	fd.field('Answer').$on('change', function(value){
		CheckAnserandHandle(value, GroupName);
	});
	CheckAnserandHandle(fd.field('Answer').value, GroupName);
		
	fd.field('IsDuplicate').$on('change', function(value){
		CheckIsDuplicate(value, "Answer");
	});
	CheckIsDuplicate(fd.field('IsDuplicate').value, "Answer");
	
	fd.field('BatchNoText').required = true;
	fd.field('EntryBatch').disabled = true;
	
	fd.field('BatchNoText').$on('change', async function(value){		
		if(isNaN(value))
		{			
			alert("Oops, please insert numerical value in Group No.");
			fd.field('EntryBatch').value = "";
		}
		else
		{
			if(value !== null)
			{
				limitCharacters('BatchNoText',2);
				var EntryBatch = String(value).padStart(2,'0');
				
				fd.field('EntryBatch').disabled = false;	
				fd.field('EntryBatch').value = "GR-" + EntryBatch;
				fd.field('EntryBatch').disabled = true;	
				
				$(fd.field('EntryUniqueNumber').$parent.$el).show();
				fd.field('EntryUniqueNumber').value = fd.field('Contractor').value.LookupValue + "-" + fd.field('EntryBatch').value + "-" + fd.field('SerialNumber').value;
				$(fd.field('EntryUniqueNumber').$parent.$el).hide();
				
				_Proceed = await IsEntryNumberExists();				
			}
		}			
	});
	
	
	fd.field('SerialNoText').required = true;
	fd.field('SerialNumber').disabled = true;
	
	fd.field('SerialNoText').$on('change', async function(value){			
		if(isNaN(value))
		{			
			alert("Oops, please insert numerical value in ID No.");
			fd.field('SerialNumber').value = "";
		}
		else
		{	
			if(value !== null)
			{	
				limitCharacters('SerialNoText', 4);	
				var SerialNumber = String(value).padStart(4,'0');
				
				fd.field('SerialNumber').disabled = false;	
				fd.field('SerialNumber').value = "ID-" + SerialNumber;
				fd.field('SerialNumber').disabled = true;
				
				$(fd.field('EntryUniqueNumber').$parent.$el).show();
				fd.field('EntryUniqueNumber').value = fd.field('Contractor').value.LookupValue + "-" + fd.field('EntryBatch').value + "-" + fd.field('SerialNumber').value;
				$(fd.field('EntryUniqueNumber').$parent.$el).hide();
				
				_Proceed = await IsEntryNumberExists();			
			}
		}		
	});	

	fd.field('Contractor').$on('change', async function(value){	
		
		if(value !== null)
		{
			$(fd.field('EntryUniqueNumber').$parent.$el).show();	
			fd.field('EntryUniqueNumber').value = value.Title + "-" + fd.field('EntryBatch').value + "-" + fd.field('SerialNumber').value;
			$(fd.field('EntryUniqueNumber').$parent.$el).hide();
			
			_Proceed = await IsEntryNumberExists();
		}
	});
}

var TQ_editForm = async function(){		

	fd.validators.push({
		name: 'MyCustomValidatorTenderCircularNo',
		error: "Tender Circular No. and Query No. for same Tenderer is already exists.",
		validate: function(value) {			
			if (_OutProceed > 0)
				return false;
			return true;			
		}
	});

	try{		
		$(fd.field('AttachFiles').$parent.$el).hide();	
		$(fd.field('SenttoClient').$parent.$el).hide();
	
		fd.toolbar.buttons[0].style="display: none;";
		fd.toolbar.buttons[1].text = "Cancel"; 
		fd.toolbar.buttons[1].style = "background-color: #c8c8c8; color: white;";
	
		$(fd.field('Reassigned').$parent.$el).hide();
		$(fd.field('AnsweredbyLead').$parent.$el).hide();
		$(fd.field('Reject').$parent.$el).hide();
		$(fd.field('SendToTM').$parent.$el).hide();
		$(fd.field('Answered').$parent.$el).hide();
		$(fd.field('IssuedDate').$parent.$el).hide();
		$(fd.field('ClientInternalNote').$parent.$el).hide();
		fd.field('Reviewedby').disabled = true;
		
		fd.field('Title').disabled = true;
		fd.field('QuerySerialNbr').disabled = true;	
		fd.field('EntryBatch').disabled = true;
		fd.field('SerialNumber').disabled = true;	
		fd.field('Package').disabled = true;
		fd.field('Contractor').disabled = true;
		
		$(fd.field('OutUniqueNumber').$parent.$el).hide();
		fd.field('NoticeNbrText').value = "";
		fd.field('QueryNoText').value = "";
		fd.field('NoticeNbr').value = "";
		fd.field('BidderQueryNo').value = "";
		
		$(fd.field('NoticeNbrText').$parent.$el).hide();
		$(fd.field('QueryNoText').$parent.$el).hide();
		$(fd.field('NoticeNbr').$parent.$el).hide();
		$(fd.field('BidderQueryNo').$parent.$el).hide();
		
		var isSiteAdmin = _spPageContextInfo.isSiteAdmin; 		
		const _firstVisitNumber = await GetColumnValueByID('WorkflowStatus');	

		var TQUsers = fd.field('TQAction').value;	
		let _isAllowed = false
		var currentUser = await pnp.sp.web.currentUser.get();	
		$(fd.field('TQAction').$parent.$el).hide();					

		var WorkflowStatus = fd.field('WorkflowStatus').value;
		fd.field('WorkflowStatus').disabled = true;
	
		const GroupName = await CheckifUserinTMorATM();	
			
		if (GroupName === "TM" || GroupName === "Admin")
		{
			if(WorkflowStatus === "Issued" && !isSiteAdmin)
				SetAttachmentToReadOnly();

			fd.field('LeadAnswer').disabled = true;
			
			$(fd.field('ClientInternalNote').$parent.$el).show();
			
			$(fd.field('NoticeNbrText').$parent.$el).show();
			$(fd.field('QueryNoText').$parent.$el).show();
			$(fd.field('NoticeNbr').$parent.$el).show();
			$(fd.field('BidderQueryNo').$parent.$el).show();		
			
			//fd.field('NoticeNbrText').required = true;
			fd.field('NoticeNbr').disabled = true;
	
			fd.field('NoticeNbrText').$on('change', async function(value){		
				if(isNaN(value))
				{			
					alert("Oops, please insert numerical value in Tender No.");
					fd.field('NoticeNbr').value = "";
				}
				else
				{
					if(value !== null)
					{
						limitCharacters('NoticeNbrText',2);
						var NoticeNbr = String(value).padStart(2,'0');
						
						fd.field('NoticeNbr').disabled = false;	
						fd.field('NoticeNbr').value = "TC-" + NoticeNbr;
						fd.field('NoticeNbr').disabled = true;	
						
						$(fd.field('OutUniqueNumber').$parent.$el).show();
						fd.field('Contractor').disabled = false;
						fd.field('OutUniqueNumber').value = fd.field('Contractor').value.LookupValue + "-" + fd.field('NoticeNbr').value + "-" + fd.field('BidderQueryNo').value;
						fd.field('Contractor').disabled = true;
						$(fd.field('OutUniqueNumber').$parent.$el).hide();
						
						_OutProceed = await IsOutNumberExists();				
					}
				}			
			});
			
			//fd.field('QueryNoText').required = true;
			fd.field('BidderQueryNo').disabled = true;
	
			fd.field('QueryNoText').$on('change', async function(value){			
				if(isNaN(value))
				{			
					alert("Oops, please insert numerical value in Query No.");
					fd.field('BidderQueryNo').value = "";
				}
				else
				{	
					if(value !== null)
					{	
						limitCharacters('QueryNoText',4);	
						var BidderQueryNo = String(value).padStart(4,'0');
						
						fd.field('BidderQueryNo').disabled = false;	
						fd.field('BidderQueryNo').value = "TQ-" + BidderQueryNo;
						fd.field('BidderQueryNo').disabled = true;
						
						$(fd.field('OutUniqueNumber').$parent.$el).show();
						fd.field('Contractor').disabled = false;
						fd.field('OutUniqueNumber').value = fd.field('Contractor').value.LookupValue + "-" + fd.field('NoticeNbr').value + "-" + fd.field('BidderQueryNo').value;
						fd.field('Contractor').disabled = true;
						$(fd.field('OutUniqueNumber').$parent.$el).hide();
						
						_OutProceed = await IsOutNumberExists();			
					}
				}		
			});		
			
			fd.field('LeadTrade').$on('change', function(value){
				$(fd.field('Reassigned').$parent.$el).show();
				fd.field('Reassigned').value = true;
				$(fd.field('Reassigned').$parent.$el).hide();
			});
			
			if(WorkflowStatus === "Issued")
			{			
				fd.toolbar.buttons.push({
						icon: 'Reply',
						class: 'btn-outline-primary',
						text: 'Rejected by Client',
						style: 'background-color: #FF5733; color: white;',
						click: function() {	    		
						
						$(fd.field('Reject').$parent.$el).show();
						fd.field('Reject').value = true;
						$(fd.field('Reject').$parent.$el).hide();	
						
						$(fd.field('Reassigned').$parent.$el).show();
						fd.field('Reassigned').value = false;
						$(fd.field('Reassigned').$parent.$el).hide();				
					
						$(fd.field('SenttoClient').$parent.$el).show();
						fd.field('SenttoClient').value = false;
						$(fd.field('SenttoClient').$parent.$el).hide();					
						
						fd.validators.push({
							name: 'Check Answer box',
							error: "Please write the reason of rejection in DAR Internal Note.",
							validate: function(value) {	
									if (fd.field('DarInternalNotes').value === null || fd.field('DarInternalNotes').value === "")
										return false;					
									
									return true;
								}	
							});			
									
						fd.save();			 
					 }
				});
			}	
			else if(WorkflowStatus === "Answered")
			{	
				//DisableNonTMColumens();
				
				fd.field('IsDuplicate').$on('change', function(value){
					CheckIsDuplicateEdit(value, "Answer");
				});
				CheckIsDuplicateEditDefaultState(fd.field('IsDuplicate').value, "Answer");
				
				fd.toolbar.buttons.push({
						icon: 'Save',
						class: 'btn-outline-primary',
						text: 'Save / Reassign',
						style: 'background-color: #e3d122; color: white;',
						click: function() {										
							
						//$(fd.field('Reassigned').$parent.$el).show();	
						$(fd.field('AnsweredbyLead').$parent.$el).show();
						$(fd.field('SendToTM').$parent.$el).show();
						$(fd.field('Answered').$parent.$el).show();
						$(fd.field('Reject').$parent.$el).show();
						
						//fd.field('Reassigned').value = false;		
						fd.field('AnsweredbyLead').value = false;	
						fd.field('SendToTM').value = false;			
						fd.field('Answered').value = false;			
						fd.field('Reject').value = false;
						
						//$(fd.field('Reassigned').$parent.$el).hide();	
						$(fd.field('AnsweredbyLead').$parent.$el).hide();
						$(fd.field('SendToTM').$parent.$el).hide();
						$(fd.field('Answered').$parent.$el).hide();
						$(fd.field('Reject').$parent.$el).hide();			
								
						fd.save();			 
					}
				});	
				
				fd.toolbar.buttons.push({
						icon: 'Reply',
						class: 'btn-outline-primary',
						text: 'Reject',
						style: 'background-color: #FF5733; color: white;',
						click: function() {	    		
						
						$(fd.field('Reject').$parent.$el).show();
						fd.field('Reject').value = true;
						$(fd.field('Reject').$parent.$el).hide();
						
						$(fd.field('Reassigned').$parent.$el).show();
						fd.field('Reassigned').value = false;
						$(fd.field('Reassigned').$parent.$el).hide();	
						
						fd.validators.push({
							name: 'Check Answer box',
							error: "Please write the reason of rejection in DAR Internal Note.",
							validate: function(value) {	
									if (fd.field('DarInternalNotes').value === null || fd.field('DarInternalNotes').value === "")
										return false;					
									
									return true;
								}	
							});			
									
						fd.save();			 
					 }
				});				
			}	
			else if(WorkflowStatus != "Issued")
			{	
			
				fd.field('IsDuplicate').$on('change', function(value){
					CheckIsDuplicateEdit(value, "Answer");
				});
				CheckIsDuplicateEditDefaultState(fd.field('IsDuplicate').value, "Answer");
				
				fd.toolbar.buttons.push({
						icon: 'Save',
						class: 'btn-outline-primary',
						text: 'Save / Reassign',
						style: 'background-color: #e3d122; color: white;',
						click: function() {
						
						var NCountofATT = 0;
						var OCountofATT = 0;
						for(i = 0; i < fd.field('Attachments').value.length; i++) 
						{
							var Val = fd.field('Attachments').value[i].extension.toString();
							if(Val === "")
							{OCountofATT++;}
							else
							{			
									NCountofATT++;			
							}
						}													
							
						//$(fd.field('Reassigned').$parent.$el).show();	
						$(fd.field('AnsweredbyLead').$parent.$el).show();
						$(fd.field('SendToTM').$parent.$el).show();
						$(fd.field('Answered').$parent.$el).show();
						$(fd.field('Reject').$parent.$el).show();
						
						//fd.field('Reassigned').value = false;		
						fd.field('AnsweredbyLead').value = false;	
						fd.field('SendToTM').value = false;			
						fd.field('Answered').value = false;			
						fd.field('Reject').value = false;
						
						//$(fd.field('Reassigned').$parent.$el).hide();	
						$(fd.field('AnsweredbyLead').$parent.$el).hide();
						$(fd.field('SendToTM').$parent.$el).hide();
						$(fd.field('Answered').$parent.$el).hide();
						$(fd.field('Reject').$parent.$el).hide();
						
						$(fd.field('AttachFiles').$parent.$el).show();
						fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
						$(fd.field('AttachFiles').$parent.$el).hide();			
							
						fd.field('QueryType').required = false;					
						fd.validators.length =0;	
						fd.save();			 
					}
				});
				
				if(WorkflowStatus === "Assigned" || WorkflowStatus === "Reassigned")
				{				
				}
				else
				{			
					fd.toolbar.buttons.push({
							icon: 'Reply',
							class: 'btn-outline-primary',
							text: 'Reject',
							style: 'background-color: #FF5733; color: white;',
							click: function() {	    		
							
							$(fd.field('Reject').$parent.$el).show();
							fd.field('Reject').value = true;
							$(fd.field('Reject').$parent.$el).hide();
							
							$(fd.field('Reassigned').$parent.$el).show();
							fd.field('Reassigned').value = false;
							$(fd.field('Reassigned').$parent.$el).hide();					
							
							fd.validators.push({
							name: 'Check Answer box',
							error: "Please write the reason of rejection in DAR Internal Note.",
							validate: function(value) {	
									if (fd.field('DarInternalNotes').value === null || fd.field('DarInternalNotes').value === "")
										return false;					
									
									return true;
								}	
							});
										
							fd.save();			 
						 }
					});
				}
				
				fd.toolbar.buttons.push({
					icon: 'Accept',
					class: 'btn-outline-primary',
					text: 'Submit to Client',
					style: 'background-color: #3CDBC0; color: white;',
					click: async function() {

						const _SecondVisitNumber = await GetColumnValueByID('WorkflowStatus');
						
						if(_SecondVisitNumber !== _firstVisitNumber)
						{
							alert("Please be informed that this record has already responded.");
							fd.close();
						}						
									
						var NCountofATT = 0;
						var OCountofATT = 0;
						for(i = 0; i < fd.field('Attachments').value.length; i++) 
						{
							var Val = fd.field('Attachments').value[i].extension.toString();
							if(Val === "")
							{OCountofATT++;}
							else
							{			
									NCountofATT++;			
							}
						}						
						
						$(fd.field('AttachFiles').$parent.$el).show();
						fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
						$(fd.field('AttachFiles').$parent.$el).hide();
						
						$(fd.field('Reassigned').$parent.$el).show();
						fd.field('Reassigned').value = false;
						$(fd.field('Reassigned').$parent.$el).hide();
						
						fd.validators.push({
						name: 'Check Answer box',
						error: "Please write your answer before submit final answer",
						validate: function(value) {	
								if (fd.field('Answer').value === null || fd.field('Answer').value === "")
									return false;					
								
								return true;
							}	
						});		
						
						fd.field('QueryType').required = true;	

						fd.validators.push({
						name: 'Check Query Type',
						error: "Please confirm that Query Type is correct.",
						validate: function(value) {	
								if (fd.field('QueryTypeChecked').value === false)
									return false;					
								
								return true;
							}	
						});		
						
						$(fd.field('Answered').$parent.$el).show();
						fd.field('Answered').value = true;
						$(fd.field('Answered').$parent.$el).hide();
						
						$(fd.field('AttachFiles').$parent.$el).show();
						fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
						$(fd.field('AttachFiles').$parent.$el).hide();
							
						fd.save();						 
					 }
				});
				
				/*if(fd.field('WorkflowStatus').value != "Assigned")
				{
					
				}*/
			}
		}
	
		else if (GroupName === "ATM")
		{	
			fd.field('LeadAnswer').disabled = true;
			DisableNonTMColumens();	

			if(WorkflowStatus !== "Sent to ATM")
				SetAttachmentToReadOnly();
			
			if(WorkflowStatus === "Issued")
			{			
				fd.toolbar.buttons.push({
						icon: 'Reply',
						class: 'btn-outline-primary',
						text: 'Reject by Client',
						style: 'background-color: #FF5733; color: white;',
						click: function() {	    		
						
						$(fd.field('Reject').$parent.$el).show();
						fd.field('Reject').value = true;
						$(fd.field('Reject').$parent.$el).hide();	
						
						$(fd.field('Reassigned').$parent.$el).show();
						fd.field('Reassigned').value = false;
						$(fd.field('Reassigned').$parent.$el).hide();				
					
						$(fd.field('SenttoClient').$parent.$el).show();
						fd.field('SenttoClient').value = false;
						$(fd.field('SenttoClient').$parent.$el).hide();					
						
						fd.validators.push({
							name: 'Check Answer box',
							error: "Please write the reason of rejection in DAR Internal Note.",
							validate: function(value) {	
									if (fd.field('DarInternalNotes').value === null || fd.field('DarInternalNotes').value === "")
										return false;					
									
									return true;
								}	
							});			
									
						fd.save();			 
					 }
				});
			}	
			
			else if(WorkflowStatus === "Assigned" || WorkflowStatus === "Reassigned" || WorkflowStatus === "Ready for Response" || WorkflowStatus === "Sent to ATM")
			{
			
				fd.field('IsDuplicate').$on('change', function(value){
						CheckIsDuplicateEdit(value, "Answer");
					});
				CheckIsDuplicateEditDefaultState(fd.field('IsDuplicate').value, "Answer");
			
				if(WorkflowStatus === "Assigned" || WorkflowStatus === "Reassigned")
				{
					fd.toolbar.buttons.push({
							icon: 'AddFriend',
							class: 'btn-outline-primary',
							text: 'Reassign',
							style: 'background-color: #e3d122; color: white;',
							click: function() {
							
							fd.field('QueryType').required = false;	
											
							var NCountofATT = 0;
							var OCountofATT = 0;
							for(i = 0; i < fd.field('Attachments').value.length; i++) 
							{
								var Val = fd.field('Attachments').value[i].extension.toString();
								if(Val === "")
								{OCountofATT++;}
								else
								{			
										NCountofATT++;			
								}
							}					
							$(fd.field('Reassigned').$parent.$el).show();
							fd.field('Reassigned').value = true;
							$(fd.field('Reassigned').$parent.$el).hide();
							
							$(fd.field('AttachFiles').$parent.$el).show();
							fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
							$(fd.field('AttachFiles').$parent.$el).hide();
							
							fd.validators.length =0;			
							fd.save();			 
						 }
					});
				}
				else
				{
					fd.field('LeadTrade').disabled = true;	
					fd.field('Trades').disabled = true;	
					
					fd.toolbar.buttons.push({
							icon: 'Reply',
							class: 'btn-outline-primary',
							text: 'Reject',
							style: 'background-color: #FF5733; color: white;',
							click: function() {	    		
							
							$(fd.field('Reject').$parent.$el).show();
							fd.field('Reject').value = true;
							$(fd.field('Reject').$parent.$el).hide();	
							
							fd.validators.push({
							name: 'Check Answer box',
							error: "Please write the reason of rejection in DAR Internal Note.",
							validate: function(value) {	
									if (fd.field('DarInternalNotes').value === null || fd.field('DarInternalNotes').value === "")
										return false;					
									
									return true;
								}	
							});
										
							fd.save();			 
						 }
					});
				}
				
				fd.toolbar.buttons.push({
						icon: 'Accept',
						class: 'btn-outline-primary',
						text: 'Send to TM',
						style: 'background-color: #3CDBC0; color: white;',
						click: async function() {

						const _SecondVisitNumber = await GetColumnValueByID('WorkflowStatus');
					
						if(_SecondVisitNumber !== _firstVisitNumber)
						{
							alert("Please be informed that this record has already responded.");
							fd.close();
						}
										
						var NCountofATT = 0;
						var OCountofATT = 0;
						for(i = 0; i < fd.field('Attachments').value.length; i++) 
						{
							var Val = fd.field('Attachments').value[i].extension.toString();
							if(Val === "")
							{OCountofATT++;}
							else
							{			
									NCountofATT++;			
							}
						}
						$(fd.field('AttachFiles').$parent.$el).show();
						fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
						$(fd.field('AttachFiles').$parent.$el).hide();
						
						fd.validators.push({
						name: 'Check Answer box',
						error: "Please write your answer before submit to TM.",
						validate: function(value) {	
								if (fd.field('Answer').value === null || fd.field('Answer').value === "")
									return false;					
								
								return true;
							}	
						});	

						fd.field('QueryType').required = true;
						
						fd.validators.push({
						name: 'Check Query Type',
						error: "Please confirm that Query Type is correct.",
						validate: function(value) {	
								if (fd.field('QueryTypeChecked').value === false)
									return false;					
								
								return true;
							}	
						});				
						
						$(fd.field('SendToTM').$parent.$el).show();
						fd.field('SendToTM').value = true;
						$(fd.field('SendToTM').$parent.$el).hide();
						
						$(fd.field('AttachFiles').$parent.$el).show();
						fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
						$(fd.field('AttachFiles').$parent.$el).hide();
							
						fd.save();						 
					 }
				});
				
				/*if(fd.field('WorkflowStatus').value != "Assigned")
				{
					
				}*/
			}
			else
			{		
				fd.field('LeadTrade').disabled = true;	
				fd.field('Trades').disabled = true;		
				fd.field('Answer').disabled = true;			
				//fd.field('Attachments').disabled = true;
				
				SetAttachmentToReadOnly();
			}		
		}
	
		else if (GroupName === "Trade")
		{	
			if(!isSiteAdmin)
			{		
				TQUsers.map(item =>{if(item.displayName == currentUser.Title)
						  _isAllowed = true;});
	
				if(!_isAllowed)
				{
					alert("Apologies, but this task has not been assigned to you.");
					fd.close();
				}		
			}
			//debugger;
			DisableNonTMColumens();
			fd.field('LeadTrade').disabled = true;	
			fd.field('Trades').disabled = true;	
			$(fd.field('Answer').$parent.$el).hide();
			$(fd.field('QueryTypeChecked').$parent.$el).hide();	
			//fd.field('Attachments').disabled = true;
			SetAttachmentToReadOnly();					
			
			if(WorkflowStatus === "Assigned" || WorkflowStatus === "Reassigned" || WorkflowStatus === "Ready for Response")
			{			
				fd.field('QueryType').required = false;
			
				fd.field('IsDuplicate').$on('change', function(value){
						CheckIsDuplicateEdit(value, "LeadAnswer");
					});
				CheckIsDuplicateEditDefaultState(fd.field('IsDuplicate').value, "LeadAnswer");
				
				if(WorkflowStatus === "Assigned" || WorkflowStatus === "Reassigned")
				{
					//fd.field('QueryType').required = false;
					
					/*
					fd.toolbar.buttons.push({
							icon: 'AddFriend',
							class: 'btn-outline-primary',
							text: 'Reassign',
							click: function() {						
											
							var NCountofATT = 0;
							var OCountofATT = 0;
							for(i = 0; i < fd.field('Attachments').value.length; i++) 
							{
								var Val = fd.field('Attachments').value[i].extension.toString();
								if(Val === "")
								{OCountofATT++;}
								else
								{			
										NCountofATT++;			
								}
							}						
							$(fd.field('Reassigned').$parent.$el).show();
							fd.field('Reassigned').value = true;
							$(fd.field('Reassigned').$parent.$el).hide();
							
							$(fd.field('AttachFiles').$parent.$el).show();
							fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
							$(fd.field('AttachFiles').$parent.$el).hide();		
							
							fd.validators.length =0;
							fd.save();			 
						 }
					});*/
				}
				else
				{
					fd.field('LeadTrade').disabled = true;	
					fd.field('Trades').disabled = true;	
				}
				
				fd.toolbar.buttons.push({
						icon: 'Accept',
						class: 'btn-outline-primary',
						text: 'Send to PMC',
						style: 'background-color: #3CDBC0; color: white;',
						click: async function() {
										
						//fd.field('QueryType').required = true;

						const _SecondVisitNumber = await GetColumnValueByID('WorkflowStatus');
						
						if(_SecondVisitNumber !== _firstVisitNumber)
						{
							alert("Please be informed that this record has already responded.");
							fd.close();
						}
						
						var NCountofATT = 0;
						var OCountofATT = 0;
						for(i = 0; i < fd.field('Attachments').value.length; i++) 
						{
							var Val = fd.field('Attachments').value[i].extension.toString();
							if(Val === "")
							{OCountofATT++;}
							else
							{			
									NCountofATT++;			
							}
						}
						$(fd.field('AttachFiles').$parent.$el).show();
						fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
						$(fd.field('AttachFiles').$parent.$el).hide();
						
						fd.validators.push({
						name: 'Check Answer box',
						error: "Please write your answer before submit to PMC.",
						validate: function(value) {	
								if (fd.field('LeadAnswer').value === null || fd.field('LeadAnswer').value === "")
									return false;					
								
								return true;
							}	
						});
						
						fd.validators.push({
							name: 'Check Query Type box',
							error: "Please select the Query Type before submit to PMC.",
							validate: function(value) {	
									if (fd.field('QueryType').value === null || fd.field('QueryType').value === "")
										return false;							
									return true;
								}	
							});			
					
						$(fd.field('AnsweredbyLead').$parent.$el).show();
						fd.field('AnsweredbyLead').value = true;
						$(fd.field('AnsweredbyLead').$parent.$el).hide();					
											
						fd.field('Reviewedby').disabled = false;
						fd.field('Reviewedby').value = currentUser.Title;
						fd.field('Reviewedby').disabled = true;					
						
						$(fd.field('AttachFiles').$parent.$el).show();
						fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
						$(fd.field('AttachFiles').$parent.$el).hide();
										
						fd.save();					 
					 }
				});			
			}
			else
			{
				fd.field('LeadAnswer').disabled = true;
				$(fd.field('Answer').$parent.$el).show();
				fd.field('Answer').disabled = true;
				fd.field('Trades').disabled = true;

				fd.field('IsDuplicate').disabled = true;
				fd.field('QueryLookupNumber').disabled = true;
				fd.field('QueryType').disabled = true;
				fd.field('DarInternalNotes').disabled = true;
				//fd.field('Attachments').disabled = true;
				
				SetAttachmentToReadOnly();
			}
		}
	
		CustomListEditor(GroupName);	
		
		//debugger;
		//var element = document.getElementById('SuiteNavPlaceHolder');
		//element.style.marginTop = '0px';	
		
		//const topBarCommandBar = document.querySelector('.shortpoint-focusMode-enabled');	
		//topBarCommandBar.classList.remove('shortpoint-focusMode-enabled--state-on');	
		
		/*
		var elements = document.getElementsByClassName('commandBarWrapper');
		for (var i = 0; i < elements.length; i++) {
		  var element = elements[i];
		  element.style.marginTop = '0px';
		  element.style.opacity = '10';
		}
		*/
	
		//const topBarCommandBar = document.querySelector('.shortpoint-focusMode-enabled.shortpoint-focusMode-enabled');	
		
		
		//.shortpoint-focusMode-enabled.shortpoint-focusMode-enabled--state-on .commandBarWrapper
		//commandBarWrapper shortpoint-proxy-neutral-light--bdr shortpoint-proxy-theme-scroll-content shortpoint-proxy-theme-scroll-mobile-wrapper	
		//topBarCommandBar.classList.add('shortpoint-focusMode-enabled--state-off');	
		//.shortpoint-focusMode-enabled.shortpoint-focusMode-enabled--state-on .od-TopBar-commandBar
		//commandBarWrapper.classList.add('shortpoint-focusMode-enabled--state-off');	
		
		//var isSiteAdmin = _spPageContextInfo.isSiteAdmin; 
		if(isSiteAdmin)
			ReopenallfieldsandSaveButton();
		
	}
	catch(e)
	{
		alert(e);
	} 

}

var GetAllPckagesByProjectNumber = async function(ProjectNo) 
{	
	var siteUrl = _spPageContextInfo.siteAbsoluteUrl;	
	var xhr = new XMLHttpRequest();
	var SERVICE_URL = siteUrl+ "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=getprojectpackages&ProjectCode=" + ProjectNo;	

	xhr.open("GET", SERVICE_URL, true);	
	
	xhr.onreadystatechange = async function() 
	{
		if (xhr.readyState == 4) 
		{
			var text;

			try 
			{
				if (xhr.status == 200)
				{
					if(this.responseText !== "")
					{
						const obj =  await JSON.parse(this.responseText, async function (key, value) {
							var _columnName = key;
							var _value = value;					

							if(_columnName === 'PackageName'){
								if(!PackageArr.includes(_value))							
								PackageArr.push(_value);						
							}						
						});					

						fd.field('PackageDropDown').widget.setDataSource({data: PackageArr});
					}
					
					if (PackageArr.length > 0)
						fd.field('PackageDropDown').required = true;

					else
					{
						fd.field('PackageDropDown').required = false;
						fd.field('PackageDropDown').widget.dataSource.data(['']);
						fd.field('PackageDropDown').disabled = true;						  
					}						
				}
				else
					console.log('Request failed with status code:', xhr.status);
			}
			catch(err) 
			{
				console.log(err + "\n" + text);				
			}
		}
	}			

	xhr.send();
}

async function CheckifUserinTMorATM() {
	var IsTMUser = "Trade";
	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(async function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {				
					if(groupsData[i].Title === "TM")
					{					
					   IsTMUser = "TM";
					   break;
				    }
					else if(groupsData[i].Title === "ATM")
					{					
					   IsTMUser = "ATM";
					   break;
				    }
					else
					{					
					    IsTMUser = "Trade";					   
				    }
				}				
			});
	     });
    }
	catch(e){alert(e);}

	var isSiteAdmin = _spPageContextInfo.isSiteAdmin; 
	if(isSiteAdmin)
		IsTMUser = "Admin";
		
	return IsTMUser;				
}

function CheckAnserandHandle(value, GroupName) 
{	
	if(value === null || value === "" || value === '')
	{
		fd.field('LeadTrade').disabled = false;
		fd.field('LeadTrade').required = true;
		fd.field('Trades').disabled = false;		
		
		$(fd.field('LeadTrade').$parent.$el).show();
		$(fd.field('Trades').$parent.$el).show();
		
		if(GroupName === "TM")
		{
			$(fd.field('Answered').$parent.$el).show();
			fd.field('Answered').value = false;
			$(fd.field('Answered').$parent.$el).hide();	
		}
		else if(GroupName === "ATM")
		{
			$(fd.field('SendToTM').$parent.$el).show();
			fd.field('SendToTM').value = false;
			$(fd.field('SendToTM').$parent.$el).hide();	
		}
		
		fd.field('QueryType').required = false;
		fd.field('QueryType').value = "";
		$(fd.field('QueryType').$parent.$el).hide();		
	}
	else
	{
		fd.field('LeadTrade').value = null;
		fd.field('LeadTrade').disabled = true;
		fd.field('LeadTrade').required = false;
		fd.field('Trades').value = null;
		fd.field('Trades').disabled = true;		
		
		//$(fd.field('LeadTrade').$parent.$el).hide();
		//$(fd.field('Trades').$parent.$el).hide();
		
		if(GroupName === "TM")
		{
			$(fd.field('Answered').$parent.$el).show();
			fd.field('Answered').value = true;
			$(fd.field('Answered').$parent.$el).hide();	
		}
		else if(GroupName === "ATM")
		{
			$(fd.field('SendToTM').$parent.$el).show();
			fd.field('SendToTM').value = true;
			$(fd.field('SendToTM').$parent.$el).hide();	
		}
		
		$(fd.field('QueryType').$parent.$el).show();
		fd.field('QueryType').required = true;
		//$(fd.field('QueryType').$parent.$el).hide();
		
		$(fd.field('ClientInternalNote').$parent.$el).show();		
	}
}

function CheckIsDuplicate(value, AnswerColumnName) 
{
	if(value === true)
	{
		$(fd.field('QueryLookupNumber').$parent.$el).show();
		fd.field('QueryLookupNumber').ready().then(function() {
		    fd.field('QueryLookupNumber').filter = "QuerySerialNbr ne null"; 		
		    fd.field('QueryLookupNumber').refresh();
	  	});
		fd.field('QueryLookupNumber').required = true;
		
		fd.field('QueryLookupNumber').$on('change', function(value){
			FillInAnswer(value.LookupValue, AnswerColumnName);
		});
		
		if(fd.field('QueryLookupNumber').value != null)
			FillInAnswer(fd.field('QueryLookupNumber').value.LookupValue, AnswerColumnName);	
	}
	else
	{
		fd.field('QueryLookupNumber').required = false;
		$(fd.field('QueryLookupNumber').$parent.$el).hide();
		
		fd.field(AnswerColumnName).value = "";
		fd.field(AnswerColumnName).disabled = false;
		
		fd.field('QueryType').value = "";
		fd.field('QueryType').disabled = false;	
		$(fd.field('QueryType').$parent.$el).hide();			
	}
}

function FillInAnswer(value, AnswerColumnName) 
{
	fd.field(AnswerColumnName).value = "Tenderer shall refer to answer to Query no. " + value + ".";
	//fd.field(AnswerColumnName).disabled = true;
	
	fd.field('QueryType').value = "Duplicate";
	fd.field('QueryType').disabled = true;
}

const IsEntryNumberExists = async function(){
	
	let result = "";
	
	$(fd.field('EntryUniqueNumber').$parent.$el).show();
	var EntryUniqueNumber = fd.field('EntryUniqueNumber').value;
	$(fd.field('EntryUniqueNumber').$parent.$el).hide();	
	var camlF = "EntryUniqueNumber eq '" + EntryUniqueNumber + "'";	
	
	await pnp.sp.web.lists.getByTitle('QApp').items.select("EntryUniqueNumber").filter(camlF).get().then(function(items)
	{
		result = items.length;						 
	}); 
	
	return result;	
}

function limitCharacters(fieldName, value){
       var newText = fd.field(fieldName).value;
       if(fd.field(fieldName).value.length > value) {
            fd.field(fieldName).value = text;
       }
       else {
           text = newText;
       }
}

function ReopenallfieldsandSaveButton() {
	$(fd.field('ClientInternalNote').$parent.$el).show();
	fd.field('LeadTrade').disabled = false;	
	fd.field('Trades').disabled = false;		
	fd.field('Answer').disabled = false;	
	fd.field('DarInternalNotes').disabled = false;		
	fd.field('Attachments').disabled = false;
	fd.field('Question').disabled = false;	
	//fd.field('Package').disabled = false;		
	//fd.field('Contractor').disabled = false;	
	fd.field('LetterReference').disabled = false;
	fd.field('LeadAnswer').disabled = false;	
	
	fd.toolbar.buttons.push({
		        icon: 'Save',
		        class: 'btn-outline-primary',
		        text: 'Save',
		        click: function() {										
					
				$(fd.field('Reassigned').$parent.$el).show();	
				$(fd.field('AnsweredbyLead').$parent.$el).show();
				$(fd.field('SendToTM').$parent.$el).show();
				$(fd.field('Answered').$parent.$el).show();
				$(fd.field('Reject').$parent.$el).show();
				
				fd.field('Reassigned').value = false;		
				fd.field('AnsweredbyLead').value = false;	
				fd.field('SendToTM').value = false;			
				fd.field('Answered').value = false;			
				fd.field('Reject').value = false;
				
				$(fd.field('Reassigned').$parent.$el).hide();	
				$(fd.field('AnsweredbyLead').$parent.$el).hide();
				$(fd.field('SendToTM').$parent.$el).hide();
				$(fd.field('Answered').$parent.$el).hide();
				$(fd.field('Reject').$parent.$el).hide();			
						
	            fd.save();			 
			}
		});
	
}
	
function DisableNonTMColumens() {
	fd.field('Question').disabled = true;	
	fd.field('Package').disabled = true;		
	fd.field('Contractor').disabled = true;	
	fd.field('LetterReference').disabled = true;	
}	
	
function CustomListEditor(GroupName) {				

	if (fd.field('WorkflowStatus').value === 'Answered' || fd.field('WorkflowStatus').value === 'Issued') 
	{		
		fd.field('LeadTrade').disabled = true;	
		fd.field('Trades').disabled = true;	
		
		if(fd.field('WorkflowStatus').value == 'Issued' && (GroupName === "TM" || GroupName === "ATM"))
		{
			fd.field('ClientInternalNote').disabled = true;
			fd.field('DarInternalNotes').disabled = false;
			fd.field('Attachments').disabled = false;
		}		
		else if(fd.field('WorkflowStatus').value == 'Answered' && GroupName === "TM")
		{
			fd.field('Attachments').disabled = false;
		}
		else
		{
			fd.field('Answer').disabled = true;	
			fd.field('DarInternalNotes').disabled = true;
			fd.field('QueryLookupNumber').disabled = true;
			fd.field('IsDuplicate').disabled = true;
			//fd.field('Attachments').disabled = true;
			
			var isSiteAdmin = _spPageContextInfo.isSiteAdmin;
			
			if(!isSiteAdmin)
				SetAttachmentToReadOnly();
		}				
	}			
}

function DisableAttachment() {
	fd.field('Attachments').disabled = false;
	
	$('div.k-upload-button').remove();
	$('button.k-upload-action').remove();
	$('.k-dropzone').remove();
}

function CheckIsDuplicateEdit(value, AnswerColumnName) 
{
	if(value === true)
	{
		$(fd.field('QueryLookupNumber').$parent.$el).show();
		
		fd.field('QueryLookupNumber').ready().then(function() {
		    fd.field('QueryLookupNumber').filter = "QuerySerialNbr ne'" + fd.field('QuerySerialNbr').value + "'"; 			
		    fd.field('QueryLookupNumber').refresh();
		});
		
		fd.field('QueryLookupNumber').required = true;
		
		fd.field('QueryLookupNumber').$on('change', function(value){
			FillInAnswer(value.LookupValue, AnswerColumnName);
		});
		
		if(fd.field('QueryLookupNumber').value != null)
			FillInAnswer(fd.field('QueryLookupNumber').value.LookupValue, AnswerColumnName);	
	}
	else
	{
		fd.field('QueryLookupNumber').required = false;
		$(fd.field('QueryLookupNumber').$parent.$el).hide();
		
		fd.field(AnswerColumnName).value = "";
		fd.field(AnswerColumnName).disabled = false;
		
		fd.field('QueryType').value = "";
		fd.field('QueryType').disabled = false;	
		//$(fd.field('QueryType').$parent.$el).hide();			
	}
}

function CheckIsDuplicateEditDefaultState(value, AnswerColumnName) 
{
	if(value === true)
	{
		$(fd.field('QueryLookupNumber').$parent.$el).show();
		
		fd.field('QueryLookupNumber').ready().then(function() {
		    fd.field('QueryLookupNumber').filter = "DARIDNo ne'" + fd.field('DARIDNo').value + "'"; 			
		    fd.field('QueryLookupNumber').refresh();
		});
		
		fd.field('QueryLookupNumber').required = true;
		
		fd.field('QueryLookupNumber').$on('change', function(value){
			FillInAnswer(value.LookupValue, AnswerColumnName);
		});
		
		if(fd.field('QueryLookupNumber').value != null)
			FillInAnswer(fd.field('QueryLookupNumber').value.LookupValue, AnswerColumnName);	
	}
	else
	{
		fd.field('QueryLookupNumber').required = false;
		$(fd.field('QueryLookupNumber').$parent.$el).hide();			
	}
}

const IsOutNumberExists = async function(){
	
	let result = "";
	
	$(fd.field('OutUniqueNumber').$parent.$el).show();
	var OutUniqueNumber = fd.field('OutUniqueNumber').value;
	$(fd.field('OutUniqueNumber').$parent.$el).hide();	
	var camlF = "OutUniqueNumber eq '" + OutUniqueNumber + "'";	
	
	await pnp.sp.web.lists.getByTitle('QApp').items.select("OutUniqueNumber").filter(camlF).get().then(function(items)
	{
		result = items.length;						 
	}); 
	
	return result;	
}

async function GetItemVersionnumber(){
	
	var VersionNumber;
	var listUrl = fd.webUrl + fd.listUrl;
	var id = fd.itemId;
	await pnp.sp.web.getList(listUrl)
		.items
		.getById(id)
		.versions
		.get()
		.then(function(versions){
			self.entries = versions.map(function(v) {
				VersionNumber = v.VersionLabel;				
			})
		});

		return VersionNumber;
}
//#endregion

//#region Not Used
function populateCountryCodesFromWeb()
{    
    // $(fd.field('IsCodeAvailable').$parent.$el).hide();
    
    // //specify your site URL
    // var siteURL = countryURL;
    // let web = new Web(siteURL);

	// var country = fd.field('Country').value; //"Lebanon";//
	// if(country === "") return;
	// if(country === "International") 
	// {
    //     		web.lists.getByTitle('International Codes').items.select('Title').getAll().then(function(items) {
	// 		fd.field('DropDownCC').widget.setDataSource({
	// 			data: items.map(function(i) { return i.Title })
	// 		});
	// 		//set the dropdown with the previously selected value
	// 		var ExTitle = fd.field('Title').value;
	// 		if(ExTitle != "") 
	// 			fd.field('DropDownCC').widget.value(ExTitle);			
	// 	});
	// 	$(fd.field('Description').$parent.$el).show();
	// 	return;
	// }	
	// else 
	// {
	// 	web.lists.getByTitle('Country Codes').items.orderBy('Title', true).filter("Country eq '" + country + "' and ApprovalStatus eq 'Approved' and ( Status eq 'Valid' or Status eq 'Uncontrolled' )").select('Title').getAll().then(function(items) {
	// 		fd.field('DropDownCC').widget.setDataSource({
	// 			data: items.map(function(i) { return i.Title })
	// 		});
	// 		//set the dropdown with the previously selected value
	// 		var ExTitle = fd.field('Title').value;
	// 		if(ExTitle != "") 
	// 			fd.field('DropDownCC').value = ExTitle;			
	// 	});
	// }
}

function GetCountryCodeDetailsFromWeb(code) {
    
    // let web = new Web(countryURL);
	// var param = encodeURIComponent(code);
	// web.lists.getByTitle("Country Codes").items.filter("Title eq '" + param + "'").select("Title", "Description", "Hyperlink").getAll().then(function(items) {
	// 	//debugger;
	// 	if (items.length > 0) {
			
	// 		$(fd.field('Description').$parent.$el).show();
	//         $(fd.field('Hyperlink').$parent.$el).show();
    //         fd.field('Description').disabled = false;
	//         fd.field('Hyperlink').disabled = false;
			
	// 		fd.field('Description').value = items[0].Description;
	// 		var folderURL = items[0].Hyperlink;
			
	// 		var SiteUrl = window.location.protocol + "//" + window.location.host
	// 	    fd.field('Hyperlink').value.url = SiteUrl + folderURL;//relativePath;
	// 		fd.field('Hyperlink').value.description = 'go to folder';
			
	// 	   fd.field('Description').disabled = true;
	//        fd.field('Hyperlink').disabled = true;
	// 	}
	// });
}

function GetInternationalCodeDetailsFromWeb(code) {
	// let web = new Web(countryURL);
	// var param = encodeURIComponent(code);
	// web.lists.getByTitle('International Codes').items.filter("Title eq '" + param + "'").select('Title', "Description", "Hyperlink").get().then(function(items) {
		
	// 	if (items.length > 0) {
	// 		$(fd.field('Description').$parent.$el).show();
	// 		fd.field('Description').value = items[0].Description;
			
	// 		if(param !== "Others")
	// 			fd.field('Description').disabled = true;
	// 	    else
	// 			fd.field('Description').disabled = false;
		
	// 		$(fd.field('Hyperlink').$parent.$el).show();
	// 		fd.field('Hyperlink').value.url = items[0].Hyperlink;
	// 		fd.field('Hyperlink').value.description = 'go to folder';
	// 		fd.field('Hyperlink').disabled = true;
	// 	/*	*/
	// 	}
	// });
}
//#endregion

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

async function DCR_IsUserInGroup(formType) 
{
    var IsPMUser = "Trade";
	try{
		var userId = _spPageContextInfo.userId;
		await pnp.sp.web.siteUsers.getById(userId).groups.get().then(function(groupsData)
		{
			for (var i = 0; i < groupsData.length; i++) 
			{			
				if(groupsData[i].Title === 'PM' || groupsData[i].Title === 'Area')
				{
					IsPMUser = groupsData[i].Title;

					fd.toolbar.buttons[0].style = "display: none;";
					// fd.toolbar.buttons[0].icon = 'Save';
					// fd.toolbar.buttons[0].text = "Save";

					// fd.toolbar.buttons.push({
					// 	icon: 'Save',
					// 	class: 'btn-outline-primary',
					// 	text: 'Save',
					// 	click: function() {
					// 		fd.field('Status').value = 'Pending';
					// 		AttachFiles();
					// 		fd.save();
					// 	}
					// }); 
					
					fd.toolbar.buttons.push({
						icon: 'DocumentApproval',
						class: 'btn-outline-primary',
						text: 'Authorized',
						click: function() {
							
							$(fd.field('Status').$parent.$el).show();
							fd.field('Status').value = 'Authorized';
							$(fd.field('Status').$parent.$el).hide();

							if(formType == 'Edit'){
								$(fd.field('Reviewed').$parent.$el).show();
								fd.field('Reviewed').value = false;
								$(fd.field('Reviewed').$parent.$el).hide();
							}

							AttachFiles();
							fd.save();
							//setTimeout(function(){ fd.save(); }, 300);
						}
					});					

					if(formType == 'Edit')
					{
						fd.toolbar.buttons.push({
							icon: 'DocumentApproval',
							class: 'btn-outline-primary',
							text: 'Unauthorized',
							click: function() {
								$(fd.field('Status').$parent.$el).show();
								fd.field('Status').value = 'Unauthorized';
								$(fd.field('Status').$parent.$el).hide();

								$(fd.field('Reviewed').$parent.$el).show();
								fd.field('Reviewed').value = false;
								$(fd.field('Reviewed').$parent.$el).hide();

								AttachFiles();
								fd.save();
								//setTimeout(function(){ fd.save(); }, 300);
							}
						});
					}	
				}								
			}
				//alert("IsPMUser = " + IsPMUser);
		});	
		 
		//IsPMUser === "Trade" && 
		if(formType === 'Edit'){
			fd.field('Country').disabled = true;	   
			fd.field('Title').disabled = true;
			fd.field('Description').disabled = true;
			$(fd.field('Status').$parent.$el).show();
			fd.field('Status').disabled = true;

			SetAttachmentToReadOnly();
		}
		else if(IsPMUser === "Trade" && formType === 'New')
		{
			$(fd.field('Status').$parent.$el).show();
			fd.field('Status').value = 'Pending';
			$(fd.field('Status').$parent.$el).hide();
		}
    }
	catch(e){alert(e);}
	return IsPMUser;
}

function new_beforeSaveDCR(){
    try
    {
     var country = fd.field('Country').value;
     var codeTitle = fd.field('Title').value;
     var countryPath = "";
     
      //$(fd.field('Hyperlink').$parent.$el).show();
      $(fd.field('IsCodeAvailable').$parent.$el).show();
       if(fd.field('DropDownCC') != "")
        fd.field('IsCodeAvailable').value = true;
       else
         {
            alert("Existing Code is required");
            return;
         }
        
        AttachFiles();
        
       $(fd.field('IsCodeAvailable').$parent.$el).hide();
       $(fd.field('Hyperlink').$parent.$el).hide();
       
       return fd._vue.$nextTick();
    }
    catch(e){alert(e);}
}

async function DCR_newForm(formType){
    
	$(fd.field('AttachFiles').$parent.$el).hide();
	$(fd.field('Title').$parent.$el).hide();
	$(fd.field('Description').$parent.$el).hide();
	$(fd.field('Hyperlink').$parent.$el).hide();
	$(fd.field('Status').$parent.$el).hide();
	//fd.field('IsStatutoryAuthReq').value = false;

    fd.toolbar.buttons[1].icon = 'Cancel';
    fd.toolbar.buttons[1].text = "Cancel";

	fd.toolbar.buttons[0].icon = 'Accept';
    fd.toolbar.buttons[0].text = "Submit";
	await DCR_IsUserInGroup(formType, 'PM');

	var webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
	var projectName = webUrl.substring(webUrl.lastIndexOf('/') + 1);	

	var content = "Please <a href='" + countryURL + "/SitePages/Main-Page.aspx' target='_blank'>click here</a> to search the complete list of Country Codes or <a href='" + countryURL + "/SitePages/PlumsailForms/CountryCodeRequest/Item/NewForm.aspx?proj=" + projectName + "' target='_blank'>click here</a> to request a new code.";
	$('#my-html').append(content);

	var country = fd.field('Country').value;
	 
	if(country !== "International")
		document.getElementById('my-html').style.visibility = 'visible';
	else
		document.getElementById('my-html').style.visibility = 'hidden';
	 
	//call function on from load
    await populateCountryCodes();

	const TradeName = await CheckifUserinSPGroup();

	if(TradeName != null || TradeName != "")
	{		
	    fd.field('Trade').value = TradeName;
		// fd.field('Trade').disabled = true;
	}

    //fill SharePoint field with the selected value
    fd.field('DropDownCC').$on('change', function(value) {
		$(fd.field('Description').$parent.$el).hide();
		$(fd.field('Hyperlink').$parent.$el).hide();				

        fd.field('Title').value = value;
		var country = fd.field('Country').value;      
		if(country === "International")
			GetInternationalCodeDetails(value);
		else
			GetCountryCodeDetails(value);
    });
	
    fd.field('DropDownCC').addValidator({
        name: 'Array Count',
        error: 'Only one Code can be selected per form, save this form then add another Code.',
        validate: function(value) {
            if(fd.field('DropDownCC').value.length > 1) {
                return false;
            }
            return true;
        }
    });	
	
    fd.field('Country').$on('change', async function(value) {
		fd.field('Title').value = "";
		fd.field('Description').value = "";
		fd.field('DropDownCC').value = "";	
		
		//var country = fd.field('Country').value;		
		
		if(value !== "International")
		 	 document.getElementById('my-html').style.visibility = 'visible';
		else
			document.getElementById('my-html').style.visibility = 'hidden';
		
		await populateCountryCodes();
    });
}

async function DCR_editForm(formType){
    
    $(fd.field('AttachFiles').$parent.$el).hide();
	$(fd.field('Title').$parent.$el).hide();
	$(fd.field('IsCodeAvailable').$parent.$el).hide();

    fd.toolbar.buttons[1].icon = 'Cancel';
    fd.toolbar.buttons[1].text = "Cancel";
	
	//call function on from load
    populateCountryCodes();

    //fill SharePoint field with the selected value
    fd.field('DropDownCC').$on('change', function() {
		var code = fd.field("DropDownCC").value;
        fd.field('Title').value = code;
		//GetCountryCodeDetails(code);
		var country = fd.field('Country').value;
		if(country === "International")
			GetInternationalCodeDetails(code);
		else
			GetCountryCodeDetails(code);
    });
	
    fd.field('Country').$on('change', function() {
		fd.field('DropDownCC').value = "";
		fd.field('Title').value = "";
		fd.field('Description').value = "";		
		populateCountryCodes();
    });  
	
    var Status = fd.field('Status').value;
	var Reviewed = fd.field('Reviewed').value;   

	$(fd.field('Title').$parent.$el).show();
	$(fd.field('Status').$parent.$el).hide();
	$(fd.field('Reviewed').$parent.$el).hide();
	$(fd.field('CodeType').$parent.$el).hide();
	$(fd.field('DropDownCC').$parent.$el).hide();
	$(fd.field('ApprovalDate').$parent.$el).hide();
	$(fd.field('ApprovedBy').$parent.$el).hide();

	fd.field('Title').title = 'Title';

    //fd.field('IsCodeAvailable').$on('change',CodeAvailable);	
	
	if(Status == 'Pending' || Status == '')
	{ 		
		fd.toolbar.buttons[0].style = "display: none;";
	    await DCR_IsUserInGroup(formType);

	    $(fd.field('ApprovalDate').$parent.$el).hide();
	    $(fd.field('ApprovedBy').$parent.$el).hide();	   
    }
    else
    {
		fd.toolbar.buttons[0].style = "display: none;";
		
        fd.field('Country').disabled = true;	   
	    fd.field('Title').disabled = true;
	    fd.field('Description').disabled = true;
		$(fd.field('Status').$parent.$el).show();
		fd.field('Status').disabled = true;

		SetAttachmentToReadOnly();
    }
}

function edit_beforeSaveDCR(){
    var Status = fd.field('Status').value;
		var Reviewed = fd.field('Reviewed').value;
		
		if(Status == "Valid" || Status == "Invalid")
		{
		  var date = new Date();
	      var todayDate = String(date.getDate()).padStart(2, '0') + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + date.getFullYear();
		  fd.field('ApprovalDate').value = todayDate;
		  fd.field('ApprovedBy').value = _spPageContextInfo.userDisplayName;
		  fd.field('Reviewed').value = false;
		}
	 return fd._vue.$nextTick();
}

async function CheckifUserinSPGroup() {
	var IsTMUser = "";
	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(async function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {				
					if(groupsData[i].Title === "PM")
					{					
					   IsTMUser = "PM";
					   break;
				    }
					else if(groupsData[i].Title === "AR")
					{					
					   IsTMUser = "AR";
					   break;
				    }
					else if(groupsData[i].Title === "EC")
					{					
					   IsTMUser = "EC";
					   break;
				    }	
					else if(groupsData[i].Title === "EL")
					{					
					   IsTMUser = "EL";
					   break;
				    }	
					else if(groupsData[i].Title === "GE")
					{					
					   IsTMUser = "GE";
					   break;
				    }
					else if(groupsData[i].Title === "LAD")
					{					
					   IsTMUser = "LAD";
					   break;
				    }	
					else if(groupsData[i].Title === "ME")
					{					
					   IsTMUser = "ME";
					   break;
				    }	
					else if(groupsData[i].Title === "PMC")
					{					
					   IsTMUser = "PMC";
					   break;
				    }
					else if(groupsData[i].Title === "PUD")
					{					
					   IsTMUser = "PUD";
					   break;
				    }
					else if(groupsData[i].Title === "WE" || groupsData[i].Title === "RE")
					{					
					   IsTMUser = "WE";
					   break;
				    }
					else if(groupsData[i].Title === "SB" || groupsData[i].Title === "ST")
					{					
					   IsTMUser = "SB";
					   break;
				    }
					else if(groupsData[i].Title === "TR")
					{					
					   IsTMUser = "TR";
					   break;
				    }	
				}				
			});
	     });
    }
	catch(e){alert(e);}
	return IsTMUser;				
}

function addLegend(){
    if(  _formType === 'New' || _formType === 'Edit'){
      var id = '#legendlbl';
      if ($(id).length === 0) {
  
        var textNote = `Click 'Save' to keep your work as a draft, or 'Submit' to officially submit your answer.`;
  
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
}

function delay(time) {
	return new Promise(function(resolve) { 
		setTimeout(resolve, time)
	});
}

var GetColumnValueByID = async function (ColumName){
	var statusValue;
	var ItemID = fd.itemId;
	var listUrl = fd.webUrl + fd.listUrl;

	await pnp.sp.web.getList(listUrl).items.getById(ItemID).select(ColumName).get().then((item) => {		
		statusValue = item[ColumName];		
	  }).catch((error) => {
		console.error("Error fetching item: ", error);
	});

	return statusValue;
}

function fixTextArea(){
	$("textarea").each(function(index){
		var height = (this.scrollHeight + 5) + "px";
        $(this).css('height', height);
	});
}
//#endregion

var loadScripts = async function(){
	const libraryUrls = [
		_layout + '/plumsail/js/commonUtils.js',
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

var onDDSRender = async function (formType){
	if(formType == 'New' || formType == 'Edit'){
        fd.toolbar.buttons[0].icon = 'Accept';
        fd.toolbar.buttons[0].text = "Submit";
        fd.toolbar.buttons[1].icon = 'Cancel';
        fd.toolbar.buttons[1].text = "Cancel";

		let query = `Title ne null and IsDeliverableModule eq '1'`;
	
		let result = await pnp.sp.web.lists.getByTitle('Trades').items.select('IsDeliverableModule').get()
		.catch(err =>{
			query = 'Title ne null'
		});

		fd.field('Trades').ready().then(() => {
			fd.field('Trades').filter = query
			fd.field('Trades').orderBy = { field: 'Title', desc: false };
			fd.field('Trades').refresh();
		});
	}
    
	if(formType == 'Edit'){
		const Status = await GetColumnValueByID('Status');
		fd.field('Stat').value = Status;
		fd.field('Stat').disabled = true;

		const Reference = await GetColumnValueByID('Reference');
		fd.field('Ref').value = Reference;
		fd.field('Ref').disabled = true;
		
		if(Status !== "Reviewed"){}
		else
		{
			fd.toolbar.buttons[0].style = "display: none;";
			fd.field('Trades').disabled = true;
			fd.field('DueDate').disabled = true;
			fd.field('Comment').disabled = true;
		}
    }    
    if(formType == 'Display'){     
        fd.toolbar.buttons[1].icon = 'Cancel';
        fd.toolbar.buttons[1].text = "Cancel"; 

		const Status = await GetColumnValueByID('Status');
		fd.field('Stat').value = Status;
		fd.field('Stat').disabled = true;

		const Reference = await GetColumnValueByID('Reference');
		fd.field('Ref').value = Reference;
		fd.field('Ref').disabled = true;
    }
}

var PreloaderScripts = async function(){
    await _spComponentLoader.loadScript(_layout + '/controls/preloader/jquery.dim-background.min.js');
    await _spComponentLoader.loadScript(_layout + '/plumsail/js/preloader.js');
    preloader();
}