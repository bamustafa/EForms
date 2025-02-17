var _layout = "/_layouts/15/PCW/General/EForms";
var _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5];
var _ListFullUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/" + _ListInternalName;

var _modulename = "", _formType = "",  _AssignedTo = "";

var _isExist = false;
let CurrentUser;

var _status;

const itemsToRemove = ['Status', 'State', 'Code', 'WorkflowStatus'];

var onRender = async function (moduleName, formType, relativeLayoutPath){
	// localStorage.clear();	
	try {	

		if(relativeLayoutPath !== undefined && relativeLayoutPath !== null && relativeLayoutPath !== '')
		_layout = relativeLayoutPath;

		await PreloaderScripts();		
		await loadScripts();

		clearLocalStorageItemsByField(itemsToRemove);

		fixTextArea();

		_modulename = moduleName;
		_formType = formType;
		if(moduleName == 'Compliance')
			await onComplianceRender(formType);	
		
		await setButtonToolTip('Cancel', cancelMesg);
		
		preloader("remove");
	}
	catch (e) {
		alert(e);
		console.log(e);		
		preloader();
	}
}

var onComplianceRender = async function (formType){	

	debugger;

	CurrentUser = await GetCurrentUser(); 
	
    if(formType == 'Edit'){
		await Compliance_editForm();        
    }    
    else if(formType == 'Display'){ 
		await Compliance_displayForm();     
    }
}

fd.spBeforeSave(function()
{ 
	return fd._vue.$nextTick();
});

var Compliance_editForm = async function(){

	debugger;

	fd.toolbar.buttons[0].style="display: none;";
    fd.toolbar.buttons[1].style="display: none;";
    //fd.toolbar.buttons[1].text = "Cancel";
    
    //fixTextArea();
    
    fd.toolbar.buttons.push({
		icon: 'Cancel',
		class: 'btn-outline-primary',
		text: 'Cancel',
		//style: 'background-color: #3CDBC0; color: white;',
		click: async function() {
			fd.close();
		}			
	});	
	
    _status = fd.field('Status').value;
	var actionCode =  fd.field('Action').value;

    $(fd.field('Status').$parent.$el).hide(); 
    $(fd.field('Action').$parent.$el).hide(); 
    $(fd.field('ReasonofRejection').$parent.$el).hide();  
        
	var isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    const MemberName = await CheckifUserinGroup();

    var AssignedToLoginName, AssignedToEmail;
	await pnp.sp.web.currentUser.get().then(user => {
        AssignedToLoginName = user.Title;		
    });    
	
	var AssignedTo;
	var AssignedToEmail;

	AssignedTo = fd.field('AssignedTo').value;
	AssignedToEmail = AssignedTo[0].email;	
	
	if(AssignedToEmail === ''){		
		let userId = AssignedTo[0].id;
		await pnp.sp.web.siteUsers.getById(userId).get().then(user => {		
			AssignedToEmail = user.Email;
			console.log("Fixed Email:", AssignedToEmail);
		}).catch(error => {
			console.log("Error:", error);
		});
	}
	else
		console.log("User Email:", AssignedToEmail);	

    // let totalTries = 10;
	// let retry = 1;
	
	// while(retry < totalTries){
	// 	AssignedTo = fd.field('AssignedTo').value;	
	// 	if(AssignedTo !== null){
	// 		if(AssignedTo[0].email !== '') {
	// 			AssignedToEmail = AssignedTo[0].email;	
	// 			break;
	// 		}
	// 		else retry++;
	// 	}
	// }
	
	// if (AssignedTo && AssignedTo.length > 0) {
	// 	AssignedToEmail = AssignedTo[0].email;		
	// } 

    if(_status === "Assigned")
    {            
        var isUserAllowed = await CheckUserInAssignedColumn(AssignedToLoginName);
        
        if(!isUserAllowed && !isSiteAdmin)
    	{
    		alert("Apologies, but this task has not been assigned to you. Name: " + AssignedToLoginName + " Assign to: " + _AssignedTo);
    		fd.close();
    	} 
        
        fd.field('Response').required = true;
        fd.field('Attachments').required = true; 
		
		if(actionCode.toLowerCase() === 'rejected'){
			$(fd.field('ReasonofRejection').$parent.$el).show();
			fd.field('ReasonofRejection').disabled = true;
		}
        
    	fd.toolbar.buttons.push({
	        icon: 'Accept',
	        class: 'btn-outline-primary',
	        text: 'Send for Review',
			//style: 'background-color: #3CDBC0; color: white;',
	        click: async function() {
            
                $(fd.field('Status').$parent.$el).show();
                fd.field('Status').value = 'Sent for Review';
                $(fd.field('Status').$parent.$el).hide();

				$(fd.field('Action').$parent.$el).show();						
				fd.field('Action').value = '';
				$(fd.field('Action').$parent.$el).hide(); 

				$(fd.field('ReasonofRejection').$parent.$el).show();						
				fd.field('ReasonofRejection').value = '';
				$(fd.field('ReasonofRejection').$parent.$el).hide();				
								
				var notificationName = 'Compliance_TaskSenttoReview';

				var Office = '',				    
					Control = '',
					Year = '',
					ComplianceType = '',
					Requirement = '',
					RecordforOffice = '',
					Response = '';			

				Control = fd.field('Control').value.LookupValue;
				Office = fd.field('Office').value;
				Year = fd.field('Year').value;
				ComplianceType = fd.field('ComplianceType').value;
				Requirement = fd.field('Requirement').value;
				RecordforOffice = fd.field('RecordforOffice').value;
				Response = fd.field('Response').value;

				var EmailbodyHeader = "Kindly be informed that I have reviewed the requirements below, and all necessary records have been attached for your reference.";

				var Result = {  Control: Control, 
					Office: Office, 
					Year: Year, 
					'Compliance Type': ComplianceType, 
					Requirement: Requirement, 
					'Record for Office': RecordforOffice,
					Response: Response
				};

				var Subject = 'ISO 27001 Compliance Check ' + Office + ' - ' + Year + ' - Sent for Review';
				var encodedSubject = htmlEncode(Subject);
				var Body = GetHTMLBody(Result, EmailbodyHeader, AssignedToLoginName, "Assigned");
				var encodedBody = htmlEncode(Body);

				const itemId = fd.itemId;
				let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;

				if(fd.isValid)
					await _sendEmail(_modulename, encodedSubject + '|' + encodedBody, query, '', notificationName, '', CurrentUser);

                fd.save();            
            }			
		});
    }
    
    else if(_status === "Sent for Review" && (MemberName == 'SecurityMember' || isSiteAdmin))
    { 
        $(fd.field('Action').$parent.$el).show();
        fd.field('Action').required = true;
        fd.field('Response').disabled = true; 
        SetAttachmentToReadOnly();   
         
        fd.field('Action').$on('change', function(value){		 
			ShowHideReasonOfRejection(value);
	    }) 
        ShowHideReasonOfRejection(fd.field('Action').value);
              
    	fd.toolbar.buttons.push({
	        icon: 'Accept',
	        class: 'btn-outline-primary',
	        text: 'Submit',
			//style: 'background-color: #3CDBC0; color: white;',
	        click: async function() { 						

				var Office = '',				    
					Control = '',
					Year = '',
					ComplianceType = '',
					Requirement = '',
					RecordforOffice = '',
					Response = '',
					Action = '',
					ReasonofRejection = '';			

				Control = fd.field('Control').value.LookupValue;
				Office = fd.field('Office').value;
				Year = fd.field('Year').value;
				ComplianceType = fd.field('ComplianceType').value;
				Requirement = fd.field('Requirement').value;
				RecordforOffice = fd.field('RecordforOffice').value;
				Response = fd.field('Response').value;
				Action = fd.field('Action').value;
				ReasonofRejection = fd.field('ReasonofRejection').value;		

				if(Action === 'Approved')
				{
					var controlStructure = Control.split('.');
					var Domain = controlStructure[0];	
					var subDomain = controlStructure[1][0];		
					var folderStructure = Year + '-' + Office + '-' + ComplianceType + '-' + Domain + '-' + (Domain + '.' + subDomain);
					await _CreatFolderStructure('Compliance', 'ComplianceRepository', fd.itemId, folderStructure);

					var notificationName = 'Compliance_TaskApproved';

					var EmailbodyHeader = "Thank you for fulfilling the requirements outlined below. We have reviewed your submission and are pleased to confirm that all necessary records have been approved.";

					var Result = {  Control: Control, 
						Office: Office, 
						Year: Year, 
						'Compliance Type': ComplianceType, 
						Requirement: Requirement, 
						'Record for Office': RecordforOffice,
						Response: Response,
						Action: Action
					};

					var Subject = 'ISO 27001 Compliance Check ' + Office + ' - ' + Year + ' - Approved';

					var encodedSubject = htmlEncode(Subject);
					var Body = GetHTMLBody(Result, EmailbodyHeader, AssignedToLoginName, "Approved");
					var encodedBody = htmlEncode(Body);

					const itemId = fd.itemId;
					let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;

					if(fd.isValid)
						await _sendEmail(_modulename, encodedSubject + '|' + encodedBody, query, AssignedToEmail, notificationName, '', CurrentUser);
				}
				else if(Action === 'Rejected')
				{
					var notificationName = 'Compliance_TaskRejected';

					var EmailbodyHeader = "Thank you for the update regarding the completion of the ISO 27001 audit in Dar " + Office + ". Kindly be informed that the provided records have been rejected. We kindly ask you to refer to the reject reason provided to understand the necessary modifications required.";

					var Result = {  Control: Control, 
						Office: Office, 
						Year: Year, 
						'Compliance Type': ComplianceType, 
						Requirement: Requirement, 
						'Record for Office': RecordforOffice,
						Response: Response,
						Action: Action,
						'Reason of Rejection': ReasonofRejection
					};

					var Subject = 'ISO 27001 Compliance Check ' + Office + ' - ' + Year + ' - Rejected';

					var encodedSubject = htmlEncode(Subject);
					var Body = GetHTMLBody(Result, EmailbodyHeader, AssignedToLoginName, "Rejected");
					var encodedBody = htmlEncode(Body);

					const itemId = fd.itemId;
					let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;

					if(fd.isValid)
						await _sendEmail(_modulename, encodedSubject + '|' + encodedBody, query, AssignedToEmail, notificationName, '', CurrentUser);
				}				
				
                fd.save();            
            }			
		});
    }
    
    else if(_status === "Reviewed")
    { 
        $(fd.field('Action').$parent.$el).show(); 
        $(fd.field('ReasonofRejection').$parent.$el).show();
                
        var ActionVal = fd.field('Action').value;     				
	    
	    if (ActionVal === 'Approved')	
        	$(fd.field('ReasonofRejection').$parent.$el).hide();             
    	
        fd.field('Response').disabled = true;
        fd.field('Action').disabled = true;
        SetAttachmentToReadOnly();
    }
    
    else
    {
        alert("Apologies, but this task has not been assigned to you.");
    	fd.close();
    }
}

var Compliance_displayForm = async function(){

	fd.toolbar.buttons[1].text = "Cancel"; 
    
    //fixTextArea(); 
    
    var Status = fd.field('Status').value; 
    if(Status === 'Reviewed')
        fd.toolbar.buttons[0].style="display: none;";
   
    $(fd.field('ReasonofRejection').$parent.$el).hide();
    $(fd.field('Status').$parent.$el).hide();
    
    var ActionVal = fd.field('Action').value;
    if (ActionVal === 'Approved')	
        $(fd.field('ReasonofRejection').$parent.$el).hide();   
    else if (ActionVal === 'Rejected')	
        $(fd.field('ReasonofRejection').$parent.$el).show();  
}

//#region Utilities
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
}

async function CheckifUserinGroup() {
	let UserGrouArray = 'IT';
	
	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			 await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {
					if(groupsData[i].Title === 'Security Perfomance Metric Members')
					{
					    UserGrouArray = 'SecurityMember';
						break;
					}								      
				}				
			});
	     });
    }
	catch(e){alert(e);}
		
	return UserGrouArray;				
}

async function CheckUserInAssignedColumn(Username){ 
		
	_AssignedTo = fd.field('AssignedTo').value[0].displayName;    

	if (_AssignedTo === Username)
		_isExist = true;
	else
		_isExist = false;

	return _isExist;	
        
    // var CC = fd.field('CC').value;
    
    // if (CC === Username)
	// 	return isExist = true;
}

function ShowHideReasonOfRejection(value){
    
    if (value === 'Rejected') {
		$(fd.field('ReasonofRejection').$parent.$el).show();
		fd.field('ReasonofRejection').required = true;
        
        $(fd.field('Status').$parent.$el).show();
        fd.field('Status').value = 'Assigned';
        $(fd.field('Status').$parent.$el).hide();		
	} 
	else if (value === 'Approved') {	
		$(fd.field('ReasonofRejection').$parent.$el).show();
		fd.field('ReasonofRejection').required = false;
		fd.field('ReasonofRejection').value = "";
		$(fd.field('ReasonofRejection').$parent.$el).hide();
        
        $(fd.field('Status').$parent.$el).show();
        fd.field('Status').value = 'Reviewed';
        $(fd.field('Status').$parent.$el).hide();				
	}	
}

function fixTextArea(){
	$("textarea").each(function(index){
		var height = (this.scrollHeight + 5) + "px";
        $(this).css('height', height);
	});
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

function GetHTMLBody(Result, EmailbodyHeader, AssignedToLoginName, actionType){	

	var Body = "<html>";
	Body += "<head>";
	Body += "<meta charset='UTF-8'>";
	Body += "<meta name='viewport' content='width=device-width, initial-scale=1.0'>";				
	Body += "</head>";
	Body += "<body style='font-family: Verdana, sans-serif; font-size: 12px; line-height: 1.5; color: #333;'>";
	Body += "<div style='max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ccc; border-radius: 5px; background-color: #f9f9f9;'>"; 			
	Body += "<p style='margin: 0 0 10px;'>"+ EmailbodyHeader + "</p>";
	if(actionType === 'Assigned'){
		Body += "<table style='table style='width:300px;border-collapse:collapse;margin-bottom:10px;'>" +
				"<tr>" +
				"<td style='padding:4px;text-align:center;font-size:12px;font-family:Verdana;'><a style='color:#616A76;font-weight: bold;font-size:13px' href='" + _ListFullUrl + "/EditForm.aspx?ID=" + fd.itemId + "'>Reply</a></td>" +
				"<td style='padding:4px;text-align:center;font-size:12px;font-family:Verdana;'><a style='color:#616A76;font-weight: bold;font-size:13px' href ='" + _ListFullUrl + "/To Review Items.aspx'>View All</a></td>" +
				"</tr>" +
				"</table>";
	}
	else if(actionType === 'Approved'){
		Body += "<table style='table style='width:300px;border-collapse:collapse;margin-bottom:10px;'>" +
				"<tr>" +				
				"<td style='padding:4px;text-align:center;font-size:12px;font-family:Verdana;'><a style='color:#616A76;font-weight: bold;font-size:13px' href ='" + _ListFullUrl + "/My Tasks.aspx'>View All</a></td>" +
				"</tr>" +
				"</table>";
	}
	else if(actionType === 'Rejected'){
		Body += "<table style='table style='width:300px;border-collapse:collapse;margin-bottom:10px;'>" +
				"<tr>" +
				"<td style='padding:4px;text-align:center;font-size:12px;font-family:Verdana;'><a style='color:#616A76;font-weight: bold;font-size:13px' href='" + _ListFullUrl + "/EditForm.aspx?ID=" + fd.itemId + "'>Reply</a></td>" +
				"<td style='padding:4px;text-align:center;font-size:12px;font-family:Verdana;'><a style='color:#616A76;font-weight: bold;font-size:13px' href ='" + _ListFullUrl + "/My Tasks.aspx'>View All</a></td>" +
				"</tr>" +
				"</table>";
	}
	Body += "<table style='width: 100%; border-collapse: collapse; margin-bottom: 10px;'>";			

	for (var column in Result) {
		Body += "<tr>";
		Body += "<th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;font-size:13px'>" + column + "</th>";
		Body += "<td style='border: 1px solid #ddd; padding: 8px; text-align: left;font-size:13px'>" + Result[column] + "</td>";
		Body += "</tr>";
	}				

	var Attachs = "";
	
	var FullPathAtt =  _ListFullUrl + "/Attachments/" + fd.itemId + "/";
	for (var i = 0; i < fd.field('Attachments').value.length; i++) {

		var attachmentUrl = FullPathAtt + fd.field('Attachments').value[i].name;
		var attachmentName = fd.field('Attachments').value[i].name;
		
		if (Attachs === "") {
			Attachs = "<a href='" + attachmentUrl + "'>" + attachmentName + "</a>";
		} else {
			Attachs += "<br/><a href='" + attachmentUrl + "'>" + attachmentName + "</a>";
		}
	}

	Body += "<tr>";
	Body += "<th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;font-size:13px'>Attachments</th>";
	Body += "<td style='border: 1px solid #ddd; padding: 8px; text-align: left;font-size:13px'>" + Attachs + "</td>";
	Body += "</tr>";
			
	Body += "</table>";
	Body += "<div style='margin-top: 10px;'>";
	Body += "<p style='margin: 0 0 10px;'>Best regards,</p>";
	Body += "<p style='margin: 0 0 10px;'>" + AssignedToLoginName + "</p>";
	Body += "</div>";				
	Body += "</div>";
	Body += "</body>";
	Body += "</html>";

	return Body;
}

const _CreatFolderStructure = async function(ListName, LibraryName, ID, folderStructure){
	let webUrl = _spPageContextInfo.siteAbsoluteUrl;
	let siteUrl = new URL(webUrl).origin;    
    let serviceUrl = `${siteUrl}/AjaxService/DarPSUtils.asmx?op=UploadAttchment`;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                        <soap:Body>
                            <UploadAttchment xmlns="http://tempuri.org/">
                                <WebURL>${webUrl}</WebURL>
                                <ListName>${ListName}</ListName>                             
                                <LibraryName>${LibraryName}</LibraryName>                                
                                <ID>${ID}</ID>
                                <folderStructure>${folderStructure}</folderStructure>                                
                            </UploadAttchment>
                        </soap:Body>
                    </soap:Envelope>`
    let res = await getSoapResponse('POST', serviceUrl, true, soapContent, 'UploadAttchmentResult');
}
//#endregion

var loadScripts = async function(){
	const libraryUrls = [
		//_layout + '/controls/preloader/jquery.dim-background.min.js',
		_layout + "/plumsail/js/customMessages.js",
		_layout + '/controls/tooltipster/jquery.tooltipster.min.js',
		//_layout + '/plumsail/js/preloader.js',
		_layout + '/plumsail/js/utilities.js',
		_layout + '/plumsail/js/commonUtils.js'
	];
  
	const cacheBusting = '?t=' + new Date().getTime();
	  libraryUrls.map(url => { 
		  $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
		});
		
	const stylesheetUrls = [
		_layout + '/controls/tooltipster/tooltipster.css',
		_layout + '/plumsail/css/CssStyle.css'		
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

const GetCurrentUser = async function(){
	try {
        const user = await pnp.sp.web.currentUser.get();
        return user;
    } catch (error) {
        console.error("Error fetching current user:", error);
        throw error; // Re-throw the error if needed
    }	
}

const _sendEmail = async function(ModuleName, emailName, query, ApprovalTradeCC, notificationName, rootFolder, currUser){
	let webUrl = _spPageContextInfo.siteAbsoluteUrl;
	let siteUrl = new URL(webUrl).origin;
    let CurrentUser;
	
	debugger;
	if(currUser !== undefined && currUser !== null && currUser !== '')
		CurrentUser = currUser
	else CurrentUser = await GetCurrentUser(); 
	

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