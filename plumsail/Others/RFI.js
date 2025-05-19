var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDisplay = false;

var _currentUser, _formFields = {}, _emailFields = {};
const greenColor = '#5FC9B3', redColor = '#F28B82', yellowColor = '#6c757d';

const itemsToRemove = ['WorkflowStatus', 'Discipline', 'Reference'];
let _nextStatus = '', _schema = '', _emailBody = '';
let _status = 'Sent to PM Delegate';

var onRender = async function (relativeLayoutPath, moduleName, formType) {

    try{

        _layout = relativeLayoutPath;

        _formFields = {
            Title: fd.field('Title'),
            Reference: fd.field('Reference'),
            Discipline: fd.field('Discipline'),
            WorkflowStatus: fd.field('WorkflowStatus'),
            Question: fd.field('Description'),
            Answer: fd.field('Answer'),

            PMDelegate: fd.field('PMDelegate'),
            PMDelegateDate: fd.field('PMDelegateDate'),          
            CB: fd.field('CB'),
            CBDate: fd.field('CMDate'), 
            PM: fd.field('OM'),
            PMDate: fd.field('OMDate'),
            Client: fd.field('Client'),
            ClientDate: fd.field('ClientDate'),
            ReviewedDate: fd.field('ReviewedDate'),

            Attachments: fd.field('Attachments')           
        } 
        
        _emailFields = {
            Title: fd.field('Title'),
            Reference: fd.field('Reference'),
            Discipline: fd.field('Discipline'),
            WorkflowStatus: fd.field('WorkflowStatus'),
            Question: fd.field('Description'),
            Answer: fd.field('Answer')           
        }

        const Generaliconsvg = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
        setIconSource("General-icon", Generaliconsvg);   
        const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;
        setIconSource("Attachments-icon", svgattachment);        
        
        await loadScripts().then(async () => {

            showPreloader();
            await extractValues(moduleName, formType);

            if (_isEdit) 
                await handleEditForm();               
            
            else if (_isNew)
                await handleNewForm();          

            else if(_isDisplay)
                await handleDisplayForm();
        })

    }
    catch (e){
      showPreloader();
      fd.toolbar.buttons[0].style = "display: none;";
      fd.toolbar.buttons[1].style = "display: none;";
      console.log(e)
    }
    finally{
        hidePreloader();
    }
}

var handleNewForm = async function () {

    clearLocalStorageItemsByField(itemsToRemove);
    
    await setCustomButtons();
    formatingButtonsBar("RFI Form"); 

    setPSHeaderMessage('');
    setPSErrorMesg(`Please complete all required fields in the form.`); 
    hideHideSection(true);

    _HideFormFields([_formFields.Answer], true);
    _setFieldsDisabled([_formFields.WorkflowStatus, _formFields.Reference, _formFields.Discipline], true);    

    let _Discipline = await CheckifUserinSPGroup();
    _formFields.Discipline.value = _Discipline; 

    _schema = await getAutReferenceFormat();    
    _formFields.Reference.value = await parseRefFormat(_schema);

    _nextStatus = 'Sent to PM Delegate';   
}

var handleEditForm = async function () {
    
    await setCustomButtons();
    formatingButtonsBar("RFI Form"); 
    setPSHeaderMessage('');
    hideHideSection(true);

    let GroupName = await CheckifUserinGroup();
    _status = _formFields.WorkflowStatus.value;  

    const transitions = {
        'Sent to PM Delegate': {
            roles: ['RFIChecker', 'Admin'],
            nextStatus: 'Sent to CB',
            hideAnswer: true,
            disableFields: [_formFields.WorkflowStatus, _formFields.Reference, _formFields.Discipline]
        },
        'Sent to CB': {
            roles: ['RFIReviewer', 'Admin'],
            nextStatus: 'Sent to PM',
            hideAnswer: true,
            disableFields: [_formFields.WorkflowStatus, _formFields.Reference, _formFields.Discipline]
        },
        'Sent to PM': {
            roles: ['RFIApprover', 'Admin'],
            nextStatus: 'Sent to Client',
            hideAnswer: true,
            disableFields: [_formFields.WorkflowStatus, _formFields.Reference, _formFields.Discipline]
        },
        'Sent to Client': {
            roles: ['Client', 'Admin'],
            nextStatus: 'Reviewed',
            hideAnswer: false,
            disableFields: [_formFields.WorkflowStatus, _formFields.Reference, _formFields.Discipline, _formFields.Title],
            disableQuestion: true,
            isAnsweredRequired: true
        }
    };

    const transition = transitions[_status];

    if (transition && transition.roles.includes(GroupName)) {
        setPSErrorMesg(`Kindly review the current form. The status is marked as '${_status}'.`);
        _HideFormFields([_formFields.Answer], transition.hideAnswer);
        _setFieldsDisabled(transition.disableFields, true);
        
        if (transition.disableQuestion) {
            disableRichTextFieldColumn(_formFields.Question);
        }

        if (transition.isAnsweredRequired) {
            _formFields.Answer.required = true;
        }

        _nextStatus = transition.nextStatus;
    } else {
        setPSErrorMesg(`Please note that the information below pertains to the RFI record.`);
        _HideFormFields([_formFields.Answer], false);
        _setFieldsDisabled([_formFields.WorkflowStatus, _formFields.Reference, _formFields.Discipline, _formFields.Title], true);
        disableRichTextFieldColumn(_formFields.Question);
        disableRichTextFieldColumn(_formFields.Answer);
        $('span').filter(function () {
            return $(this).text() === submitDefault;
        }).parent().attr("disabled", "disabled");

        SetAttachmentToReadOnly();
    }
}

var handleDisplayForm = async function () {
    await setCustomButtons();
    formatingButtonsBar("RFI Form");
    setPSHeaderMessage('');
    setPSErrorMesg(`Please note that the information below pertains to the RFI record. Click 'Edit' to make changes.`);
    hideHideSection(true);
}

//#region General
var loadScripts = async function(withSign){

    const libraryUrls = [
      _layout + '/controls/preloader/jquery.dim-background.min.js',
      _layout + "/plumsail/js/buttons.js",
      _layout + '/plumsail/js/customMessages.js',
      _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
      _layout + '/plumsail/js/preloader.js',
      _layout + '/plumsail/js/commonUtils.js',
      _layout + '/plumsail/js/utilities.js'
    ];

    const cacheBusting = `?t=${Date.now()}`;
      libraryUrls.map(url => {
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`);
        });

    const stylesheetUrls = [
      _layout + '/controls/tooltipster/tooltipster.css',
      _layout + '/plumsail/css/DarTemplate.css' + `?t=${Date.now()}`
    ];

    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}${cacheBusting}">`);
    });
}

var extractValues = async function(moduleName, formType){

    if($('.text-muted').length > 0)
      $('.text-muted').remove();

      _web = pnp.sp.web;
      _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
      _module = moduleName;
      _formType = formType;
      _webUrl = _spPageContextInfo.siteAbsoluteUrl;
      _siteUrl = new URL(_webUrl).origin;

    if(_formType === 'New')
        _isNew = true;

    else if(_formType === 'Edit'){
        _isEdit = true;
        _itemId = fd.itemId;
    }
    else if(_formType === 'Display')
        _isDisplay = true;

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;

    _currentUser = await pnp.sp.web.currentUser.get();    
}

var getAutReferenceFormat = async function(){
    let result;
    const listTitle = _MajorTypes; //list.Title;

    const camlFilter = `<View>
                            <Query>
                                <Where>
                                    <Eq><FieldRef Name='Title'/><Value Type='Text'>RFI</Value></Eq>
                                </Where>
                            </Query>
                        </View>`;
    let items = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });   
    if(items.length == 1){
        result = items[0].CDSFormat;
    } 
    return result;
}

var parseRefFormat = async function (format, updateCounter) {
    const splitRefFormat = format.split('-');
    let returnValue = "";

    for (let i = 0; i < splitRefFormat.length; i++) {
        let column = splitRefFormat[i].replace('[', '').replace(']', '');

        if (column.includes('"')) {
            column = column.replace(/"/g, '');
        } else if (column.includes('$')) {
            column = column.replace(/\$/g, '');
        } else {
            const itemValue = _formFields[column].value;
            if (itemValue !== undefined && itemValue !== "") {
                column = itemValue.toString();
                if (column.includes('#')) {
                    const splitVal = column.split('#');
                    column = splitVal[1];
                }
            }
        }

        returnValue += column + "-";
    }    

    returnValue = returnValue.slice(0, -1); // Remove the trailing dash 
    const digit = returnValue.substring(returnValue.lastIndexOf('-') + 1);
    const lastSlash = returnValue.lastIndexOf('-'); 
    returnValue = (lastSlash > -1) ? returnValue.substring(0, lastSlash) : returnValue;    
    const counter = await GetReferenceCounter(returnValue, updateCounter); // Assumes this returns a number    
    returnValue = returnValue + '-' + counter.toString().padStart(digit.length, '0');

    return returnValue;
}

var GetReferenceCounter = async function (returnValue, updateCounter) {    
   
    let listname = 'Counter';

    try {

        const camlFilter = `<View>
                                <Query>
                                    <Where>
                                        <Eq><FieldRef Name='Title'/><Value Type='Text'>${returnValue}</Value></Eq>
                                    </Where>
                                </Query>
                            </View>`;
        let items = await pnp.sp.web.lists.getByTitle(listname).getItemsByCAMLQuery({ ViewXml: camlFilter }); 
        
        if (updateCounter) {

            var _cols = {};
            
            if (items.length == 0) {
                
                _cols["Title"] = returnValue;
                let value = '2';
                _cols["Counter"] = value;                 
                await pnp.sp.web.lists.getByTitle(listname).items.add(_cols); 
                return '1';
            }
            else if (items.length == 1) {   
                
                var _item = items[0];  
                let value = parseInt(_item.Counter) + 1;
                _cols["Counter"] = `${value}`;
                await pnp.sp.web.lists.getByTitle(listname).items.getById(_item.Id).update(_cols);
                return (value - 1).toString(); //value.toString();
            }
        }
        else {            
            return items.length === 0 ? 1 : items[0].Counter;
        }

    } catch (error) {
        console.error('Error fetching/updating reference counter:', error);
        throw new Error('Error fetching/updating reference counter');
    }
}

async function CheckifUserinSPGroup() {
    var IsTMUser = "null";
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

function generateEmailBody(bodyMessage) {
    let BodyEmail = `<html>
                        <head>
                            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
                            <style>
                                body {
                                    font-family: Arial, sans-serif;
                                    background-color: #f4f4f4;
                                    margin: 0;
                                    padding: 0;
                                    font-size: 15px;
                                    color: #333333;
                                }
                                .container {
                                    width: 100%;                         
                                    margin: 0 auto;
                                    padding: 10px;
                                    background-color: #fdfdfd;
                                }
                                .header {
                                    background-color: #3bdbc0;
                                    color: #ffffff;
                                    padding: 12px 20px;
                                    font-size: 16px;
                                    font-weight: bold;
                                    text-align: center;                            
                                }
                                .content {
                                    padding: 10px;
                                    font-size: 14px;
                                    line-height: 1.6;
                                    color: #555555;
                                }
                                .content p {
                                    margin: 5px 0;
                                }
                                .footer {
                                    text-align: center;
                                    color: #888888;
                                    font-size: 12px;
                                    padding: 10px 0;
                                }
                                .action-button {
                                    display: inline-block;
                                    background-color: #4CAF50;
                                    color: #ffffff;
                                    padding: 10px 15px;
                                    border-radius: 4px;
                                    text-decoration: none;
                                    font-weight: bold;
                                    text-align: center;
                                    margin-top: 15px;
                                }
                                .action-button:hover {
                                    background-color: #45a049;
                                }
                            </style>
                        </head>
                        <body>
                            <table class='container' cellpadding='0' cellspacing='0'>                                                                     
                                <tr>
                                    <td class='content'>                                                             
                                        <p>${bodyMessage}</p>
                                        <br /> 
                                        <p>Best regards,</p>
                                        <br />                                     
                                        <img src='data:image/png;base64,base64Image' alt='Purchase Order Image'/>
                                    </td>
                                </tr> 
                                <tr>
                                    <td class='footer'>
                                        <p>This is an automated message. Please do not reply to this email.</p>
                                    </td>
                                </tr>                                       
                            </table>
                        </body>
                    </html>`;

    return BodyEmail;
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

async function getCurrentUserGroups() {
    try {
        const user = await pnp.sp.web.currentUser.get();
        const groups = await pnp.sp.web.siteUsers.getById(user.Id).groups.get();
        console.log("User is in the following groups:", groups);
    } catch (error) {
        console.error("Error getting user groups:", error);
    }
}

function hideHideSection(isHide) {
  document.querySelectorAll('.hideSection').forEach(el => {
    el.style.display = isHide ? 'none' : '';
  });
}

const _setFieldsDisabled = (fields, isDisabled) => {
  (Array.isArray(fields) ? fields : [fields]).forEach(field => {
    if (field) field.disabled = isDisabled;
  });
};

async function CheckifUserinGroup() {
	var IsTMUser = "Trade";
	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(async function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {				
					if(groupsData[i].Title === "RFIChecker")
					{					
					   IsTMUser = "RFIChecker";
					   break;
				    }
					else if(groupsData[i].Title === "RFIReviewer")
					{					
					   IsTMUser = "RFIReviewer";
					   break;
                    }
                    else if(groupsData[i].Title === "RFIApprover")
					{					
					   IsTMUser = "RFIApprover";
					   break;
                    }
                    else if(groupsData[i].Title === "Client")
					{					
					   IsTMUser = "Client";
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

function DisableAttachment() {
	fd.field('Attachments').disabled = false;
	
	$('div.k-upload-button').remove();
	$('button.k-upload-action').remove();
	$('.k-dropzone').remove();
}
//#endregion

//#region Custom Buttons
var setCustomButtons = async function () {

    if (!_isDisplay) {
        fd.toolbar.buttons[0].style = "display: none;";
        await setButtonActions("Accept", submitDefault, `${greenColor}`);
    }
    fd.toolbar.buttons[1].style = "display: none;";
    
    await setButtonActions("ChromeClose", "Cancel", `${yellowColor}`);

    setToolTipMessages();
}

const setButtonActions = async function(icon, text, bgColor){

    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          style: `background-color: ${bgColor}; color: white;`, //color of font button//
          click: async function() {

            if(text == "Close" || text == "Cancel"){
              showPreloader();
              fd.close();
            }
            else if (text == 'Submit') {

                if (fd.isValid) {

                    showPreloader(); 
                    
                    _emailBody = Object.entries(_emailFields).map(([key, val]) => {
                        let displayValue = val.value || ''; // Default to an empty string if value is undefined
                        if (key === 'WorkflowStatus') 
                            displayValue = _nextStatus; 
                        if (displayValue) {
                            return `
                                <tr>
                                    <td style="padding: 5px 5px; text-align: left; font-size: 14px; font-weight: bold; background-color: #f4f7fb;">${key}</td>
                                    <td style="padding: 5px 5px; text-align: left; font-size: 14px; background-color: #f4f7fb;">${displayValue}</td>
                                </tr>
                            `;
                        }
                    }).join('');

                    if (_nextStatus === 'Sent to PM Delegate') {
                        _formFields.Reference.value = await parseRefFormat(_schema, true);
                        _formFields.WorkflowStatus.value = _nextStatus;
                        _formFields.PMDelegateDate.value = new Date();
                        fd.save();                    
                    }                    
                    else if (_nextStatus === 'Sent to CB') {
                        _formFields.WorkflowStatus.value = _nextStatus;
                        _formFields.PMDelegate.value = _currentUser.Title;
                        _formFields.CBDate.value = new Date();
                        fd.save();                        
                    }
                    else if (_nextStatus === 'Sent to PM') {
                        _formFields.WorkflowStatus.value = _nextStatus; 
                        _formFields.CB.value = _currentUser.Title;
                        _formFields.PMDate.value = new Date();
                        fd.save();                        
                    }
                    else if (_nextStatus === 'Sent to Client') {
                        _formFields.WorkflowStatus.value = _nextStatus;
                        _formFields.PM.value = _currentUser.Title;
                        _formFields.ClientDate.value = new Date();
                        fd.save();                        
                    }
                     else if (_nextStatus === 'Reviewed') {
                        _formFields.WorkflowStatus.value = _nextStatus;
                        _formFields.Client.value = _currentUser.Title;
                        _formFields.ReviewedDate.value = new Date();
                        fd.save();                        
                    }
                }
            }
        }
    });
}

function setToolTipMessages(){

  let finalizetMesg = `Click ${submitDefault} for Manager Approval`;

  setButtonCustomToolTip(submitDefault, finalizetMesg);
  setButtonCustomToolTip('Close', closeMesg);

	if($('p').find('small').length > 0)
    $('p').find('small').remove();
}

function formatingButtonsBar(titelValue){

    $('div.ms-compositeHeader').remove();
    $('i.ms-Icon--PDF').remove();

    let toolbarElements = document.querySelectorAll('.fd-toolbar-primary-commands');
    toolbarElements.forEach(function(toolbar) {
        toolbar.style.display = "flex";
        toolbar.style.justifyContent = "flex-end";
    });

    let commandBarElement = document.querySelectorAll('[aria-label="Command Bar."]');
    commandBarElement.forEach(function(element) {
        element.style.paddingTop = "16px";
    }); 

    document.querySelector('.col-sm-12').style.setProperty('padding-top', '0px', 'important'); 
    $('.col-sm-12').attr("style", "display: block !important;justify-content:end;");   
    $('.fd-grid.container-fluid').attr("style", "margin-top: -15px !important; padding: 10px;");
    $('.fd-form-container.container-fluid').attr("style", "margin-top: -22px !important;");   

    const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';
    const linkElement = `<a href="${_spPageContextInfo.webAbsoluteUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">
                          <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">${titelValue}</a>`;
    $('span.o365cs-nav-brandingText').html(linkElement);

    $('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');

    $('.border-title').each(function() {
        $(this).css({          
            'margin-top': '-35px', /* Adjust the position to sit on the border */
            'margin-left': '20px', /* Align with the content */            
        });
    });
}

function setIconSource(elementId, iconFileName) {

  const iconElement = document.getElementById(elementId);

  if (iconElement) {
      iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
  }
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

function disableRichTextFieldColumn(field){

    let elem = $(field.$el);//$(fd.field(fieldname).$el).find('.k-editor tr');

	elem.each(function(index, element){	

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
	})
}
//#endregion

fd.spSaved(async function(result) {

    try {
        
        _itemId = result.Id; 
        let listPath = `${fd.listUrl}`;
        let listInternalName = listPath.replace("/Lists/", "").trim();
        let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${_itemId}</Value></Eq></Where>`;       

        let linkURL = _spPageContextInfo.webAbsoluteUrl + `/SitePages/PlumsailForms/${listInternalName}/Item/DisplayForm.aspx?item=` + _itemId;
        let Subject = `New Request for Information (RFI) Submitted - Status: ${_nextStatus}`;

        let bodyMessage = `<p>Dear All,</p>
                            <br /> 
                            <p>Kindly be informed that a new <strong>Request for Information (RFI)</strong> has been submitted. Please find the detailed information below.</p>
                            <br />
                            <p>You may click <a href="${linkURL}" class="link">here</a> to take the necessary action.</p> 
                            <br />                                                 
                            <table style="width: 100%; max-width: 760px; margin-top: 20px; border-collapse: collapse;" border="0">                              
                                <tr>
                                    <td style="padding: 5px 5px; text-align: left; font-size: 14px; font-weight: bold; background-color: #007BFF; color: white;">Field</td>
                                    <td style="padding: 5px 5px; text-align: left; font-size: 14px; font-weight: bold; background-color: #007BFF; color: white;">Value</td>
                                </tr>                        
                                ${_emailBody}                             
                            </table><br />`;           
        
        let BodyEmail = generateEmailBody(bodyMessage);
        let encodedSubject = htmlEncode(Subject);
        let encodedBodyEmail = htmlEncode(BodyEmail);  
        
        debugger;
      
        await _sendEmail('ClientRFI', encodedSubject + '|' + encodedBodyEmail, query, '', `ClientRFI_${_nextStatus}`, '', _currentUser);
               
    } catch (e) {
      console.log(e);
  }
});