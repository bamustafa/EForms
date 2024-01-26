var ProjectNameList="", ProjectYear=2010;

var _layout, _module = '', _formType = '', _webUrl, _siteUrl;
var hFields = [], projectArr=[];
var delayTime = 100, retryTime = 10, _timeOut, _timeOut1, _pplTimeOut, _distimeOut;

var  _isMain = true, _isLead = false, _isPart = false, _isNew = false, _isEdit = false, _isSentForReview = false, _isSentForApproval = false, _isAurClosed = false, 
     _hideCode = true, disableCode = false, _ignoreBtnCreation = false;
var _header, _main, _footer;

var auditGenSummary = 'Fill Department and start/end Audit date to generate summary';
var carGenSummary = 'Fill all fields above to proceed';

var aurFields = ['Department', 'AuditType', 'StartAuditDate', 'EndAuditDate', 'Title', 'Office', 'AuditRef'];
var aurPeopleFields = ['AuditTeam', 'Contact', 'Distribution', 'PreparedBy'];

var auditTeamGroup = 'Audit Team';
var activeTabName = '';
var auditReportTab = 'Audit Report';
var corrActionTab = 'Corrective Action';

var onRender = async function (relativeLayoutPath, moduleName, formType) {

    await getGlobalParameters(relativeLayoutPath, moduleName, formType); // GET LAYOUT PATH FROM SOAP SERVICE


    if(_module === 'DCC')
     await onDCCRender();
   else if(_module === 'AUR')
     await onAURender();

     await renderControls();
     
     if(_isAurClosed){
        activeTabName = corrActionTab;
         await renderTabs();
         await handleCarTab();
     }

     _timeOut = setInterval(fixEnhancedRichText, 1000);
     _distimeOut = setInterval(adjustDisableOpacity, 1000);
}

//#region DCC RENDER MODULE
var onDCCRender = async function() {
    if (_isNew || _isEdit) {
        hFields = ['Signature1Yes', 'IDC6Yes', 'IDC9YesCH', 'IDC11YesCH','Title'];
        HideFields(hFields, true);

        if(_isNew)
            setProjectsWithValidation();
        else
        {
            $(fd.field('Title').$parent.$el).show();
            fd.field('Title').disabled = true;
            fd.field('PKnumber').disabled = true;
            fd.field('Department1').disabled = true;
            fd.field('DepUnits').disabled = true;
        }           

        setColumnsArrayShowHide();
    } 
}
//#endregion

//#region AUR RENDER MODULE
var onAURender = async function() {
    if (_isNew || _isEdit){

        //#region HANDLE GENERATE AUDIT SUMMARY BUTTON
        fd.field('Department').$on('change', function(value)
    	{
             handleSummaryButton();
    	});	

        fd.field('StartAuditDate').$on('change',  function(value)
    	{
             handleSummaryButton();
    	});

        fd.field('EndAuditDate').$on('change', function(value)
    	{
             handleSummaryButton();
    	});
         handleSummaryButton();
        //#endregion

        if(_isEdit){
            if( activeTabName === corrActionTab)
             await handleCarTab();
            else{
                await handleAurTabs();
                if(!_isAurClosed)
                 fd.container('Tab1').tabs[1].disabled = true;
            }

            hFields = ['SummaryPlainText', 'ReasonOfRejection', 'CARSummaryPlainText', 'Status', 'Submit'];
             if(_hideCode && !disableCode)
              hFields.push('Code');
             else {
                fd.field('Code').required = true;
                fd.field('Code').$on('change', async function(value){
                    var rejFields = 'ReasonOfRejection';
                    if(value === 'Rejected'){
                        $(fd.field(rejFields).$parent.$el).show();
                        fd.field(rejFields).required = true;
                    }
                    else {
                        fd.field(rejFields).required = false;
                        $(fd.field(rejFields).$parent.$el).hide();
                    }
                });
             }

            HideFields(hFields, true);

            var SummaryPlainText = fd.field('SummaryPlainText').value;
            if(SummaryPlainText !== null && SummaryPlainText !== undefined && SummaryPlainText !== ''){
              fd.field('Summary').value = SummaryPlainText;

            var CARSummaryPlainText = fd.field('CARSummaryPlainText').value;
            if(CARSummaryPlainText !== null && CARSummaryPlainText !== undefined && CARSummaryPlainText !== '')
              fd.field('CARSummary').value = CARSummaryPlainText;

              if(disableCode)
               aurFields.push('Code');
              await disableFields(aurFields, true, false);
              await disablPeoplePickerFields(true);
            }
        }
    }
}
//#endregion

//#region DCC FUNCTIONS
var getProjectList = async function(ProjectYear) 
{ 
    var SERVICE_URL = _siteUrl+ "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=getAllProjectsbyear&ProjectYear="+encodeURIComponent(ProjectYear);
    var xhr = new XMLHttpRequest();	

	xhr.open("GET", SERVICE_URL, true);	
	
	xhr.onreadystatechange = async function() 
	{
		if (xhr.readyState == 4) 
		{ 	
			var response;
			try 
			{
				if (xhr.status == 200)
				{                        
						const obj =  await JSON.parse(this.responseText, async function (key, value) {
						var _columnName = key;
						var _value = value;					

						if(_columnName === 'ProjectCode'){
							if(!projectArr.includes(_value))							
						       projectArr.push(_value);						
						}						
					});				

                    // for(var i =0; i < 20000; i++){
                    //     projectArr.push(i);
                    // }
					fd.field('Reference').widget.setDataSource({data: projectArr});                    			
				}				
			}
			catch(err) 
			{
				console.log(err + "\n" + text);				
			}
		}
	}
	xhr.send();
} 

function showHideColumn(ColumnName, value)
{      
    if(value.toLowerCase() !== "no")
    {
        fd.field(ColumnName).required = false;
        $(fd.field(ColumnName).$parent.$el).hide();
    }    
    else
    {
        $(fd.field(ColumnName).$parent.$el).show();
        fd.field(ColumnName).required = true;
    }
}

function showHideCustomColumn(ColumnName, value)
{      
    if(value.toLowerCase() === "yes")
    {
        $(fd.field(ColumnName).$parent.$el).show();
        fd.field(ColumnName).required = true;
    }    
    else
    {
        fd.field(ColumnName).required = false;
        $(fd.field(ColumnName).$parent.$el).hide();
    }
}

function renderColumns(QAcolumn, ANcolumn, value){
    var matchingColumns = ['Signature1',   '_x0049_DC6', '_x0049_DC9', '_x0049_DC11', 'Process1',   'Authoring1',        'Authoring4',      'Coordination1',
                           'Coordination6',      'TRGM1',    'TRGM2',    'TRGM4',      'Submission1',    'Submission2'];

    var matchingAnswers = ['Signature1Yes', 'IDC6Yes',   'IDC9YesCH', 'IDC11YesCH',   'Process1Yes', 'Authoring1YesCH',  'Authoring4YesCH', 'Coordination1YesCH', 
                           'Coordination6YesCH', 'TRGM1Yes', 'TRGM2Yes', 'TRGM4YesCH', 'SubmissionYesCH', 'Submission2YesCH'];

    var isFound = false;
    if(QAcolumn === '_x0049_DC4')
    {
        showHideCustomColumn(ANcolumn, value);  
        isFound = true;
    }         
    else showHideColumn(ANcolumn, value);
    
    if(!isFound)
    {
        if(matchingColumns.includes(QAcolumn)){
            const index = matchingColumns.indexOf(QAcolumn);
            showHideCustomColumn(matchingAnswers[index],  fd.field(QAcolumn).value);
        }        
    }
}

var setProjectsWithValidation = async function(){
    fd.field('Reference').required = true;     
    fd.field('Reference').addValidator({
        name: 'Array Count',
        error: 'Only one Project can be selected per form.',
        validate: function(value) {
            if(fd.field('Reference').value.length > 1) {
                return false;
            }
            return true;
        }
    });
                 	
    fd.field('Reference').widget.dataSource.data(['Please wait while retreiving ...']); 
    $("ul.k-reset")
      .find("li")
      .first()
      .css("pointer-events", "none")
      .css("opacity", "0.6"); 
    await getProjectList(ProjectYear);


    //var id = $('div.k-multiselect').first();
    // $(id).on('click', async function(value) {
    //     var countResult = fd.field('Reference').widget.dataSource._data.length;
    //     if(countResult < 2){
    //         fd.field('Reference').widget.dataSource.data(['Please wait while retreiving ...']); 
    //         $("ul.k-reset")
    //         .find("li")
    //         .first()
    //         .css("pointer-events", "none")
    //         .css("opacity", "0.6");  
   
    //         await getProjectList(ProjectYear);
    //     }
    // });

    fd.field('Reference').$on('change', async function(value)
	{
            $(fd.field('Title').$parent.$el).show();		
            fd.field('Title').value = value;
            $(fd.field('Title').$parent.$el).hide();
	});  
}

var setColumnsArrayShowHide = async function(){
    const QAColumns = ["Registry1", "Registry2", "Regsitry3", "Checking1", "Checking2", "Checking3", "Checking4", "Revision1", "Revision2", "Revision3", 
                       "Revision4", "TitleBlock1", "TitleBlock2", "TitleBlock3",  "Signature1", "Stamp1", "Stamp2", "General1", "General2", "General3", 
                       "DrawingsOTE1", "_x0049_DC1", "_x0049_DC2", "_x0049_DC3", "_x0049_DC4", "_x0049_DC5", "_x0049_DC6", "_x0049_DC7", "_x0049_DC8", 
                       "_x0049_DC9", "_x0049_DC10", "_x0049_DC11", "Process1", "Authoring1", "Authoring2", "Authoring3", "Authoring4", "Authoring5", 
                       "Coordination1", "Coordination2", "Coordination3", "Coordination4", "Coordination5", "Coordination6", "Coordination7", "TRGM1", 
                       "TRGM2", "TRGM3", "TRGM4", "TRGM5", "TRWT1", "TRWT2", "TRWT3", "TRWT4", "TRWT5", "TRWT6", "TRWT7", "Translation1", "Translation2", 
                       "Translation3", "Submission1", "Submission2", "IsDrawing"];

    const ANColumns = ["Registry1CH", "Registry2CH", "Registry3CH", "Checking1CH", "_x201c_Checking2CH_x201d_", "Checking3CH", "Checking4CH", "Revision1CH", 
                       "Revision2CH", "Revision3CH", "Revision4CH", "TitleBlock1CH", "TitleBlock2CH", "TitleBlock3CH", "Signature1CH", "Stamp1CH", "Stamp2CH", 
                       "General1CH", "General2CH", "General3CH", "DrawingsOTE1CH", "IDC1CH", "IDC2CH", "IDC3CH", "IDC4YesCH", "IDC5CH", "IDC6CH", "IDC7CH", 
                       "IDC8YesCH", "IDC9CH", "IDC10CH", "IDC11CH", "Process1CH", "Authoring1CH", "Authoring2CH", "Authoring3CH", "Authoring4CH",
                       "Authoring5CH", "Coordination1CH", "Coordination2CH", "Coordination3CH", "Coordination4CH", "Coordination5CH", "Coordination6CH", 
                       "Coordination7CH", "TRGM1CH", "TRGM2CH", "TRGM3CH", "TRGM4CH", "TRGM5CH", "TRWT1CH", "TRWT2CH", "TRWT3CH", "TRWT4CH", "TRWT5CH", "TRWT6CH", 
                       "TRWT7CH", "Translation1CH", "Translation2CH", "Translation3CH", "Submission1CH", "Submission2CH", "DrawingAvailability"];

    for (var i = 0; i < QAColumns.length; i++) 
    {		
        let QAcolumn = QAColumns[i];
        let ANcolumn = ANColumns[i];
        

    	fd.field(QAcolumn).$on('change', function(value)
    	{
            renderColumns(QAcolumn, ANcolumn, value);
    	});	
    
        var value = fd.field(QAcolumn).value;
        renderColumns(QAcolumn, ANcolumn, value);
    }
}
//#endregion

//#region AUR FUNCTIONS
var handleSummaryButton = async function(){

    var department = '';
    if(fd.field('Department').value !== undefined)
     department = fd.field('Department').value.LookupValue;

    var startDate = fd.field('StartAuditDate').value;
    if(startDate !== null){
        var year = startDate.getFullYear();
        var month = startDate.getMonth() + 1; // January is 0, so we add 1 to get the actual month
        var day = startDate.getDate();
        startDate = `${year}-${month}-${day}`;
    }

    var endDate = fd.field('EndAuditDate').value;
    if(endDate !== null){
        var year = endDate.getFullYear();
        var month = endDate.getMonth() + 1; // January is 0, so we add 1 to get the actual month
        var day = endDate.getDate();
        endDate = `${year}-${month}-${day}`;
    }

    var btnAuditElement = $(`button:contains('Generate Audit Summary')`);
    
    $(btnAuditElement).click(async function(){
        preloader();
        
        var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GET_AUR_SUMMARY'; //https://db-sp.darbeirut.com
        var soapContent;
        soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                        '<soap:Body>' +
                          '<GET_AUR_SUMMARY xmlns="http://tempuri.org/">' +
                            '<WebURL>' + _webUrl + '</WebURL>' +
                            '<Department>' + department + '</Department>' +
                            '<StartDate>' + startDate + '</StartDate>' +
                            '<EndDate>' + endDate + '</EndDate>' +
                          '</GET_AUR_SUMMARY>' +
                        '</soap:Body>' +
                      '</soap:Envelope>';
         await getSoapRequest('POST', serviceUrl, false, soapContent)
         .then(async function(){
            var filterFields = ['Department', 'StartAuditDate', 'EndAuditDate'];
            var _operation = 'Add';
             if(_main.includes("No result")){
                await disableFields(filterFields, false, false);
                _operation = 'remove';
             }
             else await disableFields(filterFields, true, false);

             if(_operation === 'Add'){
                filterFields.forEach(function (element) {
                    if (!aurFields.includes(element)) {
                        aurFields.push(element);
                    }
             });
            }
            else{
                aurFields = aurFields.filter(function(field) {
                    return !filterFields.includes(field);
                });
            }

            await disableFields(aurFields, true, false);
            await disablPeoplePickerFields(true);
         })
         .then(function(){
            setButtonState('Print', true);
            _timeOut1 = setInterval(Remove_Preloader, 1000);
         });
        //preloader("remove");
    });

    $('ul.nav-tabs li a').on('click', function(element) {
             debugger;
             var index = $(this).parent().index()
             renderTabs(index);

            activeTabName = $(this).text();
            if(activeTabName == auditReportTab){
                if(department === null || department === undefined || startDate === null || endDate === null){
                    setErrorMesg(btnAuditElement, false, auditGenSummary);
                    $(btnAuditElement).prop('disabled', true);
                }
                else{
                    setErrorMesg(btnAuditElement, true, '');
                    $(btnAuditElement).prop('disabled', false);
                } 
                var summary = fd.field('Summary').value;
                if(summary === null || summary === undefined || summary === '')
                 setButtonState('Print', false);
                else setButtonState('Print', true);

                var status = fd.field('Status').value;
                if(status === 'Closed')
                 setButtonState('Submit', false);
                disableSummary();
            }
            else  handleCarTab();
    });
}

var handleAurTabs = async function(){
    var code = fd.field('Code').value;
    var status = fd.field('Status').value;
    var isSubmitted = fd.field('Submit').value;
    var office = fd.field('Office').value.LookupValue;
    var officeItem = await getAssignedUsers(office);

    if(status === 'Closed'){
        //#region SET AUDIT REPORT TAB READ ONLY AND ENABLE CAR TAB FOR AUDIT TEAM
        debugger;
        var _isAllowed = await ensureFunction('IsUserInGroup', auditTeamGroup);
        if(!_isAllowed)
          $("ul.nav,nav-tabs").remove();
        else {
            disableSummary();
          _isAurClosed = true;
        }
        //#endregion
    }
    else if(status === 'Sent for Review' && (code === '' || code === 'Rejected') ){
        if(isSubmitted){
          _ignoreBtnCreation = true;
          disableSummary();
        }
 
        var assignedOfficeUser = officeItem.AssignedTo;
        var _isAllowedUser = false;
        for (var i = 0; i < assignedOfficeUser.length; i++) {
            var username = assignedOfficeUser[i].Title;
            _isAllowedUser = await doesUserAllowed(username);
            if(_isAllowedUser)
                break;
        }
        if(_isAllowedUser){
            _hideCode = false;
            _isSentForReview = true;

            if(isSubmitted)
            _ignoreBtnCreation = false;
        }
    }
    else if(status === 'Sent for Approval'){
        disableSummary();
        var FinalApproverUser = officeItem.FinalApprover;
        var _isAllowedUser = false;
        for (var i = 0; i < FinalApproverUser.length; i++) {
            var username = FinalApproverUser[i].Title;
            _isAllowedUser = await doesUserAllowed(username);
            if(_isAllowedUser) break;
        }
        if(_isAllowedUser){
         _hideCode = false;
         _isSentForApproval = true;
        }
        else _ignoreBtnCreation = true
    }
    else {
        _ignoreBtnCreation = true;
        //$(`button:contains('Generate Audit Summary')`).remove();
        disableSummary();
        $("ul.nav,nav-tabs").remove();
    }

    if((status === 'Sent for Review' || status === 'Sent for Approval') && code === 'Rejected')
      disableCode = true;
}

var getAssignedUsers = async function(office){
	let result = null;
      await pnp.sp.web.lists
			.getByTitle("Offices")
			.items
			.select("Title,AssignedTo/Title,FinalApprover/Title")
            .expand("AssignedTo,FinalApprover")
			.filter(`Title eq '${  office  }'`)
			.get()
			.then((items) => {
				if(items.length > 0)
				  result = items[0];
				});
	 return result;
}

var handleCarTab = async function(){
    setButtonState('Submit', true);
    
    var btnCarElement = $(`button:contains('Generate CAR Summary')`);
    $(btnCarElement).click(async function(){
       await setSoapResult();
    });

    var contentData = fd.field('Summary').value;
    if(contentData.startsWith('<div')){
        var startIndex = contentData.indexOf('<table');
        contentData = contentData.substring(startIndex);
    }

    var elements = $.parseHTML(contentData);
    var bodyContent = $(elements).filter('#body1');
    fd.field('DescriptionCar').value = bodyContent[0].innerHTML;
    fd.field('DescriptionCar').disabled = true;
    delay(500);

    var disableBtn = false;
    var _fields = ['CarSerialNo', 'DrawingControl', 'DescriptionCar', 'CausesOfNC', 'CorrectiveActionTaken', 'ScheduledCompletionDate', 'Comments', 'FollowupDate'];
    for(var i = 0; i < _fields.length; i++){
        let _field = _fields[i];
        fd.field(_field).required = true
        var _value = fd.field(_field).value;
        if(_value === undefined || _value === null || _value === '')
            disableBtn = true;

        fd.field(_field).$on('change', async function(value){
                disableBtn = await checkCarFields(_fields);
                if(disableBtn){
                    setErrorMesg(btnCarElement, false, carGenSummary);
                    $(btnCarElement).prop('disabled', true);
                    setButtonState('Print', false);
                }
                else{
                    setErrorMesg(btnCarElement, true, '');
                    $(btnCarElement).prop('disabled', false);
                    setButtonState('Print', true);
                } 
        });
    }

    if(disableBtn){
        setErrorMesg(btnCarElement, false, carGenSummary);
        $(btnCarElement).prop('disabled', true);
        setButtonState('Print', false);
    }
    else{
        setErrorMesg(btnCarElement, true, '');
        $(btnCarElement).prop('disabled', false);
        setButtonState('Print', true);
    } 
    disableSummary();
    _distimeOut = setInterval(adjustDisableOpacity, 1000);
}

var setSoapResult = async function(){
    preloader();
    var CurrentUser = await GetCurrentUser();
    var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GET_EMAIL_BODY';
    var soapContent;
    soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body>' +
                      '<GET_EMAIL_BODY xmlns="http://tempuri.org/">' +
                        '<WebURL>' + _webUrl + '</WebURL>' +
                        '<Email_Name>CAR_FORM</Email_Name>' +
                        '<Id>' + fd.itemId + '</Id>' +
                        '<UserDisplayName>' + CurrentUser + '</UserDisplayName>' +
                      '</GET_EMAIL_BODY>' +
                    '</soap:Body>' +
                  '</soap:Envelope>';
     await getSoapRequest('POST', serviceUrl, false, soapContent)
     .then(async function(){
        if(_module === 'AUR' && activeTabName === corrActionTab){
           var fields = ['CarSerialNo', 'DrawingControl', 'DescriptionCar', 'CausesOfNC', 'CorrectiveActionTaken', 'ScheduledCompletionDate', 'Comments', 'FollowupDate'];
           await disableFields(fields, true, false);
        }
     })
     .then(function(){
        setButtonState('Print', true);
        _timeOut1 = setInterval(Remove_Preloader, 1000);
     });
}

//#endregion

//#region GENERAL
const GetCurrentUser = async function(){
	var LoginName = '';
	await pnp.sp.web.currentUser.get()
		         .then(async (user) =>{
					debugger;
					LoginName = user.Title; //LoginName;
	 });
	 return LoginName;
}

function setButtonState(text, isEnabled){
    if(isEnabled)
     $('span').filter(function() { return $(this).text() == text }).parent().removeAttr('disabled');
    else $('span').filter(function(){ return $(this).text() == text; }).parent().attr("disabled", "disabled");
}

function HideFields(fields, isHide){
	for(let i = 0; i < fields.length; i++)
	{
		if(isHide)
		  $(fd.field(fields[i]).$parent.$el).hide();
		else $(fd.field(fields[i]).$parent.$el).show();
	}
}

var setButtons = async function () {
    var status = fd.field('Status').value;
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

    if(!_ignoreBtnCreation)
    {
        await customButtons("Accept", "Submit", false);

        if(_isEdit && _module === 'AUR'){
           await customButtons("Print", "Print", false);

           var summary = fd.field('Summary').value;

           if(summary === null || summary === undefined || summary === '')
           setButtonState('Print', false);
        }

        if(activeTabName == auditReportTab && status === 'Closed')
            setButtonState('Submit', false);
    }

    await customButtons("ChromeClose", "Close", false);
}

fd.spSaved(function(result) {         
    try
    {   
        if(_isNew && _module === 'AUR' ){
            var itemId = result.Id;                     
            var webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
            result.RedirectUrl = webUrl + "/SitePages/PlumsailForms/AuditReport/Item/EditForm.aspx?item=" + itemId;   
        }                   
    }
    catch(e){alert(e);}                              
 });

 const doesUserAllowed = async function(userName){
	let _isAllowed = false
	await pnp.sp.web.currentUser.get()
		 .then(async (user) =>{
			if(user.Title == userName){
				_isAllowed = true;
			}
		});	
	return _isAllowed;	
}

function setToolTipMessages(){
    setButtonToolTip('Submit', submitMesg);
    setButtonToolTip('Close', cancelMesg);
  }

function delay(time) {
    return new Promise(function(resolve) { 
        setTimeout(resolve, time)
    });
}

var renderControls = async function(){
    var isValid = false;
    var retry = 1;
    while (!isValid)
    {
      try{
        if(retry >= retryTime) break;

         await setFormHeaderTitle();
         await setButtons();
         setToolTipMessages();
         //fixTextArea();
         isValid = true;
      }
        catch{
          retry++;
          await delay(delayTime);
        }
    } 
}

var setErrorMesg = async function(inputElement, isCorrect, mesg){
    var errorId = '#cMesg'; 
    if($(errorId).length === 0){
      var htmlContent = "<div id='" + errorId.replace('#','') + "' class='form-text text-danger small'>" + mesg + "</div>";
      $(inputElement).after(htmlContent);
    }
     
    if(!isCorrect){
        $(errorId).html(mesg).attr('style', 'color: rgba(var(--fd-danger-rgb), var(--fd-text-opacity)) !important');
    }
    else{
       $(errorId).remove();
    }
}

function Remove_Preloader(){
    if($('div.dimbackground-curtain').length > 0){
        $('div.dimbackground-curtain').remove();
    }
	
    if($('#loader').length > 0){
        $('#loader').remove();
    }
    clearInterval(_timeOut1);
}

const disableFields = async function(fields, disableControls, disableCustomButtons){

    var isAllowed = true;
    for(let i = 0; i < fields.length; i++){
        var fieldName = fields[i];

        if(fieldName == "Attachments")
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
        
        else {
            try { 
                fd.field(fieldName).disabled = disableControls; 
            }	
            catch(e){alert(`fields[i] = ${  fieldName  }<br/>${  e}`);}
        }
    }

    if(disableCustomButtons)
    {
       $('span').filter(function(){ return $(this).text() == 'Save'; }).parent().attr("disabled", "disabled");
       $('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().attr("disabled", "disabled");
    }
}

const checkCarFields = async function(_fields){
    var disableBtn = false;
    for(var i = 0; i < _fields.length; i++){
        let _field = _fields[i];
        var _value = fd.field(_field).value;
        if(_value === undefined || _value === null || _value === ''){
            disableBtn = true;
            break;
        }
    }
    return disableBtn;
}

// const disablPeoplePickerFields = async function(fields, disableControls){
//     var userFields = $('div.fd-field-user div');
//     if(userFields.length > 0){
//       userFields.addClass('k-state-disabled')
//       clearInterval(_pplTimeOut);
//     }
// }

const disablPeoplePickerFields = async function(isDisable){
     for(let i = 0; i < aurPeopleFields.length; i++){
         var fieldName = aurPeopleFields[i];
         fd.field(fieldName).ready(function(field) {
            field.disabled = isDisable;
         });
     }
 }

 var adjustDisableOpacity = async function(){
    var element = $('div.fd-editor-overlay');
    var isFound = false;
    if(element.length > 0){
        element.css('opacity', '0.4');
        isFound = true;
        //$('textarea').css('height', '100px');
    }

	if(isFound)
	 clearInterval(_distimeOut);   
}

var disableSummary = async function(){
    if(activeTabName === auditReportTab)
     fd.field('Summary').disabled =  true;
    $(`button:contains('Generate Audit Summary')`).remove();
}

var renderTabs = async function(tabIndex){
    var reqIndex = 1;
    if(tabIndex !== undefined)
         reqIndex = tabIndex;
    $('ul.nav-tabs li a').each(function(index){
        var element = $(this);
       
         
       if(index === reqIndex){
         element.attr('aria-selected', 'true');
         element.addClass('active');
       }
       else{
         element.attr('aria-selected', 'false');
         element.removeClass('active');
       }
    });

    $("div[role='tabpanel']").each(function(index){
        var element = $(this);
        if(index === reqIndex){
            element.addClass('show active');
            element.css('display', '');
            
        }
        else{
          element.removeClass('show active');
        }
     });
}

//#endregion

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
					
                    var value= xml.find("GLOBAL_PARAMResult");
                    if(value.length > 0){
                        text = value.text();
                        _layout = value[0].children[0].textContent;
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

var getGlobalParameters = async function(relativeLayoutPath, moduleName, formType){
    localStorage.clear();  
    
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl; //https://db-sp.darbeirut.com/sites/QM-eAudit
    _siteUrl = new URL(_webUrl).origin;
    _layout = relativeLayoutPath;

    var script = document.createElement("script");
    script.src = _layout + "/plumsail/js/config/configFileRouting.js";
    document.head.appendChild(script);

    if(_module === 'DCC' || _module === 'AUR') 
      activeTabName = $('a.active').text();

   if(_formType === 'New'){
    fd.clear();
    _isNew = true;
   }
   else if(_formType === 'Edit')
    _isEdit = true;

    // var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GLOBAL_PARAM';
    // var soapContent;
    // soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    //                 '<soap:Body>' +
    //                   '<GLOBAL_PARAM xmlns="http://tempuri.org/" />' +
    //                 '</soap:Body>' +
    //               '</soap:Envelope>';
    // await getSoapRequest('POST', serviceUrl, false, soapContent);

}

var ensureFunction = async function(funcName, ...params){
    var isValid = false;
    var retry = 1;
    while (!isValid)
    {
        try{
          if(retry >= retryTime) break;
          if(funcName === 'IsUserInGroup'){
            var allowed = await IsUserInGroup(...params);
            isValid = true;
             return allowed;
          }
        }
        catch{
          retry++;
          await delay(delayTime);
        }
    }
}