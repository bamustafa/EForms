var onAURender = async function() {
    $('div.SPCanvas, .commandBarWrapper').css('padding-left', '0px');
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

            debugger;
              if(disableCode)
               aurFields.push('Code');
              //await disableFields(aurFields, true, false);
              //await disablPeoplePickerFields(true);
            }
        }
        else{
            let masterID = JSON.parse(localStorage.getItem('MasterID'))
            let mpID = JSON.parse(localStorage.getItem('MPID'))
       
            fd.field('MasterID').value = masterID
            fd.field('MPID').value = mpID

            $(fd.field('MasterID').$parent.$el).hide();
            $(fd.field('MPID').$parent.$el).hide();
        }
    }
}

//#region AUR FUNCTIONS
var handleSummaryButton = async function(){

    var department = '';
    if(fd.field('Department').value)
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
    
    let liItems = $('ul.nav-tabs li a')
    if(liItems.length > 0){
        liItems.on('click', function(element) {
                
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
                    setButtonState(_submitText, false);
                    disableSummary();
                }
                else  handleCarTab();
        });
    }
}

var handleAurTabs = async function(){
    var code = fd.field('Code').value;
    var status = fd.field('Status').value;
    var isSubmitted = fd.field('Submit').value;
    var office = fd.field('Office').value.LookupValue;
    var officeItem = await getAssignedUsers(office);

    if(status === 'Closed'){
        //#region SET AUDIT REPORT TAB READ ONLY AND ENABLE CAR TAB FOR AUDIT TEAM
   
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

const disablPeoplePickerFields = async function(isDisable){
    for(let i = 0; i < aurPeopleFields.length; i++){
        var fieldName = aurPeopleFields[i];
        fd.field(fieldName).ready(function(field) {
           field.disabled = isDisable;
        });
    }
}

const disableFields = async function(fields, disableControls, disableCustomButtons){

    debugger;
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
       $('span').filter(function(){ return $(this).text() == _submitText }).parent().attr("disabled", "disabled");
    }
}

