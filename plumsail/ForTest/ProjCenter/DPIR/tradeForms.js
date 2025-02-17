var onRevDeptTardeFormsRender = async function(){


   //CurrentUser = await GetCurrentUser()
   _isPD = JSON.parse(localStorage.getItem('_isPD'));
   _isPM = JSON.parse(localStorage.getItem('_isPM'));
   _isQM = JSON.parse(localStorage.getItem('_isQM'));
   _isGLMain = JSON.parse(localStorage.getItem('_isGLMain'));
   _isGL = JSON.parse(localStorage.getItem('_isGL'));

   debugger;
   let isReadOnly = !_isGLMain ? true : false; //_isPD || _isPM || _isQM || _isGL ? true : false;
   
  
    clearLocalStorageItemsByField(['Status']);

    formFields = {
        Department: fd.field("Title"),
        Status: fd.field("Status"),
        AppDate: fd.field("AppDate"),
        RejDate: fd.field("RejDate"),

        AppBy: fd.field("AppBy"),
        RejBy: fd.field("RejBy"),

        RejRemarks: fd.field("RejRemarks"),
        MasterID: fd.field("MasterID"),
        ListOfDocs: fd.field("ListOfDocs"),

        GLMain: fd.field("GLMain"),
        GL: fd.field("GL"),
        TeamMembers: fd.field("TeamMembers")
    }

    await manageFields();

    let mId = formFields.MasterID.value;
    _itemId = mId !== undefined && mId !== null ? mId.LookupValue : localStorage.getItem('MasterId')
    if(_itemId === undefined || _itemId === null){
        alert('Master ID is NULL for DPIR tradeForms.js')
        fd.close();
    }

    debugger;
    let isDisable = false
    if(isReadOnly || (formFields.Status.value !== 'In Progress' && formFields.Status.value !== 'Reject')){
        isDisable = true
        let fields = ['GenDesc', 'DesignCrit', 'ListOfItems', 'Lesslearn']

        if(isReadOnly){
            fields = [...fields, 'TotalDesDwg', 'ListOfDocs', 'Attachments'];
            await disableBtns(true, true)
        }
        else await disableBtns(true, false)
        
        _DisableFormFields(fields, isDisable);
    }

    await onTradeContReviewRender(isDisable);

    let hideFields = [formFields.Department, formFields.Status, formFields.AppDate, formFields.RejDate, formFields.AppBy, formFields.RejBy, formFields.RejRemarks, 
                      formFields.MasterID, formFields.GLMain, formFields.GL, formFields.TeamMembers];

    await _HideFields(hideFields, true);

    
}

let  manageFields = async function(){

    $('span').filter(function(){ return $(this).text() == 'Attachments'; }).addClass('fd-required')
    setFieldsTitle();
    setHeaderText();
}

function setFieldsTitle(){
    let label = $('span').filter(function(){ return $(this).text() == '<LessLearn>'; })
    let text = `Lessons learned from previous projects 
                  <a class="FieldsetLabel" href="https://bi.dar.com/ReportServer/Pages/ReportViewer.aspx?/wapps/PMISReports/PCRLessonsLearnedPublic" 
                      target="_blank" style="color:Blue;">Click here to browse the database</a>`
    label.html(text);
    //let listDocValue = $('#ListOfDocs').text()

    let checked =  ''
    if(formFields.ListOfDocs.value === 'None'){
        checked = 'checked'
        formFields.ListOfDocs.disabled = true;
    }

    let ListOfDocs = $('span').filter(function(){ return $(this).text() == '<ListOfDocs>'; })
    text = `Tentative list of documents (spc, boq, etc...) 
                <input type="checkbox" id="tlistId" ${checked} onclick="handleChange(this)" />
                    <label>None</label>`
    
    ListOfDocs.html(text);
}

function handleChange(element){
    let val = '', disable = false;
    if(element.checked){
        val = 'None';
        disable = true;
    }

    formFields.ListOfDocs.value = val;
    formFields.ListOfDocs.disabled = disable;
}

function setHeaderText(){
    let headerText = `<span class="PageTitle">${formFields.Department.value} Department Initiation</span>`
    let status = formFields.Status.value;

   let color = 'orange', display = '';
   if(status === 'Approve'){
       color = 'green'
       display = 'Approved'
   }
   else if(status === 'Reject'){
       color = 'red'
       display = 'Rejected'
   }

   if(status !== 'Approve' && status !== 'Reject')
        display = status;

   headerText += `<br/><span style="color:${color};font-family:verdana;font-size:10pt;font-weight:bold;">${display}</span>`
   
   if(status === 'Approve' || status === 'Reject'){
       let name = formFields.AppBy.value, 
       date = formFields.AppDate.value;
       
       if(status === 'Reject'){
           name = formFields.RejBy.value
           date = formFields.RejDate.value
       }
       name = name !== undefined && name !== null ? name.DisplayText : '';
       date = date !== undefined && date !== null ? formatDate(date) : '';

       headerText += `<br/><span style="font-family:verdana;font-size:8pt;font-weight:normal;">${display} by ${name} on ${date}</span>`
       if(status === 'Reject')
         headerText += `<br/><font color="#FF0000\"> Reason: ${formFields.RejRemarks.value}</font>`
    }
    headerText += setDynamicLink();
    $('#dpirHeader').append(headerText);
}

function setDynamicLink(){
//    let imgUrl = `${_webUrl}${_layout}/Images/MS_D365_logo.png`
//    return `<br/><a href="https://ax.d365.dar.com/namespaces/AXSF/?mi=ProjProjectsListPage" 
//                 style="display:inline-block;color:Blue;font-size:9pt;font-weight:bold;width:147px;margin-top:10px;margin-bottom:10px">Manage project on D365</a>
//               <img src=${imgUrl} alt="Microsoft Dynamics D365" align="bottom" style="height:50px;width:125px;border-width:0px;vertical-align: middle;">`
    return ''
}

function formatDate(dateValue){
    let date = new Date(dateValue);

    // Extract the day, month, and year
    let day = date.getDate();
    let month = date.getMonth() + 1; // Months are zero-based in JavaScript
    let year = date.getFullYear();

    // Add leading zero if day or month is less than 10
    if (day < 10) {
    day = '0' + day;
    }
    if (month < 10) {
    month = '0' + month;
    }

    // Format the date as dd/mm/yyyy
    return day + '/' + month + '/' + year;
}

let disableBtns = async function(disableSubmit, disableSave){
    setTimeout(async () => {
        if(disableSubmit)
          $('span').filter(function(){ return $(this).text() === submitText; }).parent().attr("disabled", "disabled");
        if(disableSave)
          $('span').filter(function(){ return $(this).text() === 'Save'; }).parent().attr("disabled", "disabled");
    }, 200);
}


