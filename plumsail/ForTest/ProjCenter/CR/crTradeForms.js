let isPCR = false, isPMTask = false, department = '';

var onCRTardeFormsRender = async function(){
 
    //getConsoleLogRoles();
    clearLocalStorageItemsByField(['Status']);
 
     formFields = {
         Department: fd.field("Title"),
         Status: fd.field("Status"),

         SubDate: fd.field("SubDate"),
         AppDate: fd.field("AppDate"),
         RejDate: fd.field("RejDate"),
 
         SubBy: fd.field("SubBy"),
         AppBy: fd.field("AppBy"),
         RejBy: fd.field("RejBy"),
 
         RejRemarks: fd.field("RejRemarks"),
         IsPMTask: fd.field("IsPMTask"),
         IsProjectComp: fd.field("IsProjectComp"),

         MasterID: fd.field("MasterID"),
         ProjNo: fd.field("MasterID_x003a_Project_x0020_No_"),
         ProjTitle: fd.field("MasterID_x003a_Project_x0020_Tit"),

         LessonsLearned: fd.field("LessonsLearned"),
         AreaComments: fd.field("AreaComments"),
         SubmittalLoc: fd.field("SubmittalLoc"),
         DwgNo: fd.field("DwgNo"),
         Created: fd.field('Created'),

         Budgetdt: fd.control('budgetdt'),

         PMApprovalStatus: fd.field('PMApprovalStatus'),
         TrFormRejRemarks: fd.field('TrFormRejRemarks')
     }

     isPCR = formFields.IsProjectComp.value;
     department = formFields.Department.value;
     isPMTask = department === 'PM' ? true : false;

     renderHeaderTitle();

     setCRHeaderText();

     let isDisable = manageFormPermission();
    
     //await onTradeContReviewRender(isDisable);
 
     let hideFields = [ formFields.Department, formFields.Status, formFields.SubDate, formFields.AppDate, formFields.RejDate, formFields.SubBy, formFields.AppBy, 
                        formFields.RejBy, formFields.RejRemarks, formFields.MasterID, formFields.IsPMTask, formFields.IsProjectComp,
                        formFields.ProjNo, formFields.ProjTitle, formFields.Created];
    
    if(!isPMTask){
        hideFields.push(formFields.AreaComments, formFields.SubmittalLoc)

        formFields.PMApprovalStatus.$on("change", async function (value){
            if(value === 'Reject'){
                formFields.TrFormRejRemarks.required = true;
                $(formFields.TrFormRejRemarks.$parent.$el).show();
            }
            else{
                formFields.TrFormRejRemarks.required = false;
                $(formFields.TrFormRejRemarks.$parent.$el).hide();
            }
        });
    }
    else {
        formFields.AreaComments.required = true;
        formFields.SubmittalLoc.required = true;
        hideFields.push(formFields.DwgNo)
        formFields.Budgetdt.hidden = true;
    }

    let status = formFields.Status.value 
    if(!_isLLChecker || (_isLLChecker && (status === 'In Progress'  || status=== 'Approve' || status=== 'Reject')))
        hideFields.push(formFields.TrFormRejRemarks)

    if(!_isPM || (_isPM && (status=== 'In Progress'  || status=== 'Approve' || status=== 'Reject')))
        hideFields.push(formFields.PMApprovalStatus, formFields.TrFormRejRemarks)
    

    if(_isPM && status=== 'Submitted'){
       formFields.PMApprovalStatus.required = true;
       formFields.PMApprovalStatus.value = '';
    }

     await _HideFields(hideFields, true);   
}
 
 function setCRHeaderText(){

     let headerText = `<span class="PageTitle">${department} Department Initiation</span>`
     let status = formFields.Status.value;
 
    let color = 'orange', display = '';
    if(status === 'Approve'){
        color = '#49c4b1'
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
          headerText += `<br/><font color="#FF0000\"> Reason: ${formFields.TrFormRejRemarks.value}</font>`
     }
     else headerText += `<br/><br/>`
     $('#crHeader').append(headerText);
 }

 function formatDate(dateValue){
     let date = new Date(dateValue);
 
     // Extract the day, month, and year
     let day = date.getDate();
     let month = date.getMonth() + 1; // Months are zero-based in JavaScript
     let year = date.getFullYear();
 
     if (day < 10) 
       day = '0' + day;
     
     if (month < 10) 
       month = '0' + month;
     
     // Format the date as dd/mm/yyyy
     return day + '/' + month + '/' + year;
 }
 
 let disableCRBtns = async function(disableSubmit, disableSave){
     setTimeout(async () => {
         if(disableSubmit)
           $('span').filter(function(){ return $(this).text() === submitText; }).parent().attr("disabled", "disabled");
         if(disableSave)
           $('span').filter(function(){ return $(this).text() === 'Save'; }).parent().attr("disabled", "disabled");
     }, 200);
 }

 let manageFormPermission = async function(){
    let mId = formFields.MasterID.value;
    _itemId = mId !== undefined && mId !== null ? mId.LookupValue : localStorage.getItem('MasterId')
    if(_itemId === undefined || _itemId === null){
        alert('Master ID is NULL for CR crTradeForms.js')
        fd.close()
    }

    _isPM = JSON.parse(localStorage.getItem('_isPM'));
    _isPD = JSON.parse(localStorage.getItem('_isPD'));
    _isQM = JSON.parse(localStorage.getItem('_isQM'));

    if(!_isPM && !_isPD && !_isQM){
        await getTradeRole(formFields.MasterID.value.LookupValue, department)

        if(isPCR)
         _isLLChecker = await getLLChecker(department)

        _isGLMain = JSON.parse(localStorage.getItem('_isGLMain'));
        _isGL = JSON.parse(localStorage.getItem('_isGL'));
    }

    let isReader = _isQM || _isGL ? true : false;
    let isEditor = _isPD || _isPM || _isGLMain ? true : false;

    let status = formFields.Status.value;
    let disableForm = false;

    if(status === 'In Progress' && isEditor){
       // ALLOW EDIT AND SUBMIT TO LESSONS LEARNED
    }
    else if(status === 'Submitted'){
       
        if(isPCR && _isLLChecker){
            //allow APPROVE / REJECT
        }
        else if(!isPCR && (_isPM || _isPD)){
            //allow APPROVE / REJECT
        }
        else if(isReader || isEditor){
            disableForm = true
            await handleFormFields();
        }
        else{
            alert('Permission Denied')
            fd.close()
        }
    }
    else if(status === 'Reject' && isEditor){
        // ALLOW EDIT AND SUBMIT TO LESSONS LEARNED
    }
    else if(status === 'Pending Lessons Learned Checker' && _isLLChecker){
        // ALLOW EDIT LL FIELD AND SUBMIT TO PM
        debugger;
        let fields = [formFields.DwgNo, 
            fd.field("DesignProblems"),
            fd.field("QualityProblems"),
            fd.field("DeviationJust"),
            fd.field("OtherComments")]

        await _DisableFormFields(fields, true);
    }
    else if(isReader || isEditor || _isLLChecker){
        disableForm = true
        await handleFormFields(); 
    }
    else{
       alert('Permission Denied')
       fd.close()
    }
    return disableForm
 }

 function renderHeaderTitle (){

    let projectNo = formFields.ProjNo.value[0].LookupValue;
    let projectTitle = formFields.ProjTitle.value[0].LookupValue;
    
    let fullProjTitle;
    if(projectNo !== undefined && projectTitle !== undefined)
      fullProjTitle = `${projectNo} - ${projectTitle}`
    else fullProjTitle = localStorage.getItem('FullProjTitle');

    fullProjTitle += isPCR ? ' - Project Completion Form (PCR)' : ' - Phase Completion Form (PhCR)'

    setPageStyle(fullProjTitle);
}

let handleFormFields = async function(){
    let fields = [formFields.DwgNo, 
                  fd.field("DesignProblems"), fd.field("QualityProblems"), fd.field("DeviationJust"), fd.field("OtherComments"),
                  formFields.LessonsLearned.internalName];

    if(isPMTask)
       fields.push(formFields.AreaComments.internalName, formFields.SubmittalLoc.internalName);

    await disableCRBtns(true, true)
    _DisableFormFields(fields, true);
}

let setCRMetaInfo = async function(){
    let status = formFields.Status.value;

    let nextStatus = '';
    let pmApproval = formFields.PMApprovalStatus.value;

    let dateNow = new Date();
    let currentUser = await GetCurrentUser();
    let objUser = currentUser.LoginName;
    
    switch(status){
        case 'In Progress':
        case 'Reject':
          formFields.RejBy.value = objUser;

          formFields.RejDate
          formFields.RejDate.value = dateNow;

          if(isPMTask)
            nextStatus = 'Approve'
          else nextStatus = isPCR ? 'Pending Lessons Learned Checker' : 'Submitted'
          break;

        case 'Submitted':
            nextStatus = pmApproval
            if(nextStatus === 'Reject'){
                //$(formFields.RejBy.$parent.$el).show();
                formFields.RejBy.value = objUser;
                //$(formFields.RejBy.$parent.$el).hide();
                formFields.RejDate.value = dateNow;
            }
            else{
                formFields.AppBy.value = objUser;
                formFields.AppDate.value = dateNow;
                formFields.TrFormRejRemarks.value = '';
            }
            break;

        case 'Pending Lessons Learned Checker':
            nextStatus = 'Submitted'
            formFields.SubBy.value = objUser;
            formFields.SubDate.value = dateNow;
            break;
    }
    if(nextStatus !== '')
      formFields.Status.value = nextStatus

    let resultValue = {
        PCRtradeFormID: fd.itemId,
        PCRtradeFormStatus: nextStatus,
        _module: _module
    }
    debugger;
    if(window.opener) 
      window.opener.postMessage(resultValue, `${_webUrl}/SitePages/PlumsailForms/ProjectInfo/Item/EditForm.aspx?item=${_itemId}`)
}


// fd.spBeforeSave(function(spForm){	
//     debugger;
//     setCRMetaInfo();				
// 	return fd._vue.$nextTick();
// });