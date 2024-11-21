let gisPath = `${_spPageContextInfo.siteAbsoluteUrl}${_layout}/plumsail/ProjCenter/features/drawProjects2k14.html`;
// 'https://db-gispub.darbeirut.com/gwadar/default_clean.aspx?projid=AD23005-0100D'
var gisItem = {}, isFound = false, gisId;
let projNo, mapTimeOut;

var onGISRender = async function(){

    clearFormFields(['Option', 'Remarks', 'ApprovalStatus', 'ResRejRemarks'])

    formFields = {
        Status: fd.field("Status"),
        Option: fd.field("Option"),
        longitude: fd.field("longitude"),
        latitude: fd.field("latitude"),
        Remarks: fd.field("Remarks"),
        Attachments: fd.field('Attachments'),
        SelectionType: fd.field('SelectionType'),
        ApprovalStatus: fd.field('ApprovalStatus'),
        RejRemarks: fd.field('ResRejRemarks')
    }

    formFields.Status.disabled = true;
    formFields.longitude.disabled = true;
    formFields.latitude.disabled = true;

    debugger
    let fullProjTitle = localStorage.getItem('FullProjTitle');
    fullProjTitle += ' - GIS Location'
    setPageStyle(fullProjTitle);
    let gisControlId = '#GISId';
    let gisFields = [formFields.longitude, formFields.latitude];

    let itemId, filtQuery = '';
    
    
    if(!isNaN(fd.itemId)){
        itemId = fd.itemId
        filtQuery = `Id eq ${itemId}`

        
        let option = fd.field('Option').selectedOption;
        if( option === 'Attaching A File' && fd.field('Status').value !== 'Rejected'){
            //fd.field('Attachments').disabled = true
            $('div.k-upload-button').remove();
            $('button.k-upload-action').remove();
            $('.k-dropzone').remove();
        }
    }
    else{
        itemId = parseInt(localStorage.getItem('MasterId'));
        filtQuery = `MasterID/Id eq ${itemId}`

        formFields.Status.value = 'In Progress';
    }
    
    let result = await isRecordExist(filtQuery);
    _isGIS = await IsUserInGroup('GIS');

    if(result !== ''){
        isFound = true;
        //window.location.href = `${_webUrl}/SitePages/PlumsailForms/GISLocation/Item/EditForm.aspx?item=${itemId}`
        //IS APPROVED OR CANCELLED SET FUNCTION TO CHECK THE LOGIC
        await handleItemAfterApproval(result);
    }
    
    localStorage.setItem('isGISFound', isFound);
    
    if(isFound){

       formFields.Status.value = result.Status;

        gisId = result.Id
        let option = result.Option
        formFields.Option.value = option
        formFields.longitude.value = result.longitude
        formFields.latitude.value = result.latitude
        formFields.Remarks.value = result.Remarks
        formFields.RejRemarks.value = result.ResRejRemarks

        await checkGISApprovalFields();

        if(option === 'Attaching A File'){
            await _HideFields(gisFields, true);
            $(gisControlId).hide();
        }
        
        else {
            $(formFields.Attachments.$parent.$el).hide();
            await _HideFields(gisFields, false);

            localStorage.setItem('longitude', result.longitude);
            localStorage.setItem('latitude', result.latitude);
            localStorage.setItem('SelectionType', result.SelectionType);
            
            addMap(gisControlId);
        }
    }
    else{
        gisFields.push(formFields.Attachments);
        hideGISApprovalFields()
        await _HideFields(gisFields, true);

        localStorage.removeItem('gisStatus');  

        if(_isGIS)
         disableSubmission('Only PM is allowed to submit for GIS',true);
    }
    
    fd.field('Option').$on('change', async (value) => {
         if(value === 'Attaching A File'){
            await _HideFields(gisFields, true);
            $(formFields.Attachments.$parent.$el).show();
            $(gisControlId).hide();
         }
         else{
            debugger;
            await _HideFields(gisFields, false);
            hideGISApprovalFields()
            $(formFields.Attachments.$parent.$el).hide();
            addMap(gisControlId)

            if(!isFound){
                mapTimeOut = setInterval(async () => {
                                    await disableGISPMFields(true);
                                }, 200);
            }
            let ResRejRemarksFld =  formFields.RejRemarks

            if(ResRejRemarksFld.value !== undefined && ResRejRemarksFld.value !== null && ResRejRemarksFld.value !== ''){
                ResRejRemarksFld.disabled = true;
                $(formFields.RejRemarks.$parent.$el).show();
            }
             else $(formFields.RejRemarks.$parent.$el).hide();
         }
    }); 

    fd.field('ApprovalStatus').$on('change', async (value) => {
        let rejReason = formFields.RejRemarks;
       switch(value){
        case 'Approved': 
          $(rejReason.$parent.$el).hide();
          rejReason.required = false;
          break;
        case 'Rejected': 
          $(rejReason.$parent.$el).show();
          rejReason.required = true;
          break;
        case 'Cancelled': 
          $(rejReason.$parent.$el).show();
          rejReason.required = false;
          break;
       }
    });

    window.addEventListener('message', async function(event) {
        const value = event.data;
        await fetchGISInfo(value, itemId, isFound);
    });

    await _HideFields(['MasterID', 'SelectionType', 'Title'], true, true);

    // var hasOnClick = $._data($('#gwadar')[0], 'events') && $._data($('#gwadar')[0], 'events').click;
    // if (!hasOnClick)

    if(isFound)
     disableGISPMFields(true, true)//to make sure gwadar is removed if not approved.
}

let isRecordExist = async function(filterQuery){
        let result = "";
          await _web.lists
                .getByTitle("GIS Location")
                .items
                .select("Id,MasterID/Id,longitude,latitude,Option,Remarks,SelectionType,Status,ResRejRemarks")
                .filter(filterQuery)
                .expand('MasterID')
                .get()
                .then((items) => {
                    if(items.length > 0)
                      result = items[0];
                });
         return result;
}

function addMap(gisControlId){
    $(gisControlId).show();
    if($('#frameId').length === 0){
        let iframe = `<iframe id="frameId" src=${gisPath} width="100%" height="450px" frameborder="0" scrolling="no"></iframe>`;
        $(gisControlId).append(iframe);
    }
}

async function fetchGISInfo(res, itemId, isFound){
  fd.field('longitude').value = res.longitude;
  fd.field('latitude').value = res.latitude;
  fd.field('SelectionType').value = res.type;

  projNo = localStorage.getItem('projectNo')
  projNo = isNullOrEmpty(projNo) ? '' : projNo;
  fd.field('Title').value = projNo;
  fd.field('MasterID').value = itemId;



  gisItem = {
    longitude: res.longitude,
    latitude: res.latitude,
    SelectionType: res.type,
    Option: fd.field('Option').value,
    MasterIDId: itemId,
    Title: projNo
   }
}

var InsertGISItem = async function(data, isFound){
    
    if(!isFound){
        data['Remarks'] = fd.field('Remarks').value
        await _web.lists.getByTitle(GISLocation).items.add(data);
    }
    else {
        let approvalStatus = formFields.ApprovalStatus.value;
        data.Status = formFields.ApprovalStatus.value === 'Rejected' || formFields.Status.value === 'Rejected' ? 'Submitted' : formFields.ApprovalStatus.value;

        data.Remarks = fd.field('Remarks').value
        data.longitude = fd.field('longitude').value
        data.latitude = fd.field('latitude').value

        data.ApprovalStatus =  approvalStatus

        let rejRemarks = approvalStatus === 'Approved' || approvalStatus === 'Cancelled' ? '' : formFields.RejRemarks.value;
        data.ResRejRemarks = rejRemarks;

        await _web.lists.getByTitle(GISLocation).items.getById(gisId).update(data); 
    }
}

let checkGISApprovalFields = async function(){
    let fields = [], hideFields = false;

    let status = formFields.Status.value;
    localStorage.setItem('gisStatus', status);

    if(!_isGIS){
        fields.push(formFields.ApprovalStatus);
        if(status !== 'Rejected' && status !== 'Cancelled')
          fields.push(formFields.RejRemarks);
        hideFields = true

        let mesg = 'item is already submitted for GIS Approval'
        if(status !== 'Rejected'){

            if(status === 'Approved' || status === 'Cancelled')
                mesg = 'Item is already Closed'

            disableSubmission(mesg,true);
            await disableGISPMFields(true);
        }
    }
    else{
        if(status === 'In Progress'){
            fields.push(formFields.ApprovalStatus);
            formFields.ApprovalStatus.required = true;
            $(formFields.RejRemarks.$parent.$el).hide();
            disableGISPMFields(true)
        }

        if(status !== 'Submitted'){
            disableSubmission('',false);
        }
    }
    if(fields.length > 0)
     await _HideFields(fields, hideFields);
}

let disableGISPMFields = async function(isDisable, ignoreFields){

    if(ignoreFields){
        if(isFound){
          formFields.Option.disabled = isDisable;

          if(!_isGIS && (formFields.Status.value === 'Rejected' || formFields.Status.value === 'In Progress' )){}
          else formFields.Remarks.disabled = isDisable;
        }
    }

    if(isDisable){
        $(document).ready(function(){
            $('#frameId').on('load', function(){
                var iframeContent = $(this).contents();

                let status = localStorage.getItem('gisStatus') === null ? fd.field('Status').value : localStorage.getItem('gisStatus') ;
                if(status !== 'Approved')
                  iframeContent.find('#gwadar').remove();

                if(status === 'Approved' || status === 'Cancelled' || _isGIS || (!_isGIS && status !== 'In Progress' && status !== 'Rejected')){
                    iframeContent.find('#polyImg').remove();
                    iframeContent.find('#lineImg').remove();
                    iframeContent.find('#pointImg').remove();
                }
                else{
                    iframeContent.find('#gwadar').on('click', function(){
                        let proj = localStorage.getItem('projectNo')
                        proj = isNullOrEmpty(proj) ? '' : proj;
                        window.open(`https://db-gispub.darbeirut.com/gwadar/default_clean.aspx?projid=${proj}`,'_blank')
                    })
                }
                //localStorage.removeItem('gisStatus');
                if(mapTimeOut !== undefined && mapTimeOut !== null)
                  clearInterval(mapTimeOut);
            });
        });
    }
}

let sendGISApproval = async function(){
    debugger;
    let module = 'GIS';
    let projNo = isNullOrEmpty(projectNo) ? fd.field('Title').value : projectNo

    const ApprovalStatus = formFields.ApprovalStatus.value;

    let itemId = !isNaN(fd.itemId) ? fd.itemId : parseInt(localStorage.getItem('MasterId'));
    let filtQuery = `MasterID/Id eq ${itemId}`

    let result = await isRecordExist(filtQuery);

    let isNew = result.Status === 'Submitted' ? true : false;
    let emailName = isNew ? `${module}_New_Email` : `${module}_Email`,
    notName = isNew ? `${module}_New` : `${module}_${ApprovalStatus}`;

    let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${result.Id}</Value></Eq></Where>`;
    _sendEmail(module, `${emailName}|${projNo}`, query, '', notName, '');
}

function hideGISApprovalFields(){
    $(formFields.ApprovalStatus.$parent.$el).hide();
    $(formFields.RejRemarks.$parent.$el).hide();
}

let handleItemAfterApproval = async function(result){

  let Status =  '', ResRejRemarksFld = '';
  if(result !== undefined){
    Status = result.Status
    ResRejRemarksFld = result.ResRejRemarks
  }
  else{
    Status = formFields.Status.value;
    ResRejRemarksFld = formFields.RejRemarks.value;
  }
  
  if(Status !== 'In Progress'){
    if(ResRejRemarksFld !== undefined && ResRejRemarksFld !== null && ResRejRemarksFld !== ''){
        if(_isGIS && formFields.ApprovalStatus.value === ''){
           $(formFields.RejRemarks.$parent.$el).hide();
        }
        formFields.RejRemarks.disabled = true;
    }
       
    else $(formFields.RejRemarks.$parent.$el).hide();

    if(Status === 'Submitted'){
        setTimeout(async () => {
            gisItem['longitude'] = formFields.longitude.value
        }, 100);
        $(formFields.ApprovalStatus.$parent.$el).show();
    }
    else $(formFields.ApprovalStatus.$parent.$el).hide();

    if(Status === 'Rejected'){
      if(_isGIS)
         disableGISPMFields(true);
    }
    else disableGISPMFields(true)
  }
}

let disableSubmission = async function(mesg){
    setTimeout(async () => {
        setPSErrorMesg(mesg, true)
        $('span').filter(function(){ return $(this).text() == submitText; }).parent().attr("disabled", "disabled");
    }, 100);
}

let setGISMetaInfo = async function(){
    let isValid = fd.isValid
    if(isValid){


     let option = gisItem['Option'] === undefined ? fd.field('Option').value : gisItem['Option'];
     if(option !== 'Attaching A File'){
       let longitude = gisItem['longitude'] === undefined ? fd.field('longitude').value : gisItem['longitude'];
       if(longitude === undefined || longitude === null || longitude === ''){
         setPSErrorMesg('please mark a coordinate on the map and try again', true)
         return;
       }
     }
     else{
        let attachment = fd.field('Attachments').value;
        if(attachment.length === 0){
            setPSErrorMesg('Attachment is required Field', true)
            return;
        }
     }
     

     let itemId;
     if(!isNaN(fd.itemId)){
         itemId = fd.itemId
         if(fd.field('Status').value === 'Rejected' && formFields.ApprovalStatus.value === '')
           fd.field('Status').value = 'Submitted'
        else fd.field('Status').value = formFields.ApprovalStatus.value
     }
     else {
        _HideFields(['Title', 'MasterID'], false, true);

       itemId = parseInt(localStorage.getItem('MasterId'));
       projNo = localStorage.getItem('projectNo')
       projNo = isNullOrEmpty(projNo) ? '' : projNo;

       fd.field('Title').value = projNo;
       fd.field('Status').value = 'Submitted';
       fd.field('MasterID').value = itemId;

       fd.field('longitude').value;
       fd.field('latitude').value
       fd.field('SelectionType').value

       _HideFields(['Title', 'MasterID'], true, true);
     }

     debugger;
     fd.save()
     .then(()=>{
        window.close();
      });

    
       // await InsertGISItem(gisItem, isGISFound).then(()=>{
       //   let isNew = isGISFound ? false : true
       //   sendGISApproval(isNew)
       // })
       // fd.save()
       // .then(()=>{
       //   window.close();
       // });

       // .then(()=>{
       //   window.close();
       // })
    }
}