let gisPath = `${_spPageContextInfo.siteAbsoluteUrl}${_layout}/plumsail/ProjCenter/features/drawProjects2k14.html`;
// 'https://db-gispub.darbeirut.com/gwadar/default_clean.aspx?projid=AD23005-0100D'
var gisItem = {}, isFound = false, gisId;
let projNo;

var onGISRender = async function(){

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

    setPageStyle();
    let gisControlId = '#GISId';
    let gisFields = [formFields.longitude, formFields.latitude, formFields.Remarks];

    let itemId, filtQuery = '';
    
    if(!isNaN(fd.itemId)){
        itemId = fd.itemId
        filtQuery = `Id eq ${itemId}`
    }
    else{
        itemId = parseInt(localStorage.getItem('MasterId'));
        filtQuery = `MasterID/Id eq ${itemId}`
    }
    let result = await isRecordExist(filtQuery);
    _isGIS = await IsUserInGroup('GIS');

    if(result !== ''){
        isFound = true;

        //IS APPROVED OR CANCELLED SET FUNCTION TO CHECK THE LOGIC
        await handleItemAfterApproval(result);
    }
    
    localStorage.setItem('isGISFound', isFound);
    
    if(isFound){
        gisId = result.Id
        let option = result.Option
        formFields.Option.value = option
        formFields.longitude.value = result.longitude
        formFields.latitude.value = result.latitude
        formFields.Remarks.value = result.Remarks

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
    }
    
    fd.field('Option').$on('change', async (value) => {
         if(value === 'Attaching A File'){
            await _HideFields(gisFields, true);
            $(formFields.Attachments.$parent.$el).show();
            $(gisControlId).hide();
         }
         else{
            await _HideFields(gisFields, false);
            hideGISApprovalFields()
            $(formFields.Attachments.$parent.$el).hide();
            addMap(gisControlId)

            //let Status =  formFields.Status.value;
            let ResRejRemarksFld =  formFields.RejRemarks

            if(ResRejRemarksFld.value !== undefined || ResRejRemarksFld.value !== null || ResRejRemarksFld.value !== ''){
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

    await _HideFields(['MasterID', 'SelectionType'], true, true);

    // var hasOnClick = $._data($('#gwadar')[0], 'events') && $._data($('#gwadar')[0], 'events').click;
    // if (!hasOnClick)

    disableGISPMFields(true, true)//to make sure gwadar is removed if not approved.
}

let isRecordExist = async function(filterQuery){
        let result = "";
          await _web.lists
                .getByTitle("GIS Location")
                .items
                .select("Id,MasterID/Id,longitude,latitude,Option,Remarks,SelectionType")
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

  projNo = localStorage.getItem('projectNo')
  projNo = isNullOrEmpty(projNo) ? '' : projNo;
  gisItem = {
    longitude: res.longitude,
    latitude: res.latitude,
    SelectionType: res.type,
    Remarks: fd.field('Remarks').value,
    Option: fd.field('Option').value,
    MasterIDId: itemId,
    Title: projNo
   }
}

var InsertGISItem = async function(data, isFound){
    if(!isFound)
        await _web.lists.getByTitle(GISLocation).items.add(data);
    else {
        data.Status =  formFields.ApprovalStatus.value;
        data.ApprovalStatus =  formFields.ApprovalStatus.value;
        data.ResRejRemarks =  formFields.RejRemarks.value;
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

        if(status !== 'Rejected'){
            setTimeout(async () => {
                setPSErrorMesg('item is already submitted for GIS Approval', true)
                $('span').filter(function(){ return $(this).text() == 'Finalize'; }).parent().attr("disabled", "disabled");
            }, 100);
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
    }
    if(fields.length > 0)
     await _HideFields(fields, hideFields);
}

let disableGISPMFields = async function(isDisable, ignoreFields){

    if(ignoreFields){
      formFields.Option.disabled = isDisable;
      formFields.Remarks.disabled = isDisable;
    }

    if(isDisable){
        $(document).ready(function(){
            $('#frameId').on('load', function(){
                var iframeContent = $(this).contents();
                iframeContent.find('#polyImg').remove();
                iframeContent.find('#lineImg').remove();
                iframeContent.find('#pointImg').remove();

                let status = localStorage.getItem('gisStatus');
                if(status !== 'Approved')
                  iframeContent.find('#gwadar').remove();
                else{
                    iframeContent.find('#gwadar').on('click', function(){
                        let proj = localStorage.getItem('projectNo')
                        proj = isNullOrEmpty(proj) ? '' : proj;
                        window.open(`https://db-gispub.darbeirut.com/gwadar/default_clean.aspx?projid=${proj}`,'_blank')
                    })
                }
                //localStorage.removeItem('gisStatus');
            });
        });
    }
}

let sendGISApproval = async function(){
    let module = 'GIS';
    let projNo = isNullOrEmpty(projectNo) ? fd.field('Title').value : projectNo

    const ApprovalStatus = formFields.ApprovalStatus.value;
    let emailName = `${module}_Email`, notName = `${module}_${ApprovalStatus}`;
  

    let itemId = !isNaN(fd.itemId) ? fd.itemId : parseInt(localStorage.getItem('MasterId'));
    let result = await isRecordExist(itemId);

    let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${result.Id}</Value></Eq></Where>`;
    _sendEmail(module, `${emailName}|${projNo}`, query, '', notName, '');
}

function hideGISApprovalFields(){
    $(formFields.ApprovalStatus.$parent.$el).hide();
    $(formFields.RejRemarks.$parent.$el).hide();
}

let handleItemAfterApproval = async function(){

  let Status =  formFields.Status.value;
  let ResRejRemarksFld =  formFields.RejRemarks
  
  if(Status !== 'In Progress'){
    if(ResRejRemarksFld.value !== undefined && ResRejRemarksFld.value !== null && ResRejRemarksFld.value !== '')
        ResRejRemarksFld.disabled = true;
    else $(formFields.RejRemarks.$parent.$el).hide();

    $(formFields.ApprovalStatus.$parent.$el).hide();

    if(Status === 'Rejected'){
      if(_isGIS)
         disableGISPMFields(true);
    }
    else disableGISPMFields(true)
  }
}