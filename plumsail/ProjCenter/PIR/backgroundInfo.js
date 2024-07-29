  var onBackGroundInfoRender = async function (){  
    setStyles();
    ManageOtherFields(formFields)
  } 
  
  function setStyles(){
    fd.control('Mtddt').dialogOptions = {
      width: '80%',  
      height: '80%'
    }
  
    fd.control('Contactsdt').allowDragAndFill = true;
  
    let label = $('span').filter(function(){ return $(this).text() == '<LessLearn>'; })
    let text = 'Lessons learned from previous projects <a class="FieldsetLabel" href="https://bi.dar.com/ReportServer/Pages/ReportViewer.aspx?/wapps/PMISReports/PCRLessonsLearnedPublic" target="_blank" style="color:Blue;">Click here to browse the database</a>'
    label.html(text);
  
    // let btn = $('.fd-sp-datatable-toolbar button').first()
    // btn.after('<span style="font-family:Arial;font-weight:bold;font-size:14px; padding-left: 700px; color: var(--darkgreen)">Source of Technical Documents</span>');
  }
  
  function ManageOtherFields(formFields){
  
    let FormContractOther = formFields.FormContractOther
    let TypeContractOther = formFields.TypeContractOther
  
     _HideFields([FormContractOther,TypeContractOther], true, false);
  
    OtherFieldChange(formFields.ContractForm, FormContractOther)
    OtherFieldChange(formFields.ContractType, TypeContractOther)
  }
  
  function OtherFieldChange(field, otherField){
    field.$on('change', function(){
      let value = this.value;
  
      if(value === 'Other')
        {
          $(otherField.$parent.$el).show();
          otherField.required = true;
        }
      else {
        $(otherField.$parent.$el).hide();
        otherField.required = false;
      }
    })
  }

  // #region MTD

  var onMTDRender = async function (){    
    
    formFields = {
      IsMTD: fd.field("IsMTD"),
      MTD: fd.field("MTD"),
      Comment: fd.field("Comment"),
      Status: fd.field("Status"),
      Title: fd.field("Title")
    }
  
    if(_isNew){
      handleTitleColumn();
      await filterLookupList(formFields);
    }
  
    else if(_isEdit){
  
     let ApprovalStatus = fd.field("ApprovalStatus");
     if(ApprovalStatus.value === 'Approved' || ApprovalStatus.value === 'Rejected'){
        alert(`item is already action by quality team as ${ApprovalStatus.value}`)
        fd.close();
     }
  
      _isQM = await IsUserInGroup('QM');
      Object.assign(formFields, {
        ApprovalStatus: ApprovalStatus,
        Remarks: fd.field("Remarks")
  
      });
    }    
    
    setMTDStyle();
    manageMTDFieldChange(formFields);    
  }

  function setMTDStyle(){
    $('.pageContainer_e9e4af8d, .container_e9e4af8d, .pageContainer_5a558a10, .pageContainer_972f0ed1').css('margin-left', '0px');
    // setTimeout(function() {
    //   $('.fd-fixed-toolbar, .fd-form-toolbar[data-v-68f8f33f]').attr('style', 'padding-left: 29px !important');
    // }, 2000);
  }
  
  function  manageMTDFieldChange(formFields){
  
    const { IsMTD, MTD, Comment, ApprovalStatus, Remarks, Status, Title } = formFields;
    $(formFields.Status.$parent.$el).hide();
    $(formFields.Title.$parent.$el).hide();
  
    IsMTD.$on('change', function(isChecked){
      if(isChecked)
        {
          $(MTD.$parent.$el).show();
          $(Comment.$parent.$el).hide();
        }
      else {
          $(MTD.$parent.$el).hide();
          $(Comment.$parent.$el).show();
      }
    })
  
    if(IsMTD.value){
      $(Comment.$parent.$el).hide();
  
      if(_isEdit){
        $(IsMTD.$parent.$el).hide();
        $(ApprovalStatus.$parent.$el).hide();
        $(Remarks.$parent.$el).hide();
        _hideSubmit = true;
      }
      else if(_isDisplay){      
        $(fd.field('Remarks').$parent.$el).hide();
      }        
    }
    else{
      if(_isEdit){
        $(IsMTD.$parent.$el).hide();
        $(MTD.$parent.$el).hide();
        $(Remarks.$parent.$el).hide(); 
         
        if(!_isQM){
          $(ApprovalStatus.$parent.$el).hide();
          _hideSubmit = true;
        }
        else{
          ApprovalStatus.required = true;
          Comment.disabled = true;
          ApprovalStatus.$on('change', function(value){
              if(value === 'Rejected'){
                $(Remarks.$parent.$el).show();
                Remarks.required = true;
              }
              else{
                Remarks.clear();
                Remarks.required = false;
                $(Remarks.$parent.$el).hide();
              }
          });
        }
      }
      else if(_isDisplay){
        $(IsMTD.$parent.$el).hide();
        $(MTD.$parent.$el).hide(); 
        $(fd.field('Status').$parent.$el).show();     
      } 
    }    
  }
  
  let sendMTDApproval = async function(){
    let itemId = fd.itemId;
    const ApprovalStatus = formFields.ApprovalStatus.value;
    let emailName, notName;
  
    if(ApprovalStatus === 'Approved'){
      emailName = 'MTD_Approved_Email';
      notName = 'MTD_Approved';
    }
    else{
      emailName = 'MTD_Rejected_Email';
      notName = 'MTD_Rejected';
    }
    
    await setStatus(ApprovalStatus);
    let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
    _sendEmail(_module, `${emailName}|${projectNo}`, query, '', notName, '');
  }
  
  let setStatus = async function(statusText){
    const { IsMTD, MTD, Status } = formFields;
  
    $(Status.$parent.$el).show();
      Status.value = statusText;
    $(Status.$parent.$el).hide();
  
    if(!IsMTD.value){
      $(MTD.$parent.$el).show();
        MTD.value = null;
      $(MTD.$parent.$el).hide();
    }
  }
  
  const filterLookupList = async function(){
    
      const listUrl = fd.webUrl + fd.listUrl;
      const list = await _web.getList(listUrl).get();
      const listname = list.Title;
  
      let field = 'MasterID'
      let masterId = localStorage.getItem('MasterId');
      let query = `${field}/Id eq ${masterId}`;
  
      let opions = await getListOptions(listname, 'MTD', true, query)
  
      let filterQuery = '';
      opions.map( (option, index) => {
           let Id = option.value
           if(index === 0) 
             filterQuery += ` ID ne ${Id} `
           else filterQuery += `and ID ne ${Id} `
      });
  
      let mtd = fd.field("MTD");
      mtd.filter = filterQuery;
      mtd.orderBy = { field: "FullDesc", desc: false };
      mtd.refresh();
  }

  const handleTitleColumn = function(){    
    
    fd.field('MTD').$on('change', function(value) { 
      debugger;
      $(fd.field('Title').$parent.$el).show();
      fd.field('Title').value = value.LookupValue;
      $(fd.field('Title').$parent.$el).hide();
    });

    fd.field('Comment').$on('change', function(value) { 
      $(fd.field('Title').$parent.$el).show();
      fd.field('Title').value = value;
      $(fd.field('Title').$parent.$el).hide();
    });
}
  
  // #endregion
 