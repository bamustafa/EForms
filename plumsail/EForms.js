var _layout, _htLibraryUrl;

var isAllowed = false, _isLead = false, _isPart = false, _isMain = false, isContractor = false, isDar = true, _isTeamLeader = false, _isPMC = false
   _isMultiContracotr = false, _IsLeadAction = false, _isDigitalForm = false, _isSiteAdmin = false, _isAutoAssign = false, _AllowManualAssignment = false,
   _isCompliedWithConvention = false;

var hFields = ["AttachFiles", "Submit"];
var _module = '', _formType = '', Status, taskStatus = '', BIC, _role = '', _partTrade = '';

var element, targetList, targetFilter;
var counterType, counter;
var delayTime = 50, retryTime = 10;

var mTypeItem;

var disableButtons = false; // FOR validateDisciplineAgainstMatrix FUNCTION
var disableButtonsFNC = false; // FOR setNamingConvention FUNCTION
const itemsToRemove = ['Status', 'State', 'Code', 'WorkflowStatus'];

var onRender = async function (relativeLayoutPath, moduleName, formType, isLead, isPart) {

  const startTime = performance.now();
  if (formType !== "Display")
  { 
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";
  }  

  _layout = relativeLayoutPath;

  await loadScripts();
 
  clearLocalStorageItemsByField(itemsToRemove);

  if(formType == "New"){
    await LimitDateTimeSubmission();
    clearStoragedFields(fd.spForm.fields);
  }

  // var script = document.createElement("script"); // create a script DOM node
  // script.src = _layout + "/plumsail/js/config/configFileRouting.js"; // set its src to the provided URL
  // document.head.appendChild(script);
  _htLibraryUrl  = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

  _module = moduleName;
  _formType = formType;
  _isLead = isLead;
  _isPart = isPart;
  _isSiteAdmin = _spPageContextInfo.isSiteAdmin;

  if(_isPart){
   _role = fd.field("Role").value;
   _partTrade = fd.field("Trade").value;
    if(_partTrade === 'PMC')
     _isPMC = true;
  }

  if(_role === 'TeamLeader')
   _isTeamLeader = true;

  if(!_isLead && !_isPart)
    _isMain = true;

    const endTime1 = performance.now();
    const elapsedTime1 = endTime1 - startTime;
    console.log(`Execution time extractValues: ${elapsedTime1} milliseconds`);

  try {
    setTimeout(removePadding, 1000);
    
    var Trans = "A";
    if (_isMain || _isLead){
      if (formType == "New"){
          Status = "Initiated";
          
          try {fd.field('ORFI').clear();}
          catch{}
      }
      else Status = fd.field("Status").value;
    }
      
    if (_isLead || _isPart){
      Status = fd.field("Status").value;
      taskStatus = fd.field("Status").value;
    }
     
    if (formType == "Edit" && isPart) {
      var userNames = fd.field("AssignedTo").value;
      var _isAllowedUser = false;
      if (userNames != null) {
        for (var i = 0; i < userNames.length; i++) {
          var userName = userNames[i].displayName;
          
          //var isValid = false;
          //var retry = 1;
          //while (!isValid)
          //{
            //try{
              //if(retry >= retryTime) break;
            _isAllowedUser = await isUserAllowed(userName);
            // isValid = true;
            // }
              // catch{
              //   retry++;
              //   await delay(delayTime);
              // }
          //}

          if (_isAllowedUser) {
            _isAllowedUser = true;
            break;
          }
        }
        if (!_isAllowedUser) {
          alert(isAllowedUserMesg);
          return _isAllowedUser;
        } 
      }
    }
    
    if (formType == "Display") {
      script = document.createElement("script");
      script.src = _layout + "/plumsail/js/buttons.js";
      document.head.appendChild(script);

      if (moduleName == "IR" && isLead == false && isPart == false)
        disableIRFields(formType);

      if(isPart && moduleName !== 'MIR'){
        hFields = ["ph_dmg", "det_corr", "app_sub", "acc_spl"];
        HideFields(hFields, false);
      }
      else if (moduleName == "DPR"){
        //_spComponentLoader.loadScript(_layout + '/plumsail/js/commonUtils.js').then(()=> {
          //_spComponentLoader.loadScript(_layout + '/plumsail/js/customMessages.js').then(async() => {
            setFormHeaderTitle();
            await onDPRRender();
         //});
      //});
      return;
      }
    }
    else {
      //Edit
      if (formType == "New") Trans = "A";
      else {
        Trans = "U";
        if ((isLead && _AllowManualAssignment) || (isPart && _partTrade === 'PMC')) {
          checkTradeAssignment("Lead");
          checkTradeAssignment("Part");
        }
      }

      if(_isMain){
        // var isValid = false;
        // var retry = 1;
        // while (!isValid)
        // {
        //   try{
        //     if(retry >= retryTime) break;

            var _value = await getParameter("isDigitalForm");
            if(_value !== '' && _value !== undefined && _value.toLowerCase() === 'yes')
              _isDigitalForm = true;

             await setFormHeaderTitle();
        //      isValid = true;
        //   }
        //     catch{
        //       retry++;
        //       await delay(delayTime);
        //     }
        // }
    }

    if (_isMain || _isLead || _isPart) {
      //#region CHECK USER IF DAR OR CONTRACTOR
      var groupName = '';

      var isValid = false;
      var retry = 1;
      var mType;
      while (!isValid)
      {
         try{
          if(retry >= retryTime) break;

          if (_isMain){
              groupName = await isMultiContractor();

              if(_isMultiContracotr)
                isContractor = await IsUserInGroup(groupName);
              else isContractor = await IsUserInGroup('ContractorDC');
          }

          mTypeItem = await getMajorType(_module);
          mTypeItem = mTypeItem[0];
          if(_isPart){
            var codes = mTypeItem.Code;
            if(codes !== null && codes != '' && codes !== undefined && codes.length > 0)
              fd.field('Code').widget.dataSource.data(codes);
          }
          else{
            try { _isCompliedWithConvention = mTypeItem.IsCompliedWithConvention; } catch(e){}
          }
          isValid = true;
        }
        catch{
          retry++;
          await delay(delayTime);
        }
      }
      isValid = false;
      
      if(isContractor) 
        isDar = false;

        if (_isMain && formType === "New") {
        if (moduleName !="MOM" && moduleName !== "SLF" && moduleName !== "DPR") {
          if (moduleName == "SI" || moduleName == "NCR") 
              setOldRef("Closed");
          else setOldRef("Issued to Contractor");
        }
      }
      //#endregion
    }


    if (moduleName == "SI") 
      await onSIRender(moduleName, formType, hFields);
    else if (moduleName == "IR" || moduleName == "MIR")
        await onIRRender(moduleName, formType, hFields, isLead, isPart);
    else if (moduleName == "MAT" || moduleName == "SCR")
      await onMATRender(moduleName, formType, hFields, isLead, isPart);
    else if (moduleName == "DPR")
      await onDPRRender();
    else if (moduleName == "SLF")
      await onSLFRender(hFields, isLead, isPart);

    if(_isMain && _formType === 'New' && _module !== 'SI' && _module !== 'DPR')
        await validateDisciplineAgainstMatrix();
    
    if(_isPart && _formType === 'Edit' && (_isTeamLeader || _IsLeadAction || _partTrade === 'PMC') && moduleName != 'MIR'){
      var listname = "Part Action";
      var byRef = true;
      if (moduleName == "MAT" || moduleName == "SCR") byRef = false;

      await onPageRender(moduleName, listname, byRef); 
    }

    await setButtons(moduleName, formType, Trans, Status, isLead, isPart);

    //_spComponentLoader.loadScript(_layout + '/plumsail/js/customMessages.js').then(async() => {
       setToolTipMessages();
    //});
    fixTextArea();
    }

    // if(_isMain && _isCompliedWithConvention)
    //   await renderFieldsOnForm();
  }
  
  catch (e) {
    alert(e);
    console.log(e);
    
    preloader();
    var webURL = document.URL.substring(0, document.URL.indexOf('/PlumsailForms')).replace('SitePages','');
    //window.location.href = webURL;
  }
  const endTime = performance.now();
  const elapsedTime = endTime - startTime;
  console.log(`Execution time onRender: ${elapsedTime} milliseconds`);
};

//#region (IR & MIR) RENDER MODULE
var onIRRender = async function (moduleName, formType, hFields, isLead, isPart) {
  if (formType == "New") {
    if (moduleName == "MIR" || moduleName === 'IR'){

      if(moduleName == "MIR"){
        hFields = ["AttachFiles", "Submit", "Status"];
        fd.field('Question').placeholder = 'Question';
      }

      if(_isCompliedWithConvention !== undefined){
        let schema = await getAutReferenceFormat();
        await renderFieldsOnForm(schema);
          if(!_isCompliedWithConvention){
            try{
               fd.field('Reference').value;
               hFields.push('Reference');
            }
            catch{}
          }
          else setNamingConvention('Reference');
      }
    }
    else {
      // if(isDar){
      //   fd.field("HardcopyDate").value = new Date();
      //   fd.field("HardcopyDate").disabled = true;
      // }
      hFields = ["AttachFiles", "Submit", "BuildingType", "Level", "Room", "Road", "FromStation", "ToStation", "Zone"];
      fd.field('Remarks').placeholder = 'Remarks';
    }
  } 
  else if (formType == "Edit") {
    var fields;

    if (isLead) {
      isAllowed = await delayshowHideFunction();
          
      hFields = ["LeadTrade", "PartTrades", "AssignedDate", "Assigned", "WorkflowStatus", "BIC", "SentToPMC", "AttachFiles"];
      if(_isAutoAssign)
        hFields.push("Lead");
      fields = ["Status", "Attachments"];
    } 
    else if (isPart) {
      if (moduleName == "MIR")
        hFields = ["Role", "isPMCAssignment", "IsLeadAction", "AttachFiles", "isRejected"];
      else hFields = ["ph_dmg", "det_corr", "app_sub", "acc_spl", "Role", "isPMCAssignment", "IsLeadAction", "AttachFiles", "isRejected"];
      hFields.push('Resubmit');

      if (Status == "Completed") {
        if (moduleName == "MIR" && isPart)
          fields = ["ph_dmg", "det_corr", "app_sub", "acc_spl", "Status", "Code", "Comment", "Attachments"];
        else fields = ["Status", "Code", "Comment", "Attachments"];
      } 
      else fields = ["Status", "Attachments"];

      var RejTrades = fd.field("RejectionTrades").value;
      if(RejTrades === '' || RejTrades === null || RejTrades === undefined)
         hFields.push('RejectionTrades');
      else fields.push('RejectionTrades');

      isAllowed = await showHideTabs(moduleName, BIC, true, false);
    } 
    else {
      if (moduleName == "MIR" && isLead == false && isPart == false) {
        fields = [ "Reference", "Rev", "Title", "Question", "HardcopyDate", "MaterialSubmittalNo", "MATReferenceTitle", "MATReferenceRecDate", "MATReferenceBOQ",
                   "MATReferenceSpecs", "Quantity", "Inspection_x0020_Date", "Manufacturer", "ModelNo", "SerialNo", "DateofARRIVALonSite", "CountryofOrigin",
                   "STORAGELocation", "Submittedby", "MaterialSubmittalNo", "Attachments", "Status"];
        
        const fields1 = ["MATReferenceTitle", "MATReferenceRecDate", "MATReferenceBOQ", "MATReferenceSpecs", "Manufacturer"];
        handleHardCopyDate(fields, hFields);
        await DisableFields("Status", "Initiated", "eq", fields1, true, false, true);
      } else {
        disableIRFields(formType);
        fields = ["Status", "Attachments"];
        handleHardCopyDate(fields, hFields);
      }
      hFields = ["AttachFiles", "Submit", "Status"];
      isAllowed = await showHideTabs(moduleName);
    }

 
    if (Status == "Closed" || Status == "Issued to Contractor" || taskStatus == "Completed")
      await DisableFields("Status", "Initiated", "neq", fields, true, true, true);
    else await DisableFields("Status", "Initiated", "neq", fields, true, false, true);
  }

  handleHardCopyDate(fields, hFields);
  HideFields(hFields, true);
  if (moduleName == "MIR" && isLead == false && isPart == false)
    setMATMetaInfo(formType);
  else if (moduleName == "IR" && isLead == false && isPart == false)
    setCascadedValues();
};
//#endregion

//#region (SI MODULE) RENDER MODULE
var onSIRender = async function(modulename, formType, hFields) {
  if (formType == "New") {
    fd.field("SIDate").value = new Date();
    hFields = ["AttachFiles", "Submit"];

    if(_isCompliedWithConvention !== undefined){
      if(!_isCompliedWithConvention){
        let schema = await getAutReferenceFormat();
        await renderFieldsOnForm(schema);
        try{
           fd.field('Reference').value;
           hFields.push('Reference');
        }
        catch{}
      }
      else setNamingConvention('Reference');
    }

    HideFields(hFields, true);
    fd.field('InstDetails').placeholder = 'Instruction Details';
    
    //$("textarea.fd-textarea").attr("placeholder", "Question");
  } else if (formType == "Edit") {
    hFields = ["AttachFiles", "Submit", "Issued", "IsValid", "IsAcknowledged", "ReasonofRejection"];
    HideFields(hFields, true);

    var dFields = ["Response", "AnsweredDate", "ResDate", "Status", "ResponsedBy"];
    await DisableFields("Status", "Initiated", "neq", dFields, true, true);

    var AttdFields = ["Attachments"];
    await DisableFields("Status", "Closed", "eq", AttdFields, true, true);

    if (Status == "Initiated") SIisUserInGroup("SIOriginator");
    else {
      if (Status == "Open") SIisUserInGroup("RE");
      else if (Status == "Received by Contractor")
        SIisUserInGroup("ContractorDC");
      else if (Status == "Received by Originator")
        SIisUserInGroup("SIOriginator");
      else if (Status == "Acknowledged by Originator") SIisUserInGroup("RE");
    }
  }
}
//#endregion

//#region (MAT, SCR MODULE) RENDER MODULE
var onMATRender = async function (moduleName, formType, hFields, isLead, isPart) {
  if (formType == "New"){
    hFields = ["AttachFiles", "Submit", "Status"];

    //if(_isCompliedWithConvention !== undefined){
      if(!_isCompliedWithConvention){
        let schema = await getAutReferenceFormat();
        await renderFieldsOnForm(schema);
        try{
           fd.field('Reference').value;
           hFields.push('Reference');
        }
        catch{}
      }
      else setNamingConvention('Reference');
    //}

    if (moduleName == "SCR")
     fd.field('Question').placeholder = 'Question';
  }
  else if (formType == "Edit") {
    var fields;
    if (isLead) {
      BIC = fd.field("BIC").value;
      isAllowed = await delayshowHideFunction();

      hFields = ["AttachFiles", "LeadTrade", "PartTrades", "AssignedDate", "Assigned", "WorkflowStatus", "BIC", "SentToPMC", "Trade"];
      if(_isAutoAssign){
       hFields.push("Lead");
       hFields.push("LeadTrade");
      }
      fields = ["Status"];
    } 
    else if (isPart) {
      _IsLeadAction = fd.field("IsLeadAction").value;
      _partTrade = fd.field("Trade").value;
      hFields = ["AttachFiles", "Status", "Assigned", "LeadTrade", "PartTrades", "Role", "isPMCAssignment", "IsLeadAction", "ph_dmg", "det_corr", 
                 "app_sub", "acc_spl", "isRejected", "Resubmit"];
      if (Status == "Completed")
        fields = ["Status", "Code", "Comment", "Attachments"];
      else fields = ["Status", "Attachments"];

      var RejTrades = fd.field("RejectionTrades").value;
      if(RejTrades === '' || RejTrades === null || RejTrades === undefined)
         hFields.push('RejectionTrades');
      else fields.push('RejectionTrades');

      isAllowed = await showHideTabs(moduleName, "", isPart);
    } 
    else {
      hFields = ["AttachFiles", "Submit", "Status"];
      if (moduleName == "MAT")
        fields = [
          "Reference", "Rev", "Title", "SubmittedDate", "Area", "Discipline", "SubDiscipline", "Drawingref", "BOQ", "Specs", "Standards", "Manufacturer", "Address", "Agent", 
          "Availability", "Country", "TotalDuration", "ArrivalTime", "MatDate", "Latestdate", "SubmittedBy", "FinalResponse", "Attachments", "Status"];
      else
      {
        fields = ["Reference", "Rev", "Title", "Question", "SubmittedDate", "Discipline", "SubDiscipline", "SubmittedBy", "FinalResponse", "Attachments", "Status"];
        if (Status === "Initiated")
        hFields.push("FinalResponse");
      }
    }

    handleHardCopyDate(fields, hFields);
    // if (Status == "Closed" || Status == "Issued to Contractor" || Status == "Completed")
    //   await DisableFields("Status", "Initiated", "neq", fields, true, true, true);
     if (Status == "Initiated")
     await DisableFields("Status", "Initiated", "eq", fields, true, false, true);
    else await DisableFields("Status", "Initiated", "neq", fields, true, true, true);
  }

  HideFields(hFields, true);
  if (_isMain) 
   setMATCascadedValues(formType);
};
//#endregion

//#region SLF RENDER MODULE
var onSLFRender = async function (hFields, isLead, isPart) {
	var fields = [];
  var _Status = 'Pending';
  var zoneCount = await getZoneCount();
  var ReviewedURL;
  
  //isSubmit = fd.field("Submit").value;

  if (_formType == "New")
  {
    //Fields = ["ReviewedURL"];
  }
  else if(_formType == "Edit"){
    if(_isMain)
     ReviewedURL = fd.field("ReviewedURL").value;
  var operator = 'eq';
    if (isLead) {
      isAllowed = await delayshowHideFunction();

      hFields = ["Lead", "LeadTrade", "PartTrades", "AssignedDate", "Assigned", "WorkflowStatus", "BIC", "SentToPMC", "AttachFiles"];
      fields = ["Status", "Attachments"];
      _Status = fd.field("Status").value;
    }
    else if (_isMain || isPart) {

      if(isPart){
        filterStatus = fd.field("Status").value;
        hFields = ["Status", "PartComments", "ContAttachUrl", "ph_dmg", "det_corr", "app_sub", "acc_spl", "Role", "Code", "isPMCAssignment", 
                    "IsLeadAction", "AttachFiles", "Attachments", "isRejected","Resubmit"]; //"RejectionTrades"

        
        if(zoneCount === 0)
         hFields.push('Zone');

        if(!_isTeamLeader){
          $("p:contains('Distribution')").hide();
          $('#parttasks').hide();
        }

        var RejTrades = fd.field("RejectionTrades").value;
        if(RejTrades === '' || RejTrades === null || RejTrades === undefined)
            hFields.push('RejectionTrades');
        else fields.push('RejectionTrades');

        //else hFields.push('RejectionTrades');
        // else{
        //   $("p:contains('Distribution')").show();
        //   $('#parttasks').show();
        // }

        if (_Status == "Completed") {
            fields = ["ph_dmg", "det_corr", "app_sub", "acc_spl", "Status", "Code", "Comment", "Attachments"];
        } 
        else {
          fields = ["Status", "Comment", "Attachments"];
          //_Status = ''
        }
        $("ul.nav,nav-tabs").remove();
      }

      else {
        isSubmit = fd.field("Submit").value;
        if(isSubmit || Status !== "Initiated")
         fields = ["Reference", "Title", "BuildingType", "LeadInspector"];

         if(zoneCount > 0)
          fields.push('Zone');

         if(isContractor)
           fields.push("Part");
           operator = 'neq';

           if(ReviewedURL === null)
            hFields.push('ReviewedURL')
      }

      if(_isMain && ReviewedURL !== null){}
      else setSLFGrid(true, false);
    } 

    await DisableFields("Status", _Status, operator, fields, true, false, true);
  }
   if (_isMain){

    fd.field("BuildingType").orderBy = { field: "Title", desc: false };

    if (_formType == "New" || _formType == "Edit"){
      if(_formType === 'New'){
        fd.field("BuildingType").value = null;
    }

      var building = fd.field("BuildingType").value;

      await setLeadPartTrades();
      hFields = ["Attachments", "AttachFiles", "Status", "Submit", "PartTrades", "AssignedDate"];
     
      if(_formType == "Edit")
      hFields.push('Assigned', 'Resubmit');
    }
  }
  if(_isMain || _isPart){
   if(zoneCount === 0)
    hFields.push('Zone');
   else fd.field("Zone").required = true;
  }

   HideFields(hFields, true);
};
//#endregion

//#region DPR RENDER MODULE
var onDPRRender = async function () {
  
  if(_formType === 'New'){
    _isCompliedWithConvention = _isCompliedWithConvention !== undefined ? _isCompliedWithConvention : false;

    if(!_isCompliedWithConvention){
      //let schema = await getAutReferenceFormat();
      //await renderFieldsOnForm(schema);
      try{
          fd.field('Reference').value;
          hFields.push('Reference');
      }
      catch{}
    }
    else setNamingConvention('Reference');
    
    var currentDate = new Date();
    var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    
    fd.field('Date').value = currentDate; 
    fd.field('Day').value = daysOfWeek[currentDate.getDay()];
  }

  fd.field('Day').disabled = true;
  fd.field('TemperatureMini').$on('change', function(value) {
       var TemperatureMaxi = fd.field('TemperatureMaxi').value;
       if(TemperatureMaxi !== null) {
          if(value > TemperatureMaxi)
           {
               alert('Minimum Temparature cant be higher than the Maximum Temparature');
               fd.field('TemperatureMini').value = "";
           }
       }
  });
  
  fd.field('TemperatureMaxi').$on('change', function(value) {
       var TemperatureMini = fd.field('TemperatureMini').value;
       if(TemperatureMini !== null) {
          if(value < TemperatureMini)
          {
              alert('Minimum Temparature cant be higher than the Maximum Temparature');
              fd.field('TemperatureMaxi').value = "";
          }
       }
  });
  
  fd.field('Date').$on('change', function(value) { 
       if(value !== null) {  
           fd.field('Day').value = daysOfWeek[value.getDay()];
           fd.field('Day').disabled = true;   
       }
  });

  var isSubmit = fd.field("Submit").value;
 var padding = '22px 0px 0px 207px';
 var isReadOnly = false;
  if(_formType === 'Edit'){ // || _formType === 'Display'){

    if(Status !== 'Initiated'){
      isReadOnly = true;
      padding = '0px 0px 0px 150px';
      var fields = ["Title", "Date",  "Weather", "SignedBy", "TemperatureMini", "TemperatureMaxi", "Attachments"];
      await DisableFields("Status", "Initiated", "neq", fields, true, true, true);
    
      var gridFields = ['OnSiteStaffLabour', 'OnSitePlant', 'DeliveryofMaterial', 'VisitorsonSite', 'DescriptionofWork', 'Remarks'];
      gridFields.map((fld) =>{
        let controlField = fd.control(fld);
        controlField.ready(function(dt) {  
          this.readonly = true;  

          // var columns = this._columnsSettings;
          
          // if(fld === 'OnSiteStaffLabour')
          //   columns.Trade['width'] = 700;
          dt.buttons[2].visible = false;        
        });
      });
    }
  }

  await _spComponentLoader.loadScript( _layout + '/plumsail/js/utilities.js').then(async () =>{
    hFields = ["AttachFiles", "Submit", "Status"];

    if(!_isCompliedWithConvention)
      hFields.push('Reference');
 
    if(_formType === 'Display')
      HideFields(hFields, false);
    else HideFields(hFields, true);
	});

  fd.control('OnSiteStaffLabour').$on('edit', function(editData) {
    var elements = $('div.fd-inline-editor');
    elements.eq(0).css('width', '200px').css('min-width', '');
    elements.eq(1).add(elements.eq(2)).css('width', '80px').css('min-width', '');

    $('fd-form,fd-sp-datatable-wrapper>fd-sp-datatable,fd-inline-editor').css('min-width', '');
  });

  if(!isReadOnly)
   $('.k-grid-header table, .k-grid-footer table').css('width', '').css('table-layout', 'auto');

  $('div.col-sm-4 p span').css({
     'position': 'absolute', 
     'padding': padding
  }); 

  await setData();
}
//#endregion


//#region SI FUNCTIONS
function SIisUserInGroup(group) {
  try {
    pnp.sp.web.currentUser.get().then(function (user) {
      pnp.sp.web.siteUsers
        .getById(user.Id)
        .groups.get()
        .then(function (groupsData) {
          for (var i = 0; i < groupsData.length; i++) {
            if (groupsData[i].Title == group) {

              document.querySelectorAll('.k-button.k-button-icon.k-flat.k-upload-action').forEach(function(element) {
                element.style.display = 'none';
              });

              if (fd.field("Status").value == "Open") {
                $(fd.field("ResDate").$parent.$el).hide();
                $(fd.field("AnsweredDate").$parent.$el).hide();
                $(fd.field("ResponsedBy").$parent.$el).hide();
                $(fd.field("Response").$parent.$el).hide();

                fd.toolbar.buttons.push({
                  icon: "Accept",
                  class: "btn-outline-primary",
                  text: "Approve",
                  click: function () {
                    $(fd.field("Issued").$parent.$el).show();
                    $(fd.field("IsValid").$parent.$el).show();

                    fd.field("Status").disabled = false;
                    fd.field("Status").value = "Received by Contractor";
                    fd.field("Status").disabled = true;

                    fd.field("Issued").value = false;
                    fd.field("IsValid").value = true;

                    $(fd.field("Issued").$parent.$el).hide();
                    $(fd.field("IsValid").$parent.$el).hide();

                    var NCountofATT = 0;
                    var OCountofATT = 0;
                    for (i = 0; i < fd.field("Attachments").value.length; i++) {
                      var Val = fd
                        .field("Attachments")
                        .value[i].extension.toString();
                      if (Val === "") {
                        OCountofATT++;
                      } else {
                        IsNewAttachment = true;
                        NCountofATT++;
                      }
                    }
                    $(fd.field("AttachFiles").$parent.$el).show();
                    fd.field("AttachFiles").value =
                      "U" + "," + NCountofATT + "," + OCountofATT;
                    $(fd.field("AttachFiles").$parent.$el).hide();

                    setTimeout(function () {
                      fd.save();
                    }, 300);
                  },
                });

                fd.toolbar.buttons.push({
                  icon: "ChromeClose",
                  class: "btn-outline-primary",
                  text: "Cancel",
                  click: function () {
                    $(fd.field("Issued").$parent.$el).show();
                    $(fd.field("IsValid").$parent.$el).show();
                    fd.field("Status").disabled = false;
                    fd.field("Status").value = "Closed";
                    fd.field("Status").disabled = true;
                    fd.field("IsValid").value = false;
                    fd.field("Issued").value = false;
                    $(fd.field("Issued").$parent.$el).hide();
                    $(fd.field("IsValid").$parent.$el).hide();

                    var NCountofATT = 0;
                    var OCountofATT = 0;
                    for (i = 0; i < fd.field("Attachments").value.length; i++) {
                      var Val = fd
                        .field("Attachments")
                        .value[i].extension.toString();
                      if (Val === "") {
                        OCountofATT++;
                      } else {
                        IsNewAttachment = true;
                        NCountofATT++;
                      }
                    }
                    $(fd.field("AttachFiles").$parent.$el).show();
                    fd.field("AttachFiles").value =
                      "U" + "," + NCountofATT + "," + OCountofATT;
                    $(fd.field("AttachFiles").$parent.$el).hide();

                    setTimeout(function () {
                      fd.save();
                    }, 300);
                  },
                });
              } else if (fd.field("Status").value == "Received by Contractor") {
                fd.field("Response").disabled = false;
                fd.field("Response").required = true;
                fd.field("ResDate").disabled = false;
                fd.field("ResDate").required = true;
                fd.field("ResDate").value = new Date();
                fd.field("ResDate").disabled = true;
                fd.field("ResponsedBy").disabled = false;
                $(fd.field("ReasonofRejection").$parent.$el).hide();                

                fd.toolbar.buttons.push({
                  icon: "Accept",
                  class: "btn-outline-primary",
                  text: "Send to Originator",
                  click: function () {
                    $(fd.field("Issued").$parent.$el).show();
                    fd.field("Status").disabled = false;
                    fd.field("Status").value = "Received by Originator";
                    fd.field("Status").disabled = true;
                    fd.field("Issued").value = false;

                    $(fd.field("Issued").$parent.$el).hide();

                    var NCountofATT = 0;
                    var OCountofATT = 0;
                    for (i = 0; i < fd.field("Attachments").value.length; i++) {
                      var Val = fd
                        .field("Attachments")
                        .value[i].extension.toString();
                      if (Val === "") {
                        OCountofATT++;
                      } else {
                        IsNewAttachment = true;
                        NCountofATT++;
                      }
                    }
                    $(fd.field("AttachFiles").$parent.$el).show();
                    fd.field("AttachFiles").value =
                      "U" + "," + NCountofATT + "," + OCountofATT;
                    $(fd.field("AttachFiles").$parent.$el).hide();

                    setTimeout(function () {
                      fd.save();
                    }, 300);
                  },
                });
              } else if (fd.field("Status").value == "Received by Originator") {
                fd.field("Response").disabled = false;
                $(fd.field("ReasonofRejection").$parent.$el).show();
                fd.field("ReasonofRejection").disabled = true;

                fd.toolbar.buttons.push({
                  icon: "Accept",
                  class: "btn-outline-primary",
                  text: "Acknowledge",
                  click: function () {
                    $(fd.field("Issued").$parent.$el).show();
                    fd.field("Status").disabled = false;

                    $(fd.field("IsAcknowledged").$parent.$el).show();
                    fd.field("IsAcknowledged").value = true;
                    $(fd.field("IsAcknowledged").$parent.$el).hide();
                    fd.field("Status").value = "Acknowledged by Originator";
                    //fd.field('Response').value = "";

                    if (fd.field("IsAcknowledged").value == true)
                      fd.field("Status").disabled = true;
                    fd.field("Issued").value = false;
                    $(fd.field("Issued").$parent.$el).hide();

                    var NCountofATT = 0;
                    var OCountofATT = 0;
                    for (i = 0; i < fd.field("Attachments").value.length; i++) {
                      var Val = fd
                        .field("Attachments")
                        .value[i].extension.toString();
                      if (Val === "") {
                        OCountofATT++;
                      } else {
                        IsNewAttachment = true;
                        NCountofATT++;
                      }
                    }
                    $(fd.field("AttachFiles").$parent.$el).show();
                    fd.field("AttachFiles").value =
                      "U" + "," + NCountofATT + "," + OCountofATT;
                    $(fd.field("AttachFiles").$parent.$el).hide();

                    setTimeout(function () {
                      fd.save();
                    }, 300);
                  },
                });

                fd.toolbar.buttons.push({
                  icon: "ChromeClose",
                  class: "btn-outline-primary",
                  text: "Reject",
                  click: function () {
                    fd.field("Response").required = true;

                    $(fd.field("Issued").$parent.$el).show();
                    fd.field("Status").disabled = false;

                    $(fd.field("IsAcknowledged").$parent.$el).show();
                    fd.field("IsAcknowledged").value = false;
                    $(fd.field("IsAcknowledged").$parent.$el).hide();
                    fd.field("Status").value = "Received by Contractor";

                    fd.field("Status").disabled = true;
                    fd.field("Issued").value = false;
                    $(fd.field("Issued").$parent.$el).hide();

                    var NCountofATT = 0;
                    var OCountofATT = 0;
                    for (i = 0; i < fd.field("Attachments").value.length; i++) {
                      var Val = fd
                        .field("Attachments")
                        .value[i].extension.toString();
                      if (Val === "") {
                        OCountofATT++;
                      } else {
                        IsNewAttachment = true;
                        NCountofATT++;
                      }
                    }
                    $(fd.field("AttachFiles").$parent.$el).show();
                    fd.field("AttachFiles").value =
                      "U" + "," + NCountofATT + "," + OCountofATT;
                    $(fd.field("AttachFiles").$parent.$el).hide();

                    setTimeout(function () {
                      fd.save();
                    }, 300);
                  },
                });
              } else if (
                fd.field("Status").value == "Acknowledged by Originator"
              ) {
                fd.field("AnsweredDate").disabled = false;
                fd.field("AnsweredDate").required = true;
                fd.field("AnsweredDate").value = new Date();
                fd.field("AnsweredDate").disabled = true;

                $(fd.field("ReasonofRejection").$parent.$el).show();
                fd.field("ReasonofRejection").disabled = false;

                fd.toolbar.buttons.push({
                  icon: "Accept",
                  class: "btn-outline-primary",
                  text: "Submit Final Response",
                  click: function () {
                    $(fd.field("Issued").$parent.$el).show();
                    fd.field("Status").disabled = false;
                    fd.field("Status").value = "Closed";
                    fd.field("Status").disabled = true;
                    fd.field("Issued").value = false;
                    $(fd.field("Issued").$parent.$el).hide();

                    var NCountofATT = 0;
                    var OCountofATT = 0;
                    for (i = 0; i < fd.field("Attachments").value.length; i++) {
                      var Val = fd
                        .field("Attachments")
                        .value[i].extension.toString();
                      if (Val === "") {
                        OCountofATT++;
                      } else {
                        IsNewAttachment = true;
                        NCountofATT++;
                      }
                    }
                    $(fd.field("AttachFiles").$parent.$el).show();
                    fd.field("AttachFiles").value =
                      "U" + "," + NCountofATT + "," + OCountofATT;
                    $(fd.field("AttachFiles").$parent.$el).hide();

                    setTimeout(function () {
                      fd.save();
                    }, 300);
                  },
                });

                fd.toolbar.buttons.push({
                  icon: "ChromeClose",
                  class: "btn-outline-primary",
                  text: "Reject",
                  click: function () {
                    fd.field("ReasonofRejection").required = true;

                    $(fd.field("Issued").$parent.$el).show();
                    fd.field("Status").disabled = false;
                    fd.field("Status").value = "Received by Originator";
                    fd.field("Status").disabled = true;
                    fd.field("Issued").value = false;
                    $(fd.field("Issued").$parent.$el).hide();

                    var NCountofATT = 0;
                    var OCountofATT = 0;
                    for (i = 0; i < fd.field("Attachments").value.length; i++) {
                      var Val = fd
                        .field("Attachments")
                        .value[i].extension.toString();
                      if (Val === "") {
                        OCountofATT++;
                      } else {
                        IsNewAttachment = true;
                        NCountofATT++;
                      }
                    }
                    $(fd.field("AttachFiles").$parent.$el).show();
                    fd.field("AttachFiles").value =
                      "U" + "," + NCountofATT + "," + OCountofATT;
                    $(fd.field("AttachFiles").$parent.$el).hide();

                    setTimeout(function () {
                      fd.save();
                    }, 300);
                  },
                });
              } else if (fd.field("Status").value == "Initiated") {
                $(fd.field("ResDate").$parent.$el).hide();
                $(fd.field("AnsweredDate").$parent.$el).hide();
                $(fd.field("ResponsedBy").$parent.$el).hide();
                $(fd.field("Response").$parent.$el).hide();

                fd.toolbar.buttons.push({
                  icon: "Accept",
                  class: "btn-outline-primary",
                  text: "Submit",
                  click: function () {
                    var NCountofATT = 0;
                    var OCountofATT = 0;
                    for (i = 0; i < fd.field("Attachments").value.length; i++) {
                      var Val = fd
                        .field("Attachments")
                        .value[i].extension.toString();
                      if (Val === "") {
                        OCountofATT++;
                      } else {
                        IsNewAttachment = true;
                        NCountofATT++;
                      }
                    }
                    $(fd.field("AttachFiles").$parent.$el).show();
                    fd.field("AttachFiles").value =
                      "U" + "," + NCountofATT + "," + OCountofATT;
                    $(fd.field("AttachFiles").$parent.$el).hide();

                    $(fd.field("Submit").$parent.$el).show();
                    fd.field("Submit").value = true;
                    fd.save();
                  },
                });
              }
              break;
            }
          }
        });
    });
  } catch (e) {
    alert(e);
  }
}
//#endregion

//#region IR FUNCTIONS
function setCascadedValues() {
  clearDDL();

  fd.field("Discipline").ready()
    .then(async function () {
      fd.field("Discipline").orderBy = { field: "Title", desc: false };
      fd.field("Discipline").refresh();
      //filter Products when Category changes

      if(_isCompliedWithConvention === true){
        fd.field("Discipline").$on("change", async function (value) {
          debugger;
          let result = await pnp.sp.web.lists.getByTitle("Discipline").items.select("InspType/Id").expand("InspType").filter(`Title eq '${  value.LookupValue  }'`).get();
          if(result.length > 0){
            let inspTypesId = result[0].InspType;
            filterInspType(inspTypesId);
          }
        })
      }


      fd.field("InspType").$on("change", async function (value) {
        
        filterDiscipline(value);

        //if fnc enabled
        if(_isCompliedWithConvention === true){

          let discField = fd.field("Discipline");
          let currentDiscipline = discField.value !== null ? discField.value.LookupValue : null;
            if(currentDiscipline !== null){
              let query = `Title eq '${currentDiscipline}' and InspType/Id eq '${value.LookupId}'`
              let item = await pnp.sp.web.lists.getByTitle('Discipline').items.select("Title").filter(query).get();
              var refId = $("#null");
              if(item.length === 0){
                let inspType = value.LookupValue;
                let mesg = `${currentDiscipline} is not allowed under ${inspType}`;
                $(refId).html(mesg).attr('style', 'color: red !important');

                setDimonButton('Save', true);
                setDimonButton('Submit', true);
                fd.field('Reference').disabled = false;
              }
              else{
                $(refId).html('filename matches naming convention').attr('style', 'color: green !important');
                setDimonButton('Save', false);
                setDimonButton('Submit', false);
                fd.field('Reference').disabled = true;
              }
              filterPurpose(discField.value)
          }
        }
        else fd.field("Discipline").value = null;

        fd.field("InspDesc").value = null;
        fd.field("InspPurpose").value = null;

        showHideIRfields(value);
      });

      fd.field("InspType")
        .ready()
        .then(function (field) {
          fd.field("InspType").orderBy = { field: "Title", desc: false };
          fd.field("InspType").refresh();
          filterDiscipline(field.value);
        });
    });

  fd.field("InspPurpose").ready()
    .then(function () {
      fd.field("InspPurpose").orderBy = { field: "Title", desc: false };
      fd.field("InspPurpose").refresh();
      fd.field("Discipline").$on("change", function (value) {
        filterPurpose(value);
        fd.field("InspPurpose").value = null;
      });

      fd.field("Discipline")
        .ready()
        .then(function (field) {
          filterPurpose(field.value);
        });
    });

  fd.field("InspDesc").ready()
    .then(function () {
      fd.field("InspDesc").orderBy = { field: "Title", desc: false };
      fd.field("InspDesc").refresh();

      fd.field("InspPurpose").$on("change", function (value) {
        filterPurposeType(value);
        fd.field("InspDesc").value = null;

        if (
          fd.field("InspType").value == null ||
          fd.field("Discipline").value == null ||
          fd.field("InspPurpose").value.length == 0
        ) {
          fd.field("InspDesc").filter = "Discipline/Id eq 0";
          fd.field("InspDesc").refresh();
        }
      });

      fd.field("InspPurpose")
        .ready()
        .then(function (field) {
          fd.field("InspPurpose").orderBy = { field: "Title", desc: false };
          fd.field("InspPurpose").refresh();
          filterPurposeType(field.value);
        });
    });
}

function filterDiscipline(inspType) {
  var inspTypeId = (inspType && inspType.LookupId) || inspType || null;
  fd.field("Discipline").filter = "InspType/Id eq " + inspTypeId;
  fd.field("Discipline").orderBy = { field: "Title", desc: false };
  fd.field("Discipline").refresh();
}

function filterPurpose(discipline) {
  if (discipline != null && discipline !== undefined && fd.field("InspType").value !== null) {
    var disciplineId = (discipline && discipline.LookupId) || discipline || null;
    fd.field("InspPurpose").filter = "Discipline/Id eq " + disciplineId + " and InspType/Id eq " + fd.field("InspType").value.LookupId;
    fd.field("InspPurpose").orderBy = { field: "Title", desc: false };
    fd.field("InspPurpose").refresh();
  }
}

function filterPurposeType(purpose) {
  if (purpose !== null && purpose !== undefined && purpose.length > 0) {
    var purposeId = (purpose && purpose.LookupId) || purpose || null;
    var disciplineId = fd.field("Discipline").value.LookupId;
    var insptypeId = fd.field("InspType").value.LookupId;

    var subQuery = "";
    var j = 1;
    for (var i = 0; i < purpose.length; i++) {
      if (purpose.length == j)
        subQuery += "InspPurpose/Id eq " + purpose[i].LookupId;
      else subQuery += "InspPurpose/Id eq " + purpose[i].LookupId + " or ";
      j++;
    }

    fd.field("InspDesc").filter =
      "Discipline/Id eq " +
      disciplineId +
      " and InspType/Id eq " +
      insptypeId +
      " and ( " +
      subQuery +
      " )";
    fd.field("InspDesc").orderBy = { field: "Title", desc: false };
    fd.field("InspDesc").refresh();
  }
}

function filterInspType(InspTypes){
  var subQuery = "";
    var j = 1;
    for (var i = 0; i < InspTypes.length; i++) {
      if (InspTypes.length == j)
        subQuery += "Id eq " + InspTypes[i].Id;
      else subQuery += "Id eq " + InspTypes[i].Id + " or ";
      j++;
    }

    fd.field("InspType").filter = subQuery;
    fd.field("InspType").orderBy = { field: "Title", desc: false };
    fd.field("InspType").refresh();
}

function clearDDL() {
  if(_formType == "New") {
    fd.field("InspType").value = null;

    if(fd.field("InspType").value === null){
      fd.field("Discipline").filter = "InspType/Id eq 0";
      fd.field("Discipline").refresh();
    }

    fd.field("Discipline").value = null;
    fd.field("InspPurpose").value = null;
    fd.field("InspDesc").value = null;
    fd.field("Submit").value = false;
  }

  fd.field("InspPurpose").filter = "Discipline/Id eq 0";
  fd.field("InspPurpose").refresh();

  fd.field("InspDesc").filter = "Discipline/Id eq 0";
  fd.field("InspDesc").refresh();
}

function showHideIRfields(InspType) {
  var InspVal = InspType.LookupValue;
  if (InspVal == "Buildings") {
    $(fd.field("BuildingType").$parent.$el).show();
    fd.field("BuildingType").required = true;
    $(fd.field("Level").$parent.$el).show();
    fd.field("Level").required = true;

    $(fd.field("Room").$parent.$el).show();

    $(fd.field("Road").$parent.$el).hide();
    fd.field("Road").required = false;
    $(fd.field("FromStation").$parent.$el).hide();
    fd.field("FromStation").required = false;
    $(fd.field("ToStation").$parent.$el).hide();
    fd.field("ToStation").required = false;
    $(fd.field("Zone").$parent.$el).hide();
    fd.field("Zone").required = false;
  } else {
    $(fd.field("BuildingType").$parent.$el).hide();
    fd.field("BuildingType").required = false;
    $(fd.field("Level").$parent.$el).hide();
    fd.field("Level").required = false;
    $(fd.field("Room").$parent.$el).hide();

    $(fd.field("Road").$parent.$el).show();
    fd.field("Road").required = true;
    $(fd.field("FromStation").$parent.$el).show();
    fd.field("FromStation").required = true;
    $(fd.field("ToStation").$parent.$el).show();
    fd.field("ToStation").required = true;
    $(fd.field("Zone").$parent.$el).show();
    fd.field("Zone").required = true;
  }
}

function disableIRFields(formType) {
  var InspVal = fd.field("InspType").value.LookupValue;
  var isStatus = fd.field("Status").value;

  if (formType != "Display") {
    if (isStatus != "Initiated") {
      fd.field("Title").disabled = true;
      fd.field("InspType").disabled = true;
      fd.field("Discipline").disabled = true;
      fd.field("InspPurpose").disabled = true;
      fd.field("InspDesc").disabled = true;
      fd.field("Remarks").disabled = true;

      fd.field("HardcopyDate").disabled = true;
      fd.field("Inspection_x0020_Date").disabled = true;
      fd.field("Drawingno").disabled = true;
      fd.field("Submit").disabled = true;
      fd.field("Attachments").disabled = true;
      //fd.toolbar.buttons[3].disabled = true;
      //fd.toolbar.buttons[4].disabled = true;
    } else {
      fd.field("Title").disabled = false;
      fd.field("InspType").disabled = false;
      fd.field("Discipline").disabled = false;
      fd.field("InspPurpose").disabled = false;
      fd.field("InspDesc").disabled = false;
      fd.field("Remarks").disabled = false;

      fd.field("HardcopyDate").disabled = false;
      fd.field("Inspection_x0020_Date").disabled = false;
      fd.field("Drawingno").disabled = false;
      fd.field("Submit").disabled = false;
      fd.field("Attachments").disabled = false;
    }
  }

  if (InspVal == "Buildings") {
    $(fd.field("BuildingType").$parent.$el).show();
    fd.field("BuildingType").required = true;
    $(fd.field("Level").$parent.$el).show();
    fd.field("Level").required = true;
    $(fd.field("Room").$parent.$el).show();

    $(fd.field("Road").$parent.$el).hide();
    fd.field("Road").required = false;
    $(fd.field("FromStation").$parent.$el).hide();
    fd.field("FromStation").required = false;
    $(fd.field("ToStation").$parent.$el).hide();
    fd.field("ToStation").required = false;
    $(fd.field("Zone").$parent.$el).hide();
    fd.field("Zone").required = false;

    if (formType != "Display") {
      if (isStatus != "Initiated") {
        fd.field("BuildingType").disabled = true;
        fd.field("Level").disabled = true;
        fd.field("Room").disabled = true;
      } else {
        fd.field("BuildingType").disabled = false;
        fd.field("Level").disabled = false;
        fd.field("Room").disabled = false;
      }
    }
  } else {
    $(fd.field("Road").$parent.$el).show();
    fd.field("Road").required = true;
    $(fd.field("FromStation").$parent.$el).show();
    fd.field("FromStation").required = true;
    $(fd.field("ToStation").$parent.$el).show();
    fd.field("ToStation").required = true;
    $(fd.field("Zone").$parent.$el).show();
    fd.field("Zone").required = true;

    $(fd.field("BuildingType").$parent.$el).hide();
    fd.field("BuildingType").required = false;
    $(fd.field("Level").$parent.$el).hide();
    fd.field("Level").required = false;
    $(fd.field("Room").$parent.$el).hide();

    if (formType != "Display") {
      if (isStatus != "Initiated") {
        fd.field("Road").disabled = true;
        fd.field("FromStation").disabled = true;
        fd.field("ToStation").disabled = true;
        fd.field("Zone").disabled = true;
      } else {
        fd.field("Road").disabled = false;
        fd.field("FromStation").disabled = false;
        fd.field("ToStation").disabled = false;
        fd.field("Zone").disabled = false;
      }
    }
  }
}
//#endregion

//#region MIR FUNCTIONS
function setMATMetaInfo(formType) {
  fd.field("MaterialSubmittalNo").filter =
    "Code eq 'Approved' or Code eq 'Approved as noted'";
  //fd.field('MaterialSubmittalNo').orderBy = { field: 'MaterialSubmittalNo', desc: false };
  fd.field("MaterialSubmittalNo").refresh();

  fd.field("MaterialSubmittalNo").$on("change", function (value) {
    if (fd.field("MaterialSubmittalNo").value != null) {
      var MaterialSubmittalNo = fd.field("MaterialSubmittalNo").value.LookupId;
      pnp.sp.web.lists
        .getByTitle("Material Submittal")
        .items.getById(MaterialSubmittalNo) //filter(_query)//
        .select("Title", "HardcopyDate", "BOQ", "Specs/Title", "Specs/ID","Manufacturer")
        .expand("Specs")
        //.filter("Code eq 'Approved'")
        .get()
        .then(function (foundItem) {
          if (foundItem) {
            if (foundItem.Title != null)
              fd.field("MATReferenceTitle").value = foundItem.Title;
            if (foundItem.HardcopyDate != null)
              fd.field("MATReferenceRecDate").value = new Date(
                foundItem.HardcopyDate
              );
            if (foundItem.BOQ != null)
              fd.field("MATReferenceBOQ").value = foundItem.BOQ;

            var specs = "";
            if (foundItem.Specs != null && foundItem.Specs != "") {
              for (var i = 0; i < foundItem.Specs.length; i++) {
                specs += foundItem.Specs[i].Title + ", ";
              }
              if (specs != null && specs != "") {
                specs = specs.slice(0, -2);
                fd.field("MATReferenceSpecs").value = specs;
              }
            }

            if (foundItem.Manufacturer != null)
              fd.field("Manufacturer").value = foundItem.Manufacturer;

            const fields = ["MATReferenceTitle", "MATReferenceRecDate", "MATReferenceBOQ", "MATReferenceSpecs", "Manufacturer",];
            DisableFields("Status", "Initiated", "eq", fields, true, false, true); 
            //DisableFields(filterColumn, filterValue, operator, fields, disableControls, disableCustomButtons, defaultButtons)
          }
        });
    }
  });
  fd.field("Discipline").orderBy = { field: "Title", desc: false };
  fd.field("Discipline").refresh();
}
//#endregion

//#region MAT FUNCTIONS
function setMATCascadedValues(formType) {
	clearMATDDL(formType);

  fd.field("Discipline").ready()
    .then(function () {
      fd.field("Discipline").orderBy = { field: "Title", desc: false };
      fd.field("Discipline").refresh();
      fd.field("Discipline").$on("change", function (value) {
        filterMATDiscipline(value);
        fd.field("SubDiscipline").value = null;
      });
    });
}

function clearMATDDL(formType){
  if (formType == "New") {
    fd.field("Discipline").orderBy = { field: "Title", desc: false };
    fd.field("Discipline").refresh();
    fd.field("SubDiscipline").filter = "Discipline/Id eq " + null;
    fd.field("SubDiscipline").refresh();
  }
}

function filterMATDiscipline(discipline) {
  var disciplineId = (discipline && discipline.LookupId) || discipline || null;
  fd.field("SubDiscipline").filter = "Discipline/Id eq " + disciplineId;
  fd.field("SubDiscipline").orderBy = { field: "Title", desc: false };
  fd.field("SubDiscipline").refresh();
}
//#endregion

//#region SLF FUNCTIONS
async function setSLFGrid(showGrid, removeRef) {

  if(!showGrid){
   $('#dt').hide();

   if(removeRef)
    fd.field("Reference").value = '';
  }
  else{
    var reference = fd.field("Reference").value;

    if(_formType === 'New'){
      reference = await setReference(_module);
      fd.field("Reference").value = reference;
    }

    $('#dt').show();

    _spComponentLoader.loadScript(_layout + '/plumsail/js/grid/grid.js').then(async () =>{
      var objValue = await buildGrid();
      if(objValue !== undefined){
        element = objValue.element;
        targetList = objValue.targetList;
        targetFilter = objValue.targetFilter;
      }
    }); 
  }
}

async function setReference(moduleName){
  var _web = pnp.sp.web;
  var referenceFormat = await getSLF_ReferenceFormat_MajorType(_web, moduleName);

  var returnValue ='';
  var disAcronym = '';
  var splitReference = referenceFormat.split('-');
  for (var i = 0; i < splitReference.length; i++){
    var _column = splitReference[i].replace('[', '').replace(']', '');
    if(_column !== null && _column !== '' && _column !== undefined){
      if (_column.includes('"'))
        _column = _column.replaceAll('"', '');
      else if (_column.includes("$")){
       _column = _column.replaceAll("$", "");
       var type = moduleName + '-' + disAcronym;
       var getCounter = await GetReferenceCounter(_web, type);
       _column = _column + getCounter;

       counterType = type;
       counter = getCounter;
      }
      else {
        if(_column === 'Discipline'){
         var value = fd.field(_column).value.LookupValue;
         _column = await getDisciplineAcronym(_web, value);
         disAcronym = _column;
        }
        else if( _column === 'BuildingType')
         _column = fd.field(_column).value.LookupValue;
        else _column = fd.field(_column).value;

      }
      returnValue += _column + "-";
    }
  }

  returnValue = returnValue.slice(0, -1);

 return returnValue;
}

async function getDisciplineAcronym(_web, value){
  var result = ''
  await _web.lists
  .getByTitle("Discipline")
  .items
  .select("Acronym")
  .filter("Title eq '" + value + "'")
  .get()
  .then(async function (items) {
      if(items.length > 0)
      result = items[0].Acronym;
      });
return result;
}

async function GetReferenceCounter(_web, type){
  var result = ''
  await _web.lists
  .getByTitle("Counter")
  .items
  .select("Counter")
  .filter("Title eq '" + type + "'")
  .get()
  .then(async function (items) {
      if(items.length > 0){
        result = parseInt(items[0].Counter);
      }
      else{
        result = 1;
      }
   });
  return result;
}

async function setCounter(type, counter){
  var _web = pnp.sp.web;
  var _listname = 'Counter';
  var objValue = {};

  await _web.lists
  .getByTitle(_listname)
  .items
  .select("Id,Counter")
  .filter("Title eq '" + type + "'")
  .get()
  .then(async function (items) {
      if(items.length === 0){
         objValue['Title'] = type;
         objValue['Counter'] = counter.toString();
        _web.lists.getByTitle(_listname).items.add(objValue);
      }
      else{
        var _item = items[0];
        counter = parseInt(counter);
        counter++;
        objValue['Counter'] = counter.toString();
       _web.lists.getByTitle(_listname).items.getById(_item.Id).update(objValue);
      }
   });
}

async function setLeadPartTrades(){
  let listName;
	const mType = await getMajorType(_module);

	if(mType !=null && mType.length > 0)
		listName = mType[0].MatrixList;

    const query = pnp.sp.web.lists.getByTitle(listName).items.select("Title");
						await query.orderBy("Title", true)
							.get()
							.then(async (items) => { 
									const TradeArray = [];
									const leadArray = [];
									for(let i = 0; i < items.length; i++)
									{
										const trade = items[i].Title;
										TradeArray.push(trade);
										leadArray.push(trade);
									}
									await setTrades(TradeArray, leadArray, _module, false);
							}); 	
}

async function getZoneCount(){
  return await pnp.sp.web.lists
  .getByTitle('Zone')
  .items
  .select("Id")
  .get()
  .then(async function (items) {
      return items.length;
   });
}
//#endregion

//#region DPR FUNCTIONS
var getTemplateItems = async function(targetList, colsInternal){
  var itemArray = [];
  
  const cols = colsInternal.join(',');
  var _query = "DeliverableType eq '" + _module + "'";

  var items = await pnp.sp.web.lists.getByTitle(targetList).items.filter(_query).select('Id,' + cols).getAll();

  for(var i = 0; i < items.length; i++){
    var item = items[i];
    var rowData  = {};

    for(var j = 0; j < colsInternal.length; j++){
      var colname = colsInternal[j];
      rowData[colname] = item[colname];
    }
    itemArray.push(rowData);
  }
  return itemArray;
}

var insertItemsInBulk = async function(itemsToInsert, targetList, tblName, colsInternal) {
  const list = pnp.sp.web.lists.getByTitle(targetList);
  const batch = pnp.sp.createBatch();
  const columns = colsInternal.join(',');

  for (const item of itemsToInsert){
    var _query = '';
    
    if(tblName === '')
      _query = `Title eq '${item.Title}' and Lookup_ID eq null`;
    else _query = `Title eq '${item.Title}' and DeliverableType eq '${_module}' and TableName eq '${tblName}'`;

    const existingItem = await list.items
    .select("Id," + columns)
    .filter(_query)
    .top(1)
    .get();

    if (existingItem.length > 0) {
      var _item = existingItem[0]; // _item is the oldItem values while objValue is the new one
      var doUpdate = false;

      for(var i = 0; i < colsInternal.length; i++){
        var columnName = colsInternal[i];
        var oldValue = '', currentValue = '';

        oldValue = _item[columnName];
        currentValue = item[columnName];

        if( oldValue != currentValue ){
            doUpdate = true;
            break;
        }
      }
      if(doUpdate)
       list.items.getById(existingItem[0].Id).inBatch(batch).update(item);
    } else {
       list.items.inBatch(batch).add(item);
    }
  }
  await batch.execute();
}

var setData = async function(){ 
  const listname = 'PR Template';
  var templateCols = ['Title', 'No','Hrs', 'TableName','DeliverableType'];
  var colsInternal = ['Title', 'No','Hrs'];

  var result = await getTemplateItems(listname, templateCols);
  
  if(result.length > 0){
    var staff = [], equipment = [];
    result.filter(item => {
      if(item.TableName === 'Staff')
        staff.push(item);
      else if(item.TableName === 'Equipment')
        equipment.push(item);
    });

    staff.forEach(obj => {
        delete obj.TableName; 
        delete obj.DeliverableType;
    });
    await insertItemsInBulk(staff, 'On Site Staff and Labour', '', colsInternal);
    fd.control('OnSiteStaffLabour').refresh();

    equipment.forEach(obj => {
      delete obj.TableName; 
      delete obj.DeliverableType;
    });
    await insertItemsInBulk(equipment, 'On Site Plant & Equipment Record', '', colsInternal);
    fd.control('OnSitePlant').refresh();
  }
}
//#endregion

//#region GENERAL FUNCTIONS
const setDimonButton = async function(text, isDisabled){
  if(isDisabled)
    $('span').filter(function () { return $(this).text() == text; }).parent().css('color', '#737373').attr("disabled", "disabled");
  else $('span').filter(function () { return $(this).text() == text; }).parent().css('color', '#444').removeAttr('disabled');
}

var setButtons = async function (moduleName, formType, Trans, Status, isLead, isPart) {
  if(!_isSiteAdmin && (Status == "Closed" || Status === 'Issued to Contractor' || Status === 'Completed')){
    await customButtons("ChromeClose", "Cancel", false);
    return;
  }

  var isAttachmentRequired = true;
  if(moduleName != 'MIR' && Status !== "Completed" && isPart && (_isTeamLeader || _IsLeadAction || _partTrade == 'PMC') )
    await customButtons("Reply", "Reject", false, Trans);
 
   if (_isLead || _isPMC){ // && (moduleName === 'IR' || moduleName === 'MIR' || moduleName === 'SLF') ) {
      var Trade = fd.field("Trade").value;

     if(isLead && (_isAutoAssign && !_AllowManualAssignment) ){}
     else{
        if(_isAutoAssign && !_isPMC)
        {
            $("ul.k-reset")
              .find("li")
              .first()
              .css("pointer-events", "none")
              .css("opacity", "0.6");

            $("div.k-list-scroller ul:eq(1)")
              .find("li:contains(" + Trade + ")")
              .css("pointer-events", "none")
              .css("opacity", "0.6");
        
            fd.field("Part").$on("change", function (value) {
              $("ul.k-reset")
                .find("li")
                .first()
                .css("pointer-events", "none")
                .css("opacity", "0.6");
            });
        }
        else{
          var Lead = fd.field("Lead").value;
          // if(Lead == null || Lead == "")
          //   fd.field('Part').disabled = true;
          //else 
          await renderTradeResult(Lead, fd.field("Part").value);

          preventDuplicateTradeSelection();
        }
        //Disable PMC Trade
        if(moduleName === 'IR' || moduleName === 'MIR' ||  moduleName === 'SLF')
        {
          $('div.k-list-scroller ul').find('li').each(function(index){
            var element = $(this);
            var _value = $(element).text().trim();
        
            if(_value === 'PMC'){
              $(element).css("pointer-events", "none").css("opacity", "0.6");
              $(element).prop('disabled', true);
            }
          });
        }
    }
   }
   else if(_isMain && _module === 'SLF'){
    $('div.k-list-scroller ul').find('li').each(function(index){
      var element = $(this);
      var _value = $(element).text().trim();
  
      if(_value === 'PMC'){
        $(element).css("pointer-events", "none").css("opacity", "0.6");
        $(element).prop('disabled', true);
      }
    });
   }
 
   if(moduleName == "SLFI" && formType == "Edit"){
     fd.toolbar.buttons[0].style = "display: none;";
     fd.toolbar.buttons[1].style = "display: none;";
     
     if(Status !== 'Closed')
       await customButtons("Accept", 'Submit', true, Trans, false, true, false, false, moduleName);
 
     await customButtons("ChromeClose", "Cancel", false);
     return;
   }
   else if(moduleName == "SLF" && formType == "Edit" && isContractor && Status === 'Issued to Contractor'){
     var ReviewedURL = fd.field("ReviewedURL").value;
     if(ReviewedURL === null || ReviewedURL === undefined)
      await customButtons("Accept", 'Submit', true, Trans, false, true, false, false, moduleName);
   }
 
   if (_isMain && formType == "Edit" &&(Status == "Closed" || Status == "Issued to Contractor"))
     await customButtons("ChromeClose", "Cancel", false);
   else {
     if (isAllowed &&(Status == "Open" || Status == "Assigned" ||Status == "Reassigned" || Status == "Returned to Site")) {
       if (moduleName != "SI") {
         if (isPart) {
           await customButtons("Accept", "Submit", true, Trans, false, true, false, true, moduleName);
           if (_partTrade == 'PMC')
             await customButtons("Assign", "Assign", true, Trans, false, false, false, true, moduleName);
         } else if (isLead) {
             if ( (_AllowManualAssignment || (!_isAutoAssign && !_AllowManualAssignment)) && (Status == "Open" || Status == "Assigned" || Status == "Reassigned"))
              await customButtons("Assign", "Assign", false, Trans, false, false, false, false, moduleName);
             else $("ul.nav,nav-tabs").remove();
         }
       }
     } 
     else {
       if (formType == "Display") {
       } else {
         if (isLead && _AllowManualAssignment) {
             await customButtons("Accept", "Assign", true, Trans, false, true, false, true);
         } else {
             if (isPart && Status !== "Completed") {
 
               if(moduleName == "SLF" && isPart && _isTeamLeader){}
               else await customButtons("Save", "Save", true, Trans);
 
               await customButtons("Accept", "Submit", true, Trans, false, true, false, true, moduleName);
             } 
             else if (Status == "Initiated" || Status == "Pending" || Status == "Open" || Status == "Assigned" || Status == "Reassigned"){
                 if(moduleName !== "SLF" && Status === "Open"){
                  if(_module == 'DPR' && Status == "Open"){  // FOR EDIT ON DAILY REPORT
                    if(isDar){
                      await customButtons("Save", "Preview Form", false);
                      await customButtons("Accept", "acknowledge", false, Trans, false, true, false, false, moduleName);
                    }
                    await customButtons("ChromeClose", "Cancel", false);
                    return;
                  }
                   else if(!_isDigitalForm && isDar)
                    await customButtons("Save", "Save", true, Trans);
                 }
                 else{
                   if( (Status == "Initiated" || Status == "Pending") && (!isContractor || ( (formType == "New" || formType == "Edit") )) )
                   {
                    if (moduleName == "SLF" || moduleName == "SCR" || moduleName == "DPR")
                      isAttachmentRequired = false;

                      if(moduleName !== "DPR")
                        await customButtons("Save", "Save", true, Trans);
                      
                      let runBelow = true;
                      if(moduleName === "SI" && formType == "Edit" )
                        runBelow = false;

                      if(runBelow)
                        await customButtons("Accept", "Submit", true, Trans, isAttachmentRequired, true, true, false, moduleName); // FOR SCR
                   }
                   else if(formType == "Edit" && isDar && !_isDigitalForm && Status !== "Initiated"){
                    await customButtons("Save", "Save", true, Trans);
                   }
                   
                   
                   var text = 'Submit';
 
                   if (moduleName == "SLF"){
                     isAttachmentRequired = false;
                     if(formType == "Edit"){
                       text = 'Assign';
                       if(isContractor)
                         text = 'Submit'; 
                     }
                   }
               }
             }
           }
         }
       }
 
     if (isLead && fd.field("Status").value != "Completed")
       await customButtons("Accept", "Compile & Close", false);
 
     if (formType == "Display") {
     } 
     else {
      if(moduleName === "SI" && Status === 'Open'){}
      else await customButtons("ChromeClose", "Cancel", false);
     }
   }



};

function delay(time) {
  return new Promise(function(resolve) { 
      setTimeout(resolve, time)
  });
}

function setToolTipMessages(){
  setButtonToolTip('Save', saveMesg);
  setButtonToolTip('Submit', submitMesg);
  setButtonToolTip('Preview Form', previewMesg);
  setButtonToolTip('Reject', rejectMesg);
  setButtonToolTip('Assign', assignMesg);
  setButtonToolTip('Cancel', cancelMesg);
  setButtonToolTip('Save Attachment', 'Displayed for admin only.');

  //else addLegend();
  adjustlblText('Comment', ' (Optional)', false);
  	
	if($('p').find('small').length > 0)
    $('p').find('small').remove();
}

function removePadding(){
  $(".fd-form-block > .fd-grid[data-v-105ebe50]").attr('style', 'padding: 12px !important');

  var mainContent = $('#spPageChromeAppDiv div').next();
  var targetElement = mainContent[2];
  //targetElement.style.marginLeft = '-200px';
}

function removeImage(){
  if($('#loader').length > 0)
   $('#loader').remove();
   clearInterval(intervalId);
   
}

function hideReeasonofRejection(){
  var elem = $("textarea")[1];
  var _elemValue = elem._value
  
  if(_elemValue === '' || Status === 'Completed' || _isTeamLeader)
    $('div._vpw0q').hide();
  else{
    $(elem).prop("readonly", true);
  }
}

function handleHardCopyDate(fields, hfield){
  var _columnName = 'HCDate';
  if(_formType === 'Edit' && _isMain){
    
    if(!_isDigitalForm && Status !== 'Initiated' && isDar){

      var hardCopyDate = fd.field(_columnName).value;
      if (hardCopyDate !== null || Status === 'Issued to Contractor') 
        fd.field(_columnName).disabled = true;
      else fd.field(_columnName).required = true;
    }
    else {
      hfield.push(_columnName);
    }
  }
}

function setOldRef(_status){
  fd.field('ORFI').ready().then(() => {
   //fd.field('ORFI').filter = "Reference ne null";  
   fd.field('ORFI').filter = `Reference ne null  and Status eq '${  _status  }' and IsLatestRev eq '1'`;  
   fd.field('ORFI').orderBy = { field: 'Reference', desc: false };
   fd.field('ORFI').refresh();
   });

   if(_isCompliedWithConvention){
    fd.field('ORFI').$on('change', function(value) {

      if(_isCompliedWithConvention === true)
         setNamingConvention('Reference');

      if(value !== null){
        fd.field('Reference').value = value.LookupValue;
        fd.field('Reference').disabled = true;
      }
      else{
        fd.field('Reference').value = '';
        fd.field('Reference').disabled = false;

        if($('#fncImgId').length > 0)
           $('#fncImgId').remove();
      }
    });
   }
}

var setNamingConvention = async function(referenceField){
  let delimeter, schema, scheamResult, filenameText = '', schemaFields = '', listname = '';
  let schemFieldsLength = 0;
  var refId = $("input[title='Reference']");
  refId.css('text-transform', 'uppercase');

  fd.field('Reference').required = true;
  fd.field('Reference').clear();

    await pnp.sp.web.lists
			.getByTitle("FNC")
			.items
			.select("Delimeter,Schema")
			.filter(`Title eq '${  _module  }'`)
			.get()
			.then(async (items) => {
          if(items.length > 0)
          {
            delimeter = items[0].Delimeter;
            schema = items[0].Schema;

            schema = schema.replace(/&nbsp;/g, '');
            schema = JSON.parse(schema);

            await renderFieldsOnForm(schema);
            scheamResult = await setFilenameText(schema, delimeter);
            
             if(scheamResult.filenameText !== ''){
              schemaFields = scheamResult.schemaFields.slice(0, -1);
              filenameText = scheamResult.filenameText.slice(0, -1);
              listname = scheamResult.listname;

              fd.field('Reference').placeholder = filenameText
                                                 .replace('Discipline_x003a_Acronym','Discipline Acronym')
                                                 .replace('SubDiscipline_x003a_Title','SubDiscipline Acronym');

              if(refId.length === 0){
                refId = $(`input[title='${filenameText}']`);
                var filename = fd.field('Reference').value;
                await checkFileName(_module, delimeter, schema, filenameText, filename, true, false);
              }
              else $("input[title='Reference']").attr('title', filenameText);
            }

            $(refId).on('change', async function(value) {
              //if(fd.field('ORFI').value === undefined){
                var filename = this.value;

                if(filenameText === '' || filenameText === undefined){
                  scheamResult = await setFilenameText(schema, delimeter);
                  if(scheamResult.filenameText !== ''){
                    schemaFields = scheamResult.schemaFields.slice(0, -1);
                    filenameText = scheamResult.filenameText.slice(0, -1);
                  }
                }

                var result = await checkFileName(_module, delimeter, schema, filenameText, filename, true, false); //function checkFileName(acronym, delimeter, schema, filename, isSingle, checkRev)
                if(result !==  ''){
                  disableButtonsFNC = true;
                 await setErrorMesg(refId, result, null, true);
                }
                else{
                  if(_module !== 'DPR')
                    result = await isFilenameExist(listname, filename);

                  if(result !==  ''){
                    disableButtonsFNC = true;
                   await setErrorMesg(refId, result, null, true);
                  }
                  else {
                    disableButtonsFNC = false;
                    await setErrorMesg(refId, '', null, true);
                  }
                }
              //}
            });
          }
          else {
            disableButtonsFNC = true;
            await setErrorMesg(refId, 'Naming convention is not defined, contact your administrator.', null, true);
          }
		});
}

var isFilenameExist = async function(listname, filename){
  let mesg = '';
  var _query = "Reference eq '" + filename + "' and IsLatestRev eq '1'";

   const list = pnp.sp.web.lists.getByTitle(listname);
   const items = await list.items.select("Id, Rev, Status").filter(_query).top(1).get();
   if (items.length > 0){
     let item = items[0];
     let rev = (item['Rev'] !== undefined && item['Rev'] !== null) ? item['Rev'] : '';
     if(rev !== ''){
      let status = (item['Status'] !== undefined && item['Status'] !== null) ? item['Status'] : '';
      if(status !== "Issued to Contractor" && status !== "Closed")
        mesg = "Can't submit new revision as previous one is under review.";
      else{
        var obj = {  LookupId: item['Id'],
                     LookupValue: filename
                  };

        $(fd.field('ORFI').$parent.$el).show();
        fd.field('ORFI').value = obj;
        fd.field('ORFI').disabled = true;
      }
     }
     else mesg = "filename is already exist.";
   }
   else $(fd.field('ORFI').$parent.$el).hide();
   return mesg;
}

var setErrorMesg = async function(inputElement, mesg, elementErrorId, isFNC){
  // var webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
  // var ImageUrl =  webUrl + _layout + '/Images/Error.png';
  // var imageId = '#fncImgId';
  // if(isCorrect)
  //  ImageUrl = webUrl + _layout + '/Images/Submitted.png';

  // if($(imageId).length > 0){
  //   $(imageId).attr("src", ImageUrl);
  // }
  // else{
  //   $("<img id='fncImgId' src='" + ImageUrl + "' />")
  //   .addClass("imageValidator")
  //   .insertAfter(inputElement);
  // }

  var errorId = '#cMesg';
  if(_module !== 'FNC')
    errorId = '#' + elementErrorId;

  if($(errorId).length === 0){
    var htmlContent = "<div id='" + errorId.replace('#','') + "' class='form-text text-danger small'>" + mesg + "</div>";
    $(inputElement).after(htmlContent);
  }
   
  if(disableButtonsFNC || disableButtons){
      $(errorId).html(mesg).attr('style', 'color: rgba(var(--fd-danger-rgb), var(--fd-text-opacity)) !important');

    // $(imageId).tooltipster({
    //   content: mesg,
    //   contentAsHTML: true,
    //   delay: 100,
    //   maxWidth: 350,
    //   speed: 500,
    //   interactive: true,
    //   animation: 'slide', //fade, grow, swing, slide, fall
    //   trigger: 'hover'
    // });

    $('span').filter(function(){ return $(this).text() == 'Save'; }).parent().attr("disabled", "disabled");
		$('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().attr("disabled", "disabled");
  }
  else{
    if(_module === 'FNC' || isFNC)
      $(errorId).html(fncSuccessMesg).attr('style', 'color: green !important');
    else $(errorId).remove();

    $('span').filter(function(){ return $(this).text() == 'Save'; }).parent().removeAttr('disabled');
		$('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().removeAttr('disabled');
    //$('#cMesg').remove();

    //$(imageId).tooltipster('destroy');
    //$(imageId).removeAttr("title");
  }
}

var delayshowHideFunction = async function(){
  var isValid = false;
  var retry = 1;
  while (!isValid)
  {
    try{
      if(retry >= retryTime) break;
      isAllowed = await showHideTabs(_module, BIC, false, _isLead);
      isValid = true
      return isAllowed;
    }
    catch{
      retry++;
      await delay(delayTime);
    }
  }
}

var setFilenameText = async function(schema, delimeter){
  var filenameText = '', schemaFields = '', listname = '';
  schema.filter(item => {
    var fieldName = item.InternalName;
    if(fieldName === 'Rev' || fieldName === 'Revision'){
      schemaFields += fieldName + delimeter;
      return;
    }
      
    if(fieldName !== 'properties'){
      schemaFields += fieldName + delimeter;
      filenameText += fieldName + delimeter;

       //try {
       // fd.field(fieldName).value;
        //fd.field(fieldName).disabled = true;
        //fd.field(fieldName).required = false;
        //$(fd.field(fieldName).$parent.$el).hide();
      //}
      //catch(e){}
    }
    else listname = item.MainList;
   });
   return {
    filenameText: filenameText,
    schemaFields: schemaFields,
    listname: listname
  };
}

var validateDisciplineAgainstMatrix = async function(){
	var FieldsArray = [];
	if(mTypeItem !== undefined ){
	  var AutoAssignQueryColumns = mTypeItem.AutoAssignQueryColumns;
  
	  if(AutoAssignQueryColumns !== undefined && AutoAssignQueryColumns !== null){
		  if(AutoAssignQueryColumns.includes(","))
		  {
			AutoAssignQueryColumns = AutoAssignQueryColumns.split(',');
			for(let i = 0; i < AutoAssignQueryColumns.length; i++)
			{
			  FieldsArray.push(AutoAssignQueryColumns[i]);
			}
		  }
		  else FieldsArray.push(AutoAssignQueryColumns);
	  
		  var MatrixList = mTypeItem.MatrixList;
		  if(MatrixList !== undefined && FieldsArray.length > 0){
		  for(let j = 0; j < FieldsArray.length; j++){
			  let fieldname = FieldsArray[j];
			  fd.field(fieldname).$on('change', async function(value) {
        var elementErrorId = `${fieldname}Id`;
        if(value === null){
          $('#' + errorId).remove();
          return;
        }

        let result;
				await pnp.sp.web.lists
				.getByTitle(MatrixList)
				.items
				//.select(AutoAssignQueryColumns)
				.filter(`${fieldname} eq '${  value.LookupValue  }'`)
				.get()
				.then((items) => {
				  if(items.length > 0)
					  result = items[0];
				});
  
          var fieldnameText = fieldname;

          if(fieldname === 'BuildingType')
           fieldnameText = 'Building Name';

          var spanWithText = $("span:contains('" + fieldnameText + "'):first");
          var getLabelElement = spanWithText[0].parentElement.parentElement.children[1];       
          var inputElement= getLabelElement.children[0].children[0];
          
          //!_isCompliedWithConvention && 
          if (result === undefined){
            disableButtons = true;
            await setErrorMesg(inputElement, `Kindly contact administrator to add ${value.LookupValue} to the matrix to proceed`, elementErrorId);
          }
          else {
            disableButtons = false;
            await setErrorMesg(inputElement, '', elementErrorId);
          }
			  });
			}
		 }
	  }  
	}
}

var preventDuplicateTradeSelection = async function(){
  fd.field("Lead").$on("change", function (value) {
    //fd.field('Part').disabled = false;
    $('div.k-list-scroller ul:eq(1)').find('li').each(function(index){
      var element = $(this);
      var _value = $(element).text().trim();
  
      if(_value === value  || ( _value === 'PMC' && (_module === 'IR' || _module === 'MIR' || _module === 'SLF') )){
        $(element).css("pointer-events", "none").css("opacity", "0.6");
        $(element).prop('disabled', true);
      }
      else{ 
          $(element).css("pointer-events", "auto").css("opacity", "1");
          $(element).prop('disabled', false);
      }
    });
  });

  fd.field("Part").$on("change", function (value) {
    $('div.k-list-scroller ul:eq(0)').find('li').each(function(index){
      var element = $(this);
      var _value = $(element).text().trim();
  
      if( (value.includes(_value)) || ( _value === 'PMC' && (_module === 'IR' || _module === 'MIR' || _module === 'SLF') )){
        $(element).css("pointer-events", "none").css("opacity", "0.6");
        $(element).prop('disabled', true);
      }
      else {
          $(element).css("pointer-events", "auto").css("opacity", "1");
          $(element).prop('disabled', false);
      }
    });
  });

 
}

var renderTradeResult = async function(leadValue, partValue){

  if(leadValue !== null && leadValue !== undefined){
    $('div.k-list-scroller ul:eq(1)').find('li').each(function(index){
      var element = $(this);
      var _value = $(element).text().trim();

      if(_value === leadValue[0]  || ( _value === 'PMC' && (_module === 'IR' || _module === 'MIR' || _module === 'SLF') )){
        $(element).css("pointer-events", "none").css("opacity", "0.6");
        $(element).prop('disabled', true);
      }
      else{
          $(element).css("pointer-events", "auto").css("opacity", "1");
          $(element).prop('disabled', false);
      }
    });
}

  if(partValue !== null && partValue !== undefined){
    $('div.k-list-scroller ul:eq(0)').find('li').each(function(index){
      var element = $(this);
      var _value = $(element).text().trim();

      if( (partValue.includes(_value)) || ( _value === 'PMC' && (_module === 'IR' || _module === 'MIR' || _module === 'SLF') )){
        $(element).css("pointer-events", "none").css("opacity", "0.6");
        $(element).prop('disabled', true);
      }
      else {
          $(element).css("pointer-events", "auto").css("opacity", "1");
          $(element).prop('disabled', false);
      }
    });
  }
}

var loadScripts = async function(){
  const libraryUrls = [
    _layout + '/controls/handsonTable/libs/handsontable.full.min.js',
    _layout + '/controls/preloader/jquery.dim-background.min.js',
    _layout + "/plumsail/js/buttons.js",
    _layout + '/plumsail/js/customMessages.js',
    _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
    _layout + '/plumsail/js/preloader.js',
    _layout + '/plumsail/js/commonUtils.js',
    _layout + '/plumsail/js/utilities.js',
    _layout + '/plumsail/js/partTable.js',
    _layout + '/plumsail/js/grid/gridUtils.js',
    _layout + '/plumsail/js/grid/grid.js'
  ];

  const cacheBusting = '?t=' + new Date().getTime();
    libraryUrls.map(url => { 
        $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
      });
      
  const stylesheetUrls = [
    _layout + '/controls/tooltipster/tooltipster.css',
    _layout + '/controls/handsonTable/libs/handsontable.full.min.css',
    _layout + '/plumsail/css/CssStyle.css',
    _layout + '/plumsail/css/partTable.css'
  ];

  stylesheetUrls.map((item) => {
    var stylesheet = item;
    $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
  });
}

var LimitDateTimeSubmission = async function(){

  var EnableWorkHours = await getParameter("EnableWorkHours"); 

  if (EnableWorkHours.toLowerCase() === "yes") {

    var WorkWeek = await getParameter("workWeek");
    var WorkStartTime = await getParameter("WorkStartTime");
    var WorkEndTime = await getParameter("WorkEndTime");

    const workWeek = WorkWeek.split(',');
    const workday = new Date().toLocaleDateString('en-US', { weekday: 'long' });

    const mesg = ''
    if (workWeek.includes(workday)) {
          const arrayWorkStartTime = WorkStartTime.split(',');
          const arrayWorkEndTime = WorkEndTime.split(',');

          if (arrayWorkStartTime.length > 0 && arrayWorkEndTime.length > 0) {
              var workStartTime = new Date();
              workStartTime.setHours(arrayWorkStartTime[0], arrayWorkStartTime[1], arrayWorkStartTime[2], 0);
              var workEndTime = new Date();
              workEndTime.setHours(arrayWorkEndTime[0], arrayWorkEndTime[1], arrayWorkEndTime[2], 0);
              const now = new Date();              

              if (now > workStartTime && now < workEndTime) { /* Submit Normally*/ }
              else {
                alert('Please note that you are not allowed to submit after working hours');      
                fd.close();
              }          
          }
    } 
    else {
      alert('Please note that you are not allowed to submit on Weekend.');        
      fd.close();
    }
    
  }
}
//#endregion

var getAutReferenceFormat = async function(){
  let schema = [];
  await pnp.sp.web.lists
  .getByTitle(_MajorTypes)
  .items
  .select("CDSFormat")
  .filter(`Title eq '${  _module  }'`)
  .get()
  .then((items) => {
    if(items.length > 0){
        let result = items[0].CDSFormat;
        let splitFields = result !== undefined && result !== null && result !== '' ? result.split('-') : [];
        if(splitFields.length > 0){
          let filteredFields = splitFields.filter(field => !field.includes('"') && !field.includes('$'));
          schema = filteredFields.map(field => ({
            InternalName: field.replace('[', '').replace(']', ''),
            isList: true
          }));
        }
    }
  });
  return schema;
}

var renderFieldsOnForm = async function(schema){
  var _fieldsToShow = '';
  
 if(_module === 'IR')
   _fieldsToShow = 'Area, Location, Package, Phase';
 else if(_module === 'MAT')
   _fieldsToShow = 'Building, Package, Phase, Location, Station, Zone, Road';
 else if(_module === 'MIR' || _module === 'SCR')
   _fieldsToShow = 'Area, Building, Level, Location, Package, Phase, Road, Station, Zone';
 else if(_module === 'SI')
   _fieldsToShow = 'Area, Building, Level, Package, Phase, Road, Station, Zone';

   if(_fieldsToShow !== ''){

    let fields = fd.spForm.fields;
    let splitMatchingFields = _fieldsToShow.trim().split(',');

      if(schema !== ''){
        for (const field in fields){

          let fieldProp = fd.field(field);
          let isFound = false;

          for (const item of schema){
            let fieldname = item.InternalName;
            let isList = item.isList !== undefined ? true : false;

            if(isList && field === fieldname){
              isFound = true
              fd.field(field).required = true;
              $(fieldProp.$parent.$el).show();
              break;
            }
          }
          if(!isFound){
              let item = splitMatchingFields.filter(fld =>{
                return fld.trim() === field.trim()
              })

              if(item.length > 0){
                fd.field(field).required = false;
                $(fieldProp.$parent.$el).hide();
              }
          }
        }
      }

    }


}