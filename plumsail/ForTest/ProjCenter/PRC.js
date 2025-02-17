var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isMain = false, _isLead = false, _isPart = false, _isNew = false, _isEdit = false, _isDisplay = false
_isPD = false, _isPM = false, _isQM = false, _isSus = false, _isGIS = false, _isGLMain = false, _isGL = false, _isTeamMember = false, _isQMOwner = false, _isSusOwner = false,
_isLLChecker = false, _isReader = false, _isConfidential = false, _isBuilding = false, _isReader = false, _isUserAllowed = false, _hideSubmit = false, _isPROLE = false;

var projectNo = '', projectTitle = '', CurrentUser;

var ProjectInfo = 'Project Info', Category = 'Category', SubCategory = 'SubCategory', MatrixFields = 'Matrix Fields', MTDs = 'MTDs', Firms = 'Firms', Questions = 'Questions', 
    ContractReview = 'Contract Review', OtherPartiesFirms = 'Other Parties Firms', Sustainability = 'Sustainability', Level = 'Levels', Roles = 'Roles', 
    ProjectRoles = 'Project Roles', GISLocation = 'GIS Location', RevDepartments = 'Review Departments', CompReports = 'Completion Reports', 
    TradeContractReview = 'Trade Contract Review', LLChecker = 'LLChecker';

var formFields = {};

var firstChildTab = 'TabsMain', secondChildTab = 'Tabs6'
var tabs = [
            {
              masterTab: firstChildTab,
              title: 'Sustainability',
              tooltip: 'Sustainability is enabled for Buildings Category only'
            },
            {
              masterTab: firstChildTab,
              title: 'Background Info',
              tooltip: 'Background Info is enabled for Buildings Category only'
            }
];

var appBarItems, isgetTradeRoleFinalized = false;

const blueColor = '#6ca9d5', greenColor = '#5FC9B3', redColor = '#F28B82'; // Buttons Colors

var onRender = async function (relativeLayoutPath, moduleName, formType) {
  _layout = relativeLayoutPath;  

    try{
      
      if(moduleName === 'PINT'){
        _isMain = true;
        addMessageEventListener();
      }

      await loadScripts().then(async ()=>{
        //preloader_btn(false, true);
        showPreloader();
        await extractValues(moduleName, formType);
      })
      
      if(_isMain)
       await onPINTRender();
      else if(_module === 'MTD')
        await onMTDRender();
      else if(_module === 'PROLE'){
        _isPROLE = true;
        await onProjRoleRender()
      }
      else if(_module === 'GIS'){
        _isGIS = true;
        await onGISRender();
      }
      else if(_module === 'DPIR')
        await onRevDeptTardeFormsRender()
      else if(_module === 'CR')
        await onCRTardeFormsRender()
      
      if(!_isDisplay){
        if(_module === 'CR'){
          const intervalId = setInterval(function() {
          
            if (isgetTradeRoleFinalized) {
              setCustomButtons();
              clearInterval(intervalId);
            }
        }, 100);
        }
        else await setCustomButtons()
      }

      //Remove_Pre(true);
    }
    catch (e){
      //preloader();
      showPreloader();
      fd.toolbar.buttons[0].style = "display: none;";
      console.log(e);
    }
    finally {      
      hidePreloader();
    }

    //setPreviewForm();
    //await drawChart();
    //await hideLicense();
}

//#region GENERAL
var loadScripts = async function(withSign){

    const libraryUrls = [
      //_layout + '/controls/goJs/go-debug.js',
      _layout + '/controls/preloader/jquery.dim-background.min.js',
      _layout + "/plumsail/js/buttons.js",
      _layout + '/plumsail/js/customMessages.js',
      _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
      _layout + '/plumsail/js/preloader.js',
      _layout + '/plumsail/js/commonUtils.js',
      _layout + '/plumsail/js/utilities.js',
      _layout + '/controls/vanillaSelectBox/vanillaSelectBox.js',

      _layout + '/plumsail/ForTest/ProjCenter/utils.js',
      _layout + '/plumsail/ForTest/ProjCenter/Dictionaries.js',
      _layout + '/plumsail/ForTest/ProjCenter/PIR/metrics.js',   
      _layout + '/plumsail/ForTest/ProjCenter/PIR/backgroundInfo.js',
      _layout + '/plumsail/ForTest/ProjCenter/PIR/designProcess.js',
      _layout + '/plumsail/ForTest/ProjCenter/PIR/Sustainability.js',
      _layout + '/plumsail/ForTest/ProjCenter/PIR/contractReview.js',

      _layout + '/plumsail/ForTest/ProjCenter/features/projroles.js',
      _layout + '/plumsail/ForTest/ProjCenter/features/gis.js',

      _layout + '/plumsail/ForTest/ProjCenter/DPIR/revDepartments.js',
      _layout + '/plumsail/ForTest/ProjCenter/DPIR/tradeForms.js',
      _layout + '/plumsail/ForTest/ProjCenter/DPIR/tradeContractReview.js',
      
      _layout + '/plumsail/ForTest/ProjCenter/CR/cr.js',
      _layout + '/plumsail/ForTest/ProjCenter/CR/crTradeForms.js'
    ];

    const cacheBusting = withSign ? `?v=${Date.now()}`: '';
      libraryUrls.map(url => {
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`);
        });

    const stylesheetUrls = [
      _layout + '/controls/tooltipster/tooltipster.css',
      _layout + '/controls/vanillaSelectBox/vanillaSelectBox.css',
      _layout + '/plumsail/ForTest/pmisStyle.css' + `?v=${Date.now()}`
    ];

    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}${cacheBusting}">`);
    });
}

var extractValues = async function(moduleName, formType){
  const startTime = performance.now();
  if($('.text-muted').length > 0)
    $('.text-muted').remove();

    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;

    if(_formType === 'New'){
      //clearStoragedFields(fd.spForm.fields);
      _isNew = true;
  }
  else if(_formType === 'Edit'){
      _isEdit = true;
      _itemId = fd.itemId;
  }
  else if(_formType === 'Display')
    _isDisplay = true;


  if(_isMain || _module === 'DPIR' || _module === 'CR'){
    if(_isMain)
      _isConfidential = fd.field('IsConfidential').value;

      await getCurrentUserRole()
      if (!_isUserAllowed){
        //if(!_isSiteAdmin){
          alert(isAllowedUserMesg);
          fd.close();
        //}
      }
  }
 

    if(_isMain){
      let gisUrl = `${_webUrl}/SitePages/PlumsailForms/GISLocation/Item/NewForm.aspx`
      let result = await isRecordExist(`MasterID/Id eq ${_itemId}`);
      if(result !== '')
        gisUrl = `${_webUrl}/SitePages/PlumsailForms/GISLocation/Item/EditForm.aspx?item=${result.Id}`
      appBarItems = [
        {
          svgPath: '<path d="m7.5 13a4.5 4.5 0 1 1 4.5-4.5 4.505 4.505 0 0 1 -4.5 4.5zm6.5 11h-13a1 1 0 0 1 -1-1v-.5a7.5 7.5 0 0 1 15 0v.5a1 1 0 0 1 -1 1zm3.5-15a4.5 4.5 0 1 1 4.5-4.5 4.505 4.505 0 0 1 -4.5 4.5zm-1.421 2.021a6.825 6.825 0 0 0 -4.67 2.831 9.537 9.537 0 0 1 4.914 5.148h6.677a1 1 0 0 0 1-1v-.038a7.008 7.008 0 0 0 -7.921-6.941z"/>',
          redirectUrl: `${_webUrl}/SitePages/PlumsailForms/ProjectRoles/Item/NewForm.aspx`,
          viewBox: '0 0 24 24',
          iconTitle: 'Roles',
          tooltip: 'Set user roles',
          editors: 'QM,SUS',
          readers: 'all'
        },
        {
          svgPath: '<path d="M12,0A10.011,10.011,0,0,0,2,10c0,5.282,8.4,12.533,9.354,13.343l.646.546.646-.546C13.6,22.533,22,15.282,22,10A10.011,10.011,0,0,0,12,0Zm0,15a5,5,0,1,1,5-5A5.006,5.006,0,0,1,12,15Z"/><circle cx="12" cy="10" r="3"/>',
          redirectUrl: gisUrl,
          viewBox: '0 0 24 24',
          iconTitle: 'GIS',
          tooltip: 'Set GIS Location',
          editors: 'PM,GIS',
          readers: 'PD,QM'
        },
        {
          svgPath: `
            <!-- Top Triangle -->
            <path d="M16.5,2.5L4,8.25L15,14L19.5,10.75L16.5,2.5Z" fill="#002050"/>
            
            <!-- Bottom Triangle -->
            <path d="M4,8.25L4,21.5L15,14L4,8.25Z" fill="#002050"/>

            <!-- Inner Triangle -->
            <path d="M7,11L4,21.5L15,14L7,11Z" fill="#ffffff"/>
          `,
          redirectUrl: 'https://ax.d365.dar.com/namespaces/AXSF/?mi=ProjProjectsListPage',
          viewBox: '0 0 24 24',
          iconTitle: 'D365',
          tooltip: 'Manage project on D365',
          editors: 'all',
          readers: 'all'
        }

      ];
      appBar();

      projectNo = fd.field('Title').value;
      projectTitle = fd.field('ProjectTitle').value;
      $(fd.field('ProjectTitle').$parent.$el).hide();

      let fullProjTitle = `${projectNo} - ${projectTitle}`;
      localStorage.setItem('projectNo', projectNo);
      localStorage.setItem('ProjectTitle', projectTitle);
      localStorage.setItem('FullProjTitle', fullProjTitle);

      fullProjTitle += ' - Project Management Initiation Form'
      setPageStyle(fullProjTitle);        
    }
    else projectNo = localStorage.getItem('projectNo');

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;

    await setFormHeaderTitle()

    let titleCtlr = $('div.dumTitle');
    let rowDiv = titleCtlr.parent().parent().parent().parent();
    if(rowDiv !== undefined)
       rowDiv.css('padding', '0px 12px 0px 12px');

    $('div.multiselect-dropdown').remove();

    //_spComponentLoader.loadScript(_layout + '/plumsail/ProjCenter/utils.js').then(async ()=> {
   
   //});
   CurrentUser = await GetCurrentUser()

   const endTime = performance.now();
   const elapsedTime = endTime - startTime;
   console.log(`extractValues: ${elapsedTime} milliseconds`);
}

var setCustomButtons = async function () {

    //  let ss = fd.field('Status').value;
    //  let longitude = fd.field('longitude').value;

    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";
    //fd.toolbar.buttons[2].style = "display: none;";

   
    if(!_hideSubmit){

      if(_module !== 'GIS' && ((_module === 'DPIR' || _module === 'CR') && _isGLMain) )
        await setButtonActions("Save", "Save", `${blueColor}`);

      if(!_isPROLE){
       await setButtonActions("Accept", submitText, `${greenColor}`);
      }
    }

    await setButtonActions("ChromeClose", "Close", `${redColor}`);
    setToolTipMessages();

    if(_isQM || _isSus || _isGL || _isTeamMember || _isReader){
        $('span').filter(function(){ return $(this).text() === 'Save'; }).parent().attr("disabled", "disabled");
        $('span').filter(function(){ return $(this).text() === submitText; }).parent().attr("disabled", "disabled");
    }
}

const setButtonActions = async function(icon, text, bgColor){

    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          style: `background-color: ${bgColor}; color: white;`,
          click: async function() {
            
            if(text == "Close" || text == "Cancel"){
              const lowerText = text.toLowerCase();
              if (confirm(`Are you sure you want to ${lowerText}?`)){
                  if(_module === 'GIS' || _module === 'PROLE' || _module === 'DPIR' || _module === 'CR'){
                    window.close();
                  }
                  else fd.close();
              }
            }
            else if(text == "Save"){
               fd.validators.length = 0;
               if(_isPROLE)
                window.close();
               else if(_module === 'DPIR'){
                  let {items} = await isTradeQuestionValid();
                  await insertTradeQuestions(items, TradeContractReview, ['Question', 'Answer', 'Comments','NeedPMAction', 'Department'])
                  fd.save().then(()=>{
                    window.close();
                  });
               }
               else{
                let {items} = await getTabFields();
                let itemMetaInfo = await getFieldsData(MatrixFields);
                let query = `Title eq '${projectNo}'`;

                await addUpdate(MatrixFields, query, itemMetaInfo);
                await insertItemsInBulk(items, ContractReview, ['Question', 'Answer', 'Comments']);

                let {dpirItems} = await isDPIR_RowValid()
                await updateDPIRItems(dpirItems, ['Status','RejRemarks']);
                fd.save();
               }
            }
            else if(text == submitText){
            
               if(_module === 'MTD'){
                  if(_isNew)
                    await setStatus('Pending');
                  else if(_isEdit)
                    await sendMTDApproval();

                  fd.save();
               }

               else if(_module === 'GIS'){
                   await setGISMetaInfo();
               }

               else if(_module === 'CR'){
                  await setCRMetaInfo();

                  let dept = fd.field("Title").value
                  await isAllTradesSubmitted(dept)

                  setTimeout(() => { 
                    fd.save()
                    .then(async ()=>{
                      setTimeout(() => {
                        window.close();
                      }, 100);          
                    })
                 }, 500);
               }

                else if(_module === 'DPIR'){
                  
                  if(fd.field('Attachments').value.length == 0){
                    setPSErrorMesg("Kindly upload a tentative list of drawings under Docs & Drawings List")
                    return;
                  }
                  
                  let {mesg,items} = await getTabFields('DPIRTabs');
                
                  if(mesg === ''){
                    if (confirm('Are you sure you want to Submit?')){
                      await insertTradeQuestions(items, TradeContractReview, ['Question', 'Answer', 'Comments','NeedPMAction', 'Department'])
                      await checkAllSubmitted();

                      fd.field('Status').value = 'Submitted'
                      fd.save()
                      .then(()=>{
                        window.close();
                      });
                    }
                  }
                  else setPSErrorMesg("Kindly fill required fields as shown in below exclamation mark icon on each Tab")
                }

                else{
                    let {mesg,items,dpirItems} = await getTabFields('Tabs1');

                    if(mesg === ''){
                      if (confirm('Are you sure you want to Submit?')){
                        let itemMetaInfo = await getFieldsData(MatrixFields);
                        let query = `Title eq '${projectNo}'`;
                        let {dpirItems} = await isDPIR_RowValid()

                        await Promise.all([
                           addUpdate(MatrixFields, query, itemMetaInfo),
                           insertItemsInBulk(items, ContractReview, ['Question', 'Answer', 'Comments']),
                           updateDPIRItems(dpirItems, ['Status','RejRemarks'])
                        ]).then(()=>{
                           notifyRejectedTrades()
                        });
                        //fd.save();
                      }
                    }
                }
             }
        }
    });
}

function setToolTipMessages(){

  if(_module === 'GIS')
    finalizetMesg = `Click ${submitText} for GIS Approval`;

  setButtonCustomToolTip('Save', saveMesg);
  setButtonCustomToolTip(submitText, finalizetMesg);
  setButtonCustomToolTip('Close', closeMesg);

	if($('p').find('small').length > 0)
    $('p').find('small').remove();
}

const addUpdate = async function(listname, query, itemMetaInfo){

	return await _web.lists.getByTitle(listname).items.select("Id").filter(query).get()
	.then(async items=>{
		if(items.length === 0)
            await _web.lists.getByTitle(listname).items.add(itemMetaInfo);
		else{
			let item = items[0];
			await _web.lists.getByTitle(listname).items.getById(item.Id).update(itemMetaInfo);
		}
	})
}

function addMessageEventListener(){
  window.addEventListener('message', function(event) {

    const data = event.data;
    if(data === undefined || data === null)
      return;

    if(data._module === 'CR'){
    
      let PCRtradeFormID = data.PCRtradeFormID;
      let PCRtradeFormStatus = data.PCRtradeFormStatus;
      
      if(PCRtradeFormID !== null && PCRtradeFormStatus !== null){
        let textColor = '';

        if(PCRtradeFormStatus === 'Approve')
          textColor = 'color:green;font-weight:bold;'
        else if(PCRtradeFormStatus === 'Reject')
          textColor = 'color:red;font-weight:bold;'
        else if(PCRtradeFormStatus === 'Submitted')
          textColor = 'color:#936106cf;font-weight:bold;'

            let elementStatusTd = $(`a[itemid='${PCRtradeFormID}']`);
            elementStatusTd.parent().next().find('label').text(PCRtradeFormStatus).attr('style', textColor);
      
      }
    }
  });
}
//#endregion

//PINT (PROJECT INFO)
var onPINTRender = async function (){

  const startTime = performance.now();
 
   //let firmsdt = await getFirms()
   if(_isEdit)
     localStorage.setItem('MasterId', _itemId);   

    fd.control('firmdt').filter = 'IsActive eq 1';
    fd.control('otherfirmdt').filter = 'IsActive eq 1';

    // if(_isMain)
    //   GetDictionaries(fd.field('Title').value), // Dictionaries.js

     await Promise.all([
      
    // PIR (MAIN TAB)
      onMetricsRender(),  // metrics.js
      onBackGroundInfoRender(), // backgroundInfo.js

      onDesignProcessRender(), // designProcess.js
      onSusRender('susdt', 'Goal', 'Title', 'Level', Level), // Sustainability.js
      
       setActiveFirms('firmdt'), //current PRC.js
       setActiveFirms('otherfirmdt'), //current PRC.js

       onContReviewRender(), // contractReview.js
      
       // DPIR (2nd TAB) 
       onRevDepRender(), //  revDepartments.js

       // PhCR, PCR
       onCompletionReportRender(), //CR.js

      //GENERAL
       handelTables(), // utils.js
       setMainRules() // pcr.js
     ]).then(()=>{
      handleCascadedTabsView(); // pcr.js
      const endTime = performance.now();
      const elapsedTime = endTime - startTime;
      console.log(`overall: ${elapsedTime} milliseconds`);
  
      //getConsoleLogRoles();
     });
     
}

let setActiveFirms = async function(dtName){
  const startTime = performance.now();
  let dt = fd.control(dtName)

  dt.$on('edit', function(editData){
    editData.field('Firm').filter = 'IsActive eq false';
    editData.field('Firm').useCustomFilterOnly = true;
    editData.field('Firm').orderBy = { field: 'Title', desc: false };
  })
  
  const endTime = performance.now();
  const elapsedTime = endTime - startTime;
  console.log(`setActiveFirms: ${elapsedTime} milliseconds`);
 
}

var setMainRules = async function(){
  let categoryField = fd.field('Category')
  let category = categoryField.value !== undefined && categoryField.value !== null ? categoryField.value.LookupValue : '';
  if(category !== ''){
    let isDisabled = true;
    if(category === 'Buildings'){
      _isBuilding = true;
      isDisabled = false;
    }

    if(_isMain)
      localStorage.setItem('_isBuilding', _isBuilding);

    if(isDisabled)
      await enable_Disable_Tabs(tabs, isDisabled);
    categoryField.disabled = true
  }

  await getEditableTabFields();
  await setPINT_Permissions();
}

function handleCascadedTabsView(){
  const containers = document.querySelector('.fd-grid.container-fluid');
  const getTabs = containers.querySelectorAll('.tabset .tabs-top');
  getTabs.forEach((tab, index) => {
    if(index == 0){
  
        const tabsUL = tab.querySelectorAll('.nav-item');        
        tabsUL.forEach(navItem => {          
          navItem.style.width = '200px';
          navItem.style.textAlign = 'center'; // Center the text within the tab item            
        //   navItem.style.margin = '0 2px'; // Margin between tabs
        //   navItem.style.border = 'none'; // Remove default border
        //   navItem.style.borderBottom = '2px solid transparent'; // Default bottom border
        //   navItem.style.borderRadius = '4px'; // Rounded corners
        //   //navItem.style.backgroundColor = '#f8f9fa'; // Light background color
        //   navItem.style.color = '#333'; // Text color
        //   navItem.style.transition = 'all 0.3s ease'; // Smooth transition for hover effects

        //   // Apply styles for active tab
        //   if (navItem.classList.contains('active')) {
        //       navItem.style.borderBottom = '2px solid #007bff'; // Active tab bottom border
        //       //navItem.style.backgroundColor = '#fff'; // Active tab background color
        //       navItem.style.color = '#007bff'; // Active tab text color
        //       navItem.style.fontWeight = 'bold'; // Bold text for active tab
        //   }
        });
    }   
  });  
}

async function setPINT_Permissions(){

  if(_isMain){
    let mainPlayers = _isPD || _isPM || _isGLMain || _isGL || _isQM ? true : false
    if(!mainPlayers)
      fd.container('TabsMain').tabs[1].disabled = true; // Disable Review Department Initiation

    if(!mainPlayers && !_isLLChecker)
       fd.container('TabsMain').tabs[2].disabled = true; // Disable Completion Report
    
   if(!_isPD && !_isPM && !_isQM){
      tabs.push({
        masterTab: secondChildTab,
        title: 'Clarification Register',
        tooltip: 'Permission Denied'
      })
    }

    if(!_isPD && !_isPM)
      await handleProjectInfoFields();
  }
}

async function handleProjectInfoFields(){
  let textFields = {
    //General Info
    MainOffice: fd.field("MainOffice"),
    DesignOffice: fd.field("DesignOffice"),
    DateContSigned: fd.field("DateContSigned"),
    Country: fd.field("Country"),
    City: fd.field("City"),
    EstCost: fd.field('EstCost'),
    IsConfidential: fd.field('IsConfidential'),
    Description: fd.field('Description'),
    Attachments: fd.field('Attachments'),
    
    //Classification & Metrics
    Category: fd.field('Category'),
    SubCategory: fd.field('SubCategory'),
    showAll: fd.field('showAll'),
    MetricsFields: '#metricId',

    //Background Info
    GroundWater: fd.field('GroundWater'),
    SeisDesReq: fd.field('SeisDesReq'),
    ListOfItems: fd.field('ListOfItems'),
    Lesslearn: fd.field('Lesslearn'),
    MTD: fd.control('Mtddt'),
    ContractForm: fd.field('ContractForm'),
    FormContractOther: fd.field('FormContractOther'),
    ContractType: fd.field('ContractType'),
    TypeContractOther: fd.field('TypeContractOther'),
    Client: fd.control('Contactsdt'),

    //Design Process
    Process: fd.field('Process'),
    CADJust: fd.field('CADJust'),
    BIM: fd.control('bimdt'),
    DRDApplicable: fd.field('DRDApplicable'),
    DRDJustification: fd.field('DRDJustification'),
    DSSInvolved: fd.field('DSSInvolved'),

    //Sustainability
    SUS: fd.control('susdt'),

    //Other Parties & Firms
    FIRM: fd.control('firmdt'),
    OtherFirm: fd.control('otherfirmdt'),

    //Contract Review
    Questions : '#tblItemsId'
  }

  Object.entries(textFields).forEach(([key, field]) => {
      if(key === 'MetricsFields' || key === 'Questions'){
        setTimeout(async () => {
          $(field).find('input, select, button, textarea').attr('disabled', true);
        }, 500);
      }
      else if(key === 'MTD' || key === 'Client' || key === 'BIM' || key === 'SUS' || key === 'FIRM' || key === 'OtherFirm')
        field.readonly = true;
      else if(field === "Attachments"){	
        $('div.k-upload-button').remove();
        $('button.k-upload-action').remove();
        $('.k-dropzone').remove();
      }
      else  field.disabled = true;
  });
}


//for PCR CHECKING, so if all submitted then send email to PM
let isAllTradesSubmitted = async function(department){

  let MasterID = fd.field('MasterID').value.LookupId;
  const list = _web.lists.getByTitle(CompReports);
  let query = `MasterID/Id eq ${MasterID} and Title ne 'PM' and Title ne '${TextQueryEncode(department)}' and IsProjectComp eq 1`;
  let items = await list.items
                 .select("Id, Title, Status, IsProjectComp, MasterID/Id")
                 .expand("MasterID")
                 .filter(query)
                 .get();
  let AllSubmittedItems = items.filter((item)=>{ return item.Status === 'Submitted'})

  localStorage.setItem('IsAllSubmitted', items.length === AllSubmittedItems.length)
}



