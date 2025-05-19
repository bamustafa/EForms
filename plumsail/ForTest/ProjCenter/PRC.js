var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isMain = false, _isLead = false, _isPart = false, _isNew = false, _isEdit = false, _isDisplay = false
_isPD = false, _isPM = false, _isQM = false, _isSus = false, _isGIS = false, _isGLMain = false, _isGL = false, _isTeamMember = false, _isQMOwner = false, _isSusOwner = false,
_isLLChecker = false, _isReader = false, _isConfidential = false, _isBuilding = false, _isReader = false, _isUserAllowed = false, _hideSubmit = false, _isPROLE = false;

var projectNo = '', projectTitle = '', CurrentUser;

const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';



var ProjectInfo = 'Project Info', Category = 'Category', SubCategory = 'SubCategory', MatrixFields = 'Matrix Fields', MTDs = 'MTDs', Firms = 'Firms', Questions = 'Questions', 
    ContractReview = 'Contract Review', OtherPartiesFirms = 'Other Parties Firms', Sustainability = 'Sustainability', Level = 'Levels', Roles = 'Roles', 
    ProjectRoles = 'Project Roles', GISLocation = 'GIS Location', RevDepartments = 'Review Departments', CompReports = 'Completion Reports', 
    TradeContractReview = 'Trade Contract Review', LLChecker = 'LLChecker';

var formFields = {};

var firstChildTab = 'TabsMain', secondChildTab = 'Tabs6'

var tabs = [
            {
              masterTab: firstChildTab,
              level: 2,
              title: 'Sustainability123',
              tooltip: 'Sustainability is enabled for Buildings Category only'
            },
            {
              masterTab: firstChildTab,
              level: 2,
              title: 'Background Info123',
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
    await setLegends();
      _isEdit = true
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

      let fullProjTitle = `<img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px; margin-top:10px;" /> ${projectNo} - ${projectTitle}`;
      localStorage.setItem('projectNo', projectNo);
      localStorage.setItem('ProjectTitle', projectTitle);
      localStorage.setItem('FullProjTitle', fullProjTitle);

      fullProjTitle += ' - Project Management Initiation Form'
  const linkElement = `<a href="${_webUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">${fullProjTitle}</a>`;
  $('span.o365cs-nav-brandingText').html(linkElement);

      setPageStyle(`${fullProjTitle}`);  
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
    let mainPlayers = _isPD || _isPM || _isGLMain || _isGL || _isQM || _isSiteAdmin ? true : false
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


var setLegends = async function(){

  try {
      const svguserinfo = `<svg version="1.1" id="Layer_1" width="24" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 34 32" enable-background="new 0 0 34 32" xml:space="preserve" fill="#49c4b1" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <g> <path fill="#0000828282" d="M1.512,28H19.5c0.827,0,1.5-0.673,1.5-1.5v-19c0-0.023-0.01-0.043-0.013-0.065 c-0.003-0.022-0.007-0.041-0.013-0.062c-0.023-0.086-0.06-0.166-0.122-0.227l-6.999-6.999c-0.061-0.061-0.141-0.098-0.227-0.122 c-0.021-0.006-0.04-0.01-0.062-0.013C13.543,0.01,13.523,0,13.5,0H1.506C0.676,0,0,0.673,0,1.5v25C0,27.327,0.678,28,1.512,28z M14,1.707L19.293,7H14.5C14.225,7,14,6.776,14,6.5V1.707z M1,1.5C1,1.224,1.227,1,1.506,1H13v5.5C13,7.327,13.673,8,14.5,8H20 v18.5c0,0.276-0.225,0.5-0.5,0.5H1.512C1.229,27,1,26.776,1,26.5V1.5z"></path> <path fill="#0000828282" d="M4.5,12h12c0.276,0,0.5-0.224,0.5-0.5S16.776,11,16.5,11h-12C4.224,11,4,11.224,4,11.5S4.224,12,4.5,12z"></path> <path fill="#0000828282" d="M4.5,16h12c0.276,0,0.5-0.224,0.5-0.5S16.776,15,16.5,15h-12C4.224,15,4,15.224,4,15.5S4.224,16,4.5,16z"></path> <path fill="#0000828282" d="M4.5,8h5C9.776,8,10,7.776,10,7.5S9.776,7,9.5,7h-5C4.224,7,4,7.224,4,7.5S4.224,8,4.5,8z"></path> <path fill="#0000828282" d="M4.5,20h12c0.276,0,0.5-0.224,0.5-0.5S16.776,19,16.5,19h-12C4.224,19,4,19.224,4,19.5S4.224,20,4.5,20z"></path> <path fill="#0000828282" d="M4.5,24h12c0.276,0,0.5-0.224,0.5-0.5S16.776,23,16.5,23h-12C4.224,23,4,23.224,4,23.5S4.224,24,4.5,24z"></path> <path fill="#0000828282" d="M21.5,5H26v5.5c0,0.827,0.673,1.5,1.5,1.5H33v18.5c0,0.276-0.225,0.5-0.5,0.5H14.512 C14.229,31,14,30.776,14,30.5v-1c0-0.276-0.224-0.5-0.5-0.5S13,29.224,13,29.5v1c0,0.827,0.678,1.5,1.512,1.5H32.5 c0.827,0,1.5-0.673,1.5-1.5v-19c0-0.023-0.01-0.043-0.013-0.065c-0.003-0.022-0.007-0.041-0.013-0.062 c-0.023-0.086-0.06-0.166-0.122-0.227l-6.999-6.999c-0.061-0.062-0.142-0.099-0.228-0.122c-0.021-0.006-0.039-0.009-0.061-0.012 C26.543,4.01,26.523,4,26.5,4h-5C21.224,4,21,4.224,21,4.5S21.224,5,21.5,5z M27.5,11c-0.275,0-0.5-0.224-0.5-0.5V5.707L32.293,11 H27.5z"></path> <path fill="#0000828282" d="M23.5,16h6c0.276,0,0.5-0.224,0.5-0.5S29.776,15,29.5,15h-6c-0.276,0-0.5,0.224-0.5,0.5S23.224,16,23.5,16z "></path> <path fill="#0000828282" d="M23.5,20h6c0.276,0,0.5-0.224,0.5-0.5S29.776,19,29.5,19h-6c-0.276,0-0.5,0.224-0.5,0.5S23.224,20,23.5,20z "></path> <path fill="#0000828282" d="M23.5,24h6c0.276,0,0.5-0.224,0.5-0.5S29.776,23,29.5,23h-6c-0.276,0-0.5,0.224-0.5,0.5S23.224,24,23.5,24z "></path> <path fill="#0000828282" d="M23.5,28h6c0.276,0,0.5-0.224,0.5-0.5S29.776,27,29.5,27h-6c-0.276,0-0.5,0.224-0.5,0.5S23.224,28,23.5,28z "></path> </g> </g></svg>`;            
      setIconSource("technical-icon", svguserinfo);     
      const pncsvg = `<svg fill="#000000" width="24" viewBox="0 0 24 24" id="contract" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"><path id="primary" d="M7,13v7a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V4a1,1,0,0,0-1-1H8A1,1,0,0,0,7,4V7" style="fill: none; stroke: #49c4b1; stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="secondary" d="M3.29,7.69l1.4-1.4a1,1,0,0,1,1.4,0L11,11.2V14H8.2L3.29,9.09A1,1,0,0,1,3.29,7.69ZM17,17H14" style="fill: none; stroke: #40c4b1; stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></g></svg>`;      
      setIconSource("contract-icon", pncsvg);   
      const pmcsvg = `<svg fill="#49c4b1" width="24" version="1.1" id="Capa_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 176.646 176.646" xml:space="preserve" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <g> <g> <path d="M89.96,96.038l3.2-3.678c-1.547,0.56-3.145,0.895-4.835,0.895c-1.689,0-3.285-0.347-4.844-0.907l3.203,3.69h0.088 l-8.257,19.973l9.816,9.797l9.821-9.797l-8.263-19.973H89.96z"></path> <path d="M140.364,0H36.285C25.163,0,16.121,9.036,16.121,20.152v4.582h12.596c4.177,0,7.557,3.392,7.557,7.56 c0,4.171-3.379,7.55-7.557,7.55H16.121V52.75h12.596c4.177,0,7.557,3.392,7.557,7.557c0,4.171-3.379,7.551-7.557,7.551H16.121 v12.909h12.596c4.177,0,7.557,3.385,7.557,7.557c0,4.171-3.379,7.557-7.557,7.557H16.121v12.908h12.596 c4.177,0,7.557,3.38,7.557,7.551c0,4.165-3.379,7.557-7.557,7.557H16.121v12.903h12.596c4.177,0,7.557,3.386,7.557,7.55 c0,4.166-3.379,7.563-7.557,7.563H16.121v4.579c0,11.119,9.042,20.154,20.165,20.154h104.079c11.119,0,20.161-9.035,20.161-20.154 V20.152C160.531,9.036,151.489,0,140.364,0z M53.301,117.18c0-9.743,11.237-22.914,26.369-26.756 c-6.321-4.253-10.687-12.669-10.687-20.24c0-10.675,8.656-19.354,19.342-19.354c10.693,0,19.352,8.68,19.352,19.354 c0,7.571-4.372,15.987-10.705,20.24c15.15,3.842,26.379,17.013,26.379,26.756C123.345,128.712,53.301,128.712,53.301,117.18z"></path> </g> </g> </g></svg>`;      
      setIconSource("client-icon", pmcsvg); 
      const pmc2svg = `<svg fill="#49c4b1" width="24" viewBox="0 0 128 128" id="Layer_1" version="1.1" xml:space="preserve" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <g> <path d="M64,42c-13.2,0-24,10.8-24,24s10.8,24,24,24s24-10.8,24-24S77.2,42,64,42z M64,82c-8.8,0-16-7.2-16-16s7.2-16,16-16 s16,7.2,16,16S72.8,82,64,82z"></path> <path d="M64,100.8c-14.9,0-29.2,6.2-39.4,17.1l-2.7,2.9l5.8,5.5l2.7-2.9c8.8-9.4,20.7-14.6,33.6-14.6s24.8,5.2,33.6,14.6l2.7,2.9 l5.8-5.5l-2.7-2.9C93.2,107.1,78.9,100.8,64,100.8z"></path> <path d="M97,47.9v8c9.4,0,18.1,3.8,24.6,10.7l5.8-5.5C119.6,52.7,108.5,47.9,97,47.9z"></path> <path d="M116.1,20c0-10.5-8.6-19.1-19.1-19.1S77.9,9.5,77.9,20S86.5,39.1,97,39.1S116.1,30.5,116.1,20z M85.9,20 c0-6.1,5-11.1,11.1-11.1s11.1,5,11.1,11.1s-5,11.1-11.1,11.1S85.9,26.1,85.9,20z"></path> <path d="M31,47.9c-11.5,0-22.6,4.8-30.4,13.2l5.8,5.5c6.4-6.9,15.2-10.7,24.6-10.7V47.9z"></path> <path d="M50.1,20C50.1,9.5,41.5,0.9,31,0.9S11.9,9.5,11.9,20S20.5,39.1,31,39.1S50.1,30.5,50.1,20z M31,31.1 c-6.1,0-11.1-5-11.1-11.1S24.9,8.9,31,8.9s11.1,5,11.1,11.1S37.1,31.1,31,31.1z"></path> </g> </g></svg>`;      
      setIconSource("BIM-icon", pmc2svg); 


      const pmc3svg = `<svg fill="#49c4b1" width="30" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"><path d="M 12 6.59375 L 11.28125 7.28125 L 3.28125 15.28125 L 2.59375 16 L 3.28125 16.71875 L 11.28125 24.71875 L 12 25.40625 L 12.71875 24.71875 L 16.71875 20.71875 L 17.40625 20 L 16.71875 19.28125 L 11.71875 14.28125 L 10.28125 15.71875 L 14.5625 20 L 12 22.5625 L 5.4375 16 L 12 9.4375 L 13.28125 10.71875 L 14.71875 9.28125 L 12.71875 7.28125 Z M 20 6.59375 L 19.28125 7.28125 L 15.28125 11.28125 L 14.59375 12 L 15.28125 12.71875 L 20.28125 17.71875 L 21.71875 16.28125 L 17.4375 12 L 20 9.4375 L 26.5625 16 L 20 22.5625 L 18.71875 21.28125 L 17.28125 22.71875 L 19.28125 24.71875 L 20 25.40625 L 20.71875 24.71875 L 28.71875 16.71875 L 29.40625 16 L 28.71875 15.28125 L 20.71875 7.28125 Z"></path></g></svg>`;      
      setIconSource("general-icon", pmc3svg); 
      setIconSource("gen-icon", pmc3svg); 
      setIconSource("g-icon", pmc3svg); 


      const pmc7svg = `<svg version="1.1" width="24" id="Uploaded to svgrepo.com" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 32 32" xml:space="preserve" fill="#000000"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <style type="text/css"> .blueprint_een{fill:#49c4b1;} </style> <path class="blueprint_een" d="M8,9H4v3h4V9z M7,11H5v-1h2V11z M8,13H4v3h4V13z M7,15H5v-1h2V15z M6,17c-1.657,0-3,1.343-3,3 s1.343,3,3,3s3-1.343,3-3S7.657,17,6,17z M6,22c-1.103,0-2-0.897-2-2c0-1.103,0.897-2,2-2s2,0.897,2,2C8,21.103,7.103,22,6,22z M21.646,11.646l0.707,0.707l-2,2l-0.707-0.707L21.646,11.646z M31,4h-1.586l-1.707-1.707c-0.391-0.391-1.023-0.391-1.414,0 L24.586,4H1C0.448,4,0,4.448,0,5v22c0,0.552,0.448,1,1,1h30c0.552,0,1-0.448,1-1V5C32,4.448,31.552,4,31,4z M16,14v4h4l8-8v13H11V9 h10L16,14z M17,14.707L19.293,17H17V14.707z M20.146,16.439l-2.586-2.586L27,4.414L29.586,7L20.146,16.439z M2,26V6h22l-2,2H10v16 h19V9l1-1v18H2z"></path> </g></svg>`;      
      setIconSource("DRD-icon", pmc7svg); 
      const pmc8svg = `<svg viewBox="0 0 31 31" width="24" version="1.1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:sketch="http://www.bohemiancoding.com/sketch/ns" fill="#49c4b1" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <title>permalink</title> <desc>Created with Sketch Beta.</desc> <defs> </defs> <g id="Page-1" stroke="none" stroke-width="1" fill="none" fill-rule="evenodd" sketch:type="MSPage"> <g id="Icon-Set" sketch:type="MSLayerGroup" transform="translate(-153.000000, -723.000000)" fill="#49c4b1"> <path d="M171.702,735.282 C171.307,734.888 170.667,734.888 170.272,735.282 L165.267,740.287 C164.872,740.683 164.872,741.322 165.267,741.718 C165.662,742.112 166.302,742.112 166.697,741.718 L171.702,736.713 C172.097,736.317 172.097,735.678 171.702,735.282 L171.702,735.282 Z M182.785,725.613 L181.355,724.184 C179.775,722.604 177.214,722.604 175.635,724.184 L168.484,731.334 C167.428,732.39 167.095,733.881 167.451,735.228 L177.064,725.613 C177.854,724.824 179.135,724.824 179.925,725.613 L181.355,727.044 C182.145,727.834 182.145,729.114 181.355,729.904 L171.741,739.518 C173.088,739.873 174.579,739.54 175.635,738.484 L182.785,731.334 C184.365,729.755 184.365,727.193 182.785,725.613 L182.785,725.613 Z M159.904,751.354 C159.114,752.145 157.833,752.145 157.043,751.354 L155.614,749.925 C154.824,749.135 154.824,747.854 155.614,747.064 L165.228,737.451 C163.881,737.095 162.39,737.429 161.334,738.484 L154.184,745.635 C152.604,747.214 152.604,749.775 154.184,751.354 L155.614,752.785 C157.193,754.365 159.754,754.365 161.334,752.785 L168.484,745.635 C169.54,744.579 169.874,743.088 169.518,741.741 L159.904,751.354 L159.904,751.354 Z" id="permalink" sketch:type="MSShapeGroup"> </path> </g> </g> </g></svg>`;      
      setIconSource("DSS-icon", pmc8svg); 



      const pmc9svg = `<svg viewBox="0 0 64 64" width="28" xmlns="http://www.w3.org/2000/svg" fill="#000000"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <defs> <style>.cls-1{fill:none;stroke:#49C4B1;stroke-linejoin:round;stroke-width:2px;}</style> </defs> <title></title> <g data-name="Layer 30" id="Layer_30"> <rect class="cls-1" height="4" width="60" x="2" y="58"></rect> <rect class="cls-1" height="46" width="20" x="22" y="12"></rect> <polyline class="cls-1" points="22 12 32 2 42 12"></polyline> <polyline class="cls-1" points="22 28 6 28 6 58"></polyline> <rect class="cls-1" height="18" width="8" x="10" y="32"></rect> <line class="cls-1" x1="6" x2="22" y1="54" y2="54"></line> <line class="cls-1" x1="10" x2="18" y1="38" y2="38"></line> <line class="cls-1" x1="10" x2="18" y1="44" y2="44"></line> <line class="cls-1" x1="14" x2="14" y1="32" y2="50"></line> <polyline class="cls-1" points="42 28 58 28 58 58"></polyline> <rect class="cls-1" height="18" transform="translate(100 82) rotate(180)" width="8" x="46" y="32"></rect> <line class="cls-1" x1="58" x2="42" y1="54" y2="54"></line> <line class="cls-1" x1="54" x2="46" y1="38" y2="38"></line> <line class="cls-1" x1="54" x2="46" y1="44" y2="44"></line> <line class="cls-1" x1="50" x2="50" y1="32" y2="50"></line> <circle class="cls-1" cx="32" cy="22" r="6"></circle> <path class="cls-1" d="M26,58V52a6,6,0,0,1,6-6h0a6,6,0,0,1,6,6v6"></path> <rect class="cls-1" height="4" width="10" x="27" y="38"></rect> <line class="cls-1" x1="34" x2="34" y1="52" y2="54"></line> <polyline class="cls-1" points="32 19 32 22 34 24"></polyline> <polyline class="cls-1" points="50 28 50 16 58 16 58 22 50 22"></polyline> <polyline class="cls-1" points="14 28 14 16 6 16 6 22 14 22"></polyline> </g> </g></svg>`;      
      setIconSource("firms-icon", pmc9svg); 

      const pm12svg = `<svg fill="#49c4b1" width="28" version="1.2" baseProfile="tiny" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 256 220" xml:space="preserve" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path id="XMLID_28_" d="M251.2,60.9l-55.3-50c-1.9-1.7-4.9-1.6-6.7,0.3l-16.5,18.2l-67-14.9c-1.8-0.7-3.7-1.2-5.8-1.2 c-6.5,0-12.1,4.1-14.2,9.9c-0.4,0.6-0.8,1.5-1,2.4c0,0-10.4,45.1-13.9,60.1c-0.1,0.4-0.1,0.8-0.2,1.1c-0.1,0.8-0.2,1.7-0.2,2.6 c0,7.9,6.4,14.3,14.3,14.3c6.3,0,11.7-4.1,13.6-9.7c0.3-0.5,0.5-1.1,0.6-1.8l6.3-27.1l0-0.2c3.2-2.5,7.3-4.1,11.7-4.1 c7.3,0,13.9,4.1,16.8,10c0.1,0.2,0.2,0.4,0.3,0.6c0.2,0.4,0.5,0.8,0.9,1.2l5.6,5.6l35.5,35.5c0.1,0.3,0.3,0.6,0.6,0.9l10.5,10.5 l13.5,13.4c1.5,1.7,2.4,4,2.4,6.5c0,5.5-4.5,10-10,10c-3.2,0-6.1-1.5-8-3.9l-0.4-0.4l-7.3-7.3l-15.3-15.3c-1.1-1.1-2.9-1.1-4,0 c-1.1,1.1-1.1,2.9,0,4l15.6,15.6c0,0,0,0,0,0l7.1,7.1c0.4,0.6,0.9,1.1,1.5,1.6l0,0c1.7,1.8,2.8,4.2,2.8,6.9c0,5.5-4.5,10-10,10 c-3.2,0-6.1-1.5-8-3.9l-1-1l-6.6-6.6l-15.9-15.9c-1.1-1.1-2.9-1.1-4,0s-1.1,2.9,0,4l15.9,15.9c0,0,0,0,0,0l7.4,7.4 c0.4,0.4,0.7,0.9,1.2,1.3c1.7,1.8,2.8,4.2,2.8,6.9c0,5.5-4.5,10-10,10c-3.2,0-6.1-1.5-8-3.9l-1.3-1.3l-6.4-6.4l-16.2-16.2 c-1.1-1.1-2.9-1.1-4,0c-1.1,1.1-1.1,2.9,0,4l16.5,16.5c0,0,0,0,0,0l7.1,7.1c0.4,0.6,0.9,1.1,1.5,1.6l0,0c1.7,1.8,2.8,4.2,2.8,6.9 c0,5.5-4.5,10-10,10c-3.2,0-6.1-1.5-8-3.9l-7.7-7.6v0l-6.1-6.1l1-1c0.1-0.1,0.2-0.2,0.2-0.3c2.1-2.6,3.4-6,3.4-9.6 c0-8.3-6.7-15.1-15.1-15.1c-1.3,0-2.5,0.2-3.7,0.4c0.3-1.2,0.4-2.4,0.4-3.7c0-8.3-6.7-15.1-15.1-15.1c-1.3,0-2.5,0.2-3.7,0.4 c0.3-1.1,0.4-2.2,0.4-3.4c0-8.3-6.7-15.1-15.1-15.1c-1.1,0-2.1,0.1-3.1,0.3c0.3-1.2,0.4-2.3,0.4-3.6c0-8.3-6.7-15.1-15.1-15.1 c-4.4,0-8.3,1.9-11,4.8L8.9,84.2c-1.7-1.7-3.9-2.1-5.1-1c-1.1,1.1-0.7,3.4,1,5.1l32.1,32.1c-4.4,4.4-10.8,10.8-14.9,14.9 c-0.4,0.4-0.8,0.8-1.2,1.2c-0.3,0.3-0.7,0.7-1,1c-0.1,0.1-0.2,0.3-0.2,0.5c-1.7,2.4-2.7,5.4-2.7,8.6c0,8.3,6.7,15.1,15.1,15.1 c1.1,0,2.1-0.1,3.1-0.3c-0.3,1.2-0.4,2.3-0.4,3.6c0,8.3,6.7,15.1,15.1,15.1c1.3,0,2.5-0.2,3.7-0.4c-0.3,1.1-0.4,2.2-0.4,3.4 c0,8.3,6.7,15.1,15.1,15.1c1.3,0,2.5-0.2,3.7-0.4c-0.3,1.2-0.4,2.4-0.4,3.7c0,8.3,6.7,15.1,15.1,15.1c5,0,9.4-2.4,12.1-6.1 l14.1-14.1l13.5,13.5c0.2,0.2,0.4,0.4,0.7,0.6c2.9,3.1,7,5.1,11.5,5.1c8.7,0,15.7-7,15.7-15.7c0-1-0.1-1.9-0.3-2.9 c0.9,0.1,1.7,0.2,2.6,0.2c8.7,0,15.7-7,15.7-15.7c0-0.9-0.1-1.8-0.2-2.6c0.9,0.2,1.9,0.3,2.9,0.3c8.7,0,15.7-7,15.7-15.7 c0-1-0.1-1.9-0.3-2.9c0.9,0.1,1.7,0.2,2.6,0.2c8.7,0,15.7-7,15.7-15.7c0-4.5-1.9-8.5-4.9-11.4c-0.1-0.1-0.1-0.2-0.2-0.3l-9.2-9.2 c-0.1,0-0.1,0-0.2,0c-2.3,0.5-4.5,0.8-6.8,0.9c2.3-0.1,4.5-0.4,6.7-0.9c0.1,0,0.1,0,0.2,0c2.4-0.5,4.8-1.2,7-2.1 c4-1.6,7.7-3.7,11-6.4c2.8-2.2,5.2-4.6,7.4-7.4l12.8-19.5l19.1-21.1C253.3,65.6,253.1,62.6,251.2,60.9z"></path> </g></svg>`;
      setIconSource("OP-icon", pm12svg); 
      const pmc10svg = `<svg  width="28" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 512 512" xml:space="preserve" fill="#000000"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <circle style="fill:#ffffff;" cx="256" cy="256" r="256"></circle> <path style="fill:#49c4b1;" d="M256,145.224c61.177,0,110.776,49.6,110.776,110.776S317.177,366.776,256,366.776 S145.224,317.177,145.224,256S194.823,145.224,256,145.224z"></path> <path style="fill:#49c4b1;" d="M256,145.224c43.868,0,81.777,25.481,99.71,62.482c2.724,10.045,4.143,20.6,4.143,31.496 c0,66.852-54.196,120.991-120.991,120.991c-9.931,0-19.579-1.192-28.829-3.462c-38.193-17.479-64.752-56.012-64.752-100.788 c0-61.177,49.6-110.776,110.776-110.776L256,145.224z"></path> <path style="fill:#F7F9F7;" d="M325.235,169.513c20.6,16.514,35.185,40.179,39.895,67.306c-0.341-0.511-0.681-1.192-1.022-2.157 c-2.156,0.738-2.213,0.738-4.767,1.192l-0.114,0.17c-1.021,1.419,0.057,0.851-1.759,2.327c-2.1,1.816-3.518,4.256-5.902,5.675 c-0.284,2.554,0.227,3.121-0.397,6.129c-0.227,1.135-1.476,4.086-2.951,3.859c-0.454-0.057-2.44-4.937-2.894-6.243 c-0.794-2.27-1.646-3.973-2.213-6.583c-0.341-1.419-0.511-4.086-0.851-4.994c-1.532,0.227-1.475,1.192-2.894,0.17 c-1.022-0.738-0.568-1.192-1.135-2.327c-1.192-1.135-2.44-2.27-3.689-3.405c-0.965-0.284-4.654,0.227-6.186,0.057 c-2.781-0.284-3.065-0.113-4.71-2.327c-2.951,0.17-2.724,1.249-5.562-1.249c-1.873-1.589-1.532-4.143-4.256-2.724 c-0.113,2.327,1.021,2.838,2.043,3.916c0.738,0.851,1.646,1.419,2.043,2.497c0.057,0.057,0.227,1.078,0.227,1.192 c2.667,0.17,3.292-1.078,4.767-2.157c1.419,1.305,0.057,1.589,2.1,2.611c1.646,0.851,1.305,0.34,2.384,1.702 c0.511,1.816-3.632,6.64-5.051,7.548c-1.078,0.681-1.646,0.284-2.611,1.078c-1.419,1.192,0.17,1.021-2.327,1.93 c-2.667,0.965-5.278,3.008-8.059,3.178c-0.624-1.646-1.192-3.348-1.816-4.994c-1.759-5.732-1.702-3.519-3.518-6.47 c-0.851-1.305-0.511-2.1-1.078-3.235c-0.397-0.851-0.965-1.249-1.532-2.27c-0.795-1.646-1.759-4.143-3.632-3.916 c-0.397,1.249,1.759,5.618,2.27,6.64c0.511,1.078,1.362,1.532,1.646,3.008c0.34,2.043-0.057,1.816,1.078,3.178 c2.384,2.894-0.113,2.384,3.575,5.164c2.781,2.1,2.043,2.724,3.235,5.562c1.816,0.511,7.207-1.873,8.115-1.192 c1.078,1.986-1.249,5.108-1.873,6.81c-1.986,5.732-5.448,5.618-8.115,9.988c-1.475,2.383-2.1,2.667-3.462,5.505 c-1.532,3.235,3.065,12.088,0.341,15.833c-0.624,0.851-2.554,1.362-4.994,3.973c-2.611,2.781,0.057,2.327-0.908,7.037 c-0.794,1.078-1.759,1.135-2.667,1.93c-0.341,1.078-0.057,1.986-0.681,3.348c-0.624,1.646-0.738,0.738-1.646,2.1 c-0.681,1.022-0.624,1.476-1.362,2.44c-0.681,0.908-1.021,1.135-1.646,1.702c-3.065,2.838-9.704,2.951-10.272,2.554 c-1.589-1.305-0.511-1.703-0.908-3.575c-0.397-1.646-1.249-3.008-1.986-4.313l-0.057-0.113c-0.341-0.624-0.681-1.249-0.851-1.816 c-1.249-4.029-0.794-7.094-3.065-11.18c-2.781-5.278,1.532-8.059,1.362-11.01c-0.057-0.794-0.34-1.249-0.397-1.93 c-0.511-3.292-0.851-6.243-3.065-8.853c-2.44-2.837-2.1-1.816-1.305-6.47c0.965-5.335-1.249-3.973-4.483-4.086 c-1.589-4.824-4.597-2.497-7.661-1.022c-1.93,0.851-2.611-0.227-4.767,0.057c-4.824,0.681-1.021,2.724-7.491-2.384 c-2.667-2.157-0.908-1.021-2.497-3.689c-0.511-0.965-2.157-2.27-2.894-3.235c-2.44-3.065,0.454-6.186,0.057-9.704 c-0.738-6.526-1.816-0.851,1.816-8.74c1.93-4.2,4.824-4.256,5.562-6.129c0.624-1.703,0.227-2.894,1.419-4.427 c1.078-1.249,2.611-1.022,3.178-2.951c0.851-2.44-0.795-2.724-2.554-2.781c-0.738-2.554-0.738-1.873-0.17-4.256 c0.34-1.305,0.057-1.93,0.34-3.235c0.738-3.292,6.299,0,7.888-1.362c0.738-1.419,0.851-2.781,0-4.086 c-0.908-1.305-2.043-1.135-2.667-2.44c1.362-0.908,0.34-0.113,1.532-0.34c1.362-0.284,1.192-0.738,2.837-1.021 c1.476-0.284,2.384-1.646,3.802-2.554c2.043-1.305,1.078-2.497,3.518-2.894c3.916-0.681,1.362-0.568,1.816-4.256 c0.624-0.227,1.192-0.511,1.759-0.738c1.249,1.646-0.34,0.965,1.192,2.043c0.397-0.227,0.851-0.454,1.249-0.681 c-0.227-1.759-1.022-1.419-1.532-3.632c-1.93,0.284-2.667,1.986-4.2,1.362c-1.532-0.568-1.589-2.781-1.532-4.597 c0.568-1.362,3.405-2.1,4.483-2.44c1.759-0.624,1.532-1.816,3.178-3.405c2.1-2.157,0.568,0,1.589-2.724 c2.327-0.113,1.249-0.17,2.611-0.908l0.057-0.057l0.17-0.057c0.794-0.397,1.873-0.454,2.951-0.851 c3.916-1.305,5.505-0.681,8.342,0.511c2.213,0.908,8.853,0.965,9.307,3.746c-1.589,1.646-3.121,0.454-5.448,0.624 c0.397,2.157,0.624,2.157,2.554,2.724c0.34-1.022-0.113-0.624,0.397-1.249c0.624,0.113,1.305,0.227,1.93,0.34 c0.624-2.213,1.362-1.873,3.575-2.157c-0.34-1.532-0.738-1.078,0.057-2.27c2.213,0.227,0.397,0.057,1.532,1.816 c1.532-0.17,3.064-1.93,5.221-2.156c1.362-0.114,1.249,0.397,3.178,0.17c1.589-0.17,1.475-0.738,3.065-0.057 c0.568-1.078,0-1.021,1.362-1.305c2.1-0.34,3.973,0.908,6.016,1.249c-1.532-1.873-2.951-1.532-1.078-4.654 c2.384-0.511,2.611-0.113,3.632,1.646c0.568-0.34,0.851-0.681,1.135-0.851L325.235,169.513z M147.891,232.108 c4.994-22.7,16.912-42.79,33.483-57.942c1.305,0.454,3.008,0.17,4.256,0.114c0.454-0.851,0.738-1.589,1.532-1.93 c0.34-0.113,0.738-0.113,1.249-0.113c0.057,0.397,0.17,0.851,0.284,1.249c0,0.057,0.057,0.34,0.057,0.284 c0.057-0.057,0.057,0.17,0.114,0.227c2.213-1.305,0.34-1.305,1.532-3.689c-1.873-0.34-1.646-0.17-2.781-1.362 c1.135-1.249,3.348-1.703,4.994-1.532c-0.113,1.135-0.624,1.078-0.227,2.043c1.249-0.794,2.156-2.894,4.2-2.554 c0.397,0.284,0.851,0.568,1.249,0.795c-1.078,1.021-2.667,1.078-4.029,1.646c-0.624,6.186,1.93,1.021,1.078,5.221 c1.532-0.113,3.008-1.248,3.575-2.383c-0.568-0.738-1.362-0.738-1.816-1.646c0.34-1.135,1.135-1.986,2.213-2.497 c0.738-0.34,1.589-0.511,2.611-0.511c11.293,0.114,5.902,0.624,9.421,2.213c0.795,0.34,2.157,0.511,2.611,1.192 c0.738,1.078,0.057,2.1,1.192,2.951c0.965,0.738,1.135,0,2.44,1.646c-0.284,0.908-0.965,1.873-2.327,1.93 c-1.93,0-0.794-1.305-2.838-1.078c0.057,1.93,1.93,2.667,0.17,3.746c-1.532,0.908-0.568-0.34-1.759,0.284 c-0.057,0.057-0.113,0.057-0.17,0.113s-0.114,0.114-0.17,0.17c-0.227,0.227-0.397,0.397-0.567,0.624 c-2.838-0.795-2.327-0.397-4.029-2.327c-0.624-0.738-0.738-0.908-1.873-0.794c-0.738,0-1.646,0.34-2.327-0.284 c1.362-1.759,2.043-0.057,4.086-1.078c0.341-0.17,2.611-1.873,2.951-2.327c-0.624-1.362-0.057-0.227-0.965-1.135 c-1.986-1.986-2.554-1.532-5.221-1.532c-1.249,1.476-2.213,3.235-3.575,4.029c-0.851,0.454-1.816,0.567-3.121,0 c0.624,2.44,2.157,0.567,2.327,3.519c-1.021-0.227-1.589-0.738-2.611-0.34c-2.27,0.851-0.227,0.794-0.057,1.249 c0,0,1.078,0.113,0,0.567c-1.192,0.397-0.17-0.227-2.44-1.589c0.34-0.567,0.624-1.135,0.965-1.759 c-1.816,0.568-4.143,1.475-6.299,2.554c-2.44,1.192-4.767,2.611-6.129,3.916c-0.057,2.157,0.34,0.738,0.568,3.235 c3.292,0.17,2.213,1.192,5.505,2.1c0.624,0.17,1.476,0.34,2.44,0.511c0,1.589-1.419,2.781-0.454,4.427 c2.327,1.589,2.383-1.532,2.951-3.575c7.832-1.646,4.029-3.916,6.072-8.229c0.284-0.624,0.738-1.305,1.362-2.043 c0.624,0.057,3.065,0.284,3.575,0.397c1.759,0.284,0.794,0.738,2.724,1.589c-0.057,1.532-0.794,3.519,1.476,3.178 c1.759-0.284,1.362-1.702,3.462-2.043c0.965,2.27,0.681,3.235,1.135,5.902c5.108,1.986,2.838,3.632,1.816,7.321 c0.624,0.284,1.305,0.624,1.986,0.851c-0.057,1.078-0.057,2.157-0.113,3.178c-0.113,0-5.448-1.192-5.845-1.249 c0.794-2.327,2.327-2.383,3.292-4.767c-1.93,0.34-1.759,0.908-3.178,1.362c-1.249,0.397-2.213,0.114-3.519,0.284 c-0.454,1.532-0.057-1.362,0.113,0.624c0.057,0.738,0.397,0.057-0.057,0.851c-1.759-0.851-1.192-2.781-4.483-0.965 c0.738,0.34,1.476,0.624,2.27,0.965c-0.113,2.157-0.284,0.511-0.738,1.986c-1.135,3.746,3.064,1.419,4.2,1.192 c-0.284,1.986-1.476,1.986-3.178,2.497c-1.873,0.568-2.44,1.362-3.916,1.702c0-2.327-0.34-0.397,0.681-1.986 c-1.476,0.113-2.611,0.851-4.029,1.419l-0.17,0.057c-1.816,0.738-2.1,0.851-2.327,3.178c-1.532,0.454-3.348,1.021-4.483,2.156 c-1.078,1.022-2.213,2.1-3.292,3.121c-0.114,0.114-0.17,0.17-0.227,0.227c0.397,2.951,0,3.008-1.759,4.143 c-1.475,0.965-1.986,0.908-3.348,2.157c-4.427,3.973-2.327,2.1-2.043,8.967c-0.908-0.284-1.816-0.624-2.667-0.908 c-0.681-4.54,0.227-3.008-3.405-4.086c-3.916-1.135-3.405,0.795-5.051,1.135c-1.816,0.397-2.724-1.646-5.335,0.057 c-1.475,0.965-2.837,1.646-3.178,3.405c-0.397,2.157-1.362,3.689-1.248,6.072c0.227,6.413,6.072,6.753,7.548,3.802 c0.965-1.986,0.738-3.065,3.973-2.554c0.34,1.702-0.17,2.951-0.624,4.256c-0.284,0.795-1.078,2.157-0.681,3.235 c2.327,0.908,2.383-0.851,5.448,0.965c0.227,3.008-1.362,6.47,0.965,8.229c1.816,1.305,1.93,0.113,3.519-0.114 c1.759-0.284,1.476,0.795,3.065,1.021c0.908-1.135,0.511-0.397,1.135-1.816c1.078-2.384,1.362-1.249,3.859-2.724 c0.227-0.113,0.454-0.284,0.738-0.454c0.284,0.624,0.624,1.192,0.908,1.759c1.873-2.213,2.043-0.057,4.2,0.568 c0.908,0.227,1.589,0.34,2.043,0.34c1.532,0,1.703-0.738,3.802,0c0.227,0.114,0.568,0.227,0.851,0.284 c0.057,0,0.113,0.057,0.113,0.113c0.057,0,0.17,0.114,0.17,0.17c0.34,0.397,0.397,0.794,0.681,1.192 c1.476,1.419,2.894,2.781,4.313,4.143c2.781,1.135,2.781-0.965,6.016,2.1c1.362,1.305,0.738,1.249,1.475,2.894 c0.114,0.284,0.34,0.511,0.454,0.851c0.284,0.794,0.624,1.532,0.908,2.27c1.475,1.93,3.462,0.738,5.448,2.667 c1.192,1.135,0.284,0.795,2.894,1.249c4.029,0.681,2.44,0.284,5.505,2.611c1.192,0.851,1.476,0.341,2.554,1.078 c1.532,2.156,1.135,4.427-0.113,6.299c-3.916,5.732-2.611,4.824-2.781,9.364c-0.17,2.611-0.908,8.002-3.348,8.683 c-8.91,2.497-3.746,4.256-6.072,8.626c-0.681,1.249-1.192,1.589-1.589,2.554c-1.759,4.143-0.511,5.561-5.789,4.597 c0.511,0.851,0.17,0.284,0.681,1.078c1.532,2.44,1.816,4.256-1.362,4.824c-2.667,0.511-1.532,0-1.759,2.951 c-1.022,0.397-1.476,0-2.157,0.681c0.908,1.703,0.057,0.284,1.646,1.249c-0.284,0.908-0.511,1.873-0.738,2.724 c-0.624,2.213-0.795,0.17-0.965,2.781c0.908,0.397,1.589,0.34,2.383,1.475c-0.113,0.851-0.965,3.178-1.192,4.483 c1.305,1.986,2.384,4.086,5.164,4.256c-0.057,1.646,0.34,1.475-1.078,1.816c-1.192,0.284-3.746-0.851-4.824-1.419 c-1.816-1.022-3.519-2.327-4.767-4.2c-0.681-1.078-0.397-0.568-1.192-1.532c-1.362-1.759-0.794-2.554-1.192-3.689 c-0.227-0.681-1.305-1.419-0.851-3.632c0.113-0.624,1.078-3.632-0.114-3.121c-1.192,0.511,0.908,0.284-0.738,1.078 c-0.397-1.078-0.738-2.667-0.908-4.029c-0.397-2.27-0.965-2.327-1.192-3.689c-0.511-3.178,0.795-3.519-0.681-8.74 c-0.965-3.462-0.568-11.691-1.646-15.947c-0.17-0.681-0.34-1.249-0.624-1.702c-0.965-1.589-3.292-2.327-5.164-3.632 c-0.114,0-0.17-0.057-0.227-0.113c-2.611-1.816-5.051-8.967-7.094-11.634c-0.965-1.305-1.532-0.738-1.986-2.724 c-0.454-1.873,0.227-1.532,0.851-3.065c-3.178-5.789,4.086-5.732,2.611-12.655c-0.34-1.248-1.249-3.064-2.043-3.632 c-1.476,0.681-0.795,1.022-1.589,2.27c-1.192-0.681-2.384-1.362-3.575-2.043c-3.632-3.519-1.703-0.965-2.724-3.292 c-2.213-4.824-4.313-2.1-6.924-4.54c-1.192-1.021-1.362-2.554-3.348-2.27c-2.951,0.397-1.305,0.851-4.37-0.568 c-2.611-1.192-5.278-2.667-7.094-4.313c-0.114-3.462,0.965-3.121-1.646-6.299V232.108z M182.622,173.088 c0.681-0.568,1.362-1.192,2.043-1.759l0.057,0.114c0.114,0.227,0.227,0.227,0.397,0.851c-1.816,0.454-1.816,0.738-2.497,0.851 V173.088z M277.395,206.003c0.965,0.624,2.157,1.078,2.894,1.759c0.908,2.554-0.227,1.305,1.532,3.973 c1.022,1.419,0.795,2.44,1.873,3.235c1.93-0.227,0.227,0.34,1.475-0.851c0.227-0.227,0.794-0.965,0.965-1.192 c-0.624-0.965-1.021-1.078-1.589-2.043c2.213-2.894,3.178-1.249,3.689,0.681c1.362,5.391,3.519,3.802,5.278,3.575 c1.816-0.227,0.227,0.057,1.873,0.511c1.589,0.397,0.794,0,1.93-0.284c3.746-0.908,0.851,6.299,0.454,7.094 c-2.043,0.227-1.646-0.17-3.065-0.341c-2.951-0.284,0.17,1.759-6.753-0.511c-6.64-2.1-5.221-2.611-6.64,2.043 c-1.532-0.113-1.475-0.397-2.894-0.908c-0.568-0.227-0.851-0.34-1.078-0.397h-0.057c-0.568-0.17,0,0-1.192-1.135 c-1.022-0.965-3.519-1.192-5.335-2.157c0-1.532,1.362-2.781-0.284-4.654c-0.511-0.113-8.059,0.17-9.307,0.397 c-2.44,0.511-2.894,2.611-6.243,2.157c-1.816-0.284-1.192-0.057-1.646-1.135c0.624-0.624,0-0.284,1.021-0.681 c0.397-0.113,0.511-0.113,0.908-0.113c5.448-0.851,2.667-3.575,6.072-6.243c2.837-2.27,0.624,0.057,2.156-2.724 c1.93-0.17,1.646,0.794,3.292,0c1.362-0.681,0.908-1.249,2.837-1.362c1.703,1.703,1.362,2.724,4.143,4.029 c1.078,0.511,1.759,0.795,2.497,1.759c1.249,1.816,0.17,1.249,0.681,2.497c0.113-0.057,0.227-0.114,0.397-0.17l0.057-0.057 c1.475-0.738-0.738-0.965,1.646-3.065c-0.681-0.511-1.192-0.851-1.646-1.135l-0.057-0.057c-1.873-1.078-1.135-0.057-3.518-3.064 c-0.624-0.851-0.795-0.908-0.795-2.213c1.759-0.284,2.043,0.965,3.632,2.327c0.227,0.227,0.454,0.397,0.738,0.568l0.057,0.057 V206.003z M315.247,282.446c-1.532,0.794-1.759,2.611-2.951,3.632c-1.249,1.192-3.008,0.624-3.462,2.781 c-0.624,2.724,1.078,1.873-0.794,5.391c-0.681,1.305-0.738,4.597,0.34,5.562c3.065,2.951,4.767-4.54,5.448-6.924 c0.511-1.702,1.021-3.575,1.589-5.164c0.568-1.703,1.362-3.178-0.17-5.335V282.446z M310.14,203.222 c0.227,3.462,2.327,4.824,3.746,6.413c-0.17,1.646-0.567,1.249-0.511,3.519c1.759,1.249,3.008,2.043,5.278,1.419 c0.34-2.781-0.568-2.667-1.419-4.54c0.113-0.227,0.511-0.511,0.624-0.567c0.624-0.851,0.34,0.34,0.624-1.078 c-1.646-0.568-4.37-2.27-4.597-4.483c0.795-0.624,0.568-0.227,1.192-0.624c1.305-0.851,0.34,0.454,1.078-1.021 c-0.681-1.305-1.419-1.532-3.065-0.908c-0.738,0.284-2.611,1.646-2.951,1.93V203.222z M295.612,204.13 c-1.078-1.532-1.022-2.327-3.178-2.384c-0.851,0.908-1.816,2.27-2.327,3.689c-0.681,2.1-0.227,0.681-0.113,2.951 c-0.057,0.057,0,0.17,0,0.284c-0.057,1.362-0.34-0.284-0.17,1.021c1.646-0.34,0.397-0.34,1.646-0.738 c1.192-0.397,4.767-1.476,5.789-1.192c2.44,0.568,4.483,2.497,7.264,0.568c-0.113-2.497-3.121-3.405-5.164-4.483 c0.114-2.27,0.681-0.454,0.568-2.667c-0.681-0.057-1.135-0.113-1.816,0.114c-1.589,0.738-0.738,0.908-0.738,1.135 c0.057,1.249,1.589-0.227-0.057,1.249c-0.454,0.397-1.192,0.397-1.759,0.454H295.612z M254.184,186.935 c-0.057,3.973,1.305,2.667,2.61,4.597c-0.624,1.702-0.738,0.34-1.362,2.1c-0.624,1.475,0.114,0.17-0.227,2.667 c1.021,0.057,2.554-0.284,3.916-0.397c1.93-0.17,1.873-0.34,2.554-2.327c-1.078-1.305-2.1-2.611-3.235-3.916 c-1.362-2.1,0-0.681-0.397-2.44c-0.34-0.624-0.624-1.249-0.908-1.873c-1.816,0.624-1.646,0.454-2.894,1.589H254.184z"></path> <path style="fill:#49c4b1;" d="M70.711,162.759l31.326-12.599l-5.505-13.393l50.508,21.679l-20.6,51.359l-6.47-15.947l-26.9,10.783 c-11.974,14.812-13.053,33.823-4.086,56.693l-28.886-22.53c-12.826-9.988-18.898-24.97-16.685-41.087 C45.684,181.601,55.558,168.832,70.711,162.759z"></path> <path style="fill:#49c4b1;" d="M93.07,204.698c-48.748,18.671-5.391,55.672-5.391,55.672l1.305,1.021 C80.869,242.266,79.847,223.312,93.07,204.698z"></path> <path style="fill:#49c4b1;" d="M115.089,402.756l-2.327-33.71l-14.415,1.078l36.207-41.371l42.506,35.469l-17.139,1.249 l1.986,28.886c10.385,15.947,28.148,22.87,52.664,21.395l-30.361,20.544c-13.507,9.137-29.624,10.272-44.265,3.121 c-14.642-7.151-23.722-20.487-24.857-36.774V402.756z"></path> <path style="fill:#49c4b1;" d="M161.851,394.414c2.667,52.153,51.302,22.36,51.302,22.36l1.419-0.965 c-20.714,1.816-39.044-3.121-52.664-21.395H161.851z"></path> <path style="fill:#49c4b1;" d="M357.015,434.706l-32.745-8.229l-3.405,14.074l-28.148-47.216l46.876-29.453l-4.143,16.685 l28.091,7.037c18.387-4.937,30.418-19.749,36.661-43.471l10.158,35.242c4.483,15.663,0.624,31.326-10.669,43.073 c-11.293,11.747-26.843,16.231-42.619,12.258H357.015z"></path> <path style="fill:#49c4b1;" d="M363.542,387.66c50.394,13.563,37.115-41.882,37.115-41.882l-0.454-1.589 C395.492,364.45,385.163,380.396,363.542,387.66z"></path> <path style="fill:#49c4b1;" d="M462.173,214.516l-17.933,28.602l12.315,7.605l-53.629,12.201l-13.507-53.686l14.642,9.08 l15.379-24.573c0.965-19.011-9.364-35.015-30.021-48.294l36.604,1.249c16.287,0.567,29.964,9.08,37.625,23.495 C471.31,184.552,470.799,200.725,462.173,214.516z"></path> <path style="fill:#49c4b1;" d="M419.44,193.745c28.489-43.754-28.375-48.238-28.375-48.238l-1.703-0.057 c17.82,10.726,29.794,25.481,30.021,48.294H419.44z"></path> <path style="fill:#49c4b1;" d="M285.226,46.422l21.679,25.935l11.01-9.364l-4.994,54.764l-55.218-3.746l13.166-11.123 l-18.614-22.246c-17.82-6.81-36.207-1.93-55.218,13.62l12.485-34.447c5.562-15.323,17.933-25.708,33.937-28.545 c16.06-2.838,31.269,2.667,41.711,15.152H285.226z"></path> <path style="fill:#49c4b1;" d="M252.254,80.699c-32.802-40.576-54.65,12.031-54.65,12.031l-0.568,1.589 C212.756,80.699,230.462,73.889,252.254,80.699z"></path> </g></svg>`;      
      setIconSource("sustainability-icon", pmc10svg); 
      const pmc11svg = `<svg version="1.1" width="24" id="Uploaded to svgrepo.com" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 32 32" xml:space="preserve" fill="#000000"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <style type="text/css"> .blueprint_een{fill:#49c4b1;} </style> <path class="blueprint_een" d="M23,17h-7v-1h7V17z M23,18h-7v1h7V18z M23,8h-7v1h7V8z M23,6h-7v1h7V6z M10,22.05V19.95 c-1.141-0.232-2-1.24-2-2.45c0-1.209,0.859-2.218,2-2.45V9.95C8.859,9.718,8,8.709,8,7.5C8,6.119,9.119,5,10.5,5S13,6.119,13,7.5 c0,1.209-0.859,2.218-2,2.45v5.101c1.141,0.232,2,1.24,2,2.45c0,1.209-0.859,2.218-2,2.45v2.101c1.141,0.232,2,1.24,2,2.45 c0,1.381-1.119,2.5-2.5,2.5S8,25.881,8,24.5C8,23.291,8.859,22.282,10,22.05z M10.5,23C9.673,23,9,23.673,9,24.5S9.673,26,10.5,26 s1.5-0.673,1.5-1.5S11.327,23,10.5,23z M10.5,9C11.327,9,12,8.327,12,7.5S11.327,6,10.5,6S9,6.673,9,7.5S9.673,9,10.5,9z M10.5,19 c0.827,0,1.5-0.673,1.5-1.5S11.327,16,10.5,16S9,16.673,9,17.5S9.673,19,10.5,19z M29,2v28c0,0.552-0.447,1-1,1H4 c-0.553,0-1-0.448-1-1V2c0-0.552,0.447-1,1-1h24C28.553,1,29,1.448,29,2z M27,3H5v26h22V3z M23,25h-7v1h7V25z M23,23h-7v1h7V23z"></path> </g></svg>`;      
      setIconSource("CR-icon", pmc11svg);
      const pmc12svg = `<svg fill="#49c4b1" width= "24" viewBox="0 0 1920 1920" xmlns="http://www.w3.org/2000/svg" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <g fill-rule="evenodd"> <path fill-rule="nonzero" d="M0 53v1813.33h1386.67v-320H1280v213.34H106.667V159.667H1280V373h106.67V53z"></path> <path d="M1226.67 1439.67c113.33 0 217.48-39.28 299.6-104.96l302.37 302.65c20.82 20.84 54.59 20.85 75.42.04 20.84-20.82 20.86-54.59.04-75.43l-302.41-302.68c65.7-82.12 104.98-186.29 104.98-299.623 0-265.097-214.91-480-480-480-265.1 0-480.003 214.903-480.003 480 0 265.093 214.903 480.003 480.003 480.003Zm0-106.67c206.18 0 373.33-167.15 373.33-373.333 0-206.187-167.15-373.334-373.33-373.334-206.19 0-373.337 167.147-373.337 373.334 0 206.183 167.147 373.333 373.337 373.333Z"></path> </g> </g></svg>`;      
      setIconSource("DR-icon", pmc12svg);
      const pmc13svg = `<svg viewBox="0 0 24 24" width="24" fill="none" xmlns="http://www.w3.org/2000/svg" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path opacity="0.34" d="M5 10H7C9 10 10 9 10 7V5C10 3 9 2 7 2H5C3 2 2 3 2 5V7C2 9 3 10 5 10Z" stroke="#249c4b1" stroke-width="1.5" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"></path> <path d="M17 10H19C21 10 22 9 22 7V5C22 3 21 2 19 2H17C15 2 14 3 14 5V7C14 9 15 10 17 10Z" stroke="#249c4b1" stroke-width="1.5" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"></path> <path opacity="0.34" d="M17 22H19C21 22 22 21 22 19V17C22 15 21 14 19 14H17C15 14 14 15 14 17V19C14 21 15 22 17 22Z" stroke="#249c4b1" stroke-width="1.5" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"></path> <path d="M5 22H7C9 22 10 21 10 19V17C10 15 9 14 7 14H5C3 14 2 15 2 17V19C2 21 3 22 5 22Z" stroke="#249c4b1" stroke-width="1.5" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"></path> </g></svg>`;      
      setIconSource("CM-icon", pmc13svg);
      const pmc14svg = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;      
      setIconSource("files-icon", pmc14svg);
      const pmc15svg = `<svg  width="24" version="1.1" id="Capa_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 18.566 18.566" xml:space="preserve" fill="#000000"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <g> <path style="fill:#49c4b1;" d="M17.705,11.452h-0.73v-3.65c0.929-0.352,1.591-1.247,1.591-2.299c0-1.359-1.102-2.46-2.461-2.46 c-1.044,0-1.936,0.651-2.291,1.568h-1.668V4.522c0-0.475-0.384-0.86-0.859-0.86h-5.2c-0.475,0-0.859,0.386-0.859,0.86V4.61H3.566 V2.924c0-0.475-0.385-0.86-0.859-0.86H1.24c-0.475,0-0.859,0.385-0.859,0.86v5.199c0,0.476,0.385,0.859,0.859,0.859h1.467 c0.474,0,0.859-0.384,0.859-0.859V5.9h1.663v0.089c0,0.475,0.384,0.86,0.859,0.86h5.199c0.476,0,0.859-0.385,0.859-0.86V5.9h1.533 c0.168,1.032,0.977,1.848,2.006,2.025v3.527h-3.179c-0.475,0-0.859,0.385-0.859,0.858v0.089H9.851v-1.955 c0-0.474-0.386-0.859-0.86-0.859H7.524c-0.475,0-0.859,0.386-0.859,0.859v1.955H4.833c-0.284-1.046-1.237-1.815-2.372-1.815 C1.102,10.584,0,11.686,0,13.044c0,1.359,1.102,2.46,2.461,2.46c1.136,0,2.088-0.769,2.372-1.815h1.832v1.955 c0,0.476,0.385,0.859,0.859,0.859h1.467c0.474,0,0.86-0.384,0.86-0.859v-1.955h1.796v0.089c0,0.476,0.385,0.859,0.859,0.859h5.199 c0.476,0,0.86-0.384,0.86-0.859v-1.467C18.565,11.837,18.18,11.452,17.705,11.452z"></path> </g> </g></svg>`;      
      setIconSource("PT-icon", pmc15svg);
      const pmc16svg = `<svg viewBox="0 0 24 24" width="24" fill="none" xmlns="http://www.w3.org/2000/svg" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path opacity="0.34" d="M5 10H7C9 10 10 9 10 7V5C10 3 9 2 7 2H5C3 2 2 3 2 5V7C2 9 3 10 5 10Z" stroke="#249c4b1" stroke-width="1.5" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"></path> <path d="M17 10H19C21 10 22 9 22 7V5C22 3 21 2 19 2H17C15 2 14 3 14 5V7C14 9 15 10 17 10Z" stroke="#249c4b1" stroke-width="1.5" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"></path> <path opacity="0.34" d="M17 22H19C21 22 22 21 22 19V17C22 15 21 14 19 14H17C15 14 14 15 14 17V19C14 21 15 22 17 22Z" stroke="#249c4b1" stroke-width="1.5" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"></path> <path d="M5 22H7C9 22 10 21 10 19V17C10 15 9 14 7 14H5C3 14 2 15 2 17V19C2 21 3 22 5 22Z" stroke="#249c4b1" stroke-width="1.5" stroke-miterlimit="10" stroke-linecap="round" stroke-linejoin="round"></path> </g></svg>`;      
      setIconSource("ContractReview-icon", pmc12svg);
      const pmc17svg = `<svg viewBox="0 0 16 16" width="24" fill="none" xmlns="http://www.w3.org/2000/svg"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path fill-rule="evenodd" clip-rule="evenodd" d="M3.36066 4.98361H4.93443V7.86885H8.60656V9.44262H4.93443V13.1148H8.60656V14.6885H4.14754C3.71296 14.6885 3.36066 14.3362 3.36066 13.9016V4.98361Z" fill="#49c4b1"></path> <path fill-rule="evenodd" clip-rule="evenodd" d="M1 1.83607C1 0.822034 1.82203 0 2.83607 0H10.7049C11.7189 0 12.541 0.822034 12.541 1.83607V3.67213C12.541 4.68616 11.7189 5.5082 10.7049 5.5082H2.83607C1.82203 5.5082 1 4.68616 1 3.67213V1.83607ZM2.83607 1.57377C2.6912 1.57377 2.57377 1.6912 2.57377 1.83607V3.67213C2.57377 3.81699 2.6912 3.93443 2.83607 3.93443H10.7049C10.8498 3.93443 10.9672 3.81699 10.9672 3.67213V1.83607C10.9672 1.6912 10.8498 1.57377 10.7049 1.57377H2.83607Z" fill="#49c4b1"></path> <path fill-rule="evenodd" clip-rule="evenodd" d="M7.81967 8.39344C7.81967 7.37941 8.64171 6.55738 9.65574 6.55738H12.2787C13.2927 6.55738 14.1148 7.37941 14.1148 8.39344V8.91803C14.1148 9.93206 13.2927 10.7541 12.2787 10.7541H9.65574C8.64171 10.7541 7.81967 9.93206 7.81967 8.91803V8.39344ZM9.65574 8.13115C9.51088 8.13115 9.39344 8.24858 9.39344 8.39344V8.91803C9.39344 9.06289 9.51088 9.18033 9.65574 9.18033H12.2787C12.4235 9.18033 12.541 9.06289 12.541 8.91803V8.39344C12.541 8.24858 12.4235 8.13115 12.2787 8.13115H9.65574Z" fill="#49c4b1"></path> <path fill-rule="evenodd" clip-rule="evenodd" d="M7.81967 13.6393C7.81967 12.6253 8.64171 11.8033 9.65574 11.8033H12.2787C13.2927 11.8033 14.1148 12.6253 14.1148 13.6393V14.1639C14.1148 15.178 13.2927 16 12.2787 16H9.65574C8.64171 16 7.81967 15.178 7.81967 14.1639V13.6393ZM9.65574 13.377C9.51088 13.377 9.39344 13.4945 9.39344 13.6393V14.1639C9.39344 14.3088 9.51088 14.4262 9.65574 14.4262H12.2787C12.4235 14.4262 12.541 14.3088 12.541 14.1639V13.6393C12.541 13.4945 12.4235 13.377 12.2787 13.377H9.65574Z" fill="#49c4b1"></path> </g></svg>`;      
      setIconSource("type-icon", pmc17svg);
      const pmc18svg = `<svg viewBox="0 0 32 32" width="24" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" fill="#000000"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"><defs><style>.cls-1{fill:#49c4b1;}</style></defs><title></title><path class="cls-1" d="M15,24H13a2,2,0,0,0-2,2H7a1,1,0,0,1-1-1V22.41l1.29,1.3a1,1,0,0,0,1.42,0,1,1,0,0,0,0-1.42l-3-3a1,1,0,0,0-.33-.21,1,1,0,0,0-.76,0,1,1,0,0,0-.33.21l-3,3a1,1,0,0,0,1.42,1.42L4,22.41V25a3,3,0,0,0,3,3h4a2,2,0,0,0,2,2h2a2,2,0,0,0,2-2V26A2,2,0,0,0,15,24Zm0,4H13V26h2Z"></path><path class="cls-1" d="M4,17H6a2,2,0,0,0,2-2V13a2,2,0,0,0-2-2V7A1,1,0,0,1,7,6H9.59L8.29,7.29a1,1,0,0,0,0,1.42,1,1,0,0,0,1.42,0l3-3a1,1,0,0,0,.21-.33,1,1,0,0,0,0-.76,1,1,0,0,0-.21-.33l-3-3A1,1,0,0,0,8.29,2.71L9.59,4H7A3,3,0,0,0,4,7v4a2,2,0,0,0-2,2v2A2,2,0,0,0,4,17Zm0-4H6v2H4Z"></path><path class="cls-1" d="M28,15H26a2,2,0,0,0-2,2v2a2,2,0,0,0,2,2v4a1,1,0,0,1-1,1H22.41l1.3-1.29a1,1,0,0,0-1.42-1.42l-3,3a1,1,0,0,0-.21.33,1,1,0,0,0,0,.76,1,1,0,0,0,.21.33l3,3a1,1,0,0,0,1.42,0,1,1,0,0,0,0-1.42L22.41,28H25a3,3,0,0,0,3-3V21a2,2,0,0,0,2-2V17A2,2,0,0,0,28,15Zm0,4H26V17h2Z"></path><path class="cls-1" d="M30.71,8.29a1,1,0,0,0-1.42,0L28,9.59V7a3,3,0,0,0-3-3H21a2,2,0,0,0-2-2H17a2,2,0,0,0-2,2V6a2,2,0,0,0,2,2h2a2,2,0,0,0,2-2h4a1,1,0,0,1,1,1V9.59l-1.29-1.3a1,1,0,0,0-1.42,1.42l3,3a1,1,0,0,0,.33.21.94.94,0,0,0,.76,0,1,1,0,0,0,.33-.21l3-3A1,1,0,0,0,30.71,8.29ZM19,6H17V4h2Z"></path></g></svg>`;      
      setIconSource("phase-icon", pmc18svg);

  }    
  catch(err){
      console.log(err.message, err.stack);            
  }
}


function setIconSource(elementId, iconFileName){

  const iconElement = document.getElementById(elementId);

  if (iconElement) {
      iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
  }
}


