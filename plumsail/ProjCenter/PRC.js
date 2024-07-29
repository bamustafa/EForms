var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;
var _isSiteAdmin = false, _isMain = false, _isLead = false, _isPart = false, _isNew = false, _isEdit = false, _isDisplay = false
    _isPD = false, _isPM = false, _isQM = false, _isSus = false, _isGIS = false, _isBuilding = false, _isReader = false, _isUserAllowed = true, _hideSubmit = false, 
    _isPROLE = false, _isGIS = false;

var projectNo = '', projectTitle = '', CurrentUser;

var Category = 'Category', SubCategory = 'SubCategory', MatrixFields = 'Matrix Fields', MTDs = 'MTDs', Questions = 'Questions', ContractReview = 'Contract Review',
    OtherPartiesFirms = 'Other Parties Firms', Sustainability = 'Sustainability', Level = 'Levels', Roles = 'Roles', ProjectRoles = 'Project Roles', 
    GISLocation = 'GIS Location';
var formFields = {};

let masterTab = 'Tabs1'
var tabs = [
            {
              masterTab: masterTab,
              title: 'Sustainability',
              tooltip: 'Sustainability is enabled for Buildings Category only'
            },
            {
              masterTab: masterTab,
              title: 'Background Info',
              tooltip: 'Background Info is enabled for Buildings Category only'
            }
];

var appBarItems;

const blueColor = '#6ca9d5', greenColor = '#5FC9B3', redColor = '#F28B82'; // Buttons Colors
const submitText = 'Finalize'
var onRender = async function (relativeLayoutPath, moduleName, formType) {
 
  _layout = relativeLayoutPath;

    try{
      if(moduleName === 'PINT')
        _isMain = true;

      await loadScripts()
      await extractValues(moduleName, formType);

      if (!_isUserAllowed){
        alert(isAllowedUserMesg);
        fd.close();
      }

      
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
      
      if(!_isDisplay)
        await setCustomButtons();
    }
    catch (e){
      fd.toolbar.buttons[0].style = "display: none;";
      console.log(e);
    }

    //setPreviewForm();
    //await drawChart();
    //await hideLicense();
}

//#region GENERAL
var loadScripts = async function(){

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

      _layout + '/plumsail/ProjCenter/utils.js',
      _layout + '/plumsail/ProjCenter/PIR/metrics.js',
      _layout + '/plumsail/ProjCenter/PIR/backgroundInfo.js',
      _layout + '/plumsail/ProjCenter/PIR/designProcess.js',
      _layout + '/plumsail/ProjCenter/PIR/Sustainability.js',
      _layout + '/plumsail/ProjCenter/PIR/contractReview.js',

      _layout + '/plumsail/ProjCenter/features/projroles.js',
      _layout + '/plumsail/ProjCenter/features/gis.js'
    ];

    const cacheBusting = `?v=${Date.now()}`;
      libraryUrls.map(url => {
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`);
        });

    const stylesheetUrls = [
      _layout + '/controls/tooltipster/tooltipster.css',

      _layout + '/controls/vanillaSelectBox/vanillaSelectBox.css',

      _layout + '/plumsail/css/pmisStyle.css' + `?v=${Date.now()}`
      //_layout + '/plumsail/css/FNC.css'
      //_layout + '/controls/jqwidgets/styles/jqx.fluent.scss'
    ];

    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
    });
}

var extractValues = async function(moduleName, formType){

  if($('.text-muted').length > 0)
    $('.text-muted').remove();

    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;

    if(_isMain)
      await getCurrentUserRole();

    if(_formType === 'New'){
        clearStoragedFields(fd.spForm.fields);
        _isNew = true;
    }
    else if(_formType === 'Edit'){
        _isEdit = true;
        _itemId = fd.itemId;
    }
    else if(_formType === 'Display')
      _isDisplay = true;

    if(_module === 'PINT'){
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
          redirectUrl: `${_webUrl}/SitePages/PlumsailForms/GISLocation/Item/NewForm.aspx`,
          viewBox: '0 0 24 24',
          iconTitle: 'GIS',
          tooltip: 'Set GIS Location',
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
}

var setCustomButtons = async function () {

    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";
    //fd.toolbar.buttons[2].style = "display: none;";

    if(!_hideSubmit){
      if(_module !== 'GIS')
        await setButtonActions("Save", "Save", `${blueColor}`);  

      if(!_isPROLE){
       await setButtonActions("Accept", submitText, `${greenColor}`);
      }
    }

    await setButtonActions("ChromeClose", "Close", `${redColor}`);
    setToolTipMessages();
    
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
                fd.close();
                if(_isPROLE || _isGIS)
                  window.close();
              }
            }
            else if(text == "Save"){
               fd.validators.length = 0;
               if(_isPROLE) // || _isGIS){
                //debugger;
                //if(_isGIS)
                 // await InsertGISItem(gisItem, isFound);
                window.close();
               //}
               else{
                let {items} = await getTabFields();
                let itemMetaInfo = await getFieldsData(MatrixFields);
                let query = `Title eq '${projectNo}'`;

                await addUpdate(MatrixFields, query, itemMetaInfo);
                await insertItemsInBulk(items, ContractReview, ['Question', 'Answer', 'Comments']);
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
                   let isValid = fd.isValid
                   if(isValid){
                    debugger
                      let isGISFound = JSON.parse(localStorage.getItem('isGISFound'));
                      await InsertGISItem(gisItem, isGISFound).then(async ()=>{
                        if(isGISFound)
                          await sendGISApproval();
                      }).then(()=>{
                        //window.close();
                      });
                   }
                }

                else{
                    let {mesg,items} = await getTabFields();

                    if(mesg === ''){
                      if (confirm('Are you sure you want to Submit?')){
                        let itemMetaInfo = await getFieldsData(MatrixFields);
                        let query = `Title eq '${projectNo}'`;

                        await addUpdate(MatrixFields, query, itemMetaInfo);
                        await insertItemsInBulk(items, ContractReview, ['Question', 'Answer', 'Comments']);
                        fd.save();
                      }
                    }
                }         
            }
        }
    });
}

function setToolTipMessages(){

  if(_module === 'GIS')
    finalizetMesg = 'Click Finalize for GIS Approval';

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
//#endregion

//PINT (PROJECT INFO)
var onPINTRender = async function (){

   if(_isEdit)
     localStorage.setItem('MasterId', _itemId);   

    await onMetricsRender()  //metrics.js
    await onBackGroundInfoRender() //backgroundInfo.js

    await handelTables(); //utils.js
    //await filterDataTable_LookupField('firmdt', 'Firm', 'Title');  //utils.js for Other Parties

    await onDesignProcessRender();
    await onSusRender('susdt', 'Goal', 'Title', 'Level', Level);  //Sustainability.js

    await onContReviewRender()  //contractReview.js

   
    await setMainRules() // pcr.js

    handleCascadedTabsView();
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
    await enable_Disable_Tabs(tabs, isDisabled);
    categoryField.disabled = true
  }

  let tabFields = await getEditableTabFields();
}

function handleCascadedTabsView(){
  const containers = document.querySelector('.fd-grid.container-fluid');
  const getTabs = containers.querySelectorAll('.tabset .tabs-top');
  getTabs.forEach((tab, index) => {
    if(index == 1){
      let tabsUL = $(tab).find('.nav-item'); //tab.querySelectorAll('.nav-item'); 
      if(tabsUL.length > 0){
        tabsUL.forEach(navItem => {
          navItem.style.width = '300px';
        }); 
      }
    }   
  });  
}