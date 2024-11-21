var ProjectNameList="", ProjectYear=2010;

var _layout, _module = '', _formType = '', _web, _webUrl, _siteUrl, _itemId;
var hFields = [], projectArr=[] , fields = {};
var delayTime = 100, retryTime = 10, _timeOut, _timeOut1, _pplTimeOut, _distimeOut;

var  _isMain = true, _isLead = false, _isPart = false, _isNew = false, _isEdit = false, _isSentForReview = false, _isSentForApproval = false, _isAurClosed = false, 
     _hideCode = true, disableCode = false, _ignoreBtnCreation = false;
var _header, _main, _footer;

var auditGenSummary = 'Fill Department and start/end Audit date to generate summary';
var carGenSummary = 'Fill all fields above to proceed';

var aurFields = ['Department', 'Procedures', 'StartAuditDate', 'EndAuditDate', 'Title', 'Office', 'AuditRef'];
var aurPeopleFields = ['AuditTeam', 'Contact', 'Distribution', 'PreparedBy'];

var auditTeamGroup = 'Audit Team';
var activeTabName = '';
var auditReportTab = 'Audit Report';
var corrActionTab = 'Corrective Action';

var AuditSchedule = 'Audit Schedule', MasterPlan = 'Master Plan', departmentsList = 'Departments', Scope = 'Procedures', officeProceduresList = 'Office Procedures', 
    activitiesList = 'Activities', DrawingControl = 'Drawing Control', auditReport = 'Audit Report';

var _htLibraryUrl, colsInternal = [], counterTemp = [];

let _isDepartmental = false, _isCompany = false, _isProject = false, _isUnschedule;

var _submitText = 'Submit', _genRev = 'Generate Revision'
var tabs = [
    {
      masterTab: 'Tabs1',
      title: 'Audit Process',
      tooltip: 'Audit Process will be enabled once it is approved by Quality Manager'
    }
];

var onRender = async function (relativeLayoutPath, moduleName, formType){

    try {
        await getGlobalParameters(relativeLayoutPath, moduleName, formType); // GET LAYOUT PATH FROM SOAP SERVICE

        if(_module === 'MAP')
          await onMAPRender();
        else if(_module === 'DCC')
          await onDCCRender();
        else if(_module === 'AUR')
          await onAURender();
        else if(_module === 'AUS')
          await onAUSRender();
       
        await renderControls();
        
        if(_isAurClosed){
            activeTabName = corrActionTab;
            await renderTabs();
            await handleCarTab();
        }

        _timeOut = setInterval(fixEnhancedRichText, 1000);
        _distimeOut = setInterval(adjustDisableOpacity, 1000);
    }
    catch (e) {
        console.log(e);
    }
}

var loadScripts = async function(){
    const libraryUrls = [
      _htLibraryUrl,
      _layout + '/controls/preloader/jquery.dim-background.min.js',
      _layout + "/plumsail/js/buttons.js",
      _layout + '/plumsail/js/customMessages.js',
      _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
      _layout + '/plumsail/js/preloader.js',
      _layout + '/plumsail/js/commonUtils.js',
      _layout + '/plumsail/js/utilities.js',
      
      _layout + '/plumsail/QMAudit/masterPlan.js',
      _layout + '/plumsail/QMAudit/drawingChecklist.js',
      _layout + '/plumsail/QMAudit/auditReport.js',
      _layout + '/plumsail/QMAudit/auditSchedule.js',
      _layout + '/plumsail/QMAudit/ausMP.js',
      _layout + '/plumsail/QMAudit/qmUtils.js',
      _layout + '/controls/vanillaSelectBox/vanillaSelectBox.js',
    ];

    if(_module === 'MAP'){
        Array.prototype.push.apply(libraryUrls, [
            _layout + '/controls/jqwidgets/jqxcore.js',
            _layout + '/controls/jqwidgets/jqxdatetimeinput.js',
            _layout + '/controls/jqwidgets/jqxcalendar.js',
            _layout + '/controls/jqwidgets/jqxtooltip.js',
            _layout + '/controls/jqwidgets/globalization/globalize.js'
        ]);
    }

    const cacheBusting = `?v=${Date.now()}`;
      libraryUrls.map(url => { 
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
        });
        
    const stylesheetUrls = [
      _layout + '/controls/tooltipster/tooltipster.css',
      _layout + '/controls/handsonTable/libs/handsontable.full.min.css',
      _layout + '/plumsail/css/CssStyle.css',
      _layout + '/plumsail/css/partTable.css',
      _layout + '/controls/jqwidgets/styles/jqx.base.css',
      _layout + '/controls/vanillaSelectBox/vanillaSelectBox.css'
      //_layout + '/controls/jqwidgets/styles/jqx.fluent.scss'
    ];
  
    stylesheetUrls.map((item) => {
      var stylesheet = item + `?v=${Date.now()}`
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
    });
}

var getGlobalParameters = async function(relativeLayoutPath, moduleName, formType){
    if($('.text-muted').length > 0)
      $('.text-muted').remove();
    
    _module = moduleName;
    _formType = formType;
    _web = pnp.sp.web;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl; //https://db-sp.darbeirut.com/sites/QM-eAudit
    _siteUrl = new URL(_webUrl).origin;
    _layout = relativeLayoutPath;
    _itemId = fd.itemId;

    await loadScripts();
    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

    if(_module === 'DCC' || _module === 'AUR') 
      activeTabName = $('a.active').text();

   if(_formType === 'New'){
    clearStoragedFields(fd.spForm.fields);
    _isNew = true;
    if(_module === 'DCC'){
        if(_isNew){
          fd.field('IsDrawing').value = true;
          $(fd.field('DrawingAvailability').$parent.$el).hide();
        }
    }
   }
   else if(_formType === 'Edit')
    _isEdit = true;

    // var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GLOBAL_PARAM';
    // var soapContent;
    // soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    //                 '<soap:Body>' +
    //                   '<GLOBAL_PARAM xmlns="http://tempuri.org/" />' +
    //                 '</soap:Body>' +
    //               '</soap:Envelope>';
    // await getSoapRequest('POST', serviceUrl, false, soapContent);

}

// var getSoapRequest = async function(method, serviceUrl, isAsync, soapContent){
// 	var xhr = new XMLHttpRequest(); 
//     xhr.open(method, serviceUrl, isAsync); 
//     xhr.onreadystatechange = async function() 
//     {
//         if (xhr.readyState == 4) 
//         {   
//             try 
//             {
//                 if (xhr.status == 200)
//                 {                
// 					const obj = this.responseText;
// 					var xmlDoc = $.parseXML(this.responseText),
// 					xml = $(xmlDoc);
					
//                     var value= xml.find("GLOBAL_PARAMResult");
//                     if(value.length > 0){
//                         text = value.text();
//                         _layout = value[0].children[0].textContent;
//                     }
//                 }            
//             }
//             catch(err) 
//             {
//                 console.log(err + "\n" + text);             
//             }
//         }
//     }
// 	xhr.setRequestHeader('Content-Type', 'text/xml');
//     xhr.send(soapContent);
// }

// var ensureFunction = async function(funcName, ...params){
//     var isValid = false;
//     var retry = 1;
//     while (!isValid)
//     {
//         try{
//           if(retry >= retryTime) break;
//           if(funcName === 'IsUserInGroup'){
//             var allowed = await IsUserInGroup(...params);
//             isValid = true;
//              return allowed;
//           }
//         }
//         catch{
//           retry++;
//           await delay(delayTime);
//         }
//     }
// }