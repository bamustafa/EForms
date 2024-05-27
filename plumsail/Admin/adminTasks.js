var _web, _webUrl, _siteUrl, _list, _layout, _module = '', _formType = '', activeTabName;
var _isNew = false, _isEdit = false, _isDisplay = false, _isMain = true, _isLead = false, _isPart = false, _isSiteAdmin = false;


var onRender = async function (relativeLayoutPath, moduleName, formType){
    _layout = relativeLayoutPath;
    await loadScripts();
    await extractValues(moduleName, formType);

    const isUserAllowed = await IsUserInGroup('PowerUser');
    if(!isUserAllowed  && !_isSiteAdmin){
       alert('Permission Denied. You are not part of PowerUser group. Contact Admin')
       fd.close();;
    }

    await setButtons();
}


//#region GENERAL
var loadScripts = async function(){
    const libraryUrls = [
      _layout + '/controls/preloader/jquery.dim-background.min.js',
      _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
      _layout + '/plumsail/js/preloader.js',
      _layout + '/plumsail/js/customMessages.js',
      _layout + '/plumsail/js/commonUtils.js'
    ];
  
    const cacheBusting = `?v=${Date.now()}`;
      libraryUrls.map(url => { 
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
      });
        
    const stylesheetUrls = [
         _layout + '/controls/tooltipster/tooltipster.css',
         _layout + '/plumsail/css/CssStyle.css'
        ];
  
  
    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
    });
}

var extractValues = async function(moduleName, formType){
    if($('.text-muted').length > 0)
      $('.text-muted').remove();
    
    _module = moduleName;
    _formType = formType;
    _web = pnp.sp.web;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl; //https://db-sp.darbeirut.com/sites/QM-eAudit
    _siteUrl = new URL(_webUrl).origin;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    
    activeTabName = $('a.active').text();
    $('ul.nav-tabs li a').on('click', async function(element) {
        activeTabName = $(this).text();
        if(activeTabName !== 'Manage Users'){
            fd.field('Groups').required = false;
            fd.field('UserPicker').required = false;
        }
    });

   if(_formType === 'New'){
    clearStoragedFields(fd.spForm.fields);
    _isNew = true;
   }
   else if(_formType === 'Edit')
    _isEdit = true;
   else _isDisplay = true;

   var groups  = await getParameter('SPGroupName');
   if(groups !== ''){
    let field = fd.field('Groups');
     let splitGroups = groups.split(',');
     splitGroups = splitGroups.sort();
     field.options = splitGroups;

     field.required = true;
     fd.field('UserPicker').required = true;
   }
}

function setToolTipMessages(){
    setButtonCustomToolTip('Add', 'Add user to the selected group');
    setButtonCustomToolTip('Remove', 'remove user to the selected group');
    setButtonCustomToolTip('Submit', submitMesg);
    setButtonCustomToolTip('Cancel', cancelMesg);
}

var setButtons = async function () {
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";


    if(activeTabName === 'Manage Users'){
        await setButtonActions("Accept", "Add");
        await setButtonActions("ChromeClose", "Remove");
        await setButtonActions("ChromeClose", "Cancel");
    }

    setToolTipMessages();
}

const setButtonActions = async function(icon, text){
    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          click: async function() {
           if(text == "Close" || text == "Cancel")
               fd.close();
           else if( fd.isValid && (text == "Add" || text == "Remove")){
             preloader();
             await addRemoveUsers(text);

             let mesg = 'users added successfully';
             if(text == "Remove") 
               mesg = 'users removed successfully';

             alert(mesg)
             fd.close();
           }
          } 
    });
}
//#endregion


const addRemoveUsers = async function(transaction){
   let users = fd.field('UserPicker').value;

   const spGroups = await _web.siteGroups();
   let selectedGroup = fd.field('Groups').value;
   const group = spGroups.find(g => g.Title === selectedGroup);
   let groupId = group.Id;

   const userOperations = users.map(async user =>{
    let userLogin = user.Description;
      if(transaction === 'Add')
        await _web.siteGroups.getById(groupId).users.add(userLogin);
      else await _web.siteGroups.getById(groupId).users.removeByLoginName(userLogin);//remove
    });
    await Promise.all(userOperations);
}