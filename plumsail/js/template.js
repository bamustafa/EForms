var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDisplay = false;

var  _currentUser, _formFields = {};
const greenColor = '#5FC9B3', redColor = '#F28B82', yellowColor = '#6c757d';

var onRender = async function (relativeLayoutPath, moduleName, formType) {

    try{

        _layout = relativeLayoutPath;

        _formFields = {
        // Title: fd.field('Title'),
        // StartDate: fd.field('StartDate'),
        // EndDate: fd.field('EndDate'),
        // ApprovalStatus: fd.field('ApprovalStatus'),
        // ReviewedDate: fd.field('ReviewedDate'),

        // Level: fd.field('Level'),
        // GoalCategory: fd.field('GoalCategory'),
        // Overview: fd.field('Overview'),
        // DateFinished: fd.field('DateFinished'),

        // PercentComplete: fd.field('PercentComplete'),
        // Status: fd.field('Status'),
        // Attachments: fd.field('Attachments'),

        // Submit: fd.field('Submit'),
        // Manager: fd.field('Manager'),
        // EmployeeId: fd.field('EmployeeId')
        }

        await loadScripts().then(async () => {

            showPreloader();
            await extractValues(moduleName, formType);
            await setCustomButtons();

            if (_isEdit) {
                await handleEditForm();
                //_HideFormFields([_formFields.Submit], true);
            }
            else if (_isNew)
                await handleNewForm();
            //_HideFormFields([_formFields.Submit, _formFields.Manager, _formFields.EmployeeId], true);

            else if(_isDisplay)
                await handleDisplayForm();
        })
    }
    catch (e){
      showPreloader();
      fd.toolbar.buttons[0].style = "display: none;";
      fd.toolbar.buttons[1].style = "display: none;";
      console.log(e)
    }
    finally{
        hidePreloader();
    }
}

var handleNewForm = async function(){

}

var handleEditForm = async function(){

}

var handleDisplayForm = async function(){

}

//#region General
var loadScripts = async function(withSign){

    const libraryUrls = [
      _layout + '/controls/preloader/jquery.dim-background.min.js',
      _layout + "/plumsail/js/buttons.js",
      _layout + '/plumsail/js/customMessages.js',
      _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
      _layout + '/plumsail/js/preloader.js',
      _layout + '/plumsail/js/commonUtils.js',
      _layout + '/plumsail/js/utilities.js'
    ];

    const cacheBusting = `?t=${Date.now()}`;
      libraryUrls.map(url => {
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`);
        });

    const stylesheetUrls = [
      _layout + '/controls/tooltipster/tooltipster.css',
      _layout + '/plumsail/css/DarTemplate.css' + `?t=${Date.now()}`
    ];

    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}${cacheBusting}">`);
    });
}

var extractValues = async function(moduleName, formType){

    //const startTime = performance.now();
    if($('.text-muted').length > 0)
      $('.text-muted').remove();

      _web = pnp.sp.web;
      _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
      _module = moduleName;
      _formType = formType;
      _webUrl = _spPageContextInfo.siteAbsoluteUrl;
      _siteUrl = new URL(_webUrl).origin;

    if(_formType === 'New')
        _isNew = true;

    else if(_formType === 'Edit'){
        _isEdit = true;
        _itemId = fd.itemId;
    }
    else if(_formType === 'Display')
        _isDisplay = true;

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;

    _currentUser = await pnp.sp.web.currentUser.get();
    //  const elapsedTime = endTime - startTime;
    //  console.log(`extractValues: ${elapsedTime} milliseconds`);
}
//#endregion

//#region Custom Buttons
var setCustomButtons = async function () {

    if (!_isDisplay) {
        fd.toolbar.buttons[0].style = "display: none;";
        await setButtonActions("Accept", submitDefault, `${greenColor}`);
    }
    fd.toolbar.buttons[1].style = "display: none;";
    
    await setButtonActions("ChromeClose", "Cancel", `${yellowColor}`);

    setToolTipMessages();

    //$('span').filter(function(){ return $(this).text() === submitText; }).parent().attr("disabled", "disabled");
}

const setButtonActions = async function(icon, text, bgColor){

    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          style: `background-color: ${bgColor}; color: white;`, //color of font button//
          click: async function() {

            if(text == "Close" || text == "Cancel"){
              showPreloader();
              fd.close();
            }
            else if (text == 'Submit') {
                showPreloader();

                fd.save();
            }
        }
    });
}

function setToolTipMessages(){

  let finalizetMesg = `Click ${submitDefault} for Manager Approval`;

  setButtonCustomToolTip(submitDefault, finalizetMesg);
  setButtonCustomToolTip('Close', closeMesg);

	if($('p').find('small').length > 0)
    $('p').find('small').remove();
}

function formatingButtonsBar(titelValue){

    $('div.ms-compositeHeader').remove();
    $('i.ms-Icon--PDF').remove();

    let toolbarElements = document.querySelectorAll('.fd-toolbar-primary-commands');
    toolbarElements.forEach(function(toolbar) {
        toolbar.style.display = "flex";
        toolbar.style.justifyContent = "flex-end";
    });

    let commandBarElement = document.querySelectorAll('[aria-label="Command Bar."]');
    commandBarElement.forEach(function(element) {
        element.style.paddingTop = "16px";
    }); 

    document.querySelector('.col-sm-12').style.setProperty('padding-top', '0px', 'important'); 
    $('.col-sm-12').attr("style", "display: block !important;justify-content:end;");   
    $('.fd-grid.container-fluid').attr("style", "margin-top: -15px !important; padding: 10px;");
    const marginTopValue = _webUrl.includes("db-sp") ? "-22px" : "-10px";
    $('.fd-form-container.container-fluid').css("margin-top", marginTopValue);   

    const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';
    const linkElement = `<a href="${_spPageContextInfo.webAbsoluteUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">
                          <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">${titelValue}</a>`;
    $('span.o365cs-nav-brandingText').html(linkElement);

    $('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');

    $('.border-title').each(function() {
        $(this).css({          
            'margin-top': '-35px', /* Adjust the position to sit on the border */
            'margin-left': '20px', /* Align with the content */            
        });
    });
}

function setIconSource(elementId, iconFileName) {

  const iconElement = document.getElementById(elementId);

  if (iconElement) {
      iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
  }
}
//#endregion

fd.spSaved(async function(result) {

  try {
        // _itemId = result.Id;
        // let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${_itemId}</Value></Eq></Where>`;

        // if(_isNew)
        //     await _sendEmail(_module, 'PMG_New_Email', query, '', 'PMG_New', '', _currentUser);
  } catch(e) {
      console.log(e);
  }
});