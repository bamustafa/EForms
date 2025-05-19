var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDisplay = false, _isPM = false;

var _manager = '', _employeeId = '', _currentUser, _formFields = {};
const greenColor = '#5FC9B3', redColor = '#F28B82', yellowColor = '#6c757d';

var onRender = async function (relativeLayoutPath, moduleName, formType){
    try{

      _layout = relativeLayoutPath;

      _formFields = {
        Title: fd.field('Title'),
        StartDate: fd.field('StartDate'),
        EndDate: fd.field('EndDate'),
        ApprovalStatus: fd.field('ApprovalStatus'),
        ReviewedDate: fd.field('ReviewedDate'),

        Level: fd.field('Level'),
        GoalCategory: fd.field('GoalCategory'),
        Overview: fd.field('Overview'),
        DateFinished: fd.field('DateFinished'),

        PercentComplete: fd.field('PercentComplete'),
        Status: fd.field('Status'),
        Attachments: fd.field('Attachments'),

        Submit: fd.field('Submit'),
        Manager: fd.field('Manager'),
        EmployeeId: fd.field('EmployeeId')
    }

      await loadScripts().then(async ()=>{
        showPreloader();
        await extractValues(moduleName, formType);
        await setCustomButtons();

        if(_isEdit){
           await handleEditForm();
           _HideFormFields([_formFields.Submit], true);
        }
        else if(_isNew)
           _HideFormFields([_formFields.Submit, _formFields.Manager, _formFields.EmployeeId], true);

        if(_isDisplay){
            await handleDisplayForm();
        }
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

var Goal_newForm = async function(){

  try {

      const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;

      const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;

      setIconSource("overview-icon", svguserinfo);
      setIconSource("attachment-icon", svgattachment);

      formatingButtonsBar('Human Resources: Performance Management Goals');
  }
  catch(err){
      console.log(err.message, err.stack);
  }
}

var handleEditForm = async function(){

    const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;

    const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;

    setIconSource("overview-icon", svguserinfo);
    setIconSource("attachment-icon", svgattachment);

    formatingButtonsBar('Human Resources: Performance Management Goals');

    let arrayFields = [_formFields.Title, _formFields.StartDate, _formFields.EndDate, _formFields.Level, _formFields.GoalCategory, _formFields.DateFinished, _formFields.PercentComplete, _formFields.Status, _formFields.Attachments];

    $(_formFields.ApprovalStatus.$parent.$el).hide();
    $(_formFields.ReviewedDate.$parent.$el).hide();

    _HideFormFields([_formFields.Submit, _formFields.EmployeeId], true);
    
    if(_formFields.ApprovalStatus.value === 'Approved'){
        $('span').filter(function(){ return $(this).text() === submitDefault; }).parent().attr("disabled", "disabled");
        setPSErrorMesg('PM has Approved this Goal');
        setPSHeaderMessage('Current Status:');
        $(_formFields.ApprovalStatus.$parent.$el).show();
        fd.field('ApprovalStatus').disabled = true;
        $(_formFields.ReviewedDate.$parent.$el).show();
        fd.field('ReviewedDate').disabled = true;
        _DisableFormFields(arrayFields, true);
        disableRichTextField(_formFields.Overview.title);
        return;
    }
    else{
        let isReject = _formFields.ApprovalStatus.value  === 'Rejected' ? true : false;
        if(!isReject){

            if(!_isPM){
                $('span').filter(function(){ return $(this).text() === submitDefault; }).parent().attr("disabled", "disabled");

                let status = _formFields.ApprovalStatus.value;
                let errorMessage = 'Waiting for PM Approval'
                if(status === 'Approved')
                errorMessage = `PM has ${status}ed this Goal`
                else if(status === 'Rejected')
                errorMessage = `PM has ${status}ed this Goal`

                setPSErrorMesg(errorMessage);
                setPSHeaderMessage('Current Status:');
            }
        }
        else{
        $(_formFields.ApprovalStatus.$parent.$el).show();
        fd.field('ApprovalStatus').disabled = true;
        $(_formFields.ReviewedDate.$parent.$el).show();
        fd.field('ReviewedDate').disabled = true;
        }

        if (!isReject) {
            _DisableFormFields(arrayFields, true);
            disableRichTextField(_formFields.Overview.title);
        }

        if (_isPM) {
            _DisableFormFields(arrayFields, true);
            disableRichTextField(_formFields.Overview.title);
        }
    }
}

var handleDisplayForm = async function(){

  const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;

  const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;

  setIconSource("overview-icon", svguserinfo);
  setIconSource("attachment-icon", svgattachment);

  formatingButtonsBar('Human Resources: Performance Management Goals');
}

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

      if(_formType === 'New'){
        _isNew = true;
        await Goal_newForm();
      }

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
    let serviceUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeeManager&Email=${_currentUser.Email}`;
    let response = await getRestfulResult(serviceUrl)

    if(response && response.length > 0){
         _manager = response[0].Email;
         _manager = 'ali.hsleiman@dar.com';

         if(_currentUser.Email.toLowerCase() === _manager.toLowerCase())
            _isPM = true;

         serviceUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeeIdFromUPN&Email=${_currentUser.Email}`;
         let response1 = await getRestfulResult(serviceUrl)
         if(response1)
          _employeeId = response1.employeeId;

         if(formType === 'New'){
          fd.field('EmployeeId').value = _employeeId;
          fd.field('Manager').value = _manager;
          fd.field('Submit').value = true;
         }
    }

    console.log(_isPM)

    //  const endTime = performance.now();
    //  const elapsedTime = endTime - startTime;
    //  console.log(`extractValues: ${elapsedTime} milliseconds`);
}

var setCustomButtons = async function () {

    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

    if(_isNew)
      await setButtonActions("Accept", submitDefault, `${greenColor}`);

    else if (_isEdit && _isPM && _formFields.ApprovalStatus.value === ''){
      fd.toolbar.buttons.push({
        icon: 'Accept',
        class: 'btn-outline-primary',
        text: 'Approve',
        style: `background-color:${greenColor}; color:white;`,
          click: async function () {
              if (fd.isValid) {
                  showPreloader();
                  $(_formFields.ApprovalStatus.$parent.$el).show();
                  $(_formFields.ReviewedDate.$parent.$el).show();
                  fd.field('ApprovalStatus').value = 'Approved';
                  fd.field('ReviewedDate').value = new Date();
                  $(_formFields.ApprovalStatus.$parent.$el).hide();
                  $(_formFields.ReviewedDate.$parent.$el).hide();
                  await fetchResultToDynamics().then(() => { fd.save(); })
              }
            }
      });

      fd.toolbar.buttons.push({
        icon: 'Cancel',
        class: 'btn-outline-primary',
        text: 'Reject',
        style: `background-color:${redColor}; color:white;`,
          click: async function () {
              if (fd.isValid) {
                  showPreloader();
                  $(_formFields.ApprovalStatus.$parent.$el).show();
                  $(_formFields.ReviewedDate.$parent.$el).show();
                  fd.field('ApprovalStatus').value = 'Rejected';
                  fd.field('ReviewedDate').value = new Date();
                  $(_formFields.ApprovalStatus.$parent.$el).hide();
                  $(_formFields.ReviewedDate.$parent.$el).hide();
                  fd.save();
              }
            }
      });
    }
    else if(_isEdit && !_isPM && _formFields.ApprovalStatus.value === 'Rejected'){
      await setButtonActions("Accept", submitDefault, `${greenColor}`);
    }

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
            else if(text == submitDefault){
                //if (confirm('Are you sure you want to Submit?')){
                if (fd.isValid) {
                    if (_isEdit && _formFields.ApprovalStatus.value === 'Approve')
                        await fetchResultToDynamics().then(() => { fd.save(); })
                    else {
                        showPreloader();
                        if (_isEdit && !_isPM && _formFields.ApprovalStatus.value === 'Rejected') {
                            fd.field('ApprovalStatus').value = '';
                            fd.field('ReviewedDate').value = '';
                        }
                        fd.save();
                    }
                    //}
                }
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
  }) ;

  document.querySelector('.col-sm-12').style.setProperty('padding-top', '0px', 'important');

  const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';
  const linkElement = `<a href="${_spPageContextInfo.webAbsoluteUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">
                          <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">${titelValue}</a>`;
  $('span.o365cs-nav-brandingText').html(linkElement);

  $('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');

  $('.fd-form p').css({
      'margin-top': '0',
      'margin-bottom': '1rem',
      'display': 'none'
  });
}

function setIconSource(elementId, iconFileName) {

  const iconElement = document.getElementById(elementId);

  if (iconElement) {
      iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
  }
}

fd.spSaved(async function(result) {

  try {
        _itemId = result.Id;
        let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${_itemId}</Value></Eq></Where>`;

        if(_isNew)
           await updateItem(_itemId, query, 'New', '');

        else if(_isEdit){
          let status = _formFields.ApprovalStatus.value;

          if(status === 'Approved')
            await updateItem(_itemId, query, 'Edit', status)
          else if(status === 'Rejected')
            await updateItem(_itemId, query, 'Edit', status)
          else
            await updateItem(_itemId, query, 'New', '');
        }
  } catch(e) {
      console.log(e);
  }
});

var getRestfulResult = async function(serviceUrl){

  try {
      const response = await Promise.race([
        fetch(serviceUrl, {
            method: 'GET',
            headers: { 'Content-Type': 'application/json' },
        }),
        new Promise((_, reject) =>
            setTimeout(() => reject(new Error('Request timed out')), 10000) // 5 seconds timeout
        )
      ]);
      return response.json();
  } catch (error) {
    console.error('Error fetching data:', error.message);
    throw error; // Rethrow or handle as needed
    showPreloader();
  }
}

const updateItem = async function(formItemId, query, operation, status){

  if(operation === 'New'){
    if(_manager){
      await _sendEmail(_module, 'PMG_New_Email', query, '', 'PMG_New', '', _currentUser);
    }
  }

  else if(operation === 'Edit' && status === 'Rejected'){

    if(!_isPM){
      await _sendEmail(_module, 'PMG_New_Email', query, '', 'PMG_New', '', _currentUser);
    }
    else{
      await _sendEmail(_module, 'PMG_Reject_Email', query, '', 'PMG_Reject', '', _currentUser);
    }
  }

  else if(operation === 'Edit' && status === 'Approved'){
    await _sendEmail(_module, 'PMG_Approve_Email', query, '', 'PMG_Approve', '', _currentUser);
  }
}

const fetchResultToDynamics = async function(){
  let apiUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=PostEmployeeGoals`

  let overview = _formFields.Overview.value;
  const parser = new DOMParser();
  const doc = parser.parseFromString(overview, 'text/html');
  overview = doc.body.textContent.trim();

   let dateFinished = ''
   if(!isNullOrEmpty(_formFields.DateFinished.value)){
     let date = new Date(_formFields.DateFinished.value);
     dateFinished = date.toISOString().split('T')[0];
   }

   let date1 = new Date(_formFields.StartDate.value);
   let dateStart = date1.toISOString().split('T')[0];

   let date2 = new Date(_formFields.EndDate.value);
   let dateEnd = date2.toISOString().split('T')[0];

  let updateData =  {
    PersonnelNumber: !isNullOrEmpty(_employeeId) ? _employeeId : '',
    GoalHeadingId: _formFields.GoalCategory.value,
    Status: !isNullOrEmpty(_formFields.Status.value) ? _formFields.Status.value : '',

     DateFinished: dateFinished,
     StartDate: dateStart, //_formFields.StartDate.value,
     EndDate: dateEnd, //_formFields.EndDate.value,

    PercentComplete: !isNullOrEmpty(_formFields.PercentComplete.value) ? _formFields.PercentComplete.value : 0,
    Description: _formFields.Title.value,

    GoalLevel: _formFields.Level.value,
    Overview: overview
  };

  const response = await fetch(apiUrl, {
    method: "POST", // or "PATCH" depending on the API
    headers: {
        "Content-Type": "application/json",
        //"Authorization": "Bearer YOUR_ACCESS_TOKEN", // Include if the API requires authentication
    },
    body: JSON.stringify(updateData),
  });
  console.log(response)
}

function disableRichTextField(fieldname){

	let elem = $(fd.field(fieldname).$el).find('.k-editor tr');

	elem.each(function(index, element){

	 if(index === 0)
		$(element).remove()

	 else if(index === 1){

		let iframe = $(element).find('iframe');

		if(iframe.length > 0){

			let content = iframe.contents();
			let divElement = content.find('div');

			var lblElement = $('<label>', {
			  for: 'inputField',
			}).html(divElement.html());

			if(divElement.length === 0){
				lblElement = $('<label>', {
					for: 'inputField',
				  }).html(content[0].activeElement.innerHTML);
			}

			lblElement.css({
				'padding-top': '6px',
				'padding-bottom': '6px',
				'padding-left': '12px',
				'background-color': '#e9ecef',
				'width': '100%',
				'border-radius': '4px'
			});

			let tblElement = iframe.parent().parent().parent().parent();
			tblElement.parent().append(lblElement);
			tblElement.remove();
		}
	   }
	})
}