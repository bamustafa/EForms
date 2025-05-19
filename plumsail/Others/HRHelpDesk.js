var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDisplay = false;

var  _currentUser, _formFields = {};
const greenColor = '#5FC9B3', redColor = '#F28B82', yellowColor = '#6c757d';

var onRender = async function (relativeLayoutPath, moduleName, formType){
    try{

        _layout = relativeLayoutPath; 
      
        _formFields = {
            Title: fd.field('Title'),
            AffectedUser: fd.field('AffectedUser'),
            RequestType: fd.field('RequestType'),
            EmployeeID: fd.field('EmployeeID'),
            Description: fd.field('Body'), 
            
            Status: fd.field('Status'),
            Priority: fd.field('Priority'),
            AssignedTo: fd.field('AssignedTo'),
            Comments: fd.field('Comments'),
            StartDate: fd.field('StartDate'), 

            DueDate: fd.field('DueDate'),
            DateAssigned: fd.field('DateAssigned'),
            DateEngaged: fd.field('DateEngaged'),
            DateResolved: fd.field('DateResolved'),
            DateClosed: fd.field('DateClosed'),
            Attachments: fd.field('Attachments')
        }

        await loadScripts().then(async ()=>{
            
            showPreloader();
            await extractValues(moduleName, formType);
            await setCustomButtons();
            
            formatingButtonsBar('Human Resources: Service Desk');           

            if (_isEdit) 
                await handleEditForm();                
            
            else if (_isNew)
                await handleNewForm();            

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

var handleNewForm = async function () {    

    const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
 
    setIconSource("overview-icon", svguserinfo); 

    setPSHeaderMessage('', '-25px', '30px');
    setPSErrorMesg(`Please fill in the below fields.`);

    _formFields.RequestType.ready().then(function() {      
        _formFields.RequestType.orderBy = { field: 'Title', desc: false };
        _formFields.RequestType.refresh();
  	});
    
    _currentUser = await pnp.sp.web.currentUser.get();  

    await fillEMpID(_currentUser.Email);

    fd.field('AffectedUser').$on('change', async function (value) { 

        if (value != null) {
            if (value.EntityData != null && value.EntityData.Email) {
                var AffectedEMail = value.EntityData.Email;
                await fillEMpID(AffectedEMail);
            }    
        }            
    });
}

var handleEditForm = async function () {
    
    const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
 
    setIconSource("overview-icon", svguserinfo);  

    //setPSHeaderMessage(''); 
    setPSHeaderMessage('', '-25px', '30px');

    let arrayFields = [_formFields.Title, _formFields.AffectedUser, _formFields.RequestType, _formFields.EmployeeID, _formFields.DateAssigned,
    _formFields.DateEngaged, _formFields.DateResolved, _formFields.DateClosed];
    
    _DisableFormFields(arrayFields, true);
    //disableRichTextField('Body');
    disableRichTextFieldColumn(_formFields.Description);

    let status = _formFields.Status.value;
    //setPSErrorMesg(status);
    //setPSHeaderMessage('Current Status:');

    if (status === 'Closed') {
        _DisableFormFields([_formFields.Status, _formFields.Priority, _formFields.AssignedTo, _formFields.StartDate, _formFields.DueDate], true);
        $('span').filter(function () { return $(this).text() === submitDefault; }).parent().attr("disabled", "disabled");
        //disableRichTextField('Comments');
        disableRichTextFieldColumn(_formFields.Comments);

        _DisableFormFields( _formFields.Attachments, true);

        setPSErrorMesg(`Your request has been reviewed.`);
    }
    else {
        setPSErrorMesg(`Please review the below request.`);
    }
    
}

var handleDisplayForm = async function () {
    
    const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
 
    setIconSource("overview-icon", svguserinfo);

    //setPSHeaderMessage('');
    setPSHeaderMessage('', '-25px', '30px');
    setPSErrorMesg(`Form Display`);
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

function disableRichTextFieldColumn(field){

    let elem = $(field.$el);//$(fd.field(fieldname).$el).find('.k-editor tr');

	elem.each(function(index, element){	

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
	})
}
//#endregion

//#region Custom Buttons
var setCustomButtons = async function () {

    if (!_isDisplay)
        fd.toolbar.buttons[0].style = "display: none;"; 
    else
        fd.toolbar.buttons[0].style = `background-color: #5FC9B3; color: white;`; 
    
    fd.toolbar.buttons[1].style = "display: none;";

    if(_isNew || _isEdit)
      await setButtonActions("Accept", submitDefault, `${greenColor}`);        

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
                if (fd.isValid) {
                    showPreloader();
                    fd.save();
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
    }); 

    document.querySelector('.col-sm-12').style.setProperty('padding-top', '0px', 'important'); 
    $('.col-sm-12').attr("style", "display: block !important;justify-content:end;");   
    $('.fd-grid.container-fluid').attr("style", "margin-top: -15px !important; padding: 10px;");
    $('.fd-form-container.container-fluid').attr("style", "margin-top: -10px !important;");   

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

async function fillEMpID(EmpEmail) {
    
    if (EmpEmail) {
        let serviceUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeeIdFromUPN&Email=${EmpEmail}`;
        let response1 = await getRestfulResult(serviceUrl);

        if (response1)
            fd.field('EmployeeID').value = response1.employeeId;
    }

    _HideFormFields([fd.field('EmployeeID')], true);
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