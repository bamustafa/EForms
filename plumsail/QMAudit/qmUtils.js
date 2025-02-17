const createLabelCtlr = async function(element, key, value, padding, color){
    let html = `<label style="color: #4e778f; font-weight:bold; ${padding}">
                  <span class="overflow-hidden fd-title-wrap">${key}</span>
                </label>
                    
                <label class="fd-field-control col-sm" style='color: ${color}'><div class="fd-sp-field-text col-form-label">
                   ${value}
                </label>`;
  
    element.after(html);
}

const resolveLatestRev = async function(listname, reference){
  let query = `Reference eq '${reference}' and IsLatest eq 1`;
  let items = await _web.lists.getByTitle(listname).items.select("Id").filter(query).get();

    for (let i = 0; i < items.length; i++) {
        let itemId = items[i].Id;
        
        await _web.lists.getByTitle(listname).items.getById(itemId).update({
            IsLatest: false
        });
    }
}

fd.spSaved(function(result){         
  try
  {
    var itemId = result.Id;
    var webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;

      if(_isNew && _module === 'AUR' )
          result.RedirectUrl = webUrl + "/SitePages/PlumsailForms/AuditReport/Item/EditForm.aspx?item=" + itemId;
      else if(_module === 'AUS'){
        const isGenRev =  JSON.parse(localStorage.getItem('isGenRev'));
        if(isGenRev){
          const newId = JSON.parse(localStorage.getItem('GenRevId'));
          result.RedirectUrl = webUrl + "/SitePages/PlumsailForms/AuditSchedule/Item/EditForm.aspx?item=" + newId;
        }
      }
      else if(_module === 'DCC'){
        debugger
        let formUrl = `${webUrl}/SitePages/PlumsailForms/CKLDWG/Item/EditForm.aspx?item=` + itemId;
        localStorage.setItem('formUrl', formUrl);
        localStorage.setItem('formName', _module);

        let ProjName = localStorage.getItem('ProjName')

        let resultValue = {
          formUrl: formUrl,
          ProjName: ProjName,
          formName: _module
        }

        //window.parent.postMessage(resultValue, '*')
      }
  }
  catch(e){alert(e);}
});

let isItemExist = async function(listname, query, mesg){
  let items = await _web.lists.getByTitle(listname)
                   .items.select("Id").filter(query).get();

  if(items.length > 0)
      disableSubmission(mesg)
  else disableSubmission('')
}

//#region GENERAL
function setButtonState(text, isEnabled){
  if(isEnabled)
   $('span').filter(function() { return $(this).text() == text }).parent().removeAttr('disabled');
  else $('span').filter(function(){ return $(this).text() == text; }).parent().attr("disabled", "disabled");
}

function HideFields(fields, isHide){
for(let i = 0; i < fields.length; i++)
{
  if(isHide)
    $(fd.field(fields[i]).$parent.$el).hide();
  else $(fd.field(fields[i]).$parent.$el).show();
}
}

var setButtons = async function () {
  var status = '';
  if(_module === 'AUR')
    status = fd.field('Status').value;
  fd.toolbar.buttons[0].style = "display: none;";
  fd.toolbar.buttons[1].style = "display: none;";

  if(!_ignoreBtnCreation)
  {
      await setQMButtonActions("Accept", "Submit", false);

      if(_isEdit && _module === 'AUR'){
         await customButtons("Print", "Print", false);

         var summary = fd.field('Summary').value;
         if(summary === null || summary === undefined || summary === '')
           setButtonState('Print', false);
      }
      if(activeTabName == auditReportTab && status === 'Closed')
          setButtonState('Submit', false);
  }
  await customButtons("ChromeClose", "Close", false);
}

var setNewButtons = async function () {
  fd.toolbar.buttons[0].style = "display: none;";
  fd.toolbar.buttons[1].style = "display: none;";

  await setQMButtonActions("Accept", "Submit");

  if(_module === 'AUS' && _isEdit){
      await setQMButtonActions("Accept", _genRev);
      const isGenRev =  JSON.parse(localStorage.getItem('isGenRev'));
      if(isGenRev){
          $('span').filter(function(){ return $(this).text() == _genRev }).parent().attr("disabled", "disabled");
          localStorage.removeItem('isGenRev')
      }
  }

  await setQMButtonActions("ChromeClose", "Cancel");
}

const setQMButtonActions = async function(icon, text){
  fd.toolbar.buttons.push({
        icon: icon,
        class: 'btn-outline-primary',
        text: text,
        click: async function(){
          debugger;
         if(text == "Close" || text == "Cancel")
             fd.close()

         if(text == _submitText){
            if(_module === 'MAP' && _isNew)
              await submitMAPFunction()
           
            else if(_module === 'AUS')
              await addAUSItem()
            
            else if(_module === 'DCC'){
              localStorage.setItem('ProjName', fd.field('Reference').value);
              checkForAlertBody()
              fd.save()
            }
            else fd.save()
             
         }

         else if(text === _genRev){
            localStorage.setItem('isGenRev', true);
            await makeFormCopy();
            fd.save()
         }
     }
  });
}

const doesUserAllowed = async function(userName){
let _isAllowed = false
await pnp.sp.web.currentUser.get()
   .then(async (user) =>{
    if(user.Title == userName){
      _isAllowed = true;
    }
  });	
return _isAllowed;	
}

function setToolTipMessages(){
  setButtonToolTip(_submitText, submitMesg);
  setButtonToolTip('Close', cancelMesg);
}

function delay(time) {
  return new Promise(function(resolve) { 
      setTimeout(resolve, time)
  });
}

var renderControls = async function(){
  await setFormHeaderTitle();

  if(_module === 'MAP' || _module === 'AUS')
    await setNewButtons()
  else await setButtons();

  setToolTipMessages();
}

var setErrorMesg = async function(inputElement, isCorrect, mesg){
  var errorId = '#cMesg'; 
  if($(errorId).length === 0){
    var htmlContent = "<div id='" + errorId.replace('#','') + "' class='form-text text-danger small'>" + mesg + "</div>";
    $(inputElement).after(htmlContent);
  }
   
  if(!isCorrect){
      $(errorId).html(mesg).attr('style', 'color: rgba(var(--fd-danger-rgb), var(--fd-text-opacity)) !important');
  }
  else{
     $(errorId).remove();
  }
}

function Remove_Preloader(){
  if($('div.dimbackground-curtain').length > 0){
      $('div.dimbackground-curtain').remove();
  }

  if($('#loader').length > 0){
      $('#loader').remove();
  }
  clearInterval(_timeOut1);
}

const checkCarFields = async function(_fields){
  var disableBtn = false;
  for(var i = 0; i < _fields.length; i++){
      let _field = _fields[i];
      var _value = fd.field(_field).value;
      if(_value === undefined || _value === null || _value === ''){
          disableBtn = true;
          break;
      }
  }
  return disableBtn;
}

var adjustDisableOpacity = async function(){
  var element = $('div.fd-editor-overlay');
  var isFound = false;
  if(element.length > 0){
      element.css('opacity', '0.4');
      isFound = true;
      //$('textarea').css('height', '100px');
  }

if(isFound)
 clearInterval(_distimeOut);   
}

var disableSummary = async function(){
  if(activeTabName === auditReportTab)
   fd.field('Summary').disabled =  true;
  $(`button:contains('Generate Audit Summary')`).remove();
}

var renderTabs = async function(tabIndex){
  var reqIndex = 1;
  if(tabIndex !== undefined)
       reqIndex = tabIndex;
  $('ul.nav-tabs li a').each(function(index){
      var element = $(this);
     
       
     if(index === reqIndex){
       element.attr('aria-selected', 'true');
       element.addClass('active');
     }
     else{
       element.attr('aria-selected', 'false');
       element.removeClass('active');
     }
  });

  $("div[role='tabpanel']").each(function(index){
      var element = $(this);
      if(index === reqIndex){
          element.addClass('show active');
          element.css('display', '');
          
      }
      else{
        element.removeClass('show active');
      }
   });
}

//#endregion

//#region TabNames is an OBJECT ARRAY
const enable_Disable_Tabs = async function(TabNames, isDisabled){
	TabNames.map(tabName=>{
	  let susIndex = getTabIndex(tabName.title, tabName.tooltip, isDisabled);
	  let tab = fd.container(tabName.masterTab).tabs[susIndex];
	  tab.disabled = isDisabled;
	})
}
  
function getTabIndex(tabTitle, tabTooltip, isDisabled){
	 
  let index;

	$('div.tabset ul li a').each(function(i, element){
	  var title = $(this).text();
	   if(title.toLowerCase() === tabTitle.toLowerCase()){

		 index = i;

     let parentElement = $(this).parent();

		 if(!isDisabled)
		  try{$(parentElement).tooltipster('destroy');} catch{}

		 else{
		   $(parentElement).attr('title', tabTooltip);
		   $(parentElement).tooltipster({
				delay: 100,
				maxWidth: 350,
				speed: 500,
				interactive: true,
				animation: 'slide', //fade, grow, swing, slide, fall
				trigger: 'hover'
		   });
		}

		 return false;
	   }
	}) 
	return index
}
//#endregion



