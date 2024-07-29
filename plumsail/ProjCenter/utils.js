//#region AUTOFIT DATATABLE COLUMNS WIDTH
const handelTables = async function(singleCtl){

  let dts = [], isSingleTable = false;
  let tblName;
  let targetCtls = "div[role='tabpanel'] div.w-100";
  if(singleCtl !== undefined && singleCtl !== null && singleCtl !== ''){
    targetCtls = singleCtl
    isSingleTable = true;
  }

  $(targetCtls).each(function(index, element) {
     tblName = $(element).attr('internal-name');
    if (tblName !== undefined && tblName !== null && tblName !== '') {
        dts.push(tblName);
    }
  });

  if(dts.length > 0){
    dts.map(tblName => {
      fd.control(tblName).ready().then(function(dt) {           
          dt.$on('change', () => handleEvent(tblName, dt)); 
          dt.$on('delete', () => handleEvent(tblName, dt));                  
        }
      )
    });
  }  

  if(isSingleTable){           
    let dt = fd.control(tblName);              
    handleEvent(tblName, dt);
  }
  else{
    $('ul.nav-tabs li a').on('click', function(){
      handleSingleTable()
    });
  }
}

const handleEvent = async function(tblName, dt) {
  setTimeout(async () => { 
    await FixWidget(dt)
    let tblElement = $(`div[internal-name='${tblName}'] table`)
    FixListTabelRows(tblElement); 
  }, 100);
};
  
function FixWidget(dt) {
 
    const clientWidth = dt.$el.clientWidth;   
    const Rwidget = dt.widget;
    const columns = Rwidget.columns;
    const columnsLength = columns.length;
    const baseWidth = clientWidth / columnsLength;
  
    const fixedWidths = {
        '': 30,
        'Status': 120,
        'Firm': 500,
        'BOQ': 120,
        'IsMTD': 70, 
        'Question': 600,
        'Answer': 250,
        '#commands': 60,
        'Goal': 400,
        'Level': 400,

        'ProjRole': 450,
        'EmpName': 450,
        'Created': 350
    };
  
    let remainingWidth = 0;
    columns.forEach((column, index) => {    
  
        const field = column.field !== undefined ? column.field : '';
        let colWidth = 0;
  
        if (fixedWidths.hasOwnProperty(field)){
            colWidth = fixedWidths[field];
            remainingWidth = remainingWidth + (baseWidth - colWidth);
        }
        else if(index == (columnsLength - 1))    
          colWidth = baseWidth + remainingWidth;         
        else colWidth = baseWidth;
  
        dt._columnWidthViewStorage.set(field, colWidth);
        dt._columnWidthEditStorage.set(field, colWidth);
        Rwidget.resizeColumn(column, colWidth);
    });

    const gridContent = dt.$el.querySelector('.k-grid-content.k-auto-scrollable');
    if (gridContent) {
        gridContent.style.overflowX = 'hidden';
    }
  
    let rows = Rwidget._data;
    rows.forEach(row => {
        if (row.Status === 'Approved') {
            const rowElement = $(dt.$el).find('tr[data-uid="' + row.uid + '"]')[0];
            if (rowElement) {
                //rowElement.style.backgroundColor = 'var(--lightgreen)';
                rowElement.style.background = 'linear-gradient(to left, rgb(235, 245, 230), rgb(210, 240, 230), rgb(190, 230, 235), rgb(180, 200, 245), rgb(200, 190, 245), rgb(200, 190, 245), rgb(190, 185, 245))';        
            }
        }
    });
}
  
function FixListTabelRows(tblElement){ 
      
    const columnsToCenter = ['', 'Used for BoQ', 'IsMTD', 'Status'];
    let cols = {};
  
    tblElement.each(function(tblIndex, tbl){
        $(tbl).find('tr').each(function(trIndex, tr) {
          if (trIndex === 0 ){    	
           let childs = tr.children;
             if(childs.length > 0){
              $(childs).each(function(index, ctlr) {
                let colname = $(ctlr).text().trim();
                if (index === 0 || columnsToCenter.includes(colname)){
                  ctlr.style.textAlign = 'center';
                  cols[index] = colname;
                }
              });
            }
          }
        
          $(tr).find('td').each(function(tdIndex, td){
                let $td = $(td);
                let  colIndex = cols[tdIndex];
  
                if (colIndex !== undefined)
                    td.style.textAlign = 'center';
                else{
                    if(_formType !== 'Display'){
                      $td.children().css('whiteSpace', 'nowrap');
                      $td.css('whiteSpace', 'nowrap');
                    }
                }
          });                			
        });
    });    
}

function handleSingleTable(){
  let tabElements = $("div[role='tabpanel'].active div.fd-sp-datatable-wrapper");
  tabElements.each(function() {
      let tabElement = $(this);
      if (tabElement.length > 0) {              
          let tblName = tabElement.attr('internal-name');            
          let dt = fd.control(tblName);              
          handleEvent(tblName, dt);                       
      }
  });
}
//#endregion

//#region SET ERROR ICON ON TAB
let getTabFields = async function(){

  let errorMesg = '';
  $('div.tabset ul li.nav-item img').remove();
  const fields = fd.fields();
  const Tabs = fd.container('Tabs1').tabs;
  let tabArray=[], questResult;

    for(var i = 0; i < Tabs.length; i++){

      let tab = fd.container('Tabs1').tabs[i];
      let isTabExist = tabArray.filter(item=>{ item.tabIndex === i });
        if(isTabExist.length > 0) continue;
      
        let TabTitle = tab.title;
        if(TabTitle === 'Classification & Metrics' || TabTitle === 'Contract Review'){
          
          let mesg = '';
          if(TabTitle === 'Contract Review'){
            let obj = await isQuestionValid();
             mesg = obj.mesg
             questResult = obj.items;
          }
          else mesg = await validateFields();

          if(mesg !== ''){
            errorMesg = mesg;
            let htmlElement =  $(`div.tabset ul li.nav-item a:contains(${TabTitle})`);
            tabArray.push({
              tabIndex: i,
              tabTitle: TabTitle,
              htmlElement: htmlElement
            });
          }
        }

        else{
          for(var j = 0; j < fields.length; j++){
            let field = fields[j];
            let fieldTitle = fields[j].title;
            let isValidField = isDescendant(tab.$el, field.$el);
            let isValueNull = isNullOrEmpty(fd.field(field.internalName).value);
            let htmlElement =  $(`div.tabset ul li.nav-item a:contains(${TabTitle})`);

            if(field.required && isValidField && isValueNull){ 
              errorMesg = 'Error';  
                  tabArray.push({
                    tabIndex: i,
                    tabTitle: TabTitle,
                    htmlElement: htmlElement
                  });
                  setRequiredFieldErrorStyle(field.$el, false)
            }
            else setRequiredFieldErrorStyle(field.$el, true)
          }
        }
    }
  
    for(const tab of tabArray){
      let tabIndex = tab.tabIndex;
      let aTag = tab.htmlElement;

      if (aTag)
        await setTabErrorIcon(tabIndex, aTag);
      else console.error(`No <a> tag found in the tab element at index ${tabIndex}`);
    }
   return {
    mesg: errorMesg,
    items: questResult
   }
}

const isDescendant = function (parent, child) {
  let node = child.parentNode;
  while (node) {
      if (node === parent) {
          return true;
      }

      // Traverse up to the parent
      node = node.parentNode;
  }

  // Go up until the root but couldn't find the `parent`
  return false;
}

let setTabErrorIcon = async function(tabIndex, element){

  let errorIconId = `errorIcon${tabIndex}`;

  if($(`#${errorIconId}`).length === 0){
  
    let imgUrl = `${_webUrl}${_layout}/Images/errorIcon.png`;
    let img = $('<img id="' + errorIconId + '" src="' + imgUrl + '" />').css({
                  "position": "absolute",
                  "right": "5px",
                  "top": "50%",
                  "transform": "translateY(-50%)",
                  "width": "20px",
                  "height": "20px"
              });

     $(element).parent().css("position", "relative").append(img);
   }
  
}
//#endregion

//#region INLINE DATATABLE LOOKUP FIELDS 2 LEVELS OR ONE LEVEL (TO SELECT FROM FIRST FIELD ONCE)
const filterDataTable_LookupField = async function(dtName, fieldOnList, fieldQuery, secondList, secondFieldonTable, result){

let dt = fd.control(dtName);
dt.$on('change', async function(item){
  if(item.type === 'add'){

     removeEdit(dtName);
     let itemId = item.itemId
    
     const listUrl = dt.listRootFolder;
     const tempList = await _web.getList(listUrl).get();
     const listname  = tempList.Title;

     let spItem = await getItem(listname, itemId);
     let val = spItem[`${fieldOnList}Id`];

     if(val === undefined || val === null){
        let list = _web.lists.getByTitle(listname);
        await list.items.getById(itemId).delete()
           .then(()=>{ 
           dt.refresh();
           FixWidget(dt);
          })
      }
      else{
        //send email
        let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
          _sendEmail(_module, `SUS_New_Email|${projectNo}`, query, '', 'SUS_New', '');
      }
     }
   else if(item.type === 'delete'){
    let items = dt._data._selectedItems;
    let _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5];
    let _ListFullUrl = _spPageContextInfo.webAbsoluteUrl + "/Lists/" + _ListInternalName;

    let hrefElement = `<a style='color:#616A76;font-weight: bold;font-size:13px' href='${_ListFullUrl}/EditForm.aspx?ID=${_itemId}'>${projectNo} ${projectTitle}</a>`;
    for(const spItem of items){
      let goal = spItem.Goal.length > 0 ? spItem.Goal[0].lookupValue : '';
      let subject = `Project Sustainability Goal - ${projectNo}`
      let body = htmlEncode(`Please note that sustainability goal ${goal} has been removed for ${hrefElement}`);
      let emailTemplate = `${subject}|${body}`;
      _sendEmail(_module, `${emailTemplate}|${projectNo}`, '', '', 'SUS_New', '');
    }
   }
   setTimeout(() => { 
    fixTableCols(dtName)
   }, 200);
})

dt.$on('edit', function(editData){

  removeEdit(dtName);

  let filterQuery = '';

  if(!fieldQuery.includes(' eq ')){
    let firmOptions = dt.widget.dataSource.data()
    firmOptions.map((option, index)=>{
      let value = option[fieldOnList] !== "" ? option[fieldOnList][0].lookupValue : "";
      if(value !== ""){
        if(index === 0) 
        filterQuery += ` ${fieldQuery} ne '${value}' `
        else filterQuery += filterQuery === '' ? `${fieldQuery} ne '${value}' ` : `and ${fieldQuery} ne '${value}' `
      }
    })
  }
  else filterQuery = fieldQuery;

  editData.field(fieldOnList).filter = filterQuery;
  editData.field(fieldOnList).useCustomFilterOnly = true;

  if(secondFieldonTable !== undefined && secondFieldonTable !== null){
    if (editData.formType === 'New' || editData.formType === 'Edit'){
      setTimeout(() => { 
        if(editData.formType === 'Edit')
          editData.field(fieldOnList).disabled = true;
        editData.field(secondFieldonTable).disabled = true;
        }, 200);
    }
      
    editData.field(fieldOnList).$on('change', async function(value){
    
      if(value !== undefined && value !== null && value !== ''){
        let query = `${fieldOnList}/Id eq '${value.LookupId}'`;
        let options;

        let headerFieldText = secondFieldonTable
        if(secondList === null){
          debugger;
          options = result.filter((item) =>{ 
             if(item.title === value.LookupValue)
              return item.title
          })
          if(headerFieldText === 'EmpName')
            headerFieldText = 'Employee Name'

          if(options.length > 0){
            let users = options[0].groupMembers;
            let titles = users.map(u => u.Title);
            editData.field(secondFieldonTable).widget.setDataSource({data: titles});
          }
        }
        else {
          options = await _web.lists.getByTitle(secondList).items.select(`${fieldOnList}/Id,Title,Id`).filter(query).expand(fieldOnList).getAll();
          await filter_Inline_DataTable_SecondLookupField(dtName, headerFieldText, options)
        }
        editData.field(fieldOnList).disabled = true;
        editData.field(secondFieldonTable).disabled = false;
      }
    })
  }
  fixTableCols(dtName);
});

const getItem = async function(listname, itemId){

  let list = _web.lists.getByTitle(listname);
  return await list.items.getById(itemId).get()
}
}

const filter_Inline_DataTable_SecondLookupField = async function(dtName, fieldonTable, options){

let htmlDT = `div [internal-name='${dtName}']`;

let textHeaders = $(htmlDT).find("tr[role='row'] th");
let colIndex;
if(textHeaders.length > 0){
  textHeaders.each(function(index, element) {
  let headerText = element.innerText.trim()
    if ( headerText === fieldonTable){
      colIndex = index;
      return
    }
  });
}

setTimeout(() => { 
  let elem = $(htmlDT).find("tr[data-role='editable'] td").eq(colIndex); //3
  let input = elem.length > 0 ? $(elem).find('input'): undefined;
  let listId = '';

  if(input.length > 0)
    listId = input.attr("aria-controls");

  if(listId !== undefined && listId !== null && listId !== ''){
    let list = $(`#${listId}`);

    if(list.length > 0){

      let liElement = list.find('li');
      liElement.each(function(index, element){
      let option = element.innerText.trim()
      let isFound = options.filter(item => { 
        return option === item.Title
      });

      if ( isFound !== undefined && isFound.length === 0 )
        $(this).remove()
      });

      let lineHeight = 0;

        liElement = list.find('li');
        liElement.each(function(index, element){
          let option = element.innerText.trim()
            if(index !== 0)
              lineHeight += 21;
            let tranVal = `translateY(${lineHeight}px)`;
            $(element).css({ 'transform': tranVal });
        })

        list.next().children().css('height','auto');
    }
  }
}, 500);

}

const removeEdit = async function(dtName){
let btns = $('.k-grid-update, .k-grid-cancel')
btns.each(function(index, ahref){
  $(this).on('click', function(){
  setTimeout(() => { 
    fixTableCols(dtName)
    }, 200);
  });
})
}

function fixTableCols(dtName){

  let tables = $(`div [internal-name='${dtName}'] table`)
  tables.each(function(tblIndex, tbl){
      let editLinks = $(this).find('a.k-grid-edit')
      if(editLinks.length > 0) 
          editLinks.remove();

      $(tbl).find('tr').each(function(trIndex, tr) {
        
          if (trIndex === 0){        
             let childs = tr.children;
             if(childs.length > 0){
               childs[0].style.textAlign = 'center';
               childs[4].style.textAlign = 'center';

               //tr.removeChild(childs[1]);
              }                          
          }
          
         $(tr).find('td').each(function(tdIndex, td) {
              if (tdIndex === 4){
                  td.style.textAlign = 'center'
              }
              //else if( tdIndex === 1) tr.removeChild(td);
          });                         
      });
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
	$('div.tab-content ul li a').each(function(i, element){
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


const getCurrentUserRole = async function (){
  let currentUser = await GetCurrentUser();
  isUserFound('Manager', currentUser);

  if(!_isPM){
    isUserFound('Director', currentUser);
    
    if(!_isPD){
      _isQM = await IsUserInGroup('QM');
      _isSus = await IsUserInGroup('Sustainability');
    }
  }
  debugger;
  _isGIS = await IsUserInGroup('GIS');

   if(!_isPD && !_isPM && !_isQM && !_isSus && !_isGIS) 
    _isUserAllowed = false;

   localStorage.setItem('_isPD', _isPD);
   localStorage.setItem('_isPM', _isPM);
   localStorage.setItem('_isQM', _isQM);
   localStorage.setItem('_isSus', _isSus);
   localStorage.setItem('_isGIS', _isGIS);
   localStorage.setItem('_isBuilding', _isBuilding);
}

const isUserFound = async function(fieldname, currentUser){
  let users = fd.field(fieldname).value;
  for(const user of users){
    if(user.email.toLowerCase() === currentUser.Email.toLowerCase()){
       if(fieldname === 'Director')
        _isPD = true;
       else _isPM = true
       break;
    }
  }
}

var _HideFields = async function(fields, isHide, isText){
	for(let i = 0; i < fields.length; i++)
	{
		let field;
    field = isText === true ? fd.field(fields[i]) : fields[i];
		if(isHide || isHide == undefined)
		  $(field.$parent.$el).hide();
      //fd.field(field).hidden = true;
		else $(field.$parent.$el).show(); //fd.field(field).hidden = false; 
    
	}
}

fd.spSaved(function(result){
    try{
        let query = _isNew ? `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${result.Id}</Value></Eq></Where>` : '';
        
        if(_isNew && _module === 'MTD' && !formFields.IsMTD )
            _sendEmail(_module, 'MTD_New_Email' + '|' + projectNo, query, '', 'MTD_New', '', CurrentUser);                
        
        else if(_isNew && _module === 'GIS')
          _sendEmail(_module, 'GIS_Email' + '|' + projectNo, query, '', 'GIS_New', '', CurrentUser);                
    }
    catch(e){alert(e);}                             
});

const getListOptions = async function(listname, fieldName, expandField, query){
  let options = [];
  let expandFieldName = fieldName;

  query = query !== undefined && query !== '' ? query : '';
  let col;
  
  if(typeof fieldName === 'string' || fieldName === undefined)
    col = fieldName === undefined ? 'Title' : `${fieldName}/Title,${fieldName}/Id`
  else  col = fieldName

  if(query.includes('MasterID'))
    col += ',MasterID/Id'

  let select = _web.lists.getByTitle(listname).items.select(`Id,${col}`);
  if(query.includes('MasterID')){
    select.expand(['MasterID',expandFieldName]);
  }
  else if(expandField){
    select.expand(expandFieldName);
  }

  if(query != '')
    select.filter(query);

  return await select
    .getAll()
    .then(async function (items){
        if(items.length > 0)
        {
          for(const item of items){
            let Id, lblText = '';

            if(query.includes('MasterID')){

             if(expandField && item[expandFieldName] !== null && item[expandFieldName] !== undefined){
                Id = item[expandFieldName].Id;
                lblText = item[expandFieldName].Title
             }
             else continue
            }
            else {
              Id = expandField && item[expandFieldName] !== null && item[expandFieldName] !== undefined ? item[expandFieldName].Id : item.Id;
              if(typeof col === 'string') 
                lblText = expandField && item[expandFieldName] !== null && item[expandFieldName] !== undefined ? item[expandFieldName].Title : item[col];
              else lblText = item[fieldName[1]]
            }

            options.push({
              value: Id,
              label: lblText
            })
          }
        }
      })
    .then(()=>{ return options; });
}

const setRequiredFieldErrorStyle = async function(htmlField, isCorrect){
  if(isCorrect === true)
    htmlField.style.setProperty('border', '1px solid #dee2e6', 'important');
  else htmlField.style.setProperty('border', '1px solid red', 'important');
}

const isDRD = async function(subCategory){
  let isDRD = false;
  subCategory.map(async (value)=>{
      if(value.IsDRD)
        isDRD = true
  });
  if(isDRD) 
    fd.container('DRDAccordion').hidden = false;
  else fd.container('DRDAccordion').hidden = true;
}

const appBar = async function(){

  let masterNavBar = $('#SuiteNavPlaceHolder');
  let childMenu = `<div class="sp-appBar" id="sp-appBar" role="navigation" aria-label="App bar" tabindex="-1">
                      <ul class="sp-appBar-linkContainer">`;
                      
  appBarItems.map(item=>{
        let editors = item.editors;
        let readers = item.readers;
        let className = item.iconTitle === 'GIS' && !_isPM ? 'dimmed-svg' : '';

        childMenu += `<li class="sp-appBar-linkLi" data-automation-id="sp-appBar-linkLi-globalnav">
                              <div class="sp-appBar-linkLiDiv">
                                <a class="sp-appBar-link ${className}" role="button" href="${item.redirectUrl}" onclick="openInCustomWindow(event)" style="display: flex; flex-direction: column; align-items: center; text-decoration: none;">
                                  <svg class= ${className} xmlns="http://www.w3.org/2000/svg" x="0px" y="0px" width="20" height="20" viewBox="${item.viewBox}">
                                    ${item.svgPath}                                  
                                  </svg>
                                <span style="font-size: 10px; color: black; margin-top: 3px;">${item.iconTitle}</span>                                                             
                                </a>                              
                                <div class="sp-appBar-tooltipDiv" style="position: absolute; left: 100%; top: 50%; transform: translateX(10px) translateY(-50%);">
                                  <span class="sp-appBar-tooltip" role="tooltip" id="sp-appbar-tooltiplabel-globalnav">${item.tooltip}</span>
                                </div>
                              </div>
                          </li>`;
  })
  childMenu += "</ul></div>";
  masterNavBar.append(childMenu);
}


let previewWindow = null, checkPreviewInterval = null;

function openInCustomWindow(event) {
  event.preventDefault(); // Prevent the default link behavior

  const url = event.currentTarget.href;
  const windowWidth = window.innerWidth;
  const windowHeight = window.innerHeight;

  const newWindowWidth = Math.floor(windowWidth * 0.9); // 60% of the current window's width
  const newWindowHeight = Math.floor(windowHeight * 0.9) + 50; // 85% of the current window's height

  const newWindowTop = Math.floor((windowHeight - newWindowHeight) / 2) + 200;
  const newWindowLeft = Math.floor((windowWidth - newWindowWidth) / 2);    

  const windowFeatures = `toolbar=no,scrollbars=yes,resizable=yes,width=${newWindowWidth},height=${newWindowHeight},top=${newWindowTop},left=${newWindowLeft},toolbar=no`;
 let fullProjTitle = localStorage.getItem('FullProjTitle');
  previewWindow = window.open(url, fullProjTitle, windowFeatures);

  // Apply styles for rounded corners
  if (previewWindow) {
      previewWindow.onload = () => {
          const styles = `
              body {
                  margin: 0;
                  overflow: hidden;
                  border-radius: 15px;
              }
              html {
                  border-radius: 15px;
              }
          `;
          const styleSheet = previewWindow.document.createElement("style");
          styleSheet.type = "text/css";
          styleSheet.innerText = styles;
          previewWindow.document.head.appendChild(styleSheet);
      };
  }


  preloader_btn(false, true);
  checkPreviewInterval = setInterval(closePreview, 50);
}

let closePreview = async function() {
  if (previewWindow && previewWindow.closed) {
      previewWindow = null; // Reset previewWindow before closing to prevent infinite loop
      clearInterval(checkPreviewInterval); // Clear any remaining intervals
      Remove_Pre(true);
  }
}

async function getSPtGroupMembers(groupName) {
  const group = await _web.siteGroups.getByName(groupName);
  const members = await group.users.get();
  return members;
}


function setPageStyle(ProjTitle){
  $('div.ms-compositeHeader').remove()
  let fullProjTitle = ProjTitle;
  if(fullProjTitle === undefined)
    fullProjTitle = localStorage.getItem('FullProjTitle');
  $('span.o365cs-nav-brandingText').text(fullProjTitle);
  $('i.ms-Icon--PDF').remove();


  let toolbarElements = document.querySelectorAll('.fd-toolbar-primary-commands');
  toolbarElements.forEach(function(toolbar) {
      toolbar.style.display = "flex";
      toolbar.style.justifyContent = "flex-end";
      toolbar.style.marginRight = "25px";            
  });   
  
  document.querySelectorAll('.CanvasZoneContainer.CanvasZoneContainer--read').forEach(element => {
      element.style.marginTop = '7px';
  });


    if(_isMain)
      setTabStyles()
}

function setTabStyles(){
  let tabIndex = 1;
  $('div.tabset,.tabs-top').eq(tabIndex).attr('style', 'padding-top: 0px !important');
  $('div.tabset,.tabs-top').parent().parent().parent().parent().eq(tabIndex).attr('style', 'padding-top: 2px !important; padding-left: 0px');

  $('div.tab-content').eq(tabIndex).attr('style', 'padding: 15px 12px 0px 12px !important');

  let uls = $("ul[role = 'tablist'").eq(tabIndex)
  uls.removeClass('nav-tabs');

  let colors = ['red', 'green', 'blue', 'yellow', 'gray', 'purple', 'navy', 'pink']
  let lis = uls.find('li')

  lis.each(function(index) {
    $(this).hover(
        function() {
            let color = colors[index]
            $(this).css({
              // 'background-color': 'lightblue',
              'border-bottom': `4px solid ${color}`
            }); // On mouse over
        },
        function() {
            $(this).css({
              'background-color': 'transparent',
              'border-bottom': '0px none transparent'
            }); // On mouse out
        }
    );

    $(this).click(function() {
      $(this).css('border-bottom', `4px solid ${color}`); 
    })
  });

  let activeTab = $('a.active').first().parent()
  activeTab.css('background-color','#5dc7b2b0'); 

  $('a').on('click', function(){
    $('a.active').first().parent().parent().children().css('background-color','');
    $(this).parent().css('background-color','#5dc7b2b0'); 
  })
}

const getEditableTabFields = async function(){
  const fields = fd.fields();
  const controls = fd.control();
  const Tabs = fd.container('Tabs1').tabs;

  let tabArray={};
  for(var i = 0; i < Tabs.length; i++){

    let tab = fd.container('Tabs1').tabs[i];
    let TabTitle = tab.title;
    let isTabExist = tabArray[TabTitle]
      if(isTabExist !== undefined) continue;
    
      let internalFields = []
      for(var j = 0; j < fields.length; j++){
        let field = fields[j];
        //debugger;
        let fieldTitle = field.title;
        let isReadOnly = field._readonly !== undefined ? JSON.parse(field._readonly.toLowerCase()) : false;

        let isValidField = isDescendant(tab.$el, field.$el);
        if(isValidField && !isReadOnly)
          internalFields.push(field.internalName);
      }
      tabArray[TabTitle] = internalFields;
  }
  return tabArray;
}


