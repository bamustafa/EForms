let  _web, _webUrl, _siteUrl, _formType = '', hot, hot1, container, container1, data, delivData, cdsDesignData = [];
let _isSiteAdmin = false, _isNew = false, _isEdit = false, isPM = false, isArea = false;
var _layout, _module, _isMain = true, _isLead = false, _isPart = false, projectName, Reference;

let _htLibraryUrl, _itemListname;
let _colsInternal = [], _colsType = [];

let defaultClassName = 'TransparentRow htMiddle', activeTabName, currentUser, closingStatus = 'Submitted'
let allowMultiEmailApproval = false, allowMultiEmailRejection = false, restrictFilesToMatchExcel = false, isAllCorrect = true;

var onRender = async function (moduleName, formType, relativeLayoutPath){
  _layout = relativeLayoutPath;
  await getPageParameters(moduleName, formType);

  if(moduleName == 'INS')
    await onINSRender(formType);

  await renderControls();
}

//#region GENERAL
var loadScripts = async function(){
    const libraryUrls = [
        _layout + '/plumsail/js/commonUtils.js',
        _layout + '/plumsail/js/customMessages.js',
        _layout + '/controls/preloader/jquery.dim-background.min.js',
        _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
        _layout + '/plumsail/js/preloader.js'
      ];

    const cacheBusting = `?v=${Date.now()}`;
      libraryUrls.map(url => { 
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
        });
        
    const stylesheetUrls = [
        _layout + '/controls/tooltipster/tooltipster.css',
        _layout + '/controls/handsonTable/libs/handsontable.full.min.css',
        _layout + '/plumsail/css/CssStyle.css'
        ];
  
    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
    });
}

var getPageParameters = async function(moduleName, formType){

    if($('.text-muted').length > 0)
      $('.text-muted').remove();

    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;
    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _itemListname = list.Title;

    if(_formType === 'New'){
        clearStoragedFields();
        _isNew = true;
    }
    else if(_formType === 'Edit'){
        _isEdit = true;
        Reference = fd.field('Reference').value;
        //_itemId = fd.itemId;
        //fd.field('chklstStates').clear();
    }
    await loadScripts();
    await setFormHeaderTitle();

    isPM = await IsUserInGroup('PM');
    isArea = await IsUserInGroup('Area');

    if(isArea) {
        fd.container('Tabs1').tabs[0].disabled = true;

        setTimeout(()=>{ 
            renderTabs()}, 
            100);
    }
    setMetaInfo();
    let arrayFunctions = [
        getParameter("AllowMultiEmailApproval"),
        getParameter("AllowMultiEmailRejection"),
        getParameter('RestrictFilesToMatchExcel')
    ];

     const params = await Promise.all(arrayFunctions);
     allowMultiEmailApproval = params[0].toLowerCase() === 'yes' ? true : false,
     allowMultiEmailRejection = params[1].toLowerCase() === 'yes' ? true : false,
     restrictFilesToMatchExcel = params[2].toLowerCase() === 'yes' ? true : false;

     currentUser = await GetCurrentUser();
}

function setToolTipMessages(){
    setButtonCustomToolTip('Submit', submitMesg);
    setButtonCustomToolTip('Cancel', cancelMesg);
}

var setButtons = async function () {
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

    await setButtonActions("Accept", "Submit");
    await setButtonActions("ChromeClose", "Cancel");
}

const setButtonActions = async function(icon, text){
    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          click: async function() {
           if(text == "Close" || text == "Cancel")
               fd.close();
           else if(text == "Submit"){
            activeTabName = activeTabName === undefined ? $('ul.nav-tabs li a.active')[0].innerHTML : activeTabName;
            if(activeTabName === 'Trades Approval')
              await onTradesApprovalSubmit()
               //fd.save();
           }
          } 
    });
}

function validateButtons(){
    if(!isAllCorrect)
            $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#737373').attr("disabled", "disabled");
        else $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#444').removeAttr('disabled');
}

var renderControls = async function(){
    await setButtons();
    setToolTipMessages();

    $('ul.nav-tabs li a').on('click', async function(element) {
         activeTabName = $(this).text();
         if(activeTabName === 'Trades Approval'){
            setTimeout(()=>{ 
                disableCheckBox()
                validateButtons()
            }, 
                100);
         }
         else{
            if(isPM) 
              $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#444').removeAttr('disabled');
         }
    });

    $('.deliv-title').hide();
    await renderMainTable();
    
   
}

var renderTabs = async function(tabIndex){
    var reqIndex = 1;
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

//#region TRADES APPROVAL TABLES
var _getData = async function(){
    let _itemArray = [], _dataArray = [];

    let _query = "Reference eq '" + Reference + "' and Status ne 'Cancelled'";
    let filterArray = _colsInternal.filter(item => item !== 'isSelected' && item !== 'ImageOnly' && item !== '_Mesg');
    const _cols = filterArray.join(',');

    var items = await _web.lists.getByTitle(_DesignTasks).items.filter(_query).select(_cols).getAll();
    if(items.length > 0){
        items.forEach(item => {
            let mesg = '';
            for(var j = 0; j < _colsInternal.length; j++){
                //var _type = _colsType[j];
                var _colname = _colsInternal[j];
                var _value = item[_colname];

                if(_colname === 'ImageOnly'){
                    var iconUrl =  _layout + '/Images/Submitted.png';
                    if(mesg !== '')
                      iconUrl =  _layout + '/Images/Error.png';

                     _value  = "<img src='" + iconUrl + "' alt='na'></img>";
                }
                else if(_colname === 'isSelected')
                   _value = 'no';
                else if(_colname === 'WorkflowStatus'){
                    if(_value === 'Pending')
                      mesg = 'Trade did not yet submit to PM'
                    else if(_value.toLowerCase() === 'sent to area'){
                        if(!isArea) mesg = 'Only Area users can review this trade'
                    }
                    else if(_value.toLowerCase() === 'sent to pm'){
                        if(!isPM)  mesg = 'Only PM users can review this trade'
                    }
                    else if(_value === closingStatus)
                    mesg = 'Trade is already submitted to area'
                }
                else if(_colname === '_Mesg')
                     _value = mesg;
            
                _itemArray[_colname] = _value;
            }
            _dataArray.push(_itemArray);
            _itemArray = [];
        });
     }
     return _dataArray;
}

var getDelivData = async function(){
    let _itemArray = [], _dataArray = [];

    let _query = "SubmittalRef eq '" + Reference + "' and Status ne 'Cancelled'";
  
    const _cols = 'FileName,Revision,Title,DarTrade';
    let colsArray = _cols.split(',');

    var items = await _web.lists.getByTitle(_RLOD).items.filter(_query).select(_cols).getAll();
    if(items.length > 0){
        items.forEach(item => {
            let mesg = '';
            for(var j = 0; j < colsArray.length; j++){

                let columnName = colsArray[j];
                var _value = item[columnName];

                 if(_value === null) _value = '';
                _itemArray[columnName] = _value;
            }
            
            _dataArray.push(_itemArray);
            _itemArray = [];
        });
     }
     return _dataArray;
}

const _setData = async (Handsontable) =>{
    container = document.getElementById('dt');

    let colArray = [
        {
            title: ' ',
            data: 'isSelected',
            type: 'checkbox',
            width: '3%',
            checkedTemplate: 'yes',
            uncheckedTemplate: 'no'
        },
        {
            title: 'Trade',
            data: 'Trade',
            type: 'text',
            width: '10%',
            readOnly: true,
            className: 'htLeft'
        },
        {
            title: 'Task Status',
            data: 'Status',
            type: 'text',
            width: '10%',
            readOnly: true,
            className: 'htCenter'
        },
        {
            title: 'Workflow Status',
            data: 'WorkflowStatus',
            type: 'text',
            width: '15%',
            readOnly: true,
            className: 'htLeft'
        },
        {
            title: '_Mesg',
            data: '_Mesg',
            type: 'text',
            width: '20%',
            readOnly: true,
            className: 'htLeft ErrorMesg'
        },
        {
            title: 'Ready for Submittal',
            data: 'ImageOnly',
            type: 'text',
            width: '20%',
            readOnly: true,
            className: 'htCenter',
            renderer: 'html'
        }
    ];

    colArray.map(item =>{
        _colsInternal.push(item.data);
        _colsType.push(item.type);
    });

    data = await _getData();
    delivData = await getDelivData();

    let height = 'auto';
    // if(data.length > 5) 
    //   height = '300';
    
    hot = new Handsontable(container, {
        data: data,
        columns: colArray,
        width:'100%',
        height: height,
        //colHeaders: true,
        rowHeaders: true,
        stretchH: 'all',
        licenseKey: htLicenseKey
    });

    setTickBoxReadOnly(false);
    
    setTimeout(()=>{ 
        hot.render();
        disableCheckBox();
        validateButtons();
      }, 
    500);

    performAfterChangeActions((rowIndex, operation) => {
        let tempData = data.filter((item)=>{
           return item.isSelected === 'yes'
        })
        if(tempData.length > 0){
            _spComponentLoader.loadScript(_htLibraryUrl).then(bindDelivGrid);
            $('.deliv-title').show();
            $('#dt1').show();
        }
        else{
            $('.deliv-title').hide();
            $('#dt1').hide();
        }
        disableCheckBox();
    });
}

const bindDelivGrid = async (Handsontable) =>{
    let colArray = [
        {
            title: 'FileName',
            data: 'FileName',
            type: 'text',
            width: '15%',
            readOnly: true,
            className: 'htLeft'
        },
        {
            title: 'Revision',
            data: 'Revision',
            type: 'text',
            width: '10%',
            readOnly: true,
            className: 'htCenter'
        },
        {
            title: 'Title',
            data: 'Title',
            type: 'text',
            width: '40%',
            readOnly: true
        },
        {
            title: 'Trade',
            data: 'DarTrade',
            type: 'text',
            width: '15%',
            readOnly: true,
            className: 'htCenter'
        },
    ];

    debugger;
    let tempData = delivData.filter(delivItem=>{
            return data.some(item => {
                  return item.isSelected === 'yes' && item.Trade === delivItem.DarTrade;
            });
    })
    
    let _itemArray = {};
    tempData.filter(item=>{
        _itemArray['filename'] = item['FileName'];
        _itemArray['revision'] = item['Revision'];
        _itemArray['title'] = item['Title'];
        cdsDesignData.push(_itemArray);
        _itemArray = {};
    })


    if(hot1 !== undefined){
        hot1.updateSettings({
            data: tempData
          });
    }
    else{
        container1 = document.getElementById('dt1');
        hot1 = new Handsontable(container1, {
            data: tempData,
            columns: colArray,
            width:'99%',
            height: 'auto',
            rowHeaders: true,
            colHeaders: true,
            //manualColumnResize: true,
            stretchH: 'all',
            licenseKey: htLicenseKey
        });
    }

    setTimeout(()=>{ 
        hot1.render();
        disableCheckBox();
      }, 
    100);
}

function setTickBoxReadOnly(isReject){
    let rows = hot.getData();
    const rowCount = hot.countRows();
    let isDisable = isReject ? false : true;

    var defaultMetaObject = {
        readOnly: isDisable,
        className: defaultClassName
    };

    let isConditionMet = false;
    for(let row = 0; row < rowCount; row++){
        let tickIndex = hot.propToCol('isSelected');
        let mesgIndex = hot.propToCol('_Mesg');
        let wfIndex = hot.propToCol('WorkflowStatus');
        let message = hot.getDataAtCell(row, mesgIndex);
        let workflowStatus = hot.getDataAtCell(row, wfIndex);

        if(workflowStatus === closingStatus || (message !== undefined && message !== null && message !== '')){
          isAllCorrect = false;
          hot.setCellMetaObject(row, tickIndex, defaultMetaObject);
        }
        else if(isReject) 
          isConditionMet = true;
    }

    if(isReject && isConditionMet)
      isAllCorrect = true;
}

function disableCheckBox(){

    let colIndex = 0;
    let tickIndex, mesgIndex, wfIndex;

    hot.getSettings().columns.find(item => {
          if(item.data === 'isSelected')
            tickIndex = colIndex;
          else if(item.data === '_Mesg')
            mesgIndex = colIndex;
          else if(item.data === 'WorkflowStatus')
            wfIndex = colIndex;
          colIndex++;
    });

    var tbodyElement = $('#dt').find('table.htCore')[0].querySelector('tbody');
    var trElements = tbodyElement.querySelectorAll('tr');

    $(trElements).each(function(index, trElement) {
        var tdElements = $(trElement).find('td');
        var InputBox, mesgValue, wfValue;
        $(tdElements).each(function(tdIndex, tdElement) {
            if(tdIndex === tickIndex)
              InputBox = this.children[0];
            else if(tdIndex === mesgIndex)
              mesgValue = this.outerText;
            else if(tdIndex === wfIndex)
              wfValue = this.outerText;
          });

          if(mesgValue !== '' || wfValue === closingStatus)
           $(InputBox).prop('disabled', true);
    });
}

function setMetaInfo(){
    fd.field('Ref').value = Reference;
    fd.field('Ref').disabled = true;

    fd.field('Tit').value = fd.field('Title').value;
    fd.field('Tit').disabled = true;

    fd.control('hlink').text = Reference;
    fd.control('hlink').href = `${_webUrl}/${_Deliverables}/${Reference}`;

    let stat = fd.field('ChoiceStatus').value;
    if(stat === 'Rejected') 
        $(fd.field('RejReason').$parent.$el).show();
    else $(fd.field('RejReason').$parent.$el).hide();

    fd.field('ChoiceStatus').$on('change', function(value) {
        if(value === 'Rejected'){
          $(fd.field('RejReason').$parent.$el).show();
          fd.field('RejReason').required = true;
        }
        else {
            fd.field('RejReason').required = false;
            $(fd.field('RejReason').$parent.$el).hide();
        }
   });  
}

async function performAfterChangeActions(callback) {
    hot.addHook('afterChange', (changes, source) => {
        if (typeof callback === 'function') {
            callback();
         }
    });
}

async function renderMainTable(){
    fd.field('ChoiceStatus').$on("change", function (value) {
         if(value === 'Rejected'){
            setTickBoxReadOnly(true)
            validateButtons();
         }
         else{
            setTickBoxReadOnly(false)
            validateButtons();
         }
    });
}
//#endregion

var onINSRender = async function (){
	if(_isNew || _isEdit){
		let query = `Title ne null and IsDeliverableModule eq '1'`;
		fd.field('Trades').ready().then(() => {
			fd.field('Trades').filter = query
			fd.field('Trades').orderBy = { field: 'Title', desc: false };
			fd.field('Trades').refresh();
		});
	}
    
	if(_isEdit){
		const Status = fd.field('Status').value;
		if(Status === "Reviewed" || isArea){
            fd.field('Title').disabled = true;
			fd.field('Trades').disabled = true;
			fd.field('DueDate').disabled = true;
			fd.field('Comment').disabled = true;
		}

        _spComponentLoader.loadScript(_htLibraryUrl).then(_setData);
    }
}

var onTradesApprovalSubmit = async function(){
    let mesg = '', WorkflowStatus = '';
    let status = fd.field('ChoiceStatus').value;
    let comment = fd.field('RejReason').value;

    let isItemsSelect = data.filter(item =>{ return item.isSelected === 'yes'});
    if(isItemsSelect.length === 0){
        mesg = 'Please make sure to select at least one trade.';
        alert(mesg);
        return;
    }
    else WorkflowStatus = isItemsSelect[0].WorkflowStatus;

    let wfStepMetaInfo = await getWorkflowSteps_MetaInfo(WorkflowStatus);
    await setCoreFunctionality(status, WorkflowStatus, isItemsSelect, wfStepMetaInfo);  
}

var getWorkflowSteps_MetaInfo = async function(wfStatus){
    let _query = `CurrentStatus eq '${wfStatus}'`;
    var items = await _web.lists.getByTitle(_WorkflowSteps).items.filter(_query).getAll();
    if(items.length > 0){
        let item = items[0];
        return{
            ApprovalWorkflowStatus: item["ApprovedStatus"],
            ApprovalNotificationUser: item["NotificationUser"],

            ApprovalEmailName: item["ApprovalEmailName"],
            ApprovalTradeCC: item["ApprovalTradeCC"],
            RejectionWorkflowStatus: item["RejectedStatus"],

            RejectionNotificationUser: item["RejectionNotificationUser"],
            RejectionEmailName: item["RejectionEmailName"],
            RejectionTradeCC: item["RejectionTradeCC"],
            Role: item["Title"],
            IncludeCurrentUser: item["CurrentUser"]
        }
    }
}

var setCoreFunctionality = async function(status, WorkflowStatus, isItemsSelect, wfStepMetaInfo){
    let isAllSelected = false;
    if(isItemsSelect.length === data.length)
       isAllSelected = true;

    let partTrades = '';
    
    for (let i = 0; i < isItemsSelect.length; i++){
        let trade = isItemsSelect[i].Trade;
        partTrades += trade + ',';
        //#region DET RLOD AND DESIGN TASKS ITEMS
        let cols = 'Id,Status,Comments,ProjectName,SubmittalRef,Title,CDSTitle,FileName,Revision,DarTrade,DWGLink,DWFLink,PDFLink,CDSNumber,Category,RejectedComments';
        let query = `SubmittalRef eq '${Reference}' and DarTrade eq '${trade}' and Status ne 'Cancelled'`;
        let rlodItems = await _web.lists.getByTitle(_RLOD).items.filter(query).select(cols).getAll();

        cols = 'Id, Comments,WorkflowStatus,RejectedComments,Status';
        query = `Reference eq '${Reference}' and Trade eq '${trade}' and Status ne 'Cancelled'`;
        let designItems = await _web.lists.getByTitle(_DesignTasks).items.filter(query).select(cols).getAll();
        //#endregion
        
        let formUrl = `${_webUrl}/SitePages/PlumsailForms/${_CDS}/Item/NewForm.aspx?acronym=&Reference=${Reference}&Trade=${trade}`;
        let approvalStatus = wfStepMetaInfo.ApprovalWorkflowStatus;

        if(status === 'Approved'){
           if (WorkflowStatus != "Sent To Area"){
             await updateList(WorkflowStatus, rlodItems, true, true, wfStepMetaInfo); //RLOD to remove after
             await updateList(WorkflowStatus, designItems, false, true, wfStepMetaInfo); // DESIGN TASK to remove after
             let folder = await updateDeliverablesLibrary(trade, approvalStatus);

             if(approvalStatus === 'Sent To Area'){
                if(isAllSelected){
                    let item = await updateInitiateSubmittal(approvalStatus);
                    await setItemPermission(item, 'PM', 'Contribute', 'Read');
                }
             }
             else if(approvalStatus === closingStatus){
                cachMetaInfo(trade)
                window.location.href = formUrl;
             }
           }
           else {
            cachMetaInfo(trade)
            window.location.href = formUrl;
            } //Approved by Area
        }

        else{ // Rejected
            await updateList(WorkflowStatus, designItems, false, false, wfStepMetaInfo); // DESIGN TASK
            await updateList(WorkflowStatus, rlodItems, true, false, wfStepMetaInfo); //RLOD
            let folder = await updateDeliverablesLibrary(trade, wfStepMetaInfo.RejectionWorkflowStatus);

            if (WorkflowStatus === "Sent To PM")
                await setItemPermission(folder, trade, 'Read', 'Contribute');
            else if (WorkflowStatus === "Sent To Area"){
                let item = await updateInitiateSubmittal('Open');
                await setItemPermission(item, 'PM', 'Read', 'AddEdit');

                if (approvalStatus == 'Sent To Area')
                  await setItemPermission(folder, 'PM', 'Read', 'Contribute');
            }
        }
    }

    partTrades = partTrades.slice(0, -1);
    await sendEmail(partTrades, status, wfStepMetaInfo);
}

var updateList = async function(WorkflowStatus, items, isRLOD, isApproved, wfStepMetaInfo){
    let comment = fd.field('RejReason').value;
    let finalComment = '';
    let batch = pnp.sp.createBatch();
    const batchSize = 30;
    let row = 0, rowBatchSize = 1; 
    let totalItems = items.length;

     for (const item of items){
        let setItem = {};

        if(!isApproved && finalComment === ''){
            const now = new Date();
            const day = now.getDate();
            const month = now.getMonth() + 1; // Adding 1 because getMonth() returns zero-based index
            const year = now.getFullYear();

            const formattedDateTime = `${day}/${month}/${year}`;
            let user = `${currentUser.Title} (${item['Title']}) ${formattedDateTime}`;
            let oldRejComment = item['RejectedComments'] !== null && item['RejectedComments'] !== '' ? item['RejectedComments'] : '';
            if(oldRejComment !== '')
               finalComment = `${oldRejComment} ${user}: ${comment} <hr>`;
            else finalComment = comment;
        }

        if(!isRLOD){
            if(!isApproved){
              setItem['RejectedComments'] = finalComment;
              setItem['WorkflowStatus'] = wfStepMetaInfo.RejectionWorkflowStatus;

              if (WorkflowStatus === 'Sent To PM')
                 setItem["Status"] = 'Re-Opened';
            }
            else{
                setItem['WorkflowStatus'] = wfStepMetaInfo.ApprovalWorkflowStatus;
            }
            await _web.lists.getByTitle(_DesignTasks).items.getById(item.Id).update(setItem);
        }
        else{
            if(!isApproved){
                setItem['Status'] = 'Rejected';
                setItem['RejectedComments'] = finalComment;
            }
            else  setItem['Status'] = wfStepMetaInfo.ApprovalWorkflowStatus;
            _web.lists.getByTitle(_RLOD).items.getById(item.Id).inBatch(batch).update(setItem);
        }

        if(rowBatchSize === batchSize || rowBatchSize === totalItems){
            row++;
            detailedLoader(totalItems, row);
            rowBatchSize = 1;

            await batch.execute();
            batch = pnp.sp.createBatch();
            continue;
        }
        rowBatchSize++;
        row++;
     }
}

var updateDeliverablesLibrary = async function(trade, status){
    let relativeUrl = `${_Deliverables}/${Reference}/${trade}`;
    const folder = await _web.getFolderByServerRelativeUrl(relativeUrl).getItem();
    const folderId = folder.Id;

    await _web.lists.getByTitle(_Deliverables).items.getById(folderId).update({
        WorkflowStatus: status
    });
    return folder;
}

var setItemPermission = async function(item, group, permissionBefore, permissionAfter){
    await item.breakRoleInheritance(false);
    const { Id: contRoleDefId } = await _web.roleDefinitions.getByName(permissionBefore).get();
    const { Id: assignedRoleDefId } = await _web.roleDefinitions.getByName(permissionAfter).get();
    let roleAssignments = await item.roleAssignments();

    let promises = roleAssignments.map(async (roleAssignment) =>{
       return {
         principalId: roleAssignment.PrincipalId,
         group: await _web.siteGroups.getById(roleAssignment.PrincipalId).get().catch(()=>{ return undefined })
       }
    })

    Promise.all(promises)
    .then(results=>{
        return results.filter(result=>{
            if(result.group !== undefined && result.group.Title === group)
               return result;
        })
    })
    .then(async (groups)=>{
        Promise.all(groups.map(async group=>{
            item.roleAssignments.remove(group.principalId, contRoleDefId);
            item.roleAssignments.add(group.principalId, assignedRoleDefId);
        }))
        .then(()=>{
            console.log(`${Reference} permission is set successfully.`);
        });
    })
}

var updateInitiateSubmittal = async function(status){
    let isQuery = `Reference eq '${Reference}'`
    var items = await _web.lists.getByTitle(_InitiateSubmittals).items.filter(isQuery).select().getAll();

    if(items.length > 0){
       let item = items[0];

        await _web.lists.getByTitle(_InitiateSubmittals).items.getById(item.Id).update({
            Status: status
        });
        return item;
    }
}

var sendEmail = async function(partTrades, status, wfStepMetaInfo){
    //let query = `SubmittalRef eq '${Reference}' and Status ne 'Cancelled' and (${partTrades})`;
    let workflowStatus, emailName, notificationUser, tradeCC, includeCurrentUser;
    if(status === 'Approved'){
        workflowStatus = wfStepMetaInfo.ApprovalWorkflowStatus;
        emailName = wfStepMetaInfo.ApprovalEmailName;
        notificationUser = wfStepMetaInfo.ApprovalNotificationUser;
        //tradeCC = wfStepMetaInfo.ApprovalTradeCC;
    }
    else{
        workflowStatus = wfStepMetaInfo.RejectionWorkflowStatus;
        emailName = wfStepMetaInfo.RejectionEmailName;
        notificationUser = wfStepMetaInfo.RejectionNotificationUser;
        //tradeCC = wfStepMetaInfo.RejectionTradeCC;
    }
    includeCurrentUser = wfStepMetaInfo.IncludeCurrentUser;

    await setSoapContent(status, workflowStatus, partTrades, emailName, notificationUser);
}

var setSoapContent = async function(ApprovalStatus, workflowStatus, partTrades, emailName, notifName){
    let method = 'SEND_DESIGN_APPROVAL_EMAIL';
    let serviceUrl = `${_siteUrl}/AjaxService/DarPSUtils.asmx?op=${method}`;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                        <soap:Body>
                            <${method} xmlns="http://tempuri.org/">
                                <WebURL>${_webUrl}</WebURL>
                                <Reference>${Reference}</Reference>
                                <Status>${ApprovalStatus}</Status>
                                <WorkflowStatus>${workflowStatus}</WorkflowStatus>
                                <Trades>${partTrades}</Trades>
                                <EmailName>${emailName}</EmailName>
                                <Notification_Name>${notifName}</Notification_Name>
                                <CurrentEmail>${currentUser.Email}</CurrentEmail>
                            </${method}>
                        </soap:Body>
                    </soap:Envelope>`
    let res = await getSoapResponse('POST', serviceUrl, true, soapContent, `${method}Result`);
}

var  detailedLoader = async function(total, index, isDim){
    var targetControl = $('#ms-notdlgautosize').addClass('remove-position-preloadr');
    
    if(isDim === true){
        targetControl.dimBackground({
            darkness: 0.4
        }, function () {}); 
    }
    else{
        let mesg = `Processing ${index} out of ${total}. Please dont close this page.`;
        if($('#totalId').length > 0)
          $('#totalId').text(mesg);
        else{
            $(`<span id='totalId'>${mesg}</span>`)
            .css({
                "position": "fixed",
                "top": "50%",
                "left": "50%",
                "transform": "translate(-50%, -50%)",
                "background-color": "rgb(115 115 115 / 55%)",
                "border-radius": "15px",
                "color": "white",
                "box-shadow": "2px 10px 10px rgba(0, 0, 0, 0.1)",
                "transition": "opacity 0.3s ease-out",
                "font-weight": "bold",
                "padding": "10px",
                "font-size": "20px",
                "width": "auto"
            }).insertAfter(targetControl);
        }         
    }          
}

function cachMetaInfo(trade){
    //let tempData  = hot1.getData();
    console.log(cdsDesignData);
    const json = JSON.stringify(cdsDesignData);
    localStorage.setItem('data', json);
    localStorage.setItem('acronym', 'GEN');
    localStorage.setItem('trade', trade);
    localStorage.setItem('submittalRef', Reference);
}