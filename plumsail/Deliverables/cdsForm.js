let  _web, _webUrl, _siteUrl, _formType = '';
let _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDesign = false, _isMultiContractor = false, isAssignAllowed = false, enableSubmissionDate = false;
let _itemId, _itemListname, _htLibraryUrl, data, acronym, trade, contCode, transNo, designRef, cdsRef;

var _layout, _module, _isMain = true, _isLead = false, _isPart = false, projectName;
let screenHeight = screen.height - 500;

let hot, container;
let allowedGroupAssigner, onLoadQuery, assigneeGroupName;
let splitter = '-';
let currentUser;

var onRender = async function (relativeLayoutPath, moduleName, formType){
    _layout = relativeLayoutPath;
    await loadScripts();

    preloader();
    
    await getPageParameters(moduleName, formType);
    await renderControls();

    if(_isEdit){
        let isAllowed = await checkAllowedUser();
        if(!isAllowed){
            $('span').filter(function () { return $(this).text() == 'Assign'; }).parent().css('color', '#737373').attr("disabled", "disabled");
            return;
        }

        await getData();
    }

    if(data !== null){
      _spComponentLoader.loadScript(_htLibraryUrl).then(_setData);
      preloader('remove');
    }
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
        _itemId = fd.itemId;
        fd.field('chklstStates').clear();
    }

    let arrayFunctions = [
        getParameter("AllowedGroupAssigner"),
        getParameter("OnLoadQuery"),
        getParameter('Phase'),
        getParameter('IsAssignAllowed'),
        getParameter('EnableSubmissionDate'),
        getParameter("ProjectName"),
        isMultiContractor() // _isMultiContracotr is already global variable there
    ];

    const params = await Promise.all(arrayFunctions);
    allowedGroupAssigner = params[0];
    onLoadQuery = params[1];
    _isDesign = params[2].toLowerCase() === 'design'? true : false;
    phase = params[2], 
    isAssignAllowed = params[3].toLowerCase() === 'yes' ? true : false, 
    enableSubmissionDate = params[4].toLowerCase() === 'yes' ? true : false;
    projectName = params[5];

    currentUser = await GetCurrentUser();
}

var renderControls = async function(){
    let [projectTitle, projectDescription] = await Promise.all([
            getParameter('ProjectTitle'),
            getParameter('ProjectDescription')
        ]);
    let moduleTitle  = `${projectTitle} <br/> ${projectDescription}`;
    await addLegend('modulelblId', moduleTitle, 'moduleTitle', 'same');

    await setButtons();
    await setValues();

    var fields = ['Acronym', 'ContCode', 'TransNo', 'DesignRef', 'CDSNumber'];
    if(_isEdit){
        fields.push('CDSNumber');
        fields.push('LeadTrade');
        fields.push('PartTrade');
        fields.push('WorkflowStatus');
    }
    HideFields(fields, true);
    setToolTipMessages();
}

const setValues = async function(){
    if(_isNew){
        data = localStorage.getItem('data');
        data = JSON.parse(data);
        //localStorage.removeItem('data');

        acronym = localStorage.getItem('acronym');
        trade = localStorage.getItem('trade');

        //localStorage.removeItem('acronym');
        //localStorage.removeItem('trade');
       
        fd.field('Acronym').value = acronym;
        //fd.field('Trade').value = trade;

        if(!_isDesign){
            transNo = localStorage.getItem('transNo');
            contCode = localStorage.getItem('contCode');
    
            //localStorage.removeItem('transNo');
            //localStorage.removeItem('contCode');

            if(contCode !== undefined && contCode !== null && contCode !== '')
              fd.field('ContCode').value = contCode;
            else contCode = '';

            if(transNo !== undefined && transNo !== null && transNo !== '')
              fd.field('TransNo').value = transNo;
            else transNo = '';
        }
        else {
            designRef = localStorage.getItem('submittalRef');
            //localStorage.removeItem('submittalRef');

            if(designRef !== undefined && designRef !== null && designRef !== '')
              fd.field('DesignRef').value = designRef;
            else designRef = '';
        }

        if(_isDesign || !enableSubmissionDate){
            fd.field('Date').value = new Date();
            fd.field('Date').disabled = true;
        }
    }
    else{
        acronym = fd.field('Acronym').value;
        trade = fd.field('Trade').value;
        contCode = fd.field('ContCode').value;
        transNo = fd.field('TransNo').value;
        designRef = fd.field('DesignRef').value;
    }
}

function HideFields(fields, isHide){
	var field;
	for(let i = 0; i < fields.length; i++)
	{
		field = fd.field(fields[i]);
		if(isHide || isHide == undefined)
		  $(field.$parent.$el).hide();
		else $(field.$parent.$el).show();
	}
}

function setToolTipMessages(){
    setButtonCustomToolTip('Assign', assignMesg);
    setButtonCustomToolTip('Submit', submitMesg);
    setButtonCustomToolTip('Cancel', cancelMesg);
}

var setButtons = async function () {
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

    if(_isEdit){
        if(isAssignAllowed)
          await setButtonActions("Accept", "Assign");
    }
    else 
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
             let isCorrect = await chckRequiredFields();
             if(isCorrect){
                preloader();
                await onSubmit();
                fd.close();
             }
           }
           else if(text == "Assign"){
            let isCorrect = await chckRequiredFields();
            if(isCorrect){
                let leadTrade = fd.field('LeadTradeDDL').$el.innerText;
                if(leadTrade === ''){
                    setPSErrorMesg('Lead Trade is required <br/>' );
                }
                
               preloader();
               await onAssign(leadTrade);
            }
          }
          } 
    });
}
//#endregion

const _setData = (Handsontable) =>{
    container = document.getElementById('dt');
    
      let colArray = [{
          title: 'Filename',
          data: 'filename',
          type: 'text',
          width: '20%',
          readOnly: true,
          className: 'htLeft'
      },
      {
        title: 'Revision',
        data: 'revision',
        type: 'text',
        width: '10%',
        readOnly: true,
        className: 'htCenter'
      },
      {
          title: 'Title',
          data: 'title',
          type: 'text',
          width: '65%',
          readOnly: true,
          className: 'htLeft'
      }];
  
      hot = new Handsontable(container, {
          data: data,
          columns: colArray,
          width:'100%',
          //height: screenHeight,
          //colHeaders: true,
          rowHeaders: true,
          stretchH: 'all',
          licenseKey: htLicenseKey
      });
}

const setDimonButton = async function(text, isDisabled){
    if(isDisabled)
      $('span').filter(function () { return $(this).text() == text; }).parent().css('color', '#737373').attr("disabled", "disabled");
    else $('span').filter(function () { return $(this).text() == text; }).parent().css('color', '#444').removeAttr('disabled');
}

//#region ON SUBMIT CLICK
const onSubmit = async function(){
    debugger;
    let mItem = await getMajorType(acronym);
    if(mItem.length === 0)
     return;

    let isSplitAllowed = (mItem[0]["isSplitAllowed"] != null) ? JSON.parse(mItem[0]["isSplitAllowed"]) : false;
    let refFormat = mItem[0]["CDSFormat"] !== null ? mItem[0]["CDSFormat"] : '';
    let cdsNumber = '';

    if (isSplitAllowed){
        for (const item of data){
            cdsNumber = await checkCDSRef(refFormat);

            let fulleFileName = item.filename;
            fulleFileName = item.revision !== undefined && item.revision !== null && item.revision !== '' ? `${fulleFileName}${splitter}${item.revision}` : fulleFileName;
            result = await updateRLOD(cdsNumber, item.filename, fulleFileName, isSplitAllowed);
            await insertCDSItem(cdsNumber); //result.trade, result.title);
        }
    }
    else{
        cdsNumber = await checkCDSRef(refFormat);
        if(!_isDesign){
            for (const item of data) {
                let fulleFileName = item.filename;
                fulleFileName = item.revision !== undefined && item.revision !== null && item.revision !== '' ? `${fulleFileName}${splitter}${item.revision}` : fulleFileName;
                await updateRLOD(cdsNumber, item.filename, fulleFileName, false);
            };
        }
        cdsRef = cdsNumber.replace('"', '');
        await insertCDSItem(cdsNumber);
    }
}

const checkCDSRef = async function(refFormat){
    let referenceValue = await getReferenceValue(refFormat);
    let referenceWithoutDigits = referenceValue.substring(0, referenceValue.lastIndexOf(splitter));
    let digit = referenceValue.substring(referenceValue.lastIndexOf(splitter) + 1);

    let seqNum = await getCounter(_web, referenceWithoutDigits, true);
    let counterRef = `${referenceWithoutDigits}${splitter}${seqNum.toString().padStart(digit.length, '0')}`;
    fd.field('CDSNumber').value  = counterRef;
    return counterRef;
}

const getReferenceValue = async function(refFormat){
  let splitRefFormat = refFormat.split(splitter);
  let returnValue = '';

  for (let i = 0; i < splitRefFormat.length; i++) {
    let Column = splitRefFormat[i].replace("[", "").replace("]", "");

    if (Column.includes("\"")) 
        Column = Column.replace("\"", "");
    else if (Column.includes("$")) 
        Column = Column.replace("$", "");
    else {
        //let batch = pnp.sp.createBatch();
        for (let i = 0; i < data.length; i++) {
            let Exit = false;
            let item = data[i];
            let query = `FileName eq '${item.filename}'`;
        
            _web.lists.getByTitle(_RLOD).items.filter(query).getAll()
            .then((items) => {
                if (items.length > 0) {
                    for (const item of items) {
                        if (item[Column] != null && item[Column].toString() !== "") {
                            Column = item[Column].toString();
                            Exit = true; // Set Exit flag to true to break out of the loop
                            break;
                        }
                    }
                }
            });
        
            if (Exit) {
                break;
            }
        }
    }
    returnValue += Column.replace("$", "") + "-";
  }
  if (returnValue.endsWith(splitter))
    returnValue = returnValue.slice(0, -1);
  return returnValue;
}

const updateRLOD = async function(cdsNumber, filename, fullFileName, isSplitAllowed){
    let trade, title;
    let query = `FileName eq '${filename}'`;
    let items = await _web.lists.getByTitle(_RLOD).items.select("Id, PendingCDS, PendingFileName, CDSTitle, Title, Trade").filter(query).getAll();

    if(items.length > 0){
      let item = items[0];
      let setItem = {};
      setItem["PendingCDS"] = cdsNumber;
      setItem["PendingFileName"] = fullFileName;

      if(isSplitAllowed){
        trade = item["Trade"] != null ? item["Trade"] : '';
        title = item["Title"] != null ? item["Title"] : '';
        setItem["CDSTitle"] = title;
      }
      await _web.lists.getByTitle(_RLOD).items.getById(item.Id).update(setItem);
    }
    return{
        trade: trade,
        title: title
    }
}

const insertCDSItem = async function(concatParts){
    let item = {};
    
    if(_isNew){
        item["Trade"] = trade;
        item["Date"] = fd.field('Date').value;
        item["CDSNumber"] = cdsRef;
        item["Title"] = fd.field('Title').value;
        item["Comment"] = fd.field('Comment').value;
        item["Acronym"] = acronym;
    
        item["ContCode"] = contCode;
        item["TransNo"] = transNo;
        item["DesignRef"] = designRef;
        await _web.lists.getByTitle(_CDS).items.add(item);
    }
    else if(_isEdit){
        let remarks = fd.field('Remarks').value !== null ? fd.field('Remarks').value : '';
        item["Remarks"] = remarks;

        let type = typeof fd.field('LeadTradeDDL').value;

        if(type === 'object'){
            if(fd.field('LeadTradeDDL').value.length > 0)
              item["LeadTrade"] = fd.field('LeadTradeDDL').value[0];
        }
        else item["LeadTrade"] = fd.field('LeadTradeDDL').value;
        item["PartTrade"] = concatParts;

        if(fd.field('Location').value === undefined || fd.field('Location').value === null){
            let location = {
                Description: cdsRef,
                Url: fd.field('Location').value.url
            };
          item["Location"] = location;
        }
        await _web.lists.getByTitle(_CDS).items.getById(_itemId).update(item); 
    }
}

const getLocation = async function(item){

    let arrayFunctions = [
        getParameter("MultiRepository"),
        getParameter("MultiRepositoryFieldName"),
        getParameter('FilesRepositoryUrl'),
        getParameter('SubmittedDelivFolderName'),
        getParameter('DelivFolderStructure'),
        getParameter('RepositoryListName')
    ];

    let [MutliRepository, MultiRepositoryFieldName, RepositoryPath, SubmittedDelivFolderName, DelivFolderStructure, RepositoryListName] = await Promise.all(arrayFunctions);

    if (RepositoryPath === '' || RepositoryPath.toLowerCase().includes('tempurl'))
        RepositoryPath = _webUrl;
    
        let ConcateFolderStructure = '', locationUrl = '';

        let SplitDelivFolderStructure = DelivFolderStructure.split(splitter);
        for (let i = 0; i < SplitDelivFolderStructure.length; i++) {
            let ColumnName = SplitDelivFolderStructure[i];
            if (ColumnName === "Acronym") 
                ConcateFolderStructure += acronym + "/";
            else ConcateFolderStructure += item[ColumnName].toString() + "/";
        }

        if (MutliRepository.toLowerCase() === "yes" || MutliRepository === "1") {
            locationUrl = `${RepositoryPath}/${item[MultiRepositoryFieldName]}/${RepositoryListName}/${SubmittedDelivFolderName}/${ConcateFolderStructure}`;
            //ReamrkFileURL = `${RepositoryPath}/${item[MultiRepositoryFieldName]}/${RepositoryListName}/Remarks/${Reference}`;
        } else {
            locationUrl = `${RepositoryPath}/${RepositoryListName}/${SubmittedDelivFolderName}/${ConcateFolderStructure}`;
            //ReamrkFileURL = `${RepositoryPath}/${RepositoryListName}/Remarks/${Reference}`;
        }
    return locationUrl;
}
//#endregion

//#region ON ASSIGN CLICK
const onAssign = async function(leadTrade){
 
   //#region EXTRACT VALUES
    let UserProfile = await GetCurrentUser();
    let cc = UserProfile !== undefined ? UserProfile.Email : '';

    let leadColumnName = acronym + "Lead";
    let partColumnName = acronym + "Part";
    let ReviewLinkUrl =  `${_webUrl}${_layout}/ReviewList.aspx?CDSNumber=${cdsRef}&Type=${acronym}`;
    let reviewLink = {
        Description: 'Review Link',
        Url: ReviewLinkUrl
    };

    let cdsNumber = fd.field('CDSNumber').value;
    let cdskUrl =  `${_webUrl}/SitePages/PlumsailForms/${_CDS}/Item/EditForm.aspx?item=${_itemId}`;
    let cdsLink = {
        Description: cdsNumber,
        Url: cdskUrl
    };

    //let title = fd.field('Title').value;

    let contractorHostURL = await getParameter("ContractorHostURL");

    let partTrades = fd.field('chklstStates').value;
    let concatParts = '', existingPartTrades = [], newPartTrades = [], canPartTrades = [], tradeCC = '';
    if (partTrades.length > 0){
        for(let trade of partTrades){
            if (trade !== ''){
  
                if(leadTrade === trade){
                    setPSErrorMesg(`Set ${leadTrade} as Lead or Part Trade but not both`);
                    return;
                }

                concatParts += trade + ",";
                let partMeta = await getTradeEmails(trade, 'CC', contractorHostURL);

                if(partMeta !== undefined)
                   tradeCC += partMeta.tradeEmails;

                let query = `Reference eq '${cdsRef}' and Trade eq '${trade}'`;
                let items = await _web.lists.getByTitle(_PartTask).items.filter(query).getAll();
                if(items.length === 0)
                  newPartTrades.push(trade);
                else {
                    let status = items[0].Status;
                    if(status === 'Cancelled')
                      canPartTrades.push({Id: items[0].Id, trade: trade});
                    else existingPartTrades.push({Id: items[0].Id, trade: trade});
                }
            }
       }
    }

    if (concatParts.endsWith(','))
       concatParts = concatParts.slice(0, -1);

    leadMeta = await getTradeEmails(leadTrade, 'CC', contractorHostURL);
    if(leadMeta !== undefined)
      tradeCC += leadMeta.tradeEmails;
    //#endregion

  let result = await createLeadTask(leadTrade, leadColumnName, contractorHostURL, reviewLink, concatParts, cdsLink);
  if(result === undefined){
    await setDimonButton('Assign', true)
    return;
  }
  let {subject, sendLeadEmail, isLeadValid, userEmail} = result;

  let partTradeEmails = await createPartTask(existingPartTrades, newPartTrades, canPartTrades, concatParts, partColumnName, contractorHostURL, cdsLink);

   await updateListItem(_RLOD, concatParts);
   await updateListItem(_SLOD, concatParts);

   await insertCDSItem(concatParts);


   //#region SEND MASTER EMAIL
   let newParts = newPartTrades.length > 0 ? newPartTrades.join(',') : '';
   let containParts = concatParts !== ''? 'yes' : 'no';
   let extraEmails = leadMeta !== undefined && leadMeta.extEmails !== undefined ? leadMeta.extEmails : '';

   let to = userEmail + partTradeEmails;
   if(to !== undefined && to !== ''){
    to = await removeDuplicates(to);
    if (to.endsWith(','))
        to = to.slice(0, -1);
   }

   if(cc !== ''){
    if(tradeCC !== '')
      cc += tradeCC + ','
   }
   else cc = tradeCC;

   if(cc !== undefined && cc !== ''){
    cc = await removeDuplicates(cc);
    if (cc.endsWith(',')){}
      //cc = cc.slice(0, -1);
    else cc += ',';
   }

   if(extraEmails !== undefined && extraEmails !== ''){
    extraEmails = await removeDuplicates(extraEmails);
    if (extraEmails.endsWith(','))
      extraEmails = extraEmails.slice(0, -1);
   }

   let getNotif = `${acronym}AssignCC`;
   sendMasterEmail(newParts, containParts, to, cc, extraEmails, getNotif);
   //#endregion

   fd.close();
}

const removeDuplicates = async function(values){
    let valueArray = values.split(',');

    let uniqueValues = valueArray.filter((value, index, self) => {
        return self.indexOf(value) === index;
    });

   return  uniqueValues.join(',');
}

const getTradeEmails = async function(trade, columnName, contractorHostURL){
    let isCCFound = false, isColumnNameFound = false;
    let tradeEmails = '', extEmails = '', users;
    let userArray = [];
    let _fields = await _web.lists.getByTitle(_Trades).fields.select("Title", "InternalName").get();
    _fields.map(field=>{
        let internalName = field.InternalName;

        if(internalName === 'CC')
          isCCFound = true;

        if(internalName === columnName)
          isColumnNameFound = true;
    });

    if(columnName === 'CC' && !isCCFound)
       return;
    else if(!isColumnNameFound)
      columnName = 'DWGLead'

    let query = `Title eq '${trade}'`;
    let items = await _web.lists.getByTitle(_Trades).items
                        .select(`${columnName}/Title, ${columnName}/EMail, ${columnName}/Id, DWGLead/Title, DWGLead/EMail, DWGLead/Id`)
                        .expand(`${columnName}, DWGLead`)
                        .filter(query).get();

    if(items.length > 0){
        let item = items[0];
        users = item[columnName] === undefined ? item['DWGLead'] : item[columnName];
        if(users !== null && users.length > 0){
            for(let i = 0; i < users.length; i++){
                let userProf = users[i];
                let userEmail = userProf.EMail;
                userEmail = userEmail !== undefined && userEmail !== null && userEmail !== undefined  ? userEmail : '';

                if(userEmail !== ''){
                  if(contractorHostURL !== ''){
                    if(userEmail.includes('@dar'))
                     tradeEmails += userEmail + ',';
                    else extEmails += userEmail + ',';
                  }
                  else tradeEmails += userEmail + ',';
                }

                userArray.push(userProf.Id);
            }
        }
    }

    return {
        tradeEmails: tradeEmails,
        extEmails: extEmails,
        users: userArray
    }
}

const createLeadTask = async function(leadTrade, leadColumnName, contractorHostURL, reviewLink, concatParts, cdsLink){
    let subject = '', sendLeadEmail = false, isLeadValid = false;
    let query = `Reference eq '${cdsRef}' and Status eq 'Open'`;

    let tradeResult= await getTradeEmails(leadTrade, leadColumnName, contractorHostURL);
    let items = await _web.lists.getByTitle(_LeadTask).items.filter(query).getAll();

    if(tradeResult !== undefined){
        if(items.length === 0){
            insertTask(_LeadTask, leadTrade, tradeResult.users, reviewLink, concatParts, cdsLink);
            sendLeadEmail = true;
        }
        else{
            let item = items[0];
            let status = item.Status;
            let currentTrade = item.Trade;

            if(status === 'Completed'){
                setPSErrorMesg("Kindly note that you can't assign anymore as Lead Task is Completed. Kindly Contact PWS Support to Open/Remove Lead Task so you can Assign Again");
                return;
            }
             
            let targetStatus = 'Open'
            if(leadTrade === currentTrade){
              if(status === 'Open')
                isLeadValid = true;
              else if(status === 'Cancelled'){
                subject =  `${projectName} - ${cdsRef} - Lead Action (${leadTrade}) - Re-Open Task`;
                await updateItem(_LeadTask, item.Id, {Status: 'Open', PartTrades: concatParts});
              }
            }
            else{
                subject =  `${projectName} - ${cdsRef} - Lead Action (${currentTrade}) - Cancelled Task`;
                await updateItem(_LeadTask, item.Id, {Status: 'Cancelled'});


                query = `Reference eq '${cdsRef}' and Trade eq '${leadTrade}'`;
                items = await _web.lists.getByTitle(_LeadTask).items.filter(query).getAll();
                if(items.length === 0)
                 insertTask(_LeadTask, leadTrade, tradeResult.users, reviewLink, concatParts, cdsLink); // insert new trade

                sendLeadEmail = true;
                targetStatus = 'Cancelled';
            }

            if(subject !== '')
              setLeadPartEmail(subject, targetStatus, currentTrade, tradeResult.tradeEmails, true, item.Id)
        }
    }
    return{
        subject: subject,
        sendLeadEmail: sendLeadEmail,
        isLeadValid: isLeadValid,
        userEmail: tradeResult.tradeEmails
    }
}

const createPartTask = async function(existingPartTrades, newPartTrades, canPartTrades, concatParts, partColumnName, contractorHostURL, cdsLink){
    let subject = '';
    let query = `Reference eq '${cdsRef}' and Status ne 'Cancelled'`;
    let items = await _web.lists.getByTitle(_PartTask).items.filter(query).getAll();
    let tradeResult, tradeEmails = '';
    if (newPartTrades.length > 0){
        for(let trade of newPartTrades){
            tradeResult= await getTradeEmails(trade, partColumnName, contractorHostURL);
            if(tradeResult !== undefined){
              insertTask(_PartTask, trade, tradeResult.users, null, '', cdsLink);
              tradeEmails += tradeResult.tradeEmails + ',';
            }
        }
    }

    if (canPartTrades.length > 0){
        for(let item of canPartTrades){
            let trade = item.trade;
            let itemId = item.Id;

            tradeResult = await getTradeEmails(trade, partColumnName, contractorHostURL);
            if(tradeResult !== undefined){
             subject =  `${projectName} - ${cdsRef} - Part Action (${trade}) - Re-Open Task`;
             await updateItem(_PartTask, itemId, {Status: 'Open'});
             setLeadPartEmail(subject, 'Open', trade, tradeResult.tradeEmails, false, itemId);
             if(tradeResult !== undefined)
                  tradeEmails += tradeResult.tradeEmails + ',';
            }
        }
    }

    if(items.length > 0){
        let splitTrades = concatParts.split(',');
        for (const item of items){
            let itemTrade = item.Trade;
            let isFound = splitTrades.filter(tempTrade => tempTrade === itemTrade)

            tradeResult= await getTradeEmails(itemTrade, partColumnName, contractorHostURL);
            if(tradeResult !== undefined)
              tradeEmails += tradeResult.tradeEmails;

            if(isFound.length === 0){
                let itemId = item.Id;
                await updateItem(_PartTask, itemId, {Status: 'Cancelled'});
                subject =  `${projectName} - ${cdsRef} - Part Action (${itemTrade}) - Cancelled Task`;
                setLeadPartEmail(subject, 'Cancelled', trade, tradeResult.tradeEmails, false, itemId);
            }
        }
    }
    return tradeEmails;
}

const insertTask = async function(listname, trade, users, reviewLink, partTrades, cdsLink){

    let item = {};
    item['Reference'] = cdsRef;
    item['Date'] = fd.field('Date').value;//new Date();
    item['SubmittalType'] = acronym;
    item['Title'] = fd.field('Title').value;

    item['Trade'] = trade;
    //item['Status'] = 'Open';
    // //item['AssignedTo'] = users;
     item['AssignedToId'] = { results: users }
    
    if(reviewLink !== null){
      item['ReviewLink'] = reviewLink;
      item['PartTrades'] = partTrades;
    }

    //item['CDSLink'] = cdsLink;

    //if(!doUpdate)
      await _web.lists.getByTitle(listname).items.add(item);
    //else await _web.lists.getByTitle(listname).items.getById(itemId).update(item);
}

const updateItem = async function(listname, itemId, fields){
    await pnp.sp.web.lists.getByTitle(listname).items.getById(itemId).update(fields);
}

const setLeadPartEmail = async function(subject, status, trade, userEmail, isLead, itemId){

    let body = '', to = '', cc = '';
    let displayTradeName = isLead ? 'Lead Action' : 'Part Action';
    let listname = isLead ? _LeadTask : _PartTask;
    if(status === 'Open'){

        let fields = { };
        fields[displayTradeName] = trade;
        fields['View Task'] = `<a href='${_webUrl}/SitePages/PlumsailForms/${listname}/Item/DisplayForm.aspx?item=${itemId}'>View Task</a>`;
        fields['Edit Task'] = `<a href='${_webUrl}/SitePages/PlumsailForms/${listname}/Item/EditForm.aspx?item=${itemId}'>Edit Task</a>`;;

        body += `<p style='font-family: Verdana; font-size: 13px; font-weight: bold;'>${listname} ${cdsRef} has been <b><u>Re-Open</u></b> for your review
                   <table border='0' width='60%' cellpadding='2' cellspacing='0' style='font-family: Verdana; font-size: 13px'>`;

        for (let key in fields){
          let value = fields[key];
          body += `<tr><td width='25%' style='border: 1px solid #E8EAEC; vertical-align: middle; font-weight: bold; color: #616A76;font-family: 
                     Verdana; font-weight: bold; font-size: 13px'>${key}</td>
                    <td style='border: 1px solid #E8EAEC; background-color: #F5F6F7;font-family: Verdana; font-size: 13px'>${value}</td></tr>`;
        }
        body += "</table> ";
    }
    else{ //Cancelled
        body = `<p style='font-family: Verdana; font-size: 13px;'>No needed action is required from your end for <b> ${cdsRef}</b>`;
    }
    to = userEmail;

    if (to.endsWith(','))
      to = to.slice(0, -1);

    cc = currentUser.Email;

    if(to !== '')
      sendEmail(subject, body, to, cc);
}

const sendEmail = async function(subject, body, to, cc){
    let method = 'SEND_ASSIGN_EMAIL';
    let serviceUrl = `${_siteUrl}/AjaxService/DarPSUtils.asmx?op=${method}`;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                        <soap:Body>
                            <${method} xmlns="http://tempuri.org/">
                                <WebURL>${_webUrl}</WebURL>
                                <Subject>${subject}</Subject>
                                <Body><![CDATA[${body}]]></Body>
                                <To>${to}</To>
                                <CC>${cc}</CC>
                            </${method}>
                        </soap:Body>
                    </soap:Envelope>`
    let res = await getSoapResponse('POST', serviceUrl, true, soapContent, `${method}Result`);
}

const sendMasterEmail = async function(NewPartTrades, containParts, to, cc, extEmails, getNotif){

    //SEND_MASTER_ASSIGN_EMAIL(string WebURL, string Reference, string NewPartTrades, string ContainParts, string To, string CC, string ExtEmails)
    let LocUrl = fd.field('Location').value !== null ? fd.field('Location').value.url : '';
    let method = 'SEND_MASTER_ASSIGN_EMAIL';
    let serviceUrl = `${_siteUrl}/AjaxService/DarPSUtils.asmx?op=${method}`;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                        <soap:Body>
                            <${method} xmlns="http://tempuri.org/">
                                <WebURL>${_webUrl}</WebURL>
                                <Reference>${cdsRef}</Reference>
                                <NewPartTrades>${NewPartTrades}</NewPartTrades>
                                <ContainParts>${containParts}</ContainParts>
                                <To>${to}</To>
                                <CC>${cc}</CC>
                                <ExtEmails>${extEmails}</ExtEmails>
                                <GetNotification>${getNotif}</GetNotification>
                                <LocUrl>${LocUrl}</LocUrl>
                            </${method}>
                        </soap:Body>
                    </soap:Envelope>`
    let res = await getSoapResponse('POST', serviceUrl, true, soapContent, `${method}Result`);
}

const updateListItem = async function(listname, concatPart){

    if(concatPart !== ''){
        let fields = await getListFields(_web, listname);
        let colsInternal = [];
        let splitTrades = concatPart.split(',');
        for (let itemTrade of splitTrades) {
          let columnName = `${itemTrade}_Part`;
          colsInternal.push(columnName);
        }
        validateListFields(listname, fields, colsInternal);
    }

    let query = `CDSNumber eq '${cdsRef}' and Status ne 'Reviewed'`
    let items = await _web.lists.getByTitle(listname).items.filter(query).getAll()
    
    if(items.length > 0){
        for (let item of items){
            let tempItem = {}
            tempItem['AssignedDate'] = new Date();

            let value = fd.field('LeadTradeDDL').value;
            if(typeof value === 'object'){
                if(value.length > 0)
                tempItem["LeadTrade"] = value[0];
            }
            else tempItem["LeadTrade"] = value;

            tempItem['Status'] = 'Assigned'
            tempItem['LeadStatus'] = 'Open'

            if(concatPart !== ''){
                let splitTrades = concatPart.split(',');
                for (let itemTrade of splitTrades) {
                    let columnName = `${itemTrade}_Part`;
                    tempItem[columnName] = "Open";
                  }
            }
            await _web.lists.getByTitle(listname).items.getById(item.Id).update(tempItem);
        }
    }
}

var getListFields = async function(_web, _listname){
    var fieldInfo = [];
    await  _web.lists
           .getByTitle(_listname)
           .fields
           .filter("ReadOnlyField eq false and Hidden eq false")
           .get()
           .then(function(result) {
              for (var i = 0; i < result.length; i++) {
                  fieldInfo.push(result[i].InternalName);
                  // fieldInfo += "Title: " + result[i].Title + "<br/>";
                  // fieldInfo += "Name:" + result[i].InternalName + "<br/>";
                  // fieldInfo += "ID:" + result[i].Id + "<br/><br/>";
              }
      }).catch(function(err) {
          alert(err);
      });
      return fieldInfo;
}

var validateListFields = async function(listname, fields, colsInternal){

    for(var i = 0; i < colsInternal.length; i++){
        var isFldExist = false;
        var _fld = colsInternal[i];

        for(var j = 0; j < fields.length; j++){
            var _lstFld = fields[j];
            if(_fld === _lstFld){
                isFldExist = true;
                break;
            }
        }
        if(!isFldExist){
            await createFields(_web, listname, _fld, 'text');
        }
    }
}

var createFields = async function(_web, _listname, _field, _type){
    try{
     _type = _type.toLowerCase();
     if(_type === 'text' || _type === 'dropdown')
        await _web.lists.getByTitle(_listname).fields.addText(_field);
     else if(_type === 'multi')
        await _web.lists.getByTitle(_listname).fields.addMultilineText(_field);
     else if(_type === 'calendar')
        await _web.lists.getByTitle(_listname).fields.addDateTime(_field);
        
        await _web.lists.getByTitle(_listname).defaultView.fields.add(_field);
    }
    catch (e) {
       console.log(e);
    }
}
//#endregion

//#region BIND DATA ON EDIT
const getData = async function(){
    data = [];
    cdsRef = fd.field('CDSNumber').value;
    let query = `CDSNumber eq '${cdsRef}'`;
    let items = await _web.lists.getByTitle(_SLOD).items.filter(query).getAll();
    let isHCEmpty = false;
    let isReviewed = true;
    for (const item of items){
        let rlodRev = item['Revision'] !== null ? item['Revision'] : '';
        
        if(!isHCEmpty)
         isHCEmpty = item['SiteHCReceivedDate'] !== null ? false : true;

        data.push({
            filename: item.FileName,
            revision: rlodRev,
            title: item.Title
        });

        if(fd.field('Location').value === null || fd.field('Location').value.description === ''){
            let locationUrl = await getLocation(item);
            let cdsLink = {
              description: cdsRef,
              url: locationUrl
            };
            fd.field('Location').value = cdsLink;
        }

        if(item['Status'] !== 'Reviewed')
          isReviewed = false;
    }

    if(isAssignAllowed){
      if(isReviewed){
        $(fd.field('LeadTradeDDL').$parent.$el).hide();
        $(fd.field('chklstStates').$parent.$el).hide();
        $(fd.field('Remarks').$parent.$el).hide();
        await addCrsReply();
      }
      else{
        await setLeadTrades();
        await setHCDate(isHCEmpty);
      }
    }
    else{
        $(fd.field('LeadTradeDDL').$parent.$el).hide();
        $(fd.field('chklstStates').$parent.$el).hide();
        $(fd.field('Remarks').$parent.$el).hide();
    }
}

const setLeadTrades = async function(){
    let isPermissionEnabled = false, isCategoryFound = false;

    let _fields = await _web.lists.getByTitle(_Trades).fields.select("Title", "InternalName").get();
    _fields.map(field=>{
        let internalName = field.InternalName;

        if(internalName === 'Permission')
          isPermissionEnabled = true;

        if(internalName === 'Category')
          isCategoryFound = true;
    });

    let query = `Title ne null`;
    if(isCategoryFound)
       query = `Title ne null and Category eq 'CDS'`;
    let items = await _web.lists.getByTitle(_Trades).items
                      .select('Title, Permission/Title')
                      .expand('Permission')
                      .filter(query)
                      //.orderBy('Title', true)
                      .getAll();

    let trades = [];
    for (const item of items){
        const {Permission} = item;
        if(Permission !== undefined && Permission.length > 0){
            for(let j = 0; j < Permission.length; j++){
                const group = Permission[j].Title;
                const isUserValid = await IsUserInGroup(group);
                if(isUserValid)
                  trades.push(item.Title);
            }
        }
        else trades.push(item.Title);
    }
    trades = trades.sort();
    fd.field('LeadTradeDDL').required = true;
    fd.field('LeadTradeDDL').widget.dataSource.data(trades);
    if(_isEdit){
        let leadTrade = fd.field('LeadTrade').value;
        if(leadTrade !== null && leadTrade !== ''){
            const leadObject = [];
            leadObject.push(leadTrade);
            fd.field('LeadTradeDDL').value = leadObject;
        }
    }

    let field = fd.field('chklstStates');
    field.options = trades;

    let partTrades = fd.field('PartTrade').value;
    if(partTrades !== null && partTrades !== ''){
        let selectedTrades = partTrades.split(',');
        field.value = selectedTrades;
    }
}

const checkAllowedUser = async function(){
    let isAllowed = false;
    if(allowedGroupAssigner !== ''){
        let splitAllowedGroupAssigner = allowedGroupAssigner.split(',');
        for(let i = 0; i < splitAllowedGroupAssigner.length; i++){
            let group = splitAllowedGroupAssigner[i];
            const isUserValid = await IsUserInGroup(group);
            if(isUserValid){
                isAllowed = true;
                assigneeGroupName = group;
                break;
            }
        }
    }
    return isAllowed
}

const setHCDate = async function(isHCEmpty){
    let isHCDateMandatory = await getParameter('isHCDateMandatory');
    if(isHCDateMandatory.toLowerCase() === 'yes' && isHCEmpty){
        setPSErrorMesg('Hard Copy Date must be filled before assignment');
        $('span').filter(function () { return $(this).text() == 'Assign'; }).parent().css('color', '#737373').attr("disabled", "disabled");
    }
}

const addCrsReply = async function(){
    let arrayFunctions = [
        getParameter('SubmittedDelivFolderName'),
        getParameter('ReviewededDelivFolderName')
    ];

    let [SubmittedDelivFolderName, ReviewedDeliverables] = await Promise.all(arrayFunctions);

    let {description, url} = fd.field('Location').value;
    url = url.replace(SubmittedDelivFolderName, ReviewedDeliverables)

    let html = `<label for="fd-field-a723a99c-3151-4989-bf3c-712c77c88387" title="" class="d-flex fd-field-title col-form-label col-sm-auto" style="float: left; padding-right: 52px;">
                        <span class="overflow-hidden fd-title-wrap">CRS Reply</span>
                </label>
                    
                <div data-v-50112471="" class="fd-sp-field-note col-form-label" style="display: inline-block;">
                  <a href="${url}" target='_blank'>${description}</a>
                </div>`;

    $('#crsReply').append(html);
}
//#endregion