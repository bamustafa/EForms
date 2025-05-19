var  _layout, _module = '', _rootFolderUrl = '', _container, _hot, _transNo, _data=[], _acronym, _ContCode, _isMain = true, _isLead = false, _isPart = false;
var fileIcons, pdfTronAllowedPreviewExt, pdfTronPreviewUrl = '/_layouts/15/pdf/viewer.aspx?f=';

let _web, _webUrl, _siteUrl, _formType = '', _errorImg = '', _submitImg = '';
let _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDesign = false, isTicked = true, _isMultiContractor = false, _isRename = false, isFirstFill = true,
    _CheckRevision = true, _exact_Filename_Match = false;
let _itemId, _itemListname, _htLibraryUrl, _submittalRef, _darTrade, enableTransmittal, isWorkHoursEnabled, bimRevFormat;

let previewWindow = null, checkPreviewInterval = null;
let screenHeight = screen.height - 500;

var _mtItem, _revNumStart, _revStartsWith = '', allowMultiTrade, checkAgainstFNC, _files = [], _fncSchemas= {}, delivTypeExtensions = {}, _masterErrors = [], clonResults = [],
     checkedFiles = [], docType = [], multiContractorValues = [], distinctDelivFilename = [], _bimRevFormat;
let contentType, filenameNewValue, clsoingStatus = 'reviewed';

let isSelectedColIndex;
var disableClassName = 'disableCheckBox';
let allowedExtensionsObj = [], CheckAllowedExtensionsObj = [], cdsFormDataArray = [];
let splitter = '-';
let rowSize = 25; // Select 30 items per batch
let updateBatchSize = 50;
let missingPairErrorsFiles = [];
let rlodFilenames;

let totalRowUpdates = 0, _trade = '';


var onRender = async function (relativeLayoutPath, moduleName, formType){
    const startTime = performance.now();

    _layout = relativeLayoutPath;
    await loadScripts();

    showPreloader();
    await renderControls();

    if(_isRename){
        localStorage.setItem('filenameOldValue', fd.field('FileLeafRef').value);
        $(fd.field('FolderName').$parent.$el).hide();
        $(fd.field('FilesNo').$parent.$el).hide();
        $('#search_field').remove();
        hidePreloader();
        return;
    }
    else $(fd.field('FileLeafRef').$parent.$el).hide();

    await getLODGlobalParameters(moduleName, formType);

    if(!_isDesign)
      await handleFeatures(enableTransmittal);

    await getFileNames();

    document.addEventListener("DOMContentLoaded", function () {
       document.querySelectorAll(".preview-link").forEach(link => {
        link.addEventListener("click", function (event) {
            event.preventDefault(); // Prevent default navigation
            previewFile(this);
        });
       });
    });

    hidePreloader();
    const endTime = performance.now();
    const elapsedTime = endTime - startTime;
    console.log(`Execution onRender: ${elapsedTime} milliseconds`);
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

var getLODGlobalParameters = async function(moduleName, formType){

    if($('.text-muted').length > 0)
      $('.text-muted').remove();

    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;
    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

    let arrayFunctions = [
        getParameter("EnableTransmittal"),
        getParameter("AllowMultiTrade"),
        getParameter('Phase'),
        getParameter('EnableWorkHours'),
        getParameter('CheckAgainst_FileConvention'),
        getParameter('CheckRevision'),
        getGridMType(_web, _webUrl, 'LOD', false),
        isMultiContractor() // _isMultiContracotr is already global variable there
    ];
    const params = await Promise.all(arrayFunctions);

    enableTransmittal = params[0].toLowerCase() === 'yes'? true : false;
    allowMultiTrade = params[1].toLowerCase() === 'yes'? true : false;
    _isDesign = params[2].toLowerCase() === 'design'? true : false;
    isWorkHoursEnabled = params[3].toLowerCase() === 'yes' ? true : false;
    checkAgainstFNC = params[4].toLowerCase() === 'yes' ? true : false;
    _CheckRevision = params[5].toLowerCase() === 'yes' ? true : false;
    _mtItem = params[5];
    _revNumStart = _mtItem.revNumStart;

    if(_isDesign){
      arrayFunctions = [];
      arrayFunctions.push(getParameter("BIMRevFormat"), getParameter("Exact_Filename_Match"), getClosingStatus());
      const params1 = await Promise.all(arrayFunctions);
     _bimRevFormat = params1[0]
     _exact_Filename_Match = params1[1].toLowerCase() === 'yes' ? true : false;
    }

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _itemListname = list.Title;

    contentType = fd.spFormCtx.RowData.ContentTypeId;

    _errorImg = _layout + '/Images/Error.png';
    _submitImg = _layout + '/Images/Submitted.png';

    if(_formType === 'New'){
        fd.clear();
        _isNew = true;
    }
    else if(_formType === 'Edit'){
        _isEdit = true;
        _itemId = fd.itemId;
    }

    let fileUrl = `${_webUrl}${_layout}/Images/Icons/`;
     fileIcons = {
      pdf: fileUrl + 'Pdf.png',
      doc: fileUrl + 'Word.png',
      docx: fileUrl + 'Word.png',
      xls: fileUrl + 'Excel.gif',
      xlsx: fileUrl + 'Excel.gif',
      ppt: fileUrl + 'Ppt.png',
      pptx: fileUrl + 'Ppt.png',

      rar: fileUrl + 'Zip.gif',
      zip: fileUrl + 'Zip.gif',
      folder: fileUrl + 'Folder.gif',
      attach: fileUrl + 'attach.gif',

      dwg: fileUrl + 'dwg.jpg',
      dwf: fileUrl + 'dwf.png',
      dwfx: fileUrl + 'dwfx.png',
      ifc: fileUrl + 'dwg.jpg',

      rvt: fileUrl + 'rvt.png',
      nwc: fileUrl + 'nwc.jpg',
      nwd: fileUrl + 'nwc.jpg',

      png: fileUrl + 'Png.png',
      gif: fileUrl + 'Png.png',
      jpeg: fileUrl + 'Jpeg.gif',
      jpg: fileUrl + 'Jpeg.gif',
      jfif: fileUrl + 'jfif.png',
      tif: fileUrl + 'tif.png',
      tiff: fileUrl + 'tif.png'
      // Add more file extensions and corresponding icons here
    };
    pdfTronAllowedPreviewExt = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'pdf', 'dwg','dwf', 'jpg', 'png' , 'gif', 'jp2' , 'jfif', 'tif', 'tiff']; // 'rvt'

    var types = ['DWG','DOC','TRM','GEN', 'MOD'];
    types.forEach(async function(type) {

        let query = `Title eq '${type}'`
        let extItems = await _web.lists.getByTitle(_AllowedExtensions).items.filter(query).get();
        if(extItems !== undefined && extItems !== null && extItems.length > 0){
            let item = extItems[0];
            let AllowedExt = item['AllowedExtensions'] !== null ? item['AllowedExtensions'] : '';
            delivTypeExtensions[type] ={
                AllowedExt: AllowedExt
            };
        }

        var items = await _web.lists.getByTitle("FNC").items.select("Delimeter,Schema").filter(`Title eq '${type}'`).get();
        if(items.length > 0){
            let item = items[0];
            let delimeter = item.Delimeter;

            let schema = item.Schema;
            schema = schema.replace(/&nbsp;/g, '');
            schema = JSON.parse(schema);

            let containsRev = false;
            let revStartWith;
            schema.map(fld => {
                let fieldName = fld.InternalName.toLowerCase();
                if(fieldName === 'rev' || fieldName === 'revision'){
                  containsRev = true;
                  revStartWith = fld.RevStartWith;
                  return true;
                }
            });

            let enableTrans = false, isTransRequired = false;
            if(!_isDesign){
                let majorItems = await _web.lists.getByTitle("MajorTypes").items.select('EnableTransmittal,TransmittalRequired').filter(`Title eq '${type}'`).get();
                if(majorItems.length > 0){
                    enableTrans = majorItems[0].EnableTransmittal;
                    isTransRequired = majorItems[0].TransmittalRequired;
                }
            }

            _fncSchemas[type] = {
                delimeter: delimeter,
                schemaFields: schema, // replace with actual values
                containsRev: containsRev,
                revStartWith: revStartWith,
                enableTransmittal: enableTrans,
                isTransRequired: isTransRequired
            };
        }
    });

    // let retry = 0;
    //        var intervalId = setInterval(() => {
    //         if(retry >= 10 )
    //           clearInterval(intervalId);

    //           if(ctlrForm.length > 0){
    //             let padding = ctlrForm[0].style.padding;
    //             if (padding !== undefined && padding !== null && padding !== '') {
    //                 ctlrForm.css('padding', '');
    //                 clearInterval(intervalId); // Provide interval ID to clearInterval()
    //             }
    //           }
    //         retry++;
    //     }, 1000);
}

var renderControls = async function () {
    await setFormHeaderTitle();
    await setButtons();


    //var fields = ['Title','Status'];
    //HideFields(fields, true);
    setToolTipMessages();

    var queryString = window.location.search;
	var urlParams = new URLSearchParams(queryString);
	var item = urlParams.get('ct');
    _isRename = item !== undefined && item !== null ? true : false;
}

const setValueFields = async function(fileCount){
    let readOnlyClass = {
        'background-color': 'Transparent',
        'border': 0,
        'pointer-events': 'none'
    }

    $("div[title='Folder Name']").find('input').prop('readonly', true).css(readOnlyClass);

    fd.field('FilesNo').value = fileCount;
    $("div[title='Number of Files']").find('input').prop('readonly', true).css(readOnlyClass);
}

function setToolTipMessages(){
    //setButtonToolTip('Save', saveMesg);
    setButtonCustomToolTip('Submit', submitMesg);
    setButtonCustomToolTip('Cancel', cancelMesg);

    //else addLegend();
    //adjustlblText('Comment', ' (Optional)', false);

    //   if($('p').find('small').length > 0)
    //   $('p').find('small').remove();
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
           if(text == "Submit"){
                if(_isRename){
                    localStorage.setItem('filenameNewValue', fd.field('FileLeafRef').value);
                    fd.save();
                    window.close();
                }
                else{

                    let formUrl = `${_webUrl}/SitePages/PlumsailForms/${_CDS}/Item/NewForm.aspx`;

                    if(!_isDesign){
                        let result = await handelConstructionSubmit();

                        if(result.errorFound)
                            setErrorMessage(result.mesg);
                        else{
                            const json = JSON.stringify(cdsFormDataArray);
                            localStorage.setItem('data', json);
                            localStorage.setItem('acronym', _acronym);
                            localStorage.setItem('trade', _trade);

                            if(_transNo !== undefined)
                              localStorage.setItem('transNo', _transNo);
                            if(_ContCode !== undefined)
                              localStorage.setItem('contCode', _ContCode);

                            window.location.href = formUrl;
                        }
                    }
                    else {
                        localStorage.setItem('data', _data);
                        localStorage.setItem('submittalRef', _submittalRef);
                        await handelDesignSubmit();
                        window.close();
                    }
                }

           }
          }
    });
}

const creatErrorlbl = async(errorText) =>{
    var id = '#errorlbl';
    if ($(id).length === 0) {
      var label = $('<label>', {
        id: 'errorlbl',
        text: ''
      });
      label.css('color', 'red');
      label.css('width', '950px');
      label.css('font-size', '16px');

      $('#dt').after(label);
    }
    $(id).html(errorText);

    $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#737373').attr("disabled", "disabled");
    //else $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#444').removeAttr('disabled');
}
//#endregion

//#region GET DATA
const getFileNames = async function(){
    let folder, serverRelativeUrl;

    const list = _web.lists.getByTitle(_itemListname);
    let item = await list.items.getById(_itemId).get();

    if(item.FileSystemObjectType === 1) { //folder Type
         folder = await list.items.getById(_itemId).folder.get();
         serverRelativeUrl = folder.ServerRelativeUrl;
    } else { // file Type
        const file = await list.items.getById(_itemId).file.get();
        serverRelativeUrl = file.ServerRelativeUrl;
        serverRelativeUrl = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/')+1);
      }

      await setFolderRelativeUrl(serverRelativeUrl);

      if(_isDesign){
        let error = await checkTaskStatus();
        if(error !== ''){
            setErrorMessage(error);
            validateButtons(true);
            $(fd.field('FolderName').$parent.$el).hide();
            $(fd.field('FilesNo').$parent.$el).hide();
            preloader('remove')
            return;
        }
      }

      _files = await getAllFilesInFolderAndSubfolders(serverRelativeUrl);

      let rowLimit = 10000, row = -1;
      const startTime = performance.now();

      const promises = _files.map(async (file, rowIndex) =>{
            row++
            if(row === rowLimit || row > rowLimit)
                return Promise.resolve(undefined);

            let filename = file.Name.toUpperCase();
            let extension = filename !== undefined && filename.includes('.') ? filename.substring(filename.lastIndexOf(".") + 1).toLowerCase(): '';
            if(extension !== '')
                filename = filename.substring(0, filename.lastIndexOf("."));

            return{
                rowIndex: rowIndex,
                ServerRelativeUrl: file.ServerRelativeUrl,
                extension: extension,
                filename: filename,
                filenameWithExt: `${filename}.${extension}`,
                mesg: ''
            }
      });

       await Promise.all(promises)
       .then(results =>{
            return results.filter((item)=>{
                return item !== undefined
            });
       })
       .then(async (filterResults) =>{
            const batchSize = filterResults.length < rowSize ? filterResults.length: rowSize;
            const fileLength = filterResults.length;

            if(fileLength < updateBatchSize)
              updateBatchSize = fileLength;

            processArrayInBatches(filterResults, batchSize, undefined, fileLength) // load is here
            .then(results => {
                processArrayInBatches(results.result, batchSize, results.rows, fileLength) // last batch if found
                .then(()=> {
                    _spComponentLoader.loadScript(_htLibraryUrl).then(_setData);
                    const endTime = performance.now();
                    const elapsedTime = endTime - startTime;
                    console.log(`Execution time afterSet: ${elapsedTime} milliseconds`);
                })
                .then(async ()=>{
                    let isErrorFound = false;
                    if(_masterErrors !== undefined && _masterErrors.length === 0){
                      if(multiContractorValues.length > 1){
                         isErrorFound = true;
                         let tempMesg = `Can't Submit more than one Category, as the below filenames contains ${multiContractorValues.length} categories (${multiContractorValues.join()})`
                         setErrorMessage(tempMesg);
                      }
                      else{
                        if(_isDesign){
                            rlodFilenames = await getPending_RLOD_FullFilenames();
                            if(rlodFilenames === undefined){
                                isErrorFound = true;
                                let tempMesg = `The below filenames under submission Reference ${_submittalRef} are not submitted under trade ${_darTrade}.<br /><br />`;
                                setErrorMessage(tempMesg);
                            }
                            else{
                                let difference = rlodFilenames.filter(item => !distinctDelivFilename.includes(item));
                                if(difference.length > 0){
                                    isErrorFound = true;
                                    let tempMesg = `The below filename ${difference.length} are missing while they are listed in RLOD Under Reference ${_submittalRef},
                                                    please upload through ${_itemListname}.<br /><br />`;

                                    difference.map(item =>{tempMesg += `${item}<br/>`})
                                    setErrorMessage(tempMesg);
                                }
                            }
                        }
                        else {
                            let res = await handelConstructionSubmit(); // for construction
                            isErrorFound = res.errorFound;
                            if(isErrorFound){
                                setErrorMessage(res.mesg);
                            }
                            //here to check master error message on top
                        }
                      }
                    }
                    else if(_masterErrors.length > 0) isErrorFound = true;
                    validateButtons(isErrorFound);
                    $('#totalId').remove();
                    preloader('remove');
                })
            })
            .catch(error => {
                console.log('Error:', error);
            });
        })
        await setValueFields(_files.length);
}

async function getAllFilesInFolderAndSubfolders(folderUrl) {
    try {
        const folder = await _web.getFolderByServerRelativeUrl(folderUrl);
        const files = await folder.files.get();

        let allFiles = [...files];

        const subfolders = await folder.folders.get();
        const subfolderPromises = subfolders.map(subfolder => getAllFilesInFolderAndSubfolders(subfolder.ServerRelativeUrl));

        const subfolderFiles = await Promise.all(subfolderPromises);

        subfolderFiles.forEach(files => {
            allFiles = allFiles.concat(files);
        });

        return allFiles;
    } catch (error) {
        console.error("Error getting files:", error);
        throw error;
    }
}

const setFolderRelativeUrl = async function(relativeUrl){
    folderUrl = relativeUrl.substring(relativeUrl.indexOf(`/${_itemListname}/`) + 1);
    folderUrl = folderUrl.replace(`${_itemListname}`, '');

    if(folderUrl.includes('/')){
      let folderArray = folderUrl.split('/');
      if(_isDesign){
        _submittalRef = folderArray[1];
        _darTrade = folderArray[2];
      }
      _rootFolderUrl = !_isDesign ? `${_webUrl}/${_UploadDeliverables}/${folderArray[1]}` : `${_webUrl}/${_Deliverables}/${_submittalRef}/${folderArray[2]}`;
      fd.field('FolderName').value = !_isDesign ? folderArray[1] : _darTrade;
    }
    return _rootFolderUrl;
}

const selectFilesInBatchV1 = async function(list, filenames, results, splitRev){
    let filterQuery = '';
    let dumpRevObj = {};

    // filenames.map((filename) => {
    //     let isTypeContainsRev = false;
    //     if (_CheckRevision || splitRev) { // _CheckRevision for design

    //        let rev = filename.substring(filename.lastIndexOf(splitter) + 1);
    //        filename = filename.substring(0, filename.lastIndexOf(splitter));
    //        dumpRevObj[filename] = rev;
    //     }
    //     filterQuery += `FileName eq '${filename}' or `;
    // });


    for (const filename of filenames) {

        let isTypeContainsRev = false;
        let updatedFilename = filename; // Keep original filename reference

        if (_CheckRevision || splitRev) {
            let item = await list.items.select('DeliverableType').filter(`FileName eq '${filename}'`).get();
            if (item.length === 0) {
                let tempFilename = filename.substring(0, filename.lastIndexOf(splitter));
                item = await list.items.select('DeliverableType').filter(`FileName eq '${tempFilename}'`).get();
            }

            if (item?.length > 0) {
                let docType = item[0].DeliverableType;
                isTypeContainsRev = _fncSchemas[docType]?.containsRev ? true: false;
            }

            if (isTypeContainsRev) {
                let rev = updatedFilename.substring(updatedFilename.lastIndexOf(splitter) + 1);
                updatedFilename = updatedFilename.substring(0, updatedFilename.lastIndexOf(splitter));
                dumpRevObj[updatedFilename] = rev;
            }
        }

        filterQuery += `FileName eq '${updatedFilename}' or `;
    }

    if(filterQuery !== '')
        filterQuery = filterQuery.substring(0, filterQuery.lastIndexOf(' or'));
    else return;

    let columns = 'FileName,DeliverableType,Revision,Trade,Status,Title';
    if(!_isDesign){
        columns += ',SubDiscipline,Contractor';
    }

   return await list.items.select(columns).filter(filterQuery).get()
    .then((items)=>{
        if(items.length > 0){

          let isFileFound = true;
          let filename, delivType, containsRev, trade, contCode, rlodRev, roldStatus, title, filenameRev;
          for(let item of items){
            filename = item.FileName.toUpperCase();

            filenames = filenames.filter(arrayFilename => {
                if(_CheckRevision || splitRev)
                  arrayFilename = arrayFilename.substring(0, arrayFilename.lastIndexOf(splitter));
                return arrayFilename !== filename
            });

            isFileFound = true;

            delivType = item['DeliverableType'] !== null ? item['DeliverableType'] : 'Empty';
            containsRev = item['Revision'] !== null ? true : false;

            trade = item['Trade'] !== null ? item['Trade'] : '';
            if(_trade === '')
             _trade = trade;

            if(!_isDesign){
                if(trade === '')
                trade = item['SubDiscipline'] !== null ? item['SubDiscipline'] : '';

                if(_isMultiContractor){
                    contCode = item['Contractor'] !== null ? item['Contractor'] : '';
                    if(contCode !== '' && !multiContractorValues.includes(contCode))
                        multiContractorValues.push(contCode);
                }
            }
            if(containsRev){
                rlodRev = item['Revision'] !== null ? item['Revision'] : '';
                filenameRev = dumpRevObj[filename] !== undefined ? dumpRevObj[filename] : '';
            }
            roldStatus = item['Status'] !== null ? item['Status'].toLowerCase() : 'pending';
            title = item['Title'] !== null ? item['Title'].toLowerCase() : 'NA';

            let tempObj = _fncSchemas[delivType];
            let revStartWith;
            if(tempObj !== undefined && tempObj !== null && tempObj.length > 0)
                revStartWith = tempObj.revStartWith;

                cdsFormDataArray.push({
                    filename: filename,
                    revision: filenameRev,
                    title: title
                    //isFileFound: isFileFound,
                    //delivType: delivType,
                    //trade: trade,
                    //contCode: contCode,
                    //containsRev: containsRev,
                    //revStartWith: revStartWith,

                    //roldStatus: roldStatus,
            });

            results = results.map(fileMetaInfo=>{
                let arrayFilename = fileMetaInfo.filename.toUpperCase();

                if (distinctDelivFilename.indexOf(arrayFilename) === -1)
                  distinctDelivFilename.push(arrayFilename);

                if(containsRev)
                  arrayFilename = arrayFilename.substring(0, arrayFilename.lastIndexOf(splitter));
                if( arrayFilename === filename){
                    return{
                        ...fileMetaInfo,
                        isFileFound: isFileFound,
                        delivType: delivType,
                        title: title,
                        trade: trade,
                        contCode: contCode,
                        containsRev: containsRev,
                        revStartWith: revStartWith,
                        rlodRev: rlodRev,
                        roldStatus: roldStatus
                        }
                }
                else return fileMetaInfo
            })
          }
        }
        return {
            results: results,
            filenames: filenames
        }
    })
    // .then(results=>{
    //     return results;
    // })
}

const processBatch = async (batchItems, batchfileNames, rows) => {
   const list = await _web.lists.getByTitle(_RLOD);
   const totalFiles = fd.field('FilesNo').value;

    let batchSelectPromise = await selectFilesInBatchV1(list, batchfileNames, batchItems, true);
    if(batchSelectPromise === undefined)
        return;

    batchItems = batchSelectPromise.results;
    batchfileNames = batchSelectPromise.filenames;

    if(batchfileNames.length > 0){
        let batchSelectPromise1 = await selectFilesInBatchV1(list, batchfileNames, batchItems, false);
        batchItems = batchSelectPromise1.results;
    }

    const bacthValidation = batchItems.map(async (result) =>{
        detailedLoader(totalFiles, rows);
        return validateFileName(result)
    })
    await Promise.all(bacthValidation)
    .then(validatedResults =>{
        validatedResults.map( (validatedItem, rowIndex) =>{
                let rowData = {};
                let { delivType, mesg, status, title, extension, ServerRelativeUrl, filename } = validatedItem;
                let previewLink =  'not supported';

                let _value = '';
                if(extension !== undefined){
                if(pdfTronAllowedPreviewExt.includes(extension))
                    //previewLink = `<a href='${_webUrl}/${pdfTronPreviewUrl}${ServerRelativeUrl}' onclick='previewFile(this);'>Preview Link</a>`;
                    previewLink = `<a href='${_webUrl}/${pdfTronPreviewUrl}${ServerRelativeUrl}' target='_Blank'>Preview Link</a>`;

                    _value = "<table align='center'>";
                    let iconUrl = fileIcons[extension];
                    _value += "<tr>" +
                                "<td style='border: none'><img src='" + iconUrl + "' alt='na'></img></td>" +
                                "</tr>";
                    _value += "</table>";
                }

                rowData['isSelected'] = 'yes';
                rowData['type'] = _value;
                rowData['extension'] = extension;
                rowData['title'] = title;
                rowData['filename'] = filename;
                rowData['DelivType'] = delivType;
                rowData['_mesg'] = mesg;
                rowData['_status'] = status;
                rowData['previewLink'] = previewLink;
                rowData['fileRelativeUrl'] = ServerRelativeUrl;
                _data.push(rowData);

                if(mesg !== undefined && mesg !== null && mesg !=='' && !mesg.toLowerCase().includes('cancelled')){
                    _masterErrors.push({
                        rowIndex: rowIndex,
                        ServerRelativeUrl: ServerRelativeUrl,
                        filename: filename,
                        filenameWithExt: `${filename}.${extension}`,
                        mesg: mesg
                    })
                }
        });
    })
    .catch(err => {
        console.log('Error in selectFilesInBatch:', err);
        throw err;
    });
}

const processArrayInBatches = async (array, batchSize, rowsNo, filesLength) => {
    const results = [];
    let rows = 0
    let Original;
    for (let i = 0; i <= array.length; i++) {
        let batchfileNames = [];
        let batch = [];

        array.forEach((result) => {

            if (batchfileNames.length >= batchSize)
                return;

            //if (!batchfileNames.includes(result.filename)) {
            if (batchfileNames.indexOf(result.filename) === -1){
                if(batchfileNames.length <= batchSize){
                    batchfileNames.push(result.filename);
                    let row = array.filter((filenameResult) => { return result.filename === filenameResult.filename});
                    array = array.filter(filenameResult => result.filename !== filenameResult.filename);

                    if(row.length > 0)
                    batch.push(...row);
                }
            }
        });
        if(rowsNo !== undefined)
          rows = filesLength; //rowsNo + array.length;
        else rows = filesLength - array.length; //updateBatchSize // 50++

        const batchResults = await processBatch(batch, batchfileNames, rows);
        //console.log(array.length);
    }
    return {
        result: array,
        rows: rows
    }
}

const validateFileName = async function(result){
    let fileMesg = '', fileStatus, filenameDelivType = 'NA', filenameTitle;
    let {filenameWithExt, filename, isFileFound, extension, ServerRelativeUrl, delivType, title, containsRev, revStartWith, rlodRev, roldStatus} = result;

    if(filename.includes(' ')){
        let ahrefElement = await renameFilenameLink(filenameWithExt, filename);
        fileMesg = `${spaceFileName} ${ahrefElement}`;
    }
    else if (extension === '')
        fileMesg = extFileName;
    else{
        // REGISTER FILES THAT U CHECKED AGAINST DATABASE TO REMOVE DUPLICATE CHECK
        //let {isFileFound, delivType, title, containsRev, revStartWith, rlodRev, roldStatus}= await getFileName(filename, splitter);


        // let fileInfo = await getFileName(filename, splitter);
        // isFileFound = fileInfo.isFileFound;
        // delivType = fileInfo.delivType;
        // title = fileInfo.title;
        // containsRev = fileInfo.containsRev;
        // revStartWith = fileInfo.revStartWith;
        // rlodRev = fileInfo.rlodRev;
        // roldStatus = fileInfo.roldStatus;



        filenameDelivType = delivType;
        filenameTitle = title;

       if(!isFileFound)
        fileMesg = `${filename} is not available in ${_RLOD}, Please upload through Register Deliverables`;
       else{
           if(roldStatus.toLowerCase() === 'cancelled'){
            fileMesg = 'This reference has been marked as <b>Cancelled</b> in the RLOD list'
            let partContainer = document.getElementById('dt');
            let masterMesg = `<p style='color:#ce830e; font-size:14px; font-weight:bold; padding:7px'>
                                Attention: There are references marked as cancelled. If you proceed submitting, they will be automatically un-cancelled and sent to PM.
                                Alternatively, you can go back and unselect the cancelled references.</p>`

             if ($('#insplbl').length === 0)
              addLegend('insplbl', masterMesg, partContainer, 'before', true);
           }
          //let containsRev = result.containsRev;
          else if(containsRev){
            //let rlodRev = result.rlodRev;
            let fileRev = filename !== undefined ? filename.substring(filename.lastIndexOf(splitter) + 1).toLowerCase(): '';
            let ahrefElement = await renameFilenameLink(filenameWithExt, filename);

            //fileMesg = await checkRevision(ahrefElement, fileRev, rlodRev, result.revStartWith, result.roldStatus);
            fileMesg = await checkRevision(ahrefElement, fileRev, rlodRev, revStartWith, roldStatus);
          }
       }
    }

    if(fileMesg === ''){
        filenameWithExt = filenameWithExt.toLowerCase();
        await CheckAllowedExtensions(filename, extension, filenameDelivType, false);
        let getRows = CheckAllowedExtensionsObj.filter(error =>{ return error.filenameWithExt === filenameWithExt });
        if(getRows !== undefined && getRows.length > 0)
            fileMesg = getRows[0].statusMesg !== undefined ? getRows[0].statusMesg : '';
        else {
            await CheckAllowedExtensions(filename, extension, filenameDelivType, false);
            let getRows = CheckAllowedExtensionsObj.filter(error =>{ return error.filenameWithExt === filenameWithExt });
            if(getRows.length > 0)
              fileMesg = getRows[0].statusMesg !== undefined ? getRows[0].statusMesg : '';
        }
    }
    fileStatus = setErrorMesg(fileMesg);
    return{
        delivType: filenameDelivType,
        title: filenameTitle,
        mesg: fileMesg,
        status: fileStatus,
        extension: extension,
        ServerRelativeUrl: ServerRelativeUrl,
        filename: filename
    }
}

const getFileName = async function(filename, splitter){
    let delivType = 'NA', trade, contCode, rlodRev, roldStatus, title;
    let isFileFound = false, containsRev = false;

    let query = `FileName eq '${filename}'`
    let items = await _web.lists.getByTitle(_RLOD).items.filter(query).get();
    if(items.length === 0){
        filename = filename.substring(0, filename.lastIndexOf(splitter));
        query = `FileName eq '${filename}'`
        items = await _web.lists.getByTitle(_RLOD).items.filter(query).get();
    }

    if(items.length > 0){
        isFileFound = true;
        let item = items[0];
        delivType = item['DeliverableType'] !== null ? item['DeliverableType'] : 'Empty';
        containsRev = item['Revision'] !== null ? true : false;
        trade = item['Trade'] !== null ? item['Trade'] : '';

        if(!_isDesign){
            if(trade === '')
            trade = item['SubDiscipline'] !== null ? item['SubDiscipline'] : '';

            if(_isMultiContractor){
                contCode = item['Contractor'] !== null ? item['Contractor'] : '';
                if(contCode !== '' && !multiContractorValues.includes(contCode))
                  multiContractorValues.push(contCode);
            }
        }

        if(containsRev){
            rlodRev = item['Revision'] !== null ? item['Revision'] : '';
        }
        roldStatus = item['Status'] !== null ? item['Status'].toLowerCase() : 'pending';
        title = item['Title'] !== null ? item['Title'].toLowerCase() : 'NA';
    }

    let tempObj = _fncSchemas[delivType];
    let revStartWith;
    if(tempObj !== undefined && tempObj !== null && tempObj.length > 0)
        revStartWith = tempObj.revStartWith;

    return{
        isFileFound: isFileFound,
        delivType: delivType,
        title: title,
        trade: trade,
        contCode: contCode,
        containsRev: containsRev,
        revStartWith: revStartWith,
        rlodRev: rlodRev,
        roldStatus: roldStatus,
    }
}

function setFilename(item){
    return {
        isFileFound : item.isFileFound,
        delivType : item.delivType,
        title : item.title,
        containsRev : item.containsRev,
        revStartWith : item.revStartWith,
        rlodRev : item.rlodRev,
        roldStatus : item.roldStatus
    }
}

const renameFilenameLink = async function(filenameWithExt, filename){
    let fileItems = await _web.lists.getByTitle(_itemListname).items.filter(`FileLeafRef eq '${filenameWithExt}'`).get();
    if(fileItems !== undefined && fileItems !== null && fileItems.length > 0){
        let item = fileItems[0];
        let linkUrl = `${_webUrl}/SitePages/PlumsailForms/${_itemListname}/Document/EditForm.aspx?item=${item.Id}&ct` //=${contentType}`; //0x01010088207B89203E9F46A97B692C834A7689
        return `<a href='${linkUrl}' target='_blank' onclick='previewFile(this, true);'>${filename}</a>`;
    }
    else return '';
}

const checkRevision = async function(ahrefElement, fileRev, rlodRev, revStartWith, roldStatus){
    let mesg = '';

    if(revStartWith !== undefined && revStartWith !== null){
        fileRev = fileRev.replace(revStartWith, '');
        rlodRev = rlodRev.replace(revStartWith, '');
    }
    else revStartWith = '';

    if(roldStatus === 'pending' && fileRev != rlodRev)
        mesg = `First Submission for ${ahrefElement} Should be ${revStartWith}${rlodRev}`;

    else if( (roldStatus === 'cancelled' || roldStatus === 'rejected') && fileRev != rlodRev)
        mesg = `${ahrefElement} with revision  ${revStartWith}${fileRev} is ${roldStatus}, Kindly you need to submit same Revision as in RLOD which is ${revStartWith}${rlodRev}`;

    else if( roldStatus === clsoingStatus){
       if(fileRev == rlodRev)
         mesg = `File ${ahrefElement} is already Reviewed`;
        else {
            if(isAllDigits(fileRev) && isAllDigits(rlodRev)){
                let targetRev = parseInt(rlodRev) + 1;
                if(rlodRev.length > 1 && targetRev < 10)
                targetRev = targetRev.toString().padStart(2, '0');

                if(fileRev < rlodRev || (fileRev > rlodRev && parseInt(fileRev - rlodRev) !== 1) )
                 mesg = `${ahrefElement} should be submitted as Revision ${revStartWith}${targetRev} instead of ${revStartWith}${fileRev} as RLOD Revision is ${revStartWith}${rlodRev}`;
            }
            else if(!isAllDigits(fileRev) && !isAllDigits(rlodRev)){
                const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
                let fileRevIndex = alphabet.indexOf(fileRev.toUpperCase());
                let rlodRevIndex = alphabet.indexOf(rlodRev.toUpperCase());
                let result = fileRevIndex - rlodRevIndex;

                let targetRev = String.fromCharCode(rlodRev.charCodeAt(0) + 1); // Convert the incremented character code to a string
                if(result !== 1)
                 mesg = `${ahrefElement} should be submitted as Revision ${revStartWith}${targetRev} instead of Revision ${revStartWith}${fileRev.toUpperCase()} as RLOD Revision is ${revStartWith}${rlodRev}`;
            }
            else if( (isAllDigits(fileRev) && !isAllDigits(rlodRev)) || (!isAllDigits(fileRev) && isAllDigits(rlodRev)) )
              mesg = `Can't submit Revision ${revStartWith}${fileRev} While Revision in RLOD is ${revStartWith}${rlodRev} for ${ahrefElement}`;
        }
    }

    else{
         //If roldStatus != Pending && roldStatus != ClosingStatus(Reviewed)
         if(roldStatus !== 'pending'){
            if((roldStatus !== 'sent to pm' || roldStatus !== 'sent to pmc') && fileRev == rlodRev)
                mesg = 'File is already under Dar Review';
            else mesg = `Can't submit new Rev while Revision ${revStartWith}${rlodRev} still under Review`;
         }
    }
    return mesg;
}

function isAllDigits(str) {
    return /^\d+$/.test(str);
}

const CheckAllowedExtensions = async function(filename, extension, delivType, singleCheck, filenames){
    let result, typeExtension = delivTypeExtensions[delivType];
    if(typeExtension !== undefined && typeExtension.hasOwnProperty('AllowedExt')){
        let AllowedExt = typeExtension.AllowedExt;
        result = await isCorrectExt(delivType, filename, extension, AllowedExt, singleCheck, filenames);
    }
    return result;
}

const isCorrectExt = async function(delivType, filename, filenameExt, AllowedExt, singleCheck, filenamesToCheck){
  if (checkedFiles.includes(filename))
    return '';

  let mesg = '';
  let filenames = _files.filter((file) =>{
    let fullfilename = file.Name;
    let tempFile = fullfilename.substring(0, fullfilename.lastIndexOf("."));
    let extension = fullfilename !== undefined && fullfilename.includes('.') ? fullfilename.substring(fullfilename.lastIndexOf(".") + 1).toLowerCase(): '';

    if(singleCheck){
        if(filenamesToCheck !== null && filenamesToCheck.length > 0){
            for(let filenameWithExt of filenamesToCheck){
              if(fullfilename === filenameWithExt)
                return true;
            }
        }
        else return filename.toLowerCase() === tempFile.toLowerCase() && filenameExt.toLowerCase() !== extension;
    }
    else return filename.toLowerCase() === tempFile.toLowerCase();
  });

  let splitRules = AllowedExt.toLowerCase().split(';');
  let allowedExtensionsMesg = '';

  //let resultExtensions = allowedExtensionsObj[delivType];

    let isMatched = false, isPair = false;
    splitRules.some(rule =>{
        ruleExt = rule.toLowerCase().split('|');
        if(!isMatched)
          isMatched = ruleExt.length === filenames.length ? true: false

        if (!allowedExtensionsObj.hasOwnProperty(delivType))
        {
            for(let ext of ruleExt){
                if(allowedExtensionsMesg === ''){
                    if(ext.includes('-')){
                      ext = `(<b>${ext}</b>)`;
                      allowedExtensionsMesg += ext.replaceAll('-', ' or ');
                    }
                    else if(!allowedExtensionsMesg.includes(ext))
                    allowedExtensionsMesg += ext + ' and ';
                }

                else {
                        if(ext.includes('-')){
                          ext = `(<b>${ext}</b>)`;
                          allowedExtensionsMesg += ext.replaceAll('-', ' or ');
                        }
                        else if(!allowedExtensionsMesg.includes(ext))
                        allowedExtensionsMesg += ext + ' and ';
                }
            }
            allowedExtensionsObj[delivType] = { AllowedExt: allowedExtensionsMesg };
        }
        else allowedExtensionsMesg = allowedExtensionsObj[delivType].AllowedExt;
    });


 if(!isMatched){
    let filenameWithExt = `${filename}.${filenameExt}`.toLowerCase();
    let mesg = `Missing/Wrong File Pair. Missing Extension as the allowed extension are ${allowedExtensionsMesg}`;

    if(singleCheck){
        return{
            filename: filenameWithExt,
            statusMesg: mesg,
        }
    }

    CheckAllowedExtensionsObj.push({
        filenameWithExt: filenameWithExt,
        filename: filename.toLowerCase(),
        statusMesg: mesg
    });
 }
 else{
    let isAllCorrect = true;
    filenames.map(item => {
        let singleFileMesg = '';
        let filenameWithExt = item.Name.toLowerCase();
        let filename = filenameWithExt.substring(0, filenameWithExt.lastIndexOf(".")).toLowerCase();
        let extension = filenameWithExt.substring(filenameWithExt.lastIndexOf(".") + 1).toLowerCase();

        let isFound = splitRules.some(rule =>{
            ruleExt = [];

            let isFound = false;

            ruleExt = rule.toLowerCase().split('|');
            for(let ext of ruleExt){
                if(ext.includes(extension)){
                isFound = true;
                return true;
                }
            }
        });

        if(!isFound){
            singleFileMesg = `Pair Extension problem, ${extension} is not allowed as the allowed extension are ${allowedExtensionsMesg}`
            isAllCorrect = false;

            if(singleCheck && clonResults !== undefined && clonResults.length > 0){
                filenames.map(item =>{
                    let fullfilename = item.Name.toLowerCase();
                    let tempFile = fullfilename.substring(0, fullfilename.lastIndexOf("."));

                    clonResults = clonResults.filter(result =>{
                        let resFullfilename = result.filename.toLowerCase();
                        let resTempFile = resFullfilename.substring(0, resFullfilename.lastIndexOf("."))
                        return  resTempFile !== tempFile;
                    });
                });
            }
        }
        let getRow = CheckAllowedExtensionsObj.filter(error =>{ return error.filenameWithExt === filenameWithExt });
        if(getRow.length === 0){
            CheckAllowedExtensionsObj.push({
                filenameWithExt: filenameWithExt,
                filename: filename,
                statusMesg: singleFileMesg
            });
        }
        if(!isFound)
        return true;
    });

    if(isAllCorrect)
        checkedFiles.push(filename);
 }

}

function validateButtons(isErrorFound){
    if(isErrorFound)
            $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#737373').attr("disabled", "disabled");
        else $('span').filter(function () { return $(this).text() == 'Submit'; }).parent().css('color', '#444').removeAttr('disabled');
}

function setErrorMesg(mesg){
    var imgElement = `<img src='${_submitImg}' alt='na'></img>` ;

    if(mesg !== undefined && mesg !== null && mesg !== '' && !mesg.toLowerCase().includes('cancelled'))
      imgElement = `<img src='${_errorImg}' alt='na'></img>`;

   return imgElement

}

const getClosingStatus = async function(){
    let column = 'ApprovedStatus';
    let query = "CurrentStatus eq 'Sent To Area'"
    let wfItem = _web.lists.getByTitle(_WorkflowSteps).items.select(column).filter(query).get();
    if(wfItem !== undefined && wfItem.length > 0){
        let item = wfItem[0];
        clsoingStatus = item[column];
    }
}
//#endregion

//#region SET DATA
const _setData = (Handsontable) => {
  _container = document.getElementById('dt');
  if(_data.length === 0){
        creatErrorlbl('There are no pending items.');
        return;
   }

    let colArray = [{
        title: '',
        data: 'isSelected',
        type: 'checkbox',
        width: '3%',
        checkedTemplate: 'yes',
        uncheckedTemplate: 'no'
    },
    {
        title: 'Type',
        data: 'type',
        type: 'text',
        width: '3%',
        readOnly: true,
	    renderer: 'html'
    },
    {
        title: 'Ext.',
        data: 'extension',
        type: 'text',
        width: '3%',
        readOnly: true
    },
    {
        title: 'Filename',
        data: 'filename',
        type: 'text',
        width: '25%',
        readOnly: true,
        className: 'htLeft'
    },
    {
        title: 'Title',
        data: 'title',
        type: 'text',
        width: '1%',
        readOnly: true,
        className: 'htLeft'
    },
    {
        title: 'Deliverable Type',
        data: 'DelivType',
        type: 'text',
        width: '7%',
        readOnly: true,
        className: 'htCenter'
    },
    {
        title: 'Message',
        data: '_mesg',
        type: 'text',
        width: '43%',
        readOnly: true,
        renderer: 'html',
        className: 'htLeft ErrorMesg'
    },
    {  title: 'Status',
	   data: '_status',
       type: 'text',
       width: '5%',
       readOnly: true,
	   renderer: 'html'
   },
    {
        title: 'Preview Link',
        data: 'previewLink',
        type: 'text',
        width: '10%',
        readOnly: true,
	    renderer: 'html',
        className: 'htCenter'
    },
    {
        title: 'filenameRelativeUrl',
        data: 'fileRelativeUrl',
        type: 'text',
        width: '1%',
        readOnly: true
    }
    ];

	_hot = new Handsontable(_container, {
		data: _data,
        columns: colArray,
        width:'100%',
        height: screenHeight,
        columnSorting: true,
        search: {
            queryMethod: customSearchMatch,
            searchResultClass: 'customClass'
        },
        //fixedColumnsLeft: 4,
        //filters: true,
        //filter_action_bar: true,
        //dropdownMenu: ['filter_by_condition', 'filter_by_value', 'filter_action_bar'],
        colHeaders: true,
        rowHeaders: true,
        // colHeaders: function(col) {
        //     switch (col) {
        //       case 0:
        //         let control = `<input id='masterBoxId' class='htCheckboxRendererInput' type='checkbox' `;
        //         let element = $('#masterBoxId');
        //         if(element.length === 0)
        //           control += 'checked>'
        //         else{
        //             if(isTicked)
        //               control += 'checked>'
        //             else control += '>'
        //         }
        //          return control;
        //       default:
        //         return '';
        //     }
        // },
        afterChange: function(props, operation){
            if(operation === 'edit' || operation === 'Autofill.fill' || operation === 'CopyPaste.paste'){
                if(isFirstFill){
                  clonResults = _masterErrors;
                  isFirstFill = false;
                }
                let propResults = [];


                props.map(prop => {
                    debugger;
                    let index = prop[0]; //rowIndex
                    let column = prop[1];
                    //let isChecked = prop[3];
                    let filenames = [];

                    if(column === 'isSelected'){
                      let oldValue = prop[2];
                      let newValue = prop[3];
                      if(oldValue === newValue)
                        return;

                        let filename = _hot.getDataAtCell(index, _hot.propToCol('filename'));
                        let ext= _hot.getDataAtCell(index, _hot.propToCol('extension'));
                        let filenameWithExt = `${filename}.${ext}`;
                        checkedFiles = checkedFiles.filter(file => file !== filename);
                        let extension, delivType, _mesg;
                        let transaction = 'remove';

                        _data.some((row) =>{
                            let value = row[column];
                            extension = row.extension;
                            let rowFilenameWithExt = `${row.filename}.${extension}`;
                            if(row.filename === filename){
                            if(rowFilenameWithExt === filenameWithExt){
                                if(value === 'yes')
                                  transaction = 'add';
                            }
                                delivType = row.DelivType;
                                _mesg = row._mesg;

                                if(value === 'yes'){
                                    filenames.push(rowFilenameWithExt); //`${filename}.${extension}`);
                                }
                            }
                        });

                        propResults.push({
                            rowIndex: index,
                            filenameWithExt: filenameWithExt,
                            filename: filename,
                            extension: ext,
                            delivType: delivType,
                            mesg: _mesg,
                            transaction: transaction,
                            filenames: filenames
                        });
                    }
                });

                if(propResults.length > 0){
                    let resultAllowedExts;
                     Promise.all(propResults.map(async result =>{
                        let filenameWithExt = result.filenameWithExt;
                        let filename = result.filename;
                        let extension = result.extension;
                        let trans = result.transaction;
                        let filenames = result.filenames;

                        if (checkedFiles.includes(filename))
                          return;

                        let filemasterError = '';
                        if(trans === 'add'){
                            let row = clonResults.filter(result =>{ return result.filenameWithExt === filenameWithExt });
                            if(row.length === 0){
                              let getRow = _masterErrors.filter(error =>{ return error.filenameWithExt === filenameWithExt });
                              if(getRow.length > 0){
                                filemasterError = getRow[0].mesg;
                                clonResults.push({
                                    rowIndex: getRow[0].rowIndex,
                                    filename: getRow[0].filename,
                                    filenameWithExt: getRow[0].filenameWithExt,
                                    mesg: filemasterError
                                });
                              }
                            }
                            else filemasterError = row[0].mesg;
                        }

                        let statusMesg = '';
                        let errorItems = [];
                        if(result.mesg.startsWith('Missing') || result.mesg === ''){
                            let delivType = result.delivType;

                            //let isFound = false;
                            if(resultAllowedExts !== undefined && resultAllowedExts.length > 0){
                                let getRows = resultAllowedExts.filter(error =>{ return error.filename === filename });
                                if(getRows !== undefined && getRows.length > 0){
                                    filenames.map(file =>{
                                        resultAllowedExts.map(resultAllowedExt =>{
                                            if(file === resultAllowedExt.filenameWithExt){
                                                errorItems.push({
                                                    filename: resultAllowedExt.filenameWithExt,
                                                    statusMesg: resultAllowedExt.statusMesg,
                                                    trans: trans
                                                });
                                            }
                                        })
                                    })
                                }
                            }
                            else {
                                resultAllowedExts = await CheckAllowedExtensions(filename, extension, delivType, true, filenames);
                                let mesg = '';
                                if(resultAllowedExts !== undefined){
                                    mesg = resultAllowedExts.statusMesg
                                    // //resultAllowedExts.map(resultAllowedExt =>{
                                    //     errorItems.push({
                                    //         filename: resultAllowedExts.filename,
                                    //         statusMesg: ,
                                    //         trans: trans
                                    //     });
                                    // //})
                                }

                                errorItems.push({
                                    filename: filename,
                                    statusMesg: mesg,
                                    trans: trans
                                });

                            }
                        }
                        if(statusMesg === '' && result.mesg === '' && trans === 'add'){

                            statusMesg = filemasterError;
                            errorItems.push({
                                filename: filenameWithExt,
                                statusMesg: statusMesg,
                                trans: trans
                            });

                        }
                        return errorItems;
                    }))
                    .then(items=>{
                        let concatenatedItems = items.reduce((acc, curr) => acc.concat(curr), []);
                        concatenatedItems= removeDuplicateRows(concatenatedItems);
                            let doUpdate = false;
                            concatenatedItems.map((item)=>{
                                if(item != undefined){
                                    let filename = item.filename; //this is comment either filename or filenameWithExt
                                    let trans = item.trans;
                                    if(trans === 'remove')
                                      clonResults = clonResults.filter(result =>{
                                        if(result.mesg.startsWith('Missing')){
                                            return result.filename !== filename
                                        }
                                        else return result.filenameWithExt !== filename
                                   });
                                    //doUpdate = updateMetaInfo(item.isPairCheck, filename, trans, item.statusMesg);
                                    if(!doUpdate)
                                     doUpdate = updateMetaInfo(filename, trans, item.statusMesg);
                                    else updateMetaInfo(filename, trans, item.statusMesg);
                                }
                            });
                            return doUpdate;

                    })
                    .then((doUpdate)=>{
                        if(doUpdate)
                        _hot.updateSettings({data: _data});
                        missingPairErrorsFiles = [];
                    })
                    .then(()=>{
                        let isErrorFound = false;
                        rows = _data.filter(item =>{return item.isSelected === 'yes'});
                        if(rows.length < 1){
                          isErrorFound = true;
                          setErrorMessage('One file at least must be selected')
                        }
                        else {
                           let ErrorItems = setErrorCounter();

                           if(ErrorItems === 0 && !isErrorFound){
                                if(_isDesign && rlodFilenames !== undefined && distinctDelivFilename.length > 0){
                                    let difference = rlodFilenames.filter(item => !distinctDelivFilename.includes(item));
                                    if(difference.length > 0){
                                        isErrorFound = true;
                                        let tempMesg = `The below filename ${difference.length} are missing while they are listed in RLOD Under Reference ${_submittalRef},
                                                        please upload through ${_itemListname}.<br /><br />`;

                                        difference.map(item =>{tempMesg += `${item}<br/>`})
                                        setErrorMessage(tempMesg);
                                    }
                                }
                           }

                           if(ErrorItems === 0 && !isErrorFound){
                                var errorElement = $('.errors[data-v-386d995a]');
                                if(errorElement.length > 0){
                                    errorElement.css({
                                        'height': '0',
                                        'opacity': '0'
                                    });

                                    if($('#customErrorId'). length > 0)
                                    $('#customErrorId').remove();
                                }
                                $('.alert-errors[data-v-386d995a]').addClass('errors');
                            }
                            else isErrorFound = true;
                        }

                        validateButtons(isErrorFound);
                    });
                }
            }
        },
        //manualColumnResize: true,
        stretchH: 'all',
        licenseKey: htLicenseKey
	 });

     setErrorCounter();

     isSelectedColIndex = _hot.propToCol('isSelected');
     const columnIndex = _hot.propToCol('fileRelativeUrl');
     const titleColumnIndex = _hot.propToCol('title');
     _hot.updateSettings({
      hiddenColumns: {
        columns: [titleColumnIndex, columnIndex], //titleColumnIndex
        indicators: false // Show the hidden columns indicators
      }
     });

    handleSearchInput();

    $(document).on('click', '#masterBoxId', function() {
        CheckBoxes(this);
        $('#masterBoxId').removeAttr('checked');
        $('#masterBoxId').prop('checked', false);
    });

    //disableTickBoxes();

    setTimeout(() => {
         $('div.fd-grid').attr('style', 'padding: 8px !important;');
        // $('div[data-v-105ebe50].row').eq(2).attr('style', 'width: 100%; height: auto !important;');
        // _container.addClass('handsontable htRowHeaders htColumnHeaders');
        _hot.render();
    }, 100);
}

function customSearchMatch(queryStr, value, meta) {
    if(queryStr !== '' && queryStr !== undefined && value !== null && meta.prop !== 'type' && meta.prop !== 'fileRelativeUrl'){
        if(meta.prop === 'previewLink' && value !== 'not supported'){}
        else {
            value = value !== undefined && value !== '' ? value.toLowerCase().toString() : '';
            queryStr = queryStr !== ''? queryStr.toLowerCase().toString() : ''
            if(queryStr.length > 1 && (queryStr === value || value.includes(queryStr)) ){
                 return true;
            }
        }
    }
}

function CheckBoxes(element){
    let value = 'yes';
    isTicked = true;
    if(!element.checked){
        value = 'no';
        isTicked = false;
    }

    _hot.updateSettings({
        data: _data.forEach(obj => {
            if(obj._mesg !== undefined && obj._mesg !== '')
              obj.isSelected = value;
          })
      });
}

function handleSearchInput(){
    const searchFiled = $('#search_field');
    searchFiled.prop('placeholder', 'Search any...');

    if(searchFiled.length > 0){
       searchFiled.css({
           'width': '400px',
           'padding': '4px',
           'border': '1px solid #ccc',
           //'border-radius': '5px',
           'outline': 'none',
           'font-size': '14px',
           'box-shadow': '0px 3px #888888',
           'float': 'left',
           'display': 'inline'
       });

       searchFiled.blur(function() {
           $(this).css('border-color', '#ccc');
       });

       searchFiled[0].addEventListener('input', evt => {
           if(evt.inputType !== undefined || (!evt.inputType && (evt.data === undefined))){
               var search = _hot.getPlugin('search');
               var queryResult = search.query(evt.currentTarget.value);
               _hot.render();
           }
       });
    }
}

function setTickBoxReadOnly(rowIndex, isReadOnly){
    let metaObject = {};
    metaObject['readOnly'] = isReadOnly;
    metaObject['className'] = disableClassName;
    _hot.setCellMetaObject(rowIndex, isSelectedColIndex, metaObject);
}
//#endregion

//#region FEATURES
const renderTransmittal = async function(){
    const searchFiled = $('#search_field');
    var transmittalControl = `<div style='padding-bottom: 22px'>
                                <label for='transmittal' class='fd-field-title col-form-label col-sm-auto'>Transmittal No:</label>
                                <input type="text" id='transmittal' name='transmittal' style='width: 20%'>
                              </div>`;
    searchFiled.before(transmittalControl);
}

const evaluateWorkingHour = async function(){
    let workWeek =  await getParameter('WorkWeek');
    workWeek = workWeek.split(',');

    // Get the current day of the week
    var currentDate = new Date();
    var workday = currentDate.toLocaleString('en-US', { weekday: 'long' });

    // Find all occurrences of the current day in the 'WorkWeek' array
    var results = workWeek.filter(function(day) {
        return day === workday;
    });

    if (results.length !== 0) {
        let ArrayWorkStartTime =  await getParameter('WorkStartTime');
        ArrayWorkStartTime = ArrayWorkStartTime.split(',');

        let ArrayWorkEndTime =  await getParameter('WorkEndTime');
        ArrayWorkEndTime = ArrayWorkEndTime.split(',');

        if (ArrayWorkStartTime.length > 0 && ArrayWorkEndTime.length > 0) {
            // Convert the time components to TimeSpan
            var WorkStartTime = new Date();
            WorkStartTime.setHours(parseInt(ArrayWorkStartTime[0]), parseInt(ArrayWorkStartTime[1]), parseInt(ArrayWorkStartTime[2]));

            var WorkEndTime = new Date();
            WorkEndTime.setHours(parseInt(ArrayWorkEndTime[0]), parseInt(ArrayWorkEndTime[1]), parseInt(ArrayWorkEndTime[2]));

            var now = new Date();
            // Check if the current time is within working hours
            if (now > WorkStartTime && now < WorkEndTime) {
                // Do nothing
            } else return 'Please note that you are not allowed to submit after working hours';

        }
    } else return 'Please note that you are not allowed to submit on weekends.';
}

function previewFile(hrefElement, isRename) {

    event.preventDefault();
    let linkHref = hrefElement.getAttribute('href');
    let width = '800';

    let left = (window.screen.width - 800) / 2;
    if(isRename)
     width = '1050';

    const top = (window.screen.height - 600) / 2;
    previewWindow = window.open(linkHref, 'filePreview', `width=${width},height=600,left=${left},top=${top}`);

    //preloader_btn(false, true);
    checkPreviewInterval = setInterval(closePreview, 50);
}

var closePreview = async function() {
    if (previewWindow && previewWindow.closed) {
        previewWindow = null; // Reset previewWindow before closing to prevent infinite loop
        clearInterval(checkPreviewInterval); // Clear any remaining intervals
        Remove_Pre(true);

        let filenameNewValue = localStorage.getItem('filenameNewValue');
        let extension = filenameNewValue.substring(filenameNewValue.lastIndexOf(".")+1);
        let newFilenameWithExt = filenameNewValue;

        if(filenameNewValue !== undefined){
          let filenameOldValue = localStorage.getItem('filenameOldValue');
          let oldFilenameWithExt = filenameOldValue;

          if(filenameOldValue !== filenameNewValue){
            filenameNewValue = filenameNewValue.substring(0, filenameNewValue.lastIndexOf("."));
            filenameOldValue = filenameOldValue.substring(0, filenameOldValue.lastIndexOf("."));

            let changes = _data.map(async (row) => {
                if (row.filename === filenameOldValue) {
                    let object = await validateFileName(row);//newFilenameWithExt, filenameNewValue, extension);
                    row.filename = filenameNewValue;
                    row._mesg = object.mesg;
                    row._status = object.status;
                    localStorage.removeItem('filenameNewValue');
                    localStorage.removeItem('filenameOldValue');

                    if(object.mesg !== ''){
                     _masterErrors = _masterErrors.map(error => {
                            if(error.filename === oldFilenameWithExt){
                                error.filename = newFilenameWithExt;
                                error.mesg = object.mesg; // Update the mesg property
                            }
                            return error
                        });
                    }
                    else _masterErrors = _masterErrors.filter(error => error.filename !== oldFilenameWithExt);
                    isFirstFill = true;

                }
                return row;
            });
            _data = await Promise.all(changes);

            _hot.updateSettings({
                data: _data
            });
          }
        }
  }
}

const handleFeatures = async function(isTranRequired){
    if(isTranRequired)
      await renderTransmittal();

    if(isWorkHoursEnabled){
        let mesg = await evaluateWorkingHour();
        if(mesg !== undefined && mesg !== null & mesg !== '')
           await creatErrorlbl(mesg);
          return;
    }
}
//#endregion

function setErrorMessage(errMesg){

  function checkForErrors() {
    var errorElement = $('.errors[data-v-386d995a]');
    if (errorElement.length > 0) {
        clearInterval(intervalId);
        handleErrors(errorElement, errMesg);
    }

    $('button').hover(
        function () {
            // Mouseenter event handler
            $(this).css('color', 'white');
        },
        function () {
            // Mouseleave event handler
            $(this).css('color', 'black');
        }
    );
  }

  function handleErrors(element, errorMessage) {
    if (errorMessage !== '') {
        element.css({
            'height': 'auto',
            'opacity': '1'
        });
        errorMessage = '<br/>' + errorMessage;
        var mesgElement = $('#customErrorId');
         if(mesgElement.length === 0)
          $('p.alert-heading').append(`<p id='customErrorId'>${errorMessage}</p>`);
         else mesgElement.html(errorMessage);
        //$('span').filter(function () { return $(this).text() == buttonText; }).parent().css('color', '#737373').attr("disabled", "disabled");
    } else {
        element.css({
            'height': '0',
            'opacity': '0'
        });
        $('#customErrorId').remove();
        $('.alert-errors[data-v-386d995a]').addClass('errors');
        //$('span').filter(function () { return $(this).text() == buttonText; }).parent().css('color', '#444').removeAttr('disabled');
    }
  }

  $('button.close').on('click', () => {
    $('div.alert').css({
      'height': '0',
      'opacity': '0'
    });
  });

  var intervalId = setInterval(checkForErrors, 100);
}

function setErrorCounter(){
    let errorData = _data.filter(item=> item._mesg !== '' && !item._mesg.toLowerCase().includes('cancelled') && item.isSelected !== 'no');
    let lengthOfErrors = errorData.length;
    if(lengthOfErrors> 0){
       setErrorMessage(`${lengthOfErrors} validation errors found`);
    }
    return lengthOfErrors;
}

var disableTickBoxes_notUsed = async function(){
    //let trInput = $('#dt table.htCore tbody tr input[type="checkbox"]');;
    //if(trInput.length > 0){
        _data.map((item, rowIndex) =>{
            if(item._mesg === undefined || item._mesg === '' ){
                // let inputElement = $(trInput[rowIndex]);
                // if(inputElement.length > 0){
                //     $(inputElement).prop('disabled', true);
                // }
                setTickBoxReadOnly(rowIndex, true);
            }
        });
}

function updateMetaInfo(filename, trans, statusMesg){
    let doUpdate = false;
    let tempFileName = filename.includes('.') ? filename.substring(0, filename.lastIndexOf(".")): filename;

    if(_isDesign && trans === 'add'){
        let indexToRemove = distinctDelivFilename.indexOf(filename);
        if (indexToRemove === -1)
         distinctDelivFilename.push(filename);
    }
    else if(_isDesign && trans === 'remove'){
        let indexToRemove = distinctDelivFilename.indexOf(filename);
        if (indexToRemove !== -1)
            distinctDelivFilename.splice(indexToRemove, 1);
    }

    _data.find((row) => {
        //let proceed = isPairCheck ? row.filename === filename : `${row.filename}.${row.extension}` === filename

        let rowFileName = `${row.filename}.${row.extension}`;
        let checkExt = false;
        if(statusMesg.startsWith('Missing') || trans === 'add'){
            filename = tempFileName;
            rowFileName = row.filename;
            checkExt = true;
        }

        let proceed =  rowFileName.toLowerCase() === filename.toLowerCase();
        if(proceed && checkExt){
            proceed = row.isSelected === 'yes'? true : false; //trans === 'add' ? tempExtension.toLowerCase() === row.extension.toLowerCase(): tempExtension.toLowerCase() !== row.extension.toLowerCase();
        }

        if (proceed) {
            if (missingPairErrorsFiles.indexOf(tempFileName) !== -1)
             return false;

            if(statusMesg !== '')
              missingPairErrorsFiles.push(filename);

            if(row._mesg !== statusMesg){
                row._status = setErrorMesg(statusMesg);
                row._mesg = statusMesg;
                doUpdate = true;
            }

            else if(trans === 'remove'){
               if(row._status !== ''){
                 row._status = '';
                 doUpdate = true;
               }
            }
            else{
                fileStatus = setErrorMesg(statusMesg);
                if(row._status !== fileStatus){
                    row._status = fileStatus;
                    doUpdate = true;
                  }
            }
        }
   });
   return doUpdate;
}

function removeDuplicateRows(array) {
    let seen = new Set();
    return array.filter(item => {
        if(item === undefined)
         return false;

        let key = item.filename + '|' + item.statusMesg +  '|' + item.trans; // Combine filename and mesg as key
        if (seen.has(key)) {
            return false;
        } else {
            seen.add(key);
            return true;
        }
    });
}

const checkTaskStatus = async function(isUpdate, wfStatus){

    let mesg = '';
    let query = `Reference eq '${_submittalRef}' and Trade eq '${_darTrade}'`
    let items = await _web.lists.getByTitle(_DesignTasks).items.select("Id, Status, WorkflowStatus, ReviewLink").filter(query).get();
    if (items !== undefined && items.length > 0){
        let item = items[0];
          if(isUpdate){
            let setItem = {};
            setItem["Status"] = "Closed";
            setItem["WorkflowStatus"] = wfStatus;

            let url = `${_webUrl}${_layout.replace('EForms','CDS')}/BulkFilter.aspx?Reference=${_submittalRef}`
            let reviewLink = {
                Description: 'Review',
                Url: url
            };
            setItem["ReviewLink"] = reviewLink;

            _web.lists.getByTitle(_DesignTasks).items.getById(item.Id).update(setItem);
          }
          else{
            let workflowStatus = item['WorkflowStatus'];
            if(workflowStatus.toLowerCase() === 'sent to pm' || workflowStatus.toLowerCase() === 'sent to pm')
            mesg = `Submission ${_submittalRef} is already ${workflowStatus}`;
          }
    }
    return mesg;
}

const getPending_RLOD_FullFilenames = async function(){
    let query = `SubmittalRef eq '${_submittalRef}' and DarTrade eq '${_darTrade}' and Status ne 'Cancelled'`;
    if (!_exact_Filename_Match)
       query += " and (Status eq 'Pending' or Status eq 'Rejected')";

    return await _web.lists.getByTitle(_RLOD).items.select("FullFileName").filter(query).getAll()
    .then((items)=>{
        if(items.length > 0){
           let fileNames = [];
          for(const item of items){
            fileNames.push(item.FullFileName.toUpperCase());
          }
          return fileNames;
        }
     });
}

const handelConstructionSubmit = async function () {
    debugger;
  let mesg = '',  errorFound = false;

  if(enableTransmittal){
    let transmittalText = $('#transmittal').text();
    if(transmittalText === undefined || transmittalText === null || transmittalText === '')
        mesg = 'Transmittal Number is Required Field';
  }

   _data = _data.filter((item)=>{
      if(item.isSelected === 'yes'){
        let delivType = item.DelivType !== undefined ? item.DelivType : '';

        if(delivType !== '' && !docType.includes(delivType))
            docType.push(delivType);
        return item;
      }
   });

   console.log(cdsFormDataArray);

  if(docType.length > 1){
    if(docType[0] === 'DWG' && docType[1] === 'TRM')
      _acronym = 'DWG';
    else mesg = `Mentioned Submittal Types are not allowed Together: ${docType.join(', ')}`;
  }
  else _acronym = docType[0];

  if (_isMultiContractor && multiContractorValues.length == 0)
    mesg = 'Multi Contractor is Missing';


    if(mesg !== ''){
        errorFound = true;
        setErrorMesg(mesg);
    }
    return {
        errorFound: errorFound,
        mesg: mesg
    };
}

//#region HANDLE DESIGN SUBMIT
const handelDesignSubmit = async function(){
    preloader();
    let query = 'StepID eq 1';
    let wfColumns = ['ApprovedStatus', 'NotificationUser', 'ApprovalEmailName', 'ApprovalTradeCC'];
    let wfItem = await _web.lists.getByTitle(_WorkflowSteps).items.select(wfColumns).filter(query).get();

    let workflowStatus, NotificationUser, ApprovalEmailName, ApprovalTradeCC;
    if(wfItem.length > 0){
      let item = wfItem[0];
      workflowStatus = item.ApprovedStatus;
      NotificationUser = item.NotificationUser;
      ApprovalEmailName = item.ApprovalEmailName;
      ApprovalTradeCC = item.ApprovalTradeCC;
    }

    let UpdatedFiles = [];
    let filenames = []
    let rlodList = _web.lists.getByTitle(_RLOD);

    let cloneDdata = _data.reduce((accumulator, currentItem) => {
        if (currentItem.isSelected === 'yes'){
            const isDuplicate = accumulator.some(item => item.filename === currentItem.filename);

            // If the filename is not a duplicate, add it to accumulator array
            if(!isDuplicate) {
                accumulator.push(currentItem);
            }
        }

        return accumulator;
    }, []);

    let totalFiles = cloneDdata.length;
    const batchSize = totalFiles < rowSize ? totalFiles: rowSize;

    const dataChunks = [];
    for (let i = 0; i < cloneDdata.length; i += batchSize) {
        dataChunks.push(cloneDdata.slice(i, i + batchSize));
    }

    // Using for...of loop to iterate over dataChunks
    for (let [chunkIndex, chunk] of dataChunks.entries()) {
        let metaInfo = [];
        let filenames = []; // Initialize filenames array for the current chunk
        let rev;
        // Process each item in the current chunk
        for (let [rowIndex, item] of chunk.entries()){
            let delivType = item.DelivType;
            let _schema = _fncSchemas[delivType];
            let containsRev = _schema.containsRev;
            let delimeter = _schema.delimeter;
            let revStartWith = _schema.revStartWith;

            let filename = item.filename;
            let filenameWithoutRev = containsRev ? item.filename.substring(0, item.filename.lastIndexOf(delimeter)) : item.filename;

            if(checkAgainstFNC){
                rev = containsRev ? item.filename.substring(item.filename.lastIndexOf(delimeter) + 1) : '';
                if (revStartWith !== undefined)
                    rev = rev.replace(revStartWith, '');
            }

            let query = `FileName eq '${filenameWithoutRev}'`;
            if (checkAgainstFNC && !containsRev) {
                let items  = await _web.lists.getByTitle(_SLOD).items.select('ID').filter(query).getAll();
                rev = items.length; // FOR MOD Type
                rev = String(rev).padStart(_bimRevFormat.length, '0');

            }

            if (UpdatedFiles.indexOf(filename) === -1) {
                UpdatedFiles.push(filename);
            }

            if (filenames.indexOf(filename) === -1) {
                filenames.push(filename);
            }

            metaInfo.push({
                filename: filename,
                rev: rev
            });
        }

        await updateRLOD(rlodList, filenames, workflowStatus, metaInfo, totalFiles, ApprovalEmailName, ApprovalTradeCC, NotificationUser);
    }
}

const updateRLOD = function(rlodList, filenames, workflowStatus, metaInfo, totalFiles, ApprovalEmailName, ApprovalTradeCC, NotificationUser){
    var batch = pnp.sp.createBatch();
    let query = '';
      filenames.map((filename)=>{
        query += `FullFileName eq '${filename}' or `;
      });

      if(query !== '')
       query = query.substring(0, query.lastIndexOf(' or'));
    let cols = ['Id', 'FileName','SentToPMDate','DWGLink','PDFLink','Status','SubmittalRef','Revision','FullFileName','CDSNumber', 'CDSLink','SubmittedDate','State','DWFLink'];
    //let query = `FullFileName eq '${filename}'`;
       rlodList.items.select(cols).filter(query).getAll()
               .then((items)=>{

                    if(items.length > 0){
                        items.forEach(item => {
                            let setItem = {};
                            let files = _data.filter(file =>{
                                return file.filename === item.FullFileName;
                            })

                            files.map(file=>{
                                let ext = file.extension;
                                let url = file.fileRelativeUrl;

                                let fileLink = {
                                    Description: ext,
                                    Url: url
                                };

                                if(ext === 'pdf' || ext === 'rvt')
                                setItem["PDFLink"] = fileLink;
                                else if(ext === 'dwf' || ext === 'dwfx')
                                setItem["DWFLink"] = fileLink;
                                else setItem["DWGLink"] = fileLink;
                            });

                            setItem["Status"] = workflowStatus;

                            if(checkAgainstFNC){
                                let metaItem = metaInfo.filter(metaItem=> { return metaItem.filename === item.FullFileName });
                                if(metaItem.length > 0)
                                setItem["Revision"] = metaItem[0].rev;
                            }

                            setItem["SubmittalRef"] = _submittalRef;
                            setItem["FullFileName"] = item.FullFileName;
                            setItem["CDSNumber"] = "";
                            setItem["CDSLink"] = null;
                            setItem["SubmittedDate"] = null;
                            setItem["State"] = "Open";
                            setItem["SentToPMDate"] = new Date();

                            //rlodList.items.inBatch(batch).add(item);
                            rlodList.items.getById(item.Id).inBatch(batch).update(setItem);
                        });
                    }
               })
               .then(()=>{
                    totalRowUpdates += batch._deps.length;
                    detailedLoader(totalFiles, totalRowUpdates, 'Updating');
                    batch.execute();

                    if(totalFiles === totalRowUpdates){
                       Promise.all([
                            updateDeliverableFolderItemStatus_SetFolderPermission(workflowStatus),
                            checkTaskStatus(true, workflowStatus)
                        ])
                        .then(()=>{
                            // Construct spQuery
                            let spQuery = "<Where><And><And>";
                            if (!_exact_Filename_Match)
                                spQuery += "<And>";

                            spQuery += "<Eq><FieldRef Name='SubmittalRef'/><Value Type='Text'>" + _submittalRef + "</Value></Eq>" +
                                "<Eq><FieldRef Name='DarTrade'/><Value Type='Text'>" + _darTrade + "</Value></Eq>" +
                                "</And>" +
                                "<Neq><FieldRef Name='Status'/><Value Type='Text'>Cancelled</Value></Neq></And>";

                            if (!_exact_Filename_Match) {
                                spQuery += "<Or>" +
                                    // "<Eq><FieldRef Name='Status' /><Value Type='Text'>Rejected</Value></Eq>" +
                                    // "<Eq><FieldRef Name='Status' /><Value Type='Text'>Pending</Value></Eq>" +
                                    "<Eq><FieldRef Name='Status' /><Value Type='Text'>" + workflowStatus + "</Value></Eq>" +
                                    "<Eq><FieldRef Name='Status' /><Value Type='Text'>Sent To PMC123</Value></Eq>" +
                                    "</Or>" +
                                "</And>";
                            }
                            spQuery += "</Where>";

                            const sendEmailWithTimeout = () => {
                                _sendEmail(_module, ApprovalEmailName, spQuery, ApprovalTradeCC, NotificationUser, _rootFolderUrl);
                                preloader('remove');
                                fd.save();
                            };
                            setTimeout(sendEmailWithTimeout, 1000);
                         })
                    }
               })
}

const updateDeliverableFolderItemStatus_SetFolderPermission = async function(workflowStatus){
    let relativeUrl = `${_Deliverables}/${_submittalRef}/${_darTrade}`;
    const folder = await _web.getFolderByServerRelativeUrl(relativeUrl).getItem();
    const folderId = folder.Id;

    _web.lists.getByTitle(_Deliverables).items.getById(folderId).update({
         "WorkflowStatus": workflowStatus
     })
    .then(async () => {
        console.log('Item updated successfully.');

        // Break role inheritance for the folder
        await folder.breakRoleInheritance(false);
        const { Id: contRoleDefId } = await _web.roleDefinitions.getByName("Contribute").get();
        const { Id: readRoleDefId } = await _web.roleDefinitions.getByName("Read").get();
        let roleAssignments = await folder.roleAssignments();

        let promises = roleAssignments.map(async (roleAssignment) =>{
           return {
             principalId: roleAssignment.PrincipalId,
             group: await _web.siteGroups.getById(roleAssignment.PrincipalId).get().catch(()=>{ return undefined })
           }
        })

        Promise.all(promises)
        .then(results=>{
            return results.filter(result=>{
                if(result.group !== undefined && result.group.Title === _darTrade)
                   return result;
            })
        })
        .then(async (groups)=>{
            Promise.all(groups.map(async group=>{
                 folder.roleAssignments.remove(group.principalId, contRoleDefId);
                 folder.roleAssignments.add(group.principalId, readRoleDefId);
            }))
            .then(()=>{
                console.log(`${_darTrade} permission is set successfully.`);
            });
        })
    })
    .catch(error => {
        console.error('Error updating item:', error);
    });
}
//#endregion


var  detailedLoader = async function(total, index, trans){
    var targetControl = $('#ms-notdlgautosize').addClass('remove-position-preloadr');

    if(trans === undefined)
      trans = 'validating';

    let mesg = `${trans} ${index} out of ${total}. Please dont close this page.`;
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