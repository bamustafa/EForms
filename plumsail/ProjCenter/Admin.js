var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list;
var _isSiteAdmin = false, _isMain = true, _isLead = false, _isPart = false, _isNew = false, _isEdit = false;
var fieldsSchema = [];

var onRender = async function (relativeLayoutPath, moduleName, formType){
    
    _layout = relativeLayoutPath;
    await loadScripts();
    await extractValues(moduleName, formType);

    if(_module === 'SCAT')
       await onSCATRender();

    await setCustomButtons();
}

//#region GENERAL
var loadScripts = async function(){

    const libraryUrls = [
      _layout + '/controls/preloader/jquery.dim-background.min.js',
      _layout + "/plumsail/js/buttons.js",
      _layout + '/plumsail/js/customMessages.js',
      _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
      _layout + '/plumsail/js/preloader.js',
      _layout + '/plumsail/js/commonUtils.js',
      _layout + '/plumsail/js/utilities.js',
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

    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;

    if(_formType === 'New'){
        //clearStoragedFields(fd.spForm.fields);
        _isNew = true;
    }
    else if(_formType === 'Edit'){
        _isEdit = true;
    }

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;
}

var setCustomButtons = async function () {
  
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
              await createMatrixFields();
              fd.save();
           }
       }
    });
}
//#endregion


//#region SCAT RENDER MODULE
var onSCATRender = async function () {

    // let controlField = fd.control('dt');
    // controlField.$on('edit', function(editData) {
    //    let getControl = $('input[name=FieldType]');
    //    if(getControl.length > 0){
    //      getControl.on('change', function() {
    //         let fieldType = $(this).val()
    //         setColumnValidation(controlField, 'field is required')
    //      });
    //    }
    //  })
    //  isValidated = false;
}

var createMatrixFields = async function(){

    let matrixList = 'Matrix Fields';
    let controlField = fd.control('dt');
    var rows = controlField.widget.dataItems();

    for(const row of rows){
        let colInfo = {
            fieldName: row['Fieldname'],
            fieldDisplay: row['FieldDisplay'],
            fieldType: row['FieldType'],
            lookupListName: row['LookupListName'],
            isRequired: false
            
        };
        await getSchema(matrixList, colInfo);
    }
    await createFields('Create');
    await createFields('Update');
    await createFields('View');
}

var getSchema = async function(list, colInfo){

 try{
        let fieldName = colInfo.fieldName;
        let fieldDisplay = colInfo.fieldDisplay;
        let type = colInfo.fieldType.toLowerCase();
        let lookupListName = colInfo.lookupListName;
        let isRequired = colInfo.isRequired.toString().toUpperCase();

        let fieldXml  = '';

        try { 
            await _web.lists.getByTitle(list).fields.getByInternalNameOrTitle(fieldName).get(); 
            fieldsSchema.push({
                fieldName: fieldName,
                fieldDisplay: fieldDisplay,
                fieldType: type,
                fieldSchema: fieldXml
            });
            return;
        }
        catch(err){
            // continue to create
        } 

        if(type === 'text')
            fieldXml  = `<Field Type='Text' Name='${fieldName}'  DisplayName='${fieldName}' Title='${fieldDisplay}' Required='${isRequired}' />`

        else if(type === 'int')
            fieldXml  = `<Field Type='Number' Name='${fieldName}' DisplayName='${fieldName}' Required='${isRequired}' Decimals='0' />`

         else if(type === 'decimal')
            fieldXml  = `<Field Type='Number' Name='${fieldName}' DisplayName='${fieldName}' Required='${isRequired}' Decimals='5' />`

        else if(type === 'bool')
            //fieldXml  = `<Field Type='Boolean' Name='${fieldName}' DisplayName='${fieldName}' Required='${isRequired}' />`
           fieldXml  = `<Field Type='Boolean' Title='${fieldDisplay}' DisplayName='${fieldName}' Required='${isRequired}'><Default>FALSE</Default></Field>`
        else if(type === 'date')
            fieldXml  = `<Field Type='DateTime' Name='${fieldName}' Format='DateOnly' DisplayName='${fieldName}' Required='${isRequired}' />`
        else if(type === 'lookup' || type === 'multilookup'){
            const list = await _web.lists.getByTitle(lookupListName).select("Id").get();
            if(type === 'lookup')
              fieldXml  = `<Field Type='Lookup' Name='${fieldName}' DisplayName='${fieldName}' Required='${isRequired}' List='${list.Id}' ShowField='Title' />`
            else fieldXml  = `<Field Type='LookupMulti' Name='${fieldName}' DisplayName='${fieldName}' Required='${isRequired}' List='${list.Id}' ShowField='Title' Mult='TRUE' />`
        }
        fieldsSchema.push({
             fieldName: fieldName,
             fieldDisplay: fieldDisplay,
             fieldType: type,
             fieldSchema: fieldXml
        });
    }
    catch (e) {
       console.log(e);
    }
}

var createFields = async function(operation){
    try {
        const list = _web.lists.getByTitle('Matrix Fields');
        const batch = pnp.sp.createBatch();

        let defaultView, fieldsInDefaultView;
        
        if(operation === 'View'){
            defaultView = await list.defaultView.get();
            fieldsInDefaultView = defaultView.HtmlSchemaXml;
        }
        
        for (const schema of fieldsSchema) {
            let {fieldName, fieldDisplay, fieldType, fieldSchema} = schema;


            if(operation === 'Create' && fieldSchema !== ''){
                list.fields.inBatch(batch).createFieldAsXml(fieldSchema).then(() => {
                    console.log(`Field created with schema: ${fieldSchema}`);
                })
                .catch(error => {
                    console.error(`Error creating field with schema: ${fieldSchema}`, error);
                });
            }
            else if(operation === 'Update'){
                if(fieldType === 'bool') continue; // not supported from pnp

                let isFound = false;
                try{
                    const field = await list.fields.getByInternalNameOrTitle(fieldDisplay).get();
                    isFound = true;
                } catch(err){}

                if(!isFound){
                    const intervalId = setInterval(() => {
                        list.fields.getByInternalNameOrTitle(fieldName).update({ Title: fieldDisplay })
                            .then(() => {
                                console.log(`Field ${fieldName} updated with display: ${fieldDisplay}`);
                                clearInterval(intervalId); // Clear interval after successful update
                            })
                            .catch(error => {
                                if (error !== 'Save Conflict') {
                                    clearInterval(intervalId); // Clear interval if not a save conflict
                                }
                                console.error(`Error updating field: ${fieldName}`, error);
                            });
                    }, 200);
                }
            }

            else if(operation === 'View'){
                let IndexValue = fieldsInDefaultView.indexOf(fieldName);
                if(IndexValue === -1){
                    list.defaultView.fields.inBatch(batch).add(fieldName).then(() => {
                        console.log(`Field ${fieldName} set in view`);
                    })
                    .catch(error => {
                        console.error(`Error set field: ${fieldName} in view`, error);
                    });
                }
            }
        }

        if(operation !== 'Update')
          await batch.execute();

        if(operation === 'Create')
          console.log("All fields have been created successfully.");

        else if(operation === 'Update')
            console.log("All fields have been updated successfully.");

        else if(operation === 'View')
            console.log("All fields have been added to view successfully.");

    } catch (error) {
        console.error(`Error ${operation} fields in batch`, error);
    }
}
//#endregion


