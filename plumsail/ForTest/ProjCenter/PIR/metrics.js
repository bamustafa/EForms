let isTblCreated = false;
let single = [], multi = [], classMetArray = [];

let metricsformFields = [];

var onMetricsRender = async function () {
  
  const startTime = performance.now();
    metricsformFields = {
      Description: fd.field("Description"),
      Category: fd.field("Category"),
      SubCategory: fd.field("SubCategory"),
      IsConfidential: fd.field("IsConfidential"),
      showAll: fd.field("showAll")
    }
  
    let category = metricsformFields.Category;
    let subCategory = metricsformFields.SubCategory;
    let isConfidential = metricsformFields.IsConfidential.value;
    let showAll = metricsformFields.showAll;
    metricsformFields.Description.placeholder = 'Description and Components/Scope of Work';
  
    if(isConfidential)
      setConfidentialImage('Set Project Confidential');
    
    
    category.$on("change", async function (value){
          let catId  = null;
          if(value !== undefined && value !== null){
              subCategory.value = null; 
              catId = value.LookupId;

              let isDisabled = false;
              if(value.LookupValue === 'Buildings')
                isDisabled = true;
              await enable_Disable_Tabs(tabs, isDisabled);
          }
          
          subCategory.filter = "Category/Id eq " + catId;
          subCategory.orderBy = { field: "Title", desc: false };
          subCategory.refresh().then(()=>{
            let dataLength = metricsformFields.SubCategory.widget.dataSource.data().length;
            let isRequired = dataLength > 0 ? true : false;
            subCategory.required = isRequired;
          });
    });       
  
    subCategory.$on("change", async function (values){
              await setRenderLogic(values);
    });
  
    showAll.$on("change", async function (value){
      let catId  = null;
  
      let query = `Category/Id ne 1234567`;
      if(value === false){
        catId = category.value.LookupId;
        query = `Category/Id eq ${catId}`;
      }
          
      subCategory.filter = query;
      subCategory.orderBy = { field: "Title", desc: false };
      subCategory.refresh();
    }); 
  
    if(_isEdit){
        let categoryId = category.value !== undefined && category.value !== null ? category.value.LookupId : null;
        if(categoryId !== null){
          subCategory.filter = "Category/Id eq " + categoryId;
          subCategory.orderBy = { field: "Title", desc: false };
          subCategory.refresh();
        }
        await setRenderLogic(subCategory.value);
    }

    const endTime = performance.now();
    const elapsedTime = endTime - startTime;
    console.log(`onMetricsRender: ${elapsedTime} milliseconds`);
}
  
  const setRenderLogic = async function(subCategoryArray){
    let tableId = 'metricId';
      if(subCategoryArray !== null && subCategoryArray.length > 0){

        await isDRD(subCategoryArray);

        $(`#${tableId}`).css('display', '');
        let fieldSchemas = [];
  
        let prmoises = subCategoryArray.map(async (value)=>{
            let subCatId = value.LookupId;
            let fieldSchema = value.FieldSchema !== undefined && value.FieldSchema !== null && value.FieldSchema !== '' ? value.FieldSchema : '';
          
            if(fieldSchema === ''){
                let query = `Id eq '${subCatId}'`;
                fieldSchema =  await _web.lists.getByTitle(SubCategory).items.select("Id,FieldSchema").filter(query).get()
                    .then(async function (items) {
                        if(items.length > 0)
                        {
                          const item = items[0]
                          return item.FieldSchema !== undefined && item.FieldSchema !== null ? item.FieldSchema : '';
                        }
                    })
            }
            if(fieldSchema !== '')
              fieldSchemas.push(JSON.parse(fieldSchema))
            
            classMetArray.push({
              sucCatId: subCatId,
              sucCatTitle: value.LookupValue,
              isDRD: value.IsDRD
            })
        });
  
        await Promise.all(prmoises)
        .then(async ()=>{
  
          let flatFieldSchemas = [];
          fieldSchemas.forEach(item => {
            for (let key in item) {
              let exists = flatFieldSchemas.some(schema => schema.Fieldname === item[key].Fieldname);
              if (!exists) {
                  flatFieldSchemas.push(item[key]);
              }
            }
          });
                
          if(flatFieldSchemas.length > 0)
            showHideControls(flatFieldSchemas);
          else $('#metricId').remove();
  
          if (flatFieldSchemas.length > 0){
              await renderSubCategoryControls(tableId, flatFieldSchemas);
              await fetchData();
              showDropDownListMenu();
          }
        })
      }
      else{
          $(`#${tableId}`).css('display', 'none');
          $('#metricId').remove();
      }
  }
  
  function showHideControls(flatFieldSchemas){
    let controls = $('#metricId tr td input, #metricId tr td select');
    if(controls.length > 0){
      controls.each(function(index, element){
  
        if(flatFieldSchemas.length > 0){
          let isFound = false;
          let elementId = element.id;
          flatFieldSchemas.map(schema => {
  
              let controlId = schema.Fieldname;
              if(elementId === controlId){
                isFound = true;
                return;
              }
          })
          if(!isFound){
            let trElemenet = $(`#${elementId}`);
            if(trElemenet.length > 0)
              trElemenet.parent().parent().remove();
          }
          else {
            // $(`#${elementId}`).css('visibility', 'visible');
            // $(`label[for="${elementId}"]`).css('visibility', 'visible');
          }
        }
  
      })
    }
  }
  
  const renderSubCategoryControls = async function(tableId, flatFieldSchemas){
      let tbl = '', isTblFound = true;
      if($(`#${tableId}`).length === 0){
        isTblFound = false;
        tbl = "<table id='" + tableId + "' width='100%' class='modern-table'><tr><td colspan='2' class = 'lbldummyStyle'>Metrics</td></tr>";
      }
  
        for (let i = 0; i < flatFieldSchemas.length; i++) {
          let fld = flatFieldSchemas[i];
          let controlsRow = await renderFeaturesControls(fld);
          tbl += controlsRow;
        }
  
        if(!isTblFound){
          tbl += "</table>";
          $('#subform').after(tbl);
        }
        else $(`#${tableId}`).append(tbl);
      
      
      multi.map(ctl =>{
          let id = '#'  + ctl;
          
          let mySelect = new vanillaSelectBox(id, {
              maxWidth: '100%',
              maxHeight: 300,
              search: true
          });
  
          $('ul.multi').attr('style', 'min-width: 358px !important');
          $('div.vsb-main button').css('line-height', '17px');
      })
  
      // single.map(ctl =>{
      //     var selectElement = document.getElementById('FLDLOOKUP');
      //     selectElement.selectedIndex  = -1; // This will remove all options;
      // })
  }
  
  var renderFeaturesControls = async function(schema){
      
      let tblRows = '';
      let controlId = schema.Fieldname;
      if(controlId === '' || $(`#${controlId}`).length > 0) 
        return '';
  
      let fieldType = schema.FieldType;
      let fieldName = schema.FieldDisplay;
      let isRequired = schema.IsRequired;
      let LookupListName = schema.LookupListName;
  
      if(fieldType === 'bool')
        isRequired = false;
  
      let options = []
        if(LookupListName !== undefined && LookupListName !== null && LookupListName !== '')
          options = await getListOptions(LookupListName);
  
      tblRows += "<tr>"
      tblRows += `<th width = '20%'><label for="${controlId}">${fieldName}`;
      
      if(isRequired)
          tblRows += `<span style="color: red"> *</span></label></th>`
      else tblRows += "</label></th>";
  
      tblRows += "<td>";
      tblRows += await getControl(fieldType, controlId, fieldName, isRequired, options);
      tblRows += "</td></tr>";
      return tblRows;
  }
  
  const getControl = async function(colType, ctlId, displayText, isRequired, options){
      let control = '';
      switch (colType) {
          case 'bool':
              control = `<input Id="${ctlId}" type="checkbox" name="${ctlId}" style="vertical-align: middle; width: 20px; height: 20px;" displayName='${displayText}'`;
              return renderControl(control, false);
          case 'date':
              control = `<input Id="${ctlId}" type="date" name="${ctlId}" style="padding: 8px;" class="ctlStyle border" displayName='${displayText}'`;
              return renderControl(control, isRequired);
          case 'int':
          case 'decimal':
              let step = colType === 'int' ? '1' : '0.0';
              control = `<input Id="${ctlId}" type="number" step = ${step} name="${ctlId}" class="ctlStyle border" displayName='${displayText}' onchange='checkNumberVal(this)'`;
              return renderControl(control, isRequired);
          case 'lookup':
              let isFound = single.filter(item=> {
                 return item === ctlId
              })
              if(isFound.length === 0)
                single.push(ctlId);
              control = `<select Id="${ctlId}" name="${ctlId}" class="ctlStyle border" placeholder='select an option' displayName='${displayText}'`;
              control =  renderControl(control, isRequired);
              return `${control}
                        ${options.map(option => `<option value="${option.value}">${option.label}</option>`).join('')}
                      </select>`;
          case 'multilookup':
              let isMultiFound = multi.filter(item=> {
                return item === ctlId
              })
              if(isMultiFound.length === 0)
                multi.push(ctlId);
  
              control = `<select Id="${ctlId}" name="${ctlId}" multiple class="ctlStyle" displayName='${displayText}'`;
              control =  renderControl(control, isRequired);
  
              return `${control}
                        ${options.map(option => `<option value="${option.value}">${option.label}</option>`).join('')}
                      </select>`;
          case 'text':
              control = `<input Id="${ctlId}" type="text" name="${ctlId}" maxlength="255" class="ctlStyle border" displayName='${displayText}'`;
              return renderControl(control, isRequired);
          default:
              return '';
      }
  }
  
  function checkNumberVal(control){
    let val = control.value !== undefined && control.value !== null ? control.value.trim() : null;
    if(val !== null){
      val = val.replace(/[^\d.-]/g, '');
      let isInt = control.step === '1' ? true : false;
      let parts = val.split('.');
  
      if(isInt)
          val = parts[0];
      //else val = parts[0] + '.' + parts.slice(1).join('');
      control.value = val
    }
  }
  
  function renderControl(ctlr, isRequired){
      let control = ctlr;
      if(isRequired) 
          control += ` role = 'required'>`;
      else control += '>';
      return control;
  }
  
  const validateFields = async function(){
  
      let controls = $('#metricId tr td input[role="required"], #metricId tr td select[role="required"]');
      controls = controls.filter(function() {
          return $(this).css('display') !== 'none';
      });
  
      let mesg = '';
      controls.each(function(){
            let id = this.id;
            let val;
            if(this.type === 'select-multiple')
              val = $('.vsb-main button span.title').text();
            else val = $(`#${id}`).val();
    
            if(val === undefined || val === null || val === '' || val.length === 0){
                let label = $(`#${id}`).attr('displayname');
                mesg += `${label} is required field <br/>`;
                setRequiredFieldErrorStyle(this, false);

                this.addEventListener('change', function() {
                  let elementTag = $(this).val().trim();
                  if(elementTag === '')
                    setRequiredFieldErrorStyle(this, false);
                  else setRequiredFieldErrorStyle(this, true);
                });
            }
      });
  
      setPSErrorMesg(mesg);
      return mesg;
  }
  
  const getFieldsData = async function(listname){
  
      let tables = $('table.modern-table').first();
      tables = tables.filter(function() {
          return $(this).css('display') !== 'none';
      });
  
      var resultArray = {};
      let lookupIds = [];

      if(tables.length > 0)
        resultArray['MasterIDId'] = parseInt(_itemId)
      
      tables.each(function(){
          
          resultArray['CategoryId'] = parseInt(metricsformFields.Category.value.LookupId)
  
          let subCategories = metricsformFields.SubCategory.value;
          for(const subCat of subCategories){
            lookupIds.push(parseInt(subCat.LookupId))
          }
          
          let controls = $(this).find('input, select');
          $(controls).each(function(){
              getControlValue(this, resultArray);
          })
      });
  
      if(lookupIds.length > 0)
        resultArray['SubCategoryId'] = {results: lookupIds};
  
      resultArray['Title'] = projectNo;
  
      let listFields = await getListFields(listname);
      for (const field of listFields){
        let fieldname = field.name;
  
        if (!resultArray.hasOwnProperty(fieldname)){
          let type= field.type;
          let value = null;
  
          if(type === 'Lookup' || type === 'LookupMulti')
            fieldname = `${fieldname}Id`
  
          if(type === 'text') 
            value = '';
          else if(type === 'Boolean')
            value = false;
          else if(type === 'LookupMulti')
            value = { results: [] };
          
          if(fieldname === 'MasterID')
          
          resultArray[fieldname] = value;
        }
      }
  
      return resultArray;
  }
  
  const getListFields = async function(_listname){
    let fieldInfo = [];
    await  _web.lists
           .getByTitle(_listname)
           .fields
           .filter("ReadOnlyField eq false and Hidden eq false")
           .get()
           .then(function(result) {
              for (var i = 0; i < result.length; i++) {
                let field = result[i].InternalName;
                if(field !== 'ContentType' && field !== 'Attachments'){
                  if(_listname === MatrixFields){
                     if(field !== 'Category' && field !== 'SubCategory')
                       fieldInfo.push({
                            name: field,
                            type:  result[i].TypeAsString
                        });
                  }
                  else fieldInfo.push(field);
                }
              }
      }).catch(function(err) {
          alert(err);
      });
      return fieldInfo;
  }
  
  const getControlValue = async function(ctlr, resultArray){
  
      let colType = ctlr.type;
      let id = ctlr.id;
      if(id.includes('search_')) 
          return;
      let val = $(`#${id}`).val();
  
      switch (colType) {
          case 'checkbox':
              resultArray[id] = ctlr.checked ? true : false;
              break;
          case 'date':
          case 'number':
          case 'decimal':
          case 'text':
                 if(val !== '')
                  resultArray[id] = val;
                break;
          case 'select-one':
                  let option = ctlr.selectedOptions[0];
                  resultArray[`${id}Id`] = parseInt(option.value); //{ LookupId: option.value, LookupValue: option.text };
                  break;
          case 'select-multiple':
                  let lookupIds = [];
                  $('ul.multi li.active').each(function(index, element){
                    let Id = element.value;
                    lookupIds.push(parseInt(Id));
                  })
                  resultArray[`${id}Id`] = {results: lookupIds};
                  break;
          default:
              return '';
      }
  }
  
  const fetchData = async function(){
  
      let tables = $('table.modern-table').first();
      tables = tables.filter(function() {
          return $(this).css('display') !== 'none';
      });
  
      let classMetArray = []
      fd.field("SubCategory").value.map((item)=>{
          classMetArray.push({
              sucCatId: item.LookupId,
              sucCatTitle: item.LookupValue,
              isDRD: item.IsDRD
          })
      })
  
      let fields = [], fieldsTypes = [], spFields = [], spExpand = [];
      tables.each(function(){
          let controls = $(this).find('input, select');
          $(controls).each(function(){
              if(!this.id.includes('search_')){
  
                  let colName = this.id;
                  let colType = this.type;
  
                  fields.push(colName)
                  fieldsTypes.push(colType);
  
                  if(colType === 'select-one'){
                      spExpand.push(colName);
                      colName = `${colName}/Id`
                  }
                  else if(colType === 'select-multiple'){
                      spExpand.push(colName);
                      colName = `${colName}/Title`
                  }
                  spFields.push(colName);
              }
          })
      });
  
      if(fields.length > 0){
          const _cols = spFields.join(',');
          let query = `Title eq '${projectNo}'`
          let item = await _web.lists.getByTitle(MatrixFields).items.select(_cols).expand(spExpand).filter(query).get()
          .then(async items=>{
              if(items.length > 0){
              return items[0];
              }
          })
  
          if(item !== undefined && item !== null){
              for(let i = 0; i < fields.length; i++){
                  let fieldType = fieldsTypes[i];
                  let fieldname = fields[i];
                  let value = item[fieldname];
  
                  if(fieldType === 'date') 
                      value = value.substr(0, 10);
                  else if(fieldType === 'select-one') 
                      value = value !== undefined ? value.Id : [];
                    
                  await fetchControlData(fieldType, fieldname, value);
              }
          }
      }
  }
  
  const fetchControlData = async function(fieldType, fieldname, value){
      let fieldId = `#${fieldname}`;
      switch (fieldType) {
          case 'checkbox':
              $(fieldId).attr("checked", value);
              break;
          case 'date':
          case 'number':
          case 'decimal':
          case 'text':
          case 'select-one':
                  $(fieldId).val(value);
                break;
          case 'select-multiple':
              if(value.length > 0){
                  
                  let items = [];
                  value.map(val=>{
                      items.push(val.Title);
                  })
                  
                  $(`#btn-group-${fieldname} button`).attr('style', 'max-width: 800px !important');
                  $(`#btn-group-${fieldname} button span.title`).text(items);
                  $('.vsb-menu ul li').each(function(index, li){
                      let $li = $(li); // Use jQuery to wrap the li element
                      let isMatched = items.find(item => item.trim().toLowerCase() === $li.attr('data-text').trim().toLowerCase());
                      if(isMatched) {
                        $li.addClass('active');
                      }
                    });
              }
              break;
          default:
              return '';
      }
  }
  
  function showDropDownListMenu(){
    let ctlr = $('div.vsb-main button');
    let menu = $('div.vsb-menu');
  
    ctlr.on('click', function(){
      menu.css('position', 'relative');
    });
  
    menu.on('focusout', function() {
      closeMenu();
    });
  
    function closeMenu() {
      menu.css('position', 'absolute');
    }
  
    $(document).on('click', function(event) {
      if (!$(event.target).closest('div.vsb-menu, div.vsb-main button').length) {
          closeMenu();
      }
    });
  }
  
  function setValues(fieldId, values) {
      const selectElement = document.getElementById(fieldId);
      
      // Clear all current selections
      for (const option of selectElement.options) {
        option.selected = false;
      }
      
      // Set new selections
      values.forEach(value => {
        for (const option of selectElement.options) {
          if (option.value == value) {
            option.selected = true;
          }
        }
      });
  }
  
  function setConfidentialImage(text){
    let element = $('span').filter(function(){ return $(this).text() == text; });
    if(element.length > 0){
      let targetElement = element.parent().parent().parent();
      let appendElement = targetElement.parent();
  
      targetElement.remove();
  
      let imageUrl =  `${_webUrl}${_layout}/Images/Confidential.png`;
      $("<img id='loader' src='" + imageUrl + "' />")
      .css({
        "width": "150px",
        "height": "60px"
      }).insertAfter(appendElement);
    }
  }
  