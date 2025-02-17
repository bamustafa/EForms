var _hot, _container, _colArray = [], _targetList;
let dtTbl = '#dt', tableId = 'tblItemsId', tblDatesId = 'tblDatesId'
let activities, scopes, disciplines, multi = [];

const months = ['Jan', 'Feb', 'March', 'April', 'May', 'June', 'July', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
const statAcros = ['Audit Planned', 'Audit Under taken', 'Report Issued', 'Major items closed', 'Complete']

let _isSenior = false, _isManager = false;
var onAUSRender = async function() {

    debugger;
    let ctlr = $('[dir="ltr"] .leftMargin .nav_b5013a39 + .container_b5013a39')
    ctlr.css('left', '0px');
    $('div.SPCanvas, .commandBarWrapper').css('padding-left', '20px');

     fields = {
            AuditType: {
                field: fd.field("AuditType"),
                before: 'AUT', 
                after: ''
            },

            Reference: fd.field("Reference"),

            Year: {
                field: fd.field("Year"),
                before: 'YY', 
                after: ''
            },

            Office: {
                field: fd.field("Office"),
                before: 'XX', 
                after: ''
            },

            Discipline: {
                field: fd.field("Discipline"),
                before: 'DD', 
                after: ''
            },

            Project: {
                field: fd.field("Project"),
                before: 'PRJNUM', 
                after: ''
            },

            ProjTitle: fd.field("ProjTitle"),

            FullRef: fd.field('Title'),
            Revision: fd.field("Revision"),
            Remarks: fd.field('Remarks'),
            WFStatus: fd.field('WFStatus'),
            Code: fd.field('Code'),
            RejRemarks: fd.field('RejRemarks')
     };

     await handleApproval();

     await renderAUSFields();

      if(!_isNew){
        await buildAusMPForm(fields.Reference.value);
        await setTableRowsReadOnly();
        await setIds();
      }

      $(fields.WFStatus.$parent.$el).hide();
}

//#region APPROVAL STATUS
let handleApproval = async function(){
    _isSenior = await IsUserInGroup('Quality Seniors');
    _isManager = await IsUserInGroup('Quality Manager');

    console.log(`_isSenior is ${_isSenior}`)
    console.log(`_isManager is ${_isManager}`)

    if(_isNew){
        if(!_isSenior){
            alert('Only Quality Senior is allowed to fill the form')
            fd.close()
        }
    }
   
    if(_isEdit){
        debugger;
        let wfStatus = fields.WFStatus.value
        let code = fields.Code.value


        fields.Code.$on("change", async function (value){
            let isReq = false
            if(value === 'Reject')
                isReq = true
                
            if(isReq)
                $(fields.RejRemarks.$parent.$el).show();
            else $(fields.RejRemarks.$parent.$el).hide();
    
            fields.RejRemarks.required = isReq
        })

        if(code !== 'Approve'){
            
            disableAuditTab()
            
            if(code === '')
                $(fields.RejRemarks.$parent.$el).hide();
        }
        else  $(fields.RejRemarks.$parent.$el).hide();


       if(!_isManager && !_isSenior){
            alert('You are not allowed to view the form')
            fd.close()
            return;
       }

        if(_isManager && wfStatus === 'Sent to Manager'){
            fields.Code.required = true;
            enable_Disable_Tabs(tabs, true)
        }
        else if(_isSenior && (wfStatus === 'Pending' || wfStatus === 'Return to Senior')){
            enable_Disable_Tabs(tabs, true)

            if(wfStatus === 'Pending')
                $(fields.Code.$parent.$el).hide();
            else{
                fields.Code.disabled = true
                fields.RejRemarks.disabled = true
            }
        }
        else{
            if(wfStatus === 'Completed'){
                await disableSubmission('form is already completed');
            }
            else{
                let mesg = 'Only Quality Manager is allowed to Approve the form';
                if(_isManager && (wfStatus === 'Pending' || wfStatus === 'Return to Senior'))
                    mesg = 'Only Quality Senior is allowed to Submit the form';

                alert(mesg)
                fd.close()
            }
        }

      
    }


}

let setFormReadOnly = async function(){

}

function disableAuditTab(){
    fd.container('Tabs1').tabs[1].disabled = true;
    setTimeout(() => {
        $('.fd-form .tabset .disabled').hide();
    }, 200); 
}
//#endregion


var renderAUSFields = async function(){
    
  let hideFields = [fields.Reference, fields.Year.field, fields.Office.field, fields.Discipline.field, fields.Project.field,
         fields.ProjTitle, fields.Remarks, fields.Revision, fields.FullRef];
        disFields = [fields.Year, fields.Office.field, fields.Discipline];

   fields.Reference.disabled = true;
   fields.Revision.disabled = true;

   _HideFormFields(hideFields, true);

    if(_isEdit)
      await getDictionaries();

    if (_isNew){
        
        $(dtTbl).hide();
        
        fields.AuditType.field.$on("change", async function (value){
          
            clearStoragedFields(fd.spForm.fields, 'AuditType');
    
            let tempHideFields = await renderFieldsPerType(value);
            _HideFormFields(tempHideFields, false);

            await _DisableFormFields([fields.AuditType.field], true)

            if(_isNew)
                await handleFieldsChanges();

            setAuditType(value)
        });
    }
    else if (_isEdit){
       
        $(fields.FullRef.$parent.$el).hide();
        disFields.push(fields.Reference, fields.AuditType.field)
        _DisableFormFields(disFields, true)

        let auditType = fields.AuditType.field.value
        let tempHideFields = await renderFieldsPerType(auditType);
        _HideFormFields(tempHideFields, false);
        _DisableFormFields(tempHideFields, true)

        setAuditType(auditType)      
        await setHTMLTable().then(()=> $(dtTbl).show());
    }
}

let renderFieldsPerType = async function(auditType){

    let showfeidls = [];
    let placeHolder = '';

    _isDepartmental = auditType === 'Departmental' ? true : false;
    _isCompany = auditType === 'Company' ? true : false;
    _isProject = auditType === 'Project' ? true : false;
    _isUnschedule = auditType === 'Unscheduled' ? true : false;

    let year = fields.Year;
    let office = fields.Office.field;
    office.required = true;

    if(_isCompany){
        showfeidls.push(fields.Reference, year.field, office)
        placeHolder = `${fields.Office.before}-${year.before}-CO`

        year.field.required = true;
    }
    else if(_isDepartmental){
        let discipline = fields.Discipline
        showfeidls.push(fields.Reference, year.field, office, discipline.field)
        placeHolder = `${fields.Office.before}-${year.before}-${discipline.before}`

        year.field.required = true;
        discipline.field.required = true;
    }
    else if(_isProject){
        showfeidls.push(fields.Reference, office, fields.Project.field, fields.ProjTitle)
        placeHolder = `${fields.Project.before}`

        fields.Project.field.required = true;
        fields.ProjTitle.disabled = true;

        $("label:contains('Office')").parent().parent().parent().parent().prev().remove()
        fields.ProjTitle.value = '';
    }

    if( placeHolder !== '')
        placeHolder += '-AS'
    
    fields.Reference.placeholder = placeHolder;

    showfeidls.push(fields.Revision, fields.Remarks)
    return showfeidls;
}

let handleFieldsChanges = async function(){
    fields.Year.field.$on("change", async function (value){

        if(value !== null) // && (_isCompany || _isDepartmental))
           await setSchemaFieldMetaInfo(fields, fields.Year, value, true)
            .then(async ()=>{
                 await generateRevision();
            });
    });

    fields.Office.field.$on("change", async function (value){
        if(value !== null)
           await setSchemaFieldMetaInfo(fields, fields.Office, value.Acronym, false)
            .then(async ()=>{
                  await generateRevision();
            })
    });

    fields.Discipline.field.$on("change", async function (value){ 
        if(value !== null && _isDepartmental)
           await setSchemaFieldMetaInfo(fields, fields.Discipline, value.Acronym, false)
            .then(async ()=>{
                await generateRevision();
            });
    });

    fields.Project.field.$on("change", async function (value){
   
        if(value !== null && _isProject)
           fields.ProjTitle.value = value.ProjectName;
           await setSchemaFieldMetaInfo(fields, fields.Project, value.LookupValue, false)
            .then(async ()=>{
                await generateRevision();
            });
    });
}

const setSchemaFieldMetaInfo = async function(fields, field, value, isDate){

    let last_value = '';
    
    last_value = isDate ? value.slice(-2) : value;
 
     if(value === null){
         field.before = field.after;
         field.after = '';
         return;
     }
 
     if(field.before === field.after){
         field.before = last_value;
         field.after = last_value;
     }
     else if(field.before !== '' && field.after === '')
         field.after = last_value;
     else {
         field.before = field.after;
         field.after  = last_value;
     }
     
     replaceReference(fields, field.before, field.after);
}

const validateReferenceInput = async function(){
    let isReady = false;
    if(_isCompany || _isDepartmental){
        let year = fields.Year?.field?.value || '';
        let office = fields.Office.field?.value || '';

        if(_isDepartmental){
            let discipline = fields.Discipline?.field?.value || '';

            if(year !== '' && office !== '' && discipline !== '')
                isReady = true
        }
        else {
            if(year !== '' && office !== '')
                isReady = true
        }
    }
    else if(_isProject){
        let office = fields.Office.field?.value || '';
        let project = fields.Project?.field?.value || '';
        if(office !== '' && project !== '')
            isReady = true
    }
    return isReady;
}

var replaceReference = async function(fields, oldValue, newValue){
    let reference = fields.Reference;
     reference.placeholder = reference.placeholder.replace(oldValue, newValue);
     
     
     if(_module === 'MAP'){
        let placeholder = reference.placeholder;
        let officeAcronym = fields.Office.field.value !== undefined && fields.Office.field.value !== null ? fields.Office.field.value.Acronym : '';
        let isReady = false;

        if(_isDepartmental){
            let discAcronym = fields.Discipline.field.value !== undefined && fields.Discipline.field.value !== null? fields.Discipline.field.value.Acronym : '';
            if(discAcronym !== ''  && officeAcronym !== '') 
                isReady = true
        }
        else if(_isProject || _isUnschedule){
            let projectCode = fields.Project.field.value !== undefined && fields.Project.field.value !== null ? fields.Project.field.value.LookupValue : '';
            if(_isProject){
            if(projectCode !== '')
                isReady = true
            }
            else{
                if(projectCode !== '' && officeAcronym !== '')
                    isReady = true
            }
        }
        else if(_isCompany){
            if(officeAcronym !== '')
                isReady = true
        }

        if(isReady){
            //INITIATE SEQUENCE COUNTER
            let type = placeholder.substring(0, placeholder.lastIndexOf('/'));
            let sequenceNum = await getSequence(type);
            placeholder = `${type}/${sequenceNum}`;
            reference.placeholder = placeholder;
        }
    }
}

let generateRevision = async function(){
    let isReady = false;
   
    if(_module === 'AUS'){
        isReady =  await validateReferenceInput()

        if(isReady){
            await validateRefExist()
            await getDictionaries();
            let result = await getRev();

            let currentRev = result.currentRev;
            // let prevRev = result.prevRev;

            // let lblText = 'Previous Revision'
            // let lblCtlr = $(`label:contains(${lblText})`);

            // if(lblCtlr.length === 0)
            //     createLabelCtlr($("label:contains('Reference')").parent(), lblText, prevRev, 'padding-right: 69px; padding-bottom:15px', '');

            fields.FullRef.value = `${fields.Reference.placeholder}-${currentRev}`
            fields.Revision.value = currentRev

            $(dtTbl).show()
            await setHTMLTable()
         }
    }
}

let getRev = async function(){
    
    //$(fields.FullRef.$parent.$el).show();
    let reference = fields.Reference.value;
    //$(fields.FullRef.$parent.$el).hide();
    //let ref = FullRef.substring(0, FullRef.lastIndexOf('-'));
    //let revField = FullRef.substring(FullRef.lastIndexOf('-')+1);

    let prevRev = 'NA', currentRev = 0;
    let query = `Reference eq '${reference}' and IsLatest eq 1`
    let items = await _web.lists.getByTitle(AuditSchedule).items.select("Id, Revision").filter(query).get();

    if (items.length > 0){
        items.forEach(item => {
            let revision = parseInt(item.Revision) || 0;  // Use 0 if Revision is undefined, null, or an empty string
            if (revision > currentRev) {

                currentRev = revision;

                if(currentRev > 0)
                  prevRev = currentRev;
            }
        });
        currentRev++;
    }

    currentRev = String(currentRev).padStart(2, '0');
    prevRev = String(prevRev).padStart(2, '0'); 

    //fields.Reference.placeholder = FullRef.replace(revField, currentRev);
    return {
        currentRev: currentRev,
        prevRev: prevRev
    }
}

let setIds = async function(){
    let reference = fields.Reference.value
    let items = await _web.lists.getByTitle(AuditSchedule).items.select("Id").filter(`Reference eq '${reference}'`).get();
    if (items.length > 0){
        let masterId = items[0].Id;
        localStorage.setItem('MasterID', masterId)
    }
}


//#region GET DICTIONARIES
var getProcedure = async function(listname){
	var _colArray = [];
    let select = 'Id,Title,Description,AuditType/Title',
        expand = 'AuditType',
        filter = `Title ne null and AuditType/Title eq '${fields.AuditType.field.value}'`;

    if(listname === officeProceduresList){
        select = 'Id,Title,Description,AuditType/Title,Office/Title',
        expand = 'AuditType,Office',
        filter = `Title ne null and AuditType/Title eq '${fields.AuditType.field.value}' and Office/Title eq '${fields.Office.field.value.LookupValue}'`;
    }
    else if(listname === departmentsList)
        filter += ' and IsCustom eq 1'
	
	await pnp.sp.web.lists
		  .getByTitle(listname)
		  .items
		  .select(select)
          .expand(expand)
		  .filter(filter)
		  .get()
		  .then(async function (items) {
			  if(items.length > 0)
				for (var i = 0; i < items.length; i++){
                     let Id = items[i]['Id'];
                     let Title = items[i]['Title'];
			         let val = items[i]['Description'];

                     if(listname === officeProceduresList){
                        _colArray.push({ 'Id': Id, 'Title': Title, 'Description': val });
                     }
                     else{
                        if(!_colArray.includes(Title))                     
                            _colArray.push({ 'Id': Id, 'Title': Title });
                     }
				}
			});
   return _colArray;
}

let getDictionaries = async function(){
    activities = await getProcedure(officeProceduresList);
    activities = activities.sort((a, b) => a.Description.localeCompare(b.Description));

    disciplines = await getProcedure(departmentsList);
    disciplines = disciplines.sort((a, b) => a.Title.localeCompare(b.Title)); //.filter(Boolean).sort((a, b) => a.localeCompare(b))
}
//#endregion



var setHTMLTable = async function(){

  if($(`#${tableId}`).length === 0){
    let actTdWidth = _isDepartmental ? '35%' : '30%'
    let tbl = `<table id='${tableId}' width='100%' class='modern-table'>
                <tr>
                  <th width='3%'><label class="d-flex justify-content-center fd-field-title col-form-label">No.</label></th>
                  <th width='30%'><label class="d-flex fd-field-title col-form-label">Activity</label></th>
                  <th width='10%'><label class="d-flex fd-field-title col-form-label">Remarks</label></th>
                  `
                  //<th width='12%'><label class="d-flex fd-field-title col-form-label">Scope</label></th>`

    tbl += await renderTblCols();
      
    tbl += '</tr>';

    let mainId;
    if(_isNew){
        let rows = 1;
         for(var i = 0; i < rows; i++){
            tbl += '<tr>'
            tbl += await renderRows(i)
            tbl += '</tr>'
            mainId = i;
        }
    }
    else{
        //get rows
   
        let items = await getItems();

        const groupedData = items.reduce((acc, item) => {
            const itemNo = item.ItemNo;
        
            if (!acc[itemNo]) {
                acc[itemNo] = {
                    ItemNo: itemNo,
                    Procedures: item.Procedures,
                    Disciplines: item.Discipline,
                    Remarks: item.Remarks,
                    SubItems: []
                };
            }
        
            acc[itemNo].SubItems.push({
                SubItemNo: item.SubItemNo,
                Date: item.Date,
                Status: item.Status,
                ID: item.ID
            });
        
            return acc;
        }, {});

        console.log(groupedData)
        const groupedDataLength = Object.keys(groupedData).length;

        let i = 0;
        for (const itemNo in groupedData) {
            const item = groupedData[itemNo];
            
            tbl += '<tr>'
            tbl += await renderRows(i, item)
            tbl += '</tr>'
            mainId = i+1;
            i++;
        }
    }
   
    tbl += `<tr>
              <td colspan='5' style='font-weight:normal; text-align:right;'>
                <a href="javascript:void(0);" onclick="createNewRow(this, false);" mainId=${mainId}>Add New Activity</a>
              </td>
            </tr>`

    tbl += `</table>`;
    $(dtTbl).append(tbl);

    //setTimeout(function () {
     await bindSelect();
         
      //}, 1000);
  }
}

let renderTblCols = async function(){
    let cols = '';

    if(_isProject)
        cols += '<th  width="10%"><label class="d-flex fd-field-title col-form-label">Discipline</label></th>';
    // else if(_isDepartmental){
    //     months.forEach(month => {
    //         cols += `<th><label class="d-flex fd-field-title col-form-label">${month}</label></th>`
    //         });
    // }

    cols += '<th><label class="d-flex fd-field-title col-form-label">Dates</label></th>'
    cols += '<th></th>' //FOR DELETE
    return cols;
}

let renderRows = async function(index, item){
  //masterNo //.toString().padStart(2, '0');
  let masterNo = item === undefined ? index + 1 : item.ItemNo;
  let tblRows = `<td style="font-weight:normal; text-align: center; vertical-align: top;">${masterNo}</td>`;
  
  tblRows += `<td style="font-weight:normal; vertical-align: top;">`;

  let proc = item !== undefined ? item.Procedures : undefined;
  let remarks = item !== undefined ? item.Remarks : '';
  let actWidth = '530'; //_isDepartmental ? 530 : 530

  tblRows +=  createSelect(`act${index}`, actWidth, false, activities, proc);
  tblRows += '</td>';

  tblRows += `<td style="font-weight:normal; vertical-align: top;">
               <textarea id="rem${index}" rows="4" cols="50" maxlength="255" width="100px">${remarks}</textarea>
              </td>`;

  if(_isProject){
    tblRows += `<td style="font-weight:normal; vertical-align: top;">`;

    let disc = item === undefined ? undefined: item.Discipline;
    tblRows +=  createSelect(`dis${index}`, 300, true, disciplines, disc);
    tblRows += '</td>';
  }

  tblRows += `<td style="font-weight:normal;">`;
  tblRows +=  await DatesTable(masterNo, item);
  tblRows += '</td>';

  tblRows += addDeleteCell(masterNo);
  
  return tblRows;
}

function createSelect(id, width, isMulti, options, items){

    let multiple = '';
    if(isMulti)
        multiple = 'multiple size="4"'
    
    multi.push({
        id: id,
        width: width
    });

    let select = `<select Id="${id}" class="styled-select" ${multiple} >`;
  
    let isObject = id.startsWith('stat') ? false : true;

    options.map(option => {
        // _colArray.push({ 'Id': Id, 'Title': Title, 'Description': val });
        if(isObject){
            
            let isSelected = '';
            let text = option.Description === undefined ? option.Title : option.Description;

            if(items !== undefined){
                let isMatched = items.find(item => {
                    let itemText = item.Description === undefined ? item.Title : item.Description;
                    if(text.trim().toLowerCase() === itemText.trim().toLowerCase())
                        return true
                });           
                isSelected = isMatched ? 'selected' : '';
            }

           select += `<option value="${text}" lookupid="${option.Id}" ${isSelected}>${text}</option>`
        }
        else {
            let isSelected = '';
            if(items !== undefined){

                let isMatched;
                if (typeof items === 'string')
                {
                  isMatched = option.trim().toLowerCase() === items.trim().toLowerCase() ? true : false;
                }
                else{
                    isMatched = items.find(item => {
                        let itemText = item.Description === undefined ? item.Title : item.Description;
                        if(option.trim().toLowerCase() === itemText.trim().toLowerCase())
                            return true
                    });
                }
                isSelected = isMatched ? 'selected' : '';
            }
            select += `<option value="${option}" ${isSelected}>${option}</option>`
        }
    }).join('')
    select += '</select>';
    return select;
}

let bindSelect = async function(){

    multi.map(ctl =>{
        let id = '#'  + ctl.id;
        let width =  ctl.width;
        let maxOptionWidth = width - 61;
        
        let mySelect = new vanillaSelectBox(id, {
            maxWidth: 650,
            //maxHeight: 250,
            maxOptionWidth: maxOptionWidth,
            search: false,
            disableSelectAll: true
        });

        let vsbMain = $(id).next();
        //$('div.vsb-main')

        vsbMain.css({
            'position': 'unset',
            'width': width
        });

    
        vsbMain.find('span.caret').css({
            'position': 'relative',
            //'left': width - 40,
            'vertical-align': 'middle',
            'margin-top': 0
        });
    })
    multi = [];
}

let DatesTable = async function(masterNo, item){

    let tbl = `<table id='${tblDatesId}${masterNo}' width='80%' class='modern-table'>
                <tr>
                    <th width='5%'><label class="d-flex justify-content-center fd-field-title col-form-label">No.</label></th>
                    <th width='35%'><label class="d-flex fd-field-title col-form-label">Date</label></th>
                    <th width='35%'><label class="d-flex fd-field-title col-form-label">Status</label></th>
                    <th width="5%"></th>
                <tr>`
              
    if(_isNew || (_isEdit && item === undefined)){
        let rows = 1;
        for(var i = 0; i < rows; i++){
           tbl += '<tr>'
           tbl += await renderDateRows(masterNo, i)
           tbl += '</tr>'
        }
    }
    else{
        let subitems = item.SubItems
        for(var i = 0; i < subitems.length; i++){
            tbl += '<tr>'
            tbl += await renderDateRows(masterNo, i, subitems[i])
            tbl += '</tr>'
         }
    }

    tbl += `<tr>
              <td colspan='4' style='font-weight:normal; text-align:right;'>
                <a href="javascript:void(0);" onclick="createNewRow(this, true);" masterNo = ${masterNo}>Add New Date</a>
              </td>
            </tr>
          </table>`;
    return tbl;
}

let renderDateRows = async function(masterNo, index, item){

    let subItemNo = item === undefined ? `${masterNo}.${index+1}` : item.SubItemNo
    let subItemId = subItemNo.replace('.','-')

    let  tblRows = `<td style="font-weight:normal; text-align: center">`;
    tblRows +=  subItemNo
    tblRows += '</td>';

    let dateVal = item !== undefined ? item.Date : ''
    dateVal = dateVal !== '' ? dateVal.split('T')[0] : '';

    tblRows += `<td style="font-weight:normal;">`;
    tblRows +=  `<input type='date' id="date${subItemId}" class="dateStyle" value=${dateVal}></input>`
    tblRows += '</td>';
    
    tblRows += `<td style="font-weight:normal;">`;

    let statVal = item !== undefined ? item.Status : ''
    tblRows +=    createSelect(`stat${subItemId}`, 220, false, statAcros, statVal);
    tblRows += '</td>';

    tblRows += addDeleteCell(subItemNo);

    return tblRows;
}

let createNewRow = async function(linkCtlr, isSub){

    let masterNo = isSub ? linkCtlr.getAttribute('masterNo') : linkCtlr.getAttribute('mainId')

    let tblId = isSub ? $(linkCtlr).parent().parent().parent().parent().attr('id') : tableId
    let latestTr =  $(`#${tblId} tr:last`)
    let trIndex = latestTr[0].rowIndex - 1;
  
    let targetTrElement = $(`#${tblId} > tbody > tr`).eq(trIndex)  //isSub ? $(`#${tblDatesId} tr`).eq(trIndex) : $(`#${tableId} > tbody > tr`).eq(trIndex) 
    
    let itemNo = targetTrElement.find('td:first').text();
    if(itemNo === '')
        itemNo = isSub ? '1.1' : '0';
    
    let index;
    if(isSub){
        let parts = itemNo.split('.');
        index = parseInt(parts[1]);
    }
    else index = parseInt(itemNo)

    let newRow = '<tr>'
    newRow += isSub ? await renderDateRows(masterNo, index) : await renderRows(index);
    newRow += '</tr>'

    latestTr.before(newRow);
    
    await bindSelect();
}

function addDeleteCell(itemNo){
    let imgUrl = `${_webUrl}${_layout}/Images/delete.png`;
    return `<td style="text-align: center; vertical-align: top;">
              <a href="javascript:void(0);" onclick="deleteRow(this);" title="remove ${itemNo}"> 
                <img src="${imgUrl}" alt="Delete" style="width: 20px; height: 20px;">
              </a>
            </td>`;
}

function deleteRow(element){
    if (confirm(`Are you sure you want to ${element.title} ?`)){
        const row = element.closest('tr');
        row.parentNode.removeChild(row);
    }
}





function setAuditType(value){
    _isDepartmental = value === 'Departmental' ? true : false;
    _isCompany = value === 'Company' ? true : false;
    _isProject = value === 'Project' ? true : false;
    _isUnschedule = value === 'Unscheduled' ? true : false;
}

const getItems = async function(){
	return await _web.lists.getByTitle(activitiesList).items
    .select("Id,ItemNo,Procedures/Description,Department/Title,SubItemNo,Remarks,Date,Status")
    .expand('Procedures,Department')
    .filter(`MasterID/Reference eq '${fields.Reference.value}'`)
    .getAll()
}

function bind_MultiSelect_notUsed(items, ctlId){
$(ctlId).find('.vsb-menu ul li').each(function(index, li){
    let $li = $(li); // Use jQuery to wrap the li element
    let isMatched = items.find(item => {
        item.Description.trim().toLowerCase() === $li.attr('data-text').trim().toLowerCase()
   });
    if(isMatched)
      $li.addClass('active');
  });
}


//#region INSERTION FUNCTIONS
let addAUSItem = async function(){
   
    let item = {};

    let officeId = fields.Office.field.value ? fields.Office.field.value.LookupId : null;
    let disciplineId = fields.Discipline.field.value ? fields.Discipline.field.value.LookupId : null;
    let projectId = fields.Project.field.value ? fields.Project.field.value.LookupId : null;

    let reference = _isNew ? fields.Reference.placeholder : fields.Reference.value
    let fullRef = fields.FullRef.value

    let code = fields.Code.value;
    let rejRemarks = fields.RejRemarks.value;

    item['Title'] = fullRef //FullRef.substring(0, FullRef.lastIndexOf('-'))
    item['Reference'] = reference

    item['Year'] = fields.Year.field.value

    if(officeId !== null)
        item['OfficeId'] = officeId

    if(disciplineId !== null)
        item['DisciplineId'] = disciplineId

    item['AuditType'] = fields.AuditType.field.value

    if(projectId !== null)
        item['ProjectId'] = projectId
    item['Revision'] = fields.Revision.value

    item['Remarks'] = fields.Remarks.value

    debugger;
    if(_isSenior)
      item['WFStatus'] = 'Sent to Manager';

    if(_isManager &&  code === 'Approve')
        item['WFStatus'] = 'Completed';
    else  if(_isManager &&  code === 'Reject')
        item['WFStatus'] = 'Return to Senior';
    
    
    if(code)
      item['Code'] = code

    if(rejRemarks)
        item['RejRemarks'] = rejRemarks;

    addUpdate(AuditSchedule, `Title eq '${fullRef}'`, item, true, reference)
    await addAUSActivities(); //MASTER PLAN & ACTIVITIES
}

const addUpdate = async function (listname, query, itemMetaInfo, isAusSubmit, itemNo, reference) {
    try {
        let items = await _web.lists.getByTitle(listname).items
            .select("Id")
            .filter(query)
            .get();

        if (items.length === 0){
            if (isAusSubmit) {
                await _web.lists.getByTitle(listname).items.add(itemMetaInfo);
                await addAUSActivities();
            } else {
                if(listname === MasterPlan){
                    const mapRef = await getMAPReference().then(async (ref) => {
                        //if (_isEdit) return ref;
                        let sequence = await getMAPSequence(ref, true);
                        return ref += `/${sequence}`;
                    });
                    itemMetaInfo['Title'] = mapRef;
                }
                else if(listname === activitiesList){
                 
                    let mpQuery = `MasterID/Reference eq '${itemMetaInfo.Title}' and ItemNo eq '${itemNo}'`;
                    let mpItems = await _web.lists.getByTitle(MasterPlan).items.select("Title").filter(mpQuery).get();
                    if (mpItems.length > 0){
                        let item = mpItems[0];
                        itemMetaInfo['Title'] = item.Title;
                    }
                }
                await _web.lists.getByTitle(listname).items.add(itemMetaInfo);
            }
        } else {
            let item = items[0];
            if(listname === MasterPlan)
                delete itemMetaInfo['AuditType'];

            if(listname === MasterPlan || listname === activitiesList)
                delete itemMetaInfo['MasterIDId'];

            if (isAusSubmit) {
                await _web.lists.getByTitle(listname).items.getById(item.Id).update(itemMetaInfo);
                await addAUSActivities();
            } else {
                await _web.lists.getByTitle(listname).items.getById(item.Id).update(itemMetaInfo);
            }
        }
    } catch (error) {
        console.error(error);
    }
}

let addAUSActivities = async function (ignoreMasterId) {
    let officeId = fields.Office.field.value ? fields.Office.field.value.LookupId : null;
    let disciplineId = fields.Discipline.field.value ? fields.Discipline.field.value.LookupId : null;
    let reference = _isNew ? fields.Reference.placeholder : fields.Reference.value;
    let auditType = fields.AuditType.field.value;

    // Fetch the Master Plan ID if it doesn't exist
    if (isNaN(_itemId)){
        let items = await _web.lists.getByTitle(AuditSchedule).items.select("Id").filter(`Reference eq '${reference}'`).get();
        if (items.length > 0){
            _itemId = items[0].Id;
        }
        //return;
    }

    // Retrieve all rows (excluding the first row) and iterate through them
    const rows = Array.from($(`#${tableId} > tbody > tr:gt(0)`));

    for (let index = 0; index < rows.length; index++) {
        const element = rows[index];

        const selectedActOptions = $(element).find('select[id^="act"] option:selected');
        const selectedDisOptions = $(element).find('select[id^="dis"] option:selected');

        // Check if activities are selected
        if (selectedActOptions.length > 0) {
            let item = {};
            let itemNo = $(element).find('td:first').text();
            if (itemNo === '') return true;

            
            item['MasterIDId'] = _itemId;
            item['ItemNo'] = itemNo;

            // Multi-select Procedures/Activities
            let lookupIds = [];
            selectedActOptions.each(function () {
                let lookupId = $(this).attr('lookupid');
                lookupIds.push(parseInt(lookupId));
            });
            if (lookupIds.length > 0) item['ProceduresId'] = { results: lookupIds };

            // Multi-select Disciplines
            lookupIds = [];
            if (_isProject) {
                selectedDisOptions.each(function () {
                    let lookupId = $(this).attr('lookupid');
                    lookupIds.push(parseInt(lookupId));
                });
                if (lookupIds.length > 0) item['DepartmentId'] = { results: lookupIds };
            } else {
                if (disciplineId !== undefined && disciplineId !== null) {
                    lookupIds.push(parseInt(disciplineId));
                    item['DepartmentId'] = { results: lookupIds };
                }
            }

            let remarks = $(`#rem${index}`).val();
            if(remarks === null)
                remarks = '';
            
            item['AuditType'] = auditType;
            item['OfficeId'] = officeId;

            let query = `MasterID/Reference eq '${reference}' and ItemNo eq '${itemNo}'`;
            debugger;
            await addUpdate(MasterPlan, query, item, false);
            item['MasterIDId'] = _itemId;

            delete item['AuditType'];
            delete item['OfficeId'];
            //delete item['DepartmentId'];

            // Iterate through sub-rows for additional details
            const tblId = `tblDatesId${element.rowIndex}`;
            const subRows = $(element).find(`#${tblId} tbody tr:gt(0)`);
            for (let j = 0; j < subRows.length; j++) {
                const subRow = subRows[j];

                let subItemNo = $(subRow).find('td:first').text();
                const dateValue = $(subRow).find('input[type="date"]').val();
                const statusValue = $(subRow).find('select[id^="stat"]').val();

                if (!subItemNo || !dateValue || !statusValue) continue;

                item['SubItemNo'] = subItemNo;
                item['Date'] = dateValue;
                item['Status'] = statusValue;
                item['Remarks'] = remarks;

                let query = `MasterID/Reference eq '${reference}' and ItemNo eq '${itemNo}' and SubItemNo eq '${subItemNo}'`;
                await addUpdate(activitiesList, query, item, false, itemNo);
            }
        }
    }
}

let getMAPReference = async function(){

    // debugger;
    // if(_isEdit && isMain)
    //     return fields.Reference.value;

    let ref = ''
    let office = fields.Office.field.value.Acronym;

    //just for company and department

    let year = fields.Year?.field?.value  ? fields.Year.field.value.slice(-2) : '';

    if(_isDepartmental){
        let discipline = fields.Discipline.field.value !== null ? fields.Discipline.field.value.Acronym : '';
        ref = `${office}-${discipline}${year}`;
    }
    else if(_isProject){
        let project = fields.Project.field.value !== null ? fields.Project.field.value.LookupValue : '';
        ref = `${project}`;
    }
    else if(_isCompany) 
        ref = `${office}-${year}`;

    // let sequence = await getMAPSequence(ref)
    //     return ref += `/${sequence}`
    
    return ref;
    //else if(_isUnschedule) reference.placeholder = `X-UNS/${year}/NN`;
}

let getMAPSequence = async function(type, isAUS){
    let seq;
    let filterItem = counterTemp.filter(item => item.type === type);
    if (filterItem.length === 0 || (isAUS && _isNew)){
        seq = await getCounter(_web, type, true);
        //console.log(`${type} sequennce is ${seq}`);
      counterTemp.push({
        type: type,
        seq: seq
      })
    }
    else seq = filterItem[0].seq;
    return seq.toString().padStart(2, '0');
}
//#endregion


let setTableRowsReadOnly = async function(){
    const rows = Array.from($(`#${tableId} > tbody > tr:gt(0)`));
    let isAllComp = true;
    for (let index = 0; index < rows.length; index++) {
    
        const element = rows[index];

        const itemNo = $(element).find('td:first').text();
        const selectedActOptions = $(element).find('select[id^="act"]');
        const selectedDisOptions = $(element).find('select[id^="dis"]');
        let remarks = $(`#rem${index}`)

        const tblId = `tblDatesId${element.rowIndex}`;
        const subRows = $(element).find(`#${tblId} tbody tr:gt(0)`);

        for (let j = 0; j < subRows.length; j++) {
            const subRow = subRows[j];

            const subItemNo = $(subRow).find('td:first').text();
            const date = $(subRow).find('input[type="date"]')
            const status = $(subRow).find('select[id^="stat"]')
                
            if(status.length === 0)
                continue;

            const statusVal = status.val()
            if(statusVal === 'Complete'){
                
                const actId = status.attr('id');
                let selectBox = new vanillaSelectBox(`#${actId}`,{});
                selectBox.disable();

                if(selectedDisOptions.length > 0){
                  const disId = selectedDisOptions.attr('id');
                  const selectBox1 = new vanillaSelectBox(`#${disId}`,{});
                  selectBox1.disable();
                }

                remarks.attr('disabled', true);

                date.attr('disabled', true);

                const statusId = selectedActOptions.attr('id');
                let selectBox1 = new vanillaSelectBox(`#${statusId}`,{});
                selectBox1.disable();

                let removeLink = $(`a[title="remove ${itemNo}"]`)
                if(removeLink.length > 0)
                    removeLink.remove();

                let removeLink1 = $(`a[title="remove ${subItemNo}"]`)
                if(removeLink1.length > 0)
                    removeLink1.remove();

                $(subRow).next('tr').find('a').filter(function() {
                    return $(this).text().trim() === 'Add New Date';
                }).remove();
            }
            else isAllComp = false;
        }
    }
    
    if(isAllComp){
        localStorage.setItem('isReadOnly', true);
        $('a').filter(function() {
            return $(this).text().trim() === 'Add New Activity';
        }).remove();

        disableSubmission('Audit Schedule is Completed')
    }
}

let disableSubmission = async function(mesg){
    setTimeout(async () => {
        setPSErrorMesg(mesg, true)
        let btn = $('span').filter(function(){ return $(this).text() == _submitText }).parent()
        let genRev = $('span').filter(function(){ return $(this).text() == _genRev }).parent()
        if(mesg !== undefined && mesg !== '')
            btn.attr("disabled", "disabled");
        else btn.removeAttr("disabled");

        if(genRev.length > 0){
            const isReadOnly =  JSON.parse(localStorage.getItem('isReadOnly'));
            if(isReadOnly) 
                genRev.attr("disabled", "disabled");
            else genRev.removeAttr("disabled");
        }
    }, 100);
}

let validateRefExist = async function(){
    let items = await _web.lists.getByTitle(AuditSchedule)
                     .items.select("Id").filter(`Reference eq '${fields.Reference.placeholder}'`).get();

    if(items.length > 0)
        disableSubmission('Reference ia already exist')
    else disableSubmission('')
}


//#region MAKE FORM COPY (GENERATE REVISION BUTTON)
let makeFormCopy = async function(){

  let itemId = await insertMainFields();
  localStorage.setItem('GenRevId', itemId);
}

let insertMainFields = async function(){

  let auditType = fields.AuditType.field.value
  let reference = fields.Reference.value

  debugger;
  let result = await getRev();
  let revision = result.currentRev;

  let fullRef = `${reference}-${revision}`
  let year = fields.Year.field.value
  let officeId = fields.Office.field.value ? fields.Office.field.value.LookupId : null;
  let disciplineId = fields.Discipline.field.value ? fields.Discipline.field.value.LookupId : null;
  let projId = fields.Project.field.value ? fields.Project.field.value.LookupId : null;
  let remarks = fields.Remarks.value

  let item = {};
  item[fields.AuditType.field.internalName] = auditType;
  item[fields.Reference.internalName] = reference;
  item[fields.Revision.internalName] = parseInt(revision);
  item[fields.FullRef.internalName] = fullRef;

  item[fields.Year.field.internalName] = year;
  item[`${fields.Office.field.internalName}Id`] = officeId;

  if(disciplineId !== null)
    item[`${fields.Discipline.field.internalName}Id`] = disciplineId;

  if(projId !== null)
    item[`${fields.Project.field.internalName}Id`] = projId;

  item[fields.Remarks.internalName] = remarks;

  await resolveLatestRev(AuditSchedule, reference)
  const res = await _web.lists.getByTitle(AuditSchedule).items.add(item);
  return res.data.Id;
}
//#endregion

