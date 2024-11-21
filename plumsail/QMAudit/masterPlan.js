//let _isDepartmental = false, _isCompany = false, _isProject = false, _isUnschedule;

var onMAPRender = async function() {

    $('div.SPCanvas, .commandBarWrapper').css('padding-left', '0px');

     fields = {
            AuditType: fd.field("AuditType"),
            Reference: fd.field("Title"),

            Discipline: {
                field: fd.field("Department"),
                before: 'DD', 
                after: ''
            },

            SubDiscipline: fd.field("SubDiscipline"),

            Project: {
                field: fd.field("Project"),
                before: 'PRJNUM', 
                after: ''
            },

            Procedures: fd.field("Procedures"),

            Office: {
                field: fd.field("Office"),
                before: 'X', 
                after: ''
            },

            StartDate: {
                field: fd.field("StartDate"),
                before: 'YY', 
                after: ''
            },
            EndDate: fd.field("EndDate"),
            AuditorTeam: fd.field("AuditorTeam")
    };

    let hFields = [fields.StartDate.field, fields.EndDate];
    $('.k-clear-value').remove();

    
    if(_isNew){
      renderFields(fields, hFields, '');
      fields.AuditType.$on("change", async function (value){
          
          clearStoragedFields(fd.spForm.fields, 'AuditType');
          renderFields(fields, hFields, value);
          renderAuditType(fields);
      });
    }
    else{
        hideArrayFields(hFields, true);
        renderFields(fields, [], fields.AuditType.value);

        
        if(_isDepartmental){
            filterSubDiscipline(fields.Discipline.field.value);
                setTimeout(() => {
                    let subDiscipline = fields.SubDiscipline;

                    if (subDiscipline.widget.dataSource._data.length > 0){
                        subDiscipline.required = true;
                        subDiscipline.disabled = false;
                    }
                    else{
                        subDiscipline.value = null;
                        subDiscipline.required = false; 
                        subDiscipline.disabled = true;
                        }
                }, 300);
            }
    }
    setDateRange(fields.StartDate.field, fields.EndDate);
}

const renderFields = async function(fields, hOrdFields, auditType){

    let reference = fields.Reference;
    reference.disabled = true;

    fields.Office.before = 'X'
    fields.Office.after = ''

    fields.Discipline.before = 'DD'
    fields.Discipline.after = ''

    fields.Project.before = 'PRJNUM'
    fields.Project.after = ''

    fields.StartDate.before = 'YY'
    fields.StartDate.after = ''
    
    if(_isNew || _isEdit){
        _isDepartmental = auditType === 'Departmental' ? true : false;
        _isCompany = auditType === 'Company' ? true : false;
        _isProject = auditType === 'Project' ? true : false;
        _isUnschedule = auditType === 'Unscheduled' ? true : false;
    }

    if(_isNew){
        if(_isProject || _isUnschedule){
            await GetProjectListnofilter();

            let project = fields.Project.field;
            project.required = true;

            project.ready().then(() => {
                project.filter = "Status eq 'Open'";
                project.refresh();
            });
        }
        else fields.Project.field.required = false;

        let year = getYear(fields)
        fields.StartDate.before = year;

        if(_isDepartmental) reference.placeholder = `X-DD${year}/NN`;
        else if(_isProject) reference.placeholder = 'PRJNUM/NN'; 
        else if(_isUnschedule) reference.placeholder = `X-UNS/${year}/NN`;
        else if(_isCompany) reference.placeholder = `X-${year}/NN`;  

        if(_isProject || _isUnschedule)
            setHyperLink();
        else setHyperLink('remove')
    }
    else if(_isEdit){
       
        Array.prototype.push.apply(hOrdFields, [
            fields.AuditType,
            fields.Office.field,
            fields.Discipline.field,
            fields.SubDiscipline,
            fields.Project.field,
            fields.Procedures
        ]);
        // if(_isDepartmental){
        //     let disc = fields.Discipline.field
        //     if(disc.value)
        //       await filterDiscipline_AuditType(disc, fields, fields.Office.field.value, disc.value, officeProceduresList);
        // }

        disableArrayFields(hOrdFields, true);
        hOrdFields = [];
    }

    Array.prototype.push.apply(hOrdFields, [
        reference,
        fields.Office.field,
        fields.Discipline.field,
        fields.SubDiscipline,
        fields.Project.field,
        fields.Procedures,
        fields.AuditorTeam,
        'div._TVO6g'
    ]);

    hideArrayFields(hOrdFields, true);
    if(auditType === '') return;

    let disable = _isProject  || _isUnschedule ? false : true;
    setTimeout(function(){
        fields.Discipline.field.disabled = disable;
        fields.Procedures.disabled = disable;
    }, 200);

    let showFields = [reference, fields.Office.field, fields.Procedures, fields.AuditorTeam, 'div._TVO6g']
    if(_isDepartmental)
        Array.prototype.push.apply(showFields, [ fields.SubDiscipline, fields.Discipline.field ]);
    else if(_isProject || _isUnschedule)
        Array.prototype.push.apply(showFields, [ fields.Project.field ]);
    hideArrayFields(showFields, false);
}

const renderAuditType = async function(fields){

    let reference = fields.Reference,
        discipline = fields.Discipline.field,
        procedure = fields.Procedures;

    setCascadedValues(fields, reference); // FOR SUBDISCIPLINE DEPARTMENTAL

    let office = fields.Office.field;
    office.$on("change", async function (value) {

        if(_isProject || _isUnschedule){
            if(value !== null)
                await setSchemaFieldMetaInfo(fields, fields.Office, value);
        }
        else{
            discipline.clear();
            procedure.clear();
            if(value !== null){
                await setSchemaFieldMetaInfo(fields, fields.Office, value);
                await filterDiscipline_AuditType(this, fields, fields.Office.field.value, null, departmentsList);
            }
        }
    });

    discipline.$on("change", async function (value){
       await filterDiscipline_AuditType(this, fields, fields.Office.field.value, discipline.value, officeProceduresList);
    });


    if(_isProject || _isUnschedule){
        let project = fields.Project.field;
        project.$on("change", async function (value) {
            if(value !== null){
                let projVal = value.LookupValue
                value.Acronym = projVal;
                await setSchemaFieldMetaInfo(fields, fields.Project, value);

                let phase = 'Design';
                let lastChar = projVal.charAt(projVal.length - 1);
                if(lastChar.toLowerCase() =='s')
                    phase = 'Supervision';

                let query = `Phase eq '${phase}'`;
                await filterDiscipline_AuditType(this, fields, fields.Office.field.value, query, officeProceduresList, true);
            }
        });
    }

    procedure.$on("change", function (value) {
            if(value !== null){
                let tempCtlr = this.$el.parentElement.children[0].children[0].children[0].children[0].children[0];
                setTimeout(function(){
                    $('.k-clear-value').remove();
                    $(tempCtlr).prop('readonly', true).css('background-color', 'white'); //'.k-input,.form-control'
                }, 500);
            }
    });

    fields.StartDate.field.$on("change", async function (value){
        if(value !== null){
            const date = new Date(value);
            const year = date.getFullYear().toString().slice(-2);
            await setSchemaFieldMetaInfo(fields, fields.StartDate, year, true);
        }
    });
}

function setCascadedValues(fields) {
      let discipline = fields.Discipline.field;
      let subDiscipline = fields.SubDiscipline;
      
      subDiscipline.disabled = true; 

      discipline.ready()
      .then(function () {
        discipline.orderBy = { field: "Description", desc: false };
        discipline.refresh();
        discipline.$on("change", async function (value) {

            if(value !== undefined && value !== null){
                await setSchemaFieldMetaInfo(fields, fields.Discipline, value);

                if(_isDepartmental){
                filterSubDiscipline(value);
                    setTimeout(() => {
                        if (subDiscipline.widget.dataSource._data.length > 0){
                            subDiscipline.required = true;
                            subDiscipline.disabled = false;
                        }
                        else{
                            subDiscipline.value = null;
                            subDiscipline.required = false; 
                            subDiscipline.disabled = true;
                            }
                    }, 100);
                }
            }
        })
      }) 
}

function filterSubDiscipline(discipline) {
    
    var disciplineId = (discipline && discipline.LookupId) || discipline || null;
    
    if(disciplineId !== null){
        let subDiscipline = fields.SubDiscipline;
        subDiscipline.filter = "Discipline/Id eq " + disciplineId[0].LookupId;
        subDiscipline.orderBy = { field: "Description", desc: false };
        subDiscipline.refresh();
    }
  }

function setDateRange(startDate, endDate){
    const dateRange = $("#dateRange");
    
    dateRange.jqxDateTimeInput({ 
                    width: 300, 
                    height: 32,  
                    selectionMode: 'range', 
                    min: new Date(2024, 1, 1),
                    placeHolder: "select start and end date",
                    textAlign: "center",
                    value: null,
                    animationType: 'fade',
                    theme: 'fluent'
                });

    dateRange.on('change', function (event) {
        let selection = dateRange.jqxDateTimeInput('getRange'); //$(this); //
        if (selection.from != null) {
            startDate.value = selection.from.toLocaleDateString();
            endDate.value = selection.to.toLocaleDateString();
        }
    });
    $('#inputdateRange').attr('placeholder', 'select start and end date');

    if(_isEdit)
        dateRange.jqxDateTimeInput('setRange', startDate.value, endDate.value);
}

var replaceReference = async function(fields, oldValue, newValue){
    let reference = fields.Reference;
    
     reference.placeholder = reference.placeholder.replace(oldValue, newValue);
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

var getSequence = async function(type){
    let seq;
    let filterItem = counterTemp.filter(item => item.type === type);
    if (filterItem.length === 0){
        seq = await getCounter(_web, type, false);
        //console.log(`${type} sequennce is ${seq}`);
      counterTemp.push({
        type: type,
        seq: seq
      })
    }
    else seq = filterItem[0].seq;
    return seq.toString().padStart(2, '0');
}

 
var filterDiscipline_AuditType = async function(ctlr, fields, officeValue, departmentValue, listname, isProjFilter){
    let officeId = (officeValue && officeValue.LookupId) || officeValue || null;
    let deptId, query = '', cols = 'Id,Description,Acronym';
        

    if(_isProject || _isUnschedule){
        if(isProjFilter){
            cols = 'Id,Description';
            fields.Procedures.disabled = false;
            fields.Procedures.widget.dataSource.data([]);
            query = departmentValue;
        }
    }
    else{
        if(officeId === undefined || officeId === null){
            fields.Discipline.field.clear();
            fields.Discipline.field.disabled = true;
            fields.Procedures.disabled = true;
            return;
        }

        if(listname === departmentsList){
            fields.Discipline.field.widget.dataSource.data([]);
            query = `Offices/Id eq ${officeId} and AuditType/Title eq '${fields.AuditType.value}'`;
        }
        else if( listname === officeProceduresList){
                deptId = (departmentValue && departmentValue.LookupId) || departmentValue || null;
                if(deptId === undefined || deptId === null){
                  fields.Procedures.clear();
                  fields.Procedures.disabled = true;
                  return;
                }
                else{
                    cols = 'Id,Description';
                    fields.Procedures.widget.dataSource.data([]);
                    query = `Office/Id eq ${officeId} and Department/Id eq ${deptId}`;
                }
        }
    }
    
    let items = await _web.lists.getByTitle(listname).items
    .select(cols)
    .filter(query)
    .getAll();

    let objArrayItems = [];
    for (const item of items){
        let itemId = item.Id;
        let descValue = item.Description;

        let obj = { LookupId: itemId, LookupValue: descValue, Acronym: item.Acronym};
        objArrayItems.push(obj);
    }
    if(objArrayItems.length > 0){
        objArrayItems.sort((a, b) => {
            if (a.LookupValue < b.LookupValue){
                return -1;
            }
            if (a.LookupValue > b.LookupValue){
                return 1;
            }
            return 0;
        });

        if(listname === departmentsList){
            fields.Discipline.field.ready()
            .then(function () {
                fields.Discipline.field.widget.dataSource.data(objArrayItems);
                fields.Discipline.field.disabled = false;
            })
        }
        else if( listname === officeProceduresList){
            fields.Procedures.ready()
            .then(function () {
                fields.Procedures.widget.dataSource.data(objArrayItems);
                fields.Procedures.disabled = false;
            })
        }
    }

    let tempCtlr = ctlr.$el.parentElement.children[0].children[0].children[0].children[0].children[0];
    setTimeout(function(){
        $('.k-clear-value').remove();
        $(tempCtlr).prop('readonly', true).css('background-color', 'white'); //'.k-input,.form-control'
    }, 500);
   
}



function disableArrayFields(fields, isDisable){
   fields.map(field =>{
    field.disabled = isDisable;
   })
}

const hideArrayFields = async function(fields, isHide){
    fields.map(field =>{
        if(field.toString().includes('div.'))
            isHide ? $(field).hide() : $(field).show()
        else isHide ? $(field.$parent.$el).hide() : $(field.$parent.$el).show()
    })
}

const submitMAPFunction = async function(){
    let placeholder = fd.field("Title").placeholder;
    fd.field("Title").value = placeholder;

    let type = placeholder.substring(0, placeholder.indexOf('/'));
    await getCounter(_web, type, true);
}


async function GetProjectListnofilter(){ 
    var xhr = new XMLHttpRequest();
    var SERVICE_URL = `${_siteUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=allactive`;    
    xhr.open("GET", SERVICE_URL, true);      
    
    xhr.onreadystatechange = async function() 
    {
        if (xhr.readyState == 4)
        {
            try
            {
                if (xhr.status == 200)
                {
                    const parsedObj = await JSON.parse(this.responseText, function (key, value) {
                        if (typeof value === 'string') {
                            return value.trim();
                        }
                        return value;
                    });

                    let result = [];
                    parsedObj.map(item => {
                        result.push({
                            Title: item.ProjectCode,
                            ProjectName: item.ProjectName
                        });
                    });
                    await ensureItemsExist(result);
                }
            }
            catch(err) { console.log(err) }
        }
    }
    xhr.send();
}

async function ensureItemsExist(itemsToEnsure) {
    try {
        const listTitle = 'Projects';        
        const existingItems = await _web.lists.getByTitle(listTitle).items.select('Id,Title').getAll();       
        const existingItemTitles = new Set(existingItems.map(item => item.Title));  
        const itemsToCreate = itemsToEnsure.filter(item => !existingItemTitles.has(item.Title));       

        const itemsToUpdate = existingItems.filter(item => !itemsToEnsure.find(i => i.Title === item.Title));

        for (const item of itemsToCreate) {
            await _web.lists.getByTitle(listTitle).items.add(item);
            console.log(`Item with Title "${item.Title}" created.`);
        }

        for (const item of itemsToUpdate){
            await _web.lists.getByTitle(listTitle).items.getById(item.Id).update({
                Status: 'Completed'
            }); 
            console.log(`Project "${item.Title}" is updated and Status set to Completed.`);
        }
        
        if (itemsToCreate.length === 0) {
            console.log("All items already exist.");
        }
    }catch (error) {
        console.error("Error ensuring items exist:", error);
    }
}

function setHyperLink(trans){
    let link = $('#addDepartmentLink')
    if(trans === 'remove'){
      if(link.length > 0) 
        link.remove();
    }
    else{
        if(link.length === 0){
            var linkHtml = `<a id="addDepartmentLink" href="${_webUrl}/SitePages/PlumsailForms/Departments/Item/NewForm.aspx" target="_blank" style="padding-left: 490px;">Add New Department</a>`
            let element = $('label').filter(function(){ return $(this).text() == 'Department'; });
            let targetElement = element.next().find('select');
            $(linkHtml).insertAfter(targetElement);
        }
    }
}

function getYear(fields){
    const dateRange = $("#dateRange");
    let selection = dateRange.jqxDateTimeInput('getRange'); //$(this); //
    if (selection !== undefined && selection.from != null) {
        fields.StartDate.field.value = selection.from.toLocaleDateString();
        fields.EndDate.value = selection.to.toLocaleDateString();
    }

    let year;
    let startDate = fields.StartDate.field.value;
    if(startDate !== null && startDate !== undefined){
        const date = new Date(startDate);
        year = date.getFullYear().toString().slice(-2);
    }
    else {
        const currentDate = new Date();
        year = String(currentDate.getFullYear()).slice(-2);
    }
    return year;
}

