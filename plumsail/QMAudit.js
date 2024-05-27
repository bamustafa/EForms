var ProjectNameList="", ProjectYear=2010;

var _layout, _module = '', _formType = '', _web, _webUrl, _siteUrl;
var hFields = [], projectArr=[];
var delayTime = 100, retryTime = 10, _timeOut, _timeOut1, _pplTimeOut, _distimeOut;

var  _isMain = true, _isLead = false, _isPart = false, _isNew = false, _isEdit = false, _isSentForReview = false, _isSentForApproval = false, _isAurClosed = false, 
     _hideCode = true, disableCode = false, _ignoreBtnCreation = false;
var _header, _main, _footer;

var auditGenSummary = 'Fill Department and start/end Audit date to generate summary';
var carGenSummary = 'Fill all fields above to proceed';

var aurFields = ['Department', 'Procedures', 'StartAuditDate', 'EndAuditDate', 'Title', 'Office', 'AuditRef'];
var aurPeopleFields = ['AuditTeam', 'Contact', 'Distribution', 'PreparedBy'];

var auditTeamGroup = 'Audit Team';
var activeTabName = '';
var auditReportTab = 'Audit Report';
var corrActionTab = 'Corrective Action';

var departmentsList = 'Departments', officeProceduresList = 'Office Procedures';

var _htLibraryUrl, _hot, _container, _colArray = [], colsInternal = [], counterTemp = [], _targetList;

var onRender = async function (relativeLayoutPath, moduleName, formType) {

    try {
        await getGlobalParameters(relativeLayoutPath, moduleName, formType); // GET LAYOUT PATH FROM SOAP SERVICE

        if(_module === 'MAP')
          await onMAPRender();
        else if(_module === 'DCC')
          await onDCCRender();
        else if(_module === 'AUR')
          await onAURender();
        else if(_module === 'AUS')
          await onAUSRender();
       
        await renderControls();
        
        if(_isAurClosed){
            activeTabName = corrActionTab;
            await renderTabs();
            await handleCarTab();
        }

        _timeOut = setInterval(fixEnhancedRichText, 1000);
        _distimeOut = setInterval(adjustDisableOpacity, 1000);
    }
    catch (e) {
        console.log(e);
    }
}

//#region MAP RENDER MODULE (MASTER PLAN)
var onMAPRender = async function() {
    let fields = {
            Reference: fd.field("Title"),

            Discipline: {
                field: fd.field("Department"),
                before: 'department Acronym', 
                after: ''
            },

            SubDiscipline: fd.field("SubDiscipline"),
            Procedures: fd.field("Procedures"),

            Office: {
                field: fd.field("Office"),
                before: 'office acronym', 
                after: ''
            },

            StartDate: fd.field("StartDate"),
            EndDate: fd.field("EndDate"),
            AuditorTeam: fd.field("AuditorTeam")
    };

    let reference = fields.Reference;
    let disFields = [reference];
    let hFields = [fields.StartDate, fields.EndDate];
    $('.k-clear-value').remove();
    //$('.k-input,.form-control').prop('readonly', true)
    
    if(_isNew){
        setTimeout(function(){
            fields.Discipline.field.disabled = true;
            fields.Procedures.disabled = true;
        }, 100);
      
      setCascadedValues(fields, reference); // FOR SUBDISCIPLINE

      let office = fields.Office.field;
      office.$on("change", async function (value) {
            fields.Discipline.field.clear();
            fields.Procedures.clear();
            await setSchemaFieldMetaInfo(fields, fields.Office, value);
            await filterDiscipline_AuditType(this, fields, fields.Office.field.value, null, departmentsList);
      });

       fields.Discipline.field.$on("change", async function (value) {
        await filterDiscipline_AuditType(this, fields, fields.Office.field.value, fields.Discipline.field.value, officeProceduresList);
       });

       fields.Procedures.$on("change", function (value) {
            if(value !== null){
                let tempCtlr = this.$el.parentElement.children[0].children[0].children[0].children[0].children[0];
                setTimeout(function(){
                    $('.k-clear-value').remove();
                    $(tempCtlr).prop('readonly', true).css('background-color', 'white'); //'.k-input,.form-control'
                }, 500);
            }
        });

      const currentDate = new Date();
      let getYear = String(currentDate.getFullYear()).slice(-2);
      reference.placeholder = `${fields.Office.before}-${fields.Discipline.before}${getYear}/sequence`; // (Ex:X-DDYY/NN)'
    }
    else {
        Array.prototype.push.apply(disFields, [
            fields.Discipline.field,
            fields.SubDiscipline,
            fields.Procedures,
            fields.Office.field
        ]);
    }

    disableArrayFields(disFields, true);
    hideArrayFields(hFields, true);
    setDateRange(fields.StartDate, fields.EndDate);
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
        })
    
      })
     
      function filterSubDiscipline(discipline) {
        var disciplineId = (discipline && discipline.LookupId) || discipline || null;
        
        subDiscipline.filter = "Discipline/Id eq " + disciplineId;
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
    let discAcronym = fields.Discipline.field.value !== undefined && fields.Discipline.field.value !== null? fields.Discipline.field.value.Acronym : '';
    let officeAcronym = fields.Office.field.value !== undefined && fields.Office.field.value !== null ? fields.Office.field.value.Acronym : '';

     reference.placeholder = reference.placeholder.replace(oldValue, newValue);
     let placeholder = reference.placeholder;

    if(discAcronym !== ''  && officeAcronym !== ''){
        //INITIATE SEQUENCE COUNTER
        let type = placeholder.substring(0, placeholder.indexOf('/'));
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
    return seq;
}

 
var filterDiscipline_AuditType = async function(ctlr, fields, officeValue, departmentValue, listname){
    let officeId = (officeValue && officeValue.LookupId) || officeValue || null;
    let deptId, query = '', cols = 'Id,Description,Acronym';
        
    if(officeId === undefined || officeId === null){
        fields.Discipline.field.clear();
        fields.Discipline.field.disabled = true;
        fields.Procedures.disabled = true;
        return;
    }

    if(listname === departmentsList){
        fields.Discipline.field.widget.dataSource.data([]);
        query = `Offices/Id eq ${officeId}`;
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

    debugger;
    let tempCtlr = ctlr.$el.parentElement.children[0].children[0].children[0].children[0].children[0];
    setTimeout(function(){
        $('.k-clear-value').remove();
        $(tempCtlr).prop('readonly', true).css('background-color', 'white'); //'.k-input,.form-control'
    }, 500);
   
}

const setSchemaFieldMetaInfo = async function(fields, field, value){
    if(value === null){
        field.before = field.after;
        field.after = '';
        return;
    }

    if(field.before === field.after){
        field.before = value.Acronym;
        field.after = value.Acronym;
    }
    else if(field.before !== '' && field.after === ''){
        field.after = value.Acronym;
    }
    else {
         field.before = field.after;
         field.after  = value.Acronym;
        
    }
    replaceReference(fields, field.before, field.after);
}

function disableArrayFields(fields, isDisable){
   fields.map(field =>{
    field.disabled = isDisable;
   })
}

function hideArrayFields(fields, isHide){
    fields.map(field =>{
        isHide ? $(field.$parent.$el).hide() : $(field.$parent.$el).show()
    })
}

const submitMAPFunction = async function(){
    let placeholder = fd.field("Title").placeholder;
    fd.field("Title").value = placeholder;

    let type = placeholder.substring(0, placeholder.indexOf('/'));
    await getCounter(_web, type, true);
}
//#endregion



//#region DCC RENDER MODULE (Drawing Control)
var onDCCRender = async function() {
    if (_isNew || _isEdit) {
        hFields = ['Signature1Yes', 'IDC6Yes', 'IDC9YesCH', 'IDC11YesCH','Title', 'Department2', 'Department3'];
        HideFields(hFields, true);

        if(_isNew)
            setProjectsWithValidation();
        else
        {
            $(fd.field('Title').$parent.$el).show();
            fd.field('Title').disabled = true;
            fd.field('PKnumber').disabled = true;
            fd.field('Department1').disabled = true;
            fd.field('DepUnits').disabled = true;
        }           

        setColumnsArrayShowHide();
    } 
}
//#endregion

//#region AUR RENDER MODULE (AUDIT REPORT)
var onAURender = async function() {
    if (_isNew || _isEdit){

        //#region HANDLE GENERATE AUDIT SUMMARY BUTTON
        fd.field('Department').$on('change', function(value)
    	{
             handleSummaryButton();
    	});	

        fd.field('StartAuditDate').$on('change',  function(value)
    	{
             handleSummaryButton();
    	});

        fd.field('EndAuditDate').$on('change', function(value)
    	{
             handleSummaryButton();
    	});
         handleSummaryButton();
        //#endregion

        if(_isEdit){
            if( activeTabName === corrActionTab)
             await handleCarTab();
            else{
                await handleAurTabs();
                if(!_isAurClosed)
                 fd.container('Tab1').tabs[1].disabled = true;
            }

            hFields = ['SummaryPlainText', 'ReasonOfRejection', 'CARSummaryPlainText', 'Status', 'Submit'];
             if(_hideCode && !disableCode)
              hFields.push('Code');
             else {
                fd.field('Code').required = true;
                fd.field('Code').$on('change', async function(value){
                    var rejFields = 'ReasonOfRejection';
                    if(value === 'Rejected'){
                        $(fd.field(rejFields).$parent.$el).show();
                        fd.field(rejFields).required = true;
                    }
                    else {
                        fd.field(rejFields).required = false;
                        $(fd.field(rejFields).$parent.$el).hide();
                    }
                });
             }

            HideFields(hFields, true);

            var SummaryPlainText = fd.field('SummaryPlainText').value;
            if(SummaryPlainText !== null && SummaryPlainText !== undefined && SummaryPlainText !== ''){
              fd.field('Summary').value = SummaryPlainText;

            var CARSummaryPlainText = fd.field('CARSummaryPlainText').value;
            if(CARSummaryPlainText !== null && CARSummaryPlainText !== undefined && CARSummaryPlainText !== '')
              fd.field('CARSummary').value = CARSummaryPlainText;

              if(disableCode)
               aurFields.push('Code');
              await disableFields(aurFields, true, false);
              await disablPeoplePickerFields(true);
            }
        }
    }
}
//#endregion

//#region AUS RENDER MODULE (AUDIT SCHEDULE)
var onAUSRender = async function() {
    if (_isNew || _isEdit){

      let result = await getGridMType(_web, _webUrl, _module, false);
      if (result.colArray.length > 0) {
        _colArray = result.colArray;
        _targetList = result.targetList;
      }

      _spComponentLoader.loadScript(_htLibraryUrl).then(renderHandsonTable);
    }
}
//#endregion



//#region DCC FUNCTIONS
var getProjectList = async function(ProjectYear) 
{ 
    var SERVICE_URL = _siteUrl+ "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=getAllProjectsbyear&ProjectYear="+encodeURIComponent(ProjectYear);
    var xhr = new XMLHttpRequest();	

	xhr.open("GET", SERVICE_URL, true);	
	
	xhr.onreadystatechange = async function() 
	{
		if (xhr.readyState == 4) 
		{ 	
			var response;
			try 
			{
				if (xhr.status == 200)
				{                        
						const obj =  await JSON.parse(this.responseText, async function (key, value) {
						var _columnName = key;
						var _value = value;					

						if(_columnName === 'ProjectCode'){
							if(!projectArr.includes(_value))							
						       projectArr.push(_value);						
						}						
					});				

                    // for(var i =0; i < 20000; i++){
                    //     projectArr.push(i);
                    // }
					fd.field('Reference').widget.setDataSource({data: projectArr});                    			
				}				
			}
			catch(err) 
			{
				console.log(err + "\n" + text);				
			}
		}
	}
	xhr.send();
} 

function showHideColumn(ColumnName, value)
{      
    if(value.toLowerCase() !== "no")
    {
        fd.field(ColumnName).required = false;
        $(fd.field(ColumnName).$parent.$el).hide();
    }    
    else
    {
        $(fd.field(ColumnName).$parent.$el).show();
        fd.field(ColumnName).required = true;
    }
}

function showHideCustomColumn(ColumnName, value)
{      
    if(value !== '' && value != null && value.toLowerCase() === "no")
    {
        $(fd.field(ColumnName).$parent.$el).show();
        fd.field(ColumnName).required = true;
    }    
    else
    {
        fd.field(ColumnName).required = false;
        $(fd.field(ColumnName).$parent.$el).hide();
    }
}

function renderColumns(field, value){

    let keyField = field.key;
    let keyValue = field.value;
    let isHide = (field.isHide !== null && field.isHide !== undefined) ? field.isHide : false;
    
    if(keyField === 'IsDrawing'){
        //if no add sentence to DrawingAvailability "Drawings were not made available for the Audit)"
        //if yes then hide DrawingAvailability
        // if(_isNew){
           //fd.field(keyField).value = true;
           //fd.field(keyValue).disabled = true;
        // }


            if(value === false){	
                $(fd.field(keyValue).$parent.$el).show();
                fd.field(keyValue).value = 'Drawings were not made available for the Audit';
                fd.field(keyValue).disabled = true;
                $('.Dwg-Grid').hide();
            }
            else{
                fd.field(keyValue).value = '';
                $(fd.field(keyValue).$parent.$el).hide();
                $('.Dwg-Grid').show();
            }
    }

    else if(keyField === 'Revision2' || keyField == '_x0049_DC7'){ //
        debugger;
        let columns = [];
        if(keyValue.includes(','))
          columns = keyValue.split(',');
        else columns.push(keyValue);

        for (const valueField of columns) {
            if(value.toLowerCase() === 'no')
                showHideCustomColumn(valueField, 'no');
            else showHideCustomColumn(valueField, 'yes');
        }
    }

    else if(keyField === '_x0049_DC6'){//All IDC replies evidences are available
        if(value === '' || value.toLowerCase() === 'no') //value.toLowerCase() === 'yes' || 
         showHideCustomColumn(keyValue, 'yes');
        else showHideCustomColumn(keyValue, 'no');
    }

    //IDC is requested
    else showHideColumn(keyValue, value); // if no then hide keyValue
    
    
    if(isHide) 
        showHideCustomColumn(keyValue,  fd.field(keyField).value);
}


var setProjectsWithValidation = async function(){
    fd.field('Reference').required = true;     
    fd.field('Reference').addValidator({
        name: 'Array Count',
        error: 'Only one Project can be selected per form.',
        validate: function(value) {
            if(fd.field('Reference').value.length > 1) {
                return false;
            }
            return true;
        }
    });
                 	
    fd.field('Reference').widget.dataSource.data(['Please wait while retreiving ...']); 
    $("ul.k-reset")
      .find("li")
      .first()
      .css("pointer-events", "none")
      .css("opacity", "0.6"); 
    await getProjectList(ProjectYear);


    //var id = $('div.k-multiselect').first();
    // $(id).on('click', async function(value) {
    //     var countResult = fd.field('Reference').widget.dataSource._data.length;
    //     if(countResult < 2){
    //         fd.field('Reference').widget.dataSource.data(['Please wait while retreiving ...']); 
    //         $("ul.k-reset")
    //         .find("li")
    //         .first()
    //         .css("pointer-events", "none")
    //         .css("opacity", "0.6");  
   
    //         await getProjectList(ProjectYear);
    //     }
    // });

    fd.field('Reference').$on('change', async function(value)
	{
            $(fd.field('Title').$parent.$el).show();		
            fd.field('Title').value = value;
            $(fd.field('Title').$parent.$el).hide();
	});  
}

var setColumnsArrayShowHide = async function(){
    var formFields = [{
        key: 'Registry1',
        value: 'Registry1CH',
        category: 'Drawing Registry',
        question: 'Drawing Registry Is available and includes latest revisions of all phases'
    },
    {
        key: 'Registry2',
        value: 'Registry2CH',
        category: 'Drawing Registry',
        question: 'Drawing Registry shows correct information: (Title/Number/Revision/Date/Names/IDC)'
    },
    {
        key: 'Regsitry3',
        value: 'Registry3CH',
        category: 'Drawing Registry',
        question: 'Drawing Regsitry  shows “Not used” for skipped drawings'
    },

    {
        key: 'Checking1',
        value: 'Checking1CH',
        category: 'Drawing checking',
        question: 'Latest drawings are available in hardcopy'
    },
    {
        key: 'Checking2',
        value: '_x201c_Checking2CH_x201d_',
        category: 'Drawing checking',
        question: 'Drawing number complies with PRC-PM-08 (PRJNUM-PH (if applicable)-C-XNN Rev Z)'
    },
    {
        key: 'Checking3',
        value: 'Checking3CH',
        category: 'Drawing checking',
        question: 'Drawing number complies with PM WI'
    },
    {
        key: 'Checking4',
        value: 'Checking4CH',
        category: 'Drawing checking',
        question: 'PM WI is pre-approved by QM (applicable only in lead trade)'
    },

    
    {
        key: 'Revision1',
        value: 'Revision1CH',
        category: 'Revision Control',
        question: 'Revision history is shown (with initials)'
    },
    {
        key: 'Revision2',
        value: 'Revision2CH,Revision2Yes',
        category: 'Revision Control',
        question: 'Revision is properly described'
    },
    {
        key: 'Revision3',
        value: 'Revision3CH',
        category: 'Revision Control',
        question: 'Revision approval box is signed by GL prior to external issuance or DSR'
    },
    {
        key: 'Revision4',
        value: 'Revision4CH',
        category: 'Revision Control',
        question: 'Clouds & revision codes are used to show changes'
    },


    {
        key: 'TitleBlock1',
        value: 'TitleBlock1CH',
        category: 'Drawing Title Block',
        question: 'Initials are shown in title block'
    },
    {
        key: 'TitleBlock2',
        value: 'TitleBlock2CH',
        category: 'Drawing Title Block',
        question: 'Date in title block (if applicable) is the same as Rev A/Rev 0 date in the revision'
    },
    {
        key: 'TitleBlock3',
        value: 'TitleBlock3CH',
        category: 'Drawing Title Block',
        question: 'Revision in title block (if applicable) should be the latest or as per PM instructions'
    },


    {
        key: 'Signature1',
        value: 'Signature1CH',
        isHide: true,
        category: 'Drawings Signature',
        question: 'Drawings are signed correctly'
    },
    {
        key: 'Stamp1',
        value: 'Stamp1CH',
        category: 'Drawings Signature',
        question: 'Drawings are stamped prior to external submission'
    },
    {
        key: 'Stamp2',
        value: 'Stamp2CH',
        category: 'Drawings Signature',
        question: 'Where drawings are submitted “For Tender” then “For Construction” without introducing any modifications, the same revision is maintained, only stamp is changed'
    },


    {
        key: 'General1',
        value: 'General1CH',
        category: 'General',
        question: 'Revision 0 of drawings of all phases and subsequent revisions are retained for project duration'
    },
    {
        key: 'General2',
        value: 'General2CH',
        category: 'General',
        question: 'Old revisions are stamped superseded'
    },
    {
        key: 'General3',
        value: 'General3CH',
        category: 'General',
        question: 'Drawings are posted under P:\Drive'
    },


    {
        key: 'DrawingsOTE1',
        value: 'DrawingsOTE1CH',
        category: 'Drawings in Language other than English ',
        question: 'PM instructions for issuing drawings in other language is available'
    },


    {
        key: '_x0049_DC1',
        value: 'IDC1CH',
        category: 'Inter-Department Checking (IDC)',
        question: 'IDC is undertaken on alphabetical revision'
    },
    {
        key: '_x0049_DC2',
        value: 'IDC2CH',
        category: 'Inter-Department Checking (IDC)',
        question: 'IDC is undertaken Prior to final DSR or final submission, unless specified otherwise by PM'
    },
    {
        key: '_x0049_DC3',
        value: 'IDC3CH',
        category: 'Inter-Department Checking (IDC)',
        question: 'IDC is undertaken as per minimum requirements for type of project (Refer to IDC Requirements for Drawings document posted under Insidar)'
    },
    {
        key: '_x0049_DC4',
        value: 'IDC4CH',
        isHide: true,
        category: 'Inter-Department Checking (IDC)',
        question: 'IDC is requested'
    }, //value: 'IDC4CH,IDC4YesCH'},
    {
        key: '_x0049_DC5',
        value: 'IDC5CH',
        category: 'Inter-Department Checking (IDC)',
        question: 'Drawings were signed “By” before being sent for IDC OR in PDF format if sent by email'
    },
    {
        key: '_x0049_DC6',
        value: 'IDC6CH',
        category: 'Inter-Department Checking (IDC)',
        question: 'All IDC replies evidences are available'
    },
    {
        key: '_x0049_DC7',
        value: 'IDC7CH',
        category: 'Inter-Department Checking (IDC)',
        question: 'IDC replies received in a timely manner'
    },
    {
        key: '_x0049_DC8',
        value: 'IDC8YesCH',
        isHide: true,
        category: 'Inter-Department Checking (IDC)',
        question: 'Reviewers raised comments (if available)'
    },
    {
        key: '_x0049_DC9',
        value: 'IDC9CH',
        isHide: true,
        category: 'Inter-Department Checking (IDC)',
        question: 'Originator resolved comments correctly'
    },
    {
        key: '_x0049_DC10',
        value: 'IDC10CH',
        category: 'Inter-Department Checking (IDC)',
        question: 'Drawing revision was changed after implementing IDC comments'
    },
    {
        key: '_x0049_DC11',
        value: 'IDC11CH',
        isHide: true,
        category: 'Inter-Department Checking (IDC)',
        question: 'IDC replies are filed properly'
    },


    {
        key: 'Process1',
        value: 'Process1CH',
        isHide: true,
        category: 'Design process on PMIS is BIM or CAD',
        question: 'Design process is defined properly'
    },


    {
        key: 'Authoring1',
        value: 'Authoring1CH',
        isHide: true,
        category: 'Design authoring and development  Authoring1',
        question: 'BIM coordinator (for lead trade only) initiate the BIM model properly'
    },
    {
        key: 'Authoring2',
        value: 'Authoring2CH',
        category: 'Design authoring and development  Authoring1',
        question: 'Trade’s models are published on the shared workspace at the frequency specified in the BEP'
    },
    {
        key: 'Authoring3',
        value: 'Authoring3CH',
        category: 'Design authoring and development  Authoring1',
        question: 'Design teams resolved issues and submitted model check reports/BIMcollab reports to the GL for the SDC process'
    },
    {
        key: 'Authoring4',
        value: 'Authoring4CH',
        isHide: true,
        category: 'Design authoring and development  Authoring1',
        question: 'Group Leader finalize BIM task properly'
    },
    {
        key: 'Authoring5',
        value: 'Authoring5CH',
        category: 'Design authoring and development  Authoring1',
        question: 'SDC is conducted and finalized at the frequency defined in the BEP'
    },


    {
        key: 'Coordination1',
        value: 'Coordination1CH',
        isHide: true,
        category: 'Design Coordination',
        question: 'BIM manager complete coordination task'
    },
    {
        key: 'Coordination2',
        value: 'Coordination2CH',
        category: 'Design Coordination',
        question: 'GLs conducted visual checks and sent  BIMcollab reports to affected trades'
    },
    {
        key: 'Coordination3',
        value: 'Coordination3CH',
        category: 'Design Coordination',
        question: 'GLs clearly define in the BIMcollab reports the reason why raised issues are still pending for the non-resolved issues (if any)'
    },
    {
        key: 'Coordination4',
        value: 'Coordination4CH',
        category: 'Design Coordination',
        question: 'Clash detection/BIMcollab reports are properly signed by GL after IDC process'
    },
    {
        key: 'Coordination5',
        value: 'Coordination5CH',
        category: 'Design Coordination',
        question: 'Clash detection/BIMcollab reports are submitted to Project Manager and BIM Manager'
    },
    {
        key: 'Coordination6',
        value: 'Coordination6CH',
        isHide: true,
        category: 'Design Coordination',
        question: 'Final Clash detection/BIMcollab reports are properly signed and posted on PCW'
    },
    {
        key: 'Coordination7',
        value: 'Coordination7CH',
        category: 'Design Coordination',
        question: 'IDCs are conducted at the frequency defined in the BEP'
    },



    {
        key: 'TRGM1',
        value: 'TRGM1CH',
        isHide: true,
        category: 'Transportation design requirements for building projects including parking',
        question: 'Geometric design requirements checklist is filled for each project phase and package'
    },
    {
        key: 'TRGM2',
        value: 'TRGM2CH',
        isHide: true,
        category: 'Transportation design requirements for building projects including parking',
        question: 'Geometric design requirements checklist is signed properly'
    },
    {
        key: 'TRGM3',
        value: 'TRGM3CH',
        category: 'Transportation design requirements for building projects including parking',
        question: 'Design calculations: Geometric design criteria and Critical point are filled properly'
    },
    {
        key: 'TRGM4',
        value: 'TRGM4CH',
        isHide: true,
        category: 'Transportation design requirements for building projects including parking',
        question: 'Design calculations: Geometric design criteria and Critical point are signed properly'
    },
    {
        key: 'TRGM5',
        value: 'TRGM5CH',
        category: 'Transportation design requirements for building projects including parking',
        question: 'Design calculations: Geometric design criteria and Critical point is attached to the “Geometric design requirements checklist'
    },



    {
        key: 'TRWT1',
        value: 'TRWT1CH',
        category: 'Design water table level requirements',
        question: 'AR Provided the GHCE with (Plan for site boundaries with coordinates / Topographic plan of the site / Proposed number of basements / AR zero level in relation to national datum)'
    },
    {
        key: 'TRWT2',
        value: 'TRWT2CH',
        category: 'Design water table level requirements',
        question: 'GHCE to advise on the clearance required for the support system'
    },
    {
        key: 'TRWT3',
        value: 'TRWT3CH',
        category: 'Design water table level requirements',
        question: 'GHCE to be advised by AR of any variations in final excavation levels'
    },
    {
        key: 'TRWT4',
        value: 'TRWT4CH',
        category: 'Design water table level requirements',
        question: 'Piezometers layout issued by (GHCE) is IDCed as per minimum requirements'
    },
    {
        key: 'TRWT5',
        value: 'TRWT5CH',
        category: 'Design water table level requirements',
        question: 'Excavation issued by (ENV/ST/GHCE/TR) is IDCed as per minimum requirements'
    },
    {
        key: 'TRWT6',
        value: 'TRWT6CH',
        category: 'Design water table level requirements',
        question: 'Rigid pavement layout issued by (GHCE) is IDCed as per minimum requirements'
    },
    {
        key: 'TRWT7',
        value: 'TRWT7CH',
        category: 'Design water table level requirements',
        question: 'Typical sections issued by (TR) is IDCed as per minimum requirements'
    },



    {
        key: 'Translation1',
        value: 'Translation1CH',
        category: 'Translation to Portuguese (if applicable)',
        question: 'GL/PM sent list of drawings and drawing notes (in excel format) to translation coordinator'
    },
    {
        key: 'Translation2',
        value: 'Translation2CH',
        category: 'Translation to Portuguese (if applicable)',
        question: 'In case they were sent by PM, evidence to trades submissions are available'
    },
    {
        key: 'Translation3',
        value: 'Translation3CH',
        category: 'Translation to Portuguese (if applicable)',
        question: 'Received back translated list of drawings and drawing notes and issued them in relevant drawings'
    },


    {
        key: 'Submission1',
        value: 'Submission1CH',
        isHide: true,
        category: 'Submission',
        question: 'Drawings are submitted properly to PM'
    },
    {
        key: 'Submission2',
        value: 'Submission2CH',
        isHide: true,
        category: 'Submission',
        question: 'Drawings are submitted properly to client (applicable only in lead trade)'
    },


    {
        key: 'IsDrawing',
        value: 'DrawingAvailability'
    }
];

    // const QAColumns = ["Registry1", "Registry2", "Regsitry3", "Checking1", "Checking2", "Checking3", "Checking4", "Revision1", "Revision2", "Revision3", //10
    //                    "Revision4", "TitleBlock1", "TitleBlock2", "TitleBlock3",  "Signature1", "Stamp1", "Stamp2", "General1", "General2", "General3", //10 
    //                    "DrawingsOTE1", "_x0049_DC1", "_x0049_DC2", "_x0049_DC3", "_x0049_DC4", "_x0049_DC5", "_x0049_DC6", "_x0049_DC7", "_x0049_DC8", //9

    //                    "_x0049_DC9", "_x0049_DC10", "_x0049_DC11", "Process1", "Authoring1", "Authoring2", "Authoring3", "Authoring4", "Authoring5", 
    //                    "Coordination1", "Coordination2", "Coordination3", "Coordination4", "Coordination5", "Coordination6", "Coordination7", "TRGM1", 
    //                    "TRGM2", "TRGM3", "TRGM4", "TRGM5", "TRWT1", "TRWT2", "TRWT3", "TRWT4", "TRWT5", "TRWT6", "TRWT7", "Translation1", "Translation2", 
    //                    "Translation3", "Submission1", "Submission2", "IsDrawing"];

    // const ANColumns = ["Registry1CH", "Registry2CH", "Registry3CH", "Checking1CH", "_x201c_Checking2CH_x201d_", "Checking3CH", "Checking4CH", "Revision1CH", "Revision2CH", "Revision3CH", //10
    //                    "Revision4CH", "TitleBlock1CH", "TitleBlock2CH", "TitleBlock3CH", "Signature1CH", "Stamp1CH", "Stamp2CH", "General1CH", "General2CH", "General3CH", //10
    //                    "DrawingsOTE1CH", "IDC1CH", "IDC2CH", "IDC3CH", ": IDC4CH", "IDC4YesCH", "IDC5CH", "IDC6CH", "IDC7CH", //9
                       
    //                    "IDC8YesCH", "IDC9CH", "IDC10CH", "IDC11CH", "Process1CH", "Authoring1CH", "Authoring2CH", "Authoring3CH", "Authoring4CH",
    //                    "Authoring5CH", "Coordination1CH", "Coordination2CH", "Coordination3CH", "Coordination4CH", "Coordination5CH", "Coordination6CH", 
    //                    "Coordination7CH", "TRGM1CH", "TRGM2CH", "TRGM3CH", "TRGM4CH", "TRGM5CH", "TRWT1CH", "TRWT2CH", "TRWT3CH", "TRWT4CH", "TRWT5CH", "TRWT6CH", 
    //                    "TRWT7CH", "Translation1CH", "Translation2CH", "Translation3CH", "Submission1CH", "Submission2CH", "DrawingAvailability"];

    //for (var i = 0; i < QAColumns.length; i++)
    for (const field of formFields){		
        let keyField = field.key;
        // let ANcolumn = ANColumns[i];

    	fd.field(keyField).$on('change', function(value)
    	{
            renderColumns(field, value);
    	});	
    
         var value = fd.field(keyField).value;
         renderColumns(field, value);
    }
}
//#endregion

//#region AUR FUNCTIONS
var handleSummaryButton = async function(){

    var department = '';
    if(fd.field('Department').value !== undefined)
     department = fd.field('Department').value.LookupValue;

    var startDate = fd.field('StartAuditDate').value;
    if(startDate !== null){
        var year = startDate.getFullYear();
        var month = startDate.getMonth() + 1; // January is 0, so we add 1 to get the actual month
        var day = startDate.getDate();
        startDate = `${year}-${month}-${day}`;
    }

    var endDate = fd.field('EndAuditDate').value;
    if(endDate !== null){
        var year = endDate.getFullYear();
        var month = endDate.getMonth() + 1; // January is 0, so we add 1 to get the actual month
        var day = endDate.getDate();
        endDate = `${year}-${month}-${day}`;
    }

    var btnAuditElement = $(`button:contains('Generate Audit Summary')`);
    
    $(btnAuditElement).click(async function(){
        preloader();
        
        var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GET_AUR_SUMMARY'; //https://db-sp.darbeirut.com
        var soapContent;
        soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                        '<soap:Body>' +
                          '<GET_AUR_SUMMARY xmlns="http://tempuri.org/">' +
                            '<WebURL>' + _webUrl + '</WebURL>' +
                            '<Department>' + department + '</Department>' +
                            '<StartDate>' + startDate + '</StartDate>' +
                            '<EndDate>' + endDate + '</EndDate>' +
                          '</GET_AUR_SUMMARY>' +
                        '</soap:Body>' +
                      '</soap:Envelope>';
         await getSoapRequest('POST', serviceUrl, false, soapContent)
         .then(async function(){
            var filterFields = ['Department', 'StartAuditDate', 'EndAuditDate'];
            var _operation = 'Add';
             if(_main.includes("No result")){
                await disableFields(filterFields, false, false);
                _operation = 'remove';
             }
             else await disableFields(filterFields, true, false);

             if(_operation === 'Add'){
                filterFields.forEach(function (element) {
                    if (!aurFields.includes(element)) {
                        aurFields.push(element);
                    }
             });
            }
            else{
                aurFields = aurFields.filter(function(field) {
                    return !filterFields.includes(field);
                });
            }

            await disableFields(aurFields, true, false);
            await disablPeoplePickerFields(true);
         })
         .then(function(){
            setButtonState('Print', true);
            _timeOut1 = setInterval(Remove_Preloader, 1000);
         });
        //preloader("remove");
    });

    $('ul.nav-tabs li a').on('click', function(element) {
             debugger;
             var index = $(this).parent().index()
             renderTabs(index);

            activeTabName = $(this).text();
            if(activeTabName == auditReportTab){
                if(department === null || department === undefined || startDate === null || endDate === null){
                    setErrorMesg(btnAuditElement, false, auditGenSummary);
                    $(btnAuditElement).prop('disabled', true);
                }
                else{
                    setErrorMesg(btnAuditElement, true, '');
                    $(btnAuditElement).prop('disabled', false);
                } 
                var summary = fd.field('Summary').value;
                if(summary === null || summary === undefined || summary === '')
                 setButtonState('Print', false);
                else setButtonState('Print', true);

                var status = fd.field('Status').value;
                if(status === 'Closed')
                 setButtonState('Submit', false);
                disableSummary();
            }
            else  handleCarTab();
    });
}

var handleAurTabs = async function(){
    var code = fd.field('Code').value;
    var status = fd.field('Status').value;
    var isSubmitted = fd.field('Submit').value;
    var office = fd.field('Office').value.LookupValue;
    var officeItem = await getAssignedUsers(office);

    if(status === 'Closed'){
        //#region SET AUDIT REPORT TAB READ ONLY AND ENABLE CAR TAB FOR AUDIT TEAM
        debugger;
        var _isAllowed = await ensureFunction('IsUserInGroup', auditTeamGroup);
        if(!_isAllowed)
          $("ul.nav,nav-tabs").remove();
        else {
            disableSummary();
          _isAurClosed = true;
        }
        //#endregion
    }
    else if(status === 'Sent for Review' && (code === '' || code === 'Rejected') ){
        if(isSubmitted){
          _ignoreBtnCreation = true;
          disableSummary();
        }
 
        var assignedOfficeUser = officeItem.AssignedTo;
        var _isAllowedUser = false;
        for (var i = 0; i < assignedOfficeUser.length; i++) {
            var username = assignedOfficeUser[i].Title;
            _isAllowedUser = await doesUserAllowed(username);
            if(_isAllowedUser)
                break;
        }
        if(_isAllowedUser){
            _hideCode = false;
            _isSentForReview = true;

            if(isSubmitted)
            _ignoreBtnCreation = false;
        }
    }
    else if(status === 'Sent for Approval'){
        disableSummary();
        var FinalApproverUser = officeItem.FinalApprover;
        var _isAllowedUser = false;
        for (var i = 0; i < FinalApproverUser.length; i++) {
            var username = FinalApproverUser[i].Title;
            _isAllowedUser = await doesUserAllowed(username);
            if(_isAllowedUser) break;
        }
        if(_isAllowedUser){
         _hideCode = false;
         _isSentForApproval = true;
        }
        else _ignoreBtnCreation = true
    }
    else {
        _ignoreBtnCreation = true;
        //$(`button:contains('Generate Audit Summary')`).remove();
        disableSummary();
        $("ul.nav,nav-tabs").remove();
    }

    if((status === 'Sent for Review' || status === 'Sent for Approval') && code === 'Rejected')
      disableCode = true;
}

var getAssignedUsers = async function(office){
	let result = null;
      await pnp.sp.web.lists
			.getByTitle("Offices")
			.items
			.select("Title,AssignedTo/Title,FinalApprover/Title")
            .expand("AssignedTo,FinalApprover")
			.filter(`Title eq '${  office  }'`)
			.get()
			.then((items) => {
				if(items.length > 0)
				  result = items[0];
				});
	 return result;
}

var handleCarTab = async function(){
    setButtonState('Submit', true);
    
    var btnCarElement = $(`button:contains('Generate CAR Summary')`);
    $(btnCarElement).click(async function(){
       await setSoapResult();
    });

    var contentData = fd.field('Summary').value;
    if(contentData.startsWith('<div')){
        var startIndex = contentData.indexOf('<table');
        contentData = contentData.substring(startIndex);
    }

    var elements = $.parseHTML(contentData);
    var bodyContent = $(elements).filter('#body1');
    fd.field('DescriptionCar').value = bodyContent[0].innerHTML;
    fd.field('DescriptionCar').disabled = true;
    delay(500);

    var disableBtn = false;
    var _fields = ['CarSerialNo', 'DrawingControl', 'DescriptionCar', 'CausesOfNC', 'CorrectiveActionTaken', 'ScheduledCompletionDate', 'Comments', 'FollowupDate'];
    for(var i = 0; i < _fields.length; i++){
        let _field = _fields[i];
        fd.field(_field).required = true
        var _value = fd.field(_field).value;
        if(_value === undefined || _value === null || _value === '')
            disableBtn = true;

        fd.field(_field).$on('change', async function(value){
                disableBtn = await checkCarFields(_fields);
                if(disableBtn){
                    setErrorMesg(btnCarElement, false, carGenSummary);
                    $(btnCarElement).prop('disabled', true);
                    setButtonState('Print', false);
                }
                else{
                    setErrorMesg(btnCarElement, true, '');
                    $(btnCarElement).prop('disabled', false);
                    setButtonState('Print', true);
                } 
        });
    }

    if(disableBtn){
        setErrorMesg(btnCarElement, false, carGenSummary);
        $(btnCarElement).prop('disabled', true);
        setButtonState('Print', false);
    }
    else{
        setErrorMesg(btnCarElement, true, '');
        $(btnCarElement).prop('disabled', false);
        setButtonState('Print', true);
    } 
    disableSummary();
    _distimeOut = setInterval(adjustDisableOpacity, 1000);
}

var setSoapResult = async function(){
    preloader();
    var CurrentUser = await GetCurrentUser();
    var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GET_EMAIL_BODY';
    var soapContent;
    soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body>' +
                      '<GET_EMAIL_BODY xmlns="http://tempuri.org/">' +
                        '<WebURL>' + _webUrl + '</WebURL>' +
                        '<Email_Name>CAR_FORM</Email_Name>' +
                        '<Id>' + fd.itemId + '</Id>' +
                        '<UserDisplayName>' + CurrentUser + '</UserDisplayName>' +
                      '</GET_EMAIL_BODY>' +
                    '</soap:Body>' +
                  '</soap:Envelope>';
     await getSoapRequest('POST', serviceUrl, false, soapContent)
     .then(async function(){
        if(_module === 'AUR' && activeTabName === corrActionTab){
           var fields = ['CarSerialNo', 'DrawingControl', 'DescriptionCar', 'CausesOfNC', 'CorrectiveActionTaken', 'ScheduledCompletionDate', 'Comments', 'FollowupDate'];
           await disableFields(fields, true, false);
        }
     })
     .then(function(){
        setButtonState('Print', true);
        _timeOut1 = setInterval(Remove_Preloader, 1000);
     });
}

//#endregion

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
      fd.field('Status').value;
    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";

    if(!_ignoreBtnCreation)
    {
        await customButtons("Accept", "Submit", false);

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
              if(_module === 'MAP' && _isNew){
                await submitMAPFunction();
              }
              fd.save();
           }
       }
    });
}

fd.spSaved(function(result) {         
    try
    {   
        if(_isNew && _module === 'AUR' ){
            var itemId = result.Id;                     
            var webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
            result.RedirectUrl = webUrl + "/SitePages/PlumsailForms/AuditReport/Item/EditForm.aspx?item=" + itemId;   
        }                   
    }
    catch(e){alert(e);}                              
 });

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
    setButtonToolTip('Submit', submitMesg);
    setButtonToolTip('Close', cancelMesg);
  }

function delay(time) {
    return new Promise(function(resolve) { 
        setTimeout(resolve, time)
    });
}

var renderControls = async function(){
    await setFormHeaderTitle();

    if(_module === 'MAP')
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

const disableFields = async function(fields, disableControls, disableCustomButtons){

    var isAllowed = true;
    for(let i = 0; i < fields.length; i++){
        var fieldName = fields[i];

        if(fieldName == "Attachments")
        {	
           if(_isLead && !isAllowed)
            $(fd.field('Attachments').$parent.$el).hide();
          else if(!isAllowed)
          {		
            $("div.k-upload-sync").removeClass("k-state-disabled");
             if(taskStatus !== 'Open'){
               $('div.k-upload-button').remove();
               $('button.k-upload-action').remove();
               $('.k-dropzone').remove();
            }
          }
          else {
              fd.field('Attachments').disabled = false;			   
              //if(_status == "Completed" || _status == "Closed" || _status == "Issued to Contractor")
              await customButtons("Save", "Save Attachment", true, "U");
          }
        }
        
        else {
            try { 
                fd.field(fieldName).disabled = disableControls; 
            }	
            catch(e){alert(`fields[i] = ${  fieldName  }<br/>${  e}`);}
        }
    }

    if(disableCustomButtons)
    {
       $('span').filter(function(){ return $(this).text() == 'Save'; }).parent().attr("disabled", "disabled");
       $('span').filter(function(){ return $(this).text() == 'Submit'; }).parent().attr("disabled", "disabled");
    }
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

// const disablPeoplePickerFields = async function(fields, disableControls){
//     var userFields = $('div.fd-field-user div');
//     if(userFields.length > 0){
//       userFields.addClass('k-state-disabled')
//       clearInterval(_pplTimeOut);
//     }
// }

const disablPeoplePickerFields = async function(isDisable){
     for(let i = 0; i < aurPeopleFields.length; i++){
         var fieldName = aurPeopleFields[i];
         fd.field(fieldName).ready(function(field) {
            field.disabled = isDisable;
         });
     }
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


var loadScripts = async function(){
    const libraryUrls = [
      _htLibraryUrl,
      //_layout + '/controls/handsonTable/libs/chosen.jquery.js',
      //_layout + '/controls/handsonTable/libs/handsontable-chosen-editor.js',
      _layout + '/controls/preloader/jquery.dim-background.min.js',
      _layout + "/plumsail/js/buttons.js",
      _layout + '/plumsail/js/customMessages.js',
      _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
      _layout + '/plumsail/js/preloader.js',
      _layout + '/plumsail/js/commonUtils.js',
      _layout + '/plumsail/js/utilities.js',
      _layout + '/plumsail/js/partTable.js',
      _layout + '/plumsail/js/grid/gridUtils.js',
      _layout + '/plumsail/js/grid/grid.js'
    ];

    if(_module === 'MAP'){
        Array.prototype.push.apply(libraryUrls, [
            _layout + '/controls/jqwidgets/jqxcore.js',
            _layout + '/controls/jqwidgets/jqxdatetimeinput.js',
            _layout + '/controls/jqwidgets/jqxcalendar.js',
            _layout + '/controls/jqwidgets/jqxtooltip.js',
            _layout + '/controls/jqwidgets/globalization/globalize.js'
        ]);
    }
  
  
    const cacheBusting = `?v=${Date.now()}`;
      libraryUrls.map(url => { 
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
        });
        
    const stylesheetUrls = [
      _layout + '/controls/tooltipster/tooltipster.css',
      _layout + '/controls/handsonTable/libs/handsontable.full.min.css',
      _layout + '/plumsail/css/CssStyle.css',
      _layout + '/plumsail/css/partTable.css',
      _layout + '/controls/jqwidgets/styles/jqx.base.css',

      //_layout + '/controls/jqwidgets/styles/jqx.fluent.scss'
    ];
  
    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
    });
}

var getSoapRequest = async function(method, serviceUrl, isAsync, soapContent){
	var xhr = new XMLHttpRequest(); 
    xhr.open(method, serviceUrl, isAsync); 
    xhr.onreadystatechange = async function() 
    {
        if (xhr.readyState == 4) 
        {   
            try 
            {
                if (xhr.status == 200)
                {                
					const obj = this.responseText;
					var xmlDoc = $.parseXML(this.responseText),
					xml = $(xmlDoc);
					
                    var value= xml.find("GLOBAL_PARAMResult");
                    if(value.length > 0){
                        text = value.text();
                        _layout = value[0].children[0].textContent;
                    }
                }            
            }
            catch(err) 
            {
                console.log(err + "\n" + text);             
            }
        }
    }
	xhr.setRequestHeader('Content-Type', 'text/xml');
    xhr.send(soapContent);

}

var getGlobalParameters = async function(relativeLayoutPath, moduleName, formType){
    if($('.text-muted').length > 0)
      $('.text-muted').remove();
    
    _module = moduleName;
    _formType = formType;
    _web = pnp.sp.web;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl; //https://db-sp.darbeirut.com/sites/QM-eAudit
    _siteUrl = new URL(_webUrl).origin;
    _layout = relativeLayoutPath;

    // var script = document.createElement("script");
    // script.src = _layout + "/plumsail/js/config/configFileRouting.js";
    // document.head.appendChild(script);

    await loadScripts();
    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

    if(_module === 'DCC' || _module === 'AUR') 
      activeTabName = $('a.active').text();

   if(_formType === 'New'){
    clearStoragedFields(fd.spForm.fields);
    _isNew = true;
    if(_module === 'DCC'){
        if(_isNew){
          fd.field('IsDrawing').value = true;
          $(fd.field('DrawingAvailability').$parent.$el).hide();
        }
    }
   }
   else if(_formType === 'Edit')
    _isEdit = true;

    // var serviceUrl = _siteUrl + '/SoapServices/DarPSUtils.asmx?op=GLOBAL_PARAM';
    // var soapContent;
    // soapContent = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    //                 '<soap:Body>' +
    //                   '<GLOBAL_PARAM xmlns="http://tempuri.org/" />' +
    //                 '</soap:Body>' +
    //               '</soap:Envelope>';
    // await getSoapRequest('POST', serviceUrl, false, soapContent);

}

var ensureFunction = async function(funcName, ...params){
    var isValid = false;
    var retry = 1;
    while (!isValid)
    {
        try{
          if(retry >= retryTime) break;
          if(funcName === 'IsUserInGroup'){
            var allowed = await IsUserInGroup(...params);
            isValid = true;
             return allowed;
          }
        }
        catch{
          retry++;
          await delay(delayTime);
        }
    }
}

const renderHandsonTable = (Handsontable) => {
    _container = document.getElementById('dt');

    const newColumns = [];
    const months = ['Jan', 'Feb', 'March', 'April', 'May', 'June', 'July', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
    months.forEach(month => {
        newColumns.push({
            title: month,
            data: month,
            type: 'dropdown',
            //source: ['Audit Planned', 'Audit Under Taken', 'Report Issued', 'Major items closed', 'Completed'],
            source: ['AP', 'AUT', 'RI', 'MIC', 'COM'],
            width: '3%'
        });
    });
   _colArray = _colArray.concat(newColumns);

    _hot = new Handsontable(_container, {
        startRows: 15,
         //data: _data,
         columns: _colArray,
         width:'99%',
         height: '400',
         search: {
           searchResultClass: 'highlight-cell'
         },
         // filters: true,
         // filter_action_bar: true,
         //dropdownMenu: ['filter_by_condition', 'filter_by_value', 'filter_action_bar'],
         //fixedColumnsLeft: _editableColumns.length,
         rowHeaders: true,
         colHeaders: true,
         manualColumnResize: true,
         manualRowMove: true,
         stretchH: 'all',
         licenseKey: htLicenseKey
    });

    addFooterLegend();
}

function addFooterLegend(){
    let fileUrl = `${_webUrl}${_layout}/Images`;
    let border = `style = 'border-collapse: collapse; border: black 1px solid; text-align:center; color: black'`;
    var TradesLegend = `<table cellspacing='0' cellpadding='0' width='30%' border='1' style='border-collapse:collapse; font-size:12px;'>
                         <tr>
                          <td colspan='4' style='border-collapse: collapse; border: black 1px solid; background-color:#6fc7aa; font-weight:bold;text-align:center; font-size:13px'>Audit Legend</td>
                         </tr>

                         <tr>
                           <td ${border}>AP <img src = '${fileUrl}/ri.png' /></td>
                           <td ${border}>Audit Planned</td>

                           <td ${border}>AUT <img src = '${fileUrl}/ri.png' /></td>
                           <td ${border}>Audit Under Taken</td>
                        </tr>

                        <tr>
                          <td ${border}>RI <img src = '${fileUrl}/ri.png' /></td>
                          <td ${border}>Report Issued</td>

                          <td ${border}>MIC <img src = '${fileUrl}/ri.png' /></td>
                          <td ${border}>Major items closed</td>
                        </tr>

                        <tr>
                          <td ${border}>COM <img src = '${fileUrl}/com.png' /></td>
                          <td colspan='3' ${border}>Completed</td>
                        </tr>
                     </table><br/>`;

        //  "<tr><td style='text-align:center'>AR</td><td>Architectural</td><td style='text-align:center'>ME</td><td>Mechanical</td></tr>" +
		//  "<tr><td style='text-align:center'>Area</td><td>Area Office</td><td style='text-align:center'>PCS</td><td>Project Control Specialist</td></tr>" +
		//  "<tr><td style='text-align:center'>CA</td><td>Contract Administrator</td><td style='text-align:center'>PM</td><td>Project Manager</td></tr>" +
		//  "<tr><td style='text-align:center'>Client</td><td>Client</td><td style='text-align:center'>PMC</td><td>PMC General</td></tr>" +
		//  "<tr><td style='text-align:center'>CP</td><td>Construction Specialist</td><td style='text-align:center'>PUD</td><td>Planning and Urban Design</td></tr>" +
		//  "<tr><td style='text-align:center'>EC</td><td>Economics</td><td style='text-align:center'>QS</td><td>Quantity Surveyor</td></tr>" +
		//  "<tr><td style='text-align:center'>EL</td><td>Electrical</td><td style='text-align:center'>SB</td><td>Structural</td></tr>" +
		//  "<tr><td style='text-align:center'>GE</td><td>Geotechnical</td><td style='text-align:center'>TR</td><td>Transportation</td></tr>" +
		//  "<tr><td style='text-align:center'>LAD</td><td>Landscape</td><td style='text-align:center'>WE</td><td>Water and Environment</td></tr>" +
         
	$('#dt').after(TradesLegend);
}



