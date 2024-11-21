var onDCCRender = async function() {
    $('div.SPCanvas, .commandBarWrapper').css('padding-left', '0px');

    if (_isNew || _isEdit) {
        hFields = ['Signature1Yes', 'IDC6Yes', 'IDC9YesCH', 'IDC11YesCH','Title', 'Department2', 'Department3'];
        HideFields(hFields, true);

        if(_isNew){
            setProjectsWithValidation();

            let masterID = JSON.parse(localStorage.getItem('MasterID'))
            let mpID = JSON.parse(localStorage.getItem('MPID'))
       
            fd.field('MasterID').value = masterID
            fd.field('MPID').value = mpID

            $(fd.field('MasterID').$parent.$el).hide();
            $(fd.field('MPID').$parent.$el).hide();
        }
        else
        {
            $(fd.field('Title').$parent.$el).show();
            fd.field('Title').disabled = true;
            fd.field('PKnumber').disabled = true;
            fd.field('Department1').disabled = true;
            fd.field('DepUnits').disabled = true;
            
            let proj = [];
            proj.push(fd.field('Title').value);
            fd.field('Reference').options = proj

            // fd.field('Reference').ready(function(field){
            //     fd.field('Reference').value = proj
            // });
            //fd.field('Reference').widget.dataSource.data(proj);
            
            fd.field('Reference').disabled = true;
            $(fd.field('Title').$parent.$el).hide();

            setTimeout(function () {
                let isDrawing = fd.field('IsDrawing').value
                if(!isDrawing)
                  $('.Dwg-Grid').hide();
            },500);
        }           

        setColumnsArrayShowHide();
    } 
}

//#region DCC FUNCTIONS
var getProjectList = async function(ProjectYear) 
{ 
    if(ProjectYear === undefined)
        ProjectYear = fields.Year.field.value
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

                fd.field('IDC6CH').required = false
                $('.Dwg-Grid').show();
            }
    }

    else if(keyField === 'Revision2' || keyField == '_x0049_DC7'){ //

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

    fd.field('Reference').$on('change', async function(value){

        let query = `Title eq '${value}'`
        await isItemExist(DrawingControl, query, 'Project is already exist').then(()=>{
            checkForAlertBody()
        });

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

function checkForAlertBody() {
    let checkInterval = setInterval(function() {
        let alertBody = $('.alert-body'); // Searching for the .alert-body element

        if (alertBody.length) { // If the .alert-body is found
            console.log('The alert-body has been found!');
            
            // Add padding-left of 20px to the .alert-body
            alertBody.css('padding-left', '100px');
            
            // Clear the interval to stop checking
            clearInterval(checkInterval);
        }
    }, 500); // Check every 500 milliseconds
}
//#endregion
