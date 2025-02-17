var _layout = '/_layouts/15/PCW/General/EForms', 
    _ImageUrl = _spPageContextInfo.webAbsoluteUrl + '/Style%20Library/tooltip.png',
    _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5],
    _ListFullUrl = _spPageContextInfo.webAbsoluteUrl + '/Lists/' + _ListInternalName;

let Inputelems = document.querySelectorAll('input[type="text"]');

var _modulename = "", _formType = "", EmployeeID;
const itemsToRemove = ['Status', 'State', 'Code', 'WorkflowStatus'];
const blueColor = '#6ca9d5', greenColor = '#5FC9B3', yellowColor = '#F7D46D', redColor = '#F28B82';
let _status = '';
let CurrentUser;
let _proceed = false;

var onRender = async function (moduleName, formType){    
	try { 
        const startTime = performance.now();
		_modulename = moduleName;
		_formType = formType;
		if(moduleName == 'EmpStatus')
			await onEmpStatusRender(formType);

        const endTime = performance.now();
        const elapsedTime = endTime - startTime;
        console.log(`Execution time onEmpStatusRender: ${elapsedTime} milliseconds`);
	}
	catch (e) {
		//alert(e);
		console.log(e);	
	}
}

var onEmpStatusRender = async function (formType){	

    await PreloaderScripts(); 			
	await loadScripts();
 
    clearLocalStorageItemsByField(itemsToRemove);

    CurrentUser = await GetCurrentUser();

	if(formType == 'New'){        
		await EmpStatus_newForm();        
    } 	
    else if(formType == 'Edit'){        
		await EmpStatus_editForm();       
    }    
    else if(formType == 'Display'){      
		await EmpStatus_displayForm();     
    }    
}

var EmpStatus_newForm = async function(){ 
    
    try {
    
        fd.toolbar.buttons[0].style = "display: none;";
        fd.toolbar.buttons[1].style = "display: none;";    
            
        CustomclearStoragedFields(fd.spForm.fields);     
        
        const hideField = (field) => $(fd.field(field).$parent.$el).hide();
        const disableField = (field) => fd.field(field).disabled = true; 
        ['Title', 'EmployeeLoginName', 'EmployeeName', 'EmployeeEmail', 'EmployeeOffice', 'EmployeeTitle', 'EmployeeDepartment'].forEach(disableField);  
        
        await loadingButtons();
        formatingButtonsBar();    
        
        _status = fd.field('AccountStatus').value;

        fd.field('AccountStatus').$on('change', async function(value){         
            _status = value;            
        });
        
        fd.field('Employee').ready().then(() => {            
            fd.field('Employee').$on('change', async function(value){ 
                debugger;               
                if(value)
                {                     
                    let LoginName = value.Key.split('|')[1];
                    let DisplayName = value.DisplayText;
                    let EmpEmail = value.EntityData.Email;
                    let EmployeeOffice = LoginName.split("\\")[0].toUpperCase();
                    let EmployeeTitle = value.EntityData.Title;
                    let EmployeeDepartment = value.EntityData.Department;
                    let MobilePhone = value.EntityData.MobilePhone;

                    await doLoadFunction(LoginName); 
                    fd.field('EmployeeLoginName').value = LoginName;
                    fd.field('EmployeeName').value = DisplayName;
                    fd.field('EmployeeEmail').value = EmpEmail;
                    fd.field('EmployeeOffice').value = EmployeeOffice;
                    fd.field('EmployeeTitle').value = EmployeeTitle;
                    fd.field('EmployeeDepartment').value = EmployeeDepartment;  
                    fd.field('ContactNumber').value = ContactNumber;                                      
                }
                else {
                    fd.field('Title').value = '';
                    fd.field('EmployeeLoginName').value = '';
                    fd.field('EmployeeName').value = '';
                    fd.field('EmployeeEmail').value = '';
                    fd.field('EmployeeOffice').value = '';
                    fd.field('EmployeeTitle').value = '';
                    fd.field('EmployeeDepartment').value = '';
                    fd.field('ContactNumber').value = '';  
                }                
            });            
        });         
    }    
    catch(err){
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, LoginName, DisplayName, err.message, err.stack);        
    }      
    
	preloader("remove");
}

var EmpStatus_editForm = async function(){

    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";   
   
    const disableField = (field) => fd.field(field).disabled = true; 
    ['Title', 'EmployeeLoginName', 'EmployeeName', 'EmployeeEmail', 'EmployeeOffice', 'EmployeeTitle', 'EmployeeDepartment'].forEach(disableField);
    
    _status = fd.field('AccountStatus').value;
    if(_status === 'Disabled'){
        ['ContactNumber', 'Site', 'DateofDeparture', 'AccountStatus', 'Comments'].forEach(disableField);
        SetAttachmentToReadOnly();
    }

    fd.field('AccountStatus').$on('change', async function(value){         
        _status = value;            
    });
    
    await loadingButtons(_status);
    formatingButtonsBar();	    
    
	preloader("remove");	
}

var EmpStatus_displayForm = async function(){	

    //fd.toolbar.buttons[1].style = "display: none;";
    fd.toolbar.buttons[1].text = "Cancel";
    fd.toolbar.buttons[1].icon = "Cancel";
    fd.toolbar.buttons[1].style = `background-color:${redColor}; color:white; width:195px !important;`;

    fd.toolbar.buttons[0].icon = "Edit";
    fd.toolbar.buttons[0].text = "Back to Edit Form";
    fd.toolbar.buttons[0].class = 'btn-outline-primary';
    fd.toolbar.buttons[0].style = `background-color:${greenColor}; color:white; width:195px !important;`;

    formatingButtonsBar();     
    preloader("remove");
}

fd.spBeforeSave(function(spForm){					
	return fd._vue.$nextTick();
});

fd.spSaved(async function(result) {			
    try {        
        
        if(_status === 'Enabled' && _proceed){
            const itemId = result.Id;
            const query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
            await _sendEmail(_modulename, 'EmployeesNewEntry_Email', query, '', 'EmployeesStatus_Initiated', '', CurrentUser);   
        }  

    } catch(e) {
        console.log(e);
    }								 
});

//#region Additional Function
function CustomclearStoragedFields(fields){ 

	for (const field in fields) {
        
        if(field !== 'InkSign'){
        
    		var fieldproperties = fd.field(field);

            if(fieldproperties){

                var fieldDefaultVal = fieldproperties._fieldCtx.schema.DefaultValue;
                    
                if (fieldDefaultVal !== undefined && fieldDefaultVal !== null) {}

                else			  
                    fd.field(field).clear();
            }
        } 
	}
}

async function GetItemCount(LoginName) {

    try{
            // const listUrl = `${fd.webUrl}${fd.listUrl}`;
            // const list = await pnp.sp.web.getList(listUrl).get();
            const listTitle = 'CV'; //list.Title;  
                
            const camlFilter = `<View>
                                    <Query>
                                        <Where>									
                                            <Eq><FieldRef Name='LoginName'/><Value Type='Text'>${LoginName}</Value></Eq>						
                                        </Where>
                                    </Query>
                                </View>`; 	
            
            items = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });
            return items;	
        }
        catch(e){
            console.log(e);
        }  
}

async function CapabilitiesensureItemsExist(itemsToEnsure, DisplayName) {
    try {

	    //const listUrl = `${fd.webUrl}` + '/Lists/CVCapabilities';
	    //const list = await pnp.sp.web.getList(listUrl).get();
	    const listTitle = 'CV Capabilities';//list.Title;

        const camlFilter = `<View>
                                <Query>
                                    <Where>									
                                        <Eq><FieldRef Name='LoginName'/><Value Type='Text'>${DisplayName}</Value></Eq>						
                                    </Where>
                                </Query>
                            </View>`; 

        const existingItems = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });
        const existingItemKeys = new Set(existingItems.map(item => `${item.LoginName}|${item.FieldOfSpecialization}`));
        const itemsToCreate = itemsToEnsure.filter(item => !existingItemKeys.has(`${item.LoginName}|${item.FieldOfSpecialization}`));            
    
        if (itemsToCreate.length > 0) {
            const batch = pnp.sp.web.createBatch();

            itemsToCreate.forEach(item => {
                pnp.sp.web.lists.getByTitle(listTitle).items.inBatch(batch).add(item)
                    .then(() => console.log(`Item with Title "${item.Title}" created.`));
            });

            await batch.execute();
        } else {
            console.log("All items already exist.");
        }
    } catch (error) {
        console.log("Error ensuring items exist:", error);
    }
}

async function PreDARensureItemsExist(itemsToEnsure, DisplayName) {
    try {

	    //const listUrl = `${fd.webUrl}` + '/Lists/ProjectExperience';
	    //const list = await pnp.sp.web.getList(listUrl).get();
	    const listTitle = 'CV Previous Experience';//list.Title;
        
        const camlFilter = `<View>
                                <Query>
                                    <Where>									
                                        <Eq><FieldRef Name='LoginName'/><Value Type='Text'>${DisplayName}</Value></Eq>						
                                    </Where>
                                </Query>
                            </View>`;

        const existingItems = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });

        const existingItemKeys = new Set(existingItems.map(item => `${item.LoginName}|${item.OriginalTitle}`));
        const itemsToCreate = itemsToEnsure.filter(item => !existingItemKeys.has(`${item.LoginName}|${item.OriginalTitle}`));            
    
        if (itemsToCreate.length > 0) {
            const batch = pnp.sp.web.createBatch();

            itemsToCreate.forEach(item => {
                pnp.sp.web.lists.getByTitle(listTitle).items.inBatch(batch).add(item)
                    .then(() => console.log(`Item with Title "${item.Title}" created.`));
            });

            await batch.execute();
        } else {
            console.log("All items already exist.");
        }
    } catch (error) {
        console.log("Error ensuring items exist:", error);
    }
}

async function Top5ensureItemsExist(itemsToEnsure, LoginName, itemId) {
    let itemValue = '';
    try {	    
	    const listTitle = 'CV Top Five Projects';//list.Title;
        
        const camlFilter = `<View>
                                <Query>
                                    <Where>									
                                        <Eq><FieldRef Name='LoginName'/><Value Type='Text'>${LoginName}</Value></Eq>						
                                    </Where>
                                </Query>
                            </View>`;

        const existingItems = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });

        const existingItemKeys = new Set(existingItems.map(item => `${item.LoginName}|${item.Title}`));
        const itemsToCreate = itemsToEnsure.filter(item => !existingItemKeys.has(`${item.LoginName}|${item.Title}`));            
    
        if (itemsToCreate.length > 0) {
            const batch = pnp.sp.web.createBatch();

            itemsToCreate.forEach(async item => {

                itemValue = item;
                if(itemId !== undefined){

                    const newItemData = {                                                
                        Lookup_IDId: itemId // Use CVOwnerId to specify the user ID
                    };

                    const updatedItemData = {
                        ...item,
                        ...newItemData
                    };

                    await pnp.sp.web.lists.getByTitle(listTitle).items.inBatch(batch).add(updatedItemData)
                    .then(() => console.log(`Item with Title "${item.Title}" created.`));
                }
                else{
                    await pnp.sp.web.lists.getByTitle(listTitle).items.inBatch(batch).add(item)
                        .then(() => console.log(`Item with Title "${item.Title}" created.`));
                }
            });

            await batch.execute();
        } else {
            console.log("All items already exist.");
        }
    } catch (error) {
        console.log(`Item: ${JSON.stringify(itemValue)} Error ensuring items exist:`, error);
    }
}

async function GetCapabilitiesItemCount(DisplayName) {
    try{   
		CapabilitiesitemsCount = 0;
		//const listUrl = `${fd.webUrl}` + '/Lists/CVCapabilities';
	    //const list = await pnp.sp.web.getList(listUrl).get();
	    const listTitle = 'CV Capabilities';//list.Title;
		     
    	var camlF =  `LoginName eq '${DisplayName}'`;             	
        Capabilitiesitems = await pnp.sp.web.lists.getByTitle(listTitle).items.select('LoginName', 'NbofCapabilities').filter(camlF).get();
		if (Capabilitiesitems && Capabilitiesitems.length > 0) {
	        CapabilitiesitemsCount = Capabilitiesitems[0].NbofCapabilities;	        
    	}						 
    }
    catch(e){alert(e);}   
}

async function CheckCapabilityCount(EVal, DisplayName) {   	 	    
    // const ResponsibilityTableArr = fd.control('CapabilitiesTable')._listViewManager._dataSource._pristineData;
    // var SPDataTable1Length = await ResponsibilityTableArr.length;
	// if(EVal !== undefined) 
	// 	SPDataTable1Length += EVal;
	// if(SPDataTable1Length > 0){
	// 	await GetCapabilitiesItemCount(DisplayName);
	// 	if(CapabilitiesitemsCount == SPDataTable1Length)                		
	// 		fd.control('CapabilitiesTable').buttons[0].visible = false;               
        
	// 	else 
	// 	{
	// 		if(EVal !== undefined) {
	// 			if(SPDataTable1Length - EVal > CapabilitiesitemsCount)                
	// 				fd.control('CapabilitiesTable').buttons[0].visible = true;                             
	// 		}
	// 	}			
	// }   
}

async function GetEmpKeyQualFromSP(LoginName, DisplayName){

    let EmployeeKeySharepoint = await GetEmployeeKeyQualificationsFromSharepoint(LoginName);      

    if(EmployeeKeySharepoint.length > 0) {

        let EmployeeKeyitemsToUpdate = []; 
        let ProfessionValue = '';

        EmployeeKeySharepoint.map(item => {
            ProfessionValue = item.HRPosition !== null ? item.HRPosition : item.Profession;
            EmployeeKeyitemsToUpdate.push({
                Title: ProfessionValue,
                Qualifications: item.KeyQualifications,
                FieldOfSpecialization: item.FieldOfSpecialization,
                LoginName: DisplayName					         
            });
        });
        
        localStorage.setItem('Profession', ProfessionValue);
                 
        (async () => {    
            try {

                await CapabilitiesensureItemsExist(EmployeeKeyitemsToUpdate, DisplayName);
                
                fd.control('CapabilitiesTable').ready().then(async function(){         
                    fd.control('CapabilitiesTable').refresh().then(function(){ FixWidget(fd.control('CapabilitiesTable')); });        //FixListTabelRows                             
                });	                
                    
            } catch (error) {
                console.log("Error in getting filtered items:", error);
            }
        })();         
    } 
}

async function GetProfession(LoginName){    
    
    const listTitle = 'CV Capabilities';//list.Title;
        
    const camlFilter = `<View>
                            <ViewFields>
                                <FieldRef Name='Title' />                                                                                       
                            </ViewFields>
                            <Query>
                                <Where>									
                                    <Eq><FieldRef Name='LoginName'/><Value Type='Text'>${LoginName}</Value></Eq>						
                                </Where>
                            </Query>
                        </View>`;

    let EmployeeKeySharepoint = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });

    if(EmployeeKeySharepoint.length > 0) {   
        let ProfessionValue = '';
        EmployeeKeySharepoint.map(item => {
            ProfessionValue = item.Title !== null ? item.Title: '';            
        });        
        localStorage.setItem('Profession', ProfessionValue);                
    } 
    else
        localStorage.removeItem('Profession');
}

async function GetEmpPreDarExpFromSP(LoginName, DisplayName){

    let EmployeePreDarExperiences = await GetEmployeePreDarExperiencesFromSharepoint(LoginName);     
    
    if(EmployeePreDarExperiences.length > 0) {

        let PreDARitemsToUpdate = []; 

        EmployeePreDarExperiences.map(item => {

            let PreviousExperienceExported = item.Description !== null ? item.Description : 'NA';
            let Original = PreviousExperienceExported;
            let Project = item.Project != null ? item.Project : 'NA';

            if (Project !== 'NA') {
                PreviousExperienceExported += '\n' + Project;
                Original += Project;
            }

            PreDARitemsToUpdate.push({
                Title: item.Employer !== null ? item.Employer : 'NA',
                From: item.FromYear !== null ? item.FromYear : 'NA',
                To: item.ToYear !== null ? item.ToYear : 'NA',
                Country: item.Country !== null ? item.Country : 'NA',
                Position: item.Position !== null ? item.Position : 'NA',
                PreviousExperience: PreviousExperienceExported,
                OriginalTitle: Original,         
                LoginName: DisplayName                   				         
            });
        });        
            
        (async () => {    
            try {
    
                await PreDARensureItemsExist(PreDARitemsToUpdate, DisplayName);    
                
                fd.control('PrevExpTable').ready().then(async function(){         
                    fd.control('PrevExpTable').refresh().then(function(){                
                        FixWidget(fd.control('PrevExpTable')); 
                    });                               
                });               
                
            } catch (error) {
                console.log("Error in getting filtered items:", error);
            }
        })();        
    }  
}

async function GetEmployeesTopFiveProjects(LoginName, DisplayName, itemId){

    let GetTopFiveProjects = await GetEmployeesTopFiveProjectsFromSQL('GET', true, LoginName);
    const GetTopFiveProjectsNodes = GetTopFiveProjects.getElementsByTagName("EmplyeeTop5Projects");	     
    
    if(GetTopFiveProjectsNodes.length > 0) {

        let GetTopFiveProjectsToUpdate = []; 

        for (let i = 0; i < GetTopFiveProjectsNodes.length; i++) {

            const projectNode = GetTopFiveProjectsNodes[i];
            
            const ProjectCode = projectNode.getElementsByTagName("ProjectNumber")[0]?.textContent.trim() || 'NA';
            const RegularHours = projectNode.getElementsByTagName("TotalRegularHours")[0]?.textContent || 'NA';          
            const ProjectName = projectNode.getElementsByTagName("ProjectName")[0]?.textContent || 'NA';
            const ProjectCountry = projectNode.getElementsByTagName("Country")[0]?.textContent.trim() || 'NA';
            const ProjectYear = projectNode.getElementsByTagName("ProjectYear")[0]?.textContent || 'NA';
            const ProjectSubCategory = projectNode.getElementsByTagName("SubCategory")[0]?.textContent || 'NA';
            const ProjectCategory = projectNode.getElementsByTagName("ProjectCategory")[0]?.textContent || 'NA';
            const ProjectDescription = projectNode.getElementsByTagName("ProjectDescription")[0]?.textContent || 'NA';
            const MCProjectDescription = projectNode.getElementsByTagName("MCProjectDescription")[0]?.textContent || 'NA';            

            GetTopFiveProjectsToUpdate.push({
                Title: ProjectCode,
                ProjectName: ProjectName,
                TotalHours: RegularHours,              
                ProjectCountry: ProjectCountry,
                ProjectYear: ProjectYear,
                ProjectSubCategory: ProjectSubCategory,
                ProjectCategory: ProjectCategory,         
                ProjectDescription: ProjectDescription,
                MCProjectDescription: MCProjectDescription,
                LoginName: LoginName,
                EmployeeName: DisplayName           				         
            });
        };        
            
        (async () => {    
            try {
    
                await Top5ensureItemsExist(GetTopFiveProjectsToUpdate, LoginName, itemId);    
                
                fd.control('TopFiveTable').ready().then(async function(){ 
                    fd.control('TopFiveTable').buttons[0].visible = false;	
                    fd.control('TopFiveTable').filter = `<Eq><FieldRef Name="LoginName"/><Value Type="Text">${LoginName}</Value></Eq>`;           
                    fd.control('TopFiveTable').refresh().then(function() {
                        setTimeout(function() {
                            FixWidget(fd.control('TopFiveTable')); 
                        }, 200); // Delay of 2000 milliseconds (2 seconds)
                    });                               
                });               
                
            } catch (error) {
                console.log("Error in getting filtered items:", error);
            }
        })();        
    }  
}

async function GetEmpGrade(LoginName){     

    let xmlDoc = await GetEmployeeGrade('GET', true, LoginName);  
   
    const table1Nodes = xmlDoc.getElementsByTagName("EmployeeGrade");
    let EmployeeID = 0;
	if(table1Nodes !== undefined && table1Nodes !== null && table1Nodes.length > 0)	{		
        EmployeeID = table1Nodes[0].getElementsByTagName("EmployeeID")[0].textContent.trim();
        let EmployeeIDNumber = parseInt(EmployeeID, 10);
        EmployeeID = EmployeeIDNumber.toString().padStart(6, '0');
    }  

	fd.field('Title').value = EmployeeID; 

    if(EmployeeID === 0)
        fd.field('Title').disabled = false; 
    else
        fd.field('Title').disabled = true; 
}

async function GetItemCT(LoginName){
    (async () => {    
        try {
            items = await GetItemCount(LoginName);       
            if (items && items.length > 0) {
                alert('You have already filled in your record for the CV guideline. If you need to make any changes or additions, please review your existing submission. Thank you!');
                fd.close();	        
            }            
        } catch (error) {
            console.log("Error in getting filtered items:", error);
        }
    })();
}

async function loadingButtons(_status){    
    
    fd.toolbar.buttons.push({
        icon: 'Accept',
        class: 'btn-outline-primary',
        disabled: false,
        text: 'Submit',
        style: `background-color:${greenColor}; color:white`,
        click: async function() {  	
            if(fd.isValid){
                await PreloaderScripts(); 
                _proceed = true;
                fd.save();
            }            
	     }
    });
    
    fd.toolbar.buttons.push({
        icon: 'Cancel',
        class: 'btn-outline-primary',
        text: 'Cancel',	
        style: `background-color:${redColor}; color:white`,
        click: async function() {
            await PreloaderScripts();
            fd.close();
        }			
	});
    
    var buttons = fd.toolbar.buttons;
    var submitButton = buttons.find(button => button.text === 'Submit');
    if (submitButton)   
        submitButton.disabled = (_status === 'Disabled');        
}

function formatingButtonsBar(){
    
    $('div.ms-compositeHeader').remove()
    $('span.o365cs-nav-brandingText').text('Employees Status');
    $('i.ms-Icon--PDF').remove();
          
    let toolbarElements = document.querySelectorAll('.fd-toolbar-primary-commands');
    toolbarElements.forEach(function(toolbar) {
        toolbar.style.display = "flex";
        toolbar.style.justifyContent = "flex-end";
        //toolbar.style.marginRight = "25px";            
    });

    let commandBarElement = document.querySelectorAll('[aria-label="Command Bar."]');
        commandBarElement.forEach(function(element) {        
        element.style.paddingTop = "16px";       
    }) ;
    
    document.querySelectorAll('.CanvasZoneContainer.CanvasZoneContainer--read').forEach(element => {
        element.style.marginTop = '7px';
        element.style.marginLeft = '-100px';
    });

    var fieldTitleElements = document.querySelectorAll('.fd-form .row > .fd-field-title');

    fieldTitleElements.forEach(function(element) {
        element.style.fontWeight = 'bold';      
        element.style.borderTopLeftRadius = '6px';
        element.style.borderBottomLeftRadius = '6px';
        element.style.width = '200px';
        element.style.display = 'inline-block';
    });
}

function parseGrade(grade) {
    const match = grade.match(/^P(\d+)(\+?)$/);
    if (!match) return null;
    return {
        level: parseInt(match[1], 10),
        plus: match[2] === '+'
    };
}

function ReadOnly(elem){ 
    $(elem).prop("readonly", true).css({
	    "background-color": "transparent",
	    "border": "0px"
	});  
}

function fixTextArea(){
	$("textarea").each(function(index){		
		$(this).css('height', '150px');
		var height = (this.scrollHeight + 5) + "px";
        $(this).css('height', height);
	});
}

function FixListTabelRows(){ 
    
    let tables = $("table[role='grid']");
    tables.each(function(tblIndex, tbl){
        $(tbl).find('tr').each(function(trIndex, tr) {
    	  
    	    if (trIndex === 0 ){    	
    		   let childs = tr.children;
    		   if(childs.length > 0){
    		     childs[0].style.textAlign = 'center';
    			 childs[1].style.textAlign = 'center';
                }                  		   
    		}
    		
    	   $(tr).find('td').each(function(tdIndex, td) {
                let $td = $(td);
                
                if (tdIndex === 0 || tdIndex === 1)
                    td.style.textAlign = 'center';
                    
                else{
                    if(_formType !== 'Display')
                        $td.children().css('whiteSpace', 'nowrap');
                }
                 
                if(_formType !== 'Display')
                    $td.css('whiteSpace', 'nowrap');
    		});                			
        });
    });    
}

function FixWidget(dt){
    FixListTabelRows();
    var Clientwidth = dt.$el.clientWidth; 
    //Clientwidth = Clientwidth * 96 / 100;  
    var Rwidget = dt.widget;
    var columns = Rwidget.columns;  
    var ColumnsLength = columns.length;
    var width = Clientwidth/(ColumnsLength-1);
    
    var RemainingWidth = 0;
    var RemainingWidth2 = 0;
    var RemainingWidth3 = 0;
    var RemainingWidth4 = 0;
    var RemainingWidth5 = 0;
    var RemainingWidth6 = 0;
            
    for (let i = 1; i < ColumnsLength; i++) {
    
        var field = columns[i].field;	
        
        if(field === 'Reviewed'){
            var ReviewedWidth = 80;
            RemainingWidth = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        }  
        else if(field === 'LinkTitle'){
            var ReviewedWidth = 350;
            RemainingWidth2 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(field === 'FieldOfSpecialization'){
            var ReviewedWidth = 180;
            RemainingWidth3 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(field === 'Qualifications'){
            var ReviewedWidth = 400;
            RemainingWidth4 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        }  
        else if(field === 'From'){
            var ReviewedWidth = 90;
            RemainingWidth5 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(field === 'To'){
            var ReviewedWidth = 90;
            RemainingWidth6 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(i == (ColumnsLength - 1)){                  
            dt._columnWidthViewStorage.set(field, (width + RemainingWidth + RemainingWidth2 + RemainingWidth3 + RemainingWidth4 + RemainingWidth5 + RemainingWidth6)); 
            Rwidget.resizeColumn(columns[i], (width + RemainingWidth + RemainingWidth2 + RemainingWidth3 + RemainingWidth4 + RemainingWidth5 + RemainingWidth6)); 
        }
        else{
            dt._columnWidthViewStorage.set(field, width); 
            Rwidget.resizeColumn(columns[i], width); 
        } 
    }

    const gridContent = dt.$el.querySelector('.k-grid-content.k-auto-scrollable');
    if (gridContent) {
        gridContent.style.overflowX = 'hidden';
    }
    
    var rows = Rwidget._data;
    
    // #32DAC4
    for (let i = 0; i < rows.length; i++) {
        var Reviewed = rows[i].Reviewed; 
        if(Reviewed === 'Yes'){
            const row = $(dt.$el).find('tr[data-uid="' + rows[i].uid+ '"');
            row[0].style.background = 'linear-gradient(to right, rgb(245, 255, 245), rgb(235, 250, 245), rgb(220, 255, 250), rgb(210, 220, 255), rgb(215, 205, 255), rgb(215, 205, 255), rgb(205, 205, 255))';
        }           
    }        
}

function RenderForm(MCAdmin) {

    if(MCAdmin === 'User') {
        if(_Status === 'Pending' || _Status === 'Reviewed'){
            
            fd.toolbar.buttons.push({
                icon: 'Save',
                class: 'btn-outline-primary',
                text: 'Save',
                style: `background-color:${blueColor}; color:white; width:150px !important;`,
                click: async function() {                    
                    
                    RedirectRule = 'Redirect';
        			$(fd.field('Status').$parent.$el).show();
        			fd.field('Status').value = 'Pending';
                    $(fd.field('Status').$parent.$el).hide();
        			alert(_SaveAlertMessage);        				
                    if(fd.isValid){
                    	await PreloaderScripts(); 
                    	fd.save();
                    	preloader("remove");
                    }					 
        	     }
            });
    
            fd.toolbar.buttons.push({
                icon: 'Accept',
                class: 'btn-outline-primary',
                text: 'Submit',
                style: `background-color:${greenColor}; color:white; width:150px !important;`,
                click: async function() {                
                    
                    if (confirm('Ready to submit?')) {
                        
                        RedirectRule = 'Dont Redirect';
            			$(fd.field('Status').$parent.$el).show();
            			fd.field('Status').value = 'Submitted';
                        $(fd.field('Status').$parent.$el).hide();			
            			alert(_SubmitAlertMessage); 	
                        if(fd.isValid){
                        	await PreloaderScripts(); 
                        	fd.save();
                        	preloader("remove");
                        }
                    }						 
        	     }
            });    
        }
        else {
            alert('You have already submitted your record for the CV Form. Thank you!');
	    	fd.close();
        }
    }
    else {
        
        if(_Status === 'Submitted'){
            fd.toolbar.buttons.push({
                    icon: 'Save',
                    class: 'btn-outline-primary',
                    text: 'Save',
                    style: `background-color:${blueColor}; color:white; width:150px !important;`,
                    click: async function() {                        
                        
                        RedirectRule = 'Redirect';
            			$(fd.field('Status').$parent.$el).show();
            			fd.field('Status').value = 'Submitted';
                        $(fd.field('Status').$parent.$el).hide();        					
            			alert(_SaveAlertMessage);	
                        if(fd.isValid){
                        	await PreloaderScripts(); 
                        	fd.save();
                        	preloader("remove");
                        }						 
            	     }
                });
        
            fd.toolbar.buttons.push({
                icon: 'Accept',
                class: 'btn-outline-primary',
                disabled: true,
                text: 'Confirm',
                style: `background-color:${greenColor}; color:white; width:150px !important;`,
                click: async function() {	
                    
                    if (confirm('Are you sure you want to confirm the form?')) {
                        
                        RedirectRule = 'Dont Redirect';
                        $(fd.field('Status').$parent.$el).show();
            			fd.field('Status').value = 'Reviewed';
                        $(fd.field('Status').$parent.$el).hide();
                        if(fd.isValid){
                        	await PreloaderScripts();                            
                   
                           var EmailbodyHeader = "<p style='margin: 0 0 10px;'>We are pleased to inform you that your <a href=" + editTaskLink + ">CV</a> has been successfully reviewed by the M&C department.</p>";
                            EmailbodyHeader += "</br>";
                            EmailbodyHeader += "<p style='margin: 0 0 10px;'>You may now review your updated <a href=" + editTaskLink + ">CV</a> and suggest any further updates or changes.</p>";
                            EmailbodyHeader += "</br>";	
                            EmailbodyHeader += "<p style='margin: 0 0 10px;'>If you have any further questions or require additional assistance, please do not hesitate to contact us.</p>";
                            EmailbodyHeader += "</br>";	
                            EmailbodyHeader += "<p style='margin: 0 0 10px;'>Thank you for your attention and cooperation.</p>";
                            EmailbodyHeader += "</br>";	
                            
            				var Subject = 'Your CV Has Been Reviewed by the M&C Department';
            				var encodedSubject = htmlEncode(Subject);
            				var Body = GetHTMLBody(EmailbodyHeader, RealDisplayName, 'M&C Department');
            				var encodedBody = htmlEncode(Body);
                            var notificationName = '';

            				await _sendEmail(_modulename, encodedSubject + '|' + encodedBody, '', CreatedByEmail, notificationName, '');
                            
                        	fd.save();
                        	preloader("remove");
                        }
                    }						 
        	     }
            });
        }
        else if(_Status === 'Pending' || _Status === 'Reviewed') {
            // if(!_isforReviewer && _Impersonate === null){
            //     alert('The CV form has not yet been submitted for your review. Thank you!');
    	    // 	fd.close();
            // }
            //else{
                fd.toolbar.buttons.push({
                    icon: 'Save',
                    class: 'btn-outline-primary',
                    text: 'Save',
                    style: `background-color:${blueColor}; color:white; width:150px !important;`,
                    click: async function() {                        
                        
                        RedirectRule = 'Redirect';
            			$(fd.field('Status').$parent.$el).show();
            			fd.field('Status').value = 'Pending';
                        $(fd.field('Status').$parent.$el).hide();        					
            			alert(_SaveAlertMessage);	
                        if(fd.isValid){
                        	await PreloaderScripts(); 
                        	fd.save();
                        	preloader("remove");
                        }						 
            	     }
                });
        
                fd.toolbar.buttons.push({
                    icon: 'Accept',
                    class: 'btn-outline-primary',
                    disabled: true,
                    text: 'Confirm',
                    style: `background-color:${greenColor}; color:white; width:150px !important;`,
                    click: async function() {	
                        
                        //await PreloaderScripts();           						
            			if (confirm('Are you sure you want to confirm the form?')) {
                            
                            RedirectRule = 'Dont Redirect';
                            $(fd.field('Status').$parent.$el).show();
                			fd.field('Status').value = 'Reviewed';
                            $(fd.field('Status').$parent.$el).hide();
                            if(fd.isValid){
                            
                            	await PreloaderScripts();                                       
                             
                                var EmailbodyHeader = "<p style='margin: 0 0 10px;'>We are pleased to inform you that your <a href=" + editTaskLink + ">CV</a> has been successfully reviewed by the M&C department.</p>";
                                EmailbodyHeader += "</br>";	
                                EmailbodyHeader += "<p style='margin: 0 0 10px;'>You may now review your updated <a href=" + editTaskLink + ">CV</a> and suggest any further updates or changes.</p>";
                                EmailbodyHeader += "</br>";	
                                EmailbodyHeader += "<p style='margin: 0 0 10px;'>If you have any further questions or require additional assistance, please do not hesitate to contact us.</p>";
                                EmailbodyHeader += "</br>";	
                                EmailbodyHeader += "<p style='margin: 0 0 10px;'>Thank you for your attention and cooperation.</p>";
                                EmailbodyHeader += "</br>";	
                                
                				var Subject = 'Your CV Has Been Reviewed by the M&C Department';
                				var encodedSubject = htmlEncode(Subject);
                				var Body = GetHTMLBody(EmailbodyHeader, RealDisplayName, 'M&C Department');
                				var encodedBody = htmlEncode(Body);
                                var notificationName = '';

                				await _sendEmail(_modulename, encodedSubject + '|' + encodedBody, '', CreatedByEmail, notificationName, '');
                                 
                            	fd.save();
                            	preloader("remove");
                            }
                        }                                            						 
            	     }
                });
            //}
        }
        else if(_Status === 'Reviewed') {
            // alert('The CV form has already been reviewed. Thank you!');
	    	// fd.close();
        }
    }    
}

async function CheckifUserinSPGroup() {

	let IsTMUser = "User"; 

	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(async function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {				
					if(groupsData[i].Title === "MC_Reviewer")
					{					
					   IsTMUser = "MC_Reviewer";
                       _DipN = user.Title;
					   break;
				    }					
				}				
			});
	     });
    }
	catch(e){
        console.log(e);
    }
	return IsTMUser;				
}

function GetHTMLBody(EmailbodyHeader, ToName, DoneBy){	

	var Body = "<html>";
	Body += "<head>";
	Body += "<meta charset='UTF-8'>";
	Body += "<meta name='viewport' content='width=device-width, initial-scale=1.0'>";				
	Body += "</head>";
	Body += "<body style='font-family: Verdana, sans-serif; font-size: 12px; line-height: 1.5; color: #333;'>";
	Body += "<div style='max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ccc; border-radius: 5px; background-color: #f9f9f9;'>"; 	
    
    Body += "<div style='margin-top: 10px;'>";
	Body += "<p style='margin: 0 0 10px;'>Dear <strong>" + ToName + "</strong>,</p>";
    Body += '</br>' ;
    Body += EmailbodyHeader ;
    Body += "<p style='margin: 0 0 10px;'>Best regards,</p>";
	Body += "<p style='margin: 0 0 10px;'><strong>" + DoneBy + "</strong></p>";
	Body += "</div>";				
	Body += "</div>";
	Body += "</body>";
	Body += "</html>";

	return Body;
}

function GetCapabilitiesTableCount(dt){
   
    var ReviewedCount = 0;
    var CapabilitiesTableCount = dt._listViewManager._dataSource._pristineData.length;
    
    var Rwidget = dt.widget; 
    var rows = Rwidget._data;
    
    for (let i = 0; i < rows.length; i++) {
        var Reviewed = rows[i].Reviewed; 
        if(Reviewed === 'Yes'){  
            ReviewedCount++;          
        }           
    } 
    
    if(CapabilitiesTableCount ==  ReviewedCount)
        _CapabilitiesTableCount = true;
    else
        _CapabilitiesTableCount = false;
        
    CheckBoolians();
}

function GetCAdditionalCapabilitiesTableCount(dt){
   
    var ReviewedCount = 0;
    var CapabilitiesTableCount = dt._listViewManager._dataSource._pristineData.length;
    
    var Rwidget = dt.widget; 
    var rows = Rwidget._data;
    
    for (let i = 0; i < rows.length; i++) {
        var Reviewed = rows[i].Reviewed; 
        if(Reviewed === 'Yes'){  
            ReviewedCount++;          
        }           
    } 
    
    if(CapabilitiesTableCount ==  ReviewedCount)
        _CAdditionalCapabilitiesTableCount = true;
    else
        _CAdditionalCapabilitiesTableCount = false;
        
    CheckBoolians();
}

function GetCTopFiveTableCount(dt){
   
    var ReviewedCount = 0;
    var CapabilitiesTableCount = dt._listViewManager._dataSource._pristineData.length;
    
    var Rwidget = dt.widget; 
    var rows = Rwidget._data;
    
    for (let i = 0; i < rows.length; i++) {
        var Reviewed = rows[i].Reviewed; 
        if(Reviewed === 'Yes'){  
            ReviewedCount++;          
        }           
    } 
    
    if(CapabilitiesTableCount ==  ReviewedCount)
        _CTopFiveTableCount = true;
    else
        _CTopFiveTableCount = false;
        
    CheckBoolians();
}

function GetCProjectsTableCount(dt){
   
    var ReviewedCount = 0;
    var CapabilitiesTableCount = dt._listViewManager._dataSource._pristineData.length;
    
    var Rwidget = dt.widget; 
    var rows = Rwidget._data;
    
    for (let i = 0; i < rows.length; i++) {
        var Reviewed = rows[i].Reviewed; 
        if(Reviewed === 'Yes'){  
            ReviewedCount++;          
        }           
    } 
    
    if(CapabilitiesTableCount ==  ReviewedCount)
        _CProjectsTableCount = true;
    else
        _CProjectsTableCount = false;
        
    CheckBoolians();
}

function GetCPrevExpTableCount(dt){
   
    var ReviewedCount = 0;
    var CapabilitiesTableCount = dt._listViewManager._dataSource._pristineData.length;

    var rows = dt.widget._data;
    
    rows.forEach(function(row) {
        if (row.Reviewed === 'Yes') {
            ReviewedCount++;
        }
    }); 
    
    if(CapabilitiesTableCount ==  ReviewedCount)
        _CPrevExpTableCount = true;
    else
        _CPrevExpTableCount = false;
        
    CheckBoolians();
}

function CheckBoolians(){   
    
    var buttons = fd.toolbar.buttons;

    buttons.forEach(function(button, index) {
        if (button.text === 'Confirm') {
            
            if (_CapabilitiesTableCount && _CAdditionalCapabilitiesTableCount && _CTopFiveTableCount && _CProjectsTableCount && _CPrevExpTableCount) {
                button.disabled = false;
            } else {
                button.disabled = true;
            }           
        }
    });    
}

function htmlEncode(str) {
    return str.replace(/[&<>"']/g, function(match) {
        return {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        }[match];
    });
}

var loadScripts = async function(){
	const libraryUrls = [		
		_layout + '/controls/tooltipster/jquery.tooltipster.min.js',
		_layout + '/plumsail/js/commonUtils.js'
	];
  
	const cacheBusting = '?t=' + new Date().getTime();
	  libraryUrls.map(url => { 
		  $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
		});
		
	const stylesheetUrls = [
		_layout + '/controls/tooltipster/tooltipster.css',
		_layout + '/plumsail/css/CssStyleCVMain.css'		
	];
  
	stylesheetUrls.map((item) => {
	  var stylesheet = item;
	  $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
	});
}

var PreloaderScripts = async function(){
  
	await _spComponentLoader.loadScript(_layout + '/controls/preloader/jquery.dim-background.min.js')
		.then(() => {
			return _spComponentLoader.loadScript(_layout + '/plumsail/js/preloader.js');
		})
		.then(() => {
			preloader();
		});	    
}

function setPreviewForm(){
    // let targetElement = $('.fd-toolbar-side-commands button');
    // targetElement.removeClass('btn-link').addClass('btn-outline-primary');
   
    // let appendElement = targetElement.children();
    // $("<span data-v-fcdfbf90 class='mx-1'>Preview</span>").insertAfter(appendElement);   
    // $('.fd-toolbar-side-commands').attr('style', 'position: absolute; padding-left: 375px !important');   

    var pdfButton = fd.toolbar.buttons[2];
    pdfButton.text = 'Print';
    pdfButton.location = 0;
    pdfButton.icon = 'Print';  
    pdfButton.style = `color: #444;
                        background-color: #fdfdfd;
                        border: 1px solid #ababab;
                        padding: 5px 24px;
                        display: inline-block;
                        margin: 2px 2px;
                        border-radius: 6px;
                        font-family: Arial, Helvetica, sans-serif;
                        font-size: 0.95em;
                        font-weight: bold;
                        width: max-content;`;  
}

function setTableFilter(controlName, filterField, filterValue) {
    fd.control(controlName).ready(function() {
        fd.control(controlName).filter = `<Eq><FieldRef Name="${filterField}"/><Value Type="Text">${filterValue}</Value></Eq>`;
        fd.control(controlName).refresh().then(function() { 
            setTimeout(function() {
                if(controlName === 'CapabilitiesTable' || controlName === 'TopFiveTable') 
                    fd.control(controlName).buttons[0].visible = false;
                FixWidget(fd.control(controlName));                         
            }, 200); // Optional delay of 2000 milliseconds (2 seconds)
        });
    });
}

function updateFields(DisplayName, LoginName) {
    fd.field('Title').value = DisplayName;	
    var Titleelem = Inputelems[0];
    ReadOnly(Titleelem);

    $(fd.field('LoginName').$parent.$el).show();
    fd.field('LoginName').value = LoginName;
    $(fd.field('LoginName').$parent.$el).hide();

    localStorage.setItem('LoginName', LoginName);
    localStorage.setItem('DisplayName', DisplayName);

    $(fd.field('CVOwner').$parent.$el).show();
    fd.field('CVOwner').value = LoginName;
    $(fd.field('CVOwner').$parent.$el).hide();
}

function resetFields() {
    fd.field('Title').value = 'M&C Team'; // _DipN;
    var Titleelem = Inputelems[0];
    ReadOnly(Titleelem);

    $('.PMRes').hide();
    $('.toHide').hide();
    $('.picImg').hide();
    fd.field('Responsibilities').required = false;
    $(fd.field('Responsibilities').$parent.$el).hide();
}

function applyFilters(filterValue) {
    setTableFilter('CapabilitiesTable', 'LoginName', filterValue);
    setTableFilter('AdditionalCapabilitiesTable', 'LoginName', filterValue);
    setTableFilter('TopFiveTable', 'LoginName', filterValue);
    setTableFilter('ProjectsTable', 'LoginName', filterValue);
    setTableFilter('PrevExpTable', 'LoginName', filterValue);
}

let closePreview = async function() {
    if (previewWindow && previewWindow.closed) {
        previewWindow = null; // Reset previewWindow before closing to prevent infinite loop
        clearInterval(checkPreviewInterval); // Clear any remaining intervals
        Remove_Pre(true);
    }
}

var setButtonToolTip = async function(_btnText, toolTipMessage){  
        
    var btnElement = $('span').filter(function(){ return $(this).text() == _btnText; }).prev();
	if(btnElement.length === 0)
	  btnElement = $(`button:contains('${_btnText}')`);
	
    if(btnElement.length > 0){
	  if(btnElement.length > 1)
		btnElement = btnElement[1].parentElement;
      else btnElement = btnElement[0].parentElement;
	  
      $(btnElement).attr('title', toolTipMessage);

      $(btnElement).tooltipster({
        delay: 100,
        maxWidth: 350,
        speed: 500,
        interactive: true,
        animation: 'fade', //fade, grow, swing, slide, fall
        trigger: 'hover'
      });
    }
}

async function doLoadFunction(LoginName, DisplayName) {   
    startTime = performance.now();
    await GetEmpGrade(LoginName);
    endTime = performance.now();
    elapsedTime = endTime - startTime;
    console.log(`Execution time GetEmployeeGrade: ${elapsedTime} milliseconds`); 	
}

function SetAttachmentToReadOnly(){

	var spanATTDelElement = document.querySelector('.k-upload');
	if(spanATTDelElement !== null)
	{
		//spanATTDelElement.style.display = 'none';
		
		var spanATTUpElement = document.querySelector('.k-upload .k-upload-button');
        if(spanATTUpElement !== null)
            spanATTUpElement.style.display = 'none';
		
		var spanATTZoneElement = document.querySelector('.k-dropzone');
		if(spanATTZoneElement !== null)
			spanATTZoneElement.style.display = 'none';

        var spanRemoveElement = document.querySelector('.k-upload .k-upload-files .k-upload-action');
        if(spanRemoveElement !== null)
            spanRemoveElement.style.display = 'none';
	}	
}
  
//#endregion


//#region Soap Call Function

var GetEmployeeGrade = async function(method, isAsync, LoginName){  

    var siteUrl = _spPageContextInfo.siteAbsoluteUrl;    	
    var serviceUrl = siteUrl + "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeeGrade&LoginName=" + LoginName;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                          <GetEmployeeGrade xmlns="http://tempuri.org/">
                           <LoginName>${LoginName}</LoginName >
                          </GetEmployeeGrade>
                         </soap:Body>
                       </soap:Envelope>`;
                      
	return new Promise((resolve, reject) => {
        var xhr = new XMLHttpRequest();
        xhr.open(method, serviceUrl, isAsync);
        xhr.onreadystatechange = function() {
            if (xhr.readyState == 4) {
                try {
                    if (xhr.status == 200) {
                        const response = this.responseText;
                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(response, "text/xml");
                        resolve(xmlDoc); // Resolve the promise with the parsed XML document
                    } else {
                        reject(new Error('Failed to get valid response'));
                    }
                } catch (err) {
                    reject(err);
                }
            }
        };
        xhr.setRequestHeader('Content-Type', 'text/xml');
        if (soapContent !== '') 
		  xhr.send(soapContent);
        else xhr.send();
    });
}

var GetEmployeeKeyQualifications = async function(method, isAsync, LoginName){    

    // var siteUrl = _spPageContextInfo.siteAbsoluteUrl;    	
    // var serviceUrl = siteUrl + "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeeKeyQualifications&LoginName=" + LoginName;
    
    // let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    //                      <soap:Body>
    //                       <GetEmployeeKeyQualifications xmlns="http://tempuri.org/">
    //                        <LoginName>${LoginName}</LoginName >
    //                       </GetEmployeeKeyQualifications>
    //                      </soap:Body>
    //                    </soap:Envelope>`;
                      
	// return new Promise((resolve, reject) => {
    //     var xhr = new XMLHttpRequest();
    //     xhr.open(method, serviceUrl, isAsync);
    //     xhr.onreadystatechange = function() {
    //         if (xhr.readyState == 4) {
    //             try {
    //                 if (xhr.status == 200) {
    //                     const response = this.responseText;
    //                     const parser = new DOMParser();
    //                     const xmlDoc = parser.parseFromString(response, "text/xml");
    //                     resolve(xmlDoc); // Resolve the promise with the parsed XML document
    //                 } else {
    //                     reject(new Error('Failed to get valid response'));
    //                 }
    //             } catch (err) {
    //                 reject(err);
    //             }
    //         }
    //     };
    //     xhr.setRequestHeader('Content-Type', 'text/xml');
    //     if (soapContent !== '') 
	// 	  xhr.send(soapContent);
    //     else xhr.send();
    // });
}

var GetEmployeeKeyQualificationsFromSharepoint = async function(LoginName){   

    const listTitle = 'GetEmployeeKeyQualifications';//list.Title;
        
    const camlFilter = `<View>
                            <ViewFields>
                                <FieldRef Name='Title' />
                                <FieldRef Name='HRPosition' />
                                <FieldRef Name='KeyQualifications' />
                                <FieldRef Name='FieldOfSpecialization' />  
                                <FieldRef Name='Profession' />                             
                            </ViewFields>
                            <Query>
                                <Where>									
                                    <Eq><FieldRef Name='Title'/><Value Type='Text'>${LoginName}</Value></Eq>						
                                </Where>
                            </Query>
                        </View>`;

    const existingItems = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });
    return existingItems;       
}

var GetEmployeePreDarExperiences = async function(method, isAsync, LoginName){    

    // var siteUrl = _spPageContextInfo.siteAbsoluteUrl;    	
    // var serviceUrl = siteUrl + "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeePreDarExperiences&LoginName=" + LoginName;
    
    // let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    //                      <soap:Body>
    //                       <GetEmployeePreDarExperiences xmlns="http://tempuri.org/">
    //                        <LoginName>${LoginName}</LoginName >
    //                       </GetEmployeePreDarExperiences>
    //                      </soap:Body>
    //                    </soap:Envelope>`;
                      
	// return new Promise((resolve, reject) => {
    //     var xhr = new XMLHttpRequest();
    //     xhr.open(method, serviceUrl, isAsync);
    //     xhr.onreadystatechange = function() {
    //         if (xhr.readyState == 4) {
    //             try {
    //                 if (xhr.status == 200) {
    //                     const response = this.responseText;
    //                     const parser = new DOMParser();
    //                     const xmlDoc = parser.parseFromString(response, "text/xml");
    //                     resolve(xmlDoc); // Resolve the promise with the parsed XML document
    //                 } else {
    //                     reject(new Error('Failed to get valid response'));
    //                 }
    //             } catch (err) {
    //                 reject(err);
    //             }
    //         }
    //     };
    //     xhr.setRequestHeader('Content-Type', 'text/xml');
    //     if (soapContent !== '') 
	// 	  xhr.send(soapContent);
    //     else xhr.send();
    // });
}

var GetEmployeePreDarExperiencesFromSharepoint = async function(LoginName){    

    const listTitle = 'GetEmployeePreDarExperiences';//list.Title;
        
    const camlFilter = `<View>
                            <ViewFields>
                                <FieldRef Name='Title' />   
                                <FieldRef Name='FromYear' />
                                <FieldRef Name='ToYear' />
                                <FieldRef Name='Employer' />
                                <FieldRef Name='Country' />
                                <FieldRef Name='Description' />
                                <FieldRef Name='Project' />
                                <FieldRef Name='Position' />                                                        
                            </ViewFields>
                            <Query>
                                <Where>									
                                    <Eq><FieldRef Name='Title'/><Value Type='Text'>${LoginName}</Value></Eq>						
                                </Where>
                            </Query>
                        </View>`;

    const existingItems = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });
    return existingItems; 
}

var GetEmployeesTopFiveProjectsFromSQL = async function(method, isAsync, LoginName){ 

    const siteUrl = _spPageContextInfo.siteAbsoluteUrl;
    const serviceUrl = `${siteUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeesTopFiveProjects&LoginName=${LoginName}`;        
    const soapContent = `
        <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
            <soap:Body>
                <GetEmployeesTopFiveProjects xmlns="http://tempuri.org/">
                    <LoginName>${LoginName}</LoginName>
                </GetEmployeesTopFiveProjects>
            </soap:Body>
        </soap:Envelope>`;

    return new Promise((resolve, reject) => {
        var xhr = new XMLHttpRequest();
        xhr.open(method, serviceUrl, isAsync);
        xhr.onreadystatechange = function() {
            if (xhr.readyState == 4) {
                try {
                    if (xhr.status == 200) {
                        const response = this.responseText;
                        const parser = new DOMParser();
                        const xmlDoc = parser.parseFromString(response, "text/xml");
                        resolve(xmlDoc); // Resolve the promise with the parsed XML document
                    } else {
                        reject(new Error('Failed to get valid response'));
                    }
                } catch (err) {
                    reject(err);
                }
            }
        };
        
        xhr.setRequestHeader('Content-Type', 'text/xml');
        if (soapContent !== '') 
		  xhr.send(soapContent);
        else xhr.send();
    }); 
}
//#endregion