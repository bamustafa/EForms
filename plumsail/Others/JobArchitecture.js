var _layout = '/_layouts/15/PCW/General/EForms', 
    _ImageUrl = _spPageContextInfo.webAbsoluteUrl + '/Style%20Library/tooltip.png',
    _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5],
    _ProjectNumber = _spPageContextInfo.serverRequestPath.split('/')[2],
    _ListFullUrl = _spPageContextInfo.webAbsoluteUrl + '/Lists/' + _ListInternalName;
    _webUrl = _spPageContextInfo.webAbsoluteUrl;

let Inputelems = document.querySelectorAll('input[type="text"]');

var _modulename = "", _formType = "";
let fontSize = '17px';

let _Email, _Notification = '', rfCORVal = '', From = '', To = '', _refNo = '', _loginName; 

const itemsToRemove = ['Status', 'State', 'Code', 'WorkflowStatus'];
const blueColor = '#6ca9d5', greenColor = '#5FC9B3', yellowColor = '#F7D46D', redColor = '#F28B82';

let previewWindow = null, checkPreviewInterval = null;

let CurrentUser;
let _proceed = false;

const hideField = (field) => $(fd.field(field).$parent.$el).hide();
const showField = (field) => $(fd.field(field).$parent.$el).show();
const disableField = (field) => fd.field(field).disabled = true;

var onRender = async function (moduleName, formType){    
	try { 
        const startTime = performance.now();
		_modulename = moduleName;
		_formType = formType;

		if(moduleName == 'JobAr')
			await onJobArRender(formType);       

        const endTime = performance.now();
        const elapsedTime = endTime - startTime;
        console.log(`Execution time onJobArRender: ${elapsedTime} milliseconds`);
	}
	catch (e) {
		//alert(e);
		console.log(e);	
	}
}

var onJobArRender = async function (formType){	

    await PreloaderScripts(); 			
	await loadScripts();
 
    clearLocalStorageItemsByField(itemsToRemove);

    CurrentUser = await GetCurrentUser();

	if(formType === 'New'){        
		await JobAr_newForm();        
    } 	
    else if(formType === 'Edit'){        
		await JobAr_editForm();       
    }    
    else if(formType === 'Display'){      
		await JobAr_displayForm();     
    }    
}

var JobAr_newForm = async function(){ 
    
    try {      
     
        fd.toolbar.buttons[0].style = "display: none;";
        fd.toolbar.buttons[1].style = "display: none;";
        
        CustomclearStoragedFields(fd.spForm.fields); 
        
        let realLoginName = '';
        await pnp.sp.web.currentUser.get().then(user => {            
            realLoginName = user.LoginName.split('|')[1];
        }); 
        
        debugger;
        let RetValue = await GetEmployeeFromSharepoint(realLoginName);
        let EmpID = "";
        if(RetValue.length > 0){
            EmpID = RetValue[0].ID;
            window.location.href = `${_spPageContextInfo.webAbsoluteUrl}/SitePages/PlumsailForms/JobArchitecture/Item/EditForm.aspx?item=` + EmpID;
        }
        else{
            alert("Apologies, Form not found. You do not have access to submit this form.");
	    	fd.close();
        }
    }    
    catch(err){
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);        
    }      
    
	preloader("remove");
}

var JobAr_editForm = async function(){

    try {

        fixTextArea();

        const EmployeeName = fd.field('EmployeeName').value; 
        const Title = fd.field('Title').value;
        const Office = fd.field('Office').value;  
        const Department = fd.field('Department').value;
        const EmployeeTitle = fd.field('EmployeeTitle').value;
        const EmployeeGrade = fd.field('EmployeeGrade').value;
        const LoginName = fd.field('LoginName').value;
        const LastAccessedDate = fd.field('LastAccessedDate').value;
        const Acknowledged = fd.field('Acknowledged').value;

        // document.getElementById('ackDiv').style.display = 'none';

        ['EmployeeName', 'Title', 'Office', 'Department', 'EmployeeTitle', 'EmployeeGrade', 'LoginName', 'LastAccessedDate', 'Acknowledged', 'InkSign'].forEach(hideField);

        fd.toolbar.buttons[0].style = "display: none;";
        fd.toolbar.buttons[1].style = "display: none;";         
        
        await loadingButtons();
        formatingButtonsBar(); 
        
        let realLoginName = '';
        await pnp.sp.web.currentUser.get().then(user => {            
            realLoginName = user.LoginName.split('|')[1];
        });

        var isSiteAdmin = _spPageContextInfo.isSiteAdmin; 

        if (LoginName.toLowerCase() !== realLoginName.toLowerCase() && !isSiteAdmin){
            var content = `</br><p>Apologies, this item is not related to you.</p></br>`;
            $('#my-html').append(content);
			alert("Apologies, this item is not related to you.");
	    	fd.close();
		}
        else{

            var content = `</br><p style='font-size: ${fontSize};'>Dear <span style='font-weight: bold;font-size: ${fontSize};'>${EmployeeName},</span></p></br> 

                            <p style='font-size: ${fontSize};'>We have an important update related to the recent implementation of our new Job Architecture (JA) framework.</p></br> 

                            <p style='font-size: ${fontSize};'>As part of this new framework, there have been changes to your role, specifically in your job title and grade. 
                                These adjustments are based on the comprehensive review we have undertaken to ensure our structure is aligned 
                                with both our strategic goals and industry best practices. Please be assured that there will be no changes to your compensation or benefits as a result of this update.</p></br> 

                            <p style='font-size: ${fontSize};'>Your new Grade - <span style='font-weight: bold;font-size: ${fontSize};'>${EmployeeGrade}</span></p>
                            <p style='font-size: ${fontSize};'>Your new Job Title - <span style='font-weight: bold;font-size: ${fontSize};'>${EmployeeTitle}</span></p></br> 

                            <p style='font-size: ${fontSize};'>It is important that you acknowledge and confirm your understanding of these changes, as they form an update to your current terms of employment.</p>
                            <p style='font-size: ${fontSize};'>For any questions regarding your new title and grade, please reach out to your line manager. They will be able to provide further clarity and support.</p></br> 
                        
                            <p style='font-size: ${fontSize};'>Thank you for your continued commitment to Dar, and please do not hesitate to contact us if you need any additional information.</p></br> 
                            
                            <p style='font-size: ${fontSize};'>Best regards,</p>                       
                            <p><span style='font-weight: bold;font-size: ${fontSize};'>${Office}</span>, <span style='font-weight: bold;font-size: ${fontSize};'>${Department}</span></p>`;
            
            $('#my-html').append(content); 

            let imgUrlLegend = `${_webUrl}${_layout}/Images/legend.png`;          
            let imgUrlSettled = `${_webUrl}${_layout}/Images/Settle.png`;
            
            let legendTbl = `<table style="border: 0px solid black;" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td>
                                        <fieldset style="background-color: #ffe9991a;color: #333;padding: 4px;border-radius: 6px;border: 1px solid #ccc;">
                                            <legend style="font-family: Calibri; font-size: 13pt; font-weight: bold;padding: 1px 10px !important;float:none;width:auto;">
                                                <img src="${imgUrlLegend}" style="width: 16px; vertical-align: middle;">
                                                <span style="vertical-align: middle;">Legend</span>
                                            </legend>
                                            <table style="font-size: 11px; font-family: Calibri;" cellspacing="0" cellpadding="1" width="100%" border="0">
                                                <tbody>                                                   
                                                    <tr height="25">
                                                        <td align="center">
                                                            <img title="Issue Settled" src="${imgUrlSettled}" style="height: 14px; width: 14px; border-width: 0px;">
                                                        </td>
                                                        <td>
                                                            <span style="font-family: Calibri; font-size: 11pt;"><span style="font-weight: bold;">Acknowledge:</span> Click the "Acknowledge" button to confirm that you have read the message.</span>
                                                            <span style="margin-right: 10px;"></span>
                                                        </td>
                                                    </tr>   
                                                     <tr height="10">
                                                        <td colspan="2">&nbsp;</td>
                                                    </tr>                                                 
                                                </tbody>
                                            </table>
                                        </fieldset>
                                    </td>
                                </tr>                                                            
                            </table>`;

            $('#jobLegend').append(legendTbl);
        }  
        
        var buttons = fd.toolbar.buttons;
        var submitButton = buttons.find(button => button.text === 'Acknowledge'); 

        if(Acknowledged) {                      
            submitButton.disabled = true;
            submitButton.style = `background-color: #737373; color:white; width:200px !important`;             

            $('#noteDiv').show();
            
            const lastAccessedDate = new Date(`${LastAccessedDate}`);

            const formattedDate = lastAccessedDate.toLocaleDateString('en-GB', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric'
            });
            
            document.getElementById("acknowledgeDate").innerHTML = `<span style="font-weight: bold;">${formattedDate}</span>`;
            $('#noteDiv').parent().append("<br>");           
            $('#my-html').parent().append(`<img id="ackimage" src=${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/PCW/General/EForms/Images/MailSign.png alt="Acknowledge Image" style="display: block; margin-top: 10px;margin-left: -6px;">`);
            $('#my-html').parent().append("<br>");
            $('#jobLegend').hide();
        }
        else{
            submitButton.disabled = false;             
            $('#my-html').parent().append(`<img id="ackimage" src=${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/PCW/General/EForms/Images/MailSign.png alt="Acknowledge Image" style="display: block; margin-top: 10px;margin-left: -6px;">`);
            $('#my-html').parent().append("<br>"); 
            $('#jobLegend').parent().append("<br>");                      
        }   
    }    
    catch(err){
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);        
    } 

	preloader("remove");	
}

var JobAr_displayForm = async function(){	

    try {

        fixTextArea();
        debugger;

        fd.toolbar.buttons[1].text = "Cancel";
        fd.toolbar.buttons[1].icon = "Cancel";
        fd.toolbar.buttons[1].style = `background-color:${redColor}; color:white; width:195px !important;`;

        fd.toolbar.buttons[0].icon = "Edit";
        fd.toolbar.buttons[0].text = "Back to Edit Form";
        fd.toolbar.buttons[0].disabled = true;
        fd.toolbar.buttons[0].class = 'btn-outline-primary';
        fd.toolbar.buttons[0].style = `background-color:${greenColor}; color:white; width:195px !important;`;      

        formatingButtonsBar(); 

        const EmployeeName = fd.field('EmployeeName').value; 
        const Title = fd.field('Title').value;
        const Office = fd.field('Office').value;  
        const Department = fd.field('Department').value;
        const EmployeeTitle = fd.field('EmployeeTitle').value;
        const EmployeeGrade = fd.field('EmployeeGrade').value;
        const LoginName = fd.field('LoginName').value;
        const LastAccessedDate = fd.field('LastAccessedDate').value;
        const Acknowledged = fd.field('Acknowledged').value;

        //document.getElementById('ackDiv').style.display = 'none';

        ['EmployeeName', 'Title', 'Office', 'Department', 'EmployeeTitle', 'EmployeeGrade', 'LoginName', 'LastAccessedDate', 'Acknowledged', 'InkSign'].forEach(hideField);

        let realLoginName = '';
        await pnp.sp.web.currentUser.get().then(user => {            
            realLoginName = user.LoginName.split('|')[1];
        });

        var isSiteAdmin = _spPageContextInfo.isSiteAdmin; 

        if (LoginName.toLowerCase() !== realLoginName.toLowerCase() && !isSiteAdmin){
            var content = `</br><p>Apologies, this item is not related to you.</p></br>`;
            $('#my-html').append(content);
			alert("Apologies, this item is not related to you.");
	    	fd.close();
		}
        else{
            
            var content = `</br><p style='font-size: ${fontSize};'>Dear <span style='font-weight: bold;font-size: ${fontSize};'>${EmployeeName},</span></p></br> 

                            <p style='font-size: ${fontSize};'>We have an important update related to the recent implementation of our new Job Architecture (JA) framework.</p></br> 

                            <p style='font-size: ${fontSize};'>As part of this new framework, there have been changes to your role, specifically in your job title and grade. 
                                These adjustments are based on the comprehensive review we have undertaken to ensure our structure is aligned 
                                with both our strategic goals and industry best practices. Please be assured that there will be no changes to your compensation or benefits as a result of this update.</p></br> 

                            <p style='font-size: ${fontSize};'>Your new Grade - <span style='font-weight: bold;font-size: ${fontSize};'>${EmployeeGrade}</span></p>
                            <p style='font-size: ${fontSize};'>Your new Job Title - <span style='font-weight: bold;font-size: ${fontSize};'>${EmployeeTitle}</span></p></br> 

                            <p style='font-size: ${fontSize};'>It is important that you acknowledge and confirm your understanding of these changes, as they form an update to your current terms of employment.</p>
                            <p style='font-size: ${fontSize};'>For any questions regarding your new title and grade, please reach out to your line manager. They will be able to provide further clarity and support.</p></br> 
                        
                            <p style='font-size: ${fontSize};'>Thank you for your continued commitment to Dar, and please do not hesitate to contact us if you need any additional information.</p></br> 
                            
                            <p style='font-size: ${fontSize};'>Best regards,</p>                       
                            <p><span style='font-weight: bold;font-size: ${fontSize};'>${Office}</span>, <span style='font-weight: bold;font-size: ${fontSize};'>${Department}</span></p>`;
            
            $('#my-html').append(content);             
        }  
        
        var buttons = fd.toolbar.buttons;
        var submitButton = buttons.find(button => button.text === 'Back to Edit Form'); 

        if(Acknowledged) {                      
                      
            //document.getElementById('ackDiv').style.display = 'block';
            //document.getElementById('ackimage').src = `${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/PCW/General/EForms/Images/acknowledge.png`;

            $('#noteDiv').show();

            const lastAccessedDate = new Date(`${LastAccessedDate}`);

            const formattedDate = lastAccessedDate.toLocaleDateString('en-GB', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric'
            });
            //$('#ackdate').append(`<span style='font-weight: bold;'>Date:</span> ${formattedDate}`);

            submitButton.style = `background-color: #737373; color:white; width:200px !important`;

            document.getElementById("acknowledgeDate").innerHTML = `<span style="font-weight: bold;">${formattedDate}</span>`;
            $('#noteDiv').parent().append("<br>");
            $('#my-html').parent().append(`<img id="ackimage" src=${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/PCW/General/EForms/Images/MailSign.png alt="Acknowledge Image" style="display: block; margin-top: 10px;margin-left: -6px;">`);            
            $('#my-html').parent().append("<br>");
        }
        else{
            fd.toolbar.buttons[0].disabled = false;
            //$('#noteDiv').hide();            
            $('#my-html').parent().append(`<img id="ackimage" src=${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/PCW/General/EForms/Images/MailSign.png alt="Acknowledge Image" style="display: block; margin-top: 10px;margin-left: -6px;">`);     
        }
    }    
    catch(err){
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);        
    }
    
    preloader("remove");
}

fd.spBeforeSave(function(spForm){					
	return fd._vue.$nextTick();
});

fd.spSaved(async function(result) {			
    try {         
       
        // if(_formType === 'New' && _modulename === 'COR'){ 
        //     var itemId = result.Id;						
    	// 	var webUrl = `${_spPageContextInfo.webAbsoluteUrl}`; 
        //     var folderStructure = _refNo;
        //     await _CreatFolderStructure('COR', 'CORLibrary', itemId, folderStructure, '');   		
    	// 	result.RedirectUrl = webUrl + "/SitePages/PlumsailForms/COR/Item/EditForm.aspx?item=" + itemId;	
        //     window.location.href = result.RedirectUrl;
        // }        
        // else if(_proceed){		
        //     var itemId = result.Id;        
        //     let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
        //     await _sendEmail(_modulename, _Email, query, '', _Notification, '', CurrentUser);    
        // }  

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

async function loadingButtons(){  

    fd.toolbar.buttons.push({
        icon: 'CheckMark',
        class: 'btn-outline-primary',
        disabled: false,
        text: 'Acknowledge',
        style: `background-color:${greenColor}; color:white; width:200px !important`,
        click: async function() {  	
            if(fd.isValid){
                await PreloaderScripts();          
                fd.field('Acknowledged').value = true;
                fd.field('LastAccessedDate').value = new Date();
                fd.save();
            }            
        }
    });   
    
    fd.toolbar.buttons.push({
        icon: 'Cancel',
        class: 'btn-outline-primary',
        text: 'Cancel',	
        style: `background-color:${redColor}; color:white; width:200px !important`,
        click: async function() {
            await PreloaderScripts();
            fd.close();
        }			
	});           
}

async function processClosedWindowResult() {

    const folderUrl = `${_spPageContextInfo.siteServerRelativeUrl}/CORLibrary/${_refNo}`;    
    try {        
        const files = await pnp.sp.web.getFolderByServerRelativeUrl(folderUrl).files();
        return files.length;
    } catch (error) {
        console.error("Error fetching files:", error);
    }
}

function formatingButtonsBar(){
    
    $('div.ms-compositeHeader').remove()
    $('span.o365cs-nav-brandingText').text(`${_ProjectNumber} - Job Architecture Form`);
    $('i.ms-Icon--PDF').remove();
          
    let toolbarElements = document.querySelectorAll('.fd-toolbar-primary-commands');
    toolbarElements.forEach(function(toolbar) {
        toolbar.style.display = "flex";
        //toolbar.style.justifyContent = "flex-end";
        toolbar.style.marginLeft = "44px";            
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

    document.querySelectorAll(".homepageContainer_ce0ab33e.scrollRegion_ce0ab33e.homepageContainerNewDigestNavBar_ce0ab33e").forEach(element => {
        element.style.textAlign = "justify";
    });

    document.querySelectorAll('.homepageContainer_ce0ab33e.scrollRegion_ce0ab33e.homepageContainerNewDigestNavBar_ce0ab33e').forEach(element => {
        element.style.backgroundColor = "#f4f4f4";
    });

    document.querySelectorAll('.ControlZone').forEach(element => {
        // Uncomment the line below if you want to use the gradient background
        // element.style.background = "linear-gradient(to right, rgb(218, 237, 216), rgb(187, 229, 218), rgb(158, 214, 224), rgb(150, 182, 235), rgb(175, 169, 240), rgb(175, 168, 240), rgb(168, 165, 239))";
                
        element.style.background = "#ffffff";
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
        _layout + '/plumsail/css/CssStyleCV.css',
		_layout + '/plumsail/css/CssStyleCVMain.css'		
	];
  
	stylesheetUrls.map((item) => {
	  var stylesheet = item;
	  $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
	});

    // const fontStyles = `
	// 	@font-face {
	// 		font-family: 'SegoeUIRegular';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/segoeui-regular.woff2') format('woff2'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/segoeui-regular.woff') format('woff');
	// 	}
	// 	@font-face {
	// 		font-family: 'SegoeUILight';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/segoeui-light.woff2') format('woff2'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/segoeui-light.woff') format('woff');
	// 	}
	// 	@font-face {
	// 		font-family: 'FabricIcons53';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.53.woff2') format('woff2'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.53.woff') format('woff'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.53.ttf') format('truetype');
	// 	}
	// 	@font-face {
	// 		font-family: 'FabricIcons23';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.23.woff2') format('woff2'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.23.woff') format('woff'),
	// 			 url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-2.23.ttf') format('truetype');
	// 	}
	// 	@font-face {
	// 		font-family: 'FabricIcons354';
	// 		src: url('${_layout}/plumsail/js/FabricIcons/fabricmdl2icons-3.54.woff') format('woff');
	// 	}
	// `;

	// // Append font styles to head
	// $('<style>')
	// 	.prop('type', 'text/css')
	// 	.html(fontStyles)
	// 	.appendTo('head');
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

async function updateCounter(From, To) {
    
    let refNo = '';
    let value = 1;
    let refFormat = `COR-${From}-${To}`;
    if(From === undefined)
        refFormat = `COR`;
				
	var camlF = "Title" + ` eq '${refFormat}'`;
	
	var listname = 'Counter';
	
	await pnp.sp.web.lists.getByTitle(listname).items.select("Id, Title, Counter").filter(camlF).get().then(async function(items){
		var _cols = { };
        if(items.length == 0){
             _cols["Title"] = `${refFormat}`;
             refNo = `${refFormat}-` + String(1).padStart(5,'0'); 
			 value = parseInt(value) + 1;
             _cols["Counter"] = value.toString();                 
             await pnp.sp.web.lists.getByTitle(listname).items.add(_cols);
                             
        }
          else if(items.length > 0){

            var _item = items[0];
			value = parseInt(_item.Counter) + 1;
            refNo = `${refFormat}-` + String(parseInt(_item.Counter)).padStart(5,'0'); ;             	
			_cols["Counter"] = value.toString();                   
			await pnp.sp.web.lists.getByTitle(listname).items.getById(_item.Id).update(_cols);        			
		}                   
         
    });
    
    return refNo;
}

function AutoReference(items, rfVal) {
	var ReservedRef = "";	
	var CountNumber = 1;	
	if (items.length == 1)    
		CountNumber = items[0].Counter;
		
	CountNumber = String(CountNumber).padStart(5,'0');	
	ReservedRef = `${rfVal}-${CountNumber}`;	
	fd.field('RefNo').value = ReservedRef;
	
	['RefNo'].forEach(disableField);
}

async function referenceCORChecker(Originator, Receiver){

	if(Originator !== '' &&  Receiver !== ''){	

		rfCORVal = `COR-${Originator}-${Receiver}`;
		let camlF = "Title" + " eq '" + rfCORVal + "'";
		
		await pnp.sp.web.lists.getByTitle('Counter').items.select("Title", "Counter").filter(camlF).get().then(function(items){		
			AutoReference(items, rfCORVal);	        
		});
	}
    else{
        rfCORVal = `COR`;
		let camlF = "Title" + " eq '" + rfCORVal + "'";
		
		await pnp.sp.web.lists.getByTitle('Counter').items.select("Title", "Counter").filter(camlF).get().then(function(items){		
			AutoReference(items, rfCORVal);	        
		});
    }	
}

const _CreatFolderStructure = async function(ListName, LibraryName, ID, folderStructure, digitalForm){
	let webUrl = _spPageContextInfo.siteAbsoluteUrl;
	let siteUrl = new URL(webUrl).origin;    
    let serviceUrl = `${siteUrl}/AjaxService/DarPSUtils.asmx?op=UploadAttchment`;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                        <soap:Body>
                            <UploadAttchment xmlns="http://tempuri.org/">
                                <WebURL>${webUrl}</WebURL>
                                <ListName>${ListName}</ListName>                             
                                <LibraryName>${LibraryName}</LibraryName>                                
                                <ID>${ID}</ID>
                                <folderStructure>${folderStructure}</folderStructure>  
								<digitalForm>${digitalForm}</digitalForm>                               
                            </UploadAttchment>
                        </soap:Body>
                    </soap:Envelope>`
    let res = await getSoapResponse('POST', serviceUrl, true, soapContent, 'UploadAttchmentResult');
}

var GetEmployeeFromSharepoint = async function(LoginName){    

    const listTitle = 'Job Architecture';//list.Title;
        
    const camlFilter = `<View>
                            <ViewFields>
                                <FieldRef Name='LoginName' /> 
                                <FieldRef Name='ID' />                                                                                      
                            </ViewFields>
                            <Query>
                                <Where>									
                                    <Eq><FieldRef Name='LoginName'/><Value Type='Text'>${LoginName}</Value></Eq>						
                                </Where>
                            </Query>
                        </View>`;

    const existingItems = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });
    return existingItems; 
}
//#endregion