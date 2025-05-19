var _module, _formType, _web, _webUrl, _siteUrl, _layout, _itemId, _ImageUrl, _ListInternalName, _ProjectNumber, 
    _ListFullUrl, _CurrentUser, _fields = [];
let Inputelems = document.querySelectorAll('input[type="text"]');
let _Email, _Notification = '', _DipN = "", _htLibraryUrl, _errorImg, _submitImg, _WorkflowStatus = '';


//handson variables
let _hot, _container, _data= [], _colArray, batchSize = 15;
const blueColor = '#6ca9d5', greenColor = '#5FC9B3', yellowColor = '#F7D46D', redColor = '#F28B82';

//form type
let _isNew = false, _isEdit = false, _proceed = false;

let categoryArray = '';

const disableField = (field) => fd.field(field).disabled = true;
const enableField = (field) => fd.field(field).disabled = false;
const clearField = (field) => fd.field(field).value = '';
const showField = (field) => $(fd.field(field).$parent.$el).show();
const hideField = (field) => $(fd.field(field).$parent.$el).hide();

var onRender = async function (moduleName, formType, relativeLayoutPath){ 
       
	try { 

        const startTime = performance.now();    

        _module = moduleName;
        _formType = formType;
        _web = pnp.sp.web;
        _webUrl = _spPageContextInfo.siteAbsoluteUrl;
        _layout = '/_layouts/15/PCW/General/EForms';     

        _ImageUrl = _spPageContextInfo.webAbsoluteUrl + '/Style%20Library/tooltip.png',
        _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5],
        _ProjectNumber = _spPageContextInfo.serverRequestPath.split('/')[2],
        _ListFullUrl = _spPageContextInfo.webAbsoluteUrl + '/Lists/' + _ListInternalName,

        _errorImg = _layout + '/Images/Error.png';
        _submitImg = _layout + '/Images/Submitted.png';
        _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js'; 

        if(moduleName == 'FPR')
			await onFPRRender(formType); 		
        
        const endTime = performance.now();
        const elapsedTime = endTime - startTime;
        console.log(`Execution time FPR: ${elapsedTime} milliseconds`);
	}
	catch (e) {
		//alert(e);
		console.log(e);	
	}
}

var onFPRRender = async function (formType){	

    if(formType === "Edit")
        await FPR_editForm();
    else if(formType === "Display")
        await FPR_displayForm();  
}

var FPR_editForm = async function(){ 
    
    try 
    {
        fd.toolbar.buttons[0].style = "display: none;";
        fd.toolbar.buttons[1].style = "display: none;"; 
        
        const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;            
        setIconSource("overview-icon", svguserinfo);

        await loadScripts();
        showPreloader();        
        formatingButtonsBar('Firewall Policies Review');
        
        _CurrentUser = await GetCurrentUser();

        ['IDUnique', 'Title', 'From', 'To', 'Source', 'Destination', 'ChangeRequestID', 'WorkflowStatus', 'AssignTo', 'AssignedBy', 'AssignedDate', 'ResponseDate', 'Status'].forEach(disableField);        

        _WorkflowStatus = fd.field('WorkflowStatus').value;          
   
        await loadingButtons();        

        var buttons = fd.toolbar.buttons;

        if(_WorkflowStatus === 'Open')
        {          
            fd.field('AssignedBy').value = _CurrentUser.Title;
            fd.field('AssignedDate').value = new Date();
            ['AssignTo'].forEach(enableField);
            fd.field('AssignTo').required = true; 

            var submitButton = buttons.find(button => button.text === 'Submit');
            submitButton.text = 'Assign';
            submitButton.icon = 'Assign';            
        }

        else if(_WorkflowStatus === 'Assigned')
        {          
            fd.field('Status').required = true;
            fd.field('ResponseDate').value = new Date(); 
            ['Status'].forEach(enableField);                         
        }
        else {
            ['WorkflowStatus', 'AssignTo', 'AssignedBy', 'AssignedDate', 'ResponseDate'].forEach(hideField);
            var submitButton = buttons.find(button => button.text === 'Submit');
            submitButton.style = "display: none;";
        }        
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);        
    }
    finally{      
        hidePreloader();
    }     
}

var FPR_displayForm = async function(){ 
    
    try 
    {
        const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;            
        setIconSource("overview-icon", svguserinfo);

        await loadScripts();
        showPreloader();        
        formatingButtonsBar('Firewall Policies Review'); 
        
        fd.toolbar.buttons[1].text = "Cancel";
        fd.toolbar.buttons[1].icon = "Cancel";
        fd.toolbar.buttons[1].style = `background-color:${redColor}; color:white; width:195px !important;`;

        fd.toolbar.buttons[0].icon = "Edit";
        fd.toolbar.buttons[0].text = "Edit Form";
        fd.toolbar.buttons[0].class = 'btn-outline-primary';
        fd.toolbar.buttons[0].style = `background-color:${greenColor}; color:white; width:195px !important;`;

        ['WorkflowStatus', 'AssignTo', 'AssignedBy', 'AssignedDate', 'ResponseDate'].forEach(hideField);
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);        
    }
    finally{      
        hidePreloader();
    }     
}

fd.spSaved(async function(result) {	

    try {        
        
        if(_proceed){	            
            debugger;	
            var itemId = result.Id;        
            let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
            await _sendEmail(_module, _Email, query, '', _Notification, '', _CurrentUser);    
        }  

    } catch(e) {
        console.log(e);
    }								 
});

var loadScripts = async function(){
	const libraryUrls = [		
        _layout + '/controls/preloader/jquery.dim-background.min.js',
        _layout + '/plumsail/js/preloader.js',

		_layout + '/controls/tooltipster/jquery.tooltipster.min.js',
        _layout + '/plumsail/js/customMessages.js',
        _layout + '/plumsail/js/utilities.js',
		_layout + '/plumsail/js/commonUtils.js'
	];
  
	const cacheBusting = '?t=' + new Date().getTime();

    libraryUrls.map(url => { 
		  $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
		});
		
	const stylesheetUrls = [
		_layout + '/controls/tooltipster/tooltipster.css',
		_layout + '/plumsail/css/CssStyleCVMain.css',
        _layout + '/plumsail/css/HRDynamics.css',	
        _layout + '/controls/handsonTable/libs/handsontable.full.min.css'
	];
  
	stylesheetUrls.map((item) => {
	  var stylesheet = item;
	  $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}${cacheBusting}">`);
	});
}

const _setData = (Handsontable) => {

    if(_data.length < batchSize){
        var remainingLength = batchSize - _data.length;
        for (var i = 0; i < remainingLength; i++) {
          var rowData = { id: i + 1, value: 'Row ' + (i + 1) }
          _data.push(rowData);
        }
    }

    var contextMenu = ['row_below']; //, '---------', 'remove_row'];
    console.log(_colArray)
    debugger;
     _container = document.getElementById('dt');
	 _hot = new Handsontable(_container, {
		data: _data,
        columns: _colArray,
        width:'100%',
        height: '500',
        autoWrapRow: true,
        autoWrapCol: true,
        //filters: true,
        // filter_action_bar: true,
        // dropdownMenu: ['filter_by_condition', 'filter_by_value', 'filter_action_bar'],
        // rowHeaders: true,
        //colHeaders: true,
        //manualColumnResize: true,
        stretchH: 'all',
        licenseKey: htLicenseKey
	});

    // _hot = new Handsontable(_container, {
    //     data: [
    //       ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1'],
    //       ['A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2'],
    //       ['A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3'],
    //     ],
    //     colHeaders: [
    //       'ID',
    //       'Full name',
    //       'Position',
    //       'Country',
    //       'City',
    //       'Address',
    //       'Zip code',
    //       'Mobile',
    //       'E-mail',
    //     ],
    //     rowHeaders: true,
    //     height: 'auto',
    
    //     licenseKey: htLicenseKey
    //   });

    // setTimeout(() => {
    //     _hot.render();
    // }, 200);
}

var getSchema = async function(){
    var colArray = [];

    var fetchUrl = `${_webUrl}/Config/TR-Schema.txt`
    await fetch(fetchUrl)
        .then(response => response.text())
        .then(async data => {
            colArray = JSON.parse(data); 

        //     for (const obj of colArray) {
        //     if (obj.renderer === "customDropdownRenderer"){
        //         obj.renderer = customDropdownRenderer;
        //     }
        //     if (obj.source === "getDropDownListValues"){
        //         obj.source = await getDropDownListValues(obj.listname, obj.listColumn);
        //         }
        //     else if (obj.source === "getQMDropDownListValues"){
        //         obj.source = await getQMDropDownListValues(obj.listname, obj.listColumn);
        //     }
        // }
    });

    return  colArray
}

async function GetEmployeesInfo(LoginName, isBatch){ 

    let xmlDoc = await GetEmployeesInfoBySeparatedLoginNames('GET', true, LoginName);  
   
    const table1Nodes = xmlDoc.getElementsByTagName("NewDataSet");

	if(table1Nodes !== undefined && table1Nodes !== null && table1Nodes.length > 0)	{
		let DepartmentDesc = table1Nodes[0].getElementsByTagName("DepartmentDesc")[0]?.textContent.trim() || '';                             
        let EmployeeId = table1Nodes[0].getElementsByTagName("EmployeeId")[0]?.textContent.trim() || '';
        let EmployeeIDNumber = parseInt(EmployeeId, 10);
        EmployeeId = EmployeeIDNumber.toString().padStart(6, '0');
        let FullName = table1Nodes[0].getElementsByTagName("FullName")[0]?.textContent.trim() || '';

        if (isBatch) {
            return {
                Title: FullName,
                IDNo: EmployeeId,
                Department: DepartmentDesc,
                Position: "Position"             
            };
        }
        else {
            fd.field('IDNo').value = EmployeeId;
            fd.field('Department').value = DepartmentDesc;
            fd.field('Name').value = FullName;
            _DisableFormFields([fd.field('IDNo'), fd.field('Department'), fd.field('Name')], true);
        } 
    } 
}

var GetEmployeesInfoBySeparatedLoginNames = async function(method, isAsync, LoginName){  

    var siteUrl = _spPageContextInfo.siteAbsoluteUrl;    	
    var serviceUrl = siteUrl + "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeesInfoBySeparatedLoginNames&loginNames=" + LoginName;
    let soapContent = `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                          <GetEmployeesInfoBySeparatedLoginNames xmlns="http://tempuri.org/">
                           <loginNames>${LoginName}</loginNames >
                          </GetEmployeesInfoBySeparatedLoginNames>
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

async function CheckifUserinSPGroup() {

	let IsTMUser = "User"; 

	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(async function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {				
					if(groupsData[i].Title === "TRBulkInsertion")
					{					
					   IsTMUser = "TRBulkInsertion";
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

async function HandleCostValue(addedID, isdelete) {
    let AllItems = fd.control('TrainBulkTable').widget._data;
    let itemIdsArray = AllItems.map(item => item.id);

    // If isdelete is true, exclude the item with addedID from the AllItems array
    if (isdelete && addedID) {
        itemIdsArray = itemIdsArray.filter(id => String(id) !== String(addedID));
    }

    // If AllItems is empty, set cost to 0
    if (!itemIdsArray || itemIdsArray.length === 0) {
        fd.field('Cost').value = 0;
        _DisableFormFields([_fields.Cost], true);
        return; // Exit the function here if AllItems is empty
    }

    // Otherwise, proceed to sum the costs of all items in AllItems
    try {
        const costs = await Promise.all(itemIdsArray.map(item => {
            return pnp.sp.web.lists.getByTitle('Trainee Users').items.getById(item).select("Cost").get().then(i => {
                return parseInt(i.Cost) || 0;
            });
        }));

        // Sum all the costs retrieved from SharePoint
        let totalCost = costs.reduce((sum, cost) => sum + cost, 0);

        // If addedID is not null and isdelete is false, fetch its cost and add it to the total cost
        if (addedID && !isdelete) {
            try {

                const listTitle = 'Trainee Users';//list.Title;
        
                const camlFilter = `<View>
                                        <ViewFields>
                                            <FieldRef Name='Cost' />
                                            <FieldRef Name='Employee' />                                                                       
                                        </ViewFields>
                                        <Query>
                                            <Where>									
                                                <Eq><FieldRef Name='ID' /><Value Type='Counter'>${addedID}</Value></Eq>						
                                            </Where>
                                        </Query>
                                    </View>`;

                const addedItem = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });

                const addedCost = parseInt(addedItem[0].Cost) || 0;
                totalCost += addedCost; // Add the cost of the added item to the total if isdelete is false
                fd.field('Cost').value = totalCost; // Update the field with the total cost
                _DisableFormFields([_fields.Cost], true); // Disable the Cost field if needed

                const employeeId = addedItem[0].EmployeeId; // Get the Employee Id from the expanded lookup field
                const user = await pnp.sp.web.siteUsers.getById(employeeId).get(); // Get user details
                const userLoginName = user.LoginName.split('|')[1];

                console.log('Employee Login Name:', userLoginName); // Log employee login name
                const employeeInfo = await GetEmployeesInfo(userLoginName, true);

                // Now update the record in SharePoint (this can be the added item or another field)
                await pnp.sp.web.lists.getByTitle('Trainee Users').items.getById(addedID)
                    .update({
                        Title: employeeInfo.Title,  // Use employee info to update fields
                        IDNo: employeeInfo.IDNo,
                        Department: employeeInfo.Department,
                        Position: employeeInfo.Position
                    });

                console.log("Record updated successfully!");
                fd.control('TrainBulkTable').refresh(); // Refresh the table after updating

            } catch (error) {
                console.error("Error fetching or updating the added item:", error);
            }
        } else {
            // If addedID is null or isdelete is true, just set the totalCost directly
            fd.field('Cost').value = totalCost;
            _DisableFormFields([_fields.Cost], true);
        }
    } catch (error) {
        console.error("Error fetching costs:", error);
    }
}

function FixListTabelRows(){ 
    
    let tables = $("table[role='grid']");
    tables.each(function(tblIndex, tbl){
        $(tbl).find('tr').each(function(trIndex, tr) {
    	  
    	    if (trIndex === 0){    	
    		   let childs = tr.children;
    		   if(childs.length > 0){
    		     childs[0].style.textAlign = 'center';
    			 childs[1].style.textAlign = 'center';
                 childs[3].style.textAlign = 'center';          
                 childs[5].style.textAlign = 'center';
                }                  		   
    		}
    		
    	   $(tr).find('td').each(function(tdIndex, td) {
                let $td = $(td);
                
                if (tdIndex === 0 || tdIndex === 1 || tdIndex === 3 || tdIndex === 5)
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

        if(field === 'Employee'){
            var ReviewedWidth = 300;
            RemainingWidth = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        }                  
        else if(field === 'Cost'){
            var ReviewedWidth = 100;
            RemainingWidth2 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } //Department, Position
        else if(field === 'IDNo'){
            var ReviewedWidth = 100;
            RemainingWidth3 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        } 
        else if(field === '#commands'){
            var ReviewedWidth = 20;
            RemainingWidth4 = width - ReviewedWidth;            
            dt._columnWidthViewStorage.set(field, ReviewedWidth); 
            Rwidget.resizeColumn(columns[i], ReviewedWidth); 
        }
        else if(field === 'Department'){
            var ReviewedWidth = 300;
            RemainingWidth5 = width - ReviewedWidth;            
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
}

async function loadingButtons(){  

    fd.toolbar.buttons.push({
        icon: 'Accept',
        class: 'btn-outline-primary',
        disabled: false,
        text: 'Submit',
        style: `background-color:${greenColor}; color:white`,
        click: async function() {             

            if(fd.isValid){

                showPreloader(); 

                _proceed = true;           

                if(_WorkflowStatus === 'Open'){                  
                    _Email = 'FPRAssigned_Email'; 
                    fd.field('WorkflowStatus').value = 'Assigned';                   
                }  
                else if(_WorkflowStatus === 'Assigned'){                  
                    _Email = 'FPRReviewed_Email'; 
                    fd.field('WorkflowStatus').value = 'Reviewed';                   
                }               

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
            showPreloader();
            fd.close();
        }			
	});          
}

function formatingButtonsBar(titelValue){
    
    $('div.ms-compositeHeader').remove();
    $('i.ms-Icon--PDF').remove();
          
    let toolbarElements = document.querySelectorAll('.fd-toolbar-primary-commands');
    toolbarElements.forEach(function(toolbar) {
        toolbar.style.display = "flex";
        toolbar.style.justifyContent = "flex-end";                 
    });

    let commandBarElement = document.querySelectorAll('[aria-label="Command Bar."]');
        commandBarElement.forEach(function(element) {        
        element.style.paddingTop = "16px";       
    }) ;     

    var targetElement = document.querySelector(".o365cs-nav-centerAlign");        
    if (targetElement) 
        targetElement.innerHTML = `<strong style='font-size: 16px !important;color: white;font-family: "Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif;'>${titelValue}</strong>`;  
    
    // Select the target element
    var targetElement = document.querySelector(".o365cs-nav-header16 .o365cs-nav-centerAlign");
    if (targetElement) {
        targetElement.style.display = "table-cell";
        targetElement.style.width = "100%";
        targetElement.style.textAlign = "left";
        targetElement.style.verticalAlign = "middle";
    }       

    //$('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');   
    
    $('.fd-form p').css({
        'margin-top': '0',
        'margin-bottom': '1rem',
        'display': 'none'
    });
}

const GetCurrentUser = async function(){
	try {
        const user = await pnp.sp.web.currentUser.get();
        return user;
    } catch (error) {
        console.error("Error fetching current user:", error);
        throw error; // Re-throw the error if needed
    }	
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

function setIconSource(elementId, iconFileName) {

    const iconElement = document.getElementById(elementId);
  
    if (iconElement) {
        iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
    }
  }