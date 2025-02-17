var _module, _formType, _web, _webUrl, _siteUrl, _layout, _itemId, _ImageUrl, _ListInternalName, _ProjectNumber, 
    _ListFullUrl, _CurrentUser, _fields = [];
let Inputelems = document.querySelectorAll('input[type="text"]');
let _Email, _Notification = '', _DipN = "", _htLibraryUrl, _errorImg, _submitImg;


//handson variables
let _hot, _container, _data= [], _colArray, batchSize = 15;
const blueColor = '#6ca9d5', greenColor = '#5FC9B3', yellowColor = '#F7D46D', redColor = '#F28B82';

//form type
let _isNew = false, _isEdit = false, _proceed = false;

const disableField = (field) => fd.field(field).disabled = true;
const enableField = (field) => fd.field(field).disabled = false;
const clearField = (field) => fd.field(field).value = '';
const showField = (field) => $(fd.field(field).$parent.$el).show();
const hideField = (field) => $(fd.field(field).$parent.$el).hide();

var onRender = async function (moduleName, formType){ 
       
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

        if(moduleName == 'Bench')
			await onBenchRender(formType); 		
        
        const endTime = performance.now();
        const elapsedTime = endTime - startTime;
        console.log(`Execution time HRDynamics: ${elapsedTime} milliseconds`);
	}
	catch (e) {
		//alert(e);
		console.log(e);	
	}
}

var onBenchRender = async function (formType){	

    if(formType === "New")
        await BM_newForm();  
}

var BM_newForm = async function(){ 
    
    try 
    {
        fd.toolbar.buttons[0].style = "display: none;";
        fd.toolbar.buttons[1].style = "display: none;";        

        await loadScripts();
        showPreloader();
        await loadButtons();
        formatingPagelayout('Mechanical Benchmark Form');
        ['Year'].forEach(hideField);

        await pnp.sp.web.currentUser.get().then(user => {
            DisplayName = user.Title;
            LoginName = user.LoginName.split('|')[1];
        });
        
        let currentYear = new Date().getFullYear();
        let yearArr = [];  

        for (let i = 0; i <= 20; i++) {
            yearArr.push(currentYear - i);
        }
        fd.field('YearChoice').widget.setDataSource({
            data: yearArr
        });

        fd.field('YearChoice').$on('change', async function(value)	{
            if(value){
                ['Year'].forEach(showField);
                fd.field('Year').value = value;
                ['Year'].forEach(hideField);
            }            
        });
        
        fd.field('CreatedBy').value = DisplayName;
        ['CreatedBy', 'ProjectName', 'Country', 'Occupancy'].forEach(disableField);        
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }
    finally {      
        hidePreloader();
    }      
}

var loadScripts = async function(){
	const libraryUrls = [		
        _layout + '/controls/preloader/jquery.dim-background.min.js',
        _layout + '/plumsail/js/preloader.js',

		_layout + '/controls/tooltipster/jquery.tooltipster.min.js',
        _layout + '/plumsail/js/customMessages.js',
        _layout + '/plumsail/js/utilities.js',
		_layout + '/plumsail/js/commonUtils.js',
        _layout + '/plumsail/HRDynamics/utils.js'
	];
  
	const cacheBusting = '?t=' + new Date().getTime();

    libraryUrls.map(url => { 
		  $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
		});
		
	const stylesheetUrls = [
		_layout + '/controls/tooltipster/tooltipster.css',
        _layout + '/plumsail/css/benchmarkStyle.css'     
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

async function loadButtons(status){   
    
    fd.toolbar.buttons.push({
        icon: 'Accept',
        class: 'btn-outline-primary',
        disabled: false,
        text: 'Submit',
        style: `background-color:${greenColor}; color:white`,
        click: async function() {  	
            if(fd.isValid){

                showPreloader(); 

                _proceed = false;

                if(_formType === 'New'){

                    // let refNo = await updateCounter();              
                
                    // $(fd.field('Reference').$parent.$el).show();
                    // fd.field('Reference').value = refNo;
                    // $(fd.field('Reference').$parent.$el).hide();
                    
                    _Email = 'RFCNewEntry_Email';
                    _Notification = 'RFC_Initiated';
                }
                else if(_formType === 'Edit') {
                    
                    // fd.field('Status').disabled = false;
                    // fd.field('Status').value = 'Closed';
                    // fd.field('Status').disabled = true;                    

                    _Email = 'RFCIssued_Email';
                    _Notification = 'RFC_Reviewed';
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
    
    var buttons = fd.toolbar.buttons;
    var submitButton = buttons.find(button => button.text === 'Submit');
    if (submitButton) {  
        if (status === 'Closed') {
            submitButton.disabled = true;          

            disableRichTextField('Answer');
            
            var elem = $("textarea")[0];
	        $(elem).prop("readonly", true);   

            SetAttachmentToReadOnly();
        }
    }         
}

function formatingPagelayout(titelValue){
    
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
    
    const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';
    const linkElement = `<a href="${_spPageContextInfo.webAbsoluteUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">
                            <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">${titelValue}</a>`; 
    $('span.o365cs-nav-brandingText').html(linkElement);

    $('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');   
    
    $('.fd-form p').css({
        'margin-top': '0',
        'margin-bottom': '1rem',
        'display': 'none'
    });

    // Select the <div> element by its role and aria-label attributes    
    const commandBarDiv = document.querySelector('div[role="region"][aria-label="Command Bar."]');

    // Check if the element exists
    if (commandBarDiv) {
        // Change the padding-top
        commandBarDiv.style.paddingTop = "1px"; // Set to desired value
    } else {
        console.error("Element not found.");
    }
}