var _layout = '/_layouts/15/PCW/General/EForms', 
    _siteUrl = _spPageContextInfo.siteAbsoluteUrl,
    _ImageUrl = _siteUrl + '/Style%20Library/tooltip.png',
    _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5],
    _ListFullUrl = _siteUrl + '/Lists/' + _ListInternalName;    

var _modulename = "", _formType = "";
let _proceed = false;

const _listName = 'Tasks';
let itemId = '';

const ProjectNameList = "";
let projectArr=[];
let responseArray=[];
let CurrentUser;

let _Email, _Notification = '';

const itemsToRemove = ['Status', 'State', 'Code', 'WorkflowStatus'];

const blueColor = '#6ca9d5', greenColor = '#5FC9B3', yellowColor = '#F7D46D', redColor = '#F28B82';

var onRender = async function (moduleName, formType){    
	try { 
        const startTime = performance.now();
		_modulename = moduleName;
		_formType = formType;
		if(moduleName == 'LADTask')
			await onLADTaskRender(formType);

        const endTime = performance.now();
        const elapsedTime = endTime - startTime;
        console.log(`Execution time onLADRender: ${elapsedTime} milliseconds`);
	}
	catch (e) {
		//alert(e);
		console.log(e);	
	}
}

var onLADTaskRender = async function (formType){	

    await PreloaderScripts(); 			
	await loadScripts();
 
    clearLocalStorageItemsByField(itemsToRemove);

    CurrentUser = await GetCurrentUser(); 

	if(formType == 'New'){        
		await LADTask_newForm();        
    } 
    else if(formType == 'Edit'){        
		await LADTask_editForm();        
    } 
    else if(formType == 'Display'){        
		await LADTask_displayForm();        
    }     
}

var LADTask_newForm = async function(){    
    
    fd.toolbar.buttons[0].style = "display: none;";   
	fd.toolbar.buttons[1].style = "display: none;";   
	
    await loadingButtons()
    formatingButtonsBar();  
    
    const disableField = (field) => fd.field(field).disabled = true;
    const enableField = (field) => fd.field(field).disabled = false;
    const clearField = (field) => fd.field(field).value = '';
    ['Country', 'Client'].forEach(disableField);
    ['Country', 'Client', 'PM'].forEach(clearField);
    
    fd.field('isAvailable').value = true;
    
    fd.field('isAvailable').$on('change', async function(value)	{
        if(value){
            fd.field('JobNo').required = false;
            $(fd.field('JobNo').$parent.$el).hide(); 
            $(fd.field('Reference').$parent.$el).show(); 
            $(fd.field('YearDropDown').$parent.$el).show(); 
            fd.field('Reference').required = true;
            fd.field('YearDropDown').required = true;
        }
        else {
            $(fd.field('JobNo').$parent.$el).show(); 
            fd.field('JobNo').required = true;
            fd.field('Reference').required = false;
            fd.field('YearDropDown').required = false;
            $(fd.field('Reference').$parent.$el).hide(); 
            $(fd.field('YearDropDown').$parent.$el).hide(); 
            ['Country', 'Client'].forEach(clearField);          
        }
    });	

    let isAvailable = fd.field('isAvailable').value;
    if(isAvailable){
        fd.field('JobNo').required = false;
        $(fd.field('JobNo').$parent.$el).hide(); 
        $(fd.field('Reference').$parent.$el).show(); 
        $(fd.field('YearDropDown').$parent.$el).show(); 
        fd.field('Reference').required = true;
        fd.field('YearDropDown').required = true;
    }
    else {
        $(fd.field('JobNo').$parent.$el).show(); 
        fd.field('JobNo').required = true;
        fd.field('Reference').required = false;
        fd.field('YearDropDown').required = false;
        $(fd.field('Reference').$parent.$el).hide(); 
        $(fd.field('YearDropDown').$parent.$el).hide(); 
        ['Country', 'Client'].forEach(clearField);            
    }

    let currentYear = new Date().getFullYear();
    let yearArr = [];  

    for (let i = 0; i <= 20; i++) {
        yearArr.push(currentYear - i);
    }   

    fd.field('YearDropDown').widget.setDataSource({
        data: yearArr
    });

    let ProjectYear = currentYear;
    fd.field('YearDropDown').value = ProjectYear;     

    fd.field('Reference').widget.dataSource.data(['Please wait while retreiving ...']);
    await GetProjectListALLP(ProjectYear);             		
    
    fd.field('Reference').addValidator({
        name: 'Array Count',
        error: 'Only one JobNo can be selected per form.',
        validate: function(value) {
            if(fd.field('Reference').value && fd.field('Reference').value.length > 1) {
                return false;
            }
            return true;
        }
    });        

    fd.field('YearDropDown').$on('change', async function(value)
	{
        ProjectYear = value;
        
        ['Reference', 'Country', 'Client', 'PM'].forEach(clearField);

        if(ProjectYear){            
            
            ['Reference'].forEach(enableField); 

            fd.field('Reference').widget.dataSource.data(['Please wait while retreiving ...']);
            await GetProjectListALLP(ProjectYear);             		
            
            fd.field('Reference').addValidator({
                name: 'Array Count',
                error: 'Only one JobNo can be selected per form.',
                validate: function(value) {
                    if(fd.field('Reference').value.length > 1) {
                        return false;
                    }
                    return true;
                }
            });
        }                    
        else{  
            ['Reference'].forEach(disableField);             
        }
	});   
	
	fd.field('Reference').$on('change', async function(value)
	{        
        if(value.length > 0)
            await GetProjectList(ProjectYear, value[0]);        
        else{

            $(fd.field('JobNo').$parent.$el).show();	
            fd.field('JobNo').value  = '';
            $(fd.field('JobNo').$parent.$el).hide();
            fd.field('Reference').value  = '';
            
            ['Country', 'Client', 'PM'].forEach(clearField);
        }
	});	    

    preloader("remove");    
}

var LADTask_editForm = async function(){    
    
    fd.toolbar.buttons[0].style = "display: none;";   
	fd.toolbar.buttons[1].style = "display: none;";   
	
    await loadingButtons()
    formatingButtonsBar();        

    preloader("remove");    
}

var LADTask_displayForm = async function(){    
    
    fd.toolbar.buttons[1].text = "Cancel";
    fd.toolbar.buttons[1].icon = "Cancel";
    fd.toolbar.buttons[1].style = `background-color:${redColor}; color:white;`;

    fd.toolbar.buttons[0].icon = "Edit";
    fd.toolbar.buttons[0].text = "Edit";
    fd.toolbar.buttons[0].class = 'btn-outline-primary';
    fd.toolbar.buttons[0].style = `background-color:${greenColor}; color:white;`; 
    
    formatingButtonsBar();

    preloader("remove");    
}

fd.spBeforeSave(function(spForm){					
	return fd._vue.$nextTick();
});

//#region General Functions

async function GetProjectListALLP(ProjectYear) 
{ 
    projectArr = JSON.parse(localStorage.getItem(`${ProjectYear}LADTaskProjects`)) || [];   

    if(projectArr.length > 0){        
        setTimeout(function() {
            fd.field('Reference').widget.setDataSource({
                data: projectArr
            });
        }, 1400);
    }

    var xhr = new XMLHttpRequest();
    var restURL = _siteUrl+ "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=withfilters&ProjectYear="+encodeURIComponent(ProjectYear)+"&ProjectType=all&JobType=PrQ";
    xhr.open("GET", restURL, true);

    xhr.onreadystatechange = async function () {
        if (xhr.readyState == 4) { 
            
            try {

                if(xhr.status == 200) { 

                    const obj = JSON.parse(this.responseText);                                 
                    localStorage.setItem(`${ProjectYear}LADTaskResponseArray`, JSON.stringify(obj));                     
                    const newProjectArr = obj.map(d => d.ProjectCode.trim()); 
                    
                    const diff2 = newProjectArr.filter(x => !projectArr.includes(x));

                    if (diff2.length > 0) {                        
                        localStorage.setItem(`${ProjectYear}LADTaskProjects`, JSON.stringify(newProjectArr));                       
                        fd.field('Reference').widget.setDataSource({
                            data: newProjectArr.sort()
                        });
                    } else {
                        if(projectArr.length === 0){                        
                            fd.field('Reference').widget.setDataSource({
                                data: projectArr
                            });
                        }
                    }
                }
            }  
            catch(err) 
            {
                console.log(err);				
            } 
        }
    } 

    xhr.send();
}

function GetProjectList(ProjectYear, ProjectName) 
{
    responseArray = JSON.parse(localStorage.getItem(`${ProjectYear}LADTaskResponseArray`)) || [];

    if (ProjectName !== '') {
        let projectMetaInfo = responseArray.find(project => project.ProjectCode === ProjectName);
        
        if (projectMetaInfo) {  // Check if a matching project is found
            fd.field('Country').value = projectMetaInfo.AreaName;	
            fd.field('Client').value = projectMetaInfo.ClientName;
            fd.field('PM').value = projectMetaInfo.ProjectManagers;  // Assuming you want the Project Manager's name
            fd.field('JobNoName').value = projectMetaInfo.ProjectName;  

            $(fd.field('JobNo').$parent.$el).show();	
            fd.field('JobNo').value  = projectMetaInfo.ProjectCode;
            $(fd.field('JobNo').$parent.$el).hide();
        }
    } 
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
        //_layout + '/plumsail/css/CssStyleCV.css',
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

function formatingButtonsBar(){   

    $('i.ms-Icon--PDF').remove();
          
    // let toolbarElements = document.querySelectorAll('.fd-toolbar-primary-commands');
    // toolbarElements.forEach(function(toolbar) {
    //     toolbar.style.display = "flex";
    //     toolbar.style.justifyContent = "flex-end";
    //     //toolbar.style.marginRight = "25px";            
    // }); 

    var fieldTitleElements = document.querySelectorAll('.fd-form .row > .fd-field-title');
   
    fieldTitleElements.forEach(function(element) {
        element.style.fontWeight = 'bold'; 
        element.style.fontSize = '14px';      
        element.style.borderTopLeftRadius = '6px';
        element.style.borderBottomLeftRadius = '6px';
        element.style.width = '200px';
        element.style.display = 'inline-block';
    });
}

async function loadingButtons(){	
	
	fd.toolbar.buttons.push({
	        icon: 'Accept',
	        class: 'btn-outline-primary',
	        text: 'Submit',
            style: `background-color:${greenColor}; color:white`,
	        click: async function() {	                             

	        if(fd.isValid){
                await PreloaderScripts();                
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
}

async function updateCounter() {
    
    let refNo = '';
    let value = 1;
				
	var camlF = "Title" + " eq 'LL'";
	
	var listname = 'Counter';
	
	await pnp.sp.web.lists.getByTitle(listname).items.select("Id, Title, Counter").filter(camlF).get().then(async function(items){
		var _cols = { };
        if(items.length == 0){
             _cols["Title"] = 'LL';
             refNo = 'LL-' + String(1).padStart(5,'0'); 
			 value = parseInt(value) + 1;
             _cols["Counter"] = value.toString();                 
             await pnp.sp.web.lists.getByTitle(listname).items.add(_cols);
                             
        }
          else if(items.length > 0){

            var _item = items[0];
			value = parseInt(_item.Counter) + 1;
            refNo = 'LL-' + String(parseInt(_item.Counter)).padStart(5,'0'); ;             	
			_cols["Counter"] = value.toString();                   
			await pnp.sp.web.lists.getByTitle(listname).items.getById(_item.Id).update(_cols);        			
		}                   
         
    });
    
    return refNo;
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

async function GetHTMLBody(Result, EmailbodyHeader, AssignedToLoginName){	
    
    var taskId = await getMaxItemId(_listName);
	var Body = "<html>";
	Body += "<head>";
	Body += "<meta charset='UTF-8'>";
	Body += "<meta name='viewport' content='width=device-width, initial-scale=1.0'>";				
	Body += "</head>";
	Body += "<body style='font-family: Verdana, sans-serif; font-size: 12px; line-height: 1.5; color: #333;'>";
	Body += "<div style='max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ccc; border-radius: 5px; background-color: #f9f9f9;'>"; 
    Body += "<div style='margin-top: 10px;'>";			
	Body += "<p style='margin: 0 0 10px;'>"+ EmailbodyHeader + "</p>";
	Body += "<table style='table style='width:300px;border-collapse:collapse;margin-bottom:10px;'>" +
			"<tr>" +
			"<td style='padding:4px;text-align:center;font-size:12px;font-family:Verdana;'><a style='color:#616A76;font-weight: bold;font-size:13px' href='" + _ListFullUrl + "/EditForm.aspx?ID=" + taskId + "'>Reply</a></td>" +
			"<td style='padding:4px;text-align:center;font-size:12px;font-family:Verdana;'><a style='color:#616A76;font-weight: bold;font-size:13px' href ='" + _ListFullUrl + "/AllItems.aspx'>View All</a></td>" +
			"</tr>" +
			"</table>";
	Body += "<table style='width: 100%; border-collapse: collapse; margin-bottom: 10px;'>";			

	for (var column in Result) {
		Body += "<tr>";
		Body += "<th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;font-size:13px'>" + column + "</th>";
		Body += "<td style='border: 1px solid #ddd; padding: 8px; text-align: left;font-size:13px'>" + Result[column] + "</td>";
		Body += "</tr>";
	}				

	var Attachs = "";
	
	var FullPathAtt =  _ListFullUrl + "/Attachments/" + fd.itemId + "/";
	for (var i = 0; i < fd.field('Attachments').value.length; i++) {

		var attachmentUrl = FullPathAtt + fd.field('Attachments').value[i].name;
		var attachmentName = fd.field('Attachments').value[i].name;
		
		if (Attachs === "") {
			Attachs = "<a href='" + attachmentUrl + "'>" + attachmentName + "</a>";
		} else {
			Attachs += "<br/><a href='" + attachmentUrl + "'>" + attachmentName + "</a>";
		}
	}

	Body += "<tr>";
	Body += "<th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;font-size:13px'>Attachments</th>";
	Body += "<td style='border: 1px solid #ddd; padding: 8px; text-align: left;font-size:13px'>" + Attachs + "</td>";
	Body += "</tr>";
			
	Body += "</table>";	
    Body += "</br>";	
	Body += "<p style='margin: 0 0 10px;'>Best regards,</p>";
	Body += "<p style='margin: 0 0 10px;'>" + AssignedToLoginName + "</p>";
	Body += "</div>";				
	Body += "</div>";
	Body += "</body>";
	Body += "</html>";

	return Body;
}

async function getMaxItemId(listName) {
    try {
      const items = await pnp.sp.web.lists.getByTitle(listName).items.orderBy("ID", true).getAll();
      if (items && items.length > 0) {
          const MaxIDs = items.map(item => item.ID);
          const maxId = Math.max(...MaxIDs);
          return maxId + 1;
      } else {
        return null; // List is empty
      }
    } catch (error) {
      console.error("Error getting max item ID:", error);
      throw error;
    }
}

function MEIsUserInGroup(group, Initiator) 
{
    var IsPMUser = false
	try{
		pnp.sp.web.currentUser.get()
         .then(function(user){
			pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {
				
					if(groupsData[i].Title == group)
					{
						fd.toolbar.buttons.push({
						        icon: 'Save',
						        class: 'btn-outline-primary',
						        text: 'Save',
                                style: `background-color:${blueColor}; color:white`,
						        click: async function() {	                              

								$(fd.field('FormStatus').$parent.$el).show();
								fd.field('FormStatus').value = "Saved By PD";
								$(fd.field('FormStatus').$parent.$el).hide();								
								
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
                                style: `background-color:${greenColor}; color:white`,
						        click: async function() {                            
                                
                                let Code = fd.field('Code').value;                              

                                if (Code === ''){}
                                else {                                                                       

                                    if(Code === 'Approved'){                                        
                                        
                                        $(fd.field('FormStatus').$parent.$el).show();
                                        fd.field('FormStatus').value = "";
                                        $(fd.field('FormStatus').$parent.$el).hide();

                                        $(fd.field('ClosedDate').$parent.$el).show();
                                        fd.field('ClosedDate').value = new Date();
                                        $(fd.field('ClosedDate').$parent.$el).hide();
                                        
                                        fd.field('Status').disabled = false;
                                        fd.field('Status').value = "Completed";
                                        fd.field('Status').disabled = true;                                                                            
                                    }
                                    else{

                                        $(fd.field('FormStatus').$parent.$el).show();
                                        fd.field('FormStatus').value = "Saved By Trade";
                                        $(fd.field('FormStatus').$parent.$el).hide();                                                                                                                   
                                    }
                                }
																
                                if(fd.isValid){
                                    await PreloaderScripts();  
                                    _proceed = true;                                 
                                    
                                    if(Code === 'Approved'){                                        
                                        _Email = 'IssuedLLItem_Email';
                                        _Notification = 'LessonsLearned_Reviewed';  
                                    }
                                    else{                                        
                                        _Email = 'IssuedLLRejectedItem_Email';
                                        _Notification = 'LessonsLearned_Rejected';
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
                                await PreloaderScripts();
                                fd.close();
                            }			
                        });
					
					   IsPMUser = true;
				    }
				}				
			});
	     });
    }
	catch(e){alert(e);}
	return IsPMUser;
}

async function CheckifUserRole() {
	var IsTMUser = "Trade";
	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(async function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {				
					if(groupsData[i].Title === "PD")
					{					
					   IsTMUser = "PD";
					   break;
				    }					
					else
					{					
					    IsTMUser = "Trade";					   
				    }
				}				
			});
	     });
    }
	catch(e){alert(e);}
	return IsTMUser;				
}

function CodehideOrShow() {	

    if (fd.field('Code').value === "Returned to initiator with Comments") {
		$(fd.field('ReasonofRejection').$parent.$el).show();
		fd.field('ReasonofRejection').required = true;	
		
		$(fd.field('Classification').$parent.$el).show();
		fd.field('Classification').required = false;
		fd.field('Classification').value = "";
		$(fd.field('Classification').$parent.$el).hide();	
	} 
	else if (fd.field('Code').value === "Approved") {
		$(fd.field('Classification').$parent.$el).show();
		fd.field('Classification').required = true;		
		
		$(fd.field('ReasonofRejection').$parent.$el).show();
		fd.field('ReasonofRejection').required = false;
		fd.field('ReasonofRejection').value = "";
		$(fd.field('ReasonofRejection').$parent.$el).hide();				
	}
	else {
	
		$(fd.field('ReasonofRejection').$parent.$el).show();
		fd.field('ReasonofRejection').required = false;
		fd.field('ReasonofRejection').value = "";
		$(fd.field('ReasonofRejection').$parent.$el).hide();
		
		$(fd.field('Classification').$parent.$el).show();	
		fd.field('Classification').required = false;
		fd.field('Classification').value = "";
		$(fd.field('Classification').$parent.$el).hide();					
	}
}

function fixTextArea(){
	$("textarea").each(function(index){			
		var height = (this.scrollHeight + 5) + "px";		
        $(this).css('height', height);
	});
}

function ReadOnly(elem){ 
    $(elem).prop("readonly", true).css({
	    "background-color": "transparent",
	    "border": "0px"
	});  
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

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function clearLocalStorageItemsByField(fields) {
	fields.forEach(field => {
		let cachedFields = localStorage;
        for (let i = 0; i < cachedFields.length; i++) {
	        const key = localStorage.key(i);        
	        if (key.includes(field)){
	        	localStorage.removeItem(key);
	        }
    	}
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

//#endregion
