var _layout = '/_layouts/15/PCW/General/EForms', 
    _siteUrl = _spPageContextInfo.siteAbsoluteUrl,
    _ImageUrl = _siteUrl + '/Style%20Library/tooltip.png',
    _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5],
    _ListFullUrl = _siteUrl + '/Lists/' + _ListInternalName;    

var _modulename = "", _formType = "";
let _proceed = false;

const _listName = 'Lessons Learned';
let itemId = '';

const ProjectNameList = "";
const ProjectYear = 2010;
let projectArr=[];

const itemsToRemove = ['Status', 'State', 'Code', 'WorkflowStatus'];

const blueColor = '#6ca9d5', greenColor = '#5FC9B3', yellowColor = '#F7D46D', redColor = '#F28B82';

var onRender = async function (moduleName, formType){    
	try { 
        const startTime = performance.now();
		_modulename = moduleName;
		_formType = formType;
		if(moduleName == 'MELL')
			await onMELLRender(formType);

        const endTime = performance.now();
        const elapsedTime = endTime - startTime;
        console.log(`Execution time onMELLRender: ${elapsedTime} milliseconds`);
	}
	catch (e) {
		//alert(e);
		console.log(e);	
	}
}

var onMELLRender = async function (formType){	

    await PreloaderScripts(); 			
	await loadScripts();
 
    clearLocalStorageItemsByField(itemsToRemove);

	if(formType == 'New'){        
		await MELL_newForm();        
    } 	
    else if(formType == 'Edit'){        
		await MELL_editForm();       
    }    
    else if(formType == 'Display'){        
		await MELL_displayForm();     
    }    
}

var MELL_newForm = async function(){ 
    
    fd.toolbar.buttons[0].style = "display: none;";   
	fd.toolbar.buttons[1].style = "display: none;";
    
    const hideField = (field) => $(fd.field(field).$parent.$el).hide();
    ['AttachFiles', 'Client_Name', 'Country_Name', 'Status', 'FormStatus', 'AutoRef', 'PM'].forEach(hideField);  
	
    await loadingButtons()
    formatingButtonsBar();		
	
	fd.field('Reference').required = true;
	fd.field('Recommendation').required = true;

	var LoginNameU = _spPageContextInfo.userLoginName;
	var Domain = LoginNameU.split("\\")[0];
	if(Domain.toUpperCase() === "DMSD")
		Domain = "DarCairo";
	fd.field('Office').value = Domain.toUpperCase();
	$(fd.field('Office').$parent.$el).hide();   
    
    fd.field('Leader').value = _spPageContextInfo.userDisplayName;
    $(fd.field('Leader').$parent.$el).hide();
	
	fd.field('UserDisplayName').value = _spPageContextInfo.userDisplayName;
	$(fd.field('UserDisplayName').$parent.$el).hide();	
	
	fd.field('Trade').value = 'ME';
	fd.field('Trade').disabled = true;
	
	if(fd.field('Trade').value == "ME")
	{
		var content = "<span style='font-weight:bold;color:red'>Note:</span> Please click <a href='https://bi.dar.com/ReportServer/Pages/ReportViewer.aspx?/wapps/PMISReports/PCRLessonsLearnedPublic' target='_blank'>here</a> in case you need to check lessons learned from 'Project Completion Report in PMIS'";
		$('#my-html').append(content);
	}
	
	$(fd.field('ProjectNo').$parent.$el).hide();
	fd.field('Type').value = 'General';
	
	fd.field('Country').value = "";
	fd.field('ClientName').value = "";
	fd.field('ReferenceName').value = "";
			
	fd.field('Country').disabled = true;
	fd.field('ClientName').disabled = true;
	fd.field('ReferenceName').disabled = true;	
	  
	fd.field('ClientName').$on('change', function(value)
	{
		$(fd.field('Client_Name').$parent.$el).show();
		fd.field('Client_Name').value = value.LookupValue;
		$(fd.field('Client_Name').$parent.$el).hide();
	}); 
	
	fd.field('Country').$on('change', function(value)
	{
		$(fd.field('Country_Name').$parent.$el).show();
		fd.field('Country_Name').value = value;
		$(fd.field('Country_Name').$parent.$el).hide();
	}); 	

	fd.field('Type').$on('change', async function(value)
	{		
		if(value === "General")
		{	
			$(fd.field('Client_Name').$parent.$el).hide();
			$(fd.field('Country_Name').$parent.$el).hide();
			
			$(fd.field('ClientName').$parent.$el).show();
			$(fd.field('Country').$parent.$el).show();			
            
            //fd.field('Reference').ready(function() {
                fd.field('Reference').widget.dataSource.data(['General']);
                fd.field('Reference').value = ['General'];		
                fd.field('Reference').disabled = true;
            //});	
			
			fd.field('ReferenceName').value = "General";
			fd.field('ReferenceName').disabled = true;
			
			$(fd.field('ProjectNo').$parent.$el).show();	
	        fd.field('ProjectNo').value  = "General";
			$(fd.field('ProjectNo').$parent.$el).hide();
			
			fd.field('Country').value = "";
			fd.field('ClientName').value = "";	
			
			fd.field('Country').disabled = false;
			fd.field('ClientName').disabled = false;			
							
				
			fd.field('Country').required = true;
			fd.field('ClientName').required = true;				
		}
		else if (value === "Project")
		{
			fd.field('Country').value = "";
			fd.field('ClientName').value = "";
			fd.field('ReferenceName').value = "";
	
			fd.field('Country').disabled = true;
			fd.field('ClientName').disabled = true;
			fd.field('ReferenceName').disabled = true;
			
			fd.field('Country').required = false;
			fd.field('ClientName').required = false;	
			fd.field('ReferenceName').required = false;
						
			fd.field('Reference').disabled = false;
			fd.field('Reference').widget.dataSource.data(['']);
			fd.field('Reference').required = true; 
            
            //fd.field('Reference').ready(async function() {
                fd.field('Reference').widget.dataSource.data(['Please wait while retreiving ...']);
                await GetProjectListALLP(ProjectYear);
                fd.field('Reference').required = true;	
            //});            		
			
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
		}	 
 	});	
	
	fd.field('Reference').$on('change', async function(value)
	{        
        if(value.length > 0)
            await GetProjectList(value);        
        else{

            $(fd.field('ProjectNo').$parent.$el).show();	
            fd.field('ProjectNo').value  = '';
            $(fd.field('ProjectNo').$parent.$el).hide();
            fd.field('ReferenceName').value  = '';
            
            $(fd.field('Client_Name').$parent.$el).show();
            fd.field('Client_Name').value = '';
            fd.field('Client_Name').disabled = true;
            $(fd.field('ClientName').$parent.$el).hide();
            
            $(fd.field('Country_Name').$parent.$el).show();
            fd.field('Country_Name').value = '';
            fd.field('Country_Name').disabled = true;
            $(fd.field('Country').$parent.$el).hide();
        }
	});	
	
	if(fd.field('Type').value === "Project")
	{		
		//fd.field('Reference').ready(async function() {
            fd.field('Reference').widget.dataSource.data(['Please wait while retreiving ...']);
            await GetProjectListALLP(ProjectYear);
            fd.field('Reference').required = true;	
        //}); 		
	  
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
	} 
	else if(fd.field('Type').value === "General")
	{
		$(fd.field('Client_Name').$parent.$el).hide();
		$(fd.field('Country_Name').$parent.$el).hide();
		
		$(fd.field('ClientName').$parent.$el).show();
		$(fd.field('Country').$parent.$el).show();		
		
        //fd.field('Reference').ready(function() {
            fd.field('Reference').widget.dataSource.data(['General']);
            fd.field('Reference').value = ['General'];		
            fd.field('Reference').disabled = true;
        //});				
		
		fd.field('ReferenceName').value = "General";
		fd.field('ReferenceName').disabled = true;
		
		$(fd.field('ProjectNo').$parent.$el).show();	
        fd.field('ProjectNo').value  = "General";
		$(fd.field('ProjectNo').$parent.$el).hide();
		
		fd.field('Country').value = "";
		fd.field('ClientName').value = "";	
		
		fd.field('Country').disabled = false;
		fd.field('ClientName').disabled = false;			
						
			
		fd.field('Country').required = true;
		fd.field('ClientName').required = true;
	}

    debugger;
    const group = await pnp.sp.web.siteGroups.getByName('PD').get();
    const groupId = group.Id;    
    //const PD = await pnp.sp.web.siteGroups.getById(groupId).users();
    const PDGroups = await pnp.sp.web.siteGroups.getById(groupId).expand('Users').get(); 
    const PD = PDGroups.Users;

    if (PD.length > 0) {
        let contributors=[];
        PD.map(user => {
            contributors.push(user.Title);                      
        });

        if (contributors.length > 0) {                        
            $(fd.field('PM').$parent.$el).show();
            fd.field('PM').value = contributors; 
            $(fd.field('PM').$parent.$el).hide();                      
        }
    } 

    preloader("remove");    
}

var MELL_editForm = async function(){ 

    fd.toolbar.buttons[0].style = "display: none;"; 
    fd.toolbar.buttons[1].style = "display: none;";
    
    fixTextArea();
    itemId = fd.itemId;
	
	$(fd.field('FormStatus').$parent.$el).hide();
    $(fd.field('AutoRef').$parent.$el).hide();	
	$(fd.field('Submit').$parent.$el).hide();
	$(fd.field('AttachFiles').$parent.$el).hide();
    $(fd.field('Leader').$parent.$el).hide();
    fd.field('Status').disabled = true;   
    fd.field('UserDisplayName').disabled = true;
    fd.field('Trade').disabled = true;
    $(fd.field('PM').$parent.$el).hide();
    $(fd.field('ClosedDate').$parent.$el).hide();  
	
	//var isSiteAdmin = _spPageContextInfo.isSiteAdmin; 	
	const GroupName = await CheckifUserRole();
	
	if (fd.field('Status').value === "Initiated" && GroupName === "Trade") 
	{			
		fd.field('Title').disabled = false;
		fd.field('Problem').disabled = false;
		
		fd.field('Code').disabled = true;
		fd.field('Code').value = "";
		
		$(fd.field('Code').$parent.$el).hide();
		$(fd.field('Classification').$parent.$el).hide();
		$(fd.field('ReasonofRejection').$parent.$el).hide();
        fd.field('UserDisplayName').disabled = true;
        fd.field('Trade').disabled = true;

        // $(fd.field('Leader').$parent.$el).show();
        // fd.field('Leader').value = _spPageContextInfo.userDisplayName;
        // $(fd.field('Leader').$parent.$el).hide();        
        // fd.field('UserDisplayName').disabled = false;
        // fd.field('UserDisplayName').value = _spPageContextInfo.userDisplayName;       
        // fd.field('Trade').disabled = false;
        // fd.field('Trade').value = 'ME';
        // const group = await pnp.sp.web.siteGroups.getByName('PD').get();
        // const groupId = group.Id;    
        // //const PD = await pnp.sp.web.siteGroups.getById(groupId).users(); 
        // const PDGroups = await pnp.sp.web.siteGroups.getById(groupId).expand('Users').get(); 
        // const PD = PDGroups.Users;
        // if (PD.length > 0) {
        //     let contributors=[];
        //     PD.map(user => {
        //         contributors.push(user.Title);                      
        //     });

        //     if (contributors.length > 0) {                        
        //         $(fd.field('PM').$parent.$el).show();
        //         fd.field('PM').value = contributors; 
        //         $(fd.field('PM').$parent.$el).hide();                      
        //     }
        // }
	
		fd.toolbar.buttons.push({
		        icon: 'Save',
		        class: 'btn-outline-primary',
		        text: 'Save',
                style: `background-color:${blueColor}; color:white`,
		        click: async function() {				
            
				$(fd.field('FormStatus').$parent.$el).show();
				fd.field('FormStatus').value = "Saved By Trade";
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
				
                // await PreloaderScripts();
				// $(fd.field('FormStatus').$parent.$el).show();
				// fd.field('FormStatus').value = "";
				// $(fd.field('FormStatus').$parent.$el).hide();           

                let web = pnp.sp.web;            
                const leader = _spPageContextInfo.userDisplayName;                             
            
                let refNo = await updateCounter();             

                fd.field('Status').disabled = false;
                fd.field('Status').value = 'Open';
                fd.field('Status').disabled = true;
                
                $(fd.field('AutoRef').$parent.$el).show();
                fd.field('AutoRef').value = refNo;
                $(fd.field('AutoRef').$parent.$el).hide();

                $(fd.field('FormStatus').$parent.$el).show();
                fd.field('FormStatus').value = 'Saved By PD';
                $(fd.field('FormStatus').$parent.$el).hide();         
				
		        if(fd.isValid){
                    await PreloaderScripts();
                    fd.save().then(async function() {
                        let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
                        await _sendEmail(_modulename, 'NewLLItem_Email', query, '', 'LessonsLearned_Initiated', '');
                    });                    
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
	else if(fd.field('Status').value === "Open" && fd.field('Code').value === "Returned to initiator with Comments" && GroupName === "Trade")
	{
		$(fd.field('Code').$parent.$el).show();
		fd.field('Code').disabled = true;		
		fd.field('Title').disabled = false;
		fd.field('Problem').disabled = false;
        fd.field('Recommendation').disabled = false;		
		$(fd.field('ReasonofRejection').$parent.$el).show();		
		// var elem = $("textarea")[1]; // Select the textarea by its class 
        // $(elem).prop("readonly", true);	
        fd.field('ReasonofRejection').disabled = true;			
		$(fd.field('Classification').$parent.$el).hide();
		
		fd.toolbar.buttons.push({
		        icon: 'Save',
		        class: 'btn-outline-primary',
		        text: 'Save',
                style: `background-color:${blueColor}; color:white`,
		        click: async function() {  
				
				$(fd.field('FormStatus').$parent.$el).show();
				fd.field('FormStatus').value = "Saved By Trade";
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
				
				fd.field('Code').disabled = false;	
				fd.field('Code').value = "";
				fd.field('Code').disabled = true;
				
				$(fd.field('FormStatus').$parent.$el).show();
				fd.field('FormStatus').value = "Saved By PD";
				$(fd.field('FormStatus').$parent.$el).hide();               
		        
                if(fd.isValid){
                    await PreloaderScripts(); 
                    fd.save().then(async function() {
                        let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
                        await _sendEmail(_modulename, 'NewLLItem_Email', query, '', 'LessonsLearned_Initiated', '');
                    }); 
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
	else if(fd.field('Status').value === "Open" && GroupName === "PD")
	{
		$(fd.field('Code').$parent.$el).show();
        fd.field('Code').required = true;  
        $(fd.field('Leader').$parent.$el).show();
        let Initiator = fd.field('Leader').value;
        if (Initiator && Initiator.EntityData && Initiator.EntityData.Email) {
            Initiator = Initiator.EntityData.Email;
        }
        $(fd.field('Leader').$parent.$el).hide();  	
		MEIsUserInGroup('PD', Initiator);
	}	
	else
	{
		fd.toolbar.buttons[0].disabled = true; //Disable Save button
        fd.field('Code').disabled = true;		
		fd.field('Title').disabled = true;
		fd.field('Problem').disabled = true;
        fd.field('Recommendation').disabled = true;    
        // var textarea = document.querySelector('textarea.form-control.fd-textarea');
		// ReadOnly(textarea); 
        fd.field('Category').disabled = true;
		fd.field('ProjectPhase').disabled = true;	
		fd.field('Impact').disabled = true;	
		
		if (fd.field('Code').value === "Returned to initiator with Comments")
		{			
			fd.field('ReasonofRejection').required = false;
			fd.field('ReasonofRejection').disabled = true;	
		}
		else if (fd.field('Code').value === "Approved")
		{
			fd.field('Classification').required = false;
			fd.field('Classification').disabled = true;	
			
			$(fd.field('Classification').$parent.$el).hide();
			$(fd.field('ReasonofRejection').$parent.$el).hide();
		}
		else
		{
			$(fd.field('ReasonofRejection').$parent.$el).hide();	
			$(fd.field('Classification').$parent.$el).hide();
		}

        SetAttachmentToReadOnly();

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
	
	if(GroupName == "PD")
	{
		fd.field('Code').$on('change',CodehideOrShow);
		CodehideOrShow();
	}   

    formatingButtonsBar();

    preloader("remove"); 
}

var MELL_displayForm = async function(){

    fd.toolbar.buttons[1].text = "Cancel";
    fd.toolbar.buttons[1].icon = "Cancel";
    fd.toolbar.buttons[1].style = `background-color:${redColor}; color:white;`;

    fd.toolbar.buttons[0].icon = "Edit";
    fd.toolbar.buttons[0].text = "Edit";
    fd.toolbar.buttons[0].class = 'btn-outline-primary';
    fd.toolbar.buttons[0].style = `background-color:${greenColor}; color:white;`;

    formatingButtonsBar();
    fixTextArea();

    const GroupName = await CheckifUserRole();	
	if(GroupName === "Trade")
		$(fd.field('Classification').$parent.$el).hide();	

    preloader("remove"); 
}

fd.spBeforeSave(function(spForm){					
	return fd._vue.$nextTick();
});

fd.spSaved(async function(result) {			
    try
    { 
        if(_proceed){		
            var itemId = result.Id;        
            let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
            await _sendEmail(_modulename, 'NewLLItem_Email', query, '', 'LessonsLearned_Initiated', '');    
        }   						
    }
    catch(e){
        console.log(e);
    }								 
});

//#region General Functions

async function GetProjectListALLP(ProjectYear) 
{ 
    projectArr = JSON.parse(localStorage.getItem('MELessonsProjects')) || [];

    if(projectArr.length > 0){        
        setTimeout(function() {
            fd.field('Reference').widget.setDataSource({
                data: projectArr
            });
        }, 1400);
    }

    var xhr = new XMLHttpRequest();
    var restURL = _siteUrl+ "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=getAllProjectsbyear&ProjectYear="+encodeURIComponent(ProjectYear);
    xhr.open("GET", restURL, true);

    xhr.onreadystatechange = async function () {
        if (xhr.readyState == 4) { 
            
            try {

                if(xhr.status == 200) { 

                    const obj = JSON.parse(this.responseText);
                    const newProjectArr = obj.map(d => d.ProjectCode.trim()); 
                    
                    const diff2 = newProjectArr.filter(x => !projectArr.includes(x));

                    if (diff2.length > 0) {                        
                        localStorage.setItem('MELessonsProjects', JSON.stringify(newProjectArr));                       
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
                console.log(err + "\n" + text);				
            } 
        }
    } 

    xhr.send();
}

function GetProjectList(ProjectName) 
{
    var xhr = new XMLHttpRequest()
  
    xhr.open('GET', _siteUrl + "/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=getprojectname&ProjectCode="+encodeURIComponent(ProjectName), true);

    xhr.onreadystatechange = async function () {
        if (xhr.readyState == 4) { 
            
            try {

                if(xhr.status == 200) { 

                    const data = JSON.parse(this.responseText);

                    data.forEach((d) => 
                        {		 	
                             $(fd.field('ProjectNo').$parent.$el).show();	
                             fd.field('ProjectNo').value  = d.ProjectCode;
                             $(fd.field('ProjectNo').$parent.$el).hide();
                             fd.field('ReferenceName').value  = d.ProjectName;
                             
                             $(fd.field('Client_Name').$parent.$el).show();
                             fd.field('Client_Name').value = d.ClientName;
                             fd.field('Client_Name').disabled = true;
                             $(fd.field('ClientName').$parent.$el).hide();
                             
                             $(fd.field('Country_Name').$parent.$el).show();
                             fd.field('Country_Name').value = d.AreaName.toUpperCase();
                             fd.field('Country_Name').disabled = true;
                             $(fd.field('Country').$parent.$el).hide();				 	 
                        })
                }
            }  
            catch(err) 
            {
                console.log(err + "\n" + text);				
            } 
        }
    }     
    xhr.send() 
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
        element.style.color = '#3CDBC0';
        element.style.borderTopLeftRadius = '6px';
        element.style.borderBottomLeftRadius = '6px';
        element.style.width = '200px';
        element.style.display = 'inline-block';
    });
}

async function loadingButtons(){

$(fd.field('Submit').$parent.$el).hide();
	
	fd.toolbar.buttons.push({
	        icon: 'Save',
	        class: 'btn-outline-primary',
	        text: 'Save',
            style: `background-color:${blueColor}; color:white`,
	        click: async function() {           		            
                
            $(fd.field('Status').$parent.$el).show();
            fd.field('Status').value = 'Initiated';
            $(fd.field('Status').$parent.$el).hide();          

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

            let web = pnp.sp.web;            
			const leader = _spPageContextInfo.userDisplayName;                             
           
            let refNo = await updateCounter();                   
                
            $(fd.field('Status').$parent.$el).show();
            fd.field('Status').value = 'Open';
            $(fd.field('Status').$parent.$el).hide();
            
            $(fd.field('AutoRef').$parent.$el).show();
            fd.field('AutoRef').value = refNo;
            $(fd.field('AutoRef').$parent.$el).hide();

            $(fd.field('FormStatus').$parent.$el).show();
            fd.field('FormStatus').value = 'Saved By PD';
            $(fd.field('FormStatus').$parent.$el).hide(); 
            
            // let EmailbodyHeader = `A new Lesson Learned of reference number ${refNo} has been sent to you for review.`;
            // let Title = fd.field('Title').value; 
            // let Category = fd.field('Category').value; 
            // let Office = fd.field('Office').value;           

            // $(fd.field('Country_Name').$parent.$el).show();
            // let Country_Name = fd.field('Country_Name').value;
            // $(fd.field('Country_Name').$parent.$el).hide();

            // $(fd.field('Client_Name').$parent.$el).show();
            // let Client_Name = fd.field('Client_Name').value;
            // $(fd.field('Client_Name').$parent.$el).hide();

            // let Result = {  'Lesson Learned Title': Title, 
            //     Country: Country_Name, 
            //     Client: Client_Name, 
            //     Category: Category, 
            //     Office: Office, 
            //     ModifiedBy: leader                
            // };

            // let Subject = `New Lesson Learned - ${refNo} - ${Title}`;
            // let encodedSubject = htmlEncode(Subject);
            // let Body = await GetHTMLBody(Result, EmailbodyHeader, leader);
            // let encodedBody = htmlEncode(Body);

            //await _sendEmail(_modulename, encodedSubject + '|' + encodedBody, '', '', 'LessonsLearned_Initiated', '');
            // let itemId = await getMaxItemId(_listName);            

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
                                    fd.save().then(async function() {                                        
                                        let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
                                        if(Code === 'Approved')
                                            await _sendEmail(_modulename, 'IssuedLLItem_Email', query, Initiator, 'LessonsLearned_Reviewed', '');
                                        else
                                            await _sendEmail(_modulename, 'IssuedLLRejectedItem_Email', query, Initiator, 'LessonsLearned_Rejected', '');
                                    });                                   
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

//#endregion
