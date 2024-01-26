var _layout = "/_layouts/15/PCW/General/EForms";

var _modulename = "", _formType = "";

var script = document.createElement("script"); // create a script DOM node
script.src = _layout + "/plumsail/js/config/configFileRoutingDeliverablesTasks.js"; // set its src to the provided URL
document.head.appendChild(script);


var onRender = async function (moduleName, formType){

	localStorage.clear();

    _modulename = moduleName;
    _formType = formType;

	setTimeout(removePadding, 1000);

    if(moduleName == 'SPart')
         onSPartRender(formType);
	
	if(moduleName == 'SLead')
         onSLeadRender(formType); 
	
	var isValid = false;
	var retry = 1;
	while (!isValid)
	{
		try{
			if(retry >= 7) break;
			setButtonToolTip('Save', saveMesg);
			setButtonToolTip('Submit', submitMesg);
			setButtonToolTip('Cancel', cancelMesg);
			isValid = true;
			}
		catch{
			retry++;
			await delay(1000);
		}
	}	
}

var onSPartRender = async function (formType){
    if(formType == 'Edit'){
        fd.toolbar.buttons[0].style = "display: none;";
		fd.toolbar.buttons[1].text = "Cancel"; 	

		fd.field('Status').disabled = true;
		$(fd.field('AttachFiles').$parent.$el).hide();
		$(fd.field('FullRefs').$parent.$el).hide();

		var isSiteAdmin = _spPageContextInfo.isSiteAdmin; 
		var AssignedToUsers = fd.field('AssignedTo').value;	
		let _isAllowed = false
		var currentUser = await pnp.sp.web.currentUser.get();

		if(!isSiteAdmin)
		{		
			AssignedToUsers.map(item =>{
				if(item.displayName == currentUser.Title)
					_isAllowed = true;
			});

			if(!_isAllowed)
			{
				alert("Apologies, but this task has not been assigned to you.");
				fd.close();
			}		
		}
	
		if(fd.field('Status').value === 'Open')
		{	
			fd.toolbar.buttons.push({
							icon: 'Save',
							class: 'btn-outline-primary',
							text: 'Save',
							click: function() {					
									
							fd.save();			 
						}
					});

			fd.toolbar.buttons.push({
							icon: 'Accept',
							class: 'btn-outline-primary',
							text: 'Submit',
							click: function() {
							
							var NCountofATT = 0;
							var OCountofATT = 0;
							for(i = 0; i < fd.field('Attachments').value.length; i++) 
							{
								var Val = fd.field('Attachments').value[i].extension.toString();
								if(Val === "")
								{OCountofATT++;}
								else
								{			
										NCountofATT++;			
								}
							}
							$(fd.field('AttachFiles').$parent.$el).show();
							fd.field('AttachFiles').value = "U" + "," + NCountofATT + "," + OCountofATT;
							$(fd.field('AttachFiles').$parent.$el).hide();
							
							$(fd.field('FullRefs').$parent.$el).show();
							fd.field('FullRefs').value = "|";
							$(fd.field('FullRefs').$parent.$el).hide();
							
							fd.field('Status').disabled = false;
							fd.field('Status').value = "Completed";
							fd.field('Status').disabled = true;
								
							fd.save();						 
						}
					});			
		}
		else
		{
			fd.field('Comment').disabled = true;

			var spanATTDelElement = document.querySelector('.k-upload .k-upload-files .k-upload-status');
			spanATTDelElement.style.display = 'none';			
			var spanATTUpElement = document.querySelector('.k-upload .k-upload-button');
			spanATTUpElement.style.display = 'none';	
		}
    }  
    if(formType == 'Display'){        
        
    }
}

var onSLeadRender = async function (formType){
    if(formType == 'Edit'){
        fd.toolbar.buttons[0].style = "display: none;";
		fd.toolbar.buttons[1].text = "Cancel"; 		
    }  
    if(formType == 'Display'){        
        
    }
}

function delay(time) {
	return new Promise(function(resolve) { 
		setTimeout(resolve, time)
	});
}

function removePadding(){
	$(".fd-form-block > .fd-grid[data-v-105ebe50]").attr('style', 'padding: 12px !important');
  
	var mainContent = $('#spPageChromeAppDiv div').next();
	var targetElement = mainContent[2];
	targetElement.style.marginLeft = '-200px';
  }
