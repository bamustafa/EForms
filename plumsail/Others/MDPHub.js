var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDisplay = false;

var  _currentUser, _formFields = {};
const greenColor = '#5FC9B3', redColor = '#F28B82', yellowColor = '#6c757d';

let formBarTitle = 'Dar MDP Hub';

let _isSendEmail = false, _Discipline = '';
let _Status = '', _Decision = '', _ReviewedDate = '', _nextStatus = '';
let _isAdmin = false, _isNormal = false;
const disableField = (field) => field.disabled = true;
const hideField = (field) => $(field.$parent.$el).hide();
const showField = (field) => $(field.$parent.$el).show();

var onRender = async function (relativeLayoutPath, moduleName, formType) {
    
    try {        

        _layout = relativeLayoutPath;   

        const pmcBugsvg = `<svg fill="#49c4b1" width="22" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 324.274 324.274" stroke="#49c4b1">
            <g>
                <path d="M34.419,298V22h138.696v33.411c0,8.301,6.753,15.055,15.053,15.055h33.154v88.5c2.443-0.484,4.957-0.75,7.528-0.75
                c5.087,0,9.962,0.994,14.472,2.804V64.006c0-1.326-0.526-2.598-1.464-3.536L183.694,1.464C182.755,0.527,181.484,0,180.158,0
                H27.472c-8.3,0-15.053,6.753-15.053,15.054v289.893c0,8.301,6.753,15.054,15.053,15.054h111.884
                c-1.256-6.713,1.504-13.831,7.559-17.83c2.341-1.546,4.692-2.919,7.034-4.17H34.419z"/>
                <path d="M308.487,310.515c-12.254-8.092-25.057-11.423-33.599-12.795c6.02-9.685,9.564-21.448,9.564-34.129
                c0-9.12-1.824-17.889-5.174-25.781c8.22-1.738,18.908-5.176,29.209-11.98c3.457-2.283,4.408-6.935,2.126-10.392
                c-2.283-3.456-6.936-4.407-10.392-2.125c-10.742,7.094-22.229,9.723-29.102,10.698c-3.459-4.387-7.5-8.249-12.077-11.394
                c0.859-3.081,1.294-6.265,1.294-9.509c0-17.861-13.062-32.393-29.117-32.393c-16.055,0-29.115,14.531-29.115,32.393
                c0,3.244,0.435,6.428,1.294,9.509c-4.577,3.145-8.618,7.007-12.077,11.394c-6.873-0.975-18.358-3.603-29.102-10.698
                c-3.456-2.282-8.108-1.331-10.392,2.125c-2.282,3.456-1.331,8.109,2.126,10.392c10.301,6.803,20.988,10.241,29.208,11.979
                c-3.351,7.893-5.175,16.661-5.175,25.781c0,12.681,3.544,24.444,9.563,34.129c-8.541,1.372-21.343,4.703-33.597,12.794
                c-3.456,2.283-4.408,6.935-2.126,10.392c1.442,2.184,3.83,3.368,6.266,3.368c1.419,0,2.854-0.402,4.126-1.242
                c16.62-10.975,35.036-11.269,35.362-11.272c0.639-0.002,1.255-0.093,1.847-0.245c8.877,7.447,19.884,11.861,31.791,11.861
                c11.907,0,22.914-4.415,31.791-11.861c0.598,0.153,1.22,0.244,1.865,0.245c0.183,0,18.499,0.148,35.346,11.272
                c1.272,0.84,2.707,1.242,4.126,1.242c2.434,0,4.823-1.184,6.266-3.368C312.895,317.45,311.943,312.797,308.487,310.515z
                M238.719,296.005c0,4.142-3.357,7.5-7.5,7.5c-4.142,0-7.5-3.358-7.5-7.5v-64.83c0-4.142,3.358-7.5,7.5-7.5
                c4.143,0,7.5,3.358,7.5,7.5V296.005z"/>
                <path d="M143.627,49.624h-78c-4.418,0-8,3.582-8,8s3.582,8,8,8h78c4.418,0,8-3.582,8-8S148.045,49.624,143.627,49.624z"/>
                <path d="M143.627,99.624h-78c-4.418,0-8,3.582-8,8s3.582,8,8,8h78c4.418,0,8-3.581,8-8S148.045,99.624,143.627,99.624z"/>
                <path d="M143.627,149.624h-78c-4.418,0-8,3.582-8,8s3.582,8,8,8h78c4.418,0,8-3.581,8-8S148.045,149.624,143.627,149.624z"/>
            </g>
            </svg>`;
        setIconSource("bug-icon", pmcBugsvg); 
        const pmcSubmitsvg = `<svg version="1.1" width="24" id="Icons" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 32 32" xml:space="preserve" fill="#49c4b1" stroke="#49c4b1"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <style type="text/css"> .st0{fill:none;stroke:#49c4b1;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;stroke-miterlimit:10;} </style> <line class="st0" x1="16" y1="20" x2="16" y2="4"></line> <polyline class="st0" points="12,8 16,4 20,8 "></polyline> <polyline class="st0" points="9,13 3,16.5 3,21.5 16,29 29,21.5 29,16.5 23,13 "></polyline> </g></svg>`;      
        setIconSource("Submit-icon", pmcSubmitsvg); 
        const pmcFeedbacksvg = `<svg viewBox="0 0 24 24" width="24" fill="none" xmlns="http://www.w3.org/2000/svg"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path d="M8.24999 18L5.24999 20.25V15.75H2.25C1.85217 15.75 1.47064 15.5919 1.18934 15.3106C0.908034 15.0293 0.749999 14.6478 0.749999 14.25V2.25C0.749999 1.85217 0.908034 1.47064 1.18934 1.18934C1.47064 0.908034 1.85217 0.749999 2.25 0.749999H18.75C19.1478 0.749999 19.5293 0.908034 19.8106 1.18934C20.0919 1.47064 20.25 1.85217 20.25 2.25V6.71484" stroke="#49c4b1" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"></path> <path d="M5.24999 5.24999H15.75" stroke="#49c4b1" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"></path> <path d="M5.24999 9.74999H8.24999" stroke="#49c4b1" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"></path> <path d="M23.25 18.75H20.25V23.25L15.75 18.75H11.25V9.74999H23.25V18.75Z" stroke="#49c4b1" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"></path> <path d="M19.5 15H15" stroke="#49c4b1" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"></path> </g></svg>`;      
        setIconSource("Feedback-icon", pmcFeedbacksvg); 
    
        await loadScripts().then(async () => {
        
            showPreloader();        
            
            await extractValues(moduleName, formType);
            await setCustomButtons();

            if (moduleName == "MDPHub") {                 
                if(formType === "New")
                    await MDPHub_newForm();  
                else if(formType=="Edit")
                    await MDPHub_editForm();
                else if(formType=="Display") 
                    await MDPHub_displayForm();
            }
            else if(moduleName=="Submit"){
                if(formType === "New")
                    await Submit_newForm();
                else if(formType=="Edit")
                    await Submit_editForm();
                else if(formType=="Display") 
                    await Submit_displayForm();
            }
            else if(moduleName=="Feedback"){
                if(formType === "New")
                    await Feedback_newForm(); 
                else if(formType=="Edit")
                    await Feedback_editForm();
                else if(formType=="Display") 
                    await Feedback_displayForm();
            }
            else if(moduleName=="Ticket"){
                if(formType === "New")
                    await Ticket_newForm(); 
                else if(formType=="Edit")
                    await Ticket_editForm();
                else if(formType=="Display") 
                    await Ticket_displayForm();    
            }
        })
    }
    catch (e){
      showPreloader();
      fd.toolbar.buttons[0].style = "display: none;";
      fd.toolbar.buttons[1].style = "display: none;";
      console.log(e)
    }
    finally{
        hidePreloader();
    }
}  


var MDPHub_newForm = async function(){     
  try 
  {  
        setPSHeaderMessage('');
        setPSErrorMesg(`Please fill in the below fields.`);

        _formFields = {
            Trade: fd.field('Trade'),
            DisciplineMembers: fd.field('DisciplineMembers')
        }
      
        formatingButtonsBar(formBarTitle);

        //document.querySelector(".divGrid").closest(".row").style.paddingLeft = "15px";
        //document.querySelector(".divGrid").closest(".row").style.paddingTop = "20px";

        let queryString = window.location.search;
        let source = true;
        if (!queryString.includes("QueryVal")) {
            queryString = "?QueryVal=Clean%20Revit%20Files"; // Default query string
            source = false; // Update the source message
        }

        let actualQueryString = queryString.split('?').pop();

        let urlParams = new URLSearchParams(actualQueryString);
        let queryVal = urlParams.get('QueryVal');
        queryVal = queryVal.replace(/^'|'$/g, '');

        fd.field("Title").value = queryVal;
        fd.field("Title").disabled = true;

        console.log("QueryVal:", queryVal);
        console.log("Source:", source);
        // Call validation when files are added
        fd.field('Attachments').$on('change', validateAttachments);
        
        $(_formFields.DisciplineMembers.$parent.$el).hide();

        _formFields.Trade.$on('change', async function (value) { 
            let AllTeamMembers = await fillinDisciplineMembers(value);
            let TeamMembers = AllTeamMembers.toUsers;
            $(_formFields.DisciplineMembers.$parent.$el).show();
            _formFields.DisciplineMembers.value = TeamMembers?.map(user => user.Title) ?? '';
            $(_formFields.DisciplineMembers.$parent.$el).hide();           
        });

        _isSendEmail = true;     
  }    
  catch(err){
      console.log(err.message, err.stack)
      await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
      showPreloader();       
  }     
}
var MDPHub_editForm = async function(){   
    try 
    {
        setPSHeaderMessage('');                      
        
        _formFields = {
            Title: fd.field('Title'),
            Trade: fd.field('Trade'),
            Issue: fd.field('Issue'),
            Priority: fd.field('Priority'),          
            DisciplineMembers: fd.field('DisciplineMembers'),
            IssueDescription: fd.field('IssueDescription'),
            Attachments: fd.field('Attachments'),
            Author: fd.field('Author')
        }         
                  
        formatingButtonsBar(formBarTitle);

        formatingButtonsBar(formBarTitle);      
        [_formFields.DisciplineMembers, _formFields.Title, _formFields.Trade, _formFields.Issue, _formFields.Priority, _formFields.Attachments].forEach(disableField);
        disableRichTextFieldColumn(_formFields.IssueDescription);

        let buttons = fd.toolbar.buttons;
        let submitButton = buttons.find(button => button.text === 'Submit'); 
        disableEnableButton(submitButton, true);

        [_formFields.Author].forEach(hideField);  

        debugger;

        let Author = _formFields.Author.value.displayName;
        let TeamMembers = _formFields.DisciplineMembers.value;
        let myName = _currentUser.Title;
        _isAdmin = TeamMembers.some(user => user.DisplayText.toLowerCase() === myName.toLowerCase());
        _isNormal = myName.toLowerCase() === Author.toLowerCase();     

        if (_isAdmin || _isNormal) {
            setPSErrorMesg(`Edit Form`);           
        }
        else {    
            $('.divGrid').hide();
            setTimeout(function() {
                alert("Apologies, you have accessed a link that is not related to you.");  // Shows the alert after a short delay
            }, 50);  // 0 milliseconds delay, but lets the hide take effect before showing the alert
            fd.close();
        }
        
        //_isSendEmail = true;
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }        
}
var MDPHub_displayForm = async function(){   
    try 
    {
        _formFields = {     
            DisciplineMembers: fd.field('DisciplineMembers'), 
            Author: fd.field('Author')
        } 

        setPSHeaderMessage('');               
           
        formatingButtonsBar(formBarTitle);

        [_formFields.Author].forEach(hideField);
        
        debugger;

        let Author = _formFields.Author.value.displayName;
        let TeamMembers = _formFields.DisciplineMembers.value;
        let myName = _currentUser.Title;
        _isAdmin = TeamMembers.some(user => user.displayName.toLowerCase() === myName.toLowerCase());
        _isNormal = myName.toLowerCase() === Author.toLowerCase();     

        if (_isAdmin || _isNormal) {
            setPSErrorMesg(`Display Form`);           
        }
        else {    
            $('.divGrid').hide();
            setTimeout(function() {
                alert("Apologies, you have accessed a link that is not related to you.");  // Shows the alert after a short delay
            }, 50);  // 0 milliseconds delay, but lets the hide take effect before showing the alert
            fd.close();
        } 
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }          
}


var Submit_newForm = async function(){
    try 
    {        
        setPSHeaderMessage('');
        setPSErrorMesg(`Please fill in the below fields.`);

        _formFields = {
            Trade: fd.field('Trade'),
            DisciplineMembers: fd.field('DisciplineMembers'),
            ITUsers: fd.field('ITUsers')
        }

        formatingButtonsBar(formBarTitle);
    
        fd.field('Attachments').$on('change', validateAttachments);        
      
        [_formFields.DisciplineMembers, _formFields.ITUsers].forEach(hideField);

        _formFields.Trade.$on('change', async function (value) {              
            let AllTeamMembers = await fillinDisciplineMembers(value);
            let TeamMembers = AllTeamMembers.toUsers;
            let ITTeamMembers = AllTeamMembers.itUsers;
            $(_formFields.DisciplineMembers.$parent.$el).show();
            $(_formFields.ITUsers.$parent.$el).show();            
            _formFields.DisciplineMembers.value = TeamMembers?.map(user => user.Title) ?? '';
            _formFields.ITUsers.value = ITTeamMembers?.map(user => user.Title) ?? ''; 
            $(_formFields.DisciplineMembers.$parent.$el).hide();   
            $(_formFields.ITUsers.$parent.$el).hide();
        });       
        
        _isSendEmail = true;   
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    } 
}
var Submit_editForm= async function(){
    try 
    {
        setPSHeaderMessage('');
        //setPSErrorMesg(`Please review the below entrace.`);                
        
        _formFields = {
            Title: fd.field('Title'),
            Trade: fd.field('Trade'),
            ToolSource: fd.field('ToolSource'),
            ToolType: fd.field('ToolType'),
            Category: fd.field('Category'),
            SubCategory: fd.field('SubCategory'),
            DisciplineMembers: fd.field('DisciplineMembers'),
            ITUsers: fd.field('ITUsers'),
            ToolDescription: fd.field('ToolDescription'),
            Status: fd.field('Status'),
            Decision: fd.field('Decision'),
            ReviewedDate: fd.field('ReviewedDate'),
            ReasonofRejection: fd.field('ReasonofRejection'),
            Created: fd.field('Created'),
            Author: fd.field('Author')
        }

        //fd.toolbar.buttons[0].style = "display: none;";
        //fd.toolbar.buttons[1].style = "display: none;";   
        //document.querySelector(".divGrid").closest(".row").style.paddingLeft = "15px";
                
        formatingButtonsBar(formBarTitle);      
        [_formFields.DisciplineMembers, _formFields.Title, _formFields.Trade, _formFields.ToolSource, _formFields.ToolType, _formFields.Category, _formFields.SubCategory, _formFields.Status, _formFields.ReviewedDate].forEach(disableField);
        disableRichTextFieldColumn(_formFields.ToolDescription);
        [_formFields.ReasonofRejection, _formFields.Created, _formFields.Author, _formFields.ITUsers].forEach(hideField);    
  
        const CreatedDate = formatToLocalTimeZoneModern(_formFields.Created.value);
        let Author = _formFields.Author.value.displayName; 

        _Decision = _formFields.Decision.value.toLowerCase();
        
        let buttons = fd.toolbar.buttons;
        let submitButton = buttons.find(button => button.text === 'Submit');        

        _Status = _formFields.Status.value;
        let TeamMembers = _formFields.DisciplineMembers.value;
        let ITMembers = _formFields.ITUsers.value;
        let myName = _currentUser.Title;
        _isAdmin = TeamMembers.some(user => user.DisplayText.toLowerCase() === myName.toLowerCase());
        _isNormal = myName.toLowerCase() === Author.toLowerCase(); 
        let isITAdmin = ITMembers.some(user => user.DisplayText.toLowerCase() === myName.toLowerCase());

        //_isAdmin = false;

        if (_isAdmin || _isNormal || isITAdmin) { 

            if (_Status === 'Open') {            

                if (_isAdmin) { 
                    
                    setPSErrorMesg(`Kindly review the entry below, created on ${CreatedDate} by ${Author}.`);
                
                    _formFields.Decision.required = true;
                    _formFields.ReviewedDate.value = new Date();

                    _formFields.Decision.$on('change', async function (value) {
                        handleDecisionChange(value);
                    });

                    handleDecisionChange(_formFields.Decision.value);

                    _nextStatus = 'Reviewed';

                    _isSendEmail = true;
                }
                else {
                    setPSErrorMesg(`Your record, ${Author}, is currently under review and will be processed shortly.`);
                    [_formFields.Decision].forEach(disableField);
                    disableEnableButton(submitButton, true);
                }
            }
            else {
                setPSErrorMesg(`Your record, ${Author}, has been ${_Decision}`);
                [_formFields.Decision].forEach(disableField);
                disableEnableButton(submitButton, true);

                handleDecisionChange(_formFields.Decision.value);
                disableRichTextFieldColumn(_formFields.ReasonofRejection);
            }
        }
        else {    
            $('.divGrid').hide();
            setTimeout(function() {
                alert("Apologies, you have accessed a link that is not related to you.");  // Shows the alert after a short delay
            }, 50);  // 0 milliseconds delay, but lets the hide take effect before showing the alert
            fd.close();
        }            
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }         
}
var Submit_displayForm = async function(){  
    try 
    {
        setPSHeaderMessage('');        
        
         _formFields = {
            Title: fd.field('Title'),
            Trade: fd.field('Trade'),
            ToolSource: fd.field('ToolSource'),
            ToolType: fd.field('ToolType'),
            Category: fd.field('Category'),
            SubCategory: fd.field('SubCategory'),
            DisciplineMembers: fd.field('DisciplineMembers'),
            ITUsers: fd.field('ITUsers'),
            ToolDescription: fd.field('ToolDescription'),
            Status: fd.field('Status'),
            Decision: fd.field('Decision'),
            ReviewedDate: fd.field('ReviewedDate'),
            ReasonofRejection: fd.field('ReasonofRejection'),
            Created: fd.field('Created'),
            Author: fd.field('Author')
        }      
                
        formatingButtonsBar(formBarTitle);      
        
        [_formFields.ReasonofRejection, _formFields.Created, _formFields.Author, _formFields.ITUsers].forEach(hideField);  
      
        let Author = _formFields.Author.value.displayName; 

        _Decision = _formFields.Decision.value.toLowerCase();              

        _Status = _formFields.Status.value;
        let TeamMembers = _formFields.DisciplineMembers.value;
        let ITMembers = _formFields.ITUsers.value;
        let myName = _currentUser.Title;
        _isAdmin = TeamMembers.some(user => user.displayName.toLowerCase() === myName.toLowerCase());
        _isNormal = myName.toLowerCase() === Author.toLowerCase(); 
        let isITAdmin = ITMembers.some(user => user.displayName.toLowerCase() === myName.toLowerCase());

        if (_isAdmin || _isNormal || isITAdmin) {

            if (_Status === 'Reviewed')
                setPSErrorMesg(`Display Form - record for ${Author} is ${_Decision}`); 
            else
                setPSErrorMesg(`Display Form - record for ${Author} is under review`); 
            
            handleDecisionChange(_formFields.Decision.value);
         }
        else {    
            $('.divGrid').hide();
            setTimeout(function() {
                alert("Apologies, you have accessed a link that is not related to you.");  // Shows the alert after a short delay
            }, 50);  // 0 milliseconds delay, but lets the hide take effect before showing the alert
            fd.close();
        }        
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }   
}


var Feedback_newForm = async function(){
    try 
    {      
        formatingButtonsBar(formBarTitle);
        
        //document.querySelector(".divGrid").closest(".row").style.paddingLeft = "15px";
        //document.querySelector(".divGrid").closest(".row").style.paddingTop = "20px";
        let queryString = window.location.search;
        let source = true; 
        if (!queryString.includes("QueryVal")) {
        queryString = "?QueryVal=Clean%20Revit%20Files"; // Default query string
        source = false; // Update the source message
        }

        let actualQueryString = queryString.split('?').pop(); 

        let urlParams = new URLSearchParams(actualQueryString);
        let queryVal = urlParams.get('QueryVal');     
        queryVal = queryVal.replace(/^'|'$/g, '');

        fd.field("Title").value = queryVal;
        fd.field("Title").disabled = true;

        console.log("QueryVal:", queryVal);
        console.log("Source:", source);    
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }
}
var Feedback_editForm = async function(){
  try 
  {
      //fd.toolbar.buttons[0].style = "display: none;";
      //fd.toolbar.buttons[1].style = "display: none;";   
      //document.querySelector(".divGrid").closest(".row").style.paddingLeft = "15px";
      fd.field("Title").disabled = true;    
         
      formatingButtonsBar(formBarTitle);
  }    
  catch(err){
      console.log(err.message, err.stack)
      await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
      showPreloader();       
  }        
}
var Feedback_displayForm = async function(){     
    try 
    {
        //fd.toolbar.buttons[0].style = "display: none;";
        //fd.toolbar.buttons[1].style = "display: none;";
        //document.querySelector(".divGrid").closest(".row").style.paddingLeft = "15px";
            
        formatingButtonsBar(formBarTitle);
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }   
}


var Ticket_newForm = async function(){
    try 
    {             
        formatingButtonsBar("Qops Ticket Form");            
        
        //document.querySelector(".divGrid").closest(".row").style.paddingLeft = "15px";
        //document.querySelector(".divGrid").closest(".row").style.paddingTop = "20px";
        // Call validation when files are added
           //fd.field('Attachments').$on('change', validateAttachments);
  
                
            // await pnp.sp.web.currentUser.get().then(user => {
            //     DisplayName = user.Title;
            //     LoginName = user.LoginName.split('|')[1];
            // });             
           
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }     
}
var Ticket_editForm= async function(){
    try 
    {
        //fd.toolbar.buttons[0].style = "display: none;";
        //fd.toolbar.buttons[1].style = "display: none;";   
        //document.querySelector(".divGrid").closest(".row").style.paddingLeft = "15px";
    
        formatingButtonsBar("Qops Ticket Form");    
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }      
}
var Ticket_displayForm = async function(){ 
    try 
    {
        //fd.toolbar.buttons[0].style = "display: none;";
        //fd.toolbar.buttons[1].style = "display: none;";   
        //document.querySelector(".divGrid").closest(".row").style.paddingLeft = "15px";
        
        formatingButtonsBar("Qops Ticket Form");
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack); 
        showPreloader();       
    }
}


//#region General
var loadScripts = async function(withSign){

    const libraryUrls = [
      _layout + '/controls/preloader/jquery.dim-background.min.js',
      _layout + "/plumsail/js/buttons.js",
      _layout + '/plumsail/js/customMessages.js',
      _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
      _layout + '/plumsail/js/preloader.js',
      _layout + '/plumsail/js/commonUtils.js',
      _layout + '/plumsail/js/utilities.js'
    ];

    const cacheBusting = `?t=${Date.now()}`;
      libraryUrls.map(url => {
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`);
        });

    const stylesheetUrls = [
      _layout + '/controls/tooltipster/tooltipster.css',
      _layout + '/plumsail/css/DarTemplate.css' + `?t=${Date.now()}`
    ];

    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}${cacheBusting}">`);
    });
}

var extractValues = async function(moduleName, formType){

    //const startTime = performance.now();
    if($('.text-muted').length > 0)
      $('.text-muted').remove();

      _web = pnp.sp.web;
      _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
      _module = moduleName;
      _formType = formType;
      _webUrl = _spPageContextInfo.siteAbsoluteUrl;
      _siteUrl = new URL(_webUrl).origin;

    if(_formType === 'New')
        _isNew = true;

    else if(_formType === 'Edit'){
        _isEdit = true;
        _itemId = fd.itemId;
    }
    else if(_formType === 'Display')
        _isDisplay = true;

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;

    _currentUser = await pnp.sp.web.currentUser.get();
    //  const elapsedTime = endTime - startTime;
    //  console.log(`extractValues: ${elapsedTime} milliseconds`);
}
//#endregion

//#region Custom Buttons
var setCustomButtons = async function () {

    if (_formType !== "Display") {
        fd.toolbar.buttons[0].style = "display: none;";
        await setButtonActions("Accept", submitDefault, `${greenColor}`);
    }    
    fd.toolbar.buttons[1].style = "display: none;";       
    await setButtonActions("ChromeClose", "Cancel", `${yellowColor}`);
    setToolTipMessages();

    //$('span').filter(function(){ return $(this).text() === submitText; }).parent().attr("disabled", "disabled");
}

const setButtonActions = async function (icon, text, bgColor) {
    
    fd.toolbar.buttons.push({
        icon: icon,
        class: 'btn-outline-primary',
        text: text,
        style: `background-color: ${bgColor}; color: white;`,
        click: async function () {  

            if (text === "Close" || text === "Cancel") {
                fd.close();         
                //window.location.href = "https://4ce22.dar.com/sites/QOps/SitePages/HomePage.aspx?web=1";               
            }
            else if (text === "Submit") {

                if (fd.isValid) {                    
                    showPreloader();
                    if (_nextStatus) 
                        _formFields.Status.value = _nextStatus;                     
                    fd.save();            
                }
            }
        }
    });   
}

function setToolTipMessages(){

  let finalizetMesg = `Click ${submitDefault} for Manager Approval`;

  setButtonCustomToolTip(submitDefault, finalizetMesg);
  setButtonCustomToolTip('Close', closeMesg);

	if($('p').find('small').length > 0)
    $('p').find('small').remove();
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
    }); 

    document.querySelector('.col-sm-12').style.setProperty('padding-top', '0px', 'important'); 
    $('.col-sm-12').attr("style", "display: block !important;justify-content:end;");   
    $('.fd-grid.container-fluid').attr("style", "margin-top: -15px !important; padding: 10px;");
    $('.fd-form-container.container-fluid').attr("style", "margin-top: -10px !important;");   

    const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';
    const linkElement = `<a href="${_spPageContextInfo.webAbsoluteUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">
                          <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">${titelValue}</a>`;
    $('span.o365cs-nav-brandingText').html(linkElement);

    $('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');

    $('.border-title').each(function() {
        $(this).css({          
            'margin-top': '-35px', /* Adjust the position to sit on the border */
            'margin-left': '20px', /* Align with the content */            
        });
    });
}

function setIconSource(elementId, iconFileName) {

  const iconElement = document.getElementById(elementId);

  if (iconElement) {
      iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
  }
}

var fillinDisciplineMembers = async function(Trade){

    const listName = 'DisciplineMatrix';//list.Title;

    // const camlFilter = `<View>                            
    //                         <Query>
    //                             <Where>
    //                                 <Eq><FieldRef Name='Title'/><Value Type='Text'>${Trade}</Value></Eq>
    //                             </Where>
    //                         </Query>
    //                     </View>`;
    
    // const existingItems = await pnp.sp.web.lists
    //     .getByTitle(listName)
    //     .getItemsByCAMLQuery(
    //         { ViewXml: camlFilter },
    //         selectedColumns,
    //         expandColumns
    // );

    let selectedColumns = ['ToUsers/Id', 'ToUsers/Title', 'ITUsers/Id', 'ITUsers/Title'];
    let expandColumns = ['ToUsers', 'ITUsers'];
    let filterColumns = `Title eq '${Trade}'`;   

    const existingItems = await pnp.sp.web.lists.getByTitle(listName).items
        .select(...selectedColumns)
        .expand(...expandColumns)
        .filter(filterColumns)
        .getAll();
    
    return {
        toUsers: existingItems[0].ToUsers,
        itUsers: existingItems[0].ITUsers
    };
}

function disableRichTextFieldColumn(field){

    let elem = $(field.$el);//$(fd.field(fieldname).$el).find('.k-editor tr');

	elem.each(function(index, element){	

		let iframe = $(element).find('iframe');

		if(iframe.length > 0){

			let content = iframe.contents();
			let divElement = content.find('div');

			var lblElement = $('<label>', {
			  for: 'inputField',
			}).html(divElement.html());

			if(divElement.length === 0){
				lblElement = $('<label>', {
					for: 'inputField',
				  }).html(content[0].activeElement.innerHTML);
			}

			lblElement.css({
				'padding-top': '6px',
				'padding-bottom': '6px',
				'padding-left': '12px',
				'background-color': '#e9ecef',
				'width': '100%',
				'border-radius': '4px'
			});

			let tblElement = iframe.parent().parent().parent().parent();
			tblElement.parent().append(lblElement);
			tblElement.remove();
		}	
	})
}

function formatToLocalTimeZoneModern(dateString) {
    const date = new Date(dateString);

    // Get the time zone offset in hours
    const timeZoneOffset = -date.getTimezoneOffset() / 60;
    const formattedOffset = `UTC${timeZoneOffset >= 0 ? '+' : ''}${timeZoneOffset}`;

    return date.toLocaleString(undefined, {
        month: 'short',
        day: 'numeric',
        year: 'numeric',
        hour: 'numeric',
        minute: '2-digit',
        hour12: true
    }) + ` (${formattedOffset})`;
}

function disableEnableButton(submitButton, isDisabled, disableColor) {
    if (isDisabled) {
        submitButton.disabled = true;       
        disableColor = disableColor || '#97dccd';
        submitButton.style = `background-color:${disableColor}; color:white;`;
    }
    else {
        submitButton.disabled = false;
        submitButton.style = `background-color: #5FC9B3; color:white;`;
    }   
}

function handleDecisionChange(value) {
    if (value) {
        _Decision = value.toLowerCase();
        if (_Decision === 'rejected') {
            [_formFields.ReasonofRejection].forEach(showField);
            _formFields.ReasonofRejection.required = true;
        }
        else {
            _formFields.ReasonofRejection.required = false;
            _formFields.ReasonofRejection.value = '';
            [_formFields.ReasonofRejection].forEach(hideField);
        }
    }
}
//#endregion

fd.spSaved(async function(result) {

    try {
        
        _itemId = result.Id;
        let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${_itemId}</Value></Eq></Where>`;

        if (_isSendEmail) {

            let listPath = `${fd.listUrl}`;
            let listInternalName = listPath.replace("/Lists/", "").trim();

            let editTaskLink = _spPageContextInfo.webAbsoluteUrl + `/SitePages/PlumsailForms/${listInternalName}/Item/EditForm.aspx?item=` + _itemId;
            let displayTaskLink = _spPageContextInfo.webAbsoluteUrl + `/SitePages/PlumsailForms/${listInternalName}/Item/DisplayForm.aspx?item=` + _itemId;
            
            let Subject = `New "${_list}" Record Added by ${_currentUser.Title} - Please Review.`;
            let bodyMessage = `<p>A new record has been successfully added.</p> 
                               <p>You may click <a href="${editTaskLink}" class="link">here</a> to review it at your convenience.</p>`;
            
            if (_module === 'MDPHub') {
                bodyMessage = `<p>A new record has been successfully added.</p>
                               <p>Please click <a href="${displayTaskLink}" class="link">here</a> to view the details.</p>`;
            }
            
            if (_formType == 'Edit') {
                
                if (_module === 'Submit') {

                    if (_Decision === 'approved') {
                        bodyMessage = `<p>The record has been ${_Decision}.</p> 
                               <p>Please click <a href="${displayTaskLink}" class="link">here</a> to view the details and add the tool to DarMDPHub.</p>`;
                    }
                    else {
                        bodyMessage = `<p>The record has been ${_Decision}.</p> 
                               <p>Please click <a href="${displayTaskLink}" class="link">here</a> to view the details.</p>`;
                    }

                }
                else {
                    bodyMessage = `<p>Your record has been reviewed.</p> 
                               <p>Please click <a href="${displayTaskLink}" class="link">here</a> to view the details.</p>`;
                }
                Subject = `Your "${_list}" Record Review Completed.`;

            }
            
            let BodyEmail = `<html>
                            <head>
                                <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
                                <style>
                                    body {
                                        font-family: Arial, sans-serif;
                                        background-color: #f4f4f4;
                                        margin: 0;
                                        padding: 0;
                                        font-size: 15px;
                                        color: #333333;
                                    }
                                    .container {
                                        width: 100%;                         
                                        margin: 0 auto;
                                        padding: 10px;
                                        background-color: #fdfdfd;
                                    }
                                    .header {
                                        background-color: #3bdbc0;
                                        color: #ffffff;
                                        padding: 12px 20px;
                                        font-size: 16px;
                                        font-weight: bold;
                                        text-align: center;                            
                                    }
                                    .content {
                                        padding: 10px;
                                        font-size: 14px;
                                        line-height: 1.6;
                                        color: #555555;
                                    }
                                    .content p {
                                        margin: 5px 0;
                                    }
                                    .footer {
                                        text-align: center;
                                        color: #888888;
                                        font-size: 12px;
                                        padding: 10px 0;
                                    }
                                    .action-button {
                                        display: inline-block;
                                        background-color: #4CAF50;
                                        color: #ffffff;
                                        padding: 10px 15px;
                                        border-radius: 4px;
                                        text-decoration: none;
                                        font-weight: bold;
                                        text-align: center;
                                        margin-top: 15px;
                                    }
                                    .action-button:hover {
                                        background-color: #45a049;
                                    }
                                </style>
                            </head>
                            <body>
                                <table class='container' cellpadding='0' cellspacing='0'>                                                                     
                                    <tr>
                                        <td class='content'>                                          
                                            <p>Dear All,</p>
                                            <br />                                            
                                            <p>${bodyMessage}</p>
                                            <br /> 
                                            <p>Best regards,</p>
                                            <br />                                     
                                            <img src='data:image/png;base64,base64Image' alt='Purchase Order Image'/>
                                        </td>
                                    </tr> 
                                    <tr>
                                        <td class='footer'>
                                            <p>This is an automated message. Please do not reply to this email.</p>
                                        </td>
                                    </tr>                                       
                                </table>
                            </body>
                        </html>`;
            
            let encodedSubject = htmlEncode(Subject);
            let encodedBodyEmail = htmlEncode(BodyEmail);
            
            if (_module === 'Submit')
                _module = 'Submit New Tool';

            else if (_module === 'MDPHub')
                _module = 'Report Bug';

            await _sendEmail(_module, encodedSubject + '|' + encodedBodyEmail, query, `${_formType}|${_Decision}`, '', '', _currentUser);
        }       
        
    } catch(e) {
        console.log(e);
    }
});

// #region Function to validate file types
function validateAttachments() {
    var attachments = fd.field('Attachments').value;
    var allowedExtensions = ['.zip', '.rar'];

    if (!attachments || attachments.length === 0) {
        return true; // No file attached, allow submission
    }

    for (var i = 0; i < attachments.length; i++) {
        var fileName = attachments[i].name.toLowerCase();
        var isValid = allowedExtensions.some(ext => fileName.endsWith(ext));

        if (!isValid) {
            alert("Only .zip or .rar files are allowed!");
            fd.field('Attachments').clear(); // Remove invalid files
            return false;
        }
    }
    return true;
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
//#endregion

