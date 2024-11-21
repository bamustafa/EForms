const blueColor = '#6ca9d5', greenColor = '#5FC9B3', yellowColor = '#F7D46D', redColor = '#F28B82';

function setIconSource(elementId, iconFileName) {

    const iconElement = document.getElementById(elementId);

    if (iconElement) {
        iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
    }
}

async function loadingButtons(status){    
    
    fd.toolbar.buttons.push({
        icon: 'Accept',
        class: 'btn-outline-primary',
        disabled: false,
        text: 'Submit',
        style: `background-color:${greenColor}; color:white`,
        click: async function() {  	
            if(fd.isValid){

                preloader();

                _proceed = false;

                if(_formType === 'New'){
                    _Email = 'RFCNewEntry_Email';
                    _Notification = 'RFC_Initiated';
                }
                else if(_formType === 'Edit') {
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
            preloader();
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
}



fd.spSaved(async function(result) {			
    // try {        
        
    //     if(_proceed){		
    //         var itemId = result.Id;        
    //         let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where>`;
    //         await _sendEmail(_modulename, _Email, query, '', _Notification, '', CurrentUser);    
    //     }  

    // } catch(e) {
    //     console.log(e);
    // }								 
});