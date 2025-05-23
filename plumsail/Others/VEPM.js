var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDisplay = false;

var _currentUser, _formFields = {}, _emailFields = {};
const greenColor = '#5FC9B3', redColor = '#F28B82', yellowColor = '#6c757d', saveColor = '#F7D46D';

const itemsToRemove = ['WorkflowStatus', 'Discipline', 'Reference'];
let _nextStatus = '', _schema = '', _emailBody = '', _VEPSchema = '';
let _status = 'IDEA Submitted', _Originator = '', _Reference = '', _VEPReference = '',_code = '', _reviewFlow = '', _ideaOrigin = '', _parentID = '', _titleforEmail = '', ProjectNumber = '';
let _notReadyButtonName = 'Save Draft', _SelectCompanyForReview = '', _companyType = '', _companyGroupName = '';

let groupMap = {};

var onRender = async function (relativeLayoutPath, moduleName, formType) {

    try{

        _layout = relativeLayoutPath;

        _formFields = {
            Title: fd.field('Title'),
            Reference: fd.field('Reference'),
            Discipline: fd.field('Discipline'),
            Originator: fd.field('Originator'),
            WorkflowStatus: fd.field('WorkflowStatus'),
            Question: fd.field('Description'),
            Category: fd.field('Category'),
            ProjectComponent: fd.field('ProjectComponent'),              
            Notes: fd.field('Notes'), 
            Code: fd.field('Code'),
            ReasonofRejection: fd.field('ReasonofRejection'),
            Importance: fd.field('Importance'),
            Satisfaction: fd.field('Satisfaction'),
            RoughOrder: fd.field('RoughOrder'),
            IdeaOrigin: fd.field('IdeaOrigin'),

            UniformatLevel1: fd.field('UniformatLevel1'),
            UniformatLevel2: fd.field('UniformatLevel2'),
            UniformatLevel3: fd.field('UniformatLevel3'), 
            
            VEPReference: fd.field('VEPReference'),
            OriginalDesign: fd.field('OriginalDesign'),
            ProposedDesign: fd.field('ProposedDesign'),
            PerformanceBenefits: fd.field('PerformanceBenefits'),
            CoordinationRequired: fd.field('CoordinationRequired'),
            CoordinationCompleted: fd.field('CoordinationCompleted'),
            OriginalCost: fd.field('OriginalCost'),
            ProposedCost: fd.field('ProposedCost'),
            Savings: fd.field('Savings'),
            TimeImpact: fd.field('TimeImpact'),
            ReviewedDate: fd.field('ReviewedDate'),
            SelectCompanyForReview: fd.field('SelectCompanyForReview'),
            Comments: fd.field('Comments'),
        
            Attachments: fd.field('Attachments')           
        } 
        
        _emailFields = {
            'IDEA No.': fd.field('Reference'),
            Company: fd.field('Originator'),
            Trade: fd.field('Discipline'),   
            'Sub-Project': fd.field('Category'), 
            'Scope': fd.field('ProjectComponent'),            
            'Uniformat Level 1': fd.field('UniformatLevel1'),
            'Uniformat Level 2': fd.field('UniformatLevel2'),
            'Uniformat Level 3': fd.field('UniformatLevel3'),
            'IDEA Title': fd.field('Title'),                       
            'IDEA Description': fd.field('Description'),
            Notes: fd.field('Notes'),            
            WorkflowStatus: fd.field('WorkflowStatus'),
            Code: fd.field('Code'),
            'Reason of Rejection': fd.field('ReasonofRejection')
        }

        const Generaliconsvg = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
        setIconSource("General-icon", Generaliconsvg);        
        const VPMiconsvg = `<svg width="26" height="26" viewBox="0 0 48 48" version="1" xmlns="http://www.w3.org/2000/svg" enable-background="new 0 0 48 48" fill="#000000">
        <g id="SVGRepo_bgCarrier" stroke-width="0"/>
        <g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"/>
        <g id="SVGRepo_iconCarrier"> <path fill="#7b9c98" d="M37.4,24.6l-11.6-2.2l-3.9-11.2l-3.8,1.3L22,23.6l-7.8,9l3,2.6l7.8-9l11.6,2.2L37.4,24.6z"/> <g fill="#7bd8cc"> <path d="M24,19c-2.8,0-5,2.2-5,5c0,2.8,2.2,5,5,5s5-2.2,5-5C29,21.2,26.8,19,24,19z M24,26c-1.1,0-2-0.9-2-2 c0-1.1,0.9-2,2-2s2,0.9,2,2C26,25.1,25.1,26,24,26z"/> <path d="M40.7,27c0.2-1,0.3-2,0.3-3c0-1-0.1-2-0.3-3l3.3-2.4c0.4-0.3,0.6-0.9,0.3-1.4L40,9.8 c-0.3-0.5-0.8-0.7-1.3-0.4L35,11c-1.5-1.3-3.3-2.3-5.2-3l-0.4-4.1c-0.1-0.5-0.5-0.9-1-0.9h-8.6c-0.5,0-1,0.4-1,0.9L18.2,8 c-1.9,0.7-3.7,1.7-5.2,3L9.3,9.3C8.8,9.1,8.2,9.3,8,9.8l-4.3,7.4c-0.3,0.5-0.1,1.1,0.3,1.4L7.3,21C7.1,22,7,23,7,24 c0,1,0.1,2,0.3,3L4,29.4c-0.4,0.3-0.6,0.9-0.3,1.4L8,38.2c0.3,0.5,0.8,0.7,1.3,0.4L13,37c1.5,1.3,3.3,2.3,5.2,3l0.4,4.1 c0.1,0.5,0.5,0.9,1,0.9h8.6c0.5,0,1-0.4,1-0.9l0.4-4.1c1.9-0.7,3.7-1.7,5.2-3l3.7,1.7c0.5,0.2,1.1,0,1.3-0.4l4.3-7.4 c0.3-0.5,0.1-1.1-0.3-1.4L40.7,27z M24,35c-6.1,0-11-4.9-11-11c0-6.1,4.9-11,11-11s11,4.9,11,11C35,30.1,30.1,35,24,35z"/></g></g></svg>`;
        setIconSource("VPM-icon", VPMiconsvg);
        const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;
        setIconSource("Attachments-icon", svgattachment);        
        
        await loadScripts().then(async () => {

            showPreloader();
            await extractValues(moduleName, formType);

            if (moduleName === 'VEPM') {

                if (_isEdit) 
                    await handleEditForm();                
            
                else if (_isNew)
                    await handleNewForm();

                else if (_isDisplay)
                    await handleDisplayForm();
            }
            else if (moduleName === 'VEPMTasks') {

                if (_isEdit)
                    await handleTasksEditForm(); 
                else if (_isDisplay)
                    await handleDisplayForm();
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

var handleNewForm = async function () {

    clearLocalStorageItemsByField(itemsToRemove);   
    
    await setCustomButtons();

    let buttons = fd.toolbar.buttons;
    let SaveButton = buttons.find(button => button.text === _notReadyButtonName); 
    if (SaveButton) 
        SaveButton.visible = false;  

    formatingButtonsBar("IDEA Register Form"); 

    setPSHeaderMessage('', '-25px');
    setPSErrorMesg(`Please complete all required fields and select a satisfaction and importance value from the matrix.`); 
    hideHideSection(true); 

    const matrixbtn = document.getElementById("open-matrix-btn");      
    if (matrixbtn) {
        toggleToolbarButton('Submit', true);
        await loadMatrix(matrixbtn);
    }

    const spaceElement = document.getElementById("spaceID");
    if (spaceElement) {
        spaceElement.style.display = "none";
    }

    const divGridVPMEl = document.querySelectorAll('.divGrid.VPM');
    divGridVPMEl.forEach(el => {
        el.style.display = 'none';
    });

    //_HideFormFields([_formFields.Answer], true);
    _setFieldsDisabled([_formFields.Reference, _formFields.Originator], true);  
    _HideFormFields([_formFields.Code, _formFields.ReasonofRejection, _formFields.SelectCompanyForReview], true);

    _Originator = await CheckifUserinSPGroup();
    _formFields.Originator.value = _Originator;
    _formFields.IdeaOrigin.value = _Originator;
    
    //Title Limit Character Counter  
    var titleField = _formFields.Title;
    const maxLength = titleField.maxLength;
    var counter = document.createElement('div');
    counter.id = 'charCount';
    counter.style.marginTop = '2px';
    counter.style.fontSize = '11px';
    counter.style.color = '#6b7280';
    counter.style.fontStyle = 'italic';
    counter.style.marginBottom = '-15px';
    titleField.$el.appendChild(counter);
    if (counter) {
        const updateCounter = (value) => {
            const remaining = maxLength - value.length;
            const label = remaining <= 1 ? 'character' : 'characters';
            counter.textContent = `${remaining} ${label} left`;
            if(remaining <= maxLength*0.2)
                counter.style.color = '#ce310e';             
            else
                counter.style.color = '#6b7280';
        };
        updateCounter(titleField.value);
        titleField.$on('change', async function (value) {
            updateCounter(value);
        });
    }


    // Create the display div
    let numberField = _formFields.RoughOrder; // replace with your actual field     
    let formatDiv = document.createElement('div');
    formatDiv.id = 'formattedNumber';
    formatDiv.style.marginTop = '2px';
    formatDiv.style.fontSize = '11px';
    formatDiv.style.color = '#6b7280';
    formatDiv.style.fontStyle = 'italic';
    formatDiv.style.marginBottom = '-20px';
    numberField.$el.appendChild(formatDiv);
    updateFormatted(formatDiv, numberField.value);
    let origInput = $(_formFields.RoughOrder.$el).find('.k-input');
    origInput.on('input', (event) => {
        let thisVal = event.target.value;
        updateFormatted(formatDiv, thisVal);
    });
    
    
    //await HandleCascadedLookup();

    //await HandleCascadedUniformatLevels();    

    await HandleUnifiedCascade();

    _schema = await getAutReferenceFormat();
    
    _formFields.Reference.value = await parseRefFormat(_schema);

    _nextStatus = 'IDEA Submitted';   
}

var handleEditForm = async function () {

    clearLocalStorageItemsByField(itemsToRemove);
    
    await setCustomButtons();

    let buttons = fd.toolbar.buttons;
    let SaveButton = buttons.find(button => button.text === _notReadyButtonName); 
    if (SaveButton) 
        SaveButton.visible = false; 

    formatingButtonsBar("VEPM Form"); 
    setPSHeaderMessage('', '-25px');
    hideHideSection(true);

    _setFieldsDisabled([_formFields.Title, _formFields.Reference, _formFields.Discipline, _formFields.Originator, _formFields.Category, _formFields.ProjectComponent,
    _formFields.UniformatLevel1, _formFields.UniformatLevel2, _formFields.UniformatLevel3, _formFields.RoughOrder], true); 
    disableRichTextFieldColumn(_formFields.Question);

    _HideFormFields([_formFields.SelectCompanyForReview], true);
    
    _Originator = _formFields.Originator.value;
    _status = _formFields.WorkflowStatus.value; 
    _Reference = _formFields.Reference.value;    
    
    const matrixbtn = document.getElementById("open-matrix-btn"); 
    const divGridVPMEl = document.querySelectorAll('.divGrid.VPM');
    
    let GroupName = await CheckifUserinGroup(); 
    console.log("GroupName: ", GroupName);    

    const companyDetails = await GetCompanyDetails(_Originator);
    _companyType = companyDetails.CompanyType;
    _companyGroupName = companyDetails.SpGroup;

    const transitions = {
        'IDEA Submitted': {
            roles: ['VEPM_Reviewer', 'Client', 'Admin'],
            nextStatus: {
                'Rejected': {
                    status: 'Rejected',
                    showReasonField: true
                },
                'Approved': {
                    status: 'Request for Details',
                    showReasonField: false
                }
            },
            errorMessage: 'Kindly review the current form'           
        },
        'Request for Details': {
            roles: [_companyGroupName, 'Client', 'VEPM_Reviewer', 'Admin'],
            nextStatus: 'Details Submitted',
            errorMessage: 'Please provide the required idea details'
        },
        'Revise After Rejection': {
            roles: [_companyGroupName, 'Client', 'VEPM_Reviewer', 'Admin'],
            nextStatus: 'Details Submitted',
            errorMessage: 'Please provide the required idea details'
        },
    };

    const transition = transitions[_status]; //ok 2

    if (transition && transition.roles.includes(GroupName)) {        

        setPSErrorMesg(`${transition.errorMessage}. The current status is marked as "${_status}".`);
        _HideFormFields([_formFields.ReasonofRejection], true);

        if (_status === 'IDEA Submitted') {

            _setFieldsDisabled(_formFields.Code, true);
            
            const spaceElement = document.getElementById("spaceID");
            if (spaceElement) {
                spaceElement.style.display = "none";
            }
            
            divGridVPMEl.forEach(el => {
                el.style.display = 'none';
            });

            let x = _formFields.Importance.value;
            let y = _formFields.Satisfaction.value;

            if (x >= 6 && y >= 6)
                _formFields.Code.value = 'Approved';
            else
                _formFields.Code.value = 'Rejected';

            await updateSubmitButtonState(transition);

            if (matrixbtn) {        
                await loadMatrix(matrixbtn, transition);                
            }

            if (_companyType === 'Client' || _companyType === 'PMC') {
                _HideFormFields([_formFields.SelectCompanyForReview], false);
                _isRequiredFields([_formFields.SelectCompanyForReview], true);

                await _formFields.SelectCompanyForReview.ready();
                _formFields.SelectCompanyForReview.filter = `CompanyType eq 'Consultant' or CompanyType eq 'Contractor'`;
                _formFields.SelectCompanyForReview.refresh();                               
            }            
        }
        else if (_status === 'Request for Details') {         

            _nextStatus = transition.nextStatus;
            _setFieldsDisabled([_formFields.Code, _formFields.Discipline, _formFields.VEPReference, _formFields.Savings, _formFields.SelectCompanyForReview, _formFields.ReviewedDate], true);
            disableRichTextFieldColumn(_formFields.Notes);           

            if (matrixbtn) {        
                await readonlyloadMatrix(matrixbtn);
            }

            _HideFormFields([_formFields.Notes, _formFields.CoordinationCompleted], true);

            // const matrixdialog = document.getElementById('matrix-dialog-wrapper'); 
            // if (matrixdialog)
            //     matrixdialog.style.display = 'none';         
          
            _VEPReference = _formFields.VEPReference.value;
            _VEPSchema = await getAutReferenceFormat(true);

            if (_VEPReference === '') {                
                _formFields.VEPReference.value = await parseRefFormat(_VEPSchema);
                _formFields.ReviewedDate.value = new Date();
            }   
            
            _isRequiredFields([_formFields.OriginalDesign, _formFields.ProposedDesign, _formFields.PerformanceBenefits, _formFields.OriginalCost,
            _formFields.ProposedCost, _formFields.TimeImpact], true);

            let origInput = $(_formFields.OriginalCost.$el).find('.k-input');
            let propInput = $(_formFields.ProposedCost.$el).find('.k-input');

            updateSavings(origInput.val(), propInput.val());

            origInput.on('input', (event) => {           
                const thisVal = event.target.value;
                updateSavings(thisVal, propInput.val());
            });

            propInput.on('input', (event) => { 
                const thisVal = event.target.value;
                updateSavings(origInput.val(), thisVal);
            }); 
            
            await toggleCoordinationCompleted(_formFields.CoordinationRequired.value);      
            _formFields.CoordinationRequired.$on('change', async function (value) {
                await toggleCoordinationCompleted(value);
            });
            
            // _setFieldsDisabled([_formFields.UniformatLevel3], false); 
            // await HandleLevel3CascadedLookup();           
          
            let IdeaOrigin = _formFields.IdeaOrigin.value;

            if (IdeaOrigin === 'DAR' || IdeaOrigin === 'UDL') {
                _setFieldsDisabled([_formFields.Category, _formFields.ProjectComponent, _formFields.UniformatLevel1, _formFields.UniformatLevel2, _formFields.UniformatLevel3], false);
                // await HandleCascadedLookup();
                // await HandleCascadedUniformatLevels();

                await HandleUnifiedCascade();
            }
        }
        else if (_status === 'Revise After Rejection') {          

            _nextStatus = transition.nextStatus;
            _setFieldsDisabled([_formFields.Code, _formFields.VEPReference, _formFields.Savings, _formFields.SelectCompanyForReview, _formFields.ReviewedDate], true);
            disableRichTextFieldColumn(_formFields.Notes);

            if (matrixbtn) {        
                await readonlyloadMatrix(matrixbtn);
            }

            _HideFormFields([_formFields.UniformatLevel1, _formFields.UniformatLevel2, _formFields.UniformatLevel3, _formFields.Category,
            _formFields.ProjectComponent, _formFields.Notes, _formFields.CoordinationCompleted], true);

            const matrixdialog = document.getElementById('matrix-dialog-wrapper'); 
            if (matrixdialog)
                matrixdialog.style.display = 'none';         
          
            _VEPReference = _formFields.VEPReference.value;              
            
            _isRequiredFields([_formFields.OriginalDesign, _formFields.ProposedDesign, _formFields.PerformanceBenefits, _formFields.OriginalCost,
            _formFields.ProposedCost, _formFields.TimeImpact], true);

            let origInput = $(_formFields.OriginalCost.$el).find('.k-input');
            let propInput = $(_formFields.ProposedCost.$el).find('.k-input');

            updateSavings(origInput.val(), propInput.val());

            origInput.on('input', (event) => {           
                const thisVal = event.target.value;
                updateSavings(thisVal, propInput.val());
            });

            propInput.on('input', (event) => { 
                const thisVal = event.target.value;
                updateSavings(origInput.val(), thisVal);
            }); 
            
            await toggleCoordinationCompleted(_formFields.CoordinationRequired.value);      
            _formFields.CoordinationRequired.$on('change', async function (value) {
                await toggleCoordinationCompleted(value);
            });
        }
        else {            
            _nextStatus = transition.nextStatus;
            _setFieldsDisabled(_formFields.Code, true);
            disableRichTextFieldColumn(_formFields.Notes);

            if (matrixbtn) {        
                await readonlyloadMatrix(matrixbtn);
            }

            disableVEPSection();

            if (_formFields.VEPReference.value) {
                
                divGridVPMEl.forEach(el => {
                    el.style.display = 'none';
                });                
            }
        }
        
    } else {
        setPSErrorMesg(`This section displays the submitted Idea and its details. Current status: '${_status}'.`);   
        _setFieldsDisabled([_formFields.Code], true);
        disableRichTextFieldColumn(_formFields.ReasonofRejection);
        disableRichTextFieldColumn(_formFields.Notes);
        // $('span').filter(function () {
        //     return $(this).text() === submitDefault;
        // }).parent().attr("disabled", "disabled");
        toggleToolbarButton('Submit', true);

        SetAttachmentToReadOnly();

        if (matrixbtn) {        
            await readonlyloadMatrix(matrixbtn);
        } 
        
        disableVEPSection();
    }
}

var handleDisplayForm = async function () {
    await setCustomButtons();

    let buttons = fd.toolbar.buttons;
    let SaveButton = buttons.find(button => button.text === _notReadyButtonName); 
    if (SaveButton) 
        SaveButton.visible = false; 

    formatingButtonsBar("RFI Form");
    setPSHeaderMessage('', '-25px');
    setPSErrorMesg(`Please note that the information below pertains to the Idea details. Click 'Edit' to make changes.`);
    hideHideSection(true);

    const matrixbtn = document.getElementById("open-matrix-btn");  
    
    if (matrixbtn) {        
        await readonlyloadMatrix(matrixbtn);
    }

    let code = _formFields.Code.value; 

    _HideFormFields([_formFields.ReasonofRejection], true); 

    if (code && code === 'Rejected')
        _HideFormFields([_formFields.ReasonofRejection], false);
}

var handleTasksEditForm = async function () {   
    
    clearLocalStorageItemsByField(itemsToRemove);

    await setCustomButtons();

    let buttons = fd.toolbar.buttons;
    let SaveButton = buttons.find(button => button.text === _notReadyButtonName); 
    if (SaveButton) 
        SaveButton.visible = false; 

    formatingButtonsBar("VEP Task Form"); 
    setPSHeaderMessage('', '-25px');
    hideHideSection(true);    

    _setFieldsDisabled([_formFields.Title, _formFields.Reference, _formFields.Discipline, _formFields.Originator, _formFields.Category, _formFields.ProjectComponent], true); 
    disableRichTextFieldColumn(_formFields.Question);    

    _Originator = _formFields.Originator.value;
    _status = _formFields.WorkflowStatus.value; 
    _Reference = _formFields.Reference.value;
    _ideaOrigin = fd.field('IdeaOrigin').value;
    _parentID = fd.field('ParentID').value;
    _setFieldsDisabled([fd.field('IdeaOrigin')], true);

    _HideFormFields([_formFields.Notes], true); 
    
    let GroupName = await CheckifUserinGroup();    

    const transitions = {
        'Open': {
            roles: [_Originator, 'Admin'],
            nextStatus: {
                'Objection': {
                    status: 'Completed',
                    showReasonField: true
                },
                'No Objection': {
                    status: 'Completed',
                    showReasonField: false
                }
            }           
        }        
    };

    const transition = transitions[_status];

    if (transition && transition.roles.includes(GroupName)) {    

        setPSErrorMesg(`Kindly review the current form. The current status is marked as '${_status}'.`);
        _HideFormFields([_formFields.ReasonofRejection], true);

        if (_status === 'Open') {
            _setFieldsDisabled(transition.disableFields, true);
            await updateSubmitButtonState(transition);
            _formFields.Code.$on('change', async function (value) {
                await updateSubmitButtonState(transition);
            });
        }
        else {
            _nextStatus = transition.nextStatus;
            _setFieldsDisabled(_formFields.Code, true);
        }
        
    } else {

        setPSErrorMesg(`This section displays the submitted Idea and its details. Current status: '${_status}'.`);

        if (_status === 'Open')
            setPSErrorMesg(`Only ${_Originator} is allowed to respond. Please coordinate with them to take action.`);

        _setFieldsDisabled([_formFields.Code], true);
        disableRichTextFieldColumn(_formFields.ReasonofRejection);
        disableRichTextFieldColumn(_formFields.Notes);
        // $('span').filter(function () {
        //     return $(this).text() === submitDefault;
        // }).parent().attr("disabled", "disabled");

        toggleToolbarButton('Submit', true);

        SetAttachmentToReadOnly();
    }
}

//#region Core Functions

var loadMatrix = async function (btn, transition) {
    // Create the dialog container
    const dialog = document.createElement("div");
    dialog.id = "matrix-popup";
    dialog.style.position = "fixed";
    dialog.style.top = "50%";
    dialog.style.left = "50%";
    dialog.style.transform = "translate(-50%, -50%)";
    dialog.style.background = "#fff";
    dialog.style.border = "1px solid #ccc"; 
    dialog.style.paddingLeft = "30px";
    dialog.style.paddingBottom = "10px";
    dialog.style.paddingTop = "10px";
    dialog.style.paddingRight = "10px";
    dialog.style.zIndex = 9999;
    dialog.style.display = "none";
    dialog.style.boxShadow = "0 4px 8px rgba(0,0,0,0.2)";
    dialog.style.borderRadius = "8px";   

    // Add close (X) button
    const closeButton = document.createElement("span");
    closeButton.textContent = "×";
    closeButton.style.position = "absolute";
    closeButton.style.top = "5px";
    closeButton.style.right = "10px";
    closeButton.style.cursor = "pointer";
    closeButton.style.fontSize = "20px";
    closeButton.style.fontWeight = "bold";
    closeButton.style.color = "#888";
    closeButton.addEventListener("click", () => {
        dialog.style.display = "none";                
        hidePreloader();
    });
    dialog.appendChild(closeButton);

    // Table element
    const table = document.createElement("table");
    table.style.borderCollapse = "collapse";
    table.style.textAlign = "center";

    // Add X-axis labels row
    const labelRow = document.createElement("tr");

    // Add X-axis labels 1 to 10
    for (let x = 1; x <= 10; x++) {
        const label = document.createElement("td");
        label.textContent = x;
        label.style.textAlign = "center";
        label.style.fontSize = "12px";
        label.style.fontWeight = "bold";
        label.style.padding = "2px";
        label.style.border = "none";
        labelRow.appendChild(label);
    }

    table.appendChild(labelRow);    

    for (let y = 10; y >= 1; y--) {

        const row = document.createElement("tr");

        for (let x = 1; x <= 10; x++) {
            
            const cell = document.createElement("td");
            cell.setAttribute("data-x", x);
            cell.setAttribute("data-y", y);
            cell.style.width = "30px";
            cell.style.height = "30px";
            cell.style.border = "1px solid #ccc";
            cell.style.cursor = "pointer";
            cell.style.textAlign = "center";
            cell.title = `Importance: ${x}, Satisfaction: ${y}`;
            cell.style.backgroundColor = getColor(x, y);

            cell.addEventListener("click", async () => {
                // Save to Plumsail fields
                fd.field("Importance").value = x;
                fd.field("Satisfaction").value = y;
                let BKCol = getColor(x, y);
                fd.field("BKColor").value = BKCol; 
                
                toggleToolbarButton('Submit', false);
          
                dialog.style.display = "none";    
                hidePreloader();

                btn.textContent = `Selected: ${x}, ${y}`;
                btn.style.backgroundColor = BKCol; 
                btn.style.setProperty('color', 'white', 'important');
                btn.style.fontWeight = 'bold';

                if (_status === 'IDEA Submitted' && _isEdit)
                {
                    if (x >= 6 && y >= 6)
                        _formFields.Code.value = 'Approved';
                    else
                        _formFields.Code.value = 'Rejected';

                    await updateSubmitButtonState(transition);
                }
            });

            row.appendChild(cell);

            // If this is the last x (x = 10), add the label for Y
            if (x === 10) {
                const yLabel = document.createElement("td");
                yLabel.textContent = y;
                yLabel.style.fontSize = "12px";
                yLabel.style.fontWeight = "bold";
                yLabel.style.paddingLeft = "6px";
                yLabel.style.border = "none";
                row.appendChild(yLabel);
            }
        }

        table.appendChild(row);             
    }    
    
    // Create top container with flex for text left + table right
    const topRow = document.createElement('div');   

    // Left text
    const leftText = document.createElement('div');
    leftText.textContent = 'Importance/Cost';
    leftText.style.transform = "rotate(-90deg)";  
    leftText.style.fontWeight = 'bold';
    leftText.style.fontSize = '14px';
    leftText.style.position = 'absolute';
    leftText.style.left = '-32px';  // x coordinate
    leftText.style.top = '200px';  // y coordinate    

    // Append left text and table to top row
    topRow.appendChild(leftText);
    topRow.appendChild(table);

    // Bottom centered text
    const bottomText = document.createElement('div');
    bottomText.textContent = 'Satisfaction/ Performance';
    bottomText.style.textAlign = 'center';
    bottomText.style.fontWeight = 'bold';
    bottomText.style.fontSize = '14px';
    bottomText.style.marginTop = '5px';

    // Append top row and bottom text to dialog
    dialog.appendChild(topRow);
    dialog.appendChild(bottomText);

    document.body.appendChild(dialog);

    // Open matrix on button click
    btn.addEventListener("click", () => {
        dialog.style.display = "block";
        showPreloader();

        const savedX = fd.field("Importance").value;
        const savedY = fd.field("Satisfaction").value;

        if (savedX && savedY) {
            // Remove any existing highlights
            dialog.querySelectorAll("td[data-x][data-y]").forEach(cell => {
                cell.style.boxShadow = "none";
                cell.style.border = "1px solid #ccc";
            });

            // Find and highlight the selected cell
            const selectedCell = dialog.querySelector(`td[data-x="${savedX}"][data-y="${savedY}"]`);
            if (selectedCell) {
                selectedCell.style.boxShadow = "inset 2px 2px 5px rgba(0,0,0,0.3), inset -2px -2px 5px rgba(255,255,255,0.8)";
                selectedCell.style.border = "3px solid white";
            }
        }
    });

    let x = fd.field("Importance").value;
    let y = fd.field("Satisfaction").value;
    let bacjBKCol = fd.field("BKColor").value;

    if (bacjBKCol) {
        btn.textContent = `Selected: ${x}, ${y}`;
        btn.style.backgroundColor = bacjBKCol;
        btn.style.color = 'white';
        btn.style.fontWeight = 'bold';
    }

    // Optional: Close on outside click
    // window.addEventListener("click", function (e) {
    //     if (dialog.style.display === "block" && !dialog.contains(e.target) && e.target !== btn) {
    //         dialog.style.display = "none";
    //     }
    // });
}

var readonlyloadMatrix = async function (btn) {
    // Create overlay to detect outside clicks
    const overlay = document.createElement("div");
    overlay.style.position = "fixed";
    overlay.style.top = "0";
    overlay.style.left = "0";
    overlay.style.width = "100vw";
    overlay.style.height = "100vh";
    overlay.style.backgroundColor = "rgba(0, 0, 0, 0.3)";
    overlay.style.zIndex = 9998;
    overlay.style.display = "none";

    // Create the dialog container
    const dialog = document.createElement("div");
    dialog.id = "matrix-popup";
    dialog.style.position = "fixed";
    dialog.style.top = "50%";
    dialog.style.left = "50%";
    dialog.style.transform = "translate(-50%, -50%)";
    dialog.style.background = "#fff";
    dialog.style.border = "1px solid #ccc"; 
    dialog.style.paddingLeft = "30px";
    dialog.style.paddingBottom = "10px";
    dialog.style.paddingTop = "10px";
    dialog.style.paddingRight = "10px";
    dialog.style.zIndex = 9999;
    dialog.style.display = "none";
    dialog.style.boxShadow = "0 4px 8px rgba(0,0,0,0.2)";
    dialog.style.borderRadius = "8px";
    dialog.style.position = "fixed";

    // Add close (X) button
    const closeButton = document.createElement("span");
    closeButton.textContent = "×";
    closeButton.style.position = "absolute";
    closeButton.style.top = "5px";
    closeButton.style.right = "10px";
    closeButton.style.cursor = "pointer";
    closeButton.style.fontSize = "20px";
    closeButton.style.fontWeight = "bold";
    closeButton.style.color = "#888";
    closeButton.addEventListener("click", () => {
        dialog.style.display = "none";
        overlay.style.display = "none";        
    });
    dialog.appendChild(closeButton);

    // Table element
    const table = document.createElement("table");
    table.style.borderCollapse = "collapse";
    table.style.textAlign = "center";

    // Add X-axis labels row
    const labelRow = document.createElement("tr");
    for (let x = 1; x <= 10; x++) {
        const label = document.createElement("td");
        label.textContent = x;
        label.style.textAlign = "center";
        label.style.fontSize = "12px";
        label.style.fontWeight = "bold";
        label.style.padding = "2px";
        label.style.border = "none";
        labelRow.appendChild(label);
    }
    table.appendChild(labelRow);    

    for (let y = 10; y >= 1; y--) {
        const row = document.createElement("tr");

        for (let x = 1; x <= 10; x++) {
            const cell = document.createElement("td");
            cell.setAttribute("data-x", x);
            cell.setAttribute("data-y", y);
            cell.style.width = "30px";
            cell.style.height = "30px";
            cell.style.border = "1px solid #ccc";
            cell.style.cursor = "not-allowed";
            cell.style.textAlign = "center";
            cell.title = `Importance: ${x}, Satisfaction: ${y}`;
            cell.style.backgroundColor = getColor(x, y);         

            row.appendChild(cell);

            if (x === 10) {
                const yLabel = document.createElement("td");
                yLabel.textContent = y;
                yLabel.style.fontSize = "12px";
                yLabel.style.fontWeight = "bold";
                yLabel.style.paddingLeft = "6px";
                yLabel.style.border = "none";
                row.appendChild(yLabel);
            }
        }

        table.appendChild(row);        
    }

    const topRow = document.createElement('div');

    const leftText = document.createElement('div');
    leftText.textContent = 'Importance/Cost';
    leftText.style.transform = "rotate(-90deg)";
    leftText.style.fontWeight = 'bold';
    leftText.style.fontSize = '14px';
    leftText.style.position = 'absolute';
    leftText.style.left = '-32px';
    leftText.style.top = '200px';

    topRow.appendChild(leftText);
    topRow.appendChild(table);

    const bottomText = document.createElement('div');
    bottomText.textContent = 'Satisfaction/ Performance';
    bottomText.style.textAlign = 'center';
    bottomText.style.fontWeight = 'bold';
    bottomText.style.fontSize = '14px';
    bottomText.style.marginTop = '5px';

    dialog.appendChild(topRow);
    dialog.appendChild(bottomText);

    document.body.appendChild(overlay);
    document.body.appendChild(dialog);

    // Open matrix on button click
    btn.addEventListener("click", () => {
        dialog.style.display = "block";
        overlay.style.display = "block";       

        const savedX = fd.field("Importance").value;
        const savedY = fd.field("Satisfaction").value;

        if (savedX && savedY) {
            dialog.querySelectorAll("td[data-x][data-y]").forEach(cell => {
                cell.style.boxShadow = "none";
                cell.style.border = "1px solid #ccc";
            });

            const selectedCell = dialog.querySelector(`td[data-x="${savedX}"][data-y="${savedY}"]`);
            if (selectedCell) {
                selectedCell.style.boxShadow = "inset 2px 2px 5px rgba(0,0,0,0.3), inset -2px -2px 5px rgba(255,255,255,0.8)";
                selectedCell.style.border = "3px solid white";
            }
        }
    });   

    let x = fd.field("Importance").value;
    let y = fd.field("Satisfaction").value;
    let bacjBKCol = fd.field("BKColor").value;

    btn.textContent = `Selected: ${x}, ${y}`;
    btn.style.backgroundColor = bacjBKCol;
    btn.style.color = 'white';
    btn.style.fontWeight = 'bold';

    // Close when clicking outside the dialog
    window.addEventListener("click", function (e) {
        if (dialog.style.display === "block" && !dialog.contains(e.target) && e.target !== btn) {
            dialog.style.display = "none";   
            overlay.style.display = "none";
        }
    });
};

var GetVEPCompanies = async function(){   

    const listTitle = 'VEPCompanies';
    const camlFilter = `<View>                            
                            <Query>
                                <Where>                                    
                                    <IsNotNull><FieldRef Name='Title' /></IsNotNull>                        
                                </Where>
                            </Query>
                        </View>`;

    const existingItems = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });

    existingItems.forEach(item => {        
        if (item.SpGroup && item.Title) {            
            groupMap[item.SpGroup] = item.Title;
        }
    }); 
    
    return groupMap;
}

var HandleCascadedLookup = async function () {
    
    // (async () => {
        try {
            
            await _formFields.ProjectComponent.ready();
            _formFields.ProjectComponent.filter = `SubProjects/Id eq 0`;
            _formFields.ProjectComponent.refresh();
            
            const companiesList = await pnp.sp.web.lists
                .getByTitle("VEPCompanies")
                .items.select("Id", "Title")
                .filter(`Title eq '${_Originator}'`)
                .get();

            if (companiesList.length === 0) {
                console.warn(`Company '${_Originator}' not found in VEPCompanies list.`);
                return;
            }

            const compId = companiesList[0].Id;
            
            await _formFields.Category.ready();
            _formFields.Category.filter = `Companies/Id eq ${compId}`;
            _formFields.Category.refresh();
            
            _formFields.Category.$on('change', async function (value) {

                _formFields.ProjectComponent.value = [];

                if (!value || !value.LookupValue) {
                    console.warn("No value selected in Category.");
                    await _formFields.ProjectComponent.ready();
                    _formFields.ProjectComponent.filter = `SubProjects/Id eq 0`;
                    _formFields.ProjectComponent.refresh();
                    return;
                } 
                
                const subProjId = value.LookupId;

                await _formFields.ProjectComponent.ready();
                _formFields.ProjectComponent.filter = `SubProjects/Id eq ${subProjId}`;
                _formFields.ProjectComponent.refresh();
            });

        } catch (error) {
            console.error("An error occurred during dynamic filtering setup:", error);
        }
    // })();
}

var HandleCascadedUniformatLevels = async function () {
    try {
        
        await _formFields.UniformatLevel2.ready();
        _formFields.UniformatLevel2.filter = `UniformatLevel1/Id eq 0`;
        _formFields.UniformatLevel2.refresh();

        await _formFields.UniformatLevel3.ready();
        _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq 0`;
        _formFields.UniformatLevel3.refresh();
        
        _formFields.UniformatLevel1.$on('change', async function (level1Value) {
            _formFields.UniformatLevel2.value = [];
            _formFields.UniformatLevel3.value = [];

            if (!level1Value || !level1Value.LookupId) {
                console.warn("UniformatLevel1 not selected.");
                return;
            }

            const level1Id = level1Value.LookupId;

            await _formFields.UniformatLevel2.ready();
            _formFields.UniformatLevel2.filter = `UniformatLevel1/Id eq ${level1Id}`;
            _formFields.UniformatLevel2.refresh();
          
            await _formFields.UniformatLevel3.ready();
            _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq 0`;
            _formFields.UniformatLevel3.refresh();
        });
       
        _formFields.UniformatLevel2.$on('change', async function (level2Value) {
            _formFields.UniformatLevel3.value = [];

            if (!level2Value || !level2Value.LookupId) {
                console.warn("UniformatLevel2 not selected.");
                return;
            }

            const level2Id = level2Value.LookupId;

            await _formFields.UniformatLevel3.ready();
            _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq ${level2Id}`;
            _formFields.UniformatLevel3.refresh();
        });

    } catch (error) {
        console.error("Error during Uniformat cascade setup:", error);
    }
};

var HandleUnifiedCascade = async function () {

    try {

        if (_isNew) {

            _formFields.Category.value = [];
            _formFields.ProjectComponent.value = [];
            _formFields.UniformatLevel1.value = [];
            _formFields.UniformatLevel2.value = [];
            _formFields.UniformatLevel3.value = [];
        }      
       
        const companiesList = await pnp.sp.web.lists
            .getByTitle("VEPCompanies")
            .items.select("Id", "Title")
            .filter(`Title eq '${_Originator}'`)
            .get();

        if (companiesList.length === 0) {
            console.warn(`Company '${_Originator}' not found in VEPCompanies list.`);
            return;
        }

        const compId = companiesList[0].Id;

        await _formFields.Category.ready();
        _formFields.Category.filter = `Companies/Id eq ${compId}`;
        _formFields.Category.refresh();
    
        await _formFields.ProjectComponent.ready();
        const SubProId = _formFields.Category?.value?.LookupId ?? 0;
        _formFields.ProjectComponent.filter = `SubProjects/Id eq ${SubProId}`;
        _formFields.ProjectComponent.refresh();
      
        await _formFields.UniformatLevel1.ready();
        const ScopeId = _formFields.ProjectComponent?.value?.LookupId ?? 0;
        _formFields.UniformatLevel1.filter = `Scope/Id eq ${ScopeId}`;
        _formFields.UniformatLevel1.refresh();

        await _formFields.UniformatLevel2.ready();
        const UniformatLevel1Id = _formFields.UniformatLevel1?.value?.LookupId ?? 0;
        _formFields.UniformatLevel2.filter = `UniformatLevel1/Id eq ${UniformatLevel1Id}`;
        _formFields.UniformatLevel2.refresh();

        await _formFields.UniformatLevel3.ready();
        const UniformatLevel2Id = _formFields.UniformatLevel2?.value?.LookupId ?? 0;
        _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq ${UniformatLevel2Id}`;
        _formFields.UniformatLevel3.refresh();
       
        _formFields.Category.$on('change', async function (categoryValue) { 
            
            _formFields.ProjectComponent.value = [];
            _formFields.UniformatLevel1.value = [];
            _formFields.UniformatLevel2.value = [];
            _formFields.UniformatLevel3.value = [];

            if (!categoryValue || !categoryValue.LookupId) {
                console.warn("No Category selected.");
                _formFields.ProjectComponent.filter = `SubProjects/Id eq 0`;
                _formFields.UniformatLevel1.filter = `Scope/Id eq 0`;
            } else {
                const catId = categoryValue.LookupId;
                _formFields.ProjectComponent.filter = `SubProjects/Id eq ${catId} and Companies/Id eq ${compId}`;
                _formFields.UniformatLevel1.filter = `Scope/Id eq ${catId}`;
            }

            _formFields.ProjectComponent.refresh();
            _formFields.UniformatLevel1.refresh();

            _formFields.UniformatLevel2.filter = `UniformatLevel1/Id eq 0`;
            _formFields.UniformatLevel2.refresh();

            _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq 0`;
            _formFields.UniformatLevel3.refresh();
        });

        _formFields.ProjectComponent.$on('change', async function (projectComponentValue) {           
            _formFields.UniformatLevel1.value = [];
            _formFields.UniformatLevel2.value = [];
            _formFields.UniformatLevel3.value = [];

            if (!projectComponentValue || !projectComponentValue.LookupId) {
                console.warn("No ProjectComponent selected.");
                _formFields.UniformatLevel1.filter = `Scope/Id eq 0`;
            } else {
                const pcId = projectComponentValue.LookupId;
                _formFields.UniformatLevel1.filter = `Scope/Id eq ${pcId}`;
            }

            _formFields.UniformatLevel1.refresh();

            _formFields.UniformatLevel2.filter = `UniformatLevel1/Id eq 0`;
            _formFields.UniformatLevel2.refresh();

            _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq 0`;
            _formFields.UniformatLevel3.refresh();
        });
        
        _formFields.UniformatLevel1.$on('change', async function (level1Value) {       
            _formFields.UniformatLevel2.value = [];
            _formFields.UniformatLevel3.value = [];

            if (!level1Value || !level1Value.LookupId) {
                console.warn("No UniformatLevel1 selected.");
                _formFields.UniformatLevel2.filter = `UniformatLevel1/Id eq 0`;
            } else {
                const level1Id = level1Value.LookupId;
                _formFields.UniformatLevel2.filter = `UniformatLevel1/Id eq ${level1Id}`;
            }

            _formFields.UniformatLevel2.refresh();

            _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq 0`;
            _formFields.UniformatLevel3.refresh();
        });
       
        _formFields.UniformatLevel2.$on('change', async function (level2Value) {         
            _formFields.UniformatLevel3.value = [];

            if (!level2Value || !level2Value.LookupId) {
                console.warn("No UniformatLevel2 selected.");
                _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq 0`;
            } else {
                const level2Id = level2Value.LookupId;
                _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq ${level2Id}`;
            }

            _formFields.UniformatLevel3.refresh();
        });

    } catch (error) {
        console.error("Error during unified cascaded filtering:", error);
    }
};

function updateSavings(original, proposed) {

    let originalVal = parseFloat(original);
    let proposedVal = parseFloat(proposed);

    if (!isNaN(originalVal) && !isNaN(proposedVal)) {
        let savings = originalVal - proposedVal;
        _formFields.Savings.value = savings.toFixed(2); // optional: format to 2 decimals
    } else {
        _formFields.Savings.value = ''; // Clear if values are invalid
    }
}

async function toggleCoordinationCompleted(showCoordinationCompleted) {
  
    const coordinationField = _formFields?.CoordinationCompleted;   
    if (!coordinationField) return;
    
    _HideFormFields([coordinationField], !showCoordinationCompleted);
  
    const buttons = fd.toolbar.buttons;
    const submitBtn = buttons.find(b => b.text === 'Submit');
    const notReadyBtn = buttons.find(b => b.text === _notReadyButtonName);

    if (showCoordinationCompleted) {        
        await updateButtonVisibility(coordinationField.value, submitBtn, notReadyBtn);        
        coordinationField.$on('change', async (newValue) => {
            await updateButtonVisibility(newValue, submitBtn, notReadyBtn);
        });
    } else {       
        if (submitBtn) submitBtn.visible = true;
        if (notReadyBtn) notReadyBtn.visible = false;   
        coordinationField.value = false;         
    }
}

async function updateButtonVisibility(isCompleted, submitBtn, notReadyBtn) {

    let CoordinationRequired = _formFields.CoordinationRequired.value;
    
    if (CoordinationRequired) {
        if (submitBtn) submitBtn.visible = !!isCompleted;
        if (notReadyBtn) notReadyBtn.visible = !isCompleted;
    }    
}

function disableVEPSection() {
    let divGridVPMEl = document.querySelectorAll('.divGrid.VPM');
    if (divGridVPMEl) {
        _setFieldsDisabled([_formFields.VEPReference, _formFields.CoordinationRequired, _formFields.CoordinationCompleted,
            _formFields.OriginalCost, _formFields.ProposedCost, _formFields.Savings,
            _formFields.TimeImpact, _formFields.ReviewedDate, _formFields.SelectCompanyForReview], true);
        disableRichTextFieldColumn(_formFields.OriginalDesign);
        disableRichTextFieldColumn(_formFields.ProposedDesign);
        disableRichTextFieldColumn(_formFields.PerformanceBenefits);
        disableRichTextFieldColumn(_formFields.Comments);
    }
}

function formatNumberWithCommas(value) {
    if (!value || isNaN(value)) return '';
    return parseFloat(value).toLocaleString(); // adds commas
}

function updateFormatted(formatDiv, value) {
    const formatted = formatNumberWithCommas(value);
    formatDiv.textContent = formatted ? `Formatted: ${formatted}` : '';
}

var HandleLevel3CascadedLookup = async function () {
    
    // (async () => {
        try {           
            
            await _formFields.UniformatLevel2.ready();
            let UniformatLevel2value = _formFields.UniformatLevel2.value.LookupId;         
            
            await _formFields.UniformatLevel3.ready();
            _formFields.UniformatLevel3.filter = `UniformatLevel2/Id eq ${UniformatLevel2value}`;
            _formFields.UniformatLevel3.refresh();           

        } catch (error) {
            console.error("An error occurred during dynamic filtering setup:", error);
        }
    // })();
}
//#endregion

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

    if($('.text-muted').length > 0)
      $('.text-muted').remove();

    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;    
    groupMap = await GetVEPCompanies();
    ProjectNumber = await getParameter('ProjectName');

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
}

var getAutReferenceFormat = async function (isVep) {
    
    let result;
    const listTitle = _MajorTypes; //list.Title;

    const camlFilter = `<View>
                            <Query>
                                <Where>
                                    <Eq><FieldRef Name='Title'/><Value Type='Text'>VEPM</Value></Eq>
                                </Where>
                            </Query>
                        </View>`;
    let items = await pnp.sp.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ ViewXml: camlFilter });   
    if (items?.length === 1) {
        const item = items[0];
        result = isVep ? item?.CRSFormat ?? '' : item?.CDSFormat ?? '';
    } 
    return result;
}

var parseRefFormat = async function (format, updateCounter) {
    const splitRefFormat = format.split('-');
    let returnValue = "";

    for (let i = 0; i < splitRefFormat.length; i++) {
        let column = splitRefFormat[i].replace('[', '').replace(']', '');

        if (column.includes('"')) {
            column = column.replace(/"/g, '');
        } else if (column.includes('$')) {
            column = column.replace(/\$/g, '');
        } else {
            const itemValue = _formFields[column].value;
            if (itemValue !== undefined && itemValue !== "") {
                column = itemValue.toString();
                if (column.includes('#')) {
                    const splitVal = column.split('#');
                    column = splitVal[1];
                }
            }
        }

        returnValue += column + "-";
    }    

    returnValue = returnValue.slice(0, -1); // Remove the trailing dash 
    const digit = returnValue.substring(returnValue.lastIndexOf('-') + 1);
    const lastSlash = returnValue.lastIndexOf('-'); 
    returnValue = (lastSlash > -1) ? returnValue.substring(0, lastSlash) : returnValue;    
    const counter = await GetReferenceCounter(returnValue, updateCounter); // Assumes this returns a number    
    returnValue = returnValue + '-' + counter.toString().padStart(digit.length, '0');

    return returnValue;
}

var GetReferenceCounter = async function (returnValue, updateCounter) {    
   
    let listname = 'Counter';

    try {

        const camlFilter = `<View>
                                <Query>
                                    <Where>
                                        <Eq><FieldRef Name='Title'/><Value Type='Text'>${returnValue}</Value></Eq>
                                    </Where>
                                </Query>
                            </View>`;
        let items = await pnp.sp.web.lists.getByTitle(listname).getItemsByCAMLQuery({ ViewXml: camlFilter }); 
        
        if (updateCounter) {

            var _cols = {};
            
            if (items.length == 0) {
                
                _cols["Title"] = returnValue;
                let value = '2';
                _cols["Counter"] = value;                 
                await pnp.sp.web.lists.getByTitle(listname).items.add(_cols); 
                return '1';
            }
            else if (items.length == 1) {   
                
                var _item = items[0];  
                let value = parseInt(_item.Counter) + 1;
                _cols["Counter"] = `${value}`;
                await pnp.sp.web.lists.getByTitle(listname).items.getById(_item.Id).update(_cols);
                return (value - 1).toString(); //value.toString();
            }
        }
        else {            
            return items.length === 0 ? 1 : items[0].Counter;
        }

    } catch (error) {
        console.error('Error fetching/updating reference counter:', error);
        throw new Error('Error fetching/updating reference counter');
    }
}

async function CheckifUserinSPGroup() {
    var IsTMUser = "null";
    try{
         await pnp.sp.web.currentUser.get()
         .then(async function(user){
            await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
             .then(async function(groupsData){
                for (var i = 0; i < groupsData.length; i++) {               
                    const title = groupsData[i].Title;
                    if (groupMap.hasOwnProperty(title)) {
                        IsTMUser = groupMap[title];
                        break;
                    }                                           
                }               
            });
         });
    }
    catch(e){alert(e);}
    return IsTMUser;                
}

function generateEmailBody(bodyMessage) {
    let BodyEmail = `<html>
                        <head>
                            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
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

    return BodyEmail;
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

async function getCurrentUserGroups() {
    try {
        const user = await pnp.sp.web.currentUser.get();
        const groups = await pnp.sp.web.siteUsers.getById(user.Id).groups.get();
        console.log("User is in the following groups:", groups);
    } catch (error) {
        console.error("Error getting user groups:", error);
    }
}

function hideHideSection(isHide) {
  document.querySelectorAll('.hideSection').forEach(el => {
    el.style.display = isHide ? 'none' : '';
  });
}

const _setFieldsDisabled = (fields, isDisabled) => {
  (Array.isArray(fields) ? fields : [fields]).forEach(field => {
    if (field) field.disabled = isDisabled;
  });
};

async function CheckifUserinGroup() {
	var IsTMUser = "Trade";
	try{
		 await pnp.sp.web.currentUser.get()
         .then(async function(user){
			await pnp.sp.web.siteUsers.getById(user.Id).groups.get()
			 .then(async function(groupsData){
				for (var i = 0; i < groupsData.length; i++) {				
                    const title = groupsData[i].Title;                    
                    if (groupMap.hasOwnProperty(title)) {                        
                        IsTMUser = title;                        
                        break;
                    } 
				}				
			});
	     });
    }
	catch(e){alert(e);}

	var isSiteAdmin = _spPageContextInfo.isSiteAdmin; 
	if(isSiteAdmin)
		IsTMUser = "Admin";
		
	return IsTMUser;				
}

async function GetCompanyDetails(key) {
	let result = "";
    await pnp.sp.web.lists
        .getByTitle("VEPCompanies")
        .items
        .select("Title,CompanyType,SpGroup")
        .filter(`Title eq '${  key  }'`)
        .get()
        .then((items) => {
            if(items.length > 0)
                result = {
                    CompanyType: items[0].CompanyType,
                    SpGroup: items[0].SpGroup
                }; 
            });
    return result;			
}

function SetAttachmentToReadOnly(){
	fd.field('Attachments').disabled = false;

	var spanATTDelElement = document.querySelector('.k-upload .k-upload-files .k-upload-status');
	if(spanATTDelElement !== null)
	{
		spanATTDelElement.style.display = 'none';
		
		var spanATTUpElement = document.querySelector('.k-upload .k-upload-button');
		spanATTUpElement.style.display = 'none';
		
		var spanATTZoneElement = document.querySelector('.k-dropzone');
		if(spanATTZoneElement !== null)
			spanATTZoneElement.style.display = 'none';
	}
	else
		DisableAttachment();
}

function DisableAttachment() {
	fd.field('Attachments').disabled = false;
	
	$('div.k-upload-button').remove();
	$('button.k-upload-action').remove();
	$('.k-dropzone').remove();
}

const _isRequiredFields = (fields, isDisabled) => {
  (Array.isArray(fields) ? fields : [fields]).forEach(field => {
    if (field) field.required = isDisabled;
  });
};
//#endregion

//#region Custom Buttons
var setCustomButtons = async function () {

    if (!_isDisplay) {
        fd.toolbar.buttons[0].style = "display: none;";
        await setButtonActions("Accept", _notReadyButtonName, `${saveColor}`);
        await setButtonActions("Accept", submitDefault, `${greenColor}`);
    }
    fd.toolbar.buttons[1].style = "display: none;"; 
   
    await setButtonActions("ChromeClose", "Cancel", `${yellowColor}`);

    setToolTipMessages();
}

const setButtonActions = async function(icon, text, bgColor){

    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          style: `background-color: ${bgColor}; color: white;`, //color of font button//
          click: async function() {

            if(text == "Close" || text == "Cancel"){
                showPreloader();
                fd.close();
            }
            else if (text == _notReadyButtonName) {

                if (fd.isValid) {
                    showPreloader();
                    _nextStatus = _notReadyButtonName;                        
                    fd.save();
                }
            }
            else if (text == 'Submit') {

                if (fd.isValid) {

                    showPreloader(); 
                    
                    _titleforEmail = _formFields.Title.value;
                    
                    const entries = Object.entries(_emailFields);
                    _emailBody = entries.map(([key, val], index) => {                    
                        const isLast = index === entries.length - 1;
                        let displayValue = (val?.value.LookupValue ?? val?.value) || ''; // Default to an empty string if value is undefined
                        if (key === 'WorkflowStatus') 
                            displayValue = _nextStatus; 
                        if (displayValue) {
                            return `
                                <tr>
                                    <td style="padding: 10px 12px; font-weight: 600; background-color: #f9fafb; color: #1a202c; width: 250px;">
                                        ${key}:
                                    </td>
                                    <td style="padding: 10px 12px; background-color: #ffffff; color: #4a5568; border-bottom: 1px solid #F3F4F6;">
                                        ${displayValue}
                                    </td>
                                </tr>
                            `;
                        }
                    }).join('');

                    if (_nextStatus === 'IDEA Submitted') {
                        _Reference = await parseRefFormat(_schema, true);
                        _formFields.Reference.value = _Reference;
                        _formFields.WorkflowStatus.value = _nextStatus;                       
                        fd.save();                    
                    } 
                    else if (_nextStatus === 'Details Submitted') {
                        
                        if (_status !== 'Revise After Rejection') {
                            _VEPReference = await parseRefFormat(_VEPSchema, true);
                            _formFields.VEPReference.value = _VEPReference;
                        }
                        _formFields.WorkflowStatus.value = _nextStatus;
                        fd.field('OMDate').value = new Date();
                        fd.save();                    
                    }                        
                    else {

                        if (_nextStatus === 'Rejected' || _nextStatus === 'Request for Details')
                            fd.field('CMDate').value = new Date();                     
                        
                        if (_nextStatus === 'Request for Details' && _formFields.SelectCompanyForReview.value) {
                            _formFields.Originator.value = _formFields.SelectCompanyForReview.value.LookupValue;
                            _SelectCompanyForReview = _formFields.SelectCompanyForReview.value.LookupValue;
                        }
                            
                        _formFields.WorkflowStatus.value = _nextStatus;   
                        
                        fd.save();                    
                    }
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
    $('.fd-grid.container-fluid').attr("style", "margin-top: 10px !important; padding: 10px;");
    const marginTopValue = _webUrl.includes("db-sp") ? "-22px" : "-10px";
    $('.fd-form-container.container-fluid').css("margin-top", marginTopValue);        

    const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';
    const linkElement = `<a href="${_spPageContextInfo.webAbsoluteUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">
                          <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">${titelValue}</a>`;
    $('span.o365cs-nav-brandingText').html(linkElement);

    $('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');

    // setTimeout(function () {
    //     $('input[id^="fd-field-"]:disabled').css({
    //         'background-color': 'transparent',
    //         'opacity': '1',
    //         'border': 'none',
    //         'pointer-events': 'none', // Optional: disables interaction
    //         'color': '#495057'        // Optional: normal text color
    //     });
    // }, 300); // delay in milliseconds
    
    // $('.border-title').each(function() {
    //     $(this).css({          
    //         'margin-top': '-35px', /* Adjust the position to sit on the border */
    //         'margin-left': '20px', /* Align with the content */            
    //     });
    // });
}

function setIconSource(elementId, iconFileName) {

  const iconElement = document.getElementById(elementId);

  if (iconElement) {
      iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
  }
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

function toggleToolbarButton(label, shouldDisable) {
    let buttons = fd.toolbar.buttons;
    let targetButton = buttons.find(button => button.text === label);
    
    if (targetButton) {
        disableEnableButton(targetButton, shouldDisable);
    } else {
        console.warn(`Button with label "${label}" not found.`);
    }
}

async function updateSubmitButtonState(transition) { 

    let showReasonField = false;
    _code = _formFields.Code.value;
    if(_module === 'VEPMTasks')
        _reviewFlow = _code;
    toggleToolbarButton('Submit', !_code); // disables if code is falsy
    let config = transition.nextStatus[_code];
    if (config) {
        _nextStatus = config.status;
        showReasonField = config.showReasonField;
        _HideFormFields([_formFields.ReasonofRejection], !showReasonField);

        if (!showReasonField) {
            _formFields.ReasonofRejection.value = ''; // or '' depending on your form setup
            _formFields.ReasonofRejection.required = false; 
            
            if (_companyType === 'Client' || _companyType === 'PMC') {
                _HideFormFields([_formFields.SelectCompanyForReview], false);
                _isRequiredFields([_formFields.SelectCompanyForReview], true);

                await _formFields.SelectCompanyForReview.ready();
                _formFields.SelectCompanyForReview.filter = `CompanyType eq 'Consultant' or CompanyType eq 'Contractor'`;
                _formFields.SelectCompanyForReview.refresh();                               
            }
        }
        else {
            _formFields.ReasonofRejection.required = true;        

            if (_formFields.SelectCompanyForReview) {
                _formFields.SelectCompanyForReview.value = null; // or '' depending on your form setup
                _formFields.SelectCompanyForReview.required = false;
                _HideFormFields([_formFields.SelectCompanyForReview], true);
            }
        }
    }
    else {
        _formFields.ReasonofRejection.value = '';
        _formFields.ReasonofRejection.required = false;
        _HideFormFields([_formFields.ReasonofRejection], !showReasonField);         
    }
}

const getColor = (x, y) => {
    // Assign colors based on zones
    if (x <= 5 && y <= 5) return "#f87171";    //  Red
    if (x >= 6 && y >= 6) return "#4ade80";    //  Green
    if (x <= 5 && y >= 5) return "#fde047";    // yellow
    if (x >= 5 && y <= 5) return "#fde047";    // yellow
    return "#f0f0f0";                          // Neutral
};
//#endregion

fd.spSaved(async function(result) {
    try { 
        
        if (_nextStatus !== _notReadyButtonName) {

            _itemId = result.Id;
            let listPath = `${fd.listUrl}`;
            let listInternalName = listPath.replace("/Lists/", "").trim();
            let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${_itemId}</Value></Eq></Where>`;
        
            let formType = 'EditForm';

            let linkURL = _spPageContextInfo.webAbsoluteUrl + `/SitePages/PlumsailForms/${listInternalName}/Item/${formType}.aspx?item=` + _itemId;
        
            let Subject = '';
            let reviewFlowMessage = '';
            let DearTo = 'Dear All,';

            if (_reviewFlow) {

                linkURL = _spPageContextInfo.webAbsoluteUrl + `/SitePages/PlumsailForms/VEPM/Item/${formType}.aspx?item=` + _parentID;

                if (_reviewFlow === 'Objection') {
                    Subject = `IDEA Objected by ${_Originator} - for ${_ideaOrigin} (Ref: ${_Reference}) ${_titleforEmail}`;
                    reviewFlowMessage = 'We regret to inform you that the review process has <strong><span style="color:red;">objected</span></strong> the submitted <strong>IDEA</strong> details. Please review the remarks below.';
                } else if (_reviewFlow === 'No Objection') {
                    Subject = `All Tasks Completed - ${_ideaOrigin} (Ref: ${_Reference})`;
                    reviewFlowMessage = 'We are pleased to inform you that the review process has <strong><span style="color:green;">no objection</span></strong> all related tasks. See the details below.';
                }
            } else if (_nextStatus === 'Rejected') {
                Subject = `(Company: ${_Originator}) IDEA Rejected by PM / Client - (Title: ${_titleforEmail}) - (Ref: ${_Reference})`;
                DearTo = `Dear ${_Originator},`;
            } else if (_nextStatus === 'Request for Details') {
                Subject = `(Company: ${_Originator}) IDEA Approved by PM / Client - (Title: ${_titleforEmail}) - (Ref: ${_Reference})`;
                DearTo = `Dear ${_Originator},`;
            } else if (_nextStatus === 'Details Submitted') {
                Subject = `(Company: ${_Originator}) Details Submitted - (Title: ${_titleforEmail}) - (Ref: ${_Reference}) - (VEP Ref: ${_VEPReference})`;
            } else {
                Subject = `(Company: ${_Originator}) New IDEA Submission - (Title: ${_titleforEmail}) - (Ref: ${_Reference})`;
                DearTo = 'Dear PM,';
            }

            let bodyMessage = `
            <p>
            ${DearTo}
            </p>
            <br />
            <p>
                ${reviewFlowMessage ||
                (_nextStatus === 'Rejected'
                    ? 'Your submitted <strong>IDEA</strong> has been <strong><span style="color:red;">rejected</span></strong>. Please find the detailed below.'
                    : _nextStatus === 'Request for Details'
                        ? 'Your submitted <strong>IDEA</strong> has been <strong><span style="color:green;">approved</span></strong>. Please find the details below.'
                        : _nextStatus === 'Details Submitted'
                            ? 'Please be informed that in reference to the <strong>approved IDEA</strong>, the requested details have been submitted. Kindly review the information below and provide your feedback accordingly.'
                            : 'Kindly be informed that a new <strong>IDEA</strong> has been submitted. Please find the detailed information below.')
                }
            </p>
            <br />
            ${_nextStatus === 'Details Submitted'
                    ? `
                        <p style="padding: 12px; font-weight: bold; color: #005a9e;font-family: Segoe UI, sans-serif; font-size: 14px;text-decoration: underline;">Tasks for Each Company:</p>
                        <div id="detailsSubmittedPlaceholder"></div>
                        <p style="padding: 12px; font-weight: bold; color: #005a9e;font-family: Segoe UI, sans-serif; font-size: 14px;text-decoration: underline;">Details for the Approved IDEA:</p>
                    `
                    : `
                        ${_reviewFlow
                        ? `<p>Click <a href="${linkURL}" class="link">here</a> to ${_reviewFlow === 'Objection'
                            ? 'view the review decision and comments'
                            : _reviewFlow === 'No Objection'
                                ? 'view the final approval and review outcome'
                                : 'view the review status'
                        }.</p><div id="detailsSubmittedPlaceholder"></div><br />`
                        : _nextStatus === 'Request for Details' && _SelectCompanyForReview
                            ? `
                                Dear ${_SelectCompanyForReview},<br /><br />
                                <p>
                                Click <a href="${linkURL}" class="link">here</a> to submit VEP details.
                                </p><br />
                            `
                            : `
                                <p>
                                Click <a href="${linkURL}" class="link">here</a> to ${
                                    _nextStatus === 'Rejected'
                                    ? 'view the details'
                                    : _nextStatus === 'Request for Details'
                                        ? 'submit VEP details'
                                        : 'take the necessary action'
                                }.
                                </p><br />
                            `                            
                        }`
                }           
            <table style="width: 65%; border-collapse: collapse; font-family: Segoe UI, sans-serif; font-size: 13px; border: 1px solid #e2e8f0; border-radius: 6px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                ${_emailBody}
            </table>`;
        
            Subject = `${ProjectNumber} - ${Subject}`
        
            let BodyEmail = generateEmailBody(bodyMessage);
            let encodedSubject = htmlEncode(Subject);
            let encodedBodyEmail = htmlEncode(BodyEmail);
        
            if (_reviewFlow)
                await _sendEmail(_module, encodedSubject + '|' + encodedBodyEmail, query, '', `${_module}_${_reviewFlow}`, '', _currentUser);
            else
                await _sendEmail(_module, encodedSubject + '|' + encodedBodyEmail, query, '', `${_module}_${_nextStatus}`, '', _currentUser);
        }
               
    } catch (e) {
      console.log(e);
  }
});