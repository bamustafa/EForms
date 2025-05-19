var _web, _webUrl, _siteUrl, _layout, _module = '', _formType = '', _list, _itemId;

var _isSiteAdmin = false, _isNew = false, _isEdit = false, _isDisplay = false, _isPM = false, _isChildForm = false, _isReviewForm = false;

let _isEmployeeComment = false, _isManagerComment = false, _isFinalRatingSet = false, _isEmpSignOff = false, _isMgrSignOff = false;

var _manager = '', _employeeId = '', _managerId = '', _employeePartyNumber = '', _managerPartyNumber = '', _currentUser, _formFields = {}, _PerfomanceReviewFormFields = {};
let _EmployeeName = '', _ManagerName = '', _UserIntheForm = '', _WorkflowStatus = '', _nextStatus = '';
let _GoalID = '';
let _GoalsListName = 'Performance Management - Goals';

let _isAllEmployeeCommentDate = false, _isAllManagerCommentDate = false;

const greenColor = '#5FC9B3', redColor = '#F28B82', yellowColor = '#6c757d';

const disableField = (field) => field.disabled = true;
const enabledField = (field) => field.disabled = false;

const itemsToRemove = ['WorkflowStatus'];

const ratings = {
    "Exceptional Achievement": 5,
    "Above Target": 4,
    "On Target": 3,
    "Progressing": 2,
    "Needs Significant Improvement": 1
};

var onRender = async function (relativeLayoutPath, moduleName, formType) {
    
    try {      

        _isChildForm = moduleName === 'CPG' ? true : false;
        _isReviewForm = moduleName === 'PReview' ? true : false;
        _layout = relativeLayoutPath; 

        _formFields = {
            Title: fd.field('Title'),
            StartDate: fd.field('StartDate'),
            EndDate: fd.field('EndDate'),
            ApprovalStatus: fd.field('ApprovalStatus'),
            ReviewedDate: fd.field('ReviewedDate'),

            Level: fd.field('Level'),
            GoalCategory: fd.field('GoalCategory'),
            Overview: fd.field('Overview'),
            DateFinished: fd.field('DateFinished'),

            PercentComplete: fd.field('PercentComplete'),
            Status: fd.field('Status'),
            Attachments: fd.field('Attachments'),

            Submit: fd.field('Submit'),
            GoalID: fd.field('GoalID'),
            DiscussionID: fd.field('ReviewID_x003a_Discussion'),
            Templates: fd.field('Templates'),
            Manager: fd.field('Manager'),         
            EmployeeId: fd.field('EmployeeId'),
            EmployeePartyNumber: fd.field('EmployeePartyNumber'),
            ManagerId: fd.field('ManagerId'),
            ManagerPartyNumber: fd.field('ManagerPartyNumber'),
            goalsDT: fd.control('SPDataTable1'),
            Rating: fd.field('Rating'),
            ManagerComment: fd.field('ManagerComment'),
            EmployeeComment: fd.field('EmployeeComment'),
            EmployeeCommentDate: fd.field('EmployeeCommentDate'),
            ManagerCommentDate: fd.field('ManagerCommentDate'),
        }    

        _PerfomanceReviewFormFields = {
            Description: fd.field('Title'),
            Discussion: fd.field('Discussion'),
            PerformancePeriodId: fd.field('PerformancePeriodId'),
            DiscussionApprovalStatus: fd.field('DiscussionApprovalStatus'),
            // FinishedDate: fd.field('FinishedDate'),
            TotalRatingScore: fd.field('TotalRatingScore'),
            AverageRatingScore: fd.field('AverageRatingScore'),
            StartDate: fd.field('StartDate'),
            EndDate: fd.field('EndDate'),
            Employee: fd.field('Employee'),
            Manager: fd.field('Manager'),
            WorkflowStatus: fd.field('WorkflowStatus'),
            FinalRate: fd.field('FinalRate'),
            EmployeeSignOff: fd.field('EmployeeSignOff'),
            EmployeeSignOffDate: fd.field('EmployeeSignOffDate'),
            ManagerSignOff: fd.field('ManagerSignOff'),
            ManagerSignOffDate: fd.field('ManagerSignOffDate'),
            PersonnelNumber: fd.field('PersonnelNumber'),
            goalsDT: fd.control('SPDataTable1')
        };      

        clearLocalStorageItemsByField(itemsToRemove);

        await loadScripts().then(async () => {
            
            showPreloader();         

            await extractValues(moduleName, formType);
            await setCustomButtons();            

            if (_isEdit) {
                //await handleEditForm();               
                if (_isChildForm) {   
                    
                    initializeEvents();
                    validateGoals();
                    //fixTableColumnsWidth();
                    //setTableStyle(); 

                    let dt = _formFields.goalsDT;
                    
                    dt.ready().then(async function () {
                        dt.buttons[dt.buttons.length - 1].visible = false;
                        setTimeout(function() {
                            fixTableColumnsWidth(); 
                            setTableStyle();
                        }, 250);            
                    });
                    
                    dt.$on('change', async function (changeData) {                  
                        setTimeout(function () {                            
                            fixTableColumnsWidth(); 
                        }, 250);                       
                    });

                    let goals = _formFields.goalsDT.widget.dataItems();
                    let allReviewed = goals.every(goal => goal.ApprovalStatus === "Approved" || goal.ApprovalStatus === "Rejected");

                    if (allReviewed) {

                        let buttons = fd.toolbar.buttons;
                        let ApprovedButton = buttons.find(button => button.text === 'Approve');
                        if (ApprovedButton)
                            disableEnableButton(ApprovedButton, true);
                        let RejectedButton = buttons.find(button => button.text === 'Reject');
                        if (RejectedButton)
                            disableEnableButton(RejectedButton, true, '#F5B7B1');
                        setTimeout(function() {
                            disableDataTableCheckBoxes(); 
                        }, 200);
                    }
                }
                else if (_isReviewForm) 
                    await handlePerfomanceReviewForm();                   
                
                else _HideFormFields([_formFields.Submit, _formFields.GoalID, _formFields.DiscussionID], true);
            }
            else if (_isNew) {
                if (_isChildForm) {                  
                    // applyDataTable();
                    initializeEvents();
                    validateGoals();
                    fixTableColumnsWidth();              
                }
                else _HideFormFields([_formFields.Submit, _formFields.Manager, _formFields.EmployeeId, _formFields.EmployeePartyNumber, _formFields.ManagerId, _formFields.ManagerPartyNumber], true);
            }
            if (_isDisplay) {               
               
                await handleDisplayForm();

                if (_isReviewForm)
                    await handlePerfomanceReviewForm(); 
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

var Goal_newForm = async function(){

    try {

        const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
        const svgGoals = `<svg height="24" width="24" version="1.1" id="_x32_" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" 
        viewBox="0 0 512 512"  xml:space="preserve">
        <style type="text/css">
            .st0{fill:#000000;}
        </style>
        <g>
            <path class="st0" d="M228.31,77.506c21.358-1.961,37.07-20.901,35.093-42.259c-1.994-21.342-20.917-37.063-42.268-35.077
                c-21.358,1.986-37.054,20.893-35.076,42.259C188.045,63.778,206.968,79.498,228.31,77.506z"/>
            <path class="st0" d="M368.479,388.205c7.133,6.194,17.918,5.605,24.341-1.332l0.458-0.482c6.406-6.928,6.185-17.682-0.499-24.341
                l-45.045-50.028l-16.48-55.888c-0.752,1.34-1.536,2.663-2.402,3.922c-8.481,12.191-21.841,19.994-36.646,21.359l-6.822,0.474
                l17.387,41.997c4.118,6.848,9.224,13.032,15.14,18.409L368.479,388.205z"/>
            <path class="st0" d="M214.419,351.568c9.168-2.255,14.977-11.275,13.253-20.541l-11.128-59.76l73.455-5.099
                c10.532-0.744,20.19-6.21,26.228-14.904c6.03-8.677,7.794-19.609,4.804-29.766l-2.754-9.355l-22.968-83.668l39.374,2.132
                l31.474,30.322c-1.52,1.846-2.28,4.281-1.668,6.798l2.566,10.401l-18.294,4.51c-4.339,1.07-6.978,5.45-5.908,9.78l12.322,50.005
                c1.078,4.322,5.449,6.97,9.772,5.915l77.671-19.152c4.339-1.078,6.962-5.458,5.9-9.789l-12.338-49.988
                c-1.063-4.33-5.434-6.986-9.764-5.915l-18.302,4.51l-2.574-10.41c-1.046-4.232-5.311-6.798-9.543-5.76l-1.92,0.474
                c-0.434-1.773-1.218-3.489-2.395-5.025l-30.779-40.363c-3.325-4.347-8.048-7.419-13.384-8.662l-50.626-18.122
                c-18.131-6.504-38.328-3.333-53.6,8.4l-70.244,53.894l-45.56-16.431c-7.142-3.178-15.484-0.123-18.891,6.879l-0.474,0.956
                c-1.683,3.489-1.928,7.476-0.662,11.12c1.275,3.652,3.955,6.642,7.436,8.309l54.916,26.213c6.83,3.268,14.781,3.235,21.595-0.082
                l29.268-19.463l20.091,57.587l-48.591,4.249c-8.187,0.744-15.622,5.099-20.247,11.881c-4.641,6.79-5.982,15.296-3.693,23.189
                l24.152,82.606c2.68,9.143,12.084,14.568,21.342,12.289L214.419,351.568z M397.819,159.212l0.351,0.204l2.574,10.402l-26.318,6.488
                l-2.483-10.083c4.885,4.004,11.856,4.306,16.929,0.392l0.295-0.238c2.124-1.634,3.611-3.783,4.461-6.136L397.819,159.212z"/>
            <polygon class="st0" points="239.266,387.895 212.245,371.586 212.245,512 299.754,512 299.754,364.07 289.467,358.611 	"/>
            <polygon class="st0" points="316.488,512 403.988,512 403.988,419.377 316.488,372.951 	"/>
            <polygon class="st0" points="420.722,428.258 420.722,512 494.455,512 494.455,467.38 	"/>
            <polygon class="st0" points="108.011,512 195.512,512 195.512,361.479 108.011,308.647 	"/>
            <polygon class="st0" points="17.545,512 91.278,512 91.278,298.54 17.545,254.026 	"/>
        </g>
        </svg>`;
        const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;

        setIconSource("overview-icon", svguserinfo);
        setIconSource("goals-icon", svgGoals);
        setIconSource("attachment-icon", svgattachment);
        formatingButtonsBar('Human Resources: Performance Management Goals');        

        if (!_isChildForm) {
            
            renderAIDAHelper(25);

            //clearFormFields(itemsToRemove);           

            _formFields.Templates.required = false;
            _HideFormFields([_formFields.Templates], true); 
            const templateNote = document.getElementById("TemplateNote");
            if (templateNote) {
                templateNote.style.display = "none";
            }

            let serviceUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetGoalTemplates`;
            let offlineitems = JSON.parse(localStorage.getItem('GetGoalTemplates')) || [];     

            try {

                let response = offlineitems.length > 0 ? offlineitems : [];

                if (response.length === 0) {
                    try {
                        response = await getRestfulResult(serviceUrl);
                        if (response && response.length > 0) {
                            localStorage.setItem('GetGoalTemplates', JSON.stringify(response));
                        }
                    } catch (error) {
                        console.error("Error fetching data from API:", error);
                    }
                }
                else {
                    (async () => {
                        try {
                            let newItems = await getRestfulResult(serviceUrl);
                            if (newItems && newItems.length > 0 && JSON.stringify(newItems) !== JSON.stringify(offlineitems)) {
                                localStorage.setItem('GetGoalTemplates', JSON.stringify(newItems));
                            }
                        } catch (error) {
                            console.error("Error fetching data from API:", error);
                        }
                    })();
                }

                if (response?.length > 0) {        
                    _formFields.GoalCategory.$on('change', async function (value) {
                        await GoalCategoryTemplates(value, response);
                    });
                    await GoalCategoryTemplates(_formFields.GoalCategory.value, response);               
                    
                    _formFields.Templates.$on('change', async (selectedValue) => {
                        if (selectedValue?.length > 0) {
                            _formFields.Templates.disabled = true;

                            const selectedVal = selectedValue[0].trim();
                            const selectedItem = response.find(item => item.Description.trim() === selectedVal);

                            if (selectedItem) {
                                _formFields.Title.value = selectedItem.Description;
                                _formFields.Level.value = selectedItem.Level;
                                _formFields.Overview.value = selectedItem.Overview;
                            
                                _formFields.Title.disabled = true;
                                _formFields.Level.disabled = true;
                                _formFields.Overview.disabled = true;
                            }
                        } else {                        
                            _formFields.Templates.disabled = false;
                            _formFields.Title.value = '';
                            _formFields.Level.value = '';
                            _formFields.Overview.value = '';

                            _formFields.Title.disabled = false;
                            _formFields.Level.disabled = false;
                            _formFields.Overview.disabled = false;
                        }
                    });
                }
            }
            catch (error) {
                console.error("Error fetching data:", error);
            }            
        }
    }
    catch(err){
        console.log(err.message, err.stack);
    }
}

function hideButtons() {
    $('.fd-sp-datatable-toolbar-primary-commands').attr("style", "display: none !important;");
    $('.col-sm-12.Manager-Buttons').attr("style", "display: flex !important;justify-content:end;");
    $('.fd-grid.container-fluid').attr("style", "justify-content:space-between; gap:20px;");
}

var handleEditForm = async function () {   

    const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
    const svgGoals = `<svg height="24" width="24" version="1.1" id="_x32_" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" 
        viewBox="0 0 512 512"  xml:space="preserve">
        <style type="text/css">
            .st0{fill:#000000;}
        </style>
        <g>
            <path class="st0" d="M228.31,77.506c21.358-1.961,37.07-20.901,35.093-42.259c-1.994-21.342-20.917-37.063-42.268-35.077
                c-21.358,1.986-37.054,20.893-35.076,42.259C188.045,63.778,206.968,79.498,228.31,77.506z"/>
            <path class="st0" d="M368.479,388.205c7.133,6.194,17.918,5.605,24.341-1.332l0.458-0.482c6.406-6.928,6.185-17.682-0.499-24.341
                l-45.045-50.028l-16.48-55.888c-0.752,1.34-1.536,2.663-2.402,3.922c-8.481,12.191-21.841,19.994-36.646,21.359l-6.822,0.474
                l17.387,41.997c4.118,6.848,9.224,13.032,15.14,18.409L368.479,388.205z"/>
            <path class="st0" d="M214.419,351.568c9.168-2.255,14.977-11.275,13.253-20.541l-11.128-59.76l73.455-5.099
                c10.532-0.744,20.19-6.21,26.228-14.904c6.03-8.677,7.794-19.609,4.804-29.766l-2.754-9.355l-22.968-83.668l39.374,2.132
                l31.474,30.322c-1.52,1.846-2.28,4.281-1.668,6.798l2.566,10.401l-18.294,4.51c-4.339,1.07-6.978,5.45-5.908,9.78l12.322,50.005
                c1.078,4.322,5.449,6.97,9.772,5.915l77.671-19.152c4.339-1.078,6.962-5.458,5.9-9.789l-12.338-49.988
                c-1.063-4.33-5.434-6.986-9.764-5.915l-18.302,4.51l-2.574-10.41c-1.046-4.232-5.311-6.798-9.543-5.76l-1.92,0.474
                c-0.434-1.773-1.218-3.489-2.395-5.025l-30.779-40.363c-3.325-4.347-8.048-7.419-13.384-8.662l-50.626-18.122
                c-18.131-6.504-38.328-3.333-53.6,8.4l-70.244,53.894l-45.56-16.431c-7.142-3.178-15.484-0.123-18.891,6.879l-0.474,0.956
                c-1.683,3.489-1.928,7.476-0.662,11.12c1.275,3.652,3.955,6.642,7.436,8.309l54.916,26.213c6.83,3.268,14.781,3.235,21.595-0.082
                l29.268-19.463l20.091,57.587l-48.591,4.249c-8.187,0.744-15.622,5.099-20.247,11.881c-4.641,6.79-5.982,15.296-3.693,23.189
                l24.152,82.606c2.68,9.143,12.084,14.568,21.342,12.289L214.419,351.568z M397.819,159.212l0.351,0.204l2.574,10.402l-26.318,6.488
                l-2.483-10.083c4.885,4.004,11.856,4.306,16.929,0.392l0.295-0.238c2.124-1.634,3.611-3.783,4.461-6.136L397.819,159.212z"/>
            <polygon class="st0" points="239.266,387.895 212.245,371.586 212.245,512 299.754,512 299.754,364.07 289.467,358.611 	"/>
            <polygon class="st0" points="316.488,512 403.988,512 403.988,419.377 316.488,372.951 	"/>
            <polygon class="st0" points="420.722,428.258 420.722,512 494.455,512 494.455,467.38 	"/>
            <polygon class="st0" points="108.011,512 195.512,512 195.512,361.479 108.011,308.647 	"/>
            <polygon class="st0" points="17.545,512 91.278,512 91.278,298.54 17.545,254.026 	"/>
        </g>
    </svg>`;
    const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;

    setIconSource("overview-icon", svguserinfo);
    setIconSource("goals-icon", svgGoals);
    setIconSource("attachment-icon", svgattachment);

    formatingButtonsBar('Human Resources: Performance Management Goals');   

    if(_isEdit && _isPM){
      hideButtons();
    }

    if (_isChildForm) {
        return;
    }

    let arrayFields = [_formFields.Title, _formFields.StartDate, _formFields.EndDate, _formFields.Level, _formFields.GoalCategory, _formFields.Attachments];

    $(_formFields.ApprovalStatus.$parent.$el).hide();
    $(_formFields.ReviewedDate.$parent.$el).hide();

    _HideFormFields([_formFields.Submit, _formFields.EmployeeId, _formFields.EmployeePartyNumber, _formFields.ManagerId, _formFields.ManagerPartyNumber, _formFields.EmployeeCommentDate, _formFields.ManagerCommentDate, _formFields.DiscussionID, _formFields.Manager], true);

    if(_formFields.ApprovalStatus.value === 'Approved'){
        $('span').filter(function(){ return $(this).text() === submitDefault; }).parent().attr("disabled", "disabled");
        setPSErrorMesg('Manager has Approved this Goal');
        setPSHeaderMessage('Current Status:', '-24px');
        $(_formFields.ApprovalStatus.$parent.$el).show();
        fd.field('ApprovalStatus').disabled = true;
        $(_formFields.ReviewedDate.$parent.$el).show();
        fd.field('ReviewedDate').disabled = true;
        _DisableFormFields(arrayFields, true);
        disableRichTextFieldColumn(_formFields.Overview);
        return;
    }
    else {
        
        let isReject = _formFields.ApprovalStatus.value === 'Rejected' ? true : false;
        
        if(!isReject){

            if (!_isPM) {
                
                $('span').filter(function(){ return $(this).text() === submitDefault; }).parent().attr("disabled", "disabled");

                let status = _formFields.ApprovalStatus.value;
                let errorMessage = 'Waiting for PM Approval'
                if(status === 'Approved')
                errorMessage = `Manager has ${status}ed this Goal`
                else if(status === 'Rejected')
                errorMessage = `Manager has ${status}ed this Goal`

                setPSErrorMesg(errorMessage);
                setPSHeaderMessage('Current Status:', '-24px');
            }
        }
        else {
            
            $(_formFields.ApprovalStatus.$parent.$el).show();
            fd.field('ApprovalStatus').disabled = true;
            $(_formFields.ReviewedDate.$parent.$el).show();
            fd.field('ReviewedDate').disabled = true;
        }

        if (!isReject) {
            _DisableFormFields(arrayFields, true);
            disableRichTextField(_formFields.Overview.title);
        }

        if (_isPM) {
            _DisableFormFields(arrayFields, true);
            disableRichTextField(_formFields.Overview.title);
        }
    }
}

var handleDisplayForm = async function(){

    const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
    const svgGoals = `<svg height="24" width="24" version="1.1" id="_x32_" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" 
        viewBox="0 0 512 512"  xml:space="preserve">
        <style type="text/css">
            .st0{fill:#000000;}
        </style>
        <g>
            <path class="st0" d="M228.31,77.506c21.358-1.961,37.07-20.901,35.093-42.259c-1.994-21.342-20.917-37.063-42.268-35.077
                c-21.358,1.986-37.054,20.893-35.076,42.259C188.045,63.778,206.968,79.498,228.31,77.506z"/>
            <path class="st0" d="M368.479,388.205c7.133,6.194,17.918,5.605,24.341-1.332l0.458-0.482c6.406-6.928,6.185-17.682-0.499-24.341
                l-45.045-50.028l-16.48-55.888c-0.752,1.34-1.536,2.663-2.402,3.922c-8.481,12.191-21.841,19.994-36.646,21.359l-6.822,0.474
                l17.387,41.997c4.118,6.848,9.224,13.032,15.14,18.409L368.479,388.205z"/>
            <path class="st0" d="M214.419,351.568c9.168-2.255,14.977-11.275,13.253-20.541l-11.128-59.76l73.455-5.099
                c10.532-0.744,20.19-6.21,26.228-14.904c6.03-8.677,7.794-19.609,4.804-29.766l-2.754-9.355l-22.968-83.668l39.374,2.132
                l31.474,30.322c-1.52,1.846-2.28,4.281-1.668,6.798l2.566,10.401l-18.294,4.51c-4.339,1.07-6.978,5.45-5.908,9.78l12.322,50.005
                c1.078,4.322,5.449,6.97,9.772,5.915l77.671-19.152c4.339-1.078,6.962-5.458,5.9-9.789l-12.338-49.988
                c-1.063-4.33-5.434-6.986-9.764-5.915l-18.302,4.51l-2.574-10.41c-1.046-4.232-5.311-6.798-9.543-5.76l-1.92,0.474
                c-0.434-1.773-1.218-3.489-2.395-5.025l-30.779-40.363c-3.325-4.347-8.048-7.419-13.384-8.662l-50.626-18.122
                c-18.131-6.504-38.328-3.333-53.6,8.4l-70.244,53.894l-45.56-16.431c-7.142-3.178-15.484-0.123-18.891,6.879l-0.474,0.956
                c-1.683,3.489-1.928,7.476-0.662,11.12c1.275,3.652,3.955,6.642,7.436,8.309l54.916,26.213c6.83,3.268,14.781,3.235,21.595-0.082
                l29.268-19.463l20.091,57.587l-48.591,4.249c-8.187,0.744-15.622,5.099-20.247,11.881c-4.641,6.79-5.982,15.296-3.693,23.189
                l24.152,82.606c2.68,9.143,12.084,14.568,21.342,12.289L214.419,351.568z M397.819,159.212l0.351,0.204l2.574,10.402l-26.318,6.488
                l-2.483-10.083c4.885,4.004,11.856,4.306,16.929,0.392l0.295-0.238c2.124-1.634,3.611-3.783,4.461-6.136L397.819,159.212z"/>
            <polygon class="st0" points="239.266,387.895 212.245,371.586 212.245,512 299.754,512 299.754,364.07 289.467,358.611 	"/>
            <polygon class="st0" points="316.488,512 403.988,512 403.988,419.377 316.488,372.951 	"/>
            <polygon class="st0" points="420.722,428.258 420.722,512 494.455,512 494.455,467.38 	"/>
            <polygon class="st0" points="108.011,512 195.512,512 195.512,361.479 108.011,308.647 	"/>
            <polygon class="st0" points="17.545,512 91.278,512 91.278,298.54 17.545,254.026 	"/>
        </g>
    </svg>`;const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;

    setIconSource("overview-icon", svguserinfo);
    setIconSource("goals-icon", svgGoals);
    setIconSource("attachment-icon", svgattachment);

    formatingButtonsBar('Human Resources: Performance Management Goals');

    let dt = _formFields.goalsDT;
                    
    dt.ready().then(async function () {
        dt.buttons[0].visible = false;
        dt.buttons[dt.buttons.length - 1].visible = false;
        setTimeout(function() {
            if (_isDisplay) { 
                fixTableColumnsWidth();
                disableDataTableCheckBoxes(dt);
            }    
        }, 200);            
    });   
}

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

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;

    _currentUser = await pnp.sp.web.currentUser.get();

    let EmailCurrentUser = _currentUser.Email;
    //EmailCurrentUser = 'raji.zeitouny@dar.com';
    _UserIntheForm = EmailCurrentUser.toLowerCase();    

    let serviceUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeeManager&Email=${EmailCurrentUser}`;
    let response = await getRestfulResult(serviceUrl)

    if (response && response.length > 0) {       

        _manager = response[0].ManagerUPN; 
        _ManagerName = _manager.toLowerCase();
        _employeeId = response[0].EmployeeId;
        _managerId = response[0].ManagerId;
        _employeePartyNumber = response[0].EmployeePartyNumber;;
        _managerPartyNumber = response[0].ManagerPartyNumber;;

        if (_currentUser.Email.toLowerCase() === _manager.toLowerCase())
            _isPM = true;        

        // serviceUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=GetEmployeeIdFromUPN&Email=${_currentUser.Email}`;
        // let response1 = await getRestfulResult(serviceUrl)
        // if (response1)
        //     _employeeId = response1.employeeId;     

        if (formType === 'New'){
            if (_isChildForm) {                
                fd.field('Manager').value = _manager; 
                fd.field('Employee').value = EmailCurrentUser; 
                fd.field('Title').value = `Goal - ${_currentUser.Title} - ${formatDate(new Date())}`; 
                _HideFormFields([_formFields.Manager, fd.field('Employee'), _formFields.Title], true);
            }
            else {
                fd.field('EmployeeId').value = _employeeId;
                fd.field('EmployeePartyNumber').value = _employeePartyNumber;
                fd.field('ManagerId').value = _managerId;
                fd.field('ManagerPartyNumber').value = _managerPartyNumber;
                fd.field('Manager').value = _manager;
                fd.field('Submit').value = true;
            }
        }
    }
    else {
        
        if (formType === 'Display') {
            _ManagerName = fd.field('Manager').value.email.toLowerCase();
            _HideFormFields([_formFields.Manager], true);
        }
        else
            _ManagerName = fd.field('Manager').value && fd.field('Manager').value.EntityData && fd.field('Manager').value.EntityData.Email ? fd.field('Manager').value.EntityData.Email.toLowerCase() : null;       

        if (_ManagerName === _UserIntheForm) {_isPM = true; }
        else {
            alert("Sorry, but your aligned manager has not been set yet. Please contact the administrator for assistance.");
            fd.close();
        }
    }

    //_isPM = true; // DISABLE FOR LIVE
    console.log(`isPM: ${_isPM}`)

    if (_formType === 'New') {
        _isNew = true;
        await Goal_newForm();
    }
    else if (_formType === 'Edit') {
        _isEdit = true;
        _itemId = fd.itemId;
        localStorage.setItem('_isEdit', _isEdit);        
        if (!_isReviewForm)
            handleEditForm();
    }
    else if (_formType === 'Display') {
        _isDisplay = true;
        _HideFormFields([_formFields.Manager], true);
        if (_isReviewForm) {
            //setPSHeaderMessage('');
            setPSHeaderMessage('', '-25px', '20px');
            setPSErrorMesg(`Please click the Edit button to add your comments.`);
        }    
    }

    if (_isChildForm) {

        _formFields.goalsDT.dialogOptions = {
          width: '65%',
          height: '90%',
        //   open: function () {            
        //         //$('.fd-grid.container-fluid').attr("style", "margin-top: 10px !important;");
        //     },
        //     close: function() {                
        //         //$('.fd-grid.container-fluid').attr("style", "margin-top: -10px !important;");
        //     }
        }
        
        let GoalsEmpName = ""; fd.field('Employee').value;
        let GoalsMangeName = ""; fd.field('Manager').value;

        if (_formType === 'Edit') {
            GoalsEmpName = fd.field('Employee').value.DisplayText;
            GoalsMangeName = fd.field('Manager').value.DisplayText;
        }
        else {
            GoalsEmpName = fd.field('Employee').value.displayName;
            GoalsMangeName = fd.field('Manager').value.displayName;
        }

        $(fd.field('Employee').$parent.$el).hide();
        $(fd.field('Manager').$parent.$el).hide();

        var content = `<table role="grid" class="modern-table">
                <tr>
                    <th width="10%">Employee Name</th>
                    <td width="40%">${GoalsEmpName}</td>
                    <th width="10%">Manager Name</th>
                    <td width="40%">${GoalsMangeName}</td>
                </tr>                
        </table><br>`;

        $('#TopTable').append(content);
    }

    //  const endTime = performance.now();
    //  const elapsedTime = endTime - startTime;
    //  console.log(`extractValues: ${elapsedTime} milliseconds`);
}

var GoalCategoryTemplates = async function (value, response) {    

    if (value && value === 'IDP') {        
        _HideFormFields([_formFields.Templates], false);
        _formFields.Templates.required = true;
        const templateNote = document.getElementById("TemplateNote");
        if (templateNote) {
            templateNote.style.display = "block";
            templateNote.innerHTML = `
            <div style="margin-top: 15px;padding:8px; border:1px solid #ccc; border-radius:5px; background-color:#f9f9f9; font-size:14px; color:#333;box-shadow: rgba(0, 0, 0, 0.1) 2px 4px 4px;">
                <span style="margin-left: 5px;">
                    <strong style="color: red;">Note:</strong> IDP goal setting is a personal development objective for the Performance Review, not a training request. Employee must separately enroll in approved training program(s) through Learning & Development Department.                    
                </span>
            </div>`;           
        }
        _formFields.Templates.disabled = false;                     
        
        _formFields.Templates.widget.dataSource.data(['Please wait while retreiving ...']);                    

        const tempArray = response.map(d => d.Description.trim()).sort((a, b) => a.localeCompare(b));

        _formFields.Templates.widget.setDataSource({
            data: tempArray
        });                    
        
    } else {                    
        // _formFields.Templates.disabled = false;
        // _formFields.Templates.value = "";
        _formFields.Title.disabled = false; 
        _formFields.Level.disabled = false; 
        _formFields.Overview.disabled = false;
        _formFields.Templates.required = false; 
        _HideFormFields([_formFields.Templates], true); 
        const templateNote = document.getElementById("TemplateNote");
        if (templateNote) {
            templateNote.style.display = "none";
        }
    }
}

//#region THIS setCustomButtons FOR SINGLE ITEM FORM
var setCustomButtons = async function () {   

    if (!_isDisplay)
        fd.toolbar.buttons[0].style = "display: none;";

    else if ((_isDisplay && !_isChildForm && !_isReviewForm))
        fd.toolbar.buttons[0].style = "display: none;";

    fd.toolbar.buttons[1].style = "display: none;";    

    try {
        _HideFormFields([_formFields.EmployeeComment, _formFields.ManagerComment], true); //_formFields.Rating
    } catch (error) { }   
    
    if (_isNew) {
        if (_isChildForm)
            await setButtonActions("Save", "Save", '#dbdbdb', 'white');
        await setButtonActions("Accept", submitDefault, `${greenColor}`, 'white');
    }

    else if (_isReviewForm && _isEdit) { 
               
        _EmployeeName = _PerfomanceReviewFormFields.Employee.value.EntityData.Email.toLowerCase();//.DisplayText.toLowerCase();
        //_ManagerName = _PerfomanceReviewFormFields.Manager.value.DisplayText.toLowerCase();
        //_UserIntheForm = _currentUser.Title.toLowerCase(); 
        //_UserIntheForm = 'Bachir ElOjeil'.toLowerCase();
        _WorkflowStatus = _PerfomanceReviewFormFields.WorkflowStatus.value.toLowerCase();

        let hasAccess = false;       

        _HideFormFields([_PerfomanceReviewFormFields.FinalRate, _PerfomanceReviewFormFields.WorkflowStatus], true);
                
        [_PerfomanceReviewFormFields.EmployeeSignOff, _PerfomanceReviewFormFields.ManagerSignOff, _PerfomanceReviewFormFields.EmployeeSignOffDate, _PerfomanceReviewFormFields.ManagerSignOffDate].forEach(disableField);
          
        if (_WorkflowStatus === 'open') {

            if (_UserIntheForm === _ManagerName && _isPM) {                
                    
                //setPSHeaderMessage('');
                setPSHeaderMessage('', '-25px', '20px');
                setPSErrorMesg(`You're logged in as ${_PerfomanceReviewFormFields.Manager.value.DisplayText} (Manager). Please note that the employee has not yet filled in the goal comments.`);
            
                await setPreviewFormButtonActions("Accept", submitDefault, `${greenColor}`, 'white');

                let buttons = fd.toolbar.buttons;
                let submitButton = buttons.find(button => button.text === 'Submit');
                disableEnableButton(submitButton, true);

                setTimeout(function () {
                    disableDataTableCheckBoxes();
                }, 200);              

                hasAccess = true;
            }
            else if (_UserIntheForm === _EmployeeName) {                
                
                await setPreviewFormButtonActions("Accept", 'Submit', `${greenColor}`, 'white');

                let dt = _PerfomanceReviewFormFields.goalsDT;
                let disablebutton = true;
                let goals = dt.widget.dataItems(); // Get goal list data
                let idpCount = 0;
                let PersCount = 0;
                let compCount = 0;
                if (goals.length >= 5) {
                    goals.forEach(goal => {
                        if (goal.GoalCategory === 'IDP')
                            idpCount++;
                        
                        else if (goal.GoalCategory === 'Personal')
                            PersCount++;
                            
                        else if (goal.GoalCategory === 'Company values')
                            compCount++;
                    });
                }
                else
                    disablebutton = true;
                
                if (idpCount >= 1 && PersCount >= 3 && compCount >= 1)
                    disablebutton = false;                

                if (disablebutton) {
                    //setPSHeaderMessage('');  
                    setPSHeaderMessage('', '-25px', '20px');
                    setPSErrorMesg(`You're logged in as ${_PerfomanceReviewFormFields.Employee.value.DisplayText} (Employee). Unfortunately, your approved goals do not meet the required criteria.`);                
                    $('span').filter(function () { return $(this).text() === 'Submit'; }).parent().attr("disabled", "disabled");
                    setTimeout(function() {
                        disableDataTableCheckBoxes(); 
                    }, 200);
                }
                else {
                    //setPSHeaderMessage('');
                    setPSHeaderMessage('', '-25px', '20px');
                    setPSErrorMesg(`You're logged in as ${_PerfomanceReviewFormFields.Employee.value.DisplayText} (Employee). Please provide a comment for each goal.`);
                    _nextStatus = 'Sent to Manager';
                }

                // setTimeout(function() {
                //     disableDataTableCheckBoxes(); 
                // }, 200);                

                hasAccess = true;
            }
        }
        else if (_WorkflowStatus === 'sent to manager') {            

            if (_UserIntheForm === _ManagerName && _isPM) { 
                //setPSHeaderMessage('');
                setPSHeaderMessage('', '-25px', '20px');
                setPSErrorMesg(`You're logged in as ${_PerfomanceReviewFormFields.Manager.value.DisplayText} (Manager). Please provide a comment for each goal and select the final rating that best reflects the employee's performance.`);
                _nextStatus = 'Reviewed';
                _HideFormFields([_PerfomanceReviewFormFields.FinalRate], false);
                _PerfomanceReviewFormFields.FinalRate.required = true;
                _isFinalRatingSet = true;
                await setPreviewFormButtonActions("Accept", submitDefault, `${greenColor}`, 'white');
                hasAccess = true;
            }
            else if (_UserIntheForm === _EmployeeName) { 
                //setPSHeaderMessage('');
                setPSHeaderMessage('', '-25px', '20px');
                setPSErrorMesg(`You're logged in as ${_PerfomanceReviewFormFields.Employee.value.DisplayText} (Employee). Please note that your performance review is currently under manager review.`);                
                
                await setPreviewFormButtonActions("Accept", submitDefault, `${greenColor}`, 'white');

                let buttons = fd.toolbar.buttons;
                let submitButton = buttons.find(button => button.text === 'Submit');
                disableEnableButton(submitButton, true);

                setTimeout(function() {
                    disableDataTableCheckBoxes(); 
                }, 200);                

                hasAccess = true;
            }            
        }
        else if (_WorkflowStatus === 'published') {

            if (_UserIntheForm === _ManagerName && _isPM) { 
                //setPSHeaderMessage('');
                setPSHeaderMessage('', '-25px', '20px');
                setPSErrorMesg(`You're logged in as ${_PerfomanceReviewFormFields.Manager.value.DisplayText} (Manager). Please note that the performance review below has been ${_WorkflowStatus}.`);               
                await setPreviewFormButtonActions("Accept", submitDefault, `${greenColor}`, 'white');

                setTimeout(function() {
                    disableDataTableCheckBoxes(); 
                }, 200);

                let buttons = fd.toolbar.buttons;
                let submitButton = buttons.find(button => button.text === 'Submit');
                disableEnableButton(submitButton, true);
                
                hasAccess = true;

                let managerCheckBox = _PerfomanceReviewFormFields.ManagerSignOff.value;

                if (!managerCheckBox) {

                    [_PerfomanceReviewFormFields.ManagerSignOff].forEach(enabledField);

                    _PerfomanceReviewFormFields.ManagerSignOff.$on('change', async function (value) {
                        if (value) {
                            [_PerfomanceReviewFormFields.ManagerSignOffDate].forEach(enabledField);
                            _PerfomanceReviewFormFields.ManagerSignOffDate.value = new Date();
                            [_PerfomanceReviewFormFields.ManagerSignOffDate].forEach(disableField);

                            _isMgrSignOff = true;
                        
                            disableEnableButton(submitButton, false);
                        }
                        else {
                            [_PerfomanceReviewFormFields.ManagerSignOffDate].forEach(enabledField);
                            _PerfomanceReviewFormFields.ManagerSignOffDate.value = '';
                            [_PerfomanceReviewFormFields.ManagerSignOffDate].forEach(disableField);

                            _isMgrSignOff = false;

                            disableEnableButton(submitButton, true);
                        }
                    });
                }                
            }
            else if (_UserIntheForm === _EmployeeName) { 
                //setPSHeaderMessage('');
                setPSHeaderMessage('', '-25px', '20px');
                setPSErrorMesg(`You're logged in as ${_PerfomanceReviewFormFields.Employee.value.DisplayText} (Employee). Please note that the performance review below has been ${_WorkflowStatus}.`);              
                await setPreviewFormButtonActions("Accept", submitDefault, `${greenColor}`, 'white');          
              
                setTimeout(function() {
                    disableDataTableCheckBoxes(); 
                }, 200);  

                let buttons = fd.toolbar.buttons;
                let submitButton = buttons.find(button => button.text === 'Submit');
                disableEnableButton(submitButton, true);
                
                hasAccess = true;

                let employeeCheckBox = _PerfomanceReviewFormFields.EmployeeSignOff.value;

                if (!employeeCheckBox) {

                    [_PerfomanceReviewFormFields.EmployeeSignOff].forEach(enabledField);

                    _PerfomanceReviewFormFields.EmployeeSignOff.$on('change', async function (value) {
                        if (value) {
                            [_PerfomanceReviewFormFields.EmployeeSignOffDate].forEach(enabledField);
                            _PerfomanceReviewFormFields.EmployeeSignOffDate.value = new Date();
                            [_PerfomanceReviewFormFields.EmployeeSignOffDate].forEach(disableField);

                            _isEmpSignOff = true;
                        
                            disableEnableButton(submitButton, false);
                        }
                        else {
                            [_PerfomanceReviewFormFields.EmployeeSignOffDate].forEach(enabledField);
                            _PerfomanceReviewFormFields.EmployeeSignOffDate.value = '';
                            [_PerfomanceReviewFormFields.EmployeeSignOffDate].forEach(disableField);

                            _isEmpSignOff = false;

                            disableEnableButton(submitButton, true);
                        }
                    });
                }
            }

            _HideFormFields([_PerfomanceReviewFormFields.FinalRate], false);
            [_PerfomanceReviewFormFields.FinalRate].forEach(disableField);
        }
        else if (_WorkflowStatus === 'calibrated' || _WorkflowStatus === 'recalibration requested' || _WorkflowStatus === 'reviewed') {

            if ((_UserIntheForm === _ManagerName && _isPM) || (_UserIntheForm === _EmployeeName)) {

                //setPSHeaderMessage('');
                setPSHeaderMessage('', '-25px', '20px');
                setPSErrorMesg(`You're logged in as ${_currentUser.Title}. Please note that the performance review below has been ${_WorkflowStatus}.`);
                await setPreviewFormButtonActions("Accept", submitDefault, `${greenColor}`, 'white');
            
                setTimeout(function () {
                    disableDataTableCheckBoxes();
                }, 200);

                let buttons = fd.toolbar.buttons;
                let submitButton = buttons.find(button => button.text === 'Submit');
                disableEnableButton(submitButton, true);
            
                hasAccess = true;                
            }            
        }
        if (!hasAccess) {
            alert("Apologies, you have accessed a link that is not related to your.");
            fd.close();
        }
    }

    else if (!_isChildForm && _isEdit && _isPM && _formFields.ApprovalStatus.value === '') {   
        fd.toolbar.buttons.push({
            icon: 'Accept',
            class: 'btn-outline-primary',
            text: 'Approve',
            style: `background-color:${greenColor}; color: white;`,

            click: async function () {
                if (fd.isValid) {
                    showPreloader();
                    $(_formFields.ApprovalStatus.$parent.$el).show();
                    $(_formFields.ReviewedDate.$parent.$el).show();
                    fd.field('ApprovalStatus').value = 'Approved';
                    fd.field('ReviewedDate').value = new Date();
                    $(_formFields.ApprovalStatus.$parent.$el).hide();
                    $(_formFields.ReviewedDate.$parent.$el).hide();
                }
            }
        });

        fd.toolbar.buttons.push({
            icon: 'Cancel',
            class: 'btn-outline-primary',
            text: 'Reject',
            style: `background-color:${redColor}; color: white;`,
            click: async function () {
                if (fd.isValid) {
                    showPreloader();
                    $(_formFields.ApprovalStatus.$parent.$el).show();
                    $(_formFields.ReviewedDate.$parent.$el).show();
                    fd.field('ApprovalStatus').value = 'Rejected';
                    fd.field('ReviewedDate').value = new Date();
                    $(_formFields.ApprovalStatus.$parent.$el).hide();
                    $(_formFields.ReviewedDate.$parent.$el).hide();
                    fd.save();
                }
            }
        });
    }
        
    else if (_isEdit && !_isPM && !_isChildForm && _formFields.ApprovalStatus.value !== 'Approved' /*&& _formFields.ApprovalStatus.value === 'Rejected'*/) {
        await setButtonActions("Accept", submitDefault, `${greenColor}`, 'white');
    }
        
    else if (_isEdit && _isPM && _isChildForm) {
        await setButtonActions("Accept", "Approve", `${greenColor}`, 'white');
        await setButtonActions("ChromeClose", "Reject", "red", 'white');        
    }

    else if (_isEdit && !_isPM && _isChildForm) {
        await setButtonActions("Accept", "Submit", `${greenColor}`, 'white');        
    }
    
    if (!_isChildForm && !_isReviewForm) {   

        if (_isEdit && _isPM && _formFields.ApprovalStatus.value === 'Approved') {
            _HideFormFields([_formFields.EmployeeComment, _formFields.ManagerComment, _formFields.GoalID], false); //_formFields.Rating,   
            await setPreviewFormButtonActions("Accept", submitDefault, `${greenColor}`, 'white');
            disableRichTextFieldColumn(_formFields.EmployeeComment);

            _GoalID = _formFields.GoalID.value;
            console.log(`Goal ID = ${_GoalID}`);
            _HideFormFields([_formFields.GoalID], true); 

            let buttons = fd.toolbar.buttons;
            let submitButton = buttons.find(button => button.text === 'Submit');  
            
            renderAIDAHelper(25);

            let MgmComment = _formFields.ManagerComment.value;            
            
            if (MgmComment && MgmComment.trim() !== "") {
                disableEnableButton(submitButton, true);
                disableRichTextFieldColumn(_formFields.ManagerComment);
            }            
            else {
                disableEnableButton(submitButton, false);
                _isManagerComment = true;
            }            
        }

        else if (_isEdit && !_isPM && _formFields.ApprovalStatus.value === 'Approved') {           

            _HideFormFields([_formFields.EmployeeComment, _formFields.ManagerComment, _formFields.GoalID], false);
            await setPreviewFormButtonActions("Accept", submitDefault, `${greenColor}`, 'white');
            disableRichTextFieldColumn(_formFields.ManagerComment);         

            _GoalID = _formFields.GoalID.value;
            console.log(`Goal ID = ${_GoalID}`);
            _HideFormFields([_formFields.GoalID], true);       

            let buttons = fd.toolbar.buttons;
            let submitButton = buttons.find(button => button.text === 'Submit');
            
            renderAIDAHelper();

            let EMpComment = _formFields.EmployeeComment.value;       
            
            if (EMpComment && EMpComment.trim() !== "") {
                disableEnableButton(submitButton, true);
                disableRichTextFieldColumn(_formFields.EmployeeComment);
            }            
            else {
                disableEnableButton(submitButton, false);
                _isEmployeeComment = true;
            }
        }
    } 
    
    if (_isDisplay && _isReviewForm) {
        _HideFormFields([_PerfomanceReviewFormFields.WorkflowStatus], true);
    }
    
    await setButtonActions("ChromeClose", "Cancel", `${yellowColor}`, 'white');

    setToolTipMessages();

    //$('span').filter(function(){ return $(this).text() === submitText; }).parent().attr("disabled", "disabled");
}

const setButtonActions = async function(icon, text, bgColor, color){

    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          style: `background-color: ${bgColor}; color: ${color || 'white'};`, //color of font button//
        click: async function () {            
            
            if (text == "Close" || text == "Save" || text == "Cancel"){
                showPreloader();
                fd.close();
            }
            if (text == "Approve" || text == "Reject") {            
                
                let selectedItems = _formFields.goalsDT.selectedItems;
                let goals = _formFields.goalsDT.widget.dataItems();

                if (!selectedItems || selectedItems.length === 0) {
                    alert("Please select at least one goal before proceeding.");
                    return;
                }

                try {

                    showPreloader();

                    if (text === 'Approve')
                        localStorage.setItem('ApprovalStatus', 'Approved')
                    else localStorage.setItem('ApprovalStatus', 'Rejected')

                    let isApproved = text === "Approve" ? true : false;
                    const fetchPromises = selectedItems
                            //.filter(goal => goal.ApprovalStatus === 'Approved')
                            .map(goal => fetchResultToDynamics(goal, isApproved));
                    await Promise.all(fetchPromises);          
                    
                    await _formFields.goalsDT.refresh().then(() => {                      
                        setTimeout(() => {
                            fixTableColumnsWidth();                             
                        }, 250);
                    });
             
                    goals = _formFields.goalsDT.widget.dataItems();  
                    
                    let allReviewed = goals.every(goal => goal.ApprovalStatus === "Approved" || goal.ApprovalStatus === "Rejected");

                    if (allReviewed) {
                        hidePreloader();
                        alert("All Goals are Now Reviewed");
                        fd.save();
                    }
                    else {
                        hidePreloader();
                        alert("Goals Updated Successfully.");                        
                    }        
                }
                catch (error) {
                    console.error("Error updating Approval Status:", error);
                }
            }
            else if (text == submitDefault) {
                //if (confirm('Are you sure you want to Submit?')){
                if (fd.isValid) {

                    showPreloader();

                    if (_isChildForm && _isNew) {
                        //await handleSubmit();                       
                        console.log('new child form')
                        fd.save();
                    }

                    else {
                        if (_isChildForm)
                            fd.save();
                        else {
                            if (_isEdit && _formFields.ApprovalStatus.value === 'Approve')
                                await fetchResultToDynamics(null, true).then(() => { fd.save(); })
                            else {
                                showPreloader();
                                if (_isEdit && !_isPM && _formFields.ApprovalStatus.value === 'Rejected') {
                                    fd.field('ApprovalStatus').value = '';
                                    fd.field('ReviewedDate').value = '';
                                }
                                fd.save();
                            }
                        }
                    }
                }
            }
        }
    });
}

const setPreviewFormButtonActions = async function (icon, text, bgColor, color) {   

    fd.toolbar.buttons.push({
          icon: icon,
          class: 'btn-outline-primary',
          text: text,
          style: `background-color: ${bgColor}; color: ${color || 'white'};`, //color of font button//
        click: async function () { 

            if (fd.isValid) {

                showPreloader();                 
                
                let errorMessage = 'There was an issue with the integration to Dynamics. Please try again later or contact the IT department if the issue persists.';

                if (_isEmployeeComment) {

                    _HideFormFields([_formFields.EmployeeCommentDate], false);
                    _formFields.EmployeeCommentDate.value = new Date();
                    _HideFormFields([_formFields.EmployeeCommentDate], true);                 

                    let EmpComment = _formFields.EmployeeComment.value;

                    if (EmpComment) {
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(EmpComment, 'text/html');
                        EmpComment = doc.body.textContent.trim();
                    }

                    let updateData = {
                        DiscussionWorkerPersonnelNumber: _formFields.EmployeeId.value,
                        DiscussionId: _formFields.DiscussionID.value[0].LookupValue,        
                        GoalId: _formFields.GoalID.value,
                        GoalHeadingId: _formFields.GoalCategory.value, //_formFields.EndDate.value,        
                        GoalWorkerPersonnelNumber: _formFields.EmployeeId.value,
                        CommenterPartyNumber: _formFields.EmployeePartyNumber.value,
                        // CommentDateTime: new Date().toISOString().split('T')[0],
                        Comment: EmpComment
                    };                  
                    
                    let apiUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=PostDiscussionGoalComments`;
                    let result = await fetchCommentsToDynamics(apiUrl, updateData);

                    if (!result) {
                        hidePreloader();
                        alert(errorMessage);
                        return;
                    }
                }
                else if (_isManagerComment) {                   

                    _HideFormFields([_formFields.ManagerCommentDate], false);
                    _formFields.ManagerCommentDate.value = new Date();
                    _HideFormFields([_formFields.ManagerCommentDate], true);

                    let MnGComment = _formFields.ManagerComment.value;

                    if (MnGComment) {
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(MnGComment, 'text/html');
                        MnGComment = doc.body.textContent.trim();
                    }

                    let updateData = {
                        DiscussionWorkerPersonnelNumber: _formFields.EmployeeId.value,
                        DiscussionId: _formFields.DiscussionID.value[0].LookupValue,        
                        GoalId: _formFields.GoalID.value,
                        GoalHeadingId: _formFields.GoalCategory.value, //_formFields.EndDate.value,        
                        GoalWorkerPersonnelNumber: _formFields.EmployeeId.value, //_formFields.EmployeeId.value,  
                        CommenterPartyNumber: _formFields.ManagerPartyNumber.value,                     
                        Comment: MnGComment
                    };                    
                    
                    let apiUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=PostDiscussionGoalComments`;
                    let result = await fetchCommentsToDynamics(apiUrl, updateData);

                    if (!result) {
                        hidePreloader();
                        alert(errorMessage);
                        return;
                    }
                }           
                else if (_nextStatus && _nextStatus.trim() !== "") {

                    _HideFormFields([_PerfomanceReviewFormFields.WorkflowStatus], false);
                    _PerfomanceReviewFormFields.WorkflowStatus.value = _nextStatus;
                    _HideFormFields([_PerfomanceReviewFormFields.WorkflowStatus], true);

                    if (_isFinalRatingSet) {                       
                        
                        let EmpFinalRate = _PerfomanceReviewFormFields.FinalRate.value;
                        const ratingNumber = ratings[EmpFinalRate];

                        let updateData = {
                            PersonnelNumber: _PerfomanceReviewFormFields.PersonnelNumber.value,
                            DiscussionId: _PerfomanceReviewFormFields.Discussion.value,        
                            FinalEmployeeRatingId: ratingNumber,
                            RatingModelId: 'Standard'                           
                        };                     
                        
                        let apiUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=PostDiscussionsRating`;
                        let result = await fetchCommentsToDynamics(apiUrl, updateData);

                        if (!result) {
                            hidePreloader();
                            alert(errorMessage);
                            return;
                        }                       
                    }                   
                }
                else if (_isMgrSignOff) {
                    let ManagerSignOffDateVal = _PerfomanceReviewFormFields.ManagerSignOffDate.value;                    
                }
                else if (_isEmpSignOff) {
                    let EmployeeSignOffDateVal = _PerfomanceReviewFormFields.EmployeeSignOffDate.value                    
                }
            
                fd.save();                
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

function formatingButtonsBar(titelValue) {  

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
    $('.fd-grid.container-fluid').attr("style", "margin-top: -2px !important; padding: 10px;");
    $('.fd-form-container.container-fluid').attr("style", "margin-top: -22px !important;");    

    if (_module === 'PMG') {
        $('.fd-form-container.container-fluid').attr("style", "margin-top: 5px !important;");
        if(_isEdit) $('.fd-grid.container-fluid.divGrid').attr("style", "margin-top: 18px !important; padding: 10px;");       
    }

    const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';
    const linkElement = `<a href="${_spPageContextInfo.webAbsoluteUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">
                          <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">${titelValue}</a>`;
    $('span.o365cs-nav-brandingText').html(linkElement);

    $('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');
    
    $('.border-title').each(function () {
        $(this).css({
            'margin-top': '-25px', /* Adjust the position to sit on the border */
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

fd.spSaved(async function(result){

    try {

        if (!_isEmployeeComment && !_isManagerComment) {
      
            _itemId = result.Id;
            let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${_itemId}</Value></Eq></Where>`;

            if (_isEdit) {
                let status = _isChildForm ? localStorage.getItem('ApprovalStatus') : _formFields.ApprovalStatus.value;

                if (status === 'Approved')
                    await updateItem(_itemId, query, 'Edit', status)
                else if (status === 'Rejected')
                    await updateItem(_itemId, query, 'Edit', status)
                else await updateItem(_itemId, query, 'New', '');
            }
            else { // New
                let newGoals = localStorage.getItem('newGoals');
                let submitGoals = newGoals ? JSON.parse(newGoals) : null;

                if (_isNew && _isChildForm) {
                    _isChildForm = true;
                    localStorage.removeItem('newGoals');
                    await updateItem(_itemId, query, 'New', '');

                }
            }
        }    
    }
    catch (e) {
        console.log(e);
    }
});

var getRestfulResult = async function(serviceUrl){

    try {

        const response = await Promise.race([
          fetch(serviceUrl, {
              method: 'GET',
              headers: { 'Content-Type': 'application/json' },
          }),
          new Promise((_, reject) =>
              setTimeout(() => reject(new Error('Request timed out')), 30000) // 20 seconds timeout
          )
        ]);
   
        const responseText = await response.text();

        let data;
        try {
            data = JSON.parse(responseText);
        } catch (error) {
            throw new Error(`Unexpected response format: ${error.message}`);
        }
        
        if (data.detail) {
            throw new Error(data.detail);
        }

        return data; 

    } catch (error) {
      console.error('Error getRestfulResult:', error.message);
      throw error; // Rethrow or handle as needed
      showPreloader();
    }
}

const updateItem = async function(formItemId, query, operation, status){

    localStorage.removeItem('ApprovalStatus')
    if (operation === 'New') {
        if (_isChildForm)
            await _sendEmail('CPG', 'CPG_New_Email', query, '', 'CPG_New', '', _currentUser)
        else if(_manager) await _sendEmail(_module, 'PMG_New_Email', query, '', 'PMG_New', '', _currentUser);
    }

    else if(operation === 'Edit' && status === 'Rejected'){

        if (_isChildForm)
            await _sendEmail(_module, 'CPG_Reject_Email', query, '', 'CPG_Reject', '', _currentUser);
        else {
            if (!_isPM)
                await _sendEmail(_module, 'PMG_New_Email', query, '', 'PMG_New', '', _currentUser);
            else await _sendEmail(_module, 'PMG_Reject_Email', query, '', 'PMG_Reject', '', _currentUser);
        }
    }

    else if (operation === 'Edit' && status === 'Approved'){
        if (_isChildForm)
            await _sendEmail(_module, 'CPG_Approve_Email', query, '', 'CPG_Approve', '', _currentUser);
        else await _sendEmail(_module, 'PMG_Approve_Email', query, '', 'PMG_Approve', '', _currentUser);
    }
}

const fetchResultToDynamics = async function (goal, isApproved) {
    
    let apiUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=PostEmployeeGoals`;

    let overview = _isChildForm ? goal.Overview : _formFields.Overview.value;   

    //    let dateFinished = _isChildForm ? goal.DateFinished.value : _formFields.DateFinished.value;
    //    if(!isNullOrEmpty(dateFinished)){
    //      let date = new Date(dateFinished);
    //      dateFinished = date.toISOString().split('T')[0];
    //    }

    let dateStartTemp = _isChildForm ? goal.StartDate : _formFields.StartDate.value;
    let date1 = new Date(dateStartTemp);
    let dateStart = date1.toISOString().split('T')[0];

    let dateEndTemp = _isChildForm ? goal.EndDate : _formFields.EndDate.value;
    let date2 = new Date(dateEndTemp);
    let dateEnd = date2.toISOString().split('T')[0];

    // let stat = _isChildForm ? goal.Status.value : _formFields.Status.value;
    // stat = !isNullOrEmpty(stat) ? stat : ''

    if (!isApproved) {

        try {
            if (_isChildForm) {
                const goalsList = await _web.lists.getByTitle(_GoalsListName);
                await goalsList.items.getById(goal.ID).update({
                    ApprovalStatus: 'Rejected',
                    ReviewedDate: new Date()                    
                });
                console.log(`Updated item with ID: ${goal.ID}`);
            }
        } catch (error){
            console.error(`Failed to update item with ID: ${goal.ID}`, error);
        }

        return;
    }

    else {

        try {
            if (_isChildForm) {
                const goalsList = await _web.lists.getByTitle(_GoalsListName);
                const item = await goalsList.items.getById(goal.ID).get();
                if (item) {                   
                    overview = item.Overview;
                    _employeeId = item.EmployeeId;
                }
                            }
        } catch (error){
            console.error(`Failed to update item with ID: ${goal.ID}`, error);
        }      
    }
    
    if (overview) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(overview, 'text/html');
        overview = doc.body.textContent.trim();
    }

    let updateData = {
        PersonnelNumber: !isNullOrEmpty(_employeeId) ? _employeeId : '',
        GoalHeadingId: _isChildForm ? goal.GoalCategory : _formFields.GoalCategory.value,
        //Status: '', //stat,
        //DateFinished: null, //dateFinished,
        StartDate: dateStart, //_formFields.StartDate.value,
        EndDate: dateEnd, //_formFields.EndDate.value,
        //PercentComplete: 0, //!isNullOrEmpty(_formFields.PercentComplete.value) ? _formFields.PercentComplete.value : 0,
        Description: _isChildForm ? goal.Title : _formFields.Title.value, //_formFields.Title.value,
        GoalLevel: _isChildForm ? goal.Level : _formFields.Level.value, //_formFields.Level.value,
        Overview: overview
    };

    const response = await fetch(apiUrl, {
        method: "POST", // or "PATCH" depending on the API
        headers: {
            "Content-Type": "application/json",
            //"Authorization": "Bearer YOUR_ACCESS_TOKEN", // Include if the API requires authentication
        },
        body: JSON.stringify(updateData),
    });  

    let data;
    const responseText = await response.text();

    try {

        data = JSON.parse(responseText);        

        if (response.ok && isApproved) {
            try {
                
                if (_isChildForm) {
                    console.log(`Goal ID = ${data.Goal}`);
                    const goalsList = await _web.lists.getByTitle(_GoalsListName);
                    await goalsList.items.getById(goal.ID).update({
                        ApprovalStatus: isApproved ? 'Approved' : goal.ApprovalStatus,
                        ReviewedDate: new Date(),
                        GoalID: data.Goal,
                        IsSynched: true
                    });
                    console.log(`Updated item with ID: ${goal.ID}`);
                }
            } catch (error) {
                console.error(`Failed to update item with ID: ${goal.ID}`, error);
            }
        }        
    }
    catch (err) {
        console.error(`Failed to update item with ID: ${goal.ID}`, responseText);
        alert(`Failed to update item with ID: ${goal.ID}: ${responseText}`);
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);
    }          
}

const fetchCommentsToDynamics = async function (apiUrl, updateData) {

    const response = await fetch(apiUrl, {
        method: "POST", // or "PATCH" depending on the API
        headers: {
            "Content-Type": "application/json",
            //"Authorization": "Bearer YOUR_ACCESS_TOKEN", // Include if the API requires authentication
        },
        body: JSON.stringify(updateData),
    });
    
    try {

        if (response.ok) {          
            const text = await response.text(); 
            if (text.toLowerCase().includes("error response from server".toLowerCase())) {
                console.error("An error occurred: ", text);
                return false; 
            }
            console.log(`Successfully Posted item`);
            return true; // Return the API response data
        }
        else 
            return false;                
    }
    catch (err) {      
        console.error(`Failed to update item - Goal ID: ${updateData.GoalId}, Discussion ID: ${updateData.DiscussionId}`, err.stack);   
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);
    }    
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

function formatDate(dateString) {

    if (!dateString) return ''; // Handle empty values gracefully
    let date = new Date(dateString);
    return date.toLocaleDateString(undefined, { 
        year: 'numeric', 
        month: 'short', 
        day: 'numeric' 
    }); // Example: Mar 20, 2025

}

function formatDateforSPQuery(dateString) {
    const StartDateinputDate = new Date(dateString);
    const StartDateformattedDate = StartDateinputDate.toISOString().split('.')[0] + 'Z';
    return StartDateformattedDate;
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

function disableRichTextField(fieldname){

	let elem = $(fd.field(fieldname).$el).find('.k-editor tr');

	elem.each(function(index, element){

	 if(index === 0)
		$(element).remove()

	 else if(index === 1){

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
	   }
	})
}

function disableDataTableCheckBoxes(dt) { 
    
    document.querySelectorAll('.k-checkbox').forEach(checkbox => {
        checkbox.disabled = true;
    });
    
    setTableStyle(true);
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

function renderAIDAHelper(marginBottom = 8) {  

    const aidaContainer = document.getElementById("idAIDA");

    if (!aidaContainer) return;

    aidaContainer.innerHTML = `
        <div style="margin-top:0px;margin-bottom:${marginBottom}px; padding:8px; border:1px solid #ccc; border-radius:5px; background-color:#f9f9f9; font-size:14px; color:#333;box-shadow: rgba(0, 0, 0, 0.1) 2px 4px 4px;">
            <span style="margin-left: 5px;">
                <strong style="color: #FF9800;">Disclaimer:</strong> For assistance in filling in your goal, you can use our AI platform 
                <a href="#" id="aidaLink" style="text-decoration:underline; color:#0078d4;"><strong>AIDA</strong></a>.
            </span>
        </div>
    `;

    const link = document.getElementById("aidaLink");

    if (link) {
        link.addEventListener("click", function (e) {
            e.preventDefault();

            const windowWidth = window.innerWidth;
            const windowHeight = window.innerHeight;

            const newWindowWidth = windowWidth; // Full width
            const newWindowHeight = windowHeight;  // Full height

            const newWindowTop = Math.floor((windowHeight - Math.floor(windowHeight * 0.9)) / 2) + 100;
            const newWindowLeft = Math.floor((windowWidth - Math.floor(windowWidth * 0.5)) / 2) + 150;

            const windowFeatures = [
                `width=${newWindowWidth}`,
                `height=${newWindowHeight}`,
                `top=${newWindowTop}`,
                `left=${newWindowLeft}`,
                'resizable=yes',
                'scrollbars=yes'
            ].join(',');

            window.open('https://askaida.dar.com/', 'AIDA', windowFeatures);
        });
    }
}

//#endregion

//#region CHILD FORM
let applyDataTable = async function (){
    customizeColumns();
    _formFields.goalsDT.widget.bind('dataBound', customizeColumns);  // Apply customizations after data is bound
    _formFields.goalsDT.widget.bind('change', customizeColumns);  // Apply customizations after data change
}

function customizeColumns()
{
    let goalsList = _formFields.goalsDT;    
        // Loop through data and append the custom status text, icon, and completion details
    let data = goalsList.widget.dataSource.view(); // Get the data from the DataTable

        data.forEach(function (dataItem) {
           // ApprovalStatus Handling
            let approvalStatus = dataItem.ApprovalStatus;
            let approvalText = '';
            let approvalIcon = '';

            if (approvalStatus === "Approved") {
                approvalIcon = '✔️';
                approvalText = 'Approved';
            } else if (approvalStatus === "Rejected") {
                approvalIcon = '❌';
                approvalText = 'Rejected';
            } else {
                approvalIcon = '⏳';
                approvalText = 'Pending';
            }
            dataItem.ApprovalStatus = `<span class="approval-status">${approvalIcon} ${approvalText}</span>`;
    });

    // Re-render the DataTable after modifications
    goalsList.widget.refresh(); // Refresh to reflect data changes
}

function validateGoals(addRow) {   
    
    let goals = _formFields.goalsDT.widget.dataItems(); // Get goal list data
    let idpCount = 0;
    let PersCount = 0;
    let compCount = 0;
    let isAllApproved = true;
    let isAllReviewed= true;
    let patternResult = [], allResult = [];  
    let disableButton = true;

    if (goals.length >= 5) {
        goals.forEach(goal => {
            if (goal.GoalCategory === 'IDP' && goal.ApprovalStatus !== 'Rejected' && idpCount < 1) {
                patternResult.push({ Id: goal.ID, ApprovalStatus: goal.ApprovalStatus });
                idpCount++;
            }
            else if (goal.GoalCategory === 'Personal' && goal.ApprovalStatus !== 'Rejected' && PersCount < 3) {
                patternResult.push({ Id: goal.ID, ApprovalStatus: goal.ApprovalStatus });
                PersCount++;
            }
            else if (goal.GoalCategory === 'Company values' && goal.ApprovalStatus !== 'Rejected' && compCount < 1) {
                patternResult.push({ Id: goal.ID, ApprovalStatus: goal.ApprovalStatus });
                compCount++;
            }

            if (goal.ApprovalStatus !== 'Approved')
                isAllApproved = false;

             if (goal.ApprovalStatus === '')
                isAllReviewed = false;

            allResult.push({ Id: goal.ID, ApprovalStatus: goal.ApprovalStatus });
        });
    }

    if (_isEdit) {       

        if (isAllApproved) {
            //setTimeout(() => { setPageReadOnly('All Goals are Reviewed.'); }, 100);
            setPSHeaderMessage('', '-25px', '20px');
            setPSErrorMesg('All Goals are Reviewed.');  
            $('span').filter(function () { return $(this).text() == 'Approve'; }).parent().css('color', '#737373').attr("disabled", "disabled");
            $('span').filter(function () { return $(this).text() == 'Reject'; }).parent().css('color', '#737373').attr("disabled", "disabled");
            _formFields.goalsDT.buttons[0].visible = false;
        }
        else if (_isPM){

            setPSHeaderMessage('', '-25px', '20px');
            setPSErrorMesg('Please review the goals submitted by the employee. You can batch approve or reject them as needed.');                     
        }

        if (!_isPM) {       

            setTableStyle(true);
            setPSHeaderMessage('', '-25px', '20px');         

            let emptyItems = allResult.filter(item => item.ApprovalStatus === '');
            if (!addRow && emptyItems.length > 0) {
                setTimeout(() => { setPSErrorMesg('Pending Manager Approval.'); }, 100);
                let dt = _formFields.goalsDT;
                dt.buttons[0].visible = false;
                let buttons = fd.toolbar.buttons;
                let submitButton = buttons.find(button => button.text === 'Submit');
                if (submitButton) {
                    submitButton.visible = false;
                }
            }  
            
            let isMatched = false;
            if (idpCount >= 1 && PersCount >= 3 && compCount >= 1)
                isMatched = true;

            let rejectedItems = allResult.filter(item => item.ApprovalStatus === 'Rejected');
            if (!addRow && emptyItems.length === 0 && rejectedItems.length > 0) {
                if (!isMatched) {
                    setTimeout(() => { setPSErrorMesg('Please add a new goal for review to replace the rejected one.'); }, 100);
                    disableButton = false;
                }
                else {
                    setTimeout(() => { setPSErrorMesg('All goals have been reviewed and meet the required criteria.'); }, 100);                
                    let dt = _formFields.goalsDT;
                    dt.buttons[0].visible = false;
                    let buttons = fd.toolbar.buttons;
                    let submitButton = buttons.find(button => button.text === 'Submit');
                    if (submitButton) {
                        submitButton.visible = false;
                    }
                }
            }
            else {
                $('span').filter(function () { return $(this).text() === submitDefault; }).parent().attr("disabled", "disabled");                
            }
        }
    }

    //if (goals.length >= 5 && idpCount >= 1 && PersCount>=3 &&compCount>=1){
    let rejectedItems = patternResult.filter(item => item.ApprovalStatus === 'Rejected');
    if (patternResult.length >= 5 && rejectedItems.length === 0) {
        console.log("Eligible for submission");
        console.log("goals length:", goals.length);
        console.log("ip count:", idpCount);
        console.log("personal count:", PersCount);
        console.log("company values count:", compCount);

        if ((_isEdit && _isPM) || _isNew)
            $('span').filter(function () { return $(this).text().trim() === submitDefault }).parent().removeAttr("disabled");        
        //setPSErrorMesg('', true)

        if (addRow)
          localStorage.setItem('newGoals', true);
    }
    else{
        console.log(" NOT Eligible for submission");
        console.log("goals length:",goals.length);
        console.log("ip count:",idpCount);
        console.log("personal count:",PersCount);
        console.log("company values count:", compCount);
        if (disableButton)
            $('span').filter(function () { return $(this).text() === submitDefault; }).parent().attr("disabled", "disabled");
        // if (_isEdit && !_isPM) {
        //     setPSErrorMesg(`Your goals below are under review.`);
        //     setPSHeaderMessage(' ');
        // }
        // if (_isEdit && _isPM) {
        //     setPSErrorMesg(`Please review the goals submitted by the employee. You can batch approve or reject them as needed.`);
        //     setPSHeaderMessage(' ');
        // }
    }
}

function initializeEvents() {

    let goalsList = _formFields.goalsDT;    

    if (_isNew) {
        setPSHeaderMessage('', '-25px', '20px');
        setPSErrorMesg(`Please note that the employee is required to submit at least 5 goals: 3 Personal, 1 aligned with Company Values, and 1 Individual Development Plan (IDP).`);  
        goalsList.widget.bind('dataBound', validateGoals);        
    }
    else if(_isEdit)
        goalsList.widget.bind('dataBound', function () {
        validateGoals(true);
    });
    // goalsList.widget.bind('change', validateGoals);
        // goalsList.widget.bind('save', validateGoals);
    // console.log("Events bound successfully.");

    // goalsList.$on('edit', async function(item){
    //     if (item.type === 'add') {
    //         await validateGoals(true);
    //     }
    // });

}

function fixTableColumnsWidth() {

    let dt = _PerfomanceReviewFormFields.goalsDT; //_formFields.goalsDT;
    var Clientwidth = dt.$el.clientWidth;

    let widthMultiplier = 1;
    if (_isReviewForm) 
        widthMultiplier = (_isPM) ? 0.95 : 0.96; // 95% if PM, 96% if not PM
    else 
        widthMultiplier = 0.95; // Default to 95% if not in review form    
    Clientwidth *= widthMultiplier; // Apply the multiplier

    var Rwidget = dt.widget;
    var columns = Rwidget.columns;
    var ColumnsLength = columns.length;

    let totalAssignedWidth = 0;
    let columnWidths = {
        'ID': 40,
        'Level': 130,
        'GoalCategory': 140,
        'StartDate': 140,
        'EndDate': 140,
        'ReviewedDate': 140,
        'ApprovalStatus': 140 
    };

    if (_isReviewForm) {

        columnWidths = {
            'ID': 40,
            'Level': 150,
            'GoalCategory': 160,
            //'GoalID': 100,
            'StartDate': 160,
            'EndDate': 160,
            'ReviewedDate': 160,
            'EmployeeCommentDate': 180,
            'ManagerCommentDate': 180
        };
    }

    // First, set the widths for predefined columns
    for (let i = 1; i < ColumnsLength; i++) {
        let field = columns[i].field;

        if (columnWidths.hasOwnProperty(field)) {
            let reviewedWidth = columnWidths[field];
            totalAssignedWidth += reviewedWidth;
            dt._columnWidthViewStorage.set(field, reviewedWidth);
            Rwidget.resizeColumn(columns[i], reviewedWidth);
        }
    }

    // Assign remaining width to LinkTitle || Title
    let remainingWidth = Clientwidth - totalAssignedWidth;
    for (let i = 1; i < ColumnsLength; i++) {
        let field = columns[i].field;

        if (field === 'LinkTitle' || field === 'Title') {
            dt._columnWidthViewStorage.set(field, remainingWidth);
            Rwidget.resizeColumn(columns[i], remainingWidth);
            break; // Exit loop since LinkTitle is found
        }
    }

    //Background Colors based on condition
    var rows = Rwidget._data;

    let validEmployeeCount = 0;
    let validManagerCount = 0;

    for (let i = 0; i < rows.length; i++) {        

        let bkCol = '#fff3e0';    // Orange
        let MNGbkCol = '#d0f4de'; // (minty green)

        let isEmployeeCommentDate, isManagerCommentDate;

        if (_isReviewForm) {
            isEmployeeCommentDate = rows[i].EmployeeCommentDate;
            isManagerCommentDate = rows[i].ManagerCommentDate;
        } else {
            isEmployeeCommentDate = rows[i].ApprovalStatus;
            bkCol = (isEmployeeCommentDate === 'Approved') ? MNGbkCol : '#fcd0d3ba';
        }

        if (isEmployeeCommentDate && isEmployeeCommentDate.trim() !== "") {
            validEmployeeCount++;
            const row = $(dt.$el).find('tr[data-uid="' + rows[i].uid+ '"');
            row[0].style.cssText = `background: ${bkCol} !important;`;            
        }
        if (isManagerCommentDate && isManagerCommentDate.trim() !== "") {
            validManagerCount++;
            const row = $(dt.$el).find('tr[data-uid="' + rows[i].uid+ '"');
            row[0].style.cssText = `background: ${MNGbkCol} !important;`;            
        }
    }    

    if (_isReviewForm) {        
        _isAllEmployeeCommentDate = (validEmployeeCount === rows.length);
        _isAllManagerCommentDate = (validManagerCount === rows.length);

        enableDisableSubmitButton();
    }

    //Centralize the fields and wrap it
    let tables = $("table[role='grid']");
    tables.each(function(tblIndex, tbl){
        $(tbl).find('tr').each(function(trIndex, tr) {

    	    if (trIndex === 0 ){
    		   let childs = tr.children;
    		   for (let i = 0; i < childs.length; i++) {
                    if (i !== 2) { // Skip child[2]
                        childs[i].style.textAlign = 'center';
                    }
                }
    		}

    	   $(tr).find('td').each(function(tdIndex, td) {
                let $td = $(td);

                if (tdIndex !== 2)
                    td.style.textAlign = 'center';             

                if(_formType !== 'Display')
                    $td.css('whiteSpace', 'nowrap');
    		});
        });
    });
}

function setTableStyle(disableForReviewer) {
   
    var approvedBgColor = 'linear-gradient(to left, rgb(235, 245, 230), rgb(210, 240, 230), rgb(190, 230, 235), rgb(180, 200, 245), rgb(200, 190, 245), rgb(200, 190, 245), rgb(190, 185, 245))'

    let dt = _formFields.goalsDT;   
    let rows = dt.widget._data;

    rows.forEach(row => {      
        if (disableForReviewer || row.ApprovalStatus === 'Approved' || row.ApprovalStatus === 'Rejected') {
            const rowElement = $(dt.$el).find('tr[data-uid="' + row.uid + '"]')[0];
            if (rowElement) {

                let checkBox = $(rowElement).children().find('input[type="checkbox"]')[0];
                if (checkBox)
                    checkBox.disabled = true;

                // if(!disableForReviewer)
                // rowElement.style.background = approvedBgColor;
            }
        }
    });   
}

function setPageReadOnly(mesg) {
    //_formFields.goalsDT.readonly = true;
    disableDataTableCheckBoxes();
    $('span').filter(function () { return $(this).text() == submitDefault; }).parent().css('color', '#737373').attr("disabled", "disabled");
    setPSErrorMesg(mesg);
    setPSHeaderMessage(' ');
}

let linkdtRecords = async function (dt) {

    let apiUrl = `${_webUrl}/_layouts/15/NewsLetter/HSEIncidentForm.aspx?command=PostDiscussionGoals`;

    let updateResult = [];
    let formID = fd.itemId;    
    let PerformanceReviewItems = dt.widget.dataItems();

    if (PerformanceReviewItems.length > 0) {

        PerformanceReviewItems.forEach(PerformanceReviewItem => {            
            updateResult.push({ GoalId: PerformanceReviewItem.ID, ReviewID: formID, GoalCatg:  PerformanceReviewItem.GoalCategory});             
        });

        for (const updateItem of updateResult) {

            try {

                const goalItem = await _web.lists.getByTitle(_GoalsListName).items.getById(updateItem.GoalId).select("GoalID,ReviewID/Id").expand("ReviewID").get();         
                
                if (!goalItem.ReviewID || !goalItem.ReviewID.Id)
                {                    
                    let updateData = {
                        PersonnelNumberDiscussion: _PerfomanceReviewFormFields.PersonnelNumber.value,
                        Discussion: _PerfomanceReviewFormFields.Discussion.value,        
                        Goal: goalItem.GoalID,
                        GoalHeading: updateItem.GoalCatg,                           
                    };                  
                    
                    let result = await fetchCommentsToDynamics(apiUrl, updateData);
                    
                    if (result) {
                        await _web.lists.getByTitle(_GoalsListName).items.getById(updateItem.GoalId).update({
                            ReviewIDId: updateItem.ReviewID
                        });

                        console.log(`Goal ID: ${updateItem.GoalId} was updated successfully`);
                    }                  
                }
            }
            catch (error) {
                console.error(`Error updating GoalId ${updateItem.GoalId}:`, error);
            }
        }
    }    
}

function enableDisableSubmitButton() { 
    
    let buttons = fd.toolbar.buttons;
    let submitButton = buttons.find(button => button.text === 'Submit');    

    if (_WorkflowStatus === 'open' && !_isAllEmployeeCommentDate)
        disableEnableButton(submitButton, true);
    
    else if (_WorkflowStatus === 'open' && _isAllEmployeeCommentDate)
        disableEnableButton(submitButton, false);

    else if (_WorkflowStatus === 'sent to manager' && !_isAllManagerCommentDate)
        disableEnableButton(submitButton, true);
        
    else if (_WorkflowStatus === 'sent to manager' && _isAllManagerCommentDate)
        disableEnableButton(submitButton, false);
}

const handlePerfomanceReviewForm = async function () {   

    const svguserinfo = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
    const svgGoals = `<svg height="24" width="24" version="1.1" id="_x32_" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" 
        viewBox="0 0 512 512"  xml:space="preserve">
        <style type="text/css">
            .st0{fill:#000000;}
        </style>
        <g>
            <path class="st0" d="M228.31,77.506c21.358-1.961,37.07-20.901,35.093-42.259c-1.994-21.342-20.917-37.063-42.268-35.077
                c-21.358,1.986-37.054,20.893-35.076,42.259C188.045,63.778,206.968,79.498,228.31,77.506z"/>
            <path class="st0" d="M368.479,388.205c7.133,6.194,17.918,5.605,24.341-1.332l0.458-0.482c6.406-6.928,6.185-17.682-0.499-24.341
                l-45.045-50.028l-16.48-55.888c-0.752,1.34-1.536,2.663-2.402,3.922c-8.481,12.191-21.841,19.994-36.646,21.359l-6.822,0.474
                l17.387,41.997c4.118,6.848,9.224,13.032,15.14,18.409L368.479,388.205z"/>
            <path class="st0" d="M214.419,351.568c9.168-2.255,14.977-11.275,13.253-20.541l-11.128-59.76l73.455-5.099
                c10.532-0.744,20.19-6.21,26.228-14.904c6.03-8.677,7.794-19.609,4.804-29.766l-2.754-9.355l-22.968-83.668l39.374,2.132
                l31.474,30.322c-1.52,1.846-2.28,4.281-1.668,6.798l2.566,10.401l-18.294,4.51c-4.339,1.07-6.978,5.45-5.908,9.78l12.322,50.005
                c1.078,4.322,5.449,6.97,9.772,5.915l77.671-19.152c4.339-1.078,6.962-5.458,5.9-9.789l-12.338-49.988
                c-1.063-4.33-5.434-6.986-9.764-5.915l-18.302,4.51l-2.574-10.41c-1.046-4.232-5.311-6.798-9.543-5.76l-1.92,0.474
                c-0.434-1.773-1.218-3.489-2.395-5.025l-30.779-40.363c-3.325-4.347-8.048-7.419-13.384-8.662l-50.626-18.122
                c-18.131-6.504-38.328-3.333-53.6,8.4l-70.244,53.894l-45.56-16.431c-7.142-3.178-15.484-0.123-18.891,6.879l-0.474,0.956
                c-1.683,3.489-1.928,7.476-0.662,11.12c1.275,3.652,3.955,6.642,7.436,8.309l54.916,26.213c6.83,3.268,14.781,3.235,21.595-0.082
                l29.268-19.463l20.091,57.587l-48.591,4.249c-8.187,0.744-15.622,5.099-20.247,11.881c-4.641,6.79-5.982,15.296-3.693,23.189
                l24.152,82.606c2.68,9.143,12.084,14.568,21.342,12.289L214.419,351.568z M397.819,159.212l0.351,0.204l2.574,10.402l-26.318,6.488
                l-2.483-10.083c4.885,4.004,11.856,4.306,16.929,0.392l0.295-0.238c2.124-1.634,3.611-3.783,4.461-6.136L397.819,159.212z"/>
            <polygon class="st0" points="239.266,387.895 212.245,371.586 212.245,512 299.754,512 299.754,364.07 289.467,358.611 	"/>
            <polygon class="st0" points="316.488,512 403.988,512 403.988,419.377 316.488,372.951 	"/>
            <polygon class="st0" points="420.722,428.258 420.722,512 494.455,512 494.455,467.38 	"/>
            <polygon class="st0" points="108.011,512 195.512,512 195.512,361.479 108.011,308.647 	"/>
            <polygon class="st0" points="17.545,512 91.278,512 91.278,298.54 17.545,254.026 	"/>
        </g>
    </svg>`;
    const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;
    const signOff = `<svg width="24" height="24" viewBox="0 0 28 28" version="1.1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
                        <!-- Uploaded to: SVG Repo, www.svgrepo.com, Generator: SVG Repo Mixer Tools -->
                        <title>ic_fluent_signature_28_filled</title>
                        <desc>Created with Sketch.</desc>
                        <g id="🔍-Product-Icons" stroke="none" stroke-width="1" fill="none" fill-rule="evenodd">
                            <g id="ic_fluent_signature_28_filled" fill="#212121" fill-rule="nonzero">
                                <path d="M16.4798956,21.0019578 L16.75,21 C17.9702352,21 18.6112441,21.5058032 19.4020627,22.7041662 L19.7958278,23.3124409 C20.1028266,23.766938 20.2944374,23.9573247 20.535784,24.0567929 C20.9684873,24.2351266 21.3271008,24.1474446 22.6440782,23.5133213 L23.0473273,23.3170319 C23.8709982,22.9126711 24.4330286,22.6811606 25.0680983,22.5223931 C25.4699445,22.4219316 25.8771453,22.6662521 25.9776069,23.0680983 C26.0780684,23.4699445 25.8337479,23.8771453 25.4319017,23.9776069 C25.0371606,24.0762922 24.6589465,24.2178819 24.1641364,24.4458997 L23.0054899,25.0032673 C21.4376302,25.7436944 20.9059009,25.8317321 19.964216,25.4436275 C19.3391237,25.1860028 18.9836765,24.813298 18.4635639,24.0180227 L18.2688903,23.7140849 C17.6669841,22.7656437 17.3640608,22.5 16.75,22.5 L16.5912946,22.5037584 C16.1581568,22.5299816 15.8777212,22.7284469 14.009281,24.1150241 C12.2670395,25.4079488 10.9383359,26.0254984 9.24864243,26.0254984 C7.18872869,26.0254984 5.24773367,25.647067 3.43145875,24.8905363 L6.31377803,24.2241784 C7.25769404,24.4250762 8.23567143,24.5254984 9.24864243,24.5254984 C10.5393035,24.5254984 11.609129,24.0282691 13.1153796,22.9104743 L14.275444,22.0545488 C15.5468065,21.1304903 15.8296113,21.016032 16.4798956,21.0019578 L16.4798956,21.0019578 Z M22.7770988,3.22208979 C24.4507223,4.8957133 24.4507566,7.60916079 22.7771889,9.28281324 L21.741655,10.3184475 C22.8936263,11.7199657 22.8521526,13.2053774 21.7811031,14.279556 L18.7800727,17.2805874 L18.7800727,17.2805874 C18.4870374,17.5733384 18.0121637,17.573108 17.7194126,17.2800727 C17.4266616,16.9870374 17.426892,16.5121637 17.7199273,16.2194126 L20.7188969,13.220444 C21.2039571,12.7339668 21.2600021,12.1299983 20.678941,11.3818945 L10.0845437,21.9761011 C9.78635459,22.2743053 9.41036117,22.482705 8.99944703,22.5775313 L2.91864463,23.9807934 C2.37859061,24.1054212 1.89457875,23.6214094 2.0192066,23.0813554 L3.42247794,17.0005129 C3.51729557,16.5896365 3.72566589,16.2136736 4.0238276,15.9154968 L16.7165019,3.22217992 C18.3900415,1.54855555 21.1034349,1.54851059 22.7770988,3.22208979 Z" id="🎨-Color">

                    </path>
                            </g>
                        </g>
                    </svg>`;
    setIconSource("overview-icon", svguserinfo);
    setIconSource("goals-icon", svgGoals);
    setIconSource("signOff-icon", signOff);
    setIconSource("attachment-icon", svgattachment);

    formatingButtonsBar('Human Resources: Performance Management Review');

    $('.toHide').hide();

    let EmployeevalueDisp = _PerfomanceReviewFormFields.Employee.value.DisplayText;
    let ManagervalueDisp = _PerfomanceReviewFormFields.Manager.value.DisplayText;

    if(_isDisplay){
        EmployeevalueDisp = _PerfomanceReviewFormFields.Employee.value.displayName;
        ManagervalueDisp = _PerfomanceReviewFormFields.Manager.value.displayName;
    }   

    // <tr>
    //     <th>Discussion Approval Status</th>
    //     <td>${_PerfomanceReviewFormFields.DiscussionApprovalStatus?.value || 'N/A'}</td>
    //     <th>Finished Date</th>
    //     <td>${_PerfomanceReviewFormFields.FinishedDate?.value ? formatDate(_PerfomanceReviewFormFields.FinishedDate.value) : 'N/A'}</td>
    // </tr> 

    var content = `<table role="grid" class="modern-table">
                <tr>
                    <th width="10%">Employee</th>
                    <td width="40%">${EmployeevalueDisp || ''}</td>
                    <th width="10%">Manager</th>
                    <td width="40%">${ManagervalueDisp || ''}</td>
                </tr>
                <tr>
                    <th>Description</th>
                    <td>${_PerfomanceReviewFormFields.Description.value || ''}</td>
                    <th>Performance Period</th>
                    <td>${_PerfomanceReviewFormFields.PerformancePeriodId.value || ''}</td>
                </tr>                            
                <tr>
                    <th>Start Date</th>
                    <td>${_PerfomanceReviewFormFields.StartDate?.value ? formatDate(_PerfomanceReviewFormFields.StartDate.value) : 'N/A'}</td>
                    <th>End Date</th>
                    <td>${_PerfomanceReviewFormFields.EndDate?.value ? formatDate(_PerfomanceReviewFormFields.EndDate.value) : 'N/A'}</td>
                </tr>                
            </table><br>`;

    $('#TopTable').append(content);   
    
    let dt = _PerfomanceReviewFormFields.goalsDT;
    
    let formatedStartDate = formatDateforSPQuery(_PerfomanceReviewFormFields.StartDate.value);
    let formatedEndDate = formatDateforSPQuery(_PerfomanceReviewFormFields.EndDate.value);  

    let EmployeeName = _PerfomanceReviewFormFields.Employee.value.DisplayText;
    if (_isDisplay)
        EmployeeName = _PerfomanceReviewFormFields.Employee.value.displayName;  

    dt.ready().then(function () {
        
        dt.buttons[0].visible = false;   
        dt.buttons[dt.buttons.length - 1].visible = false;
        dt.filter = `<And>
                        <And>
                            <And>
                                <Eq><FieldRef Name="Author"/><Value Type="User">${EmployeeName}</Value></Eq>
                                <Geq><FieldRef Name="StartDate"/><Value Type="DateTime">${formatedStartDate}</Value></Geq>
                            </And>
                            <Eq><FieldRef Name="ApprovalStatus"/><Value Type="Choice">Approved</Value></Eq>
                        </And>
                        <Leq><FieldRef Name="EndDate"/><Value Type="DateTime">${formatedEndDate}</Value></Leq>
                    </And>`;
        dt.refresh().then(function () { 
            setTimeout(async function () {
                if (_isDisplay) { 
                    fixTableColumnsWidth();
                    disableDataTableCheckBoxes(dt);
                }
                else {

                    fixTableColumnsWidth();                    

                    (async () => {                         
                        await linkdtRecords(dt);
                    })();
                }    
            }, 250);     
        });          
    }); 

    dt.$watch('selectedItems', function(items, pervItem) {

        if (items.length > 1) {

            var itemtoUncheck = pervItem[0];

            var row = document.querySelector('tr[data-uid="' + itemtoUncheck.uid + '"]').children[0];

            var checkbox = row.querySelector('input.k-checkbox');

            if (checkbox){
                checkbox.checked = false;
                row.parentElement.classList.remove('k-state-selected');
            }

            var indexOfPrevItem = 0;

            items.forEach((item, index) => {
                if(item.ID === itemtoUncheck.ID){indexOfPrevItem = index;}
            });

            items.splice(indexOfPrevItem, 1);
        }
    });

    dt.$on('change', async function(changeData) {
        dt.ready().then(async function () {
             setTimeout(function() {
                 fixTableColumnsWidth();                 
            }, 250);            
    	});
    });     

    dt.dialogOptions = {
        width: '75%',
        height: '90%'      
    }    
}
//#endregion

