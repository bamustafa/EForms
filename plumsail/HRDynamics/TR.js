var _module, _formType, _web, _webUrl, _siteUrl, _layout, _itemId, _ImageUrl, _ListInternalName, _ProjectNumber, 
    _ListFullUrl, _CurrentUser, _fields = [];
let Inputelems = document.querySelectorAll('input[type="text"]');
let _Email, _Notification = '', _DipN = "", _htLibraryUrl, _errorImg, _submitImg;

//handson variables
let _hot, _container, _data = [], _colArray, batchSize = 15;
const greenColor = '#5FC9B3', redColor = '#F28B82', yellowColor = '#6c757d';
let _formFields = {};

const disableField = (field) => field.disabled = true;
const enabledField = (field) => field.disabled = false;
const requiredField = (field) => field.required = true;

//form type
let _isNew = false, _isEdit = false, _isDisplay = false; _proceed = false;

var onRender = async function (relativeLayoutPath, moduleName, formType){ 
       
    try {      
        
        const startTime = performance.now();

        _layout = relativeLayoutPath;

        await loadScripts().then(async () => {

            showPreloader();
            await extractValues(moduleName, formType);          

            if (moduleName == 'TR') {
                if (formType === "New")                    
                    await TR_newForm();
            }
            else if (moduleName == "Training") {
                
                const TrainingEvaluationsvg = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
                setIconSource("TrainingEvaluation-icon", TrainingEvaluationsvg);
                const Contenticonsvg = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
                setIconSource("Content-icon", Contenticonsvg);
                const TrainingCommentsvg = `<svg fill="#000000" width="24" height="24" viewBox="0 0 24 24" id="create-note-alt" data-name="Line Color" xmlns="http://www.w3.org/2000/svg" class="icon line-color"><path id="secondary" d="M19.44,8.22C17.53,10.41,14,10,14,10s-.39-4,1.53-6.18a3.49,3.49,0,0,1,.56-.53L18,4l.47-1.82A8.19,8.19,0,0,1,21,2S21.36,6,19.44,8.22ZM14,10l-2,2" style="fill: none; stroke: rgb(44, 169, 188); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path><path id="primary" d="M12,3H4A1,1,0,0,0,3,4V20a1,1,0,0,0,1,1H20a1,1,0,0,0,1-1V12" style="fill: none; stroke: rgb(0, 0, 0); stroke-linecap: round; stroke-linejoin: round; stroke-width: 2;"></path></svg>`;
                setIconSource("Comments-icon", TrainingCommentsvg);

                _formFields = {
                    Title: fd.field('Title'),
                    TypeOfTraining: fd.field('TypeOfTraining'),
                    Instructor: fd.field('Instructor'),
                    DurationFrom: fd.field('DurationFrom'),
                    Location: fd.field('Location'),
                    
                    OrganizationRate: fd.field('OrganizationRate'),
                    TrainerSkillRate: fd.field('TrainerSkillRate'),
                    DurationRate: fd.field('DurationRate'),
                    TechnicalLevel: fd.field('TechnicalLevel'),
                    InteractivityRate: fd.field('InteractivityRate'),
                    MaterialUsefulness: fd.field('MaterialUsefulness'),
                    TechnicalFunctionalityRate: fd.field('TechnicalFunctionalityRate'),
                    TrainingExpectations: fd.field('TrainingExpectations'),
                    RecommendationToColleagues: fd.field('RecommendationToColleagues'),
                    OverallImpression: fd.field('OverallImpression'),

                    OverallImpressionComment: fd.field('OverallImpressionComment'),
                    TrainingBenefitInWork: fd.field('TrainingBenefitInWork'),
                    TrainingProblems: fd.field('TrainingProblems'),
                    SuggestionsToImprove: fd.field('SuggestionsToImprove'),
                    MissingTopics: fd.field('MissingTopics')
                }

                if(formType === "New")
                    await Training_newForm();
                else if(formType=="Edit")
                    await Training_editForm();
                else if(formType=="Display")
                    await Training_displayForm();
            }            		
            
            const endTime = performance.now();
            const elapsedTime = endTime - startTime;
            console.log(`Execution time HRDynamics: ${elapsedTime} milliseconds`);            
        })
	}
	catch (e) {
		//alert(e);
		console.log(e);	
	}
    finally{      
        hidePreloader();
    }
}

var TR_newForm = async function(){ 
    
    try {

        _fields = {
            Name: fd.field("Name"),
            IDNo: fd.field("IDNo"),
            Department: fd.field("Department"),
            Position: fd.field("Position"),

            TypeofTraining: fd.field("TypeofTraining"),
            Title: fd.field("Title"),

            From: fd.field("From"),
            To: fd.field("To"),

            VendorName: fd.field("VendorName"),
            Tel: fd.field("Tel"),
            Address: fd.field("Address"),
            Email: fd.field("E_x002d_mail"),
            Fax: fd.field("Fax"),


            location: fd.field("location"),
            Cost: fd.field("Cost"),
            chargedTo: fd.field("chargedTo"),
            AdditionalInformation: fd.field("AdditionalInformation"),
            Justifications: fd.field("Justifications"),
            Status: fd.field("Status"),
            Employee: fd.field("Employee")
        }
        
        const svguserinfo = `<svg width="24" height="24" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" aria-hidden="true" role="img" class="iconify iconify--emojione" preserveAspectRatio="xMidYMid meet">
                                <circle cx="32" cy="32" r="30" fill="#4fd1d9">
                                </circle>
                                <g fill="#ffffff">
                                <path d="M27 27.8h10v24H27z">
                                </path>
                                <circle cx="32" cy="17.2" r="5">
                                </circle>
                                </g>
                            </svg>`;

        const svgtraining = `<svg width="24" height="24" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
                                <g data-name="21_Certificate" id="_21_Certificate">
                                <path d="M61,50H3a1,1,0,0,1-1-1V5A1,1,0,0,1,3,4H61a1,1,0,0,1,1,1V49A1,1,0,0,1,61,50Z" style="fill:#febc00"/>
                                <path d="M25.583,48l-5.767,5.675a1.051,1.051,0,0,0-.269,1.061,1.075,1.075,0,0,0,.824.734l3.1.6.634,3.077a1.067,1.067,0,0,0,.747.806,1.089,1.089,0,0,0,1.075-.265L32.5,53.218A9.72,9.72,0,0,1,25.583,48Z" style="fill:#f74e0c"/>
                                <path d="M38.417,48l5.767,5.675a1.051,1.051,0,0,1,.269,1.061,1.075,1.075,0,0,1-.824.734l-3.1.6L39.9,59.148a1.067,1.067,0,0,1-.747.806,1.089,1.089,0,0,1-1.075-.265L31.5,53.218A9.72,9.72,0,0,0,38.417,48Z" style="fill:#f74e0c"/>
                                <path d="M57,46H7a1,1,0,0,1-1-1V9A1,1,0,0,1,7,8H57a1,1,0,0,1,1,1V45A1,1,0,0,1,57,46Z" style="fill:#fdfeff"/>
                                <path d="M21.051,46a10.9,10.9,0,0,0,1.163,4H41.786a10.9,10.9,0,0,0,1.163-4Z" style="fill:#edaa03"/>
                                <path d="M43,45a11,11,0,0,0-22,0c0,.338.021.67.051,1h21.9C42.979,45.67,43,45.338,43,45Z" style="fill:#dfeaef"/>
                                <path d="M8,45V9A1,1,0,0,1,9,8H7A1,1,0,0,0,6,9V45a1,1,0,0,0,1,1H9A1,1,0,0,1,8,45Z" style="fill:#dfeaef"/>
                                <path d="M32,54a9,9,0,1,1,9-9A9.01,9.01,0,0,1,32,54Z" style="fill:#f74e0c"/>
                                <path d="M34,52a8.984,8.984,0,0,1-7.276-14.276A8.99,8.99,0,1,0,39.276,50.276,8.944,8.944,0,0,1,34,52Z" style="fill:#cc2600"/>
                                <circle cx="32" cy="45" r="4" style="fill:#febc00"/>
                                <path d="M34,47a4,4,0,0,1-4-4,3.964,3.964,0,0,1,.36-1.64,3.995,3.995,0,1,0,5.28,5.28A3.964,3.964,0,0,1,34,47Z" style="fill:#edaa03"/>
                                <rect height="2" style="fill:#dfeaef" width="28" x="18" y="16"/>
                                <rect height="2" style="fill:#dfeaef" width="28" x="18" y="20"/>
                                <rect height="2" style="fill:#dfeaef" width="28" x="18" y="24"/>
                                <rect height="2" style="fill:#dfeaef" width="14" x="32" y="28"/>
                                <rect height="2" style="fill:#dfeaef" width="8" x="38" y="32"/>
                                </g>
                            </svg>`;
        
        const svgvendor = `<svg height="24" width="24" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" 
                                    viewBox="0 0 512.016 512.016" xml:space="preserve">
                                <circle cx="377.128" cy="450.56" r="61.456"/>
                                <circle style="fill:#999999;" cx="377.128" cy="450.56" r="29.856"/>
                                <polygon points="145.544,151.36 105.184,31.92 0.68,31.92 0.68,0 128.096,0 175.784,141.136 "/>
                                <circle cx="191.752" cy="450.56" r="61.456"/>
                                <polygon style="fill:#FFFFFF;" points="147.824,353.584 78.992,132.552 499.992,132.552 421.216,353.584 "/>
                                <path style="fill:#E21B1B;" d="M488.648,140.56L415.576,345.6H153.704l-63.84-205.04L488.648,140.56 M511.336,124.56H68.128
                                    l73.808,237.04h284.92L511.336,124.56z"/>
                                <circle style="fill:#999999;" cx="191.752" cy="450.56" r="29.856"/>
                                <circle style="fill:#E21B1B;" cx="315.664" cy="151.36" r="108"/>
                                <polygon style="fill:#FFFFFF;" points="352.784,102.312 297.84,170.072 274.752,147.28 259.52,164.056 299.624,203.632 
                                    369.696,117.216 "/>
                            </svg>`; 

        const svgdetails = `<svg fill="#000000" height="24" width="24" version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" 
                                    viewBox="0 0 500 500" xml:space="preserve">
                                <g>
                                    <g>
                                        <path d="M406.192,109.332H97.716c-2.212,0-4,1.792-4,4c0,2.208,1.788,4,4,4h308.476c2.212,0,4-1.792,4-4
                                            C410.192,111.124,408.404,109.332,406.192,109.332z"/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M160.192,148.38H97.716c-2.212,0-4,1.792-4,4s1.788,4,4,4h62.476c2.212,0,4-1.792,4-4S162.404,148.38,160.192,148.38z"/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M406.192,148.38h-214.76c-2.212,0-4,1.792-4,4s1.788,4,4,4h214.76c2.212,0,4-1.792,4-4S408.404,148.38,406.192,148.38z"/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M406.192,187.428h-62.476c-2.212,0-4,1.792-4,4s1.788,4,4,4h62.476c2.212,0,4-1.792,4-4S408.404,187.428,406.192,187.428z
                                            "/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M308.572,187.428H97.716c-2.212,0-4,1.792-4,4s1.788,4,4,4h210.856c2.208,0,4-1.792,4-4S310.784,187.428,308.572,187.428z
                                            "/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M406.192,226.476H97.716c-2.212,0-4,1.792-4,4c0,2.208,1.788,4,4,4h308.476c2.212,0,4-1.792,4-4
                                            C410.192,228.268,408.404,226.476,406.192,226.476z"/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M406.192,343.62H97.716c-2.212,0-4,1.792-4,4s1.788,4,4,4h308.476c2.212,0,4-1.792,4-4S408.404,343.62,406.192,343.62z"/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M160.192,265.524H97.716c-2.212,0-4,1.792-4,4s1.788,4,4,4h62.476c2.212,0,4-1.792,4-4S162.404,265.524,160.192,265.524z"
                                            />
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M406.192,265.524h-214.76c-2.212,0-4,1.792-4,4s1.788,4,4,4h214.76c2.212,0,4-1.792,4-4S408.404,265.524,406.192,265.524z
                                            "/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M406.192,304.572h-62.476c-2.212,0-4,1.792-4,4s1.788,4,4,4h62.476c2.212,0,4-1.792,4-4S408.404,304.572,406.192,304.572z
                                            "/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M308.572,304.572H97.716c-2.212,0-4,1.792-4,4s1.788,4,4,4h210.856c2.208,0,4-1.792,4-4S310.784,304.572,308.572,304.572z
                                            "/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M160.192,382.668H97.716c-2.212,0-4,1.792-4,4s1.788,4,4,4h62.476c2.212,0,4-1.792,4-4S162.404,382.668,160.192,382.668z"
                                            />
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M367.144,382.668H191.428c-2.212,0-4,1.792-4,4s1.788,4,4,4h175.716c2.212,0,4-1.792,4-4S369.356,382.668,367.144,382.668
                                            z"/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M466.66,0H37.144C17.356,0,0,16.436,0,35.176v429.516C0,483.5,17.356,500,37.144,500h331.908
                                            c1.068,0,2.096-0.428,2.852-1.192L498.852,369.9c0.732-0.748,1.148-1.756,1.148-2.808V35.176C500,15.124,485.668,0,466.66,0z
                                            M492,365.456L367.376,492H37.144C21.616,492,8,479.24,8,464.692V35.176C8,20.7,21.616,8,37.144,8H466.66
                                            C481.104,8,492,19.68,492,35.176V365.456z"/>
                                    </g>
                                </g>
                                <g>
                                    <g>
                                        <path d="M445.236,363.144h-44.948c-19.788,0-37.144,16.436-37.144,35.176V496c0,2.208,1.788,4,4,4c2.212,0,4-1.792,4-4v-97.68
                                            c0-14.476,13.616-27.176,29.144-27.176h44.948c2.212,0,4-1.792,4-4S447.448,363.144,445.236,363.144z"/>
                                    </g>
                                </g>
                            </svg>`;

        const svgattachment = `<svg width="24" height="24" viewBox="0 0 32 32" data-name="Layer 1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><defs><style>.cls-1{fill:url(#linear-gradient);}.cls-2{fill:url(#linear-gradient-2);}.cls-3{fill:#f8edeb;}.cls-4{fill:#577590;}</style><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient" x1="4" x2="25" y1="16" y2="16"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.01" stop-color="#f8dbc2"/><stop offset="0.12" stop-color="#f3ceb1"/><stop offset="0.26" stop-color="#f0c5a6"/><stop offset="0.46" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient><linearGradient gradientUnits="userSpaceOnUse" id="linear-gradient-2" x1="7.48" x2="27.52" y1="8.98" y2="29.02"><stop offset="0" stop-color="#f9dcc4"/><stop offset="0.32" stop-color="#f8d9c0"/><stop offset="0.64" stop-color="#f4cfb3"/><stop offset="0.98" stop-color="#eebf9f"/><stop offset="1" stop-color="#edbe9d"/></linearGradient></defs><rect class="cls-1" height="22" rx="2.5" width="21" x="4" y="5"/><rect class="cls-2" height="22" rx="2.5" width="21" x="7" y="8"/><path class="cls-3" d="M20,2a5,5,0,0,0-5,5v5a2,2,0,0,0,3,1.73V15.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/><path class="cls-4" d="M20,2a5,5,0,0,0-5,5v4.5a.5.5,0,0,0,1,0V7a4,4,0,0,1,8,0v8.5a2.5,2.5,0,0,1-5,0V7a1,1,0,0,1,2,0v8.5a.5.5,0,0,0,1,0V7a2,2,0,0,0-4,0v8.5a3.5,3.5,0,0,0,7,0V7A5,5,0,0,0,20,2Z"/></svg>`;

        setIconSource("overview-icon", svguserinfo);
        setIconSource("training-icon", svgtraining);
        setIconSource("vendor-icon", svgvendor);
        setIconSource("details-icon", svgdetails);
        setIconSource("attachment-icon", svgattachment);

        await setButtonActions("Accept", "Submit", `${greenColor}`, 'white');
        await setButtonActions("ChromeClose", "Cancel", `${yellowColor}`, 'white');

        formatingButtonsBar('Human Resources: Employee Training Application'); 
        
        setPSHeaderMessage('');
        setPSErrorMesg(`Please complete the following form and submit.`);

        const TRAdmin = await CheckifUserinSPGroup();
        _fields.Employee.value = _CurrentUser.Title; 

        if(TRAdmin === "User") { 
            fd.control('TrainBulkTable').hidden = true;
            document.querySelector('.description-area').style.display = 'none';           
            let LoginName = _CurrentUser.LoginName.split('|')[1];
            await GetEmployeesInfo(LoginName); 
        }  
        else  {

            _HideFormFields([_fields.Name, _fields.IDNo, _fields.Department, _fields.Position], true);

            fd.control('TrainBulkTable').$on('change', async function(changeData) {  

                if (changeData.type === 'add' ) {                
                    await HandleCostValue(changeData.itemId);                     
                }
                else if (changeData.type === 'edit') {
                    await HandleCostValue();                 
                }
                else if (changeData.type === 'delete') {
                    for (const itemId of changeData.itemIds) {
                        await HandleCostValue(itemId, true);
                    }                
                }                
            });         

            fd.control('TrainBulkTable').readonlyRow = function(row) {           
                //return row.Cost !== '';
            };

            fd.control('TrainBulkTable').ready().then(async function(){                       
                fd.control('TrainBulkTable').refresh().then(async function() {
                    setTimeout(async function() {
                        FixWidget(fd.control('TrainBulkTable')); 
                        await HandleCostValue();                        
                    }, 200); // Delay of 2000 milliseconds (2 seconds)
                });                       
            }); 

            fd.control('TrainBulkTable').$on('edit', function(editData) {                  

                const cells = document.querySelectorAll('td[data-container-for="Department"], td[data-container-for="IDNo"], td[data-container-for="Position"]');
                cells.forEach(cell => {
                    const inputElement = cell.querySelector('input');
                    if (inputElement) {
                        inputElement.disabled = true;
                    }
                });                     
                
                $('.k-button.k-button-icontext.k-grid-cancel').on('click', function(event) {
                    setTimeout(async function() {
                        FixWidget(fd.control('TrainBulkTable'));                                               
                    }, 100);
                });
              
            });            
        }

        _HideFormFields([_fields.Employee, _fields.Status], true);
    }    
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);        
    }      
}

var Training_newForm = async function(){
    try
    {        
        formatingButtonsBar("Evaluation of Training Event Form"); 

        setPSHeaderMessage('');
        setPSErrorMesg(`Please complete the following form and submit to the technical training manager.`); 
        
        await setButtonActions("Accept", "Submit", `${greenColor}`, 'white');
        await setButtonActions("ChromeClose", "Cancel", `${yellowColor}`, 'white'); 
        
        Object.values(_formFields).forEach(requiredField);
    }
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);
        showPreloader();
    }
}
var Training_editForm= async function(){
    try
    {
        formatingButtonsBar("Evaluation of Training Event Form");
    }
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);
        showPreloader();
    }
}
var Training_displayForm = async function(){
    try
    {
        formatingButtonsBar("Evaluation of Training Event Form");
    }
    catch(err){
        console.log(err.message, err.stack)
        await _generateErrorEmail(_spPageContextInfo.siteAbsoluteUrl, '', '', err.message, err.stack);
        showPreloader();
    }
}

//#region General Functions

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

    _CurrentUser = await pnp.sp.web.currentUser.get();
    //  const elapsedTime = endTime - startTime;
    //  console.log(`extractValues: ${elapsedTime} milliseconds`);
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

    const formContainer = document.querySelector('.fd-form-container.container-fluid');
    if (formContainer) {
        formContainer.style.setProperty('margin-top', '-8px', 'important');
        formContainer.style.setProperty('padding-left', '5px', 'important');
    }

    const iconPath = _spPageContextInfo.webAbsoluteUrl + '/_layouts/15/Images/animdarlogo1.png';
    const linkElement = `<a href="${_spPageContextInfo.webAbsoluteUrl}" style="text-decoration: none; color: inherit; display: flex; align-items: center; font-size: 18px;">
                            <img src="${iconPath}" alt="Icon" style="width: 50px; height: 26px; margin-right: 14px;">${titelValue}</a>`;
    $('span.o365cs-nav-brandingText').html(linkElement);

    $('.o365cs-base.o365cs-topnavBGColor-2').css('background', 'linear-gradient(to bottom, #808080, #4d4d4d, #1a1a1a, #000000, #1a1a1a, #4d4d4d, #808080)');
  
    let co = 0;
    $('.fd-field-choice').each(function () { 

        if (co === 0) {

            const $choice = $(this); // reference to the first fd-field-choice

            $choice.find('.fd-fill-in-choice').css({
                display: 'flex',
                'margin-top': '0px'
            });

            $choice.find('.fd-fill-in-choice .row').each(function () {
                this.style.setProperty('width', '200px', 'important');
            });

            $choice.css('display', 'flex');
            $choice.find('.col-sm.form-check.d-none.d-sm-block').attr("style", "display: none !important;");
            $choice.find('.col-sm.form-check:not(.d-none)').css('width', '500px');
        }

        co++;
    });

    $('.fd-grid.container-fluid').attr("style", "margin-top: 5px !important; padding: 10px !important;");
  
    $('.divGrid').each(function() {
        $(this).css({             
            'margin-left': '0px', /* Align with the content */            
        });
    }); 

    if (_module === 'TR') {        

        $('.divGrid').each(function() {
            $(this).css({             
                'margin-left': '5px', /* Align with the content */            
            });
        }); 
        $('.border-title').each(function() {
            $(this).css({          
                'margin-top': '-22px', /* Adjust the position to sit on the border */                  
            });
        });
    }    
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

function setIconSource(elementId, iconFileName) {

  const iconElement = document.getElementById(elementId);

  if (iconElement) {
      iconElement.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconFileName)}`;
  }
}

const setButtonActions = async function (icon, text, bgColor, color) {
    
    if (!_isDisplay)
        fd.toolbar.buttons[0].style = "display: none;";

    fd.toolbar.buttons[1].style = "display: none;"; 

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
            else if (text == "Submit" ) {               
                if (fd.isValid) {
                    showPreloader();                                       
                    console.log('new child form')
                    fd.save();                    
                }
            }
        }
    });
}
//#endregion