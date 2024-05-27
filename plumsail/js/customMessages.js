var htLicenseKey = '4afed-93e31-596f2-64130-b1031';
var _RLOD = 'RLOD', _SLOD = 'SLOD', _MIDP = 'MIDP', _AllowedExtensions = 'AllowedExtensions', _MajorTypes = 'MajorTypes', 
    _InitiateSubmittals = 'InitiateSubmittals', _WorkflowSteps = 'WorkflowSteps', _DesignTasks = 'DesignTasks',
    _CDS = 'CDS', _UploadDeliverables = 'Upload Deliverables', _Deliverables = 'Deliverables', _Trades = 'Trades',
    _LeadTask = 'LeadTasks', _PartTask = 'PartTasks';

//#region EFORMS MAIN MODULE HEADER TITLE
var leadTaskHeaderTitle = 'CM Task',
    partTaskHeaderTitle = 'Edit Part Task',
    scrHeaderTitle = 'Site Clarification Request',
    irHeaderTitle = 'Inspection Request',
    mirHeaderTitle = 'Material Inspection Request',
    matHeaderTitle = 'Material Submittal',
    siHeaderTitle = 'Site Instruction',
    slfHeaderTitle = 'Snag List',
    lodHeaderTitle = 'Tentative LOD',
    dtrdHeaderTitle = 'Digital Ready Design (DRD) Guidelines',
    dprHeaderTitle = 'Daily Progress Report',
    fncHeaderTitle = 'Filenaming Convention',
    ckdHeaderTitle = 'Check Deliverables',
    insHeaderTitle = 'Initiate New Submission';
//#endregion

//#region EFORMS BUTTON MESSAGES
var submit_SLF_TeamLeader_ValidationMesg = 'Inspector Tasks must be completed before you submit to RE.'
var reject_SLF_TeamLeader_ValidationMesg = 'Reason of rejection is required field';
var rejectButton_SLF_TeamLeader_ValidationMesg = 'One Trade at least should selected for rejection';
//#endregion

//#region EFORMS ERROR MESSAGES (GENERAL)
var isAllowedUserMesg = 'You are not allowed to review this task as you are not part of the assignees';
var InspectorLabelMesg = "<div class='instruction'><u>Inspectors:</u>" +
                         "<h6 class='redColor'> (To return the task to inspector, select the completed task in the tick column then click Reject)</h6>";

var leadTradeRequiredFieldMesg = 'Lead Trade is Required to be filled';
var partTradeRequiredFieldMesg = 'Part Trade is Required to be filled';
var partError = "Can not assign same trade as lead and part.";
//#endregion

//#region EFORMS ERROR MESSAGES (SLF)
var SLFLabelMesg = '<u>Snag List Items:</u>';
var htPreventDeleteMesg = 'can not remove an item that were previously inserted before.';

var htRequiredFieldMesg = 'field is required';
var htdropdownMesg = 'value out of range';
var htremoveValueMesg = 'field value should be removed'


var contractorLabelMesg = 'Kindly fill Inspection date for any completed Items.';
var contractorRequiredFieldMesg = 'One Inspection Date at least should be set for an item';
//#endregion

//#region GENERAL BUTTON MESSAGES
var saveMesg = "Click 'Save' to store your progress and keep your work as a draft.";
var submitMesg = "Click 'Submit' to finalize and send officially.";
var cancelMesg = "Click 'Cancel' to discard changes and exit without saving.";
var assignMesg = "Click 'Assign' to allocate this task or responsibility.";
var rejectMesg = "To return the task to inspector, select the completed task in the tick column then click Reject."
var previewMesg = "Click here to preview the document before submitting to contractor";
var assignMesg = 'Select Lead and Part(optional) then click Assign'
//#endregion

//#region FNC MESSAGES
var fncLengthMesg = 'filename length is not matching',
fncSuccessMesg = 'filename matches naming convention';
//#endregion

//#region AUDIT REPORT MESSAGES
var mapHeaderTitle = 'Master Plan',
    ausHeaderTitle = 'Audit Schedule',
    aurHeaderTitle = 'Audit Report',
    dccHeaderTitle = 'Drawing Control Checklist',
    genSummary = 'Fill Department and start/end Audi date to generate summary';
//#endregion

//#region DIGITAL TWIN READY DESIGN (DTRD)
var dtrdPermission= 'Apologies, but this task has not been assigned to you.'
var dtrdTradeValidation = 'Items is already submitted';
//#endregion

//#region DELIVERABLES CHECKPAGE ERROR MESSAGES
var spaceFileName = 'Space character is not allowed';
var extFileName = 'No Extension for filename';
//#endregion



