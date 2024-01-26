var htLicenseKey = '4afed-93e31-596f2-64130-b1031';

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
    dtrdHeaderTitle = 'Digital Twin Ready Design Guidelines',
    dprHeaderTitle = 'Daily Progress Report';
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
//#endregion

//#region FNC MESSAGES
fncLengthMesg = 'filename length is not matching';
fncSuccessMesg = 'filename matches naming convention';
//#endregion

//#region AUDIT REPORT MESSAGES
aurHeaderTitle = 'Audit Report';
dccHeaderTitle = 'Drawing Control Checklist';
genSummary = 'Fill Department and start/end Audi date to generate summary';
//#endregion