let dept, deptId;

var onTradeContReviewRender = async function(isDisable){

    let fullProjTitle = localStorage.getItem('FullProjTitle');
    fullProjTitle += ' - Departmental Project Initiation Form (DPIR)'
    setPageStyle(fullProjTitle);

    deptId = fd.itemId;
    dept = $('#dpirHeader span.PageTitle').text().replace('Department Initiation','').trim();
    var result = await getTradeTemplateItems(Questions, ['Id', 'Title', 'IsRequired']);
    await renderTradeCustomTable(result, isDisable);
}

var getTradeTemplateItems = async function(targetList, colsInternal){
    var itemArray = [];
    
    const list = _web.lists.getByTitle(TradeContractReview);
    let items = await list.items
              //  .select("Id,MasterID/Id,Question/Id,Question/Title,Question/IsRequired,Answer,Comments,NeedPMAction,PMResponse,PMStatus")
              //  .expand("MasterID,Question")
              //  .filter(`MasterID/Id eq ${_itemId} and Department eq '${TextQueryEncode(dept)}'`).get();

              .select("Id,RDID/Id,Question/Id,Question/Title,Question/IsRequired,Answer,Comments,NeedPMAction,PMResponse,PMStatus")
               .expand("RDID,Question")
               .filter(`RDID/Id eq ${deptId} and Department eq '${TextQueryEncode(dept)}'`).get();

    //let items = await list.items.get();
    
    if(items.length > 0){
      for(var i = 0; i < items.length; i++){
        var item = items[i];
        var rowData  = {};
  
        rowData["Id"] = item["Id"];
        rowData["Title"] = item.Question.Title;
        rowData["TitleId"] = item.Question.Id;
        rowData["IsRequired"] = item.Question.IsRequired;
        rowData["Answer"] = item["Answer"];
        rowData["Comments"] = item["Comments"];

        rowData["NeedPMAction"] = item["NeedPMAction"];
        rowData["PMResponse"] = item["PMResponse"];
        rowData["PMStatus"] = item["PMStatus"];
        itemArray.push(rowData);
      }
    }
  
    else{
      const cols = colsInternal.join(',');
      var _query = "Category eq 'DPIR'";
      items = await pnp.sp.web.lists.getByTitle(targetList).items.filter(_query).select(cols).getAll();
  
      for(var i = 0; i < items.length; i++){
        var item = items[i];
        var rowData  = {};
  
        for(var j = 0; j < colsInternal.length; j++){
          var colname = colsInternal[j];
          rowData[colname] = item[colname];
        }
        itemArray.push(rowData);
      }
    }
    return itemArray;
}
  
let renderTradeCustomTable = async function(result, isDisable){
    let resultCount = result !== null ? result.length : 0;
    
    let tableId = 'tbltradeQId'
    let divElement = $('');
    if($(`#${tableId}`).length === 0){
      let tbl = `</br><table id='${tableId}' width='100%' class='modern-table'> 
                    <tr>
                      <th width='30%'><label class="d-flex fd-field-title col-form-label">Question</label></th>
                      <th width='15%'><label class="d-flex fd-field-title col-form-label">Answer</label></th>
                      <th width='15%'><label class="d-flex fd-field-title col-form-label">Comments</label></th>

                      <th width='10%'><label class="d-flex fd-field-title col-form-label">Action Requested by PM</label></th>
                      <th width='20%'><label class="d-flex fd-field-title col-form-label">Response by PM</label></th>
                      <th width='10%'><label class="d-flex fd-field-title col-form-label">Status</label></th>
                    </tr>`;
  
      for(var i = 0; i < resultCount; i++){

        let {Id, Title, TitleId, IsRequired, Answer, Comments, NeedPMAction, PMResponse, PMStatus} = result[i];
        TitleId = TitleId === undefined ? Id : TitleId;
        tbl += '<tr>'
        tbl += await renderTradeQuestionControls(i, Title, TitleId, Answer, IsRequired.toLowerCase(), Comments, NeedPMAction, PMResponse, PMStatus, isDisable);
        tbl += '</tr>'
      }


      let imgUrlLegend = `${_webUrl}${_layout}/Images/legend.png`
      let imgUrlNotSettled = `${_webUrl}${_layout}/Images/redclock.png`
      let imgUrlSettled = `${_webUrl}${_layout}/Images/Settle.png`
      
      let legendTbl = `<table style="border: 0px solid black; margin: 10px 0px 0px 10px" cellpadding="0" cellspacing="0">
                          <tr>
                              <td>
                                  <fieldset style="padding:1px !important; border: 1px solid #666 !important;">
                                      <legend style="font-family: Calibri; font-size: 10pt; font-weight: bold;padding: 1px 10px !important;float:none;width:auto;">
                                          <img src="${imgUrlLegend}" style="width: 16px; vertical-align: middle;">
                                          <span style="vertical-align: middle;">Legend</span>
                                      </legend>
                                      <table style="font-size: 11px; font-family: Calibri;" cellspacing="0" cellpadding="1" width="100%" border="0">
                                          <tbody>
                                              <tr height="25">
                                                  <td align="center" width="22px">
                                                      <img title="Issue Not Settled" src="${imgUrlNotSettled}" style="height: 14px; width: 14px; border-width: 0px;">
                                                  </td>
                                                  <td>
                                                      <span class="Legend">Issue Not Settled</span>
                                                  </td>
                                              </tr>
                                              <tr height="25">
                                                  <td align="center">
                                                      <img title="Issue Settled" src="${imgUrlSettled}" style="height: 14px; width: 14px; border-width: 0px;">
                                                  </td>
                                                  <td>
                                                      <span class="Legend">Issue settled</span>
                                                  </td>
                                              </tr>
                                          </tbody>
                                      </table>
                                  </fieldset>
                              </td>
                          </tr>

                          <tr>
                            <td>
                                <span style="font-family:Verdana;font-size:9.5pt;font-style:italic;">Answers to questions marked with <font color="red"><b>*</b></font> must be accompanied with additional comments.</span>
                            </td>
                          </tr>

                          <tr>
                            <td>
                                <span style="font-family:Verdana;font-size:9.5pt;font-style:italic;">Responses are filled by PM <font color="red"><b>+</b></font> only when requested by GL</span>
                            </td>
                          </tr>

      </table>`;
      


      tbl += `</table>`;
      tbl += legendTbl;
      $('#dpirQL').append(tbl);
    }
}
  
let renderTradeQuestionControls = async function(index, Title, TitleId, Answer, IsRequired, Comments, NeedPMAction, PMResponse, PMStatus, isDisable){
      
    let quesId = `q${index}`;
    let padding = "padding: 0px 30px 0px 6px;"
    let yesChecked = Answer === 'Yes' ? 'checked' : '';
    let noChecked = Answer === 'No' ? 'checked' : '';
    let disable = isDisable ? 'disabled' : '';
    let tblRows = `<td style="vertical-align: top; color: black; font-weight:normal;"><label id="${quesId}title" titleId="${TitleId}">${Title}</label></td>`;
    
    tblRows += '<td style="vertical-align: top; color: black; font-weight:normal;">'
    tblRows += IsRequired === 'yes' ? 
               `<input type="radio" id="${quesId}yes" name="${quesId}" ${disable} value="yes" ${yesChecked} isRequiredOn =${IsRequired} onchange="handleRadioChange(this)" /><label for="${quesId}yes" style="${padding}">Yes<span style="color: red"> *</span></label>`:
               `<input type="radio" id="${quesId}yes" name="${quesId}" ${disable} value="yes" ${yesChecked} isRequiredOn =${IsRequired} onchange="handleRadioChange(this)" /><label for="${quesId}yes" style="${padding}">Yes</label>`
               
    tblRows += IsRequired === 'no' ?  
              `<input type="radio" id="${quesId}no" name="${quesId}" ${disable} value="no" ${noChecked} isRequiredOn =${IsRequired} onchange="handleRadioChange(this)" /><label for="${quesId}no" style="${padding}">No<span style="color: red"> *</span></label>`:
              `<input type="radio" id="${quesId}no" name="${quesId}" ${disable} value="no" ${noChecked} isRequiredOn =${IsRequired} onchange="handleRadioChange(this)" /><label for="${quesId}no" style="${padding}">No</label>`
    tblRows += '</td>';
  
    Comments = Comments === undefined || Comments === null? '' : Comments;
    let style = Answer !== undefined && Answer !== null && Answer.toLowerCase() === IsRequired ? 'placeholder="Required"' : 'placeholder="optional"';
    
    tblRows += `<td><textarea id="${quesId}com" rows="4" cols="50" ${style} ${disable}>${Comments}</textarea></td>`;
  

    let checked = NeedPMAction ? 'checked' : ''
    tblRows += `<td style="text-align: center"><input type="checkbox" id="${quesId}PMA" ${checked} ${disable} /></td>`;

    PMResponse = PMResponse !== undefined && PMResponse !== null && PMResponse !== '' ? PMResponse : ''
    tblRows += `<td><label id="${quesId}PMR">${PMResponse}</label></td>`;

    if(PMStatus !== undefined && PMStatus !== null && PMStatus !== ''){
      let imgIcon = PMStatus === 'Issue Not Settled' ? 'redclock.png' : 'Settle.png';
      let imgUrl = `${_webUrl}${_layout}/Images/${imgIcon}`
      tblRows += `<td><img id="${quesId}PMS" src=${imgUrl} alt="${imgIcon}" width="30px" title='${PMStatus}'/></td>`;
    }
    else if(NeedPMAction && PMResponse === ''){
      let imgIcon = 'redclock.png';
      let imgUrl = `${_webUrl}${_layout}/Images/${imgIcon}`
      tblRows += `<td><img id="${quesId}PMS" src=${imgUrl} alt="${imgIcon}" width="30px" title='Issue Not Settled'/></td>`;
    }
    else tblRows += `<td></td>`;
   
    return tblRows;
}
  
function handleRadioChange(element){
    let checked = element.defaultValue === 'yes' ? 'yes' : 'no';
    let isRequiredOn = element.getAttribute('isRequiredOn');
    let placeholderText = isRequiredOn === checked ? 'Required' : 'Optional';
  
    let Id = element.id.replace('yes', 'com').replace('no', 'com');
    $(`#${Id}`).attr('placeholder', placeholderText);
}
  
const isTradeQuestionValid = async function(validate) {
    let tblRows = $('#tbltradeQId tr').length - 1;
    let _mesg = '';
  
    var itemsToInsert = [];
    for(let i = 0; i < tblRows; i++){
      var _objValue = { };
      var radios = document.getElementsByName(`q${i}`);
      var formValid = false;
  
      for (var j = 0; j < radios.length; j++){
          if (radios[j].checked) {
              let checked = radios[j].defaultValue === 'yes' ? 'Yes' : 'No';
              _objValue['Answer'] = checked;
              formValid = true;
              break;
          }
      }
  
      let comment = document.getElementById(`q${i}com`);
      if (!formValid){
         _mesg = 'Please select one Answer option.'
          comment.placeholder = _mesg;
          setRequiredFieldErrorStyle(comment);
          if(validate)
             break;
      }
      else _objValue['Comments'] = comment.value.trim();
  
  
      let element = document.getElementById(`q${i}title`);
      let titleId = element.getAttribute('titleId');
      let titleLookupValue = element.innerText.trim();

      _objValue['Title'] = titleLookupValue;
      _objValue['QuestionId'] = parseInt(titleId);
      //_objValue['MasterIDId'] = _itemId;

      debugger;
      _objValue['RDIDId'] = deptId;
      let reqPMAction = document.getElementById(`q${i}PMA`);
      let pmChecked = reqPMAction.checked ? true : false;
      _objValue['NeedPMAction'] = pmChecked;
      _objValue['Department'] = dept;

      itemsToInsert.push(_objValue);
    }
  
    if(_mesg === ''){
      for(let i = 0; i < tblRows; i++){
        let comment = document.getElementById(`q${i}com`);
        if(comment.placeholder === 'Required'){
          if(comment.value.trim() === ''){
            _mesg = 'Comment is Required Field';
            setRequiredFieldErrorStyle(comment);
            
            comment.addEventListener('change', function() {
              if(this.value.trim() === '')
                setRequiredFieldErrorStyle(this, false);
              else setRequiredFieldErrorStyle(this, true);
            });
            
          }
        }
      }
    }
  
    return {
      mesg: _mesg, // Allow form submission
      items: itemsToInsert
    }
}

let insertTradeQuestions = async function(itemsToInsert, targetList, colsInternal) {
    const list = _web.lists.getByTitle(targetList);
    const batch = pnp.sp.createBatch();
    const columns = colsInternal.join(',');
  
    for (const item of itemsToInsert){

      // let _query = `Question/Title eq '${item.Title}' and MasterID/Id eq ${_itemId} and Department eq '${TextQueryEncode(item.Department)}'`;
  
      // const existingItem = await list.items
      // .select("Id,MasterID/Id,Question/Title," + columns)
      // .expand("MasterID,Question")
      // .filter(_query)
      // .top(1)
      // .get();


      let _query = `Question/Title eq '${item.Title}' and RDID/Id eq ${deptId} and Department eq '${TextQueryEncode(item.Department)}'`;
  
      const existingItem = await list.items
      .select("Id,RDID/Id,Question/Title," + columns)
      .expand("RDID,Question")
      .filter(_query)
      .top(1)
      .get();
  
      if (existingItem.length > 0) {
        var _item = existingItem[0]; // _item is the oldItem values while objValue is the new one
        var doUpdate = false;
  
        for(var i = 0; i < colsInternal.length; i++){
          var columnName = colsInternal[i];
          var oldValue = '', currentValue = '';
  
          oldValue = _item[columnName];
          currentValue = item[columnName];
  
          if( oldValue != currentValue ){
              doUpdate = true;
              break;
          }
        }
        if(doUpdate)
         list.items.getById(existingItem[0].Id).inBatch(batch).update(item);
      } else list.items.inBatch(batch).add({
                 //MasterIDId: _itemId,
                 RDIDId: deptId,
                 QuestionId: item.QuestionId,
                 Title: item.Title,
                 Comments: item.Comments,
                 Answer: item.Answer,
                 NeedPMAction: item.NeedPMAction,
                 Department: item.Department
              });
      
  
    }
    if(batch._index > -1)
     await batch.execute();
}

let checkAllSubmitted = async function(){

  let items = await getReviewTrades(_itemId);
  debugger;
  let openTrades = items.filter(item=> { return item.Status === 'In Progress' || item.Status === 'Reject' })
  if(openTrades.length === 0){
    
    let emailName = 'AllSubmitted_Trade_Email', notName = '';
    let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${items[0].Id}</Value></Eq></Where>`;
      await _sendEmail(_module, emailName, query, '', notName, '')
  }
}