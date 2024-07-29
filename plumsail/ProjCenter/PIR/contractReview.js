var getTemplateItems = async function(targetList, colsInternal){
    var itemArray = [];
    
    const list = _web.lists.getByTitle(ContractReview);
    let items = await list.items.select("Id,MasterID/Id,Question/Id,Question/Title,Question/IsRequired,Answer,Comments").expand("MasterID,Question").filter(`MasterID/Id eq ${_itemId}`).get();
    
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
        itemArray.push(rowData);
      }
    }
  
    else{
      const cols = colsInternal.join(',');
      var _query = "Title ne null";
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
  
  var insertItemsInBulk = async function(itemsToInsert, targetList, colsInternal) {
    const list = _web.lists.getByTitle(targetList);
    const batch = pnp.sp.createBatch();
    const columns = colsInternal.join(',');
  
    for (const item of itemsToInsert){
  
      let _query = `Question/Title eq '${item.Title}' and MasterID/Id eq ${_itemId}`;
  
      const existingItem = await list.items
      .select("Id,MasterID/Id,Question/Title," + columns)
      .expand("MasterID,Question")
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
                 MasterIDId: _itemId,
                 QuestionId: item.QuestionId,
                 Title: item.Title,
                 Comments: item.Comments,
                 Answer: item.Answer
              });
      
  
    }
    if(batch._index > -1)
     await batch.execute();
  }
  
  var onContReviewRender = async function(){
  
    var result = await getTemplateItems(Questions, ['Id', 'Title', 'IsRequired']);
    await renderCustomTable(result);
  }
  
  var renderCustomTable = async function(result){
    let resultCount = result !== null ? result.length : 0;
    
    let tableId = 'tblItemsId'
    let divElement = $('');
    if($(`#${tableId}`).length === 0){
      let tbl = `</br><table id='${tableId}' width='100%' class='modern-table'> 
                    <tr>
                      <th width='50%'><label class="d-flex fd-field-title col-form-label">Question</label></th>
                      <th width='20%'><label class="d-flex fd-field-title col-form-label">Answer</label></th>
                      <th width='20%'><label class="d-flex fd-field-title col-form-label">Comments</label></th>
                    </tr>`;
  
      for(var i = 0; i < resultCount; i++){
        let {Id, Title, TitleId, IsRequired, Answer, Comments} = result[i];
        TitleId = TitleId === undefined ? Id : TitleId;
        tbl += '<tr>'
        tbl += await renderQuestionControls(i, Title, TitleId, Answer, IsRequired.toLowerCase(), Comments);
        tbl += '</tr>'
      }
      tbl += `</table>`;
      $('#QL').append(tbl);
    }
  }
  
  var renderQuestionControls = async function(index, Title, TitleId, Answer, IsRequired, Comments){
      
    let quesId = `q${index}`;
    let padding = "padding: 0px 30px 0px 6px;"
    let yesChecked = Answer === 'Yes' ? 'checked' : '';
    let noChecked = Answer === 'No' ? 'checked' : '';
  
    let tblRows = `<td style="vertical-align: top; color: black; font-weight:normal;"><label id="${quesId}title" titleId="${TitleId}">${Title}</label></td>`;
    
    tblRows += '<td style="vertical-align: top; color: black; font-weight:normal;">'
    tblRows += IsRequired === 'yes' ? 
               `<input type="radio" id="${quesId}yes" name="${quesId}" value="yes" ${yesChecked} isRequiredOn =${IsRequired} onchange="handleRadioChange(this)" /><label for="${quesId}yes" style="${padding}">Yes<span style="color: red"> *</span></label>`:
               `<input type="radio" id="${quesId}yes" name="${quesId}" value="yes" ${yesChecked} isRequiredOn =${IsRequired} onchange="handleRadioChange(this)" /><label for="${quesId}yes" style="${padding}">Yes</label>`
               
    tblRows += IsRequired === 'no' ?  
              `<input type="radio" id="${quesId}no" name="${quesId}" value="no" ${noChecked} isRequiredOn =${IsRequired} onchange="handleRadioChange(this)" /><label for="${quesId}no" style="${padding}">No<span style="color: red"> *</span></label>`:
              `<input type="radio" id="${quesId}no" name="${quesId}" value="no" ${noChecked} isRequiredOn =${IsRequired} onchange="handleRadioChange(this)" /><label for="${quesId}no" style="${padding}">No</label>`
    tblRows += '</td>';
  
    Comments = Comments === undefined || Comments === null? '' : Comments;
    let style = Answer !== undefined && Answer !== null && Answer.toLowerCase() === IsRequired ? 'placeholder="Required"' : 'placeholder="optional"';
    
    tblRows += `<td><textarea id="${quesId}com" rows="4" cols="50" ${style}>${Comments}</textarea></td>`;
  
    return tblRows;
  }
  
  function handleRadioChange(element){
    let checked = element.defaultValue === 'yes' ? 'yes' : 'no';
    let isRequiredOn = element.getAttribute('isRequiredOn');
    let placeholderText = isRequiredOn === checked ? 'Required' : 'Optional';
  
    let Id = element.id.replace('yes', 'com').replace('no', 'com');
    $(`#${Id}`).attr('placeholder', placeholderText);
  }
  
  const isQuestionValid = async function() {
    let tblRows = $('#tblItemsId tr').length - 1;
    let _mesg = '';
  
    var itemsToInsert = [];
    for(let i = 0; i < tblRows; i++){
      var _objValue = { };
      var radios = document.getElementsByName(`q${i}`);
      var formValid = false;
  
      for (var j = 0; j < radios.length; j++) {
          if (radios[j].checked) {
              let checked = radios[j].defaultValue === 'yes' ? 'Yes' : 'No';
              _objValue['Answer'] = checked;
              formValid = true;
              break;
          }
      }
  
      let comment = document.getElementById(`q${i}com`);
      if (!formValid) {
         _mesg = 'Please select one Answer option.'
          comment.placeholder = _mesg;
          break;
      }
      else _objValue['Comments'] = comment.value.trim();
  
  
      let element = document.getElementById(`q${i}title`);
      let titleId = element.getAttribute('titleId');
      let titleLookupValue = element.innerText.trim();
  
      _objValue['Title'] = titleLookupValue;
      _objValue['QuestionId'] = parseInt(titleId);
      _objValue['MasterIDId'] = _itemId;
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