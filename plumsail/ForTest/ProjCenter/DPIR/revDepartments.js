
let pmTblname = "pmdt"

var onRevDepRender = async function(){
    const startTime = performance.now();
    
    let items = JSON.parse(localStorage.getItem('GLTrades')); //await getReviewTrades(_itemId);
    if(items === null || (!_isPM && !_isPD)){
      $('button:contains("Notify")').remove();
      if(items === null)
        return;
      else fd.control(pmTblname).readonly = true;
    }
      
    await handlePMTable();

    items = items.filter((item)=>{ return item.Title !== 'Area Operations' })
    await renderTable(items);

    let disableButton = false;

    let isFound = items.filter(item=>{
      return !item.IsNotified;
   })

   if(isFound.length === 0)
    disableButton = true;
   
    await notifyTrade(items, disableButton);

    const endTime = performance.now();
    const elapsedTime = endTime - startTime;
    console.log(`onRevDepRender: ${elapsedTime} milliseconds`);
}

let renderTable = async function(result){
    let resultCount = result !== null ? result.length : 0;
    
    let tableId = 'tblTradeId'
    if($(`#${tableId}`).length === 0){

      const list = await _web.lists.getByTitle(RevDepartments).select("RootFolder/ServerRelativeUrl").expand("RootFolder").get();
      const listUrl = list.RootFolder.ServerRelativeUrl;
      const listnameFromUrl = listUrl.substring(listUrl.lastIndexOf('/') + 1);
      
      
      let tbl = `</br><table id='${tableId}' width='100%' class='modern-table'> 
                    <tr>
                      <th width='40%'><label class="d-flex fd-field-title col-form-label center-text">Department</label></th>
                      <th width='10%'><label class="d-flex fd-field-title col-form-label center-text">Status</label></th>
                      <th width='50%'><label class="d-flex fd-field-title col-form-label center-text">Action</label></th>
                    </tr>`;
  
      for(var i = 0; i < resultCount; i++){
        let {Id, Title, Status, RejRemarks, ISLead, GLMainId} = result[i];
        let tradeItemUrl = `${_webUrl}/SitePages/PlumsailForms/${listnameFromUrl}/Item/EditForm.aspx?item=${Id}`;

        let customText = GLMainId.length === 0 ? 'GL Main is not assigned' : '';
        let bgcolor = Status === 'Approve' ? `background: ${approvedBgColor}`: ''
        tbl += `<tr style="${bgcolor}">`
        tbl += await renderItemControls(i+1, Id, Title, Status, RejRemarks, tradeItemUrl, ISLead, customText);
        tbl += '</tr>'
      }
      tbl += `</table>`;
      $('#DPR').append(tbl);
    }
}

let renderItemControls = async function(index, Id, Trade, Status, RejRemarks, tradeItemUrl, ISLead, customText){
      
    if(RejRemarks === null || RejRemarks === undefined)
       RejRemarks = ''

    let quesId = `${index}`;
    let padding = "padding: 0px 30px 0px 6px;"
    let appChecked = Status === 'Approve' ? 'checked' : '';
    let rejChecked = Status === 'Reject' ? 'checked' : '';


    let style = Status === 'Reject' ? '' : 'display:none';
    let textColor = '';
    
    if(Status === 'Approve')
      textColor = 'color:green;font-weight:bold;'
    else if(Status === 'Reject')
      textColor = 'color:red;font-weight:bold;'
     else if(Status === 'Submitted')
      textColor = 'color:#936106cf;font-weight:bold;'

    let isDiabled = '', isReadOnly = '';
    if(Status !== 'Submitted' || (!_isPM && !_isPD)){
      isDiabled = 'disabled'
      isReadOnly = 'readonly'
    }
    //<label id="${quesId}trade" itemId = ${Id}>${Trade}</label>
    
    //<img src=${_webUrl}${_layout}/Images/BallGreen.ico} style="height:14px;width:14px;border-width:0px;"> 

    let leadTag = ISLead ? `<span style="color: green"><b>(Lead Dept)</b></span>` : '';
   

    let customMesg = customText !== '' ? `<span style="font-size: 12px; background-color: yellow;">${customText}<span>` : ''
    let tblRows = `<td style="vertical-align: top; color: black; font-weight:normal;">
                     <a id="${quesId}trade" itemId = ${Id} role="button" href="${tradeItemUrl}" onclick="openInCustomWindow(event)">${Trade} ${leadTag}</a>
                     ${customMesg}
                   </td>


                   <td style="vertical-align: top; color: black; font-weight:normal; text-align: center;${textColor}"><label id="${quesId}status">${Status}</label></td>

                   <td style="vertical-align: top; color: black; font-weight:normal; text-align: center">
                     <input type="radio" id="${quesId}Approve" name="${quesId}" value="Approve" ${appChecked} ${isDiabled} onchange="handleTradeRadioChange(this)" />
                        <label for="${quesId}Approve" style="${padding}">Approve</label>

                     <input type="radio" id="${quesId}Reject" name="${quesId}" value="Reject" ${rejChecked} ${isDiabled} onchange="handleTradeRadioChange(this)" />
                        <label for="${quesId}Reject" style="${padding}">Reject</label>

                        <br/>
                        <textarea id="${quesId}com" rows="4" cols="70" style="${style}" ${isReadOnly} placeholder="Required">${RejRemarks}</textarea>
                   </td>`
    return tblRows;
}
  
function handleTradeRadioChange(element){
    let value = element.value
    let Id = element.id.replace('Approve', 'com').replace('Reject', 'com');
   
    if(value === 'Reject')
       $(`#${Id}`).show();
    else $(`#${Id}`).hide();  
}

let updateDPIRItems = async function(itemsToInsert, colsInternal){

    const list = _web.lists.getByTitle(RevDepartments);
    const batch = pnp.sp.createBatch();
   
    for (const item of itemsToInsert){
      let columns = colsInternal.join(',');
      let RejBy = item['RejById']
      let AppBy = item['AppById']

      let _query = `Id eq ${item.Id}`;
      if(RejBy !== undefined)
        columns += ',RejBy/Id'
      else if(AppBy !== undefined)
        columns += ',AppBy/Id'
  
      let existingItem = await list.items.select("Id," + columns)

       if(RejBy !== undefined)
        existingItem = existingItem.expand('RejBy')
      else if(AppBy !== undefined)
        existingItem = existingItem.expand('AppBy')
       

       let result = await existingItem
                          .filter(_query)
                          .top(1)
                          .get();
  
      if (result.length > 0){
        var _item = result[0]; // _item is the oldItem values while objValue is the new one
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
          //await list.items.getById(_item.Id).update(item); 
         list.items.getById(_item.Id).inBatch(batch).update(item);
      }
    }
    if(batch._index > -1)
     await batch.execute();
}

const isDPIR_RowValid = async function() {
  let tblRows = $('#tblTradeId tr').length - 1;
  let _mesg = '';

  var itemsToInsert = [];
  let currentUser = await GetCurrentUser();
  //currentUser = currentUser; //.LoginName.split('|')[1];

  for(let i = 0; i < tblRows; i++){

    let ctlIndex = i +1;
    let _objValue = { };
    let appRadioCtl = document.getElementById(`${ctlIndex}Approve`);
    let rejRadioCtl = document.getElementById(`${ctlIndex}Reject`);
    let rejChecked = rejRadioCtl.checked;
    let formValid = false;

    if (rejChecked){
      let comment = document.getElementById(`${ctlIndex}com`);

      if(comment.value.trim() !== ''){
        _objValue['RejRemarks'] = comment.value.trim();
        formValid = true;
      }
      else _mesg = 'NA';

      comment.addEventListener('change', function(){
        if(this.value.trim() === '')
          setRequiredFieldErrorStyle(this, false);
        else setRequiredFieldErrorStyle(this, true);
      });
    }
  
    let trade = $(`#${ctlIndex}trade`).text()
    let itemId = $(`#${ctlIndex}trade`).attr('itemId');

    let status = $(`#${ctlIndex}status`).text();

    if(rejChecked) 
      status = 'Reject'
    else if(appRadioCtl.checked)
      status = 'Approve'

    _objValue['Id'] = itemId;
    _objValue['Title'] = trade;
    _objValue['Status'] = status;


    if(status === 'Reject'){
       let col = 'RejById'
      _objValue['RejDate'] = new Date();
      _objValue[col] = currentUser.Id
    }
    else if(status === 'Approved'){
       let col = 'AppById'
      _objValue['AppDate'] = new Date();
      _objValue[col] = currentUser.Id
    }
    //_objValue['MasterIDId'] = _itemId;
    itemsToInsert.push(_objValue);
  }

  return {
    mesg: _mesg, // Allow form submission
    dpirItems: itemsToInsert
  }
}

const handlePMTable = async function(){
    
    let dt = fd.control(pmTblname);
    dt.selectable = false;

    let filterQuery = ` <And>
                          <Eq><FieldRef Name='MasterID' /><Value Type='Lookup'>${_itemId}</Value></Eq>
                          <Eq><FieldRef Name='NeedPMAction' /><Value Type='Boolean'>1</Value></Eq>
                        </And>`

    // let filterQuery = ` <And>
    //                       <Eq><FieldRef Name='RDID' /><Value Type='Lookup'>${_itemId}</Value></Eq>
    //                       <Eq><FieldRef Name='NeedPMAction' /><Value Type='Boolean'>1</Value></Eq>
    //                     </And>`
    dt.filter =  filterQuery; //'<Eq><FieldRef Name="NeedPMAction"/><Value Type="Text">Test</Value></Eq>';
    dt.refresh();
    
    dt.$on('edit', function(editData){

      //setTimeout(async () => {
         editData.field('Department').disabled = true;
         editData.field('Question').disabled = true;
         editData.field('Answer').disabled = true;
         editData.field('Comments').disabled = true;
         editData.field('PMStatus').disabled = true;
    //}, 500);
      


   
      let input = $("div[internal-name='pmdt']").find("input.form-check-input");
      $(input).on('change', function(){
         let ctlr = this;
         let stat = ''
         if (ctlr.checked){
            stat = 'Issue Settled'
            //imgElement();
         }
         //else imgElement('remove');

         editData.field('PMStatus').value = stat;
         
      })
    });

    setTimeout(async () => {
       let imgUrlSettled = `${_webUrl}${_layout}/Images/Settle.png`;
       let res = $('td:contains("Issue Settled")')

       $('td:contains("Issue Settled")').parent().find('a.k-grid-edit').remove()

      $('td:contains("Issue Settled")')
       .css('text-align','center')
      .html(`<img title="Issue Settled" src="${imgUrlSettled}" style="height: 14px; width: 14px; border-width: 0px;">`);

      $('td:contains("No")')
      .css('text-align','center')

      $('td:contains("Yes")')
      .css('text-align','center')

      $('i.ms-Icon--CheckMark').parent().css('text-align','center')
    }, 300);

}

function imgElement(trans){
  let imgElement = $("img[Title='Issue Settled']")
  if(trans === 'remove'){
    $("img[Title='Issue Settled']").remove();
    return
  }
  if(imgElement.length === 0){
      let selector = $('td[data-container-for="PMStatus"]');
      let imgUrlSettled = `${_webUrl}${_layout}/Images/Settle.png`
      $('<img>', {
        title: "Issue Settled",
        src: imgUrlSettled,
        css: {
            height: "14px",
            width: "14px",
            borderWidth: "0px"
        }
    }).appendTo(selector);
  }
}

let notifyTrade = async function(result, disableButton){

  let btn = $('button').filter(function(){ return $(this).text().trim() == 'Notify'; });

  if(disableButton)
    disableBtn(btn)
  
  btn.hover(
    function() {
        $(this).css('cursor', 'pointer');
    },
    function() {
        $(this).css('cursor', 'default');
    }
);

  btn.on('click', async function(){
    for(var i = 0; i < result.length; i++){
      
      let {Id, Title, Status, RejRemarks, IsNotified} = result[i];
      //let tradeItemUrl = `${_webUrl}/SitePages/PlumsailForms/${listnameFromUrl}/Item/EditForm.aspx?item=${Id}`;

      if(IsNotified)
        continue;

      
      let emailName = 'Notify_Email', notName = '';
      let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${Id}</Value></Eq></Where>`;
      await _sendEmail(_module, emailName, query, '', notName, '')
      .then(async ()=>{
        // await _web.lists.getByTitle(RevDepartments).items.getById(Id).update({
				// 	IsNotified: true
				// });
      });
    }
    disableBtn(btn);
  })
}

function disableBtn(btn){
  btn.prop('disabled', true);
  btn.css({
    'background-color': 'white',
    'color': 'black'
  })
}

let notifyRejectedTrades = async function(){
  
  $('#tblTradeId tr').each(async function(rowIndex, rowElement) {

    if (rowIndex === 0)
        return true;

    let statusCell = $(`#${rowIndex}status`).text()
    let rejCell = $(`#${rowIndex}Reject`).prop('checked')

    if(statusCell === 'Submitted' && rejCell){
      let Id = $(`#${rowIndex}trade`).attr('itemid');
      let emailName = 'Reject_Trade_Email', notName = '';
      let query = `<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${Id}</Value></Eq></Where>`;
      await _sendEmail(_module, emailName, query, '', notName, '')
    }

  });
}

