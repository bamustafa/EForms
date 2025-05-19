var  phases = [], isProjCompleted = false, Reptype = '';

let isTrade = false;
let listnameFromUrl;

var onCompletionReportRender = async function(){

    const startTime = performance.now();
    if(!_isPM && !_isPD && !_isQM)
        isTrade = true;

    formFields = {
        ReportType: fd.field("ReportType"),
        Phases: fd.field("Phases")
    }

    const list = await _web.lists.getByTitle(CompReports).select("RootFolder/ServerRelativeUrl").expand("RootFolder").get();
    const listUrl = list.RootFolder.ServerRelativeUrl;
    listnameFromUrl = listUrl.substring(listUrl.lastIndexOf('/') + 1);

    clearFormFields([formFields.ReportType.internalName, formFields.Phases.internalName]);
    
    await getItems()
    .then(async (items)=>{
        if(items.length > 0){
            handleType(items);
        }else{
          var compReportTab = [
                      {
                        masterTab: firstChildTab,
                        title: 'Completion Report',
                        tooltip: 'no phases found'
                      }
                    ];
          await enable_Disable_Tabs(compReportTab, true)
        }
    })

    if(isTrade)
       await _HideFields([formFields.Phases], true);

    const endTime = performance.now();
    const elapsedTime = endTime - startTime;
    console.log(`onCompletionReportRender: ${elapsedTime} milliseconds`);
}

let getItems = async function(){
var itemArray = [];
  
  const list = _web.lists.getByTitle(CompReports);
  let query = `MasterID/Id eq ${_itemId}`;

  if(isTrade){
    let items = JSON.parse(localStorage.getItem('GLTrades'));
    if(items === null){query += ' and Title eq null'}
    else {
        
        let suQuery = ' and ('
        items.map((item, i)=>{
            let trade= TextQueryEncode(item.Title)
            if(i > 0)
                suQuery += ' or '
           suQuery += ` Title eq '${trade}'`
        })
        suQuery += ')'

        query += suQuery;
    }
  }

  let items = await list.items
                 .select("Id,MasterID/Id,Title,Status,Phases,IsProjectComp,RejRemarks")
                 .expand("MasterID")
                 .filter(query)
                 .orderBy("Title", true)
                 .getAll();
  
  if(items.length > 0){
    for(var i = 0; i < items.length; i++){
      var item = items[i];
      var rowData  = {};

      rowData["Id"] = item["Id"];
      rowData["Title"] = item.Title;
      rowData["Status"] = item.Status;

      let phase = item["Phases"];
      rowData["Phases"] =  phase !== undefined && phase !== null ? phase : '';
      if(phase !== '' && !phases.includes(phase))
        phases.push(phase);

      rowData["IsProjectComp"] = item.IsProjectComp;
      if(item.IsProjectComp)
        isProjCompleted = true

      rowData["RejRemarks"] = item.RejRemarks;
      itemArray.push(rowData);
    }
  }
  return itemArray;
}

let handleType = async function(items){

    if(phases.length === 0 || _isLLChecker){
        formFields.ReportType.value = 'Project';
        $(formFields.Phases.$parent.$el).hide();

        if(_isLLChecker){
          let inputElement = $("input[value='Phase']");
          inputElement.attr('disabled', true);
          inputElement.parent().attr('title', 'permission Denied');
        }
    }
    else {
        formFields.ReportType.value = 'Phase';
        formFields.Phases.options = phases;
        setTimeout(async () => { 
            await setPhaseStatus(items)
          }, 100);
    }

    formFields.Phases.$on("change", async function (phase){
       
        $('#pmtrade').remove()
     
        if(phase !== null && phase !== ''){
            let pmLink = '', isTradesApproved = true;
            let filterItems = items.filter((item)=>{ 
                if(item.Phases === phase && item.Title === 'PM'){
                  let linkUrl = `${_webUrl}/SitePages/PlumsailForms/${listnameFromUrl}/Item/EditForm.aspx?item=${item.Id}`;
                  pmLink = `<a id="pmtrade" itemid="${item.Id}" role="button" href= ${linkUrl} onclick="openInCustomWindow(event)">Proceed to finalize PM task</a>`
                }
                
              if(item.Phases === phase && item.Title !== 'PM' && item.Status !== 'Approve')
                isTradesApproved = false
                
              return item.Phases === phase && item.Title !== 'PM'
            });

            if(!isTradesApproved)
              pmLink = ''; // remove pmlink if not all trades approved for phase

            if(isTrade)
              filterItems = items.filter((item)=>{ return item.IsProjectComp !== true });
           
            filterItems = await sortItems(filterItems);
            await renderCRTable(filterItems, pmLink)
        }
    })

    if(formFields.ReportType.value !== undefined && formFields.ReportType.value !== null && formFields.ReportType.value !== '')
      Reptype = formFields.ReportType.value  === 'Phase' ? 'PhCR' : 'PCR'

    formFields.ReportType.$on("change", async function (value){
       //debugger;
        $('#pmtrade').remove()
        clearFormFields([formFields.Phases.internalName]);
        Reptype = value === 'Phase' ? 'PhCR' : 'PCR'
        let pmLink = '', isTradesApproved = true;
       
        let filterItems = items.filter((item)=>{
            if(Reptype === 'PCR'){
              if(item.IsProjectComp && item.Title === 'PM'){
                let linkUrl = `${_webUrl}/SitePages/PlumsailForms/${listnameFromUrl}/Item/EditForm.aspx?item=${item.Id}`;
                pmLink = `<a id="pmtrade" itemid="${item.Id}" role="button" href= ${linkUrl} onclick="openInCustomWindow(event)">Proceed to finalize PM task</a>`
              }
             
              if(item.IsProjectComp === true && item.Title !== 'PM' && item.Status !== 'Approve')
                isTradesApproved = false

              $(formFields.Phases.$parent.$el).hide();
              return item.IsProjectComp === true && item.Title !== 'PM'
            }
            else {
                if(!isTrade)
                  $(formFields.Phases.$parent.$el).show();
                else return item.IsProjectComp === false
            }
        });

        if(!isTradesApproved)
          pmLink = ''; // remove pmlink if not all trades approved for phase

        filterItems = await sortItems(filterItems);
        await renderCRTable(filterItems, pmLink)
        console.log(filterItems);
    })


     let filterItems = items.filter((item)=>{
        return item.IsProjectComp === true
     });

     if(filterItems.length === 0){
        let inputElement = $("input[value='Project']");
        inputElement.attr('disabled', true);
        inputElement.parent().attr('title', 'once the project status is completed PCR can be filled');
     }
}

let setPhaseStatus = async function(items){

    const list = _web.lists.getByTitle(CompReports);
    for(var i = 0; i < phases.length; i++){

        let phase = phases[i];
        let status = 'In Progress';
        let textColor = 'color:#8677b1';

        let allItemsTotal = items.filter((item)=>{
              return item.Phases === phase && item.Status !== 'Approve'
        })

        if(allItemsTotal.length === 0){
            status = 'Approved'
            textColor = 'color:#49c4b1';
        }
        else if(allItemsTotal.length === 1){
          let item = allItemsTotal[0];

          if(item.Title === 'PM'){
            status = 'Ready to Finalize'
            textColor = 'color:red';
          }
        }

        if(!isTrade){
          let lblText = $(`label:contains('${phase}')`)
          let spanStat = `<span style="font-size: 12px; font-weight:bold; text-align: center;${textColor}">${status}</span>`;
          $(lblText).append(spanStat);
        }
    }
}

let sortItems = async function(filterItems){
    return filterItems.sort((a, b) => {
        if (a.Title < b.Title) return -1;
        if (a.Title > b.Title) return 1;
        return 0;
    });
}

let renderCRTable = async function(items, pmLink){
    let resultCount = items !== null ? items.length : 0;
    
    let tableId = 'tblCRId'
    if($(`#${tableId}`).length > 0)
        $(`#${tableId}`).remove()

    if(resultCount === 0)
        return;

    
      let thdCol = isTrade ? 'Phase' : 'Action'
      let tbl = `<table id='${tableId}' width='100%' class='modern-table'> 
                    <tr>
                      <th width='40%'><label class="d-flex fd-field-title col-form-label center-text">Department</label></th>
                      <th width='10%'><label class="d-flex fd-field-title col-form-label center-text">Status</label></th>
                      <th width='50%'><label class="d-flex fd-field-title col-form-label center-text">${thdCol}</label></th>
                    </tr>`;
  
      for(var i = 0; i < resultCount; i++){
        let {Id, Title, Status, RejRemarks, Phases} = items[i];
        let tradeItemUrl = `${_webUrl}/SitePages/PlumsailForms/${listnameFromUrl}/Item/EditForm.aspx?item=${Id}`;

        //let customText = GLMainId.length === 0 ? 'GL Main is not assigned' : '';
        let bgcolor = Status === 'Approve' ? `color:#49c4b1 `: ''
        tbl += `<tr style="${bgcolor}">`
        tbl += await renderCRItemControls(i+1, Id, Title, Status, RejRemarks, Phases, tradeItemUrl);
        tbl += '</tr>'
      }
      if(pmLink !== '')
        tbl += `<tr style="text-align: right">
                  <td colspan='3'>
                    ${pmLink}
                  </td>
                </tr>`;
      
      tbl += `</table>`;
      $('#PCR').append(tbl);
    
}

let renderCRItemControls = async function(index, Id, Trade, Status, RejRemarks, Phase, tradeItemUrl){
      
    if(RejRemarks === null || RejRemarks === undefined)
       RejRemarks = ''

    let rowId = `${index}`;
    let padding = "padding: 0px 30px 0px 6px;"
    let appChecked = Status === 'Approve' ? 'checked' : '';
    let rejChecked = Status === 'Reject' ? 'checked' : '';

    let style = Status === 'Reject' ? '' : 'display:none';
    let textColor = '';
    
    if(Status === 'Approve')
      textColor = 'color:#49c4b1;font-weight:bold;'
    else if(Status === 'Reject')
      textColor = 'color:red ;font-weight:bold;'
     else if(Status === 'In Progress')
      textColor = 'color:#8677b1;font-weight:bold;'

    let isDiabled = '', isReadOnly = '';
    if(Status !== 'Submitted' || (!_isPM && !_isPD)){
      isDiabled = 'disabled'
      isReadOnly = 'readonly'
    }

    let leadTag = ''; //ISLead ? `<span style="color: green"><b>(Lead Dept)</b></span>` : '';
   
    let customMesg = ''; //customText !== '' ? `<span style="font-size: 12px; background-color: yellow;">${customText}<span>` : ''

    let thdCol = isTrade ? 'Phase' : 'Action'
    let tblRows = `<td style="vertical-align: top; color: black; font-weight:normal;">
                     <a id="${rowId}trade" itemId = ${Id} role="button" href="${tradeItemUrl}" onclick="openInCustomWindow(event)">${Trade} ${leadTag}</a>
                     ${customMesg}
                   </td>

                   <td style="vertical-align: top; color: black; font-weight:normal; text-align: center;${textColor}"><label id="${rowId}status">${Status}</label></td>`

       tblRows += isTrade ? 
                    `<td style="vertical-align: top; color: black; font-weight:normal; text-align: center;"><label id="${rowId}phase">${Phase}</label></td>`
                    : 
                    `<td style="vertical-align: top; color: black; font-weight:normal; text-align: center">
                     <input type="radio" id="${rowId}${Reptype}Approve" name="${rowId}" value="Approve" ${appChecked} ${isDiabled} onchange="handleCRRadioChange(this)" />
                        <label for="${rowId}Approve" style="${padding}">Approve</label>

                     <input type="radio" id="${rowId}${Reptype}Reject" name="${rowId}" value="Reject" ${rejChecked} ${isDiabled} onchange="handleCRRadioChange(this)" />
                        <label for="${rowId}Reject" style="${padding}">Reject</label>

                        <br/>
                        <textarea id="${rowId}${Reptype}com" rows="4" cols="70" style="${style}" ${isReadOnly} placeholder="Required">${RejRemarks}</textarea>
                   </td>`
    return tblRows;
}

function handleCRRadioChange(element){
    let value = element.value
    let Id = element.id.replace(`${Reptype}Approve`, `${Reptype}com`).replace(`${Reptype}Reject`, `${Reptype}com`);
   
    if(value === 'Reject')
       $(`#${Id}`).show();
    else $(`#${Id}`).hide();  
}