let buildAusMPForm = async function(reference){
 
    let items = await getMPItems(reference);
    let htmlContent = `<div class="container">
                            <!-- Left side: Items -->
                            <div class="items">
                                <h3>Master Plan Procedures</h3>
                                 <ul id="item-list" style="list-style-type: square">`

    for (const itemNo in items){
        const item = items[itemNo];
        const procedures = item.Procedures;

        let procs = ''
        for (const procedure of procedures){
            procs += procedure.Description + '<br/>'
        }

        const itemUrl = `${_webUrl}/SitePages/PlumsailForms/MasterPlan/Item/EditForm.aspx?item=${item.Id}`
        htmlContent += `<li><a href="${itemUrl}" mpId = "${item.Id}" target="formFrame">${procs}</a></li>`
    }
        
    htmlContent += `</ul>
                    </div>
                        <!-- Right side: Form in an iframe -->
                        <div class="form">
                            <iframe Id="formFrame" name="formFrame"></iframe>
                        </div>
                   </div>`
   $('#ausMP').append(htmlContent)

   $('#ausMP').children().css({ 
        'margin': '0px', 
        'max-width': '100%' 
    });

   listItems();

//    window.addEventListener('message', async function(event){
//         debugger;
//         const value = event.data;
//         let formUrl = value.formUrl
//         let ProjName = value.ProjName
//         let formName = value.formName
//    });
}

const getMPItems = async function(reference){
   
	return await _web.lists.getByTitle(MasterPlan).items
    .select("Id,Procedures/Description")
    .expand('Procedures')
    .filter(`MasterID/Reference eq '${reference}'`)
    .getAll()
}

function listItems(){

    const links = document.querySelectorAll('.items a');

    links.forEach(link => {
      link.addEventListener('click', function() {
        links.forEach(link => link.classList.remove('active'));
        this.classList.add('active');
        const linkText = this.textContent;

        let mpIdValue = $(this).attr('mpId');
        localStorage.setItem('MPID', mpIdValue)

        $('#formFrame').on('load', function() {
            setTimeout(function () {
                let content = $('#formFrame').contents();
                adjustMPFormStyles(content, linkText, mpIdValue);
            }, 1000);
        });
      });
    });
}

function adjustMPFormStyles(iframeDocument, linkText) {
   // const iframeDocument = $('#formFrame').contents(); // Get the iframe document

    if (iframeDocument.length) {

            // iframeDocument.find('.pageContainer_e9e4af8d, .container_e9e4af8d, .pageContainer_5a558a10, .pageContainer_972f0ed1')
            // .css({ 'margin-left': '-98px'});
            
            iframeDocument.find('div.fd-toolbar-side-commands').remove()

            let xx =  iframeDocument.find('.homePageContent_39ae25bf, .SPCanvas-canvas, .CanvasZone')
            iframeDocument.find('.homePageContent_39ae25bf, .SPCanvas-canvas, .CanvasZone')
            .first().attr('style', 'max-width: 100% !important');
         
            const isReadOnly =  JSON.parse(localStorage.getItem('isReadOnly'));
            if(isReadOnly)
              iframeDocument.find("button:contains('Submit')").prop('disabled', true);

            renderFormTab(iframeDocument, linkText);
    }
}



//#region GENERATE CHILD TAB FOR MASTER FORM
const renderFormTab = async function(iframeDocument, linkText){
   
    const tabList = iframeDocument.find("ul[role='tablist'] li").eq(1);
    if (tabList.length){

        const firstSpaceIndex = linkText.indexOf(" ");
        let listname = linkText.substring(firstSpaceIndex + 1);
        let alink = tabList.children()

        let formName = listname;
        if(listname === 'Drawing Control')
            formName = 'Drawing Checklist'
        alink.text(formName);

        const reference = iframeDocument.find("input[title='Reference']").val()
        let formItems = await getFormItems(listname, reference);
        await listMenu(iframeDocument, formItems)
    }
}

const getFormItems = async function(formName, reference){
    let listname = formName
    let urlListName = '';

     if(listname === 'Drawing Control')
        urlListName = 'CKLDWG'
    
    // debugger;
    // const tempList = await _web.getList(listname).get();
    // const listnameUrl  = tempList.Url;

    let items = await _web.lists.getByTitle(listname).items
    .select("Id,Title")
    .filter(`MPID/Title eq '${reference}'`)
    .getAll()

    return {
        items: items,
        urlListName: urlListName
    }

}

const listMenu = async function(iframeDocument, MetaInfo){
    //<h2>Master Plan Procedures</h2>
    let htmlContent = `<div class="container">
                         <div class="items">
                           <ul id="mpItem-list" style="list-style-type: square">`

    let tempUrl;
    for (const item of MetaInfo.items){
        const projectNo = item.Title;

        tempUrl = `${_webUrl}/SitePages/PlumsailForms/${MetaInfo.urlListName}/Item/EditForm.aspx`
        let id = item.Id
        const itemUrl = `${tempUrl}?item=${id}`
        htmlContent += `<li><a href="${itemUrl}" target="MPFormFrame" onclick="adjustChildTab(this)">${projectNo}</a></li>`
    }

    tempUrl = tempUrl.replace('EditForm.aspx', 'NewForm.aspx')

            htmlContent += `</ul>
                             <a style="text-align: right" href="${tempUrl}" target="MPFormFrame" onclick="adjustChildTab(this)">add New</a>
                           </div>

                            <div class="form">
                                <iframe Id="MPFormFrame" name="MPFormFrame"></iframe>
                            </div>
                        </div>`

    
    let ctlr = iframeDocument.find("#mpFormTab")
    ctlr.append(htmlContent)
    ctlr.children().css({ 
        'margin': '0px', 
        'max-width': '100%' 
    });
    await renderAURTab(iframeDocument);
}

function adjustChildTab(aLinkCtlr){
  
    let childTab = $('#mpFormTab').find('#MPFormFrame');

    childTab.on('load',  function() {
        setTimeout(async function () {
           
            let content = childTab.contents();
            //content.find('.fd-form-container .container-fluid').attr('style', 'padding-left: 95px !important');

            let formUrl = localStorage.getItem('formUrl');
            let formName = localStorage.getItem('formName');
            let ProjName = localStorage.getItem('ProjName');

            if(formUrl){
              let newLi = `<li><a href="${formUrl}" target="MPFormFrame" onclick="adjustChildTab(this)">${ProjName}</a></li>`
              $('#mpFormTab div.items ul').append(newLi);

              localStorage.removeItem('formUrl');
              localStorage.removeItem('formName');
              localStorage.removeItem('ProjName');
            }

            let gridDWG =  content.find('div.Dwg-Grid');
            gridDWG.attr('style', 'padding-left: 0px !important');
            // setTimeout(function () {
            //   content.find('.fd-grid, .container-fluid, .Dwg-Grid').attr('style', 'padding-left: 0px !important');
            // }, 1000);
        }, 1000);
    });

    
}
//#endregion


//#region RENDER AUDIT REPORT AND CORRECTIVE ACTION
const renderAURTab = async function(iframeDocument){

    let aurUrl = `${_webUrl}/SitePages/PlumsailForms/AuditReport/Item/NewForm.aspx`
    let mpID = JSON.parse(localStorage.getItem('MPID'))
    if(mpID){
        let query = `MPID/ID eq '${mpID}'`
        let items = await _web.lists.getByTitle(auditReport).items.select("Id").filter(query).get();

        if(items.length > 0)
        aurUrl = `${_webUrl}/SitePages/PlumsailForms/AuditReport/Item/EditForm.aspx?item=${items[0].Id}`
    }

    let ctlr = iframeDocument.find("#aurFormTab")

    const iframe = $('<iframe>', {
        id: 'aurId',
        src: aurUrl,
        width: '100%',
        height: '600px',
        frameborder: '0',
        scrolling: 'auto'
    }).css({
        'margin': '0px',
        'max-width': '100%'
    });
    
    ctlr.append(iframe);

    // iframe.on('load', function() {
    //         let content = $(this).contents();
    //         $(content).find('div.SPCanvas').css({ 'margin-left': '85px'});
    // });
}


//#endregion