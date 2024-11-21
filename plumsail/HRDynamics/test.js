

var _module, _formType, _web, _webUrl, _siteUrl, _layout, _itemId, _ImageUrl, _ListInternalName, _ProjectNumber, 
    _ListFullUrl, _CurrentUser, _fields = [];
let Inputelems = document.querySelectorAll('input[type="text"]');
let _Email, _Notification = '', _htLibraryUrl, _errorImg, _submitImg;


//handson variables
let _hot, _container, _data= [], _colArray, batchSize = 15;
var onRender = async function (moduleName, formType, relativeLayoutPath){ 

    await getGlobalParameters(relativeLayoutPath, moduleName, formType);
    //_colArray = await getSchema()
    //_spComponentLoader.loadScript(_htLibraryUrl).then(_setData);

    await dynamicPageSample();
}

let dynamicPageSample = async function(){
  let tempUrl = 'https://ax.d365s.dar.global/namespaces/AXSF/?cmp=BEI&mi=ProjManagementWorkspace';
  let htmlContent = `<div class="form">
                       <iframe Id="formFrame" name="formFrame" src="${tempUrl}" width="600" height="400">
                       </iframe>
                     </div>`
   $('#dt').append(htmlContent)

   $('#formFrame').on('load', function() {
    setTimeout(function () {
        let content = $('#formFrame').contents();
        //adjustMPFormStyles(content, linkText, mpIdValue);
        alert('hello iframe load');
    }, 1000);
   });
}

var getSchema = async function(){
    var colArray = [];

    var fetchUrl = `${_webUrl}/Config/TR-Schema.txt`
    await fetch(fetchUrl)
        .then(response => response.text())
        .then(async data => {
            colArray = JSON.parse(data); 
    });

    return  colArray
}

const _setData = (Handsontable) => {

    // if(_data.length < batchSize){
    //     var remainingLength = batchSize - _data.length;
    //     for (var i = 0; i < remainingLength; i++) {
    //       var rowData = { id: i + 1, value: 'Row ' + (i + 1) }
    //       _data.push(rowData);
    //     }
    // }

    var contextMenu = ['row_below']; //, '---------', 'remove_row'];
    console.log(_colArray)
    debugger;
     _container = document.getElementById('dt2');

    _hot = new Handsontable(_container, {
        data: [
          ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1'],
          ['A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2'],
          ['A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3'],
        ],
        colHeaders: [
          'ID',
          'Full name',
          'Position',
          'Country',
          'City',
          'Address',
          'Zip code',
          'Mobile',
          'E-mail',
        ],
        rowHeaders: true,
        height: 'auto',
    
        licenseKey: 'e9cca-b5ee4-f18a2-5492c-9fd49'
      });

    setTimeout(() => {
        _hot.render();
    }, 200);
}

var getGlobalParameters = async function(relativeLayoutPath, moduleName, formType){

    fd.toolbar.buttons[0].style = "display: none;";
    fd.toolbar.buttons[1].style = "display: none;";   

    if($('.text-muted').length > 0)
      $('.text-muted').remove();
    
    _module = moduleName;
    _formType = formType;
    _web = pnp.sp.web;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl; //https://db-sp.darbeirut.com/sites/QM-eAudit
    _siteUrl = new URL(_webUrl).origin;
    _layout = relativeLayoutPath;
    _itemId = fd.itemId;

    _ImageUrl = _spPageContextInfo.webAbsoluteUrl + '/Style%20Library/tooltip.png',
    _ListInternalName = _spPageContextInfo.serverRequestPath.split('/')[5],
    _ProjectNumber = _spPageContextInfo.serverRequestPath.split('/')[2],
    _ListFullUrl = _spPageContextInfo.webAbsoluteUrl + '/Lists/' + _ListInternalName,

    _errorImg = _layout + '/Images/Error.png';
    _submitImg = _layout + '/Images/Submitted.png';
    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';
}