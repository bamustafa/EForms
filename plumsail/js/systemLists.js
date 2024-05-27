var _layout, _web, _webUrl, _module = '', _formType = '',_list;
let _isNew = false, _isEdit = false, _isSiteAdmin = false;

var onRender = async function (relativeLayoutPath, moduleName, formType){
    _layout = relativeLayoutPath;
    await loadScripts();
    await getLODGlobalParameters(moduleName, formType);
    await handleSystemList();
}

var loadScripts = async function(){
    const libraryUrls = [
        _layout + '/plumsail/js/commonUtils.js',
        _layout + '/plumsail/js/customMessages.js'
      ];

    const cacheBusting = `?v=${Date.now()}`;
      libraryUrls.map(url => { 
          $('head').append(`<script src="${url}${cacheBusting}" async></script>`); 
        });
        
    const stylesheetUrls = [
        _layout + '/plumsail/css/CssStyle.css'
    ];
  
    stylesheetUrls.map((item) => {
      var stylesheet = item;
      $('head').append(`<link rel="stylesheet" type="text/css" href="${stylesheet}">`);
    });
}

var getLODGlobalParameters = async function(moduleName, formType){
    _web = pnp.sp.web;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    
    if(_formType === 'New'){
        fd.clear();
        _isNew = true;
    }
    else if(_formType === 'Edit')
        _isEdit = true;

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _list = list.Title;
}

var handleSystemList = async function(){
    if(!_isSiteAdmin){
        fd.toolbar.buttons[0].style = "display: none;";
        //fd.toolbar.buttons[1].style = "display: none;";

        if(_list === 'Counter'){
            fd.field('Title').disabled = true;
            fd.field('Counter').disabled = true;
        }
    }
}