let _layout, _web, _isSiteAdmin, _container, _hot, _data, _fieldSchema;

let _module, _formType, _webUrl, _siteUrl, _listTitle;
let _isNew = false, _isEdit = false;

let screenHeight = screen.height - 500;

var onRender = async function (relativeLayoutPath, moduleName, formType){
   _layout = relativeLayoutPath

    await extractValues();
   _data = await getData();
   _fieldSchema = [
                    {  "title": "Filename", // display name
                    "data":"FullFileName", // internal name
                    "type": "text",
                    "width": "20%",
                    "allowEmpty": false,
                    "className": "toUpper",
                    "length": 60
                    },
                    
                    {  "title": "Title",
                    "data":"Title",
                    "type": "text",
                    "width": "25%",
                    "allowEmpty": false,
                    "length": 254
                    },
                    
                    {  "title": "Deliverable Type",
                    "data":"DeliverableType",
                    "type": "dropdown",
                    "source": ["DOC", "DWG", "TRM"],
                    "width": "10%",
                    "allowEmpty": false
                    }
                  ]
  _spComponentLoader.loadScript(_htLibraryUrl).then(renderHandsonTable)
}

var extractValues = async function(moduleName, formType){
    _web = pnp.sp.web;
    _isSiteAdmin = _spPageContextInfo.isSiteAdmin;
    _module = moduleName;
    _formType = formType;
    _webUrl = _spPageContextInfo.siteAbsoluteUrl;
    _siteUrl = new URL(_webUrl).origin;

    _htLibraryUrl = _layout + '/controls/handsonTable/libs/handsontable.full.min.js';

    if(_formType === 'New')
        _isNew = true;
    
    else if(_formType === 'Edit'){
        _isEdit = true;
        //_itemId = fd.itemId;
    }

    const listUrl = fd.webUrl + fd.listUrl;
    const list = await _web.getList(listUrl).get();
    _listTitle = list.Title;
}

const renderHandsonTable = (Handsontable) => {
  _container = document.getElementById('dt');

  _hot = new Handsontable(_container, {
          data: _data,
          columns: _fieldSchema,
          width:'99%',
          height: screenHeight,
          search: {
            searchResultClass: 'highlight-cell'
          },
          rowHeaders: true,
          colHeaders: true,
          manualColumnResize: true,
          manualRowMove: false,
          stretchH: 'all',
          licenseKey: htLicenseKey
   });
}

var getData = async function(){
    var _itemArray = [];
  
    var _query = 'Title ne null';
    let columns = ['col1','col2','col3'];

    await _web.lists.getByTitle(_listTitle).items.filter(_query)
    .select(columns)
    .getAll().then(async function(items){
        let itemCount = items.length;
        if (itemCount > 0) {
            for(var i = 0; i < itemCount; i++){

                var item = items[i];
                var rowData  = {};

                for(var j = 0; j < columns.length; j++){
                        var _colname = columns[j];
                        var _value = item[_colname];
                        rowData[_colname] = _value;
                }
                _itemArray.push(rowData);
            }
        }
    });
   return _itemArray;
}

//subtract two numbers












