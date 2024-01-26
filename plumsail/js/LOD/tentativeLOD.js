var _layout = "/_layouts/15/PCW/General/EForms";
var errorImg = _layout + '/Images/Error.png';
var submitImg = _layout + '/Images/Submitted.png';

var webURL = '', formType = '', _list = '', _itemId = '', _lodRef = '', _status = '', tblRows, 
    RLODList = 'RLOD', tempLOD = 'tempLOD', RegDeli = 'Register Deliverables';
var haveError = false, _isNew = false;
var _colsInternal = [],  _colsType = [];

$(document).ready(async function() {
    var _lodArray = [];
    var _colsArray = [];

    var _colsDisplay = [], _colsWidth = [], _colsReq = [], _colsUrl = [];
    var _listname, _filterfld;

    //const params = new URLSearchParams(window.location.search);
    formType = await GetParameterValues('formtype');
    _list = await GetParameterValues('List');
    _itemId = await GetParameterValues('ID');

    webURL = document.URL.substring(0, document.URL.indexOf('/_layout'));

    var _web = $pnp.sp.web;
    var _lodArray = await getMajorType(_web, webURL, formType);
    if(_lodArray.length > 0){
        var _lodVal = _lodArray[0];
        _colsDisplay = formatStringToArray(_lodVal.display);
        _colsInternal = formatStringToArray(_lodVal.internal);
        _colsType = formatStringToArray(_lodVal.type);
        _colsWidth = formatStringToArray(_lodVal.width);
        _colsReq = formatStringToArray(_lodVal.required);
        _colsUrl = formatStringToArray(_lodVal.url);
        _listname = _lodVal.listname;
        _filterfld = _lodVal.filterfld;
    }
    
    //GET CURRENT LISTNAME TO GET REFERENCE TO ECTRACT ITEMS "CDS-DWG-00003"
    // var _listTitle = await _web.lists.getById(_ListGUI).get().then(async function(res) {
    //          return res.Title;
    //     });

    //GET REFERENCE TO ECTRACT ITEMS "CDS-DWG-00003"
    if(_itemId == "" || _itemId == null || _itemId == undefined)
      _isNew = true;
    else _isNew = false;

    
    var listItem = "", _query = "";
    if(!_isNew){
     _query = "ID eq " + _itemId;
    listItem = _web.lists.getByTitle(_list).items.filter(_query);
    
    _lodRef = await listItem.select("Title").getAll()
        .then(async function(items) {
             return items[0].Title;
        });

     _status = await listItem.select("Status").getAll()
     .then(async function(items) {
          return items[0].Status;
      });
    }
        
    var _colsArray = await getFieldTypes(_colsDisplay, _colsType, _colsWidth, _colsUrl);
    var _dataArray = await getData(_web, _listname, _filterfld, _lodRef , _lodArray[0].internal, _colsInternal, _colsType);
    
// Set your JSS license key (The following key only works for one day)
jspreadsheet.setLicense('MjI5OWY1YjkxZDg4NDQwYzUzYjMyNjVjZmMxMmY4Mjc1YjY5MTY1NDhkZjBmM2U1NmU0YjQ4YzVmMTk5OTI5ZWRlZWM0ZDNmZWNjMTVhOTIyZGVhNDIxMDc3NmJhZDg3ZjAyOGE0M2QxODgzNzdmN2NjMGJjNmZiMGRlZjNkMTAsZXlKdVlXMWxJam9pUVd4cElGTnNaV2x0WVc0aUxDSmtZWFJsSWpveE5qYzRPRE00TkRBd0xDSmtiMjFoYVc0aU9sc2laR0Z5WW1WcGNuVjBMbU52YlNJc0lteHZZMkZzYUc5emRDSmRMQ0p3YkdGdUlqb3dMQ0p6WTI5d1pTSTZXeUoyTnlJc0luWTRJaXdpZGpraVhYMD0=');
//var _sheetCols = _cols;
//_sheetCols.push("_MesgError","_Status")
// Create the spreadsheet
//_colsArray = _colsArray.replaceAll('"clockEditor"', 'clockEditor')

// _colsArray = _colsArray.map(function(v) {
//     debugger;
//     var val = v;
//     return eval('(' + v + ')');
//   });

var table = jspreadsheet(document.getElementById('spreadsheet'), {
    
    worksheets: [{
        minDimensions: [_colsInternal.length,15],
        data: _dataArray,
        columns: _colsArray,
        csvHeaders: true,
        freezeColumns: 1,
        search: true,
        tableOverflow:true,
        lazyLoading:true,
        loadingSpin:true,
        tableWidth:'1885px',
        tableHeight:'500px',
        filters: true,
        // pagination: 10,
        //paginationOptions: [10,25,50,100],
        columnDrag: false,
        allowManualInsertColumn: false,
        allowInsertColumn: false,
        allowRenameColumn: false,
        allowDeleteColumn: false,
    }],

    // onchange: function(event,a,col,row,after,before) {
    //     debugger;
    //     var prevRow = event.table.childNodes[2];
    //     var currentRow = prevRow;
    //     prevRow = prevRow.childNodes[row-1];
    //     var filename = prevRow.childNodes[1].innerHTML;
 
    //     if( filename == ""){
    //         //prevRow.childNodes[3].innerHTML = "<p style='color:red;'>empty row is not allowed</p>";
    //         var pRow = parseInt(row);
    //         alert('empty row is not allowed. Row ' + pRow + ' is empty');
    //         currentRow = currentRow.childNodes[row];
    //         var colNum = parseInt(col+1);
    //         currentRow.childNodes[colNum].innerHTML = "";
    //         return;
    //     }
    //     console.log(event,a,col,row,after,before);
    // }
});

document.addEventListener('keydown', function(e) {
    var _type = e.target.className;
    if ( _type.includes('_input')) {
        if (e.target.innerText.length > 252) {
            alert('maximum 255 characters are allowed to be set in _FileName or _Title at row ' + (e.target.y +1));
            return false;
        }
    }
});

$("div").on('scroll', function (e) {
    var div = $(this);
    if (div[0].scrollHeight - div.scrollTop() == div.height())
      table[0].insertRow(10);
});

if(!_isNew){
 setStyle(table, _status);
 $('div.jss_search_container').find('div:eq(1)').append("<br/><div style='padding-top: 10px;'><label id='lodid'>Reference:  " + _lodRef + "</label></div>");
}

preloader(_web, _colsInternal, _colsType, table);
});

var getFieldTypes = async function(_colsDisplay, _colsType, _colsWidth, _colsUrl){
    var _itemArray = {};
    var _colArray = [];
    //var _clockArray = clockEditor;


    for (var i = 0; i < _colsDisplay.length; i++) {          
            var _display, _type, _width;

            _display = _colsDisplay[i];
            _type = _colsType[i].toLowerCase();
            _width = _colsWidth[i];

          if( _type == "clockEditor")
             _itemArray.type = clockEditor;
           else if(_type == "calendar")
             _itemArray.options = dateOption();
            else if(_type == "autocomplete")
             _itemArray.url = _colsUrl[0]; //"json/workflow.txt";//['Assigned','Open'];
            else _itemArray.type = _type;


            if( _type == "text"){
                _itemArray.max = '10'
                _itemArray.maxlength = '10';
            }
          
            //_itemArray.filter = true;
            _itemArray.title = _display;
            _itemArray.width = _width;
            if(_display === '_Mesg' || _display === '_Status' || _display === '_Type')
             _itemArray.readOnly = true;
            _colArray.push(_itemArray);
            _itemArray = {};
    }
    return _colArray;
}

var getData = async function(_web, _listname, _filterfld, _refNo, _cols, _colsInternalArray, _colsType){
    var _itemArray = [];
    var _dataArray = [];

    var _query = _filterfld + " eq '" + _refNo + "'";

    //_cols = _cols.replace(",Mesg", "").replace(",Status", "");

    await _web.lists.getByTitle(_listname).items.filter(_query)
    .select(_cols)
    //.expand(_expandColumns)
    .getAll().then(async function(items) {
      _itemCount = items.length;
      if (_itemCount > 0) {
        for(var i = 0; i < _itemCount; i++){
           var item = items[i];

           for(var j = 0; j < _colsInternalArray.length; j++){
            var _type = _colsType[j];
            var _colname = _colsInternalArray[j];
            var _value = item[_colname];
            if( _type.toLowerCase() == "calendar"){
                //_itemArray.push(formatDate(_value));
                _itemArray.push(_value);
            }
            else if( _type == "clockEditor"){
              _itemArray.push(formatTime(_value));
            }
            else if(_colname === 'Status'){
                var iconUrl = _value;
                _value  = "<img src='" + iconUrl + "' alt='na'></img>";
                _itemArray.push(_value);
            }
            else _itemArray.push(_value);
           }

           _dataArray.push(_itemArray);
           _itemArray = [];
        }
      }
  });
  return _dataArray;
}

var validateInput = async function(_web, table, _data, allData){
    haveError = false;
    var isErrorFound = checkEmptyRows(table, allData);
    if(isErrorFound)
       return;

    var xmlhttp = new XMLHttpRequest();

    xmlhttp.open('POST', 'http://db-sp.darbeirut.com/SoapServices/JSpreadSheetService.asmx?op=CHECK_FILENAMES', false);
    var sr = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
               '<soap:Body>' +
                 '<CHECK_FILENAMES xmlns="http://tempuri.org/">' +
                    '<WebURL>' + webURL + '</WebURL>' +
                      '<SpreadSheet>';

                var valArray = table[0].getColumnData(0);
                for(var i =0; i < valArray.length; i++){
                    if(valArray[i] != "")
                     sr += '<string>' + valArray[i] + '</string>';
                }
                sr += '</SpreadSheet>';

                table[0].setColumnData(3,'',true);
                table[0].setColumnData(4,'',true);
            

            sr += '</CHECK_FILENAMES>' +
                     '</soap:Body>' +
                        '</soap:Envelope>';

      xmlhttp.onreadystatechange = async function () {
        if (xmlhttp.readyState == 4) {
            if (xmlhttp.status == 200) {
                var xmlDoc = $.parseXML( xmlhttp.responseText ),
                $xml = $( xmlDoc ),
                $value= $xml.find( "CHECK_FILENAMESResult" );
                var text = $value.text();
                tblRows = _data.length;
                const obj =  await JSON.parse(text, async function (key, value) {
                    var baseType = '';
                    for(var i = 0; i < tblRows; i++){
                        var _filename = table[0].getCellFromCoords(0,i).innerHTML;
                        var _Mesg = "";
                        
                        var columIndex = i + 1;

                        if(key.includes("|")){
                            var keyArray = key.split("|");
                            key = keyArray[0];
                            baseType = keyArray[1];
                        }

                        if(_filename == key){
                            if(value != "")
                              _Mesg = value;
                            else _Mesg  = table[0].getCellFromCoords(3,i).innerHTML;

                            if(_Mesg === "" || _Mesg === undefined || _Mesg === null){
                                // debugger;
                                // var isExist = await checkRLODListItem(_web, RLODList, _filename);
                                // if(isExist.length > 0){
                                //     setMessage(table, i, columIndex, 'filename is already exist in ' + RLODList, errorImg, baseType);
                                //     haveError = true;
                                // }
                            
                               // else{
                                    var rowNum = i + 1;
                                    table[0].resetStyle('A' + rowNum);
                                    table[0].setReadOnly('A' + rowNum, true);
                                    table[0].setReadOnly('B' + rowNum, true);
                                    table[0].updateCell(2, i, baseType, true);
                                    table[0].updateCell(4, i, submitImg, true);
                                //}
                            } 
                            else{
                                setMessage(table, i, columIndex, value, errorImg, baseType);
                                haveError = true;
                            }
                            break;
                         }
                     }
                });
        }
    }
}
    // Send the POST request
    xmlhttp.setRequestHeader('Content-Type', 'text/xml');
    xmlhttp.send(sr);
    // send request
}

var submitData= async function(_web, table, _data, _lodRef){
    var xmlhttp = new XMLHttpRequest();

    xmlhttp.open('POST', 'http://db-sp.darbeirut.com/SoapServices/JSpreadSheetService.asmx?op=SUBMIT_FILENAMES', true);
    var sr = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
               '<soap:Body>' +
                 '<SUBMIT_FILENAMES xmlns="http://tempuri.org/">' +
                    '<WebURL>' + webURL + '</WebURL>' +  
                    '<LODRef>' + _lodRef + '</LODRef>' +
                    '<formType>LOD</formType>';
                 
                sr += '<RowValues>';
                for(var i =0; i < _data.length; i++){
                    sr += '<ArrayOfString>';
                    var rowData = table[0].getRowData(i);
                    for(var j =0; j < rowData.length; j++){
                        var col = rowData[j];
                        sr += '<string>' + col + '</string>';
                    }
                    sr += '</ArrayOfString>';
                }
                sr += '</RowValues>';

                var columns = _colsInternal + '|' + _colsType;
                sr += '<Columns>' + columns + '</Columns>';

        sr += '</SUBMIT_FILENAMES>' +
                '</soap:Body>' +
                  '</soap:Envelope>';

     xmlhttp.onreadystatechange = async function () {
        if (xmlhttp.readyState == 4) {
            if (xmlhttp.status == 200) {
                if(_isNew){
                    var _query = "Title eq '" + _lodRef + "'";
                    var listItem = _web.lists.getByTitle(RegDeli).items.filter(_query);
                    _itemId = await listItem.select("ID").getAll()
                                    .then(async function(items) {
                                        return items[0].ID;
                                    });
                }
                if(_itemId != "" && _itemId != null && _itemId != undefined){
                    _web.lists.getByTitle(RegDeli).items.getById(_itemId).update({ Status: 'Submitted' });
                    closeForm(webURL, true);
                }
            }
        }
    }
    // Send the POST request
    xmlhttp.setRequestHeader('Content-Type', 'text/xml');
    xmlhttp.send(sr);
    // send request
}

var preloader = async function(_web, _colsInternal, _colsType, table){
    $('input[type=button]').click(async function () {

        var btnText = $(this).attr("value");
        await preloader_btn(btnText);
        $("input").attr('disabled', true).css("background-color", "");

       if(btnText === 'Close'){
         closeForm(webURL);
         return;
       }
       else if(btnText === 'Click Me'){
        var loader = setInterval(function(){
            if($('#loader').length > 0){
                var divCheckingInterval = setInterval(async function(){
                if($('div.dimbackground-curtain').length > 0){
                    Remove_Pre();
                    if(haveError)
                      $('#subid').attr('disabled', true).removeClass("button1").addClass("disableButton");
                    else $('#subid').attr('disabled', false).removeClass("disableButton").addClass("button1");
                    clearInterval(divCheckingInterval);
                }
               }, 500);
            clearInterval(loader);
            }
        }, 500);  
        return;
       }
       
       var _data = table[0].getData(false,false);
       var allData = _data;
       _data = filterData(_data);
       var lodCtl = $('#lodid').text();

       if(_isNew && lodCtl === '' && _data.length > 0){
         _lodRef = 'LOD-' + await getCounter(_web, "LOD");
         $('div.jss_search_container').find('div:eq(1)').append("<br/><div style='padding-top: 10px;'><label id='lodid'>Reference:  " + _lodRef + "</label></div>");
       }

        var loader = setInterval(function(){
            if($('#loader').length > 0){
                var divCheckingInterval = setInterval(async function(){
                if($('div.dimbackground-curtain').length > 0){
                    Remove_Pre();
                    if(haveError)
                      $('#subid').attr('disabled', true).removeClass("button1").addClass("disableButton");
                    else $('#subid').attr('disabled', false).removeClass("disableButton").addClass("button1");
                    clearInterval(divCheckingInterval);
                }
               }, 500);
            clearInterval(loader);
            }
        }, 500);  


        if(_data.length > 0){
           if(btnText === 'Validate'){
            debugger;
            haveError = false;
            for(var i = 0; i < _data.length; i++){
               var row = _data[i];
               var columIndex = i + 1;
               var result = await checkRLODListItem(_web, RLODList, row[0]);
               if(result.length > 0){
                   setMessage(table, i, columIndex, 'filename is already exist in ' + RLODList, errorImg, result[0].DeliverableType);
                   haveError = true;
               }
            }

            if(!haveError){
                await validateInput(_web, table, _data, allData);
                isValidate = true;
            }
                _data = table[0].getData(false,false);
                _data = filterData(_data);
                await createList(_web, tempLOD, _colsInternal, _colsType, table, _data, _lodRef, btnText, RegDeli);        
            
           }
           else submitData(_web, table, _data, _lodRef);
         }
    });
}

var preloader_btn = async function(textVal){
	try{
        var inputId = '';
		var webUrl = window.location.protocol + "//" + window.location.host;
        const html = $('input[type=button]').length;

        if($('#loader').length > 0)
           $('#loader').remove();

        if(textVal == 'Validate')
          inputId = '#valid';
        else if(textVal == 'Submit')
          inputId = '#subid';
        else inputId = '#cloid';


			 //$('input[type=button]').dimBackground({
            $(inputId).dimBackground({
						 darkness: 0.1
					 }, async function() {
                        if($('#loader').length > 0)
                          $('#loader').remove();

						$("<img id='loader' src='" + webUrl + _layout + "/Images/Loading.gif' />")
						  .css({
								"position": "absolute",
								"top": "400px",
								"left": "828px",
								"width": "100px",
								"height": "100px",
								"visibility":"hidden"
							}).insertAfter('body');
						 $('#loader').css("visibility", "visible").dimBackground({ darkness: 0.8 });
				});        
	}
   catch(err) { console.log(err.message); }
}

function Remove_Pre(){
    var bgDim = setInterval(function(){
        if($('div.dimbackground-curtain').length > 0){
            $('div.dimbackground-curtain').remove();
            clearInterval(bgDim);
        }
    }, 500);

    var loader = setInterval(function(){
        if($('#loader').length > 0){
            $('#loader').remove();
            clearInterval(loader);
        }
    }, 500);

    $("input").attr('disabled', false); //.hover(function(){$(this).css("background-color", "#1b1b1b");});
}

function setStyle(table, _status){
    var _data = table[0].getData(false,false);
    _data = filterData(_data);
    var isAllCoeect = true;
    for(var i =0; i < _data.length; i++){
        var rowNum = i + 1;
        var rowData = table[0].getRowData(i);

        for(var j =0; j < rowData.length; j++){
            var headerMesg = table[0].getHeader(j);
            if(headerMesg == '_Mesg'){
                var colVal = rowData[j];
                var isReadOnly = false;
                if(colVal != ""){
                    table[0].setStyle('D' + rowNum, 'color', 'red');
                    isAllCoeect = false;
                }
                else isReadOnly = true;

                table[0].setReadOnly('A' + rowNum, isReadOnly);
                table[0].setReadOnly('B' + rowNum, isReadOnly);
            }
        }
    }

if(isAllCoeect){
   $('#valid').attr('disabled', true).removeClass("button1").addClass("disableButton");
   $('#subid').attr('disabled', false).removeClass("disableButton").addClass("button1");
}

  if(_status == 'Submitted'){
    $('#valid').remove();
    $('#subid').remove();
  }
}

function checkEmptyRows(table, _data){
    var isEmptyRow = false, isErrorFound = false;
    var mesg = "";
    for(var i =0; i < _data.length; i++){
        var rowData = table[0].getRowData(i);
        var rowNum = i + 1;

        var filename = rowData[0];
        var title = rowData[1];

          if(filename == "") {
            isEmptyRow = true;
            mesg += 'row ' + rowNum + ' is empty <br/>';
          }
          else if(filename != "" && title == ""){
            isEmptyRow = true;
            mesg = 'Title is required field';
          }

          if(isEmptyRow && filename != ""){
            isErrorFound = true;
            table[0].updateCell(3, i, 'empty rows are not allowed<br/>' + mesg, true);
            table[0].updateCell(4, i, errorImg, true);
            table[0].resetStyle('D' + rowNum);
            table[0].setStyle('D' + rowNum, 'color', 'red');
            isEmptyRow = false;
            mesg = "";
          }
    }
    return isErrorFound;
}

function setMessage(table, i, columnIndx, value, imgStatus, baseType){
    table[0].updateCell(2, i, baseType, true);
    table[0].updateCell(3, i, value, true);
    table[0].updateCell(4, i, imgStatus, true);
    table[0].resetStyle('D' + columnIndx);
    table[0].setStyle('D' + columnIndx, 'color', 'red');
    table[0].setReadOnly('A' + columnIndx, false);
    table[0].setReadOnly('B' + columnIndx, false);
}

function clickFunc(){
    debugger;
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
         if (this.readyState == 4 && this.status == 200) {
             alert(this.responseText);
         }
    };
    xhttp.open("GET", "http://db-sp.darbeirut.com:8080/api/v1/hw", true);

    //xhttp.open("GET", "https://jsonplaceholder.typicode.com/todos/1", true);
    //xhttp.open("GET", "http://db-sp16.darbeirut.com/api/v1/hw", true);
    xhttp.setRequestHeader("Content-type", "application/json");
    xhttp.send("Your JSON Data Here");
}