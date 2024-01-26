var partContainer;

var _itemCount = 0;
var _isFirstRow = false, _isbyRef = false, _haveOpenItems = false;
var _module, _listname;

var _attachFiles = "";
var rejectedStatus = '', rejectedRemarks = '';

var hTable;
var _heigth = '100';
var _cols = [], _partData = [];
var _colsInternal = [], _colsType = [], _colsht = [];
var linkIndex = 0, tradeIndex = 0, roleIndex = 0, statusIndex = 0, remarkIndex = 0, AttachmentCount = 0;
var setTimeOut = false;

var onPageRender = async function(module, listname, isbyRef){
     partContainer = document.getElementById("parttasks");
    _module = module;
    _listname = listname;
    _isbyRef = isbyRef;
    setFormHeaderTitle();
    //const numRows = hTable.countRows();
    
    if(hTable === undefined){
       getHandsonTableSchema();
        _cols.map(item =>{
            _colsInternal.push(item.InternalName);
            _colsType.push(item.type);
            _colsht.push(item.data);
       });  
    
       await getInfo();
       setTimeOut = true;
    }
    _spComponentLoader.loadScript(_htLibraryUrl).then(setHandsonTable);
    
    addLegend('insplbl', InspectorLabelMesg, partContainer);

    if(_IsLeadAction){
        const linkElement = $('.nav-link.active');
        if(linkElement.length > 0){
            var isVisited = false;
            linkElement.click(function(event) {
                debugger;
                if(!isVisited){
                    event.preventDefault(); // Prevent the default link behavior
                    onPageRender(module, listname, isbyRef);
                    isVisited = true;
                }
          });
        }
    }
}

var getInfo = async function(){

    var siteURL = window.location.protocol + "//" + window.location.host;
    var webUrl = siteURL + _spPageContextInfo.siteServerRelativeUrl;
    
    var fileUrl = webUrl + _layout + '/Images/Icons/'
    var fileIcons = {
      pdf: fileUrl + 'Pdf.png',
      doc: fileUrl + 'Word.png',
      docx: fileUrl + 'Word.png',
      xls: fileUrl + 'Excel.gif',
      xlsx: fileUrl + 'Excel.gif',
      ppt: fileUrl + 'Ppt.png',
      pptx: fileUrl + 'Ppt.png',
      png: fileUrl + 'Png.png',
      gif: fileUrl + 'Png.png',
      jpeg: fileUrl + 'Jpeg.gif',
      jpg: fileUrl + 'Jpeg.gif',
      rar: fileUrl + 'Zip.gif',
      zip: fileUrl + 'Zip.gif',
      attach: fileUrl + 'attach.gif'
      // Add more file extensions and corresponding icons here
    };

    let web = pnp.sp.web;//new Web(webUrl);
    //let web = new Web(webUrl);
    var _Ref = fd.field("Reference").value;
    var _query, _spColumns, _expandColumns;

    const _cols = _colsInternal.join(',');
    _expandColumns = ["AssignedTo", "AttachmentFiles"];
    
    if(_module === 'SLF'){
        _query = "Reference eq '" + _Ref + "' and Role eq 'Inspectors' ";
    }
    else
    {
        if(_isbyRef){
            if(_isTeamLeader && !_isPMCAssignment) 
            _query = "Reference eq '" + _Ref + "' and isPMCAssignment eq '0' and Trade ne '" + _partTrade + "'";
            else _query = "Reference eq '" + _Ref + "'";
        }
        else{
            var _isPMCAssignment = fd.field("isPMCAssignment").value;
            
            if(_role === '' && _IsLeadAction)
              _query = "Reference eq '" + _Ref + "' and Trade ne 'PMC'";

            else  if( (_role === '' && !_IsLeadAction)) 
             _query = "Reference eq '" + _Ref + "' and isPMCAssignment eq '1' and IsLeadAction eq '0' and Trade ne 'PMC'";

            else if( (_isTeamLeader || _role === 'Inspectors') && _isPMCAssignment)
             _query = "Reference eq '" + _Ref + "' and isPMCAssignment eq '1' and Trade ne '" + _partTrade + "' and Trade ne 'PMC'";

            else  if(_isTeamLeader && !_isPMCAssignment)
             _query = "Reference eq '" + _Ref + "' and (isPMCAssignment eq '0' or Trade eq 'PMC') and Trade ne '" + _partTrade + "'";
       
        }
    }
        
	 await web.lists.getByTitle(_listname).items.filter(_query)
          .select('Id,' + _cols)
          .expand(_expandColumns)
          .getAll().then(function(items) {
            _itemCount = items.length;
            if (_itemCount > 0) {

               for(var i = 0; i < _itemCount; i++){
                  var item = items[i];
                  var rowData  = {};
                  var _colStatus = '';
                  for(var j = 0; j < _colsInternal.length; j++){
                    
                    var _type = _colsType[j];
                    var _colname = _colsInternal[j];
                    var htColumn = _colsht[j];
                    var _value;
                    
                    if(htColumn === 'checkitem'){
                        _value = item[_colname];

                        if(_value !== null)
                            _value = true;
                         else _value = false;
                    }
                    else if(htColumn === 'view item')
                    {
                        var formItemUrl = webUrl + '/SitePages/PlumsailForms/' + _listname.replace(' ','') + '/Item/EditForm.aspx?item=' + item.Id;
                       _value = "<a target='_blank' href='" +  formItemUrl + "'>link</a>";
                    }
                    else if(htColumn === 'assignedto'){
                        _value = item['AssignedTo'];
                        var _user = '';
                        var userCount = _value.length;
                        if(userCount === 3)_heigth = '130'
                        else if(userCount === 4)_heigth = '160'
                        else if (userCount > 4) _heigth = '200';
                        // else if(userCount > )
                        //  _heigth = '130'
                        for(var k =0; k < userCount; k++)
                        {
                            _user += _value[k].Title;
                            if(userCount > 0)
                             _user += "<br />";
                        }
                        _value = _user;
                    }
                    else if(htColumn === 'status'){
                        _value = item[_colname];
                        _colStatus = _value;
                        if(_value === 'Open'){
                            _haveOpenItems = true;
                        }
                    }

                    else if(htColumn === 'attachment'){
                        _value = '';
                        var attachments = item.AttachmentFiles;
                        //var tblHeight = '20%';
                        AttachmentCount = attachments.length;
                        if (AttachmentCount > 0) {
        
                          //  if(attachments.length === 1)
                          //     tblHeight = '61%';
                          
                          _value = "<table align='center'>";
                          for (var k = 0; k < attachments.length; k++) {
                            var attachment = attachments[k];
                            var attachmentName = attachment.FileName;
                            var fileExtension = attachmentName.split('.').pop().toLowerCase();
                            var iconUrl = fileIcons[fileExtension] || fileUrl + 'icon-default.png';
        
                            var attachmentUrl = attachment.ServerRelativeUrl;
                            //var thumbnailUrl = attachmentUrl + "?&thumbnailMode=1";
                            _value += "<tr>" +
                                        "<td style='border: none'><img src='" + iconUrl + "' alt='na'></img></td>  " +
                                        "<td style='border: none'><a target='_blank' href='" +  attachmentUrl + "'>" + attachment.FileName + "</a></td>" +
                                      "</tr>";
                            //_value += "<a target='_blank' href='" +  attachmentUrl + "'><img src='" + thumbnailUrl + "' alt='Attachment " + (k + 1) + "' width='100'></img></a><br/>";
                          }
                          _value += "</table>";
                        }
        
                        if(_isPart && Status === 'Completed'){}
                        else if(_isPart && _colStatus !== 'Closed' && !_isTeamLeader && !_isPMC){
                          var itemId = item.Id;
                          var link = webURL + 'SitePages/PlumsailForms/' + _listname.replace('Snag Items','SLFItems') + '/Item/EditForm.aspx?item=' + itemId;
                          //_value = "<a target='_blank' href='" +  link + "'>upload files</a>";
                          var iconUrl = fileIcons['attach'];
                          //_value += "<table align='left' height='" + tblHeight + "'>" +
                          _value += "<table align='left'>" +
                                      "<tr>" +
                                        "<td style='border: none; padding-top: 5px; vertical-align: bottom'> " + 
                                            "<img src='" + iconUrl + "' alt='na'></img>  " + 
                                            "<a href='#' onclick=\"openSmallWindow('" + link + "')\">Attach files</a>" + 
                                          "</td>" +
                                      "</tr>" +
                                    "</table>";
                        }
                        rowData[htColumn] = _value;
                    }


                    else _value = item[_colname];

                    if (_value === null)
                     _value = '';
                    rowData[htColumn] = _value;
                  }
                  _partData.push(rowData);
               }
            }
        });
}

function setHTMLContent(siteURL, _item, _cols){
    var _htmlRow;
    var _values = [_item.AssignedTo, _item.Trade, _item.Status, _item.Code, _item.Comment, _item.AttachmentFiles];

    _htmlRow = "<tr>";
    for(var i=0; i< _cols.length; i++){
         var _colname = _cols[i];
         var _val = _values[i];

        if(_val == null) _val = "";

         if(_colname == "Reference"){
            if(!_isFirstRow){
                _htmlRow += "<td rowspan='" + _itemCount + "'>" + _val + "</td>";
                _isFirstRow = true;
            }
         }

         else if(_colname == "Assigned To"){
            var  _users = "";
            for(var j =0; j < _val.length; j++)
            {
              _users += _val[j].Title + "<br />";
            }
            _htmlRow += "<td>" + _users + "</td>";
         }

         else if(_colname == "Attachments"){
            var  _fileUrls = "";
            for(var j =0; j < _val.length; j++)
            {
                var filename = _val[j].FileName;
                var filePath = siteURL + _val[j].ServerRelativeUrl;
                var _href = '<a href="' + filePath + '" target="_blank">' + filename + '</a>';
                _fileUrls +=  _href + "<br />";
            }
            _htmlRow += "<td>" + _fileUrls + "</td>";
        }

         else _htmlRow += "<td>" + _val + "</td>";
    }

    _htmlRow += "</tr>";
    return _htmlRow;
}

const setHandsonTable = (Handsontable) => {
    if(_partData.length === 1){
      _heigth = '120';
      if(AttachmentCount > 2)
       _heigth = '150';
    }
    else if(_partData.length === 2)
    {
     _heigth = '200';
     if(AttachmentCount >= 2)
       _heigth = '250';
    }
    else _heigth = '300';

     hTable = new Handsontable(partContainer, {
        data: _partData,
        height: _heigth,
        width:'100%',
        dropdownMenu: false,
        columns: _cols,
        rowHeaders: true,
        stretchH: 'all',
        afterChange: function(changes, source) {
            if (source === 'edit') {
                disableCheckBox();
            }
        },
        licenseKey: htLicenseKey
    });
    if(setTimeOut){
     setTimeout(removeOverflow, 1000);
     setTimeOut = false;
    }
    else removeOverflow();
}

function getHandsonTableSchema(){
    //var _tablecols = ["RejectionTrades", "AssignedTo", "Trade", "Status", "Code", "Comment", "RejectionRemarks", "Attachments"]; //"Reference"
    //if(_module === 'SLF'){
        var newColumn;
        _cols = [
            {title: 'Tick', data: 'checkitem', type: 'checkbox', width: '3%', InternalName: 'RejectionTrades'},
            {title: 'View', data: 'view item', type: 'text', renderer: 'html', width: '3%', readOnly: true, InternalName: 'Title'},
            {title: 'Assigned To', data: 'assignedto', type: 'text', renderer: 'html', width: '15%', readOnly: true, InternalName: 'AssignedTo/Title'},
            {title: 'Trade', data: 'trade', type: 'text', width: '8%', readOnly: true, InternalName: 'Trade'},
            {title: 'Role', data: 'role', type: 'text', width: '6%', readOnly: true, InternalName: 'Role'},
            {title: 'Status', data: 'status', type: 'text', width: '5%', readOnly: true, InternalName: 'Status'}];

            if(_module !== 'SLF')
            {
                 newColumn = {
                    title: 'Code',
                    data: 'code',
                    type: 'text',
                    width: '10%',
                    readOnly: true,
                    InternalName: 'Code'
                };
                _cols.push(newColumn);
            }

            var newColumns = [
                {
                    title: 'Comment',
                    data: 'comment',
                    type: 'text',
                    width: '14%',
                    readOnly: true,
                    InternalName: 'Comment'
                },
                {
                    title: 'Reason of rejection',
                    data: 'RejectionTrades',
                    type: 'text',
                    width: '20%',
                    InternalName: 'RejectionTrades'
                }
            ];

            _cols.push(...newColumns);

            if(_module !== 'SLF')
            {
                 newColumn = {
                    title: 'attachment',
                    data: 'attachment',
                    type: 'text',
                    width: '10%',
                    renderer: 'html',
                    readOnly: true,
                    InternalName: 'Attachments'
                };
                _cols.push(newColumn);
            }
    //}
}

function disableCheckBox(){

    var colIndex = 0;
    hTable.getSettings().columns.find(item => {
          if(item.data === 'checkitem')
            linkIndex = colIndex;
          else if(item.data === 'trade')
            tradeIndex = colIndex;
          else if(item.data === 'status')
            statusIndex = colIndex;
          else if(item.data === 'RejectionTrades')
            remarkIndex = colIndex;
          else if(item.data === 'role')
            roleIndex = colIndex;
          colIndex++;
    });

    var tbodyElement = $('#parttasks').find('table.htCore')[0].querySelector('tbody');
    var trElements = tbodyElement.querySelectorAll('tr');
    $(trElements).each(function(index, trElement) {
        var tdElements = $(trElement).find('td');
        var InputBox, statusValue, roleValue;
        $(tdElements).each(function(tdIndex, tdElement) {
            if(tdIndex === linkIndex)
              InputBox = this.children[0];
            else if(tdIndex === statusIndex)
               statusValue = this.outerText;
            else if(tdIndex === roleIndex)
               roleValue = this.outerText;
          });

          if(statusValue === 'Open' || taskStatus === 'Completed' || (_isPMC && roleValue != 'TeamLeader')){
            var _mesg = 'Only completed tasks can be rejected.';
            if(_isPMC && !_isTeamLeader)
             _mesg = 'Only TeamLeader tasks can be rejected.';
             else if (statusValue === 'Completed')
               _mesg = '';

           InputBox.checked = false;
           $(InputBox).prop('disabled', true);
           $(InputBox).attr('title', _mesg);
           hTable.setCellMeta(index, remarkIndex, 'readOnly', true);
          }
    });
}

function removeOverflow(){
    $('.ht_master .wtHolder').css('overflow-x', 'hidden');
    disableCheckBox();
}