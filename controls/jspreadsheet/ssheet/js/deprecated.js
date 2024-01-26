var getFieldTypes_Deprecated = async function(_listname, _cols){
    var _itemArray = {};
    var _dateArray = {};
    var _colArray = [];
    
       await $pnp.sp.web.lists.getByTitle(_listname).fields.filter("ReadOnlyField eq false and Hidden eq false").get().then(async function(items) {
            var fieldInfo = "";
            for(var i = 0; i < _cols.length; i++)
            {
                for (var j = 0; j < items.length; j++) {
                    var _item = items[j];
                    var _itemfld = _item.InternalName;

                    var fld = _cols[i];
                    if(_itemfld == fld)
                    {
                        var _type = await setSheetFieldTypes_Deprecate(_item.TypeAsString); //.TypeDisplayName;
                        
                        _itemArray.type = _type;
                        if(_type == "calendar"){
                         _dateArray.format = 'DD/MM/YYYY';
                         _itemArray.options = _dateArray;
                        }

                        _itemArray.title = _itemfld;
                        _itemArray.width = 250;
                        _colArray.push(_itemArray);
                        _itemArray = {};
                        break;
                    }
                }
            }
        }).catch(function(err) {
            alert(err);
        });

    _itemArray.type = "text";
    _itemArray.title = "_Mesg";
    _itemArray.width = 250;
    _colArray.push(_itemArray);
    _itemArray = {};

    _itemArray.type = "text";
    _itemArray.title = "_Status";
    _itemArray.width = 250;
    _colArray.push(_itemArray);
    _itemArray = {};
    return _colArray;
}

var setSheetFieldTypes_Deprecate = async function(fldtype){
   var _type = "text";
    if(fldtype == "DateTime")
    {
        _type = "calendar";
    }
    return _type;
}

function formatDate_Deprecated(date) {
    var lDate = new Date(date);
    var month = lDate.getMonth() + 1;
    if (month < 10) month = "0" + month;

    var dateOfMonth = lDate.getDate();
    if (dateOfMonth < 10) dateOfMonth = "0" + dateOfMonth;

    var year = lDate.getFullYear();
    var formattedDate = dateOfMonth + "/" + month + "/" + year;

    //console.log(lDate);
    //lDate.setDate(lDate.getDate() + 1);
    return formattedDate;
}