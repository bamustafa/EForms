function formatStringToArray(_cols){
    var partArray = [];
    if(_cols.includes(","))
	{
        _cols = _cols.split(',');
		for(var i = 0; i < _cols.length; i++)
		{
			partArray.push(_cols[i]);
		}
	}
    else partArray.push(_cols);
    return partArray;
}

var GetParameterValues = async function(param){
    var url = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');  
    for (var i = 0; i < url.length; i++) {  
        var urlparam = url[i].split('=');  
        if (urlparam[0] == param) {  
            return urlparam[1];  
        }  
    }  
} 

function formatTime(date) {
    var lDate = new Date(date);
    var hours = lDate.getHours();// + 1;
    if (hours < 10) hours = "0" + hours;

    var minutes = lDate.getMinutes()
    if (minutes < 10) minutes = "0" + minutes;

    var formattedTime = hours + ":" + minutes;

    return formattedTime;
}

function dateOption(){
    var _dateArray = {};
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    var weekdays = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    var weekdays_short = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];

    _dateArray.format = 'dd/mm/yyyy h:mm';
    _dateArray.time = 0;
    _dateArray.readonly = 0
    _dateArray.today = 0;
    _dateArray.time = 1;
    _dateArray.clear = 1;
    _dateArray.mask = 1;
    _dateArray.months = months;
    _dateArray.weekdays = weekdays;
    _dateArray.weekdays_short = weekdays_short;
    return _dateArray;
}

function parseXmlToJson(xml) {
    const json = {};
    for (const res of xml.matchAll(/(?:<(\w*)(?:\s[^>]*)*>)((?:(?!<\1).)*)(?:<\/\1>)|<(\w*)(?:\s*)*\/>/gm)) {
        const key = res[1] || res[3];
        const value = res[2] && parseXmlToJson(res[2]);
        json[key] = ((value && Object.keys(value).length) ? value : res[2]) || null;

    }
    return json;
}

function filterData(_data){
    _data = _data.filter((str) => {
         if(str[0] == "") // || str[3] == submitImg) 
           return false;
         else return true;
         });
     return _data;
}

//PNP FUNCTIONS

var getParameter = async function(_web, key){
	var result = "";
     _web.lists
			.getByTitle("Parameters")
			.items
			.select("Title,Value")
			.filter("Title eq '" + key + "'")
			.get()
			.then(function (items) {
				if(items.length > 0)
				  result = items[0].Value;
				});
	 return result;
}

var getMajorType = async function(_web, webURL, key){
  var _colArray = [];
  var targetList;
  var targetFilter;

  await _web.lists
        .getByTitle("MajorTypes")
        .items
        .select("LODListName, LODFilterColumn, HandsonTblSchema")
        .filter("Title eq '" + key + "'")
        .get()
        .then(async function (items) {
            if(items.length > 0)
              //for (var i = 0; i < items.length; i++) {
                var item = items[0];
                var HandsonTblSchema = item.HandsonTblSchema;
                targetList = item.LODListName;
                targetFilter = item.LODFilterColumn;

                  var fetchUrl = webURL + HandsonTblSchema;
                  await fetch(fetchUrl)
                      .then(response => response.text())
                      .then(async data => {
                        _colArray = JSON.parse(data); 
                  });
            });
  return {
    colArray: _colArray,
    targetList: targetList,
    targetFilter: targetFilter
  };
}

var getCounter = async function(_web, key){
    var value = 1;
    var listname = 'Counter';

    await _web.lists
          .getByTitle(listname)
          .items
          .select("Id, Title, Counter")
          .filter("Title eq '" + key + "'")
          .get()
          .then(async function (items) {
            var _cols = { };
              if(items.length == 0){
                 _cols["Title"] = key;
                 _cols["Counter"] = value.toString();
                 value = String(value).padStart(4,'0');
                 var blabla = await _web.lists.getByTitle(listname).items.add(_cols);
                 return value;
              }

              else if(items.length > 0)
                    var _item = items[0];
                    value = parseInt(_item.Counter) + 1;
                    _cols["Counter"] = value.toString();
                    value = String(value).padStart(4,'0');
                    var blabla = await _web.lists.getByTitle(listname).items.getById(_item.Id).update(_cols);
                    return value;
              });
   return value;
}

var getListFields = async function(_web, _listname){
  var fieldInfo = [];
  await  _web.lists
         .getByTitle(_listname)
         .fields
         .filter("ReadOnlyField eq false and Hidden eq false")
         .get()
         .then(function(result) {
            for (var i = 0; i < result.length; i++) {
                fieldInfo.push(result[i].InternalName);
                // fieldInfo += "Title: " + result[i].Title + "<br/>";
                // fieldInfo += "Name:" + result[i].InternalName + "<br/>";
                // fieldInfo += "ID:" + result[i].Id + "<br/><br/>";
            }
    }).catch(function(err) {
        alert(err);
    });
    return fieldInfo;
}

var createList = async function(_web, _listname, _colsInternal, _colsType, table, _data, lodRef, btnText, RegDeli){
  var isSubmit = false;
  if(btnText == 'Submit') isSubmit = true;

    //ADD RECORD IF NOT EXIST TO REGISTER DELIVERABLES
    if(isSubmit)
      addRegistryListItem(_web, RegDeli, lodRef, isSubmit);
    else if(lodRef != '')
      addRegistryListItem(_web, RegDeli, lodRef, false);

    _web.lists.filter("Title eq '" + _listname + "'").get().then(function(result) 
    {
       if (result.length == 0) {
         _web.lists.add(_listname);
       } 
    });

    var _listFields = [];
    _listFields = await getListFields(_web, _listname);

    for(var i = 0; i < _colsInternal.length; i++){
        var isFldExist = false;
        var _fld = _colsInternal[i];

        for(var j = 0; j < _listFields.length; j++){
            var _lstFld = _listFields[j];
            if(_fld === _lstFld){
                isFldExist = true;
                break;
            }
        }
        if(!isFldExist){
            await createFields(_web, _listname, _fld, _colsType[i]);
        }
    }

    var isFldExist = false;
    var _fld = "LODRef";

    for(var j = 0; j < _listFields.length; j++){
      var _lstFld = _listFields[j];
       if(_fld === _lstFld){
           isFldExist = true;
           break;
        }
    }
    if(!isFldExist){
      await createFields(_web, _listname, _fld, 'text');
    }

    var filenameItems = [];
    var dataLength = _data.length;

    for(var i =0; i < dataLength; i++){
        var _cols = { };
        var _filename = "";
        var rowData = _data[i];//table[0].getRowData(i);

        for(var j = 0; j < _colsInternal.length; j++){
            var _intCol = _colsInternal[j];
            if(isSubmit){
              if(_intCol == 'Mesg' || _intCol == 'Status'){
                continue;
              }
            }
            _cols[_intCol] = rowData[j];
            if( _intCol == 'FullFileName' ){
              _filename = rowData[j];
              filenameItems.push(_filename.trim());
            }
        }
      _cols[_fld] = lodRef;
      if(isSubmit){
        await addRLODListItem(_web, RLODList, _cols, _filename, lodRef);
      }
      else if(_filename != '')
       await addListItem(_web, _listname, _cols, _filename, lodRef);

       $('#counter').text( (i+1) + ' of ' + dataLength);
    }

    if(isSubmit)
      closeForm(true);
    else await removeItems(_web, _listname, lodRef, filenameItems, isSubmit);
}

var createFields = async function(_web, _listname, _field, _type){
    try{
    _type = _type.toLowerCase();
    if(_type === 'text' || _type === 'dropdown'|| _type === 'image')
      await _web.lists.getByTitle(_listname).fields.addText(_field);
    else if(_type === 'multi')
      await _web.lists.getByTitle(_listname).fields.addMultilineText(_field);
    else if(_type === 'calendar'){
      await _web.lists.getByTitle(_listname).fields.addDateTime(_field);
            // ,{ DisplayFormat: DateTimeFieldFormatType.DateOnly, 
            //   DateTimeCalendarType: CalendarType.Gregorian, 
            //   FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Disabled, 
            //   Group: "My Group"
            //  });
       }
       await _web.lists.getByTitle(_listname).defaultView.fields.add(_field);
    }
    catch (e) {
       console.log(e);
    }
}

var addListItem = async function(_web, _listname, objValue, _filename, lodRef){
    await _web.lists
    .getByTitle(_listname)
    .items
    //.select("Id, Title, Counter")
    .filter("FullFileName eq '" + _filename + "'")
    .get()
    .then(async function (items) {
      var _cols = { };
        if(items.length == 0){
            await _web.lists.getByTitle(_listname).items.add(objValue);
            return;
        }

        else if(items.length > 0)
              var _item = items[0];
              var mesg = _item.Mesg;
              var objmesg = objValue['Mesg'];

              if( mesg != objmesg)
                await _web.lists.getByTitle(_listname).items.getById(_item.Id).update(objValue);
              return;
        });  
}

var addRLODListItem = async function(_web, _listname, objValue, _filename, lodRef){
  await _web.lists
  .getByTitle(_listname)
  .items
  .filter("FullFileName eq '" + _filename + "'")
  .get()
  .then(function (items) {
    var _cols = { };
      if(items.length == 0){
           _web.lists.getByTitle(_listname).items.add(objValue);
          return;
      }

      else if(items.length > 0){
        var _item = items[0];
       _web.lists.getByTitle(_listname).items.getById(_item.Id).update(objValue);
        return;
      }
    });  
}

var checkRLODListItem = async function(_web, _listname, _filename){
  var result = '';
  result = _web.lists
            .getByTitle(_listname)
            .items
            .filter("FullFileName eq '" + _filename + "'")
            .get();
            // .then(function (items) {
            //     if(items.length > 0){
            //         isExist = true;
            //         return isExist;
            //       }
            //    });
    return result;
}

var addRegistryListItem = async function(_web, _listname, lodRef, isSubmit){
    await _web.lists
    .getByTitle(_listname)
    .items
    //.select("Id, Title, Counter")
    .filter("Title eq '" + lodRef + "'")
    .get()
    .then(async function (items) {
      var _cols = { };
        if(items.length == 0){
            await _web.lists.getByTitle(_listname).items.add({
                Title: lodRef
            });
        }
        else{
          if(isSubmit){
            var _item = items[0];
            await _web.lists.getByTitle(_listname).items.getById(_item.Id).update({
              Status: 'Submitted'
            });
          }
          return;
        }
    });  
}

var getTempLODListItems = async function(_web, _listname, lodRef){
    var _itemArray = [];
    await _web.lists
    .getByTitle(_listname)
    .items
    //.select("Id, Title, Counter")
    .filter("LODRef eq '" + lodRef + "'")
    .top(10000)
    .getAll()
    .then(function (items) {
      var _cols = { };
      for(var i = 0; i < items.length; i++){
        var item = items[i];
        _itemArray.push(item.FullFileName);
      }
    });  
    return _itemArray;
}

var removeItems = async function(_web, _listname, lodRef, filenameItems, isSubmit){
    var _tempLODItems = [];
    _tempLODItems = await getTempLODListItems(_web, _listname, lodRef);

    let difference = "";
    if(isSubmit)
      difference = filenameItems;
    else difference = _tempLODItems.filter(x => !filenameItems.includes(x));

    for(var i =0; i < difference.length; i++){
    var list = _web.lists.getByTitle(_listname);

    list.items.filter("FullFileName eq '" + difference[i] + "'").get().then((items) =>{ 
        items.forEach(i =>{
            list.items.getById(i["ID"]).delete().then(r => {
            //console.log("deleted");
            });
        });
    });
  }
}

function closeForm(webURL, _isSubmit){
  if(_isSubmit)
   alert('Deliverables submitted successfully!!');
  window.location = webURL + '/Lists/RegisterDeliverables';
}

