var lineSymbol;
var symbol2;
var pointSymbol;
var selectedTool;
var wmUtils;

$(document).ready(function () {

    require([
        "esri/Map", 
        "esri/views/MapView",
        "esri/widgets/ScaleBar", 
        "esri/widgets/Home",
        "esri/widgets/BasemapToggle", 
        "esri/views/draw/Draw", 

        "esri/Graphic",
        "esri/layers/GraphicsLayer",

        "esri/symbols/SimpleMarkerSymbol",
        "esri/symbols/SimpleLineSymbol",
        "esri/symbols/SimpleFillSymbol",
  
        // "esri/PopupTemplate", 
        // "esri/lang", 
        // "dojo/dom-style", 
        // "dojo/on", 

        "esri/geometry/Point", 
        "esri/geometry/Polygon", 
        "esri/geometry/Polyline", 
        "esri/geometry/Extent", 
        "esri/geometry/SpatialReference",
        //"esri/InfoTemplate", 
        "esri/Color", 
        "esri/symbols/PictureMarkerSymbol", 
        //"esri/dijit/BasemapGallery", 
        "esri/geometry/support/webMercatorUtils",
        "esri/core/reactiveUtils"
        //"dojo/domReady!"
      ], function (
        Map,
        MapView,
        Scalebar, 
        Home,
        BasemapToggle, 
        Draw, 
        Graphic,
        GraphicsLayer,
        SimpleMarkerSymbol, 
        SimpleLineSymbol, 
        SimpleFillSymbol, 

        //TooltipDialog, PopupTemplate, esriLang, domStyle, on, 

        Point, Polygon, Polyline, Extent, SpatialReference,
        //InfoTemplate, 
        Color, PictureMarkerSymbol, 
        //BasemapGallery, 
        webMercatorUtils
      
      ) {
      
          wmUtils = webMercatorUtils;
          graphicsLayer = new GraphicsLayer();
          
        //lineSymbol = new esri.symbol.SimpleLineSymbol(esri.symbol.SimpleLineSymbol.STYLE_SOLID, new dojo.Color([0, 255, 255, 1]), 2)
         lineSymbol = new SimpleLineSymbol({
            color: new Color([0, 255, 255, 1]), // RGBA color
            width: 2, // Line width
            style: "solid" // Line style
          });


          pointSymbol = new SimpleMarkerSymbol(SimpleMarkerSymbol.STYLE_CIRCLE, 10, new SimpleLineSymbol(SimpleLineSymbol.STYLE_SOLID, new Color([0, 255, 255, 1]), 2), new Color([0, 255, 0, 0.25]));
          
          
          //symbol = new esri.symbol.SimpleFillSymbol(esri.symbol.SimpleFillSymbol.STYLE_SOLID, new esri.symbol.SimpleLineSymbol(esri.symbol.SimpleLineSymbol.STYLE_SOLID, new dojo.Color([255, 240, 0, 1]), 2), new dojo.Color([255, 0, 0, 0.7]));
          var outlineSymbol = new SimpleLineSymbol({
            color: new Color([255, 240, 0, 1]), // RGBA color
            width: 2, // Line width
            style: "solid" // Line style
          });
          
          symbol = new SimpleFillSymbol({
            color: new Color([255, 0, 0, 0.7]), // Fill color with transparency
            outline: outlineSymbol, // Set the outline symbol
            style: "solid" // Fill style
          });

          //symbol2 = new esri.symbol.SimpleFillSymbol(esri.symbol.SimpleFillSymbol.STYLE_SOLID, new esri.symbol.SimpleLineSymbol(esri.symbol.SimpleLineSymbol.STYLE_SOLID, new dojo.Color([0, 255, 255, 1]), 2), new dojo.Color([255, 0, 0, 0]));
          outlineSymbol = new SimpleLineSymbol({
            color: new Color([0, 255, 255, 1]), // RGBA color
            width: 2, // Line width
            style: "solid" // Line style
          });
        
          // Create the fill symbol
          symbol2 = new SimpleFillSymbol({
            color: new Color([255, 0, 0, 0]), // Fill color with transparency (fully transparent)
            outline: outlineSymbol, // Set the outline symbol
            style: "solid" // Fill style
          });

          pms = new PictureMarkerSymbol('images/location.png', 20, 27);
          map = new Map({
            basemap: "streets-vector"
        });
      
          var scalebar = new Scalebar({
              map: map,
              // "dual" displays both miles and kilmometers
              // "english" is the default, which displays miles
              // use "metric" for kilometers
              scalebarUnit: "dual"
          });
      

          var view = new MapView({
            container: "viewDiv",
            map: map,
            zoom: 4,
            center: [15, 65]
          });
          view.ui.add("srchDiv", "top-right");


        var homeWidget  = new Home({
            view: view,
            extent: new Extent(-170.049611, -59.792392, 177.438150, 71.930667)
        }, "Home");
        view.ui.add(homeWidget, "top-left");
      
      
          var basemapToggle  = new BasemapToggle({
              view: view,
              basemap: "hybrid"
          }, "BasemapToggle");
          view.ui.add(basemapToggle, "top-right");
           

          map.on("load", createToolbar);
          function createToolbar(themap) {
              toolbar = new Draw(map);
              toolbar.on("draw-end", addToMap);
              map.addLayer(graphicsLayer);
              doZoomToCountry()
          }
          function doZoomToCountry()
          {
              var Country = getQueryVariable("ac", false);
              
              var targetURL = "http://DB-PUBGIS.DARBEIRUT.COM/darAPIS/service.svc/LocateCountry/" + Country;
              $.ajax({
                  type: "GET",
                  url: targetURL,
                  contentType: "application/json; charset=utf-8", timeout: 999999999,
                  success: function (data) {
                      var jsonData = data;
                      infoData = jsonData;
                      dozoomMap(infoData.LocateCountryResult.xMin , infoData.LocateCountryResult.xMax , infoData.LocateCountryResult.yMin ,infoData.LocateCountryResult.yMax);
                      var i;
                  }, dataType: "jsonp", complete: function () { }
              }
              )
          }
          function closeDialog() {
              //ijitPopup.close(dialog);
          }

          function getQueryVariable(variable, parent) {
              var query;
             
              if (parent == true) {
                  query = window.parent.location.search.substring(1);
              }
              else {
                  query = window.location.search.substring(1);
              }
              var vars = query.split("&");
              for (var i = 0; i < vars.length; i++) {
                  var pair = vars[i].split("=");
                  if (pair[0] == variable) {
                      return pair[1];
                  }
              }
              return undefined
              //alert('Query Variable ' + variable + ' not found');
          }

          function applyCoordsOnMap(longitude, latitude, SelectionType){

            let symbol, coords
            switch (SelectionType) {
                case "point":
                    symbol = pointSymbol;
                    coords = new Point(longitude, latitude);
                    break;
                case "polygon":
                case "polyline":
                    symbol = SelectionType === 'polygon' ? symbol2 : lineSymbol;

                    let polyCoords = [];
                    let splitLong = longitude.split(',');
                    let splitLat = latitude.split(',');

                    splitLong.map((res, index)=>{
                        let parseX = parseFloat(splitLong[index].trim())
                        let parseY = parseFloat(splitLat[index].trim())
                        polyCoords.push([parseX, parseY])
                    })

                    if(SelectionType === 'polygon'){
                      coords = new Polygon(new SpatialReference({ wkid: 4326  }));
                      coords.addRing(polyCoords);
                    }
                    else{
                        coords = new Polyline(new SpatialReference({ wkid: 4326  }));
                        coords.paths = [polyCoords];
                    }
                    break;
                default:
                    break;
            }
        
            if(symbol !== undefined){
                var graphic = new Graphic(coords, symbol);
                graphicsLayer.add(graphic);

                if(SelectionType === 'point')
                  map.centerAndZoom(coords, 3);
                else{
                    let extent = coords.extent;
                    view.extent = extent;

                    // let bufferFactor = 0.5; 
                    // let bufferX = (extent.xmax - extent.xmin) * bufferFactor;
                    // let bufferY = (extent.ymax - extent.ymin) * bufferFactor;

                    // let bufferedExtent = new Extent(
                    //     extent.xmin - bufferX,
                    //     extent.ymin - bufferY,
                    //     extent.xmax + bufferX,
                    //     extent.ymax + bufferY,
                    //     extent.spatialReference
                    // );
                    // map.setExtent(bufferedExtent, true); // zoomOut
                }
            }
            
          }

          let isGISFound = JSON.parse(localStorage.getItem('isGISFound'));
      
          if(isGISFound){
              let longitude = localStorage.getItem('longitude');
              let latitude = localStorage.getItem('latitude');
              let SelectionType = localStorage.getItem('SelectionType');
      
              applyCoordsOnMap(longitude, latitude, SelectionType);
          }
      });
});

var graphicsLayer;
var map;
var toolbar;
var drawToolActive = false;

var queryGeom = undefined;

function dozoomMap(xMin, xMax, yMin, yMax) {
    var extent = new esri.geometry.Extent();

    extent.xmin = xMin;
    extent.ymin = yMin;
    extent.xmax = xMax;
    extent.ymax = yMax;
    map.setExtent(extent, true);
}
function dozoom(long, lat, ) {
    var extent = new esri.geometry.Extent();

    extent.xmin = long - 0.1;
    extent.ymin = lat - 0.1;
    extent.xmax = long + 0.1;
    extent.ymax = lat + 0.1;
    map.setExtent(extent, true);
}
function addToMap(evt) {
    //alert('hohoh');
    
    debugger;
    var symbol;
    toolbar.deactivate();
    //penImg.src = "images/poly_off.png";
    //penImg.attributes["mode"].value = "off";
    map.showZoomSlider();
    var str;
    var regExp;
    var res;
    var coordStr;

    //alert(evt.geometry.type);
    let selectionType = evt.geometry.type;
    switch (selectionType) {
        case "point":
            symbol = pointSymbol;
            //var value = wmUtils.xyToLngLat(evt.geometry.toJson().x, evt.geometry.toJson().y, true)
            //alert(value);
            res = JSON.stringify(wmUtils.webMercatorToGeographic(evt.geometry).toJson().x) + "," + JSON.stringify(wmUtils.webMercatorToGeographic(evt.geometry).toJson().y);
            //alert(res);
            break;
        case "polygon":
            symbol = symbol2;
            str = JSON.stringify(wmUtils.webMercatorToGeographic(evt.geometry).toJson().rings);
            regExp = /\],\[/gi;
            res = str.replace(regExp, "*");
            //alert(res);
            break;
        case "polyline":
            symbol = lineSymbol;
            str = JSON.stringify(wmUtils.webMercatorToGeographic(evt.geometry).toJson().paths);
            regExp = /\],\[/gi;
            res = str.replace(regExp, "*");
            //alert(res);
            break;
        default:
            //symbol = new SimpleFillSymbol();
            break;
    }


    var graphic = new Graphic(evt.geometry, symbol);
    graphicsLayer.add(graphic);
    //map.graphics.add(graphic);
    
    //alert(JSON.stringify(evt.geometry.toJson()));
    //alert(res);
   
    drawToolActive = false;
    coordStr = "\"X\",\"Y\"" + "*";
    let response = res.replace("[[[", "").replace("]]]", "");
    coordStr = coordStr + response;

    sendCoords(response, selectionType);
    
    
    //alert(coordStr);

    // parent.document.getElementById("ctl00_ContentPlaceHolder_txCoor").value = coordStr;
    // if (parent.document.getElementById("ctl00_ContentPlaceHolder_RadioButtonList1_0").checked == true) {
    //     parent.document.getElementById("ctl00_ContentPlaceHolder_txType").value = parent.document.getElementById("ctl00_ContentPlaceHolder_RadioButtonList1_0").value;
    // }

    // else {
    //     parent.document.getElementById("ctl00_ContentPlaceHolder_txType").value = parent.document.getElementById("ctl00_ContentPlaceHolder_RadioButtonList1_1").value;
    // } 
}

function activateTool(toolName) {
    drawToolActive = true;

    if (toolName == "POLYGON") {
        toolbar.activate(esri.toolbars.Draw.POLYGON);
    }

    if (toolName == "LINE") {
        toolbar.activate(esri.toolbars.Draw.POLYLINE);
    }

    if (toolName == "POINT") {
        toolbar.activate(esri.toolbars.Draw.POINT);
    }

    

    map.hideZoomSlider();
}

function makeSelection(sender) {

    //debugger;
    //alert(sender.attributes["type"].value);
    if (selectedTool != undefined) {

        selectedTool.src = selectedTool.src.replace("on", "off");
        

    }

    selectedTool = sender;

    if (sender.attributes["mode"].value == "off") {

        activateTool(sender.attributes["type"].value);
        sender.src = sender.src.replace("off","on"), //"images/poly_on.png";
        sender.attributes["mode"].value = "on";
    }

    else {
        toolbar.deactivate();
        map.showZoomSlider();
        sender.src = sender.src.replace("on","off");
        sender.attributes["mode"].value = "off";
        sendCoords(",", '');
    }

//    if (lineImg.attributes["mode"].value == "off") {

//        activateTool("POLYGON");
//        lineImg.src = "images/line_on.png";
//        lineImg.attributes["mode"].value = "on";
//    }

//    else {
//        toolbar.deactivate();
//        map.showZoomSlider();
//        lineImg.src = "images/line_off.png";
//        lineImg.attributes["mode"].value = "off";
//    }

//    if (pointImg.attributes["mode"].value == "off") {

//        activateTool("POLYGON");
//        pointImg.src = "images/Pen_on.png";
//        pointImg.attributes["mode"].value = "on";
//    }

//    else {
//        toolbar.deactivate();
//        map.showZoomSlider();
//        pointImg.src = "images/Pen_off.png";
//        pointImg.attributes["mode"].value = "off";
//    }

    graphicsLayer.clear();

}

function getDocHeight() {

    var myWidth = 0, myHeight = 0;
    if (typeof (parent.window.innerWidth) == 'number') {
        myWidth = parent.window.innerWidth;
        myHeight = parent.window.innerHeight;
    } else if (parent.document.documentElement && (parent.document.documentElement.clientWidth || parent.document.documentElement.clientHeight)) {
        myWidth = parent.document.documentElement.clientWidth;
        myHeight = parent.document.documentElement.clientHeight;
    } else if (parent.document.body && (parent.document.body.clientWidth || parent.document.body.clientHeight)) {
        myWidth = parent.document.body.clientWidth;
        myHeight = parent.document.body.clientHeight;
    }

    return myHeight;
}

function getDocWidth() {

    var myWidth = 0, myHeight = 0;
    if (typeof (parent.window.innerWidth) == 'number') {
        myWidth = parent.window.innerWidth;
        myHeight = parent.window.innerHeight;
    } else if (parent.document.documentElement && (parent.document.documentElement.clientWidth || parent.document.documentElement.clientHeight)) {
        myWidth = parent.document.documentElement.clientWidth;
        myHeight = parent.document.documentElement.clientHeight;
    } else if (parent.document.body && (parent.document.body.clientWidth || parent.document.body.clientHeight)) {
        myWidth = parent.document.body.clientWidth;
        myHeight = parent.document.body.clientHeight;
    }
    return myWidth;
}

function clearText() {
    txtSearch.value = "";
}

function sendCoords(res, selectionType){
    if(res !== undefined && res !== null){
        let resultValue, splitRes, longitude = '', latitude = '', isMulti = false;

        if(res.includes('*')){
            splitRes = res.split('*');
            isMulti = true;
        }
        
        else splitRes = res.split(',');
        
        if(splitRes.length > 0){

            if(isMulti){
                for (const items of splitRes) {

                    let splitItems = items.split(',');
                    if(splitItems.length > 0){
                        longitude += splitItems[0] + ','
                        latitude += splitItems[1] + ','
                    }
                }
                longitude = longitude.slice(0, -1); //.trimEnd(',');
                latitude = latitude.slice(0, -1); //.trimEnd(',');
            }
            else{
                longitude = splitRes[0]
                latitude = splitRes[1]
            }

            resultValue = {
                longitude: longitude,
                latitude: latitude,
                type: selectionType
            }
            window.parent.postMessage(resultValue, '*')
        }
    }
}

