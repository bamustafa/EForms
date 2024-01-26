<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head>
  <script src="../../common libs/jquery/jquery-3.6.3.js"></script>
  <script src="../../common libs/pnp/pnp-3.0.10.js"></script>
  
  <!-- <script src="../../controls/handsonTable/libs/handsontable.full.min.js"></script> -->
  <script src="../LOD1.js"></script>

</head>

<body>
  <div id="dt"></div>
  <div class="section">
    <input id="valid" type="button" value="Validate" class="button1 bouncy" />
    <input id="subid" type="button" value="Submit" class="disableButton bouncy" disabled= "disabled"/>
    <input id="cloid" type="button" value="Close" class="button1 bouncy" />
    <input id="testid" type="button" value="Click Me" onclick="clickFunc()" class="button1 bouncy" />
  </div>
  <div><label id="counter" style="font-size: 20px;"></label></div>

  <script>
    onRender('LOD', 'New');
  </script>
</body>
</html>

