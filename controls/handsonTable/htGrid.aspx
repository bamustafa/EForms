<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head>
    <script type="text/javascript" src="libs/handsontable.full.min.js"></script>
    <link rel="stylesheet" href="libs/handsontable.full.min.css" />
</head>

<body>
    <div id="example"></div>
  <script>
    const container = document.querySelector('#example');
    const hot = new Handsontable(container, {
        data: [
            ['', 'Tesla', 'Volvo', 'Toyota', 'Ford'],
            ['2019', 10, 11, 12, 13],
            ['2020', 20, 11, 14, 13],
            ['2021', 30, 15, 12, 13]
        ],
        rowHeaders: true,
        colHeaders: true,
        licenseKey: 'non-commercial-and-evaluation' // for non-commercial use only
    });
  </script>
</body>
</html>