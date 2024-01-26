//const cacheBusting = `?cache=${Date.now()}`;

const libraryUrls = [
  _layout + '/controls/preloader/jquery.dim-background.min.js',
  
  _layout + '/plumsail/js/commonUtils.js',
  _layout + '/plumsail/js/utilities.js',
  _layout + '/plumsail/js/preloader.js',
  
  // _layout + '/controls/jqwidgets/scripts/jquery-1.11.1.min.js',
  _layout + '/controls/jqwidgets/scripts/demos.js',
  _layout + '/controls/jqwidgets/jqxcore.js',
  _layout + '/controls/jqwidgets/jqxlistbox.js',
  _layout + '/controls/jqwidgets/jqxscrollbar.js',
  _layout + '/controls/jqwidgets/jqxbuttons.js',
  _layout + '/controls/jqwidgets/jqxexpander.js',
  _layout + '/controls/jqwidgets/jqxvalidator.js',
  _layout + '/controls/jqwidgets/jqxinput.js',
  _layout + '/controls/jqwidgets/jqxdragdrop.js'
];

const stylesheetUrls = [
  _layout + '/plumsail/css/CssStyle.css',
  _layout + '/plumsail/css/FNC.css',
 _layout + '/controls/jqwidgets/styles/jqx.base.css'
];

function loadScript(url) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = url;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
    script.defer = true;
  });
}

function loadScripts(urls) {
  // const promises = urls.map(url => loadScript(url));
  // return Promise.all(promises);
  return urls.reduce((prevPromise, url) => {
    return prevPromise.then(() => loadScript(url));
  }, Promise.resolve());
}

function loadStylesheet(url) {
  return new Promise((resolve, reject) => {
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.type = 'text/css';
    link.href = url;
    link.onload = resolve;
    link.onerror = reject;
    document.head.appendChild(link);
  });
}

function loadStylesheets(urls) {
  // const promises = urls.map(url => loadStylesheet(url));
  // return Promise.all(promises);

  return urls.reduce((prevPromise, url) => {
    return prevPromise.then(() => loadStylesheet(url));
  }, Promise.resolve());
}

loadStylesheets(stylesheetUrls)
    .then(() => {
      //console.log('Stylesheets are loaded.');
      return loadScripts(libraryUrls);
    })
    .then(() => {
      console.log('Libraries are loaded. Executing the main JavaScript code...');
    })
    .catch(error => {
      console.error('Error loading scripts or stylesheets:', error);
 });