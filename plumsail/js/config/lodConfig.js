const cacheBusting = `?cache=${Date.now()}`;

const libraryUrls = [
  _layout + '/controls/preloader/jquery.dim-background.min.js',
  _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
  _layout + '/plumsail/js/customMessages.js',
  _layout + '/plumsail/js/commonUtils.js',
  _layout + '/controls/preloader/jquery.dim-background.min.js',
  _layout + '/plumsail/js/preloader.js'
];

const stylesheetUrls = [
  _layout + '/controls/tooltipster/tooltipster.css',
 _layout + '/controls/handsonTable/libs/handsontable.full.min.css',
 _layout + '/plumsail/css/CssStyle.css'
];

if(_module === 'DTRD'){
  //libraryUrls.push(_layout + '/controls/jqwidgets/jqxcore.js');
  libraryUrls.push(_layout + '/plumsail/js/preloader.js');
  libraryUrls.push(_layout + '/controls/jqwidgets/jqxbuttons.js');
  libraryUrls.push(_layout + '/controls/jqwidgets/jqxscrollbar.js');
  libraryUrls.push(_layout + '/controls/jqwidgets/jqxlistbox.js');
  libraryUrls.push(_layout + '/controls/jqwidgets/jqxcheckbox.js');
  libraryUrls.push(_layout + '/controls/jqwidgets/jqxsplitter.js');

  stylesheetUrls.push(_layout + '/plumsail/css/CssStyleRACI.css');
  stylesheetUrls.push(_layout + '/controls/jqwidgets/styles/jqx.base.css');
  stylesheetUrls.push(_layout + '/controls/jqwidgets/styles/jqx.summer.css');
}

function loadScript(url) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = url;
    script.onload = resolve;
    script.onerror = reject;
    script.async = true;
    document.head.appendChild(script);
  });
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

async function loadResources() {
  try {
    await Promise.all([
      Promise.all(libraryUrls.map(url => loadScript(url))),
      Promise.all(stylesheetUrls.map(url => loadStylesheet(url)))
    ]);

    console.log('Stylesheets and libraries are loaded. Executing the main JavaScript code...');
    // Place your main JavaScript code here
  } catch (error) {
    console.error('Error loading scripts or stylesheets:', error);
  }
}

loadResources();