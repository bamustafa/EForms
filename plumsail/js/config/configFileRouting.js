const cacheBusting = `?cache=${Date.now()}`;

// const libraryUrls = [
//   _layout + '/controls/handsonTable/libs/handsontable.full.min.js',
//   _layout + '/controls/preloader/jquery.dim-background.min.js',
//   _layout + "/plumsail/js/buttons.js",
//   _layout + '/plumsail/js/customMessages.js',
//   _layout + '/controls/tooltipster/jquery.tooltipster.min.js',
//   _layout + '/plumsail/js/preloader.js',
//   _layout + '/plumsail/js/commonUtils.js',
//   _layout + '/plumsail/js/utilities.js',
//   _layout + '/plumsail/js/partTable.js',
//   _layout + '/plumsail/js/grid/gridUtils.js', // + cacheBusting,
//   _layout + '/plumsail/js/grid/grid.js' // + cacheBusting
// ];

const stylesheetUrls = [
  _layout + '/controls/tooltipster/tooltipster.css',
  _layout + '/controls/handsonTable/libs/handsontable.full.min.css',
  _layout + '/plumsail/css/CssStyle.css',
  _layout + '/plumsail/css/partTable.css'
];

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