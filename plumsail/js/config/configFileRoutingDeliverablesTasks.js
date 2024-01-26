const cacheBusting = `?cache=${Date.now()}`;

const libraryUrls = [
  _layout + '/controls/preloader/jquery.dim-background.min.js' + cacheBusting,
  _layout + "/plumsail/js/customMessages.js",
  _layout + '/controls/tooltipster/jquery.tooltipster.min.js' + cacheBusting,
  _layout + '/plumsail/js/preloader.js' + cacheBusting,
  _layout + '/plumsail/js/utilities.js' + cacheBusting,
];

const stylesheetUrls = [
  _layout + '/controls/tooltipster/tooltipster.css',
   _layout + '/plumsail/css/CssStyle.css'
];

function loadScript(url) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = url;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
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