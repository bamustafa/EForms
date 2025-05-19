var layout = '/_layouts/15/PCW/General/EForms';
var scriptUrl = layout + '/plumsail/Others/ShortpointFunctions.js' + '?t=' + new Date().getTime();
 
async function loadScriptAsync(src) {
    return new Promise((resolve, reject) => {
        var script = document.createElement('script');
        script.type = 'text/javascript';
        script.src = src;
 
        script.onload = function() {
            resolve(); // Resolve the promise when the script loads successfully
        };
 
        script.onerror = function() {
            console.error('Error loading script:', src); // Log an error if the script fails to load
            reject(new Error(`Failed to load script: ${src}`)); // Reject the promise on error
        };
 
        document.head.appendChild(script);
    });
}
 
async function loadMyScript() {
    try {
        loadScriptAsync(scriptUrl).then(() => {
			console.log('Script loaded successfully.');
			return onRender("SPointLoad");
		});       
    } catch (error) {
        console.error(error); // Handle any errors that occur during script loading
    }
}
 
(async function() {
    await loadMyScript(); // Use await to ensure the script is loaded before proceeding
})();