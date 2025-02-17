var preTimeOut;
var iscurtainRemoved = false, isloaderRemoved = false;
var mainDimmerElement = 'pageLayout_5a558a10,pageLayoutDesktop_5a558a10';

const preloaderMain = document.querySelector('[class^="pageLayout"]');

if (preloaderMain) {

	const style = document.createElement('style');

	style.textContent = `
	/* Overlay to dim the page */
	.preloader {
	  position: fixed;
	  top: 0;
	  left: 0;
	  width: 100%;
	  height: 100%;
	  background: rgba(0, 0, 0, 0.5); /* Semi-transparent background */
	  z-index: 9999; /* High z-index to cover all elements */
	  display: none; /* Hidden by default */
	}
	
	/* Loader animation */
	.loader {
	  position: absolute;
	  top: 50%;
	  left: 50%;
	  transform: translate(-50%, -50%);
	  border: 8px solid #f3f3f3; /* Light gray border */
	  border-top: 8px solid #33D9C2; /* Blue border for animation */
	  border-radius: 50%;
	  width: 60px;
	  height: 60px;
	  animation: spin 1s linear infinite;
	}
	
	/* Keyframes for the spinning loader */
	@keyframes spin {
	  0% {
		  transform: rotate(0deg);
	  }
	  100% {
		  transform: rotate(360deg);
	  }
	}`;

	document.head.appendChild(style);

	const preloaderMain = document.body;
	const preloaderDiv = document.createElement('div');
	preloaderDiv.id = 'preloader';
	preloaderDiv.className = 'preloader';
	
	const loaderDiv = document.createElement('div');
	loaderDiv.className = 'loader';
	
	preloaderDiv.appendChild(loaderDiv);
	preloaderMain.appendChild(preloaderDiv);
}

const preloaderItem = document.getElementById('preloader');

//var preloader_btn = async function(){
var  preloader_btn = async function(isCC, ignoreImage){
	try {
		var targetControl = 'div.ControlZone';
		 //'div.ControlZone-control'//'div.fd-toolbar-primary-commands';
		
		var webUrl = '';
		var darknessFloat = 0.8;
		let top = "300px", left = "790px";
		if (isCC) {
			webUrl = await getCurrentWebUrl();
			targetControl = '#dt';
			darknessFloat = 0.3;
			top = "250px";
			left = "750px"
		}
		else {
			webUrl = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;
		}

		var ImageUrl = webUrl + _layout + '/Images/Loading.gif';
		if (isCC)
			ImageUrl = webUrl + _layout + '/General/EForms/Images/Loading.gif';
		$(mainDimmerElement).dimBackground({
					darkness: 0.7
				}, function () {});

				if(!ignoreImage){
					$("<img id='loader' src='" + ImageUrl + "' />")
						.css({
							"position": "absolute",
							"top": top,
							"left": left,
							"width": "100px",
							"height": "100px"
						}).insertAfter(targetControl);
				}
			//$('#loader').css("visibility", "visible").dimBackground({ darkness: darknessFloat });
	}
   catch(err) { console.log(err.message); }
}
 
//var preloader = async function(isDefaultClose){
function preloader(isDefaultClose, isCC){
	try{
		if(isDefaultClose == "remove")
		{
			try{
				preTimeOut = setInterval(Remove_Pre, 200);
			    return;
			}
			catch(er){console.log(er);}

		}
		
	   var cloButton = $('div.fd-toolbar-primary-commands button:contains("Cancel")');
	   var EditLength = $('div.fd-toolbar-primary-commands button:contains("Edit")').length;

	   if(cloButton.length > 0 && isDefaultClose == true)
	   {
		   preloader_btn(); 
	   }
	   
	   else if(EditLength > 0 && isDefaultClose == true)
	   {
		  $('div.fd-toolbar-primary-commands button:contains("Edit")').click(function () {
		     preloader_btn();
		  });   
	   }
	   else preloader_btn(isCC);
	}
	catch(err) { console.log(err.message); }
}

//var Remove_Pre = async function(){
function Remove_Pre(ignoreImage){
	if(ignoreImage){
		let control = $('div.dimbackground-curtain');
		if(control.length > 0){
		control.remove();
		//clearInterval(preTimeOut);
		}
	}
	else{
			if(!iscurtainRemoved){
				console.log('Removing curtain...');
				if($('div.dimbackground-curtain').length > 0){
					$('div.dimbackground-curtain').remove();
					iscurtainRemoved = true;
				}
				else
					iscurtainRemoved = true;
			}

			if(!isloaderRemoved){
				console.log('Removing loader...');
				if($('#loader').length > 0){
					$('#loader').remove();
					isloaderRemoved = true;
				}
				else
					isloaderRemoved = true;
			}
			console.log('im here');
			if(iscurtainRemoved && isloaderRemoved){
				console.log('Both curtain and loader removed');
				clearInterval(preTimeOut);
			}
	}    
}

function showPreloader() {
	preloaderItem.style.display = 'block';
}  

function hidePreloader() {
	preloaderItem.style.display = 'none';
}