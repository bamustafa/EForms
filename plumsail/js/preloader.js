var preTimeOut;
var iscurtainRemoved = false, isloaderRemoved = false;
var mainDimmerElement = 'pageLayout_5a558a10,pageLayoutDesktop_5a558a10';

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