
if ($("i.ms-Icon.ms-Icon--DocumentApproval").length === 0){
  $('span')
    .filter(function(){ 
      return $(this).text() == 'Submit to Area'; })
    .before("<i class='ms-Icon ms-Icon--DocumentApproval'></i>");
}


if ($("i.ms-Icon.ms-Icon--PageEdit").length === 0){
  $('span')
    .filter(function(){ return $(this).text() == 'Edit'; })
    .before("<i class='ms-Icon ms-Icon--PageEdit'></i>");
 }


 if ($("i.ms-Icon.ms-Icon--Accept").length === 0){
  $('span')
    .filter(function(){ return $(this).text() == 'Save'; })
    .before("<i class='ms-Icon ms-Icon--Accept'></i>");
 }

if ($("i.ms-Icon.ms-Icon--Accept").length === 0){
$('span')
  .filter(function(){ return $(this).text() == 'Submit'; })
  .before("<i class='ms-Icon ms-Icon--Accept'></i>");
}


if ($("i.ms-Icon.ms-Icon--Accept").length === 0){
$('span')
  .filter(function(){ return $(this).text() == 'Approve'; })
  .before("<i class='ms-Icon ms-Icon--Accept'></i>");
}


if ($("i.ms-Icon.ms-Icon--ChromeClose").length === 0){
  $('span')
  .filter(function(){ return $(this).text() == 'Close'; })
  .before("<i class='ms-Icon ms-Icon--ChromeClose'></i>");
}

if ($("i.ms-Icon.ms-Icon--ChromeClose").length === 0){
$('span')
  .filter(function(){ return $(this).text() == 'Reject'; })
  .before("<i class='ms-Icon ms-Icon--ChromeClose'></i>");
}
