// A workaround for XSL-to-XHTML systems that don't
//  implement XSL 'disable-output-escaping="yes"'.

var output_escaping;
var DEBUG = 0;


//Checks whether or not the application supports disable_output_escaping option
function check_app (){
 var browser=navigator.appName;

 //If the browser doesn't support disable_output_escaping="true" attribute
 //the result will have
 if (browser == "Netscape"){
  output_escaping = true;
 }
 else {
  output_escaping = false;
 }

}


function decodeHTML () {

  check_app();

  if(!output_escaping) {
    DEBUG && alert("Nothing to do!");
    return;
  }

  var to_decode = document.getElementsByTagName('decodeable');
  if(!( to_decode && to_decode.length )) {
    DEBUG && alert("No work needs doing -- no elements to decode!");
    return;
  }

  var s;
  for(var i = to_decode.length - 1; i >= 0; i--) { 
    s = to_decode[i].textContent;

    if(
      s == undefined ||
      (s.indexOf('&') == -1 && s.indexOf('<') == -1)
    ) {
      // the null or markupless element needs no reworking
    } else {
      to_decode[i].innerHTML = s;  // here is the decoding
    }
  }

  return;
}
