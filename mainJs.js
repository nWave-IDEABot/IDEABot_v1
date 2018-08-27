var local = {};
local.avatar = "iqlogo.png";

var remote = {};
remote.avatar = "https://developers.viber.com/images/apps/apiai-icon.png";
var SESSIONID=generateUUID();
var accessToken = "3c44974b43934bbdb1fdc030b17df30e";
var baseUrl = "https://api.api.ai/v1/";
var Opurl="https://nwave-output-v1.herokuapp.com/getop/";
var url=Opurl+SESSIONID;
function formatTime(date) {
    var hours = date.getHours();
    var minutes = date.getMinutes();
    var ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // the hour '0' should be '12'
    minutes = minutes < 10 ? '0'+minutes : minutes;
    var strTime = hours + ':' + minutes + ' ' + ampm;
    return strTime;
}            

function insertChat(who, text){
    var control = "";
    var date = formatTime(new Date());
    
     if (who == "local"){
        
         control = '<li style="width:100%;float:right;">' +
                        '<div class="msj-rta macro">' +
                            '<div class="text text-r">' +
                                '<p>'+text+'</p>' +
                                '<p><small>'+date+'</small></p>' +
                            '</div>' +                                                       
                  '</li>';                   
    }else{
        control = '<li style="width:100%;align:right;">' +
                        '<div class="msj macro">' +
                            '<div class="text text-l">' +
                                '<p>'+ text +'</p>' +
                                '<p><small>'+date+'</small></p>' +
                            '</div>' +
                        '</div>' +
                    '</li>';   
    }
    $("#messages").append(control);
    var objDiv = document.getElementById("messages");
    objDiv.scrollTop = objDiv.scrollHeight;
}

$("#chat-panel").on('click',function(){
var framewidth = $("#frame").width();
var op="";
if(framewidth > 200 ){
	framewidth = "175";
	op="0.1";
	
}else {
	framewidth = "370";
	op="1";
	
}
$(".innerframe").animate({height:'toggle',opacity: op});
$('#frame').animate({ width: framewidth,background:"black"});
});
 
function resetChat(){
    $("#messages").empty();
}


$(".mytext").on("keyup", function(e){
    if (e.which == 13){
        var text = $(this).val();
        if (text !== ""){
            insertChat("local", text);              
            $(this).val('');
            queryBot(text)
        }
    }
});
function buttonclick(text){
insertChat("local", text);              
queryBot(text);
}
resetChat();

function queryBot(text) {
            $.ajax({
                type: "POST",
                url: baseUrl + "query?v=20150910",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                headers: {
                    "Authorization": "Bearer " + accessToken
                },
                data: JSON.stringify({ query: text, lang: "en", sessionId: SESSIONID }),
                
		success: function(data) {
		    displayOutput(data.result.fulfillment.displayText);
                    insertChat("remote",data.result.fulfillment.speech);
			var msg = new SpeechSynthesisUtterance(data.result.fulfillment.speech);
            		window.speechSynthesis.speak(msg);
                },
                error: function() {
                    insertChat("remote","Sorry Bot has faced some issues! Please try again later");
                }
            });
    }
function generateUUID() { // Public Domain/MIT
    var d = new Date().getTime();
    if (typeof performance !== 'undefined' && typeof performance.now === 'function'){
        d += performance.now(); //use high-precision timer if available
    }
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = (d + Math.random() * 16) % 16 | 0;
        d = Math.floor(d / 16);
        return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    });
}

$(document).ready(function(){
    $("#myHref").click(function(event){
document.getElementById("myData").setAttribute('data',url);
//document.getElementById("myData").setAttribute('data','file:///D:/Guna/POCs/ML/nwave-UI/output.html');
	$("#myData").show();
	$("#myHref").hide();
	$("#closeOp").show();
	return false;
    });
});
$("#closeOp").click(function(){
	$("#loading").hide();
	document.getElementById("myData").setAttribute('data',"");
	$("#closeOp").hide();	
	$("#myData").hide();
	$("#myHref").show();
});
function displayOutput(input){
if (input === 'LOAD-PAGE'){
document.getElementById("myData").setAttribute('data',url);
//document.getElementById("myData").setAttribute('data','https://nwave-ideabot-flask-webhook-p.herokuapp.com/getop/TESTINPUT1');
	$("#myData").show();
	$("#myHref").hide();
	$("#closeOp").show();}
}
$(function () {
    var body = $('#myHref');
    var backgrounds = [
      'url(https://gcn.com/articles/2017/04/06/~/media/GIG/GCN/Redesign/Articles/2017/April/AIrobot.png)',
      'url(https://marketingland.com/wp-content/ml-loads/2017/09/AI1920-1-800x450.png)',
'url(http://crazzfiles.com/wp-content/uploads/2015/10/artificial-intelligence-617x416.jpg)',
'url(https://techinwire.com/wp-content/uploads/2017/07/dcx_doc6q01z349m08qrpwz2sd.jpg)',
'url(http://www.cellphoneage.com/media/magpleasure/mpblog/post_thumbnail_file/5/0/cache/1/ece9a24a761836a70934a998c163f8c8/50e2975740b22aa588236eb17342dcaf.jpeg)',
'url(http://www.innovationmanagement.se/wp-content/uploads/2016/10/artificial-intelligence-digital-transformation.png)',
];
    var current = 0;

    function nextBackground() {
        body.css(
            'background',
        backgrounds[current = ++current % backgrounds.length]);

        setTimeout(nextBackground, 3500);
    }
    setTimeout(nextBackground, 3500);
    body.css('background', backgrounds[0]);
});
 $("#menu").click(function () {
    if ($(this).find("img").attr("data-state") == "static") {
      $(this).find("img").attr("src", "https://digitalsynopsis.com/wp-content/uploads/2015/10/gif-icons-menu-transition-animations-dancing-icon.gif");
    } else {
      $(this).find("img").attr("src", "https://cdn1.iconfinder.com/data/icons/cool-cirlce/128/menu.png");
    }
  });
