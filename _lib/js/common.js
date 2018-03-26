function bluring(){if(event.srcElement.tagName=="A"||event.srcElement.tagName=="IMG") document.body.focus();}
document.onfocusin=bluring;

//ÆÄ¶ó¹ÌÅÍ ¸®Äù½ºÆ®
var Request = function(){
	this.getParameter = function( name ){
	var rtnval = '';
	var nowAddress = unescape(location.href);
	var parameters = (nowAddress.slice(nowAddress.indexOf('#')+1,nowAddress.length)).split('&');
		for(var i = 0 ; i < parameters.length ; i++){
			var varName = parameters[i].split('=')[0];
			if(varName.toUpperCase() == name.toUpperCase()){
				rtnval = parameters[i].split('=')[1];break;
			}
		}
	return rtnval;
	}
}
var request = new Request();

/*=====================================================================
 * ´Þ·Â.
 *=====================================================================*/
var calendarDivId    = "__DIV_CALENDAR__";
var calendarIframeId = "__IFRAME_CALENDAR__";
//´Þ·Â º¸¿©ÁÙ À§Ä¡ Ã£±â:Top
function getRealOffsetTop(o) {
	return o ? o.offsetTop + getRealOffsetTop(o.offsetParent) : 3;
}
//´Þ·Â º¸¿©ÁÙ À§Ä¡ Ã£±â:Left
function getRealOffsetLeft(o) {
	return o ? o.offsetLeft + getRealOffsetLeft(o.offsetParent) : 2;
}
function hideCalendar()	{
	var cal = document.getElementById(calendarDivId);
		if(cal) cal.style.display = "none";
}
function callCalendar(obj)	{
	var top  = getRealOffsetTop(obj)+17;
	var left = getRealOffsetLeft(obj)-6;
	
	var param = '';
	if(obj) {
		var tokens = obj.value.split("-");
		if(tokens.length==3)
		{
			param   = "&year="+tokens[0];
			param  += "&month="+tokens[1];
			param  += "&day="+tokens[2];
		}
	}
	url = "/_lib/calender.asp?obj="+obj.form.name+"."+obj.name;	
	if(param!="") url += param;
	var width = 215;
	var height = 160;
	var ifrm = document.getElementById(calendarIframeId);
	var div = document.getElementById(calendarDivId);
	if(!div)
	{
		div = document.createElement("DIV");
		div.id = calendarDivId;
		div.style.display = "none";
		div.style.position = "absolute";
		div.style.left = left + 'px';
		div.style.top = top + 'px';
		div.style.zIndex = 100;
		document.body.appendChild(div);
	}else{
		div.style.left = left;
		div.style.top = top;
	}
	if(!ifrm)
	{
		ifrm = document.createElement("IFRAME");
		ifrm.id = calendarIframeId;
		ifrm.width = width;
		ifrm.height = height;
		ifrm.frameBorder = 0;
		ifrm.scrolling = "no";
		div.appendChild(ifrm);
	}
	div.style.display = "inline";
	ifrm.src = url;
}

//-------------------------------------------------------
// ¿À´Ã , 7ÀÏ , 30ÀÏÈÄ
//-------------------------------------------------------
function date_input(Indate,Outdate,value1,value2){
	Indate = document.getElementById(Indate);
	Outdate = document.getElementById(Outdate);
	Indate.value=value1;
	Outdate.value=value2;
}

/*===========================================================================
 * DIV ÆË¾÷Ã¢ ¿­±â
 *===========================================================================*/
function layerPopupOpen(wall_id,wall_zindex,pop_id,pop_zindex,top_px){
	var $tmp_wall = '<div class="wall" id="'+wall_id+'"></div>';
	var $layerPopupObj = $('#'+pop_id);
	var left = ( $(window).scrollLeft() + ($(window).width() - $layerPopupObj.width()) / 2 ); 
	var top_center = ( $(window).scrollTop() + ($(window).height() - $layerPopupObj.height()) / 2 ); 
	var top_pix = ( $(window).scrollTop() + top_px ); 
	var top = top_px ? top_pix : top_center;

	$('body').append($tmp_wall);
	$('#'+wall_id).css({"z-index":wall_zindex,"height":$(document).height(),"opacity":0.5});

	$layerPopupObj.css({'left':left,'top':top,"z-index":pop_zindex});
}
/*===========================================================================
 * DIV ÆË¾÷Ã¢ »èÁ¦
 *===========================================================================*/
function layerPopupClose(wall_id,pop_id){
	var $tmp_wall = $('#'+wall_id);
	var $layerPopupObj = $('#'+pop_id);
	$tmp_wall.remove();
	$layerPopupObj.remove();
}

/*===========================================================================
 * ¼ýÀÚ¸¸
 *===========================================================================*/
function onlyNumber(str){
	var strobj = str; //ÀÔ·Â°ªÀ» ´ãÀ»º¯¼ö.
	re = /[^0-9]/gi;
	if(re.test(strobj.value)){
		//alert("Æ¯¼ö¹®ÀÚ´Â ÀÔ·ÂÇÏ½Ç¼ö ¾ø½À´Ï´Ù.");
		strobj.value=strobj.value.replace(re,"");
	}
}
/*===========================================================================
 * °ø¹é, Æ¯¼ö¹®ÀÚ 
 *===========================================================================*/
function CheckSpace(str) {
	var strobj = str; //ÀÔ·Â°ªÀ» ´ãÀ»º¯¼ö.
	var re = /[\s~!|@\#$%^&*\()\-=+\\/\[\]?<>,."']/gi;
	if(re.test(strobj.value)){
		//alert("Æ¯¼ö¹®ÀÚ´Â ÀÔ·ÂÇÏ½Ç¼ö ¾ø½À´Ï´Ù.");
		strobj.value=strobj.value.replace(re,"");
	}
}

/*===========================================================================
 * ±âÃÊÄÚµå ÄÞº¸¹Ú½º ¿É¼Ç Ãß°¡ÇÏ±â fc_lib.asp [fc_code_list]
 *===========================================================================*/
function getCodeAdd_combobox(objId,txt,mode,val){
	$('#'+objId+' option').remove();
	var oTex = "<option value=''>¼±ÅÃ</option>";
	$(oTex).appendTo($('#'+objId));

	var tmp_arry = txt.split('|_ARRY_|');
	if(tmp_arry.length > 0){
		for(var i=0 ; i < tmp_arry.length ; i++) {
			var o = tmp_arry[i].split("|_KEY_|");
			var k = mode == "idx" ? o[0] : o[1] ;
			var s = k == val ? "selected" : "";
			var paramLi = "";
			paramLi += "<option value='" + k + "' "+s+">";
			paramLi += o[1];
			paramLi += "</option>";
			$(paramLi).appendTo($('#'+objId));
		}
	}
}

/*===========================================================================
 * TRIM È­ÀÌÆ®½ºÆäÀÌ½º Á¦°Å
 *===========================================================================*/
function trim(str){
	str = str.replace(/^\s*/,'').replace(/\s*$/, '');
	return str; //º¯È¯ÇÑ ½ºÆ®¸µÀ» ¸®ÅÏ.
}

var reg_id      = /[\s~!|@\#$%^&*\(){}`;:\-=+\\/\[\]?<>,."'¤¡-¤¾|¤¿-¤Ó|°¡-ÆR]/gi;
var reg_name    = /[\s~!|@\#$%^&*\(){}`;:\-_=+\\/\[\]?<>,."'0-9a-zA-Z]/gi;
var reg_number  = /[\s~!|@\#$%^&*\(){}`;:\-_=+\\/\[\]?<>,."'¤¡-¤¾|¤¿-¤Ó|°¡-ÆRa-zA-Z]/gi;
var reg_default = /[\s~!|@\#$%^&*\(){}`;:\-_=+\\/\[\]?<>,."'¤¡-¤¾|¤¿-¤Ó|°¡-ÆR]/gi;
var reg_email   = /[\s~!|\#$%^&*\(){}`;:\=+\\/\[\]?<>,"'¤¡-¤¾|¤¿-¤Ó|°¡-ÆR]/gi;
var reg_check   = /^([0-9a-zA-Z_-]+)@([0-9a-zA-Z_-]+)(\.[0-9a-zA-Z_-]+){1,2}$/;

//°ø¹é,Æ¯¼ö¹®ÀÚ Ã¼Å©
function CheckSpace(str,mode) {
   var reg = eval('reg_'+mode);
   return (reg.test(str));
}

//°ø¹é,Æ¯¼ö¹®ÀÚ °æ°í, »èÁ¦
function CheckSpace_alert(str,mode){
  var strobj = str;
  var reg = eval('reg_'+mode);
  if(reg.test(strobj.val())){
    alert("ÇÑ±Û ¶Ç´Â Æ¯¼ö¹®ÀÚ´Â ÀÔ·ÂÇÏ½Ç¼ö ¾ø½À´Ï´Ù.");
    strobj.val( strobj.val().replace(reg,"") );
  }
}

//ºñ¹Ð¹øÈ£ ¿µ¹® ¼ýÀÚ Á¶ÇÕ Ã¼Å©
function CheckPass(str){ 
	var reg1 = /^[a-zA-Z0-9]{5,20}/g;
	var reg2 = /[a-zA-Z]/g; 
	var reg3 = /[0-9]/g; 
	return ( reg1.test(str) && reg2.test(str) && reg3.test(str) ); 
}

//ÀÌ¸ÞÀÏ Çü½ÄÃ¼Å©
function check_email(txt){
	var reg = reg_check;
	return (reg.test(txt));
}

/*===========================================================================
 * TagDecode script
 *===========================================================================*/
function TagDecode(str){
	var temp;
	temp = str.replace(/&quot;/gi,"\"")
	temp = temp.replace(/&#39;/gi,"\'")
	temp = temp.replace(/&lt;/gi,"<")
	temp = temp.replace(/&gt;/gi,">")
	temp = temp.replace(/<br>/gi,"\n")
	temp = temp.replace(/&amp;/gi,"&")
	return temp;
}

/*===========================================================================
 * ¿ìÆí¹øÈ£
 *===========================================================================*/
function zipAddrSearch(zip1,zip2,addr1,addr2){
	var html_txt = '' +
	'<div class="ZIP_POP" id="ZIP_POP">' + 
		'<div style="float:right;padding:10px 20px 10px 0px;" class="close"><a href="#"><img src="/_lib/img/zip/icon_close.gif" align=absmiddle></a></div>' +
		'<table cellpadding=0 cellspacing=0 border=0 width=400>' +
			'<tr>' +
				'<td align=center><img src="/_lib/img/zip/zip_01.jpg"></td>' +
			'</tr>' +
			'<tr>' +
				'<td align=center>' +
					'<table cellpadding=0 cellspacing=5 width=370 bgcolor="#eaeaea">' +
						'<tr>' +
							'<td style=padding:1 bgcolor="#ffffff">' +					
								'<table cellpadding=0 cellspacing=0 width=100% bgcolor="#eaeaea">' +
									'<tr><td height=10></td></tr>' +
									'<tr>' +
										'<td width=250 align=center><input type="text" name="schDong" id="schDong" class="input" style="width:225px;ime-mode:active;"></td>' +
										'<td><a href="#"><img src="/_lib/img/zip/zip_btn.jpg" align=absmiddle id="schDong_search_btn"></a></td>' +
									'</tr>' +
									'<tr><td height=10></td></tr>' +
								'</table>' +
							'</td>' +
						'</tr>' +
					'</table>' +
				'</td>' +
			'</tr>' +
			'<tr><td height=10></td></tr>' +
			'<tr>' +
				'<td align=center>' +
					'<div style="width:370px;margin:0 auto;background-color:#f6f6f6;border-top:1px solid #d9d9d9">' +
						'<div style="clear:both;float:left;width:61px;line-height:23px;color:#555555;border-right:1px solid #d9d9d9;border-bottom:1px solid #d9d9d9">¿ìÆí¹øÈ£</div>' +
						'<div style="float:left;width:308px;line-height:23px;border-bottom:1px solid #d9d9d9;">ÁÖ¼Ò</div>' +
					'</div>'+
					'<div style="clear:both;width:370px;height:180px;margin:0 auto;overflow-y:auto;" id=zipArea></div>' +
				'</td>' +
			'</tr>' +
		'</table>' +
	"</div>";
		

	$('body').append(html_txt);
	$('#ZIP_POP .close a').click(function(e){
		e.preventDefault();
		layerPopupClose('wall_zip','ZIP_POP');
	});
	$('#schDong').keyup(function(e){
		if (e.keyCode == 13){
			zipAddrSearchData(zip1,zip2,addr1,addr2);
		}
	});
	$('#schDong_search_btn').click(function(e){
		e.preventDefault();
		zipAddrSearchData(zip1,zip2,addr1,addr2);
	});
	layerPopupOpen('wall_zip',500,'ZIP_POP',520);
	$('#schDong').focus()
}

function zipAddrSearchData(zip1,zip2,addr1,addr2){
	var html = '';
	$('#zipArea').html('<div style="text-align:center;line-height:23px;">·ÎµùÁßÀÔ´Ï´Ù.</div>');
	$.ajax({
		type: "GET",
		dataType: "xml",
		url: "/_lib/zip_rss.asp",
		data: {
			schDong : $('#schDong').val()
		} ,
		success: function(xml){
			if ($(xml).find("data").find("item").length > 0) {
				$(xml).find("data").find("item").each(function(idx) {

					var zipcode = $(this).find("zipcode").text();
					var addr    = $(this).find("addr").text();
					var bunji   = $(this).find("bunji").text();

					html += '' +
					'<table cellpadding=0 cellspacing=0 height=23><tr align=center>' +
					'<td style="border-right:1px solid #d9d9d9;border-bottom:1px solid #d9d9d9" width=65><a href=# zipcode="'+zipcode+'" addr="'+addr+'" bunji="'+bunji+'">' + zipcode + '</a></td>' +
					'<td style="border-bottom:1px solid #d9d9d9;padding-left:5px;" width=308 align=left><a href=# zipcode="'+zipcode+'" addr="'+addr+'" bunji="'+bunji+'">'+ addr +'  '+ bunji +'</a></td>' +
					'</tr></table>';
				});
			}else{
				html = '<div style="text-align:center;line-height:23px;">µ¥ÀÌÅÍ°¡ ¾ø½À´Ï´Ù.</div>'
			}
			$('#zipArea').html(html);
			$('#zipArea a').click(function(e){
				e.preventDefault();
				var tmp_zip = $(this).attr('zipcode').split('-');
				$('#'+zip1).val( tmp_zip[0] );
				$('#'+zip2).val( tmp_zip[1] );
				$('#'+addr1).val( $(this).attr('addr') + ' ' + $(this).attr('bunji') );
				$('#'+addr2).focus();
				layerPopupClose('wall_zip','ZIP_POP');
			});
		},error:function(err){
			alert('ERR [502] : °í°´¼¾ÅÍ¿¡ ¹®ÀÇÇÏ¼¼¿ä.' + err.responseText);
		}
	});
}

//ÀÌ¹ÌÁö ·Ñ¿À¹ö
$(function() {
	$("img.rollover").mouseover(function() {
		$(this).attr("src", $(this).attr("src").replace("_off","_on"));
	});
	$("img.rollover").mouseout(function() {
		$(this).attr("src", $(this).attr("src").replace("_on", "_off"));
	});
});

//¿ÞÂÊ¸Þ´º li ·Ñ¿À¹ö
$(function() {
	$("li.rollover").mouseover(function() {
		$(this).find('img').attr("src", $(this).find('img').attr("src").replace("_off","_on"));
	});
	$("li.rollover").mouseout(function() {
		$(this).find('img').attr("src", $(this).find('img').attr("src").replace("_on", "_off"));
	});
	//li ¸µÅ©
	$("li.rollover").click(function(){
		location.href = $(this).find('a').attr('href');
	})
});

 //·Îµù
function pop_loading(){
	var html = '<div class="pop_box" id="pop_loading" style="text-align:center;width:200px;line-height:100px;"><b>LOADING..</b></div>';
	$('body').append(html);
	layerPopupOpen('wall_loading',1000,'pop_loading',1200);
}


function chkEmail(str){
	var reg_email = /^([0-9a-zA-Z_\.-]+)@([0-9a-zA-Z_-]+)(\.[0-9a-zA-Z_-]+){1,2}$/;
	if(!reg_email.test(str)){
		return false;
	}
	return true;
}

//Å¬¸¯±Ûº¹»ç
function TextClipBoard(x){
	if(x){
		alert("Å¬¸³º¸µå¿¡ ¼³Ä¡ ÄÚµå°¡ ÀúÀåµÇ¾ú½À´Ï´Ù.\n\nCtrl+VÅ°¸¦ ´­·¯ ºÙ¿©³Ö¾î »ç¿ëÇÏ¼¼¿ä.")
		window.clipboardData.setData('Text',x)
	}
}
/* ÆÄÀÏÈ®ÀåÀÚ ¸®ÅÏ */
function fileTypeCheck( f ){
	var p = f.lastIndexOf('.');
	var t = f.substring(p+1,f.length);
	return t;
}