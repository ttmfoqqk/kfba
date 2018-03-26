/*
이메일:heyjou@hotmail.com
■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
■■  file       = 파일명
■■  version = 버젼
■■  quality  = 품질
■■  id         = 아이디값
■■  width    = 넓이값
■■  height   = 높이값
■■  wmode = 투명모드
■■  color    = 배경색상
■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
*/

function flash(file,version,quality,id,width,height,wmode,color) {
	//파일값 있는지 체크
	if(file == ''||file == "") var file = 'random_file' + Math.round(Math.random()*10);

	//버젼값 있는지 체크
	if(version == ''||version == "") var version = '9';

	//퀄리티값 체크
	if(quality == ''||quality == "") var quality = 'best';

	//아이디값 체크
	if(id == ''||id == "") var id = 'random_id' + Math.round(Math.random()*10);

	//width값 체크
	if(width == ''||width == "") var width = '100%';

	//height값 체크
	if(height == ''||height == "") var height = '100%';

	//배경색 체크
	if(color == ''||color == "") var color = '#ffffff';
	
	document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase=http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=' + version +',0,0,0 width='+width+' height='+height+' id='+id+' align=middle>');
	document.write('<param name=allowScriptAccess value=sameDomain />');
	document.write('<param name=movie value='+file+' />');
	//투명처리
	//opaque
	if(wmode == 'yes' || wmode == '1') document.write('<param name=wmode value=transparent />');
	document.write('<param name=quality value='+quality+' />');
	document.write('<param name=bgColor value='+color+' />');
	document.write('<embed src='+file+' quality=best bgcolor='+color+' width='+width+' height='+height+' name='+id+' align=middle allowScriptAccess=sameDomain type=application/x-shockwave-flash pluginspage=http://www.macromedia.com/go/getflashplayer />');
	document.write('</object>');
}