/*
�̸���:heyjou@hotmail.com
������������������������������������������
���  file       = ���ϸ�
���  version = ����
���  quality  = ǰ��
���  id         = ���̵�
���  width    = ���̰�
���  height   = ���̰�
���  wmode = ������
���  color    = ������
������������������������������������������
*/

function flash(file,version,quality,id,width,height,wmode,color) {
	//���ϰ� �ִ��� üũ
	if(file == ''||file == "") var file = 'random_file' + Math.round(Math.random()*10);

	//������ �ִ��� üũ
	if(version == ''||version == "") var version = '9';

	//����Ƽ�� üũ
	if(quality == ''||quality == "") var quality = 'best';

	//���̵� üũ
	if(id == ''||id == "") var id = 'random_id' + Math.round(Math.random()*10);

	//width�� üũ
	if(width == ''||width == "") var width = '100%';

	//height�� üũ
	if(height == ''||height == "") var height = '100%';

	//���� üũ
	if(color == ''||color == "") var color = '#ffffff';
	
	document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase=http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=' + version +',0,0,0 width='+width+' height='+height+' id='+id+' align=middle>');
	document.write('<param name=allowScriptAccess value=sameDomain />');
	document.write('<param name=movie value='+file+' />');
	//����ó��
	//opaque
	if(wmode == 'yes' || wmode == '1') document.write('<param name=wmode value=transparent />');
	document.write('<param name=quality value='+quality+' />');
	document.write('<param name=bgColor value='+color+' />');
	document.write('<embed src='+file+' quality=best bgcolor='+color+' width='+width+' height='+height+' name='+id+' align=middle allowScriptAccess=sameDomain type=application/x-shockwave-flash pluginspage=http://www.macromedia.com/go/getflashplayer />');
	document.write('</object>');
}