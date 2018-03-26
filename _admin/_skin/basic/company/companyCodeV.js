$(document).ready(function(){
	getCodelist(1,0)
})

// 리스트 읽기
function getCodelist(mode,Idx){
	$('#code2').attr({ "value":mode == 2 ? Idx : $('#code2').attr('value') });
	$('#code'+mode).html('<ul><li style="width:100%;">데이터 로딩중입니다.</li></ul>');
	$.ajax({
		type: "POST",
		dataType: "xml",
		url: "companyCodeL.asp",
		data: {
			mode : mode    ,
			Idx  : Idx
		} ,
		success: function(xml){
			var admin_login = $(xml).find("admin_login").text();
			if(admin_login=='login'){
				alert('로그인 세션 만료!');location.reload();return false;
			}
			var tmp_html = '';
			var tmp_style= mode == 1 ? 'cursor:pointer;' : '';

			if ($(xml).find("data").find("item").length > 0) {
				$(xml).find("data").find("item").each(function(idx) {
					var code_idx   = $(this).find("code_idx").text();
					var code_name  = $(this).find("code_name").text();
					var code_order = $(this).find("code_order").text();
					var code_bigo  = $(this).find("code_bigo").text();
					var code_usfg  = $(this).find("code_usfg").text();
					var usfg_txt   = code_usfg == 0 ? '사용' : '미사용';
					var tmp_click  = mode == 1 ? 'getCodelist(2,' + code_idx + ')' : '';

					tmp_html += '<ul>' +
						'<li style="width:30px;"><input type="checkbox" name="codecheck" style="margin-top:3px" value="'+code_idx+'"></li>' +
						'<li style="width:50px;'+tmp_style+'" onclick="'+tmp_click+'">'+code_order+'</li>' +
						'<li style="width:89px;'+tmp_style+'" onclick="'+tmp_click+'">'+code_name+'</li>' +
						'<li style="width:50px;'+tmp_style+'" onclick="'+tmp_click+'">'+usfg_txt+'</li>' +
						'<li style="width:50px;"><img src="../_skin/basic/images/center_btn_edite_Code.gif" style="margin-top:3px;cursor:pointer;" onclick="getCodeView(\'UPDATE\',\''+mode+'\',\''+code_idx+'\')"></li>' +
					'</ul>';
				});
			}else{
				tmp_html = '<ul><li style="width:100%;">등록된 내용이 없습니다.</li></ul>';
			}			
			$('#code' + mode).html(tmp_html);
		},error:function(err){
			alert('ERR [502] : 고객센터에 문의하세요.' + err.responseText);
		}
	});
}

function getCodeView(action,mode,Idx){
	var html_btn_write = '<img src="../_skin/basic/images/center_btn_write_ok.gif" style="cursor:pointer;" value="'+action+'">';
	var html_btn_dell = ' <img src="../_skin/basic/images/center_btn_delete.gif" style="cursor:pointer;" value="DELETE">';
	var html_btn_area = html_btn_write;
	if(action == 'UPDATE'){
		html_btn_area += html_btn_dell;
	}
	var html_txt = '' +
	'<div class="admin_popup" id="admin_popup">' +
		'<div class="top_area">' +
			'<div class="title"><img src="../_skin/basic/images/pop/title_common_code.gif"></div>' +
			'<div class="close"><a href="#">[닫기]</a></div>' +
		'</div>' +
		'<div class="cont">' +
			'<table cellpadding=0 cellspacing=0 width=100%>'+
				'<tr>' +
					'<td class="line_box" align=right bgcolor="f0f0f0">순서</td>'+
					'<td class="line_box"><input type="text" id="code_ord" name="code_ord" class="input" size=7 maxlength=7 onkeyup=onlyNumber(this)></td>'+
				'</tr>'+
				'<tr>' +
					'<td class="line_box" align=right bgcolor="f0f0f0">내용</td>'+
					'<td class="line_box"><input type="text" id="code_name" name="code_name" class="input"></td>'+
				'</tr>'+
				'<tr>' +
					'<td class="line_box" align=right bgcolor="f0f0f0">사용여부</td>'+
					'<td class="line_box"><input type="radio" name="code_usfg" value=0 checked>사용 <input type="radio" name="code_usfg" value=1> 미사용</td>'+
				'</tr>'+
				'<tr>' +
					'<td class="line_box" align=right bgcolor="f0f0f0">비고</td>'+
					'<td class="line_box"><textarea id="code_bigo" name="code_bigo" style="width:100%;height:80px;"></textarea></td>'+
				'</tr>'+
			'</table>'+
		'</div>' +
		'<div class="btn_area pop_btn">' + html_btn_area + '</div>' +
	'</div>';
	

	if(action == 'UPDATE'){
		pop_loading()
		$.ajax({
			type: "POST",
			dataType: "xml",
			url: "companyCodeV.asp",
			data: {
				mode : mode ,
				Idx  : Idx
			} ,
			success: function(xml){
				$('body').append(html_txt);

				var admin_login = $(xml).find("admin_login").text();
				if(admin_login=='login'){
					alert('로그인 세션 만료!');location.reload();return false;
				}
				if ($(xml).find("data").find("item").length > 0) {
					$(xml).find("data").find("item").each(function(idx) {
						var code_idx   = $(this).find("code_idx").text();
						var code_name  = $(this).find("code_name").text();
						var code_order = $(this).find("code_order").text();
						var code_bigo  = $(this).find("code_bigo").text();
						var code_usfg  = $(this).find("code_usfg").text();

						$('#code_ord').val( code_order );
						$("#code_name").val( code_name );
						$('input[name=code_usfg]').filter("input[value="+code_usfg+"]").attr("checked", "checked");
						$('#code_bigo').val( code_bigo );
					});
				}
				$('#admin_popup .close a').click(function(e){
					e.preventDefault();
					layerPopupClose('wall','admin_popup');
				});
				$('#admin_popup .pop_btn img').click(function(e){
					e.preventDefault();
					goAction( $(this).attr('value') , mode , Idx )
				});
				layerPopupOpen('wall',10,'admin_popup',20);
				layerPopupClose('wall_loading','pop_loading');

			},error:function(err){
				alert('ERR [502] : 고객센터에 문의하세요.' + err.responseText);
				layerPopupClose('wall_loading','pop_loading');
			}
		});
	}else{
		$('body').append(html_txt);
		$('#admin_popup .close a').click(function(e){
			e.preventDefault();
			layerPopupClose('wall','admin_popup');
		});
		$('#admin_popup .pop_btn img').click(function(e){
			e.preventDefault();
			goAction( $(this).attr('value') , mode , Idx )
		});
		layerPopupOpen('wall',10,'admin_popup',20);		
	}
}

function goAction( actType , mode , Idx ){
	if(actType == 'DELETE'){
		if(confirm("삭제 하시겠습니까?")){
			
		}else{
			return false;
		}
	}
	$('#admin_popup .pop_btn').html('처리중입니다.');
	var html_btn_write = '<img src="../_skin/basic/images/center_btn_write_ok.gif" style="cursor:pointer;" value="'+actType+'">';
	var html_btn_dell = ' <img src="../_skin/basic/images/center_btn_delete.gif" style="cursor:pointer;" value="DELETE">';
	var html_btn_area = html_btn_write;
	if(actType == 'UPDATE'){
		html_btn_area += html_btn_dell;
	}

	$.ajax({
		type: "POST",
		url: "companyCodeP.asp",
		data: {
			actType : actType ,
			mode    : mode ,
			Name    : $('#code_name').val() ,
			Ord     : $('#code_ord').val() ,
			Idx     : Idx ,
			UsFg    : $(':radio[name="code_usfg"]:checked').val() ,
			Bigo    : $('#code_bigo').val()
		} ,
		success: function(msg){
			if(msg == 'login'){
				alert('로그인 세션이 만료되었습니다.');
				location.reload();
			}else if(msg == 'success'){
				alert('정상처리 되었습니다');
				layerPopupClose('wall','admin_popup');
				getCodelist(mode, mode == 1 ? Idx : $('#code2').attr('value') )
			}else{
				alert('데이터 처리 오류');
				layerPopupClose('wall','admin_popup');
			}
		},error:function(err){
			alert('ERR [502] : 고객센터에 문의하세요.' + err.responseText);
		}
	});
}

function go_delete(mode){
	if( $( "#code"+mode+" :checkbox[name='codecheck']:checked").length==0 ){
		alert("삭제할 항목을 하나이상 체크해주세요.");
		return;
	}
	var chked_val = "";
	$( "#code"+mode+" :checkbox[name='codecheck']:checked").each(function(pi,po){
		chked_val += ","+po.value;
	});
	if(chked_val!="")chked_val = chked_val.substring(1);

	goAction( 'DELETE' , mode , chked_val )
}