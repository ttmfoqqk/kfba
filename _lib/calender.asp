<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%@ Language=VBScript %>
<%
	Dim obj	: obj = request("obj")
	Dim yy	: yy = request("year")
	Dim mm	: mm = request("month")
	Dim dd	: dd = request("day")
%>
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
	<META http-equiv="imagetoolbar" content="no">
	<title>>SPWEB</title>
	<link href="./css/common.css" rel="stylesheet" type="text/css" />

	<script language="JScript">
	<!-- 
		// Begin  //place these scripts within BODY tag if you are using IE 4.0 or below.
		//****************************************************************************
		// PopCalendar 4.22, Emailware (please mail&commend me if u like it)
		// Originally coded by Liming(Victor) Weng, email: victorwon@netease.com
		// Release date: 2001.1.4
		// Anyone may modify it to satify his needs, but please leave this comment ahead.
		// It is free only for non-commercial usage. Contact me if you use it on a commercial site.
		//****************************************************************************
		
		var gdCtrl = new Object();
		var gcGray = "#d0d0d0";
		var gcToggle = "#FF6699";
		var gcBG = "#ffffff";
		var gcTodayBG = "#FF6699";
		var gcTemp = gcBG;
		
		var gdCurDate = new Date();
		var giYear = gdCurDate.getFullYear();
		var giMonth = gdCurDate.getMonth()+1;
		var giDay = gdCurDate.getDate();
		var VicPopCal = new Object();

		var today  = new Date();
		var giStartYear	= (new Date()).getFullYear()-5;

		var giEndYear	= (new Date()).getFullYear()+1;
		
		var obj = eval('parent.document.<%=obj%>');
		
		// 매개변수로 컨트롤이 넘어온 경우 값을 직접 넣어줌.
		function fSetDate(iYear, iMonth, iDay){
			// 월이나 날짜가 10보다 작은 경우 앞에 0을 붙여 줌.
			var tMonth0 = iMonth;
			var tDay0 	= iDay;
			
			if (iMonth < 10) tMonth0 = "0" + iMonth;
			if (iDay   < 10) tDay0   = "0" + iDay;
			
			gdCtrl.value = iYear + "-" + tMonth0 + "-" + tDay0;

			// 매개변수로 컨트롤이 넘어온 경우 직접 값을 넣어줌.
			if(obj) obj.value = gdCtrl.value;

			parent.hideCalendar(); // 달력 숨기기.
		}
		
		// 특정 일짜를 클릭했을때.
		function fSetSelected(aCell){
		  var iOffset = 0;
		  var iYear = parseInt(tbSelYear.value);
		  var iMonth = parseInt(tbSelMonth.value);
		
		  aCell.bgColor = gcBG;
		  with (aCell.children["cellText"]){
		  	var iDay = parseInt(innerText);
		  	if (color==gcGray)
				iOffset = (Victor<10)?-1:1;
			iMonth += iOffset;
			if (iMonth<1) {
				iYear--;
				iMonth = 12;
			}else if (iMonth>12){
				iYear++;
				iMonth = 1;
			}
		  }
		  fSetDate(iYear, iMonth, iDay);
		}
		
		function fBuildCal(iYear, iMonth) {
		  var aMonth=new Array();
		  for(i=1;i<7;i++)
		  	aMonth[i]=new Array(i);
		
		  var dCalDate=new Date(iYear, iMonth-1, 1);
		  var iDayOfFirst=dCalDate.getDay();
		  var iDaysInMonth=new Date(iYear, iMonth, 0).getDate();
		  var iOffsetLast=new Date(iYear, iMonth-1, 0).getDate()-iDayOfFirst+1;
		  var iDate = 1;
		  var iNext = 1;
		
		  for (d = 0; d < 7; d++)
			aMonth[1][d] = (d<iDayOfFirst)?-(iOffsetLast+d):iDate++;
		  for (w = 2; w < 7; w++)
		  	for (d = 0; d < 7; d++)
				aMonth[w][d] = (iDate<=iDaysInMonth)?iDate++:-(iNext++);
		  return aMonth;
		}
		
		function fDrawCal(iYear, iMonth, iCellWidth, iDateTextSize) {
		  var styleTD = " height=13 width='"+iCellWidth+"' style='cursor:hand;font-family:Verdana;font-size:"+iDateTextSize+"px;font-weight: bold; text-align:right"; //Coded by Liming Weng(Victor Won)  email:victorwon@netease.com
		
		  with (document) {
		  	for (w = 1; w < 7; w++) {
				write("<tr class='C_ver9'>");
				for (d = 0; d < 7; d++) {
					write("<td id=calCell class='calendar_D' "+styleTD+";cursor:pointer;text-align:right;' onMouseOver='gcTemp=this.bgColor;this.bgColor=gcToggle' onMouseOut='this.bgColor=gcTemp' onclick='fSetSelected(this)'>");
					write("<font id=cellText Victor='Liming Weng'> </font>");
					write("</td>")
				}
				write("</tr>");
			}
		  }
		}
		
		function fUpdateCal(iYear, iMonth) {
		  var today = new Date();

		  myMonth = fBuildCal(iYear, iMonth);
		  var i = 0;
		  var iDate = 0;
		  for (w = 0; w < 6; w++)
			for (d = 0; d < 7; d++)
				with (cellText[(7*w)+d]) {
					Victor = i++;
					if (myMonth[w+1][d]<0) {
						color = gcGray;
						innerText = -myMonth[w+1][d];
						iDate = 0;
					}else{
						//if(today.getDate()==i) alert(i); //오늘날짜인지 여부.
						
						color = (today.getDate()==(iDate+1))?"#FFFFFF":(d==6)?"#184485":(d==0)?"#EF70A9":"#97AAC8";
						innerText = myMonth[w+1][d];
						iDate++;
					}
					parentElement.bgColor = ((iYear==giYear)&&(iMonth==giMonth)&&(iDate==giDay))?gcTodayBG:gcBG;
				}
		}
		
		function fSetYearMon(iYear, iMon) {
			tbSelMonth.options[iMon-1].selected = true;
			if (iYear>giEndYear) iYear=giEndYear;
			if (iYear<giStartYear) iYear=giStartYear;
			tbSelYear.options[iYear-giStartYear].selected = true;
			fUpdateCal(iYear, iMon);
		}
		
		function fPrevMonth(){
		  var iMon = tbSelMonth.value;
		  var iYear = tbSelYear.value;
		
		  if (--iMon<1) {
			  iMon = 12;
			  iYear--;
		  }
		
		  fSetYearMon(iYear, iMon);
		}
		
		function fNextMonth(){
		  var iMon = tbSelMonth.value;
		  var iYear = tbSelYear.value;
		
		  if (++iMon>12) {
			  iMon = 1;
			  iYear++;
		  }
		
		  fSetYearMon(iYear, iMon);
		}
		
		// 오늘 날짜를 기준으로 달력 초기화.
		function initCalendar()	{
			var today  = new Date();
			var year   = today.getFullYear();
			var month  = today.getMonth()+1;

			// 매개변수로 컨트롤이 넘어온 경우 기존 값의 년월로 초기화.
			var tArguments = window.dialogArguments;
			if (tArguments && tArguments.dateObject)	{
				var pDate	= tArguments.dateObject.value;
					if(pDate.length==10)	{
						year	= parseInt(pDate.substr(0,4), 10);
						month	= parseInt(pDate.substr(5,2), 10);
					}
			}
			
			fSetYearMon(year, month);
		}
		
		var gMonths = new Array("01","02","03","04","05","06","07","08","09","10","11","12");
	//-->
	</script>
</head>

<body leftmargin=0 topmargin=0 onLoad="JavaScript:initCalendar();document.body.focus();" style="overflow:hidden;background-color:#ffffff;margin:0px;padding:0px;">
	<script language="JScript">
	<!-- 
		with (document) {

		write("<table width='207' border='0' cellpadding='0' cellspacing='0' class='d_event_calen'>");
		write("<tr>");
		write("  <td height='26' class='text_align_M' align='center' ><a href='javascript:fPrevMonth()' onfocus='blur()'><<</a> &nbsp;&nbsp;<span class='txt3'>");

		// 년도 출력.
		write("&nbsp;<SELECT name='tbSelYear' id='tbSelYear' onChange='fUpdateCal(tbSelYear.value, tbSelMonth.value)' Victor='Won'>");
		for(i=giStartYear;i<=giEndYear;i++)
			write("<OPTION value='"+i+"'>"+i+"</OPTION>");
		write("</SELECT>");
		// 월 출력.
		write("&nbsp;<SELECT name='tbSelMonth' id='tbSelMonth' onChange='fUpdateCal(tbSelYear.value, tbSelMonth.value)' Victor='Won'>");
		for (i=0; i<12; i++)
			write("<option value='"+(i+1)+"'>"+gMonths[i]+"</option>");
		write("</SELECT>");

		write("  </span> &nbsp; <a href='javascript:fNextMonth()'>>></a></td>");
		write("</tr>");
		//write("<tr>");
		//write("  <td height='22' valign='top' style='padding-left:7px;'><img src='/images/event/calendar_02.gif' width='191' height='20'></td>");
		//write("</tr>");
		write("<TR>");
		write("<td align='center'>");
		write("<table class='C_ver9' border='0' cellpadding='2'>");
		fDrawCal(giYear, giMonth, 26, 10);
		write("</table>");
		write("</td>");
		write("</TR>");

		write("<tr>");
		write("	<td class='calendar_bar' height='3px'></td>");
		write("</tr>");
		write("<tr>");
		write("	<td height='24' align='right' style='padding-right:5px;font-weight: bold;'><a href='javascript:parent.hideCalendar();'>close</a></td>");
		write("</tr>");
		
		write("</TD></TR><TR></TR>");
		write("</TABLE>");
		}
		// End -- Coded by Liming Weng, email: victorwon@netease.com 
	//-->
	</script>
</body>
</html>

