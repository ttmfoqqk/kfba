<!-- #include file = "../_lib/header.asp" -->
<%
Session("UserIdx")	= ""
Session("UserId")	= ""
Session("UserName")	= ""

Session.Contents.RemoveAll()
Session.Abandon()

response.redirect FRONT_ROOT_DIR
%>