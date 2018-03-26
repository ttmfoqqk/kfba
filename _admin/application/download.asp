<%
Dim pach : pach = Server.MapPath( Request("pach") )
Dim file : file = Request("file")
Dim name : name = Request("name")

Response.ContentType = "application/unknown"
Response.AddHeader "Content-Disposition","attachment; filename=" & name 
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type = 1
objStream.LoadFromFile pach &"\"&  file
download = objStream.Read
Response.BinaryWrite download 
Set objstream = nothing 
%>
