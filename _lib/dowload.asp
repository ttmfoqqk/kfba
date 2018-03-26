<%
Dim pach : pach = Server.MapPath( Request("pach") )
Dim file : file = Request("file")

Response.ContentType = "application/unknown"
Response.AddHeader "Content-Disposition","attachment; filename=" & file 
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type = 1
objStream.LoadFromFile pach &"\"&  file
download = objStream.Read
Response.BinaryWrite download 
Set objstream = nothing 
%>
