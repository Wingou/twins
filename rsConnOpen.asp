<%
Dim cnx
Dim rs
Set cnx = Server.CreateObject("ADODB.Connection") 
cnx.Open ("Provider= Microsoft.ACE.OLEDB.12.0;Data Source="& Server.MapPath("../DAL/twins.accdb")&"") 
%>