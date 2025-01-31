
<%

choix=request("choix")

if choix="delete" then
    sql="DELETE FROM file"
    Cnx.execute sql

    sql="DELETE FROM source"
    Cnx.execute sql

    
    <!-- #include file="../rsConnClose.asp"-->
    response.redirect "index.asp"
end if

%>