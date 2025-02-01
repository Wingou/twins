<%

choix=request("choix")
sourceFolder=request("sourceFolder")
deltaFolder=request("deltaFolder")
saveFolder=request("saveFolder")

if choix="reset" then
    sql="DELETE FROM file"
    Cnx.execute sql

    sql="DELETE FROM source"
    Cnx.execute sql
    
    <!-- #include file="../rsConnClose.asp"-->
    response.redirect "index.asp?sourceFolder="&sourceFolder&"&deltaFolder="&deltaFolder&"&saveFolder="&saveFolder&"&choix=resetDone"
end if

%>