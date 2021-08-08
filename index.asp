<% response.Buffer=false %>
<% Server.ScriptTimeout = 1000000 %> 

<!--#include file="functions.asp"-->
<!--#include file="viewCpt.asp"-->

<HTML>
<HEAD Title="Recherche de doublons">
</HEAD>
<BODY style="font-size:12px;font-family:arial">

<%
set fs=Server.CreateObject("Scripting.FileSystemObject")

DIM t_Date(1000000)
DIM t_Size(1000000)
DIM t_Type(1000000)
DIM t_Name(1000000)
DIM t_Folder(1000000)
DIM t_Uniq(1000000)

parentOriginFolder=parentFolder(originFolder)
originFolder=request("folder")

dirDoublon = "C:\Users\WingQiqi\Desktop\doublons"

response.write "<center><font style='font-size:1rem;'><b>TWINS - Recherche de doublons</b></font><br><font style='font-size:0.6rem;'>08/08/2021 - WTL</font></center>"
response.write "<hr>Analyse du dossier <b>"&originFolder&"</b><hr>"

viewCpt

i=0
ListFolderContents(originFolder)
iMax=i

response.write "<hr>"
hasDoublon=false
for i=1 to iMax
    incrementCptRecherche(i)
    for j=1+i to iMax
        if t_Date(i)=t_Date(j) AND t_Size(i)=t_Size(j) AND t_Type(i)=t_Type(j) AND (NOT (t_Name(i)=t_Name(j) AND t_Folder(i)=t_Folder(j))) AND t_Uniq(i)=True then
            response.write "original : "&t_Folder(i)&"\"&t_Name(i)&"<br>"
            copier t_folder(j), t_Name(j) 
            incrementCptDoublon
            t_Uniq(j)=False
            hasDoublon=true
        end if
    next
next

set fo=nothing
set fs_=nothing
set fs=nothing

if hasDoublon=false then
    resultMsg="Aucun doublon !"
else
    resultMsg="Recherche OK !"
end if
response.write "<hr>"&resultMsg&"<hr>"
%>

</BODY>
</HTML>