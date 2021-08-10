<% response.Buffer=false %>
<% Server.ScriptTimeout = 1000000 %> 

<!--#include file="functions.asp"-->
<!--#include file="view.asp"-->

<HTML>
<HEAD>
    <meta charset="UTF-8">
    <Title>TWINS - Recherche de doublons</Title>
</HEAD>
<BODY style="font-size:12px;font-family:arial">
<%
set fs=Server.CreateObject("Scripting.FileSystemObject")

DIM t_Date(1000000)
DIM t_Size(1000000)
DIM t_Type(1000000)
DIM t_Name(1000000)
DIM t_Folder(1000000)

sourceFolder=request("sourceFolder")
saveFolder=request("saveFolder")

viewTitle
viewGetFolder sourceFolder, saveFolder
viewCpt(sourceFolder)

i=0
ListFolderContents sourceFolder,sourceFolder
iMax=i

response.write "<hr>"
hasDoublon=false
for i=1 to iMax
    sourceFullName = t_Folder(i)&"\"&t_Name(i)      
    incrementCptRecherche(i)
    updateLabelCtp "idLabelCptRecherche", (relativeFolderForCpt(sourceFullName, sourceFolder))
                     
    for j=1+i to iMax
        if t_Date(i)=t_Date(j) AND t_Size(i)=t_Size(j) AND t_Type(i)=t_Type(j) AND (NOT (t_Name(i)=t_Name(j) AND t_Folder(i)=t_Folder(j))) then

            fromFullName = t_folder(j)&"\"&t_Name(j)
            
            toFolder = replace(t_folder(j), sourceFolder, saveFolder)
            toFullName = toFolder&"\"&replace(t_Name(j), " ", "_")

            displaySourceFullName = replace(sourceFullName, sourceFolder, "<b>"&sourceFolder&"</b>")
            displayFromFullName =replace(fromFullName, sourceFolder, "<b>"&sourceFolder&"</b>")
            displayToFullName =replace(toFullName, saveFolder, "<b>"&saveFolder&"</b>")

            viewTwinsFiles displaySourceFullName, displayFromFullName, displayToFullName

            setValidFolder toFolder, saveFolder
            copier fromFullName, toFullName
            incrementCptDoublon
            updateLabelCtp "idLabelCptDoublon", (relativeFolderForCpt (sourceFullName, sourceFolder))

            hasDoublon=true

        end if
    next
next

set fo=nothing
set fs_=nothing
set fs=nothing

if sourceFolder<>"" then
    if hasDoublon=false then
        resultMsg="Aucun doublon !"
    else
        resultMsg="Recherche OK !"
    end if
    response.write "<hr>"&resultMsg&"<hr>"
end if
%>
</BODY>
</HTML>