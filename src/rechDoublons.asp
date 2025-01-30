

<HTML>
<HEAD>
    <meta charset="UTF-8">
    <Title>TWINS - Recherche de doublons</Title>
</HEAD>
<BODY style="font-size:12px;font-family:arial">
<%
set fs=Server.CreateObject("Scripting.FileSystemObject")

DIM t_Id(1000000)
DIM t_Date(1000000)
DIM t_Size(1000000)
DIM t_Type(1000000)
DIM t_Name(1000000)
DIM t_Folder(1000000)

DIM t_NameTrueTwinsSameName(1000000)
DIM t_FolderTrueTwinsSameName(1000000)

DIM t_NameTrueTwinsDiffName(1000000)
DIM t_FolderTrueTwinsDiffName(1000000)

sourceFolder=request("sourceFolder")
saveFolder=request("saveFolder")

viewTitle
viewGetFolder sourceFolder, saveFolder
viewCpt(sourceFolder)


        ' Vérif si l'on a le dossier en base
        sql="SELECT COUNT(file.id) AS nbFiles FROM source INNER JOIN file ON source.id = file.id_source WHERE source.folder='"&sourceFolder&"';"
        Set rs = cnx.execute(sql)        
        nbFiles = rs("nbFiles")

        if nbFiles=0 then
            ' Si la base est vide, on la remplit + Récup id_source
            sql="INSERT INTO source (folder) VALUES ('"&sourceFolder&"');"
            cnx.execute(sql)
            sql="SELECT @@IDENTITY AS sourceId;"
            Set rs = cnx.execute(sql)
            id_source = rs.Fields("sourceId").value

            i=0
            ListFolderContents sourceFolder,sourceFolder,id_source
            iMax=i
        else
            ' Si la base est remplie,récup id_source
            sql="SELECT id AS sourceId FROM source WHERE folder='"&sourceFolder&"';"
            Set rs = cnx.execute(sql)
            id_source = rs.Fields("sourceId").value
        end if

        ' on identifie en base les doublons
        'sql = "SELECT id, name, folder, fileDate, fileSize, fileType from file WHERE id_source="&id_source&" "

        sql="SELECT id, name, folder, fileDate, fileSize, fileType "
        sql=sql&" FROM file "
        sql=sql&" WHERE filesize in (SELECT filesize FROM file WHERE id_source="&id_source&" GROUP BY filesize HAVING count(filesize)>1) "
        sql=sql&" AND id_source="&id_source&" "
        sql=sql&" ORDER BY filesize, name, folder "
        set rsFile = cnx.execute(sql)

        i=0
        WHILE not rsFile.EOF
            i=i+1
            t_Id(i) = rsFile("id")
            t_Name(i) = rsFile("name")
            t_Folder(i) = rsFile("folder")
            t_Date(i) = rsFile("fileDate")
            t_Size(i) = rsFile("fileSize")
            t_Type(i) = rsFile("fileType")
        rsFile.moveNext
        WEND
        iMax=i
        rsFile.close


response.write "<hr>"

hasDoublon=false
%>

<table>
<tr>
        <td><b>Taille</b></td>
        <td><b>Fichier</b></td>
        <td><b>Dossier</b></td>
        <td><b>Aperçu</b></td>
</tr>
<%

for i=1 to iMax
    sourceFullName = t_Folder(i)&"\"&t_Name(i)      
    incrementCptRecherche(i)
    updateLabelCtp "idLabelCptRecherche", (relativeFolderForCpt(sourceFullName, sourceFolder))

    folderPath = t_Folder(i)
    folderPath = replace(folderPath, "H:\Principal\GO\Photos\","")
    folderPath = replace(folderPath, "\","/")
    srcPreview = "http://localhost/photos/"&folderPath&"/"&t_Name(i)
            

    %>

    <tr>
        <td><%=t_Size(i)%></td>
        <td><%=t_Name(i) %></td>
        <td><%=t_Folder(i)%></td>
        <td><img src="<%=srcPreview%>" style="width:50px;height:50px;"></td>
    </tr>
    <%

    for j=1+i to iMax
        if t_Date(i)=t_Date(j) AND t_Size(i)=t_Size(j) AND t_Type(i)=t_Type(j) AND (NOT (t_Name(i)=t_Name(j) AND t_Folder(i)=t_Folder(j))) then

            fromFullName = t_folder(j)&"\"&t_Name(j)
            
            toFolder = replace(t_folder(j), sourceFolder, saveFolder)
            toFullName = toFolder&"\"&replace(t_Name(j), " ", "_")

            displaySourceFullName = replace(sourceFullName, sourceFolder, "<b>"&sourceFolder&"</b>")
            displayFromFullName =replace(fromFullName, sourceFolder, "<b>"&sourceFolder&"</b>")
            displayToFullName =replace(toFullName, saveFolder, "<b>"&saveFolder&"</b>")

            ' viewTwinsFiles displaySourceFullName, displayFromFullName, displayToFullName

            setValidFolder toFolder, saveFolder
            copier fromFullName, toFullName
            incrementCptDoublon
            updateLabelCtp "idLabelCptDoublon", (relativeFolderForCpt (sourceFullName, sourceFolder))

            hasDoublon=true
        end if
        if t_Size(i)=t_Size(j) AND t_Type(i)=t_Type(j) AND t_Name(i)=t_Name(j) AND (NOT (t_Folder(i)=t_Folder(j))) then
            sql="UPDATE file set template=10 WHERE id="&t_Id(i)&";"
            cnx.execute(sql)
            
            sql="UPDATE file set template=15 WHERE id="&t_Id(j)&";"
            cnx.execute(sql)
        end if
        if t_Size(i)=t_Size(j) AND t_Type(i)=t_Type(j) AND (NOT (t_Name(i)=t_Name(j) OR t_Folder(i)=t_Folder(j))) then
            sql="UPDATE file set template=20 WHERE id="&t_Id(i)&";"
            cnx.execute(sql)
        
            sql="UPDATE file set template=25 WHERE id="&t_Id(j)&";"
            cnx.execute(sql)
        end if
    next
next
%>
</table>
<%
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



displayTrueTwins 10, 15, "Same Name", id_source

displayTrueTwins 20, 25, "Diff Name", id_source




%>
</BODY>
</HTML>
