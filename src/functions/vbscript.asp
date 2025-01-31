<%
sub ListFolderContents(path_, sourceFolder_, firstSourceFolderId_)
        dim fs_, folder_, item_, url_
        set fs_ = CreateObject("Scripting.FileSystemObject")



    if fs_.FolderExists(path_) then

            set folder_ = fs_.GetFolder(path_)

            for each item_ in folder_.SubFolders
                ListFolderContents item_, sourceFolder_, firstSourceFolderId_
            next
            for each x in folder_.Files
                folder=x.ParentFolder
                folder = replace(folder, "'", "''")
                
                name=x.Name
                name = replace(name, "'", "''")
                
                fileDate=x.DateLastModified
                fileSize=x.Size
                fileType=x.Type
                
                sql="INSERT INTO file (name, folder, id_source, fileDate, fileSize, fileType) VALUES ( '"&name&"', '"&folder&"', "&firstSourceFolderId_&", '"&fileDate&"', "&fileSize&", '"&fileType&"');"
                cnx.execute(sql)
                sql="SELECT @@IDENTITY AS fileId;"
                Set rs = cnx.execute(sql)


                incrementCpt
                updateLabelCtp "idLabelCptLecture", (relativeFolderForCpt(x.Path, sourceFolder_))
            next
    end if
end sub


sub FindDuplicateInCompareFolder(path_, compareFolder_, firstFolderId_)
        dim fs_, folder_, item_, url_
        set fs_ = CreateObject("Scripting.FileSystemObject")

    if fs_.FolderExists(path_) then

            set folder_ = fs_.GetFolder(path_)

            for each item_ in folder_.SubFolders
                FindDuplicateInCompareFolder item_, compareFolder_, firstFolderId_
            next
            for each x in folder_.Files
                folder = x.ParentFolder
                folder = replace(folder, "'", "''")

                name=x.Name
                name = replace(name, "'", "''")
                
                fileDate=x.DateLastModified
                fileSize=x.Size
                fileType=x.Type
                
                sql="SELECT count(id) FROM file WHERE name='"&name&"' AND id_source="&firstFolderId_&" AND fileSize="&fileSize&";"
                'response.write "<br>"&sql
                Set rs = cnx.execute(sql)
                incrementCptRecherche
                'response.write "<br>count = "&rs(0)
                if rs(0)>0 then
                    toFolder = replace(folder, compareFolder, saveFolder)
                    fromFullName = folder&"\"&name
                    toFullName = toFolder&"\"&replace(name, " ", "_")
                    setValidFolder toFolder, saveFolder
                    copier fromFullName, toFullName
                    incrementCptDoublon
                end if

            next
    end if
end sub

function fileRename(f__)            
    Randomize Timer
    indRandom=int(Rnd*99999)+1
    b=split (f__, ".")
    last=b(Ubound(b))
    fileRename = replace(f__, last, "-Copie("&indRandom&")."&last)
end function

function setValidFolder(toFolder_, copieDirDoublon_)
        singleFolders = split(replace(toFolder_, copieDirDoublon_ , ""), "\")
        currentFolder=copieDirDoublon_
        for iFolder=0 to UBound(singleFolders)
            singleFolder=singleFolders(iFolder)
            if singleFolder<>"" then
                currentFolder=currentFolder&"\"&singleFolders(iFolder)
                if Not fs.FolderExists(currentFolder) then
                          fs.CreateFolder(currentFolder)
                end if
            end if
        next
end function

function copier(from_, to_)
        if fs.FileExists(to_) then to_ = fileRename(to_)

       ' response.write "<br>from: "&from_
       ' response.write "<br>to: "&to_

        if fs.FileExists(from_) then fs.MoveFile from_, to_

        if fs.FolderExists(from_) then
            if fs.SubFolders(from_).Files.Count=0 and fs.SubFolders(from_).SubFolders.Count=0 then
                    fs.SubFolders(from_) = replace(from_, last, "xxxxxxxxxxxxxxxxxxxxxxxxxx")
                    end if
        end if
end function

function log(logLabel_, logVal)
    response.write "<hr><b>"&logLabel_&"</b>:"&logVal&"<hr>"
end function


function displayTrueTwins (tpMin, tpMax, title, idSource_)

    sql="SELECT name, folder FROM file WHERE template in ("&tpMin&","&tpMax&") AND id_source="&idSource_&" ORDER BY name, template"
    set rs=cnx.execute(sql)

    i=0
    WHILE not rs.EOF 
    i=i+1
        t_NameTrueTwinsSameName(i)=rs("name")
        t_FolderTrueTwinsSameName(i)=rs("folder")
    rs.MoveNext
    WEND
    iTrueTwinsSameName=i

    
    if i>0 then
        response.write "<hr><b>TRUE TWINS - "&title&"</b><hr>"
        for i=1 to iTrueTwinsSameName
            name=t_NameTrueTwinsSameName(i)
            fullFolder=t_FolderTrueTwinsSameName(i)
            if i>1 then
                if name=t_NameTrueTwinsSameName(i-1) then
                    name=""
                end if
            end if  
            viewTrueTwinsFiles name, fullFolder
        next
    end if

end function


function fillDatabaseReturnIdSource (sourceFolder_)

        ' Vérif si l'on a le dossier en base
        sql="SELECT COUNT(file.id) AS nbFiles FROM source INNER JOIN file ON source.id = file.id_source WHERE source.folder='"&sourceFolder_&"';"
        Set rs = cnx.execute(sql)        
        nbFiles = rs("nbFiles")

        if nbFiles=0 then
            ' Si la base est vide, on la remplit + Récup id_source
            sql="INSERT INTO source (folder) VALUES ('"&sourceFolder_&"');"
            cnx.execute(sql)
            sql="SELECT @@IDENTITY AS sourceId;"
            Set rs = cnx.execute(sql)
            id_source = rs.Fields("sourceId").value

            i=0
            ListFolderContents sourceFolder_,sourceFolder_,id_source
            iMax=i
        else
            ' Si la base est remplie,récup id_source
            sql="SELECT id AS sourceId FROM source WHERE folder='"&sourceFolder_&"';"
            Set rs = cnx.execute(sql)
            id_source = rs.Fields("sourceId").value
        end if

        fillDatabaseReturnIdSource = id_source

end function
%>