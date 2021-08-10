<%
sub ListFolderContents(path_, sourceFolder_)
        dim fs_, folder_, item_, url_
        set fs_ = CreateObject("Scripting.FileSystemObject")
    if fs_.FolderExists(path_) then 
        set folder_ = fs_.GetFolder(path_)
            for each item_ in folder_.SubFolders
                ListFolderContents item_,sourceFolder_
            next
            for each x in folder_.Files
                i=i+1
                t_Folder(i)=x.ParentFolder
                t_Name(i)=x.Name

                t_Date(i)=x.DateLastModified
                t_Size(i)=x.Size
                t_Type(i)=x.Type

                incrementCpt(i)
                
                updateLabelCtp "idLabelCptLecture", (relativeFolderForCpt(x.Path, sourceFolder_))
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

        if fs.FileExists(to_) then
            to_ = fileRename(to_)
        end if
        fs.MoveFile from_, toFullName

end function

function incrementCpt(iCptVal)
%>
    <script language=javascript>
    document.getElementById("idCptLecture").value='<%=iCptVal%>';
    </script>
<%
end function

function incrementCptRecherche(iCptRechVal)
%>
    <script language=javascript>
    document.getElementById("idCptRecherche").value='<%=iCptRechVal%>';
    </script>
<%
end function

function incrementCptDoublon
%>
    <script language=javascript>
    var valStr=document.getElementById("idCptDoublon");
    var valNum=parseInt(valStr.value, 10)+1;
    valStr.value=valNum
    </script>
<%
end function

function relativeFolderForCpt(cptVal__, sourceFolder__)
    cptVal_ = cptVal__
    cptVal_=replace(cptVal_, sourceFolder__, "")
    cptVal_=replace(cptVal_, "\", "\\")
    relativeFolderForCpt = cptVal_
end function

function updateLabelCtp( idCpt_, cptNewVal_)
%>
    <script language=javascript>
    document.getElementById('<%=idCpt_%>').value='<%=cptNewVal_%>';
    </script>
<%
end function

function log(logLabel_, logVal)
    response.write "<hr><b>"&logLabel_&"</b>:"&logVal&"<hr>"
end function
%>