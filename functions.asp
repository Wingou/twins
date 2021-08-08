<%
sub ListFolderContents(path_)
    dim fs_, folder_, item_, url_
    set fs_ = CreateObject("Scripting.FileSystemObject")
    set folder_ = fs_.GetFolder(path_)

    for each item_ in folder_.SubFolders
        ListFolderContents(item_)
    next

    for each x in folder_.Files
    
                i=i+1
                t_Folder(i)=x.ParentFolder
                t_Name(i)=x.Name

                t_Date(i)=x.DateLastModified
                t_Size(i)=x.Size
                t_Type(i)=x.Type

                t_Uniq(i)=True

                incrementCpt(i)
    next
end sub

function parentFolder(f_)
    if f_<>"" then
        a=split (f_, "\")
        aa=a(0)
        for k=1 to Ubound(a)-1
            aa=aa&"\"&a(k)
        next
        parentFolder=aa
    end if
end function

function fileRename(f__)
    Randomize Timer
    indRandom=int(Rnd*99999)+1
    b=split (f__, ".")
    last=b(Ubound(b))
    fileRename = replace(f__, last, "-Copie("&indRandom&")."&last)
end function

function copier(copieFolder, copieName)

        destFolder = replace(copieFolder, originFolder, dirDoublon)
        doublonName = replace(copieName, " ", "_")
        destFullName=destFolder&"\"&doublonName

        destParentFolder = parentFolder(replace(copieFolder, originFolder, dirDoublon))

        sourceFolder = copieFolder
        sourceFullName = sourceFolder&"\"&copieName

        if Not fs.FolderExists(destFolder) and destFolder<>parentOriginFolder Then  
                if Not fs.FolderExists(destParentFolder) and destParentFolder<>parentOriginFolder Then  
                        'response.write "<br>parentFolder created : " & destParentFolder
                        fs.CreateFolder(destParentFolder)
                end if
                'response.write "<br>folder created : "& destFolder
                fs.CreateFolder (destFolder)
        end if

        response.write "source : "&sourceFullName
        response.write "<br>"
        response.write "dest : "&destFullName


        if fs.FileExists(destFullName) then
            destFullName = fileRename(destFullName)
        end if
        fs.MoveFile sourceFullName, destFullName

        response.write "<hr>"
end function


function incrementCpt(iCptVal)
%>
    <script language=javascript>
    document.getElementById("cpt").value='<%=iCptVal%>';
    </script>
<%
end function

function incrementCptDoublon
%>
    <script language=javascript>
    var valStr=document.getElementById("cptDoublon");
    var valNum=parseInt(valStr.value, 10)+1;
    valStr.value=valNum
    </script>
<%
end function

function incrementCptRecherche(iCptRechVal)
%>
    <script language=javascript>
    document.getElementById("cptTest").value='<%=iCptRechVal%>';
    </script>
<%
end function
%>