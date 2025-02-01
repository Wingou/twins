<%
set fs=Server.CreateObject("Scripting.FileSystemObject")
        viewCpt(sourceFolder)
        id_source=fillDatabaseReturnIdSource(sourceFolder)
        FindDuplicateInDeltaFolder deltaFolder, deltaFolder, id_source
        FindEmptyFolders deltaFolder, deltaFolder, id_source
set fs=nothing
%>