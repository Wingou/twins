<%
set fs=Server.CreateObject("Scripting.FileSystemObject")

        viewCpt(sourceFolder)
        id_source=fillDatabaseReturnIdSource(sourceFolder)

        FindDuplicateInCompareFolder compareFolder, compareFolder, id_source

'set fo=nothing
'set fs_=nothing
set fs=nothing

%>