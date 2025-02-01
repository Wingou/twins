<!--#include file="views/common.asp"-->
<!--#include file="functions.asp"-->
<%
    choix=request("choix")
    sourceFolder=request("sourceFolder")
    saveFolder=request("saveFolder")
    deltaFolder=request("deltaFolder")

    viewTitle
    viewGetFolder sourceFolder, saveFolder, deltaFolder

    if sourceFolder<>"" and deltaFolder="" and choix="" then
%>
    <!--#include file="views/duplicates.asp" -->
    <!--#include file="findDuplicates.asp"-->
<%
    end if

    if sourceFolder<>"" and deltaFolder<>"" and choix="" then
%>
    <!--#include file="findDelta.asp"-->
<%
    end if
%>


