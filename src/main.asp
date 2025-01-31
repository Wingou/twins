<!--#include file="views/common.asp"-->
<!--#include file="functions.asp"-->
<%
    sourceFolder=request("sourceFolder")
    saveFolder=request("saveFolder")
    compareFolder=request("compareFolder")

    viewTitle
    viewGetFolder sourceFolder, saveFolder, compareFolder

    if sourceFolder<>"" and compareFolder="" then
%>
    <!--#include file="views/duplicates.asp" -->
    <!--#include file="findDuplicates.asp"-->
<%
    end if

    if sourceFolder<>"" and compareFolder<>"" then
%>
  <!--#include file="compare.asp"-->
<%
    end if
%>


