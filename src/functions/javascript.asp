<%
function incrementCpt
%>
    <script language=javascript>
    document.getElementById("idCptLecture").value++;
    </script>
<%
end function

function incrementCptRecherche
%>
    <script language=javascript>
    document.getElementById("idCptRecherche").value++;
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


%>