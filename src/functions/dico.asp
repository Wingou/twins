<%
Function Dico(valueToTranslate)

    Select Case valueToTranslate
    Case "OUI"
        translatedValue = "Y"
    Case else
        translatedValue=valueToTranslate
    End Select

    Dico=translatedValue

End Function
%>