<%
Function Dico(valueToTranslate)

    Select Case valueToTranslate
    Case "OUi"
        translatedValue = "Y"
    Case else
        translatedValue=valueToTranslate
    End Select

    Dico=translatedValue

End Function
%>