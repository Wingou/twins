<%
function viewCpt(viewCptFolder_)
    if viewCptFolder_<>"" then
%>
    <hr>
    <div style="display:flex;flex-direction:column">
        <div style="display:flex;flex-direction:row">
            <div style="width:5rem;">Lecture</div>
            <div><input style="width:5rem;" id="idCptLecture" value="0" type="text"></div>
            <div style="flex:auto"><input style="width:100%;" id="idLabelCptLecture" value="" type="text"></div>
        </div>
        <div style="display:flex;flex-direction:row">
            <div style="width:5rem;">Recherche</div>
            <div><input style="width:5rem;" id="idCptRecherche" value="0" type="text"></div>
            <div style="flex:auto"><input style="width:100%;" id="idLabelCptRecherche" value="" type="text"></div>
        </div>
        <div style="display:flex;flex-direction:row">
            <div style="width:5rem;">Doublons</div>
            <div><input style="width:5rem;" id="idCptDoublon" value="0" type="text"></div>
            <div style="flex:auto"><input style="width:100%;" id="idLabelCptDoublon" value="" type="text"></div>
        </div>
    </div>

<%
    end if
end function

function viewTitle
%>
    <center><font style="font-size:1rem;"><b>TWINS - Recherche de doublons</b></font><br><font style="font-size:0.6rem;">08/08/2021 - WTL</font></center>
<%
end function

function viewGetFolder(initSourceFolder_, initSaveFolder_)
    if initSourceFolder_="" then initSourceFolder_="C:\XXX"
    if initToCompareFolder_="" then initToCompareFolder_="C:\XXX"
    if initSaveFolder_="" then initSaveFolder_="C:\Users\WingQiqi\Desktop\doublons"
%>
    <hr>
        <form method=GET action="index.asp">
        <div style="display:flex;flex-direction:column">
            <div style="display:flex;flex-direction:row">
                <div style="width:8rem;">
                    Source Folder
                </div>
                <div style="flex:auto">
                    <input type="text" name="sourceFolder" value="<%=initSourceFolder_%>" style="width:100%;">
                </div>
            </div>
            
            
            
            <div style="display:flex;flex-direction:row">
                <div style="width:8rem;">
                    <input type="checkbox" name="toCompareFolder" value="<%=initToCompareFolder_%>">To Compare Folder
                </div>
                <div style="flex:auto">
                    <input type="text" name="toCompareFolder" value="<%=initToCompareFolder_%>" style="width:100%;">
                </div>
            </div>
            <div style="display:flex;flex-direction:row">
                <div style="width:8rem;">
                    Sauvegarde
                </div>
                <div style="flex:auto">
                    <input type="text" name="saveFolder" value="<%=initSaveFolder_%>" style="width:100%;">
                </div>
            </div>
            <div style="display:flex;flex-direction:row">
                <div style="width:8rem;">
                </div>
                <div style="flex:auto">
                    <input type="submit" value="    GO    "> <font style="font-size: 0.7rem;"> <i>le profil IUSR doit avoir un accès total sur ces dossiers</i></font>
                </div>
            </div>
        </div>
    </form>

<%
end function

function viewTwinsFiles(displaySourceFullName_, displayFromFullName_, displayToFullName_)
%>
            <div style="display:flex; flex-direction:column;">
                <div style="display:flex; flex-direction:row;">
                    <div style="width:5rem;;">
                        Original
                    </div>
                    <div style="flex:auto;">
                        <%=displaySourceFullName_ %>
                    </div>
                </div>
                <div style="display:flex; flex-direction:row;">
                    <div style="width:5rem;;">
                        Doublon
                    </div>
                    <div style="flex:auto;">
                        <%=displayFromFullName_ %>
                    </div>
                </div>
                <div style="display:flex; flex-direction:row;">
                    <div style="width:5rem;;">
                        Sauvegarde
                    </div>
                    <div style="flex:auto;">
                        <%=displayToFullName_ %>
                    </div>
                </div>
                <div>
                    <div>
                        <hr>
                    </div>
                </div>
            </div>
<%
end function


function viewTrueTwinsFiles(displayName_, displayFullFolder_)
%>
            <div style="display:flex; flex-direction:column;">
                <div style="display:flex; flex-direction:row;">
                    <div style="width:15rem;">
                        <%=displayName_%>
                    </div>
                    <div style="flex:auto">
                        <%=displayFullFolder_ %>
                    </div>
                </div>
                    
                        
                
            </div>
            <% if displayName_="" then %>
            <hr>
            <% end if %>
<%
end function
%>