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