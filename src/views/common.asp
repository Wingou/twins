<%
function viewTitle
%>
    <center><font style="font-size:1rem;"><b>TWINS</b></font>
    <br><font style="font-size:0.8rem;">Find duplicated media by WTL</font>
    <br><font style="font-size:0.6rem;">08/08/2021 - last update : 31/01/2025</font></center>
<%
end function

function viewGetFolder(initSourceFolder_, initSaveFolder_, initDeltaFolder_)
    
    if initSaveFolder_="" then initSaveFolder_="C:\Users\WingQiqi\Desktop\doublons"
    placeholderSourceFolder="Enter the SOURCE path"
    placeholderDeltaFolder="Enter the path of the DELTA folder. Don't if the media are all in the SOURCE Folder"
    %>
    <hr>
        <form method=GET action="index.asp">
        <div style="display:flex;flex-direction:column">
            <div style="display:flex;flex-direction:row">
                <div style="width:8rem;">
                    Source Folder
                </div>
                <div style="flex:auto">
                    <input type="text" name="sourceFolder" placeholder="<%=placeholderSourceFolder%>" value="<%=initSourceFolder_%>" style="width:100%;">
                </div>
            </div>
            <div style="display:flex;flex-direction:row">
                <div style="width:8rem;">
                    Delta Folder
                </div>
                <div style="flex:auto">
                    <input type="text" name="deltaFolder" placeholder="<%=placeholderDeltaFolder%>" value="<%=initDeltaFolder_%>" style="width:100%;">
                </div>
            </div>
            <div style="display:flex;flex-direction:row">
                <div style="width:8rem;">
                    Backup Folder
                </div>
                <div style="flex:auto">
                    <input type="text" name="saveFolder" value="<%=initSaveFolder_%>" style="width:100%;">
                </div>
            </div>
            <div style="display:flex;flex-direction:row">
                <div style="width:8rem;">
                </div>
                <div style="flex:auto">
                    <input type="submit" value="GO"> <font style="font-size: 0.7rem;"> <i>IUSR account MUST have a total access on these folders</i></font>
                     - <a href="index.asp?sourceFolder=<%=sourceFolder%>&deltaFolder=<%=deltaFolder%>&saveFolder=<%=saveFolder%>&choix=reset">RESET</a> -
                </div>
            </div>
        </div>
    </form>

<%
end function
%>