
<%if convertGetal(page.iID)<>0 then 
if convertGetal(page.iListPageID)<>0 then
response.write "<div style=""float:right"">" & getArtLink(C_DIRECTORY_QUICKERSITE & "/default.asp?iID=" & encrypt(page.iListPageID) & "&amp;item=" & encrypt(page.iId) & "#" & encrypt(page.iId),l("preview"),"","_blank","") & "</div>"
else 
response.write "<div style=""float:right"">"
if secondAdmin.bSetupPageElements then
'response.write getArtLink("bs_editPageBlocksIF.asp?iPageID=" & encrypt(page.iID),"Edit Page Blocks","","_blank","example1") & "&nbsp;"
end if
response.write getArtLink(C_DIRECTORY_QUICKERSITE & "/default.asp?iID=" & encrypt(page.iID),l("preview"),"","_blank","") & "</div>"
end if
end if%>
