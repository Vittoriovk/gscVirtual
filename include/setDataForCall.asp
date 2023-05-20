<%
Session("CryptKey")=generatePassword(256)
sendData="IdAccount=" & IdAccount
sendData=CryptWithKey(sendData,Session("CryptKey"))
 
localVirtualPath = virtualPath 
if right(localVirtualPath,1)="\" or right(localVirtualPath,1)="/" then 
   localVirtualPath=mid(localVirtualPath,1,len(localVirtualPath)-1)
end if    
   
%>
<input type="hidden" name="localVirtualPath" id="localVirtualPath" value="<%=VirtualPath%>">
<input type="hidden" name="SendDataForCall"  id="SendDataForCall"  value="<%=sendData%>">
