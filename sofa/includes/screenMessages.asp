<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: get screen message from database
%>
<%
function getMsg(pIdScreenMessage, pMem)  
 
 ' check if the product is inside the cart
 mySQL="SELECT screenMessage FROM screenMessages WHERE idScreenMessage=" &pIdScreenMessage& " AND idStore=" &pIdStore
 
 call getFromDatabase(mySQL, rstempMsg, "screenMessages")  

 if not rstempMsg.eof then
  getMsg=rstempMsg("screenMessage")
 else
  getMsg=""
 end if  
 
end function

%>