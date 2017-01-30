<% 
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: get 2 news records
%>
       
<%
' gets news

if pShowNews="-1" then

%>

<tr><td colspan=2>&nbsp;</td></tr>
<tr>
<td colspan=2>

 <br><b><%=getMsg(723,"Top news")%></b><%
 
 pNewsCount=0

 mySQL="SELECT newsText, newsDate FROM storeNews WHERE idStore="& pIdStore& " ORDER BY idStoreNews DESC"

 call getFromDatabase(mySql, rstemp,"index")

do while not rstemp.eof and pNewsCount<2

%>
<br>• <%=rstemp("newsDate")%> &nbsp; <%=rstemp("newsText")%>
      
<%
 pNewsCount=pNewsCount+1
 rstemp.movenext
loop
end if
%>
</td></tr>