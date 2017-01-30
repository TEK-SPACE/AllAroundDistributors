<% 
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: get suppliers list
%>
<%
dim rsTempSup

pShowSuppliersList		= getSettingKey("pShowSuppliersList")

if pShowSuppliersList="-1" then

 mySQL = "SELECT DISTINCT suppliers.idSupplier, supplierName FROM suppliers, products WHERE suppliers.idSupplier=products.idSupplier AND products.idStore=" &pIdStore& " ORDER BY supplierName"
 call getFromDatabase (mySql, rsTempSup, "getSuppliers")

 do while not rsTempSup.eof%>  
	
  <p class="b01"><img src="images/e02.gif" width="6" height="5" alt="" border="0" align="absmiddle">&nbsp;&nbsp;<a href="comersus_listItems.asp?idSupplier=<%=rsTempSup("idSupplier")%>"><%=rsTempSup("supplierName")%></a></p>
  <div align="center"><img src="images/hr01.gif" width="137" height="3" alt="" border="0"></div>
  <%
  rsTempSup.movenext
 loop

end if
%>
