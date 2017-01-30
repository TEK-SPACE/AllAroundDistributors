<%
' Comersus BackOffice 
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' February 2006
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->    
<!--#include file="../includes/databaseFunctions.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/settings.asp"--> 

<!--#include file="header.asp"--> 
<br><b>Import products</b>
<br>

<%
' on error resume next
dim mySQL, conntemp, rstemp

pIdStoreBackOffice	= getSessionVariable("idStoreBackOffice", 1)

mySQL="SELECT * FROM stores ORDER BY idStore"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_importUtilityForm.asp") 

%>
<br><br><b>File format details:</b><br>
- Record fields order:<br>
<i>
sku description details price btobprice imageUrl smallImageUrl 
listHidden cost sales weight length width height category supplier
emailText keywords clearance donation personalization</i>

<br><br>- Fields separator must be a tab

<br><br><b>Import file example:</b><br>
	sku1	Phone	Great Phone for you	26	22	phone.jpg	phonesmall.jpg	0	5	4	1	2	3	2	Cell Phones	Phone Supplier Inc	We will send your phone soon	phone, cell, sale	0	0	0			
                  <br>
                  <form name="selectform" method="post" action="comersus_backoffice_importProductsExec.asp">
	                <table width="100%">
                      <tr> 
                        <td width="202">Please select the Store</td> 
                        <td width="455"> 
                          <select name="idStore" size=1>
                            <%Do While Not rstemp.eof%> 
                            <option value="<%=rstemp("idStore")%>"
                    			<%if cInt(rstemp("idStore"))=cInt(pIdStoreBackOffice) then response.write "selected"%>
                    			><%=rstemp("StoreDescription")%></option>
                                                <%  rstemp.movenext 
                    		  Loop
                    		 %> 
                          </select>
                        </td>
                      </tr>
                      <tr> 
                        <td width="202">Check to mark the imported products as 
                          actives</td>
                        <td width="455"> 
                          <input type="checkbox" name="active">
                        </td>
                      </tr>
                      <tr> 
                        <td align="left" colspan=2>&nbsp;&nbsp; 
                          <TEXTAREA cols=80 rows=5 name="InfoToImport"></TEXTAREA>
                          <p align=left>
                          <input type=submit name="BtnTextField" value="Import from textfield">
              </p>
                        </td>



                      </tr>
                    </table>   
	</form>   

<!--#include file="footer.asp"--> 
<%call closeDb()%>