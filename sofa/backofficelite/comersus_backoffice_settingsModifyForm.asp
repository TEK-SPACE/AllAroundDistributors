<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    

<% 
on error resume next 

dim mySQL, conntemp, rstemp

mySQL="SELECT idSettingGroup, idSetting, settingKey, settingValue, comments FROM settings WHERE settingKey<>'Wizard Test' ORDER by idSettingGroup, idSetting "

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_settingsform.asp") 

if rstemp.eof then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Empty settings table. ")
end if
%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"-->

<br><br><div class=title>Settings</div><br><br>
<i>Note: some settings are related to features included in Power Pack</i>
<form name="settings" method="post" action="comersus_backoffice_settingsModifyExec.asp">
  <table border=0>
	<%
    pPrevGroup =""
    do while not rstemp.eof
		if trim(pPrevGroup) <> trim(rstemp("idSettingGroup")) then 
			pPrevGroup = rstemp("idSettingGroup")
			
			Select Case pPrevGroup
				Case 1:
					pGroupName = "Company"	    
				Case 2:
					pGroupName = "Email"	    
				Case 3:
					pGroupName = "Shipment"	    
				Case 4:
					pGroupName = "Payment"	    
				Case 5:
					pGroupName = "PowerPack"	    
				Case 6:
					pGroupName = "Cart Behavior"	    
				Case 7:
					pGroupName = "Prefixes"	    
				Case 8:
					pGroupName = "Stock"	    
				Case 9:
					pGroupName = "Date"	    
				Case 10:
					pGroupName = "Currency"	    
				Case 11:
					pGroupName = "Customer"	    
				Case 12:
					pGroupName = "Fraud Prevention"	    
				Case 13:
					pGroupName = "Digital Goods"	    
				Case 14:
					pGroupName = "Taxes"	    
				Case 15:
					pGroupName = "Other"	    
					
			End Select
			pGroupName = pGroupName & " Settings"
			%>
			<TR> 
	         <TD colspan=3 width="450" height="21"> 
	         <br><b><font size=2><% response.write pGroupName %></font></b>
	         <hr>
	         </TD>      			
	        </TR>
	<%    	 
		end if	
  	
  	 	if rstemp("settingKey")="pStoreFrontDemoMode" _ 
    	or trim(rstemp("settingKey"))="pRunInstallationWizard" _
    	or trim(rstemp("settingKey"))="pCompareWithAmazon" _
    	or trim(rstemp("settingKey"))="pDisableState" _
    	or trim(rstemp("settingKey"))="pSuppliersList" _
    	or trim(rstemp("settingKey"))="pShowNews" _
    	or trim(rstemp("settingKey"))="pShowCategories" _
    	or trim(rstemp("settingKey"))="pShowSearchBox" _
    	or trim(rstemp("settingKey"))="pShowSuppliersList" _
    	or trim(rstemp("settingKey"))="pShowAddedToday" _
    	or trim(rstemp("settingKey"))="pAutogenerateNews" _
    	or trim(rstemp("settingKey"))="pEmailNoStock" _
    	or trim(rstemp("settingKey"))="pEnableCustomerTracking" _
    	or trim(rstemp("settingKey"))="pUseVatNumber" _
    	or trim(rstemp("settingKey"))="pPayPalExpressCheckout" _
    	or trim(rstemp("settingKey"))="pDateSwitch" _
    	or trim(rstemp("settingKey"))="pTaxIncluded" _
    	or trim(rstemp("settingKey"))="pRssFeedServer" _
    	or trim(rstemp("settingKey"))="pSecureStoreGraph"  then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <select  Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCompany" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCompanyAddress" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>	 
    	<%
    	elseif rstemp("settingKey")="pCompanyZip" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCompanyCity" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>	 
    	<%
    	elseif rstemp("settingKey")="pCompanyStateCode" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCompanyCountryCode" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCompanyPhone" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCompanyFax" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCompanyLogo" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCartQuantityLimit" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pMaxAddCartQuantity" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pOrderPrefix" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pCustomerPrefix" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pEmailAdmin" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>
    	<%
    	elseif rstemp("settingKey")="pEmailSender" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    
    	<%
    	elseif rstemp("settingKey")="pSmtpServer" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    
    	<%
    	elseif rstemp("settingKey")="pDebugEmail" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	        
    	<%
    	elseif rstemp("settingKey")="pEmailComponent" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			       <option value="none" <% if ucase(rstemp("settingValue")) = "NONE" then response.write "SELECTED" %>> NONE </OPTION>
			 	<option value="jmail" <% if ucase(rstemp("settingValue")) = "JMAIL" then response.write "SELECTED" %>> JMAIL </OPTION>
			 	<option value="jmailhtml" <% if ucase(rstemp("settingValue")) = "JMAILHTML" then response.write "SELECTED" %>> JMAIL HTML</OPTION>
			 	<option value="serverobjectsaspmail1" <% if ucase(rstemp("settingValue")) = "SERVEROBJECTSASPMAIL1" then response.write "SELECTED" %>> SERVEROBJECTSASPMAIL1 </OPTION> 
			 	<option value="serverobjectsaspmail2" <% if ucase(rstemp("settingValue")) = "SERVEROBJECTSASPMAIL2" then response.write "SELECTED" %>> SERVEROBJECTSASPMAIL2 </OPTION>
			 	<option value="persitsaspmail" <% if ucase(rstemp("settingValue")) = "PERSITSASPMAIL" then response.write "SELECTED" %>> PERSITSASPMAIL </OPTION> 			 	
			 	<option value="cdonts" <% if ucase(rstemp("settingValue")) = "CDONTS" then response.write "SELECTED" %>> CDONTS </OPTION> 			 	
			 	<option value="cdontshtml" <% if ucase(rstemp("settingValue")) = "CDONTSHTML" then response.write "SELECTED" %>> CDONTS HTML</OPTION> 			 	
			 	<option value="cdo" <% if ucase(rstemp("settingValue")) = "CDO" then response.write "SELECTED" %>> CDOSYS WIN-2003 </OPTION> 			 	
			 	<option value="bamboosmtp" <% if ucase(rstemp("settingValue")) = "BAMBOOSMTP" then response.write "SELECTED" %>> BAMBOOSMTP </OPTION> 			 	
			 	<option value="ocxmail" <% if ucase(rstemp("settingValue")) = "OCXMAIL" then response.write "SELECTED" %>> OCXMail </OPTION> 			 	
			 	<option value="aspsmartmail" <% if ucase(rstemp("settingValue")) = "ASPSMARTMAIL" then response.write "SELECTED" %>> AspSmartMail </OPTION> 			 	
			 	<option value="dundasmail" <% if ucase(rstemp("settingValue")) = "DUNDASMAIL" then response.write "SELECTED" %>> DundasMail </OPTION> 			 	
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	        
	        
	        
<%
    	elseif rstemp("settingKey")="pUnderStockBehavior" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="none" <% if ucase(rstemp("settingValue")) = "NONE" then response.write "SELECTED" %>> NONE </OPTION>
			 	<option value="dontadd" <% if ucase(rstemp("settingValue")) = "DONTADD" then response.write "SELECTED" %>> DONTADD </OPTION> 
			 	<option value="backorder" <% if ucase(rstemp("settingValue")) = "BACKORDER" then response.write "SELECTED" %>> BACKORDER </OPTION>			 	
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>  
	        	        
    	<%
    	elseif rstemp("settingKey")="pIdStore" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	     		 
    	<%
    	elseif rstemp("settingKey")="pShowStockView" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			

			 <td>
			 <select  Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pCurrencySign" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pDecimalSign" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			

			 <td>
			 <select  Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="." <% if rstemp("settingValue") = "." then response.write "SELECTED" %>> . </OPTION>
			 	<option value="," <% if rstemp("settingValue") = "," then response.write "SELECTED" %>> , </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pDateFormat" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select  Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="DD/MM/YY" <% if rstemp("settingValue") = "DD/MM/YY" then response.write "SELECTED" %>> DD/MM/YY </OPTION>
			 	<option value="MM/DD/YY" <% if rstemp("settingValue") = "MM/DD/YY" then response.write "SELECTED" %>> MM/DD/YY </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pDefaultLanguage" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pUseShippingAddress" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select  Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pRandomPassword" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pMinimumPurchase" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pAllowNewCustomer" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pStoreLocation" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pForceSelectOptionals" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pEncryptionPassword" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pWishList" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pIndexVisitsCounter" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pShowBtoBPrice" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pStoreNews" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pUseEncryptedTotal" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pForgotPassword" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pListBestSellers" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pRelatedProducts" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pGetRelatedProductsLimit" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	     
    	<%
    	elseif rstemp("settingKey")="pProductReviews" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	     
    	<%
    	elseif rstemp("settingKey")="pProductReviewsAutoActive" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
    	<%
    	elseif rstemp("settingKey")="pEmailToFriend" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pNewsLetter" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 

        <%
    	elseif rstemp("settingKey")="pGiftRandom" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pPercentageToDiscount" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pAuctions" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        
        <%
    	elseif rstemp("settingKey")="pRealTimeShipping" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="USPS" <% if ucase(rstemp("settingValue")) = "USPS" then response.write "SELECTED" %>> USPS </OPTION>
			 	<option value="UPS" <% if ucase(rstemp("settingValue")) = "UPS" then response.write "SELECTED" %>> UPS </OPTION> 
			 	<option value="CA" <% if ucase(rstemp("settingValue")) = "CA" then response.write "SELECTED" %>> CA </OPTION> 
			 	<option value="FEDEX" <% if ucase(rstemp("settingValue")) = "FEDEX" then response.write "SELECTED" %>> FEDEX </OPTION> 
			 	<option value="INTERSHIPPER" <% if ucase(rstemp("settingValue")) = "INTERSHIPPER" then response.write "SELECTED" %>> INTERSHIPPER </OPTION> 
			 	<option value="NONE" <% if ucase(rstemp("settingValue")) = "NONE" then response.write "SELECTED" %>> NONE </OPTION> 
			 </select> 			 
	
			 </td>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pUpsDefaultFee" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pUpsDefaultServiceType" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pUSPSDefaultServiceType" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pUSPSDefaultFee" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 	        
        <%
    	elseif rstemp("settingKey")="pPriceList" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pSendPlainText" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        
        <%
    	elseif rstemp("settingKey")="pCurrencyConversion" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="NONE" <% if ucase(rstemp("settingValue")) = "NONE" then response.write "SELECTED" %>> NONE </OPTION>
			 	<option value="STATIC" <% if ucase(rstemp("settingValue")) = "STATIC" then response.write "SELECTED" %>> STATIC </OPTION> 
			 	<option value="DYNAMIC" <% if ucase(rstemp("settingValue")) = "DYNAMIC" then response.write "SELECTED" %>> DYNAMIC </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pOneStepCheckout" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	     
        <%
    	elseif rstemp("settingKey")="pPayPalMerchantEmail" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	     
        <%
    	elseif rstemp("settingKey")="pRowsShown" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pPriceListShowStock" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pFraudPreventionMode" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="paranoid" <% if rstemp("settingValue") = "paranoid" then response.write "SELECTED" %>> paranoid </OPTION>
			 	<option value="caution" <% if rstemp("settingValue") = "caution" then response.write "SELECTED" %>> caution </OPTION>
			 	<option value="none" <% if rstemp("settingValue") = "none" then response.write "SELECTED" %>> none </OPTION>
			 </select> 			 

	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pDownloadGoodsType" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="NONE" <% if rstemp("settingValue") = "NONE" then response.write "SELECTED" %>> NONE </OPTION>
			 	<option value="ASPUPLOAD" <% if rstemp("settingValue") = "ASPUPLOAD" then response.write "SELECTED" %>> ASPUPLOAD </OPTION>
			 	<option value="ZIP" <% if rstemp("settingValue") = "ZIP" then response.write "SELECTED" %>> ZIP </OPTION>
			 	<option value="STRONGCUBE" <% if rstemp("settingValue") = "STRONGCUBE" then response.write "SELECTED" %>> STRONGCUBE </OPTION>
			 </select> 			 

	      	 </TD>      	 
	        </TR>    	 

        <%
    	elseif rstemp("settingKey")="pFraudPreventionSmallField" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 	        
        <%
    	elseif rstemp("settingKey")="pFraudPreventionPriceField" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 	        
        <%
    	elseif rstemp("settingKey")="pWhatsNewIdProductFrom" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			<TD width="330" height="21"> 		
			 <INPUT TYPE="Text" Name="<%=rstemp("idSetting")%>" VALUE="<%=rstemp("settingValue")%>" SIZE="<%=len(rstemp("settingValue"))+5%>">
			 <i><%=rstemp("comments")%></i>
	      	 </TD>      	 
	        </TR>    	 	        	        
        <%
    	elseif rstemp("settingKey")="pSerialCodeOnlyOnce" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        
        <%
    	elseif rstemp("settingKey")="pStockQuantityControl" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        	        
        <%
    	elseif rstemp("settingKey")="pDisableSaveOrderEmail" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 
        <%
    	elseif rstemp("settingKey")="pStoreClosed" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        
        <%
    	elseif rstemp("settingKey")="pDownloadDigitalGoods" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        
	        
	<%
    	elseif rstemp("settingKey")="pShippingTaxable" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        
	<%
    	elseif rstemp("settingKey")="pMoneyDontRound" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        	
	<%
    	elseif rstemp("settingKey")="pAllowDelayPayment" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        	                	        
	<%
    	elseif rstemp("settingKey")="pOfferBoxEnabled" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        	                	        	        
	<%
    	elseif rstemp("settingKey")="pCompareProducts" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        	                	        	        	        
	<%
    	elseif rstemp("settingKey")="pBoxCodeRegistration" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>    	 	        	                	        	        	        	        
	<%
    	elseif rstemp("settingKey")="pSendDigitalGoodsByEmail" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        
	<%
    	elseif rstemp("settingKey")="pLastChanceListing" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        
       <%
    	elseif rstemp("settingKey")="pByPassShipping" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        
	<%
    	elseif rstemp("settingKey")="pUseJavaCalendar" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        
<%
    	elseif rstemp("settingKey")="pListProductsByLetter" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        
<%
    	elseif rstemp("settingKey")="pAffiliatesStoreFront" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        
<%
    	elseif rstemp("settingKey")="pAffiliatesStoreFront" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        
	        
	<%
    	elseif rstemp("settingKey")="pCookiesDetection" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        
	<%
    	elseif rstemp("settingKey")="pSearchLog" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        	        
<%
    	elseif rstemp("settingKey")="pRecommendations" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        	        	        
<%
    	elseif rstemp("settingKey")="pCategoriesAlphOrder" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        	        	        	        
<%
    	elseif rstemp("settingKey")="pChangeDecimalPoint" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        	        	        	        
	        
	  <%
    	elseif rstemp("settingKey")="pUpsPickupType" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="01" <% if rstemp("settingValue") = "01" then response.write "SELECTED" %>> Daily Pickup</OPTION>
			 	<option value="03" <% if rstemp("settingValue") = "03" then response.write "SELECTED" %>> Customer Counter</OPTION> 
			 	<option value="06" <% if rstemp("settingValue") = "06" then response.write "SELECTED" %>> One Time Pickup</OPTION> 
			 	<option value="07" <% if rstemp("settingValue") = "07" then response.write "SELECTED" %>> On Call Air Pickup</OPTION> 
			 	<option value="11" <% if rstemp("settingValue") = "11" then response.write "SELECTED" %>> Suggested Retail Rates(UPS Store)</OPTION> 
			 	<option value="19" <% if rstemp("settingValue") = "19" then response.write "SELECTED" %>> Letter Center</OPTION> 
			 	<option value="20" <% if rstemp("settingValue") = "20" then response.write "SELECTED" %>> Air Service Center</OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        	        	        	        
	        
	  <%
    	elseif rstemp("settingKey")="pUpsPackagingType" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="01" <% if rstemp("settingValue") = "01" then response.write "SELECTED" %>> UPS Letter/Envelope</OPTION>
			 	<option value="02" <% if rstemp("settingValue") = "02" then response.write "SELECTED" %>> Package</OPTION> 
			 	<option value="03" <% if rstemp("settingValue") = "03" then response.write "SELECTED" %>> UPS Tube</OPTION> 
			 	<option value="04" <% if rstemp("settingValue") = "04" then response.write "SELECTED" %>> UPS Pak</OPTION> 
			 	<option value="21" <% if rstemp("settingValue") = "21" then response.write "SELECTED" %>> UPS Express Box</OPTION> 
			 	<option value="24" <% if rstemp("settingValue") = "24" then response.write "SELECTED" %>> UPS 25Kg Box</OPTION> 
			 	<option value="25" <% if rstemp("settingValue") = "25" then response.write "SELECTED" %>> UPS 10Kg Box</OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        	        	        	        
	  <%
    	elseif rstemp("settingKey")="pUpsExcludeServiceTypes" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
		        01 (Next Day Air)<input type="checkbox" name="<%=rstemp("idSetting")%>" value="01" <%If Instr(rstemp("settingValue"), "01") > 0 Then response.write "Checked" %>> <BR>
		        02 (2nd Day Air) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="02" <%If Instr(rstemp("settingValue"), "02") > 0 Then response.write "Checked" %>> <BR>
		        03 (Ground) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="03" <%If Instr(rstemp("settingValue"), "03") > 0 Then response.write "Checked" %>> <BR>
		        07 (Worldwide Express) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="07" <%If Instr(rstemp("settingValue"), "07") > 0 Then response.write "Checked" %>> <BR>
		        08 (Worldwide Expedited) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="08" <%If Instr(rstemp("settingValue"), "08") > 0 Then response.write "Checked" %>> <BR>
		        11 (Standard) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="11" <%If Instr(rstemp("settingValue"), "11") > 0 Then response.write "Checked" %>> <BR>
		        12 (3 Day Select) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="12" <%If Instr(rstemp("settingValue"), "12") > 0 Then response.write "Checked" %>> <BR>
		        13 (Next Day Air Saver) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="13" <%If Instr(rstemp("settingValue"), "13") > 0 Then response.write "Checked" %>> <BR>
		        14 (Next Day Air Early A.M.) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="14" <%If Instr(rstemp("settingValue"), "14") > 0 Then response.write "Checked" %>> <BR>
		        54 (Worldwide Express Plus) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="54" <%If Instr(rstemp("settingValue"), "54") > 0 Then response.write "Checked" %>> <BR>
		        59 (2nd Day Air A.M.) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="59" <%If Instr(rstemp("settingValue"), "59") > 0 Then response.write "Checked" %>> <BR>
		        65 (Express Saver) <input type="checkbox" name="<%=rstemp("idSetting")%>" value="65" <%If Instr(rstemp("settingValue"), "65") > 0 Then response.write "Checked" %>> <BR>
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        	        	        	        
	        
	  <%
    	elseif rstemp("settingKey")="pUpsResidentialAddress" then 
    	%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <select Name="<%=rstemp("idSetting")%>">			 	
			 	<option value="-1" <% if rstemp("settingValue") = "-1" then response.write "SELECTED" %>> YES </OPTION>
			 	<option value="0" <% if rstemp("settingValue") = "0" then response.write "SELECTED" %>> NO </OPTION> 
			 </select> 			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>   	        	        	        	        	        	        	        	        	        
	        
	  <%
	  ' default for new additions
	  else%>
	  	<TR> 
	         <TD width="187" height="21"> <%=rstemp("settingKey")%></TD>      			
			 <td>
			 <input type="text" name="<%=rstemp("idSetting")%>" value="<%=rstemp("settingValue")%>">			 
			 <i><%=rstemp("comments")%></i>
			 </td>
	      	 </TD>      	 
	        </TR>
	  <%  
	 end if 
	%>
    <%rstemp.movenext
       loop%>    
  </table>    
  <br><input type="Submit" name="Submit1" value="Save"><br>
</form>
<%call closeDb()%>
<!--#include file="footer.asp"-->
