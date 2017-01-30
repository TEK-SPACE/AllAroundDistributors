<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->   
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp" --> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"-->   
<!--#include file="../includes/stringFunctions.asp"-->   
<!--#include file="../includes/encryption.asp"--> 
<!--#include file="../store/comersus_optDes.asp"--> 
<!--#include file="includes/settings.asp" --> 

<%
on error resume next

dim conntemp, rstemp, mysql, rstemp2, rstemp3, rstemp4

' get settings 
pEncryptionPassword 	= getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pCurrencySign    	= getSettingKey("pCurrencySign")
pDecimalSign    	= getSettingKey("pDecimalSign")
pMoneyDontRound    	= getSettingKey("pMoneyDontRound")


if pBackOfficeDemoMode=-1 then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Function disabled in demo mode")
end if

if request.form("Modify")="Modify" then

	pIdcustomer	= request.form("idCustomer")
	pName		= request.form("name")
	pLastName	= request.form("lastName")
	pCustomerCompany= request.form("customerCompany")
	pEmail		= request.form("email")	
	pPassword	= EnCrypt(request.form("password"), pEncryptionPassword)
	pCity		= request.form("city")
	pZip		= request.form("zip")
	pCountryCode	= request.form("countryCode")
	pState		= request.form("state")
	pStateCode	= request.form("stateCode")
	pPhone		= request.form("phone")
	pAddress	= request.form("address")
	pBonusPoints	= getUserInput(request.form("bonusPoints"),10)       
  	pBonusPoints    = formatNumberForDb(pBonusPoints)

        pIdCustomerType	= request.form("idCustomerType")
        pWapAccess	= request.form("wapAccess")
        pActive		= request.form("active")
        
        ' validation
        
        if pName="" or pLastName="" or pEmail="" or pAddress="" or pPhone="" or pPassword="" or pCity="" then
           response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Please enter all required fields")
        end if
        
        if pState<>"" and pStateCode<>"" then
           response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("You cannot select a state and a non listed state")
        end if

	pName=formatForDb(pName)
	pLastName=formatForDb(pLastName)
	pCustomerCompany=formatForDb(pCustomerCompany)
	
	' update customer record
	mysql="UPDATE customers SET name='" &pName& "', lastName='" &pLastName& "', customerCompany='" &pCustomerCompany& "', email='" &pEmail& "', password='" &ppassword&"',city='" &pcity& "',zip='" &pzip& "',countryCode='" &pCountryCode& "', state='" &pState& "', stateCode='" &pStateCode& "', phone='" &pPhone& "', address='" &pAddress& "', idCustomerType=" &pIdCustomerType& ", wapAccess=" &pWapAccess&", active=" &pActive&", bonusPoints="&pBonusPoints&" WHERE idCustomer="&pIdCustomer	
	
	call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyCustomerexec.asp") 

        call closeDb()
        response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Customer record modified")  	
  	
end if 

' modify
if request("Delete")="Delete" then  

 pIdCustomer	= request.form("idCustomer")
 
  ' delete orders first, for integrity reasons
  
   mysql="SELECT idOrder FROM orders WHERE idCustomer=" &pIdCustomer
   call getFromDatabase(mySQL, rstemp, "comersus_backoffice_modifyCustomerexec.asp") 
   
   do while not rstemp.eof
	   
	   ' get idDbSessionCart for that order
	   mysql="SELECT idDbSessionCart FROM dbSessionCart WHERE idOrder=" &rstemp("idOrder")
	   call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_modifyCustomerexec.asp") 		  	  
	  
	   ' get idDbSessionCart for that order
	   mysql="SELECT idDbSessionCart FROM dbSessionCart WHERE idOrder=" &rstemp("IdOrder")
	   call getFromDatabase(mySQL, rstemp2, "comersus_backoffice_modifyCustomerexec.asp") 	
	  
	   if not rstemp2.eof then
	     pIdDbSessionCart=rstemp2("idDbSessionCart")
	   else
	     pIdDbSessionCart=0
	   end if
	   
	   ' get idCartRows for that idDbSessionCart
	   mysql="SELECT idCartRow FROM cartRows WHERE idDbSessionCart=" &pIdDbSessionCart
	   call getFromDatabase(mySQL, rstemp3, "comersus_backoffice_modifyCustomerexec.asp") 
	
	   do while not rstemp3.eof
	   
	     ' delete all cartRowsOptions
	     mysql="DELETE FROM cartRowsOptions WHERE idCartRow="&rstemp3("idCartRow")
	     call updateDatabase(mySQL, rstemp4, "comersus_backoffice_modifyCustomerexec.asp") 
	    
	     rstemp3.movenext
	   loop
	  
	     ' delete cartRows	  
	     mysql="DELETE FROM cartRows WHERE idDbSessionCart="&pIdDbSessionCart
	     call updateDatabase(mySQL, rstemp4, "comersus_backoffice_modifyCustomerexec.asp") 	
	  
	     ' delete dbSessionCart
	     mysql="DELETE FROM dbSessionCart WHERE idDbSessionCart="&pIdDbSessionCart
	     call updateDatabase(mySQL, rstemp4, "comersus_backoffice_modifyCustomerexec.asp") 
		
	     ' delete credit cards
	     mysql="DELETE FROM creditCards WHERE idOrder=" &rstemp("IdOrder")
	     call updateDatabase(mySQL, rstemp2, "comersus_backoffice_modifyCustomerexec.asp") 
	  
	     ' delete recurring billing
	     mysql="DELETE FROM recurringBilling WHERE idOrder=" &rstemp("IdOrder")
	     call updateDatabase(mySQL, rstemp2, "comersus_backoffice_modifyCustomerexec.asp") 
	     
	     ' delete orders
	     mysql="DELETE FROM orders WHERE idOrder=" &rstemp("IdOrder")
	     call updateDatabase(mySQL, rstemp2, "comersus_backoffice_modifyCustomerexec.asp") 
	 
	   rstemp.movenext 
	 loop
	  
	 ' delete wish list	 
	 mysql="DELETE FROM wishlist WHERE idCustomer=" &pIdCustomer
	 call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyCustomerexec.asp") 
	
	 ' delete customer_specialPrices
	 mysql="DELETE FROM customer_specialPrices WHERE idCustomer=" &pIdCustomer
	 call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyCustomerexec.asp") 

	 ' delete customerTracking
	 mysql="DELETE FROM customerTracking WHERE idCustomer=" &pIdCustomer
	 call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyCustomerexec.asp") 
	 
	 ' delete current customer
	 mysql="DELETE FROM customers WHERE idCustomer=" &pIdCustomer
	 call updateDatabase(mySQL, rstemp, "comersus_backoffice_modifyCustomerexec.asp") 

 call closeDb()
 
 response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Customer record deleted")

end if 
%>