<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: customer tracking
%>
<%


function customerTracking(pAction, pMisc)

 dim rstempgSN, pIdCustomer, pDate, pHour
 
 pDateFormat			= getSettingKey("pDateFormat")
 pEnableCustomerTracking	= getSettingKey("pEnableCustomerTracking")
 
 if pEnableCustomerTracking="-1" then
 
  pIdCustomer     = getSessionVariable("idCustomer",0)
 
  if pIdCustomer<>0 then
   
   pDate	=fixDate(Date())
   pHour	=Time
   
   mySQL="INSERT INTO customerTracking (idCustomer, trackingDate, trackingHour, action, misc, idStore) VALUES (" &pIdCustomer& ",'" &pDate& "','" &pHour& "','" &pAction& "','" &pMisc& "'," &pIdStore& ")"
   call updateDatabase(mySQL, rstempgSN, "customerTracking")    
  
  end if
 
 end if ' enabled
 
end function


%>