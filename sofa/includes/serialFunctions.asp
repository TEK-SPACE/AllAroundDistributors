<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: serial functions
%>
<%
function getSerials(idProduct, quantity)
  
  dim pCompiledSerials, rstempSerial, rstempSerial2, pDescriptionS, pHasSerials
  
  pCompiledSerials=""
  
  ' get from settings
  
  pSerialCodeOnlyOnce	= getSettingKey("pSerialCodeOnlyOnce")
  
  ' get item description
  mySQL="SELECT description FROM products WHERE idProduct=" &idProduct
  call getFromDatabase(mySQL, rstempSerial, pGatewayName&" Silent Response, get serial codes")        
  
  if not rstempSerial.eof then  
   pDescriptionS=rstempSerial("description")
  else
   pDescriptionS="N/A"
  end if
  
  ' check if the item has any serial loaded 
   mySQL="SELECT idSerial FROM serials WHERE idProduct=" &idProduct  
   call getFromDatabase(mySQL, rstempSerial, pGatewayName&" Silent Response, get serial codes")            
   
   if not rstempSerial.eof then
    pHasSerials=-1
   else
    pHasSerials=0
   end if
    
  ' get serials, one for every item sold
 if pHasSerials=-1 then
  
  for f=1 to quantity
      
   ' get serial
   mySQL="SELECT idSerial, serialCode FROM serials WHERE used=0 AND idProduct=" &idProduct  
   call getFromDatabase(mySQL, rstempSerial, pGatewayName&" Silent Response, get serial codes")            
  
   if not rstempSerial.eof then
    
    ' compile the serial
    pCompiledSerials=pCompiledSerials&pDescriptionS&" -> "&rstempSerial("serialCode")&Vbcrlf   
        
    if pSerialCodeOnlyOnce="-1" then
    
      ' mark serial code as used
      mySQL="UPDATE serials SET used=-1 WHERE idSerial="&rstempSerial("idSerial")  
      call updateDatabase(mySQL, rstempSerial2, pGatewayName&" Silent Response, mark used serial code")                     
      
    end if 'only once
  
   else
    pCompiledSerials=pCompiledSerials&pDescriptionS&" -> " &getMsg(709,"No more available")&Vbcrlf   
   end if
      
  next
  
  end if
  
  getSerials=pCompiledSerials
      
end function
%>
