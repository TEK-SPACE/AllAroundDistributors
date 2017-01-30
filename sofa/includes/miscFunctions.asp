<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: miscellaneous functions
%>
<%

' discount code Generator (65^n)
function DiscountCodeGenerator(n) 
 dim s
 randomize()
 s="1234567890AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
 
 do
  a=""
  for i = 1 to n
   a = a + mid(s,cint(rnd()*len(s))+1,1)
  next
  
  ' search if the code is being used  
  mySQL = "SELECT discountCode FROM discounts WHERE discountCode='" &a& "'"	
  call getFromDatabase (mySql, rsTemp3, "miscFunctions")
 
 ' eof implies that the code is available		
 loop until rstemp3.eof 
 
 
 DiscountCodeGenerator = cstr(a)
end function

' randomNumber function, generates a number between 1 and limit
function randomNumber(limit)
 randomize
 randomNumber=int(rnd*limit)+1
end function

' min value
function min(byval value1, byval value2)
 if value1>value2 then
    min = value2
 else
    min = value1
 end if
end function

function getMax(arrProducts)
 dim tempMax, k
 tempMax=0
 for k=0 to uBound(arrProducts)
  if arrProducts(k,7)>tempMax then tempMax=arrProducts(k,7)
 next 
 getMax=tempMax
end function

function getPopularity(currentPopularity, maxPopularity)
 ' to avoid errors 
 if maxPopularity=0 then maxPopularity=1 
 getPopularity=int(currentPopularity*4/maxPopularity)
end function

function removePrefix(input,prefix)
 removePrefix=""
 if input<>"" and prefix<>"" then
     dim longPrefix
     longPrefix		=len(prefix)+1
     removePrefix	=mid(input,longPrefix)      
 end if
end function

function saveCookie()
 response.cookies("test")="-1"
end function

function readCookie() 
 if request.cookies("test")<>"-1" then 
  readCookie=0
 else
  readCookie=-1  
 end if
end function

function writeLog(line, file)
 set FSO 	= Server.CreateObject("scripting.FileSystemObject") 
 set myFile    = fso.OpenTextFile(file, 8, true) 
 myFile.WriteLine(line) 
 myFile.Close 
end function


sub parseShipmentMethod(pPlainString, pShipmentDesc, pShipmentPrice)
 
 pByPassShipping=getSettingKey("pByPassShipping")
  
 if  pByPassShipping="0" then
  arrayShipment=split(pPlainString,"%%")
 
  pShipmentDesc		=arrayShipment(0)
  pShipmentPrice	=Cdbl(arrayShipment(1))+Cdbl(arrayShipment(2))
 
  if pChangeDecimalPoint="-1" then 
   pShipmentPrice=replace(pShipmentPrice,".",",") 
  end if
  
 else
  
  pShipmentPrice	=0
  pShipmentDesc		=""
  
 end if
 
end sub

function getShipmentString(pShipmentIndex, pIdDbSession)
 
 dim rstemp3
 
 pByPassShipping=getSettingKey("pByPassShipping")
  
 getShipmentString=""
 
 if pByPassShipping="0" then
  
  mySQL = "SELECT sessionData FROM dbSession WHERE idDbSession=" &pIdDbSession
  call getFromDatabase (mySql, rsTemp3, "getShipmentString")
 
  if not rstemp3.eof then
   pSessionData=rstemp3("sessionData")
  
   ' parse rows
   arrayRows=split(pSessionData,"||")
  
   ' retrieve index
   getShipmentString=arrayRows(pShipmentIndex)
  
  end if
  
 end if 
 
end function


sub checkRentalAvailability(pIdProduct, pFrom, pUntil, pIsAvailable, pReason, pQuantity)
 
 pDateFormat	= getSettingKey("pDateFormat")
 pQuantity	=datediff("d",pFrom,pUntil) 

 if pQuantity<1 then
  pReason="Rental interval is not correct. Please check date format."
  pIsAvailable=0
 end if
 
 ' get availability

 mySQL="SELECT * FROM rentals WHERE idProduct="&pIdProduct
 call getFromDatabase(mySQL, rstemp, "comersus_rentalListAvailability.asp") 

 do while not rstemp.eof

  pRentalFromMinusSelectedFrom=datediff("d",pFrom,rstemp("fromDate"))
  pRentalUntilMinusSelectedFrom=datediff("d",pFrom,rstemp("untilDate"))

  pRentalFromMinusSelectedUntil=datediff("d",pUntil,rstemp("fromDate"))  
  pRentalUntilMinusSelectedUntil=datediff("d",pUntil,rstemp("untilDate"))
  
  pSelectedUntilMinusRentalFrom=datediff("d",rstemp("fromDate"),pUntil)
  pSelectedUntilMinusRentalUntil=datediff("d",rstemp("untilDate"),pUntil)
   
 
  if pRentalFromMinusSelectedFrom<=0 and pRentalUntilMinusSelectedFrom>=0 then 
   pReason=pReason&"selected start date reserved"
   pIsAvailable=0
  end if
 
  if pRentalFromMinusSelectedUntil<=0 and pRentalUntilMinusSelectedUntil>=0 then 
   if pReason<>"" then pReason=pReason&" and "
   pReason=pReason&"selected end date reserved"
   pIsAvailable=0
  end if
     
  if pRentalFromMinusSelectedFrom>0 and pSelectedUntilMinusRentalFrom>0 then
   if pReason<>"" then pReason=pReason&" and "
   pReason=pReason&"reservation start date inside selected period"
   pIsAvailable=0  
  end if
  
  if pRentalFromMinusSelectedFrom>0 and pSelectedUntilMinusRentalUntil>0 then
   if pReason<>"" then pReason=pReason&" and "
   pReason=pReason&"reservation end date inside selected period"
   pIsAvailable=0
  end if
	 
 rstemp.movenext
loop
  
  'response.end
  
end sub

function getStateName(pStateCode)
 
 dim rstempgSN
 
 mySQL="SELECT stateName FROM stateCodes WHERE stateCode='"&pStateCode&"'"
 call getFromDatabase(mySQL, rstempgSN, "getStateName()")    
 if not rstempgSN.eof then 
  getStateName=rstempgSN("stateName")
 else
  getStateName="-"
 end if
 
end function

function getCountryName(pCountryCode)

 dim rstempgSN
 mySQL="SELECT countryName FROM countryCodes WHERE countryCode='"&pCountryCode&"'"
 call getFromDatabase(mySQL, rstempgSN, "getCountryName()")    
 if not rstempgSN.eof then 
  getCountryName=rstempgSN("countryName")
 else
  getCountryName="-"
 end if

end function

function isDownloadFile(pFile)

 isDownloadFile=0
 
 if instr(pFile,".zip")<>0 then
   isDownloadFile=-1
 end if
 
 if instr(pFile,".exe")<>0 then
   isDownloadFile=-1
 end if
 
 if instr(pFile,".mp3")<>0 then
   isDownloadFile=-1
 end if
 
 if instr(pFile,".msi")<>0 then
   isDownloadFile=-1
 end if

 if instr(pFile,".rar")<>0 then
   isDownloadFile=-1
 end if
 
 if instr(pFile,".wma")<>0 then
   isDownloadFile=-1
 end if
 
end function

function isVideoFile(pFile)

 isVideoFile=0
 
 if instr(pFile,".avi")<>0 then
   isVideoFile=-1
 end if
 
 if instr(pFile,".wmv")<>0 then
   isVideoFile=-1
 end if
  
end function

function videoExtension(pFile)

 videoExtension=""
 
 if instr(pFile,".avi")<>0 then
   videoExtension=".avi"
 end if
 
 if instr(pFile,".wmv")<>0 then
   videoExtension=".wmv"
 end if
  
end function


function checkAvailability()
 
 dim rstempgSN
 
 checkAvailability=0
 
 mySQL="SELECT MAX(idOrder) AS maxIdOrder FROM orders"
 call getFromDatabase(mySQL, rstempgSN, "orderVerification")   
  
 if not rstempgSN.eof then 
  if rstempgSN("maxIdOrder")>50 then
   checkAvailability=-1
  end if
 end if
 
end function

function getBonusPoints(pIdCustomer)

 dim mysql, rstemp
 
 getBonusPoints=0
 
 ' retrieve available Bonus Points 
 if pIdCustomer<>0 then
 
 ' get current points
 mySQL="SELECT bonusPoints FROM customers WHERE idCustomer=" &pIdCustomer
 call getFromDatabase(mySQL, rstemp, "orderVerify")    
 
 getBonusPoints=rstemp("bonusPoints")   
 
 end if


end function

function affiliateValid(pIdAffiliate)

dim mysql, rstemp

 affiliateValid 	= Cint(1)

 if pIdAffiliate<>1 then

  ' check if idAffiliate is valid
  mySQL="SELECT idAffiliate FROM affiliates WHERE idAffiliate=" &pIdAffiliate

  call getFromDatabase(mySQL, rsTemp, "affiliateIsValid")    
 
  if rsTemp.eof then
   affiliateValid	= 0
  end If
 
 end if 

end function

function restBonusPoints(pBonusPointsToRest, pIdCustomer)

 dim mysql, rsTempB
 
  ' replace decimal comma
 pBonusPointsToRest	=Cstr(pBonusPointsToRest)
 pBonusPointsToRest	=replace(pBonusPointsToRest,",",".") 
 
 mySQL="UPDATE customers SET bonusPoints=bonusPoints-" &pBonusPointsToRest& " WHERE idCustomer="&pIdCustomer 
 call updateDatabase(mySQL, rsTempB, "restBonusPoints")       
 
end function
%>