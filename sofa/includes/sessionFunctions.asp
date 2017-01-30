<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: all session functions
%>
<!--#include file="../includes/dateFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<%

function sessionInit()
 
 ' makes an init over all session variables used
 if session("idCustomer")="" or isNull(session("idCustomer")) then 
  session("language")		= pDefaultLanguage      
  session("idAffiliate")	= Cint(1)     
  session("idCustomer")		= Cint(0)    
  session("idCustomerType")	= Cint(1) 
  session("cartItems")		= Cint(0)
  session("cartSubTotal")	= CDbl(0)
  session("idDbSession")	= Cint(0)
  session("idDbSessionCart")    = Cint(0)
  
  ' now check if the client browser supports session variables  
  if session("idCustomer")<>0 then
    response.redirect "comersus_message.asp?message="&Server.Urlencode(getMsg(662,"It seems that your browser does not support Cookies. You need Cookies enabled to purchase in this store."))        
  end if
 
  sessionInit=-1
 
 else
  sessionInit=0
 end if
end function



function getSessionVariable(pSessionVariable, pDefault)
 ' try to get one session variable, if is empty, returns default value
 if session(pSessionVariable)="" or isNull(session(pSessionVariable)) then
  getSessionVariable=pDefault  
 else
  getSessionVariable=session(pSessionVariable)
 end if
end function


function checkSessionData()

pIdDbSession		= getSessionVariable("idDbSession", 0)

if Cdbl(pIdDbSession)=0 then 
  
 ' dbSession not defined
 
 pRandomKey	= randomNumber(99999999)
 pDbSessionDate	= formatDate(Date())

 mySQL="INSERT INTO dbSession (randomKey, sessionType, dbSessionDate) VALUES (" &pRandomKey& ",'web','" &pDbSessionDate& "')"
 
 call updateDatabase(mySQL, rsTemp, "sessionFunctions, checkSessionData")
 
 ' retrieve idDbSession
 mySQL="SELECT MAX(idDbSession) AS maxIdDbSession FROM dbSession WHERE randomKey=" &pRandomKey& " AND dbSessionDate='" &pDbSessionDate& "'"
 
 call getFromDatabase(mySQL, rsTemp, "sessionFunctions, checkSessionData")
   
 pIdDbSession=rstemp("maxIdDbSession")

end if ' not NULL
 
 ' load in session
 session("idDbSession")=pIdDbSession
 
 checkSessionData = pIdDbSession
 
end function

function checkDbSessionCartOpen()
 
 dim pIdDbSession2  
 
 ' check if there is some cart open for current dbSession
 
 pIdDbSession2=getSessionVariable("idDbSession", 0)
 
 mySQL="SELECT idDbSessionCart FROM dbSessionCart WHERE idDbSession=" &pIdDbSession2& " AND cartOpen=-1 ORDER BY idDbSessionCart"
 
 call getFromDatabase(mySQL, rsTemp, "sessionFunctions")  

 if rstemp.eof then
  checkDbSessionCartOpen = 0
 else
  session("idDbSessionCart") = rstemp("idDbSessionCart")
  checkDbSessionCartOpen     = rstemp("idDbSessionCart")
 end if  

end function


function createNewDbSessionCart()  
 
 pIdDbSession=getSessionVariable("idDbSession", 0) 
 
 if pIdDbSession=0 then
    response.redirect "comersus_supportError.asp?error="&Server.UrlEncode("idDbSession was not defined while creating dbSessionCart. idDbSession:"&session("idDbSession")&" - Type:"&varType(session("idDbSession")))
 end if
 
 mySQL="INSERT INTO dbSessionCart (idDbSession, cartOpen) VALUES (" &pIdDbSession& ",-1)"
 
 call updateDatabase(mySQL, rsTemp, "sessionFunctions, createNewDbSessionCart")  
 
 ' retrieve idDbSession
  
 mySQL="SELECT MAX(idDbSessionCart) AS maxIdDbSessionCart FROM dbSessionCart WHERE idDbSession=" &pIdDbSession
 
 call getFromDatabase(mySQL, rsTemp, "sessionFunctions, createNewDbSessionCart")  
 
 pIdDbSessionCart=rstemp("maxIdDbSessionCart") 

 createNewDbSessionCart = pIdDbSessionCart  
 
end function

function sessionLost()
 pIdDbSession		= getSessionVariable("idDbSession", 0)
 
 if pIdDbSession=0 then
  sessionLost=-1
 else 
  sessionLost=0
 end if
end function

%>

