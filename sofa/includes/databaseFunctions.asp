<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: functions to open and close db connection 
%>
<!--#include file="../includes/adovbs.inc"--> 
<%

sub openDb()
 if varType(connTemp)=0 or varType(connTemp)=1 then   
 
   ' create the connection
   set connTemp	= server.createObject("adodb.connection")      
    
   connTemp.Open pDatabaseConnectionString
 
   if err.number <> 0 then
     response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Error while opening DB read:"&Err.Description& "<br><br><b>Common solutions</b><br><br>1. Check that you haven't change default database path and name <br>2. Check that your web server has Access 97 or 2000 ODBC installed <br>3. Check that you have read, modify and delete permissions over database folder and database file  <br>4. Open your database with Access program and select Repair Database option <br>5. Select other connection method like other connection string or DSN")    
   end if
  
 end if 
end sub

sub getFromDatabase(mySQL, rsTemp, scriptName) 
  
   call openDb()
 
   set rsTemp = server.createObject("adodb.recordset")

   ' set locktype
   rsTemp.lockType = adLockReadOnly

   ' set the cursor
   rsTemp.cursorType = adOpenForwardOnly 
   
   rsTemp.open mySQL, connTemp    
   
   if err.number <> 0 then
      response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Error in " &scriptName& ", error: "&Err.Description& " - Err.Number:"&Err.number&" - SQL:"&mySQL)            
   end if
 
end sub

sub getFromDatabasePerPage(mySQL, rsTemp, scriptName)

 call openDb()
 
 set rsTemp 		= server.createObject("adodb.recordset")     
 
 rsTemp.cursorLocation 	= adUseClient
 rsTemp.cacheSize 	= pNumPerPage
 
 rsTemp.open mySQL, connTemp
 
 if err.number <> 0 then  
    response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Error in " &scriptName& ", error: "&Err.Description& " - Err.Number:"&Err.number&" - SQL:"&mySQL)        
 end if
 
end sub

sub getFromDatabaseSeek(mySQL, rsTemp, scriptName)

 call openDb()
 
 set rsTemp 		= server.createObject("adodb.recordset")
 rsTemp.cursorType 	= 3
 rsTemp.lockType 	= 3

 rsTemp.Open mySQL, connTemp
 
 if err.number <> 0 then  
    response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Error in " &scriptName& ", error: "&Err.Description& " - Err.Number:"&Err.number&" - SQL:"&mySQL)          
 end if
 
end sub

sub updateDatabase(mySQL, rsTemp, scriptName) 
  
 call openDb()    
   
 set rsTemp=connTemp.execute(mySQL)
 
 if err.number <> 0 then  
   response.redirect "comersus_supportError.asp?error="&Server.Urlencode("Update Error in " &scriptName& ", error: "&Err.Description& " - Err.Number:"&Err.number&" - SQL:"&mySQL)          
 end if
  
end sub


function closeDB()
  on error resume next
  
  rsTemp.close
  set rsTemp 		= nothing 
  connTemp.close
  set connTemp	 	= nothing
 
end function

%>