<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: main initial settings (other settings can be found in database)
%>
<%

dim pDatabaseConnectionString, pSupportErrorEmailFrom, pSupportErrorSMTP, pSupportErrorEmailComponent, pSupportErrorShowDetails, pTrapDbErrors, pIdStore

' storeFront Version (don't change this constant)
private const pStoreFrontVersion	= "7.099"

' Options: Windows or unix/linux
private const pServerOS			= "windows"

' database (access, sqlserver, mysql)
private const pDataBase			= "access"

' input string filterin type (1-3 being 1 hard, 3 light)
pFilteringLevel				= 1

' id # if you have several stores connected to the same database
pIdStore			= 1

' Alternate connection strings (change only for SQl Server, mySQL or if you get errors with Access)

' SQL Server local or remote IP in SERVER=
' pDatabaseConnectionString = "Driver={SQL Server};UID=comersus;password=123456;DATABASE=comersus6;SERVER=127.0.0.1"

' mySQL  Server 2.5
' pDatabaseConnectionString = "Driver={mySQL};Server=localhost;database=comersus;Uid=Root;Pwd="

' mySQL  Server 3.51 local
' pDatabaseConnectionString = "Driver={MySQL ODBC 3.51 Driver};Server=localhost;database=comersus;user=root;password=123456;OPTION=3"                        
 
' DSN connection, you must define the DSN first in your server
' pDatabaseConnectionString = "DSN=comersus"
 
' DSN less connection
pDatabaseConnectionString = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" &server.MapPath("../database/comersus.mdb")&";" 



' Email Settings for Support Error script (if there's a DB error, this data cannot be retrieved from the DB)
pSupportErrorEmailFrom 		= "you@yourDomain.com"
pSupportErrorSMTP      		= "smtp.yourDomain.com"
pSupportErrorEmailComponent	= "Jmail" ' options are Jmail, ServerObjectsASPMail1, ServerObjectsASPMail2, PersitsASPMail, CDONTS and BambooSMTP
pSupportErrorShowDetails	= -1

' error debug level, trap common DB errors

pTrapDbErrors=0

%>


