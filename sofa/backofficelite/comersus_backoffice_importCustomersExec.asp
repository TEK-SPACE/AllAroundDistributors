<%
' Comersus Diagnostics
' e-commerce ASP Open Source
' CopyRight Rodrigo S. Alhadeff, Comersus
' April 2005
' http://www.comersus.com 
' Import Customers 
%>
<!--#include file="comersus_backoffice_adminVerify.asp"--> 
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/itemFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->    
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/stringFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/encryption.asp"-->    
<%
on error resume next

if pBackOfficeDemoMode=-1 then
 response.redirect "comersus_backoffice_message.asp?message="&Server.Urlencode("Function disabled in demo mode.")
end if

Const ForReading = 1, ForWriting = 2
Dim mySQL, conntemp, rstemp, categoryDescription, pIdCategory
Dim fso, MyFile, pLine, valuesToInsert, fsoOut, MyFileOut

Set fso 	= CreateObject("Scripting.FileSystemObject")
Set MyFile 	= fso.OpenTextFile(Server.MapPath("customers.txt"), ForReading)       
set fsoOut  = Server.CreateObject("scripting.FileSystemObject") 

pIdStoreBackOffice	= getSessionVariable("idStoreBackOffice", 1)

%>
<!--#include file="header.asp"--> 

<b>Customers import</b>
<br><br>
<i>Importing file</i>: <%=Server.MapPath("customers.txt")%><br>
<%	

pEncryptionPassword = getSettingKey("pEncryptionPassword")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")

pFields		 = request("fields")

dim arrfields

pos_key = -1
arrfields=split(pFields,",")
for a = Lbound(arrfields) to ubound(arrfields)
	if trim(arrfields(a))= "email" then
		pos_key = a
	end if
next 
Quantity_fields = a 
pCounter=0   
if pos_key = -1 then
	response.write "Incorrect Format. the email field is required." 
else
	Do While MyFile.AtEndOfStream <> True
	 	
	  	pLine 		  = FormatForDb(trim(MyFile.ReadLine))
	 	
	 	if len(pLine)>5 then 
			response.write "<br><br>Trying to Import:<br>"  &pLine
	
		 	pPassword   = randomNumber(99999)
		 	'response.write pPassword
		 	pPassword	= EnCrypt(pPassword, pEncryptionPassword)
	
			call openDb()    
			arrdata=split(pLine,",")
			Quantity_info = ubound(arrdata)+1
			
			' verify duplicated.
			if pos_key >= Lbound(arrdata) and pos_key =< ubound(arrdata) and Quantity_info = Quantity_fields  then
				mysql = "select idCustomer from customers where email ='" & trim(arrdata(pos_key)) & "'" 
				'response.write mysql
				set rsTemp=connTemp.execute(mySQL)
				vflgDuplicated = 0
				if not rsTemp.eof then	 
					response.write "<br>Import error. Email duplicated:" & trim(arrdata(pos_key))
					vflgDuplicated = 1
				end if
			
				a = 0
				if vflgDuplicated = 0 then
					mySQL="INSERT INTO customers (idCustomerType, password, active, idStore, " &pFields&") VALUES (1,'" & pPassword & "',-1," & pIdStoreBackOffice	
					while a >= Lbound(arrdata) and a =< ubound(arrdata) 
						mySQL= mySQL & ",'"
						mySQL= mySQL & trim(arrdata(a)) & "'"  	
						a = a + 1 
					wend
					mySQL= mySQL & ")"
					
					set rsTemp=connTemp.execute(mySQL)
	
					if err.number <> 0 then 		 	  			
						Response.write " <br>-Error found!"
					else    	 	
				 		pCounter=pCounter+1
					end if	'error = 0
				end if 'customer duplicated
			else
				response.write "<br>Format File Error " 
			end if ' same number of items.
		else
			response.write "<br>Incorrect format." 
		end if	' len >5		
	Loop
end if 'email required
call closeDb()

%>

<br><br><%=pCounter%> records imported.<br>
<br><br>
<!--#include file="footer.asp"--> 