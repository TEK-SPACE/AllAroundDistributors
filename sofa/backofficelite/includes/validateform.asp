<%
' Comersus 5.0x Sophisticated Cart 
' Developed by Rodrigo S. Alhadeff for Comersus Open Technologies
' USA - 2003
' Open Source License can be found at documentation/readme.txt
' http://www.comersus.com  
' Details: validate form fields 
%>
<% 
const errorSymbol = "<font color=red><b>*</b></font>" 

dim dicError 
set dicError = server.createObject("scripting.dictionary") 


sub checkForm 

 dim fieldName, fieldValue, pFieldValue 

 for each field in request.form 

  if left(field, 1) = "_" then 
   ' is validation field , obtain field name
   fieldName  = right( field, len( field ) - 1) 
   ' obtain field value
   fieldValue = request.form(field) 
  
   select case lCase(fieldValue) 
  
   case "required" 
   if trim(request.form(fieldName)) = "" then     
    dicError(fieldName) = "You must enter a value for  " & "<i>" & fieldName & "</i>" 
   end if

   case "date" 
   if Not isDate(request.form(fieldName)) then 
     dicError(fieldName) = "You must enter a date for "& "<i>" & fieldName & "</i>" 
   end if  

   case "number"    
   pFieldValue=request.form(fieldName)   
   if Not isNumeric(pFieldValue) or (instr(pFieldValue,",")<>0) then 
     dicError(fieldName) = " You must enter a valid number for <i>" & fieldName & "</i>" 
   end if 
   
   case "intnumber"    
   pFieldValue=request.form(fieldName)   
   if Not isNumeric(pFieldValue) or (instr(pFieldValue,",")<>0) or (instr(pFieldValue,".")<>0) then 
     dicError(fieldName) = "You must enter an integer number for <i>" & fieldName & "</i>" 
   end if 
   
   case "positiveNumber" 
   if Not isNumeric(request.form(fieldName)) or (fieldName<0) then 
     dicError(fieldName) = "You must enter a positive number for <i>" & fieldName & "</i>" 
   end if 
   
   case "email" 
   if instr(request.form(fieldName),"@")=0 or instr(request.form(fieldName),".")=0 then 
     dicError(fieldName) = "You must enter a valid email for <i>" & fieldName & "</i>" 
   end if 
   
   case "phone" 
   pFieldValue=request.form(fieldName)
   pFieldValue=replace(pFieldValue," ","")
   pFieldValue=replace(pFieldValue,"-","")
   pFieldValue=replace(pFieldValue,"(","")
   pFieldValue=replace(pFieldValue,")","")
   if Not isNumeric(pFieldValue)  then 
     dicError(fieldName) = "You must enter a valid phone number for <i>" & fieldName & "</i>" 
   end if 
  
  end select 
 end if 
next 

end sub 

sub validateForm(byVal successPage) 
 
 if request.ServerVariables("CONTENT_LENGTH") > 0 then 
 checkForm 

 ' if no errors, then successPage 
  if dicError.Count = 0 then 
   
   ' build success querystring
  tString=Cstr("")
  
  for each field in request.form   
   if left(field, 1) <> "_" then 
    fieldName = field
    fieldValue = request.form(fieldName) 
    tString=tString &fieldName& "=" &Server.UrlEncode(fieldValue)& "&"
   end if   
  next
    
  response.redirect successPage&"?"& tString
  end if 

 end if 

end sub 

sub validateError  
 dim countRow
 countRow=cInt(0)
 
 for each field in dicError 
  
  if countRow=0 then
     response.write "<br>"
  end if
  
  response.write "<br> - " & dicError(field) 
  countRow=countRow+1
  
 next 
 
 if countRow>0 then
    response.write "<br><br>"
 end if
 
end sub 


sub validate( byVal fieldName, byVal validType ) 
%> <input name="_<%=fieldName%>" type="hidden" value="<%=validType%>"> <% 
 if dicError.Exists(fieldName) then 
  response.write errorSymbol 
 end if 
end sub 


sub textbox(byVal fieldName , byVal fieldValue, byVal fieldSize, byVal fieldType) 
 dim lastValue 
 lastValue = request.form(fieldName) 
 
   select case fieldType
  
   case "textbox" 
   %>
   <input name="<%=fieldName%>" size="<%=fieldSize%>" value="<%
   if trim(fieldValue)<>"" then
    response.write fieldValue
   else
    response.write Server.HTMLEncode(lastValue)    
   end if%>"> 
   <%
  
   case "password" 
   %><input name="<%=fieldName%>" type="password" size="<%=fieldSize%>" value="<%
   if trim(fieldValue)<>"" then
    response.write fieldValue
   else
    response.write server.HTMLEncode(lastValue)    
   end if%>"> 
   <%
   
   case "textarea" 
   %>
   <textarea name="<%=fieldName%>" rows="5" cols="<%=fieldSize%>"><%
   if trim(fieldValue)<>"" then
     response.write fieldValue    
   else
    response.write Server.HTMLEncode(trim(lastValue))
   end if
   %></textarea>
   <%
   
   end select 
   
end sub %> 

