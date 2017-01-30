<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: string functions
%>
<%

function getLoginField(input,stringLength)
 
 ' to filter login fields
 
 dim regEx
 Set regEx = New RegExp
 
 getLoginField		= left(trim(input),stringLength) 
 
 regEx.Pattern 		= "([^-_A-Za-z0-9@.])"
 regEx.IgnoreCase 	= True
 regEx.Global 		= True
 getLoginField 		= regEx.Replace(getLoginField, "")
 Set regEx 		= nothing
     
end function

function getScreenMessage(input,stringLength)
 
 ' to filter screenMessage
 
 dim regEx
 Set regEx = New RegExp
 
 getScreenMessage	= left(trim(input),stringLength) 
 
 regEx.Pattern 		= "([^-_A-Za-z0-9@., ])"
 regEx.IgnoreCase 	= True
 regEx.Global 		= True
 getScreenMessage 	= regEx.Replace(getScreenMessage, "")
 Set regEx 		= nothing
     
end function

function getUserInput(input, stringLength)

 dim newString, regEx
 Set regEx = New RegExp
 
 ' only specified length
 newString = left(trim(input),stringLength)   
 
 if pFilteringLevel=1 then  
 
  regEx.Pattern 	= "([^A-Za-z0-9@=:/*|' _-]+.%)"
  regEx.IgnoreCase 	= True
  regEx.Global 		= True
  newString 		= regEx.Replace(newString, "")
  Set regEx 		= nothing   
  
  newString		= replace(newString,"--","")
  newString		= replace(newString,";","")      
  newString	        = replace(newString,"'","&#39;") 
  newString	        = replace(newString,"<script>","[script]") 
  
 end if
 
 if pFilteringLevel=2 then
    
  newString	= replace(newString,"--","")
  newString	= replace(newString,";","&#59;") 
  newString	= replace(newString,"=","&#61;") 
  newString	= replace(newString,"(","&#40;") 
  newString	= replace(newString,")","&#41;")   
  newString	= replace(newString,"'","&#39;") 
  newString	= replace(newString,"""","&#34;") 
  newString	= replace(newString,"<script>","[script]") 
 
 end if
 
 if pFilteringLevel=3 then
    
  newString	= replace(newString,"'","&#39;") 
  newString	= replace(newString,"""","&#34;") 
  newString	= replace(newString,"<script>","[script]") 
 
 end if 
 
 getUserInput		= newString 
 
end function

function getUserInputL(input,stringLength)

 ' light filtering
 dim tempStr
 
 tempStr	= left(input,stringLength)    
 tempStr	= replace(tempStr,"--","")
 tempStr	= replace(tempStr,";","&#59;") 
 tempStr	= replace(tempStr,"=","&#61;") 
 tempStr	= replace(tempStr,"(","&#40;") 
 tempStr	= replace(tempStr,")","&#41;") 
 tempStr	= replace(tempStr,"CHAR","&#67;&#72;&#65;&#82;") 
 tempStr	= replace(tempStr,"'","&#39;") 
 tempStr	= replace(tempStr,"""","&#34;") 
 tempStr	= replace(tempStr,"<","") 
 tempStr	= replace(tempStr,">","") 
 
 
 getUserInputL	= tempStr 
end function

function formatForDb(input)
 dim tempStr  
 tempStr=input
 if isNull(tempStr)=false then
 ' replace to avoid DB errors 
  tempStr	= replace(tempStr,"'","''") 
  tempStr	= replace(tempStr,"''''","''") 
  tempStr	= replace(tempStr,"''''''","''") 
  tempStr	= replace(tempStr,"''''''''","''") 
  tempStr	= replace(tempStr,"""","''")  
 end if
 formatForDb	= tempStr 
 
end function 

function formatNumberForDb(input)
 formatNumberForDb=replace(input,",",".")
end function

function getEditorInput(input, stringLength)

 dim newString, regEx
 Set regEx = New RegExp
 
 ' only specified length
 newString = left(trim(input),stringLength)   
 
 
  regEx.Pattern 	= "([^A-Za-z0-9@=:/*|' _-]+.%)"
  regEx.IgnoreCase 	= True
  regEx.Global 		= True
  newString 		= regEx.Replace(newString, "")
  Set regEx 		= nothing   
  
  newString		= replace(newString,"--","")
  newString	        = replace(newString,"'","&#39;") 
  newString	        = replace(newString,"<script>","[script]") 
  
  getEditorInput = newString
end function

function getQueryString(input,stringLength)
  
  dim regEx, newString    
  
  ' only specified length
  newString = left(trim(input),stringLength)  

  set regEx = New RegExp
  regEx.IgnoreCase = True
  regEx.Global  = True
  regEx.Pattern = "([^-A-Z a-z0-9])"
  newString 	= regEx.Replace(newString, "")
  
  getQueryString=newstring
end function

%>