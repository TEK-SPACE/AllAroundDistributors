<%

' create array to store codes      
dim arrayCodes(15)
    
arrayCodes(1)="91D53"            
arrayCodes(2)="14A56"            
arrayCodes(3)="58298"            
arrayCodes(4)="7BB41"            
arrayCodes(5)="A5FEE"            
arrayCodes(6)="8AD3A"            
arrayCodes(7)="FCA25"            
arrayCodes(8)="37C92"            
arrayCodes(9)="72761"            
arrayCodes(10)="C3356"            
arrayCodes(11)="C97E7"            
arrayCodes(12)="EF2D5"            
arrayCodes(13)="8F9D8"            
arrayCodes(14)="E3E62"            
arrayCodes(15)="D2F76"            

' select random code
pSelection=randomNumber(15)            

' store in session
session("boxCode")=arrayCodes(pSelection)

%>
<%if pVerificationCodeEnabled="-1" then%>
 <img src="images/secret<%=pSelection%>.gif"> 
<%end if%>
