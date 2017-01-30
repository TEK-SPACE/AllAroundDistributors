<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2005
' http://www.comersus.com 
' Details: misc currency functions
%>
<%

function money(number)

dim pos, indexPoint, moneyLarge, decPart, g

if isNull(number) then 
 money="0.00"
else

 if pMoneyDontRound="-1" then
   money = number  
 else
   money	= round(number,2)
 end if 

 if pDecimalSign="," then
 ' replace . by ,
  money		= replace(money,".",",")
 else
 ' replace , by .
  money		= replace(money,",",".")
 end if

 if pMoneyDontRound="-1" then 
  Pos=inStr(money,pDecimalSign)
  if Pos<>0 then
  	money=mid(money,1,Pos+2)
  end if
 end if

 ' locate dec division
 indexPoint	= instr(money, pDecimalSign)

 ' for integer, add .00
 if indexPoint=0 then
    money	= Cstr(money)+pDecimalSign+"00"
 end if

 ' calculate if 0 or 00
 moneyLarge	= len(money)
 decPart		= right(money,moneyLarge-indexPoint)

 ' add to original numbers
 for g=0 to (1- (moneyLarge-indexPoint))
    money	= Cstr(money)+"0"    
 next

 ' money=separatorMil(money)

end if ' not empty
end function


' use this function to display numbers like 125,000,000
function SeparatorMil(number)

' locate pDecimalSign position
pDecimalSignExist = instr(number, pDecimalSign) - 1
if pDecimalSignExist <= 0 then 
	pDecimalSignExist = len(number)
end if

' add separatorMil to integers
pCountNumbers = 0
if pDecimalSignExist > 3 then
 for indexPoint = pDecimalSignExist to 1 step -1
	pNumber = mid(number,indexPoint,1)
	if pNumber <> pDecimalSign then
		pCountNumbers  = pCountNumbers + 1
		pSeparatorMil  = pNumber & pSeparatorMil 
 	end if	
	if pCountNumbers = 3 and (indexPoint > 1) then
        	pSeparatorMil  = "," & pSeparatorMil 
		pCountNumbers  = 0
	end if
	pNumber = ""
 next 

 ' add decimals
 pSeparatorMil = pSeparatorMil & mid(number, pDecimalSignExist+1, Len(number))
 SeparatorMil  = pSeparatorMil
else
 SeparatorMil  = number
end if

end function
%>