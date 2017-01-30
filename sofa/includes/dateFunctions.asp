<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: date functions
%>
<%
function MediumDate (str)
  
  Dim aDay, aMonth, aYear
  aDay   	= Day(str)
  aMonth 	= Monthname(Month(str),True)
  aYear 	= Year(str)    
  MediumDate  = aDay & "-" & aMonth & "-" & aYear        
  
end Function

function formatDate(originalDate)
 
 if pDateSwitch="-1" then
  formatDate=mid(originalDate,4,2)&"/"&mid(originalDate,1,2)&"/"&right(originalDate,4)
 else
  formatDate=originalDate
 end if
 
end function

function fixDate(pNewFormat)        

if inStr(pNewFormat, "/")>0 then

    pToken1 = inStr(pNewFormat, "/")        
    
    pToken2 = inStr(mid(pNewFormat,pToken1+1, len(pNewFormat)) , "/" ) 

    pPart1 = mid(pNewFormat,1,pToken1-1) 
    pPart2 = mid(pNewFormat,pToken1+1,pToken2-1)
    pPart3 = mid(pNewFormat,pToken2+pToken1+1,len(pNewFormat))
    
    if len(pPart1) = 1 then
        pPart1  = "0" & pPart1
    end if
    if len(pPart2) = 1 then
        pPart2  = "0" & pPart2
    end if
    if len(pPart3) = 2 then
        pPart3  = "20" & pPart3
    end if
    pNewFormat =   pPart1 & "/" & pPart2 & "/" & pPart3  
   
  else
  
    pToken1 = inStr(pNewFormat, "-")        
    
    pToken2 = inStr(mid(pNewFormat,pToken1+1, len(pNewFormat)) , "-" ) 

    pPart1 = mid(pNewFormat,1,pToken1-1) 
    pPart2 = mid(pNewFormat,pToken1+1,pToken2-1)
    pPart3 = mid(pNewFormat,pToken2+pToken1+1,len(pNewFormat))
    
    if len(pPart1) = 1 then
        pPart1  = "0" & pPart1
    end if
    if len(pPart2) = 1 then
        pPart2  = "0" & pPart2
    end if
    if len(pPart3) = 2 then
        pPart3  = "20" & pPart3
    end if
    pNewFormat =   pPart1 & "/" & pPart2 & "/" & pPart3  
  
end if
   
    fixDate = pNewFormat
   
end function

function getServerDateFormat()
 
 if datediff("d","01/11/2005","01/12/2005")=1 then
  getServerDateFormat="MM/DD/YYYY"
 else
  getServerDateFormat="DD/MM/YYYY"
 end if
 
end function

Function return_RFC822_Date(myDate, offset)
   Dim myDay, myDays, myMonth, myYear
   Dim myHours, myMonths, mySeconds

   myDate = CDate(myDate)
   myDay = WeekdayName(Weekday(myDate),true)
   myDays = Day(myDate)
   myMonth = MonthName(Month(myDate), true)
   myYear = Year(myDate)
   myHours = zeroPad(Hour(myDate), 2)
   myMinutes = zeroPad(Minute(myDate), 2)
   mySeconds = zeroPad(Second(myDate), 2)

   return_RFC822_Date = myDay&", "& _
                                  myDays&" "& _
                                  myMonth&" "& _ 
                                  myYear&" "& _
                                  myHours&":"& _
                                  myMinutes&":"& _
                                  mySeconds&" "& _ 
                                  offset
End Function 

Function zeroPad(m, t)
   zeroPad = String(t-Len(m),"0")&m
End Function

%>
