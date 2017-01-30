<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<!--#include file="comersus_backoffice_adminVerify.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/settings.asp"-->  
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/currencyFormat.asp"-->  
<!--#include file="includes/clsBarChart.asp"-->  

<% 
on error resume next

dim mySQL, conntemp, rstemp, counter

' get settings 
pDefaultLanguage 	= getSettingKey("pDefaultLanguage")
pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pCompany	 	= getSettingKey("pCompany")
pCompanyLogo	 	= getSettingKey("pCompanyLogo")
pDateFormat		= getSettingKey("pDateFormat")

counter = 0

' count statistic registers


If Left(pDateFormat,2) = "DD" Then
	
	Select Case pDataBase
		Case "sqlserver":
			pMonth = "SUBSTRING(orderDate,4,2)"
		Case "mysql":
			pMonth = "SUBSTRING(orderDate,4,2)"
		Case Else:
			pMonth = "Mid(Left(orderDate,InStr(orderDate,Right(orderDate,5))-1),InStr(orderDate,'/')+1)"
	End Select
	
Else
	Select Case pDataBase
		Case "sqlserver":
			pMonth = "SUBSTRING(orderDate,1,2)"
		Case "mysql":
			pMonth = "SUBSTRING(orderDate,1,2)"
		Case Else:
			pMonth = "LEFT(orderDate, INSTR(orderDate, '/')-1)"
	End Select
	
End If

mySQL="SELECT COUNT(*) FROM orders GROUP BY " & pMonth

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_salesReport.asp")  	

quantity=Cint(0)

do  until rstemp.eof 
   quantity=quantity+1   
   rstemp.movenext
loop

' creates array for chart
ReDim arrValues(quantity-1)
ReDim arrLabels(quantity-1)

mySQL="SELECT SUM(total) AS totalSum, " & pMonth & " AS monthsql FROM orders GROUP BY " & pMonth

call getFromDatabase(mySQL, rstemp, "comersus_backoffice_salesreport.asp")  	

if  rstemp.eof then 
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("There is not enough information for this report")
end if

%>

<!--#include file="header.asp"-->
<!--#include file="./includes/settings.asp"--> 

<br><b>Sales report</b><br>
<%
set objChart = NEW BarChart 

do  until rstemp.eof 
   pTotalSum		= rstemp("totalSum")
   pMonth	 	= rstemp("monthsql")
   if isnull(pMonth) then
   		pMonth = 0
   else
   		if len(trim(pMonth)) = 0 then
   			pMonth = 0
   		else
   			pMonth = cInt(pMonth)
   		end if
   end if		
   if pMonth < 13 and pMonth > 0 then
		   arrValues(counter)	= Cdbl(pTotalSum)
		   arrLabels(counter)	= MonthName(pmonth, True)
		   counter=counter+1
		   %>
		  <br><%=MonthName(pmonth, True) %>, Total <%=pCurrencySign & money(pTotalSum)%><%
	end if
  rstemp.movenext
loop
%><br><br><%
With objChart
    .chartBGcolor = "#FFCC00"
    .chartTitle = "Monthly sales report"
    .chartWidth = "160"
    'the array holding the values to plot    
    .chartValueArray = arrValues
    'the array holding the labels for the values    
    .chartLabelsArray = arrLabels
    .chartColorArray = array("#FF9966" , "#009900" , "#000099")
    .chartViewDataType = "N"
    .chartBarHeight = 10
    .chartTextColor = "#990000"
END WITH

objChart.Draw 
%>
<br><br><i>Get more charts and reports with <a href="http://www.comersus.com/power-pack.html">Power Pack</a></i>
<%call closeDb()%>
<!--#include file="footer.asp"-->