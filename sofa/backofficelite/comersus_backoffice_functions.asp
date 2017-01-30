<%
' Comersus BackOffice Lite
' e-commerce ASP Open Source
' Comersus Open Technologies LC
' 2005
' http://www.comersus.com 
%>

<%
' replace user enters by breaks

Function showBreaks(str)
	showBreaks = replace(str, chr(10), "&nbsp;<br>")
End Function

' replace single quote by two single quotes

Function TwoSingleQ(str)
 TwoSingleQ	= replace(str,"'", "''" )
End Function

q = Chr(34)
%>