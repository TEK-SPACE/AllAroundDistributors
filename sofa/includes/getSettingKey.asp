<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: Get one setting key 
%>

<%
function getSettingKey(pSettingKey)

dim rstempGetSetting

mySQL="SELECT settingValue FROM settings WHERE settingKey='" &pSettingKey& "' AND idStore=" &pIdStore

call getFromDatabase(mySQL, rstempGetSetting, "getSettingKey")  

if not rstempGetSetting.eof then 
 getSettingKey=rstempGetSetting("settingValue")
else
 getSettingKey=""
end if

end function
%>

