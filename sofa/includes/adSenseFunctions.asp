<%
' Comersus Sophisticated Cart 
' Comersus Open Technologies
' USA - 2006
' http://www.comersus.com 
' Details: Google adSense functions
%>
<%
function displayAdPermanent()  

 ' get settings
 pAdSenseClient=getSettingKey("pAdSenseClient")
 pAdSenseType=getSettingKey("pAdSenseType")

if pAdSenseClient<>"0" and pAdSenseClient<>"" then
 
 randomize
 pRandomNumber=int(rnd*10)+1
 
 if pAdSenseType="permanent" then
 
  if pRandomNumber>2 then
 
	 %>
	 <script type="text/javascript"><!--
	 google_ad_client = "<%=pAdSenseClient%>";
	 google_ad_width = 468;
	 google_ad_height = 60;
	 google_ad_format = "468x60_as";
	 google_color_border = "BBBBBA";
	 google_color_bg = "FFFFFF";
	 google_color_link = "333333";
	 google_color_url = "008000"; 
	 google_color_text = "000000";
	 //--></script>
	 <script type="text/javascript"
	   src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
	 </script>
 	<%
 	
   else
 	 %>
	 <script type="text/javascript"><!--
	 google_ad_client = "pub-0428296380990344";
	 google_ad_width = 468;
	 google_ad_height = 60;
	 google_ad_format = "468x60_as";
	 google_color_border = "BBBBBA";
	 google_color_bg = "FFFFFF";
	 google_color_link = "333333";
	 google_color_url = "008000"; 
	 google_color_text = "000000";
	 //--></script>
	 <script type="text/javascript"
	   src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
	 </script>
 	<%
  end if ' replacement ID
 
 end if ' permanent type
 
end if ' adsense client loaded

end function

function displayAdConfirmation()  

 ' get settings
 pAdSenseClient=getSettingKey("pAdSenseClient")
 pAdSenseType=getSettingKey("pAdSenseType")

if pAdSenseClient<>"0" and pAdSenseClient<>"" then
 
 randomize
 pRandomNumber=int(rnd*10)+1
 
 if pAdSenseType="confirmation" then
 
  if pRandomNumber>2 then
 
	 %>
	 <script type="text/javascript"><!--
	 google_ad_client = "<%=pAdSenseClient%>";
	 google_ad_width = 468;
	 google_ad_height = 60;
	 google_ad_format = "468x60_as";
	 google_color_border = "BBBBBA";
	 google_color_bg = "FFFFFF";
	 google_color_link = "333333";
	 google_color_url = "008000"; 
	 google_color_text = "000000";
	 //--></script>
	 <script type="text/javascript"
	   src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
	 </script>
 	<%
 	
   else
 	 %>
	 <script type="text/javascript"><!--
	 google_ad_client = "pub-0428296380990344";
	 google_ad_width = 468;
	 google_ad_height = 60;
	 google_ad_format = "468x60_as";
	 google_color_border = "BBBBBA";
	 google_color_bg = "FFFFFF";
	 google_color_link = "333333";
	 google_color_url = "008000"; 
	 google_color_text = "000000";
	 //--></script>
	 <script type="text/javascript"
	   src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
	 </script>
 	<%
  end if ' replacement ID
 
 end if ' permanent type
 
end if ' adsense client loaded

end function

%>
