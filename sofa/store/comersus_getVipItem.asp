<% 
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: get VIP item
%>
       
<%

dim pIdProductSelected, rstempVip

pIdProductSelected	=0
 
pRandomSelection=randomNumber(5)

if pRandomSelection=1 then
 
 ' get best seller
 mySQL="SELECT idProduct FROM products WHERE idStore="&pIdStore&" AND active=-1 AND isBundleMain=0 AND listHidden=0 ORDER BY sales"
 call getFromDatabase (mySql, rstempVip, "getVipItem")
 if not rstempVip.eof then 
  pIdProductSelected=rstempVip("idProduct")
 end if
 
end if

if pRandomSelection=2 then
 ' item marked as VIP 
 pIdProductSelected	= getSettingKey("pIdProductVip")
 
 if pIdProductSelected="0" then
  ' VIP not defined, force expensive
  pRandomSelection=3
 end if
 
end if

if pRandomSelection=3 then
 
 ' get expensive item
 mySQL="SELECT idProduct FROM products WHERE idStore="& pIdStore&" AND active=-1 AND isBundleMain=0 AND listHidden=0 ORDER BY price"
 call getFromDatabase (mySql, rstempVip, "getVipItem")
 if not rstempVip.eof then 
  pIdProductSelected=rstempVip("idProduct")
 end if
 
end if

if pRandomSelection=4 then
 
 ' get most visited item
 mySQL="SELECT idProduct FROM products WHERE idStore="& pIdStore&" AND active=-1 AND isBundleMain=0 AND listHidden=0 ORDER BY visits"
 call getFromDatabase (mySql, rstempVip, "getVipItem")
 if not rstempVip.eof then 
  pIdProductSelected=rstempVip("idProduct")
 end if
 
end if

if pRandomSelection=5 then
 
 ' get recently added
 mySQL="SELECT idProduct FROM products WHERE idStore="& pIdStore&" AND active=-1 AND isBundleMain=0 AND listHidden=0 ORDER BY idProduct DESC"
 call getFromDatabase (mySql, rstempVip, "getVipItem")
 if not rstempVip.eof then 
  pIdProductSelected=rstempVip("idProduct")
 end if
 
end if
 
' get item details

 mySQL="SELECT description, details, map, visits FROM products WHERE active=-1 AND isBundleMain=0 AND idProduct="&pIdProductSelected& " AND idStore=" &pIdStore   
 call getFromDatabase (mySql, rstempVip, "getVipItem")

 if not rstempVip.eof then
   pDescriptionVip =rstempVip("description") 
   pDetailsVip	   =rstempVip("details")    
   pMapPriceVip	   =rstempVip("map")
   pVisits  	   =rstempVip("visits")
 end if
 

%> 
  
<font class=bar01 size="12px" color="#DA0008"><%=pDescriptionVip%></font>

<%
 pRateReview=getRateReview(pIdProductSelected)
	
if Cdbl(pRateReview)>0 then%>            
   <img src="images/<%=Cint(pRateReview)%>stars.gif">    
<%end if%>  

<br><div align=right>
<font class=bar01 size="14px" color="#DA0008">
<a href="comersus_viewItem.asp?idProduct=<%=pIdProductSelected%>">Special Price 

<%if pMapPriceVip = "-1" then%>
       	[<%=getMsg(745,"add to cart to find out")%>] 	
<%else%>
       	<%=pCurrencySign & money(getPrice(pIdProductSelected, pIdCustomerType, pIdCustomer))%>
<%end if%>  
</a>                    
</font>
</div>
&nbsp;



