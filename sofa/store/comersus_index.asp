<%
' Comersus Shopping Cart 
' Comersus Open Technologies
' United States 
' Software License can be found at License.txt
' http://www.comersus.com  
' Details: get products marked as showInHome
%>

<!--#include file="../includes/settings.asp"--> 
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/sessionFunctions.asp"--> 
<!--#include file="../includes/screenMessages.asp"-->
<!--#include file="../includes/databaseFunctions.asp"--> 
<!--#include file="../includes/currencyFormat.asp"--> 
<!--#include file="../includes/itemFunctions.asp"--> 
<!--#include file="../includes/cartFunctions.asp"--> 
<!--#include file="../includes/miscFunctions.asp"--> 
<!--#include file="../includes/dateFunctions.asp"--> 
<!--#include file="../includes/adSenseFunctions.asp"-->
<% 

on error resume next 

dim connTemp, rsTemp, rsTemp3, pItemsShown, pMaxPopularity, indexShowInHome, pIdCustomer, pIdCustomerType, howManyHome, mySQL, pDefaultLanguage, pStoreFrontDemoMode, pCurrencySign, pDecimalSign, pMoneyDontRound, pCompany, pCompanyLogo, pHeaderKeywords, pAuctions, pListBestSellers, pNewsLetter, pPriceList, pStoreNews, pOneStepCheckout, pAffiliatesStoreFront, pShowStockView, pIdProductVip, pLastChanceListing, pAllowNewCustomer, pListProductsByLetter, pRunInstallationWizard, pIdProduct, pDescription, pDetails, pDetailsVip, pImageUrlVip, pLanguage, pCustomerName, pHeaderCartItems, pHeaderCartSubtotal, pItemCounter, f, h

call saveCookie()

' get settings 

pStoreFrontDemoMode 	= getSettingKey("pStoreFrontDemoMode")
pCurrencySign	 	= getSettingKey("pCurrencySign")
pDecimalSign	 	= getSettingKey("pDecimalSign")
pMoneyDontRound	 	= getSettingKey("pMoneyDontRound")
pCompany	 	= getSettingKey("pCompany")
pCompanySlogan 		= getSettingKey("pCompanySlogan")
pAboutUsLink 		= getSettingKey("pAboutUsLink")
pStoreLocation		= getSettingKey("pStoreLocation")

pAdSenseClient 		= getSettingKey("pAdSenseClient")

pHeaderKeywords		= getSettingKey("pHeaderKeywords")
pDateSwitch		= getSettingKey("pDateSwitch")
pDateFormat		= getSettingKey("pDateFormat")

pSuppliersList		= getSettingKey("pSuppliersList")
pAuctions		= getSettingKey("pAuctions")
pAffiliatesStoreFront   = getSettingKey("pAffiliatesStoreFront")
pRssFeedServer		= getSettingKey("pRssFeedServer")
pNewsLetter 		= getSettingKey("pNewsLetter")

pShowStockView		= getSettingKey("pShowStockView")
pAllowNewCustomer	= getSettingKey("pAllowNewCustomer")
pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")
pItemsShown 		= Cint(getSettingKey("pItemsShown"))
pShowNews		= getSettingKey("pShowNews")
pShowAddedToday		= getSettingKey("pShowAddedToday")

pShowSearchBox		= getSettingKey("pShowSearchBox")

if pRunInstallationWizard="-1" then 
 response.redirect "../backofficeLite/comersus_backoffice_install0.asp"
end if

pPreviousProducts=""
pMaxPopularity 	= 0
indexShowInHome	= 0
pTodayCount	= 0
pIdCustomer     = getSessionVariable("idCustomer",0)
pIdCustomerType = getSessionVariable("idCustomerType",1)
pIdDbSessionCart= getSessionVariable("idDbSessionCart",0)

' get how many products are marked as showInHome

mySQL="SELECT COUNT(idProduct) AS howManyHome FROM products WHERE showInHome=-1 AND active=-1 AND listHidden=0 AND idStore=" &pIdStore
   
call getFromDatabase (mySql, rsTemp3, "dynamic_index")

if rsTemp3.eof then
 howManyHome=0
else
 howManyHome=rstemp3("howManyHome")
end if

if Cint(howManyHome)<Cint(pItemsShown) then 
 ' redirect to categories listing since there are no enough products to show in home
 response.redirect "comersus_listCategories.asp"
end if

' get products marked as showInHome

mySQL="SELECT idProduct, description, imageUrl, visits, map FROM products WHERE showInHome=-1 AND listHidden=0 AND active=-1 AND idStore=" &pIdStore

call getFromDatabase (mySql, rsTemp3, "dynamic_index")

' define the array to store products (2 columns)
reDim arrProducts(pItemsShown,9)

' load the array only with spots available 

do while (not rstemp3.eof) and indexShowInHome<(pItemsShown)  
 
 ' load product array 
 arrProducts(indexShowInHome,0)= rstemp3("idProduct")
 arrProducts(indexShowInHome,1)= rstemp3("description") 
 arrProducts(indexShowInHome,4)= rstemp3("imageUrl")
 arrProducts(indexShowInHome,5)= getStock(rstemp3("idProduct"))
 arrProducts(indexShowInHome,7)= rstemp3("visits")
 arrProducts(indexShowInHome,8)= rstemp3("map")
 
 ' store to avoid same random product
 pPreviousProducts=pPreviousProducts&" AND idProduct<>"&arrProducts(indexShowInHome,0)
 
 indexShowInHome = indexShowInHome+1
 
 rstemp3.moveNext
loop 

' calculate maximum visits
pMaxPopularity=getMax(arrProducts)

' get random product

reDim arrRandom(20,9)

mySQL="SELECT idProduct, description, smallImageUrl, visits, map FROM products WHERE idStore="& pIdStore& " AND listHidden=0 AND active=-1 and isBundleMain=0 "&pPreviousProducts
 
call getFromDatabase(mySql, rstemp, "index")

indexRandomProducts=0

do while not rstemp.eof and indexRandomProducts<20
    
 arrRandom(indexRandomProducts,0)= rstemp("idProduct")
 arrRandom(indexRandomProducts,1)= rstemp("description") 
 arrRandom(indexRandomProducts,4)= rstemp("smallImageUrl")
 arrRandom(indexRandomProducts,5)= getStock(rstemp("idProduct"))
 arrRandom(indexRandomProducts,7)= rstemp("visits")
 arrRandom(indexRandomProducts,8)= rstemp("map")
 
 indexRandomProducts=indexRandomProducts+1
 
rstemp.movenext
loop

' random spots to fill
pSpotsToFill=4

' only select random products to print if there are enough products

if indexRandomProducts>indexShowInHome+4 then

 reDim arrRandomSelections(indexRandomProducts)

 for f=0 to pSpotsToFill 
 
  pAccepted=0
 
  do while pAccepted=0
 
   arrRandomSelections(f)=randomNumber(indexRandomProducts-1)
  
   pAccepted=-1
  
   if f=0 then 
    pAccepted=-1
   else  
    for g=0 to f-1   
       if arrRandomSelections(g)=arrRandomSelections(f) then        
         pAccepted=0
       end if    
     next   
    end if 

  loop    
 next
end if ' enough products


%>
<!--#include file="header.asp"-->      
<table width="95%" border="0" cellspacing="0" cellpadding="0" bgcolor="#EAEAEA">     

<%

pColNumber=1

for f=0 to pItemsShown-1
%> 

   <%if pColNumber=1 then%>
    <tr><td width="50%" height="50">
   <%end if%>     


<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="center" class="description"><%=arrProducts(f,1)%></div></td></tr>
        <tr>
          <td><div class=price align="center">
              
	      <%if arrProducts(f,8) = "-1" Then%>
	     	<a href="comersus_addItem.asp?idProduct=<%=arrProducts(f,0)%>&quantity=1">?</a>	
	     <%else%>
	     	<%=pCurrencySign & money(getPrice(arrProducts(f,0), pIdCustomerType, pIdCustomer))%>     
	     <%end if%></div>
      </td></tr>
       
       <tr> 
          <td><div align="center"><%if pShowStockView="-1" then%>
      	   <%=getMsg(3,"Stock")%>: <%=arrProducts(f,5)%>
      	  <%end if%> &nbsp;
          <%=getMsg(4,"Pop")%>: &nbsp;<%=arrProducts(f,7)%></div></td>
        </tr>
        
        <tr> 
          <td><div align="center">
          <%if arrProducts(f,4)<>"" then%>
       		<a href="comersus_viewItem.asp?idProduct=<%=arrProducts(f,0)%>"><img alt="<%=arrProducts(f,1)%>" border=0 src="catalog/<%=arrProducts(f,4)%>" vspace=3></a> 
     	  <%else%>
       		<img src='catalog/imageNA.gif' alt="<%=getMsg(5,"Not availab")%>">              
          <%end if%>
          </div></td>
        </tr>
        <tr> 
          <td><div align="center">
          <img src="Images/buttonDetailsAdd.gif" width="176" height="31" usemap="#<%=f%>" border="0"><map name="<%=f%>"><area shape="rect" coords="6,10,78,28" href="comersus_viewItem.asp?idProduct=<%=arrProducts(f,0)%>" alt="View item details" title="View item details"><area shape="rect" coords="98,10,168,30" href="comersus_addItem.asp?idProduct=<%=arrProducts(f,0)%>&quantity=1" alt="Add to cart" title="Add to cart"></map>
          </div></td>
        </tr>
</table>           

  <%if pColNumber=1 then%>
   </td><td width="50%" height="50">
   <%pColNumber=2%>
  <%else%> 
   </td></tr>
   <%pColNumber=1%>
  <%end if%>     
        
<%next%>   


<%if indexRandomProducts>indexShowInHome+4 then%>

<tr><td colspan=2 align=center>&nbsp;</td></tr>
<tr><td colspan=2 align=center><div class="description"><%=getMsg(637,"Featured")%></div></td></tr>

<tr>
<td colspan=2>

<table width=95%>
<tr align=center>

<td>
<%if arrRandom(arrRandomSelections(0),4)<>"" then%>
 <a href="comersus_viewItem.asp?idProduct=<%=arrRandom(arrRandomSelections(0),0)%>"><img alt="<%=arrRandom(arrRandomSelections(0),1)%>" border=0 src="catalog/<%=arrRandom(arrRandomSelections(0),4)%>" vspace=3></a> 
<%else%>
 <img src='catalog/imageNA.gif' alt="<%=getMsg(5,"Not availab")%>">              
<%end if%>   
<br><a href="comersus_viewItem.asp?idProduct=<%=arrRandom(arrRandomSelections(0),0)%>"><%=arrRandom(arrRandomSelections(0),1)%></a>      
</td>

<td>
<%if arrRandom(arrRandomSelections(1),4)<>"" then%>
 <a href="comersus_viewItem.asp?idProduct=<%=arrRandom(arrRandomSelections(1),0)%>"><img alt="<%=arrRandom(arrRandomSelections(1),1)%>" border=0 src="catalog/<%=arrRandom(arrRandomSelections(1),4)%>" vspace=3></a> 
<%else%>
 <img src='catalog/imageNA.gif' alt="<%=getMsg(5,"Not availab")%>">              
<%end if%>  
<br> <a href="comersus_viewItem.asp?idProduct=<%=arrRandom(arrRandomSelections(1),0)%>"><%=arrRandom(arrRandomSelections(1),1)%></a>       
</td>

<td>
<%if arrRandom(arrRandomSelections(2),4)<>"" then%>
 <a href="comersus_viewItem.asp?idProduct=<%=arrRandom(arrRandomSelections(2),0)%>"><img alt="<%=arrRandom(arrRandomSelections(2),1)%>" border=0 src="catalog/<%=arrRandom(arrRandomSelections(2),4)%>" vspace=3></a> 
<%else%>
 <img src='catalog/imageNA.gif' alt="<%=getMsg(5,"Not availab")%>">              
<%end if%>   
<br><a href="comersus_viewItem.asp?idProduct=<%=arrRandom(arrRandomSelections(2),0)%>"><%=arrRandom(arrRandomSelections(2),1)%></a>      
</td>

<td>
<%if arrRandom(arrRandomSelections(3),4)<>"" then%>
 <a href="comersus_viewItem.asp?idProduct=<%=arrRandom(arrRandomSelections(3),0)%>"><img alt="<%=arrRandom(arrRandomSelections(3),1)%>" border=0 src="catalog/<%=arrRandom(arrRandomSelections(3),4)%>" vspace=3></a> 
<%else%>
 <img src='catalog/imageNA.gif' alt="<%=getMsg(5,"Not availab")%>">              
<%end if%>     
<br> <a href="comersus_viewItem.asp?idProduct=<%=arrRandom(arrRandomSelections(3),0)%>"><%=arrRandom(arrRandomSelections(3),1)%></a>    
</td>

</tr>
</table>

</td>   
</tr>       

<%end if%>
<%if pShowAddedToday="-1" then
 
 mySQL="SELECT idProduct, description FROM products WHERE dateAdded='" &Date()& "' AND idStore="& pIdStore& " AND listHidden=0 AND active=-1 and isBundleMain=0 ORDER BY idProduct DESC"
 
 call getFromDatabase(mySql, rstemp,"index")
 
  if not rsTemp.eof then 
    pIdProduct1		=rstemp("idProduct")
    pDescription1	=rstemp("description")
    pTodayCount	=pTodayCount+1
 end if

 if not rsTemp.eof then rstemp.movenext

 if not rsTemp.eof then 
    pIdProduct2		=rstemp("idProduct")
    pDescription2	=rstemp("description")
    pTodayCount	=pTodayCount+1
 end if
 
end if%>

<%if pShowAddedToday="-1" and pTodayCount>1 then%>

<tr><td colspan=2>&nbsp;</td></tr>
<tr>

<td colspan=2>

<b><%=getMsg(724,"added today")%></b>

<br>• <a href="comersus_viewItem.asp?idProduct=<%=pIdProduct1%>"><%=pDescription1%></a> &nbsp;<%=pCurrencySign & money(getPrice(pIdProduct1, pIdCustomerType, pIdCustomer))%>  &nbsp; • <a href="comersus_viewItem.asp?idProduct=<%=pIdProduct2%>"><%=pDescription2%></a> &nbsp;<%=pCurrencySign & money(getPrice(pIdProduct2, pIdCustomerType, pIdCustomer))%>

</td></tr>
<%end if%>

<!--#include file="comersus_getNews.asp"--> 

</table>
      
<!--#include file="footer.asp"--> 
<%call closeDb()%>

