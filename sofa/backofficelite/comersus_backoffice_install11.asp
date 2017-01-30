<!--#include file="header.asp"--> 
<!--#include file="./includes/settings.asp"-->
<!--#include file="../includes/settings.asp"-->  
<!--#include file="../includes/databaseFunctions.asp"-->    
<!--#include file="../includes/itemFunctions.asp"-->    
<!--#include file="../includes/dateFunctions.asp"-->    
<!--#include file="../includes/getSettingKey.asp"--> 
<!--#include file="../includes/encryption.asp"--> 

<%
on error resume next

dim mySQL, conntemp, rstemp

pRunInstallationWizard 	= getSettingKey("pRunInstallationWizard")
pEncryptionMethod 	= getSettingKey("pEncryptionMethod")
pEncryptionPassword	= getSettingKey("pEncryptionPassword")
pDateFormat		= getSettingKey("pDateFormat")

if pRunInstallationWizard<>"-1" then 
 response.redirect "comersus_backoffice_message.asp?message="&Server.UrlEncode("Sorry, the Installation Wizard is disabled")
end if

pPowerPacksInstalled=request.form("pPowerPacksInstalled")

' reset all settings
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pEmailToFriend'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pForgotPassword'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
   
 mySQL="UPDATE settings SET settingValue='none' WHERE settingKey='pCurrencyConversion'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pNewsLetter'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pProductReviews'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pRelatedProducts'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")   
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pAuctions'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pDownloadDigitalGoods'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pRecommendations'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pAffiliatesStoreFront'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pCompareProducts'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")    

 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pRssFeedServer'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pSuppliersList'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")
 
 mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pAllowDelayPayment'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
if pPowerPacksInstalled="PPP" then
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pEmailToFriend'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pForgotPassword'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='static' WHERE settingKey='pCurrencyConversion'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pAllowDelayPayment'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")      
  
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pNewsLetter'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pProductReviews'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pRelatedProducts'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")   

 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pAffiliatesStoreFront'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pCompareProducts'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")    
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pSuppliersList'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")    

 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pAuctions'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pDownloadDigitalGoods'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pRecommendations'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
 
 mySQL="UPDATE settings SET settingValue='-1' WHERE settingKey='pRssFeedServer'"
 call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp")  
    
  
end if

pPassword =Cstr(EnCrypt("123456", pEncryptionPassword))

' insert demo customer
mySQL="INSERT INTO customers (name, lastName, customerCompany, phone, email, password, address, zip, stateCode, city, countryCode, active, idCustomerType) VALUES ('John','Doe','Comersus','305 735 8008','test@comersus.com','"&pPassword&"','8345 NW 66 ST 3537','33166','FL','Miami','US',-1,1)"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 

' insert demo category
mySQL="INSERT INTO categories (categoryDesc, idParentCategory, active) VALUES ('Test Category',1,-1)"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 

' get id of demo supplier
mySQL="SELECT idSupplier FROM suppliers"
call getFromDatabase(mySQL, rstemp, "install11") 

pIdSupplier=0

pIdSupplier=rstemp("idSupplier")

' get ID of category
mySQL="SELECT idCategory FROM categories WHERE categoryDesc='Test Category'"
call getFromDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 

pIdCategory=0

if not rstemp.eof then
 pIdCategory=rstemp("idCategory")
end if

pDateAdded="10/10/2008"

' insert demo product
mySQL="INSERT INTO products (sku, description, details, price, listPrice, bToBPrice, cost, imageUrl, listhidden, weight, active, idSupplier, hotDeal, emailText, deliveringTime, formQuantity, smallImageUrl, showInHome, dateAdded, isBundleMain, idStore) VALUES ('T1','Succeeding in e-commerce','Rodrigo Alhadeff, founder and CEO of Comersus Open Technologies, goes the extra mile to explain useful tips and secrets that will help you create and maintain a successful online store. Maybe one of the great advantages of this book is that it is a wonderful source of quick solutions for e-commerce administrators, whether they manage a small business e-commerce or an extensive store. Covering everything from payment gateway selection to site design and fraud prevention, and even pro-sales displays and the concept of purchase drive, there is finally a practical and updated book available that will bring -and keep- your store afloat.',15,20,12,5,'Succ.jpg',0,1,-1,"&pIdSupplier&",-1,'-',2,5,'Succ_sm.jpg',-1,'" &pDateAdded& "',0," &pIdStore& ")"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 

' get product inserted
mySQL="SELECT idProduct FROM products WHERE sku='T1'"
call getFromDatabase(mySQL, rstemp, "install11") 

if rstemp.eof then
  response.redirect "comersus_backoffice_message.asp?message="& Server.Urlencode("Cannot locate inserted item")
end if

pIdProduct = rstemp("idProduct")

' insert categories for the item
mySQL="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory& ")"
call updateDatabase(mySQL, rstemp, "install11") 

' create stock record
call createStock(pIdProduct, 100)

' autoselect XMLHTTP component
pXmlHTTPComponent="NONE" 

'response.write "<br>Testing Msxml2.ServerXMLHTTP"
'Err.Clear
'Set xml1 = Server.CreateObject("Msxml2.ServerXMLHTTP")
'if(Err.number = 0) then
' pXmlHTTPComponent="Msxml2.ServerXMLHTTP"
'end if			

'response.write "<br>Testing Msxml2.ServerXMLHTTP.4.0"
'Err.Clear
'Set xml3 = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'if(Err.number = 0) then
' pXmlHTTPComponent="Msxml2.ServerXMLHTTP.4.0"
'end if

'response.write "<br>Testing Msxml2.ServerXMLHTTP.5.0"
'Err.Clear
'Set xml3 = Server.CreateObject("Msxml2.ServerXMLHTTP.5.0")
'if(Err.number = 0) then
' pXmlHTTPComponent="Msxml2.ServerXMLHTTP.5.0"
'end if

'response.write "<br>Testing Microsoft.XMLHTTP"
'Err.Clear
'set xml2 = Server.CreateObject("Microsoft.XMLHTTP")
'if(Err.number = 0) then
' pXmlHTTPComponent="Microsoft.XMLHTTP"
'end if

' update XML component
mySQL="UPDATE settings SET settingValue='"&pXmlHTTPComponent&"' WHERE settingKey='pXmlHTTPComponent'"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 

' disable Wizard
mySQL="UPDATE settings SET settingValue='0' WHERE settingKey='pRunInstallationWizard'"
call updateDatabase(mySQL, rstemp, "comersus_backoffice_install10.asp") 

' login admin
session("admin") = 1

%>
<br><font size="3"><b>Installation Wizard</b></font><br>
<br><br>You have completed the wizard!
<br><br>Things you can do now: 
<br>1. <a href="../default.asp"  target="_blank">Test the store</a>
<br>2. <a href="comersus_backoffice_importProductsExcelForm.asp"  target="_blank">Import your entire catalog from an Excel file</a>
<br>3. <a href="comersus_backoffice_addCategoryForm.asp" target="_blank">Add your own categories</a> and <a href="comersus_backoffice_addProductForm.asp"  target="_blank">products</a>
<br>4. <a href="http://www.comersus.org/">Request free technical assistance </a>
<br>5. <a href="http://www.comersus.com/power-pack.html">Purchase the Power Pack to increase functionality</a>

<!--#include file="footer.asp"-->