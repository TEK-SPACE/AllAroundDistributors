<HTML>
<HEAD>

 	<%if pTitle<>"" then%>
	  <title><%=pCompany%> :: <%=pTitle%></title>
	<%else%>
	  <title><%=pCompany%> :: OnLine Store :: <%=chr(67)&chr(111)&chr(109)&chr(101)&chr(114)&chr(115)&chr(117)&chr(115)%></title>
	<%end if%>
	
	<%if pMetaDescription="" then%>
	 <meta name="DESCRIPTION" content="<%=pCompany%> Shopping Cart :: Purchase OnLine">
	<%else%>
	 <meta name="DESCRIPTION" content="<%=pMetaDescription%>">
	<%end if%>
	
	<%if pSearchKeywords="" or isNull(pSearchKeywords) then%>
	 <meta name="KEYWORDS" content="<%=pHeaderKeywords%>">
	<%else%>
	 <meta name="KEYWORDS" content="<%=pSearchKeywords%>">
	<%end if%>
	
	<link rel="stylesheet" type="text/css" href="images/style.css">	
	<meta name="robots" CONTENT="index, follow"> 
	<meta name="distribution" content="global">
	
</HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 background="Images/pixelbg.gif">
<TABLE WIDTH=960 BORDER=0 CELLPADDING=0 CELLSPACING=0 align="center">
  <TR>
		<TD>

<TABLE WIDTH=960 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD background="Images/upbarbancl_01.gif" WIDTH=320 HEIGHT=98>
			<br><br><h1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=pCompany%></h1>
		</TD>
		<TD>			
            <table WIDTH=503 HEIGHT=98 border="0" cellpadding="0" cellspacing="0" background="Images/pixelup.gif" >
              <tr> 
                <td>
                  <div align="center">

		  <%call displayAdPermanent()%>
                  
                  </div>
                </td>
              </tr></table>
</TD>
		<TD>
			<IMG SRC="Images/upbarban_03.gif" WIDTH=137 HEIGHT=98 ALT=""></TD>
	</TR>
</TABLE>		
		</TD>
	</TR>
	<TR>
		<TD>

<TABLE WIDTH=960 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		  <TD> <a href="comersus_listItems.asp"><IMG SRC="Images/midbar_01.gif" WIDTH=217 HEIGHT=115 ALT="Products" border="0"></a></TD>
		  <TD> <a href="comersus_customerUtilitiesMenu.asp"><IMG SRC="Images/midbar_02.gif" WIDTH=188 HEIGHT=115 ALT="My Account" border="0"></a></TD>
		  <TD> <a href="comersus_showCart.asp"><IMG SRC="Images/midbar_03.gif" WIDTH=209 HEIGHT=115 border="0" ALT="Shopping Cart"></a></TD>
		  <TD> <a href="comersus_contactUs.asp"><IMG SRC="Images/midbar_04.gif" WIDTH=203 HEIGHT=115 ALT="Contact Us" border="0"></a></TD>
		  <TD> <a href="comersus_listItems.asp?hotDeal=-1"><IMG SRC="Images/midbar_05.gif" WIDTH=143 HEIGHT=115 ALT="Sale Listing" border="0"></a></TD>
	</TR>
</TABLE>		
		</TD>
	</TR>
	<TR>
		<TD>

<TABLE WIDTH=960 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD>
			<IMG SRC="Images/content_01.gif" WIDTH=35 HEIGHT=643 ALT=""></TD>
		<TD>
		
		<TABLE WIDTH=128 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD>
			<IMG SRC="Images/leftbar_01.gif" WIDTH=128 HEIGHT=28 ALT=""></TD>
	</TR>
	<TR>
		<TD>
            <table WIDTH=128 HEIGHT=103 ><tr><td>

Date: <%=formatDate(fixDate(Date))%>
<br>News: <a href="comersus_newsPageExec.asp">Click here</a>
<br>Cart: <a href="comersus_showCart.asp"><%=session("cartItems")%> (item/s) </a>  
<br>Subtotal: <%=pCurrencySign & money(getSessionVariable("cartSubTotal",0))%>
            </td></tr></table>
</TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="Images/leftbar_03.gif" WIDTH=128 HEIGHT=53 ALT=""></TD>
	</TR>
	<TR>
		<TD>
 <table WIDTH=128 HEIGHT=112><tr><td>
 
     <%if pIdCustomer=0 then%>
      <form name="login" method="post" action="comersus_customerAuthenticateExec.asp">
      
      <%if pStoreFrontDemoMode="-1" then%>
        <input type="text" name="email" size="15" value="test@comersus.com">
        <input type="password" name="password" size="15" value="123456">
      <%else%>
        <input type="text" name="email" size="15" value="Email">
        <input type="password" name="password" size="15" value="">
       <%end if%> 
        <br><br>
        <input type="submit" name="login" value="login">
      </form>
    <%else%>
    
     Welcome back <br><%=session("customerName")%>     
     <div class="b01">
     <br>&nbsp;<img src="images/arrow.gif"> &nbsp;<a href="comersus_customerUtilitiesMenu.asp">Your Account</a> 
     <br>&nbsp;<img src="images/arrow.gif"> &nbsp;<a href="comersus_customerLogout.asp">Logout</a></div>
     <%end if%>

 
 </td></tr></table>
</TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="Images/leftbar_05.gif" WIDTH=128 HEIGHT=53 ALT=""></TD>
	</TR>
	<TR>
		<TD>
 <table WIDTH=128 HEIGHT=183><tr><td>
 
 <!--#include file="comersus_getRootCategories.asp"--> 
  
 </td></tr></table>

</TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="Images/leftbar_07.gif" WIDTH=128 HEIGHT=53 ALT=""></TD>
	</TR>
	<TR>
		<TD>
 <table WIDTH=128 HEIGHT=58 border=0><tr><td>
     <%if pShowSearchBox="-1" then%>
     <form method=post action=comersus_listItems.asp><input type=text name=strSearch size=12></a><input type=submit value=">>"></form>
     <%end if%>
 </td></tr></table>
</TD>
	</TR>
</TABLE>
		</TD>
		<TD>
		<table WIDTH=753 HEIGHT=643>
		<tr><td>
                <div style="overflow: auto; height: 637; width: 725;"> 