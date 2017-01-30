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
	
	<style type="text/css">


.F_VeryLight
{
	color:#ffffff;
}
.MSC_HeaderText{font-size:24px;margin:0}

.MSC_HeaderDescription
{
	font-size:12px;
	margin:0px;
}
	


.F_Light
{
	color:#FF9900;
}
   
.MSC_SiteWidth,.MS_MasterHeader,.MS_MasterGlobalLinks,.MS_MasterPrimaryNav,.MS_MasterFooter,.MS_MasterTopAD,.MS_MasterBottomAD
{
	width:100%;
}



.MSC_SiteWidth,.MS_MasterHeader,.MS_MasterGlobalLinks,.MS_MasterPrimaryNav,.MS_MasterFooter,.MS_MasterTopAD,.MS_MasterBottomAD
{
	width:1000;
}



.BG_Dark

	body {
	background-repeat: no-repeat;
}
body {
	background-repeat: no-repeat;
}
.style2 {
	font-size: 24px
}
.style12 {
	font-size: 12px;
	font-style: italic;
}
    </style>
	
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 background="Images/pixelbg.gif">



<table width="1100" height="121" border="0" cellpadding="0" cellspacing="0" class="MSC_SiteWidth BG_Dark">

<tr>
                <td width="163" align="center" valign="middle" background="images/top0.jpg" style="background-repeat:repeat-x" class="BG_Dark">&nbsp;</td>
                <td width="278"  valign="top" background="images/top.1.jpg"  class="BG_Dark" style="background-repeat:no-repeat;"><span class="BG_Dark" style="background-repeat:no-repeat;"><img src="images/top.1.jpg" width="278"  height="121" align="top" > </td>
    <td width="408"    align="left" valign="middle" background="images/top3.jpg" class="menu02 style2" style="background-repeat:no-repeat; background-position:left">

<div align="center">OneStopSofas.com</div>                <div align="right" class="style12">
                    <div align="center">Where   Taste Is A Matter Of Choice<br> 
                      And Quality Is A Matter Of Fact</div>
    </div>                  </td>
    <td width="151" background="images/top4.jpg" style="background-repeat:repeat-x" class="BG_Dark">&nbsp;</td>
  </tr>
</table>


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