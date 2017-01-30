<%
pPayPalExpressCheckout	= getSettingKey("pPayPalExpressCheckout")
pGoogleCheckoutLevel1	= getSettingKey("pGoogleCheckoutLevel1")
%>
</div>
</td>
              </tr></table></TD>
		<TD>
			<IMG SRC="Images/content_04.gif" WIDTH=44 HEIGHT=643 ALT=""></TD>
	</TR>
</TABLE>
		
		</TD>
	</TR>
	<TR>
		<TD>

      <TABLE WIDTH=960 BORDER=0 CELLPADDING=0 CELLSPACING=0>
        <TR> 
          <TD> <a href="http://www.comersus.com"><IMG SRC="Images/footer_01.gif" WIDTH=169 HEIGHT=44 ALT="Comersus ASP Cart" border=0></a></TD>
          <TD> <IMG SRC="Images/footer_02.gif" WIDTH=622 HEIGHT=44 ALT=""></TD>
          <TD> <IMG SRC="Images/footer_03.gif" WIDTH=169 HEIGHT=44 ALT=""></TD>
        </TR>
        <TR>
          <TD>&nbsp;</TD>
          <TD>
            <div align="center">

<a href="comersus_index.asp">Home</a>&nbsp;&nbsp;
<a href="comersus_listCategories.asp">Categories</a>&nbsp;&nbsp;
<a href="comersus_customerUtilitiesMenu.asp">Your Account</a>&nbsp;&nbsp;
<a href="comersus_newsPageExec.asp">News</a>&nbsp;&nbsp;
<a href="comersus_advancedSearchForm.asp">Search</a>&nbsp;&nbsp;
<a href="comersus_listItems.asp?lastChance=1">Last Chance</a>&nbsp;&nbsp;
<a href="comersus_listItems.asp?orderBy=sales">Top Sellers</a>&nbsp;&nbsp;

<%if pAuctions="-1" then%>
 <a href="comersus_optAuctionListAll.asp">Auctions</a>&nbsp;&nbsp;
<%end if%>

<%if pNewsLetter="-1" then%>
 <a href="comersus_optNewsLetterAddEmailForm.asp">Newsletter</a>&nbsp;&nbsp;
<%end if%>

<br> 
<a href="comersus_wapInfo.asp">WAP</a>&nbsp;&nbsp;

<%if pAllowNewCustomer="-1" then%>
 <a href="comersus_customerRegistrationForm.asp">Register</a>&nbsp;&nbsp;
<%end if%>

<%if pSuppliersList="-1" then%>
 <a href="comersus_optListSuppliers.asp">Suppliers List</a>&nbsp;&nbsp;
<%end if%>

<%if pAffiliatesStoreFront="-1" then%>
 <a href="comersus_optAffiliateRegistrationForm.asp">Affiliate Registration</a>&nbsp;&nbsp;
<%end if%>

<a href="comersus_storeSurveyForm.asp">Survey</a>&nbsp;&nbsp;
<a href="comersus_contactUs.asp">Contact us</a>

<br><a href="comersus_conditions.asp" target="_blank"><%=getMsg(470,"terms")%></a>

<br><br>

<%if pRssFeedServer="-1" then%>
 <a href="comersus_rssInfo.asp"><img src="images/rssfeedicon.gif" border=0 alt="RSS Feeds Information">
<%end if%>

<br><br><a href="http://www.comersus.com">Powered by Comersus ASP Shopping Cart Application</a>
                    
            </div>
          </TD>
          <TD>&nbsp;</TD>
        </TR>
      </TABLE>		
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>