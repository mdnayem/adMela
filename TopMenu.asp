<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="keywords" content="Bangladesh, Dhaka,  ejuty, marketing, bangla, bangali, bengal, bengali, bangladeshi, deshi commercial, deshi biggapon, deshi ads">
<meta name=description content="Bangla ads, advertisements, bangladeshi abroad, sell, buy, service">
<META 
content="Bangladeshi advertisement, Quality Bangladeshi biggapon center" 
name=description>
<META 
content="Bangladeshi Matrimonial , Matrimonial , bangladeshi Matrimonials , Matrimonials , Bangladesh Marriage , Match-Making , Bangladesh Marriages , Internet Matrimonials , bangladeshi Sites , netMela , Love, Sex, Bor, Koney, Brides , Grooms , Life-partner , Dhaka , Comilla , Mymensing , Noakhali , Chittagong , New York , America , Montreal , Toronto , London , America , Bangla , Rajshahi , Bogra , Dinajpur, wedding , Laksam matrimonials , Sylhet matrimonial , Rangpur matrimonials , bangladesh matrimonial , abu dhabi , sharjah , marriage proposal , Bangladesh , Bangali , Bangladesh  , Dallywood , bangladeshi brides , Bangladeshi grooms , wife , husband , dating , Bangladeshi dating , matrimonial services , Bangladeshi matrimonial services" 
name=keywords>
<title>Ad Mela for Bangladeshis Abroad!</title>
<link rel = "Shortcut Icon" href = "image/favicon.ico" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
</head>
<!---#include file = "IncDSN.asp"--->
<% strName = Request.Cookies("AdMelaUserName")
   strHello = Request.Cookies("Hello")
 %>
<body>
<table width="836" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td valign="top"><table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td colspan="5" height="83px" valign="top" align="center" class="online">Beta Release<table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td height="57px"  width="400"align="left" valign="bottom"><A href="Default.asp"><img src="Image/melaLogo.png" border="0" /></a> <Font size="-1" color="#FF8040" face="Verdana">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%= strHello %>&nbsp;<%= strName %></font></td>
            <td height="57px" width="400" align="center" valign="bottom">
			<Form name="searcform" method="post" action="Search.asp">
			<input name="SearchWord" type="text" value="Coming Soon!" size="30" disabled/>&nbsp;<input type="Image" align="absbottom" src="image/search.jpg" height="22" name="searchbutton"  onclick="this.form.submit();" align="bottom" /></form></td>
          </tr>
        
        </table></td>
        </tr>
      <tr>
        <td colspan="5" height="82px" valign="top"><table width="800" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td class="left"></td>
            <td background="Image/business-top-bg.jpg" width="784px"> </td>
            <td class="right"></td>
          </tr>
          <tr>
            <td height="66px" width="8px" background="Image/business-bg.jpg">&nbsp;</td>
            <td height="66px" width="784px" background="Image/business-bg.jpg" class="homelink">
			<a href="Default.asp" class="homelink" title="home">Home</a>         |        <a href="MyAccount.asp?LandOn=AddAd" class="homelink" title="submit your ad"> PostAd  </a>       |     <a href="contact.asp" class="homelink" title="contact us">    Contact   </a>      |  <% If Request.Cookies("AdMelaUser") = "" then %>     <a href="Login.asp" class="homelink" title="login">  Login     </a>         |        <a href="Registration.asp" class="homelink" title="join ad-mela"> Register  </a> <% Else %>      <a href="MyAccount.asp?LandOn=profile" class="homelink" title="my account"> MyAccount  </a>   |   <a href="Login.asp?Action=Logout" class="homelink" title="logout"> Logout  </a> <% End if %></td>
            <td height="66px" width="8px" background="Image/business-bg.jpg">&nbsp;</td>
          </tr>

        </table>
