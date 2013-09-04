<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Ad-Mela</title>
</head>

<SCRIPT LANGUAGE="JavaScript">
<!--
 function nameDefined(ckie,nme)
{
   var splitValues
   var i
   for (i=0;i<ckie.length;++i)
   {
      splitValues=ckie[i].split("=")
      if (splitValues[0]==nme) return true
   }
   return false
}

function delBlanks(strng)
{
   var result=""
   var i
   var chrn
   for (i=0;i<strng.length;++i) {
      chrn=strng.charAt(i)
      if (chrn!=" ") result += chrn
   }
   return result
}
function getCookieValue(ckie,nme)
{
   var splitValues
   var i
   for(i=0;i<ckie.length;++i) {
      splitValues=ckie[i].split("=")
      if(splitValues[0]==nme) return splitValues[1]
   }
   return ""
}
function insertCounter() {
 readCookie()
 displayCounter()
}
 function displayCounter() {
 document.write('<H5 ALIGN="CENTER">')
 document.write("Welcome, you're the # ")
 if(counter==1) document.write(" ")
 else document.write(counter+" visitor")
 document.writeln('</H5>')
 }
function readCookie() {
 var cookie=document.cookie
 counter=0
 var chkdCookie=delBlanks(cookie)  //are on the client computer
 var nvpair=chkdCookie.split(";")
 if(nameDefined(nvpair,"pageCount"))
 counter=parseInt(getCookieValue(nvpair,"pageCount"))
 ++counter
 var futdate = new Date()
 var expdate = futdate.getTime()
 expdate += 3600000 * 24 *30  //expires in 1 hour
 futdate.setTime(expdate)

 var newCookie="pageCount="+counter
 newCookie += "; expires=" + futdate.toGMTString()
 window.document.cookie=newCookie
}
// -->
</SCRIPT>

<script language="JavaScript">
TargetDate = "07/30/2012 5:00 AM";
BackColor = "yellow";
ForeColor = "navy";
CountActive = true;
CountStepper = -1;
LeadingZero = true;
DisplayFormat = "%%D%% Days, %%H%% Hours, %%M%% Minutes, %%S%% Seconds.";
FinishMessage = "It is finally here!";
</script>


<body>
<table width="190" border="0" cellspacing="0" cellpadding="0" align="center">
                              <TR>
							        
							        <TD align="center">
									<FONT color="#0000FF" size="+1">If you have any question please contat at <A href="mailto:postmaster@ad-mela.com">postmaster@ad-mela.com</A></FONT>
									<BR><FONT color="#FF0000" size="+2">Site launch countdown:</FONT>
									<script language="JavaScript" src="http://scripts.hashemian.com/js/countdown.js"></script></td>
							  </tr>
                              <tr>
                                <td><img src="Image/comingsoon.bmp" border="0"/></td>
                              </tr>
                            
                          </table>

</body>
</html>
