<!---#include file = "TopMenu.asp"--->
		<% 'On Error Resume Next
		If Request.Cookies("AdMelaUser") = "" then
		Response.Redirect "Login.asp"
		Else
		strAddedByEmail = Trim(Request.Cookies("AdMelaUser"))
		End if
		'********************************************************************************************
		
		
		Dim adClassify
    adClassify = Request.QueryString("Action")
	SubID = Request.QueryString("SubID")
	    UserID = Request.Cookies("UserID")
		
		incDate = Now
		AddedBy = Request.Cookies("MemberID")
		'Response.write UserID
		'Response.write AddedBy
		
		
		
			Sub checkUser()
					    
							Response.Write "<script language='JavaScript'>alert('You did not fillout all required fields!');</script>"
							Response.Write "<script language='JavaScript'>history.back();</script>"
							Response.End
						
					End Sub
					
	if adClassify = "strClassified" then
	
	    Dim summary, description, contact, location, email, telephone, web,o_Active
		Dim UserID, incDate, UserIDNumber
		'Dim 
		
		summary = Request("Summary")
		 summary = Replace(summary,"'","""")
		description = Request("message")
		 description = Replace(description,"'","""")
		contact = Request("myName")
		 contact = Replace(contact,"'","""")
		location = Request("Location")
		 location = Replace(location,"'","""")
		email = Request("myEmail")
		telephone = Request("myTelephone")
		web = Request("myURL")
		o_ProductID = Request("id_style")
		myProductID = Request("ADID")
		staticImage = "no_photo.gif"
		o_Active = 0
		o_deActive = Request("deactivate")
		   if o_deActive <> "" then
		      o_deActive = 2
			  Else
			  o_deActive = o_deActive
		   End if	  
		strDate = Now 
		
		'response.write o_deActive
		'response.end
		
		If summary = "" or description="" or contact="" or email="" or o_ProductID = "Select Next" then
	    Call checkUser()
		End if
		Dim incDo
		    incDo = Request.queryString("incDo")
			
	   If incDo = "add"	then 
	  
		objSql = "Insert into AM_ADS(SubCatID,ProductName,Description,URL,DateAdded,LastUpdated,ContactName,AddedBy,Location,Telephone,Email_Address,Active,UserID) Values ('"& o_ProductID &"','"& summary &"','"& description &"','"& web &"','"& strDate &"','"& strDate &"','"& contact &"','"& AddedBy &"','"& Location &"','"& telephone &"','"& email &"','"& o_Active &"','"& UserID &"')"
		Elseif incDo = "edit" then
		objSql = "Update AM_ADS set ProductName='"& summary &"',Description='"& description &"',ContactName='"& contact &"',Location='"& location &"',Email_Address='"& email &"',Telephone='"& telephone &"',URL='"& web &"',LastUpdated='"& Now &"',OnOff='"& o_deActive &"' WHERE ProductID='"& myProductID &"'"
		End if
		'response.write objSql
		'response.end
		Set objRs = DBopen.Execute(objSql)
		If incDo = "add"	then
		Response.Redirect "http://localhost/admela/myaccount.asp?LandOn=AdList"
		Else
		Response.Redirect "myClassified.asp?message=2"
		End if
		
	End if
	


Dim pullClassified, pullRs, ADID
    ADID = Request.Querystring("ADID")
    If ADID = "" then
    pullClassified = "Select * From AM_ADS where Email_Address = '"& strAddedByEmail &"' "
	ElseIf ADID <> "" Then
	 pullClassified = "Select * From AM_ADS where ADID = '"& ADID &"' AND Email_Address = '"& strAddedByEmail &"' "
	End if
	Set pullRs = Dbopen.Execute(pullClassified)
	
If Request.queryString("LandOn") <> "" then	
    'Change the names here before turning them ON
     Dim Cat_Sql, Cat_Rs, SubCat_Sql, SubCat_Rs

    Cat_Sql = "Select * From AM_Category where OnOff = 1 order by CategoryID ASC"
	Set Cat_Rs = Dbopen.Execute(Cat_Sql)
	
	'response.write Cat_Rs("CategoryName")
	'response.end
	
	SubCat_Sql = "Select * From AM_subcategory order by SubCatID ASC"
	Set SubCat_Rs = Dbopen.Execute(SubCat_Sql)
	writeMessage = SubCat_Rs("SubCatID")
	'response.write SubCat_Rs("SubCatID")
	'response.end

'************************************************************************************************

%>

<%GetParent = Request("combo0")
response.write GetParent & "-"
GetChild = Request("combo1")
response.write GetChild

    
	
	 Cat_Sql = "Select * From AM_Category where OnOff = 1 order by CategoryID ASC"
	Set Cat_Rs = Dbopen.Execute(Cat_Sql)
	
	SubCat_Sql = "Select * From AM_subcategory order by CategoryID ASC"
	Set SubCat_Rs = Dbopen.Execute(SubCat_Sql)
	writeMessage = SubCat_Rs("SubCatID")
	'response.write SubCat_Rs("SubCatID")
	'response.end


 %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<script language="JavaScript" type="text/javascript">
<!--
/*
*** Multiple dynamic combo boxes
*** by Mirko Elviro, 9 Mar 2005
*** Script featured and available on JavaScript Kit (http://www.javascriptkit.com)
***
***Please do not remove this comment
*/
// This script supports an unlimited number of linked combo boxed
// Their id must be "combo_0", "combo_1", "combo_2" etc.
// Here you have to put the data that will fill the combo boxes
// ie. data_2_1 will be the first option in the second combo box
// when the first combo box has the second option selected
 
// first combo box
<% While not Cat_Rs.Eof 
 
%>

data_<%=Cat_Rs("OrderBy")%> = new Option("<%=Cat_Rs("CategoryID")%>", "$");

			<%    SubCat_Sql = "Select * From AM_subcategory where CategoryID="& Cat_Rs("CategoryID")&" order by OrderBy ASC"
				Set SubCat_Rs = Dbopen.Execute(SubCat_Sql) %>
			        // second combo box
					<% While Not SubCat_Rs.Eof %>
			         data_<%=Cat_Rs("OrderBy")%>_<%=SubCat_Rs("OrderBy")%> = new Option("<%=SubCat_Rs("SubCategoryName")%>", "<%=SubCat_Rs("SubCatID")%>");
					<%  SubCat_Rs.MoveNext
					 Wend %>

<%  Cat_Rs.MoveNext
    wend %>

 
// other parameters
displaywhenempty=""
valuewhenempty=-1
displaywhennotempty="-select-"
valuewhennotempty=0
 
function change(currentbox) {
numb = currentbox.id.split("_");
currentbox = numb[1];
i=parseInt(currentbox)+1
// I empty all combo boxes following the current one
while ((eval("typeof(document.getElementById(\"combo_"+i+"\"))!='undefined'")) &&
(document.getElementById("combo_"+i)!=null)) {
son = document.getElementById("combo_"+i);
// I empty all options except the first one (it isn't allowed)
for (m=son.options.length-1;m>0;m--) son.options[m]=null;
// I reset the first option
son.options[0]=new Option(displaywhenempty,valuewhenempty)
i=i+1
}
 
// now I create the string with the "base" name ("stringa"), ie. "data_1_0"
// to which I'll add _0,_1,_2,_3 etc to obtain the name of the combo box to fill
stringa='data'
i=0
while ((eval("typeof(document.getElementById(\"combo_"+i+"\"))!='undefined'")) &&
(document.getElementById("combo_"+i)!=null)) {
eval("stringa=stringa+'_'+document.getElementById(\"combo_"+i+"\").selectedIndex")
if (i==currentbox) break;
i=i+1
}
 
// filling the "son" combo (if exists)
following=parseInt(currentbox)+1
if ((eval("typeof(document.getElementById(\"combo_"+following+"\"))!='undefined'")) &&
(document.getElementById("combo_"+following)!=null)) {
son = document.getElementById("combo_"+following);
stringa=stringa+"_"
i=0
while ((eval("typeof("+stringa+i+")!='undefined'")) || (i==0)) {
// if there are no options, I empty the first option of the "son" combo
// otherwise I put "-select-" in it
if ((i==0) && eval("typeof("+stringa+"0)=='undefined'"))
if (eval("typeof("+stringa+"1)=='undefined'"))
eval("son.options[0]=new Option(displaywhenempty,valuewhenempty)")
else
eval("son.options[0]=new Option(displaywhennotempty,valuewhennotempty)")
else
eval("son.options["+i+"]=new Option("+stringa+i+".text,"+stringa+i+".value)")
i=i+1
}
//son.focus()
i=1
combostatus=''
cstatus=stringa.split("_")
while (cstatus[i]!=null) {
combostatus=combostatus+cstatus[i]
i=i+1
}
return combostatus;
}
}
//-->
</script>


	
		
	<%	End if 
		'******************************************************************************************
		%>
		
		
		 
		</td>
        </tr>
     
      <tr>
        <td colspan="5" valign="top"><table width="800" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="190px" valign="top"><table width="190" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="top">
				        <!---#include file = "LeftNav.asp"--->
				   
				   </td>
                                    </tr>
                                    <tr>
                                      <td><img src="image/right-bot.jpg" width="194" height="25" /></td>
                                    </tr>
                                </table>
								
					  </td>
								
              </tr>
			  <tr>
			  <td align="center"  valign="top"  height="10px"></td>
			  </tr>
              <tr>
                <td>
				
				<table width="190" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td></td>
                              </tr>
                              <tr>
                                
                              </tr>
                          </table>
						  
						 
						  
					  </td>
              </tr>
			  <tr>
			  <td align="center"  valign="top"  height="10px"></td>
			  </tr>
              
            </table>
				<!----removed--->
				
				</td>
			<td width="30px">&nbsp;</td>
            <td valign="top"><table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="image/blog-bg-top.jpg" /></td>
                </tr>
              <tr>
                <td class="blogbg"><table width="530" border="0" cellspacing="0" cellpadding="0" align="center">
                  <tr>
                    <td colspan="3"><table width="530" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#FFFFFF"><img src="image/People.png" height="75" width="75"  align="left"/></td>
                        <td width="0px" bgcolor="#FFFFFF"></td>
						<td class="centerheading" align="left" bgcolor="#FFFFFF">My Account</td>
                      </tr>
                      <tr>
                   <td align="right" colspan="3" bgcolor="#FFFFFF" class="matter"><A href="myaccount.asp?LandOn=profile">my account</a> | <A href="myaccount.asp?LandOn=AdList">my ads</a> | <A href="myaccount.asp?LandOn=SavedList">my saved ads</a> &nbsp;
				   <BR><hr color="#FF8040" /></td>
                      </tr>
                    </table></td>
                    </tr>
					
					
                  <tr bgcolor="#FFFFFF"> 
				  
		<% If Request.queryString("LandOn") = "SavedList" then %>		  
                    <td colspan="3" height="225px" class="matter" valign="top"> Here is the list of Ads you've SAVED
					<BR>
					<!---insert here--->
					<%   If Request.Cookies("AdMelaUser") <> "" then
				         
						 'Lets check if has saved any listing before or not
						 chkAdSQL = "Select * From AM_SavedAds where AddedByEmail='"& strAddedByEmail &"' order by DateAdded DESC"
						 set chkRs = DbOpen.Execute(chkAdSQL) %>
						<TABLE cellpadding="5" bgcolor="#ffffff" cellspacing="3" width="100%" align="center">
							<TH bgcolor="#69ACB8"><Font color="#FFFFFF">AD ID</FONT></th>
				            <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Description</FONT></th>
				            <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Saved On</FONT></th>
							<TH bgcolor="#69ACB8"><Font color="#FFFFFF">Action</FONT></th>
						 		<%chkRs.MoveFirst 
					  While not chkRs.Eof %>
				  <TR>
				      
				      <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= chkRs("ADID") %></font></td>
				      <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= chkRs("AdTitle") %></font></td>
					  <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= chkRs("DateAdded") %></font></td>
					   <TD bgcolor="#E8E8E8"><a href="AdDetails.asp?ADID=<%= chkRs("ADID") %>&SubCatName=&IMG_Path=">View Ad</a></td>
					  
					  </td>
				   </tr>	  
					 <% chkRs.MoveNext
					    Wend
						End if %>
					<!----ends here--->
					
					</table>
			<% ElseIf Request.queryString("LandOn") = "AdList" then 
			
			    Dim myEmail, mySql, myRs
				myEmail = Request.Cookies("AdMelaUser")
				
				mySql = "Select * From AM_ADS AS A INNER JOIN AM_Status As S on A.Active = S.StatusValue Where A.Email_Address = '"& myEmail &"' order by A.ADID DESC"
				Set myRs = Dbopen.Execute(mySql)
			
			%>		
			<td colspan="3" height="225px" class="matter" valign="top"> Here is the list of Ads you've submitted
					<TABLE cellpadding="5" bgcolor="#ffffff" cellspacing="3" width="100%" align="center">
								
							<TH bgcolor="#69ACB8"><Font color="#FFFFFF">AD ID</FONT></th>
				            <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Description</FONT></th>
				            <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Submitte On</FONT></th>
							<TH bgcolor="#69ACB8"><Font color="#FFFFFF">Status</FONT></th>
				         
					<% while not myRs.Eof %>		
				<TR>
				     <td><A href="myaccount.asp?LandOn=EditAd&ADID=<%= myRs("ADID") %>&incDo=edit"><%= myRs("ADID") %></a></td>
					 <td><%= myRs("ProductName") %></td>
					  <td><%= myRs("DateAdded") %></td>
					 <td><%= myRs("StatusName") %></td>
					
				</tr>	
				<% myRs.MoveNext
				   Wend %>
				</Table>				 
						 
			<% ElseIf Request.queryString("LandOn") = "AddAd" then
			          SubCat_Rs.MoveFirst
			          'response.write SubCat_Rs("SubCatID")
					  'response.end
			
			 %>	
				<td colspan="3" height="225px" class="matter" valign="top"> Submit your ad here
					<BR>
				<table cellpadding="5" width="100%">	
				<FORM name=product method=post target="" action="MyAccount.asp?Action=strClassified&incDo=add&SubID=<%= SubID %>"> 		
								<TR>     <TD><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Choose Category:</font></td>
								          <TD><select name="combo0" id="combo_0" onChange="change(this);" style="width:200px;">
												<option value="value0">-select-</option>
												<% Cat_Rs.MoveFirst
												While not Cat_Rs.Eof %>
												<option value="<%=Cat_Rs("OrderBy")%>"><%=Cat_Rs("CategoryName")%></option>
												<%  Cat_Rs.MoveNext
												    wend %>
												
												</select>
												<BR><BR>
												<select name="combo1" id="combo_1" onChange="change(this)" style="width:200px;">
												<option value="value1"> </option>
												</select>
												<BR><BR></td>
								
								</tr>		  
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Ad Name/ summary:</td>
									<td><input type="text" name="Summary" maxlength="50" size="50"></td>
								</tr>
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Ad Description:</td>
									<td>
									 <textarea name=message wrap=physical cols=40 rows=4 onKeyDown="textCounter(this.form.message,this.form.remLen,300);" onKeyUp="textCounter(this.form.message,this.form.remLen,300);"></textarea>
								                         <br>
								                         <input readonly type=text name=remLen size=3 maxlength=3 value="300"> <font size="2">Max Charecter.</font> 
												
									
									
									</td>
								</tr>
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Contact Name:</td>
									<td><input type="text" name="myName" maxlength="30" size="50"></td>
								</tr>
								<tr>
									<td><font color="#004080" size="2">Location:</td>
									<td><input type="text" name="Location" maxlength="30" size="50"></td>
								</tr>
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Contact Email:</td>
									<td><%= Request.Cookies("AdMelaUser") %><input type="hidden" name="myEmail" value="<%= Request.Cookies("AdMelaUser") %>" maxlength="30" size="50"></td>
								</tr>
								<tr>
									<td><font color="#004080" size="2">Contact Telephone:</td>
									<td><input type="text" name="myTelephone" maxlength="30" size="50"></td>
								</tr>
								<tr>
									<td><font color="#004080" size="2">Web Address:</td>
									<td><input type="text" name="myURL" maxlength="50" size="50"></td>
								</tr>
								<tr>
									<td></td>
									<td>&nbsp;<input class="submit" type="submit" name="incMessage" value="Submit Ad!"></td>
								</tr>
								</Form>
								</table>
						<% ElseIf Request.queryString("LandOn") = "profile" then %>		
						
						<% Dim my_Sql, my_Rs
						       my_Sql = "Select * From AM_Member M INNER JOIN AM_Users U ON M.EmailAddress = U.UserName where M.EmailAddress='"& strAddedByEmail &"'"
							   Set my_Rs = DbOpen.Execute(my_Sql)
							   
						   
						 %>
						  
						        <tr bgcolor="#FFFFFF" class="matter" >
									<td align="Left">First Name:</td>
									<td>&nbsp;</td>
									<td align="left">&nbsp;<%= my_Rs("FirstName") %></td>
									
								</tr>
								  <tr bgcolor="#FFFFFF" class="matter">
									<td align="Left">Last Name:</td>
									<td>&nbsp;</td>
									<td align="left">&nbsp;<%= my_Rs("LastName") %></td>
									
								</tr>
								  <tr bgcolor="#FFFFFF" class="matter">
									<td align="Left">Zip Code:</td>
									<td>&nbsp;</td>
									<td align="left">&nbsp;<%= my_Rs("ZipCode") %></td>
									
								</tr>
								  <tr bgcolor="#FFFFFF" class="matter">
									<td align="Left">Email Address:</td>
									<td>&nbsp;</td>
									<td align="left">&nbsp;<%= my_Rs("EmailAddress") %></td>
									
								 </tr>
								   <tr bgcolor="#FFFFFF" class="matter">
									<td align="Left">Password:</td>
									<td>&nbsp;</td>
									<td align="left">&nbsp;<%= my_Rs("Password") %></td>
									
								</tr>
								   <tr bgcolor="#FFFFFF" class="matter">
									<td align="Left">Member Since:</td>
									<td>&nbsp;</td>
									<td align="left">&nbsp;<%= my_Rs("MemberSince") %></td>
									
								</tr>
								</tr>
								   <tr bgcolor="#FFFFFF" class="matter">
									<td align="Left">Last Login:</td>
									<td>&nbsp;</td>
									<td align="left">&nbsp;<%= my_Rs("LastLogin") %></td>
									
								</tr>
						   
						<% 
						
						End if %>		
								<% if request.queryString("LandOn") = "EditAd" then %>
								 <td colspan="3" height="225px" class="matter" valign="top">
											  <table class="matter">
								<FORM name=Register method=post onSubmit="return verifyEmail(this)" target="" action="myClassified.asp?Action=strClassified&incDo=update"> 	
								<input type="hidden" name="ProductID" value="<%= pullRs("ADID") %>">	
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">DE-ACTIVATE AD?:</td>
									<td><input type="Checkbox" name="deactivate" value="<%= pullRs("Active") %>" <% If pullRs("Active") = 2 then response.write "checked" end if %>> <font color="#FF0000" size="2">Check for YES</FONT></td>
								</tr>		  
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Ad summary:</td>
									<td><input type="text" name="Summary" maxlength="50" size="50" value="<%= pullRs("ProductName") %>"></td>
								</tr>
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Ad Description:</td>
									<td>
									 <textarea name=message wrap=physical cols=40 rows=6 onKeyDown="textCounter(this.form.message,this.form.remLen,300);" onKeyUp="textCounter(this.form.message,this.form.remLen,300);"><%= pullRs("Description") %></textarea>
								                         <br>
								                         <input readonly type=text name=remLen size=3 maxlength=3 value="300"> <font size="2">Max Charecter.</font> 
												
									
									
									</td>
								</tr>
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Contact Name:</td>
									<td><input type="text" name="myName" maxlength="30" size="50" value="<%= pullRs("ContactName") %>"></td>
								</tr>
								<tr>
									<td><font color="#004080" size="2">Location:</td>
									<td><input type="text" name="Location" maxlength="30" size="50" value="<%= pullRs("Location") %>"></td>
								</tr>
								<tr>
									<td><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Contact Email:</td>
									<td><input type="text" name="myEmail" maxlength="30" size="50" value="<%= pullRs("Email_Address") %>"></td>
								</tr>
								<tr>
									<td><font color="#004080" size="2">Contact Telephone:</td>
									<td><input type="text" name="myTelephone" maxlength="30" size="50" value="<%= pullRs("Telephone") %>"></td>
								</tr>
								<tr>
									<td><font color="#004080" size="2">Web Address:</td>
									<td><input type="text" name="myURL" maxlength="50" size="50" value="<%= pullRs("URL") %>"></td>
								</tr>
								<tr>
									<td></td>
									<td><input type="button" name="incBack" value="Return Back" class="submit" onclick="history.back();">&nbsp;<input class="submit" type="submit" name="incMessage" value="Update Ad!"></td>
								</tr>
								</Form>
								</table>
					
					            <% End if %>
					
					
					
					
					
					
					
					
					
					
					
					
					
					
					
					
					
					
					
					<!---end here--->
					</td>
                    </tr>
                  <tr>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                </table></td>
                </tr>
              <tr>
                <td><img src="image/blog-bg-bottom.jpg" /></td>
                </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
     
      
    </table></td>
  </tr>
  
</table>

<!---#include file = "Footer.asp"--->
</body>
</html>
