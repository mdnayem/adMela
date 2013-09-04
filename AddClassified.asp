<!---#include file = "TopMenu.asp"--->
		<% 'On Error Resume Next
		'If Request.Cookies("AdMelaUser") = "" then
		'Response.Redirect "Login.asp"
		'Else
		'strAddedByEmail = Trim(Request.Cookies("AdMelaUser"))
		'End if
		'********************************************************************************************
		
		
		Dim adClassify
    adClassify = Request.QueryString("Action")
	SubID = Request.QueryString("SubID")
	    UserID = Request.Cookies("UserID")
		incDate = Now
		UserIDNumber = Request.Cookies("MemberID")
		
		
		
			'Sub checkUser()
					    
							'Response.Write "<script language='JavaScript'>alert('You did not fillout all required fields!');</script>"
							''Response.Write "<script language='JavaScript'>history.back();</script>"
							'Response.End
						
					'End Sub
					
	'if adClassify = "strClassified" then
	
	    Dim summary, description, contact, location, email, telephone, web
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
		myProductID = Request("ProductID")
		staticImage = "no_photo.gif"
		o_deActive = Request("deactivate")
		   if o_deActive <> "" then
		      o_deActive = 2
			  Else
			  o_deActive = o_deActive
		   End if	  
		
		
		'response.write o_deActive
		'response.end
		
		'If summary = "" or description="" or contact="" or email="" or o_ProductID = "Select Next" then
	   'Call checkUser()
		'End if
		Dim incDo
		    incDo = Request.queryString("incDo")
			
	   'If incDo = "add"	then 
	  
		'objSql = "Insert into juty_products(SubID,ProductName,Description,ImagePath,ContactName,Location,Email_Address,Telephone,URL,DateAdded,AddedBy,UserID) Values ('"& o_ProductID &"','"& summary &"','"& description &"','"& staticImage &"','"& contact &"','"& location &"','"& email &"','"& telephone &"','"& web &"','"& incDate &"','"& UserIDNumber &"','"& UserID &"')"
		'Else
		
		'objSql = "Update juty_products set ProductName='"& summary &"',Description='"& description &"',ContactName='"& contact &"',Location='"& location &"',Email_Address='"& email &"',Telephone='"& telephone &"',URL='"& web &"',LastUpdated='"& Now &"',OnOff='"& o_deActive &"' WHERE ProductID='"& myProductID &"'"
		'End if
		'response.write objSql
		'response.end
		'Set objRs = DBopen.Execute(objSql)
		'If incDo = "add"	then
		'Response.Redirect "myClassified.asp?message=1"
		'Else
		'Response.Redirect "myClassified.asp?message=2"
		'End if
		
	'End if
	


Dim pullClassified, pullRs, ADID
    ADID = Request.Querystring("ADID")
    If ADID = "" then
    pullClassified = "Select * From AM_ADS where Email_Address = '"& strAddedByEmail &"' "
	ElseIf ADID <> "" Then
	 pullClassified = "Select * From AM_ADS where ADID = '"& ADID &"' AND Email_Address = '"& strAddedByEmail &"' "
	End if
	Set pullRs = Dbopen.Execute(pullClassified)
	
	
    'Change the names here before turning them ON
Dim Cat_Sql, Cat_Rs

    Cat_Sql = "Select * From AM_Category where OnOff = 1 order by CategoryID ASC"
	Set Cat_Rs = Dbopen.Execute(Cat_Sql)
	
	'response.write Cat_Rs("CategoryName")
	'response.end
	
	SubCat_Sql = "Select * From AM_subcategory order by SubCatID ASC"
	Set SubCat_Rs = Dbopen.Execute(SubCat_Sql)

'************************************************************************************************

%>
<SCRIPT LANGUAGE="JavaScript">

<!-- Begin
function textCounter(field, countfield, maxlimit) {
if (field.value.length > maxlimit) // if too long...trim it!
field.value = field.value.substring(0, maxlimit);
// otherwise, update 'characters left' counter
else 
countfield.value = maxlimit - field.value.length;
}
// End -->

<!--
	//	Initialize class for Type and Style

	function Type(id, type){
		this.id = id;
		this.type = type;
	}
	function Style(id, id_type, style){
		this.id = id;
		this.id_type = id_type;
		this.style = style;
	}

    //	Initialize Array's Data for Type and Style

	TypeArray = new Array(
      <% While not Cat_Rs.Eof %>
			new Type(<%= Cat_Rs("CategoryID") %>, "<%= Cat_Rs("CategoryName") %> "),
        <%  Cat_Rs.MoveNext
		    wend %>
			new Type(5, "")

	);

	StyleArray = new Array(
       <% While not SubCat_Rs.Eof %>
			new Style(<%= SubCat_Rs("SubCatID") %>, <%= SubCat_Rs("CategoryID") %>, "<%= SubCat_Rs("SubCategoryName") %>"),
       <% SubCat_Rs.MoveNext
	      Wend %>
			new Style(31, 5, "")
	);
	function init(sel_type, sel_style){
		document.product.id_type.options[0]	= new Option("Select First");
		document.product.id_style.options[0] = new Option("Select Next");
		for(i = 1; i <= TypeArray.length; i++){
			document.product.id_type.options[i]	= new Option(TypeArray[i-1].type, TypeArray[i-1].id);
			if(TypeArray[i-1].id == sel_type)
				document.product.id_type.options[i].selected = true;
		}
		OnChange(sel_style);
	}
	function OnChange(sel_style){
		sel_type_index = document.product.id_type.selectedIndex;
		sel_type_value = parseInt(document.product.id_type[sel_type_index].value);
		for(i = document.product.id_style.length - 1; i > 0; i--)
			document.product.id_style.options[i]	= null;
		        j=1;
		    for(i = 1; i <= StyleArray.length; i++){
			 if(StyleArray[i-1].id_type == sel_type_value){
				document.product.id_style.options[j]	= new Option(StyleArray[i-1].style, StyleArray[i-1].id);
				if(StyleArray[i-1].id == sel_style)	document.product.id_style.options[j].selected = true;
				j++;
			}
		}
	}

//-->
</SCRIPT>



	




	
		
	<table cellpadding="5" width="100%">	
				<FORM name=product method=post target="" action="myClassified.asp?Action=strClassified&incDo=add&SubID=<%= SubID %>"> 		
								<TR>
								          <TD><font color="#FF0000" size="2">*</font><font color="#004080" size="2">Choose Category:</font></td>
										  <TD><select name="id_type" size="1" style="width: 150px;" onChange="OnChange()">
																	
								  </select>
								
									     
								
								  <select name="id_style" size="1" style="width: 150px;">
								  
								  </select></td>
								
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
									<td><input type="text" name="myEmail" maxlength="30" size="50"></td>
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
				
						<!---#include file = "Footer.asp"--->	