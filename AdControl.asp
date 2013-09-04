<% on error resume next
Dim adClassify
    adClassify = Request.QueryString("Action")
	SubID = Request.QueryString("SubID")
	    UserID = Request.Cookies("UserID")
		incDate = Now
		UserIDNumber = Request.Cookies("MemberID")
		
		
		
			Sub checkUser()
					    
							Response.Write "<script language='JavaScript'>alert('You did not fillout all required fields!');</script>"
							Response.Write "<script language='JavaScript'>history.back();</script>"
							Response.End
						
					End Sub
					
	if adClassify = "strClassified" then
	
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
		
		If summary = "" or description="" or contact="" or email="" or o_ProductID = "Select Next" then
	    Call checkUser()
		End if
		Dim incDo
		    incDo = Request.queryString("incDo")
			
	   If incDo = "add"	then 
	  
		objSql = "Insert into AM_ADS(SubCatID,ProductName,Description,ImagePath,URL,DateAdded,ContactName,AddedBy,Location,Telephone,Email_Address,Action,UserID) Values 
		                            ('"& o_ProductID &"','"& summary &"','"& description &"','"& staticImage &"','"& contact &"','"& location &"','"& email &"','"& telephone &"','"& web &"','"& incDate &"','"& UserIDNumber &"','"& UserID &"')"
		Else
		
		objSql = "Update juty_products set ProductName='"& summary &"',Description='"& description &"',ContactName='"& contact &"',Location='"& location &"',Email_Address='"& email &"',Telephone='"& telephone &"',URL='"& web &"',LastUpdated='"& Now &"',OnOff='"& o_deActive &"' WHERE ProductID='"& myProductID &"'"
		End if
		'response.write objSql
		'response.end
		Set objRs = DBopen.Execute(objSql)
		If incDo = "add"	then
		Response.Redirect "myClassified.asp?message=1"
		Else
		Response.Redirect "myClassified.asp?message=2"
		End if
		
	End if
	


Dim pullClassified, pullRs, ADID
    ADID = Request.Querystring("ADID")
    If ADID = "" then
    pullClassified = "Select * From juty_products where AddedBy = '"& UserIDNumber &"' "
	ElseIf ADID <> "" Then
	 pullClassified = "Select * From juty_products where ProductID = '"& ADID &"' AND AddedBy = '"& UserIDNumber &"' "
	End if
	Set pullRs = Dbopen.Execute(pullClassified)
	
	
    'On Error resume next
Dim ad_Sql, ad_Rs

    ad_Sql = "Select * From juty_classified where OnOff = 1 order by ID ASC"
	Set ad_Rs = Dbopen.Execute(ad_Sql)
	
	p_Sql = "Select * From juty_subcategory order by CategoryID ASC"
	Set p_Rs = Dbopen.Execute(p_Sql)



 


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
</script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">

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
      <% While not ad_Rs.Eof %>
			new Type(<%= ad_Rs("ID") %>, "<%= ad_Rs("CategoryName") %> "),
        <%  ad_Rs.MoveNext
		    wend %>
			new Type(5, "")

	);



	StyleArray = new Array(
       <% While not p_Rs.Eof %>
			new Style(<%= p_Rs("Sub_ID") %>, <%= p_Rs("CategoryID") %>, "<%= p_Rs("SubCategoryName") %>"),
       <% p_Rs.MoveNext
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
