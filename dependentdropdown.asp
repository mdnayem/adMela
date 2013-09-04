<% 
Set DbOpen = Server.CreateObject("ADODB.Connection")
   connectString ="Provider=SQLOLEDB.1;uid=admela;pwd=lailabanu;database=AdMela;Data source=SUNOCO-PC\SQLEXPRESS;"
   DbOpen.Open ConnectString

   
   	   'On Error resume next
	   'I am adding following SQL statement on include file because this need for all pages for leftnav menus.
        Dim x_Sql, x_Rs

        C_Sql = "Select * From AM_Category where OnOff = 1 order by CategoryID ASC"
	    Set C_Rs = Dbopen.Execute(C_Sql)
		
		
	
 %>	
   
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
 <html xmlns="http://www.w3.org/1999/xhtml">
 <head>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
 <script language="javascript" type="text/javascript">
 function dropdownlist(listindex)
 {
 
document.formname.subcategory.options.length = 0;
 switch (listindex)
 {
<%  while not C_Rs.Eof 
      x_Sql = "Select * From AM_SubCategory WHERE CategoryID='"& C_Rs("CategoryID")&"' AND OnOff = 1 order by CategoryID ASC"
	    Set x_Rs = Dbopen.Execute(x_Sql)
	 %>
case "<%= C_Rs("CategoryID")%>" :
 document.formname.subcategory.options[0]=new Option("Select Sub-Category","");
 <% x_Rs.MoveFirst
   while not x_Rs.Eof %>
 document.formname.subcategory.options[<%= x_Rs("SubCatID")%>]=new Option("<%= x_Rs("SubCategoryName")%>","<%= x_Rs("SubCategoryName")%>");
 <% x_Rs.MoveNext
   Wend %>
 
break;
 <% C_Rs.MoveNext
   Wend %>
}
 return true;
 }
 </script>
 </head>
 <title>Dynamic Drop Down List</title>
 <body>
 
<form id="formname" name="formname" method="post" action="submitform.asp" >
 <table width="50%" border="0" cellspacing="0" cellpadding="5">
 <tr>
 <td width="41%" align="right" valign="middle">Category :</td>
 <td width="59%" align="left" valign="middle"><select name="category" id="category" onchange="javascript: dropdownlist(this.options[this.selectedIndex].value);">
 <option value="">Select Category</option>
 <%C_Rs.MoveFirst 
   while not C_Rs.Eof %> 
 <option value="<%= C_Rs("CategoryID")%>"><%= C_Rs("CategoryName") %></option>
<% C_Rs.MoveNext
   Wend %> 
 </select></td>
 </tr>
 <tr>
 <td align="right" valign="middle">Sub Category :
 </td>
 <td align="left" valign="middle"><script type="text/javascript" language="JavaScript">
 document.write('<select name="subcategory"><option value="">Select Sub-Category</option></select>')
 </script>
 <noscript><select name="subcategory" id="subcategory" >
 <option value="">Select Sub-Category</option>
 </select>
 </noscript></td>
 </tr>
 </table>
 
</form> 
 
 
</body>
 </html>
 <%     DBOpen.close
   Set DBOpen = nothing %>

