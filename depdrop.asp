<!---#include file = "TopMenu.asp"--->
<% 'on error resume next
GetParent = Request("combo0")
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
<form action="depdrop.asp">
<select name="combo0" id="combo_0" onChange="change(this);" style="width:200px;">
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
<BR><BR>
<input type="submit" name="submit" value="Submit"> 
 
</form>
<p align="center">This free script provided by<br />
<a href="http://javascriptkit.com">JavaScript
Kit</a></p>
<!---#include file = "Footer.asp"--->

</body>
</html>
