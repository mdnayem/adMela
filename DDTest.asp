<!---#include file = "TopMenu.asp"--->
<% 
GetParent = Request("combo0")
'response.write GetParent
GetChild = Request("combo1")
response.write GetChild

    'Change the names here before turning them ON
     Dim Cat_Sql, Cat_Rs, SubCat_Sql, SubCat_Rs

    Cat_Sql = "Select * From AM_Category where OnOff = 1 order by CategoryID ASC"
	Set Cat_Rs = Dbopen.Execute(Cat_Sql)
	
	'response.write Cat_Rs("CategoryName")
	'response.end
	
	SubCat_Sql = "Select * From AM_subcategory order by CategoryID ASC"
	Set SubCat_Rs = Dbopen.Execute(SubCat_Sql)
	writeMessage = SubCat_Rs("SubCatID")
	'response.write SubCat_Rs("SubCatID")
	'response.end



 %>


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

data_100 = new Option("100", "$");
data_200 = new Option("200", "$$");
data_300 = new Option("300", "$$");
// second combo box
data_1_1 = new Option("11", "11");
data_1_2 = new Option("12", "12");
data_3_1 = new Option("31", "31");
data_4_1 = new Option("21", "21");
data_2_2 = new Option("22", "22");
data_2_3 = new Option("23", "23");
data_2_4 = new Option("24", "24");
data_2_5 = new Option("25", "25");



 
// first combo box
 <% While not Cat_Rs.Eof %>
data_<%= Cat_Rs("CategoryID")%> = new Option("<%= Cat_Rs("CategoryID")%>", "$");
<%  Cat_Rs.MoveNext
    wend %>
// second combo box
<% While not SubCat_Rs.Eof 
   'objS = "Select * From AM_Category where OnOff = 1 order by OrderBy ASC"
   'Set objR = Dbopen.Execute(objS)
%>
data_<%= SubCat_Rs("CategoryID")%>_<%= SubCat_Rs("OrderBy")%> = new Option("<%= SubCat_Rs("SubCatID")%>", "-");
<% SubCat_Rs.MoveNext
	      Wend %>
 
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
<form method="post" action="DDTest.asp">
<select name="combo0" id="combo_0" onChange="change(this);" style="width:200px;">
<option value="value1">-select-</option>
<%Cat_Rs.MoveFirst 
  While not Cat_Rs.Eof %>
<option value="<%= Cat_Rs("CategoryID")%>"><%= Cat_Rs("CategoryID")%></option>
<%  Cat_Rs.MoveNext
    wend %>

</select>
<BR><BR>
<select name="combo1" id="combo_1" onChange="change(this)" style="width:200px;">
<option value="value1"> Select one </option>
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