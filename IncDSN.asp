
<% 
Set DbOpen = Server.CreateObject("ADODB.Connection")
   connectString ="Provider=SQLOLEDB.1;uid=admela;pwd=lailabanu;database=AdMela;Data source=SUNOCO-PC\SQLEXPRESS;"
   'connectString ="Provider=SQLOLEDB.1;uid=mchowdh_laila01;pwd=lailabanu88;database=mchowdh_admela;Data source=IT-19;"
   DbOpen.Open ConnectString

   
   	   'On Error resume next
	   'I am adding following SQL statement on include file because this need for all pages for leftnav menus.
        Dim o_Sql, o_Rs

        o_Sql = "Select * From AM_Category where OnOff = 1 order by CategoryName"
	    Set o_Rs = Dbopen.Execute(o_Sql)
		



 %>	
   
   
 

