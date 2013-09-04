<%@ LANGUAGE = VBScript %>
<!---#include file="incDSN.asp"--->
<!---#include file="Email.asp"--->
<% 'On Error Resume Next
Dim objFName, objLName, objEmail, objPassword, objZip
Dim objSession, objRs, objSql 
 
		objFName = Request.Form("Fname")
		objLName = Request.Form("Lname")
		objEmail = Trim(Request.Form("Email"))
		objPassword = Trim(Request.Form("password"))
		objPassword1 = Trim(Request.Form("password1"))
		objZip = Request.Form("Zip")
				     If objZip = "" then
				      objZip = 00000
					  Else
					  objZip = objZip
				   End if
		
		
	   'If objPassword <> "" then
		   If objPassword <> objPassword1 then
		    Response.Redirect "Registration.asp?incPasswords=donotmatch&Fname="& objFname &"&Lname="& objLname &"&Zip="& objZip &"&Email="&objEmail
		   End if
		'End if 
		
		objSiteRelID = Session.SessionID 
		objLoginCount = 1		   
				   	  
		If objFName <>"" and objLName <>"" and objEmail <> "" and objEmail <> "" and objPassword <> "" then
		   Dim chkUserName, chkUserRs
		   'response.write "email found"
		   'response.end
		'Lests check if the email address is a dups or not
		chkUserName = "Select * From AM_Member Where EmailAddress ='"& objEmail &"'"
		Set chkUserRs = DbOpen.Execute(chkUserName)
		
		   If chkUserRs.Eof then
		      
			    objSql = "Insert into AM_Member (FirstName,LastName,ZipCode,EmailAddress,SiteRelID,MemberSince) Values ('"& objFName &"','"& objLName &"','"& objZip &"','"& objEmail &"','"& objSiteRelID &"','"& Now &"')"
				userSql = "Insert into AM_Users(UserName,Password,MemberRelID,TimesLoged) values('"& objEmail &"','"& objPassword &"','"& objSiteRelID &"','"& objLoginCount &"')"
				Set objRs = DbOpen.Execute(objSql)
				Set objUserSql = DbOpen.Execute(userSql)
				Response.Cookies("AdMelaUser") = objEmail
				Response.Cookies("AdMelaUserName") = objFName
				Response.Cookies("Hello") = "Hello"
				Response.Redirect "MyAccount.asp?LandOn=AdList"
			Else
			Response.Redirect "Registration.asp?_dups=yes&Fname="& objFname &"&Lname="& objLname &"&Zip="& objZip&"&Email="&objEmail	
		   End if 
		 
		 Else
		 Response.Redirect "Registration.asp?_doPostBack=yes&Fname="& objFname &"&Lname="& objLname &"&Zip="& objZip&"&Email="&objEmail	
		 
		 End if 	
		
		
		    DbOpen.Close
		Set DbOpen = Nothing
		
		
		
		%>