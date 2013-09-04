



<!---#include file="incDSN.asp"--->

<% 

'On error resume next

		Dim UserID, Password
		    UserID = Request.Form("UserID")
			UserID = LCase(UserID) 
                                UserID = Replace (UserID,"'","""")
                        'Response.write UserID
						'response.end	
			Password = Request.Form("Password")
                               Password = Replace (Password,"'","""")
			
			
				Dim objSQL, objRs
					    objSQL = "Select * From AM_Users U INNER JOIN AM_Member M On U.UserName = M.EmailAddress WHERE (U.UserName='"& UserID &"' AND U.Password='"& Password &"')" 
						Set objRs = DbOpen.Execute(objSQL)
						strTimesLogin = objRs("TimesLoged")
						AddTimesLogin = strTimesLogin + 1
			
			
				Sub checkUser()
				            				    
							Response.Write "<script language='JavaScript'>alert('Login combination is incorrect! Please try again.');</script>"
							Response.Write "<script language='JavaScript'>history.back();</script>"
							Response.End
						
					End Sub
					
				If objRs.Eof then
				     'response.write "user could not found"
					 'response.end
							Call checkUser()
				 Else
						Application.Lock		
						        Response.cookies("remember")= UserID
								Response.Cookies("AdMelaUser") = objRs("EmailAddress")
				                Response.Cookies("AdMelaUserName") = objRs("FirstName")
								Response.Cookies("MemberID")= objRs("MemberID")
				                Response.Cookies("Hello") = "Hello: " 
								Session("password") = Password
								'***********************************************************
								
							   strRember = Request.Form("rememberpassword")  
							         If strRember = "yes" then
								       	 
										      
											  Response.cookies("remember").expires = dateadd("d",1,now)
											  Response.cookies("remember").path = "/" 
										 
									 
								       End if	
								   Dim upDateLogin, RsupDateLogin
								       upDateLogin = "Update AM_Users Set LastLogin='"& Now &"', TimesLoged ='"& AddTimesLogin &"' WHERE UserName = '"& UserID &"'" 
									   Set RsupDateLogin = DbOpen.Execute(upDateLogin)
		               Application.UnLock							   
								   
	              Response.Redirect "MyAccount.asp?LandOn=AdList"
  
  	      
	   End if		
								
								
								
								'*************************************************
								'response.write "user found"
							'Response.end	
			   	
						'If UserID <> "" And Password <> "" then
						
						
						
						'*************************************************************
						' Lets boot out all restricted users now
						' Lets check the DB juty_users table and restricted column. 
						'If the Restricted coulmn is populated then user is restricted, boot him out to Restricted.asp page
						
						'Dim ResRs, ResSQL
						    'ResSQL = "Select * From juty_users where UserID = '"& UserID &"'"
							'Set ResRs = DBopen.Execute (ResSQL)
							  'If ResRs("Restricted") = UserID then
							  'Response.Cookies("Restricted") = UserID
							  'Response.cookies("Restricted").path = "/Cookies/"
						      'Response.cookies("Restricted").expires = dateadd("d",1,now)
							  'Dim RemoteHOST
							  'RemoteHOST = Request.ServerVariables("REMOTE_ADDR")
							  'RemoteSQL = "Update juty_users Set Remote_IP='"& RemoteHOST &"' WHERE UserID ='"& UserID &"'"
							  'Set RemoteRs = DBopen.Execute(RemoteSQL)
							  'Response.Redirect "../Restricted.asp"
						    'End if
						
						
						
						'*****************************************************************
						
						
								'Lets pull the record of GOLD members and if they have passed 365 days, make them non-gold members
								'First look at PayCode column in Members table to see how many days users have gold membership for
								
								'Dim PayCodeSQL, PayCodeRs
								
								    'PayCodeSQL = "Select * From Juty_Members WHERE (User_ID = '"& UserID &"')"
									'Set PayCodeRs = Dbopen.Execute(PayCodeSQL)
									'o_PayCode = Trim (PayCodeRs("PayCode"))
									'o_HotSpot = PayCodeRs("DateHotSpotActivated")
									
									'Lets see if the user has HotSpot and is it time to expire...
									'If o_HotSpot <> "" then
									
									'sqlhotspot = "SELECT DATEDIFF(day,DateHotspotActivated,getdate()) AS Expire_HotSpot FROM Juty_Members WHERE (User_ID = '"& UserID &"')"
									'Set RsHotSpot = Dbopen.Execute(sqlhotspot)
									'ReadNumbOfDays = Trim (RsHotSpot("Expire_HotSpot"))
									'Response.write ReadNumbOfDays
									'If ReadNumbOfDays > 31 then
									'sqlExpire = "Update Juty_members set hotSpot = null, datehotspotactivated = null where (User_ID = '"& UserID &"')"
									'Set ExpireHotSpot = Dbopen.execute(sqlExpire)
									'End if
									'Response.end
									
									'End if 
									
									
								
						'Now Access Level check....
								
						'Dim Expire_Read_Access
						    'Expire_Read_Access = Trim (objRs("Access_Level"))
					'If Expire_Read_Access > 3 then		
						'Dim Expire_Sql, Expire_Rs
						
						    'Expire_Sql = "SELECT DATEDIFF(day,PayDateTime,getdate()) AS Expire_Look FROM Juty_Members WHERE (User_ID = '"& UserID &"')"
							'Set Expire_Rs = Dbopen.Execute(Expire_Sql)
							    'TimeDiffNumber = Trim (Expire_Rs("Expire_Look"))
								 
								'Diff_result = trim (o_PayCode - TimeDiffNumber)
                                                                'Response.Cookies("CountDown") = Diff_result
								
								
							'If Diff_result =< 0 then
							   
							
							 'strExpiredOn = "Expired on "& Now
							  
							   'o_expire_mem = "UPdate Juty_Members Set Access_Level = 3 where User_ID='"& UserID &"'"
							   'Set o_expire_Rs_mem = Dbopen.Execute(o_expire_mem)
							   'o_expire_user =  "UPdate Juty_Users Set Paid_NonPaid='"& strExpiredOn &"',Access_Level = 3 where UserID='"& UserID &"'"
							   'Set o_expire_Rs_mem = Dbopen.Execute(o_expire_user)
							   'Response.cookies("PermissionID")= 3
							   
							'End if 
							
					'End if	
					
					
					
					
					
					'**************************************************************************************************
						         'strCookies = Request.cookies("UserID")
								 'strPermission = Request.cookies("PermissionID")
								  
								 
								 'if strCookies <> "" then
								 'Response.Cookies(strCookies).Expires = Date() - 1
								 'End if
								 
								 'If strPermission <> "" then
								 'Response.Cookies(strPermission).Expires = Date() - 1
								 'End if
								 
						          'Response.cookies("Full_Name") = objRs("Full_Name")
						          'Response.cookies("UserID") = UserID
								  'Response.cookies("Password") = Password
								  'Response.cookies("MemberID") = objRs("ID")
								  
								  'If objRs("OldEmail") <> "" then
								  'Response.cookies("oldEmail") = objRs("OldEmail")
								  'End if
								  
								  'Session("PermissionID") = objRs("Access_Level")
								  'Response.cookies("PermissionID")= Trim (objRs("Access_Level"))
								  'Response.cookies("Email")= objRs("Email_Address")
								  'Response.cookies("Gender")= Trim (objRs("Gender"))
								  
							     
								  'Response.cookies("Full_Name").path = "/" 
							      'Response.cookies("UserID").path = "/" 
								  'Response.cookies("Password").path = "/" 
								  'Response.cookies("PermissionID").path = "/" 
								  'Response.cookies("Email").path = "/" 
								  'Response.cookies("Gender").path = "/" 
								  'Response.cookies("MemberID").path = "/"											


                  'iTrackRemote = Request.ServerVariables("REMOTE_ADDR")
                 
                  'iTrackSQL = "INSERT INTO TrackUser (UserID,TimeLogged,IP_Address) values('"& UserID &"','"& Now &"','"& iTrackRemote &"')"
                     'Set itrackRS = DBOpen.Execute(iTrackSQL)

   		
	

  
DbOpen.Close
Set DbOpen = Nothing	   		   




 %>

