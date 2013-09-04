<!---#include file = "TopMenu.asp"--->
		
		<% Dim strLogout
		       strLogout = Request.QueryString("Action")
			   If strLogout = "Logout"  then 
			   Response.Cookies("AdMelaUser").Expires = Date() - 1
			   Response.Cookies("AdMelaUserName").Expires = Date() - 1
			   Response.Cookies("Hello").Expires = Date() - 1
			   Response.Redirect "Login.asp"
			   End if
			      
		 %>
		
		</td>
        </tr>
     
      <tr>
        <td colspan="5" valign="top"><table width="800" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="190px" valign="top"><table width="190" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
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
                        <td bgcolor="#FFFFFF"><img src="image/Login1.jpg" height="75" width="75"  align="left"/></td>
                        <td width="0px" bgcolor="#FFFFFF"></td>
						<td class="centerheading" align="left" bgcolor="#FFFFFF">Member's Login</td>
                      </tr>
                      <tr>
                   <td colspan="3" bgcolor="#FFFFFF"><hr color="#FF8040" /></td>
                      </tr>
                    </table></td>
                    </tr>
					
					
                  <tr bgcolor="#FFFFFF">
                    <td colspan="3" height="225px" class="matter" valign="top">
					<!---insert here--->
					<% 
					If Request.Cookies("AdMelaUser") = "" then %>
					<form name="myForm" method="post" action="ProcessLogin.asp">
					<table width="400" height="175" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td colspan="3" height="22px"></td>
                              </tr>
                              <tr>
                                <td colspan="1" class=""> </td>
                              </tr>
                              <tr>
							    <td colspan="3" class="matter" align="right">UserID: </td>
                                <td colspan="3"><input name="UserID" type="text" size="20px" value="<%= Request.cookies("remember") %>"/>
                                  &nbsp;</td>
                              </tr>
                              <tr>
							  <td colspan="3" class="matter" align="right" valign="middle">Pwd:</td>
                                <td colspan="3"><input name="Password" type="password" size="20px"/>
                                  &nbsp;</td>
                              </tr>
                        
                              <tr>
							    <td colspan="3" class="">&nbsp;</td>
                                <td colspan="3" height="18px"><input type="submit" name="login" value="Login"></td>
								<td colspan="3" height="18px"><BR /></td>
                              </tr>
                              <tr>
                                <td colspan="3" class=""><u></u></td>
                              </tr>
                              <tr>
							   <td colspan="3" class="">&nbsp;</td>
							   <td colspan="3" class=""><font face="Arial, Helvetica, sans-serif" color="#000000" size="2px">No account yet? </font>  <span class="matter">
							   <A href="Registration.asp">Create </a></span></td>
                                <td colspan="3" class="use"></td>
                              </tr>
							   <tr>
                          <td colspan="3">&nbsp;</td>
                        </tr>
                          </table>
						  </form>
                   <% Else
					Response.write "You are already logged in"
					End if %>
					<!---end here--->
					</td>
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
