<!---#include file = "TopMenu.asp"--->




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
                        <td bgcolor="#FFFFFF"><img src="image/Register.jpg" height="75" width="75"  align="left"/></td>
                        <td width="0px" bgcolor="#FFFFFF"></td>
						<td class="centerheading" align="left" bgcolor="#FFFFFF">Registration
						<BR class="center"><font color="#FF0000" size="1"> [* = Required]</font>
						</td>
                      </tr>
                      <tr>
                   <td colspan="3" bgcolor="#FFFFFF"><hr color="#FF8040" /></td>
                      </tr>
                    </table></td>
                    </tr>
					
					
                  <tr bgcolor="#FFFFFF">
                    <td colspan="3" height="225px" class="matter" valign="top">
					<% 
					If Request.Cookies("AdMelaUser") = "" then 
					       If Request.querystring("_doPostBack") ="yes" then
					                      %>
					<!---insert here--->
                     
							<script type="text/javascript">
								alert("Please fill out all fields before clicking submit!");
							</script>
							<% Elseif Request.querystring("incPasswords") ="donotmatch" then %>	
							<script type="text/javascript">
								alert("Passwords do not match!");
							</script>
                            <% End if %>
					<Form id="myform" name="myform" method="post" action="ProcessRegistration.asp">
					<table width="400" height="275" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td colspan="3" height="22px"></td>
                              </tr>
                              <tr>
                                
                              </tr>
                              <tr>
							    <td colspan="3" class="matter" align="right">*First Name: &nbsp;</td>
                                <td colspan="3"><input id="Fname" name="Fname" value="<%= Request.queryString("Fname") %>" type="text" size="20px" maxlength="20"/>
                                  &nbsp;</td>
                              </tr>
							      <tr>
							    <td colspan="3" class="matter" align="right">*Last Name: &nbsp;</td>
                                <td colspan="3"><input  id="Lname" value="<%= Request.queryString("Lname") %>" name="Lname" type="text" size="20px" maxlength="20"/>
                                  &nbsp;</td>
                              </tr>
							  
							  
							  <tr>
							    <td colspan="3" class="matter" align="right">Zip Code: &nbsp;</td>
                                <td colspan="3"><input id="Zip" value="<%= Request.queryString("Zip") %>" name="Zip" type="text" size="5px" maxlength="5"/>
                                  &nbsp;</td>
                              </tr>
							  
							  <tr><% if request.queryString("_dups")="yes" then strDupsMSG  ="Email already exists!" End if %>
							    <td colspan="3" class="matter" align="right">*Email Address: &nbsp;</td>
                                <td colspan="3">
								<input  id="Email" value="<%= Request.queryString("Email") %>" name="Email" type="text" size="20px" maxlength="30"/>
								<font color="#FF0000"><%= strDupsMSG %></font>
                                  &nbsp;</td>
                              </tr>
							  
											                             
                              <tr>
							    <td colspan="3" class="matter" align="right">*Password:&nbsp;</td>
                                <td colspan="3"><input  id="Password" name="Password" type="password" size="20px" maxlength="20"/>
                                  &nbsp;</td>
                              </tr>
							    <tr>
							    <td colspan="3" class="matter" align="right">*Re-type Password:&nbsp;</td>
                                <td colspan="3"><input  id="Password1" name="Password1" type="password" size="20px" maxlength="20"/>
                                  &nbsp;</td>
                              </tr>
                              <tr>
                                <td colspan="3" height="14px"></td>
                              </tr>
                              <tr>
                                <td width="26" valign="top">
                                  </td>
                                <td colspan="2" class="matter" valign="top"></td>
                              </tr>
                              <tr>
                                <td colspan="3" height="18px">&nbsp;</td>
								<td colspan="3" height="18px" align="center"><input type="submit" name="register" value="Register!"></td>
                              </tr>
                              <tr>
                    </table>
					</form>
					
					<% Else
					Response.write "You already have registedred and logged in"
					End if %>
					
					
					<!---end here--->
					</td>
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
