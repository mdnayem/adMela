<!---#include file = "TopMenu.asp"--->
	
	
	
	<% On Error Resume Next
	   strIMG = Request.Querystring("IMG_Path")
	   strAdID = Trim(Request.Querystring("ADID"))
	   
	   
	  Dim objSql, objRs, AdCatID

     SubCatName = Request.Querystring("SubCatName")
  
    objSql = "Select * From AM_ADS where ADID='"& strAdID &"' And Active = 1"
	Set objRs = Dbopen.Execute(objSql)
	strProductTitle = objRs("ProductName")
	
	IMG_Path = objRs("ImagePath")
	  if IMG_Path <>"" then
	     IMG_Path = "nophoto.jpg"
		 else
		 IMG_Path = IMG_Path
	  End if	 
	
	o_dateDiff = trim (Now - (objRs("DateAdded") ))
	'response.write o_datediff - this is working fine
	expires_On = (objRs("DateAdded") + 30)
                  
				  If Request.queryString("saveAd") = "yes" And Request.Cookies("AdMelaUser") <> "" then
				  
				  AddedByEmail = Request.Cookies("AdMelaUser")
				  strSubCatName= Request.queryString("SubCatName")
				  strIMGPath = Request.queryString("IMG_Path")
				  
				  objAdSaveSQL = "Insert into am_SavedAds(ADID,AddedByEmail,AdTitle,DateAdded) values ('"& strAdID &"','"& AddedByEmail &"','"& strProductTitle &"','"& Now &"')"
				  Set objAdSaveRs = DbOpen.Execute(objAdSaveSQL)
				  Response.Redirect "AdDetails.asp?saveAd=saved&ADID="& strAdID &"&SubCatName="& strSubCatName &"&IMG_Path="& strIMGPath
				  
                  End if

 %>
	  	
		</td>
        </tr>
     
      <tr>
        <td colspan="5" valign="top">
		<table width="800" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="190px" valign="top"><table width="190" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="top">
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
                <td class="blogbg">
				<table width="530" border="0" cellspacing="0" cellpadding="0" align="center">
                  <tr>
                    <td colspan="3">
					<table width="530" border="0" cellspacing="0" cellpadding="0">
                      <tr>
					    <td bgcolor="#FFFFFF">&nbsp;</td>
                        <td bgcolor="#FFFFFF"><img src="image/<%= strIMG %>" height="55" width="55" />
						</td>
                        <td class="centerheading" align="left" bgcolor="#FFFFFF"><%= SubCatName %> &raquo; Ad ID # <Font color="#FF8040" face="Impact"><%= strAdID %></font>
						<BR>
						</td>
                      </tr>
                      <tr>
                   <td colspan="3" bgcolor="#FFFFFF"><hr color="#FF8040" /></td>
                      </tr>
                    </table></td>
                    </tr>
					
					
                  <tr bgcolor="#FFFFFF">
                    <td colspan="3" height="225px" class="matter" valign="top">
					<!---insert here--->
					
					
			<TABLE cellSpacing=0 cellPadding=5 width=500 border=0>
                <TBODY> 
                <TR> 
                  <TD width="25%"><A href="AdList.asp?SubCatID=<%= objRs("SubCatID") %>&ADID=&IMG_Path="><img src="image/bullet.png" border="0"/></a> </TD>
				  <TD><% 
				  If Request.Cookies("AdMelaUser") <> "" and Request.QueryString("saveAd")<> "saved" then
				         strAddedByEmail = Trim(Request.Cookies("AdMelaUser"))
						 'Lets check if the Ad has been saved before or not
						 chkAdSQL = "Select * From AM_SavedAds where ADID ='"& strAdID &"' AND AddedByEmail='"& strAddedByEmail &"'"
						 set chkRs = DbOpen.Execute(chkAdSQL)
						 strWrite = "You have saved this listing"
				       If chkRs.Eof then
					  					     
				  %>
				  <A onclick="javascript:return confirm('Are you sure you want to save this listing?');" href="AdDetails.asp?saveAd=yes&<%= Request.queryString %>"><IMG height="60" width="70" SRC="image/adsave.jpg" border=0></a>
				     <% Else
					 strWrite = "You have saved this listing"
					 End if
				Else
				Response.Write ="" 
				
				End if %>
				<%= strWrite %>
				  </td>
				  <TD>
				  <BR>
				  
				   <tr bgcolor="#FFFFFF">
				       <TD width="25%">
					   <BR>
					   <IMG width="100" height="90" SRC="Uploads/<%= IMG_Path %>" border=0>
					   <BR>
					   <font color="#804040" size="-2">Date Added:<%= objRs("DateAdded") %></font>
					   <!---<br><font color="#804040" size="1">Expires on:<%= expires_On %></font>--->
					   <br><font color="#804040" size="1">Ad Number:<%= objRs("ADID") %></font>
					   </td>
					   
					   <TD>
					   <TABLE>
					   
					   
					   <TR>
					        <TD><font color="#008040" size="2">Summary:</font></td>
							<TD><font color="#800000" size="2"><%= objRs("ProductName") %></font></td>
					   </tr>
					   <TR>
					        <td><font color="#008040" size="2">Description:</font></td>
							<td><font color="#800000" size="2"><%= objRs("Description") %></font></td>
					        
					   </TR>
					   
					   <TR>
					      <TD><font color="#008040" size="2">Contact:</font></td>
					      <TD><font color="#800000" size="2"><%= o_Rs("ContactName") %></font></td>
					   </tr>
					   <TR>
					      <TD><font color="#008040" size="2">Location:</font></td>
					      <TD><font color="#800000" size="2"><%= objRs("Location") %></font></td>
					   </tr>
					   <TR>
					      <TD><font color="#008040" size="2">Telephone:</font></td>
					      <TD><font color="#800000" size="2"><%= objRs("Telephone") %></font></td>
					   </tr>
					   <TR>
					      <TD><font color="#008040" size="2">Email:</font></td>
					      <TD><A href="mailto:<%= objRs("Email_Address") %>"><%= objRs("Email_Address") %></A></td>
					   </tr>
					   <TR>
					      <TD><font color="#008040" size="2">Website:</font></td>
					      <TD><A href="<%= objRs("URL") %>" Target="_blank"><%= objRs("URL") %></a></td>
					   </tr>
					   
					   </table>
					  
										  
					   </td>
				  </tr>
				  <TR>
				  <TD>&nbsp;</td>
				  <td></td>
				  </tr>
				  
				  
				  </TD>
                </TR>
				<TR>
				
				          <TD>&nbsp;</td>
				</tr>
			
                </TBODY>
              </TABLE>
					
					
					
					
					
					
					
					
					
					
					<!---end here--->
					</td>
                    </tr>
                  <tr bgcolor="#FFFFFF">
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr bgcolor="#FFFFFF">
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr bgcolor="#FFFFFF">
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
