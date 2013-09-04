<!---#include file = "TopMenu.asp"--->
		<% 'On Error Resume Next
		Dim objRs, objSql, objKeyword
		objKeyword = Request.Form ("Searchword")
		objKeyword = Replace(objKeyword,"'","""")
		objKeywork = Replace(objKeyword,"drop","car")
		
		'response.write objKeywork
		'response.end
		
				
		 %>
		</td>
        </tr>
     
      <tr>
        <td colspan="5" valign="top"><table width="800" border="0" cellspacing="0" cellpadding="0">
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
            <td valign="top">
			<table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="image/blog-bg-top.jpg" /></td>
                </tr>
              <tr>
                <td class="blogbg">
				<table width="530" border="0" cellspacing="0" cellpadding="0" align="center">
                  <tr>
                    <td colspan="3"><table width="530" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#FFFFFF"><img src="image/searchresult.jpg" height="75" width="75"  align="left"/></td>
                        <td width="0px" bgcolor="#FFFFFF"></td>
						<td class="centerheading" align="left" bgcolor="#FFFFFF">Search Results...</td>
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
    if objKeyword <> "" then
		                 objSql = "Select * From AM_ADS A INNER JOIN AM_SubCategory S ON A.SubCatID = S.SubCatID where A.Description LIKE ('%"& objKeyword &"%') and A.Active =1 order by ADID DESC"
		                 Set objRs = DbOpen.Execute(objSql)	 %>	
				 
				 
			
				  
					<%	
					
				If Not objRs.Eof then
					objRs.MoveFirst %>
					
					
			<TABLE cellpadding="5" bgcolor="#ffffff" cellspacing="3" width="100%" align="center">
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">AD ID</FONT></th>
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Description</FONT></th>
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Posted On</FONT></th>
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Photo</FONT></th>
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Details</FONT></th>
					
					
					
					<%  While not objRs.Eof %>
				  <TR>
				      
				      <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= objRs("ADID") %></font></td>
				      <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= objRs("ProductName") %></font></td>
					  <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= objRs("DateAdded") %></font></td>
					  <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><Img src="member/Uploads/<%= objRs("ImagePath") %>" border="0" height=20 width=30></font></td>
					  <TD bgcolor="#E8E8E8"><a href="AdDetails.asp?ADID=<%= objRs("ADID") %>&SubCatName=<%= SubCatName %>&IMG_Path=<%= objRs("ProductIMG") %>">Description</a></td>
					  
					  </td>
				   </tr>	  
					 <% objRs.MoveNext
					    Wend%>
					<!----ends here--->
					
					</table>
				
					<% Else %>
					   <font color="#FF0000" size="-1">"No Results have been found!<BR /> Index is up-to-date."</font>
					   
					  <%  End if 
					  Else %>
					  <font color="#FF0000" size="-1">"You must type few letters to conduct a search!" </font>
					<% End if  %>
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
