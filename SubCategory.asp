<% Languange = VBScript %>
<!---#include file = "TopMenu.asp"--->
<% 'On Error Resume Next
       Dim strCategory
       strCategory = Request.QueryString("CategoryName")
	   strImage = Request.QueryString("IMG")
	   strCatID = Request.QueryString("AdID")
	   
	   Dim objRs, objSql, objSubCatID
	   objSql = "Select * From AM_SubCategory where CategoryID = '"& strCatID &"' AND OnOff = 1 order by SubCategoryName"
	   Set objRs = DBOpen.Execute(objSql)
	   strIMG = objRs("ProductIMG")
	    
	  

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
            <td valign="top"><table width="600" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="image/blog-bg-top.jpg" /></td>
                </tr>
              <tr>
                <td class="blogbg">
				<table width="530" border="0" cellspacing="0" cellpadding="0" align="center">
                  <tr>
                    <td colspan="3">
					<table width="530" border="0" cellspacing="0" cellpadding="0" cellpadding="1">
                   
                      <tr>
                   <td colspan="3" bgcolor="#FFFFFF"></td>
                      </tr>
                    </table></td>
                    </tr>
                  <tr>
                    <td class="last" colspan="3" height="20px" bgcolor="#FFFFFF" valign="baseline">
					<Img src="Image/<%= strIMG %>" border="0" height="55" width="55">
					<%= strCategory %> &raquo; Sub Categories
					
					</td>
                    </tr>
					<TR>
					  <td colspan="3" bgcolor="#FFFFFF"><hr color="#FF8040" /></td>
					</tr>
					<% 'objRs.MoveFirst
					   While not objRs.Eof 
					       
				           'Count how many ads been posted....
						   CountSql = "Select Count (*) as myCount From AM_ADS where Active=1 and SubCatID ="& objRs("SubCatID")
						   Set CountRs = Dbopen.Execute(CountSql)
						   'If objRs("myCount")> 0 then
						   'aHref = "<A HREF='AdList.asp?SubCatID="& objRs("SubCatID")&"'>"
						   'aHrefClose = "</a>"
						   'End if 
						  %>
					
					
					
                    <tr>
                      
				      <td class="matter" colspan="3" bgcolor="#FFFFFF" valign="top" width="100"><img src="image/left-arrow1.gif" border="0">
					   <A href="AdList.asp?SubCatID=<%= objRs("SubCatID") %>"><%= objRs("SubCategoryName") %></a> &nbsp;(<%= CountRs("MyCount") %><font class="home1" color="#69ACB8" size="2"> records found!)</Font>
					  
					  </td>
					  <TD bgcolor="#FFFFFF"></td>
				   </tr>	  
					 <% objRs.MoveNext
					    Wend%>
                  <tr>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
				   <tr>
                    <td height="160">&nbsp;</td>
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
<%   If Error.number > 0 then
		 Response.Redirect "Default.asp"
		 End if %>