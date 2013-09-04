<!---#include file = "TopMenu.asp"--->
<% On Error resume next
Dim AllProdSql, AllProdRs, SubCatID, strIMG

   SubCatID = Request.Querystring("SubCatID")
   
   
   PID = Request.Querystring("PID")
   strQuery = "AdList.asp?"& Request.QueryString
   CategoryName = Request.Querystring("CategoryName")
   productcategory = Request.Querystring("productcategory")
    AllProdSql = "Select * From AM_ADS A INNER JOIN AM_Subcategory S ON A.SubCatID = S.SubCatID where A.SubCatID='"& SubCatID &"' And A.Active = 1 order by A.ADID DESC"
	Set AllProdRs = Dbopen.Execute(AllProdSql)
	SubCatName = AllProdRs("SubCategoryName")
	
	strIMG = AllProdRs("ProductIMG")



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
					<table width="530" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td bgcolor="#FFFFFF">&nbsp;</td>
                        <td bgcolor="#FFFFFF"><img src="image/<%= strIMG %>" height="55" width="55" /></td>
						<td class="centerheading" bgcolor="#FFFFFF">
						
						<br>Classified List for " <Font color="#FF8040" face="Impact"><%= SubCatName %></font> " category 
						<BR>
						<% If AllProdRs("CategoryID") = "" then
						   'Response.write "category found"
						   'response.end
						 %>
						 <BR><A href="#" onclick="history.back();"><img src="image/bullet.png" border="0"/></a>
						 <% Elseif AllProdRs("CategoryID") <> "" then %>
						<BR><A href="SubCategory.asp?ADID=<%= AllProdRs("CategoryID") %>&SubCatName=&IMG_Path="><img src="image/bullet.png" border="0"/></a>
						
						
						<% End if %>
						</td>
                      </tr>
                      <tr>
                   <td colspan="3" bgcolor="#FFFFFF"><hr color="#FF8040" /></td>
                      </tr>
                    </table></td>
                    </tr>
                  <tr>
                    <td colspan="3" height="225px" class="matter" valign="top" bgcolor="#FFFFFF">
					
					<!---pulling products start here--->
					
				  
				   <TABLE cellpadding="5" bgcolor="#ffffff" cellspacing="3" width="100%" align="center">
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">AD ID</FONT></th>
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Description</FONT></th>
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Posted On</FONT></th>
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Photo</FONT></th>
				   <TH bgcolor="#69ACB8"><Font color="#FFFFFF">Details</FONT></th>
				  
					<%AllProdRs.MoveFirst 
					  While not AllProdRs.Eof %>
				  <TR>
				      
				      <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= AllProdRs("ADID") %></font></td>
				      <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= AllProdRs("ProductName") %></font></td>
					  <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><%= AllProdRs("DateAdded") %></font></td>
					  <TD bgcolor="#E8E8E8"><Font color="#004000" size="-1"><Img src="member/Uploads/<%= AllProdRs("ImagePath") %>" border="0" height=20 width=30></font></td>
					  <TD bgcolor="#E8E8E8"><a href="AdDetails.asp?ADID=<%= AllProdRs("ADID") %>&SubCatName=<%= SubCatName %>&IMG_Path=<%= AllProdRs("ProductIMG") %>">Description</a></td>
					  
					  </td>
				   </tr>	  
					 <% AllProdRs.MoveNext
					    Wend%>
					<!----ends here--->
					
					</table>
					
					
					
					
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
