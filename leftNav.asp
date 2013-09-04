  <table width="190" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td colspan="3" valign="top">
								<table width="190" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td><img src="Image/top.jpg" width="194" height="18" /></td>
                                    </tr>
                                    <tr>
                                      <td align="center" height="54" class="topbg"><img src="image/classifieds_logo.jpg" border="0" width="150"/> </td>
                                    </tr>
                                </table></td>
                              </tr>
                              <tr>
                                <td colspan="3" valign="top">
								<table width="190" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td></td>
                                    </tr>
                                    <tr>
                                      <td height="190px" width="180px" valign="top" class="centerbg">
									  
									  <table background="Image/right-bg3.jpg" width="185" border="0" cellspacing="0" cellpadding="0" align="center">
									    <% While not o_Rs.Eof %>
                                          <tr>
                                            <td width="26">&nbsp;</td>
                                            <td width="26" height="30px"><Img src="Image/<%= o_Rs("IMG_Path") %>" border="0" height="25" width="25"></td>
                                            <td class="home"><a href="Subcategory.asp?AdID=<%= o_Rs("CategoryID") %>&CategoryName=<%= o_Rs("CategoryName") %>&IMG=<%= o_Rs("IMG_Path") %>" class="home" title="<%= o_Rs("CategoryName") %>"><%= o_Rs("CategoryName") %></a>
											</td>
                                          <tr>
										  <% o_Rs.MoveNext
					                         Wend%>
										  
                                          										  
                                      </table>