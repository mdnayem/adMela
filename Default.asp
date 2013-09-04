<!---#include file = "TopMenu.asp"--->
<% On Error Resume Next %>
		</td>
        </tr>
    
      <tr>
        <td colspan="5" valign="top"><table width="800" height="704px" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="179" height="708px" valign="top"><table width="190" height="704px" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="top"><table width="190"  height="704px" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td valign="top"><table width="190" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td colspan="3" height="285px" valign="top">
						  
						
								<!---#include file = "LeftNav.asp"--->	  
									  </td>
                                    </tr>
                                    <tr>
                                      <td><img src="Image/right-bot.jpg" width="194" height="25" /></td>
                                    </tr>
                                </table></td>
                              </tr>
                          </table></td>
                        </tr>
                    </table></td>
                  </tr>
                 
				 <tr>
				        <td>&nbsp;</td>
				 </tr>
			   
			   	  <!---insert left lower callout here--->

				   <tr>
                    <td height="235" colspan="3" valign="top"><table width="180" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="Image/right-top.jpg" width="184" height="22" /></td>
                        </tr>
						 <tr>
                          <td background="Image/right-bg2.jpg" height="100px" width="180px" valign="top">
						  <table width="180" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td colspan="2" align="left">
								<p class="online">Find your Match,<BR>visit ejuty.com today
						      <BR><A href="http://www.ejuty.com" target="_blank"><img src="image/ejutygreen.gif" width="150" height="40" /></a>
							 
								
								
								</td>
                              </tr>
                            
                          </table></td>
                        </tr>
                        <tr>
                          <td height="22"><img src="Image/right-bot.jpg" width="184" height="22" /></td>
                        </tr>
                    </table>
					</td>
                  </tr>
				  
				  
				  <!---callout ends here--->
			   
			   
			   
                  <tr>
                    <td valign="top"><table width="190" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td colspan="3" valign="top">
						  <table width="190" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td></td>
                              </tr>
                              <tr>
                                <td height="54"> </td>
                              </tr>
                          </table></td>
                        </tr>
                        <tr>
                          <td colspan="3">
						  <table width="190" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td></td>
                              </tr>
                              <tr>
                                <td valign="top">
								<table width="180" border="0" cellspacing="0" cellpadding="0" align="center">
                               
                                </table></td>
                              </tr>
                             
                          </table></td>
                        </tr>
                    </table></td>
                  </tr>
                  
               <!---left callout removed--->
                </table></td>
              </tr>
              
            </table></td>
            <td width="14" height="704px" ></td>
            <td width="389" valign="top">
			  <table width="388"  height="783" border="0" cellspacing="0" cellpadding="0" >
              <tr>
                <td valign="top"></td>
              </tr>
              <tr>
                <td  valign="top"><table width="389" height="745" border="0" align="center" cellpadding="0" cellspacing="0">
                 
                  <tr>
                    <td colspan="3"><table width="389" border="0" cellspacing="0" cellpadding="0" >
                      <tr>
                        <td>&nbsp;</td>
                        <td><img src="Image/bangla.bmp" /></td>
                        
                      </tr>
                      <tr>
                        <td colspan="3" height="10px" class="matter">Bangladeshis living abroad, please place your advertisement via our portal.
						No matter where you live; whether in 50 US states or in a different country, your ad can been seen by the Bangladeshis living
						through-out the world! Ad-mela is NY based and a subsidiary of ejuty.com. We appreciate your participation and patronage.
						</td>
                        </tr>
                      <tr>
					  <td width="8px"></td>
                        <td colspan="3"><hr color="#CECEFF" /></td>
						<td width="8px" class="matter"></td>
                        </tr>
                    </table></td>
                    </tr>
					<tr>
					<td width="8px" class="matter">Advertisement &raquo; Browse Sub-category ads</td>
					</tr>
                  <tr>
                    <td colspan="3" class="matter" >
					
				    <% o_Rs.MoveFirst
					While not o_Rs.Eof 
					 oDescription = Left(o_Rs("Description"),30)
					 oDescription = oDescription & "..."
					%>
							
					<BR /><B><Font color="#FF8040"><%= o_Rs("CategoryName") %></font></b> - 
					 <%= oDescription %>
					<BR /><A class="home2" href="Subcategory.asp?AdID=<%= o_Rs("CategoryID") %>&CategoryName=<%= o_Rs("CategoryName") %>&IMG=<%= o_Rs("IMG_Path") %>">more...</a>
					 <% 
					    o_Rs.MoveNext
					    Wend%>
					</td>
                    </tr>
                             
                  <tr>
                           <td height="100">&nbsp;</td>   
                  <tr>
                    <td colspan="3" valign="top">
					<table width="349" border="0" cellspacing="0" cellpadding="0" align="center">
                                           
                    </table></td>
                    </tr>
                  <tr>
                   
                    </tr>
                  <tr>
                    <td colspan="3" valign="top"><table width="389" border="0" cellspacing="0" cellpadding="0">
                      
                     
                    
                      <tr>
					  <td width="15"></td>
                        
						<td width="8"></td>
                      </tr>
                    </table></td>
                    </tr>
              
                 
                  <tr>
                    <td colspan="3"><table width="349" border="0" cellspacing="0" cellpadding="0" align="center">
                   
                    </table></td>
                    </tr>
                  
                  
                </table></td>
              </tr>
           
            </table>
			  </td>
            <td width="14" height="704px"></td>
            <td valign="top" height="704px">
			<table width="183" border="0" cellspacing="0" cellpadding="1">
                <tr>
                    <td colspan="3" valign="top"><table width="180" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="Image/right-top.jpg" width="184" height="15" /></td>
                        </tr>
                        <tr>
                          <td background="Image/right-bg2.jpg" height="169px" width="180px" valign="top"><table width="180" height="175" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td colspan="2" align="center" class="home2"><B />Sitar Evening!<hr color="#CECEFF" />
								<img src="Image/apu1.jpg" width="111" height="44" /></td>
                              </tr>
                              <tr>
                                <td width="10" align="center"></td>
                                <td width="148" class="home2">Do you want to arrange a Sitar evening? Contact Ustad Morshed Khan (Son of famous late Ustad Khurshid Khan) for a dazzling sitar evening in your State. 
								Please contact: <a href="mailto:mkhanmusic@juno.com" class="home">mkhanmusic@yahoo.com</a> or call at 917-923-3406 
								<BR /></td>
                              </tr>
                       
                          </table></td>
                        </tr>
                        <tr>
                          <td><img src="Image/right-bot.jpg" width="184" height="20" /></td>
                        </tr>
                    </table>
					</td>
                  </tr>
                  <tr>
                    <td colspan="3">&nbsp;</td>
                  </tr>
				 <!---right callout starts here---> 
                  <tr>
                    <td height="235" colspan="3" valign="top"><table width="180" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="Image/right-top.jpg" width="184" height="15" /></td>
                        </tr>
						 <tr>
                          <td background="Image/right-bg2.jpg" height="179px" width="180px" valign="top">
						  <table width="180" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td colspan="2" align="center" class="home2">Grab this spot for $15 a month!<hr color="#CECEFF" />
								<img src="Image/youradhere.jpg" width="120" height="100" /></td>
                              </tr>
                            
                              <tr>
                                <td width="10" align="center"></td>
								<td width="148" class="home2" align="center"><BR />Contact: <a href="mailto:ad@ad-mela.com" class="home">ad@ad-mela.com</a></td>
								
                              </tr>
                          </table></td>
                        </tr>
                        <tr>
                          <td height="22"><img src="Image/right-bot.jpg" width="184" height="20" /></td>
                        </tr>
                    </table>
					</td>
                  </tr>
				  <!---right callout ends here--->
				  
				
				  
				  <!---right callout starts here---> 
                  <tr>
                    <td height="235" colspan="3" valign="top"><table width="180" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="Image/right-top.jpg" width="184" height="15" /></td>
                        </tr>
						 <tr>
                          <td background="Image/right-bg2.jpg" height="179px" width="180px" valign="top">
						  <table width="180" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td colspan="2" align="center" class="home2">Grab this spot for $15 a month!<hr color="#CECEFF" />
								<img src="Image/youradhere.jpg" width="120" height="100" /></td>
                              </tr>
                            
                              <tr>
                                <td width="10" align="center"></td>
								<td width="148" class="home2" align="center"><BR />Contact: <a href="mailto:ad@ad-mela.com" class="home">ad@ad-mela.com</a></td>
								
                              </tr>
                          </table></td>
                        </tr>
                        <tr>
                          <td height="22"><img src="Image/right-bot.jpg" width="184" height="20" /></td>
                        </tr>
                    </table>
					</td>
                  </tr>
				  <!---right callout ends here--->
				  
				  
				  
                </table>
	</table>
<!---#include file = "Footer.asp"--->
</body>
</html>
