<%Response.ContentType = "text/html; charset=iso-8859-1"%>
<!--#include file="_cnx.asp"-->
<!--#include file="config.asp"-->
                        <table border="0" width="99%" align="center" cellspacing="0">
							<tr>
							<td height="20" colspan="2" bgcolor="<%= cCorTopoBase %>" class="tittext"><font color="<%= ca05corLink %>">TV CÂMARA</font></td>
							</tr>
                        	<tr>
                        		<td class="texto">
								
									<SCRIPT LANGUAGE = "JScript"  FOR = Player EVENT = StatusChange()>
									
									    // Display status. This requires Windows Media Player 9 Series or later.
									    divStatus.innerHTML = "Status: " + Player.status;
									
									</SCRIPT>
									<span class="navtext">Assista a programação de segunda a sexta.</span><br><br>
									<table cellpadding="0" cellspacing="0" border="0" align="center">
                                    	<tr>
                                    		<td align="center" class="texto">
												
												<div id="fora_do_ar" style="visibility:hidden;display:none">
										            <img src="imagens/aviso_tv.jpg" alt="" width="240" height="180" border="0">
										        </div>
										        <div id="ao_vivo" style="visibility:visible;display:inline">
										            <OBJECT id="Player" classid="CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6" width="280" height="220">
										            <PARAM NAME="autoStart" VALUE="true">
										            <param name="url" value="<%= tvcamara %>">
										            </OBJECT>
										        </div>
												<br>
												<!-- <input type="button" name="btnPlayVideo" value="Assistir" class="ud_caixa" onClick="PlayVideo()">
									            <input type="button" name="btnStop" value="Parar" class="ud_caixa" onClick="StopMe()"> -->
									            <font face="Verdana" size="1">
									            	<b><DIV id=divStatus size=50><span class="navtext">Status:</span> </DIV></b>
									            </font>
												<p align="center">
											    <a href="<%= tvcamara %>" class="texto"><span class="navtext">Clique para assistir a programação da TV Câmara AO VIVO</span></a>&nbsp;<img src="imagens/ic_windowsmedia.gif" alt="" border="0" align="absmiddle"><br>
											</td>
                                    	</tr>
                                    </table>
