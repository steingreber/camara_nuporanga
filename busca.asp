           <table width="401" cellpadding="0" cellspacing="0">
		   <form action="busca_ok.asp">
               <tr>
                   <td align="center">
                       <span class="ms_href">BUSCA NO SITE</span>
                   </td>
                   <td align="right">
                       <input type="text" name="search" value="<% =Request.QueryString("search") %>" size="39" class="caixa">
                   </td>
                   <td align="center">
                       <input type="image" name="btvai" src="imagens/botao_buscar.gif" align="absmiddle" width="88" height="16" border="0">
                   </td>
               </tr>
			    <tr>
			        <td colspan="3" align="center" class="ms_href">
						Procurar por: 
						<input type="radio" name="mode" value="allwords" CHECKED>
						Todas as palavras
						<input type="radio" name="mode" value="anywords">
						Qualquer uma das palavras
					</td>
			    </tr>
			</form>
            </table>