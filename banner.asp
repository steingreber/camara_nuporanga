<% 
randomize
set result = objConect.Execute("SELECT * FROM BANNERS Where a05exibir=1 AND a07tamanho=1")
nums = ""
while not result.eof
 nums = nums & result("a00id") & "|"
 result.movenext
wend

arrTemp = Split(nums,"|")
id_sort = Round((ubound(arrTemp)-1) * Rnd)

set result = objConect.Execute("SELECT * FROM BANNERS Where a00id = " & arrTemp(id_sort))
sImg1 = result(3)
idesc = result(1)
iLink = result(2)
iTipG = result(8)
iNovo = result(6)
%>

<% If iTipG = "0" Then %>
                    <a href="<%= iLink %>" <% If iNovo = "0" Then %>target="_blank"<% End If %>><img src="config/banners/<%= sImg1 %>" alt="<%= idesc %>" width="468" height="60" border="0"></a>
<% Else %>
                    <script>ConteudoFlash('config/banners/<%= sImg1 %>','60','468');</script>
<% End If %>