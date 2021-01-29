<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<%
'nao é aqui,essa função!
session("ID_usu")
dim existeCadastro
function Verificar_Exploracao()
call abreConexao
 SQL2 = "SELECT COUNT (*) as qt FROM ExploracaoPecuaria  WHERE CPFCNPJProdutor = '"&session("CPF_Usu")&"' AND CodigoPropriedade = '"&session("Id_Usu")&"'"
 set rs = conn.Execute(SQL2)
		if rs("qt") <> 0 then
			existeCadastro = true
		else
			existeCadastro = false
		end if
		rs.close
		set rs = nothing
		Verificar_Exploracao = existeCadastro 
call fechaConexao
end function
IF REQUEST ("Operacao") = 1 THEN ' Cadastrar
    if Verificar_Exploracao() = false then
    call abreConexao
    sql = "INSERT INTO ExploracaoPecuaria (CPFCNPJProdutor, CodigoPropriedade, TipoExploracao) VALUES ('"&session("CPF_usu")&"', '"&session("ID_usu")&"', '"&request.Form("TipoExploracao")&"' )"
conn.execute (sql)
call fechaConexao
END IF
response.Redirect("Exploracao.asp")
end if


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Exploracao</title>
<script type="text/javascript">
function Validar()
{
if(document.frmExploracao.TipoExploracao.value == "")
{
	alert("Obrigatorio Selecionar o TipoExploracao");
	document.frmExploracao.TipoExploracao.focus();
	return false;
	}
	return true;
}
function cadastrar()
{
	if(Validar() == false)
	return false;
	
	document.frmExploracao.Operacao.value = 1;
	document.frmExploracao.action = "Exploracao.asp";
	document.frmExploracao.submit();
}	

	

</script>
</head>
<body>
<form name="frmExploracao" onsubmit="Validar();" method="post">
<input type="hidden" name="Operacao" id="Operacao" />

<p>
Tipo Exploracao:</br>
<%
call abreConexao
sql = "SELECT Codigo, Descricao FROM TipoExploracao order by Codigo"

set rs = conn.execute(sql)

%>
<select name="TipoExploracao" id="TipoExploracao">
<option value="">Selecionar</option>
<%do while not rs.eof%>
<option value="<%=rs("Codigo")%>"><%=rs("Descricao")%></option>
<%
rs.movenext
loop
call fechaConexao%>
</select>
</p>
<input type="submit" name="btnCadastrar" value="Vincular" onclick=" cadastrar();"/>
</p>
</form>
<%
   call abreConexao
   sql = "SELECT Produtor.CPF, Produtor.Nome, Propriedade.Nome AS NomePropriedade, Propriedade.UF,Propriedade.Codigo,  TipoExploracao.Descricao FROM ExploracaoPecuaria INNER JOIN Produtor ON Produtor.CPF = ExploracaoPecuaria.CPFCNPJProdutor INNER JOIN Propriedade ON Propriedade.Codigo = ExploracaoPecuaria.CodigoPropriedade INNER JOIN TipoExploracao ON TipoExploracao.Codigo = ExploracaoPecuaria.TipoExploracao ORDER BY Nome; "
 
    
   set rs = conn.execute(sql)
%>

<table width="650" border="1" align="center">
  <%if rs.eof then%>
  <tr><td>Não Existe Nenhum Registro na base de Dados!</td></tr>
  <%else%>
  <tr>
  <th>CPFCNPJProdutor</th>
  <th>Nome</th>
  <th>NomePropriedade</th>
  <th>UF</th>
  <th>Codigo</th>
  <th>TipoExploracao</th>
  </tr>
  <%do while not rs.eof%>
  <tr>
  <td align="center"><%=rs("CPF")%></td>
  <td align="center"><%=rs("Nome")%></td>
  <td align="center"><%=rs("NomePropriedade")%></td>
  <td align="center"><%=rs("UF")%></td>
  <td align="center"><%=rs("Codigo")%></td>
  <td align="center"><%=rs("Descricao")%></td>
  </tr>
  <%
     rs.movenext
	 loop
  %>
  <%end if%>
</table>

<%call fechaConexao%>

</body>
</html>
