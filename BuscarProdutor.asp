<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Buscar_Produtor</title>
<script type="text/javascript">
function Buscar(Codigo)     
{

    opener.document.frmCadastro.txtCNPJProdutor.value = Codigo;
	close();
}
</script>

</head>

<body>
<%
   call abreConexao
   sql = "SELECT Produtor.CPF, Produtor.CodigoMunicipio, Produtor.Nome, Produtor.Endereco, Produtor.Telefone,Produtor.Status, Municipio.Nome AS NomeMunicipio, Municipio.UF FROM Produtor INNER JOIN Municipio ON Municipio.Codigo = Produtor.CodigoMunicipio"
   set rs = conn.execute(sql)
%>

<table width="650" border="1" align="center">
  <%if rs.eof then%>
  <tr><td>Não Existe Nenhum Registro na base de Dados!</td></tr>
  <%else%>
  <tr>
  <th>CPF</th>
  <th>Nome Completo</th>
  <th>Endereco</th>
  <th>Telefone</th>
 
  </tr>
  <%do while not rs.eof%>
  <tr>
  <td align="center"><a href="#" onclick="Buscar(<%=rs("CPF")%>)"><%=rs("CPF")%></a></td>
 
  <td align="center"><%=rs("Nome")%></td>
  <td align="center"><%=rs("Endereco")%></td>
  <td align="center"><%=rs("Telefone")%></td>

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
