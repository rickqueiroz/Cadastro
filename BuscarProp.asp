<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Buscar_Prop</title>
<script type="text/javascript">
function Buscar(Codigo)     
{

    opener.document.frmCadastro.txtCodigoProp.value = Codigo;
	close();
}

</script>
</head>

<body>


<%
call abreConexao
   sql = "SELECT Propriedade.Codigo AS CodPropriedade, Propriedade.Nome AS NomePropriedade, Propriedade.Endereco, Propriedade.UF, Municipio.Nome AS NomeMunicipio FROM Propriedade INNER JOIN Municipio ON Municipio.Codigo = Propriedade.CodigoMunicipio"
   set rs = conn.execute(sql)
%>

<table width="650" border="1" align="center">
  <%if rs.eof then%>
  <tr><td>Não Existe Nenhum Registro na base de Dados!</td></tr>
  <%else%>
  <tr>
  <th>Codigo</th>
  <th>Nome</th>
  <th>Endereco</th>
  <th>UF</th>
  <th>NomeMunicipio</th>
 
  </tr>
  <%do while not rs.eof%>
  <tr>
  <td align="center"><a href="#" onclick="Buscar(<%=rs("CodPropriedade")%>)"><%=rs("CodPropriedade")%></a></td>
  <td align="center"><%=rs("NomePropriedade")%></td>
  <td align="center"><%=rs("Endereco")%></td>
  <td align="center"><%=rs("UF")%></td>
  <td align="center"><%=rs("NomeMunicipio")%></td> 
 
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
