<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<%

IF REQUEST("Operacao") = 1 THEN 'Cadastrar
	if request.ServerVariables("CONTENT_LENGTH") <> 0 then
	call abreConexao
	sql = "INSERT INTO Propriedade(Codigo, Nome, Endereco, UF, CodigoMunicipio) VALUES ('"&request.form("txtCodigo")&"', '"&request.Form("txtNome")&"', '"&request.Form("txtEndereco")&"', '"&request.Form("Estado")&"', '"&request.form("municipio")&"')"
	conn.execute(sql)
	call fechaConexao
	session("ID_usu") = request.Form("txtCodigo")
	response.Redirect("Propriedade.asp?Resp=1&Operacao=2&CodVisualizar="&request("txtCodigo")&"") 
	end if


ELSEIF REQUEST("Operacao") = 2 THEN 'Visualizar

call abreConexao
sql = "select Codigo, Nome, Endereco, UF, CodigoMunicipio from Propriedade WHERE Codigo = '"&request("CodVisualizar")&"'"
set rs = conn.execute(sql)

if not rs.eof then
Codigo=rs("Codigo")
Nome =rs("Nome")
Endereco=rs("Endereco")
Estado=rs("UF")
CodigoMunicipio=rs("CodigoMunicipio")
existe = 1
session("ID_usu") = Codigo

end if

ELSEIF REQUEST("Operacao") = 3 THEN 'ALTERAR
if request.ServerVariables("CONTENT_LENGTH") <> 0 then
call abreConexao
sql = "UPDATE Propriedade SET Nome = '"&request.Form("txtNome")&"', Endereco = '"&request.Form("txtEndereco")&"', UF = '"&request.Form("Estado")&"', CodigoMunicipio = '"&request.Form("municipio")&"' WHERE Codigo = '"&request.Form("txtCodigo")&"' "
conn.execute(sql)
session("ID_usu") = request.Form("txtCodigo")
call fechaConexao
response.Redirect("Propriedade.asp?Resp=2")

end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html >
<head>
<meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="description" content="Creative - Bootstrap 3 Responsive Admin Template">
  <meta name="author" content="GeeksLabs">
  <meta name="keyword" content="Creative, Dashboard, Admin, Template, Theme, Bootstrap, Responsive, Retina, Minimal">
  <link rel="shortcut icon" href="img/favicon.png">

  <title>Elements | Creative - Bootstrap 3 Responsive Admin Template</title>

  <!-- Bootstrap CSS -->
  <link href="css/bootstrap.min.css" rel="stylesheet">
  <!-- bootstrap theme -->
  <link href="css/bootstrap-theme.css" rel="stylesheet">
  <!--external css-->
  <!-- font icon -->
  <link href="css/elegant-icons-style.css" rel="stylesheet" />
  <link href="css/font-awesome.min.css" rel="stylesheet" />
  <!-- Custom styles -->
  <link href="css/style.css" rel="stylesheet">
  <link href="css/style-responsive.css" rel="stylesheet" />

<title>Propriedade</title>
<script src="javascript/Mascara.js"></script>
<script type="text/javascript">
function validar()
{
	if(document.frmPropriedade.txtCodigo.value == "")
	{
		alert("Obrigatorio digitar o Codigo!");
		document.frmPropriedade.txtCodigo.focus();
		return false;
	}
	
	
		if(document.frmPropriedade.txtNome.value == "")
	{
		alert("Obrigatorio digitar o Nome!");
		document.frmPropriedade.txtNome.focus();
		return false;
		}
	if(document.frmPropriedade.txtEndereco.value == "")
	{
		alert("Obrigatorio Digitar o Endereco!");
		document.frmPropriedade.txtEndereco.focus();
		return false;
		}	
	if(document.frmPropriedade.Estado.value == "")
	{
		alert("Obrigatorio Selecionar o Estado !")
		document.frmPropriedade.Estado.focus();
		return false;
		}	
	if(document.frmPropriedade.municipio.value == "")	
	{
		alert("Obrigatorio Selecionar o municipio !")
		document.frmPropriedade.municipio.focus();
		return false;		
	}
		return true;
}
function cadastrar()

{
	if(validar() == false)
	return false;
	
	document.frmPropriedade.Operacao.value = 1;
	document.frmPropriedade.action = "propriedade.asp";
	document.frmPropriedade.submit();
}
function Visualizar(Codigo)
{
	document.frmPropriedade.Operacao.value = 2;
	document.frmPropriedade.CodVisualizar.value = Codigo;
	document.frmPropriedade.action = "propriedade.asp";
	document.frmPropriedade.submit();

}
function verificar_cadastro()
{
	document.frmPropriedade.Operacao.value = 2;
	document.frmPropriedade.CodVisualizar.value = document.frmPropriedade.txtCodigo.value;
	document.frmPropriedade.action = "propriedade.asp"
	document.frmPropriedade.submit();
	}
function Alterar()
{
	
	if(validar() ==  false)
	return false;
	document.frmPropriedade.Operacao.value = 3;
	document.frmPropriedade.action = "propriedade.asp";
	document.frmPropriedade.submit(); 
	}
function Novo()
{
	window.location.href = "propriedade.asp";	
	}	

</script>
</head>
<center>
<%if request("Resp") = 1 or request("Resp") = 2 then
	if request("Resp") = 1 then
	mensagem = "Cadastro Realizado com Sucesso!"
	else
	mensagem = "Alteração Realizada com Sucesso!"
	end if
%>
<div class="messagem_informacao" align="center"><%=mensagem%></div>
<%end if%>
</center>

<body>
<section class="panel">
              <header class="panel-heading tab-bg-primary ">
                <ul class="nav nav-tabs">
                  <li class="">
                    <a data-toggle="tab" href="index.asp">Home</a>
                  </li>
                  <li class="active">
                    <a data-toggle="tab" href="Propriedade.asp">Propriedade</a>
                  </li>
                  <li class="">
                    <a data-toggle="tab" href="Cadastro.asp">Vinculacao</a>
                  </li>
                  <section class="panel">
                <ul class="nav nav-tabs pull-right">
                  <li class="">
                    <a data-toggle="tab" href="CadLogin.asp">Sair</a>
                  </li>
                
                </ul>
              </header>
<form name="frmPropriedade" onsubmit="validar();" id="frmPropriedade" method="post">
<input type="hidden" name="Operacao" id="Operacao" />
<input type="hidden" name="CodVisualizar" id="CodVisualizar" />
<p> 
Codigo:</br>
<input type="text" onkeypress="somente_numero(txtCodigo)" name="txtCodigo" id="txtCodigo" value="<%if existe = 1 then response.Write(Codigo) ELSE response.Write(Request("txtCodigo")) end if  %>" maxlength="6" onblur="somente_numero(txtCodigo);verificar_cadastro();"<%IF REQUEST ("Operacao") = 1 AND Existe = 1 THEN%> readonly<%END IF%>/>

</p>
<p>
Nome Propriedade:</br>
<input type="text" name="txtNome" id="txtNome" value="<%=Nome%>" />
</p>
<p>
Endereco Propriedade:</br>
<input type="text" name="txtEndereco" id="txtEndereco" value="<%=Endereco%>"  />
</p>
<p>
Estado:<select name="Estado" id="Estado"></br>
<option value="">Selecionar</option>
<option value="TO" <% if Estado = "TO" then %>selected<%end if%>>TO</option>
<option value="PA" <% if Estado = "PA" then%> selected<%end if%>>PA</option>
</select> 
</p>
<p>

Municipio:</br>
<%
call abreConexao
sql = "SELECT Codigo, Nome FROM Municipio order by Nome"
set rs = conn.execute(sql)
%>
<select name="municipio" id="municipio" value"<%=CodigoMunicipio%>">
<option value="">Selecionar</option>
<%do while not rs.eof%>
<option value="<%=rs("Codigo")%>"<%if rs ("Codigo") =  CodigoMunicipio then%>Selected<%end if%>><%=rs("Nome")%></option>


<%
rs.movenext
loop
call fechaConexao%>
</select>
</br>
<p>
<input type="submit"  name="btnCadastro"  value="<%if existe = 1 then%>Alterar<%else%>Cadastrar<%end if%>" onclick="return <%if existe = 1 then%>Alterar();<%else%> cadastrar();<%end if%>"/>
<%if existe = 1 then%>
<input type="button" onClick="return Novo();" value=" Novo Cadastro">
<a href = "Cadastro.asp"><input type="button" value="Ir Exploracao"></a> 
<%end if%>

</form>
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
  <th>Operacao</th>
  </tr>
  <%do while not rs.eof%>
  <tr>
  <td align="center"><%=rs("CodPropriedade")%></td>
  <td align="center"><%=rs("NomePropriedade")%></td>
  <td align="center"><%=rs("Endereco")%></td>
  <td align="center"><%=rs("UF")%></td>
  <td align="center"><%=rs("NomeMunicipio")%></td> 
  <td align="center"><a href="#" onclick="Visualizar(<%=rs("CodPropriedade")%>)"><img src="Imagem\editar.png" width="30"/></a></td>
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
