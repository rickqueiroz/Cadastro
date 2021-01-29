<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<%
session("ID_usu")
dim existeCadastro
function Verificar_Exploracao()
call abreConexao
 SQL2 = "SELECT COUNT (*) as qt FROM ExploracaoPecuaria  WHERE CPFCNPJProdutor = '"&request.form("txtCNPJProdutor")&"' AND CodigoPropriedade = '"&request.Form("txtCodigoProp")&"'"
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
	sql = "INSERT INTO ExploracaoPecuaria(CPFCNPJProdutor,CodigoPropriedade, TipoExploracao) VALUES ('"&request.form("txtCNPJProdutor")&"', '"&request.Form("txtCodigoProp")&"', '"&request.Form("TipoExploracao")&"')"
	conn.execute(sql)
	call fechaConexao 
	END IF
	
	response.Redirect("Cadastro.asp?Resp=1")
ELSEIF REQUEST("Operacao") = 2 THEN 'Visualizar
	call abreConexao
 	sql = "select CPFCNPJProdutor, CodigoPropriedade, TipoExploracao  from ExploracaoPecuaria WHERE CPFCNPJProdutor = '"&request.form("CPFVisualizar")&"' AND CodigoPropriedade = '"&request.Form("CodPropriedade")&"'"
 
	  	set rs = conn.execute(sql)
	if not rs.eof then
	CPFCNPJProdutor=rs("CPFCNPJProdutor")
	Codigo =rs("CodigoPropriedade")
	TipoExploracao= clng(rs("TipoExploracao"))
	existe = 1
	end if
	call fechaConexao
 
  ELSEIF REQUEST ("Operacao") = 3 THEN ' Alterar
  call abreConexao
	sql = "UPDATE ExploracaoPecuaria SET TipoExploracao = '"&request.Form("TipoExploracao")&"' WHERE CPFCNPJProdutor = '"&request.Form("txtCNPJProdutor")&"' AND CodigoPropriedade = '"&request. Form("txtCodigoProp")&"'"
       conn.execute(sql)
 	   call  fechaConexao
response.Redirect("Cadastro.asp?Resp=2")
   END IF
  
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title> Vinculação de Propriedade/Produtor </title>
<script src="javascript/Mascara.js"></script>
<script type="text/javascript">

//<style>
//error {
	//border-color: #e8273b;
	//color: #FFF;
	//background-color: #ed5565;
//}
//</style>
function validar()
{
	if(document.frmCadastro.txtCodigoProp.value == "")
	{
		alert("Obrigatorio digitar o Codigo !");
		document.frmCadastro.txtCodigoProp.focus();
		return false;
	}
	if (document.frmCadastro.txtCNPJProdutor.value == "")
	{
		alert("Obrigatorio Digitar o CNPJ/CPF !");
		document.frmCadastro.txtCNPJProdutor.focus();
		return false;
		}
	if(document.frmCadastro.TipoExploracao.value == "")
	{
		alert("Obrigatorio Selecionar o TipoExploracao !");
		document.frmCadastro.TipoExploracao.focus();
		return false;
		}
	return true;
}
function cadastrar()
{
	if(validar() == false)
	return false;
	
	document.frmCadastro.Operacao.value = 1;
	document.frmCadastro.action = "Cadastro.asp";
	document.frmCadastro.submit();
}	
function Visualizar(cpf, Codigo)
{
	
	document.frmCadastro.Operacao.value = 2;
	document.frmCadastro.CPFVisualizar.value = cpf;
	document.frmCadastro.CodPropriedade.value = Codigo;
	document.frmCadastro.action = "Cadastro.asp";
	document.frmCadastro.submit();

}
function Alterar()
{
	if(validar() == false)
	return false;
	
	document.frmCadastro.Operacao.value = 3;
	document.frmCadastro.action  =  "Cadastro.asp";
	document.frmCadastro.submit();
	}
	function Novo()
{
	window.location.href = "Cadastro.asp";
	}


//funçao para quando clicar no botão ele abrir uma janela.
function abrir() {
  window.open("BuscarProp.asp", 'janela', 'width=795, height=590, top=100, left=699, scrollbars=no, status=no, toolbar=no, location=no, menubar=no, resizable=no, fullscreen=no')
}
function abrir2() {
  window.open("BuscarProdutor.asp", 'janela', 'width=795, height=590, top=100, left=699, scrollbars=no, status=no, toolbar=no, location=no, menubar=no, resizable=no, fullscreen=no')
}
function abrir3() {
  window.open("Exploracao.asp", 'janela', 'width=795, height=590, top=100, left=699, scrollbars=no, status=no, toolbar=no, location=no, menubar=no, resizable=no, fullscreen=no')
}
</script>
</head>
<center>
<%if request("Resp") = 1 or request("Resp") = 2 then
	if request("Resp") = 1 then
	mensagem = "Vinculação Realizada com Sucesso!"
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
                  <li class="">
                    <a data-toggle="tab" href="Propriedade.asp">Propriedade</a>
                  </li>
                  <li class="active">
                    <a data-toggle="tab" href="Cadastro.asp">Vinculacao</a>
                  </li>
                    <section class="panel">
                <ul class="nav nav-tabs pull-right">
                  <li class="">
                    <a data-toggle="tab" href="CadLogin.asp">Sair</a>
                  </li>
                
                </ul>
              </header>
<fieldset style="text-align:center" style="width:30%;" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cadastro De Vinculação
<form name="frmCadastro" onsubmit="validar();" id="frmCadastro" method="post">
<input type="hidden" name="Operacao" id="Operacao" />
<input type="hidden" name="CPFVisualizar" id="CPFVisualizar" />
<input type="hidden" name="CodPropriedade" id="CodPropriedade" />
Codigo da Propriedade:
<input type="text" name="txtCodigoProp"  id="txtCodigoProp" value="<%=Codigo%>" readonly  />
<input type="button"  onclick="abrir();"  value=" Propriedade"> 
</br>
<p>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CNPJ/CPF:
<input type="text" name="txtCNPJProdutor" id="txtCNPJProdutor"  value="<%=CPFCNPJProdutor%>"readonly/>
<input type="button" onclick="abrir2();" value=" Produtor"> </p>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<p>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tipo Exploracao:</br>
<%
call abreConexao
sql = "SELECT Codigo, Descricao FROM TipoExploracao order by Codigo"

set rs_Exploracao = conn.execute(sql)

%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<select name="TipoExploracao" id="TipoExploracao">
<option value="">Selecionar</option>
<%do while not rs_Exploracao.eof%>
<option value="<%=rs_Exploracao("Codigo")%>"<%if clng(rs_Exploracao("Codigo")) = TipoExploracao then%>selected<%end if%>><%=rs_Exploracao("Descricao")%></option>
<%
rs_Exploracao.movenext
loop
call fechaConexao%>
</select>
<p>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="btnCadastrar" value="<%if existe = 1 then%>Alterar Cadastro<%else%>Vincular Cadastro<%end if%>" onclick=" return <%if existe = 1 then%>Alterar();<%else%> cadastrar();<%end if%>"/>
<input type="button" onClick="return Novo();" value=" Novo">

</form>
</fieldset>
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
  <td align="center"><a href="#" onclick="Visualizar(<%=rs("CPF")%>, <%=rs("Codigo")%>)"><%=rs("CPF")%></a></td>
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

<%if existeCadastro = true then %>
<div class="alerta error">Este Produtor já esta vinculado a um Tipo Exploracao</div>
<%end if %>


</body>
</html>
