<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<%
IF REQUEST ("Operacao") = 1 THEN 'Cadastrar
	if request.ServerVariables("CONTENT_LENGTH") <> 0 then
	 call abreConexao
     sql = "INSERT INTO Produtor(CPF, CodigoMunicipio, Nome, Endereco, Telefone, CEP, Status) VALUES ('"&replace(replace(request.form("txtCPF"),".",""),"-","")&"', '"&request.form("municipio")&"','"&request.form("txtNome")&"','"&request.form("txtEndereco")&"', '"&request.form("txtTelefone")&"', '"&request.form("txtCEP")&"', 1)"
	 conn.execute(sql)
     call fechaConexao
	 session("CPF_usu") =  replace(replace(request.Form("txtCPF"),".",""),"-","")
	response.Redirect("index.asp?Operacao=2&CpfVisualizar="&replace(replace(request.form("txtCPF"),".",""),"-","")&"") 
	end if
ELSEIF REQUEST ("Operacao") = 2 THEN ' Visualizar
call abreConexao
sql = "select CPF, CodigoMunicipio, Nome, Endereco, Telefone, CEP, Status from Produtor  WHERE  CPF = '"&request("CpfVisualizar")&"'"

set rs = conn.execute(sql)
if not rs.eof then
CPF = rs("CPF") 
CodMunicipio = rs("CodigoMunicipio")
Nome =  rs("Nome")
Endereco = rs("Endereco")
Telefone = rs("Telefone")
CEP  = rs("CEP")
StatusUsuario = rs("Status")
CPF =  left((CPF),3)&"."&mid((CPF),4,3)&"."&mid((CPF),7,3)&"-"&right((CPF),2)
existe = 1

session("CPF_usu") = rs("CPF") 

 end if

ELSEIF REQUEST ("Operacao") = 3 THEN ' Alterar
 if request.ServerVariables("CONTENT_LENGTH")<> 0 then
 call abreConexao
 sql = "UPDATE Produtor SET CodigoMunicipio = '"&request.Form("municipio")&"', Nome = '"&request.Form("txtNome")&"', Endereco = '"&request.Form("txtEndereco")&"', Telefone = '"&request.Form("txtTelefone")&"', CEP = '"&request.Form("txtCEP")&"', Status = '"&request.Form("Status")&"' WHERE CPF= '"&replace(replace(request.Form("txtCPF"),".",""),"-","")&"' "
 conn.execute(sql)
 session("CPF_usu") =  replace(replace(request.Form("txtCPF"),".",""),"-","")
call fechaConexao
response.Redirect("index.asp")
  end if
END IF
%>

<!DOCTYPE HTML>
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
<title>Inscricao</title>
<script src="javascript/Mascara.js"></script>
<script type="text/javascript">
function validar()
{
	if(document.frmProdutor.txtCPF.value == "")
	{
		alert("Obrigatorio digitar o CPF");
		document.frmProdutor.txtCPF.focus();
		return false;
	}
	else
	 if(!ValidarCPF(document.frmProdutor.txtCPF.value.replace(".","").replace(".","").replace("-","")))
	 {
	   alert("CPF Invalido ou está na formatação errada 000.000.000-00");
	   return false
	 }
	if(document.frmProdutor.municipio.value == "")
	{
		alert("Obrigatorio Selecionar o Municipio" );
		document.frmProdutor.municipio.focus();
		return false;
		}
	if(document.frmProdutor.txtNome.value == "")
	{
		alert("Obrigatorio Digitar o Nome" );
		document.frmProdutor.txtNome.focus();
		return false;
		}	
	if(document.frmProdutor.txtEndereco.value == "")
	{
		alert("Obrigatorio Digitar o Endereco" );
		document.frmProdutor.txtEndereco.focus();
		return false;
		}	
	if(document.frmProdutor.txtTelefone.value == "")
	{
		alert("Obrigatorio Digitar o Celular" );
		document.frmProdutor.txtTelefone.focus();
		return false;
		}	
		else
	if(!ValidaCelular(document.frmProdutor.txtTelefone.value))
	{
	  document.frmProdutor.txtTelefone.focus();
	  return false
	}
	if(document.frmProdutor.txtCEP.value == "")
	{
		alert("Obrigatorio Digitar o CEP!" );
		document.frmProdutor.txtCEP.focus();
		return false;
		}		
		else
	if(!ValidaCep(document.frmProdutor.txtCEP.value))
	{
		document.frmProdutor.txtCEP.focus();
	  return false
	}
	return true;
}
function cadastrar()
{
	if(validar() == false)
	return false;
	
	document.frmProdutor.Operacao.value = 1;
	document.frmProdutor.action = "index.asp";
	document.frmProdutor.submit();
}
function Visualizar(cpf)
{
	document.frmProdutor.Operacao.value = 2;
	document.frmProdutor.CpfVisualizar.value = cpf;
	document.frmProdutor.action = "index.asp";
	document.frmProdutor.submit();
}
function verificar_cadastro()
{
	document.frmProdutor.Operacao.value = 2
	document.frmProdutor.CpfVisualizar.value = document.frmProdutor.txtCPF.value;
	document.frmProdutor.action = "index.asp"
	document.frmProdutor.submit();
	}
function Alterar()
{
	if(validar() == false)
	return false;
	
	document.frmProdutor.Operacao.value = 3;
	document.frmProdutor.action = "index.asp";
	document.frmProdutor.submit();
}	

function Novo()
{
	window.location.href = "index.asp";
	}

</script>

</head>
<body>
  <section class="panel">
              <header class="panel-heading tab-bg-primary ">
                <ul class="nav nav-tabs">
                  <li class="active">
                    <a data-toggle="tab" href="index.asp">Home</a>
                  </li>
                  <li class="">
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
              
              
<form name="frmProdutor" onsubmit="validar();" id="frmProdutor" method="post">
<input type="hidden" name="Operacao" id="Operacao">
<input type="hidden" name="CpfVisualizar" id="CpfVisualizar">
<p>
 CPF:<br/>
<input type="text" onKeyPress="MascaraCPF(txtCPF)" name="txtCPF" id="txtCPF" value="<%if existe = 1 then response.Write(CPF) ELSE response.Write(Request("txtCPF")) end if  %>" maxlength="14" onblur="MascaraCPF(txtCPF);verificar_cadastro();"<%IF REQUEST ("Operacao") = 2 AND Existe = 1 THEN%> readonly<%END IF%>/>

<p/>
<p>
Municipio:</br>
<%
call abreConexao
sql = "SELECT Codigo, Nome FROM Municipio order by Nome"

set rs = conn.execute(sql)

%>
<select name="municipio" id="municipio">
<option value="">Selecionar</option>
<%do while not rs.eof%>
<option value="<%=rs("Codigo")%>" <%IF rs("Codigo") = CodMunicipio THEN%>SELECTED<%END IF%>><%=rs("Nome")%></option>
<%
rs.movenext
loop
call fechaConexao%>
</select>

</p>
<p>
Nome: <br/>
<input type="text" name="txtNome" id= "txtNome" placeholder="Seu Nome" value="<%=Nome%>" />
<p/>

<p> 
Endereco:<br/>
<input type="text" name="txtEndereco" id= "txtEndereco" placeholder="Seu Endereco" value="<%=Endereco%>" />
<p/>

<p> 
Telefone:<br/>
<input type="text" name="txtTelefone" id= "txtTelefone" placeholder="Contato" maxlength="15"  value="<%=Telefone%>"  onKeyPress="MascaraCelular(txtTelefone)" onBlur="MascaraCelular(txtTelefone)"/>
<p/>

<p>
CEP: <br/>
<input type="text" name="txtCEP" id= "txtCEP" placeholder="Digite seu Cep"  maxlength="10" value="<%=Cep%>" onKeyPress="MascaraCep(txtCEP)" onBlur="MascaraCep(txtCEP)">
<p/>
<%IF existe = 1 THEN%>
<p>
Status:<select name="Status" ></br>
<option  value="1" <% if StatusUsuario = true then %> selected<%end if%>>Ativo</option>
<option value="0" <% if StatusUsuario = false  then%> selected<% end if %>>Desativado</option>
</select> 
</p>
<%END IF%>

<input type="submit" name="btnCadastro" value="<%if existe = 1 then%>Alterar<%else%>Cadastrar<%end if%>" onclick="return <%if existe = 1 then%>Alterar();<%else%>cadastrar();<%end if%>"/>
<%if existe = 1 then%>
<input type="button" onClick="return Novo();" value=" Novo Cadastro">
<a href = "propriedade.asp"><input type="button" value="Ir Propriedade"></a> 
<%end if%>
</form>
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
  <th>Operacao</th>
  </tr>
  <%do while not rs.eof%>
  <tr>
  <td align="center"><%response.write rs("CPF")%></td>
  <td align="center"><%=rs("Nome")%></td>
  <td align="center"><%=rs("Endereco")%></td>
  <td align="center"><%=rs("Telefone")%></td>

  <td align="center"><a href="#" onclick="Visualizar(<%=rs("CPF")%>)"><img src="Imagem\editar.png" width="30"/></a></td>
  </tr>
  <%
     rs.movenext
	 loop
  %>
  <%end if%>
</table>
<%call fechaConexao%>

</body>
</HTML>
 



