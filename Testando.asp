<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<%
if request("Operacao") = 1 then
call abreConexao
sql =" INSERT INTO Testando(Nome,Endereco,Cep) VALUES ('"&request.Form("txtNome")&"', '"&request.Form("txtEndereco")&"', '"&request.Form("txtCep")&"') "
conn.execute(sql)
call fechaConexao
end if 
%>

<%
'dim  cpf
'CPF = "04980227147"
'CPF = left ((CPf),3)&"."&mid((CPF),4,3)&"."&mid((CPF),7,3)&"."&right((Cpf),2)
'response.Write(CPF)
'response.End()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Testando</title>
<script type="text/javascript">
function validar()
{
if(document.frmTestando.txtNome.value == "")
{
	alert("Obrigatorio Digitar o Nome!");
	document.frmTestando.txtNome.focus();
	return false;	
	
	}
}
function cadastrar()
{
	if(validar()== false)
	return false;
	
	document.frmTestando.Operacao.value = 1;
	document.frmTestando.action = "Testando.asp";
	document.frmTestando.focus();
	}
</script>
</head>
<body>
<fieldset style="text-align:center" style="width:30;">Cadastro
<form  name="frmTestando"  onsubmit="validar()" id="frmTestando" method="post">
<input type="hidden" id="Operacao" name="Operacao" />
Nome:
<input type="text" name="txtNome"  id="txtNome" />
<p>
Endereco:
<input type="text" name="txtEndereco" id="txtEndereco" />
<p>

Cep:
<input type="text" name="txtCep" id="txtCep" /><p>

</p>
<input type="submit"  onclick="return cadastrar(); "value="Cadastrar" />
</form>
</fieldset>
</body>
</html>
gffhfghfghfhfgh