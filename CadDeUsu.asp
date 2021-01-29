<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<%
IF REQUEST ("Operacao") = 1 THEN 'CADASTRAR
call abreConexao
sql = "INSERT into CadLogin(CPF,Senha) VALUES ('"&replace(replace(request.form("txtCPF"),".",""),"-","")&"' ,'"&request.Form("pswSenha")&"')"
conn.execute (sql)
call fechaConexao 
response.Redirect("CadLogin.asp")
END IF
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Cadastro de Usuario</title>
<script src="javascript/Mascara.js"></script>
<script type="text/javascript">
function validar()
{
	if(document.frmCadUsu.txtCPF.value == "")
	{
		alert("Obrigatorio digitar o CPF !");
		document.frmCadUsu.txtCPF.focus();
		return false;
	}
	if (document.frmCadUsu.pswSenha.value == "")
	{
		alert("Obrigatorio Digitar a Senha !");
		document.frmCadUsu.pswSenha.focus();
		return false;
		}
	return true;
}
function Cadastrar()
{
	if(validar() == false )
	return false;
	
	document.frmCadUsu.Operacao.value = 1;
	document.frmCadUsu.action = "CadDeUsu.asp";
	document.frmCadUsu.submit();
	
	}
function Novo()
	{
	 window.location.href = "CadDeUsu.asp";
		}

</script>
</head>
<center>
<%if request("Resp") = 1 then
	mensagem = "Cadastro Realizado com Sucesso!"
%>
<div class="messagem_informacao" align="center"><%=mensagem%></div>
<%end if%>
</center>

<body>
<fieldset style="text-align:center" style="width:30%;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cadastro
<form name="frmCadUsu" onsubmit="validar();" id="frmCadUsu" method="post">
<input type="hidden" name="Operacao" id="Operacao" />
CPF:
<input type="text" onkeypress="MascaraCPF(txtCPF)" name="txtCPF" id="txtCPF" maxlength="14" onblur="MascaraCPF(txtCPF)" />
<p>
Senha:
<input type="password" name="pswSenha" id="pswSenha" maxlength="10"/>
</form>
<p>
<input type="submit" onclick="return Cadastrar();"  value="Cadastrar" />
<input type="button"  onclick="return Novo();" value="Novo" />
</p>

</fieldset>

</body>
</html>
