<!DOCTYPE html>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/Conexao.asp"-->
<%


dim existeCadastro
IF REQUEST("Operacao") = 1 THEN
call abreConexao
sql = "SELECT CPF,Senha from CadLogin WHERE CPF = '"&replace(replace(request.form("txtCPF"),".",""),"-","")&"' AND Senha = '"&   request.Form("pswSenha")&"'" 
set rs = conn.Execute(sql)
		if not rs.eof then
		if replace(replace(request.form("txtCPF"),".",""),"-","") = rs("CPF") and request.form("pswSenha") = rs("Senha") then
			existeCadastro = true
		else
			existeCadastro = false
		rs.close
		set rs = nothing
		call fechaConexao
		end if
		end if
if existeCadastro = true then
response.Redirect("index.asp")
elseif existeCadastro = false then
response.write("Usuario ou Senha Incorreto")
else
response.Write("Usuario ou Senha Incorreto!")
end if
end if 

%>          

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="description" content="Creative - Bootstrap 3 Responsive Admin Template">
  <meta name="author" content="GeeksLabs">
  <meta name="keyword" content="Creative, Dashboard, Admin, Template, Theme, Bootstrap, Responsive, Retina, Minimal">
  <link rel="shortcut icon" href="img/favicon.png">

  <title>Home</title>
  <script src="javascript/Mascara.js"></script>
  <script type="text/javascript">
function validar()

{
	if(document.frmCadLogin.txtCPF.value == "")
	{
	
		alert("Obrigatorio digitar o CPF");
		document.frmCadLogin.txtCPF.focus();
		return false;
	}
	else
	 if(!ValidarCPF(document.frmCadLogin.txtCPF.value.replace(".","").replace(".","").replace("-","")))
	 {
	   alert("CPF Invalido ou está na formatação errada 000.000.000-00");
	   return false

	 
  }
  
	  if(document.frmCadLogin.pswSenha.value == "")
	  {
		  alert("Obrigatorio Digitar a Senha !");
		  document.frmCadLogin.pswSenha.focus();
		  return false;
		  }
	  
	  return true;
  }
function login()
{
	if(validar() == false)
	  return false;
	document.frmCadLogin.Operacao.value = 1;
	document.frmCadLogin.action = "CadLogin.asp";
	document.frmCadLogin.submit();
}  
  
  </script>
  <!-- Bootstrap CSS -->
  <link href="css/bootstrap.min.css" rel="stylesheet">
  <!-- bootstrap theme -->
  <link href="css/bootstrap-theme.css" rel="stylesheet">
  <!--external css-->
  <!-- font icon -->
  <link href="css/elegant-icons-style.css" rel="stylesheet" />
  <link href="css/font-awesome.css" rel="stylesheet" />
  <!-- Custom styles -->
  <link href="css/style.css" rel="stylesheet">
  <link href="css/style-responsive.css" rel="stylesheet" />

  <!-- HTML5 shim and Respond.js IE8 support of HTML5 -->
  <!--[if lt IE 9]>
    <script src="js/html5shiv.js"></script>
    <script src="js/respond.min.js"></script>
    
    <![endif]-->

    <!-- =======================================================
      Theme Name: NiceAdmin
      Theme URL: https://bootstrapmade.com/nice-admin-bootstrap-admin-html-template/
      Author: BootstrapMade
      Author URL: https://bootstrapmade.com
    ======================================================= -->
</head>


<body class=" login-img3-body">

  <div class="container">
    <form name="frmCadLogin" class="login-form" method="post" id="frmCadLogin">
    <input type="hidden" name="Operacao" id="Operacao">    
      <div class="login-wrap">
        <p class="login-img">CADASTRO</p>
        <div class="input-group">
          <span class="input-group-addon"><i class="icon_profile"></i></span>
          <input type="text" onKeyPress="MascaraCPF(txtCPF)" name="txtCPF"  id="txtCPF" class="form-control" placeholder="CPF" autofocus maxlength="14"   onBlur="MascaraCPF(txtCPF)">
        </div>
        <div class="input-group">
          <span class="input-group-addon"><i class="icon_key_alt"></i></span>
          <input type="password" class="form-control" name="pswSenha" id="pswSenha" placeholder="Senha" maxlength="10">
        </div>
        <label class="checkbox">
                <span class="pull-right"> <a href="#"> Esqueceu a Senha?</a></span>
            </label> 
        <button class="btn btn-primary btn-lg btn-block" type="submit" onClick="return login();"> Entrar</button>
        <button class="btn btn-info btn-lg btn-block" type="submit"> <a href="CadDeUsu.asp">Cadastre-se</a></button>
        
      </div>
    </form>

    <div class="text-right">
      <div class="credits">
          Desenvolvido por <a href="http://sidato.adapec.to.gov.br/sistemas/">sidato.adapec.to.gov.br/sistemas/</a>
        </div>
    </div>
  </div>


</body>

</html>
