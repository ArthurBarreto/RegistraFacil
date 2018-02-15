<!--#include virtual="/registrafacil/include/db_open.asp"-->
<!--#include virtual="/registrafacil/FCKeditor/fckeditor.asp"-->

<%
wCNPJ_CPF = Request.QueryString("cnpj_cpf")
wCod_Grupo = Request.QueryString("grupo")
call logado()
Session.LCID = 1046
'Response.Write("SELECT a.status,a.cod_grupo,a.cod_sub,a.status_patente,a.status_marca,a.status_avaliacao,a.status_software,a.status_internacional,a.status_direito,a.status_juridico, a.EMITENOTA, a.SUBST_TRIB, a.TIPO_TRIBUTACAO , * FROM GEDCLIENTES AS a LEFT JOIN GEDUSU AS b on a.CNPJ_CPF = b.CNPJ_CPF Where a.CNPJ_CPF='" & Tira_Pontuacao(wCNPJ_CPF) & "' and a.COD_GRUPO='" & wCod_Grupo & "'")'
'Response.End()'

rs.open "SELECT a.status,a.cod_grupo,a.cod_sub,a.status_patente,a.status_marca,a.status_avaliacao,a.status_software,a.status_internacional,a.status_direito,a.status_juridico, a.*, b.desusuario, b.senha FROM GEDCLIENTES AS a LEFT JOIN GEDUSU AS b on a.CNPJ_CPF = b.CNPJ_CPF Where a.CNPJ_CPF='" & Tira_Pontuacao(wCNPJ_CPF) & "' and a.COD_GRUPO='" & wCod_Grupo & "'", conn, 3, 3
IF NOT RS.EOF THEN
   wstatus_cliente = rs.fields(0)
END IF
%>
<html>
<head>
<title>RegistraFacil - Alteração de Cliente </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link id="luna-tab-style-sheet" type="text/css" rel="stylesheet" href="../tabs/tabpane.css" />
<script type="text/javascript" src="../tabs/tabpane.js"></script>
<script language="javascript" src="../script/script.js" type="text/javascript"></script>
<script language="javascript" src="../include/micoxAjax.js" type="text/javascript"></script>
<script type="text/javascript" src="../js/mascaras.js"></script>
<script language="javascript" src="../js/validacoes.js" type="text/javascript"></script>

<script language="javascript">

function mens_orgao(){
		alert("Não esqueça de cadastrar o Orgão Expedidor junto com a Identidade. Ex: (SSP/CE, MD/CD, SSP/RS)");
}


function Formulario(){
	if (document.frm.cod_sub.value != ''){
		document.getElementById('status_marca').checked = true;
	}

	if (document.frm.cod_sub_patente.value != '')
	{
		document.getElementById('status_patente').checked = true;
	}
	else{
		document.getElementById('status_patente').checked = false;
	}

	if (document.frm.cod_sub_avaliacao.value != '')
	{
		document.getElementById('status_avaliacao').checked = true;
	}
	else{
		document.getElementById('status_avaliacao').checked = false;
	}

	if (document.frm.cod_sub_internacional.value != '')
	{
		document.getElementById('status_internacional').checked = true;
	}
	else{
		document.getElementById('status_internacional').checked = false;
	}

	if (document.frm.cod_sub_software.value != '')
	{
		document.getElementById('status_software').checked = true;
	}
	else{
		document.getElementById('status_software').checked = false;
	}

	var direito = document.frm.cod_sub_direito.value;
	if (direito != '')
	{
		document.getElementById('status_direito').checked = true;
	}
	else{
		document.getElementById('status_direito').checked = false;
	}

	MM_validateForm('login','','R','RG','','R','nome','','R','data_nasc','','R','endereco','','R','bairro','','R','cep','','R','email1','','R','senha','','R');
	if(document.MM_returnValue){
		if(document.frm.fone_res.value == "" && document.frm.fone_com.value == "" && document.frm.celular.value == "" && document.frm.fax.value == ""){
			alert('É necessário preencher ao menos um dos telefones');
		} else {
			if(ValidaNome(document.frm.nome) == true && ValidaNome(document.frm.contato1) == true && ValidaNome(document.frm.representante1) == true && ValidaNome(document.frm.representante2) == true && ValidaNome(document.frm.representante3) == true){
  			    document.frm.submit();
				//alert(ValidaNome(document.frm.nome));
			}
		}
	} else {
	 	return document.MM_returnValue;
	}
}
MM_preloadImages('../tabs/images/tab.png','../tabs/images/tab_active.png','../tabs/images/tab_hover.png');
function copia_dados(){
document.frm.representante1.value = document.frm.contato1.value;
document.frm.aniver_representante1.value = document.frm.aniver1.value;
document.frm.cargo_representante1.value = document.frm.cargo1.value;
document.frm.email_representante1.value = document.frm.email_contato1.value;
}

function copia_endereco(valor){
	if (valor == 1){
		document.getElementById('endereco2').value = document.getElementById('endereco').value;
		document.getElementById('numero2').value = document.getElementById('numero').value;
		document.getElementById('complemento2').value = document.getElementById('complemento').value;
		document.getElementById('bairro2').value = document.getElementById('bairro').value;
		document.getElementById('cep2').value = document.getElementById('cep').value;
		document.getElementById('txt_estado2').selectedIndex = document.getElementById('txt_estado').selectedIndex;
		document.getElementById('txt_cidade2').selectedIndex = document.getElementById('txt_cidade').selectedIndex;
		document.getElementById('pais2').selectedIndex = document.getElementById('pais').selectedIndex;
	}else{
		document.getElementById('endereco3').value = document.getElementById('endereco').value;
		document.getElementById('numero3').value = document.getElementById('numero').value;
		document.getElementById('complemento3').value = document.getElementById('complemento').value;
		document.getElementById('bairro3').value = document.getElementById('bairro').value;
		document.getElementById('cep3').value = document.getElementById('cep').value;
		document.getElementById('txt_estado3').selectedIndex = document.getElementById('txt_estado').selectedIndex;
		document.getElementById('txt_cidade3').selectedIndex = document.getElementById('txt_cidade').selectedIndex;
		document.getElementById('pais3').selectedIndex = document.getElementById('pais').selectedIndex;
	}
}

function validaRepresentante()
{
	d = document.getElementById;
	ValidaNome(document.frm.representante1);
	if (d("representante1").value == "" || d("cargo_representante1").value == "" || d("cpf_representante1").value == "" || d("email_representante1").value == ""|| d("aniver_representante1").value == "" || d("rg_representante1").value == "")
	{
		alert("Todos os dados do Primeiro Representante devem ser preenchidos!");
		d("representante1").focus();
		return false;
	}
	else
	{
		if (d("representante2").value != "" && (d("cargo_representante2").value == "" || d("cpf_representante2").value == "" || d("email_representante2").value == ""|| d("aniver_representante2").value == "" || d("rg_representante2").value == "") )
		{
			alert("Preencha todos os dados do segundo representante!");
			d("representante2").focus();
			return false;
		}

		if (d("representante3").value != "" && (d("cargo_representante3").value == "" || d("cpf_representante3").value == "" || d("email_representante3").value == "" || d("aniver_representante3").value == "" || d("rg_representante3").value == "") )
		{
			alert("Preencha todos os dados do terceiro representante!");
			d("representante3").focus();
			return false;
		}
	}
	return true;
}

function validarEmail(email,campo){
	if (!checkMail(email) && document.getElementById(campo).value != "") {
		alert("Email inválido, tente novamente");
		document.getElementById(campo).value="";
		document.getElementById(campo).focus();
	}
}

	function checkMail(mail){
        var er = new RegExp(/^[A-Za-z0-9_\-\.]+@[A-Za-z0-9_\-\.]{2,}\.[A-Za-z0-9]{2,}(\.[A-Za-z0-9])?/);
        if(typeof(mail) == "string"){
                if(er.test(mail)){ return true; }
        }else if(typeof(mail) == "object"){
                if(er.test(mail.value)){
                                        return true;
                                }
        }else{
                return false;
                }
	}
	

</script>
<style type="text/css">
<!--
.style2 {font-size: 9px}
.style3 {font-size: 12px}
.style26 {color: #FFFFFF}
.style27 {
	color: #000000;
	font-weight: bold;
}
-->
</style>

</head>
<body>
<%
If rs.EOF then
	Call CaixaMensagem ("erro","ERRO!","Cliente não encontrado. Certifique-se que o mesmo esteja cadastrado",true,"Avançar",false,"")
	response.End()
else
%>
<fieldset>
<legend>
<table border="0">
  <tr>
    <td width="32"><div align="center"><span class="style2 style3"><img src="../imagens/btn_web-users_bg.gif" width="32" height="32"></span></div></td>
    <td width="154"><div align="left" class="style2 style3">Altera&ccedil;&atilde;o de Clientes:</div></td>
  </tr>
</table>
</legend>
<table  border="0" align="center">
  <tr>
    <td width="828"  valign="top"><form action="alt_cli1.asp" method="post"  name="frm" target="_self">
      <table width="100%" border="0" align="center">
          <tr>
            <td height="15" colspan="6" bgcolor="#d0d8f7" class="txt_normal"><div align="center" class="style27">DADOS <span class="txt_normal ">CADASTRAIS</span></div></td>
          </tr>
          <tr>
            <td colspan="3" class="txt_normal"> Data Cadastro: <%
			sql = "SELECT a.data_cad FROM GEDCLIENTES AS a LEFT JOIN GEDUSU AS b on a.CNPJ_CPF = b.CNPJ_CPF Where a.CNPJ_CPF='" & Tira_Pontuacao(wCNPJ_CPF) & "'"
			rs1.open sql, conn, 3, 3
			if not rs.eof then
			   wdata_cad = rs1.fields(0)
			end if

			if (wdata_cad) < (date) then
			   response.write DATE_BR(wdata_cad)
			else
			   response.write "15/01/2007"
            end if
			%>              </td>
            <td colspan="2" class="txt_normal"><div align="left">Data  Altera&ccedil;&atilde;o: <%=DATE_BR(rs("DATA_ULT_ALT"))%>
                <input name="DATA_ULT_ALT" type="hidden" id="DATA_ULT_ALT" value="<%=DATE_BR(DATE())%>">
            </div></td>
          </tr>
          <tr>
            <td colspan="5" class="txt_normal"><strong>Login:</strong>              <input name="login" type="text" class="caixa_mensagem" id="login" value="<%=rs("Desusuario")%>" size="15" maxlength="30">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Senha:</strong>              <input name="senha" type="text" class="caixa_mensagem" id="Senha" value="<%=rs("senha")%>" size="15" maxlength="10">
              <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Status:</strong>              <select name="status" class="txt_normal" id="status">
                <option value="1" <%if wstatus_cliente = "1" then response.write "selected"%>>Ativo</option>
                <option value="0"<%if wstatus_cliente = "0" then response.write "selected"%>>Inativo</option>
              </select>
			  <span class="aviso">&nbsp;
			  <%
			  'if rs("codigo_cnpj") = "9" then'
			  '   response.write "Login e senha temporário."'
			  'end if'
			  %>
			  </span>

			  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <%
			if session("admin") = 1 then
			%>

			  <strong>&nbsp; tipo:</strong>
              <select name="perfil" class="txt_normal" id="perfil">
                <option value="1" <%if rs("perfil") = "1" then response.write "selected"%>>Cliente</option>
                <option value="2"<%if rs("perfil") = "2" then response.write "selected"%>>Prospect</option>
              </select>
			  <%
			  end if
			  %>			  </td>
          </tr>

          <tr>
            <td class="txt_normal">nome indica&ccedil;&atilde;o:</td>
            <td colspan="4" class="txt_normal"><input name="nome_indicacao" type="text" class="txt_normal" id="nome_indicacao" value="<%=rs("nome_indicacao")%>" style="width:540;" maxlength="70"></td>
            </tr>
          <tr>
            <td class="txt_normal"><strong>CNPJ/CPF:</strong></td>
            <td colspan="2" class="txt_normal"><%=FormataCNPJCPF(Request("cnpj_cpf"))%>
            <input name="cnpj_cpf" type="hidden" id="cnpj_cpf" value="<%=Request("cnpj_cpf")%>"></td>
            <td width="104" class="txt_normal"><div align="right"><strong>RG/INSCRI&Ccedil;&Atilde;O:</strong></div></td>
            <td width="249" class="txt_normal"><input name="RG" type="text" class="caixa_mensagem" id="CGF/RG" value="<%=rs("rg")%>" size="25" maxlength="50" onFocus="mens_orgao();"></td>
          </tr>		  
          <tr>
          <tr>
            <td class="txt_normal"><strong>empresa:</strong></td>
            <td colspan="4" class="txt_normal"><span class="caixa_mensagem"><strong>
              <%
			if session("admin") = 1 then
			%>
              </strong></span><span class="caixa_mensagem"><strong>
              <strong>
              <select name="cod_grupo" class="txt_normal" id="cod_grupo" onChange="ajaxGet('../include/carrega_solicitante.asp?cnpj_cpf='+this.value,document.getElementById('solicitante'),true);"style="width:544;">
                <option selected>--- SELECIONE ---</option>
                <%
			  set rsFranquia = Server.CreateObject("ADODB.Recordset")

			  rsFranquia.open "SELECT cod_grupo, DesGrupo FROM gedgrp order by DesGrupo", conn, 2, 3
			  if not rsFranquia.EOF then
			  	while not rsFranquia.EOF
				if rsFranquia("cod_grupo") = rs(1) then
					response.Write("<option value='" & rsFranquia("cod_grupo") & "' selected>" & rsFranquia("DesGrupo") & "</option>")
					else
					response.Write("<option value='" & rsFranquia("cod_grupo") & "'>" & rsFranquia("DesGrupo") & "</option>")
				end if
				rsFranquia.MoveNext
				wend
			  end if
			  rsfranquia.close
			%>
              </select>
              </strong>
              <%
			else
				Response.Write ExibeNome(session("franquia"), "franquia")
				%>
              <input name="cod_grupo" type="hidden" id="cod_grupo" value="<%=session("franquia")%>">
              <% end if %>
              </strong></span></td>
          </tr>
          <tr>
            <td colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center" class="style27">consultores<span class="txt_normal "></span></div></td>
          </tr>
          <tr>
            <td class="txt_normal">
              marca: </td>
            <td colspan="4" class="txt_normal"><div align="left" id="solicitante">
              <%
			if session("admin") = 1 then
			%>
              <select name="cod_sub" class="txt_normal" style="width:544;">
			  <option>--- SELECIONE ---</option>
                <%
				Set rsSub = conn.execute("SELECT CNPJ_CPF, NOME FROM GEDCLIENTES WHERE COD_GRUPO='" & rs(1) & "' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 ORDER BY NOME")

				while not rsSub.EOF
					if rsSub("CNPJ_CPF") = rs(2) then
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "' selected>"& rsSub("nome") & "</option>")
					else
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "'>"& rsSub("nome") & "</option>")
					end if
					rsSub.MoveNext
				wend
				rssub.close
				%>
              </select>
              <%
				else
					if session("admin") = 3 or session("admin") = 2 then
					%>
              <select name="cod_sub" class="txt_normal" id="cod_sub" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
						  set rsSubFranqueado = Server.CreateObject("ADODB.Recordset")
						  rsSubFranqueado.open "select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Tira_Pontuacao(rs(1)) &"' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 order by  nome", conn, 2, 3
							while not rsSubFranqueado.EOF
							if rsSubFranqueado("cnpj_cpf") = rs(2) then
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "' selected>" & rsSubFranqueado("nome") & "</option>")
								else
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "'>" & rsSubFranqueado("nome") & "</option>")
							end if
							rsSubFranqueado.MoveNext
							wend
						rssubfranqueado.close
						%>
              </select>
              <%
				end if

				if session("admin") = 4 or session("admin") = 0 or session("admin") = 8  or session("admin") = 5 or session("admin") = 6 then
					'Response.Write(rs(2))'
					Response.Write ExibeNome(rs(2), "subfranqueado2")
					%>
              <input name="cod_sub" type="hidden" id="cod_sub" value="<%=rs("cod_sub")%>">
              <%
				end if
		end if
				%>
            </div></td>
          </tr>

          <tr>
            <td class="txt_normal">
              patente: </td>
            <td colspan="4" class="txt_normal"><div align="left" id="solicitante_patente">
              <%
			if session("admin") = 1 then
			%>

              <select name="cod_sub_patente" class="txt_normal" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				Set rsSub = conn.execute("SELECT CNPJ_CPF, NOME FROM GEDCLIENTES WHERE COD_GRUPO='" & rs(1) & "' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 ORDER BY NOME")

				while not rsSub.EOF
					if rsSub("CNPJ_CPF") = rs("cod_sub_patente") then
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "' selected>"& rsSub("nome") & "</option>")
						else
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "'>"& rsSub("nome") & "</option>")
					end if
					rsSub.MoveNext
				wend
				rssub.close
				%>
              </select>
              <%
				else
					if session("admin") = 3 or session("admin") = 2 then
					%>
              <select name="cod_sub_patente" class="txt_normal" id="cod_sub_patente" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				          RESPONSE.WRITE rs("cod_sub_patente")
						  set rsSubFranqueado = Server.CreateObject("ADODB.Recordset")

						  rsSubFranqueado.open "select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Tira_Pontuacao(Request.QueryString("grupo")) &"' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 order by  nome", conn, 2, 3
						  if not rsSubFranqueado.EOF then
							while not rsSubFranqueado.EOF
							if rsSubFranqueado("cnpj_cpf") = rs("cod_sub_patente") then
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "' selected>" & rsSubFranqueado("nome") & "</option>")
								else
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "'>" & rsSubFranqueado("nome") & "</option>")
							end if
							rsSubFranqueado.MoveNext
							wend
						end if
						rssubfranqueado.close
						%>
              </select>
              <%
				end if

				if session("admin") = 4 or session("admin") = 0 or session("admin") = 8 or session("admin") = 5 or session("admin") = 6 then
					'Response.Write(rs("cod_sub_patente"))'
					Response.Write ExibeNome(rs("cod_sub_patente"), "subfranqueado2")
					%>
              <input name="cod_sub_patente" type="hidden" id="cod_sub_patente" value="<%=rs("cod_sub_patente")%>">
              <%
				end if
		end if
				%>
           </div></td>
          </tr>



          <tr>
            <td class="txt_normal"> avalia&ccedil;&atilde;o:</td>
            <td colspan="4" class="txt_normal">
			              <%
			if session("admin") = 1 then
			%>
              <strong><strong>
              <select name="cod_sub_avaliacao" class="txt_normal" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				Set rsSub = conn.execute("SELECT CNPJ_CPF, NOME FROM GEDCLIENTES WHERE COD_GRUPO='" & rs(1) & "' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 ORDER BY NOME")

				while not rsSub.EOF
					if rsSub("CNPJ_CPF") = rs("cod_sub_avaliacao") then
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "' selected>"& rsSub("nome") & "</option>")
						else
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "'>"& rsSub("nome") & "</option>")
					end if
					rsSub.MoveNext
				wend
				rssub.close
				%>
              </select>
              </strong></strong>
              <%
				else
					if session("admin") = 3 or session("admin") = 2 then
					%>
              <select name="cod_sub_avaliacao" class="txt_normal" id="cod_sub_avaliacao" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				          RESPONSE.WRITE rs("cod_sub_avaliacao")
						  set rsSubFranqueado = Server.CreateObject("ADODB.Recordset")

						  rsSubFranqueado.open "select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Tira_Pontuacao(Request.QueryString("grupo")) &"' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 order by  nome", conn, 2, 3
						  if not rsSubFranqueado.EOF then
							while not rsSubFranqueado.EOF
							if rsSubFranqueado("cnpj_cpf") = rs("cod_sub_avaliacao") then
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "' selected>" & rsSubFranqueado("nome") & "</option>")
								else
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "'>" & rsSubFranqueado("nome") & "</option>")
							end if
							rsSubFranqueado.MoveNext
							wend
						end if
						rssubfranqueado.close
						%>
              </select>
              <%
				end if

				if session("admin") = 4 or session("admin") = 0 or session("admin") = 8 or session("admin") = 5 or session("admin") = 6 then
					Response.Write ExibeNome(rs("cod_sub_avaliacao"), "subfranqueado2")
					%>
              <input name="cod_sub_avaliacao" type="hidden" id="cod_sub_avaliacao" value="<%=rs("cod_sub_avaliacao")%>">
            <%
				end if
		end if
				%>			</td>
          </tr>
          <tr>
            <td class="txt_normal"> internacional:</td>
            <td colspan="4" class="txt_normal">
			              <%
			if session("admin") = 1 then
			%>
              <strong><strong>
              <select name="cod_sub_internacional" class="txt_normal" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				Set rsSub = conn.execute("SELECT CNPJ_CPF, NOME FROM GEDCLIENTES WHERE COD_GRUPO='" & rs(1) & "' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 ORDER BY NOME")

				while not rsSub.EOF
					if rsSub("CNPJ_CPF") = rs("cod_sub_internacional") then
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "' selected>"& rsSub("nome") & "</option>")
						else
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "'>"& rsSub("nome") & "</option>")
					end if
					rsSub.MoveNext
				wend
				rssub.close
				%>
              </select>
              </strong></strong>
              <%
				else
					if session("admin") = 3 or session("admin") = 2 then
					%>
              <select name="cod_sub_internacional" class="txt_normal" id="cod_sub_internacional" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				          RESPONSE.WRITE rs("cod_sub_internacional")
						  set rsSubFranqueado = Server.CreateObject("ADODB.Recordset")

						  rsSubFranqueado.open "select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Tira_Pontuacao(Request.QueryString("grupo")) &"' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 order by  nome", conn, 2, 3
						  if not rsSubFranqueado.EOF then
							while not rsSubFranqueado.EOF
							if rsSubFranqueado("cnpj_cpf") = rs("cod_sub_internacional") then
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "' selected>" & rsSubFranqueado("nome") & "</option>")
								else
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "'>" & rsSubFranqueado("nome") & "</option>")
							end if
							rsSubFranqueado.MoveNext
							wend
						end if
						rssubfranqueado.close
						%>
              </select>
              <%
				end if

				if session("admin") = 4 or session("admin") = 0 or session("admin") = 8 or session("admin") = 5 or session("admin") = 6 then
					Response.Write ExibeNome(rs("cod_sub_internacional"), "subfranqueado2")
					%>
              <input name="cod_sub_internacional" type="hidden" id="cod_sub_internacional" value="<%=rs("cod_sub_internacional")%>">
            <%
				end if
		end if
				%>			</td>
          </tr>
          <tr>
            <td class="txt_normal"> SOFTWARE:</td>
            <td colspan="4" class="txt_normal">
			              <%
			if session("admin") = 1 then
			%>
              <strong><strong>
              <select name="cod_sub_software" class="txt_normal" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				Set rsSub = conn.execute("SELECT CNPJ_CPF, NOME FROM GEDCLIENTES WHERE COD_GRUPO='" & rs(1) & "' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 ORDER BY NOME")

				while not rsSub.EOF
					if rsSub("CNPJ_CPF") = rs("cod_sub_software") then
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "' selected>"& rsSub("nome") & "</option>")
						else
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "'>"& rsSub("nome") & "</option>")
					end if
					rsSub.MoveNext
				wend
				rssub.close
				%>
              </select>
              </strong></strong>
              <%
				else
					if session("admin") = 3 or session("admin") = 2 then
					%>
              <select name="cod_sub_software" class="txt_normal" id="cod_sub_software" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				          RESPONSE.WRITE rs("cod_sub_software")
						  set rsSubFranqueado = Server.CreateObject("ADODB.Recordset")

						  rsSubFranqueado.open "select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Tira_Pontuacao(Request.QueryString("grupo")) &"' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 order by  nome", conn, 2, 3
						  if not rsSubFranqueado.EOF then
							while not rsSubFranqueado.EOF
							if rsSubFranqueado("cnpj_cpf") = rs("cod_sub_software") then
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "' selected>" & rsSubFranqueado("nome") & "</option>")
								else
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "'>" & rsSubFranqueado("nome") & "</option>")
							end if
							rsSubFranqueado.MoveNext
							wend
						end if
						rssubfranqueado.close
						%>
              </select>
              <%
				end if

				if session("admin") = 4 or session("admin") = 0 or session("admin") = 8 or session("admin") = 5 or session("admin") = 6 then
					Response.Write ExibeNome(rs("cod_sub_software"), "subfranqueado2")
					%>
              <input name="cod_sub_software" type="hidden" id="cod_sub_software" value="<%=rs("cod_sub_software")%>">
            <%
				end if
		end if
				%>			</td>
          </tr>
          <tr>
            <td class="txt_normal"> d.autoral:</td>
            <td colspan="4" class="txt_normal">
			              <%
			if session("admin") = 1 then
			%>
              <strong><strong>
              <select name="cod_sub_direito" class="txt_normal" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				Set rsSub = conn.execute("SELECT CNPJ_CPF, NOME FROM GEDCLIENTES WHERE COD_GRUPO='" & rs(1) & "' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 ORDER BY NOME")

				while not rsSub.EOF
					if rsSub("CNPJ_CPF") = rs("cod_sub_direito") then
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "' selected>"& rsSub("nome") & "</option>")
						else
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "'>"& rsSub("nome") & "</option>")
					end if
					rsSub.MoveNext
				wend
				rssub.close
				%>
              </select>
              </strong></strong>
              <%
				else
					if session("admin") = 3 or session("admin") = 2 then
					%>
              <select name="cod_sub_direito" class="txt_normal" id="cod_sub_direito" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				          RESPONSE.WRITE rs("cod_sub_direito")
						  set rsSubFranqueado = Server.CreateObject("ADODB.Recordset")

						  rsSubFranqueado.open "select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Tira_Pontuacao(Request.QueryString("grupo")) &"' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 order by  nome", conn, 2, 3
						  if not rsSubFranqueado.EOF then
							while not rsSubFranqueado.EOF
							if rsSubFranqueado("cnpj_cpf") = rs("cod_sub_direito") then
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "' selected>" & rsSubFranqueado("nome") & "</option>")
								else
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "'>" & rsSubFranqueado("nome") & "</option>")
							end if
							rsSubFranqueado.MoveNext
							wend
						end if
						rssubfranqueado.close
						%>
              </select>
              <%
				end if

				if session("admin") = 4 or session("admin") = 0 or session("admin") = 8 or session("admin") = 5 or session("admin") = 6 then
					Response.Write ExibeNome(rs("cod_sub_direito"), "subfranqueado2")
					%>
              <input name="cod_sub_direito" type="hidden" id="cod_sub_direito" value="<%=rs("cod_sub_direito")%>">
            <%
				end if
		end if
				%>			</td>
          </tr>

          <tr>
            <td class="txt_normal"> JURÍDICO:</td>
            <td colspan="4" class="txt_normal">
			              <%
			if session("admin") = 1 then
			%>
              <strong><strong>
              <select name="cod_sub_juridico" class="txt_normal" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				Set rsSub = conn.execute("SELECT CNPJ_CPF, NOME FROM GEDCLIENTES WHERE COD_GRUPO='" & rs(1) & "' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 ORDER BY NOME")

				while not rsSub.EOF
					if rsSub("CNPJ_CPF") = rs("cod_sub_juridico") then
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "' selected>"& rsSub("nome") & "</option>")
					else
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "'>"& rsSub("nome") & "</option>")
					end if
					rsSub.MoveNext
				wend
				rssub.close
				%>
              </select>
              </strong></strong>
              <%
				else
					if session("admin") = 3 or session("admin") = 2 then
					%>
              <select name="cod_sub_juridico" class="txt_normal" id="cod_sub_juridico" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				          RESPONSE.WRITE rs("cod_sub_juridico")
						  set rsSubFranqueado = Server.CreateObject("ADODB.Recordset")

						  rsSubFranqueado.open "select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Tira_Pontuacao(Request.QueryString("grupo")) &"' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 order by  nome", conn, 2, 3
						  if not rsSubFranqueado.EOF then
							while not rsSubFranqueado.EOF
							if rsSubFranqueado("cnpj_cpf") = rs("cod_sub_juridico") then
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "' selected>" & rsSubFranqueado("nome") & "</option>")
								else
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "'>" & rsSubFranqueado("nome") & "</option>")
							end if
							rsSubFranqueado.MoveNext
							wend
						end if
						rssubfranqueado.close
						%>
              </select>
              <%
				end if

				if session("admin") = 4 or session("admin") = 0 or session("admin") = 8 or session("admin") = 5 or session("admin") = 6 then
					Response.Write ExibeNome(rs("cod_sub_juridico"), "subfranqueado2")
					%>
              <input name="cod_sub_juridico" type="hidden" id="cod_sub_direito" value="<%=rs("cod_sub_juridico")%>">
            <%
				end if
		end if
				%>			</td>
          </tr>

          <tr>
            <td  class="txt_normal">ASSISTENTE:</td>
            <td colspan="4" class="txt_normal">

									<select name="cod_ass" class="txt_normal" style="width:544;" id="subfranqueado2">
						  <option value="" selected>--- SELECIONE ---</option>
						  <%
						  set rsSubFranqueado = Server.CreateObject("ADODB.Recordset")

						  rsSubFranqueado.open "select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Session("franquia") &"' and TIP_CLI <> 0 and TIP_CLI <> 2 and STATUS = 1 and TIPO_CAD=0 order by  nome", conn
						  if not rsSubFranqueado.EOF then
							while not rsSubFranqueado.EOF

							if rsSubFranqueado("cnpj_cpf") = rs("cod_ass") then
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "' selected>" & rsSubFranqueado("nome") & "</option>")
								else
								response.Write("<option value='" & rsSubFranqueado("cnpj_cpf") & "'>" & rsSubFranqueado("nome") & "</option>")
							end if
							rsSubFranqueado.MoveNext
							wend
						end if
						rssubfranqueado.close
						set rssubfranqueado = nothing
						%>
                </select></td>
          </tr>
          <tr>
            <td class="txt_normal"> COBRADOR:</td>
            <td colspan="4" class="txt_normal"><%
			  if session("admin") = 1 then %>
              <select name="cod_sub_cob" class="txt_normal" style="width:544;">
                <option>--- SELECIONE ---</option>
                <%
				Set rsSub = conn.execute("select cnpj_cpf, nome from gedclientes where COD_GRUPO  ='" & Session("franquia")&"' and TIP_CLI = 8 AND status='1' and TIPO_CAD=0 order by nome")

				while not rsSub.EOF
					if rsSub("CNPJ_CPF") = rs("cod_sub_cob") then
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "' selected>"& rsSub("nome") & "</option>")
						else
						Response.Write("<option value='" & rsSub("CNPJ_CPF") & "'>"& rsSub("nome") & "</option>")
					end if
					rsSub.MoveNext
				wend
				rssub.close
				%>
              </select>
              <%
				else
					Response.Write ExibeNome(rs("cod_sub_cob"), "subfranqueado2")%>
              		<input name="cod_sub_cob" type="hidden" id="cod_sub_cob" value="<%=rs("cod_sub_cob")%>">
             <% end if %>
             </td>
          </tr>
          <tr>
            <td class="txt_normal"><strong>CLIENTE:</strong></td>
            <td colspan="4" class="txt_normal"><input name="nome" type="text" class="txt_normal" id="nome" value="<%=rs("nome")%>" size="88" maxlength="200"></td>
          </tr>
          <tr>
            <td class="txt_normal">Nome Fantasia:</td>
            <td colspan="4" class="txt_normal"><input name="nome_fantasia" type="text" class="txt_normal" id="nome_fantasia" value="<%=rs("nome_fantasia")%>" size="88" maxlength="75"></td>
          </tr>
          <tr>
            <td class="txt_normal"><strong>Fund/Nasc:</strong></td>
            <td colspan="2" class="txt_normal"><input name="data_nasc" type="text" class="txt_normal" id="Data de Nascimento" onKeyPress="return txtBoxFormat(this, '##/##/####', event);" value="<%=date_br(rs("DATA_NASC"))%>" size="12" maxlength="10"></td>
            <td class="txt_normal"><div align="right">Sexo:</div></td>
            <td class="txt_normal">
			<select name="sexo" class="txt_normal" id="select4" style="width:130;" >
                <option value="">--- SELECIONE ---</option>
                <option value="0" <% if rs("sexo") = "0" then response.Write "selected" %>>Masculino</option>
                <option value="1" <% if rs("sexo") = "1" then response.Write "selected" %>>Feminino</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="5" class="txt_normal">USA A atual raz&atilde;o social desde a funda&ccedil;&atilde;o?
              <input type="radio" name="razao_fund" id="razao_fund" value="0" onClick="mostrarDataRazao(0);" <% if rs("razao1") = "0" then %> checked <% end if %>> Sim &nbsp; <input type="radio" name="razao_fund" id="razao_fund" value="1" onClick="mostrarDataRazao(1);" <% if rs("razao1") = "1" then %> checked <% end if %>>
              N&atilde;o
            <span id="tr_data_razao" style="visibility:hidden">
            &nbsp;&nbsp;&nbsp; - INFORME A NOVA DATA:
            <input name="data_razao" type="text" class="txt_normal" id="data_razao" onKeyPress="return txtBoxFormat(this, '##/##/####', event);" value="<%=date_br(rs("razao2"))%>" size="12" maxlength="10">
            </span>            </td>
          </tr>

          <tr>
            <td class="txt_normal"><strong>gaveta/pasta:</strong></td>
            <td colspan="4" class="aviso"><input name="gaveta_pasta" type="text" class="txt_normal" id="Gaveta/Pasta" value="<%=rs("gaveta_pasta")%>" size="15" maxlength="10">
            * G999/P9999 </td>
          </tr>
          <tr>
            <td class="txt_normal"><strong>TIPO   cliente:</strong></td>
            <td colspan="4" class="txt_normal">
			<% if rs(4) = 1 then %>
			<input name="status_marca" type="checkbox" id="status_marca" value="1" checked>
			<% else %>
			<input name="status_marca" type="checkbox" id="status_marca" value="1">
			<% end if %>

 marca &nbsp;

			<% if rs(3) = 1 then %>
                <input name="status_patente" type="checkbox" id="status_patente" value="1" checked>
			<% else %>
                <input name="status_patente" type="checkbox" id="status_patente" value="1">
			<% end if
			%>




 patente

 			<% if rs(5) = 1 then %>
                <input name="status_avaliacao" type="checkbox" id="status_avaliacao" value="1" checked>
			<% else %>
                <input name="status_avaliacao" type="checkbox" id="status_avaliacao" value="1">
			<% end if
			%>
avaliação

 			<% if rs(6) = 1 then %>
                <input name="status_software" type="checkbox" id="status_software" value="1" checked>
			<% else %>
                <input name="status_software" type="checkbox" id="status_software" value="1">
			<% end if
			%>
software

 			<% if rs(7) = 1 then %>
                <input name="status_internacional" type="checkbox" id="status_internacional" value="1" checked>
			<% else %>
                <input name="status_internacional" type="checkbox" id="status_internacional" value="1">
			<% end if
			%>
Internacional

 			<% if rs(8) = 1 then %>
                <input name="status_direito" type="checkbox" id="status_direito" value="1" checked>
			<% else %>
                <input name="status_direito" type="checkbox" id="status_direito" value="1">
			<% end if
			%>
D.Autoral

 			<% if rs(9) = 1 then %>
                <input name="status_juridico" type="checkbox" id="status_juridico" value="1" checked>
			<% else %>
                <input name="status_juridico" type="checkbox" id="status_juridico" value="1">
			<% end if
			%>
JURÍDICO</td>
          </tr>
          <tr>
            <td class="txt_normal">grupo cliente:</td>
            <td colspan="4" class="txt_normal"><select name="grupo_cliente" class="txt_normal" id="Grupo Cliente" >
              <option value="">--- SELECIONE ---</option>
              <%
			set rsGrupo_Cliente = Server.CreateObject("ADODB.Recordset")

			rsGrupo_Cliente.open "Select * from gedfingrupo where cod_grupo = '" & session("franquia") & "' ORDER BY NOME", conn, 3, 3

			while not rsGrupo_Cliente.EOF
			    if rs("grupo_fin") = rsgrupo_cliente("codigo") then
				   Response.Write("<option value='" & rsGrupo_Cliente("codigo") & "' selected='selected' >" & ucase(rsGrupo_Cliente("nome")) & "</option>")
				else
				   Response.Write("<option value='" & rsGrupo_Cliente("codigo") & "'>" & ucase(rsGrupo_Cliente("nome")) & "</option>")
				end if
				rsGrupo_Cliente.MoveNext
			Wend
			rsGrupo_Cliente.close
			Set rsGrupo_Cliente = nothing
			%>
            </select></td>
          </tr>
          <%
		  wvalor = rs("capital")&""
		  if (trim(wvalor)="") then
		  	  wcapital = "0"
		  else
		  	  wcapital = wvalor
		  end if
		 %>
          <tr>
            <td colspan="9" class="txt_normal">Capital Social Integralizado (R$):
               <input name="capital" type="text" class="txt_normal" id="capital" size="18" maxlength="20" value="<%=wcapital%>">            </td>
          </tr>
          <tr bgcolor="#CCFFCC">
            <td height="15" colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center"><strong>ENDERE&Ccedil;OS</strong></div></td>
          </tr>
          <tr>
            <td colspan="5" class="txt_normal"><div class="tab-page" id="modules-cpanel" style="width:96%;"><script type="text/javascript">
   var tabPane1 = new WebFXTabPane( document.getElementById( "modules-cpanel" ), 1 )
</script>
	  <div class="tab-page" id="module20" style="width:100%;">
	    <h2 class="tab">COMERCIAL </h2>
	    <script type="text/javascript">tabPane1.addTabPage(document.getElementById("module20"));</script><table width="100%" border="0" class="txt_normal">
          <tr>
            <td width="15%" class="txt_normal"><strong>Endere&ccedil;o:</strong></td>
            <td><input name="endereco" type="text" class="txt_normal" id="Endereco" value="<%=rs("endereco")%>" size="50" maxlength="200">
              N&ordm;:
<input name="numero" type="text" class="txt_normal" id="numero" value="<%=rs("numero")%>" size="5" maxlength="10"></td>
            <td colspan="2" rowspan="2" valign="top"><img src="../imagens/logo_correios.jpg" alt="Busca de CEP apartir do endere&ccedil;o." width="80" height="18" border="0" style="cursor:pointer;" onClick="javascript:window.open('http://www.buscacep.correios.com.br/sistemas/buscacep/buscaCepEndereco.cfm');"></td>
            </tr>
			<!--https://www.registrafacil.com.br/registrafacil/utilitario/busca_cep_java.asp?endereco='+document.frm.endereco.value+'&numero='+document.frm.numero.value+'&estado='+document.frm.txt_estado.value+'&cidade='+document.frm.txt_cidade.value,'_blank','width=350,height=350   -->
          <tr>
            <td class="txt_normal">Complemento:</td>
            <td><input name="complemento" type="text" class="txt_normal" id="complemento" value="<%=rs("complemento")%>" size="50" maxlength="40"></td>
            </tr>
          <tr>
            <td class="txt_normal"><strong>Bairro:</strong></td>
            <td width="53%" class="txt_normal"><input name="bairro" type="text" class="txt_normal" id="Bairro" value="<%=rs("bairro")%>" size="30"></td>
            <td width="7%" class="txt_normal"><div align="right"><strong>CEP:</strong></div></td>
            <td width="29%" class="txt_normal" valign="middle"><input name="cep" type="text" class="txt_normal" id="CEP" onKeyPress="return txtBoxFormat(this, '#####-###', event);" value="<%=rs("cep")%>" size="10" maxlength="9"></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="left"><strong>Estado:</strong></div></td>
            <td class="txt_normal"><select name="txt_estado" style="width:150;" class="txt_normal" id="txt_estado" onChange="ajaxGet('../include/carrega_cidades.asp?uf='+this.value,document.getElementById('cidades'),true);">
                <option value="">--- SELECIONE ---</option>
                <%
			set rsEstados = Server.CreateObject("ADODB.Recordset")

			rsEstados.open "Select * from UF", conn, 3, 3

			while not rsEstados.EOF
				if rsEstados("txt_uf") = rs("estado") then
					Response.Write("<option value='" & rsEstados("txt_uf") & "' selected>" & rsEstados("txt_estado") & "</option>")
				else
					Response.Write("<option value='" & rsEstados("txt_uf") & "'>" & rsEstados("txt_estado") & "</option>")
				end if
				rsEstados.MoveNext
			Wend
			rsEstados.close
			Set rsEstados = nothing
			%>
            </select></td>
            <td class="txt_normal">&nbsp;</td>
            <td class="txt_normal"><img src="../imagens/logo_correios.jpg" alt="Busca de CEP apartir do CEP." width="80" height="18" border="0" onClick="javascript:window.open('https://www.registrafacil.com.br/registrafacil/utilitario/busca_por_cep_java.asp?cep='+document.frm.cep.value,'_blank','width=350,height=280')" style="cursor:pointer;"></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="left"><strong>Cidade:</strong></div></td>
            <td class="txt_normal"><div align="left" id="cidades">
                <select name="txt_cidade" class="txt_normal" id="txt_cidade" style="width:150;">
                  <option selected><%=rs("cidade") %></option>
                  <%
			set rsMunicipios = Server.CreateObject("ADODB.Recordset")

			rsMunicipios.open "Select * from MUNICIPIOS WHERE TXT_UF='" & rs("estado") & "'", conn, 3, 3

			while not rsMunicipios.EOF
				if rsMunicipios("txt_municipio") = rs("cidade") then
					Response.Write("<option value='" & rsMunicipios("txt_municipio") & "' selected>" & rsMunicipios("txt_municipio") & "</option>")
				else
					Response.Write("<option value='" & rsMunicipios("txt_municipio") & "'>" & rsMunicipios("txt_municipio") & "</option>")
				end if
				rsMunicipios.MoveNext
			Wend
			rsMunicipios.close
			Set rsMunicipios = nothing
			%>
                </select>
            </div></td>
            <td class="txt_normal">&nbsp;</td>
            <td class="txt_normal">&nbsp;</td>
          </tr>
          <tr>
            <td class="txt_normal"><strong>Pa&iacute;s:</strong></td>
            <td class="txt_normal"><select name="pais" class="txt_normal" id="pais" style="width:150;">
              <option value="">SELECIONE</option>
              <%
			set rsPais = Server.CreateObject("ADODB.Recordset")
			rsPais.open "Select * from pais order by cnome", conn, 3, 3

			while not rsPais.EOF
				If UCASE(rsPais("cnome")) = Ucase(rs("pais")) then
					Response.Write ("<option value='" & rsPais("cnome") & "' selected>"& rsPais("cnome") & "</option>")
				Else
					Response.Write ("<option value='" & rsPais("cnome") & "'>"& rsPais("cnome") & "</option>")
				End if
				rsPais.MoveNext
			wend
			rsPais.close
			Set rsPais = nothing
			%>
            </select></td>
            <td class="txt_normal">&nbsp;</td>
            <td class="txt_normal">&nbsp;</td>
          </tr>

         </table>
	  </div>
	  <div class="tab-page" id="module21" style="width:100%;">
	  <h2 class="tab">INPI </h2>
	  <script type="text/javascript">tabPane1.addTabPage(document.getElementById("module21"));</script><table width="100%" border="0" class="txt_normal">
        <tr>
          <td width="15%" class="txt_normal">Endere&ccedil;o:</td>
          <td><input name="endereco_inpi" type="text" class="txt_normal" id="endereco2" value="<%=rs("endereco_inpi")%>" size="50" maxlength="200" />
            N&ordm;:
            <input name="numero_inpi" type="text" class="caixa_mensagem" id="numero2" value="<%=rs("numero_inpi")%>" size="5" maxlength="10" />&nbsp;<img src="../imagens/all_arrow_back_red.gif" alt="Copiar Dados" width="14" height="12" onClick="javascript:copia_endereco(1)" style="cursor:pointer;"></td>
          <td colspan="2"><img src="../imagens/logo_correios.jpg" alt="Busca de CEP apartir do endere&ccedil;o." width="80" height="18" border="0" style="cursor:pointer;" onClick="javascript:window.open('https://www.registrafacil.com.br/registrafacil/utilitario/busca_cep_java.asp?endereco='+document.frm.endereco2.value+'&numero='+document.frm.numero_inpi.value+'&estado='+document.frm.txt_estado_inpi.value+'&cidade='+document.frm.txt_cidade_inpi.value,'_blank','width=350,height=350');">&nbsp;</td>
          </tr>
        <tr>
          <td class="txt_normal">Complemento:</td>
          <td colspan="3"><input name="complemento_inpi" type="text" class="txt_normal" id="complemento2" value="<%=rs("complemento_inpi")%>" size="50" maxlength="40" /></td>
        </tr>
        <tr>
          <td class="txt_normal">Bairro:</td>
          <td width="53%"><input name="bairro_inpi" type="text" class="txt_normal" id="bairro2" value="<%=rs("bairro_inpi")%>" size="30" /></td>
          <td width="7%"><div align="right">CEP:</div></td>
          <td width="29%"><input name="cep_inpi" type="text" class="txt_normal" id="cep2" onKeyPress="return txtBoxFormat(this, '99.999-999', event);" value="<%=rs("cep_inpi")%>" size="12" maxlength="10" /></td>
        </tr>
        <tr>
          <td class="txt_normal"><div align="left">Estado:</div></td>
          <td><select name="txt_estado_inpi" class="txt_normal" id="txt_estado2" style="width:150;" onChange="ajaxGet('../include/carrega_cidades2.asp?comp=_inpi&uf='+this.value,document.getElementById('cidades_inpi'),true);">
              <option value="">--- SELECIONE ---</option>
               <%
			set rsEstados = Server.CreateObject("ADODB.Recordset")

			rsEstados.open "Select * from UF", conn, 3, 3

			while not rsEstados.EOF
				if rsEstados("txt_uf") = rs("estado_inpi") then
					Response.Write("<option value='" & rsEstados("txt_uf") & "' selected>" & rsEstados("txt_estado") & "</option>")
				else
					Response.Write("<option value='" & rsEstados("txt_uf") & "'>" & rsEstados("txt_estado") & "</option>")
				end if
				rsEstados.MoveNext
			Wend
			rsEstados.close
			Set rsEstados = nothing
			%>
          </select></td>
          <td>&nbsp;</td>
          <td><img src="../imagens/logo_correios.jpg" alt="Busca de CEP apartir do CEP." width="80" height="18" border="0" onClick="javascript:window.open('https://www.registrafacil.com.br/registrafacil/utilitario/busca_por_cep_java.asp?cep='+document.getElementById('cep2').value,'_blank','width=350,height=280')" style="cursor:pointer;"></td>
        </tr>
        <tr>
          <td class="txt_normal"><div align="left">Cidade:</div></td>
          <td colspan="3"><div id="cidades_inpi">
            <select name="txt_cidade_inpi" class="txt_normal"  style="width:150;" id="txt_cidade2">
              <option value="" selected>
                <% = rs("cidade_inpi") %>
                </option>
              <%
			set rsMunicipios = Server.CreateObject("ADODB.Recordset")

			rsMunicipios.open "Select * from MUNICIPIOS WHERE TXT_UF='" & rs("estado") & "'", conn, 3, 3

			while not rsMunicipios.EOF
				if rsMunicipios("txt_municipio") = rs("cidade_inpi") then
					Response.Write("<option value='" & rsMunicipios("txt_municipio") & "' selected>" & rsMunicipios("txt_municipio") & "</option>")
				else
					Response.Write("<option value='" & rsMunicipios("txt_municipio") & "'>" & rsMunicipios("txt_municipio") & "</option>")
				end if
				rsMunicipios.MoveNext
			Wend
			rsMunicipios.close
			Set rsMunicipios = nothing
			%>
            </select>
          </div></td>
        </tr>
        <tr>
            <td class="txt_normal">Pa&iacute;s:</td>
            <td class="txt_normal"><select name="pais2" class="txt_normal" id="pais2" style="width:150;">
              <option>SELECIONE</option>
              <%
			set rsPais = Server.CreateObject("ADODB.Recordset")
			rsPais.open "Select * from pais order by cnome", conn, 3, 3

			while not rsPais.EOF
				If UCASE(rsPais("cnome")) = Ucase(rs("pais")) then
					Response.Write ("<option value='" & rsPais("cnome") & "' selected>"& rsPais("cnome") & "</option>")
				Else
					Response.Write ("<option value='" & rsPais("cnome") & "'>"& rsPais("cnome") & "</option>")
				End if
				rsPais.MoveNext
			wend
			rsPais.close
			Set rsPais = nothing
			%>
            </select></td>
            <td class="txt_normal">&nbsp;</td>
            <td class="txt_normal">&nbsp;</td>
          </tr>
        </table>
	  </div>
	  <div class="tab-page" id="module22" style="width:100%;">
		  <h2 class="tab">COBRAN&Ccedil;A</h2>
		  <script type="text/javascript">tabPane1.addTabPage(document.getElementById("module22"));</script>
		  <table width="100%" border="0" class="txt_normal">
			<tr>
			  <td width="15%" class="txt_normal">Endere&ccedil;o:</td>
			  <td><input name="endereco_corr" type="text" class="txt_normal" id="endereco3" value="<%=rs("endereco_corr")%>" size="50" maxlength="200" />
				N&ordm;:
				<input name="numero_corr" type="text" class="caixa_mensagem" id="numero3" value="<%=rs("numero_corr")%>" size="5" maxlength="10" />&nbsp;<img src="../imagens/all_arrow_back_red.gif" alt="Copiar Dados" width="14" height="12" onClick="javascript:copia_endereco(2)" style="cursor:pointer;"></td>
			  <td colspan="2"><img src="../imagens/logo_correios.jpg" alt="Busca de CEP apartir do endere&ccedil;o." width="80" height="18" border="0" style="cursor:pointer;" onClick="javascript:window.open('https://www.registrafacil.com.br/registrafacil/utilitario/busca_cep_java.asp?endereco='+document.frm.endereco_corr.value+'&numero='+document.frm.numero_corr.value+'&estado='+document.frm.txt_estado_corr.value+'&cidade='+document.frm.txt_cidade_corr.value,'_blank','width=350,height=350');"></td>
			  </tr>
			<tr>
			  <td class="txt_normal">Complemento:</td>
			  <td colspan="3"><input name="complemento_corr" type="text" class="txt_normal" id="complemento3" value="<%=rs("complemento_corr")%>" size="50" maxlength="40" /></td>
			</tr>
			<tr>
			  <td class="txt_normal">Bairro:</td>
			  <td width="53%"><input name="bairro_corr" type="text" class="txt_normal" id="bairro3" value="<%=rs("bairro_corr")%>" size="30" /></td>
			  <td width="7%"><div align="right">CEP:</div></td>
			  <td width="29%"><input name="cep_corr" type="text" class="txt_normal" id="cep3" onKeyPress="return txtBoxFormat(this, '99.999-999', event);" value="<%=rs("cep_corr")%>" size="12" maxlength="10" /></td>
			</tr>
			<tr>
			  <td class="txt_normal"><div align="left">Estado:</div></td>
			  <td><select name="txt_estado_corr" style="width:150;" class="txt_normal" id="txt_estado3" onChange="ajaxGet('../include/carrega_cidades3.asp?comp=_corr&uf='+this.value,document.getElementById('cidades_corr'),true);">
				  <option value="" selected="selected">--- SELECIONE ---</option>
				   <%
			set rsEstados = Server.CreateObject("ADODB.Recordset")

			rsEstados.open "Select * from UF", conn, 3, 3

			while not rsEstados.EOF
				if rsEstados("txt_uf") = rs("estado_corr") then
					Response.Write("<option value='" & rsEstados("txt_uf") & "' selected>" & rsEstados("txt_estado") & "</option>")
				else
					Response.Write("<option value='" & rsEstados("txt_uf") & "'>" & rsEstados("txt_estado") & "</option>")
				end if
				rsEstados.MoveNext
			Wend
			rsEstados.close
			Set rsEstados = nothing
			%>
			  </select></td>
			  <td rowspan="2"><div align="center"></div></td>
			  <td><img src="../imagens/logo_correios.jpg" alt="Busca de CEP apartir do CEP." width="80" height="18" border="0" onClick="javascript:window.open('https://www.registrafacil.com.br/registrafacil/utilitario/busca_por_cep_java.asp?cep='+document.frm.cep_corr.value,'_blank','width=350,height=280')" style="cursor:pointer;"></td>
			</tr>
			<tr>
			  <td class="txt_normal"><div align="left">Cidade:</div></td>
			  <td><div id="cidades_corr">
				  <select name="txt_cidade_corr" class="txt_normal" style="width:150;" id="txt_cidade3">
					<option value=""><%= rs("cidade_corr") %></option>
					 <%
			set rsMunicipios = Server.CreateObject("ADODB.Recordset")

			rsMunicipios.open "Select * from MUNICIPIOS WHERE TXT_UF='" & rs("estado") & "'", conn, 3, 3

			while not rsMunicipios.EOF
				if rsMunicipios("txt_municipio") = rs("cidade_corr") then
					Response.Write("<option value='" & rsMunicipios("txt_municipio") & "' selected>" & rsMunicipios("txt_municipio") & "</option>")
				else
					Response.Write("<option value='" & rsMunicipios("txt_municipio") & "'>" & rsMunicipios("txt_municipio") & "</option>")
				end if
				rsMunicipios.MoveNext
			Wend
			rsMunicipios.close
			Set rsMunicipios = nothing
			%>
				  </select>
			  </div></td>
			  <td>&nbsp;</td>
			</tr>
            <tr>
            <td class="txt_normal">Pa&iacute;s:</td>
            <td class="txt_normal"><select name="pais3" class="txt_normal" id="pais3" style="width:150;">
              <option>SELECIONE</option>
              <%
			set rsPais = Server.CreateObject("ADODB.Recordset")
			rsPais.open "Select * from pais order by cnome", conn, 3, 3

			while not rsPais.EOF
				If UCASE(rsPais("cnome")) = Ucase(rs("pais")) then
					Response.Write ("<option value='" & rsPais("cnome") & "' selected>"& rsPais("cnome") & "</option>")
				Else
					Response.Write ("<option value='" & rsPais("cnome") & "'>"& rsPais("cnome") & "</option>")
				End if
				rsPais.MoveNext
			wend
			rsPais.close
			Set rsPais = nothing
			%>
            </select></td>
            <td class="txt_normal">&nbsp;</td>
            <td class="txt_normal">&nbsp;</td>
          </tr>
			</table>
		</div>
	</div></td>
          </tr>
          <tr>
            <td class="txt_normal">&nbsp;</td>
            <td colspan="4" class="txt_normal">&nbsp;</td>
          </tr>
          <tr bgcolor="#CCFFCC">
            <td height="15" colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center"><strong>TELEFONES</strong></div></td>
          </tr>
          <tr>
            <td class="txt_normal"> <div align="right"></div></td>
            <td width="125" class="txt_normal"><div align="right">Residencial: </div></td>
            <td width="222" class="txt_normal">
              <div align="left">
                <input name="fone_res" type="text" class="caixa_mensagem" id="fone_res" onKeyPress="return txtBoxFormat(this, '(99)9999-9999', event);" value="<%=rs("fone_res")%>" size="16" maxlength="13">
            </div></td>
            <td class="txt_normal">
              <div align="right">Comercial: </div></td>
            <td class="txt_normal"><input name="fone_com" type="text" class="caixa_mensagem" id="fone_com" onKeyPress="return txtBoxFormat(this, '(99)9999-9999', event);" value="<%=rs("fone_com")%>" size="16" maxlength="13"></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="right"></div></td>
            <td class="txt_normal"><div align="right">Fax:</div></td>
            <td class="txt_normal">
              <div align="left">
                <input name="fax" type="text" class="caixa_mensagem" id="fone_res3" onKeyPress="return txtBoxFormat(this, '(99)9999-9999', event);" value="<%=rs("fax")%>" size="16" maxlength="13">
              </div></td>
            <td class="txt_normal">
              <div align="right">Celular:</div></td>
            <td class="txt_normal"><input name="celular" type="text" class="caixa_mensagem" id="celular" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<%=rs("celular")%>" size="16" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal">&nbsp;</td>
            <td class="txt_normal"><div align="right">r&aacute;dio:</div></td>
            <td class="txt_normal"><input name="radio" type="text" class="caixa_mensagem" id="radio"  value="<%=rs("radio")%>" size="16" maxlength="50"></td>
            <td class="txt_normal">&nbsp;</td>
            <td class="txt_normal">&nbsp;</td>
          </tr>
          <tr>
            <td colspan="5" ><span class="aviso">* pelo menos um telefone &eacute; obrigat&oacute;rio </span></td>
          </tr>
          <tr bgcolor="#CCFFCC">
            <td height="15" colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center"><strong>COMUNICA&Ccedil;&Atilde;O ELETR&Ocirc;NICA</strong></div></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="right"><strong>Email 1: </strong></div></td>
            <td colspan="4" class="txt_normal"><input name="email1" type="text" class="caixa_mensagem" id="Email" value="<%=rs("email1")%>" size="90" maxlength="200" onBlur="validarEmail(this.value, this.name);"></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="right">Email 2: </div></td>
            <td colspan="4" class="txt_normal"><input name="email2" type="text" class="caixa_mensagem" id="email2" value="<%=rs("email2")%>" size="90" maxlength="200" onBlur="validarEmail(this.value, this.name);"></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="right">msn: </div></td>
            <td colspan="4" class="txt_normal"><input name="email_msn" type="text" class="caixa_mensagem" id="email_msn" value="<%=rs("email_msn")%>" size="50" maxlength="150"></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="right">skype: </div></td>
            <td colspan="4" class="txt_normal"><input name="skype" type="text" class="caixa_mensagem" id="skype" value="<%=rs("skype")%>" size="50" maxlength="50"></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="right">orkut: </div></td>
            <td colspan="4" class="txt_normal"><input name="orkut" type="text" class="caixa_mensagem" id="orkut" value="<%=rs("orkut")%>" size="50" maxlength="50"></td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="right">Site:</div></td>
            <td colspan="4" class="txt_normal"><input name="site" type="text" class="caixa_mensagem" id="site" value="<%=rs("site")%>" size="90" maxlength="75"></td>
          </tr>
          <tr bgcolor="#CCFFCC">
            <td height="15" colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center"><strong>CONTATOS</strong></div></td>
          </tr>
          <tr>
            <td width="106" class="txt_normal"><div align="left">NOME:</div></td>
            <td colspan="4" class="txt_normal"><div align="left"><strong>
              <input name="contato1" type="text" class="caixa_mensagem" id="CONTATO" value="<% = rs("contato1") %>" size="70" maxlength="50">
            </strong></div></td>
          </tr>
          <tr>
            <td class="txt_normal">CARGO/SETOR:</td>
            <td colspan="3" class="txt_normal"><strong>
              <input name="cargo1" type="text" class="caixa_mensagem" id="cargo1" value="<% = rs("cargo1") %>" size="40">
            </strong></td>
            <td class="txt_normal">ANIV.:
              <input name="aniver1" type="text" class="caixa_mensagem" id="aniver1" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<% = rs("aniver1") %>" size="5" maxlength="5">
              <span class="aviso">(dd/mm)</span></td>
          </tr>
          <tr>
            <td class="txt_normal">EMAIL:</td>
            <td colspan="3" class="txt_normal"><strong>
              <input name="email_contato1" type="text" class="caixa_mensagem" id="email_contato1" value="<% = rs("email_contato1") %>" size="50" maxlength="50">
            </strong></td>
            <td class="txt_normal">TELEFONE1:
            <input name="fone_contato1" type="text" class="caixa_mensagem" id="fone_contato1" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<%=rs("fone_contato1") %>" size="15" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal" >&nbsp;</td>
            <td class="txt_normal" colspan="3">&nbsp;</td>
            <td class="txt_normal" colspan="3">TELEFONE2:
            <input name="fone2_contato1" type="text" class="caixa_mensagem" id="fone2_contato1" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<%=rs("fone2_contato1")%>" size="16" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal">&nbsp;</td>
            <td colspan="4" class="txt_normal">&nbsp;</td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="left">NOME:</div></td>
            <td colspan="4" class="txt_normal"><div align="left">
              <input name="contato2" type="text" class="caixa_mensagem" id="site23" value="<% = rs("contato2") %>" size="70" maxlength="50">
            </div></td>
          </tr>
          <tr>
            <td class="txt_normal">CARGO/setor:</td>
            <td colspan="3" class="txt_normal"><input name="cargo2" type="text" class="caixa_mensagem" id="cargo2" value="<% = rs("cargo2") %>" size="40" maxlength="30"></td>
            <td class="txt_normal"> <div align="left">ANIV.:<span class="aviso">
              <input name="aniver2" type="text" class="caixa_mensagem" id="aniver2" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<% = rs("aniver2") %>" size="8" maxlength="5">
              (dd/mm)</span></div></td></tr>
          <tr>
            <td class="txt_normal">EMAIL:</td>
            <td colspan="3" class="txt_normal"><strong>
              <input name="EMAIL_CONTATO2" type="text" class="caixa_mensagem" id="email_contato2" value="<% = rs("email_contato2") %>" size="50" maxlength="50">
            </strong></td>
            <td class="txt_normal">TELEFONE1:
              <input name="FONE_CONTATO2" type="text" class="caixa_mensagem" id="fone_contato2" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<% = rs("fone_contato2") %>" size="15" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal" >&nbsp;</td>
            <td class="txt_normal" colspan="3">&nbsp;</td>
            <td class="txt_normal" colspan="3">TELEFONE2:
            <input name="fone2_contato2" type="text" class="caixa_mensagem" id="fone2_contato2" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<%=rs("fone2_contato2")%>" size="16" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal">&nbsp;</td>
            <td colspan="4" class="txt_normal">&nbsp;</td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="left">NOME:</div></td>
            <td colspan="4" class="txt_normal"><div align="left">
              <input name="contato3" type="text" class="caixa_mensagem" id="site23" value="<% = rs("contato3") %>" size="70" maxlength="50">
            </div></td>
          </tr>
          <tr>
            <td class="txt_normal">CARGO/setor:</td>
            <td colspan="3" class="txt_normal"><input name="cargo3" type="text" class="caixa_mensagem" id="cargo3" value="<% = rs("cargo3") %>" size="40" maxlength="30"></td>
            <td class="txt_normal"> <div align="left">ANIV.:<span class="aviso">
              <input name="aniver3" type="text" class="caixa_mensagem" id="aniver3" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<% = rs("aniver3") %>" size="8" maxlength="5">
              (dd/mm)</span></div></td></tr>
          <tr>
            <td class="txt_normal">EMAIL:</td>
            <td colspan="3" class="txt_normal"><strong>
              <input name="EMAIL_CONTATO3" type="text" class="caixa_mensagem" id="email_contato3" value="<% = rs("email_contato3") %>" size="50" maxlength="50">
            </strong></td>
            <td class="txt_normal">TELEFONE1:
              <input name="FONE_CONTATO3" type="text" class="caixa_mensagem" id="fone_contato3" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<% = rs("fone_contato3") %>" size="15" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal" >&nbsp;</td>
            <td class="txt_normal" colspan="3">&nbsp;</td>
            <td class="txt_normal" colspan="3">TELEFONE2:
            <input name="fone2_contato3" type="text" class="caixa_mensagem" id="fone2_contato3" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<%=rs("fone2_contato3")%>" size="16" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal">&nbsp;</td>
            <td colspan="4" class="txt_normal">&nbsp;</td>
          </tr>
          <tr>
            <td class="txt_normal"><div align="left">NOME:</div></td>
            <td colspan="4" class="txt_normal"><div align="left">
              <input name="contato4" type="text" class="caixa_mensagem" id="site23" value="<% = rs("contato4") %>" size="70" maxlength="50">
            </div></td>
          </tr>
          <tr>
            <td class="txt_normal">CARGO/setor:</td>
            <td colspan="3" class="txt_normal"><input name="cargo4" type="text" class="caixa_mensagem" id="cargo4" value="<% = rs("cargo4") %>" size="40" maxlength="30"></td>
            <td class="txt_normal"> <div align="left">ANIV.:<span class="aviso">
              <input name="aniver4" type="text" class="caixa_mensagem" id="aniver4" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<% = rs("aniver4") %>" size="8" maxlength="5">
              (dd/mm)</span></div></td></tr>
          <tr>
            <td class="txt_normal">EMAIL:</td>
            <td colspan="3" class="txt_normal"><strong>
              <input name="EMAIL_CONTATO4" type="text" class="caixa_mensagem" id="email_contato4" value="<% = rs("email_contato4") %>" size="50" maxlength="50">
            </strong></td>
            <td class="txt_normal">TELEFONE1:
              <input name="FONE_CONTATO4" type="text" class="caixa_mensagem" id="fone_contato4" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<% = rs("fone_contato4") %>" size="15" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal" >&nbsp;</td>
            <td class="txt_normal" colspan="3">&nbsp;</td>
            <td class="txt_normal" colspan="3">TELEFONE2:
            <input name="fone2_contato4" type="text" class="caixa_mensagem" id="fone2_contato4" onKeyPress="return txtBoxFormat(this, '(99)99999-9999', event);" value="<%=rs("fone2_contato4")%>" size="16" maxlength="14"></td>
          </tr>
          <tr>
            <td class="txt_normal">&nbsp;</td>
            <td colspan="3" class="txt_normal">&nbsp;</td>
            <td class="txt_normal">&nbsp;</td>
          </tr>

		  <%
		  set rsSocios = Server.CreateObject("ADODB.Recordset")
		  rsSocios.open "SELECT * FROM GEDSOCIOS WHERE CNPJ_CPF ='" & rs("cnpj_cpf") & "'", conn, 3, 3
		  cont = 0
		  if not rsSocios.EOF then
		  	cont = cont + 1
			  socio1 = rsSocios("nome_socio")
			  aniver1 = rsSocios("aniver")
			  cargo1 = rsSocios("cargo")
			  cod1 = rsSocios("codigo")
			  profissao1 = rsSocios("profissao")
			  email_socio1 = rsSocios("email")
			  cpf_socio1 = rsSocios("cpf_socio")
			  rg1 = rsSocios("rg_socio")
			  status_socio1 = rsSocios("status_socio")
			  rsSocios.MoveNext
		  end if
		  if not rsSocios.EOF then
		  	cont = cont + 1
		  	  socio2 = rsSocios("nome_socio")
			  aniver2 = rsSocios("aniver")
			  cargo2 = rsSocios("cargo")
			  cod2 = rsSocios("codigo")
			  profissao2 = rsSocios("profissao")
			  email_socio2 = rsSocios("email")
			  cpf_socio2 = rsSocios("cpf_socio")
			  rg2 = rsSocios("rg_socio")
			  status_socio2 = rsSocios("status_socio")
		      rsSocios.MoveNext
		  end if
		  if not rsSocios.EOF then
		  	cont = cont + 1
		  	  socio3 = rsSocios("nome_socio")
			  aniver3 = rsSocios("aniver")
			  cargo3 = rsSocios("cargo")
			  cod3 = rsSocios("codigo")
			  profissao3 = rsSocios("profissao")
			  email_socio3 = rsSocios("email")
			  cpf_socio3 = rsSocios("cpf_socio")
			  rg3 = rsSocios("rg_socio")
			  status_socio3= rsSocios("status_socio")
		      rsSocios.MoveNext
		  end if
		  if not rsSocios.EOF then
		  	cont = cont + 1
		  	  socio4 = rsSocios("nome_socio")
			  aniver4 = rsSocios("aniver")
			  cargo4 = rsSocios("cargo")
			  cod4 = rsSocios("codigo")
			  profissao4 = rsSocios("profissao")
			  email_socio4 = rsSocios("email")
			  cpf_socio4 = rsSocios("cpf_socio")
			  rg4 = rsSocios("rg_socio")
			  status_socio4 = rsSocios("status_socio")
		      rsSocios.MoveNext
		  end if
		  if not rsSocios.EOF then
		  	cont = cont + 1
		  	  socio5 = rsSocios("nome_socio")
			  aniver5 = rsSocios("aniver")
			  cargo5 = rsSocios("cargo")
			  cod5 = rsSocios("codigo")
			  profissao5 = rsSocios("profissao")
			  email_socio5 = rsSocios("email")
			  cpf_socio5 = rsSocios("cpf_socio")
			  rg5 = rsSocios("rg_socio")
			  status_socio5 = rsSocios("status_socio")
		      rsSocios.MoveNext
		  end if
		  %>
          <tr bgcolor="#CCFFCC">
            <td height="15" colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center"><strong>SOCIOS / REPRESENTANTES LEGAIS</strong></div></td>
          </tr>  
          <tr>
            <td colspan="5" class="txt_normal"><table width="100%" border="0" class="txt_normal">
              <tr class="style2">
                <td colspan="4"><div align="center"><strong>N&ordm; 1</strong> <span class="aviso">*</span> </div></td>
                </tr>
              <tr>
                <td width="16%">Nome:</td>
                <td colspan="3"><input name="representante1" type="text" class="caixa_mensagem" id="representante1" value="<%=socio1%>" size="90" maxlength="90">
                  <input name="codigo1" type="hidden" id="codigo1" value="<%=cod1%>">
                  <img src="../imagens/all_arrow_back_red.gif" alt="Copiar Dados" width="11" height="9" onClick="javascript:copia_dados()"></td>
              </tr>
              <tr>
                <td>CARGO:</td>
                <td><input name="cargo_representante1" type="text" class="caixa_mensagem" id="cargo_representante1" value="<%=cargo1%>" size="40" maxlength="75"></td>
                <td><div align="right">Aniv.:</div></td>
                <td width="44%"><span class="aviso">
                  <input name="aniver_representante1" type="text" class="caixa_mensagem" id="aniver_representante1" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<%=aniver1%>" size="5" maxlength="5">
                  (dd/mm)</span></td>
                </tr>
              <tr>
                <td>Profiss&atilde;o:</td>
                <td width="33%"><input name="prof_representante1" type="text" class="caixa_mensagem" id="prof_representante1" value="<%=profissao1%>" size="40"></td>
                <td width="7%"><div align="right">cnpj/CPF:</div></td>
                <td><input name="cpf_representante1" type="text" class="caixa_mensagem" id="cpf_representante1"  value="<%=FormataCNPJCPF(cpf_socio1)%>" size="22" maxlength="20"></td>
                </tr>
              <tr>
                <td>Email:</td>
                <td><input name="email_representante1" type="text" class="caixa_mensagem" id="email_representante1" value="<%=email_socio1%>" size="40" maxlength="90"></td>
                <td><div align="right">RG/insc:</div></td>
                <td><input name="rg_representante1" type="text" class="caixa_mensagem" id="rg_representante1" value="<%=rg1%>" size="16" maxlength="75"></td>
                </tr>
              <tr>
                <td>contrato:</td>
                <td><input name="status_socio1" type="hidden" value="0"> Sim</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td colspan="5" class="txt_normal"><div align="center">
              <table width="100%" border="0" class="txt_normal">
                <tr class="style2">
                  <td colspan="4" class="style2"><div align="center"><strong>N&ordm; 2 </strong> </div></td>
                  </tr>
                <tr>
                  <td width="16%">Nome:</td>
                  <td colspan="3"><input name="representante2" type="text" class="caixa_mensagem" id="representante2" value="<%=socio2%>" size="90" maxlength="90">
                    <input name="codigo2" type="hidden" id="codigo2" value="<%=cod2%>"></td>
                </tr>
                <tr>
                  <td>Cargo:</td>
                  <td width="33%"><input name="cargo_representante2" type="text" class="caixa_mensagem" id="cargo_representante2" value="<%=cargo2%>" size="40" maxlength="75"></td>
                  <td width="7%"><div align="right">Aniv.:</div></td>
                  <td width="44%"><input name="aniver_representante2" type="text" class="caixa_mensagem" id="aniver_representante2" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<%=aniver2%>" size="5" maxlength="5">
                    <span class="aviso">(dd/mm)</span></td>
                  </tr>
                <tr>
                  <td>Profiss&atilde;o:</td>
                  <td><input name="prof_representante2" type="text" class="caixa_mensagem" id="prof_representante2" value="<%=profissao2%>" size="40"></td>
                  <td><div align="right">cnpj/CPF:</div></td>
                  <td><input name="cpf_representante2" type="text" class="caixa_mensagem" id="cpf_representante2"  value="<%=FormataCNPJCPF(cpf_socio2)%>" size="22" maxlength="20"></td>
                  </tr>
                <tr>
                  <td>Email:</td>
                  <td><input name="email_representante2" type="text" class="caixa_mensagem" id="email_representante2" value="<%=email_socio2%>" size="40" maxlength="90"></td>
                  <td><div align="right">RG/insc:</div></td>
                  <td><input name="rg_representante2" type="text" class="caixa_mensagem" id="rg_representante2" value="<%=rg2%>" size="16" maxlength="75"></td>
                  </tr>
                <tr>
                  <td>contrato:</td>
                  <td>
				  				  <input type="radio" name="status_socio2" id="status_socio2" value="0" <% if status_socio2 = "0" then %> checked <% end if %>> Sim &nbsp; <input type="radio" name="status_socio2" id="status_socio2" value="1"  <% if status_socio2 = "1" or status_socio2 = "" then %> checked <% end if %>>
              N&atilde;o				  </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
            </div>              <div align="center"></div></td>
          </tr>
          <tr>
            <td colspan="5" class="txt_normal"><table width="100%" border="0" class="txt_normal">
              <tr class="style2">
                <td colspan="4"><div align="center" class="style2"><strong>N&ordm; 3 </strong> </div></td>
                </tr>
              <tr>
                <td width="16%">Nome:</td>
                <td colspan="3"><input name="representante3" type="text" class="caixa_mensagem" id="representante3" value="<%=socio3%>" size="90" maxlength="90">
                  <input name="codigo3" type="hidden" id="codigo3" value="<%=cod3%>"></td>
              </tr>
              <tr>
                <td>Cargo:</td>
                <td width="33%"><input name="cargo_representante3" type="text" class="caixa_mensagem" id="cargo_representante3" value="<%=cargo3%>" size="40" maxlength="75"></td>
                <td width="7%"><div align="right">Aniv.:</div></td>
                <td width="44%"><input name="aniver_representante3" type="text" class="caixa_mensagem" id="aniver_representante3" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<%=aniver3%>" size="5" maxlength="5">
                  <span class="aviso">(dd/mm)</span></td>
                </tr>
              <tr>
                <td>Profiss&atilde;o:</td>
                <td><input name="prof_representante3" type="text" class="caixa_mensagem" id="prof_representante3" value="<%=profissao3%>" size="40"></td>
                <td><div align="right">cnpj/CPF:</div></td>
                <td><input name="cpf_representante3" type="text" class="caixa_mensagem" id="cpf_representante3"  value="<%=FormataCNPJCPF(cpf_socio3)%>" size="22" maxlength="20"></td>
                </tr>
              <tr>
                <td>Email:</td>
                <td><input name="email_representante3" type="text" class="caixa_mensagem" id="email_representante3" value="<%=email_socio3%>" size="40" maxlength="90"></td>
                <td><div align="right">RG/insc:</div></td>
                <td><input name="rg_representante3" type="text" class="caixa_mensagem" id="rg_representante3" value="<%=rg3%>" size="16" maxlength="75"></td>
                </tr>
              <tr>
                <td>contrato:</td>
                <td>
								  <input type="radio" name="status_socio3" id="status_socio3" value="0"  <% if status_socio3 = "0" then %> checked <% end if %>> Sim &nbsp; <input type="radio" name="status_socio3" id="status_socio3" value="1"  <% if status_socio3 = "1" or status_socio3 = "" then %> checked <% end if %>>
              N&atilde;o                </td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>

          <tr>
            <td colspan="5" class="txt_normal"><table width="100%" border="0" class="txt_normal">
              <tr class="style2">
                <td colspan="4"><div align="center" class="style2"><strong>N&ordm;4</strong> </div></td>
                </tr>
              <tr>
                <td width="16%">Nome:</td>
                <td colspan="3"><input name="representante4" type="text" class="caixa_mensagem" id="representante4" value="<%=socio4%>" size="90" maxlength="90">
                  <input name="codigo4" type="hidden" id="codigo4" value="<%=cod4%>"></td>
              </tr>
              <tr>
                <td>Cargo:</td>
                <td width="33%"><input name="cargo_representante4" type="text" class="caixa_mensagem" id="cargo_representante4" value="<%=cargo4%>" size="40" maxlength="75"></td>
                <td width="7%"><div align="right">Aniv.:</div></td>
                <td width="44%"><input name="aniver_representante4" type="text" class="caixa_mensagem" id="aniver_representante4" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<%=aniver4%>" size="5" maxlength="5">
                  <span class="aviso">(dd/mm)</span></td>
                </tr>
              <tr>
                <td>Profiss&atilde;o:</td>
                <td><input name="prof_representante4" type="text" class="caixa_mensagem" id="prof_representante4" value="<%=profissao4%>" size="40"></td>
                <td><div align="right">cnpj/CPF:</div></td>
                <td><input name="cpf_representante4" type="text" class="caixa_mensagem" id="cpf_representante4"  value="<%=FormataCNPJCPF(cpf_socio4)%>" size="22" maxlength="20"></td>
                </tr>
              <tr>
                <td>Email:</td>
                <td><input name="email_representante4" type="text" class="caixa_mensagem" id="email_representante4" value="<%=email_socio4%>" size="40" maxlength="90"></td>
                <td><div align="right">RG/insc:</div></td>
                <td><input name="rg_representante4" type="text" class="caixa_mensagem" id="rg_representante4" value="<%=rg4%>" size="16" maxlength="75"></td>
                </tr>
              <tr>
                <td>contrato:</td>
                <td>
								  <input type="radio" name="status_socio4" id="status_socio4" value="0"  <% if status_socio4 = "0" then %> checked <% end if %>> Sim &nbsp; <input type="radio" name="status_socio4" id="status_socio4" value="1"  <% if status_socio4 = "1" or status_socio4 = "" then %> checked <% end if %>>
              N&atilde;o                </td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>

          <tr>
            <td colspan="5" class="txt_normal"><table width="100%" border="0" class="txt_normal">
              <tr class="style2">
                <td colspan="4"><div align="center" class="style2"><strong>N&ordm;5</strong></div></td>
                </tr>
              <tr>
                <td width="16%">Nome:</td>
                <td colspan="3"><input name="representante5" type="text" class="caixa_mensagem" id="representante5" value="<%=socio5%>" size="90" maxlength="90">
                  <input name="codigo5" type="hidden" id="codigo5" value="<%=cod5%>"></td>
              </tr>
              <tr>
                <td>Cargo:</td>
                <td width="33%"><input name="cargo_representante5" type="text" class="caixa_mensagem" id="cargo_representante5" value="<%=cargo5%>" size="40" maxlength="75"></td>
                <td width="7%"><div align="right">Aniv.:</div></td>
                <td width="44%"><input name="aniver_representante5" type="text" class="caixa_mensagem" id="aniver_representante5" onKeyPress="return formatar_Mascara(this,'##/##');" onBlur="return checkDateDiaMes(this)" value="<%=aniver5%>" size="5" maxlength="5">
                  <span class="aviso">(dd/mm)</span></td>
                </tr>
              <tr>
                <td>Profiss&atilde;o:</td>
                <td><input name="prof_representante5" type="text" class="caixa_mensagem" id="prof_representante5" value="<%=profissao5%>" size="40"></td>
                <td><div align="right">cnpj/CPF:</div></td>
                <td><input name="cpf_representante5" type="text" class="caixa_mensagem" id="cpf_representante5"  value="<%=FormataCNPJCPF(cpf_socio5)%>" size="22" maxlength="20"></td>
                </tr>
              <tr>
                <td>Email:</td>
                <td><input name="email_representante5" type="text" class="caixa_mensagem" id="email_representante5" value="<%=email_socio5%>" size="40" maxlength="90"></td>
                <td><div align="right">RG/insc:</div></td>
                <td><input name="rg_representante5" type="text" class="caixa_mensagem" id="rg_representante5" value="<%=rg5%>" size="16" maxlength="75"></td>
                </tr>
              <tr>
                <td>contrato:</td>
                <td>
								  <input type="radio" name="status_socio5" id="status_socio5" value="0"  <% if status_socio5 = "0" then %> checked <% end if %>> Sim &nbsp; <input type="radio" name="status_socio5" id="status_socio5" value="1"  <% if status_socio5 = "1" or status_socio5 = "" then %> checked <% end if %>>
              N&atilde;o                </td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>


            </table></td>
          </tr>
		
		
	 <tr bgcolor="#CCFFCC">
		    <td height="15" colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center"><strong>para utiliza&ccedil;&Atilde;o fiscal  </strong></div></td>
          </tr>  
          <tr>
            <td colspan="5" class="txt_normal"><table width="100%" border="0" class="txt_normal">
     
                <td width="16%"></td>
                <td </td>
              </tr>
 		
			 <td class="txt_normal"><div align="left">Emiti nota:</div></td>
            <td class="txt_normal">
			<select name="emite_nota" class="txt_normal" id="select20" style="width:130;" >
                <option value="">--- SELECIONE ---</option>
                <option value="S" <% if rs("EMITENOTA") = "S" then response.Write "selected" %>>SIM</option>
                <option value="N" <% if rs("EMITENOTA") = "N" then response.Write "selected" %>>NÃO</option>
              </select></td>
          </tr> 
		    
		   <td class="txt_normal"><div align="left">Subtituto Tributário:</div></td>
            <td class="txt_normal">
			<select name="substituto" class="txt_normal" id="select20" style="width:130;" >
                <option value="">--- SELECIONE ---</option>
                <option value="S" <% if rs("SUBST_TRIB") = "S" then response.Write "selected" %>>SIM</option>
                <option value="N" <% if rs("SUBST_TRIB") = "N" then response.Write "selected" %>>NÃO</option>
              </select></td>
          </tr> 
		  
	  
		  <td class="txt_normal"><div align="left">Tipo Tributação:</div></td>
            <td class="txt_normal">
			<select name="tipo_tributacao" class="txt_normal" id="select20" style="width:130;" >
                <option value="">--- SELECIONE ---</option>
                <option value="1" <% if rs("tipo_tributacao") = "1" then response.Write "selected" %>>Simples Nacional</option>
                <option value="2" <% if rs("tipo_tributacao") = "2" then response.Write "selected" %>>Presumido + Real</option>
              </select></td>
          </tr>         
          
		  <td class="txt_normal"><div align="left">Percentual ISS:</div></td>
            <td class="txt_normal">
			<select name="percentual_ISS" class="txt_normal" id="select20" style="width:130;" >
                <option value="">--- SELECIONE ---</option>
                <option value="0,00" <% if rs("percentual_ISS") = "0,00" then response.Write "selected" %>>0%</option>
				<option value="0,01" <% if rs("percentual_ISS") = "0,01" then response.Write "selected" %>>1%</option>
                <option value="0,02" <% if rs("percentual_ISS") = "0,02" then response.Write "selected" %>>2%</option>
				<option value="0,03" <% if rs("percentual_ISS") = "0,03" then response.Write "selected" %>>3%</option>
				<option value="0,04" <% if rs("percentual_ISS") = "0,04" then response.Write "selected" %>>4%</option>
				<option value="0,05" <% if rs("percentual_ISS") = "0,05" then response.Write "selected" %>>5%</option>
              </select></td>
          </tr>  		  
		 			  
            </table></td>
          </tr>	
		
		 <tr bgcolor="#CCFFCC">
            <td height="15" colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center"><strong>bloco de notas / anota&ccedil;&otilde;es / agenda do cliente </strong></div></td>
          </tr>
          <tr>
            <td colspan="5" class="txt_normal"><div align="right">
						 <%
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "../FCKeditor/"
oFCKeditor.Value = RS("BLOCO_NOTA")
oFCKeditor.ToolbarSet = "Default"
oFCKeditor.Create "bloco_nota"
%>


			</div></td>
          </tr>

		            <!--<tr bgcolor="#CCFFCC">
            <td height="15" colspan="5" bgcolor="#d0d8f7" class="txt_normal"><div align="center"><strong>outros dados </strong></div></td>
          </tr>
                    <tr>
                      <td class="txt_normal">autor:</td>
                      <td colspan="4" class="txt_normal"><strong>
                        <input name="autor" type="text" class="txt_normal" id="autor" value="<%=rs("autor")%>" size="80" maxlength="500">
                      </strong></td>
                    </tr>
          <tr>
            <td class="txt_normal">



              <div align="left">inventor:</div></td>
            <td colspan="4" class="txt_normal"><strong>
              <input name="inventor" type="text" class="txt_normal" id="inventor" value="<%=rs("inventor")%>" size="80" maxlength="500">
            </strong></td>
          </tr>-->

          <tr>
            <td height="47" colspan="5"><div align="center">
                <input name="Button" type="button" class="buttonNormal2" id="Install2" value="  Alterar" onClick="Formulario();">
                &nbsp;&nbsp;
                <input name="Submit2" type="reset" class="buttonNormal2" id="Uninstall2" value="   Limpar">
            </div></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
<p>&nbsp;</p>
</fieldset>
</body>
<script language="javascript">
 function mostrarDataRazao(valor){
 	var d = document.getElementById;
	if (valor == 1){
		d('tr_data_razao').style.visibility = '';
		d('data_razao').value = '';
	}
	else{
		d('tr_data_razao').style.visibility = 'hidden';
	}
 }
</script>
<% if rs("razao1") = 1 then %>
<script language="javascript">
	document.getElementById('tr_data_razao').style.visibility = '';
</script>
<% end if %>
</html>
<% end if %>
<!--#include virtual="/registrafacil/include/db_close.asp"-->