<!--#include virtual="/registrafacil/include/db_open.asp"-->
<%
call logado()
Session.LCID = 1046
%>
<html>
<head>
<title>RegistraFacil :. Financeiro do Contrato</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../script/script.js" type="text/javascript"></script>
<script language="javascript" src="../script/toggle.js" type="text/javascript"></script>
<style type="text/css">
<!--
.style2 {font-size: 9px}
.style3 {font-size: 12px}
.style4 {font-size: 10px}
.style6 {color: #FFFFFF}
.style27 {font-weight: bold; font-size: 10px;}
.style28 {color: #FFFFFF; font-weight: bold; }
.style41 {font-size: 10px}
.btn-acoes {
	background-color: white; color: black; border: 1px solid #ccc; border-radius: 3px;
	padding: 2px 6px;
    font-size: small;
    cursor: pointer;
    text-align: center;
}
-->
.container {
	width: 100%;
    overflow: auto;
	white-space: nowrap;
}
</style>
</head>
<script type="text/javascript" >
function selecionarTodos(cxSel) {
  checkboxes = document.getElementsByName('cxItemBol');
  for(var i=0, n=checkboxes.length;i<n;i++) {
    checkboxes[i].checked = cxSel.checked;
  }
}
</script>
<body>
<%
SQL1=""
PagAtual = Request.QueryString("PagAtual") 'página atual'
Session("PagAtual") = PagAtual

wtaxa = request.form("taxa")
wimposto = request.form("imposto")


if Request.Form("data_ini") <> "" and Request.Form("data_fim") <> "" then
   session("data1") = Request.Form("data_ini")
   session("data2") = Request.Form("data_fim")
end if

cont = 0
For each Obj In Request.Form
	If Request.Form(Obj) <> "" and Request.Form(obj) <> "Pesquisar" and Obj <> "txtCliente" and CStr(Obj) <> "status"  then
		cont = cont + 1
	end if
Next
'Response.Write cont & "<br>"'
if cont = 0 then
	Response.Write ""
	Else
	SQL = "SELECT *, E.COD_SUB AS SUBFRANQUEADO FROM  GEDPROPOSTA b inner join GEDCONTRATO c on b.cod_proposta = c.cod_proposta inner join GEDPARCELAMENTO d on B.COD_PROPOSTA = d.COD_PROPOSTA inner join GEDCLIENTES E on (E.CNPJ_CPF = B.CNPJ_CPF) left join gedtipoprop f on (f.cod_tipo = b.cod_tipo) where (d.imp_boleto = 1 or d.imp_boleto is null) and d.data_pagamento is null and " 
'	SQL = "SELECT a.*, b.cod_tipo, b.desc_tipo, c.* FROM GEDPROPOSTA a, GEDTIPOPROP b, GEDCONTRATO c, GEDPARCELAMENTO d WHERE a.COD_TIPO = b.COD_TIPO and  A.COD_PROPOSTA = C.COD_PROPOSTA and a.cod_proposta = d.cod_proposta and "'
	i = 1
	For Each Obj In Request.Form ' Supondo o uso do method = "post"'
		If Request.Form(Obj) <> "" and Request.Form(Obj) <> "Pesquisar" and Obj <> "txtCliente" and CStr(Obj) <> "data_ini" and CStr(Obj) <> "imposto" and CStr(Obj) <> "taxa" and CStr(Obj) <> "data_fim" and CStr(Obj) <> "status" then
			if i = cont then
				if CStr(Obj) = "marca" then
					SQL = SQL & Obj & " LIKE '%" & Request.Form(obj) & "%' "
				else
					if CStr(Obj) = "historico" then
						SQL = SQL & "d.historico LIKE '%" & Request.Form(obj) & "%' "
					elseif CStr(Obj) = "cnpj_cpf" then
						SQL = SQL & "e.cnpj_cpf " & "='" & Request.Form(Obj) & "' "
					else
						SQL = SQL & Obj & "='" & Request.Form(Obj) & "' "
					end if
				end if
			else 
				if CStr(Obj) = "marca" then
					SQL = SQL & Obj & " LIKE '%" & Request.Form(obj) & "%' and "
				else
					if CStr(Obj) = "historico" then
						SQL = SQL & "d.historico LIKE '%" & Request.Form(obj) & "%' and "
					elseif CStr(Obj) = "cnpj_cpf" then
						SQL = SQL & "e.cnpj_cpf " & "='" & Request.Form(Obj) & "' and "
					else
						SQL = SQL & Obj & "='" & Request.Form(Obj) & "' and "
					end if 
				end if
				i = i + 1
			End if
		End if
	Next 
	
		if Request.Form("data_ini") <> "" and Request.Form("data_fim") <> "" then
			SQL = SQL & " data_parcela >= '" & Date_US(Request.Form("data_ini")) & "' and data_parcela <= '"& Date_US(Request.Form("data_fim")) & "' "
		else
			SQL = SQL & " "
		End if 

		if Request.Form("status") = "0" then
			SQL = SQL & " and DATEDIFF(day, GETDATE(), data_parcela) > 40 "
		elseif Request.Form("status") = "1" then
			SQL = SQL & " and (DATEDIFF(day, GETDATE(), data_parcela) > 30 AND DATEDIFF(day, GETDATE(), data_parcela) <= 40) "
		elseif Request.Form("status") = "2" then
			SQL = SQL & " and (DATEDIFF(day, GETDATE(), data_parcela) > 15 AND DATEDIFF(day, GETDATE(), data_parcela) <= 30) "
		elseif Request.Form("status") = "3" then
			SQL = SQL & " and (DATEDIFF(day, GETDATE(), data_parcela) > 0 AND DATEDIFF(day, GETDATE(), data_parcela) <= 15) "
		end if
	
	'SQL = SQL & " order by data_parcela ASC, cod_contrato DESC"
	'SQL = SQL & " order by cliente asc, data_parcela desc"
	SQL = SQL & " order by nome ASC, data_parcela ASC"
end if

'response.write(SQL)

if sql <> "" and session("sql_contrato") = "" then
	session("sql_contrato") = SQL
end if

if sql <> "" and session("sql_contrato") <> "" then
	session("sql_contrato") = SQL
end if

if sql = "" then
	SQL = session("sql_contrato")
end if


if SQL = "" then 
	call CaixaMensagem ("erro","ERRO NA CONSULTA! ","Voc&ecirc; deve preencher ao menos um dos campos para realizar a consulta.",true,"Avançar",false,"")
	response.End()
	else
	rs.open SQL, conn, 3, 3
	end if
if rs.eof then
	call CaixaMensagem ("erro","ERRO NA CONSULTA! ","Nenhum registro foi encontrado. Cheque os par&acirc;metros da consulta.",true,"Avançar",false,"")
   	response.end
end if


wValor_Total = 0 
if not rs.eof then
	while not rs.eof
	
		if RS("tipo_moeda") = 0 then
			 wValor_Total = wValor_Total + rs("valor_parcela")
			 wvalor_parcela_ind =  rs("valor_parcela")
		else
		  	set rsSalario = Server.CreateObject("ADODB.Recordset")
			rsSalario.open "select * from gedsalario_minimo where data_salario <= '" & Date_US(RS("data_parcela")) & "' order by cod_salario desc",conn,2,3
			if not rssalario.eof then
			   wsalario = ccur(rssalario("valor_salario")) 
			else
			   wsalario = 1   
			end if
			wValor_Total = wValor_Total + (RS("valor_parcela") * wsalario)
			wvalor_parcela_ind = (RS("valor_parcela") * wsalario)
			rssalario.close	
		end if	  

		rs.movenext
	wend
	rs.movefirst
else
	wValor_Total = 0
end if

if IsNull(wValor_Total) then
	wValor_Total = 0
end if	

%>

<fieldset>
<legend>
<table border="0">
  <tr>
    <td width="32"><div align="center"><img src="../menu_imagens/btn_certificates_bg.gif" alt="Proposta" width="32" height="32"></div></td>
    <td width="268"><div align="left" class="style2 style3">Listagem para Emiss&atilde;o de Boleto: </div></td>
  </tr>
</table>
</legend>
<table width="100%" border="0">
  <tr class="txt_normal">
    <td width="167" nowrap> <strong> encontrados:</strong> </td>
    <td width="76"><strong>
    <% = RS.RECORDCOUNT %> </strong></td>
    <td width="130"><strong>valor total:</strong> </td>
    <td width="155"><%=formatcurrency(wValor_Total,2)%></td>
    <td width="443"><div style="text-align: right; margin: 0 10px 0 0;" ><a href="../contrato/Relatorio_contrato.asp?wsql=<%=sql%>" alt="Exportar para Excel" ><button type="button" onClick="return confirm('Exportar Dados para Excel?')" class="btn-acoes" alt="Exportar para Excel" ><img src="../imagens/excel.jpg" alt="" width="18" height="18" border="0"></button></a>
      <!--<input type="button" name="btImp" class="buttonNormal2" id="Print2" value="  Imprimir" onClick="this.style.visibility='hidden';self.print();">--></div></td>
  </tr>
</table>

<div class="container" >
<table width="100%"  border="0" align="center">
  <tr>
    <td valign="top"><table width="100%" border="0" align="center" cellspacing="1" id="foo">
      <tr bgcolor="#0099CC">
		<td class="txt_normal"><div align="center" class="style6"><span class="style27"><input type="checkbox" name="cxSelTodos" alt="Selecionar Todos" onClick="selecionarTodos(this)" ></span></div></td>
		<td class="txt_normal"><div align="center" class="style6"><span class="style27">Cont N&ordm; </span></div></td>
        <td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> CPF/CNPJ </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Cliente </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Valor Bruto </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Valor ISS </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Valor IR </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Valor PIS </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Valor COFINS </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Valor CSLL </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Dedução </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Valor Líquido </span> </div></td>
		<td bgcolor="#0099CC" class="txt_normal"><div align="center" class="style6"><span class="style27"> Vencimento </span> </div></td>
        <td class="txt_normal"><div align="center" class="style6"><span class="style27">Parcela</span></div></td>
        <td class="txt_normal"><div align="center" class="style6"><strong><span class="style27"> Consultor </span></strong></div></td>
        <td class="txt_normal"><div align="center" class="style28">Processos</div></td>
        <td class="txt_normal"><div align="center" class="style28">Emite Nota?</div></td>
        <td class="txt_normal"><div align="center" class="style28">Substituto Tributário?</div></td>
        <td class="txt_normal"><div align="center"><span class="style6"><span class="style26"><b>Tipo Tributação</b></span></span></div></td>
        <td class="txt_normal" colspan="2"><div align="center" class="style6"><span class="style27"> A&ccedil;&atilde;o </span></div></td>
      </tr>
      <%

'############## paginacao Introdução  #################

'------- Coloque aqui a quantidade de registros que você deseja por página --------

Const NumPorPage = 15

'Verifica qual a página solicitada
   Dim PagAtual

   IF Request.QueryString("PagAtual") = "" Then
               PagAtual = 1 'Primeira página
         Else
                PagAtual = Request.QueryString("PagAtual")
   End If

   'Cria conexão com o Banco de Dados, já abrir anteriormente
   'Criado anteriormente Set RS = Server.CreateObject("ADODB.Recordset")
   '>>> FIZ EM CIMA RS.CursorLocation = 3        Acerta a posição do cursor . 3 ou adUseClient
  
   RS.CacheSize = NumPorPage 'Define o tamanho do Cache = para o número de registros

   'Cria a String SQL
   '>>> FIZ EM CIMA Dim SQLpag
   '>>> FIZ EM CIMA SQLpag = "SELECT * FROM jogos"
   '>>> FIZ EM CIMA RS.Open SQLpag, Conn    Abre o RecordSet


    RS.MoveFirst                'Move o RecorSet para o início 
    RS.PageSize = NumPorPage    'Coloca a quantidade de páginas

    Dim TotalPages              'Pega o número total de páginas
    TotalPages = RS.PageCount

    RS.AbsolutePage = PagAtual  'Configura a página atual

'############## paginacao Introdução  - FIM #################

On Error Resume Next
Err.Clear

Count = 0       'Zera o contador
   
'Inicia a Função DO, utilizando a quantidade de páginas especificadas
'Ou seja ele irá executar a ação até que o valor Count seja menor que "20" como está no nosso exemplo
 
i = 0 
cor = "#FFFFFF"

Set rsConsFecha = Server.CreateObject("ADODB.Recordset")

DO WHILE NOT RS.EOF And Count < RS.PageSize  'paginacao And Count < RS.PageSize '
%>
      <tr bgcolor="<%=cor%>">
	  <td width="4%"><div align="center"><span class="aviso"><input type="checkbox" name="cxItemBol" ></div></td>
<% if rs.fields("impresso") = 1 then %>
          <td width="4%"><div align="left"><span class="aviso"><%=rs("cod_contrato")%></span></a></div></td>
<% else %>
          <td width="3%"><div align="left"><span class="style4"><%=rs("cod_contrato")%></span></a></div></td>
<% end if %>
        <td><div align="center" class="style4">
            <div align="center"><%=rs("CNPJ_CPF")%></div>
        </div></td>
		<td><div align="center" class="style4">
            <div align="center"><%=ExibeNome(rs("CNPJ_CPF"), "cliente")%></div>
        </div></td>
        
                <td><div align="center" class="style41"> 
		 <%
		' cálculo da parcela
        wvalor_parcela_ind = 0
        wvalor_parcela = 0 
		
		Dim vValorParcelaBase
		vValorParcelaBase = IIf(IsNull(rs("valor_parcela")), 0, rs("valor_parcela"))
		
		'''Response.Write IIf(IsNull(rs("valor_parcela")), rs("valor_parcela"), 0)
		'''Response.End
		
		if RS("tipo_moeda") = 0 then
			 wvalor_parcela = wvalor_parcela + vValorParcelaBase
			 wvalor_parcela_ind =  vValorParcelaBase
		else
		  	set rsSalario = Server.CreateObject("ADODB.Recordset")
			rsSalario.open "select * from gedsalario_minimo where data_salario <= '" & Date_US(RS("data_parcela")) & "' order by cod_salario desc",conn,2,3
			if not rsSalario.eof then
			   wsalario = ccur(IIf(IsNull(rsSalario("valor_salario")), 0, rsSalario("valor_salario"))) 
			else
			   wsalario = 1   
			end if
			   wvalor_parcela = wvalor_parcela + (vValorParcelaBase * wsalario)
			   wvalor_parcela_ind = (vValorParcelaBase * wsalario)
			rsSalario.close	
		end if	  
		
		if IsNull(wvalor_parcela) then
			wvalor_parcela = 0
		end if 
		response.write formatcurrency(wvalor_parcela,2)

		%>		
		</div></td>

		 <%
		 ' Código Elano 11/01/18
		 
		 dim wvalor_desconto , wvalor_liquido
		 dim valorISS, valorIR, ValorPis, ValorCofins, ValorCSLL
		 dim vPercISS
		 wvalor_desconto = 0
		 wvalor_liquido = 0 
		 valorISS = 0
		 valorIR = 0
		 ValorPis = 0
		 ValorCofins = 0
		 ValorCSLL = 0
		 
		 vPercISS = IIf(IsNull(RS("percentual_ISS")), 0, RS("percentual_ISS"))
		 
		if (RS("EMITENOTA") = "N") or (RS("EMITENOTA") = "") or (IsNull(RS("EMITENOTA"))) then
		wvalor_liquido = wvalor_parcela
		wvalor_desconto = 0
		end if 
		
		if (RS("EMITENOTA") = "S" and RS("SUBST_TRIB") = "S" and RS("TIPO_TRIBUTACAO") = 1 ) then ' Simples Nacional
			wvalor_desconto = (wvalor_parcela * vPercISS)			
			wvalor_liquido = wvalor_parcela - (wvalor_parcela * vPercISS)
			valorISS = (wvalor_parcela * vPercISS)
		end if
		
		if (RS("EMITENOTA") = "S" and RS("SUBST_TRIB") = "S" and RS("TIPO_TRIBUTACAO") = 2 ) then ' Lucro Presumido
			set rsimpostofederais = Server.CreateObject("ADODB.Recordset")
			'rsimpostofederais.open "SELECT IR, PIS, COFINS, CSLL, ANO FROM minhamarca2.IMPOSTOS where ano <= '" & Date_US(RS("data_parcela")) & "' order by ano",conn,2,3
	        rsimpostofederais.open "SELECT IR, PIS, COFINS, CSLL, ANO FROM minhamarca2.IMPOSTOS",conn,2,3
        	
			dim vIR, vPIS, vCOFINS, vCSLL
			
			' vIR = 0
		    ' vPIS = 0
		    ' vCOFINS = 0
		    ' vCSLL = 0
			
			if not rsimpostofederais.eof then
			  
			  vIR = IIf(IsNull(rsimpostofederais("IR")),0,rsimpostofederais("IR"))
			  vPIS = IIf(IsNull(rsimpostofederais("PIS")),0,rsimpostofederais("PIS"))
			  vCOFINS = IIf(IsNull(rsimpostofederais("COFINS")),0,rsimpostofederais("COFINS"))
			  vCSLL = IIf(IsNull(rsimpostofederais("CSLL")),0,rsimpostofederais("CSLL"))
			  
			  wvalor_desconto = (wvalor_parcela * vPercISS) + (wvalor_parcela * vIR) + (wvalor_parcela * vPIS)
			  wvalor_desconto = wvalor_desconto + (wvalor_parcela * vCOFINS) + (wvalor_parcela * vCSLL)
			  
			  valorISS = ( wvalor_parcela * vPercISS)
			  valorIR = ( wvalor_parcela * vIR )
			  ValorPis = ( wvalor_parcela * vPIS)
			  ValorCofins =  ( wvalor_parcela * vCOFINS )
			  ValorCSLL = ( wvalor_parcela * vCSLL )
			
			end if
			
		    ' para obter os percentuais
			'valorISS = RS("percentual_ISS") * 100
			'valorIR = rsimpostofederais("IR") * 100
			'ValorPis = rsimpostofederais("PIS") * 100
			'ValorCofins =  rsimpostofederais("COFINS") * 100
			'ValorCSLL = rsimpostofederais("CSLL") * 100
			
			rsimpostofederais.close	
            wvalor_liquido = wvalor_parcela - wvalor_desconto
		end if 
		
		%>

		<td><div align="center" class="style41"> 
		 <% response.write formatcurrency(valorISS , 2) %>		
		</div></td>
		
		<td><div align="center" class="style41"> 
		 <% response.write formatcurrency(valorIR , 2) %>		
		</div></td>
		
		<td><div align="center" class="style41"> 
		 <% response.write formatcurrency(ValorPis , 2)%>	
		</div></td>
		
		<td><div align="center" class="style41"> 
		 <% response.write formatcurrency(ValorCofins , 2) %>	
		</div></td>
		
		<td><div align="center" class="style41"> 
		<% response.write formatcurrency(ValorCSLL , 2)%>	
		</div></td>
		
		<td><div align="center" class="style41"> 
		 <%
		response.write formatcurrency(wvalor_desconto,2)
 		%>		
		</div></td>	
	
		<td><div align="center" class="style41"> 
		 <%
		response.write formatcurrency(wvalor_liquido,2)
 		%>		
		</div></td>
		
		
		<td><div align="center"><span class="style41"><%=Date_BR(rs("data_parcela"))%></span></div></td>
		<td><div align="center"><span class="style4"><%=rs("historico")%></span></div></td>
		
		<td><div align="center" class="style4">
            <div align="center"><%=ExibeNome(rs("cod_sub"), "subfranqueado")%></div>
        </div></td>
		
		
      <% 
		  	set rsMarcas = Server.CreateObject("ADODB.Recordset")
		  	rsMarcas.open "SELECT * FROM GEDSERVPROP WHERE exibir_proposta = 1 and  COD_PROPOSTA='" & rs("cod_proposta") & "'", conn, 3, 3
		  	
			dim wprocessos
			
			if rsMarcas.eof then
		  		'Marca = "Não cadastradas"
			else
		   		while not rsMarcas.eof
				    if Not isNull(rsMarcas("num_processo")) then
					 wprocessos = wprocessos & rsMarcas("num_processo") & " | "
					end if
					
					rsMarcas.MoveNext
				Wend
				
				if not isNull(wprocessos) and Len(wprocessos) > 0 then
					wprocessos = Left(wprocessos, Len(wprocessos) - 3)
				end if				
				
			end if
		  %>	
	
	   <td><div align="center" class="style41"> 
		 <% response.write (wprocessos) %>		
		</div></td>
        
		<%	
		rsMarcas.Close
		Set rsMarcas = Nothing
		 wprocessos = ""
		%>

		<td><div align="center"><span class="style4"><%=rs("EMITENOTA")%></span></div></td>
		<td><div align="center"><span class="style4"><%=rs("SUBST_TRIB")%></span></div></td>
		<td><div align="center"><span class="style4"><%=rs("TIPO_TRIBUTACAO")%></span></div></td>
		
        <td width="5%" valign="top" bgcolor="<%=cor%>"><div align="center">
            <div align="center"><a href="tela_taxa_imposto.asp?cod_contrato=<%=rs("cod_contrato")%>&cnpj_cpf=<%=rs("cnpj_cpf")%>&data=<%=rs("data_parcela")%>&taxa=<%=wtaxa%>&imposto=<%=wimposto%>&titulo=<%=rs("titulo")%>"><img src="../menu_imagens/147.jpg" title="Exibir Boleto" width="20" height="20" border="0">
            </a> </div>
        </div></td>
        <td width="6%" valign="top" bgcolor="<%=cor%>"><div align="center"><a href="cad_nota_fiscal.asp?acao=1"><img src="../menu_imagens/btn_edit_bg.gif" title="Cadastro de Nota Fiscal" width="22" height="22" border="0"></a>
		<%
		wemitir = rs("emitir")
	    wretencao = rs("retencao")
		
		 if wemitir = "1" and wretencao = "0" then
		  	response.write("<img src=""../imagens/ok.gif"" width=""18"" height=""18"" border=""0"" ><br /><br />")
			
			else if wemitir = "1" and wretencao = "1" then
			response.write("<img src=""../imagens/ok.gif"" width=""18"" height=""18"" border=""0"" >")
			response.write("<img src=""../imagens/btn_aviso.gif"" width=""18"" height=""18"" border=""0"" ><br /><br />")
		  end if
		  end if
		
		%>
		
		</div></td>
      </tr>
      
      <%
    i = i+1
 Count = Count + 1   'paginacao
    RS.MoveNext
' ----- linhas coloridas -------
if cor = "#d0d8f7" then
cor = "#FFFFFF"
else
cor = "#d0d8f7"
end if 
'-------------------------------
    LOOP                'tb paginacao
Set rsConsFecha = nothing
%>
      <tr class="txt_normal">
        <td colspan="10" class="txt_normal"><div align="left">
            <%
			'############## paginacao 02 ####################
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@@@@@@ ALEX TEIXEIRA - alexct@hotmail.com @@@@@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

 'Coloca o N&ordm; p&aacute;gina atual / N&ordm; Total de p&aacute;ginas

 Response.Write("<div align='right'><strong> P&aacute;gina " & PagAtual & " de " & TotalPages & " </strong></div>")   

'Mostra os bot&otilde;es: Anterior e Pr&oacute;ximo, utilizando da op&ccedil;&atilde;o de IF 

Response.Write("<strong> P&aacute;ginas: </strong>") 

'----------- Numeros - Calculos ---------------------------------------------

var01 = Len(PagAtual) 'L&ecirc; o tamanho do numero
var02 = var01 - 1 'subtrai um da variavel , retirando o digito menos sig.
var03 = Left(PagAtual,var02) 'obtem os digitos mais  sig. do numero
var04 = Right(PagAtual,1)    'obtem o digito menos sig. do numero
var05 = var03 & 0 ' Acrecenta ZERO no final
IF var04 <> 0 THEN     ' condi&ccedil;&atilde;o se o digito menos sig. &eacute; Zero
	inicial = var05 + 1
	final = inicial + 9  
ELSE
	inicial = var05 - 9  
	final = var05
END IF

indice_i = var04 - 1 'ultimo digito  - 1
indice_f = 10 - var04 ' 10 - digito menos sig.


' If CInt(inicial) < 1 Then inicial = 1
     
'If CInt(final) > CInt(TotalPages) Then final = TotalPages

final = TotalPages
inicial = 1
'------------------------------------------------------------------------------


IF PagAtual > 1 THEN 

'Se for a primeira p&aacute;gina, Mostra apenas o bot&atilde;o Pr&oacute;ximo e Ultima
      Response.Write("<B><font color=""#660066"" size=""1"" face=""Arial"">") 
      'Response.Write("<a href='"& Request.ServerVariables("SCRIPT_NAME")&"?PagAtual=" &  1 & "'>")
      'Response.Write("Primeira") 
      Response.Write("</font></B>  ")
      
      Response.Write("<B><font color=""#660066"" size=""2"" face=""Arial"">") 
      'Response.Write("<a href='"& Request.ServerVariables("SCRIPT_NAME")&"?PagAtual=" & PagAtual - 1 & "'>")
      'Response.Write("Anterior") 
      Response.Write("</font></B>  ")

      IF PagAtual > 10 THEN

       Response.Write("<B><font color=""#660066"" size=""2"" face=""Arial"">") 
       'Response.Write("<a href='"& Request.ServerVariables("SCRIPT_NAME")&"?PagAtual=" & inicial - 1 & "'>")
       'Response.Write("...") 
    ' Response.Write("</a></font></B>  ")

          ELSE

      Response.Write("<B><font color=""#CCCCCC"" size=""2"" face=""Arial"">") 
       'Response.Write("....") 
       Response.Write("</font></B>  ")

    END IF

  Else

      Response.Write("<B><font color=""#EEEEEE"" size=""1"" face=""Arial"">") 
      'Response.Write("Primeira") 
      Response.Write("</font></B>  ")

      Response.Write("<B><font color=""#CCCCCC"" size=""2"" face=""Arial"">") 
      'Response.Write("Anterior") 
      Response.Write("</font></B>  ")

      'Response.Write("<B><font color=""#CCCCCC"" size=""2"" face=""Arial"">") 
      'Response.Write("...") 
      'Response.Write("</font></B>  ")

End If

'---------------------- NUMEROS  ---------------------------

For i = inicial To final
     If CInt(i)=CInt(PagAtual) Then
         Response.Write "<font color=""#660066"" size=""1"" face=""Arial"">[ <B>" & i & "</B> <font color=""#660066"">]</font>  "
     END IF
     If CInt(i) < CInt(PagAtual) Then
      Response.Write "<font color=""#660066"" size=""1"" face=""Arial""><a href='"& Request.ServerVariables("SCRIPT_NAME")&"?PagAtual=" & i & "'>" & i & "</a></font>  "
     END IF
     If CInt(i) > CInt(PagAtual) Then
         Response.Write "<font color=""#660066"" size=""1"" face=""Arial""><a href='"& Request.ServerVariables("SCRIPT_NAME")&"?PagAtual=" & i & "'>" & i & "</a></font>  "
  END IF
Next

'------------------------------------------------------

IF CInt(PagAtual) <> CInt(TotalPages) THEN 


p1 = Left(PagAtual,var02) 
p2 = Left(TotalPages,var02)
p3 = Left(TotalPages,var02) & 0


'##### CONDI&Ccedil;&Otilde;ES ########
'digitos mais significativos do Numero com 1 no fim > PagAtual
'EX:  21   [ 22 ]  23   24  25         2 com 1 => 21 > 22 (F)
'OU
'PagAtual <= 10      E    TotalPages > 10
'EX:  ... 1  2  3  4 [ 5 ]  6  ...              5 <= 10 (V)  E   6 > 10 (F) 

IF (p1 > PagAtual) or ((PagAtual <= 10) and (TotalPages > 10)) THEN 


       Response.Write("<B><font color=""#660066"" size=""2"" face=""Arial"">")
       'Response.Write("<a href='"& Request.ServerVariables("SCRIPT_NAME")&"?PagAtual=" & final + 1 & "'>")
       'Response.Write("...")
       Response.Write("</font></B>  ") 

         ELSE

       Response.Write("<B><font color=""#CCCCCC"" size=""2"" face=""Arial"">")
       'Response.Write("...") 
       Response.Write("</font></B>  ")

   END IF

      Response.Write("<B><font color=""#660066"" size=""2"" face=""Arial"">")
      'Response.Write("<a href='"& Request.ServerVariables("SCRIPT_NAME")&"?PagAtual=" & PagAtual + 1 & "'>")
      'Response.Write("Pr&oacute;xima")
      Response.Write("</font></B>  ") 

      Response.Write("<B><font color=""#660066"" size=""1"" face=""Arial"">")
      'Response.Write("<a href='"& Request.ServerVariables("SCRIPT_NAME")&"?PagAtual=" & TotalPages & "'>")
      'Response.Write("Ultima")
      Response.Write("</font></B>  ")        

 ELSE

      Response.Write("<B><font color=""#CCCCCC"" size=""2"" face=""Arial"">")
      'Response.Write("...") 
      Response.Write("</font></B>  ")

      Response.Write("<B><font color=""#CCCCCC"" size=""2"" face=""Arial"">")
      'Response.Write("Pr&oacute;xima") 
      Response.Write("</font></B>  ")

      Response.Write("<B><font color=""#EEEEEE"" size=""1"" face=""Arial"">")
      'Response.Write("Ultima") 
      Response.Write("</font></B>  ")
End If 
'################## fim paginacao 02 #######################'
Rs.Close  
Set RS = Nothing
%>
        </div></td>
      </tr>
      
    </table></td>
  </tr>
</table>
</div>
</fieldset>
</body>
</html>
<!--#include virtual="/registrafacil/include/db_close.asp"-->