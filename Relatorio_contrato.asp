<!--#include virtual="/registrafacil/include/db_open.asp"-->
<%
call logado()
Session.LCID = 1046
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition","attachment;filename=listagem.xls" 
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!--<link href="../css/style.css" rel="stylesheet" type="text/css">-->
<script language="javascript" src="../script/script.js" type="text/javascript"></script>
<script language="javascript" src="../script/toggle.js" type="text/javascript"></script>
<style type="text/css">
<!--
.style4 {font-size: 10px}
.style26 {color: #000000}
.style27 {font-weight: bold; font-size: 10px;}
.style41 {font-size: 10px}
-->
</style>
</head>
<body>
<%
			set rsPropostas = Server.CreateObject("ADODB.Recordset")
			
			rsPropostas.open "select * from gedtipoprop", conn, 3, 3
			
PagAtual = Request.QueryString("PagAtual") 'página atual

cont = 0
For each Obj In Request.Form
	If Request.Form(Obj) <> "" and Request.Form(obj) <> "Pesquisar" then
		cont = cont + 1
	end if
Next
'Response.Write cont & "<br>"
if cont = 0 then
	Response.Write ""
	Else
	SQL = "SELECT a.*, b.desc_tipo, c.* FROM GEDPROPOSTA a, GEDTIPOPROP b, GEDCONTRATO c WHERE a.COD_TIPO = b.COD_TIPO and  A.COD_PROPOSTA = C.COD_PROPOSTA and "
	i = 1
	For Each Obj In Request.Form ' Supondo o uso do method = "post"
		If Request.Form(Obj) <> "" and Request.Form(Obj) <> "Pesquisar" and CStr(Obj) <> "data_ini" and CStr(Obj) <> "data_fim"  then
			if i = cont then
				if CStr(Obj) = "marca" then
					SQL = SQL & Obj & " LIKE '%" & Request.Form(obj) & "%' "
					else
					SQL = SQL & Obj & "='" & Request.Form(Obj) & "' "
				end if
			else 
				if CStr(Obj) = "marca" then
					SQL = SQL & Obj & " LIKE '%" & Request.Form(obj) & "%' and "
					else
					SQL = SQL & Obj & "='" & Request.Form(Obj) & "' and "
				end if
				i = i + 1
			End if
		End if
	Next 
	
	if Request.Form("data_ini") <> "" and Request.Form("data_fim") <> "" then
		SQL = SQL & " DATA_CONTRATO >= '" & Date_US(Request.Form("data_ini")) & "' and DATA_CONTRATO <= '"& Date_US(Request.Form("data_fim")) & "'"
	End if 
	
	SQL = SQL & " order by data_contrato DESC, cod_contrato DESC"
end if

'Response.Write(SQL)
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
	
	'''Response.Write(SQL)
	'''Response.End
	
	rs.open SQL, conn, 3, 3
end if
if rs.eof then
	call CaixaMensagem ("erro","ERRO NA CONSULTA! ","Nenhum registro foi encontrado. Cheque os par&acirc;metros da consulta.",true,"Avançar",false,"")
   	response.end
end if

%>

<!--<table width="100%" border="0">
  <tr class="txt_normal">
    <td width="70%" nowrap class="style2"><div align="left" class="style28">RELA&Ccedil;&Atilde;O DE CONTRATOS </div></td>
    <td width="463"><div align="right"><input type="button" name="btImp" class="buttonNormal2" id="Print2" value="  Imprimir" onClick="this.style.visibility='hidden';self.print();">
    </div></td>
  </tr>
</table>-->
<table width="100%"  border="0" align="center">
  <tr>
    <td height="115" valign="top"><table width="100%" border="0" align="center" cellspacing="0" id="foo">
      <tr bgcolor="#CCCCCC">
 <td width="3%" class="txt_normal"><div align="center" class="style26"><span class="style27">Cont N&ordm; </span></div></td>  
		<td width="16%" class="txt_normal"><div align="center" class="style26"><span class="style27"> CNPJ/CPF </span> </div></td>
		<td width="16%" class="txt_normal"><div align="center" class="style26"><span class="style27"> Cliente </span> </div></td>     
        <td width="22%" class="txt_normal"><div align="center"><span class="style27">Valor Bruto</span></div></td>    
		<td width="22%" class="txt_normal"><div align="center"><span class="style27">Valor ISS</span></div></td>
		<td width="22%" class="txt_normal"><div align="center"><span class="style27">Valor IR</span></div></td>
		<td width="22%" class="txt_normal"><div align="center"><span class="style27">Valor PIS</span></div></td>
		<td width="22%" class="txt_normal"><div align="center"><span class="style27">Valor COFINS</span></div></td>
		<td width="22%" class="txt_normal"><div align="center"><span class="style27">Valor CSLL</span></div></td>
		<td width="6%" class="txt_normal"><div align="center"><span class="style27">Deducão</span></div></td>
		<td width="6%" class="txt_normal"><div align="center"><span class="style27">Valor Líquido</span></div></td>
		<td width="6%" class="txt_normal"><div align="center"><span class="style27">Vencimento</span></div></td>
        <td width="6%" class="txt_normal"><div align="center"><span class="style27">Parcela</span></div></td>
		<td width="18%" class="txt_normal"><div align="center" class="style26"><strong><span class="style27"> Consultor </span></strong></div></td>
		<td width="6%" class="txt_normal"><div align="center"><span class="style27">Processos</span></div></td>		
		<td width="6%" class="txt_normal"><div align="center"><span class="style27">Emite Nota?</span></div></td>
		<td width="6%" class="txt_normal"><div align="center"><span class="style27">Substituto Tributário?</span></div></td>
		<td width="6%" class="txt_normal"><div align="center"><span class="style27">Tipo Tributação</span></div></td>

        </tr>
      <%
On Error Resume Next
Err.Clear

Count = 0       'Zera o contador
   
'Inicia a Função DO, utilizando a quantidade de páginas especificadas
'Ou seja ele irá executar a ação até que o valor Count seja menor que "20" como está no nosso exemplo
 
i = 0 
cor = "#FFFFFF"

Set rsConsFecha = Server.CreateObject("ADODB.Recordset")

DO WHILE NOT RS.EOF 'And Count < RS.PageSize  paginacao And Count < RS.PageSize 
	
	if not isNull(rs("cod_proposta")) then
		rsConsFecha.open("SELECT TOP 1 * FROM GEDHONORARIOS_PROP WHERE COD_PROPOSTA = '" & rs("cod_proposta") & "'"), conn, 3, 3
	end if
	
	if rsConsFecha.EOF then
		Fechamento = False
	else
		Fechamento = True
	end if
	
	rsConsFecha.close
	%>
   <tr bgcolor="<%=cor%>" class="txt_normal">
        <td height="24"><div align="left"><span class="style4"><%=rs("cod_contrato")%></span></div></td>
		<td><div align="center" class="style4">
            <div align="center">.<%=rs("CNPJ_CPF")%></div>
        </div></td>
        <td><div align="center" class="style4">
        <div align="left"><%=ExibeNome(rs("CNPJ_CPF"), "cliente")%></div>
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
		response.write formatcurrency(wvalor_parcela,3)

		%>		
		</div></td>
		

		 <%
		 ' Código Elano 11/01/18
		 
		 dim wvalor_desconto , wvalor_liquido
		 dim valorISS, valorIR, ValorPis, ValorCofins, ValorCSLL
		 dim vPercISS
		 vPercISS = 0
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
		
		
		if (RS("EMITENOTA") = "S" and RS("SUBST_TRIB") = "S" and RS("TIPO_TRIBUTACAO") = 1 ) then  ' Simples Nacional
			valorISS = 0
			wvalor_liquido = 0
			wvalor_desconto = (wvalor_parcela * vPercISS)			
			wvalor_liquido = wvalor_parcela - wvalor_desconto
			valorISS = (wvalor_parcela * vPercISS)
		end if
		
		
		if (RS("EMITENOTA") = "S"  and RS("TIPO_TRIBUTACAO") = 1 ) and RS("SUBST_TRIB") = "N"  then  ' Simples Nacional sem substituição 
			wvalor_liquido = 0
			wvalor_liquido = wvalor_parcela
		    wvalor_desconto = 0
		end if
		
		
		
		if (RS("EMITENOTA") = "S" and RS("SUBST_TRIB") = "S" and RS("TIPO_TRIBUTACAO") = 2 ) then ' Lucro Presumido
			set rsimpostofederais = Server.CreateObject("ADODB.Recordset")
			'rsimpostofederais.open "SELECT IR, PIS, COFINS, CSLL, ANO FROM minhamarca2.IMPOSTOS where ano <= '" & Date_US(RS("data_parcela")) & "' order by ano",conn,2,3
	        rsimpostofederais.open "SELECT IR, PIS, COFINS, CSLL, ANO FROM minhamarca2.IMPOSTOS",conn,2,3
        	
			dim vIR, vPIS, vCOFINS, vCSLL, VreceitasFederais
			
			 vIR = 0
		     vPIS = 0
		     vCOFINS = 0
		     vCSLL = 0
			 VreceitasFederais = 0
			
			vPercISS = IIf(IsNull(RS("percentual_ISS")), 0, RS("percentual_ISS"))
			
			if not rsimpostofederais.eof then
			  
			  vISS = IIf(IsNull(rsimpostofederais("IR")),0,rsimpostofederais("IR"))
			  vIR = IIf(IsNull(rsimpostofederais("IR")),0,rsimpostofederais("IR"))
			  vPIS = IIf(IsNull(rsimpostofederais("PIS")),0,rsimpostofederais("PIS"))
			  vCOFINS = IIf(IsNull(rsimpostofederais("COFINS")),0,rsimpostofederais("COFINS"))
			  vCSLL = IIf(IsNull(rsimpostofederais("CSLL")),0,rsimpostofederais("CSLL"))
			  
			  'wvalor_desconto = (wvalor_parcela * vPercISS) + (wvalor_parcela * vIR) + (wvalor_parcela * vPIS)
			 ' wvalor_desconto = wvalor_desconto + (wvalor_parcela * vCOFINS) + (wvalor_parcela * vCSLL)
			  
			  valorISS = ( wvalor_parcela * vPercISS)
			  
			  valorIR = ( wvalor_parcela * vIR )
			  if valorIR < 10 then
			  valorIR = 0
			  end if 
			  
			  ValorPis = ( wvalor_parcela * vPIS)
			  'if ValorPis < 10 then
			  'ValorPis = 0
			  'end if 
			  
			  ValorCofins =  ( wvalor_parcela * vCOFINS )
			  'if ValorCofins < 10 then
			  'ValorCofins = 0
			  'end if 
			  
			  
			  ValorCSLL = ( wvalor_parcela * vCSLL )
			 ' if ValorCSLL < 10 then
			  'ValorCSLL = 0
			 ' end if 
			  
			 
			 VreceitasFederais = ValorPis + ValorCofins + ValorCSLL
			  
			  if VreceitasFederais < 10 then
			  ValorPis = 0 
			  ValorCofins = 0
			  ValorCSLL = 0
			  end if 
			  
			  
			  wvalor_desconto = valorISS + (valorIR) + (ValorPis)
			  wvalor_desconto = wvalor_desconto + (ValorCofins) + (ValorCSLL)
			
			end if
				
			rsimpostofederais.close	
            wvalor_liquido = wvalor_parcela - wvalor_desconto
		end if 
		
		%>		
	

		<td><div align="center" class="style41"> 
		 <% response.write formatcurrency(valorISS , 3) %>		
		</div></td>
		
		<td><div align="center" class="style41"> 
		 <% response.write formatcurrency(valorIR , 3) %>		
		</div></td>
		
		<td><div align="center" class="style41"> 
		 <% response.write formatcurrency(ValorPis , 3)%>	
		</div></td>
		
		<td><div align="center" class="style41"> 
		 <% response.write formatcurrency(ValorCofins , 3) %>	
		</div></td>
		
		<td><div align="center" class="style41"> 
		<% response.write formatcurrency(ValorCSLL , 3)%>	
		</div></td>
		
		<td><div align="center" class="style41"> 
		 <%
		response.write formatcurrency(wvalor_desconto,3)
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
		
		
    
	   <td><div align="center" class="style41"> 
		 <% response.write (wprocessos) %>		
		</div></td>
        
		
		<td><div align="center"><span class="style4"><%=rs("EMITENOTA")%></span></div></td>
		<td><div align="center"><span class="style4"><%=rs("SUBST_TRIB")%></span></div></td>
		<td><div align="center"><span class="style4"><%=rs("TIPO_TRIBUTACAO")%></span></div></td>
		
		
        
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

If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Number
    WScript.Echo "Error (Hex): " & Hex(Err.Number)
    WScript.Echo "Source: " &  Err.Source
    WScript.Echo "Description: " &  Err.Description
    Err.Clear             ' Clear the Error
End If

    LOOP                'tb paginacao
Set rsConsFecha = nothing
%>
	  
      <tr bgcolor="" class="txt_normal">
        <td height="24">&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td><%=FormatNumber(wvalor_total)%></td>
        <td>&nbsp;</td>
      </tr>
      <tr class="txt_normal">
          <td height="32" colspan="5" class="txt_normal"><strong>total de CONTRATOS:</strong> <strong>
              <% = RS.RECORDCOUNT %>
          </strong></td>
        </tr>
    </table></td>
  </tr>
</table>

</body>
</html>

<!--#include virtual="/registrafacil/include/db_close.asp"-->