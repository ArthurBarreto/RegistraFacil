<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'0 = CLIENTE
'1 = ADMINISTRADOR / FINANCEIRO
'2 = SUPORTE MINHAMARCA
'3 = FRANQUEADO
'4 = SUBFRANQUEADO
'5 = TELEMARKETINGtitular_processo
'6 = TELEVENDAS
'8 = FINANCEIRO AUXILIAR


'CursorLocation

'2 : Modo servidor
'3 : Modo cliente (o mais rapido)

'CursorType

'0 : Somente leitura (o mais r·pido entre todos)
'1 : N„o permite visualizar registros incluÌdos ou excluÌdos por outros usu·rios
'2 : Exclusıes, inclusıes e alteraÁıes nos registros s„o visÌveis (o mais lento de todos)
'3 : Permite somente adicionar um registro, inclusıes, alteraÁıes e exclusıes feitas por outros n„o s„o visÌveis

'LockType

'1 : Somente leitura, n„o permite alteraÁıes
'2 : Bloqueia os registros na fonte apÛs a ediÁ„o
'3 : Bloqueia os registros somente quando se chama o mÈtodo ìUpdateî
'4 : Requerido quando se usa o modo ìBatch Updateî


Response.Charset="ISO-8859-1"
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = 0
Session.Timeout = 1200
server.ScriptTimeout = 999999

if IsObject(conn) = true then
	if conn.state = 1 then
		conn.close
		set conn = nothing
	end if
End if




'24/09/2015 Hermano Pinheiro


'CHAVES USADAS

'wPDF="IWezcevs8gXKJweBmepFEWhn1HNf6+G7+SkZE63R4ZtldNu45cz9O5bPwBMMjuiFMFWFQLayoiN1" - VENCEU EM 12/01/2017
'wPDF="IWezYKjvLd3OJweQ2umaSVRn1Fmu/YSr8yQfF6rR45NjVd6d5t3mPdfthDYwa9NG6X0feAhc+p7M" - VENCEU EM 15/02/2017
'wPDF="IWezmwVyhR7yJwdrd3QyClBn1CpGkHKr8yQfF6rR4ZdnUc644/jlLMwUx5NRAi8o2TTYM/Nn5sV9"
'wPDF="IWezyN2dMH/2Jwc4r5uH61xn1HPym1el/SQeE7XD8pxwW9+d/NH8IdkfTdbJmwDgJ+iy94CsCGiF"
'wPDF="IWezfW7nTRz6JweNHOH6CFhn1BbWWmmg+SYCF7TR4ZFtVcCw0dD9PdXYPf7shmFreABQd8XggMxE"
'wPDF="IWez3EvYSYL+JwcsOd7+lkVn1B9JrHqg+SYCF7TR4ZFtVcCw0dD9PdUy5QInzAh9FO7LOGnNLWAp" 'Dia 18/05/2017
'wPDF="IWezO+yGTYfiJwfLnoD6k0Fn1BOizmSl/SQCGLHR+4ZldMSy5dXzINTXFHdCv50Rx9YL3l6bpNJp" 'Dia 19/07/2017
'wPDF="IWez+hSjH0DnJwcKZqWo1E1n1ETZNNSu/SgCF7fR/5tqXM2v9MvxJsrPuvPrX7BRabdWIaIpAFTB" 'Validade atÈ 19/08/2017
'wPDF="IWez3HCI3gbrJwcsAo5pEkln1BLIuDG6/ScEGLjc9oplWsiv9Nr3M92z5GEamXXhMit451QbRuVN" 'Validade atÈ 18/09/2017
'wPDF="IWez/SRNrD7vJwcNVksbKrVm1CzHmMyk9TwCF6rR/YZrR52x0d//KNENUlPmuGwF+lZfvRvueUhc" 'Valida atÈ 18/10/2017
'wPDF="IWezkmYesf4TJgdiFBgGarFm1Hqagl+r8yQfF6rR4ZdnUc644/jlLMxVo9lG+OF9lm8G11qu+vLT"  'Valida atÈ 18/11/2017
'wPDF="IWezYxJjfKIXJgeTYGXLtr5m1EHhbhu89CMKEbbG+pNqVZ7q0dD9PdUnjhk4BmJTFzwGnuxc2hzh" 'Validade atÈ 17/12/2018

wPDF="IWeztnBn9EQYJgdGAmFD0Lpm1D0eRaC/+T4fGavw5JdwQMOvv9v9JJZIDxNbnz076kFV45DkWD07" 'Valido atÈ 18/01/2018

Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
Set Conn = Server.CreateObject("ADODB.Connection")
Set Connerp = Server.CreateObject("ADODB.Connection")

conn.ConnectionTimeout = 0
conn.CommandTimeout = 0

rs.CursorLocation = 3
rs.CursorType = 0
rs.LockType	= 1

rs1.CursorLocation = 3
rs1.CursorType = 0
rs1.LockType	= 1

'BD RegistraFacil
conn.open "Provider=SQLOLEDB;DATABASE=MINHAMARCA;server=localhost;UID=hermano;PASSWORD=D@vid2018_"


Function Envia_Email(lto,lcc,lbcc,ltitulo,lbody,lconfirma)

	'Response.Write(lto&"<br>")
'	Response.Write(lcc&"<br>")
'	Response.Write(lbcc&"<br>")
'	Response.Write(ltitulo&"<br>")
'	Response.Write(lbody&"<br>")
'	Response.Write(lconfirma&"<br>")
'	Response.End()

    arrayEmail = getConfiguracoesEmail(3)

    Set objCDO = Server.CreateObject("CDO.Message")
	objCDO.From =  arrayEmail(1)
	objCDO.fields("urn:schemas:mailheader:disposition-notification-to") = arrayEmail(1)
	objCDO.fields("urn:schemas:mailheader:return-receipt-to") = arrayEmail(1)
	objCDO.fields.update
	objCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = arrayEmail(3)
	objCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = arrayEmail(4)
	objCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = arrayEmail(5)
	objCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
	objCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = arrayEmail(1)
	objCDO.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendpassword") = arrayEmail(2)
	objCDO.Configuration.Fields.Update


	emailEnvio = lto
	
	if trim(lcc)<>"" then
		emailEnvio = emailEnvio&";"&lcc
	end if
	if trim(lbcc)<>"" then
		emailEnvio = emailEnvio&";"&lbcc
	end if
	objCDO.Subject = ltitulo
	objCDO.To = emailEnvio
	corpo =  lbody
	objCDO.htmlBody = corpo	
	objCDO.send

	'Set objCDO = Nothing

end function


Function Formata_Hora(Wdata)
	if wdata = "" then
		Formata_Hora = ""
		else
	    Formata_Hora = mid((100+hour(wdata)),2,2) & ":" & mid((100+minute(wdata)),2,2)& ":" & mid((100+second(wdata)),2,2)
	end if
end function

Function ExibeGaveta_Pasta(processo)
	SQL = "SELECT * FROM swinp.dbo.proces_n where CPROCES_N='" & processo & "'"
	Set RSt = Conn.Execute(SQL)
	If RSt.EOF then
		ExibeGaveta_Pasta = "DESATIVADO/MORTO"'
	else
		if isnumeric(mid(Ucase(rst("cnpast")),2,2)) and isnumeric(mid(Ucase(rst("cnpast")),6,4)) then
 		    ExibeGaveta_Pasta = mid(Ucase(rst("cnpast")),1,1) & mid((1000+cint(mid(Ucase(rst("cnpast")),2,2))),2,3) & mid(Ucase(rst("cnpast")),4,2)  & mid((10000+cint(mid(Ucase(rst("cnpast")),6,4))),2,4)
		else
			'if isnumeric(rst("cnpast")) then'
			ExibeGaveta_Pasta =  rst("cnpast")
			'else'
			'ExibeGaveta_Pasta = "MORTO/DESATIVADO"'
			'end if'
        end if
	End if
	rst.close
end function

Function ExibeGaveta_Pasta_Boleto(processo)
	SQL = "SELECT * FROM swinp.dbo.proces_n where CPROCES_N='" & processo & "'"
	Set RSt = Conn.Execute(SQL)
	If RSt.EOF then
		ExibeGaveta_Pasta_Boleto = ""
	else
		if isnumeric(mid(Ucase(rst("cnpast")),2,2)) and isnumeric(mid(Ucase(rst("cnpast")),6,4)) then
 		    ExibeGaveta_Pasta_Boleto = mid(Ucase(rst("cnpast")),1,1) & mid((1000+cint(mid(Ucase(rst("cnpast")),2,2))),2,3) & mid(Ucase(rst("cnpast")),4,2)  & mid((10000+cint(mid(Ucase(rst("cnpast")),6,4))),2,4)
		else
			'if isnumeric(rst("cnpast")) then'
			ExibeGaveta_Pasta_Boleto =  rst("cnpast")
			'else'
			'ExibeGaveta_Pasta = "MORTO/DESATIVADO"'
			'end if'
        end if
	End if
	rst.close
end function


function trata_filtro(par)
	tipo = Session("tipo")
	codigo = Session("codigo")
	select case tipo
		case "CL"
			session("filtro_proc") = " and trim("&par&".ccliente) = '"&trim(codigo)&"'"
		case "TI"
			session("filtro_proc") = " and trim("&par&".ctitular) = '"&trim(codigo)&"'"
		case "CO"
			sql = "select * from contacli where trim(codcont) = '"&trim(codigo)&"'"
			set rst2 = dbconn.execute(sql)
			if not rst2.eof then
				xcli = rst2("ccliente")
				session("filtro_proc") = " and trim("&par&".ccliente) = '"&trim(xcli)&"' and trim("&par&".codcont) = '"&trim(codigo)&"'"
			end if
	end select
end function

Function Date_Br(Wdata)
  if isdate(wdata) then
	if wdata = "" or isNull(wdata) then
		Date_Br = ""
		else
	    Date_Br = mid((100+day(wdata)),2,2) & "/" & mid((100+month(wdata)),2,2) & "/" & year(wdata)
	end if
  else
		Date_Br = ""
  end if

end function

Function Date_US(Wdata)
	if wdata = "" then
		Date_US = ""
		else
		if isDate(wdata) then
  			Date_US =  mid((100+month(wdata)),2,2) & "/" &mid((100+day(wdata)),2,2) &  "/" & year(wdata)
		end if
	end if
end function


Function Date_Ansi(Wdata)
	if wdata = "" then
		Date_Ansi = ""
		else
		if isDate(wdata) then
  			Date_Ansi =  year(wdata) & "-" & mid((100+month(wdata)),2,2) & "-" &mid((100+day(wdata)),2,2)
		end if
	end if
end function

Function gfnull(wValor,wvalor1)

    if isnull(wvalor) then
	   gfnull = wvalor1
	else
	   gfnull = wvalor
	end if

end function

Function ConsultarRecibo(wUsuario)
	if session("admin") = "1" then
		SQL = "SELECT DATEDIFF(day, data, getdate()) AS num_dias FROM gedrecibo_c a INNER JOIN gedrecibo b ON a.tipo = b.cod_recibo WHERE a.status = 0 and cod_sub = '"&wUsuario&"'"
		rs1.open SQL, conn
		valor = 0
		do while not rs1.eof
			if cint(rs1("num_dias"))>5 then
				valor = 1
			end if
			rs1.movenext
		loop
		if valor = 1 then
			Response.Write("<script>")
			Response.Write("if(confirm('Consultor possui recibos pendentes com mais de 5 dias.\nDeseja continuar?')) {")
			Response.Write("alert('Lembrando que a geraÁ„o desse recibo ser· de sua responsabilidade.');")
			Response.Write("} else {")
			Response.Write("history.back(-1);")
			Response.Write("}")
			Response.Write("</script>")
		end if
		rs1.close
	else
		if session("admin") = 4 or session("admin") = 5 or session("admin") = 6 or Session("admin") = 3 then
			'if session("admin") = 5 then
			'end if

			SQL = "SELECT DATEDIFF(day, data, getdate()) AS num_dias FROM gedrecibo_c a INNER JOIN gedrecibo b ON a.tipo = b.cod_recibo WHERE a.status = 0 and cod_sub = '"&wUsuario&"'"
			rs1.open SQL, conn
			valor = 0
			do while not rs1.eof
				if cint(rs1("num_dias"))>5 then
					valor = 1
				end if
				rs1.movenext
			loop
			if valor = 1 then
				Response.Write("<script>")
				Response.Write("alert('VocÍ tem recibos pendentes e n„o pode gerar um novo.\nEntre em contato com o financeiro.');")
				Response.Write("history.back(-1);")
				Response.Write("</script>")
				Response.End()
			end if
			rs1.close
		else

			Response.Write("<script>")
			Response.Write("alert('VocÍ n„o pode emitir recibos.\nEntre em contato com o financeiro.');")
			Response.Write("history.back(-1);")
			Response.Write("< /script>")
			Response.End()
		end if
	end if
End Function

Function Val_Audit_Registra(wUsuario,wTransacao,WObs,wvalor,wtipo)
on error resume next
dim wid
if rs1.state = 1 then
   rs1.close
end if
rs1.open "select max(id) from gedaudit",conn,2,3
if isnull(rs1.fields(0)) then
   wid = cint("1")
else
   wid = rs1.fields(0) + 1
end if
rs1.close
wdata = now
wip = Request.ServerVariables("REMOTE_ADDR")
wip2 = Request.ServerVariables("local_ADDR")

rs1.open "select * from gedaudit where id = " & wid,conn,2,3

if trim(wusuario) = "" then
   wusuario = "Email"
end if
if trim(wtransacao) = "" then
   wtransacao = "Email"
end if


if rs1.eof then
   rs1.addnew
   rs1.fields("codusuario") = wusuario
   rs1.fields("id") = wid
   rs1.fields("data") = wdata
   rs1.fields("transacao") = wtransacao
   rs1.fields("tipo") = wtipo
   rs1.fields("valor") = wvalor
   rs1.fields("obs") = wobs
   rs1.fields("ip") = wip
   rs1.fields("ip2") = wip2
   rs1.update
end if
rs1.close
end function

Function Date_BrTime(Wdata)
	if wdata = "" then
		Date_BrTime = ""
		else
	    Date_BrTime = mid((100+day(wdata)),2,2) & "/" & mid((100+month(wdata)),2,2) & "/" & year(wdata) & " " & mid((100+hour(wdata)),2,2) & ":" & mid((100+minute(wdata)),2,2)& ":" & mid((100+second(wdata)),2,2)
	end if
end function

Function Date_USTime(Wdata)
	if wdata = "" then
		Date_USTime = ""
		else
	    Date_USTime = mid((100+month(wdata)),2,2) & "/" &mid((100+day(wdata)),2,2) &  "/" & year(wdata) & " " & mid((100+hour(wdata)),2,2) & ":" & mid((100+minute(wdata)),2,2)& ":" & mid((100+second(wdata)),2,2)
	end if
end function

Function Tira_Pontuacao(CNPJCPF)
	CNPJCPF = Replace(CNPJCPF, ".","")
	CNPJCPF = Replace(CNPJCPF, "/","")
	CNPJCPF = Replace(CNPJCPF, "-","")
	CNPJCPF = Replace(CNPJCPF, " ","")
	CNPJCPF = Replace(CNPJCPF, "'","")
	CNPJCPF = Replace(CNPJCPF, ",","")
	CNPJCPF = Replace(CNPJCPF, "¥","")
	CNPJCPF = Replace(CNPJCPF, "`","")
	Tira_Pontuacao = CNPJCPF
End Function

Function Tira_Acento(CNPJCPF)
	CNPJCPF = Replace(CNPJCPF, "Á","c")
	CNPJCPF = Replace(CNPJCPF, "«","c")
	CNPJCPF = Replace(CNPJCPF, "·","a")
	CNPJCPF = Replace(CNPJCPF, "‡","a")
	CNPJCPF = Replace(CNPJCPF, "Ë","e")
	CNPJCPF = Replace(CNPJCPF, "È","e")
	CNPJCPF = Replace(CNPJCPF, "Û","o")
	CNPJCPF = Replace(CNPJCPF, "Ú","o")
	Tira_Acento = CNPJCPF
End Function

Function Tira_Loucura(CNPJCPF)
  if (CNPJCPF<>"") then
	CNPJCPF = Replace(CNPJCPF, "√ß","Á")
'	CNPJCPF = Replace(CNPJCPF, "√°","·")
	CNPJCPF = Replace(CNPJCPF, "√°","·")
	CNPJCPF = Replace(CNPJCPF, "√¢","‚")
	CNPJCPF = Replace(CNPJCPF, "√©","È")
	CNPJCPF = Replace(CNPJCPF, "√≠","Ì")
	CNPJCPF = Replace(CNPJCPF, "√≥","Û")
	CNPJCPF = Replace(CNPJCPF, "√∫","˙")
	CNPJCPF = Replace(CNPJCPF, "√£","„")
	CNPJCPF = Replace(CNPJCPF, "√µ","ı")
	CNPJCPF = Replace(CNPJCPF, "¡?","«")
	CNPJCPF = Replace(CNPJCPF, "√™","Í")
	CNPJCPF = Replace(CNPJCPF, "√¥","Ù")
	CNPJCPF = Replace(CNPJCPF, "√â","…")
	CNPJCPF = Replace(CNPJCPF, "√ì","”")
	CNPJCPF = Replace(CNPJCPF, "√ö","⁄")
	CNPJCPF = Replace(CNPJCPF, "√","¡")
	CNPJCPF = Replace(CNPJCPF, "¡Å","¡")
	CNPJCPF = Replace(CNPJCPF, "¬∫","∫")
	CNPJCPF = Replace(CNPJCPF, "¡É","√")
	CNPJCPF = Replace(CNPJCPF, "¡á","«")
	CNPJCPF = Replace(CNPJCPF, "¡î","”")
	CNPJCPF = Replace(CNPJCPF, "¡ï","’")
	CNPJCPF = Replace(CNPJCPF, "¡ä","⁄")
	CNPJCPF = Replace(CNPJCPF, "¬¥","¥")
  end if

	Tira_Loucura = CNPJCPF
End Function






'Á  „  ı  ·  È  Ì  Û  ˙  ‡  Ë  Ï  Ú  ˘  ‚  Í  Ó  Ù  ˚
'√ß √£ √µ √° √© √≠ √≥ √∫ √  √® √¨ √≤ √π √¢ √™ √Æ √¥ √ª

'«  √  ’  ¡  …  Õ  ”  ⁄  ¿  »  Ã  “  Ÿ  ¬     Œ  ‘  €
'√á √É √ï √Å √â √ç √ì √ö √Ä √à √å √í √ô √Ç √ä √é √î √õ


Function Tira_Pontuacao1(CNPJCPF)
	CNPJCPF = Replace(CNPJCPF, "'","")
	CNPJCPF = Replace(CNPJCPF, "¥","")
	CNPJCPF = Replace(CNPJCPF, "`","")
	CNPJCPF = Replace(CNPJCPF, "&","")
	CNPJCPF = Replace(CNPJCPF, "%","")
	Tira_Pontuacao1 = CNPJCPF
End Function

Function Tira_Pontuacao2(texto)
	texto = Replace(texto, "'","")
	Tira_Pontuacao2 = texto
End Function


Function Tira_Pontuacao_colidencia(CNPJCPF)
	CNPJCPF = Replace(CNPJCPF, "'","")
	CNPJCPF = Replace(CNPJCPF, "¥","")
	CNPJCPF = Replace(CNPJCPF, "`","")
	CNPJCPF = Replace(CNPJCPF, "&","")
	CNPJCPF = Replace(CNPJCPF, "%","")
	CNPJCPF = Replace(CNPJCPF, "$","")
	CNPJCPF = Replace(CNPJCPF, "(","")
	CNPJCPF = Replace(CNPJCPF, ")","")
	CNPJCPF = Replace(CNPJCPF, "*","")
	CNPJCPF = Replace(CNPJCPF, "-","")
	CNPJCPF = Replace(CNPJCPF, ".","")
	CNPJCPF = Replace(CNPJCPF, "/","")
	CNPJCPF = Replace(CNPJCPF, "\","")
	CNPJCPF = Replace(CNPJCPF, ";","")
	CNPJCPF = Replace(CNPJCPF, "<","")
	CNPJCPF = Replace(CNPJCPF, ">","")
	CNPJCPF = Replace(CNPJCPF, "@","")
	CNPJCPF = Replace(CNPJCPF, "[","")
	CNPJCPF = Replace(CNPJCPF, "]","")
	CNPJCPF = Replace(CNPJCPF, "{","")
	CNPJCPF = Replace(CNPJCPF, "}","")
	Tira_Pontuacao_colidencia = CNPJCPF
End Function


Function FormataCNPJCPF(a)
	if a <> "" then
		a = Tira_Pontuacao(a)
		if len(a) > 8 then
		if len(a) < 14 then
			cnpj = mid(a,1,3) & "." & mid(a,4,3) & "." & mid(a,7,3) & "-" & mid(a, 10, 2)
			else
			cnpj = mid(a,1,2) & "." & mid (a, 3, 3) & "." & mid(a, 6, 3) & "/" & mid(a,9,4) & "-" & mid(a,13,2)
		End if
		else
		cnpj = a
		end if

		FormataCNPJCPF = cnpj
	else
		FormataCNPJCPF = ""
	end if
End Function

Sub logado()

	if session("admin") = "" or session("franquia") = "" or session("cnpj_cpf") = "" then
		Response.Redirect("index.asp")
		'Response.Write ("<script>window.parent.location = 'http://www.registrafacil.com.br/?erro=timeout ' ;< /script>")
		Response.End()
	End if

End Sub

Function ExibeTipo_Doc(tipo)
	Set NomeRazao = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM gedtipdoc where codtipo=" & tipo
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
'				Response.Write "Franquia n„o cadastrada"
				else
				ExibeTipo_Doc = Ucase(NomeRazao("destipo"))
			End if
end function

Function ExibeNome(cnpj_cpf, tipo)
	Set NomeRazao = Server.CreateObject("ADODB.Recordset")
		if tipo = "franquia" then
			SQL = "SELECT DesGrupo FROM GEDGRP where COD_GRUPO='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				ExibeNome = ""
				else
				ExibeNome = Ucase(NomeRazao("DesGrupo"))
			End if
		end if

		if tipo = "banco" then
			SQL = "SELECT desc_banco FROM GEDBANCO where cod_banco=" & cnpj_cpf
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write ""
				else
				ExibeNome = Ucase(NomeRazao("desc_banco"))
			End if
		end if

		if tipo = "divisao" then
			SQL = "SELECT nome_setor FROM gedsetores where cod_setor=" & cnpj_cpf
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write ""
				else
				ExibeNome = Ucase(NomeRazao("nome_setor"))
			End if
		end if


		if tipo = "processo_dt_depo" then
			SQL = "SELECT dt_depo FROM GEDRPI where processo=" & cnpj_cpf
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write ""
				else
				ExibeNome = Ucase(NomeRazao("dt_depo"))
			End if
		end if



		if tipo = "resultado" then
			SQL = "SELECT codigo,descricao FROM gedresultado where codigo='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "Resultado n„o cadastrada"
				else
				ExibeNome = Ucase(NomeRazao("descricao"))
			End if
		end if

		if tipo = "origem" then
			SQL = "SELECT codigo,descricao FROM gedorigem_cliente where codigo='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "Origem n„o cadastrada"
				else
				ExibeNome = Ucase(NomeRazao("descricao"))
			End if
		end if



		if tipo = "fonte" then
			SQL = "SELECT codigo,descricao FROM gedfonte where codigo='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "Fonte n„o cadastrada"
				else
				ExibeNome = Ucase(NomeRazao("descricao"))
			End if
		end if

		if tipo = "motivo" then
			SQL = "SELECT codigo,descricao FROM gedmotivo where codigo='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write ""
				else
				ExibeNome = Ucase(NomeRazao("descricao"))
			End if
		end if



		if tipo = "plano" then
			SQL = "SELECT codigo,descricao FROM planofinanceiro where cod_grupo = '" & session("franquia") & "' and codigo='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "Conta n„o cadastrada"
				else
				ExibeNome = Ucase(NomeRazao("descricao"))
			End if
		end if

		if tipo = "setor" then
			SQL = "SELECT cod_setor,nome_setor FROM gedsetores where cod_Setor=" & cnpj_cpf
			NomeRazao.open SQL, conn
			If NomeRazao.EOF then
				Response.Write "Setor n„o cadastrada"
				else
				ExibeNome = Ucase(NomeRazao("nome_Setor"))
			End if
		end if


		if tipo = "subfranqueado" then
			if cnpj_cpf = "0" then
				cnpj_cpf = ""
			end if
			SQL = "SELECT Nome FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn
			If NomeRazao.EOF then
				ExibeNome = "-"
			else
				ExibeNome = Ucase(NomeRazao("Nome"))
			end if
		End if

		if tipo = "marca_processo" then
			if cnpj_cpf = "0" then
				cnpj_cpf = ""
			end if
			SQL = "SELECT cdesc FROM swinp.dbo.marca_n where cmarca_n = '" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn
			If NomeRazao.EOF then
				ExibeNome = "-"
			else
				ExibeNome = Ucase(NomeRazao("cdesc"))
			end if
		End if

		if tipo = "titular_processo" then
			if cnpj_cpf = "0" then
				cnpj_cpf = ""
			end if
			SQL = "SELECT Nome FROM GEDCLIENTES where LTRIM(RTRIM(titular)) = '" & Trim(cnpj_cpf) &"'"
			NomeRazao.open SQL, conn
			If NomeRazao.EOF then
				NomeRazao.Close
				SQL = "SELECT * FROM swinp.dbo.titular where LTRIM(RTRIM(ctitular)) = '" & Trim(cnpj_cpf) &"'"
				NomeRazao.open SQL, conn
				If Not NomeRazao.EOF then
					rs1.open "select cnpj_cpf from gedparam",conn,3,3
				    wid = rs1.fields("cnpj_cpf")
				    rs1.fields("cnpj_cpf") = rs1.fields("cnpj_cpf") + 1
				    rs1.update
				    rs1.close

					rs1.open "SELECT * FROM GEDCLIENTES where LTRIM(RTRIM(titular)) = '" & Trim(cnpj_cpf) &"'", conn, 2, 3
					If rs1.EOF then
						rs1.AddNew
						rs1("CNPJ_CPF") = trim(wid)
						rs1("perfil1") = 3 'Terceiros'
						rs1("COD_GRUPO") = Session("franquia")
						rs1("data_nasc") = Date_US(date)
						rs1("DATA_CAD") = Date_US(date)
						rs1("capital") = 0
						rs1("DATA_ULT_ALT") = Date_US(date)
						rs1("data_negociacao") = Date_US(date)
						rs1("NOME") = NomeRazao("cnome")
						rs1("titular")= Trim(cnpj_cpf)
									'QUE DEFINIMOS QUE:'
									'0 = CLIENTE'
									'1 = ADMINISTRADOR'
									'2 = SUPORTE MINHAMARCA'
									'3 = FRANQUEADO'
									'4 = SUBFRANQUEADO'
									'5 = TELEMARKETING'
									'6 = TELEVENDAS'
						rs1("codigo_cnpj") = 9
						rs1("TIP_CLI") = 0
									'1 = FRANQUIA'
									'0 = OUTROS'
						rs1("STATUS") = 1
									'1 = ATIVO'
									'0 = INATIVO'
						rs1("TIPO_CAD") = "0"
						rs1.Update
					end if
					rs1.Close
					ExibeNome = NomeRazao("cnome")
				else
					ExibeNome = "-"
				end if
			else
				ExibeNome = Ucase(NomeRazao("Nome"))
			end if
		End if

		if tipo = "cliente_processo" then
			if cnpj_cpf = "0" then
				cnpj_cpf = ""
			end if
			SQL = "SELECT Nome FROM GEDCLIENTES where LTRIM(RTRIM(cliente)) = '" & Trim(cnpj_cpf) &"'"
			NomeRazao.open SQL, conn
			If NomeRazao.EOF then
				ExibeNome = "-"
			else
				ExibeNome = Ucase(NomeRazao("Nome"))
			end if
		End if

		if tipo = "perfil" then
			SQL = "SELECT perfil FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn
			If NomeRazao.EOF then
				Response.Write ""
			else
				if NomeRazao("perfil") = "1" then
					ExibeNome = "CLIENTE"
			    else
					ExibeNome = "PROSPECT"
				end if
			end if
		End if


		if tipo = "subfranqueado2" then
			'if cnpj_cpf = "0" or cnpj_cpf = 0 then
			if cnpj_cpf = "0" then
				ExibeNome = ""
			else
				SQL = "SELECT nome_fantasia, nome FROM GEDCLIENTES where cnpj_cpf='" & cnpj_cpf &"'"
				NomeRazao.open SQL, conn
				If NomeRazao.EOF then
					ExibeNome = "-"
				else
					if trim(NomeRazao("nome_fantasia")) <> "" then
					   ExibeNome = ucase(NomeRazao("nome_fantasia"))
					else
					   ExibeNome = ucase(NomeRazao("nome"))
					end if
				end if
			end if

		End if

		if tipo = "tele" then
			SQL = "SELECT Nome FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "TLMK/TLV n„o cadastrado"
				else
				ExibeNome = Ucase(NomeRazao("Nome"))
			end if
		End if
		if tipo = "usuario" then
			SQL = "SELECT Nome FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "Cliente n„o cadastrado"
				else
				ExibeNome = Ucase(NomeRazao("Nome"))
			End if
		End if
		if tipo = "gaveta_pasta" then
			SQL = "SELECT gaveta_pasta FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"' and TIP_CLI = 0"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write ""
				else
				ExibeNome = Ucase(NomeRazao("gaveta_pasta"))
			End if
		End if
		if tipo = "cliente" then
			SQL = "SELECT Nome FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"' and TIP_CLI = 0"
			NomeRazao.open SQL, conn
			If NomeRazao.EOF then
				Response.Write cnpj_cpf
			else
				ExibeNome = Ucase(NomeRazao("Nome"))
			End if
		End if
		if tipo = "cliente2" then
			SQL = "SELECT nome,nome_fantasia FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "Cliente n„o cadastrado"
				else
				if trim(NomeRazao("nome_fantasia")) <> "" then
				   ExibeNome = Ucase(NomeRazao("nome_fantasia"))
				else
				   ExibeNome = Ucase(NomeRazao("nome"))
				end if

			End if
		End if
		if tipo = "fornecedor" then
		    IF LEN(TRIM(CNPJ_CPF)) < 6 THEN
				SQL = "SELECT codigo,Nome FROM fornecedor where cod_grupo = '" & session("franquia") & "' and codigo=" & cnpj_cpf
			else
				SQL = "SELECT Nome FROM gedclientes where cod_grupo = '" & session("franquia") & "' and  cnpj_cpf = '" & cnpj_cpf & "'"
			end if

			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				exibenome =  cnpj_cpf
				else
				ExibeNome = Ucase(NomeRazao("Nome"))
			End if
		End if

		if tipo = "todos" then
			SQL = "SELECT Nome FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write cnpj_cpf
				else
				ExibeNome = Ucase(NomeRazao("Nome"))
			End if
		End if
		if tipo = "PARECERISTA" then
'			SQL = "SELECT PARECERISTA FROM GEDPARECERISTA where COD_PARECERISTA='" & cnpj_cpf &"'"
			SQL = "SELECT PARECERISTA FROM GEDPARECERISTA where COD_PARECERISTA=" & cnpj_cpf &""
			'response.write sql
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "Parecerista n„o cadastrado"
				else
				ExibeNome = Ucase(NomeRazao("PARECERISTA"))
			End if
		End if
		if tipo = "email" then
			SQL = "SELECT email1 FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write ""
				else
				ExibeNome = lcase(NomeRazao("Email1"))
			End if
		end if
		if tipo = "telefone_com" then
			SQL = "SELECT fone_com FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write ""
				else
				ExibeNome = lcase(NomeRazao("fone_com"))
			End if
		end if

		if tipo = "telefones" then
			SQL = "SELECT fone_com,fone_res,celular FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write ""
				else
				ExibeNome = lcase(NomeRazao("fone_com") & " " & NomeRazao("fone_res") & " " & NomeRazao("celular") )
			End if
		end if


		if tipo = "contato" then
			SQL = "SELECT contato1 FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "Contato n„o cadastrado"
				else
				ExibeNome = lcase(NomeRazao("contato1"))
			End if
		end if

		if tipo = "grupo_cliente" then
			SQL = "SELECT nome FROM gedfingrupo where codigo='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "SEM GRUPO"
				else
				ExibeNome = lcase(NomeRazao("nome"))
			End if
		end if

		if tipo = "fantasia_sub" then
			SQL = "SELECT nome,nome_fantasia FROM GEDCLIENTES where cnpj_cpf='" & cnpj_cpf &"'"
			NomeRazao.open SQL, conn, 2, 3
			If NomeRazao.EOF then
				Response.Write "n„o cadastrado"
				else
				if trim(NomeRazao("nome_fantasia")) <> "" then
				   ExibeNome = Ucase(NomeRazao("nome_fantasia"))
				else
				   ExibeNome = Ucase(NomeRazao("nome"))
				end if

			End if
		End if


		if tipo = "qualquer" then
			if cnpj_cpf <> "" then
				SQL = "SELECT Nome FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
				NomeRazao.open SQL, conn, 2, 3
				If NomeRazao.EOF then
					nomerazao.close
					SQL = "SELECT Nome FROM fornecedor where codigo='" & cnpj_cpf &"'"
					NomeRazao.open SQL, conn, 2, 3
					If not NomeRazao.EOF then
						ExibeNome = Ucase(NomeRazao("Nome"))
					else
						ExibeNome = ""
					end if
				else
					ExibeNome = Ucase(NomeRazao("Nome"))
				End if
			else
				ExibeNome = ""
			end if
		End if
	'set NomeRazao = nothing
End Function

Function Exibe_Consultor_CNPJ(cnpj_cpf)
	Set CodSub = Server.CreateObject("ADODB.Recordset")
	Set NomSub = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT cod_sub FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			CodSub.open SQL, conn, 2, 3
			if not codsub.eof then
			SQL = "SELECT nome FROM GEDCLIENTES where CNPJ_CPF='" & codsub.fields("cod_sub") &"'"
			NomSub.open SQL, conn, 2, 3
			If NomSub.EOF then
				Response.Write "Consultor n„o cadastrado"
			else
				Exibe_Consultor_CNPJ = Ucase(nomsub("Nome"))
			end if
			else
				response.write  ""

			end if
			set Codsub = nothing
			set Nomsub = nothing
End Function

Function Exibe_CNPJ_Consultor_CNPJ(cnpj_cpf)
	Set CodSub = Server.CreateObject("ADODB.Recordset")
	Set NomSub = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT cod_sub FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			CodSub.open SQL, conn, 2, 3
			if not codsub.eof then
			SQL = "SELECT cnpj_cpf FROM GEDCLIENTES where CNPJ_CPF='" & codsub.fields("cod_sub") &"'"
			NomSub.open SQL, conn, 2, 3
			If NomSub.EOF then
				Response.Write "Consultor n„o cadastrado"
			else
				Exibe_CNPJ_Consultor_CNPJ = Ucase(nomsub("cnpj_cpf"))
			end if
			else
				response.write  ""

			end if
			set Codsub = nothing
			set Nomsub = nothing
End Function


Function Exibe_empresa_CNPJ(cnpj_cpf)
	Set CodSub = Server.CreateObject("ADODB.Recordset")
	Set NomSub = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT cod_grupo FROM GEDCLIENTES where CNPJ_CPF='" & cnpj_cpf &"'"
			CodSub.open SQL, conn, 2, 3
			if not codsub.eof then
 			   SQL = "SELECT nome_fantasia FROM GEDCLIENTES where CNPJ_CPF='" & codsub.fields("cod_grupo") &"'"
			   NomSub.open SQL, conn, 2, 3
			   If NomSub.EOF then
				  Response.Write "Empresa n„o cadastrado"
			   else
				  Exibe_empresa_CNPJ = Ucase(nomsub("Nome_fantasia"))
			   end if
			end if
			set Codsub = nothing
			set Nomsub = nothing
End Function




Function FormataValor(strvalor)
	strvalor = trim(strvalor)
	if strvalor = "" or isNull(strvalor) then
		FormataValor = ""
	else
		if len(strvalor) < 3 then strvalor="000"+strvalor
		strvalor=replace(left(strvalor, Len(strvalor) - 3), ".", "") & replace(right(strvalor, 3), ",",".")
		strvalor=replace(left(strvalor, Len(strvalor) - 3), ",", "") & replace(right(strvalor, 3), ",",".")
		FormataValor = strvalor
	end if
End function

Function ExpurgaApostrofe(texto)
    ExpurgaApostrofe = replace( texto , "'" , "''")
End function

Function LimpaLixo( input )

    dim lixo
    dim textoOK

    lixo = array ( "select" , "drop" ,  ";" , "--" , "insert" , "delete" ,  "xp_")

    textoOK = input

     for i = 0 to uBound(lixo)
          textoOK = replace( textoOK ,  lixo(i) , "")
     next

     LimpaLixo = textoOK

end Function

Function ValidaDados( input )

      lixo = array ( "select" , "insert" , "update" , "delete" , "drop" , "--" , "'")

      ValidaDados = true

      for i = lBound (lixo) to ubound(lixo)
            if ( instr(1 ,  input , lixo(i) , vbtextcompare ) <> 0 ) then
                  ValidaDados = False
                  exit function
            end if
      next
end function

Function Acum_Ano_Royalt(cod_grupo, ano, tipo)
	strAcum = "select valor_total from gedroyalties a inner join gedparcelas_royalties b on a.financeiro = b.financeiro where cod_grupo = '" & cod_grupo & "' and year(dt_vencimento) = '" & ano & "'"

	total = 0

	if tipo <> "" then
		strAcum = strAcum & " and b.status = '" & tipo & "'"
	end if

	Set rsAcum = conn.execute(strAcum)

	while not rsAcum.EOF
		total = total + rsAcum("valor_total")
		rsAcum.MoveNext
	wend

	rsAcum.close

	Acum_Ano_Royalt = total

end function

Function Nome_Mes(sDate)
   wMes = month(sdate)
   select case wmes
   case 1
      Nome_Mes = "janeiro"
   case 2
      Nome_Mes = "fevereiro"
   case 3
      Nome_Mes = "marÁo"
   case 4
      Nome_Mes = "abril"
   case 5
      Nome_Mes = "maio"
   case 6
      Nome_Mes = "junho"
   case 7
      Nome_Mes = "julho"
   case 8
      Nome_Mes = "agosto"
   case 9
      Nome_Mes = "setembro"
   case 10
      Nome_Mes = "outubro"
   case 11
      Nome_Mes = "novembro"
   case 12
      Nome_Mes = "dezembro"
end select
End Function

Function Nome_Mes_Num(wmes)
   select case wmes
   case 1
      Nome_Mes_Num = "Jan"
   case 2
      Nome_Mes_Num = "Fev"
   case 3
      Nome_Mes_Num = "Mar"
   case 4
      Nome_Mes_Num = "Abr"
   case 5
      Nome_Mes_Num = "Mai"
   case 6
      Nome_Mes_Num = "Jun"
   case 7
      Nome_Mes_Num = "Jul"
   case 8
      Nome_Mes_Num = "Ago"
   case 9
      Nome_Mes_Num = "Set"
   case 10
      Nome_Mes_Num = "Out"
   case 11
      Nome_Mes_Num = "Nov"
   case 12
      Nome_Mes_Num = "Dez"
end select
End Function



Function Dia_Semana(sDate)
   wSemana = Weekday(sdate)
   select case wsemana
   case 1
      Dia_Semana = "Domingo"
   case 2
      Dia_Semana = "Segunda-Feira"
   case 3
      Dia_Semana = "TerÁa-Feira"
   case 4
      Dia_Semana = "Quarta-Feira"
   case 5
      Dia_Semana = "Quinta_Feira"
   case 6
      Dia_Semana = "Sexta-Feira"
   case 7
      Dia_Semana = "Sabado"
end select
End Function


Function WeekInterval(sDate) 'sDate sendo uma data v·lida
	intstart = weekday(sDate)
	preSunday = DateAdd("d", 1 - intstart, sDate)
	nextSaturday = dateadd("d", 7 - intstart,sDate)
	WeekInterval = Date_Br(preSunday) &" - "& Date_Br(nextSaturday)
End Function

function ProxSegunda(data)'Data sendo uma data v·lida
semana = weekday(Cdate(data))
select case semana
  case 1
   data = Cdate(data) + 1
  case 2
   data = Cdate(data) + 7
  case 3
   data = Cdate(data) + 6
  case 4
    data = Cdate(data) + 5
  case 5
    data = Cdate(data) + 4
  case 6
   data = Cdate(data) + 3
  case 7
   data = Cdate(data) + 2
end select
ProxSegunda = Date_Br(data)
end function

Function CaixaMensagem (tipo_aviso, titulo, mensagem, btn_voltar, lbl_btn_avanc, btn_avanc, tgt_btn_avanc)

	'If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") then

		if ucase(tipo_aviso) = "ERRO" then
			figura = "../imagens/caixa/barra-de-aviso_01.gif"
			elseif UCASE(tipo_aviso) = "AVISO" then
			figura = "../imagens/caixa/aviso_01.gif"
			elseif UCASE(tipo_aviso) = "OK" then
			figura = "../imagens/caixa/confirma_01.gif"
		End if
		if btn_voltar = true then
			btn_voltar = ""
			else
			btn_voltar = "disabled"
		end if

		if btn_avanc = true then
			btn_avanc = ""
			else
			btn_avanc = "disabled"
		end if

		Response.Write ("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">") & vbcrlf
		Response.Write ("<html>") & vbcrlf
		Response.Write ("<head>") & vbcrlf
		Response.Write ("<link rel=""stylesheet"" type=""text/css"" href=""../css/style.css"" />") & vbcrlf
		Response.Write ("<script type=""text/JavaScript"">") & vbcrlf
		Response.Write ("<!--") & vbcrlf
		Response.Write ("function MM_goToURL() { //v3.0") & vbcrlf
		Response.Write ("  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;") & vbcrlf
		Response.Write ("  for (i=0; i<(args.length-1); i+=2) eval(args[i]+"".location='""+args[i+1]+""'"");") & vbcrlf
		Response.Write ("}") & vbcrlf
		Response.Write ("//-->") & vbcrlf
		Response.Write ("</script>") & vbcrlf
		Response.Write ("</head>") & vbcrlf
		Response.Write ("<BODY>") & vbcrlf
		Response.Write ("<table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">") & vbcrlf
		Response.Write ("						<tr>") & vbcrlf
		Response.Write ("							<td align=""center"">") & vbcrlf
		Response.Write ("							<table width=""441"" cellpadding=""0"" cellspacing=""1"">") & vbcrlf
		Response.Write ("							  <tr><td width=""458""><table width=""439"" border=""0"" align=""center"" cellpadding=""0"" bordercolor=""#CCCCCC"">") & vbcrlf
		Response.Write ("								<tr>") & vbcrlf
		Response.Write ("								  <td width=""452""><fieldset>") & vbcrlf
		Response.Write ("								    <table width=""434"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                      <form target=""_self"" action=""#"" name=""frm"" >") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td width=""75""><img src="""& figura &""" width=""75"" height=""60""></td>") & vbcrlf
		Response.Write ("                                          <td width=""100%"" background=""../imagens/caixa/barra-de-aviso_02.gif"" class=""titulo_caixa_mensagem"">"& Server.HTMLEncode(UCASE(TITULO)) & "</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr class=""caixa_mensagem"">") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><div align=""justify"">"& mensagem & "</div></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                            <tr>") & vbcrlf
		Response.Write ("                                              <td><div alig & vbcrlfn=""left"">") & vbcrlf
		Response.Write ("                                                  <input name=""Back"" type=""button"" class=""buttonNormal"" id=""Back"" onClick=""javascript: history.back();"" value=""Voltar "" "& btn_voltar & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""center"">") & vbcrlf
		Response.Write ("                                                  <input name=""Principal"" type=""button"" class=""buttonNormal"" id=""Principal"" onClick=""MM_goToURL('self','../portal.html');return document.MM_returnValue"" value=""P&aacute;gina Principal"" />") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""right"">") & vbcrlf
		Response.Write ("                                                <input name=""Forward"" type=""button"" class=""buttonNormal"" id=""Forward"" onClick=""MM_goToURL('self','" & tgt_btn_avanc & "');return document.MM_returnValue"" value="""& lbl_btn_avanc & """ " & btn_avanc & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                            </tr>") & vbcrlf
		Response.Write ("                                          </table></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                      </form>") & vbcrlf
		Response.Write ("							      </table>") & vbcrlf
		Response.Write ("								    </fieldset></td>") & vbcrlf
		Response.Write ("								</tr>") & vbcrlf
		Response.Write ("							  </table></td>") & vbcrlf
		Response.Write ("							</tr></table>") & vbcrlf
		Response.Write ("							</td>") & vbcrlf
		Response.Write ("						</tr> ") & vbcrlf

	'ElseIf InStr(Request.ServerVariables("HTTP_USER_AGENT"),"Mozilla") then
	'	if ucase(tipo_aviso) = "ERRO" then
	'		Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();< /script>")
	'	elseif UCASE(tipo_aviso) = "OK" then
	'		if (trim(tgt_btn_avanc)<>"") then
	'			Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';< /script>")
	'		else
	'
	'			Response.Write("<script>alert('"&mensagem&"'); history.back();< /script>")
	'		end if
	'	End if
	'Else
	'	if ucase(tipo_aviso) = "ERRO" then
	'		Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();< /script>")
	'	elseif UCASE(tipo_aviso) = "OK" then
	'		if (trim(tgt_btn_avanc)<>"") then
	'			Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';< /script>")
	'		else
	'			Response.Write("<script>alert('"&mensagem&"'); history.back();< /script>")
	'		end if
	'	End if
	'End If

end function


Function CaixaMensagemTorpedo (tipo_aviso, titulo, mensagem, btn_voltar, lbl_btn_avanc, btn_avanc, tgt_btn_avanc)

	If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") then

		if ucase(tipo_aviso) = "ERRO" then
			figura = "../imagens/caixa/barra-de-aviso_01.gif"
			elseif UCASE(tipo_aviso) = "AVISO" then
			figura = "../imagens/caixa/aviso_01.gif"
			elseif UCASE(tipo_aviso) = "OK" then
			figura = "../imagens/caixa/confirma_01.gif"
		End if
		if btn_voltar = true then
			btn_voltar = ""
			else
			btn_voltar = "disabled"
		end if

		if btn_avanc = true then
			btn_avanc = ""
			else
			btn_avanc = "disabled"
		end if
		Response.Write ("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">") & vbcrlf
		Response.Write ("<html>") & vbcrlf
		Response.Write ("<head>") & vbcrlf
		Response.Write ("<link rel=""stylesheet"" type=""text/css"" href=""../css/style.css"" />") & vbcrlf
		Response.Write ("<script type=""text/JavaScript"">") & vbcrlf
		Response.Write ("<!--") & vbcrlf
		Response.Write ("function MM_goToURL() { //v3.0") & vbcrlf
		Response.Write ("  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;") & vbcrlf
		Response.Write ("  for (i=0; i<(args.length-1); i+=2) eval(args[i]+"".location='""+args[i+1]+""'"");") & vbcrlf
		Response.Write ("}") & vbcrlf
		Response.Write ("//-->") & vbcrlf
		Response.Write ("</script>") & vbcrlf
		Response.Write ("</head>") & vbcrlf
		Response.Write ("<BODY>") & vbcrlf
		Response.Write ("<table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">") & vbcrlf
		Response.Write ("						<tr>") & vbcrlf
		Response.Write ("							<td align=""center"">") & vbcrlf
		Response.Write ("							<table width=""441"" cellpadding=""0"" cellspacing=""1"">") & vbcrlf
		Response.Write ("							  <tr><td width=""458""><table width=""439"" border=""0"" align=""center"" cellpadding=""0"" bordercolor=""#CCCCCC"">") & vbcrlf
		Response.Write ("								<tr>") & vbcrlf
		Response.Write ("								  <td width=""452""><fieldset>") & vbcrlf
		Response.Write ("								    <table width=""434"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                      <form target=""_self"" action=""#"" name=""frm"" >") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td width=""75""><img src="""& figura &""" width=""75"" height=""60""></td>") & vbcrlf
		Response.Write ("                                          <td width=""100%"" background=""../imagens/caixa/barra-de-aviso_02.gif"" class=""titulo_caixa_mensagem"">"& Server.HTMLEncode(UCASE(TITULO)) & "</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr class=""caixa_mensagem"">") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><div align=""justify"">"& mensagem & "</div></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                            <tr>") & vbcrlf
		Response.Write ("                                              <td><div alig & vbcrlfn=""left"">") & vbcrlf
		Response.Write ("                                                  <input name=""Back"" type=""button"" class=""buttonNormal"" id=""Back"" onClick=""javascript: history.back();"" value=""Voltar "" "& btn_voltar & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""center"">") & vbcrlf
		Response.Write ("                                                  <input name=""Principal"" type=""button"" class=""buttonNormal"" id=""Principal"" onClick=""MM_goToURL('self','../principal.asp');return document.MM_returnValue"" value=""P&aacute;gina Principal"" disabled='disabled' />") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""right"">") & vbcrlf
		Response.Write ("                                                <input name=""Forward"" type=""button"" class=""buttonNormal"" id=""Forward"" onClick=""MM_goToURL('self','" & tgt_btn_avanc & "');return document.MM_returnValue"" value="""& lbl_btn_avanc & """ " & btn_avanc & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                            </tr>") & vbcrlf
		Response.Write ("                                          </table></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                      </form>") & vbcrlf
		Response.Write ("							      </table>") & vbcrlf
		Response.Write ("								    </fieldset></td>") & vbcrlf
		Response.Write ("								</tr>") & vbcrlf
		Response.Write ("							  </table></td>") & vbcrlf
		Response.Write ("							</tr></table>") & vbcrlf
		Response.Write ("							</td>") & vbcrlf
		Response.Write ("						</tr> ") & vbcrlf

	ElseIf InStr(Request.ServerVariables("HTTP_USER_AGENT"),"Mozilla") then
		if ucase(tipo_aviso) = "ERRO" then
			Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();</script>")
		elseif UCASE(tipo_aviso) = "OK" then
			if (trim(tgt_btn_avanc)<>"") then
				Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';</script>")
			else
				Response.Write("<script>alert('"&mensagem&"'); history.back();</script>")
			end if
		End if
	Else
		if ucase(tipo_aviso) = "ERRO" then
			Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();</script>")
		elseif UCASE(tipo_aviso) = "OK" then
			if (trim(tgt_btn_avanc)<>"") then
				Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';</script>")
			else
				Response.Write("<script>alert('"&mensagem&"'); history.back();</script>")
			end if
		End if
	End If

end function

Function CaixaMensagem1 (tipo_aviso, titulo, mensagem, btn_voltar, lbl_btn_avanc, btn_avanc, tgt_btn_avanc)

	If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") then

		if ucase(tipo_aviso) = "ERRO" then
			figura = "../imagens/caixa/barra-de-aviso_01.gif"
			elseif UCASE(tipo_aviso) = "AVISO" then
			figura = "../imagens/caixa/aviso_01.gif"
			elseif UCASE(tipo_aviso) = "OK" then
			figura = "../imagens/caixa/confirma_01.gif"
		End if
		if btn_voltar = true then
			btn_voltar = ""
			else
			btn_voltar = "disabled"
		end if

		if btn_avanc = true then
			btn_avanc = ""
			else
			btn_avanc = "disabled"
		end if
		Response.Write ("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">") & vbcrlf
		Response.Write ("<html>") & vbcrlf
		Response.Write ("<head>") & vbcrlf
		Response.Write ("<link rel=""stylesheet"" type=""text/css"" href=""../css/style.css"" />") & vbcrlf
		Response.Write ("<script type=""text/JavaScript"">") & vbcrlf
		Response.Write ("<!--") & vbcrlf
		Response.Write ("function MM_goToURL() { //v3.0") & vbcrlf
		Response.Write ("  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;") & vbcrlf
		Response.Write ("  for (i=0; i<(args.length-1); i+=2) eval(args[i]+"".location='""+args[i+1]+""'"");") & vbcrlf
		Response.Write ("}") & vbcrlf
		Response.Write ("//-->") & vbcrlf
		Response.Write ("</script>") & vbcrlf
		Response.Write ("</head>") & vbcrlf
		Response.Write ("<BODY>") & vbcrlf
		Response.Write ("<table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">") & vbcrlf
		Response.Write ("						<tr>") & vbcrlf
		Response.Write ("							<td align=""center"">") & vbcrlf
		Response.Write ("							<table width=""441"" cellpadding=""0"" cellspacing=""1"">") & vbcrlf
		Response.Write ("							  <tr><td width=""458""><table width=""439"" border=""0"" align=""center"" cellpadding=""0"" bordercolor=""#CCCCCC"">") & vbcrlf
		Response.Write ("								<tr>") & vbcrlf
		Response.Write ("								  <td width=""452""><fieldset>") & vbcrlf
		Response.Write ("								    <table width=""434"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                      <form target=""_self"" action=""#"" name=""frm"" >") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td width=""75""><img src="""& figura &""" width=""75"" height=""60""></td>") & vbcrlf
		Response.Write ("                                          <td width=""100%"" background=""../imagens/caixa/barra-de-aviso_02.gif"" class=""titulo_caixa_mensagem"">"& Server.HTMLEncode(UCASE(TITULO)) & "</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr class=""caixa_mensagem"">") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><div align=""justify"">"& mensagem & "</div></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                            <tr>") & vbcrlf
		Response.Write ("                                              <td><div alig & vbcrlfn=""left"">") & vbcrlf
		Response.Write ("                                                  <input name=""Back"" type=""button"" class=""buttonNormal"" id=""Back"" onClick=""javascript: history.back();"" value=""Voltar "" "& btn_voltar & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""center"">") & vbcrlf
		Response.Write ("                                                  <input name=""Principal"" type=""button"" class=""buttonNormal"" id=""Principal"" onClick=""MM_goToURL('self','../index.asp');return document.MM_returnValue"" value=""P&aacute;gina Principal"" />") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""right"">") & vbcrlf
		Response.Write ("                                                <input name=""Forward"" type=""button"" class=""buttonNormal"" id=""Forward"" onClick=""MM_goToURL('self','" & tgt_btn_avanc & "');return document.MM_returnValue"" value="""& lbl_btn_avanc & """ " & btn_avanc & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                            </tr>") & vbcrlf
		Response.Write ("                                          </table></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                      </form>") & vbcrlf
		Response.Write ("							      </table>") & vbcrlf
		Response.Write ("								    </fieldset></td>") & vbcrlf
		Response.Write ("								</tr>") & vbcrlf
		Response.Write ("							  </table></td>") & vbcrlf
		Response.Write ("							</tr></table>") & vbcrlf
		Response.Write ("							</td>") & vbcrlf
		Response.Write ("						</tr> ") & vbcrlf

	ElseIf InStr(Request.ServerVariables("HTTP_USER_AGENT"),"Mozilla") then
		if ucase(tipo_aviso) = "ERRO" then
			Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();</script>")
		elseif UCASE(tipo_aviso) = "OK" then
			if (trim(tgt_btn_avanc)<>"") then
				Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';</script>")
			else
				Response.Write("<script>alert('"&mensagem&"'); history.back();</script>")
			end if
		End if
	Else
		if ucase(tipo_aviso) = "ERRO" then
			Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();</script>")
		elseif UCASE(tipo_aviso) = "OK" then
			if (trim(tgt_btn_avanc)<>"") then
				Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';</script>")
			else
				Response.Write("<script>alert('"&mensagem&"'); history.back();</script>")
			end if
		End if
	End If

end function

Function SoundEx(WordString ,LengthOption ,CensusOption)

Dim WordStr
Dim b, b2, b3, SoundExLen, FirstLetter
Dim i


' Sanity
'
If (CensusOption > 0) Then
    LengthOption = 4
End If
If (LengthOption > 0) Then
    SoundExLen = LengthOption
End If
If (SoundExLen > 250) Then
    SoundExLen = 250
End If
If (SoundExLen < 4) Then
    SoundExLen = 4
End If
If (Len(WordString) < 1) Then
    Exit Function
End If


' Copy to WordStr
' and UpperCase
'
WordStr = UCase(WordString)


WordStr = replace(WordStr,"¡","A")
WordStr = replace(WordStr,"¿","A")
WordStr = replace(WordStr,"√","A")
WordStr = replace(WordStr,"¬","A")

WordStr = replace(WordStr,"…","E")
WordStr = replace(WordStr," ","E")

WordStr = replace(WordStr,"Õ","I")
WordStr = replace(WordStr,"Œ","I")

WordStr = replace(WordStr,"”","O")
WordStr = replace(WordStr,"‘","O")

WordStr = replace(WordStr,"⁄","U")
WordStr = replace(WordStr,"€","U")
WordStr = replace(WordStr,"«","C")


' Convert all non-alpha
' chars to spaces. (thanks John)
'
For i = 1 To Len(WordStr)
    b = Mid(WordStr, i, 1)
	if ( instr("ABCDEFGHIJKLMNOPQRSTUVWXYZ",b) = 0 ) THEN
        WordStr = Replace(WordStr, b, " ")
    End If
Next



' Remove leading and
' trailing spaces
'
WordStr = Trim(WordStr)

' sanity
'
If (Len(WordStr) < 1) Then
    Exit Function
End If


' Perform our own multi-letter
' improvements
'
' double letters will be effectively
' removed in a later step.
'
If (CensusOption < 1) Then
    b = Mid(WordStr, 1, 1)
    b2 = Mid(WordStr, 2, 1)
    If (b = "P" And b2 = "S") Then
        WordStr = Replace(WordStr, "PS", "S", 1, 1)
    End If
    If (b = "P" And b2 = "F") Then
        WordStr = Replace(WordStr, "PF", "F", 1, 1)
    End If

    WordStr = Replace(WordStr, "DG", "_G")
    WordStr = Replace(WordStr, "GH", "_H")
    WordStr = Replace(WordStr, "KN", "_N")
    WordStr = Replace(WordStr, "GN", "_N")
    WordStr = Replace(WordStr, "MB", "M_")
    WordStr = Replace(WordStr, "PH", "F_")
    WordStr = Replace(WordStr, "TCH", "_CH")
    WordStr = Replace(WordStr, "MPS", "M_S")
    WordStr = Replace(WordStr, "MPT", "M_T")
    WordStr = Replace(WordStr, "MPZ", "M_Z")

End If


' end if(Not CensusOption)
'

' Sqeeze out the extra _ letters
' from above (not strictly needed
' in VB but used in C code)
'
WordStr = Replace(WordStr, "_", "")



' This must be done AFTER our
' multi-letter replacements
' since they could change
' the first letter
'
FirstLetter = Mid(WordStr, 1, 1)

' in case first letter is
' an h, a w ...
' we'll change it to something
' that doesn't match anything
'
If (FirstLetter = "H" Or FirstLetter = "W") Then
    b = Mid(WordStr, 2)
    WordStr = "-" + b
End If


' In properly done census
' SoundEx, the H and W will
' be squezed out before
' performing the test
' for adjacent digits
' (this differs from how
' 'real' vowels are handled)
'
If (CensusOption = 1) Then
    WordStr = Replace(WordStr, "H", ".")
    WordStr = Replace(WordStr, "W", ".")
End If


' Perform classic SoundEx
' replacements
' Here, we use ';' instead of zero '0'
' because of MS strangeness with leading
' zeros in some applications.
'
WordStr = Replace(WordStr, "A", ";")
WordStr = Replace(WordStr, "E", ";")
WordStr = Replace(WordStr, "I", ";")
WordStr = Replace(WordStr, "O", ";")
WordStr = Replace(WordStr, "U", ";")
WordStr = Replace(WordStr, "Y", ";")
WordStr = Replace(WordStr, "H", ";")
WordStr = Replace(WordStr, "W", ";")

WordStr = Replace(WordStr, "B", "1")
WordStr = Replace(WordStr, "P", "1")
WordStr = Replace(WordStr, "F", "1")
WordStr = Replace(WordStr, "V", "1")

WordStr = Replace(WordStr, "C", "2")
WordStr = Replace(WordStr, "S", "2")
WordStr = Replace(WordStr, "G", "2")
WordStr = Replace(WordStr, "J", "2")
WordStr = Replace(WordStr, "K", "2")
WordStr = Replace(WordStr, "Q", "2")
WordStr = Replace(WordStr, "X", "2")
WordStr = Replace(WordStr, "Z", "2")

WordStr = Replace(WordStr, "D", "3")
WordStr = Replace(WordStr, "T", "3")

WordStr = Replace(WordStr, "L", "4")

WordStr = Replace(WordStr, "M", "5")
WordStr = Replace(WordStr, "N", "5")

WordStr = Replace(WordStr, "R", "6")
'
' End Clasic SoundEx replacements
'



' In properly done census
' SoundEx, the H and W will
' be squezed out before
' performing the test
' for adjacent digits
' (this differs from how
' 'real' vowels are handled)
'
If (CensusOption = 1) Then
    WordStr = Replace(WordStr, ".", "")
End If



' squeeze out extra equal adjacent digits
' (don't include first letter)
'
b = ""
b2 = ""
' remove from v1.0c djr: b3 = Mid(WordStr, 1, 1)
b3 = ""
For i = 1 To Len(WordStr) ' i=1 (not 2) in v1.0c
    b = Mid(WordStr, i, 1)
    b2 = Mid(WordStr, (i + 1), 1)
    If (Not (b = b2)) Then
        b3 = b3 + b
    End If
Next

WordStr = b3
If (Len(WordStr) < 1) Then
    Exit Function
End If



' squeeze out spaces and zeros (;)
' Leave the first letter code
' to be replaced below.
' (In case it made a zero)
'
WordStr = Replace(WordStr, " ", "")
b = Mid(WordStr, 1, 1)
WordStr = Replace(WordStr, ";", "")
If (b = ";") Then ' only if it got removed above
    WordStr = b + WordStr
End If



' Right pad with zero characters
'
b = String(SoundExLen, "0")
WordStr = WordStr + b


' Replace first digit with
' first letter
'
WordStr = Mid(WordStr, 2)
WordStr = FirstLetter + WordStr

' Size to taste
'
WordStr = Mid(WordStr, 1, SoundExLen)


' Copy WordStr to SoundEx
'
SoundEx = WordStr

End Function


Function Fu_consistir_CgcCpf (Vl_CgcCpf)


Fu_consistir_CgcCpf = False
on error resume next
Dim VA_CgcCpf
Dim VA_Digito
dim Numero(15)
Dim VA_Resto
Dim VA_Resultado
Dim VA_SomaDigito10
Dim VA_resto1
if isnull(Vl_CgcCpf) then
   Exit Function
end if

VA_CgcCpf = Vl_CgcCpf & space(14-len(Vl_CgcCpf))
VA_Digito = Mid(VA_CgcCpf, 13, 2)
Numero(1) = val(Mid(VA_CgcCpf, 1, 1))
Numero(2) = val(Mid(VA_CgcCpf, 2, 1))
Numero(3) = val(Mid(VA_CgcCpf, 3, 1))
Numero(4) = val(Mid(VA_CgcCpf, 4, 1))
Numero(5) = val(Mid(VA_CgcCpf, 5, 1))
Numero(6) = val(Mid(VA_CgcCpf, 6, 1))
Numero(7) = val(Mid(VA_CgcCpf, 7, 1))
Numero(8) = val(Mid(VA_CgcCpf, 8, 1))
Numero(9) = val(Mid(VA_CgcCpf, 9, 1))
Numero(10) = val(Mid(VA_CgcCpf, 10, 1))
Numero(11) = val(Mid(VA_CgcCpf, 11, 1))
Numero(12) = val(Mid(VA_CgcCpf, 12, 1))
Numero(13) = val(Mid(VA_CgcCpf, 13, 1))
Numero(14) = val(Mid(VA_CgcCpf, 14, 1))

if err <> "" then
   exit function
end if

If Len(Trim(Vl_CgcCpf)) > 11 Then ' Cgc
   VA_Resultado = Numero(1) * 2
   If VA_Resultado > 9 Then
      VA_SomaDigito10 = VA_Resultado + 1
   Else
      VA_SomaDigito10 = VA_Resultado
   End If
   VA_Resultado = Numero(3) * 2
   If VA_Resultado > 9 Then
      VA_SomaDigito10 = VA_SomaDigito10 + VA_Resultado + 1
   Else
      VA_SomaDigito10 = VA_SomaDigito10 + VA_Resultado
   End If
   VA_Resultado = Numero(5) * 2
   If VA_Resultado > 9 Then
     VA_SomaDigito10 = VA_SomaDigito10 + VA_Resultado + 1
   Else
     VA_SomaDigito10 = VA_SomaDigito10 + VA_Resultado
   End If
   VA_Resultado = Numero(7) * 2
   If VA_Resultado > 9 Then
      VA_SomaDigito10 = VA_SomaDigito10 + VA_Resultado + 1
   Else
      VA_SomaDigito10 = VA_SomaDigito10 + VA_Resultado
   End If
   VA_SomaDigito10 = VA_SomaDigito10 + Numero(2) + Numero(4) + Numero(6)
   If Mid(Str(VA_SomaDigito10), Len(Str(VA_SomaDigito10)), 1) = "0" Then
      VA_Resto = 0
   Else
   VA_Resto = 10 - Val(Mid(Str(VA_SomaDigito10), Len(Str(VA_SomaDigito10)), 1))
   End If
   If VA_Resto <> Numero(8) Then
      Exit Function
   End If
   VA_Resultado = (Numero(1) * 5) + (Numero(2) * 4) _
   + (Numero(3) * 3) + (Numero(4) * 2) _
   + (Numero(5) * 9) + (Numero(6) * 8) + _
   (Numero(7) * 7) + (Numero(8) * 6) + _
   (Numero(9) * 5) + (Numero(10) * 4) + _
   (Numero(11) * 3) + (Numero(12) * 2)
   ' Atribui para resto o resto da divis„o
   ' de VA_resultado dividido por 11
   VA_Resto = VA_Resultado Mod 11
   If VA_Resto < 2 Then
      VA_resto1 = 0
   Else
      VA_resto1 = 11 - VA_Resto
   End If
   If VA_resto1 <> Numero(13) Then
      Exit Function
   End If
   VA_Resultado = (Numero(1) * 6) + _
   (Numero(2) * 5) + (Numero(3) * 4) + _
   (Numero(4) * 3) + (Numero(5) * 2) + _
   (Numero(6) * 9) + (Numero(7) * 8) + _
   (Numero(8) * 7) + (Numero(9) * 6) + _
   (Numero(10) * 5) + (Numero(11) * 4) + _
   (Numero(12) * 3) + (Numero(13) * 2)
   ' Atribui para resto o resto da divis„o
   ' de VA_resultado dividido por 11
   VA_Resto = VA_Resultado Mod 11
   If VA_Resto < 2 Then
      VA_resto1 = 0
   Else
      VA_resto1 = 11 - VA_Resto
   End If
   If VA_resto1 <> Numero(14) Then
      Exit Function
   End If
Else ' Cpf
   VA_Resultado = (Numero(4) * 1) + _
   (Numero(5) * 2) + (Numero(6) * 3) _
   + (Numero(7) * 4) + (Numero(8) * 5) _
   + (Numero(9) * 6) + (Numero(10) * 7)_
   + (Numero(11) * 8) + (Numero(12) * 9)
   VA_Resto = VA_Resultado Mod 11
   If VA_Resto > 9 Then
      VA_resto1 = VA_Resto - 10

   Else
      VA_resto1 = VA_Resto
   End If
   If VA_resto1 <> Numero(13) Then
      Exit Function
   End If

   VA_Resultado = (Numero(5) * 1) _
   + (Numero(6) * 2) + (Numero(7) * 3) _
   + (Numero(8) * 4) + (Numero(9) * 5) + _
   (Numero(10) * 6) + (Numero(11) * 7) + _
   (Numero(12) * 8) + (VA_Resto * 9)
   VA_Resto = VA_Resultado Mod 11
   If VA_Resto > 9 Then
      VA_resto1 = VA_Resto - 10
   Else
      VA_resto1 = VA_Resto
   End If
   If VA_resto1 <> Numero(14) Then
      Exit Function
   End If
End If
Fu_consistir_CgcCpf = True
End Function

 Function StripTags(ByRef Str)

    Str = Str & ""
    If Str = "" Then
 StripTags = ""
 Exit Function
    End If

    Dim nw, i, j, c, k, x

    i = 1
    x = 1
    k = Len(Str)
    nw = ""

    Do While i > 0

 i = InStr(i, Str, "<")
 If i = 0 Then Exit Do

 j = InStr(i, Str, ">")
 If j < i Then Exit Do

 c = Mid(Str, i + 1, 1)
 If ereg("^[a-zA-Z/!]$", c) Then
     nw = nw & Mid(Str, x, i - x)
     i = j + 1
     x = i
 Else
     i = i + 1
 End If

 If i >= k Then Exit Do

    Loop

    nw = nw & Mid(Str, x, k - x + 1)
    StripTags = nw

End Function

Function ereg(ByVal Expr, ByVal Varn)
    Dim Regex
    Set Regex = New RegExp
    Regex.Pattern = Expr
    Regex.IgnoreCase = False
    Regex.Global = True
    ereg = Regex.Test(Varn)
    Set Regex = Nothing
End Function

'Conta o n˙mero de linhas em um arquivo.
Function FileCountLines(ByVal FileName)

    Dim sStream
    Dim iLines
    Dim i, j, k

    If Mid(FileName, 2, 1) <> ":" Then
 FileName = Server.MapPath(FileName)
    End If
    sStream = ReadFile(FileName)

    i = 1
    j = 1
    iLines = 0
    k = Len(sStream)

    Do While i <= k
 i = InStr(i, sStream, vbCrLf)
 If i = 0 Then Exit Do
 iLines = iLines + 1
 i = i + 2
    Loop

    If iLines > 0 Then iLines = iLines + 1
    FileCountLines = iLines

End Function

 Function BinaryToString(strBinary)
Dim intCount
BinaryToString =""
For intCount = 1 to LenB(strBinary)
BinaryToString = BinaryToString & chr(AscB(MidB(strBinary,intCount,1)))
Next
End Function

Function CaixaMensagem_sicat (tipo_aviso, titulo, mensagem, btn_voltar, lbl_btn_avanc, btn_avanc, tgt_btn_avanc)
	If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") then
		if ucase(tipo_aviso) = "ERRO" then
			figura = "../registrafacil/imagens/caixa/barra-de-aviso_01.gif"
			elseif UCASE(tipo_aviso) = "AVISO" then
			figura = "../registrafacil/imagens/caixa/aviso_01.gif"
			elseif UCASE(tipo_aviso) = "OK" then
			figura = "../registrafacil/imagens/caixa/confirma_01.gif"
		End if
		if btn_voltar = true then
			btn_voltar = ""
			else
			btn_voltar = "disabled"
		end if

		if btn_avanc = true then
			btn_avanc = ""
			else
			btn_avanc = "disabled"
		end if
		Response.Write ("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">") & vbcrlf
		Response.Write ("<html>") & vbcrlf
		Response.Write ("<head>") & vbcrlf
		Response.Write ("<link rel=""stylesheet"" type=""text/css"" href=""../registrafacil/css/style.css"" />") & vbcrlf
		Response.Write ("<script type=""text/JavaScript"">") & vbcrlf
		Response.Write ("<!--") & vbcrlf
		Response.Write ("function MM_goToURL() { //v3.0") & vbcrlf
		Response.Write ("  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;") & vbcrlf
		Response.Write ("  for (i=0; i<(args.length-1); i+=2) eval(args[i]+"".location='""+args[i+1]+""'"");") & vbcrlf
		Response.Write ("}") & vbcrlf
		Response.Write ("//-->") & vbcrlf
		Response.Write ("</script>") & vbcrlf
		Response.Write ("</head>") & vbcrlf
		Response.Write ("<BODY>") & vbcrlf
		Response.Write ("<table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">") & vbcrlf
		Response.Write ("						<tr>") & vbcrlf
		Response.Write ("							<td align=""center"">") & vbcrlf
		Response.Write ("							<table width=""441"" cellpadding=""0"" cellspacing=""1"">") & vbcrlf
		Response.Write ("							  <tr><td width=""458""><table width=""439"" border=""0"" align=""center"" cellpadding=""0"" bordercolor=""#CCCCCC"">") & vbcrlf
		Response.Write ("								<tr>") & vbcrlf
		Response.Write ("								  <td width=""452""><fieldset>") & vbcrlf
		Response.Write ("								    <table width=""434"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                      <form target=""_self"" action=""#"" name=""frm"" >") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td width=""75""><img src="""& figura &""" width=""75"" height=""60""></td>") & vbcrlf
		Response.Write ("                                          <td width=""100%"" background=""../registrafacil/imagens/caixa/barra-de-aviso_02.gif"" class=""titulo_caixa_mensagem"">"& Server.HTMLEncode(UCASE(TITULO)) & "</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr class=""caixa_mensagem"">") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><div align=""justify"">"& mensagem & "</div></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                            <tr>") & vbcrlf
		Response.Write ("                                              <td><div alig & vbcrlfn=""left"">") & vbcrlf
		Response.Write ("                                                  <input name=""Back"" type=""button"" class=""buttonNormal"" id=""Back"" onClick=""javascript: history.back();"" value=""Voltar "" "& btn_voltar & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""center"">") & vbcrlf
		Response.Write ("                                                  <input name=""Principal"" type=""button"" class=""buttonNormal"" id=""Principal"" onClick=""MM_goToURL('top','../registrafacil/indexprincipal.asp');return document.MM_returnValue"" value=""P&aacute;gina Principal"" />") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""right"">") & vbcrlf
		Response.Write ("                                                <input name=""Forward"" type=""button"" class=""buttonNormal"" id=""Forward"" onClick=""MM_goToURL('self','" & tgt_btn_avanc & "');return document.MM_returnValue"" value="""& lbl_btn_avanc & """ " & btn_avanc & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                            </tr>") & vbcrlf
		Response.Write ("                                          </table></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                      </form>") & vbcrlf
		Response.Write ("							      </table>") & vbcrlf
		Response.Write ("								    </fieldset></td>") & vbcrlf
		Response.Write ("								</tr>") & vbcrlf
		Response.Write ("							  </table></td>") & vbcrlf
		Response.Write ("							</tr></table>") & vbcrlf
		Response.Write ("							</td>") & vbcrlf
		Response.Write ("						</tr> ") & vbcrlf

	ElseIf InStr(Request.ServerVariables("HTTP_USER_AGENT"),"Mozilla") then
		if ucase(tipo_aviso) = "ERRO" then
			Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();</script>")
		elseif UCASE(tipo_aviso) = "OK" then
			if (trim(tgt_btn_avanc)<>"") then
				Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';</script>")
			else
				Response.Write("<script>alert('"&mensagem&"'); history.back();</script>")
			end if
		End if
	Else
		if ucase(tipo_aviso) = "ERRO" then
			Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();</script>")
		elseif UCASE(tipo_aviso) = "OK" then
			if (trim(tgt_btn_avanc)<>"") then
				Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';</script>")
			else
				Response.Write("<script>alert('"&mensagem&"'); history.back();</script>")
			end if
		End if
	End If

end function

Function CaixaMensagemTorpedo (tipo_aviso, titulo, mensagem, btn_voltar, lbl_btn_avanc, btn_avanc, tgt_btn_avanc)

	If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") then

		if ucase(tipo_aviso) = "ERRO" then
			figura = "../registrafacil/imagens/caixa/barra-de-aviso_01.gif"
			elseif UCASE(tipo_aviso) = "AVISO" then
			figura = "../registrafacil/imagens/caixa/aviso_01.gif"
			elseif UCASE(tipo_aviso) = "OK" then
			figura = "../registrafacil/imagens/caixa/confirma_01.gif"
		End if
		if btn_voltar = true then
			btn_voltar = ""
			else
			btn_voltar = "disabled"
		end if

		if btn_avanc = true then
			btn_avanc = ""
			else
			btn_avanc = "disabled"
		end if
		Response.Write ("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">") & vbcrlf
		Response.Write ("<html>") & vbcrlf
		Response.Write ("<head>") & vbcrlf
		Response.Write ("<link rel=""stylesheet"" type=""text/css"" href=""../registrafacil/css/style.css"" />") & vbcrlf
		Response.Write ("<script type=""text/JavaScript"">") & vbcrlf
		Response.Write ("<!--") & vbcrlf
		Response.Write ("function MM_goToURL() { //v3.0") & vbcrlf
		Response.Write ("  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;") & vbcrlf
		Response.Write ("  for (i=0; i<(args.length-1); i+=2) eval(args[i]+"".location='""+args[i+1]+""'"");") & vbcrlf
		Response.Write ("}") & vbcrlf
		Response.Write ("//-->") & vbcrlf
		Response.Write ("</script>") & vbcrlf
		Response.Write ("</head>") & vbcrlf
		Response.Write ("<BODY>") & vbcrlf
		Response.Write ("<table width=""100%"" height=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">") & vbcrlf
		Response.Write ("						<tr>") & vbcrlf
		Response.Write ("							<td align=""center"">") & vbcrlf
		Response.Write ("							<table width=""441"" cellpadding=""0"" cellspacing=""1"">") & vbcrlf
		Response.Write ("							  <tr><td width=""458""><table width=""439"" border=""0"" align=""center"" cellpadding=""0"" bordercolor=""#CCCCCC"">") & vbcrlf
		Response.Write ("								<tr>") & vbcrlf
		Response.Write ("								  <td width=""452""><fieldset>") & vbcrlf
		Response.Write ("								    <table width=""434"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                      <form target=""_self"" action=""#"" name=""frm"" >") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td width=""75""><img src="""& figura &""" width=""75"" height=""60""></td>") & vbcrlf
		Response.Write ("                                          <td width=""100%"" background=""../registrafacil/imagens/caixa/barra-de-aviso_02.gif"" class=""titulo_caixa_mensagem"">"& Server.HTMLEncode(UCASE(TITULO)) & "</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr class=""caixa_mensagem"">") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><div align=""justify"">"& mensagem & "</div></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF"">&nbsp;</td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                        <tr>") & vbcrlf
		Response.Write ("                                          <td colspan=""2"" bgcolor=""#FFFFFF""><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">") & vbcrlf
		Response.Write ("                                            <tr>") & vbcrlf
		Response.Write ("                                              <td><div alig & vbcrlfn=""left"">") & vbcrlf
		Response.Write ("                                                  <input name=""Back"" type=""button"" class=""buttonNormal"" id=""Back"" onClick=""javascript: history.back();"" value=""Voltar "" "& btn_voltar & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""center"">") & vbcrlf
		Response.Write ("                                                  <input name=""Principal"" type=""button"" class=""buttonNormal"" id=""Principal"" onClick=""MM_goToURL('self','../registrafacil/principal.asp');return document.MM_returnValue"" value=""P&aacute;gina Principal"" disabled='disabled' />") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                              <td><div align=""right"">") & vbcrlf
		Response.Write ("                                                <input name=""Forward"" type=""button"" class=""buttonNormal"" id=""Forward"" onClick=""MM_goToURL('self','" & tgt_btn_avanc & "');return document.MM_returnValue"" value="""& lbl_btn_avanc & """ " & btn_avanc & ">") & vbcrlf
		Response.Write ("                                              </div></td>") & vbcrlf
		Response.Write ("                                            </tr>") & vbcrlf
		Response.Write ("                                          </table></td>") & vbcrlf
		Response.Write ("                                        </tr>") & vbcrlf
		Response.Write ("                                      </form>") & vbcrlf
		Response.Write ("							      </table>") & vbcrlf
		Response.Write ("								    </fieldset></td>") & vbcrlf
		Response.Write ("								</tr>") & vbcrlf
		Response.Write ("							  </table></td>") & vbcrlf
		Response.Write ("							</tr></table>") & vbcrlf
		Response.Write ("							</td>") & vbcrlf
		Response.Write ("						</tr> ") & vbcrlf

	ElseIf InStr(Request.ServerVariables("HTTP_USER_AGENT"),"Mozilla") then
		if ucase(tipo_aviso) = "ERRO" then
			Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();</script>")
		elseif UCASE(tipo_aviso) = "OK" then
			if (trim(tgt_btn_avanc)<>"") then
				Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';</script>")
			else
				Response.Write("<script>alert('"&mensagem&"'); history.back();</script>")
			end if
		End if
	Else
		if ucase(tipo_aviso) = "ERRO" then
			Response.Write("<script>alert('Nenhum registro encontrado!'); history.back();</script>")
		elseif UCASE(tipo_aviso) = "OK" then
			if (trim(tgt_btn_avanc)<>"") then
				Response.Write("<script>alert('"&mensagem&"'); self.location.href='"&tgt_btn_avanc&"';</script>")
			else
				Response.Write("<script>alert('"&mensagem&"'); history.back();</script>")
			end if
		End if
	End If

end function

function exibir_consultor(wCodigo,wAlign)




codigo = wCodigo

Select Case codigo
	case 21289786372
		'Andre
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/andre_color.jpg' width='86' height='86'>&nbsp;"
	case 10437686884
		'Alessiana
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/alessiana_foto.jpg' width='86' height='86'>&nbsp;"
	case 52521893315
		'Carlene
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/carlene costa_foto.jpg' width='86' height='86'>&nbsp;"
	case 32283962315
		'Claudete
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/claudete_foto.jpg' width='86' height='86'>&nbsp;"
	case 65935543320
		'Christiano
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/christiano_foto.jpg' width='86' height='86'>&nbsp;"

	case 65845080304
		'Luis Wagner
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/luiswagner_foto.jpg' width='86' height='86'>&nbsp;"
	case 76589242372
		'Monica Braga
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/MÙnica_foto.jpg' width='86' height='86'>&nbsp;"

	case 62865900304
		'Nadja Karla
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/Nadja_foto.jpg' width='86' height='86'>&nbsp;"
	case 20345127315
		'Wagner Alencar
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/wagner_foto.jpg' width='86' height='86'>&nbsp;"
	case 66176123372
		'Monique
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/monique_foto.jpg' width='86' height='86'>&nbsp;"
	case 70395896304
		'Paulo AndrÈ
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/paulo AndrÈ_foto.jpg' width='86' height='86'>&nbsp;"
	case 64161005334
		'Marta
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/marta_foto.jpg' width='86' height='86'>&nbsp;"
	case 01401435300
		'Luis AndrÈ
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/lui_andre_foto.jpg' width='86' height='86'>&nbsp;"

End Select




Select Case codigo

	case 10437686884
		'Alessiana
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/alessiana_as.jpg' width='200' height='60'></div>"
	case 52521893315
		'Carlene
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/carlene_as.jpg' width='160' height='60'></div>"
	case 32283962315
		'Claudete
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/claudete_as.jpg' width='60' height='60'></div>"
	case 65935543320
		'Christiano
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/christiano_as.jpg' width='200' height='60'></div>"
	case 65845080304
		'Luis Wagner
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/Luis Wagner.jpg' width='110' height='60'></div>"
	case 76589242372
		'Monica Braga
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/monica_as.jpg' width='196' height='60'></div>"
	case 62865900304
		'Nadja Karla
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/nadja.jpg' width='115' height='60'></div>"
	case 20345127315
		'Wagner Alencar
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/wagner_as.jpg' width='41' height='100'></div>"
	case 66176123372
		'Monique
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/monique_as.jpg' width='165' height='60'></div>"
	case 70395896304
		'Paulo AndrÈ
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/paulo_as.jpg' width='135' height='60'></div>"
	case 21289786372
		'AndrÈ
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/andre_as.jpg' width='140' height='54'><br>"
	case 64161005334
		'Marta
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/marta_as.jpg' width='201' height='45'></div>"
	case 01401435300
		'Luis AndrÈ
		response.write "<div align='"& wAlign &"'><img src='/registrafacil/imagens/wettor/lui_andre_as.jpg' width='220' height='60'></div>"
End Select





assinatura = wcodigo

Select Case assinatura
	case 21289786372
		'Andre
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>SEU CONSULTOR NA WETTOR:<BR></STRONG><u>ANDR… LUÕS</u> C.SILVA<br>Telefones: (85)3066.3999 / 8711.3830<br>andreluis@wettor.com.br</div>"
	case 10437686884
		'Alessiana
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>Sua CONSULTORa NA WETTOR:<BR></STRONG><u>ALESSIANA SOUZA</u> do carmo pedro<br>Telefones: (85)3066.3999 / 8711.3827<br>ALESSIANASOUZA@wettor.com.br</div>"
	case 52521893315
		'karlene
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>Sua CONSULTORa NA WETTOR:<BR></STRONG><u>KARLENE COSTA</u><br>Telefones: (85)3066.3999 / 8607.6942<br>kARLENECOSTA@wettor.com.br</div>"
	case 32283962315
		'Claudete
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>Sua CONSULTORa NA WETTOR:<BR></STRONG><u>CLAUDETE</u> LOPES<br>Telefones: (85)3066.3999 / 87119914 / 8850.3899<br>CLAUDETELOPES@wettor.com.br</div>"
	case 65935543320
		'Christiano
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>SEU CONSULTOR NA WETTOR:<BR></STRONG><u>christiano</u> g. de <u>Matos</u> rocha<br>Telefones: 3066.3999 / 8760.9771<br>christianomatos@wettor.com.br</div>"

	case 65845080304
		'Luis Wagner
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>SEU CONSULTOR NA WETTOR:<BR></STRONG><u>Luis Wagner</u> Domingos<br>Telefones: (85)3066.3999 / 8711.3829<br>luiswagner@wettor.com.br</div>"
	case 76589242372
		'Monica Braga
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>Sua CONSULTORa NA WETTOR:<BR></STRONG><u>Monica braga</u><br>Telefones: (85)3066.3999 / 8828.9368<br>monicabraga@wettor.com.br</div>"

	case 62865900304
		'Nadja Karla
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>Sua CONSULTORa NA WETTOR:<BR></STRONG><u>NADJA KARLA</u> ANDRADE SALES<br>Telefones: (85)3066.3999 / 9617.9084<br>nadjakarla@wettor.com.br</div>"
	case 20345127315
		'Wagner Alencar
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>SEU CONSULTOR NA WETTOR:<BR></STRONG><u>Wagner Alencar</u> Domingos<br>Telefones: (85)3066.3999 / 8711.3832<br>wagneralencar@wettor.com.br</div>"
	case 66176123372
		'Monique
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>Sua CONSULTORa NA WETTOR:<BR></STRONG><u>Monique Barros</u> Sampaio<br>Telefones: (85)8849.4774 / 8636.6184<br>moniquebarros@wettor.com.br</div>"
	case 70395896304
		'Paulo AndrÈ
		response.write "<br><div align='"& wAlign &"' class='txt_normal'><strong>SEU CONSULTOR NA WETTOR:<BR></STRONG><u>Paulo AndrÈ</u> Oliveira da Costa<br>Telefones: (85)3066.3999 / 9108.8210<br>pauloandre@wettor.com.br</div>"
	case 64161005334
		'Marta
		response.write "<br><div align='"& wAlign &"' class='txt_normal' align='left'><strong>Sua CONSULTORa NA WETTOR:<BR></STRONG><u>Marta Alencar</u> Domingos<br>Telefones: (85)3066.3999 / 8840.9259<br>martaalencar@wettor.com.br</div>"
	case 01401435300
		'Luis AndrÈ
		 response.write "<br><div align='"& wAlign &"' class='txt_normal' align='left'><strong>seu CONSULTOR NA WETTOR:<BR></STRONG><u>LuÌs AndrÈ</u> santos domingos<br>Telefones: (85)3066.3999 / 8887.0925<br>patentes@wettor.com.br</div>"

End Select
response.write "<br>"

end function

function CalculaCPF(WCPF)

Dim RecebeCPF, Numero(11), soma, resultado1, resultado2
RecebeCPF = WCPF
'Retirar todos os caracteres que nao sejam 0-9
s=""
for x=1 to len(RecebeCPF)
	ch=mid(RecebeCPF,x,1)
	if asc(ch)>=48 and asc(ch)<=57 then
		s=s & ch
	end if
next
RecebeCPF = s



if len(RecebeCPF) <> 11 then
	CalculaCPF = false
elseif RecebeCPF = "00000000000" then
	CalculaCPF = false
else


Numero(1) = Cint(Mid(RecebeCPF,1,1))
Numero(2) = Cint(Mid(RecebeCPF,2,1))
Numero(3) = Cint(Mid(RecebeCPF,3,1))
Numero(4) = Cint(Mid(RecebeCPF,4,1))
Numero(5) = Cint(Mid(RecebeCPF,5,1))
Numero(6) = CInt(Mid(RecebeCPF,6,1))
Numero(7) = Cint(Mid(RecebeCPF,7,1))
Numero(8) = Cint(Mid(RecebeCPF,8,1))
Numero(9) = Cint(Mid(RecebeCPF,9,1))
Numero(10) = Cint(Mid(RecebeCPF,10,1))
Numero(11) = Cint(Mid(RecebeCPF,11,1))

soma = 10 * Numero(1) + 9 * Numero(2) + 8 * Numero(3) + 7 * Numero(4) + 6 * Numero(5) + 5 * Numero(6) + 4 * Numero(7) + 3 * Numero(8) + 2 * Numero(9)
soma = soma -(11 * (int(soma / 11)))
if soma = 0 or soma = 1 then
	resultado1 = 0
else
	resultado1 = 11 - soma
end if

if resultado1 = Numero(10) then
	soma = Numero(1) * 11 + Numero(2) * 10 + Numero(3) * 9 + Numero(4) * 8 + Numero(5) * 7 + Numero(6) * 6 + Numero(7) * 5 + Numero(8) * 4 + Numero(9) * 3 + Numero(10) * 2
	soma = soma -(11 * (int(soma / 11)))
	if soma = 0 or soma = 1 then
		resultado2 = 0
	else
		resultado2 = 11 - soma
	end if

	if resultado2 = Numero(11) then
		CalculaCPF = True
	else
		CalculaCPF = false
	end if
else
	CalculaCPF = false
end if
end if
end function

function CalculaCNPJ(wCNPJ)

Dim RecebeCNPJ, Numero(14), soma, resultado1, resultado2
RecebeCNPJ = wCNPJ

s=""
for x=1 to len(RecebeCNPJ)
	ch=mid(RecebeCNPJ,x,1)
	if asc(ch)>=48 and asc(ch)<=57 then
		s=s & ch
	end if
next

RecebeCNPJ = s
if len(RecebeCNPJ) <> 14 then
	CalculaCNPJ = false
elseif RecebeCNPJ = "00000000000000" then
	CalculaCNPJ = false
else

Numero(1) = Cint(Mid(RecebeCNPJ,1,1))
Numero(2) = Cint(Mid(RecebeCNPJ,2,1))
Numero(3) = Cint(Mid(RecebeCNPJ,3,1))
Numero(4) = Cint(Mid(RecebeCNPJ,4,1))
Numero(5) = Cint(Mid(RecebeCNPJ,5,1))
Numero(6) = CInt(Mid(RecebeCNPJ,6,1))
Numero(7) = Cint(Mid(RecebeCNPJ,7,1))
Numero(8) = Cint(Mid(RecebeCNPJ,8,1))
Numero(9) = Cint(Mid(RecebeCNPJ,9,1))
Numero(10) = Cint(Mid(RecebeCNPJ,10,1))
Numero(11) = Cint(Mid(RecebeCNPJ,11,1))
Numero(12) = Cint(Mid(RecebeCNPJ,12,1))
Numero(13) = Cint(Mid(RecebeCNPJ,13,1))
Numero(14) = Cint(Mid(RecebeCNPJ,14,1))

soma = Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 + Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2
soma = soma -(11 * (int(soma / 11)))
if soma = 0 or soma = 1 then
	resultado1 = 0
else
	resultado1 = 11 - soma
end if

if resultado1 = Numero(13) then
soma = Numero(1) * 6 + Numero(2) * 5 + Numero(3) * 4 + Numero(4) * 3 + Numero(5) * 2 + Numero(6) * 9 + Numero(7) * 8 + Numero(8) * 7 + Numero(9) * 6 + Numero(10) * 5 + Numero(11) * 4 + Numero(12) * 3 + Numero(13) * 2
soma = soma - (11 * (int(soma/11)))
if soma = 0 or soma = 1 then
	resultado2 = 0
else
	resultado2 = 11 - soma
end if

if resultado2 = Numero(14) then
	CalculaCNPJ = true
else
	CalculaCNPJ = false
end if
else
	CalculaCNPJ = false
end if
end if

end function

Function QuebraTexto(Texto, Caracteres)

    If Len(Texto) => Eval(Caracteres) Then

        QuebraTexto = Left(Texto, Caracteres) &"<BR>"& QuebraTexto(Right(Texto,Len(Texto)-Caracteres), Caracteres)

    Else

        QuebraTexto = Texto

    End If

End Function

'Dados da 2™ Testemunha dos contratos da Wettor
session("testemunha_wet") = "RAQUEL SANTANA GOMES"
Session("testemunha_cpf_wet") = "625.229.733-04"
'Variavel com o ip do registrafacil'
'Variavel com o ip do registrafacil'
'const g_ipWettor = "200.253.26.82"
const g_ipWettor = "187.109.34.242"
const g_PortaSMTP = "587"
const solicitantePadrao = "41572819000100"

'g_LinkLDSOFT = "http://ld2.ldsoft.com.br/siteld/webseek/marca_editar.asp"

'g_LinkLDSOFT = "http://ld2.ldsoft.com.br/webseek7/marca_editar.asp"

g_LinkLDSOFT = "http://ld2.ldsoft.com.br/webseek7/marca_editar_todos.asp"

'USU¡RIOS PERMITIDOS PARA EXCLUIR MARCAS
const usuarioExclusaoProcesso_1 = "ELINDA"
const usuarioExclusaoProcesso_2 = "RAQUEL"
const usuarioExclusaoProcesso_3 = "FABIANA"

Function IIf(bClause, sTrue, sFalse)
    If CBool(bClause) Then
        IIf = sTrue
    Else 
        IIf = sFalse
    End If
End Function

%>