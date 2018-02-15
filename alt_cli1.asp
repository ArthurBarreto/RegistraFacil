<!--#include virtual="/registrafacil/include/db_open.asp"-->
<%
Session.LCID = 1046
if trim(request.form("rg")) = "" then
   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO DO REG/INSCRIÇÃO!","Caso não exista Incrição Estadual para a Empresa/Órgão digite ISENTO.",true,"Avançar",false,"")
   response.end
end if   
'Hp
'if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> trim(Tira_Pontuacao(request.Form("cnpj_cpf"))) then
  'rs1.open "select * from gedclientes where cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "'",conn,3,3
   'if not rs1.eof then
       'Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","CNPJ/CPF já cadastrado no banco de dados.",true,"Avançar",false,"")
       'response.end 
  ' end if
   'rs1.close
'end if'

rs1.open "select * from gednegra where cod_grupo = '" & session("franquia") & "'",conn,3,3
while not rs1.eof 
      if instr(UCASE(request.form("email1")),UCASE(rs1("lista_negra"))) > 0 then
		   Call CaixaMensagem ("erro","E-MAIL INVÁLIDO (LISTA NEGRA)!","Verifique o E-Mail pois ele encontra-se na Lista Negra e alimente um E-Mail válido.  E-Mail = " & request.form("email1"),true,"Avançar",false,"")
    	   response.end 
	  end if
	rs1.movenext
wend	  	   
rs1.close

if request.form("status_marca") = "1" then
   wstatus_marca = 1
else
   wstatus_marca = 0
end if

if  NOT isnumeric(request.Form("cod_sub")) AND not isnumeric(request.Form("cod_sub_patente")) AND not isnumeric(request.Form("cod_sub_avaliacao")) AND not isnumeric(request.Form("cod_sub_internacional")) AND not isnumeric(request.Form("cod_sub_software")) AND not isnumeric(request.Form("cod_sub_direito")) then
	   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","Consultor inválido.",true,"Avançar",false,"")
       response.end 
end if


if request.form("status_patente") = "1" then
	SQL = "update GEDCLIENTES set status_patente = 1 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
else
	SQL = "update GEDCLIENTES set status_patente= 0 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
end if 

if request.form("status_avaliacao") = "1" then
	SQL = "update GEDCLIENTES set status_avaliacao= 1 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
else
	SQL = "update GEDCLIENTES set status_avaliacao= 0 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
end if   

if request.form("status_internacional") = "1" then
	SQL = "update GEDCLIENTES set status_internacional = 1 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
else
	SQL = "update GEDCLIENTES set status_internacional = 0 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
end if   

if request.form("status_software") = "1" then
	SQL = "update GEDCLIENTES set status_software = 1 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
else
	SQL = "update GEDCLIENTES set status_software = 0 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
end if   

if request.form("status_direito") = "1" then
	SQL = "update GEDCLIENTES set status_direito = 1 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
else
	SQL = "update GEDCLIENTES set status_direito = 0 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
end if   


if request.form("status_juridico") = "1" then
	SQL = "update GEDCLIENTES set status_juridico = 1 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
else
	SQL = "update GEDCLIENTES set status_juridico = 0 where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	conn.execute (sql)
end if   


if not isnumeric(request.Form("cod_sub_patente")) then
   wcod_sub_patente = 0 
else
   wcod_sub_patente = request.Form("cod_sub_patente")
end if      

if not isnumeric(request.Form("cod_sub_avaliacao")) then
   wcod_sub_avaliacao = 0 
else
   wcod_sub_avaliacao = request.Form("cod_sub_avaliacao")
end if      

if not isnumeric(request.Form("cod_sub_internacional")) then
   wcod_sub_internacional = 0 
else
   wcod_sub_internacional = request.Form("cod_sub_internacional")
end if      

if not isnumeric(request.Form("cod_sub_software")) then
   wcod_sub_software = 0 
else
   wcod_sub_software = request.Form("cod_sub_software")
end if      

if not isnumeric(request.Form("cod_sub_direito")) then
   wcod_sub_direito = 0 
else
   wcod_sub_direito = request.Form("cod_sub_direito")
end if      

if not isnumeric(request.Form("cod_sub_juridico")) then
   wcod_sub_juridico = 0 
else
   wcod_sub_juridico = request.Form("cod_sub_juridico")
end if      

if not isnumeric(request.Form("cod_sub_cob")) then
   wcod_sub_cob = 0 
else
   wcod_sub_cob = request.Form("cod_sub_cob")
end if

if Request.Form("representante1") <> "" and Tira_Pontuacao(Request.Form("cpf_representante1")) = "" then
	   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","Todos os campos do representante legal são obrigatório.",true,"Avançar",false,"")
       response.end 
end if

if Request.Form("representante1") <> "" and trim(Request.Form("aniver_representante1")) = "" then
	   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","Todos os campos do representante legal são obrigatório.",true,"Avançar",false,"")
   response.end 
end if

if Request.Form("representante1") <> "" and trim(Request.Form("cargo_representante1")) = "" then
	   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","Todos os campos do representante legal são obrigatório.",true,"Avançar",false,"")
   response.end 
end if

if Request.Form("representante1") <> "" and trim(Request.Form("rg_representante1")) = "" then
	   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","Todos os campos do representante legal são obrigatório.",true,"Avançar",false,"")
   response.end 
end if

if Request.Form("representante1") <> "" and trim(Request.Form("prof_representante1")) = "" then
	   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","Todos os campos do representante legal são obrigatório.",true,"Avançar",false,"")
   response.end 
end if


if Request.Form("representante1") <> "" and trim(Request.Form("email_representante1")) = "" then
	   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","Todos os campos do representante legal são obrigatório.",true,"Avançar",false,"")
   response.end 
end if

if trim(request.form("cod_grupo")) = "" then
   Call CaixaMensagem ("erro","ERRO NO CADASTRAMENTO!","Franquia Inválida, tente novamente.",true,"Avançar",false,"")
   response.end 
end if
   
conn.begintrans

if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> "" then
	conn.execute "update gedusu set cod_grupo='"& Request.Form("cod_grupo") & "', desusuario='"& Tira_Pontuacao(Request.Form("login")) & "', senha='"& Tira_Pontuacao(Request.Form("senha")) & "', status="& replace(Request.Form("status"),"'","") & ", cnpj_cpf = '" & trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) & "' WHERE cnpj_cpf ='" & request.Form("cnpj_cpf") & "'"
else
	sql = "select * from GEDUSU where CNPJ_CPF = '"&request.Form("cnpj_cpf")&"'" 
	rs1.open sql, conn, 3, 3
	if not rs1.eof then
		conn.execute "update gedusu set cod_grupo='"& Request.Form("cod_grupo") & "', desusuario='"& Tira_Pontuacao(Request.Form("login")) & "', senha='"& Tira_Pontuacao(Request.Form("senha")) & "', status="& replace(Request.Form("status"),"'","") & " WHERE cnpj_cpf ='" & request.Form("cnpj_cpf") & "'"
	else
		conn.execute "insert gedusu (cod_grupo, desusuario, senha, status, cnpj_cpf, admin) VALUES ('"& Request.Form("cod_grupo") & "', '"& Tira_Pontuacao(Request.Form("login")) & "', '"& Tira_Pontuacao(Request.Form("senha")) & "', "& replace(Request.Form("status"),"'","") & ", '" & request.Form("cnpj_cpf") & "', 9)"
	end if	
end if	


	if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> "" then
	   conn.execute "update gedproposta set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
	   conn.execute "update gedparecer set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
   	   conn.execute "update gedos set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
   	   conn.execute "update gedcheque set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
   	   conn.execute "update GEDARQDOC set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
   	   conn.execute "update gedclasseos set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
   	   conn.execute "update gedmarcaos set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
   	   conn.execute "update gedhist_atend set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
   	   conn.execute "update gedhist_atend1 set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"
   	   conn.execute "update gedhist_cliente set cnpj_cpf = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where cnpj_cpf = '" & request.Form("cnpj_cpf") & "'"

	 if isnumeric(request.Form("cnpj_cpf")) then
	   if (request.Form("cnpj_cpf")) > 100000 then
     	   conn.execute "update lancamentos set codigo = '" & Tira_Pontuacao(Request.Form("cnpj_cpf_novo")) & "' where codigo = '" & request.Form("cnpj_cpf") & "'"
	   end if
	 end if
	end if 	   
	
	if (Request.Form("cnpj_cpf_novo")="") then
		cpf_cnpj = ""
	else
		cpf_cnpj = trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo")))
	end if
	
	
	SQL = "update GEDCLIENTES set "
	sql1 = "cod_grupo = '" & Request.Form("cod_grupo") & "', "
	if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> "" then
	   sql1 = sql1 & "CNPJ_CPF='" & cpf_cnpj & "', "
    end if
	sql1 = sql1 & "rg='" & Tira_Pontuacao(request.Form("rg")) & "', "
	sql1 = sql1 & "data_nasc='" & DATE_US(request.Form("data_nasc")) & "', "
	if Request.Form("razao_fund") = "1" then
		if trim(request.Form("data_razao")) = "" then
			sql1 = sql1 & "razao1 = 1, razao2 = '" & DATE_US(request.Form("data_nasc")) & "',"
		else
			sql1 = sql1 & "razao1 = 1, razao2 = '" & DATE_US(request.Form("data_razao")) & "',"
		end if
	else
		sql1 = sql1 & "razao1 = 0, razao2 = '" & DATE_US(request.Form("data_nasc")) & "',"
	end if
	
	sql1 = sql1 & "data_ult_alt='" & DATE_US(date()) & "', "
	sql1 = sql1 & "data_negociacao='" & DATE_US(date()) & "', "
	sql1 = sql1 & "nome='" & Replace(ucase(request.Form("nome")),"'","''") & "', "
	sql1 = sql1 & "nome_fantasia='" & ucase(Replace(request.Form("nome_fantasia"),"'","''")) & "', "
	sql1 = sql1 & "sexo='" & request.Form("sexo") & "', "
	
	sql1 = sql1 & "endereco='" & Tira_Pontuacao1(request.Form("endereco")) & "', "
	sql1 = sql1 & "numero='" & Tira_Pontuacao1(request.Form("numero")) & "', "
	sql1 = sql1 & "complemento='" & Tira_Pontuacao1(request.Form("complemento")) & "', "
	sql1 = sql1 & "bairro='" & Tira_Pontuacao1(request.Form("bairro")) & "', "
	sql1 = sql1 & "cep='" & request.Form("cep") & "', "
	sql1 = sql1 & "cidade='" & Tira_Pontuacao1(request.Form("txt_cidade")) & "', "
	sql1 = sql1 & "estado='" & request.Form("txt_estado") & "', "
	
	if trim(request.form("cod_ass")) <> "" then
		sql1 = sql1 & "cod_ass='" & request.Form("cod_ass") & "', "
	else
		sql1 = sql1 & "cod_ass='', "
    end if		
	if isnumeric(Request.Form("capital")) then
		sql1 = sql1 & "capital='" & request.Form("capital") & "', "
	else
		sql1 = sql1 & "capital='0', "
	end if
	sql1 = sql1 & "endereco_inpi='" & Mid(Tira_Pontuacao1(request.Form("endereco_inpi")),1,50) & "', "
	sql1 = sql1 & "numero_inpi='" & request.Form("numero_inpi") & "', "
	sql1 = sql1 & "complemento_inpi='" & Tira_Pontuacao1(request.Form("complemento_inpi")) & "', "
	sql1 = sql1 & "bairro_inpi='" & Tira_Pontuacao1(request.Form("bairro_inpi")) & "', "
	sql1 = sql1 & "cep_inpi='" & request.Form("cep_inpi") & "', "
	sql1 = sql1 & "cidade_inpi='" & Tira_Pontuacao1(request.Form("txt_cidade_inpi")) & "', "
	sql1 = sql1 & "estado_inpi='" & request.Form("txt_estado_inpi") & "', "
	
	sql1 = sql1 & "endereco_corr='" & Tira_Pontuacao1(request.Form("endereco_corr")) & "', "
	sql1 = sql1 & "numero_corr='" & request.Form("numero_corr") & "', "
	sql1 = sql1 & "complemento_corr='" & Tira_Pontuacao1(request.Form("complemento_corr")) & "', "
	sql1 = sql1 & "bairro_corr='" & Tira_Pontuacao1(request.Form("bairro_corr")) & "', "
	sql1 = sql1 & "cep_corr='" & request.Form("cep_corr") & "', "
	sql1 = sql1 & "cidade_corr='" & Tira_Pontuacao1(request.Form("txt_cidade_corr")) & "', "
	sql1 = sql1 & "estado_corr='" & request.Form("txt_estado_corr") & "', "

	sql1 = sql1 & "status_marca=" & wstatus_marca & ", "
	sql1 = sql1 & "pais='" & request.Form("pais") & "', "
	sql1 = sql1 & "fone_res='" & Tira_Pontuacao1(request.Form("fone_res")) & "', "
	sql1 = sql1 & "fone_com='" & Tira_Pontuacao1(request.Form("fone_com")) & "', "
	sql1 = sql1 & "fax='" & Tira_Pontuacao1(request.Form("fax")) & "', "
	sql1 = sql1 & "celular='" & request.Form("celular") & "', "
	sql1 = sql1 & "radio='" & request.Form("radio") & "', "
	sql1 = sql1 & "gaveta_pasta='" & request.Form("gaveta_pasta") & "', "
	sql1 = sql1 & "site='" & Tira_Pontuacao1(request.Form("site")) & "', "
	sql1 = sql1 & "email1='" & Tira_Pontuacao1(request.Form("email1")) & "', "
	sql1 = sql1 & "email2='" & Tira_Pontuacao1(request.Form("email2")) & "', "
	sql1 = sql1 & "contato1='" & Tira_Pontuacao1(request.Form("contato1")) & "', "
	sql1 = sql1 & "contato2='" & Tira_Pontuacao1(request.Form("contato2")) & "', "
	sql1 = sql1 & "contato3='" & Tira_Pontuacao1(request.Form("contato3")) & "', "
	sql1 = sql1 & "contato4='" & Tira_Pontuacao1(request.Form("contato4")) & "', "
	sql1 = sql1 & "cargo1='" & Tira_Pontuacao1(request.Form("cargo1")) & "', "
	sql1 = sql1 & "cargo2='" & Tira_Pontuacao1(request.Form("cargo2")) & "', "
	sql1 = sql1 & "cargo3='" & Tira_Pontuacao1(request.Form("cargo3")) & "', "
	sql1 = sql1 & "cargo4='" & Tira_Pontuacao1(request.Form("cargo4")) & "', "
	sql1 = sql1 & "autor='" & Tira_Pontuacao1(request.Form("autor")) & "', "
	sql1 = sql1 & "inventor='" & Tira_Pontuacao1(request.Form("inventor")) & "', "
	sql1 = sql1 & "aniver1='" & request.Form("aniver1") & "', "
	sql1 = sql1 & "aniver2='" & request.Form("aniver2") & "', "
	sql1 = sql1 & "aniver3='" & request.Form("aniver3") & "', "
	sql1 = sql1 & "aniver4='" & request.Form("aniver4") & "', "
	sql1 = sql1 & "email_contato1='" & request.Form("email_contato1") & "', "
	sql1 = sql1 & "email_msn='" & request.Form("email_msn") & "', "
	sql1 = sql1 & "skype='" & request.Form("skype") & "', "
	sql1 = sql1 & "orkut='" & Tira_Pontuacao1(request.Form("orkut")) & "', "
	sql1 = sql1 & "email_contato2='" & request.Form("email_contato2") & "', "
	sql1 = sql1 & "email_contato3='" & request.Form("email_contato3") & "', "
	sql1 = sql1 & "email_contato4='" & request.Form("email_contato4") & "', "
	sql1 = sql1 & "fone_contato1='" & request.Form("fone_contato1") & "', "
	sql1 = sql1 & "fone_contato2='" & request.Form("fone_contato2") & "', "
	sql1 = sql1 & "fone_contato3='" & request.Form("fone_contato3") & "', "
	sql1 = sql1 & "fone_contato4='" & request.Form("fone_contato4") & "', "
	sql1 = sql1 & "fone2_contato1='" & request.Form("fone2_contato1") & "', "
	sql1 = sql1 & "fone2_contato2='" & request.Form("fone2_contato2") & "', "
	sql1 = sql1 & "fone2_contato3='" & request.Form("fone2_contato3") & "', "
	sql1 = sql1 & "fone2_contato4='" & request.Form("fone2_contato4") & "', "
	sql1 = sql1 & "nome_indicacao='" & request.Form("nome_indicacao") & "', "
	sql1 = sql1 & "grupo_fin='" & request.Form("grupo_cliente") & "', "
	sql1 = sql1 & "status='" & request.Form("status") & "', "
	if session("admin") = 1 then
		sql1 = sql1 & "perfil=" & request.Form("perfil") & ", "
	end if	
	sql1 = sql1 & "cod_sub='" & request.Form("cod_sub") & "', "
	sql1 = sql1 & "cod_sub_patente='" & wcod_sub_patente & "', "
	sql1 = sql1 & "cod_sub_avaliacao='" & wcod_sub_avaliacao & "', "
	sql1 = sql1 & "cod_sub_internacional ='" & wcod_sub_internacional & "', "
	sql1 = sql1 & "cod_sub_software ='" & wcod_sub_software & "', "
	sql1 = sql1 & "cod_sub_direito ='" & wcod_sub_direito & "', "
	sql1 = sql1 & "cod_sub_juridico ='" & wcod_sub_juridico & "', "
	sql1 = sql1 & "cod_sub_cob ='" & wcod_sub_cob & "', "
	sql1 = sql1 & "bloco_nota ='" & request.form("bloco_nota") & "', "
	sql1 = sql1 & "perfil1 ='" & request.form("perfil") & "', "
	sql1 = sql1 & "ramo ='" & request.form("ramo") & "', "
	sql1 = sql1 & "origem1 ='" & request.form("origem1") & "', "
	sql1 = sql1 & "origem2 ='" & request.form("origem2") & "', "
	
	' alteração Elano 
	sql1 = sql1 & "EMITENOTA='" & request.Form("emite_nota") & "', "
	sql1 = sql1 & "SUBST_TRIB='" & request.Form("substituto") & "', "
	sql1 = sql1 & "tipo_tributacao='" & request.Form("tipo_tributacao") & "', "
	sql1 = sql1 & "percentual_ISS='" & request.Form("percentual_ISS") & "', "
	
	' Fim da Alteração 
	
	sql1 = sql1 & "numero_ficha ='" & request.form("numero_ficha") & "' "
	sql1 = sql1 & "where CNPJ_CPF ='" & request.Form("cnpj_cpf") & "'"
	sqlFinal = sql & sql1

	conn.execute (sqlFinal)
	'response.write(sqlFinal)


	
'Alteração Elano - Atualização de cliente na base do alterdata  

	Dim DbConn, rs , rserp
	Dim sql , SQL1, SQL2, SQL3 
	dim CodPessoa , tipoPes , CNPJCPF_P
	dim nomepes , nomecurtopes , varcpfcnpj
	Dim ender, tipolog , codenderecopes 
	Dim email,email2, Telefone, Celular 
    Dim TelefoneCont1, CelularCont1, emailCont1 
    Dim TelefoneCont2, CelularCont2, emailCont2
    Dim TelefoneCont3, CelularCont3, emailCont3
    Dim TelefoneCont4, CelularCont4, emailCont4 


   Set DbConn = Server.CreateObject("ADODB.Connection")
   Set rserp = Server.CreateObject("ADODB.Recordset")
   Set rserp1 = Server.CreateObject("ADODB.Recordset")
   Set rs = Server.CreateObject("ADODB.Recordset")

   'DbConn.Open "Provider=SQLOLEDB;DATABASE=BIMER_ALTERDATA_TEST;server=SRVFINANCEIRO,1433;UID=livia;PASSWORD=123456"
   DbConn.Open "Provider=SQLOLEDB;DATABASE=BIMER_ALTERDATA;server=localhost;UID=livia;PASSWORD=123456"  'Mudar em produção
   
   CNPJCPF_P = FormataCNPJCPF(request.Form("cnpj_cpf"))

   'Response.write(CNPJCPF_P)  & "<BR>"    
	
	SQL1 = "SELECT  IdPessoa, CdChamada, TpPessoa, CdCPF_CGC, NmPessoa,  NmCurto FROM  dbo.Pessoa where CdCPF_CGC= '" & CNPJCPF_P & "'"
	rserp.open SQL1, DbConn
	if not rserp.eof then
	CodPessoa = rserp("IdPessoa")
 '	Response.write(CodPessoa)  & "<BR>" 
	end if  
    rserp.close 
	
	tipoPes = "J"  ' Não precisa para atualizar, apenas para não passar o paramentro em branco
	 
     nomepes = request.Form("nome")
	 nomecurtopes = request.Form("nome_fantasia")
	 varcpfcnpj = FormataCNPJCPF(request.Form("cnpj_cpf"))
	 
     SQL = "execute Insere_Pessoa  '" & CodPessoa & "', '" & tipoPes & "', '" & varcpfcnpj &"','" & nomepes &"',  '"& nomecurtopes &"'"
     set rs = DbConn.Execute(SQL)
	 
	 ' Atualiza tabela de pessoa endereço 
	 tipolog = "-"
	 
	 SQL2 = "execute Insere_PessoaEndereco  '"& CodPessoa &"', '"& "01" &"', '"& request.Form("endereco")	&"', '"& request.Form("numero") &"',  '"& request.Form("complemento") &"', '"& request.Form("rg") &"', '"& tipolog &"' , '"& request.Form("cep") &"', '"& request.Form("txt_estado") & "' ,  '"& request.Form("bairro") &"' , '"& request.Form("txt_cidade") &"' "
	 set rs = DbConn.Execute(SQL2)
	 
	
	' Atualiza tabela de pessoa endereço 
'---------------------
	
	 'Caso existam registros para este clientes  serão excluidos
	 SQL2 = "DELETE FROM dbo.PessoaEndereco_Contato WHERE dbo.PessoaEndereco_Contato.IdPessoa =  '" & CodPessoa &"'"
     set rs = DbConn.Execute(SQL2)

	 SQL2 = "DELETE FROM dbo.PessoaEndereco_TipoContato WHERE dbo.PessoaEndereco_TipoContato.IdPessoa =  '" & CodPessoa &"'"
     set rs = DbConn.Execute(SQL2)
	

					'insere pessoa endereco Principal  
					SQL3 =  "execute  stp_GetMultiCode  PessoaEndereco_Contato , IdPessoaEndereco_Contato , 1 "
					rserp1.open SQL3, DbConn
					if not rserp1.eof then
					codcontato = rserp1(0)  ' pega paramentro para o incremento
					end if     
					rserp1.close 
									
					'insere pessoa endereço 
					SQL3 = "INSERT INTO PessoaEndereco_Contato (IdPessoaEndereco_Contato, IdPessoa, CdEndereco, DsContato, StContatoPrincipal) VALUES('"& codcontato &"', '"& CodPessoa &"', '"& "01" &"', '"& nomecurtopes &"','"& "S" &"')"
					set rs = DbConn.Execute(SQL3)
					
					
					if request.Form("fone_com") <> "" then
					Telefone = request.Form("fone_com")
					else
					Telefone = "AT.FONE"
					end if 
				  
				    if request.Form("celular") <> "" then
				    Celular = request.Form("celular")
				    else
				    Celular = "AT.CELULAR"
				    end if 
				  
				    if request.Form("email1") <> "" then
				    email = request.Form("email1")
					email = mid (email,1,50)
				    else
				    email = "AT.EMAIL"
				    end if 
					
					if request.Form("email2") <> "" then
				    email2 = request.Form("email2")
					email2 = mid (email2,1,50)
				    else
				    email2 = "AT.EMAIL 2"
				    end if 
			  
				   'insere o telefone do contato
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('"& codcontato &"', '"& "0000000001" &"', '"& CodPessoa &"', '"& "01" &"', '"& Telefone &"')"
				   set rs = DbConn.Execute(SQL3)
				   			   
				   'insere o celular do contato       
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('"& codcontato &"', '"& "0000000003" &"', '"& CodPessoa &"', '"& "01" &"', '"& celular &"' )"
				   set rs = DbConn.Execute(SQL3)
											
				   'insere o email do contato
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000004" & "', '" & CodPessoa & "', '" & "01" & "', '" & email & "' )"
				   set rs = DbConn.Execute(SQL3)
					
            	   '0000000009 email de cobrança
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000009" & "', '" & CodPessoa & "', '" & "01" & "', '" & email2 & "' )"
				   set rs = DbConn.Execute(SQL3)
					   
  			      'fim do contato Prinicipal

				  
			   '---------------------------------------------------------------------------
				
			      'insere pessoa endereco contato1  
					
					if request.Form("contato1") <> "" then
					
					codcontato = ""
					SQL1 =  "execute  stp_GetMultiCode  PessoaEndereco_Contato , IdPessoaEndereco_Contato , 1 "
					rserp.open SQL1, DbConn
					if not rserp.eof then
					codcontato = rserp(0)  ' pega paramentro para o incremento
					end if     
					rserp.close 
					
					'insere pessoa endereço 
					SQL3 = "INSERT INTO PessoaEndereco_Contato (IdPessoaEndereco_Contato, IdPessoa, CdEndereco, DsContato, StContatoPrincipal) VALUES('"& codcontato &"', '"& CodPessoa &"', '" & "01" & "', '"& request.Form("contato1") &"','"& "N" &"' )"
					set rserp = DbConn.Execute(SQL3)
					
					
					if request.Form("fone_contato1") <> "" then
					TelefoneCont1 = request.Form("fone_contato1")
					else
					TelefoneCont1 = "AT.FONE"
					end if 
				  
				    if request.Form("fone2_contato1") <> "" then
				    CelularCont1 = request.Form("fone2_contato1")
				    else
				    CelularCont1 = "AT.CELULAR"
				    end if 
				  
				    if request.Form("email_contato1") <> "" then
				    emailCont1 = request.Form("email_contato1")
					emailCont1 = mid (emailCont1,1,50)
				    else
				    emailCont1 = "AT.EMAIL"
				    end if 
				  
				     'insere o telefone do contato1
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000001" & "', '" & CodPessoa & "', '" & "01" & "', '" & TelefoneCont1 & "' )"
				   set rserp = DbConn.Execute(SQL3)
				              
				   'insere o celular do contato1       
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000003" & "', '" & CodPessoa & "', '" & "01" & "', '" & CelularCont1 & "' )"
				   set rserp = DbConn.Execute(SQL3)
											
				'   'insere o email do contato1
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000004" & "', '" & CodPessoa & "', '" & "01" & "', '" & emailCont1 & "' )"
				   set rserp = DbConn.Execute(SQL3)
				
				  else  
					
				   End If
				
 				   'fim do contato 1
				  
				'   Response.write("Passou pelo contato 1 ") & "<BR>"
				'---------------------------------------------------------------------------

				   'insere pessoa endereco contato2  
					
					if request.Form("contato2") <> "" then
					
					codcontato = ""
					SQL1 =  "execute  stp_GetMultiCode  PessoaEndereco_Contato , IdPessoaEndereco_Contato , 1 "
					rserp.open SQL1, DbConn
					if not rserp.eof then
					codcontato = rserp(0)  ' pega paramentro para o incremento
					end if     
					rserp.close 
					
					'insere pessoa endereço 2 
					SQL3 = "INSERT INTO PessoaEndereco_Contato (IdPessoaEndereco_Contato, IdPessoa, CdEndereco, DsContato, StContatoPrincipal) VALUES('" & codcontato & "', '" & CodPessoa & "', '" & "01" & "', '" & request.Form("contato2") & "','" & "N" & "' )"
					set rserp = DbConn.Execute(SQL3)
					
					
					if request.Form("fone_contato2") <> "" then
					TelefoneCont2 = request.Form("fone_contato2")
					else
					TelefoneCont2 = "AT.FONE"
					end if 
				  
				    if request.Form("fone2_contato2") <> "" then
				    CelularCont2 = request.Form("fone2_contato2")
				    else
				    CelularCont2 = "AT.CELULAR"
				    end if 
				  
				    if request.Form("email_contato2") <> "" then
				    emailCont2 = request.Form("email_contato2")
					emailCont2 = mid (emailCont2,1,50)
				    else
				    emailCont2 = "AT.EMAIL"
				    end if 
				  
				     'insere o telefone do contato2
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000001" & "', '" & CodPessoa & "', '" & "01" & "', '" & TelefoneCont2 & "' )"
				   set rserp = DbConn.Execute(SQL3)
				              
				   'insere o celular do contato2       
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000003" & "', '" & CodPessoa & "', '" & "01" & "', '" & CelularCont2 & "' )"
				   set rserp = DbConn.Execute(SQL3)
											
				   'insere o email do contato2
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000004" & "', '" & CodPessoa & "', '" & "01" & "', '" & emailCont2 & "' )"
				   set rserp = DbConn.Execute(SQL3)
				
 				   else  
				   
				   End If
				  'fim do contato 2
                 
				' Response.write("Passou pelo contato 2 ") & "<BR>"
				'---------------------------------------------------------------------------
				
				'insere pessoa endereco contato3  
					
					if request.Form("contato3") <> "" then
					
					codcontato = ""
					SQL1 =  "execute  stp_GetMultiCode  PessoaEndereco_Contato , IdPessoaEndereco_Contato , 1 "
					rserp.open SQL1, DbConn
					if not rserp.eof then
					codcontato = rserp(0)  ' pega paramentro para o incremento
					end if     
					rserp.close 
					
					'insere pessoa endereço 3 
					SQL3 = "INSERT INTO PessoaEndereco_Contato (IdPessoaEndereco_Contato, IdPessoa, CdEndereco, DsContato, StContatoPrincipal) VALUES('" & codcontato & "', '" & CodPessoa & "', '" & "01" & "', '" & request.Form("contato3") & "','" & "N" & "' )"
					set rserp = DbConn.Execute(SQL3)
					
					
					if request.Form("fone_contato3") <> "" then
					TelefoneCont3 = request.Form("fone_contato3")
					else
					TelefoneCont3 = "AT.FONE"
					end if 
				  
				    if request.Form("fone2_contato3") <> "" then
				    CelularCont3 = request.Form("fone2_contato3")
				    else
				    CelularCont3 = "AT.CELULAR"
				    end if 
				  
				    if request.Form("email_contato3") <> "" then
				    emailCont3 = request.Form("email_contato3")
					emailCont3 = mid (emailCont3,1,50) 
				    else
				    emailCont3 = "AT.EMAIL"
				    end if 
				  
				     'insere o telefone do contato3
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000001" & "', '" & CodPessoa & "', '" & "01" & "', '" & TelefoneCont3 & "' )"
				   set rserp = DbConn.Execute(SQL3)
				              
				   'insere o celular do contato3       
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000003" & "', '" & CodPessoa & "', '" & "01" & "', '" & CelularCont3 & "' )"
				   set rserp = DbConn.Execute(SQL3)
											
				   'insere o email do contato3
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000004" & "', '" & CodPessoa & "', '" & "01" & "', '" & emailCont3 & "' )"
				   set rserp = DbConn.Execute(SQL3)
				
                    else  
				   
				   End If
				   
 				  'fim do contato 3
				
				'Response.write("Passou pelo contato 3 ") & "<BR>"
				
				'---------------------------------------------------------------------------
				
					'insere pessoa endereco contato4  
					if request.Form("contato4") <> "" then
					
					codcontato = ""
					SQL1 =  "execute  stp_GetMultiCode  PessoaEndereco_Contato , IdPessoaEndereco_Contato , 1 "
					rserp.open SQL1, DbConn
					if not rserp.eof then
					codcontato = rserp(0)  ' pega paramentro para o incremento
					end if     
					rserp.close 
					
					'insere pessoa endereço 4 
					SQL3 = "INSERT INTO PessoaEndereco_Contato (IdPessoaEndereco_Contato, IdPessoa, CdEndereco, DsContato, StContatoPrincipal) VALUES('" & codcontato & "', '" & CodPessoa & "', '" & "01" & "', '" & request.Form("contato4") & "','" & "N" & "' )"
					set rserp = DbConn.Execute(SQL3)
					
					
					if request.Form("fone_contato4") <> "" then
					TelefoneCont4 = request.Form("fone_contato4")
					else
					TelefoneCont4 = "AT.FONE"
					end if 
				  
				    if request.Form("fone2_contato4") <> "" then
				    CelularCont4 = request.Form("fone2_contato4")
				    else
				    CelularCont4 = "AT.CELULAR"
				    end if 
				  
				    if request.Form("email_contato4") <> "" then
				    emailCont4 = request.Form("email_contato4")
					emailCont4 = mid (emailCont4,1,50) 
				    else
				    emailCont4 = "AT.EMAIL"
				    end if 
				  
				   'insere o telefone do contato4
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000001" & "', '" & CodPessoa & "', '" & "01" & "', '" & TelefoneCont4 & "' )"
				   set rserp = DbConn.Execute(SQL3)
				              
				   'insere o celular do contato4       
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000003" & "', '" & CodPessoa & "', '" & "01" & "', '" & CelularCont4 & "' )"
				   set rserp = DbConn.Execute(SQL3)
											
				   'insere o email do contato4
				   SQL3 = "INSERT INTO PessoaEndereco_TipoContato(IdPessoaEndereco_Contato, IdTipoContato, IdPessoa, CdEndereco, DsContato) VALUES ('" & codcontato & "', '" & "0000000004" & "', '" & CodPessoa & "', '" & "01" & "', '" & emailCont4 & "' )"
				   set rserp = DbConn.Execute(SQL3)
				
				   else  
				   
				   End If
				   
 				  'fim do contato 4
				
	  'Response.write("Passou pelo contato 4 ") & "<BR>"

	  '----------------------- Fim da integração 

	
	if Request.Form("representante1") <> "" then 
		rs.open ("select * from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=1"), conn, 3, 3
		if rs.EOF then
			rs.AddNew
				rs("CNPJ_CPF") = Tira_Pontuacao(request.Form("cnpj_cpf"))
				rs("CODIGO") = "1"
				rs("NOME_SOCIO") = Request.Form("representante1")
				rs("CARGO") = Request.Form("cargo_representante1")
				rs("ANIVER") = Request.Form("aniver_representante1")
				rs("RG_SOCIO") = Request.Form("rg_representante1")
				rs("CPF_SOCIO") = Tira_Pontuacao(Request.Form("cpf_representante1"))
				rs("EMAIL") = Request.Form("email_representante1")
				rs("PROFISSAO") = Request.Form("prof_representante1")
				rs("STATUS_SOCIO") = "0"
			rs.Update
		else
			if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> "" then
				conn.execute "update gedsocios set profissao = '" & replace(request.form("prof_representante1"),"'","") & "', email = '" & Request.Form("email_representante1") & "', cpf_socio = '" & Request.Form("cpf_representante1") & "', rg_socio = '" &Request.Form("rg_representante1") & "', nome_socio = '" & Request.Form("representante1") & "', cargo = '"& Request.Form("cargo_representante1") & "', aniver = '" & request.Form("aniver_representante1") & "', cnpj_cpf = '" & trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) & "', status_socio = '0'  where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo1") & "'"
			else
				conn.execute "update gedsocios set profissao = '" & replace(request.form("prof_representante1"),"'","") & "', email = '" & Request.Form("email_representante1") & "', cpf_socio = '" & Request.Form("cpf_representante1") & "', rg_socio = '" &Request.Form("rg_representante1") & "', nome_socio = '" & Request.Form("representante1") & "', cargo = '"& Request.Form("cargo_representante1") & "', aniver = '" & request.Form("aniver_representante1") & "', status_socio = '0' where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo1") & "'"
			end if			
		End if
		rs.close
	else
		conn.execute("delete from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=1")		
	End if
	wstatus_socio2 = Replace(Request.Form("STATUS_SOCIO2"),"0, ","")
	
	if Request.Form("representante2") <> "" then 
		rs.open ("select * from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=2"), conn, 3, 3
		if rs.EOF then
			rs.AddNew
				rs("CNPJ_CPF") = Tira_Pontuacao(request.Form("cnpj_cpf"))
				rs("CODIGO") = "2"
				rs("NOME_SOCIO") = Request.Form("representante2")
				rs("CARGO") = Request.Form("cargo_representante2")
				rs("ANIVER") = Request.Form("aniver_representante2")
				rs("RG_SOCIO") = Request.Form("rg_representante2")
				rs("CPF_SOCIO") = Tira_Pontuacao(Request.Form("cpf_representante2"))
				rs("EMAIL") = Request.Form("email_representante2")
				rs("PROFISSAO") = Request.Form("prof_representante2")
				if isnumeric(wstatus_socio2) then
					rs("STATUS_SOCIO") = cint(wstatus_socio2)
				else
					rs("STATUS_SOCIO") = 0
				end if
				
			rs.Update
			else
	if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> "" then
		 if Request.Form("codigo2")<>"" then
		 	conn.execute "update gedsocios set profissao = '" & replace(request.form("prof_representante2"),"'","") & "', email = '" & Request.Form("email_representante2") & "', cpf_socio = '" & Request.Form("cpf_representante2") & "', rg_socio = '" & Request.Form("rg_representante2") & "', nome_socio = '" & Request.Form("representante2") & "', cargo = '"& Request.Form("cargo_representante2") & "', aniver = '" & request.Form("aniver_representante2") & "', cnpj_cpf = '" & trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) & "', status_socio = '" & cint(wstatus_socio2) & "' where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo2") & "'"
		end if
	else
		if Request.Form("codigo2")<>"" then
			conn.execute "update gedsocios set profissao = '" & replace(request.form("prof_representante2"),"'","") & "', email = '" & Request.Form("email_representante2") & "', cpf_socio = '" & Request.Form("cpf_representante2") & "', rg_socio = '" & Request.Form("rg_representante2") & "', nome_socio = '" & Request.Form("representante2") & "', cargo = '"& Request.Form("cargo_representante2") & "', aniver = '" & request.Form("aniver_representante2") & "', status_socio = '" & cint(wstatus_socio2) & "' where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo2") & "'"
		end if
    end if			
		End if
		rs.close
		ELSE
		conn.execute("delete from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=2")
	End if
	
	if Trim(Request.Form("representante3")) <> "" then 
		rs.open ("select * from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=3"), conn, 3, 3
		if rs.EOF then
			rs.AddNew
				rs("CNPJ_CPF") = Tira_Pontuacao(request.Form("cnpj_cpf"))
				rs("CODIGO") = "3"
				rs("NOME_SOCIO") = Request.Form("representante3")
				rs("CARGO") = Request.Form("cargo_representante3")
				rs("ANIVER") = Request.Form("aniver_representante3")
				rs("RG_SOCIO") = Request.Form("rg_representante3")
				rs("CPF_SOCIO") = Tira_Pontuacao(Request.Form("cpf_representante3"))
				rs("EMAIL") = Request.Form("email_representante3")
				rs("PROFISSAO") = Request.Form("prof_representante3")
				Response.Write Request.Form("STATUS_SOCIO3")
				wstatus3 = Trim(Request.Form("STATUS_SOCIO3"))
				if wstatus3 > 0then
					rs("STATUS_SOCIO") = Request.Form("STATUS_SOCIO3")
				else
					rs("STATUS_SOCIO") = "0"
				end if
			rs.Update
			else
	if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> "" then
			conn.execute "update gedsocios set profissao = '" & request.form("prof_representante3") & "', email = '" & Request.Form("email_representante3") & "', cpf_socio = '" & Request.Form("cpf_representante3") & "', rg_socio = '" & Request.Form("rg_representante3") & "', nome_socio = '" & Request.Form("representante3") & "', cargo = '"& Request.Form("cargo_representante3") & "', aniver = '" & request.Form("aniver_representante3") & "', cnpj_cpf = '" & trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) & "', status_socio = " & request.form("status_socio3") & " where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo3") & "'"
	else
			conn.execute "update gedsocios set profissao = '" & request.form("prof_representante3") & "', email = '" & Request.Form("email_representante3") & "', cpf_socio = '" & Request.Form("cpf_representante3") & "', rg_socio = '" & Request.Form("rg_representante3") & "', nome_socio = '" & Request.Form("representante3") & "', cargo = '"& Request.Form("cargo_representante3") & "', aniver = '" & request.Form("aniver_representante3") & "', status_socio = '" & request.form("status_socio3") & "' where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo3") & "'"
   end if			
		End if
		rs.close
		ELSE
		conn.execute("delete from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=3")
	End if
	
	
	if Request.Form("representante4") <> "" then 
		rs.open ("select * from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=4"), conn, 3, 3
		if rs.EOF then
			rs.AddNew
				rs("CNPJ_CPF") = Tira_Pontuacao(request.Form("cnpj_cpf"))
				rs("CODIGO") = "4"
				rs("NOME_SOCIO") = Request.Form("representante4")
				rs("CARGO") = Request.Form("cargo_representante4")
				rs("ANIVER") = Request.Form("aniver_representante4")
				rs("RG_SOCIO") = Request.Form("rg_representante4")
				rs("CPF_SOCIO") = Tira_Pontuacao(Request.Form("cpf_representante4"))
				rs("EMAIL") = Request.Form("email_representante4")
				rs("PROFISSAO") = Request.Form("prof_representante4")
				rs("STATUS_SOCIO") = Request.Form("STATUS_SOCIO4")
			rs.Update
		else
			if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> "" then
			
		
					conn.execute "update gedsocios set profissao = '" & replace(request.form("prof_representante4"),"'","") & "', email = '" & Request.Form("email_representante4") & "', cpf_socio = '" & Request.Form("cpf_representante4") & "', rg_socio = '" &Request.Form("rg_representante4") & "', nome_socio = '" & Request.Form("representante4") & "', cargo = '"& Request.Form("cargo_representante4") & "', aniver = '" & request.Form("aniver_representante4") & "', cnpj_cpf = '" & trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) & "', status_socio = '" & request.form("status_socio4") & "'  where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo4") & "'"
			else
					conn.execute "update gedsocios set profissao = '" & replace(request.form("prof_representante4"),"'","") & "', email = '" & Request.Form("email_representante4") & "', cpf_socio = '" & Request.Form("cpf_representante4") & "', rg_socio = '" &Request.Form("rg_representante4") & "', nome_socio = '" & Request.Form("representante4") & "', cargo = '"& Request.Form("cargo_representante4") & "', aniver = '" & request.Form("aniver_representante4") & "', status_socio = " & request.form("status_socio4") & "  where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo4") & "'"
			end if			
		End if
		rs.close
	else
		conn.execute("delete from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=4")		
	End if
	
	if Request.Form("representante5") <> "" then 
		rs.open ("select * from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=5"), conn, 3, 3
		if rs.EOF then
			rs.AddNew
				rs("CNPJ_CPF") = Tira_Pontuacao(request.Form("cnpj_cpf"))
				rs("CODIGO") = "5"
				rs("NOME_SOCIO") = Request.Form("representante5")
				rs("CARGO") = Request.Form("cargo_representante5")
				rs("ANIVER") = Request.Form("aniver_representante5")
				rs("RG_SOCIO") = Request.Form("rg_representante5")
				rs("CPF_SOCIO") = Tira_Pontuacao(Request.Form("cpf_representante5"))
				rs("EMAIL") = Request.Form("email_representante5")
				rs("PROFISSAO") = Request.Form("prof_representante5")
				rs("STATUS_SOCIO") = Request.Form("STATUS_SOCIO5")
			rs.Update
		else
			if trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) <> "" then
			
		
					conn.execute "update gedsocios set profissao = '" & replace(request.form("prof_representante5"),"'","") & "', email = '" & Request.Form("email_representante5") & "', cpf_socio = '" & Request.Form("cpf_representante5") & "', rg_socio = '" &Request.Form("rg_representante5") & "', nome_socio = '" & Request.Form("representante5") & "', cargo = '"& Request.Form("cargo_representante5") & "', aniver = '" & request.Form("aniver_representante5") & "', cnpj_cpf = '" & trim(Tira_Pontuacao(Request.Form("cnpj_cpf_novo"))) & "', status_socio = '" & request.form("status_socio5") & "'  where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo5") & "'"
			else
					conn.execute "update gedsocios set profissao = '" & replace(request.form("prof_representante5"),"'","") & "', email = '" & Request.Form("email_representante5") & "', cpf_socio = '" & Request.Form("cpf_representante5") & "', rg_socio = '" &Request.Form("rg_representante5") & "', nome_socio = '" & Request.Form("representante5") & "', cargo = '"& Request.Form("cargo_representante5") & "', aniver = '" & request.Form("aniver_representante5") & "', status_socio = " & request.form("status_socio5") & "  where CNPJ_CPF = '" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and CODIGO = '" & Request.Form("codigo5") & "'"
			end if			
		End if
		rs.close
	else
		conn.execute("delete from gedsocios where cnpj_cpf='" & Tira_Pontuacao(request.Form("cnpj_cpf")) & "' and codigo=5")		
	End if


if Err.Number <> 0 then
    conn.RollBackTrans
    response.write Err.Description
else
	call Val_Audit_Registra(session("usuario_registra"),"ALTERAÇÃO DE CLIENTE","O cliente " & ucase(Request.Form("nome")) & " foi alterado.",Tira_Pontuacao1(request.Form("nome")),session("franquia"))
    conn.committrans

	'Call CaixaMensagem ("Ok","ALTERAÇÃO REALIZADA COM SUCESSO!","O Cliente: " & ucase(Request.Form("nome")) & " foi atualizado com sucesso!!",true,"Avançar",false,"LISTAGEM.asp")
	
	Call CaixaMensagem ("Ok","ALTERAÇÃO REALIZADA COM SUCESSO!","O Cliente: " & ucase(Request.Form("nome")) & " foi atualizado com sucesso!!",true,"Avançar",false,"")
		
end if

%>
<!--#include virtual="/registrafacil/include/db_close.asp"-->