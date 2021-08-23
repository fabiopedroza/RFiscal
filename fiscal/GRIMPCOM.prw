#INCLUDE "PROTHEUS.CH"
#include "RWMAKE.ch"
#include "Topconn.ch"
#INCLUDE 'FWMVCDef.ch'

User Function GRIMPCOM()

	//programa para importar notas

	Local Nz, NX, x		:= 0
	Local aLogX			:= {}			// array para os logs

	Private cCnpjEmit	:= ""
	Private cMsgErro
	Private _xFornece 	:= ""
	Private _xloja    	:= ""
	Private cDescProd	:= ""
	Private cFile 		:= "" // Space(10)
	Private Caminho 	:= ""
	Private CamProc		:= ""
	Private CamErro		:= ""
	Private cArq
	Private cArq2
	Private nHdl
	Private	aFiles		:= {}
	Private lPcNfe		:= GETMV("MV_PCNFE")

	Private lMsHelpAuto		:= .T.		// Controle de erro e log
	Private lAutoErrNoFile 	:= .T.		// Controle de erro e log - Utilizado junto com GetAutoGrLog()
	Private lMsErroAuto		:= .T.		// Controle de erro e log

	///////////////////////////////////
	// NFe
	///////////////////////////////////
	Private oNF
	Private oEmitente
	Private oIdent
	Private oDestino
	Private oTotal
	Private oTransp
	Private oDet
	Private oICM
	Private oFatura
	Private xNCMZero	:= "0|00|000|0000|00000|000000|0000000"
	///////////////////////////////////
	// NFe - FIM
	///////////////////////////////////

	CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Diretorios"  )

	cFile:= cGetFile( "Arquivo NFe (*.xml) | *.xml", "Selecione diretório",0,'\conexaonfe\',.F., )

	// Rotina abaixo é para retirar o xml escolhido e deixar somente o caminho que estão os xml's
	Caminho	:= cFile
	nPos	:= 0
	Do While AT( "\", Caminho )  <> 0
		nPos 	:= AT( "\", Caminho )
		Caminho := Substr(Caminho,nPos+1, Len(Caminho))
	Enddo

	Caminho	:= substr(cfile,1,len(cfile) - len(caminho))

	// Verifica se existe o diretório
	If !ExistDir(Caminho) .or. nPos == 0
		MsgAlert("O diretório "+cFile+" nao existe!","Atencao!")
	endif

	CamProc	:= STRTRAN(Caminho, "entrada\", "backup\")
	If !ExistDir(CamProc)
		MakeDir(CamProc)
	endif

	CamErro	:= STRTRAN(Caminho, "entrada\", "erro\")
	If !ExistDir(CamErro)
		MakeDir(CamErro)
	endif

	// Fim verifica diretório
	CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Fim diretorios"  )

	PutMV("MV_PCNFE",.f.)

	CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Busca XML"  )

	AchaFile()

	// Verifica se encontrou os xmls
	IF Empty(aFiles)
		CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Não localizou xml ou não tinha xml"  )
		Return
	endif


	CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Fim Busca XML"  )

	For Nz	:= 1 to Len(aFiles)

		cFile	:= aFiles[Nz][1]

		nHdl    := fOpen(Caminho+cFile,0)

		If nHdl == -1
			If !Empty(cFile)
				// SAG MCOMlert("O arquivo de nome "+cFile+" nao pode ser aberto! Verifique os parametros.","Atencao!")
				CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - O arquivo de nome "+cFile+" nao pode ser aberto!"  )

				cMsgErro 	:= "ERRO!!! O arquivo de nome "+cFile+" nao pode ser aberto!"  + Chr(13) + Chr(10)
				cMsgErro	+= CHR(13) + CHR(10)

				COPY FILE &(Caminho + cFile) TO &(CamProc + cFile)
				FErase((Caminho + cFile))
				MemoWrit( CamProc + cFile + ".TXT", cMsgErro)
			endif

			LOOP

		endif
		nTamFile := fSeek(nHdl,0,2)
		fSeek(nHdl,0,0)
		cBuffer  := Space(nTamFile)                // Variavel para criacao da linha do registro para leitura
		nBtLidos := fRead(nHdl,@cBuffer,nTamFile)  // Leitura  do arquivo XML
		fClose(nHdl)

		cAviso := ""
		cErro  := ""
		oNfe := XmlParser(cBuffer,"_",@cAviso,@cErro)

		If Type("oNFe:_NfeProc")<> "U"
			oNF := oNFe:_NFeProc:_NFe
		endif

		oEmitente  := oNF:_InfNfe:_Emit
		oIdent     := oNF:_InfNfe:_IDE
		oDestino   := oNF:_InfNfe:_Dest
		oTotal     := oNF:_InfNfe:_Total
		oTransp    := oNF:_InfNfe:_Transp
		oDet       := oNF:_InfNfe:_Det
		oFatura    := IIf(Type("oNF:_InfNfe:_Cobr")=="U",Nil,oNF:_InfNfe:_Cobr)
		cCgc := AllTrim(IIf(Type("oDestino:_CPF")=="U",oDestino:_CNPJ:TEXT,oDestino:_CPF:TEXT))
		cCnpjEmit	:= oEmitente:_CNPJ:TEXT
		cCnpjDest	:= oDestino:_CNPJ:TEXT

		oDet := IIf(ValType(oDet)=="O",{oDet},oDet)

		cNota 	:= 	Strzero(Val(OIdent:_nNF:TEXT),TAMSX3("F1_DOC")[1])
		cSerie	:= Padr(OIdent:_serie:TEXT,3)

/* dev avc
		IF cCgc <> SM0->M0_CGC

			CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - O CNPJ destino não pertence a empresa/filial: " + AllTrim(SM0->M0_NOME))
			PutMV("MV_PCNFE",lPcNfe)

			//fMsgErro("ATENÇÃO! ERRO NA GRAVAÇÃO NA IMPORTAÇÃO DA NOTA  " + cNota + "/" + cSerie + " POIS A RAIZ DO CNPJ DESTINO NÃO PERCENTE A EMPRESA: " + AllTrim(SM0->M0_NOME))

			cMsgErro 	:= "[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - O CNPJ destino não pertence a empresa/filial: " + AllTrim(SM0->M0_NOME) + Chr(13) + Chr(10)
			cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
			cMsgErro	+= CHR(13) + CHR(10)

//			COPY FILE &(Caminho + cFile) TO &(CamErro + cFile)
//			FErase((Caminho + cFile))
			MemoWrit( CamErro + cFile + ".TXT", cMsgErro)

			LOOP

		endif
*/

		// Cadastros
		If !fNCMInc()  // Inclusão o NCM
			LOOP
		Endif
		IF !fGrpInc()
			LOOP
		EndIf
		If !fForInc()  // Inclusão ou alteração do fornecedor
			LOOP
		Endif
/*		 dev avc
If !M020inc()  // Inclusão de produto
			LOOP
		Endif
*/
/*	A chamada ficou para o cadastro de protuto x fornecedor ficou junto com M020inc
		If !M060inc()  // Produto x Fornecedor
			LOOP
		Endif
*/
// Fim Cadastros

// -- Nota Fiscal já existe na base
//	SA2->(dbSetOrder(3), dbseek(xFilial("SA2")+cCnpjEmit))
		If SF1->(dbSetOrder(1), dbseek(xFilial("SF1")+cNota+cSerie+_xFornece+_xloja))  // F1_FILIAL+F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA+F1_FORMUL+F1_TIPO
			CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Nota fiscal já existe: " + cNota + "/" + cSerie + " do fornecedor  " + SA2->A2_NOME + "  e CNPJ/CPF " + cCgc )
			PutMV("MV_PCNFE",lPcNfe)

			fMsgErro("ATENÇÃO! ERRO NA GRAVAÇÃO NA IMPORTAÇÃO DA NOTA (JÁ LANÇADA)  " + cNota + "/" + cSerie + " DO FORNECEDOR " + Alltrim(SA2->A2_NOME) + "  e CNPJ/CPF " + cCgc)

			LOOP
		endif

		If Type("oNF:_InfNfe:_ICMS")<> "U"
			oICM       := oNF:_InfNfe:_ICMS
		Else
			oICM		:= nil
		endif

		lMsErroAuto := .f.
		lMsHelpAuto := .T.

		aCabec := {}
		aItens := {}
		cChave := oNFe:_NfeProc:_PROTNFE:_INFPROT:_CHNFE:TEXT
		cNota  := OIdent:_nNF:TEXT
		aadd(aCabec,{"F1_TIPO"   ,"N" ,Nil,Nil})
		aadd(aCabec,{"F1_FORMUL" ,"S" ,Nil,Nil})  // DEV AVC
		/*
		aadd(aCabec,{"F1_DOC"    ,strzero(val(OIdent:_nNF:TEXT),9),Nil,Nil})
		aadd(aCabec,{"F1_CHVNFE"  ,oNFe:_NfeProc:_PROTNFE:_INFPROT:_CHNFE:TEXT,Nil,Nil})
		aadd(aCabec,{"F1_SERIE"  ,OIdent:_serie:TEXT,Nil,Nil})
*/
		aadd(aCabec,{"F1_DOC"    ,strzero(val(OIdent:_nNF:TEXT),9),Nil,Nil})
		aadd(aCabec,{"F1_CHVNFE"  ,oNFe:_NfeProc:_PROTNFE:_INFPROT:_CHNFE:TEXT,Nil,Nil})
		aadd(aCabec,{"F1_SERIE"  ,OIdent:_serie:TEXT,Nil,Nil})


		cData:=Alltrim(OIdent:_dhEmi:TEXT)
		_dData:=CTOD(Subs(cData,9,2)+'/'+Substr(cData,6,2)+'/'+Subs(cData,1,4))

		aadd(aCabec,{"F1_EMISSAO",_dData,Nil,Nil})
		aadd(aCabec,{"F1_FORNECE",_xFornece,Nil,Nil})
		aadd(aCabec,{"F1_LOJA"   ,_xLOJA,Nil,Nil})
		aadd(aCabec,{"F1_ESPECIE","SPED",Nil,Nil})
		aAdd(aCabec,{"F1_STATUS", "" ,NIL, NIL})
		aAdd(aCabec,{"F1_MENNOTA", "IMPORTADOR" ,NIL, NIL})



	aCab := {	{"F1_FILIAL"	,xFilial('SF1')						,NIL},;
	{"F1_TIPO"		,"N"								,NIL},;
	{"F1_FORMUL"	,"S"								,NIL},;	
	{"F1_DOC"		,"1"								,NIL},;	
	{"F1_SERIE"		,"3"								,NIL},;	
	{"F1_EMISSAO"	,dDataBase							,NIL},;
	{"F1_DTDIGIT"	,dDataBase							,NIL},;
	{"F1_LOJA"		,_xLOJA     						,NIL},;
	{"F1_FORNECE"	,_xFornece  						,NIL},;
	{"F1_FRETE"		,0		 			  				,NIL},;
	{"F1_FORMUL"	,' '    							,NIL},;
	{"F1_ESPECIE"	,'SPED'    							,NIL},;
	{"F1_EST"		,oDestino:_enderdest:_UF:TEXT 		,NIL},;
	{"F1_RECBMTO"	,dDatabase    						,NIL},;
	{"F1_DESPESA"	,0		    						,NIL},;
	{"F1_VALMERC"	,0		    						,NIL},;
	{"F1_MENNOTA"	,oNFe:_NfeProc:_PROTNFE:_INFPROT:_CHNFE:TEXT  						,NIL},;
	{"F1_COND"		,"001"	    						,NIL},;	
	{"F1_HORA"		,SUBSTR(TIME(), 1, 5)				,NIL}} 

		For nX := 1 To Len(oDet)

			cProduto:= PadR(AllTrim(oDet[nX]:_Prod:_cProd:TEXT),15)
			xProduto:=cProduto

//			cNCM:=IIF(Type("oDet[nX]:_Prod:_NCM")=="U",space(12),oDet[nX]:_Prod:_NCM:TEXT)

			SB1->(dbSetOrder(1))


//	Next nX

//	For nX := 1 To Len(oDet)

			aLinha := {}
			// Validacao: Produto Existe no SB1 ?
			// Se não existir, abrir janela c/ codigo da NF e descricao para digitacao do cod. substituicao.
			// Descricao: oDet[nX]:_Prod:_xProd:TEXT

			cDescProd	:= Upper(oDet[nX]:_Prod:_xProd:TEXT)

			//faz o calculo dos espacos a inserir no campo F1_Doc
//	cPro := len(xProduto)
//	cTam := 20 - cPro
			cProdt := oDet[nX]:_Prod:_cProd:TEXT    //  xProduto+SPACE(cTam)

			cNCM:=IIF(Type("oDet[nX]:_Prod:_NCM")=="U",space(12),oDet[nX]:_Prod:_NCM:TEXT)

			SA5->(dbSetOrder(14), dbseek(xFilial("SA5")+_xFornece+_xloja+cProdt))  //SA5->(DbSetOrder(1))		// (E) A5_FILIAL+A5_FORNECE+A5_LOJA+A5_CODPRF
			SB1->(dbSetOrder(1) , dbseek(xFilial("SB1")+SA5->A5_PRODUTO))
/*			Else
				SA5->(dbSetOrder(14), dbseek(xFilial("SA5")+SA2->A2_COD+SA2->A2_LOJA+cProdt))
				SB1->(dbSetOrder(1) , dbseek(xFilial("SB1")+SA5->A5_PRODUTO))
		endif
*/
		aadd(aLinha,{"D1_COD",SB1->B1_COD,Nil,Nil})
		aadd(aLinha,{"D1_ITEM",Strzero(nX,4),Nil,Nil})
		If Val(oDet[nX]:_Prod:_qTrib:TEXT) != 0
			aadd(aLinha,{"D1_QUANT",Val(oDet[nX]:_Prod:_qTrib:TEXT),Nil,Nil})
			aadd(aLinha,{"D1_VUNIT",Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qTrib:TEXT),6),Nil,Nil})
		Else
			aadd(aLinha,{"D1_QUANT",Val(oDet[nX]:_Prod:_qCom:TEXT),Nil,Nil})
			aadd(aLinha,{"D1_VUNIT",Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qCom:TEXT),6),Nil,Nil})
		endif
		aadd(aLinha,{"D1_TOTAL",Val(oDet[nX]:_Prod:_vProd:TEXT),Nil,Nil})
		_cfop:=oDet[nX]:_Prod:_CFOP:TEXT

		aadd(aLinha,{"D1_TES","104",Nil,Nil})


		If Type("oDet[nX]:_Prod:_vDesc")<> "U"
			aadd(aLinha,{"D1_VALDESC",Val(oDet[nX]:_Prod:_vDesc:TEXT),Nil,Nil})
		endif
		Do Case
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS00")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS00
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS10")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS10
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS20")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS20
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS30")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS30
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS40")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS40
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS51")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS51
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS60")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS60
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS70")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS70
		Case Type("oDet[nX]:_Imposto:_ICMS:_ICMS90")<> "U"
			oICM:=oDet[nX]:_Imposto:_ICMS:_ICMS90
		EndCase

//CST_Aux:=Alltrim(oICM:_orig:TEXT)+Alltrim(oICM:_CST:TEXT)
//aadd(aLinha,{"D1_CLASFIS",CST_Aux,Nil,Nil})

		If Type("oICM:_orig:TEXT")<> "U" .And. Type("oICM:_CST:TEXT")<> "U"
			IF alltrim(oICM:_CST:TEXT) == "00"
				CST_Aux := Alltrim(oICM:_orig:TEXT)+Alltrim(oICM:_CST:TEXT)
				picm_aux:= Alltrim(oICM:_picms:TEXT)
				vicm_aux:= Alltrim(oICM:_vicms:TEXT)
				aadd(aLinha,{"D1_CLASFIS",CST_Aux,Nil,Nil})
				aadd(aLinha,{"D1_PICM",val(picm_aux),Nil,Nil})
				aadd(aLinha,{"D1_VALICM",val(vicm_aux),Nil,Nil})
			endif
		endif


				aLinha := { {'D1_FILIAL ', xFilial('SD1')   			,Nil},;
		{"D1_COD"	,"27.3.0020"			     				,NIL},;
		{"D1_UM"	,"UN"										,NIL},;
		{"D1_QUANT"	,Val(oDet[nX]:_Prod:_qCom:TEXT)				,NIL},;
		{"D1_VUNIT"	,Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qCom:TEXT),6)					,NIL},;
		{"D1_TOTAL"	,Val(oDet[nX]:_Prod:_vProd:TEXT)			,NIL},;
		{"D1_ITEM" 	,StrZero(1,TamSX3("D1_ITEM")[1], 0)			,NIL},;
		{"D1_TES" 	,"104"								 		,NIL},;
		{"D1_LOCAL"	,"01"										,NIL}}

/*
		{"D1_NFORI" ,strzero(val(OIdent:_nNF:TEXT),6)	 		,NIL},;
		{"D1_SERIORI",OIdent:_serie:TEXT				 		,NIL},;
		{"D1_ITEMORI","01"								 		,NIL},;
*/

		aadd(aItens,aLinha)
	Next nX


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//| Inclusao Pré-Nota                                          |
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	If Len(aItens) > 0
		lMsErroAuto := .f.
		lMsHelpAuto := .T.

		SB1->( dbSetOrder(1) )
		SA1->( dbSetOrder(1) )

		nModulo := 2  //ESTOQUE
	//	MSExecAuto({|x,y,z,a,b| MATA140(x,y,z,a,b)}, aCabec, aItens, 3 ,,) //3 = Inclusão, 4 = Alteração, 5 = Exclusão, 7 = Estorna Classificação
	MSExecAuto({|x,y,z| MATA103(x,y,z)},aCab,aItens,3)  // DEV AVC

		IF lMsErroAuto
			aLogX	:= {}

			CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Erro na geração da nota fiscal: " + cNota + "/" + cSerie + ". Processamento interrompido!"  )
			aLogX 			:= GetAutoGRLog()
			cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DA NOTA FISCAL  "  + cNota + "/" + cSerie + Chr(13) + Chr(10)
			cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
			cMsgErro	+= CHR(13) + CHR(10)

			For x:=1 To Len(aLogX)
				cMsgErro += aLogX[x] + CHR(13) + CHR(10)
			Next x

			fMsgErro(cMsgErro)

			LOOP

		Else

			CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Sucesso!!! Nota fiscal: " + cNota + "/" + cSerie + ". Gerada com sucesso!"  )

			cMsgErro 	:= "SUCESSO! GRAVAÇÃO DA NOTA FISCAL  "  + cNota + "/" + cSerie + Chr(13) + Chr(10)

			COPY FILE &(Caminho + cFile) TO &(CamProc + cFile)
			FErase((Caminho + cFile))
			MemoWrit( CamProc + cFile + ".TXT", cMsgErro)

			LOOP
		Endif

	endif

Next Nz


PutMV("MV_PCNFE",lPcNfe)
Return

Static Function AchaFile(cCodBar)
	Local cCaminho := Caminho

//	aFiles := Directory(cCaminho+"*-nfeproc.XML", "D")  // pega os xmls que estão no diretório Caminho (escolhido pelo usuário)
aFiles := Directory(cCaminho+"*.XML", "D")  // pega os xmls que estão no diretório Caminho (escolhido pelo usuário) 

Return

///////////////////////////////////////////////////////////////////
// Mensagem de erro
///////////////////////////////////////////////////////////////////
Static Function fMsgErro(cMsg )

	cMsgErro 	:= cMsg  + Chr(13) + Chr(10)
	cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
	cMsgErro	+= CHR(13) + CHR(10)

	COPY FILE &(Caminho + cFile) TO &(CamErro + cFile)
	FErase((Caminho + cFile))
	MemoWrit( CamErro + cFile + ".TXT", cMsgErro)

Return


///////////////////////////////////////////////////////////////////
// Inclusão de Produto
///////////////////////////////////////////////////////////////////
Static Function M020inc()

	Local cCodNew	:= ""
	Local aInclSB1	:= {}
	Local Nx		:= 0
	Local X			:= 0
	Local CodProd	:= ""
	Local cQry		:= ""

	cQry	+= "SELECT MAX(B1_COD) AS PROD "
	cQry	+= "FROM "+ RetSqlName("SB1") +" SB1 "
	cQry	+= "WHERE SB1.D_E_L_E_T_ = '' "
	cQry 	:= ChangeQuery(cQry)

	TCQUERY cQry NEW ALIAS "ALIAS010" NEW

	cCodNew := soma1(ALIAS010->PROD)

	ALIAS010->(DbCloseArea())

	For Nx := 1 to Len(oDet)

		lMsErroAuto 	:= .F.

		CodProd		:= oDet[nX]:_Prod:_cProd:TEXT

		SA5->(dbSetOrder(1), dbseek(xFilial("SA5")+_xFornece+_xloja+CodProd))  //SA5->(DbSetOrder(1))		// (E) A5_FILIAL+A5_FORNECE+A5_LOJA+A5_CODPRF
		SB1->(dbSetOrder(1) , dbseek(xFilial("SB1")+SA5->A5_PRODUTO))

		//	If !SB1->(dbSetOrder(1), dbseek(xFilial("SB1")+CodProd)) .and. SA5->(!EOF())
		If !SB1->(dbSetOrder(1), dbseek(xFilial("SB1")+SA5->A5_PRODUTO)) .and. (SA5->(!EOF()) .or. Empty(SA5->A5_PRODUTO))

			aInclSB1 := { {"B1_COD" ,Subs(cCodNew,1,15)															,NIL},;
				{"B1_DESC" 		, FwCutOff(Subs(Alltrim(odet[Nx]:_prod:_xprod:text),1,50) , .T.)				,NIL},;
				{"B1_TIPO" 		,"PA" 																			,Nil},;
				{"B1_UM" 		,PADR(Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2)),2)					,Nil},;
				{"B1_GRUPO"		,"01" 																			,Nil},;
				{"B1_LOCPAD" 	,"01" 																			,Nil},;
				{"B1_PICM" 		,0 																				,Nil},;
				{"B1_IPI" 		,0 																				,Nil},;
				{"B1_POSIPI" 	,IIF(Empty(odet[Nx]:_prod:_ncm:text) .or. odet[Nx]:_prod:_ncm:text $ xNCMZero ,"00000000",odet[Nx]:_prod:_ncm:text )	,Nil},;
				{"B1_CLASFIS" 	,"00" 																			,Nil},;
				{"B1_SITTRIB" 	,"0" 																			,Nil},;
				{"B1_CONTRAT" 	,"N" 																			,Nil},;
				{"B1_GARANT" 	,"2" 																			,Nil},;
				{"B1_CODGTIN" 	,"SEM GTIN" 																	,Nil}}

			MSExecAuto({|x,y| Mata010(x,y)},aInclSB1,3)

			IF lMsErroAuto

				CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Erro no cadastro de produto - "+CodProd+". Processamento interrompido!"  )
				aLogX 			:= GetAutoGRLog()
				cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DO CADASTRO DE PRODUTO  " + cCgc + Chr(13) + Chr(10)
				cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
				cMsgErro	+= CHR(13) + CHR(10)

				For x:=1 To Len(aLogX)
					cMsgErro += aLogX[x] + CHR(13) + CHR(10)
				Next x

				fMsgErro(cMsgErro)
				return(.F.)
			endif
		endif

		M060inc(cCodNew,CodProd)
		cCodNew	:= soma1(cCodNew)

	Next Nx

Return .T.

///////////////////////////////////////////////////////////////////
// Inclusão de Fornecedor x Produto
///////////////////////////////////////////////////////////////////
Static Function M060inc(cCodNew,CodProd)

	Local cProd061 	:= ""
//	Local Nx		:= 0


	cProd061 := CodProd     //PadR(AllTrim(oDet[01]:_Prod:_cProd:TEXT),15)
	DbSelectArea("SA5")
	SA5->(DbSetOrder(14))		// (1) A5_FILIAL+A5_FORNECE+A5_LOJA+A5_PRODUTO+A5_FABR+A5_FALOJA     (14) A5_FILIAL+A5_FORNECE+A5_LOJA+A5_CODPRF


//	For Nx := 1 to Len(oDet)

	//cProd061 := PadR(AllTrim(oDet[Nx]:_Prod:_cProd:TEXT),15)

	If !SA5->(DbSeek(xFilial('SA5')+_xFornece+_xloja+cProd061))

		RecLock("SA5", .T.)

		SA5->A5_PRODUTO	:= cCodNew
		SA5->A5_FORNECE	:= _xFornece
		SA5->A5_LOJA	:= _xloja
		SA5->A5_CODPRF	:= cProd061
		SA5->A5_NOMEFOR	:= Subs(Posicione("SA2",1,xFilial("SA2") + _xFornece+_xloja,"A2_NOME"),1,40)
		SA5->A5_NOMPROD	:= Posicione("SB1",1,xFilial("SB1") + cCodNew,"B1_DESC")

		MsUnLock("SA5")

	Endif

//	Next Nx

Return .T.


/*
	Local nOpc		:= 3
	Local oModel 	:= Nil
//	Local cProd061 	:= ""
//	Local Nx		:= 0
	Local lRet		:= .F.

//	For Nx := 1 to Len(oDet)
		oModel := FWLoadModel('MATA061')

		oModel:SetOperation(nOpc)
		oModel:Activate()

		cProd061 := CodProd     //PadR(AllTrim(oDet[01]:_Prod:_cProd:TEXT),15)
		DbSelectArea("SA5")
		SA5->(DbSetOrder(14))		// (1) A5_FILIAL+A5_FORNECE+A5_LOJA+A5_PRODUTO+A5_FABR+A5_FALOJA     (14) A5_FILIAL+A5_FORNECE+A5_LOJA+A5_CODPRF    


	If !SA5->(DbSeek(xFilial('SA5')+_xFornece+_xloja+cProd061))

			//Cabeçalho
			oModel:SetValue('MdFieldSA5','A5_PRODUTO',cCodNew)
			oModel:SetValue('MdFieldSA5','A5_NOMPROD',Posicione("SB1",1,xFilial("SB1") + cCodNew,"B1_DESC") )

			//Grid
			oModel:SetValue('MdGridSA5','A5_FORNECE',_xFornece)
			oModel:SetValue('MdGridSA5','A5_LOJA' ,_xloja)
			oModel:SetValue('MdGridSA5','A5_NOMEFOR', Subs(Posicione("SA2",1,xFilial("SA2") + _xFornece+_xloja,"A2_NOME"),1,40) ) 
			oModel:SetValue('MdGridSA5','A5_CODPRF' , cProd061)

		If oModel:VldData()
				oModel:CommitData()
				lRet		:= .T.
		Else

				aErro := oModel:GetErrorMessage()

				//Monta o Texto que será mostrado na tela
				AutoGrLog("Id do formulário de origem:"  + ' [' + AllToChar(aErro[01]) + ']')
				AutoGrLog("Id do campo de origem: "      + ' [' + AllToChar(aErro[02]) + ']')
				AutoGrLog("Id do formulário de erro: "   + ' [' + AllToChar(aErro[03]) + ']')
				AutoGrLog("Id do campo de erro: "        + ' [' + AllToChar(aErro[04]) + ']')
				AutoGrLog("Id do erro: "                 + ' [' + AllToChar(aErro[05]) + ']')
				AutoGrLog("Mensagem do erro: "           + ' [' + AllToChar(aErro[06]) + ']')
				AutoGrLog("Mensagem da solução: "        + ' [' + AllToChar(aErro[07]) + ']')
				AutoGrLog("Valor atribuído: "            + ' [' + AllToChar(aErro[08]) + ']')
				AutoGrLog("Valor anterior: "             + ' [' + AllToChar(aErro[09]) + ']')

				//Mostra a mensagem de Erro
				MostraErro()
				Return .F.
		Endif

			oModel:DeActivate()

			oModel:Destroy()
	Endif



Return lRet
*/


///////////////////////////////////////////////////////////////////
// Inclusão de Fornecedor x Produto
///////////////////////////////////////////////////////////////////
Static Function fForInc()

	Local oSA2Mod 	:= Nil
	Local oModel 	:= Nil
	Local lAchou	:= .F.
	Local lDeuCerto := .F.
	Local nOpcao	:= 0

	lAchou  := SA2->(dbSetOrder(3), dbseek(xFilial("SA2")+cCnpjDest))  // se achou (.T.) então altera, caso se não achou (.F.) então inclui

	cCodFor := IIf(lAchou, SA2->A2_COD , GetSxeNum("SA2","A2_COD") )
	clojFor := IIf(lAchou, SA1->A1_LOJA, "00" )

	nOpcao	:= Iif(!lAchou, 3, 4)

//Pegando o modelo de dados, setando a operação de inclusão
	oModel := FWLoadModel("MATA020M")  // Não substituir o objeto oModel
//oSA2Mod:SetOperation(3)
	oModel:SetOperation(Iif(!lAchou, 3, 4)) // 3 - Inclui e 4 - Altera  // // Não substituir o objeto oModel
	oModel:Activate()  // Não substituir o objeto oModel

//Pegando o model dos campos da SA2
	oSA2Mod:= oModel:getModel("SA2MASTER")  // // Não substituir o objeto oModel
	if nOpcao == 3
		oSA2Mod:setValue("A2_COD",       cCodFor   													) // Codigo
		oSA2Mod:setValue("A2_LOJA",      cLojFor       												) // Loja
	Endif
	oSA2Mod:setValue("A2_NOME",      FwCutOff(Subs(odestino:_xNome:TEXT,1,50) , .T.)				) // Nome
	oSA2Mod:setValue("A2_NREDUZ",   IIF( Type("odestino:_xFANT:TEXT") != "U" ,    FwCutOff(Subs(odestino:_xFANT:TEXT,1,50) , .T.)	,  FwCutOff(Subs(odestino:_xNome:TEXT,1,50) , .T.)	) ) // Nome reduz.
	oSA2Mod:setValue("A2_END",       FwCutOff(Subs(odestino:_enderdest:_xLgr:TEXT+" "+odestino:_enderdest:_nro:TEXT,1,40) , .T.)   ) // Endereco
	oSA2Mod:setValue("A2_NR_END",    odestino:_enderdest:_nro:TEXT   								) // Nro Endereco
	oSA2Mod:setValue("A2_BAIRRO",    FwCutOff(Subs(odestino:_enderdest:_xBairro:TEXT,1,20) , .T.)  ) // Bairro
	oSA2Mod:setValue("A2_TIPO",      IF(Len(cCnpjDest) == 14,"J", "F" )    							) // Tipo
	oSA2Mod:setValue("A2_EST",       odestino:_enderdest:_UF:TEXT        							) // Estado
	oSA2Mod:setValue("A2_COD_MUN",   Subs(odestino:_enderdest:_cMun:TEXT,3,5)     					) // Codigo Municipio
	oSA2Mod:setValue("A2_MUN",       odestino:_enderdest:_xMun:TEXT    							) // Municipio
	oSA2Mod:setValue("A2_CEP",       odestino:_enderdest:_CEP:TEXT        							) // CEP
	IF Type("odestino:_ie:TEXT") != "U"
		oSA2Mod:setValue("A2_INSCR",     odestino:_ie:TEXT         								) // Inscricao Estadual
	else
		oSA2Mod:setValue("A2_INSCR",     "ISENTO"        											) // Inscricao Estadual
	endif
	oSA2Mod:setValue("A2_CGC",       cCnpjDest      												) // CNPJ/CPF
	oSA2Mod:setValue("A2_PAIS",      "105"    ) // Pais
	IF Type("odestino:_enderdest:_fone:TEXT") != "U"
		oSA2Mod:setValue("A2_DDD",      Subs(odestino:_enderdest:_fone:TEXT,1,2)   				) // DDD
		oSA2Mod:setValue("A2_TEL",      Subs(odestino:_enderdest:_fone:TEXT,3,10)   				) // Fone
	ELSE // Incluído
		oSA2Mod:setValue("A2_DDD",    	"999"        												) // DDD
		oSA2Mod:setValue("A2_TEL",      "9999-9999"   												) // Fone
	endif
	oSA2Mod:setValue("A2_CONTA",       	"211011101"   												) // Conta Contábil
//	oSA2Mod:setValue("A2_NATUREZ",     	"201091"   													) // Natureza  DEV AVC
	oSA2Mod:setValue("A2_NATUREZ",     	"112015"   													) // Natureza
	oSA2Mod:setValue("A2_RECISS",     	"N"   														) // Recolhe ISS

//	oSA2Mod:setValue("A2_TPESSOA",   cTipPessoa  ) // Tipo Pessoa
	oSA2Mod:setValue("A2_CODPAIS",   "01058"   														) // Pais Bacen


/* dev avc
	lAchou  := SA2->(dbSetOrder(3), dbseek(xFilial("SA2")+cCnpjEmit))  // se achou (.T.) então altera, caso se não achou (.F.) então inclui

	cCodFor := IIf(lAchou, SA2->A2_COD , GetSxeNum("SA2","A2_COD") )
	clojFor := IIf(lAchou, SA1->A1_LOJA, "00" )

	nOpcao	:= Iif(!lAchou, 3, 4)

//Pegando o modelo de dados, setando a operação de inclusão
	oModel := FWLoadModel("MATA020M")  // Não substituir o objeto oModel
//oSA2Mod:SetOperation(3)
	oModel:SetOperation(Iif(!lAchou, 3, 4)) // 3 - Inclui e 4 - Altera  // // Não substituir o objeto oModel
	oModel:Activate()  // Não substituir o objeto oModel

//Pegando o model dos campos da SA2
	oSA2Mod:= oModel:getModel("SA2MASTER")  // // Não substituir o objeto oModel
	if nOpcao == 3
		oSA2Mod:setValue("A2_COD",       cCodFor   													) // Codigo
		oSA2Mod:setValue("A2_LOJA",      cLojFor       												) // Loja
	Endif
	oSA2Mod:setValue("A2_NOME",      FwCutOff(Subs(oEmitente:_xNome:TEXT,1,50) , .T.)				) // Nome
	oSA2Mod:setValue("A2_NREDUZ",   IIF( Type("oEmitente:_xFANT:TEXT") != "U" ,    FwCutOff(Subs(oEmitente:_xFANT:TEXT,1,50) , .T.)	,  FwCutOff(Subs(oEmitente:_xNome:TEXT,1,50) , .T.)	) ) // Nome reduz.
	oSA2Mod:setValue("A2_END",       FwCutOff(Subs(oEmitente:_enderemit:_xLgr:TEXT+" "+oEmitente:_enderemit:_nro:TEXT,1,40) , .T.)   ) // Endereco
	oSA2Mod:setValue("A2_NR_END",    oEmitente:_enderemit:_nro:TEXT   								) // Nro Endereco
	oSA2Mod:setValue("A2_BAIRRO",    FwCutOff(Subs(oEmitente:_enderemit:_xBairro:TEXT,1,20) , .T.)  ) // Bairro
	oSA2Mod:setValue("A2_TIPO",      IF(Len(cCnpjEmit) == 14,"J", "F" )    							) // Tipo
	oSA2Mod:setValue("A2_EST",       oEmitente:_enderemit:_UF:TEXT        							) // Estado
	oSA2Mod:setValue("A2_COD_MUN",   Subs(oEmitente:_enderemit:_cMun:TEXT,3,5)     					) // Codigo Municipio
	oSA2Mod:setValue("A2_MUN",       oEmitente:_enderemit:_xMun:TEXT    							) // Municipio
	oSA2Mod:setValue("A2_CEP",       oEmitente:_enderemit:_CEP:TEXT        							) // CEP
	IF Type("oEmitente:_ie:TEXT") != "U"
		oSA2Mod:setValue("A2_INSCR",     oEmitente:_ie:TEXT         								) // Inscricao Estadual
	else
		oSA2Mod:setValue("A2_INSCR",     "ISENTO"        											) // Inscricao Estadual
	endif
	oSA2Mod:setValue("A2_CGC",       cCnpjEmit      												) // CNPJ/CPF
	oSA2Mod:setValue("A2_PAIS",      "105"    ) // Pais
	IF Type("oEmitente:_enderdest:_fone:TEXT") != "U"
		oSA2Mod:setValue("A2_DDD",      Subs(oEmitente:_enderemit:_fone:TEXT,1,2)   				) // DDD
		oSA2Mod:setValue("A2_TEL",      Subs(oEmitente:_enderemit:_fone:TEXT,3,10)   				) // Fone
	ELSE // Incluído
		oSA2Mod:setValue("A2_DDD",    	"999"        												) // DDD
		oSA2Mod:setValue("A2_TEL",      "9999-9999"   												) // Fone
	endif
	oSA2Mod:setValue("A2_CONTA",       	"211011101"   												) // Conta Contábil
//	oSA2Mod:setValue("A2_NATUREZ",     	"201091"   													) // Natureza  DEV AVC
	oSA2Mod:setValue("A2_NATUREZ",     	"112015"   													) // Natureza
	oSA2Mod:setValue("A2_RECISS",     	"N"   														) // Recolhe ISS

//	oSA2Mod:setValue("A2_TPESSOA",   cTipPessoa  ) // Tipo Pessoa
	oSA2Mod:setValue("A2_CODPAIS",   "01058"   														) // Pais Bacen
*/

//Se conseguir validar as informações
	If oModel:VldData()  // Não substituir o objeto oModel

		//Tenta realizar o Commit
		If oModel:CommitData()  // Não substituir o objeto oModel
			lDeuCerto := .T.

			_xFornece := SA2->A2_COD
			_xloja    := SA2->A2_LOJA

			If !lAchou
				ConfirmSx8() // Confirma a gravaçãp do número sxf/sxf
			Endif


			//Se não deu certo, altera a variável para false
		Else
			lDeuCerto := .F.

			If !lAchou
				RollBackSX8()
			Endif

			CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Erro no cadastro do fornecedor - "+cCgc+". Processamento interrompido!"  )
			cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DO CADASTRO DO FORNECEDOR  " + cCgc + Chr(13) + Chr(10)
			cMsgErro	+= CHR(13) + CHR(10)

			fMsgErro(cMsgErro)

/*
			COPY FILE &(Caminho + cFile) TO &(CamErro + cFile)
			FErase((Caminho + cFile))
			MemoWrit( CamErro + cFile + ".TXT", cMsgErro)
*/
		EndIf

//Se não conseguir validar as informações, altera a variável para false
	Else
		lDeuCerto := .F.

	EndIf

//Se não deu certo a inclusão, mostra a mensagem de erro
	If ! lDeuCerto
		//Busca o Erro do Modelo de Dados
		aErro := oModel:GetErrorMessage()   // Não substituir o objeto oModel

/*
		//Monta o Texto que será mostrado na tela
		AutoGrLog("Id do formulário de origem:"  + ' [' + AllToChar(aErro[01]) + ']')
		AutoGrLog("Id do campo de origem: "      + ' [' + AllToChar(aErro[02]) + ']')
		AutoGrLog("Id do formulário de erro: "   + ' [' + AllToChar(aErro[03]) + ']')
		AutoGrLog("Id do campo de erro: "        + ' [' + AllToChar(aErro[04]) + ']')
		AutoGrLog("Id do erro: "                 + ' [' + AllToChar(aErro[05]) + ']')
		AutoGrLog("Mensagem do erro: "           + ' [' + AllToChar(aErro[06]) + ']')
		AutoGrLog("Mensagem da solução: "        + ' [' + AllToChar(aErro[07]) + ']')
		AutoGrLog("Valor atribuído: "            + ' [' + AllToChar(aErro[08]) + ']')
		AutoGrLog("Valor anterior: "             + ' [' + AllToChar(aErro[09]) + ']')
*/

		CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Erro no cadastro do fornecedor - "+cCgc+". Processamento interrompido!"  )
		cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DO CADASTRO DO FORNECEDOR  " + cCgc + Chr(13) + Chr(10)
		cMsgErro	+= CHR(13) + CHR(10)

		cMsgErro 	+= "Id do formulário de origem:"  + ' [' + AllToChar(aErro[01]) + ']' + CHR(13) + CHR(10)
		cMsgErro 	+= "Id do campo de origem: "      + ' [' + AllToChar(aErro[02]) + ']' + CHR(13) + CHR(10)
		cMsgErro 	+= "Id do formulário de erro: "   + ' [' + AllToChar(aErro[03]) + ']' + CHR(13) + CHR(10)
		cMsgErro 	+= "Id do campo de erro: "        + ' [' + AllToChar(aErro[04]) + ']' + CHR(13) + CHR(10)
		cMsgErro 	+= "Id do erro: "                 + ' [' + AllToChar(aErro[05]) + ']' + CHR(13) + CHR(10)
		cMsgErro 	+= "Mensagem do erro: "           + ' [' + AllToChar(aErro[06]) + ']' + CHR(13) + CHR(10)
		cMsgErro 	+= "Mensagem da solução: "        + ' [' + AllToChar(aErro[07]) + ']' + CHR(13) + CHR(10)
		cMsgErro 	+= "Valor atribuído: "            + ' [' + AllToChar(aErro[08]) + ']' + CHR(13) + CHR(10)
		cMsgErro 	+= "Valor anterior: "             + ' [' + AllToChar(aErro[09]) + ']' + CHR(13) + CHR(10)

		cMsgErro	+= CHR(13) + CHR(10)

		fMsgErro(cMsgErro)

		If !lAchou
			RollBackSX8()
		Endif

		//Mostra a mensagem de Erro
		//	MostraErro()
	EndIf

//Desativa o modelo de dados
	oModel:DeActivate()   // Não substituir o objeto oModel
	oModel:Destroy()   // Não substituir o objeto oModel

Return lDeuCerto


///////////////////////////////////////////////////////////////////
// Inclusão da unidade de medida
///////////////////////////////////////////////////////////////////
Static Function fGrpInc()


	Local aVetor 	:= {}
	Local Nx, x
	Local lRet		:= .T.

	For Nx := 1 to Len(oDet)

		IF !SAH->(dbSetOrder(1), dbseek(xFilial("SAH")+Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2))))
			lMsErroAuto 	:= .F.

			//--- Inclusao --- //
			aVetor:= {{"AH_UNIMED", Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2)) ,NIL},; // Un. Medida
			{"AH_UMRES" , "IMPORTADO"  ,NIL},; // Desc. Resum.
			{"AH_DESCPO", "IMPORTADOR - " + Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2)) ,Nil}} // Descr. Portug.

			MSExecAuto({|x,y| QIEA030(x,y)}, aVetor, 3)

			IF lMsErroAuto

				CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPCOM] - Erro no cadastro de grupo - "+Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2))+". Processamento interrompido!"  )
				aLogX 			:= GetAutoGRLog()
				cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DO CADASTRO DE GRUPO  " + cCgc + Chr(13) + Chr(10)
				cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
				cMsgErro	+= CHR(13) + CHR(10)

				For x:=1 To Len(aLogX)
					cMsgErro += aLogX[x] + CHR(13) + CHR(10)
				Next x

				fMsgErro(cMsgErro)
				return(.F.)
			Endif
		Endif

	Next Nx


/*
	Local oModel 	:= Nil
	Local Nx
	Local lRet		:= .T.

	For Nx := 1 to Len(oDet)

		IF !SAH->(dbSetOrder(1), dbseek(xFilial("SAH")+Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2))))
			oModel 	:= Nil
			oModel  := FwLoadModel("QIEA030M")
			
			oModel:SetOperation(3)
			oModel:Activate()
			oModel:SetValue("SAHMASTER","AH_UNIMED", Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2))) // Un. Medida
			oModel:SetValue("SAHMASTER","AH_UMRES" , "Importador - " + Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2))) // Desc. Resum.
			oModel:SetValue("SAHMASTER","AH_DESCPO", "Importador - " + Upper(Subs(Alltrim(odet[Nx]:_prod:_ucom:text),1,2))) // Descr. Portug.

			If oModel:VldData()
				oModel:CommitData()
			Else
				VarInfo("",oModel:GetErrorMessage())
				lRet		:= .F.
			EndIf

			oModel:DeActivate()
			oModel:Destroy()

			oModel := NIL
		EndIf

	Next Nx
*/

Return lRet


///////////////////////////////////////////////////////////////////
// Inclusão do NM
///////////////////////////////////////////////////////////////////
Static Function fNCMInc()
	Local Nx
	Local xNCM

	For Nx := 1 to Len(oDet)

		xNCM	:= IF(Empty(odet[Nx]:_prod:_ncm:text) .or. odet[Nx]:_prod:_ncm:text $ xNCMZero ,"00000000",odet[Nx]:_prod:_ncm:text)

		IF !SYD->(dbSetOrder(1), dbseek(xFilial("SYD")+xNCM))
			Reclock("SYD",.T.)
			SYD->YD_TEC		:= xNCM
			SYD->YD_DESC_P	:= "IMPORTADOR"
			SYD->(Msunlock())
		Endif

	Next Nx

Return .T.
