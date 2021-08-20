#INCLUDE "PROTHEUS.CH"
#include "Topconn.ch"
#Include "Totvs.ch"

User Function GRIMPSGA()

	//programa para importar notas da empresa 31


	Private cMsgErro
	Private _xcliente 	:= ""
	Private _xloja    	:= ""
	Private cDescProd	:= ""
	Private cFile 		:= Space(10)
	Private CPERG   	:="NOTAXML"
	Private Caminho 	:= "" 
	Private CamErro		:= ""
	Private CamProc		:= ""
	Private _cMarca   	:= GetMark()
	Private aFields   	:= {}
	Private cArq
	Private aFields2  	:= {}
	Private cArq2
	Private	aFiles		:= {}
	Private lPcNfe		:= GETMV("MV_PCNFE")

	Private lMsHelpAuto		:= .T.		// Controle de erro e log
	Private lAutoErrNoFile 	:= .T.		// Controle de erro e log - Utilizado junto com GetAutoGrLog()
	Private lMsErroAuto		:= .T.		// Controle de erro e log

	CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Diretorios"  )

	// Verifica e se necessário crio o diretório
	If !ExistDir("\LOGXMLFER\")
		MakeDir("\LOGXMLFER\")
	endif

	If !ExistDir("\LOGXMLFER\ERRO\")
		MakeDir("\LOGXMLFER\ERRO\")
	endif

	If !ExistDir("\LOGXMLFER\PROCESSADAS\")
		MakeDir("\LOGXMLFER\PROCESSADAS\")
	endif

	Caminho	:= "\LOGXMLFER\"
	CamErro	:= "\LOGXMLFER\ERRO\"
	CamProc	:= "\LOGXMLFER\PROCESSADAS\"
	// Fim verifica diretório
	CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Fim diretorios"  )

	PutMV("MV_PCNFE",.f.)

	CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Busca XML"  )

	AchaFile()

	Processa({|lEnd| GRFISXML()} , "Aguarde, processando informacoes...")

Return

Static Function GRFISXML()

	Local x , Nz, Nx, Ndp	:= 0

	ProcRegua(Len(aFiles))

	// Verifica se encontrou os xmls
	IF Empty(aFiles)
		CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Não localizou xml ou não tinha xml"  )

		Return
	endif

//	MV_PAR01 := nTipo
	CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Fim Busca XML"  )

	For Nz	:= 1 to Len(aFiles)

		cFile	:= aFiles[Nz][1]

		IncProc(aFiles[Nz][1])

		Private nHdl    := fOpen(Caminho+cFile,0)

		If nHdl == -1
			If !Empty(cFile)
				// SAG MsgAlert("O arquivo de nome "+cFile+" nao pode ser aberto! Verifique os parametros.","Atencao!")
				CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - O arquivo de nome "+cFile+" nao pode ser aberto!"  )
			endif

			fMsgErro("ATENÇÃO! ERRO O ARQUIVO NÃO PODE SER ABERTO " +cFile+ "!")

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
		Private oNF
		lMsErroAuto := .f.
		lMsHelpAuto := .T.

		If Type("oNFe:_NfeProc")<> "U"
			oNF := oNFe:_NFeProc:_NFe
			// SGA - Verificar quando for fazer os testes sobre xlm´s de cancelamento????????
		endif

		Private oEmitente  := oNF:_InfNfe:_Emit
		Private oIdent     := oNF:_InfNfe:_IDE
		Private oDestino   := oNF:_InfNfe:_Dest
		Private oTotal     := oNF:_InfNfe:_Total
		Private oTransp    := oNF:_InfNfe:_Transp
		Private oDet       := oNF:_InfNfe:_Det
		Private oFatura    := IIf(Type("oNF:_InfNfe:_Cobr")=="U",Nil,oNF:_InfNfe:_Cobr)
		Private cEdit1	   := Space(15)
		Private _DESCdigit :=space(55)
		Private _NCMdigit  :=space(8)

		cCnpjEmit	:= oEmitente:_CNPJ:TEXT

		oDet := IIf(ValType(oDet)=="O",{oDet},oDet)

		//faz o calculo dos espacos a inserir no campo F2_Doc
		cDoc := len(SF2->F2_DOC)
		cCn  := len(OIdent:_nNF:TEXT)
		cSpa := cDoc - cCn

		cNota 	:= 	Strzero(Val(OIdent:_nNF:TEXT),TAMSX3("F2_DOC")[1])		// SGA OIdent:_nNF:TEXT+SPACE(cSpa)
		cSerie	:= Padr(OIdent:_serie:TEXT,3)

		cCgc := AllTrim(IIf(Type("oDestino:_CPF")=="U",oDestino:_CNPJ:TEXT,oDestino:_CPF:TEXT))

		SA1->(dbSetOrder(3), dbseek(xFilial("SA1")+cCgc))

		IF Substr(cCnpjEmit,1,8) <> Substr(SM0->M0_CGC,1,8)

			CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Raiz do CNPJ destino não pertence a empresa: " + AllTrim(SM0->M0_NOME))
			PutMV("MV_PCNFE",lPcNfe)

			fMsgErro("ATENÇÃO! ERRO NA GRAVAÇÃO NA IMPORTAÇÃO DA NOTA  " + cNota + "/" + cSerie + " POIS A RAIZ DO CNPJ DESTINO NÃO PERCENTE A EMPRESA: " + AllTrim(SM0->M0_NOME))

			LOOP

		endif

		// -- Nota Fiscal já existe na base ?
		If SF2->(dbSetOrder(1), dbseek(xFilial("SF2")+cNota+cSerie+SA1->A1_COD+SA1->A1_LOJA))  // F2_FILIAL+F2_DOC+F2_SERIE+F2_CLIENTE+F2_LOJA+F2_FORMUL+F2_TIPO
			CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Nota fiscal já existe: " + cNota + "/" + cSerie + " do cliente  " + SA1->A1_NOME + "  e CNPJ/CPF " + cCgc )
			PutMV("MV_PCNFE",lPcNfe)

			fMsgErro("ATENÇÃO! ERRO NA GRAVAÇÃO NA IMPORTAÇÃO DA NOTA (JÁ LANÇADA)  " + cNota + "/" + cSerie + " DO CLIENTE " + Alltrim(SA1->A1_NOME) + "  e CNPJ/CPF " + cCgc)

			LOOP
		endif

		Private oICM
		If Type("oNF:_InfNfe:_ICMS")<> "U"
			oICM       := oNF:_InfNfe:_ICMS
		Else
			oICM		:= nil
		endif

		If !MCLIinc()  // Inclusão ou alteração do cliente
			LOOP
		Endif

		lMsErroAuto := .f.
		lMsHelpAuto := .T.

		aCabec := {}
		aItens := {}
		cChave := oNFe:_NfeProc:_PROTNFE:_INFPROT:_CHNFE:TEXT
		cNota  := OIdent:_nNF:TEXT
		aadd(aCabec,{"F2_TIPO"   ,"N" ,Nil,Nil}) //			If(MV_PAR01==1,"N",If(MV_PAR01==2,'B','D')),Nil,Nil})
		aadd(aCabec,{"F2_FORMUL" ,"N" ,Nil,Nil})
		aadd(aCabec,{"F2_DOC"    ,strzero(val(OIdent:_nNF:TEXT),9),Nil,Nil})
		aadd(aCabec,{"F2_CHVNFE"  ,oNFe:_NfeProc:_PROTNFE:_INFPROT:_CHNFE:TEXT,Nil,Nil})
		aadd(aCabec,{"F2_SERIE"  ,OIdent:_serie:TEXT,Nil,Nil})

		cData:=Alltrim(OIdent:_dhEmi:TEXT)
		_dData:=CTOD(Subs(cData,9,2)+'/'+Substr(cData,6,2)+'/'+Subs(cData,1,4))

		aadd(aCabec,{"F2_EMISSAO",_dData,Nil,Nil})
		aadd(aCabec,{"F2_CLIENTE",_xcliente,Nil,Nil})
		aadd(aCabec,{"F2_LOJA"   ,_xLOJA,Nil,Nil})
		aadd(aCabec,{"F2_ESPECIE","SPED",Nil,Nil})
		aadd(aCabec,{"F2_COND","001"})
		aadd(aCabec,{"F2_DESCONT",0})
		aadd(aCabec,{"F2_FRETE",0})
		aadd(aCabec,{"F2_SEGURO",0})
		aadd(aCabec,{"F2_DESPESA",0})

		// Primeiro Processamento
		// Busca de Informações para Pedidos de Compras

		cProds := ''
		aPedIte:={}

		For nX := 1 To Len(oDet)
			cEdit1 := Space(15)
			_DESCdigit :=space(55)
			_NCMdigit  :=space(8)

			cProduto:= "20.1.0113"   //PadR(AllTrim(oDet[nX]:_Prod:_cProd:TEXT),15)
			xProduto:=cProduto

			cNCM:=IIF(Type("oDet[nX]:_Prod:_NCM")=="U",space(12),oDet[nX]:_Prod:_NCM:TEXT)
			Chkproc=.F.

			//faz o calculo dos espacos a inserir no campo F1_Doc
			cPro := len(xProduto)
			cTam := 20 - cPro // 20 eh o tamanho do campo sa5_codprf
			cProdt :=xProduto+SPACE(cTam)

			//	If MV_PAR01 = 1

			SB1->(dbSetOrder(1))

			cProds += ALLTRIM(SB1->B1_COD)+'/'

			AAdd(aPedIte,{SB1->B1_COD,Val(oDet[nX]:_Prod:_qTrib:TEXT),Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qCom:TEXT),6),Val(oDet[nX]:_Prod:_vProd:TEXT)})

//	Next nX


			PUBLIC ticm_aux := 0
//	For nX := 1 To Len(oDet)

			aLinha := {}
			// Validacao: Produto Existe no SB1 ?
			// Se não existir, abrir janela c/ codigo da NF e descricao para digitacao do cod. substituicao.
			// Descricao: oDet[nX]:_Prod:_xProd:TEXT

			cDescProd	:= Upper(oDet[nX]:_Prod:_xProd:TEXT)

			//faz o calculo dos espacos a inserir no campo F1_Doc
			cPro := len(xProduto)
			cTam := 20 - cPro
			cProdt :=xProduto+SPACE(cTam)

			cNCM:=IIF(Type("oDet[nX]:_Prod:_NCM")=="U",space(12),oDet[nX]:_Prod:_NCM:TEXT)
			Chkproc=.F.


			SA7->(Dbsetorder(3))
			If	!SA7->(dbseek(xFilial("SA7")+SA1->A1_COD+SA1->A1_LOJA+cProduto))
				// SGA
				If !SB1->(dbSetOrder(1), dbseek(xFilial("SB1")+xProduto))
					lRetInc010 := M010inc()  // Inclusão de produto
					IF !lRetInc010
						//Return
						LOOP
					endif
				endif
				// SGA
				If !SA7->(dbSetOrder(1), dbseek(xFilial("SA7")+SA1->A1_COD+SA1->A1_LOJA+SB1->B1_COD))
					M370inc(nX)		// Inclusão de cliente x produto
				endif
			endif

			SB1->(dbSetOrder(1) , dbseek(xFilial("SB1")+SA7->A7_PRODUTO))
 	
	 		aadd(aLinha,{"D2_COD",SB1->B1_COD,Nil,Nil})
			aadd(aLinha,{"D2_ITEM","0001",Nil,Nil})
			If Val(oDet[nX]:_Prod:_qTrib:TEXT) != 0
				aadd(aLinha,{"D2_QUANT",Val(oDet[nX]:_Prod:_qTrib:TEXT),Nil,Nil})
				aadd(aLinha,{"D2_PRCVEN",Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qTrib:TEXT),6),Nil,Nil})
			Else
				aadd(aLinha,{"D2_QUANT",Val(oDet[nX]:_Prod:_qCom:TEXT),Nil,Nil})
				aadd(aLinha,{"D2_PRCVEN",Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qCom:TEXT),6),Nil,Nil})
			endif
			aadd(aLinha,{"D2_TOTAL",Val(oDet[nX]:_Prod:_vProd:TEXT),Nil,Nil})
			_cfop:=oDet[nX]:_Prod:_CFOP:TEXT


			IF Alltrim(_cfop) == "5101" .or. Alltrim(_cfop) == "6101"
				aadd(aLinha,{"D2_TES","602",Nil,Nil})  // Estava 539
			ELSEIF Alltrim(_cfop) == "5910" .or. Alltrim(_cfop) == "6910"
				aadd(aLinha,{"D2_TES","555",Nil,Nil})
			ELSEIF Alltrim(_cfop) == "5922" .or. Alltrim(_cfop) =="6922"
				aadd(aLinha,{"D2_TES","606",Nil,Nil})
			ELSEIF Alltrim(_cfop) == "5116" .or. Alltrim(_cfop) == "6116"
				aadd(aLinha,{"D2_TES","607",Nil,Nil})
//		ELSEIF Alltrim(_cfop) == "5933" .or. Alltrim(_cfop) == "6933"
//			aadd(aLinha,{"D2_TES","521",Nil,Nil})
			ELSE
				ALERT("CFOP SEM REGRA PARA TES, ACIONE SUPORTE !!!")
				Return(.t.)
			endif

			If Type("oDet[nX]:_Prod:_vDesc")<> "U"
				aadd(aLinha,{"D2_DESCONT",Val(oDet[nX]:_Prod:_vDesc:TEXT),Nil,Nil})
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
					aadd(aLinha,{"D2_CLASFIS",CST_Aux,Nil,Nil})
					aadd(aLinha,{"D2_PICM",val(picm_aux),Nil,Nil})
					aadd(aLinha,{"D2_VALICM",val(vicm_aux),Nil,Nil})
					ticm_aux += val(vicm_aux)
				endif
			endif


			aadd(aItens,aLinha)
		Next nX


		//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
		//| Teste de Inclusao                                            |
		//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

		Public _mvDATPGTO := GetMv("MV_DATPGTO")
		Public _venc1par  := GetMv("MV_DATPGTO")+30

		If Len(aItens) > 0
			lMsErroAuto := .f.
			lMsHelpAuto := .T.

			SB1->( dbSetOrder(1) )
			SA1->( dbSetOrder(1) )

			nModulo := 9  //ESTOQUE
			MSExecAuto({|x,y,z|MATA920(x,y,z)},aCabec,aItens,3)

			IF lMsErroAuto
				aLogX	:= {}

				CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Erro na geração da nota fiscal: " + cNota + "/" + cSerie + ". Processamento interrompido!"  )
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

				////////////////////////
				// Contabilização
				///////////////////////
				cSerieAju	:= 	OIdent:_serie:TEXT +  space(3 - len(alltrim(OIdent:_serie:TEXT)))  // conter os espaços corretos para localizar a nota

				if 	SD2->(dbSetOrder(3),dbseek( xFilial("SD2") + strzero(val(OIdent:_nNF:TEXT),9) + cSerieAju + _xcliente + _xLOJA , .f. ))  // (3) D2_FILIAL+D2_DOC+D2_SERIE+D2_CLIENTE+D2_LOJA+D2_COD+D2_ITEM

					// Rotina IMPSGA limpa o conteúdo do campo D2_ORIGLAN (LF) para utilizar a contabilização normal do faturamento
					//No chamado 11687030 informaram que só contabilizava LP 6B7 somente o campo D2_TOTAL sem IF e outras coisas (e ainda estava com erro e Totvs estava corrigindo)
					RecLock("SD2",.f.)
					SD2->D2_ORIGLAN	:=	""
					MsUnlock("SD2")

					if 	SF2->(dbSetOrder(1),dbseek( xFilial("SF2") + strzero(val(OIdent:_nNF:TEXT),9) + cSerieAju + _xcliente + _xLOJA , .f. ))

						// No MSAUTOEXEC não estava gravando esses campos
						RecLock("SF2",.f.)
						SF2->F2_MENNOTA	:=	"IMPORTADA SGA"
						SF2->F2_COND	:=	"001"
						MsUnlock("SF2")

					EndIf
				EndIf

				////////////////////////
				// Contabilização - FIM
				////////////////////////

				CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Nota fiscal gerada com sucesso: " + cNota + "/" + cSerie + "!"  )

				If Type("oFatura:_dup:_vDup:TEXT") <> "U"

					nPosFin := Ascan(aLinha, { |X| Alltrim(X[1]) == "D2_TES" })

					If nPosFin > 0 .and. SF4->(dbSetOrder(1), dbseek(xFilial("SF4")+aLinha[nPosFin,2])) //.and. val(oFatura:_fat:_vLiq:TEXT) > 0

						If SF4->F4_DUPLIC == "S"

							cTipoXML := valtype(oFatura:_dup)

							//Gera contas a receber - inicio
							aCabSE1     := {}

							If cTipoXML == "A"
								For Ndp := 1 to Len(oFatura:_dup)

									dDVenc		:= StrTran(oFatura:_dup[ndp]:_dVenc:TEXT,'-','')
									cParcela	:= oFatura:_dup[ndp]:_nDup:TEXT   // Iif(Type("oFatura:_dup:_nDup:TEXT") <> "U" , oFatura:_dup:_nDup:TEXT , "" )
									dDVenc		:= CTOD(Subs(dDVenc,7,2)+'/'+Substr(dDVenc,5,2)+'/'+Subs(dDVenc,1,4))


									// desabilita contabilização on-line
									SX1->(DbSeek("FIN040    03"))
									SX1->(Reclock("SX1", .F.))
									SX1->X1_CNT01 := "2"
									SX1->X1_PRESEL := 2
									SX1->(MSUNLOCK())
									//MV_PAR03 := 2

									lMsErroAuto := .f.
									lMsHelpAuto := .T.

									nModulo := 6  //ESTOQUE
									aadd(aCabSE1,{"E1_PREFIXO"    ,OIdent:_serie:TEXT,Nil,Nil})
									aadd(aCabSE1,{"E1_NUM"		  ,strzero(val(OIdent:_nNF:TEXT),9),Nil,Nil})
									aadd(aCabSE1,{"E1_TIPO"		  ,"NF",Nil,Nil})
									aadd(aCabSE1,{"E1_PARCELA"	  ,cParcela,Nil,Nil})
									aadd(aCabSE1,{"E1_CLIENTE"	  ,_xcliente,Nil,Nil})
									aadd(aCabSE1,{"E1_LOJA"		  ,_xloja,Nil,Nil})
									aadd(aCabSE1,{"E1_NATUREZ"	  ,"611011",Nil,Nil})
									aadd(aCabSE1,{"E1_EMISSAO"	  ,_dData,Nil,Nil})
									aadd(aCabSE1,{"E1_VENCTO"	  ,Iif(dDVenc >= dDatabase , dDVenc ,dDatabase),Nil,Nil})
									aadd(aCabSE1,{"E1_VALOR"	  ,val(oFatura:_dup[ndp]:_vDup:TEXT),Nil,Nil})
									aadd(aCabSE1,{"E1_MOEDA"	  ,1,Nil,Nil})
									aadd(aCabSE1,{"E1_ORIGEM"	  ,"MATA920",Nil,Nil})

									MSExecAuto({|x,y| FINA040(x,y)},aCabSE1,3) //Inclusao

									// habilita contabilização on-line
									SX1->(DbSeek("FIN040    03"))
									SX1->(Reclock("SX1", .F.))
									SX1->X1_CNT01 := "1"
									SX1->X1_PRESEL := 1
									SX1->(MSUNLOCK())



									// Melhorar e colocar esse o else em uma única rotina
									IF lMsErroAuto
										aLogX	:= {}

										CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Erro na geração do contas a receber: " + cNota + "/" + cSerie + ". Processamento interrompido!"  )
										aLogX 			:= GetAutoGRLog()
										cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DO CONTAS A RECEBER  "  + cNota + "/" + cSerie + Chr(13) + Chr(10)
										cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
										cMsgErro	+= CHR(13) + CHR(10)

										For x:=1 To Len(aLogX)
											cMsgErro += aLogX[x] + CHR(13) + CHR(10)
										Next x

										fMsgErro(cMsgErro)

										//	LOOP
									Elseif Ndp == Len(oFatura:_dup)

										aLogX	:= {}
										CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Titulo a receber gerado com sucesso: " + cNota + "/" + cSerie + "!"  )
										aLogX 			:= GetAutoGRLog()

										cMsgErro 	:= "SUCESSO!!! NA GERAÇÃO DO TÍTULO A RECEBER  "  + cNota + "/" + cSerie + Chr(13) + Chr(10)
										cMsgErro	+= CHR(13) + CHR(10)

										COPY FILE &(Caminho + cFile) TO &(CamProc + cFile)
										FErase((Caminho + cFile))
										MemoWrit( CamProc + cFile + ".TXT", cMsgErro)
									endif

								Next Ndp
							Else


								dDVenc		:= StrTran(oFatura:_dup:_dVenc:TEXT,'-','')
								cParcela	:= Iif(Type("oFatura:_dup:_nDup:TEXT") <> "U" , oFatura:_dup:_nDup:TEXT , "" )
								dDVenc		:= CTOD(Subs(dDVenc,7,2)+'/'+Substr(dDVenc,5,2)+'/'+Subs(dDVenc,1,4))


								// desabilita contabilização on-line
								SX1->(DbSeek("FIN040    03"))
								SX1->(Reclock("SX1", .F.))
								SX1->X1_CNT01 := "2"
								SX1->X1_PRESEL := 2
								SX1->(MSUNLOCK())

								lMsErroAuto := .f.
								lMsHelpAuto := .T.

								nModulo := 6  //ESTOQUE
								aadd(aCabSE1,{"E1_PREFIXO"    ,OIdent:_serie:TEXT,Nil,Nil})
								aadd(aCabSE1,{"E1_NUM"		  ,strzero(val(OIdent:_nNF:TEXT),9),Nil,Nil})
								aadd(aCabSE1,{"E1_TIPO"		  ,"NF",Nil,Nil})
								aadd(aCabSE1,{"E1_PARCELA"	  ,cParcela,Nil,Nil})
								aadd(aCabSE1,{"E1_CLIENTE"	  ,_xcliente,Nil,Nil})
								aadd(aCabSE1,{"E1_LOJA"		  ,_xloja,Nil,Nil})
								aadd(aCabSE1,{"E1_NATUREZ"	  ,"611011",Nil,Nil})
								aadd(aCabSE1,{"E1_EMISSAO"	  ,_dData,Nil,Nil})
								aadd(aCabSE1,{"E1_VENCTO"	  ,Iif(dDVenc >= dDatabase , dDVenc ,dDatabase),Nil,Nil})
								aadd(aCabSE1,{"E1_VALOR"	  ,val(oFatura:_dup:_vDup:TEXT),Nil,Nil})
								aadd(aCabSE1,{"E1_MOEDA"	  ,1,Nil,Nil})
								aadd(aCabSE1,{"E1_ORIGEM"	  ,"MATA920",Nil,Nil})

								MSExecAuto({|x,y| FINA040(x,y)},aCabSE1,3) //Inclusao

								// habilita contabilização on-line
								SX1->(DbSeek("FIN040    03"))
								SX1->(Reclock("SX1", .F.))
								SX1->X1_CNT01  := "1"
								SX1->X1_PRESEL := 1
								SX1->(MSUNLOCK())


								IF lMsErroAuto
									aLogX	:= {}

									CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Erro na geração do contas a receber: " + cNota + "/" + cSerie + ". Processamento interrompido!"  )
									aLogX 			:= GetAutoGRLog()
									cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DO CONTAS A RECEBER  "  + cNota + "/" + cSerie + Chr(13) + Chr(10)
									cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
									cMsgErro	+= CHR(13) + CHR(10)

									For x:=1 To Len(aLogX)
										cMsgErro += aLogX[x] + CHR(13) + CHR(10)
									Next x

									fMsgErro(cMsgErro)

									//	LOOP
								Else
									aLogX	:= {}
									CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Titulo a receber gerado com sucesso: " + cNota + "/" + cSerie + "!"  )
									aLogX 			:= GetAutoGRLog()

									cMsgErro 	:= "SUCESSO!!! NA GERAÇÃO DO TÍTULO A RECEBER  "  + cNota + "/" + cSerie + Chr(13) + Chr(10)
									cMsgErro	+= CHR(13) + CHR(10)

									COPY FILE &(Caminho + cFile) TO &(CamProc + cFile)
									FErase((Caminho + cFile))
									MemoWrit( CamProc + cFile + ".TXT", cMsgErro)
								endif


							endif


							//	endif

							//endif


						Else
							CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Nota não tinha titulo a receber: " + cNota + "/" + cSerie + "!" )

							cMsgErro 	:= "SUCESSO!!! NOTA NÃO TINHA TÍTULO A RECEBER  " + cNota + "/" + cSerie +  "!" + Chr(13) + Chr(10)
							cMsgErro	+= CHR(13) + CHR(10)

							COPY FILE &(Caminho + cFile) TO &(CamProc + cFile)
							FErase((Caminho + cFile))
							MemoWrit( CamProc + cFile + ".TXT", cMsgErro)
						endif


					endif
				else //

					CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Nota não tinha titulo a receber: " + cNota + "/" + cSerie + "!" )

					cMsgErro 	:= "SUCESSO!!! NOTA NÃO TINHA TÍTULO A RECEBER  " + cNota + "/" + cSerie +  "!" + Chr(13) + Chr(10)
					cMsgErro	+= CHR(13) + CHR(10)

					COPY FILE &(Caminho + cFile) TO &(CamProc + cFile)
					FErase((Caminho + cFile))
					MemoWrit( CamProc + cFile + ".TXT", cMsgErro)
				endif
			endif

		endif

	Next Nz

//endif
// SGA 	Enddo   // Do While .T.

	PutMV("MV_PCNFE",lPcNfe)
Return


Static Function a370Cabec(nHdlPrv,cArquivo,lCriar)

	lCriar		:=	iif( lCriar == Nil , .f. , lCriar )
	nHdlPrv		:=	HeadProva( "008810" , "FINA040" , Substr(cUsuario,7,6) , @cArquivo , lCriar )
	lCabecalho	:=	.t.

Return

Static Function ValProd()
	_DESCdigit=Alltrim(GetAdvFVal("SB1","B1_DESC",xFilial("SB1")+cEdit1,1,""))
	_NCMdigit=GetAdvFVal("SB1","B1_POSIPI",xFilial("SB1")+cEdit1,1,"")

	dbselectarea("SB1")
	dbsetorder(1)
	dbseek(xFilial("SB1")+cEdit1)

Return 	ExistCpo("SB1")

Static Function Troca()
	Chkproc=.T.
	cProduto=cEdit1
	If Empty(SB1->B1_POSIPI) .and. !Empty(cNCM) .and. cNCM != '00000000'
		RecLock("SB1",.F.)
		Replace B1_POSIPI with cNCM
		MSUnLock()
	endif
	_oDlg:End()
Return

Static Function GetArq(cFile)
	cFile:= cGetFile( "Arquivo NFe (*.xml) | *.xml", "Selecione o Arquivo de Nota Fiscal XML",,'C:\Xml',.F., )
Return cFile

Static Function AchaFile(cCodBar)
	Local cCaminho := Caminho

	aFiles := Directory(cCaminho+"*.XML", "D")  // pega os xmls que estão no diretório // SGA - cCaminho+"\*.XML"

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
Static Function M010inc()

	Local cCodPro	:= GETMV("GR_SGASB1")
	Local aInclSB1	:= {}
	Local x			:= 0
	Local cQry		:= ""

	lMsErroAuto 	:= .F.

	cQry	+= "SELECT TOP 1 B1_COD "
	cQry	+= "FROM "+ RetSqlName("SB1") +" SB1 "
	cQry	+= "WHERE SB1.D_E_L_E_T_ = '' "
	cQry	+= "AND UPPER(SB1.B1_DESC) = UPPER('"+cDescProd+"') "

	cQry 	:= ChangeQuery(cQry)

	TCQUERY cQry NEW ALIAS "ALIAS010" NEW

	xProduto := ALIAS010->B1_COD
//	(ALIAS010)->(DbGoTop())  

//	IF (ALIAS010)->(!EOF())
//		xProduto	:= (ALIAS010)->B1_COD
//	Else
	If Empty(xProduto)
		aInclSB1 := { {"B1_COD" ,cCodPro ,NIL},;
			{"B1_DESC" 		, cDescProd								,NIL},;
			{"B1_TIPO" 		,"PA" 									,Nil},;
			{"B1_UM" 		,"UN" 									,Nil},;
			{"B1_GRUPO"		,"01" 									,Nil},;
			{"B1_LOCPAD" 	,"01" 									,Nil},;
			{"B1_PICM" 		,0 										,Nil},;
			{"B1_IPI" 		,0 										,Nil},;
			{"B1_POSIPI" 	,IF(Empty(cNCM),"00000000",cNCM)		,Nil},;
			{"B1_CLASFIS" 	,"00" 									,Nil},;
			{"B1_SITTRIB" 	,"0" 									,Nil},;
			{"B1_CONTRAT" 	,"N" 									,Nil},;
			{"B1_GARANT" 	,"2" 									,Nil},;
			{"B1_CODGTIN" 	,"SEM GTIN" 							,Nil}}

		MSExecAuto({|x,y| Mata010(x,y)},aInclSB1,3)

		IF lMsErroAuto

			CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Erro no cadastro de produto - "+cCgc+". Processamento interrompido!"  )
			aLogX 			:= GetAutoGRLog()
			cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DO CADASTRO DE PRODUTO  " + cCgc + Chr(13) + Chr(10)
			cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
			cMsgErro	+= CHR(13) + CHR(10)

			For x:=1 To Len(aLogX)
				cMsgErro += aLogX[x] + CHR(13) + CHR(10)
			Next x

			fMsgErro(cMsgErro)

			RollBackSX8()
			ALIAS010->(DbCloseArea())
			return(.F.)
		Else
			PutMV("GR_SGASB1", soma1(cCodPro)) // Atualiza a numeração para o cadastro de produtos
			xProduto	:= cCodPro
		endif
	endif

	ALIAS010->(DbCloseArea())
Return .T.

///////////////////////////////////////////////////////////////////
// Inclusão de Cliente x Produto
///////////////////////////////////////////////////////////////////
Static Function M370inc(nX)

	Reclock("SA7",.t.)
	SA7->A7_FILIAL := xFilial("SA7")
	SA7->A7_CLIENTE := SA1->A1_COD
	SA7->A7_LOJA 	:= SA1->A1_LOJA
	SA7->A7_DESCCLI := cDescProd 	//oDet[nX]:_Prod:_xProd:TEXT
	SA7->A7_PRODUTO := xProduto   	//SB1->B1_COD
	SA7->A7_CODCLI  := cProduto
	SA7->(MsUnlock())

Return

///////////////////////////////////////////////////////////////////
// Inclusão de Cliente x Produto
///////////////////////////////////////////////////////////////////
Static Function MCLIinc()

	Local lRet 		:= .T.
	Local aDados	:= {}
	Local x			:= 0
	Local IAchou	:= .F.
	Local aRetCod	:= {"",""}
	Local cInscEst	:= ""

	If oDestino:_INDIEDEST:TEXT == "2"
		cInscEst	:= "ISENTO"
	elseif Type("oDestino:_ie:TEXT") != "U"
		cInscEst	:= oDestino:_ie:TEXT
	Endif

	lMsErroAuto := .f.
	lMsHelpAuto := .T.

	IAchou  := SA1->(dbSetOrder(3), dbseek(xFilial("SA1")+cCgc))  // se achou (.T.) então altera, caso se não achou (.F.) então inclui

	If IAchou .and. Len(cCgc) < 14  // Conferência para produtor rural que utiliza CPF com inscrição estadual diferente para as fazendas

		aRetCod	:= fCLIincIE(oDestino:_CPF:TEXT, cInscEst)

	Endif

	cCodCli := IIf(IAchou, IIf(!Empty(aRetCod[1]) ,  aRetCod[1]    ,  	SA1->A1_COD   )     ,     GetSXENum("SA1")   )
	clojCli := IIf(IAchou, IIf(!Empty(aRetCod[1]) ,  aRetCod[2]    ,  	SA1->A1_LOJA   )     ,     "00"				 )

	Aadd( aDados, { "A1_COD" ,cCodCli,nil} )
	Aadd( aDados, { "A1_LOJA" ,cLojCli,nil} )
	Aadd( aDados, { "A1_NOME" ,oDestino:_xNome:TEXT,nil} )
	Aadd( aDados, { "A1_NREDUZ" ,oDestino:_xNome:TEXT,nil} )
	Aadd( aDados, { "A1_END" ,oDestino:_enderdest:_xLgr:TEXT+""+oDestino:_enderdest:_nro:TEXT,nil} )
	Aadd( aDados, { "A1_TIPO" ,"R",nil} )
	Aadd( aDados, { "A1_PESSOA" ,IF(Len(cCgc) == 14,"J", "F" ),nil} )
	Aadd( aDados, { "A1_MUN" ,oDestino:_enderdest:_xMun:TEXT,nil} )
	Aadd( aDados, { "A1_EST" ,oDestino:_enderdest:_UF:TEXT,nil} )
	Aadd( aDados, { "A1_COD_MUN" ,Subs(oDestino:_enderdest:_cMun:TEXT,3,5),nil} )
	Aadd( aDados, { "A1_BAIRRO" ,oDestino:_enderdest:_xBairro:TEXT,nil} )
	Aadd( aDados, { "A1_CEP" ,IIf(Type("oDestino:_enderdest:_CEP:TEXT") != "U" , oDestino:_enderdest:_CEP:TEXT , "77000000"  ),nil} )
	Aadd( aDados, { "A1_DDI" , "55",nil} )
	IF Type("oDestino:_enderdest:_fone:TEXT") != "U"
		Aadd( aDados, { "A1_DDD" ,Subs(oDestino:_enderdest:_fone:TEXT,1,2),nil} )
		Aadd( aDados, { "A1_TEL" ,Subs(oDestino:_enderdest:_fone:TEXT,3,10),nil} )
	ELSE // SGA - Incluído
		Aadd( aDados, { "A1_DDD" ,"999",nil} )
		Aadd( aDados, { "A1_TEL" ,"9999-9999",nil} )
	endif
	IF Type("oDestino:_CPF:TEXT") != "U"
		Aadd( aDados, { "A1_CGC" ,oDestino:_CPF:TEXT,nil} )
		Aadd( aDados, { "F" ,"F",nil} )
	Else
		Aadd( aDados, { "A1_CGC" ,oDestino:_CNPJ:TEXT,nil} )
		Aadd( aDados, { "J" ,"F",nil} )
	endif
/* 
	IF Type("oDestino:_ie:TEXT") != "U"
		Aadd( aDados, { "A1_INSCR" ,oDestino:_ie:TEXT,nil} )
	else
		Aadd( aDados, { "A1_INSCR" ,"ISENTO",nil} )
	endif
*/
	Aadd( aDados, { "A1_INSCR" , cInscEst ,nil} )
	Aadd( aDados, { "A1_NATUREZ" ,"101001",nil} )
	Aadd( aDados, { "A1_PAIS" ,"105",nil} )
	Aadd( aDados, { "A1_CONTA" ,"112014102",nil} )
	Aadd( aDados, { "A1_CODPAIS" ,"01058",nil} )

	If IAchou // .T. achou -> altera
		MSExecAuto({|x,y| MATA030(x,y)},aDados,4) //Alteração
	else // .F. não achou -> inclui
		MSExecAuto({|x,y| MATA030(x,y)},aDados,3) //Inclusão
	EndIf

	IF lMsErroAuto

		CONOUT("[" + DTOC(DDATABASE) + "][" + TIME() + "] # [GRIMPSGA] - Erro no cadastro do cliente - "+cCgc+". Processamento interrompido!"  )
		aLogX 			:= GetAutoGRLog()
		cMsgErro 	:= "ATENÇÃO! ERRO NA GRAVAÇÃO DO CADASTRO DE CLIENTE  " + cCgc + Chr(13) + Chr(10)
		cMsgErro	+= "PROCESSO INTERROMPIDO." + CHR(13) + CHR(10)
		cMsgErro	+= CHR(13) + CHR(10)

		For x:=1 To Len(aLogX)
			cMsgErro += aLogX[x] + CHR(13) + CHR(10)
		Next x

		fMsgErro(cMsgErro)

		If !IAchou
			RollBackSX8()
		Endif

		return(.F.)

	Else

		If !IAchou
			ConfirmSx8() // Confirma a gravaçãp do número sxf/sxf
		Endif

		_xcliente := SA1->A1_COD
		_xloja    := SA1->A1_LOJA
	endif

Return lRet


Static Function fCLIincIE(xCGC , xIE)

	Local cQry		:= ""
	Local xCodCliIE	:= ""
	Local xLojCliIE	:= ""

	cQry	+= "SELECT A1_COD , A1_LOJA "
	cQry	+= "FROM "+ RetSqlName("SA1") +" SA1 "
	cQry	+= "WHERE SA1.D_E_L_E_T_ = '' "
	cQry	+= "AND A1_CGC = '" + xCGC +"' "
	cQry	+= "AND A1_INSCR = '" + xIE +"' "

	cQry 	:= ChangeQuery(cQry)

	TCQUERY cQry NEW ALIAS "_SA1" NEW

	xCodCliIE	:=  _SA1->A1_COD
	xLojCliIE	:=  _SA1->A1_LOJA

	_SA1->(DbCloseArea())

Return({xCodCliIE,xLojCliIE})
