#INCLUDE "PROTHEUS.CH"
#include "RWMAKE.ch"
#include "Colors.ch"
#include "Font.ch"
#Include "HBUTTON.CH"
#include "Topconn.ch"

User Function PreNotaXML

	//programa para importar notas da empresa 31

	Local aTipo			:={'N','B','D'}
	Local cFile 		:= Space(10)
	Private CPERG   	:="NOTAXML"
	Private Caminho 	:= "C:\Rialma\XML_FERTILIZANTES\"    // ???? "C:\Xml\"
	Private _cMarca   := GetMark()
	Private aFields   := {}
	Private cArq
	Private aFields2  := {}
	Private cArq2
	Private lPcNfe		:= GETMV("MV_PCNFE")

	PutMV("MV_PCNFE",.f.)


	nTipo := 1
	Do While .T.
		cCodBar := Space(100)

		DEFINE MSDIALOG _oPT00005 FROM  50, 050 TO 400,500 TITLE OemToAnsi('Busca de XML de Notas Fiscais de Entrada') PIXEL

		@ 003,005 Say OemToAnsi("Cod Barra NFE") Size 040,030
		@ 030,005 Say OemToAnsi("Tipo Nota Saida:") Size 070,030

		@ 003,060 Get cCodBar  Picture "@!S80" Valid (AchaFile(@cCodBar),If(!Empty(cCodBar),_oPT00005:End(),.t.))  Size 150,030
		@ 020,060 RADIO oTipo VAR nTipo ITEMS "Nota Normal","Nota Beneficiamento","Nota Devolução" SIZE 70,10 OF _oPT00005


		@ 135,060 Button OemToAnsi("Arquivo") Size 036,016 Action (GetArq(@cCodBar),_oPT00005:End())
		@ 135,110 Button OemToAnsi("Ok")  Size 036,016 Action (_oPT00005:End())
		@ 135,160 Button OemToAnsi("Sair")   Size 036,016 Action Fecha()

		Activate Dialog _oPT00005 CENTERED

		MV_PAR01 := nTipo

		cFile := cCodBar

		If !File(cFile) .and. !Empty(cFile)
			MsgAlert("Arquivo Não Encontrado no Local de Origem Indicado!")
			PutMV("MV_PCNFE",lPcNfe)
			Return
		Endif

		Private nHdl    := fOpen(cFile,0)


		aCamposPE:={}

		If nHdl == -1
			If !Empty(cFile)
				MsgAlert("O arquivo de nome "+cFile+" nao pode ser aberto! Verifique os parametros.","Atencao!")
			Endif
			PutMV("MV_PCNFE",lPcNfe)
			Return
		Endif
		nTamFile := fSeek(nHdl,0,2)
		fSeek(nHdl,0,0)
		cBuffer  := Space(nTamFile)                // Variavel para criacao da linha do registro para leitura
		nBtLidos := fRead(nHdl,@cBuffer,nTamFile)  // Leitura  do arquivo XML
		fClose(nHdl)

		cAviso := ""
		cErro  := ""
		oNfe := XmlParser(cBuffer,"_",@cAviso,@cErro)
		Private oNF
		Private lMsErroAuto := .f.
		Private lMsHelpAuto := .T.

		If !ExistDir("C:\XML")
			MakeDir("C:\XML")
		Endif

		If Type("oNFe:_NfeProc")<> "U"
			oNF := oNFe:_NFeProc:_NFe
		Else
			_chvnfe := oNFe:_ProcEventoNFe:_Evento:_InfEvento:_Chnfe:TEXT  //cancelamento de nota fiscal
//			alert(_chvnfe)
			aCabe02 := {}
			Dbselectarea("SF2")
			Dbsetorder(10)//filial+chave da nota
			IF Dbseek(xFilial("SF2")+_chvnfe)
				aadd(aCabe02,{"F2_FILIAL" ,SF2->F2_FILIAL,Nil,Nil})
				aadd(aCabe02,{"F2_TIPO"   ,SF2->F2_TIPO,Nil,Nil})
				aadd(aCabe02,{"F2_FORMUL" ,"S",Nil,Nil})
				aadd(aCabe02,{"F2_DOC"    ,SF2->F2_DOC,Nil,Nil})
				aadd(aCabe02,{"F2_SERIE"  ,SF2->F2_SERIE,Nil,Nil})
				aadd(aCabe02,{"F2_EMISSAO",SF2->F2_EMISSAO,Nil,Nil})
				aadd(aCabe02,{"F2_CLIENTE",SF2->F2_CLIENTE,Nil,Nil})
				aadd(aCabe02,{"F2_LOJA"   ,SF2->F2_LOJA,Nil,Nil})
				aadd(aCabe02,{"F2_ESPECIE","SPED",Nil,Nil})
				aadd(aCabe02,{"F2_CHVNFE"  ,SF2->F2_CHVNFE,Nil,Nil})
				aadd(aCabe02,{"F2_COND","001"})
				aadd(aCabe02,{"F2_DESCONT",0})
				aadd(aCabe02,{"F2_FRETE",0})
				aadd(aCabe02,{"F2_SEGURO",0})
				aadd(aCabe02,{"F2_DESPESA",0})

				_xnumnf := SF2->F2_DOC
				aIte02 := {}
				Dbselectarea("SD2")
				dbsetorder(3) //filial+nota+serie
				Dbseek(xFilial("SD2")+SF2->F2_DOC+SF2->F2_SERIE)
				Do While !eof() .and. SF2->F2_DOC+SF2->F2_SERIE == SD2->D2_DOC+SD2->D2_SERIE

					aLinh2 := {}
					aadd(aLinh2,{"D2_COD",SD2->D2_COD,Nil,Nil})
					aadd(aLinh2,{"D2_ITEM",SD2->D2_ITEM,Nil,Nil})
					aadd(aLinh2,{"D2_QUANT",SD2->D2_QUANT,Nil,Nil})
					aadd(aLinh2,{"D2_PRCVEN",SD2->D2_PRCVEN,Nil,Nil})
					aadd(aLinh2,{"D2_TOTAL",SD2->D2_TOTAL,Nil,Nil})
					aadd(aLinh2,{"D2_TES",SD2->D2_TES,Nil,Nil})
					aadd(aLinh2,{"D2_CLASFIS",SD2->D2_CLASFIS,Nil,Nil})
//					aadd(aLinh2,{"D2_PICM",SD2->D2_PCIM,Nil,Nil})
//					aadd(aLinh2,{"D2_VALICM",SD2->D2_PICM,Nil,Nil})
					aadd(aIte02,aLinh2)

					SD2->(Dbskip())
				Enddo

				MSExecAuto({|x,y,z|MATA920(x,y,z)},aCabe02,aIte02,5)
				IF lMsErroAuto
					MSGALERT("ERRO NO PROCESSO - PRÉ NOTA")
					mostraerro()
				Else
					Alert("Nota "+_xnumnf+" excluida com sucesso!!!")
				Endif

			Else
				Alert("Nota a ser excluida não encontrada, verifique XML")
			Endif
			Return(.t.)
		Endif

		Private oEmitente  := oNF:_InfNfe:_Emit
		Private oIdent     := oNF:_InfNfe:_IDE
		Private oDestino   := oNF:_InfNfe:_Dest
		Private oTotal     := oNF:_InfNfe:_Total
		Private oTransp    := oNF:_InfNfe:_Transp
		Private oDet       := oNF:_InfNfe:_Det
		If Type("oNF:_InfNfe:_ICMS")<> "U"
			Private oICM       := oNF:_InfNfe:_ICMS
		Else
			Private oICM		:= nil
		Endif
		Private oFatura    := IIf(Type("oNF:_InfNfe:_Cobr")=="U",Nil,oNF:_InfNfe:_Cobr)
		Private cEdit1	   := Space(15)
		Private _DESCdigit :=space(55)
		Private _NCMdigit  :=space(8)

		oDet := IIf(ValType(oDet)=="O",{oDet},oDet)
		// Validações -------------------------------
		// -- CNPJ da NOTA = CNPJ do CLIENTE ? oEmitente:_CNPJ
		If MV_PAR01 = 1
			cTipo := "N"
		ElseIF MV_PAR01 = 2
			cTipo := "B"
		ElseIF MV_PAR01 = 3
			cTipo := "D"
		Endif

		// CNPJ ou CPF

		cCgc := AllTrim(IIf(Type("oDestino:_CPF")=="U",oDestino:_CNPJ:TEXT,oDestino:_CPF:TEXT))
		aDados := {}

		If !SA1->(dbSetOrder(3), dbSeek(xFilial("SA1")+cCgc))
			cCodCli := GetSXENum("SA1")
			clojCli := "01"
			//Aadd( aDados, { "A1_COD" ,cCodCli,nil} )
			//Aadd( aDados, { "A1_LOJA" ,cLojCli,nil} )
			Aadd( aDados, { "A1_NOME" ,oDestino:_xNome:TEXT,nil} )
			Aadd( aDados, { "A1_NREDUZ" ,oDestino:_xNome:TEXT,nil} )
			Aadd( aDados, { "A1_END" ,oDestino:_enderdest:_xLgr:TEXT+""+oDestino:_enderdest:_nro:TEXT,nil} )
			Aadd( aDados, { "A1_TIPO" ,"R",nil} )
			Aadd( aDados, { "A1_PESSOA" ,"J",nil} )
			Aadd( aDados, { "A1_MUN" ,oDestino:_enderdest:_xMun:TEXT,nil} )
			Aadd( aDados, { "A1_EST" ,oDestino:_enderdest:_UF:TEXT,nil} )
			Aadd( aDados, { "A1_COD_MUN" ,Subs(oDestino:_enderdest:_cMun:TEXT,3,5),nil} )
			Aadd( aDados, { "A1_BAIRRO" ,oDestino:_enderdest:_xBairro:TEXT,nil} )
			Aadd( aDados, { "A1_CEP" ,oDestino:_enderdest:_CEP:TEXT,nil} )
			IF Type("oDestino:_enderdest:_fone:TEXT") != "U"
				Aadd( aDados, { "A1_DDD" ,Subs(oDestino:_enderdest:_fone:TEXT,1,2),nil} )
				Aadd( aDados, { "A1_TEL" ,Subs(oDestino:_enderdest:_fone:TEXT,3,10),nil} )
			ENDIF
			IF Type("oDestino:_CPF:TEXT") != "U"
				Aadd( aDados, { "A1_CGC" ,oDestino:_CPF:TEXT,nil} )
				Aadd( aDados, { "F" ,"F",nil} )
			Else
				Aadd( aDados, { "A1_CGC" ,oDestino:_CNPJ:TEXT,nil} )
				Aadd( aDados, { "J" ,"F",nil} )
			Endif
			IF Type("oDestino:_ie:TEXT") != "U"
				Aadd( aDados, { "A1_INSCR" ,oDestino:_ie:TEXT,nil} )
			else
				Aadd( aDados, { "A1_INSCR" ,"ISENTO",nil} )
			Endif
			Aadd( aDados, { "A1_NATUREZ" ,"101001",nil} )
			Aadd( aDados, { "A1_PAIS" ,"105",nil} )
			Aadd( aDados, { "A1_CONTA" ,"112014102",nil} )
			Aadd( aDados, { "A1_CODPAIS" ,"01058",nil} )


			MSExecAuto({|x,y| MATA030(x,y)},aDados,3) //Inclusao

			IF lMsErroAuto
				MSGALERT("ERRO NO PROCESSO - CAD. CLIENTE")
				MostraErro()

				return(.T.)
			endif


			_xcliente := SA1->A1_COD
			_xloja    := SA1->A1_LOJA
		Else
			_xcliente := SA1->A1_COD
			_xloja    := SA1->A1_LOJA
		Endif

		//faz o calculo dos espacos a inserir no campo F1_Doc
		cDoc := len(SF2->F2_DOC)
		cCn  := len(OIdent:_nNF:TEXT)
		cSpa := cDoc - cCn

		cNota :=OIdent:_nNF:TEXT+SPACE(cSpa)

		// -- Nota Fiscal já existe na base ?
		If SF2->(DbSeek(XFilial("SF2")+cNota+Padr(OIdent:_serie:TEXT,3)+SA1->A1_COD+SA1->A1_LOJA))
			MsgAlert("Nota No.: "+OIdent:_nNF:TEXT+"/"+OIdent:_serie:TEXT+" do Cliente "+SA1->A1_COD+"/"+SA1->A1_LOJA+" Ja Existe. A Importacao sera interrompida")
			PutMV("MV_PCNFE",lPcNfe)
			Return Nil
		EndIf

		Private lMsErroAuto := .f.
		Private lMsHelpAuto := .T.
		aCabec := {}
		aItens := {}
		cChave := oNFe:_NfeProc:_PROTNFE:_INFPROT:_CHNFE:TEXT
		cNota  := OIdent:_nNF:TEXT
		aadd(aCabec,{"F2_TIPO"   ,If(MV_PAR01==1,"N",If(MV_PAR01==2,'B','D')),Nil,Nil})
		aadd(aCabec,{"F2_FORMUL" ,"N",Nil,Nil})
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

			cProduto:=PadR(AllTrim(oDet[nX]:_Prod:_cProd:TEXT),20)
			xProduto:=cProduto

			cNCM:=IIF(Type("oDet[nX]:_Prod:_NCM")=="U",space(12),oDet[nX]:_Prod:_NCM:TEXT)
			Chkproc=.F.

			//faz o calculo dos espacos a inserir no campo F1_Doc
			cPro := len(xProduto)
			cTam := 20 - cPro // 20 eh o tamanho do campo sa5_codprf
			cProdt :=xProduto+SPACE(cTam)

			If MV_PAR01 = 1
				SA7->(dbsetorder(3))   // FILIAL + FORNECEDOR + LOJA + CODIGO PRODUTO NO FORNECEDOR

				If !SA7->(dbSeek(xFilial("SA7")+SA1->A1_COD+SA1->A1_LOJA+cProduto))
					If !MsgYesNo ("Produto Cod.: "+cProduto+" Nao Encontrado. Digita Codigo de Substituicao?")
						PutMV("MV_PCNFE",lPcNfe)
						Return Nil
					Endif
					DEFINE MSDIALOG _oDlg TITLE "Dig.Cod.Substituicao" FROM C(177),C(192) TO C(509),C(659) PIXEL

					// Cria as Groups do Sistema
					@ C(002),C(003) TO C(071),C(186) LABEL "Dig.Cod.Substituicao " PIXEL OF _oDlg

					// Cria Componentes Padroes do Sistema
					@ C(012),C(027) Say "Produto: "+cProduto+" - NCM: "+cNCM Size C(150),C(008) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(020),C(027) Say "Descricao: "+oDet[nX]:_Prod:_xProd:TEXT Size C(150),C(008) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(028),C(070) MsGet oEdit1 Var cEdit1 F3 "SB1" Valid(ValProd()) Size C(060),C(009) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(040),C(027) Say "Produto digitado: "+cEdit1+" - NCM: "+_NCMdigit Size C(150),C(008) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(048),C(027) Say "Descricao: "+_DESCdigit Size C(150),C(008) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(004),C(194) Button "Processar" Size C(037),C(012) PIXEL OF _oDlg Action(Troca())
					@ C(025),C(194) Button "Cancelar" Size C(037),C(012) PIXEL OF _oDlg Action(_oDlg:End())
					oEdit1:SetFocus()

					ACTIVATE MSDIALOG _oDlg CENTERED
					If Chkproc!=.T.
						MsgAlert("Produto Cod.: "+cProduto+" Nao Encontrado. A Importacao sera interrompida")
						PutMV("MV_PCNFE",lPcNfe)
						Return Nil
					Else
						If SA7->(dbSetOrder(1), dbSeek(xFilial("SA7")+SA1->A1_COD+SA1->A1_LOJA+SB1->B1_COD))
							RecLock("SA7",.f.)
						Else
							Reclock("SA7",.t.)
						Endif

						SA7->A7_FILIAL := xFilial("SA7")
						SA7->A7_CLIENTE := SA1->A1_COD
						SA7->A7_LOJA 	:= SA1->A1_LOJA
						SA7->A7_DESCCLI := oDet[nX]:_Prod:_xProd:TEXT
						SA7->A7_PRODUTO := SB1->B1_COD
						SA7->A7_CODCLI  := xProduto
						SA7->(MsUnlock())

					EndIf
				Else
					SB1->(dbSetOrder(1), dbSeek(xFilial("SB1")+SA7->A7_PRODUTO))
					If !Empty(cNCM) .and. cNCM != '00000000' .And. SB1->B1_POSIPI <> cNCM
						RecLock("SB1",.F.)
						Replace B1_POSIPI with cNCM
						MSUnLock()
					Endif
				Endif
			Else

				// FILIAL + FORNECEDOR + LOJA + CODIGO PRODUTO NO FORNECEDOR
				SA5->(dbSetOrder(14))
				IF  !SA5->(dbSeek(xFilial("SA5")+SA2->A2_COD+SA2->A2_LOJA+cProdt))
					If !MsgYesNo ("Produto Cod.: "+cProduto+" Nao Encontrado. Digita Codigo de Substituicao?")
						PutMV("MV_PCNFE",lPcNfe)
						Return Nil
					Endif
					DEFINE MSDIALOG _oDlg TITLE "Dig.Cod.Substituicao" FROM C(177),C(192) TO C(509),C(659) PIXEL

					// Cria as Groups do Sistema
					@ C(002),C(003) TO C(071),C(186) LABEL "Dig.Cod.Substituicao " PIXEL OF _oDlg

					// Cria Componentes Padroes do Sistema
					@ C(012),C(027) Say "Produto: "+cProduto+" - NCM: "+cNCM Size C(150),C(008) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(020),C(027) Say "Descricao: "+oDet[nX]:_Prod:_xProd:TEXT Size C(150),C(008) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(028),C(070) MsGet oEdit1 Var cEdit1 F3 "SB1" Valid(ValProd()) Size C(060),C(009) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(040),C(027) Say "Produto digitado: "+cEdit1+" - NCM: "+_NCMdigit Size C(150),C(008) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(048),C(027) Say "Descricao: "+_DESCdigit Size C(150),C(008) COLOR CLR_HBLUE PIXEL OF _oDlg
					@ C(004),C(194) Button "Processar" Size C(037),C(012) PIXEL OF _oDlg Action(Troca())
					@ C(025),C(194) Button "Cancelar" Size C(037),C(012) PIXEL OF _oDlg Action(_oDlg:End())
					oEdit1:SetFocus()

					ACTIVATE MSDIALOG _oDlg CENTERED
					If Chkproc != .T.
						MsgAlert("Produto Cod.: "+cProduto+" Nao Encontrado. A Importacao sera interrompida")
						PutMV("MV_PCNFE",lPcNfe)
						Return Nil
					Else
						If SA5->(dbSetOrder(1), dbSeek(xFilial("SA5")+SA2->A2_COD+SA2->A2_LOJA+SB1->B1_COD))
							RecLock("SA5",.f.)
						Else
							Reclock("SA5",.t.)
						Endif

						SA5->A5_FILIAL := xFilial("SA5")
						SA5->A5_FORNECE := SA2->A2_COD
						SA5->A5_LOJA 	:= SA2->A2_LOJA
						SA5->A5_NOMEFOR := SA2->A2_NOME
						SA5->A5_PRODUTO := SB1->B1_COD
						SA5->A5_NOMPROD := oDet[nX]:_Prod:_xProd:TEXT
						SA5->A5_CODPRF  := xProduto
						SA5->(MsUnlock())

					EndIf
				Else
					SB1->(dbSetOrder(1), dbSeek(xFilial("SB1")+SA5->A5_PRODUTO))
				Endif
			Endif
			SB1->(dbSetOrder(1))

			cProds += ALLTRIM(SB1->B1_COD)+'/'

			AAdd(aPedIte,{SB1->B1_COD,Val(oDet[nX]:_Prod:_qTrib:TEXT),Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qCom:TEXT),6),Val(oDet[nX]:_Prod:_vProd:TEXT)})

		Next nX


		PUBLIC ticm_aux := 0
		For nX := 1 To Len(oDet)

			aLinha := {}
			// Validacao: Produto Existe no SB1 ?
			// Se não existir, abrir janela c/ codigo da NF e descricao para digitacao do cod. substituicao.
			// Descricao: oDet[nX]:_Prod:_xProd:TEXT

			cProduto:=Right(AllTrim(oDet[nX]:_Prod:_cProd:TEXT),15)
			xProduto:=cProduto

			//faz o calculo dos espacos a inserir no campo F1_Doc
			cPro := len(xProduto)
			cTam := 20 - cPro
			cProdt :=xProduto+SPACE(cTam)

			cNCM:=IIF(Type("oDet[nX]:_Prod:_NCM")=="U",space(12),oDet[nX]:_Prod:_NCM:TEXT)
			Chkproc=.F.

			If MV_PAR01 == 1
				SA7->(Dbsetorder(3))
				SA7->(dbSeek(xFilial("SA7")+SA1->A1_COD+SA1->A1_LOJA+cProduto))
				SB1->(dbSetOrder(1) , dbSeek(xFilial("SB1")+SA7->A7_PRODUTO))
			Else
				SA5->(dbSetOrder(14), dbSeek(xFilial("SA5")+SA2->A2_COD+SA2->A2_LOJA+cProdt))
				SB1->(dbSetOrder(1) , dbSeek(xFilial("SB1")+SA5->A5_PRODUTO))
			Endif

			//SE FOR NOTA DE PALLET
			//		if cProduto = "46901" .or. cProduto = "46072"
			//			SB1->(dbSetOrder(1) , dbSeek(xFilial("SB1")+"PALLET"))
			//		endif

			aadd(aLinha,{"D2_COD",SB1->B1_COD,Nil,Nil})
			aadd(aLinha,{"D2_ITEM","0001",Nil,Nil})
			If Val(oDet[nX]:_Prod:_qTrib:TEXT) != 0
				aadd(aLinha,{"D2_QUANT",Val(oDet[nX]:_Prod:_qTrib:TEXT),Nil,Nil})
				aadd(aLinha,{"D2_PRCVEN",Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qTrib:TEXT),6),Nil,Nil})
			Else
				aadd(aLinha,{"D2_QUANT",Val(oDet[nX]:_Prod:_qCom:TEXT),Nil,Nil})
				aadd(aLinha,{"D2_PRCVEN",Round(Val(oDet[nX]:_Prod:_vProd:TEXT)/Val(oDet[nX]:_Prod:_qCom:TEXT),6),Nil,Nil})
			Endif
			aadd(aLinha,{"D2_TOTAL",Val(oDet[nX]:_Prod:_vProd:TEXT),Nil,Nil})
			_cfop:=oDet[nX]:_Prod:_CFOP:TEXT


			IF Alltrim(_cfop) == "5101" .or. Alltrim(_cfop) == "6101"
				aadd(aLinha,{"D2_TES","539",Nil,Nil})
			ELSEIF Alltrim(_cfop) == "5922" .or. Alltrim(_cfop) =="6922"
				aadd(aLinha,{"D2_TES","606",Nil,Nil})
			ELSEIF Alltrim(_cfop) == "5116" .or. Alltrim(_cfop) == "6116"
				aadd(aLinha,{"D2_TES","607",Nil,Nil})
			ELSEIF Alltrim(_cfop) == "5910" .or. Alltrim(_cfop) == "6910"
				aadd(aLinha,{"D2_TES","555",Nil,Nil})
			ELSEIF Alltrim(_cfop) == "5933" .or. Alltrim(_cfop) == "6933"
				aadd(aLinha,{"D2_TES","521",Nil,Nil})
			ELSE
				ALERT("CFOP SEM REGRA PARA TES, ACIONE SUPORTE !!!")
				Return(.t.)
			ENDIF

//			If Left(Alltrim(_cfop),1)="5"
//				_cfop:=Stuff(_cfop,1,1,"1")
			//Else
//				_cfop:=Stuff(_cfop,1,1,"2")
			//Endif

			If Type("oDet[nX]:_Prod:_vDesc")<> "U"
				aadd(aLinha,{"D2_DESCONT",Val(oDet[nX]:_Prod:_vDesc:TEXT),Nil,Nil})
			Endif
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
				Endif
			Endif


			aadd(aItens,aLinha)
		Next nX


		//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
		//| Teste de Inclusao                                            |
		//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

		Public _mvDATPGTO := GetMv("MV_DATPGTO")
		Public _venc1par  := GetMv("MV_DATPGTO")+30

		If Len(aItens) > 0
			Private lMsErroAuto := .f.
			Private lMsHelpAuto := .T.

			SB1->( dbSetOrder(1) )
			SA1->( dbSetOrder(1) )

			nModulo := 9  //ESTOQUE
			MSExecAuto({|x,y,z|MATA920(x,y,z)},aCabec,aItens,3)

			IF lMsErroAuto
				If !ExistDir("C:\XML\ERRO\")
					MakeDir("C:\XML\ERRO\")
				Endif
				xFile := STRTRAN(Upper(cFile),"C:\XML\","C:\XML\ERRO\")
				COPY FILE &cFile TO &xFile
				FErase(cFile)
				MSGALERT("ERRO NO PROCESSO")
				MostraErro()
			Else
				If !ExistDir("C:\XML\PROCESSADAS\")
					MakeDir("C:\XML\PROCESSADAS\")
				Endif
				xFile := STRTRAN(Upper(cFile),"C:\XML\",  "C:\XML\PROCESSADAS\")
				COPY FILE &cFile TO &xFile
				FErase(cFile)
				MSGALERT(Alltrim(aCabec[3,2])+"  -  Nota Gerada Com Sucesso!")

				//Gera contas a receber - inicio
				aCabSE1     := {}
				cArquivo    :=	""
				nHdlPrv	    :=	0
				cArquivo	:=	""
				nTotal		:=	0
				cNomProg	:=	"CONT_SE1"

				IF Type("oFatura:_fat:_vLiq:TEXT") <> "U"
					IF val(oFatura:_fat:_vLiq:TEXT) > 0

						nModulo := 6  //ESTOQUE
						aadd(aCabSE1,{"E1_PREFIXO"    ,OIdent:_serie:TEXT,Nil,Nil})
						aadd(aCabSE1,{"E1_NUM"		  ,strzero(val(OIdent:_nNF:TEXT),9),Nil,Nil})
						aadd(aCabSE1,{"E1_TIPO"		  ,"NF",Nil,Nil})
						aadd(aCabSE1,{"E1_CLIENTE"	  ,_xcliente,Nil,Nil})
						aadd(aCabSE1,{"E1_LOJA"		  ,_xloja,Nil,Nil})
						aadd(aCabSE1,{"E1_NATUREZ"	  ,"611011",Nil,Nil})
						aadd(aCabSE1,{"E1_EMISSAO"	  ,_dData,Nil,Nil})
						aadd(aCabSE1,{"E1_VENCTO"	  ,_dData+30,Nil,Nil})
						aadd(aCabSE1,{"E1_VALOR"	  ,val(oFatura:_fat:_vLiq:TEXT),Nil,Nil})
						aadd(aCabSE1,{"E1_MOEDA"	  ,1,Nil,Nil})

						MSExecAuto({|x,y| FINA040(x,y)},aCabSE1,3) //Inclusao

					Endif

/*
					//Gera contas a receber - fim

					//contabiliza contas a receber - inicio
					a370Cabec(@nHdlPrv,@cArquivo)

					ddatabase := SE1->E1_EMISSAO
					nTotal	+=	DetProva( nHdlPrv , "400" , cNomProg , "008810" )

					RodaProva( nHdlPrv , nTotal )
					cA100Incl( cArquivo , nHdlPrv , 3 , "008810" , .T. , .f. )

					RecLock("SE1",.f.)
					SE1->E1_LA	:=	"S"
					MsUnlock("SE1")
					//contabiliza contas a receber - inicio
*/
				Endif
			Endif

		EndIf

		//Endif
	Enddo
	PutMV("MV_PCNFE",lPcNfe)
Return

Static Function a370Cabec(nHdlPrv,cArquivo,lCriar)

	lCriar		:=	iif( lCriar == Nil , .f. , lCriar )
	nHdlPrv		:=	HeadProva( "008810" , "FINA040" , Substr(cUsuario,7,6) , @cArquivo , lCriar )
	lCabecalho	:=	.t.

Return

Static Function ValProd()
	_DESCdigit=Alltrim(GetAdvFVal("SB1","B1_DESC",XFilial("SB1")+cEdit1,1,""))
	_NCMdigit=GetAdvFVal("SB1","B1_POSIPI",XFilial("SB1")+cEdit1,1,"")

	Dbselectarea("SB1")
	Dbsetorder(1)
	Dbseek(xFilial("SB1")+cEdit1)

Return 	ExistCpo("SB1")

Static Function Troca()
	Chkproc=.T.
	cProduto=cEdit1
	If Empty(SB1->B1_POSIPI) .and. !Empty(cNCM) .and. cNCM != '00000000'
		RecLock("SB1",.F.)
		Replace B1_POSIPI with cNCM
		MSUnLock()
	Endif
	_oDlg:End()
Return

Static Function GetArq(cFile)
	cFile:= cGetFile( "Arquivo NFe (*.xml) | *.xml", "Selecione o Arquivo de Nota Fiscal XML",,'C:\Xml',.F., )
Return cFile


StatiC Function Fecha()
	Close(_oPT00005)
Return


Static Function AchaFile(cCodBar)
	Local aCompl := {}
	Local cCaminho := Caminho
	Local lOk 	:= .f.
	Local oNf
	Local oNfe
	Local nArq	:= 0

	If Empty(cCodBar)
		Return .t.
	Endif

	aFiles := Directory(cCaminho+"\*.XML", "D")

	For nArq := 1 To Len(aFiles)
		cFile := AllTrim(cCaminho+aFiles[nArq,1])

		nHdl    := fOpen(cFile,0)
		nTamFile := fSeek(nHdl,0,2)
		fSeek(nHdl,0,0)
		cBuffer  := Space(nTamFile)                // Variavel para criacao da linha do registro para leitura
		nBtLidos := fRead(nHdl,@cBuffer,nTamFile)  // Leitura  do arquivo XML
		fClose(nHdl)
		If AT(AllTrim(cCodBar),AllTrim(cBuffer)) > 0
			cCodBar := cFile
			lOk := .t.
			Exit
		Endif
	Next nArq
	If !lOk
		Alert("Nenhum Arquivo Encontrado, Por Favor Selecione a Opção Arquivo e Faça a Busca na Arvore de Diretórios!")
	Endif

Return lOk

