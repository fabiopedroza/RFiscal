#INCLUDE "TOPCONN.CH"
#include "COLORS.CH"
#include "Protheus.CH"
#include "inkey.ch"
#include "RWMAKE.ch"

Static __lFirst := .T.

User Function F820Reg()

	Local nRec := PARAMIXB[1] // Recno do registro na SE5 (mov. banc�rio)
	Local aTpReg := PARAMIXB[2] // Array com os tipos de t�tulos a serem considerados como

// aTpReg[1] = Tipos de Nota Fiscal, al�m do NF
// aTpReg[2] = Tipos de Fatura, al�m do FT
// aTpReg[3] = Tipos de Recibo, al�m do RC
// aTpReg[4] = Tipos de Contrato, al�m do C01
// aTpReg[5] = Tipos de Fol. Pagto, al�m do FOL
// O formato � sempre dos demais t�tulos separador por '|' (pipe) entre si. Ex.: 'NCC|NDF'

	Local aRet := {}
	Local aArea := SE5->(GetArea())
	Local lRegValid := .T.

	DbSelectArea("SE5")
	DbGoTo(nRec)


	cQuery := " SELECT REGISTRO FROM R_LCDPR_EXCLUIDOS WHERE REGISTRO = "+alltrim(str(nRec))+" "
	TCQUERY cQuery NEW ALIAS "TCOL"
	nExc:= TCOL->REGISTRO
	TCOL->(dbCloseArea())



	If nRec == nExc
		lRegValid := .F.

	Else

		TCSqlExec("INSERT INTO R_LCDPR (REGISTRO) VALUES ("+alltrim(str(nRec))+")")

	EndIf



	If __lFirst

		aTpReg[1] := {"NCC|NDF"}
		__lFirst := .F.
		aAdd(aRet, lRegValid )
		aAdd(aRet, aTpReg )

	Else

		aAdd(aRet, lRegValid )

	EndIf

	RestArea(aArea)

Return aRet
