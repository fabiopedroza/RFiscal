#INCLUDE "TOPCONN.CH"
#include "COLORS.CH"
#include "Protheus.CH"
#include "inkey.ch"
#include "RWMAKE.ch"

Static __lFirst := .T.

User Function F820Reg()

	Local nRec := PARAMIXB[1] // Recno do registro na SE5 (mov. bancário)
	Local aTpReg := PARAMIXB[2] // Array com os tipos de títulos a serem considerados como

// aTpReg[1] = Tipos de Nota Fiscal, além do NF
// aTpReg[2] = Tipos de Fatura, além do FT
// aTpReg[3] = Tipos de Recibo, além do RC
// aTpReg[4] = Tipos de Contrato, além do C01
// aTpReg[5] = Tipos de Fol. Pagto, além do FOL
// O formato é sempre dos demais títulos separador por '|' (pipe) entre si. Ex.: 'NCC|NDF'

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
