#include "rwmake.ch"
#include "protheus.ch"
#include "TOPCONN.ch"



/*/
+-----------------------------------------------------------------------+
¦Programa  ¦ BLQFIS  | Autor | Hermano Nobre        ¦ Data ¦ 29.04.09 ¦
+----------+------------------------------------------------------------+
¦Descriçào ¦ Programa para alteração do parâmetro MV_DATAFIN. Ao final  ¦
|          | é gravado um log de alteração. 							¦
+----------+------------------------------------------------------------+
¦ Uso      ¦ Microlog - Financeiro                                      ¦
+-----------------------------------------------------------------------+
¦           ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL            ¦
+-----------------------------------------------------------------------+
¦PROGRAMADOR ¦ DATA   ¦ MOTIVO DA ALTERACAO                             ¦
+------------+--------+-------------------------------------------------+
¦            ¦        |                                                 ¦
+------------+--------+-------------------------------------------------+
/*/

User Function BLQFINAN()

	public anterior
	Private oDlg

	SetPrvt("_atual,_data,_ret")

	_atual  := Getmv("MV_DATAFIN")
	anterior:= DtoS(_atual)
	_Data := CtoD(space(8))  
	
	DEFINE MSDIALOG oDlg TITLE "Data para bloqueio de operações financeiras" FROM 96,42 TO 323,505 PIXEL
	
	@ 33,20 Say "Data atual:" Size 043,008 COLOR CLR_BLACK PIXEL OF oDlg
	@ 33,84 Say _atual Size 043,020 COLOR CLR_BLACK PIXEL OF oDlg
	@ 53,20 Say "Informe nova data:" Size 053,008 COLOR CLR_BLACK PIXEL OF oDlg
	@ 53,84 Get _data Picture "99/99/99" Size 053,008 COLOR CLR_BLACK PIXEL OF oDlg
	@ 91,140 Button "Confirmar" Size 40,15 PIXEL OF oDlg Action Atual()
	@ 91,185 Button "Cancelar" Size 40,15 PIXEL OF oDlg Action Close(oDlg)
	
	ACTIVATE MSDIALOG oDlg CENTERED
	
Return


// Gravando novo valor do parâmetro.
Static Function Atual()

If _data = ctod(space(8))
	MsgBox("Você deve preencher uma data.")
	Return
EndIf

dbSelectArea("SX6")
dbSetOrder(1)
If dbSeek("  "+"MV_DATAFIN")

	Reclock("SX6",.F.)
	If _data != ctod(space(8))
		SX6->X6_CONTEUD:= DtoS(_data)
	Endif
	MsUnlock()

	// Gravando log de alteração.
	Grava()
	// Fechando a janela.
	Close(oDlg)

Else
      APMSGSTOP("O parâmetro MV_DATAFIN não foi encontrado.", "TRFINA04")   
Endif

Return


// Gravando log de alteração.
Static Function Grava()

dbSelectArea("SM0")

EMP := M0_CODIGO

_ARQUIVO := "\LOG\DATAFIN"+EMP+DTOS(DATE())+".LOG"


If  !file(_ARQUIVO)
	handle := FCREATE(_ARQUIVO)
ELSE
	handle := FOPEN(_ARQUIVO,2)
	FSEEK(M->handle,0,2)
Endif

_Conteudo := dtoc(date())+" "+time()+" "+Subs(CUSUARIO,7,15)+" "+anterior+" "+dtoc(_data)
FWRITE(M->handle,_Conteudo+CHR(13)+CHR(10))

FCLOSE(M->handle)
Return
