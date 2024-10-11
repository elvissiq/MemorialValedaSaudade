#Include 'Protheus.ch'
#Include "TOPCONN.ch"

/*/
/*----------------------------------------------------------------------------------
  Função: LJRECCONF - Customizações na opção Recebimentos (CTRL+D)
  O ponto de entrada LJRECCONF está localizado antes da função de estorno 
  do recebimento de títulos, no módulo Controle de Lojas, rotina Venda Assistida,
  que permite fazer customizações na opção Recebimentos (CTRL+D).
  Retorno:
    lRet (logico)
    Retorno lógico:
    True (.T.) = Confirma
    False (.F.) = Não confirma
  @history
  11/10/2024 - Desenvolvimento da Rotina.
/*/
//------------------------------------------------------------------------------------

User Function LJRECCONF()
  Local lRet  := .T.
  Local cOper := ParamIxB[1]

  If cOper <> '2'
    Return lRet
  EndIF

  IF !LjProfile(8)
    FWAlertError("Usuário sem permissão para realizar estorno.","Estorno de Recebimento.")
    lRet := .F.
  EndIF 
  
Return lRet
