#INCLUDE "RWMAKE.CH" 

User Function XSCRPED()   

Local aArea 		:= GetArea()
Local aAreaSL1 		:= SL1->(GetArea())
Local aAreaSL2 		:= SL2->(GetArea())
Local nOrcam		:= 0
Local sTexto        := ""              
Local nCheques		:= 0
Local nCartaoC		:= 0
Local nCartaoD		:= 0
Local nConveni		:= 0
Local nVales		:= 0
Local nFinanc		:= 0
Local nCredito		:= 0
Local nOutros		:= 0
Local cQuant 		:= ""	
Local cVrUnit		:= "" 								// Valor unitário
Local cDesconto		:= ""
Local cVlrItem		:= ""
Local nVlrIcmsRet	:= 0								// Valor do icms retido (Substituicao tributaria)
Local nTroco		:= 0
Local lMvLjTroco  	:= SuperGetMV("MV_LJTROCO", ,.F. )	// Verifica se utiliza troco nas diferentes formas de pagamento
Local lMvGarFP		:= SuperGetMV("MV_LJGarFP", ,.F.)	// Define o conc
Local lLibQtdGE		:= SuperGetMv("MV_LJLIBGE", , .F.) 	// Libera quantidade da garantia estendida  .. por default é Falso
Local cServType		:= SuperGetMv("MV_LJTPSF",,"SF")	// Define o tipo do produto de servico financeiro
Local nValTot		:= 0
Local nDescTot		:= 0
Local nFatorRes		:= 1
Local nValPag		:= 0
Local nVlrDescIt	:= 0
Local cGarType		:= SuperGetMv("MV_LJTPGAR",,"GE")	// Define o tipo do produto de garantia estendida
Local nVlrFSD		:= 0								// Valor do frete + seguro + despesas
Local cDocPed       := SL1->L1_DOC      				// Documento do Pedido
Local cSerPed       := SL1->L1_SERIE					// Serie do pedido
Local nVlrTot       := 0                                // Valor Total
Local nVlrAcres     := 0                                // Valor Acrescimo
Local lMvArrefat    := SuperGetMv("MV_ARREFAT") == "S"
Local nTotDesc		:= 0 								// Total de desconto de acordo com L2_DESCPRO
Local lPedido		:= .F.								// Indica se a venda tem itens com pedido
Local aVlrFormas	:= SCRPRetPgt()						// Resgata os valores de cada forma de pagamento
Local nMaxChar 		:= 47
Local aFieldSM0     := { ;
							"M0_NOMECOM",;   //Posição [1]
							"M0_ENDENT",;    //Posição [2]
							"M0_BAIRENT",;   //Posição [3]
							"M0_CIDENT",;    //Posição [4]
							"M0_ESTENT",;    //Posição [5]
							"M0_CEPENT",;    //Posição [6]
							"M0_CGC",;       //Posição [7]
							"M0_INSC",;      //Posição [8]
							"M0_COMPENT",;   //Posição [9]
							"M0_TEL";		 //Posição [10]
							}        
Local aSM0Data 		:= FWSM0Util():GetSM0Data(, SL1->L1_FILIAL, aFieldSM0)
Local cNomCom       := aSM0Data[1,2] // Nome Comercial da Empresa
Local cEndEnt       := aSM0Data[2,2] // Endereço de Entrega
Local cBaiEnt       := aSM0Data[3,2] // Bairro de Entrega
Local cCidEnt       := aSM0Data[4,2] // Cidade de Entrega
Local cEstEnt       := aSM0Data[5,2] // Estado de Entrega
Local cCepEnt       := aSM0Data[6,2] // Cep de Entrega
Local cCgcEnt       := aSM0Data[7,2] // CNPJ 
Local cInsEnt       := aSM0Data[8,2] // Inscrição Estadual
Local cNomCli       := Posicione("SA1",1,xFILIAL("SA1")+SL1->L1_CLIENTE+SL1->L1_LOJA,"A1_NREDUZ") //Nome do Cliente
Local cNomVen       := Posicione("SA3",1,xFilial("SA3")+SL1->L1_VEND,"A3_NOME")  // Nome do Vendedor
Local cNomOpe       := Posicione("SA6",1,xFilial("SA6")+SL1->L1_OPERADO,"A6_NOME")  // Nome do Operador

sTexto:= '<ce>'+ alltrim(cNomCom) +'</ce>'+ Chr(13)+ Chr(10)
sTexto:= sTexto+'<ce>'+ alltrim(cEndEnt) + ' - '+ alltrim(cBaiEnt) +'</ce>'+ Chr(13)+ Chr(10)
sTexto:= sTexto+'<ce>'+ alltrim(cCidEnt) + ' - '+ alltrim(cEstEnt) + ' CEP:'+ alltrim(cCepEnt) +'</ce>'+ Chr(13)+ Chr(10)
sTexto:= sTexto+'<ce> CNPJ: '+ alltrim(cCgcEnt) + ' IE: '+ alltrim(cInsEnt) +'</ce>'+ Chr(13)+ Chr(10)

sTexto := sTexto + '==============================================='	+Chr(13)+Chr(10)
sTexto := sTexto + '      COMPROVANTE DE VENDA NAO FISCAL          '	+Chr(13)+Chr(10)
sTexto := sTexto + '==============================================='	+Chr(13)+Chr(10)

sTexto:= sTexto + 'Codigo         Descricao'+Chr(13)+Chr(10)
sTexto:= sTexto + 'Qtd             VlrUnit              VlrTot'+ Chr(13)+ Chr(10)
sTexto:= sTexto + '-----------------------------------------------'+ Chr(13)+ Chr(10)

dbSelectArea("SL1")                                                                  
dbSetOrder(1)  

nOrcam		:= SL1->L1_NUM
nTroco		:= Iif(SL1->(FieldPos("L1_TROCO1")) > 0,(nFatorRes * SL1->L1_TROCO1), 0)
nDinheir	:= (nFatorRes * aVlrFormas[01][02] )
nCheques	:= (nFatorRes * aVlrFormas[02][02] )
nCartaoC 	:= (nFatorRes * aVlrFormas[03][02] )
nCartaoD 	:= (nFatorRes * aVlrFormas[04][02] )
nPIX	 	:= (nFatorRes * aVlrFormas[05][02] )
nCartDig 	:= (nFatorRes * aVlrFormas[06][02] )
nFinanc		:= (nFatorRes * aVlrFormas[07][02] )
nConveni	:= (nFatorRes * aVlrFormas[08][02] )
nVales  	:= (nFatorRes * aVlrFormas[09][02] )
nCredito	:= (nFatorRes * aVlrFormas[10][02] )
nOutros		:= (nFatorRes * aVlrFormas[11][02] )
nValTot		:= 0
nDescTot	:= 0

/* Soma o valor de todas as formas de pagamento
Necessariio dar um round em cada forma para verificar se ha diferença de arredondamento no somatorio dos pagamentos*/
nValPag :=	Round(nDinheir,2)	+	Round(nCheques,2)	+	Round(nCartaoC,2)	+	Round(nCartaoD,2)	+;
			Round(nConveni,2)	+	Round(nVales,2)	+	Round(nCredito,2)	+	Round(nFinanc,2)	+;
			Round(nOutros,2)

dbSelectArea("SL2")
dbSetOrder(1)  
dbSeek(xFilial("SL2") + nOrcam)

While !SL2->(Eof()) .AND. SL2->L2_FILIAL + SL2->L2_NUM == xFilial("SL2") + nOrcam
      /* Entrega ou Retira Posterior ou Vale Presente*/	
	If SL2->L2_ENTREGA $("1|2|3|4|5") .OR. !Empty(SL2->L2_VALEPRE) ;
	       .OR. (Posicione("SB1",1,xFilial("SB1") + SL2->L2_PRODUTO,"SB1->B1_TIPO") == cServType) ;
	       .OR. (Posicione("SB1",1,xFilial("SB1") + SL2->L2_PRODUTO,"SB1->B1_TIPO") == cGarType)
		
		If SL2->L2_ENTREGA $("3|5")
			lPedido := .T.
		EndIf

		//Faz o tratamento do valor do ICMS ret.
		If SL2->(FieldPos("L2_ICMSRET")) > 0
			nVlrIcmsRet	:= SL2->L2_ICMSRET
		Endif

		cQuant 		:= StrZero(SL2->L2_QUANT, 8, 3)

		If (!lMvGarFP .AND. !lLibQtdGE ) .AND. (cGarantia == "*" .OR. (Posicione("SB1",1,xFilial("SB1")+SL2->L2_PRODUTO, "B1_TIPO") == SuperGetMV("MV_LJTPGAR",,"GE")))
			cVrUnit		:= Str(((SL2->L2_QUANT * SL2->L2_VLRITEM) + SL2->L2_VALIPI + nVlrIcmsRet) / SL2->L2_QUANT, 15, 2)
		Else
			cVrUnit		:= Str(((SL2->L2_QUANT * SL2->L2_PRCTAB) + SL2->L2_VALIPI + nVlrIcmsRet) / SL2->L2_QUANT, 15, 2)
		EndIf

		//Valor de desconto no item
		nVlrDescIt += SL2->L2_VALDESC
		nTotDesc   += SL2->L2_DESCPRO
		cVlrItem   := Str(Val(cVrUnit) * SL2->L2_QUANT, 15, 2)

		sTexto		:= sTexto + SL2->L2_PRODUTO + SL2->L2_DESCRI + Chr(13) + Chr(10)
		sTexto		:= sTexto + cQuant + '  ' + cVrUnit + '      ' + cVlrItem + Chr(13) + Chr(10)
		If SL2->L2_VALDESC > 0
			sTexto	:= sTexto + 'Desconto no Item:              ' + Str(SL2->L2_VALDESC, 15, 2) + Chr(13) + Chr(10)
		EndIf
		
		nValTot  += Val(cVlrItem)
 
		SL2->(DbSkip())
	Else
		SL2->(DbSkip())
	EndIf
EndDo                    

cDesconto	:= Str(nVlrDescIt, TamSx3("L2_VALDESC")[1], TamSx3("L2_VALDESC")[2])
nVlrFSD		:= SL1->L1_FRETE + SL1->L1_SEGURO + SL1->L1_DESPESA

If SL1->L1_DESCONTO > 0
	//O valor de desconto deve ser encontrada atraves da soma de todos os produtos (seus valores originais) (sem frete / sem desconto / sem acrescimo)
	//Porem quando é selecionado a NCC para pagamento os valores vem um pouco diferentes.
	If SL1->L1_CREDITO > 0 
		nDescTot	:= SL1->L1_DESCONT * (  (nValTot-nVlrDescIt) / (SL1->L1_VALBRUT - nVlrFSD + SL1->L1_DESCONT))
	Else
		nDescTot	:= nTotDesc
	EndIf 
	
	sTexto	:= sTexto + 'Desconto no Total:             ' + Str(nDescTot, 15, 2) + Chr(13) + Chr(10)
EndIf

//Armazena Valor Total
If lMvArrefat
    nVlrTot := Round((nValTot - nDescTot - nVlrDescIt + nTroco), TamSX3("D2_TOTAL")[2])
Else
    nVlrTot := NoRound((nValTot - nDescTot - nVlrDescIt + nTroco), TamSX3("D2_TOTAL")[2])
EndIf

//Calcula juros
If SL1->L1_VLRJUR > 0    

   nVlrAcres := SL1->L1_VLRJUR  
    
	nVlrTot   += nVlrAcres //Adiciona acrescimo no valor total
    sTexto    := sTexto + 'Valor do acrescimo no Total:     ' + Transform(SL1->L1_VLRJUR, "@R 99,999,999.99") + Chr(13) + Chr(10)

EndIf

//Adiciona frete somente quando existe um item com pedido na venda
If nVlrFSD > 0 .And. lPedido
    nVlrTot += nVlrFSD
EndIf

/* Ajusta o valor proporcionalizado na condição de pagamento em $
Necessario para evitar diferença de 0,01 centavos em determinados casos de venda mista*/
If nDinheir > 0
	If nValPag <> nVlrTot
		// Ajusto o valor em dinheiro para impressão no comprovante não fiscal		
		//nDinheir := nDinheir + Round(nValTot + nVlrFSD - nDescTot - nVlrDescIt + nTroco + nVlrAcres,2) - nValPag
		nDinheir := nDinheir + Round(nValTot + nVlrFSD - nDescTot - nVlrDescIt + nTroco, 2) - nValPag

	EndIf	            
EndIf
                                                                    
If nVlrFSD > 0 .And. lPedido
	sTexto	:= sTexto + 'Frete:                         ' + Transform(nVlrFSD, PesqPict("SL1","L1_FRETE")) + Chr(13) + Chr(10)
EndIf

sTexto	:= sTexto + '-----------------------------------------------' + Chr(13) + Chr(10)
sTexto	:= sTexto + 'TOTAL                          ' + Str(nVlrTot, 15, 2) + Chr(13) + Chr(10)
If nDinheir > 0 
	sTexto := sTexto + 'DINHEIRO' + '                   ' + Str( nDinheir , 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf
If nCheques > 0 
	sTexto := sTexto + 'CHEQUE' + '                     ' + Str(nCheques, 15, 2) + ' (+)' +  Chr(13) + Chr(10)
EndIf
If nCartaoC > 0 
	sTexto := sTexto + 'CARTAO CRED' + '                ' + Str(nCartaoC, 15, 2) + ' (+)' +  Chr(13) + Chr(10)
EndIf
If nCartaoD > 0 
	sTexto := sTexto + 'CARTAO DEB' + '                 ' + Str(nCartaoD, 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf

If nPIX > 0 
	sTexto := sTexto + 'PIX' + '                        ' + Str(nCartaoD, 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf
If nCartDig > 0 
	sTexto := sTexto + 'CARTEIRA DIGITAL' + '           ' + Str(nCartaoD, 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf

If nConveni > 0 
	sTexto := sTexto + 'CONVENIO' + '                   ' + Str(nConveni, 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf
If nOutros > 0 
	sTexto := sTexto + 'OUTROS' + '                      ' + Str(nOutros, 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf
If nVales > 0 
	sTexto := sTexto + 'VALES' + '                      ' + Str(nVales, 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf
If nFinanc > 0 
	sTexto := sTexto + 'FINANCIADO' + '                 ' + Str(nFinanc, 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf  
If nCredito > 0
	sTexto := sTexto + 'CREDITO ' + '                   ' + Str(nCredito, 15, 2) + ' (+)' + Chr(13) + Chr(10)
EndIf			
sTexto := sTexto + '-----------------------------------------------' + Chr(13) + Chr(10) 
If lMvLjTroco .And. nTroco > 0
	sTexto := sTexto + 'TROCO   ' + '                   ' + Str(nTroco, 15, 2) +' (-)'+ Chr(13) + Chr(10)
EndIf			                                                                                        
sTexto := sTexto + '-----------------------------------------------' + Chr(13) + Chr(10) 

If !Empty(cDocPed) .and. !Empty(cSerPed) 
    sTexto := sTexto + 'DOCUMENTO: ' + cDocPed + Chr(13) + Chr(10)
	sTexto := sTexto + 'SERIE: ' + cSerPed + Chr(13) + Chr(10)
EndIf

sTexto := sTexto + Replicate("-", nMaxChar)						   + Chr(13)+ Chr(10) 

sTexto := sTexto + '<b>Orcamento: </b>' + AllTrim(SL1->L1_NUM) + Chr(13) + Chr(10)
sTexto := sTexto + ' ' + Chr(13) + Chr(10)
sTexto := sTexto + '<b>Cliente:</b> ' +  AllTrim(SL1->L1_CLIENTE) + "-" + Alltrim(cNomCli)    + Chr(13) + Chr(10)
sTexto := sTexto + Replicate("-", nMaxChar)						     + Chr(13) + Chr(10) 

sTexto := sTexto + '<b>Data:</b> ' + DtoC(dDatabase) + ' <b>Hora: </b>' +Time() + Chr(13) + Chr(10)
sTexto := sTexto + '<b>Vendedor:</b> ' + Alltrim(SL1->L1_VEND)+' - ' +  Alltrim(cNomVen) + Chr(13) + Chr(10)
sTexto := sTexto + '<b>Caixa:</b> ' + Alltrim(SL1->L1_ESTACAO)+'<b> Operador: </b>' + Alltrim(SL1->L1_OPERADO)+' - ' +  Alltrim(cNomOpe) + Chr(13) + Chr(10)
sTexto := sTexto + Replicate("-", nMaxChar)						     + Chr(13) + Chr(10)
sTexto := sTexto + ' ' + Chr(13) + Chr(10)

sTexto := sTexto + Chr(13) + Chr(10)

STWManagReportPrint(sTexto,1) //Envia comando para a Impressora

RestArea(aAreaSL1)
RestArea(aAreaSL2)  
RestArea(aArea)

Return

/*/{Protheus.doc} SCRPRetPgt
Retorna os valores de cada Forma de Pagamento da venda conforma os valores gravados na SL4
@type  Static Function
@author joao.marcos
@since 26/09/2023
@version version
@return aVlrFormas, arrray, array com os valores de cada Forma de Pagamento
/*/
Static Function SCRPRetPgt()
Local aAreaSL4		:= SL4->(GetArea())
Local aVlrFormas	:= {{"R$",0},;	// 01
						{"CH",0},;	// 02
						{"CC",0},;	// 03
						{"CD",0},;	// 04
						{"PX",0},;	// 05
						{"PD",0},;	// 06
						{"FI",0},;	// 07
						{"CO",0},;	// 08
						{"VA",0},;	// 09
						{"CR",0},;	// 10
						{"OUTRO",0}} // 11

SL4->(dbSetOrder(1))
If SL4->(dbSeek(SL1->L1_FILIAL + SL1->L1_NUM))
	While SL4->(!EOF()) .AND. SL4->L4_FILIAL == SL1->L1_FILIAL .AND. SL4->L4_NUM == SL1->L1_NUM
		Do Case
			Case AllTrim(SL4->L4_FORMA) == "R$"
				aVlrFormas[01][02] += SL4->L4_VALOR
			Case AllTrim(SL4->L4_FORMA) == "CH"
				aVlrFormas[02][02] += SL4->L4_VALOR
			Case AllTrim(SL4->L4_FORMA) == "CC"
				aVlrFormas[03][02] += SL4->L4_VALOR
			Case AllTrim(SL4->L4_FORMA) == "CD"
				aVlrFormas[04][02] += SL4->L4_VALOR
			Case AllTrim(SL4->L4_FORMA) == "PX"
				aVlrFormas[05][02] += SL4->L4_VALOR
			Case AllTrim(SL4->L4_FORMA) == "PD"
				aVlrFormas[06][02] += SL4->L4_VALOR
			Case AllTrim(SL4->L4_FORMA) == "FI"
				aVlrFormas[07][02] += SL4->L4_VALOR
			Case AllTrim(SL4->L4_FORMA) == "CO"
				aVlrFormas[08][02] += SL4->L4_VALOR
			Case AllTrim(SL4->L4_FORMA) == "VA"
				aVlrFormas[09][02] += SL4->L4_VALOR	
			Case AllTrim(SL4->L4_FORMA) == "CR"
				aVlrFormas[10][02] += SL4->L4_VALOR
			Otherwise
				aVlrFormas[11][02] += SL4->L4_VALOR	
		EndCase

		SL4->(dbSkip())
	EndDo
EndIf

RestArea(aAreaSL4)
	
Return aVlrFormas
