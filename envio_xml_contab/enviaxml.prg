#Include "hwgui.ch"
STATIC Thisform

FUNCTION envioprgcontab(lOpt)
  Local hWin:= GetActiveWindow();
		  , vDataUltimoEnvio, vDataAtual		//adicionadas

  PRIVATE vDate1 := date() , vDate2 := date() , vEmail := cdpar000->ctb_email, vZip:= .F. ;
  			 , dDataInicialProximoEnvio, dDataFinalProximoEnvio, lAutoEnvio		//adicionadas

  //Adicionada verificacao para realizar procedimento automatico  
  if pCount() < 1 .OR. !HB_IsLogical(pValue(1))
	  lAutoEnvio := .F.
  else
	  lAutoEnvio := lOpt
  endif

IF lAutoEnvio		//Rotina de envio automatico

	vDataUltimoEnvio := cdpar000->xml_auto_envio
	
	vDataAtual := DtoS( Date() )
	
	if Empty(vDataUltimoEnvio)
		vDataUltimoEnvio := DtoS( ( StoD(vDataAtual) - Val( Right(vDataAtual,2) ) ) )
	endif
	
	//Verificar se deve ser executada
	if !( Val(Left(vDataAtual,6)) > Val(Left(vDataUltimoEnvio,6)) )		//se continuar no mesmo mês, retorna.
	
		release vDate1, vDate2, vEmail, vZip, dDataInicialProximoEnvio, dDataFinalProximoEnvio	

		Return NIL
	
	endif
	
	Private Notas := {}, nRadiogroup_TipoDoc, cMsgFalha := '', nEtapaEnvio := 0, aArqZip := {}, cPathArqZip
	
	dDataInicialProximoEnvio := BoM( StoD(vDataUltimoEnvio) ) //StoD( Left(vDataUltimoEnvio,6) + '01' )

  	dDataFinalProximoEnvio := StoD(vDataAtual) - Val( Right(vDataAtual,2) )
	
	vZip := .T.
	
	vDate1 := dDataInicialProximoEnvio
	vDate2 := dDataFinalProximoEnvio
	 
	if Coleta_Envia() //caso de tudo certo
		cdpar000->xml_auto_envio := vDataAtual
		MsgInfo('Os Arquivos Fiscais foram enviados, para o email ' + Alltrim(vEmail) + ', com sucesso.','Aviso do Sistema')
	else		
		if !Empty(cMsgFalha)
			MsgInfo(cMsgFalha,'Aviso do Sistema')
		endif
	endif

	Release vDate1, vDate2, vEmail, vZip, dDataInicialProximoEnvio, dDataFinalProximoEnvio ;
			  , Notas, nRadiogroup_TipoDoc, lAutoEnvio, cMsgFalha, nEtapaEnvio, aArqZip, cPathArqZip
	
ELSE	//Rotina manual, como já funcionava

  PRIVATE oDlgEnviaXML, oDate1, oLabel1, oDate2, oLabel2, oButtonex1, oEmail, oLabel3 ;
        , oBrowse1EnviaXML, oButtonex2, oButtonex3, notas:= {}, oZip, oAnimation1
  //variaveis adicionadas
  Private nRadiogroup_TipoDoc := 1
  Private oRadiogroup_TipoDoc, oRadiobutton_NFE, oRadiobutton_NFCE ;
        , oRadiobutton_SAT, oRadiobutton_CTE, oRadiobutton_MDFE


  INIT DIALOG oDlgEnviaXML TITLE "Envio de arquivos fiscais" ; //"Notas fiscais eletrônicas - Envio de XML p/ contabilidade" ;  //Titulo alterado
    ICON HIcon():AddResource(100)  ;
    AT 0, 0 SIZE 884,590 ;  //AT 0, 0 SIZE 884,467 ; (valor anterior)
    FONT HFont():Add( 'Tahoma',0,-11,400,,,) CLIPPER  NOEXIT  ;
    STYLE WS_POPUP+WS_CAPTION+WS_SYSMENU+DS_CENTER ;  
    ON INIT {|| SetTransparent( hWin, 200 ) } 
    Thisform:= oDlgEnviaXML

	//Adicionado grupo para escolha do tipo de doc fiscal a ser enviado
   @ 7,10 GET RADIOGROUP oRadiogroup_TipoDoc VAR nRadiogroup_TipoDoc  ;
         CAPTION "Selecione o tipo de documento fiscal a ser enviado"  SIZE 280,105 ;
        STYLE BS_LEFT  
   @ 30,25 RADIOBUTTON oRadiobutton_NFE CAPTION "NFe"  SIZE 90,22  
   @ 30,55 RADIOBUTTON oRadiobutton_NFCE CAPTION "NFCe"  SIZE 90,22  
   @ 30,85 RADIOBUTTON oRadiobutton_SAT CAPTION "SAT"  SIZE 90,22  
   @ 180,24 RADIOBUTTON oRadiobutton_CTE CAPTION "CTe"  SIZE 90,22  
   @ 180,55 RADIOBUTTON oRadiobutton_MDFE CAPTION "MDFe"  SIZE 90,22  
   END RADIOGROUP oRadiogroup_TipoDoc SELECTED nRadiogroup_TipoDoc
    
	@ 417,208 ANIMATION oAnimation1 SIZE 50,50   OF oDlgEnviaXML ;
        FROM RESOURCE 10001 ;
        AUTOPLAY
        
   //Bloco abaixo foi movido 120px à baixo, para acomodar o groupbox
   @ 7,141 GET DATEPICKER oDate1 VAR vDate1 SIZE 98,32  
   @ 7,125 SAY oLabel1 CAPTION "Data inicial"  TRANSPARENT SIZE 55,21  
   @ 123,141 GET DATEPICKER oDate2 VAR vDate2 SIZE 98,32  
   @ 123,125 SAY oLabel2 CAPTION "Data final"  TRANSPARENT SIZE 46,21  
   @ 228,141 BUTTONEX oButtonex1 CAPTION ""   TRANSPARENT SIZE 44,32 ;
        STYLE BS_CENTER +WS_TABSTOP  ;
        BITMAP (HBitmap():AddFile('res\eanok.bmp')):handle  ;
        BSTYLE 0 ;
        ON CLICK {|| query_xmls(), oButtonex3:SetFocus() }   //adicionada mudança de foco para o botao marcar tudo
   @ 303,141 GET oEmail VAR vEmail SIZE 377,32  MAXLENGTH 60  
   @ 303,125 SAY oLabel3 CAPTION "Email contab."  TRANSPARENT SIZE 64,21  
   @ 8,181 BROWSE oBrowse1EnviaXML ARRAY SIZE 869,354 ;
        STYLE WS_TABSTOP       NOBORDER 

    // CREATE oBrowse1EnviaXML   //  SCRIPT GENARATE BY DESIGNER
    oBrowse1EnviaXML:lSep3d := .T.
    oBrowse1EnviaXML:aArray := notas
    oBrowse1EnviaXML:ShowMark := .F.
    oBrowse1EnviaXML:AddColumn( HColumn():New('Emissão', ColumnArBlock() ,'D',10, 0 ,.F.,1,1,,,,,,,,,))
    oBrowse1EnviaXML:AddColumn( HColumn():New('Número', ColumnArBlock() ,'N',9, 0 ,.F.,1,1,,,,,,,,,))
    oBrowse1EnviaXML:AddColumn( HColumn():New('Destinatário', ColumnArBlock() ,'C',30, 0 ,.F.,1,,,,,,,,,,))
    oBrowse1EnviaXML:AddColumn( HColumn():New('XML', ColumnArBlock() ,'C',65, 0 ,.T.,1,1,,,,,,,,, {|value,this| pega_path_xml(value,this) }))
    oBrowse1EnviaXML:AddColumn( HColumn():New('Enviado', ColumnArBlock() ,'D',10, 0 ,.F.,1,1,,,,,,,,,))
    oBrowse1EnviaXML:AddColumn( HColumn():New('Enviar', ColumnArBlock() ,'L',2, 0 ,.T.,1,1,,,,,,,,,))

    // END BROWSE SCRIPT  -  oBrowse1EnviaXML
   @ 756,546 BUTTONEX oButtonex2 CAPTION "Envia para contab."   SIZE 121,32 ;
        STYLE BS_CENTER +WS_TABSTOP  ;
        ON CLICK {|| oButtonEx2:Disable(),enviaparacontab(),oButtonEx2:Enable()}

   @ 628,546 BUTTONEX oButtonex3 CAPTION "Marcar todas"   SIZE 121,32 ;
        STYLE BS_CENTER +WS_TABSTOP ;
        ON CLICK {|| marcatudo() }

   @ 723,151 GET CHECKBOX oZip VAR vZip  ;
        CAPTION "Compactar"  SIZE 110,22  ; 
        TOOLTIP 'Reúne todos os arquivos num arquivo compactado, formato ZIP.'
        
   ACTIVATE DIALOG oDlgEnviaXML ;
   on activate {|| oAnimation1:Hide() }    
        
	rele vDate1, vDate2, vEmail, oDlgEnviaXML,  oDate1, oLabel1, oDate2, oLabel2, oButtonex1, oEmail, oLabel3 ;
        , oBrowse1EnviaXML, oButtonex2, oButtonex3, notas, vZip, oZip, oAnimation1 ;
		  , nRadiogroup_TipoDoc, oRadiogroup_TipoDoc, oRadiobutton_NFE, oRadiobutton_NFCE ;	//adicionadas
        , oRadiobutton_SAT, oRadiobutton_CTE, oRadiobutton_MDFE ;
		  , dDataInicialProximoEnvio, dDataFinalProximoEnvio		  
		  
   SetTransparent( hWin, 255 ) 

ENDIF
   
RETURN NIL


Static Function marcatudo
local a, b:= oBrowse1EnviaXML:nCurrent
if len(notas) = 0
   return nil
endif
oBrowse1EnviaXML:Top()
for a:= 1 to len(notas)
    oBrowse1EnviaXML:aArray[a,6]:= .t.
    oBrowse1EnviaXML:refreshline(.t.)
    oBrowse1EnviaXML:LineDown()
next
oBrowse1EnviaXML:nCurrent = b
oBrowse1EnviaXML:Refresh()	//adicionado refresh no browser
return nil


Static Function pega_path_xml(value,this)   
Local cSALVA_PATH:= CAMINHO_EXE(), cArq:= SELECTFILE("Arquivos Nota Fiscal Eletrônica", "*.xml",Substr( HB_ArgV(0), 1, ( Rat('\',HB_ArgV(0)) - 1 ) )+'\nfe\criadas','Selecione o arquivo XML correspondente')	//substituido curdrive/curdir SELECTFILE("Arquivos Nota Fiscal Eletrônica", "*.xml",curdrive()+':\'+curdir()+'\nfe\criadas','Selecione o arquivo XML correspondente')
Local lEnv:= .f.
this:value:= cArq
oBrowse1EnviaXML:refreshline(.t.)
Dirchange(cSALVA_PATH)
RETURN .t.


Static Function query_xmls
Local oSql:= SR_GetConnection(), cSql, n:= {}, hArqNFe ;
    , rede:= iif(file("c:\acbrmonitorplus\acbrmonitor.exe"),.f.,.t.), patth ; //alterado patth:= curdrive()+':\'+curdir()+'\nfe\criadas\'
    , i, path_exe := Left(HB_ArgV(0),Rat('\',HB_ArgV(0))), hArqAux ;	//adicionadas
    , path_sat_aut, path_sat_tra, path_sat_can, path_sat_can_tra ;
	 , path_nfce_aut, path_nfce_tra, path_mdfe_aut, path_mdfe_can ;
	 , path_cte_cri, lFalta_cte := .F.
notas = {}
//Bloco adicionado para identificar caminho das pastas
Do Case
	Case nRadiogroup_TipoDoc == 1		//NFe    
		patth := path_exe + 'nfe\criadas\'
		
	Case nRadiogroup_TipoDoc == 2		//NFCe
		//equivale a patth
		path_nfce_aut := path_exe + 'nfce\autorizados\'

		path_nfce_tra := path_exe + 'nfce\autorizados\transmitidos\'
		
	Case nRadiogroup_TipoDoc == 3		//SAT
		//equivale a patth
		path_sat_aut := path_exe + 'sat\autorizados\'

		path_sat_tra := path_exe + 'sat\autorizados\transmitidos\'
		
		//equivale a pattz
		path_sat_can := path_exe + 'sat\cancelados\'

		path_sat_can_tra := path_exe + 'sat\cancelados\transmitidos\'

	Case nRadiogroup_TipoDoc == 4		//CTe
		//equivale a patth
		 path_cte_cri := path_exe + 'nfe\criadas\'
		 
	Case nRadiogroup_TipoDoc == 5		//MDFe
		//equivale a patth
		path_mdfe_aut :=  path_exe + 'mdfe\autorizados\'

EndCase

/*	bloco movido para o case abaixo (pode ser excluido)
cSql:= "select distinct cab_dt,cabnfe,cabnom,'',xmlctbdt,false,cabkey,sr_recno,cabarqxml from cdcab000 where cab_dt between "
cSql += sr_cdbvalue(vDate1) + " and " + sr_cdbvalue(vDate2)
oSql:Exec(cSql,.t.,.t.,@notas)
for a:= 1 to len(notas)
    notas[a,6] = .f.
    if file(patth + notas[a,7] + '-nfe.xml') // arquivo existe na pasta
       notas[a,4] = patth + notas[a,7] + '-nfe.xml'
    else                                     // arquivo não existe na pasta e será criado com base no BD
       hArqNFe:= Fcreate(patth + notas[a,7] + '-nfe.xml',0)
       Fwrite(hArqNFe,notas[a,9])
       Fclose(hArqNFe)
       notas[a,4] = patth + notas[a,7] + '-nfe.xml'
    endif   
next
*/
//Adicionada verificacao do tipo de doc selecionado
Do Case	
	Case nRadiogroup_TipoDoc == 1		//NFe    
		cSql:= "select distinct cab_dt,cabnfe,cabnom,'',xmlctbdt,false,cabkey,sr_recno,cabarqxml from cdcab000 where cab_dt between "
		cSql += sr_cdbvalue(vDate1) + " and " + sr_cdbvalue(vDate2)
		
		oSql:Exec(cSql,.t.,.t.,@notas)
		
		for a:= 1 to len(notas)
    		notas[a,6] := lAutoEnvio //.f. valor anterior
		    if file(patth + notas[a,7] + '-nfe.xml') // arquivo existe na pasta
		       notas[a,4] := patth + notas[a,7] + '-nfe.xml'
		    else                                     // arquivo não existe na pasta e será criado com base no BD
		       hArqNFe:= Fcreate(patth + notas[a,7] + '-nfe.xml',0)
		       Fwrite(hArqNFe,notas[a,9])
		       Fclose(hArqNFe)
		       notas[a,4] := patth + notas[a,7] + '-nfe.xml'
    		endif   
		next
	Case nRadiogroup_TipoDoc == 2		//NFCe
									//	1					2		3		4		5		6			7				8			9				10						
		cSql:= "select distinct usu_nfce_data,usufis,usunom,'',xmlctbdt,false,usu_nfce_key,sr_recno,usu_nfce_xml,usu_nfce_canc_xml from cdusu000 where usu_nfce_data between "
		cSql += sr_cdbvalue(vDate1) + " and " + sr_cdbvalue(vDate2) + " order by usufis"
		
		oSql:Exec(cSql,.t.,.t.,@notas)

		if len(notas) > 0
			for i := 1 to len(notas)
				notas[i,6] := lAutoEnvio //.f. valor anterior
				//verifica aqui se houve o cancelamento do cupom, se o arquivo existe e cria se for o caso
				if File( path_nfce_tra + notas[i,7] + '-NFCe.xml' )
					notas[i,4] := path_nfce_tra + notas[i,7] + '-NFCe.xml'
				elseif File( path_nfce_tra + rAnoMes(notas[i,7],3) + '\' + notas[i,7] + '-NFCe.xml' )
					notas[i,4] := path_nfce_tra + rAnoMes(notas[i,7],3) + '\' + notas[i,7] + '-NFCe.xml'
				elseif File( path_nfce_aut + notas[i,7] + '-NFCe.xml' )
					notas[i,4] := path_nfce_aut + notas[i,7] + '-NFCe.xml'
				else
				   hArqAux := FCreate(path_nfce_aut + notas[i,7] + '-NFCe.xml',0)
					if !Empty(notas[i,10])
					   FWrite(hArqAux,notas[i,10])	//coloca o conteudo do usu_nfce_canc_xml
		   		else
					   FWrite(hArqAux,notas[i,9])		//coloca o conteudo do usu_nfce_xml
		   		endif
				   FClose(hArqAux)
					notas[i,4] := path_nfce_aut + notas[i,7] + '-NFCe.xml'
				endif
			next
		endif
		
	Case nRadiogroup_TipoDoc == 3		//SAT  
									//	1					2		3		4		5		6			7				8			9				10						11
		cSql:= "select distinct usu_sat_data,usufis,usunom,'',xmlctbdt,false,usu_sat_key,sr_recno,usu_sat_xml,usu_sat_canc_key,usu_sat_canc_xml from cdusu000 where usu_sat_data between "
		cSql += sr_cdbvalue(vDate1) + " and " + sr_cdbvalue(vDate2) + " order by usufis"
		
		oSql:Exec(cSql,.t.,.t.,@notas)

		if len(notas) > 0
			for i := 1 to len(notas)
				notas[i,6] := lAutoEnvio //.f. valor anterior
				//verifica primeiro em transmitidos
				if File( path_sat_tra + notas[i,7] + '.xml' )
					notas[i,4] := path_sat_tra + notas[i,7] + '.xml'
				elseif File( path_sat_tra + rAnoMes(notas[i,7],6) + '\' + notas[i,7] + '.xml' )
					notas[i,4] := path_sat_tra + rAnoMes(notas[i,7],6) + '\' + notas[i,7] + '.xml'
				elseif File( path_sat_aut + notas[i,7] + '.xml' )
					notas[i,4] := path_sat_aut + notas[i,7] + '.xml'
				else
				   hArqAux := FCreate(path_sat_aut + notas[i,7] + '.xml',0)
				   FWrite(hArqAux,notas[i,9])
				   FClose(hArqAux)
					notas[i,4] := path_sat_aut + notas[i,7] + '.xml'
				endif
				
				//verifica aqui se houve o cancelamento do cupom
				if !Empty(notas[i,10]) .AND. !Empty(notas[i,11]) 	
					//procurar arquivo na pasta de cancelados e criar caso nao encontre
					if !File( path_sat_can_tra + notas[i,10] + '.xml' ) .AND. !File( path_sat_can_tra + rAnoMes(notas[i,10],6) + '\' + notas[i,10] + '.xml' ) .AND. !File( path_sat_can + notas[i,10] + '.xml' )	//arquivo nao existe na pasta cancelados nem trasmitidas
					   hArqAux := FCreate(path_sat_can + notas[i,10] + '.xml',0)
					   FWrite(hArqAux,notas[i,11])
					   FClose(hArqAux)
					endif
				endif
			next
		endif
		
	Case nRadiogroup_TipoDoc == 4		//CTe
									//	1					2			3		 4		5			6			7			8
		cSql:= "select distinct cteautoriz,ctenumero,ctedesnome,'',xmlctbdt,false,ctechave,sr_recno from cdcte000 where cteautoriz between "
		cSql += sr_cdbvalue(vDate1) + " and " + sr_cdbvalue(vDate2) + " order by ctenumero"
		
		oSql:Exec(cSql,.t.,.t.,@notas)

		if len(notas) > 0
			for i := 1 to len(notas)
				notas[i,2] := Val( CStr(notas[i,2]) )
				notas[i,6] := lAutoEnvio //.f. valor anterior
				if File( path_cte_cri + notas[i,7] + '-cte.xml' )
					notas[i,4] := path_cte_cri + notas[i,7] + '-cte.xml'
				else
					lFalta_cte := .T.
				endif
			next
		endif
		
	Case nRadiogroup_TipoDoc == 5		//MDFe
									//	1			2		3	4		5		6			7			8		9
		cSql:= "select distinct mdfdat,mdfnum,'','',xmlctbdt,false,mdfkey,sr_recno,mdfxml from cdmdf000 where mdfdat between "
		cSql += sr_cdbvalue(vDate1) + " and " + sr_cdbvalue(vDate2) + " order by mdfnum"
		
		oSql:Exec(cSql,.t.,.t.,@notas)
		
		if len(notas) > 0
			For i := 1 to len(notas)
				notas[i,2] := Val( CStr(notas[i,2]) )
				notas[i,6] := lAutoEnvio //.f. valor anterior
				//verifica se arquivo existe na pasta autorizados (cancelados serão incluidos na hora do envio como nfe)
				if file(path_mdfe_aut + notas[i,7] + '-mdf.xml') // arquivo existe na pasta
					notas[i,4] := path_mdfe_aut + notas[i,7] + '-mdf.xml'
				else                                     // arquivo não existe na pasta e será criado com base no BD
					hArqAux:= Fcreate(path_mdfe_aut + notas[i,7] + '-mdf.xml',0)
					Fwrite(hArqAux,notas[i,9])
					Fclose(hArqAux)
					notas[i,4] := path_mdfe_aut + notas[i,7] + '-mdf.xml'
				endif   
			Next
		endif
		
EndCase

if !lAutoEnvio		//adicionada verificacao de auto envio
	CreateArList(oBrowse1EnviaXML,notas)
	if lFalta_cte
		MsgInfo('Não foi possível localizar 1 ou mais arquivos CTe na pasta ' + path_cte_cri + Chr(10) + Chr(10) + 'Mova os arquivos necessarios para a pasta ou indique o caminho do mesmo com um duplo clique no campo xml.' + Chr(10) + Chr(10) + 'Em caso de dúvidas, favor entrar em contato com o suporte.' )
	endif
	Return nil
else
	if lFalta_cte
		cMsgFalha := 'Não foi possível enviar os arquivos fiscais para o email ' + Alltrim(vEmail) + ', pois o sistema não localizou 1 ou mais arquivos CTe.' + Chr(10) + Chr(10) + 'Utilize a opção Fiscal->Envio XMLs Contabilidade para realizar o envio manual.' + Chr(10) + Chr(10) + 'Em caso de dúvidas, favor entrar em contato com o suporte.'
		Return .T. //falhou
	endif
	Return .F.
endif


Static Function enviaparacontab                   //Aqui verificar e alterar caminho de acordo com tipo de doc
/* Alterado
Local a, vai:= .f., xmls:= {}, b, patth:= curdrive()+':\'+curdir()+'\nfe\criadas\'; 
    , pattz:= curdrive()+':\'+curdir()+'\nfe\canceladas\', cZip:= '', cSql, oSql:= SR_GetConnection()
*/
Local i, a, vai:= .f., xmls:= {}, b, patth; 
    , pattz, cZip:= '', cSql, oSql:= SR_GetConnection();
    , path_exe := Left(HB_ArgV(0),Rat('\',HB_ArgV(0))) ; //adicionadas
    , path_sat_aut, path_sat_tra, path_sat_can, path_sat_can_tra ;
	 , path_nfce_aut, path_nfce_tra, path_mdfe_aut, path_mdfe_can, path_cte_cri
    
//Bloco adicionado para estrutura de pastas    
Do Case
	Case nRadiogroup_TipoDoc == 1		//NFe
		patth := path_exe + 'nfe\criadas\'		
		pattz := path_exe + 'nfe\canceladas\'
		
	Case nRadiogroup_TipoDoc == 2		//NFCe
		//equivale a patth
		path_nfce_aut := path_exe + 'nfce\autorizados\'

		path_nfce_tra := path_exe + 'nfce\autorizados\transmitidos\'
		
	Case nRadiogroup_TipoDoc == 3		//SAT
		//equivale a patth
		path_sat_aut := path_exe + 'sat\autorizados\'

		path_sat_tra := path_exe + 'sat\autorizados\transmitidos\'
		
		//equivale a pattz
		path_sat_can := path_exe + 'sat\cancelados\'

		path_sat_can_tra := path_exe + 'sat\cancelados\transmitidos\'

	Case nRadiogroup_TipoDoc == 4		//CTe
		//equivale a patth
		 path_cte_cri := path_exe + 'nfe\criadas\'
	
	Case nRadiogroup_TipoDoc == 5		//MDFe
		//equivale a patth
		path_mdfe_aut := path_exe + 'mdfe\autorizados\'

		//equivale a pattz
		path_mdfe_can := path_exe + 'mdfe\cancelados\'
		
EndCase

//adicionada verificacao de auto envio
if !lAutoEnvio
	for a:= 1 to len(notas)
   	 if notas[a,6]
      	 vai = .t.
	    endif
	next
	if empty(vEmail)
	   msgstop('Informe o email da contabilidade !')
	   return nil
	endif
	if !vai
   	msginfo('Nenhuma nota selecionada para envio !')
   	return nil
	endif
	oAnimation1:Show()
	MilliSec(250)
endif

/* Bloco Movido para o Case
for a:= 1 to len(notas)
    if notas[a,6]        //Aqui verificar o tipo de doc, estrutura de pastas e nomenclatura
       aadd(xmls,notas[a,4])
       if file(pattz + notas[a,7] + '-cancelamento.xml')
          aadd(xmls,pattz + notas[a,7] + '-cancelamento.xml')
       endif
    endif
next
*/
//Bloco adicionado
Do Case	
	Case nRadiogroup_TipoDoc == 1		//NFe
		for a:= 1 to len(notas)
			if notas[a,6]
				if File( notas[a,4] )
					aadd(xmls,notas[a,4])
				endif
				if file(pattz + notas[a,7] + '-cancelamento.xml')
					aadd(xmls,pattz + notas[a,7] + '-cancelamento.xml')
				endif
			endif
		next
	Case nRadiogroup_TipoDoc == 2		//NFCe
		For i := 1 to len(notas)
			if notas[i,6] .AND. File( notas[i,4] )
				aadd(xmls,notas[i,4])
			endif
		Next
	Case nRadiogroup_TipoDoc == 3		//SAT
		For i := 1 to len(notas)
			if notas[i,6]
				if File( notas[i,4] )
					aadd(xmls,notas[i,4])
				endif
				//verifica aqui se houve o cancelamento do cupom
				if !Empty(notas[i,10]) .AND. !Empty(notas[i,11]) 	
					if File( path_sat_can_tra + notas[i,10] + '.xml' )
						aadd(xmls, path_sat_can_tra + notas[i,10] + '.xml' )						
					elseif File( path_sat_can_tra + rAnoMes(notas[i,10],6) + '\' + notas[i,10] + '.xml' )
						aadd(xmls, path_sat_can_tra + rAnoMes(notas[i,10],6) + '\' + notas[i,10] + '.xml' )
					elseif File( path_sat_can + notas[i,10] + '.xml' )
						aadd(xmls, path_sat_can + notas[i,10] + '.xml' )
					endif
				endif
			endif
		Next
	Case nRadiogroup_TipoDoc == 4		//CTe
		For i := 1 to len(notas)
			if notas[i,6] .AND. File( notas[i,4] )
				aadd(xmls,notas[i,4])
			endif
		Next
	Case nRadiogroup_TipoDoc == 5		//MDFe
		For i := 1 to len(notas)
			if notas[i,6]
				if File( notas[i,4] )
					aadd(xmls,notas[i,4])
				endif
				//verifica se existe cancelamento
				if file(path_mdfe_can + notas[i,7] + '-mdf.xml')
					aadd(xmls,path_mdfe_can + notas[i,7] + '-mdf.xml')
				endif
			endif
		Next
EndCase

if lAutoEnvio
	if nEtapaEnvio < 6
		if len(xmls) <= 0
			cMsgFalha := 'Ocorreu uma falha inesperada e não foi possível anexar os arquivos fiscais ' + iif(nRadiogroup_TipoDoc==1,'NFe',iif(nRadiogroup_TipoDoc==2,'NFCe',iif(nRadiogroup_TipoDoc==3,'SAT',iif(nRadiogroup_TipoDoc==4,'CTe','MDFe')))) + ' ao email!' + Chr(10) + Chr(10) + 'Verifique o caminho dos arquivos e tente novamente; caso o problema persista, favor entrar em contato com o suporte.'
			Return .T.	//falhou
		endif	

		cPathArqZip := path_exe + "swap\" + cdpar000->cnpj + '_' + iif(nRadiogroup_TipoDoc==1,'NFe',iif(nRadiogroup_TipoDoc==2,'NFCe',iif(nRadiogroup_TipoDoc==3,'SAT',iif(nRadiogroup_TipoDoc==4,'CTe','MDFe')))) + '.zip'
		if !hb_zipfile(cPathArqZip,xmls,5,,.T.,,.F.,.F.,)
			cMsgFalha := 'Falha na compactação dos arquivos !' + Chr(10) + 'Verifique se a pasta ' + path_exe + 'swap existe.' + Chr(10) + 'Crie a pasta, caso não encontre; em caso de dúvidas, favor entrar em contato com o suporte.'
			Return .T.	//falhou
		else
			//adiciona o caminho do arquivo zip gerado para cada tipo de doc fiscal
			aAdd( aArqZip, cPathArqZip )			
		endif
		Return .F.
	endif
	xmls := aArqZip
else
	//adicionada verificação do array xmls
	if len(xmls) <= 0
   	msginfo('Ocorreu uma falha inesperada e não foi possível anexar os arquivos ao email!' + Chr(10) + Chr(10) + 'Verifique o caminho dos arquivos e tente novamente; caso o problema persista, favor entrar em contato com o suporte.')
   	return nil
	endif

	if vZip
   	*cZip:= curdrive()+':\'+curdir()+"\swap\"+cdpar000->cnpj+'.zip'	//alterada
   	cZip:= path_exe + "swap\" + cdpar000->cnpj + '_' + iif(nRadiogroup_TipoDoc==1,'NFe',iif(nRadiogroup_TipoDoc==2,'NFCe',iif(nRadiogroup_TipoDoc==3,'SAT',iif(nRadiogroup_TipoDoc==4,'CTe','MDFe')))) + '.zip'							
   	if !hb_zipfile(cZip,xmls,5,,.T.,,.F.,.F.,)
      	msgstop('Falha na compactação dos arquivos !' + Chr(10) + 'Verifique se a pasta ' + path_exe + 'swap existe.' + Chr(10) + 'Crie a pasta, caso não encontre; em caso de dúvidas, favor entrar em contato com o suporte.' )	//adicionada observacao
	      return nil
   	endif
	endif
	if vZip
   	xmls := {cZip}
	endif
endif

cSubject := 'Arquivos XML - ' + cdpar000->razao
aQuem    := alltrim(vEmail)
cMsg     := 'XML enviado pelo próprio software de emissão.' + chr(13)+chr(10) + 'Daxxi Tecnologia Ltda. (12) 3833-4366'
cServerIp:= alltrim(cdpar000->msmtp)
cFrom    := alltrim(cdpar000->mfrom)
cUser    := alltrim(cdpar000->muser)
cPass    := hb_decrypt(alltrim(cdpar000->mpswd),"senhadoconsultacpf.com")
vPORTSMTP:= val(alltrim(cdpar000->mport))
aCC      := "" // caracteres entre aspas
aBCC     := "" // caracteres entre aspas
lCONF    := .T.
lSSL     := cdpar000->mssl
lAut     := cdpar000->mauth
if config_mail(xmls,cSubject,aQuem,cMsg,cServerIp,cFrom,cUser,cPass,vPORTSMTP,aCC,aBCC,lCONF,lSSL,lAut)        
	/* Bloco movido para o case (pode ser excluido)
   for b:= 1 to len(notas)
       if notas[b,6]
       	//Aqui Acrescentar verificacao do tipo de doc para efetuar o update
          cSql:= "update cdcab000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
          cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
          oSql:Execute(cSql,.t.) 
          oBrowse1EnviaXML:aArray[b,5] = date()
          oBrowse1EnviaXML:RefreshLine(.t.)
       endif
   next    
   oSql:Commit()
	*/
	
	//Bloco adicionado	
	if lAutoEnvio	//modo de auto envio
		Return .F. 
		
	else	//modo manual
	
		Do Case	//Aqui Acrescentar verificacao do tipo de doc para efetuar o update
			Case nRadiogroup_TipoDoc == 1		//NFe
				for b:= 1 to len(notas)
					if notas[b,6] .AND. File( notas[b,4] )
						cSql:= "update cdcab000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
						cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
						oSql:Execute(cSql,.t.) 
						oBrowse1EnviaXML:aArray[b,5] = date()
						oBrowse1EnviaXML:RefreshLine(.t.)
					endif
				next    
				oSql:Commit()
      	
			Case nRadiogroup_TipoDoc == 2		//NFCe
				for b:= 1 to len(notas)
					if notas[b,6] .AND. File( notas[b,4] )
						cSql:= "update cdusu000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
						cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
						oSql:Execute(cSql,.t.) 
						oBrowse1EnviaXML:aArray[b,5] = date()
						oBrowse1EnviaXML:RefreshLine(.t.)
					endif
				next    
				oSql:Commit()

			Case nRadiogroup_TipoDoc == 3		//SAT
				for b:= 1 to len(notas)
					if notas[b,6] .AND. File( notas[b,4] )
						cSql:= "update cdusu000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
						cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
						oSql:Execute(cSql,.t.) 
						oBrowse1EnviaXML:aArray[b,5] = date()
						oBrowse1EnviaXML:RefreshLine(.t.)
					endif
				next    
				oSql:Commit()

			Case nRadiogroup_TipoDoc == 4		//CTe
				for b:= 1 to len(notas)
					if notas[b,6] .AND. File( notas[b,4] )
						cSql:= "update cdcte000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
						cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
						oSql:Execute(cSql,.t.) 
						oBrowse1EnviaXML:aArray[b,5] = date()
						oBrowse1EnviaXML:RefreshLine(.t.)
					endif
				next    
				oSql:Commit()

				/* Bloco substituido por update no cdcte000 (pode ser excluido)
				for b:= 1 to len(notas)
					if notas[b,6]
						oBrowse1EnviaXML:aArray[b,5] = date()
						oBrowse1EnviaXML:RefreshLine(.t.)
					endif
				next
				*/

			Case nRadiogroup_TipoDoc == 5		//MDFe
				for b:= 1 to len(notas)
					if notas[b,6] .AND. File( notas[b,4] )
						cSql:= "update cdmdf000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
						cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
						oSql:Execute(cSql,.t.) 
						oBrowse1EnviaXML:aArray[b,5] = date()
						oBrowse1EnviaXML:RefreshLine(.t.)
					endif
				next    
				oSql:Commit()

		EndCase
	
	endif

else	//adicionada verificacao de auto envio
 
 if lAutoEnvio
 	cMsgFalha := 'Ocorreu uma falha durante o envio do email.' + Chr(10) + Chr(10) + 'Verifique a configuração do email na opção Ferramentas->Email - Configuração' + Chr(10) + Chr(10) + 'Em caso de dúvidas, favor entrar em contato com o suporte.'
 	Return .T. //falhou
 endif
	 
endif
oBrowse1EnviaXML:Refresh()
oAnimation1:Hide()
return nil


//Funcao adicionada para efetuar a coleta e envio automatico dos arquivos fiscais
Static Function Coleta_Envia()
	Local i, b, lRet := .F., nTipo, lFalha := .F., aTodasNotas := {} ;
			, cSql, oSql:= SR_GetConnection();
			
	//consulta se existem arquivos a serem enviados
	For nTipo := 1 to 5
	
		nRadiogroup_TipoDoc := nTipo
		 
		lFalha := Query_Xmls()
		
		mostra(iif(nTipo==1,'NFe',iif(nTipo==2,'NFCe',iif(nTipo==3,'SAT',iif(nTipo==4,'CTe','MDFe')))),len(notas))
		
		if lFalha	//não localizou algum dos arquivos CTe
			aTodasNotas := {}
			Exit
		endif

		nEtapaEnvio := nTipo

		if Len(Notas) > 0
			//adiciona o array de notas para update na data de envio xmlctbdt
			aAdd( aTodasNotas, {notas,nTipo} )

			lFalha := EnviaParaContab()
		
			if lFalha	//não conseguiu compactar algum dos arquivos
				aTodasNotas := {}
				Exit
			endif
			
		endif
	
	Next
	
	if Len(aTodasNotas) > 0
		if !Empty(vEmail)
			MsgInfo('Existem arquivos fiscais prontos para enviar ao email do contabilista.' + Chr(10) + Chr(10) + 'O sistema irá realizar o envio para o email ' + Alltrim(vEmail),'Aviso do Sistema')
		else
			cMsgFalha := 'Existem arquivos fiscais prontos para envio, mas o email do contabilista não foi cadastrado.' + Chr(10) + Chr(10) + 'Efetue o cadastro do email do contabilista e reinicie o aplicativo.' + Chr(10) + Chr(10) + 'Em caso de dúvidas, favor entrar em contato com o suporte.'
			lFalha := .T.	//sem email para enviar
		endif
	endif
	
	if !lFalha .AND. Len(aTodasNotas) > 0
	
		nEtapaEnvio := 6
	
		lFalha := EnviaParaContab()
		
		if !lFalha	//conseguiu enviar o email
			lRet := .T.
			
			//realiza o update dos campos xmlctbdt e xmlctbhr dos arquivos fiscais enviados
			For i := 1 to Len(aTodasNotas)

				Do Case	//Aqui Acrescentar verificacao do tipo de doc para efetuar o update
					Case aTodasNotas[i,2] == 1		//NFe
						notas := aTodasNotas[i,1]
						for b:= 1 to len(notas)
							if notas[b,6] .AND. File( notas[b,4] )
								cSql:= "update cdcab000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
								cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
								oSql:Execute(cSql,.t.) 
							endif
						next    
						oSql:Commit()
      	
					Case aTodasNotas[i,2] == 2		//NFCe
						notas := aTodasNotas[i,1]
						for b:= 1 to len(notas)
							if notas[b,6] .AND. File( notas[b,4] )
								cSql:= "update cdusu000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
								cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
								oSql:Execute(cSql,.t.) 
							endif
						next    
						oSql:Commit()
   	
					Case aTodasNotas[i,2] == 3		//SAT
						notas := aTodasNotas[i,1]
						for b:= 1 to len(notas)
							if notas[b,6] .AND. File( notas[b,4] )
								cSql:= "update cdusu000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
								cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
								oSql:Execute(cSql,.t.) 
							endif
						next    
						oSql:Commit()
            	
					Case aTodasNotas[i,2] == 4		//CTe
						notas := aTodasNotas[i,1]
						for b:= 1 to len(notas)
							if notas[b,6] .AND. File( notas[b,4] )
								cSql:= "update cdcte000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
								cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
								oSql:Execute(cSql,.t.) 
							endif
						next    
						oSql:Commit()
         	
					Case aTodasNotas[i,2] == 5		//MDFe
						notas := aTodasNotas[i,1]
						for b:= 1 to len(notas)
							if notas[b,6] .AND. File( notas[b,4] )
								cSql:= "update cdmdf000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
								cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
								oSql:Execute(cSql,.t.) 
							endif
						next    
						oSql:Commit()
         	
				EndCase

			Next
			
		endif
			
	endif
	
Return lRet


//Funcao adicionada para recortar ano e mes da chave do doc fiscal
Static Function rAnoMes(cTxt,nPos)
Return Substr(cTxt,nPos,4)


*Static Function Caminho_Exe()	//REMOVER
*msginfo(Substr( HB_ArgV(0), 1, ( Rat('\',HB_ArgV(0)) - 1 ) ))
*Return Substr( HB_ArgV(0), 1, ( Rat('\',HB_ArgV(0)) - 1 ) ) //retornar o caminho do executavel sem a barra

*Static Function config_mail() //REMOVER
*Return .T. //procede a configuracao e envio do email (retorna como tendo funcionado)
