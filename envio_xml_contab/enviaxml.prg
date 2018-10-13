#Include "hwgui.ch"
STATIC Thisform

FUNCTION envioprgcontab
  Local hWin:= GetActiveWindow()
  PRIVATE vDate1 := date() , vDate2 := date() , vEmail := cdpar000->ctb_email, vZip:= .F.
  PRIVATE oDlgEnviaXML, oDate1, oLabel1, oDate2, oLabel2, oButtonex1, oEmail, oLabel3 ;
        , oBrowse1EnviaXML, oButtonex2, oButtonex3, notas:= {}, oZip, oAnimation1

  INIT DIALOG oDlgEnviaXML TITLE "Notas fiscais eletrônicas - Envio de XML p/ contabilidade" ;
    ICON HIcon():AddResource(100)  ;
    AT 0, 0 SIZE 884,467 ;
    FONT HFont():Add( 'Tahoma',0,-11,400,,,) CLIPPER  NOEXIT  ;
    STYLE WS_POPUP+WS_CAPTION+WS_SYSMENU+DS_CENTER ;
    ON INIT {|| SetTransparent( hWin, 200 ) }
    Thisform:= oDlgEnviaXML
    
   @ 417,208 ANIMATION oAnimation1 SIZE 50,50   OF oDlgEnviaXML ;
        FROM RESOURCE 10001 ;
        AUTOPLAY
   @ 7,21 GET DATEPICKER oDate1 VAR vDate1 SIZE 98,32  
   @ 7,5 SAY oLabel1 CAPTION "Data inicial"  TRANSPARENT SIZE 55,21  
   @ 123,21 GET DATEPICKER oDate2 VAR vDate2 SIZE 98,32  
   @ 123,5 SAY oLabel2 CAPTION "Data final"  TRANSPARENT SIZE 46,21  
   @ 228,21 BUTTONEX oButtonex1 CAPTION ""   TRANSPARENT SIZE 44,32 ;
        STYLE BS_CENTER +WS_TABSTOP  ;
        BITMAP (HBitmap():AddFile('res\eanok.bmp')):handle  ;
        BSTYLE 0 ;
        ON CLICK {|| query_xmls() }
   @ 303,21 GET oEmail VAR vEmail SIZE 377,32  MAXLENGTH 60  
   @ 303,5 SAY oLabel3 CAPTION "Email contab."  TRANSPARENT SIZE 64,21  
   @ 8,61 BROWSE oBrowse1EnviaXML ARRAY SIZE 869,354 ;
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
   @ 756,426 BUTTONEX oButtonex2 CAPTION "Envia para contab."   SIZE 121,32 ;
        STYLE BS_CENTER +WS_TABSTOP  ;
        ON CLICK {|| oButtonEx2:Disable(),enviaparacontab(),oButtonEx2:Enable()}

   @ 628,426 BUTTONEX oButtonex3 CAPTION "Marcar todas"   SIZE 121,32 ;
        STYLE BS_CENTER +WS_TABSTOP ;
        ON CLICK {|| marcatudo() }

   @ 723,31 GET CHECKBOX oZip VAR vZip  ;
        CAPTION "Compactar"  SIZE 110,22  ; 
        TOOLTIP 'Reúne todos os arquivos num arquivo compactado, formato ZIP.'
        
   ACTIVATE DIALOG oDlgEnviaXML ;
   on activate {|| oAnimation1:Hide() }
        
	rele vDate1, vDate2, vEmail, oDlgEnviaXML,  oDate1, oLabel1, oDate2, oLabel2, oButtonex1, oEmail, oLabel3 ;
        , oBrowse1EnviaXML, oButtonex2, oButtonex3, notas, vZip, oZip, oAnimation1
   SetTransparent( hWin, 255 )
   
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
return nil

Static Function pega_path_xml(value,this)
Local cSALVA_PATH:= CAMINHO_EXE(), cArq:= SELECTFILE("Arquivos Nota Fiscal Eletrônica", "*.xml",curdrive()+':\'+curdir()+'\nfe\criadas','Selecione o arquivo XML correspondente')
Local lEnv:= .f.
this:value:= cArq
oBrowse1EnviaXML:refreshline(.t.)
Dirchange(cSALVA_PATH)
RETURN .t.

Static Function query_xmls
Local oSql:= SR_GetConnection(), cSql, n:= {}, hArqNFe ;
    , rede:= iif(file("c:\acbrmonitorplus\acbrmonitor.exe"),.f.,.t.), patth:= curdrive()+':\'+curdir()+'\nfe\criadas\'
notas = {}
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
CreateArList(oBrowse1EnviaXML,notas)
return nil

Static Function enviaparacontab
Local a, vai:= .f., xmls:= {}, b, patth:= curdrive()+':\'+curdir()+'\nfe\criadas\'; 
    , pattz:= curdrive()+':\'+curdir()+'\nfe\canceladas\', cZip:= '', cSql, oSql:= SR_GetConnection()
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
for a:= 1 to len(notas)
    if notas[a,6]
       aadd(xmls,notas[a,4])
       if file(pattz + notas[a,7] + '-cancelamento.xml')
          aadd(xmls,pattz + notas[a,7] + '-cancelamento.xml')
       endif
    endif
next
if vZip
   cZip:= curdrive()+':\'+curdir()+"\swap\"+cdpar000->cnpj+'.zip'
   if !hb_zipfile(cZip,xmls,5,,.T.,,.F.,.F.,)
      msgstop('Falha na compactação dos arquivos !')
      return nil
   endif
endif
if vZip
   xmls = {cZip}
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
   for b:= 1 to len(notas)
       if notas[b,6]
          cSql:= "update cdcab000 set xmlctbdt="+sr_cdbvalue(date())+",xmlctbhr="+sr_cdbvalue(time())
          cSql = cSql + " where sr_recno="+sr_cdbvalue(notas[b,8])
          oSql:Execute(cSql,.t.) 
          oBrowse1EnviaXML:aArray[b,5] = date()
          oBrowse1EnviaXML:RefreshLine(.t.)
       endif
   next    
   oSql:Commit()
endif
oBrowse1EnviaXML:Refresh()
oAnimation1:Hide()
return nil