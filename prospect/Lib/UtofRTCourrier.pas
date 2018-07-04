unit UtofRTCourrier;

interface

uses  Windows,Controls,Classes,forms,sysutils,
      HCtrls,HMsgBox,UTOF,UTOB,UTom,M3FP,ShellAPI,Hent1,FileCtrl,
{$IFDEF EAGLCLIENT}
      Maineagl,
{$ELSE}
      db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fe_Main,
{$ENDIF}
       Vierge,
{$IFDEF PAUL}
       Web,
{$ENDIF}
       Paramsoc,UtilPGI,UtomAction,HRichOLE, UtilWord,UtilRT,EntGC,
       menus,
//bur 11110 {$ifndef BUREAU}
       UTOFDOCEXTERNE,
//bur 11110 {$endif}
       UtofMailing,HTB97,EntRT,Web,AglInit
{$IF Defined(CRM) or Defined(GRCLIGHT)}
      ,UtofRTSaisieInfosEsker,EskerInterface
{$IFEND}
       ;

Type
    TOF_RTCourrier = Class (TOF)
    Private
     TC,TA,TI,TY,TP, TF : Tob ;
     NumeroContact : string;
     ClientParticulier : Boolean;
     SOrtdoclienaction : Boolean;
     SOrtactresp,stProduitpgi : string;
     stMaquetteValide : string;
     stNatureDoc : string;
     procedure ChargeAdresse ;
     procedure GenereAction (stDocWord,TypeAction,stMaquette : string) ;
     Procedure InitTobFusion(  Prefixe : string; T: TOB);
     procedure FusionneTob ;
//bur 11110{$ifndef BUREAU}
     procedure MAQUETTE_OnElipsisClick(Sender: TObject);
//bur 11110{$endif}
{$IF Defined(CRM) or Defined(GRCLIGHT)}
     procedure BESKERCOURRIER_OnClick(Sender: TObject);
     procedure BESKEREMAIL_OnClick(Sender: TObject);
     procedure BESKERFAX_OnClick(Sender: TObject);
     procedure BESKERSUIVI_OnClick(Sender: TObject);
{$IFEND}

    Public
     procedure OnLoad ; override ;
     procedure OnUpdate ; override ;
     procedure OnClose ; override ;
     procedure ChangeContact ;
     procedure ChangeMaquette ;
     procedure ChangeDocument ;
     procedure ChangeAdresse (NumAdresse : integer );
     procedure OnArgument(Arguments : String ) ; override ;
     procedure OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
//bur 11110{$ifndef BUREAU}
     procedure OuvreMaquette (Mode : string);
//bur 11110{$ENDIF}
     procedure BAdresse_OnClick(Sender: TObject);
{$IFDEF PAUL}
     procedure Maporama ;
{$ENDIF}
     END;

function GetNextAction (Aux : string) : integer;

implementation

procedure TOF_RTCourrier.OnArgument(Arguments : String ) ;
var     x : integer;
        Critere : string ;
        ChampMul,ValMul: string;
begin
inherited ;
{$IFDEF PAUL}
SetControlVisible('BMAPORAMA', True);
{$ENDIF}

Numerocontact:='';
ClientParticulier := False;
Critere := Arguments ;
Repeat
    Critere:=uppercase(ReadTokenSt(Arguments)) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));

           if ChampMul='T_AUXILIAIRE' then SetControlText('T_AUXILIAIRE', ValMul);
           if ChampMul='C_NUMEROCONTACT' then Numerocontact:=ValMul;
           end;
        end;
until  Critere='';
Tf:=Nil;
TI:=Tob.create('TIERS', Nil,-1);
TI.selectDB('"'+GetControlText('T_AUXILIAIRE')+'"',Nil,false);
TA:=Tob.create('ADRESSES',Nil,-1);
ChargeAdresse ;
TC:=Tob.create('CONTACT',Nil,-1) ;
TC.InitValeurs(false);
TP:=Tob.create('PROSPECTS', Nil,-1);
TP.selectDB('"'+GetControlText('T_AUXILIAIRE')+'"',Nil,false);
TY:=Tob.create('TIERSCOMPL', Nil,-1);
TY.selectDB('"'+GetControlText('T_AUXILIAIRE')+'"',Nil,false);

if (TI.GetValue('T_NATUREAUXI') = 'CLI') or (TI.GetValue('T_NATUREAUXI') = 'PRO') or
   ((TI.GetValue('T_NATUREAUXI') = 'FOU') and (Not VH_GC.GRFSeria)) then
   begin
//   SetControlText('MAQUETTE',GetParamsocSecur('SO_RTDIRMAQUETTE','C:\PGI00\STD\MAQUETTE')+'\*.doc');
   SetControlText('DOCUMENT',GetParamsocSecur('SO_RTDIRSORTIE','')+'\Courrier_'+TI.GetValue('T_LIBELLE')+'.doc');
   SetControlChecked('OUVRIRDOC',GetParamSocSecur('SO_RTDOCOUVRIRDOC',False));
   end
else
   begin
//   SetControlText('MAQUETTE',GetParamsocSecur('SO_RFDIRMAQUETTE','C:\PGI00\STD\MAQUETTE')+'\*.doc');
   SetControlText('DOCUMENT',GetParamsocSecur('SO_RFDIRSORTIE','')+'\Courrier_'+TI.GetValue('T_LIBELLE')+'.doc');
   SetControlChecked('OUVRIRDOC',GetParamSocSecur('SO_RFDOCOUVRIRDOC',False));
   end;
if (TI.GetValue('T_NATUREAUXI') = 'CLI') or (TI.GetValue('T_NATUREAUXI') = 'PRO') or
   ((TI.GetValue('T_NATUREAUXI') = 'FOU') and (VH_GC.GRFSeria)) then
    begin
    SetControlVisible('GENEREACTION',true); SetControlVisible('TYPEACTION',true);
    if (TI.GetValue('T_NATUREAUXI') = 'CLI') or (TI.GetValue('T_NATUREAUXI') = 'PRO') then
       begin
       SetControlChecked('GENEREACTION',GetParamSocSecur('SO_RTDOCGENEREACTION',True));
       SetControlText( 'TYPEACTION',GetParamSocSecur('SO_RTDOCTYPEACTION',''));
       SOrtdoclienaction := GetParamSocSecur('SO_RTDOCLIENACTION',False);
       SOrtactresp := GetParamSocSecur('SO_RTACTRESP','UTI');
       stProduitpgi := 'GRC';
       end
    else
       begin
       SetControlChecked('GENEREACTION',GetParamSocSecur('SO_RFDOCGENEREACTION',True));
       SetControlText( 'TYPEACTION',GetParamSocSecur('SO_RFDOCTYPEACTION',''));
       SOrtdoclienaction := GetParamSocSecur('SO_RFDOCLIENACTION',False);
       SOrtactresp := 'UTI';
       stProduitpgi := 'GRF';
       SetControlProperty ('TYPEACTION','Datatype','RTTYPEACTIONFOU');
       end;
    end;

SetControlEnabled('BMAQUETTE',  true) ;
if stProduitPgi = 'GRF' then stNatureDoc := 'MLF' else stNatureDoc := 'MLC';
if TI.GetValue('T_PARTICULIER') = 'X' then
   begin
   SetControlProperty ('ADR_JURIDIQUE', 'DataType', 'TTCIVILITE');
   ClientParticulier := True;
   end;

if Assigned(GetControl('BADRESSE_ALL')) then
   TMenuItem(GetControl('BADRESSE_ALL')).OnClick := BAdresse_OnClick;
if Assigned(GetControl('BADRESSE_CONT')) then
   TMenuItem(GetControl('BADRESSE_CONT')).OnClick := BAdresse_OnClick;
//bur 11110{$ifndef BUREAU}
if Assigned(GetControl('MAQUETTE')) then
  THEDIT(GetControl('MAQUETTE')).OnElipsisClick := MAQUETTE_OnElipsisClick;
(* //bur 11110 {$else}
  SetControlVisible ('MAQUETTE',false); // dans le bureau, il ne faut pas les fichiers en base pour pas avoir affaireOle
{$endif} *)
{$IF Defined(CRM) or Defined(GRCLIGHT)}
if GetParamSocSecur ('SO_RTGESTIONFLYDOC',False) = True then
  begin
  SetControlVisible ('BESKERSUIVI',True);
  SetControlVisible ('BESKERCOURRIER',True);
  SetControlVisible ('BESKEREMAIL',True);
  SetControlVisible ('BESKERFAX',True);
  end;
if Assigned(GetControl('BESKERCOURRIER')) then
   TToolbarButton97(GetControl('BESKERCOURRIER')).OnClick := BESKERCOURRIER_OnClick;
if Assigned(GetControl('BESKEREMAIL')) then
   TToolbarButton97(GetControl('BESKEREMAIL')).OnClick := BESKEREMAIL_OnClick;
if Assigned(GetControl('BESKERFAX')) then
   TToolbarButton97(GetControl('BESKERFAX')).OnClick := BESKERFAX_OnClick;
if Assigned(GetControl('BESKERSUIVI')) then
   TToolbarButton97(GetControl('BESKERSUIVI')).OnClick := BESKERSUIVI_OnClick;
{$IFEND}
end;

procedure AddChamp(TF, T : tob; nom : string;CliPart : Boolean);
var st2 : string;
    begin
    TF.addChampSup(nom, false);
    TF.putValue(nom,T.getValue(nom));
    if (ChampToType(nom)='COMBO') or (copy(nom,1,9)='YTC_TABLE') or (nom='T_REPRESENTANT') then
       begin
       st2:= ChampToLibelle(nom)+' libelle' ;
       TF.addChampSup(st2 , false);
       if ((nom = 'ADR_JURIDIQUE') and (CliPart = True)) then TF.putValue(st2 ,rechdom('TTCIVILITE',T.getValue(nom),false))
       else TF.putValue(st2 ,rechdom(ChampToTT(nom),T.getValue(nom),false));
       end;
    end;

procedure TOF_RTCourrier.FusionneTob ;
begin
if TF<>nil then Tf.free;
TF := TOB.Create('LaTobFusion',Nil,-1);
addchamp(TF,TI, 'T_TIERS' ,ClientParticulier);
addchamp(TF,TI, 'T_REPRESENTANT' ,ClientParticulier );
addchamp(TF,TI, 'T_NATUREAUXI' ,ClientParticulier);
addchamp(TF,TI, 'T_SIRET' ,ClientParticulier );
addchamp(TF,TI, 'T_TELEPHONE' ,ClientParticulier);
addchamp(TF,TI, 'T_FAX' ,ClientParticulier);
addchamp(TF,TI, 'T_TELEX' ,ClientParticulier);
addchamp(TF,TI, 'T_ZONECOM' ,ClientParticulier);
addchamp(TF,TI, 'T_SECTEUR' ,ClientParticulier);
addchamp(TF,TI, 'T_ORIGINETIERS' ,ClientParticulier);
addchamp(TF,TI, 'T_TARIFTIERS' ,ClientParticulier);
addchamp(TF,TI, 'T_RVA' ,ClientParticulier);

// InitTobFusion('T', TI);
InitTobFusion('ADR', TA);
InitTobFusion('C', TC);
InitTobFusion('YTC', TY);
InitTobFusion('RPR', TP);
end;


Procedure TOF_RTCourrier.InitTobFusion(  Prefixe : string; T: TOB);
var Q : TQuery;
    st : string;
begin
Q:=OpenSql('SELECT * FROM DECHAMPS WHERE (DH_PREFIXE="'+Prefixe+'") AND (DH_CONTROLE like "%L%")',True);
while not Q.EOF  do
    begin
    st := Q.Findfield('DH_NOMCHAMP').asString;
    if copy(ChampToLibelle(st),1,2)<>'.-' then  addchamp(TF,T,st,ClientParticulier);
    Q.next ;
    end;
Ferme(Q);
end;

procedure TOF_RTCourrier.ChargeAdresse ;
begin
TA.InitValeurs(false);
TA.putvalue('ADR_JURIDIQUE', TI.GetValue('T_JURIDIQUE'));
TA.putvalue('ADR_LIBELLE', TI.GetValue('T_LIBELLE'));
TA.putvalue('ADR_LIBELLE2', TI.GetValue('T_PRENOM'));
TA.putvalue('ADR_ADRESSE1', TI.GetValue('T_ADRESSE1'));
TA.putvalue('ADR_ADRESSE2', TI.GetValue('T_ADRESSE2'));
TA.putvalue('ADR_ADRESSE3', TI.GetValue('T_ADRESSE3'));
TA.putvalue('ADR_CODEPOSTAL', TI.GetValue('T_CODEPOSTAL'));
TA.putvalue('ADR_VILLE', TI.GetValue('T_VILLE'));
TA.putvalue('ADR_PAYS', TI.GetValue('T_PAYS'));
end;

procedure TOF_RTCourrier.ChangeContact ;
var QQ : TQuery ;
begin

if (getControlText('C_NUMEROCONTACT')<> '') then
   begin
   TC.InitValeurs(false);
   QQ:=OpenSql('SELECT * from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+GetControlText('T_AUXILIAIRE')+'" AND C_NUMEROCONTACT='+GetControlText('C_NUMEROCONTACT'),True) ;
   TC.SelectDB('',QQ,False );
   ferme(QQ);
   TC.PutEcran (TFVierge(Ecran),TWinControl(getControl('GB_CONTACT')));
// cd 17/07/01
   end else
   begin
   TC.InitValeurs(false);
   TC.PutEcran (TFVierge(Ecran),TWinControl(getControl('GB_CONTACT')));
// cd 17/07/01
   end;
end;

procedure TOF_RTCourrier.ChangeAdresse (NumAdresse : integer );
var Q : TQuery;
begin
if NumAdresse=0 then exit ;
Q:=OpenSql('SELECT * from ADRESSES where ADR_TYPEADRESSE = "TIE" AND ADR_NUMEROADRESSE='+inttostr(NumAdresse),True) ;
if Not Q.EOF then TA.SelectDB('',Q,False );
ferme(Q);
//TA.SelectDB(inttostr(NumAdresse),Nil,False );
TA.PutEcran (TFVierge(Ecran),TWinControl(getControl('GB_ADRESSE')));
end;

procedure TOF_RTCourrier.ChangeMaquette;
var //stMaquette: string;
    OkOk : Boolean ;
begin
{OkOk:=False;
stMaquette:=getControlText('MAQUETTE');
if (stMaquette<>'') and (pos('\*.',stMaquette)=0) and  FileExists(stMaquette) then OkOk:=true;   }
if StMaquetteValide <> '' then DeleteFile(pchar(StMaquetteValide));
StMaquetteValide := '';
OkOk := Ctrl_Maquette (GetControlText('MAQUETTE'),Ecran.Caption,stNatureDoc,StMaquetteValide);
//SetControlEnabled('BMAQUETTE',  OkOk) ;
SetControlEnabled('BVALIDER',  OkOk) ;
end;

procedure TOF_RTCourrier.ChangeDocument;
var stDocument: string;
    OkOk : boolean ;
begin
OkOk:=False;
stDocument:=getControlText('DOCUMENT');
if (stDocument<>'') and (pos('\*.',stDocument)=0) and  FileExists(stDocument) then OkOk:=true;
SetControlEnabled('BDOCUMENT',OkOk) ;
SetControlEnabled('BDETRUIREDOC',OkOk) ;
SetControlEnabled('BMAIL',  ((GetControlText('C_RVA')<>'') and OkOk) ) ;
{$IF Defined(CRM) or Defined(GRCLIGHT)}
SetControlEnabled('BESKEREMAIL', ((GetControlText('C_RVA')<>'') and OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ) ;
SetControlEnabled('BESKERFAX',  (((TC.GetString('C_FAX')<>'') or (TI.GetString('T_FAX') <> '')) and OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ) ;
SetControlEnabled('BESKERCOURRIER',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
{$IFEND}
end;

//bur 11110 {$ifndef BUREAU}
procedure TOF_RTCourrier.OuvreMaquette (Mode : string);
var //stMaquette: string;
    OkOk : Boolean;
begin
TI.getecran (TFVierge(Ecran),Nil);
TA.getecran (TFVierge(Ecran),TWinControl(getControl('GB_ADRESSE')));
if (getControlText('C_NUMEROCONTACT')<> '') then TC.getEcran (TFVierge(Ecran),TWinControl(getControl('GB_CONTACT')));
FusionneTob;


if (Mode = 'C') then LancePublipostage('TOB','','',TF)
else
    begin
{    stMaquette:=getControlText('MAQUETTE');
//   if (stMaquette<>'') and (pos('\*.',stMaquette)<>0)  then stMaquette:='' ;
    if (stMaquette<>'') and (pos('\*.',stMaquette)=0) and FileExists(stMaquette) then LancePublipostage('TOB',stMaquette,'',TF)
    else PGIBox('La maquette n''existe pas', ecran.caption);   }
    AFLanceFiche_DocExterneMul('', 'TYPE=WOR;NATURE='+stNatureDoc+';TOBADR='+intToStr(longint(TF)));
    if (GetControlText('MAQUETTE') <> '') and (stMaquetteValide <> '') and (not FileExists(stMaquetteValide)) then
      begin
      StMaquetteValide := '';
      OkOk := Ctrl_Maquette (GetControlText('MAQUETTE'),Ecran.Caption,stNatureDoc,StMaquetteValide);
      SetControlEnabled('BVALIDER',  OkOk) ;
      end;
      end;
end;
//bur 11110 {$endif}

procedure TOF_RTCourrier.OnLoad ;
begin
inherited ;
TI.putecran (TFVierge(Ecran),Nil);
TA.putecran (TFVierge(Ecran),TWinControl(getControl('GB_ADRESSE')));
if NumeroContact<>'' then begin SetControlText('C_NUMEROCONTACT', NumeroContact); ChangeContact; end;
TC.PutEcran (TFVierge(Ecran),TWinControl(getControl('GB_CONTACT')));
end;

procedure TOF_RTCourrier.OnUpdate ;
var stDocWord , AQui: string;
begin
inherited ;
TI.getecran (TFVierge(Ecran),Nil);
TA.getecran (TFVierge(Ecran),TWinControl(getControl('GB_ADRESSE')));
if (getControlText('C_NUMEROCONTACT')<> '') then TC.getEcran (TFVierge(Ecran),TWinControl(getControl('GB_CONTACT')));
stDocWord:=getControlText('DOCUMENT');
{stMaquette:=getControlText('MAQUETTE');
if not FileExists(stMaquette) then begin PGIBox('La maquette n''existe pas', ecran.caption); exit; end; }
if (ExtractFileDrive(stDocWord) = '') or (not DirectoryExists(ExtractFilePath(stDocWord))) then begin PGIBox(Format(TraduireMemoire('Le répertoire %s n''existe pas'),[ExtractFilePath(stDocWord)]), ecran.caption); exit; end;
if FileExists(stDocWord) then
   if PGIAsk('Ce fichier existe déjà, voulez-vous l''écraser ?',ecran.caption)=mryes then DeleteFile(stDocWord ) else exit;
//ConvertDocFile(stMaquette,stDocWord,Nil,Nil,Nil,OnGetVar,Nil,Nil) ;
FusionneTob;
LancePublipostage('TOB',stMaquetteValide,stDocWord,TF,'',Nil,True);
if getControlText('GENEREACTION')='X' then GenereAction (stDocWord,getControlText('TYPEACTION'),stMaquetteValide);
if getControlText('OUVRIRDOC')='X' then ShellExecute( 0, PCHAR('open'),PChar(stDocWord), Nil,Nil,SW_RESTORE);   //PCS
Aqui:=TC.Getvalue('C_RVA');
if (Aqui<>'') and (getControlText('ENVOYERDOC')='X') then PGIEnvoiMail ('',Aqui,'',nil,stDocWord,false) ;
//if getControlText('GENEREACTION')='X' then GenereAction (stDocWord,getControlText('TYPEACTION'),stMaquette);
ChangeDocument ;
end;

procedure TOF_RTCourrier.OnClose ;
begin
TA.free;TC.free;TI.free;TY.free;TP.free;TF.free;
if StMaquetteValide <> '' then DeleteFile(pchar(StMaquetteValide));
inherited ;
end;

function GetNextAction (Aux : string) : integer;
var Q : Tquery ;
begin
result:=1;
Q:=OpenSQL('SELECT MAX(RAC_NUMACTION) FROM ACTIONS WHERE RAC_AUXILIAIRE = "'+Aux+'"', True);
if not Q.Eof then result := Q.Fields[0].AsInteger+1;
Ferme(Q) ;
end;


procedure TOF_RTCourrier.GenereAction (stDocWord,TypeAction,stMaquette : string);
var TA : Tob;
    TM : Tom_actions;
    b_result:boolean;
    oleDoc : THRichEditOLE ;
    Q : Tquery;
    Select : string;
begin
TM:=Tom_Actions(CreateTOM('ACTIONS',Nil,False,False) );
TM.stTiers:=TI.GetValue('T_TIERS') ;
TM.stProduitpgi := stProduitpgi;
TM.soRtactresp := SOrtactresp;
TA:=Tob.create('ACTIONS',nil,-1);
TM.InitTob(TA);

// creation lien Ole vers le Document
if SOrtdoclienaction then
   begin
   oleDOc:=THRichEditOLE(GetControl('RAC_BLOCNOTE'));
   oleDoc.InsertFileLink(stDocWord, True{iconic});
   TA.GetECran(TFVierge(Ecran),Nil);
   end;

TA.putvalue('RAC_TYPEACTION', typeAction);
TA.putvalue('RAC_NUMEROCONTACT',TC.getValue('C_NUMEROCONTACT'));
TA.putvalue('RAC_LIBELLE', ExtractFileName(stMaquette));
TA.putvalue('RAC_NUMACTION', GetNextAction(TI.GetValue('T_AUXILIAIRE')));
TA.putvalue('RAC_ETATACTION', 'REA');
if SOrtactresp='COM'  then
   TA.putvalue('RAC_INTERVENANT', RTRechResponsable (TI.GetValue('T_TIERS')));
    //mcd 30/11/2005 il faut prendre en compte le paramétrage possible intervenent que ressource
if (( soRtactresp = 'RE1' )
 or ( soRtactresp = 'RE2' )
 or ( soRtactresp = 'RE3' ))
 and (CtxAffaire in V_PGI.PGIContexte) then
 begin
   Select :='SELECT YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3 FROM  TIERSCOMPL WHERE YTC_TIERS="'+TI.GetValue('T_TIERS')+'"';
   Q := OpenSQL(Select, True);
   if not Q.Eof then
        begin
        if ( soRtactresp = 'RE1' ) then TA.putvalue('RAC_INTERVENANT',Q.FindField('YTC_RESSOURCE1').asstring);
        if ( soRtactresp = 'RE2' ) then TA.putvalue('RAC_INTERVENANT',Q.FindField('YTC_RESSOURCE2').asstring);
        if ( soRtactresp = 'RE3' ) then TA.putvalue('RAC_INTERVENANT',Q.FindField('YTC_RESSOURCE3').asstring);
        end;
   Ferme(Q);
   end;
      //fin mcd 30/11/2005
TA.putvalue('RAC_DATECREATION', Date);

if (TM.VerifTOB( TA )) then
  begin
  b_result := TA.InsertDB (nil);
  If Not b_result Then PGIBox('Impossible de générer l''action', ecran.caption)
  end else
      PGIBox(TM.LastErrorMsg, ecran.caption);

if TM <> NIL then TM.Free ;
if TA <> NIL then TA.Free ;
end;

procedure TOF_RTCourrier.BAdresse_OnClick(Sender: TObject);
var Noadr, Params, Range : string;
begin
  Range :='ADR_TYPEADRESSE=TIE;ADR_REFCODE=' + TI.GetValue('T_TIERS') ;
  if (Uppercase(TMenuItem(Sender).name) = 'BADRESSE_CONT') and (GetControlText('C_NUMEROCONTACT') <> '') then Range := Range + ';ADR_NUMEROCONTACT=' + GetControlText('C_NUMEROCONTACT');
  Params := 'TYPEADRESSE=TIE;NATUREAUXI=' + TI.GetValue('T_NATUREAUXI')
          +';YTC_TIERSLIVRE=' + TI.GetValue('T_TIERS')
          +';CLI=' + GetControlText('T_AUXILIAIRE')
          +';TIERS=' + TI.GetValue('T_TIERS')
          +';PART=' + TI.GetValue('T_PARTICULIER')
          +';ACTION=MODIFICATION';

  Noadr := AglLanceFiche('GC', 'GCADRESSES', Range, '', Params);
  ChangeAdresse (Valeuri(Noadr));
end;

//bur 11110 {$ifndef BUREAU}
procedure TOF_RTCourrier.MAQUETTE_OnElipsisClick(Sender: TObject);
var NoMaquette : string;
begin
  NoMaquette := AFLanceFiche_DocExterneMulSelec('ADE_DOCEXETAT=UTI', 'TYPE=WOR;NATURE='+stNatureDoc+';SELECTION');
  SetControlText('MAQUETTE',NoMaquette)
end;
//bur 11110 {$endif}

{$IF Defined(CRM) or Defined(GRCLIGHT)}
procedure TOF_RTCourrier.BESKERCOURRIER_OnClick(Sender: TObject) ;
var Infos : string;
    TobInfos,TobAdr : Tob;
begin
  if LoginEsker = False then Exit;
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
  TobAdr.AddChampSupValeur ('ADR_JURIDIQUE',TA.GetString('ADR_JURIDIQUE'),False);
  TobAdr.AddChampSupValeur ('ADR_LIBELLE',TA.GetString('ADR_LIBELLE'),False);
  TobAdr.AddChampSupValeur ('ADR_SUITELIBELLE',TA.GetString('ADR_LIBELLE2'),False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE1',TA.GetString('ADR_ADRESSE1'),False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE2',TA.GetString('ADR_ADRESSE2'),False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE3',TA.GetString('ADR_ADRESSE3'),False);
  TobAdr.AddChampSupValeur ('ADR_CODEPOSTAL',TA.GetString('ADR_CODEPOSTAL'),False);
  TobAdr.AddChampSupValeur ('ADR_VILLE',TA.GetString('ADR_VILLE'),False);
  TobAdr.AddChampSupValeur ('ADR_PAYS',TA.GetString('ADR_PAYS'),False);
  TobAdr.AddChampSupValeur ('C_NOM',TC.GetString('C_NOM'),False);
  TobAdr.AddChampSupValeur ('C_PRENOM',TC.GetString('C_PRENOM'),False);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','C',False);
//  TheTob := TobInfos;
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIECOURRIER','','','Document='+GetControlText('DOCUMENT')+';TobAdr='+intToStr(longint(TobAdr))+';TobInfos='+intToStr(longint(TobInfos)));
//  TobInfos := TheTob;
  if Infos <> '' then
  begin
    SendEsker (0,TobInfos,GetControlText('DOCUMENT'),TobAdr);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_RTCourrier.BESKEREMAIL_OnClick(Sender: TObject) ;
var Infos,Chp : string;
    TObInfos,TobAdr : Tob;
    Q : TQuery;
begin
  if LoginEsker = False then Exit;
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','C',False);
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIEEMAIL','','','TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    Q:=OpenSQL('Select ARS_EMAIL FROM RESSOURCE WHERE ARS_UTILASSOCIE="'+V_PGI.User+'"',true);
    if Not Q.EOF then Chp:=Q.FindField('ARS_EMAIL').asstring;
    ferme(Q);
    TobInfos.AddChampSupValeur ('EMAILEMETTEUR',Chp,False);
    TobAdr.AddChampSupValeur ('NOMCONTACT',TC.GetString('C_NOM'),False);
    TobAdr.AddChampSupValeur ('NOMSOCCONTACT',TA.GetString('ADR_LIBELLE'),False);
    TobAdr.AddChampSupValeur ('EMAILCONTACT',TC.GetString('C_RVA'),False);
    SendEsker (2,TobInfos,GetControlText('DOCUMENT'),TobAdr);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_RTCourrier.BESKERFAX_OnClick(Sender: TObject) ;
var Infos,Chp : string;
    TObInfos,TobAdr : Tob;
begin
  if LoginEsker = False then Exit;
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','C',False);
  TobAdr.AddChampSupValeur ('NOMCONTACT',TC.GetString('C_PRENOM') + ' ' + TC.GetString('C_NOM'),False);
  TobAdr.AddChampSupValeur ('NOMSOCCONTACT',TA.GetString('ADR_LIBELLE'),False);
  if TC.GetString('C_FAX') <> '' then Chp := TC.GetString('C_FAX')
  else Chp := TI.GetString('T_FAX');
  TobAdr.AddChampSupValeur ('FAXCONTACT',CleTelephone (Chp),False);
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIEFAX','','','Document='+GetControlText('DOCUMENT')+';TobAdr='+intToStr(longint(TobAdr))+';TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    SendEsker (1,TobInfos,GetControlText('DOCUMENT'),TobAdr);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_RTCourrier.BESKERSUIVI_OnClick(Sender: TObject) ;
var sHttp : String;
begin
  sHttp := 'https://as1.ondemand.esker.com/ondemand/webaccess/40/documents.aspx';
  LanceWeb(sHttp,False);
end;
{$IFEND}

// fonction appelée par RT_Parser

procedure TOF_RTCourrier.OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
var TTOB : tob ;
BEGIN
if Pos('T_',VarName)>0 then TTob:=TI
else if Pos('ADR_',VarName)>0 then TTob:=TA
else if Pos('C_',VarName)>0 then  TTob:=TC
else if Pos('YTC_',VarName)>0 then  TTob:=TY
else if Pos('RPR_',VarName)>0 then  TTob:=TP
else exit ;

if Pos('BLOCNOTE',VarName)>0
    then  begin BlobToFile(VarName,TTob.GetValue(VarName)); Value:=''; end else
          if ChampToType( VarName)= 'COMBO' then Value:=RechDom(ChampToTT( VarName),TTob.GetValue(VarName),false)
             else  Value:=TTob.GetValue(VarName);
END ;

//====== pour script ===========================

Procedure AGLRTCourrier_ChangeContact( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_RTCourrier(totof).ChangeContact;
end;
{Procedure AGLRTCourrier_ChangeAdresse( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_RTCourrier(totof).ChangeAdresse(Integer(parms[1])) ;
end;   }
Procedure AGLRTCourrier_ChangeMaquette( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_RTCourrier(totof).ChangeMaquette;
end;
Procedure AGLRTCourrier_ChangeDocument( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_RTCourrier(totof).ChangeDocument;
end;

Procedure AGLRTCourrier_OuvreMaquette( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
//bur 11110 {$ifndef BUREAU}
  TOF_RTCourrier(totof).OuvreMaquette(Parms[1]);
//bur 11110 {$endif}
end;

{$IFDEF PAUL}
Procedure AGLRTCourrier_Maporama( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_RTCourrier(totof).Maporama;
end;

Function PaysISO(Pays : string) :string;
var Q : Tquery;
begin
Result:='';
Q:=OpenSql('Select PY_CODEISO2 from PAYS where PY_PAYS="'+Pays+'"', True);
If not Q.EOF then result:=Q.Fields[0].asString;
ferme(Q);
end;

procedure TOF_RTCourrier.Maporama;
var Fichier: TextFile;
    OkOk : boolean ;
begin

  AssignFile(Fichier, IncludeTrailingBackSlash(TCBPPath.GetCegidUserDocument) + 'maporama.htm');
  Rewrite (Fichier);

  WriteLn (Fichier,'<html>') ;
  WriteLn (Fichier,'<head>') ;
  WriteLn (Fichier,'<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">') ;
  WriteLn (Fichier,'<meta name="GENERATOR" content="Microsoft FrontPage 4.0">') ;
  WriteLn (Fichier,'<meta name="ProgId" content="FrontPage.Editor.Document">') ;
  WriteLn (Fichier,'<title>Fill the fields below to get the map of the required localization</title>') ;
  WriteLn (Fichier,'</head>') ;

  WriteLn (Fichier,'<body>') ;

  WriteLn (Fichier,'<form action="http://www.maporama.com/share/Map.asp?PDT=maposearch" method="post" name="FindAdr"> ') ;
  WriteLn (Fichier,'  <input type="hidden" name="_XgoGCAdrCommand" value="run">') ;
  WriteLn (Fichier,'  <font size="1" face="Arial, Helvetica, sans-serif">') ;
  WriteLn (Fichier,'    Fill the fields below to get the map of the required localization. City is the only mandatory field.') ;
  WriteLn (Fichier,'    <hr size="3">') ;
  WriteLn (Fichier,'    Address ') ;
  WriteLn (Fichier,'  </font><br> ') ;
  WriteLn (Fichier,'  <input type="Text" name="_XgoGCAddress" value="'+TA.getValue('ADR_ADRESSE1') +'" size="38" maxlength="70"><br>') ;
  WriteLn (Fichier,'  <font size="1" face="Arial, Helvetica, sans-serif">') ;
  WriteLn (Fichier,'    Post Code') ;
  WriteLn (Fichier,'  </font><br>') ;
  WriteLn (Fichier,'  <input type="Text" name="Zip" size="10" maxlength="10" value="'+TA.getValue('ADR_CODEPOSTAL') +'"><br>') ;
  WriteLn (Fichier,'  <font size="1" face="Arial, Helvetica, sans-serif">') ;
  WriteLn (Fichier,'    City ') ;
  WriteLn (Fichier,'  </font><br> ') ;
  WriteLn (Fichier,'  <input type="text" name="_XgoGCTownName" value="'+TA.getValue('ADR_VILLE') +'" size="25" maxlength="70"> ') ;
  WriteLn (Fichier,'  <font size="1" face="arial"><br>Pays</font><br> ') ;
  WriteLn (Fichier,'  <input type="text" name="COUNTRYCODE" value="'+PaysISO(TA.getValue('ADR_PAYS')) +'" size="2" maxlength="2"> ') ;

  {WriteLn (Fichier,'  <select name="COUNTRYCODE"> ') ;
  WriteLn (Fichier,'      <option value="AD">Andorra</option><option value="AT">Austria</option>') ;
  WriteLn (Fichier,'      <option value="BE">België</option><option value="DK">Danmark</option>') ;
  WriteLn (Fichier,'      <option value="DE">Deutschland</option><option value="ES">Espana</option>') ;
  WriteLn (Fichier,'      <option value="FR" SELECTED >France</option><option value="IT">Italia</option> ') ;
  WriteLn (Fichier,'      <option value="LI">Liechtenstein</option><option value="LU">Luxembourg</option> ') ;
  WriteLn (Fichier,'      <option value="NL">Nederland</option><option value="PT">Portugal</option> ') ;
  WriteLn (Fichier,'      <option value="CH">Suisse</option><option value="SE">Sverige</option>') ;
  WriteLn (Fichier,'      <option value="UK">United Kingdom</option> ') ;

  WriteLn (Fichier,'  </select>') ;   }
  WriteLn (Fichier,'  <br> ') ;
  WriteLn (Fichier,'  <input type="Image" name="submit" src="go.gif" alt="Afficher la carte" value="submit" border="0" width="22" height="22">') ;
  WriteLn (Fichier,'  <img src="PoweredBy.jpg" width="88" height="31">') ;
  WriteLn (Fichier,'</form>') ;
  WriteLn (Fichier,'<SCRIPT LANGUAGE=VBSCRIPT>') ;
  WriteLn (Fichier,'   FindAdr.submit') ;
  WriteLn (Fichier,'</SCRIPT> ') ;

  WriteLn (Fichier,'</body>  ') ;

  WriteLn (Fichier,'</html> ') ;

  CloseFile(Fichier);

  LanceWeb(IncludeTrailingBackSlash(TCBPPath.GetCegidUserDocument) + 'maporama.htm',true);

end;
{$ENDIF}

Initialization
registerclasses([Tof_RTCourrier]);
RegisterAglProc( 'RTCourrier_ChangeContact', TRUE , 0, AGLRTCourrier_ChangeContact);
//RegisterAglProc( 'RTCourrier_ChangeAdresse', TRUE , 1, AGLRTCourrier_ChangeAdresse);
RegisterAglProc( 'RTCourrier_ChangeMaquette', TRUE ,0 , AGLRTCourrier_ChangeMaquette);
RegisterAglProc( 'RTCourrier_ChangeDocument', TRUE , 0, AGLRTCourrier_ChangeDocument);
RegisterAglProc( 'RTCourrier_OuvreMaquette', TRUE , 1, AGLRTCourrier_OuvreMaquette);

{$IFDEF PAUL}
RegisterAglProc( 'RTCourrier_Maporama', TRUE , 0, AGLRTCourrier_Maporama);
{$ENDIF}
end.
