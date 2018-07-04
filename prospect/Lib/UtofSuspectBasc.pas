unit UtofSuspectBasc;

interface

uses  ComCtrls,Sysutils, HCtrls,StdCtrls, HEnt1, Controls,Numconv,HeureUtil,
{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
      Forms,M3FP,HMsgBox,Classes, HStatus,ed_tools,UTob,utof,UtilGC,UtilRT,ParamSoc,Ent1,EntGC;

Const
     NbFillesMax : integer = 100;
    TexteMessage: array[1..2] of string 	= (
      {1}         'Saisir le début code prospect'
      {2}        ,'Attribution auto code client non-paramétrée'
                         );

type
    TOF_SuspectBasc = class(TOF)
     private
      Tobsuspect,TobsuspectTous,TOBsuspectMaj,TobSuspectcompl,TobCorrespond : Tob;
      Tobtiers,Tobtierscompl,TobProspect,TobContact: Tob;
      CodeAuxi,CodeTiers,CodeSuspect,NumChrono,SqlOrigine,ModeCreation,LmodeCreation:String;
      NbSuspect,LgCode : Integer;
      NbSuspectTotal,NbInsertions : Integer;
      ListRecap : Tmemo;
      BStop : boolean;
      procedure InsertLesTobs;
      procedure CreationTiers ;
      procedure CreationTiersCompl;
      procedure CreationProspect;
      procedure CreationContact;
      procedure JournalEvenement;
      procedure BasculeSuspect;
      procedure TOBCopieChamp(FromTOB, ToTOB : TOB);
      procedure CorrespondSuspectcompl(CleChoixCode : string; FromTOB, ToTOB : TOB);
     public
      procedure OnUpdate ; override ;
    end;


implementation


Procedure TOF_SuspectBasc.InsertLesTobs ;
begin
  InitMove(NbFillesMax,'') ;
  if V_PGI.IoError=oeOk then TobTiers.InsertDB(nil,False);
  if V_PGI.IoError=oeOk then  Tobtierscompl.InsertDB(nil,false);
  if V_PGI.IoError=oeOk then  TobProspect.InsertDB(nil,False);
  if V_PGI.IoError=oeOk then  TobContact.InsertDB(nil,False);
  if V_PGI.IoError=oeOk then   TOBsuspectMaj.UpdateDB(False);

  TobTiers.ClearDetail;
  Tobtierscompl.ClearDetail;
  TobProspect.ClearDetail;
  TobContact.ClearDetail;
  TOBsuspectMaj.ClearDetail;
  FiniMove ;
end;

procedure TOF_SuspectBasc.TOBCopieChamp(FromTOB, ToTOB : TOB);
var i_pos,i_ind1: integer;
    FieldNameTo,FieldNameFrom,St:string;
    PrefixeTo,PrefixeFrom : string;
begin
  PrefixeFrom := TableToPrefixe (FromTOB.NomTable);
  PrefixeTo := TableToPrefixe (ToTOB.NomTable);
  for i_ind1 := 1 to FromTOB.NbChamps do
    begin
      FieldNameFrom := FromTOB.GetNomChamp(i_ind1);
      St := FieldNameFrom ;
      i_pos := Pos ('_', St) ;
      System.delete (St, 1, i_pos-1) ;
      FieldNameTo := PrefixeTo + St ;
      if ToTOB.FieldExists(FieldNameTo) then
        ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom)) ;
    end;
end;

Procedure TOF_SuspectBasc.CorrespondSuspectcompl(CleChoixCode : string; FromTOB, ToTOB : TOB);
var Q : TQuery;
    PrefixeTo,FieldNameFrom,FieldNameTo :string;
    ind : integer;
begin
  PrefixeTo := TableToPrefixe (ToTOB.NomTable);
  if (TobCorrespond=Nil) then
  begin
    Q := OpenSQL('SELECT CC_LIBELLE,CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="'+CleChoixCode+'"', true);
    TobCorrespond:=TOB.create ('_CHOIXCOD',NIL,-1);
    TobCorrespond.LoadDetailDB('CHOIXCOD','','',Q, false, false);
    Ferme(Q);
  end; 
  if TobCorrespond.detail.count>0 then
  begin
    For ind:=0 to TobCorrespond.detail.count-1 do
      begin
        FieldNameFrom := TobCorrespond.detail[ind].getvalue('CC_LIBELLE');
        FieldNameTo := TobCorrespond.detail[ind].getvalue('CC_LIBRE');
        if (copy (FieldNameTo ,1,3) = PrefixeTo) then
          ToTOB.PutValue (FieldNameTo, FromTOB.GetValue (FieldNameFrom));
      end;
  end;
end;

Procedure  TOF_SuspectBasc.CreationTiers ;
Var TOBL : TOB ;
BEGIN
  TOBL:=Tob.create('TIERS',TobTiers,-1);
  TOBL.InitValeurs;
  TOBCopieChamp(TOBSuspect, TOBL);
  TOBL.PutValue('T_TIERS',CodeTiers) ;
  TOBL.PutValue('T_AUXILIAIRE',CodeAuxi) ;
  TOBL.PutValue ('T_FACTURE',CodeAuxi);
  TOBL.putvalue('T_NATUREAUXI','PRO');
  TOBL.putvalue('T_DATEINTEGR', V_PGI.DateEntree);

  TOBL.PutValue('T_ABREGE',Copy(TOBL.getvalue('T_LIBELLE'),1,17)) ;
  TOBL.PutValue('T_PHONETIQUE',phoneme(TOBL.getvalue('T_LIBELLE')));
  if (TOBL.GetValue ('T_PAYS')='') then TOBL.PutValue ('T_PAYS', GetParamSoc('SO_GcTiersPays'));
  TOBL.PutValue ('T_NATIONALITE',TOBL.getvalue('T_PAYS'));
  TOBL.PutValue ('T_CREERPAR', 'GEN');
  TOBL.PutValue ('T_DATEOUVERTURE', V_PGI.DateEntree);
  TOBL.PutValue ('T_DEVISE', V_PGI.devisepivot);
  TOBL.PutValue ('T_LETTRABLE', 'X');
  TOBL.PutValue ('T_REGIMETVA',VH^.RegimeDefaut);
  if VH_GC.GCDefFactureHT then TOBL.PutValue ('T_FACTUREHT','X') ;
  TOBL.PutValue ('T_COLLECTIF',VH^.DefautCli);
  TOBL.PutValue ('T_MODEREGLE',GetParamSoc('SO_GcModeRegleDefaut'));
  TOBL.PutValue ('T_TVAENCAISSEMENT', GetParamSoc('SO_TVAENCAISSEMENT'));
  if (Getparamsoc('SO_GCVENTEEURO')) then TOBL.PutValue ('T_EURODEFAUT','X')
                                    else TOBL.PutValue ('T_EURODEFAUT','-');
  TOBL.PutValue ('T_CONFIDENTIEL','0') ;
  TOBL.PutValue ('T_PUBLIPOSTAGE', 'X');
  TOBL.PutValue ('T_PARTICULIER','-');
  TOBL.PutValue ('T_DATECREATION',V_PGI.DateEntree);
  if VH_GC.GCIfDefCEGID then
  begin
    TOBL.PutValue ('T_REPRESENTANT','996');
    TOBL.PutValue ('T_NATIONALITE','FRA');
    TOBL.PutValue ('T_REGIMETVA','001');
    TOBL.PutValue ('T_COMPTATIERS','001');
  end;
end;

Procedure  TOF_SuspectBasc.CreationTiersCompl ;
Var TOBL : TOB ;
begin
  TOBL:=Tob.create('TIERSCOMPL',Tobtierscompl,-1);
  TOBL.InitValeurs;
  if TOBSuspectCompl.selectDB ('"'+CodeSuspect+'"',Nil,False) then
     CorrespondSuspectcompl('RSC',TOBSuspectCompl, TOBL);
  TOBL.PutValue ('YTC_TIERS',CodeTiers);
  TOBL.PutValue ('YTC_AUXILIAIRE',CodeAuxi);
end;

Procedure  TOF_SuspectBasc.CreationProspect ;
Var TOBL : TOB ;
begin
  TOBL:=Tob.create('PROSPECTS',TobProspect,-1);
  TOBL.InitValeurs;
  TOBL.PutValue ('RPR_AUXILIAIRE',CodeAuxi);
end;

Procedure  TOF_SuspectBasc.CreationContact ;
Var TOBL : TOB ;
begin
  if ( trim (TobSuspect.Getvalue('RSU_CONTACTNOM')) <> '' ) then
  begin
    TOBL:=Tob.create('CONTACT',TobContact,-1);
    TOBL.InitValeurs;
    TOBL.PutValue('C_TYPECONTACT','T') ;
    TOBL.PutValue('C_AUXILIAIRE',CodeAuxi) ;
    TOBL.PutValue('C_NUMEROCONTACT','1') ;
    TOBL.PutValue('C_NATUREAUXI','PRO') ;
    TOBL.PutValue('C_CIVILITE',TobSuspect.Getvalue('RSU_CONTACTCIVILITE')) ;
    TOBL.PutValue('C_NOM',TobSuspect.Getvalue('RSU_CONTACTNOM')) ;
    TOBL.PutValue('C_PRENOM',TobSuspect.Getvalue('RSU_CONTACTPRENOM')) ;
    TOBL.PutValue('C_FONCTIONCODEE',TobSuspect.Getvalue('RSU_CONTACTFONCTION')) ;
    TOBL.PutValue('C_PRINCIPAL','X') ;
    TOBL.PutValue('C_PUBLIPOSTAGE','X') ;
  end;
end;

procedure TOF_SuspectBasc.JournalEvenement ;
var TOBJnal : TOB ;
    NumEvt : integer ;
    QQ : TQuery ;
begin
NumEvt:=0 ;
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'RSP');
TOBJnal.PutValue('GEV_LIBELLE', 'Transfert des Suspects en Prospects');
TOBJnal.PutValue('GEV_DATEEVENT', V_PGI.DateEntree);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
ListRecap.lines.Add ('');
if BStop = True then
    begin
    TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
    ListRecap.lines.Add ('Interrompue par l''utilisateur');
    end else
    begin
    if V_PGI.IoError = oeOk then
        begin
        TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
        ListRecap.lines.Add('Traitement terminé');
        end else
        begin
        TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
        ListRecap.lines.Add ('Traitement non terminé');
        end;
    end;
TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Text);
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
end;


Procedure TOF_SuspectBasc.BasculeSuspect ;
var  Q : TQuery ;
     i_ind,index,i_Chrono : integer;
{$IFNDEF AGL550B}
     QRPFStop : TQRProgressForm;
{$ENDIF}
     BSTop : boolean;
     MessTransfert : string;
     datedebut,datefin :Tdatetime;
begin
  if PGIAsk('Voulez-vous transférer les Suspects en Prospects ?', 'Suspects en Prospects') <> mrYes then exit;
  Q:=OpenSQL('SELECT count(RSU_SUSPECT) FROM SUSPECTS WHERE RSU_FERME = "-"' + SqlOrigine , True) ;
  if Not Q.EOF then NbSuspectTotal:=Q.Fields[0].AsInteger;
  Ferme(Q) ;
{$IFDEF AGL550B}
  InitMoveProgressForm (nil,'Transfert des suspects en Prospects...','',NbSuspectTotal,TRUE,TRUE) ;
{$ELSE}
  QRPFStop := DebutProgressForm (Ecran,'Transfert des suspects en Prospects...','', 0, true, true) ;
  QRPFStop.MaxValue := NbSuspectTotal;
{$ENDIF}
  LgCode:=VH^.Cpta[fbAux].Lg ;
//  ListRecap:=TStringList.Create;
  BStop := False;
  TobCorrespond := Nil;
  datedebut := CurrentDate;
try
  TobsuspectTous := TOB.Create('Les Suspects', Nil, -1) ;
  Tobtiers := TOB.Create('les tiers', nil, -1);
  Tobtierscompl := TOB.Create('complementaires tiers', nil, -1);
  TobProspect := TOB.Create('complementaires prospects', nil, -1);
  TobContact := TOB.Create('les contacts', nil, -1);
  TOBsuspectMaj := TOB.Create('Maj Suspect', nil, -1);
  TobSuspectcompl := TOB.Create('SUSPECTSCOMPL', nil, -1);
{$IFDEF AGL550B}
  MoveCurProgressForm('Transfert en prospects de '+IntToStr(NbSuspectTotal)+' suspects');
{$ELSE}
  QRPFStop.Text := 'Transfert en prospects de '+IntToStr(NbSuspectTotal)+' suspects';
  QRPFStop.Value := 0;
{$ENDIF}

  ListRecap.lines[0] := ('Transfert des suspects en prospects - début à '+ FormateDate('hh "h" mm', datedebut));
  ListRecap.lines.Add ('');
  ListRecap.lines.Add ('');
  ListRecap.lines.Add (LmodeCreation);
  ListRecap.lines.Add ('');
  ListRecap.lines.Add ('Nombre de suspects à transférer :'+IntToStr(NbSuspectTotal));
  ListRecap.lines.Add ('');

  if (ModeCreation = 'debut') then
  begin
    i_Chrono := StrToInt(Numchrono);
    if (i_Chrono>0) then i_Chrono := i_Chrono-1;
    CodeTiers := IntToStr(i_Chrono)
  end else
  if (ModeCreation = 'auto') then Numchrono := GetParamsoc('SO_GCCOMPTEURTIERS');
  if (ModeCreation <> 'reprise') then
    ListRecap.lines.Add ( 'Compteur initial des TIERS : '+Numchrono);
  
  Q:=OpenSQL('SELECT TOP 500 * FROM SUSPECTS WHERE RSU_FERME = "-" '+SqlOrigine+' ORDER BY RSU_SUSPECT', True) ;
  while (not Q.EOF ) and not BStop do
  begin
    NbSuspect := Q.recordcount;
{$IFDEF AGL550B}
    MoveCurProgressForm('Lecture de '+IntToStr(NbSuspect)+' suspects');
{$ELSE}
    if QRPFStop.Canceled then  begin BStop := true; break; end;
    QRPFStop.Text := 'Lecture de '+IntToStr(NbSuspect)+' suspects';
{$ENDIF}
    TobsuspectTous.LoadDetailDB('SUSPECTS', '', '', Q, False, True);
    Ferme(Q) ;
    NbSuspect := TobsuspectTous.Detail.Count;
    if (ModeCreation = 'auto') then
    begin
      Numchrono := GetParamsoc('SO_GCCOMPTEURTIERS');
      CodeTiers := Numchrono;
      i_Chrono := StrToInt(Numchrono);
      i_Chrono := i_Chrono + NbSuspect;
      Numchrono :=  IntToStr(i_Chrono);
      SetParamSoc('SO_GCCOMPTEURTIERS', NumChrono) ;
    end;

    for i_ind := 0 to NbSuspect -1 do
    begin
      Tobsuspect := TobsuspectTous.Detail[i_ind] ;
      CodeSuspect := Tobsuspect.getvalue ('RSU_SUSPECT');
      if (ModeCreation = 'reprise') then  CodeTiers := CodeSuspect  else
      begin
        i_Chrono := StrToInt(CodeTiers);
        i_Chrono := i_Chrono+1;
        CodeTiers := IntToStr(i_Chrono);
      end;

      CodeAuxi := CodeTiers;
      CodeAuxi:=BourreLaDonc(CodeAuxi,fbAux);
      CreationTiers ;
      CreationTiersCompl ;
      CreationProspect ;
      CreationContact ;
      index := TOBsuspectMaj.Detail.Count;
      TOB.Create ('SUSPECT', TOBsuspectMaj, index) ;
      TOBsuspectMaj.Detail[index].Dupliquer (Tobsuspect, False, True,True);
      TOBsuspectMaj.Detail[index].putvalue('RSU_FERME','X');
      TOBsuspectMaj.Detail[index].putvalue('RSU_DATEFERMETURE',V_PGI.DateEntree);
      TOBsuspectMaj.Detail[index].putvalue('RSU_DATESUSPRO',V_PGI.DateEntree);
      TOBsuspectMaj.Detail[index].putvalue('RSU_MOTIFFERME','PRO');
      if (TOBsuspectMaj.Detail.Count >= NbFillesMax) or (i_ind=NbSuspect-1)then
      begin
{$IFDEF AGL550B}
      if Not MoveCurProgressForm('Insertion dans les tables de '+IntToStr(TOBsuspectMaj.Detail.Count)+' prospects') then break
      else NbInsertions := NbInsertions + TOBsuspectMaj.Detail.Count;
{$ELSE}
       NbInsertions := QRPFStop.Value ;
       QRPFStop.Value := QRPFStop.Value + TOBsuspectMaj.Detail.Count;
       QRPFStop.Text := 'Insertion dans les tables de '+IntToStr(TOBsuspectMaj.Detail.Count)+' prospects';
{$ENDIF}
       if (Transactions (InsertLesTobs,1)<>oeOk) then
        begin
        MessTransfert := 'Erreur Import au compteur tiers :'+CodeTiers+' et nombre suspects :'+IntToStr(NbInsertions);
        ListRecap.lines.Add (MessTransfert);
        MessageAlerte(MessTransfert );
        BStop := true; break;
        end;
{$IFDEF AGL550B}
        if Not MoveCurProgressForm('Début tiers :'+numchrono+' - Code Tiers :'+CodeTiers+' - Nombre Insertions :'+IntToStr(NbInsertions)) then break
{$ELSE}
        QRPFStop.SubText := 'Début tiers :'+numchrono+' - Code Tiers :'+CodeTiers+' - Nombre Insertions :'+IntToStr(QRPFStop.Value);
{$ENDIF}
      end;
    end;
//    if  (QRPFStop.Value >= 1500) then BStop:=true;

    Q:=OpenSQL('SELECT TOP 500 * FROM SUSPECTS WHERE RSU_FERME = "-" '+SqlOrigine+ ' ORDER BY RSU_SUSPECT', True) ;
  end;

//  MessTransfert:='Compteur final des TIERS : '+GetParamsoc('SO_GCCOMPTEURTIERS');
//  ListRecap.lines.Add (MessTransfert);

finally
  MessTransfert:='Dernier Code TIERS : '+CodeTiers;
  ListRecap.lines.Add (MessTransfert);
{$IFDEF AGL550B}
  MessTransfert:='Nombre total de Suspects transférés : '+IntToStr(NbInsertions);
  FiniMoveProgressForm ;
{$ELSE}
  MessTransfert:='Nombre total de Suspects transférés : '+IntToStr(QRPFStop.Value);
  QRPFStop.Free ;
{$ENDIF}
  ListRecap.lines.Add (MessTransfert);
  ListRecap.lines.Add ('');
  datefin := CurrentDate;
  MessTransfert:= 'Fin du traitement à '+ FormateDate('hh "h" mm', datefin)+ ' - durée :'+ FormateDate('hh "h" mm', datefin-datedebut);
  ListRecap.lines.Add (MessTransfert);
  JournalEvenement ;
  PGIBox(MessTransfert,'Fin du transfert des Suspects');
  Ferme(Q) ;

//  ListRecap.Free ;
  TobsuspectTous.free;
  TOBsuspectMaj.free;
  TobSuspectcompl.free;
  Tobtiers.free;
  Tobtierscompl.free;
  TobProspect.free;
  TobContact.free;
end;
end;

procedure TOF_SuspectBasc.OnUpdate;
var CodeOrigine : string;
    codecreation : integer;
begin
Inherited;
ListRecap := Tmemo(getcontrol('RAPPORT_BASC'));
ListRecap.lines.clear;
NumChrono := getcontroltext('COMPTEURDEBUT');
CodeOrigine := GetControlText('RSU_ORIGINETIERS');
codecreation := THRadioGroup(GetControl('BASCCREATION')).ItemIndex ;
ModeCreation := THRadioGroup(GetControl('BASCCREATION')).values[codecreation];
LmodeCreation := THRadioGroup(GetControl('BASCCREATION')).items[codecreation];
SqlOrigine := '';
if (trim (CodeOrigine) <> '') then
   SqlOrigine := ' AND RSU_ORIGINETIERS="'+CodeOrigine+'" ';
if ( (ModeCreation='debut') and (trim(NumChrono)='') ) then
begin
  PGIBox('Vous devez saisir le code début prospect',Ecran.Caption);
  exit;
end;
if ( (ModeCreation='debut') and Presence('TIERS','T_TIERS',trim(NumChrono)) ) then
begin
  PGIBox('Une fiche existe déjà avec ce code prospect ',Ecran.Caption);
  exit;
end;

if ( (ModeCreation = 'auto') and (GetParamsoc('SO_GCNUMTIERSAUTO') = FALSE) ) then
begin
  PGIBox('Pas d''attribution automatique du code client',Ecran.Caption);
  exit;
end;
BasculeSuspect;
end;

initialization
RegisterClasses([TOF_SuspectBasc]);
end.
