unit UtilSuspect;

interface

uses  ComCtrls,Sysutils, HCtrls,StdCtrls, HEnt1, Controls,
{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
      Forms,M3FP,HMsgBox,Classes, HStatus,ed_tools,UTob,utof,UtilGC,UtilRT,ParamSoc,Ent1,EntGC;

Const
     NbFillesMax : integer = 20;

type
    T_SuspectToProspect = Class(TObject)
    procedure InsertLesTobs;
    end;

Procedure RTSuspectToProspect ;

implementation

var
      Tobsuspect,TobsuspectTous,TOBsuspectMaj,TobSuspectcompl,Tobtiers,Tobtierscompl,TobProspect,TobContact : Tob;
      CodeAuxi,CodeTiers,CodeSuspect,NumChrono:String;
      NbSuspect,LgCode : Integer;
      NbSuspectTotal : longint;
      ListRecap : TStringList;
      BStop : boolean;

Procedure T_SuspectToProspect.InsertLesTobs ;
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

procedure TOBCopieChamp(FromTOB, ToTOB : TOB);
var iChamp , iTableLigne, i_pos,i_ind1: integer;
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
      Delete (St, 1, i_pos-1) ;
      FieldNameTo := PrefixeTo + St ;
      if ToTOB.FieldExists(FieldNameTo) then
        ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom)) ;
    end;
end;

Procedure  CreationTiers ;
Var TOBL : TOB ;
    i_Chrono : integer;
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
  if (TOBL.GetValue ('T_PAYS')='') then TOBL.PutValue ('T_PAYS', GetParamSocSecur('SO_GcTiersPays','FRA'));
  TOBL.PutValue ('T_CREERPAR', 'GEN');
  TOBL.PutValue ('T_DATEOUVERTURE', V_PGI.DateEntree);
  TOBL.PutValue ('T_DEVISE', V_PGI.devisepivot);
  TOBL.PutValue ('T_LETTRABLE', 'X');
  TOBL.PutValue ('T_REGIMETVA',VH^.RegimeDefaut);
  if VH_GC.GCDefFactureHT then TOBL.PutValue ('T_FACTUREHT','X') ;
  TOBL.PutValue ('T_COLLECTIF',VH^.DefautCli);
  TOBL.PutValue ('T_MODEREGLE',GetParamSocSecur('SO_GcModeRegleDefaut','001'));
  TOBL.PutValue ('T_TVAENCAISSEMENT', GetParamSocSecur('SO_TVAENCAISSEMENT','TE'));
{  if (Getparamsoc('SO_GCVENTEEURO')) then TOBL.PutValue ('T_EURODEFAUT','X')
                                    else TOBL.PutValue ('T_EURODEFAUT','-');  }
  TOBL.PutValue ('T_EURODEFAUT','X');
  TOBL.PutValue ('T_CONFIDENTIEL','0') ;
  TOBL.PutValue ('T_PUBLIPOSTAGE', 'X');
  TOBL.PutValue ('T_PARTICULIER','-');
end;

Procedure  CreationTiersCompl ;
Var TOBL : TOB ;
begin
  TOBL:=Tob.create('TIERSCOMPL',Tobtierscompl,-1);
  TOBL.InitValeurs;
  if TOBSuspectCompl.selectDB ('"'+CodeSuspect+'"',Nil,False) then
     RTCorrespondSuspectProspect('RSC',TOBSuspectCompl, TOBL);
  TOBL.PutValue ('YTC_TIERS',CodeTiers);
  TOBL.PutValue ('YTC_AUXILIAIRE',CodeAuxi);
end;

Procedure  CreationProspect ;
Var TOBL : TOB ;
begin
  TOBL:=Tob.create('PROSPECTS',TobProspect,-1);
  TOBL.InitValeurs;
  TOBL.PutValue ('RPR_AUXILIAIRE',CodeAuxi);
end;

Procedure  CreationContact ;
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
  end;
end;

procedure SuspectJournalEvenement ;
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
ListRecap.Add ('');
if BStop = True then
    begin
    TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
    ListRecap.add ('Interrompue par l''utilisateur');
    end else
    begin
    if V_PGI.IoError = oeOk then
        begin
        TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
        ListRecap.Add('Traitement terminé');
        end else
        begin
        TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
        ListRecap.Add ('Traitement non terminé');
        end;
    end;
TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Text);
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
end;


Procedure RTSuspectToProspect ;
var  Q : TQuery ;
     i_ind,index,i_Chrono : integer;
     QRPFStop : TQRProgressForm;
     BSTop : boolean;
     MessTransfert : string;
     TransactSuspect :T_SuspectToProspect;
begin
  if PGIAsk('Voulez-vous transférer les Suspects en Prospects ?', 'Suspects en Prospects') <> mrYes then exit;

  QRPFStop := DebutProgressForm (Nil,'Traitement en cours...','', 0, true, true) ;
  Q:=OpenSQL('SELECT count(RSU_SUSPECT) FROM SUSPECTS WHERE RSU_FERME = "-" and RSU_ORIGINETIERS="048"', True) ;
  if Not Q.EOF then NbSuspectTotal:=Q.Fields[0].AsInteger;
  Ferme(Q) ;
  QRPFStop.MaxValue := NbSuspectTotal;
  LgCode:=VH^.Cpta[fbAux].Lg ;
  ListRecap:=TStringList.Create;
  BStop := False;
try
  TobsuspectTous := TOB.Create('Les Suspects', Nil, -1) ;
  Tobtiers := TOB.Create('les tiers', nil, -1);
  Tobtierscompl := TOB.Create('complementaires tiers', nil, -1);
  TobProspect := TOB.Create('complementaires prospects', nil, -1);
  TobContact := TOB.Create('les contacts', nil, -1);
  TOBsuspectMaj := TOB.Create('Maj Suspect', nil, -1);
  TobSuspectcompl := TOB.Create('SUSPECTSCOMPL', nil, -1);
  QRPFStop.Text := 'Transfert des suspects';
  QRPFStop.Value := 0;
  ListRecap.Add ('Transfert des suspects');
  ListRecap.Add ('');
  MessTransfert:='Compteur initial des TIERS : '+GetParamsocSecur('SO_GCCOMPTEURTIERS','0');
  ListRecap.Add (MessTransfert);
  TransactSuspect :=T_SuspectToProspect.Create;

  Q:=OpenSQL('SELECT TOP 100 * FROM SUSPECTS WHERE RSU_FERME = "-" and RSU_ORIGINETIERS="048"', True) ;
  while (not Q.EOF ) and not BStop do
  begin
    if QRPFStop.Canceled then  begin BStop := true; break; end;
    TobsuspectTous.LoadDetailDB('SUSPECTS', '', '', Q, False, True);
    Ferme(Q) ;
    NbSuspect := TobsuspectTous.Detail.Count;
    Numchrono := GetParamsocSecur('SO_GCCOMPTEURTIERS','0');
    CodeTiers := Numchrono;
    i_Chrono := StrToInt(Numchrono);
    i_Chrono := i_Chrono + NbSuspect;
    Numchrono :=  IntToStr(i_Chrono);
    SetParamSoc('SO_GCCOMPTEURTIERS', NumChrono) ;

    for i_ind := 0 to NbSuspect -1 do
    begin
      Tobsuspect := TobsuspectTous.Detail[i_ind] ;
      CodeSuspect := Tobsuspect.getvalue ('RSU_SUSPECT');
      i_Chrono := StrToInt(CodeTiers);
      i_Chrono := i_Chrono+1;
      CodeTiers := IntToStr(i_Chrono);
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
       NbInsertions := QRPFStop.Value ;
       QRPFStop.Value := QRPFStop.Value + TOBsuspectMaj.Detail.Count;
       if (Transactions (TransactSuspect.InsertLesTobs,1)<>oeOk) then
        begin
        MessTransfert := 'Erreur Import au compteur tiers :'+numchrono+' et nombre suspects :'+IntToStr(NbInsertions);
        ListRecap.Add (MessTransfert);
        MessageAlerte(MessTransfert );
        BStop := true; break;
        end;
      end;
    end;
    Q:=OpenSQL('SELECT TOP 100 * FROM SUSPECTS WHERE RSU_FERME = "-" and RSU_ORIGINETIERS="048"', True) ;
  end;

  MessTransfert:='Compteur final des TIERS : '+GetParamsocSecur('SO_GCCOMPTEURTIERS','0');
  ListRecap.Add (MessTransfert);
  MessTransfert:='Nombre de Suspects transférés : '+IntToStr(QRPFStop.Value);
  ListRecap.Add (MessTransfert);
  SuspectJournalEvenement ;
  PGIBox(MessTransfert,'Fin du transfert des Suspects');

finally
  Ferme(Q) ;
  QRPFStop.Free ;
  ListRecap.Free ;
  TobsuspectTous.free;
  TOBsuspectMaj.free;
  TobSuspectcompl.free;
  Tobtiers.free;
  Tobtierscompl.free;
  TobProspect.free;
  TobContact.free;
  TransactSuspect.free;
end;

end;

end.
