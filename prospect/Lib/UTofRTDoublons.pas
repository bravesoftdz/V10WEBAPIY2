unit UTofRTDoublons;

interface
uses UTOF,GCMZSUtil, UtilGC,
{$IFDEF GIGI}
     EntGc,
{$ENDIF}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
{$IFDEF EAGLCLIENT}
     eMul,
{$ELSE}
     Mul,
{$ENDIF}
     Forms,Classes,M3FP,HMsgBox,Controls,HEnt1,HCtrls,Hstatus;

Type
{$ifdef AFFAIRE}
                //mcd 11/05/2006 12940  pour faire affectation depuis ressource si paramétré
     TOF_RTDoublons = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_RTDoublons = Class(TOF)
{$endif}
       public
       Argument : string;
       procedure OnArgument (stArgument : string); override;
       procedure OnClose ; override ;
       procedure DoublonsProspect;
       procedure OnLoad ; override;
     end;
     TOF_RTFermeSuspect = Class(TOF)
       public
       Argument : string;
       procedure OnArgument (stArgument : string); override;
       procedure FermeSuspect(Motif : string);
       procedure MajSuspect(Motif : string);
     end;

const
	// libellés des messages
  TexteMessage: array[1..1] of string  = (
          {1}       'Le traitement de dédoublonnage est effectué sur le champ téléphone formaté.' + #13 + ' Il est donc nécessaire d''effectuer le traitement ''Initialisation clé téléphone tiers''' + #13 + 'pour que le dédoublonnage soit complet.'
                          );

procedure AGLDoublonsProspect(parms : array of variant; nb : integer);

implementation

procedure TOF_RTDoublons.OnArgument (stArgument : string);
var ListeChamp,ListeLibelle,ListeChamp1,ListeLibelle1: TStrings;
begin
{$ifdef AFFAIRE}
inherited;	//mcd 11/05/2006 pour passer dans Aftraducaffaire
{$endif}
ListeChamp := TStringList.Create ;
ListeLibelle := TStringList.Create;
ListeChamp1 := TStringList.Create ;
ListeLibelle1 := TStringList.Create;
ListeChamp.add('');                  ListeLibelle.add (TraduireMemoire('<<Aucun>>'));
ListeChamp.add('T_LIBELLE');	     ListeLibelle.add (TraduireMemoire('Raison sociale'));
ListeChamp.add('T_ADRESSE1');	     ListeLibelle.add (TraduireMemoire('Adresse 1'));
ListeChamp.add('T_ADRESSE2');	     ListeLibelle.add (TraduireMemoire('Adresse 2'));
ListeChamp.add('T_ADRESSE3');	     ListeLibelle.add (TraduireMemoire('Adresse 3'));
ListeChamp.add('T_CODEPOSTAL');	     ListeLibelle.add (TraduireMemoire('Code postal'));
ListeChamp.add('T_VILLE');	     ListeLibelle.add (TraduireMemoire('Ville'));
ListeChamp.add('T_DIVTERRIT');	     ListeLibelle.add (TraduireMemoire('Division territoriale'));
ListeChamp.add('T_PAYS');	     ListeLibelle.add (TraduireMemoire('Pays'));
ListeChamp.add('T_SIRET');	     ListeLibelle.add (TraduireMemoire('Code SIRET'));
ListeChamp.add('T_APE');	     ListeLibelle.add (TraduireMemoire('Code APE'));
ListeChamp.add('T_CLETELEPHONE');    ListeLibelle.add (TraduireMemoire('Téléphone '));
ListeChamp.add('T_FAX');	     ListeLibelle.add (TraduireMemoire('Fax'));
ListeChamp.add('T_TELEX');	     ListeLibelle.add (TraduireMemoire('Telex'));
ListeChamp.add('T_SECTEUR');	     ListeLibelle.add (TraduireMemoire('Secteur activité'));
{$ifndef GIGI}
ListeChamp.add('T_ZONECOM');	     ListeLibelle.add (TraduireMemoire('Zone commerciale'));
{$endif}
ListeChamp.add('T_SOCIETE');	     ListeLibelle.add (TraduireMemoire('Société'));
ListeChamp.add('T_RVA');	     ListeLibelle.add (TraduireMemoire('Site internet'));
ListeChamp.add('T_CODEIMPORT');	     ListeLibelle.add (TraduireMemoire('Code application amont'));
ListeChamp.add('T_PRESCRIPTEUR');    ListeLibelle.add (TraduireMemoire('Prescripteur'));
ListeChamp.add('T_ORIGINETIERS');    ListeLibelle.add (TraduireMemoire('Code origine'));
ListeChamp.add('T_SOCIETEGROUPE');   ListeLibelle.add (TraduireMemoire('Groupe Maison mère'));
ListeChamp.add('T_PHONETIQUE');      ListeLibelle.add (TraduireMemoire('Libellé phonétique'));
ListeChamp.add('T_DOMAINE');         ListeLibelle.add (TraduireMemoire('Domaine d''activité'));

ListeChamp1.add('');                 ListeLibelle1.add (TraduireMemoire('<<Aucun>>'));
ListeChamp1.add('T_AUXILIAIRE');     ListeLibelle1.add (TraduireMemoire('Code tiers')+' 1');
ListeChamp1.add('T_AUXILIAIRE_1');   ListeLibelle1.add (TraduireMemoire('Code tiers')+' 2');
ListeChamp1.add('T_LIBELLE');        ListeLibelle1.add (TraduireMemoire('Raison sociale')+' 1');
ListeChamp1.add('T_LIBELLE_1');      ListeLibelle1.add (TraduireMemoire('Raison sociale')+' 2');
ListeChamp1.add('T_PHONETIQUE');     ListeLibelle1.add (TraduireMemoire('Libellé phonétique')+' 1');
ListeChamp1.add('T_PHONETIQUE_1');   ListeLibelle1.add (TraduireMemoire('Libellé phonétique')+' 2');
{$ifndef GIGI}
ListeChamp1.add('T_ZONECOM');        ListeLibelle1.add (TraduireMemoire('Zone commerciale')+' 1');
ListeChamp1.add('T_ZONECOM_1');      ListeLibelle1.add (TraduireMemoire('Zone commerciale')+' 2');
{$endif}

THValcomboBox(GetControl('LISTECHAMPS')).Items.Assign(ListeLibelle);
THValcomboBox(GetControl('LISTECHAMPS')).Values.Assign(ListeChamp);
THValcomboBox(GetControl('LISTECHAMPS1')).Items.Assign(ListeLibelle1);
THValcomboBox(GetControl('LISTECHAMPS1')).Values.Assign(ListeChamp1);
ListeChamp.free;ListeLibelle.Free;
ListeChamp1.free;ListeLibelle1.Free;

//V_PGI.ExtendedFieldSelection:='2';
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    end;
  end;
{$Ifdef GIGI}
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
 SetControlText ('TT_JURIDIQUE','Forme juridique');
{$endif}
end;

procedure TOF_RTDoublons.OnClose;
begin
V_PGI.ExtendedFieldSelection:='' ;
end;

procedure TOF_RTDoublons.OnLoad;
begin
if (GetControl('LISTECHAMPS') <> nil)  then
   begin
   if GetControlText('LISTECHAMPS') = 'T_CLETELEPHONE' then
     begin
     if ExisteSql ('SELECT T_TIERS FROM TIERS WHERE (T_NATUREAUXI = "CLI" OR T_NATUREAUXI = "PRO") AND T_CLETELEPHONE = "" AND T_TELEPHONE <> ""') then
          PGIInfo(TexteMessage[1],'');
     end;
   end;
end;

/////////////// ModificationParLotDesTiers //////////////
procedure TOF_RTDoublons.DoublonsProspect;
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
V_PGI.ExtendedFieldSelection:='2';
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin MessageAlerte('Aucun élément sélectionné'); V_PGI.ExtendedFieldSelection:=''; exit; end;

TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe; TheModifLot.Q := F.Q;
TheModifLot.NatureTiers := Argument;
TheModifLot.Titre := Ecran.Caption;
TheModifLot.TableName:='TIERS;TIERSCOMPL';
TheModifLot.FCode := 'T_AUXILIAIRE';
TheModifLot.Nature := 'GC';
TheModifLot.FicheAOuvrir := 'GCTIERS';
// mng 18-11-03 pour restriction fiche
TheModifLot.stAbrege:='T_NATUREAUXI';
ModifieEnSerie(TheModifLot, Parametrages) ;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLDoublonsProspect(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_RTDoublons) then TOF_RTDoublons(TOTOF).DoublonsProspect else exit;
end;

{ TOF_RTFermeSuspect }

procedure TOF_RTFermeSuspect.OnArgument(stArgument: string);
begin
  inherited;
{$Ifdef GIGI}
 if (GetControl('RSU_REPRESENTANT') <> nil) then  SetControlVisible('RSU_REPRESENTANT',false);
 if (GetControl('TRSU_REPRESENTANT') <> nil) then  SetControlVisible('TRSU_REPRESENTANT',false);
 if (GetControl('RSU_ZONECOM') <> nil) then  SetControlVisible('RSU_ZONECOM',false);
 if (GetControl('TRSU_ZONECOM') <> nil) then  SetControlVisible('TRSU_ZONECOM',false);
{$endif}
end;

procedure TOF_RTFermeSuspect.FermeSuspect(Motif : string);
var  i : integer;

begin

  With TFMul(Ecran) do
  begin
    if(FListe.NbSelected=0)and(not FListe.AllSelected) then
      begin
      MessageAlerte('Aucun élément sélectionné'); V_PGI.ExtendedFieldSelection:=''; exit;
      end;

    try
    BeginTrans;
    if FListe.AllSelected then
    begin
      InitMove(Q.RecordCount,'');
      Q.First;
      while Not Q.EOF do
        begin
        MoveCur(False);
        MajSuspect(Motif);
        Q.NEXT;
        end;
      FListe.AllSelected:=False;
    end else
    begin
      InitMove(FListe.NbSelected,'');
      for i:=0 to FListe.NbSelected-1 do
          BEGIN
          FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
          Q.TQ.Seek(FListe.Row-1) ;
{$ENDIF}
          MajSuspect(Motif);
          MoveCur(False);
          END ;
      FListe.ClearSelected;
    end;
    FiniMove;
    Q.EnableControls ;
    CommitTrans;
    except
      RollBack;
    end;
  end;
end;

procedure TOF_RTFermeSuspect.MajSuspect(Motif : string);
var Suspect : string;
Begin
Suspect :=TFmul(Ecran).Q.FindField('RSU_SUSPECT').asstring ;
ExecuteSQL('UPDATE SUSPECTS SET RSU_FERME = "X", '+
                               'RSU_MOTIFFERME="'+motif+'", '+
                               'RSU_DATEFERMETURE="'+USDateTime(NowH)+'" '+
                               'WHERE RSU_SUSPECT="'+Suspect+'" ');
end;

/////////////////// Appellé par le bouton Valider /////////////////
procedure AGLFermeSuspect(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then TOTOF:=TFMul(F).LaTOF else exit;

if (TOTOF is TOF_RTFermeSuspect) then TOF_RTFermeSuspect(TOTOF).FermeSuspect(Parms[1]) else exit;
end;



Initialization
registerclasses([TOF_RTDoublons]) ;
registerclasses([TOF_RTFermeSuspect]) ;
RegisterAglProc('DoublonsProspect',TRUE,2,AGLDoublonsProspect);
RegisterAglProc('RTFermeSuspect',TRUE,1,AGLFermeSuspect);
end.
