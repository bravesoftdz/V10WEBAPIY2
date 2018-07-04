{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTEVENTCHA_TOF ()
Mots clefs ... : TOF;BTEVENTCHA_TOF
*****************************************************************}
Unit BTEVENTCHA_TOF ;

Interface
                                             
Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
{$else}
     eMul, 
     uTob,
{$ENDIF}
     forms,
     HPanel,
     Vierge,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTOB,
     HTB97,
     UTOF ;

Type
  TOF_BTEVENTCHA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Action        : TActionFiche;
    AppelPlan		  : Boolean;
    //
    HDeb          : TDateTime;
    HFin          : TDateTime;
    NbHJour       : TDateTime;
    //
    OldDDeb       : TDateTime;
    OldDFin       : TDateTime;
    OldDuree      : Double;
    //
    NumEventCha   : THLabel;
    //
    Ressource     : THEdit;
    NomRessource  : THLabel;
    TypeAction    : THEdit;
    LibTypeAction : THLabel;
    //
    Affaire       : THEdit;
    Affaire0      : THEdit;
    Affaire1      : THEdit;
    Affaire2      : THEdit;
    Affaire3      : THEdit;
    Avenant       : THEdit;
    //
    TNumPhase     : TLabel;
    NumPhase      : THEdit;
    LibPhase      : TLabel;
    //
    LibelleAction : THEdit;
    //
    DateDeb       : THEdit;
    DateFin       : THEdit;
    NbHeure       : THEdit;
    //
    BTSelect      : TToolbarButton97;
    BTEfface      : TToolbarButton97;
    BtSelect2     : TToolbarButton97;
    BtEfface2     : TToolbarButton97;
    BtEfface3     : TToolbarButton97;
    BTDelete      : TToolBarButton97;
    //
    BPlanning     : TToolBarButton97;
    BInsert       : TToolBarButton97;
    BValider      : TToolBarButton97;
    //
    TobEventCha   : Tob;
    TobCalendrier : TOB;
    TobJFerie     : TOB;
    //
    StSQL         : String;
    CodeClient    : String;
    //
    FF            : String;
    //
    QQ            : TQuery;
    //
    HeureDeb      : TDateTime;
    HeureFin      : TDateTime;
    //
    procedure AffaireOnExit(Sender: TObject);
    procedure AffEffaceOnClick(Sender: TObject);
    procedure AffSelectOnClick(Sender: TObject);
    procedure CalculNoCompteur;
    procedure ChargeEcranWithLaTob;
    procedure ChargeInfoAffaire;
    procedure ChargeInfoPhase;
    procedure ChargeInfoTypeAction;
    procedure ChargeZoneEcran;
    procedure ChargeZoneTOB;
    procedure Controlechamp(Champ, Valeur: String);
    procedure ControleRessource;
    function  ControleSaisie: Boolean;
    procedure CreateTOB;
    procedure DateDebOnExit(Sender: TObject);
    procedure DateFinOnExit(Sender: TObject);
    procedure DestroyTOB;
    procedure GetObjects;
    procedure NBHeureOnExit(Sender: TObject);
    procedure NumPhaseOnChange(Sender: TObject);
    procedure NumPhaseOnExit(Sender: TObject);
    procedure PhaseEffaceOnClick(Sender: TObject);
    procedure RazZoneEcran;
    procedure RechEquipeCreation;
    procedure RechEquipeModification;
    procedure RessourceEffaceOnClick(Sender: TObject);
    procedure RessourceOnExit(Sender: TObject);
    procedure RessourceSelectOnClick(Sender: TObject);
    procedure SetScreenEvents;
    procedure TypeActionOnExit(Sender: TObject);
    procedure TypeActionOnChange(Sender: TObject);


  end ;

const
  TexteMessage: array[1..12] of string  = (
          {1}   'Le Type d''action est Obligatoire'
          {2},  'Le code Affaire est Obligatoire'
          {3},  'Le code Ressource est Obligatoire'
          {4},  ''
          {5},  'La suppression a échoué'
          {6},  'Création impossible, numéro non déterminé !'
          {7},  'Le Type d''action est inexistant'
          {8},  'Le code Affaire est inexistant'
          {9},  'Le code Ressource est inexistant'
          {10}, 'Le nombre d''heures est obligatoire si votre événement est réalisé'
          {11}, 'La date de début est supérieure à la date de fin'
          {12}, 'la date de fin est inférieure à la date de début'
                                          );

Implementation

uses AffaireUtil,
     AGLInitGC,
     FactUtil,
     UtilsParc,
     HeureUtil,
     UtilConso,
     ParamDBG,
     GerePiece,
     ParamSoc,
     DateUtils,
     UDateUtils,
     MsgUtil,
     UtilActionPlanning,
     UtofRessource_Mul,
     BTPUtil;

procedure TOF_BTEVENTCHA.OnNew ;
begin
  Inherited ;

  RazZoneEcran;

  if LaTOB <> nil then ChargeEcranwithLaTob;

  if (AppelPlan = False) Or (LaTOB = nil) then
  begin
    DateDeb.Text := DateToStr(Now);
    Datefin.Text := DateToStr(Now+1);
  end;
  //
  if ressource.Text <> '' then
  begin
    BtSelect2.visible     := False;
    BtEfface2.visible     := False;
    ressource.Enabled     := False;
  end;

  //Recherche du type d'action par défaut
  StSQL := 'SELECT BTA_BTETAT FROM BTETAT WHERE BTA_TYPEACTION="PCA" AND BTA_DEFAUT="X"';
  QQ    := OpenSQL(StSQL, False,-1,'',True);
  if not QQ.eof then
  begin
    TypeAction.text := QQ.Findfield('BTA_BTETAT').AsString;
    ChargeInfoTypeAction;
  end;
  Ferme(QQ);

  if Affaire.Text <> '' then
  begin
    if Ressource.Text = '' then
    begin
      BtSelect.Visible        := False;
      BtEfface.Visible        := False;
      BtEfface3.visible       := False;
      Affaire.Enabled         := False;
    end;
  end;

  if NumPhase.Text <> '0' then
  begin
    ChargeInfoPhase;
    if Ressource.Text = '' then
    begin
      BtSelect.Visible        := False;
      BtEfface.Visible        := False;
      Affaire.Enabled         := False;
      NumPhase.Enabled        := False;
      BtEfface3.Visible       := False;
    end;
  end;
  //

end;

procedure TOF_BTEVENTCHA.OnDelete ;
begin
  Inherited ;

  //Suppression évènement chantier
  if PGIAsk('Confirmez-vous la suppression de cet évènement ?', 'Saisie Evénement') = Mryes then
  begin
    StSQL := 'DELETE BTEVENTCHA WHERE BEC_IDEVENTCHA=' + NumEventcha.caption;
    if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[4], 'Evénement Chantier/Ressource');

    OnClose;
  end;

end ;

procedure TOF_BTEVENTCHA.OnUpdate ;
begin
  Inherited ;

  if ControleSaisie then
  begin
    Ecran.ModalResult := 1;

    //si la TOB est vide on fait rien et surtout cela signifie qu'on a un gros problème
    if TOBeventCha = nil then Exit;

    If Action = TaCreat then CalculNoCompteur;

    if  NumEventCha.Caption = '0' then exit;

    //mise à jour de la table des Evénements Parc/Matériel
    ChargeZoneTOB;

    //On met à jour la table des événements
    TOBeventCha.SetAllModifie(true);
    TOBeventCha.InsertOrUpdateDB(true);

    //Gestion de l'équipe
    if (TobEventCha.getvalue('BEC_EQUIPERESS') <> '') then
    begin
      If Action = TaCreat       then RechEquipeCreation
      else if Action = TaModif  then RechEquipeModification
    end;

    If (AppelPlan) and (LaTob <> nil) then
    begin
      LaTob.PutValue('RETOUR',          1);
      LaTob.PutValue('BEC_AFFAIRE',   Affaire.Text);
      LaTOB.PutValue('BEC_NUMPHASE',  NumPhase.text);
      LaTOB.PutValue('BEC_LIBPHASE',  LibPhase.Caption);
    end;
  end
  else
  begin
    If (AppelPlan) and (LaTob <> nil) then
    begin
      LaTob.PutValue('RETOUR',            0);
      LaTob.PutValue('BEC_AFFAIRE',       '');
      LaTOB.PutValue('BEC_NUMPHASE',      '');
      LaTOB.PutValue('BEC_LIBPHASE',      '');
      exit;
    end;
    Ecran.ModalResult := 0;
  end;

end;


Function TOF_BTEVENTCHA.ControleSaisie : Boolean;
begin

  Result := False;

  if not TypeActionExist(TypeAction.text, ' AND BTA_TYPEACTION="PCA"') then
  begin
    PGIError(TexteMessage[7], 'Erreur Saisie');
    TypeAction.SetFocus;
    exit;
  end;

  if not RessourceExist(Ressource.text, ' AND ARS_TYPERESSOURCE in ("INT", "SAL", "LOC", "ST")') then
  begin
    PGIError(TexteMessage[9], 'Erreur Saisie');
    Ressource.SetFocus;
    exit;
  end;

  if not AffaireExist(Affaire.text, ' AND AFF_AFFAIRE0="A"') then
  begin
    PGIError(TexteMessage[8], 'Erreur Saisie');
    Affaire.SetFocus;
    exit;
  end;

  if StrToDate(DateDeb.text) > StrToDate(Datefin.text) then
  begin
    PGIError(TexteMessage[11], 'Erreur Saisie');
    DateDeb.SetFocus;
    exit;
  end;

  if StrToDate(DateFin.text) < StrToDate(DateDeb.text) then
  begin
    PGIError(TexteMessage[12], 'Erreur Saisie');
    DateFin.SetFocus;
    exit;
  end;

  Result := true;

end ;

Procedure TOF_BTEVENTCHA.OnLoad ;
begin
  Inherited ;

  if NumEventCha.caption = '' then exit;

  StSQL := 'SELECT * FROM BTEVENTCHA WHERE BEC_IDEVENTCHA=' + NumEventCha.caption;
  QQ := OpenSQL(StSQL, True);

  If not QQ.Eof then
  begin
    TOBeventCha.SelectDB('BTEVENTCHA', QQ);
    ChargeZoneEcran;
  end;

  ferme(QQ);

end ;

procedure TOF_BTEVENTCHA.OnArgument (S : String ) ;
var x       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  AppelPlan := False;
  //
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  FF:='#';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for x:=1 to V_PGI.OkDecV-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  //Gestion des évènement des zones écran
  SetScreenEvents;
  //
  //Heure de début de journée, heure de fin de journée et nombre d'heure par jour
  //
  //Chargement de la table des jours fériés....
  ChargeJoursFeries(TobJferie);
  //
  //Heure de début de journée, heure de fin de journée et nombre d'heure par jour
  //
  Hdeb := TimeToFloat(GetParamSocSecur('SO_BTAMDEBUT', '08:30'));
  HFin := TimeToFloat(GetParamSocSecur('SO_BTPMFIN',   '17:30'));
  NbHJour := ChargeParametreTemps;
  //
  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.text, False);
  //
  TypeAction.plus := 'BTA_TYPEACTION="PCA"';
  //
  ressource.Enabled      := True;
  Affaire.Enabled        := True;
  //
  BtSelect2.Visible      := True;
  BtEfface2.Visible      := True;
  BtEfface3.Visible      := True;
  BtSelect.Visible       := True;
  BtEfface.Visible       := True;
  //
  TNumPhase.Visible      := False;
  NumPhase.Visible       := False;
  LibPhase.Visible       := False;
  //
  if Action = Tacreat then OnNew;

end ;

procedure TOF_BTEVENTCHA.OnClose ;
begin
  Inherited ;

  DestroyTOB;

end ;

procedure TOF_BTEVENTCHA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTEVENTCHA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTEVENTCHA.GetObjects;
begin

  NumEventcha   := THLabel(Getcontrol('BEC_IDEVENTCHA'));
  //
  Ressource     := THEdit(Getcontrol('BEC_RESSOURCE'));
  NomRessource  := THLabel(Getcontrol('NOMRESSOURCE'));
  //
  TypeAction    := THEdit(Getcontrol('BTETAT'));
  LibTypeAction := THLabel(Getcontrol('LIBTYPEACTION'));
  LibelleAction := THEdit(Getcontrol('LIBACTION'));
  DateDeb       := THEdit(Getcontrol('DATEEVT'));
  DateFin       := THEdit(Getcontrol('DATEEVT_'));
  NbHeure       := THEdit(Getcontrol('DUREE'));
  //
  Affaire       := THEdit(Getcontrol('BEC_AFFAIRE'));
  Affaire0      := THEdit(Getcontrol('BEC_AFFAIRE0'));
  Affaire1      := THEdit(Getcontrol('BEC_AFFAIRE1'));
  Affaire2      := THEdit(Getcontrol('BEC_AFFAIRE2'));
  Affaire3      := THEdit(Getcontrol('BEC_AFFAIRE3'));
  Avenant       := THEdit(Getcontrol('BEC_AVENANT'));
  //
  TNumPhase     := TLabel(GetControl('FE__HLabel6'));
  NumPhase      := THEdit(Getcontrol('BEC_NUMPHASE'));
  LibPhase      := TLabel(Getcontrol('LIBPHASE'));
  //
  BTSelect      := TToolBarButton97(Getcontrol('BSELECT'));
  BTEfface      := TToolBarButton97(Getcontrol('BEFFACE'));
  BTEfface3     := TToolBarButton97(Getcontrol('BEFFACE3'));
  //
  BTSelect2     := TToolBarButton97(Getcontrol('BSELECT2'));
  BTEfface2     := TToolBarButton97(Getcontrol('BEFFACE2'));
  //
  BtDelete      := TToolBarButton97(GetCONTROL('BDELETE'));
  //
  BPlanning     := TToolBarButton97(GetCONTROL('BPLANNING'));
  BInsert       := TToolBarButton97(GetCONTROL('BINSERT'));
  BValider      := TToolBarButton97(GetCONTROL('BVALIDER'));

end;

procedure TOF_BTEVENTCHA.SetScreenEvents;
begin
  //
  Affaire1.Onexit         := AffaireOnExit;
  Affaire2.OnExit         := AffaireOnExit;
  Affaire3.OnExit         := AffaireOnExit;
  //
  NumPhase.OnExit         := NumPhaseOnExit;
  NumPhase.OnChange       := NumPhaseOnChange;
  //
  TypeAction.OnExit       := TypeActionOnExit;
  TypeAction.OnChange     := TypeActionOnChange;
  //
  Ressource.OnExit        := RessourceOnExit;
  //
  BTSelect.OnClick        := AffSelectOnClick;
  BTEfface.OnClick        := AffEffaceOnclick;
  //
  BTEfface3.OnClick       := PhaseEffaceOnclick;
  //
  BTSelect2.OnClick       := RessourceSelectOnClick;
  BTEfface2.OnClick       := RessourceEffaceOnclick;
  //
  DateDeb.OnExit          := DateDebOnExit;
  DateFin.OnExit          := DateFinOnExit;

  NBHeure.OnExit          := NBHeureOnExit;

end;

Procedure TOF_BTEVENTCHA.AffaireOnExit(Sender : TObject);
var IP      : Integer;
begin

  Affaire.text := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, CodeClient, Action, False, True, false, IP);

  if Affaire.text <> '' then  ChargeInfoAffaire;

end;

procedure TOF_BTEVENTCHA.NumPhaseOnExit(Sender : TObject);
begin

  if Numphase.text <> '' then ChargeInfoPhase;

end;

Procedure TOF_BTEVENTCHA.RessourceOnExit(Sender : TObject);
begin

  if Ressource.text <> '' then ControleRessource;

end;

Procedure TOF_BTEVENTCHA.RessourceSelectOnClick(Sender : TObject);
Var StChamps  : string;
    StWhere   : string;
    Action    : string;
begin

  StChamps  := ressource.Text;

  StWhere := 'TYPERESSOURCE=INT,SAL,LOC,ST';
 	Action := ';ACTION=RECH';

  Ressource.Text := AFLanceFiche_Rech_Ressource ('ARS_RESSOURCE='+Trim(Ressource.Text),stwhere + Action);

  //DispatchRecherche(Ressource, 3, stwhere + Action,'ARS_RESSOURCE=' + Trim(Ressource.Text), '');

  ControleRessource;


end;


Procedure TOF_BTEVENTCHA.RessourceEffaceOnClick(Sender : TObject);
begin

  Ressource.Text  := '';
  NomRessource.Caption := '...';

end;

Procedure TOF_BTEVENTCHA.ControleRessource;
Var StWhere : string;
    TobRessource : TOB;
begin

  if ressource.text = '' then ressource.SetFocus;

  TobRessource := Tob.create('RESSOURCE', nil, -1);

  StWhere := ' AND ARS_TYPERESSOURCE in ("INT", "SAL", "LOC", "ST")';

  ChargeInfoRessource(Ressource.text, StWhere, TobRessource);

  NomRessource.caption := TobRessource.GetValue('ARS_LIBELLE2') + ' ' + TobRessource.GetValue('ARS_LIBELLE');

  if NomRessource.Caption = '' then Ressource.setfocus;

  RechercheCalendrier(Ressource.text, Tobcalendrier);

  TOBeventCha.PutValue('BEC_FONCTION',    TobRessource.GetValue('ARS_FONCTION1'));
  TOBeventCha.PutValue('BEC_RESSOURCE',   TobRessource.GetValue('ARS_RESSOURCE'));
  TOBeventCha.PutValue('BEC_EQUIPERESS',  TobRessource.GetValue('ARS_EQUIPERESS'));

  RechercheCalendrier(Ressource.text, Tobcalendrier);

  FreeAndNil(TobRessource);

end;

Procedure TOF_BTEVENTCHA.TypeActionOnExit(Sender : TObject);
begin

  ChargeInfoTypeAction;

  if TypeAction.Text = '' then TypeAction.SetFocus;

end;

procedure TOF_BTEVENTCHA.DateDebOnExit(Sender: TObject);
Var Duree : Double;
    DDeb  : TDateTime;
    DFin  : TDateTime;
begin

  DDeb    := StrToDate(DateDeb.Text);
  DFin    := StrToDate(DateFin.Text);
  Duree   := StrToFloat(NbHeure.Text);

  if DDeb > Dfin then
  begin
    PGIError(TexteMessage[11], 'Erreur Saisie');
    DateDeb.Text := DateToStr(OldDDeb);
    Datedeb.SetFocus;
    Exit;
  end;

  if DDeb = idate1900 then exit;

  if Ddeb = OldDdeb then exit;

  Duree := CalculDureeEvent(DDeb, DFin, TobJFerie, TobCalendrier);
  NbHeure.Text  := FloatToStr(Duree/60);

  OldDDeb := DDeb;

end;

procedure TOF_BTEVENTCHA.DateFinOnExit(Sender: TObject);
Var DDeb : TDateTime;
    DFin : TDateTime;
    Duree: Double;
begin

  DDeb := StrToDate(DateDeb.text);
  DFin := StrToDate(DateFin.text);
  Duree:= StrToFloat(Nbheure.Text);

  if DFin < DDeb then
  begin
    PGIError(TexteMessage[12], 'Erreur Saisie');
    DateFin.Text := DateToStr(OldDFin);
    DateFin.SetFocus;
    exit;
  end;

  if DFin = idate2099 then exit;

  if DFin = OldDFin   then exit;

  if DFin <> idate2099 then Duree := CalculDureeEvent(DDeb, DFin, TobJFerie, TobCalendrier);

  NbHeure.Text  := FloatToStr(Duree/60);

  OldDFin := DFin;

end;

procedure TOF_BTEVENTCHA.NBHeureOnExit(Sender: TObject);
Var DureeMin  : Integer;
    Duree     : Double;
    Datedebut : TDatetime;
    Datefinal : TDateTime;
    An, Mois, Jour  : Word;
    H, M, S, Msec   : Word;
begin

  DateDebut := StrToDate(DateDeb.Text) + HeureDeb;
  Datefinal := StrToDate(DateFin.Text) + HeureFin;
  Duree     := StrToFloat(NbHeure.text);
  DureeMin  := Round(Duree * 60);

  if Duree = OldDuree then exit;

  Datefinal := CalculDateFinParcMat(DateDebut, DateFinal, NbHJour, DureeMin, TobJFerie);

  Datefin.text := DateToStr(DateFinal);

  OldDuree := Duree;

end;

Procedure TOF_BTEVENTCHA.AffSelectOnClick(Sender : TObject);
Var StChamps  : String;
begin

  StChamps  := Affaire.Text;

  if GetAffaireEnteteSt(Affaire0, Affaire1, Affaire2, Affaire3, AVENANT, nil, StChamps, false, false, false, True, true,'',False,True,False,'ACP') then Affaire.text := StChamps;

  ChargeInfoAffaire;

end;
//
Procedure TOF_BTEVENTCHA.AffEffaceOnClick(Sender : TObject);
begin

  Affaire.Text  := '';
  Affaire0.Text := '';
  Affaire1.Text := '';
  Affaire2.Text := '';
  Affaire3.Text := '';
  Avenant.Text  := '';
  //
  LibelleAction.Text := '';
  //
  NumPhase.text    := '0';
  LibPhase.Caption := '...';
  NumPhase.Plus    := '';
  //
  NumPhase.Visible := False;
  LibPhase.Visible := False;
  TNumPhase.Visible:= False;

end;

Procedure TOF_BTEVENTCHA.PhaseEffaceOnClick(Sender : TObject);
begin
  //
  NumPhase.text    := '0';
  LibPhase.Caption := '...';

end;


Procedure TOF_BTEVENTCHA.Controlechamp(Champ, Valeur : String);
Var Date1 : TdateTime;
    Date2 : TDateTime;
begin

  if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if Champ='IDEVENEMENT' then NumEventCha.caption := Valeur;

  if Champ='RESSOURCE'   then Ressource.Text := Valeur;

  if Champ='AFFAIRE'     then Affaire.Text   := Valeur;
  if Champ='PHASE'       then NumPhase.Text  := Valeur;

  if Champ='DATEDEB'     then
  begin
    Date1         := StrToDate(Valeur);
    DateDeb.Text  := DateToStr(Date1);
  end;

  if Champ='DATEFIN'     then
  begin
    Date2         := StrToDate(Valeur);
    DateFin.Text  := DateToStr(Date2);
  end;

  If Champ='PLANNING' Then AppelPlan := True;

end;

Procedure TOF_BTEVENTCHA.CreateTOB;
begin

  TobEventCha := Tob.Create('BTEVENTCHA', nil, -1);

  TobCalendrier := Tob.Create('CALENDRIER', nil, -1);
  TobJFerie     := Tob.Create('JOURFERIE', nil, -1);

end;

procedure TOF_BTEVENTCHA.DestroyTOB;
begin

  FreeAndNil(TobEventCha);

  FreeAndNil(TobCalendrier);
  FreeAndNil(TobJFerie);

end;

Procedure TOF_BTEVENTCHA.ChargeInfoTypeAction;
var TOBTypeAction : TOB;
    StWhere       : String;
    Temp 	        : Double;
    DureeMin  : Integer;
    Duree     : Double;
    Datedebut : TDatetime;
    Datefinal : TDateTime;
    An, Mois, Jour  : Word;
    H, M, S, Msec   : Word;
begin

  if TypeAction.text = '' then Exit;;

  TOBTypeAction := TOB.Create('BTETAT',nil, -1);

  Stwhere := ' AND BTA_TYPEACTION="PCA"';

  StSQL := 'SELECT * FROM BTETAT WHERE BTA_BTETAT = "' + TypeAction.Text + '" ' + StWhere;
  QQ    := OpenSQL(StSQL, False);
  if QQ.eof then
  begin
    TypeAction.text := '';
    Ferme(QQ);
    Exit;
  end;

  TOBTypeAction.SelectDB('Le type Action', QQ);
  Ferme(QQ);

  LibTypeAction.caption := TOBTypeAction.GetString('BTA_LIBELLE');
  //
  if Action = TaCreat then
  begin
    if not AppelPlan then
    begin
      Temp          := TOBTypeAction.GetValue('BTA_DUREEMINI');
      NbHeure.text  := FloatToStr(Temp);
      Datefinal     := CalculDateFinParcMat(DateDebut, DateFinal, NbHJour, DureeMin, TobJFerie);
      Datefin.text  := DatetoStr(DateFinal);
      OldDDeb       := StrToDate(DateDeb.Text);
      OldDFin       := StrtoDate(Datefin.Text);
      OldDuree      := StrToFloat(NbHeure.text);
    end;
  end;
  //
  FreeAndNil(TOBTypeAction);

  ferme(QQ);

end;

procedure TOF_BTEVENTCHA.ChargeInfoAffaire;
begin

  StSQL := 'SELECT AFF_LIBELLE, AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE = "' + Affaire.Text + '"';
  QQ    := OpenSQL(StSQL, False);

  if QQ.Eof then
  begin
    LibelleAction.text := '';
    NumPhase.Plus      := '';
    TNumPhase.Visible  := False;
    NumPhase.Visible   := False;
    LibPhase.Visible   := False;
  end
  else
  begin
    ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.text, False);
    //if LibelleAction.text = '' then
    LibelleAction.text := QQ.findfield('AFF_LIBELLE').AsString;
    CodeClient := QQ.findfield('AFF_TIERS').AsString;
    //Contrôle si des phase existent sur le chantier
    StSQL := 'SELECT GL_LIBELLE FROM LIGNE WHERE GL_NATUREPIECEG="PBT" AND GL_TYPELIGNE="DP1" AND GL_AFFAIRE="' + Affaire.Text + '"';
    If ExisteSQL(StSQL) then
    begin
      TNumPhase.Visible  := True;
      NumPhase.Visible   := True;
      LibPhase.Visible   := True;
      NumPhase.Plus := 'AND GL_AFFAIRE="' + Affaire.Text + '"';
    end;
  end;

  ferme(QQ);

end;

procedure TOF_BTEVENTCHA.ChargeInfoPhase;
begin

  StSQL := 'SELECT GL_LIBELLE FROM LIGNE WHERE GL_NATUREPIECEG="PBT" AND GL_NUMORDRE=' + NumPhase.text + ' AND GL_AFFAIRE="' + Affaire.Text + '"';
  QQ := OpenSQL(STSQL, False, -1,'', True);
  if QQ.Eof then
    LibPhase.Caption := '...'
  else
    LibPhase.Caption := QQ.findfield('GL_LIBELLE').AsString;

  Ferme(QQ);

end;

procedure TOF_BTEVENTCHA.RazZoneEcran;
begin

  Ressource.text        := '';
  NomRessource.caption  := '...';
  //
  TypeAction.text       := '';
  LibTypeAction.caption := '...';
  //
  Affaire.text          := '';
  Affaire0.text         := '';
  Affaire1.text         := '';
  Affaire2.text         := '';
  Affaire3.text         := '';
  Avenant.text          := '';
  //
  NumPhase.Text         := '';
  LibPhase.Caption      := '...';
  //
  LibelleAction.text    := '';
  //
  DateDeb.Text          := DateToStr(now);
  DateFin.Text          := DateToStr(idate2099);
  //
  HeureDeb              := GetDebutMatinee;
  HeureFin              := GetFinApresMidi;
  //
  OldDDeb               := StrToDate(DateDeb.Text);
  OldDFin               := StrtoDate(Datefin.Text);
  OldDuree              := 1;
  //
  NbHeure.text          := '1';
  //
  BtDelete.visible      := False;
  //
  TNumPhase.Visible  := False;
  NumPhase.Visible   := False;
  LibPhase.Visible   := False;
end ;

Procedure TOF_BTEVENTCHA.ChargeZoneEcran;
Var Temp      : Integer;
begin

  NumEventCha.Caption := TOBeventCha.GetString('BEC_IDEVENTCHA');
  //
  DateDeb.Text        := DateToStr(TOBeventCha.GetValue('BEC_DATEDEB'));
  DateFin.Text        := DateToStr(TOBeventCha.GetValue('BEC_DATEFIN'));
  //
  HeureDeb            := TobEventCha.GetValue('BEC_HEUREDEB');
  HeureFin            := TobEventCha.GetValue('BEC_HEUREFIN');
  //
  Ressource.Text      := TOBeventCha.GetString('BEC_RESSOURCE');
  ControleRessource;
  //
  TypeAction.Text     := TOBeventCha.GetString('BEC_BTETAT');
  ChargeInfoTypeAction;
  //
  Affaire.Text        := TOBeventCha.GetString('BEC_AFFAIRE');
  ChargeInfoAffaire;
  //
  NumPhase.Text       := TOBeventCha.GetString('BEC_NUMPHASE');
  LibPhase.Caption    := TOBeventCha.GetString('BEC_LIBPHASE');
  if LibPhase.caption = '' then ChargeInfoPhase;
  //
  LibelleAction.Text  := TOBeventCha.GetString('BEC_LIBACTION');
  //
  OldDDeb             := StrToDate(DateDeb.Text);
  OldDFin             := StrtoDate(DateFin.Text);

  //Chargement de la durée mini
  Temp                := TOBeventCha.GetValue('BEC_DUREE');
  NbHeure.text        := FloatToStr(Temp/60);

end;

procedure TOF_BTEVENTCHA.CalculNoCompteur;
Var NumAff : Integer;
begin

  if not GetNumCompteur('BEC',iDate1900, NumAff) then
  begin
    PGIError(TexteMessage[7], 'Evènement Ressource/Chantier');
    OnClose;
  end
  else NumEventCha.Caption := IntToStr(NumAff);

end;

Procedure TOF_BTEVENTCHA.ChargeZoneTOB;
Var DureeMin  : Integer;
    Duree     : Double;
    Datedebut : TDatetime;
    Datefinal : TDateTime;
    An, Mois, Jour  : Word;
    H, M, S, Msec   : Word;
begin

  TOBeventCha.PutValue('BEC_IDEVENTCHA',    NumEventCha.Caption);
  TOBeventCha.PutValue('BEC_RESSOURCE',     Ressource.Text);
  TOBeventCha.PutValue('BEC_BTETAT',        TypeAction.Text);
  TOBeventCha.PutValue('BEC_AFFAIRE',       Affaire.Text);
  TOBeventCha.PutValue('BEC_TIERS',         CodeClient);
  if NumPhase.text = '' then NumPhase.Text := '0';
  TOBeventCha.PutValue('BEC_NUMPHASE',      NumPhase.Text);
  TOBeventCha.PutValue('BEC_LIBPHASE',      LibPhase.Caption);
  TOBeventCha.PutValue('BEC_ARTICLE',       '');
  TOBeventCha.PutValue('BEC_LIBACTION',     LibelleAction.Text);

  if Action = TaCreat then
  begin
    TOBeventCha.PutValue('BEC_IDAFFECT',    NumEventCha.Caption);
  end;

  TOBeventCha.PutValue('BEC_RESPRINCIPALE', 'X');
  TOBeventCha.PutValue('BEC_EQUIPESEP',     '-');
  //
  //Calcul de l'heure de fin en fonction du nombre d'heure saisie et de la date de début
  //
  DateDebut := StrToDate(DateDeb.Text) + HeureDeb;
  Datefinal := StrToDate(DateFin.Text) + HeureFin;
  Duree     := StrToFloat(NbHeure.text);
  DureeMin  := Round(Duree * 60);
  Datefinal := CalculDateFinParcMat(DateDebut, DateFinal, NbHJour, DureeMin, TobJFerie);
  //
  TOBeventCha.PutValue('BEC_DATEDEB',       DateDebut);
  TOBeventCha.PutValue('BEC_DATEFIN',       Datefinal);
  //
  DecodeDateTime(DateDebut, An, Mois, Jour, H, M, S, Msec);
  TOBeventCha.PutValue('BEC_HEUREDEB',      TimeToStr(EncodeTime(H, M, 0, 0)));
  DecodeDateTime(Datefinal, An, Mois, Jour, H, M, S, Msec);
  TOBeventCha.PutValue('BEC_HEUREFIN',      TimeToStr(encodeTime(H, M, 0, 0)));
  //
  Duree := StrToFloat(Nbheure.text);
  duree := duree * 60;
  TOBeventCha.PutValue('BEC_DUREE',       duree);
  //
end;

Procedure TOF_BTEVENTCHA.ChargeEcranWithLaTob;
Var Duree : Double;
begin

  //
  Ressource.Text    := LaTob.GetValue('BEC_RESSOURCE');
  ControleRessource;
  //
  TypeAction.Text   := LaTob.GetValue('BEC_BTETAT');
  if TypeAction.text<> '' then ChargeInfoTypeAction;
  //
  Affaire.Text      := LaTob.GetValue('BEC_AFFAIRE');
  ChargeInfoAffaire;
  //
  NumPhase.Text     := LaTob.GetValue('BEC_NUMPHASE');
  if (Numphase.text  <> '') or (Numphase.text  <> '0') then ChargeInfoPhase;
  LibPhase.Caption  := LaTob.GetValue('BEC_LIBPHASE');

  NbHeure.Text      := LaTob.GetValue('BEC_DUREE')/60;
  //
  DateDeb.Text      := LaTob.GetValue('BEC_DATEDEB');
  DateFin.Text      := LaTob.GetValue('BEC_DATEFIN');

  HeureDeb          := LaTob.GetValue('BEC_HEUREDEB');
  HeureFin          := LaTob.GetValue('BEC_HEUREFIN');

  OldDDeb           := LaTob.GetValue('BEC_DATEDEB');
  OldDFin           := LaTob.GetValue('BEC_DATEFIN');

  OldDDeb := Trunc(OldDDeb) + HeureDeb;
  OldDFin := Trunc(OldDFin) + Heurefin;

  NbHeure.Text      := LaTob.GetValue('BEC_DUREE');

  Duree             := CalculDureeEvent(OldDDeb, OldDFin, TobJFerie, TobCalendrier);
  //if Duree <> StrToFloat(NbHeure.Text) then
  //begin
  Duree           := Round(Duree/60);
  NbHeure.Text    := FloatToStr(Duree);
  LaTob.PutValue('BEC_DUREE', NbHeure.text);
  OldDuree        := StrToFloat(NbHeure.text);
  //end;

end;

procedure TOF_BTEVENTCHA.TypeActionOnChange(Sender: TObject);
begin

  if TypeAction.Text <> '' then ChargeInfoTypeAction;

end;

procedure TOF_BTEVENTCHA.NumPhaseOnChange(Sender: TObject);
begin

  if NumPhase.Text <> '' then   ChargeInfoPhase;

end;

//Procedure de recherche de l'équipe associée à une ressource
procedure TOF_BTEVENTCHA.RechEquipeCreation;
var StSql     : string;
    Rep       : string;
    I         : Integer;
    QQ        : TQuery;
    TobEquipe : Tob;
    TobL      : TOB;
    TobEvtE   : TOB;
    TobLEvtE  : TOB;
    Equipe    : string;
    NumEvtcha : Integer;
begin

  if TOBeventCha = nil then Exit;
  
  Equipe := TobEventCha.getvalue('BEC_EQUIPERESS');

  if PGIAsk(TraduireMemoire('La Ressource appartient à l''Equipe ' + Equipe + '.' + Chr(10) + 'Voulez-vous gérer l''equipe ?'), rep) = mrno then exit;

  StSQL := 'SELECT * FROM RESSOURCE ';
  StSQL := StSQL + 'WHERE ARS_EQUIPERESS="' + Equipe + '" ';
  StSQL := StSQL + 'ORDER BY ARS_TYPERESSOURCE';

  QQ := OpenSQL(StSql, True,-1,'',true);
  if QQ.eof then
  begin
    Ferme(QQ);
    Exit;
  end;

  //Création de la Tob equipe uniquement si Equipe existe et
  //si réponse oui à la question
  TobEquipe := Tob.Create('Equipe', nil, -1);
  TobEquipe.LoadDetailDB('LEQUIPE', '', '', QQ, False);

  TobEvtE := TOB.Create('Evt Equipe', nil, -1);

  Ferme(QQ);

  NumEvtCha := StrToInt(NumEventCha.Caption);

  //Vérification si les info de la tob n'existe pas déjà dans le grille
  for i := 0 to TobEquipe.detail.count - 1 do
  begin
    Tobl := TobEquipe.Detail[I];
    if Tobl.GetValue('ARS_RESSOURCE') <>  TobEventCha.getvalue('BEC_RESSOURCE') then
    begin
      If Action = TaCreat then CalculNoCompteur;
      if NumEventCha.Caption = '0' then Continue;
      //mise à jour de la table des Evénements Parc/Matériel
      TobLEvtE := TOB.Create('BTEVENTCHA',TobEvtE,-1);
      TobLEvtE.Dupliquer(TobEventCha, False, True);
      TobLEvtE.PutValue('BEC_IDEVENTCHA',NumEventCha.caption);
      TobLEvtE.PutValue('BEC_RESSOURCE', Tobl.GetValue('ARS_RESSOURCE'));
      TobLEvtE.PutValue('BEC_FONCTION',  Tobl.GetValue('ARS_FONCTION1'));
      TobLEvtE.PutValue('BEC_RESPRINCIPALE', '-');
      TobLEvtE.PutValue('BEC_IDAFFECT',  NumEvtCha);
    end;
  end;

  //On met à jour la table des événement
  TobEvtE.SetAllModifie(true);
  TobEvtE.InsertOrUpdateDB(true);

  FreeAndNil(TobEquipe);
  FreeAndNil(TobEvtE);

end;

//Procedure de recherche de l'équipe associée à une ressource
procedure TOF_BTEVENTCHA.RechEquipeModification;
var StSql     : string;
    Rep       : string;
    I         : Integer;
    QQ        : TQuery;
    TobL      : TOB;
    TobEvtE   : TOB;
    Equipe    : string;
    //
    NbMinute  : Double;
begin

  if TOBeventCha = nil then Exit;

  Equipe := TobEventCha.getvalue('BEC_EQUIPERESS');

  if PGIAsk(TraduireMemoire('La Ressource appartient à l''Equipe ' + Equipe + '.' + Chr(10) + 'Voulez-vous gérer l''equipe ?'), rep) = mrno then exit;

  StSQL := 'SELECT * FROM BTEVENTCHA ';
  StSQL := StSQL + 'WHERE BEC_EQUIPERESS="' + Equipe + '" ';
  StSQL := StSQL + ' AND  BEC_IDAFFECT="' + NumEventCha.Caption  + '" ';
  StSQL := StSQL + 'ORDER BY BEC_RESSOURCE';

  QQ := OpenSQL(StSql, True,-1,'',true);
  if QQ.eof then
  begin
    Ferme(QQ);
    Exit;
  end;

  //Création de la Tob equipe uniquement si Equipe existe et
  //si réponse oui à la question
  TobEvtE := Tob.Create('BTEVENTCHA', nil, -1);
  TobEvtE.LoadDetailDB('BTEVENTCHA', '', '', QQ, False);

  Ferme(QQ);

  //Vérification si les info de la tob n'existe pas déjà dans le grille
  for i := 0 to TobEvtE.detail.count - 1 do
  begin
    Tobl := TobEvtE.Detail[I];
    if Tobl.GetValue('ARS_RESSOURCE') <>  TobEventCha.GetValue('BEC_RESSOURCE') then
    begin
      //mise à jour de la table des Evénements MO
      TobL.PutValue('BEC_BTETAT',    TypeAction.Text);
      Tobl.PutValue('BEC_AFFAIRE',   Affaire.Text);
      Tobl.PutValue('BEC_TIERS',     CodeClient);
      Tobl.PutValue('BEC_LIBACTION', LibelleAction.Text);
      Tobl.PutValue('BEC_NUMPHASE',  NumPhase.Text);
      Tobl.PutValue('BEC_LIBPHASE',  LibPhase.caption);
      Tobl.PutValue('BEC_DATEDEB',   DateDeb.Text);
      Tobl.PutValue('BEC_DATEFIN',   DateFin.Text);
      //
      NbMinute := (StrtoFloat(NbHeure.text) * 60);
      //
      Tobl.PutValue('BEC_DUREE',     NbMinute);
        //On met à jour la table des événement
      TobL.SetAllModifie(true);
      TobL.UpdateDB(true);
    end;
  end;


  FreeAndNil(TobEvtE);

end;


Initialization
  registerclasses ( [ TOF_BTEVENTCHA ] ) ;
end.

