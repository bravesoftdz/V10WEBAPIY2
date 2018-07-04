{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTEVENPLAN_TOF ()
Mots clefs ... : TOF;BTEVENPLAN_TOF
*****************************************************************}
Unit BTEvenPlan_TOF ;

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
     HRichOle,
     UTOF ;

Type
  TOF_BTEVENPLAN = Class (TOF)
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
    CtrlCal       : Boolean;
    //
    HDeb          : TDateTime;
    HFin          : TDateTime;
    NbHJour       : TDateTime;
    //
    OldDDeb       : TDateTime;
    OldDFin       : TDateTime;
    //
    NumEvenPlan   : THEdit;
    //
    Ressource     : THEdit;
    NomRessource  : THLabel;
    TypeAction    : THEdit;
    TNbHeure      : THLabel;
    LibTypeAction : THLabel;
    //
    DateDeb       : THEdit;
    DateFin       : THEdit;
    HeureDeb      : THEdit;
    HeureFin      : THEdit;
    //
    FF            : String;
    //
    NbHeure       : THEdit;
    //
    Modifiable    : TCheckBox;
    TravailFerie  : TCheckBox;
    Surbook       : TCheckBox;
    //
    BValider      : TToolBarButton97;
    BTSelect      : TToolbarButton97;
    BTEfface      : TToolbarButton97;
    //
    Descriptif    : THRichEditOLE;
    //
    TobEvenPlan   : Tob;
    TobCalendrier : Tob;
    TobJferie     : Tob;
    //
    StSQL         : String;
    //
    QQ            : TQuery;
    //
    NbMinute      : Double;
    //
    Ok_Obligatoire  : Boolean;
    OK_GerePlanning : Boolean;
    OK_Modifiable   : Boolean;
    //
    procedure CalculNoCompteur;
    procedure ChargeInfoTypeAction;
    procedure ChargeZoneEcranWithLaTob;
    procedure ChargeZoneTOB;
    procedure Controlechamp(Champ, Valeur: String);
    procedure ControleRessource;
    function  ControleSaisie: Boolean;
    procedure CreateTOB;
    procedure DateDebOnExit(Sender: TObject);
    procedure DateFinOnExit(Sender: TObject);
    procedure HeureDebOnExit(Sender: TObject);
    procedure HeureFinOnExit(Sender: TObject);
    procedure DestroyTOB;
    procedure GetObjects;
    procedure NBHeureOnExit(Sender: TObject);
    procedure RazZoneEcran;
    procedure RechEquipeCreation;
    procedure RechEquipeModification;
    procedure RessourceEffaceOnClick(Sender: TObject);
    procedure RessourceOnExit(Sender: TObject);
    procedure RessourceSelectOnClick(Sender: TObject);
    procedure SetScreenEvents;
    procedure TypeActionOnExit(Sender: TObject);
    procedure TypeActionOnChange(Sender: TObject);
    function ConstitueRequeteEvenementPlanning(NumEvent: String): string;

  end ;

const
  TexteMessage: array[1..9] of string  = (
          {1}   'Le Type d''action est Obligatoire'
          {2},  'Le code Ressource est Obligatoire'
          {3},  'La suppression a échoué'
          {4},  'Création impossible, numéro non déterminé !'
          {5},  'Le Type d''action est inexistant'
          {6},  'Le code Ressource est inexistant'
          {7},  'La date de début est supérieure à la date de fin'
          {8},  'la date de fin est inférieure à la date de début'
          {9},  'les dates saisies ne correspondent à aucune plage horaire'
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
     BTPUtil;

procedure TOF_BTEVENPLAN.OnArgument (S : String ) ;
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
  CtrlCal   := True;
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
  AppliqueFontDefaut (Descriptif);
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
  TypeAction.plus := 'BTA_TYPEACTION="INT"'; //==> voir pour enlever ABSPAIE ET ACT-GRC (???)
  //
  ressource.Enabled      := True;
  //
  BtSelect.Visible       := True;
  BtEfface.Visible       := True;
  //
  //if Action = Tacreat then OnNew;
  RazZoneEcran;

end ;

Procedure TOF_BTEVENPLAN.OnLoad ;
begin
  Inherited ;

  if LaTOB = nil then
  begin
    if NumEvenPlan.Text = '' then exit;
    //StSQL := 'SELECT * FROM BTEVENPLAN WHERE BEP_CODEEVENT=' + NumEvenPlan.Text;
    StSQL := ConstitueRequeteEvenementPlanning(NumEvenPlan.Text);
    QQ := OpenSQL(StSQL, True);

    If not QQ.Eof then
    begin
      LaTob.SelectDB('BTEVENEMENTPLA', QQ);
    end;

    ferme(QQ);
  end;

  ChargeZoneEcranWithLaTob;

  if (AppelPlan = False) Or (LaTOB = nil) then
  begin
    DateDeb.Text := DateToStr(Now);
    Datefin.Text := DateToStr(Now+1);
  end;
  //
  if ressource.Text <> '' then
  begin
    BtSelect.visible     := False;
    BtEfface.visible     := False;
    ressource.Enabled    := False;
  end;

  //Recherche du type d'action par défaut
  if TypeAction.text = '' then
  begin
    StSQL := 'SELECT BTA_BTETAT, BTA_LIBELLE FROM BTETAT WHERE BTA_TYPEACTION="INT" AND BTA_DEFAUT="X"';
    QQ    := OpenSQL(StSQL, False,-1,'',True);
    if not QQ.eof then
    begin
      TypeAction.text       := QQ.Findfield('BTA_BTETAT').AsString;
      LibTypeAction.Caption := QQ.Findfield('BTA_LIBELLE').AsString;
    end;
    Ferme(QQ);
  end;

end ;

procedure TOF_BTEVENPLAN.OnNew ;
begin
  Inherited ;
end;

procedure TOF_BTEVENPLAN.OnDelete ;
begin
  Inherited ;

  if PGIAsk('Confirmez-vous la suppression de cet évènement ?', 'Saisie Evénement') = Mryes then
  begin
    StSQL := 'DELETE BTEVENPLAN WHERE BEP_CodeEvent=' + NumEvenPlan.Text;
    if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[3], 'Evénement Matériel');
    OnClose;
  end;

end ;

procedure TOF_BTEVENPLAN.OnUpdate ;
begin
  Inherited ;

  if ControleSaisie then
  begin
    Ecran.ModalResult := 1;

    //si la TOB est vide on fait rien et surtout cela signifie qu'on a un gros problème
    if LaTOB = nil then Exit;

    If Action = TaCreat then CalculNoCompteur;

    if  NumEvenPlan.Text = '0' then exit;

    //mise à jour de la table des Evénements Parc/Matériel
    ChargeZoneTOB;

    TobEvenPlan.PutValue('BEP_EQUIPERESS', LaTob.GetValue('BPL_EQUIPERESS'));

    //On met à jour la table des événements
    TOBEvenPlan.SetAllModifie(true);
    TOBEvenPlan.InsertOrUpdateDB(true);

    //Gestion de l'équipe
    if (TobEvenPlan.getvalue('BEP_EQUIPERESS') <> '') then
    begin
      If Action = TaCreat       then RechEquipeCreation
      else if Action = TaModif  then RechEquipeModification
    end;

    If (AppelPlan) and (LaTob <> nil) then LaTob.PutValue('RETOUR',          1);
  end
  else
  begin
    If (AppelPlan) and (LaTob <> nil) then LaTob.PutValue('RETOUR',          0);
    Ecran.ModalResult := 0;
  end;

end;


Function TOF_BTEVENPLAN.ControleSaisie : Boolean;
begin

  Result := False;

  if TypeAction.text = '' then
  begin
    PGIError(TexteMessage[1], 'Erreur Saisie');
    TypeAction.SetFocus;
    exit;

  end;
  if not TypeActionExist(TypeAction.text, ' AND BTA_TYPEACTION="INT"') then
  begin
    PGIError(TexteMessage[5], 'Erreur Saisie');
    TypeAction.SetFocus;
    exit;
  end;

  if Ressource.Text = '' then
  begin
    PGIError(TexteMessage[2], 'Erreur Saisie');
    TypeAction.SetFocus;
    exit;
  end;

  if not RessourceExist(Ressource.text, ' AND ARS_TYPERESSOURCE in ("SAL", "INT")') then
  begin
    PGIError(TexteMessage[6], 'Erreur Saisie');
    Ressource.SetFocus;
    exit;
  end;

  if StrToDate(DateDeb.text) > StrToDate(Datefin.text) then
  begin
    PGIError(TexteMessage[7], 'Erreur Saisie');
    DateDeb.SetFocus;
    exit;
  end;

  if StrToDate(DateFin.text) < StrToDate(DateDeb.text) then
  begin
    PGIError(TexteMessage[8], 'Erreur Saisie');
    DateFin.SetFocus;
    exit;
  end;

  //controle de l'heure de début et de l'heure de fin...
  //Elles doivent être comprise entre les heures de début et de fin du paramètre sociéte
  if not ControleRessourceDispo(LaTob, CtrlCal) then
  Begin
    PGIError(TexteMessage[9], 'Erreur Saisie');
    DateDeb.SetFocus;
    exit;
  end;

  Result := true;

end ;

procedure TOF_BTEVENPLAN.OnClose ;
begin
  Inherited ;

  DestroyTOB;

end ;

procedure TOF_BTEVENPLAN.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTEVENPLAN.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTEVENPLAN.GetObjects;
begin

  NumEvenPlan   := THEdit(Getcontrol('CODEEVENT'));
  //
  Ressource     := THEdit(Getcontrol('RESSOURCE'));
  NomRessource  := THLabel(Getcontrol('NOMRESSOURCE'));
  //
  TypeAction    := THEdit(Getcontrol('BTETAT'));
  LibTypeAction := THLabel(Getcontrol('LIBTYPEACTION'));
  DateDeb       := THEdit(Getcontrol('DATEEVT'));
  DateFin       := THEdit(Getcontrol('DATEEVT_'));
  HeureDeb      := THEdit(Getcontrol('HEUREDEB'));
  HeureFin      := THEdit(Getcontrol('HEUREFIN'));

  TNbHeure      := THLabel(Getcontrol('TNbHeure'));

  NbHeure       := THEdit(Getcontrol('DUREE'));
  //
  Descriptif    := THRichEditOLE(GetControl('DESCRIPTIF'));
  //
  Modifiable    := TCheckBox(Getcontrol('MODIFIABLE'));
  TravailFerie  := TCheckBox(Getcontrol('CHECKTRAVAILFERIE'));
  Surbook       := TCheckBox(Getcontrol('SURBOOKING'));
  //
  BTSelect      := TToolBarButton97(Getcontrol('BSELECT'));
  BTEfface      := TToolBarButton97(Getcontrol('BEFFACE'));
  BValider      := TToolBarButton97(GetCONTROL('BVALIDER'));

end;

procedure TOF_BTEVENPLAN.SetScreenEvents;
begin
  //
  TypeAction.OnExit       := TypeActionOnExit;
  TypeAction.OnChange     := TypeActionOnChange;
  //
  Ressource.OnExit        := RessourceOnExit;
  //
  BTSelect.OnClick        := RessourceSelectOnClick;
  BTEfface.OnClick        := RessourceEffaceOnClick;
  //
  DateDeb.OnExit          := DateDebOnExit;
  DateFin.OnExit          := DateFinOnExit;
  HeureDeb.OnExit         := HeureDebOnExit;
  HeureFin.OnExit         := HeureFinOnExit;

  NBHeure.OnExit          := NBHeureOnExit;

end;

Procedure TOF_BTEVENPLAN.RessourceOnExit(Sender : TObject);
begin

  if Ressource.text <> '' then ControleRessource;

end;

Procedure TOF_BTEVENPLAN.RessourceSelectOnClick(Sender : TObject);
Var StChamps  : string;
    StWhere   : string;
begin

  StChamps  := ressource.Text;

  StWhere := 'ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")';

  DispatchRecherche(Ressource, 3, stwhere,'ARS_RESSOURCE=' + Trim(Ressource.Text), '');

  ControleRessource;

end;

Procedure TOF_BTEVENPLAN.RessourceEffaceOnClick(Sender : TObject);
begin

  Ressource.Text  := '';
  NomRessource.Caption := '...';

end;

Procedure TOF_BTEVENPLAN.ControleRessource;
Var StWhere       : string;
    TobRessource  : TOB;
    Calendrier    : string;
begin

  if ressource.text = '' then ressource.SetFocus;

  TobRessource := Tob.create('RESSOURCE', nil, -1);

  //StWhere := ' AND ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")';

  ChargeInfoRessource(Ressource.text, StWhere, TobRessource);

  NomRessource.caption := TobRessource.GetValue('ARS_LIBELLE2') + ' ' + TobRessource.GetValue('ARS_LIBELLE');

  RechercheCalendrier(Ressource.text, Tobcalendrier);

  if NomRessource.Caption = '' then Ressource.setfocus;

  LaTob.PutValue('BPL_RESSOURCE',   TobRessource.GetValue('ARS_RESSOURCE'));
  LaTob.PutValue('BPL_EQUIPERESS',  TobRessource.GetValue('ARS_EQUIPERESS'));

  Calendrier := TobRessource.GetString('ARS_STANDCALEN');

  if (NumEvenPlan.Text <> '') Or (NumEvenPlan.Text <> '0') then
  begin
   if not ControleRessourceDispo(LaTob, CtrlCal) then
    begin
      Ressource.text := '';
      NomRessource.caption := '...';
      Ressource.SetFocus;
    end;
  end;

  FreeAndNil(TobRessource);

end;

Procedure TOF_BTEVENPLAN.TypeActionOnExit(Sender : TObject);
begin

  ChargeInfoTypeAction;

  if TypeAction.Text = '' then TypeAction.SetFocus;

end;

procedure TOF_BTEVENPLAN.DateDebOnExit(Sender: TObject);
Var TDateDeb  : TDateTime;
    TDateFin  : TDateTime;
begin

  TDateDeb    := StrToDate(DateDeb.Text);
  TDateFin    := StrToDate(DateFin.Text);

  if TDateDeb > TDatefin then
  begin
    PGIError(TexteMessage[7], 'Erreur Saisie');
    DateDeb.Text := DateToStr(OldDDeb);
    Datedeb.SetFocus;
    Exit;
  end;

  if TDatedeb = OldDdeb then exit;

  NbMinute := CalculDureeEvenement(TDateDeb, TDatefin, CtrlCal);
  if NbMinute > 59 then
  begin
    TNbHeure.Caption := 'Nb Heure(s):';
    //Temp := Arrondi(Temp/60,2);
  end
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

  OldDDeb := StrToDate(DateDeb.text);

end;

procedure TOF_BTEVENPLAN.DateFinOnExit(Sender: TObject);
Var TDateDeb  : TDateTime;
    TDateFin  : TDateTime;
begin

  TDateDeb    := DateToFloat (StrToDate(DateDeb.Text));
  TDateFin    := DateToFloat (StrToDate(DateFin.Text));

  if TDateFin < TDateDeb then
  begin
    PGIError(TexteMessage[8], 'Erreur Saisie');
    DateFin.Text := DateToStr(OldDFin);
    DateFin.SetFocus;
    exit;
  end;

  if TDateFin = idate2099 then exit;

  if TDateFin = OldDFin   then exit;

  NbMinute := CalculDureeEvenement(StrToDate(DateDeb.text)+StrToTime(HeureDeb.text),StrToDate(DateFin.text)+StrToTime(HeureFin.text), CtrlCal);
  if NbMinute > 59 then
  begin
    TNbHeure.Caption := 'Nb Heure(s):';
    //Temp := Arrondi(Temp/60,2);
  end
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

  OldDFin := StrToDate(DateFin.text);

end;

procedure TOF_BTEVENPLAN.HeureDebOnExit(Sender: TObject);
Var	Heure       : TdateTime;
    HeureDebAM  : TdateTime;
    HeureFinAM  : TDateTime;
    HeureDebPM  : Tdatetime;
    HeureFinPM  : TDateTime;
Begin

  if copy(HeureDeb.text, 0, 2) > '24' then
  Begin
    Heure         := StrtoDateTime(LaTob.GetValue('BPL_HEUREDEB'));
    HeureDeb.Text := TimetoStr(Heure);
    Exit;
  end;

  if copy(HeureDeb.text, 4, 2) > '59' then
  Begin
    Heure         := StrtoDateTime(LaTob.GetValue('BPL_HEUREDEB'));
    HeureDeb.Text := TimetoStr(Heure);
    Exit;
  end;

  //Il faut contrôler que la date saisie soit cohérente avec les heure de début et de fin....
  //soit de la ressource soit du paramétrage...
  Heure           := StrTotime(HeureDeb.text);

  //Si jamais il y a un calendrier sur la ressource celui-ci ne sera pas pris en compte ici...
  if CtrlCal then
  begin
    HeureDebAm      := GetDebutMatinee;
    HeureFinAm      := GetFinMatinee;
    HeureDebPm      := GetDebutApresMidi;
    HeurefinPm      := GetfinApresMidi;

    if (Heure < heureDebAM) then Heure := heureDebAM;

    if (Heure > heureFinPM) then Heure := heureFinPM;

    if (Heure > heureFinAM) And (Heure < heureDebPM) then
      Heure         := StrtoDateTime(LaTob.GetValue('BPL_HEUREDEB'));
  end;

  HeureDeb.Text := TimetoStr(Heure);

  NbMinute := CalculDureeEvenement(StrToDate(DateDeb.text)+StrToTime(HeureDeb.text),StrToDate(DateFin.text)+StrToTime(HeureFin.text),CtrlCal);
  if NbMinute > 59 then
  begin
    TNbHeure.Caption := 'Nb Heure(s):';
    //Temp := Arrondi(Temp/60,2);
  end
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

end;

procedure TOF_BTEVENPLAN.HeureFinOnExit(Sender: TObject);
Var	Heure       : TdateTime;
    HeureDebAM  : TdateTime;
    HeureFinAM  : TDateTime;
    HeureDebPM  : Tdatetime;
    HeureFinPM  : TDateTime;
begin


  if copy(HeureFin.text, 0, 2) > '24' then
  Begin
    Heure         := StrtoDateTime(LaTob.GetValue('BPL_HEUREFIN'));
    Heurefin.Text := TimetoStr(Heure);
    Exit;
  end;

  if copy(Heurefin.text, 4, 2) > '59' then
  Begin
    Heure         := StrtoDateTime(LaTob.GetValue('BPL_HEUREFIN'));
    Heurefin.Text := TimetoStr(Heure);
    Exit;
  end;

  if (StrToTime(HeureFin.Text) < StrToTime(HeureDeb.text)) then
  begin
    Heure         := StrtoDateTime(LaTob.GetValue('BPL_HEUREFIN'));
    Heurefin.Text := TimetoStr(Heure);
    Exit;
  end;

  //Il faut contrôler que la date saisie soit cohérente avec les heure de début et de fin....
  //soit de la ressource soit du paramétrage...
  Heure           := StrTotime(HeureFin.text);

  //Si jamais il y a un calendrier sur la ressource celui-ci ne sera pas pris en compte ici...
  if CtrlCal then
  begin
    HeureDebAm      := GetDebutMatinee;
    HeureFinAm      := GetFinMatinee;
    HeureDebPm      := GetDebutApresMidi;
    HeurefinPm      := GetfinApresMidi;

    if (Heure < HeureDebAM) then Heure := heureDebAM;

    if (Heure > HeureFinPM) then Heure := heureFinPM;

    if (Heure > HeureFinAM) And (Heure < heureDebPM) then
      Heure         := StrtoDateTime(LaTob.GetValue('BPL_HEUREFIN'));
  end;

  HeureFin.Text := TimetoStr(Heure);

  NbMinute := CalculDureeEvenement(StrToDate(DateDeb.text)+StrToTime(HeureDeb.text),StrToDate(DateFin.text)+StrToTime(HeureFin.text),CtrlCal);
  if NbMinute > 59 then
  begin
    TNbHeure.Caption := 'Nb Heure(s):';
    //Temp := Arrondi(Temp/60,2);
  end
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

end;

procedure TOF_BTEVENPLAN.NBHeureOnExit(Sender: TObject);
begin
end;

Procedure TOF_BTEVENPLAN.Controlechamp(Champ, Valeur : String);
Var Date1 : TdateTime;
    Date2 : TDateTime;
begin

  if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if Champ='IDEVENEMENT' then NumEvenPlan.Text := Valeur;

  if Champ='RESSOURCE'   then Ressource.Text := Valeur;

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

Procedure TOF_BTEVENPLAN.CreateTOB;
begin

  TOBEvenPlan   := Tob.Create('BTEVENPLAN', nil, -1);

  TobJferie     := TOB.Create('JOURFERIE', nil, -1);
  Tobcalendrier := TOB.Create('CALENDRIER', nil, -1);

end;

procedure TOF_BTEVENPLAN.DestroyTOB;
begin

  FreeAndNil(TobEvenPlan);
  FreeAndNil(TobJFerie);
  FreeAndNil(TobCalendrier);

end;

Procedure TOF_BTEVENPLAN.ChargeInfoTypeAction;
var TOBTypeAction : TOB;
    StWhere       : String;
begin

  if TypeAction.text = '' then Exit;;

  TOBTypeAction := TOB.Create('BTETAT',nil, -1);

  Stwhere := ' AND BTA_TYPEACTION="INT"';

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
  Ok_Obligatoire        := TOBTypeAction.GetBoolean('BTA_OBLIGATOIRE');
  OK_GerePlanning       := TOBTypeAction.GetBoolean('BTA_ASSOSRES');
  OK_Modifiable         := TOBTypeAction.GetBoolean('BTA_MODIFIABLE');
  //
  FreeAndNil(TOBTypeAction);

  ferme(QQ);

end;

procedure TOF_BTEVENPLAN.RazZoneEcran;
begin

  Ressource.text        := '';
  NomRessource.caption  := '...';
  //
  TypeAction.text       := '';
  LibTypeAction.caption := '...';
  //
  DateDeb.Text          := DateToStr(now);
  DateFin.Text          := DateToStr(idate2099);
  //
  HeureDeb.text         := TimeToStr(GetDebutMatinee);
  HeureFin.Text         := TimetoStr(GetFinApresMidi);
  //
  OldDDeb               := StrToDate(DateDeb.Text);
  OldDFin               := StrtoDate(Datefin.Text);
  //
  Modifiable.Checked    := False;
  TravailFerie.Checked  := False;
  Surbook.Checked       := False;
  //

  NbMinute     := 0;
  NbHeure.Text := LibelleDuree(NbMinute, False);

  NumEvenPlan.Text      := '0';
  //

end ;

Procedure TOF_BTEVENPLAN.ChargeZoneEcranWithLaTob;
Var Heure     : TTime;
begin

  CtrlCal             := Latob.GetBoolean('BPL_CONTROLCAL');

  NumEvenPlan.Text    := LaTob.GetString('BPL_IDEVENT');
  //
  DateDeb.Text        := DateToStr(LaTob.GetValue('BPL_DATEDEB'));
  DateFin.Text        := DateToStr(LaTob.GetValue('BPL_DATEFIN'));
  //
  Heure               := StrtoDateTime(LaTob.GetValue('BPL_HEUREDEB'));
  HeureDeb.Text       := TimetoStr(Heure);
  Heure               := StrtoTime(LaTob.GetValue('BPL_HEUREFIN'));
  HeureFin.Text       := TimeToStr(Heure);
  //
  Ressource.Text      := LaTob.GetString('BPL_RESSOURCE');
  ControleRessource;
  //
  TypeAction.Text       := LaTob.GetString('BPL_BTETAT');
  LibTypeAction.caption := Latob.GetString('BPL_LIBACTION');   
  //
  OldDDeb             := StrToDate(DateDeb.Text);
  OldDFin             := StrtoDate(DateFin.Text);

  Descriptif.Text     := LaTOB.GetValue('BLOCNOTE');

  //Chargement de la durée mini
  //Pourquoi ne pas lancer le calcul ici ==> certitude que cela corresponde aux dates affichées...
  NbMinute := CalculDureeEvenement(StrToDate(DateDeb.text)+StrToTime(HeureDeb.text),StrToDate(DateFin.text)+StrToTime(HeureFin.text), CtrlCal);
  if NbMinute > 59 then
  begin
    TNbHeure.Caption := 'Nb Heure(s):';
    //Temp := Arrondi(Temp/60,2);
  end
  else
    TNbHeure.Caption := 'Nb Minutes :';

  NbHeure.Text := LibelleDuree(NbMinute, False);

end;

procedure TOF_BTEVENPLAN.CalculNoCompteur;
Var NumAff : Integer;
begin

  if not GetNumCompteur('BEP',iDate1900, NumAff) then
  begin
    PGIError(TexteMessage[4], 'Evènement Intervention');
    OnClose;
  end
  else NumEvenPlan.text := IntToStr(NumAff);

end;

Procedure TOF_BTEVENPLAN.ChargeZoneTOB;
begin

  TOBEvenPlan.PutValue('BEP_CODEEVENT',     NumEvenPlan.Text);
  TOBEvenPlan.PutValue('BEP_RESSOURCE',     Ressource.Text);
  TOBEvenPlan.PutValue('BEP_BTETAT',        TypeAction.Text);
  TOBEvenPlan.PutValue('BEP_AFFAIRE',       '');
  TOBEvenPlan.PutValue('BEP_TIERS',         '');
  TOBEvenPlan.PutValue('BEP_NUMEROADRESSE', 0);

  //calcul de l'heure de fin en fonction de la durée
  //NbMinute := StrToInt(FloatToStr((StrToFloat(NbHeure.text) * 60)));

  TOBEvenPlan.PutValue('BEP_HEUREDEB',      DateTimeToStr(StrToDateTime(HeureDeb.Text)));
  TOBEvenPlan.PutValue('BEP_HEUREFIN',      DateTimeToStr(StrToDateTime(HeureFin.Text)));

  TOBEvenPlan.PutValue('BEP_RESPRINCIPALE', 'X');
  TOBEvenPlan.PutValue('BEP_EQUIPESEP',     '-');

  TOBEvenPlan.PutValue('BEP_DATEDEB',       DateTimeToStr(StrToDate(DateDeb.Text) + StrToTime(HeureDeb.Text)));
  TOBEvenPlan.PutValue('BEP_DATEFIN',       DateTimeToStr(StrToDate(DateFin.Text) + StrToTime(HeureFin.Text)));
  //
  TOBEvenPlan.PutValue('BEP_DUREE',         NbMinute);
  //
  TOBEvenPlan.PutValue('BEP_BLOCNOTE',      Descriptif.text);
  //
  TobEvenPlan.PutValue('BEP_OBLIGATOIRE',   BoolToStr(Ok_Obligatoire));
  TobEvenPlan.PutValue('BEP_GEREPLAN',      BoolToStr(OK_GerePlanning));
  TobEvenPlan.PutValue('BEP_MODIFIABLE',    BoolToStr(OK_Modifiable));

end;

procedure TOF_BTEVENPLAN.TypeActionOnChange(Sender: TObject);
begin

  if TypeAction.Text <> '' then ChargeInfoTypeAction;

end;

//Procedure de recherche de l'équipe associée à une ressource
procedure TOF_BTEVENPLAN.RechEquipeCreation;
var StSql     : string;
    Rep       : string;
    I         : Integer;
    QQ        : TQuery;
    TobEquipe : Tob;
    TobL      : TOB;
    TobEvtE   : TOB;
    TobLEvtE  : TOB;
    Equipe    : string;
    EvenPlan  : Integer;
begin

  if TOBEvenPlan = nil then Exit;

  Equipe := TobEvenPlan.getvalue('BEP_EQUIPERESS');

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

  EvenPlan := StrToInt(NumEvenPlan.Text);

  //Vérification si les info de la tob n'existe pas déjà dans le grille
  for i := 0 to TobEquipe.detail.count - 1 do
  begin
    Tobl := TobEquipe.Detail[I];
    if Tobl.GetValue('ARS_RESSOURCE') <>  TobEvenPlan.getvalue('BEP_RESSOURCE') then
    begin
      If Action = TaCreat then CalculNoCompteur;
      if NumEvenPlan.text = '0' then Continue;
      //mise à jour de la table des Evénements Parc/Matériel
      TobLEvtE := TOB.Create('BTEVENPLAN',TobEvtE,-1);
      TobLEvtE.Dupliquer(TobEvenPlan, False, True);
      TobLEvtE.PutValue('BEP_CodeEvent',NumEvenPlan.Text);
      TobLEvtE.PutValue('BEP_RESSOURCE', Tobl.GetValue('ARS_RESSOURCE'));
      TobLEvtE.PutValue('BEP_FONCTION',  Tobl.GetValue('ARS_FONCTION1'));
      TobLEvtE.PutValue('BEP_RESPRINCIPALE', '-');
      TobLEvtE.PutValue('BEP_IDAFFECT',  EvenPlan);
    end;
  end;

  //On met à jour la table des événement
  TobEvtE.SetAllModifie(true);
  TobEvtE.InsertOrUpdateDB(true);

  FreeAndNil(TobEquipe);
  FreeAndNil(TobEvtE);

end;

//Procedure de recherche de l'équipe associée à une ressource
procedure TOF_BTEVENPLAN.RechEquipeModification;
var StSql     : string;
    Rep       : string;
    I         : Integer;
    QQ        : TQuery;
    TobL      : TOB;
    TobEvtE   : TOB;
    Equipe    : string;
begin

  if TOBEvenPlan = nil then Exit;

  Equipe := TobEvenPlan.getvalue('BEP_EQUIPERESS');

  if PGIAsk(TraduireMemoire('La Ressource appartient à l''Equipe ' + Equipe + '.' + Chr(10) + 'Voulez-vous gérer l''equipe ?'), rep) = mrno then exit;

  StSQL := 'SELECT * FROM BTEVENPLAN ';
  StSQL := StSQL + 'WHERE BEP_EQUIPERESS="' + Equipe + '" ';
  StSQL := StSQL + ' AND  BEP_IDAFFECT="' + NumEvenPlan.Text  + '" ';
  StSQL := StSQL + 'ORDER BY BEP_RESSOURCE';

  QQ := OpenSQL(StSql, True,-1,'',true);
  if QQ.eof then
  begin
    Ferme(QQ);
    Exit;
  end;

  //Création de la Tob equipe uniquement si Equipe existe et
  //si réponse oui à la question
  TobEvtE := Tob.Create('BTEVENPLAN', nil, -1);
  TobEvtE.LoadDetailDB('BTEVENPLAN', '', '', QQ, False);

  Ferme(QQ);

  //Vérification si les info de la tob n'existe pas déjà dans le grille
  for i := 0 to TobEvtE.detail.count - 1 do
  begin
    Tobl := TobEvtE.Detail[I];
    if Tobl.GetValue('ARS_RESSOURCE') <>  TobEvenPlan.GetValue('BEP_RESSOURCE') then
    begin
      //mise à jour de la table des Evénements MO
      TobL.PutValue('BEP_BTETAT',    TypeAction.Text);
      Tobl.PutValue('BEP_AFFAIRE',   '');
      Tobl.PutValue('BEP_TIERS',     '');
      Tobl.PutValue('BEP_DATEDEB',   DateDeb.Text);
      Tobl.PutValue('BEP_DATEFIN',   DateFin.Text);
      Tobl.PutValue('BEP_HEUREDEB',  HeureDeb.Text);
      Tobl.PutValue('BEP_HEUREFIN',  HeureFin.Text);
      Tobl.PutValue('BEP_DUREE',     NbMinute);
      Tobl.PutValue('BEP_BLOCNOTE',  Descriptif.Text);

      //On met à jour la table des événement
      TobL.SetAllModifie(true);
      TobL.UpdateDB(true);
    end;
  end;


  FreeAndNil(TobEvtE);

end;

function TOF_BTEVENPLAN.ConstitueRequeteEvenementPlanning(NumEvent : String) : string;
begin

  Result := '';

  Result := 'SELECT "INTERV"            AS BPL_ORIGINEITEM, ';
  Result := Result + 'BEP_CODEEVENT     AS BPL_IDEVENT, ';
  Result := Result + 'BEP_AFFAIRE       AS BPL_CODEAFFAIRE, ';
  Result := Result + 'AFF_LIBELLE       AS BPL_LIBAFFAIRE, ';
  Result := Result + 'BEP_TIERS         AS BPL_CODETIERS, ';
  Result := Result + 'T_AUXILIAIRE      AS BPL_AUXILIAIRE, ';
  Result := Result + 'T_LIBELLE         AS BPL_LIBTIERS, ';
  Result := Result + 'ADR_JURIDIQUE     AS BPL_NOMJURIDIQUE, ';
  Result := Result + 'ADR_LIBELLE       AS BPL_NOM1, ';
  Result := Result + 'ADR_LIBELLE2      AS BPL_NOM2, ';
  Result := Result + 'ADR_ADRESSE1      AS BPL_ADR1, ';
  Result := Result + 'ADR_ADRESSE2      AS BPL_ADR2, ';
  Result := Result + 'ADR_ADRESSE3      AS BPL_ADR3, ';
  Result := Result + 'ADR_CODEPOSTAL    AS BPL_CP,  ';
  Result := Result + 'ADR_VILLE         AS BPL_VILLE, ';
  Result := Result + 'BEP_DATEDEB       AS BPL_DATEDEB, ';
  Result := Result + 'BEP_DATEFIN       AS BPL_DATEFIN, ';
  Result := Result + 'BEP_HEUREDEB      AS BPL_HEUREDEB, ';
  Result := Result + 'BEP_HEUREFIN      AS BPL_HEUREFIN, ';
  Result := Result + 'BEP_DUREE         AS BPL_DUREE, ';
  Result := Result + 'BEP_BTETAT        AS BPL_ACTION, ';
  Result := Result + 'BTA_LIBELLE       AS BPL_LIBACTION, ';
  Result := Result + 'BTA_ASSOSRES      AS BPL_ASSOSRES, ';
  Result := Result + 'BEP_RESSOURCE     AS BPL_CODERESSOURCE, ';
  Result := Result + 'ARS_LIBELLE       AS BPL_LIBRESSOURCE, ';
  Result := Result + 'ARS_EQUIPERESS    AS BPL_EQUIPE, ';
  Result := Result + 'BEP_RESSOURCE     AS BPL_MATERIEL, ';
  Result := Result + 'ARS_LIBELLE       AS BPL_LIBMATERIEL, ';
  Result := Result + 'ARS_TYPERESSOURCE AS BPL_FAMILLEMAT, ';
  Result := Result + '""                AS BPL_LIBFAMILLEMAT, ';
  Result := Result + 'AFF_ETATAFFAIRE   AS BPL_CODEETAT, ';
  Result := Result + '""                AS BPL_LIBEVENEMENT, ';
  Result := Result + '""                AS BPL_LIBREALISE, ';
  Result := Result + 'BEP_BLOCNOTE      AS BPL_BLOCNOTE, ';
  Result := Result + 'ARS_FONCTION1     AS BPL_CODEFONCTION, ';
  Result := Result + 'AFO_LIBELLE       AS BPL_LIBFONCTION, ';
  Result := Result + '"0"               AS BPL_NUMPHASE, ';
  Result := Result + '""                AS BPL_LIBPHASE, ';
  Result := Result + 'ARS_EMAIL         AS BPL_EMAIL, ';
  Result := Result + 'ARS_STANDCALEN    AS BPL_STANDCALEN, ';
  Result := Result + 'ARS_CALENSPECIF   AS BPL_CALENSPECIF, ';
  Result := Result + 'ARS_TYPERESSOURCE AS BPL_TYPERESSOURCE, ';
  Result := Result + 'ARS_CHAINEORDO    AS BPL_STYPERESSOURCE, ';
  Result := Result + 'BTR_LIBELLE       AS BPL_LIBSTYPERESSOURCE, ';
  Result := Result + '""                AS BPL_TYPEACTION, ';
 	Result := Result + '""                AS BPL_TYPEMVT, ';
  Result := Result + '""                AS BPL_TYPECONGE, ';
  Result := Result + 'ARS_SALARIE       AS BPL_SALARIE ';
  Result := Result + 'FROM BTEVENPLAN ';
  Result := Result + 'LEFT JOIN BTETAT    ON (BTA_BTETAT=BEP_BTETAT) ';
  Result := Result + 'LEFT JOIN AFFAIRE   ON (AFF_AFFAIRE=BEP_AFFAIRE) ';
  Result := Result + 'LEFT JOIN TIERS     ON (T_TIERS=BEP_TIERS) ';
  Result := Result + 'LEFT JOIN ADRESSES  ON (ADR_REFCODE=AFF_AFFAIRE) ';
  Result := Result + 'LEFT JOIN RESSOURCE ON (ARS_RESSOURCE=BEP_RESSOURCE) ';
  Result := Result + 'LEFT JOIN FONCTION  ON (AFO_FONCTION=ARS_FONCTION1) ';
  Result := Result + 'LEFT JOIN BTTYPERES ON (BTR_TYPRES=ARS_CHAINEORDO) ';

  Result := Result + 'WHERE BEP_CODEEVENT="' + NumEvent + '"';

end;

Initialization
  registerclasses ( [ TOF_BTEVENPLAN ] ) ;
end.

