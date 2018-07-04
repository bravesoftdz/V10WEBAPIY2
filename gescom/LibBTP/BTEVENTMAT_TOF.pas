{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTEVENTMAT_TOF ()
Mots clefs ... : TOF;BTEVENTMAT_TOF
*****************************************************************}
Unit BTEVENTMAT_TOF ;

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
  TOF_BTEVENTMAT = Class (TOF)
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
    Affecttiers   : Boolean;
    AffectAffaire : Boolean;
    AffectRessource : Boolean;
    Ok_GereConso  : Boolean;
    //
    HDeb          : TDateTime;
    HFin          : TDateTime;
    NbHJour       : Double;
    NbJourOuvre   : Integer;
    //
    OldDDeb       : TDateTime;
    OldDFin       : TDateTime;
    OldDuree      : Double;
    //
    Heuredeb      : TDateTime;
    HeureFin      : TDateTime;
    //
    Hpanel1       : THPanel;
    NumEventMat   : THLabel;
    //
    Hpanel2       : THPanel;
    CodeMateriel  : THEdit;
    LibCodeMat    : THLabel;
    CodeEtat      : THValComboBox;
    TypeAction    : THEdit;
    LibTypeAction : THLabel;
    LibelleAction : THEdit;
    DateDeb       : THEdit;
    DateFin       : THEdit;
    Dateimputation: THEdit;
    NbHeure       : THEdit;
    LibelleReal   : THEdit;
    //
    HPanel3       : THPanel;
    NatureAuxi    : THEdit;
    CodeClient    : THEdit;
    StJuridique   : THLabel;
    StAdresse1    : THLabel;
    StAdresse2    : THLabel;
    StAdresse3    : THLabel;
    StCP          : THLabel;
    STVille       : THLabel;
    //
    HPanel4       : THPanel;
    Affaire       : THEdit;
    Affaire0      : THEdit;
    Affaire1      : THEdit;
    Affaire2      : THEdit;
    Affaire3      : THEdit;
    Avenant       : THEdit;
    LibAffaire    : THLabel;
    StDateImputation : THLabel;
    //
    HPanel5       : THPanel;
    Ressource     : THEdit;
    NomRessource  : THLabel;
    //
    Hpanel6       : THPanel;
    PA            : THNumEdit;
    PR            : THNumEdit;
    PV            : THNumEdit;
    Devise        : String;
    //
    Hpanel7       : THPanel;
    StRefPiece    : THLabel;
    RefPiece      : THEdit;
    //
    BTSelect      : TToolbarButton97;
    BTSelect1     : TToolbarButton97;
    BTSelect2     : TToolbarButton97;
    BTEfface      : TToolbarButton97;
    BTEfface1     : TToolBarButton97;
    BTEfface2     : TToolBarButton97;
    BTDelete      : TToolBarButton97;
    BTDuplication : TToolBarButton97;
    //
    BPlanning     : TToolBarButton97;
    BInsert       : TToolBarButton97;
    BValider      : TToolBarButton97;
    //
    TobEventMat   : Tob;
    TobMateriel   : TOb;
    TobCalendrier : TOB;
    TobJFerie     : TOB;
    //
    StSQL         : String;
    //
    FF            : String;
    //
    QQ            : TQuery;
    //
    TailleEcranMaxi : Integer;
    TailleEcranMini : Integer;
    //
    procedure AffaireOnExit(Sender: TObject);
    procedure AffEffaceOnClick(Sender: TObject);
    procedure AffSelectOnClick(Sender: TObject);
    procedure CalculNoCompteur;
    procedure ChargeEcranWithLaTob;
    procedure ChargeInfoAffaire;
    procedure ChargeInfoTiers;
    procedure ChargeInfoTypeAction;
    procedure ChargeZoneEcran;
    procedure ChargeZoneTOB;
    procedure CodeMaterielOnExit(Sender: TObject);
    procedure Controlechamp(Champ, Valeur: String);
    procedure ControleMateriel;
    procedure ControleRessource;
    function  ControleSaisie: Boolean;
    procedure CodeEtatOnChange(Sender: TObject);
    procedure CreateTOB;
    procedure DateDebOnExit(Sender: TObject);
    procedure DateFinOnExit(Sender: TObject);
    procedure DateImputationOnExit(Sender: TObject);
    procedure DestroyTOB;
    procedure GetObjects;
    procedure NBHeureOnExit(Sender: TObject);
    procedure OnDupliqueFiche(Sender: TObject);
    procedure RazZoneEcran;
    procedure RessourceEffaceOnClick(Sender: TObject);
    procedure RessourceOnExit(Sender: TObject);
    procedure RessourceSelectOnClick(Sender: TObject);
    procedure SetScreenEvents;
    procedure TiersEffaceOnClick(Sender: TObject);
    procedure TiersOnExit(Sender: TObject);
    procedure TiersSelectOnClick(Sender: TObject);
    procedure TypeActionOnExit(Sender: TObject);
    procedure TypeActionOnChange(Sender: TObject);



  end ;

const
  TexteMessage: array[1..22] of string  = (
          {1}   'Le Code Materiel est obligatoire'
          {2},  'Le Type d''action est Obligatoire'
          {3},  'Le Code Tiers est Obligatoire'
          {4},  'Le Code Affaire est Obligatoire'
          {5},  'Le code Ressource est Obligatoire'
          {6},  'Suppression impossible ce Matériel se trouve sur une Affectation'
          {7},  'La suppression a échoué'
          {8},  'Création Impossible, numéro non déterminé !'
          {9},  'Le Code Materiel est inexistant'
          {10}, 'Le Type d''action est inexistant'
          {11}, 'Le Code Tiers est inexistant'
          {12}, 'Le Code Affaire est inexistant'
          {13}, 'Le code Ressource est inexistant'
          {14}, 'La Mise à jour des consommations est impossible, la Prestation associée au Matériel est inexistante'
          {15}, 'La Mise à jour des consommations est impossible, le Tiers associé est inexistant'
          {16}, 'La Mise à jour des consommations est impossible, l''affaire associée est inexistante'
          {17}, 'La Mise à jour des consommations ne s''est pas effectuée'
          {18}, 'Le libellé de la réalisation est obligatoire si votre évènement est réalisé'
          {19}, 'Suppression de la consommation associée impossible'
          {20}, 'Le nombre d''heures est obligatoire si votre événement est réalisé'
          {21}, 'La date de début est supérieure à la date de fin'
          {22}, 'la date de fin est inférieure à la date de début'
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
     BTPUtil,
     UdateUtils;

procedure TOF_BTEVENTMAT.OnNew ;
begin
  Inherited ;

end;

procedure TOF_BTEVENTMAT.OnDelete ;
begin
  Inherited ;

  //contrôle si Famille non présente sur Matériel...
  if PGIAsk('Confirmez-vous la suppression de cet évènement ?', 'Saisie Evénement') = Mryes then
  begin
    StSQL := 'DELETE BTEVENTMAT WHERE BEM_IDEVENTMAT=' + NumEventMat.caption;
    if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[3], 'Evénement Matériel');

    OnClose;
  end;

end ;

procedure TOF_BTEVENTMAT.OnUpdate ;
begin
  Inherited ;

  if ControleSaisie then
  begin
    Ecran.ModalResult := 1;

    //si la TOB est vide on fait rien et surtout cela signifie qu'on a un gros problème
    if TOBEventMat = nil then Exit;

    If Action = TaCreat then CalculNoCompteur;

    //mise à jour de la table des Evénements Parc/Matériel
    ChargeZoneTOB;

    //On met à jour la table des événement
    TOBeventMat.SetAllModifie(true);
    TobEventMat.InsertOrUpdateDB(true);

    //Mise à jour des consommations si événement Réalisé
    if (CodeEtat.Value = 'REA') then TraitementRealise(TobEventMat);

    If (AppelPlan = true) and (LaTob <> nil) then
    begin
      LaTOB.PutValue('BEM_IDEVENTMAT',  TobEventMat.GetValue('BEM_IDEVENTMAT'));
      LaTob.PutValue('BEM_CODEMATERIEL',CodeMateriel.Text);
      LaTob.PutValue('BEM_BTETAT',      TypeAction.text);
      LaTob.PutValue('BEM_CODEETAT',    CodeEtat.Value);
      LaTob.PutValue('BEM_DATEDEB',     DateDeb.text);
      LaTob.PutValue('BEM_DATEFIN',     DateFin.text);
      LaTob.PutValue('BEM_NBHEURE',     NbHeure.text);
      LaTob.PutValue('BEM_AFFAIRE',     Affaire.text);
      LaTob.PutValue('BEM_TIERS',       CodeClient.text);
      LaTob.PutValue('BEM_RESSOURCE',   Ressource.text);
      LaTob.PutValue('RETOUR',          1);
    end;
    //Ecran.ModalResult := 0;
  end
  else
  begin
    If (AppelPlan) and (LaTob <> nil) then
    begin
      Ecran.ModalResult := 0;
      LaTob.PutValue('RETOUR',            0);
      exit;
    end;
  end;

end;


Function TOF_BTEVENTMAT.ControleSaisie : Boolean;
begin

  Result := False;

  if not CodeMaterielExist(CodeMateriel.Text, '') then
  begin
    PGIError(TexteMessage[9], 'Erreur Saisie');
    CodeMateriel.SetFocus;
    exit;
  end;

  if not TypeActionExist(TypeAction.text, ' AND BTA_TYPEACTION="PMA"') then
  begin
    PGIError(TexteMessage[10], 'Erreur Saisie');
    TypeAction.SetFocus;
    exit;
  end;

  if AffectTiers then
  begin
    if not TiersExist(CodeClient.text, ' AND T_NATUREAUXI="CLI"') then
    begin
      PGIError(TexteMessage[11], 'Erreur Saisie');
      CodeClient.SetFocus;
      exit;
    end;
  end;

  If AffectRessource then
  begin
    if not RessourceExist(Ressource.text, ' AND ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")') then
    begin
      PGIError(TexteMessage[13], 'Erreur Saisie');
      Ressource.SetFocus;
      exit;
    end;
  end;

  if AffectAffaire then
  begin
    if not AffaireExist(Affaire.text, ' AND AFF_AFFAIRE0="' + Affaire0.text + '"') then
    begin
      PGIError(TexteMessage[12], 'Erreur Saisie');
      exit;
    end;
  end;

  if (LibelleReal.text = '') And (CodeEtat.Value = 'REA') then
  begin
    PGIError(TexteMessage[18], 'Erreur Saisie');
    LibelleReal.SetFocus;
    exit;
  end;

  if StrToDate(DateDeb.text) > StrToDate(Datefin.text) then
  begin
    PGIError(TexteMessage[21], 'Erreur Saisie');
    DateDeb.SetFocus;
    exit;
  end;

  if StrToDate(DateFin.text) < StrToDate(DateDeb.text) then
  begin
    PGIError(TexteMessage[22], 'Erreur Saisie');
    DateFin.SetFocus;
    exit;
  end;

  if ((NbHeure.text = '') Or (NbHeure.Text = '0')) And (CodeEtat.Value = 'REA') then
  begin
    PGIError(TexteMessage[20], 'Erreur Saisie');
    NbHeure.SetFocus;
    exit;
  end;

  Result := true;

end ;

Procedure TOF_BTEVENTMAT.OnLoad ;
begin
  Inherited ;

  if NumEventMat.caption = 'O' then
  Begin
    if      AffectAffaire    then Affaire1.SetFocus
    else if Affecttiers      then CodeClient.setfocus
    else if AffectRessource  then Ressource.SetFocus;
    exit;
  end;

  StSQL := 'SELECT * FROM BTEVENTMAT WHERE BEM_IDEVENTMAT=' + NumEventMat.caption;
  QQ := OpenSQL(StSQL, True);

  If not QQ.Eof then
  begin
    TobEventMat.SelectDB('BTEVENTMAT', QQ);
    ChargeZoneEcran;
    if TobEventMat.GetValue('BEM_NUMMOUV') <> 0 then
    Begin
      BInsert.Visible   := False;
      BValider.Visible  := False;
      BTDelete.Visible  := False;
    end;
  end;

  ferme(QQ);

  if      AffectAffaire    then Affaire1.SetFocus
  else if Affecttiers      then CodeClient.setfocus
  else if AffectRessource  then Ressource.SetFocus;

end ;

procedure TOF_BTEVENTMAT.OnArgument (S : String ) ;
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

  //effacement des différents panels
  Hpanel3.Visible := False;
  Hpanel4.Visible := False;
  Hpanel5.Visible := False;
  Hpanel6.Visible := False;
  Hpanel7.Visible := False;
  //
  TailleEcranMaxi := Ecran.Height;
  TailleEcranMini := TailleEcranMaxi - (HPanel3.Height + HPanel4.Height + HPanel5.Height + Hpanel6.Height + Hpanel7.Height);
  //
  Ecran.Height    := TailleEcranMini;
  //
  //Chargement de la table des jours fériés....
  ChargeJoursFeries(TobJferie);
  //
  //Heure de début de journée, heure de fin de journée et nombre d'heure par jour
  //
  Hdeb := GetParamSocSecur('SO_BTAMDEBUT', '08:30');
  HFin := GetParamSocSecur('SO_BTPMFIN',   '17:30');

  //Voir si il faudrait pas à un moment utiliser les calendriers...
  NbHJour := ChargeParametreTemps;
  //
  TypeAction.plus := 'BTA_TYPEACTION="PMA"';
  //
  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.text, False);
  //
  if Action = Tacreat then
  begin
    RazZoneEcran;
    if LaTOB <> nil then ChargeEcranwithLaTob;
    //
    if CodeMateriel.Text <> '' then
    begin
      CodeMateriel.ElipsisButton:= False;
      CodeMateriel.Enabled      := False;
      ControleMateriel;
      if (AppelPlan = False) Or (LaTOB = nil) then
      begin
        DateDeb.Text := DateToStr(Now);
        Datefin.Text := DateToStr(Now+1);
      end;
    end
    else
    begin
      CodeMateriel.Enabled      := True;
      CodeMateriel.ElipsisButton:= True;
    end;
    BTDuplication.Visible := False;
  end
  else
  begin
    CodeMateriel.Enabled      := False;
    CodeMateriel.ElipsisButton:= False;
    BTDuplication.Visible     := True;
  end;

end ;

procedure TOF_BTEVENTMAT.OnClose ;
begin
  Inherited ;

  DestroyTOB;

end ;

procedure TOF_BTEVENTMAT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTEVENTMAT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTEVENTMAT.GetObjects;
begin

  NumEventMat   := THLabel(Getcontrol('BEM_IDEVENTMAT'));
  //
  CodeMateriel  := THEdit(Getcontrol('BEM_CODEMATERIEL'));
  LibCodeMat    := THLabel(Getcontrol('STMATERIEL'));
  CodeEtat      := THValComboBox(Getcontrol('BEM_CODEETAT'));
  TypeAction    := THEdit(Getcontrol('BEM_TYPEACTION'));
  LibTypeAction := THLabel(Getcontrol('STLIBTYPEACTION'));
  LibelleAction := THEdit(Getcontrol('BEM_LIBACTION'));
  DateDeb       := THEdit(Getcontrol('DATEEVT'));
  DateFin       := THEdit(Getcontrol('DATEEVT_'));
  DateImputation:= THEdit(Getcontrol('DATEIMPUTATION'));
  NbHeure       := THEdit(Getcontrol('BEM_NBHEURE'));
  LibelleReal   := THEdit(Getcontrol('BEM_LIBREALISE'));
  //
  NatureAuxi    := THEdit(Getcontrol('BEM_NATUREAUXI'));
  CodeClient    := THEdit(Getcontrol('BEM_TIERS'));
  StJuridique   := THLabel(Getcontrol('STJURIDIQUE'));
  StAdresse1    := THLabel(Getcontrol('STADRESSE1'));
  StAdresse2    := THLabel(Getcontrol('STADRESSE2'));
  StAdresse3    := THLabel(Getcontrol('STADRESSE3'));
  StCP          := THLabel(Getcontrol('STCP'));
  STVille       := THLabel(Getcontrol('STVILLE'));
  //
  Affaire       := THEdit(Getcontrol('BEM_AFFAIRE'));
  Affaire0      := THEdit(Getcontrol('BEM_AFFAIRE0'));
  Affaire1      := THEdit(Getcontrol('BEM_AFFAIRE1'));
  Affaire2      := THEdit(Getcontrol('BEM_AFFAIRE2'));
  Affaire3      := THEdit(Getcontrol('BEM_AFFAIRE3'));
  Avenant       := THEdit(Getcontrol('BEM_AVENANT'));
  LibAffaire    := THLabel(Getcontrol('LIBAFFAIRE'));
  StDateImputation := THLabel(Getcontrol('STDATEIMPUTATION'));
  //
  Ressource     := THEdit(Getcontrol('BEM_RESSOURCE'));
  NomRessource  := THLabel(Getcontrol('NOMRESSOURCE'));
  //
  PA            := THNumEdit(Getcontrol('BEM_PA'));
  PR            := THNumEdit(Getcontrol('BEM_PR'));
  PV            := THNumEdit(Getcontrol('BEM_PV'));
  //
  StRefPiece    := THLabel(GetControl('STREFPIECE'));
  RefPiece      := THEdit(Getcontrol('BEM_REFPIECE'));
  //
  BTSelect      := TToolBarButton97(Getcontrol('BSELECT'));
  BTSelect1     := TToolBarButton97(Getcontrol('BSELECT1'));
  BTSelect2     := TToolBarButton97(Getcontrol('BSELECT2'));
  //
  BTEfface      := TToolBarButton97(Getcontrol('BEFFACE'));
  BTEfface1     := TToolBarButton97(Getcontrol('BEFFACE1'));
  BTEfface2     := TToolBarButton97(Getcontrol('BEFFACE2'));
  //
  BtDelete      := TToolBarButton97(GetCONTROL('BDELETE'));
  //
  BPlanning     := TToolBarButton97(GetCONTROL('BPLANNING'));
  BInsert       := TToolBarButton97(GetCONTROL('BINSERT'));
  BValider      := TToolBarButton97(GetCONTROL('BVALIDER'));
  BTDuplication := TToolBarButton97(GetCONTROL('BDUPLICATION'));

  Hpanel1       := THPanel(GetControl('HPANEL1'));
  Hpanel2       := THPanel(GetControl('HPANEL2'));
  Hpanel3       := THPanel(GetControl('HPANEL3'));
  Hpanel4       := THPanel(GetControl('HPANEL4'));
  Hpanel5       := THPanel(GetControl('HPANEL5'));
  Hpanel6       := THPanel(GetControl('HPANEL6'));
  Hpanel7       := THPanel(GetControl('HPANEL7'));
  //

end;

procedure TOF_BTEVENTMAT.SetScreenEvents;
begin
  //
  Affaire1.Onexit         := AffaireOnExit;
  Affaire2.OnExit         := AffaireOnExit;
  Affaire3.OnExit         := AffaireOnExit;
  //
  CodeMateriel.OnExit     := CodeMaterielOnExit;
  TypeAction.OnExit       := TypeActionOnExit;
  CodeClient.OnExit       := TiersOnExit;
  Ressource.OnExit        := RessourceOnExit;
  //
  BTSelect.OnClick        := AffSelectOnClick;
  BTEfface.OnClick        := AffEffaceOnclick;
  //
  BTSelect1.OnClick       := TiersSelectOnClick;
  BTEfface1.OnClick       := TiersEffaceOnclick;
  //
  BTSelect2.OnClick       := RessourceSelectOnClick;
  BTEfface2.OnClick       := RessourceEffaceOnclick;
  //
  BTDuplication.OnClick   := OnDupliqueFiche;
  //
  DateDeb.OnExit          := DateDebOnExit;
  DateFin.OnExit          := DateFinOnExit;
  Dateimputation.OnExit   := DateImputationOnExit;

  NBHeure.OnExit          := NBHeureOnExit;

  TypeAction.OnChange     := TypeActionOnChange;
  CodeEtat.OnChange       := CodeEtatOnChange;

end;

Procedure TOF_BTEVENTMAT.AffaireOnExit(Sender : TObject);
var IP      : Integer;
begin

  Affaire.text := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, CodeClient.text, Action, False, True, false, IP);

  if Affaire.text <> '' then  ChargeInfoAffaire;

end;

Procedure TOF_BTEVENTMAT.TiersOnExit(Sender : TObject);
begin

  if CodeClient.text <> '' then  ChargeInfoTiers;

end;

Procedure TOF_BTEVENTMAT.RessourceOnExit(Sender : TObject);
begin

  if Ressource.text <> '' then ControleRessource;

end;

Procedure TOF_BTEVENTMAT.ControleRessource;
Var StWhere : string;
begin

  if ressource.text = '' then exit;

  StWhere := ' AND ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")';

  NomRessource.caption := ChargeinfoRessource(Ressource.text, StWhere, True);

  if NomRessource.Caption = '' then Ressource.setfocus;

  RechercheCalendrier(ressource.text, Tobcalendrier);

end;

Procedure  TOF_BTEVENTMAT.ControleMateriel;
Var codeRes : String;
begin

  if CodeMateriel.text <> '' then ChargeInfoMateriel(CodeMateriel.Text, TobMateriel);

  LibCodeMat.Caption  := TobMateriel.GetString('BMA_LIBELLE');

  if Action = Tacreat then PR.text :=  FormatFloat(FF,TobMateriel.GetVALUE('BMA_COUT'))
  else if (Action = Tamodif) then
  begin
    if (StrtoFloat(PR.text)=0) then PR.text := FormatFloat(FF,TobMateriel.GetVALUE('BMA_COUT'));
  end;

  PA.Text             := FormatFloat(FF, TobMateriel.GetVALUE('BMA_PA'));
  PV.Text             := FormatFloat(FF,TobMateriel.GetVALUE('BMA_PV'));

  //gestion du calendrier si matériel associé à une ressource
  if TobMateriel.GetValue('BMA_RESSOURCE') <> '' then
  begin
    codeRes := TobMateriel.GetValue('BMA_RESSOURCE');
    RechercheCalendrier(CodeRes,TobCalendrier);
  end;

end;

Procedure TOF_BTEVENTMAT.CodeMaterielOnExit(Sender : TObject);
begin

  if CodeMateriel.text <> '' then
  begin
    ControleMateriel;
    if LibCodeMat.caption = '' then
    begin
      CodeMateriel.Text := '';
      CodeMateriel.SetFocus;
    end;

  end
  else CodeMateriel.Setfocus;

end;

Procedure TOF_BTEVENTMAT.TypeActionOnExit(Sender : TObject);
begin

  if TypeAction.text <> '' then
  begin
    Ecran.Height := TailleEcranMini;
    //
    hpanel3.Visible := False;
    hpanel4.Visible := False;
    hpanel5.Visible := False;
    hpanel6.Visible := False;
    hpanel7.Visible := False;
    //
    ChargeInfoTypeAction;
    if TypeAction.text = '' then TypeAction.SetFocus;
  end
  else
  begin
    if TypeAction.Text = '' then TypeAction.SetFocus;
  end;

end;

procedure TOF_BTEVENTMAT.DateDebOnExit(Sender: TObject);
Var Duree : Double;
    DDeb  : TDateTime;
    DFin  : TDateTime;
begin

  DDeb    := StrToDate(DateDeb.Text);
  DFin    := StrToDate(DateFin.Text);
  Duree   := StrToFloat(NbHeure.Text);

  if DDeb > Dfin then
  begin
    PGIError(TexteMessage[21], 'Erreur Saisie');
    DateDeb.Text := DateToStr(OldDDeb);
    Datedeb.SetFocus;
    Exit;
  end;

  if Ddeb = OldDdeb then exit;

  Duree := CalculDureeEvent(DDeb, DFin, TobJFerie, TobCalendrier);
  NbHeure.Text  := FloatToStr(Duree/60);

  OldDDeb := DDeb;

end;

procedure TOF_BTEVENTMAT.DateFinOnExit(Sender: TObject);
Var DDeb : TDateTime;
    DFin : TDateTime;
    Duree: Double;
begin

  DDeb := StrToDate(DateDeb.text);
  DFin := StrToDate(DateFin.text);
  Duree:= StrToFloat(Nbheure.Text);

  if DFin < DDeb then
  begin
    PGIError(TexteMessage[22], 'Erreur Saisie');
    DateFin.Text := DateToStr(OldDFin);
    DateFin.SetFocus;
    exit;
  end;

  if not ControleCoherence(ddeb,DFin,Duree, CodeMateriel.Text, StrToInt(NumEventMat.caption)) then exit;

  if DFin = idate2099 then exit;

  if DFin = OldDFin   then exit;

  //if OldDFin <> idate2099 then NbM := CalculDureeEvent(DDeb, DFin);
  if DFin <> idate2099 then Duree := CalculDureeEvent(DDeb, DFin, TobJFerie, TobCalendrier);

  NbHeure.Text  := FloatToStr(Duree/60);

  OldDFin := DFin;

end;

Procedure TOF_BTEVENTMAT.DateImputationOnExit(Sender : TObject);
begin
  TobEventMat.AddChampSupValeur('DATEIMPUTATION', DateImputation.text);
end;

procedure TOF_BTEVENTMAT.NBHeureOnExit(Sender: TObject);
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


Procedure TOF_BTEVENTMAT.CodeEtatOnChange(Sender : TObject);
begin

  if CodeEtat.Value = 'REA' then
  begin
    LibelleReal.Text := LibelleAction.Text;
    If ok_GereConso then
    begin
      StDateImputation.visible  := True;
      Dateimputation.Visible    := True;
      if StrToDate(DateFin.text) <= Now then
        Dateimputation.Text := DateFin.text
      else
        Dateimputation.Text := DateToStr(Now);
      Dateimputation.SetFocus;
    end;
  end
  else
  begin
    LibelleReal.text := '';
    StDateImputation.visible  := False;
    Dateimputation.Visible    := False;
  end;

end;

Procedure TOF_BTEVENTMAT.AffSelectOnClick(Sender : TObject);
Var StChamps  : String;
begin

  StChamps  := Affaire.Text;

  if GetAffaireEnteteSt(Affaire0, Affaire1, Affaire2, Affaire3, AVENANT, nil, StChamps, false, false, false, True, true,'') then Affaire.text := StChamps;

  ChargeInfoAffaire;

end;

Procedure TOF_BTEVENTMAT.TiersSelectOnClick(Sender : TObject);
Var StChamps  : string;
Begin

  StChamps  := CodeClient.Text;

  DispatchRecherche(CodeClient, 2, 'T_NATUREAUXI="CLI"','T_TIERS=' + Trim(CodeClient.Text), '');

  ChargeInfoTiers;

end;
Procedure TOF_BTEVENTMAT.RessourceSelectOnClick(Sender : TObject);
Var StChamps  : string;
    StWhere   : string;
begin

  StChamps  := ressource.Text;

  StWhere := 'ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")';

  DispatchRecherche(Ressource, 3, stwhere,'ARS_RESSOURCE=' + Trim(Ressource.Text), '');

  ControleRessource;

  //if Ressource.text <> '' then NomRessource.caption := ChargeinfoRessource(Ressource.text,  'AND ' + StWhere, True);

end;
//
Procedure TOF_BTEVENTMAT.AffEffaceOnClick(Sender : TObject);
begin

  Affaire.Text  := '';
  Affaire0.Text := '';
  Affaire1.Text := '';
  Affaire2.Text := '';
  Affaire3.Text := '';
  Avenant.Text  := '';
  //
  LibAffaire.caption := '...';

end;
Procedure TOF_BTEVENTMAT.TiersEffaceOnClick(Sender : TObject);
begin

  CodeClient.Text  := '';
  StJuridique.Caption := '...';

end;
Procedure TOF_BTEVENTMAT.RessourceEffaceOnClick(Sender : TObject);
begin

  Ressource.Text  := '';
  NomRessource.Caption := '...';

end;

Procedure TOF_BTEVENTMAT.Controlechamp(Champ, Valeur : String);
Var Date1 : TdateTime;
    Date2 : TDateTime;
begin

  if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if Champ='IDEVENEMENT' then NumEventMat.caption := Valeur;

  if Champ='MATERIEL'    then CodeMateriel.Text   := Valeur;

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

Procedure TOF_BTEVENTMAT.CreateTOB;
begin

  TobEventMat := Tob.Create('BTEVENTMAT', nil, -1);

  TobMateriel := Tob.create('BTMATERIEL', nil, -1);

  TobCalendrier := Tob.create('CALENDRIER', nil, -1);

  TobJFerie     := Tob.create('JOURFERIE', nil, -1);

end;

procedure TOF_BTEVENTMAT.DestroyTOB;
begin

  FreeAndNil(TobEventMat);

  FreeAndNil(TobMateriel);

  FreeAndNil(TobCalendrier);

  FreeAndNil(TobJFerie);

end;

Procedure TOF_BTEVENTMAT.ChargeZoneEcran;
var	Temp 	  : Double;
begin

  NumEventMat.Caption := TobEventMat.GetString('BEM_IDEVENTMAT');
  CodeMateriel.Text   := TobEventMat.GetString('BEM_CODEMATERIEL');
      
  DateDeb.Text        := DateToStr(TobEventMat.GetValue('BEM_DATEDEB'));
  DateFin.Text        := DateToStr(TobEventMat.GetValue('BEM_DATEFIN'));
  //
  OldDDeb             := Trunc(StrToDate(DateDeb.Text));
  OldDFin             := Trunc(StrtoDate(DateFin.Text));
  //
  Heuredeb            := TobEventMat.GetDateTime('BEM_HEUREDEB');
  HeureFin            := TobEventMat.GetDateTime('BEM_HEUREFIN');
  //
  if HeureDeb = 0 then HeureDeb := Hdeb;
  if HeureFin = 0 then HeureFin := HFin;
  //
  TypeAction.Text     := TobEventMat.GetString('BEM_BTETAT');
  Ok_GereConso        := GereConso(TypeAction.Text);
  //
  NatureAuxi.Text     := TobEventMat.GetString('BEM_NATUREAUXI');
  CodeClient.Text     := TobEventMat.GetString('BEM_TIERS');

  Affaire.Text        := TobEventMat.GetString('BEM_AFFAIRE');

  Ressource.Text      := TobEventMat.GetString('BEM_RESSOURCE');
  //
  PA.Text             := FormatFloat(FF,TobEventMat.GetValue('BEM_PA'));
  PR.Text             := FormatFloat(FF,TobEventMat.GetValue('BEM_PR'));
  PV.Text             := FormatFloat(FF,TobEventMat.GetValue('BEM_PV'));
  //
  Devise              := TobEventMat.GetString('BEM_Devise');

  //Chargement de la durée mini
  Temp                := TobEventMat.GetValue('BEM_NBHEURE');
  NbHeure.text        := FloatToStr(Temp); //TimeToStr(Duree);
  OldDuree            := Temp;
  //
  CodeEtat.Value      := TobEventMat.GetString('BEM_CODEETAT');
  //
  ControleMateriel;
  //
  if TypeAction.text <> '' then ChargeInfoTypeAction;
  LibelleAction.Text  := TobEventMat.GetString('BEM_LIBACTION');

  if Affecttiers Then
  begin
    if CodeClient.text <> '' then ChargeInfoTiers;
  end;
  //
  if AffectAffaire then
  begin
    if Affaire.text <> '' then ChargeInfoAffaire;
  end;

  if AffectRessource then ControleRessource;
  //
  RefPiece.Text       := TobEventMat.GetString('BEM_REFPIECE');
  if refpiece.text <> '' then
  begin
    HPanel7.Visible     := True;
    Ecran.Height        := Ecran.Height + HPanel7.Height;
    StRefPiece.Visible  := True;
    Refpiece.visible    := true;
  end
  else
  begin
    HPanel7.Visible     := False;
    StRefPiece.visible  := False;
    Refpiece.visible    := False;
  end;
  //
  LibelleReal.Text    := TobEventMat.GetString('BEM_LIBREALISE');
  //
  //si evénement réalisé et numéro de conso renseigné on récupère la date de mouvement pour la mettre en date d'imputation
  //la zone sera dès lors enabled à false...
  //
  StSQL := 'SELECT BCO_DATEMOUV FROM CONSOMMATIONS WHERE BCO_NUMMOUV=' + FloatToStr(TobEventMat.GetValue('BEM_NUMMOUV'));
  QQ    := OpenSQL(StSQL, False);
  If not QQ.Eof then
  begin
    StDateImputation.visible  := True;
    Dateimputation.Visible    := True;
    DateImputation.Enabled    := False;
    DateImputation.Text       := QQ.findField('BCO_DATEMOUV').AsString;
    CodeEtat.Enabled          := False;
    TobEventMat.AddChampSupValeur('DATEIMPUTATION', DateImputation.text);
  end
  else
  begin
    StDateImputation.visible  := False;
    Dateimputation.Visible    := False;
  end;
  Ferme(QQ);

end;

Procedure TOF_BTEVENTMAT.ChargeZoneTOB;
Var Prix      : Double;
    DureeMin  : Integer;
    Duree     : Double;
    Datedebut : TDatetime;
    Datefinal : TDateTime;
    An, Mois, Jour  : Word;
    H, M, S, Msec   : Word;
begin

  TobEventMat.PutValue('BEM_IDEVENTMAT',    NumEventMat.Caption);
  TobEventMat.PutValue('BEM_CODEMATERIEL',  CodeMateriel.Text);
  //
  //Calcul de l'heure de fin en fonction du nombre d'heure saisie et de la date de début
  //
  DateDebut := StrToDate(DateDeb.Text) + HeureDeb;
  Datefinal := StrToDate(DateFin.Text) + HeureFin;
  Duree     := StrToFloat(NbHeure.text);
  DureeMin  := Round(Duree * 60);
  Datefinal := CalculDateFinParcMat(DateDebut, DateFinal, NbHJour, DureeMin, TobJFerie);
  //
  TobEventMat.PutValue('BEM_DATEDEB',       DateDebut);
  TobEventMat.PutValue('BEM_DATEFIN',       Datefinal);
  //
  DecodeDateTime(DateDebut, An, Mois, Jour, H, M, S, Msec);
  TobEventMat.PutValue('BEM_HEUREDEB',      TimeToStr(EncodeTime(H, M, 0, 0)));
  DecodeDateTime(Datefinal, An, Mois, Jour, H, M, S, Msec);
  TobEventMat.PutValue('BEM_HEUREFIN',      TimeToStr(encodeTime(H, M, 0, 0)));
  //
  TobEventMat.PutValue('BEM_BTETAT',        TypeAction.Text);
  TobEventMat.PutValue('BEM_LIBACTION',     LibelleAction.Text);
  //
  TobEventMat.PutValue('BEM_NATUREAUXI',    NatureAuxi.Text);
  TobEventMat.PutValue('BEM_TIERS',         CodeClient.Text);
  //
  TobEventMat.PutValue('BEM_AFFAIRE',       Affaire.Text);
  //
  Prix := StrToFloat(PA.Text);
  TobEventMat.PutValue('BEM_PA',            Prix);

  Prix := StrToFloat(PV.Text);
  TobEventMat.PutValue('BEM_PV',            Prix);

  Prix := StrToFloat(PR.Text);
  TobEventMat.PutValue('BEM_PR',            Prix);
  //
  TobEventMat.PutValue('BEM_DEVISE',        Devise);
  TobEventMat.PutValue('BEM_CODEETAT',      CodeEtat.Value);

  //Duree := StrToTime(NbHeure.Text);
  //Temp  := TimeToFloat(Duree, true);
  TobEventMat.PutValue('BEM_NBHEURE',       Nbheure.text);
  //
  TobEventMat.PutValue('BEM_RESSOURCE',     Ressource.Text);
  //
  TobEventMat.PutValue('BEM_REFPIECE',      RefPiece.Text);
  TobEventMat.PutValue('BEM_LIBREALISE',    LibelleReal.Text);
  //
end;

Procedure TOF_BTEVENTMAT.ChargeInfoTypeAction;
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

  TOBTypeAction := TOB.Create('BTETAT',nil, -1);

  Stwhere := ' AND BTA_TYPEACTION="PMA"';

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
  if LibelleAction.text = '' then LibelleAction.text := LibTypeAction.caption;
  //
  AffectRessource := TOBTypeAction.GetBoolean('BTA_AFFECTRESS');
  AffectAffaire   := TOBTypeAction.GetBoolean('BTA_AFFECTCHANTIER');
  Affecttiers     := TOBTypeAction.GetBoolean('BTA_AFFECTIERS');

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

  if TOBTypeAction.GetValue('BTA_GESTIONCONSO') = 'X' then Ok_GereConso := True else Ok_GereConso := False;

  If TOBTypeAction.GetValue('BTA_VALORISE') = 'X' then
  begin
    HPanel6.Visible := True;
    Ecran.Height := (Ecran.Height + HPanel6.Height);
  end
  Else HPanel6.Visible := False;

  if Affecttiers Then
  begin
    if Action = Tacreat then
    begin
      CodeClient.text := TOBTypeAction.GetString('BTA_TIERS');
      if CodeClient.Text <> '' then ChargeInfoTiers;
    end;
    HPanel3.Visible := True;
    Ecran.Height := (Ecran.Height + HPanel3.Height);
  end
  else HPanel3.Visible := False;

  if AffectAffaire Then
  begin
    if Action = Tacreat then
    begin
      Affaire.text  := TOBTypeAction.GetString('BTA_AFFAIRE');
      ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.text, False);
      if Affaire.Text <> '' then ChargeInfoAffaire;
    end;
    HPanel4.Visible := True;
    Ecran.Height    := (Ecran.Height + HPanel4.Height);
  end
  Else HPanel4.Visible := False;

  if AffectRessource Then
  begin
    if Action = Tacreat then
    begin
      Ressource.text  := TOBTypeAction.GetString('BTA_RESSOURCE');
      ControleRessource;
    end;
    HPanel5.Visible   := True;
    Ecran.Height      := Ecran.Height + HPanel5.Height;
  end
  else HPanel5.Visible := False;

  FreeAndNil(TOBTypeAction);

  Ecran.Left  :=(Screen.Width-Ecran.Width)  div 2;
  Ecran.Top   :=(Screen.Height-Ecran.Height) div 2;

  ferme(QQ);

end;

procedure TOF_BTEVENTMAT.ChargeInfoAffaire;
begin

  StSQL := 'SELECT AFF_LIBELLE, AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE = "' + Affaire.Text + '"';
  QQ    := OpenSQL(StSQL, False);

  if QQ.Eof then
    LibAffaire.Caption := ''
  else
  begin
    ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.text, False);
    LibAffaire.caption := QQ.findfield('AFF_LIBELLE').AsString;
    if AffectAffaire then
    begin
      CodeClient.text := QQ.findfield('AFF_TIERS').AsString;
      ChargeInfoTiers;
    end;
  end;

  ferme(QQ);

end;

procedure TOF_BTEVENTMAT.ChargeInfoTiers;
begin

  StSQL := 'SELECT * FROM TIERS WHERE T_TIERS = "' + CodeClient.Text + '" AND T_NATUREAUXI="' + NatureAuxi.Text + '"';
  QQ    := OpenSQL(StSQL, False);

  if QQ.Eof then
    StJuridique.Caption := ''
  else
  begin
    IF QQ.findfield('T_PARTICULIER').AsString = 'E' then
      StJuridique.caption := QQ.findfield('T_JURIDIQUE').AsString + ' ' + QQ.findfield('T_LIBELLE').AsString + ' '
    else
      StJuridique.caption := QQ.findfield('T_JURIDIQUE').AsString + ' ' + QQ.findfield('T_LIBELLE').AsString + ' ' + QQ.findfield('T_PRENOM').AsString;
    //chargement de l'adresse
    StAdresse1.Caption    := QQ.findfield('T_ADRESSE1').AsString;
    StAdresse2.Caption    := QQ.findfield('T_ADRESSE2').AsString;
    StAdresse3.Caption    := QQ.findfield('T_ADRESSE3').AsString;
    StCP.Caption          := QQ.findfield('T_CODEPOSTAL').AsString;
    StVille.Caption       := QQ.findfield('T_VILLE').AsString;
  end;

  ferme(QQ);

end;

procedure TOF_BTEVENTMAT.RazZoneEcran;
begin

  LibCodeMat.caption    := '...';
  CodeEtat.value        := 'ARE';
  TypeAction.text       := '';
  LibTypeAction.caption := '...';
  LibelleAction.text    := '';
  DateDeb.Text          := DateToStr(now);
  DateFin.Text          := DateToStr(idate2099);
  //
  Heuredeb              := HDeb;
  HeureFin              := HFin;
  //
  OldDDeb               := StrToDate(DateDeb.Text);
  OldDFin               := StrtoDate(Datefin.Text);
  OldDuree              := 1;
  //
  NbHeure.text          := '1';
  LibelleReal.text      := '';
  //
  NatureAuxi.text       := 'CLI';
  CodeClient.text       := '';
  StJuridique.caption   := '...';
  StAdresse1.caption    := '...';
  StAdresse2.caption    := '...';
  StAdresse3.caption    := '...';
  StCP.caption          := '...';
  STVille.caption       := '...';
  //
  Affaire.text          := '';
  Affaire0.text         := '';
  Affaire1.text         := '';
  Affaire2.text         := '';
  Affaire3.text         := '';
  Avenant.text          := '';
  LibAffaire.caption    := '...';
  //
  Ressource.text        := '';
  NomRessource.caption  := '...';
  //
  PA.text               := FormatFloat(FF,0);
  PR.text               := FormatFloat(FF,0);
  PV.text               := FormatFloat(FF,0);
  //
  Devise                := V_PGI.DevisePivot;
  //
  RefPiece.text         := '';
  //
  BtDelete.visible      := False;
  BPlanning.visible     := False;
  //
end ;

procedure TOF_BTEVENTMAT.CalculNoCompteur;
Var NumAff : Integer;
begin       

  if not GetNumCompteur('BEM',iDate1900, NumAff) then
  begin
    PGIError(TexteMessage[7], 'Evènement Matériel');
    OnClose;
  end
  else NumEventMat.Caption := IntToStr(NumAff);

end;

Procedure Tof_BTEVENTMAT.ChargeEcranWithLaTob;
Var Duree : Double;
begin

  //
  CodeMateriel.Text := LaTob.GetValue('BEM_CODEMATERIEL');
  CodeEtat.Value    := LaTob.GetValue('BEM_CODEETAT');
  TypeAction.Text   := LaTob.GetValue('BEM_BTETAT');

  DateDeb.Text      := DateToStr(LaTob.GetValue('BEM_DATEDEB'));
  DateFin.Text      := DateToStr(LaTob.GetValue('BEM_DATEFIN'));

  HeureDeb          := Hdeb; //LaTob.GetDateTime('BEM_HEUREDEB');
  HeureFin          := HFin; //LaTob.GetDateTime('BEM_HEUREFIN');

  OldDDeb           := LaTob.GetValue('BEM_DATEDEB');
  OldDFin           := LaTob.GetValue('BEM_DATEFIN');

  OldDDeb := Trunc(OldDDeb) + HeureDeb;
  OldDFin := Trunc(OldDFin) + Heurefin;

  NbHeure.Text      := LaTob.GetValue('BEM_NBHEURE');

  Duree             := CalculDureeEvent(OldDDeb, OldDFin, TobJFerie, TobCalendrier);
  //if Duree <> StrToFloat(NbHeure.Text) then
  //begin
  Duree             := Round(Duree/60);
  NbHeure.Text      := FloatToStr(Duree);
  LaTob.PutValue('BEM_NBHEURE', NbHeure.text);
  OldDuree          := Duree;
  //end;

  Affaire.Text      := LaTob.GetValue('BEM_AFFAIRE');
  CodeClient.Text   := LaTob.GetValue('BEM_TIERS');
  Ressource.Text    := LaTob.GetValue('BEM_RESSOURCE');
  //

end;


Procedure TOF_BTEVENTMAT.OnDupliqueFiche(Sender : TObject);
begin

  if action = TaCreat then Exit;

  Action            := TaCreat;

  CodeMateriel.ElipsisButton:= True;
  CodeMateriel.Enabled      := True;

  CodeEtat.value            := 'ARE';

  OldDDeb := StrToDate(DateDeb.Text);
  OldDFin := StrtoDate(Datefin.Text);
  OldDuree:= StrToFloat(NbHeure.text);

  BTDuplication.Visible := False;
  DateDeb.SetFocus;
  //
end;


procedure TOF_BTEVENTMAT.TypeActionOnChange(Sender: TObject);
begin

end;

Initialization
  registerclasses ( [ TOF_BTEVENTMAT ] ) ;
end.

