{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTYPEACTIONMAT ()
Mots clefs ... : TOF;BTTYPEACTIONMAT
*****************************************************************}
Unit BTTYPEACTIONMAT_TOF ;

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
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     uTOB,
     HTB97,
     HPanel,
     UTOF,
     graphics,
     ExtCtrls;

Type
  TOF_BTTYPEACTIONMAT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Action          : TActionFiche;
    //
    TypeAction      : String;
    //
    GBASSOCIATION_INT : TGroupBox;
    GBASSOCIATION_PMA : TGroupBox;
    GBASSOCIATION_PCA : TGroupBox;
    PANEL2            : THPanel;
    //
    BtEtat          : THEdit;
    Libelle         : THEdit;
    Abrege          : THEdit;
    //
    Affaire         : THEdit;
    Affaire0        : THEdit;
    Affaire1        : THEdit;
    Affaire2        : THEdit;
    Affaire3        : THEdit;
    Avenant         : THEdit;
    //
    Tiers           : THEdit;
    Ressource       : THEdit;
    Duree           : THEdit;
    //
    StAffaire       : THLabel;
    StTiers         : THLabel;
    StRessource     : THLabel;
    //
    TBAM_Affaire    : THLabel;
    TBAM_Tiers      : THLabel;
    TBAM_Ressource  : THLabel;
    TBAM_Duree      : THLabel;
    //
    //Cadencement     : THValComboBox;
    //
    AffectChantier  : TCheckBox;
    AffectTiers     : TCheckBox;
    AffectRessource : TCheckBox;
    GerePlanning    : TCheckBox;
    Defaut          : TCheckBox;
    GereConso       : TCheckBox;
    SaisieValorisee : TCheckBox;
    CumulWorkingHour: TCheckBox;
    Modifiable      : TCheckBox;
    Surbook         : TCheckBox;
    //
    BTSelect        : TToolbarButton97;
    BTSelect1       : TToolbarButton97;
    BTSelect2       : TToolbarButton97;
    BTEfface        : TToolbarButton97;
    BTEfface1       : TToolBarButton97;
    BTEfface2       : TToolBarButton97;
    BTDelete        : TToolBarButton97;
    BFOND	 	        : TToolbarButton97;
    BCOULEUR	      : TToolbarButton97;
    BIMGICONE	      : TToolbarButton97;
    BFONTE		      : TToolbarButton97;

    NumeroIcone     : THEdit;
    ImageIcone      : TImage;

    LFOND			      : THEdit;
    LCOULEUR	      : THEdit;
    CouleurFond     : THEdit;
    Couleur         : THEdit;
    Fonte           : THEdit;

    LFONTE		      : THLabel;

    //
    TobTypeAction   : Tob;
    //
    StWhere         : String;
    StSQL           : String;
    //
    QQ              : TQuery;
    //
    procedure AffaireOnExit(Sender: TObject);
    procedure AffCheckOnClick(Sender: TObject);
    procedure AffEffaceOnClick(Sender: TObject);
    procedure AffSelectOnClick(Sender: TObject);
    procedure ChargeZoneEcran;
    procedure ChargeZoneTOB;
    procedure CliCheckOnClick(Sender: TObject);
    procedure Controlechamp(Champ, Valeur: String);
    procedure CreateTOB;
    procedure CtrlAffaireTiers;
    procedure CtrlTierAffaire;
    procedure DestroyTOB;
    procedure GetObjects;
    procedure LibelleOnExit(Sender: TObject);
    procedure RazZoneEcran;
    procedure ResCheckOnClick(Sender: TObject);
    procedure RessourceEffaceOnClick(Sender: TObject);
    procedure RessourceOnExit(Sender: TObject);
    procedure RessourceSelectOnClick(Sender: TObject);
    procedure SetScreenEvents;
    procedure TiersEffaceOnClick(Sender: TObject);
    procedure TiersOnExit(Sender: TObject);
    procedure TiersSelectOnClick(Sender: TObject);
    //
    procedure BFond_OnClick(Sender: TObject);
    procedure BFonte_OnClick(Sender: TObject);
    procedure BCouleur_OnClick(Sender: TObject);
	  procedure BImgIcone_OnClick(Sender: TObject);
    //
  end ;

const
	// libellés des messages
  TexteMessage: array[1..14] of string  = (
          {1}  'Le Type Action est obligatoire'
          {2}, 'Suppression impossible ce Type Action est utilisé sur un Evénement'
          {3}, 'La suppression a échoué'
          {4}, 'Client inexistant. Veuillez renseigner un client Valide'
          {5}, 'Client Fermé. Veuillez saisir un nouveau client'
          {6}, 'Client inexistant sur ce chantier. Veuillez saisir un client appartenant au chantier'
          {7}, 'Chantier inexistant pour ce client. Veuillez saisir un chantier valide'
          {8}, 'Chantier inexistant. Veuillez renseigner un chantier valide'
          {9}, 'Chantier fermé. Veuillez recommencer votre saisie'
         {10}, 'Ressource inexistante. Veuillez renseigner une ressource valide'
         {11}, 'Ressource fermée. Veuillez recommencer votre saisie avec une nouvelle ressource'
         {12}, 'Code chantier  Inexistant'
         {13}, 'Code Ressource Inexistant'
         {14}, 'Code Tiers Inexistant'

            );

Implementation

uses AffaireUtil,
     AGLInitGC,
     UtilsParc,
     HRListeIcone,
     MsgUtil,
     HeureUtil,
     FactUtil;

procedure TOF_BTTYPEACTIONMAT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT.OnDelete ;
begin
  Inherited ;

  if PGIASK('Confirmez-vous la suppression ?', 'Suppression Type Action')=MrNo then Exit;

  //FV1 - 03/04/2018 - FS#3031 - ECHAFAUDAGE SERVICE : Interdire de supprimer un type action si présent sur planning
  if TypeAction = 'PMA' then
  begin
    //contrôle si Famille non présente sur Matériel...
    StSQL := 'SELECT BEM_BTETAT FROM BTEVENTMAT WHERE BEM_BTETAT="' + BtEtat.Text + '"';
  end
  else if TypeAction = 'INT' then
  begin
    //contrôle si Famille non présente sur Matériel...
    StSQL := 'SELECT BEP_BTETAT FROM BTEVENPLAN WHERE BEP_BTETAT="' + BtEtat.Text + '"';
  end
  else if TypeAction = 'PCA' then
  begin
    //contrôle si Famille non présente sur Matériel...
    StSQL := 'SELECT BEC_BTETAT FROM BTEVENTCHA WHERE BEC_BTETAT="' + BtEtat.Text + '"';
  end;
  //
  
  if existeSQL(STSQL) then
    PGIError(TexteMessage[2], 'Type Action')
  else
  Begin
    //suppression pure et simple de l'enregistrement avec confirmation
    StSQL := 'DELETE BTETAT WHERE BTA_BTETAT="' + BtEtat.Text + '"' + StWhere;
    if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[3], 'Type Action');
  end;

  Ecran.ModalResult := 1;

  OnClose;

end ;

procedure TOF_BTTYPEACTIONMAT.OnUpdate ;
begin
  Inherited ;

  Ecran.ModalResult := 0;

  if BtEtat.text = '' then
  begin
    PGIError(TexteMessage[1], 'Type Action');
    Exit;
  end;

  IF AffectChantier.Checked then
  begin
    if not AffaireExist(Affaire.text, '') then
    begin
      PGIError(TexteMessage[12], 'Erreur Saisie');
      Affaire1.SetFocus;
      Exit;
    end;
  end;

  if AffectRessource.Checked then
  begin
    if not RessourceExist(Ressource.text, 'AND ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")') then
    begin
      PGIError(TexteMessage[13], 'Erreur Saisie');
      Ressource.SetFocus;
      Exit;
    end;
  end;

  ChargeZoneTOB;

  //mise à jour de la table Famille Matériel
  TobTypeAction.InsertOrUpdateDB(true);

  Ecran.ModalResult := 1;


end ;

procedure TOF_BTTYPEACTIONMAT.OnLoad ;
begin
  Inherited ;

  StSQL := 'SELECT * FROM BTETAT WHERE BTA_BTETAT="' + BtEtat.text + '"' + StWhere;
  QQ := OpenSQL(StSQL, True);

  If not QQ.Eof then
  begin
      TobTypeAction.SelectDB('BTETAT', QQ);
  end;

  ChargeZoneEcran;

  ferme(QQ);

end ;

procedure TOF_BTTYPEACTIONMAT.OnArgument (S : String ) ;
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

  //Gestion des évènement des zones écran
  SetScreenEvents;

  //Cadencement.Plus := 'CO_TYPE="ACT" AND CO_ABREGE<>"X"';

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.Text, False);

  Affaire1.Visible := false;
  Affaire2.Visible := false;
  Affaire3.Visible := false;
  Avenant.Visible  := false;

  if TypeAction = 'PMA' then
  begin
    GBASSOCIATION_INT.Visible := false;
    GBASSOCIATION_PMA.Visible := True;
    GBASSOCIATION_PCA.Visible := False;
    Panel2.visible            := True;
    Ecran.Caption             := 'Type Action Parc/Matériel';
  end
  else if TypeAction = 'INT' then
  begin
    GBASSOCIATION_INT.Visible := True;
    GBASSOCIATION_PCA.Visible := False;
    GBASSOCIATION_PMA.Visible := False;
    Panel2.visible            := False;
    Ecran.Caption             := 'Type Action Intervention';
  end
  else if TypeAction = 'PCA' then
  begin
    GBASSOCIATION_INT.Visible := false;
    GBASSOCIATION_PMA.Visible := False;
    GBASSOCIATION_PCA.Visible := False;
    Panel2.visible            := False;
    Duree.Visible             := False;
    CumulWorkingHour.Visible  := False;
    TBAM_Duree.Visible        := False;
    Modifiable.Visible        := False;
    BTEtat.Enabled            := True;
    Ecran.Caption             := 'Type Action Planning Chantier';
  end;

  //Contrôle si au moins un type d'action paramétré. Si oui defaut is enabled
  StSQL := 'SELECT BTA_BTETAT FROM BTETAT WHERE BTA_TYPEACTION="' + TypeAction + '" AND BTA_DEFAUT="X"';
  if ExisteSQL(StSQL) then Defaut.Enabled := False else Defaut.Enabled := True;

  if Action = Tacreat then
    RazZoneEcran
  else
    BtEtat.Enabled := false;

end ;

procedure TOF_BTTYPEACTIONMAT.OnClose ;
begin
  Inherited ;

  DestroyTOB;

end ;

procedure TOF_BTTYPEACTIONMAT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPEACTIONMAT.GetObjects;
begin

  GBASSOCIATION_INT := TGroupBox(Getcontrol('GBASSOCIATION_INT'));
  GBASSOCIATION_PMA := TGroupBox(Getcontrol('GBASSOCIATION_PMA'));
  GBASSOCIATION_PCA := TGroupBox(Getcontrol('GBASSOCIATION_PCA'));
  PANEL2            := THPanel(Getcontrol('PANEL2'));

  //Gestion du type Action Intervention
  BTEtat      := THEdit(Getcontrol('BAM_TYPEACTION'));
  Libelle     := THEdit(Getcontrol('BAM_LIBELLE'));
  Abrege      := THEdit(Getcontrol('BTA_ABREGE'));
  //
  Affaire     := THEdit(GetControl('BAM_AFFAIRE'));
  Affaire0    := THEdit(GetControl('BAM_AFFAIRE0'));
  Affaire1    := THEdit(GetControl('BAM_AFFAIRE1'));
  Affaire2    := THEdit(GetControl('BAM_AFFAIRE2'));
  Affaire3    := THEdit(GetControl('BAM_AFFAIRE3'));
  Avenant     := THEdit(GetControl('BAM_AVENANT'));
  //
  Tiers       := THEdit(Getcontrol('BAM_TIERS'));
  Ressource   := THEdit(Getcontrol('BAM_RESSOURCE'));
  Duree       := THEdit(Getcontrol('BAM_DUREE'));
  //Cadencement := THValComboBox(GetControl('BAM_CADENCEMENT'));
  //
  StAffaire       := THLabel(GetControl('StAFFAIRE'));
  StTiers         := THLabel(GetControl('StTiers'));
  StRessource     := THLabel(GetControl('StRESSOURCE'));
  //
  TBAM_Affaire    := THLabel(GetControl('TBAM_AFFAIRE'));
  TBAM_Tiers      := THLabel(GetControl('TBAM_TIERS'));
  TBAM_Ressource  := THLabel(GetControl('TBAM_RESSOURCE'));
  TBAM_Duree      := THLabel(GetControl('TBAM_DUREE'));
  //
  AffectChantier  := THCheckBox(Getcontrol('BAM_AFFECTCHANTIER'));
  AffectRessource := THCheckBox(Getcontrol('BAM_AFFECTRESS'));
  AffectTiers     := THCheckBox(Getcontrol('BAM_AFFECTIERS'));
  //
  GerePlanning    := THCheckBox(Getcontrol('BTA_ASSOSRES'));
  Defaut          := THCheckBox(Getcontrol('BTA_DEFAUT'));
  GereConso       := THCheckBox(Getcontrol('BAM_GESTIONCONSO'));
  SaisieValorisee := THCheckBox(Getcontrol('BAM_VALORISE'));

  if TypeAction = 'PCA' then
  Begin
    CumulWorkingHour:= TCheckBox(Getcontrol('BTA_ASSOSDOS1'));
    Modifiable      := TCheckBox(Getcontrol('BTA_MODIFIABLE1'));
    Surbook         := TCheckBox(Getcontrol('BTA_OBLIGATOIRE1'));
  end
  else
  begin
    CumulWorkingHour:= TCheckBox(Getcontrol('BTA_ASSOSDOS'));
    Modifiable      := TCheckBox(Getcontrol('BTA_MODIFIABLE'));
    Surbook         := TCheckBox(Getcontrol('BTA_OBLIGATOIRE'));
  end;
  //
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
  BFond 		    := TToolbarButton97(GetCONTROL('B_FOND'));
  BCouleur 	    := TToolbarButton97(GetCONTROL('B_COULEUR'));
  BFonte		    := TToolbarButton97(GetCONTROL('BFONTE'));
  BImgIcone     := TToolbarButton97(GetCONTROL('B_ICONE'));
  //
  NumeroIcone   := THEdit(GetControl('BTA_NUMEROICONE'));
  ImageIcone    := Timage(GetCONTROL('IMAGEICONE'));
  //
  LCouleur 	    := THEdit(GetControl('LCOULEUR'));
  LFond 		    := THEdit(GetControl('LFOND'));
  //
  CouleurFond   := THEdit(GetControl('BTA_COULEURFOND'));
  Couleur       := THEdit(GetControl('BTA_COULEUR'));
  Fonte         := THEdit(GetControl('BTA_FONTE'));
  //
	LFonte 		    := THLabel(ecran.FindComponent('LFONTE'));

end;

procedure TOF_BTTYPEACTIONMAT.SetScreenEvents;
begin

  AffectChantier.OnClick  := AffCheckOnClick;
  AffectTiers.OnClick     := CliCheckOnClick;
  AffectRessource.OnClick := ResCheckOnClick;
  //
  Affaire1.Onexit         := AffaireOnExit;
  Affaire2.OnExit         := AffaireOnExit;
  Affaire3.OnExit         := AffaireOnExit;
  Tiers.Onexit            := TiersOnExit;
  Ressource.Onexit        := RessourceOnExit;
  Libelle.OnExit          := LibelleOnExit;
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
  BCouleur.onclick        := BCouleur_OnClick;
  BFond.onclick           := BFond_OnClick;
  BFonte.onclick          := BFonte_OnClick;
  BImgIcone.onclick       := BImgIcone_OnClick;

end;

Procedure TOF_BTTYPEACTIONMAT.AffaireOnExit(Sender : TObject);
var IP      : Integer;
    CodeAff : String;
begin

  Codeaff := DechargeCleAffaire(Affaire0, Affaire1, Affaire2, Affaire3, Avenant, Tiers.text, Action, False, True, false, IP);

  if Codeaff <> '' then
  begin
    Affaire.text      := CodeAff;
    StAffaire.caption := ChargeInfoAffaire(Affaire.Text, True);
  end;

  CtrlAffaireTiers;

end;

Procedure TOF_BTTYPEACTIONMAT.TiersOnExit(Sender : TObject);
begin

  if Tiers.text <> '' then  StTiers.Caption := ChargeInfoTiers(Tiers.Text, True);
  
  CtrlTierAffaire;

end;

Procedure TOF_BTTYPEACTIONMAT.LibelleOnExit(Sender : TObject);
begin

  if Libelle.text <> '' then Abrege.Text := Copy(Libelle.Text,1, 25);

end;

Procedure TOF_BTTYPEACTIONMAT.RessourceOnExit(Sender : TObject);
begin

  if Ressource.text <> '' then  StRessource.Caption := ChargeInfoRessource(Ressource.Text,'AND  ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")', True);

end;

Procedure TOF_BTTYPEACTIONMAT.AffCheckOnClick(Sender : TObject);
begin

  Affaire1.Visible      := AffectChantier.Checked;
  Affaire2.Visible      := AffectChantier.Checked;
  Affaire3.Visible      := AffectChantier.Checked;

  if AffectChantier.Checked then
  begin
    ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.Text, False);
    AffectTiers.Visible := False;
    BTEfface1.Visible   := False;
    BTSelect1.Visible   := False;
    Tiers.Visible       := False;
  end
  else
  begin
    AffEffaceOnClick(Self);
    AffectTiers.Visible := True;
    AffectTiers.Visible := True;
    BTEfface1.Visible   := True;
    BTSelect1.Visible   := True;
    Tiers.Visible       := True;
  end;
  //
  BTEfface.Visible      := AffectChantier.Checked;
  BtSelect.visible      := AffectChantier.Checked;
  StAffaire.Visible     := AffectChantier.Checked;
  TBAM_Affaire.Visible  := AffectChantier.Checked;

end;

Procedure TOF_BTTYPEACTIONMAT.CliCheckOnClick(Sender : TObject);
begin

  Tiers.Visible         := AffectTiers.Checked;
  BtEfface1.Visible     := Affecttiers.checked;
  BTSelect1.Visible     := Affecttiers.Checked;
  StTiers.Visible       := Affecttiers.Checked;
  TBAM_Tiers.Visible    := Affecttiers.Checked;

  if not AffectTiers.Checked then TiersEffaceOnClick(Self);

end;

Procedure TOF_BTTYPEACTIONMAT.ResCheckOnClick(Sender : TObject);
begin

  Ressource.Visible     := AffectRessource.checked;
  BtEfface2.visible     := AffectRessource.checked;
  BtSelect2.Visible     := AffectRessource.checked;
  StRessource.Visible   := AffectRessource.Checked;
  TBAM_Ressource.Visible:= AffectRessource.Checked;

  if not AffectRessource.Checked then RessourceEffaceOnClick(Self);

end;

Procedure TOF_BTTYPEACTIONMAT.AffSelectOnClick(Sender : TObject);
Var StChamps  : String;
begin

  StChamps  := Affaire.Text;

  if GetAffaireEnteteSt(Affaire0, Affaire1, Affaire2, Affaire3, AVENANT, nil, StChamps, false, false, false, True, true,'') then Affaire.text := StChamps;

  ChargeCleAffaire (Affaire0,Affaire1,Affaire2,Affaire3,Avenant, BtSelect, Action, Affaire.Text, False);

  StAffaire.caption := ChargeInfoAffaire(Affaire.text, True);
  CtrlAffaireTiers;

end;

Procedure TOF_BTTYPEACTIONMAT.TiersSelectOnClick(Sender : TObject);
Var StChamps  : string;
Begin

  StChamps  := Tiers.Text;

  DispatchRecherche(Tiers, 2, 'T_NATUREAUXI="CLI"','T_TIERS=' + Trim(Tiers.Text), '');

  StTiers.Caption := ChargeInfoTiers(tiers.Text, True);
  
  CtrlTierAffaire;

end;
Procedure TOF_BTTYPEACTIONMAT.RessourceSelectOnClick(Sender : TObject);
Var StChamps  : string;
begin

  StChamps  := ressource.Text;

  DispatchRecherche(Ressource, 3, 'ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")','ARS_RESSOURCE=' + Trim(Ressource.Text), '');

  StRessource.Caption := ChargeInfoRessource(Ressource.Text, 'AND ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")', True);

end;
//
Procedure TOF_BTTYPEACTIONMAT.AffEffaceOnClick(Sender : TObject);
begin

  Affaire.Text  := '';
  Affaire0.Text := '';
  Affaire1.Text := '';
  Affaire2.Text := '';
  Affaire3.Text := '';
  Avenant.Text  := '';
  //
  stAffaire.caption := '...';

end;
Procedure TOF_BTTYPEACTIONMAT.TiersEffaceOnClick(Sender : TObject);
begin

  Tiers.Text  := '';
  StTiers.Caption := '...';

end;
Procedure TOF_BTTYPEACTIONMAT.RessourceEffaceOnClick(Sender : TObject);
begin

  Ressource.Text  := '';
  StRessource.Caption := '...';

end;

Procedure TOF_BTTYPEACTIONMAT.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if Champ='TYPEACTION' then
  begin
    TypeAction := Valeur;
    StWhere := ' AND BTA_TYPEACTION="' + TypeAction + '"';
  end;

  if Champ='BTETAT'     then BTETAT.text := Valeur;

end;

Procedure TOF_BTTYPEACTIONMAT.CreateTOB;
begin

  TobTypeAction := Tob.Create('BTETAT', nil, -1);

end;

procedure TOF_BTTYPEACTIONMAT.DestroyTOB;
begin

  FreeAndNil(TobTypeAction);

end;

Procedure TOF_BTTYPEACTIONMAT.ChargeZoneEcran;
var sttmp 			: String;
		Temp 				: Double;
    DureeMini		: TDateTime;
begin

  BtEtat.Text      := TobTypeAction.GetString('BTA_BTETAT');
  Libelle.Text     := TobTypeAction.GetString('BTA_LIBELLE');
  Abrege.Text      := TobTypeAction.GetString('BTA_ABREGE');
  //
  Affaire.Text     := TobTypeAction.GetString('BTA_AFFAIRE');

  Tiers.Text       := TobTypeAction.GetString('BTA_TIERS');

  Ressource.Text   := TobTypeAction.GetString('BTA_RESSOURCE');

  //Chargement de la durée mini
  Temp        := TobTypeAction.GetValue('BTA_DUREEMINI');
  //DureeMini   := FloatToTime(Temp, true);
  //Duree.text  := TimeToStr(DureeMini);
  Duree.text  := FloatToStr(Temp);

  //Cadencement.Value:= TobTypeAction.GetString('BTA_CADENCEMENT');
  //
  if TobTypeAction.GetString('BTA_ASSOSRES')='X' then
    GerePlanning.Checked := True
  else
    GerePlanning.Checked := False;
  //
  if TobTypeAction.GetString('BTA_DEFAUT')='X' then
    Defaut.Checked := True
  else
    Defaut.Checked := False;
  //
  if TobTypeAction.GetString('BTA_VALORISE')='X' then
    SaisieValorisee.checked := True
  else
    SaisieValorisee.checked := False;
  //
  if TobTypeAction.GetString('BTA_GESTIONCONSO')='X' Then
    GereConso.checked := True
  else
    GereConso.checked := False;
  //
  if TobTypeAction.GetString('BTA_AFFECTCHANTIER')='X' Then
    AffectChantier.checked := True
  else
    AffectChantier.checked := False;
  //
  if TobTypeAction.GetString('BTA_AFFECTRESS')='X' Then
    AffectRessource.checked := True
  else
    AffectRessource.checked := False;
  //
  if TobTypeAction.GetString('BTA_AFFECTIERS')='X' Then
    AffectTiers.checked := True
  else
    AffectTiers.checked := False;
  //
  if TobTypeAction.GetString('BTA_OBLIGATOIRE')='X' Then
    Surbook.checked := True
  else
    Surbook.checked := False;

  if Affaire.text <> ''   then StAffaire.caption    := ChargeInfoAffaire(Affaire.Text, False);

  if Tiers.text <> ''     then StTiers.Caption      := ChargeInfoTiers(tiers.text, False);

  if Ressource.text <> '' then
  begin
    StRessource.Caption  := ChargeInfoRessource(Ressource.Text, 'AND  ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")', False);
  end;

  // Chargement de l'icône associé
  NumeroIcone.Text := TobTypeAction.GetString('BTA_NUMEROICONE');
  If TobTypeAction.GetValue('BTA_NUMEROICONE') <> 9999 then
     ChargeIcone(StrToInt(NumeroIcone.text), ImageIcone)
  else
     ImageIcone.Picture:= Nil;

  Couleurfond.text  := TobTypeAction.GetValue('BTA_COULEURFOND');
  Couleur.text      := TobTypeAction.GetValue('BTA_COULEUR');

  // Chargement Couleurs & Fonte
  AfficheCouleur(LFOND, CouleurFond);
  AfficheCouleur(LCOULEUR, Couleur);

  LFONTE.Font.Name := TobTypeAction.GetString('BTA_FONTE');
  LFONTE.Font.Color := LCOULEUR.Color;
  StTmp := TobTypeAction.GetString('BTA_FONTESTYLE');

  if StTmp <> '' Then
  begin
    If StTmp = 'fsBold' Then
      LFONTE.Font.Style := [fsBold]
    else if StTmp = 'fsItalic' Then
      LFONTE.Font.Style := [fsItalic]
    else if StTmp = 'fsUnderline' Then
      LFONTE.Font.Style := [fsUnderline]
    else if StTmp = 'fsStrikeout' Then
      LFONTE.Font.Style := [fsStrikeout];
  End;

  StTmp := TobTypeAction.GetValue('BTA_FONTESIZE');
  if (StrToInt (StTmp) >= 1) Then LFONTE.Font.Size := StrToInt(StTmp);

end;

Procedure TOF_BTTYPEACTIONMAT.ChargeZoneTOB;
Var Temp 			: Double;
    DureeMini	: TDateTime;
    Color     : Integer;
begin

  TobTypeAction.PutValue('BTA_TYPEACTION', TypeAction);
  //
  TobTypeAction.PutValue('BTA_BTETAT',     BtEtat.Text);
  TobTypeAction.PutValue('BTA_LIBELLE',    Libelle.Text);
  TobTypeAction.PutValue('BTA_ABREGE',     Abrege.Text);
  //
  if (Affaire1.text = '') And (Affaire2.text = '') and (Affaire3.text='') then
    TobTypeAction.PutValue('BTA_AFFAIRE',    '')
  else
    TobTypeAction.PutValue('BTA_AFFAIRE',    Affaire.Text);

  TobTypeAction.PutValue('BTA_TIERS',      Tiers.Text);
  TobTypeAction.PutValue('BTA_RESSOURCE',  Ressource.Text);

  //Mise à jour de la durée mini
  //DureeMini := StrToTime(Duree.Text);
  //emp      := TimeToFloat(DureeMini, true);
  TobTypeAction.PutValue('BTA_DUREEMINI', Duree.text);

  //TobTypeAction.PutValue('BTA_CADENCEMENT',Cadencement.Value);
  //
  if GerePlanning.Checked then
    TobTypeAction.PutValue('BTA_ASSOSRES','X')
  else
    TobTypeAction.PutValue('BTA_ASSOSRES','-');
  //
  if Defaut.Checked then
    TobTypeAction.PutValue('BTA_DEFAUT','X')
  else
    TobTypeAction.PutValue('BTA_DEFAUT','-');
  //
  if SaisieValorisee.checked then
    TobTypeAction.PutValue('BTA_VALORISE','X')
  else
    TobTypeAction.PutValue('BTA_VALORISE','-');
  //
  if GereConso.checked Then
    TobTypeAction.PutValue('BTA_GESTIONCONSO','X')
  else
    TobTypeAction.PutValue('BTA_GESTIONCONSO','-');
  //
  if AffectChantier.checked Then
    TobTypeAction.PutValue('BTA_AFFECTCHANTIER','X')
  else
    TobTypeAction.PutValue('BTA_AFFECTCHANTIER','-');
  //
  if AffectRessource.checked Then
    TobTypeAction.PutValue('BTA_AFFECTRESS','X')
  else
    TobTypeAction.PutValue('BTA_AFFECTRESS','-');
  //
  if Surbook.checked Then
    TobTypeAction.PutValue('BTA_OBLIGATOIRE','X')
  else
    TobTypeAction.PutValue('BTA_OBLIGATOIRE','-');
  //
  if AffectTiers.checked  Then
    TobTypeAction.PutValue('BTA_AFFECTIERS','X')
  else
    TobTypeAction.PutValue('BTA_AFFECTIERS','-');
  //
  TobTypeAction.PutValue('BTA_NUMEROICONE', NumeroIcone.text);
  //
  Color := Lfond.Color;
  TobTypeAction.PutValue('BTA_COULEURFOND', IntToStr(Color));
  Color := LCOULEUR.Color;
  TobTypeAction.PutValue('BTA_COULEUR', IntToStr(Color));

  TobTypeAction.PutValue('BTA_FONTE',LFONTE.Font.Name);
  TobTypeAction.PutValue('BTA_FONTESIZE',LFONTE.Font.Size);

  If fsBold in LFONTE.Font.Style Then
     TobTypeAction.PutValue('BTA_FONTESTYLE','fsBold')
  else if fsItalic in LFONTE.Font.Style Then
     TobTypeAction.PutValue('BTA_FONTESTYLE','fsItalic')
  else if fsUnderline in LFONTE.Font.Style Then
     TobTypeAction.PutValue('BTA_FONTESTYLE','fsUnderline')
  else if fsStrikeout in LFONTE.Font.Style Then
     TobTypeAction.PutValue('BTA_FONTESTYLE','fsStrikeout')
  Else
     TobTypeAction.PutValue('BTA_FONTESTYLE','');

end;

procedure TOF_BTTYPEACTIONMAT.CtrlAffaireTiers;
begin

  if not AffectTiers.Checked then Exit;

  if Tiers.Text = '' then Exit;

  StSQL := 'SELECT AFF_LIBELLE FROM AFFAIRE LEFT JOIN TIERS ON T_TIERS=AFF_TIERS WHERE AFF_AFFAIRE = "'+ Affaire.Text + '" AND AFF_TIERS = "' + Tiers.Text + '" AND T_NATUREAUXI="CLI"';
  QQ    := OpenSQL(StSQL, False);

  if QQ.Eof then
  begin
    PGIError(TexteMessage[7], 'Type Action');
    Affaire.Text      := '';
    Affaire0.Text     := '';
    Affaire1.Text     := '';
    Affaire2.Text     := '';
    Affaire3.Text     := '';
    Avenant.Text      := '';
    StAffaire.Caption := '';
    Affaire1.Setfocus;
  end
  else
    StAffaire.caption := QQ.findfield('AFF_LIBELLE').AsString;

  ferme(QQ);

end;


procedure TOF_BTTYPEACTIONMAT.CtrlTierAffaire;
begin

  if Tiers.Text = '' then Exit;

  //contrôle si le tiers appartient à l'affaire si affectation affaire coché
  if not AffectChantier.Checked then Exit;

  StSQL := 'SELECT T_LIBELLE FROM TIERS LEFT JOIN AFFAIRE ON AFF_TIERS=T_TIERS WHERE AFF_AFFAIRE = "'+ Affaire.Text + '" AND AFF_TIERS = "' + Tiers.Text + '" AND T_NATUREAUXI="CLI"';

  QQ    := OpenSQL(StSQL, False);

  if QQ.Eof then
  begin
    PGIError(TexteMessage[6], 'Type Action');
    Tiers.Text      := '';
    StTiers.Caption := '';
    Tiers.SetFocus;
  end
  else
    StTiers.caption := QQ.findfield('T_LIBELLE').AsString;

  ferme(QQ);

end;

procedure TOF_BTTYPEACTIONMAT.RazZoneEcran;
begin

  BtEtat.Text     := '';
  Libelle.Text    := '';
  Abrege.Text     := '';
  //
  Affaire.Text    := '';
  Affaire0.Text   := '';
  Affaire1.Text   := '';
  Affaire2.Text   := '';
  Affaire3.Text   := '';
  Avenant.Text    := '';
  //
  Tiers.Text      := '';
  Ressource.Text  := '';
  //
  StAffaire.Caption   := '...';
  StTiers.Caption     := '...';
  StRessource.Caption := '...';
  //
  Duree.Text := '01:00';
  //Cadencement.Value   := '0';

  // Init. couleurs ,fontes,icone
  NumeroIcone.text    := '9999';
  COULEURFOND.Text    := Inttostr(16777215);
  COULEUR.text        := IntToStr(0);

  AfficheCouleur(LFOND, COULEURFOND);
  AfficheCouleur(LCOULEUR, COULEUR);

  Fonte.text          := 'Ms Sans Serif';
  //
  LFonte.Font.Name    := 'Ms Sans Serif';
  LFonte.Font.Size    := 8;
  LFonte.Font.Style   := [fsBold];
  LFonte.Font.Color   := LCouleur.Color;
  //
  AffectChantier.checked  := False;
  AffectTiers.checked     := False;
  AffectRessource.checked := False;
  GerePlanning.checked    := False;
  Defaut.Checked          := False;
  GereConso.checked       := False;
  SaisieValorisee.checked := False;
  CumulWorkingHour.Checked:= False;
  Modifiable.Checked      := False;
  Surbook.Checked         := False;
  //
  BTEfface.Visible  := False;
  BTEfface1.Visible := False;
  BTEfface2.Visible := False;
  //
  BTSelect.Visible  := False;
  BTSelect1.Visible := False;
  BTSelect2.Visible := False;
  //
  StAffaire.visible    := False;
  StTiers.Visible     := False;
  StRessource.Visible := False;
  //
  Affaire1.Visible := False;
  Affaire2.Visible := False;
  Affaire3.Visible := False;
  //
  BtDelete.visible := False;
end;

procedure TOF_BTTYPEACTIONMAT.BFOND_OnClick(Sender: TObject);
begin

  SelColorNew(LFOND,Couleurfond, TForm(Ecran));

end;

procedure TOF_BTTYPEACTIONMAT.BCOULEUR_OnClick(Sender: TObject);
begin

  SelColorNew(LCOULEUR, COULEUR, TForm(Ecran));

  //Affichage de la police dans la nouvelle couleur
  LFONTE.Font.Color := LCOULEUR.Color;

end;

procedure TOF_BTTYPEACTIONMAT.BFONTE_OnClick(Sender: TObject);
begin

	SelFonteNew(LFONTE, Fonte, TForm(Ecran));

  //Affichage de la couleur en fonction de la police
  LCOULEUR.Color := LFONTE.Font.Color;

end;

procedure TOF_BTTYPEACTIONMAT.BImgIcone_OnClick(Sender: TObject);
var NumIcone : Variant;
begin

  ImageIcone.Picture := nil;

  NumeroIcone.Text := IntToStr(ListeIcone(True));
  NumIcone := StrToInt(NumeroIcone.Text);

  If (NumIcone <> -1) or (NumIcone=9999) then
  begin
    ChargeIcone(NumIcone,ImageIcone);
  end;

end;


Initialization
  registerclasses ( [ TOF_BTTYPEACTIONMAT ] ) ;
end.

