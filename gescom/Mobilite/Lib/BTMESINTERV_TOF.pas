{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/10/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMESINTERV ()
Mots clefs ... : TOF;BTMESINTERV
*****************************************************************}
Unit BTMESINTERV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     //
     ShellAPI,
     ImgList,
     Graphics,
{$IFDEF EAGLCLIENT}
      eMul,
      Maineagl,
      EFiche,
      UtileAGL,
      HQry,
{$ELSE}
      db,
      Grids,
      {$IFNDEF DBXPRESS}
      dbTables,
      {$ELSE}
      uDbxDataSet,
      {$ENDIF}
      DBGrids,
      mul,
      FE_Main,
      Fiche,
{$ENDIF}
      Hpanel,
      HTB97,
      HRichOle,
      Types,
      windows,
      HDB,
      ParamSoc,
      ParamDBG,
      UTOM,
      MsgUtil,
      HSysMenu,
      AglInit,
      Vierge,
      utob;


Type
  TOF_BTMESINTERV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private

    CodeUser      : String;
    UserName      : String;
    Domaine       : String;
    ListeParam    : String;


    GestGrille    : AffGrille;

    ValoModePrestation    : Boolean;

    P0            : String;
    P1            : String;
    P2            : String;
    P3            : String;
    Avn           : String;

    CodeAppel     : String;
    Avenant       : String;

    CodeTiers     : String;
    Auxiliaire    : String;
    CodeContrat   : String;
    CodeContact   : String;

    PageControl   : TPageControl;
    
    TT_Client     : THLabel;
    TT_Civilite   : THLabel;
    TT_Libelle    : THLabel;
    TT_Prenom     : THLabel;
    TT_Adresse1   : THLabel;
    TT_Adresse2   : THLabel;
    TT_Adresse3   : THLabel;
    TT_CodePostal : THLabel;
    TT_Ville      : THLabel;
    TT_Pays       : THLabel;
    TT_Region     : THLabel;
    //
    TT_Contact    : THLabel;
    TT_TelDom     : THLabel;
    TT_TelBur     : THLabel;
    TT_TelMob     : THLabel;
    TT_Email      : THLabel;
    TT_Contrat    : THLabel;
    TT_TypeInter  : THLabel;
    TT_DateDebut  : THLabel;
    TT_DateFin    : THLabel;
    TT_DateSouhait: THLabel;

    TT_DESCRIPTIF : THRichEditOLE;
    TT_DESCRIPTIF1: THRichEditOLE;

    PBarreBtn     : THPanel;
    PAppel        : THPanel;

    PBouton       : TToolWindow97;

    Dock971       : TDock97;

    BSuivant      : TToolBarButton97;
    BPrecedent    : TToolBarButton97;
    BParametre    : TToolBarButton97;
    BAjout        : TToolBarButton97;
    BQuitter      : TToolBarButton97;
    BValide       : TToolBarButton97;
    BTBasse       : TToolBarButton97;
    BtHaute       : TToolBarButton97;
    BtNormale     : TToolBarButton97;

    TobAppel      : TOB;
    TobRessource  : TOB;
    TobConso      : TOB;
    TobOle        : TOB;
    TobTiers      : TOB;
    TobContact    : TOB;
    TobContrat    : TOB;
    TobAdrInt     : TOB;

    GrilleConso   : THGrid;
    HmTrad        : ThSystemMenu;

    TWindows      : TToolWindow97;
    BDefinitif    : TToolbarButton97;
    BProvisoire   : TToolbarButton97;

    function  AddLigneConso: TOB;
    procedure AfficheAdrInt(TobAdrInt: TOB);
    procedure AfficheContact(TobContact: TOB);
    procedure AfficheContrat(TobContrat: tob);
    procedure AfficheTiers(TobTiers: Tob);
    procedure BAjout_OnClick(Sender: TObject);
    procedure BDefinitif_Click(Sender: TObject);
    procedure BDelete_Click;
    procedure BParametre_OnClick(Sender: TObject);
    procedure BPrecedent_OnClick(Sender: TObject);
    procedure BProvisoire_Click(Sender: TObject);
    procedure BSuivant_OnClick(Sender: TObject);
    procedure BValide_Click(Sender: TObject);
    procedure ChargeEcran(TOBL : TOB);
    procedure ChargeTOBAppel;
    procedure ControleChamp(Champ, Valeur: String);
    procedure CreateTOB;
    procedure CreateTToolwindows;
    procedure DesTroyTob;
    procedure GestionSaisieNumAppel;
    procedure GetObjects;
    procedure GrilleConso_OnRowClick(Sender: TObject);
    procedure InitCreation(TOBCO, TOBEchange: TOB);
    procedure InitEcran;
    procedure InitGrille;
    procedure InitNewConso(TOBCO, TOBEchange: TOB);
    procedure LectureAdrInt;
    procedure LectureConso;
    procedure LectureContact;
    procedure LectureContrat;
    procedure LectureLienOle;
    procedure LectureRessource;
    procedure LectureTiers;
    procedure MAJConso(TOBCO, TOBEchange: TOB);
    procedure MajDescriptifRealisation;
    procedure GrilleConso_PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure RefreshGrille;
    procedure SetScreenEvents(Etat : Boolean);
    procedure SetImportance;
    function  SetUniqueNumber(TOBCO: TOB): boolean;
    procedure ValorisationConsommation(TOBCO, TobEchange: TOB);
//
		procedure AfterFormShow;
  end ;

Implementation

Uses  AppelsUtil,
      UtilPGI,
      BTPUtil,
      UtilScreen,
      UtilSaisieConso,
      FactDomaines,
      CalcOLEGenericBTP,
      BTSIGNATURE_TOF;


procedure TOF_BTMESINTERV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMESINTERV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMESINTERV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMESINTERV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMESINTERV.OnArgument (S : String ) ;
var X       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
    EtatApp : String;
begin
  Inherited ;

  // ---
  CodeUser  := V_PGI.User;
  UserName  := V_PGI.UserName;


  HMTrad    := ThsystemMenu.Create(Ecran);

  CreateTOB;

  //traitement Arguments
	Critere :=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
    begin
      if Critere <> '' then
      begin
        X := pos (':', Critere) ;
        if x = 0 then
          X := pos ('=', Critere) ;
        if x <> 0 then
        begin
          Champ := copy (Critere, 1, X - 1) ;
          Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        	ControleChamp(champ, valeur);
				end
      end;
      Critere := (Trim(ReadTokenSt(S)));
    end;

  //Chargement des objets écran dans l'uses
  GetObjects;

  //remise à zéro des zone écran
  InitEcran;

  //lecture de la table des ressources
  LectureRessource;

  //Chargement des données de la table (LISTE/Vue)
  ChargeTOBAppel;

  //Chargement des données de la table (LISTE/Vue)
  ChargeEcran(TobAppel);

  if TobAppel.GetString('AFF_ETATAFFAIRE') = 'AFF' then
    EtatApp := 'Affecté'
  else
    EtatApp := 'En-cours de Réalisation';


  //Titre de la fiche
  if Assigned(PBarreBtn)  then PBarreBtn.Caption := 'Interventions pour ' + UserName;

  If Assigned(PAppel)     then
  begin
    if Avenant <> '00' then
      Pappel.Caption := 'APPEL N° ' + P1 + P2 + P3 + ' - (' + Avn + ')'
    else
      Pappel.Caption := 'APPEL N° ' + P1 + P2 + P3;
    Pappel.Caption := Pappel.Caption + '    '+ EtatApp;
  end;

  //Définition des Evènement des zones écrans
  SetScreenEvents(True);
	TFVierge(Ecran).OnAfterFormShow := AfterFormShow;
end ;

procedure TOF_BTMESINTERV.OnClose ;
begin

  FreeAndNil(HMTrad);

  DestroyTob;

  Inherited ;

end ;

procedure TOF_BTMESINTERV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMESINTERV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMESINTERV.ControleChamp(Champ, Valeur: String);
begin
  if Champ = 'APPEL' then CodeAppel := Valeur;
end;

Procedure TOF_BTMESINTERV.GetObjects;
begin

  if assigned(GetControl('PBouton'))        then PBouton      := TToolWindow97(GetControl('PBOUTON'));
  if Assigned(GetControl('Dock971'))        then Dock971      := TDock97(Getcontrol('Dock971'));
  //
  if Assigned(Getcontrol('PAGECONTROL'))    Then PAGECONTROL  := TPageControl(Getcontrol('PAGECONTROL'));
  //
  if Assigned(Getcontrol('PBARREBTN'))      then PBarreBtn    := THPanel(GetControl('PBARREBTN'));
  if Assigned(Getcontrol('PAPPEL'))         then PAppel       := THPanel(GetControl('PAPPEL'));
  //
  if Assigned(GetControl('BSUIVANT'))       then BSuivant     := TToolBarButton97(GetControl('BSUIVANT'));
  if Assigned(GetControl('BPRECEDENT'))     then BPrecedent   := TToolBarButton97(GetControl('BPRECEDENT'));
  if Assigned(GetControl('BPARAMETRE'))     then BParametre   := TToolBarButton97(GetControl('BPARAMETRE'));
  if Assigned(GetControl('BAJOUT'))         then BAjout       := TToolBarButton97(GetControl('BAJOUT'));
  if Assigned(GetControl('BVALIDER_'))      then BValide      := TToolBarButton97(GetControl('BVALIDER_'));

  if Assigned(GetControl('BQUITTER'))       then BQuitter     := TToolBarButton97(GetControl('BQUITTER'));
  //
  if Assigned(GetControl('BTNORMALE'))      then BtNormale    := TToolBarButton97(GetControl('BTNORMALE'));
  if Assigned(GetControl('BTBASSE'))        then BtBasse      := TToolBarButton97(GetControl('BTBASSE'));
  if Assigned(GetControl('BTHAUTE'))        then BtHaute      := TToolBarButton97(GetControl('BTHAUTE'));
  //
  if Assigned(Getcontrol('TT_CLIENT'))      then TT_Client    := THLabel(GetControl('TT_CLIENT'));
  if Assigned(Getcontrol('TT_CIVILITE'))    then TT_Civilite  := THLabel(GetControl('TT_CIVILITE'));
  if Assigned(Getcontrol('TT_LIBELLE'))     then TT_Libelle   := THLabel(GetControl('TT_LIBELLE'));
  if Assigned(Getcontrol('TT_PRENOM'))      then TT_Prenom    := THLabel(GetControl('TT_PRENOM'));
  if Assigned(Getcontrol('TT_ADRESSE1'))    then TT_Adresse1  := THLabel(GetControl('TT_ADRESSE1'));
  if Assigned(Getcontrol('TT_ADRESSE2'))    then TT_Adresse2  := THLabel(GetControl('TT_ADRESSE2'));
  if Assigned(Getcontrol('TT_ADRESSE3'))    then TT_Adresse3  := THLabel(GetControl('TT_ADRESSE3'));
  if Assigned(Getcontrol('TT_CODEPOSTAL'))  then TT_CodePostal:= THLabel(GetControl('TT_CODEPOSTAL'));
  if Assigned(Getcontrol('TT_VILLE'))       then TT_Ville     := THLabel(GetControl('TT_VILLE'));
  if Assigned(Getcontrol('TT_PAYS'))        then TT_Pays      := THLabel(GetControl('TT_PAYS'));
  if Assigned(Getcontrol('TT_REGION'))      then TT_Region    := THLabel(GetControl('TT_REGION'));
  //
  if Assigned(Getcontrol('TT_CONTACT'))     then TT_Contact   := THLabel(GetControl('TT_CONTACT'));
  if Assigned(Getcontrol('TT_TELDOM'))      then TT_TelDom    := THLabel(GetControl('TT_TELDOM'));
  if Assigned(Getcontrol('TT_TELBUR'))      then TT_TelBur    := THLabel(GetControl('TT_TELBUR'));
  if Assigned(Getcontrol('TT_TELMOB'))      then TT_TelMob    := THLabel(GetControl('TT_TELMOB'));
  if Assigned(Getcontrol('TT_EMAILCLIENT')) then TT_Email     := THLabel(GetControl('TT_EMAILCLIENT'));
  //
  if Assigned(Getcontrol('TT_CONTRAT'))     then TT_Contrat     := THLabel(GetControl('TT_CONTRAT'));
  if Assigned(Getcontrol('TT_TYPEINTER'))   then TT_TypeInter   := THLabel(GetControl('TT_TYPEINTER'));

  if Assigned(Getcontrol('TT_DATEDEB'))     then TT_DateDebut   := THLabel(GetControl('TT_DATEDEB'));
  if Assigned(Getcontrol('TT_DATEFIN'))     then TT_DateFin     := THLabel(GetControl('TT_DATEFIN'));
  if Assigned(Getcontrol('TT_DATESOUHAIT')) then TT_DateSouhait := THLabel(GetControl('TT_DATESOUHAIT'));
  //
  If assigned(GetControl('TT_DESCRIPTIF'))  then TT_Descriptif  := THRichEditOLE(Getcontrol('TT_Descriptif'));
  If assigned(GetControl('TT_DESCRIPTIF1')) then TT_Descriptif1 := THRichEditOLE(Getcontrol('TT_Descriptif1'));
  //
  If Assigned(GetControl('GRILLE_CONSO'))   then GrilleConso    := THGrid(GetControl('GRILLE_CONSO'));
  //
  CreateTToolwindows;

end;

procedure TOF_BTMESINTERV.SetScreenEvents(Etat : Boolean);
begin

  if etat then
  begin
    if Assigned(BSuivant)       then BSuivant.OnClick     := BSuivant_OnClick;
    if Assigned(BPrecedent)     then BPrecedent.OnClick   := BPrecedent_OnClick;
    if Assigned(BParametre)     then BParametre.OnClick   := BParametre_OnClick;
    if Assigned(BAjout)         then BAjout.OnClick       := BAjout_OnClick;
    if Assigned(BValide)        then BValide.OnClick      := BValide_Click;
    if Assigned(BProvisoire)    then BProvisoire.OnClick  := BProvisoire_Click;
    if assigned(BDefinitif)     then Bdefinitif.OnClick   := Bdefinitif_Click;
    if assigned(GrilleConso)    then
    begin
      GrilleConso.OnClick       := GrilleConso_OnRowClick;
      GrilleConso.PostDrawCell  := GrilleConso_PostDrawCell;
    end;
  end
  else
  begin
    if Assigned(BSuivant)       then BSuivant.OnClick     := Nil;
    if Assigned(BPrecedent)     then BPrecedent.OnClick   := Nil;
    if Assigned(BParametre)     then BParametre.OnClick   := Nil;
    if Assigned(BAjout)         then BAjout.OnClick       := Nil;
    if Assigned(BValide)        then BValide.OnClick      := Nil;
    if assigned(GrilleConso)    then
    begin
      GrilleConso.OnClick       := Nil;
      GrilleConso.PostDrawCell  := Nil;
    end;
    if Assigned(BProvisoire)    then BProvisoire.OnClick  := Nil;
    if assigned(BDefinitif)     then Bdefinitif.OnClick   := Nil;
  end;

end;

Procedure TOF_BTMESINTERV.InitEcran;
begin

  if Assigned(TT_CLIENT)      then TT_Client.caption    := '';
  if Assigned(TT_CIVILITE)    then TT_Civilite.caption  := '';
  if Assigned(TT_LIBELLE)     then TT_Libelle.caption   := '';
  if Assigned(TT_PRENOM)      then TT_Prenom.caption    := '';
  if Assigned(TT_ADRESSE1)    then TT_Adresse1.caption  := '';
  if Assigned(TT_ADRESSE2)    then TT_Adresse2.caption  := '';
  if Assigned(TT_ADRESSE3)    then TT_Adresse3.caption  := '';
  if Assigned(TT_CODEPOSTAL)  then TT_CodePostal.caption:= '';
  if Assigned(TT_VILLE)       then TT_Ville.caption     := '';
  if Assigned(TT_PAYS)        then TT_Pays.caption      := '';
  if Assigned(TT_REGION)      then TT_Region.caption    := '';
  //
  if Assigned(TT_CONTACT)     then TT_Contact.caption   := '';
  if Assigned(TT_TELDOM)      then TT_TelDom.caption    := '';
  if Assigned(TT_TELBUR)      then TT_TelBur.caption    := '';
  if Assigned(TT_TELMOB)      then TT_TelMob.caption    := '';
  if Assigned(TT_EMAIL)       then TT_Email.caption     := '';
  //
  if Assigned(TT_CONTRAT)     then TT_Contrat.caption   := '';
  if Assigned(TT_TYPEINTER)   then TT_TypeInter.caption := '';

  if Assigned(TT_DATEDEBUT)   then TT_DateDebut.caption   := DatetimeToStr(idate1900);
  if Assigned(TT_DATEFIN)     then TT_DateFin.caption     := DatetimeToStr(idate1900);
  if Assigned(TT_DATESOUHAIT) then TT_DateSouhait.Caption := DatetimeToStr(V_PGI.DateEntree);

  //Le Bouton de Validation ne s'affichera que si on a au moins une ligne de saisie
//  if Assigned(BValide)        then BValide.Visible    := False;

  if Assigned(BPrecedent)     then BPrecedent.Visible := False;
  if Assigned(BAjout)         then BAjout.Visible     := False;

  //Le bouton de paramétrage ne s'affichera que si on est administrateur....
  if Assigned(BParametre)     then BParametre.Visible := False;
  //
  If assigned(TT_DESCRIPTIF)  then
  begin
    TT_Descriptif.text  := '';
    TT_Descriptif.Enabled := False;
  end;

  If assigned(TT_DESCRIPTIF1) then TT_Descriptif1.text  := '';

  if Assigned(PageControl) then
  begin
    PageControl.ActivePageIndex := 0;
    //PageControl.TabIndex := 0;
    //Pagecontrol.pages[0].tabvisible := false;
    //PageControl.Pages[0].Visible := True;
    //
  end;

end;

Procedure TOF_BTMESINTERV.ChargeTOBAppel;
Var StSQL   : String;
    QAppel  : TQuery;
begin

  StSQL := 'SELECT * FROM AFFAIRE ';
  StSQL := StSQL + ' LEFT JOIN LIENSOLE ON AFF_AFFAIRE = LO_IDENTIFIANT';
  StSQL := StSQL + ' WHERE AFF_AFFAIRE0 = "W"';
  StSql := StSQL + ' AND AFF_AFFAIRE = "' + CodeAppel + '" ';

  QAppel := OpenSQL(StSQL, True);

  if QAppel.Eof then
  begin
    AfficheErreur('BTMESINTERV','1', 'Saisie Retour Intervention');
    ecran.ModalResult := mrCancel;
    exit;
  end;

  TobAppel.SelectDb('',QAppel);

  GestionSaisieNumAppel;

  Ferme(QAppel);

end;

Procedure TOF_BTMESINTERV.ChargeEcran(TOBL : TOB);
Begin

  If Assigned(BtBasse)    then BTBasse.visible   := False;
  If Assigned(BtHaute)    then BTHaute.Visible   := False;
  If Assigned(BtNormale)  then BTNormale.Visible := False;

  With TOBL do
  Begin
    Avenant := GetString('AFF_AVENANT');
    //Chargement du domaine D'activité de l'Affaire
    Domaine := GetString('AFF_DOMAINE');
    //
    //Affichage de l'importance de l'appel
    SetImportance;
    //
    //chargement des informations du tiers
    CodeTiers := GetString('AFF_TIERS');
    LectureTiers;
    //
    //Chargement des infos contrat
    CodeContrat := GetString('AFF_AFFAIREINIT');
    LectureContrat;
    //
    //Chargement des infos Contact
    CodeContact := GetString('AFF_NUMEROCONTACT');
    LectureContact;
    //
    //Chargement de l'adresse d'intervention
    LectureAdrInt;

    if Assigned(TT_DateDebut)    Then TT_DateDebut.Caption   := DateToStr(GetDateTime('AFF_DATEDEBUT'));
    if Assigned(TT_Datefin)      Then TT_Datefin.Caption     := DateToStr(GetDateTime('AFF_DATEFIN'));
    if Assigned(TT_DateSouhait)  Then TT_DateSouhait.Caption := DateToStr(GetDateTime('AFF_DATESIGNE'));
    //
    if Assigned(TT_Descriptif)   Then
    begin
    	AppliqueFontDefaut (TT_Descriptif);
      TT_Descriptif.Text     := GetValue('AFF_DESCRIPTIF');
    end;

    LectureLienOle;

    if Assigned(TT_Descriptif1)   Then
    begin
    	AppliqueFontDefaut (TT_Descriptif1);
      TT_Descriptif1.Text    := TobOle.GetValue('LO_OBJET');
    end;

    if Assigned(GrilleConso) then
    begin
      InitGrille;
      //chargement de la page saisie intervention ecran 2 (La saisie proprement dite)
      CodeUser := GetString('AFF_RESPONSABLE');
      //Chargement de la liste Associées
      ChargeListeAssociee(ListeParam, 'Liste des Consommations Par Appel', Nil, GestGrille);
      //Dessin de la grille en fonction des options de la liste
      DessineGrille(Nil, GrilleConso,GestGrille );
      //Chargement des données de la table (LISTE/Vue)
      LectureConso;

      //Affichage des données de la table (LISTE/Vue)
      RefreshGrille;
      //
    end;
  end;

end;

Procedure TOF_BTMESINTERV.LectureTiers;
Var StSQL     : String;
    QTiers    : TQuery;
Begin

  StSQL := '';
  StSQL := 'SELECT * FROM TIERS ';
  StSQL := StSQL + 'WHERE T_NATUREAUXI ="CLI" AND T_TIERS="' + CodeTiers + '"';

  QTiers := OpenSQL(StSQL, True);

  if QTiers.Eof then
  begin
     AfficheErreur('BTMESINTERV','2', 'Saisie Retour Intervention');
     Ferme(QTiers);
     ecran.ModalResult := mrCancel;
    exit;
  end;

  TobTiers.SelectDb('',Qtiers);

  Ferme(QTiers);

  if TobTiers.GetString('T_FERME')='X' then
  begin
    AfficheErreur('BTMESINTERV','1', 'Saisie Retour Intervention');
    TobTiers.free;
    ecran.ModalResult := mrCancel;
    exit;
  end;

	AfficheTiers(TobTiers);
end;

procedure TOF_BTMESINTERV.LectureContrat;
Var StSQL       : String;
    QContrat    : TQuery;
Begin

  if CodeContrat = '' then exit;

  StSQL := '';

  StSQL := 'SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE ="' + CodeContrat + '"';

  if CodeTiers <> '' then
	  StSQL := StSQL + 'AND AFF_TIERS ="' + CodeTiers + '"';

  QContrat := OpenSQL(StSQL, True);

  if QContrat.Eof then
  begin
    AfficheErreur('BTMESINTERV','4', 'Saisie Retour Intervention');
    Ferme(QContrat);
    ecran.ModalResult := mrCancel;
    exit;
  end;

  TobContrat.SelectDb('', QContrat);

  Ferme(QContrat);

  AfficheContrat(TobContrat);


End;

Procedure TOF_BTMESINTERV.LectureContact;
Var StSQL       : String;
    QContact    : TQuery;
Begin

  if CodeContact = '0' then exit;

  StSql := '';
  StSql := 'SELECT * FROM CONTACT WHERE C_AUXILIAIRE ="' + Auxiliaire + '" AND C_NUMEROCONTACT=' + CodeContact;

  QContact := OpenSQL(StSQL, True);

  if QContact.Eof then
  begin
    AfficheErreur('BTMESINTERV','5', 'Saisie Retour Intervention');
    Ferme(QContact);
    ecran.ModalResult := mrCancel;
    exit;
  end;

  TobContact.SelectDB('', QContact);

  Ferme(QContact);

  AfficheContact(TobContact);

end;

Procedure TOF_BTMESINTERV.LectureAdrInt;
Var StSQL       : String;
    QAdrInt     : TQuery;
Begin

  StSQL := 'SELECT * FROM ADRESSES WHERE ADR_REFCODE="' + CodeAppel + '" ';
  StSQL := StSQL + 'AND ADR_TYPEADRESSE="INT"';

  QAdrInt := OpenSQL(StSQL, true);

  if QAdrInt.eof then
  Begin
    ferme(QAdrInt);
    exit;
  end;

  TobAdrInt.selectdb('', QAdrInt);

  Ferme(QAdrInt);

  AfficheAdrInt(TobAdrInt);

end;
Procedure TOF_BTMESINTERV.LectureLienOle;
var StSql : STring;
		QOle  : Tquery;
Begin

  StSql := 'SELECT * FROM LIENSOLE  WHERE LO_TABLEBLOB="APP" AND LO_QUALIFIANTBLOB="MOT" AND LO_EMPLOIBLOB="REA"' ;
  StSql := StSql + ' AND LO_IDENTIFIANT="' + CodeAppel + '" AND LO_RANGBLOB = 1';

  QOle := OpenSQL (StSql,true,-1,'',true);

  if QOle.eof then
  Begin
    ferme(QOle);
    exit;
  end;

  TOBOle.SelectDB ('',QOle);

  Ferme(QOle);

end;

Procedure TOF_BTMESINTERV.LectureRessource;
var StSql      : STring;
		QRessource : Tquery;
begin

  StSql := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+CodeUser+'"';

  QRessource := OpenSQL (StSql,true,-1,'',true);

  if QRessource.eof then
  Begin
    ferme(QRessource);
    exit;
  end;

  TOBRessource.SelectDB ('',QRessource);

  Ferme(QRessource);

end;

//Affichage des informations du tiers
Procedure TOF_BTMESINTERV.AfficheTiers(TobTiers : Tob);
Begin

  With TobTiers do
  Begin
    Auxiliaire          := GetString('T_AUXILIAIRE');
    //
    If Assigned(TT_Client)    then TT_Client.Caption   := CodeTiers;
    If Assigned(TT_Civilite)  then TT_CIVILITE.Caption := GetString('T_JURIDIQUE');
    //
    TT_LIBELLE.caption  := GetString('T_LIBELLE');
    TT_PRENOM.Caption   := GetString('T_PRENOM');
    TT_ADRESSE1.Caption := GetString('T_ADRESSE1');
    TT_ADRESSE2.Caption := GetString('T_ADRESSE2');
    TT_ADRESSE3.caption := GetString('T_ADRESSE3');
    TT_CODEPOSTAL.caption := GetString('T_CODEPOSTAL');
    TT_VILLE.caption    := GetString('T_VILLE');
    TT_PAYS.caption     := GetString('T_PAYS');
    TT_REGION.caption   := GetString('T_REGION');
    //
    if TT_CONTACT.Caption = '' then
    Begin
      TT_TELDOM.Caption := GetString('T_TELEPHONE');
    	TT_TELBUR.Caption := GetString('T_FAX');
    	TT_TELMOB.Caption := GetString('T_TELEX');
      TT_EMAIL.Caption  := GetString('T_EMAIL');
    End;
  end;

end;

//Affichage des informations du contrat
Procedure TOF_BTMESINTERV.AfficheContrat(TobContrat : tob);
Var C0      : String;
    C1      : String;
    C2      : String;
    C3      : String;
    C4      : String;
    TypeContrat : String;
Begin

  With TobContrat do
  begin
    BTPCodeAffaireDecoupe(CodeContrat,C0,C1,C2,C3,C4, TaConsult, False);
    If Assigned(TT_CONTRAT) then
    begin
      if C1 <> '' then
      begin
        TT_CONTRAT.Caption := C1;
      end;
      If C2 <> '' then
      begin
        TT_CONTRAT.Caption := TT_CONTRAT.Caption + '-' + C2;
      end;
      If C3 <> '' then
      begin
        TT_CONTRAT.Caption := TT_CONTRAT.Caption + '-' + C3;
      end;
      If C4 <> '' then
      begin
        TT_CONTRAT.Caption := TT_CONTRAT.Caption + '/' + C4;
      end;
    end;
    //
    TypeContrat := GetString('AFF_PERIODICITE');
    If Assigned(TT_TYPEINTER) then TT_TYPEINTER.Caption := RechDom('AFTPERIODICITE', TypeContrat, False);
  end;

End;

Procedure TOF_BTMESINTERV.AfficheContact(TobContact: TOB);
Begin

	With TobContact do
  Begin
    TT_CONTACT.caption    := GetString('C_NOM');
		TT_TELDOM.caption     := GetString('C_TELEPHONE');
    TT_TELBUR.Caption     := GetString('C_TELEX');
  	TT_TELDOM.Caption     := GetString('C_FAX');
		TT_EMAIL.Caption      := GetValue('C_RVA');
  End;

End;

Procedure TOF_BTMESINTERV.AfficheAdrInt(TobAdrInt: TOB);
Begin

  With TobAdrInt do
  Begin
    TT_CIVILITE.caption   := GetValue('ADR_JURIDIQUE');
    TT_LIBELLE.caption    := GetValue('ADR_LIBELLE');
    TT_PRENOM.caption     := GetValue('ADR_LIBELLE2');
    TT_ADRESSE1.caption   := GetValue('ADR_ADRESSE1');
    TT_ADRESSE2.caption   := GetValue('ADR_ADRESSE2');
    TT_ADRESSE3.caption   := GetValue('ADR_ADRESSE3');
    TT_CODEPOSTAL.caption := GetValue('ADR_CODEPOSTAL');
    TT_VILLE.caption      := GetValue('ADR_VILLE');
    TT_PAYS.caption       := GetValue('ADR_PAYS' );
    TT_REGION.caption     := GetValue('ADR_REGION');

    TT_TELDOM.caption     := GetValue('ADR_TELEPHONE');

    TT_CONTACT.caption    := GetValue('ADR_CONTACT');
    TT_EMAIL.caption      := GetValue('ADR_EMAIL' );
  end;

end;

//détermination de l'importance de l'appel
procedure TOF_BTMESINTERV.SetImportance;
begin

  if TobAppel.GetString('AFF_PRIOCONTRAT') = '3' then
  begin
     if Assigned(BtBasse)    Then BTBasse.visible := True;
  end
  else if TobAppel.GetString('AFF_PRIOCONTRAT') = '1' then
  begin
     if Assigned(BTHaute)    Then BTHaute.Visible := True;
  end
  else
  Begin
     if Assigned(BTNormale)  Then BTNormale.Visible := True;
  end;

end;

Procedure TOF_BTMESINTERV.InitGrille;
begin

  //Initialisation par défaut de la liste de lecture
  ListeParam := 'BTMESINTERV';

  GrilleConso.ColCount := 1;
	GrilleConso.RowCount := 1;

end;

procedure TOF_BTMESINTERV.Bsuivant_OnClick(Sender: TObject);
var NumPage       : Integer;
begin

  PageControl.SelectNextPage(True, True);
  Numpage := Pagecontrol.ActivePageIndex+1;

  if Numpage <= PageControl.PageCount - 1 then
  begin
    bprecedent.Visible := True;
  end;

  if Numpage = PageControl.PageCount - 1 then BSuivant.Visible := False;

  Pagecontrol.ActivePageIndex := NumPage;

  if PageControl.Pages[NumPage].Name = 'PAGE_2' then
  begin
    if Assigned(BParametre) then BParametre.Visible := V_PGI.Superviseur;;

    BAjout.Visible := True;
(*
    if TobConso.Detail.count > 0 then
    begin
      BValide.Visible := True;
    end;
*)
  end;
end;

procedure TOF_BTMESINTERV.BPrecedent_OnClick(Sender: TObject);
Var NumPage       : Integer;
begin

  PageControl.SelectNextPage(False, True);
  Numpage := Pagecontrol.ActivePageIndex;

  if Numpage >= 0 then
  begin
    BSuivant.Visible    := True;
    BPrecedent.Visible  := False;
    BAjout.visible      := False;
    Bparametre.Visible  := False;
  end;

  Pagecontrol.ActivePageIndex := NumPage;

end;

procedure TOF_BTMESINTERV.GestionSaisieNumAppel;
begin
  //
  CodeAppelDecoupe(CodeAppel, P0, P1, P2, P3, Avn);
  //
end;

procedure TOF_BTMESINTERV.LectureConso;
Var StSQL       : String;
    ZoneTable   : String;
    QConso  : TQuery;
begin

  If Not Assigned(TobAppel) then exit;

  //Initialisation de la TOB
  TOBConso.clearDetail;

  //Chargement de la grille en utilisant le paramétrage
  ZoneTable   := '';

  StSql  := GestGrille.ColGAppel;

  //Chargement de la TOB
  while  StSql <> '' do
     begin
     if ZoneTable = '' Then
	 	  	ZoneTable := ReadTokenst(StSql)
     else
	      ZoneTable := ZoneTable + ',' + ReadTokenst(StSql);
     end;

  if zoneTable = '' then ZoneTable := '*';

  StSQL := 'SELECT * ';
  StSQL := StSQL + ' FROM ' + GestGrille.TableGapp;
  StSQL := StSQL + ' WHERE BCO_AFFAIRE = "' + Codeappel +'" ';
  StSQL := StSQl + ' AND BCO_NATUREMOUV IN ("FOU", "MO")';

  if GestGrille.TriGapp <> '' then  StSQL := StSQL + ' ORDER BY ' + GestGrille.TriGapp;

  QConso := OpenSQL(StSQL, True);

  TobConso.LoadDetailDB('CONSOMMATIONS', '', '', QConso, False);

  Ferme(QConso);

end;


procedure TOF_BTMESINTERV.DesTroyTob () ;
begin
  FreeAndNil(TobAppel);
  FreeAndNil(TobOle);
  FreeAndNil(TobRessource);
  FreeAndNil(TobConso);
  FreeAndNil(TobTiers);
  FreeAndNil(TobContrat);
  FreeAndNil(TobContact);
  FreeAndNil(TobAdrInt);
  //
  FreeAndNil (TWindows);
  //
end ;

Procedure TOF_BTMESINTERV.CreateTOB;
begin

  TobConso    := Tob.create('Les Consommations', nil, -1);
  TobRessource:= TOB.Create('RESSOURCES', nil, -1);
  TobAppel    := Tob.create('AFFAIRE', nil, -1);
  TobTiers    := Tob.Create('TIERS',Nil, -1);
  TobContrat  := Tob.Create('CONTRAT',Nil, -1);
  TobContact  := Tob.Create('CONTACT',Nil, -1);
  TobAdrInt   := Tob.Create('ADRINT',Nil, -1);
  TobOle      := Tob.Create('LIENSOLE',Nil, -1);

end;

procedure TOF_BTMESINTERV.BParametre_OnClick(Sender: TObject);
begin


  SetScreenEvents(False);

  if assigned(BParametre) then
  begin
    {$IFDEF EAGLCLIENT}
    ParamListe(ListeParam, nil, '');
    {$ELSE}
    ParamListe(ListeParam, nil, nil, '');
    {$ENDIF}
  end;

  //Chargement de la liste Associées
  ChargeListeAssociee(ListeParam, 'Liste des Consommation par Appel', Nil, GestGrille);

  //Dessin de la grille en fonction des options de la liste
  DessineGrille(Nil, GrilleConso,GestGrille );

  RefreshGrille;


end;

procedure TOF_BTMESINTERV.GrilleConso_OnRowClick(Sender: TObject);

Var Arow        : Integer;
    TobEchange  : Tob;
    TOBCo       : Tob;
    NumMouv     : String;
begin

  if Assigned(GrilleConso) then
  begin
    Arow := GrilleConso.Row;
    if Arow < 0 then exit;
    if TOBConso.detail.count < Arow then exit;
    NumMouv := TobConso.detail[Arow-1].GetValue('BCO_NUMMOUV');

    TOBEchange := TOB.Create ('UN ECHANGE',nil,-1);

    With TobConso.Detail[Arow-1] Do
    begin
      TOBECHANGE.AddChampSupValeur('DATEMOUV', GetString('BCO_DateMouv'));

      TOBECHANGE.AddChampSupValeur('FamArt1', '');
      TOBECHANGE.AddChampSupValeur('FamArt2', '');
      TOBECHANGE.AddChampSupValeur('FamArt3', '');
      TOBECHANGE.AddChampSupValeur('CODEART', GetString('BCO_CODEARTICLE'));
      TOBECHANGE.AddChampSupValeur('ARTICLE', GetString('BCO_ARTICLE'));
      TOBECHANGE.AddChampSupValeur('LIBART',  GetString('BCO_LIBELLE'));
      TOBECHANGE.AddChampSupValeur('QTEART',  GetDouble('BCO_QUANTITE'));
      TOBECHANGE.AddChampSupValeur('TYPEARTICLE',  GetTypeArticle('BCO_ARTICLE'));

      TOBEchange.AddChampSupValeur('VALIDE', '');
      (*
      If GetString('BCO_NATUREMOUV') = 'MO' Then
        TOBEchange.AddChampSupValeur('TYPEARTICLE', 'PRE')
      else if GetString('BCO_NATUREMOUV') = 'FOU' Then
        TOBEchange.AddChampSupValeur('TYPEARTICLE', 'MAR');
      *)
      TheTOB := TOBEchange;
    end;

    AGLLanceFiche('BTP', 'BTSAISIECONSOMOB', '', '','APPEL=' + CodeAppel + '; NUMMOUV=' + NumMouv);

    if (TOBECHANGE.GetString('VALIDE')='M') then
    begin
      TOBCO := TobConso.detail[GrilleConso.Row-1];
      MajConso(TOBCO, TobEchange);
      ValorisationConsommation(TOBCO, TOBEchange);
    end
    else if (TOBECHANGE.GetString('VALIDE')='S') then
      BDelete_Click
  end;

  RefreshGrille;

end;

function TOF_BTMESINTERV.AddLigneConso : TOB;
begin
  result := TOB.Create ('CONSOMMATIONS',nil,-1);
end;

procedure TOF_BTMESINTERV.InitCreation(TOBCO, TOBEchange : TOB);
begin
  //
  if not SetUniqueNumber (TOBCO) then
  begin
    TOBCO.free;
    exit;
  end;
  //
  InitNewConso (TOBCo, TOBEchange);
  //
end;

procedure TOF_BTMESINTERV.MAJConso (TOBCO, TOBEchange : TOB);
var year        : Word;
    Month       : Word;
    Day         : word;
    TypeArticle : String;
    DateMouv    : TDateTime;
begin
  if Not Assigned(TOBCO) then exit;

  //
  TypeArticle := TOBEchange.GetString('TYPEARTICLE');
  //
  If TypeArticle = 'PRE' then
    TOBCO.PutValue('BCO_NATUREMOUV','MO')
  else if (TypeArticle = 'MAR') or (TypeArticle='ARP') Then
    TOBCO.PutValue('BCO_NATUREMOUV','FOU');
  //
  DateMouv := TobEchange.GetDateTime('DATEMOUV');
  DecodeDate (DateMouv,year,Month,Day);
  TOBCO.PutValue ('BCO_DATEMOUV', DateMouv);

  TOBCO.putValue('JOUR',IntToStr (day));
  TOBCO.putValue('BCO_SEMAINE',NumSemaine(TOBCO.GetValue('BCO_DATEMOUV')));
  TOBCO.PutValue('BCO_MOIS',month);

  TOBCO.putValue('BCO_QTEFACTUREE', TOBEchange.GetDouble('QTEART'));

  TOBCO.putValue ('BCO_QUANTITE', TOBEchange.GetDouble('QTEART'));

  TOBCO.PUTVALUE('BCO_LIBELLE',   TOBEchange.GetString('LIBART'));
  //
  //chargement par défaut de la ressource si affaire = appel
  TOBCO.putValue ('BCO_AFFAIRESAISIE', CodeContrat);
  TOBCO.putValue ('BCO_QTEFACTUREE', TOBEchange.GetDouble('QTEART'));

end;

procedure TOF_BTMESINTERV.InitNewConso (TOBCO, TOBEchange : TOB);
var year        : Word;
    Month       : Word;
    Day         : word;
    TypeArticle : String;
    DateMouv    : TDateTime;
    Aff0        : String;
    Aff1        : String;
    Aff2        : String;
    Aff3        : String;
    Aff4        : String;
    //
begin

  if Not Assigned(TOBCO) then exit;

  DateMouv := TobEchange.GetDateTime('DATEMOUV');

  // initialisation des champs de la table conso (voir SaisieConsommations -- InitNewConso
  DecodeDate (DateMouv,year,Month,Day);
  //
  TypeArticle := TOBEchange.GetString('TYPEARTICLE');
  //
  If TypeArticle = 'PRE' then
    TOBCO.PutValue('BCO_NATUREMOUV','MO')
  else if (TypeArticle = 'MAR') or (TypeArticle='ARP') Then
    TOBCO.PutValue('BCO_NATUREMOUV','FOU');
  //
  TOBCO.PutValue ('BCO_DATEMOUV', DateMouv);
  //
  TOBCO.PutValue ('BCO_AFFAIRE',  CodeAppel);
  BTPCodeAffaireDecoupe (TOBCO.GetValue('BCO_AFFAIRE'),Aff0,Aff1,Aff2,Aff3,Aff4,tacreat,false);
  TOBCO.PutValue ('BCO_AFFAIRE0',Aff0);
  TOBCO.PutValue ('BCO_AFFAIRE1',Aff1);
  TOBCO.PutValue ('BCO_AFFAIRE2',Aff2);
  TOBCO.PutValue ('BCO_AFFAIRE3',Aff3);

  //FV1 : 28/08/2013 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions
  TOBCO.SetString('DOMAINE', Domaine);
  //
  TOBCO.PutValue ('BCO_PHASETRA','');
  //
  TOBCO.putValue ('BCO_RESSOURCE', CodeUser);
  TOBCO.putValue ('LIBELLEMO', TobRessource.GetString('ARS_LIBELLE'));

  TOBCO.putValue('BCO_QUALIFQTEMOUV', TOBRessource.GetValue('ARS_UNITETEMPS'));

  TOBCO.putValue('BCO_QTEFACTUREE', TOBEchange.GetDouble('QTEART'));

  TOBCO.putValue('JOUR',IntToStr (day));
  TOBCO.putValue('BCO_SEMAINE',NumSemaine(TOBCO.GetValue('BCO_DATEMOUV')));
  TOBCO.PutValue('BCO_MOIS',month);

  TOBCO.putValue ('BCO_QUANTITE', TOBEchange.GetDouble('QTEART'));

  TOBCO.PutValue('BCO_ARTICLE',       TOBEchange.Detail[0].GetString('GA_ARTICLE'));
  TOBCO.PutValue('BCO_CODEARTICLE',   TOBEchange.Detail[0].GetString('GA_CODEARTICLE'));
  TOBCO.PUTVALUE('BCO_LIBELLE',       TOBEchange.GetString('LIBART'));

  //
  TOBCO.putValue('BCO_QUALIFQTEMOUV', TOBEchange.Detail[0].GetValue('GA_QUALIFUNITEVTE'));

  ValoModePrestation := GetParamSocSecur('SO_BTVALOAPPELS',True);

  //chargement par défaut de la ressource si affaire = appel
  TOBCO.putValue ('BCO_AFFAIRESAISIE', CodeContrat);
  TOBCO.putValue ('BCO_QTEFACTUREE', TOBEchange.GetDouble('QTEART'));

  if CodeContrat = '' then
		TOBCO.PutValue ('BCO_FACTURABLE','A')
  else
	  TOBCO.PutValue ('BCO_FACTURABLE','N');

end;

Procedure TOF_BTMESINTERV.ValorisationConsommation(TOBCO, TOBEchange : TOB);
var DPA         : Double;
    DPR         : Double;
    PUHT        : Double;
    //
    CoefMarge   : Double;
    CoefFG     : Double;
begin

  DPA := 0;
  DPR := 0;
  PUHT:= 0;
  //
  CoefMarge := 0;
  CoefFG    := 0;
  If (TOBEchange.GetString('TYPEARTICLE') = 'MAR') or (TOBEchange.GetString('TYPEARTICLE') = 'ARP')  then
  begin
    // ----------------------------------------------------------------------------------------------
    // Si article alors je recup de PA de l'article et si application systematique des coef de l'article alors je recup aussi le PR et le PV
    //  "  "      "                                       pas application alors je positionne les coef du domaine d'activité s'il existe
    // ----------------------------------------------------------------------------------------------
    //
    SetvaloArticle (TOBCO,TOBEchange.Detail[0]);
    //
    DPA   := TOBEchange.Detail[0].GetDouble('GA_PAHT');
    DPR   := TOBEchange.Detail[0].GetDouble('GA_DPR');
    PUHT   := TOBEchange.Detail[0].GetDouble('GA_PVHT');
    if TobEchange.Detail[0].GetString('GA_PRIXPASMODIF')<>'X' then
    begin
      if Domaine    <> '' Then  GetCoefDomaine(Domaine,CoefFG,CoefMarge, 'W');
      if CoefFG     <> 0  Then  DPR   := Arrondi(DPA*CoefFG,   V_PGI.okdecP);
      If CoefMarge  <> 0  Then  PUHT  := Arrondi(DPR*CoefMarge,V_PGI.okdecP);
    end;
  end
  else If TOBEchange.GetString('TYPEARTICLE') = 'PRE' then
  begin
    // --------------------------------------------------------------------------------------------------------
    // Si prestation (Mo) et valorisation a la ressource alors les PA PR PV --> ressource
    //    ++++++ Si Domaine activité alors j'applique les coef du domaine pour recalculer le PR et le PV
    // --------------------------------------------------------------------------------------------------------
    DPA   := TOBRessource.GetDouble('ARS_TAUXUNIT');
    DPR := TOBRessource.GetDouble('ARS_TAUXREVIENTUN');
    if DPR = 0 then DPR := TOBCO.GetDouble('BCO_DPA');
    PUHT   := TOBEchange.Detail[0].GetDouble('GA_PVHT');
    if ValoModePrestation then
    begin
      if (Domaine <> '') and (TobEchange.Detail[0].GetString('GA_PRIXPASMODIF')<>'X') then
      begin
        GetCoefDomaine(Domaine,CoefFG,CoefMarge, 'W');
        if CoefFG     <> 0  Then DPR   := Arrondi(DPA*CoefFG,   V_PGI.okdecP);
        If CoefMarge  <> 0  Then  PUHT  := Arrondi(DPR*CoefMarge,V_PGI.okdecP);
      end;
    end else
    begin
      if (Domaine <> '') then
      begin
        GetCoefDomaine(Domaine,CoefFG,CoefMarge, 'W');
        if CoefFG     <> 0  Then DPR   := Arrondi(DPA*CoefFG,   V_PGI.okdecP);
        If CoefMarge  <> 0  Then  PUHT  := Arrondi(DPR*CoefMarge,V_PGI.okdecP);
      end;
    end;
    //
  end;
  //
  TOBCO.PutVALUE('BCO_DPA',  DPA);
  TOBCO.PutVALUE('BCO_DPR',  DPR);
  TOBCO.PutVALUE('BCO_PUHT', PUHT);

  // Calcul des montants de la ligne
  CalculeLaLigne (TOBCO);

end;

procedure TOF_BTMESINTERV.BAjout_OnClick(Sender: TObject);
var TOBEchange,TOBCO : TOB;
begin

  TOBEchange := TOB.Create ('UN ECHANGE',nil,-1);

  TOBECHANGE.AddChampSupValeur('DATEMOUV', V_PGI.DateEntree);

  TOBECHANGE.AddChampSupValeur('FamArt1', '');
  TOBECHANGE.AddChampSupValeur('FamArt2', '');
  TOBECHANGE.AddChampSupValeur('FamArt3', '');
  TOBECHANGE.AddChampSupValeur('CODEART', '');
  TOBECHANGE.AddChampSupValeur('ARTICLE', '');
  TOBECHANGE.AddChampSupValeur('LIBART',  '');
  TOBECHANGE.AddChampSupValeur('QTEART',  0);

  TOBEchange.AddChampSupValeur('VALIDE', '');

  TOBEchange.AddChampSupValeur('TYPEARTICLE', '');

  TRY
    TheTOB  := TOBEchange;
    AGLLanceFiche('BTP', 'BTSAISIECONSOMOB', '', '','APPEL=' + CodeAppel );
    if (TOBECHANGE.GetString('VALIDE')='C') then
    begin
      TOBCO := AddLigneConso;
      TOBCO.InitValeurs;
      InitCreation(TOBCO, TOBEchange);
      ValorisationConsommation(TOBCo, TOBEchange);
      TOBCo.ChangeParent(TOBConso,-1);
      //
    end
    Else if (TOBECHANGE.GetString('VALIDE')='S') then
    Begin
      BDelete_Click;
    end;
  FINALLY
    TheTob := Nil;
    TOBEchange.free;
  END;

  SetScreenEvents(True);
  RefreshGrille;
(*
  if TobConso.detail.count > 0 Then
  Begin
    BValide.Visible := True;
  end;
*)

end;

function TOF_BTMESINTERV.SetUniqueNumber (TOBCO : TOB) : boolean;
var TheRetour : TGncERROR;
    UnNumero  : double;
begin

  result := true;

  TheRetour := GetNumUniqueConso (UnNumero);

  if TheRetour = gncAbort then
  Begin
    result := false;
    Exit;
  End;

  TOBCO.putValue ('BCO_NUMMOUV',UnNumero);

end;

Procedure TOF_BTMESINTERV.CreateTToolwindows;
var R             : Trect;
    SW, SH, W, H  : Integer;
Begin

  //création de tToolwindows de validation
  TWindows := TToolWindow97.Create(Ecran);
  TWindows.Parent       := Ecran;
  TWindows.BorderStyle  := bsNone;
  TWindows.CloseButton  := False;
  TWindows.Caption      := 'Validation Saisie';
  TWindows.Height       := 140;
  TWindows.Width        := 366;

  //permet de centrer la TtoolWindows par rapport à l'écran.
  SW := GetSystemMetrics(SM_CXSCREEN);
  SH := GetSystemMetrics(SM_CYSCREEN);
  R  := Twindows.ClientRect;
  GetWindowRect(TWindows.ParentWindow, R);
  W := Twindows.ClientRect.Right - Twindows.ClientRect.Left;
  H := Twindows.ClientRect.Bottom - Twindows.ClientRect.Top;
  R.Left := (SW - W) div 2;
  if R.Left < 0 then R.Left := 0;
  R.Top := (SH - H) div 2;
  if R.Top < 0 then R.Top := 0;
  TWindows.left := R.Left;
  TWindows.Top  := R.Top;
  TWindows.Resizable := False;

  //Création des Boutons Associés à la ttoolwindows
  BProvisoire := TtoolBarButton97.Create(Twindows);
  BProvisoire.Parent  := TWindows;
  BProvisoire.Caption := 'Validation Provisoire';
  Bprovisoire.Flat    := True;
  Bprovisoire.DisplayMode := dmTextOnly;
  BProvisoire.Anchors := [akLeft,akTop,akRight];
  Bprovisoire.Font.Name  := 'Tahoma';
  BProvisoire.Font.Style :=[fsBold];
  Bprovisoire.Font.Size  := 10;
  Bprovisoire.Height  := 52;
  Bprovisoire.Width   := 350;
  Bprovisoire.Top     := 0;
  Bprovisoire.Left    := 0;
  //
  BDefinitif  := TtoolBarButton97.Create(Twindows);
  BDefinitif.Parent   := TWindows;
  BDefinitif.Caption  := 'Validation Définitive';
  BDefinitif.Flat     := True;
  BDefinitif.DisplayMode  := dmTextOnly;
  Bdefinitif.Anchors  := [akLeft,akBottom,akRight];
  Bdefinitif.Font.Name  := 'Tahoma';
  Bdefinitif.Font.Style :=[fsBold];
  Bdefinitif.Font.Size  := 10;
  Bdefinitif.Height   := 52;
  Bdefinitif.width    := 350;
  Bdefinitif.Top      := 52;
  Bdefinitif.Left     := 0;

  Twindows.Visible      := False;

end;

procedure TOF_BTMESINTERV.BValide_Click(Sender: TObject);
begin

  TWindows.Visible    := True;
  BValide.Visible     := False;
  BQuitter.Visible    := False;
  BSuivant.Visible    := False;
  BPrecedent.Visible  := False;
  BParametre.Visible  := False;
  BAjout.Visible      := False;

  TWindows.Setfocus;

end;

procedure TOF_BTMESINTERV.BDelete_Click;
begin

  if PGIAsk('Désirez-vous réellement supprimer cet enregistrement', 'Suppression Ligne Intervention')=mrNo then exit;

  //suppression de l'enregistrement dans la Tobconso uniquement
  TobConso.detail[GrilleConso.Row-1].free;

  RefreshGrille;
(*
  if TobConso.detail.count = 0 Then
  Begin
    BValide.Visible := False;
  end;
*)
end;

//dans le cas d'une validation provisoire on mettra à jour la table des conso mais on laissera l'appel comme affecté
procedure TOF_BTMESINTERV.BProvisoire_Click(Sender: TObject);
Var StsQL : String;
begin

  ///suppression physique dans la table...
  StSQL := 'DELETE CONSOMMATIONS WHERE BCO_AFFAIRE="' + CodeAppel + '" ';

  ExecuteSQL(StSQL);

  Tobconso.SetAllModifie(True);
  Tobconso.InsertOrUpdateDB(True);

  TobAppel.PutValue('AFF_ETATAFFAIRE', 'ECR');
  TobAppel.SetAllModifie(True);
  TobAppel.UpdateDB(True,True);

  MajDescriptifRealisation;

  Ecran.ModalResult := mrOk;

  TWindows.Visible    := False;
  BValide.Visible     := True;
  BQuitter.Visible    := True;
  BSuivant.Visible    := True;
  BPrecedent.Visible  := True;
  BParametre.Visible  := True;
  BAjout.Visible      := True;

end;

//dans le cas d'une validation définitive on mettra à jour la table des consos mais on passera l'appel comme terminé
procedure TOF_BTMESINTERV.Bdefinitif_Click(Sender: TObject);
Var StSQL : String;
		TOBTT : TOB;
    TheDate : TDateTime;
begin
  TheDate := Now();
  TOBTT := TOB.Create ('LIENSOLE',nil,-1);
  TRY
    if not LanceFormeSignature (TOBTT) then exit;
    EnregistreSignature(CodeAppel,TOBTT);
  ///suppression physique dans la table...
  StSQL := 'DELETE CONSOMMATIONS WHERE BCO_AFFAIRE="' + CodeAppel + '" ';

  Tobconso.SetAllModifie(True);
  Tobconso.InsertOrUpdateDB(True);
  //
  TobAppel.PutValue('AFF_DATEFIN', TheDate);
  TobAppel.PutValue('AFF_ETATAFFAIRE', 'REA');
  TobAppel.SetAllModifie(True);
  TobAppel.UpdateDB(True,True);

  MajDescriptifRealisation;

  Ecran.ModalResult := mrOk;

  TWindows.Visible    := False;
  BValide.Visible     := True;
  BQuitter.Visible    := True;
  BSuivant.Visible    := True;
  BPrecedent.Visible  := True;
  BAjout.Visible      := True;
  BParametre.Visible  := True;
  FINALLY
    TOBTT.Free;
  end;

end;

Procedure TOF_BTMESINTERV.RefreshGrille;
Begin

  SetScreenEvents(False);

  AfficheGrille(GrilleConso, TobConso, GestGrille);

  HMTrad.ResizeGridColumns(GrilleConso);
  GrilleConso.DefaultRowHeight := 40;


  SetScreenEvents(True);
  GrilleConso.Refresh;

end;

Procedure TOF_BTMESINTERV.MajDescriptifRealisation;
begin

  if (Length(tt_descriptif1.Text) = 0) or (tt_descriptif1.Text = #$D#$A) then Exit;

  //Chargement des Tob liens OLE pour mise à jour du Résumé d'intervention
  TobOle.PutValue('LO_TABLEBLOB', 'APP');
  TobOle.PutValue('LO_QUALIFIANTBLOB', 'MOT');
  TobOle.PutValue('LO_EMPLOIBLOB', 'REA');
  TobOle.PutValue('LO_IDENTIFIANT', CodeAppel);
  TobOle.PutValue('LO_RANGBLOB', 1);

  TobOle.PutValue('LO_LIBELLE', 'Intervention sur ' + CodeAppel + 'le ' + DateToStr( V_PGI.DateEntree));

  TobOle.PutValue('LO_PRIVE', '-');
  TobOle.PutValue('LO_DATEBLOB',  V_PGI.DateEntree);
  TobOle.PutValue('LO_OBJET', TT_DESCRIPTIF1.Text);

  TobOle.SetAllModifie(True);
  TobOle.InsertOrUpdateDB(True);

end;

procedure TOF_BTMESINTERV.GrilleConso_PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

  DessinerLaCellule(GrilleConso, ACol, ARow, Canvas, AState);

end;
procedure TOF_BTMESINTERV.AfterFormShow;
begin
  //passage de la fiche en plein écran pour mode tablette
  PassagePleinEcran(Ecran);
end;

Initialization
  registerclasses ( [ TOF_BTMESINTERV ] ) ;
end.

