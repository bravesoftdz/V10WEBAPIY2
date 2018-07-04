{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 01/10/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : APPELPARRESS ()
Mots clefs ... : TOF;APPELPARRESS
*****************************************************************}
Unit BTAPPELPARRESS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Forms,
     Graphics,
     ComCtrls,
     DbCtrls,
{$IFDEF EAGLCLIENT}
      eMul,
      Maineagl,
      EFiche,
      UtileAGL,
      HQry,
{$ELSE}
      db,
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
      sysutils,
      HCtrls,
      HEnt1,
      HMsgBox,
      Hpanel,
      HTB97,
      HRichOle,
      vierge,
      Types,
      windows,
      UTOF,
      HDB,
      ParamSoc,
      ParamDBG,
      UTOM,
      MsgUtil,
      HSysMenu,
      Grids,
      utob;


Type
  TOF_APPELPARRESS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    ListeParam  : String;
    CodeUser    : String;
    UserName    : String;

    CodeAppel   : String;

    GestGrille  : AffGrille;

    PBarreBtn   : THPanel;

    PBouton     : TToolWindow97;

    Dock971     : TDock97;

    BParametre  : TToolBarButton97;
    BQuitter    : TToolBarButton97;
    BRefresh    : TToolBarButton97;

    Fliste      : THGrid;

    HMTrad      : ThSystemMenu;

    TobAppel    : TOB;

    //procedure AfficheGrille;
    Procedure BParametre_OnClick(Sender: TObject);
    procedure BRefresh_OnClick(Sender: TObject);
    procedure ChargeTOBAppelRessource;
    procedure DesTroyTob;
    procedure Fliste_RowClick(Sender: TObject);
    procedure GestionGrille;
    procedure GetObjects;
    procedure InitGrille;
    procedure PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure SetScreenEvents(Etat : boolean=true);

  end ;

Implementation

uses TntGrids,
     DateUtils,
     UtilScreen;

procedure TOF_APPELPARRESS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_APPELPARRESS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_APPELPARRESS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_APPELPARRESS.OnLoad ;
begin
  Inherited ;
  //passage de la fiche en plein écran pour mode tablette
  PassagePleinEcran(Ecran);
end ;

procedure TOF_APPELPARRESS.OnArgument (S : String ) ;
begin
  Inherited ;

  HMTrad    := ThsystemMenu.Create(Ecran);

  CodeUser  := V_PGI.User;
  UserName  := V_PGI.UserName;


  //Chargement des objets écran dans l'uses
  GetObjects;

  //Le bouton de paramétrage ne s'affichera que si on est administrateur....
  BParametre.Visible := V_PGI.Superviseur;

  //Initialisation de la grille
  InitGrille;

  GestionGrille;

  //Titre de la fiche
  if Assigned(PBarreBtn) then PBarreBtn.Caption := 'Liste des Interventions pour ' + UserName;

  //Définition des Evènement des zones écrans
  SetScreenEvents;

  //positionnement dans la grille pour attente action
  Fliste.SetFocus;

end ;

Procedure TOF_APPELPARRESS.GestionGrille;
begin

  //Chargement de la liste Associées
  ChargeListeAssociee(ListeParam, 'Liste des Appels par Ressource', Nil, GestGrille);

  //Dessin de la grille en fonction des options de la liste
  DessineGrille(Nil, Fliste,GestGrille );

  //Chargement des données de la table (LISTE/Vue)
  ChargeTOBAppelRessource;

  //Chargement des données de la table (LISTE/Vue)
  AfficheGrille(FListe, TobAppel, GestGrille);

  HMTrad.ResizeGridColumns(Fliste);
  FListe.DefaultRowHeight := 40;


end;

Procedure TOF_APPELPARRESS.ChargeTOBAppelRessource;
Var StSQL       : String;
    StDate      : TDateTime;
    ZoneTable   : String;
    QAppel		  : TQuery;
Begin

  Fliste.ClearSelected;

  //Initialisation de la TOB
  FreeAndNil(TobAppel);

  //Chargement de la grille en utilisant le paramétrage
  ZoneTable   := '';

  //recherche de la ressource associée à l'utilisateur
  RechRessourceAssociee(CodeUser,UserName);

  StDate := V_PGI.DateEntree;

  StSql  := GestGrille.ColGAppel;

  //Chargement de la TOB
  while  StSql <> '' do
     begin
     if ZoneTable = '' Then
	 	  	ZoneTable := ReadTokenst(StSql)
     else
	      ZoneTable := ZoneTable + ',' + ReadTokenst(StSql);
     end;

  TobAppel := Tob.create('Les Appels', nil, -1);

  StSQL := 'SELECT ' + ZoneTable + ' FROM ' + GestGrille.TableGapp;

  //On ne charge que les appels Affectés
  StSQL := StSQL + ' WHERE AFF_ETATAFFAIRE IN ("AFF", "ECR") ';

  //si code responsable pas renseigné on charge tout les salarié
  if CodeUser <> '' then
    StSql := StSQL + ' AND AFF_RESPONSABLE = "' + CodeUser + '" ';

  //si date pas renseignée ou erronée on prends la totalité
  if (StDate <> idate1900) Or (StDate <> Idate2099) then
    StSql := StSQL + 'AND AFF_DATESIGNE <= "' + UsDateTime(EndofTheDay(StDate)) + '" ';

  if GestGrille.TriGapp <> '' then  StSQL := StSQL + ' ORDER BY ' + GestGrille.TriGapp;

  QAppel := OpenSQL(StSQL, True);

  TobAppel.LoadDetailDB('APPEL', '', '', QAppel, False);

  Ferme(QAppel);

  StSql := GestGrille.ColGAppel;

end;

Procedure TOF_APPELPARRESS.InitGrille;
begin

  //Initialisation par défaut de la liste de lecture
  ListeParam := 'BTAPPELSRESSOURCE';

  FListe.ColCount := 1;
	FListe.RowCount := 1;

end;

Procedure TOF_APPELPARRESS.GetObjects;
begin

  if assigned(GetControl('PBouton'))    then PBouton    := TToolWindow97(GetControl('PBOUTON'));
  if Assigned(GetControl('Dock971'))    then Dock971    := TDock97(Getcontrol('Dock971'));
  //
  if Assigned(Getcontrol('PBARREBTN'))  then PBarreBtn  := THPanel(GetControl('PBARREBTN'));
  if Assigned(GetControl('BPARAMETRE')) then BParametre := TToolBarButton97(GetControl('BPARAMETRE'));
  if Assigned(GetControl('BQUITTER'))   then BQuitter   := TToolBarButton97(GetControl('BQUITTER'));
  If Assigned(GetControl('BREFRESH'))   then BRefresh   := TToolBarButton97(Getcontrol('BREFRESH'));
  if Assigned(GetControl('FLISTE'))     then Fliste     := THGrid (GetControl('FLISTE'));

end;

procedure TOF_APPELPARRESS.SetScreenEvents (Etat : boolean=true);
begin
  if Etat then
  begin
    BParametre.OnClick := BParametre_OnClick;
    BRefresh.Onclick   := BRefresh_OnClick;
    if assigned(FListe)    then
    begin
      FListe.OnClick       := FListe_RowClick;
      FListe.PostDrawCell  := PostDrawCell;
    end;
  end else
  begin
    BParametre.OnClick := nil;
    BRefresh.Onclick   := nil;
    if assigned(FListe)    then
    begin
      FListe.OnClick       := Nil;
      FListe.PostDrawCell  := Nil;
    end;
  end;
end;

procedure TOF_APPELPARRESS.OnClose ;
begin

  FreeAndNil(HMTrad);

  DestroyTob;

  Inherited ;

end ;

procedure TOF_APPELPARRESS.DesTroyTob () ;
begin
  FreeAndNil(TobAppel);
end ;

procedure TOF_APPELPARRESS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_APPELPARRESS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_APPELPARRESS.BParametre_OnClick(Sender: TObject);
begin

  if assigned(BParametre) then
  begin
    {$IFDEF EAGLCLIENT}
    ParamListe(ListeParam, nil, '');
    {$ELSE}
    ParamListe(ListeParam, nil, nil, '');
    {$ENDIF}
  end;

  GestionGrille;
                          
  BRefresh_OnClick(self);

end;

procedure TOF_APPELPARRESS.BRefresh_OnClick(Sender: TObject);
begin
  SetScreenEvents (false);

  //Chargement des données de la table (LISTE/Vue)
  ChargeTOBAppelRessource;

  //Chargement des données de la table (LISTE/Vue)
  AfficheGrille(FListe, TobAppel, GestGrille);

  HMTrad.ResizeGridColumns(Fliste);
  FListe.DefaultRowHeight := 40;
  
  SetScreenEvents (true);

end;

procedure TOF_APPELPARRESS.Fliste_RowCLick(Sender: TObject);
Var Arow : Integer;
begin


  if Assigned(Fliste) then
  begin
    Arow := Fliste.Row;
    CodeAppel := TobAppel.Detail[Arow-1].GetString('AFF_AFFAIRE');
    AGLLanceFiche('BTP', 'BTMESINTERV', '', '','APPEL=' + CodeAppel);
    Fliste.SetFocus;
  end;

  BRefresh_OnClick(self); 

end;

procedure TOF_APPELPARRESS.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

  DessinerLaCellule(Fliste, ACol, ARow, Canvas, AState);

end;

Initialization
  registerclasses ( [ TOF_APPELPARRESS ] ) ;
end.

