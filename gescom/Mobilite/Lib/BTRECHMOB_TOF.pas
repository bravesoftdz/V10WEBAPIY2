{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 09/10/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECHMOB ()
Mots clefs ... : TOF;BTRECHMOB
*****************************************************************}
Unit BTRECHMOB_TOF ;

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
  TOF_BTRECHMOB = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  Private

    ListeParam  : String;
    CritRech    : String;
    CodeRecup   : String;
    AlphaSel    : String;
    SaisInit    : string;
    OrderBy     : string; 
    //
    ZoneRech    : THEdit;
    //
    PBouton     : TToolWindow97;
    Dock971     : TDock97;
    //
    PBarreBtn   : THPanel;
    BParametre  : TToolBarButton97;
    BQuitter    : TToolBarButton97;
    //
    BRefresh    : TToolBarButton97;
    //
    Fliste      : THGrid;
    //
    GestGrille  : AffGrille;
    //
    TobRecherche: Tob;
    HMTrad      : ThSystemMenu;
    //
    LstAlpha    : TListBox;
    //
    procedure BParametre_OnClick(Sender: TObject);
    procedure BRefresh_OnClick(Sender: TObject);
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Valeur: String);
    procedure ChargeTOBRecherche;
    procedure FListe_OnRowClick(Sender: TObject);
    procedure GestionGrille;
    procedure GetObjects;
    procedure InitEcran;
    procedure LstAlpha_OnClick(Sender: TObject);
    procedure FListe_PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure SetScreenEvents;
    procedure ZoneRech_Change(Sender: Tobject);

  end ;

Implementation

uses TntGrids,
     DateUtils,
     UtilScreen;

procedure TOF_BTRECHMOB.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHMOB.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHMOB.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHMOB.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHMOB.OnArgument (S : String ) ;
var X       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;

  HMTrad    := ThsystemMenu.Create(Ecran);
  SaisInit := ''; 
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
      ControleCritere(Critere);
      Critere := (Trim(ReadTokenSt(S)));
    end;

  //Chargement des objets écran dans l'uses
  GetObjects;


  //remise à zéro des zone écran
  InitEcran;
  //
  if ListeParam      = 'BTRECHFAMNIV1MOB' then
  begin
    if Assigned(PBarreBtn) then PBarreBtn.Caption := 'Recherche Par Famille de Niveau 1';
  end
  else if ListeParam = 'BTRECHFAMNIV2MOB' then
  Begin
    if Assigned(PBarreBtn) then PBarreBtn.Caption := 'Recherche Par Famille de Niveau 2';
  end
  Else if ListeParam = 'BTRECHFAMNIV3MOB' then
  begin
    if Assigned(PBarreBtn) then PBarreBtn.Caption := 'Recherche Par Famille de Niveau 3';
  end
  Else if ListeParam = 'BTRECHARTMOBILE'  then
  begin
    if Assigned(PBarreBtn) then PBarreBtn.Caption := 'Recherche Par Article';
  end;
  //
  GestionGrille;
  if OrderBy <> '' then
  begin
  	GestGrille.TriGapp := OrderBy;
	end;
  //Définition des Evènement des zones écrans
  SetScreenEvents;
  //
  if SaisInit <> '' then
  begin
    ZoneRech.Text := SaisInit;
    BRefresh_OnClick(self);
  end;


end ;

Procedure TOF_BTRECHMOB.GetObjects;
begin

  if assigned(GetControl('PBouton'))      then PBouton    := TToolWindow97(GetControl('PBOUTON'));
  if Assigned(GetControl('Dock971'))      then Dock971    := TDock97(Getcontrol('Dock971'));
  //
  if Assigned(Getcontrol('PBARREBTN'))    then PBarreBtn  := THPanel(GetControl('PBARREBTN'));
  //
  if Assigned(GetControl('BPARAMETRE'))   then BParametre := TToolBarButton97(GetControl('BPARAMETRE'));
  if Assigned(GetControl('BQUITTER'))     then BQuitter   := TToolBarButton97(GetControl('BQUITTER'));
  If Assigned(GetControl('BRECHERCHE'))     then BRefresh   := TToolBarButton97(Getcontrol('BRECHERCHE'));
  //
  If Assigned(GetControl('LSTALPHA'))     then LstAlpha   := TListBox(GetControl('LSTALPHA'));
  //
  if Assigned(GetControl('GRILLE_RECH'))  then Fliste     := THGrid (GetControl('GRILLE_RECH'));
  //
  if Assigned(GetControl('ZoneRech'))     then ZoneRech   := THEdit (GetControl('ZONERECH'));

end;

Procedure TOF_BTRECHMOB.SetScreenEvents;
begin

  If Assigned(Fliste)     then
  begin
    Fliste.OnClick        := FListe_OnRowClick;
    FListe.PostDrawCell   := FListe_PostDrawCell;
  end;
  If assigned(BParametre) then BParametre.OnClick := BParametre_OnClick;
  If Assigned(BRefresh)   then BRefresh.OnClick   := BRefresh_OnClick;
  If Assigned(LstAlpha)   Then LstAlpha.OnClick   := LstAlpha_OnClick;
  If Assigned(ZoneRech)   Then ZoneRech.OnChange  := ZoneRech_Change;

end;

Procedure TOF_BTRECHMOB.ZoneRech_Change(Sender : Tobject);
Var ZoneDeTri : String;
begin

  //recherche de la première zone de tri
//  Zonedetri := ReadTokenst(GestGrille.TriGapp);

  if Assigned(ZoneRech) then
  begin
//    if ZoneDeTri <> '' then
//       AlphaSel := ZoneDeTri + ' LIKE "' + ZoneRech.text + '%"'
//    else
//       AlphaSel := CodeRecup + ' LIKE "' + ZoneRech.text + '%"';
       AlphaSel := OrderBy + ' LIKE "' + ZoneRech.text + '%"';
    //
    ChargeTOBRecherche;
    //
    AfficheGrille(FListe, TobRecherche, GestGrille);
    //
    if assigned(HMTrad) then HMTrad.ResizeGridColumns(Fliste);
    FListe.DefaultRowHeight := 40;
  end;

  ZoneRech.SetFocus;
  ZoneRech.SelStart := Length(ZoneRech.Text)+1;

end;

Procedure TOF_BTRECHMOB.LstAlpha_OnClick(Sender : TObject);
Var ZoneDeTri : String;
Begin
  (*
  //recherche de la première zone de tri
  Zonedetri := ReadTokenst(GestGrille.TriGapp);

  if Assigned(LstAlpha) then
  begin
    if ZoneDeTri <> '' then
       AlphaSel := ZoneDeTri + ' LIKE "' + LstAlpha.Items.Strings[LstAlpha.ItemIndex] + '%"'
    else
       AlphaSel := CodeRecup + ' LIKE "' + LstAlpha.Items.Strings[LstAlpha.ItemIndex] + '%"';
    //
    ChargeTOBRecherche;
    //
    AfficheGrille(FListe, TobRecherche, GestGrille);
    //
    if assigned(HMTrad) then HMTrad.ResizeGridColumns(Fliste);
    FListe.DefaultRowHeight := 40;
  end;
  *)
  ZoneRech.Text := '';
  if LstAlpha.Items.Strings[LstAlpha.ItemIndex] <> '...' then
  begin
  	ZoneRech.Text := LstAlpha.Items.Strings[LstAlpha.ItemIndex];
  end else if SaisInit <> '' then
  begin
    ZoneRech.Text := SaisInit;
  end;
  BRefresh_OnClick(self);
end;

Procedure TOF_BTRECHMOB.BParametre_OnClick(Sender : TObject);
begin

  {$IFDEF EAGLCLIENT}
  ParamListe(ListeParam, nil, '');
  {$ELSE}
  ParamListe(ListeParam, nil, nil, '');
  {$ENDIF}

end;

Procedure TOF_BTRECHMOB.FListe_OnRowClick(Sender : TObject);
begin

  if Assigned(Fliste) then
  begin
    LaTOB.PutValue('CODE', TobRecherche.Detail[Fliste.Row-1].GetString(CodeRecup));
    ecran.ModalResult := mrCancel;
    exit;
  end;

end;

procedure TOF_BTRECHMOB.BRefresh_OnClick(Sender: TObject);
begin

  //Chargement des données de la table (LISTE/Vue)
  ChargeTOBRecherche;

  if assigned(FListe) then
  begin
    //Chargement des données de la table (LISTE/Vue)
    AfficheGrille(FListe, TobRecherche, GestGrille);
    //
    if assigned(HMTrad) then HMTrad.ResizeGridColumns(Fliste);
    FListe.DefaultRowHeight := 40;
  end;

end;


Procedure TOF_BTRECHMOB.InitEcran;
begin

  //Initialisation par défaut de la liste de lecture
  If Assigned(FListe) then
  begin
    FListe.DefaultRowHeight := 40;

    FListe.ColCount := 1;
    FListe.RowCount := 1;
  end;

end;

procedure TOF_BTRECHMOB.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHMOB.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHMOB.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECHMOB.ControleCritere(Valeur: String);
begin
end;

procedure TOF_BTRECHMOB.ControleChamp(Champ, Valeur: String);
begin
  if Champ = 'LISTEPARAM' then ListeParam := Valeur;
  if Champ = 'CRITRECH'   then CritRech   := Valeur;
  If Champ = 'CODERECUP'  then CodeRecup  := Valeur;
  If Champ = 'SAISINIT'  then SaisInit  := Valeur;
  If Champ = 'ORDERBY'  then OrderBy  := Valeur;

end;

Procedure TOF_BTRECHMOB.GestionGrille;
begin

  //Chargement de la liste Associées
  ChargeListeAssociee(ListeParam, '???', Nil, GestGrille);

  //Dessin de la grille en fonction des options de la liste
  DessineGrille(Nil, Fliste,GestGrille );

  HMTrad.ResizeGridColumns(Fliste);
  FListe.DefaultRowHeight := 40;

end;

Procedure TOF_BTRECHMOB.ChargeTOBRecherche;
Var StSQL       : String;
    ZoneTable   : String;
    QRecherche  : TQuery;
Begin

  Fliste.ClearSelected;

  //Initialisation de la TOB
  FreeAndNil(TobRecherche);

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

  TobRecherche := Tob.create('Les Recherches', nil, -1);

  StSQL := 'SELECT ' + ZoneTable + ' FROM ' + GestGrille.TableGapp;

  if CritRech <> '' then StSQL := StSQL + ' WHERE ' + CritRech;

  if AlphaSel <> '' then
  begin
    if CritRech = '' then
      StSQl := StSQl + ' WHERE ' + AlphaSel
    Else
      StSQl := StSql + ' AND ' + AlphaSel;
  end;

  if GestGrille.TriGapp <> '' then  StSQL := StSQL + ' ORDER BY ' + GestGrille.TriGapp;

  Qrecherche := OpenSQL(StSQL, True);

  TobRecherche.LoadDetailDB('Recherche', '', '', QRecherche, False);

  Ferme(QRecherche);

  StSql := GestGrille.ColGAppel;

end;

procedure TOF_BTRECHMOB.FListe_PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

  DessinerLaCellule(Fliste, ACol, ARow, Canvas, AState);

end;

Initialization
  registerclasses ( [ TOF_BTRECHMOB ] ) ;
end.

