{***********UNITE*************************************************
Auteur  ...... : FV
Créé le ...... : 15/02/2012
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
unit BTVALIDEDEMPRIX_TOF;

interface
Uses StdCtrls,
     Controls,
     Classes,
     HCtrls,
{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
     uTob,
{$ELSE}
     DBCtrls, Db,
     {$IFNDEF DBXPRESS}
     dbTables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF DBXPRESS}
     fe_main,
     HDB,
     Mul,
{$ENDIF EAGLCLIENT}
     forms,
     sysutils,
     ComCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     uTOFComm,
     uEntCommun,
     uTob,
     Menus,
     Udefexport,
     HPanel,
     graphics,
     grids,
     Vierge,
     AglInit,
     DialogEx,
     UTOF;

Type

  TOF_BTVALIDEDEMPRIX = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  Private
    fFromMem      : boolean;
    TOBparam      : TOB;
    //
    NaturePiece   : THEdit;
    Souche        : THEdit;
    Numero        : THEdit;
    Indice        : THEdit;
    //
    Unique        : THValComboBox;
    TypeArticle   : THValComboBox;
    FamilleNiv1   : THValComboBox;
    FamilleNiv2   : THValComboBox;
    FamilleNiv3   : THValComboBox;
    NbLigHistoPx  : THValComboBox;
    //
    HTitre        : THLabel;
    HistoFrs      : THLabel;
    HistoRef      : THLabel;
    LibNbLig      : THLabel;
    LibNbLig2     : THLabel;
    //
    TheAction     : String;
    LesColArticle : string;
    LesColFourn   : string;
    LesColHisto   : String;
    FF            : String;
    STUnique      : String;
    LastValue     : string;
    //
    RechDdePrix   : TToolbarButton97;
    BtFicheFrs    : TToolBarButton97;
    BtLastPxAchat : TToolBarButton97;
    BTValider     : TToolBarButton97;
    BCherche      : TToolBarButton97;
    BTSupprime    : TToolBarButton97;
    //
    Panel_Rech    : THPanel;
    TPanelEntete  : THPanel;
    TPanelCorps   : THPanel;
    //
    GrilleArt     : THGrid;
    GrilleFrs     : THGrid;
    GrilleHistoPx : THGrid;
    //
    TToolHistoPx  : TToolWindow97;
    //
    ColRefArt     : Integer;
    ColLibelle    : Integer;
    ColQte        : Integer;
    ColUnite      : Integer;
    ColPrixAchat  : Integer;
    ColPrixVente  : Integer;
    ColTiers      : Integer;
    ColReference  : Integer;
    ColNbjour     : Integer;
    ColDatePiece  : Integer;
    ColPa         : Integer; // pour la saisie directe du PA fournisseur
    //
    Cledoc        : R_Cledoc;
    //
    TobPiece      : TOB;
    TobGestion    : TOB;
    TobFrsArt     : TOB;
    TobDemPrx     : TOB;
    TobDetailDemPx: TOB;
    //
    TobArtDde     : TOB;
    TobFrsLig     : TOB;
    //
    procedure AppelDemPrixOnClick(Sender: Tobject);
    procedure AppelFicheFrsOnClick(Sender: Tobject);
    procedure ChargeArtDemPrix;
    procedure ChargeCleDoc;
    procedure ChargeEnteteEcran;
    procedure ChargeHistoPx;
    procedure ChargeTobPiece;
    procedure ChercheOnClick(Sender: Tobject);
    procedure ConditionTOBArtDde;
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Critere: String);
    procedure DefinitActionZone;
    procedure DefiniEnteteHistoPx;
    procedure DefiniGrilleHistoPx;
    procedure DefiniToolHistoPx;
    procedure DefinitZone;
    procedure DessineLaGrille(Grille : ThGrid; LibColonne : String);

    function EnregInSelect(TOBA: TOB): boolean;
    function ExistCatalogueFrs(CodeFrs, CodeRef: String): Boolean;


    procedure GrilleArtDblclick(Sender: TObject);
    procedure GrilleArtEnter(Sender: TObject);
    procedure GrilleArtGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure GrilleArtRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrilleArtRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    //
    procedure GrilleFrsDblclick(Sender: TObject);
    procedure GrilleFrsEnter(Sender: TObject);
    procedure GrilleFrsExit(Sender: TObject);
    procedure GrilleFrsGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure GrilleFrsRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GrilleFrsCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleFrsCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    //
    procedure InitTob;
    //
    procedure NbLigOnchange(Sender: TObject);
    procedure OnCreateTOB;
    procedure OnDeleteRecord(Sender : Tobject);
    procedure RechDdePrixOnClick(Sender: Tobject);
    procedure RecupCompteAuxiFrs(var CodeFrs: String);
    procedure UniqueOnChange(Sender: Tobject);
    procedure ValideOnClick(Sender: Tobject);
    procedure AjoutFourClick (Sender : TObject);
    procedure SupprimeFourClick (sender : TObject);
    procedure InserelesFournisseurs(TOBDetailFour,TOBDemandeArticle : TOB);
    procedure ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function ZoneAccessible(ACol, ARow: Longint): boolean;
    procedure DessineLaGrilleArt(Grille: ThGrid; LibColonne: String);
    procedure ValideLaSaisie;

  end ;

const
  // libellés des messages de la TOM Affaire
  TexteMessage : array [1..6] of string = (
    {1} 'Confirmez-vous le retrait de cet article sur cette demande de prix ?',
    {2} 'Confirmez-vous la suppression de ce fournisseur dans cette demande de prix ?',
    {3} '',
    {4} '',
    {5} '',
    {6} ''
    ) ;


implementation
uses  UtilTOBPiece,
      UdemandePrix,
      facttob,
      EntGC,
      Windows,
      Messages;

{ TOF_BTVALIDEDEMPRIX}

procedure TOF_BTVALIDEDEMPRIX.OnArgument(S: String);
Var Critere : String;
    Champ   : String;
    Valeur  : Variant;
    X       : Integer;
    i       : Integer;
    Acol,Arow : Integer;
    CC : Boolean;
begin
  inherited;

  fFromMem := false;

  if LaTOB <> nil then
  begin
    TOBparam    := LaTOB;               //entete de saisie
    TobPiece    := TOB(TOBParam.data);  //Piece et ligne de Piece
    TobDemPrx   := TOB(TOBpiece.data);  //entete demande de Prix
    TOBgestion  := TOB(TobDemPrx.data);    //
    TobFrsArt   := TOB(TOBGestion.data);//
    TobDetailDemPx := Tob(TobFrsArt.Data);//
    fFromMem    := True;
  end;

  OnCreateTOB;

  LesColArticle := 'BDP_ARTICLE;BDP_LIBELLE;BDP_QTEBESOIN;BDP_QUALIFUNITEVTE;BDP_DPA;BDP_PUHTDEV';
  LesColFourn   := 'FIXED;BD1_TIERS;BD1_REFERENCE;BD1_PRIXACH;BD1_NBJOUR';
  LesColHisto   := 'FIXED;GL_DATEPIECE;GL_DPA;GL_QUALIFQTEVTE';

  DefinitZone;

  //Récupération valeur de argument
  Critere:=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
  begin
    if Critere <> '' then
      begin
      X := pos (';', Critere) ;
      if x = 0 then
        X := pos ('=', Critere) ;
      if x <> 0 then
        begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur:= Copy (Critere, X + 1, length (Critere) - X) ;
        ControleChamp(champ, valeur);
        end
      end;
    ControleCritere(Critere);
    Critere   := (Trim(ReadTokenSt(S)));
  end;

  FF:='#';
  if V_PGI.OkDecV>0 then
   begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
       begin
       FF:=FF+'0';
       end;
    FF:=FF+'0';
   end;
  ColNbJour     := -1;
  ColPa := -1;
  //dessin de la grille Article demande de prix
  DessineLaGrilleArt(GrilleArt,LesColArticle);

  //Dessin de la Grille Fournisseurs
  DessineLaGrille(GrilleFrs, LesColFourn);
  //
  GrilleFrs.SynEnabled := false;
  Acol := ColPa;
  Arow := GrilleFrs.Row;
  GrilleFrs.Col := Acol;
  GrilleFrs.SynEnabled := true;
  GrilleFrsCellEnter(Self,Acol,Arow,CC);
  //
  //gestion de l'historique des demandes de prix
  DefiniToolHistoPx;

  DefiniEnteteHistoPx;

  DefiniGrilleHistoPx;

  DessineLaGrille(GrilleHistoPx, LesColHisto);

  //chargement des zones Cledoc
  ChargeCleDoc;

  //si on ne vienty pas de la pièce on charge les tobs par lectures...
  if not fFromMem then
  begin
    InitTob;
    ChargeDemande(Cledoc,0, TobDemPrx, TobGestion, TobFrsArt, TobDetailDemPx);
  end;

  //chargement du combo de sélection des demandes de prix avec la TobDemPrix
  ChargeComboWithTob(Unique, TobDemPrx, 'BPP_LIBELLE', 'BPP_UNIQUE');

  //Si on a le numéro de demande de prix on la charge dans le combo
  if STUnique <> '' then
    Unique.Value := STUnique
  else
    Unique.itemindex := 0;

  //lecture de la pièce associée aux demandes de prix
  if Unique.value <> '' then ChargeTobPiece;

  if Tobpiece <> nil then ChargeEnteteEcran;

  //on charge la tob qui seraé affiché dans la grille en fonction des critères de sélection
  ChargeArtDemPrix;

  DefinitActionZone;
	GrilleArtEnter(self); 
end;


Procedure TOF_BTVALIDEDEMPRIX.ChargeEnteteEcran;
begin

  Htitre.Caption := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_LIBELLE');

  SetControlText('TGP_NUMERO', 'N° ' + TobPiece.GetString('GP_NUMERO'));
  SetControlText('TGP_DATEPIECE',  'Du ' + DateToStr(TobPiece.GetValue('GP_DATEPIECE')));
  SetControlText('TGP_TIERS',      'Tiers : ' + TobPiece.GetString('GP_TIERS'));
  SetControlText('TGP_AFFAIRE',    'Affaire : ' + TobPiece.GetString('GP_AFFAIRE'));
  SetControlText('TGP_REFEXTERNE', 'Ref. Externe : ' + TobPiece.GetString('GP_REFEXTERNE'));
  SetControlText('TGP_REFINTERNE', 'Ref. Interne : ' + TobPiece.GetString('GP_REFINTERNE'));

end;

Procedure TOF_BTVALIDEDEMPRIX.DefinitZone;
begin

  NaturePiece := THEdit(ecran.FindComponent('BPP_NATUREPIECEG'));
  Souche      := THEdit(ecran.FindComponent('BPP_SOUCHE'));
  Numero      := THEdit(ecran.FindComponent('BPP_NUMERO'));
  Indice      := THEdit(ecran.FindComponent('BPP_INDICEG'));
  //
  HTitre      := THLAbel(ecran.FindComponent('TBPP_PIECEG'));
  //
  Unique      := THValComboBox(ecran.FindComponent('SELDEMANDEPRIX'));
  TypeArticle := THValComboBox(ecran.FindComponent('TYPEARTICLE'));
  FamilleNiv1 := THValComboBox(ecran.FindComponent('FAMILLENIV1'));
  FamilleNiv2 := THValComboBox(ecran.FindComponent('FAMILLENIV2'));
  FamilleNiv3 := THValComboBox(ecran.FindComponent('FAMILLENIV3'));
  //
  RechDdePrix := TToolbarButton97(ecran.FindComponent('BTRECHERCHE'));
  BtFicheFrs  := TToolbarButton97(ecran.FindComponent('BTFICHEFRS'));
  BtLastPxAchat:= TToolbarButton97(ecran.FindComponent('BTLASTPXACHAT'));
  BTValider   := TToolbarButton97(ecran.FindComponent('BVALIDER'));
  BCherche    := TToolbarButton97(ecran.FindComponent('BTCHERCHE'));
  BtSupprime  := TToolBarButton97(ecran.FindComponent('BTSUPPRIME'));
  //
  Panel_Rech  := THPanel(ecran.findcomponent('PANEL_RECH'));
  //
  GrilleArt   := THGrid(ecran.FindComponent('GrilleArt'));
  GrilleFrs   := THGrid(ecran.FindComponent('GRILLEFRS'));
  //

end;

//creation de ttoolwindows container de la grille histo demande de prix !
procedure TOF_BTVALIDEDEMPRIX.DefiniToolHistoPx;
begin

  TToolHistoPx:= TToolWindow97.Create(ecran);

  TToolHistoPx            := TToolWindow97.Create (Ecran);
	TToolHistoPx.parent     := Ecran;
  TToolHistoPx.visible    := False;
  TToolHistoPx.CloseButton:= True;
	TToolHistoPx.fullsize   := false;
	TToolHistoPx.Resizable  := True;

  TToolHistoPx.Width      := Trunc(THPanel(ecran.FindComponent('PANEL_RIGHT')).width * 1.5);   //Panel container de la grille Frs
  TToolHistoPx.Height     := Trunc(THPanel(ecran.FindComponent('PANEL_RIGHT')).Height);  //Panel Container de la grille Frs

  TToolHistoPx.Top       := round(screen.height / 4);
  TToolHistoPx.Left      := round(screen.width / 4);

	TToolHistoPx.caption    := 'Historique Prix Achat';

end;

Procedure Tof_BtValideDEMPrix.DefiniEnteteHistoPx;
var iInc : integer;
begin

  TPanelEntete  := THPanel.Create (TToolHistoPx);
  TPanelEntete.parent   := TToolHistoPx;
  TPanelEntete.Align    := alTop;

  HistoFrs      := THLabel.Create(TPanelEntete);
  HistoFrs.parent := TPanelEntete;
  //
  HistoRef      := THLabel.Create(TPanelEntete);
  HistoRef.parent := TPanelEntete;
  //
  LibNbLig      := THLabel.Create(TPanelEntete);
  LibNbLig.Parent := TPanelEntete;
  //
  NbLigHistoPx  := THValComboBox.Create(TPanelentete);
  NbLigHistoPx.parent := TPanelEntete;
  NbLigHistoPx.visible:= true;
  For iinc := 1 to 5 do
  begin
    NbLigHistoPx.items.Add(intToStr(iinc*5));
  end;
  NbLigHistoPx.OnChange := NbLigOnchange;
  //
  LibNbLig2     := THLabel.Create(TPanelEntete);
  LibNbLig2.Parent := TPanelEntete;
  //
  TPanelEntete.height := HistoFrs.height * 5;

end;
//création de la grille Histo de prix Achat Frs dans container ttoolHistoPx
procedure TOF_BTVALIDEDEMPRIX.DefiniGrilleHistoPx;
begin

  TPanelCorps := THPanel.Create(TToolHistoPx);
  TpanelCorps.Parent  := TToolHistoPx;
  TPanelCorps.Align   := alClient;

  GrilleHistoPx := THGrid.create (TPanelCorps);
  GrilleHistoPx.parent     := TPanelCorps;
  GrilleHistoPx.Height     := TToolHistoPx.Height - TpanelEntete.Height;
  GrilleHistoPx.Width      := TToolHistoPx.Width - 120;
  GrilleHistoPx.Align      := alClient;
  GrilleHistoPx.ScrollBars := ssBoth ;
  GrilleHistoPx.Enabled    := True;
  GrilleHistoPx.Options    := GrilleHistoPx.Options+[goRowSelect];

end;

//chargement de l'historique des demande de prix...
procedure TOF_BTVALIDEDEMPRIX.ChargeHistoPx;
Var TOBHistoPx: Tob;
    Article   : String;
    CodeFrs   : String;
    StSQL     : String;
    QQ        : TQuery;
begin

  if TobArtDde.detail.count = 0 then exit;
  if (not assigned(TobFrsLig)) Or (TobFrsLig.detail.count = 0) then exit;

  TobHistoPx := Tob.Create('HISTOPX', nil, -1);

  Article := TobArtDde.Detail[GrilleArt.Row-1].GetString('BDP_ARTICLE');
  CodeFrs := TobFrsLig.Detail[GrilleFrs.Row-1].GetString('BD1_TIERS');

  HistoFrs.caption     := 'Fournisseur : ' + CodeFrs + ' ' + TobFrsLig.Detail[GrilleFrs.Row-1].GetString('T_LIBELLE');
  HistoRef.Caption     := 'Article  : ' + Article + ' ' + TobArtDde.Detail[GrilleArt.Row-1].GetString('BDP_LIBELLE');

  StSQL := 'SELECT ##TOP ' + NbLigHistoPx.Text + '##GL_DATEPIECE, GL_ARTICLE, GL_FOURNISSEUR, GL_DPA, GL_QUALIFQTEVTE ';
  StSQL := StSQL + 'GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG ';
  StSQL := StSQL + 'FROM LIGNE ';
  StSQL := StSQL + 'WHERE GL_FOURNISSEUR="' + CodeFrs + '" ';
  StSQL := StSQL + 'AND GL_ARTICLE="' + Article + '" ';
  StSQL := StSQL + 'ORDER BY GL_DATEPIECE DESC';

  QQ := OpenSQL(StSQL, True, -1, '', true);

  if Not QQ.eof then
  begin
    TobHistoPx.LoadDetailDB('HISTOPX', '','', QQ, False);
    TToolHistoPx.Height := TPanelEntete.height + (GrilleHistoPx.DefaultRowHeight * (TOBHistoPx.detail.count + 4));
  end;

  ferme(QQ);

  //chargement de la grille fournisseur
  if TOBHistoPx.detail.count <> 0 then
    ChargementGrilleWithTob(TobHistoPx, GrilleHistoPx, LesColHisto)
  else
    GrilleHistoPx.VidePile(False);

  FreeAndNil(TobHistoPx);

end;

Procedure TOF_BTVALIDEDEMPRIX.DefinitActionZone;
begin

  Unique.OnChange := UniqueOnChange;

  if assigned(RechDdePrix)   then RechDdePrix.OnClick := RechDdePrixOnClick;
  if assigned(BtFicheFrs)    then BtFicheFrs.OnClick  := AppelFicheFrsOnClick;
  if assigned(BtLastPxAchat) then BtLastPxAchat.OnClick  := AppelDemPrixOnClick;
  if assigned(BTValider)     then BTValider.Onclick   := ValideOnClick;
  if assigned(BCherche)      then BCherche.Onclick    := ChercheOnClick;
  if assigned(BTSupprime)    then BTSupprime.OnClick  := OnDeleteRecord;
  if assigned(GrilleArt) then
  begin
    GrilleArt.OnEnter      := GrilleArtEnter;
    GrilleArt.OnDblClick   := GrilleArtDblClick;
    GrilleArt.OnRowEnter   := GrilleArtRowEnter;
    GrilleArt.onRowExit 	 := GrilleArtRowExit;
    GrilleArt.GetCellCanvas:= GrilleArtGetCellCanvas;
  end;

  if Assigned(GrilleFrs) then
  begin
    GrilleFrs.OnEnter      := GrilleFrsEnter;
    GrilleFrs.OnExit       := GrilleFrsExit;
    GrilleFrs.OnDblClick   := GrilleFrsDblClick;
    GrilleFrs.GetCellCanvas:= GrilleFrsGetCellCanvas;
    GrilleFrs.OnRowEnter   := GrilleFrsRowEnter;
    GrilleFrs.OnCellExit   := GrilleFrsCellExit;
    GrilleFrs.OnCellEnter  := GrilleFrsCellEnter;
  end;
  //
  if Assigned(TMenuItem (GetControl('MnAjoutFour'))) then TMenuItem (GetControl('MnAjoutFour')).onclick := AjoutFourClick;
  if Assigned(TMenuItem (GetControl('MnSupFour')))   then TMenuItem (GetControl('MnSupFour')).onclick := SupprimeFourClick;
  if Assigned(TMenuItem (GetControl('MnFicheFour'))) then TMenuItem (GetControl('MnFicheFour')).onclick := AppelFicheFrsOnClick;


end;

procedure TOF_BTVALIDEDEMPRIX.UniqueOnChange(Sender : Tobject);
begin

  ChargeArtDemPrix;

end;

procedure TOF_BTValideDEMPRIX.NbLigOnchange(Sender : TObject);
begin

  ChargeHistoPx;

end;

Procedure TOF_BTVALIDEDEMPRIX.RechDdePrixOnClick(Sender : Tobject);
begin

  Panel_Rech.Visible := not Panel_Rech.Visible;

end;

Procedure TOF_BTVALIDEDEMPRIX.AppelFicheFrsOnClick(Sender : Tobject);
var CodeFournisseur : String;
    CodeReference   : String;
begin

  if (not assigned(TobFrsLig)) Or (TobFrsLig.Detail.count = 0) then exit;

  if GrilleFrs.Row < 1 then exit;

  CodeFournisseur := TobFrsLig.Detail[GrilleFrs.Row-1].GetString('BD1_TIERS');
  CodeReference   := TobFrsLig.Detail[GrilleFrs.Row-1].GetString('BD1_REFERENCE');

  if ExistcatalogueFrs(CodeFournisseur, Codereference) then
    AGLLanceFiche('GC', 'GCCATALOGU_SAISI3', '', CodeReference + ';' + CodeFournisseur, 'ACTION=CONSULTATION')
  else
  begin
    RecupCompteAuxiFrs(CodeFournisseur);
    AGLLanceFiche('GC','GCTIERS','',CodeFournisseur,'ACTION=CONSULTATION;T_NATUREAUXI=FOU');
  end;

end;

Function TOF_BTVALIDEDEMPRIX.ExistCatalogueFrs(CodeFrs, CodeRef : String) : Boolean;
var StSQl : String;
begin

  StSQL := 'SELECT ##TOP 1## GCA_TIERS FROM CATALOGU WHERE GCA_TIERS="' + CodeFrs + '" AND GCA_REFERENCE="' + CodeRef + '"';

  Result := ExisteSQL(StSQL);

end;

Procedure TOF_BTVALIDEDEMPRIX.RecupCompteAuxiFrs(var CodeFrs : String);
var StSQl : String;
    QQ    : TQuery;
begin

  StSQL := 'SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="' + CodeFrs + '" AND T_NATUREAUXI="FOU"';

  QQ := OpenSQL(StSQL, True, -1, '', true);

  if not QQ.eof then
    CodeFrs := QQ.findfield('T_AUXILIAIRE').AsString
  else
    CodeFrs := '';

  ferme(QQ);

end;

Procedure TOF_BTVALIDEDEMPRIX.AppelDemPrixOnClick(Sender : Tobject);
begin

  TToolHistoPx.visible := True;

  HistoFrs.left     := 8;
  HistoFrs.Top      := 8;

  HistoRef.Left     := 8;
  HistoRef.Top      := HistoFrs.Top + HistoFrs.Height + 5;

  LibNbLig.Caption  := 'Afficher les ';
  LibNbLig.left     := 8;
  LibNbLig.Top      := HistoRef.Top + HistoRef.Height + 5;

  NbLigHistoPx.Width:= 60;
  NbLigHistoPx.ItemIndex := 1;
  NbLigHistoPx.Top  := LibNbLig.Top;
  NbLigHistoPx.left := LibNbLig.Left + LibNbLig.Width + 8;

  LibNbLig2.Caption := ' derniers Prix d''achat ';
  LibNbLig2.Top     := LibNbLig.Top;
  LibNbLig2.left    := NbLigHistoPx.left + NbLigHistoPx.Width + 8;

  ChargeHistoPx;

end;

Procedure TOF_BTVALIDEDEMPRIX.OnCreateTOB;
begin
  //
  if not fFromMem then
  begin
    TOBPiece        := Tob.Create('PIECE', nil, -1);
    TobDemPrx       := tob.create('LES ENTETEDEMPRIX', nil, -1);
    TobGestion      := Tob.Create('LES ARTICLEDEMPRIX', nil, -1);
    TobFrsArt       := Tob.Create('LES FOURNLIGDEMPRIX', nil, -1);
    TobDetailDemPx  := Tob.Create('LES DETAILLIGDEMPRIX', nil, -1);
  end;
  //
  TobArtDde:= Tob.Create('LES ARTICLEDEMPRIX', nil, -1);
  //
end;

procedure TOF_BTVALIDEDEMPRIX.OnCancel;
begin
  inherited;
end;

procedure TOF_BTVALIDEDEMPRIX.OnClose;
begin
  if not fFromMem then
  begin
    FreeAndNil(TobPiece);
    FreeAndNil(TobFrsArt);
    FreeAndNil(TobGestion);
  end;
  FreeAndNil(TobArtDde);

  FreeAndNil(GrilleHistoPx);
  FreeAndNil(TToolHistoPx);

  inherited;

end;

procedure TOF_BTVALIDEDEMPRIX.OnDisplay;
begin
  inherited;

end;

procedure TOF_BTVALIDEDEMPRIX.OnLoad;
begin
  inherited;
end;


Procedure TOF_BTVALIDEDEMPRIX.InitTob;
begin
  TOBPiece.ClearDetail;
  Tobgestion.ClearDetail;
  TobDemPrx.ClearDetail;
  TobArtDde.ClearDetail;
  TobFrsArt.ClearDetail;
  TobDetailDemPx.ClearDetail;
end;

Procedure TOF_BTVALIDEDEMPRIX.ChargeTobPiece;
Var StSQL : String;
    QQ    : TQuery;
begin

  if fFromMem then exit;

  StSQL := 'SELECT * FROM PIECE INNER JOIN PIECEDEMPRIX ON GP_NATUREPIECEG=BPP_NATUREPIECEG ';
  StSQL := StSQL + 'AND GP_SOUCHE=BPP_SOUCHE ';
  StSQL := StSQL + 'AND GP_NUMERO=BPP_NUMERO ';
  StSQL := StSQL + 'AND GP_INDICEG=BPP_INDICEG ';
  StSQL := StSQL + 'WHERE ' + WherePiece(Cledoc, ttdPieceDemPrix, true) + ' AND BPP_UNIQUE='+ Unique.Value;
  QQ := OpenSQL(StSQL, True);

  if Not QQ.eof then
  begin
    TobPiece.SelectDB('', QQ);
  end;

  ferme (QQ);

end;

procedure TOF_BTVALIDEDEMPRIX.OnNew;
begin
  inherited;

end;

procedure TOF_BTVALIDEDEMPRIX.ValideLaSaisie;
begin
  ExecuteSQL('DELETE FROM PIECEDEMPRIX WHERE '   + WherePiece(CleDoc, ttdPieceDemPrix   , False));
  ExecuteSQL('DELETE FROM DETAILDEMPRIX WHERE '  + WherePiece(CleDoc, TtdDetailDemPrix  , False));
  ExecuteSQL('DELETE FROM ARTICLEDEMPRIX WHERE ' + WherePiece(CleDoc, TTdArticleDemPrix , False));
  ExecuteSQL('DELETE FROM FOURLIGDEMPRIX WHERE ' + WherePiece(CleDoc, TtdFournDemprix   , False));
  ValidelesDemPrix (TOBPiece,TobDemPrx,TobGestion,TobFrsArt,TobDetailDemPx);
end;

procedure TOF_BTVALIDEDEMPRIX.OnUpdate;
var IoResult : TIOErr;
		Indice : Integer;
    okOk : Boolean;
begin
  inherited;
  OkOk := true;
  if fFromMem then
  begin
  	exit;
  end;

  For Indice := 0 to TobArtDde.detail.count -1 do
  begin
		if not TobArtDde.detail[Indice].GetBoolean('BDP_TRAITE') then
    begin
      okOk := false;
      break;
    end;
  end;
  TobDemPrx.detail[0].SetBoolean('BPP_TRAITE',Okok);
  IoResult := TRANSACTIONS(VALIDELASAISIE,0);

end;

procedure TOF_BTVALIDEDEMPRIX.ControleChamp(Champ, Valeur: String);
begin

  if Champ = 'NATUREPIECEG' then
    NaturePiece.text := Valeur
  else if Champ = 'SOUCHE' then
    Souche.text := Valeur
  else if Champ = 'NUMERO' then
    Numero.Text := Valeur
  else if Champ = 'INDICEG' then
    Indice.text := Valeur
  else if Champ = 'UNIQUE' then
    StUnique    := Valeur
  else if champ = 'ACTION' then
    TheAction   := Valeur;

end;

procedure TOF_BTVALIDEDEMPRIX.ControleCritere(Critere: String);
begin
end;

Procedure TOF_BTVALIDEDEMPRIX.ChargeCleDoc;
begin

  if fFromMem then
  begin
    cledoc := TOB2Cledoc(TOBpiece);
    NaturePiece.Text := cledoc.NaturePiece ;
    Souche.Text := cledoc.Souche;
    Numero.text := IntToStr(cledoc.NumeroPiece ) ;
    Indice.text := IntToStr(Cledoc.indice );
  end else
  begin
    cledoc.NaturePiece := NaturePiece.Text;
    cledoc.Souche      := Souche.Text;
    cledoc.NumeroPiece := StrToInt(Numero.text);
    Cledoc.indice      := StrToInt(Indice.text);
  end

end;
procedure TOF_BTVALIDEDEMPRIX.DessineLaGrilleArt(Grille : ThGrid;LibColonne : String);
var st      : string;
    Nam     : String;
    i       : Integer;
    depart  : Integer;
begin
  //
  Grille.FixedCols := 1;
  Grille.FixedRows := 1;
  //
  Grille.DefaultRowHeight := 18;
  //
  Grille.ColCount     := 7;
  Grille.ColWidths[0] := 10;
  //
  St := LibColonne;
  //
  ColRefArt     := -1;
  ColLibelle    := -1;
  ColQte        := -1;
  ColUnite      := -1;
  ColPrixAchat  := -1;
  ColPA         := -1;
  ColPrixVente  := -1;
  //

  depart := 1;
  for i := 1 to Grille.ColCount -1 do
  begin
    if i > depart then Grille.ColWidths[i] := 100;
    Nam := ReadTokenSt(St);
    if Nam = 'BDP_ARTICLE' then
    begin
      Grille.ColWidths[i] := 100;
      Grille.ColEditables[i] := False;
      ColRefArt := i;
    end
    else if Nam = 'BDP_LIBELLE' then
    begin
      Grille.ColWidths[i] := 280;
      Grille.ColEditables[i] := False;
      ColLibelle := i;
    end
    else if Nam = 'BDP_QTEBESOIN' then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColFormats[i] := FF;
      Grille.ColAligns[i] := taRightJustify;
      Grille.ColEditables[i] := False;
      ColQte := i;
    end
    else if Nam = 'BDP_QUALIFUNITEVTE' then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColAligns[i] := taCenter;
      Grille.ColEditables[i] := False;
      ColUnite := i;
    end
    else if (Pos(Nam ,'BDP_DPA')>0) then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColFormats[i] := FF;
      Grille.ColAligns[i] := taRightJustify;
      Grille.ColEditables[i] := false;
      ColPrixAchat := i;
    end
    else if (Nam = 'BDP_PUHTDEV') then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColFormats[i] := FF;
      Grille.ColAligns[i] := taRightJustify;
      ColPrixVente := i;
      Grille.ColEditables[i] := true;
    end;
  end;
  //
  if ColRefArt     >=0 then Grille.Cells[ColRefArt,0]    := 'Ref. Article';
  if ColLibelle    >=0 then Grille.Cells[ColLibelle,0]   := 'Libellé';
  if ColQte        >=0 then Grille.Cells[ColQte,0]       := 'Qté';
  if ColUnite      >=0 then Grille.Cells[ColUnite,0]     := 'Unité';
  if ColPrixAchat  >=0 then Grille.Cells[ColPrixAchat,0] := 'Px Achat';
  if ColPrixVente  >=0 then Grille.Cells[ColPrixVente,0] := 'Px Vente';
  //

end;


procedure TOF_BTVALIDEDEMPRIX.DessineLaGrille(Grille : ThGrid; LibColonne : String);
var st      : string;
    Nam     : String;
    i       : Integer;
    depart  : Integer;
begin
  //
  Grille.FixedCols := 1;
  Grille.FixedRows := 1;
  //
  Grille.DefaultRowHeight := 18;
  //
  Grille.ColCount     := findOccurenceString(LibColonne, ';') + 1;
  Grille.ColWidths[0] := 10;
  //
  St := LibColonne;
  //
  ColPrixAchat  := -1;
  ColPrixVente  := -1;
  ColTiers      := -1;
  ColReference  := -1;
  ColDatePiece  := -1;
  //
  for i := 0 to Grille.ColCount - 1 do
  begin
    depart := 1;
    if i > depart then Grille.ColWidths[i] := 100;
    Nam := ReadTokenSt(St);
    if Nam = 'BDP_ARTICLE' then
    begin
      Grille.ColWidths[i] := 100;
      Grille.ColEditables[i] := False;
      ColRefArt := i;
    end
    else if Nam = 'BDP_LIBELLE' then
    begin
      Grille.ColWidths[i] := 300;
      Grille.ColEditables[i] := False;
      ColLibelle := i;
    end
    else if Nam = 'BDP_QTEBESOIN' then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColFormats[i] := FF;
      Grille.ColAligns[i] := taRightJustify;
      Grille.ColEditables[i] := False;
      ColQte := i;
    end
    else if Nam = 'BDP_QUALIFUNITEVTE' then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColAligns[i] := taCenter;
      Grille.ColEditables[i] := False;
      ColUnite := i;
    end
    else if (Pos(Nam ,'GL_DPA;BDP_DPA;BD1_PRIXACH')>0) then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColFormats[i] := FF;
      Grille.ColAligns[i] := taRightJustify;
      Grille.ColEditables[i] := False;
      ColPrixAchat := i;
      if Nam = 'BD1_PRIXACH' then
      begin
      	ColPa := i;
      	Grille.ColEditables[i] := true;
      end;
    end
    else if (Nam = 'BDP_PUHTDEV') then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColFormats[i] := FF;
      Grille.ColAligns[i] := taRightJustify;
      Grille.ColEditables[i] := False;
      ColPrixVente := i;
    end
    Else if Nam = 'BD1_TIERS' then
    begin
      Grille.ColWidths[i] := 100;
      Grille.ColEditables[i] := False;
      ColTiers := i;
    end
    Else if Nam = 'BD1_REFERENCE' then
    begin
      Grille.ColWidths[i] := 100;
      Grille.ColEditables[i] := False;
      ColReference := i;
    end
    Else if Nam = 'BD1_NBJOUR' then
    begin
      Grille.ColWidths[i] := 50;
      Grille.ColFormats[i] := '#0';
      Grille.ColAligns[i] := taRightJustify;
      Grille.ColEditables[i] := true;
      ColNbJour := i;
    end
    else if Nam = 'GL_DATEPIECE' Then
    Begin
      Grille.ColWidths[i] := 50;
      Grille.ColFormats[i] := 'dd/mm/yyyy';
      Grille.ColAligns[i] := taCenter;
      Grille.ColEditables[i] := False;
      ColDatePiece := i;
    end
    else if Nam = 'GL_QUALIFQTEVTE' then
    begin
      Grille.ColWidths[i] := 60;
      Grille.ColAligns[i] := taCenter;
      Grille.ColEditables[i] := False;
      ColUnite := i;
    end;
  end;
  //
  if ColRefArt     >=0 then Grille.Cells[ColRefArt,0]    := 'Ref. Article';
  if ColLibelle    >=0 then Grille.Cells[ColLibelle,0]   := 'Libellé';
  if ColQte        >=0 then Grille.Cells[ColQte,0]       := 'Qté';
  if ColUnite      >=0 then Grille.Cells[ColUnite,0]     := 'Unité';
  if ColPrixAchat  >=0 then Grille.Cells[ColPrixAchat,0] := 'Px Achat';
  if ColPrixVente  >=0 then Grille.Cells[ColPrixVente,0] := 'Px Vente';
  if ColTiers      >=0 Then Grille.Cells[ColTiers, 0]    := 'Fournisseur';
  if ColReference  >=0 Then Grille.Cells[ColReference, 0]:= 'Ref. Art. Frs';
  if ColNbJour     >=0 Then Grille.Cells[ColNbjour, 0]   := 'Nb Jours';
  if ColDatePiece  >=0 Then Grille.Cells[ColDatePiece,0] := 'DatePiece';
  //

end ;

procedure TOF_BTVALIDEDEMPRIX.GrilleArtEnter(Sender: TObject);
var Arow  : integer;
    cancel: boolean;
begin

  cancel := false;
  if GrilleArt.RowCount = 1 then
  begin
    // pas de ligne détail
    GrilleArt.rowCount := 2;
  end;
  Arow := GrilleArt.row;
  if Arow = 0 then Arow := 1;
  if Arow > TobArtDde.detail.count - 1 then  Arow := 1;

  GrilleArtRowEnter(self, Arow, cancel, false);

end;


procedure TOF_BTVALIDEDEMPRIX.GrilleArtRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end;


procedure TOF_BTVALIDEDEMPRIX.GrilleArtRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
Var IInd : Integer;
begin

  if TobArtDde.detail.count > 0 then
  begin
    IInd := ou - 1;
    BtFicheFrs.Visible    := True;
    BtLastPxAchat.Visible := True;
    BTSupprime.Visible    := True;
    TobFrsLig := FindFournDemPrix(TobArtDde.detail[IInd],TobFrsArt);
  end else
  begin
    BtFicheFrs.Visible    := False;
    BtLastPxAchat.Visible := False;
    BTSupprime.Visible    := False;
    TobFrsLig :=nil;
  end;

  //chargement de la grille fournisseur
  ChargementGrilleWithTob(TobFrsLig, GrilleFrs, LesColFourn);
  GrilleFrsEnter(Self);
  if TToolHistoPx.Visible then ChargeHistoPx;

end;

procedure TOF_BTVALIDEDEMPRIX.GrilleArtDblclick(Sender: TObject);
var ArticleAff : String;
begin

  if TobArtDde.detail.count = 0 then exit;

  ArticleAff := TobArtDde.detail[GrilleArt.row-1].GetString('BDP_ARTICLE');
  AGLLanceFiche('BTP', 'BTARTICLE', '', ArticleAff, 'ACTION=CONSULTATION');

end;

procedure TOF_BTVALIDEDEMPRIX.GrilleArtGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

  if TobArtDde.Detail.count = 0 then exit;

  if ACol < GrilleFrs.FixedCols then Exit;
  IF Arow < GrilleFrs.FixedRows then exit;
  //
  if TobArtDde.Detail[Arow-1].GetBoolean('BDP_SELECTIONNE') then
  BEGIN
    canvas.Brush.Color:= clGray;
    canvas.Font.Color := clYellow;
  END;

end;

procedure TOF_BTVALIDEDEMPRIX.GrilleFrsEnter(Sender: TObject);
var Arow,Acol  : integer;
    cancel: boolean;
begin

  cancel := false;

  Arow := GrilleFrs.row;
  if GrilleFrs.rowCount < 2 then
  begin
    GrilleFrs.rowCount := 2;
    Arow := 1;
  end;

  if assigned(TobFrsLig) then
  begin
    if Arow > TobFrsLig.detail.count - 1 then  Arow := 1;
    
    GrilleFrs.SynEnabled := false;
    Acol := ColPa;
    Arow := ARow;
    GrilleFrs.Col := Acol;
    GrilleFrs.SynEnabled := true;

    GrilleFrsRowEnter(self, Arow, cancel, false);
  	GrilleFrsCellEnter(Self,Acol,Arow,cancel);
  end;

end;

procedure TOF_BTVALIDEDEMPRIX.GrilleFrsRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

  if TToolHistoPx.Visible then ChargeHistoPx;

end;

procedure TOF_BTVALIDEDEMPRIX.GrilleFrsDblclick(Sender: TObject);
Var iInd      : Integer;
    ARow      : Integer;
    RowArt    : Integer;
begin

  if not assigned(TobFrsLig) then exit;
  if not assigned(TobArtDde) then exit;

  if TobFrsLig.detail.count = 0 then exit;
  if TobArtDde.detail.count = 0 then exit;

  ARow    := GrilleFrs.Row-1;
  RowArt  := GrilleArt.Row-1;
(*
  //si la ligne a déjà été traitée il ne faut pas pouvoir intervenir dessus...
  if TobArtDde.detail[RowArt].GetBoolean('BDP_TRAITE') then exit;
*)
  //Déselection de la ligne fournisseur sans prise en compte d'une nlle ligne
  if TobFrsLig.Detail[ARow].GetBoolean('BD1_SELECTED') then
  begin
    TobFrsLig.Detail[ARow].PutValue('BD1_SELECTED', '-');
    TobArtDde.Detail[RowArt].PutValue('BDP_SELECTIONNE', '-');
    TOB(TobArtDde.Detail[RowArt].Data).PutValue('BDP_SELECTIONNE', '-');
    TobArtDde.Detail[RowArt].PutValue('BDP_TRAITE', '-');
    TOB(TobArtDde.Detail[RowArt].Data).PutValue('BDP_TRAITE', '-');
  end else
  begin
    //si la ligne à déjà été traitée il faut déselectionner la ligne fournisseurs selected !!!
    if TobArtDde.Detail[RowArt].GetBoolean('BDP_SELECTIONNE') then
    begin
      for iInd := 0 to TobFrsLig.Detail.count - 1 do
      begin
        TobFrsLig.Detail[iInd].PutValue('BD1_SELECTED', '-');
      end;
    end;
    //Mise à jour des tobs après sélection
    TobFrsLig.Detail[ARow].PutValue('BD1_SELECTED', 'X');
    TobArtDde.Detail[RowArt].PutValue('BDP_SELECTIONNE', 'X');
    TOB(TobArtDde.Detail[RowArt].Data).PutValue('BDP_SELECTIONNE', 'X');
    TobArtDde.Detail[RowArt].PutValue('BDP_TRAITE', 'X');
    TOB(TobArtDde.Detail[RowArt].Data).PutValue('BDP_TRAITE', 'X');
  end;

  GrilleArt.InvalidateRow(RowArt + 1);
  GrilleFrs.Refresh;

end;

procedure TOF_BTVALIDEDEMPRIX.GrilleFrsGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

  if (not assigned(TobFrsLig)) Or (TobFrsLig.detail.count = 0) then exit;

  if ACol < GrilleFrs.FixedCols then Exit;
  IF Arow < GrilleFrs.FixedRows then exit;
  //
  if TobFrsLig.Detail[Arow-1].GetBoolean('BD1_SELECTED') then
  BEGIN
    canvas.Brush.Color:= clGray;
    canvas.Font.Color := clYellow;
  END;


end;

Procedure TOF_BTVALIDEDEMPRIX.ChercheOnClick(Sender : Tobject);
begin
  ChargeArtDemPrix;
end;

Procedure TOF_BTVALIDEDEMPRIX.ChargeArtDemPrix;
begin

  ConditionTobArtDde;

  ChargementGrilleWithTob(TobArtDde, GrilleArt, LesColArticle);

  GrilleArtEnter (self);
  GrilleArt.Invalidate;

end;

Procedure TOF_BTVALIDEDEMPRIX.ConditionTobArtDde;
var TOBA       : TOB;
    TOBB       : TOB;
    iInd  : Integer;
begin

  if assigned(TobArtDde) then TobArtDde.ClearDetail;

	for iInd := 0 to TobGestion.detail.count -1 do
  begin
		TOBA := TobGestion.detail[iInd];
    if EnregInSelect (TOBA) then
    begin
      TOBB := TOB.Create ('ARTICLEDEMPRIX',TobArtDde,-1);
      TOBB.Dupliquer(TOBA,False,true);
      TOBB.data := TOBA;
    end;
  end;

  TobArtDde.Detail.Sort('GL_TYPEARTICLE,GL_FAMILLENIV1,GL_FAMILLENIV2,GL_FAMILLENIV3,BDP_ARTICLE');

end;

function TOF_BTVALIDEDEMPRIX.EnregInSelect(TOBA: TOB): boolean;
begin

  result := true;

  if TOBA.GetString('BDP_UNIQUE') <> unique.Value then
  begin
    Result := False;
    Exit;
  end;

  if (TYPEARTICLE.value <> '') and (TOBA.GetString('GA_TYPEARTICLE')<>TYPEARTICLE.Value) then
  begin
    result := false;
    exit;
  end;

  if (FAMILLENIV1.value <> '') and (TOBA.GetString('GA_FAMILLENIV1')<>FAMILLENIV1.Value) then
  begin
    result := false;
    exit;
  end;

  if (FAMILLENIV2.value <> '') and (TOBA.GetString('GA_FAMILLENIV2')<>FAMILLENIV2.Value) then
  begin
    result := false;
    exit;
  end;

  if (FAMILLENIV3.value <> '') and (TOBA.GetString('GA_FAMILLENIV3')<>FAMILLENIV3.Value) then
  begin
    result := false;
    exit;
  end;

end;

Procedure TOF_BTVALIDEDEMPRIX.ValideOnClick(Sender : Tobject);
begin

  if FfromMem then
    TOBParam.setString('RETOUR','X')
  else
    OnUpdate;

end;

procedure TOF_BTVALIDEDEMPRIX.OnDeleteRecord(Sender : Tobject);
var TobDelete : TOB;
    TobA      : TOB;
    TobD      : TOB;
    NumUnique : Integer;
    i         : Integer;
    CodeArt   : String;
begin
  inherited;

  NumUnique := StrToInt(Unique.value);
  CodeArt   := TobArtDde.Detail[GrilleArt.row-1].GetString('BDP_ARTICLE');

  //lecture entete demande de prix
  if (PGIAsk(TexteMessage[1], ecran.Caption) = mrYes) then
  begin
    TobDemPrx.PutValue('MODIFIED','X');
    //désaffectation de la ligne article demande de prix
    Tobdelete := TOBgestion.FindFirst(['BDP_UNIQUE', 'BDP_ARTICLE'],[NumUnique, CodeArt], True);
    if TobDelete <> nil then
    begin
      //désafectation du fournisseur demande de prix...
      TobA := FindFournDemPrix(TobDelete,TobFrsArt);
      TOBA.SetInteger('UNIQUE', 0);
      For i:=0 to TobA.detail.count -1 do
      begin
        TOBA.Detail[i].SetInteger('BD1_UNIQUE', 0);
      end;
      //désafectation des détail de demande de prix...
      TOBD := FindFournDemPrix(TobDelete, TobDetailDemPx);
      TOBD.SetInteger('UNIQUE', 0);
      For i:=0 to TobD.detail.count -1 do
      begin
        TOBD.Detail[i].SetInteger('BD0_UNIQUE', 0);
      end;
      Tobdelete.SetInteger('BDP_UNIQUE', 0);
      //suppression de la ligne de la tob d'affichage de la grille
      TobDelete := TobArtDde.Detail[GrilleArt.row-1];
      if TobDelete <> nil then TobDelete.free;
    end;

    ChargeArtDemPrix;
  end;

end;


procedure TOF_BTVALIDEDEMPRIX.InserelesFournisseurs (TOBDetailFour,TOBDemandeArticle : TOB);
var indice : Integer;
		TOBBF : TOB;
begin

  TobFrsLig := FindFournDemPrix(TOBDemandeArticle,TobFrsArt);

  if TobFrsLig = nil then
  begin
    TobFrsLig := TOB.Create ('UN DETAIL DE LIGNE',TobFrsArt,-1);
    TobFrsLig.AddChampSupValeur ('UNIQUE',TOBDemandeArticle.GetInteger('BDP_UNIQUE'));
    TobFrsLig.AddChampSupValeur ('UNIQUELIG',TOBDemandeArticle.GetInteger('BDP_UNIQUELIG'));
  end;

	for Indice := 0  to TOBDetailFour.Detail.count -1 do
  begin
    IF TobFrsLig.FindFirst(['BD1_TIERS'],[TOBDetailFour.Detail[Indice].GetString('T_TIERS')],true) = nil then
    begin
      TOBBF := TOB.Create ('FOURLIGDEMPRIX',TobFrsLig,-1);
      TOBBF.SetDouble  ('BD1_QTEACH',0);
      TOBBF.SetString  ('BD1_TIERS',TOBDetailFour.Detail[Indice].GetString('T_TIERS'));
      TOBBF.SetString  ('BD1_NATUREPIECEG',TOBDemandeArticle.GetString('BDP_NATUREPIECEG'));
      TOBBF.SetString  ('BD1_SOUCHE',TOBDemandeArticle.GetString('BDP_SOUCHE'));
      TOBBF.SetInteger ('BD1_NUMERO',TOBDemandeArticle.GetInteger('BDP_NUMERO'));
      TOBBF.SetInteger ('BD1_INDICEG',TOBDemandeArticle.GetInteger('BDP_INDICEG'));
      TOBBF.SetInteger ('BD1_UNIQUE',TOBDemandeArticle.GetInteger('BDP_UNIQUE'));
      TOBBF.SetInteger ('BD1_UNIQUELIG',TOBDemandeArticle.GetInteger('BDP_UNIQUELIG'));
      TOBBF.SetString  ('BD1_REFERENCE',TOBDemandeArticle.GetString('BDP_ARTICLE'));
      TOBBF.SetString  ('BD1_QTEACH',TOBDemandeArticle.GetString('BDP_QTEBESOIN'));
      TOBBF.SetString  ('BD1_QUALIFUNITEACH',TOBDemandeArticle.GetString('BDP_QUALIFUNITEVTE'));
    end;
  end;
  if TobFrsLig.detail.count = 0 then TobFrsLig.free;
  ChargementGrilleWithTob(TobFrsLig, GrilleFrs, LesColFourn);
	GrilleFrsEnter (self);
end;

procedure TOF_BTVALIDEDEMPRIX.AjoutFourClick(Sender: TObject);
var TOBRecupFou : TOB;
begin
  if TobArtDde.detail.Count = 0 then Exit;
  TOBRecupFou := TOB.Create ('LES FOURNISSEURS',nil,-1);
  TRY
    TheTOB := TOBrecupFou;
    AGLLanceFiche('BTP', 'BTFOURNMULTI_MUL', 'T_NATUREAUXI=FOU', '', 'SELECTION');
    if TheTOB = nil then exit;
    if TheTOB.detail.count = 0 then exit;
    InserelesFournisseurs (TheTOB,TobArtDde.Detail[GrilleArt.row-1]);
    TobDemPrx.PutValue('MODIFIED','X');
  FINALLY
    TOBRecupFou.Free;
  end;
end;

procedure TOF_BTVALIDEDEMPRIX.SupprimeFourClick(sender: TObject);
begin

	if TobFrsLig.Detail.count = 0 Then Exit;
  if (PGIAsk(TexteMessage[2], ecran.Caption) = mrYes) then
  begin
		TobFrsLig.Detail[GrilleFrs.Row-1].Free;
    TobDemPrx.PutValue('MODIFIED','X');
  end;

  if TobFrsLig.detail.count = 0 then FreeAndNil(TobFrsLig);

  ChargementGrilleWithTob(TobFrsLig, GrilleFrs, LesColFourn);

	GrilleFrsEnter (self);

end;

procedure TOF_BTVALIDEDEMPRIX.GrilleFrsCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  if not Cancel then LastValue := GrilleFrs.Cells[GrilleFrs.Col, GrilleFrs.Row];
end;

procedure TOF_BTVALIDEDEMPRIX.GrilleFrsCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  if (GrilleFrs.Cells[ACol, ARow] = LastValue) then Exit;
  if Acol = ColPa then
  begin
    TobFrsLig.Detail[ARow-1].SetDouble('BD1_PRIXACH', VALEUR(GrilleFrs.cells[Acol,Arow]));
  end else if Acol = ColNbjour then
  begin
    TobFrsLig.Detail[ARow-1].SetInteger('BD1_NBJOUR', StrToInt(GrilleFrs.cells[Acol,Arow]));
  end;
  TobFrsLig.detail[ARow-1].PutLigneGrid (GrilleFrs, Arow,false, false, LesColFourn);

end;


procedure TOF_BTVALIDEDEMPRIX.ZoneSuivanteOuOk(var ACol, ARow: Integer; var Cancel: boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
  RowFirst : integer;
begin
	RowFirst := ARow;
  OldEna := GrilleFrs.SynEnabled;
  GrilleFrs.SynEnabled := False;
  Sens := -1;
  ChgLig := (GrilleFrs.Row <> ARow);
  ChgSens := False;
  if GrilleFrs.Row > ARow then Sens := 1 else if ((GrilleFrs.Row = ARow) and (ACol <= GrilleFrs.Col)) then Sens := 1;
  ACol := GrilleFrs.Col;
  ARow := GrilleFrs.Row;
  ii := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      Lim := GrilleFrs.RowCount - 1;
      // ---
      if ((ACol = GrilleFrs.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then Break else
        begin
          Sens := -1;
          Continue;
          ChgSens := True;
        end;
      end;
      if ChgLig then
      begin
        ACol := GrilleFrs.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GrilleFrs.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GrilleFrs.FixedCols;
      end;
    end else
    begin
      if ((ACol = GrilleFrs.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := GrilleFrs.ColCount;
        ChgLig := False;
      end;
      if ACol > GrilleFrs.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GrilleFrs.ColCount - 1;
      end;
    end;
  end;
  GrilleFrs.SynEnabled := OldEna;
end;

function TOF_BTVALIDEDEMPRIX.ZoneAccessible(ACol, ARow: Integer): boolean;
begin
  Result := false;
  if (ACol = ColNbJour) or (Acol = ColPA) then Result := True;
end;


procedure TOF_BTVALIDEDEMPRIX.GrilleFrsExit(Sender: TObject);
var Acol,Arow : Integer;
		CC : Boolean;
begin
  Acol := GrilleFrs.col;
  Arow := GrilleFrs.row;
	GrilleFrsCellExit(Self,Acol,Arow,CC);
end;

Initialization
  registerclasses ( [ TOF_BTVALIDEDEMPRIX] ) ;
end.

