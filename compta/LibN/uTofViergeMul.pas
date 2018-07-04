{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/06/2002
Modifié le ... : 20/02/2004
Description .. : Source TOF de la FICHE CPCONSECR, CPCONSGENE, CPPOINTAGEECR
Suite ........ : GCO - 29/05/2004 - FQ 17871
Suite ........ : Modification pour la colonne "POURCENTAGE", à savoir qu'on ne
Suite ........ : peut trier correctement dessus, étant donné que c'est un affichage
Suite ........ : de type "chaine". alors on doit définir un champ XPOURCENTAGE
Suite ........ : dans la TOF Fille, du coup, l'ancetre fera le tri sur ce champ
Suite ........ : et non pas sur celui que l'utilisateur pense avoir cliqué.
Mots clefs ... : TOF;
*****************************************************************}
unit uTofViergeMul;

interface

uses StdCtrls,
  Controls,
  Classes,
  Graphics,
  extctrls,
{$IFDEF VER150}
  Variants,
{$ENDIF}
{$IFDEF EAGLCLIENT}
//  UtileAGL, // PrintDBGrid
{$ELSE}
  db,
  dbGrids,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  UTOF,
  uTob,
  Grids,        // TGridDrawState
  Menus,        // PopUpMenu
  Htb97,        // ToolBarButton97
  HSysMenu,     // HSystemMenu
  Dialogs,      // TFindDialog
  HXlspas,      // ExportGrid
  uLibWindows,  // NotifyErrorComponent ou TMulPanelCumul
  uObjFiltres,  // TObjFiltre -> Objet de gestion des filtres
  CritEdt,      // ClassCritEdt
  HRichOle;     // THRichEditOle


type

 _RecChamps = record
  Nom   : string ;
  Index : integer ;
 end ;

 _TChamps = array of _RecChamps ;

  TOF_ViergeMul = class(TOF)
    FListe        : THGrid;
    Z_SQL         : THSQLMemo;

    FPanelCumul   : TPanel;
    FPanelCumulSelect : TPanel;

    PageControl   : TPageControl;
    PopF11        : TPopUpMenu;
    FFiltres      : THValComboBox;
    AFindDialog   : TFindDialog;
    ASaveDialog   : TSaveDialog;
    BCherche      : TToolBarButton97;
    BCherche_     : TToolBarButton97;
    BAgrandir     : TToolBarButton97;
    BReduire      : TToolBarButton97;
    BSelectAll    : TToolBarButton97;
    BValider      : TToolBarButton97;
    BRechercher   : TToolBarButton97;
    BExport       : TToolBarButton97;
    BImprimer     : TToolBarButton97;
    BEffaceAvance : TToolBarButton97;
    BParamListe   : TToolBarButton97;
    BFiltre       : TToolBarButton97;

{$IFDEF EAGLCLIENT}
    BFetchGauche  : TToolBarButton97;
    BFetchDroit   : TToolBarButton97;
{$ENDIF}

    BBlocNote     : TToolBarButton97;
    FBlocNote     : THRichEditOle;
    HPB           : TToolWindow97;

    Z_C1          : THValComboBox;
    Z_C2          : THValComboBox;
    Z_C3          : THValComboBox;

    procedure OnNew;                 override;
    procedure OnDelete;              override;
    procedure OnUpdate;              override;
    procedure OnLoad;                override;
    procedure OnArgument(S: string); override;
    procedure OnClose;               override;

    // Evenements pour les TToolBarButton97
    procedure OnClickBSelectAll    (Sender : TObject); virtual;
    procedure OnClickBValider      (Sender : TObject); virtual;
    procedure TraitementBValider   (Sender : TObject); virtual;
    procedure OnClickBCherche      (Sender : TObject); virtual;
    procedure OnClickBAgrandir     (Sender : Tobject); virtual;
    procedure OnClickBReduire      (Sender : Tobject); virtual;
    procedure OnClickBRechercher   (Sender : Tobject); virtual;
    procedure OnClickBExport       (Sender : Tobject); virtual;
    procedure OnClickBImprimer     (Sender : Tobject); virtual;
    procedure OnClickBEffaceAvance (Sender : TObject); virtual;
    procedure OnClickBParamListe   (Sender : TObject); virtual;
    procedure OnClickBBlocNote     (Sender : TObject); virtual;

    procedure AfterShow;                               virtual;
    procedure OnBeforeFlipFListe   (Sender: TObject; ARow: Integer; var Cancel: Boolean); virtual;
    procedure OnFlipSelectionFListe(Sender : TObject); virtual;

    // Evénements pour le AFindDialog
    procedure OnFindAFindDialog    (Sender : TObject);
    // Evénements pour le PopF11
    procedure OnPopUpPopF11        (Sender : TObject); virtual ; 
    // Evénements pour la fiche
    procedure OnKeyDownEcran       (Sender : TObject; var Key : Word; Shift : TShiftState); virtual;
    procedure OnResizeEcran        (Sender : TObject);

    // Evénements de la Grille
    procedure OnKeyDownFListe      (Sender : TObject; var Key : Word; Shift : TShiftState ); virtual ;
    procedure OnMouseDownFListe    (Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure OnRowEnterFListe     (Sender : TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean); virtual;
    procedure OnSortedFListe       (Sender : Tobject);
    procedure OnPostDrawCellFListe (ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState); virtual;
    //procedure OnColumnMovedFListe  (Sender: TObject; FromIndex, ToIndex: Integer);
    procedure OnColumnWidthsChangedFListe (Sender : TObject);

    procedure RefreshPclPge;

  private
    FObjFiltre      : TObjFiltre;
    FControlFiltre  : TControlFiltre;

    FTobFListe      : Tob;
    FTobRow         : Tob;

    FFindFirst          : Boolean;
    FFocusFListe        : Boolean;        // Prise de Focus autoriser pour FLISTE sur le BCherche.Click
    FFiltreDisabled     : Boolean;        // Désactivation des FILTRES
    FBoSansFiltreDefaut : Boolean ;       // Autorisation de charger le filtre par défaut
    FBoMultiSelected    : Boolean;        // Activation du MultiSelect
    FAutoSearch         : TAutoSearch;
    FLoading            : Boolean;        // Indicateur pour en cours de chargement
    FCritModified       : Boolean;        // Indicateur de modifiaction dans le PAGECONTROL
    FSearchTimer        : TTimer;
    FMulPanelFiltre     : TMulPanelCumul; // Panel des CUMULS de la LISTE
    FCritEdt            : ClassCritEdt;   //

    FStTriTobFListe : string;         // Champ sur lequel il faut trier la TOBFListe
    FStSqlTobFListe : string;         // Text SQL avec lequel on doit charger TOBFLISTE
    FStSqlWhereTobFListe : string;    // Condition WHERE de la requête SQL de l'écran
    FNewSources     : string;
    FNewLiaison     : string;
    FStNewTris      : string;
    FNewParams      : string;
    FStNewTitres    : hstring;
    FNewJustifs     : string;
    FNewLibelle     : hstring;
    FNewNumCols     : hstring;
    FNewPerso       : string;

    FNewOkTri       : Boolean;
    FNewOkNumCol    : Boolean;
    FBoOkRefresh    : Boolean;

    FQuery          : TQuery;

    FChamps         : _TChamps ;

    FOldKeyDown    : TKeyEvent ;
    FboFetchLesTous: boolean ;

{$IFDEF EAGLCLIENT}
    procedure OnClickBFetchGauche     ( Sender : TObject );
    procedure OnClickBFetchDroit      ( Sender : TObject );
{$ENDIF}

    procedure VerifCritere;
    procedure SearchTimerTimer (Sender: TObject);
    procedure ResetTimer       (Sender: TObject);
    procedure CritereChange    (Sender: TObject);
    procedure SetCritModified  (Value: Boolean);
    procedure UpdateGrille;
    procedure LoadCumulControls;
    procedure UnloadCumulControls;
    {$IFNDEF JOHN}
    procedure FaitOpenSql( vBoFetch : Boolean = False ; vboFetchLesTous : boolean = False);
    {$ENDIF JOHN}
    procedure TraiteStSqlTobFListe( vValue : string );
    procedure SelectionneTout( vBoSelected : Boolean );
    procedure SelectionneLigne;

{$IFDEF EAGLCLIENT}
    function FetchLesTous : boolean;
{$ENDIF}

  protected
    procedure RemplitATobFListe;                           virtual; abstract;
    procedure InitControl;                                 virtual; abstract;
    function  AjouteATobFListe    ( vTob : Tob) : Boolean; virtual;
    procedure CalculPourAffichage ( vTob : Tob);           virtual;
    procedure RefreshFListe( vBoFetch : Boolean );         virtual;

    function  BeforeLoad : boolean;                        virtual;
    function  AfterLoad  : boolean;                        virtual;
    {$IFDEF JOHN}
    procedure FaitOpenSql( vBoFetch : Boolean = False ; vboFetchLesTous : boolean = False); virtual;
    {$ENDIF JOHN}

    procedure AvantChangementFiltre;                   virtual;
    procedure ApresChangementFiltre;                   virtual;
    procedure NouveauFiltre;                           virtual;
    procedure SupprimeFiltre;                          virtual;

  public

    FStListeParam   : string;
    FFI_TABLE       : string;            // Champ FI_TABLE de la Table Filtres
    FStListeChamps  : string;            // Liste des champs de la LISTE PARAMETRABLE
    FStNewLargeurs  : string;            // Largeur des colonnes dans la liste paramètrable
    FMaxRow         : integer;           // Nombre Max de lignes dans la grille pour le CWAS
    FBoOkAncetreFaitAffichage : Boolean; // Détermine si utofViergeMul doit faire
                                         // l'affichage dans la grille ou si la
                                         // fiche qui herite désire en faire un
                                         // particulier.

    procedure InitAutoSearch;
    procedure TraitementListeParam;
    procedure RemplitListeChamps ( vTOB : TOB ) ;
    procedure TOBVersTHGrid( vTOB : TOB ) ;
    function  GetCumulChamps ( vStChamps : string ) : double ;
    property ATobFListe           : Tob          read FTobFListe           write FTobFListe;
    property ATobRow              : Tob          read FTobRow              write FTobRow;
    property AFocusFliste         : Boolean      read FFocusFliste         write FFocusFListe;
    property AFiltreDisabled      : Boolean      read FFiltreDisabled      write FFiltreDisabled;
    property ABoSansFiltreDefaut  : Boolean      read FBoSansFiltreDefaut  write FBoSansFiltreDefaut;
    property CritModified         : Boolean      read FCritModified        write SetCritModified;
    property AutoSearch           : TAutoSearch  read FAutoSearch          write FAutoSearch;
    property AStTriTobFListe      : string       read FStTriTobFListe      write FStTriTobFListe;
    property AStSqlTobFListe      : string       read FStSqlTobFListe      write TraiteStSqlTobFListe;
    property AStSqlwhereTobFListe : string       read FStSqlWhereTobFliste ;
    property ACritEdt             : ClassCritEdt read FCritEdt             write FCritEdt;
    property ABoMultiSelected     : Boolean      read FBoMultiSelected     write FBoMultiSelected;
    property AboFetchLesTous      : boolean      read FboFetchLesTous      write FboFetchLesTous default false ;
  end;

function  CSupprimeFauxChamps(vStChaineSql: string): string;
function  CSqlTextFromList(vStListe: string): string;
procedure CAjouteItemPopUp(F: TForm; vPopF11: TPopUpMenu);
procedure ActivationMenuItem( vMenuItem : TMenuItem );

implementation

uses
 {$IFDEF MODENT1}
 CPProcGen,
 {$ENDIF MODENT1}
 uTobDebug,    // TobDebug
 Vierge,
 Ent1,
 SaisComm,
 ParamSoc,     // GetParamSocSecur
 SaiSUtil,     // pour le StrSO
 HMsgBox,
 Windows,      // VK
 UtilPGI,      // EstMultiSoc
 uObjEtats,    //  TObjEtats.GenereEtatGrille
 Filtre,       // Filtres
 Messages,     // WM_KeyDown
 uTXML,        // XMLDecodeSt
 ParamDBG;     // ParamList


procedure CAjouteItemPopUp(F: TForm; vPopF11: TPopUpMenu);
var i, j, k     : integer;
    lPopUp      : TPopUpMenu;
    lMenuItem   : THMenuItem;
    lMenuItem2  : THMenuItem;
    lSMenuItem  : THMenuItem;
    lBoSupprimeEditions : Boolean;
begin
  lMenuItem2 := nil;

  if (vPopF11 = nil) or (CtxStandard in V_Pgi.PgiContexte) then Exit;

  {Purge des précédents items}
  PurgePopup(vPopF11);
  for i := 0 to F.ComponentCount - 1 do
  begin
    if ((F.Components[i]) is TPopUpMenu) and (F.Components[i].Tag > 0) then
    begin
      lPopUp := TPopUpMenu(F.Components[i]);

      if (vPopF11.Items.Count <> 0) then
      begin
        lMenuItem := THMenuItem.Create(F);
        lMenuItem.Caption := '-';
        vPopF11.Items.Add(lMenuItem);
      end;

      if lPopUp.Name = 'POPUPEDITION' then
      begin
        lMenuItem2 := THMenuItem.Create(F);
        lMenuItem2.Caption := '&Editions';
        vPopF11.Items.Add(lMenuItem2);
      end;

      // Parcous le PopUpmenu et on récupère tous les items Enabled := True
      for j := 0 to lPopUp.Items.Count - 1 do
      begin
        if lPopUp.Items[j] is THMenuItem then
        begin
          if (lPopUp.Items[j].Enabled) and (lPopUp.Items[j].Visible) then
          begin
            // Ajout de l' item dans le lPopUpF11
            lMenuItem := THMenuItem.Create(F);
            lMenuItem.Caption  := lPopUp.Items[j].Caption;
            lMenuItem.ShortCut := lPopUp.Items[j].ShortCut;
            lMenuItem.OnClick  := lPopUp.Items[j].OnClick;
            lMenuItem.Checked  := lPopUp.Items[j].Checked;

            if lPopUp.Name = 'POPUPEDITION' then
              lMenuItem2.Add(lMenuItem)
            else
              vPopF11.Items.add(lMenuItem);

            if lPopUp.Items[j].Count > 0 then
            begin
              // GCO - 04/01/2007 - Ajout des SOUS MENU
              for k := 0 to lPopUp.Items[j].Count - 1 do
              begin
                if (lPopUp.Items[j].Items[k].Enabled) and
                   (lPopUp.Items[j].Items[k].Visible) then
                begin
                  lSMenuItem := THMenuItem.Create(F);
                  lSMenuItem.Caption  := lPopUp.Items[j].Items[k].Caption;
                  lSMenuItem.ShortCut := lPopUp.Items[j].Items[k].ShortCut;
                  lSMenuItem.OnClick  := lPopUp.Items[j].Items[k].OnClick;
                  lSMenuItem.Checked  := lPopUp.Items[j].Items[k].Checked;
                  lMenuItem.Add(lSMenuItem);
                end;
              end;

              if lMenuItem.Count = 0 then
                FreeAndNil(lMenuItem);
           end;
          end;
        end;
      end;

      if lPopUp.Name = 'POPUPEDITION' then
      begin
        if (lMenuItem2 <> nil) then
        begin
          lBoSupprimeEditions := True;
          for j := 0 to lMenuItem2.Count -1 do
          begin
            if lMenuItem2.Items[j].Caption <> '-' then
              lBoSupprimeEditions := False;
          end;

          if lBoSupprimeEditions then
            FreeAndNil(lMenuItem2);
        end;
      end;

    end;
  end;
  ActivateXpPopUp( vPopF11 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure ActivationMenuItem( vMenuItem : TMenuItem );
var i : integer;
    lBoActive : Boolean;
begin
  lBoActive := False;
  for i := 0 to vMenuItem.Count - 1 do
  begin
    if vMenuItem.Items[i].isLine then
    begin
      vMenuItem.Items[i].Enabled := True;
    end
    else
    begin
      vMenuItem.Items[i].Visible := vMenuItem.Items[i].Enabled;
      if vMenuItem.Items[i].Enabled then
      begin
        if lBoActive = False then
          lBoActive := True;
      end;
    end;  
  end;
  vMenuItem.Visible := ((vMenuItem.Count = 0) and (vMenuItem.Enabled)) or (lBoActive);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 20/01/2005
Description .. : Supprime de la requete les champs dont le prefixe de table
Suite ........ : n'existe pas, donc des champs calculés que l'on traite par
Suite ........ : code. Exemple : SOLDE pour le solde progressif à la ligne
Mots clefs ... :
*****************************************************************}
function CSupprimeFauxChamps(vStChaineSql: string): string;
var lStChamp   : string;
    lStPrefixe : string;
begin
  Result := '';
  while vStChaineSql <> '' do
  begin
    lStChamp   := READTOKENST( vStChaineSql );
    lStPrefixe := ExtractPrefixe( lStChamp );
    if PrefixeToNum( lStPrefixe ) <> 0 then
      Result := Result + lStChamp + ';';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
function CSqlTextFromList(vStListe: string): string;
begin
  Result := CSupprimeFauxChamps( vStListe );
  Result := FindEtReplace(Result, ';', ',', True);
  Result := Copy(result, 0, Length(Result) - 1);
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ViergeMul.OnNew;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ViergeMul.OnDelete;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ViergeMul.OnUpdate;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2003
Modifié le ... : 25/01/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnArgument(S: string);
begin
  FSearchTimer := TTimer.Create(Ecran);
  FListe := THGrid(GetControl('FLISTE', True));
  FPanelCumul := TPanel(GetControl('PCUMUL', True));

  FPanelCumulSelect := TPanel(GetControl('PCUMULSELECT', False));

  FCritEdt := ClassCritEdt.Create;

  AFindDialog := TFindDialog.Create(Ecran);
  AFindDialog.OnFind := OnFindAFindDialog;
  
  ASaveDialog := TSaveDialog.Create(Ecran);
  ASaveDialog.Filter := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
  ASaveDialog.DefaultExt := 'XLS';
  ASaveDialog.FilterIndex := 1;
  ASaveDialog.Options := ASaveDialog.Options + [ofOverwritePrompt, ofPathMustExist, ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];

  // Création de ATobFListe en Mode 2/3 uniquement
  ATobFListe := Tob.Create('', nil, -1);

  FMulPanelFiltre := TMulPanelCumul.Create;
  // Gestion du filtre
  PageControl := TPageControl(GetControl('PAGECONTROL', True));
  FFiltres := THValComboBox(GetControl('FFILTRES', True));

  // Gestion des Filtres avec l'objet de JP
  FControlFiltre.Filtre   := TToolBarButton97(GetControl('BFILTRE', True));
  FControlFiltre.Filtres  := FFiltres;
  FControlFiltre.PageCtrl := PageControl;
  FObjFiltre := TObjFiltre.Create(FControlFiltre, FFI_TABLE, False);

  FObjFiltre.AvantChangementFiltre := AvantChangementFiltre;
  FObjFiltre.ApresChangementFiltre := ApresChangementFiltre;
  FObjFiltre.NouveauFiltre         := NouveauFiltre;
  FObjFiltre.SupprimeFiltre        := SupprimeFiltre;

  // Les TToolBarButton97
  BCherche      := TToolBarButton97(GetControl('BCherche', True));
  BCherche_     := TToolBarButton97(GetControl('BCherche_', True));
  BAgrandir     := TToolBarButton97(GetControl('BAgrandir', True));
  BReduire      := TToolBarButton97(GetControl('BReduire', True));
  BSelectAll    := TToolBarButton97(GetControl('BSelectAll', True));
  BValider      := TToolBarButton97(GetControl('BValider', True));
  BRechercher   := TToolBarButton97(GetControl('BRechercher', True));
  BExport       := TToolBarButton97(GetControl('BExport', True));
  BImprimer     := TToolBarButton97(GetControl('BImprimer', True));
  BEffaceAvance := TToolBarButton97(GetControl('BEffaceAvance', True));
  BParamListe   := TToolBarButton97(GetControl('BParamListe', True));
  BFiltre       := TToolBarButton97(GetControl('BFILTRE', True));
  BBlocNote     := TToolBarButton97(GetControl('BBLOCNOTE', True));

  // Création de la fenêtre BLOC NOTE
  HPB := TToolWindow97.Create(nil);
  HPB.Parent  := Ecran;
  HPB.Caption := 'Bloc note';
  HPB.Visible := False;
  HPB.Left    := 350;
  HPB.Top     := 260;
  HPB.Height  := 140;
  HPB.width   := 250;

  FBlocNote := THRichEditOle.Create(nil);
  FBlocNote.Parent := HPB;
  FBlocNote.Align := AlClient;
  FBlocNote.ReadOnly := True;
  FBlocNote.Clear;

  // Composants de l'onglet Avancés
  Z_C1          := THValComboBox(GetControl('Z_C1', True));
  Z_C2          := THValComboBox(GetControl('Z_C2', True));
  Z_C3          := THValComboBox(GetControl('Z_C3', True));

  BAgrandir.OnClick     := OnClickBAgrandir;
  BAgrandir.Visible     := True;
  BReduire.Visible      := False;
  BReduire.OnClick      := OnClickBReduire;
  BSelectAll.OnClick    := OnClickBSelectAll;
  BValider.OnClick      := OnClickBValider;
  BCherche.OnClick      := OnClickBCherche;
  BCherche_.OnClick     := OnClickBCherche;
  BRechercher.OnClick   := OnClickBRechercher;
  BExport.OnClick       := OnClickBExport;
  BImprimer.OnClick     := OnClickBImprimer;
  BEffaceAvance.OnClick := OnClickBEffaceAvance;
  BParamListe.OnClick   := OnClickBParamListe;
  BBlocNote.OnClick     := OnClickBBlocNote;


  PopF11 := TPopUpMenu(GetControl('POPF11', True));
  PopF11.OnPopup := OnPopUpPopF11;

  FOldKeyDown           := Ecran.OnKeyDown ;
  Ecran.OnKeyDown       := OnKeyDownEcran;
  Ecran.OnResize        := OnResizeEcran;

  VideFiltre(FFiltres, PageControl);

  // Init des variables globales
  AFiltreDisabled  := False; // Désactivation des filtres
  AFocusFListe     := True;  // Blocage de Focus de la grille
  ABoMultiSelected := False; // Activation du MultiSelect

  PageControl.ActivePageIndex := 0;
  FSearchTimer.Enabled := False;
  FSearchTimer.OnTimer := SearchTimerTimer;

  // GCO - 13/07/2005
  FListe.EnabledBlueButton := True;

  // Evénements de la grille
  FListe.ListeParam      := FStListeParam;
  FListe.OnKeyDown       := OnKeyDownFListe;
  FListe.OnMouseDown     := OnMouseDownFListe;
  FListe.OnRowEnter      := OnRowEnterFListe;
  FListe.OnSorted        := OnSortedFListe;
  FListe.PostDrawCell    := OnPostDrawCellFListe;
  FListe.OnBeforeFlip    := OnBeforeFlipFListe;
  FListe.OnFlipSelection := OnFlipSelectionFListe;
  FListe.OnColumnWidthsChanged := OnColumnWidthsChangedFListe;

  // GCO - 30/05/2006 - FQ 17399
  TraitementListeParam;

  FMulPanelFiltre.Grid             := FListe;
  FMulPanelFiltre.PanelCumul       := FPanelCumul;
  {$IFDEF EAGLCLIENT}
  FMulPanelFiltre.FetchLesTous     := FboFetchLesTous ;
  {$ENDIF}
  
  if Assigned(FPanelCumulSelect) then
    FMulPanelFiltre.PanelCumulSelect := FPanelCumulSelect;

  FMulPanelFiltre.TobListe         := FTobFListe;
  FMulPanelFiltre.InitializeField;
  FMulPanelFiltre.InitializeControl;

  ResizeOnglets(PageControl);

  BParamListe.Visible := ExJaiLeDroitConcept(ccParamListe, False);
  BExport.Visible     := ExJaiLeDroitConcept(ccExportListe, False);

  TFVierge(Ecran).OnAfterFormShow := AfterShow;

{$IFDEF EAGLCLIENT}
  // Ajout GCO pour la gestion du FETCH
  BFetchGauche         := TToolBarButton97(GetControl('BFLETCHG', True));
  BFetchDroit          := TToolBarButton97(GetControl('BFLETCHD', True));
  BFetchGauche.Visible := True;
  BFetchDroit.Visible  := True;
  BFetchGauche.OnClick := OnClickBFetchGauche;
  BFetchDroit.OnClick  := OnClickBFetchDroit;
  FListe.SortEnabled   := False;
{$ELSE}
  FListe.SortEnabled   := True;
{$ENDIF}

  // Autorise ou Interdit à la fiche ancetre de faire l'affichage dans la grille
  // Si False alors il faut faire l'affichage soit même dans la fenêtre qui hérite
  FBoOkAncetreFaitAffichage := True;

  if CtxPCl in V_Pgi.PgiContexte then
    FBoOkRefresh := True
  else
    FBoOkRefresh := GetParamSocSecur('SO_BDEFGCPTE', False);

  // GCO - 21/04/2004
  if FStListeParam = '' then
    PgiInfo('Vous devez renseigner la variable globale FStListeParam.', 'Erreur de conception')
  else
    CUpdateGridListe ( FListe , FStListeParam ) ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. : Chargement de la fiche
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnLoad;
begin
  inherited;
  // On récupère l'ordre de tri désiré de la LISTE PARAMETRABLE
  AStTriTobFListe := FindEtReplace(FStNewTris, ',', ';', True);

  // Demande aux fiches héritées d' initialiser les contrôles
  InitControl;
  UpdateGrille;

  // Chargement des filtres
  if not AFiltreDisabled then
  begin
    // Chargement des Filtre spar L'AGL, on peut travailler par dessus
    FObjFiltre.Charger;

    // GCO - 13/11/2007 - Nouvelle recherche car on ne veut pas du filtre par
    // défaut à l'ouverture de l'écran
    if FBoSansFiltreDefaut then
      FObjFiltre.NouveauFiltre;
  end
  else
  begin
    FFiltres.Enabled := False;
    BFiltre.Enabled  := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.OnClose;
begin
  UnloadCumulControls;

  // GCO - 31/08/2006 - FQ 18628 - Violation d'acces si la liberation se fait
  // apres le AFindDialog
  FreeAndNil(FBlocNote);
  FreeAndNil(HPB);
  // FIN GCO

  FreeAndNil(FMulPanelFiltre);
  FreeAndNil(ASaveDialog);
  FreeAndNil(AFindDialog);
  FreeAndNil(FTobFListe);
  FreeAndNil(FSearchTimer);
  FreeAndNil(FObjFiltre);
  FreeAndNil(FCritEdt);

  if Assigned(FQuery) then Ferme(FQuery);

  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.OnClickBCherche(Sender: TObject);
var lRow: integer;
    lStNewLargeurs : string;
    lStNewTitres   : hstring;
    lStListeChamps : string;
    lCancel : Boolean;
begin
  // Pour forcer la validation du champ
  if AutoSearch = asMouette then
    if (not NextPrevControl(Ecran, True, TRUE)) then Exit;

  SourisSablier;
  FLoading := True;
  HGBeginUpdate(Fliste);

  try
    lRow := FListe.Row;

    // GCO - FQ 12693 Problème de Resize car modif pour éviter la bataille navale
    // On remet systématiquement la largeur des colonnes avant chaque Resize de Grille
    // afin d'éviter la disparation des plus petites colonnes
    lStNewLargeurs := FStNewLargeurs;
    lStNewTitres   := FStNewTitres;
    lStListeChamps := FStListeChamps;

    if FListe.Titres[0] <> lStListeChamps then
    begin
      // Mise à jour de la chaine des titres de la grille
      FListe.Titres[0] := FStListeChamps;

      // GCO - 13/06/2007 - FQ 20536 - code déplacer dans traitementlistepararam
      // Recalcul du nombre de colonnes à afficher
      //i := 0;
      //lStListeChamps := FStListeChamps;
      //while lStListechamps <> '' do
      //begin
      //  ReadTokenSt(lStListeChamps);
      //  Inc(i);
      //end;
      //FListe.ColCount := i + 1;
      FMulPanelFiltre.InitializeField;
      FMulPanelFiltre.InitializeControl;
    end;

    CUpdateGridListe(FListe, FStListeParam);

    VerifCritere;
    // BeforeLoad sert à bloquer l'éxécution de la requête SQL, c'est la fiche
    // héritière de uTofViergeMul qui choisit ( cf : uTofPointageEcr.pas )
    if BeforeLoad then
    begin
      RemplitATobFListe;
      if Sender = bExport then FaitOpenSql ( False, True)
      else FaitOpenSql(false,FboFetchLesTous) ;
    end;

    RefreshFListe( False );
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
    FMulPanelFiltre.UpdateCumul;
    if Assigned(FPanelCumulSelect) then
      FMulPanelFiltre.UpdateCumulSelect;

    if (ATobFliste.Detail.Count > 0) then
    begin
      if (FListe.CanFocus) and (AFocusFListe) then
        FListe.SetFocus;

      // Placement sur le numéro de ligne enregistré dans lRow
      // GCO - 10/05/2007 - FQ 20179
      if lRow > FListe.RowCount then
        lRow := 1;

      if lRow = FListe.RowCount then
        lRow := FListe.RowCount -1;

      FListe.Row := lRow;
    end;

    UpdateGrille;

    // GCO - 02/05/2007 - FQ 20061
    FListe.OnRowEnter( nil, FListe.Row, lCancel, False );
    if not AfterLoad then Exit;

  finally
    HGEndUpdate(Fliste);
    SourisNormale;
    FLoading := false;
    CritModified := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/11/2002
Modifié le ... :   /  /
Description .. : Met à Blanc tous les composants de l' onglet Avancés
Mots clefs ... :
*****************************************************************}

procedure TOF_ViergeMul.OnClickBEffaceAvance(Sender: TObject);
begin
  THValComboBox(GetControl('Z_C1')).ItemIndex := -1;
  THValComboBox(GetControl('ZO1')).ItemIndex := -1;
  TEdit(GetControl('ZV1')).Text := '';
  TComboBox(GetControl('ZG1')).ItemIndex := -1;

  THValComboBox(GetControl('Z_C2')).ItemIndex := -1;
  THValComboBox(GetControl('ZO2')).ItemIndex := -1;
  TEdit(GetControl('ZV2')).Text := '';
  TComboBox(GetControl('ZG2')).ItemIndex := -1;

  THValComboBox(GetControl('Z_C3')).ItemIndex := -1;
  THValComboBox(GetControl('ZO3')).ItemIndex := -1;
  TEdit(GetControl('ZV3')).Text := '';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/06/2002
Modifié le ... : 26/04/2007
Description .. : - LG - 22/01/2003 - FB 11811 - F5 lance le lookup des
Suite ........ : generaux. Utilisation d'une fct de ULibWindows
Suite ........ : - LG - 11/10/2005 - FB 16677 - Ajout du ctrl+P
Suite ........ : - LG - 26/04/2007 - on lance la routine de l'ancetre pour 
Suite ........ : toutes les keys que l'on ne gere pas ici vk_delete etc etc
Mots clefs ... : 
*****************************************************************}

procedure TOF_ViergeMul.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of

    VK_F5: if (Shift = []) and FListe.Focused then
           begin
             Key := 0;
             if assigned(FListe.OnDblClick) then
              FListe.OnDblClick(Self);
           end;

    VK_F9  : BCherche.Click;

    VK_F10 : if (Shift = []) then BValider.Click;

    VK_F11 : begin
               if CritModified then BCherche.Click;
               Key := 0 ;
               PopF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);
             end;

    VK_F12 : if FListe.Focused then
               PageControl.SetFocus
             else
               FListe.SetFocus;

   80 : if ssCtrl in Shift then
           begin
            key := 0;
            BImprimer.Click;
           end;

    {$IFDEF EAGLCLIENT}
    VK_NEXT : if  (Shift = []) and ((FListe.RowCount - FListe.Row) < FMaxRow) and (BFetchDroit.Enabled) then
              begin
                BFetchDroit.Click;
              end;
    {$ENDIF}

    // Ctrl + A
    65: if (FListe.Focused) and (Shift = [ssCtrl]) then
        BSelectAll.Click;

    //Ctrl + H
    70: if Shift = [ssCtrl] then
        BRechercher.Click;
  else
   FOldKeyDown( Sender , Key , Shift );
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.OnClickBValider(Sender: TObject);
begin
  TraitementBValider( Sender );
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.OnClickBAgrandir(Sender: Tobject);
begin
  BAgrandir.Visible := False;
  BReduire.Visible := True;
  PageControl.Visible := False;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.OnClickBReduire(Sender: Tobject);
begin
  BReduire.Visible := False;
  BAgrandir.Visible := True;
  PageControl.Visible := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 21/06/2006
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.OnClickBSelectAll(Sender: TObject);
begin
  if FListe.AllSelected then
  begin
    FListe.AllSelected := False;
    SelectionneTout(False);
  end
  else
  begin
    FListe.AllSelected := True;
    SelectionneTout(True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_ViergeMul.OnClickBRechercher(Sender: Tobject);
begin
  FFindFirst := True;
  AFindDialog.Execute;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.OnFindAFindDialog(Sender: TObject);
begin
  Rechercher(FListe, AFindDialog, FFindFirst);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.OnClickBExport(Sender: Tobject);
var lStHint: string;
begin
  if not ExJaiLeDroitConcept(ccExportListe, True) then Exit;
  {$IFDEF EAGLCLIENT}
  { FQ 16061 - CA - 22/06/2005 - en eAGL, il faut exporter TOUTES les enregistrements de la requête,
            pas uniquement ceux affichés dans la grille }
  if FetchLesTous then
  {$ENDIF}
    if ASaveDialog.Execute then
    begin
      if ASaveDialog.FilterIndex = 5 then //html
      begin
        lStHint := FListe.Hint;
        FListe.Hint := Ecran.Caption;
        ExportGrid(FListe, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
        FListe.Hint := lStHint;
      end
      else
        ExportGrid(FListe, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 19/05/2006
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.OnPopUpPopF11(Sender: TObject);
begin
  CAjouteItemPopUp(Ecran, popF11);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/11/2002
Modifié le ... : 16/02/2007
Description .. : Impression des enregistrements de la grille
Suite ........ : - LG - 16/02/2007 - utilise du ObjEtat pour imprimer
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.OnClickBImprimer(Sender: Tobject);
begin
  if ATobFListe.Detail.Count > 0 then
   TObjEtats.GenereEtatGrille (FListe,Ecran.Caption,False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. : Paramètrage de la liste
Mots clefs ... :
*****************************************************************}

procedure TOF_ViergeMul.OnClickBParamListe(Sender: TObject);
begin
  if V_Pgi.SAV then
  begin
    PgiInfo('Vous ne pouvez pas modifier la liste des champs en mode SAV.', Ecran.Caption);
    Exit;
  end;

{$IFDEF EAGLCLIENT}
  ParamListe(FStListeParam, nil, 'Personnalisation des listes');
{$ELSE}
  ParamListe(FStListeParam, nil, nil, 'Personnalisation des listes');
{$ENDIF}

  // GCO - 30/05/2006 - FQ 17399
  TraitementListeParam;

  FMulPanelFiltre.InitializeField;
  FMulPanelFiltre.InitializeControl;
  FListe.ListeParam := FStListeParam;

  BCherche.Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/10/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnClickBBlocNote(Sender: TObject);
begin
  //
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.VerifCritere;
var i: integer;
begin
  if not V_PGI.EgaliseFourchette then Exit;

  for i := 0 to Ecran.ComponentCount - 1 do
  begin
    if (Ecran.Components[i] is THCritMaskEdit) and
      (TControl(Ecran.Components[i]).Parent is TTabSheet) and
      (TControl(Ecran.Components[i]).Visible) and
        (TControl(Ecran.Components[i]).Enabled) then
    begin
      V_PGI.EgaliseOnEnter(THCritMaskEdit(Ecran.Components[i]));
    end;
  end; // for
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.InitAutoSearch;
var
  i: integer;
begin
  //exit;

  //if (V_Pgi.Zoomole) then
  //  BCherche.Click;

  if (V_PGI.AutoSearch) and (AutoSearch <> AsMouetteForce) then
    AutoSearch := asTimer;
  for i := 0 to Ecran.ComponentCount - 1 do
  begin
    if (Ecran.Components[i] is TControl) and
      (TControl(Ecran.Components[i]).Parent is TTabSheet) then
      if (Ecran.Components[i] is TControl) and
        TControl(Ecran.Components[i]).Visible and
        (TControl(Ecran.Components[i]).Enabled) then
      begin
        if (Ecran.Components[i] is THCritMaskEdit) and (not
          Assigned(THCritMaskEdit(Ecran.Components[i]).OnEnter)) then
          THCritMaskEdit(Ecran.Components[i]).OnEnter := V_PGI.EgaliseOnEnter;

        case AutoSearch of
          asChange:
            begin
              if (Ecran.Components[i] is TEdit) and not
                assigned(TEdit(Ecran.Components[i]).OnChange) then
                TEdit(Ecran.Components[i]).OnChange := SearchTimerTimer;
              if (Ecran.Components[i] is THValComboBox) and not
                assigned(THValComboBox(Ecran.Components[i]).OnClick) then
                THValComboBox(Ecran.Components[i]).OnClick := SearchTimerTimer;
              if (Ecran.Components[i] is TCheckBox) and not
                assigned(TCheckBox(Ecran.Components[i]).OnClick) then
                TCheckBox(Ecran.Components[i]).OnClick := SearchTimerTimer;
              if (Ecran.Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Ecran.Components[i]).OnChange) then
                THCritMaskEdit(Ecran.Components[i]).OnChange :=
                  SearchTimerTimer;
              if (Ecran.Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Ecran.Components[i]).OnClick) then
                THRadioGroup(Ecran.Components[i]).OnClick := SearchTimerTimer;
            end;
          asExit:
            begin
              if (Ecran.Components[i] is TEdit) and not
                assigned(TEdit(Ecran.Components[i]).onExit) then
                TEdit(Ecran.Components[i]).onExit := SearchTimerTimer;
              if (Ecran.Components[i] is THValComboBox) and not
                assigned(THValComboBox(Ecran.Components[i]).onExit) then
                THValComboBox(Ecran.Components[i]).onExit := SearchTimerTimer;
              if (Ecran.Components[i] is TCheckBox) and not
                assigned(TCheckBox(Ecran.Components[i]).onExit) then
                TCheckBox(Ecran.Components[i]).onExit := SearchTimerTimer;
              if (Ecran.Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Ecran.Components[i]).onExit) then
                THCritMaskEdit(Ecran.Components[i]).onExit := SearchTimerTimer;
              if (Ecran.Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Ecran.Components[i]).onExit) then
                THRadioGroup(Ecran.Components[i]).onExit := SearchTimerTimer;
            end;
          asTimer:
            begin
              if (Ecran.Components[i] is TEdit) and not
                assigned(TEdit(Ecran.Components[i]).OnChange) then
                TEdit(Ecran.Components[i]).OnChange := ResetTimer;
              if (Ecran.Components[i] is THValComboBox) and not
                assigned(THValComboBox(Ecran.Components[i]).OnClick) then
                THValComboBox(Ecran.Components[i]).OnClick := ResetTimer;
              if (Ecran.Components[i] is TCheckBox) and not
                assigned(TCheckBox(Ecran.Components[i]).OnClick) then
                TCheckBox(Ecran.Components[i]).OnClick := ResetTimer;
              if (Ecran.Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Ecran.Components[i]).OnChange) then
                THCritMaskEdit(Ecran.Components[i]).OnChange := ResetTimer;
              if (Ecran.Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Ecran.Components[i]).OnClick) then
                THRadioGroup(Ecran.Components[i]).OnClick := ResetTimer;
            end;
        else
          begin
            if (Ecran.Components[i] is TEdit) and not
              assigned(TEdit(Ecran.Components[i]).OnChange) then
              TEdit(Ecran.Components[i]).OnChange := CritereChange;
            if (Ecran.Components[i] is THValComboBox) and not
              assigned(THValComboBox(Ecran.Components[i]).OnClick) then
              THValComboBox(Ecran.Components[i]).OnClick := CritereChange;
            if (Ecran.Components[i] is TCheckBox) and not
              assigned(TCheckBox(Ecran.Components[i]).OnClick) then
              TCheckBox(Ecran.Components[i]).OnClick := CritereChange;
            if (Ecran.Components[i] is THCritMaskEdit) and not
              assigned(THCritMaskEdit(Ecran.Components[i]).OnChange) then
              THCritMaskEdit(Ecran.Components[i]).OnChange := CritereChange;
            if (Ecran.Components[i] is THRadioGroup) and not
              assigned(THRadioGroup(Ecran.Components[i]).OnClick) then
              THRadioGroup(Ecran.Components[i]).OnClick := CritereChange;
          end;

        end;
      end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.SearchTimerTimer(Sender: TObject);
begin
  if csDestroying in Ecran.ComponentState then Exit ;
  FSearchTimer.Enabled := False;
  CritModified := True;
  OnClickBCherche(nil);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.ResetTimer(Sender: TObject);
begin

  if csDestroying in Ecran.ComponentState then Exit ;

  if FLoading then Exit;
  
  CritModified := True;
  FSearchTimer.Enabled := False;
  FSearchTimer.Enabled := True;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.CritereChange(Sender: TObject);
begin
  CritModified := True;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.SetCritModified(Value: Boolean);
begin
  FCritModified     := Value;
  BCherche.Visible  := not Value;
  BCherche_.Visible := Value;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.LoadCumulControls;
var
  i: integer;
  lStName: string;
begin
  for i := 0 to FPanelCumul.ControlCount - 1 do
  begin
    lStName := FPanelCumul.Controls[i].Name;
    if THSystemMenu(GetControl('HMTrad')).LockedCtrls.IndexOf(lStName) = -1 then
      THSystemMenu(GetControl('HMTrad')).LockedCtrls.Add(lStName);
  end; // for
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.UnloadCumulControls;
var
  i, idx: integer;
  lStName: string;
begin
  for i := 0 to FPanelCumul.ControlCount - 1 do
  begin
    lStName := FPanelCumul.Controls[i].Name;
    idx := THSystemMenu(GetControl('HMTrad')).LockedCtrls.IndexOf(lStName);
    if idx <> -1 then
      THSystemMenu(GetControl('HMTrad')).LockedCtrls.Delete(idx);
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_ViergeMul.UpdateGrille;
begin
  LoadCumulControls;
  Exit; // MBAMF
  if ((THSystemMenu(GetControl('HMTrad')).ActiveResize) and (V_PGI.Outlook)) then
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_ViergeMul.BeforeLoad: boolean;
begin
  Result := True;
  FTobRow := nil; // GCO - 30/08/2007
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_ViergeMul.AfterLoad: boolean;
begin
  Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2003
Modifié le ... :   /  /
Description .. : Tri de la grille avec les touches ALT + [Colonne]
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnKeyDownFListe(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  (* GCO - 19/12/2007 - FQ 22032
  if FListe.MultiSelect then
  begin
    if (key = VK_Space) and (ATobFListe <> nil) and
       (ATobFliste.Detail.Count > 0) then
    begin

    end;
  end; *)

  if FListe.SortEnabled then
  begin
    if (Shift = [ssAlt]) and (Key in [Ord('1')..Ord('9')]) then
    begin
      case key of
        Ord('1') : FListe.SortGrid(1, (FListe.SortedCol = 1) xor FListe.SortDesc);
        Ord('2') : FListe.SortGrid(2, (FListe.SortedCol = 2) xor FListe.SortDesc);
        Ord('3') : FListe.SortGrid(3, (FListe.SortedCol = 3) xor FListe.SortDesc);
        Ord('4') : FListe.SortGrid(4, (FListe.SortedCol = 4) xor FListe.SortDesc);
        Ord('5') : FListe.SortGrid(5, (FListe.SortedCol = 5) xor FListe.SortDesc);
        Ord('6') : FListe.SortGrid(6, (FListe.SortedCol = 6) xor FListe.SortDesc);
        Ord('7') : FListe.SortGrid(7, (FListe.SortedCol = 7) xor FListe.SortDesc);
        Ord('8') : FListe.SortGrid(8, (FListe.SortedCol = 8) xor FListe.SortDesc);
        Ord('9') : FListe.SortGrid(9, (FListe.SortedCol = 9) xor FListe.SortDesc);
      else
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/09/2006
Modifié le ... :   /  /    
Description .. : GCO - 13/09/2006 - FQ 16728 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.OnMouseDownFListe(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (Button in [mbLeft]) and (ssCtrl in Shift) then
  begin
    SelectionneLigne;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/09/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.SelectionneLigne;
begin
  if Fliste.AllSelected then
    SelectionneTout(False);

  if ATobFListe.Detail[FListe.Row-1].GetString('SELECTED') = 'X' then
    ATobFListe.Detail[FListe.Row-1].SetString('SELECTED', '-')
  else
    ATobFListe.Detail[FListe.Row-1].SetString('SELECTED', 'X');
  FMulPanelFiltre.UpdateCumulSelect;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/10/2005
Modifié le ... : 30/08/2007
Description .. : On passe systématiquement dedans, pour que ATobRow soit bon
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  FTobRow := GetO(FListe);
{$IFDEF EAGLCLIENT}
  BFetchGauche.Enabled := (FListe.TopRow > 1);
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2003
Modifié le ... : 06/11/2003
Description .. : Recherche le nom du champ qui corresopnd à la colonne
Suite ........ : sélectionnée dans la grille FLISTE, et demande un Refresh
Suite ........ : de la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnSortedFListe(Sender: Tobject);
var  lStListeChamps : string;
     i : integer;
begin
{$IFDEF EAGLCLIENT}
  if BFetchDroit.Enabled then Exit;
{$ENDIF}  

  // Ajouter ici tri sur la TOB pour recalculer ensuite le Solde
  lStListeChamps := FStListeChamps;

  i := 0;
  while (i <> FListe.SortedCol) do
  begin
    AStTriTobFListe := ReadTokenSt( lStlisteChamps );
    Inc(i);
  end;

//  FBoOkSortFListe := True;

  // GCO - 29/05/2006 - FQ 17871
  if AStTriTobFListe = 'POURCENTAGE' then
  begin
    AStTriTobFListe := 'XPOURCENTAGE';
  end;

  // Tri sur la colone demandée + Ordre de Tri initial de la LISTE PARAMETRABLE
  AStTriTobFListe := AStTriTobFListe + ';' + FindEtReplace(FStNewTris, ',', ';', True);

  // GCO - 02/01/2008 - FQ 19306 - RefreshFliste apparement inutile, et provoque
  // la perte de la sélection car fait passé critmodified à true, ce qui envoie
  // un BCherche.Click;
  // RefreshFListe( False );
end;

//////////////////////////// GESTION DU FETCH //////////////////////////////////
{$IFDEF EAGLCLIENT}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/10/2003
Modifié le ... : 26/01/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnClickBFetchGauche(Sender: TObject);
begin
  SendNotifyMessage( TWinControl(FListe).Handle, WM_KeyDown, 33, 0 ) ;
  BFetchGauche.Enabled := (FListe.TopRow > 1);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/10/2003
Modifié le ... : 26/01/2005
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLCLIENT}
procedure TOF_ViergeMul.OnClickBFetchDroit(Sender: TObject);
begin
  FaitOpenSql(True);
  RefreshFListe( True );
  BFetchGauche.Enabled := True;
  BFetchDroit.Enabled  := not VarAsType(FQuery.GetValue('EOF'), varBoolean);
  FListe.SortEnabled := not BFetchDroit.Enabled;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/10/2003
Modifié le ... : 25/01/2005
Description .. : Prepare la Grille FLLISTE ainsi que ATobFLsite à l'affichage
Suite ........ : (1) - Resize les colonnes de la grille ( FBUG 12693 )
Suite .........: (2) - Tri TOBFLISTE par la colonne sélectionée
Suite ........ : (3) - Tri TOBFLISTE par ordre Croissant ou Décroissant
Suite ........ : (4) - Gestion de la sélection de la ligne avec 'SELECTED'
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.RefreshFListe( vBoFetch : Boolean );
var i : integer;
    lInDepBoucle : integer ;
//    lRec : _TTRec ;
begin
  if not vBoFetch then
    FListe.AllSelected := False;

  // Si Nil, premier chargement de l'écran donc on vide la grille
  if (ATobFListe = nil) or (ATobFListe.Detail.Count = 0) then
  begin
    FListe.VidePile( False );
  end
  else
  begin
    // GCO - 24/10/2005
    if FListe.SortEnabled then
    begin
      ATobFListe.Detail.Sort(AStTriTobFListe);
      {$IFDEF EAGLCLIENT}
      ATobFListe.SetDouble('DEPARTBOUCLE', 0);
      {$ENDIF}
    end;

    //if FBoOkSortFListe then
    //begin
    //  // Tri TobFListe sur le colonne de la grille désirée
    //  ATobFListe.Detail.Sort( AStTriTobFListe );
    //  FBoOkSortFListe := False;
    //  ATobFListe.SetDouble('DEPARTBOUCLE', 0);
    //end;

    // Inverse le tri de TOBFListe si grille demande à être trié à l'envers
    if FListe.SortDesc then
    begin
      for i := ATobFListe.Detail.Count - 1 downto 0 do
        ATobFListe.Detail[i].ChangeParent(ATobFListe, -1);
    end;

    if FBoOkAncetreFaitAffichage then
    begin
      HGBeginUpdate(FListe);
      try
        FListe.RowCount := ATobFListe.Detail.Count + 1;
      {$IFDEF EAGLCLIENT}
        lInDepBoucle := ATobFListe.GetValue('DEPARTBOUCLE');
      {$ELSE}
        lInDepBoucle := 0;
      {$ENDIF}
        RemplitListeChamps( ATobFListe.Detail[0] ) ;

        for i := lInDepBoucle to (ATobFListe.Detail.Count - 1) do
        begin
          // Affichage de la ligne dans la Grille
          FListe.Row := i + 1;
          CalculpourAffichage(ATobFListe.Detail[i]);
          TOBVersTHGrid(ATobFListe.Detail[i] );
          // GCO - NB : sert à resélectionner les lignes après un sort par exemple
          if ATobFListe.Detail[i].GetString('SELECTED') = 'X' then
            FListe.FlipSelection(FListe.Row);
        end;
      finally
        HGEndUpdate(FListe);
      end;
    end; //if FBoOkAncetreFaitAffichage then
  end;

{$IFDEF EAGLCLIENT}
  if not vBoFetch then
  begin
    BFetchGauche.Enabled := FListe.TopRow > 1;
    BFetchDroit.Enabled  := False;
    if (FQuery <> nil) then
    begin
      if not VarAsType(FQuery.GetValue('EOF'), varBoolean) then
        BFetchDroit.Enabled := True;

      FListe.SortEnabled := not (BFetchDroit.Enabled);
    end;
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2003
Modifié le ... : 22/06/2005
Description .. : 
Suite ........ : CA - 22/06/2005 : ajout paramètre pour forcer le 
Suite ........ : chargement de tous les enregistrements de la requête en 
Suite ........ : WebAccess
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.FaitOpenSql( vBoFetch : Boolean = False; vboFetchLesTous : boolean = False );
var lTobTemp   : Tob;
{$IFDEF EAGLCLIENT}
    lInAdd   : Integer;
    lInDepBoucle : Integer;
    i : Integer;
{$ENDIF}
    lBoTableToViewActif : Boolean ;
begin
  if not vBoFetch then
  begin
    ATobFListe.ClearDetail;
  end;

  lBoTableToViewActif := False ;
  // Mise en place de zap table -> vue poru MultiSoc // SBO 18/08/2005
  if EstMultiSoc then
    begin
    lBoTableToViewActif := V_PGI.EnableTableToView ;
    V_PGI.EnableTableToView  := true ;
    end ;

  // Recaulcul du FMaxRow à chaque OpenSQL ou AppendSQL, car l'utilisateur a
  // peut être retaillé son écran.
  FMaxRow := (FListe.Height div FListe.DefaultRowHeight) ;
  if vboFetchLesTous then FMaxRow := -1; ///astuce du fetchlestous

  if AStSqlTobFListe <> '' then
  begin
    try
      try
      {$IFDEF EAGLCLIENT}
        if not vBoFetch then
        begin
          if Assigned(FQuery) then Ferme(FQuery);
          FQuery := OpenSql( AStSqlTobFListe, True, FMaxRow );
          lInDepBoucle := 0;
          ATobFListe.AddChampSupValeur('DEPARTBOUCLE', lInDepBoucle, False);
        end
        else
        begin
          // Ajout de nouvelles lignes dans la grille ( Appel par le FETCH )
          lInAdd := AppendSQL( FQuery , FMaxRow ) ;
          lInDepBoucle := FQuery.Detail.Count - lInAdd;
          ATobFListe.AddChampSupValeur('DEPARTBOUCLE', lInDepBoucle, False);
        end;

        for i := lInDepBoucle to (FQuery.Detail.Count - 1) do
        begin
          lTobTemp := Tob.Create('', nil, -1);
          lTobTemp.SelectDB('', FQuery, False);
          lTobTemp.AddChampSupValeur('SELECTED', '-', False);
          if AjouteATobFListe( lTobTemp ) then
            lTobTemp.ChangeParent(ATobFListe, -1)
          else
            lTobTemp.Free;
          FQuery.Next;
        end;

      {$ELSE}
        FQuery := OpenSql( AStSqlTobFListe, True );
        if not FQuery.Eof then
        begin
          while not FQuery.Eof do
          begin
            lTobTemp := Tob.Create('', nil, -1);
            lTobTemp.SelectDB('', FQuery, False);
            lTobTemp.AddChampSupValeur('SELECTED', '-', False);

            if AjouteATobFListe( lTobTemp ) then
              lTobTemp.ChangeParent(ATobFListe, -1)
               else
                lTobTemp.Free ;

            FQuery.Next;
          end;
        end;
      {$ENDIF}
      except
        on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : FaitOpenSql');
      end;
    finally
    {$IFDEF EAGLCLIENT}
    {$ELSE}
      if Assigned(FQuery) then Ferme(FQuery);
    {$ENDIF}
    end;
  end;

  // GCO - 01/03/2004
  // Si pas d'enregistrement dans la grille, on envoie une chaine vide en
  // paramètre au PanelCumul afin de lui dire de ne pas faire le SELECt SUM
  if ATobFListe.Detail.Count > 0 then
    FMulPanelFiltre.StWhereSql := FStSqlWhereTobFListe
  else
    FMulPanelFiltre.StWhereSql := '';

  // Annulation mise en place du zapping table -> vue pour MultiSoc // SBO 18/08/2005
  if EstMultiSoc then
    V_PGI.EnableTableToView  := lBoTableToViewActif ;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/02/2004
Modifié le ... : 27/02/2004
Description .. : Enregistre vValue dans FStSqlTobFListe et récupère la condition
Suite ........ : WHERE de vValue pour la stocker dans FStSqlWhereTobFListe
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.TraiteStSqlTobFListe(vValue: string);
var lSt : string;
begin
  FStSqlTobFListe := vValue;
  lSt := vValue;
  System.Delete( lSt, Pos('ORDER', UpperCase( lSt )), Length( lSt ));
  lSt := Copy( lSt, Pos('WHERE', Uppercase( lSt )), Length( lSt ));
  FStSqlWhereTobFListe := lSt;
end;

///////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnPostDrawCellFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TheRect : TRect;
begin

//GP le 18/09/2008 : report debug PCL
  If ARow=0 Then Exit ;
  If ACol=0 Then Exit ;

  if (FListe.ColTypes[ACol] <> 'R') then Exit;

  if FListe.Cells[ACol, ARow]= 'NS' then Exit;

  //lValeur := Valeur(FListe.Cells[ACol, ARow]);

  //lStValeur := FloatToStr(lValeur);

  //if IsNumeric(lStValeur) and (lValeur = 0) then

  if ((Trim(FListe.Cells[ACol, ARow]) = '') or
      (Valeur(FListe.Cells[ACol, ARow]) = 0)) then
  begin
    TheRect := FListe.CellRect(ACol, ARow);
    Canvas.TextRect(TheRect, TheRect.Left, TheRect.Top, '');
    Canvas.Brush.Color := FListe.FixedColor;
    Canvas.Brush.Style := bsBDiagonal;
    Canvas.Pen.Color := FListe.FixedColor;
    Canvas.Pen.Mode := pmCopy;
    Canvas.Pen.Style := psClear;
    Canvas.Pen.Width := 1;
    Canvas.Rectangle(TheRect);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/07/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.AfterShow;
begin
{$IFDEF CCSTD}
  FObjFiltre.ForceAccessibilite := faPublic;
{$ENDIF}
  BCherche.Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.OnBeforeFlipFListe(Sender: TObject; ARow: Integer; var Cancel: Boolean);
begin
  //
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/08/2005
Modifié le ... : 19/12/2007
Description .. :
Suite ........ : GCO - 19/12/2007 - FQ 22032
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnFlipSelectionFListe(Sender: TObject);
begin
  if (FLoading) or (not FListe.MultiSelect) or (ATobFListe.Detail.Count = 0) then Exit;
  FBoMultiSelected := (FListe.nbSelected > 0) or (FListe.AllSelected);
  SelectionneLigne;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/10/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnColumnWidthsChangedFListe(Sender: TObject);
begin
  Exit;
  FMulPanelFiltre.UpdateLayoutCumul;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.OnResizeEcran(Sender: TObject);
begin
  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
  FMulPanelFiltre.UpdateLayoutCumul;

  if Assigned(FPanelCumulSelect) then
    FMulPanelFiltre.UpdateLayoutCumulSelect;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.TraitementBValider(Sender: TObject);
begin
 if assigned(FListe.OnDblClick) then
  FListe.OnDblClick(Self);
end;

////////////////////////////////////////////////////////////////////////////////

{$IFDEF EAGLCLIENT}
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/06/2005
Modifié le ... :   /  /
Description .. : Chargement de la requête complète dans la grille
Mots clefs ... :
*****************************************************************}
function TOF_ViergeMul.FetchLesTous: boolean;
begin
  Result := FALSE;
  if FPanelCumul.Tag <= 0 then
  begin
    result := TRUE;
  end else
  begin
    if (FPanelCumul.Tag > 1000) and
      (HShowMessage('0;ATTENTION Traitement long !;La sélection concerne $$ enregistrements. #13#10 Confirmez -vous le traitement ?;W;YNC;N;C', '', IntToStr(FPanelCumul.Tag)) <> mrYes)
      then exit;
    result := TRUE;
    OnClickBCherche(bExport);
    Application.ProcessMessages;
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/05/2006
Modifié le ... :   /  /    
Description .. : Factorisation des appels à la fonction ChargeHListe
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.TraitementListeParam;
var i : integer;
    lStListeChamps : string;
begin
  // Récupération des nouveaux paramètres de la LISTE
  // Recup des infos de la liste paramètrable
  // GCO - 12/05/2006 - FQ 18074
  ChargeHListe (FStListeParam, FNewSources, FNewLiaison, FStNewTris, FStListeChamps, FStNewTitres, FStNewLargeurs, FNewJustifs,
               FNewParams, FNewLibelle, FNewNumCols, FNewPerso, FNewOkTri, FNewOkNumCol);

  // GCO - 13/06/2007 - FQ 20536
  // A la différence d'un vrai mul, le colcount de notre grille doit toujours
  // être égale au nombre de champs inscrits dans la liste, qu'il soit visible ou pas
  // ce que ne fait pas l'agl, qui affecte un colcount avec le nombre de visible
  // Recalcul du nombre de colonnes à afficher
  i := 0;
  lStListeChamps := FStListeChamps;
  while lStListechamps <> '' do
  begin
    ReadTokenSt(lStListeChamps);
    Inc(i);
  end;
  FListe.ColCount := i + 1;

  //Chargement des composants Z_C1, Z_C2, Z_C3
  CChargeZ_C( Z_C1, FStListeChamps, FStNewTitres );
  CChargeZ_C( Z_C2, FStListeChamps, FStNewTitres );
  CChargeZ_C( Z_C3, FStListeChamps, FStNewTitres );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/10/2005
Modifié le ... : 01/12/2005
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_ViergeMul.AjouteATobFListe(vTob: Tob): Boolean;
begin
  Result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/12/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.CalculPourAffichage(vTob: Tob);
begin
  //
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.RefreshPclPge;
begin
  if BCherche = nil then Exit;
  if FBoOkRefresh then
    BCherche.Click
  else
    FCritModified := true;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ViergeMul.SelectionneTout(vBoSelected: Boolean);
var i : integer;
begin
  for i := 0 to ATobFliste.Detail.Count - 1 do
  begin
    ATobFListe.Detail[i].SetString('SELECTED', IIF(vBoSelected, 'X', '-'));
  end;
  FMulPanelFiltre.UpdateCumulSelect;
end;


procedure TOF_ViergeMul.RemplitListeChamps( vTOB : TOB ) ;
var
 lStListeChamps    : string ;
 lStListeChampsOld : string ;
 lStChamps         : string ;
 lIndex            : integer ;
 i                 : integer ;
begin

 lIndex := 0 ;

 if (FListe.Titres.Count > 0) then
  lStListeChamps := Fliste.Titres[0]
   else
    lStListeChamps := FStListeParam ;

 lStListeChampsOld := lStListeChamps ;

 repeat
  lStChamps  := ReadTokenSt(lStListeChamps) ;
  if lStChamps <> '' then
   Inc(lIndex) ;
 until (lStListeChamps = '') or (lStChamps = '') ;

 SetLength(FChamps,lIndex) ;

 for i := 0 to lIndex - 1 do
  begin
   lStChamps         := ReadTokenSt(lStListeChampsOld) ;
   FChamps[i].Nom    := lStChamps ;
   FChamps[i].Index  := vTOB.GetNumChamp(lStChamps) ;
  end ;

end ;

procedure TOF_ViergeMul.TOBVersTHGrid( vTOB : TOB ) ;
var
  lStChamps: string;
  lStMask : String ;
  lValue  : variant ;
  lDt     : TDatetime ;
  lStCell : string ;
  i       : integer ;
begin

  if vTOB = nil then Exit;

  for i := low(FChamps) to high(FChamps) do
   begin
    lValue     := vTOB.GetValeur(FChamps[i].Index) ;
    lStMask    := THGrid(FListe).ColFormats[i+1] ;
    lStChamps  := FChamps[i].Nom ;

    case VarType(lValue) of

      varInteger : begin
                     if lStMask = '' then
                       lStMask := '#0';
                     lStCell := FormatFloat(lStMask,VarAsType(lValue, varInteger)) ;
                   end;

      varDouble,
      varSingle :  begin
                     if lStMask = '' then lStMask := '#,##0.00' ;
                     if Copy(lStChamps, 1, 5) = 'SOLDE' then
                     begin
                       // GCO - 05/10/2007 - PB Bureau FQ 11740
                       THGrid(FListe).Coltypes[i+1]   := #0 ;
                       THGrid(FListe).ColFormats[i+1] := '' ;
                       lStCell := AfficheDBCR(lValue, lStMask);
                     end  
                     else
                       lStCell := FormatFloat(lStMask,VarAsType(lValue, varDouble)) ;
                   end;

     varDate :     begin
                     if Pos('#',lStMask) > 0 then
                       lStMask := '' ;
                     lDt := VarAsType(lValue, VarDate) ;
                     // GCO - 04/10/2007 - pkoi réaffecter Coltype ??
                     THGrid(FListe).Coltypes[i+1] := 'D' ;
                     if lStMask = '' then
                       // GCO - 18/03/2008 - Correction du format des dates
                       if Frac(lDt) <> 0 then
                         lStMask := TraduitDateFormat('dd/mm/yyyy hh:nn:ss')
                       else
                         lStMask := TraduitDateFormat('dd/mm/yyyy');

                     if lDt = iDate1900 then
                       lStCell := ''
                     else
                       lStCell := FormatDatetime(lStMask,lDt) ;
                   end ;

     varNull : lStCell := '' ;

     else
       lStCell := VarAsType(lValue, VarString) ;

    end ; // case

    FListe.Cells[i+1, FListe.Row] := lStCell ;
  end ;

  SetO(FListe, vTOB);

end;

function TOF_ViergeMul.GetCumulChamps ( vStChamps : string ) : double ;
begin
 result := FMulPanelFiltre.GetCumulChamps(vStChamps) ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.AvantChangementFiltre;
begin
  //
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.ApresChangementFiltre;
begin
  //
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.SupprimeFiltre;
begin

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ViergeMul.NouveauFiltre;
begin
  VideFiltre(FFiltres, PageControl) ;
  InitControl;
  BCherche.Click;
end;

initialization
  registerclasses([TOF_ViergeMul]);
end.



