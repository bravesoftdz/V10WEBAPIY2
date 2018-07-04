{***********UNITE*************************************************
Auteur  ...... : Thierry PETETIN
Créé le ...... : 03/10/2001
Description .. : La Tomme de Savoie à ne pas confondre avec la Tome des Bauges
Mots clefs ... : TOMME;TOME;REBLOCHON;BEAUFORT;ABONDANCE;CAMEMBERT;BLEU D'AUVERGNE;CROTIN DE CHAVIGNOL;ST FELICIEN;
Mots clefs ... : GRUYERE;COMTE;CAPRICE DES DIEUX;MORBIER;MUNSTER;CHAUSSE AUX MOINES;TAMIER;VACHERIN;LIVARO;BRESSE BLEU;
Mots clefs ... : GOUDA;MIMOLETTE;KIRI;BABIBEL;VACHE QUI RIT;ROCHEBARON;EPOISSE;BOURSIN;TARTARE;ST MORIN;CHAOURCE;
Mots clefs ... : BRIE DE MAUX;COULOMIER;CANCOILLOTTE;ST MARCELIN;BROCCIU;RACLETTE;BRESIN;EMMENTAL;FETA;MOZZARRELLA;PARMESAN;
Mots clefs ... : GORGONZOLA;MASCARPONE;BRIQUE;CHAVROUX;SOIGNON;ST LOUP;ROQUEFORT;ST AGUR;VIEUX PANE;CHAMOIX D'OR;CHAUME;
Mots clefs ... : GRUYERE SUISSE;COMTE SUISSE;ST ALBRAY;TOMME AU MARC;BREBIOU;MAROIL;PAVE DAPHINOIX;BELLE DES CHAMPS;
Mots clefs ... : LE PETIT LOUIS;FOL EPI;MONT D'OR;
*****************************************************************}
Unit wTom;

Interface

Uses
  Classes,
  StdCtrls,
  Controls,
  Extctrls,
  Windows,
  Messages,
  Graphics,
  Dialogs,
  ImgList,
  {$IFNDEF EAGLCLIENT}
    db,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    {$IFNDEF ERADIO}
      Fe_Main,
      Fiche,
      FichList,
    {$ENDIF !ERADIO}
  {$ELSE}
    MainEagl,
    eFiche,
    eFichList,
  {$ENDIF}
  {$IFNDEF ERADIO}
    SaisieList,
  {$ENDIF !ERADIO}
  forms,
  sysutils,
  ComCtrls,
  Spin,
  dbCtrls,
  hCtrls,
  hEnt1,
  hMsgBox,
  hTB97,
  hRichOle,
  Menus,
  Hdb,
  uTOMComm,
  UTob,
  UtilConfid {JTR - Confidentialité sur la fiche}
  ;

Type
  TOnGetArticle = procedure(const Article: String) of object;
  TOnGetTiers   = procedure(const Tiers: String) of object;

  TWTom = Class (tTOMComm)
    procedure OnNewRecord             ; override;
    procedure OnUpdateRecord          ; override;
    procedure OnAfterUpdateRecord     ; override;
{$IFNDEF PGIMAJVER}
    procedure OnDeleteRecord          ; override;
    procedure OnLoadRecord            ; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: String)   ; override;
    procedure OnClose                 ; override;
    procedure OnCancelRecord          ; override;
{$ENDIF !PGIMAJVER}
  private
    {$IFNDEF PGIMAJVER}
      {$IFNDEF ERADIO}
        PnChgUnite1, PnChgUnite2, PnChgUnite1_1, PnChgUnite2_1: tPanel;
      {$ENDIF !ERADIO}
      PageControlName: String;
      TabSheetName   : String;
      MemoFieldName  : String;
    {$ENDIF !PGIMAJVER}
    FOnGetTiers       : TOnGetTiers;
    FOnGetArticle     : TOnGetArticle;
    FArticleByElipsis,
    FFromBValider     : Boolean;

    { Boutons }
    {$IFNDEF PGIMAJVER}
      {$IFNDEF ERADIO}
        procedure BFermeOnClick(Sender: TObject);
        procedure BValiderOnClick(Sender: TObject);
        procedure BRECHAFFAIRE_OnClick(Sender: TObject);
        procedure BStop_OnClick(Sender: tObject);
        procedure BTChgUniteOnClick(Sender: TObject);
        procedure BTChgUnite1OnClick(Sender: TObject);
      {$ENDIF !ERADIO}
    {$ENDIF !PGIMAJVER}

		Procedure GetRetour;

    {$IFNDEF PGIMAJVER}
      { ImgMemo }
      procedure GereImgMemo(Const FieldName: String);

      { image sur la fiche }
      procedure GereImgFic(Const FieldName: String);
    {$ENDIF PGIMAJVER}


    {$IFNDEF ERADIO}
      {$IFNDEF PGIMAJVER}
        { Image Liste }
        procedure SetImageList;
      {$ENDIF !PGIMAJVER}
    {$ENDIF !ERADIO}

    { Action }
    {$IFDEF GPAOLIGHT}
      {$IFNDEF ERADIO}
        procedure PmAction_OnPopUp(Sender: TObject);
        procedure PmBaseTech_OnPopUp(Sender: TObject);
        procedure MnJournalAction_OnClick(Sender: TObject);
      {$ENDIF !ERADIO}
    {$ENDIF GPAOLIGHT}

    {$IFNDEF PGIMAJVER}
    {$IFNDEF ERADIO}
      { Historique }
      procedure PmHistorique_OnPopUp(Sender: TObject);

      { Utilitaire }
      procedure PmUtilitaire_OnPopUp(Sender: TObject);

      { Analyse }
      procedure PmAnalyse_OnPopUp(Sender: TObject);
    {$ENDIF !ERADIO}
    {$ENDIF !PGIMAJVER}

    { Set }
    {$IFNDEF PGIMAJVER}
    {$IFNDEF ERADIO}
      { Moulination des controls posés sur la fiche (Elipsis & on Hint) }
      procedure ScanControls;
    {$ENDIF !ERADIO}
    {$ENDIF !PGIMAJVER}
    procedure wSetBmemo;
    procedure wSetIdentifiant;
    procedure wSetInitFields(const FromRecordIsValid: Boolean = False);
    procedure SetUppercase;

    { Vérification des champs obligatoires }
    function VerifChampsObligatoires: boolean;
    function VerifInfosLibres: Boolean;
    function GetArticleByElipsis: Boolean;
  public
    Contexte: string;
    ModifLotRetour:string;

    {$IFNDEF ERADIO}
      procedure wOuvreMul(Const Fiche, TableFille, Range, Params: string);
    {$ENDIF !ERADIO}

    procedure SetControlsVisible(Const FieldsName: Array of String; Const Visibility: Boolean);
    procedure SetControlsEnabled(Const FieldsName: Array of String; Const Enability: Boolean);
    {$IFNDEF ERADIO}
    procedure SetArguments(sChampsModifLot: string); //modification en série
    {$ENDIF !ERADIO}

    { Divers }
    function EcranIs(const FormName: String): Boolean;

    property OnGetArticle: TOnGetArticle read FOnGetArticle write FOnGetArticle;
    property OnGetTiers  : TOnGetTiers   read FOnGetTiers   write FOnGetTiers;
    property ArticleByElipsis: Boolean read GetArticleByElipsis write FArticleByElipsis;
  protected
    Action : string;
		Droit  : string;
    function RecordIsvalid: boolean; override;
    {$IFNDEF PGIMAJVER}
      function GetLoupeCtx: string; Override;
      {$IFNDEF ERADIO}
        { Elipsis Article }
        procedure ArticleElipsisClick(Sender: TObject); Virtual;
{GPAO_BUG_TS Déb}
        { Elipsis Tiers }
        procedure TiersElipsisClick(Sender: TObject); Virtual;
      {$ENDIF !ERADIO}
    {$ENDIF !PGIMAJVER}
  end;

const
  ErrTTLibreValue     = 998;
  ErrChampObligatoire = 999;

Implementation

Uses
  {$IFDEF GPAO}
    wAction,
    wInitChamp,
    EntGP,
    wRecordCMS,
  {$ENDIF GPAO}
  {$IFDEF GPAOLIGHT}
    {$IFNDEF ERADIO}
      wGereForms,
      GereDiff_GC_GP,
    {$ENDIF !ERADIO}
    wNatureTravail,
    wJournalAction,
    wOrdrePhase,
  {$ENDIF GPAOLIGHT}
  wJetons,
  wCommuns,
  {$IFNDEF PGIMAJVER}
    {$IFNDEF ERADIO}
      wMnu,
    {$ENDIF !ERADIO}
    {$IFNDEF BTP}
    	{$IFDEF GCGC}
      yTarifs,
      yTarifsCommun,
    	{$ENDIF GCGC}
    {$ELSE}
    	tarifs,
  	{$ENDIF}
    AGLInitGC,
    UtilArticle,
    FactUtil,
  {$ENDIF !PGIMAJVER}
  Math,
  ParamSoc,
  UtilPGI,
  CbpMCD,
  CbpEnumerator
  ;

{$IFNDEF PGIMAJVER}
procedure tWTom.OnArgument ( S: String ) ;
begin
  { TsDimensions actif? }
  SetControlProperty('TSDIMENSIONS', 'TABVISIBLE', GetParamSoc('SO_WDIMENSION'));

  Inherited ;

  Action := GetArgumentValue(S, 'ACTION');
  Droit  := GetArgumentValue(S, 'DROIT');

  {$IFNDEF ERADIO}
    { On est toujours en mode MONOFICHE }
    if Assigned(Ecran) and (Ecran is TFFiche) then
      TFFiche(Ecran).MonoFiche := True;

    { Boutons }
    if Assigned(GetControl('BFERME')) then
      TToolBarButton97(GetControl('BFERME')).OnClick := bFermeOnClick;
    if (Ecran is TFFICHE) and Assigned(GetControl('BVALIDER')) then
      TToolBarButton97(GetControl('BVALIDER')).OnClick := bValiderOnClick;
    if Assigned(GetControl('BTCHGUNITE')) then
      TToolBarButton97(GetControl('BTCHGUNITE')).OnClick := BtChgUniteOnClick;
    if Assigned(GetControl('BTCHGUNITE1')) then
      TToolBarButton97(GetControl('BTCHGUNITE1')).OnClick := BtChgUnite1OnClick;
    if Assigned(GetControl('BRECHAFFAIRE')) then
      TToolBarButton97(GetControl('BRECHAFFAIRE')).OnClick := BRECHAFFAIRE_OnCLick ;
    {Pour la modification par lot}
    if Assigned(GetControl('BSTOP')) then
    	TToolBarButton97(GetControl('BSTOP')).OnClick := BStop_OnClick;

    { Panels }
    PnChgUnite1 := tPanel(GetControl('PNCHGUNITE1'));
    if Assigned(PnChgUnite1) then
      PnChgUnite1.BevelOuter := bvNone;

    PnChgUnite2 := tPanel(GetControl('PNCHGUNITE2'));
    if Assigned(PnChgUnite2) then
    begin
      PnChgUnite2.Visible := false;
      PnChgUnite2.BevelOuter := bvNone;
    end;

    PnChgUnite1_1 := tPanel(GetControl('PNCHGUNITE1_1'));
    if Assigned(PnChgUnite1_1) then
      PnChgUnite1_1.BevelOuter := bvNone;

    PnChgUnite2_1 := tPanel(GetControl('PNCHGUNITE2_1'));
    if Assigned(PnChgUnite2_1) then
    begin
      PnChgUnite2_1.Visible := false;
      PnChgUnite2_1.BevelOuter := bvNone;
    end;

    { Bouton Elipsis recherche Article }
    ScanControls;

    { Image List pour les TabSheet }
    SetImageList;

    { Action }
    {$IFDEF GPAOLIGHT}
    if Assigned(Ecran) and not (Ecran is TFSAISIELIST) then
    begin
      { Action }
      if Assigned(GetControl('PMACTION')) then
        TPopUpMenu(GetControl('PMACTION')).OnPopUp := PmAction_OnPopUp;
      if Assigned(GetControl('PMBASETECH')) then
        TPopUpMenu(GetControl('PMBASETECH')).OnPopUp := PmBaseTech_OnPopUp;
      if Assigned(GetControl('MNJOURNALACTION')) then
        TMenuItem(GetControl('MNJOURNALACTION')).OnClick := MnJournalAction_OnClick;
    end;
    {$ENDIF GPAOLIGHT}

    { Historique }
    if Assigned(GetControl('PMHISTORIQUE')) then
      TPopUpMenu(GetControl('PMHISTORIQUE')).OnPopUp := PmHistorique_OnPopUp;

    { Utilitaire }
    if Assigned(GetControl('PMUTILITAIRE')) then
      TPopUpMenu(GetControl('PMUTILITAIRE')).OnPopUp := PmUtilitaire_OnPopUp;

    { Analyse }
    if Assigned(GetControl('PMANALYSE')) then
      TPopUpMenu(GetControl('PMANALYSE')).OnPopUp := PmAnalyse_OnPopUp;

    if Assigned(GetControl('BTUTILITAIRE')) then
      SetControlVisible('BTUTILITAIRE', Action='MODIFICATION');

    { ActivePage }
    if Assigned(Getcontrol('PAGES')) and Assigned(GetControl('PGENERAL')) then
      tPageControl(GetControl('PAGES')).ActivePage := tTabSheet(GetControl('PGENERAL'));

    { Tab Sheet TsSav: Visible en mode SAV }
    SetControlProperty('TSSAV', 'TABVISIBLE', V_PGI.Sav);
    SetControlProperty('TSSAV', 'CAPTION'   , TraduireMemoire('Vision SAV'));

    {$IFDEF GPAOLIGHT}
      {$IFNDEF GPAO}
	      { on va cacher ou rendre disabled certains champs }
	      AppliqueModifAssemblage(nil, Self, 'OnArgument');
      {$ENDIF GPAO}
    {$ENDIF GPAOLIGHT}

    if Assigned(GetControl('BTLOUPE')) then
    begin
      with TToolBarButton97(GetControl('BTLOUPE')) do
      begin
        Hint := TraduireMemoire('Menu zoom');
        ShowHint := True;
      end
    end;
  {$ENDIF !ERADIO}

  {$IFNDEF ERADIO}
    {$IFDEF GPAOLIGHT}
    if Assigned(Ecran) then
      wGereFormsByCtx(Ecran, Self, 'OnArgument');
    {$ENDIF GPAOLIGHT}

    { Confidentialité }
    AppliquerConfidentialite(Ecran, ''); {JTR - Confidentialité sur la fiche}
  {$ENDIF !ERADIO}
end;
{$ENDIF !PGIMAJVER}

procedure tWTom.OnNewRecord;
begin
  wSetIdentifiant;

  Inherited ;

  { Initialisation des champs }
  wSetInitFields;

  {$IFNDEF ERADIO}
    {$IFDEF GPAOLIGHT}
    if Assigned(Ecran) then
      wGereFormsByCtx(Ecran, Self, 'OnNewRecord');
    {$ENDIF GPAOLIGHT}
  {$ENDIF !ERADIO}
end;

{$IFNDEF PGIMAJVER}
procedure tWTom.OnDeleteRecord;
begin
  Inherited;

  if LastError = 0 then
  begin
    { Gestion du RecordCMS }
    {$IFDEF GPAO }
    if FieldExists(GetPrefixe + '_IDENTIFIANT') then
    begin
      wCreateWRD(GetPrefixe, GetInteger(GetPrefixe + '_IDENTIFIANT'), 'SUP');
    end;
    {$ENDIF GPAO}
  end;  
end;
{$ENDIF !PGIMAJVER}

procedure tWTom.OnUpdateRecord ;
begin
  Inherited;
  DisableControl;
  try
    if LastError = 0 then
    begin
      wSetBMemo;
      SetUpperCase;

      { Getion du RecordCMS }
      {$IFDEF GPAO}
      if FieldExists(GetPrefixe + '_IDENTIFIANT') then
      begin
        wCreateWRD(GetPrefixe, GetInteger(GetPrefixe + '_IDENTIFIANT'), iif(Ikc='C', 'CRE', 'MOD'));
      end;
      {$ENDIF GPAO}

      { Retourne la clé actuelle de la fiche pour 'refresher' le Mul }
      GetRetour;
    end;
  finally
    EnableControl;
  end;
end;

procedure TWTom.OnAfterUpdateRecord ;
begin
  Inherited;
end ;

{$IFNDEF PGIMAJVER}
procedure tWTom.OnLoadRecord ;
begin
  DisableControl;
  try
    Inherited;
	  if Ikc <> 'C' then
    begin
      GetRetour;
    end;
  finally
    EnableControl;
  end;

  {$IFNDEF ERADIO}
    {$IFDEF GPAOLIGHT}
    if Assigned(Ecran) then
      wGereFormsByCtx(Ecran, Self, 'OnLoadRecord')
    {$ENDIF GPAOLIGHT}
  {$ENDIF !ERADIO}
end ;
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
procedure tWTom.OnChangeField(F: TField);
var
  FieldName: string;
begin
  FieldName := F.FieldName; { Mémorisé car change en CWas... }
  
  Inherited ;

  { Gestion de l'image Mémo }
  GereImgMemo(FieldName);

  { Gestion de l'image associée au champ }
  GereImgFic(FieldName);
end ;
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
procedure tWTom.OnClose ;
begin
  Inherited ;

	{ Pour gérer l'arrêt du traitement de modif. en série fiche par fiche }
  if ModifLotRetour='Stop' then
    if ecran is tfFiche then
    	tfFiche(Ecran).Retour:=ModifLotRetour
    else if ecran is tfFicheListe then
    	tfFicheListe(Ecran).Retour:=ModifLotRetour;

end ;
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
procedure tWTom.OnCancelRecord ;
begin
  Inherited ;
end ;
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure tWTom.BFermeOnClick(Sender: TObject);
begin
	if not FFromBValider and (Ikc = 'C') then
    TFFICHE(ECRAN).Retour := '';
  TFFICHE(ECRAN).BFermeClick(Sender);
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 11/10/2001
Modifié le ... :   /  /
Description .. : Ouverture d'un Mul depuis la fiche associée à la TOM
Description .. : -----------------------------------------------------
Description .. : sFiche      => Nom de la fiche à ouvrir
Decription ..  : sTableFille => Nom de la table fille gérée dans le MUL,
Decription ..  :                pour fabrication du range et des paramètres de la clé [ ]
Description .. : sRange  	  => Range du mul (utilisé si sTablefille est à blanc)
Description .. : sParams     => Paramètres à passer au Mul
Description .. : 					  (Si pas de [ ] dedans, et que sTablefille <> '' fabrique les paramètres de la clé)
Description .. : -----------------------------------------------------
Mots clefs ... : AGLLANCEFICHE;MUL
*****************************************************************}
procedure tWTom.WOuvreMul(Const Fiche, TableFille, Range, Params: string);
var
	PrefixeMere, PrefixeFille: String;
  MyRange, MyParams: String;
begin
  if TableFille <> '' then
  begin
    { Récupère les préfixes des tables pour le range }
    PrefixeMere := GetPrefixe;
    PrefixeFille := TableToPrefixe(TableFille);
    MyParams := Params;

    { Fabrique le range du Mul }
    MyRange := WMakeRange(PrefixeFille, PrefixeMere, True, WGetNbrFieldInKey(GetTableName), DS);
    if (Pos('[', Range) = 0) and (Pos(']', Range) = 0) then
    begin
      { Fabrique les paramètres à passer au mul }
      MyParams := '[' + WMakeRange( PrefixeFille, PrefixeMere, False, WGetNbrFieldInKey(GetTableName), DS) + ']';
      if Params <> '' then MyParams := MyParams + ';' + Params;
    end;
    { Ouvre la fiche }
    AglLanceFiche ('W', Fiche, MyRange, '', MyParams);
  end
  else
    AglLanceFiche ('W', Fiche, Range, '', Params);
end;
{$ENDIF !ERADIO}

Procedure TWTom.GetRetour;
begin
  {$IFNDEF ERADIO}
    if Assigned(Ecran) then
    begin
      TFFICHE(ECRAN).Retour := wGetValueClef1(GetTableName, TFFICHE(ECRAN));
    end
  {$ENDIF !ERADIO}
end;

{$IFNDEF PGIMAJVER}
procedure tWTom.GereImgMemo(Const FieldName: string);
const
  ImgMemo = 4;
begin
  if (FieldName = MemoFieldName) and (PageControlName <> '') then
  begin
    if assigned(GetControl('TSBLOCNOTE')) then
    	TTabSheet(GetControl(TabSheetName)).ImageIndex := iif((GetString(MemoFieldName) <> ''), imgMemo - 1, -1);
  end;
end;
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure tWTom.SetImageList;
var
  i        : integer;
  CMemo, C : tControl;

  function GetMemoName: String;
  begin
    Result := GetPrefixe + '_BLOCNOTE'
  end;
begin
  if Assigned(Ecran) then
  begin
    { Gestion de l'imageList pour les mémos }
    MemoFieldName := GetMemoName;

    { Recherche si il y a un controle xxx_BLOCNOTE sur la fiche }
    CMemo := GetControl(MemoFieldName);
    if Assigned(CMemo) then
    begin
      { Recherche du nom du tabSheet sur lequel est posé le champ xxx_BLOCNOTE }
      C := CMemo.Parent;
      while Assigned(C) and (not (C is TTabSheet)) do
      begin
        C := C.Parent;
      end;

      if Assigned(C) then
      begin
        { Sauve le nom du tabSheet }
        TabSheetName := TTabSheet(C).Name;

        { Sauve le nom du pagecontrol }
        PageControlName := TPageControl(C.Parent).Name;

        for i := 0 to TPageControl(GetControl(PageControlName)).PageCount - 1 do
          TPageControl(GetControl(PageControlName)).Pages[i].ImageIndex := -1;

        ChargeImageList;
        TPageControl(GetControl(PageControlName)).Images := V_Pgi.GraphList;
      end;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 03/04/2002
Modifié le ... :   /  /
Description .. : Mise à jour du champ WBMEMO à vrai si BLOCNOTE <> ''
Mots clefs ... :
*****************************************************************}
procedure tWTom.wSetBmemo;
var
  xLastError     : integer;
  xLastErrorMsg  : string;
begin
  if (ChampToNum(GetPrefixe + '_WBMEMO') > 0) and (ChampToNum(GetPrefixe + '_BLOCNOTE') > 0) then
  begin
    { Bien mémoriser le LastError avant }
    xLastError := LastError;
    xLastErrorMsg := LastErrorMsg;
    try
      SetBoolean(GetPrefixe + '_WBMEMO', (GetString(GetPrefixe + '_BLOCNOTE') <> ''));
    finally
      LastError := xLastError;
      LastErrorMsg := xLastErrorMsg;
    end;
  end;
end;

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure tWTom.BTChgUniteOnClick(Sender: TObject);
begin
	if (not Assigned(GetControl('PNCHGUNITE1'))) or (not Assigned(GetControl('PNCHGUNITE2')))then
    Abort;
	SetControlVisible('PNCHGUNITE1', Not TToolbarButton97(GetControl('BTCHGUNITE')).Down);
	SetControlVisible('PNCHGUNITE2', TToolbarButton97(GetControl('BTCHGUNITE')).Down);
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure tWTom.BTChgUnite1OnClick(Sender: TObject);
begin
	if (not Assigned(GetControl('PNCHGUNITE1_1'))) or (not Assigned(GetControl('PNCHGUNITE2_1'))) then
    Abort;
	SetControlVisible('PNCHGUNITE1_1', Not TToolbarButton97(GetControl('BTCHGUNITE1')).Down);
	SetControlVisible('PNCHGUNITE2_1', TToolbarButton97(GetControl('BTCHGUNITE1')).Down);
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure TWTom.PmHistorique_OnPopUp(Sender: TObject);

  function GetArgument: String;
  begin
    if GetPrefixe = 'WOB' then
    begin
      Result := ';TYPELIEN='  + GetString('WOB_TYPELIEN')
              + ';MODECONSO=' + GetString('WOB_MODECONSO')
    end;

    Result := Result + ';ACTION=' + Action;
  end;

begin
  wSetMnuHisto(false, GetArgument, tPopUpMenu(GetControl('PMHISTORIQUE')));
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{$IFDEF GPAOLIGHT}
{$IFNDEF ERADIO}
procedure tWTom.PmAction_OnPopUp(Sender: TObject);
const
  sCannotDoActionOnTypeOrdre = '[STA];[VTE]';

  {$IFDEF GPAO}
    function GetCleWOP: tCleWOP;
    begin
      Result.NatureTravail  := GetString('WOP_NATURETRAVAIL');
      Result.LigneOrdre     := GetInteger('WOP_LIGNEORDRE');
      Result.OpeCirc        := GetString('WOP_OPECIRC');
    end;
  {$ENDIF GPAO}
	{ Calcul du contexte }
  function GetCtx: string;
  Var
    TypeOrdre: string;
  begin
    if GetPrefixe = 'WOT' then
    begin
      if GetString('WOT_TYPEORDRE')<>'STA' then
        Result := 'ETATTET=' + GetString('WOT_ETATTET')
      else
        Result := '';
    end
	  else if GetPrefixe = 'WOL' then
    begin
      if (GetString('WOL_TYPEORDRE')<>'STA') and (GetString('WOL_MISEENPROD') <> 'ALL') then
      begin
        Result := 'ETATTET=' + wGetSqlFieldValue('WOT_ETATTET', 'WORDRETET', 'WOT_NATURETRAVAIL="' + GetString('WOL_NATURETRAVAIL') + '" AND WOT_CODEORDRE=' + GetString('WOL_CODEORDRE'))
                + ';ETATLIG=' + GetString('WOL_ETATLIG')
                + ';MISEENPROD=' + GetString('WOL_MISEENPROD');
      end
      else
        Result := '';
    end
   	else if GetPrefixe = 'WOP' then
    begin
      TypeOrdre:= wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOP_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOP_LIGNEORDRE'));
      if Pos('[' + TypeOrdre + ']', sCannotDoActionOnTypeOrdre) = 0 then
        Result := 'ETATLIG=' + wGetSqlFieldValue('WOL_ETATLIG', 'WORDRELIG', 'WOL_NATURETRAVAIL="' + GetString('WOP_NATURETRAVAIL') + '" AND WOL_LIGNEORDRE=' + GetString('WOP_LIGNEORDRE'))
		            + ';ETATPHASE=' + GetString('WOP_ETATPHASE')  + ';QLANSAIS=' + FloatToStr(GetDouble('WOP_QLANSAIS')) + ';QACCSAIS=' + FloatToStr(GetDouble('WOP_QACCSAIS'))
                + iif((Pos('[' + TypeOrdre + ']', sCannotDoActionOnTypeOrdre) = 0) and (GetString('WOP_TIERS') <> ''),';TYPEORDRE=STP','')
                {$IFDEF GPAO}
                  + iif(ExistePiecesCSPEnCours(GetCleWOP),';PIECESCSPENCOURS=X','')
                {$ENDIF GPAO}
      else
        Result := '';
    end
	  else if GetPrefixe = 'WOB' then
    begin
      if wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOB_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOB_LIGNEORDRE'))<>'STA' then
        Result := 'ETATLIG=' + wGetSqlFieldValue('WOL_ETATLIG', 'WORDRELIG', 'WOL_NATURETRAVAIL="' + GetString('WOB_NATURETRAVAIL') + '" AND WOL_LIGNEORDRE=' + GetString('WOB_LIGNEORDRE'))
                + ';ETATPHASE=' + wGetSqlFieldValue('WOP_ETATPHASE', 'WORDREPHASE', 'WOP_NATURETRAVAIL="' + GetString('WOB_NATURETRAVAIL') + '" AND WOP_LIGNEORDRE=' + GetString('WOB_LIGNEORDRE') + ' AND WOP_OPECIRC="' + GetString('WOB_OPECIRC') + '"')
                + ';ETATBES=' + GetString('WOB_ETATBES')
                + ';TYPELIEN=' + GetString('WOB_TYPELIEN')
                + ';MODECONSO=' + GetString('WOB_MODECONSO')
      else
        Result := '';
    end
   	else if GetPrefixe = 'WOG' then
    begin
      Result := 'ETATOPE=' + GetString('WOG_ETATOPE');
    end
   	else if GetPrefixe = 'WOR' then
    begin
      Result := 'ETATOPE=' + wGetSqlFieldValue('WOG_ETATOPE', 'WORDREGAMME', 'WOG_NATURETRAVAIL="' + GetString('WOR_NATURETRAVAIL') + '" AND WOG_LIGNEORDRE=' + IntToStr(GetInteger('WOR_LIGNEORDRE')) + ' AND WOG_OPECIRC="' + GetString('WOR_OPECIRC') + '" AND WOG_NUMOPERGAMME=' + intToStr(GetInteger('WOR_NUMOPERGAMME')) )
    end
    else if GetPrefixe = 'WPC' then
    begin
      Result := 'ETATPARC=' + GetString('WPC_ETATPARC')
    end
    else if GetPrefixe = 'WVS' then
    begin
      Result := 'ETATVER=' + GetString('WVS_ETATVER')
    end
   	else if GetPrefixe = 'WLS' then
    begin
      if Ecran is TFSAISIELIST then
        Result := 'ETATLAS=' + TFSaisieList(Ecran).LeFiltre.TOBFiltre.GetString('WLS_ETATLAS')
      else
      	Result := 'ETATLAS=' + GetString('WLS_ETATLAS');
    end
    else if GetPrefixe = 'QSI' then
      begin
      	if Ecran is TFSAISIELIST then
         	Result := 'FOURNI=' + TFSaisieList(Ecran).LeFiltre.TOBFiltre.GetString('QSI_FOURNI')
         else
         	Result := 'FOURNI=' + GetString('QSI_FOURNI');
      end
      else if GetPrefixe = 'ULS' then
      begin
      	if Ecran is TFSAISIELIST then
         	Result := 'ETATLASST=' + TFSaisieList(Ecran).LeFiltre.TOBFiltre.GetString('ULS_ETATLASST')
         else
         	Result := 'ETATLASST=' + GetString('ULS_ETATLASST');
      end
	    else if GetPrefixe = 'RQN' then
  	  begin
    	  Result := 'ETATLIG=' + GetString('RQN_ETATQNC')
    	end
      else if GetPrefixe = 'RQD' then
      begin
        Result := 'ETATLIG=' + GetString('RQD_ETATDR');
      end
      else if GetPrefixe = 'RQP' then
      begin
        Result := 'ETATLIG=' + GetString('RQP_ETATPC');
      end
      else if GetPrefixe = 'RAC' then
      begin
        Result := 'ETATLIG=' + GetString('RAC_ETATACTION');
      end
  	;

    Result := Result + ';ACTION=' + Action;
  end;
{ Main PmAction_OnPopUp }
begin
	wSetMnuAction(false, GetCtx, tPopUpMenu(GetControl('PMACTION')));
end;
{$ENDIF !ERADIO}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF ERADIO}
procedure tWTom.PmBaseTech_OnPopUp(Sender: TObject);
	{ Calcul du contexte }
  function GetCtx: string;
  begin
	  if GetPrefixe = 'WOL' then
    begin
      Result := 'ETATLIG=' + GetString('WOL_ETATLIG')
    end;

    Result := Result + ';ACTION=' + Action;
  end;
begin
	wSetMnuBaseTech(false, GetCtx, tPopUpMenu(GetControl('PMBASETECH')));
end;
{$ENDIF !ERADIO}
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
{$IFNDEF ERADIO}
procedure TWTom.MnJournalAction_OnClick(Sender: TObject);
var
  BoutonAnnule : Boolean;
const
  TypeOrdre = '[STA];[VTE]';
begin
  BoutonAnnule := True;
  if      GetPrefixe = 'WOL' then
    BoutonAnnule := (Pos('[' + wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOL_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOL_LIGNEORDRE')) + ']', TypeOrdre) = 0)
  else if GetPrefixe = 'WOP' then
    BoutonAnnule := (Pos('[' + wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOP_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOP_LIGNEORDRE')) + ']', TypeOrdre) = 0)
                    and (GetString('WOP_TIERS') = '')  // N'est pas une phase de sous-traitance
  else if GetPrefixe = 'WOB' then
    BoutonAnnule := (Pos('[' + wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOB_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOB_LIGNEORDRE')) + ']', TypeOrdre) = 0);

 	wCallMulWJA(GetPrefixe, GetInteger(GetPrefixe + '_IDENTIFIANT'), Action, BoutonAnnule);
  if DS.State = DsBrowse then
    RefreshDB;
end;
{$ENDIF !ERADIO}
{$ENDIF GPAOLIGHT}

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 26/09/2002
Description .. : Déroute le onElipsisclick sur les contrôles THedit
Suite ........ : ou THdbEdit pour les contrôles 'articles'
NDP : Cette procédure est répliquée dans la wtof, merci de la mettre
à jour si vous modifiez cette procédure la.
*****************************************************************}
procedure TWTom.ScanControls;
var
	iComp : Integer;
  Ok    : Boolean;
  C     : TComponent;
begin
	if Assigned(Ecran) then
  begin
    for iComp := 0 to TForm(Ecran).ComponentCount - 1 do
    begin
      C := TForm(Ecran).Components[iComp];

      { Traitement du CLick sur l'Elipsis des 'articles' }
      Ok := false;

      {$IFNDEF EAGLCLIENT}
      if (C is THDBEdit) and (THDBEdit(C).ElipsisButton) then
      begin
        if wIsDataTypeArticle(ThDBEdit(C).DataType) or wIsDataTypeCodeArticle(ThDBEdit(C).DataType) then
        begin
          THDBEdit(C).OnElipsisClick := ArticleElipsisClick;
          Ok := True;
        end
        else if wIsDataTypeTiers(ThDBEdit(C).DataType) then
        begin
          THDBEdit(C).OnElipsisClick := TiersElipsisClick;
          Ok := True;
        end;
      end;
      {$ENDIF !EAGLCLIENT}

      if (not Ok) and (C is THEdit) and (THEdit(C).ElipsisButton) then
      begin
        { Essai avec le suffixe du nom du champ, Si le name est pas ok, ré-essaye avec le DataType }
        if wIsDataTypeArticle(THEdit(C).DataType) or wIsDataTypeCodeArticle(THEdit(C).DataType) then
          THEdit(C).OnElipsisClick := ArticleElipsisClick
        else if wIsDataTypeTiers(THEdit(C).DataType) then
          THEdit(C).OnElipsisClick := TiersElipsisClick;
      end;
    end;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 26/09/2002
Description .. : Appel le mul ou le lookup de recherche des articles
NDP : Cette procédure est sauvagement répliquée dans la wTOF, merci
de la mettre à jour si vous modifiez cette procédure si...
*****************************************************************}
procedure TWTom.ArticleElipsisClick(Sender: TObject);
var
	Text, Params, Plus, Name, DataType: String;
  ReturnCodeArticle: Boolean;
  MaxLen: Integer;
  {$IFDEF EAGLCLIENT}
    AEdit: THEdit;
  {$ELSE  EAGLCLIENT}
    AEdit: THDBEdit;
  {$ENDIF EAGLCLIENT}
begin
  Plus := ''; Name := ''; Text := ''; DataType := '';
  Params := '';

  {$IFDEF EAGLCLIENT}
    AEdit := THEdit(Sender);
  {$ELSE  EAGLCLIENT}
    AEdit := THDBEdit(Sender);
  {$ENDIF EAGLCLIENT}

  Plus := AEdit.Plus;
  Name := AEdit.Name;
  Text := AEdit.Text;
  DataType := AEdit.DataType;

  { Retourne le code article ou l'article ? }
  ReturnCodeArticle := wIsDataTypeCodeArticle(DataType);

  { Initialise le code article (ou libellé si Ctrl enfoncé) dans le mul }
  if Text <> '' then
  begin
    if Params <> '' then Params := Params + ';';
    if ssCtrl in GetKBDState() then
      Params := Params + 'GA_LIBELLE=' + Text
    else
      Params := Params + 'GA_CODEARTICLE=' + Text
  end;
 	if Params <> '' then
    Params := Params + ';';
  Params := Params + 'RECHERCHEARTICLE';

  { Accès à tous les articles ou seulement aux article profils dans le mul de recherche des articles }

  { Etat du DataSet lié à la fiche }
  if (not (ecran is TFSaisieList)) and Assigned(DS) and (DS.State = dsBrowse) then
    DS.Edit;

  { Appel du mul + Mise à jour du contrôle }
  MaxLen := AEdit.MaxLength;
  AEdit.MaxLength := TailleArticle;
  try
    if Pos('XX_WHERE', Plus) > 0 then
    begin
      { on vire le XX_WHERE du plus car si l'on met à jour le .Text du THEdit, le noyau effectue une requête
        pour mettre à jour le libellé associé. Et si on a XX_WHERE... Erreur SQL ! Le plus est rétabli plus loin }
      AEdit.Plus := GetArgumentString(Plus, 'XX_WHERE');
      GetArticleRecherche_Disp(ThCritMaskEdit(Sender), '', Params + ';' + Plus, NomMulRechArticle, ReturnCodeArticle);
    end
    else
      GetArticleRecherche_Disp(ThCritMaskEdit(Sender), Plus, Params, NomMulRechArticle, ReturnCodeArticle);

    FArticleByElipsis := True;

    if Assigned(FOnGetArticle) then
      FOnGetArticle(AEdit.Text);
  finally
    { comme promis, on rétablit le .Plus }
    AEdit.Plus := Plus;
    AEdit.MaxLength := MaxLen
  end
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 01/06/2006
Description .. : Appel le mul ou le lookup de recherche des tiers
NDP : Cette procédure est sauvagement répliquée dans la wTOF, merci
de la mettre à jour si vous modifiez cette procédure si...
*****************************************************************}
procedure TWTom.TiersElipsisClick(Sender: TObject);
begin
  DispatchRecherche(THCritMaskEdit(Sender), 2, THDBEdit(Sender).Plus,
                    'T_TIERS=' + Trim(GetControlText(TWinControl(Sender).Name)), '');

  if Assigned(FOnGetTiers) then
    FOnGetTiers(GetControlText(TWinControl(Sender).Name));
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 28/01/2003
Modifié le ... :   /  /
Description .. : Gère l'image associée à la valeur du champ
suite ........ : dans le OnChangeField
Mots clefs ... :
*****************************************************************}
procedure TWTom.GereImgFic(Const FieldName: String);
var
	sFieldNameTemp : String;
  sFieldNameTemp2: String;
  NumImage       : Integer;
  C              : TControl;

  procedure ExtractUnderScore(var sFieldName: String);
  begin
    while Pos('_', sFieldName) <> 0 do
      Delete(sFieldName, Pos('_', sFieldName), 1);
  end;

begin
  sFieldNameTemp := FieldName;
  { sFieldNameTemp de la forme : XXX_YYYYYYY ... }
  Delete(sFieldNameTemp, 1, 4);
  { ...devient YYYYYYY ... }
  sFieldNameTemp := 'IMG' + sFieldNameTemp;
  { ...puis IMGYYYYYYY ... }
  { ...Pour traiter le cas multiple : YYYYYYY_ ; YYYYYYY__ ... }
  sFieldNameTemp2 := sFieldNameTemp;
  { ...Enlève les '_'... }
  if Pos('_', sFieldNameTemp2) <> 0 then
    ExtractUnderScore(sFieldNameTemp2);
  if Assigned(GetControl(sFieldNameTemp2)) then    // Test si le contrôle YYYYYYY ou YYYYYYY_ ... existe
  begin
    { ...récupération du contrôle auquel on doit associer une image... }
    C := GetControl(FieldName);
		ChargeImageList;
    { ...init de l'image à rien... }
    tImage(GetControl(sFieldNameTemp)).Picture.Bitmap := nil;
    { on triate les différents cas de contrôles pour récupérer le DataType et ainsi récupérer le N° d'image associé par "GetNumImage" }
    if (C is THDBValComboBox) or (C is THValComboBox) then
      NumImage := GetNumImage((C as THValComboBox).DataType, 0, GetField(FieldName)) - 1
    else if (C is THDBEdit) or (C is THEdit) then
      NumImage := GetNumImage((C as THEdit).DataType, 0, GetField(FieldName)) - 1
    else
      NumImage := -1;

    { ...on affecte l'image au contrôle TImage }
    if NumImage > -1 then
      V_PGI.GraphList.GetBitmap(NumImage, tImage(GetControl(sFieldNameTemp)).Picture.Bitmap);
  end;
end;
{$ENDIF !PGIMAJVER}

function TWTom.VerifChampsObligatoires: boolean;
{$IFDEF GPAO}
  var
    ErrorFieldName : String;
{$ENDIF GPAO}
begin
  {$IFDEF GPAO}
    { Vérification des champs obligatoires }
    ErrorFieldName := '';

    if Assigned(DS) then
       Result := wInitFieldsbyDS(DS, GetTableName, iaVerif, ErrorFieldName)
    else
       Result := wInitFields(fTob, GetTableName, iaVerif, ErrorFieldName);

    if not Result then
    begin
      LastError := ErrChampObligatoire;
      LastErrorMsg := wGetLastErrorMessageFromWIL(ErrorFieldName);
      if Assigned(ecran) then
      begin
        SetFocusControl(ErrorFieldName);
      end
    end;
  {$ELSE}
    Result := True;
  {$ENDIF GPAO}
end;

{ Vérification des champs libres : Donnée existe dans la tablette libre }
function TWTom.VerifInfosLibres: Boolean;
var
  i, iTable,
  CptChampLibres: Integer;
  TT: String;
  ACombo: TDECombo;
  //
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
  Field     : IFieldCOM ;

  function ExistInDeCombo(const Value: String): Boolean;
  var
    iVal: Integer;
  begin
    Result := False;
    iVal := 0;
    while not Result and (iVal < ACombo.Valeurs.Count) do
    begin
      Result := Pos(Value + #9, ACombo.Valeurs[iVal]) = 1;
      Inc(iVal);
    end
  end;

  { Est un champ libre si le nom du champ et son datatype correspondent }
  function IsChampLibre(var Tablette: String): Boolean;
  begin
    if Pos(GetPrefixe + '_LIBRE', Field.Name) = 1 then
    begin
      Tablette := ChampToTT(Field.Name);
      Result := (Pos('WLIBRE', Tablette) = 1) or (Pos('GCLIBRE', Tablette) = 1)
    end
    else
      Result := False;
  end;

begin
	MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  //
  Result := True;
  FillChar(ACombo, SizeOf(ACombo), #0);
  if not Assigned(Ecran) then
  begin
    Table := Mcd.GetTable(GetTableName);

    if Assigned(Table) then
    begin
      FieldList := Table.Fields;
      FieldList.Reset();
      CptChampLibres := 0;
      While (result) and  (CptChampLibres < 10) and (FieldList.MoveNext) do
      begin
        Field := (FieldList.current as IFieldCOM);
        if IsChampLibre(TT) then
        begin
          ACombo := v_Pgi.DECombos[TTToNum(TT)];
          Inc(CptChampLibres);
          if GetString(Field.Name) <> '' then
          begin
            if TabletteNonChargee(TT, '') then
              RemplirListe(TT, '');
            Result := not Assigned(ACombo.Valeurs) or ExistInDeCombo(GetString(Field.Name))
          end
          else
            Result := True
        end;
        Inc(i);
      end
    end
  end;

  if not Result then
  begin
    LastError := ErrTTLibreValue;
    LastErrorMsg := GetString(Field.Name) + TraduireMemoire(' : valeur inexistante dans la table libre : ') + ACombo.Libelle + ' (' + ACombo.TT + ') !'
  end
end;

function TWTom.RecordIsvalid: boolean;
begin
  Result := false;
  
  { Contrôle bas niveau }
  if not inherited RecordIsvalid then exit;

  { Init des champs }
  DisableControl;
  try
    wSetInitFields(True);
  finally
    EnableControl;
  end;

  { Vérification des champs obligatoires }
  Result := VerifChampsObligatoires and VerifInfosLibres;
  if not Result and not Assigned(Ecran) then
    FTob.AddChampSupValeur('Error', LastErrorMsg, False)
end;

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Attention, c'est pour éviter de contrôler deux fois...
Mots clefs ... :
*****************************************************************}
procedure TWTom.BValiderOnClick(Sender: TObject);
begin
  LastError := 0;
  try
    {$IFDEF EAGLCLIENT}
      if (Action <> 'CONSULTATION') and (Ds.State = dsBrowse) then
        Ds.Edit;
    {$ENDIF}

    NextPrevControl(ecran);
    FFromBValider := LastError = 0;
    if LastError = 0 then
      TFFICHE(ECRAN).BValiderClick(Sender);
  finally
    LastError := 0;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

procedure TWTom.wSetInitFields(const FromRecordIsValid: Boolean = False);
{$IFDEF GPAO}
var
	sInitializedFields: String;
{$ENDIF GPAO}
begin
  if Pos(GetTableName + ';', 'DISPODETAIL;STKMOUVEMENT;') > 0 then exit;

  {$IFDEF GPAO}
//  PutToDebugLog('TWTom.wSetInitFields', true, 'Début');
  try
    sInitializedFields := '';

    { Initialisation des champs }
    if Assigned(DS) then
      wInitFieldsbyDS(DS, GetTableName, iaInit, sInitializedFields)
    else
      wInitFields(fTob, GetTableName, iaInit, sInitializedFields);

    { affiche la liste des champs initialisés dans le debug log }
    if sInitializedFields <> '' then
      PutToDebugLog('TWTom.wSetInitFields', false, 'Liste des champs initialisés: ' + sInitializedFields);

    if Assigned(Ecran) and not FromRecordIsValid then
    begin
      LoadBufferAvantModif;
      while sInitializedFields <> '' do
        ControlField(ReadTokenSt(sInitializedFields));
    end;
  finally
//    PutToDebugLog('TWTom.wSetInitFields', false, 'Fin');
  end;  
  {$ENDIF GPAO}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. : permet de grouper les SetControlEnabled
Mots clefs ... :
*****************************************************************}
procedure TWTom.SetControlsEnabled(Const FieldsName: array of String; Const Enability: Boolean);
var
  i: integer;
begin
  for i := 0 to Length(FieldsName) - 1 do
    SetControlEnabled(FieldsName[i], Enability);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 08/04/2003
Modifié le ... : 08/04/2003
Description .. : permet de grouper les SetControlVisible
Mots clefs ... :
*****************************************************************}
procedure TWTom.SetControlsVisible(Const FieldsName: array of String; Const Visibility: Boolean);
var
  i: integer;
begin
  for i := 0 to Length(FieldsName) - 1 do
    SetControlVisible(FieldsName[i], Visibility);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 31/01/2003
Modifié le ... :   /  /
Description .. : Génération de l'identifiant au niveau de la wTom
Mots clefs ... :
*****************************************************************}
procedure TWTom.wSetIdentifiant;
begin
  if Pos('SANSJETON', Contexte) > 0 then exit;

  if FieldExists(GetPrefixe + '_IDENTIFIANT') then
  begin
    SetInteger(GetPrefixe + '_IDENTIFIANT', wSetJeton(GetPrefixe));
  end
end;

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure TWTom.PmUtilitaire_OnPopUp(Sender: TObject);
  { Calucl du contexte }
  function GetCtx: string;
  begin
    if GetPrefixe = 'WOL' then
    begin
      Result := 'TYPEORDRE=' + GetString('WOL_TYPEORDRE') + ';ETATLIG=' + GetString('WOL_ETATLIG')
    end
    else if GetPrefixe = 'WOP' then
    begin
      Result := 'ETATPHASE=' + GetString('WOP_ETATPHASE')
    end
    else if GetPrefixe = 'WOB' then
    begin
      Result := 'ETATBES='    + GetString('WOB_ETATBES')
              + ';MODECONSO=' + GetString('WOB_MODECONSO')
              + ';TYPELIEN='  + GetString('WOB_TYPELIEN')
    end
    else if GetPrefixe = 'GQ' then
    begin
      Result := 'DEPOT='         + GetString('GQ_DEPOT')
              + ';ARTICLE='      + GetString('GQ_ARTICLE')
              + ';DATECLOTURE='  + GetString('GQ_DATECLOTURE')
              {$IFNDEF PGIMAJVER}
                + ';CONTREMARQUE=' + wGetFieldFromGA('GA_CONTREMARQUE', GetString('GQ_ARTICLE'))
              {$ENDIF !PGIMAJVER}
    end
    else if GetPrefixe = 'GSM' then
    begin
      Result := 'ETATMVT=' + GetString('GSM_ETATMVT')
    end
    {$IFDEF EDI}
      else if GetPrefixe = 'ELI' then
      begin
        Result := 'IDENTIFIANTELI=' + GetString('ELI_IDENTIFIANT')
                + ';NATUREPIECEG='  + GetString('ELI_NATUREPIECEG')
      end
    {$ENDIF EDI}
    ;

    Result := Result + ';ACTION=' + Action;
  end;
{ Main PmUtilitaire_OnPopUp }
begin
  wSetMnuUtilitaire(false, GetCtx, tPopUpMenu(GetControl('PMUTILITAIRE')));
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure TWTom.PmAnalyse_OnPopUp(Sender: TObject);
  function GetCtx: string;
  begin
    if GetPrefixe = 'WOB' then
    begin
      Result := ';TYPELIEN='  + GetString('WOB_TYPELIEN')
              + ';MODECONSO=' + GetString('WOB_MODECONSO')
    end;

    Result := Result + ';ACTION=' + Action;
  end;
begin
  wSetMnuAnalyse(false, GetCtx, tPopUpMenu(GetControl('PMANALYSE')));
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

procedure TWTom.SetUpperCase;
  procedure uCase(Const FieldName: string);
  begin
    if FieldExists(FieldName) then
      SetString(FieldName, UpperCase(GetString(FieldName)));
  end;
begin
  uCase(GetPrefixe + '_' + 'ARTICLE');
  uCase(GetPrefixe + '_' + 'CODEARTICLE');

  uCase(GetPrefixe + '_' + 'COMPOSANT');
  uCase(GetPrefixe + '_' + 'CODECOMPOSANT');
end;

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
procedure TWTom.BRECHAFFAIRE_OnClick(Sender: TObject);
var
  Affaire: string;
begin
  if Assigned(GetControl(GetPrefixe + '_AFFAIRE')) then
  begin
    GetAffaireMul(Affaire, nil, nil, nil, nil,nil, ThCritMaskEdit(GetControl(GetPrefixe + '_TIERS')), false, true, false, true, true, false);
    if Assigned(Ecran) and (Ecran is TFSaisieList) then
      SetControlText(GetPrefixe + '_AFFAIRE', Affaire)
    else
      SetString(GetPrefixe + '_AFFAIRE', Affaire);
    if FieldExists(GetPrefixe + '_TIERS') then
      SetString(GetPrefixe + '_TIERS', ThCritMaskEdit(GetControl(GetPrefixe + '_TIERS')).text);
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

function TWTom.GetArticleByElipsis: Boolean;
begin
  Result := FArticleByElipsis;
  FArticleByElipsis := False;
end;

function TWTom.EcranIs(const FormName: String): Boolean;
begin
  if Assigned(Ecran) then
    Result := Pos(UpperCase(Trim(FormName)), UpperCase(Ecran.Name)) = 1
  else
    Result := False
end;

{$IFNDEF ERADIO}
procedure TWTom.SetArguments(sChampsModifLot: string);
Var
	Critere,ChampMul, ValMul :string;
  x,y				: integer;
  Ctrl			: TControl;
begin
  SetControlVisible('BSTOP', TRUE);

  if not (DS.State in [dsInsert, dsEdit]) then
  	DS.Edit;
  repeat
    Critere := Uppercase(Trim(ReadTokenPipe(sChampsModifLot, '|')));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul 	:= copy(Critere, 1, x - 1);
        ValMul 		:= copy(Critere, x + 1, length(Critere));
        y := pos(',', ValMul);
        if y <> 0 then ValMul := copy(ValMul, 1, length(ValMul) - 1);
        if copy(ValMul, 1, 1) = '"' then
        	ValMul := copy(ValMul, 2, length(ValMul));
        if copy(ValMul, length(ValMul), 1) = '"' then
        	ValMul := copy(ValMul, 1, length(ValMul) - 1);
				if ChampMul<>'' then
        begin
					if (wGetTypeField(ChampMul) = 'DATE') then
        		SetField(ChampMul, StrToDateTime(ValMul))
        	else
          	SetField(ChampMul, ValMul);
        end;
        Ctrl := TControl(TfFiche(Ecran).FindComponent(ChampMul));
        if Ctrl = nil then
        	exit;
        {$IFDEF EAGLCLIENT}
        	if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) or (Ctrl is TCustomEdit) then
        		TEdit(Ctrl).Font.Color := clRed
        	else if Ctrl is TSpinEdit then
        		TSpinEdit(Ctrl).Font.Color := clRed;
        {$ELSE}
        	if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) or (Ctrl is THDBEdit) then
        		TEdit(Ctrl).Font.Color := clRed
        	else if Ctrl is THDBSpinEdit then
        		THDBSpinEdit(Ctrl).Font.Color := clRed;
        {$ENDIF}
      end;
    end;
  until Critere = '';
end;
{$ENDIF !ERADIO}

{$IFNDEF PGIMAJVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Dominique sclavopoulos
Créé le ...... : 18/11/2005
Modifié le ... :   /  /
Description .. : STOP : Pour gérer l'arrêt du traitement de modification en
Suite ........ : série (Si mise à jour fiche par fiche)
Mots clefs ... :
*****************************************************************}
procedure TWTom.BStop_OnClick(Sender: tObject);
begin
	if PGIAsk('Voulez-vous arrêter le traitement de modification en série ?', Ecran.Caption) = MrYes then
  begin
    ModifLotRetour := 'Stop';
  	TfFiche(ECRAN).Close;
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF !PGIMAJVER}

{$IFNDEF PGIMAJVER}
function TWTom.GetLoupeCtx: string;
const
  sCannotDoActionOnTypeOrdre = '[STA];[VTE]';
Var
  TypeOrdre: string;

  {$IFDEF GCGC}
  {---------------------------------------------------------------------------
    Construit l'argument avec les champs du MUL et la valeur
  ----------------------------------------------------------------------------}
  function GetArgumentYTS : string;
  var
    sParametreYFOActif, sParametreYFOSelection, sParametreYFOEntete, sParametreYFOClef : string;
    sQuelschamps, sArgument, sArgChamp, sYTSChamp, sChampsActifs : string;
  begin
    Result := '';
    GetParametresFromYFO(GetString('YTS_FONCTIONNALITE'),GetString('YTS_CODEPORT'),'TIE', sParametreYFOActif, sParametreYFOSelection, sParametreYFOEntete, sParametreYFOCLef);

    if (Ecran<>nil) and (Ecran.name='YTARIFSENTETE_FIC') then
      sQuelsChamps := 'ChampsEntêteClef'
    else
      sQuelschamps := 'ChampsEntêteLignesClef';

    sArgument     := GetNomChampsFromPopMenu(False, sParametreYFOActif);
    sChampsActifs := GetNomChampsFromPopMenu(True , sParametreYFOActif);
    while (sArgument<>'') do
    begin
      sYTSChamp  := ReadTokenSt(sArgument);
      sArgChamp := StringReplace(sYTSChamp, 'YTS_', '', [rfReplaceAll]);
      if (pos(sArgChamp,sChampsActifs)>0) then
      begin
        if (pos(sYTSChamp,'YTS_DATEDEBUT/YTS_DATEFIN')>0) then
          Result := Result + sArgChamp +'='+ usDateTime(GetDateTime(sYTSChamp))+';'
        else
          Result := Result + sArgChamp +'='+ iif(GetString(sYTSChamp)='-','',GetString(sYTSChamp)) +';';
      end
      else
        Result := Result + sArgChamp +'='+ ';';
    end;
  end; {GetArgumentYTS}
  {$ENDIF GCGC}

  function GetCtxWithPiece: String;
  begin
    Result := 'WITHPIECE=' + BoolToStr_((GetString (GetPrefixe + '_NATUREPIECEG') <> '') and
                                       (GetString (GetPrefixe + '_SOUCHE')       <> '') and
                                       (GetInteger(GetPrefixe + '_NUMERO')       <> 0 ))
  end;

begin
  Result := inherited GetLoupeCtx;
  if GetPrefixe = 'WOL' then
  begin
    Result := 'ORDREPERE=' + GetString('WOL_ORDREPERE')
            + ';' + GetCtxWithPiece
  end
  else if GetPrefixe = 'WOT' then
  begin
    Result := GetCtxWithPiece
  end
  else if GetPrefixe = 'WOP' then
  begin
    TypeOrdre:= wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOP_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOP_LIGNEORDRE'));
    Result:= iif((Pos('[' + TypeOrdre + ']', sCannotDoActionOnTypeOrdre) = 0) and (GetString('WOP_TIERS') <> ''),';TYPEORDRE=STP','')
            + ';ETATPHASE=' + GetString('WOP_ETATPHASE');
  end
  else if GetPrefixe = 'WOG' then
  begin
    Result := 'ETATOPE=' + GetString('WOG_ETATOPE')
  end
  else if GetPrefixe = 'WOR' then
  begin
    Result := 'ETATOPE=' + wGetSqlFieldValue('WOG_ETATOPE', 'WORDREGAMME', 'WOG_NATURETRAVAIL="' + GetString('WOR_NATURETRAVAIL') + '" AND WOG_LIGNEORDRE=' + IntToStr(GetInteger('WOR_LIGNEORDRE')) + ' AND WOG_OPECIRC="' + GetString('WOR_OPECIRC') + '" AND WOG_NUMOPERGAMME=' + IntToStr(GetInteger('WOR_NUMOPERGAMME')) )
  end
  {$IFDEF GCGC}
    else if GetPrefixe = 'YTS' then
    begin
      Result := 'APPEL=TARIF; '+GetArgumentYTS;
    end
  {$ENDIF GCGC}
  else if GetPrefixe = 'GVP' then
  begin
    Result := 'APPEL=VALORISATION;';
    Result := Result + 'FAMILLEVALO='+GetString('GVP_FAMILLEVALO')+';';
    Result := Result + 'FAMILLENIV1='+GetString('GVP_FAMILLENIV1')+';';
    Result := Result + 'FAMILLENIV2='+GetString('GVP_FAMILLENIV2')+';';
    Result := Result + 'FAMILLENIV3='+GetString('GVP_FAMILLENIV3')+';';
  end
  {$IFDEF STK}
    else if GetPrefixe = 'GST' then
    begin
      Result := 'ARTICLE='       + GetString('GST_ARTICLE')
              + ';LOTINTERNE='   + GetString('GST_LOTINTERNE')
              + ';SERIEINTERNE=' + GetString('GST_SERIEINTERNE')
              + ';TIERS='        + GetString('GST_TIERS')
    end
    else if GetPrefixe = 'GSM' then
    begin
      Result := 'ARTICLE='       + GetString('GA_ARTICLE')
              + ';LOTINTERNE='   + GetString('GSM_LOTINTERNE')
              + ';SERIEINTERNE=' + GetString('GSM_SERIEINTERNE')
    end
    else if GetPrefixe = 'GQD' then
    begin
      Result := ';ARTICLE='    + GetString('GQD_ARTICLE')
              + ';LOTINTERNE=' + GetString('GQD_LOTINTERNE')
              + ';SERIEINTERNE=' + GetString('GQD_SERIEINTERNE')
    end
  {$ENDIF STK}
  {$IFDEF QUALITE}
  else if GetPrefixe = 'RQN' then
  begin
    Result := ';ARTICLE='     + GetString('RQN_ARTICLE')
            + ';TIERS=' 		  + GetString('RQN_TIERS')
            + ';LIGNEORDRE='  + GetString('RQN_LIGNEORDRE')
            + ';NATURETRAVAIL='+ GetString('RQN_NATURETRAVAIL')
            + ';CLEPIECE='	  + GetString('RQN_PIECEORIGINE')
  end
  else if GetPrefixe = 'RQD' then
  begin
    Result := ';TIERS=' 		 + GetString('RQD_TIERS')
            + ';QNCNUM='		 + GetString('RQD_QNCNUM')
            + ';IDACTION='	 + GetString('RQD_IDACTION')
  end
  else if GetPrefixe = 'RAC' then
  begin
    Result := ';DEMDEROGNUM='+ GetString('RAC_QDEMDEROGNUM')
            + ';PLANCORRNUM='+ GetString('RAC_QPLANCORRNUM')
            + ';QNCNUM='		 + GetString('RAC_QNCNUM')
  end
  {$ENDIF QUALITE}
end;
{$ENDIF !PGIMAJVER}
                 
Initialization
	RegisterClasses([tWTom]);
end.
