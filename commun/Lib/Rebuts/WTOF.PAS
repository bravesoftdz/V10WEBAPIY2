{***********UNITE*************************************************
Auteur  ...... : TP; JLS
Créé le ...... : 26/09/2001
Description .. : TOF Métier
Mots clefs ... : WTOF
*****************************************************************}
Unit wTof;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
    db,
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    dbGrids,
    Mul,
    Fiche,
    Fe_Main,
    QRS1,
    MenuOLG,
  {$ELSE}
    MainEagl,
    eMul,
    eFiche,
    eQRS1,
    MenuOLX,
  {$ENDIF}
  Cube, 
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  hdb,
  UTOF,
  HTB97,
  HQry,
  Menus,
  Dialogs,
  AglInit,
  SaisieList,
  dbCtrls,
  ParamSoc,
  Windows,
  uTob,
  wMnu,
  Vierge,  
  GRS1,
  graphics,
  wCommuns,
  stat,
  Messages,
  uTreeTobFrame,
  UtilConfid {JTR - Confidentialité sur la fiche} 
  ;

Type
  MyArrayInt = Array of Integer;


  tWTOF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
  private
    FTabWidths          : MyArrayInt;
    FDBListe            : String;
    FViergeGridName     : String;
    FEditColsProperties : Boolean;
    FWListe             : TWListe;
    FTreeTobFrame       : TFFrameTreeTob;

    procedure PmUtilitaire_OnPopUp(Sender: TObject);
    procedure PmAnalyse_OnPopUp(Sender: TObject);
    procedure PmHistorique_OnPopUp(Sender: TObject);
    {$IFDEF GPAOLIGHT}
      procedure PmAction_OnPopUp(Sender: TObject);
      procedure PmBaseTech_OnPopUp(Sender: TObject);
      procedure MnJournalAction_OnClick(Sender: TObject);
    {$ENDIF GPAOLIGHT}

    procedure MnProperties_OnClick(Sender: TObject);
    procedure MnLpBlocNote_OnClick(Sender: TObject);

    { Dimensions }
    procedure GA_STATUTART_OnChange(Sender: TObject);
    procedure GA_GRILLEDIM_OnChange(Sender: TObject);

    { Accès au mul ou au lookup article }
    procedure ScanControls;
    procedure ArticleElipsisClick(Sender: TObject);

    { Get }
    function GetChampsMUL: String;

    { Gestion des colonnes dans les muls }
    procedure LoadColsList;
    function GetColWidths: MyArrayInt;
    {$IFNDEF EAGLCLIENT}
      procedure DBGrid_OnColMoved(Sender: TObject; FromIndex, ToIndex: Longint);
    {$ENDIF !EAGLCLIENT}
    procedure SetEditColsProperties(const Value: Boolean);

    { Particularité pour les date de décomposition }
    procedure BTEFFACEDATEDEC_OnClick(Sender: tObject);

    { Gestion des ruptures dans les QR1 }
    procedure RUPT_OnClick(Sender: TObject);
    procedure SAUTRUPT_OnClick(Sender: TObject);
    procedure ChangeGroup(Const Indice: integer);
    procedure AffectGroup(Const Indice: integer);
    procedure ChangeSautPage(Const Indice: integer);
    function GetMultiSelected: Boolean;
    function GetMultiSelection: Boolean;
    procedure SetMultiSelection(const Value: Boolean);
  public
    fNature, fFiche, fTableName			   										              : String;
    fMessageToDelete, fChamp1OfMessageToDelete, fChamp2OfMessageToDelete: String;
    fRange, fLequel, fParamsLanceFiche, fBlobFieldSuffixe								: String;
    StArgument, FSuffixeUserForBtUser                                   : String;
    fMulDeRecherche, fNoElipsisArticle, FBlocNoteObligatoire, wMonoFiche: Boolean;
    wSelectAll: Boolean;

    { Get }
    function GetTableName: string;
    function GetWhere: string;
    function GetQuery: tQuery;
    function GetPrefixe: string;
    function GetInteger(Const FieldName: string): integer;
    function GetString(Const FieldName: string): string;
    function GetDateTime(Const FieldName: string): tDateTime;
    function GetDouble(Const FieldName: string): Double;
    function GetBoolean(Const FieldName: string): Boolean;

    { Boutons }
    procedure BINSERT_OnClick(Sender: TOBject);
    procedure BSELECTALL_OnClick(Sender: TObject);
    procedure FLISTE_OnDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BDELETE_OnClick(Sender: TObject);
    procedure BTUSER_OnClick(Sender: TObject);
    procedure BRECHAFFAIRE_OnClick(Sender: TObject);
    procedure BSUPPRIME_OnClick(Sender: TObject);
    procedure BTSHOWBLOCNOTE_OnClick(Sender: tObject);

    { Set }
    procedure SetMemoFrame;
    procedure SetControlsVisible(Const FieldsName: Array of String; Const Visibility: Boolean);
    procedure SetControlsEnabled(Const FieldsName: Array of String; Const Enability: Boolean);
    procedure SetColVisible(const ColName: String; const Visible: Boolean);
    procedure SetColsVisible(const FieldsName: Array of String; const Visible: Boolean);
    procedure SetColCaption(const ColName, Caption: String);
    procedure SetColsCaption(const FieldsName: Array of String; const Caption: String);
    procedure SetLibelle;
    procedure SetDomaine;
    procedure SetWhereDomaine;
    procedure SetDepot;
    procedure SetEtablissement;

    { Divers }
    procedure ReloadColsList;
    procedure wGotoNewLine(Save: String);
    procedure RefreshDB;
    procedure wOuvreMul(Const Fiche, TableFille, Range, Param: string);
    function RechLib(sdataType, sFieldname : string; lAbrege : Boolean; lNotCheckBox : Boolean; Plus : string = '') : string;
    function wOuvreFiche(Const Fiche, Lequel: string; Action: string; Const Param: string): String;
    function IsEmpty: Boolean;
    function EcranIs(const FormName: String): Boolean;
    function MakeWhere: String;

    { Dimensions }
    procedure SetDimensions;
    procedure SetArticle;

     { BCherche }
    procedure BCherche_OnClick(Sender: TOBject);

    { Gestion des colonnes }
    property EditColsProperties: Boolean read FEditColsProperties write SetEditColsProperties;
    property wListe: TWListe read FWListe;
    property TreeTobFrame: TFFrameTreeTob read FTreeTobFrame write FTreeTobFrame;
    property MultiSelected: Boolean read GetMultiSelected;
    property MultiSelection: Boolean read GetMultiSelection write SetMultiSelection;
    property SuffixeUserForBtUser: String read FSuffixeUserForBtUser write FSuffixeUserForBtUser;
  protected
    Action: string;
    Droit : string;

    { Gestion des colonnes }
    property ChampsMUL: String read GetChampsMUL;
    property ViergeGridName: String read FViergeGridName write FViergeGridName;

    { Loupe }
    procedure PmLoupe_OnPopUp(Sender: TObject);
    { -------------------- }
    procedure RUPT_OnEnter(Sender: TObject); 
  end ;

Implementation

uses
  AGLInitGC,
  Grids,
  {$IFDEF GCGC}
    yTarifs,
    yTarifsCommun,
  {$ENDIF GCGC}
  {$IFDEF GPAOLIGHT}
    wNatureTravail,
    wGereForms,
    wJournalAction,
    wOrdrePhase,
    {$IFNDEF GPAO}
      GereDiff_GC_GP,
    {$ENDIF GPAO}
  {$ENDIF GPAOLIGHT}
  {$IFDEF GPAO}
    wAction,
    EntGP,
    wPrixDeRevient,
  {$ENDIF GPAO}
  wProperties_Tof,
  wBlocNote_Tof,
  {$IFDEF STK}
    EntGC,
  {$ENDIF STK}
  FactUtil,
  UtilArticle,
  UtilGC,
  Ent1
  ;

function NombreDeToken ( s: String ) : Integer ;
begin
	Result := 0;
	while ReadTokenSt(S) <> '' do
    Inc(Result);
end ;

procedure tWTOF.OnNew ;
begin
	Inherited;
end;

procedure tWTOF.OnDelete ;
begin
	Inherited;
end;

procedure tWTOF.OnUpdate ;
begin
	Inherited;

  { Pour être sûr de réassigner les événements }
  if Assigned(Ecran) and (Ecran is TFMul) and FMulDeRecherche and (GetTableName <> '') and (not MultiSelection) then
  begin
    if Assigned(GetControl('FLISTE')) then
    begin
      THGrid(GetControl('FLISTE')).OnDblClick := FLISTE_OnDblCLick;
      Ecran.OnKeyDown := FormKeyDown;
    end;
    if Assigned(GetControl('BOUVRIR')) then
      TToolBarButton97(GetControl('BOUVRIR')).OnClick := FLISTE_OnDblCLick;
  end;

  { -------- Gestion des colonnes -------- }
  if FEditColsProperties and not (Ecran is TFVierge) then
  begin
    {$IFDEF EAGLCLIENT}
      FTabWidths := GetColWidths;
    {$ENDIF EAGLCLIENT}
    { Chargement de la liste des champs }
    if FViergeGridName <> 'FListe' then
      FTabWidths := GetColWidths;
    LoadColsList
  end;

  { Libellés paramétrés}
  if ecran is TFMUL then
    SetLibelle;

  { Dimensions }
  SetDimensions;

  {$IFDEF GPAOLIGHT}
  wGereFormsByCtx(Ecran, Self, 'OnUpdate');
  {$ENDIF GPAOLIGHT}

  SetControlProperty('BSUPPRIME'   , 'ENABLED', (Action = 'MODIFICATION') and (not IsEmpty));
  SetControlProperty('BDELETE'     , 'ENABLED', (Action = 'MODIFICATION') and (not IsEmpty));
  SetControlProperty('BSELECTALL'  , 'ENABLED', (Action = 'MODIFICATION') and (not IsEmpty));
  SetControlProperty('BTDUPLIC'    , 'ENABLED', (Action = 'MODIFICATION') and (not IsEmpty));
  SetControlProperty('BDUPLICATION', 'ENABLED', (Action = 'MODIFICATION') and (not IsEmpty));

//===========================================================================
//  MERCI DE LAISSER CET APPEL DE FONCTION EN DERNIER
//===========================================================================
  {$IFDEF GPAOLIGHT}
    {$IFNDEF GPAO}
    // on va cacher ou rendre disabled certains champs
    AppliqueModifAssemblage(Ecran, Self, 'OnUpdate');
    {$ENDIF GPAO}
  {$ENDIF GPAOLIGHT}

end;

procedure tWTOF.OnLoad ;
begin
	Inherited;
  if Action = 'CONSULTATION' then
    SetControlVisible('BSELECTALL', False);

  { Gestion des colonnes }
  if FEditColsProperties and (Ecran is TFVierge) then
  begin
    FTabWidths := GetColWidths;
    { Chargement de la liste des champs }
    LoadColsList
  end;

  { Libellé article }
  if Assigned(GetControl('GA_CODEARTICLE')) then
    THEdit(GetControl('GA_CODEARTICLE')).UpdateLibelle;

  { Déduction du domaine }
  SetWhereDomaine;
end;

procedure tWTOF.OnArgument (S : String ) ;
var
  c: char;
begin
  { Paramètre }
  StArgument := S;
  FViergeGridName := 'FListe';
  FEditColsProperties := False;
  FWListe := nil;
  FDBListe := '';
  FSuffixeUserForBtUser := 'CREATEUR';

  { Action }
  Action := GetArgumentValue(S, 'ACTION');
  Droit  := GetArgumentValue(S, 'DROIT');

  { Nature par défaut des fiches }
  if fNature = '' then
    fNature := 'W';

  { Propriété par défaut }
  wSelectAll := False;
  wMonoFiche := False;
  if Ecran is tfMul then
    TFMUL(ECRAN).CheckBoxStyle := CsCoche;

  { Gestion des dimensions: Début }
  { Filtre par défaut }
  if Assigned(GetControl('GA_STATUTART')) then
  begin
    ThMultiValComboBox(GetControl('GA_STATUTART')).Text := 'GEN;UNI;';
    ThMultiValComboBox(GetControl('GA_STATUTART')).OnChange := GA_STATUTART_OnChange;
  end;

  { GrillesDim }
  for c := '1' to '5' do
  begin
    if Assigned(GetControl('GA_GRILLEDIM' + c)) then
      ThValComboBox(GetControl('GA_GRILLEDIM' + c)).OnChange := GA_GRILLEDIM_OnChange;
  end;
  { Gestion des dimensions: Fin }

  FBlobFieldSuffixe := 'BLOCNOTE';
  FBlocNoteObligatoire := False;

  {$IFDEF GPAOLIGHT}
  wGereFormsByCtx(Ecran, Self, 'OnArgument', True);
  {$ENDIF GPAOLIGHT}

  Inherited;

  { Evénements standard }
  if (Ecran is TFMul) and not fMulDeRecherche then
  begin
    if Assigned(GetControl('FLISTE')) then
      THGrid(GetControl('FLISTE')).OnDblClick := FLISTE_OnDblCLick ;
    if not (Ecran is TFSaisieList) and Assigned(GetControl('BINSERT')) then
      TToolBarButton97(GetControl('BINSERT')).OnClick := BINSERT_OnClick ;
    if not (Ecran is TFSaisieList) and Assigned(GetControl('BDELETE')) then
      TToolBarButton97 (GetControl('BDELETE')).OnClick := BDELETE_OnCLick ;
    if Assigned(GetControl('BSELECTALL')) then
      TToolBarButton97(GetControl('BSELECTALL')).OnClick := BSELECTALL_OnCLick ;
  end;

  if (Ecran is TFMul) or (Ecran is TFCube) then
  begin
    if Assigned(GetControl('BTUSER')) then
      TToolBarButton97(GetControl('BTUSER')).OnClick := BTUSER_OnCLick ;
    if Assigned(GetControl('BRECHAFFAIRE')) then
      TToolBarButton97(GetControl('BRECHAFFAIRE')).OnClick := BRECHAFFAIRE_OnCLick ;
  end;

  { La propriété n'est pas gérée dans les Tof Métiers }
  { Géré ici pour compatibilité avec les Tof métiers actuelles }
  { ...A voir pour opération : virer la notion de FMulDeRecherche dans les Tof métier pour n'être géré qu'ici }
  if Ecran is TFMul then
  begin
    if not FMulDeRecherche then
    begin
      FMulDeRecherche := GetArgumentBoolean(S, 'RECHERCHE');
      MultiSelection := MultiSelection and (not FMulDeRecherche);
    end;

    if FMulDeRecherche and (GetTableName <> '') and Assigned(GetControl(FViergeGridName)) then
      THGrid(GetControl(FViergeGridName)).OnDblClick := FLISTE_OnDblCLick
  end;

  { Loupe }
  if Assigned(GetControl('PMLOUPE')) then
    TPopUpMenu(GetControl('PMLOUPE')).OnPopUp := PmLoupe_OnPopUp;
  if Assigned(GetControl('MNPROPERTIES')) then
    TMenuItem(GetControl('MNPROPERTIES')).OnClick := MnProperties_OnClick;
  if Assigned(GetControl('MNLPBLOCNOTE')) then
    TMenuItem(GetControl('MNLPBLOCNOTE')).OnClick := MnLpBlocNote_OnClick;

  { Historique }
    if Assigned(GetControl('PMHISTORIQUE')) then
      TPopUpMenu(GetControl('PMHISTORIQUE')).OnPopUp := PmHistorique_OnPopUp;

  { Analyse }
    if Assigned(GetControl('PMANALYSE')) then
      TPopUpMenu(GetControl('PMANALYSE')).OnPopUp := PmAnalyse_OnPopUp;

  { Utilitaire }
    if Assigned(GetControl('PMUTILITAIRE')) then
      TPopUpMenu(GetControl('PMUTILITAIRE')).OnPopUp := PmUtilitaire_OnPopUp;

  { Action }
  {$IFDEF GPAOLIGHT}
    if Assigned(GetControl('PMACTION')) then
      TPopUpMenu(GetControl('PMACTION')).OnPopUp := PmAction_OnPopUp;
    if Assigned(GetControl('PMBASETECH')) then
      TPopUpMenu(GetControl('PMBASETECH')).OnPopUp := PmBaseTech_OnPopUp;
    if Assigned(GetControl('MNJOURNALACTION')) then
      TMenuItem(GetControl('MNJOURNALACTION')).OnClick := MnJournalAction_OnClick;
  {$ENDIF GPAOLIGHT}

  { Rendre invisible le bouton 'LOUPE' et 'LIGNES' en création }
  if Assigned(GetControl('BTLOUPE')) then
    SetControlVisible('BTLOUPE', Action <> '');
  if Assigned(GetControl('BTLIGNES')) then
    SetControlVisible('BTLIGNES', Action <> '');
  { Rendre invisible en fonction du contexte }
  if Assigned(GetControl('BTUTILITAIRE')) then
    SetControlVisible('BTUTILITAIRE', Action <> 'CONSULTATION');

  { Gestion des droits sur les boutons Création, Duplication et Suppression }
  if Droit <> '' then
  begin
    if Assigned(GetControl('BINSERT')) then
      SetControlVisible('BINSERT', (Action = '') or (pos('C', Droit) > 0));
    if Assigned(GetControl('BDUPLICATION')) then
      SetControlVisible('BDUPLICATION', (Action = '') or (pos('D', Droit) > 0));
    if Assigned(GetControl('BDELETE')) then
      SetControlVisible('BDELETE'     , (Action = '') or (pos('A', Droit) > 0));
  end;

  if  (not( ecran is TFSaisieList )) and (Assigned(GetControl('BCHERCHE'))) then
    TToolBarButton97(GetControl('BCHERCHE')).OnClick := BCHERCHE_OnCLick ;
  if Assigned(GetControl('BTSHOWBLOCNOTE')) then
    TToolBarButton97(GetControl('BTSHOWBLOCNOTE')).OnClick := BTSHOWBLOCNOTE_OnClick;

  { Bouton Elipsis recherche Article }
  if not fNoElipsisArticle then
    ScanControls;

  { BlocNote }
  if Assigned(GetControl('SPLITPNMEMO')) then
  begin
    ThSplitter(GetControl('SPLITPNMEMO')).Height := 3;
    ThSplitter(GetControl('SPLITPNMEMO')).Beveled := true;
  end;


  { 'Menu Zoom' en dur }
  if Assigned(GetControl('BTLOUPE')) then
  begin
    with TToolBarButton97(GetControl('BTLOUPE')) do
    begin
      Hint := TraduireMemoire('Menu zoom');
      ShowHint := True;
    end
  end;

  { Gestion des ruptures dans les QR1 }
  for c := '1' to '6' do
  begin
    if Assigned(GetControl('RUPT' + c )) then
    begin
      THValComboBox(GetControl('RUPT' + c )).OnEnter  := RUPT_OnEnter;
      THValComboBox(GetControl('RUPT' + c )).OnChange := RUPT_OnClick;
      THValComboBox(GetControl('RUPT' + c )).OnClick := RUPT_OnClick;
      TCheckBox(GetControl('SAUTRUPT' + c )).OnClick  := SAUTRUPT_OnClick;
    end;
  end;

  { Efface la date de décomposition }
  if Assigned(GetControl('BTEFFACEDATEDEC')) then
    tToolBarButton97(GetControl('BTEFFACEDATEDEC')).OnClick := BTEFFACEDATEDEC_OnClick;

  { Contrôle des boutons en fonction du statut d'ouverture de la fiche }
  SetControlVisible('BSUPPRIME'    , (Action = 'MODIFICATION') and GetControlVisible('BSUPPRIME'));
  SetControlVisible('BDELETE'      , (Action = 'MODIFICATION') and GetControlVisible('BDELETE'));
  SetControlVisible('BTDUPLIC'     , (Action = 'MODIFICATION') and GetControlVisible('BTDUPLIC'));
  SetControlVisible('BTDUPLICATION', (Action = 'MODIFICATION') and GetControlVisible('BTDUPLICATION'));
  SetControlVisible('BSELECTALL'   , (Action = 'MODIFICATION') and GetControlVisible('BSELECTALL'));

  { Hint imposés }
  SetControlProperty('BSUPPRIME'   , 'Hint', TraduireMemoire('Supprimer la sélection'));
  SetControlProperty('BDELETE'     , 'Hint', TraduireMemoire('Supprimer la sélection'));
  SetControlProperty('BTUSER'      , 'Hint', TraduireMemoire('Renseigner avec l''utilisateur connecté'));
  SetControlProperty('BRECHAFFAIRE', 'Hint', TraduireMemoire('Recherche affaire'));

  { Libellés paramétrés}
  if (Ecran is TFQRS1) or (Ecran is TFStat) then
    SetLibelle;

  { Pas de filtre par défaut }
  if Ecran is TFMul then
    Tfmul(Ecran).FiltreDisabled := (Pos('FILTREDISABLED', S) > 0);

  { Domaine; Dépôt; Etablissement par défaut }
  SetDomaine;
  SetDepot;
  SetEtablissement;

  { Confidentialité }
  AppliquerConfidentialite(Ecran, ''); {JTR - Confidentialité sur la fiche}

//===========================================================================
//  MERCI DE LAISSER CET APPEL DE FONCTION EN DERNIER
//===========================================================================
{$IFDEF GPAOLIGHT}
  {$IFNDEF GPAO}
  // on va cacher ou rendre disabled certains champs
  AppliqueModifAssemblage(Ecran, Self, 'OnArgument');
  {$ENDIF GPAO}
{$ENDIF GPAOLIGHT}
end ;

procedure tWTOF.OnClose ;
begin
  Inherited ;

  if FEditColsProperties and Assigned(FWListe) then
    FWListe.Free
end ;

procedure TWTOF.BINSERT_OnClick ( Sender: TOBject ) ;
begin
  if FFiche <> '' then
    wOuvreFiche(FFiche, '', 'ACTION=CREATION;MONOFICHE', FParamsLanceFiche);
end ;

procedure TWTOF.FLISTE_OnDblClick ( Sender: TOBject ) ;
var
  Lequel, sClef: String;
begin
  { Abort }
  if (FFiche = '') or (isEmpty) then Abort;

  if not FMulDeRecherche then
  begin
    { Lequel }
{$IFDEF CONVU}
    if Assigned(FTreeTobFrame) then
      Lequel := wGetValueClef1(FTreeTobFrame.CurrentTob)
    else
{$ENDIF CONVU}
      Lequel := WFabriqueRange(FTableName, FLequel, GetDataSet);

    { Ajout du Lequel }
    if FParamsLanceFiche = '' then
      FParamsLanceFiche := '[' + Lequel + ']';

    wOuvreFiche(FFiche, Lequel, iif(Pos('ACTION', fParamsLanceFiche) > 0, '', 'ACTION=' + iif(Action = '', 'MODIFICATION', Action)), FParamsLanceFiche)
  end
  else
  begin
    TFMul(Ecran).Retour := '';
    sClef := wMakeFieldString(GetTableName, ';');
    while sClef <> '' do
    begin
      if (sClef = 'GA_ARTICLE') and GetArgumentBoolean(StArgument, 'RETOUR_CODEARTICLE') then
        sClef := 'GA_CODEARTICLE';
      TFMul(Ecran).Retour := TFMul(Ecran).Retour + iif(TFMul(Ecran).Retour <> '', ';', '') + GetString(ReadTokenSt(sClef));
    end;
    TFMul(Ecran).Close
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier. PERSOUYRE
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : Cette procédure permet de se repositionner sur le bon
Suite ........ : enregistrement.
Suite ........ : En Paramètre, les valeurs de la clef qui a été traité en
Suite ........ : dernier.
Mots clefs ... :
*****************************************************************}
procedure TWTOF.wGotoNewLine(Save: String);
var
  RV            : Array of Variant ;
  S, St, sClef1 : String ;
  Indice        : Integer ;

  {$IFDEF EAGLCLIENT}
  DS    : TDataSet;
  iRec  : integer;
  found : boolean;
  {$ENDIF}
begin
  { Récupération des champs de la clef 1 }
  if fLequel <> '' then
    sClef1 := fLequel
  else
  begin
  	sClef1 := wGetFieldsClef1(FTableName);
	  sClef1 := StringReplace(sClef1, ',', ';', [ rfReplaceAll ]);
  end;

	if NombreDeToken(sClef1) = 1 then
  begin
    {$IFDEF EAGLCLIENT}
      DS := GetDataSet;
      DS.First;
      while (not DS.eof) and (Save <> DS.FindField(sCLef1).value) do
      begin
        DS.Next;
      end;
      if (DS.CurrentFilleIndex + 1) < TFMul(Ecran).FListe.RowCount then
        TFMul(Ecran).FListe.Row := DS.CurrentFilleIndex + 1
      else
        TFMul(Ecran).FListe.Row := TFMul(Ecran).FListe.RowCount-1;
    {$ELSE}
		  GetDataSet.Locate(sClef1, Save, [ LoCaseInsensitive ])
    {$ENDIF}
  end
	else
  begin
		SetLength ( RV, 0 ) ;
		S      := sCLef1 ;
		Indice := 0 ;
		St     := ReadTokenSt ( S ) ;
		while St <> '' do
		begin
			Inc ( Indice ) ;
			SetLength ( RV, Indice ) ;
			RV[Indice-1] := WConvertStringToField ( FTableName, St, ReadTokenSt ( Save ) ) ;
			St := Trim(ReadTokenSt ( S )) ;
		end ;
    {$IFDEF EAGLCLIENT}
      DS := GetDataSet;  
      iRec := DS.CurrentFilleIndex;
      DS.First;
      Found := false;
      while not DS.eof do
      begin
        S:= sCLef1 ;
        Indice := 0 ;
        St:= Trim(ReadTokenSt ( S ));
        Found := true;
        while St <> '' do
        begin
          Inc ( Indice ) ;
          if RV[Indice-1]<>DS.FindField(St).value then
          begin
            Found := False;
            Break;
          end;
          St := Trim(ReadTokenSt ( S )) ;
        end ;
        if Found then Break;
        DS.Next;
      end;
      if not Found then
        DS.Seek( iRec )
      else
      begin
        if (Ecran<>nil) And (Ecran is TFMul) then
        begin
          if (DS.CurrentFilleIndex + 1) < TFMul(Ecran).FListe.RowCount then
            TFMul(Ecran).FListe.Row := DS.CurrentFilleIndex + 1
          else
            TFMul(Ecran).FListe.Row := TFMul(Ecran).FListe.RowCount-1;
        end
      end;
    {$ELSE  EAGLCLIENT}
  		GetDataSet.Locate ( sClef1, RV, [ LoCaseInsensitive ] ) ;
    {$ENDIF EAGLCLIENT}
		SetLength ( RV, 0 ) ;
	end ;
end ;

procedure TWTOF.BDELETE_OnClick ( Sender: TOBject ) ;
var
  Save, S, St   : String;
  Champ1, Champ2: String;
  DS            : TDataSet;
  R             : Integer;
  RS            : Array of String;
  RV            : Array of Variant;
  IndiceRS      : Integer;
begin
  { Init des champs }
	Champ1 := '' ;
	Champ2 := '' ;
	if FChamp1OfMessageToDelete <> '' then Champ1 := GetString(FChamp1OfMessageToDelete);
	if FChamp2OfMessageToDelete <> '' then Champ2 := GetString(FChamp2OfMessageToDelete);

	{ Récupération du DataSet }
	DS := GetDataSet ;

  {$IFDEF EAGLCLIENT}
  DS.Next; {PMJEAGL}
  if DS.Eof then
  {$ELSE}
	if Not DS.FindNext then
  {$ENDIF}
  begin
	  { Est-ce que je suis au début ? }
    {$IFDEF EAGLCLIENT}
    DS.Prior; {PMJEAGL}
    if DS.BOF then
    {$ELSE}
		if not DS.FindPrior then
    {$ENDIF}
		begin
				{ Je suis donc le seul, sniff }
			Save := '' ;
		end
		else
		begin
			{ Récupération des clefs }
			Save := WGetValueClef1 ( FTableName, TFMUL(ECRAN), fLequel ) ;
			{ Revenir sur le suivant }
		  DS.Next ;
		end;
	end
	else
	begin
	  { Récupération des clefs }
		Save := WGetValueClef1 ( FTableName, TFMUL(ECRAN), fLequel ) ;
		{ Revenir sur le Précédent }
		DS.Prior ;
	end;

	R := HShowMessage(FMessageToDelete + ';Q;YN;N;N', Champ1, Champ2);
	if R = mrYes then
  begin
	  { Création de l'array des champs }
		IndiceRS := 0 ;
		S := WGetFieldsClef1 ( FTableName ) ;
		S := StringReplace ( S, ',', ';', [ rfReplaceAll ] ) ;
		St := ReadTokenSt ( S ) ;
		while St <> '' do
		begin
			Inc ( IndiceRS ) ;
			SetLength ( RS, IndiceRS ) ;
			RS[IndiceRS-1] := St ;
			St := ReadTokenSt ( S ) ;
		end ;

		{ Création de l'array des valeurs des champs }
		IndiceRS := 0 ;
		S := WGetFieldsClef1 ( FTableName ) ;
		S := StringReplace ( S, ',', ';', [ rfReplaceAll ] ) ;
		St := ReadTokenSt ( S ) ;
		while St <> '' do
		begin
			Inc ( IndiceRS ) ;
			SetLength ( RV, IndiceRS ) ;
			RV[IndiceRS-1] := GetField ( St ) ;
			St := ReadTokenSt ( S ) ;
		end ;

		WDeleteSQL(FTableName, RS, RV);

		SetLength(RS, 0);
		SetLength(RV, 0);

		{ Raffraichissement }
		TFMUL(ECRAN).BChercheClick ( Nil ) ;

		{ Repositionnement }
		wGotoNewLine(Save);
  end ;
end ;

function tWTOF.WOuvreFiche(Const Fiche, Lequel: string; Action: string; Const  Param: string): String;
var
	Argument: String;
begin
  { Action }
  if (Action <> '') and (Pos('ACTION', Action) <= 0) then
    Action := 'ACTION=' + Action;

  { Argument }
  Argument := Action;
  if wMonoFiche then
    Argument := Argument + ';MONOFICHE';
  Argument := iif((Argument <> '') and (Param <> ''), Argument + ';', Argument);
  Argument := iif(Param <> '', Argument + Param, Argument);

  TheMulQ := Nil;

  {$IFDEF EAGLCLIENT}
  if ECRAN is TFMul then
    TheMulQ := TFMul(ECRAN).Q.TQ;
  {$ELSE}
  if ECRAN is TFMul then
    TheMulQ := TFMul(ECRAN).Q;
  {$ENDIF}

  Result := AglLanceFiche (fNature, Fiche, '', Lequel, Argument);
  if Result <> '' then
  begin
    if Assigned(TFMUL(ECRAN)) and (ECRAN is TFMul) then
    begin
      TFMUL(ECRAN).BChercheClick (Nil);
      { Repositionnement }
      WGotoNewLine(Result);
    end;
  end
  else
  begin
    RefreshDB;
  end;

  { Si le mémo est à afficher }
  SetMemoFrame;
end;

procedure tWTOF.WOuvreMul(Const Fiche, TableFille, Range, Param: string);
var
	PrefixeMere, PrefixeFille, MyRange, MyParam: String;
begin
  if TableFille <> '' then
  begin
    { Récupère les préfixes des tables pour le range }
    PrefixeMere := TableToPrefixe(fTableName);
    PrefixeFille := TableToPrefixe(TableFille);
    MyParam := Param;

    { Fabrique le range du Mul }
    MyRange := WMakeRange(PrefixeFille, PrefixeMere, True, WGetNbrFieldInKey(fTableName), GetDataSet);
    if (Pos('[', Range) = 0) and (Pos(']', Range) = 0) then
    begin
      { Fabrique les paramètres à passer au mul }
      MyParam := '[' + WMakeRange(PrefixeFille, PrefixeMere, False, WGetNbrFieldInKey(fTableName), GetDataSet) + ']';
      if Param <> '' then MyParam := MyParam + ';' + Param;
    end;
    { Ouvre la fiche }
    AglLanceFiche('W', Fiche, MyRange, '', MyParam);
  end
  else
    AglLanceFiche('W', Fiche, Range, '', Param);
end;

procedure tWTOF.RefreshDB;
var
	Save: String ;
begin
  if Assigned(Ecran) and (Ecran is TFMul) then
  begin
    if isEmpty then
    begin
      TFMUL(Ecran).BChercheClick(Nil);
    end
    else
    begin
      Save := WGetValueClef1(fTableName, TFMUL(Ecran), fLequel);
      TFMUL(Ecran).BChercheClick(Nil);
      wGotoNewLine(Save);
    end;

    { On remet à Down le bouton de multi-sélection }
    if wSelectAll then
      SetControlProperty('BSELECTALL', 'DOWN', false);
  end
  else if (Ecran is TFVierge) and Assigned(FTreeTobFrame) then
    FTreeTobFrame.RefreshTreeTob
end;

function tWTOF.GetDateTime(Const FieldName: string): tDateTime;
begin
  if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetDateTime(FieldName)
  else if GetField(FieldName) <> null then
   	Result := GetField(FieldName)
  else
   	Result := iDate1900;
end;

function tWTOF.GetDouble(Const FieldName: string): Double;
begin
	if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetDouble(FieldName)
  else if GetField(FieldName) <> null then
   	Result := GetField(FieldName)
  else
   	Result := 0;
end;

function tWTOF.GetInteger(Const FieldName: string): integer;
begin
  if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetInteger(FieldName)
  else if GetField(FieldName) <> null then
   	Result := GetField(FieldName)
  else
   	Result := 0;
end;

function tWTOF.GetString(Const FieldName: string): string;
begin
  if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetString(FieldName)
	else if GetField(FieldName) <> null then
   	Result := GetField(FieldName)
  else
   	Result := '';
end;

function tWTOF.GetBoolean(Const FieldName: string): Boolean;
begin
  if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetBoolean(FieldName)
  else if GetField(FieldName) <> null then
    Result := StrToBool(GetField(FieldName))
  else
  	Result := False;
end;

procedure tWTOF.BSELECTALL_OnClick(Sender: TObject);
var
  {$IFDEF EAGLCLIENT}
	wBookMark: integer;
  i: integer;
  {$ELSE}
	wBookMark: tBookMark;
  {$ENDIF}
begin
  if not (ecran is TFMUL) then exit;
	if wSelectAll then
  begin
    with TFMul(Ecran) do
    begin
      Q.DisableControls;
      try
        {$IFDEF EAGLCLIENT}
        if FListe.RowCount>1 then
        begin
          wBookMark := Fliste.Row;
          try
            for i := 1 to Fliste.RowCount - 1 do
            begin
              FListe.Row := i;
              if FListe.IsSelected( FListe.Row ) = not TToolbarButton97(GetControl('BSELECTALL')).down then
                FListe.FlipSelection( FListe.Row );
            end;
          finally
            Fliste.Row := wBookMark;
          end;
        end;
        {$ELSE}
        wBookMark := GetDataSet.GetBookmark;
        try
          Q.First;
          while not Q.Eof do
          begin
            if FListe.IsCurrentSelected = not TToolbarButton97(GetControl('BSELECTALL')).down then
              FListe.FlipSelection;
            Q.Next;
          end;
        finally
          Q.GotoBookmark(wBookMark);
        end;
        {$ENDIF}
      finally
        Q.EnableControls;
      end;
    end;
  end
  else
    TFMUL(ECRAN).BSelectAllClick(Self);
end;

procedure tWTOF.BTUSER_OnClick(Sender: TObject);
begin
  if Assigned(GetControl(GetPrefixe + '_' + FSuffixeUserForBtUser)) then
  begin
    ThValComboBox(GetControl(GetPrefixe + '_' + FSuffixeUserForBtUser)).Value := V_Pgi.User;
    if Ecran is TFMul then
    begin
      TFMul(Ecran).BCherche.Glyph := TFMul(Ecran).iCritGlyphModified.Picture.BitMap;
      if V_PGI.AutoSearch then
        TFMul(Ecran).ChercheClick;
    end
    else if Ecran is TFStat then
    begin
      TFStat(Ecran).BCherche.Glyph := TFStat(Ecran).iCritGlyphModified.Picture.BitMap;
      if V_PGI.AutoSearch then
        TFStat(Ecran).ChercheClick;
    end;
  end;
end;

procedure tWTOF.MnLpBlocNote_OnClick(Sender: TObject);
  function GetParamEtatRev: string;
  begin
    if Pos(GetPrefixe, 'WNT;WGT') > 0 then
      Result := GetString(GetPrefixe + '_ETATREV')
    else
      Result := '';
  end;
Var
  prefixe: string;

begin
  if Ecran is TFSAISIELIST then
  // Pour récupérer la table ligne et pas en-tête.
  // Pas de modification de la fonction GETPREFIXE car dans certains cas on a besoin du prefixe de la table en-tête.
    Prefixe:= TableToPrefixe(TFSaisieList(Ecran).LeFiltre.TableName) // Pour récupérer la table ligne et pas en-tête.
  else
    Prefixe:= GetPrefixe;

  AGLLanceFiche('W','WBLOCNOTE', '', '', 'ACTION=' + iif(Action = '', 'MODIFICATION', Action) + ';PREFIXE=' + Prefixe + ';ETATREV=' + GetParamEtatRev + ';WHERE=' + wFabriqueWhere(Prefixe, GetDataSet) + ';OBLIGATOIRE=' + BoolToStr(FBlocNoteObligatoire));
  RefreshDb;
  if Ecran is TFSAISIELIST then
    TFSaisieList( ecran ).LeFiltre.Refreshlignes;
  SetMemoFrame;
end;

procedure tWTOF.MnProperties_OnClick(Sender: TObject);
var
	Identifiant, Clef: String;
begin
  Identifiant := wGetValueClef1(GetTableName, Ecran, fLequel);
  Identifiant := StringReplace(Identifiant, ';', '~', [rfIgnoreCase, rfReplaceAll]);
  Clef := wMakeFieldString(GetTableName, '~');

  if not IsEmpty() then
    wCallProperties(GetPrefixe, Identifiant, Clef, IntToStr(LongInt(Ecran)));
end;

procedure tWTOF.PmLoupe_OnPopUp(Sender: TObject);
const
  sCannotDoActionOnTypeOrdre = '[STA];[VTE]';

   { Calucl du contexte }
  function GetCtx: string;
 {$IFDEF GPAO}
    Var
      TypeOrdre: string;
 {$ENDIF GPAO}
    {$IFDEF GCGC}
    {---------------------------------------------------------------------------
      Construit l'argument avec les champs du MUL et la valeur
    ----------------------------------------------------------------------------}
    function GetArgumentYTS : string;
    var
      sArgument, sArgChamp, sChampsActifs, sYTSChamp, sParametreYFOActif, sParametreYFOSelection, sParametreYFOEntete, sParametreYFOClef : string;
    begin
      Result := '';
      GetParametresFromYFO(GetString('YTS_FONCTIONNALITE'), GetString('YTS_CODEPORT'), GetArgumentValue(stArgument, 'ORIENTATION'), sParametreYFOActif, sParametreYFOSelection, sParametreYFOEntete, sParametreYFOCLef);
      sArgument     := GetNomChampsFromPopMenu(False, sParametreYFOActif);
      sChampsActifs := GetNomChampsFromPopMenu(True , sParametreYFOActif);
      while (sArgument<>'') do
      begin
        sYTSChamp  := ReadTokenSt(sArgument);
        sArgChamp := StringReplace(sYTSChamp, 'YTS_', '', [rfReplaceAll]);
        if assigned(GetControl(sYTSChamp)) then
        begin
          if (pos(sArgChamp,sChampsActifs)>0) then
          begin
            if (pos(sYTSChamp,'YTS_DATEDEBUT;YTS_DATEFIN')>0) then
              Result := Result + sArgChamp +'='+ usDateTime(GetDateTime(sYTSChamp))+';'
            else
              Result := Result + sArgChamp +'='+ iif(GetString(sYTSChamp)='-','',GetString(sYTSChamp)) +';';
          end
          else
            Result := Result + sArgChamp +'='+ ';';
        end;
      end;
    end; {GetArgumentYTS}
    {$ENDIF GCGC}

    function GetCtxWithPiece: String;
    begin
      Result := 'WITHPIECE=' + BoolToStr((GetString (GetPrefixe + '_NATUREPIECEG') <> '') and
                                         (GetString (GetPrefixe + '_SOUCHE')       <> '') and
                                         (GetInteger(GetPrefixe + '_NUMERO')       <> 0 ))
    end;

  begin
    Result := '';

    { Laisser en premier pour concordance if...else if... et directives de compil. }
    if GetPrefixe = 'GVP' then
    begin
      Result := 'APPEL=VALORISATION;';
      Result := Result + 'FAMILLEVALO='+GetString('GVP_FAMILLEVALO')+';';
      Result := Result + 'FAMILLENIV1='+GetString('GVP_FAMILLENIV1')+';';
      Result := Result + 'FAMILLENIV2='+GetString('GVP_FAMILLENIV2')+';';
      Result := Result + 'FAMILLENIV3='+GetString('GVP_FAMILLENIV3')+';';
    end
    {$IFDEF GPAO}
      else if GetPrefixe = 'WOP' then
      begin
        TypeOrdre:= wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOP_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOP_LIGNEORDRE'));
        Result:= iif((Pos('[' + TypeOrdre + ']', sCannotDoActionOnTypeOrdre) = 0) and (GetString('WOP_TIERS') <> ''),';TYPEORDRE=STP','')
	             + ';ETATPHASE=' + GetString('WOP_ETATPHASE');
      end
      else if GetPrefixe = 'WOT' then
      begin
        Result := GetCtxWithPiece
      end
      else if GetPrefixe = 'WOL' then
      begin
        Result := 'ORDREPERE=' + GetString('WOL_ORDREPERE')
                + ';' + GetCtxWithPiece
      end
      else if GetPrefixe = 'WOG' then
        Result := 'ETATOPE=' + GetString('WOG_ETATOPE')
      else if GetPrefixe = 'WOR' then
        Result := 'ETATOPE=' + wGetSqlFieldValue('WOG_ETATOPE', 'WORDREGAMME', 'WOG_NATURETRAVAIL="' + GetString('WOR_NATURETRAVAIL') + '" AND WOG_LIGNEORDRE=' + IntToStr(GetInteger('WOR_LIGNEORDRE')) + ' AND WOG_OPECIRC="' + GetString('WOR_OPECIRC') + '" AND WOG_NUMOPERGAMME=' + IntToStr(GetInteger('WOR_NUMOPERGAMME')))
    {$ENDIF GPAO}
    {$IFDEF GCGC}
      else if GetPrefixe = 'YTS' then
      begin
        Result := 'APPEL=TARIF; '+GetArgumentYTS;
      end
    {$ENDIF GCGC}
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
      if Assigned(GetDataSet.FindField('GSM_ARTICLE')) then
        Result := 'ARTICLE=' + GetString('GSM_ARTICLE')
      else
        Result := 'ARTICLE=' + GetString('GA_ARTICLE');
      Result := Result
              + ';LOTINTERNE='   + GetString('GSM_LOTINTERNE')
              + ';SERIEINTERNE=' + GetString('GSM_SERIEINTERNE');
    end
    else if GetPrefixe = 'GQD' then
    begin
      Result := ';ARTICLE='      + GetString('GQD_ARTICLE')
              + ';DEPOT='        + GetString('GQD_DEPOT')
              + ';LOTINTERNE='   + GetString('GQD_LOTINTERNE')
              + ';SERIEINTERNE=' + GetString('GQD_SERIEINTERNE')
    end
    {$ENDIF STK}
  end;

begin
  if IsEmpty then Abort;

  if (Assigned(Ecran)) and ((Ecran is TFSaisieList) or (Ecran is TFVierge) or (Ecran is TFGRS1)) then   // Fiche de saisie en liste
   	wSetMnuLoupe(False, GetCtx, tPopUpMenu(GetControl('PMLOUPE')))
  else
    //Rend poosible le clic droit contextuel pour les lignes sélectionnées
    //si validation par la qualité retirer argument (MultiSelected) de la fonction qui devient alors inutile
   	//wSetMnuLoupe(wMultiSelected(TFMul(Ecran).FListe), GetCtx, tPopUpMenu(GetControl('PMLOUPE')));
   	wSetMnuLoupe(False, GetCtx, tPopUpMenu(GetControl('PMLOUPE')));
end;

{$IFDEF GPAOLIGHT}
procedure tWTOF.PmAction_OnPopUp(Sender: TObject);
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

    function AllowActionByTypeOrdre(const TypeOrdre: String; const MiseEnProd: String = ''): Boolean;
    begin
      Result := Pos('[' + TypeOrdre + ']', sCannotDoActionOnTypeOrdre) = 0;
      if (MiseEnProd <> '') and not Result and (TypeOrdre = 'VTE') then
        Result := MiseEnProd <> 'ALL'
    end;

  begin
    if GetPrefixe = 'WOT' then
    begin
      if (GetString('WOT_TYPEORDRE')<>'STA')  and (GetString('WOT_TYPEORDRE') <> 'VTE') then
        Result := 'ETATTET=' + GetString('WOT_ETATTET')
      else
        Result := '';
    end
	  else if GetPrefixe = 'WOL' then
    begin
      if (GetString('WOL_TYPEORDRE')<>'STA') and (GetString('WOL_MISEENPROD') <> 'ALL') then
      begin
        Result := 'ETATTET=' + wGetSqlFieldValue('WOT_ETATTET', 'WORDRETET', 'WOT_NATURETRAVAIL="' + GetString('WOL_NATURETRAVAIL') + '" AND WOT_CODEORDRE=' + GetString('WOL_CODEORDRE'))
                + ';ETATLIG=' + GetString('WOL_ETATLIG');
        if GetString('WOL_TYPEORDRE') = 'VTE' then
          Result := Result + ';MISEENPROD=' + GetString('WOL_MISEENPROD');
      end
      else
        Result := 'ETATSTA=' + GetString('WOL_ETATLIG');
    end
   	else if GetPrefixe = 'WOP' then
    begin
      TypeOrdre:= wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOP_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOP_LIGNEORDRE'));
      if AllowActionByTypeOrdre(TypeOrdre, wGetSqlFieldValue('WOL_MISEENPROD', 'WORDRELIG', 'WOL_NATURETRAVAIL="' + GetString('WOP_NATURETRAVAIL') + '" AND WOL_LIGNEORDRE=' + GetString('WOP_LIGNEORDRE'))) then
        Result := 'ETATLIG=' + wGetSqlFieldValue('WOL_ETATLIG', 'WORDRELIG', 'WOL_NATURETRAVAIL="' + GetString('WOP_NATURETRAVAIL') + '" AND WOL_LIGNEORDRE=' + GetString('WOP_LIGNEORDRE'))
                + ';ETATPHASE=' + GetString('WOP_ETATPHASE')  + ';QLANSAIS=' + FloatToStr(GetDouble('WOP_QLANSAIS')) + ';QACCSAIS=' + FloatToStr(GetDouble('WOP_QACCSAIS'))
                + iif((Pos('[' + TypeOrdre + ']', sCannotDoActionOnTypeOrdre) = 0) and (GetString('WOP_TIERS') <> ''),';TYPEORDRE=STP','')
                {$IFDEF GPAO}
                  + iif(ExistePiecesCSPEnCours(GetCleWOP),';PIECESCSPENCOURS=X','')
                {$ENDIF GPAO}
      else
        Result:='';
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
        Result:='';
    end
   	else if GetPrefixe = 'WOG' then
    begin
      Result := 'ETATOPE=' + GetString('WOG_ETATOPE');
    end
   	else if GetPrefixe = 'WOR' then
    begin
      Result := 'ETATRES=' + GetString('WOR_ETATRES');
    end
    else if GetPrefixe = 'WPC' then
    begin
      Result := 'ETATPARC=' + GetString('WPC_ETATPARC');
    end
    else if GetPrefixe = 'WVS' then
    begin
      Result := 'ETATVER=' + GetString('WVS_ETATVER')
    end
   	else if GetPrefixe = 'WLS' then
    begin
      if ecran is TFSAISIELIST then
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
    end;
    Result := Result + ';ACTION=' + Action;
   end;
{ Main PmAction_OnPopUp }
begin
  if IsEmpty then Abort;

  if Ecran is TFSAISIELIST then
    wSetMnuAction(False, GetCtx , tPopUpMenu(GetControl('PMACTION')))
  else
		wSetMnuAction(MultiSelected, GetCtx, tPopUpMenu(GetControl('PMACTION')))
end;
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
procedure tWTOF.PmBaseTech_OnPopUp(Sender: TObject);

  { Calcul du contexte }
  function GetCtx: string;
  begin
	  if GetPrefixe = 'WOL' then
    begin
      Result := 'ETATLIG=' + GetString('WOL_ETATLIG');
    end;
    Result := Result + ';ACTION=' + Action;
  end;
begin
  if IsEmpty then Abort;

  wSetMnuBaseTech(MultiSelected, GetCtx, tPopUpMenu(GetControl('PMBASETECH')))
end;
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
procedure tWTOF.MnJournalAction_OnClick(Sender: TObject);
var
  BoutonAnnule: Boolean;
  TabWOL      : MyArrayValue;
  SSTP        : boolean;
const
  MISEENPROD = 0;
  TYPEORDRE  = 1;
begin
  SSTP:= false;
  if (GetPrefixe = 'WOL') and (Assigned(FTreeTobFrame) or
  {$IFDEF EAGLCLIENT}
    (GetDataSet.FieldExists('WOL_MISEENPROD'))) then
  {$ELSE  EAGLCLIENT}
    Assigned(GetDataSet.FindField('WOL_MISEENPROD'))) then
  {$ENDIF EAGLCLIENT}
  begin
    SetLength(TabWOL, 2);
    TabWOL[MISEENPROD] := GetString('WOL_MISEENPROD');
    TabWOL[TYPEORDRE]  := GetString('WOL_TYPEORDRE')
  end
  else if GetPrefixe = 'WOP' then
  begin
    TabWOL := wGetSqlFieldsValues(['WOL_MISEENPROD', 'WOL_TYPEORDRE'], 'WORDRELIG', 'WOL_NATURETRAVAIL="' + GetString('WOP_NATURETRAVAIL') + '" AND WOL_LIGNEORDRE=' + GetString('WOP_LIGNEORDRE'));
    SSTP   := GetString('WOP_TIERS') <> '';  // est une phase de sous-traitance
  end
  else if GetPrefixe = 'WOB' then
    TabWOL := wGetSqlFieldsValues(['WOL_MISEENPROD', 'WOL_TYPEORDRE'], 'WORDRELIG', 'WOL_NATURETRAVAIL="' + GetString('WOB_NATURETRAVAIL') + '" AND WOL_LIGNEORDRE=' + GetString('WOB_LIGNEORDRE'))
  else
  begin
    SetLength(TabWOL, 2);
    TabWOL[MISEENPROD] := '';
    TabWOL[TYPEORDRE]  := ''
  end;

  BoutonAnnule := (TabWOL[MISEENPROD] <> 'ALL') and (TabWOL[TYPEORDRE] <> 'STA') and ( not SSTP);

  if Ecran is TFSAISIELIST then
    wCallMulWJA(GetPrefixe, TFSaisieList(Ecran).LeFiltre.TOBFiltre.GetInteger(GetPrefixe + '_IDENTIFIANT'), Action, BoutonAnnule)
  else
    wCallMulWJA(GetPrefixe, GetInteger(GetPrefixe + '_IDENTIFIANT'), Action, BoutonAnnule);
  RefreshDB;
end;
{$ENDIF GPAOLIGHT}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 26/09/2002
Description .. : Déroute le onElipsisclick sur les contrôles THedit
Suite ........ : ou THdbEdit pour les contrôles 'articles'
NDP : Cette procédure est répliquée dans la wtom, merci de la mettre
à jour si vous modifiez cette procédure la.
*****************************************************************}
procedure TWTOF.ScanControls;
var
	iComp: Integer;
  Ok: Boolean;
  C: TComponent;
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
        end;
      end;
      {$ENDIF}

      if (not Ok) and (C is ThEdit) and (ThEdit(C).ElipsisButton) then
      begin
        { Essai avec le suffixe du nom du champ, Si le name est pas ok, ré-essaye avec le DataType }
        if wIsDataTypeArticle(ThEdit(C).DataType) or wIsDataTypeCodeArticle(ThEdit(C).DataType) then
          ThEdit(C).OnElipsisClick := ArticleElipsisClick;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 26/09/2002
Description .. : Appel le mul ou le lookup de recherche des articles
NDP : Cette procédure est sauvagement répliquée dans la wTOM, merci
de la mettre à jour si vous modifiez cette procédure si...
*****************************************************************}
procedure TWTOF.ArticleElipsisClick(Sender: TObject);
var
	Text, Params, Plus, Name, DataType: String;
begin
  Plus := '';
  Name := '';
  Text := '';
  DataType := '';

	if Sender is ThEdit then
  begin
    Plus := ThEdit(Sender).Plus;
    Name := ThEdit(Sender).Name;
    Text := ThEdit(Sender).Text;
    DataType := ThEdit(Sender).DataType;
  end;

  {$IFNDEF EAGLCLIENT}
  if Sender is ThdbEdit then
  begin
    Plus := ThdbEdit(Sender).Plus;
    Name := ThdBEdit(Sender).DataField;
    Text := ThdbEdit(Sender).Text;
    DataType := ThdbEdit(Sender).DataType;
  end;
  {$ENDIF}

  { Retourne le code article ou l'article }
  if wIsSuffixeCodeArticle(Name) or wIsDataTypeCodeArticle(DataType) then
    Params := 'RETOUR_CODEARTICLE=' + wTrue
  else
    Params := '';

  { Ajoute le plus dans le where du mul }
  if Plus <> '' then
  begin
    if Params <> '' then Params := Params + ';';
    Params := Params + 'XX_WHERE=' + Plus;
  end;

  { Initialise le code article dans le mul }
  if Text <> '' then
  begin
    if Params <> '' then Params := Params + ';';
    Params := Params + 'GA_CODEARTICLE=' + Text;
  end;

  { Accès à tous les articles ou seulement aux article profils }
  if Params <> '' then Params := Params + ';';

 	if Params <> '' then Params := Params + ';';
    Params := Params + 'RECHERCHEARTICLE';

  { Appel du mul + Mise à jour du contrôle }
	DispatchRecherche(ThCritMaskEdit(Sender), 1, Plus, Params, NomMulRechArticle);
end;

function tWTOF.IsEmpty: Boolean;
var
  Q : tQuery;
begin
  Q := GetQuery;
  if Assigned(Q) then
  begin
    {$IFNDEF EAGLCLIENT}
      Result := (Q.Eof) and (Q.Bof);
    {$ELSE}
      Result := Q.RecordCount = 0;
    {$ENDIF}
  end
  else
    Result := false;
end;

procedure tWTOF.PmHistorique_OnPopUp(Sender: TObject);
  { Calcul du contexte }
  function GetArgument: string;
  begin
    if GetPrefixe = 'WOB' then
    begin
      Result := ';TYPELIEN='  + GetString('WOB_TYPELIEN')
              + ';MODECONSO=' + GetString('WOB_MODECONSO')
    end;

    Result := Result + ';ACTION=' + Action;
  end;
begin
  if IsEmpty then Abort;
 	wSetMnuHisto(MultiSelected, GetArgument, tPopUpMenu(GetControl('PMHISTORIQUE')));
end;

procedure tWTOF.BSUPPRIME_OnClick(Sender: TObject);
begin
  if isEmpty then Abort;
end;

function TWTOF.GetTableName: string;
begin
//  if Ecran is TFSAISIELIST then Result := TFSaisieList(Ecran).LeFiltre.TableName
  if Ecran is TFSAISIELIST then Result := TFSaisieList(Ecran).LeFiltre.TableEntete
  else                          Result := FTableName;
end;

procedure TWTOF.BCherche_OnClick(Sender: TOBject);
begin
  if Ecran is TFMul then
  begin
    try
      SetArticle;
      SetControlProperty('BSELECTALL', 'DOWN', false);
    finally
      TFMUL(ECRAN).BChercheClick(Nil);
    end;
  end
  else if ecran is TFStat then
    TFStat(Ecran).BChercheClick(Nil)
  else if Ecran is TfCube then
    TFCube(Ecran).BChercheClick(Nil)
end;

function tWTOF.GetQuery: tQuery;
begin
  Result := nil;
  if Ecran is TfMul then
  begin
    {$IFNDEF EAGLCLIENT}
    Result := TFMul(Ecran).Q
    {$ELSE}
    if TFMul(Ecran).Q.TQ<>nil then
      begin
      TFMul(Ecran).Q.TQ.seek(TFMul(Ecran).FListe.Row-1);
      Result := TFMul(Ecran).Q.TQ;
      end
    {$ENDIF}
  end;
end;

function tWTOF.GetPrefixe: string;
begin
  Result := TableToPrefixe(GetTableName);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. : affiche le mémo dans le cadre
Mots clefs ... :
*****************************************************************}
procedure tWTOF.SetMemoFrame;
var
  M: TMemo;
begin
  if Assigned(GetControl('BTSHOWBLOCNOTE')) and Assigned(GetControl(GetPrefixe + '_' + fBlobFieldSuffixe)) then
  begin
    if TToolBarButton97(GetControl('BTSHOWBLOCNOTE')).Down then
    begin
      M := TMemo(GetControl(GetPrefixe + '_' + fBlobFieldSuffixe));
      M.Clear;
      if not isEmpty then
      begin
        {$IFNDEF EAGLCLIENT}
        if Assigned(GetDataSet.FindField(GetPrefixe + '_' + 'WBMEMO')) then
        {$ELSE}
        if GetDataSet.FieldExists(GetPrefixe + '_' + 'WBMEMO') then
        {$ENDIF}
        begin
          if GetBoolean(GetPrefixe + '_' + 'WBMEMO') then
            M.Text := wGetSQLFieldValue(GetPrefixe + '_' + fBlobFieldSuffixe, GetTableName, wMakeWhereSQL(wMakeFieldString(GetTableName, ';'), wGetValueClef1(GetTableName, Ecran, fLequel)));
        end
        else
          M.Text := wGetSQLFieldValue(GetPrefixe + '_' + fBlobFieldSuffixe, GetTableName, wMakeWhereSQL(wMakeFieldString(GetTableName, ';'), wGetValueClef1(GetTableName, Ecran, fLequel)));
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. : Cadre permettant de visualiser les blobs dans les muls
Suite ........ : Sert surtout aux journaux
Mots clefs ... :
*****************************************************************}
procedure tWTOF.BTSHOWBLOCNOTE_OnClick(Sender: tObject);
begin
  if Assigned(GetControl('BTSHOWBLOCNOTE')) then
  begin
    SetControlVisible(GetPrefixe + '_' + fBlobFieldSuffixe, TToolBarButton97(GetControl('BTSHOWBLOCNOTE')).Down);
    THSplitter(GetControl('SPLITPNMEMO')).Visible := TToolBarButton97(GetControl('BTSHOWBLOCNOTE')).Down;
  end;
  SetMemoFrame;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. : permet de grouper les SetControlEnabled
Mots clefs ... :
*****************************************************************}
procedure TWTof.SetControlsEnabled(Const FieldsName: array of String; Const Enability: Boolean);
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
procedure TWTof.SetControlsVisible(Const FieldsName: array of String; Const Visibility: Boolean);
var
  i: integer;
begin
  for i := 0 to Length(FieldsName) - 1 do
    SetControlVisible(FieldsName[i], Visibility);
end;


function TWTof.GetChampsMUL: String;
var
  l, GridCount, i: Integer;
  sToReplace: String;
  Grid: THGrid;
begin
  Result := '';

  if Ecran is TFVierge then
  begin
    Grid := THGrid(GetControl(FViergeGridName));
    GridCount := Grid.ColCount;

    for i := 0 to Pred(GridCount) do
      Result := Result + iif(i >= Grid.FixedCols, Grid.ColNames[i], '') + iif(i < Pred(GridCount), ';', '')
  end
  else if Ecran is TFMul then
  begin
      Result := Trim(UpperCase(TFMul(Ecran).Q.Champs));
      Result := StringReplace(Result, ',', ';', [rfIgnoreCase, rfReplaceAll]);
      while Pos(' AS ', Result) <> 0 do
      begin
        l := Pos(' AS ', Result) - 1;
        sToReplace := '';
        while (Result[l] <> ';') and (l > 0) do
        begin
          sToReplace := Result[l] + sToReplace;
          Dec(l)
        end;
        Result := StringReplace(Result, sToReplace + ' AS ', '', [rfIgnoreCase, rfReplaceAll])
      end;
      Result := StringReplace(Result, ' ', '', [rfIgnoreCase, rfReplaceAll])
  end;
end;


function TWTof.GetColWidths: MyArrayInt;
var
  i, GridCount: Integer;
  Grid: THGrid;
begin
  {$IFNDEF EAGLCLIENT}
    if (Ecran is TFMul) and (FViergeGridName = 'FListe') then
    begin
      if V_PGI.Sav then
        PGIError('Utilisation 2/3 non conforme.');
      Exit;
    end;
  {$ENDIF EAGLCLIENT}
  SetLength(Result, 0);

  Grid := THGrid(GetControl(FViergeGridName));

  if Assigned(Grid) then
  begin
    GridCount := Grid.ColCount;

    for i := (0 + Grid.FixedCols) to Pred(GridCount) do
    begin
      SetLength(Result, Length(Result) + 1);
      Result[Length(Result) - 1] := Grid.ColWidths[i];
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 28/07/2003
Modifié le ... :   /  /
Description .. : Change le libelle de la colonne portant sur le ColName
Mots clefs ... :
*****************************************************************}
procedure tWTOF.SetColCaption(const ColName, Caption: String);
var
  iCol  : Integer;
  wField: TWFieldListe;
begin
  if FEditColsProperties and Assigned(FWListe) then
  begin
    wField := FWListe.ChampsByName[iif(Pos('@', ColName) > 0, '(' + ColName + '(', ColName)];
    if Assigned(wField) and (wField.Index > -1) then
    begin
      iCol := wField.Index;

      if ((Ecran is TFVierge) or (Ecran is TFSaisieList)) and (FViergeGridName <> '') and Assigned(GetControl(FViergeGridName)) then
      begin
        with THGrid(GetControl(FViergeGridName)) do
          Cells[iCol + FixedCols, 0] := Caption
      end
      else if Ecran is TFMul then
      begin
        {$IFDEF EAGLCLIENT}
          if FViergeGridName <> 'FListe' then
            THGrid(GetControl(FViergeGridName)).Cells[iCol + THGrid(GetControl(FViergeGridName)).FixedCols, 0] := Caption
          else
            TFMul(Ecran).FListe.Cells[iCol + TFMul(Ecran).FListe.FixedCols, 0] := Caption;
        {$ELSE EAGLCLIENT}
          TFMul(Ecran).SetDisplayLabel(ColName, Caption)
        {$ENDIF EAGLCLIENT}
      end
    end;

    if Copy(Caption, 1, 2) = '.-' then
      SetColVisible(ColName, False);
  end
  {$IFDEF GPAODEV}
    else if V_Pgi.Sav then
    begin
      if not FEditColsProperties then
        PGIInfo('Renseigner la propriété "EditColsProperties" à "True" (après inherited du OnArgument)', 'Débug mode SAV')
      else if FViergeGridName = '' then
        PGIInfo('Renseigner la propriété "ViergeGridName" de la wTof du nom de la THGrid (après inherited du OnArgument)', 'Débug mode SAV')
      else if not (Ecran is TFMul) and (FViergeGridName <> '') and (GetControl(FViergeGridName) is THGrid) and (THGrid(GetControl(FViergeGridName)).ListeParam = '') then
        PGIInfo('Renseigner la propriété "ListeParam" de la THGrid', 'Débug mode SAV')
    end
  {$ENDIF GPAODEV}
end;

procedure tWTOF.SetColsCaption(const FieldsName: array of String; const Caption: String);
var
  i: Integer;
begin
  for i := Low(FieldsName) to High(FieldsName) do
    SetColCaption(FieldsName[i], Caption)
end;

procedure tWTOF.PmUtilitaire_OnPopUp(Sender: TObject);
{$IFDEF GPAO}
  Var
    sPeriodeSauv : string;
{$ENDIF GPAO}
    
  { Calcul du contexte }
  function GetArgument: string;
  begin
    if GetPrefixe = 'GA' then
    begin
      Result := 'TENUESTOCK='  + GetString('GA_TENUESTOCK')
    end
{$IFDEF GPAOLIGHT}
    else if GetPrefixe = 'WOL' then
    begin
      Result := 'TYPEORDRE='  + GetString('WOL_TYPEORDRE')
              + ';ETATLIG='   + GetString('WOL_ETATLIG')
              + ';ORDREPERE=' + GetString('WOL_ORDREPERE')
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
{$ENDIF GPAOLIGHT}
{$IFDEF STK}
    else if GetPrefixe = 'GQ' then
    begin
      Result := 'DEPOT='         + GetString('GQ_DEPOT')
              + ';ARTICLE='      + GetString('GQ_ARTICLE')
              + ';DATECLOTURE='  + GetString('GQ_DATECLOTURE')
              + ';CONTREMARQUE=' + wGetFieldFromGA('GA_CONTREMARQUE', GetString('GQ_ARTICLE'))
    end
    else if GetPrefixe = 'GQD' then
    begin
      Result := 'DEPOT='        + GetString('GQD_DEPOT')
              + ';ARTICLE='     + GetString('GQD_ARTICLE')
              + ';STATUTDISPO=' + GetString('GQD_STATUTDISPO')
              + ';STATUTFLUX='  + GetString('GQD_STATUTFLUX')
              + ';CONTREMARQUE='+ GetString('GQD_CONTREMARQUE')
    end
    else if GetPrefixe = 'GSM' then
    begin
      Result := 'ETATMVT=' + GetString('GSM_ETATMVT')
    end
{$ENDIF STK}
{$IFDEF GPAO}
    else if GetPrefixe = 'WPE' then
    begin
		  if (Getstring('WPE_NATUREPDR')='BTH') then
  		begin
    		sPeriodeSauv := wPadRight(wDateTimeToStr(V_PGI.DateEntree+time,wCalculPeriodeSauvegarde(GetString('WPE_UNPDRPAR'))), iLongueurPeriodeSauvegarde, '_');
    		Result:= 'PERIODESAUVOK='+iif(GetString('WPE_PERIODESAUV')=sPeriodeSauv, wTrue, wFalse);
  		end
		  else
    		Result:= 'PERIODESAUVOK=X';
    end;
{$ENDIF GPAO}
    ;

    Result := Result + ';ACTION=' + Action;
  end;
{ Main PmUtilitaire_OnPopUp }
begin
  if IsEmpty then Abort;

  if Ecran is TFSAISIELIST then
    wSetMnuUtilitaire(False, GetArgument , tPopUpMenu(GetControl('PMUTILITAIRE')))
  else
    wSetMnuUtilitaire(MultiSelected, GetArgument, tPopUpMenu(GetControl('PMUTILITAIRE')));
end;

procedure tWTOF.OnDisplay;
begin
  inherited;

  SetMemoFrame;
end;

{$IFNDEF EAGLCLIENT}
{ empêche de modifier la position d'une colonne en 2/3 non cwas si l'on active
la possibilité de jouer sur la visibilité des colonnes et de leur libellé }
procedure tWTof.DBGrid_OnColMoved(Sender: TObject; FromIndex, ToIndex: Longint);
begin
  TDBGrid(Sender).Columns[ToIndex].Index := FromIndex
end;
{$ENDIF EAGLCLIENT}

procedure tWTof.SetEditColsProperties(const Value: Boolean);
begin
  FEditColsProperties := Value;
  if Value then
  begin
    if Ecran is TFMul then
    {$IFDEF EAGLCLIENT}
      TFMul(Ecran).FListe.Options := TFMul(Ecran).FListe.Options - [goColMoving]
    {$ELSE}
      TFMul(Ecran).FListe.OnColumnMoved := DBGrid_OnColMoved
    {$ENDIF EAGLCLIENT}
    else
    begin
      if Ecran is TFSaisieList then
      begin
        FViergeGridName := 'FListe';
        LoadColsList
      end;
      THGrid(GetControl(FViergeGridName)).Options := THGrid(GetControl(FViergeGridName)).Options - [goColMoving]
    end;
  end
end;

procedure tWTOF.SetDimensions;
var
  Affiche : boolean;
  Dim     : char;
  Libelle : string;
begin
  if Assigned(GetControl('TSDIMENSIONS')) then
  begin
    Affiche := (Pos('DIM', ThMultiValComboBox(GetControl('GA_STATUTART')).Text) > 0) or (ThMultiValComboBox(GetControl('GA_STATUTART')).Text = traduireMemoire('<<Tous>>'));

    SetControlVisible('TSDIMENSIONS', Affiche);

    if Affiche then
    begin
      for Dim := '1' to '5' do
      begin
        Libelle := RechDom('GCCATEGORIEDIM', 'DI' + Dim, False);

        { Gestion des libellés }
        if Assigned(GetControl('DIMENSION' + Dim)) then
        with THLabel(GetControl('DIMENSION' + Dim)) do
        begin
          Caption := Libelle;
          Visible := Affiche and (Copy(Libelle, 1, 2) <> '.-');
        end;

        { Gestion des grilles }
        if Assigned(GetControl('TGA_GRILLEDIM')) then THLabel(GetControl('TGA_GRILLEDIM')).Visible := Affiche;
        if Assigned(GetControl('GA_GRILLEDIM' + Dim)) then
        with ThValComboBox(GetControl('GA_GRILLEDIM' + Dim)) do
        begin
          Visible := THLabel(GetControl('DIMENSION' + Dim)).Visible;
          Value := iif(Visible, Value, '');
        end;

        { Gestion des valeurs }
        if Assigned(GetControl('TGA_LIBELLEDIM')) then THLabel(GetControl('TGA_LIBELLEDIM')).Visible := Affiche;
        if Assigned(GetControl('GA_CODEDIM' + Dim)) then
        with ThValComboBox(GetControl('GA_CODEDIM' + Dim)) do
        begin
          Visible := THLabel(GetControl('DIMENSION' + Dim)).Visible;
          Value := iif(Visible, Value, '');
        end;

        { Gestion des libellés des colonnes }
        if FEditColsProperties then
        begin
          SetColCaption('GA_CODEDIM' + Dim, Libelle);
          SetColVisible('GA_CODEDIM' + Dim, Affiche and (Copy(Libelle, 1, 2) <> '.-'));
        end;
      end;
    end
    else
    begin
      for Dim := '1' to '5' do
      begin
        ThValComboBox(GetControl('GA_GRILLEDIM' + Dim)).Value := '';
        ThValComboBox(GetControl('GA_CODEDIM' + Dim)).Value := '';
        SetColVisible('GA_CODEDIM' + Dim, False);
      end;

      { HmTrad }
      if (ecran is TFMUL) and (not TfMul(Ecran).NotResizeDbGrid) then
      begin
        {$IFDEF EAGLCLIENT}
          TFMUL(Ecran).HMTrad.ResizeGridColumns(TFMUL(Ecran).FListe)
        {$ELSE EAGLCLIENT}
          TFMUL(Ecran).HMTrad.ResizeDBGridColumns(TFMUL(Ecran).FListe)
        {$ENDIF EAGLCLIENT}
      end;
    end;
  end;
end;

procedure tWTOF.GA_STATUTART_OnChange(Sender: TObject);
begin
  SetDimensions;
end;

procedure tWTOF.GA_GRILLEDIM_OnChange(Sender: TObject);
  function GetDim: string;
  begin
    Result := wRight(ThValComboBox(sender).Name, 1)
  end;
begin
  if GetControlText('GA_GRILLEDIM' + GetDim) <> '' then
    SetControlProperty('GA_CODEDIM' + GetDim, 'PLUS', ' AND GDI_GRILLEDIM="' + GetControlText('GA_GRILLEDIM' + GetDim)+ '"');
end;

procedure tWTOF.LoadColsList;
var
  DBLst: String;
begin
  if Ecran is TFMul then
  begin
    if FViergeGridName = 'FListe' then
      DBLst := TFMul(Ecran).DBListe
    else
      DBLst := THGrid(GetControl(FViergeGridName)).ListeParam
  end
  else if (Ecran is TFVierge) and (FViergeGridName <> '') then
    DBLst := THGrid(GetControl(FViergeGridName)).ListeParam
  else if (Ecran is TFSaisieList) and (FViergeGridName <> '') then
    DBLst := TFSaisieList(Ecran).DBListe;

  if DBLst <> '' then
  begin
    if Assigned(FWListe) then
      FWListe.Free;
    FWListe := TWListe.Create(DBLst, Ecran)
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 24/02/2004
Modifié le ... :   /  /    
Description .. : Gère la visibilité des colonnes d'un mul (couche supérieure 
Suite ........ : à l'AGL : Besoin dans les stocks, tarifs, cbn, etc)
Suite ........ : ColName est :
Suite ........ : - soit un nom de champ
Suite ........ : - soit un nom de formule sans les paramètres (ATTENTION
Suite ........ : les parenthèses ouvrante et fermantes sont obligatoire dans
Suite ........ : le déclaration de la liste)
Mots clefs ... :
*****************************************************************}
procedure tWTOF.SetColVisible(const ColName: String; const Visible: Boolean);
var
  iCol  : Integer;
  wField: TWFieldListe;
begin
  if FEditColsProperties and Assigned(FWListe) then
  begin
    wField := FWListe.ChampsByName[StringReplace(ColName, ';', '~', [rfIgnoreCase, rfReplaceAll])];
    if Assigned(wField) and (wField.Index > -1) then
    begin
      iCol := wField.Index;

      if ((Ecran is TFVierge) or (Ecran is TFSaisieList)) and (FViergeGridName <> '') and Assigned(GetControl(FViergeGridName)) then
      begin
        with THGrid(GetControl(FViergeGridName)) do
        begin
          if (iCol + FixedCols) < ColCount then
          begin
            if Visible then
              ColWidths[iCol + FixedCols] := FTabWidths[iCol]
            else
              ColWidths[iCol + FixedCols] := -1
          end;
        end
      end
      else if Ecran is TFMul then
      begin
        {$IFDEF EAGLCLIENT}
          if FViergeGridName <> 'FListe' then
          begin
            if Visible then
              THGrid(GetControl(FViergeGridName)).ColWidths[iCol + THGrid(GetControl(FViergeGridName)).FixedCols] := FTabWidths[iCol]
            else
              THGrid(GetControl(FViergeGridName)).ColWidths[iCol + THGrid(GetControl(FViergeGridName)).FixedCols] := -1
          end
          else
          begin
            if Visible then
              TFMul(Ecran).FListe.ColWidths[iCol + TFMul(Ecran).FListe.FixedCols] := FTabWidths[iCol]
            else
              TFMul(Ecran).FListe.ColWidths[iCol + TFMul(Ecran).FListe.FixedCols] := -1
          end
        {$ELSE  EAGLCLIENT}
          TFMul(Ecran).SetVisibleColumn(ColName, Visible)
        {$ENDIF EAGLCLIENT}
      end
    end
  end
  {$IFDEF GPAODEV}
    else if V_Pgi.Sav then
    begin
      if not FEditColsProperties then
        PGIInfo('Renseigner la propriété "EditColsProperties" à "True" (après inherited du OnArgument)', 'Débug mode SAV')
      else if FViergeGridName = '' then
        PGIInfo('Renseigner la propriété "ViergeGridName" de la wTof du nom de la THGrid (après inherited du OnArgument)', 'Débug mode SAV')
      else if not (Ecran is TFMul) and (FViergeGridName <> '') and (GetControl(FViergeGridName) is THGrid) and (THGrid(GetControl(FViergeGridName)).ListeParam = '') then
        PGIInfo('Renseigner la propriété "ListeParam" de la THGrid', 'Débug mode SAV')
    end
  {$ENDIF GPAODEV}
end;

procedure twTof.SetColsVisible(const FieldsName: Array of String; const Visible: Boolean);
var
  i: Integer;
begin
  for i := Low(FieldsName) to High(FieldsName) do
    SetColVisible(FieldsName[i], Visible)
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/03/2004
Modifié le ... : 02/03/2004
Description .. : Recharge la liste associée à un THGrid.
Suite ........ : N'a d'utilité que dans une TFVierge
Mots clefs ... : 
*****************************************************************}
procedure tWTOF.ReloadColsList;
begin
  if FEditColsProperties then
  begin
    LoadColsList;
    if (Ecran is TFVierge) or ((Ecran is TFMul) and  (FViergeGridName <> 'FListe')) then
      FTabWidths := GetColWidths
  end
end;

procedure tWTOF.SetArticle;
Var
  FieldName: string;

  function GetCodeDim(Dim:integer): string;
  begin
    Result := tHValComboBox(GetControl('GA_CODEDIM' + IntToStr(Dim))).Value
  end;

  function GetArticle: string;
  var
    Article: String;
  begin
    Article := wGetArticleFromCodeArticle(GetControlText('GA_CODEARTICLE'), GetCodeDim(1), GetCodeDim(2), GetCodeDim(3), GetCodeDim(4), GetCodeDim(5));
    if wExistGA(Article) and (GetStatutArt(Article) = 'DIM') then
      Result := Article
    else
      Result := ''
  end;

begin
  if   Assigned(GetControl('GA_ARTICLE')) then FieldName := 'GA_ARTICLE'
  else FieldName := GetPrefixe + '_ARTICLE'; 

  if Assigned(GetControl('TSDIMENSIONS')) and Assigned(GetControl(FieldName)) and Assigned(GetControl('GA_CODEARTICLE')) and
     Assigned(GetControl('GA_CODEDIM1')) and Assigned(GetControl('GA_CODEDIM2')) and Assigned(GetControl('GA_CODEDIM3')) and
     Assigned(GetControl('GA_CODEDIM4')) and Assigned(GetControl('GA_CODEDIM5')) then
  begin
    SetControlText(FieldName, GetArticle);
  end;
end;

procedure tWTOF.AffectGroup(Const Indice: integer);
var
  St_Plus, St_Value, St : string;
  i_ind : integer;
begin
  if not (Ecran is TFQRS1) then exit ;
  St_Plus := string (THValComboBox (GetControl('RUPT' + intToStr(indice))).Plus);
  St_Value := GetControlText('RUPT' + InttoStr(Indice));
  for i_ind := 1 to 6 do
  begin
    if i_ind = Indice then continue;
    St := string (THValComboBox (GetControl('RUPT' + InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"';
  end;
  SetControlProperty('RUPT' + InttoStr(Indice), 'PLUS', St_Plus);
  if St_Value<>'' then SetControlProperty('RUPT' + InttoStr(Indice), 'VALUE', St_Value)
                  else SetControlProperty('RUPT' + InttoStr(Indice), 'ITEMINDEX', 0);
end;

procedure tWTOF.ChangeGroup(Const Indice: integer);
var
  St_Plus, St_Value, Tablette: string;
  i : integer;
begin
  if not (Ecran is TFQRS1) then exit ;
  St_Plus := string (THValComboBox (GetControl('RUPT' + intToStr(indice))).Plus);
  St_Value := GetControlText('RUPT' + InttoStr(Indice));
  if St_Value = '' then
  begin
    SetControlProperty('XX_RUPTURE' + InttoStr(Indice), 'TEXT'    , ''    );
    SetControlProperty('XX_VARIABLE'+ InttoStr(Indice), 'TEXT'    , ''    );
    SetControlProperty('XX_VARIABLE'+ InttoStr(Indice+10), 'TEXT'    , ''    );
    SetControlProperty('SAUTRUPT'   + InttoStr(Indice), 'CHECKED' , False );
    SetControlProperty('SAUTRUPT'   + InttoStr(Indice), 'ENABLED' , False );
    for i := Indice + 1 to 6 do
    begin
      SetControlProperty('RUPT'       + InttoStr(i), 'ENABLED'  , True     );
      SetControlProperty('RUPT'       + InttoStr(i), 'VALUE'    , ''        );
      SetControlProperty('RUPT'       + InttoStr(i), 'ENABLED'  , False     );
      SetControlProperty('RUPT'       + InttoStr(i), 'COLOR'    , clBtnFace );
      SetControlProperty('XX_RUPTURE' + InttoStr(i), 'TEXT'     , ''        );
      SetControlProperty('XX_VARIABLE'+ InttoStr(i+10), 'TEXT'     , ''        );
      SetControlProperty('XX_VARIABLE'+ InttoStr(i), 'TEXT'     , ''    );
      SetControlProperty('SAUTRUPT'   + InttoStr(i), 'CHECKED'  , False     );
      SetControlProperty('SAUTRUPT'   + InttoStr(i), 'ENABLED'  , False     );
    end;
  end
  else
  begin
    if Indice < 6 then
    begin
      SetControlProperty('RUPT' + InttoStr(Indice + 1), 'ENABLED'   , True    );
      SetControlProperty('RUPT' + InttoStr(Indice + 1), 'COLOR'     , clWindow);
      SetControlProperty('RUPT' + InttoStr(Indice + 1), 'ITEMINDEX' , 0       );
    end;
    SetControlProperty('SAUTRUPT' + InttoStr(Indice), 'ENABLED', True);
    Tablette:= THValComboBox(GetControl('RUPT' + IntToStr(Indice))).DataType;
    SetControlProperty('XX_RUPTURE' + InttoStr(Indice), 'TEXT', RechDom (Tablette, St_Value, True));
    // Attention ; GetControlText ne ramène que le code de la variable
    // XX_VARIABLE1 contient le libellé du combo / WW_VARIABLE11 contient le code du combo
    SetControlProperty('XX_VARIABLE'+ InttoStr(Indice), 'TEXT', string (THValComboBox (GetControl('RUPT'+InttoStr(Indice))).Text));
    SetControlProperty('XX_VARIABLE'+ InttoStr(Indice + 10), 'TEXT', GetControlText('RUPT' + InttoStr(Indice)));
  end;
end;

procedure tWTOF.ChangeSautPage(Const Indice: integer);
var
  i_ind : integer;
begin
  if not (Ecran is TFQRS1) then exit ;
  if TCheckBox (GetControl('SAUTRUPT' + InttoStr(Indice))).Checked then
  begin
    for i_ind := 1 to Indice - 1 do
      TCheckBox (GetControl('SAUTRUPT' + InttoStr(i_ind))).Checked := True;
  end
  else
  begin
    for i_ind := Indice + 1 to 6 do
      TCheckBox (GetControl('SAUTRUPT' + InttoStr(i_ind))).Checked := False;
  end;
end;

procedure tWTOF.RUPT_OnClick(Sender: TObject);
Var
  Indice : integer;
  NomControl: string;
begin
  if not (Ecran is TFQRS1) then exit ;
  NomControl:= Tcontrol(Sender).Name;
  Indice:= StrtoInt(Copy(NomControl,5,1));
  ChangeGroup(Indice);
end;

procedure tWTOF.RUPT_OnEnter(Sender: TObject);
Var
  Indice : integer;
  NomControl: string;
begin
  if not (Ecran is TFQRS1) then exit ;
  NomControl:= Tcontrol(Sender).Name;
  Indice:= StrToInt(Copy(NomControl,5,1));
  AffectGroup(Indice);
end;

procedure tWTOF.SAUTRUPT_OnClick(Sender: TObject);
Var
  Indice : integer;
  NomControl: string;
begin
  if not (Ecran is TFQRS1) then exit ;
  NomControl:= Tcontrol(Sender).Name;
  Indice:= StrtoInt(Copy(NomControl,9,1));
  ChangeSautPage(Indice);
end;

function tWTOF.RechLib(sdataType, sFieldname : string; lAbrege : Boolean; lNotCheckBox : Boolean; Plus : string = '') : string;
var s : string;
begin
  s := RechDom(sdataType, 'T' + sFieldName, lAbrege, Plus);
  if lNotCheckBox then
    sFieldName := 'T' + sFieldName;
  if s = 'Error' then
  begin
    if lNotCheckBox then
      s := GetControlText(sFieldName)
    else
      s := tCheckBox(GetControl(sFieldName)).Caption;
  end;
  result := s;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 03/08/2004
Modifié le ... :   /  /
Description .. : Permet la gestion des libellés spéciaux...
Mots clefs ... :
*****************************************************************}
procedure tWTOF.SetLibelle;
var
  Affiche                                           : Boolean;
  i                                                 : integer;
  MasterPrefixe, FieldName, MasterFieldName, Libelle, NaturePrefixe : string;

  function GetNaturePrefixe: string;
  begin
    if Pos(UpperCase(Copy(MasterPrefixe, 1, 2)), 'QW;QO;QU;QJQK') > 0 then
      Result := 'QMES'
    else if Copy(MasterPrefixe, 1, 2) = 'GC' then
      Result := 'GC'
    else
      Result := 'W'
  end;
begin
  { Familles article }
  for i := 1 to 3 do
  begin
    { GA }
    FieldName := 'GA_FAMILLENIV' + IntToStr(i);
    Libelle   := RechDom('GCLIBFAMILLE', 'LF' + IntToStr(i), False);
    Affiche   := Copy (Libelle, 1, 2) <> '.-'; 
    SetControlProperty('T' + FieldName, 'Caption', Libelle);
    SetControlVisible(FieldName, Affiche);
    SetControlVisible('T' + FieldName, Affiche);

    if FEditColsProperties then
    begin
      SetColCaption(FieldName, Libelle);
      SetColVisible(FieldName, Affiche);
    end;

    { Autre }
    FieldName := GetPrefixe + '_FAMILLENIV' + IntToStr(i);
    SetControlProperty('T' + FieldName, 'Caption', Libelle);
    SetControlVisible(FieldName, Affiche);
    SetControlVisible('T' + FieldName, Affiche);

    if FEditColsProperties then
    begin
      SetColCaption(FieldName, Libelle);
      SetColVisible(FieldName, Affiche);
    end;
  end;

  { Pour les articles et les ressources cela est géré par GCMAJChampLibre() }
  if (GetPrefixe <> 'GA') and (GetPrefixe <> 'ARS') and (GetPrefixe <> 'WAP') then
  begin
    { Table libres }
    {$IFDEF GPAOLIGHT}
      MasterPrefixe := GetMasterPrefixe(GetPrefixe);
    {$ELSE GPAOLIGHT}
      MasterPrefixe := GetPrefixe;
    {$ENDIF GPAOLIGHT}
    for i := 1 to 10 do
    begin
      FieldName       := GetPrefixe    + '_LIBRE' + GetPrefixe    + intToHex(i, 1);
      MasterFieldName := MasterPrefixe + '_LIBRE' + MasterPrefixe + intToHex(i, 1);
      NaturePrefixe := GetNaturePrefixe;

      if Assigned(GetControl(FieldName)) then
      begin
        SetControlProperty('T' + FieldName, 'CAPTION', RechLib(NaturePrefixe + 'LIBELLELIBRE' + MasterPrefixe, MasterFieldName, false, true));
        SetControlVisible('T' + FieldName, Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible(FieldName      , Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
      end;
    end;

    { Valeurs libres }
    for i := 1 to 3 do
    begin
      FieldName       := GetPrefixe    + '_VALLIBRE' + intToStr(i);
      MasterFieldName := MasterPrefixe + '_VALLIBRE' + intToStr(i);
      if Assigned(GetControl(FieldName)) then
      begin
        SetControlProperty('T' + FieldName, 'CAPTION', RechLib(NaturePrefixe + 'VALLIBRE' + MasterPrefixe, MasterFieldName, false, true));
        SetControlVisible('T' + FieldName      , Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible(FieldName            , Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible('T' + FieldName + '_', Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible(FieldName + '_'      , Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
      end;
    end;

    { Dates libres }
    for i := 1 to 3 do
    begin
      FieldName       := GetPrefixe    + '_DATELIBRE' + intToStr(i);
      MasterFieldName := MasterPrefixe + '_DATELIBRE' + intToStr(i);
      if Assigned(GetControl(FieldName)) then
      begin
        SetControlProperty('T' + FieldName, 'CAPTION', RechLib(NaturePrefixe + 'DATELIBRE' + MasterPrefixe, MasterFieldName, false, true));
        SetControlVisible('T' + FieldName, Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible(FieldName      , Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible('T' + FieldName + '_', Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible(FieldName + '_'      , Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
      end;
    end;

    { Décisions libres }
    for i := 1 to 3 do
    begin
      FieldName       := GetPrefixe + '_BOOLLIBRE' + intToStr(i);
      MasterFieldName := MasterPrefixe + '_BOOLLIBRE' + intToStr(i);
      if Assigned(GetControl(FieldName)) then
      begin
        SetControlProperty(FieldName, 'CAPTION', RechLib(NaturePrefixe + 'BOOLLIBRE' + MasterPrefixe, MasterFieldName, false, false));
        SetControlVisible (FieldName, Copy(tCheckBox(GetControl(FieldName)).Caption, 1, 2) <> '.-');
      end;
    end;

    { Textes libres }
    for i := 1 to 3 do
    begin
      FieldName       := GetPrefixe + '_CHARLIBRE' + intToStr(i);
      MasterFieldName := MasterPrefixe + '_CHARLIBRE' + intToStr(i);
      if Assigned(GetControl(FieldName)) then
      begin
        SetControlText    ('T' + FieldName, RechLib(NaturePrefixe + 'CHARLIBRE' + MasterPrefixe, MasterFieldName, false, true));
        SetControlVisible ('T' + FieldName, Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible (FieldName      , Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
      end;
    end;
  end;
end;

procedure tWTOF.BRECHAFFAIRE_OnClick(Sender: TObject);
var
  Affaire,Prefixe : string;
begin
  { cas particulier élément parc, on prend le tiers de l'entete }
  if Prefixe = 'WPN' then 
    Prefixe:='WPC'
  else
    Prefixe:=GetPrefixe;
  if Assigned(GetControl(GetPrefixe + '_AFFAIRE')) then
  begin
    if   ((Ecran is TFMul) and (pos('WCBNPROP_MUL', TFMul(Ecran).Name) > 0))
      or ((Ecran is TFCube) and (pos('WCBNCBB_CUBE', TFMul(Ecran).Name) > 0)) then
      GetAffaireMul(Affaire, nil, nil, nil, nil, nil, false, true, false, true, true, false)
    else
      GetAffaireMul(Affaire, nil, nil, nil, nil, ThCritMaskEdit(GetControl(Prefixe + '_TIERS')), false, true, false, true, true, false);
    SetControlText(GetPrefixe + '_AFFAIRE', Affaire);
  end;
end;

function tWTOF.MakeWhere: string;
begin
  Result := '';
  if FTableName <> '' then
  begin
    if Assigned(FTreeTobFrame) then
      Result := FTreeTobFrame.Where
      else if wSelectAll or (not TFMul(Ecran).FListe.AllSelected) then
    begin
      {$IFNDEF EAGLCLIENT}
      Result := wMakeWhereFromList(TableToPrefixe(FTableName), TFMul(Ecran).FListe, TFMul(Ecran).Q);
      {$ELSE}
      Result := wMakeWhereFromList(TableToPrefixe(FTableName), TFMul(Ecran).FListe, TFMul(Ecran).Q.TQ);
      {$ENDIF}
    end
    else
      Result := GetWhere
  end
{$IFDEF GPAODEV}
  else
  begin
    PgiError('FTableName vide!', 'tWTOF.MakeWhere')
  end;
{$ENDIF GPAODEV}
end;

function tWTOF.GetWhere: string;
begin
  Result := RecupWhereCritere(TPageControl(GetControl('Pages')));
  Result := StringReplace (Result, 'WHERE', '', [rfIgnoreCase]);
end;

procedure tWTOF.SetDomaine;
begin
  if Assigned(GetControl('DOMAINE')) then
  begin
    SetControlText   ('DOMAINE', VH^.ProfilUserC[PrEtablissement].Domaine);
    SetControlEnabled('DOMAINE', not VH^.ProfilUserC[PrEtablissement].ForceDomaine);
  end;
end;

procedure tWTOF.SetWhereDomaine;
begin
  if Assigned(GetControl('DOMAINE')) and Assigned(GetControl('XX_WHERE_DOMAINE')) then
  begin
    if (VH^.ProfilUserC[PrEtablissement].ForceDomaine) or (not Assigned(GetControl('XX_WHERE'))) or (Pos('_DOMAINE', GetControlText('XX_WHERE')) <= 0) then
    begin
      if VH^.ProfilUserC[PrEtablissement].ForceDomaine then
        SetControlText('DOMAINE', VH^.ProfilUserC[PrEtablissement].Domaine);

      if GetControlText('DOMAINE') <> '' then
        SetControlText('XX_WHERE_DOMAINE', 'GA_DOMAINE="' + GetControlText('DOMAINE') + '" OR GA_DOMAINE=""')
      else
        SetControlText('XX_WHERE_DOMAINE', '')
    end
    else
    begin
      SetControlText('XX_WHERE_DOMAINE', '')
    end;
  end;
end;

procedure tWTOF.SetDepot;
  function GetDepot: string;
  begin
    Result := GetPrefixe + '_DEPOT';
  end;
begin
  if Assigned(GetControl(GetDepot)) then
  begin
    SetControlText   (GetDepot, VH^.ProfilUserC[PrEtablissement].Depot);
    SetControlEnabled(GetDepot, not VH^.ProfilUserC[PrEtablissement].ForceDepot);
  end;
end;

procedure tWTOF.SetEtablissement;
  function GetEtablissement: string;
  begin
    Result := GetPrefixe + '_ETABLISSEMENT';
  end;
begin
  if Assigned(GetControl(GetEtablissement)) then
  begin
    SetControlText   (GetEtablissement, VH^.ProfilUserC[PrEtablissement].Etablissement);
    SetControlEnabled(GetEtablissement, not VH^.ProfilUserC[PrEtablissement].ForceEtab);
  end;
end;

{ Permet de savoir si la grille est en multisélection }
function tWTOF.GetMultiSelected: Boolean;
begin
  if Ecran is TFMul then
    Result := wMultiSelected(TFMul(Ecran).FListe)
  else if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.MultiSelected
  else
    Result := False;
end;

procedure tWTOF.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  TFMul(Ecran).FormKeyDown(Sender, Key, Shift);

  case Key of
    VK_RETURN: if Assigned(GetControl('FListe')) and FMulDeRecherche then
                 FLISTE_OnDblCLick(TObject(GetControl('FListe')));
  end
end;

function tWTOF.EcranIs(const FormName: String): Boolean;
begin
  Result := Pos(UpperCase(Trim(FormName)), UpperCase(Ecran.Name)) = 1
end;

function tWTOF.GetMultiSelection: Boolean;
begin
  if Assigned(GetControl(FViergeGridName)) then
  begin
    {$IFDEF EAGLCLIENT}
      Result := ThGrid(GetControl(FViergeGridName)).MultiSelect
    {$ELSE EAGLCLIENT}
      Result := ThDBGrid(GetControl(FViergeGridName)).MultiSelection
    {$ENDIF  EAGLCLIENT}
  end
  else
    Result := false;
end;

procedure tWTOF.SetMultiSelection(const Value: Boolean);
begin
  if Assigned(GetControl(FViergeGridName)) then
  begin
    {$IFDEF EAGLCLIENT}
      ThGrid(GetControl(FViergeGridName)).MultiSelect := Value;
    {$ELSE EAGLCLIENT}
      ThDBGrid(GetControl(FViergeGridName)).MultiSelection := Value;
    {$ENDIF  EAGLCLIENT}
  end
end;

procedure tWTOF.PmAnalyse_OnPopUp(Sender: TObject);
  { Calcul du contexte }
  function GetArgument: string;
  begin
    if GetPrefixe = 'WOB' then
    begin
      Result := ';TYPELIEN='  + GetString('WOB_TYPELIEN')
              + ';MODECONSO=' + GetString('WOB_MODECONSO')
    end;

    Result := Result + ';ACTION=' + Action;
  end;
begin
  if IsEmpty then Abort;

  if Ecran is TFSAISIELIST then
    wSetMnuAnalyse(False, GetArgument , tPopUpMenu(GetControl('PMANALYSE')))
  else
    wSetMnuAnalyse(MultiSelected, GetArgument, tPopUpMenu(GetControl('PMANALYSE')));
end;

procedure tWTOF.BTEFFACEDATEDEC_OnClick(Sender: tObject);
begin
  SetControlText('DATEDEC', '  /  /    ')
end;

Initialization
	registerclasses([tWTOF]);
end.
