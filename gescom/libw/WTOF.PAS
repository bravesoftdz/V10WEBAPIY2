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
  uTOFComm,
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
  uTreeTobFrame
  ;

Type
  MyArrayInt = Array of Integer;


  tWTOF = Class (tTOFComm)
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
    FEditColsProperties : Boolean;
    FWListe             : TWListe;

    procedure PmUtilitaire_OnPopUp(Sender: TObject);
    procedure PmAnalyse_OnPopUp(Sender: TObject);
    procedure PmHistorique_OnPopUp(Sender: TObject);
    {$IFDEF GPAOLIGHT}
      procedure PmAction_OnPopUp(Sender: TObject);
      procedure PmBaseTech_OnPopUp(Sender: TObject);
      procedure MnJournalAction_OnClick(Sender: TObject);
    {$ENDIF GPAOLIGHT}

    { Accès au mul ou au lookup article ou tiers }
    procedure ScanControls;
    procedure ArticleElipsisClick(Sender: TObject);
    procedure TiersElipsisClick(Sender: TObject);

    { Gestion des colonnes dans les muls }
    function GetChampsMUL: String;
    procedure LoadColsList(const OtherGridName: String = '');
    function GetColWidths(const OtherGridName: String = ''): MyArrayInt;
    procedure SetEditColsProperties(const Value: Boolean);
    {$IFNDEF EAGLCLIENT}
      procedure DBGrid_OnColMoved(Sender: TObject; FromIndex, ToIndex: Longint);
    {$ENDIF !EAGLCLIENT}

    { Particularité pour les date de décomposition }
    procedure BTEFFACEDATEDEC_OnClick(Sender: tObject);

    { Gestion des ruptures dans les QR1 }
    procedure RUPT_OnClick(Sender: TObject);
    procedure SAUTRUPT_OnClick(Sender: TObject);
    procedure ChangeGroup(Const Indice: integer);
    procedure AffectGroup(Const Indice: integer);
    procedure ChangeSautPage(Const Indice: integer);
    procedure MnLpBlocNote_OnClick(Sender: TObject);
  public
    fRange 								: String;
    fNoElipsisArticle, FBlocNoteObligatoire : Boolean;

    { Set }
    procedure SetControlsVisible(Const FieldsName: Array of String; Const Visibility: Boolean);
    procedure SetControlsEnabled(Const FieldsName: Array of String; Const Enability: Boolean);
    procedure SetColVisible(const ColName: String; const Visible: Boolean; const OtherGridName: String = '');
    procedure SetColsVisible(const FieldsName: Array of String; const Visible: Boolean; const OtherGridName: String = '');
    procedure SetColCaption(const ColName, Caption: String; const OtherGridName: String = '');
    procedure SetColsCaption(const FieldsName: Array of String; const Caption: String; const OtherGridName: String = '');

    { Divers }
    procedure ReloadColsList(const OtherGridName: String = '');
    procedure wOuvreMul(Const Fiche, TableFille, Range, Param: string);

    function EcranIs(const FormName: String): Boolean;

     { BCherche }
    procedure BCherche_OnClick(Sender: TOBject);

    { Gestion des colonnes }
    property EditColsProperties: Boolean read FEditColsProperties write SetEditColsProperties;
    property wListe: TWListe read FWListe;
  protected
    Droit : string;

    function GetLoupeCtx: string; Override;

    { Gestion des colonnes }
    property ChampsMUL: String read GetChampsMUL;

    { QRS1 }
    procedure RUPT_OnEnter(Sender: TObject);

    { Recherche affaire }
    {$IFDEF AFFAIRE}
    procedure NomsChampsAffaire(Var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit); Override;
    {$ENDIF AFFAIRE}
  end ;

Implementation

uses
  AGLInitGC,
  Grids,
  {$IFNDEF BTP}
  {$IFDEF GCGC}
    yTarifs,
    yTarifsCommun,
  {$ENDIF GCGC}
  {$ENDIF}
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
  {$IFDEF STK}
    EntGC,
  {$ENDIF STK}
  FactUtil,
  UtilArticle,
  UtilPGI
  ;

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

  { -------- Gestion des colonnes -------- }
  if FEditColsProperties and (Ecran is TFSaisieList) then
  begin
    { Chargement de la liste des champs }
    LoadColsList
  end;

  {$IFDEF GPAOLIGHT}
  wGereFormsByCtx(Ecran, Self, 'OnUpdate');
  {$ENDIF GPAOLIGHT}

  SetControlProperty('BSUPPRIME'   , 'ENABLED', (Action = 'MODIFICATION') and (not IsEmpty));
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
  if FEditColsProperties and not (Ecran is TFMul) then
  begin
    FTabWidths := GetColWidths;
    { Chargement de la liste des champs }
    LoadColsList
  end;
end;

procedure tWTOF.OnArgument (S : String ) ;
var
  c: char;
begin
  { Paramètre }
  FEditColsProperties := False;
  FWListe := nil;
  FDBListe := '';

  { Action }
  Action := GetArgumentValue(S, 'ACTION');
  Droit  := GetArgumentValue(S, 'DROIT');

  { Nature par défaut des fiches }
  if fNature = '' then
    fNature := 'W';

  { Propriété par défaut }
  fMonoFiche := False;
  if Ecran is tfMul then
    TFMUL(ECRAN).CheckBoxStyle := CsCoche;

  FBlocNoteObligatoire := False;

  {$IFDEF GPAOLIGHT}
  wGereFormsByCtx(Ecran, Self, 'OnArgument', True);
  {$ENDIF GPAOLIGHT}

  if Ecran is TFMul then
    TFMul(Ecran).Caption := GetWindowsCaptionFromArgument(S, Ecran.Caption)
  else
    Ecran.Caption := GetWindowsCaptionFromArgument(S, Ecran.Caption);

  Inherited;

  { Loupe }
  if Assigned(GetControl('MNLPBLOCNOTE')) then
    TMenuItem(GetControl('MNLPBLOCNOTE')).OnClick := MnLpBlocNote_OnClick;

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

    if FMulDeRecherche and (GetTableName <> '') and Assigned(GetControl(GridName)) then
      THGrid(GetControl(GridName)).OnDblClick := FLISTE_OnDblCLick
  end;

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

  { Bouton Elipsis recherche Article }
  if not fNoElipsisArticle then
    ScanControls;

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
  if (Ecran is TFQRS1) then
  begin
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

  { Pas de filtre par défaut }
  if Ecran is TFMul then
    Tfmul(Ecran).FiltreDisabled := (Pos('FILTREDISABLED', S) > 0);

  {$IFDEF GPAO}
    { Le menu zoom est toujours contextuel à la grille en GPAO }
    if Assigned(GetControl('PMLOUPE')) and (Ecran is TFMul) then
      TFMul(Ecran).FListe.PopupMenu := TPopupMenu(GetControl('PMLOUPE'));
  {$ENDIF GPAO}

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

  AGLLanceFiche('W','WBLOCNOTE', '', '', 'ACTION=' + iif(Action = '', 'MODIFICATION', Action) + ';PREFIXE=' + Prefixe + ';ETATREV=' + GetParamEtatRev + ';WHERE=' + wFabriqueWhere(Prefixe, GetDataSet) + ';OBLIGATOIRE=' + BoolToStr_(FBlocNoteObligatoire));
  RefreshDb;
  if Ecran is TFSAISIELIST then
    TFSaisieList( ecran ).LeFiltre.Refreshlignes;
  SetMemoFrame;
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
    end
    else if GetPrefixe = 'RQN' then
    begin
     	Result := 'ETATLIG=' + GetString('RQN_ETATQNC');
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
  if GetPrefixe = 'WOL' then
  begin
   	SetLength(TabWOL, 2);
  	TabWOL[MISEENPROD] := '';
   	TabWOL[TYPEORDRE]  := '';
  	if TreeTobFrame <> nil then
    begin
    	TabWOL[MISEENPROD] := GetString('WOL_MISEENPROD');
    	TabWOL[TYPEORDRE]  := GetString('WOL_TYPEORDRE')
    end
    else
  	begin
    	{$IFDEF EAGLCLIENT}
        if GetDataSet.FieldExists('WOL_MISEENPROD') then
 			{$ELSE  EAGLCLIENT}
        if Assigned(GetDataSet.FindField('WOL_MISEENPROD')) then
 			{$ENDIF EAGLCLIENT}
		    	TabWOL[MISEENPROD] := GetString('WOL_MISEENPROD');
  		{$IFDEF EAGLCLIENT}
      	if GetDataSet.FieldExists('WOL_TYPEORDRE') then
  		{$ELSE  EAGLCLIENT}
        if Assigned(GetDataSet.FindField('WOL_TYPEORDRE')) then
 			{$ENDIF EAGLCLIENT}
		    	TabWOL[TYPEORDRE]  := GetString('WOL_TYPEORDRE');
    end;
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
Suite ........ : pour les contrôles 'articles'
NDP : Cette procédure est répliquée dans la wtom, merci de la mettre
à jour si vous modifiez cette procédure la.
*****************************************************************}
procedure TWTOF.ScanControls;
var
	iComp: Integer;
  C: TComponent;
begin
	if Assigned(Ecran) then
  begin
    for iComp := 0 to TForm(Ecran).ComponentCount - 1 do
    begin
      C := TForm(Ecran).Components[iComp];
      if (C is ThEdit) and (THEdit(C).ElipsisButton) then
      begin
        { Essai avec le suffixe du nom du champ, Si le name est pas ok, ré-essaye avec le DataType }
        if wIsDataTypeArticle(THEdit(C).DataType) or wIsDataTypeCodeArticle(THEdit(C).DataType) then
          THEdit(C).OnElipsisClick := ArticleElipsisClick
        else if wIsDataTypeTiers(THEdit(C).DataType) then
          THEdit(C).OnElipsisClick := TiersElipsisClick
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


 	if Params <> '' then Params := Params + ';';

  Params := Params + 'RECHERCHEARTICLE';

  { Appel du mul + Mise à jour du contrôle }
	DispatchRecherche(ThCritMaskEdit(Sender), 1, Plus, Params, NomMulRechArticle);
end;

procedure TWTOF.TiersElipsisClick(Sender: TObject);
begin
  DispatchRecherche(THCritMaskEdit(Sender), 2, THEdit(Sender).Plus,
                    'T_TIERS=' + Trim(GetControlText(TWinControl(Sender).Name)), '')
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
    Grid := THGrid(GetControl(GridName));
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


function TWTof.GetColWidths(const OtherGridName: String = ''): MyArrayInt;
var
  i, GridCount: Integer;
  Grid: THGrid;
begin
  if not (Ecran is TFMul) or (OtherGridName <> '') then
  begin
    SetLength(Result, 0);

    if OtherGridName <> '' then
      Grid := THGrid(GetControl(OtherGridName))
    else
      Grid := THGrid(GetControl(GridName));

    if Assigned(Grid) then
    begin
      GridCount := Grid.ColCount;

      for i := (0 + Grid.FixedCols) to Pred(GridCount) do
      begin
        SetLength(Result, Length(Result) + 1);
        Result[Length(Result) - 1] := Grid.ColWidths[i];
      end;
    end;
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 28/07/2003
Modifié le ... :   /  /
Description .. : Change le libelle de la colonne portant sur le ColName
Mots clefs ... :
*****************************************************************}
procedure tWTOF.SetColCaption(const ColName, Caption: String; const OtherGridName: String = '');
var
  iCol  : Integer;
  wField: TWFieldListe;
  locGridName: String;
begin
  if (Ecran is TFMul) and (OtherGridName = '') then
  begin
      TFMul(Ecran).SetDisplayLabel(ColName, Caption)
  end
  else if FEditColsProperties and Assigned(FWListe) then
  begin
    wField := FWListe.ChampsByName[iif(Pos('@', ColName) > 0, '(' + ColName + '(', ColName)];
    if Assigned(wField) and (wField.Index > -1) then
    begin
      iCol := wField.Index;

      if ((Ecran is TFVierge) or (Ecran is TFSaisieList)) and Assigned(GetControl(GridName)) then
        locGridName := GridName
      { cas d'une seconde grille dans un mul autre que la FListe }
      else if (OtherGridName <> '') and Assigned(GetControl(OtherGridName)) then
        locGridName := OtherGridName
      else
        locGridName := '';

      if locGridName <> '' then
      begin
        with THGrid(GetControl(locGridName)) do
          Cells[iCol + FixedCols, 0] := Caption
      end
    end;

    if Copy(Caption, 1, 2) = '.-' then
      SetColVisible(ColName, False, OtherGridName);
  end
end;
procedure tWTOF.SetColsCaption(const FieldsName: array of String; const Caption: String; const OtherGridName: String = '');
var
  i: Integer;
begin
  for i := Low(FieldsName) to High(FieldsName) do
    SetColCaption(FieldsName[i], Caption, OtherGridName)
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
        Result := 'DEPOT='           + GetString('GQD_DEPOT')
                + ';ARTICLE='        + GetString('GQD_ARTICLE')
                + ';STATUTDISPO='    + GetString('GQD_STATUTDISPO')
                + ';STATUTFLUX='     + GetString('GQD_STATUTFLUX')
                + ';REFAFFECTATION=' + GetString('GQD_REFAFFECTATION')
                + ';CONTREMARQUE='   + GetString('GQD_CONTREMARQUE')
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
          sPeriodeSauv := wPadRight(wDateTimeToStr(V_PGI.DateEntree+time,wCalculPeriodeSauvegarde(GetString('WPE_UNPDRPAR')), False), iLongueurPeriodeSauvegarde, '_');
          Result:= 'PERIODESAUVOK='+iif(GetString('WPE_PERIODESAUV')=sPeriodeSauv, wTrue, wFalse);
        end
        else
          Result:= 'PERIODESAUVOK=X';
      end
    {$ENDIF GPAO}
    {$IFDEF EDI}
      else if GetPrefixe = 'ELI' then
      begin
        Result := 'IDENTIFIANTELI=' + GetString('ELI_IDENTIFIANT')
                + ';NATUREPIECEG='  + GetString('ELI_NATUREPIECEG')
      end
      else if GetPrefixe = 'EDD' then
      begin
        Result := 'IDENTIFIANTELI=' + GetString('EDD_IDENTIFIANTELI')
                + ';NATUREPIECEG='  + wGetSQLFieldValue('ELI_NATUREPIECEG', 'EDILIGNE', 'ELI_IDENTIFIANT=' + GetString('EDD_IDENTIFIANTELI'))
      end
    {$ENDIF EDI}
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
  FEditColsProperties := Value or (Ecran is TFMul);
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
        GridName := 'FListe';
        LoadColsList
      end;
      if Assigned(GetControl(GridName)) then
        THGrid(GetControl(GridName)).Options := THGrid(GetControl(GridName)).Options - [goColMoving]
    end;
  end
end;

procedure tWTOF.LoadColsList(const OtherGridName: String = '');
var
  DBLst: String;
begin
  if (Ecran is TFVierge) and (GridName <> '') and Assigned(GetControl(GridName)) then
    DBLst := THGrid(GetControl(GridName)).ListeParam
  else if (Ecran is TFSaisieList) then
    DBLst := TFSaisieList(Ecran).DBListe
  else if (OtherGridName <> '') and Assigned(GetControl(OtherGridName)) then
    DBLst := THGrid(GetControl(OtherGridName)).ListeParam
  else
    DBLst := '';

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
procedure tWTOF.SetColVisible(const ColName: String; const Visible: Boolean; const OtherGridName: String = '');
var
  iCol  : Integer;
  wField: TWFieldListe;
  locGridName: String;
begin
  if Assigned(TreeTobFrame) then
    TreeTobFrame.ColVisible[TreeTobFrame.ColField[ColName]] := Visible
  else if (Ecran is TFMul) and (OtherGridName = '') then
  begin
  	TFMul(Ecran).SetVisibleColumn(ColName, Visible)
  end
  else if FEditColsProperties and Assigned(FWListe) then
  begin
    wField := FWListe.ChampsByName[StringReplace(ColName, ';', '~', [rfIgnoreCase, rfReplaceAll])];
    if Assigned(wField) and (wField.Index > -1) then
    begin
      iCol := wField.Index;

      if ((Ecran is TFVierge) or (Ecran is TFSaisieList)) and Assigned(GetControl(GridName)) then
        locGridName := GridName
      { cas d'une seconde grille dans un mul autre que la FListe }
      else if (OtherGridName <> '') and Assigned(GetControl(OtherGridName)) then
        locGridName := OtherGridName
      else
        locGridName := '';

      if locGridName <> '' then
      begin
        with THGrid(GetControl(locGridName)) do
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
    end
  end
end;

procedure twTof.SetColsVisible(const FieldsName: Array of String; const Visible: Boolean; const OtherGridName: String = '');
var
  i: Integer;
begin
  for i := Low(FieldsName) to High(FieldsName) do
    SetColVisible(FieldsName[i], Visible, OtherGridName)
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/03/2004
Modifié le ... : 02/03/2004
Description .. : Recharge la liste associée à un THGrid.
Suite ........ : N'a d'utilité que dans une TFVierge
Mots clefs ... : 
*****************************************************************}
procedure tWTOF.ReloadColsList(const OtherGridName: String = '');
begin
  if FEditColsProperties then
  begin
    LoadColsList(OtherGridName);
    if not (Ecran is TFMul) or (OtherGridName <> '') then
      FTabWidths := GetColWidths(OtherGridName)
  end
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

function tWTOF.EcranIs(const FormName: String): Boolean;
begin
  Result := Pos(UpperCase(Trim(FormName)), UpperCase(Ecran.Name)) = 1
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

function tWTOF.GetLoupeCtx: string;
{$IFDEF BTP}
begin
{$ELSE}
Const
  sCannotDoActionOnTypeOrdre = '[STA];[VTE]';
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
    Result := 'WITHPIECE=' + BoolToStr_((GetString (GetPrefixe + '_NATUREPIECEG') <> '') and
                                       (GetString (GetPrefixe + '_SOUCHE')       <> '') and
                                       (GetInteger(GetPrefixe + '_NUMERO')       <> 0 ))
  end;

begin
  Result := inherited GetLoupeCtx;

  { Laisser en premier pour concordance if...else if... et directives de compil. }
  if GetPrefixe = 'GVP' then
  begin
    Result := Result
            + 'APPEL=VALORISATION;'
            + 'FAMILLEVALO=' + GetString('GVP_FAMILLEVALO')+';'
            + 'FAMILLENIV1=' + GetString('GVP_FAMILLENIV1')+';'
            + 'FAMILLENIV2=' + GetString('GVP_FAMILLENIV2')+';'
            + 'FAMILLENIV3=' + GetString('GVP_FAMILLENIV3')+';'
  end
  {$IFDEF GPAO}
    else if GetPrefixe = 'WOP' then
    begin
      TypeOrdre:= wGetSqlFieldValue('WOL_TYPEORDRE','WORDRELIG','WOL_NATURETRAVAIL="'+GetString('WOP_NATURETRAVAIL')+'" AND WOL_LIGNEORDRE='+GetString('WOP_LIGNEORDRE'));
      Result := Result
              + iif((Pos('[' + TypeOrdre + ']', sCannotDoActionOnTypeOrdre) = 0) and (GetString('WOP_TIERS') <> ''),';TYPEORDRE=STP','')
              + ';ETATPHASE=' + GetString('WOP_ETATPHASE');
    end
    else if GetPrefixe = 'WOT' then
    begin
      Result := Result
              + GetCtxWithPiece
    end
    else if GetPrefixe = 'WOL' then
    begin
      Result := Result
              + 'ORDREPERE=' + GetString('WOL_ORDREPERE')
              + ';' + GetCtxWithPiece
    end
    else if GetPrefixe = 'WOG' then
      Result := Result + 'ETATOPE=' + GetString('WOG_ETATOPE')
    else if GetPrefixe = 'WOR' then
      Result := Result + 'ETATOPE=' + wGetSqlFieldValue('WOG_ETATOPE', 'WORDREGAMME', 'WOG_NATURETRAVAIL="' + GetString('WOR_NATURETRAVAIL') + '" AND WOG_LIGNEORDRE=' + IntToStr(GetInteger('WOR_LIGNEORDRE')) + ' AND WOG_OPECIRC="' + GetString('WOR_OPECIRC') + '" AND WOG_NUMOPERGAMME=' + IntToStr(GetInteger('WOR_NUMOPERGAMME')))
  {$ENDIF GPAO}
  {$IFDEF GCGC}
    else if GetPrefixe = 'YTS' then
    begin
      Result := Result + 'APPEL=TARIF; '+GetArgumentYTS;
    end
  {$ENDIF GCGC}
  {$IFDEF STK}
  else if GetPrefixe = 'GST' then
  begin
    Result := Result
            + 'ARTICLE='       + GetString('GST_ARTICLE')
            + ';LOTINTERNE='   + GetString('GST_LOTINTERNE')
            + ';SERIEINTERNE=' + GetString('GST_SERIEINTERNE')
            + ';TIERS='        + GetString('GST_TIERS')
  end
  else if GetPrefixe = 'GSM' then
  begin
    if Assigned(GetDataSet.FindField('GSM_ARTICLE')) then
      Result := Result + 'ARTICLE=' + GetString('GSM_ARTICLE')
    else
      Result := Result + 'ARTICLE=' + GetString('GA_ARTICLE');
    Result := Result
            + ';LOTINTERNE='   + GetString('GSM_LOTINTERNE')
            + ';SERIEINTERNE=' + GetString('GSM_SERIEINTERNE');
  end
  else if GetPrefixe = 'GQD' then
  begin
    Result :=  Result
            + ';ARTICLE='      + GetString('GQD_ARTICLE')
            + ';DEPOT='        + GetString('GQD_DEPOT')
            + ';LOTINTERNE='   + GetString('GQD_LOTINTERNE')
            + ';SERIEINTERNE=' + GetString('GQD_SERIEINTERNE')
  end
  {$ENDIF STK}
  {$IFDEF QUALITE}
  else if GetPrefixe = 'RQN' then
  begin
    Result := ';ARTICLE='       + GetString('RQN_ARTICLE')
            + ';TIERS='   		  + GetString('RQN_TIERS')
            + ';LIGNEORDRE='    + GetString('RQN_LIGNEORDRE')
            + ';NATURETRAVAIL=' + GetString('RQN_NATURETRAVAIL')
            + ';CLEPIECE='  		+ GetString('RQN_PIECEORIGINE')
  end
  else if GetPrefixe = 'RQD' then
  begin
    Result := ';TIERS='    + GetString('RQD_TIERS')
            + ';QNCNUM='   + GetString('RQD_QNCNUM')
            + ';IDACTION=' + GetString('RQD_IDACTION')
  end
  else if GetPrefixe = 'RAC' then
  begin
    Result := ';DEMDEROGNUM=' + GetString('RAC_QDEMDEROGNUM')
            + ';PLANCORRNUM=' + GetString('RAC_QPLANCORRNUM')
            + ';QNCNUM='      + GetString('RAC_QNCNUM')
  end
  {$ENDIF QUALITE}
{$ENDIF}
end;

{$IFDEF AFFAIRE}
procedure tWTOF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  inherited;
  Aff   := THEdit(GetControl(GetPrefixe + '_AFFAIRE'));
  Aff0  := THEdit(GetControl('AFFAIRE0'));
  Aff1  := THEdit(GetControl('AFFAIRE1'));
  Aff2  := THEdit(GetControl('AFFAIRE2'));
  Aff3  := THEdit(GetControl('AFFAIRE3'));
  Aff4  := THEdit(GetControl('AVENANT'));
  Tiers := THEdit(GetControl(GetPrefixe + '_TIERS'));
end;
{$ENDIF AFFAIRE}

Initialization
	registerclasses([tWTOF]);
end.
