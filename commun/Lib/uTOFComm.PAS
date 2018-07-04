{***********UNITE*************************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 20/01/2006
Description .. : TOF Commune à tous les métiers
Mots clefs ... : uTOFComm

TOF            : TOF AGL
 |
 +- tTOFComm   : TOF Commune à tous les métiers
     |           - Gestion des domaines sur les muls
     |           - Gestion des dépôts sur les muls
     |           - Gestion des établissements sur les muls
     |           - Gestion de l'onglet dimensions (voir utilisation sur wArticleMul_TOF)
     |           - Gestion des libellés des tables libres
     |           - Gestion de la fiche propriétés
     |           - Gestion de la loupe contextuelle (Voir getLoupeCtx pour contexte à donner à wMnu.SetMnuLoupe)
     |           - Boutons "Standard":
     |              > Bouton BTUSER        : Assigne le user connecté au contrôle associé
     |              > Bouton BINSERT       : Appel de la fiche associée en mode Insert
     |              > Bouton BDELETE       : Suppression depuis un mul
     |              > Bouton BTSHOWBLOCNOTE: Affichage du bloc-note en bas du mul
     |           - Méthodes:
     |              > GetWhere    : Donne le where global du mul
     |              > MakeWhere   : Donne le Where du mul en fonction des sélections
     |              > RefreshDB   : Rafraîchit le mul en gardant la mémoire de la ligne
     |              > IsEmpty     : Dit si le mul est vide
     |              > GotoNewLine : Se positionne sur une nouvelle ligne
     |              > OuvreFiche  : Appel de la fiche liée au mul
     |              > EcranIs     : Dis oui si c'est l'acran demandé qui est actif
     |
     +- twTOF  : TOF Métier W (Gestion de Production)
*****************************************************************}
Unit uTOFComm;

Interface

Uses
  Variants,
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
  Vierge,  
  GRS1,
  graphics,
  stat,
  Messages,
  ENtGC,
	uTreeTobFrame
  ;

Type
  tTOFComm = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
  private
    fGridName       : String;         { Nom de la grille }
    fTreeTobFrame   : TFFrameTreeTob; { Fiche treeView }

{$IFDEF AFFAIRE}
    fCreatAff: Boolean; {mode création de l'affaire}
    fSaisieAffaire:  boolean;
    fProposition: Boolean;
    fActionFiche : TActionFiche;
{$ENDIF AFFAIRE}
    { Get }
    function GetMultiSelection: Boolean;
    function GetMultiSelected: Boolean;
    function GetRechDom(DataType, Fieldname : string; Abrege : Boolean; NotCheckBox : Boolean; Plus : string = '') : string;

    { Set }
    procedure SetDomaine;
    procedure SetDepot;
    procedure SetEtablissement;

    { Dimensions }
    procedure GA_STATUTART_OnChange(Sender: TObject);
    procedure GA_GRILLEDIM_OnChange(Sender: TObject);
{$IFDEF AFFAIRE}
    {Code Affaire}
    procedure EffaceAffaire(EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4: THEdit);
    procedure bSelectAffMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{$ENDIF AFFAIRE}
  protected
    Action: string;

    function GetLoupeCtx: string; Virtual;

    property GridName: String read FGridName write FGridName;
{$IFDEF AFFAIRE}
    { Code Affaire }
    procedure bEffaceRef(Sender: TObject);
    procedure NomsChampsAffaireREf(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4,Tiers: THEdit); virtual;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); virtual;
    procedure TypeAppelChargeCle(var AvActionFiche: TActionFiche); virtual;
    procedure bSelectAff1Click(Sender: TObject); virtual;
    procedure bSelectAff2Click(Sender: TObject); virtual;
    procedure bSelectRefClick(Sender: TObject); virtual;
    procedure bEffaceAff1Click(Sender: TObject); virtual;
    procedure bEffaceAff2Click(Sender: TObject); virtual;
    procedure SelectionAffaire(EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4: THEdit;
      bChangeStatut: boolean = false; bModele: boolean = false; sContexte: string = '';
      bReduit: boolean = false; bOuvreSiUn: boolean = true; bItDatesMax: boolean = false;
      fQueAffaire: boolean = false; fProposition: boolean = false; fEcole: boolean = false; IsAffaireRef: boolean = false
      ; AvecPc: boolean = false; fTypeAffaire: string='');

      {$IFDEF GPAO}
        procedure EditAFFExit(Sender:TOBject);
      {$ENDIF GPAO}
{$ENDIF AFFAIRE}
  public
    fNature               : string;  { Nature de la fiche associée au mul }
    fFiche                : string;  { Fiche associée au mul }
    fParamsLanceFiche     : string;  { Paramètres à passer à la fiche}
    fMonoFiche            : Boolean; { MonoFiche forcé }
    StArgument            : String;  { Argument passé à la fiche }
    fTableName            : string;  { Nom de la table maître du MUL }
    fLequel               : string;  { Clef 1 de la table maître }
    fSelectAll            : Boolean; { Sélection avec déselection possible (Mul avec peu de données) }
    fMulDeRecherche       : Boolean; { Le mul sert à faire de la recherche }
		fMulDeTraitement			: Boolean; { Le mul sert à faire des traitements}
    fBlobSuffixe          : String;  { Suffixe du bloc note (Champ blob)}
    fMessageToDelete      : String;  { Message avant de Deleter }
    fDeleteEvent          : String;  { Type d'évènement pour le journal }
    fSuffixeUserForBtUser : String;  { Suffixe du champ relié au bouton USER }
    fPrefixeTable 				: String;  { Préfixe de la table liée à la tof }
    fSansRestrictEtab 		: Boolean; { Possibilité de ne pas appliquer la restriction sur les établissements }
{$IFDEF AFFAIRE}
    BVoirStatut						: Boolean; {indique dans le mul concerné si on voit propo et affaire}
    BEcole								: Boolean; {Permet de voir que les affaires "ecole"}
    BChangeTiers					: Boolean;
    fQueAffaire: Boolean; {permet de ne voir que les "vraies" affaires AFF, pas les affaires associées
                          aux piéces gescom pour le planning}
    bEffaceAff1						: TToolBarButton97;
    bEffaceAff2						: TToolBarButton97;
    bSelectAff1						: TToolBarButton97;
    bSelectAff2						: TToolBarButton97;
    bSelectRef						:  TToolBarButton97;
    EditTiers							: THEdit;
    EditTiers_						: THEdit;
    EditAff								: THEdit;
    EditAff0							: THEdit;
    EditAff1							: THEdit;
    EditAff2							: THEdit;
    EditAff3							: THEdit;
    EditAff4							: THEdit;
    EditAff_							: THEdit;
    EditAff0_							: THEdit;
    EditAff1_							: THEdit;
    EditAff2_							: THEdit;
    EditAff3_							: THEdit;
    EditAff4_							: THEdit;
    EditAffref						: THEdit; //mcd 05/12/2006 13490
    EditAffref0						: THEdit;
    EditAffref1						: THEdit;
    EditAffref2						: THEdit;
    EditAffref3						: THEdit;
    EditAffref4						: THEdit;
    EditTiersRef					: THEdit;
    fTypeAffaire          : string;
    fRessourceCourante    : string;
    fTypeTiers            : string;
{$ENDIF AFFAIRE}

    { Evénements génériques }
    procedure FLISTE_OnDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BINSERT_OnClick(Sender: TOBject);
    procedure BSELECTALL_OnClick(Sender: TObject);
    procedure BDELETE_OnClick(Sender: TObject);
    procedure BTUSER_OnClick(Sender: TObject);
    procedure BTSHOWBLOCNOTE_OnClick(Sender: tObject);
    procedure BchercheClick (Sender : Tobject);
    { Loupe }
    procedure PmLoupe_OnPopUp(Sender: TObject);
    procedure MnProperties_OnClick(Sender: TObject);

    { Get }
    function GetWhere: string;
    function GetQuery: tQuery;
    function GetTableName: string;
    function GetPrefixe: string;
    function GetInteger(Const FieldName: string): integer;
    function GetString(Const FieldName: string): string;
    function GetDateTime(Const FieldName: string): tDateTime;
    function GetDouble(Const FieldName: string): Double;
    function GetBoolean(Const FieldName: string): Boolean;

    { Set }
    procedure SetMemoFrame(const aValueClef1: String = '');
    procedure SetLibelle;
    procedure SetWhereDomaine;
    procedure SetMultiSelection(const Value: Boolean);

    { Méthodes génériques }
    function MakeWhere: String;
    procedure RefreshDB (lastSave : string=''; WithSearch:boolean=true);
    function IsEmpty: Boolean;
    function GotoNewLine(Save: String) : boolean;
    function OuvreFiche(Const Fiche, Lequel: string; Action: string; Const Param: string): String;
    function FieldExists(const FieldName: string): boolean;

    { Dimensions }
    procedure SetDimensions;
    procedure SetArticle;

    function EcranIs(const FormName: String): Boolean;

    property TreeTobFrame: TFFrameTreeTob read FTreeTobFrame write FTreeTobFrame;
    property MultiSelection: Boolean read GetMultiSelection write SetMultiSelection;
    property MultiSelected: Boolean read GetMultiSelected;
    property SuffixeUserForBtUser: String read FSuffixeUserForBtUser write FSuffixeUserForBtUser;

{$IFDEF AFFAIRE}
    {Code Affaire}
    procedure SetCodeAffaire(stArgument: string);
{$ENDIF AFFAIRE}

  end;

Implementation

uses
  wCommuns,
  {$IFNDEF PGIMAJVER}
    wMnu,
  {$ENDIF PGIMAJVER}
  wRapport,
  UtilArticle,
  UtilGC,
  Ent1,
  UtilPGI,
  UtilConfid, {JTR - Confidentialité sur la fiche}
  HRichOle
  {$IFDEF AFFAIRE}
	,Dicobtp,
	AffaireUtil,
	FactUtil,
	CalcOLEGenericAFF
  {$ENDIF AFFAIRE}
  ;

{ tTOFComm }

procedure tTOFComm.FLISTE_OnDblClick(Sender: TObject);
var
  Lequel, sClef: String;
begin
  { Abort }
  if isEmpty then
    Abort;

  if not FMulDeRecherche then
  begin
    if (FFiche = '') then
      Abort;
    { Lequel }
    if Assigned(FTreeTobFrame) then
      Lequel := wGetValueClef1(FTreeTobFrame.CurrentTob)
    else
      Lequel := WFabriqueRange(FTableName, FLequel, GetDataSet);

    { Ajout du Lequel }
    if FParamsLanceFiche = '' then
      FParamsLanceFiche := '[' + Lequel + ']';
    if GetArgumentString(FParamsLanceFiche, 'ACTION') = 'CREATION' then
      FParamsLanceFiche := StringReplace(FParamsLanceFiche, 'ACTION=CREATION', 'ACTION=MODIFICATION', []);

    OuvreFiche(fFiche, Lequel, iif(Pos('ACTION', fParamsLanceFiche) > 0, '', 'ACTION=' + iif(Action = '', 'MODIFICATION', Action)), FParamsLanceFiche)
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

procedure tTOFComm.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  TFMul(Ecran).FormKeyDown(Sender, Key, Shift);

  case Key of
    VK_RETURN:  if Assigned(GetControl('FListe')) and FMulDeRecherche then
                begin
                  FLISTE_OnDblCLick(TObject(GetControl('FListe')));
                end;
  end
end;

function tTOFComm.GetBoolean(const FieldName: string): Boolean;
begin
  if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetBoolean(FieldName)
  else if GetField(FieldName) <> null then
    Result := StrToBool_(GetField(FieldName))
  else
  	Result := False;
end;

function tTOFComm.GetDateTime(const FieldName: string): tDateTime;
begin
  if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetDateTime(FieldName)
  else if GetField(FieldName) <> null then
   	Result := GetField(FieldName)
  else
   	Result := iDate1900;
end;

function tTOFComm.GetDouble(const FieldName: string): Double;
begin
	if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetDouble(FieldName)
  else if GetField(FieldName) <> null then
   	Result := GetField(FieldName)
  else
   	Result := 0;
end;

function tTOFComm.GetInteger(const FieldName: string): integer;
begin
  if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetInteger(FieldName)
  else if GetField(FieldName) <> null then
   	Result := GetField(FieldName)
  else
   	Result := 0;
end;

function tTOFComm.GetMultiSelection: Boolean;
begin
  if Assigned(GetControl(GridName)) then
  begin
    {$IFDEF EAGLCLIENT}
      Result := ThGrid(GetControl(GridName)).MultiSelect
    {$ELSE EAGLCLIENT}
      Result := ThDBGrid(GetControl(GridName)).MultiSelection
    {$ENDIF  EAGLCLIENT}
  end
  else
    Result := false;
end;

function tTOFComm.GetMultiSelected: Boolean;
begin
  if Ecran is TFMul then
    Result := wMultiSelected(TFMul(Ecran).FListe)
  else if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.MultiSelected
  else if Ecran is TFSaisieList then
    Result := wMultiSelected(TFSaisieList(Ecran).FListe)
  else
    Result := False;
end;

function tTOFComm.GetPrefixe: string;
begin
	if (fPrefixeTable <> '') then
		Result := fPrefixeTable
	else
  	Result := TableToPrefixe(GetTableName);
end;

function tTOFComm.GetQuery: tQuery;
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

function tTOFComm.GetString(const FieldName: string): string;
begin
  if Assigned(FTreeTobFrame) then
    Result := FTreeTobFrame.GetString(FieldName)
	else if GetField(FieldName) <> null then
   	Result := GetField(FieldName)
  else
   	Result := '';
end;

function tTOFComm.GetTableName: string;
begin
  if Ecran is TFSAISIELIST then
    Result := TFSaisieList(Ecran).LeFiltre.TableEntete
  else
    Result := FTableName;
end;

function tTOFComm.GetWhere: string;
begin
  Result := RecupWhereCritere(TPageControl(GetControl('Pages')));
  Result := StringReplace (Result, 'WHERE', '', [rfIgnoreCase]);
end;

function tTOFComm.GotoNewLine(Save: String) : boolean;
var
  RV            : Array of Variant ;
  S, St, sClef1 : String ;
  Indice        : Integer ;
  {$IFDEF EAGLCLIENT}
    DS    : TDataSet;
    iRec  : integer;
    Found,
    Cancel: Boolean;
  {$ENDIF EAGLCLIENT}

  function NombreDeToken ( s: String ) : Integer ;
  begin
    Result := 0;
    while ReadTokenSt(S) <> '' do
      Inc(Result);
  end;

  procedure DoOnDisplay();
  begin
    {$IFDEF EAGLCLIENT}
      TFMul(Ecran).FListeRowEnter(TFMul(Ecran).FListe, TFMul(Ecran).FListe.Row, Cancel, False);
    {$ELSE  EAGLCLIENT}
      TFMul(Ecran).FListeRowEnter(TFMul(Ecran).FListe);
    {$ENDIF EAGLCLIENT}
  end;

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
        if DS.Eof and TFMul(Ecran).BNext.Enabled then
        begin
          TFMul(Ecran).BNext.Click;
          DS := GetDataSet;
        end
      end;
      if (DS.CurrentFilleIndex + 1) < TFMul(Ecran).FListe.RowCount then
        TFMul(Ecran).FListe.Row := DS.CurrentFilleIndex + 1
      else
        TFMul(Ecran).FListe.Row := TFMul(Ecran).FListe.RowCount-1;
    {$ELSE}
		  result := GetDataSet.Locate(sClef1, Save, [ LoCaseInsensitive ]);
    {$ENDIF}
    { pour provoquer le OnDisplay de la Tof }
    DoOnDisplay()
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
        S := sCLef1;
        Indice := 0 ;
        St:= Trim(ReadTokenSt(S));
        Found := true;
        while St <> '' do
        begin
          Inc ( Indice ) ;
          if RV[Indice-1]<>DS.FindField(St).value then
          begin
            Found := False;
            Break;
          end;
          St := Trim(ReadTokenSt(S)) ;
        end;
        if Found then
          Break;
        DS.Next;
        if DS.Eof and TFMul(Ecran).BNext.Enabled then
        begin
          TFMul(Ecran).BNext.Click;
          DS := GetDataSet;
        end
      end;
      if not Found then
        DS.Seek(iRec)
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
  		result := GetDataSet.Locate ( sClef1, RV, [ LoCaseInsensitive ] ) ;
    {$ENDIF EAGLCLIENT}
		SetLength ( RV, 0 ) ;
    { pour provoquer le OnDisplay de la Tof }
    DoOnDisplay()
	end;
end;

function tTOFComm.IsEmpty: Boolean;
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

function tTOFComm.MakeWhere: String;
begin
  Result := '';
  if FTableName <> '' then
  begin
    if Assigned(FTreeTobFrame) then
      Result := FTreeTobFrame.Where
    else if fSelectAll or (not tToolBarButton97(GetControl('BSELECTALL')).Down) then
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

procedure tTOFComm.OnArgument(S: String);
var
  c: Char;
	BParamListe : TToolbarButton97;
begin

		BparamListe := TToolbarButton97 (GetCOntrol('BparamListe'));

{Uniquement en line
    if V_PGI.Sav then
    begin
    	if BparamListe <> nil then BParamListe.Visible := True
    end
    else
    begin
    	if BparamListe <> nil then BParamListe.Visible := False;
    end;
}

  StArgument            := S;
  fBlobSuffixe          := 'BLOCNOTE';
  fSuffixeUserForBtUser := 'CREATEUR';
	fPrefixeTable					:= '';
  fDeleteEvent          := '';
  GridName := 'FListe';

  inherited;

     // MB : Pour KPMG, ne pas chercher automatiquement
  if (Ecran is tfMUL) and (GetParamSocSecur ('SO_AFCLIENT', 0) =cInClientKPMG ) then   //mcd 18/02/08: test = au lieu de <>
        TFMul(Ecran).AutoSearch :=  asMouette ;


  { Action }
  Action := GetArgumentValue(S, 'ACTION');

  { Propriété par défaut }
  fSelectAll := False;

  { Evénements standard }
	// PL le 22/02/06 : ajout de FMulDeTraitement pour la gestion des MUL de traitement
  if (Ecran is TFMul) and not fMulDeRecherche and not fMulDeTraitement then
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
  end;

  { Loupe }
  if Assigned(GetControl('PMLOUPE')) then
    TPopUpMenu(GetControl('PMLOUPE')).OnPopUp := PmLoupe_OnPopUp;
  if Assigned(GetControl('MNPROPERTIES')) then
    TMenuItem(GetControl('MNPROPERTIES')).OnClick := MnProperties_OnClick;

  { Raccourci clavier }
  if Assigned(GetControl('MNLPDISPO')) then
    TMenuItem(GetControl('MNLPDISPO')).ShortCut := TextToShortCut('Ctrl+S');
  if Assigned(GetControl('MNDISPO')) then
    TMenuItem(GetControl('MNDISPO')).ShortCut := TextToShortCut('Ctrl+S');
  { Raccourci clavier }
  if Assigned(GetControl('MNLPDISPO')) then
    TMenuItem(GetControl('MNLPDISPO')).ShortCut := TextToShortCut('Ctrl+S');
  if Assigned(GetControl('MNDISPO')) then
    TMenuItem(GetControl('MNDISPO')).ShortCut := TextToShortCut('Ctrl+S');
  if Assigned(GetControl('MNLPARTICLE')) then
    TMenuItem(GetControl('MNLPARTICLE')).ShortCut := TextToShortCut('Ctrl+A');

  { Bloc-notes }
  if Assigned(GetControl('BTSHOWBLOCNOTE')) then
    TToolBarButton97(GetControl('BTSHOWBLOCNOTE')).OnClick := BTSHOWBLOCNOTE_OnClick;
  if Assigned(GetControl('SPLITPNMEMO')) then
  begin
    ThSplitter(GetControl('SPLITPNMEMO')).Height := 3;
    ThSplitter(GetControl('SPLITPNMEMO')).Beveled := true;
  end;


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

  { Libellés paramétrés}
  if (Ecran is TFQRS1) or (Ecran is TFStat) or (Ecran is TFVierge) then
    SetLibelle;

  { Hint imposés }
  //C.B 29/05/2007 
  {$IFNDEF AFFAIRE}
  SetControlProperty('BDELETE', 'Hint', TraduireMemoire('Supprimer la sélection'));
  {$ENDIF AFFAIRE}
  SetControlProperty('BTUSER', 'Hint', TraduireMemoire('Renseigner avec l''utilisateur connecté'));

  { Par défaut selon paramétrage utilisateur/groupe}
  SetDomaine;
  SetDepot;
  SetEtablissement;

  { Confidentialité }
  AppliquerConfidentialite(Ecran, ''); {JTR - Confidentialité sur la fiche}

{$IFDEF AFFAIRE}
  {Code Affaire}
  SetCodeAffaire(StArgument);
{$ENDIF AFFAIRE}
  if (Ecran is TFMul) then
  begin
		if Ttoolbarbutton97(GetControl('Bcherche')) <> nil then Ttoolbarbutton97(GetControl('Bcherche')).Onclick := BchercheClick;
  end;
  Ecran.Caption := GetWindowsCaptionFromArgument(S, Ecran.Caption);
  UpdateCaption(Ecran);
end;

procedure tTOFComm.OnClose;
begin
  inherited;

end;

procedure tTOFComm.OnDelete;
begin
  inherited;

end;

procedure tTOFComm.OnDisplay;
 begin
  inherited;

  { Bloc-notes }
  SetMemoFrame;
end;

procedure tTOFComm.OnLoad;
begin
  inherited;

  { Déduction du domaine }
  SetWhereDomaine;

  { Libellé article }
  if Assigned(GetControl('GA_CODEARTICLE')) then
    THEdit(GetControl('GA_CODEARTICLE')).UpdateLibelle;
end;

procedure tTOFComm.OnNew;
begin
  inherited;
end;

procedure tTOFComm.OnUpdate;
begin
  inherited;
  { Pour être sûr de réassigner les événements }
  if Assigned(Ecran) and (Ecran is TFMul) and (FMulDeRecherche) and (GetTableName <> '') and (not MultiSelection) then
  begin
    if Assigned(GetControl('FLISTE')) then
    begin
      THGrid(GetControl('FLISTE')).OnDblClick := FLISTE_OnDblCLick;
      Ecran.OnKeyDown := FormKeyDown;
    end;
    if Assigned(GetControl('BOUVRIR')) then
      TToolBarButton97(GetControl('BOUVRIR')).OnClick := FLISTE_OnDblCLick;
  end;

  { Libellés paramétrés}
  if ecran is TFMUL then
    SetLibelle;

  { Dimensions }
  SetDimensions;

  { Enabled }
  SetControlProperty('BDELETE'     , 'ENABLED', (Action = 'MODIFICATION') and (not IsEmpty));
	// PL le 22/02/06 : ajout de FMulDeTraitement pour la gestion des MUL de traitement
  SetControlProperty('BSELECTALL'  , 'ENABLED',  ((FMulDeTraitement) or (Action = 'MODIFICATION')) and (not IsEmpty));
end;

function tTOFComm.OuvreFiche(const Fiche, Lequel: string; Action: string; const Param: string): String;
var
	Argument: String;
begin
  { Action }
  if (Action <> '') and (Pos('ACTION', Action) <= 0) then
    Action := 'ACTION=' + Action;

  { Argument }
  Argument := Action;
  if fMonoFiche then
    Argument := iif(Argument<>'', Argument+';', '')+'MONOFICHE';
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
    	(*
      TFMUL(ECRAN).BChercheClick (Nil);
      { Repositionnement }
      GotoNewLine(Result);
      *)
      Refreshdb (Result);
    end;
  end
  else
  begin
    if ((Fiche <> 'WARTICLE_FIC') and (FNature<>'GC')) or GetParamSocSecur ('SO_REFRESHMUL', true) then RefreshDB;
  end;

  { Si le mémo est à afficher }
  SetMemoFrame;
end;

procedure tTOFComm.RefreshDB (LastSave : string; WithSearch:boolean);
var save : string;
    DS    : TDataSet;
    found : boolean;
begin
	found := false;
  if Assigned(Ecran) and (Ecran is TFMul) then
  begin
    if isEmpty then
    begin
      TFMUL(Ecran).BChercheClick(Nil);
    end else
    begin
    	if WithSearch then
      begin
        if LastSave <> '' then Save := LastSave
                          else Save := WGetValueClef1(fTableName, TFMUL(Ecran), fLequel);
      end;
      TFMUL(Ecran).BChercheClick(Nil);
      if (WithSearch) and (Save <> '') then
      begin
      	found := GotoNewLine(Save);
      end;
    end;
    if not found then
    begin
      DS := GetDataSet;
      if not DS.Eof then DS.First;
    end;

    { On remet à Down le bouton de multi-sélection }
    if fSelectAll then
      SetControlProperty('BSELECTALL', 'DOWN', false);
  end
  else if (Ecran is TFVierge) and Assigned(FTreeTobFrame) then
    FTreeTobFrame.RefreshTreeTob
end;

procedure tTOFComm.SetMemoFrame(const aValueClef1: String = '');
var
  C: TControl;
  s: HString;

  function GetMemoValue: String;
  var
    Cle1, ValueCle1, Where: String;
  begin
    Cle1 := wMakeFieldString(GetTableName, ';');
    if aValueClef1 = '' then
      ValueCle1 := wGetValueClef1(GetTableName, Ecran, fLequel)
    else
      ValueCle1 := aValueClef1;
    Where := wMakeWhereSQL(Cle1, ValueCle1);

    Result := wGetSQLFieldValue(GetPrefixe + '_' + fBlobSuffixe, GetTableName, Where);
  end;

begin
  if Assigned(GetControl('BTSHOWBLOCNOTE')) and Assigned(GetControl(GetPrefixe + '_' + fBlobSuffixe)) then
  begin
    if TToolBarButton97(GetControl('BTSHOWBLOCNOTE')).Down then
    begin
      C := GetControl(GetPrefixe + '_' + fBlobSuffixe);
      if Assigned(C) and (aValueClef1 <> '') or ( (aValueClef1 = '') and (not IsEmpty)) then
      begin
        if (C is THMemo) then
          THMemo(C).Clear
        else if (C is THRichEditOle) then
          THRichEditOle(C).Clear;

        s := '';
        if FieldExists(GetPrefixe + '_' + 'WBMEMO') then
        begin
          if GetBoolean(GetPrefixe + '_' + 'WBMEMO') then
            s := GetMemoValue;
        end
        else
          s := GetMemoValue;

        if s <> '' then
        begin
          if (C is THMemo) then
            THMemo(C).Text := s
          else if (C is THRichEditOle) then
            THRichEditOle(C).Text := s;
        end;

      end;
    end;
  end;
end;

procedure tTOFComm.SetMultiSelection(const Value: Boolean);
begin
  if Assigned(GetControl(GridName)) then
  begin
    {$IFDEF EAGLCLIENT}
      ThGrid(GetControl(GridName)).MultiSelect := Value;
    {$ELSE EAGLCLIENT}
      ThDBGrid(GetControl(GridName)).MultiSelection := Value;
    {$ENDIF  EAGLCLIENT}
  end
end;

procedure tTOFComm.MnProperties_OnClick(Sender: TObject);
var
	Identifiant, Clef: String;
begin
  Identifiant := wGetValueClef1(GetTableName, Ecran, fLequel);
  Identifiant := StringReplace(Identifiant, ';', '~', [rfIgnoreCase, rfReplaceAll]);
  Clef := wMakeFieldString(GetTableName, '~');

  if not IsEmpty() then
    wCallProperties(GetPrefixe, Identifiant, Clef, IntToStr(LongInt(Ecran)));
end;

procedure tTOFComm.PmLoupe_OnPopUp(Sender: TObject);
begin
  if IsEmpty then
    Abort;

  {$IFNDEF PGIMAJVER}
    if (Assigned(Ecran)) and ((Ecran is TFSaisieList) or (Ecran is TFVierge) or (Ecran is TFGRS1)) then   // Fiche de saisie en liste
      wSetMnuLoupe(False, GetLoupeCtx, tPopUpMenu(GetControl('PMLOUPE')))
    else
      //Rend poosible le clic droit contextuel pour les lignes sélectionnées
      //si validation par la qualité retirer argument (MultiSelected) de la fonction qui devient alors inutile
      //wSetMnuLoupe(wMultiSelected(TFMul(Ecran).FListe), GetCtx, tPopUpMenu(GetControl('PMLOUPE')));
      wSetMnuLoupe(False, GetLoupeCtx, tPopUpMenu(GetControl('PMLOUPE')));
  {$ENDIF PGIMAJVER}
  
end;

function tTOFComm.GetLoupeCtx: string;
begin
  Result := '';
end;

procedure tTOFComm.GA_GRILLEDIM_OnChange(Sender: TObject);
  function GetDim: string;
  begin
    Result := wRight(ThValComboBox(sender).Name, 1)
  end;
begin
  if GetControlText('GA_GRILLEDIM' + GetDim) <> '' then
    SetControlProperty('GA_CODEDIM' + GetDim, 'PLUS', ' AND GDI_GRILLEDIM="' + GetControlText('GA_GRILLEDIM' + GetDim)+ '"');
end;

procedure tTOFComm.GA_STATUTART_OnChange(Sender: TObject);
begin
  SetDimensions;
end;

procedure tTOFComm.SetArticle;
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
  if Assigned(GetControl('GA_ARTICLE')) then
    FieldName := 'GA_ARTICLE'
  else
    FieldName := GetPrefixe + '_ARTICLE';

  if Assigned(GetControl('TSDIMENSIONS')) and Assigned(GetControl(FieldName)) and Assigned(GetControl('GA_CODEARTICLE')) and
     Assigned(GetControl('GA_CODEDIM1')) and Assigned(GetControl('GA_CODEDIM2')) and Assigned(GetControl('GA_CODEDIM3')) and
     Assigned(GetControl('GA_CODEDIM4')) and Assigned(GetControl('GA_CODEDIM5')) then
  begin
    SetControlText(FieldName, GetArticle);
  end;
end;

procedure tTOFComm.SetDimensions;
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
        if Assigned(GetControl('TGA_GRILLEDIM')) then
          THLabel(GetControl('TGA_GRILLEDIM')).Visible := Affiche;
        if Assigned(GetControl('GA_GRILLEDIM' + Dim)) then
        with ThValComboBox(GetControl('GA_GRILLEDIM' + Dim)) do
        begin
          Visible := THLabel(GetControl('DIMENSION' + Dim)).Visible;
          Value := iif(Visible, Value, '');
        end;

        { Gestion des valeurs }
        if Assigned(GetControl('TGA_LIBELLEDIM')) then
          THLabel(GetControl('TGA_LIBELLEDIM')).Visible := Affiche;
        if Assigned(GetControl('GA_CODEDIM' + Dim)) then
        with ThValComboBox(GetControl('GA_CODEDIM' + Dim)) do
        begin
          Visible := THLabel(GetControl('DIMENSION' + Dim)).Visible;
          Value := iif(Visible, Value, '');
        end;

        { Gestion des libellés des colonnes }
        if Ecran is TFMul then
        begin
          TFMul(Ecran).SetDisplayLabel('GA_CODEDIM' + Dim, Libelle);
          TFMul(Ecran).SetVisibleColumn('GA_CODEDIM' + Dim, Affiche and (Copy(Libelle, 1, 2) <> '.-'))
        end;
      end;
    end
    else
    begin
      if (Ecran is TFMUL) or (Ecran is TFSTAT) or (Ecran is TFCUBE) then
      begin
        for Dim := '1' to '5' do
        begin
          ThValComboBox(GetControl('GA_GRILLEDIM' + Dim)).Value := '';
          ThValComboBox(GetControl('GA_CODEDIM' + Dim)).Value := '';
          if Ecran is TFMul then
            TFMul(Ecran).SetVisibleColumn('GA_CODEDIM' + Dim, False);
        end;
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

procedure tTOFComm.SetLibelle;
var
  Affiche                                           : Boolean;
  i                                                 : integer;
  MasterPrefixe, FieldName, MasterFieldName, Libelle, NaturePrefixe : string;

  function GetNaturePrefixe: string;
  begin
    if Pos(UpperCase(Copy(MasterPrefixe, 1, 2)), 'QW;QO;QU;QJQK') > 0 then
      Result := 'QMES'
{$IFDEF QUALITE}
    else if UpperCase(Copy(MasterPrefixe, 1, 2)) = 'RQ' then
      Result := 'RQ'
{$ENDIF QUALITE}
    else
      if Copy(MasterPrefixe, 1, 2) = 'GC' then
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

    if Ecran is TFMul then
    begin
      TFMul(Ecran).SetDisplayLabel(FieldName, Libelle);
      TFMul(Ecran).SetVisibleColumn(FieldName, Affiche);
    end;

    { Autre }
    FieldName := GetPrefixe + '_FAMILLENIV' + IntToStr(i);
    SetControlProperty('T' + FieldName, 'Caption', Libelle);
    SetControlVisible(FieldName, Affiche);
    SetControlVisible('T' + FieldName, Affiche);

    if Ecran is TFMul then
    begin
      TFMul(Ecran).SetDisplayLabel(FieldName, Libelle);
      TFMul(Ecran).SetVisibleColumn(FieldName, Affiche);
    end;
  end;

  { Pour les articles, les dépôts et les ressources cela est géré par GCMAJChampLibre() }
  if (GetPrefixe <> 'GA') and (getprefixe <> 'AFF') and (GetPrefixe <> 'GDE') and (GetPrefixe <> 'ARS') and (GetPrefixe <> 'WAP') then
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
        SetControlProperty('T' + FieldName, 'CAPTION', GetRechDom(NaturePrefixe + 'LIBELLELIBRE' + MasterPrefixe, MasterFieldName, false, true));
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
        SetControlProperty('T' + FieldName, 'CAPTION', GetRechDom(NaturePrefixe + 'VALLIBRE' + MasterPrefixe, MasterFieldName, false, true));
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
        SetControlProperty('T' + FieldName, 'CAPTION', GetRechDom(NaturePrefixe + 'DATELIBRE' + MasterPrefixe, MasterFieldName, false, true));
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
        SetControlProperty(FieldName, 'CAPTION', GetRechDom(NaturePrefixe + 'BOOLLIBRE' + MasterPrefixe, MasterFieldName, false, false));
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
        SetControlText    ('T' + FieldName, GetRechDom(NaturePrefixe + 'CHARLIBRE' + MasterPrefixe, MasterFieldName, false, true));
        SetControlVisible ('T' + FieldName, Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
        SetControlVisible (FieldName      , Copy(GetControlText('T' + FieldName), 1, 2) <> '.-');
      end;
    end;
  end;
end;

function tTOFComm.GetRechDom(DataType, Fieldname: string; Abrege,  NotCheckBox: Boolean; Plus: string): string;
var
  s : string;
begin
  s := RechDom(DataType, 'T' + FieldName, Abrege, Plus);
  if NotCheckBox then
    FieldName := 'T' + FieldName;
  if s = 'Error' then
  begin
    if NotCheckBox then
      s := GetControlText(FieldName)
    else
      s := tCheckBox(GetControl(FieldName)).Caption;
  end;
  Result := s;
end;

procedure tTOFComm.BDELETE_OnClick(Sender: TObject);
var
  Rapport: tWRapport;
  bDeleteOk : boolean;
begin
  if isEmpty then
    Abort;

  bDeleteOk:=false;
  if (fMessageToDelete='') or (PgiAsk(fMessageTodelete, Ecran.Caption)=MrYes) then
  begin
    Rapport := TWRapport.Create(TraduireMemoire('Suppression'));
    try
      bDeleteOk := wDeleteTable(fTableName, MakeWhere, true, Rapport);
      Rapport.Display;
    finally
      Rapport.free;
    end;
    RefreshDB('',false);
  end;

  { maj du journal d'évènement }
  if fDeleteEvent <> '' then
    MajJnalEvent(fDeleteEvent, iif(bDeleteOk,'OK','ERR'), Ecran.Caption, '');
end;

procedure tTOFComm.BINSERT_OnClick(Sender: TOBject);
begin
  if fFiche <> '' then
    OuvreFiche(FFiche, '', 'ACTION=CREATION;MONOFICHE', FParamsLanceFiche);
end;

procedure tTOFComm.BSELECTALL_OnClick(Sender: TObject);
var
  i : integer;
  {$IFDEF EAGLCLIENT}
	wBookMark: integer;
  {$ELSE}
	wBookMark: tBookMark;
  {$ENDIF}
begin
  if (ecran is TFSaisieList) and (fSelectAll) then
  begin
    if TFSaisieList(Ecran).FListe.RowCount>1 then
    begin
      for i := 1 to TFSaisieList(Ecran).FListe.RowCount do
      begin
        if TFSaisieList(Ecran).FListe.IsSelected(i) = not TToolbarButton97(GetControl('BSELECTALL')).down then
          TFSaisieList(Ecran).FListe.FlipSelection(i);
      end;
    end;
  end;
  if not (ecran is TFMUL) then
    exit;
  if fSelectAll then
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

procedure tTOFComm.BTUSER_OnClick(Sender: TObject);
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

procedure tTOFComm.BTSHOWBLOCNOTE_OnClick(Sender: tObject);
begin
  if Assigned(GetControl('BTSHOWBLOCNOTE')) then
  begin
    SetControlVisible(GetPrefixe + '_' + fBlobSuffixe, TToolBarButton97(GetControl('BTSHOWBLOCNOTE')).Down);
    THSplitter(GetControl('SPLITPNMEMO')).Visible := TToolBarButton97(GetControl('BTSHOWBLOCNOTE')).Down;
  end;
  SetMemoFrame;
end;

procedure tTOFComm.SetDepot;
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

procedure tTOFComm.SetDomaine;
begin
  if Assigned(GetControl('DOMAINE')) then
  begin
    SetControlText   ('DOMAINE', VH^.ProfilUserC[PrEtablissement].Domaine);
    SetControlEnabled('DOMAINE', not VH^.ProfilUserC[PrEtablissement].ForceDomaine);
  end;
end;

procedure tTOFComm.SetWhereDomaine;
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

procedure tTOFComm.SetEtablissement;
  function GetEtablissement: string;
  begin
    Result := GetPrefixe + '_ETABLISSEMENT';
  end;
begin
  if fSansRestrictEtab then exit;
  if Assigned(GetControl(GetEtablissement)) then
  begin
    SetControlText   (GetEtablissement, VH^.ProfilUserC[PrEtablissement].Etablissement);
    SetControlEnabled(GetEtablissement, not VH^.ProfilUserC[PrEtablissement].ForceEtab);
  end;
end;

{$IFDEF AFFAIRE}

procedure tTOFComm.SetCodeAffaire(stArgument: string);
Var StAff  : String ;
begin
  fQueAffaire := True;
  BChangeTiers := True;
  BVoirStatut := False;
  fCreatAff := True;
  fActionFiche := taCreat;
  fSaisieAffaire := false;

	if Pos('AFFAIREENCREATION', stArgument) > 0 then
		fSaisieAffaire := True;

	// Affectation des champs du code affaire et des boutons
  NomsChampsAffaire(EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, EditAff_, EditAff0_, EditAff1_, EditAff2_, EditAff3_, EditAff4_, EditTiers, EditTiers_);
  bSelectAff1 := TToolBarButton97(GetControl('BSELECTAFF1'));
  bSelectAff2 := TToolBarButton97(GetControl('BSELECTAFF2'));
  bSelectRef := TToolBarButton97(GetControl('BSELECTREF'));

	// Affectation des événements et des infos-bulles
  if (bSelectAff1 <> nil) then
  begin
    bSelectAff1.OnClick := bSelectAff1Click;
		if (V_PGI.EgaliseFourchette = true) then
    	bSelectAff1.OnMouseDown := bSelectAffMouseDown;
    SetControlProperty('BSelectAff1', 'hint', traduitGa('Rechercher client/affaire')); //mcd 26/10/2005
  end;
  if (bSelectAff2 <> nil) then
  begin
    bSelectAff2.OnClick := bSelectAff2Click;
		if (V_PGI.EgaliseFourchette = true) then
    	bSelectAff2.OnMouseDown := bSelectAffMouseDown;
    SetControlProperty('BSelectAff2', 'hint', traduitGa('Rechercher client/affaire')); //mcd 26/10/2005
  end;
  if (bSelectRef <> nil) then
  begin
    bSelectRef.OnClick := bSelectRefClick;
		if (V_PGI.EgaliseFourchette = true) then
    	bSelectRef.OnMouseDown := bSelectAffMouseDown;
        //mcd 05/12/2006
  NomsChampsAffaireRef(EditAffref, EditAffref0, EditAffref1, EditAffref2, EditAffref3, EditAffref4,EditTiersRef);
  if (EditAffRef1 <> nil) then
  begin
    if EditAffref <> nil then
      StAff := EditAffRef.Text
    else
      StAff := '';
    ChargeCleAffaire(EditAffREf0, EditAffref1, EditAffRef2, EditAffRef3, EditAffREf4, nil, fActionFiche, StAff, fSaisieAffaire);
  end;
  if not (GetParamSocSecur('SO_AFGERESSAFFAIRE',False)) and (GetControl('AFFAIREREF1') <> nil) then
  begin   // a refaire dans onargument fct appelante si nom champs différent
    SetcontrolVisible('TAFF_AFFAIREREF', False);
    SetcontrolVisible('AFFAIREREF1', False);
    SetcontrolVisible('AFFAIREREF2', False);
    SetcontrolVisible('AFFAIREREF3', False);
    SetcontrolVisible('AFFAIREREF4', False);
    SetcontrolVisible('AFF_ISAFFAIREREF', False);
    SetcontrolVisible('BSELECTREF', False);
  end;

  end;

	// ActionFiche
  if (EditAff1 <> nil) or (EditAff1_ <> nil) then
    TypeAppelChargeCle(fActionFiche);

	// Formatage du code affaire décomposé
  if (EditAff1 <> nil) then
  begin
    if EditAff <> nil then
      StAff := EditAff.Text
    else
      StAff := '';
    ChargeCleAffaire(EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, nil, fActionFiche, StAff, fSaisieAffaire);
    {$IFDEF GPAO}
    EditAff1.OnExit := EditAFFExit;
    {$ENDIF GPAO}
  end;
  if (EditAff1_ <> nil) then
  begin
    if EditAff_ <> nil then
      StAff := EditAff_.Text
    else
      StAff := '';
    ChargeCleAffaire(EditAff0_, EditAff1_, EditAff2_, EditAff3_, EditAff4_, nil, fActionFiche, StAff, fSaisieAffaire);
    {$IFDEF GPAO}
    EditAff1_.OnExit := EditAFFExit;
    {$ENDIF GPAO}
  end;

  {$IFDEF GPAO}
  if (EditAff2 <> nil) then
    EditAff2.OnExit := EditAFFExit;
  if (EditAff2_ <> nil) then
    EditAff2_.OnExit := EditAFFExit;
  if (EditAff3 <> nil) then
    EditAff3.OnExit := EditAFFExit;
  if (EditAff3_ <> nil) then
    EditAff3_.OnExit := EditAFFExit;
  {$ENDIF GPAO}

  // Gestion de l'effacement des champs de l'affaire et du tiers
  bEffaceAff1 := TToolBarButton97(GetControl('BEFFACEAFF1'));
  if (bEffaceAff1 <> nil) then
  begin
    SetControlProperty('BeffaceAff1', 'hint', traduitGa('Effacer critère client/affaire')); //mcd 26/10/2005
    bEffaceAff1.OnClick := bEffaceAff1Click;
  end;

  bEffaceAff2 := TToolBarButton97(GetControl('BEFFACEAFF2'));
  if (bEffaceAff2 <> nil) then
  begin
    SetControlProperty('BeffaceAff2', 'hint', traduitGa('Effacer critère client/affaire')); //mcd 26/10/2005
    bEffaceAff2.OnClick := bEffaceAff2Click;
  end;

  // Cacher les boutons si pas en mode affaire
  if not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte) then
  begin
    if (bSelectAff1 <> nil) then
      SetControlVisible('BSELECTAFF1', False);
    if (bSelectAff2 <> nil) then
      SetControlVisible('BSELECTAFF2', False);
    if (bEffaceAff1 <> nil) then
      SetControlVisible('BEFFACEAFF1', False);
    if (bEffaceAff2 <> nil) then
      SetControlVisible('BEFFACEAFF2', False);
  end;

end;

procedure tTOFComm.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  // Le champ AFFAIRE est obligatoire dans la fiche
  // Il sera renseigné par la fct appelante si accès direct à une affaire
  // Ce code ne doit autrement jamais être renseigné par le prog
  Aff := THEdit(GetControl('ACT_AFFAIRE'));

  Aff0 := THEdit(GetControl('ACT_AFFAIRE0'));
  Aff1 := THEdit(GetControl('ACT_AFFAIRE1'));
  Aff2 := THEdit(GetControl('ACT_AFFAIRE2'));
  Aff3 := THEdit(GetControl('ACT_AFFAIRE3'));
  Aff4 := THEdit(GetControl('ACT_AVENANT'));
  Aff_ := THEdit(GetControl('ACT_AFFAIRE_'));
  Aff0_ := THEdit(GetControl('ACT_AFFAIRE0_'));
  Aff1_ := THEdit(GetControl('ACT_AFFAIRE1_'));
  Aff2_ := THEdit(GetControl('ACT_AFFAIRE2_'));
  Aff3_ := THEdit(GetControl('ACT_AFFAIRE3_'));
  Aff4_ := THEdit(GetControl('ACT_AVENANT_'));
  Tiers := THEdit(GetControl('ACT_TIERS'));
  Tiers_ := THEdit(GetControl('ACT_TIERS_'));
end;

procedure tTOFComm.NomsChampsAffaireREf(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4,tiers: THEdit);
begin    //mcd 05/12/2006 si le champs affaireref1 existe sur le mul, on le stock
    //si nom de champ différent, surcharger cette fct
  Aff := THEdit(GetControl('AFF_AFFAIREREF'));
  Aff1 := THEdit(GetControl('AFFAIREREF1'));
  Aff2 := THEdit(GetControl('AFFAIREREF2'));
  Aff3 := THEdit(GetControl('AFFAIREREF3'));
  Aff4 := THEdit(GetControl('AFFAIREREF4'));
  Tiers := THEdit(GetControl('aff_tiersref'));
end;

procedure tTOFComm.bEffaceRef(Sender: TObject);
begin
  EffaceAffaire(EditTiersref, EditAffref, EditAffref0, EditAffref1, EditAffref2, EditAffref3, EditAffref4);
end;

procedure tTOFComm.bSelectAff1Click(Sender: TObject);
begin
  SelectionAffaire(EditTiers , EditAff , EditAff0 , EditAff1 , EditAff2 , EditAff3 , EditAff4 , BVoirStatut, false, '', false, true, false, fQueAffaire, fProposition, BEcole, False, False, fTypeAffaire);
end;

procedure tTOFComm.bSelectAff2Click(Sender: TObject);
begin
  SelectionAffaire(EditTiers_, EditAff_, EditAff0_, EditAff1_, EditAff2_, EditAff3_, EditAff4_, BVoirStatut, false, '', false, true, false, fQueAffaire, fProposition, BEcole, False, False, fTypeAffaire);
end;

procedure tTOFComm.bSelectRefClick(Sender: TObject);
var
  IsAffaireRef: boolean;
begin //mcd 30/09/2005 pour n'afficher que affaire ref
  IsAffaireRef := true;
  SelectionAffaire(EditTiersRef, EditAffref, EditAffref0, EditAffref1, EditAffref2, EditAffref3, EditAffref4,
    BVoirStatut, false, '', false, true, false, fQueAffaire, fProposition, BEcole, IsAffaireRef);
//mcd 05/12/06  SelectionAffaire(EditTiers_, EditAff_, EditAff0_, EditAff1_, EditAff2_, EditAff3_, EditAff4_,
//    BVoirStatut, false, '', false, true, false, fQueAffaire, fProposition, BEcole, IsAffaireRef);
end;

procedure tTOFComm.SelectionAffaire(EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4: THEdit;
  bChangeStatut: boolean = false; bModele: boolean = false; sContexte: string = '';
  bReduit: boolean = false; bOuvreSiUn: boolean = true; bItDatesMax: boolean = false;
  fQueAffaire: boolean = false; fProposition: boolean = false;
  fEcole: boolean = false; IsAffaireRef: boolean = false ; AvecPc: boolean = false; fTypeAffaire:string='');
var
  //GP_DS_GP14214_20071011
  Aff0, Aff1, Aff2, Aff3, Aff4, Affaire, Tiers: string;

begin
  Aff0 := '';
  Aff1 := '';
  Aff2 := '';
  Aff3 := '';
  Aff4 := '';
  Affaire := '';
  // Pour ne pas passer  le tiers à la recherche affaire dans un contexte Achat (Ex: Tarifs d'achat)
  if FtypeTiers='FOU' then
  begin
    Tiers:= EditTiers.Text;
    EditTiers.Text:='';
  end;

  if fTypeAffaire='' then fTypeAffaire:= 'NOR'; // Par défaut affaires dervices

  if (EditAff0 <> nil) then
    if EditAff0.Text = 'P' then
      fProposition := true;

  if EditAff <> nil then Affaire := EditAff.Text;

    // PL le 04/05/07 : FQ12280, on passe la ressource si nécessaire pour préinitialiser la recherche
  GetAffaireEnteteSt(editaff0,EditAff1, EditAff2, EditAff3, EditAff4, EditTiers, Affaire, fProposition, bChangeStatut, bModele, BChangeTiers,
    fCreatAff, sContexte, bReduit, bOuvreSiUn, bItDatesMax);

  CodeAffaireDecoupe(Affaire, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);

  if EditAff <> nil then
    EditAff.Text := Affaire;
  if EditAff0 <> nil then
    EditAff0.Text := Aff0;
  if EditAff1 <> nil then
    EditAff1.Text := Aff1;

  // mcd 01/10/01 ajout test sur zones visibles, sinon sélection 'de à' ne fonctionne pas !!!
  // mcd 17/05/02 ajout effacement zone si non géré, sinon, laisse ancienne valeur (bug complligne si change mission)
  if EditAff2 <> nil then
  begin
    if GetParamSocSecur('SO_AFFCO2VISIBLE', False) then
      EditAff2.Text := Aff2
    else
      EditAff2.text := '';
  end;
  if (EditAff3 <> nil) then
  begin
    if GetParamSocSecur('SO_AFFCO3VISIBLE', False) then
      EditAff3.Text := Aff3
    else
    begin //mcd 30/06/2005 si pour affaire ref, ne pas effacer, sinon pas OK GI
      if not IsAffaireRef then
        EditAff3.text := ''
      else
        EditAff3.Text := Aff3;
    end;
  end;
  if (EditAff4 <> nil) and (GetParamSocSecur('SO_AFFGESTIONAVENANT', False)) then
    EditAff4.Text := Aff4;
  //GP_DS_GP14214_20071011

  // On renseigne le tiers pour les recherches affaires issues des ventes uniquement
  if (fTypeTiers='FOU') then
    EditTiers.text:= Tiers;

  if (Ecran is TFMul) then
    TFMul(Ecran).ChercheClick;
end;

procedure tTOFComm.bEffaceAff1Click(Sender: TObject);
var
  bFocus: boolean;
begin
  bFocus := false;
  // PL le 15/01/02 pour V575
  if (EditAff1_ <> nil) then
    EffaceAffaire(EditTiers_, EditAff_, EditAff0_, EditAff1_, EditAff2_, EditAff3_, EditAff4_);

  EffaceAffaire(EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4);

  // gm 29/12/06  pas de refresh sue Gomme  FQ 13121 le 11/01/06
//  if (Ecran is TFMul) then
//    TFMul(Ecran).ChercheClick;



  if (Editaff0 <> nil) then
    if EditAff0.CanFocus then
    begin
      EditAff0.SetFocus;
      bFocus := true;
    end;

  if not (bFocus) then
    if EditAff1.CanFocus then
      EditAff1.SetFocus;
end;

procedure tTOFComm.bEffaceAff2Click(Sender: TObject);
var
  bFocus: boolean;
begin
  bFocus := false;
  EffaceAffaire(EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4);
  EffaceAffaire(EditTiers_, EditAff_, EditAff0_, EditAff1_, EditAff2_, EditAff3_, EditAff4_);

  // gm 29/12/06  pas de refresh sur Gomme   FQ13121
//  if (Ecran is TFMul) then
//    TFMul(Ecran).ChercheClick;



  if (EditAff0_ <> nil) then
    if EditAff0_.CanFocus then
    begin
      EditAff0_.SetFocus;
      bFocus := true;
    end;

  if not (bFocus) then
    EditAff1_.SetFocus;
end;

procedure tTOFComm.EffaceAffaire(EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4: THEdit);
begin
  if (EditTiers <> nil) then
  begin
    if (GetKeyState(VK_CONTROL) >= 0) then
      // PL le 31/05/05 : si on appuit sur la touche controle en même temps
      // on efface seulement les codes affaire, pas le tiers
    begin
      if (EditTiers.Enabled = true) and (fTypeTiers<>'FOU') then
        EditTiers.Text := ''
      else
      begin
        // PL le 18/05/04 : le tiers n'était pas toujours remis à vide :
        // Soit le tiers est tout seul à être protégé et on n'y touche pas,
        // soit tous les champs affaires sont protégés et alors il faut tout effacer
        if (EditAff1.Enabled = false) and (fTypeTiers<>'FOU') then
          EditTiers.Text := '';
      end;
    end;
  end;

  if EditAff4<>nil then
    EditAff4.Text := '';
  if EditAff3<>nil then
    EditAff3.Text := '';
  if EditAff2<>nil then
    EditAff2.Text := '';
  if EditAff1<>nil then
    EditAff1.Text := '';
{$IFNDEF BTP}
  if (Editaff0 <> nil) then
  begin
    EditAff0.Text := '';
  end;
{$ENDIF}
  if (EditAff <> nil) then
    EditAff.Text := '';
end;

procedure tTOFComm.TypeAppelChargeCle(var AvActionFiche: TActionFiche);
begin
  AvActionFiche := taCreat;
end;

// fonction de copie automatique des champs du code affaire "de" vers ceux du "à"
// et inversement
procedure tTOFComm.bSelectAffMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  Tiers, Aff0, Aff1, Aff2, Aff3, Aff4, Affaire: string;
  ChampDestTiers, ChampDestAff, ChampDestAff0, ChampDestAff1, ChampDestAff2, ChampDestAff3, ChampDestAff4: THEdit;

begin
  if (Shift <> [ssCtrl, ssLeft]) then
    exit;

  Aff0 := '';
  Aff1 := '';
  Aff2 := '';
  Aff3 := '';
  Aff4 := '';
  Affaire := '';

  if (TToolBarButton97(Sender).Name = 'BSELECTAFF1') then
  begin
    ChampDestTiers := EditTiers;
    ChampDestAff := EditAff;
    ChampDestAff0 := EditAff0;
    ChampDestAff1 := EditAff1;
    ChampDestAff2 := EditAff2;
    ChampDestAff3 := EditAff3;
    ChampDestAff4 := EditAff4;
    if (EditTiers_ <> nil) then
      Tiers := EditTiers_.Text;
    if (EditAff0_ <> nil) then
      Aff0 := EditAff0_.Text;
    if (EditAff1_ <> nil) then
      Aff1 := EditAff1_.Text;
    if (EditAff2_ <> nil) then
      Aff2 := EditAff2_.Text;
    if (EditAff3_ <> nil) then
      Aff3 := EditAff3_.Text;
    if (EditAff4_ <> nil) then
      Aff4 := EditAff4_.Text;
    if (EditAff_ <> nil) then
      Affaire := EditAff_.Text;
  end
  else
    if (TToolBarButton97(Sender).Name = 'BSELECTAFF2') then
  begin
    ChampDestTiers := EditTiers_;
    ChampDestAff := EditAff_;
    ChampDestAff0 := EditAff0_;
    ChampDestAff1 := EditAff1_;
    ChampDestAff2 := EditAff2_;
    ChampDestAff3 := EditAff3_;
    ChampDestAff4 := EditAff4_;
    if (EditTiers <> nil) then
      Tiers := EditTiers.Text;
    if (EditAff0 <> nil) then
      Aff0 := EditAff0.Text;
    if (EditAff1 <> nil) then
      Aff1 := EditAff1.Text;
    if (EditAff2 <> nil) then
      Aff2 := EditAff2.Text;
    if (EditAff3 <> nil) then
      Aff3 := EditAff3.Text;
    if (EditAff4 <> nil) then
      Aff4 := EditAff4.Text;
    if (EditAff <> nil) then
      Affaire := EditAff.Text;
  end
  else
    exit;

  if (ChampDestTiers <> nil) then
    ChampDestTiers.Text := Tiers;
  if (ChampDestAff <> nil) then
    ChampDestAff.Text := Affaire;
  if (ChampDestAff0 <> nil) then
    ChampDestAff0.Text := Aff0;
  if (ChampDestAff1 <> nil) then
    ChampDestAff1.Text := Aff1;
  if (ChampDestAff2 <> nil) and GetParamSocSecur('SO_AFFCO2VISIBLE', False) then
    ChampDestAff2.Text := Aff2;
  if (ChampDestAff3 <> nil) and GetParamSocSecur('SO_AFFCO3VISIBLE', False) then
    ChampDestAff3.Text := Aff3;
  if (ChampDestAff4 <> nil) and GetParamSocSecur('SO_AFFGESTIONAVENANT', False) then
    ChampDestAff4.Text := Aff4;

  Abort;
end;
{$ENDIF AFFAIRE}

function tTOFComm.FieldExists(const FieldName: string): boolean;
begin
  {$IFNDEF EAGLCLIENT}
  Result := Assigned(GetDataSet.FindField(FieldName))
  {$ELSE !EAGLCLIENT}
  Result := GetDataSet.FieldExists(FieldName)
  {$ENDIF !EAGLCLIENT}
end;

function TTOFComm.EcranIs(const FormName: String): Boolean;
begin
  Result := Pos(UpperCase(Trim(FormName)), UpperCase(Ecran.Name)) = 1
end;

{$IFDEF AFFAIRE}
{$IFDEF GPAO}
procedure tTOFComm.EditAFFExit(Sender: TOBject);
var
  Aff0, Aff1, Aff2, Aff3, Aff4, Affaire: string;
begin
  Aff0 := '';
  Aff1 := '';
  Aff2 := '';
  Aff3 := '';
  Aff4 := '';
  Affaire := '';
  if EditAff <> nil then
  begin
    if EditAff0<>nil then
      Aff0:= EditAff0.Text
    else
      Aff0:= Copy(EditAFF.Text,1,1);
    if EditAff1<>nil then
      Aff1:= EditAff1.Text;
    if EditAff2<>nil then
      Aff2:= EditAff2.Text;
    if EditAff3<>nil then
      Aff3:= EditAff3.text;
    if EditAFF4<>nil then
      Aff4:= EditAFF4.text;

    EditAFF.Text:='';

    if (Aff1<>'') or (Aff2<>'') or (Aff3<>'') or (Aff4<>'') then
    //GP_DS_GP14214_20071011
    begin
      if Aff4 = '' then
        Aff4:= 'W';
      EditAFF.Text:= CodeAffaireRegroupe(Aff0, Aff1, Aff2, Aff3, Aff4, taModif, false, false, false);
    end;
  end;
  if EditAff_ <> nil then
  begin
    if EditAff0_<>nil then
      Aff0:= EditAff0_.Text
    else
      Aff0:= Copy(EditAFF.Text,1,1);
    if EditAff1_<>nil then
      Aff1:= EditAff1_.Text;
    if EditAff2_<>nil then
      Aff2:= EditAff2_.Text;
    if EditAff3_<>nil then
      Aff3:= EditAff3_.text;
    if EditAFF4_<>nil then
      Aff4:= EditAFF4_.text;

    EditAFF_.Text:='';

    if (Aff1<>'') or (Aff2<>'') or (Aff3<>'') or (Aff4<>'') then
      EditAFF_.Text:= CodeAffaireRegroupe(Aff0, Aff1, Aff2, Aff3, Aff4, taModif, false, false, false);
  end;
end;
{$ENDIF GPAO}
{$ENDIF AFFAIRE}


procedure tTOFComm.BchercheClick(Sender: Tobject);
begin
  RefreshDB;
end;

Initialization
	registerclasses([tTOFComm]);
end.
