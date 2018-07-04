unit TobUtil;

interface

uses
  Math,
  SysUtils,
  hMsgBox,
  HEnt1,
  Classes,
  variants,
  {$IFNDEF EAGL}
  EBizUtil,
  {$ENDIF EAGL}
  UTob,
  wCommuns;


type
  TVerifChamp = ( vcAucun, vcSupp, vcTous );

  TChampRech = record
    Nom: string;
    Sens: Char;
    Indice: Integer;
    Valeur: Variant;
  end;
  PChampRech = ^TChampRech;
  // description de TOB
  TChampsTob = class
  private
    FnomTOB : string;
    FVerifChamp: TVerifChamp;
    FTob: TOB;
    FListChamps: TStringList;
    procedure Init( pTob: TOB );
    procedure AddTousLesChamps;
    procedure AddlesChamps(const NomChamps: array of string);
    procedure InclusLesChampsSUp(pTOB: TOB);
    procedure DelChampsSUp;
    procedure ReinitChamps(pTOB: TOB);
  public
    constructor Create( pTob: TOB ); overload;
    constructor Create( pTob: TOB; const NomChamps: array of string ); overload;
    destructor Destroy; override;
    function FieldExists (sNomChamp : string) : boolean;
    function GetIndice( pTOB : TOB; const sNomChamp: string ): Integer;
    function GetValue( pTob: TOB; const sNomChamp: string ): Variant;
    function GetValeur( pTob: TOB; const sNomChamp: string ): Variant;
    function GetString( pTob: TOB; const sNomChamp: string ): string;
    function GetDouble( pTob: TOB; const sNomChamp: string ): Double;
    function GetInteger( pTob: TOB; const sNomChamp: string ): Integer;
    function GetBoolean( pTob: TOB; const sNomChamp: string ): Boolean;
    function GetDateTime( pTob: TOB; const sNomChamp: string ): TDateTime;
    procedure PutValue (pTOB : TOB; const sNomChamp : string; sValeur : variant);
    property VerifChamp: TVerifChamp read FVerifChamp write FVerifChamp;
    procedure AddChampSup(pTOB: TOB);
  end;


  // liste des descriptifs de TOB

  TlisteChampsTOB = class (Tlist)
  private
    function GetItems(Indice: integer): TChampsTob;
    procedure SetItems(Indice: integer; const Value: TChampsTob);
    function Add(AObject: TChampsTob): Integer;
  public
    property Items [Indice : integer] : TChampsTob read GetItems write SetItems;
    constructor create;
    destructor destroy; override;
    //
    procedure Libere;
    procedure Initialise;
    //

    function findEqual (pTOB : TOB) : TchampsTOB;
    function FindWhere(const S: TChampsTob; var Index: Integer): Boolean;
    //
    function FieldExists (pTOB : TOB; const SNomChamp : string) : boolean;
    function GetIndice( pTOB : TOB; const sNomChamp: string ): Integer;
    function GetNumChamp( pTOB : TOB; const sNomChamp: string ): Integer;
    function GetValue( pTob: TOB; const sNomChamp: string ): Variant;
    function GetValeur( pTob: TOB; const sNomChamp: string ): Variant;
    function GetString( pTob: TOB; const sNomChamp: string ): string;
    function GetDouble( pTob: TOB; const sNomChamp: string ): Double;
    function GetInteger( pTob: TOB; const sNomChamp: string ): Integer;
    function GetBoolean( pTob: TOB; const sNomChamp: string ): Boolean;
    function GetDateTime( pTob: TOB; const sNomChamp: string ): TDateTime;
    procedure PutValue (pTOB : TOB; const sNomChamp : string; sValeur : variant);
    procedure AddChampSup(pTOB: TOB; const SnomChamp: string;const toutelesfilles: boolean=false);
  end;

function ChRech( const sNom: string; vVal: Variant; Sens: Char = '+' ): TChampRech;
function ChTri( const sNom: string; Sens: Char = '+' ): TChampRech;

function DEChampTipeToVarType( const sTipe: string ): Integer;
procedure IndicesMinMaxChampsTob( pTob: Tob; var iMin, iMax: Integer; bTobVirtuelle: Boolean = False );
function FOSauveTobBin( sNomFichier: TFileName; pTob: Tob ): boolean;
procedure CopieChampsTob( const TOBSrc: TOB; TobDest: TOB; RemplaceChampsNonVides: boolean );
function TOBRechercheOptimise( const TobSrc: TOB; const Field, Valeur: string ): TOB; overload;
function TOBRechercheOptimise( const TobSrc: TOB; Champs: array of TChampRech ): TOB; overload;
function TOBRechercheOptimise( const TobSrc: TOB; Champs: array of TChampRech; var iIndice: Integer ): TOB; overload;
function TOBRechercheOuCree( const TobSrc: TOB; Champs: array of TChampRech;
                             const sNomTable: string; var bCreee: Boolean ): TOB;
function TOBTriOptimise( const TOBSrc: TOB; Champs: array of TChampRech; bVideSrc: Boolean ): TOB; overload;
procedure TOBTriOptimise( const TOBSrc: TOB; Champs: array of TChampRech ); overload;

procedure TOBRenameChampSup( pTob: TOB; OldName, NewName: string );
procedure TOBRenameChampSupAllFille( TobMere: TOB; OldName, NewName: string );
function ModeTOBtoXMLString( Entree: TOB; Supp: boolean = false; Zipped : boolean = false ) : string;
function ModeTOBtoBINString( Entree: TOB; Supp: boolean = false; Zipped : boolean = false ) : string;

function TOBFindWhere(TOBF: TOB; TheChamp : string; ATrouver : variant; var Index: Integer): Boolean;

//procedure UpdateTobStruc( aTob: Tob; const pChamps: PDEChamps );
var TheOptiTOB : TlisteChampsTOB;

implementation
{$IFDEF GCGC}
uses
//  UViewTob,
  UtilPGI,
  EntGC;
{$ENDIF GCGC}


procedure InitTOBUtil;
begin
  TheOptiTOB := TlisteChampsTOB.create;
end;

procedure FinTOBUtil;
begin
  TheOptiTOB.free;
end;


{ TChampsTob }

procedure TChampsTob.AddChampSup (pTOB : TOB);
begin
  ReinitChamps (pTOB);
end;

{***********UNITE*************************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 08/02/2006
Modifié le ... : 15/02/2006
Description .. : Objet qui maintient un cache des indices de quelques
Suite ........ : champs d'une tob utilisés intensivement (typiquement dans
Suite ........ : une boucle) pour réduire leur temps d'accès lorsque
Suite ........ : beaucoup d'autres champs sont présents dans la tob.
Suite ........ :
Suite ........ : Par défaut, bCacheChampsSup est à False : les champs
Suite ........ : supplémentaires qui n'ont pas été initialisés dans le
Suite ........ : constructeur ne sont pas mis en cache lors d'un appel à
Suite ........ : GetIndice
Mots clefs ... :
*****************************************************************}

constructor TChampsTob.Create( pTob: TOB );
begin
  inherited Create();
  Init( pTob );
  AddTousLesChamps;
end;

{***********UNITE*************************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 08/02/2006
Modifié le ... : 15/02/2006
Description .. : Variante du constructeur qui permet d'initialiser directement
Suite ........ : les indices d'un tableau champs.
Mots clefs ... :
*****************************************************************}
constructor TChampsTob.Create( pTob: TOB; const NomChamps: array of string );
begin
  inherited Create();
  Init( pTob );
  AddlesChamps (NomChamps);
end;

procedure TChampsTob.ReinitChamps (pTOB : TOB);
begin
  FListChamps.Clear;
  InclusLesChampsSUp (pTOB);
  AddTousLesChamps;
end;

procedure TChampsTob.Init( pTob: TOB );
begin
  FTob := TOB.Create(pTOB.nomtable, nil, -1 );
  InclusLesChampsSUp (pTOB);
  //
  FNomTOB := pTOB.NomTable;
  FListChamps := TStringList.Create;
  FListChamps.Sorted := True;           // active la recherche dichotomique
  FVerifChamp := vcTous;
end;

procedure TChampsTob.InclusLesChampsSUp (pTOB : TOB);
var i : integer;
    NomChamp : string;
begin
  DelChampsSUp;
  for i := 1000 to (1000 + pTob.ChampsSup.Count-1) do
  begin
    NomChamp := pTOB.GetNomChamp(i);
    fTOB.AddChampSup (Nomchamp,false);
  end;
end;

procedure TChampsTob.DelChampsSUp;
var i : integer;
    NomChamp : string;
begin
  for i := 1000 to (1000 + fTob.ChampsSup.Count-1) do
  begin
    NomChamp := fTOB.GetNomChamp(i);
    fTOB.DelChampSup (Nomchamp,false);
  end;
end;

procedure TChampsTob.AddlesChamps (const NomChamps: array of string );
var i : integer;
begin
  for i := Low( NomChamps ) to High( NomChamps ) do
    FListChamps.AddObject( NomChamps[i], Pointer(fTob.GetNumChamp(NomChamps[i])) );
end;

procedure TChampsTob.AddTousLesChamps ;
var I : integer;
    NomChamp : string;
begin
  for i := 1 to fTOB.NbChamps do
  begin
    NomChamp := fTOB.GetNomChamp (i);
    FListChamps.AddObject( NomChamp, Pointer(fTOB.GetNumChamp(NomChamp)) );
  end;
  for i := 1000 to (1000 + fTob.ChampsSup.Count-1) do
  begin
    NomChamp := fTOB.GetNomChamp(i);
    FListChamps.AddObject( NomChamp, Pointer(fTOB.GetNumChamp(NomChamp)) );
  end;
end;

destructor TChampsTob.Destroy;
begin
//  FTob.Free;
  FListChamps.Free;
  inherited;
end;

function TChampsTob.FieldExists (sNomChamp : string) : boolean;
begin
  result := (FListChamps.IndexOf( sNomChamp ) >= 0);
end;

{***********UNITE*************************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 08/02/2006
Modifié le ... : 16/02/2006
Description .. : Retourne l'indice du champ demandé.
Suite ........ :
Suite ........ : Si il n'est pas en cache, il est récupéré à partir de la copie
Suite ........ : de la tob utilisée dans le constructeur et rajouté au cache.
Mots clefs ... :
*****************************************************************}
function TChampsTob.GetIndice( pTOB : TOB; const sNomChamp: string ): Integer;
var
  iIndiceChamp: Integer;
begin
  iIndiceChamp := FListChamps.IndexOf( sNomChamp );
  if( iIndiceChamp < 0 ) then
  begin
    // dans le cas ou le champs n'est pas présent dans la liste
    // on reprend la structure de la tob
    ReinitChamps (pTOB);
    iIndiceChamp := FListChamps.IndexOf( sNomChamp );
  end;

  if( iIndiceChamp >= 0 ) then
    Result := Integer( FListChamps.Objects[iIndiceChamp] )
  else
  begin
    result := -1;
  end;
end;

function TChampsTob.GetValue(pTob: TOB; const sNomChamp: string): Variant;
var
  iIndice: Integer;
begin
  Result := varEmpty;
  if( Assigned(pTob) ) then
  begin
    iIndice := GetIndice( pTOB, sNomChamp );
    if iIndice >= 0 then Result := pTob.GetValeur( iIndice );
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 16/02/2006
Modifié le ... :   /  /
Description .. : Retourne la valeur de la tob en utilisant l'indice mis en
Suite ........ : cache et en vérifiant qu'il correspond bien au champ
Suite ........ : demandé.
Mots clefs ... :
*****************************************************************}
function TChampsTob.GetValeur( pTob: TOB; const sNomChamp: string ): Variant;
var
  iIndice: Integer;
begin
  Result := varEmpty;
  if( Assigned(pTob) ) then
  begin
    iIndice := GetIndice( pTOB, sNomChamp );
(*
    if( (VerifChamp = vcAucun) or ((VerifChamp = vcSupp) and (iIndice < 1000)) or
        SameText(sNomChamp, pTob.GetNomChamp(iIndice)) ) then
      Result := pTob.GetValeur( iIndice )
    else
      Result := pTob.GetValue( sNomChamp );
*)
    if iIndice >= 0 then Result := pTob.GetValeur( iIndice )
                    else PgiBox ('le champ '+sNomchamp+' n''existe pas dans la TOB '+pTOB.NomTable);
  end;
end;

function TChampsTob.GetDouble( pTob: TOB; const sNomChamp: string ): Double;
var
  val: Variant;
begin
  try
    val := GetValeur( pTob, sNomChamp );
    Result := val;
  except
    on E: Exception do
    begin
      E.Message := E.Message + #13#10
                 + Format( 'La valeur retournée (%s) pour le champ "%s" de "%s" n''est pas un double',
                           [VarToStr(val), sNomChamp, pTob.NomTable] );
      raise;
    end;
  end;
end;

function TChampsTob.GetInteger( pTob: TOB; const sNomChamp: string ): Integer;
var
  val: Variant;
begin
  try
    val := GetValeur( pTob, sNomChamp );
    Result := val;
  except
    on E: Exception do
    begin
      E.Message := E.Message + #13#10
                 + Format( 'La valeur retournée (%s) pour le champ "%s" de "%s" n''est pas un entier',
                           [VarToStr(val), sNomChamp, pTob.NomTable] );
      raise;
    end;
  end;
end;

function TChampsTob.GetString( pTob: TOB; const sNomChamp: string ): string;
var
  val: Variant;
begin
  try
    val := GetValeur( pTob, sNomChamp );
    Result := val;
  except
    on E: Exception do
    begin
      E.Message := E.Message + #13#10
                 + Format( 'La valeur retournée (%s) pour le champ "%s" de "%s" n''est pas une chaîne',
                           [VarToStr(val), sNomChamp, pTob.NomTable] );
      raise;
    end;
  end;
end;

function TChampsTob.GetBoolean(pTob: TOB; const sNomChamp: string): Boolean;
var
  val: Variant;
begin
  try
    val := GetValeur( pTob, sNomChamp );
    Result := val;
  except
    on E: Exception do
    begin
      E.Message := E.Message + #13#10
                 + Format( 'La valeur retournée (%s) pour le champ "%s" de "%s" n''est pas un booléen',
                           [VarToStr(val), sNomChamp, pTob.NomTable] );
      raise;
    end;
  end;
end;

function TChampsTob.GetDateTime(pTob: TOB; const sNomChamp: string): TDateTime;
var
  val: Variant;
begin
  try
    val := GetValeur( pTob, sNomChamp );
    Result := val;
  except
    on E: Exception do
    begin
      E.Message := E.Message + #13#10
                 + Format( 'La valeur retournée (%s) pour le champ "%s" de "%s" n''est pas une date',
                           [VarToStr(val), sNomChamp, pTob.NomTable] );
      raise;
    end;
  end;
end;

procedure TChampsTob.PutValue(pTOB: TOB; const sNomChamp: string;sValeur: variant);
var Iindice : integer;
begin
  if( Assigned(pTob) ) then
  begin
    iIndice := GetIndice( pTOB, sNomChamp );
    if iIndice >= 0 then pTob.PutValeur( iIndice,sValeur);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 27/01/2005
Modifié le ... : 27/01/2005
Description .. : Retourne le type de variant correspondant au
Suite ........ : DEChamp[][].Tipe
Mots clefs ... :
*****************************************************************}

function DEChampTipeToVarType( const sTipe: string ): Integer;
begin
  if SameText( sTipe, 'BOOLEAN' ) then
    Result := varBoolean
  else if SameText( sTipe, 'SMALLINT' ) or SameText( sTipe, 'INTEGER' ) then
    Result := varInteger
  else if SameText( sTipe, 'DOUBLE' ) or SameText( sTipe, 'EXTENDED' ) or SameText( sTipe, 'RATE' ) then
    Result := varDouble
  else if SameText( sTipe, 'DATE' ) then
    Result := varDate
  else
    Result := varString;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 15/09/2004
Modifié le ... : 15/09/2004
Description .. : Retourne les indices min et max des champs d'une tob.
Suite ........ : Si la Tob est virtuelle, retourne les indices correspondant
Suite ........ : aux champs supplémentaires
Mots clefs ... :
*****************************************************************}

procedure IndicesMinMaxChampsTob( pTob: Tob; var iMin, iMax: Integer; bTobVirtuelle: Boolean = False );
const
  OffsetChampsSup = 999;
begin
  iMin := 1;
  if ( pTob is TOB ) then
  begin
    iMax := pTob.NbChamps;
{$IFNDEF BTP}
{$IFDEF GCGC}
    if VH_GC.ModeDeconnecte then
      bTobVirtuelle := True;
{$ENDIF GCGC}
{$ENDIF BTP}
    if ( iMax = 0 ) and bTobVirtuelle then
    begin
      Inc( iMin, OffsetChampsSup );
      iMax := OffsetChampsSup + pTob.ChampsSup.Count;
    end;
  end
  else
    iMax := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 10/15/2004
Modifié le ... :   /  /
Description .. : Sauvegarde d'une Tob au format binaire en préservant le
Suite ........ : fichier précédent si la fonction échoue
Mots clefs ... :
*****************************************************************}
function FOSauveTobBin( sNomFichier: TFileName; pTob: Tob ): boolean;
var
  tempFileName: TFileName;
begin
  Result := False;
  try
    if ( pTob is TOB ) then
    try
      tempFileName := ChangeFileExt( sNomFichier, '.$$$' );
      Result := pTob.SaveToBinFile( tempFileName, False, True, True, True );
    finally
      if Result then
      begin
        DeleteFile( sNomFichier );
        RenameFile( tempFileName, sNomFichier );
      end
      else
        DeleteFile( tempFileName );
    end;
  except
{$IFNDEF BTP}
{$IFDEF GCGC}
    on E: Exception do
      MODEAfficheErreurSQL( TraduireMemoire('Exception :') + ' ' + E.Message );
{$ENDIF GCGC}
{$ENDIF BTP}
  end;
end;

function ValeurNulle( Valeur: Variant ): Boolean;
var
  iVarType: Integer;
begin
  Result := VarIsEmpty( Valeur ) or VarIsNull( Valeur );
  if ( not ( Result ) ) then
  begin
    iVarType := VarType( Valeur );
    case iVarType of
      varByte,
        varSmallint,
        varInteger: Result := ( Valeur = 0 );
      varSingle,
        varDouble,
        varCurrency: Result := ( Abs( Valeur ) < 1E-6 );
      varDate: Result := Abs( Valeur - IDate1900 ) < 1E-6;
      varOleStr,
        varStrArg,
        varString: Result := ( Valeur = '' );
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 29/10/2004
Modifié le ... : 29/10/2004
Description .. : Copie les champs de TOBSrc vers TobDest
Mots clefs ... :
*****************************************************************}

procedure CopieChampsTob( const TOBSrc: TOB; TobDest: TOB; RemplaceChampsNonVides: boolean );
var
  idxSrc, idxDest, IndMin, IndMax: Integer;
  NomChamp: string;
begin
  if ( TOBSrc is TOB ) and ( TobDest is TOB ) then
  begin
    IndicesMinMaxChampsTob( TOBSrc, IndMin, IndMax );
    for idxSrc := IndMin to IndMax do
    begin
      NomChamp := TOBSrc.GetNomChamp( idxSrc );
      idxDest := TobDest.GetNumChamp( NomChamp );
      if ( idxDest <= 0 ) then
        TobDest.AddChampSupValeur( NomChamp, TOBSrc.GetValeur( idxSrc ) )
      else if RemplaceChampsNonVides or ValeurNulle( TobDest.GetValeur( idxDest ) ) then
        TobDest.PutValeur( idxDest, TOBSrc.GetValeur( idxSrc ) );
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 29/11/2005
Modifié le ... :   /  /
Description .. : Initialise un champ comme critère de recherche pour la
Suite ........ : fonction TOBRechercheOptimise
Mots clefs ... :
*****************************************************************}
function ChRech( const sNom: string; vVal: Variant; Sens: Char ): TChampRech;
begin
  Result := ChTri( sNom, Sens );
  Result.Valeur := vVal;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 29/11/2005
Modifié le ... : 30/11/2005
Description .. : Initialise un champ comme critère de tri pour la
Suite ........ : fonction TOBTriOptimise
Suite ........ : Le critère de tri doit être '+' ou '-'
Mots clefs ... :
*****************************************************************}
function ChTri( const sNom: string; Sens: Char ): TChampRech;
begin
  Result.Nom := sNom;
  Result.Sens := Sens;
  Result.Indice := 0;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Y. PEUILLON
Créé le ...... : 20/05/2005
Modifié le ... : 30/11/2005
Description .. : Fonction de recherche d'une TOB fille par méthode
Suite ........ : dichotomique. Attention : la TOB doit être triée et si le tri est
Suite ........ : décroissant, il est indispensable de le spécifier
Mots clefs ... :
*****************************************************************}
function TOBRechercheOptimise( const TOBSrc: TOB; const Field, Valeur: string ): TOB;
var
  TobSrc0: TOB;
  iStart, iMax, iNumChamp,
  iCount, iPrevCount, iComp: Integer;
begin
  Result := nil;
  if( Assigned(TOBSrc) and (TOBSrc.Detail.Count > 0) ) then
  begin
    // Récupération du n° de champ
    TobSrc0 := TOBSrc.Detail[0];
    iNumChamp := TobSrc0.GetNumChamp( Field );
    if( iNumChamp > 0 ) then
    begin
      iStart := 0;
      iPrevCount := -1;
      iCount := 0;
      iMax := TOBSrc.Detail.Count - 1;
      // Vérifie que ce n'est ni le premier, ni le dernier
      if SameText( TOBSrc.Detail[iMax].GetValeur(iNumChamp), Valeur ) then
        Result := TOBSrc.Detail[iMax]
      else if SameText( TobSrc0.GetValeur(iNumChamp ), Valeur ) then
        Result := TobSrc0
      else
        // Recherche dichotomique
        while( not(Assigned(Result)) ) do
        begin
          if( iPrevCount = iCount ) then
            Break;
          iPrevCount := iCount;
          iCount := ( iMax - iStart ) div 2 + iStart;
          iComp := CompareText( TOBSrc.Detail[iCount].GetValeur(iNumChamp), Valeur );
          if( iComp = 0 ) then
            Result := TOBSrc.Detail[iCount]
          else if( iComp < 0 ) then
            iStart := iCount
          else
            iMax := iCount;
        end;
    end;
  end;
end;

function TestClefs( const TOBSrc: TOB; const Champs: array of TChampRech ): Integer;
var
  i: integer;

  function CompareFloat( fValeur, fCible: Extended ): Integer;
  begin
    if( Abs(fCible - fValeur) < 1E-5 ) then
      Result := 0
    else if( fCible > fValeur ) then
      Result := Min( -1, Round(fCible - fValeur) )
    else
      Result := Max( 1, Round(fCible - fValeur) );
  end;

begin
  Result := 0;
  for i := Low(Champs) to High(Champs) do
  begin
    case VarType(Champs[i].Valeur) of
      varByte,
      varSmallint,
      varInteger:   Result := Champs[i].Valeur - TOBSrc.GetValeur( Champs[i].Indice );
      varSingle,
      varDouble,
      varCurrency,
      varDate:      Result := CompareFloat( Champs[i].Valeur, TOBSrc.GetValeur(Champs[i].Indice) );
      varOleStr,
      varStrArg,
      varString:    Result := CompareText( Champs[i].Valeur, TOBSrc.GetValeur(Champs[i].Indice) );
    end;
    if( Result <> 0 ) then
    begin
      if( Champs[i].Sens = '-' ) then
        Result := - Result;
      Break;
    end;
  end;
end;

function TOBRechercheRecursive( const TOBSrc: TOB; const Champs: array of TChampRech;
                                iDeb, iFin: Integer; var iIndice: Integer ): Integer;
var
  iItem,
  iTestClefs: Integer;
begin
  if( iDeb > iFin ) then
  begin
    Result := -1;
    iIndice := iDeb;   // c'est l'indice ou peut s'insérer l'enreg
    Exit;
  end;
  iItem := ( iDeb + iFin ) div 2;
  iTestClefs := TestClefs( TOBSrc.Detail[iItem], Champs );
  if( iTestClefs = 0 ) then
  begin
    if( iItem <> 0 ) then
    begin              // on recule pour se positionner sur le 1er enreg ok
      repeat
        iItem := iItem - 1;
        iTestClefs := TestClefs( TOBSrc.Detail[iItem], Champs );
      until (iTestClefs <> 0) or (iItem = 0);
      if( iTestClefs <> 0 ) then
        Inc( iItem );
    end;
    Result := iItem;
    iIndice := iItem;  // pour insertion avec clé non unique
  end
  else
    if( iTestClefs > 0 ) then
      Result := TOBRechercheRecursive( TOBSrc, Champs, iItem + 1, iFin, iIndice )
    else
      Result := TOBRechercheRecursive( TOBSrc, Champs, iDeb, iItem - 1, iIndice );
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 15/11/2005
Modifié le ... :   /  /
Description .. : Recherche dichotomique sur plusieurs champs, qui
Suite ........ : DOIVENT ETRE TRIES.
Suite ........ : L'indice retourné peut servir pour l'insertion d'une nouvelle
Suite ........ : valeur.
Mots clefs ... :
*****************************************************************}
function TOBRechercheOptimise( const TOBSrc: TOB; Champs: array of TChampRech; var iIndice: Integer ): TOB;
var
  TobSrc0: TOB;
  i, iDeb, iFin: Integer;
begin
  Result := nil;
  if( Assigned(TOBSrc) and (TOBSrc.Detail.Count > 0) ) then
  begin
    // Récupération du n° de champ
    TobSrc0 := TOBSrc.Detail[0];
    for i := Low(Champs) to High(Champs) do
      Champs[i].Indice := TobSrc0.GetNumChamp( Champs[i].Nom );
    iDeb := 0;
    iFin := TOBSrc.Detail.Count - 1;
    if( TOBRechercheRecursive(TOBSrc, Champs, iDeb, iFin, iIndice) >= 0 ) then
      Result := TOBSrc.Detail[iIndice];
  end
  else
    iIndice := -1;
end;

function TOBRechercheOptimise( const TOBSrc: TOB; Champs: array of TChampRech ): TOB;
var
  iIndice: Integer;
begin
  Result := TOBRechercheOptimise( TOBSrc, Champs, iIndice );
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 06/03/2006
Modifié le ... :   /  /
Description .. : Recherche dichotomique - si aucune TOB trouvée,
Suite ........ : création d'une TOB avec les valeurs recherchées
Mots clefs ... :
*****************************************************************}
function TOBRechercheOuCree( const TobSrc: TOB; Champs: array of TChampRech;
                             const sNomTable: string; var bCreee: Boolean ): TOB;
var
  i,
  iIndice: Integer;
begin
  Result := TOBRechercheOptimise( TOBSrc, Champs, iIndice );
  bCreee := not( Assigned(Result) );
  if( bCreee ) then
  begin
    Result := TOB.Create( sNomTable, TobSrc, iIndice );
    for i := Low(Champs) to High(Champs) do
      Result.AddChampSupValeur( Champs[i].Nom, Champs[i].Valeur );
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 15/11/2005
Modifié le ... : 30/11/2005
Description .. : Tri utilisant la recherche dichotomique sur plusieurs champs.
Suite ........ : Si bVideSrc est True, on supprime les filles de la TOBSrc
Mots clefs ... :
*****************************************************************}
function TOBTriOptimise( const TOBSrc: TOB; Champs: array of TChampRech; bVideSrc: Boolean ): TOB;
var
  i, j, iIndice: Integer;
  pTOBD0, pTOBDI: TOB;
begin
  if( Assigned(TOBSrc) ) then
  begin
    Result := TOB.Create( TOBSrc.NomTable, nil, -1 );
    for i := 0 to TOBSrc.Detail.Count - 1 do
    begin
      pTOBD0 := TOBSrc.Detail[0];
      // Initialise les indices des champs
      for j := Low(Champs) to High(Champs) do
        Champs[j].Indice := pTOBD0.GetNumChamp( Champs[j].Nom );
      if( i = 0 ) then
        iIndice := -1
      else
      // Recherche dans la TOB destionation
      begin
        for j := Low(Champs) to High(Champs) do
          Champs[j].Valeur := pTOBD0.GetValeur( Champs[j].Indice );
        TOBRechercheOptimise( Result, Champs, iIndice );
      end;
      // Insertion à l'indice trouvé
      if( bVideSrc ) then
        pTOBDI := pTOBD0
      else
      begin
        pTOBDI := TOB.Create( '', nil, -1 );
        pTOBDI.Dupliquer( pTOBD0, True, True, True );
        pTOBD0.ChangeParent( pTOBD0.Parent, -1 );
      end;
      pTOBDI.ChangeParent( Result, iIndice );
    end;
  end
  else
    Result := nil;
end;

procedure TOBTriOptimise( const TOBSrc: TOB; Champs: array of TChampRech );
var
  pTOBTriee: TOB;
  i: Integer;
begin
  if( Assigned(TOBSrc) ) then
  begin
    pTOBTriee := TOBTriOptimise( TOBSrc, Champs, True );
    for i := 0 to pTOBTriee.Detail.Count - 1 do
      pTOBTriee.Detail[0].ChangeParent( TOBSrc, -1 );
    pTOBTriee.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Y. PEUILLON
Créé le ...... : 20/05/2005
Modifié le ... :   /  /
Description .. : Renomme un champ supplémentaire d'une TOB
Mots clefs ... :
*****************************************************************}

procedure TOBRenameChampSup( pTob: TOB; OldName, NewName: string );
var
  iNumChamp: integer;
  ChampSup: TCS;
begin
  if not Assigned( pTob ) then
    exit;

  iNumChamp := pTob.GetNumChamp( OldName );
  if iNumChamp >= 1000 then
  begin
    ChampSup := TCS( pTob.ChampsSup[iNumChamp - 1000] );
    ChampSup.Nom := NewName;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Y. PEUILLON
Créé le ...... : 20/05/2005
Modifié le ... :   /  /
Description .. : Renomme un champ supplémentaire pour toutes les filles
Suite ........ : d'une TOB
Mots clefs ... :
*****************************************************************}

procedure TOBRenameChampSupAllFille( TobMere: TOB; OldName, NewName: string );
var
  iNumChamp, i: integer;
  ChampSup: TCS;
begin
  if not Assigned( TobMere ) then
    exit;

  iNumChamp := -1;
  for i := 0 to TobMere.Detail.Count - 1 do
  begin
    if iNumChamp < 0 then
    begin
      iNumChamp := TobMere.Detail[i].GetNumChamp( OldName );
      if iNumChamp < 1000 then
        Break;
    end;
    ChampSup := TCS( TobMere.Detail[i].ChampsSup[iNumChamp - 1000] );
    ChampSup.Nom := NewName;
  end;
end;

function ModeTOBtoXMLString( Entree: TOB; Supp: boolean = false; Zipped : boolean = false ) : string;
var
  Sortie: TStringStream;
begin
  Sortie := TStringStream.Create( '' );
  Entree.SaveToXmlStream( Sortie, true, Supp );
{$IFNDEF BTP}
  {$IFNDEF EAGL}
  if Zipped then
    ZipStream( Sortie );
  {$ENDIF EAGL}
{$ENDIF BTP}
  Result := Sortie.DataString;
  Sortie.Free;
end;

function ModeTOBtoBINString( Entree: TOB; Supp: boolean = false; Zipped : boolean = false ) : string;
var
  Sortie: TStringStream;
begin
  Sortie := TStringStream.Create( '' );
  Entree.SaveToBinStream( Sortie, true, true, Supp, true );
{$IFNDEF BTP}
  {$IFNDEF EAGL}
  if Zipped then
    ZipStream( Sortie );
  {$ENDIF EAGL}
{$ENDIF BTP}

  Result := Sortie.DataString;
  Sortie.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 16/03/2005
Modifié le ... :   /  /
Description .. : Met à jour la tob récursivement afin que toutes les filles
Suite ........ : aient les mêmes champs supplémentaires
Suite ........ :
Mots clefs ... :
*****************************************************************}
{
procedure UpdateTobStruc( aTob: Tob; const pChamps: PDEChamps );
var
  i, j, idx: Integer;
  Champs: PDEChamps;
  ChampsFilles: TDEChamps;
  Champ: TDEChamp;
  ChampSup: TCS;
  bStructOk: boolean;
  pListeChampsSup: TStringList;

  function InitValue(iValType: Integer): Variant;
  begin
    case iValType of
      varBoolean:   Result := '-';
      varSmallint,
      varInteger,
      varSingle,
      varDouble,
      varCurrency:  Result := VarAsType(0, iValType);
      varDate:      Result := iDate1900;
      varString,
      varOleStr:    Result := VarAsType('', iValType);
    else
                    Result := varNull;
    end;
  end;


begin
  Champs := pChamps;
  if( aTob is TOB ) then
  begin
    (* if( Champs = nil ) then
    begin
      iNumTable := TrouveNumTable( aTob );
      if( iNumTable > 0 ) then
        Champs := PDEChamps( Pointer(Integer(V_PGI.DEChamps) + (iNumTable * SizeOf(PDEChamp))) );
    end;                                                                                            *)
    if( Assigned(Champs) ) then
    begin
      // Vérifie que l'on a la même structure
      bStructOk := (Length(Champs^) = aTob.ChampsSup.Count);
      if bStructOk then
        for i := 0 to Length(Champs^) - 1 do
          if not(SameText(Champs^[i].nom, TCS(aTob.ChampsSup[i]).Nom)) then
          begin
            bStructOk := False;
            Break;
          end;
      // Sinon, on l'a recrée
      if not(bStructOk) then
      begin
        pListeChampsSup := TStringList.Create;
        pListeChampsSup.Sorted := True;
        for i := 0 to aTob.ChampsSup.Count - 1 do
        begin
          pListeChampsSup.AddObject( UpperCase(TCS(aTob.ChampsSup[i]).Nom), aTob.ChampsSup[i] );
          aTob.ChampsSup[i] := nil;
        end;
        aTob.ChampsSup.Clear;
        for i := 0 to Length(Champs^) - 1 do
        begin
          Champ := Champs^[i];
          aTob.AddChampSupValeur( Champ.Nom, InitValue(DEChampTipeToVarType(Champ.tipe)) );
          idx := pListeChampsSup.IndexOf( UpperCase(Champ.Nom) );
          if( idx >= 0) then
          begin
            TCS(aTob.ChampsSup[i]).Valeur := TCS(pListeChampsSup.Objects[idx]).Valeur;
            TCS(pListeChampsSup.Objects[idx]).Free;
            pListeChampsSup.Delete(idx);
          end;
        end;
        for i := 0 to pListeChampsSup.Count - 1 do
        begin
          ChampSup := TCS( pListeChampsSup.Objects[0] );
          aTob.AddChampSupValeur( ChampSup.Nom, ChampSup.valeur );
          ChampSup.Free;
          pListeChampsSup.Delete(0);
        end;
        pListeChampsSup.Free;
      end;
    end;
    if( aTob.Detail.Count = 1 ) then
      UpdateTobStruc( aTob.Detail[0], nil )
    // On s'assure que toutes les filles ont la même structure
    else if( aTob.Detail.Count > 1 ) then
    begin
      pListeChampsSup := TStringList.Create;
      pListeChampsSup.Sorted := True;
      (* iNumTable := TrouveNumTable( aTob.Detail[0] );
      if( iNumTable > 0 ) then
      begin
        Champs := PDEChamps(Pointer(Integer(V_PGI.DEChamps) + (iNumTable * SizeOf(PDEChamp))));
        SetLength( ChampsFilles, Length(Champs^) + 128 );
        for i := 0 to Length(Champs^) - 1 do
        begin
          ChampsFilles[i] := Champs^[i];
          pListeChampsSup.Add( ChampsFilles[i].Nom );
        end;
      end
      else                                          *)
        SetLength( ChampsFilles, 1024 );
      for i := 0 to aTob.Detail.Count - 1 do
      begin
        for j := 0 to aTob.Detail[i].ChampsSup.Count - 1 do
          if( pListeChampsSup.IndexOf(TCS(aTob.Detail[i].ChampsSup[j]).Nom) < 0 ) then
          begin
            idx := pListeChampsSup.Count + 1;
            if (Length(ChampsFilles) < idx) then
              SetLength( ChampsFilles, idx * 2 );
            Dec( idx );
            ChampSup := TCS( aTob.Detail[i].ChampsSup[j] );
            if (ChampSup.Nom = '') then
              ChampSup.Nom := '_CHAMPSUP' + IntToStr(idx);
            ChampsFilles[idx].Nom := ChampSup.Nom;
            case VarType(ChampSup.Valeur) of
              varBoolean:  ChampsFilles[idx].tipe := 'BOOLEAN';
              varInteger,
              varSmallint: ChampsFilles[idx].tipe := 'INTEGER';
              varSingle,
              varDouble:   ChampsFilles[idx].tipe := 'DOUBLE';
              varDate:     ChampsFilles[idx].tipe := 'DATE';
            else
                           ChampsFilles[idx].tipe := 'VARCHAR';
            end;
            pListeChampsSup.Add( ChampsFilles[idx].Nom );
          end;
      end;
      SetLength( ChampsFilles, pListeChampsSup.Count );
      pListeChampsSup.Free;
      for i := 0 to aTob.Detail.Count - 1 do
        UpdateTobStruc( aTob.Detail[i], @ChampsFilles );
      SetLength( ChampsFilles, 0 );
    end;
  end;
end;
}


{ TlisteChampsTOB }

function TlisteChampsTOB.Add(AObject: TChampsTob): Integer;
begin
  if FindWhere(AObject, Result) then exit; // element déjà présent
  inherited Insert(Result, Aobject);
end;

constructor TlisteChampsTOB.create;
begin
  inherited create;
end;

destructor TlisteChampsTOB.destroy;
begin
  libere;
  inherited;
end;

function TlisteChampsTOB.FindEqual(pTOb: TOB): TchampsTOB;
var
  L, H, I, C: Integer;
  IInterm : TChampsTOB;
begin
  Result := nil;
  L := 0;
  H := Count - 1;
  while L <= H do
  begin
    I := (L + H) shr 1; // divise l'intervale par 2
    C := AnsiCompareText(Items[I].FnomTOB, PTob.NomTable);
    if C < 0 then L := I + 1 else
    begin
      H := I - 1;
      if C = 0 then
      begin
        Result := Items[I]; break;
      end;
    end;
  end;
  //
  if result = nil then
  begin
    iInterm := TchampsTOB.Create (pTOB);
    add(Iinterm);
    result := Iinterm;
  end;
end;

// --------------------------
// recherche dichotomique
// --------------------------
function TlisteChampsTOB.FindWhere(const S: TChampsTob; var Index: Integer): Boolean;
var
  L, H, I, C: Integer;
begin
  Result := False;
  L := 0;
  H := Count - 1;
  while L <= H do
  begin
    I := (L + H) shr 1; // divise l'intervale par 2^1 (soit 2)
    C := AnsiCompareText(Items[I].FnomTOB, S.FnomTOB);
    if C < 0 then L := I + 1 else
    begin
      H := I - 1;
      if C = 0 then
      begin
        Result := True;
      end;
    end;
  end;
  Index := L;
end;

function TlisteChampsTOB.GetBoolean(pTob: TOB;const sNomChamp: string): Boolean;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetBoolean (pTob,sNomChamp);
end;

function TlisteChampsTOB.GetDateTime(pTob: TOB;const sNomChamp: string): TDateTime;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetDateTime (pTob,sNomChamp);
end;

function TlisteChampsTOB.GetDouble(pTob: TOB;const sNomChamp: string): Double;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetDouble (pTob,sNomChamp);
end;

function TlisteChampsTOB.GetIndice(pTOB: TOB;const sNomChamp: string): Integer;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetIndice (pTOB, sNomChamp);
end;

function TlisteChampsTOB.FieldExists (pTOB : TOB; const SNomChamp : string) : boolean;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.FieldExists (sNomChamp);
end;

function TlisteChampsTOB.GetNumChamp(pTOB: TOB;const sNomChamp: string): Integer;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetIndice (pTOB, sNomChamp);
end;

function TlisteChampsTOB.GetInteger(pTob: TOB;const sNomChamp: string): Integer;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetInteger (pTob,sNomChamp);
end;

function TlisteChampsTOB.GetItems(Indice: integer): TChampsTob;
begin
  result := TchampsTOB(inherited Items [Indice]);
end;

function TlisteChampsTOB.GetString(pTob: TOB;const sNomChamp: string): string;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetString (pTob,sNomChamp);
end;

function TlisteChampsTOB.GetValeur(pTob: TOB;const sNomChamp: string): Variant;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetValeur (pTob,sNomChamp);
end;

function TlisteChampsTOB.GetValue(pTob: TOB;const sNomChamp: string): Variant;
var IInterm : TchampsTOB;
begin
  IInterm := findEqual (pTOB);
  result := IINterm.GetValue (pTob,sNomChamp);
end;

procedure TlisteChampsTOB.Initialise;
begin
  Libere;
end;

procedure TlisteChampsTOB.Libere;
var Indice : integer;
begin
  for Indice := 0 to count -1 do
  begin
    Items[Indice].Free;
    Items[Indice]:= nil;
  end;
  clear;
end;

procedure TlisteChampsTOB.AddChampSup(pTOB: TOB; const SnomChamp: string; const toutelesfilles : boolean=false);
var IInterm : TChampsTob;
begin
  pTOB.AddChampSup (SnomChamp,toutelesfilles);
  IInterm := findEqual (pTOB);
  IInterm.AddChampSup (pTOB);
end;

procedure TlisteChampsTOB.PutValue(pTOB: TOB; const sNomChamp: string;sValeur: variant);
var IInterm : TChampsTob;
begin
  IInterm := findEqual (pTOB);
  IInterm.PutValue (pTOB,sNomChamp,sValeur);
end;

procedure TlisteChampsTOB.SetItems(Indice: integer;const Value: TChampsTob);
begin
  Inherited Items[Indice]:= Value;
end;


function TOBFindWhere(TOBF: TOB; TheChamp : string; ATrouver : variant; var Index: Integer): Boolean;

  function compareVariant (A,B : variant ) : integer;
  begin
    TRY
      if A > B then result := 1 else if A < B then result := -1 else result := 0;
    EXCEPT
      raise;
    END;
  end;

var
  L, H, I, C : Integer;
  TOBL : TOB;
begin
  Result := False;
  L := 0;
  H := TOBF.detail.Count - 1;
  while L <= H do
  begin
    I := (L + H) shr 1; // divise l'intervale par 2^1 (soit 2)
    TOBL := TOBF.detail[i];
    C := CompareVariant(TheOptiTOB.GetValue(TOBL,TheChamp), Atrouver);
    if C < 0 then L := I + 1 else
    begin
      H := I - 1;
      if C = 0 then
      begin
        Result := True;
      end;
    end;
  end;
  Index := L;
end;


initialization
  InitTOBUtil;
finalization
  FinTOBUtil;

end.


