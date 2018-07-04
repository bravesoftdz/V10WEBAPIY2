unit TobUtil;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  SysUtils,
  HEnt1,
  {$IFDEF GCGC}
  wCommuns,
  {$ENDIF GCGC}
  IniFiles,
  UTob;


type
  {$IFNDEF GCGC}
  TDEChamps = array of TDEChamp;
  PDEChamps = ^TDEChamps;
  {$ENDIF GCGC}

  TVerifChamp = ( vcAucun, vcSupp );

  TChampRech = record
    Nom: hstring;
    Sens: Char;
    Indice: Integer;
    Valeur: Variant;
  end;
  PChampRech = ^TChampRech;

  TDeuxChamps = record
    Champ1,
    Champ2: string;
  end;

  TChampsTob = class
  private
    FVerifChamp: TVerifChamp;
    FTob: TOB;
    FListChamps: THashedStringList;
    procedure Init( pTob: TOB );
  public
    constructor Create( pTob: TOB ); overload;
    constructor Create( pTob: TOB; const NomChamps: array of hstring ); overload;
    destructor Destroy; override;
    function GetIndice( const sNomChamp: hstring ): Integer;
    function GetValeur( pTob: TOB; const sNomChamp: hstring ): Variant;
    function GetString( pTob: TOB; const sNomChamp: hstring ): hstring;
    function GetDouble( pTob: TOB; const sNomChamp: hstring ): Double;
    function GetInteger( pTob: TOB; const sNomChamp: hstring ): Integer;
    function GetBoolean( pTob: TOB; const sNomChamp: hstring ): Boolean;
    function GetDateTime( pTob: TOB; const sNomChamp: hstring ): TDateTime;
    property VerifChamp: TVerifChamp read FVerifChamp write FVerifChamp;
  end;

function ChRech( const sNom: hstring; vVal: Variant; Sens: Char = '+' ): TChampRech;
function ChTri( const sNom: hstring; Sens: Char = '+' ): TChampRech;

function DEChampTipeToVarType( const sTipe: hstring ): Integer;
procedure IndicesMinMaxChampsTob( pTob: Tob; var iMin, iMax: Integer; bTobVirtuelle: Boolean = False );
function FOSauveTobBin( sNomFichier: TFileName; pTob: Tob ): boolean;
procedure CopieChampsTob( const TOBSrc: TOB; TobDest: TOB; RemplaceChampsNonVides: boolean );
function TOBRechercheOptimise( const TobSrc: TOB; const Field, Valeur: hstring;
                               bMinMaj: Boolean = False ): TOB; overload;
function TOBRechercheOptimise( const TobSrc: TOB; Champs: array of TChampRech;
                               bMinMaj: Boolean = False ): TOB; overload;
function TOBRechercheOptimise( const TobSrc: TOB; Champs: array of TChampRech;
                               var iIndice: Integer; bMinMaj: Boolean = False ): TOB; overload;
function TOBRechercheOuCree( const TobSrc: TOB; Champs: array of TChampRech;
                             const sNomTable: hstring; var bCreee: Boolean;
                             bMinMaj: Boolean = False ): TOB;
function TOBRechercheSuivant( const TobSrc: TOB; Champs: array of TChampRech;
                              var iIndice: Integer; bMinMaj: Boolean = False ): TOB;
function TOBTriOptimise( const TOBSrc: TOB; Champs: array of TChampRech;
                         bVideSrc: Boolean; bMinMaj: Boolean = False ): TOB; overload;
procedure TOBTriOptimise( const TOBSrc, TOBDest: TOB; Champs: array of TChampRech;
                          bVideSrc: Boolean; bMinMaj: Boolean = False ); overload;
procedure TOBTriOptimise( const TOBSrc: TOB; Champs: array of TChampRech;
                          bMinMaj: Boolean = False ); overload;

procedure TOBRenameChampSup( pTob: TOB; OldName, NewName: string );
procedure TOBRenameChampSupAllFille( TobMere: TOB; OldName, NewName: string );
procedure TOBRenameChampsSupAllFille( TobMere: TOB; Champs: array of TDeuxChamps );

function ModeTOBtoXMLString( Entree: TOB; Supp: boolean = false; Zipped : boolean = false ) : hstring;
function ModeTOBtoBINString( Entree: TOB; Supp: boolean = false; Zipped : boolean = false ) : hstring;

procedure UpdateTobStruc( aTob: Tob; const pChamps: PDEChamps );

implementation

uses
{$IFDEF VER150}
  Variants,
{$ENDIF VER150}
{$IFDEF GCGC}
  EntGC,
  UtilPGI,
{$ENDIF GCGC}
  {$IFNDEF EAGL}
  EBizUtil,
  {$ENDIF EAGL}
  Math,
  Classes,
  TntWideStrUtils;                                     


{ TChampsTob }


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
end;

{***********UNITE*************************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 08/02/2006
Modifié le ... : 15/02/2006
Description .. : Variante du constructeur qui permet d'initialiser directement
Suite ........ : les indices d'un tableau champs.
Mots clefs ... :
*****************************************************************}
constructor TChampsTob.Create( pTob: TOB; const NomChamps: array of hstring );
var
  i: Integer;
begin
  inherited Create();
  Init( pTob );
  for i := Low( NomChamps ) to High( NomChamps ) do
    FListChamps.AddObject( NomChamps[i], Pointer(FTob.GetNumChamp(NomChamps[i])) );
end;

procedure TChampsTob.Init( pTob: TOB );
begin
  FTob := TOB.Create( '', nil, -1 );
  FTob.Dupliquer( pTob, False, False );
  FListChamps := THashedStringList.Create;
  FVerifChamp := vcSupp;
end;

destructor TChampsTob.Destroy;
begin
  FTob.Free;
  FListChamps.Free;
  inherited;
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
function TChampsTob.GetIndice( const sNomChamp: hstring ): Integer;
var
  iIndiceChamp: Integer;
begin
  iIndiceChamp := FListChamps.IndexOf( sNomChamp );
  if( iIndiceChamp >= 0 ) then
    Result := Integer( FListChamps.Objects[iIndiceChamp] )
  else
  begin
    Result := FTob.GetNumChamp( sNomChamp );
    if ( Result > 0 ) then
      FListChamps.AddObject( sNomChamp, Pointer(Result) );
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
function TChampsTob.GetValeur( pTob: TOB; const sNomChamp: hstring ): Variant;
var
  iIndice: Integer;
begin
  Result := varEmpty;
  if( Assigned(pTob) ) then
  begin
    iIndice := GetIndice( sNomChamp );
    if( (VerifChamp = vcAucun) or ((VerifChamp = vcSupp) and (iIndice < 1000)) or
        SameText(sNomChamp, pTob.GetNomChamp(iIndice)) ) then
      Result := pTob.GetValeur( iIndice )
    else
      Result := pTob.GetValue( sNomChamp );
  end;
end;

function TChampsTob.GetDouble( pTob: TOB; const sNomChamp: hstring ): Double;
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

function TChampsTob.GetInteger( pTob: TOB; const sNomChamp: hstring ): Integer;
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

function TChampsTob.GetString( pTob: TOB; const sNomChamp: hstring ): hstring;
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

function TChampsTob.GetBoolean(pTob: TOB; const sNomChamp: hstring): Boolean;
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

function TChampsTob.GetDateTime(pTob: TOB; const sNomChamp: hstring): TDateTime;
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


{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 27/01/2005
Modifié le ... : 27/01/2005
Description .. : Retourne le type de variant correspondant au
Suite ........ : DEChamp[][].Tipe
Mots clefs ... :
*****************************************************************}

function DEChampTipeToVarType( const sTipe: hstring ): Integer;
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
{$IFDEF GCGC}
    if VH_GC.ModeDeconnecte then
      bTobVirtuelle := True;
{$ENDIF GCGC}
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
{$IFDEF GCGC}
    on E: Exception do
      MODEAfficheErreurSQL( TraduireMemoire('Exception :') + ' ' + E.Message );
{$ENDIF GCGC}
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
  NomChamp: hstring;
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
function ChRech( const sNom: hstring; vVal: Variant; Sens: Char ): TChampRech;
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
function ChTri( const sNom: hstring; Sens: Char ): TChampRech;
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
function TOBRechercheOptimise( const TOBSrc: TOB; const Field, Valeur: hstring; bMinMaj: Boolean ): TOB;
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

function TestClefs( const TOBSrc: TOB; const Champs: array of TChampRech; bMinMaj: Boolean ): Integer;
var
  i: integer;
  val, cible: variant;
  sVal, sCible: WideString;

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
    val := Champs[i].Valeur;
    if ( Champs[i].Indice > 0 ) then
      cible := TOBSrc.GetValeur( Champs[i].Indice )
    else
      cible := TOBSrc.GetValue( Champs[i].Nom );
    case VarType(val) of
      varByte,
      varSmallint,
      varInt64,
      varInteger:   Result := val - cible;
      varSingle,
      varDouble,
      varCurrency,
      varDate:      Result := CompareFloat( val, cible );
      varStrArg,
      varString:    if bMinMaj then
                      Result := AnsiCompareStr( val, cible )
                    else
                      Result := CompareText( val, cible );
      varOleStr:    if bMinMaj then
                    begin
                      sVal := val;
                      sCible := cible;
                      Result := WStrComp( PWideChar(sVal), PWideChar(sCible) );
                    end
                    else
                      Result := CompareText( val, cible );
    end;
    if( Result <> 0 ) then
    begin
      if( Champs[i].Sens = '-' ) then
        Result := - Result;
      Break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : G. EYDAN
Créé le ...... : 17/09/2007
Modifié le ... :   /  /    
Description .. : On ne se repositionne pas sur le premier élément 
Suite ........ : si iIndice = MaxInt
Mots clefs ... : 
*****************************************************************}
function TOBRechercheRecursive( const TOBSrc: TOB; const Champs: array of TChampRech;
                                iDeb, iFin: Integer; var iIndice: Integer; bMinMaj: Boolean ): Integer;
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
  iTestClefs := TestClefs( TOBSrc.Detail[iItem], Champs, bMinMaj );
  if( iTestClefs = 0 ) then
  begin
    if( iItem <> 0 ) and ( iIndice <> MaxInt ) then
    begin              // on recule pour se positionner sur le 1er enreg ok
      repeat
        iItem := iItem - 1;
        iTestClefs := TestClefs( TOBSrc.Detail[iItem], Champs, bMinMaj );
      until (iTestClefs <> 0) or (iItem = 0);
      if( iTestClefs <> 0 ) then
        Inc( iItem );
    end;
    Result := iItem;
    iIndice := iItem;  // pour insertion avec clé non unique
  end
  else
    if( iTestClefs > 0 ) then
      Result := TOBRechercheRecursive( TOBSrc, Champs, iItem + 1, iFin, iIndice, bMinMaj )
    else
      Result := TOBRechercheRecursive( TOBSrc, Champs, iDeb, iItem - 1, iIndice, bMinMaj );
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
function TOBRechercheOptimise( const TOBSrc: TOB; Champs: array of TChampRech;
                               var iIndice: Integer; bMinMaj: Boolean ): TOB;
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
      if ( Champs[i].Indice <= 0 ) or  ( Champs[i].Indice >= 1000 ) then
        Champs[i].Indice := TobSrc0.GetNumChamp( Champs[i].Nom );
    iDeb := 0;
    iFin := TOBSrc.Detail.Count - 1;
    if( TOBRechercheRecursive(TOBSrc, Champs, iDeb, iFin, iIndice, bMinMaj) >= 0 ) then
      Result := TOBSrc.Detail[iIndice];
  end
  else
    iIndice := -1;
end;

function TOBRechercheOptimise( const TOBSrc: TOB; Champs: array of TChampRech; bMinMaj: Boolean ): TOB;
var
  iIndice: Integer;
begin
  Result := TOBRechercheOptimise( TOBSrc, Champs, iIndice, bMinMaj );
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
                             const sNomTable: hstring; var bCreee: Boolean; bMinMaj: Boolean ): TOB;
var
  i,
  iIndice: Integer;
begin
  Result := TOBRechercheOptimise( TOBSrc, Champs, iIndice, bMinMaj );
  bCreee := not( Assigned(Result) );
  if( bCreee ) then
  begin
    Result := TOB.Create( sNomTable, TobSrc, iIndice );
    for i := Low(Champs) to High(Champs) do
      Result.AddChampSupValeur( Champs[i].Nom, Champs[i].Valeur );
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 06/07/2007
Modifié le ... : 06/07/2007
Description .. : Recherche la fille suivante sur plusieurs champs dans une 
Suite ........ : TOB mère qui doit être triée sur les champs de recherche.
Suite ........ : 
Suite ........ : L'indice passé est celui de la fille courante.
Mots clefs ... : 
*****************************************************************}
function TOBRechercheSuivant( const TobSrc: TOB; Champs: array of TChampRech; var iIndice: Integer; bMinMaj: Boolean ): TOB;
begin
  Result := nil;
  if ( Assigned( TobSrc ) ) and ( iIndice >= 0 ) then
  begin
    Inc( iIndice );
    if iIndice < TobSrc.Detail.Count then
    begin
      if TestClefs( TobSrc.Detail[ iIndice ], Champs, bMinMaj ) = 0 then
        Result := TobSrc.Detail[ iIndice ];
    end;
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
function TOBTriOptimise( const TOBSrc: TOB; Champs: array of TChampRech;
                         bVideSrc, bMinMaj: Boolean ): TOB;
begin
  Result := TOB.Create( TOBSrc.NomTable, nil, -1 );
  TOBTriOptimise( TOBSrc, Result, Champs, bVideSrc, bMinMaj );
end;

procedure TOBTriOptimise( const TOBSrc, TOBDest: TOB; Champs: array of TChampRech;
                          bVideSrc, bMinMaj: Boolean ); overload;
var
  i, j, iIndice: Integer;
  pTOBD0, pTOBDI: TOB;
begin
  if Assigned( TOBSrc ) and Assigned( TOBDest ) then
  begin
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
        TOBRechercheOptimise( TOBDest, Champs, iIndice, bMinMaj );
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
      pTOBDI.ChangeParent( TOBDest, iIndice );
    end;
  end;
end;

procedure TOBTriOptimise( const TOBSrc: TOB; Champs: array of TChampRech; bMinMaj: Boolean );
var
  pTOBTriee: TOB;
  i: Integer;
begin
  if( Assigned(TOBSrc) ) then
  begin
    pTOBTriee := TOBTriOptimise( TOBSrc, Champs, True, bMinMaj );
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

{***********A.G.L.***********************************************
Auteur  ...... : O. Grosclaude
Créé le ...... : 13/03/2008
Modifié le ... :   /  /    
Description .. : Renomme les champs sup des filles
Suite ........ : Ne fonctionne que si ces champs ont les mêmes indices 
Suite ........ : dans toutes les filles
Mots clefs ... : 
*****************************************************************}
procedure TOBRenameChampsSupAllFille( TobMere: TOB; Champs: array of TDeuxChamps );
var
  i, j: integer;
  iNumChamp: array of integer;
  TobD: TOB;
begin
  if not Assigned( TobMere ) then
    exit;

  SetLength( iNumChamp, High(Champs) - Low(Champs) + 1 );
  for i := 0 to TobMere.Detail.Count - 1 do
  begin
    TobD := TobMere.Detail[i];
    for j := 0 to Length( iNumChamp ) - 1 do
    begin
      // Récupère l'indice sur le champ
      if ( i <= 0 ) then
        iNumChamp[j] := TobD.GetNumChamp( Champs[j + Low(Champs)].Champ1 ) - 1000;
      // Change le nom
      if ( iNumChamp[j] >= 0 ) and
         SameText( TCS( TobD.ChampsSup[iNumChamp[j]] ).Nom, Champs[j + Low(Champs)].Champ1 ) then
        TCS( TobD.ChampsSup[iNumChamp[j]] ).Nom := Champs[j + Low(Champs)].Champ2;
    end;
  end;
end;


function ModeTOBtoXMLString( Entree: TOB; Supp: boolean = false; Zipped : boolean = false ) : hstring;
var
  Sortie: TStringStream;
begin
  Sortie := TStringStream.Create( '' );
  Entree.SaveToXmlStream( Sortie, true, Supp );

  {$IFNDEF EAGL}
  if Zipped then
    ZipStream( Sortie );
  {$ENDIF EAGL}

  Result := Sortie.DataString;
  Sortie.Free;
end;

function ModeTOBtoBINString( Entree: TOB; Supp: boolean = false; Zipped : boolean = false ) : hstring;
var
  Sortie: TStringStream;
begin
  Sortie := TStringStream.Create( '' );
  Entree.SaveToBinStream( Sortie, true, true, Supp, true );

  {$IFNDEF EAGL}
  if Zipped then
    ZipStream( Sortie );
  {$ENDIF EAGL}

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
              varInt64,
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


(*
procedure TestTri();
var
  i, j: Integer;
  sVal: hstring;
  fVal: Single;
  pTOB, pTOB_D: TOB;
begin
  Randomize();
  SetLength( sVal, 5 );
  pTOB := TOB.Create( '', nil, -1 );
  for i := 0 to 15 do
  begin
    pTOB_D := TOB.Create( '', pTOB, -1 );
    if( i mod 3 = 0 ) then
      for j := 1 to Length(sVal) do
        sVal[j] := Chr( 65 + Random(26) );
    fVal := Round( Random(9999) ) / 10;
    pTOB_D.AddChampSupValeur( 'CH1', sVal );
    pTOB_D.AddChampSupValeur( 'CH2', fVal );
  end;
  ShowTob( pTOB );
  TOBTriOptimise( pTOB, [ChTri('CH1'), ChTri('CH2', '-')] );
  ShowTob( pTOB );
  pTOB.Free;
end;

initialization
  TestTri(); *)

end.



