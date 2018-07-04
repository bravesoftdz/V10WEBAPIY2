unit ZRubrique;

interface

uses
  { Delphi }
  SysUtils
  , classes
{$IFDEF EAGLCLIENT}
{$ELSE}
  , dbtables
{$ENDIF}
  { AGL }
  , uTob
  , hctrls
  , HEnt1
  { Compta }
  , Ent1
  , CalcOLE
  , CpteUtil
  ;

type
  TZRubrique = class
  private
    fRubrique: TOB;
    fIndex: integer;
    fCourante: TOB;
    { Filtres }
    fDevise: string;
    fEtablissement: string;
    fExercice: string;
    fDateDebut: TDateTime;
    fDateFin: TDateTime;
    fTypeEcr: SetttTypePiece;
    fBalSit : string;
    fAxe: string;
    fSection1: string;
    fSection2: string;
    fPrefixe: string;
    function GetRubrique(stRubrique: string): integer;
    function FaitRequeteDonnees(stSQLCompte: string): string;
    procedure QuelTypEcr(St: string; var SetTyp: SetttTypePiece);
    function PositionneCumuls(pTotalDebit, pTotalCredit: double; stCompte:
      string; var TResult: TabloExt): double;
  public
    constructor Create;
    destructor Destroy; override;
    function InitCriteres(pstEtablissement, pstDevise, pstDate,
      pstTypeEcr: string; pstBalsit: string = ''; pstAxe: string = '';
        pstSection1: string = ''; pstSection2: string = ''): boolean;
    function GetValeur(stRubrique: string; LesComptes: TStringList): variant;
    { Filtres }
    property Devise: string read fDevise write fDevise;
    property Etablissement: string read fEtablissement write fEtablissement;
    property Exercice: string read fExercice write fExercice;
    property DateDebut: TDateTime read fDateDebut write fDateDebut;
    property DateFin: TDateTime read fDateFin write fDateFin;
    property TypeEcr: SetttTypePiece read fTypeEcr write fTypeEcr;
  end;

implementation

{ TZRubrique }

const
  TAILLE_CACHE = 20;

constructor TZRubrique.Create;
begin
  fDevise := '';
  fAxe := '';
  fRubrique := TOB.Create('ZRUBRIQUE', nil, -1);
end;

destructor TZRubrique.Destroy;
begin
  if fRubrique <> nil then
    FreeAndNil(fRubrique);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/05/2005
Modifié le ... :   /  /
Description .. : Constitution de la requête de calcul des cumuls sur la table
Suite ........ : écriture. Cette requête est directement inspirée du
Suite ........ : FabricReqCpt3 de cpteutil.pas
Mots clefs ... :
*****************************************************************}

function TZRubrique.FaitRequeteDonnees(stSQLCompte: string): string;
var
  stSQL: string;
  stDev: string;
  stCompte1, stCompte2 : string;
begin
  if fBalsit <> '' then
  begin
    stCompte1 := fPrefixe + 'COMPTE1';
    stCompte2 := fPrefixe + 'COMPTE2';
  end else
  begin
    stCompte1 := fPrefixe + 'GENERAL';
    stCompte2 := fPrefixe + 'AUXILIAIRE';
  end;
  { SELECT }
  stSQL := 'SELECT ' + stCompte1 + ' COMPTE1 , ' + stCompte2 + ' COMPTE2 , ';
  if (fDevise <> '') and (fDevise <> V_PGI.DevisePivot) then
    stDev := 'DEV'
  else
    stDev := '';
  stSQL := stSQL + ' SUM (' + fPrefixe + 'DEBIT' + stDev + ') TOTDEBIT, SUM (' +
    fPrefixe + 'CREDIT' + stDev + ') TOTCREDIT ';
  if fAxe <> '' then
    stSQL := stSQL + ' , Y_AXE, Y_SECTION ';

  if fBalsit <> '' then
    stSQL := stSQL + ' FROM CBALSITECR '
  else if fAxe = '' then
    stSQL := stSQL + ' FROM ECRITURE '
  else
    stSQL := stSQL + ' FROM ANALYTIQ ';

  { JOINTURE SUR LES GENERAUX }
  stSQL := stSQL + ' LEFT JOIN GENERAUX ON (G_GENERAL=' + stCompte1 + ') ';

  { WHERE }
  stSQL := stSQL + ' WHERE ' + stCompte1 + ' IN (' + stSQLCompte + ')';

  if fBalsit = '' then
  begin
    if fExercice <> '' then
      stSQL := stSQL + ' AND ' + fPrefixe + 'EXERCICE = "' + fExercice + '" ';

    if fDateDebut <> iDate1900 then
      stSQL := stSQL + ' AND ' + fPrefixe + 'DATECOMPTABLE>="' +
        UsDateTime(fDateDebut) + '" ';

    if fDateFin <> iDate1900 then
      stSQL := stSQL + ' AND ' + fPrefixe + 'DATECOMPTABLE<="' +
        UsDateTime(fDateFin) + '" ';

    if fDevise <> '' then
      stSQL := stSQL + ' AND ' + fPrefixe + 'DEVISE = "' + fDevise + '" ';

    if fEtablissement <> '' then
      stSQL := stSQL + ' AND ' + fPrefixe + 'ETABLISSEMENT = "' + fEtablissement +
        '" ';

    stSQL := stSQL + WhereSupp(fPrefixe, fTypeEcr);
  end;

  if fAxe <> '' then
  begin
    stSQL := stSQL + ' AND Y_AXE="' + fAxe + '" ';
    if fSection1 <> '' then
      stSQL := stSQL + ' AND Y_SECTION>="' + fSection1 + '" ';
    if fSection2 <> '' then
      stSQL := stSQL + ' AND Y_SECTION<="' + fSection2 + '" ';
  end;

  { GROUP BY }
  stSQL := stSQL + ' GROUP BY ' + stCompte1 + ',' + stCompte2 + ' ';
  if fAxe <> '' then
    stSQL := stSQL + ' , Y_AXE, Y_SECTION ';

  { ORDER BY }
  stSQL := stSQL + ' ORDER BY ' + stCompte1 + ',' + stCompte2 + ' ';
  if fAxe <> '' then
    stSQL := stSQL + ' , Y_AXE, Y_SECTION ';

  Result := stSQL;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/05/2005
Modifié le ... :   /  /
Description .. : Chargement des informations d'une rubrique
Mots clefs ... :
*****************************************************************}

function TZRubrique.GetRubrique(stRubrique: string): integer;
var
  i: integer;
  Q: TQuery;
  lStColonne: string;
begin
  Result := -1;
  fIndex := -1;
  fCourante := nil;

  { On commence par chercher la rubrique en mémoire }
  for i := 0 to fRubrique.Detail.Count - 1 do
  begin
    if (fRubrique.Detail[i].GetValue('RB_RUBRIQUE') = stRubrique) then
    begin
      Result := i;
      fIndex := i;
      fCourante := fRubrique.Detail[i];
      break;
    end;
  end;
  { Rubrique non trouvée en mémoire }
  if (Result = -1) then
  begin
    lStColonne := ' RB_RUBRIQUE, RB_COMPTE1, RB_EXCLUSION1, RB_SIGNERUB ';
    Q := OpenSQL('SELECT ##TOP ' + IntToStr(TAILLE_CACHE) + '## ' + lStColonne +
      ' FROM RUBRIQUE WHERE RB_RUBRIQUE>="' + stRubrique + '"', True);
    try
      if not Q.Eof then
      begin
        if fRubrique.LoadDetailDB('RUBRIQUE', '', '', Q, False) then
        begin
          Result := 0;
          fIndex := 0;
          fCourante := fRubrique.Detail[0];
        end;
      end;
    finally
      Ferme(Q);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/05/2005
Modifié le ... :   /  /
Description .. : Renvoie la valeur d'une rubrique.
Mots clefs ... :
*****************************************************************}

function TZRubrique.GetValeur(stRubrique: string; LesComptes: TStringList):
  variant;
var
  stSQL, stSQLCompte, stExclu, stDetailCompte: string;
  Q: TQuery;
  TResult: TabloExt;
  i: integer;
  dValRub: double;
begin
  if (GetRubrique(stRubrique) >= 0) then
  begin
    Result := 0;
    { Liste des comptes concernés par la rubrique }
    stSQLCompte := 'SELECT G_GENERAL FROM GENERAUX WHERE '
      + AnalyseCompte(fCourante.GetValue('RB_COMPTE1'), fbGene, False, False);
    stExclu := AnalyseCompte(fCourante.GetValue('RB_EXCLUSION1'), fbGene, True,
      False);
    if stExclu <> '' then
      stSQLCompte := stSQLCompte + ' AND ' + stExclu;

    { Constitution de la requête sur la table des données }
    stSQL := FaitRequeteDonnees(stSQLCompte);

    {Calcul des cumuls }
    Q := OpenSQL(stSQL, True);
    try
      while not Q.Eof do
      begin

        dValRub := PositionneCumuls(Q.FindField('TOTDEBIT').AsFloat,
          Q.FindField('TOTCREDIT').AsFloat, Q.FindField('COMPTE1').AsString, TResult);

        { Mémorisation du détail des comptes }
        if LesComptes <> nil then
        begin
          stDetailCompte := Q.FindField('COMPTE1').AsString;
          if (fAxe <> '') then
            stDetailCompte := stDetailCompte + ';' + Q.FindField(fPrefixe +
              'SECTION').AsString;
          for i := 1 to 6 do
            stDetailCompte := stDetailCompte + ':' + FloatToStr(TResult[i]);
          stDetailCompte := stDetailCompte + ':' +
            fCourante.GetValue('RB_SIGNERUB');
          LesComptes.Add(stDetailCompte);
        end;

        { Calcul du cumul de la rubrique }
        Result := Arrondi(Result + dValRub, V_PGI.OkDecV);
        { Passage au compte suivant }
        Q.Next;
      end;
    finally
      Ferme(Q);
    end;
  end
  else
    Result := 'Erreur - rubrique introuvable.';
end;

function TZRubrique.InitCriteres(pstEtablissement, pstDevise, pstDate,
  pstTypeEcr: string; pstBalsit: string = ''; pstAxe: string = ''; pstSection1:
    string = ''; pstSection2: string = ''): boolean;
var
  Err: integer;
  Exo: TExoDate;
begin
  fEtablissement := pstEtablissement;
  fDevise := pstDevise;
  fBalsit := pstBalSit;
  fAxe := pstAxe;
  if fBalsit <> '' then
    fPrefixe := 'BSE_'
  else if fAxe = '' then
    fPrefixe := 'E_'
  else
    fPrefixe := 'Y_';

  fSection1 := pstSection1;
  fSection2 := pstSection2;
  QuelTypEcr(pStTypeEcr, fTypeEcr);
  Result := (WhatDate(pstDate, fDateDebut, fDateFin, Err, Exo) or (fBalsit<>''));
  if Result then
    fExercice := Exo.Code;
end;

procedure TZRubrique.QuelTypEcr(St: string; var SetTyp: SetttTypePiece);
begin
  SetTyp := [];
  if Pos('N', St) > 0 then
    SetTyp := SetTyp + [tpReel];
  if Pos('S', St) > 0 then
    SetTyp := SetTyp + [tpSim];
  if Pos('P', St) > 0 then
    SetTyp := SetTyp + [tpPrev];
  if Pos('U', St) > 0 then
    SetTyp := SetTyp + [tpSitu];
  if Pos('R', St) > 0 then
    SetTyp := SetTyp + [tpRevi];
  if Pos('I', St) > 0 then
    SetTyp := SetTyp + [tpIfrs];
  if SetTyp = [] then
    SetTyp := [tpReel];
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 01/06/2005
Modifié le ... :   /  /
Description .. : Calcul des cumuls pour une ligne de détail de la rubrique
Mots clefs ... :
*****************************************************************}

function TZRubrique.PositionneCumuls(pTotalDebit, pTotalCredit: double;
  stCompte: string; var TResult: TabloExt): double;
var
  stTypeSolde: string;
  dValRub: double;
  bUneFois: boolean;
begin
  stTypeSolde := QuelTypDeSolde(fCourante.GetValue('RB_COMPTE1'), stCompte,
    fbGene, bUneFois);
  TResult[1] := Arrondi(pTotalDebit - pTotalCredit, V_PGI.OkDecV);
  if (pTotalCredit > pTotalDebit) then
    TResult[2] := Arrondi(pTotalCredit - pTotalDebit, V_PGI.OkDecV)
  else
    TResult[2] := 0;
  if (pTotalDebit > pTotalCredit) then
    TResult[3] := Arrondi(pTotalDebit - pTotalCredit, V_PGI.OkDecV)
  else
    TResult[3] := 0;
  TResult[4] := pTotalCredit;
  TResult[5] := pTotalDebit;

  dValRub := 0;
  if (fCourante.GetValue('RB_SIGNERUB') = 'NEG') then
  begin
    if (stTypeSolde = 'SM') then
      dValRub := (-1) * TResult[1]
    else if (stTypeSolde = 'SC') then
      dValRub := TResult[2]
    else if (stTypeSolde = 'SD') then
      dValRub := (-1) * TResult[3]
    else if (stTypeSolde = 'TC') then
      dValRub := TResult[4]
    else if (stTypeSolde = 'TD') then
      dValRub := (-1) * TResult[5];
  end
  else
  begin
    if (stTypeSolde = 'SM') then
      dValRub := TResult[1]
    else if (stTypeSolde = 'SC') then
      dValRub := (-1) * TResult[2]
    else if (stTypeSolde = 'SD') then
      dValRub := TResult[3]
    else if (stTypeSolde = 'TC') then
      dValRub := (-1) * TResult[4]
    else if (stTypeSolde = 'TD') then
      dValRub := TResult[5];
  end;
  if (stTypeSolde = 'SM') then TResult[6] := 7
  else if (stTypeSolde = 'SC') then TResult[6] := 6
  else if (stTypeSolde = 'SD') then TResult[6] := 5
  else if (stTypeSolde = 'TC') then TResult[6] := 3
  else if (stTypeSolde = 'TD') then TResult[6] := 2;
  Result := dValRub;
end;

end.

