{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 03/05/2006
Modifié le ... :   /  /
Description .. : Liste des cartes de séjour qui arrivent à expiration
               : Vignette : PG_VIG_CARTSEJOUR
               : Tablette : PGPERIODEVIGNETTE
               : Table    : SALARIES
Mots clefs ... :
*****************************************************************}
unit PGVIGCARTSEJOUR;

interface
uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  PGVIGUTIL,
  uToolsOWC,
  uToolsPlugin,
  HCtrls;
type
  PG_VIG_CARTSEJOUR = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeCartSejour(AParam : string);

 protected
    procedure RecupDonnees; override;
    function GetInterface (NomGrille: string = ''): Boolean; override;
    procedure GetClauseWhere; override;
    procedure DrawGrid (Grille: string); override;
    function SetInterface : Boolean ; override;
 public
 end;


 implementation
 uses
  SysUtils,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_CARTSEJOUR.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_CARTSEJOUR.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_CARTSEJOUR.RecupDonnees;
begin
  inherited;
    AfficheListeCartSejour('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_CARTSEJOUR.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'DATEXPIRSEJOUR', 'C.0 ---');
end;

function PG_VIG_CARTSEJOUR.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
//TOBResponse.savetofile('c:\datatob.txt',false,true,true);
end;


procedure PG_VIG_CARTSEJOUR.AfficheListeCartSejour(AParam : string);
var
//@@  Q                      : TQuery;
  StSQL                  : String;
  DataTob                : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i: integer;
  T: TOB;

begin

  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

  Try
    DataTob := TOB.Create('~TEMP', nil, -1);

    StSQL := 'SELECT  PSA_LIBELLE, PSA_PRENOM, PY_NATIONALITE, '+
             'PSA_DATEXPIRSEJOUR FROM SALARIES '+
             'LEFT JOIN PAYS ON PY_PAYS=PSA_NATIONALITE '+
             'WHERE '+
             'PSA_DATEXPIRSEJOUR >="' + USDATETIME (EnDateDu) + '" AND '+
             'PSA_DATEXPIRSEJOUR <="' + USDATETIME (EnDateAu) + '"'+
             ' ORDER BY PSA_DATEXPIRSEJOUR';

    DataTob := OpenSelectInCache (StSQL);
    ConvertFieldValue(DataTob);
    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
      T.AddChampSupValeur('NATIONALITE',DataTob.Detail[i].GetValue('PY_NATIONALITE'));
      T.AddChampSupValeur('DATEXPIRSEJOUR',DateToStr(DataTob.Detail[i].GetValue('PSA_DATEXPIRSEJOUR')));
    end;

    if (TobDonnees.detail.count = 0) then
      PutGridDetail('GRSALARIE', TobDonnees);

  finally
    FreeAndNil (DataTob);
  end;

end;

end.

