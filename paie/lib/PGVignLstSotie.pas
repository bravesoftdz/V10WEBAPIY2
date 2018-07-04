{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 04/2006
Modifié le ... :   /  /
Description .. : Liste des salariés sortis  (Liste)
               : Vignette : PGVIGNLSTSORTIE
               : Tablette : PGPERIODEVIGNETTE
               : Table    : SALARIES
Mots clefs ... :
*****************************************************************}
unit PGVignLstSotie;

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
  PGVignLstSortie = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeSorties(AParam : string);

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

function PGVignLstSortie.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');

  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PGVignLstSortie.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PGVignLstSortie.RecupDonnees;
begin
  inherited;
    AfficheListeSorties('');

end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PGVignLstSortie.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'PSA_DATEENTREE', 'C.0 ---');
    SetFormatCol(Grille, 'PSA_DATESORTIE', 'C.0 ---');
end;

function PGVignLstSortie.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
end;


procedure PGVignLstSortie.AfficheListeSorties(AParam : string);
var
  StSQL : String;
  DataTob : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i: integer;
  T: TOB;
begin

  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

  Try
    DataTob := TOB.Create('~TEMP', nil, -1);

    StSQL := 'SELECT  PSA_LIBELLE, PSA_PRENOM, PSA_DATEENTREE, '+
             'PSA_DATESORTIE FROM SALARIES '+
             'WHERE '+
             'PSA_DATESORTIE >="' + USDATETIME (EnDateDu) + '" AND '+
             'PSA_DATESORTIE <="' + USDATETIME (EnDateAu) + '"'+
             ' ORDER BY PSA_DATESORTIE';
    DataTob := OpenSelectInCache (StSQL);
    ConvertFieldValue(DataTob);
    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
      T.AddChampSupValeur('PSA_DATEENTREE',DateToStr(DataTob.Detail[i].GetValue('PSA_DATEENTREE')));
      T.AddChampSupValeur('PSA_DATESORTIE',DateToStr(DataTob.Detail[i].GetValue('PSA_DATESORTIE')));
    end;
//@@    TobDonnees.savetofile('c:\departs.txt',false,true,true);

    if (TobDonnees.detail.count =  0) then
      PutGridDetail('GRSALARIE', TobDonnees);

  finally
    FreeAndNil (DataTob);
  end;

end;

end.
