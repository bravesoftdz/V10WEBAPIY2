{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 03/05/2006
Modifié le ... :   /  /
Description .. : Liste des salariés entrés  (Liste)
               : Vignette : PG_VIG_LSTENTREE
               : Tablette : PGPERIODEVIGNETTE
               : Table    : SALARIES
Mots clefs ... :
*****************************************************************}
unit PGVIGLSTENTREE;

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
  PG_VIG_LSTENTREE = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeEntrees(AParam : string);

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

function PG_VIG_LSTENTREE.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_LSTENTREE.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_LSTENTREE.RecupDonnees;
begin
  inherited;
    AfficheListeEntrees('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_LSTENTREE.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'DATEENTREE', 'C.0 ---');
    SetFormatCol(Grille, 'DATESORTIE', 'C.0 ---');
end;

function PG_VIG_LSTENTREE.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
//TOBResponse.savetofile('c:\datatob.txt',false,true,true);
end;


procedure PG_VIG_LSTENTREE.AfficheListeEntrees(AParam : string);
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
             'PSA_DATEENTREE >="' + USDATETIME (EnDateDu) + '" AND '+
             'PSA_DATEENTREE <="' + USDATETIME (EnDateAu) + '"'+
             ' ORDER BY PSA_DATEENTREE';
    DataTob := OpenSelectInCache (StSQL);
    ConvertFieldValue(DataTob);
    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
      T.AddChampSupValeur('DATEENTREE',DateToStr(DataTob.Detail[i].GetValue('PSA_DATEENTREE')));
      T.AddChampSupValeur('DATESORTIE',DateToStr(DataTob.Detail[i].GetValue('PSA_DATESORTIE')));
    end;
    if (TobDonnees.detail.count = 0) then
        PutGridDetail('GRSALARIE', TobDonnees);


  finally
    FreeAndNil (DataTob);
  end;

end;

end.
