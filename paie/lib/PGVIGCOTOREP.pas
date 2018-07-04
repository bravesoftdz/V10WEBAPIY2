{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 13/06/2006
Modifié le ... :   /  /
Description .. : liste des handicapés en fin de droit COTOREP
               : Vignette : PG_VIG_COTOREP
               : Tablette : PGPERIODEVIGNETTE
               : Table    : HANDICAPE, SALARIES
Mots clefs ... :
*****************************************************************}
unit PGVIGCOTOREP;

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
  PG_VIG_COTOREP = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeCotorep(AParam : string);

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

function PG_VIG_COTOREP.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_COTOREP.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_COTOREP.RecupDonnees;
begin
  inherited;
    AfficheListeCotorep('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_COTOREP.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'DATEFINCOT', 'C.0 ---');
end;

function PG_VIG_COTOREP.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
end;


procedure PG_VIG_COTOREP.AfficheListeCotorep(AParam : string);
var
  StSQL                  : String;
  DataTob                : TOB;
  EnDateDu, EnDateAu     : TDatetime;
  i: integer;
  T: TOB;
begin

  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);

  Try
    DataTob := TOB.Create('~TEMP', nil, -1);
    StSQL := 'SELECT PSA_LIBELLE, PSA_PRENOM, PGH_DATEFINCOT '+
             'FROM HANDICAPE '+
             'LEFT JOIN SALARIES '+
             'ON PGH_SALARIE = PSA_SALARIE WHERE '+
             'PGH_COTOREP ="X" AND '+
             'PGH_DATEFINCOT >= "'+USDATETIME (EnDateDu)+'" AND '+
             'PGH_DATEFINCOT <= "'+USDATETIME (EnDateAu)+'"';

    DataTob := OpenSelectInCache (StSQL);
    DataTob.SaveToFile ('C:\DataTob.TXT',FALSE,TRUE, TRUE);
    ConvertFieldValue(DataTob);
    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
      T.AddChampSupValeur('DATEFINCOT',DateToStr(DataTob.Detail[i].GetValue('PGH_DATEFINCOT')));
    end;
    if (TobDonnees.Detail.Count = 0) then
      PutGridDetail('GRSALARIE', TobDonnees);

  finally
    FreeAndNil (DataTob);
  end;

end;

end.


