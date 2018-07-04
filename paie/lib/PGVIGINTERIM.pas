{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 12/06/2006
Modifié le ... :   /  /
Description .. : liste des intérimaires de la période
               : Vignette : PG_VIG_INTERIM
               : Tablette : PGPERIODEVIGNETTE
               : Table    : INTERIMAIRES, EMPLOIINTERIM
Mots clefs ... :
*****************************************************************}
unit PGVIGINTERIM;

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
  PG_VIG_INTERIM = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeInterim(AParam : string);

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

function PG_VIG_INTERIM.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_INTERIM.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_INTERIM.RecupDonnees;
begin
  inherited;
    AfficheListeInterim('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_INTERIM.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'DEBUTEMPLOI', 'C.0 ---');
    SetFormatCol(Grille, 'FINEMPLOI', 'C.0 ---');
end;

function PG_VIG_INTERIM.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
//TOBResponse.savetofile('c:\datatob.txt',false,true,true);
end;


procedure PG_VIG_INTERIM.AfficheListeInterim(AParam : string);
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

    StSQL := 'SELECT PSI_LIBELLE, PSI_PRENOM,PEI_DEBUTEMPLOI, '+
             'PEI_FINEMPLOI FROM INTERIMAIRES LEFT JOIN EMPLOIINTERIM '+
             'ON PEI_INTERIMAIRE = PSI_INTERIMAIRE WHERE '+
             'PEI_DEBUTEMPLOI <= "'+USDATETIME (EnDateAu)+'" AND '+
             'PEI_FINEMPLOI >= "'+USDATETIME (EnDateDu)+'" '+
             'ORDER BY PEI_DEBUTEMPLOI';

    DataTob := OpenSelectInCache (StSQL);
    ConvertFieldValue(DataTob);
    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSI_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSI_PRENOM'));
      T.AddChampSupValeur('DEBUTEMPLOI',DateToStr(DataTob.Detail[i].GetValue('PEI_DEBUTEMPLOI')));
      T.AddChampSupValeur('FINEMPLOI',DateToStr(DataTob.Detail[i].GetValue('PEI_FINEMPLOI')));
    end;
    if (TobDonnees.detail.count = 0) then
      PutGridDetail('GRSALARIE', TobDonnees);

  finally
    FreeAndNil (DataTob);
  end;

end;

end.

 
