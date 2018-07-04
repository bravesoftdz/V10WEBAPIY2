{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 03/05/2006
Modifié le ... :   /  /    
Description .. : Liste des fins de CDD de la période
               : Vignette : PG_VIG_FINCDD
               : Tablette : PGPERIODEVIGNETTE
               : Table    : CONTRATTRAVAIL, SALARIES
Mots clefs ... :
*****************************************************************}
unit PGVIGFINCDD;

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
  PG_VIG_FINCDD = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeFinCDD(AParam : string);

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

function PG_VIG_FINCDD.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_FINCDD.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_FINCDD.RecupDonnees;
begin
  inherited;
  AfficheListeFinCDD('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_FINCDD.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'DEBUTCONTRAT', 'C.0 ---');
    SetFormatCol(Grille, 'FINCONTRAT', 'C.0 ---');
end;

function PG_VIG_FINCDD.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
//TOBResponse.savetofile('c:\datatob.txt',false,true,true);
end;


procedure PG_VIG_FINCDD.AfficheListeFinCDD(AParam : string);
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

    StSQL := 'SELECT  PSA_LIBELLE, PSA_PRENOM, PCI_DEBUTCONTRAT, '+
             'PCI_FINCONTRAT FROM CONTRATTRAVAIL '+
             'LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE '+
             'WHERE '+
             'PCI_FINCONTRAT >="' + USDATETIME (EnDateDu) + '" AND '+
             'PCI_FINCONTRAT <="' + USDATETIME (EnDateAu) + '"'+
             ' ORDER BY PCI_FINCONTRAT';

    DataTob := OpenSelectInCache (StSQL);
    ConvertFieldValue(DataTob);
    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
      T.AddChampSupValeur('DEBUTCONTRAT',DateToStr(DataTob.Detail[i].GetValue('PCI_DEBUTCONTRAT')));
      T.AddChampSupValeur('FINCONTRAT',DateToStr(DataTob.Detail[i].GetValue('PCI_FINCONTRAT')));
    end;
    if (TobDonnees.detail.count = 0) then
      PutGridDetail('GRSALARIE', TobDonnees);

  finally
    FreeAndNil (DataTob);
  end;

end;

end.

