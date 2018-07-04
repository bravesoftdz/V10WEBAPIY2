{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 03/05/2006
Modifié le ... :   /  /    
Description .. : Liste des fins de période d'essai
               : Vignette : PG_VIG_FINESSAi
               : Tablette : PGPERIODEVIGNETTE
               : Table    : CONTRATTRAVAIL, SALARIES
Mots clefs ... : 
*****************************************************************}
unit PGVIGFINESSAI;

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
  PG_VIG_FINESSAI = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeFinEssai(AParam : string);

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

function PG_VIG_FINESSAI.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_FINESSAI.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_FINESSAI.RecupDonnees;
begin
  inherited;
  AfficheListeFinEssai('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_FINESSAI.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'ESSAIDEBUT', 'C.0 ---');
    SetFormatCol(Grille, 'ESSAIFIN', 'C.0 ---');
end;

function PG_VIG_FINESSAI.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
//TOBResponse.savetofile('c:\datatob.txt',false,true,true);
end;


procedure PG_VIG_FINESSAI.AfficheListeFinEssai(AParam : string);
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

    StSQL := 'SELECT  PSA_LIBELLE, PSA_PRENOM, PCI_ESSAIDEBUT, '+
             'PCI_ESSAIFIN FROM CONTRATTRAVAIL '+
             'LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE '+
             'WHERE '+
             'PCI_ESSAIFIN >="' + USDATETIME (EnDateDu) + '" AND '+
             'PCI_ESSAIFIN <="' + USDATETIME (EnDateAu) + '"'+
             ' ORDER BY PCI_ESSAIFIN';
    DataTob := OpenSelectInCache (StSQL);
    ConvertFieldValue(DataTob);
    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
      T.AddChampSupValeur('ESSAIDEBUT',DateToStr(DataTob.Detail[i].GetValue('PCI_ESSAIDEBUT')));
      T.AddChampSupValeur('ESSAIFIN',DateToStr(DataTob.Detail[i].GetValue('PCI_ESSAIFIN')));
    end;
    if (TobDonnees.detail.count = 0) then
      PutGridDetail('GRSALARIE', TobDonnees);

  finally
    FreeAndNil (DataTob);
  end;

end;

end.

