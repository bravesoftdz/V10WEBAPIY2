{***********UNITE*************************************************
Auteur  ...... : Paie Pgi - MF
Créé le ...... : 14/06/2006
Modifié le ... :   /  /
Description .. : liste des handicapés en fin de droit COTOREP
               : Vignette : PG_VIG_VISITMED
               : Tablette : PGPERIODEVIGNETTE
               : Table    : VISITEMEDTRAV
Mots clefs ... :
*****************************************************************
PT1  | 29/05/2008 | FLO | Prise en compte du reponsable
PT2  | 29/05/2008 | FLO | Adaptation pour multidossier
}
unit PGVIGVISITMED;

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
  PG_VIG_VISITMED = class (TGAVignettePaie)
 private
    sN1            : string;

    procedure AfficheListeVisites(AParam : string);

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
  PGOutilsFormation,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_VISITMED.GetInterface (NomGrille: string): Boolean;
begin
  inherited GetInterface ('');
  Result := true;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_VISITMED.GetClauseWhere;
begin
  inherited;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_VISITMED.RecupDonnees;
begin
  inherited;
    AfficheListeVisites('');
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_VISITMED.DrawGrid (Grille: string);
begin
  inherited;
    SetFormatCol(Grille, 'DATEVISITE', 'C.0 ---');
end;

function PG_VIG_VISITMED.SetInterface: Boolean;
begin
  inherited SetInterface;
  result:=true;
end;


procedure PG_VIG_VISITMED.AfficheListeVisites(AParam : string);
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

    StSQL := 'SELECT PSA_LIBELLE, PSA_PRENOM, PVM_DATEVISITE '+
             'FROM VISITEMEDTRAV '+
             'LEFT JOIN SALARIES ON PVM_SALARIE = PSA_SALARIE '+
             'LEFT JOIN DEPORTSAL ON PVM_SALARIE = PSE_SALARIE '+ //PT1
             'WHERE PVM_DATEVISITE >= "'+USDATETIME (EnDateDu)+'" AND '+
             'PVM_DATEVISITE <= "'+USDATETIME (EnDateAu)+'" AND '+
             'PVM_APTE <> "X" AND PVM_INAPTE <> "X" AND '+
             'PVM_RECLASSE <> "X" AND PVM_APTESR <> "X" '+
             'AND '+AdaptByTypeResp('RESPONSABS',V_PGI.UserSalarie,False,'PSE'); //PT1

	StSQL := AdaptMultiDosForm (StSQL); //PT2
    DataTob := OpenSelectInCache (StSQL);
    ConvertFieldValue(DataTob);
    
    for i:=0 to DataTob.detail.count-1 do
    begin
      T := TOB.Create('£REPONSE',TobDonnees,-1);
      T.AddChampSupValeur('NOM',DataTob.Detail[i].GetValue('PSA_LIBELLE'));
      T.AddChampSupValeur('PRENOM',DataTob.Detail[i].GetValue('PSA_PRENOM'));
      T.AddChampSupValeur('DATEVISITE',DateToStr(DataTob.Detail[i].GetValue('PVM_DATEVISITE')));
    end;
    if (TobDonnees.detail.count = 0) then
      PutGridDetail('GRSALARIE', TobDonnees);

  finally
    FreeAndNil (DataTob);
  end;
end;

end.

