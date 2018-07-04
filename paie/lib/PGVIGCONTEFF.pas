{***********UNITE*************************************************
Auteur  ...... : Paie -MF
Créé le ...... : 04/05/2006
Modifié le ... :
Description .. : Vignette des effectifs par contrat (Camembert)
               : Vignette : PGV_VIG_CONTEFF
               : Tablette : PGPERIODEVIGNETTE
               : Table    : CONTRATTRAVAIL, SALARIES
Mots clefs ... :
*****************************************************************}
unit PGVIGCONTEFF;

interface

uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  uToolsOWC,
  PGVignettePaie,
  PGVIGUTIL,
  HCtrls;

type
  PG_VIG_CONTEFF = class (TGAVignettePaie)
  private
    EnDateDu, EnDateAu: TDateTime;

  procedure TobToGraph(var Obj : TOWCChart);

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
  HEnt1,
  SysUtils;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_CONTEFF.GetInterface (NomGrille: string): Boolean;
var
//@@  Periode : string;
  sN1        : string;

begin
  Result := inherited GetInterface ('');
  EnDateDu := iDate1900;
  EnDateAu := iDate2099;

  // calcul des dates de début et fin de période
  DatesPeriode(DateRef, EnDateDu, EnDateAu, Periode,sN1);


end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_CONTEFF.GetClauseWhere;
begin
  inherited;
   ClauseWhere := ' WHERE  (PCI_DEBUTCONTRAT <= "'+USDATETIME (EnDateAu)+'") AND '+
                  '((PCI_FINCONTRAT >= "'+USDATETIME (EnDateDu)+'") or '+
                  '(PCI_FINCONTRAT) = "'+USDATETIME(IDate1900)+'") '+
                  'GROUP BY CO_LIBELLE '+
                  'ORDER BY NB DESC';
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_CONTEFF.RecupDonnees;
var
  TobL, MaTob: TOB;
  i_ind: integer;
  St : String;
  Nbval : integer;
  NbAutres : double;

begin
  inherited;

  NbAutres := 0;
  Nbval := StrToInt(GetCritereValue('NBVAL'))  ;
  if (Nbval = 0) then
    Nbval := 6;

  st := 'SELECT COUNT (DISTINCT PCI_SALARIE) NB , CO_LIBELLE '+
        'FROM CONTRATTRAVAIL LEFT JOIN COMMUN ON '+
        'CO_TYPE="PCT"  AND CO_CODE=PCI_TYPECONTRAT';

  try
    try
      MaTob := OpenSelectInCache (st + ClauseWhere);

      if MaTob.Detail.Count < Nbval then
          Nbval := MaTob.Detail.Count  ;

      for i_ind := 0 to Nbval - 1 do
      begin
       TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
       TobL.AddChampSupValeur ('TYPECONTRAT', MaTob.detail [i_ind] .Getstring ('CO_LIBELLE'));
       TobL.AddChampSupValeur ('LENOMBRE', MaTob.detail [i_ind] .Getvalue ('NB'));
      end;

      if MaTob.Detail.Count > Nbval then
      begin
       TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
       TobL.AddChampSupValeur ('TYPECONTRAT', 'Autres');

       for i_ind := Nbval to MaTob.Detail.Count - 1 do
       begin
        NbAutres := NbAutres +  MaTob.detail [i_ind] .GetDouble ('NB');
       end;
       TobL.AddChampSupValeur ('LENOMBRE', NbAutres);
      end;


    except
      on E: Exception do
        MessageErreur := 'Erreur lors du traitement des données avec le message :'#13#13 + E.Message;
    end;
  finally
    FreeAndNil (MaTob);
  end;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_CONTEFF.DrawGrid (Grille: string);
begin
  inherited;
end;

function PG_VIG_CONTEFF.SetInterface: Boolean;
var  Obj : TOWCChart;
begin
  result := TRUE;
  if TOBDonnees.detail.Count < 1 then exit;

  Obj := TOWCChart.Create(TobResponse, AspectGraph);
  try
    try
      TobToGraph(Obj);
      Obj.TypeGraph := tyCamembert;
      Obj.SetWidth (250);
      Obj.SetHEight(250);
      Obj.owLegend.Position := poPositionLeft;
      Obj.owLabel.Visible := True;
      Obj.owLabel.Color := colLightYellow;

      Obj.Build;
    except
      on E : Exception do
        MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
    end;
  finally
    if Assigned(Obj) then FreeAndNil(Obj);
  end;
     result := TRUE;
end;

procedure PG_VIG_CONTEFF.TobToGraph(var Obj: TOWCChart);
var
  n : Integer;
begin
  for n := 0 to TobDonnees.Detail.Count - 1 do
    Obj.SetValue(TobDonnees.Detail[n].GetString('TYPECONTRAT'), 'Effectif', TobDonnees.Detail[n].GetDouble('LENOMBRE'));
end;

end.
