{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  10/04/06   JP   Création de l'unité : Répartition de la saisie par type d'écritures
--------------------------------------------------------------------------------------}
unit CPVignetteGraphTypeEcr;

interface

uses
  Classes, UTob, DBTables, uCPVignettePlugIn, uToolsOWC, HCtrls,
  OWC11_TLB;

type
  TObjGraphTypeEcr = class(TCPVignettePlugIn)
  private
    tQualif : TOB;
    
    procedure RemplitTobDonnees;
    procedure TobToGraph(var Obj : TOWCChart);
    function  GetCouleur(QualifPiece : string) : string;
    procedure AddSpecific1 (oChart : chChart);
    procedure AddSpecific2 (oChart : chChart);
    procedure ColorSpecific(oChart : chChart);
  protected
    procedure RecupDonnees ;             override;
    function  GetInterface(NomGrille : string = '') : Boolean; override;
  public
    constructor Create(TobIn, TobOut : TOB); override;
    destructor  Destroy; override;
    function    InitPeriode : Char; override;
  end;

implementation

uses
  SysUtils, HEnt1, CalcOLE, UProcGen, NumConv, eSession;

{---------------------------------------------------------------------------------------}
procedure TObjGraphTypeEcr.RecupDonnees;
{---------------------------------------------------------------------------------------}
var
  Obj : TOWCChart;
  T   : TOB;
begin
  inherited;
  ddwriteln('PGICPGRAPHTYPECR : RecupDonnees = CHART');
  RemplitTobDonnees;
  if TobDonnees.Detail.Count > 0 then begin
    Obj := TOWCChart.Create(TobResponse, AspectGraph);
    try
      try
        ddWriteLN('PGICPGRAPHTYPECR : Debut du chargement des écritures');
        TobToGraph(Obj);
        {27/06/06 : Pour la gestion des couleurs des séries}
        Obj.AddSpecific  := AddSpecific1;
        Obj.AddSpecific2 := AddSpecific2;

        Obj.owIsAxeRightDegroupe := False;
        Obj.owHasTwoCharts := True;

        Obj.TypeGraph := tyCamembert;
        Obj.owTitre.Visible := True;
        Obj.owTitre.Bold := True;
        Obj.owTitre.FontSize := 11;

        Obj.owLegend.Visible := True;
        Obj.owLegend.Position := poPositionBottom;
        Obj.owAxeBot.Visible := False;
        Obj.owAxeRight.Visible := False;
        Obj.owAxeLeft.Visible := False;
        Obj.owAxeRight.InfoAxe.HasLabel := True;
        Obj.owAxeLeft.InfoAxe.HasLabel := True;

        if TobDonnees.Detail[0].GetString('TYPE') = 'Charge' then begin
          Obj.owTitre.Caption := 'Charges';
          T := TobDonnees.FindFirst(['TYPE'], ['Produit'], True);
          if T <> nil then
            Obj.owTitreChart2 := 'Produits'
          else
            Obj.owHasTwoCharts := False;
        end
        else begin
          Obj.owTitre.Caption := 'Produits';
          Obj.owHasTwoCharts := False;
        end;

        Obj.owLabel.Visible := False; {True; Le résultat n'est pas probant}
        Obj.owLabel.InfoAxe.LabelPerct := False;
        Obj.owLabel.InfoAxe.LabelValue := True;
        Obj.owLabel.InfoAxe.NumFormat := '#,##0';
        Obj.owLabel.Color := colWhite;


        Obj.Build;
        ddWriteLN('PGICPGRAPHTYPECR : Fin du chargement des écritures');
      except
        on E : Exception do
          MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
      end;
    finally
      if Assigned(Obj) then FreeAndNil(Obj);
    end;
  end;
end;

{Affectation des critères de traitement à partir des critères de sélection de la vignette
{---------------------------------------------------------------------------------------}
function TObjGraphTypeEcr.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetInterface;
end;

{---------------------------------------------------------------------------------------}
constructor TObjGraphTypeEcr.Create(TobIn, TobOut: TOB);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(TobIn, TobOut);
  tQualif := TOB.Create('AAAA', nil, -1);
end;

{---------------------------------------------------------------------------------------}
destructor TObjGraphTypeEcr.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tQualif) then FreeAndNil(tQualif);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphTypeEcr.RemplitTobDonnees;
{---------------------------------------------------------------------------------------}
var
  TP  : TOB;
  TC  : TOB;
  TCL : TOB; 
  Per : Char;
  Q   : TOB;
  QQ  : TOB;
  SQL : string;
  dd  : string;
  n   : Integer;
begin
  SansInterface := True;

  TobDonnees.ClearDetail;

  ddWriteLN('PGICPGRAPHTYPECR : Début des GetCumuls');
  Per := Periode;

  case Per of
    'M' : dd := 'E_DATECOMPTABLE BETWEEN "' + UsDateTime(DebutDeMois(DateRef)) + '" AND "' + UsDateTime(FinDeMois(DateRef)) + '" ';
    'T' : dd := 'E_DATECOMPTABLE BETWEEN "' + UsDateTime(StrToDate(DebutTrimestre(DateRef))) + '" AND "' + UsDateTime(StrToDate(FinTrimestre(DateRef))) + '" ';
    'A' : dd := 'E_DATECOMPTABLE BETWEEN "' + UsDateTime(DebutAnnee(DateRef)) + '" AND "' + UsDateTime(FinAnnee(DateRef)) + '" ';
    'E' : dd := 'E_EXERCICE = "' + GereExercice('EX_EXERCICE', DateRef) + '" ';
  end;

  SQL := 'SELECT ABS(SUM(E_DEBIT - E_CREDIT)) MONTANT, G_NATUREGENE, E_QUALIFPIECE, CO_LIBELLE FROM ECRITURE ' +
         'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
         'LEFT JOIN COMMUN ON CO_CODE = E_QUALIFPIECE AND CO_TYPE="QFP" ' +
         'WHERE G_NATUREGENE IN ("CHA", "PRO") AND ' + dd +
         'GROUP BY G_NATUREGENE, E_QUALIFPIECE, CO_LIBELLE  ORDER BY G_NATUREGENE, E_QUALIFPIECE';

  ddWriteLN('PGICPGRAPHIMPAYE : ' + SQL);

  Q := OpenSelectInCache(SQL);
  try
    for n := 0 to Q.Detail.Count - 1 do begin
      QQ := Q.Detail[n];
      TCL := tQualif.FindFirst(['QUALIF'], [QQ.GetString('CO_LIBELLE')], True);
      if not Assigned(TCL) then begin
        TCL := TOB.Create('AAAA', tQualif, -1);
        TCL.AddChampSupValeur('QUALIF' , QQ.GetString('CO_LIBELLE'));
        TCL.AddChampSupValeur('COULEUR', GetCouleur(QQ.GetString('E_QUALIFPIECE')));
      end;

      if QQ.GetString('G_NATUREGENE') = 'CHA' then begin
        TC := TOB.Create('ùùùù', TobDonnees, -1);
        TC.AddChampSupValeur('TYPE', 'Charge');
        TC.AddChampSupValeur('QUALIF', QQ.GetString('CO_LIBELLE'));
        TC.AddChampSupValeur('E_MONTANT', QQ.GetDouble('MONTANT'));
      end else begin
        TP := TOB.Create('ùùùù', TobDonnees, -1);
        TP.AddChampSupValeur('TYPE', 'Produit');
        TP.AddChampSupValeur('QUALIF', QQ.GetString('CO_LIBELLE'));
        TP.AddChampSupValeur('E_MONTANT', QQ.GetDouble('MONTANT'));
      end;
    end;
  finally
    FreeAndNil(Q);
  end;
  ddWriteLN('PGICPGRAPHTYPECR : Charge écritures');
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphTypeEcr.TobToGraph(var Obj : TOWCChart);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 0 to TobDonnees.Detail.Count - 1 do
    Obj.SetValue(TobDonnees.Detail[n].GetString('QUALIF'),
                 TobDonnees.Detail[n].GetString('TYPE'),
                 TobDonnees.Detail[n].GetDouble('E_MONTANT'));
end;

{---------------------------------------------------------------------------------------}
function TObjGraphTypeEcr.InitPeriode : Char;
{---------------------------------------------------------------------------------------}
begin
  Result := 'E';
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphTypeEcr.ColorSpecific(oChart : chChart);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  s : string;
  T : TOB;
begin
  if TobDonnees.Detail.Count > 0 then begin
    for n := 0 to oChart.SeriesCollection[0].Points.Count - 1 do begin
      s := oChart.SeriesCollection[0].Points[n].GetValue(chDimCategories, False);
      T := tQualif.FindFirst(['QUALIF'], [s], True);
      if Assigned(T) then
        oChart.SeriesCollection[0].Points[n].Interior.Set_Color(T.GetString('COULEUR'))
      else
        oChart.SeriesCollection[0].Points[n].Interior.Set_Color(TOWCChart.GetCouleurGraph(ColMediumSeaGreen));
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjGraphTypeEcr.GetCouleur(QualifPiece : string) : string;
{---------------------------------------------------------------------------------------}
var
  c : Char;
begin
  Result := TOWCChart.GetCouleurGraph(ColMediumSeaGreen);

  if QualifPiece = '' then Exit;

  c := QualifPiece[1];

  case c of
    'I' : Result := TOWCChart.GetCouleurGraph(colLightSkyBlue);
    'S' : Result := TOWCChart.GetCouleurGraph(colLightPink);
    'R' : Result := TOWCChart.GetCouleurGraph(colLightSalmon);
    'N' : Result := TOWCChart.GetCouleurGraph(colLightSeaGreen);
    'P' : Result := TOWCChart.GetCouleurGraph(colLightCyan);
    'U' : Result := TOWCChart.GetCouleurGraph(colLightGoldenrodYellow);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphTypeEcr.AddSpecific1(oChart : chChart);
{---------------------------------------------------------------------------------------}
begin
  ColorSpecific(oChart);
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphTypeEcr.AddSpecific2(oChart : chChart);
{---------------------------------------------------------------------------------------}
begin
  ColorSpecific(oChart);
end;

end.

