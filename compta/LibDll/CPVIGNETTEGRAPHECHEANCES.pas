{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  13/02/06   JP   Création de l'unité : Exemple de vignettes Graph
--------------------------------------------------------------------------------------}
unit CPVIGNETTEGRAPHECHEANCES;

interface

uses
  Classes, UTob,
  {$IFNDEF DBXPRESS}
  dbtables,
  {$ELSE}
  uDbxDataSet,
  {$ENDIF}
  uCPVignettePlugIn,
  uToolsOWC, HCtrls, OWC11_TLB;

type
  TObjGraphEche = class(TCPVignettePlugIn)
  private
    EtatOk : Boolean;

    procedure RemplitTobDonnees;
    procedure RemplitTobDonneesEtat;
    procedure TobToGraph(var Obj : TOWCChart);
    procedure ColorSpecific(oChart : chChart);
  protected
    procedure RecupDonnees ;             override;
    function  GetInterface(NomGrille : string = '') : Boolean; override;
    procedure TraiteParam(Param : string); override;
  public
    constructor Create(TobIn, TobOut : TOB); override;
    function InitPeriode : Char; override;
  end;

implementation

uses
  SysUtils, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEche.RecupDonnees;
{---------------------------------------------------------------------------------------}
var
  Obj : TOWCChart;
begin
  inherited;
  SansInterface := True;

  if EtatOk then begin
    ddwriteln('PGICPGRAPHECHEANCE : RecupDonnees = ETAT');
    RemplitTobDonneesEtat;
    FTip := 'E';
    FNat := 'VCP';
    FModele := 'VCE';
    FLanguePrinc := 'FRA';
    Rapport;
  end
  else begin
    ddwriteln('PGICPGRAPHECHEANCE : RecupDonnees = CHART');
    RemplitTobDonnees;
    if TobDonnees.Detail.Count > 0 then begin
      Obj := TOWCChart.Create(TobResponse, AspectGraph);
      try
        try
          ddWriteLN('PGICPGRAPHECHEANCE : Debut du chargement des écritures');
          TobToGraph(Obj);
          Obj.AddSpecific := ColorSpecific;

          Obj.owHasMajorLine := False;
          Obj.TypeGraph := tyHistogramme;
          Obj.owTitre.Visible := False;

          Obj.owLegend.Visible := True;
          Obj.owLegend.Position := poPositionBottom;

//          Obj.owAxeLeft.Caption := 'Montant';
  //        Obj.owAxeLeft.Bold := True;
          Obj.owAxeLeft.InfoAxe.NumFormat := '#,##0' +  ' K' + GetDevisePivot(True);
          Obj.Build;
          ddWriteLN('PGICPGRAPHECHEANCE : Fin du chargement des écritures');
        except
          on E : Exception do
            MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
        end;
      finally
        if Assigned(Obj) then FreeAndNil(Obj);
      end;
    end;
  end;
end;

{Affectation des critères de traitement à partir des critères de sélection de la vignette
{---------------------------------------------------------------------------------------}
function TObjGraphEche.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetInterface;
end;

{---------------------------------------------------------------------------------------}
constructor TObjGraphEche.Create(TobIn, TobOut: TOB);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(TobIn, TobOut);
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEche.RemplitTobDonnees;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Q   : TOB;
  QQ  : TOB;
  TF  : TOB;
  TC  : TOB;
  n   : Integer;
  dd  : TDateTime;
begin
  SansInterface := True;
  
  TobDonnees.ClearDetail;
  dd := iDate1900;
  TF := nil;
  TC := nil;

  n := 7;
       if Periode = 'Z' then n := 14
  else if Periode = 'M' then n := 30;

  SQL := 'SELECT ABS(SUM(E_CREDIT - E_DEBIT - E_COUVERTURE)) MONTANT, E_DATEECHEANCE, E_NATUREPIECE FROM ECRITURE WHERE ' +
         'E_DATEECHEANCE BETWEEN "' + UsDateTime(DateRef) + '" AND "' + UsDateTime(DateRef + n) + '" AND ' +
         'E_NATUREPIECE IN ("FF", "AF", "AC", "FC") AND E_QUALIFPIECE = "N" AND E_ETATLETTRAGE IN ("PL", "AL") ' +
         'AND E_TYPEMVT = "TTC" GROUP BY E_DATEECHEANCE, E_NATUREPIECE ' +
         ' UNION ' +
         ' SELECT ABS(SUM(E_CREDIT - E_DEBIT - E_COUVERTURE)) MONTANT, E_DATEECHEANCE, E_NATUREPIECE ' +
         ' FROM ECRITURE,JOURNAL,GENERAUX ' +
         ' where E_JOURNAL=E_JOURNAL ' +
         ' and E_AUXILIAIRE<>"" ' +
         ' and J_EFFET="X" ' +
         ' and E_ETATLETTRAGE IN ("PL", "AL") ' +
         ' and G_GENERAL=E_GENERAL ' +
         ' and G_EFFET="X" ' +
         ' and E_DATEECHEANCE BETWEEN "' + UsDateTime(DateRef) + '" AND "' + UsDateTime(DateRef + n) + '" ' +
         ' GROUP BY E_DATEECHEANCE, E_NATUREPIECE ' +
         ' ORDER BY E_NATUREPIECE, E_DATEECHEANCE ' ;


  ddWriteLN('PGICPGRAPHECHEANCE : ' + SQL);
  Q := OpenSelectInCache(SQL);
  try
    for n := 0 to Q.Detail.Count - 1 do begin
      QQ := Q.Detail[n];
      if dd <> QQ.GetDateTime('E_DATEECHEANCE') then begin
        dd := QQ.GetDateTime('E_DATEECHEANCE');
        TC := TOB.Create('ùùùù', TobDonnees, -1);
        TF := TOB.Create('ùùùù', TobDonnees, -1);
        TC.AddChampSupValeur('ADATE', FormatDateTime('dd/mm/yy', dd));
        TF.AddChampSupValeur('ADATE', FormatDateTime('dd/mm/yy', dd));
        TC.AddChampSupValeur('TIERS', 'Clients');
        TF.AddChampSupValeur('TIERS', 'Fournisseurs');
        TC.AddChampSupValeur('MONTANT', 0.0);
        TF.AddChampSupValeur('MONTANT', 0.0);
      end;

      if Assigned(TF) and Assigned(TC) then begin
        SQL := QQ.GetString('E_NATUREPIECE');
        if (SQL = 'FF') then
          TF.SetDouble('MONTANT', TF.GetDouble('MONTANT') + QQ.GetDouble('MONTANT') / 1000)
        else if (SQL = 'AF') then
          TF.SetDouble('MONTANT', TF.GetDouble('MONTANT') - QQ.GetDouble('MONTANT') / 1000)
        else if (SQL = 'AC') then
          TC.SetDouble('MONTANT', TC.GetDouble('MONTANT') - QQ.GetDouble('MONTANT') / 1000)
        else if (SQL = 'FC') then
          TC.SetDouble('MONTANT', TC.GetDouble('MONTANT') + QQ.GetDouble('MONTANT') / 1000);
      end;
    end;
  finally
    FreeAndNil(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEche.TobToGraph(var Obj : TOWCChart);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 0 to TobDonnees.Detail.Count - 1 do
    Obj.SetValue(TobDonnees.Detail[n].GetString('ADATE'),
                 TobDonnees.Detail[n].GetString('TIERS'),
                 TobDonnees.Detail[n].GetDouble('MONTANT'));
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEche.RemplitTobDonneesEtat;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  n   : Integer;
  Q   : TQuery;
  T   : TOB;
begin
  TobDonnees.ClearDetail;
  n := 7;
  if GetString('PERIODE', False) = 'Z' then n := 14
  else if GetString('PERIODE', False) = 'M' then n := 30;

  SQL := 'SELECT ABS(E_CREDIT + E_DEBIT - E_COUVERTURE) E_MONTANT, E_DATEECHEANCE, E_NATUREPIECE, T_LIBELLE, E_LIBELLE, MP_LIBELLE ' +
         ' FROM ECRITURE LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE LEFT JOIN MODEPAIE ON MP_MODEPAIE = E_MODEPAIE ' +
         'WHERE E_DATEECHEANCE BETWEEN "' + UsDateTime(DateRef) + '" AND "' + UsDateTime(DateRef + n) + '" AND ' +
         'E_NATUREPIECE IN ("FF", "AF", "AC", "FC") AND E_QUALIFPIECE = "N" AND E_ETATLETTRAGE IN ("PL", "AL") ' +
         'AND E_TYPEMVT = "TTC" ' +
         ' union ' +
         'SELECT ABS(E_CREDIT + E_DEBIT - E_COUVERTURE) E_MONTANT, E_DATEECHEANCE, E_NATUREPIECE, T_LIBELLE, E_LIBELLE, MP_LIBELLE ' +
         ' FROM ECRITURE LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE LEFT JOIN MODEPAIE ON MP_MODEPAIE = E_MODEPAIE , ' +
         ' JOURNAL, GENERAUX ' +
         ' where G_GENERAL=E_GENERAL ' +
         ' and J_JOURNAL=E_JOURNAL ' +
         ' and J_EFFET="X" ' +
         ' and G_EFFET="X" ' +
         ' and E_QUALIFPIECE = "N" AND E_ETATLETTRAGE IN ("PL", "AL") ' +
         ' ORDER BY E_NATUREPIECE, E_DATEECHEANCE ';

  ddWriteLN('PGICPGRAPHECHEANCE : ' + SQL);
  try
  Q := OpenSelect(SQL);
  except
   on e: exception do
    begin
     ddWriteLN('Erreur : ' + E.Message );
     exit ;
    end ;
  end ;
  try
    while not Q.EOF do begin
      T := TOB.Create('AAA', TobDonnees, -1);
      T.AddChampSupValeur('E_MONTANT', FormatFloat('#,##0.00', Q.FindField('E_MONTANT').AsFloat));
      T.AddChampSupValeur('E_DATEECHEANCE', Q.FindField('E_DATEECHEANCE').AsDateTime);
      T.AddChampSupValeur('E_NATUREPIECE', Q.FindField('E_NATUREPIECE').AsString);
      T.AddChampSupValeur('T_LIBELLE', Q.FindField('T_LIBELLE').AsString);
      T.AddChampSupValeur('E_LIBELLE', Q.FindField('E_LIBELLE').AsString);
      T.AddChampSupValeur('MP_LIBELLE', Q.FindField('MP_LIBELLE').AsString);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
  ddWriteLN('PGICPGRAPHECHEANCE : ' + IntToStr(TobDonnees.Detail.Count));
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEche.TraiteParam(Param: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  EtatOk := Pos('ETAT', Param) > 0;
  ddwriteln('PGICPGRAPHECHEANCE : TraiteParam = ' + FParam);
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEche.ColorSpecific(oChart: chChart);
{---------------------------------------------------------------------------------------}
begin
  if TobDonnees.Detail.Count > 0 then begin
    if TobDonnees.Detail[0].GetString('TIERS') = 'Clients' then begin
      oChart.SeriesCollection[0].Interior.Set_Color(TOWCChart.GetCouleurGraph(ColMediumSeaGreen));
      oChart.SeriesCollection[1].Interior.Set_Color('Red');
    end else begin
      oChart.SeriesCollection[0].Interior.Set_Color('Red');
      oChart.SeriesCollection[1].Interior.Set_Color(TOWCChart.GetCouleurGraph(ColMediumSeaGreen));
    end;
  end;

end;

{---------------------------------------------------------------------------------------}
function TObjGraphEche.InitPeriode : Char;
{---------------------------------------------------------------------------------------}
begin
  Result := 'N';
end;

end.



