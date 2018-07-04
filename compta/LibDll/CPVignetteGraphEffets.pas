{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  10/04/06   JP   Création de l'unité : Répartition de la saisie par type d'écritures
--------------------------------------------------------------------------------------}
unit CPVignetteGraphEffets;

interface

uses
  Classes, UTob,
  {$IFNDEF DBXPRESS}
  dbtables,
  {$ELSE}
  uDbxDataSet,
  {$ENDIF}
  uCPVignettePlugIn, uToolsOWC, HCtrls, OWC11_TLB;

type
  TObjGraphEffets = class(TCPVignettePlugIn)
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
  end;

implementation

uses
  SysUtils, HEnt1, CalcOLE, UProcGen, eSession;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEffets.RecupDonnees;
{---------------------------------------------------------------------------------------}
var
  Obj : TOWCChart;
begin
  inherited;
  if EtatOk then begin
    ddwriteln('PGICPGRAPHEFFETS : RecupDonnees = ETAT');
    RemplitTobDonneesEtat;
    FTip := 'E';
    FNat := 'VCP';
    FModele := 'VCF';
    FLanguePrinc := 'FRA';
    Rapport;
  end
  else begin
    ddwriteln('PGICPGRAPHEFFETS : RecupDonnees = CHART');
    RemplitTobDonnees;
    if TobDonnees.Detail.Count > 0 then begin
      Obj := TOWCChart.Create(TobResponse, AspectGraph);
      try
        try
          ddWriteLN('PGICPGRAPHEFFETS : Debut du chargement des écritures');
          TobToGraph(Obj);
          Obj.AddSpecific := ColorSpecific;
          Obj.owIsAxeRightDegroupe := False;
          Obj.owHasTwoCharts := False;

          Obj.TypeGraph := tyHistogramme;
          Obj.owTitre.Visible := False;

          Obj.owLegend.Visible := False;
          Obj.owAxeBot.Visible := False;
          Obj.owAxeRight.Visible := False;
          Obj.owAxeLeft.Visible := True;
          Obj.owAxeLeft.InfoAxe.NumFormat := '#,##0' +  ' K' + GetDevisePivot(True);
  //        Obj.owAxeLeft.Bold := True;
//          Obj.owAxeLeft.Caption := 'Montant';

          Obj.owLabel.Visible := False;

          Obj.Build;
          ddWriteLN('PGICPGRAPHEFFETS : Fin du chargement des écritures');
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
function TObjGraphEffets.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetInterface;
end;

{---------------------------------------------------------------------------------------}
constructor TObjGraphEffets.Create(TobIn, TobOut: TOB);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(TobIn, TobOut);
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEffets.RemplitTobDonnees;
{---------------------------------------------------------------------------------------}
var
  TC  : TOB;
  Per : Char;
  Q   : TOB;
  QQ  : TOB;
  SQL : string;
  Chp : string;
  dd  : TDateTime;
  a, m, j : Word;
  n : Integer;
begin
  SansInterface := True;
  
  ddWriteLN('PGICPGRAPHEFFETS : Début des GetCumuls');

  TobDonnees.ClearDetail;

  if IsValidDate(GetString('EDDATE', False)) then
    dd := StrToDate(GetString('EDDATE', False))
  else begin
    dd := DateRef;
    SetControlValue('EDDATE', DateToStr(DateRef));
  end;

  Per := StrToChr(GetString('CBAFFICHAGE', False));
  if Per = '' then begin
    Per := 'B';
    SetControlValue('CBAFFICHAGE', 'B');
  end;

  case Per of
    'B' : Chp := 'BQ_LIBELLE';
    'M' : Chp := 'E_DATEECHEANCE';
  end;

  SQL := 'SELECT ABS(SUM(E_DEBIT - E_CREDIT)) MONTANT, ' + Chp + ' FROM ECRITURE ' +
         'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
         'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
         'LEFT JOIN BANQUECP ON BQ_GENERAL = E_BANQUEPREVI ' +
         'AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ' + // 24/10/2006 YMO Multisociétés
         'LEFT JOIN MODEPAIE ON MP_MODEPAIE = E_MODEPAIE ' +
         'WHERE G_EFFET = "X" AND G_SUIVITRESO = "ENC" AND J_EFFET = "X" AND E_ETATLETTRAGE = "TL" ' +
         'AND E_DATEECHEANCE > "' + UsDateTime(dd) + '" AND E_DATEPAQUETMAX <= "' + UsDateTime(dd) + '" ' +
         'AND MP_CATEGORIE = "LCR" GROUP BY ' + Chp + ' ORDER BY ' + Chp;

  ddWriteLN('PGICPGRAPHEFFETS : ' + SQL);

  Q := OpenSelectInCache(SQL);
  try
    for n := 0 to Q.Detail.Count - 1 do begin
      QQ := Q.Detail[n];
      TC := TOB.Create('ùùùù', TobDonnees, -1);
      if Per = 'B' then begin
        SQL := AnsiUpperCase(StrLeft(QQ.GetString('BQ_LIBELLE'), 15));
        if SQL = '' then SQL := 'NON RENSEIGNÉE';
        TC.AddChampSupValeur('TYPE', SQL);
      end else begin
        DecodeDate(QQ.GetDateTime('E_DATEECHEANCE'), a, m, j);
        SQL := 'Mois de ' + PadL(IntToStr(m), '0', 2) + '/' + Copy(IntToStr(a), 3, 2);
        TC.AddChampSupValeur('TYPE', SQL);
      end;
      TC.AddChampSupValeur('MONTANT', QQ.GetDouble('MONTANT') / 1000);
    end;
  finally
    FreeAndNil(Q);
  end;
  ddWriteLN('PGICPGRAPHEFFETS : Charge écritures');
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEffets.TobToGraph(var Obj : TOWCChart);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 0 to TobDonnees.Detail.Count - 1 do
    Obj.SetValue(TobDonnees.Detail[n].GetString('TYPE'),
                 'Effets',
                 TobDonnees.Detail[n].GetDouble('MONTANT'));
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEffets.TraiteParam(Param: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  EtatOk := Pos('ETAT', Param) > 0;
  ddwriteln('PGICPGRAPHEFFETS : TraiteParam = ' + FParam);
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEffets.RemplitTobDonneesEtat;
{---------------------------------------------------------------------------------------}
var
  Per : Char;
  Q   : TQuery;
  SQL : string;
  Chp : string;
  dd  : TDateTime;
  T   : TOB;
begin
  ddWriteLN('PGICPGRAPHEFFETS : Requête état');

  TobDonnees.ClearDetail;

  if IsValidDate(GetString('EDDATE', False)) then
    dd := StrToDate(GetString('EDDATE', False))
  else begin
    dd := DateRef;
    SetControlValue('EDDATE', DateToStr(DateRef));
  end;

  Per := StrToChr(GetString('CBAFFICHAGE', False));
  if Per = '' then begin
    Per := 'B';
    SetControlValue('CBAFFICHAGE', 'B');
  end;

  case Per of
    'B' : Chp := 'E_BANQUEPREVI E_RUPTURE, "B" E_TYPE, ';
    'M' : Chp := 'E_DATEECHEANCE E_RUPTURE, "M" E_TYPE,'; //(STR(MONTH(E_DATEECHEANCE)) + STR(YEAR(E_DATEECHEANCE)))
  end;

  SQL := 'SELECT (E_DEBIT - E_CREDIT) E_MONTANT, G_LIBELLE, T_LIBELLE, E_LIBELLE, BQ_LIBELLE, ' + Chp +
         'E_DATEECHEANCE, E_MODEPAIE FROM ECRITURE ' +
         'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
         'LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
         'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
         'LEFT JOIN BANQUECP ON BQ_GENERAL = E_BANQUEPREVI ' +
         'AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ' + // 24/10/2006 YMO Multisociétés
         'LEFT JOIN MODEPAIE ON MP_MODEPAIE = E_MODEPAIE ' +
         'WHERE G_EFFET = "X" AND G_SUIVITRESO = "ENC" AND J_EFFET = "X" AND E_ETATLETTRAGE = "TL" ' +
         'AND E_DATEECHEANCE > "' + UsDateTime(dd) + '" AND E_DATEPAQUETMAX <= "' + UsDateTime(dd) + '" ' +
         'AND MP_CATEGORIE = "LCR" ORDER BY E_RUPTURE';

  ddWriteLN('PGICPGRAPHEFFETS : ' + SQL);
  Q := OpenSelect(SQL);
  try
    while (not Q.EOF) do begin
      T := TOB.Create('§§§', TobDonnees, -1);;
      T.AddChampSupValeur('G_LIBELLE'     , Q.FindField('G_LIBELLE').AsString);
      T.AddChampSupValeur('T_LIBELLE'     , Q.FindField('T_LIBELLE').AsString);
      T.AddChampSupValeur('E_LIBELLE'     , Q.FindField('E_LIBELLE').AsString);
      T.AddChampSupValeur('BQ_LIBELLE'    , Q.FindField('BQ_LIBELLE').AsString);
      T.AddChampSupValeur('E_RUPTURE'     , Q.FindField('E_RUPTURE').AsString);
      T.AddChampSupValeur('E_TYPE'        , Q.FindField('E_TYPE').AsString);
      T.AddChampSupValeur('E_DATEECHEANCE', Q.FindField('E_DATEECHEANCE').AsDateTime);
      T.AddChampSupValeur('E_MODEPAIE'    , Q.FindField('E_MODEPAIE').AsString);
      T.AddChampSupValeur('E_MONTANT'     , FormatFloat('#,##0.00', Q.FindField('E_MONTANT').AsFloat));
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
  ddWriteLN('PGICPGRAPHEFFETS : ' + IntToStr(TobDonnees.Detail.Count) + ' enregistrements');
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphEffets.ColorSpecific(oChart : chChart);
{---------------------------------------------------------------------------------------}
begin
  oChart.SeriesCollection[0].Interior.Set_Color(TOWCChart.GetCouleurGraph(ColOrange));
end;

end.

