{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  28/03/06   JP   Création de l'unité : Vignettes des impayés
--------------------------------------------------------------------------------------}
unit CPVignetteGraphImpayes;

interface

uses
  Classes, UTob,
  {$IFNDEF DBXPRESS}
  dbtables,
  {$ELSE}
  uDbxDataSet,
  {$ENDIF}
  uCPVignettePlugIn, uToolsOWC, HCtrls;

type
  TObjGraphImpaye = class(TCPVignettePlugIn)
  private
    EtatOk : Boolean;

    procedure RemplitTobDonnees;
    procedure RemplitTobDonneesEtat;
    procedure TobToGraph(var Obj : TOWCChart);
  protected
    procedure RecupDonnees ;             override;
    function  GetInterface(NomGrille : string = '') : Boolean; override;
    procedure TraiteParam(Param : string); override;
  public
    constructor Create(TobIn, TobOut : TOB); override;
  end;

implementation

uses
  SysUtils, HEnt1, Ent1, ParamSoc, UProcGen;

{---------------------------------------------------------------------------------------}
procedure TObjGraphImpaye.RecupDonnees;
{---------------------------------------------------------------------------------------}
var
  Obj : TOWCChart;
begin
  inherited;
  if EtatOk then begin
    RemplitTobDonneesEtat;
    FTip := 'E';
    FNat := 'VCP';
    FModele := 'VCI';
    FLanguePrinc := 'FRA';
    Rapport;
  end
  else begin
    RemplitTobDonnees;
    if TobDonnees.Detail.Count > 0 then begin
      {Il faut laisser se graph en 2D, il passe mal en 3D}
      AspectGraph := '2D';
      Obj := TOWCChart.Create(TobResponse, AspectGraph);
      try
        try
          ddWriteLN('PGICPGRAPHIMPAYE : Debut du chargement des écritures');
          TobToGraph(Obj);
          Obj.TypeGraph := tyHistogramme;
          Obj.owTitre.Visible := False;
          {Obj.owTitre.Caption := 'Impayés clients';
          Obj.owTitre.Bold := True;
          Obj.owTitre.FontSize := 14;}

          Obj.owLegend.Visible := True;
          Obj.owLegend.Position := poPositionBottom;

          Obj.owAxeLeft.Caption := 'Montant';
          Obj.owAxeLeft.InfoAxe.NumFormat := '#,##0' + ' K' + GetDevisePivot(True);
          Obj.owAxeLeft.Bold := True;
          Obj.owAxeLeft.InfoAxe.TypeSerie := tyHistogramme;
          Obj.owAxeLeft.InfoAxe.HasLabel := False;

          Obj.owAxeRight.Bold := True;
          Obj.owAxeRight.Visible := True;
          Obj.owAxeRight.Caption := 'Nombre';
          Obj.owAxeRight.InfoAxe.NumFormat := '#,##0';
          Obj.owAxeRight.InfoAxe.TypeSerie := tyLigne;
          Obj.owAxeRight.InfoAxe.HasLabel := True;

          Obj.owLabel.Visible := True;
          Obj.owLabel.Color   := colLightYellow;
          Obj.owIsAxeRightDegroupe := True;
          Obj.Build;
          ddWriteLN('PGICPGRAPHIMPAYE : Fin du chargement des écritures');
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
function TObjGraphImpaye.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetInterface;
end;

{---------------------------------------------------------------------------------------}
constructor TObjGraphImpaye.Create(TobIn, TobOut: TOB);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(TobIn, TobOut);
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphImpaye.RemplitTobDonnees;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Per : string;
  Q   : TOB;
  QQ  : TOB;
  TF  : TOB;
  TC  : TOB;
  dd  : TDateTime;
  Chp : string;
  n   : Integer;
  P   : Integer;
  y   : Word;
  m   : Word;
  j   : Word;
  f   : Double;
begin
  SansInterface := True;

  TobDonnees.ClearDetail;
  Per := Periode;

  if Per = 'M' then Chp := 'E_PERIODE'
               else Chp := 'E_PERIODE, E_SEMAINE';

  {On travaille sur 6 barres}
  DecodeDate(DateRef - 35, y, m, j);
  dd := PremierJourSemaine(NumSemaine(DateRef - 35), y);
  if Per = 'M' then
    dd := DebutDeMois(DateRef - 155);

  SQL := 'SELECT ABS(SUM(E_CREDIT - E_DEBIT)) MONTANT, COUNT(*) NOMBRE, ' + Chp + ' FROM ECRITURE ' +
         'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
         'WHERE E_DATECOMPTABLE BETWEEN "' + UsDateTime(dd) + '" AND "' + UsDateTime(DateRef) + '" AND ' +
         'G_NATUREGENE = "COC" AND E_MODEPAIE = "' + GetParamSocSecur('SO_CPMODEPAIEIMPAYE', '') + '"' ;
  if GetString('CBCOMPTE', False) <> '' then
    SQL := SQL + ' AND E_CONTREPARTIEGEN = "' + GetString('CBCOMPTE', False) + '" GROUP BY ' + Chp + ' ORDER BY ' + Chp
  else
    SQL := SQL + ' GROUP BY ' + Chp + ' ORDER BY ' + Chp;

  ddWriteLN('PGICPGRAPHIMPAYE : ' + SQL);
  Q := OpenSelectInCache(SQL);
  try
    for p := 0 to Q.Detail.Count - 1 do begin
      QQ := Q.Detail[p];
      TC := TOB.Create('ùùùù', TobDonnees, -1);
      TF := TOB.Create('ùùùù', TobDonnees, -1);
      m := StrToInt(Copy(QQ.GetString('E_PERIODE'), 5, 2));
      y := StrToInt(Copy(QQ.GetString('E_PERIODE'), 1, 4));
      if Per = 'M' then begin
        Chp := MajLibPeriode(EncodeDate(Y, m, 1));
        f := QQ.GetDouble('MONTANT');
        n := QQ.GetInteger('NOMBRE');
      end
      else begin
        dd := PremierJourSemaine(QQ.GetInteger('E_SEMAINE'), y);
        DecodeDate(dd, y, m, j);
        DateTimeToString(SQL, 'dd/mm/yy', dd + 7);
        Chp := MajLibPeriode(dd);
        f := QQ.GetDouble('MONTANT');
        n := QQ.GetInteger('NOMBRE');
      end;

      TC.AddChampSupValeur('PERIODE', Chp);
      TF.AddChampSupValeur('PERIODE', Chp);
      TC.AddChampSupValeur('DONNEES', 'Montant');
      TF.AddChampSupValeur('DONNEES', 'Nombre');
      TC.AddChampSupValeur('MONTANT', Format('%.2n', [f / 1000]));
      TF.AddChampSupValeur('MONTANT', n); //Format('%n', [n])
    end;
  finally
    FreeAndNil(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphImpaye.TobToGraph(var Obj : TOWCChart);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 0 to TobDonnees.Detail.Count - 1 do
    Obj.SetValue(TobDonnees.Detail[n].GetString('PERIODE'),
                 TobDonnees.Detail[n].GetString('DONNEES'),
                 TobDonnees.Detail[n].GetDouble('MONTANT'));
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphImpaye.RemplitTobDonneesEtat;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Chp : string;
  Deb : TDateTime;
  a   : Word;
  m   : Word;
  j   : Word;
  Q   : TQuery;
  T   : TOB;
begin
  TobDonnees.ClearDetail;
  if Periode = 'M' then Chp := 'E_PERIODE'
                   else Chp := 'E_PERIODE, E_SEMAINE';

  {On travaille sur 6 barres}
  DecodeDate(DateRef - 35, a, m, j);
  Deb := PremierJourSemaine(NumSemaine(DateRef - 35), a);
  if Periode = 'M' then
    Deb := DebutDeMois(DateRef - 155);

  SQL := 'SELECT ABS(E_CREDIT - E_DEBIT - E_COUVERTURE) E_MONTANT, E_DATECOMPTABLE, E_NUMEROPIECE, E_LIBELLE, ' +
         'E_REFINTERNE, E_MODEPAIE, E_CONTREPARTIEGEN, BQ_LIBELLE, T_LIBELLE FROM ECRITURE ' +
         'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
         'LEFT JOIN BANQUECP ON BQ_GENERAL = E_CONTREPARTIEGEN ' +
         'AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ' + // 24/10/2006 YMO Multisociétés
         'LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
         'WHERE E_DATECOMPTABLE BETWEEN "' + UsDateTime(Deb) + '" AND "' + UsDateTime(DateRef) + '" AND ' +
         'G_NATUREGENE = "COC" AND E_MODEPAIE = "' + GetParamSocSecur('SO_CPMODEPAIEIMPAYE', '') + '"' ;
  if GetString('CBCOMPTE', False) <> '' then
    SQL := SQL + ' AND E_CONTREPARTIEGEN = "' + GetString('CBCOMPTE', False) + '" ';
  SQL := SQL + 'ORDER BY E_CONTREPARTIEGEN, E_DATECOMPTABLE';

  ddWriteLN('PGICPGRAPHIMPAYE : ' + SQL);
  Q := OpenSelect(SQL);
  try
    while (not Q.EOF) do begin
      T := TOB.Create('§§§', TobDonnees, -1);;
      T.AddChampSupValeur('E_LIBELLE'        , Q.FindField('E_LIBELLE').AsString);
      T.AddChampSupValeur('E_REFINTERNE'     , Q.FindField('E_REFINTERNE').AsString);
      T.AddChampSupValeur('E_MODEPAIE'       , Q.FindField('E_MODEPAIE').AsString);
      T.AddChampSupValeur('E_CONTREPARTIEGEN', Q.FindField('E_CONTREPARTIEGEN').AsString);
      T.AddChampSupValeur('BQ_LIBELLE'       , Q.FindField('BQ_LIBELLE').AsString);
      T.AddChampSupValeur('T_LIBELLE'        , Q.FindField('T_LIBELLE').AsString);
      T.AddChampSupValeur('E_DATECOMPTABLE'  , Q.FindField('E_DATECOMPTABLE').AsDateTime);
      T.AddChampSupValeur('E_NUMEROPIECE'    , Q.FindField('E_NUMEROPIECE').AsString);
      T.AddChampSupValeur('E_MONTANT'        , FormatFloat('#,##0.00', Q.FindField('E_MONTANT').AsFloat));
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
  ddWriteLN('PGICPGRAPHIMPAYE : ' + IntToStr(TobDonnees.Detail.Count));
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphImpaye.TraiteParam(Param: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  EtatOk := Pos('ETAT', Param) > 0;
end;

end.

