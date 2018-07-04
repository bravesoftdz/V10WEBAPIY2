{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  12/04/06   JP   Création de l'unité : Répartition de la saisie par nature de journaux
--------------------------------------------------------------------------------------}
unit CPVignetteGraphVolumetrie;

interface

uses
  Classes, UTob, DBTables, uCPVignettePlugIn, uToolsOWC, HCtrls, OWC11_TLB;

type
  TObjGraphVolume = class(TCPVignettePlugIn)
  private
    DeuxSeries : Boolean;

    procedure CompleteSerie(Lse : TStringList; S1, S2 : string);
    procedure ColorSpecific(oChart : chChart);
    procedure RemplitTobDonnees;
    procedure TobToGraph(var Obj : TOWCChart);
  protected
    procedure RecupDonnees ;             override;
    function  GetInterface(NomGrille : string = '') : Boolean; override;
  public
    constructor Create(TobIn, TobOut : TOB); override;
  end;

implementation

uses
  SysUtils, HEnt1, CalcOLE, UProcGen, NumConv, eSession;

{---------------------------------------------------------------------------------------}
procedure TObjGraphVolume.RecupDonnees;
{---------------------------------------------------------------------------------------}
var
  Obj : TOWCChart;
begin
  inherited;
  SansInterface := True;
  
  ddwriteln('PGICPGRAPHVOLUME : RecupDonnees = CHART');
  RemplitTobDonnees;
  if TobDonnees.Detail.Count > 0 then begin
    Obj := TOWCChart.Create(TobResponse, AspectGraph);
    try
      try
        ddWriteLN('PGICPGRAPHVOLUME : Debut du chargement des écritures');
        TobToGraph(Obj);
        Obj.AddSpecific := ColorSpecific;
        Obj.owIsAxeRightDegroupe := False;
        Obj.owHasTwoCharts := False;

        Obj.TypeGraph := tyHistogramme;
        Obj.owTitre.Visible := False;

        Obj.owLegend.Visible := DeuxSeries;
        Obj.owLegend.Position := poPositionBottom;
        Obj.owAxeBot.Visible := False;
        Obj.owAxeRight.Visible := False;
        Obj.owAxeLeft.Visible := True;
        Obj.owAxeLeft.Bold := True;
        Obj.owAxeLeft.Caption := 'Nombre d''écritures';
        Obj.owAxeLeft.InfoAxe.NumFormat := '#,##0';
        Obj.owAxeLeft.InfoAxe.HasLabel := True;
        Obj.owLabel.Visible := True;
        Obj.owLabel.InfoAxe.LabelValue := True;
        Obj.owLabel.InfoAxe.LabelPerct := False;
        Obj.owLabel.InfoAxe.NumFormat := '#,##0';
        Obj.owLabel.FontSize := 6;
        Obj.owLabel.Color := ColPapayaWhip;

        Obj.Build;
        ddWriteLN('PGICPGRAPHVOLUME : Fin du chargement des écritures');
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
function TObjGraphVolume.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetInterface;
end;

{---------------------------------------------------------------------------------------}
constructor TObjGraphVolume.Create(TobIn, TobOut: TOB);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(TobIn, TobOut);
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphVolume.RemplitTobDonnees;
{---------------------------------------------------------------------------------------}
var
  TC  : TOB;
  Per : Char;
  Com : Char;
  SQL : string;
  dd  : string;
  S1  : string;
  S2  : string;
  d1  : TDateTime;
  d2  : TDateTime;
  d3  : TDateTime;
  Lse : TStringList;

    {-----------------------------------------------------------------------}
    procedure RemplitTob(Serie : string);
    {-----------------------------------------------------------------------}
    var
      Q  : TOB;
      QQ : TOB;
      n  : Integer;
    begin
      Q := OpenSelectInCache(SQL);
      try
        for n := 0 to Q.Detail.Count - 1 do begin
          QQ := Q.Detail[n];
          TC := TOB.Create('ùùùù', TobDonnees, -1);
          TC.AddChampSupValeur('TYPE', Serie);
          TC.AddChampSupValeur('NATJAL', QQ.GetString('CO_LIBELLE'));
          TC.AddChampSupValeur('NOMBRE', QQ.GetInteger('NOMBRE'));
          Lse.Add(QQ.GetString('CO_LIBELLE'));
        end;
      finally
        FreeAndNil(Q);
      end;
    end;

    {-----------------------------------------------------------------------}
    function ConstruitRequete : string;
    {-----------------------------------------------------------------------}
    begin
      Result := 'SELECT COUNT(*) NOMBRE, J_NATUREJAL, CO_LIBELLE FROM ECRITURE ' +
         'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
         'LEFT JOIN COMMUN ON CO_CODE = J_NATUREJAL AND CO_TYPE="NTJ" ' +
         'WHERE E_QUALIFPIECE = "N" ' + dd +
         'GROUP BY J_NATUREJAL, CO_LIBELLE ORDER BY CO_LIBELLE';
    end;
begin
  SansInterface := True;

  TobDonnees.ClearDetail;

  Lse := TStringList.Create;
  Lse.Sorted := True;
  Lse.Duplicates := dupIgnore;
  try
    ddWriteLN('PGICPGRAPHVOLUME : Début des GetCumuls');
    Per := Periode;

    Com := StrToChr(GetString('CBCOMPARATIF', False));
    if Com = '' then begin
      Com := 'A';
      SetControlValue('CBCOMPARATIF', 'A');
    end;

    DeuxSeries := Com <> 'A';

    case Per of
      'M' : begin
              dd := 'AND E_DATECOMPTABLE BETWEEN "' + UsDateTime(DebutDeMois(DateRef)) + '" AND "' + UsDateTime(FinDeMois(DateRef)) + '" ';
              S1 := DateToMois(DateRef);
            end;
      'T' : begin
              dd := 'AND E_DATECOMPTABLE BETWEEN "' + UsDateTime(StrToDate(DebutTrimestre(DateRef))) + '" AND "' + UsDateTime(StrToDate(FinTrimestre(DateRef))) + '" ';
              S1 := DateToTrimestre(DateRef);
            end;
      'A' : begin
              dd := 'AND E_DATECOMPTABLE BETWEEN "' + UsDateTime(DebutAnnee(DateRef)) + '" AND "' + UsDateTime(FinAnnee(DateRef)) + '" ';
              S1 := IntToStr(RetourneAnneeW(DateRef));
            end;
    end;

    SQL := ConstruitRequete;
    RemplitTob(S1);
    S2 := S1;

    if DeuxSeries then begin
      d1 := DebutDeMois(DebutDeMois(DateRef) - 1);
      d2 := FinDeMois  (DebutDeMois(DateRef) - 1);

      case Com of
        'P' : case Per of
                'M' : begin
                        d1 := DebutDeMois(DebutDeMois(DateRef) - 1);
                        d2 := FinDeMois  (DebutDeMois(DateRef) - 1);
                        S1 := DateToMois(d1);
                      end;
                'T' : begin
                        d1 := DebutDeTrimestre(DebutDeTrimestre(DateRef) - 1);
                        d2 := FinDeTrimestre  (DebutDeTrimestre(DateRef) - 1);
                        S1 := DateToTrimestre(d1);
                      end;
                'A' : begin
                        d1 := DebutAnnee(DebutAnnee(DateRef) - 1);
                        d2 := FinAnnee  (DebutAnnee(DateRef) - 1);
                        S1 := IntToStr(RetourneAnneeW(d1));
                      end;
              end;

        'N' : begin
                if (IsLeapYear(RetourneAnneeW(DateRef)) and
                    (DateRef >= StrToDate('01/03/' + IntToStr(RetourneAnneeW(DateRef))))) or
                   (IsLeapYear(RetourneAnneeW(DateRef - 365)) and
                    (DateRef < StrToDate('01/03/' + IntToStr(RetourneAnneeW(DateRef))))) then
                  d3 := DateRef - 366
                else
                  d3 := DateRef - 365;

                case Per of
                  'M' : begin
                          d1 := DebutDeMois(d3);
                          d2 := FinDeMois  (d3);
                          S1 := DateToMois(d1);
                        end;
                  'T' : begin
                          d1 := DebutDeTrimestre(d3);
                          d2 := FinDeTrimestre  (d3);
                          S1 := DateToTrimestre(d1);
                        end;
                  'A' : begin
                          d1 := DebutAnnee(d3);
                          d2 := FinAnnee  (d3);
                          S1 := IntToStr(RetourneAnneeW(d1));
                        end;
                end;
              end;
      end;
      dd := 'AND E_DATECOMPTABLE BETWEEN "' + UsDateTime(d1) + '" AND "' + UsDateTime(d2) + '" ';

      SQL := ConstruitRequete;
      RemplitTob(S1);
    end;
    {Du fait des deux requêtes, les 2 séries n'ont pas frocément le même nombre de catégories :
     Cela pose un problème à l'objet Chart si l'une des deux séries n'a pas de données dans la dernière
     catégorie ; c'est celle-ci qu'utilise l'objet pour constituer la légende : comme il ne trouve pas
     l'une des séries dans la dernière catégorie, il reprend le titre de la série présente}
    CompleteSerie(Lse, S1, S2);
  finally
    FreeAndNil(Lse);
  end;
  ddWriteLN('PGICPGRAPHVOLUME : Charge écritures');
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphVolume.TobToGraph(var Obj : TOWCChart);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 0 to TobDonnees.Detail.Count - 1 do
    Obj.SetValue(TobDonnees.Detail[n].GetString('NATJAL'),
                 TobDonnees.Detail[n].GetString('TYPE'),
                 TobDonnees.Detail[n].GetInteger('NOMBRE'));
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphVolume.ColorSpecific(oChart : chChart);
{---------------------------------------------------------------------------------------}
begin
  oChart.SeriesCollection[0].Interior.Set_Color(TOWCChart.GetCouleurGraph(ColPeachPuff));
  if DeuxSeries then
    oChart.SeriesCollection[1].Interior.Set_Color(TOWCChart.GetCouleurGraph(ColPapayaWhip));
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphVolume.CompleteSerie(Lse : TStringList; S1, S2 : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin
  for n := 0 to Lse.Count - 1 do begin
    T := TobDonnees.FindFirst(['TYPE', 'NATJAL'], [S1, Lse[n]], True);
    if T = nil then begin
      T := TOB.Create('ùùùù', TobDonnees, -1);
      T.AddChampSupValeur('TYPE', S1);
      T.AddChampSupValeur('NATJAL', Lse[n]);
      T.AddChampSupValeur('NOMBRE', 0);
    end;

    T := TobDonnees.FindFirst(['TYPE', 'NATJAL'], [S2, Lse[n]], True);
    if T = nil then begin
      T := TOB.Create('ùùùù', TobDonnees, -1);
      T.AddChampSupValeur('TYPE', S2);
      T.AddChampSupValeur('NATJAL', Lse[n]);
      T.AddChampSupValeur('NOMBRE', 0);
    end;
  end;
end;

end.

