{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  30/03/06   JP   Création de l'unité : Marge brute et valeur ajoutée
--------------------------------------------------------------------------------------}
unit CPVIGNETTEGRAPHMBVA;

interface

uses
  Classes, UTob,
  {$IFNDEF DBXPRESS}
  dbtables,
  {$ELSE}
  uDbxDataSet,
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF VER150}
  uCPVignettePlugIn, uToolsOWC, HCtrls, OWC11_TLB;

type
  TObjGraphMBVA = class(TCPVignettePlugIn)
  private
    procedure RemplitTobDonnees;
    procedure TobToGraph(var Obj : TOWCChart);
    function  CalcPeriode(Ind : Byte; Periode : Char; NomChpOk : Boolean) : string;
    procedure ColorSpecific(oChart : chChart);
  protected
    procedure RecupDonnees ;             override;
    function  GetInterface(NomGrille : string = '') : Boolean; override;
  public
    constructor Create(TobIn, TobOut : TOB); override;
    function    InitPeriode : Char; override;
  end;

implementation

uses
  SysUtils, HEnt1, CalcOLE, UProcGen, NumConv, eSession;

{---------------------------------------------------------------------------------------}
procedure TObjGraphMBVA.RecupDonnees;
{---------------------------------------------------------------------------------------}
var
  Obj : TOWCChart;
begin
  inherited;
  ddwriteln('PGICPGRAPHMBVA : RecupDonnees = CHART');
  RemplitTobDonnees;
  if TobDonnees.Detail.Count > 0 then begin
    Obj := TOWCChart.Create(TobResponse, AspectGraph);
    try
      try
        ddWriteLN('PGICPGRAPHMBVA : Debut du chargement des écritures');
        TobToGraph(Obj);
        Obj.AddSpecific := ColorSpecific;
        Obj.TypeGraph := tyHistogramme;
        Obj.owTitre.Visible := False;
        {Obj.owTitre.Caption := 'Marge Brute et Valeur Ajoutée';
        Obj.owTitre.Bold := True;
        Obj.owTitre.FontSize := 14;}

        Obj.owLegend.Visible := True;
        Obj.owLegend.Position := poPositionBottom;

//        Obj.owAxeLeft.Caption := 'Montant';
  //      Obj.owAxeLeft.Bold := True;
        Obj.owAxeLeft.InfoAxe.NumFormat := '#,##0' + ' K' + GetDevisePivot(True);
        Obj.Build;
        ddWriteLN('PGICPGRAPHMBVA : Fin du chargement des écritures');
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
function TObjGraphMBVA.GetInterface(NomGrille : string = '') : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetInterface;
end;

{---------------------------------------------------------------------------------------}
constructor TObjGraphMBVA.Create(TobIn, TobOut: TOB);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(TobIn, TobOut);
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphMBVA.RemplitTobDonnees;
{---------------------------------------------------------------------------------------}
var
  TF  : TOB;
  TC  : TOB;
  Eta : string;
  Dat : string;
  Dos : string;
  Dev : string;
  dd  : string;
  VA  : Double;
  MB  : Double;
  n   : Byte;
begin
  SansInterface := True;

  TobDonnees.ClearDetail;

  ddWriteLN('PGICPGRAPHMBVA : Début des GetCumuls');
  Eta := GetString('CBETAB', False);
  Dos := GetString('MULTIDOSSIERPORTAIL');

  Dev := GetDevisePivot(False);

  for n := 4 downto 1 do begin
    VA := 0;
    MB := 0;

    Dat := CalcPeriode(n, Periode, False);
    dd  := CalcPeriode(n, Periode, True);
    ddWriteLN('PGICPGRAPHMBVA : Valeur ajoutée n°' + IntToStr(n));
    VA := VA + VarToDle(Get_Cumul2MS('RUBRIQUE', '@CRP/000', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    VA := VA + VarToDle(Get_Cumul2MS('RUBRIQUE', '@CRP/001', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    VA := VA + VarToDle(Get_Cumul2MS('RUBRIQUE', '@CRP/002', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    VA := VA + VarToDle(Get_Cumul2MS('RUBRIQUE', '@SIG/003', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    VA := VA - VarToDle(Get_Cumul2MS('RUBRIQUE', '@SIG/010', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    VA := VA - VarToDle(Get_Cumul2MS('RUBRIQUE', '@SIG/011', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    VA := VA - VarToDle(Get_Cumul2MS('RUBRIQUE', '@SIG/012', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));

    ddWriteLN('PGICPGRAPHMBVA : Marge brute n°' + IntToStr(n));
    MB := MB + VarToDle(Get_Cumul2MS('RUBRIQUE', '@CRP/000', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    MB := MB + VarToDle(Get_Cumul2MS('RUBRIQUE', '@CRP/001', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    MB := MB + VarToDle(Get_Cumul2MS('RUBRIQUE', '@CRP/002', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    MB := MB + VarToDle(Get_Cumul2MS('RUBRIQUE', '@SIG/003', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    MB := MB - VarToDle(Get_Cumul2MS('RUBRIQUE', '@SIG/010', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));
    MB := MB - VarToDle(Get_Cumul2MS('RUBRIQUE', '@SIG/011', 'N', Eta, Dev, Dat, '', '', '', '', '', Null, Dos));

    TC := TOB.Create('ùùùù', TobDonnees, -1);
    TF := TOB.Create('ùùùù', TobDonnees, -1);
    TC.AddChampSupValeur('ADATE', dd);
    TF.AddChampSupValeur('ADATE', dd);
    TC.AddChampSupValeur('TIERS', 'Marge Brute');
    TF.AddChampSupValeur('TIERS', 'Valeur Ajoutée');
    TC.AddChampSupValeur('MONTANT', MB / 1000);
    TF.AddChampSupValeur('MONTANT', VA / 1000);
  end;
  ddWriteLN('PGICPGRAPHMBVA : Fin des GetCumuls');
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphMBVA.TobToGraph(var Obj : TOWCChart);
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
function TObjGraphMBVA.CalcPeriode(Ind : Byte; Periode : Char; NomChpOk : Boolean) : string;
{---------------------------------------------------------------------------------------}

      {---------------------------------------------------------------------}
      function GetADate(aDate : TDateTime; Nb : Byte; DebOk : Boolean) : TDateTime;
      {---------------------------------------------------------------------}
      begin
        Result := aDate;
        if Nb >= Ind then
          case Periode of
            'M' : begin
                    if DebOk then Result := DebutDeMois(aDate)
                             else Result := FinDeMois(aDate);
                  end;
            'T' : begin
                    if DebOk then Result := StrToDate(DebutTrimestre(aDate))
                             else Result := StrToDate(FinTrimestre(aDate));
                  end;
            'S' : begin
                    if DebOk then Result := StrToDate(DebutSemestre(aDate))
                             else Result := StrToDate(FinSemestre(aDate));
                  end;
          end
        else
          case Periode of
            'M' : Result := GetADate(DebutDeMois(aDate)               - 1, Nb + 1, DebOk);
            'S' : Result := GetADate(StrToDate(DebutSemestre(aDate))  - 1, Nb + 1, DebOk);
            'T' : Result := GetADate(StrToDate(DebutTrimestre(aDate)) - 1, Nb + 1, DebOk);
          end;
      end;

begin
  if NomChpOk then
    Result := MajLibPeriode(GetADate(DateRef, 1, True))
    //'Du ' + FormateDate('dd/mm/yy', GetADate(DateRef, 1, True)) + ' au ' + FormateDate('dd/mm/yy', GetADate(DateRef, 1, False))
  else
    Result := '(' + DateToStr(GetADate(DateRef, 1, True)) + ')(' + DateToStr(GetADate(DateRef, 1, False)) + ')';
end;

{---------------------------------------------------------------------------------------}
procedure TObjGraphMBVA.ColorSpecific(oChart : chChart);
{---------------------------------------------------------------------------------------}
begin
  oChart.SeriesCollection[0].Interior.Set_Color(TOWCChart.GetCouleurGraph(ColMoccasin));
  oChart.SeriesCollection[1].Interior.Set_Color(TOWCChart.GetCouleurGraph(ColOrchid));
end;

{---------------------------------------------------------------------------------------}
function TObjGraphMBVA.InitPeriode : Char;
{---------------------------------------------------------------------------------------}
begin
  Result := 'S';
end;

end.                        

