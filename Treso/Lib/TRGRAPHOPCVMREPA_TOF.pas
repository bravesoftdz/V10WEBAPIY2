{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.2X.xxx.xxx    06/11/05   JP   Création de l'unité  : Graph de la fiche de suivi des
                                 portefeuilles, basé sur des camemberts
 8.00.001.018    05/06/07   JP   FQ 10469 : désactvation des filtres et des boutons
--------------------------------------------------------------------------------------}
unit TRGRAPHOPCVMREPA_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Controls, Classes, UTOF,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  Forms, SysUtils, UTob, TeEngine, HCtrls;


type

  TOF_TRGRAPHOPCVMREPA = class (TOF)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
  private
    TobSeries : TOB;
    ListeCourbes1 : string;
    ListeCourbes2 : string;
    IndexMere     : Integer;

    procedure Chart1MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure Chart2MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure SerieChart1GetMarkText(Sender : TChartSeries; ValueIndex : Longint; var MarkText : hString);
    procedure SerieChart2GetMarkText(Sender : TChartSeries; ValueIndex : Longint; var MarkText : hString);
  end ;

procedure TRLanceFiche_GraphOpcvmRepart(Arguments : string);

implementation

uses
  UObjFiltres, HTB97, HEnt1, GRS1, Series, GraphUtil, Graphics, Chart;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_GraphOpcvmRepart(Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('TR', 'TRGRAPHOPCVMREPA', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVMREPA.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  F : TFGRS1;
  stTitre,
  stColonnes, stChampTitre,
  stCriteres, stWhere,
  stTitresCol1, stTitresCol2,
  stColGraph1, stColGraph2 : string;
  tstTitres                : Tstrings;
  ft                       : TFont;
begin
  inherited;

  stCriteres := '';
  stColGraph1 := '';
  stColGraph2 := '';
  stTitresCol1 := '';
  stTitresCol2 := '';
  stTitre := '';

  tstTitres := TStringList.Create;
  ft := TFont.Create;
  try
    {Definition des fonts des titres}
    ft.Style := [fsBold];
    ft.Color := clBlue;
    ft.Size  := 12;

    F := TFGRS1(Ecran);
    TstTitres.Add('Répartition des portefeuilles et des OPCVM');

    stColonnes  := ListeCourbes1;
    stTitre     := ListeCourbes1;
    stColGraph1 := ListeCourbes1;
    {Par défaut on cahche le Graph 2}
    stColGraph2 := '';
    stChampTitre := stTitre;
    LanceGraph(F, TobSeries, '', stColonnes, stWhere , stTitre , stColGraph1,
                 stColGraph2, tstTitres, nil, TPieSeries , stChampTitre, False);

    if F.Fchart1.SeriesCount > 0 then begin
      {Pour un affichage en traits continus sans relief}
      F.FChart1.View3D := False;
      F.FChart2.View3D := False;

      {Titre des légendes}
      F.FChart1.Legend.Visible := False;
      F.FChart2.Legend.Visible := False;
      {Le MouseDown sur les graphiques}
      F.FChart1.OnMouseDown := Chart1MouseDown;
      F.FChart2.OnMouseDown := Chart2MouseDown;
      {Definition du cercle et des marques}
      TPieSeries(F.FChart1.SeriesList.Series[1]).CustomYRadius := Round(TFGRS1(Ecran).FChart1.Height / 2) - 50;
      TPieSeries(F.FChart1.SeriesList.Series[1]).Circled := True;
      F.FChart1.SeriesList.Series[1].Marks.Clip := False;
      F.FChart1.SeriesList.Series[1].OnGetMarkText := SerieChart1GetMarkText;
    end;
  finally
    FreeAndNil(ft);
    tstTitres.Free;
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVMREPA.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  IndexMere := -1;
  SetControlVisible('BUNDO'  , False); {FQ 10128}
  SetControlVisible('BSAUVE' , False); {FQ 10128}
  SetControlVisible('BDELETE', False); {FQ 10128}

  SetControlCaption('TITRE', ReadTokenSt(S));
  {Récupération des colonnes des graphes}
  ListeCourbes1 := ReadTokenSt(S);
  ListeCourbes2 := ReadTokenSt(S);
  {Remplacement des ":" par des ";"}
  ListeCourbes1 := FindEtReplace(ListeCourbes1, ':', ';', True);
  ListeCourbes2 := FindEtReplace(ListeCourbes2, ':', ';', True);
  {Récupération de la Tob}
  TobSeries := TFGRS1(Ecran).LaTOF.LaTOB;

  {05/06/07 : FQ 10469 : on cache les filtres}
  CacheFiltreGraph(Ecran);

  {Lancement du traitement, pour qu'en arrivée sur la fiche, le graphe soit affiché}
  TToolbarButton97(GetControl('BVALIDER')).Click;
  TToolbarButton97(GetControl('BAFFGRAPH')).Click;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVMREPA.Chart1MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : TPieSeries;
  s : TPieSeries;
  T : TOB;
begin
  IndexMere := -1;
  p := TPieSeries(TFGRS1(Ecran).FChart1.SeriesList.Series[1]);
  n := p.CalcClickedPie(x, y);
  if n > -1 then begin
    {Recherche de la fille correspondant à l'endroit où l'on a cliqué}
    T := TobSeries.FindFirst(['PORTEFEUILLE'], [p.XLabel[n]], True);
    {Si l'on a trouvé la tob, on va afficher le graph 2}
    if Assigned(T) and (T.Detail.Count > 0) then begin
      IndexMere := T.GetIndex;
      {Initialisation du titre}
      TFGRS1(Ecran).FChart2.Title.Text.Clear;
      TFGRS1(Ecran).FChart2.Title.Text.Add('Detail des OPCVM du Portefeuille '+ p.XLabel[n]);
      {Récupération ou création des séries du graph}
      if TFGRS1(Ecran).FChart2.SeriesList.Count > 0 then begin
        p := TPieSeries(TFGRS1(Ecran).FChart2.SeriesList.Series[0]);
        p.Clear;
        s := TPieSeries(TFGRS1(Ecran).FChart2.SeriesList.Series[1]);
        s.Clear;
      end
      else begin
        p := TPieSeries.Create(TFGRS1(Ecran).FChart2);
        TFGRS1(Ecran).FChart2.AddSeries(p);
        s := TPieSeries.Create(TFGRS1(Ecran).FChart2);
        TFGRS1(Ecran).FChart2.AddSeries(s);
      end;

      TPieSeries(TFGRS1(Ecran).FChart2.SeriesList.Series[1]).CustomXRadius := Round(Ecran.Width / 4) - 40;
      TPieSeries(TFGRS1(Ecran).FChart1.SeriesList.Series[1]).CustomXRadius := Round(Ecran.Width / 4) - 40;
      TFGRS1(Ecran).FChart2.SeriesList.Series[1].OnGetMarkText := SerieChart2GetMarkText;
      TPieSeries(TFGRS1(Ecran).FChart2.SeriesList.Series[1]).Circled := True;
      TPieSeries(TFGRS1(Ecran).FChart1.SeriesList.Series[1]).Circled := True;
      TFGRS1(Ecran).FChart2.SeriesList.Series[1].Marks.ArrowLength := 2;
      TFGRS1(Ecran).FChart1.SeriesList.Series[1].Marks.ArrowLength := 2;
      TFGRS1(Ecran).FChart2.SeriesList.Series[1].Marks.Clip := False;
      TFGRS1(Ecran).FChart1.SeriesList.Series[1].Marks.Clip := False;

      {La première série est celle des OPCVM, mais il n'est pas besoin de la gérer}
      for n := 0 to T.Detail.count - 1 do
        s.Add(T.Detail[n].GetDouble('MONTANT'), T.Detail[n].GetString('OPCVM'));

      TFGRS1(Ecran).Splitter1.Visible := True;
      TFGRS1(Ecran).FChart2.Visible := True;
      TFGRS1(Ecran).FChart2.Width := Round(Ecran.Width / 2) - 10;
    end;
  end;
end;


{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVMREPA.Chart2MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{---------------------------------------------------------------------------------------}
begin
  {On cache le Graph 2}
  TFGRS1(Ecran).Splitter1.Visible := False;
  TFGRS1(Ecran).FChart2.Visible := False;
  TFGRS1(Ecran).FChart2.Width   := 0;
  TPieSeries(TFGRS1(Ecran).FChart1.SeriesList.Series[1]).CustomYRadius := Round(TFGRS1(Ecran).FChart1.Height / 2) - 50;
  TPieSeries(TFGRS1(Ecran).FChart1.SeriesList.Series[1]).Circled := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVMREPA.SerieChart1GetMarkText(Sender : TChartSeries; ValueIndex : Longint; var MarkText : hString);
{---------------------------------------------------------------------------------------}
begin
  MarkText := MarkText + '(' + VarToStr(TobSeries.Detail[ValueIndex].GetValue('MONTANT')) + ' %)';
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVMREPA.SerieChart2GetMarkText(Sender : TChartSeries; ValueIndex : Longint; var MarkText : hString);
{---------------------------------------------------------------------------------------}
begin
  if IndexMere > -1 then
    MarkText := MarkText + '(' + VarToStr(TobSeries.Detail[IndexMere].Detail[ValueIndex].GetValue('MONTANT')) + ' %)';
end;

initialization
  RegisterClasses([TOF_TRGRAPHOPCVMREPA]);

end.

