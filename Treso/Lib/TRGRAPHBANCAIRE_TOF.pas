{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 6.00.018.001    14/10/04    JP   Création de l'unité
 8.00.001.018    05/06/07    JP   FQ 10469 : désactvation des filtres et des boutons
--------------------------------------------------------------------------------------}
unit TRGRAPHBANCAIRE_TOF;

interface

uses
  Controls, Classes, Graphics, TeEngine, Chart,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, 
  {$ENDIF}
  Forms, SysUtils, HCtrls, GRS1, Series, GraphUtil, UTOF, UTob,
  HTB97;


type
  TOF_TRGRAPHBANCAIRE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    ListeCourbes : string;
    procedure DessineLigne(Sender : TObject);
  end ;

procedure TRLanceFiche_TRGRAPHBANCAIRE(Arguments : string; TobG : TOB);

implementation

uses
  UObjFiltres;

var
  TobGraph : Tob;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_TRGRAPHBANCAIRE(Arguments : string; TobG : TOB);
{---------------------------------------------------------------------------------------}
begin
  TobGraph := TobG;
  AGLLanceFiche('TR', 'TRGRAPHBANCAIRE', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHBANCAIRE.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  F            : TFGRS1;
  sttitre,
  stColonnes,
  stChampTitre,
  stCriteres,
  stWhere,
  stTitresCol1,
  stTitresCol2,
  stColGraph1,
  stColGraph2  : string;
  tstTitres    : Tstrings;
  ft           : TFont;
begin
  inherited;
  stCriteres   := '';
  stColGraph1  := '';
  stColGraph2  := '';
  stTitresCol1 := '';
  stTitresCol2 := '';
  sttitre      := '';

  tstTitres := TStringList.Create;
  {Definition des fonts des titres}
  ft := TFont.Create;
  ft.Style := [fsBold];
  ft.Color := clBlue;
  ft.Size  := 12;

  F := TFGRS1(Ecran);
  stColonnes := ListeCourbes;{Liste des champs}
  stColGraph1 := ListeCourbes;{Liste des champs}
  stColGraph2 := '';
  stChampTitre := stTitre;
  tstTitres.Add (Ecran.Caption);
  stTitre := 'Dates;Total;Pointées'; {légende des colonnes}

  LanceGraph(F, TobGraph, '', stColonnes, stWhere , stTitre , stColGraph1,
               stColGraph2, tstTitres, nil, TLineSeries, stChampTitre, False);
  if F.Fchart1.SeriesCount > 0 then begin
    F.FChart1.SeriesList.series[0].Marks.Visible := False;
    F.FChart1.SeriesList.series[0].DataSource := nil;
    {Pour ne pas afficher le relief}
    F.FChart1.View3D := False;
    {Couleur des séries}
    F.FChart1.SeriesList.series[1].SeriesColor := clRed;
    F.FChart1.SeriesList.series[2].SeriesColor := clNavy;
    {Pour dessiner une partie de la série dans une couleur
    F.FChart1.SeriesList.series[1].ColorRange(F.FChart1.SeriesList.series[1].XValues, F.FChart1.SeriesList.series[1].XValues.First, F.FChart1.SeriesList.series[1].XValues.Last, clRed);
    F.FChart1.SeriesList.series[2].ColorRange(F.FChart1.SeriesList.series[2].XValues, F.FChart1.SeriesList.series[2].XValues.First, F.FChart1.SeriesList.series[2].XValues.Last, clBlue);}

    {Pour dessiner des lignes horizontales sur le graphe}
    F.FChart1.SeriesList.Series[1].BeforeDrawValues := DessineLigne;
    {Format des montant de l'échelle de gauche}
    F.FChart1.LeftAxis.AxisValuesFormat:='# ##0,00';

    F.FChart1.LeftAxis  .Title.caption := 'Soldes';
    F.FChart1.BottomAxis.Title.caption := 'Date';
    {Pour empiler les bares et mbStacked100 sur une échelle de 100%
    TBarSeries(F.FChart1.SeriesList.series[1]).MultiBar := mbSide;
    TBarSeries(F.FChart1.SeriesList.series[2]).MultiBar := mbSide;}
    {Pour rendre invisible les marks}
    F.FChart1.SeriesList.Series[1].Marks.Visible := False;
    F.FChart1.SeriesList.Series[2].Marks.Visible := False;
    {Pour n'afficher qu'une rubrique dans la légende : la série 0 est la date, on s'en moque dans la légende}
    F.FChart1.Legend.FirstValue := 1;
    {Font des titres des axes du graph}
    F.FChart1.LeftAxis  .Title.Font.Assign(ft);
    F.FChart1.BottomAxis.Title.Font.Assign(ft);
    F.FChart1.LeftAxis  .Title.Font.Color := clGreen;
    ft.Size := 9;
    ft.Style := [];
    {Font des échelles des axes}
    F.FChart1.LeftAxis  .LabelsFont.Assign(ft);
    F.FChart1.BottomAxis.LabelsFont.Assign(ft);
    F.FChart1.LeftAxis  .LabelsFont.Color := clGreen;
  end ;
  FreeAndNil(ft);
  tstTitres.Free;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHBANCAIRE.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  SetControlVisible('BUNDO'  , False); {FQ 10128}
  SetControlVisible('BSAUVE' , False); {FQ 10128}
  SetControlVisible('BDELETE', False); {FQ 10128}

  Ecran.Caption := ReadTokenSt(S);
  UpdateCaption(Ecran);
  SetControlCaption('TITRE', Ecran.Caption);
  SetControlCaption('DEVISE', ReadTokenSt(S));
  SetControlCaption('SOLDE', ReadTokenSt(S));
  {Récupération des colonnes du graphe}
  ListeCourbes := S;

  {05/06/07 : FQ 10469 : on cache les filtres}
  CacheFiltreGraph(Ecran);

  TToolbarButton97(GetControl('BVALIDER')).Click;
  TToolbarButton97(GetControl('BAFFGRAPH')).Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHBANCAIRE.DessineLigne(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  with TFGRS1(Ecran).FChart1, Canvas do begin
    {Définition du crayon}
    Pen.Width := 2;
    Pen.Style := psSolid;
    Pen.Color := clPurple;
    {Dessin de la ligne}
    MoveTo(ChartRect.Left, Series[1].CalcYPosValue(0));
    LineTo(ChartRect.Left , Series[1].CalcYPosValue(0));
    LineTo(ChartRect.Right, Series[1].CalcYPosValue(0));
    {Dessin d'un bout de ligne sur la gauche de l'axe pour faire le lien avec la marque}
    MoveTo(ChartRect.Left, Series[1].CalcYPosValue(0));
    LineTo(ChartRect.Left - 5, Series[1].CalcYPosValue(0));
  end
end;

initialization
  RegisterClasses([TOF_TRGRAPHBANCAIRE]);

end.
