{***********UNITE*************************************************
Auteur  ...... : Jean Pastéris
Créé le ...... : 23/01/2004
Modifié le ... : 23/01/2004
Description .. : Graph des opérations de trésorerie et comptable
Mots clefs ... : TRLANCEFICHE_TRGRAPHTROCTA
*****************************************************************}
{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 1.01            23/01/04    JP     Création de l'unité
 6.00.014.001    17/09/04    JP     FQ 10128 : on cache les boutons inutiles
 8.00.001.018    05/06/07    JP     FQ 10469 : désactvation des filtres et des boutons
--------------------------------------------------------------------------------------}
unit TRGRAPHTROCTA_TOF ;

interface

uses
  Controls, Classes, Graphics, TeEngine, Chart,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, 
  {$ENDIF}
  Forms, SysUtils, HCtrls, GRS1, Series, GraphUtil,UTOF, UTob,
  HTB97;


type
  TOF_TRGRAPHTROCTA = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    ListeCourbes : string;
  end ;

procedure TRLanceFiche_TRGRAPHTROCTA(Arguments : string; TobG : TOB);

implementation

uses
  UObjFiltres;

var
  TobGraph : Tob;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_TRGRAPHTROCTA(Arguments : string; TobG : TOB);
{---------------------------------------------------------------------------------------}
begin
  TobGraph := TobG;
  AGLLanceFiche('TR', 'TRGRAPHTROCTA', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHTROCTA.OnUpdate ;
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
  CumulOk      : Boolean;
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
  CumulOk := Pos('CUMTRE', ListeCourbes) > 0;
  if CumulOk then begin
    tstTitres.Add ('Cumuls des opérations de trésorerie et de comptabilité');
    stTitre := 'Dates;Cumuls de trésorerie;Cumuls de comptabilité'; {légende des colonnes}
    LanceGraph(F, TobGraph, '', stColonnes, stWhere , stTitre , stColGraph1,
                 stColGraph2, tstTitres, nil, TBarSeries , stChampTitre, False);
    F.FChart1.SeriesList.series[1].Marks.Visible := False;
    F.FChart1.SeriesList.series[2].Marks.Visible := False;
    {Pour empiler les bars : ici cela n'a aucun sens puisque les montants peuvent de signes
     opposés. mbStacked100 pour empiler sur une échelle de 100
     TBarSeries(F.FChart1.SeriesList.series[1]).MultiBar := mbStacked};
  end
  else begin
    tstTitres.Add ('Suivi des soldes de la trésorerie et de la comptabilité');
    stTitre := 'Dates;Solde avant opérations;Solde après opé. comptables;Solde après opé. de trésorerie'; {légende des colonnes}
    LanceGraph(F, TobGraph, '', stColonnes, stWhere , stTitre , stColGraph1,
                 stColGraph2, tstTitres, nil, TLineSeries , stChampTitre, False);
  end;

  if F.Fchart1.SeriesCount > 0 then begin
    F.FChart1.SeriesList.series[0].Marks.Visible := False;
    F.FChart1.SeriesList.series[0].DataSource := nil;
    if CumulOk then F.FChart1.LeftAxis.Title.caption := 'Cumuls'
               else F.FChart1.LeftAxis.Title.caption := 'Soldes';
    F.FChart1.BottomAxis.Title.caption := 'Date';

    {Pour n'afficher qu'une rubrique dans la légende : la série 0 est la date, on s'en moque dans la légende}
    F.FChart1.Legend.FirstValue := 1;
    F.FChart1.LeftAxis .LabelsFont.Color := clGreen;
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
procedure TOF_TRGRAPHTROCTA.OnArgument (S : String ) ;
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

initialization
  RegisterClasses([TOF_TRGRAPHTROCTA]);

end.
