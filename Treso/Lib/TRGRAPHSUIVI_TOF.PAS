{ Unité :Source TOF de la FICHE : TRGRAPHSUIVI
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 0.91            17/09/03    JP     Création de l'unité
 6.00.014.001    17/09/04    JP     FQ 10128 : on cache les boutons inutiles
 8.00.001.018    05/06/07    JP     FQ 10469 : désactvation des filtres et des boutons
--------------------------------------------------------------------------------------}
unit TRGRAPHSUIVI_TOF ;

interface

uses
  Controls, Classes, Graphics, TeEngine, Chart, 
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, 
  {$ENDIF}
  sysutils, HCtrls, Forms, GRS1, Series, GraphUtil,UTOF, UTob,
  HTB97;


type
  TOF_TRGRAPHSUIVI = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    ListeCourbes : string;
  end ;

procedure TRLanceFiche_GraphSuivi(Arguments : string; TobG : TOB);

implementation

uses
  UObjFiltres;

var
  TobGraph : Tob;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_GraphSuivi(Arguments : string; TobG : TOB);
{---------------------------------------------------------------------------------------}
begin
  TobGraph := TobG;
  AGLLanceFiche('TR', 'TRGRAPHSUIVI', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSUIVI.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  F : TFGRS1;
  sttitre,
  stColonnes, stChampTitre,
  stCriteres, stWhere,
  stTitresCol1, stTitresCol2,
  stColGraph1, stColGraph2 : string;
  tstTitres                : Tstrings;
  ft                       : TFont;
begin
  Inherited ;
  stCriteres := '';
  stColGraph1 := '';
  stColGraph2 := '';
  stTitresCol1 := '';
  stTitresCol2 := '';
  sttitre := '';

  tstTitres := TStringList.Create;
  {Definition des fonts des titres}
  ft := TFont.Create;
  ft.Style := [fsBold];
  ft.Color := clBlue;
  ft.Size  := 12;

  F := TFGRS1(Ecran);
  tstTitres.Add ('Suivi de solde');
  stColonnes := ListeCourbes;//'DATE;SOLDE';
  stTitre := ListeCourbes;//'Date;Solde'; {légende des colonnes}
  stColGraph1 := ListeCourbes;//'DATE;SOLDE';
  stColGraph2 := '';
  stChampTitre := stTitre;
  LanceGraph(F, TobGraph, '', stColonnes, stWhere , stTitre , stColGraph1,
               stColGraph2, tstTitres, nil, TLineSeries , stChampTitre, False);

  if F.Fchart1.SeriesCount > 0 then begin
    F.FChart1.View3D := False;
    F.FChart1.SeriesList.series[0].Marks.Visible := False;
    F.FChart1.SeriesList.series[0].DataSource := nil;
    F.FChart1.LeftAxis.Title.caption := 'Solde';
    F.FChart1.BottomAxis.Title.caption := 'Date';
//    F.FChart1.SeriesList.series[1].Title := F.FChart1.LeftAxis.Title.caption;
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
    {Initialisation des échelles}
  end ;
  FreeAndNil(ft);
  tstTitres.Free;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSUIVI.OnArgument (S : String ) ;
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
  RegisterClasses ( [ TOF_TRGRAPHSUIVI ] ) ;

end.
