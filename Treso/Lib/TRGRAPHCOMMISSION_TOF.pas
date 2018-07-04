{ Unité :Source TOF de la FICHE : TRGRAPHCOMMISSION
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.0x.xxx.xxx    03/08/04    JP   Création de l'unité
 6.00.014.001    17/09/04    JP   FQ 10128 : on cache les boutons inutiles
 8.00.001.018    05/06/07    JP   FQ 10469 : désactvation des filtres et des boutons
--------------------------------------------------------------------------------------}
unit TRGRAPHCOMMISSION_TOF;

interface

uses
  Classes, Graphics, TeEngine, Chart, Windows,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, 
  {$ENDIF}
  SysUtils, HCtrls, GRS1, Series, GraphUtil, UTOF, UTob,
  HTB97;


type
  TOF_TRGRAPHCOMMISSION = class(TOF)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
  private
    TobSeries : TOB;
    ListeCourbes : string;
    procedure DessineLigne  (Sender : TObject);
  end ;

procedure TRLanceFiche_GraphCommission(Arguments : string);


implementation

uses
  UObjFiltres;


{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_GraphCommission(Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('TR', 'TRGRAPHCOMMISSION', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHCOMMISSION.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
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
    TstTitres.Add ('Courbes des commissions');

    stColonnes  := ListeCourbes;
    stTitre     := ListeCourbes;
    stColGraph1 := ListeCourbes;
    stColGraph2 := '';
    stChampTitre := stTitre;
    LanceGraph(F, TobSeries, '', stColonnes, stWhere , stTitre , stColGraph1,
                 stColGraph2, tstTitres, nil, TLineSeries , stChampTitre, False);

    if F.Fchart1.SeriesCount > 0 then begin
      {Pour un affichage en traits continus sans relief}
      F.FChart1.View3D := False;

      {Pour dessiner des lignes horizontales sur le graphe}
      F.FChart1.SeriesList.Series[1].BeforeDrawValues := DessineLigne;

      {Les comptes ne servent sue de marques, on ne les affiche pas dans le graphe}
      F.FChart1.SeriesList.Series[0].Marks.Visible := False;
      F.FChart1.SeriesList.Series[0].DataSource := nil;

      {Titre des axes}
      F.FChart1.LeftAxis.Title.Caption := 'Montants';
      F.FChart1.BottomAxis.Title.Caption := 'Mois';

      {Titre des légendes}
      ReadTokenSt(ListeCourbes);
      for n := 1 to F.FChart1.SeriesList.Count - 1 do
        F.FChart1.SeriesList.Series[n].Title := ReadTokenSt(ListeCourbes);

      {Pour une question de place, on met la légende en bas de page}
      F.FChart1.Legend.Alignment := laBottom;

      {Définition du format des montants de l'échelle de gauche}
      F.FChart1.LeftAxis.AxisValuesFormat := '#,##0.00';

      {Pour ne pas afficher les comptes dans la légende : la série 0 concerne les comptes, on s'en moque dans la légende}
      F.FChart1.Legend.FirstValue := 1;
      F.FChart1.LeftAxis.LabelsFont.Color := clRed;
      {Font des titres des axes du graph}
      F.FChart1.LeftAxis  .Title.Font.Assign(ft);
      F.FChart1.BottomAxis.Title.Font.Assign(ft);
      F.FChart1.LeftAxis  .Title.Font.Color := clRed;

      ft.Size := 9;
      ft.Style := [];
      {Font des échelles des axes}
      F.FChart1.LeftAxis  .LabelsFont.Assign(ft);
      F.FChart1.BottomAxis.LabelsFont.Assign(ft);
      F.FChart1.LeftAxis  .LabelsFont.Color := clRed;
    end;
  finally
    FreeAndNil(ft);
    tstTitres.Free;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHCOMMISSION.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetControlVisible('BUNDO'  , False); {FQ 10128}
  SetControlVisible('BSAUVE' , False); {FQ 10128}
  SetControlVisible('BDELETE', False); {FQ 10128}

  SetControlCaption('TITRE', ReadTokenSt(S));
  SetControlCaption('DEVISE', ReadTokenSt(S));
  {Récupération des colonnes du graphe}
  ListeCourbes := S;
  TobSeries := TFGRS1(Ecran).LaTOF.LaTOB;

  {05/06/07 : FQ 10469 : on cache les filtres}
  CacheFiltreGraph(Ecran);

  {Lancement du traitement, pour qu'en arrivée sur la fiche, le graphe soit affiché}
  TToolbarButton97(GetControl('BVALIDER')).Click;
  TToolbarButton97(GetControl('BAFFGRAPH')).Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHCOMMISSION.DessineLigne(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  YPosition  : Longint;
  Ecart, Inc : Double;
  Mini, Maxi : Double;
  n, p, k    : Integer;
    {---------------------------------------------------------------------}
    procedure Dessine(Zero : Boolean);
    {---------------------------------------------------------------------}
    begin
      with TFGRS1(Ecran).FChart1, Canvas do begin
        {Définition du crayon}
        if Zero then Pen.Width := 2
                else Pen.Width := 1;
        Pen.Style := psSolid;
        Pen.Color := clBlue;
        {Dessin de la ligne}
        MoveTo(ChartRect.Left, YPosition);
        LineTo(ChartRect.Left  + Width3D, YPosition - Height3D);
        LineTo(ChartRect.Right + Width3D, YPosition - Height3D);
        {Dessin d'un bout de ligne sur la gauche de l'axe pour faire le lien avec la marque}
        MoveTo(ChartRect.Left, YPosition);
        LineTo(ChartRect.Left - 5, YPosition);
      end
    end;

begin
  p := 0;
  with TFGRS1(Ecran).FChart1 do begin
    {On ne déssine une ligne pour la valeur nulle que s'il y a des montants de sens différents}
    {Récupération de la position de 0 sur l'axe vertical}
    YPosition := Series[1].CalcYPosValue(0);
    Dessine(True);

    {Dessin des 7 principales marques}
    {On commence par calculer les extrêmes de la graduation et l'incrément, car à ce moment cela
     n'a pas été encore fait : Increment = 0 et Min et Max sont les valeurs Min et Max de la série}
    LeftAxis.CalcMinMax(Mini, Maxi);
    Inc   := LeftAxis.CalcIncrement;
    Ecart := Maxi - Mini;
    {On va mettre 7 lignes plus celui de zéro}
    Ecart := Ecart / 7;
    if Inc <> 0 then Ecart := Ecart / Inc
                else Exit;
    Ecart := Round(Ecart);
    for n := 1 to 7 do begin
      {Tant que le maximun n'est pas atteint, on dessine les lignes du côté positif}
      if n * Ecart * Inc < Maxi then begin
        YPosition := Series[1].CalcYPosValue(n * Ecart * Inc);
        Dessine(False);
        p := n;
      end
      {... sinon on passe du côté négatif}
      else begin
        for k := 1 to 7 - p do
          if k * Ecart * Inc * -1 > Mini then begin
            YPosition := Series[1].CalcYPosValue(k * Ecart * Inc * -1);
            Dessine(False);
          end;
        Exit;
      end;
    end;
  end;
end;

initialization
  RegisterClasses([TOF_TRGRAPHCOMMISSION]);

end.
