{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.2X.xxx.xxx    04/11/04   JP   Création de l'unité  : Graph des cours des OPCVM calculés
                                 sur une Base 100 sur le premier jour de la sélection
 8.00.001.018    05/06/07   JP   FQ 10469 : désactvation des filtres et des boutons
 8.10.006.005    18/01/08   JP   FQ 10499 : la FQ 10469 n'aurait pas due être appliquée ici
--------------------------------------------------------------------------------------}
unit TRGRAPHOPCVM_TOF;

interface

uses
  Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, HMsgBox, GRS1, Series, GraphUtil,
  Graphics, TeEngine, Chart, UTOF, UTob, ExtCtrls;

type

  TOF_TRGRAPHOPCVM = class (TOF)
    procedure OnUpdate              ; override;
    procedure OnClose               ; override;
    procedure OnArgument(S : string); override;
  private
    DateDe : THEdit;
    DateA  : THEdit;
    NbCol  : Byte;

    {Prépare la grille en fonction du nombre de cours sélectionnés}
    procedure PreparerAffichage(var Champs, Libelles : string);
    {Prepare la tob en base 100}
    procedure PrepareLaTob(var TT : TOB);
    {Prépare le graph en fonction du nombre de cours sélectionnés}
    procedure PrepareGraph;
    {Calcule les cours en base 100 à la date de début de sélection}
    procedure CalculeBase100(var TT : TOB);
  protected
    {Gestion d'un Timer pour les Hint}
    FPause    : TTimer;
    IsCache   : Boolean;
    HintGraph : THintWindow;
    {procedure sur le OnTimer}
    procedure CacheHint(Sender : TObject);
  public
    {Tob  contenant les valeurs des séries}
    TobSerie : Tob;
    {Pour dessiner une ligne jaune sur la base 100 dans le BeforeDrawValues}
    procedure DessineLigne(Sender : TObject);
    procedure SerieClick  (Sender : TChartSeries; ValueIndex : Longint; Button : TMouseButton; Shift : TShiftState; X, Y : Integer);
    procedure BAideClick  (Sender : TObject);
  end ;

procedure TRLanceFiche_OPCVMBase100(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  UObjFiltres, UProcGen, Windows;

type
  TabOpcvm = array [1..3] of string;
  BasOpcvm = array [1..3] of Double;

var
  tOpcvm : TabOpcvm;
  tColGr : TabOpcvm;

const
  NBOPCVM = 3;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_OPCVMBase100(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  sttitre, stColonnes,
  stChampTitre, stCriteres,
  stTitresCol2,
  stColGraph1,
  stColGraph2 : string;
  tstTitres   : TStrings;
  n           : Byte;
begin
  inherited;
  NbCol := 1;
  for n := 1 to NBOPCVM do begin
    tOpcvm[n] := GetControlText('OPCVM' + IntToStr(n));
    if tOpcvm[n] <> '' then begin
      Inc(NbCol);
      tColGr[NbCol-1] := RechDom('TROPCVM', tOpcvm[n], False);
    end;
  end;

  if (tOpcvm[1] = '') and (tOpcvm[2] = '') and (tOpcvm[3] = '') then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner au moins un OPCVM.;W;O;O;O;', '', '');
    Exit;
  end;

  TobSerie.ClearDetail;
  tstTitres := TStringList.Create;
  try
    {Prépare la grille et le graph en fonction du nombre de cours sélectionnés}
    PreparerAffichage(stColonnes, stTitre);
    {Prepare la tob en base 100}
    PrepareLaTob(TobSerie);

    stCriteres := '';
    stColGraph2 := '';
    stTitresCol2 := '';
    stColGraph1 := stColonnes;
    stChampTitre := stTitre;

    tstTitres.Add ('Cours Comparés (en base 100 au ' + DateDe.Text + ')');

    LanceGraph(TFGRS1(Ecran), TobSerie, '', stColonnes, '', stTitre, stColGraph1, stColGraph2,
                tstTitres, nil, TLineSeries, stChampTitre, False);

    if TFGRS1(Ecran).Fchart1.SeriesCount > 0 then begin
      PrepareGraph;
      {Initialisation des échelles}
    end ;
  finally
    if Assigned(tstTitres) then tstTitres.Free;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetControlVisible('BUNDO'  , False); {FQ 10128}
  SetControlVisible('BSAUVE' , False); {FQ 10128}
  SetControlVisible('BDELETE', False); {FQ 10128 / 19/08/08 : FQ 10499 : on cache le bouton}
  TFGRS1(Ecran).BDelete.OnClick := BAideClick;
  DateDe := THEdit(GetControl('CRITDATE')); {FQ 10129}
  DateA  := THEdit(GetControl('CRITDATE_'));{FQ 10129}
  TFGRS1(Ecran).HelpContext := 150;
  {Timer utilisé dans la gestion manuelle des Hints}
  FPause := TTimer.Create(Ecran);
  FPause.Interval := 2000;
  FPause.Enabled  := False;
  FPause.OnTimer := CacheHint;
  TobSerie := Tob.Create('µµµ', nil, -1);

  {05/06/07 : FQ 10469 : on cache les filtres
   18/01/08 : FQ 10499 : En fait ici, il ne faut pas cacher !!!
  CacheFiltreGraph(Ecran);}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FPause  ) then FreeAndNil(FPause  );
  if Assigned(TobSerie) then FreeAndNil(TobSerie);
  inherited;
end;

{Prépare la grille et le graph en fonction du nombre de cours sélectionnés}
{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.PreparerAffichage(var Champs, Libelles : string);
{---------------------------------------------------------------------------------------}
var
  n : Byte;
begin
  Champs   := 'TTO_DATE;';
  Libelles := 'Date;';
  with TFGRS1(Ecran).FListe do begin
    {Pour forcer la réinitialisation du nom des colonnes qui contiennent, outre
     la date pour la première d'entre elles, les code OPCVM pour les autres. Le
     nom de la colonne me sert à récupérer le code OPCVM de la série courante
     dans le OnClick}
    for n := 0 to ColCount - 1 do ColNames[n] := '';

    ColCount := NbCol;

    ColWidths[0] := Round(Width - 25 / NbCol); {25 pour la ScrollBar}
    for n := 1 to NbCol - 1 do begin
      ColFormats[n] := '#,####0.0000';
      ColWidths [n] := Round(Width - 25 / NbCol);
      ColAligns [n] := taRightJustify;
    end;

    for n := 1 to NBOPCVM do begin
      if tOpcvm[n] <> '' then
        Champs := Champs + tOpcvm[n] + ';';
      if tColGr[n] <> '' then
        Libelles := Libelles + tColGr[n] + ';';
    end;
  end;
end;

{Prepare la tob en base 100}
{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.PrepareLaTob(var TT : TOB);
{---------------------------------------------------------------------------------------}
var
  s : string;
  c : string;
  T : TOB;
  F : TOB;
  n : Integer;
  p : Byte;
begin
  c := ' AND TTO_CODEOPCVM IN (';
  for p := 1 to 3 do begin
    if tOpcvm[p] <> '' then
      c := c + '"' + tOpcvm[p] + '",';
  end;

  System.Delete(c, Length(c), 1);
  c := c + ') ORDER BY TTO_DATE, TTO_CODEOPCVM';

  s := 'SELECT * FROM TRCOTATIONOPCVM WHERE TTO_DATE BETWEEN "';
  s := s + UsDateTime(StrToDate(DateDe.Text)) + '" AND "';
  s := s + UsDateTime(StrToDate(DateA.Text)) + '" ' + c;

  T := TOB.Create('$$$', nil, -1);
  try
    T.LoadDetailFromSQL(s);
    s := '';
    for n := 0 to T.Detail.Count - 1 do begin
      {On regarde s'il y a déjà une ligne pour cette date ...}
      F := TT.FindFirst(['TTO_DATE'], [T.Detail[n].GetString('TTO_DATE')], False);
      {... Sinon on crée une nouvelle Tob fille}
      if not Assigned(F) then begin
        F := TOB.Create('£££', TT, -1);
        F.AddChampSupValeur('TTO_DATE', T.Detail[n].GetString('TTO_DATE'));
      end;

      {On stocke le cours de L'OPCVM courante}
      for p := 1 to NBOPCVM do
        if (tOpcvm[p] = T.Detail[n].GetString('TTO_CODEOPCVM')) then
          F.AddChampSupValeur(tOpcvm[p], T.Detail[n].GetString('TTO_COTATION'))
        else
          F.AddChampSup(tOpcvm[p], False);
    end;
    {On met les montants en Base 100 pour pouvoir les comparer}
    CalculeBase100(TT);
  finally
    if Assigned(T) then FreeAndNil(T);
  end;
end;

{Prépare le graph en fonction du nombre de cours sélectionnés}
{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.PrepareGraph;
{---------------------------------------------------------------------------------------}
var
  F  : TFGRS1;
  ft : TFont;
  n  : Byte;
  Mini,
  Maxi : Double;
begin
  F := TFGRS1(Ecran);
  if F.Fchart1.SeriesCount <= 1 then Exit;

  ft := TFont.Create;
  try
    {Pour dessiner des lignes horizontales sur le graphe}
    F.FChart1.SeriesList.Series[1].BeforeDrawValues := DessineLigne;
    {Pour un affichage en traits continus sans relief}
    F.FChart1.View3D := False;

    {Definition des fonts des titres}
    ft.Style := [fsBold];
    ft.Color := clBlue;
    ft.Size  := 12;

    F.FChart1.SeriesList.series[0].Marks.Visible := False;
    F.FChart1.SeriesList.series[0].DataSource := nil;

    F.FChart1.LeftAxis.Title.caption := 'Cours';
    F.FChart1.BottomAxis.Title.caption := 'Date';

    for n := 1 to NbCol - 1 do begin
      F.FChart1.SeriesList.Series[n].Title := tColGr[n];
      F.FChart1.SeriesList.Series[n].OnClick := SerieClick;
      TLineSeries(F.FChart1.SeriesList.Series[n]).Pointer.Visible := True;
    end;

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
    {Pour éviter le jaune qui n'est pas très visible dans la légende et le vert qui est la couleur
     de l'axe de gauche et de la ligne de la base 100}
    F.FChart1.SeriesList.series[1].SeriesColor := clPurple;
    if F.Fchart1.SeriesCount > 2 then
      F.FChart1.SeriesList.series[2].SeriesColor := clRed;

    {Définition du format des montants de l'échelle de gauche}
    F.FChart1.LeftAxis.AxisValuesFormat := '#,##0.0';

    {Gestion de l'échelle de l'axe de gauche}
    F.FChart1.LeftAxis.CalcMinMax(Mini, Maxi);
    F.FChart1.LeftAxis.Increment := ArrondirTreso('01', (Maxi - Mini) / 20);

  finally
    if Assigned(ft) then FreeAndNil(ft);
  end;
end;

{Calcule les cours en base 100 à la date de début de sélection}
{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.CalculeBase100(var TT : TOB);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Byte;
  B : BasOpcvm;
  M : BasOpcvm;
  T : TOB;
begin
  {On récupère les bases}
  for p := 1 to NBOPCVM do begin
    B[p] := 0;
    M[p] := 100;
    if tOpcvm[p] <> '' then begin
      {Par défaut, la base sera le premier cours}
      for n := 0 to TT.Detail.Count - 1 do begin
        T := TT.Detail[n];
        B[p] := T.GetDouble(tOpcvm[p]);
        if B[p] <> 0 then Break;
      end;
    end;
  end;

  {Calcul des montants en base 100}
  for n := 0 to TT.Detail.Count - 1 do begin
    T := TT.Detail[n];
    for p := 1 to NBOPCVM do begin
      {Si sur le premier jour de la période, il n'y avait pas de valeur ...}
      if B[p] = 0 then begin
        {... si à la date courante, il y a une valeur, c'est la première de la série ...}
        if T.GetDouble(tOpcvm[p]) <> 0 then
          {.. on récupère cette valeur qui va servir de base}
          B[p] := T.GetDouble(tOpcvm[p])
        {... sinon on passe à la série suivante
        else
          Continue;}
      end;

      if (tOpcvm[p] <> '') and (T.GetDouble(tOpcvm[p]) <> 0) then
        T.SetDouble(tOpcvm[p], T.GetDouble(tOpcvm[p]) / B[p] * 100)
      {Sur ce jour, le cours n'a pas été saisi, on reprend le précedent}
      else if (tOpcvm[p] <> '') {and (T.GetString(tOpcvm[p]) <> '')} then
        T.SetDouble(tOpcvm[p], M[p]);
      {On mémorise le base calculée au cas où le cours suivant ne serait pas saisi
       et ainsi pouvoir reporter}
      M[p] := T.GetDouble(tOpcvm[p]);
    end;
  end;
end;

{Pour dessiner une ligne jaune sur la base 100 dans le BeforeDrawValues
{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.DessineLigne(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Y : LongInt;
begin
  with TFGRS1(Ecran).FChart1, Canvas do begin
    Y := Series[1].CalcYPosValue(100);
    {Définition du crayon}
    Pen.Width := 1;
    Pen.Style := psSolid;
    Pen.Color := clGreen;
    {Dessin de la ligne}
    MoveTo(ChartRect.Left , Y);
    LineTo(ChartRect.Right, Y);
  end
end;

{Ferme le hint contenant le cours d'une série
{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.CacheHint(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not IsCache then begin
    {Cache le hint}
    HintGraph.ReleaseHandle;
    FreeAndNil(HintGraph);
    {Désactive le Timer}
    IsCache := True;
    FPause.Enabled := False;
  end;
end;

{Affichage dans une bulle du cours en devise de la série courante à la date en cours
{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.SerieClick(Sender : TChartSeries; ValueIndex : Longint; Button : TMouseButton; Shift : TShiftState; X, Y : Integer);
{---------------------------------------------------------------------------------------}
var
  R : TRect;
  p : TPoint;
  s : string;
  m : string;
  Q : TQuery;

    {Recherche du code de l'OPCVM en cours : on recherche dans la grille
     la colonne qui porte le même nom que la série en cours et on récupère
     le nom du champ qui correspond à cette colonne
    {-------------------------------------------------------------}
    function GetNom : string;
    {-------------------------------------------------------------}
    var
      n : Byte;
      c : string;
    begin
      Result := '';
      for n := 1 to NbCol - 1 do begin
        c := UpperCase(TFGRS1(Ecran).FListe.Cells[n, 0]);
        if c = UpperCase(Sender.Title) then begin
          Result := TFGRS1(Ecran).FListe.ColNames[n];
          Exit;
        end;
      end;
    end;

begin
  {S'il y a un Hint en cours, on sort}
  if Assigned(HintGraph) then Exit;
  {Recherche du cours et de la devise de l'OPCVM à la date en cours}
  Q := OpenSQL('SELECT TTO_COTATION, TOF_DEVISE FROM TRCOTATIONOPCVM LEFT JOIN TROPCVMREF ON TOF_CODEOPCVM = TTO_CODEOPCVM ' +
               'WHERE TTO_DATE = "' + UsDateTime(TobSerie.Detail[ValueIndex].GetDateTime('TTO_DATE')) +
               '" AND TTO_CODEOPCVM = "' + GetNom + '"', True);
  try
    if not Q.EOF then begin
      s := Q.FindField('TTO_COTATION').AsString + ' ' + Q.FindField('TOF_DEVISE').AsString;
      m := Q.FindField('TOF_DEVISE').AsString;
    end;
  finally
    Ferme(Q);
  end;

  {Création du Hint}
  HintGraph := THintWindow.Create(TFGRS1(Ecran).FChart1);
  HintGraph.Color := clInfoBk;
  {Positionnement du hint}
  p := TFGRS1(Ecran).FChart1.ClientToScreen(Point(X, Y));
  {Pour la largeur, il s'agit d'une approximation, mais TextWidth "refuse" de me renvoyer la bonne valeur}
  R := Rect(p.X + 30 , p.Y, p.X + HintGraph.Canvas.TextWidth(m) + HintGraph.Canvas.TextWidth(s) + 25, p.Y + 16);
  {Affichage du hint}
  HintGraph.ActivateHint(r, s); 
  {Activation du timer pour la destruction du hint}
  IsCache := False;
  FPause.Enabled := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHOPCVM.BAideClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  PGIInfo(TraduireMemoire('Pour avoir les cours en devise, il suffit de cliquer sur les points des courbes.'), Ecran.Caption); 
end;

initialization
  RegisterClasses([TOF_TRGRAPHOPCVM]);

end.

