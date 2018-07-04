unit utoolsOWC;
{ Fonctionnement :
}


{$IFNDEF EAGLSERVER}abghgahgh{$ENDIF}

interface

uses classes,
  uTob,
  sysutils, OleCtrls, OWC11_TLB;

const
  OWCChartClsId = 'CLSID:0002E55D-0000-0000-C000-000000000046';
  OWCSpreadSheetClsId = 'CLSID:0002E559-0000-0000-C000-000000000046';
  OWCPivotTableClsId = 'CLSID:0002E55A-0000-0000-C000-000000000046';

type

  TTypeGraph = (tyCamembert, tyHistogramme, tyBar, tyLigne);
  TPositionText = (poPositionAutomatic, poPositionBottom, poPositionLeft, poPositionRight, poPositionTop);

  TCouleurGraph = (colNone, colAliceBlue, colAntiqueWhite, colAqua, colAquamarine, colAzure, colBeige, colBisque, colBlack,
    colBlanchedAlmond, colBlue, colBlueViolet, colBrown, colBurlyWood, colCadetBlue, colChartreuse,
    colChocolate, colCoral, colCornflower, colCornsilk, colCrimson, colCyan, colDarkBlue, colDarkCyan,
    colDarkGoldenrod, colDarkGray, colDarkGreen, colDarkKhaki, colDarkMagenta, colDarkOliveGreen,
    colDarkOrange, colDarkOrchid, colDarkRed, colDarkSalmon, colDarkSeaGreen, colDarkSlateBlue,
    colDarkSlateGray, colDarkTurquoise, colDarkViolet, colDeepPink, colDeepSkyBlue, colDimGray, colDodgerBlue,
    colFirebrick, colFloralWhite, colForestGreen, colFuchsia, colGainsboro, colGhostWhite, colGold,
    colGoldenrod, colGray, colGreen, colGreenYellow, colHoneydew, colHotPink, colIndianRed, colIndigo,
    colIvory, colKhaki, colLavender, colLavenderBlush, colLawnGreen, colLemonChiffon, colLightBlue,
    colLightCoral, colLightCyan, colLightGoldenrodYellow, colLightGray, colLightGreen, colLightPink,
    colLightSalmon, colLightSeaGreen, colLightSkyBlue, colLightSlateGray, colLightSteelBlue, colLightYellow,
    colLime, colLimeGreen, colLinen, colMagenta, colMaroon, colMediumAquamarine, colMediumBlue, colMediumOrchid,
    colMediumPurple, colMediumSeaGreen, colMediumSlateBlue, colMediumSpringGreen, colMediumTurquoise,
    colMediumVioletRed, colMidnightBlue, colMintCream, colMistyRose, colMoccasin, colNavajoWhite, colNavy,
    colOldLace, colOlive, colOliveDrab, colOrange, colOrangeRed, colOrchid, colPaleGoldenrod, colPaleGreen,
    colPaleTurquoise, colPaleVioletRed, colPapayaWhip, colPeachPuff, colPeru, colPink, colPlum, colPowderBlue,
    colPurple, colRed, colRosyBrown, colRoyalBlue, colSaddleBrown, colSalmon, colSandyBrown, colSeaGreen,
    colSeaShell, colSienna, colSilver, colSkyBlue, colSlateBlue, colSlateGray, colSnow, colSpringGreen,
    colSteelBlue, colTan, colTeal, colThistle, colTomato, colTransparent, colTurquoise, colViolet, colWheat,
    colWhite, colWhiteSmoke, colYellow, colYellowGreen);

  TObjParam = record
    NumFormat : string;
    MajorUnit : string;
    TypeSerie : TTypeGraph;
    LabelPerct: Boolean;
    LabelValue: Boolean;
    HasLabel  : Boolean;
  end;

  TObjText = record
    Visible: Boolean;
    Position: TPositionText;
    Bold: Boolean;
    Color: TCouleurGraph;
    Italic: Boolean;
    FontName: string;
    FontSize: Byte;
    FontColor : TCouleurGraph;
    Caption: string;
    InfoAxe: TObjParam;
  end;

  TAddSpecific = procedure(oChart : chChart) of object;

  TOWCChart = class(TObject)
  private
    FList: TOB;
    FTobResponse: TOB;
    FWidth: integer;
    FHeight: integer;
    FTypeGraph: TTypeGraph;
    FTitle: string;
    {12/05/06 : Nouvelle gestion des dimensions des graph : c'est paramétrable depuis le portail}
    FAspectGraph : string;

    function GetPositionI(Po : TPositionText) : Integer;
    function GetTypeChart(Typ : TTypeGraph) : ChartChartTypeEnum;
    function GetSerieCount : Integer;
    procedure SetTypeGraph(const Value: TTypeGraph);
    procedure SetGraphLegend(var oChart : chChart);
    procedure SetGraphTitle(var oChart : chChart);
    procedure SetInteriorsColor(var oChart : chChart);
    procedure SetGraphAxes(var oAxes : ChAxis; owObjAxe : TObjText);
    procedure SetDefaultValue;
  public
    owTitre   : TObjText;
    owLegend  : TObjText;
    owAxeLeft : TObjText;
    owAxeBot  : TObjText;
    owAxeRight: TObjText;
    owLabel   : TObjText;
    owContourColor: TCouleurGraph;
    owPlotColor: TCouleurGraph;
    owHasMajorLine: Boolean;
    owHasMinorLine: Boolean;
    {Si l'axe droite a une échelle différente de celui de droite}
    owIsAxeRightDegroupe : Boolean;
    owHasTwoCharts       : Boolean;
    owTitreChart2        : string;

    AddSpecific   : TAddSpecific;
    AddSpecific2  : TAddSpecific; {27/06/06 : Si deux Charts}

    constructor Create(TobResponse: TOB; NbDimension : string);
    destructor Destroy; override;
    procedure SetTitle(ATitle: string);
    procedure SetValue(ACategoryName, ASerieName: string; AValue: variant);
    procedure Build;
    procedure ExportXMLData;
    procedure ExportPicture(FluxXml : string);
    procedure SetWidth(AWidth: integer);
    procedure SetHEight(AHeight: integer);
    class function GetCouleurGraph(Co: TCouleurGraph): string;

    property TypeGraph: TTypeGraph read FTypeGraph write SetTypeGraph;
  end;

type
  TOWCPivotTable = class(TObject)
  end;

implementation

uses
  Math, HPanel,
  windows,
  activeX,
  comobj,
  {$IFDEF VER150}
  Variants,
  {$ENDIF VER150}
  hent1,
  esession;

{ TOWCChart }

constructor TOWCChart.Create(TobResponse: TOB; NbDimension : string);
begin
  FList := TOB.Create('the chart', nil, -1);
  FTobResponse := TobResponse;
  FWidth := FTobResponse.GetInteger('FWIDTH');
  FHeight := FTobResponse.GetInteger('FHEIGHT');

  FAspectGraph := NbDimension;

  {Initialisation des variables pour dessiner le graph}
  SetDefaultValue;
end;

destructor TOWCChart.Destroy;
begin
  FreeAndNil(FList);
end;

procedure TOWCChart.Build;
begin
  ExportXMLData;
end;

procedure TOWCChart.SetValue(ACategoryName, ASerieName: string; AValue: variant);
var
  TC, TS: TOB;
  ok: boolean;
begin
  //on cherche la tob de la categorie
  TC := FList.FindFirst(['CAPTION'], [ACategoryName], false);
  ok := TC <> nil;
  //si pas trouvé, on la crée
  if not ok then
  begin
    TC := TOB.Create('the category', FList, -1);
    TC.AddChampSupValeur('CAPTION', ACategoryName);
  end;
  //on cherche la tob de la serie
  TS := nil;
  if ok then
    TS := TC.FindFirst(['CAPTION'], [ASerieName], false);
  ok := TS <> nil;
  //si pas trouvé, on la crée
  if not ok then
  begin
    TS := TOB.Create('the serie', TC, -1);
    TS.AddChampSupValeur('CAPTION', ASerieName);
  end;
  TS.AddChampSupValeur('VALUE', AValue);
end;

function TOWCChart.GetSerieCount : Integer;
var
  L : TStringList;
  n : Integer;
  p : Integer;
  s : string;
begin
  L := TStringList.Create;
  try
    L.Sorted := True;
    L.Duplicates := dupIgnore;
    for n := 0 to FList.Detail.Count - 1 do
      for p := 0 to FList.Detail[n].Detail.Count - 1 do
        L.Add(FList.Detail[n].Detail[p].GetString('CAPTION'));

    Result := L.Count;
    s := L.Text;
  finally
    FreeAndNil(L);
  end;
end;

procedure TOWCChart.SetTitle(ATitle: string);
begin
  FTitle := ATitle;
  owTitre.Caption := FTitle;
  owTitre.Visible := True;
end;

procedure TOWCChart.ExportPicture(FluxXml : string);
var
  oConst, loSerie, loGraph, loChartSpace: OleVariant;
  stCaption, stValues, stVal, stCategories, stCat, stFileName: string;
  SeriesCount, CategoriesCount: integer;
  i, j: integer;
  T, TS, TC: tob;
begin
  CoInitialize(nil);
  {JP/EPZ 26/01/07 : Nouvelle gestion des graphiques}
  loChartSpace := CreateOleObject('OWC11.ChartSpace');
  loChartSpace.XmlData := FluxXml;
  //loChartSpace.Update;

  stFileName := LookupCurrentSession.eAglDocPath +
    'Portail\' + LookupCurrentSession.SessionId + '_owc.gif';
  if FileExists(stFileName) then
    SysUtils.DeleteFile(stFileName);
  loChartSpace.ExportPicture(stFileName, 'gif', FWidth, FHeight);

  if not FTobResponse.FieldExists('URL') then
    FTobResponse.AddChampSup('URL', false);
  FTobResponse.PutValue('URL', '/Portail/' + LookupCurrentSession.SessionId + '_owc.gif');

  loChartSpace := unassigned;
  CoUninitialize;
end;

procedure TOWCChart.ExportXMLData;
var
  stCaption, stValues, stVal, stCategories, stCat, stFileName: string;
  SeriesCount, CategoriesCount: integer;
  i, j: integer;
  T, TS, TC: tob;
  aString: string;
  aFileStream: TFileStream;

  ChartSpace1 : TChartSpace;
  oChart : chChart;
  oChart2 : chChart;
  oSeries1, oSeries2 : chSeries;
  oAxis : chAxis;
  ar : chAxis;
  sca1, sca2 : chScaling;
  dLab : chDataLabels;
  DeuxAxesDegroupes : Boolean;
  Min1, Min2 : Double;
  Max1, Max2 : Double;
  Pan : THPanel;
begin
  CoInitialize(nil);

(*
  oChart.HasSelectionMarks := False;
  oChart.AllowPropertyToolbox := False;
  oChart.DisplayPropertyToolbox := False;
*)
  pan := THPanel.Create(nil);
  try
      ddWriteLN ('UtoolsOWC : avant ChartSpace1 := TChartSpace.Create(Pan)');
    ChartSpace1 := TChartSpace.Create(Pan);
      ddWriteLN ('UtoolsOWC : après ChartSpace1 := TChartSpace.Create(Pan)');
    ChartSpace1.Clear;
    oChart := ChartSpace1.Charts.Add(0);

    { 1 / CHARGEMENT DES SÉRIES }
    CategoriesCount := FList.Detail.Count;
    SeriesCount := GetSerieCount;//FList.Detail[0].Detail.Count;

    stCategories := '';
    for i := 0 to CategoriesCount - 1 do
    begin
      stCat := FList.Detail[i].GetValue('CAPTION');
      if IsValidDate(stCat) then
        stCat := ' ' + stCat + ' '; {Alt + 0160 pour éviter l'interprétation des dates}
      if stCategories <> '' then
        stCategories := stCategories + ',';
      stCategories := stCategories + stCat;
    end;

    DeuxAxesDegroupes := (SeriesCount = 2) and owAxeRight.Visible and owIsAxeRightDegroupe;
    Min1 := 0; Min2 := 0; Max1 := 0; Max2 := 0;

    for i := 0 to SeriesCount - 1 do
    begin
      stValues := '';

      for j := 0 to CategoriesCount - 1 do
      begin
        TC := FList.Detail[j];
        {JP 12/04/06 : Il se peut que les séries n'aient pas toutes le même nombre de catégories}
        if TC.Detail.Count - 1 < i then Continue;

        TS := TC.Detail[i];
        stCaption := TS.GetValue('CAPTION');
        T := TC.Detail[i];

        {Calcul des échelles des deux axes si dégroupés}
        if DeuxAxesDegroupes then begin
          if i = 0 then begin
            Min1 := Min(Min1, Valeur(T.GetValue('VALUE')));
            Max1 := Max(Max1, Valeur(T.GetValue('VALUE')));
          end else begin
            Min2 := Min(Min2, Valeur(T.GetValue('VALUE')));
            Max2 := Max(Max2, Valeur(T.GetValue('VALUE')));
          end;
        end;

        stVal := Format('%f', [Valeur(T.GetValue('VALUE'))]);
        stVal := FindEtReplace(stVal, ',', '.', true);
        if stValues <> '' then
          stValues := stValues + ',';
        stValues := stValues + stVal;
      end;

      if (i = 1) and (DeuxAxesDegroupes or owHasTwoCharts) then begin
        if owHasTwoCharts then begin
          oChart2 := ChartSpace1.Charts.Add(1);
          oSeries2 := oChart2.SeriesCollection.Add(CategoriesCount);
        end
        else
          oSeries2 := oChart.SeriesCollection.Add(CategoriesCount);

        oSeries2.Caption := stCaption;
        oSeries2.SetData(chDimCategories, chDataLiteral, stCategories);
        oSeries2.SetData(chDimValues, chDataLiteral, stValues);
        {Si les deux axes dégroupés, il faut préciser le type de chaques série}
        if DeuxAxesDegroupes then
          oSeries2.Type_ := GetTypeChart(owAxeRight.InfoAxe.TypeSerie);

        if owLabel.Visible and owAxeRight.InfoAxe.HasLabel then begin
          dLab := oSeries2.DataLabelsCollection.Add;
          dLab.HasPercentage := owLabel.InfoAxe.LabelPerct;
          dLab.HasValue := owLabel.InfoAxe.LabelValue;
          dLab.Interior.Set_Color(GetCouleurGraph(owLabel.Color));
          dLab.Font.Set_Color(GetCouleurGraph(owLabel.FontColor));
          dLab.Font.Size  := owLabel.FontSize;
          dLab.Font.Name  := owLabel.FontName;
          dLab.NumberFormat := owLabel.InfoAxe.NumFormat;
        end;
      end
      else begin
        oSeries1 := oChart.SeriesCollection.Add(CategoriesCount);
        oSeries1.Caption := stCaption;
        oSeries1.SetData(chDimCategories, chDataLiteral, stCategories);
        oSeries1.SetData(chDimValues, chDataLiteral, stValues);
        {Si les deux axes dégroupés, il faut préciser le type de chaques série}
        if DeuxAxesDegroupes then
          oSeries1.Type_ := GetTypeChart(owAxeLeft.InfoAxe.TypeSerie);

        if owLabel.Visible and owAxeLeft.InfoAxe.HasLabel then begin
          dLab := oSeries1.DataLabelsCollection.Add;
          dLab.HasPercentage := owLabel.InfoAxe.LabelPerct;
          dLab.HasValue := owLabel.InfoAxe.LabelValue;
          dLab.Interior.Set_Color(GetCouleurGraph(owLabel.Color));
          dLab.Font.Set_Color(GetCouleurGraph(owLabel.FontColor));
          dLab.Font.Size  := owLabel.FontSize;
          dLab.Font.Name  := owLabel.FontName;
          dLab.NumberFormat := owLabel.InfoAxe.NumFormat;
        end;
      end;
    end;

    {Paramétrage des échelles si les deux axes sont dégroupés}
    if DeuxAxesDegroupes and not owHasTwoCharts then begin
      oSeries2.Ungroup(True);
      ar := oChart.Axes.Add(oSeries2.Get_Scalings(ChartDimensionsEnum (chDimValues)));
      ar.Position := ChartAxisPositionEnum(chAxisPositionRight);

      sca2 := oChart.Axes[chAxisPositionRight].Scaling;

      if Max2 <> 0 then sca2.Maximum := Max2 + Round(Abs(Max2) / 5)
                   else sca2.Maximum := 0;
      if Min2 <> 0 then sca2.Minimum := Min2 - Round(Abs(Min2) / 5)
                   else sca2.Minimum := 0;
      sca1 := oChart.Axes[chAxisPositionLeft].Scaling;
      if Max1 <> 0 then sca1.Maximum := Max1 + Round(Abs(Max1) / 5)
                   else sca1.Maximum := 0;
      if Min1 <> 0 then sca1.Minimum := Min1 - Round(Abs(Min1) / 5)
                   else sca1.Minimum := 0;
    end
    else if owAxeRight.Visible and not owIsAxeRightDegroupe then begin
      ar := oChart.Axes.Add(oSeries1.Get_Scalings(ChartDimensionsEnum (chDimValues)));
      ar.Position := ChartAxisPositionEnum(chAxisPositionRight);
    end;

    if not owIsAxeRightDegroupe then begin
      oChart.Type_ := GetTypeChart(FTypeGraph);
      if owHasTwoCharts then oChart2.Type_ := GetTypeChart(FTypeGraph);
    end;

    { 2 / PARAMÉTRAGE DE L'ESPACE CONTENANT LE GRAPH }
    SetInteriorsColor(oChart);
    if owHasTwoCharts then SetInteriorsColor(oChart2);

    { 3 / PARAMÉTRAGE DES LÉGENDES }
    SetGraphLegend(oChart);
    if owHasTwoCharts then SetGraphLegend(oChart2);

    { 4 / PARAMÉTRAGE DU TITRE }
    SetGraphTitle(oChart);
    owTitre.Caption := owTitreChart2;
    if owHasTwoCharts then SetGraphTitle(oChart2);

    { 7 / PARAMÉTRAGE DES AXES }
    if FTypeGraph <> tyCamembert then begin
      if owAxeLeft.Visible then begin
        oAxis := oChart.Axes[chAxisPositionLeft];
        SetGraphAxes(oAxis, owAxeLeft);
      end;

      if owAxeRight.Visible then begin
        oAxis := oChart.Axes[chAxisPositionRight];
        SetGraphAxes(oAxis, owAxeRight);
      end;

      if owAxeBot.Visible then begin
        oAxis := oChart.Axes[chAxisPositionBottom];
        SetGraphAxes(oAxis, owAxeBot);
      end;
    end;

    { 8 / TRAITEMENTS SPÉCIFIQUES }
    if Assigned(AddSpecific) then
      AddSpecific(oChart);
    if owHasTwoCharts and Assigned(AddSpecific2) then
      AddSpecific2(oChart2);

    { 9 / GÉNÉRATION DU FICHIER XML }
    stFileName := LookupCurrentSession.eAglDocPath +
      'Portail\' + LookupCurrentSession.SessionId + '_owc.xml';
    if FileExists(stFileName) then
      SysUtils.DeleteFile(stFileName);

    ChartSpace1.XMLData := ''; {On vide XMLData pour éviter toute pollution de Microsoft}
    ChartSpace1.Update; {Chargement de la propriété XMLData}
    aString := ChartSpace1.XMLData;
    aFileStream := nil;
    try
      aFileStream := TFileStream.Create(stFileName, fmCreate);
      aFileStream.Write(PChar(aString)^, Length(aString));
    finally
      aFileStream.Free;
    end;

    if not FTobResponse.FieldExists('XMLDATA') then
      FTobResponse.AddChampSup('XMLDATA', false);
    FTobResponse.PutValue('XMLDATA', astring);

    if not FTobResponse.FieldExists('URL') then
      FTobResponse.AddChampSup('URL', false);
    FTobResponse.PutValue('URL', '/Portail/' + LookupCurrentSession.SessionId + '_owc.xml');

    {JP/EPZ 26/01/07 : Nouvelle gestion des graphiques}
    ExportPicture(AString);
  finally
    FreeAndNil(Pan);
  end;
//  ChartSpace1 := unassigned;
  CoUninitialize;
end;

procedure TOWCChart.SetWidth(AWidth: integer);
begin
  FWidth := AWidth;
end;

procedure TOWCChart.SetHEight(AHeight: integer);
begin
  FHeight := AHeight;
end;

procedure TOWCChart.SetTypeGraph(const Value: TTypeGraph);
begin
  FTypeGraph := Value;
end;

procedure TOWCChart.SetGraphLegend(var oChart : chChart);
begin
  if not Assigned(oChart) or not owLegend.Visible then Exit;

  oChart.HasLegend := True;
  oChart.Legend.Border.Set_Color('white');
  oChart.Legend.Position := GetPositionI(owLegend.Position);
  oChart.Legend.Font.Name := owLegend.FontName;
  oChart.Legend.Font.Size := owLegend.FontSize;
  oChart.Legend.Font.Set_Color(GetCouleurGraph(owLegend.FontColor));
  oChart.Legend.Interior.Set_Color(GetCouleurGraph(owTitre.Color));
  if owLegend.Bold then
    oChart.Legend.Font.Bold := True;
  if owLegend.Italic then
    oChart.Legend.Font.Italic := True;
end;

procedure TOWCChart.SetGraphTitle(var oChart : chChart);
begin
  if not Assigned(oChart) then Exit;

  if owTitre.Visible and (owTitre.Caption <> '') then begin
    oChart.HasTitle := True;
    oChart.Title.Border.Set_Color('White');
    oChart.Title.Caption := owTitre.Caption;
    oChart.Title.Position := GetPositionI(owTitre.Position);
    oChart.Title.Font.Name := owTitre.FontName;
    oChart.Title.Font.Size := owTitre.FontSize;
    oChart.Title.Interior.Set_Color(GetCouleurGraph(owTitre.Color));
    oChart.Title.Font.Set_Color(GetCouleurGraph(owTitre.FontColor));
    if owTitre.Bold then
      oChart.Title.Font.Bold := True;
    if owTitre.Italic then
      oChart.Title.Font.Italic := True;
  end
  else
    oChart.HasTitle := False;
end;

procedure TOWCChart.SetInteriorsColor(var oChart : chChart);
begin
  if not Assigned(oChart) then Exit;

  oChart.Border.Set_Color('white');

  if FTypeGraph <> tyCamembert then begin
    oChart.Interior.Set_Color('White');
    if owPlotColor <> colNone then
      oChart.PlotArea.Interior.Set_Color(GetCouleurGraph(owPlotColor));
    if owContourColor <> colNone then
      oChart.Interior.Set_Color(GetCouleurGraph(owContourColor));

    oChart.Axes[chAxisPositionLeft].HasMajorGridlines := owHasMajorLine;
    oChart.Axes[chAxisPositionLeft].HasMinorGridlines := owHasMinorLine;

    if owAxeRight.Visible then begin
      oChart.Axes[chAxisPositionRight].HasMajorGridlines := owHasMajorLine;
      oChart.Axes[chAxisPositionRight].HasMinorGridlines := owHasMinorLine;
    end;
  end;
end;

procedure TOWCChart.SetGraphAxes(var oAxes : ChAxis; owObjAxe : TObjText);
begin
  if  not Assigned(oAxes) then Exit;

  if (owObjAxe.Caption <> '') and owObjAxe.Visible then begin
    oAxes.HasTitle := True;
    oAxes.Title.Caption := owObjAxe.Caption;
    oAxes.Title.Font.Name := owObjAxe.FontName;
    oAxes.Title.Font.Size := owObjAxe.FontSize;
    if owObjAxe.Bold then
      oAxes.Title.Font.Bold := True;
    if owObjAxe.Italic then
      oAxes.Title.Font.Italic := True;
  end
  else
    oAxes.HasTitle := False;

  if Trim(owObjAxe.InfoAxe.NumFormat) <> '' then
    oAxes.NumberFormat := owObjAxe.InfoAxe.NumFormat;
  if Trim(owObjAxe.InfoAxe.MajorUnit) <> '' then
    oAxes.MajorUnit := Valeur(owObjAxe.InfoAxe.MajorUnit);
end;

procedure TOWCChart.SetDefaultValue;
begin
  owLegend.Position := poPositionAutomatic;
  owLegend.FontName := 'Trebuchet MS';
  owLegend.FontSize := 8;
  owLegend.FontColor := colBlack;
  owLegend.Color := colWhite;
  owLegend.Visible := True;

  owTitre.Position := poPositionAutomatic;
  owTitre.FontName := 'Trebuchet MS';
  owTitre.FontSize := 8;
  owTitre.Visible := True;
  owTitre.FontColor := colBlack;
  owTitre.Color := colWhite;

  owContourColor := colWhite;
  owPlotColor := colWhite;

  owHasTwoCharts := False;
  owIsAxeRightDegroupe := False;

  owAxeBot.Visible := True;
  owAxeBot.FontName := 'Trebuchet MS';
  owAxeBot.FontSize := 8;

  owAxeLeft.Visible := True;
  owAxeLeft.FontName := 'Trebuchet MS';
  owAxeLeft.FontSize := 8;

  owAxeRight.Visible := False;
  owAxeRight.FontName := 'Trebuchet MS';
  owAxeRight.FontSize := 8;

  owAxeLeft.Bold := False;
  owAxeBot.Bold := False;
  owAxeRight.Bold := False;
  owAxeLeft.Color := ColBlack;
  owAxeBot.Color := ColBlack;
  owAxeRight.Color := ColBlack;
  owAxeBot.InfoAxe.NumFormat := '';
  owAxeLeft.InfoAxe.NumFormat := '';
  owAxeRight.InfoAxe.NumFormat := '';
  owAxeBot.InfoAxe.MajorUnit := '';
  owAxeLeft.InfoAxe.MajorUnit := '';
  owAxeRight.InfoAxe.MajorUnit := '';
  owAxeLeft.InfoAxe.HasLabel := True;
  owAxeRight.InfoAxe.HasLabel := True;

  owAxeRight.FontColor := colBlack;
  owAxeLeft.FontColor := colBlack;
  owAxeBot.FontColor := colBlack;
  owLegend.FontColor := colBlack;
  owTitre.FontColor := colBlack;

  {Si les deux axes sont dégroupés, les 2 séries se superposent, mieux vaut qu'elles
   soient de type différent}
  owAxeRight.InfoAxe.TypeSerie := tyLigne;
  owAxeLeft.InfoAxe.TypeSerie := tyHistogramme;

  owLabel.Visible := False;
  owLabel.FontName := 'Trebuchet MS';
  owLabel.FontSize := 8;
  owLabel.InfoAxe.NumFormat := '0';
  owLabel.InfoAxe.LabelPerct := False;
  owLabel.InfoAxe.LabelValue := True;
  owLabel.Color := colYellow;
  owLabel.FontColor := colBlack;
end;

class function TOWCChart.GetCouleurGraph(Co: TCouleurGraph): string;
begin
  Result := '';
  case Co of
    ColAliceBlue: Result := 'AliceBlue';
    ColAntiqueWhite: Result := 'AntiqueWhite';
    ColAqua: Result := 'Aqua';
    ColAquamarine: Result := 'Aquamarine';
    ColAzure: Result := 'Azure';
    ColBeige: Result := 'Beige';
    ColBisque: Result := 'Bisque';
    ColBlack: Result := 'Black';
    ColBlanchedAlmond: Result := 'BlanchedAlmond';
    ColBlue: Result := 'Blue';
    ColBlueViolet: Result := 'BlueViolet';
    ColBrown: Result := 'Brown';
    ColBurlyWood: Result := 'BurlyWood';
    ColCadetBlue: Result := 'CadetBlue';
    ColChartreuse: Result := 'Chartreuse';
    ColChocolate: Result := 'Chocolate';
    ColCoral: Result := 'Coral';
    ColCornflower: Result := 'Cornflower';
    ColCornsilk: Result := 'Cornsilk';
    ColCrimson: Result := 'Crimson';
    ColCyan: Result := 'Cyan';
    ColDarkBlue: Result := 'DarkBlue';
    ColDarkCyan: Result := 'DarkCyan';
    ColDarkGoldenrod: Result := 'DarkGoldenrod';
    ColDarkGray: Result := 'DarkGray';
    ColDarkGreen: Result := 'DarkGreen';
    ColDarkKhaki: Result := 'DarkKhaki';
    ColDarkMagenta: Result := 'DarkMagenta';
    ColDarkOliveGreen: Result := 'DarkOliveGreen';
    ColDarkOrange: Result := 'DarkOrange';
    ColDarkOrchid: Result := 'DarkOrchid';
    ColDarkRed: Result := 'DarkRed';
    ColDarkSalmon: Result := 'DarkSalmon';
    ColDarkSeaGreen: Result := 'DarkSeaGreen';
    ColDarkSlateBlue: Result := 'DarkSlateBlue';
    ColDarkSlateGray: Result := 'DarkSlateGray';
    ColDarkTurquoise: Result := 'DarkTurquoise';
    ColDarkViolet: Result := 'DarkViolet';
    ColDeepPink: Result := 'DeepPink';
    ColDeepSkyBlue: Result := 'DeepSkyBlue';
    ColDimGray: Result := 'DimGray';
    ColDodgerBlue: Result := 'DodgerBlue';
    ColFirebrick: Result := 'Firebrick';
    ColFloralWhite: Result := 'FloralWhite';
    ColForestGreen: Result := 'ForestGreen';
    ColFuchsia: Result := 'Fuchsia';
    ColGainsboro: Result := 'Gainsboro';
    ColGhostWhite: Result := 'GhostWhite';
    ColGold: Result := 'Gold';
    ColGoldenrod: Result := 'Goldenrod';
    ColGray: Result := 'Gray';
    ColGreen: Result := 'Green';
    ColGreenYellow: Result := 'GreenYellow';
    ColHoneydew: Result := 'Honeydew';
    ColHotPink: Result := 'HotPink';
    ColIndianRed: Result := 'IndianRed';
    ColIndigo: Result := 'Indigo';
    ColIvory: Result := 'Ivory';
    ColKhaki: Result := 'Khaki';
    ColLavender: Result := 'Lavender';
    ColLavenderBlush: Result := 'LavenderBlush';
    ColLawnGreen: Result := 'LawnGreen';
    ColLemonChiffon: Result := 'LemonChiffon';
    ColLightBlue: Result := 'LightBlue';
    ColLightCoral: Result := 'LightCoral';
    ColLightCyan: Result := 'LightCyan';
    ColLightGoldenrodYellow: Result := 'LightGoldenrodYellow';
    ColLightGray: Result := 'LightGray';
    ColLightGreen: Result := 'LightGreen';
    ColLightPink: Result := 'LightPink';
    ColLightSalmon: Result := 'LightSalmon';
    ColLightSeaGreen: Result := 'LightSeaGreen';
    ColLightSkyBlue: Result := 'LightSkyBlue';
    ColLightSlateGray: Result := 'LightSlateGray';
    ColLightSteelBlue: Result := 'LightSteelBlue';
    ColLightYellow: Result := 'LightYellow';
    ColLime: Result := 'Lime';
    ColLimeGreen: Result := 'LimeGreen';
    ColLinen: Result := 'Linen';
    ColMagenta: Result := 'Magenta';
    ColMaroon: Result := 'Maroon';
    ColMediumAquamarine: Result := 'MediumAquamarine';
    ColMediumBlue: Result := 'MediumBlue';
    ColMediumOrchid: Result := 'MediumOrchid';
    ColMediumPurple: Result := 'MediumPurple';
    ColMediumSeaGreen: Result := 'MediumSeaGreen';
    ColMediumSlateBlue: Result := 'MediumSlateBlue';
    ColMediumSpringGreen: Result := 'MediumSpringGreen';
    ColMediumTurquoise: Result := 'MediumTurquoise';
    ColMediumVioletRed: Result := 'MediumVioletRed';
    ColMidnightBlue: Result := 'MidnightBlue';
    ColMintCream: Result := 'MintCream';
    ColMistyRose: Result := 'MistyRose';
    ColMoccasin: Result := 'Moccasin';
    ColNavajoWhite: Result := 'NavajoWhite';
    ColNavy: Result := 'Navy';
    ColOldLace: Result := 'OldLace';
    ColOlive: Result := 'Olive';
    ColOliveDrab: Result := 'OliveDrab';
    ColOrange: Result := 'Orange';
    ColOrangeRed: Result := 'OrangeRed';
    ColOrchid: Result := 'Orchid';
    ColPaleGoldenrod: Result := 'PaleGoldenrod';
    ColPaleGreen: Result := 'PaleGreen';
    ColPaleTurquoise: Result := 'PaleTurquoise';
    ColPaleVioletRed: Result := 'PaleVioletRed';
    ColPapayaWhip: Result := 'PapayaWhip';
    ColPeachPuff: Result := 'PeachPuff';
    ColPeru: Result := 'Peru';
    ColPink: Result := 'Pink';
    ColPlum: Result := 'Plum';
    ColPowderBlue: Result := 'PowderBlue';
    ColPurple: Result := 'Purple';
    ColRed: Result := 'Red';
    ColRosyBrown: Result := 'RosyBrown';
    ColRoyalBlue: Result := 'RoyalBlue';
    ColSaddleBrown: Result := 'SaddleBrown';
    ColSalmon: Result := 'Salmon';
    ColSandyBrown: Result := 'SandyBrown';
    ColSeaGreen: Result := 'SeaGreen';
    ColSeaShell: Result := 'SeaShell';
    ColSienna: Result := 'Sienna';
    ColSilver: Result := 'Silver';
    ColSkyBlue: Result := 'SkyBlue';
    ColSlateBlue: Result := 'SlateBlue';
    ColSlateGray: Result := 'SlateGray';
    ColSnow: Result := 'Snow';
    ColSpringGreen: Result := 'SpringGreen';
    ColSteelBlue: Result := 'SteelBlue';
    ColTan: Result := 'Tan';
    ColTeal: Result := 'Teal';
    ColThistle: Result := 'Thistle';
    ColTomato: Result := 'Tomato';
    ColTransparent: Result := 'Transparent';
    ColTurquoise: Result := 'Turquoise';
    ColViolet: Result := 'Violet';
    ColWheat: Result := 'Wheat';
    ColWhite: Result := 'White';
    ColWhiteSmoke: Result := 'WhiteSmoke';
    ColYellow: Result := 'Yellow';
    ColYellowGreen: Result := 'YellowGreen';
  end;
end;

function TOWCChart.GetPositionI(Po : TPositionText) : Integer;
begin
  Result := chTitlePositionAutomatic;
  {Pour les légendes et les titres des graphs}
  case Po of
    poPositionAutomatic : Result := chTitlePositionAutomatic;
    poPositionBottom    : Result := chTitlePositionBottom;
    poPositionLeft      : Result := chTitlePositionLeft;
    poPositionRight     : Result := chTitlePositionRight;
    poPositionTop       : Result := chTitlePositionTop;
  end;
end;

function TOWCChart.GetTypeChart(Typ : TTypeGraph) : ChartChartTypeEnum;
begin
  Result := chChartTypeLine;
  if FAspectGraph = '2D' then
    case Typ of
      tyCamembert     : Result := chChartTypePie;
      tyHistogramme   : Result := chChartTypeColumnClustered;
      tyBar           : Result := chChartTypeBarClustered;
      tyLigne         : Result := chChartTypeLine;
    end
  else
    case Typ of
      tyCamembert     : Result := chChartTypePie3D;
      tyHistogramme   : Result := chChartTypeColumnClustered3D;
      tyBar           : Result := chChartTypeBar3D;
      tyLigne         : Result := chChartTypeLine3D;
    end;
//    tyBubble        : Result := chChartTypePie3D;
  //  tyHistogramme3D : Result := chChartTypeColumnClustered3D;
//    tyBar3D         : Result := chChartTypeBar3D;
  //  tyCamembert3D   : Result := chChartTypePie3D;
end;

end.

