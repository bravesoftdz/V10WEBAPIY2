unit uToolsOWC;

{$IFNDEF EAGLSERVER} abghgahgh {$ENDIF}

interface

uses classes, uTob, sysutils;

const
  OWCChartClsId = 'CLSID:0002E55D-0000-0000-C000-000000000046';
  OWCSpreadSheetClsId = 'CLSID:0002E559-0000-0000-C000-000000000046';
  OWCPivotTableClsId = 'CLSID:0002E55A-0000-0000-C000-000000000046';

  chChartTypeColumnClustered = $00000000;
  chChartTypeBarClustered = $00000003;
  chChartTypeLine = $00000006;
  chChartTypePie = $00000012;
  chChartTypeBubble = $0000001B;
  chChartTypeColumn3D = $0000002E;
  chChartTypeBar3D = $00000032;
  chChartTypePie3D = $0000003A;

  chAxisPositionBottom = -2;
  chAxisPositionCategory = -7;
  chAxisPositionCircular = -6;
  chAxisPositionLeft = -3;
  chAxisPositionPrimary = -10;
  chAxisPositionRadial = -5;
  chAxisPositionRight = -4;
  chAxisPositionSecondary = -11;
  chAxisPositionSeries = -9;
  chAxisPositionTimescale = -7;
  chAxisPositionTop = -1;
  chAxisPositionValue = -8;

(*
type TOWCSpreadSheet = Class(TObject)
     private
       FSheetCount: integer;
       FXlsFileName: string;
     public
       constructor Create(TobResponse: TOB);
       destructor Destroy; override;
       procedure SetValue(ACategoryName, ASerieName: string; AValue: variant);
       procedure Build;
       property SheetCount: integer read FSheetCount write FSheetCount;
       property XlsFileName: string read FXlsFileName write FXlsFileName;
     end;
*)

type
  TTypeGraph = (tyCamembert, tyHistogramme, tyBar, tyLigne);
  TPositionText = (poPositionAutomatic, poPositionBottom, poPositionLeft, poPositionRight, poPositionTop);
  TCouleurGraph = (colNone, colYellow, colGreen, colBlue, colSienna, colPurple, colRed, colWhite, colBlack,
                   colOrange, colDarkSalmon, ColGray, colSalmon);

  TObjAxeParam = record
    NumFormat  : string;
    MajorUnit  : string;
  end;

  TObjText = record
    Visible   : Boolean;
    Position  : TPositionText;
    Bold      : Boolean;
    Color     : TCouleurGraph;
    Italic    : Boolean;
    FontName  : string;
    FontSize  : Byte;
    Caption   : string;
    InfoAxe   : TObjAxeParam;
  end;

  TSetColorSeries = procedure (var L : TStringList) of object;

     TOWCChart = Class(TObject)
     private
       FList       : TOB;
       FTobResponse: TOB;
       FWidth     : integer;
       FHeight    : integer;
       FTypeGraph : TTypeGraph;
       FTitle     : string;
       function  GetPosition      (Po : TPositionText) : string;
       function  GetCouleurGraph  (Co : TCouleurGraph) : string;
       function  GetTypeGraph     : string;
       procedure SetTypeGraph     (const Value : TTypeGraph);
       procedure SetGraphLegend   (var L : TStringList);
       procedure SetGraphTitle    (var L : TStringList);
       procedure SetGraphAxeLeft  (var L : TStringList);
       procedure SetGraphAxeBottom(var L : TStringList);
       procedure SetInteriorsColor(var L : TStringList);
       procedure SetDefaultValue;
     public
       owTitre  : TObjText;
       owLegend : TObjText;
       owAxeLeft: TObjText;
       owAxeBot : TObjText;
       owContourColor : TCouleurGraph;
       owPlotColor    : TCouleurGraph;
       owHasMajorLine : Boolean;
       owHasMinorLine : Boolean;
       SetSpecific    : TSetColorSeries;

       constructor Create(TobResponse: TOB);
       destructor Destroy; override;
       procedure SetTitle(ATitle: string);
       procedure SetValue(ACategoryName, ASerieName: string; AValue: variant);
       procedure Build;
       procedure ExportPicture;
       procedure SetWidth(AWidth: integer);
       procedure SetHEight(AHeight: integer);
       property TypeGraph : TTypeGraph read FTypeGraph write SetTypeGraph;
     end;

type TOWCPivotTable = Class(TObject)
     end;


implementation

uses
  windows,activeX,comobj,hent1,esession;

{ TOWCChart }

constructor TOWCChart.Create(TobResponse: TOB);
begin
  FList := TOB.Create('the chart',nil,-1);
  FTobResponse := TobResponse;
  FWidth:=450;
  FHeight:=200;
  {Initialisation des variables pour dessiner le graph}
  SetDefaultValue;
end;

destructor TOWCChart.Destroy;
begin
FreeAndNil(FList);
end;

procedure TOWCChart.Build;
var L: TStringList;
    i,j: integer;
    ValuesCount,SeriesCount,CategoriesCount: integer;
    T,TC,TS: TOB;
    stFileName,stVal: string;
begin
  L := TStringList.Create;

  CategoriesCount := FList.Detail.Count;
  ValuesCount := FList.Detail.Count;
  SeriesCount := FList.Detail[0].Detail.Count;

  L.Add('<html>');
  L.Add('<body oncontextmenu="return false">');
  L.Add('<object id=ChartSpace1 classid='+OWCChartClsId+
        ' style="width:100%;height:'+IntToStr(FHeight)+'"></object>');
  L.Add('');
  L.Add('<script language=vbs>');
  L.Add('');

(*
Sub ChartSpace1_BeforeContextMenu(x, y, Menu, Cancel)
    Dim cmContextMenu(4)
    Dim cmClearSubMenu(2)

    cmClearSubMenu(0) = Array("&All", "ClearAll")
    cmClearSubMenu(1) = Array("&Formats", "ClearFormats")
    cmClearSubMenu(2) = Array("&Values", "ClearValues")

    cmContextMenu(0) = Array("Cu&t", "owc2")
    cmContextMenu(1) = Array("&Copy", "owc3")
    cmContextMenu(2) = Array("&Paste", "owc4")
    cmContextMenu(3) = Empty
    cmContextMenu(4) = Array("Clea&r", cmClearSubMenu)
    Menu.Value = cmContextMenu
End Sub
*)

  L.Add(''' Supprime le popupmenu du composant');
  L.Add('Sub ChartSpace1_BeforeContextMenu(x, y, Menu, Cancel)');
  L.Add('    Menu.Value = nothing');
  L.Add('End Sub');
  L.Add('');

  L.Add('Sub Window_OnLoad()');
  L.Add('');
  L.Add('Dim categories('+IntToStr(CategoriesCount-1)+')');
  L.Add('Dim values('+IntToStr(ValuesCount-1)+')');
  L.Add('Dim chConstants');
  L.Add(''' Create an array of strings representing the categories.');
  L.Add(''' The categories will be the same for all three series.');
  L.Add('');

  for i := 0 to FList.Detail.Count - 1 do begin
    if IsValidDate(FList.Detail[i].GetString('CAPTION')) then
      L.Add('categories(' + IntToStr(i) + ') = #' + FList.Detail[i].GetValue('CAPTION') + '#')
    else if IsNumeric(FList.Detail[i].GetString('CAPTION'), True) then
      L.Add('categories(' + IntToStr(i) + ') = ' + FList.Detail[i].GetValue('CAPTION'))
    else
      L.Add('categories(' + IntToStr(i) + ') = "' + FList.Detail[i].GetValue('CAPTION') + '"');
  end;
  L.Add('');

  L.Add(''' Clear the contents of the chart workspace. This removes');
  L.Add(''' any old charts that may already exist and leaves the chart workspace');
  L.Add(''' completely empty. One chart object is then added.');
  L.Add('');

  L.Add('ChartSpace1.Clear');
  L.Add('ChartSpace1.Charts.Add');
  L.Add('ChartSpace1.Charts(0).Border.Color = "white"');
  L.Add('');
  L.Add('Set chConstants = ChartSpace1.Constants');
  L.Add('');
  L.Add(''' Add  series to the chart.');
  L.Add('');

  for i:=0 to SeriesCount-1 do
    L.Add('ChartSpace1.Charts(0).SeriesCollection.Add');
  L.Add('');

  for i:=0 to SeriesCount-1 do
  begin
    TS := nil;
    for j:=0 to CategoriesCount-1 do
    begin
      TC := FList.Detail[j];
      TS := TC.Detail[i];
      T := TC.Detail[i];
      stVal := Format('%f',[Valeur(T.GetValue('VALUE'))]);
      stVal := FindEtReplace(stVal,',','.',true);
      L.Add('values('+IntToStr(j)+') = '+stVal);
    end;
    L.Add('ChartSpace1.Charts(0).SeriesCollection('+IntToStr(i)+').Caption = "'+TS.GetValue('CAPTION')+'"');
    L.Add('ChartSpace1.Charts(0).SeriesCollection('+IntToStr(i)+').SetData chConstants.chDimCategories, chConstants.chDataLiteral, categories');
    L.Add('ChartSpace1.Charts(0).SeriesCollection('+IntToStr(i)+').SetData chConstants.chDimValues, chConstants.chDataLiteral, values');
    L.Add('');
  end;

  L.Add('ChartSpace1.Charts(0).Type = chConstants.' + GetTypeGraph);
  L.Add(''' Make the chart legend visible, format the left value axis as percentage,');
  L.Add(''' and specify that value gridlines are at 10% intervals.');
  L.Add('');
  L.Add('''Show selection marks for individual chart elements.');
  L.Add('Chartspace1.HasSelectionMarks = False');
  L.Add('');

  L.Add('''Allow the user to display the Commands and Options dialog box.');
  L.Add('Chartspace1.AllowPropertyToolbox = False');
  L.Add('Chartspace1.DisplayPropertyToolbox = False');
  L.Add('');

  SetInteriorsColor(L);
  L.Add('');
  SetGraphTitle(L);
  L.Add('');
  SetGraphLegend(L);
  L.Add('');
  SetGraphAxeLeft(L);
  L.Add('');
  SetGraphAxeBottom(L);


  L.Add('');
  L.Add('End Sub');
  L.Add('</script>');
  L.Add('</body>');
  L.Add('</html>');

  {$IFDEF EAGLSERVER}
  stFileName := 'Portail\'+LookupCurrentSession.SessionId + '_owc.html';
  {$ELSE}
  {$ENDIF}
  if FileExists(LookupCurrentSession.eAglDocPath + stFileName) then
    Sysutils.DeleteFile(LookupCurrentSession.eAglDocPath + stFileName);
  L.SaveToFile(LookupCurrentSession.eAglDocPath + stFileName);
  L.Free;

  if not FTobResponse.FieldExists('URL') then FTobResponse.AddChampSup('URL',FALSE);
  FTobResponse.PutValue('URL','/Portail/' + LookupCurrentSession.SessionId + '_owc.html');
end;

procedure TOWCChart.SetValue(ACategoryName, ASerieName: string; AValue: variant);
var TC, TS: TOB;
    ok: boolean;
begin
//on cherche la tob de la categorie
TC := FList.FindFirst(['CAPTION'],[ACategoryName],false);
ok := TC<>nil;
//si pas trouvé, on la crée
if not ok then
begin
  TC := TOB.Create('the category',FList,-1);
  TC.AddChampSupValeur('CAPTION',ACategoryName);
end;
//on cherche la tob de la serie
TS := nil;
if ok then TS := TC.FindFirst(['CAPTION'],[ASerieName],false);
ok := TS<>nil;
//si pas trouvé, on la crée
if not ok then
begin
  TS := TOB.Create('the serie',TC,-1);
  TS.AddChampSupValeur('CAPTION',ASerieName);
end;
TS.AddChampSupValeur('VALUE',AValue);
end;

procedure TOWCChart.SetTitle(ATitle: string);
begin
FTitle := ATitle;
end;

procedure TOWCChart.ExportPicture;
var oConst,loSerie,loGraph,loChartSpace: OleVariant;
    stCaption,stValues, stVal, stCategories, stCat, stFileName: string;
    SeriesCount,CategoriesCount: integer;
    i,j: integer;
    T, TS, TC: tob;
begin
  CoInitialize(nil);

  loChartSpace := CreateOleObject('OWC11.ChartSpace');
  oConst := loChartSpace.Constants;
  loGraph := loChartSpace.Charts.Add();

  ddWriteLN('PGITRGRAPHSOLDE : Fin du chargement des écritures');
//  loGraph.Type := FTypeGraph;//chChartTypePie;
  if FTitle<>'' then
  begin
    loGraph.HasTitle := True;
    loGraph.Title.Caption := FTitle;
    loGraph.Title.Font.Name := '"verdana"';
    loGraph.Title.Font.Size := 8;
  end;

  loGraph.HasLegend := True;
  loGraph.Legend.Font.Name := '"verdava"';
  loGraph.Legend.Font.Size := 8;

  loGraph.PlotArea.Interior.Color := 'White';

  CategoriesCount := FList.Detail.Count;
  SeriesCount := FList.Detail[0].Detail.Count;

  stCategories := '';
  for i:=0 to CategoriesCount-1 do
  begin
    stCat := FList.Detail[i].GetValue('CAPTION');
    if stCategories<>'' then stCategories := stCategories + ',';
    stCategories := stCategories + stCat;
  end;

  for i:=0 to SeriesCount-1 do
  begin
    stValues := '';
    loSerie := loGraph.SeriesCollection.Add();
    for j:=0 to CategoriesCount-1 do
    begin
      TC := FList.Detail[j];
      TS := TC.Detail[i];
      stCaption := TS.GetValue('CAPTION');
      T := TC.Detail[i];
      stVal := Format('%f',[Valeur(T.GetValue('VALUE'))]);
      stVal := FindEtReplace(stVal,',','.',true);
      if stValues<>'' then stValues := stValues + ',';
      stValues := stValues + stVal;
    end;
    loSerie.Caption := stCaption;
    loSerie.SetData(oConst.chDimCategories, oConst.chDataLiteral, stCategories);
    loSerie.SetData(oConst.chDimValues, oConst.chDataLiteral, stValues);
  end;

  loGraph.Axes(0).Font.Name := '"Verdana"';
  loGraph.Axes(0).Font.Size := 8;

  loGraph.Axes(1).Font.Name := '"Verdana"';
  loGraph.Axes(1).Font.Size := 8;

  stFileName := LookupCurrentSession.eAglDocPath +
                'Portail\'+LookupCurrentSession.SessionId + '_owc.gif';
  if FileExists(stFileName) then
    SysUtils.DeleteFile(stFileName);
  loChartSpace.ExportPicture(stFileName, 'gif', FWidth, FHeight);

  if not FTobResponse.FieldExists ('URL') then
    FTobResponse.AddChampSup('URL',false);
  FTobResponse.PutValue('URL','/Portail/'+LookupCurrentSession.SessionId + '_owc.gif');

  loChartSpace := unassigned;
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

procedure TOWCChart.SetTypeGraph(const Value : TTypeGraph);
begin
  FTypeGraph := Value;
end;

procedure TOWCChart.SetGraphLegend(var L : TStringList);
begin
  if not Assigned(L) or not owLegend.Visible then Exit;

  L.Add('ChartSpace1.HasChartSpaceLegend = True');
  L.Add('ChartSpace1.ChartSpaceLegend.Font.Name = "' + owLegend.FontName + '"');
  L.Add('ChartSpace1.ChartSpaceLegend.Font.Size = ' + IntToStr(owLegend.FontSize));
  L.Add('ChartSpace1.ChartSpaceLegend.Font.Color = "' + GetCouleurGraph(owLegend.Color) + '"');
  L.Add('ChartSpace1.ChartSpaceLegend.Position = ' + GetPosition(owLegend.Position));
  if owLegend.Bold   then L.Add('ChartSpace1.ChartSpaceLegend.Font.Bold = True');
  if owLegend.Italic then L.Add('ChartSpace1.ChartSpaceLegend.Font.Italic = True');
end;

procedure TOWCChart.SetGraphTitle(var L : TStringList);
begin
  if not Assigned(L) or not owTitre.Visible or (owTitre.Caption = '') then Exit;
  L.Add('ChartSpace1.HasChartSpaceTitle = True');
  L.Add('ChartSpace1.ChartSpaceTitle.Caption = "' + owTitre.Caption + '"');
  L.Add('ChartSpace1.ChartSpaceTitle.Font.Name = "' + owTitre.FontName + '"');
  L.Add('ChartSpace1.ChartSpaceTitle.Font.Size = ' + IntToStr(owTitre.FontSize));
  L.Add('ChartSpace1.ChartSpaceTitle.Font.Color = "' + GetCouleurGraph(owTitre.Color) + '"');
  L.Add('ChartSpace1.ChartSpaceTitle.Position = ' + GetPosition(owTitre.Position));
  if owTitre.Bold   then L.Add('ChartSpace1.ChartSpaceTitle.Font.Bold = True');
  if owTitre.Italic then L.Add('ChartSpace1.ChartSpaceTitle.Font.Italic = True');
end;

procedure TOWCChart.SetInteriorsColor(var L : TStringList);
begin
  if FTypeGraph = tyCamembert then Exit;

  if owPlotColor <> colNone then
    L.Add('ChartSpace1.Charts(0).PlotArea.Interior.Color = "' + GetCouleurGraph(owPlotColor) + '"');
  if owContourColor <> colNone then
    L.Add('ChartSpace1.Charts(0).Interior.Color = "' + GetCouleurGraph(owContourColor) + '"');
  {Par défaut à True}
  if not owHasMajorLine then
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).HasMajorGridlines = False');
  {Par défaut à False}
  if owHasMinorLine then
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).HasMinorGridlines = True');

  if Assigned(SetSPecific) then SetSPecific(L);
end;

procedure TOWCChart.SetGraphAxeLeft  (var L : TStringList);
begin
  if FTypeGraph = tyCamembert then Exit;

  if owAxeLeft.Caption <> '' then begin
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).HasTitle = True');
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).Title.Caption = "' + owAxeLeft.Caption + '"');
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).Title.Font.Name = "' + owAxeLeft.FontName + '"');
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).Title.Font.Size = ' + IntToStr(owAxeLeft.FontSize));
    if owAxeLeft.Bold then
      L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).Font.Bold = True');
    if owAxeLeft.Italic then
      L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).Font.Italic = True');

  end
  else
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).HasTitle = False');
  if owAxeLeft.InfoAxe.NumFormat <> '' then
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).NumberFormat = "' + owAxeLeft.InfoAxe.NumFormat + '"');

  if owAxeLeft.InfoAxe.MajorUnit <> '' then
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionLeft).MajorUnit = ' + owAxeLeft.InfoAxe.MajorUnit);
end;

procedure TOWCChart.SetGraphAxeBottom(var L : TStringList);
begin
  if FTypeGraph = tyCamembert then Exit;

  if owAxeBot.Caption <> '' then begin
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).HasTitle = True');
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).Title.Caption = "' + owAxeBot.Caption + '"');
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).Title.Font.Name = "' + owAxeBot.FontName + '"');
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).Title.Font.Size = ' + IntToStr(owAxeBot.FontSize));
    if owAxeBot.Bold then
      L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).Font.Bold = True');
    if owAxeBot.Italic then
      L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).Font.Italic = True');
  end
  else
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).HasTitle = False');

  if owAxeLeft.InfoAxe.NumFormat <> '' then
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).NumberFormat = "' + owAxeBot.InfoAxe.NumFormat + '"');

  if owAxeBot.InfoAxe.MajorUnit <> '' then
    L.Add('ChartSpace1.Charts(0).Axes(chConstants.chAxisPositionBottom).MajorUnit = ' + owAxeBot.InfoAxe.MajorUnit);
end;


procedure TOWCChart.SetDefaultValue;
begin
  owLegend.Position := poPositionAutomatic;
  owLegend.FontName := 'verdana';
  owLegend.FontSize := 8;
  owLegend.Color    := colBlack;
  owLegend.Visible  := True;
  
  owTitre.Position := poPositionAutomatic;
  owTitre.FontName := 'verdana';
  owTitre.FontSize := 10;
  owTitre.Visible  := True;
  owTitre.Color    := colBlack;

  owContourColor := colWhite;
  owPlotColor    := colWhite;

  owAxeBot.Visible := True;
  owAxeBot.FontName := 'Verdana';
  owAxeBot.FontSize := 8;

  owAxeLeft.Visible := True;
  owAxeLeft.FontName := 'Verdana';
  owAxeLeft.FontSize := 8;

  owAxeLeft.Bold := False;
  owAxeBot.Bold := False;
  owAxeLeft.Color  := ColBlack;
  owAxeBot.Color  := ColBlack;
  owAxeBot.InfoAxe.NumFormat := '';
  owAxeLeft.InfoAxe.NumFormat := '';
  owAxeBot.InfoAxe.MajorUnit := '';
  owAxeLeft.InfoAxe.MajorUnit := '';

end;

function TOWCChart.GetTypeGraph : string;
begin
  Result := 'chChartTypePie';
  case FTypeGraph of
    tyCamembert : Result := 'chChartTypePie';
    tyHistogramme : Result := 'chChartTypeColumnClustered';
    tyBar : Result := 'chChartTypeBarClustered';
    tyLigne : Result := 'chChartTypeLine';
  end;
end;

function TOWCChart.GetCouleurGraph(Co : TCouleurGraph) : string;
begin
//colYellow, colGreen, colBlue, colSienna, colPurple, colRed, colWhite, colBlack,
  //                 colOrange, colDarkSalmon, ColGray
  Result := '';
  case Co of
    colYellow     : Result := 'Yellow';
    colGreen      : Result := 'Green';
    colBlue       : Result := 'Blue';
    colSienna     : Result := 'Sienna';
    colPurple     : Result := 'Purple';
    colRed        : Result := 'Red';
    ColWhite      : Result := 'White';
    ColBlack      : Result := 'Black';
    ColOrange     : Result := 'Orange';
    colDarkSalmon : Result := 'DarkSalmon';
    ColSalmon     : Result := 'Salmon';
  end;
end;

function TOWCChart.GetPosition(Po : TPositionText) : string;
begin
  Result := '0';
  {Pour les légendes et les titres des graphs}
  case Po of
    poPositionAutomatic : Result := '0';//chPositionAutomatic =
    poPositionBottom    : Result := '2';//chPositionBottom =
    poPositionLeft      : Result := '3';//chPositionLeft =
    poPositionRight     : Result := '4';//chPositionRight =
    poPositionTop       : Result := '1';//chPositionTop =
  end;
end;

end.

