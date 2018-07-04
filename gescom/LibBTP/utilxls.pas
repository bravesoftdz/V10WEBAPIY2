unit UtilXls;

interface

uses Windows,
     classes,
     Graphics,
     math;

const
  xlBorderStyleNone=-4142;
  xlCellTypeLastCell = $0000000B;
  xlLastCell = $0000000B;
  xlLeft = $FFFFEFDD;
  xlRight = $FFFFEFC8;
  xlCenter = $FFFFEFF4;
//
  xlDatabase = $00000001;
//axes
  xlColumnField = $00000002;
  xlDataField = $00000004;
  xlRowField = $00000001;
//function
  xlAverage = $FFFFEFF6;
  xlCount = $FFFFEFF0;
  xlCountNums = $FFFFEFEF;
  xlMax = $FFFFEFD8;
  xlMin = $FFFFEFD5;
  xlProduct = $FFFFEFCB;
  xlStDev = $FFFFEFC5;
  xlStDevP = $FFFFEFC4;
  xlSum = $FFFFEFC3;
  xlVar = $FFFFEFBC;
  xlVarP = $FFFFEFBB;

  xlEdgeBottom = $00000009;
  xlEdgeLeft = $00000007;
  xlEdgeRight = $0000000A;
  xlEdgeTop = $00000008;

  xlContinuous = $00000001;
  xlDash = $FFFFEFED;
  xlDashDot = $00000004;
   xlDashDotDot = $00000005;
  xlDot = $FFFFEFEA;
  xlDouble = $FFFFEFE9;
  xlSlantDashDot = $0000000D;
  xlLineStyleNone = $FFFFEFD2;

  xlUnderlineStyleNone = $FFFFEFD2;
  xlUnderlineStyleSingle = $00000002;
  xlUnderlineStyleSingleAccounting = $00000004;

  xlTextWindows = $00000014;
  xlNoChange = $00000001;

  xlSolid = $00000001;
  xlGrid = $0000000F;
  msoShapeChevron = $00000034;
  msoShapeRectangle = $00000001;
  msoShapeParallelogram = $00000002;
  msoShapeRoundedRectangle = $00000005;
  msoShapeHexagon = $0000000A;

  xlDate = $00000002;
  xlNumber = $FFFFEFCF;
  xlText = $FFFFEFC2;

type
  TXlsProc = procedure(ObjVariant : OleVariant) of object;
  TTFormatColonneExcel=(fceTexte) ;
  TTFormatDessin=(fdFleche,fdBerceau,fdBisot,fdLosange,fdEtoile,fdRoundRect,fdRectangle) ;
function GetOfficeDir : string;
procedure AddVbModuleFromFile (RepertSource,nom : string);
function lanceAppliWindows (var si: TStartupInfo;var pi:  TProcessInformation;st: string) : boolean;
function OpenExcel( NewInst: Boolean; var WinExcel: OleVariant; var WinNew: Boolean ): Boolean;
function OpenWorkBook(FileName : string ; var WinExcel : OleVariant ; deleteexist : boolean = false) : OleVariant;
function GetFicImport : string;
function GetFicExport : string;
function GetRefColonne( ACol : integer) : string;
function ExportToExcelViaOle(FileName : string;LData : TStrings) : boolean;
procedure GetListeOfSheets(FileName : string ; var Sheets : TStringList; var WinExcel : OleVariant );
function ExportListeToExcel (WinExcel : OleVariant ; FileName : string ; LData : TStrings ; var Ferme : boolean) : boolean;
function OfficeExcelDispo : boolean;
function IsExcelLaunched  : boolean;
function GetExcelPath : string;
function GetIndiceColonne( aCol : string) : integer;

function AddSheet(WE,WB : OleVariant ; SheetName : string) : olevariant;
function SelectSheet (WB : OleVariant ; SheetName : string) : olevariant;
procedure DelSheetReserved (WE,WB : Olevariant);
function AddSheetReserved(WE,WB : OleVariant) : olevariant;
procedure DeleteSheets (WB : Olevariant); // suppression des feuilles
procedure ClearSheetWorkbook(WE, WB : OleVariant);
function TrouveSheet(WB : OleVariant ; SheetName : string) : boolean;
procedure DelSheet (WE,WB : Olevariant ; name : string);
function ValideSheetName(st : string) : boolean;
function GetExcelVersion: Double;
function GetCelluleName(Colonne,Ligne:Integer):String;

procedure ExcelColorRange(CurrentSheet:OleVariant ; LigneFrom, ColonneFrom, LigneTo, ColonneTo: Integer; Value: TColor ; AvecEncadrement:Boolean=False) ;
procedure ExcelFontRange(CurrentSheet:OleVariant ; LigneFrom, ColonneFrom, LigneTo, ColonneTo: Integer; F: TFont) ;
procedure ExcelMergeRange(CurrentSheet:OleVariant ; LigneFrom, ColonneFrom, LigneTo, ColonneTo: Integer) ;
procedure ExcelColorCell(CurrentSheet:OleVariant ; Ligne,Colonne:Integer;Value:TColor;AvecEncadrement:Boolean=False) ;
procedure ExcelRangeWidth(CurrentSheet:OleVariant ; ColonneFrom,LigneFrom,ColonneTo,LigneTo,Width: Integer) ;
procedure ExcelColWidth(CurrentSheet:OleVariant ; Index, Width: Integer) ;
procedure ExcelAlignCell(CurrentSheet:OleVariant ; Ligne,Colonne: Integer; CodeAligne: TAlignment) ;
procedure ExcelAlignCol(CurrentSheet:OleVariant ; Colonne: Integer; CodeAligne: TAlignment) ;
procedure ExcelAlignRange(CurrentSheet:OleVariant ; ColonneFrom,LigneFrom,ColonneTo,LigneTo: Integer; CodeAligne: TAlignment) ;
procedure ExcelText(CurrentSheet:OleVariant ; Ligne, Colonne: Integer; Value: String) ;
function  ExcelCellRect(CurrentSheet:OleVariant ; Ligne, Colonne: Integer): TRect;
procedure ExcelFormatColonne(CurrentSheet:OleVariant ; Colonne: Integer; CodeFormat: TTFormatColonneExcel) ;
procedure ExcelFormeDessin(CurrentSheet:OleVariant ; F: TTFormatDessin; Gauche, Haut, Largeur, Hauteur: Integer; Couleur: TColor ; Text:String='' ; Aligne:TAlignment=taCenter) ;
procedure ExcelClose(var Handle: OleVariant);
procedure ExcelSave(CurrentWorkbook:OleVariant;FileName:String) ;
FUNCTION ExcelDupSheet(We,Wb:OleVariant;SheetName:STRING;NewName:STRING=''):BOOLEAN ;
function GetExcelText (CurrentSheet:OleVariant ; Ligne, Colonne: Integer):string ;
procedure ExcelValue(CurrentSheet:OleVariant ; Ligne, Colonne: Integer; Value: variant) ;
var
FWinExcel : OleVariant;

implementation

uses sysutils, dialogs, forms, comobj, ActiveX, hctrls, hent1, controls,hmsgbox;

function GetExcelVersion: Double;
var  S: String;
begin
  S := GetFromRegistry(HKEY_CLASSES_ROOT,'Excel.Application\CurVer','','',TRUE);
  result := Valeur(Copy(S,Length('Excel.Application.')+1,Length(S)-Length('Excel.Application.')));
end;

function IsExcelLaunched  : boolean;
begin
result := (FindWindow('XLMAIN',nil)<>0) or (FindWindow('EXCEL.EXE',nil)<>0);
end;

function OfficeExcelDispo : boolean;
var V : string;
begin
  V:='' ;
  V:=GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe','Path',V,TRUE) ;
  result := (V<>'');
end;

function GetExcelPath : string;
var V : string;
begin
  V:='' ;
  V:=GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe','Path',V,TRUE) ;
  result := V;
end;

function GetOfficeDir : string;
var V,CurrentVersion : string;
begin
  if GetExcelVersion = 0 then BEGIN result := ''; Exit; END;
  CurrentVersion := floattostr (GetExcelVersion)+'.0';
  V:='' ;
  V:=GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\Microsoft\office\'+CurrentVersion+'\Common\InstallRoot','Path',V,TRUE) ;
  result := V;
end;

function GetFicImport : string;
var
  OD : TOpenDialog;
begin
  OD := TOpenDialog.Create(Application);
  OD.DefaultExt := '*.xls';
  OD.Title := 'Importation de fichiers';
  OD.Filter := 'Fichiers excel (*.xls)|*.xls|Fichiers texte (*.txt)|*.txt|Fichiers csv (*.csv)|*.csv';
  if OD.execute then result := OD.FileName else result := '';
  OD.Free;
end;

function GetFicExport : string;
var
  SD : TSaveDialog;
begin
  SD := TSaveDialog.Create(Application);
  SD.DefaultExt := '*.xls';
  SD.Title := 'Exportation de fichiers';
  SD.Filter := 'Fichiers excel (*.xls)|*.xls|Fichiers texte (*.txt)|*.txt|Fichiers csv (*.csv)|*.csv';
  if SD.execute then result := SD.FileName else result := '';
  SD.Free;
end;

function GetRefColonne( aCol : integer) : string;
const letters : array[0..25] of char = ('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
var i,j : integer;
begin
  if (aCol <= 26) then result := letters[aCol-1]
  else if (aCol > 0) then
    begin
    i := (aCol div 26);
    j := (aCol mod 26); if j=0 then begin j:=26; Dec(i); end;
    result := letters[i-1] + letters[j-1];
    end;
end;


function GetIndiceColonne( aCol : string) : integer;

         function IndiceCaracter (caracter : char): integer;
         const letters : array[0..25] of char = ('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
         var Indice : integer;
         begin
         result := -1;
         for Indice := 0 to 25 do
             begin
             if (letters[indice]=caracter) then
                begin
                result := Indice +1;
                break;
                end;
             end;
         end;

var Indice : integer;
    puissance : integer; { Dans Excel les colonnes sont en base 26 avec valeur de depart a 1}
begin
  puissance := 0;
  result := 0;
  for Indice := length(Acol) downto 1 do
      begin
      result :=  result + (trunc(Intpower (26,Puissance)) * IndiceCaracter (Acol[Indice]));
      inc(puissance);
      end;
end;

function lanceAppliWindows (var si: TStartupInfo;var pi:  TProcessInformation;st: string) : boolean;
begin
ZeroMemory(@si, SizeOf(si) );
si.cb:=SizeOf(si);    // Start the child process.
if not CreateProcess(nil, // No module name (use command line).
    PChar(st),       // Command line.
    nil,             // Process handle not inheritable.
    nil,             // Thread handle not inheritable.
    FALSE,           // Set handle inheritance to FALSE.
    0,               // No creation flags.
    nil,             // Use parent's environment block.
    nil,             // Use parent's starting directory.
    si,              // Pointer to STARTUPINFO structure.
    pi)              // Pointer to PROCESS_INFORMATION structure.
then result:=false
else result := true;
end;

function OpenExcel( NewInst: Boolean; var WinExcel: OleVariant; var WinNew: Boolean ): Boolean;
var
    Dispatch: IDispatch;
    function GetActiveOleObjectEx(const ClassName: String): Boolean;
    var
      ClassID: TCLSID;
      Unknown: IUnknown;
      AResult: HResult;
    begin
      Result := False;
      ClassID := ProgIDToClassID(ClassName);
      AResult := GetActiveObject(ClassID, nil, Unknown);
      if (Unknown = nil) or not Succeeded(AResult) then Exit;
      AResult := Unknown.QueryInterface(IDispatch, Dispatch);
      if not Succeeded(AResult) then Exit;
      Result := True;
    end;
begin
  Result := False;
  WinExcel := UnAssigned;
  WinNew := False;
  if not NewInst then
    if GetActiveOleObjectEx( 'Excel.Application' ) then
    begin
      WinExcel := Dispatch;
      Result := True;
      Exit;
    end;
  try
    WinExcel := CreateOleObject( 'Excel.Application' );
    Result := True;
  except
    WinExcel := UnAssigned;
  end;
  if VarType(WinExcel) = varEmpty then Exit;
  WinNew := True;
end;

function ExportToExcelViaOle(FileName : string;LData : TStrings) : boolean;
var
  WinNew,FermeExcel : boolean;
begin
result := false;
if FileName='' then FileName := GetFicExport;
if FileName='' then exit;
if not OpenExcel(true{false}, FWinExcel, WinNew) then exit;
ExportListeToExcel(FWinExcel,FileName,LData,FermeExcel);
if not FermeExcel then exit;
if not VarIsEmpty(FWinExcel) then FWinExcel.Quit;
FWinExcel := unassigned;
end;

function AddSheet(WE,WB : OleVariant ; SheetName : string) : olevariant;
begin
// suppression feuille si existe deja ....
DelSheet (WE,WB,SheetName);
//
WB.Worksheets.add;
WE.ActiveSheet.Name := SheetName;
result := WE.ActiveSheet;
end;

function TrouveSheet(WB : OleVariant ; SheetName : string) : boolean;
var i : integer;
    bTrouve : boolean;
begin
  bTrouve := false;
  for i:=1 to WB.Sheets.Count do
    begin
    bTrouve := (WB.Sheets[i].Name=SheetName);
    if bTrouve then break;
    end;
  result := bTrouve;
end;

function SelectSheet (WB : OleVariant ; SheetName : string) : olevariant;
begin
result := UnAssigned;
if TrouveSheet(WB,SheetName) then
  begin
  WB.Sheets[SheetName].Select;
  result := WB.ActiveSheet;
  end else result := WB.ActiveSheet;
end;

{
function initEnv : olevariant;
begin
result := WE.ActiveWorkBook;
end;
}

procedure DelSheetReserved (WE,WB : Olevariant);
var Sheet : OleVariant;
begin
  Sheet := SelectSheet (WB,'&#@');
  if VarType(Sheet)<>varEmpty then Sheet.Delete;
end;

procedure DelSheet (WE,WB : Olevariant ; name : string);
var Sheet : OleVariant;
begin
  Sheet := SelectSheet (WB,name);
  if VarType(Sheet)<>varEmpty then Sheet.Delete;
end;

function AddSheetReserved(WE,WB : OleVariant) : olevariant;
begin
WB.Worksheets.add;
WE.ActiveSheet.Name := '&#@';
result := WE.ActiveSheet;
end;

procedure DeleteSheets (WB : Olevariant); // suppression des feuilles
var i,iCount : integer;
    Sheet : OleVariant;
begin
  i := 1;
  iCount := WB.Sheets.Count;
  while i <= iCount do
    begin
    if WB.Sheets[i].Name='&#@' then begin Inc(i); continue; end; // on garde celle là : au - 1 dans un classeur
    Sheet := SelectSheet (WB,WB.Sheets[i].Name);
    if VarType(Sheet)<>varEmpty then
      begin
      Sheet.Delete;
      Dec(iCount);
      end;
    end
end;

procedure ClearSheetWorkbook(WE, WB : OleVariant);
begin
//  AddSheetReserved(WE,WB); // permet d'avoir au moins une feuille dans le classeur
  DeleteSheets(WB);
end;

function ExportListeToExcel (WinExcel : OleVariant ; FileName : string ; LData : TStrings ; var Ferme : boolean) : boolean;
const
  stDataSheetName = 'Stats PGI - Données';
  stCubeSheetName = 'Stats PGI - Tableau';
  stPivotName = 'Cube PGI';
var
  SaveCursor:TCursor;
  WorkBook,Sheet,Cell: Variant;
  stAxe,stData,stTip,stCaption,stMask,stWidth,stOpen,stSummary,stAlign,stOperator : string;
  LMask,LFormatsPGI,LFormatsXLS,LCube : TStrings;

  procedure SetAlignment(ARow,ACol : integer);
  begin
    Cell:=Sheet.Cells[ARow,ACol];
    Cell.Style.IncludeAlignment := true;
    if TAlignment(StrToInt(stAlign)) = taLeftJustify then Cell.HorizontalAlignment := xlLeft else
    if TAlignment(StrToInt(stAlign)) = taCenter then Cell.HorizontalAlignment := xlCenter else
    if TAlignment(StrToInt(stAlign)) = taRightJustify then Cell.HorizontalAlignment := xlRight;
  end;
  procedure SetMask(ARow,ACol : integer); //seulement pour date... voir pour reste si besoin....
  var imask : integer;
  begin
    imask := LFormatsPGI.IndexOF(stMask);
    if imask<>-1 then Sheet.columns[GetRefColonne(ACol)].NumberFormat := LFormatsXLS[iMask]
                 else ;
  end;

  procedure SetFormule(stData : variant ; ARow,ACol : integer);
  begin
    Cell:=Sheet.Cells[ARow,ACol];
    Cell.Formula := '='+stData;  //data
  end;

  procedure SetData ( Tip, V : variant ; ARow,ACol : integer);
	type TDataType  = (dtString, dtBoolean, dtInteger, dtDouble, dtDateTime, dtBitmap); // cf UTobCube
  begin
    Cell:=Sheet.Cells[ARow,ACol];
    if Vartype(v) in [varnull,varempty] then else
     case TDataType(ValeurI(Tip)) of
      dtString:   Cell.Value := V;
      dtBoolean:  Cell.Value := V;
      dtInteger:  Cell.Formula := ValeurI(V);
      dtDouble:   Cell.Formula := Valeur(V);
      dtDateTime: Cell.Formula := V;//VarAsType(V,varDate)
      end;
  end;
  procedure DoExportData;
  var ACol, ARow : integer;
      st : string;
  begin
  Sheet := AddSheet(WinExcel,Workbook,stDataSheetName);
  for ACol:=1 to LData.Count  do
    begin
    ARow := 1;
    st := LData[ACol-1];
    // st est de la forme : tip|caption|mask|width|Ouvert|summary|Align|oper|tip
    stAxe := ReadTokenPipe(st,'|');  //axe
    stCaption := ReadTokenPipe(st,'|');
    stMask := ReadTokenPipe(st,'|');  //mask
    LMask.Add(stMask);
    stWidth := ReadTokenPipe(st,'|'); if not isnumeric(stWidth) then stWidth:='70'; //width
    Sheet.Columns[Format('%s:%s',[GetRefColonne(ACol),GetRefColonne(ACol)])].ColumnWidth := StrToInt(stWidth) div 10;
    stOpen := ReadTokenPipe(st,'|');  //Ouvert
    stSummary := ReadTokenPipe(st,'|');  //summary
    stAlign := ReadTokenPipe(st,'|');  //Align
    stOperator := ReadTokenPipe(st,'|');  //Operation
    stTip := ReadTokenPipe(st,'|');  //tip
    LCube.Add(stAxe+'|'+stCaption+'|'+stOperator); // utiliser lors de la création du cube
    SetData (0,stCaption,ARow,ACol);
    SetMask (ARow,ACol);
    while st<>'' do
      begin
      Inc(ARow); // les data
      stData := ReadTokenPipe(st,'|');
      SetData(stTip,stData,ARow,ACol);
      SetMask(ARow,ACol);
      //SetAlignment(ARow,ACol);
      end;
    end;
  end;
  function GetFunctionOperator(func : string) : cardinal;
  begin
  result := xlSum;// default;
  exit; // 27/11/2001 : TOUJOURS UNE SOMME ON T'A DIT !
  if func='Aucun' then
  else if func='Comptage' then result := xlCount
  else if func='Somme' then result := xlSum
  else if func='Moyenne' then result := xlAverage
  else if func='Minimum' then result := xlMin
  else if func='Maximum' then result := xlMax;
  end;

  procedure CreatePivotTableNT;
  var iMask,ACol,iPos,iH,iV,iD : integer;
      st,axe,txt,func : string;
      RangeDst,ZoneSrc,ZoneDst : string;
      orient : cardinal;
      DataSheet,PV,PT,PF : Variant;
  begin
  RangeDst := 'C1'+':C'+IntToStr(LCube.Count);
  ZoneSrc := stDataSheetName+'!'+RangeDst;
  ZoneDst := stCubeSheetName+'!'+RangeDst;

  DataSheet := SelectSheet(Workbook,stDataSheetName);
  Sheet := AddSheet(WinExcel,Workbook,stCubeSheetName);

  PV := WorkBook.PivotCaches.Add(xlDataBase,ZoneSrc);
  Sheet.Move(EmptyParam,WorkBook.Sheets[stDataSheetName]);
  PT := PV.CreatePivotTable(Sheet.Range['A1'],stPivotName);

  iPos := 0; iH := 0; iV := 0; iD := 0;
  for ACol:=1 to LCube.Count  do
    begin
    st := LCube[ACol-1];
    axe := ReadTokenPipe(st,'|');
    txt := ReadTokenPipe(st,'|');
    func := ReadTokenPipe(st,'|');
    if axe='H' then begin inc(iH); iPos:=iH; orient := xlColumnField; end // champs horizontaux
    else if axe='V' then begin inc(iV); iPos:=iV; orient := xlRowField; end // champs verticaux
    else if axe='D' then begin inc(iD); iPos:=iD; orient := xlDataField; end // champs data
    else continue; // gros probleme !
    PF := PT.PivotFields(txt);
    PF.Orientation := orient;
    if orient=xlDataField then
      begin
      PF.function := GetFunctionOperator(func);
      iMask := LFormatsPGI.IndexOf(LMask[ACol]);
      if iMask>-1 then PF.NumberFormat := LFormatsXLS[iMask];
      end;
    PF.Position := iPos;
    end;

  end;

  procedure CreatePivotTable97;
  var iMask,nbAddFields,ACol,iPos,iH,iV,iD : integer;
      st,axe,txt,func,RangeDst,ZoneSrc,ZoneDst : string;
      orient : cardinal;
      fields : variant;
      DataSheet,PT,PF : Variant;
  begin
  RangeDst := 'C1'+':C'+IntToStr(LCube.Count);
  ZoneSrc := stDataSheetName+'!'+RangeDst;
  ZoneDst := stCubeSheetName+'!'+RangeDst;

  DataSheet := SelectSheet(Workbook,stDataSheetName);
  Sheet := AddSheet(WinExcel,Workbook,stCubeSheetName);

  Sheet.PivotTableWizard(xlDatabase, ZoneSrc, EmptyParam, stPivotName);
  PT := Sheet.PivotTables(stPivotName);

  nbAddFields := 0;
  for ACol:=1 to LCube.Count  do
    begin
    st := LCube[ACol-1];
    axe := ReadTokenPipe(st,'|');
    if axe= 'D' then continue;
    Inc(nbAddFields);
    end;
  fields:=VarArrayCreate([0,nbAddFields-1], varVariant);
  for ACol:=1 to LCube.Count  do
    begin
    st := LCube[ACol-1];
    axe := ReadTokenPipe(st,'|');
    if axe= 'D' then continue;
    txt := ReadTokenPipe(st,'|');
    fields[ACol-1]:=txt;
    end;
  PT.AddFields(fields);

  iPos := 0; iH := 0; iV := 0; iD := 0;
  for ACol:=1 to LCube.Count  do
    begin
    st := LCube[ACol-1];
    axe := ReadTokenPipe(st,'|');
    txt := ReadTokenPipe(st,'|');
    func := ReadTokenPipe(st,'|');
    if axe='H' then begin inc(iH); iPos:=iH; orient := xlColumnField; end // champs horizontaux
    else if axe='V' then begin inc(iV); iPos:=iV; orient := xlRowField; end // champs verticaux
    else if axe='D' then begin inc(iD); iPos:=iD; orient := xlDataField; end // champs data
    else continue; // gros probleme !
    PF := PT.PivotFields(txt);
    PF.Orientation := orient;
    if orient=xlDataField then
      begin
      PF.function := GetFunctionOperator(func);
      iMask := LFormatsPGI.IndexOf(LMask[ACol]);
      if iMask>-1 then PF.NumberFormat := LFormatsXLS[iMask];
      end;
    PF.Position := iPos;
    end;
  end;

  procedure InitListFormats;
  begin
  LFormatsPGI := TStringList.Create; // liste contenant les informations utiles pour le cube
  LFormatsPGI.Add('0');
  LFormatsPGI.Add('0.0');
  LFormatsPGI.Add('0.00');
  LFormatsPGI.Add('#,##0.00');
  LFormatsPGI.Add('#,##0.00;; ;');
  LFormatsPGI.Add('#,##0');
  LFormatsPGI.Add('#,##0;; ;');
  LFormatsPGI.Add('dd/mm/yy');
  LFormatsPGI.Add('dd/mm/yyyy');
  LFormatsPGI.Add('dd mmmm yyyy');
  LFormatsPGI.Add('$');
  LFormatsPGI.Add('$$');
  LFormatsPGI.Add('$#');

  LFormatsXLS := TStringList.Create; // liste contenant les informations utiles pour le cube
  LFormatsXLS.Add('0');
  LFormatsXLS.Add('0'+V_PGI.SepDecimal+'0');
  LFormatsXLS.Add('0'+V_PGI.SepDecimal+'00');
  LFormatsXLS.Add('#'+V_PGI.SepMillier+'##0'+V_PGI.SepDecimal+'00');
  LFormatsXLS.Add('#'+V_PGI.SepMillier+'##0'+V_PGI.SepDecimal+'00;; ;');
  LFormatsXLS.Add('#'+V_PGI.SepMillier+'##0');
  LFormatsXLS.Add('#'+V_PGI.SepMillier+'##0;; ;');
  LFormatsXLS.Add('jj/mm/aa');
  LFormatsXLS.Add('jj/mm/aaaa');
  LFormatsXLS.Add('j mmmm aaaa');
  LFormatsXLS.Add('');
  LFormatsXLS.Add('');
  LFormatsXLS.Add('');
  end;

begin
  Ferme := false;
  SaveCursor := Screen.Cursor;
  Screen.Cursor := crHourGlass;

  WorkBook:=OpenWorkBook(FileName,WinExcel);

  AddSheetReserved(WinExcel,Workbook); // permet d'avoir au moins une feuille dans le classeur
  ClearSheetWorkbook(WinExcel,Workbook);

  LCube := TStringList.Create; // liste contenant les informations utiles pour le cube

  LMask := TStringList.Create; // liste contenant les informations utiles pour le cube
  LMask.Add('Item inutile mais nécessaire pour cohérence Indice / num colonne....');

  InitListFormats;

  DoExportData;
//PB Version Office 2K/XP vs 97
  if GetExcelVersion >= 9 then CreatePivotTableNT
                          else CreatePivotTable97;

  LMask.free; LCube.Free; LFormatsPGI.Free; LFormatsXLS.Free;

  DelSheetReserved(WinExcel,Workbook);
  WorkBook.SaveAs(FileName);

  Screen.Cursor := SaveCursor;
  result := true;

  if PGIAsk('Souhaitez-vous ouvrir Excel ?','Information')=mrYes then
    begin
    Ferme := false;
    SelectSheet(Workbook,stCubeSheetName);
    FWinExcel.Visible := true;
    exit;
    end;

  WorkBook.Close;
end;

function OpenWorkBook(FileName : string ; var WinExcel : OleVariant ; deleteexist : boolean = false) : OleVariant;
begin
  WinExcel.Visible:=false;
  WinExcel.DisplayAlerts:=false;
{
  if FileExists(FileName) then DeleteFile(FileName);
  WinExcel.WorkBooks.Add;
}
  if (FileExists(FileName)) and (deleteexist) then DeleteFile(FileName)
  else if (FileExists(FileName)) then WinExcel.WorkBooks.Open(FileName)
  else WinExcel.WorkBooks.Add;
  result := WinExcel.ActiveWorkBook;
end;

procedure GetListeOfSheets(FileName : string ; var Sheets : TStringList; var WinExcel : OleVariant );
var
  i : integer;
  AW : variant;
  bNew : boolean;
begin
  WinExcel.Visible:=false;
  WinExcel.DisplayAlerts:=false;
  bNew := not FileExists(FileName);
  if bNew then WinExcel.WorkBooks.Add
          else WinExcel.WorkBooks.Open(FileName);
  AW:=WinExcel.ActiveWorkBook;
  AddSheetReserved(WinExcel,AW); // permet d'avoir au moins une feuille dans le classeur
  if bNew then ClearSheetWorkbook(WinExcel,AW)
  else
    if VarIsEmpty(AW)=false then
      for i:=1 to AW.Sheets.Count do
        if AW.Sheets[i].Name='&#@' then else Sheets.Add(AW.Sheets[i].Name);
  WinExcel.Visible:=false;
end;

function ValideSheetName(st : string) : boolean;
var
  i : integer;
  s : string;
begin
  result := false;
  s := UPPERCASE(st);
  for i:=1 to length(s) do
    if s[i] in ['0'..'9','A'..'Z', '.', '_'] then
    else
      begin
      PGIInfo(st+' est un nom de feuille excel invalide !'#13#10'(caractères autorisés 0 à 9, A à Z et _)','ATTENTION');
      exit;
      end;
  result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Calcul le Nom de la colonne/ligne.
Mots clefs ... :
*****************************************************************}
function GetCelluleName(Colonne,Ligne:Integer):String;
var
  I,J: Integer ;
  S: String ;
begin
  if Ligne=-1 then S:='' else S:=IntToStr(Ligne);
  Result:='';
  if Colonne<27 then Result:=Chr(Ord('A')+Colonne-1)+S
  else
    begin
    J:=Colonne DIV 26 ;
    I:=Colonne MOD 26 ;
    Result:=Chr(Ord('A')+J-1)+Chr(Ord('A')+I-1)+S;
    end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 25/04/2002
Modifié le ... :   /  /
Description .. : Cette procédure permet de définir une couleur (avec encadrement ou pas)
Suite ........ : pour un ensemble de cellule.
Mots clefs ... :
*****************************************************************}
procedure ExcelColorRange(CurrentSheet:OleVariant ; LigneFrom, ColonneFrom, LigneTo, ColonneTo: Integer; Value: TColor ; AvecEncadrement:Boolean=False) ;
var
  LesCellules: String ;
  VCellule: OleVariant ;
begin
  if Not VarIsEmpty(CurrentSheet) then
    begin
    LesCellules:=GetCelluleName(ColonneFrom+1,LigneFrom+1)+':'+GetCelluleName(ColonneTo+1,LigneTo+1);
    VCellule:=CurrentSheet.Range[LesCellules];
    if Not VarIsEmpty(VCellule) then
      begin
      VCellule.Interior.Color := Value ;
      VCellule.Interior.Pattern := xlSolid ;
      VCellule.Interior.PatternColor := Value ;
      if AvecEncadrement then VCellule.Borders.LineStyle := xlContinuous;
      end ;
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 25/04/2002
Modifié le ... :   /  /
Description .. : Cette procédure permet de définir une fonte pour un
Suite ........ : ensemble de cellule.
Mots clefs ... :
*****************************************************************}
procedure ExcelFontRange(CurrentSheet:OleVariant ; LigneFrom, ColonneFrom, LigneTo, ColonneTo: Integer; F: TFont) ;
var
  LesCellules: String ;
  VCellule: OleVariant ;
begin
  if Not VarIsEmpty(CurrentSheet) then
    begin
    LesCellules:=GetCelluleName(ColonneFrom+1,LigneFrom+1)+':'+GetCelluleName(ColonneTo+1,LigneTo+1);
    VCellule:=CurrentSheet.Range[LesCellules];
    if Not VarIsEmpty(VCellule) then
      begin
      VCellule.Font.Name:=F.Name ;
      VCellule.Font.Size:=F.Size ;
      VCellule.Font.Color:=F.Color ;
      VCellule.Font.Bold:=fsBold in F.Style;
      VCellule.Font.Italic:=fsItalic in F.Style;
      VCellule.Font.Strikethrough:=fsStrikeOut in F.Style;
      if fsUnderline in F.Style then VCellule.Font.UnderLine := xlUnderlineStyleSingle
                                else VCellule.Font.UnderLine := xlUnderlineStyleNone;
      end ;
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 25/04/2002
Modifié le ... :   /  /    
Description .. : Cette procédure permet de fusionner un ensemble de 
Suite ........ : cellule.
Mots clefs ... :
*****************************************************************}
procedure ExcelMergeRange(CurrentSheet:OleVariant ; LigneFrom, ColonneFrom, LigneTo, ColonneTo: Integer) ;
var
  LesCellules: String ;
  VCellule: OleVariant ;
begin
  if Not VarIsEmpty(CurrentSheet) then
    begin
    LesCellules:=GetCelluleName(ColonneFrom+1,LigneFrom+1)+':'+GetCelluleName(ColonneTo+1,LigneTo+1);
    VCellule:=CurrentSheet.Range[LesCellules];
    if Not VarIsEmpty(VCellule) then VCellule.MergeCells:=True ;
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 25/04/2002
Modifié le ... :   /  /
Description .. : Cette procédure permet de définir une couleur (avec encadrement ou pas)
Suite ........ : pour une cellule.
Mots clefs ... :
*****************************************************************}
procedure ExcelColorCell(CurrentSheet:OleVariant ; Ligne,Colonne:Integer;Value:TColor;AvecEncadrement:Boolean=False) ;
var
  LaCellule: String ;
  VCellule: OleVariant ;
begin
  if Not VarIsEmpty(CurrentSheet) then
    begin
    LaCellule:=GetCelluleName(Colonne+1,Ligne+1);
    VCellule:=CurrentSheet.Range[LaCellule];
    if Not VarIsEmpty(VCellule) then
      begin
      VCellule.Interior.Color := Value ;
      VCellule.Interior.Pattern := xlGrid ;
      VCellule.Interior.PatternColor := Value ;
      if AvecEncadrement then VCellule.Borders.LineStyle := xlContinuous;
      end ;
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 25/04/2002
Modifié le ... :   /  /    
Description .. : Cette procédure permet de définir la largeur d'une colonne.
Mots clefs ... :
*****************************************************************}
procedure ExcelColWidth(CurrentSheet:OleVariant ; Index, Width: Integer) ;
var
  VColl:OleVariant ;
begin
  if Not VarIsEmpty(CurrentSheet) then
    begin
    VColl:=CurrentSheet.Columns ;
    if Not VarIsEmpty(VColl) then VColl.Columns[Index+1].ColumnWidth :=Width ;
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 25/04/2002
Modifié le ... :   /  /
Description .. : Cette procédure permet de définir la largeur d'un ensemble de cellule.
Mots clefs ... :
*****************************************************************}
procedure ExcelRangeWidth(CurrentSheet:OleVariant ; ColonneFrom,LigneFrom,ColonneTo,LigneTo,Width: Integer) ;
var
  VColl:OleVariant ;
begin
  if Not VarIsEmpty(CurrentSheet) then
    begin
    VColl:=CurrentSheet.Range[GetCelluleName(ColonneFrom+1,LigneFrom+1)+':'+GetCelluleName(ColonneTo+1,LigneTo+1)] ;
    if Not VarIsEmpty(VColl) then VColl.ColumnWidth :=Width ;
    end ;
end;

function TranslateAlign(Aligne:TAlignment):LongWord;
begin
  case Aligne of
    taCenter : Result:=xlCenter;
    taLeftJustify : Result:=xlLeft;
    taRightJustify : Result:=xlRight;
    else Result:=xLLeft ;
    end ;
end ;

procedure ExcelAlignRange(CurrentSheet:OleVariant ; ColonneFrom,LigneFrom,ColonneTo,LigneTo: Integer; CodeAligne: TAlignment) ;
var
  VColl: OleVariant ;
begin
  if Not VarIsEmpty(CurrentSheet) then
    begin
    VColl:=CurrentSheet.Range[GetCelluleName(ColonneFrom+1,LigneFrom+1)+':'+GetCelluleName(ColonneTo+1,LigneTo+1)];
    if Not VarIsEmpty(VColl) then
      begin
      VColl.Style.IncludeAlignment := true;
      VColl.HorizontalAlignment := TranslateAlign(CodeAligne) ;
      end ;
    end ;
end ;

procedure ExcelAlignCol(CurrentSheet:OleVariant ; Colonne: Integer; CodeAligne: TAlignment) ;
var
  S: String ;
  VColl: OleVariant ;
begin
  if Not varIsEmpty(CurrentSheet) then
    begin
    S:=GetCelluleName(Colonne+1,-1);
    VColl:=CurrentSheet.Range[S+':'+S];
    if Not VarIsEmpty(VColl) then VColl.HorizontalAlignment := TranslateAlign(CodeAligne) ;
    end ;
end;

procedure ExcelAlignCell(CurrentSheet:OleVariant ; Ligne,Colonne: Integer; CodeAligne: TAlignment) ;
var
  S: String ;
  VColl: OleVariant ;
begin
  if Not varIsEmpty(CurrentSheet) then
    begin
    S:=GetCelluleName(Colonne+1,Ligne+1);
    VColl:=CurrentSheet.Range[S+':'+S] ;
    if Not VarIsEmpty(VColl) then VColl.HorizontalAlignment := TranslateAlign(CodeAligne) ;
    end ;
end;

procedure ExcelText(CurrentSheet:OleVariant ; Ligne, Colonne: Integer; Value: String) ;
begin
  if Not varIsEmpty(CurrentSheet) then CurrentSheet.Cells[Ligne+1,Colonne+1]:=Value ;
end;

procedure ExcelValue(CurrentSheet:OleVariant ; Ligne, Colonne: Integer; Value: variant) ;
begin
  if Not varIsEmpty(CurrentSheet) then CurrentSheet.Cells[Ligne+1,Colonne+1].value :=Value ;
end;

function ExcelCellRect(CurrentSheet:OleVariant ; Ligne, Colonne: Integer): TRect;
var
  VCell: OleVariant ;
begin
  Result:=Rect(0,0,0,0);
  if Not varIsEmpty(CurrentSheet) then
    begin
    VCell:=CurrentSheet.Range[GetCelluleName(Colonne+1,Ligne+1)];
    if Not VarIsEmpty(VCell) then
      Result:=Rect(VCell.Left,VCell.Top,VCell.Left+VCell.Width,VCell.Top+VCell.Height) ;
    end ;
end;

function TranslateDessin(F: TTFormatDessin): LongWord ;
begin
  case F of
    fdLosange   : Result:=msoShapeHexagon ;
    fdFleche    : Result:=msoShapeChevron;
    fdBerceau   : Result:=msoShapeRectangle;
    fdBisot     : Result:=msoShapeParallelogram;
    fdEtoile    : Result:=msoShapeRectangle;
    fdRoundRect : Result:=msoShapeRoundedRectangle;
    fdRectangle : Result:=msoShapeRectangle;
    else Result:=msoShapeChevron;
    end ;
end ;

procedure ExcelFormeDessin(CurrentSheet:OleVariant ; F: TTFormatDessin; Gauche, Haut, Largeur, Hauteur: Integer; Couleur: TColor ; Text:String='' ; Aligne:TAlignment=taCenter) ;
var
  V: OleVariant ;
begin
  if Not varIsEmpty(CurrentSheet) then
    begin
    V:=CurrentSheet.Shapes.AddShape(TranslateDessin(F), Gauche,Haut,Largeur,Hauteur) ;
    V.Fill.ForeColor.Rgb := Couleur;
    V.TextFrame.Characters.Text:=Text ;
    V.TextFrame.HorizontalAlignment:=TranslateAlign(Aligne);
    end ;
end;

function GetFormatCellule(CodeFormat: TTFormatColonneExcel): String ;
begin
  case CodeFormat of
    fceTexte : Result:='@' ;
    end ;
end ;

procedure ExcelFormatColonne(CurrentSheet:OleVariant ; Colonne: Integer; CodeFormat: TTFormatColonneExcel) ;
var
  S: String ;
  VColl: OleVariant ;
begin
  if Not varIsEmpty(CurrentSheet) then
    begin
    S:=GetCelluleName(Colonne+1,-1);
    VColl:=CurrentSheet.Range[S+':'+S];
    if Not VarIsEmpty(VColl) then VColl.NumberFormat:=GetFormatCellule(CodeFormat) ;
    end ;
end;

procedure ExcelClose(var Handle: OleVariant);
begin
  if Not VarIsEmpty(Handle) then
    begin
    { Affiche ou pas 'Voulez-vous sauvegarder ...' }
    Handle.DisplayAlerts:=False ;

    { Comme son nom l'indique, non mais ... }
    Handle.Quit ;

    Handle:=Unassigned;
    end ;
end ;

procedure ExcelSave(CurrentWorkbook:OleVariant;FileName:String) ;
begin
  { Suppression du fichier }
  DeleteFile(FileName) ;

  if Not VarisEmpty(CurrentWorkBook) then
    CurrentWorkBook.SaveAs(FileName) ;
end ;

FUNCTION ExcelDupSheet(We,Wb:OleVariant;SheetName:STRING;NewName:STRING=''):BOOLEAN ;
VAR
  Ws: OleVariant ;
BEGIN
  TRY
    IF Not VarIsEmpty(Wb.Sheets[SheetName]) then
      BEGIN
      Wb.Worksheets.item[SheetName].Copy ;
      Ws:=AddSheet(We,Wb,NewName);
      Ws.Paste ;
      RESULT:=True ;
      END ELSE RESULT:=False;
  EXCEPT
    ON E: Exception DO Result:=False ;
  END ;
END ;

function GetExcelText (CurrentSheet:OleVariant ; Ligne, Colonne: Integer):string ;
begin
if Not varIsEmpty(CurrentSheet) then result := CurrentSheet.Cells[Ligne,Colonne];
end;

function CopieFichier (nomOrig,Nomdest : string):boolean;
var
FromF, ToF: file;
NumRead, NumWritten: Integer;
Buf: array[1..2048] of Char;
begin
AssignFile(FromF, NomOrig);
{$I-}
Reset(FromF, 1);
{$I+}
if IoResult = 0 then
  begin
  AssignFile(ToF, Nomdest);
  {$I-}
  Rewrite(ToF, 1);
  {$I+}
  if IoResult = 0 then
     begin
     result := true;
     repeat
         BlockRead(FromF, Buf, SizeOf(Buf), NumRead);
         BlockWrite(ToF, Buf, NumRead, NumWritten);
     until (NumRead = 0) or (NumWritten <> NumRead);
     end else Result := false;
  CloseFile(ToF);
  end else result := false;
CloseFile(FromF);
end;

procedure AddVbModuleFromFile (RepertSource,nom : string);
var dest : string;
begin
// installe un add-in pour Excel dans le repertoire de demarrage auto des add-ins
if GetOfficeDir = '' then exit;
dest := GetOfficeDir+'XlStart\';
CopieFichier (RepertSource+'\'+nom,dest+nom);
end;

end.
