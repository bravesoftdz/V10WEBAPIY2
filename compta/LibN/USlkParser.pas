{**********************************************
     Parser de fichier au format SYLK
            E. PLIEZ 12/12/2000
***********************************************}

unit USlkParser;
{$IFDEF EAGLPKG} ggg {$ENDIF}

interface

uses windows, classes, forms,
{$IFDEF VER140}
     Variants,
{$ENDIF}
 	dialogs, graphics, sysutils, math, inifiles, HEnt1;

type

    TIDRec = record         //First record in the Sylk File
        strProduceProg : string;
        strStyleCellProtect : string;
        strNEIgnored : string;
    end;

    TBRec = record          //Sheet bound
        iRowsCount : integer;       //;Yn
        iColsCount : integer;       //;Xn
        theRect : TRect;              //;Dy1 x1 y2 x2
    end;

    TPRec = record        //Picture format
//        strType : string;                 // 'P' 'E' ou 'F'
        strFormat : string;               //;P
//        theFont : TFont;                    //;E ;F
    end;

    TCRec = record        //Cell definition
        strExpression : string;     //;E
        strValue : string;          //;K
        varValue : variant;          //;K
        iRowIndex : integer;        //;R
        iCOlIndex : integer;        //;C
        strSharedValue : string;    //;G
        strSharedExpr : string;     //;D
        strSharedInfo : string;     //;S at row ;R ;C records
        strNotProtectCell : string; //;N
        strProtectCell : string;    //;P
        strCellHidden : string;     //;H
        strULExpre : string;        //;M ;R et ;C
        strUlTable : string;        //;T ;R et ;C
        strInsideMatrix : string;   //;I at row ;R ;C record
    end;

    TSingleFormatCell = record
      strType : string;                 // ;F ou ;D
      strFormat : string;               // [DCEFG$*%] Default, Currency (ext), Exponent, Fixed,
                                        //            General, Dollar, Graph, Percent
      iNbDigits : integer;              //   Number of digits
      strAlignment : string;            // [DCGLR-X] Default, Center, General (standard), Left,
                                        // Right, ignored, Fill
      iwidth : integer;                 // for D format only : width
    end;

    TBorderCell = record                     // [IDTLBR] Italic, bolD, gridlines Top, Left, Bottom, Right
        bGridLinesTop : boolean;
        bGridLinesLeft : boolean;
        bGridLinesBottom : boolean;
        bGridLinesRight : boolean;
    end;

    TDisplayRec = record        //Formats
        bNoShowHeader : boolean;        //;H
        bNoShowDefGridlines : boolean;  //;G
        bShowFormula : boolean;         //;E
        bShowCommas : boolean;          //;K
    end;

    TSlkCell = record
{
        iRow : integer;
        iCol : integer;
}
        strName : string;
        strValue : string;
        strComment : string;
        varValue : variant;
        strFormule : string;
        theType : string;
        theMask : TPRec;              //format d'affichage
        theDescription : TCRec;         //
        theFormat : TSingleFormatCell;
        theFont : TFont;
        theBorder : TBorderCell;
    end;

    TSlkColumn = record
        iNum : integer;
        iWidth : integer;
    end;

    TSlkRow = record
        iNum : integer;
        iHeight : integer;
        bAffected : boolean;
    end;

    TNERec = record       //External Link; ignore if ;E present in ID record
        strEpr : string;                //;E
        strFileName : string;           //;F
        strSourceArea : string;         //;S
    end;

    TNNRec = record       //Names
        strError : string;
        iAffectNamePriority : integer;         // sert a determiner si le nom doit etre affecter
        iRow : integer;
        iCol : integer;
        strName : string;               //;N
        strExpr : string;               //;E
        strRunableMacro : string;       //;G
        strOrdinaryMacro : string;      //;K
        bUsableAsFunction : boolean;    //;F
    end;

    TORec = record        //Global options
        bIterationOn : boolean;           //;A
        bCompletionTestCurCell : boolean; //;C
        bSheetProtectNoPW : boolean;      //;P
        bRowColModeReferences : boolean;  //;L
        bManualRecalc : boolean;          //;M
        bPrecisionFormatted : boolean;    //;R
        bMacroSheet : boolean;            //;E
    end;

    TColorWindow = record
        iForeground : integer;
        iBackground : integer;
        iBorders : integer;
    end;

    TSplitWindow = record
        strSplitMode : string;                   //T title, H Horizontal, V Vertical
        iNbScreenRows : integer;
        iCursorPosition : integer;
    end;

    TGrpColumnWidths = record             //;W n1 n2 n3
        iFirstCol : integer;
        iLastCol : integer;
        iWidth : integer;
    end;

{
    TERec = record        //End of File
    end;
}

{* IGNORED
    TNLRec = record;       //Chart external link
    end;
*}

{* IGNORED
    TNURec = record;       //File name substitution
      private
    end;
*}

{* IGNORED
    TCurrentCellCoord = record
        iRow : integer;
        iCol : integer;
    end;
*}

{*
    TWRec = record        //Window record
        strInfos : string;                       //;R n1  ... n8  : Title freeze info
                                                 //   n9  ... n12 : Scroll bar info
                                                 //   n13 ... n14 : Split bar info
        iWindowNumber : integer;                 //;N
        tCurrentCellCoord : TCurrentCellCoord;   //;A
        bBorderedWindow : boolean;               //;B
        tColorWindow : TColorWindow;             //;C
        tSplitWindows : TSplitWindow;            //;S
    end;
*}


TSlkGeneral = class
  private
//theTob : TOB;
//    theLines : TStringList;
//    TheIDInfo : TIDRec;

    TabFormats : array of TPRec;
    TabFontes : array of TFont;
    TabFontesDefault : array of TFont;
    theSheet : array of array of TSlkCell;
    theNames : array of array of TNNRec;
//EPZ 07/03/01    theSlkRows : array of TSlkRow;
    theSlkCols : array of TSlkColumn;
    theDisplayOpt : TDisplayRec;


    theSheetBounds : TBRec;
    theCurrentRow : integer;
    theCurrentCol : integer;
    theActiveCell : TSlkCell;

    theEntireRow : integer;
    theEntireCol : integer;

    theIdFontSize : string;


//    theGlobalOpt : TORec;
//    theCurrentCellCoord : TCurrentCellCoord;
//    theWindowsInfos : TWRec;
    theCurrentFont : TFont;
//  theCurrentMask : TPRec;
    theCurrentName : TNNRec;
    theRectCells : TRect;

//    iMinRowAffected : integer;
//    iMaxRowAffected : integer;

    function BaliseSlk(theBalise: string): string;
    procedure ParseLine(theLine: string);
    procedure ParseID(theLine: string);
    procedure ParseP(theLine: string);
    procedure ParseP_EF(theLine,theBalise : string);
    procedure ParseP_P(theLine: string);
    procedure ParseFont(Ligne: string ; var laFonte : TFont);
    procedure ParseB(theLine: string);
    procedure ParseF(theLine: string);
    procedure ParseC(theLine: string);
    procedure ParseF_W(var theLine: string);
    procedure ParseF_C(var theLine: string);
    procedure ParseF_E(var theLine: string);
    procedure ParseF_G(var theLine: string);
    procedure ParseF_H(var theLine: string);
    procedure ParseF_K(var theLine: string);
    procedure ParseF_M(var theLine: string);
    procedure ParseF_N(var theLine: string);
    procedure ParseF_P(var theLine: string);
//    procedure ParseF_R(var theLine: string);
    procedure ParseF_S(var theLine: string);
    procedure ParseF_FD(var theLine: string ; theBalise : string);
    procedure ParseC_K(var theLine: string);
    procedure ParseC_E(var theLine: string);
    procedure SetValCurrentCell(theValeur: string);
    procedure SetFormuleCurrentCell(theFormule: string);
    procedure Parse_X(var theLine: string);
    procedure Parse_Y(var theLine: string);
    procedure SetColProperties(Quoi : string);
    procedure SetRowProperties(Quoi : string);
    procedure InitActiveCell;
    procedure InitTheCurrentName;
    procedure ParseNN(theLine: string);
    procedure ParseNN_E(var theLine: string);
    procedure ParseNN_N(var theLine: string);
    procedure SetTheNamesProperties;
    procedure AffecteFormatCell(ARow, ACol: integer; theActiveCell: TSlkCell);
    procedure AffecteValCell(ARow, ACol: integer; theActiveCell: TSlkCell);
    procedure ParseE(theLine: string);
    function GetTheTipe(theCell : TSlkCell) : Char;
    function SetTheVariant(theCell: TSlkCell): variant;
    function SetTheValue(theValue: string): string;
    procedure ParseC_A(var theLine: string);
    procedure SetCommentCurrentCell(theComment: string);
    function UnselectedRow(ARow: integer): boolean;
    function UnselectedCol(ACol: integer): boolean;
    procedure TraductCol(var ACol: integer ; Sens : integer = 1);
    procedure TraductRow(var ARow: integer ; Sens : integer = 1);
    procedure GetActiveCell(ARow, ACol : integer);
    procedure GetInfoCible(Ligne: string);
    procedure GetInfoCurrent(Ligne: string; var ARow, ACol: integer);
    procedure GetInfoEntire(Ligne: string; var ARow, ACol: integer);
  public
    constructor Create;
    destructor Destroy; override;
    procedure SetLines(lesLignes : TStrings);
    function GetLines : String;
    function GetName(ARow, ACol : integer) : String;
    function GetFormat(ARow, ACol : integer) : String;
    function GetFonte(ARow, ACol : integer) : String;
//    function ParseFile : boolean;
    function ParseLines(theLines: TStrings): boolean;
    function ParseString(theString: string): boolean;
    procedure GetDim(var iRowCount, iColCount : integer);
    function GetWitdh(laColonne : integer) : integer;
    function GetFormule(ARow, ACol: integer): String;
    function GetMask(ARow, ACol: integer): String;
    function GetValue(ARow, ACol: integer): variant;
    function GetBordure(ARow, ACol: integer): String;
    function GetType(ARow, ACol: integer): string;
    function GetAlign(ARow, ACol: integer): TAlignment;
    function GetComment(ARow, ACol: integer): String;
    procedure GetInfos(ARow, ACol: integer ; var leNom : string ; var laValeur : variant );
end;

var F : TextFile;

implementation

uses {UTob,}SlkOutils;

{ TSlkGeneral }

procedure trace(st : string);
begin
WriteLN(F,st);
end;

function TSlkGeneral.ParseLines(theLines : TStrings) : boolean;
var
  i : integer;
//  inIFic : TIniFile;
//  theDir : string;
begin
  for i:=0 to theLines.Count-1 do
  begin
    if Length(theLines[i])=0 then continue;
    ParseLine(theLines[i]);
  end;
  result := true;
end;

function TSlkGeneral.ParseString (theString : string) : boolean;
var
  PCur : PChar;
  St : string;
  theLines : TStrings;
begin
theLines := TStringList.Create;
theLines.Text := theString;
theLines.SaveToFile('c:\theLines.txt');
AssignFile(F, 'c:\debug.txt');
Rewrite(F);
try
theLines.free;
  St := '';
  PCur := @theString[1];
  while PCur^<>#0 do
    begin
    St := St + PCur^;
    if (St[Length(St)]=#10) and (St[Length(St)-1]=#13) then
      begin
      SetLength(St,Length(St)-2);
      ParseLine(St);
      St := '';
      end;
    Inc(PCur);
    end;
  result := true;
except
CloseFile(F);
exit;
end;
CloseFile(F);
end;

procedure TSlkGeneral.ParseLine(theLine: string);
var
  idxSlk : string;
begin
trace(theLine);
  idxSlk := BaliseSlk(theLine);
  if idxSlk = SLK_ID then ParseID(theLine) else
  if idxSlk = SLK_P then ParseP(theLine) else
  if idxSlk = SLK_B then ParseB(theLine) else
  if idxSlk = SLK_O then {....} else
  if idxSlk = SLK_F then ParseF(theLine) else
  if idxSlk = SLK_C then ParseC(theLine) else
  if idxSlk = SLK_NN then ParseNN(theLine) else
  if idxSlk = SLK_E then ParseE(theLine) else
  ;
end;

procedure TSlkGeneral.ParseID(theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_ID);
{* initialisation
  strProduceProg : string;
  strStyleCellProtect : string;
  bNEIgnored : string;
*}
end;

procedure TSlkGeneral.ParseP(theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_P);
  if Pos(SLK_P_P,Ligne)=1 then ParseP_P(Ligne) else
  if Pos(SLK_P_F,Ligne)=1 then ParseP_EF(Ligne,SLK_P_F)  else
  if Pos(SLK_P_E,Ligne)=1 then ParseP_Ef(Ligne,SLK_P_E);
end;

procedure TSlkGeneral.ParseP_P(theLine: string);
var Balise,Ligne : string;
    idx : integer;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_P_P);
  idx := High(TabFormats)+1;
  SetLength(TabFormats,idx+1);
  TabFormats[idx].strFormat:=Ligne;
end;

procedure TSlkGeneral.ParseP_EF(theLine, theBalise : string);
var Balise,Ligne : string;
    idx : integer;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,theBalise);
  idx := High(TabFontes)+1;
  SetLength(TabFontes,idx+1);
  TabFontes[idx] := TFont.Create;
  ParseFont(Ligne,TabFontes[idx]);
end;

procedure TSlkGeneral.ParseFont(Ligne : string ; var laFonte : TFont);
var
    strBalise,strStyles : string;
    iSize : integer;
begin
  laFonte.Name := ReadTokenPlus(Ligne,';',false);
  if Pos(SLK_P_M,Ligne)=1 then // taille de la font * 2
  begin
    strBalise := SkipBalise(Ligne,SLK_P_M);
    iSize := ReadTokenPlus(Ligne,';',false);
    laFonte.Size := iSize div 20;
  end;
  if Pos(SLK_P_S,Ligne)=1 then  // options : rien, gras (B) italic (I) souligné (U)
  begin
    laFonte.Style := [];
    strBalise := SkipBalise(Ligne,SLK_P_S);
    strStyles := ReadTokenPlus(Ligne,';',false);
    if Pos('B',strStyles)<>0 then laFonte.Style := laFonte.Style + [fsBold];
    if Pos('I',strStyles)<>0 then laFonte.Style := laFonte.Style + [fsItalic];
    if Pos('U',strStyles)<>0 then laFonte.Style := laFonte.Style + [fsUnderline];
  end;
  if Pos(SLK_P_L,Ligne)=1 then  // ?
  begin
  end;
end;

procedure TSlkGeneral.ParseB(theLine: string);
var Balise,Ligne : string;
    i,j,iDimY,iDimX : integer;
    rRect : TRect;
begin
  Ligne := theLine; {iDimY := 0; iDimX := 0;}
  Balise := SkipBalise(Ligne,SLK_B);
  if Pos(SLK_Y,Ligne)=1 then
  begin
    Balise := SkipBalise(Ligne,SLK_Y);
    {iDimY := }ReadTokenPlus(Ligne,';',false);
  end;
  if Pos(SLK_X,Ligne)=1 then
  begin
    Balise := SkipBalise(Ligne,SLK_X);
    {iDimX := }ReadTokenPlus(Ligne,';',false);
  end;
  if Pos(SLK_B_D,Ligne)=1 then
  begin
    Balise := SkipBalise(Ligne,SLK_B_D);
    rRect.Top := ReadTokenPlus(Ligne,' ',true);
    rRect.Left := ReadTokenPlus(Ligne,' ',true);
    rRect.bottom := ReadTokenPlus(Ligne,' ',true);
    rRect.Right := ReadTokenPlus(Ligne,' ',true);
  end;

{
iDimY := rRect.bottom-rRect.Top+1;
iDimX := rRect.Right-rRect.Left+1;
}
iDimY := rRect.bottom+1;
iDimX := rRect.Right+1;

  InitActiveCell;
  SetLength(theSheet,iDimY,iDimX);
  for i:=0 to iDimY-1 do
    for j:=0 to iDimX-1 do
    begin
      theSheet[i][j].theFont := TFont.Create;
      AffecteFormatCell(i,j,theActiveCell); // initialisation par defaut
      AffecteValCell(i,j,theActiveCell);    // initialisation par defaut
    end;
  SetLength(theNames,iDimY,iDimX);
  SetLength(theSlkCols,iDimX);
//EPZ 07/03/01  SetLength(theSlkRows,iDimY);
//EPZ 07/03/01  iMinRowAffected := 0;
//EPZ 07/03/01  iMaxRowAffected := 0;
//EPZ 07/03/01  for i:=0 to iDimY-1 do theSlkRows[i].bAffected := false;

  theSheetBounds.iRowsCount := iDimY;
  theSheetBounds.iColsCount := iDimX;
  theSheetBounds.theRect := rRect;
end;

function TSlkGeneral.BaliseSlk(theBalise : string) : string;
begin
  if Pos(SLK_ID,theBalise)=1 then result := SLK_ID else
  if Pos(SLK_P,theBalise)=1 then result := SLK_P else
  if Pos(SLK_F,theBalise)=1 then result := SLK_F else
  if Pos(SLK_B,theBalise)=1 then result := SLK_B else
  if Pos(SLK_O,theBalise)=1 then result := SLK_O else
  if Pos(SLK_C,theBalise)=1 then result := SLK_C else
  if Pos(SLK_NN,theBalise)=1 then result := SLK_NN else
  if Pos(SLK_E,theBalise)=1 then result := SLK_E else
  ;
end;

constructor TSlkGeneral.Create;
begin
//  theLines := TStringList.Create;
  SetLength(TabFormats,0);
  SetLength(TabFontes,0);
  SetLength(TabFontesDefault,0);

  theCurrentRow := -1;
  theCurrentCol := -1;
  theEntireRow := -1;
  theEntireCol := -1;

  SetLength(theSheet,0,0);
  SetLength(theNames,0,0);
  SetLength(theSlkCols,0); //EPZ 07/03/01
  theCurrentFont := TFont.Create;
  theActiveCell.theFont := TFont.Create;
end;

destructor TSlkGeneral.Destroy;
var i,j : integer;
begin
  for i:=Low(theSheet) to High(theSheet) do
    for j:=Low(theSheet[i]) to High(theSheet[i]) do
      theSheet[i][j].theFont.Free;
  for i:=Low(TabFontes) to High(TabFontes) do TabFontes[i].Free;
  for i:=Low(TabFontesDefault) to High(TabFontesDefault) do TabFontesDefault[i].Free;
  SetLength(theSheet,0,0);
  SetLength(theNames,0,0);
  SetLength(theSlkCols,0); //EPZ 07/03/01
  theActiveCell.theFont.Free;
  SetLength(TabFormats,0);
  SetLength(TabFontes,0);
  SetLength(TabFontesDefault,0);
  theCurrentFont.Free;
end;

function TSlkGeneral.GetLines: String;
begin
//  result := theLines.Text;
end;

procedure TSlkGeneral.SetLines(lesLignes : TStrings);
begin
//  theLines.Assign(lesLignes);
end;

procedure TSlkGeneral.GetDim(var iRowCount, iColCount: integer);
begin
  iRowCount := theSheetBounds.iRowsCount;
  iColCount := theSheetBounds.iColsCount;
end;

procedure TSlkGeneral.ParseF(theLine: string);
var Balise,Ligne : string;
//    ARow,ACol : integer;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F);
// Recup format de la cellule/colonne/ligne eventuellement deja defini
  GetInfoCible(Ligne);
  while Ligne<>'' do
  begin
    if Pos(SLK_F_P,Ligne)=1 then ParseF_P(Ligne) else
    if Pos(SLK_F_S,Ligne)=1 then ParseF_S(Ligne) else
    if Pos(SLK_F_W,Ligne)=1 then ParseF_W(Ligne) else
    if Pos(SLK_F_K,Ligne)=1 then ParseF_K(Ligne) else
    if Pos(SLK_F_F,Ligne)=1 then ParseF_FD(Ligne,SLK_F_F) else
    if Pos(SLK_F_D,Ligne)=1 then ParseF_FD(Ligne,SLK_F_D) else
    if Pos(SLK_F_C,Ligne)=1 then ParseF_C(Ligne) else
//09/03/01 on ignore les formats des lignes : perturbe format des colonnes
// if Pos(SLK_F_R,Ligne)=1 then ParseF_R(Ligne) else
    if Pos(SLK_F_E,Ligne)=1 then ParseF_E(Ligne) else
    if Pos(SLK_F_G,Ligne)=1 then ParseF_G(Ligne) else
    if Pos(SLK_F_H,Ligne)=1 then ParseF_H(Ligne) else
    if Pos(SLK_F_M,Ligne)=1 then ParseF_M(Ligne) else
    if Pos(SLK_F_N,Ligne)=1 then ParseF_N(Ligne) else
    if Pos(SLK_X,Ligne)=1 then Parse_X(Ligne) else
    if Pos(SLK_Y,Ligne)=1 then Parse_Y(Ligne) else
    SkipBalise(Ligne,SLK_UNKOWN);
  end;
  //traite les infos parsées... affectation a l'objet theSheet...
  theActiveCell.theFont.Assign(theCurrentFont);
  if (theEntireRow>-1) or (theEntireCol>-1) then
    begin
    if not UnselectedRow(theEntireRow) then begin SetRowProperties(SLK_F); theEntireRow := -1; end;
    if not UnselectedCol(theEntireCol) then begin SetColProperties(SLK_F); theEntireCol := -1; end;
    end;
end;

procedure TSlkGeneral.ParseF_W(var theLine: string);
var Balise,Ligne : string;
    i,iMax,iMin : integer;
    GrpColumnWidths : TGrpColumnWidths;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_W);
  GrpColumnWidths.iFirstCol := ReadTokenPlus(Ligne,' ',true);
  GrpColumnWidths.iLastCol := ReadTokenPlus(Ligne,' ',true);
  GrpColumnWidths.iWidth := ReadTokenPlus(Ligne,' ',true);
  iMin := Max(theSheetBounds.theRect.Left,GrpColumnWidths.iFirstCol);
  iMax := Min(theSheetBounds.theRect.Right,GrpColumnWidths.iLastCol);
  for i:=iMin to iMax do
  begin
    if UnselectedCol(i-1) then continue;
    theSlkCols[i-1].iWidth := GrpColumnWidths.iWidth;
  end;
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_S(var theLine: string);
var Balise,Ligne,Format : string;
    PCur : PChar;
    iPos,Idx : integer;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_S);
  Format := ReadTokenPlus(Ligne,';',false);
  theLine := Ligne;
  if High(theSheet)=0 then exit;
  iPos := Pos('M',Format);
  if iPos<>0 then   //sélection de la fontes % tableau P;E et P;F
  begin
    PCur := @Format[iPos+1];
    Idx := StrToInt(ReadNumberSt(PCur))-1;
    if Idx<High(TabFontes) then theCurrentFont.Assign(TabFontes[Idx]);
  end;
  if Pos('D',Format)<>0 then theCurrentFont.Style := theCurrentFont.Style + [fsBold];
  if Pos('I',Format)<>0 then theCurrentFont.Style := theCurrentFont.Style + [fsItalic];
  if Pos('U',Format)<>0 then theCurrentFont.Style := theCurrentFont.Style + [fsUnderline];
  theActiveCell.theBorder.bGridLinesTop := Pos('T',Format)<>0;
  theActiveCell.theBorder.bGridLinesLeft := Pos('L',Format)<>0;
  theActiveCell.theBorder.bGridLinesBottom := Pos('B',Format)<>0;
  theActiveCell.theBorder.bGridLinesRight := Pos('R',Format)<>0;
{
  iPos := Pos('M',Format);
  if iPos<>0 then   //sélection de la fontes % tableau P;E et P;F
  begin
    PCur := @Format[iPos+1];
    Idx := StrToInt(ReadNumberSt(PCur));
    if Idx<High(TabFontes) then theCurrentFont.Assign(TabFontes[Idx]);
  end;
}
end;

procedure TSlkGeneral.ParseF_C(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_C);
  theEntireCol := ReadTokenPlus(Ligne,';',true)-1;
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_E(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_E);
  theDisplayOpt.bShowFormula := (Balise=';E');
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_G(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_G);
  theDisplayOpt.bNoShowDefGridlines := (Balise=';G');
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_H(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_S);
  theDisplayOpt.bNoShowHeader := (Balise=';H');
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_K(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_K);
  theDisplayOpt.bShowCommas := (Balise=';K');
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_M(var theLine: string);
var Balise,Ligne,Police : string;
    PCur : PChar;
    Idx : integer;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_M);
  Police := ReadTokenPlus(Ligne,';',true);
  theLine := Ligne;
  PCur := @Police[1];
  Idx := StrToInt(ReadNumberSt(PCur));
  if Idx<=High(TabFontes) then theCurrentFont.Assign(TabFontes[Idx]);
end;

procedure TSlkGeneral.ParseF_N(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_N);
  theIdFontSize := ReadTokenPlus(Ligne,';',true);
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_P(var theLine: string);
var Balise,Ligne : string;
    idxFont : integer;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_P);
  idxFont := ReadTokenPlus(Ligne,';',false);
//theCurrentMask := TabFormats[idxFont];
  theActiveCell.theMask := TabFormats[idxFont];
  theLine := Ligne;
end;

{
procedure TSlkGeneral.ParseF_R(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_R);
  theEntireRow := ReadTokenPlus(Ligne,';',true)-1;
  theLine := Ligne;
end;
}

{
procedure TSlkGeneral.ParseF_X(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_X);
  theCurrentCol := ReadTokenPlus(Ligne,';',true);
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_Y(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_F_Y);
  theCurrentRow := ReadTokenPlus(Ligne,';',true);
  theLine := Ligne;
end;
}

procedure TSlkGeneral.Parse_X(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_X);
  theCurrentCol := ReadTokenPlus(Ligne,';',false)-1;
  theLine := Ligne;
end;

procedure TSlkGeneral.Parse_Y(var theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_Y);
  theCurrentRow := ReadTokenPlus(Ligne,';',false)-1;
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseF_FD(var theLine: string; theBalise : string);
var Balise,Ligne,Format,Dim : string;
    PCur : PChar;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,theBalise);
  Format := ReadTokenPlus(Ligne,';',false);
  theLine := Ligne;
  if High(theSheet)=0 then exit;
  PCur := @Format[1];
  theActiveCell.theFormat.strType := Balise[2];
  theActiveCell.theFormat.strFormat := PCur^;
  Inc(PCur);
  Dim := ReadNumberSt(PCur);
  theActiveCell.theFormat.iNbDigits := StrToInt(Dim);
  theActiveCell.theFormat.strAlignment := PCur^;
//normalement on pourrait s'arreter la pour le type F... reserve pour D :
  Inc(PCur);
  Dim := ReadNumberSt(PCur);
  theActiveCell.theFormat.iwidth := StrToInt(Dim);
end;

function TSlkGeneral.GetWitdh(laColonne : integer) : integer;
begin
  result := theSlkCols[laColonne].iWidth * (theSheet[0][laColonne].theFont.Size);
end;

procedure TSlkGeneral.ParseC(theLine: string);
var Balise,Ligne : string;
    ARow, ACol : integer;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_C);
  GetInfoCible(Ligne);
  while Ligne<>'' do
  begin
    if Pos(SLK_C_K,Ligne)=1 then ParseC_K(Ligne) else
    if Pos(SLK_C_E,Ligne)=1 then ParseC_E(Ligne) else
    if Pos(SLK_X,Ligne)=1 then Parse_X(Ligne) else
    if Pos(SLK_Y,Ligne)=1 then Parse_Y(Ligne) else
    if Pos(SLK_C_A,Ligne)=1 then ParseC_A(Ligne) else
{
    if Pos(SLK_C_C,Ligne)=1 then ParseC_C(Ligne) else
    if Pos(SLK_C_R,Ligne)=1 then ParseC_R(Ligne) else
    if Pos(SLK_C_G,Ligne)=1 then ParseC_G(Ligne) else
    if Pos(SLK_C_D,Ligne)=1 then ParseC_D(Ligne) else
    if Pos(SLK_C_S,Ligne)=1 then ParseC_S(Ligne) else
    if Pos(SLK_C_N,Ligne)=1 then ParseC_N(Ligne) else
    if Pos(SLK_C_P,Ligne)=1 then ParseC_P(Ligne) else
    if Pos(SLK_C_H,Ligne)=1 then ParseC_H(Ligne) else
    if Pos(SLK_C_M,Ligne)=1 then ParseC_M(Ligne) else
    if Pos(SLK_C_T,Ligne)=1 then ParseC_T(Ligne) else
}
    SkipBalise(Ligne,SLK_UNKOWN);
  end;
  //traite les infos parsées... affectation a l'objet theSheet...
  if UnselectedRow(theCurrentRow) or UnselectedCol(theCurrentCol) then exit;
  ARow := theCurrentRow; TraductRow(ARow,-1);
  ACol := theCurrentCol; TraductCol(ACol,-1);
  AffecteValCell(ARow,ACol,theActiveCell);
end;

procedure TSlkGeneral.ParseC_K(var theLine: string);
var Balise,Ligne,Valeur : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_C_K);
  Valeur := ReadTokenPlus(Ligne,';',false);
  SetValCurrentCell(Valeur);
  theLine := Ligne;
end;

procedure TSlkGeneral.ParseC_E(var theLine: string);
var Balise,Ligne,Nom : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_C_E);
  Nom := ReadTokenPlus(Ligne,';',false);
  SetFormuleCurrentCell(Nom);
  theLine := Ligne;
end;

// gestion d'un commentaire
procedure TSlkGeneral.ParseC_A(var theLine: string);
var Balise,Ligne,St : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_C_A);
  St := ReadTokenPlus(Ligne,';',false);
  SetCommentCurrentCell(St);
  theLine := Ligne;
end;

{
procedure TSlkGeneral.SetNomCurrentCell(theNom : string);
begin
  theActiveCell.strName := theNom;
end;
}

procedure TSlkGeneral.SetValCurrentCell(theValeur : string);
begin
  theActiveCell.strValue := TraduitValeur(theValeur);
end;

procedure TSlkGeneral.SetFormuleCurrentCell(theFormule : string);
begin
  theActiveCell.strFormule := theFormule;
end;

procedure TSlkGeneral.SetCommentCurrentCell(theComment : string);
begin
  theActiveCell.strComment := theComment;
end;

procedure TSlkGeneral.InitActiveCell;
begin
  theActiveCell.strName := '';
  theActiveCell.strValue := '';
  theActiveCell.varValue := unassigned;
  theActiveCell.theType := '';
  theActiveCell.strFormule := '';
  theActiveCell.strComment := '';
  Fillchar(theActiveCell.theMask,SizeOf(TPRec),#0) ;
  Fillchar(theActiveCell.theDescription,SizeOf(TCRec),#0) ;
  Fillchar(theActiveCell.theFormat,SizeOf(TSingleFormatCell),#0) ;
  Fillchar(theActiveCell.theBorder,SizeOf(TBorderCell),#0) ;
end;

procedure TSlkGeneral.GetActiveCell(Arow,ACol : integer);
begin
  theActiveCell.strName := theSheet[ARow][ACol].strName;
  theActiveCell.strValue := theSheet[ARow][ACol].strValue;
  theActiveCell.varValue := unassigned;
  theActiveCell.theType := theSheet[ARow][ACol].theType;
  theActiveCell.strFormule := theSheet[ARow][ACol].strFormule;
  theActiveCell.strComment := theSheet[ARow][ACol].strComment;
  theActiveCell.theMask := theSheet[ARow][ACol].theMask;
  theActiveCell.theDescription := theSheet[ARow][ACol].theDescription;
  theActiveCell.theFormat := theSheet[ARow][ACol].theFormat;
  theActiveCell.theBorder := theSheet[ARow][ACol].theBorder ;
end;

procedure TSlkGeneral.SetRowProperties(Quoi : string);
var i,iRow : integer;
begin
  iRow := theEntireRow;
  TraductRow(iRow,-1);
  for i:=Low(theSheet[iRow]) to High(theSheet[iRow]) do
  begin
//    if UnselectedCol(i) then continue;//EPZ 07/03/01
    if Quoi = SLK_F then AffecteFormatCell(iRow,i,theActiveCell)
    else if Quoi = SLK_C then AffecteValCell(iRow,i,theActiveCell);
  end;
end;

procedure TSlkGeneral.SetColProperties(Quoi : string);
var i,iCol : integer;
begin
  iCol := theEntireCol;
  TraductCol(iCol,-1);
  for i:=Low(theSheet) to High(theSheet) do
  begin
//    if UnselectedRow(i) then continue;//EPZ 07/03/01
    if Quoi = SLK_F then AffecteFormatCell(i,iCol,theActiveCell)
    else if Quoi = SLK_C then AffecteValCell(i,iCol,theActiveCell);
  end;
end;

procedure TSlkGeneral.AffecteFormatCell(ARow, ACol : integer ; theActiveCell : TSlkCell);
begin
  theSheet[ARow][ACol].strName := theActiveCell.strName;
  theSheet[ARow][ACol].strFormule := theActiveCell.strFormule;
  theSheet[ARow][ACol].strComment := theActiveCell.strComment;
  theSheet[ARow][ACol].theMask := theActiveCell.theMask;
  theSheet[ARow][ACol].theType := GetTheTipe(theSheet[ARow][ACol]);
  theSheet[ARow][ACol].theDescription := theActiveCell.theDescription;
  theSheet[ARow][ACol].theFormat := theActiveCell.theFormat;
  theSheet[ARow][ACol].theFont.Assign(theActiveCell.theFont);
  theSheet[ARow][ACol].theBorder := theActiveCell.theBorder;
end;

function TSlkGeneral.GetTheTipe(theCell : TSlkCell) : Char;
begin
  result := #0;
  if TipeDate(theCell.theMask.strFormat) then result := 'D' else
  if TipeHeure(theCell.theMask.strFormat) then result := 'H' else
  if TipeNombre(theCell.theMask.strFormat) then result := 'R';
end;

function TSlkGeneral.SetTheVariant(theCell : TSlkCell) : variant;
var jj,mm,aa,hh,ss,mss : word;
begin
  if Pos('"',theCell.strValue)>0 then
    result := theCell.strValue
  else
  if (theCell.strValue<>'') and (theCell.theType = 'D') and (IsNumeric(theCell.strValue,true)) then //DATE
  begin
    DecodeDate(StrToInt(theCell.strValue),aa,mm,jj);
    result := EncodeDate(aa,mm,jj);
  end
  else
  if (theCell.strValue<>'') and (theCell.theType='H') and (IsNumeric(theCell.strValue,true)) then //HEURE
  begin
    DecodeTime(StrToInt(theCell.strValue),hh,mm,ss,mss);
    result := EncodeTime(hh,mm,ss,mss);
  end
  else //NOMBRE
  if (theCell.strValue<>'') and (theCell.theType='R') and (IsNumeric(theCell.strValue,true)) then
    result := Valeur(theCell.strValue)
  else result := theCell.strValue;
end;

procedure TSlkGeneral.AffecteValCell(ARow, ACol : integer ; theActiveCell : TSlkCell);
begin
//EPZ 07/03/01  if iMinRowAffected =0 then iMinRowAffected := ARow else iMinRowAffected := Min(iMinRowAffected,ARow);
//EPZ 07/03/01  iMaxRowAffected := Max(iMaxRowAffected,ARow);
//EPZ 07/03/01  theSlkRows[ARow].bAffected := true;
  theSheet[ARow][ACol].strValue := SetTheValue(theActiveCell.strValue);
  theSheet[ARow][ACol].varValue := SetTheVariant(theSheet[ARow][ACol]);
end;

function TSlkGeneral.SetTheValue(theValue : string) : string;
var St : string;
begin
  St := theValue;
  if (St<>'') and (St[1]='"') and (St[Length(St)]='"') then St := Copy(theValue,2,Length(St)-2);
  result:=St;
end;

procedure TSlkGeneral.ParseNN(theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  if Pos(SLK_REF,Ligne)<>0 then exit; //erreur dans la liaison ; x feuilles
  Balise := SkipBalise(Ligne,SLK_NN);
  InitTheCurrentName;
  while (Ligne<>'') and (Ligne<>#0) do
  begin
    if Pos(SLK_NN_N,Ligne)=1 then ParseNN_N(Ligne) else
    if Pos(SLK_NN_E,Ligne)=1 then ParseNN_E(Ligne) else
//    if Pos(SLK_NN_G,Ligne)=1 then ParseNN_G(Ligne) else
//    if Pos(SLK_NN_K,Ligne)=1 then ParseNN_K(Ligne) else
//    if Pos(SLK_NN_F,Ligne)=1 then ParseNN_F(Ligne) else
    SkipBalise(Ligne,SLK_UNKOWN);
  end;
  //traite les infos parsées... affectation a l'objet theNames...
  SetTheNamesProperties;
end;

procedure TSlkGeneral.InitTheCurrentName;
begin
  theCurrentName.iRow := -1;
  theCurrentName.iCol := -1;
  theCurrentName.strError := '';
  theCurrentName.strName := '';
  theCurrentName.strExpr := '';
  theCurrentName.strRunableMacro := '';
  theCurrentName.strOrdinaryMacro := '';
  theCurrentName.bUsableAsFunction := false;

  theRectCells := Rect(0,0,0,0);

end;

procedure TSlkGeneral.SetTheNamesProperties;
var priority,i,j,iRowDeb,iRowFin,iColDeb,iColFin,ARow,ACol : integer;
    bOk,bRectIsNull,bRectIsOk : boolean;
begin
  bRectIsNull := (theRectCells.Top=0) and (theRectCells.bottom=0) and
                 (theRectCells.Left=0) and (theRectCells.Right=0) ;
  bRectIsOk := ( ((theRectCells.Top=theRectCells.bottom) and //idem ligne ou colonne
                 (theRectCells.Left<>theRectCells.Right)) or
                 ((theRectCells.Top<>theRectCells.bottom) and //idem ligne ou colonne
                 (theRectCells.Left=theRectCells.Right)) );
  if (bRectIsNull) and (theCurrentName.iRow=-1) and (theCurrentName.iCol=-1) then exit;
  if theCurrentName.strExpr<>'' then exit; // reference entre feuilles...
  if theCurrentName.strError<>'' then exit; // reference entre feuilles...
  iColDeb := -1; iColFin := -1; iRowDeb := -1; iRowFin := -1;
  bOk := false;

  priority := -1;//05/04/01
  if (not bRectIsOk) and (theCurrentName.iRow<>-1) and (theCurrentName.iCol<>-1) then
    priority := 3; // une cellule référencée
  if (bRectIsNull) and (theCurrentName.iRow=-1) then
    begin
    iColDeb := theCurrentName.iCol-1; TraductCol(iColDeb,-1);
    iColFin := iColDeb;
    iRowDeb := Low(theSheet) ;
    iRowFin := High(theSheet) ;
    bOk := true;
    priority := 1; // une colonne référencée
    end
  else if (bRectIsNull) and (theCurrentName.iCol=-1) then
    begin
    iRowDeb := theCurrentName.iRow-1; TraductRow(iRowDeb,-1);
    iRowFin := iRowDeb;
    iColDeb := Low(theSheet[0]);
    iColFin := High(theSheet[0]);
    bOk := true;
    priority := 1; // une ligne référencée
    end
  else if (bRectIsOk) and (not bRectIsNull) then
    begin
    iRowDeb := theRectCells.Top-1; TraductRow(iRowDeb);
    iRowFin := theRectCells.Bottom-1; TraductRow(iRowFin);
    iColDeb := theRectCells.Left-1; TraductCol(iColDeb);
    iColFin := theRectCells.Right-1; TraductCol(iColFin);
    bOk := true;
    priority := 2; // un groupe de cellules en ligne ou colonne référencée
    end;
  if not bOk then exit;
  for i:=iRowDeb to iRowFin do
  begin
    ARow := i;
    TraductRow(ARow);
    if UnselectedRow(ARow) then continue;
    for j:=iColDeb to iColFin do
      begin
      ACol := j;
      TraductCol(ACol);
      if UnselectedCol(ACol) then continue;
      if (theNames[i][j].strName<>'') and (priority<theNames[i][j].iAffectNamePriority) then continue;
      theNames[i][j].iAffectNamePriority := priority;
      theNames[i][j].iRow := i;
      theNames[i][j].iCol := j;
      theNames[i][j].strError := theCurrentName.strError;
      theNames[i][j].strName := theCurrentName.strName;//+'_'+IntToStr(j-1);
      theNames[i][j].strExpr := theCurrentName.strExpr;
      theNames[i][j].strRunableMacro := theCurrentName.strRunableMacro;
      theNames[i][j].strOrdinaryMacro := theCurrentName.strOrdinaryMacro;
      theNames[i][j].bUsableAsFunction := theCurrentName.bUsableAsFunction;
      end;
    end;
end;

procedure TSlkGeneral.ParseNN_N(var theLine : string); // recup du nom de la cellule
var St,Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_NN_N);
  St := ReadTokenPlus(Ligne,';',false);
  theCurrentName.strName := TraduitValeur(St);
  theLine := Ligne;
end;

// ATTENTION : de la forme
//             ;E'chemin_d_un_doc_ou_feuille'!CoordonnéeCellule
//             ;E'chemin_d_un_doc_ou_feuille'!#REF! si pb de liaison
//             ;E"Libellé" : ?????
//             ;ECn        : nom affecté a toute la colonne n
//             ;ERn        : nom affecté a toute la ligne n
//             ;ERnCp      : nom affecté a la ligne n et colonne p
//             ;ECnCp      : nom affecté aux groupe de colonnes n a p
//             ;ERnRp      : nom affecté aux groupe de lignes n a p
procedure TSlkGeneral.ParseNN_E(var theLine : string); // recup expression....
var Balise,Ligne : string;
    PCur : PChar;
    bBlocCells,bRows : boolean;

    function IsValidePCur (c : pchar): boolean;
    begin
    result := (PCur^<>#0) and (Pcur^<>c);
    end;

    procedure Add;
    begin
    if PCur^<>#0 then Inc(PCur);
    end;

begin
  bBlocCells := false; bRows := false;
  theRectCells := Rect(0,0,0,0);
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_NN_E);
  PCur := @Ligne[1];
  while (PCur^<>#0) do
    begin
    if Pcur^=#13 then Inc(PCur) else
    if Pcur^=#10 then Inc(PCur) else
    if Pcur^='''' then
      begin
      Add;
      while IsValidePCur('''') do
        begin
        theCurrentName.strExpr := theCurrentName.strExpr + PCur^;
        Add;
        end;
      Add; // on saute '
      Add;  // on saute le !
      end
    else if PCur^='"' then
      begin
      theCurrentName.strExpr := '';
      Add;
      while IsValidePCur('"') do
        begin
        theCurrentName.strExpr := theCurrentName.strExpr + PCur^;
        Add;
        end;
      Add;  // on saute le "
      end
    else if PCur^='R' then // de la forme Rn ou RnCm
         begin
         Add;
         bRows := true;
         theCurrentName.iRow := StrToInt(ReadNumberSt(PCur));
         if bBlocCells then
           begin
           theRectCells.Bottom := theCurrentName.iRow;
           theCurrentName.iRow := -1;
           end;
         end
    else if PCur^='C' then // de la forme Rn ou RnCm
         begin
         Add;
         theCurrentName.iCol := StrToInt(ReadNumberSt(PCur));
         if bBlocCells then
           begin
           theRectCells.Right := theCurrentName.iCol;
           theCurrentName.iCol := -1;
           end;
         end
    else if PCur^=':' then // indique la présence d'un bloc
         begin
         bBlocCells := true;
         theRectCells.Top := theCurrentName.iRow;
         theRectCells.Left := theCurrentName.iCol;
         Add;
         end
    else if PCur^='!' then // indique la présence d'un bloc
      begin
      Add;  // on saute le !
      end
    else // autre cas !.....
      begin
      theCurrentName.strExpr := '';
      Add;
      while IsValidePCur('!') do
        begin
        theCurrentName.strExpr := theCurrentName.strExpr + PCur^;
        Add;
        end;
      Add;  // on saute le !
      end;
    end;
// on ne gère pas pour l'instant les blocs et les titres de lignes/
// seules les colonnes sont gérées
    if (bBlocCells) or (bRows) then InitTheCurrentName;
//
  theLine := PCur^;
end;



function TSlkGeneral.GetFonte(ARow, ACol: integer): String;
var st : string;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
  st := Format('%s (%d) - ',[theSheet[ARow][ACol].theFont.Name,theSheet[ARow][ACol].theFont.Size]);
  if fsBold in theSheet[ARow][ACol].theFont.Style then st := st + 'Bold  - ';
  if fsItalic in theSheet[ARow][ACol].theFont.Style then st := st + 'Italic - ';
  if fsUnderline in theSheet[ARow][ACol].theFont.Style then st := st + 'Underline - ';
  result := st;
end;

function TSlkGeneral.GetFormat(ARow, ACol: integer): String;
//var st : string;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
  result := Format('(%s) - %s - %d - %s',[theSheet[ARow][ACol].theFormat.strType,
                                          theSheet[ARow][ACol].theFormat.strFormat,
                                          theSheet[ARow][ACol].theFormat.iNbDigits,
                                          theSheet[ARow][ACol].theFormat.strAlignment]);
end;

function TSlkGeneral.GetName(ARow, ACol: integer): String;
var st : string;
begin
  st := theNames[ARow][ACol].strName;{ +' - '+
        theNames[ARow][ACol].strExpr +' - '+
        theNames[ARow][ACol].strError;}
  if st = '' then
    begin
      if ACol>=26 then begin St := letters[(ACol div 26)-1]; Dec(ACol,26); end;
      St := St + letters[ACol mod 26];
    end;
  result := st;
end;

function TSlkGeneral.GetValue(ARow, ACol: integer): variant;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
  result := theSheet[ARow][ACol].varValue;
end;

function TSlkGeneral.GetFormule(ARow, ACol: integer): String;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
  result := theSheet[ARow][ACol].strFormule;
end;

function TSlkGeneral.GetComment(ARow, ACol: integer): String;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
  result := theSheet[ARow][ACol].strComment;
end;

function TSlkGeneral.GetMask(ARow, ACol: integer): String;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
  result := theSheet[ARow][ACol].theMask.strFormat;
end;

function TSlkGeneral.GetBordure(ARow, ACol: integer): String;
var st : string;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
  if theSheet[ARow][ACol].theBorder.bGridLinesTop then st := 'Top -';
  if theSheet[ARow][ACol].theBorder.bGridLinesBottom then st := St + 'Bottom -';
  if theSheet[ARow][ACol].theBorder.bGridLinesLeft then st := St + 'Left -';
  if theSheet[ARow][ACol].theBorder.bGridLinesRight then st := St + 'Right -';
  result := St;
end;

procedure TSlkGeneral.ParseE(theLine: string);
var Balise,Ligne : string;
begin
  Ligne := theLine;
  Balise := SkipBalise(Ligne,SLK_E);
end;

function TSlkGeneral.GetType(ARow, ACol : integer) : string;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
result := theSheet[ARow][ACol].theType;
end;

function TSlkGeneral.GetAlign(ARow, ACol : integer) : TAlignment;
begin
{
result:=unassigned;
  if (UnselectedRow(ARow)) or (UnselectedCol(ACol)) then exit; //EPZ 07/03/01
}
if theSheet[ARow][ACol].theFormat.strAlignment='C' then result := taCenter else
if theSheet[ARow][ACol].theFormat.strAlignment='L' then result := taLeftJustify else
if theSheet[ARow][ACol].theFormat.strAlignment='R' then result := taRightJustify else
if (theSheet[ARow][ACol].theType='R') then result := taRightJustify else
result := taLeftJustify;
end;

{
function TSlkGeneral.ParseFile: boolean;
begin
;
end;
}

//EPZ 07/03/01 function TSlkGeneral.RowAffected (ARow : integer) : boolean;
function TSlkGeneral.UnselectedRow (ARow : integer) : boolean;
begin
//EPZ 07/03/01 result := not ((ARow >= iMinRowAffected) and (ARow <= iMaxRowAffected));
result := (ARow<theSheetBounds.theRect.Top) or (ARow>theSheetBounds.theRect.Bottom);
end;

function TSlkGeneral.UnselectedCol(ACol : integer) : boolean; //EPZ 07/03/01
begin
result := (ACol<theSheetBounds.theRect.Left) or (ACol>theSheetBounds.theRect.Right);
end;

procedure TSlkGeneral.TraductRow(var ARow : integer ; Sens : integer = 1);
begin
if Sens=-1 then Dec(ARow,theSheetBounds.theRect.Top) else Inc(ARow,theSheetBounds.theRect.Top);
end;

procedure TSlkGeneral.TraductCol(var ACol : integer ; Sens : integer = 1);
begin
if Sens=-1 then Dec(ACol,theSheetBounds.theRect.Left) else Inc(ACol,theSheetBounds.theRect.Left);
end;

procedure TSlkGeneral.GetInfos(ARow, ACol: integer; var leNom: string; var laValeur: variant);
var st : string;
begin
  laValeur := theSheet[ARow][ACol].varValue;
  st := theNames[ARow][ACol].strName;
  if st = '' then
    begin
      if ACol>=26 then begin St := letters[(ACol div 26)-1]; Dec(ACol,26); end;
      St := St + letters[ACol mod 26];
    end;
  leNom := st;
end;

// Recup format de la cellule/colonne/ligne eventuellement deja defini
procedure TSlkGeneral.GetInfoCible(Ligne : string);
var ACol, ARow : integer;
begin
  ACol := -1; ARow := -1;
  if (Pos(SLK_X,Ligne)>0) or (Pos(SLK_Y,Ligne)>0) then
    begin
    GetInfoCurrent(Ligne,ARow,ACol)
    end
    else if (Pos(SLK_F_C,Ligne)>0) or (Pos(SLK_F_R,Ligne)>0) then
    begin
    GetInfoEntire(Ligne,ARow,ACol);
    end;

  if (ARow=-1) and (ACol=-1) then exit;
  if (ARow=-1) then else if UnselectedRow(ARow) then exit;
  if (ACol=-1) then else if UnselectedCol(ACol) then exit;

  if (ARow=-1) then ARow := 0 else TraductRow(ARow,-1);
  if (ACol=-1) then ACol := 0 else TraductCol(ACol,-1);
  
  GetActiveCell(Arow,ACol);
end;

procedure TSlkGeneral.GetInfoCurrent(Ligne : string ; var ARow,ACol : integer);
var iPos : integer;
    St : string;
begin
  iPos := Pos(SLK_X,Ligne)+2;
  if iPos>2 then
    begin
    St := Copy(Ligne,iPos,Length(Ligne));
    ACol := ReadTokenPlus(St,';',false)-1;
    end;
  iPos := Pos(SLK_Y,Ligne)+2;
  if iPos>2 then
    begin
    St := Copy(Ligne,iPos,Length(Ligne));
    ARow := ReadTokenPlus(St,';',false)-1;
    end;
  if ACol=-1 then ACol := theCurrentCol;
  if ARow=-1 then ARow := theCurrentRow;
end;

procedure TSlkGeneral.GetInfoEntire(Ligne : string ; var ARow,ACol : integer);
var iPos : integer;
    St : string;
begin
  iPos := Pos(SLK_F_C,Ligne)+2;
  if iPos>2 then
    begin
    St := Copy(Ligne,iPos,Length(Ligne));
    ACol := ReadTokenPlus(St,';',false)-1;
    end;
  St := Ligne;
  iPos := Pos(SLK_F_R,Ligne)+2;
  if iPos>2 then
    begin
    St := Copy(Ligne,iPos,Length(Ligne));
    ARow := ReadTokenPlus(St,';',false)-1;
    end;
  if ACol=-1 then ACol := theEntireCol;
  if ARow=-1 then ARow := theEntireRow;
end;


end.

