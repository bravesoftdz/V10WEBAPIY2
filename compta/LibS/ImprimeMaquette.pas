unit ImprimeMaquette;

interface

uses
 Windows, Messages, Graphics, Dialogs, classes, sysutils,Printers;


const
//  HT             =  #09 ;
  HT             =  '|' ;

function ChargeFile (Fichier : string; List: TStringList) : Boolean;
//function ControlTextToPrinter(Fichier : string): Boolean;
function ControlTextToPrinter(Fichier : string; Orientation: TPrinterOrientation =poLandscape): Boolean;

implementation

function GetColumnsCount(S: TStrings; i : integer): Integer;
var
  p: PChar;
begin
  Result := 0;
  if S.Count=0 then Exit;
  p := PChar(S[i]);
  repeat
    p := StrScan(p, HT);
    if p=nil then Break;
    Inc(p);
    Inc(Result);
  until false;
  Inc(Result);
end;
function GetParam(var Start: PChar; Separator: Char): string;
var
  p, pQ: PChar;
  Len : Integer;
  Quotes: Boolean;
label Loop;
begin
  while (Start^ = ' ') do Inc(Start);
  Quotes := (Start^ = '"');
  if Quotes then begin
    Inc(Start);
    p := Start;
Loop:
    pQ := StrScan(p, '"');
    if (pQ <> nil) then begin
      p := pQ;   Inc(p);
      if (p^ = '"') then begin
        Inc(p);
        goto Loop;
      end;
      Len := pQ - Start;
      p := StrScan(pQ, Separator);
      if (p = nil) then p := pQ;
      Inc(p);
    end else begin
      Len := StrLen(Start);
      p := Start;
      Inc(p, Len);
    end;
    Result := System.Copy(string(Start), 1, Len);
  end else begin
    p := StrScan(Start, Separator);
    if (p <> nil) then begin
      Len := p - Start;
      Inc(p);
    end else begin
      Len := StrLen(Start);
      p := Start;
      Inc(p, Len);
    end;
    Result := Trim(System.Copy(string(Start), 1, Len));
  end;
  Start := p;
end;



type
  Integers = array of Integer;
function GetWidths(S: TStrings; Canvas: TCanvas; var Count: Integer; ind : integer): Integers;
var
  i, j: Integer;
  p: PChar;
  Field: string;
  w: Integer;
begin
  Count := GetColumnsCount(S, ind);
  SetLength(Result, Count);
  if Count=0 then Exit;
  for j := 0 to Pred(Count) do begin
    Result[j] := 0;
  end;
  for i := 0 to Pred(S.Count) do begin
    p := PChar(S[i]);
    for j := 0 to Pred(Count) do begin
      Field := GetParam(p, HT);
      w := Canvas.TextWidth(Field);
      if Result[j] < w then Result[j] := w;
    end;
  end;
end;


//function ControlTextToPrinter(Fichier : string): Boolean;
function ControlTextToPrinter(Fichier : string; Orientation: TPrinterOrientation =poLandscape): Boolean;
var
  Lines: TStringList;
  i, j: Integer;
  x, y: Integer;
  LineHeight: Integer;
  Line: string;
  PixelsPerInch: Integer;
  W: Integers;
  ColCount: Integer;
  p: PChar;
  Col: string;
  SepWidth: Integer;
  TimeStamp: TDateTime;
  Indice     : integer;
begin
  W := (nil);
  Lines := TStringList.Create;
  Result := ChargeFile (Fichier, Lines);
//  Printer.Orientation := poLandscape; // pour impression horizontal
  Printer.Orientation := Orientation; // pour impression horizontal

  if (Result and (Lines.Count > 0)) then
  begin
//    TimeStamp := UsTime(NowH);
    TimeStamp := Now;
    with Printer do
    begin
      if (Orientation = poLandscape) then
        PixelsPerInch := Round(PageHeight / 8)
      else
        PixelsPerInch := Round(PageHeight / 11);
      LineHeight := Round(PixelsPerInch / 8);
      BeginDoc;
      y := 0;
{
      W := GetWidths(Lines, Canvas, ColCount);
      SepWidth := Canvas.TextWidth('/');
}
// parcours sur nombre de ligne du fichier
      for i := 0 to Pred(Lines.Count) do
      begin
        if (y = 0) then  // édition entete
        begin
          Line := Format('Page %d - %s', [PageNumber, DateTimeToStr(TimeStamp)]);
          Canvas.font.Style := [fsItalic];
          Canvas.TextOut((PageWidth - Canvas.TextWidth(Line)) div 2, y, Line);
          y := y + 2*LineHeight;
        end;
        x := 0;
        p := PChar(Lines[i]);
        W := GetWidths(Lines, Canvas, ColCount, i);
        SepWidth := Canvas.TextWidth('/');
        Indice := 0;
        // parcours colonne par colonne et édition
        for j := 0 to Pred(ColCount) do
        begin
          // HT séparateur de la colonne, récupération contenu de la colonne
          Col := GetParam(p, HT);
          if (Col = 'G') and (j=1) then inc(Indice);
          if ((Indice <> 0)  and (j=3)) or (y=1) then
                    Canvas.Font.Style := [fsItalic,fsBold]
          else
                    Canvas.Font.Style := [fsItalic];
          // élimination de la colonne 1 et 2 pour ne pas imprimer
          if (j=1) or (j=2) then continue;
          if (Col = 'R') and (j = 0) then Canvas.TextOut(x, y, 'Rubrique')
          else
          if (Col = 'C') and (j = 0) then Canvas.TextOut(x, y, 'Commentaire')
          else
          if (Col = 'F') and (j = 0) then Canvas.TextOut(x, y, 'Formule')
          else
          Canvas.TextOut(x, y, Col);
          x := x + SepWidth + W[j];
        end;
        y := y + LineHeight;
        if (y >= PageHeight) then
        begin
          NewPage;
          y := 0;
        end;
      end;
      EndDoc;
    end;
  end;
  Lines.Free;
end;

function ChargeFile (Fichier : string; List: TStringList) : Boolean;
var
  F       : TextFile;
  S       : string;
begin
  AssignFile(F, Fichier);
  Reset(F);
  if EOF(F) then
  begin
    Result := FALSE;
    exit;
  end;
  While Not EOF(F) do
  begin
    Readln(F, S);
    List.Add(S);
  end;
  CloseFile(F);
  Result := TRUE;
end;

end.
