unit UtilScreen;

interface
Uses StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Types,
     AffaireUtil,
     windows,
     DBGrids,
     UTOB,
     MsgUtil,
     Grids,
     Graphics,
     UTOF ;
type

  TAlignementVertical=(alVTop,alVCenter,alVBottom);
  TAlignementHorizontal=(alHLeft,alHCenter,alHRight);
  TJustification=(JustLeft,JustCenter,JustRight);

  procedure DessinerLaCellule(Fliste : THGrid; ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
  procedure DessineTexte(AString:string;ACanvas:TCanvas;ARect: TRect; AlignementHorizontal:TAlignementHorizontal;AlignementVertical:TAlignementVertical;TextJustification:TJustification);
  Function TextSize(Phrase : string; Police : TFont = nil) : TPoint;
  Procedure AfficheGrille(FListe : THGrid; TOBGrille : TOB; GestGrille : AffGrille);
  Procedure PassagePleinEcran(var XX : THform);

implementation


Procedure PassagePleinEcran(var XX : THform);
var RR : Trect;
begin
//  XX.position := application.mainform.position;
	SystemParametersInfo (SPI_GETWORKAREA, 0, @RR, 0);
  XX.SetBounds (Application.MainForm.left, Application.MainForm.Top, Application.MainForm.BoundsRect.right  - Application.MainForm.left, screen.WorkAreaHeight - 20);
  XX.ClientHeight := Screen.WorkAreaHeight ;
  XX.ClientWidth := Screen.WorkAreaWidth;
end;

Procedure AfficheGrille(FListe : THGrid; TOBGrille : TOB; GestGrille : AffGrille);
Var StSql : string;
    I		        : integer;
    J				    : integer;
Begin

  if Not Assigned(TobGrille) then exit;

  with Fliste do
  Begin
    if TOBGrille.detail.count = 0 then
    Begin
      RowCount := 2;
      For J := 1 to ColCount - 1 do
      Begin
        Cells[J, 1] := '';
      end;
      end
      else
      Begin
        RowCount := TOBGrille.detail.count + 1;
        For i := 0 to TOBGrille.detail.count - 1 Do
        Begin
          Cells[0, I] := '';
          StSql := GestGrille.ColGAppel;
          TOBGRille.detail[i].PutLigneGrid(Fliste,i+1,false,false,stSql);
        end;
      row := 1;
    end;
  end;
end;

Function TextSize(Phrase : string; Police : TFont = nil) : TPoint;
var DC  : HDC;
    X   : Integer;
    Rect: TRect;
    C   : TBitmap;
begin

  C := TBitmap.create;

  if police <> nil then  C.canvas.Font := police;

  Rect.Left := 0;
  Rect.Top:=0;
  Rect.Right:=0;
  Rect.Bottom:=0;

  DC := GetDC(0);

  C.Canvas.Handle := DC;

  DrawText(C.Canvas.Handle, PChar(Phrase), -1, Rect, (DT_EXPANDTABS or DT_CALCRECT));

  C.Canvas.Handle := 0;

  ReleaseDC(0, DC);

  result.X:=Rect.Right-Rect.Left;
  result.Y:=Rect.Bottom-Rect.Top;

  C.Free;

end;

procedure DessinerLaCellule(Fliste : THGrid; ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
Var Alignement  : String;
    AlignH      : TAlignementHorizontal;
    AlignV      : TAlignementVertical;
    Justifie    : TJustification;
    TexteCell   : string;
    Arect       : Trect;
    OkImg       : String;
    OkLib       : String;
begin

  Fliste.Font.Name  := 'Tahoma';
  Fliste.Font.Size  := 12;

  if (ACol < Fliste.FixedCols) or (ARow < Fliste.FixedRows) then
  begin
    TexteCell := Fliste.Cells[ACol, Arow];
    ARect     := Fliste.CellRect(ACol, ARow);
    Alignement:= 'C';
  end
  else
  begin
    TexteCell := Fliste.Cells[ACol, Arow];
    ARect     := Fliste.CellRect(ACol, ARow);
    //Alignement:= GestGrille.TabloMEF[1,ACol];;
    //OkImg := GestGrille.TabloMEF[3,ACol];
    //OkLib := GestGrille.TabloMEF[2,ACol];
    //if (OkLib='X') or (okImg='X') then
    //begin
    //  Fliste.ColFormats[ACol] := 'CB=' + Get_Join(GestGrille.TabloMEF[0,ACol]);
    //end;
  end;

  if Alignement = 'C' then
    AlignH := alHCenter
  else if alignement = 'D' then
    AlignH := alHRight
  else if Alignement = 'G' then
    AlignH := alHLeft
  else //par défaut l'ensemble des cellules sont cadrées à gauche...
    AlignH := alHLeft;

  AlignV   := alVCenter;
  Justifie := JustLeft;

  DessineTexte(TexteCell,Fliste.Canvas,Arect,AlignH,AlignV,Justifie);

end;


procedure DessineTexte(AString: string;ACanvas:TCanvas; ARect: TRect; AlignementHorizontal:TAlignementHorizontal; AlignementVertical:TAlignementVertical; TextJustification:TJustification);
var AHeight,AWidth      :integer;
    Rect,oldClipRect    :TRect;
    ATop,ALeft,H,W      :Integer;
    AText               :string;
    JustificationDuTexte:Integer;
    MyRgn               :HRGN;
begin

  with ACanvas do
  begin
    Lock;
    AHeight:=ARect.Bottom-ARect.Top;

    AWidth:=ARect.Right-ARect.Left;
    //on calcule la taille du rectangle dans lequel va tenir le texte
    W:=TextSize(AString,ACanvas.Font).X;
    H:=TextSize(AString,ACanvas.Font).Y;
 
    //on calcule la position (Haut,Gauche) du rectangle dans lequel va tenir le texte
    //en fonction de l'alignement horizontal et vertical choisi
    ATop:=ARect.Top;
    ALeft:=ARect.Left;
 

    case AlignementVertical of
      alVBottom : ATop:=ARect.Bottom-H;
      alVCenter : ATop:=ARect.Top+((AHeight-H) div 2);
      alVTop    : ATop:=ARect.Top;
    end;

    case AlignementHorizontal of
      alHLeft  : ALeft:=ARect.Left;
      alHCenter: ALeft:=ARect.Left+(AWidth-W) div 2;
      alHRight : ALeft:=ARect.Right-W;
    end;

    //Fin du calcul du rectangle, on met le resultat dans Rect
    Rect:=Bounds(ALeft,ATop,W,H);

    //On remplit le rectangle de la zone sinon on voit le texte que delphi à dessiné
    FillRect(ARect);

    //On détermine les paramètres de justification à passer à Windows
    case TextJustification of
      JustLeft  : JustificationDuTexte:=DT_LEFT;
      JustCenter: JustificationDuTexte:=DT_CENTER;
      JustRight : JustificationDuTexte:=DT_RIGHT;
    end;

    //Si le texte est plus grand que notre zone, on prend cette précaution (Clipping)
    with ARect do MyRgn :=CreateRectRgn(Left,Top,Right,Bottom);
    SelectClipRgn(Handle,MyRgn);

    //On dessine le texte
    DrawText(Handle,PChar(AString),-1,Rect,JustificationDuTexte or DT_NOPREFIX or DT_WORDBREAK );

    //On a plus besoin de la zone de clipping
    SelectClipRgn(Handle,0);
    DeleteObject(MyRgn);
    Unlock;
  end;

end;



end.
