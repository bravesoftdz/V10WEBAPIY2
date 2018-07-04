unit UtofRess_Occup;

interface
uses  Controls,Classes,sysutils,
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB,
      graphics,Grids,Windows,SaisUtil , vierge,Messages,AglInit
{$IFDEF EAGLCLIENT}
     ,MaineAGL
{$ELSE}
     ,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}FE_Main
{$ENDIF}
;
Type
     TOF_Ress_Occup = Class (TOF)
     private
        LesColonnes : string ;
        GS : THGRID ;
        TobRess : Tob;
//        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
//        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
//        procedure GSLigneDClick (Sender: TObject);
        procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
     public
        procedure OnArgument(Arguments : String ) ; override ;
        Procedure OnNew  ; override ;        
        procedure OnLoad ; override ;
     END ;

Procedure RTLanceFiche_Ress_Occup(Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation

{ TOF_Ress_Occup }

Procedure RTLanceFiche_Ress_Occup(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Ress_Occup.OnArgument(Arguments: String);
begin
inherited ;
TobRess := LaTob;

LesColonnes :='FIXED;NOMRESSOURCE;LIBRE;MOTIF';
GS:=THGRID(GetControl('GBRESSOURCES'));
//GS.OnRowEnter:=GSRowEnter ;
//GS.OnRowExit:=GSRowExit ;
//GS.OnDblClick:=GSLigneDClick ;
GS.PostDrawCell:= DessineCell;
GS.ColWidths[0]:=15;
GS.ColWidths[1]:=100;
GS.ColWidths[2]:=20;
GS.ColWidths[3]:=100;
GS.ColCount:=4;
//GS.ColFormats[1]:='CB=RTREPRESENTANT';
GS.ColTypes[2] := 'B';
GS.ColAligns[2]:=taCenter;
GS.ColFormats[2]:=intToStr(integer(csCoche));
AffecteGrid(GS,TaConsult) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;
end;

procedure TOF_Ress_Occup.OnLoad;
begin
inherited ;
TOBRess.PutGridDetail(GS,False,False,LesColonnes,True);
end;

procedure TOF_Ress_Occup.OnNew;
begin
inherited ;
end;

procedure TOF_Ress_Occup.DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GS.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GS.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GS.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;

Initialization
registerclasses([TOF_Ress_Occup]);
end.
