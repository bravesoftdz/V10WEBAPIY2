unit UTofPerspective_Var;

interface
uses  Controls,Classes,sysutils,
      HCtrls,HEnt1,UTOF,UTOB, AglInit,
      graphics,Grids,Windows,SaisUtil, vierge
{$IFDEF EAGLCLIENT}
     ,MaineAGL
{$ELSE}
     ,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}db,FE_Main
{$ENDIF}
;
Type
     TOF_PERSP_VAR = Class (TOF)
     private
        LesColonnes,StVariante : string ;
        GS : THGRID ;
        TOBPersp : TOB;
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSLigneDClick (Sender: TObject);
        procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);        
     public
        Action   : TActionFiche ;     
        procedure OnArgument(Arguments : String ) ; override ;
        Procedure OnClose  ; override ;        
        procedure OnLoad ; override ;
     END ;
var
  Cell_Intervenant,Cell_Perspective  : integer;

implementation

{ TOF_PERSP_VAR }

procedure TOF_PERSP_VAR.OnArgument(Arguments: String);
var     x : integer;
        Critere : string ;
        ChampMul,ValMul,StAction : string;
begin
inherited ;
  Critere := Arguments ;
  Repeat
        Critere:=uppercase(ReadTokenSt(Arguments)) ;
        if Critere<>'' then
            begin
            x:=pos('=',Critere);
            if x<>0 then
               begin
               ChampMul:=copy(Critere,1,x-1);
               ValMul:=copy(Critere,x+1,length(Critere));

               if ChampMul='RPE_VARIANTE' then StVariante := ValMul;
               if ChampMul='ACTION' then StAction := ValMul;
               end;
            end;
    until  Critere='';

    if StAction='CONSULTATION' then BEGIN Action:=taConsult ; END ;
    if StAction='MODIFICATION' then BEGIN Action:=taModif ; END ;

    LesColonnes :='FIXED;PRINCIPALE;RPE_DATEMODIF;RPE_PERSPECTIVE;RPE_LIBELLE;RPE_INTERVENANT;RPE_TYPEPERSPECTIV;RPE_ETATPER;RPE_MONTANTPER;RPE_POURCENTAGE;RPE_DATEREALISE';
    GS:=THGRID(GetControl('GVARPERSPECTIVES'));
    GS.OnRowEnter:=GSRowEnter ;
    GS.OnRowExit:=GSRowExit ;
    GS.OnDblClick:=GSLigneDClick ;
    GS.PostDrawCell:= DessineCell;
    GS.ColWidths[0]:=0;
    GS.ColWidths[1]:=15;
    GS.ColWidths[1]:=20;
    GS.ColTypes[1] := 'B';
    GS.ColAligns[1]:=taCenter;
    GS.ColFormats[1]:=intToStr(integer(csCoche));

    GS.ColWidths[2]:=100;
    GS.ColWidths[3]:=70; Cell_perspective:=3;
    GS.ColFormats[3]:='0';
    GS.ColWidths[4]:=160;
    GS.ColWidths[5]:=105;Cell_Intervenant:=5;
    GS.ColWidths[6]:=100;
    GS.ColWidths[7]:=90;
    GS.ColWidths[8]:=100;
    GS.ColWidths[9]:=100;
    GS.ColWidths[10]:=100;

    GS.ColCount:=11;
    GS.ColFormats[5]:='CB=RTREPRESENTANT';
    GS.ColFormats[6]:='CB=RTTYPEPERSPECTIVE';
    GS.ColFormats[7]:='CB=RTETATPERSPECTIVE';
    GS.ColAligns[8]:=taRightJustify;GS.ColFormats[8]:='#,##0';
    GS.ColAligns[9]:=taRightJustify;GS.ColFormats[9]:='#,##0';

    AffecteGrid(GS,taConsult) ;
    TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;

    TOBPersp:=tob.create('Variantes d''une perspective',Nil,-1) ;

end;

procedure TOF_PERSP_VAR.OnLoad;
var Q : TQuery;
    Select : String;
    i : Integer;
    Principale : String;    
begin
inherited ;
TOBPersp.ClearDetail;
Select := 'SELECT RPE_PERSPECTIVE,RPE_DATEMODIF,RPE_LIBELLE,RPE_INTERVENANT,RPE_TYPEPERSPECTIV,RPE_ETATPER,RPE_MONTANTPER,RPE_POURCENTAGE,RPE_DATEREALISE,RPE_OPERATION,RPE_VARIANTE '+
       'FROM PERSPECTIVES ' +
       'WHERE RPE_VARIANTE='+StVariante+' ORDER BY RPE_PERSPECTIVE DESC';

Q:=OpenSQL(Select, True);

TOBPersp.LoadDetailDB('PERSPECTIVES','','',Q,false,true) ;
Ecran.Caption:=TraduireMemoire('Variantes de la proposition : ')+ StVariante ;
Ferme(Q);

if (TOBPersp.Detail.count > 0) then
   begin
   TOBPersp.Detail[0].AddChampSup('PRINCIPALE', True);
   for i:=0 to TOBPersp.Detail.count-1 do
       begin
       Principale := '-';
       if (TOBPersp.detail[i].GetValue('RPE_VARIANTE')= 0) or (TOBPersp.detail[i].GetValue('RPE_VARIANTE') = TOBPersp.detail[i].GetValue('RPE_PERSPECTIVE')) then
           begin
           Principale := 'X';
           end;
       TOBPersp.Detail[i].PutValue('PRINCIPALE', Principale);
       end;
   end;


TOBPersp.PutGridDetail(GS,False,False,LesColonnes,True);
//GS.RowCount:=GS.RowCount+1 ;

end;

procedure TOF_PERSP_VAR.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOF_PERSP_VAR.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOF_PERSP_VAR.GSLigneDClick (Sender: TObject);
var St_Persp : string;
   QQ:TQuery;
begin
  St_Persp := string(GS.CellValues[Cell_Perspective, GS.Row]);
  if (St_Persp = '') then Exit;
  AGLLanceFiche('RT','RTPERSPECTIVES','',St_Persp,ActionToString(Action));
  QQ:=OpenSQL('Select RPE_VARIANTE FROM PERSPECTIVES WHERE RPE_PERSPECTIVE='+ StVariante , true );
  if not QQ.EOF then stVariante:=intToStr(QQ.findfield('RPE_VARIANTE').asInteger);
  ferme(QQ);
  Onload;
end;

procedure TOF_PERSP_VAR.Onclose;
begin
TobPersp.free;
inherited ;
end;

procedure TOF_PERSP_VAR.DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
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
registerclasses([TOF_PERSP_VAR]);
end.
