unit UtofPerspective_his;
// 22/06/07 : suppression double-click (FQ 10657)

interface
uses  Controls,Classes,sysutils,
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB,
      graphics,Grids,Windows,SaisUtil , vierge,Messages
{$IFDEF EAGLCLIENT}
     ,MaineAGL
{$ELSE}
     ,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}FE_Main
{$ENDIF}
;
Type
     TOF_PERSP_HISTO = Class (TOF)
     private
        LesColonnes,StAuxiliaire,StPerspective : string ;
        GS : THGRID ;
        TOBPersp : TOB;
        procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
//        procedure GSLigneDClick (Sender: TObject);
        procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
     public
        Action   : TActionFiche ;     
        procedure OnArgument(Arguments : String ) ; override ;
        Procedure OnNew  ; override ;        
        procedure OnLoad ; override ;
     END ;

const
        Cell_Commercial  : integer = 4;
	// libellés des messages
	TexteMessage: array[1..1] of string 	= (
          {1}        'Pas d''historique pour cette proposition'
          );
implementation

{ TOF_PERSP_HISTO }

procedure TOF_PERSP_HISTO.OnArgument(Arguments: String);
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

               if ChampMul='RPE_AUXILIAIRE' then StAuxiliaire := ValMul;
               if ChampMul='RPE_PERSPECTIVE' then StPerspective := ValMul;
               if ChampMul='ACTION' then StAction := ValMul;
               end;
            end;
    until  Critere='';

    if StAction='CONSULTATION' then BEGIN Action:=taConsult ; END ;
    if StAction='MODIFICATION' then BEGIN Action:=taModif ; END ;
    LesColonnes :='FIXED;RPH_DATEMODIF;RPH_TYPEPERSPECTIV;RPH_LIBELLE;RPH_REPRESENTANT;RPH_OPERATION;RPH_PROJET;RPH_INTERVENANT;C_NOM;RPH_ETATPER;RPH_MONTANTPER;RPH_POURCENTAGE;RPH_DATEREALISE';
    GS:=THGRID(GetControl('GPERSPECTIVES'));
    GS.OnRowEnter:=GSRowEnter ;
    GS.OnRowExit:=GSRowExit ;
//    GS.OnDblClick:=GSLigneDClick ;
    GS.PostDrawCell:= DessineCell;
    GS.ColWidths[0]:=0;
    GS.ColWidths[1]:=100;
    GS.ColWidths[2]:=100;
    GS.ColWidths[3]:=100;
    GS.ColWidths[4]:=100;
    GS.ColWidths[5]:=100;
    GS.ColWidths[6]:=100;
    GS.ColWidths[7]:=120;
    GS.ColWidths[8]:=100;
    GS.ColWidths[9]:=80;
    GS.ColWidths[10]:=100;
    GS.ColWidths[11]:=100;
    GS.ColWidths[12]:=100;
    GS.ColCount:=13;
    GS.ColFormats[2]:='CB=RTTYPEPERSPECTIVE';
    GS.ColFormats[4]:='CB=GCREPRESENTANT';
    //GS.ColFormats[6]:='CB=RTPROJET';
    GS.ColFormats[7]:='CB=RTREPRESENTANT';
    //GS.ColFormats[4]:='CB=';
    GS.ColFormats[9]:='CB=RTETATPERSPECTIVE';
    GS.ColAligns[10]:=taRightJustify;GS.ColFormats[10]:='#,##0';
    GS.ColAligns[11]:=taRightJustify;GS.ColFormats[11]:='#,##0';
    AffecteGrid(GS,Action) ;
    TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;
end;

procedure TOF_PERSP_HISTO.OnLoad;
var Q : TQuery;
    Select,StCaption : String;
begin
inherited ;
TOBPersp:=tob.create('Historique de la perspective',Nil,-1) ;

Select := 'SELECT RPH_DATEMODIF,RPH_TYPEPERSPECTIV,RPH_PERSPECTIVE,RPH_REPRESENTANT,RPH_OPERATION,RPH_INTERVENANT,RPH_LIBELLE,C_NOM,RPH_ETATPER,RPH_MONTANTPER,RPH_POURCENTAGE,RPH_DATEREALISE,RPH_OPERATION,RPH_PROJET '+
       'FROM PERSPHISTO left join CONTACT on C_AUXILIAIRE = RPH_AUXILIAIRE and C_TYPECONTACT = "T"'+
       ' and C_NUMEROCONTACT = RPH_NUMEROCONTACT ' +
       'WHERE RPH_PERSPECTIVE="'+StPerspective+'" ORDER BY RPH_INDICE DESC';

Q:=OpenSQL(Select, True);
if Q.EOF then
  begin
  PGIBox(TexteMessage[1],TraduireMemoire('Historique de la proposition : ')+ StPerspective) ;
  PostMessage(TFVierge(Ecran).Handle,WM_CLOSE,0,0) ;
  Ferme(Q); TOBPersp.free; exit;
  end
else begin
  TOBPersp.LoadDetailDB('PERSPHISTO','','',Q,false,true);
  StCaption := IntToStr(TOBPersp.detail[0].GetValue ('RPH_PERSPECTIVE'))+' - '+TOBPersp.detail[0].GetValue ('RPH_LIBELLE');
  Ecran.Caption:=TraduireMemoire('Historique de la proposition : ')+ StCaption;
  TOBPersp.PutGridDetail(GS,False,False,LesColonnes,True);
end;
Ferme(Q);
TOBPersp.free;TOBPersp:=Nil;
end;

procedure TOF_PERSP_HISTO.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOF_PERSP_HISTO.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

{procedure TOF_PERSP_HISTO.GSLigneDClick (Sender: TObject);
var st_interv : string;
begin
  st_interv := GS.CellValues[Cell_Commercial, GS.Row];
  if (st_interv = '') then Exit;
  AGLLanceFiche('GC','GCCOMMERCIAL','',st_interv,'ACTION=MODIFICATION');
end;   }

procedure TOF_PERSP_HISTO.OnNew;
begin
inherited ;
end;

procedure TOF_PERSP_HISTO.DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
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
registerclasses([TOF_PERSP_HISTO]);
end.
