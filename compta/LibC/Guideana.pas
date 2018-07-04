unit GuideAna;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, ComCtrls, StdCtrls, ExtCtrls, Spin, Hctrls, Buttons, Mask, HSysMenu, HTB97,
{$IFDEF EAGLCLIENT}
  uTob,
{$ELSE}
  DBGrids, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  Hcompte,
  HEnt1,
  SaisUtil,
  hmsgbox,
  ent1,
  GuidUtil,
  SaisComm ;

Procedure ParamGuideAna ( AnaGui : TVentGuide ; Gene,TypeGuide : string ) ;

Const GCXSec=1 ; GCSec=2 ;
      GCXPou=3 ; GCPou=4 ;
      GCXQt1=5 ; GCQt1=6 ;
      GCXQt2=7 ; GCQt2=8 ;
type
  TFGuideAna = class(TForm)
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    PEntete: TPanel;
    PAxe: TTabControl;
    Messages: THMsgBox;
    Bevel2: TBevel;
    FListe: THGrid;
    HS: THCpteEdit;
    HMTrad: THSystemMenu;
    Valide97: TToolbar97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    PPied: TToolbar97;
    BInsLigne: TToolbarButton97;
    BDelLigne: TToolbarButton97;
    BZoom: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BInsLigneClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BDelLigneClick(Sender: TObject);
    procedure PAxeChange(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: string);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    FAxe,FGene,FTypeGuide : String ;
    FAnaGui : TVentGuide ;
    procedure ChargeGuideAna ;
    procedure SauveGuideAna ;
    Function CtrlGuideAna : boolean;
    Function CtrlSection(Lig:integer) : boolean;
    Function CtrlPourcent(Lig:integer) : boolean;
    Function CtrlVentil : boolean;
 public
  end;


implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  CPSECTION_TOM;

{$R *.DFM}

Procedure ParamGuideAna ( AnaGui : TVentGuide ; Gene,TypeGuide : string ) ;
var  FGuideAna: TFGuideAna;
BEGIN
FGuideAna:=TFGuideAna.Create(Application);
try
 FGuideAna.FAnaGui:=AnaGui ;
 FGuideAna.FGene:=Gene ;
 FGuideAna.FTypeGuide:=TypeGuide ;
 FGuideAna.ShowModal ;
 finally
 FGuideAna.Free ;
 END ;
Screen.Cursor:=crDefault ;
END ;

procedure TFGuideAna.FormShow(Sender: TObject);
var i,ia : integer;
begin
ia:=-1;
For i:=1 to 5 do
   BEGIN
   if OkVentil(FGene,IntToStr(i)) then BEGIN if ia<0 then ia:=i; END
                                  else BEGIN PAxe.Tabs.Strings[i-1]:='.....';END;
   END;
FAxe:='A'+IntToStr(ia); PAxe.TabIndex:=ia-1 ;
ChargeGuideAna ;
DelTabsSerie(PAxe) ; 
HS.ZoomTable:=TZoomTable(Ord(tzSection)+PAxe.TabIndex) ;
if ((EstSerie(S5)) or (EstSerie(S3))) then FListe.ColCount:=5 ;
end;

procedure TFGuideAna.ChargeGuideAna ;
Var i,ix : integer ;
    St : String ;
BEGIN
ix:=StrToInt(Copy(FAxe,2,1)) ;
FListe.VidePile(TRUE);
For i:=0 to FAnaGui.Ventil[ix].Count-1 do
   BEGIN
   St:=FAnaGui.Ventil[ix].Strings[i] ;

   FListe.Cells[0,i+1]:=IntToStr(i+1) ;

   FListe.Cells[GCSec,i+1]:=Trim(Copy(St,1,35)); Delete(St,1,35) ;
   FListe.Cells[GCXSec,i+1]:=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCPou,i+1]:=Trim(Copy(St,1,100)); Delete(St,1,100) ;
   FListe.Cells[GCXPou,i+1]:=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCQt1,i+1]:=Trim(Copy(St,1,100)); Delete(St,1,100) ;
   FListe.Cells[GCXQt1,i+1]:=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCQt2,i+1]:=Trim(Copy(St,1,100)); Delete(St,1,100) ;
   FListe.Cells[GCXQt2,i+1]:=St[1] ; Delete(St,1,1) ;

   FListe.RowCount:=FListe.RowCount+1 ;
   END ;
END ;

procedure TFGuideAna.SauveGuideAna ;
Var i,ix : integer ;
    St : String ;
begin
ix:=StrToInt(Copy(FAxe,2,1)) ;
FAnaGui.Ventil[ix].Clear ;
For i:=1 to FListe.RowCount-1 do if Not G_LigneVide(FListe,i) then
   BEGIN
   St:='' ;

   St:=St+Format_String(FListe.Cells[GCSec,i],35) ;
   if G_Croix(FListe.Cells[GCXSec,i]) then St:=St+'X' else St:=St+'-' ;

   St:=St+Format_String(FListe.Cells[GCPou,i],100) ;
   if G_Croix(FListe.Cells[GCXPou,i]) then St:=St+'X' else St:=St+'-' ;

   St:=St+Format_String(FListe.Cells[GCQt1,i],100) ;
   if G_Croix(FListe.Cells[GCXQt1,i]) then St:=St+'X' else St:=St+'-' ;

   St:=St+Format_String(FListe.Cells[GCQt2,i],100) ;
   if G_Croix(FListe.Cells[GCXQt2,i]) then St:=St+'X' else St:=St+'-' ;

   FAnaGui.Ventil[ix].Add(St) ;
   END ;
END ;

procedure TFGuideAna.BInsLigneClick(Sender: TObject);
begin
FListe.InsertRow(FListe.Row) ;
G_Renum(FListe) ;
FListe.SetFocus ;
end;

procedure TFGuideAna.BValiderClick(Sender: TObject);
begin
if Not CtrlGuideAna then Exit;
if Transactions(SauveGuideAna,5)<>oeOk then MessageAlerte(Messages.Mess[17]) ;
Close ;
end;

procedure TFGuideAna.BFermeClick(Sender: TObject);
begin
Close ;
end;

procedure TFGuideAna.BDelLigneClick(Sender: TObject);
begin
if ((Fliste.Row<=0) or (FListe.RowCount<=2)) then Exit ;
if FListe.Objects[0,FListe.Row]<>NIL then TVentGuide(FListe.Objects[0,FListe.Row]).Free ;
FListe.DeleteRow(FListe.Row) ;
G_Renum(FListe) ;
FListe.SetFocus ;
end;


procedure TFGuideAna.PAxeChange(Sender: TObject);
Var OldAxe : Integer ;
begin
OldAxe:=StrToInt(Copy(FAxe,2,1)) ;
if OkVentil(FGene,IntToStr(PAxe.TabIndex+1)) then
   BEGIN
   SauveGuideAna ;
   FAxe:='A'+IntToStr(PAxe.TabIndex+1) ;
   HS.ZoomTable:=TZoomTable(Ord(tzSection)+PAxe.TabIndex) ;
   ChargeGuideAna ;
   END else
   BEGIN
   Messages.Execute(0,FGene+' : ','') ;
   PAxe.TabIndex:=OldAxe-1 ;
   END ;
G_Renum(FListe) ;
end;


Function TFGuideAna.CtrlGuideAna : boolean;
var i    : integer;
    Good : boolean;
    lErr : integer ;
BEGIN
CtrlGuideAna:=False; lErr:=-1 ;
For i:=1 to FListe.RowCount-1 do
    BEGIN
    if Not G_LigneVide(FListe,i) then
       BEGIN
       Good:=CtrlSection(i) ;
       if Good then Good:=CtrlPourcent(i);
       if Good then lErr:=CtrlFormule(FListe,i,TRUE) ;
       Good := lErr = - 1 ;
       if Not Good then BEGIN FListe.Row:=i ; Messages.Execute(lErr,'','') ; Exit; END ;
       END else
       BEGIN
       if i<>FListe.RowCount-1 then
          BEGIN
          FListe.Row:=i ;
          Messages.Execute(14,'',''); Exit;
          END;
       END;
    END;
Good:=CtrlVentil;
if Not Good then exit;
CtrlGuideAna:=True;
END;

Function TFGuideAna.CtrlVentil : boolean;
var i    : integer;
    P    : array[1..3] of double ;
BEGIN
CtrlVentil:=false;
for i:=1 to 3 do P[i]:=0;
for i:=1 to FListe.RowCount-1 do
   BEGIN
   P[1]:=P[1]+Valeur(FListe.Cells[GCPou,i]) ;
   P[2]:=P[2]+Valeur(FListe.Cells[GCQt1,i]) ;
   P[3]:=P[3]+Valeur(FListe.Cells[GCQt2,i]) ;
   END;
if FTypeGuide<>'NOR' then
   BEGIN
   if Arrondi(P[1]-100.0,ADecimP)<>0 then BEGIN Messages.Execute(16,'','') ; exit ; END ;
   END ;
CtrlVentil:=true;
END;

Function TFGuideAna.CtrlPourcent(Lig:integer) : boolean;
var i,j  : integer;
    PC   : double ;
    St   : String;
BEGIN
CtrlPourcent:=False;
For i:=1 to 3 do
   BEGIN
   j:=GCPou ;
   case i of 1:  j:=GCPou ; 2:  j:=GCQt1 ; 3:  j:=GCQt2; end;
   St:= FListe.Cells[j,Lig] ;
   if St<>'' then
      if IsMontant(St,False) then
         BEGIN
         PC:=Valeur(St);
         if ((PC<0) or (PC>100)) then
            BEGIN FListe.Col:=j ; FListe.Row:=lig ; FListe.SetFocus; Messages.Execute(4,'','');exit; END;
         END else
         BEGIN
         FListe.Col:=j ; FListe.Row:=lig ; FListe.SetFocus;
         Messages.Execute(4,'','');exit;
         END;
   END;
CtrlPourcent:=True;
END;


Function TFGuideAna.CtrlSection(Lig:integer) : boolean;
var SExist,SClose : boolean;
BEGIN
CtrlSection:=False;
HS.Text:=FListe.Cells[GCSec,Lig] ;
SExist:=(HS.ExisteH>0) ;
SExist:=((SExist) and (HS.Text=FListe.Cells[GCSec,Lig])) ; 
SClose:=IsClose(HS.Text,'SECTION');
if Not SExist then BEGIN FListe.Col:=GCSec ; FListe.Row:=Lig ; FListe.Setfocus; Messages.Execute(5,'',''); Exit; END;
if (SExist and SClose) then BEGIN FListe.Col:=GCSec ; FListe.Row:=Lig ; FListe.Setfocus; Messages.Execute(1,'',''); Exit; END;
CtrlSection:=True;
END;


procedure TFGuideAna.BZoomClick(Sender: TObject);
begin
HS.text:=FListe.Cells[GCSec,FListe.Row];
if (HS.ExisteH<=0) then exit ;
FicheSection(Nil,FAxe,HS.Text,taConsult,0) ;
FListe.SetFocus ;
end;

procedure TFGuideAna.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TFGuideAna.FListeDblClick(Sender: TObject);
begin
Case FListe.Col of
   GCSec : BEGIN
           HS.Text:=uppercase(FListe.Cells[GCSec,FListe.Row]) ;
           if GChercheCompte(HS,Nil) then FListe.Cells[GCSec,FListe.Row]:=HS.Text ;
           END ;
   END ;
end;

procedure TFGuideAna.FListeSetEditText(Sender: TObject; ACol,ARow: Longint; const Value: string);
begin
if ((FListe.Row=FListe.RowCount-1) and (Not G_LigneVide(FListe,FListe.Row))) then
   BEGIN
   FListe.RowCount:=FListe.RowCount+1 ;
   END;
if ((FListe.Row=FListe.RowCount-3) and (G_LigneVide(FListe,FListe.Rowcount-1))
                                   and (G_LigneVide(FListe,FListe.Rowcount-2))) then
   BEGIN
   if FListe.Objects[0,FListe.Rowcount-1]<>NIL then TVentGuide(FListe.Objects[0,FListe.Rowcount-1]).Free ;
   FListe.RowCount:=FListe.RowCount-1 ;
   END;
G_Renum(FListe) ;
end;

procedure TFGuideAna.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : boolean ;
begin
Vide:=(Shift=[]) ; 
Case Key of
  VK_RETURN : if ((ActiveControl=Fliste) and (Vide)) then KEY:=VK_TAB ;
     VK_F10 : if Vide then BEGIN Key:=0 ; BValiderClick(Nil) ; END ;
  END ;
end;

procedure TFGuideAna.FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin
if ((Acol=GCSec) and (FListe.Cells[ACol,ARow]<>'')) then FListe.Cells[ACol,ARow]:=uppercase(FListe.Cells[ACol,ARow]) ; 
end;

procedure TFGuideAna.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ; 
end;

procedure TFGuideAna.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(TRUE);
end;

end.
