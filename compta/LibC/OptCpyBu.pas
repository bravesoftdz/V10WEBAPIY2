{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit OptCpyBu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  Hent1, StdCtrls, Hctrls, hmsgbox, Spin, Grids, Buttons, ExtCtrls,
  HSysMenu;

Procedure OptiondeCopideBudget(Cod,Lib,Nat : String ; Var NatCpy : String ; Var Coefficient : Double) ;

type
  TFOptCpyBu = class(TForm)
    HPB: TPanel;
    Nb1: TLabel;
    Tex1: TLabel;
    Panel1: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
    FListe: THGrid;
    PTop: TPanel;
    TSeCoef: TLabel;
    HM: THMsgBox;
    NatBud: THValComboBox;
    SeCoef: TEdit;
    Label1: TLabel;
    HMTrad: THSystemMenu;
    procedure BFermeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    Cod,Lib,Nat,NatCpy : String ;
    Coefficient : Double ;
    WMinX,WMinY : Integer ;
    TotalSelec : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure CompteElemSelectionner ;
    Procedure RempliFListe ;
    Function  ControleCoefOk : Boolean ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Procedure OptiondeCopideBudget(Cod,Lib,Nat : String ; Var NatCpy : String ; Var Coefficient : Double) ;
var FOptCpyBu : TFOptCpyBu ;
BEGIN
FOptCpyBu:=TFOptCpyBu.Create(Application) ;
  Try
   FOptCpyBu.Cod:=Cod ;
   FOptCpyBu.Lib:=Lib ;
   FOptCpyBu.Nat:=Nat ;
   FOptCpyBu.NatCpy:=NatCpy ;
   FOptCpyBu.Coefficient:=Coefficient ;
   FOptCpyBu.ShowModal ;
   NatCpy:=FOptCpyBu.NatCpy ;
   Coefficient:=FOptCpyBu.Coefficient ;
  Finally
   FOptCpyBu.Free ;
  End ;
SourisNormale ;
END ;

procedure TFOptCpyBu.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFOptCpyBu.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFOptCpyBu.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; end;

Procedure TFOptCpyBu.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFOptCpyBu.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if ((Shift=[]) And (Key=VK_SPACE)) or ((ssShift in Shift) And (Key=VK_DOWN)) or
   ((ssShift in Shift) And (Key=VK_UP))then FlisteMouseDown(Nil,mbLeft,[ssCtrl],0,0) ;
if (Shift=[]) And (Key=VK_SPACE) then
    BEGIN
    FlisteMouseDown(Nil,mbLeft,[ssCtrl],0,0) ;
    if ((FListe.Row<FListe.RowCount-1) and (Key<>VK_SPACE)) then FListe.Row:=FListe.Row+1 ;
    END ;
end;

procedure TFOptCpyBu.FListeMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if FListe.Cells[0,1]='' then Exit ;
if Not(ssCtrl in Shift) then Exit ;
if Button<>mbLeft then Exit ;
if Fliste.Cells[FListe.ColCount-1,FListe.Row]='*'
   then Fliste.Cells[FListe.ColCount-1,FListe.Row]:=''
   else Fliste.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
FListe.Invalidate ; CompteElemSelectionner ;
end;

Procedure TFOptCpyBu.CompteElemSelectionner ;
Var i : Integer ;
BEGIN
TotalSelec:=0 ;
for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(TotalSelec) ;
Nb1.Caption:=IntToStr(TotalSelec) ;
if TotalSelec>1 then Tex1.Caption:=HM.Mess[1] else Tex1.Caption:=HM.Mess[0] ;
END ;

Procedure TFOptCpyBu.RempliFListe ;
Var i : integer ;
    St,St1 : String ;
BEGIN
St:=Nat ; i:=2 ;
While St<>'' do
  BEGIN
  FListe.RowCount:=i ;
  FListe.Cells[0,FListe.RowCount-1]:=Cod ;
  FListe.Cells[1,FListe.RowCount-1]:=Lib ;
  St1:=ReadTokenSt(St) ;
  FListe.Cells[2,FListe.RowCount-1]:=St1 ;
  FListe.Cells[3,FListe.RowCount-1]:=NatBud.Items[NatBud.Values.IndexOf(St1)] ;
  if Pos(St1,NatCpy)>0 then FListe.Cells[FListe.ColCount-1,FListe.RowCount-1]:='*' ;
  Inc(i) ;
  END ;
FListe.Invalidate ; CompteElemSelectionner ;
END ;

procedure TFOptCpyBu.FormShow(Sender: TObject);
begin
FListe.GetCellCanvas:=GetCellCanvas ;
RempliFListe ; SeCoef.Text:=StrFPoint(Coefficient) ;
end;

procedure TFOptCpyBu.BValiderClick(Sender: TObject);
Var i : Integer ;
begin
if (TotalSelec=0) And (FListe.Cells[0,1]<>'') then BEGIN HM.Execute(2,'','') ; Exit ; END ;
NatCpy:='' ;
for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then NatCpy:=NatCpy+FListe.Cells[2,i]+';' ;
if Not ControleCoefOk then Exit ;
Coefficient:=Valeur(SeCoef.Text) ; Close ;
end;

Function TFOptCpyBu.ControleCoefOk : Boolean ;
BEGIN
Result:=False ;
if Not IsNumeric(SeCoef.Text) then BEGIN HM.Execute(3,'','') ; SeCoef.SetFocus ; Exit ; END ;
Result:=True ;
END ;

procedure TFOptCpyBu.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
