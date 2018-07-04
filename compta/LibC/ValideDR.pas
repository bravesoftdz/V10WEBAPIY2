{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 15/04/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit ValideDR;

interface

uses
  Forms, Classes, Controls, StdCtrls, Buttons, ExtCtrls, Grids, Hctrls,
  SysUtils,  // StrToInt, IntToStr
  Graphics,  // TCanvas, clRed
  LettUtil,  // FormatCheque
  HSysMenu,  // THSystemMenu
  HMsgBox,   // THMsgBox
  NumChek    // SaisieNumCheque
  ;

Const colTiers = 0 ;
      colLibelle = 1 ;
      colMontant = 2 ;
      colNumCheque = 3 ;

function ValideImpDR(G : THGrid ; LaColMontant : Integer) : Boolean ;

type
  TFValideDR = class(TForm)
    G: THGrid;
    pBoutons: TPanel;
    bAnnuler: THBitBtn;
    BAide: THBitBtn;
    BValider: THBitBtn;
    bNumerote: THBitBtn;
    Msg: THMsgBox;
    HMTrad: THSystemMenu;
    bDetruit: THBitBtn;
    Panel1: TPanel;
    HLabel1: THLabel;
    procedure FormShow(Sender: TObject);
    procedure bDetruitClick(Sender: TObject);
    procedure bNumeroteClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    function  ChequeEdite(lig : Integer) : Boolean ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

function ValideImpDR(G : THGrid ; LaColMontant : Integer) : Boolean ;
var
  i,j : integer ;
  X   : TFValideDR ;
begin
result:=False ;
X:=TFValideDR.Create(Application) ;
X.G.RowCount:=G.RowCount-1 ;
for i:=1 to G.RowCount-2 do
   begin
   for j:=0 to colNumCheque do
     BEGIN
     If j=ColMontant Then X.G.Cells[j,i]:=G.Cells[LaColMontant,i]
                       Else X.G.Cells[j,i]:=G.Cells[j,i] ;
     END ;
   end ;
try
   X.ShowModal ;
   if X.ModalResult=mrOK then
      begin
      for i:=1 to X.G.RowCount-1 do
         begin
         if X.ChequeEdite(i) then G.Cells[colNumCheque,i]:=X.G.Cells[colNumCheque,i]
                             else G.Cells[colNumCheque,i]:='' ;
         end ;
      result:=True ;
      end ;
   finally
   X.Free ;
   end ;
end ;

function TFValideDR.ChequeEdite(lig : Integer) : Boolean ;
begin
result:=not (G.Cells[colNumCheque,lig]=Msg.Mess[0]) ;
end ;

{================ Evènements ================}

procedure TFValideDR.FormShow(Sender: TObject);
begin
G.GetCellCanvas:=GetCellCanvas ;
end;

procedure TFValideDR.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
if (ARow=0) then exit ;
if (ACol=colNumCheque) then
   begin
   if not ChequeEdite(ARow) then G.Canvas.Font.Color:=clRed ;
   end ;
end ;

procedure TFValideDR.bDetruitClick(Sender: TObject);
begin
  if Msg.Execute(1,caption,'')<>mrYes then Exit;  // Marquer le chèque sélectionné comme étant détruit ? Attention cette opération est définitive.
  G.Cells[colNumCheque,G.Row] := Msg.Mess[0]; // Détruit
G.Invalidate ;
end;

procedure TFValideDR.bNumeroteClick(Sender: TObject);
var
  i,noChq : integer ;
  sRef,snChq,sTransit : string ;
begin
sRef:='' ; snChq:='0' ;
if SaisieNumCheque(sRef,snChq,True) then
   begin
   for i:=G.Row to G.RowCount-1 do
      begin
      if ChequeEdite(i) then
         begin
         G.Cells[colNumCheque,i]:=FormatCheque(sRef,snChq) ;
         noChq:=StrToInt(Trim(snChq)) ; Inc(noChq) ;
         sTransit:=IntToStr(noChq) ;
         While Length(sTransit)<Length(snChq) do sTransit:='0'+sTransit ;
         snChq:=sTransit ; 
         end ;
      end ;
   end ;
G.Invalidate ;
end;

procedure TFValideDR.FormCreate(Sender: TObject);
begin
  PopupMenu := AddMenuPop(PopupMenu,'','');
end;

procedure TFValideDR.BAideClick(Sender: TObject);
begin
  CallHelptopic(Self); 
end;

end.
