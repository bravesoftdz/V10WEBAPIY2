{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit SaisBase;

interface

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls, Ent1, HEnt1, Echeance, SaisUtil,
  hmsgbox, HSysMenu, SaisComm, HTB97 ;

Type Enr_Base = RECORD
                Regime,Nature : String3 ;
                Client        : Boolean ;
                Action        : TActionFiche ;
                CodeTva       : String3 ;
                END ;

Function SaisieBasesHT ( OMODR : TMOD ; Var RTVA : Enr_Base ) : Boolean ;

type
  TFSaisBase = class(TForm)
    POutils: TPanel;
    BValide: THBitBtn;
    BAbandon: THBitBtn;
    BAide: THBitBtn;
    G: THGrid;
    GEche: THGrid;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    TotHT: THNumEdit;
    Label1: TLabel;
    Label2: TLabel;
    TotTTC: THNumEdit;
    BCalcHT: TToolbarButton97;
    H_TVA: THLabel;
    E_TVA: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure GEcheRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure BAbandonClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BCalcHTClick(Sender: TObject);
    procedure GExit(Sender: TObject);
  private
    XX : Array[1..MaxEche,1..5] of Double ;
    StCur                       : String ;
    procedure RempliBasesEches ;
    procedure MontreBase ( NumEche : integer ) ;
    procedure SommeHT ( Lig : integer ) ;
  public
    OM   : TMOD ;
    RTVA : Enr_Base ;
    BaseModifie : Boolean ;
  end;


implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  utilPGI;

{$R *.DFM}

Function SaisieBasesHT ( OMODR : TMOD ; Var RTVA : Enr_Base ) : Boolean ;
Var X : TFSaisBase ;
BEGIN
X:=TFSaisBase.Create(Application) ;
 Try
  X.OM:=OMODR ; X.RTVA:=RTVA ;
  X.ShowModal ;
  Result:=X.BaseModifie ;
  if Result then RTVA.CodeTva:=X.RTVA.CodeTva ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFSaisBase.SommeHT ( Lig : integer ) ;
Var Tot,TotL : Double ;
    i,k : integer ;
BEGIN
Tot:=0 ; TotL:=0 ;
for i:=1 to GEche.RowCount-1 do for k:=1 to 5 do
    BEGIN
    Tot:=Tot+XX[i,k] ;
    if i=Lig then TotL:=TotL+XX[i,k] ;
    END ;
TotHT.Value:=Tot ;
GEche.Cells[4,Lig]:=StrS0(TotL) ;
END ;

procedure TFSaisBase.RempliBasesEches ;
Var i,k : integer ;
    TotL : Double ;
BEGIN
GEche.RowCount:=OM.ModR.NbEche+1 ;
for i:=1 to OM.ModR.NbEche do
    BEGIN
    GEche.Cells[0,i]:=IntToStr(i) ;
    GEche.Cells[1,i]:=DateToStr(OM.ModR.TabEche[i].DateEche) ;
    GEche.Cells[2,i]:=RechDom('ttModePaie',OM.ModR.TabEche[i].ModePaie,False) ;
    GEche.Cells[3,i]:=StrS0(OM.ModR.TabEche[i].MontantP) ;
    TotL:=0 ;
    for k:=1 to 5 do BEGIN XX[i,k]:=OM.ModR.TabEche[i].TAV[k] ; TotL:=TotL+XX[i,k] ; END ;
    GEche.Cells[4,i]:=StrS0(TotL) ;
    END ;
for i:=1 to 5 do
    BEGIN
    G.Cells[0,i]:=HM.Mess[i-1] ;
    if i<5 then
       BEGIN
       G.Cells[1,i]:=VH^.NumCodeBase[i] ;
       G.Cells[2,i]:=StrS0(100.0*Tva2Taux(RTVA.Regime,VH^.NumCodeBase[i],Not RTVA.Client)) ;
       END;

END ;
MontreBase(1) ;
SommeHT(1) ;
END ;



procedure TFSaisBase.MontreBase ( NumEche : integer ) ;
Var i : integer ;
BEGIN
for i:=1 to 5 do G.Cells[3,i]:=StrS0(XX[NumEche,i]) ;
END ;

procedure TFSaisBase.FormShow(Sender: TObject);
begin
ChangeMask(TotHT,V_PGI.OkDecV,'') ; ChangeMask(TotTTC,V_PGI.OkDecV,'') ;
if VH^.PaysLocalisation=CodeISOES then
   Begin
   G.ColWidths[0]:=-1 ;
   HMTrad.ResizeGridColumns(G) ;
   End ; //XVI 24/02/2005
G.ColAligns[1]:=taCenter ; G.ColAligns[2]:=taRightJustify ; G.ColAligns[3]:=taRightJustify ;
GEche.ColAligns[0]:=taCenter ; GEche.ColAligns[1]:=taCenter ; GEche.ColAligns[2]:=taCenter ;
GEche.ColAligns[3]:=taRightJustify ; GEche.ColAligns[4]:=taRightJustify ;
if RTVA.Action=taConsult then BEGIN AffecteGrid(G,RTVA.Action) ; BCalcHT.Visible:=False ; END ;
RempliBasesEches ;
G.Col:=3 ; G.Row:=1 ;
TotTTC.Value:=OM.ModR.TotalAPayerP ;
BaseModifie:=FALSE ;
if ((RTVA.Nature='OC') or (RTVA.Nature='OF')) then
   BEGIN
   H_TVA.Visible:=True ; E_TVA.Visible:=True ;
   E_TVA.Value:=RTVA.CodeTva ;
   END ;
if Valeur(G.Cells[3,5])=0 then BEGIN E_TVA.Value:='' ; E_TVA.Enabled:=False ; END ;
end;

procedure TFSaisBase.GEcheRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
MontreBase(GEche.Row) ;
end;

procedure TFSaisBase.GCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin
if RTVA.Action=taConsult then Exit ;
if G.Col<>3 then BEGIN ACol:=3 ; Cancel:=True ; END else
 if ((G.Row<5) and (VH^.NumCodeBase[G.Row]='')) then BEGIN ACol:=3 ; ARow:=5 ; Cancel:=True ; END else
   StCur:=G.Cells[G.Col,G.Row] ;
end;

procedure TFSaisBase.GCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin
if RTVA.Action=taConsult then Exit ;
if G.Cells[ACol,ARow]=StCur then Exit ;
G.Cells[ACol,ARow]:=StrS0(Valeur(G.Cells[ACol,ARow])) ;
XX[GEche.Row,ARow]:=Valeur(G.Cells[ACol,ARow]) ;
SommeHT(GEche.Row) ;
if ((ACol=3) and (ARow=5)) then E_TVA.Enabled:=(Valeur(G.Cells[3,5])<>0) ;
if Not E_TVA.Enabled then E_TVA.Value:='' ;
end;

procedure TFSaisBase.BAbandonClick(Sender: TObject);
begin
Close ;
end;

procedure TFSaisBase.BValideClick(Sender: TObject);
Var i,k : integer ;
begin
if RTVA.Action<>taConsult then
   BEGIN
   if Abs(TotHT.Value)>Abs(TotTTC.Value) then if HM.Execute(5,'','')<>mrYes then Exit ;
   for i:=1 to GEche.RowCount-1 do
      for k:=1 to 5 do
         BEGIN
         if OM.ModR.TabEche[i].TAV[k]<>XX[i,k] then BaseModifie:=TRUE ;
         OM.ModR.TabEche[i].TAV[k]:=XX[i,k] ;
         END ;
   OM.ModR.ModifTva:=True ;
   RTVA.CodeTva:=E_TVA.Value ;
   END ;
Close ;
end;

procedure TFSaisBase.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFSaisBase.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; StCur:='azeaze' ; 
end;

procedure TFSaisBase.BCalcHTClick(Sender: TObject);
Var
  Taux, Coef, Total : double ;
  i : integer ;
begin
if RTVA.Action=taConsult then Exit ;

Coef:=1.0 ; if ((RTVA.Nature='AF') or (RTVA.Nature='AC')) then Coef:=-1.0 ;

Total:=0.0;
{14/09/2007 YMO Calcul prenant en compte les valeurs des autres régimes}
For i:=1 to G.RowCount-1 do
begin
  if G.Cells[2,i]<>'' then
    Taux:=Valeur(G.Cells[2,i]) ;

  If i=G.RowCount-1 then {On prend par convention le taux 1 pour la ligne DEBIT}
    Taux:=100.0*Tva2Taux(RTVA.Regime,VH^.NumCodeBase[1],Not RTVA.Client) ;

  If (i<>G.Row) and (Taux<>0) then
      Total:=Total+Valeur(G.Cells[3,i])*(1.0+(Taux/100.0))
end;

if G.Cells[2,G.Row]<>'' then
  Taux:=Valeur(G.Cells[2,G.Row]) ;

If G.Row=G.RowCount-1 then {On prend par convention le taux 1 pour la ligne DEBIT}
  Taux:=100.0*Tva2Taux(RTVA.Regime,VH^.NumCodeBase[1],Not RTVA.Client) ;

G.Cells[3,G.Row]:=StrS0((Coef*Valeur(GEche.Cells[3,GEche.Row])-Total)/(1.0+(Taux/100.0))) ;
XX[GEche.Row,G.Row]:=Valeur(G.Cells[3,G.Row]) ;
SommeHT(GEche.Row) ;

end;


procedure TFSaisBase.GExit(Sender: TObject);
Var Acol,Arow : Longint ;
    Cancel : boolean ;
begin
Cancel:=False ; ACol:=G.Col; ARow:=G.Row ;
GCellExit(Nil,ACol,ARow,Cancel) ;
end;

end.
