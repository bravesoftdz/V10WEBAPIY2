unit SaisEcar;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Grids, Hctrls, HSysMenu, hmsgbox,
  SaisComm, SaisUtil, Ent1, HEnt1, LettUtil, UTob;

Function CorrigeEcartConvert ( DEV : RDEVISE ; GS : THGrid ; Action : TActionFiche ) : boolean ;
function TOBCorrigeEcartConvert(DEV : RDEVISE; Ecrs : TOB; Action : TActionFiche) : Boolean ;

type
  TFSaisEcar = class(TForm)
    POutils: TPanel;
    BValide: THBitBtn;
    BAbandon: THBitBtn;
    BAide: THBitBtn;
    PEntete: TPanel;
    PFen: TPanel;
    G: THGrid;
    LPivot: TLabel;
    LOppose: TLabel;
    HMTrad: THSystemMenu;
    HDiv: THMsgBox;
    HOppose: THLabel;
    SA_SOLDEOppose: THNumEdit;
    Bevel4: TBevel;
    HPivot: THLabel;
    SA_SoldePivot: THNumEdit;
    Bevel1: TBevel;
    Label1: TLabel;
    BSolde: THBitBtn;
    BModif: THBitBtn;
    ISigneEuro: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure GCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BModifClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BAideClick(Sender: TObject);
  private
    StCellCur : String ;
    DEV       : RDEVISE ;
    ModeModif,EccPivot,OkModif : Boolean ;
    procedure AfficheTitres ;
    procedure ChargeGrid ;
    procedure TOBChargeGrid ;
    procedure CalculSoldes ;
    procedure GVersO ;
    procedure TOBGVersO ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure FormateMontant ( ACol,ARow : integer ) ;
    Function  JeValide : boolean ;
    Function  ConfirmeEcart : boolean ;
  public
    GS : THGrid ;
    Ecrs : TOB ;
    Action : TActionFiche ;
    OKVal : boolean ;
  end;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  UtilPGI;


Function CorrigeEcartConvert ( DEV : RDEVISE ; GS : THGrid ; Action : TActionFiche ) : boolean ;
Var X : TFSaisEcar ;
BEGIN
X:=TFSaisEcar.Create(Application) ;
 Try
  X.GS:=GS ; X.OkVal:=False ; X.DEV:=DEV ; X.Action:=Action ;
  X.ShowModal ; X.Ecrs:=nil ;
  Result:=X.OkVal ;
 Finally
  X.Free ;
 End ;
END ;

function TOBCorrigeEcartConvert(DEV : RDEVISE; Ecrs : TOB; Action : TActionFiche) : Boolean ;
var X : TFSaisEcar ;
BEGIN
X:=TFSaisEcar.Create(Application) ;
 try
  X.Ecrs:=Ecrs ; X.OkVal:=FALSE ; X.DEV:=DEV ; X.Action:=Action ;
  X.ShowModal ; X.GS:=nil ;
  Result:=X.OkVal ;
 Finally
  X.Free ;
 end ;
end ;

procedure TFSaisEcar.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
G.GetCellCanvas:=GetCellCanvas ;
StCellCur:='' ; ModeModif:=False ; EccPivot:=True ; OkModif:=True ;
end;

procedure TFSaisEcar.FormClose(Sender: TObject; var Action: TCloseAction);
begin
G.VidePile(False) ;
end;

procedure TFSaisEcar.AfficheTitres ;
Var k : integer ;
BEGIN
LPivot.Caption:=RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False) ;
HPivot.Caption:=HPivot.Caption+' '+RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False) ;
if Not VH^.TenueEuro then
   BEGIN
   LOppose.Caption:=HDiv.Mess[0] ;
   HOppose.Caption:=HOppose.Caption+' '+HDiv.Mess[0] ;
   END else
   BEGIN
   LOppose.Caption:=RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False) ;
   HOppose.Caption:=HOppose.Caption+' '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False) ;
   END ;
G.ColAligns[0]:=taCenter ; G.ColAligns[1]:=taLeftJustify ; G.ColAligns[2]:=taLeftJustify ;
for k:=3 to 8 do G.ColAligns[k]:=taRightJustify ;
END ;

procedure TFSaisEcar.CalculSoldes ;
Var i : integer ;
    TotDE,TotCE,TotDP,TotCP : Double ;
BEGIN
TotDP:=0 ; TotCP:=0 ; TotDE:=0 ; TotCE:=0 ;
for i:=1 to G.RowCount-1 do
    BEGIN
    TotDP:=TotDP+Valeur(G.Cells[3,i]) ; TotCP:=TotCP+Valeur(G.Cells[4,i]) ;
    TotDE:=TotDE+Valeur(G.Cells[6,i]) ; TotCE:=TotCE+Valeur(G.Cells[7,i]) ;
    G.Cells[5,i]:=StrS(TotDP-TotCP,V_PGI.OkDecV) ;
    G.Cells[8,i]:=StrS(TotDE-TotCE,V_PGI.OkDecE) ;
    END ;
AfficheLeSolde(SA_SoldePivot,TotDP,TotCP) ; AfficheLeSolde(SA_SoldeOppose,TotDE,TotCE) ;
END ;

procedure TFSaisEcar.ChargeGrid ;
Var i : integer ;
    O : TOBM ;
    Okok : boolean ;
BEGIN
G.VidePile(False) ; Okok:=False ;
for i:=1 to GS.RowCount-2 do
    BEGIN
    O:=GetO(GS,i) ; if O=Nil then Break ;
    G.Cells[0,G.RowCount-1]:=IntToStr(i) ; Okok:=True ;
    G.Cells[1,G.RowCount-1]:=O.GetMvt('E_GENERAL') ;
    G.Cells[2,G.RowCount-1]:=O.GetMvt('E_AUXILIAIRE') ;
    G.Cells[3,G.RowCount-1]:=StrF00(O.GetMvt('E_DEBIT'),V_PGI.OkDecV)     ;
    G.Cells[4,G.RowCount-1]:=StrF00(O.GetMvt('E_CREDIT'),V_PGI.OkDecV) ;
//    G.Cells[6,G.RowCount-1]:=StrF00(O.GetMvt('E_DEBITEURO'),V_PGI.OkDecE) ;
//    G.Cells[7,G.RowCount-1]:=StrF00(O.GetMvt('E_CREDITEURO'),V_PGI.OkDecE) ;
    G.RowCount:=G.RowCount+1 ;
    END ;
if Okok then G.RowCount:=G.RowCount-1 ;
END ;

procedure TFSaisEcar.TOBChargeGrid ;
var Ecr : TOB ; i : Integer ; bOK : Boolean ;
begin
G.VidePile(False) ; bOK:=FALSE ;
for i:=0 to Ecrs.Detail.Count-1 do
    begin
    Ecr:=Ecrs.Detail[i] ; if Ecr=nil then Break ;
    G.Cells[0,G.RowCount-1]:=IntToStr(i+1) ; bOK:=TRUE ;
    G.Cells[1,G.RowCount-1]:=Ecr.GetValue('E_GENERAL') ;
    G.Cells[2,G.RowCount-1]:=Ecr.GetValue('E_AUXILIAIRE') ;
    G.Cells[3,G.RowCount-1]:=StrF00(Ecr.GetValue('E_DEBIT'),V_PGI.OkDecV) ;
    G.Cells[4,G.RowCount-1]:=StrF00(Ecr.GetValue('E_CREDIT'),V_PGI.OkDecV) ;
//    G.Cells[6,G.RowCount-1]:=StrF00(Ecr.GetValue('E_DEBITEURO'),V_PGI.OkDecE) ;
//    G.Cells[7,G.RowCount-1]:=StrF00(Ecr.GetValue('E_CREDITEURO'),V_PGI.OkDecE) ;
    G.RowCount:=G.RowCount+1 ;
    end ;
if bOK then G.RowCount:=G.RowCount-1 ;
END ;

procedure TFSaisEcar.FormShow(Sender: TObject);
begin
AfficheTitres ;
ChangeMask(SA_SoldePivot,V_PGI.OkDecV,'') ; ChangeMask(SA_SoldeOppose,V_PGI.OkDecE,'') ;
if GS<>nil then ChargeGrid else TOBChargeGrid ;
CalculSoldes ;
EccPivot:=(SA_SoldePivot.Value<>0) ;
if Action=taConsult then
   BEGIN
   BSolde.Visible:=False ; BModif.Visible:=False ;
   Caption:=HDiv.Mess[4]+' '+RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False) ;
   if VH^.TenueEuro then Caption:=Caption+' + '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False)
                    else Caption:=Caption+' + '+HDiv.Mess[0] ;
   if ((Action=taConsult) and (DEV.Code<>V_PGI.DevisePivot) and (DEV.Code<>V_PGI.DeviseFongible)) then Caption:=Caption+' '+HDiv.Mess[6] ;
   HelpContext:=7244700 ;                  
   END ;
end;

procedure TFSaisEcar.BAbandonClick(Sender: TObject);
begin
Close ;
end;

procedure TFSaisEcar.BSoldeClick(Sender: TObject);
Var XP,XX : Double ;
    Lig   : integer ;
begin
if Action=taConsult then Exit ;
BModif.Enabled:=False ;
Lig:=G.Row ;
XX:=SA_SoldePivot.Value ; if Not SA_SoldePivot.Debit then XX:=-XX ;
if XX<>0 then
   BEGIN
   XP:=Valeur(G.Cells[3,Lig]) ;
   if XP<>0 then BEGIN XP:=XP-XX ; G.Cells[3,Lig]:=StrF00(XP,V_PGI.OkDecV) ; END
            else BEGIN XP:=Valeur(G.Cells[4,Lig]) ; XP:=XP+XX ; G.Cells[4,Lig]:=StrF00(XP,V_PGI.OkDecV) ; END ;
   END ;
XX:=SA_SoldeOppose.Value ; if Not SA_SoldeOppose.Debit then XX:=-XX ;
if XX<>0 then
   BEGIN
   XP:=Valeur(G.Cells[6,Lig]) ;
   if XP<>0 then BEGIN XP:=XP-XX ; G.Cells[6,Lig]:=StrF00(XP,V_PGI.OkDecE) ; END
            else BEGIN XP:=Valeur(G.Cells[7,Lig]) ; XP:=XP+XX ; G.Cells[7,Lig]:=StrF00(XP,V_PGI.OkDecE) ; END ;
   END ;
CalculSoldes ;
end;

procedure TFSaisEcar.GVersO ;
Var i : integer ;
    O : TOBM ;
    X : Double ;
BEGIN
if Action=taConsult then Exit ;
for i:=1 to G.RowCount-1 do
    BEGIN
    O:=GetO(GS,i) ; if O=Nil then Break ;
    X:=Valeur(G.Cells[3,i]) ; O.PutMvt('E_DEBIT',X) ;
    X:=Valeur(G.Cells[4,i]) ; O.PutMvt('E_CREDIT',X) ;
//    X:=Valeur(G.Cells[6,i]) ; O.PutMvt('E_DEBITEURO',X) ;
//    X:=Valeur(G.Cells[7,i]) ; O.PutMvt('E_CREDITEURO',X) ;
    if DEV.Code=V_PGI.DevisePivot then
       BEGIN
       O.PutMvt('E_DEBITDEV',O.GetMvt('E_DEBIT')) ;
       O.PutMvt('E_CREDITDEV',O.GetMvt('E_CREDIT')) ;
       END ;
    END ;
END ;

procedure TFSaisEcar.TOBGVersO ;
var Ecr : TOB ; i : Integer ; X : Double ;
begin
if Action=taConsult then Exit ;
for i:=1 to G.RowCount-1 do
    begin
    Ecr:=Ecrs.Detail[i-1] ; if Ecr=nil then Break ;
    X:=Valeur(G.Cells[3,i]) ; Ecr.PutValue('E_DEBIT', X) ;
    X:=Valeur(G.Cells[4,i]) ; Ecr.PutValue('E_CREDIT', X) ;
//    X:=Valeur(G.Cells[6,i]) ; Ecr.PutValue('E_DEBITEURO', X) ;
//    X:=Valeur(G.Cells[7,i]) ; Ecr.PutValue('E_CREDITEURO', X) ;
    if DEV.Code=V_PGI.DevisePivot then
       begin
       Ecr.PutValue('E_DEBITDEV',  Ecr.GetValue('E_DEBIT')) ;
       Ecr.PutValue('E_CREDITDEV', Ecr.GetValue('E_CREDIT')) ;
       end ;
    end ;
end ;

Function TFSaisEcar.ConfirmeEcart : boolean ;
Var XF,XE,EcartE,EcartF,Coef : Double ;
    OkErr,GrossErr : boolean ;
    i,Nb : integer ;
BEGIN
OkErr:=False ; GrossErr:=False ; Result:=True ; Nb:=G.RowCount-1 ;
if Action=taConsult then Exit ; Coef:=1.0 ;
if V_PGI.TauxEuro>6 then Coef:=V_PGI.TauxEuro/5 ;
for i:=1 to G.RowCount-1 do
    BEGIN
    if VH^.TenueEuro then
       BEGIN
       XE:=Valeur(G.Cells[3,i])+Valeur(G.Cells[4,i]) ;
       XF:=Valeur(G.Cells[6,i])+Valeur(G.Cells[7,i]) ;
       EcartF:=Arrondi(EuroToPivot(XE)-XF,V_PGI.OkDecE) ;
       EcartE:=Arrondi(PivotToEuro(XF)-XE,V_PGI.OkDecV) ;
       if Abs(EcartF)>Coef*Nb*3*Resolution(V_PGI.OkDecE) then GrossErr:=True else
        if Abs(EcartE)>Coef*Nb*1*Resolution(V_PGI.OkDecV) then GrossErr:=True else
         if Abs(EcartF)>Coef*3*Resolution(V_PGI.OkDecE) then OkErr:=True else
          if Abs(EcartE)>Coef*1*Resolution(V_PGI.OkDecV) then OkErr:=True ;
       END else
       BEGIN
       XF:=Valeur(G.Cells[3,i])+Valeur(G.Cells[4,i]) ;
       XE:=Valeur(G.Cells[6,i])+Valeur(G.Cells[7,i]) ;
       EcartF:=Arrondi(EuroToPivot(XE)-XF,V_PGI.OkDecV) ;
       EcartE:=Arrondi(PivotToEuro(XF)-XE,V_PGI.OkDecE) ;
       if Abs(EcartF)>Coef*Nb*3*Resolution(V_PGI.OkDecV) then GrossErr:=True else
        if Abs(EcartE)>Coef*Nb*1*Resolution(V_PGI.OkDecE) then GrossErr:=True else
         if Abs(EcartF)>Coef*3*Resolution(V_PGI.OkDecV) then OkErr:=True else
          if Abs(EcartE)>Coef*1*Resolution(V_PGI.OkDecE) then OkErr:=True ;
       END ;
    if OkErr then Break ;
    END ;
if GrossErr then BEGIN Result:=False ; HDiv.Execute(5,'','') ; END else
   if OkErr then Result:=(HDiv.Execute(3,'','')=mrYes) ;
END ;

procedure TFSaisEcar.BValideClick(Sender: TObject);
begin
if Action<>taConsult then
   BEGIN
   if ((Arrondi(SA_SoldePivot.Value,V_PGI.OkDecV)<>0) or (Arrondi(SA_SoldeOppose.Value,V_PGI.OkDecE)<>0)) then BEGIN HDiv.Execute(1,'','') ; Exit ; END ;
   if Not ConfirmeEcart then Exit ;
   if GS<>nil then GVersO else TOBGVersO ;
   OkVal:=True ; OkModif:=False ;
   END ;
Close ;
end;

Procedure TFSaisEcar.GetCellCanvas(Acol,ARow : Longint ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if ((ARow<=0) or (ACol<3)) then Exit ;
if ARow=G.Row then Exit ;
if ACol in [3,4] then G.Canvas.Font.Color:=clGreen else
 if ACol in [6,7] then G.Canvas.Font.Color:=clRed else
  if ACol in [5,8] then G.Canvas.Font.Style:=G.Canvas.Font.Style+[fsBold] ;
END ;

procedure TFSaisEcar.FormateMontant ( ACol,ARow : integer ) ;
BEGIN
if ACol in [3,4] then G.Cells[ACol,ARow]:=StrF00(Valeur(G.Cells[ACol,ARow]),V_PGI.OkDecV) else
 if ACol in [6,7] then G.Cells[ACol,ARow]:=StrF00(Valeur(G.Cells[ACol,ARow]),V_PGI.OkDecE) ;
END ;

procedure TFSaisEcar.GCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not ModeModif then Exit ;
if G.Col in [1,2,5,8] then Cancel:=True else
 if ((Not EccPivot) and (G.Col in [3,4])) then Cancel:=True else
   if ((EccPivot) and (G.Col in [6,7])) then Cancel:=True else
    if ((G.Col=3) and (G.Cells[4,G.Row]<>'')) then Cancel:=True else
     if ((G.Col=4) and (G.Cells[3,G.Row]<>'')) then Cancel:=True else
      if ((G.Col=6) and (G.Cells[7,G.Row]<>'')) then Cancel:=True else
       if ((G.Col=7) and (G.Cells[6,G.Row]<>'')) then Cancel:=True ;
if Not Cancel then StCellCur:=G.Cells[G.Col,G.Row] ;
end;

procedure TFSaisEcar.GCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not ModeModif then Exit ;
if G.Cells[ACol,ARow]=StCellCur then Exit ;
FormateMontant(ACol,ARow) ;
CalculSoldes ;
end;

procedure TFSaisEcar.BModifClick(Sender: TObject);
begin
if Action=taConsult then Exit ;
BSolde.Enabled:=False ; ModeModif:=True ;
AffecteGrid(G,taModif) ;
end;

Function TFSaisEcar.JeValide : boolean ;
BEGIN
Result:=(HDiv.Execute(2,'','')=mrYes) ;
END ;

procedure TFSaisEcar.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if Action=taConsult then Exit ;
if OkModif then if Not JeValide then CanClose:=False ;
end;

procedure TFSaisEcar.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
