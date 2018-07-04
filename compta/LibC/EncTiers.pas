unit EncTiers;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Grids, Hctrls, Ent1, HEnt1, PrintDBG, Saisutil,
{$IFDEF VER150}
   Variants,
{$ENDIF}  
  Lettutil, SaisComm, EcheUnit ;

function EncaDecaTiers(Aux : string ; TEche : TList ; var TotDebCur,TotCredCur,TotCouvCur : Double ; Enc : Boolean ; DEV : RDEVISE ; LettreCheque,ModeOppose : Boolean) : boolean ;

type
  TFEncTiers = class(TForm)
    G: THGrid;
    P: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
    H_TOTDEVISE: TLabel;
    ER_SOLDED: THNumEdit;
    bZoomPiece: THBitBtn;
    bModifEche: THBitBtn;
    procedure FormShow(Sender: TObject);
    procedure GDblClick(Sender: TObject);
    procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure bZoomPieceClick(Sender: TObject);
    procedure bModifEcheClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Déclarations privées }
    Aux  : string ;
    TEche : TList ;
    TotDebCur,TotCredCur,TotCouvCur : Double ;
    GX,GY : integer ;
    Enc,LettreCheque,ModeOppose : Boolean ;
    WMinX,WMinY : Integer ;
    procedure MontreEcr ;
    procedure ComComEcr ( O : TOBM ) ;
    procedure CalculDebitCredit ;
    procedure CocheDecoche ( Lig : integer ; Next : boolean ) ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
  public
    DEV : RDEVISE ;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  Saisie, SaisBor;

function EncaDecaTiers(Aux : string ; TEche : TList ; var TotDebCur,TotCredCur,TotCouvCur : Double ; Enc : Boolean ; DEV : RDEVISE ; LettreCheque,ModeOppose : Boolean) : boolean ;
var X : TFEncTiers ;
begin
X:=TFEncTiers.Create(Application) ;
 try
  X.Aux:=Aux ; X.TEche:=TEche ; X.Enc:=Enc ;
  X.DEV:=DEV ; X.LettreCheque:=LettreCheque ; X.ModeOppose:=ModeOppose ;
  Result:=(X.ShowModal=mrOK) ;
  TotDebCur:=X.TotDebCur ; TotCredCur:=X.TotCredCur ; TotCouvCur:=X.TotCouvCur ;
 finally
  X.Free ;
 end ;
end ;

procedure TFEncTiers.MontreEcr ;
var i     : integer ;
    O     : TOBM ;
    XD,XC,XV : Double ;
BEGIN
G.VidePile(True) ;
for i:=0 to TEche.Count-1 do
    BEGIN
    O:=TOBM(TEche[i]) ;
    {Sauvegarde zones OBM}
    XD:=O.GetMvt('E_DEBIT') ; XC:=O.GetMvt('E_CREDIT') ; XV:=O.GetMvt('E_COUVERTURE') ;
    {Manipulation pour affichage}
    if DEV.Code<>V_PGI.DevisePivot then
       BEGIN
       O.PutMvt('E_DEBIT',O.GetMvt('E_DEBITDEV')) ;
       O.PutMvt('E_CREDIT',O.GetMvt('E_CREDITDEV')) ;
       O.PutMvt('E_COUVERTURE',O.GetMvt('E_COUVERTUREDEV')) ;
       END ;
    {Affichage}
    ComComEcr(O) ;
    {Restauration zones OBM}
    O.PutMvt('E_DEBIT',XD) ; O.PutMvt('E_CREDIT',XC) ; O.PutMvt('E_COUVERTURE',XV) ;
    G.RowCount:=G.RowCount+1 ;
    END ;
G.RowCount:=G.RowCount-1 ;
CalculDebitCredit ;
END ;

procedure TFEncTiers.ComComEcr ( O : TOBM ) ;
Var StC,STitreCol,StCell : String ;
    C,Lig : integer ;
    V : Variant ;
begin
StC:=G.Titres[0] ; C:=0 ; Lig:=G.RowCount-1 ;
repeat
 STitreCol:=ReadTokenSt(StC) ; if STitreCol='' then Break ;
 V:=O.GetMvt(STitreCol) ;
 if ((G.ColFormats[C]<>'') and (VarType(V)=VarDouble)) then StCell:=FormatFloat(G.ColFormats[C],VarAsType(V,VarDouble))
                                                       else StCell:=VarAsType(V,VarString) ;
 if C<>G.ColCount-1 then G.Cells[C,Lig]:=StCell ;
 inc(C) ;
until ((StC='') or (STitreCol='') or (C>=G.ColCount)) ;
G.Cells[G.ColCount-1,Lig]:=O.GetMvt('E_ETAT') ;
end ;

procedure TFEncTiers.CocheDecoche ( Lig : integer ; Next : boolean ) ;
var O : TOBM ;
BEGIN
if Lig<=0 then Exit ; O:=TOBM(TEche[Lig-1]) ; if O=Nil then Exit ;
if EstSelect(G,Lig) then G.Cells[G.ColCount-1,Lig]:=' ' else G.Cells[G.ColCount-1,Lig]:='+' ;
TEche[Lig-1]:=O ;
if Lig=G.RowCount-1 then Next:=False ;
if Next then G.Row:=Lig+1 ;
G.Invalidate ;
CalculDebitCredit ;
END ;

procedure TFEncTiers.CalculDebitCredit ;
var i : integer ;
    O : TOBM ;
    DebCur,CredCur,CouvDebCur,CouvCredCur,CouvCur : Double ;
BEGIN
CouvDebCur:=0 ; CouvCredCur:=0 ;
TotDebCur:=0 ; TotCredCur:=0 ; TotCouvCur:=0 ;
for i:=1 to G.RowCount-1 do if EstSelect(G,i) then
    begin
    DebCur:=Valeur(G.Cells[3,i]) ; CredCur:=Valeur(G.Cells[4,i]) ; CouvCur:=Valeur(G.Cells[5,i]) ;
    O:=TOBM(TEche[i-1]) ; if O=Nil then Break ;
    TotDebCur:=TotDebCur+DebCur ; TotCredCur:=TotCredCur+CredCur ;
    if DebCur<>0 then CouvDebCur:=CouvDebCur+CouvCur else CouvCredCur:=CouvCredCur+CouvCur ;
    end ;
if Enc then TotCouvCur:=CouvDebCur-CouvCredCur else TotCouvCur:=CouvCredCur-CouvDebCur ;
AfficheLeSolde(ER_SOLDED,TotDebCur-CouvDebCur,TotCredCur-CouvCredCur) ;
END ;

{---------------------------------- Evènements -----------------------------------------}
procedure TFEncTiers.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if (ARow<>0) And (EstSelect(G,ARow)) then
   G.Canvas.Font.Style:=G.Canvas.Font.Style+[fsItalic] else
   G.Canvas.Font.Style:=G.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFEncTiers.FormShow(Sender: TObject);
Var Deci : integer ;
begin
G.GetCellCanvas:=GetCellCanvas ;
if DEV.Code<>V_PGI.DevisePivot then Deci:=DEV.Decimale else
   if ModeOppose then Deci:=V_PGI.OkDecE else Deci:=V_PGI.OkDecV ;
ChangeMask(ER_SOLDED,Deci,DEV.Symbole) ;
Caption:=Caption+' '+Aux ;
bModifEche.Enabled:=not LettreCheque ;
G.ListeParam:='ENCEUROTIERS' ; G.RowCount:=2 ;
MontreEcr ;
UpdateCaption(Self) ;
end;

procedure TFEncTiers.GDblClick(Sender: TObject);
begin
if G.Row>0 then CocheDecoche(G.Row,True) ;
end;

procedure TFEncTiers.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var C,R : Longint ;
begin
GX:=X ; GY:=Y ;
if ((ssCtrl in Shift) and (Button=mbLeft)) then
   BEGIN
   G.MouseToCell(X,Y,C,R) ;
   if (R>0) then CocheDecoche(G.Row,False) ;
   END ;
end;

procedure TFEncTiers.BValiderClick(Sender: TObject);
var i : integer ;
    O : TOBM ;
begin
for i:=1 to G.RowCount-1 do
   begin
   O:=TOBM(TEche[i-1]) ;
   if EstSelect(G,i) then O.PutMvt('E_ETAT','*') else O.PutMvt('E_ETAT',' ') ;
   end ;
ModalResult:=mrOK ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 18/08/2004
Suite ........ : - LG - 18/08/2004 - Suppression de la fct debutdemois pour 
Suite ........ : l'appel de la saisie bor, ne fct pas avec les exercices 
Suite ........ : decalees
Mots clefs ... : 
*****************************************************************}
procedure TFEncTiers.bZoomPieceClick(Sender: TObject);
var Lig : Integer ;
    O   : TOBM ;
    M   : RMVT ;
    P   : RParFolio ;
begin
{Zoom pièce}
Lig:=G.Row ; if ((Lig<=0) or (Lig>G.RowCount-1)) then Exit ;
if Lig>TEche.Count then Exit ;
O:=TOBM(TEche[Lig-1]) ; if O=Nil then Exit ;
M:=OBMToIdent(O,False) ;
if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>'')) then
   BEGIN
   FillChar(P, Sizeof(P), #0) ;
   P.ParPeriode:=DateToStr(O.GetMvt('E_DATECOMPTABLE')) ;
   P.ParCodeJal:=O.GetMvt('E_JOURNAL') ;
   P.ParNumFolio:=IntToStr(O.GetMvt('E_NUMEROPIECE')) ;
   P.ParNumLigne:=O.GetMvt('E_NUMLIGNE') ;
   ChargeSaisieFolio(P, taConsult) ;
   END else
   BEGIN
   LanceSaisie(Nil,taConsult,M) ;
   END ;
end;

procedure TFEncTiers.bModifEcheClick(Sender: TObject);
var Lig,k : Integer ;
    O   : TOBM ;
    M   : RMVT ;
    EU  : T_ECHEUNIT ;
    TAN : String3 ;
    Coll  : String ;
begin
Lig:=G.Row ; if ((Lig<=0) or (Lig>G.RowCount-1)) then Exit ;
if Lig>TEche.Count then Exit ;
O:=TOBM(TEche[Lig-1]) ; if O=Nil then Exit ;
if O.GetMvt('E_ANA')='X' then Exit ; 
M:=OBMToIdent(O,True) ;
FillChar(EU,Sizeof(EU),#0) ;
EU.DateEche:=O.GetMvt('E_DATEECHEANCE') ; EU.ModePaie:=O.GetMvt('E_MODEPAIE') ;
EU.DebitDEV:=O.GetMvt('E_DEBITDEV') ; EU.CreditDEV:=O.GetMvt('E_CREDITDEV') ;
//EU.DebitEuro:=O.GetMvt('E_DEBITEURO') ;
//EU.CreditEuro:=O.GetMvt('E_CREDITEURO') ;
EU.Debit:=O.GetMvt('E_DEBIT') ; EU.Credit:=O.GetMvt('E_CREDIT') ;
EU.DEVISE:=O.GetMvt('E_DEVISE') ; EU.TauxDEV:=O.GetMvt('E_TAUXDEV') ;
EU.DateComptable:=O.GetMvt('E_DATECOMPTABLE') ;
EU.DateModif:=O.GetMvt('E_DATEMODIF') ;
EU.ModeSaisie:=O.GetMvt('E_MODESAISIE') ;
//if EuroOK then EU.SaisieEuro:=O.GetMvt('E_SAISIEEURO')='X' else EU.SaisieEuro:=False ;
{#TVAENC}
if VH^.OuiTvaEnc then
   BEGIN
   Coll:=O.GetMvt('E_GENERAL') ;
   if EstCollFact(Coll) then
      BEGIN
      for k:=1 to 4 do EU.TabTva[k]:=O.GetMvt('E_ECHEENC'+IntToStr(k)) ;
      EU.TabTva[5]:=O.GetMvt('E_ECHEDEBIT') ;
      END ;
   END ;
TAN:=O.GetMvt('E_ECRANOUVEAU') ;
if TAN='OAN' then
   BEGIN
   if M.CodeD<>V_PGI.DevisePivot then Exit ;
   if ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then Exit ;
   END ;
if ModifUneEcheance(M,EU) then MontreEcr ;
end;

procedure TFEncTiers.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFEncTiers.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=163 ;
end;

procedure TFEncTiers.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : Boolean ;
begin
OkG:=(G=ActiveControl) ; Vide:=(Shift=[]) ;
Case Key of
   VK_SPACE  : if ((OkG) and (Vide)) then CocheDecoche(G.Row,False) ;
   VK_F5     : if Vide then BZoomPieceClick(Nil) ;
   VK_F10    : if Vide then BValiderClick(Nil) ;
   END ;
end;

end.
