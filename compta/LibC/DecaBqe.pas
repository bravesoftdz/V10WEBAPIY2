unit DecaBqe;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Grids, Hctrls, ExtCtrls, Ent1, Hspliter,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  LettUtil,
  SaisComm,SaisUtil,PrintDBG, hmsgbox, HEnt1, HSysMenu ;

type TBQESolde = Class
  SoldeBanque : Double ;
  SoldeAttrib : Double ;
  SoldePrevu  : Double ;
  end ;

type
  TFAffBqe = class(TForm)
    Panel1: TPanel;
    BImprimer: THBitBtn;
    BValider: THBitBtn;
    BAnnuler: THBitBtn;
    BAide: THBitBtn;
    bZoomPiece: THBitBtn;
    GBqe: THGrid;
    Split: TSplitControl;
    Panel2: TPanel;
    G: THGrid;
    Panel3: TPanel;
    H_TOTDEVISE: TLabel;
    bUp: THBitBtn;
    bDown: THBitBtn;
    ER_SOLDED: THNumEdit;
    Msg: THMsgBox;
    cVisu: TCheckBox;
    Devise: TEdit;
    HMTrad: THSystemMenu;
    bSelectAll: THBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bDownClick(Sender: TObject);
    procedure bUpClick(Sender: TObject);
    procedure GDblClick(Sender: TObject);
    procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BImprimerClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure bZoomPieceClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    GX,GY : integer ;
    procedure FormateG; // FQ 21458 BVE
    procedure RempliBanques ;
    procedure AttribSelection(Plus : Boolean) ;
    procedure UpdateGridBqe(OBqe : TBqeSolde ; lig : integer) ;
    procedure CocheDecoche ( Lig : integer ; Next : boolean ) ;
    procedure CalculDebitCredit ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    function  CalcSoldeAttrib(CompteBanque : string) : Double ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  Saisie, SaisBor;

procedure TFAffBqe.RempliBanques ;
var sSQL  : string ;
    Q     : TQuery ;
    lig,n : Integer ;
    Solde : Double ;
    OBqe  : TBQESolde ;
    CompteBanque : string ;
begin
n:=0 ;
GBqe.VidePile(True) ;
GBqe.ColAligns[0]:=taLeftJustify ;
GBqe.ColAligns[1]:=taLeftJustify ;
GBqe.ColAligns[2]:=taRightJustify ;
GBqe.ColAligns[3]:=taRightJustify ;
GBqe.ColAligns[4]:=taRightJustify ;
sSQL:='SELECT * FROM BANQUECP LEFT JOIN GENERAUX ON G_GENERAL=BQ_GENERAL WHERE BQ_DEVISE="'+Devise.Text
     +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"' ; // 19/10/2006 YMO Multisociétés
Q:=OpenSQL(sSQL,True) ;
while not Q.EOF do
   begin
   CompteBanque:=Q.FindField('BQ_GENERAL').AsString ;
   Lig:=GBqe.RowCount-1 ;
   if GBqe.Cells[0,lig]<>'' then begin GBqe.RowCount:=GBqe.RowCount+1 ; Inc(lig) ; end ;
   GBqe.Cells[0,lig]:=CompteBanque ;
   GBqe.Cells[1,lig]:=Q.FindField('G_LIBELLE').AsString ;
   Solde:=Q.FindField('G_TOTALDEBIT').AsFloat-Q.FindField('G_TOTALCREDIT').AsFloat ;
   OBqe:=TBqeSolde.Create ;
   OBqe.SoldeBanque:=Solde ;
   OBqe.SoldeAttrib:=CalcSoldeAttrib(CompteBanque) ;
   OBqe.SoldePrevu:=OBqe.SoldeBanque+OBqe.SoldeAttrib ;
   GBqe.Objects[0,Lig]:=OBqe ;
   UpdateGridBqe(OBqe,lig) ;
   Inc(n) ;
   Q.Next ;
   end ;
Ferme(Q) ;
if n=0 then begin bDown.Enabled:=False ; bUp.Enabled:=False ; end ;
end ;

procedure TFAffBqe.UpdateGridBqe(OBqe : TBqeSolde ; lig : integer) ;
begin
GBqe.Cells[2,lig]:=StrFMontant(OBqe.SoldeBanque,15,2,'',True) ;
GBqe.Cells[3,lig]:=StrFMontant(OBqe.SoldeAttrib,15,2,'',True) ;
GBqe.Cells[4,lig]:=StrFMontant(OBqe.SoldePrevu,15,2,'',True) ;
end ;

function TFAffBqe.CalcSoldeAttrib(CompteBanque : string) : Double ;
var i : integer ;
    O : TOBM ;
    Solde : Double ;
begin
Solde:=0 ;
for i:=1 to G.RowCount-1 do
    begin
    O:=TOBM(G.Objects[0,i]) ;
    if O.GetMvt('E_BANQUEPREVI')=CompteBanque then Solde:=Solde+(O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT')) ;
    end ;
Result:=Solde ;
end ;

procedure TFAffBqe.AttribSelection(Plus : Boolean) ;
var i,j,n   : integer ;
    CpteBqe : string ;
    O       : TOBM ;
    OBqe    : TBqeSolde ;
    Solde   : Double ;
begin
OBqe:=nil ; j:=0 ; n:=0 ;
for i:=1 to G.RowCount-1 do if EstSelect(G,i) then Inc(n) ; // Test si éléments sélectionnés
if n=0 then begin Msg.Execute(1,caption,'') ; exit ; end ;
for i:=1 to G.RowCount-1 do
   begin
   if EstSelect(G,i) then
      begin
      O:=TOBM(G.Objects[0,i]) ;
      if Plus and (Trim(O.GetMvt('E_BANQUEPREVI'))<>'') then continue ; // déjà attribuée
      if (not Plus) and (Trim(O.GetMvt('E_BANQUEPREVI'))='') then continue ;
      if Plus then CpteBqe:=GBqe.Cells[0,GBqe.Row]
              else CpteBqe:=O.GetMvt('E_BANQUEPREVI') ;
      Solde:=O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT') ;
      if Plus then
         begin
         // Affectation
         OBqe:=TBqeSolde(GBqe.Objects[0,GBqe.Row]) ;
         OBqe.SoldeAttrib:=OBqe.SoldeAttrib+Solde ;
         OBqe.SoldePrevu:=OBqe.SoldeBanque+OBqe.SoldeAttrib ;
         O.PutMvt('E_BANQUEPREVI',CpteBqe) ;
         j:=GBqe.Row ;
         end else
         begin
         // Désaffectation -> recherche de la ligne banque concernée (j)
         if Trim(CpteBqe)='' then continue ;
         for j:=1 to GBqe.RowCount-1 do if GBqe.Cells[0,j]=CpteBqe then break ;
         OBqe:=TBqeSolde(GBqe.Objects[0,j]) ;
         OBqe.SoldeAttrib:=OBqe.SoldeAttrib-Solde ;
         OBqe.SoldePrevu:=OBqe.SoldeBanque-OBqe.SoldeAttrib ;
         O.PutMvt('E_BANQUEPREVI','') ;
         end ;
      // Mise à jour des grids
      ComComLigne(G,O,i) ;
      UpdateGridBqe(OBqe,j) ;
      CocheDecoche(i,False) ;
      end ;
   end ;
end ;

procedure TFAffBqe.CocheDecoche ( Lig : integer ; Next : boolean ) ;
BEGIN
if Lig<=0 then Exit ;
if EstSelect(G,Lig) then G.Cells[G.ColCount-1,Lig]:=' ' else G.Cells[G.ColCount-1,Lig]:='+' ;
if Lig=G.RowCount-1 then Next:=False ;
if Next then G.Row:=Lig+1 ;
G.Invalidate ;
CalculDebitCredit ;
END ;

procedure TFAffBqe.CalculDebitCredit ;
var i : integer ;
    O : TOBM ;
    CouvP,CouvD : Double ;
    TDP,TCP,TDD,TCD : Double ;
    CDD,CCD,CDP,CCP : Double ;
    // TCouvP,TCouvD : Double ; BVE 20.09.07
BEGIN
TDP:=0 ; TCP:=0 ; TDD:=0 ; TCD:=0 ; CDD := 0 ; CCD := 0 ; CDP := 0 ; CCP := 0 ;
for i:=1 to G.RowCount-1 do
   begin
   if EstSelect(G,i) then
      begin
      O:=TOBM(G.Objects[0,i]) ; if O=Nil then Break ;
      TDP:=TDP+O.GetMvt('E_DEBIT') ; TCP:=TCP+O.GetMvt('E_CREDIT') ; CouvP:=O.GetMvt('E_COUVERTURE') ;
      if O.GetMvt('E_DEBIT')<>0 then CDP:=CDP+CouvP else CCP:=CCP+CouvP ;
      TDD:=TDD+O.GetMvt('E_DEBITDEV') ; TCD:=TCD+O.GetMvt('E_CREDITDEV') ; CouvD:=O.GetMvt('E_COUVERTUREDEV') ;
      if O.GetMvt('E_DEBITDEV')<>0 then CDD:=CDD+CouvD else CCD:=CCD+CouvD ;
      end ;
   //TCouvP:=CCP-CDP ;   BVE 20.09.07 plus utilisé ?
   //TCouvD:=CCD-CDD ;
   end ;
AfficheLeSolde(ER_SOLDED,TDD-CDD,TCD-CCD) ;
END ;

{------------------------- Evenements ---------------------------}

Procedure TFAffBqe.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if (ARow<>0) then
   begin
   if EstSelect(G,ARow) then G.Canvas.Font.Style:=G.Canvas.Font.Style+[fsItalic]
                        else G.Canvas.Font.Style:=G.Canvas.Font.Style-[fsItalic] ;
   end ;
END ;

procedure TFAffBqe.FormCreate(Sender: TObject);
begin
G.VidePile(True) ;
{ FQ 21458 BVE 19.09.07
G.ListeParam:='AFFBQE' ; }
FormateG;
{ END FQ 21458 }
PopupMenu:=AddMenuPop(PopupMenu,'','') ;
end;

procedure TFAffBqe.FormShow(Sender: TObject);
begin
G.GetCellCanvas:=GetCellCanvas ;
RempliBanques ;
if cVisu.Checked then
   begin
   bDown.Visible:=False ;
   bUp.Visible:=False ;
   end ;
end;

procedure TFAffBqe.bDownClick(Sender: TObject);
begin
AttribSelection(True) ;
end;

procedure TFAffBqe.bUpClick(Sender: TObject);
begin
AttribSelection(False) ;
end;

procedure TFAffBqe.GDblClick(Sender: TObject);
begin
if G.Row>0 then CocheDecoche(G.Row,True) ;
end ;

procedure TFAffBqe.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var C,R : Longint ;
begin
GX:=X ; GY:=Y ;
if ((ssCtrl in Shift) and (Button=mbLeft)) then
   BEGIN
   G.MouseToCell(X,Y,C,R) ;
   if R>0 then CocheDecoche(G.Row,False) ;
   END ;
end;

procedure TFAffBqe.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : boolean ;
begin
OkG:=G.Focused ; Vide:=(Shift=[]) ;
Case Key of
   VK_F10    : if Vide then ModalResult:=mrOk ;
   VK_SPACE  : if ((OkG) and (Vide)) then CocheDecoche(G.Row,False) ;
   VK_F5     : BEGIN Key:=0 ; BZoomPieceClick(Nil) ; END ;
          80 : if Shift=[ssCtrl] then BImprimerClick(Nil) ;
  END ;
end;

procedure TFAffBqe.BImprimerClick(Sender: TObject);
begin
PrintDBGrid(G,nil,Caption,'') ;
end;

procedure TFAffBqe.BAnnulerClick(Sender: TObject);
Var i : integer ;
    O : TOBM ;
begin
if not cVisu.Checked then
   begin
   for i:=1 to G.RowCount-1 do
       BEGIN
       O:=TOBM(G.Objects[0,i]) ;
       O.PutMvt('E_BANQUEPREVI',O.GetMvt('E_NUMPIECEINTERNE')) ;
       O.PutMvt('E_NUMPIECEINTERNE','') ;
       END ;
   end ;
Close ;
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
procedure TFAffBqe.bZoomPieceClick(Sender: TObject);
var Lig : Integer ;
    O   : TOBM ;
    M   : RMVT ;
    P   : RParFolio ;
begin
{Zoom pièce}
Lig:=G.Row ; if ((Lig<=0) or (Lig>G.RowCount-1)) then Exit ;
O:=TOBM(G.Objects[0,Lig]) ; if O=Nil then Exit ;
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

procedure TFAffBqe.bSelectAllClick(Sender: TObject);
var i : integer ;
begin
for i:=1 to G.RowCount-1 do
   begin
   if not EstSelect(G,i) then CocheDecoche(i,False) ;
   end ;
G.Invalidate ;
end;

procedure TFAffBqe.BValiderClick(Sender: TObject);
var i,n : Integer ;
    O : TOBM ;
begin
n:=0 ;
for i:=1 to G.RowCount-1 do
   begin
   O:=TOBM(G.Objects[0,i]) ;
   if O.GetMvt('E_BANQUEPREVI')='' then Inc(n) ;
   end ;
if n>0 then if Msg.Execute(2,caption,'')<>mrYes then exit ;
ModalResult:=mrOK ;
end;

procedure TFAffBqe.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/09/2007
Modifié le ... :   /  /    
Description .. : Formate la liste G afin de remplacé le fonctionnement par 
Suite ........ : ListeParam, car la liste AFFBQE n'est plus dans la SOCREF 
Suite ........ : 850.
Suite ........ : FQ 21458
Mots clefs ... : 
*****************************************************************}
procedure TFAffBqe.FormateG; // FQ 21458 BVE
begin
  //                  0         1             2           3             4         5           6           7               8           9           10        11          12
  G.Titres.Text := 'E_JOURNAL;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_DEBITDEV;E_CREDITDEV;E_COUVERTURE;E_DATEECHEANCE;E_MODEPAIE;E_NUMLIGNE;E_NUMECHE;E_DATECOMPTABLE;;';
  G.ColWidths[0] := 30;
  G.ColWidths[1] := 50;
  G.ColWidths[2] := 65;
  G.ColWidths[3] := 65;
  G.ColWidths[4] := 60;
  G.ColWidths[5] := 60;
  G.ColWidths[6] := 60;
  G.ColWidths[7] := 60;
  G.ColWidths[8] := 75;
  G.ColWidths[9] := 30;
  G.ColWidths[10] := 30;
  G.ColWidths[11] := 30;
  G.ColWidths[12] := 75;
  G.ColWidths[13] := 0;
  G.ColAligns[0] := taCenter;
  G.ColAligns[1] := taRightJustify;
  G.ColAligns[2] := taLeftJustify;
  G.ColAligns[3] := taLeftJustify;
  G.ColAligns[4] := taLeftJustify;
  G.ColAligns[5] := taRightJustify;
  G.ColAligns[6] := taRightJustify;
  G.ColAligns[7] := taRightJustify;
  G.ColAligns[8] := taLeftJustify;
  G.ColAligns[9] := taCenter;   
  G.ColAligns[10] := taLeftJustify;
  G.ColAligns[11] := taLeftJustify;
  G.ColAligns[12] := taCenter;
  G.ColAligns[13] := taLeftJustify;

end;
end.
