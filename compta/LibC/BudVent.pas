{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/01/2005
Modifié le ... :   /  /    
Description .. : Portage en eAGL
Mots clefs ... : 
*****************************************************************}
unit BudVent;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms, StdCtrls,
  ExtCtrls, Buttons, Grids, HSysMenu, hmsgbox, HTB97, Hctrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  Ent1,     // ChangeMask
  Hent1,    // TActionFiche, Valeur, Arrondi, Format_String, BeginTrans, CommitTrans
  SaisUtil; // TModeSaisBud, TOBM, StrS, ADecimP

Type X_BV = RECORD
            T : TList ;
            Action : TActionFiche ;
            ModeSaisBud : TModeSaisBud ;
            OkD         : Boolean ;
            CurDec      : integer ;
            CurCoef     : Double ;
            LaPer       : String ;
            StC,Jal,Cpt,CatJal,GeneAttente,SectAttente : String ;
            END ;

Function VentilBud ( Var X : X_BV ) : Boolean ;

type
  TFBudVent = class(TForm)
    FListe: THGrid;
    PTot: TPanel;
    TTot: TLabel;
    TotalP: THNumEdit;
    TotalM: THNumEdit;
    HM: THMsgBox;
    HmTrad: THSystemMenu;
    Label1: TLabel;
    ResteP: THNumEdit;
    ResteM: THNumEdit;
    Label2: TLabel;
    TTotIni: TLabel;
    TotIni: THNumEdit;
    PVentil: TPanel;
    H_CODEVENTIL: THLabel;
    HLabel1: THLabel;
    H_TITREVENTIl: TLabel;
    Y_CODEVENTIL: TEdit;
    Y_LIBELLEVENTIL: TEdit;
    BNewVentil: THBitBtn;
    PFenCouverture: TPanel;
    BAbandonVentil: THBitBtn;
    Panel1: TPanel;
    TitreVentil: TLabel;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Outils97: TToolbar97;
    H_CVType: THLabel;
    CVType: THValComboBox;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    BSauveVentil: TToolbarButton97;
    BDelete: TToolbarButton97;
    BSolde: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure FListeCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BSoldeClick(Sender: TObject);
    procedure BSauveVentilClick(Sender: TObject);
    procedure BNewVentilClick(Sender: TObject);
    procedure CVTypeChange(Sender: TObject);
    procedure BAbandonVentilClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
  private
    MemoCell : String;
    Modifier,ModifVent : Boolean ;
    Total : Double ;
    ListeAuto : HTStrings ;
    Procedure RempliLeGrid ;
    Procedure RecupereMontant ;
    Procedure CalculPourcent ;
    Procedure FaitTotal ;
    Procedure FormatZone(ACol,ARow : Integer) ;
    Procedure CalculColonne(ACol,ARow : Integer ) ;
    Function  RechercheLobm(Cpte : String) : TOBM ;
    Procedure EcritLOBM ;
    Procedure MajOBM(O : TOBM ; Lig : integer) ;
    procedure ClickSolde ;
    procedure ClickSauveVentil ;
    procedure SauveLaVentil ;
    procedure ChargeVentils ;
    procedure ChargeAutos ;
    procedure LoadLaVentil ;
    procedure FinVentil ( OkVal : boolean ) ;
    Function  CodeVentil : String ;
  public
    X : X_BV ;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  UTob;

Function VentilBud ( Var X : X_BV ) : Boolean ;
Var FBudVent : TFBudVent ;
BEGIN
Result:=False ;
if (X.T=Nil) or (X.StC='') then Exit ;
FBudVent:=TFBudVent.Create(Application) ;
 Try
  FBudVent.X:=X ;
  if FBudVent.ShowModal=mrOk then BEGIN Result:=True ; X:=FBudVent.X ; END ;
 Finally
  FBudVent.Free ;
 End ;
END ;

Procedure TFBudVent.RempliLeGrid ;
Var i : Integer ;
    St,StC : String ;
BEGIN
i:=1 ;
{Comptes d'attente}
if X.ModeSaisBud=msbGene then FListe.Cells[0,i]:=X.SectAttente
                         else FListe.Cells[0,i]:=X.GeneAttente ;
FListe.RowCount:=FListe.RowCount+1 ;
Inc(i) ;
{Autres comptes}
St:=X.Stc ;
While St<>'' do
   BEGIN
   StC:=ReadTokenSt(St) ;
   if ((StC<>'') and (St<>'')) then // St<>'' --> on ne prend pas le dernier ie cpt attente
      BEGIN
      if ListeAuto.IndexOf(StC)<0 then Continue ;
      FListe.Cells[0,i]:=StC ;
      FListe.RowCount:=FListe.RowCount+1 ;
      Inc(i) ;
      END ;
   END ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
RecupereMontant ; CalculPourcent ; FaitTotal ;
END ;

Procedure TFBudVent.CalculPourcent ;
Var i : Integer ;
BEGIN
for i:=1 to FListe.RowCount-1 do if Fliste.Cells[2,i]<>'' then
    FListe.Cells[1,i]:=StrS((Valeur(Fliste.Cells[2,i])*X.CurCoef/Total)*100,ADecimP) ;
END ;

Procedure TFBudVent.FaitTotal ;
Var i : Integer ;
    TotP,TotM,XResteP,XResteM,X1,X2 : Double ;
BEGIN
TotP:=0 ; TotM:=0 ; XResteP:=0 ; XResteM:=0 ;
for i:=1 to FListe.RowCount-1 do
   BEGIN
   X1:=Valeur(Fliste.Cells[1,i]) ; X2:=Valeur(Fliste.Cells[2,i]) ;
   TotP:=TotP+X1 ; TotM:=TotM+X2 ;
   if i>1 then BEGIN XResteP:=XResteP+X1 ; XResteM:=XResteM+X2 ; END ;
   END ;
XResteP:=100.0-XResteP ; XResteM:=TotIni.Value-XResteM ;
TotalP.Value:=TotP ; TotalM.Value:=TotM ;
ResteP.Value:=XResteP ; ResteM.Value:=XResteM ;
END ;

Procedure TFBudVent.RecupereMontant ;
Var O : TOBM ;
    Tot : Double ;
    i,Lig : Integer ;
    Nom,DC : String ;
BEGIN
if X.ModeSaisBud=msbGene then  Nom:='BE_BUDSECT' else Nom:='BE_BUDGENE' ;
  if X.OkD then
    DC:='BE_DEBIT'
  else
    DC:='BE_CREDIT' ;

for i:=0 to X.T.Count-1 do
   BEGIN
   O:=TOBM(X.T[i]) ;
   Lig:=FListe.Cols[0].IndexOf(O.GetMvt(Nom)) ; if Lig<=0 then Continue ;
   Tot:=O.GetMvt(DC) ;
   FListe.Cells[2,Lig]:=StrS((Valeur(FListe.Cells[2,Lig])+Tot)/X.CurCoef,X.CurDec) ;
   Total:=Total+Tot ;
   END ;
TotIni.Value:=Total/X.CurCoef ;
END ;

procedure TFBudVent.ChargeAutos ;
Var Q   : TQuery ;
    SQL,Nom,QuelChamp : String ;
    i   : integer ;
    O   : TOBM ;
BEGIN
QuelChamp:=X.Jal ; If X.CatJal<>'' Then QuelChamp:=X.CatJal ;
if X.ModeSaisBud=msbGene then
   BEGIN
   SQL:='Select CX_SECTION from CROISCPT Where CX_TYPE="BUD" AND CX_JAL="'+QuelChamp+'" AND CX_COMPTE="'+X.Cpt+'"' ;
   Nom:='BE_BUDSECT' ;
   END else
   BEGIN
   SQL:='Select CX_COMPTE  from CROISCPT Where CX_TYPE="BUD" AND CX_JAL="'+QuelChamp+'" AND CX_SECTION="'+X.Cpt+'"' ;
   Nom:='BE_BUDGENE' ;
   END ;
{Comptes déjà saisis}
for i:=0 to X.T.Count-1 do
    BEGIN
    O:=TOBM(X.T[i]) ; if O<>Nil then ListeAuto.Add(O.GetMvt(Nom)) ;
    END ;
{Croisements autorisés}
Q:=OpenSQL(SQL,True) ;
While Not Q.EOF do BEGIN ListeAuto.Add(Q.Fields[0].AsString) ; Q.Next ; END ;
Ferme(Q) ;
END ;

procedure TFBudVent.ChargeVentils ;
Var Q,QB : TQuery ;
    Code,Libelle : String ;
BEGIN
CVType.Items.Clear ; CVType.Values.Clear ;
if X.ModeSaisBud<>msbGene then Exit ;
Q:=OpenSQL('Select CC_CODE, CC_LIBELLE from CHOIXCOD Where CC_TYPE="VTB"',True) ;
While Not Q.EOF do
   BEGIN
   Code:=Q.Fields[0].AsString ; Libelle:=Q.Fields[1].AsString ;
   QB:=OpenSQL('SELECT * from VENTIL Where V_NATURE="BUD" AND V_COMPTE like"'+Code+'%"',True) ;
   if Not QB.EOF then
      BEGIN
      CVType.Items.Add(Libelle) ;
      CVType.Values.Add(Code) ;
      END ;
   Ferme(QB) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

procedure TFBudVent.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(True) ; ListeAuto.Free ;
RegSaveToolbarPos(Self,'BudVent') ;
end;

procedure TFBudVent.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
ListeAuto:= HTStringList.Create ;
RegLoadToolbarPos(Self,'BudVent') ;
end;

procedure TFBudVent.FormShow(Sender: TObject);
begin
ChargeAutos ; ChargeVentils ; ModifVent:=False ;
ChangeMask(TotIni,X.CurDec,'') ; ChangeMask(TotalM,X.CurDec,'') ; ChangeMask(TotalP,ADecimP,'') ;
ChangeMask(ResteP,ADecimP,'') ; ChangeMask(ResteM,X.CurDec,'') ;
MemoCell:='' ; Modifier:=False ; Total:=0 ;
RempliLeGrid ;
Case X.Action of
     taCreat : Caption:=HM.Mess[0] ;
     taModif : Caption:=HM.Mess[1] ;
   taConsult : BEGIN Caption:=HM.Mess[2] ; HelpContext:=15227100 ; END ; 
   End ;
AffecteGrid(FListe,X.Action) ;
if X.Action=taConsult then
   BEGIN
   BSolde.Visible:=False ; CVType.Visible:=False ; H_CVType.Visible:=False ;
   BSauveVentil.Visible:=False ; BDelete.Visible:=False ;
   Outils97.Visible:=False ;
   END ;
if X.ModeSaisBud in [msbSect,msbSectGene] then
   BEGIN
   BSolde.Hint:=HM.Mess[5] ;
   BSauveVentil.Visible:=False ;
   BDelete.Visible:=False ;
   CVType.Visible:=False ; H_CVType.Visible:=False ;
   END ;
FListe.Col:=1 ; FListe.Row:=1 ;
if FListe.CanFocus then FListe.SetFocus ;
if X.Action<>taConsult then FListe.MontreEdit ;
TitreVentil.Caption:=TitreVentil.Caption+X.Cpt+'   '+X.LaPer ;
UpdateCaption(Self) ;
end;

procedure TFBudVent.BFermeClick(Sender: TObject);
begin
if Modifier then
   BEGIN
   if HM.Execute(3,'','')=mrYes then BEGIN BValideClick(Nil) ; Exit ; END ;
   END ;
Close ;
end;

procedure TFBudVent.BValideClick(Sender: TObject);
Var ii : integer ;
begin
if X.Action=taConsult then BEGIN Close ; Exit ; END ;
if Not Modifier then BEGIN Close ; Exit ; END ;
if Arrondi(TotalM.Value-(Total/X.CurCoef),X.CurDec)<>0 then
   BEGIN
   ii:=HM.Execute(4,'','') ;
   Case ii of
      mrYes    : ;
      mrNo     : BEGIN Modifier:=False ; ModalResult:=mrCancel ; END ;
      mrCancel : Exit ;
      END ;
   END ;              
if Modifier then BEGIN EcritLOBM ; ModalResult:=mrOk ; END ;
end;

Procedure TFBudVent.FormatZone(ACol,ARow : Integer) ;
Var LaVal : Double ;
    D     : integer ;
BEGIN
if ACol=2 then D:=X.CurDec else D:=ADecimP ;
LaVal:=Valeur(Fliste.Cells[ACol,ARow]) ;
if LaVal=0 then FListe.Cells[ACol,ARow]:='' else FListe.Cells[ACol,ARow]:=StrS(LaVal,D) ;
END ;

Procedure TFBudVent.CalculColonne(ACol,ARow : Integer ) ;
BEGIN
if ACol=1 then FListe.Cells[2,ARow]:=StrS((Total*Valeur(FListe.Cells[1,ARow])/X.CurCoef)/100,X.CurDec)
          else FListe.Cells[1,ARow]:=StrS((Valeur(FListe.Cells[2,ARow])*100*X.CurCoef)/Total,ADecimP) ;
END ;

procedure TFBudVent.FListeCellExit(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
begin
if X.Action=taConsult then Exit ;
if MemoCell=FListe.Cells[ACol,ARow] then Exit ;
FormatZone(ACol,ARow) ;
CalculColonne(ACol,ARow) ;
FaitTotal ;
Modifier:=True ;
end;

procedure TFBudVent.FListeCellEnter(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
begin
MemoCell:=FListe.Cells[FListe.Col,FListe.Row] ;
end;

Function TFBudVent.RechercheLobm(Cpte : String) : TOBM ;
Var i   : Integer ;
    O   : TOBM ;
    Nom : String ;
BEGIN
Result:=Nil ;
if X.ModeSaisBud=msbGene then Nom:='BE_BUDSECT' else Nom:='BE_BUDGENE' ;
for i:=0 to X.T.Count-1 do
    BEGIN
    O:=TOBM(X.T[i]) ;
    if O.GetMvt(Nom)=Cpte then BEGIN Result:=O ; Break ; END ;
    END ;
END ;

Procedure TFBudVent.MajOBM(O : TOBM ; Lig : integer ) ;
Var Nom : String ;
    XSais,XD,XC : Double ;
BEGIN
if O=Nil then Exit ;
XSais:=Valeur(FListe.Cells[2,Lig])*X.CurCoef ;
XD:=0 ; XC:=0 ;
if X.OkD then XD:=XSais else XC:=XSais ;
if X.ModeSaisBud=msbGene then Nom:='BE_BUDSECT' else Nom:='BE_BUDGENE' ;
O.SetMontantsBUD(XD,XC) ;
O.PutMvt(Nom,FListe.Cells[0,Lig]) ;
END ;

Procedure TFBudVent.EcritLOBM ;
Var i : Integer ;
    OB,OS : TOBM ;
BEGIN
if Not Modifier then Exit ;
for i:=1 to FListe.RowCount-1 do
    BEGIN
    OB:=RechercheLobm(FListe.Cells[0,i]) ;
    if OB<>Nil then MajOBM(OB,i) else
       BEGIN
       if FListe.Cells[2,i]<>'' then
          BEGIN
          OB:=Nil ; OS:=TOBM(X.T[0]) ;
          EgaliseOBM(OS,OB) ; MajOBM(OB,i) ;
          X.T.Add(OB) ;
          END ;
       END ;
    END ;
Modifier:=False ;
END ;

procedure TFBudVent.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
     VK_RETURN : if Shift=[] then Key:=VK_TAB ;
     VK_ESCAPE : if Shift=[] then Close ;
     VK_F10    : if Shift=[] then BEGIN Key:=0 ; BValideClick(Nil) ; END ;
     VK_F6     : if Shift=[] then BEGIN Key:=0 ; ClickSolde ; END ;
     END ;
end;

procedure TFBudVent.LoadLaVentil ;
Var Q     : TQuery ;
    i,Lig : integer ;
    Sect  : String ;
BEGIN
if X.Action=taConsult then Exit ;
if Not CVType.Visible then Exit ;
if X.ModeSaisBud in [msbSect,msbSectGene] then Exit ;
if ((X.Jal='') or (CVType.Value='')) then Exit ;
Q:=OpenSQL('SELECT V_SECTION, V_TAUXMONTANT From VENTIL Where V_NATURE="BUD" AND V_COMPTE="'+CodeVentil+'"',True) ;
if Not Q.EOF then
   BEGIN
   for i:=1 to FListe.RowCount-1 do BEGIN FListe.Cells[1,i]:='' ; FListe.Cells[2,i]:='' ; END ;
   While Not Q.EOF do
      BEGIN
      Sect:=Q.Fields[0].AsString ;
      Lig:=Fliste.Cols[0].IndexOf(Sect) ;
      if Lig>=1 then
         BEGIN
         FListe.Cells[1,Lig]:=StrS(Q.Fields[1].AsFloat,ADecimP) ;
         CalculColonne(1,Lig) ;
         END ;
      Q.Next ;
      END ;
   END ;
Ferme(Q) ;
FaitTotal ;
Modifier:=True ;
END ;

procedure TFBudVent.SauveLaVentil ;
Var StW    : String ;
    i,NumV : integer ;
    Ventil : Tob;
BEGIN
if X.Action=taConsult then Exit ;
if Not BSauveVentil.Visible then Exit ;
if X.ModeSaisBud in [msbSect,msbSectGene] then Exit ;
if X.Jal='' then Exit ;
StW:='V_NATURE="BUD" AND V_COMPTE="'+Format_String(Y_CODEVENTIL.Text,3)+Format_String(X.Jal,3)+'"' ; NumV:=0 ;
ExecuteSQL('DELETE FROM VENTIL Where '+StW) ;

  for i:=1 to FListe.RowCount-1 do if Valeur(FListe.Cells[2,i])<>0 then begin
    Inc(NumV) ;
    Ventil := TOB.Create('VENTIL', nil, -1) ;
    Ventil.InitValeurs ;
    Ventil.PutValue('V_NATURE', 'BUD');
    Ventil.PutValue('V_COMPTE', Format_String(Y_CODEVENTIL.Text,3)+Format_String(X.Jal,3));
    Ventil.PutValue('V_SECTION', FListe.Cells[0,i]);
    Ventil.PutValue('V_TAUXMONTANT', Valeur(FListe.Cells[1,i]));
    Ventil.PutValue('V_NUMEROVENTIL', NumV);
    Ventil.InsertDB(nil);
    Ventil.Free;
  end;
end;

procedure TFBudVent.ClickSolde ;
Var Delta : Double ;
BEGIN
if Not BSolde.Visible then Exit ;
if X.Action=taConsult then Exit ;
Delta:=Arrondi(TotIni.Value-TotalM.Value,X.CurDec) ; if Delta=0 then Exit ;
FListe.Cells[2,1]:=StrS(Valeur(FListe.Cells[2,1])+Delta,X.CurDec) ;
CalculColonne(2,1) ;
FaitTotal ;
END ;

procedure TFBudVent.BSoldeClick(Sender: TObject);
begin ClickSolde ; end;

procedure TFBudVent.FinVentil ( OkVal : Boolean ) ;
BEGIN
PTot.Enabled:=True ; FListe.Enabled:=True ;
Outils97.Enabled:=True ; Valide97.Enabled:=True ;
FListe.SetFocus ; PVentil.Visible:=False ;
if OkVal then
   BEGIN
   CVType.Items.Add(Y_LIBELLEVENTIL.Text) ;
   CVType.Values.Add(Y_CODEVENTIL.Text) ;
   ModifVent:=True ; CVType.ItemIndex:=CVType.Items.Count-1 ; ModifVent:=False ;
   END ;
END ;

procedure TFBudVent.ClickSauveVentil ;
BEGIN
if X.Action=taConsult then Exit ;
PVentil.Left:=FListe.Left+(FListe.Width-PVentil.Width) div 2 ;
PVentil.Visible:=True ; Y_CODEVENTIL.SetFocus ;
PTot.Enabled:=False ; FListe.Enabled:=False ;
Outils97.Enabled:=False ; Valide97.Enabled:=False ;
END ;

procedure TFBudVent.BSauveVentilClick(Sender: TObject);
begin
ClickSauveVentil ;
end;

procedure TFBudVent.BNewVentilClick(Sender: TObject);
Var
  Lib       : String ;
  ChoixCod : Tob;
begin
if ((Y_CODEVENTIL.Text='') or (Y_LIBELLEVENTIL.Text='')) then BEGIN HM.Execute(6,'','') ; Exit ; END ;
Lib:=uppercase(Y_LIBELLEVENTIL.Text) ;
if ExisteSQL('Select * from CHOIXCOD Where CC_TYPE="VTB" AND (CC_CODE="'+Y_CODEVENTIL.Text+'" OR UPPER(CC_LIBELLE)="'+Lib+'")') then BEGIN HM.Execute(7,'','') ; Exit ; END ;
BeginTrans ;
ChoixCod := TOB.Create('CHOIXCOD', nil, -1) ;
ChoixCod.InitValeurs ;
ChoixCod.PutValue('CC_TYPE', 'VTB');
ChoixCod.PutValue('CC_CODE', Y_CODEVENTIL.Text);
ChoixCod.PutValue('CC_LIBELLE', Y_LIBELLEVENTIL.Text);
ChoixCod.InsertDB(nil);
ChoixCod.Free;
SauveLaVentil ;
CommitTrans ;
FinVentil(True) ;
end;

procedure TFBudVent.CVTypeChange(Sender: TObject);
begin
if Not ModifVent then LoadLaVentil ;
end;

procedure TFBudVent.BAbandonVentilClick(Sender: TObject);
begin
FinVentil(False) ;
end;

Function TFBudVent.CodeVentil : String ;
BEGIN
Result:=Format_String(CVType.Value,3)+Format_String(X.Jal,3) ;
END ;

procedure TFBudVent.BDeleteClick(Sender: TObject);
Var ii : integer ;
begin
if X.Action=taConsult then Exit ;
if X.ModeSaisBud<>msbGene then Exit ;
if CVType.Value='' then Exit ;
if HM.Execute(8,'','')<>mrYes then Exit ;
BeginTrans ;
ExecuteSQL('DELETE from VENTIL Where V_NATURE="BUD" and V_COMPTE="'+CodeVentil+'"') ;
ExecuteSQL('DELETE from CHOIXCOD Where CC_TYPE="VTB" and CC_CODE="'+CVType.Value+'"') ;
CommitTrans ;
ii:=CVType.ItemIndex ;
CVType.Items.Delete(ii) ; CVType.Values.Delete(ii) ; CVType.Value:='' ;
end;

procedure TFBudVent.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
