{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/03/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit TvaLettr;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,  // TCanvas, clRed, clGreen
  Buttons,
  Controls,
  Forms,
  Dialogs,
  Grids,
  Hctrls,
  StdCtrls,
  ExtCtrls,
  HSysMenu,
  HMsgBox,
  SaisUtil,
  Ent1,
  HEnt1,
  SaisComm,
{$IFDEF EAGLCLIENT}
  uTOB,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Lettrage,
  ed_tools
  ;

Procedure TvaAccLettrage ( O : TOBM ) ;

type
  TFTvaLettr = class(TForm)
    POutils: TPanel;
    BValide: THBitBtn;
    BAbandon: THBitBtn;
    BAide: THBitBtn;
    PFen: TPanel;
    G: THGrid;
    Label4: TLabel;
    Label6: TLabel;
    GFact: THGrid;
    TTC_ACC: THNumEdit;
    TTC_FACT: THNumEdit;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Label8: TLabel;
    COUV_ACC: THNumEdit;
    Label9: TLabel;
    COUV_FACT: THNumEdit;
    Label3: TLabel;
    Ratio: THNumEdit;
    Label10: TLabel;
    Label11: TLabel;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    BZoomFact: THBitBtn;
    BZoomLettre: THBitBtn;
    BImprimer: THBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BZoomFactClick(Sender: TObject);
    procedure BZoomLettreClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    Client : boolean ;
    RegimeTva : String3 ;
    TEcr   : TList ;
    NowFutur : TDateTime ;
    TabTva : Array[1..5] of Double ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure RempliEcran ;
    procedure RempliBases ;
    procedure CalculeLereste ;
    procedure ValideAcc ;
  public
    O  : TOBM ;
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  Saisie;


{$R *.DFM}

Procedure TvaAccLettrage ( O : TOBM ) ;
Var X : TFTvaLettr ;
BEGIN
X:=TFTvaLettr.Create(Application) ;
 Try
  X.O:=O ;
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFTvaLettr.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TEcr:=TList.Create ;
end;

procedure TFTvaLettr.CalculeLereste ;
Var i,j : integer ;
    OEcr : TOBM ;
    Okok : boolean ;
    XX,Signe,TotFact,TotCouv,Couv : Double ;
BEGIN
TotFact:=0 ; TotCouv:=0 ; Okok:=False ;
for i:=0 to TEcr.Count-1 do
    BEGIN
    OEcr:=TOBM(TEcr[i]) ; Signe:=1.0 ; Okok:=True ;
    GFact.Cells[0,GFact.RowCount-1]:=IntToStr(OEcr.GetMvt('E_NUMEROPIECE'))+' / '+IntToStr(OEcr.GetMvt('E_NUMLIGNE'))+' / '+IntToStr(OEcr.GetMvt('E_NUMECHE')) ;
    GFact.Cells[1,GFact.RowCount-1]:=OEcr.GetMvt('E_MODEPAIE') ;
    GFact.Cells[2,GFact.RowCount-1]:=DateToStr(OEcr.GetMvt('E_DATEECHEANCE')) ;
    if Client then
       BEGIN
       XX:=OEcr.GetMvt('E_DEBIT')-OEcr.GetMvt('E_CREDIT') ;
       if OEcr.GetMvt('E_DEBIT')=0 then Signe:=-1.0 ;
       END else
       BEGIN
       XX:=OEcr.GetMvt('E_CREDIT')-OEcr.GetMvt('E_DEBIT') ;
       if OEcr.GetMvt('E_CREDIT')=0 then Signe:=-1.0 ;
       END ;
    for j:=1 to 4 do TabTva[j]:=TabTva[j]+OEcr.GetMvt('E_ECHEENC'+IntToStr(j))(**Signe*) ;
    TabTva[5]:=TabTva[5]+OEcr.GetMvt('E_ECHEDEBIT') ;
    GFact.Cells[3,GFact.RowCount-1]:=StrS0(XX) ;
    Couv:=OEcr.GetMvt('E_COUVERTURE') ;
    TotFact:=TotFact+XX ;
    TotCouv:=TotCouv+(Signe*Couv) ;
    GFact.RowCount:=GFact.RowCount+1 ;
    END ;
TTC_Fact.Value:=TotFact ;
Couv_Fact.Value:=TotCouv ;
if TotCouv<>0 then Ratio.Value:=Couv_Acc.Value/TotCouv ;
if Okok then GFact.RowCount:=GFact.RowCount-1 ;
END ;

procedure TFTvaLettr.RempliBases ;
Var i : integer ;
BEGIN
for i:=1 to 5 do
    BEGIN
    G.Cells[0,i]:=HM.Mess[i-1] ;
    if i<5 then
       BEGIN
       G.Cells[1,i]:=VH^.NumCodeBase[i] ;
       G.Cells[2,i]:=StrS0(100.0*Tva2Taux(RegimeTva,VH^.NumCodeBase[i],Not Client)) ;
       G.Cells[4,i]:=StrS0(O.GetMvt('E_ECHEENC'+IntToStr(i))) ;
       END else
       BEGIN
       G.Cells[4,i]:=StrS0(O.GetMvt('E_ECHEDEBIT')) ;
       END ;
    G.Cells[3,i]:=StrS0(TabTva[i]) ;
    G.Cells[5,i]:=StrS0(TabTva[i]*Ratio.Value) ;
    G.Cells[6,i]:=StrS0(Valeur(G.Cells[5,i])-Valeur(G.Cells[4,i])) ;
    END ;
END ;

procedure TFTvaLettr.RempliEcran ;
Var MtAcc,CouvAcc : Double ;
    StNat,SQL     : String ;
    Q             : TQuery ;
    OEcr          : TOBM ;
BEGIN
CouvAcc:=O.GetMvt('E_COUVERTURE') ;
if Client then
   BEGIN
   MtAcc:=O.GetMvt('E_CREDIT')-O.GetMvt('E_DEBIT') ;
   StNat:='(E_NATUREPIECE="AC" OR E_NATUREPIECE="FC")' ;
   END else
   BEGIN
   MtAcc:=O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT') ;
   StNat:='(E_NATUREPIECE="AF" OR E_NATUREPIECE="FF")' ;
   END ;
Couv_Acc.Value:=CouvAcc ;
TTC_Acc.Value:=MtAcc ;
SQL:='Select * from Ecriture Where E_AUXILIAIRE="'+O.GetMvt('E_AUXILIAIRE')+'" AND E_GENERAL="'+O.GetMvt('E_GENERAL')+'" AND E_ETATLETTRAGE="'+O.GetMvt('E_ETATLETTRAGE')+'"'
    +' AND E_LETTRAGE="'+O.GetMvt('E_LETTRAGE')+'" AND '+StNat ;
Q:=OpenSQL(SQL,True) ;
While Not Q.EOF do
   BEGIN
   OEcr:=TOBM.Create(EcrGen,'',False) ;
   OEcr.ChargeMvt(Q) ; TEcr.Add(OEcr) ;
   Q.Next ;
   END ;
Ferme(Q) ;
CalculeLeReste ;
RempliBases ;
END ;

procedure TFTvaLettr.FormShow(Sender: TObject);
begin
G.GetCellCanvas:=GetCellCanvas ;
ChangeFormatDevise(Self,V_PGI.OkDecV,'') ;
Client:=(O.GetMvt('E_NATUREPIECE')='OC') ;
RegimeTva:=O.GetMvt('E_REGIMETVA') ;
FillChar(TabTva,Sizeof(TabTva),#0) ;
G.ColAligns[1]:=taCenter ; G.ColAligns[2]:=taCenter ;
G.ColAligns[3]:=taRightJustify ; G.ColAligns[4]:=taRightJustify ;
G.ColAligns[5]:=taRightJustify ; G.ColAligns[6]:=taRightJustify ;
GFact.ColAligns[2]:=taCenter ; GFact.ColAligns[3]:=taRightJustify ;
RempliEcran ;
if TEcr.Count<=0 then BValide.Enabled:=False ;
end;

procedure TFTvaLettr.FormClose(Sender: TObject; var Action: TCloseAction);
begin
VideListe(TEcr) ; TEcr.Free ;
GFact.VidePile(True) ;
G.VidePile(True) ;
end;

procedure TFTvaLettr.BZoomFactClick(Sender: TObject);
Var M : RMVT ;
    OBM : TOBM ;
begin
if TEcr.Count<=0 then Exit ;
OBM:=TOBM(TEcr[GFact.Row-1]) ; if OBM=Nil then Exit ;
M:=OBMToIdent(OBM,False) ;
LanceSaisie(Nil,taConsult,M) ;
end;

procedure TFTvaLettr.BZoomLettreClick(Sender: TObject);
Var R : RLETTR ;
begin
FillChar(R,Sizeof(R),#0) ;
R.General:=O.GetMvt('E_GENERAL') ;
R.Auxiliaire:=O.GetMvt('E_AUXILIAIRE') ;
R.CritDev:=O.GetMvt('E_DEVISE') ; R.DeviseMvt:=R.CritDev ;
R.GL:=NIL ; R.CritMvt:='' ; R.Appel:=tlMenu ;
R.CodeLettre:=O.GetMvt('E_LETTRAGE') ;
R.DeviseMvt:=R.CritDev ;
R.LettrageDevise:=(O.GetMvt('E_LETTRAGEDEV')='X') ;
LettrageManuel(R,False,taConsult) ;
end;

procedure TFTvaLettr.ValideAcc ;
Var M   : RMVT ;
    Nb  : integer ;
    SQL : String ; 
BEGIN
M:=OBMToIdent(O,True) ; NowFutur:=NowH ;
SQL:='UPDATE ECRITURE SET E_ECHEENC1=0, E_ECHEENC2=0, E_ECHEENC3=0, E_ECHEENC4=0, E_ECHEDEBIT=0, E_DATEMODIF="'+USTIME(NowFutur)+'" WHERE ' ;
SQL:=SQL+WhereEcriture(tsGene,M,True)+' AND E_DATEMODIF="'+USTIME(O.GetMvt('E_DATEMODIF'))+'"' ;
Nb:=ExecuteSQL(SQL) ; if Nb<>1 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TFTvaLettr.BValideClick(Sender: TObject);
Var io : TIoErr ;
begin
if HM.Execute(5,caption,'')<>mrYes then Exit ;
io:=Transactions(ValideAcc,3) ;
if io<>oeOk then MessageAlerte(HM.Mess[6]) else Close ; 
end;

Procedure TFTvaLettr.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
Var X : Double ;
BEGIN
if ((ARow<=0) or (ACol<>G.ColCount-1)) then Exit ;
X:=Valeur(G.Cells[ACol,ARow]) ; if X=0 then Exit ;
if X<0 then G.Canvas.Font.Color:=clRed else
 if X>0 then G.Canvas.Font.Color:=clGreen ; 
END ;


procedure TFTvaLettr.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
// A FAIRE Voir QRTVAACC.PAS
{$ELSE}
//  EditionTvaAccZoom(O) ;
{$ENDIF}
end;

procedure TFTvaLettr.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
