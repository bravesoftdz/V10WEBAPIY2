unit TvaContr;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, Grids, DBGrids, StdCtrls, Hctrls,
  ExtCtrls, ComCtrls, Buttons, Ent1, HEnt1, HCompte, Mask, SaisUtil,
  hmsgbox, Saisie, SaisComm, HStatus,
{$IFNDEF ESP}
  TVA,
{$ELSE}
  CPTVATPF_TOF,
{$ENDIF ESP}//XMG 19/01/04
  Filtre, HRichEdt, HSysMenu, HDB,
  HTB97, ed_tools, ColMemo, HPanel, UiUtil, HRichOLE, ADODB ;

procedure ControleTVAFactures ( EXO : String3 ; D1,D2 : TDateTime ; Etab : String ) ;
procedure RecupInfosTVA ( TEcr,TGen,TAux : TList ; Var RegTVA : String3 ; Var SoumisTPF : boolean ) ;
procedure ConstitueListesTVA ( M : RMVT ; TEcr,TGen,TAux : TList ; Var Nbl : integer ) ;
function  EstTiersGeneTVA ( CGen : TGGeneral ) : boolean ;
function  NatureHTTVA ( NatG : String ) : boolean ;

type X_TVA = Class
             CHT,CTVA,CTPF  : String17 ;
             RegTVA,CodeTva : String3 ;
             Achat          : boolean ;
             XHTP,XTVAAP,XTVADP,XTPFAP,XTPFDP : Double ;
             XHTD,XTVAAD,XTVADD,XTPFAD,XTPFDD : Double ;
             END ;

type
  TFTvaContr = class(TFMul)
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    TE_QUALIFPIECE: THLabel;
    E_QUALIFPIECE: THValComboBox;
    E_ETABLISSEMENT: THMultiValComboBox;
    HLabel2: THLabel;
    E_JOURNAL: THValComboBox;
    HLabel3: THLabel;
    E_NUMLIGNE: THCritMaskEdit;
    E_NUMECHE: THCritMaskEdit;
    HM: THMsgBox;
    BMenuZoom: TToolbarButton97;
    POPZ: TPopupMenu;
    E_CREERPAR: THCritMaskEdit;
    BParamTPF: TToolbarButton97;
    BParamTVA: TToolbarButton97;
    BZoomPiece: TToolbarButton97;
    XX_WHEREMODE: TEdit;
    TE_REFINTERNE: THLabel;
    E_REFINTERNE: TEdit;
    E_SAISIEEURO: TCheckBox;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    TE_GENERAL: THLabel;
    E_GENERAL: THCpteEdit;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    TOLERECALC: THCritMaskEdit;
    HTOLERECALC: THLabel;
    XX_WHERE: TEdit;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure E_EXERCICEChange(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure BChercheClick(Sender: TObject); override;
    procedure FListeDblClick(Sender: TObject); override;
    procedure BZoomPieceClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BParamTVAClick(Sender: TObject);
    procedure BParamTPFClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
  private
    NowFutur : TDateTime ;
    TT_TVA   : TList ;
    DEV      : RDEVISE ;
    ModeOppose : boolean ;  
    procedure ControleLaTva ;
    function  ControlePiece ( M : RMVT ; Var NbL : integer ) : String ;
    procedure MajEtat ( M : RMVT ; EtatTVA : String3 ; NbL : integer ) ;
// Conversions, calculs
    function  PasDeCodeTVA ( TGen,TEcr : TList ) : boolean ;
// Etude de la pièce
    function  EtudieTVA ( TEcr,TGen,TAux : TList ; Achat : boolean ) : String3 ;
    function  PasDeTiers ( TGen,TAux : TList ) : boolean ;
    function  PasDeHT ( TGen : TList ) : boolean ;
    function  MultiTiers ( TEcr,TGen,TAux : TList ) : boolean ;
    function  TiersSansRegime ( TGen,TAux : TList ) : boolean ;
    procedure RempliTVA ( TEcr,TGen : TList ; RegTVA : String3 ; Achat,SoumisTPF : boolean ) ;
    procedure CumulTVA ( XX : X_TVA ; O : TOBM ; CGen : TGGeneral ; RegTVA : String3 ; Achat,SoumisTPF : boolean ) ;
    function  SommeSais ( CPT : String ; TEcr : TList ; tsm : TSorteMontant ) : Double ;
    function  SommeCalc ( CPT : String ; OkTVA,Arr : Boolean ; tsm : TSorteMontant ) : Double ;
    function  PresentTVA ( CPT : String ; TEcr : TList ) : boolean ;
  public
    LeExo : String3 ;
    LaD1,LaD2 : TDateTime ;
    LesEtab : String ;
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  UtilPgi ;

procedure ControleTVAFactures ( EXO : String3 ; D1,D2 : TDateTime ; Etab : String ) ;
Var X : TFTvaContr ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture','nrBatch'],False,'nrBatch') then Exit ;
PP:=FindInsidePanel ;
X:=TFTvaContr.Create(Application) ;
X.LeExo:=EXO; X.LaD1:=D1 ; X.LaD2:=D2 ; X.LesEtab:=Etab ;
if PP=Nil then
   BEGIN
   Try
    X.ShowModal ;
   Finally
    X.Free ;
    _Bloqueur('nrBatch',False) ;
   End ;
  Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

{============================== Procédures communes ========================================}
function EstTiersGeneTVA ( CGen : TGGeneral ) : boolean ;
BEGIN
Result:=((CGen.NatureGene='TID') or (CGen.NatureGene='TIC')) ;
END ;

function NatureHTTVA ( NatG : String ) : boolean ;
BEGIN
Result:=((NatG='CHA') or (NatG='PRO') or (NatG='IMO')) ;
END ;

procedure ConstitueListesTVA ( M : RMVT ; TEcr,TGen,TAux : TList ; Var Nbl : integer ) ;
Var QEcr : TQuery ;
    O    : TOBM ;
    CGen : TGGeneral ;
    CAux : TGTiers ;
    CptGen,CptAux : String ;
BEGIN
Nbl:=0 ;
QEcr:=OpenSQL('Select * from Ecriture Where '+WhereEcriture(tsGene,M,False),True) ;
While Not QEcr.EOF do
   BEGIN
   O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(QEcr) ;
   CptGen:=O.GetMvt('E_GENERAL') ; CGen:=TGGeneral.Create(CptGen) ;
   CptAux:=O.GetMvt('E_AUXILIAIRE') ; if CptAux<>'' then CAux:=TGTiers.Create(CptAux) else CAux:=Nil ;
   inc(NbL) ; QEcr.Next ;
   TEcr.Add(O) ; TGen.Add(CGen) ; TAux.Add(CAux) ;
   END ;
Ferme(QEcr) ;
END ;

procedure RecupInfosTVA ( TEcr,TGen,TAux : TList ; Var RegTVA : String3 ; Var SoumisTPF : boolean ) ;
Var i : integer ;
    CGen : TGGeneral ;
    CAux : TGTiers ;
    O    : TOBM ;
BEGIN
for i:=0 to TEcr.Count-1 do
    BEGIN
    O:=TOBM(TEcr[i]) ; if O=Nil then Continue ;
    if O.GetMvt('E_NUMECHE')<=1 then
       BEGIN
       RegTva:=O.GetMvt('E_REGIMETVA') ;
       CGen:=TGGeneral(TGen[i]) ;
       if EstTiersGeneTVA(CGen) then
          BEGIN
          if RegTva='' then RegTVA:=CGen.RegimeTVA ;
          SoumisTPF:=CGen.SoumisTPF ;
          Break ;
          END else if TGTiers(TAux[i])<>Nil then
          BEGIN
          CAux:=TGTiers(TAux[i]) ;
          if RegTva='' then RegTVA:=CAux.RegimeTVA ;
          SoumisTPF:=CAux.SoumisTPF ;
          Break ;
          END ;
       END ;
    END ;
if RegTva='' then RegTva:=VH^.RegimeDefaut ;     
END ;

{============================== Procédures de la form ====================================}
procedure TFTvaContr.FormShow(Sender: TObject);
begin
if LeExo='' then
   BEGIN
   if VH^.CPExoRef.Code<>'' then
      BEGIN
      E_EXERCICE.Value:=VH^.CPExoRef.Code ;
      E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
      E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
      END else
      BEGIN
      E_EXERCICE.Value:=VH^.Entree.Code ;
      E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
      E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
      END ;
   END else
   BEGIN
   E_EXERCICE.Value:=LeExo ;
   E_DATECOMPTABLE.Text:=DateToStr(LaD1) ;
   E_DATECOMPTABLE_.Text:=DateToStr(LaD2) ;
   E_ETABLISSEMENT.Text:=LesEtab ;
   END ;
if E_JOURNAL.Values.Count>0 then E_JOURNAL.Value:=E_JOURNAL.Values[0] ;
  inherited;
Pages.Pages[2].TabVisible:=FALSE ; Pages.Pages[3].TabVisible:=FALSE ;
Q.Manuel:=FALSE ;
Q.UpdateCriteres ;
CentreDBGRid(FListe) ; GereSelectionsGrid(FListe,Q) ;
Q.Manuel := False ;
end;

procedure TFTvaContr.FormCreate(Sender: TObject);
begin
 Q.Manuel:=TRUE ;
  inherited;
TypeAction:=taModif ; FNomFiltre:='CONTROLETVA' ;
E_QUALIFPIECE.Value:='N' ;
TT_TVA:=TList.Create ; ModeOppose:=False ;
Q.Liste:='CONTROLETVAS' ;
end;

procedure TFTvaContr.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

{===================== CONTROLES =========================}
function TFTvaContr.PasDeHT ( TGen : TList ) : boolean ;
Var i : integer ;
BEGIN
PasDeHT:=False ;
for i:=0 to TGen.Count-1 do
//  if ((TGGeneral(TGen[i]).General<>VH^.EccEuroDebit) and (TGGeneral(TGen[i]).General<>VH^.EccEuroCredit)) then
    if NatureHTTVA(TGGeneral(TGen[i]).NatureGene) then Exit ;
PasDeHT:=True ;
END ;

function TFTvaContr.PasDeCodeTVA ( TGen,TEcr : TList ) : boolean ;
Var i    : integer ;
    CGen : TGGeneral ;
    OEcr : TOBM ;
BEGIN
PasDeCodeTVA:=True ;
for i:=0 to TGen.Count-1 do
    BEGIN
    CGen:=TGGeneral(TGen[i]) ;
//    if ((TGGeneral(TGen[i]).General<>VH^.EccEuroDebit) and (TGGeneral(TGen[i]).General<>VH^.EccEuroCredit)) then
       if NatureHTTVA(CGen.NatureGene) then
          BEGIN
          OEcr:=TOBM(TEcr[i]) ;
          if OEcr.GetMvt('E_TVA')='' then Exit ;
          END ;
    END ;
PasDeCodeTVA:=False ;
END ;

function TFTvaContr.PasDeTiers ( TGen,TAux : TList ) : boolean ;
Var i : integer ;
BEGIN
PasDeTiers:=False ;
for i:=0 to TGen.Count-1 do if EstTiersGeneTVA(TGGeneral(TGen[i])) then Exit ;
for i:=0 to TAux.Count-1 do if TGTiers(TAux[i])<>Nil then Exit ;
PasDeTiers:=True ;
END ;

function TFTvaContr.TiersSansRegime ( TGen,TAux : TList ) : boolean ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    i    : integer ;
BEGIN
TiersSansRegime:=True ;
for i:=0 to TGen.Count-1 do
    BEGIN
    CGen:=TGGeneral(TGen[i]) ;
    if EstTiersGeneTVA(CGen) then if CGen.RegimeTVA='' then Exit ;
    END ;
for i:=0 to TAux.Count-1 do
    BEGIN
    CAux:=TGTiers(TAux[i]) ;
    if CAux<>Nil then if CAux.RegimeTVA='' then Exit ;
    END ;
TiersSansRegime:=False ;
END ;

function TFTvaContr.MultiTiers ( TEcr,TGen,TAux : TList ) : boolean ;
Var i,Nb : integer ;
BEGIN
Nb:=0 ;
for i:=0 to TEcr.Count-1 do if TOBM(TEcr[i]).GetMvt('E_NUMECHE')<=1 then
    BEGIN
    if EstTiersGeneTVA(TGGeneral(TGen[i])) then Inc(Nb) ;
    if TGTiers(TAux[i])<>Nil then Inc(Nb) ;
    END ;
MultiTiers:=(Nb>1) ;
END ;

procedure TFTvaContr.CumulTVA ( XX : X_TVA ; O : TOBM ; CGen : TGGeneral ; RegTVA : String3 ; Achat,SoumisTPF : boolean ) ;
Var CodeTva,CodeTpf : String ;
BEGIN
XX.CHT:=O.GetMvt('E_GENERAL') ;
// HT
XX.XHTP:=XX.XHTP+O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT') ;
XX.XHTD:=XX.XHTD+O.GetMvt('E_DEBITDEV')-O.GetMvt('E_CREDITDEV') ;
// TVA
CodeTva:=O.GetMvt('E_TVA') ; CodeTpf:=O.GetMvt('E_TPF') ;
XX.CTVA:=CGen.CpteTVA ;
XX.XTVADP:=HT2TVA(XX.XHTP,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5) ; XX.XTVAAP:=Arrondi(XX.XTVADP,V_PGI.OkDecV) ;
XX.XTVADD:=HT2TVA(XX.XHTD,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5) ; XX.XTVAAD:=Arrondi(XX.XTVADD,DEV.Decimale) ;
// TPF
XX.CTPF:=CGen.CpteTPF ;
XX.XTPFDP:=HT2TPF(XX.XHTP,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5) ; XX.XTPFAP:=Arrondi(XX.XTPFDP,V_PGI.OkDecV) ;
XX.XTPFDD:=HT2TPF(XX.XHTD,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5) ; XX.XTPFAD:=Arrondi(XX.XTPFDD,DEV.Decimale) ;
END ;

procedure TFTvaContr.RempliTVA ( TEcr,TGen : TList ; RegTVA : String3 ; Achat,SoumisTPF : boolean ) ;
Var O   : TOBM ;
    i,k : integer ;
    XX  : X_TVA ;
    CGen : TGGeneral ;
    Trouv : boolean ;
BEGIN
VideListe(TT_TVA) ;
for i:=0 to TEcr.Count-1 do
    BEGIN
    CGen:=TGGeneral(TGen[i]) ; O:=TOBM(TEcr[i]) ;
//    if ((TGGeneral(TGen[i]).General<>VH^.EccEuroDebit) and (TGGeneral(TGen[i]).General<>VH^.EccEuroCredit)) then
       if NatureHTTVA(CGen.NatureGene) then
       BEGIN
       Trouv:=False ;
       for k:=0 to TT_TVA.Count-1 do
           BEGIN
           XX:=X_TVA(TT_TVA[k]) ;
           if XX.CHT=CGen.General then BEGIN Trouv:=True ; CumulTVA(XX,O,CGen,RegTVA,Achat,SoumisTPF) ; Break ; END ;
           END ;
       if Not Trouv then
          BEGIN
          XX:=X_TVA.Create ;
          XX.RegTVA:=RegTVA ; XX.Achat:=Achat ; XX.CodeTva:=O.GetMvt('E_TVA') ;
          XX.XHTP:=0 ; XX.XTVADP:=0 ; XX.XTVAAP:=0 ; XX.XTPFDP:=0 ; XX.XTPFAP:=0 ;
          XX.XHTD:=0 ; XX.XTVADD:=0 ; XX.XTVAAD:=0 ; XX.XTPFDD:=0 ; XX.XTPFAD:=0 ;
          CumulTVA(XX,O,CGen,RegTVA,Achat,SoumisTPF) ;
          TT_TVA.Add(XX) ;
          END ;
       END ;
    END ;
END ;

function TFTvaContr.SommeCalc ( CPT : String ; OkTVA,Arr : Boolean ; tsm : TSorteMontant ) : Double ;
Var d : double ;
    XX : X_TVA ;
    i  : integer ;
BEGIN
D:=0 ;
for i:=0 to TT_TVA.Count-1 do
    BEGIN
    XX:=X_TVA(TT_TVA[i]) ;
    if ((OkTVa) and (XX.CTVA=CPT)) then
       BEGIN
       Case tsm of
          tsmDevise : if Arr then D:=D+XX.XTVAAD else D:=D+XX.XTVADD ;
          tsmPivot  : if Arr then D:=D+XX.XTVAAP else D:=D+XX.XTVADP ;
          END ;
       END ;
    if ((Not OkTVa) and (XX.CTPF=CPT)) then
       BEGIN
       Case tsm of
          tsmDevise : if Arr then D:=D+XX.XTPFAD else D:=D+XX.XTPFDD ;
          tsmPivot  : if Arr then D:=D+XX.XTPFAP else D:=D+XX.XTPFDP ;
          END ;
       END ;
    END ;
Result:=D ;
END ;

function TFTvaContr.SommeSais ( CPT : String ; TEcr : TList ; tsm : TSorteMontant ) : Double ;
Var d : double ;
    O : TOBM ;
    i  : integer ;
BEGIN
D:=0 ;
for i:=0 to TEcr.Count-1 do
    BEGIN
    O:=TOBM(TEcr[i]) ;
    if O.GetMvt('E_GENERAL')=CPT then
       BEGIN
       Case tsm of
          tsmDevise : D:=D+O.GetMvt('E_DEBITDEV')-O.GetMvt('E_CREDITDEV') ;
          tsmPivot  : D:=D+O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT') ;
          tsmEuro   : D:=D+O.GetMvt('E_DEBITEURO')-O.GetMvt('E_CREDITEURO') ;
          END ;
       END ;
    END ;
Result:=D ;
END ;

function TFTvaContr.PresentTVA ( CPT : String ; TEcr : TList ) : boolean ;
Var i : integer ;
BEGIN
Result:=True ;
for i:=0 to TEcr.Count-1 do if TOBM(TEcr[i]).GetMvt('E_GENERAL')=CPT then Exit ;
Result:=False ;
END ;

function TFTvaContr.EtudieTVA ( TEcr,TGen,TAux : TList ; Achat : boolean ) : String3 ;
Var RegTVA,CodeTva,CodeTpf : String3 ;
    SoumisTPF,CoherTVA,EncON : boolean ;
    i                  : integer ;
    CGen               : TGGeneral ;
    XX                 : X_TVA ;
    OHT                : TOBM ;
    Diff,STaxeA,STaxeD,Tolere : Double ;
    tsm                : TSorteMontant ;
    Ladec              : integer ;
  szProrata          : String;
BEGIN
Result:='COR' ; CoherTVA:=True ; Tolere:=Valeur(TolereCalc.Text) ;
{Tests premier niveau}
if PasDeTiers(TGen,Taux) then BEGIN Result:='STI' ; Exit ; END ;
if MultiTiers(TEcr,TGen,Taux) then BEGIN Result:='MTI' ; Exit ; END ;
if PasDeHT(TGen) then BEGIN Result:='SHT' ; Exit ; END ;
if PasDeCodeTVA(TGen,TEcr) then BEGIN Result:='STV' ; Exit ; END ;
if TiersSansRegime(TGen,TAux) then BEGIN Result:='RTI' ; Exit ; END ;
{Cohérence de la TVA}
RecupInfosTVA(TEcr,TGen,TAux,RegTVA,SoumisTPF) ;
for i:=0 to TGen.Count-1 do
    BEGIN
    CGen:=TGGeneral(TGen[i]) ;
//    if ((CGen.General<>VH^.EccEuroDebit) and (CGen.General<>VH^.EccEuroCredit)) then
       if NatureHTTVA(CGen.NatureGene) then
       BEGIN
       OHT:=TOBM(TEcr[i]) ;
       CodeTva:=OHT.GetMvt('E_TVA') ; CodeTpf:=OHT.GetMvt('E_TPF') ;
       EncON:=(OHT.GetMvt('E_TVAENCAISSEMENT')='X') ;
       RenseigneHTByTva(CGen,RegTVA,CodeTva,CodeTpf,SoumisTPF,Achat,EncON,CoherTva) ;
       END ;
    if Not CoherTVA then BEGIN Result:='INC' ; Exit ; END ;
    END ;
{Calculs}
RempliTVA(TEcr,TGen,RegTVA,Achat,SoumisTPF) ;
for i:=0 to TT_TVA.Count-1 do
    BEGIN
    XX:=X_TVA(TT_TVA[i]) ;
    // Tests de cohérence su monnaie Pivot
    if ((XX.CHT<>'') and (XX.XHTP<>0) and (XX.XTVAAP<>0) and (XX.CTVA='')) then BEGIN Result:='STV' ; Exit ; END ;
    if ((XX.CHT<>'') and (XX.XHTP<>0) and (XX.XTPFAP<>0) and (XX.CTPF='')) then BEGIN Result:='STV' ; Exit ; END ;
    if ((XX.XTVAAP<>0) and (XX.CTVA<>'')) then if Not PresentTVA(XX.CTVA,TEcr) then BEGIN Result:='ERR' ; Exit ; END ;
    if ((XX.XTPFAP<>0) and (XX.CTPF<>'')) then if Not PresentTVA(XX.CTPF,TEcr) then BEGIN Result:='ERR' ; Exit ; END ;
    if ((XX.CHT<>'') and (XX.XHTP<>0) and (XX.CodeTva<>'') and (Tva2Taux(XX.RegTVA,XX.CodeTva,XX.Achat)=0)) then BEGIN Result:='EXO' ; Exit ; END ;
    // Tests de montant sur monnaie de saisie
    if DEV.Code<>V_PGI.DevisePivot then BEGIN LaDec:=DEV.Decimale ; tsm:=tsmDevise ; END else
     if ModeOppose then BEGIN LaDec:=V_PGI.OkDecE ; tsm:=tsmEuro ; END else
        BEGIN LaDec:=V_PGI.OkDecV ; tsm:=tsmPivot ; END ;
    STaxeA:=SommeCalc(XX.CTVA,True,True,tsm) ; STaxeD:=SommeCalc(XX.CTVA,True,False,tsm) ;
    Diff:=Abs(STaxeA-STaxeD) ; if Arrondi(Diff,LaDec)>Diff then Diff:=Arrondi(Diff,LaDec) ;
    if Abs(Arrondi(SommeCalc(XX.CTVA,True,False,tsm)-SommeSais(XX.CTVA,TEcr,tsm),LaDec))>Diff+Tolere then BEGIN
      szProrata := TOBM(TEcr[i]).GetMvt('E_QUALIFORIGINE');
      {JP 06/06/03 : correction de l'ordre des conditions}
      if (szProrata = 'TV1') then begin Result:='PR1'; Exit; end else
      if (szProrata = 'TVI') then begin Result:='PR2'; Exit; end else
      if (szProrata = 'TV2') then begin Result:='PR3'; Exit; end
         else begin
              Result:='MTV'; Exit; end ;
    end;
    STaxeA:=SommeCalc(XX.CTPF,False,True,tsm) ; STaxeD:=SommeCalc(XX.CTPF,False,False,tsm) ;
    Diff:=Abs(STaxeA-STaxeD) ; if Arrondi(Diff,LaDec)>Diff then Diff:=Arrondi(Diff,LaDec) ;
    if Abs(Arrondi(SommeCalc(XX.CTPF,False,False,tsm)-SommeSais(XX.CTPF,TEcr,tsm),LaDec))>Diff+Tolere then BEGIN Result:='MTV' ; Exit ; END ;
    END ;
END ;

function TFTvaContr.ControlePiece ( M : RMVT ; Var NbL : integer ) : String ;
Var QJal : TQuery ;
    TEcr,TGen,TAux : TList ;
    Etat      : String3 ;
    Achat     : boolean ;
BEGIN
DEV.Code:=M.CodeD ; GetInfosDevise(DEV) ; 
Etat:='COR' ; TEcr:=TList.Create ; TGen:=TList.Create ; TAux:=TList.Create ;
QJal:=OpenSQL('Select J_NATUREJAL from JOURNAL Where J_JOURNAL="'+M.Jal+'"',True) ;
if Not QJal.EOF then Achat:=(QJal.Fields[0].AsString='ACH') else Achat:=False ;
Ferme(QJal) ;
ConstitueListesTVA(M,TEcr,TGen,TAux,Nbl) ;
Etat:=EtudieTVA(TEcr,TGen,TAux,Achat) ;
TEcr.Free ; TGen.Free ; TAux.Free ;
Result:=Etat ;
END ;

{========================= VALIDATIONS, MISES A JOUR ==============================}
procedure TFTvaContr.MajEtat ( M : RMVT ; EtatTVA : String3 ; NbL : integer ) ;
Var SQL : String ;
    NbE : integer ;
BEGIN
SQL:='Update Ecriture Set E_CONTROLETVA="'+EtatTVA+'", E_DATEMODIF="'+UsTime(NowFutur)+'" '
    +'Where '+WhereEcriture(tsGene,M,False) ;
NbE:=ExecuteSQL(SQL) ;
if NbE<>NbL then V_PGI.IOError:=oeUnknown ;
END ;

procedure TFTvaContr.ControleLaTva ;
Var i,Nbl : integer ;
    M : RMVT ;
    EtatTVA : String3 ;
BEGIN
if FListe.AllSelected then
   BEGIN
   InitMove(100,'') ;
   Q.First ;
   While Not Q.EOF do
      BEGIN
      MoveCur(False) ;
      if TrouveSaisie(Q,M,E_QUALIFPIECE.Value) then
         BEGIN
         EtatTVA:=ControlePiece(M,NbL) ;
         MajEtat(M,EtatTVA,NbL) ;
         END ;
      if V_PGI.IOError<>oeOk then Break ;
      Q.Next ;
      END ;
   END else
   BEGIN
   InitMove(FListe.NbSelected,'') ;
   for i:=0 to FListe.NbSelected-1 do
       BEGIN
       MoveCur(False) ; FListe.GotoLeBookMark(i) ;
       if TrouveSaisie(Q,M,E_QUALIFPIECE.Value) then
          BEGIN
          EtatTVA:=ControlePiece(M,NbL) ;
          MajEtat(M,EtatTVA,NbL) ;
          END ;
       if V_PGI.IOError<>oeOk then Break ;
       END ;
   END ;
FiniMove ;
END ;

procedure TFTvaContr.BOuvrirClick(Sender: TObject);
Var Nb : longint ;
    ii : integer ;
    io : TIOErr ;
begin
  inherited;
if Not FListe.AllSelected then
   BEGIN
   Nb:=FListe.NbSelected ; if Nb<=0 then BEGIN HM.Execute(1,'','') ; Exit ; END ;
   END ;
ii:=HM.Execute(0,'','') ; Application.ProcessMessages ;
Case ii of
   mrYes : BEGIN
           NowFutur:=NowH ; io:=Transactions(ControleLaTva,3) ;
           if io<>oeOk then MessageAlerte(HM.Mess[2]) ;
           BChercheClick(Nil) ;
           END ;
   mrNo  : FListe.ClearSelected ;
   mrCancel : ;
   END ;
end;

procedure TFTvaContr.BChercheClick(Sender: TObject);
begin
  inherited;
GereSelectionsGrid(FListe,Q) ;
end;

procedure TFTvaContr.FListeDblClick(Sender: TObject);
begin
  inherited;
TrouveEtLanceSaisie(Q,taModif,E_QUALIFPIECE.Value) ;
end;

procedure TFTvaContr.BZoomPieceClick(Sender: TObject);
begin
  inherited;
FListeDBlClick(Nil) ;
end;

procedure TFTvaContr.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
VideListe(TT_TVA) ; TT_TVA.Free ;
if IsInside(Self) then _Bloqueur('nrBatch',False) ;
end;

procedure TFTvaContr.BParamTVAClick(Sender: TObject);
begin
  inherited;
ParamTVATPF(True) ;
end;

procedure TFTvaContr.BParamTPFClick(Sender: TObject);
begin
  inherited;
ParamTVATPF(False) ;
end;

procedure TFTvaContr.BMenuZoomMouseEnter(Sender: TObject);
begin
  inherited;
PopZoom97(BMenuZoom,POPZ) ;
end;

end.

