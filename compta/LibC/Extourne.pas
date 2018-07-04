unit Extourne;
                              
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, ComCtrls, HRichEdt, Grids,
  DBGrids, StdCtrls, Hctrls, ExtCtrls, Buttons, Mask, Ent1, HEnt1, SaisComm, SaisUtil,
  Saisie, hmsgbox, HStatus, LetBatch, DelVisuE, ValPerio, HDB, HTB97, ed_tools,
  ColMemo, HPanel, UiUtil,
 {$IFNDEF CCS3}
  TiersPayeur,
 {$ENDIF}
  UtilSoc, HRichOLE, Hcompte, ADODB, TimpFic ;

procedure LanceExtourne ;

type
  TFExtourne = class(TFMul)
    Bevel5: TBevel;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    E_JOURNAL: THValComboBox;
    TE_JOURNAL: THLabel;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_NUMECHE: THCritMaskEdit;
    E_ECRANOUVEAU: THCritMaskEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHERENAT: TEdit;
    HLabel2: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    TE_DEVISE: THLabel;
    E_DEVISE: THValComboBox;
    E_NUMLIGNE: THCritMaskEdit;
    E_CREERPAR: THCritMaskEdit;
    TE_DATEECHEANCE: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    TE_DATEECHEANCE2: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    TE_QUALIFPIECE: THLabel;
    E_QUALIFPIECE: THValComboBox;
    Label2: TLabel;
    HLabel6: THLabel;
    CONTRE_DATE: THCritMaskEdit;
    CONTRE_NEGATIF: TCheckBox;
    BListePIECES: TToolbarButton97;
    BZoom: TToolbarButton97;
    HM: THMsgBox;
    HLabel3: THLabel;
    CONTRE_TYPE: THValComboBox;
    TEcrGen: TQuery;
    TEcrAna: TQuery;
    XX_WHERESAIS: TEdit;
    PComplements: TTabSheet;
    TE_REFINTERNE: THLabel;
    E_REFINTERNE: TEdit;
    TE_GENERAL: THLabel;
    E_GENERAL: THCpteEdit;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    E_VALIDE: TCheckBox;
    CDejaExt: TCheckBox;
    XX_WHEREEXT: TEdit;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure BListePIECESClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BOuvrirClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure CDejaExtClick(Sender: TObject);
  private
    ListeEcr,ListeAna,TPIECE : TList ;
    GeneDate,NowFutur : TDateTime ;
    NewNum            : Longint ;
    InfoImp           : TInfoImport ;
    QFiche            : TQFiche ;
    function  DateOK  : boolean ;
    function  JePeuxValider : boolean ;
    procedure ChargeLaPiece ( M : RMVT ) ;
    procedure NewPieceEcr ;
    procedure TripoteO ( O : TOBM ; ts : TTypeSais ) ;
    procedure GenereLesPieces ;
    Procedure TraiteTPVal ;
  public
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcGen,
  {$ENDIF MODENT1}
  UtilPgi ;

procedure LanceExtourne ;
Var X  : TFExtourne ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
if _Blocage(['nrBatch','nrCloture'],True,'nrBatch') then Exit ;
X:=TFExtourne.Create(Application) ;
X.FNomFiltre:='EXTOURNE' ;
PP:=FindInsidePanel ;
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

procedure TFExtourne.FormShow(Sender: TObject);
begin
E_DEVISE.Value:=V_PGI.DevisePivot ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   E_EXERCICE.Value:=VH^.CPExoRef.Code ; 
   E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ; E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   E_EXERCICE.Value:=VH^.Entree.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ; E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
E_QUALIFPIECE.Value:='N' ; CONTRE_Type.Value:='N' ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
CONTRE_DATE.Text:=DateToStr(V_PGI.DateEntree) ;
if E_JOURNAL.Items.Count>0 then
   BEGIN
   if E_JOURNAL.Vide then E_JOURNAL.ItemIndex:=1 else E_JOURNAL.ItemIndex:=0 ;
   END ;
PositionneEtabUser(E_ETABLISSEMENT) ;
  inherited;
Pages.Pages[3].TabVisible:=FALSE ;
if Not VH^.MontantNegatif then BEGIN CONTRE_NEGATIF.Checked:=False ; CONTRE_NEGATIF.Enabled:=False ; END ;
Q.Manuel:=FALSE ;
CentreDBGRid(FListe) ;
end;

procedure TFExtourne.FormCreate(Sender: TObject);
begin
  inherited;
ListeEcr:=TList.Create ; ListeAna:=TList.Create ; TPIECE:=TList.Create ;
Q.Manuel:=True ;
Q.Liste:='EXTOURNE' ;
ChangeSQL(TEcrGen) ; ChangeSQL(TEcrAna) ;
end;

procedure TFExtourne.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFExtourne.BListePIECESClick(Sender: TObject);
begin
  inherited;
if TPiece.Count>0 then VisuPiecesGenere(TPiece,EcrGen,10) ;
end;

procedure TFExtourne.BZoomClick(Sender: TObject);
begin
  inherited;
FListeDBLClick(Nil) ;
end;

procedure TFExtourne.FListeDblClick(Sender: TObject);
begin
  inherited;
TrouveEtLanceSaisie(Q,taConsult,E_QUALIFPIECE.Value) ;
end;

procedure TFExtourne.BChercheClick(Sender: TObject);
begin
  inherited;
GereSelectionsGrid(FListe,Q) ;
end;

procedure TFExtourne.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
VideListe(ListeEcr) ; ListeEcr.Free ;
VideListe(ListeAna) ; ListeAna.Free ;
VideListe(TPIECE) ; TPIECE.Free ;
if Parent is THPanel then _Bloqueur('nrBatch',False) ;
end;

function TFExtourne.DateOK : boolean ;
Var Err : integer ;
BEGIN
Result:=False ;
GeneDate:=0 ;
Err:=ControleDate(CONTRE_DATE.Text) ;
if Err>0 then BEGIN HM.Execute(1+Err,'','') ; Exit ; END ;
GeneDate:=StrToDate(CONTRE_DATE.Text) ;
Result:=True ;
END ;

function TFExtourne.JePeuxValider : boolean ;
BEGIN
Result:=False ;
if ((Not FListe.AllSelected) and (FListe.NbSelected<=0)) then BEGIN HM.Execute(1,'','') ; Exit ; END ;
if ((FListe.AllSelected) and (Q.EOF) and (Q.BOF)) then BEGIN HM.Execute(1,'','') ; Exit ; END ;
if ((E_QUALIFPIECE.Value<>'N') and (CONTRE_TYPE.Value='N')) then BEGIN HM.Execute(8,'',E_QUALIFPIECE.Text) ; Exit ; END ;
if Not DATEOK then Exit ;
Result:=True ;
END ;

procedure TFExtourne.ChargeLaPiece ( M : RMVT ) ;
Var QEcr,QAna : TQuery ;
    O    : TOBM ;
BEGIN
VideListe(ListeEcr) ; VideListe(ListeAna) ;
QEcr:=OpenSQL('Select * from Ecriture Where '+WhereEcriture(tsGene,M,False),False) ;
While Not QEcr.EOF do
   BEGIN
   O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(QEcr) ;
   ListeEcr.Add(O) ;

   {Marquer les pièces comme étant extounées}
   QEcr.Edit ;
   QEcr.FindField('E_QUALIFORIGINE').AsString:='URN' ;
   QEcr.Post ;

   QEcr.Next ;
   END ;
Ferme(QEcr) ;
QAna:=OpenSQL('Select * from Analytiq Where '+WhereEcriture(tsAnal,M,False),True) ;
While Not QAna.EOF do
   BEGIN
   O:=TOBM.Create(EcrAna,'',False) ; O.ChargeMvt(QAna) ;
   ListeAna.Add(O) ;
   QAna.Next ;
   END ;
Ferme(QAna) ;
END ;

procedure TFExtourne.TripoteO ( O : TOBM ; ts : TTypeSais ) ;
Var EL : String ;
    X  : double ;
BEGIN
if O=Nil then Exit ;
if ts=tsGene then
   BEGIN
   O.PutMvt('E_DATECOMPTABLE',GeneDate) ; O.PutMvt('E_DATEECHEANCE',GeneDate) ;
{$IFNDEF SPEC302}
   O.PutMvt('E_PERIODE',GetPeriode(GeneDate)) ; O.PutMvt('E_SEMAINE',NumSemaine(GeneDate)) ;
{$ENDIF}
   O.PutMvt('E_EXERCICE',QuelExoDT(GeneDate)) ;
   O.PutMvt('E_QUALIFPIECE',CONTRE_TYPE.Value) ; O.PutMvt('E_NUMEROPIECE',NewNum) ;
   O.PutMvt('E_VALIDE','-') ; O.PutMvt('E_LETTRAGE','') ;
   O.PutMvt('E_TRACE','')   ; O.PutMvt('E_QUALIFORIGINE','EXT') ;
   {Etat lettrage}
   EL:=O.GetMvt('E_ETATLETTRAGE') ;
   if ((EL='TL') or (EL='PL') or (EL='AL')) then O.PutMvt('E_ETATLETTRAGE','AL')
                                            else O.PutMvt('E_ETATLETTRAGE','RI') ;
   O.PutMvt('E_LETTRAGEDEV','-') ; 
   O.PutMvt('E_COUVERTURE',0) ; O.PutMvt('E_COUVERTUREDEV',0) ; 
   O.PutMvt('E_DATEPAQUETMIN',GeneDate) ; O.PutMvt('E_DATEPAQUETMAX',GeneDate) ;
   O.PutMvt('E_REFPOINTAGE','') ; O.PutMvt('E_DATEPOINTAGE',IDate1900) ;
   O.PutMvt('E_REFRELEVE','') ; O.PutMvt('E_FLAGECR','') ; O.PutMvt('E_ETAT','0000000000') ;
   O.PutMvt('E_NIVEAURELANCE',0) ; O.PutMvt('E_DATERELANCE',IDate1900) ;
   O.PutMvt('E_DATECREATION',Date) ; O.PutMvt('E_DATEMODIF',NowFutur) ;
   O.PutMvt('E_SUIVDEC','') ; O.PutMvt('E_NOMLOT','') ; O.PutMvt('E_EDITEETATTVA','-') ;
   // GG COM
   O.PutMvt('E_IO','X') ;
   if Contre_Negatif.Checked then
      BEGIN
      O.PutMvt('E_DEBIT',-O.GetMvt('E_DEBIT')) ;
      O.PutMvt('E_CREDIT',-O.GetMvt('E_CREDIT')) ;
      O.PutMvt('E_DEBITDEV',-O.GetMvt('E_DEBITDEV')) ;
      O.PutMvt('E_CREDITDEV',-O.GetMvt('E_CREDITDEV')) ;
      O.PutMvt('E_QTE1',-O.GetMvt('E_QTE1')) ;
      O.PutMvt('E_QTE2',-O.GetMvt('E_QTE2')) ;
      END else
      BEGIN
      X:=O.GetMvt('E_DEBIT') ; O.PutMvt('E_DEBIT',O.GetMvt('E_CREDIT')) ; O.PutMvt('E_CREDIT',X) ;
      X:=O.GetMvt('E_DEBITDEV') ; O.PutMvt('E_DEBITDEV',O.GetMvt('E_CREDITDEV')) ; O.PutMvt('E_CREDITDEV',X) ;
      END ;
   END else
   BEGIN
   O.PutMvt('Y_DATECOMPTABLE',GeneDate) ;
{$IFNDEF SPEC302}
   O.PutMvt('Y_PERIODE',GetPeriode(GeneDate)) ; O.PutMvt('Y_SEMAINE',NumSemaine(GeneDate)) ;
{$ENDIF}
   O.PutMvt('Y_EXERCICE',QuelExoDT(GeneDate)) ;
   O.PutMvt('Y_QUALIFPIECE',CONTRE_TYPE.Value) ;
   O.PutMvt('Y_NUMEROPIECE',NewNum) ; O.PutMvt('Y_TRACE','') ; O.PutMvt('Y_VALIDE','-') ;
   O.PutMvt('Y_DATECREATION',Date) ; O.PutMvt('Y_DATEMODIF',NowFutur) ;
   if Contre_Negatif.Checked then
      BEGIN
      O.PutMvt('Y_DEBIT',-O.GetMvt('Y_DEBIT')) ;
      O.PutMvt('Y_CREDIT',-O.GetMvt('Y_CREDIT')) ;
      O.PutMvt('Y_DEBITDEV',-O.GetMvt('Y_DEBITDEV')) ;
      O.PutMvt('Y_CREDITDEV',-O.GetMvt('Y_CREDITDEV')) ;
      O.PutMvt('Y_QTE1',-O.GetMvt('Y_QTE1')) ;
      O.PutMvt('Y_QTE2',-O.GetMvt('Y_QTE2')) ;
      O.PutMvt('Y_TOTALQTE1',-O.GetMvt('Y_TOTALQTE1')) ;
      O.PutMvt('Y_TOTALQTE2',-O.GetMvt('Y_TOTALQTE2')) ;
      O.PutMvt('Y_TOTALECRITURE',-O.GetMvt('Y_TOTALECRITURE')) ;
      O.PutMvt('Y_TOTALDEVISE',-O.GetMvt('Y_TOTALDEVISE')) ;
      END else
      BEGIN
      X:=O.GetMvt('Y_DEBIT') ; O.PutMvt('Y_DEBIT',O.GetMvt('Y_CREDIT')) ; O.PutMvt('Y_CREDIT',X) ;
      X:=O.GetMvt('Y_DEBITDEV') ; O.PutMvt('Y_DEBITDEV',O.GetMvt('Y_CREDITDEV')) ; O.PutMvt('Y_CREDITDEV',X) ;
      END ;
   END ;
END ;

procedure TFExtourne.NewPieceEcr ;
Var O,OHisto : TOBM ;
    i : integer ;
    Premier : boolean ;
    StRef : String35 ;
BEGIN
Premier:=True ; OHisto:=Nil ;
for i:=0 to ListeEcr.Count-1 do
    BEGIN
    O:=TOBM(ListeEcr[i]) ; if O=Nil then Break ;
    StRef:=Copy(HM.Mess[9]+' '+Inttostr(O.GetMvt('E_NUMEROPIECE'))+'  '+DateToStr(O.GetMvt('E_DATECOMPTABLE')),1,35) ;
    TripoteO(O,tsGene) ;
    TEcrGen.Insert ;
    O.EgalChamps(TEcrGen) ;
    TEcrGen.FindField('E_REFINTERNE').AsString:=StRef ;
    TEcrGen.Post ;
    if CONTRE_TYPE.Value='N' then MajSoldeCompte(TEcrGen) ;
    if ((Premier) and (OHisto=Nil)) then
       BEGIN
       OHisto:=TOBM.Create(EcrGen,'',False) ; OHisto.ChargeMvt(TEcrGen) ; TPIECE.Add(OHisto) ;
       END ;
    Premier:=False ;
    END ;
for i:=0 to ListeAna.Count-1 do
    BEGIN
    O:=TOBM(ListeAna[i]) ; if O=Nil then Break ;
    TripoteO(O,tsAnal) ;
    TEcrAna.Insert ; O.EgalChamps(TEcrAna) ; TEcrAna.Post ;
    if CONTRE_TYPE.Value='N' then MajSoldeSection(TEcrAna,True) ;
    END ;
END ;

procedure TFExtourne.GenereLesPieces ;
Var i : integer ;
    M,RR : RMVT ;
    O : TOBM ;
BEGIN
InitMove(FListe.NbSelected,'') ;
TEcrGen.Open ; TEcrAna.Open ;
if FListe.AllSelected then
   BEGIN
   Q.First ;
   While Not Q.EOF do
      BEGIN
      MoveCur(False) ;
      if TrouveSaisie(Q,M,E_QUALIFPIECE.Value) then
         BEGIN
         ChargeLaPiece(M) ; NewNum:=GetNewNumJal(M.Jal,(CONTRE_TYPE.Value='N'),GeneDate) ;
         NewPieceEcr ;
         ADevalider(M.Jal,GeneDate) ;
         END ;
      Q.Next ;
      END ;
   END else
   BEGIN
   for i:=0 to FListe.NbSelected-1 do
       BEGIN
       MoveCur(False) ; FListe.GotoLeBookMark(i) ;
       if TrouveSaisie(Q,M,E_QUALIFPIECE.Value) then
          BEGIN
          ChargeLaPiece(M) ; NewNum:=GetNewNumJal(M.Jal,(CONTRE_TYPE.Value='N'),GeneDate) ;
          NewPieceEcr ;
          ADevalider(M.Jal,GeneDate) ;
          END ;
       END ;
   END ;
TEcrGen.Close ; TEcrAna.Close ;
FiniMove ;
{#TVAENC}
if VH^.OuiTvaEnc then
   BEGIN
   InitMove(TPiece.Count,'') ;
   for i:=0 to TPiece.Count-1 do
       BEGIN
       MoveCur(False) ;
       O:=TOBM(TPiece[i]) ; RR:=OBMToIdent(O,False) ;
       //ElementsTvaEnc(RR,True) ;
       ElementsTvaEncRevise (M, True, InfoImp, QFiche) ;
       END ;
   FiniMove ;
   END ;
MarquerPublifi(True) ;
CPStatutDossier ; 
Screen.Cursor:=SyncrDefault ;
END ;

Procedure TFExtourne.TraiteTPVal ;
Var i : integer ;
    O : TOBM ;
    MM : RMVT ;
BEGIN
{$IFNDEF CCS3}
for i:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[i]) ; MM:=OBMToIdent(O,False) ;
    GenerePiecesPayeur(MM) ;
    END ;
{$ENDIF}
END ;

procedure TFExtourne.BOuvrirClick(Sender: TObject);
begin
  inherited;
if Not JePeuxValider then BEGIN Screen.Cursor:=SyncrDefault ; Exit ; END ;
if HM.Execute(0,'','')<>mrYes then Exit ;
Application.ProcessMessages ; NowFutur:=NowH ;
if Transactions(GenereLesPieces,3)<>oeOK then MessageAlerte(HM.Mess[7]) else
   BEGIN
   if TPiece.Count>0 then
      BEGIN
      if VH^.OuiTP then TraiteTPVal ;
      if HM.Execute(10,'','')=mrYes then VisuPiecesGenere(TPiece,EcrGen,10) ;
      END ;
   END ;
FListe.ClearSelected ;
end;

procedure TFExtourne.BNouvRechClick(Sender: TObject);
Var OldT,OldC : String ;
begin
OldT:=E_QUALIFPIECE.Value ;
OldC:=CONTRE_TYPE.Value ;
  inherited;
E_QUALIFPIECE.Value:=OldT ;
CONTRE_TYPE.Value:=OldC ;
end;

procedure TFExtourne.CDejaExtClick(Sender: TObject);
begin
  inherited;
Case CDejaExt.State of
   cbChecked   : XX_WHEREEXT.Text:='E_QUALIFORIGINE="URN"' ;
   cbUnChecked : XX_WHEREEXT.Text:='((E_QUALIFORIGINE<>"TP" AND E_QUALIFORIGINE<>"URN") OR E_QUALIFORIGINE="") ' ;
      cbGrayed : XX_WHEREEXT.Text:='(E_QUALIFORIGINE<>"TP" OR E_QUALIFORIGINE="")' ;
   END ;
end;

end.
