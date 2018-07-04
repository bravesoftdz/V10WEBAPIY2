unit ImputAna;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, StdCtrls, Grids, DBGrids, HDB,
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Hcompte, Mask, hmsgbox, Ent1, HEnt1,
  SaisUtil, SaisComm, LettUtil, LetBatch, SaisODA, DelVisuE, HStatus, Ventil,
  HTB97, ed_tools, ColMemo, HPanel, Uiutil, HRichOLE, ADODB ;

Procedure ReImputAnalytiques ;

type
  TFImputAna = class(TFMul)
    TabSheet1: TTabSheet;
    Bevel5: TBevel;
    TY_AXE: TLabel;
    Y_AXE: THValComboBox;
    HLabel4: THLabel;
    Y_GENERAL_SUP: THCpteEdit;
    HLabel5: THLabel;
    Y_GENERAL_INF: THCpteEdit;
    TE_EXERCICE: THLabel;
    Y_EXERCICE: THValComboBox;
    TY_SECTION: TLabel;
    Y_SECTION: THCpteEdit;
    TE_ETABLISSEMENT: THLabel;
    Y_ETABLISSEMENT: THValComboBox;
    TE_DATECOMPTABLE: THLabel;
    Y_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    Y_DATECOMPTABLE_: THCritMaskEdit;
    TE_JOURNAL: THLabel;
    Y_JOURNAL: THValComboBox;
    TE_DEVISE: THLabel;
    Y_DEVISE: THValComboBox;
    TE_NATUREPIECE: THLabel;
    Y_NATUREPIECE: THValComboBox;
    TE_NUMEROPIECE: THLabel;
    Y_NUMEROPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    Y_NUMEROPIECE_: THCritMaskEdit;
    RImput: TRadioGroup;
    HLabel2: THLabel;
    JALGENERE: THValComboBox;
    HLabel3: THLabel;
    DATEGENERE: THCritMaskEdit;
    ContreP: TCheckBox;
    Ratio: THNumEdit;
    HLabel6: THLabel;
    HLabel7: THLabel;
    VENTTYPE: THValComboBox;
    HM: THMsgBox;
    QEcrAna: TQuery;
    Y_ECRANOUVEAU: THCritMaskEdit;
    Y_QUALIFPIECE: THCritMaskEdit;
    XX_WHERE: TEdit;
    Y_TYPEANALYTIQUE: THCritMaskEdit;
    Y_NUMLIGNE: THCritMaskEdit;
    BListePIECES: TToolbarButton97;
    REFGENERE: TEdit;
    HLabel8: THLabel;
    BZoomVentil: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure Y_EXERCICEChange(Sender: TObject);
    procedure Y_AXEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RImputClick(Sender: TObject);
    procedure DATEGENEREExit(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure BChercheClick(Sender: TObject); override;
    procedure Y_SECTIONChange(Sender: TObject);
    procedure RatioExit(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BListePIECESClick(Sender: TObject);
    procedure BZoomVentilClick(Sender: TObject);
  private
    DateGene,NowFutur : TDateTime ;
    ExoGene,OldGene,OldEtab,OldDevise : String ;
    NumGene,NumVentil,OldNumL  : Longint ;
    OldTaux            : double ;
    NeedRech           : boolean ;
    Grille   : HTStrings ;
    TPieces  : TList ;
    DEV      : RDEVISE ;
    Function  LaDateOk : boolean ;
    Function  ChargeGrille : boolean ;
    Function  ControleOk : boolean ;
    Function  JournalOk : boolean ;
    procedure ODDevise ( QA : TQuery ) ;
    procedure CreerLignesAna ( QA : TQuery ) ;
    procedure ContrePasse ( QA : TQuery ) ;
    procedure Supprime ( QA : TQuery ) ;
    procedure TraiteDC ( QA : TQuery ; ii : integer ; Var TotD,TotP,TotE,Pourc : Double ) ;
    procedure AffecteDevise ( QA : TQuery ) ;
    procedure ReImputAna ;
    procedure ImputeLesAna ;
    procedure SoldeJournal ( QA : TQuery ) ;
    procedure TraiteQte(QA: TQuery; ii: integer; var TauxQte1, TauxQte2, TotalQte1, TotalQte2, TotalTauxQte1, TotalTauxQte2 : Double );
    function VerifModelRestriction: Boolean; {FP 29/12/2005}
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  CPProcMetier,
  ULibExercice,
  {$ENDIF MODENT1}
  UtilPgi,ULibAnalytique ; // CPMODRESTANA_TOM;       {FP 29/12/2005}

Procedure ReImputAnalytiques ;
Var X  : TFImputAna ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
if _Blocage(['nrCloture','nrBatch','nrSaisieModif'],False,'nrBatch') then Exit ;
X:=TFImputAna.Create(Application) ;
X.FNomFiltre:='REIMPUTANA' ;
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

procedure TFImputAna.FormShow(Sender: TObject);
begin
if Not VH^.MontantNegatif then BEGIN ContreP.Enabled:=False ; ContreP.Checked:=False ; END ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   Y_EXERCICE.Value:=VH^.CPExoRef.Code ;
   Y_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   Y_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   Y_EXERCICE.Value:=VH^.Entree.Code ;
   Y_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
   Y_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
Y_DEVISE.Value:=V_PGI.DevisePivot ;
PositionneEtabUser(Y_ETABLISSEMENT) ;
DATEGENERE.Text:=DateToStr(V_PGI.DateEntree) ; DateGene:=V_PGI.DateEntree ;
if JALGENERE.Values.Count>0 then JALGENERE.Value:=JALGENERE.Values[0] ;
if Y_AXE.Values.Count>0 then
   BEGIN
   Y_AXE.Value:=Y_AXE.Values[0] ;
   Y_Section.Text:=VH^.Cpta[AxeToFb(Y_Axe.Value)].Attente ;
   END ;
ChangeSQL(QEcrAna) ;
  inherited;
Pages.Pages[2].TabVisible:=False ; Pages.Pages[3].TabVisible:=False ;
Q.Manuel:=FALSE ;
CentreDBGRid(FListe) ; NeedRech:=False ; 
end;

procedure TFImputAna.Y_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(Y_EXERCICE.Value,Y_DATECOMPTABLE,Y_DATECOMPTABLE_) ;
end;

procedure TFImputAna.Y_AXEChange(Sender: TObject);
begin
  inherited;
if Y_AXE.Value='' then Exit ;
Y_SECTION.ZoomTable:=AxeToTz(Y_AXE.Value) ;
if Y_SECTION.ExisteH<=0 then Y_SECTION.Text:='' ;
NeedRech:=True ; 
end;

procedure TFImputAna.FormCreate(Sender: TObject);
begin
  inherited;
Grille:=HTStringList.Create ;
TPieces:=TList.Create ; NeedRech:=False ;
Q.Manuel:=True ; Q.Liste:='REIMPUTANAL' ;
end;

procedure TFImputAna.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
Grille.Clear ; Grille.Free ;
VideListe(TPieces) ; TPieces.Free ;
if isInside(Self) then _Bloqueur('nrBatch',False) ;
end;

procedure TFImputAna.RImputClick(Sender: TObject);
begin
  inherited;
if RImput.ItemIndex=0 then
   BEGIN
   JALGENERE.Enabled:=False ; JALGENERE.Value:='' ;
   REFGENERE.Enabled:=False ; REFGENERE.Text:='' ;
   DATEGENERE.Enabled:=False ; Ratio.Enabled:=False ; ContreP.Enabled:=False ;
   END else
   BEGIN
   JALGENERE.Enabled:=True ; REFGENERE.Enabled:=True ;
   DATEGENERE.Enabled:=True ; Ratio.Enabled:=True ;
   if Not VH^.MontantNegatif then
      BEGIN
      ContreP.Checked:=False ; ContreP.Enabled:=False ;
      END else
      BEGIN
      ContreP.Enabled:=True ;
      END ; 
   END ;
NeedRech:=True ;
end;

procedure TFImputAna.DATEGENEREExit(Sender: TObject);
begin
  inherited;
if Not LaDateOk then if DATEGENERE.CanFocus then DATEGENERE.SetFocus ; 
end;

Function TFImputAna.LaDateOk : boolean ;
Var DD : TDateTime ;
    Err : integer ;
BEGIN
Result:=False ;
if Not IsValidDate(DATEGENERE.Text) then
   BEGIN
   HM.Execute(4,Caption,'') ;
   DATEGENERE.Text:=DateToStr(V_PGI.DateEntree) ; DateGene:=V_PGI.DateEntree ;
   END else
   BEGIN
   DD:=StrToDate(DATEGENERE.Text) ; Err:=DateCorrecte(DD) ;
   if Err>0 then
      BEGIN
      HM.Execute(Err+4,Caption,'') ;
      DATEGENERE.Text:=DateToStr(V_PGI.DateEntree) ;
      DateGene:=V_PGI.DateEntree ;
      END else
      BEGIN
      if RevisionActive(DD) then
         BEGIN
         DATEGENERE.Text:=DateToStr(V_PGI.DateEntree) ;
         DateGene:=V_PGI.DateEntree ;
         END else
         BEGIN
         DateGene:=DD ; Result:=True ;
         END ;
      END ;
   END ;
END ;

Function TFImputAna.ChargeGrille : boolean ;
Var QG : TQuery ;
    X,Tot : double ;
    Find  : boolean ;
    St    : String ;
BEGIN
Find:=False ;
Grille.Clear ; Result:=False ; Tot:=0 ;
QG:=OpenSQL('Select * from VENTIL Where V_NATURE="TY'+Y_AXE.Value[2]+'" AND V_COMPTE="'+VENTTYPE.Value+'" Order By V_NUMEROVENTIL',True) ;
While Not QG.EOF do
   BEGIN
   X:=QG.FindField('V_TAUXMONTANT').AsFloat ; Tot:=Tot+X ;
   // MODIF FICHE 12570
   St := QG.FindField('V_SECTION').AsString + ';'
        + StrfPoint(X) + ';'
        + StrfPoint(QG.FindField('V_TAUXQTE1').AsFloat) + ';'
        + StrfPoint(QG.FindField('V_TAUXQTE2').AsFloat) ;
   // FIN MODIF FICHE 12570
   Grille.Add(St) ;
   Find:=True ;
   QG.Next ;
   END ;
Ferme(QG) ;
if Not Find then Exit ;
if Arrondi(Tot-100.0,ADecimP)<>0 then Exit ;
Result:=True ;
END ;

Function TFImputAna.JournalOk : boolean ;
Var Q : TQuery ;
BEGIN
Result:=True ;
Q:=OpenSQL('SELECT J_FERME, J_COMPTEURNORMAL, J_AXE from JOURNAL Where J_JOURNAL="'+JALGENERE.Value+'"',True) ;
if Not Q.EOF then
   BEGIN
   if Q.Fields[0].AsString='X' then BEGIN HM.Execute(12,Caption,'') ; Result:=False ; END else
    if Q.Fields[1].AsString='' then BEGIN HM.Execute(13,Caption,'') ; Result:=False ; END else
     if Q.Fields[2].AsString<>Y_AXE.Value then BEGIN HM.Execute(14,Caption,'') ; Result:=False ; END ;
   END else Result:=False ;
Ferme(Q) ;
END ;

Function TFImputAna.ControleOk : boolean ;
BEGIN
Result:=False ;
if Y_AXE.Value='' then BEGIN if Y_AXE.CanFocus then Y_AXE.SetFocus ; HM.Execute(0,Caption,'') ; Exit ; END ;
if Y_SECTION.Text='' then BEGIN if Y_SECTION.CanFocus then Y_SECTION.SetFocus ; HM.Execute(1,Caption,'') ; Exit ; END ;
if Y_SECTION.ExisteH<=0 then BEGIN if Y_SECTION.CanFocus then Y_SECTION.SetFocus ; HM.Execute(2,Caption,'') ; Exit ; END ;
if VENTTYPE.Value='' then BEGIN if VENTTYPE.CanFocus then VENTTYPE.SetFocus ; HM.Execute(3,Caption,'') ; Exit ; END ;
if Not ChargeGrille then BEGIN if VENTTYPE.CanFocus then VENTTYPE.SetFocus ; HM.Execute(10,Caption,'') ; Exit ; END ;
if RImput.ItemIndex=1 then
   BEGIN
   if JALGENERE.Value='' then BEGIN if JALGENERE.CanFocus then JALGENERE.SetFocus ; HM.Execute(9,Caption,'') ; Exit ; END ;
   if Not LaDateOk then Exit ;
   if Not JournalOk then Exit ;
   END else
   BEGIN
   if Grille.Count>1 then HM.Execute(21,Caption,VENTTYPE.Text) ;
   END ;
{b FP 19/04/2006 FQ 17729}
if not VerifModelRestriction then
  begin
  HM.Execute(22,Caption,'') ;
  Exit;
  end;
{e FP 19/04/2006}
Result:=True ;
END ;

procedure TFImputAna.SoldeJournal ( QA : TQuery ) ;
Var D,C         : Double ;
    GeneTypeExo : TTypeExo ;
    FRM         : TFRM ;
    ll          : LongInt ;
BEGIN
if RImput.itemIndex=0 then Exit ;
if ExoGene=VH^.Encours.Code then GeneTypeExo:=teEncours else
 if ExoGene=VH^.Suivant.Code then GeneTypeExo:=teSuivant else Exit ;

  FRM.Cpt   := QA.FindField('Y_JOURNAL').AsString ;
  FRM.NumD  := QA.FindField('Y_NUMEROPIECE').AsInteger ;
  FRM.DateD := QA.FindField('Y_DATECOMPTABLE').AsDateTime ;
  D         := QA.FindField('Y_DEBIT').AsFloat ;
  C         := QA.FindField('Y_CREDIT').AsFloat ;
  AttribParamsNew ( FRM, D, C, GeneTypeExo ) ;
  ll := ExecReqMAJ ( fbJal, False, False, FRM ) ;
  If ll<>1 then V_PGI.IoError:=oeUnknown ;

END ;

procedure TFImputAna.ImputeLesAna ;
Var QA    : TQuery ;
    SQL   : String ;
    Rupt  : boolean ;
    F     : TField ;
    OkSpe : Boolean ;
BEGIN
OkSpe:=False ;
if ((Grille.Count>1) and (RImput.ItemIndex=0)) then
   BEGIN
   if Q.FindField('Y_NUMVENTIL').AsInteger<>1 then Exit ;
   F:=Q.FindField('Y_POURCENTAGE') ;
   if F<>Nil then BEGIN if Arrondi(F.AsFloat-100,ADecimP)<>0 then Exit ; END ;
   OkSpe:=True ;
   END ;
SQL:='Select * from ANALYTIQ Where Y_JOURNAL="'+Q.FindField('Y_JOURNAL').AsString+'" '
    +'AND Y_EXERCICE="'+Q.FindField('Y_EXERCICE').AsString+'" AND Y_DATECOMPTABLE="'+UsDateTime(Q.FindField('Y_DATECOMPTABLE').AsDateTime)+'" '
    +'AND Y_NUMEROPIECE='+IntToStr(Q.FindField('Y_NUMEROPIECE').AsInteger)+' AND Y_NUMLIGNE='+IntToStr(Q.FindField('Y_NUMLIGNE').AsInteger)+' '
    +'AND Y_AXE="'+Q.FindField('Y_AXE').AsString+'" AND Y_NUMVENTIL='+IntToStr(Q.FindField('Y_NUMVENTIL').AsInteger)+' '
    +'AND Y_QUALIFPIECE="N"' ;
if OkSpe then SQL:=SQL+' AND Y_POURCENTAGE=100' ;
QA:=OpenSQL(SQL,True) ; if RImput.ItemIndex=0 then NumVentil:=0 ;
if Not QA.EOF then
   BEGIN
   if RImput.ItemIndex=1 then
      BEGIN
      Rupt:=False ;
      { 1 pièce par Général / Etablissement / Devise / TauxDev }
      if ((OldGene<>QA.FindField('Y_GENERAL').AsString) or
          (OldEtab<>QA.FindField('Y_ETABLISSEMENT').AsString) or
          (OldDevise<>QA.FindField('Y_DEVISE').AsString) or
          (OldNumL<>QA.FindField('Y_NUMLIGNE').AsInteger) or
          (OldTaux<>QA.FindField('Y_TAUXDEV').AsFloat)) then Rupt:=True ;
      OldGene:=QA.FindField('Y_GENERAL').AsString ;
      OldEtab:=QA.FindField('Y_ETABLISSEMENT').AsString ;
      OldDevise:=QA.FindField('Y_DEVISE').AsString ;
      OldTaux:=QA.FindField('Y_TAUXDEV').AsFloat ;
      OldNumL:=QA.FindField('Y_NUMLIGNE').AsInteger ;
      if Rupt then
         BEGIN
         NumGene:=GetNewNumJal(JALGENERE.Value,True,DateGene) ; if NumGene<=0 then V_PGI.IoError:=oeUnknown ;
         NumVentil:=0 ;
         END ;
      END ;
   AffecteDevise(QA) ;
   CreerLignesAna(QA) ;
   END ;
Ferme(QA) ;
END ;

procedure TFImputAna.ReImputAna ;
Var i : integer ;
BEGIN
DEV.Code:='' ; OldGene:='' ; OldEtab:='' ; OldDevise:='' ;
OldTaux:=0 ; OldNumL:=-1 ; NumVentil:=0 ;
QEcrAna.Open ;
{b FP 19/04/2006 FQ 17729 Dans  une transaction, pas de ShowMessage
if not VerifModelRestriction then
  begin
  HM.Execute(22,Caption,'') ;
  end
else
  begin
}
  QEcrAna.First;

  if Fliste.AllSelected then
     BEGIN
     InitMove(100,'') ;
     Q.First ;
     While Not Q.EOF do
        BEGIN
        MoveCur(False) ;
        ImputeLesAna ; if V_PGI.IoError<>oeOk then Break ;
        Q.Next ;
        END ;
     END else
     BEGIN
     InitMove(FListe.NbSelected,'') ;
     for i:=0 to FListe.NbSelected-1 do
         BEGIN
         MoveCur(False) ;
         FListe.GotoLeBookmark(i) ;
         ImputeLesAna ;  if V_PGI.IoError<>oeOk then Break ;
         END ;
     END ;
QEcrAna.Close ;
FiniMove ;
if Not FListe.AllSelected then Fliste.ClearSelected else FListe.AllSelected:=False ;
END ;

procedure TFImputAna.BOuvrirClick(Sender: TObject);
begin
if NeedRech then BEGIN HM.Execute(19,caption,'') ; Exit ; END ;
  inherited;
if Not ControleOK then Exit ;
if ((Not FListe.AllSelected) and (Fliste.NbSelected<=0)) then BEGIN HM.Execute(18,caption,'') ; Exit ; END ;
if HM.Execute(16,caption,'')<>mrYes then Exit ;
ExoGene:=QuelExoDT(DateGene) ; NowFutur:=NowH ;
VideListe(TPieces) ;
if Transactions(ReimputAna,2)<>oeOk then MessageAlerte(HM.Mess[15]) else
   BEGIN
   if Not FListe.AllSelected then Fliste.ClearSelected else FListe.AllSelected:=False ;
   if ((RImput.ItemIndex=1) and (TPieces.Count>0)) then
      BEGIN
      BListePieces.Visible:=True ;
      if HM.Execute(20,caption,'')=mrYes then VisuPiecesGenere(TPieces,EcrAna,12) ;
      END else
      BEGIN
      HM.Execute(17,caption,'') ;
      END ;
   NeedRech:=True ;
   END ;
end;

procedure TFImputAna.AffecteDevise ( QA : TQuery ) ;
Var StDev : String ;
BEGIN
StDev:=QA.FindField('Y_DEVISE').AsString ;
if StDev<>DEV.Code then BEGIN DEV.Code:=StDev ; GetInfosDevise(DEV) ; END ;
END ;

procedure TFImputAna.TraiteDC ( QA : TQuery ; ii : integer ; Var TotD,TotP,TotE,Pourc : Double ) ;
Var Rati : double ;
BEGIN
if Rimput.ItemIndex=0 then Rati:=1.0 else Rati:=Ratio.Value ;
if ii<Grille.Count-1 then
   BEGIN
   QEcrAna.FindField('Y_DEBIT').AsFloat:=Arrondi(QA.FindField('Y_DEBIT').AsFloat*Pourc*Rati/100.0,V_PGI.OkDecV) ;
   QEcrAna.FindField('Y_CREDIT').AsFloat:=Arrondi(QA.FindField('Y_CREDIT').AsFloat*Pourc*Rati/100.0,V_PGI.OkDecV) ;
   QEcrAna.FindField('Y_DEBITDEV').AsFloat:=Arrondi(QA.FindField('Y_DEBITDEV').AsFloat*Pourc*Rati/100.0,DEV.Decimale) ;
   QEcrAna.FindField('Y_CREDITDEV').AsFloat:=Arrondi(QA.FindField('Y_CREDITDEV').AsFloat*Pourc*Rati/100.0,DEV.Decimale) ;
   TotD:=Arrondi(TotD+QEcrAna.FindField('Y_DEBITDEV').AsFloat+QEcrAna.FindField('Y_CREDITDEV').AsFloat,DEV.Decimale) ;
   TotP:=Arrondi(TotP+QEcrAna.FindField('Y_DEBIT').AsFloat+QEcrAna.FindField('Y_CREDIT').AsFloat,V_PGI.OkDecV) ;
   END else
   BEGIN
   if (QA.FindField('Y_DEBITDEV').AsFloat<>0) Or (QA.FindField('Y_DEBIT').AsFloat<>0) {FQ22178 01.08 YMO} then
      BEGIN
      QEcrAna.FindField('Y_DEBITDEV').AsFloat:=Arrondi((QA.FindField('Y_DEBITDEV').AsFloat*Rati)-TotD,DEV.Decimale) ;
      QEcrAna.FindField('Y_DEBIT').AsFloat:=Arrondi((QA.FindField('Y_DEBIT').AsFloat*Rati)-TotP,V_PGI.OkDecV) ;
      END else
      BEGIN
      QEcrAna.FindField('Y_CREDITDEV').AsFloat:=Arrondi((QA.FindField('Y_CREDITDEV').AsFloat*Rati)-TotD,DEV.Decimale) ;
      QEcrAna.FindField('Y_CREDIT').AsFloat:=Arrondi((QA.FindField('Y_CREDIT').AsFloat*Rati)-TotP,V_PGI.OkDecV) ;
      END ;
   END ;
END ;

procedure TFImputAna.ODDevise ( QA : TQuery ) ;
BEGIN
if RImput.ItemIndex=0 then Exit ;
if QA.FindField('Y_DEVISE').AsString=V_PGI.DevisePivot then Exit ;
QA.FindField('Y_DEBITDEV').AsFloat:=QA.FindField('Y_DEBIT').AsFloat ;
QA.FindField('Y_CREDITDEV').AsFloat:=QA.FindField('Y_CREDIT').AsFloat ;
QA.FindField('Y_TAUXDEV').AsFloat:=1.0 ;
QA.FindField('Y_DEVISE').AsString:=V_PGI.DevisePivot ;
END ;

procedure TFImputAna.Supprime ( QA : TQuery ) ;
Var SW : String ;
BEGIN
MajSoldeSection(QA,False) ;
SW:='Y_JOURNAL="'+QA.FindField('Y_JOURNAL').AsString+'" AND Y_EXERCICE="'+QA.FindField('Y_EXERCICE').AsString+'" AND Y_DATECOMPTABLE="'+USDateTime(QA.FindField('Y_DATECOMPTABLE').AsDateTime)+'" '
   +'AND Y_NUMEROPIECE='+Inttostr(QA.FindField('Y_NUMEROPIECE').AsInteger)+' AND Y_AXE="'+Y_AXE.Value+'" '
   +'AND Y_NUMLIGNE='+Inttostr(QA.FindField('Y_NUMLIGNE').AsInteger)+' AND Y_QUALIFPIECE="N"' ;
if ExecuteSQL('UPDATE ANALYTIQ Set Y_NUMVENTIL=-1 where '+SW+' AND Y_NUMVENTIL='+IntToStr(QA.FindField('Y_NUMVENTIL').AsInteger))<=0 then V_PGI.Ioerror:=oeUnknown ;
if ExecuteSQL('UPDATE ANALYTIQ Set Y_NUMVENTIL=Y_NUMVENTIL-1000 Where '+SW+' AND Y_NUMVENTIL>1000')<=0 then V_PGI.Ioerror:=oeUnknown ;
if ExecuteSQL('DELETE FROM ANALYTIQ Where '+SW+' AND Y_NUMVENTIL=-1')<=0 then V_PGI.Ioerror:=oeUnknown ;
END ;

procedure TFImputAna.ContrePasse ( QA : TQuery ) ;
Var Rati : double ;
BEGIN
if Rimput.ItemIndex=0 then Rati:=1.0 else Rati:=Ratio.Value ;
QEcrAna.Insert ; InitNew(QEcrAna) ;
Inc(NumVentil) ;
QEcrAna.FindField('Y_SECTION').AsString:=QA.FindField('Y_SECTION').AsString ;
QEcrAna.FindField('Y_GENERAL').AsString:=QA.FindField('Y_GENERAL').AsString ;
QEcrAna.FindField('Y_ETABLISSEMENT').AsString:=QA.FindField('Y_ETABLISSEMENT').AsString ;
QEcrAna.FindField('Y_TYPEANOUVEAU').AsString:=QA.FindField('Y_TYPEANOUVEAU').AsString ;
QEcrAna.FindField('Y_DEVISE').AsString:=QA.FindField('Y_DEVISE').AsString ;
QEcrAna.FindField('Y_TAUXDEV').AsFloat:=QA.FindField('Y_TAUXDEV').AsFloat ;
QEcrAna.FindField('Y_DATETAUXDEV').AsDateTime:=QA.FindField('Y_DATETAUXDEV').AsDateTime ;
QEcrAna.FindField('Y_JOURNAL').AsString:=JALGENERE.Value ;
QEcrAna.FindField('Y_NUMVENTIL').AsInteger:=NumVentil ;
If RefGenere.Text<>'' Then QEcrAna.FindField('Y_REFINTERNE').AsString:=REFGENERE.Text
                      Else QEcrAna.FindField('Y_REFINTERNE').AsString:=QA.FindField('Y_REFINTERNE').AsString ;
QEcrAna.FindField('Y_LIBELLE').AsString:=QA.FindField('Y_LIBELLE').AsString ;
QEcrAna.FindField('Y_AXE').AsString:=Y_AXE.Value ; QEcrAna.FindField('Y_QUALIFPIECE').AsString:='N' ;
QEcrAna.FindField('Y_ECRANOUVEAU').AsString:='N' ; QEcrAna.FindField('Y_CREERPAR').AsString:='GEN' ;
QEcrAna.FindField('Y_POURCENTAGE').AsFloat:=0    ; QEcrAna.FindField('Y_DATEMODIF').AsDateTime:=NowFutur ;
QEcrAna.FindField('Y_TYPEMVT').AsString:='AL'    ; QEcrAna.FindField('Y_DATECOMPTABLE').AsDateTime:=DateGene ;
{$IFNDEF SPEC302}
QEcrAna.FindField('Y_PERIODE').AsInteger:=GetPeriode(DateGene) ;
QEcrAna.FindField('Y_SEMAINE').AsInteger:=NumSemaine(DateGene) ;
{$ENDIF}
QEcrAna.FindField('Y_NUMLIGNE').AsInteger:=0     ; QEcrAna.FindField('Y_NUMEROPIECE').AsInteger:=NumGene ;
QEcrAna.FindField('Y_TOTALECRITURE').AsFloat:=0  ; QEcrAna.FindField('Y_NATUREPIECE').AsString:='OD' ;
QEcrAna.FindField('Y_TOTALDEVISE').AsFloat:=0    ; QEcrAna.FindField('Y_EXERCICE').AsString:=ExoGene ;
QEcrAna.FindField('Y_TYPEANALYTIQUE').AsString:='X' ;
if ContreP.Checked then
   BEGIN
   QEcrAna.FindField('Y_DEBIT').AsFloat:=Arrondi(-QA.FindField('Y_DEBIT').AsFloat*Rati,V_PGI.OkDecV) ;
   QEcrAna.FindField('Y_CREDIT').AsFloat:=Arrondi(-QA.FindField('Y_CREDIT').AsFloat*Rati,V_PGI.OkDecV) ;
   QEcrAna.FindField('Y_DEBITDEV').AsFloat:=Arrondi(-QA.FindField('Y_DEBITDEV').AsFloat*Rati,DEV.Decimale) ;
   QEcrAna.FindField('Y_CREDITDEV').AsFloat:=Arrondi(-QA.FindField('Y_CREDITDEV').AsFloat*Rati,DEV.Decimale) ;
   END else
   BEGIN
   QEcrAna.FindField('Y_DEBIT').AsFloat:=Arrondi(QA.FindField('Y_CREDIT').AsFloat*Rati,V_PGI.OkDecV) ;
   QEcrAna.FindField('Y_CREDIT').AsFloat:=Arrondi(QA.FindField('Y_DEBIT').AsFloat*Rati,V_PGI.OkDecV) ;
   QEcrAna.FindField('Y_DEBITDEV').AsFloat:=Arrondi(QA.FindField('Y_CREDITDEV').AsFloat*Rati,DEV.Decimale) ;
   QEcrAna.FindField('Y_CREDITDEV').AsFloat:=Arrondi(QA.FindField('Y_DEBITDEV').AsFloat*Rati,DEV.Decimale) ;
   END ;
// MODIF FICHE 12070 : renseigner les qtés
  // Renseignement généraux
  QEcrAna.FindField('Y_QUALIFECRQTE1').AsString := QA.FindField('Y_QUALIFECRQTE1').AsString ;
  QEcrAna.FindField('Y_QUALIFECRQTE2').AsString := QA.FindField('Y_QUALIFECRQTE2').AsString ;
  QEcrAna.FindField('Y_QUALIFQTE1').AsString  := QA.FindField('Y_QUALIFQTE1').AsString ;
  QEcrAna.FindField('Y_QUALIFQTE2').AsString  := QA.FindField('Y_QUALIFQTE2').AsString ;
  // Totaux des qtes, reprise de la ligne géné ou à 0 ???? a voir...
  QEcrAna.FindField('Y_TOTALQTE1').AsFloat    := Arrondi( QA.FindField('Y_QTE1').AsFloat * Rati , 3 ) ;
  QEcrAna.FindField('Y_TOTALQTE2').AsFloat    := Arrondi( QA.FindField('Y_QTE2').AsFloat * Rati , 3 ) ;
  // Taux
  QEcrAna.FindField('Y_POURCENTQTE1').AsFloat := 100.0 ;
  QEcrAna.FindField('Y_POURCENTQTE2').AsFloat := 100.0 ;
  // Qtés
  QEcrAna.FindField('Y_QTE1').AsFloat := -1 * Arrondi( QA.FindField('Y_QTE1').AsFloat * rati , 3 ) ;
  QEcrAna.FindField('Y_QTE2').AsFloat := -1 * Arrondi( QA.FindField('Y_QTE2').AsFloat * rati , 3 ) ;
// FIN MODIF FICHE 12070
ODDevise(QEcrAna) ;
QEcrAna.Post ;
MajSoldeSection(QEcrAna,True) ;
SoldeJournal(QEcrAna) ;
END ;

procedure TFImputAna.CreerLignesAna ( QA : TQuery ) ;
Var i : integer ;
    O : TOBM ;
    Sect,St,StP  : String ;
    Pourc,Pourc1,TotD,TotP,TotE : Double ;
    TauxQte1, TauxQte2    : Double ; // MODIF FICHE 12570 SBO
    TotalQte1, TotalQte2  : Double ; // MODIF FICHE 12570 SBO
    TotalTauxQte1         : Double ; // MODIF FICHE 12570 SBO
    TotalTauxQte2         : Double ; // MODIF FICHE 12570 SBO
BEGIN
TotD:=0 ; TotP:=0 ; TotE:=0 ;
  TotalQte1     := 0 ; // MODIF FICHE 12570 SBO
  TotalQte2     := 0 ; // MODIF FICHE 12570 SBO
  TotalTauxQte1 := 0 ; // MODIF FICHE 12570 SBO
  TotalTauxQte2 := 0 ; // MODIF FICHE 12570 SBO
for i:=0 to Grille.Count-1 do
    BEGIN
    St:=Grille[i] ; Sect:=ReadtokenSt(St) ; StP:=ReadtokenSt(St) ;
// MODIF FICHE 12570 SBO
// Récupération des données de la ventilation type pour les qté
    TauxQte1   := Valeur(ReadTokenSt(St)) ;
    TauxQte2   := Valeur(ReadTokenSt(St)) ;
// FIN MODIF FICHE 12570 SBO
    (* GP le 07/03/2000 A vérifier ....
    Pourc:=Valeur(ReadtokenSt(St))*QA.FindField('Y_POURCENTAGE').AsFloat/100.0 ;
    Et pour cause :*)
    (* Piot, il a bu !! signé piot
    Pourc1:=Valeur(ReadtokenSt(St))*QA.FindField('Y_POURCENTAGE').AsFloat/100.0 ;
    Pourc:=Valeur(ReadtokenSt(St)) ;
    *)
    Pourc1:=Valeur(StP)*QA.FindField('Y_POURCENTAGE').AsFloat/100.0 ;
    Pourc:=Valeur(StP) ;
    QEcrAna.Insert ; InitNew(QEcrAna) ;
    QEcrAna.FindField('Y_GENERAL').AsString:=QA.FindField('Y_GENERAL').AsString ;
    QEcrAna.FindField('Y_ETABLISSEMENT').AsString:=QA.FindField('Y_ETABLISSEMENT').AsString ;
    QEcrAna.FindField('Y_LIBELLE').AsString:=QA.FindField('Y_LIBELLE').AsString ;
    QEcrAna.FindField('Y_TYPEANOUVEAU').AsString:=QA.FindField('Y_TYPEANOUVEAU').AsString ;
    QEcrAna.FindField('Y_DEVISE').AsString:=QA.FindField('Y_DEVISE').AsString ;
    QEcrAna.FindField('Y_TAUXDEV').AsFloat:=QA.FindField('Y_TAUXDEV').AsFloat ;
    QEcrAna.FindField('Y_DATETAUXDEV').AsDateTime:=QA.FindField('Y_DATETAUXDEV').AsDateTime ;
    QEcrAna.FindField('Y_AXE').AsString:=Y_AXE.Value ; QEcrAna.FindField('Y_SECTION').AsString:=Sect ;
    QEcrAna.FindField('Y_QUALIFPIECE').AsString:='N' ; QEcrAna.FindField('Y_ECRANOUVEAU').AsString:='N' ;
    QEcrAna.FindField('Y_CREERPAR').AsString:='GEN'  ; QEcrAna.FindField('Y_DATEMODIF').AsDateTime:=NowFutur ;
    if RImput.ItemIndex=0 then
       BEGIN
       QEcrAna.FindField('Y_DATECOMPTABLE').AsDateTime:=QA.FindField('Y_DATECOMPTABLE').AsDateTime ;
{$IFNDEF SPEC302}
       QEcrAna.FindField('Y_PERIODE').AsInteger:=QA.FindField('Y_PERIODE').AsInteger ;
       QEcrAna.FindField('Y_SEMAINE').AsInteger:=QA.FindField('Y_SEMAINE').AsInteger ;
{$ENDIF}
       QEcrAna.FindField('Y_NUMEROPIECE').AsInteger:=QA.FindField('Y_NUMEROPIECE').AsInteger ;
       QEcrAna.FindField('Y_NUMLIGNE').AsInteger:=QA.FindField('Y_NUMLIGNE').AsInteger ;
       QEcrAna.FindField('Y_EXERCICE').AsString:=QA.FindField('Y_EXERCICE').AsString ;
       QEcrAna.FindField('Y_NATUREPIECE').AsString:=QA.FindField('Y_NATUREPIECE').AsString ;
       QEcrAna.FindField('Y_TYPEANALYTIQUE').AsString:='-' ;
       QEcrAna.FindField('Y_TOTALECRITURE').AsFloat:=QA.FindField('Y_TOTALECRITURE').AsFloat ;
       QEcrAna.FindField('Y_TOTALDEVISE').AsFloat:=QA.FindField('Y_TOTALDEVISE').AsFloat ;
       QEcrAna.FindField('Y_JOURNAL').AsString:=QA.FindField('Y_JOURNAL').AsString ;
       QEcrAna.FindField('Y_NUMVENTIL').AsInteger:=QA.FindField('Y_NUMVENTIL').AsInteger+1000+i ;
       QEcrAna.FindField('Y_POURCENTAGE').AsFloat:=Pourc1 ;
       QEcrAna.FindField('Y_CONTREPARTIEGEN').AsString:=QA.FindField('Y_CONTREPARTIEGEN').AsString ;
       QEcrAna.FindField('Y_CONTREPARTIEAUX').AsString:=QA.FindField('Y_CONTREPARTIEAUX').AsString ;
       QEcrAna.FindField('Y_REFINTERNE').AsString:=QA.FindField('Y_REFINTERNE').AsString ;
       END else
       BEGIN
       Inc(NumVentil) ;
       if NumVentil+i=1 then QEcrAna.FindField('Y_TYPEMVT').AsString:='AE' else QEcrAna.FindField('Y_TYPEMVT').AsString:='AL' ;
       QEcrAna.FindField('Y_DATECOMPTABLE').AsDateTime:=DateGene ; QEcrAna.FindField('Y_NUMEROPIECE').AsInteger:=NumGene ;
{$IFNDEF SPEC302}
       QEcrAna.FindField('Y_PERIODE').AsInteger:=GetPeriode(DateGene) ;
       QEcrAna.FindField('Y_SEMAINE').AsInteger:=NumSemaine(DateGene) ;
{$ENDIF}
       QEcrAna.FindField('Y_NUMLIGNE').AsInteger:=0              ; QEcrAna.FindField('Y_EXERCICE').AsString:=ExoGene ;
       QEcrAna.FindField('Y_NATUREPIECE').AsString:='OD'         ; QEcrAna.FindField('Y_TYPEANALYTIQUE').AsString:='X' ;
       QEcrAna.FindField('Y_TOTALECRITURE').AsFloat:=0           ; QEcrAna.FindField('Y_TOTALDEVISE').AsFloat:=0 ;
       QEcrAna.FindField('Y_JOURNAL').AsString:=JALGENERE.Value ;
       QEcrAna.FindField('Y_POURCENTAGE').AsFloat:=0             ;
       If RefGenere.Text<>'' Then QEcrAna.FindField('Y_REFINTERNE').AsString:=REFGENERE.Text
                             Else QEcrAna.FindField('Y_REFINTERNE').AsString:=QA.FindField('Y_REFINTERNE').AsString ;
//       QEcrAna.FindField('Y_REFINTERNE').AsString:=REFGENERE.Text ;
       QEcrAna.FindField('Y_NUMVENTIL').AsInteger:=NumVentil     ;
       END ;
    TraiteDC(QA,i,TotD,TotP,TotE,Pourc) ;
// MODIF FICHE 12570 SBO
    TraiteQte(QA,i,TauxQte1,TauxQte2,TotalQte1,TotalQte2,TotalTauxQte1,TotalTauxQte2) ;
// FIN MODIF FICHE 12570 SBO
    ODDevise(QEcrAna) ;
    QEcrAna.Post ;
    if ((NumVentil=1) and (RImput.ItemIndex=1)) then
       BEGIN
       O:=TOBM.Create(EcrAna,Y_AXE.Value,False) ;
       O.ChargeMvt(QEcrAna) ; TPieces.Add(O) ;
       END ;
    MajSoldeSection(QEcrAna,True) ;
    SoldeJournal(QEcrAna) ;
    END ;
if RImput.ItemIndex=0 then Supprime(QA) else ContrePasse(QA) ;
END ;

procedure TFImputAna.BChercheClick(Sender: TObject);
begin
if Y_SECTION.Text='' then BEGIN if Y_SECTION.CanFocus then Y_SECTION.SetFocus ; HM.Execute(1,Caption,'') ; Exit ; END ;
if Y_SECTION.ExisteH<=0 then BEGIN if Y_SECTION.CanFocus then Y_SECTION.SetFocus ; HM.Execute(2,Caption,'') ; Exit ; END ;
  inherited;
NeedRech:=False ;
end;

procedure TFImputAna.Y_SECTIONChange(Sender: TObject);
begin
  inherited;
NeedRech:=True ;
end;

procedure TFImputAna.RatioExit(Sender: TObject);
begin
  inherited;
if ((Ratio.Value<=0) or (Ratio.Value>1)) then Ratio.Value:=1 ;
end;

procedure TFImputAna.FListeDblClick(Sender: TObject);
begin
  inherited;
if ((Q.Eof) And (Q.Bof)) then Exit ;
TrouveEtLanceSaisieODA(Q,taConsult) ;
end;                         

procedure TFImputAna.BListePIECESClick(Sender: TObject);
begin
  inherited;
if TPieces.Count<=0 then Exit ;
VisuPiecesGenere(TPieces,EcrAna,12) ;
end;

procedure TFImputAna.BZoomVentilClick(Sender: TObject);
begin
  inherited;
if VENTTYPE.Value<>'' then ParamVentil('TY',VENTTYPE.Value,'12345',taConsult,True) ; 
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 16/09/2003
Modifié le ... : 16/09/2003
Description .. :
Suite ........ : Recalcule les infos des qtés dans QEcrAna en fonctions
Suite ........ : des paramètres, en prenant en compte la gestion des
Suite ........ : arrondis
Mots clefs ... :
*****************************************************************}
procedure TFImputAna.TraiteQte(QA: TQuery; ii: integer; var TauxQte1, TauxQte2, TotalQte1, TotalQte2, TotalTauxQte1, TotalTauxQte2: Double);
var Qte1, Qte2   : Double ;
begin

  // Renseignement généraux
  QEcrAna.FindField('Y_QUALIFECRQTE1').AsString := QA.FindField('Y_QUALIFECRQTE1').AsString ;
  QEcrAna.FindField('Y_QUALIFECRQTE2').AsString := QA.FindField('Y_QUALIFECRQTE2').AsString ;
  QEcrAna.FindField('Y_QUALIFQTE1').AsString  := QA.FindField('Y_QUALIFQTE1').AsString ;
  QEcrAna.FindField('Y_QUALIFQTE2').AsString  := QA.FindField('Y_QUALIFQTE2').AsString ;
  QEcrAna.FindField('Y_POURCENTQTE1').AsFloat := Arrondi( TauxQte1 , 4 ) ;
  QEcrAna.FindField('Y_POURCENTQTE2').AsFloat := Arrondi( TauxQte2 , 4 ) ;

  // Totaux des qtes, reprise de la ligne géné ou à 0 ???? a voir...
  if Rimput.ItemIndex = 0 then
    begin
    QEcrAna.FindField('Y_TOTALQTE1').AsFloat    := QA.FindField('Y_TOTALQTE1').AsFloat ;
    QEcrAna.FindField('Y_TOTALQTE2').AsFloat    := QA.FindField('Y_TOTALQTE2').AsFloat ;
    end
  else
    begin
    QEcrAna.FindField('Y_TOTALQTE1').AsFloat    := Arrondi( QA.FindField('Y_QTE1').AsFloat * Ratio.Value , 3 ) ;
    QEcrAna.FindField('Y_TOTALQTE2').AsFloat    := Arrondi( QA.FindField('Y_QTE2').AsFloat * Ratio.Value , 3 ) ;
    end ;

  // Cumul des taux pour gestion des arrondi sur du 100%
  TotalTauxQte1 := TotalTauxQte1 + TauxQte1 ;
  TotalTauxQte2 := TotalTauxQte2 + TauxQte2 ;

  // GESTION QTE1
  if (QEcrAna.FindField('Y_TOTALQTE1').AsFloat = 0.0) or ( TauxQte1 = 0.0 ) then
    QEcrAna.FindField('Y_QTE1').AsFloat := 0.0
  else
    begin
    // Quantités à répartir :
    Qte1 := QEcrAna.FindField('Y_TOTALQTE1').AsFloat ;
    if ( ii = Grille.Count - 1 ) and ( Arrondi( TotalTauxQte1 - 100.0 , 4 ) = 0 ) then
      // ligne N : gestion des arrondi si taux de 100% -> on calcul le reste a placer
      QEcrAna.FindField('Y_QTE1').AsFloat := Qte1 - TotalQte1
    else
      // ligne 1 à n-1, application des taux sur qte à répartir
      begin
      Qte1 := Arrondi( Qte1 * (TauxQte1 / 100) , 3 ) ;
      QEcrAna.FindField('Y_QTE1').AsFloat := Qte1 ;
      TotalQte1     := TotalQte1 + Qte1 ;
      end ;
    end ;

  // GESTION QTE2
  if (QEcrAna.FindField('Y_TOTALQTE2').AsFloat = 0.0) or ( TauxQte2 = 0.0 ) then
    QEcrAna.FindField('Y_QTE2').AsFloat := 0.0
  else
    begin
    // Quantités à répartir :
    Qte2 := QEcrAna.FindField('Y_TOTALQTE2').AsFloat ;
    if ( ii = Grille.Count - 1 ) and ( Arrondi( TotalTauxQte2 - 100.0 , 4 ) = 0 ) then
      // ligne N : gestion des arrondi -> on calcul le reste a placer
      QEcrAna.FindField('Y_QTE2').AsFloat := Qte2 - TotalQte2
    else
      // ligne 1 à n-1, application des taux sur qte à répartir
      begin
      Qte2 := Arrondi( Qte2 * (TauxQte2 / 100) , 3 ) ;
      QEcrAna.FindField('Y_QTE2').AsFloat := Qte2 ;
      TotalQte2     := TotalQte2 + Qte2 ;
      end ;
    end ;

end;

function TFImputAna.VerifModelRestriction: Boolean;
var
  i:           integer;
  Lst:         HTStrings;
  RestrictAna: TRestrictionAnalytique;
BEGIN
  {b FP 29/12/2005}
  Result      := True;
  Lst         := HTStringList.Create;
  RestrictAna := TRestrictionAnalytique.Create;
  try
    if Fliste.AllSelected then
      BEGIN
      InitMove(100,'') ;
      Q.First ;
      While Not Q.EOF do
        BEGIN
        MoveCur(False) ;
        if Lst.IndexOf(Q.FindField('Y_GENERAL').AsString) = -1 then
          begin
          if not RestrictAna.IsCompteGeneAutorise(Q.FindField('Y_GENERAL').AsString,
                   Y_AXE.Value, 'TY'+Y_AXE.Value[2], VENTTYPE.Value) then
            begin
            Result := False;
            break;               
            end
          else
            begin
            Lst.Add(Q.FindField('Y_GENERAL').AsString);
            end;
          end;
        Q.Next ;
        END ;
      END
    else
      BEGIN
      InitMove(FListe.NbSelected,'') ;
      for i:=0 to FListe.NbSelected-1 do
        BEGIN
        MoveCur(False) ;
        FListe.GotoLeBookmark(i) ;
        if Lst.IndexOf(Q.FindField('Y_GENERAL').AsString) = -1 then
          begin
          if not RestrictAna.IsCompteGeneAutorise(Q.FindField('Y_GENERAL').AsString,
                   Y_AXE.Value, 'TY'+Y_AXE.Value[2], VENTTYPE.Value) then
            begin
            Result := False;
            break;
            end
          else
            begin
            Lst.Add(Q.FindField('Y_GENERAL').AsString);
            end;
          end;
        END ;
      END ;
  finally
    Lst.Free;
    RestrictAna.Free;
    end;
  {e FP 29/12/2005}
END ;

end.
