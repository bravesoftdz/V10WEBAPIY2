unit ReImput ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, Grids, DBGrids, StdCtrls, Hctrls,
  ExtCtrls, ComCtrls, Buttons, Mask, Hcompte, HEnt1, Ent1, SaisComm, SaisUtil,
  hmsgbox, Saisie, HStatus, LetBatch, DelVisuE, AboUtil, HRichEdt, HSysMenu, ValPerio,
  HDB, HTB97, ed_tools, ColMemo, LettUtil, Utilsais, UtilSoc, HRichOLE
  ,HPanel, UIUtil, ADODB // MODIF PACK AVANCE pour gestion mode inside
  ;

procedure Reimputation ;

type
  TFReImput = class(TFMul)
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    TE_DEVISE: THLabel;
    E_DEVISE: THValComboBox;
    HLabel2: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    TE_DATEECHEANCE: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    TE_DATEECHEANCE2: THLabel;
    E_NUMECHE: THCritMaskEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    XX_WHERE: TEdit;
    XX_WHERE2: TEdit;
    E_ECRANOUVEAU: THCritMaskEdit;
    XX_WHERE3: TEdit;
    PParam: TTabSheet;
    Bevel5: TBevel;
    HLabel3: THLabel;
    CONTRE_GENE: THCpteEdit;
    HLabel4: THLabel;
    CONTRE_AUXI: THCpteEdit;
    HLabel5: THLabel;
    CONTRE_JOURNAL: THValComboBox;
    HLabel6: THLabel;
    CONTRE_DATE: THCritMaskEdit;
    HM: THMsgBox;
    Label2: TLabel;
    CONTRE_NEGATIF: TCheckBox;
    BListePIECES: TToolbarButton97;
    BZoom: TToolbarButton97;
    E_TRESOLETTRE: THCritMaskEdit;
    TE_REFINTERNE: THLabel;
    E_REFINTERNE: TEdit;
    TE_GENERAL: THLabel;
    E_GENERAL: THCpteEdit;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    HLabel7: THLabel;
    CONTRE_ETABLISSEMENT: THValComboBox;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    CModeS: THValComboBox;
    HModeS: THLabel;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);  override;
    procedure BOuvrirClick(Sender: TObject);   override;
    procedure E_EXERCICEChange(Sender: TObject);
    procedure BChercheClick(Sender: TObject);  override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BListePIECESClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);  override;
    procedure BZoomClick(Sender: TObject);
    procedure CONTRE_JOURNALExit(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
  private
    GeneDate,NowFutur : TDateTime ;
    ListeSel,ListeAna,TPIECE : TList ;
    MemeVentil : boolean ;
    MultiDEV,EncON   : boolean ;
    Gen2Vent   : Array[1..MaxAxe] of boolean ;
    function  ComptesOK : boolean ;
    function  DateOK : boolean ;
    function  CoherenceOK : boolean ;
    function  SoucheOK : boolean ;
    procedure ChargeSelection ;
    function  JePeuxValider : boolean ;
    procedure TripoteO ( O : TOBM ; M : RMVT ; EcrG : boolean ) ;
    procedure TraiteLesAna ( TAna : TDataSet ; OEcr : TOBM ; NewNumL : integer ; OldM,NewM : RMVT ) ;
    procedure ReimputeAna ( TAna : TDataSet ) ;
    procedure ContrePasseAna ( TAna : TDataSet ) ;
    procedure Reimpute ( TEcr : TDataSet ) ;
    procedure ContrePasse ( TEcr : TDataSet ) ;
    procedure GenereLesPieces ;
  public
  end;

implementation

Uses UtilPgi ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 09/09/2003
Modifié le ... :   /  /    
Description .. : 
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion 
Suite ........ : mode inside
Mots clefs ... : 
*****************************************************************}
procedure Reimputation ;
Var X : TFReImput ;
    PP : THPanel ;
BEGIN
  if PasCreerDate(V_PGI.DateEntree) then Exit ;
  if _Blocage(['nrBatch','nrCloture'],True,'nrBatch') then Exit ;

  X:=TFReImput.Create(Application) ;
  X.FNomFiltre:='REIMPUT' ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    try
      X.ShowModal ;
      finally
      X.Free ;
      end ;
    end
  else
    begin
    InitInside(X,PP) ;
    X.Show ;
    end ;

  _Bloqueur('nrBatch',False) ;
  Screen.Cursor:=SyncrDefault ;
END ;

{$R *.DFM}

procedure TFReImput.FormShow(Sender: TObject);
begin
E_DEVISE.Value:=V_PGI.DevisePivot ; CModeS.Value:=V_PGI.DevisePivot ;
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
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
CONTRE_DATE.Text:=DateToStr(V_PGI.DateEntree) ;
if E_JOURNAL.Items.Count>0 then
   BEGIN
   if E_JOURNAL.Vide then E_JOURNAL.ItemIndex:=1 else E_JOURNAL.ItemIndex:=0 ;
   END ;
  inherited;
Pages.Pages[2].TabVisible:=FALSE ; Pages.Pages[3].TabVisible:=FALSE ;
if Not VH^.MontantNegatif then BEGIN CONTRE_NEGATIF.Checked:=False ; CONTRE_NEGATIF.Enabled:=False ; END ;
CONTRE_ETABLISSEMENT.Value:=VH^.EtablisDefaut ; MultiDEV:=False ;
Q.Manuel:=FALSE ;
CentreDBGRid(FListe) ;
end;

procedure TFReImput.FormCreate(Sender: TObject);
begin
  inherited;
RempliSelectEuro(CModeS,HM.Mess[25]) ;
ListeSel:=TList.Create ; ListeAna:=TList.Create ; TPIECE:=TList.Create ;
Q.Manuel:=True ; Q.Liste:='REIMPUTATION' ; EncON:=False ;
end;

function TFReImput.ComptesOK : boolean ;
Var Gen1,Gen2 : TGGeneral ;
    Aux1,Aux2 : TGTiers ;
    O1,OLOC   : TOBM ;
    Okok,OkVent,OkLocal : boolean ;
    k,ErrLoc  : integer ;
    LaDevise  : String3 ;
    SAJAL     : TSAJAL ;
BEGIN
Result:=False ;
Aux1:=Nil ; Aux2:=Nil ; SAJAL:=Nil ;
MemeVentil:=False ; FillChar(Gen2Vent,Sizeof(Gen2Vent),#0) ; OkVent:=False ;
{Présences obligatoires}
if CONTRE_GENE.Text='' then BEGIN HM.Execute(9,'','') ; Exit ; END ;
if CONTRE_GENE.ExisteH<=0 then BEGIN HM.Execute(9,'','') ; Exit ; END ;
if CONTRE_AUXI.Text<>'' then
   BEGIN
   if CONTRE_AUXI.ExisteH<=0 then BEGIN HM.Execute(15,'','') ; Exit ; END
                             else Aux2:=TGTiers.Create(CONTRE_AUXI.Text) ;
   END ;
O1:=TOBM(ListeSel[0]) ;
if O1.GetMvt('E_AUXILIAIRE')<>'' then Aux1:=TGTiers.Create(O1.GetMvt('E_AUXILIAIRE')) ;
Gen1:=TGGeneral.Create(O1.GetMvt('E_GENERAL')) ; Gen2:=TGGeneral.Create(CONTRE_GENE.Text) ;
if Gen2<>Nil then EncON:=Gen2.TvaSurEncaissement ;
if CONTRE_JOURNAL.Value<>'' then SAJAL:=TSAJAL.Create(CONTRE_JOURNAL.Value,False) ;
Okok:=True ;
{Cohérence comptes de sortie}
OkLocal:=True ; ErrLoc:=0 ;
if ((Gen2<>Nil) and (Aux2<>Nil)) then
   BEGIN
   if Not NatureGenAuxOk(Gen2.NatureGene,Aux2.NatureAux) then BEGIN ErrLoc:=18 ; OkLocal:=False ; END ;
   END ;
if ((SAJAL<>Nil) and (Gen2<>Nil)) then if EstInterdit(SAJAL.COMPTEINTERDIT,Gen2.General,0)>0 then
   BEGIN
   ErrLoc:=19 ; OkLocal:=False ;
   END ;
if ((Gen2<>Nil) and (ListeSel.Count>0)) then
   BEGIN
   OLOC:=TOBM(ListeSel[0]) ; LaDevise:=OLOC.GetMvt('E_DEVISE') ;
   if Gen2.NatureGene='BQE' then if PasDeviseBanque(Gen2.General,LaDevise) then BEGIN ErrLoc:=20 ; OkLocal:=False ; END ;
   END ;
if Gen2<>Nil then if Not NaturePieceCompteOk('OD',Gen2.NatureGene) then BEGIN ErrLoc:=21 ; OkLocal:=False ; END ;
if Not OkLocal then
   BEGIN
   Gen1.Free ; Gen2.Free ; Aux1.Free ; Aux2.Free ; SAJAL.Free ;
   HM.Execute(ErrLoc,'','') ; Exit ;
   END ;
{Cohérence 1 et 2}
if ((Gen1=Nil) or (Gen2=Nil)) then Okok:=False else
if ((Aux1<>Nil) and (Aux2=Nil)) or ((Aux1=Nil) and (Aux2<>Nil)) then Okok:=False else
if Gen1.Lettrable<>Gen2.Lettrable then Okok:=False else
if Gen1.Collectif<>Gen2.Collectif then Okok:=False else
if ((Gen1.Collectif) and (Aux1=Nil)) or ((Gen2.Collectif) and (Aux2=Nil)) then Okok:=False else
if ((Not Gen1.Collectif) and (Aux1<>Nil)) or ((Not Gen2.Collectif) and (Aux2<>Nil)) then Okok:=False else
   BEGIN
   if ((Aux1<>Nil) and (Aux2<>Nil)) then
      BEGIN
      if Aux1.Lettrable<>Aux2.Lettrable then Okok:=False else
       if ((Gen1.General=Gen2.General) and (Aux1.Auxi=Aux2.Auxi)) then Okok:=False ;
      END ;
   if Okok then
      BEGIN
      if Gen1.General=Gen2.General then
         BEGIN
         if ((Aux1=Nil) and (Aux2=Nil)) or ((Aux1.Auxi=Aux2.Auxi)) then Okok:=False ;
         END ;
      MemeVentil:=True ;
      for k:=1 to 5 do
          BEGIN
          if Gen1.Ventilable[k]<>Gen2.Ventilable[k] then MemeVentil:=False ;
          Gen2Vent[k]:=Gen2.Ventilable[k] ; if Gen1.Ventilable[k] then OkVent:=True ;
          END ;
      if Not OKVent then MemeVentil:=False ;
      END ;
   END ;
{Fin tests}
Result:=Okok ;
Gen1.Free ; Gen2.Free ; Aux1.Free ; Aux2.Free ; SAJAL.Free ;
if Not Okok then HM.Execute(16,'','') ;
END ;

function TFReImput.DateOK : boolean ;
Var Err : integer ;
BEGIN
Result:=False ;
GeneDate:=0 ;
Err:=ControleDate(CONTRE_DATE.Text) ;
if Err>0 then BEGIN HM.Execute(3+Err,'','') ; Exit ; END ;
GeneDate:=StrToDate(CONTRE_DATE.Text) ;
Result:=True ;
END ;

procedure TFReImput.ChargeSelection ;
Var i : integer ;
    M : RMVT ;
BEGIN
InitMove(FListe.NbSelected,'') ; VideListe(ListeSel) ;
for i:=0 to FListe.NbSelected-1 do
    BEGIN
    MoveCur(False) ; FListe.GotoLeBookMark(i) ;
    if TrouveSaisie(Q,M,E_QUALIFPIECE.Text) then
       BEGIN
       M.NumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ;
       RemplirOFromM(M,ListeSel) ;
       END ;
    END ;
FiniMove ; Screen.Cursor:=SyncrDefault ;
END ;

function TFReImput.CoherenceOK : boolean ;
Var i,Err : integer ;
    O : TOBM ;
    OAux,OGene,OEtab,ODevise : string ;
    OTaux            : Double ;
BEGIN
Err:=0 ; OTaux:=0 ;
ChargeSelection ;
for i:=0 to ListeSel.Count-1 do
    BEGIN
    O:=TOBM(ListeSel[i]) ;
    if i=0 then
       BEGIN
       OGene:=O.GetMvt('E_GENERAL') ; OAux:=O.GetMvt('E_AUXILIAIRE') ;
       OEtab:=O.GetMvt('E_ETABLISSEMENT') ; ODevise:=O.GetMvt('E_DEVISE') ;
       OTaux:=O.GetMvt('E_TAUXDEV') ;
       if ((ODevise<>V_PGI.DevisePivot) and (Not MultiDEV)) then Err:=22 ;
       END else
       BEGIN
       if OGene<>O.GetMvt('E_GENERAL') then Err:=1 else
        if OAux<>O.GetMvt('E_AUXILIAIRE') then Err:=2 else
         if OEtab<>O.GetMvt('E_ETABLISSEMENT') then Err:=3 else
          if ODevise<>O.GetMvt('E_DEVISE') then Err:=4 else
           if OTaux<>O.GetMvt('E_TAUXDEV') then Err:=5 ;
       END ;
    if Err>0 then Break ;
    END ;
Result:=(Err<=0) ;
if Err>0 then BEGIN if Err<20 then HM.Execute(9+Err,'','') else HM.Execute(Err,'','') ; END ;  
END ;

function TFReImput.SoucheOK : boolean ;
Var Q : TQuery ;
BEGIN
Result:=False ;
if CONTRE_JOURNAL.Value='' then Exit ;
Q:=OpenSQL('Select J_COMPTEURNORMAL from Journal Where J_JOURNAL="'+CONTRE_JOURNAL.Value+'"',True) ;
if Not Q.EOF then
   BEGIN
   if Q.Fields[0].AsString<>'' then Result:=True ;
   END ;
Ferme(Q) ;
if Not Result then HM.Execute(23,'','') ; 
END ;

function TFReImput.JePeuxValider : boolean ;
BEGIN
Result:=False ;
if FListe.NbSelected<=0 then BEGIN HM.Execute(3,'','') ; Exit ; END ;
if CONTRE_JOURNAL.Value='' then BEGIN HM.Execute(0,'','') ; Exit ; END ;
if Not SoucheOK then Exit ;
if Not DATEOK then Exit ;
if Not CoherenceOK then Exit ;
if Not ComptesOK then Exit ;
Result:=True ;
END ;

procedure TFReImput.TripoteO ( O : TOBM ; M : RMVT ; EcrG : Boolean ) ;
BEGIN
if EcrG then
   BEGIN
   O.PutMvt('E_JOURNAL',M.Jal) ; O.PutMvt('E_DATECREATION',Date) ; O.PutMvt('E_DATEMODIF',NowFutur) ;
   O.PutMvt('E_EXERCICE',M.Exo) ; O.PutMvt('E_DATECOMPTABLE',M.DateC) ;
{$IFNDEF SPEC302}
   O.PutMvt('E_PERIODE',GetPeriode(M.DateC)) ; O.PutMvt('E_SEMAINE',NumSemaine(M.DateC)) ;
{$ENDIF}
   O.PutMvt('E_NUMEROPIECE',M.Num) ; O.PutMvt('E_QUALIFPIECE',M.Simul) ; O.PutMvt('E_NATUREPIECE',M.Nature) ;
   O.PutMvt('E_DEVISE',M.CodeD) ; O.PutMvt('E_TAUXDEV',M.TauxD) ; O.PutMvt('E_DATETAUXDEV',M.DateTaux) ;
   O.PutMvt('E_ETABLISSEMENT',M.Etabl) ; O.PutMvt('E_QUALIFORIGINE','RIM') ;
   O.PutMvt('E_REFRELEVE','') ; O.PutMvt('E_FLAGECR','') ; O.PutMvt('E_TRACE','') ;
   O.PutMvt('E_ETAT','0000000000') ; O.PutMvt('E_NIVEAURELANCE',0) ;
   O.PutMvt('E_DATERELANCE',IDate1900) ; O.PutMvt('E_VALIDE','-') ;
   O.PutMvt('E_SUIVDEC','')       ; O.PutMvt('E_NOMLOT','') ;
   O.PutMvt('E_FLAGECR','')       ; O.PutMvt('E_CONTROLETVA','') ;
   O.PutMvt('E_EDITEETATTVA','-') ; O.PutMvt('E_NIVEAURELANCE',0) ;
   END else
   BEGIN
   O.PutMvt('Y_JOURNAL',M.Jal) ; O.PutMvt('Y_DATECREATION',Date) ; O.PutMvt('Y_DATEMODIF',NowFutur) ;
   O.PutMvt('Y_EXERCICE',M.Exo) ; O.PutMvt('Y_DATECOMPTABLE',M.DateC) ;
{$IFNDEF SPEC302}
   O.PutMvt('Y_PERIODE',GetPeriode(M.DateC)) ; O.PutMvt('Y_SEMAINE',NumSemaine(M.DateC)) ;
{$ENDIF}
   O.PutMvt('Y_NUMEROPIECE',M.Num) ; O.PutMvt('Y_QUALIFPIECE',M.Simul) ; O.PutMvt('Y_NATUREPIECE',M.Nature) ;
   O.PutMvt('Y_DEVISE',M.CodeD) ; O.PutMvt('Y_TAUXDEV',M.TauxD) ; O.PutMvt('Y_DATETAUXDEV',M.DateTaux) ;
   O.PutMvt('Y_ETABLISSEMENT',M.Etabl) ; O.PutMvt('Y_TRACE','') ;
   O.PutMvt('Y_VALIDE','-') ;
   END ;
END ;

procedure TFReImput.ContrePasse ( TEcr : TDataSet ) ;
BEGIN
TEcr.FindField('E_GENERAL').AsString:=CONTRE_GENE.Text ;
TEcr.FindField('E_AUXILIAIRE').AsString:=CONTRE_AUXI.Text ;
if EncON then TEcr.FindField('E_TVAENCAISSEMENT').AsString:='X'
         else TEcr.FindField('E_TVAENCAISSEMENT').AsString:='-' ; 
END ;

procedure TFReImput.ContrePasseAna ( TAna : TDataSet ) ;
BEGIN
TAna.FindField('Y_GENERAL').AsString:=CONTRE_GENE.Text ;
END ;

procedure TFReImput.ReimputeAna ( TAna : TDataSet ) ;
Var X : Double ;
BEGIN
if Contre_Negatif.Checked then
   BEGIN
   TAna.FindField('Y_DEBIT').AsFloat:=-TAna.FindField('Y_DEBIT').AsFloat ;
   TAna.FindField('Y_CREDIT').AsFloat:=-TAna.FindField('Y_CREDIT').AsFloat ;
   TAna.FindField('Y_DEBITDEV').AsFloat:=-TAna.FindField('Y_DEBITDEV').AsFloat ;
   TAna.FindField('Y_CREDITDEV').AsFloat:=-TAna.FindField('Y_CREDITDEV').AsFloat ;
   TAna.FindField('Y_QTE1').AsFloat:=-TAna.FindField('Y_QTE1').AsFloat ;
   TAna.FindField('Y_QTE2').AsFloat:=-TAna.FindField('Y_QTE2').AsFloat ;
   TAna.FindField('Y_TOTALQTE1').AsFloat:=-TAna.FindField('Y_TOTALQTE1').AsFloat ;
   TAna.FindField('Y_TOTALQTE2').AsFloat:=-TAna.FindField('Y_TOTALQTE2').AsFloat ;
   TAna.FindField('Y_TOTALECRITURE').AsFloat:=-TAna.FindField('Y_TOTALECRITURE').AsFloat ;
   TAna.FindField('Y_TOTALDEVISE').AsFloat:=-TAna.FindField('Y_TOTALDEVISE').AsFloat ;
   END else
   BEGIN
   X:=TAna.FindField('Y_DEBIT').AsFloat ; TAna.FindField('Y_DEBIT').AsFloat:=TAna.FindField('Y_CREDIT').AsFloat ; TAna.FindField('Y_CREDIT').AsFloat:=X ;
   X:=TAna.FindField('Y_DEBITDEV').AsFloat ; TAna.FindField('Y_DEBITDEV').AsFloat:=TAna.FindField('Y_CREDITDEV').AsFloat ; TAna.FindField('Y_CREDITDEV').AsFloat:=X ;
   END ;
END ;

procedure TFReImput.TraiteLesAna ( TAna : TDataSet ; OEcr : TOBM ; NewNumL : integer ; OldM,NewM : RMVT) ;
Var QAna  : TQuery ;
    NumLE,i,k : integer ;
    OA    : TOBM ;
    Ax    : String ;
    DEV   : RDEVISE ;
BEGIN
VideListe(ListeAna) ; NumLE:=OEcr.GetMvt('E_NUMLIGNE') ;
QAna:=OpenSQL('Select * from Analytiq Where '+WhereEcriture(tsAnal,OldM,False)+' AND Y_NUMLIGNE='+InttoStr(NumLE),True) ;
While Not QAna.EOF do
   BEGIN
   Ax:=QAna.FindField('Y_AXE').AsString ;
   OA:=TOBM.Create(EcrAna,Ax,False) ; OA.ChargeMvt(QAna) ; ListeAna.Add(OA) ;
   QAna.Next ;
   END ;
Ferme(QAna) ;
for i:=0 to ListeAna.Count-1 do
    BEGIN
    OA:=TOBM(ListeAna[i]) ; TripoteO(OA,NewM,False) ;
    {Anal reimpute}
    TAna.Insert ;
    OA.EgalChamps(TAna) ; ReimputeAna(TAna) ; TAna.FindField('Y_NUMLIGNE').AsInteger:=NewNumL ;
    TAna.Post ;
    if MemeVentil then
       BEGIN
       {Anal contrepasse}
       TAna.Insert ;
       OA.EgalChamps(TAna) ; ContrePasseAna(TAna) ; TAna.FindField('Y_NUMLIGNE').AsInteger:=NewNumL+1 ;
       TAna.Post ;
       END ;
    END ;
if MemeVentil then Exit ;
{anal sur section attente}
FillChar(DEV,Sizeof(DEV),#0) ; VideListe(ListeAna) ;
DEV.Code:=OEcr.GetMvt('E_DEVISE') ; GetInfosDevise(DEV) ;
DEV.DateTaux:=GeneDate ; DEV.Taux:=OEcr.GetMvt('E_DATETAUXDEV') ;
for i:=1 to MaxAxe do if Gen2Vent[i] then
    BEGIN
    Ax:='A'+Inttostr(i) ;
    VentileGenerale(Ax,OEcr,DEV,ListeAna,False) ;
    for k:=0 to ListeAna.Count-1 do
        BEGIN
        OA:=TOBM(ListeAna[k]) ; TripoteO(OA,NewM,False) ;
        TAna.Insert ;
        OA.EgalChamps(TAna) ; ContrePasseAna(TAna) ; TAna.FindField('Y_NUMLIGNE').AsInteger:=NewNumL+1 ;
        TAna.Post ;
        END ;
    END ;
END ;

procedure TFReImput.Reimpute ( TEcr : TDataSet ) ;
Var X : Double ;
BEGIN
if Contre_Negatif.Checked then
   BEGIN
   TEcr.FindField('E_DEBIT').AsFloat:=-TEcr.FindField('E_DEBIT').AsFloat ;
   TEcr.FindField('E_CREDIT').AsFloat:=-TEcr.FindField('E_CREDIT').AsFloat ;
   TEcr.FindField('E_DEBITDEV').AsFloat:=-TEcr.FindField('E_DEBITDEV').AsFloat ;
   TEcr.FindField('E_CREDITDEV').AsFloat:=-TEcr.FindField('E_CREDITDEV').AsFloat ;
   TEcr.FindField('E_QTE1').AsFloat:=-TEcr.FindField('E_QTE1').AsFloat ;
   TEcr.FindField('E_QTE2').AsFloat:=-TEcr.FindField('E_QTE2').AsFloat ;
   END else
   BEGIN
   X:=TEcr.FindField('E_DEBIT').AsFloat ; TEcr.FindField('E_DEBIT').AsFloat:=TEcr.FindField('E_CREDIT').AsFloat ; TEcr.FindField('E_CREDIT').AsFloat:=X ;
   X:=TEcr.FindField('E_DEBITDEV').AsFloat ; TEcr.FindField('E_DEBITDEV').AsFloat:=TEcr.FindField('E_CREDITDEV').AsFloat ; TEcr.FindField('E_CREDITDEV').AsFloat:=X ;
   END ;
END ;

procedure TFReImput.GenereLesPieces ;
Var NewM,OldM   : RMVT ;
    O1,O,OHisto : TOBM ;
    i,NumL,k    : integer ;
    TEcr,TAna   : TQUERY ;
    Premier,OkAna : boolean ;
BEGIN
O1:=TOBM(ListeSel[0]) ; OHisto:=Nil ;
{Nouvelle pièce}
FillChar(NewM,Sizeof(NewM),#0) ;
NewM.Jal:=CONTRE_JOURNAL.Value ;
NewM.Exo:=QuelExoDT(GeneDate) ; NewM.DateC:=GeneDate ;
NewM.Num:=GetNewNumJal(NewM.Jal,False,GeneDate) ; if NewM.Num<=0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
NewM.Simul:='R' ; NewM.Nature:='OD' ;
NewM.CodeD:=O1.GetMvt('E_DEVISE') ; NewM.TauxD:=O1.GetMvt('E_TAUXDEV') ; NewM.DateTaux:=GeneDate ;
NewM.Etabl:=CONTRE_ETABLISSEMENT.Value ;
{Ecritures}
TEcr:=OpenSQL('SELECT * FROM ECRITURE WHERE E_GENERAL="ErE"',False) ;
TAna:=OpenSQL('SELECT * FROM ANALYTIQ WHERE Y_SECTION="ErE"',False) ;
InitMove(ListeSel.Count,'') ; NumL:=-1 ;
for i:=0 to ListeSel.Count-1 do
    BEGIN
    MoveCur(False) ; O:=TOBM(ListeSel[i]) ; OldM:=OBMToIdent(O,False) ;
    TripoteO(O,NewM,True) ; Premier:=(O.GetMvt('E_NUMECHE')<=1) ;
    {Ligne de ré-imputation}
    TEcr.Insert ; if Premier then Inc(NumL,2) ;
    O.SetCotation(0) ;
    O.SetMPACC ;
    O.EgalChamps(TEcr) ;
    Reimpute(TEcr) ; TEcr.FindField('E_NUMLIGNE').AsInteger:=NumL ;
    // GG COM
    TEcr.FindField('E_IO').AsString:='X' ;
    TraiteLesAna(TAna,O,NumL,OldM,NewM) ;
    TEcr.Post ;
    {Ligne de contrepassation}
    OkAna:=False ; for k:=1 to MaxAxe do if Gen2Vent[k] then OkAna:=True ;
    TEcr.Insert ;
    O.SetCotation(0) ;
    O.SetMPACC ;
    O.EgalChamps(TEcr) ; ContrePasse(TEcr) ; TEcr.FindField('E_NUMLIGNE').AsInteger:=NumL+1 ;
    if OkAna then TEcr.FindField('E_ANA').AsString:='X 'else TEcr.FindField('E_ANA').AsString:='-' ;
    TEcr.Post ;
    if ((Premier) and (OHisto=Nil)) then
       BEGIN
       OHisto:=TOBM.Create(EcrGen,'',False) ; OHisto.ChargeMvt(TEcr) ; TPIECE.Add(OHisto) ;
       END ;
    END ;
Ferme(TEcr) ; Ferme(TAna) ;
ADevalider(NewM.Jal,GeneDate) ;
MarquerPublifi(True) ;
CPStatutDossier ;
FiniMove ;
END ;

procedure TFReImput.BOuvrirClick(Sender: TObject);
begin
  inherited;
MemeVentil:=False ;
if Not JePeuxValider then BEGIN Screen.Cursor:=SyncrDefault ; Exit ; END ;
if HM.Execute(1,'','')<>mrYes then Exit ;
Application.ProcessMessages ; NowFutur:=NowH ;
if Transactions(GenereLesPieces,3)<>oeOK then MessageAlerte(HM.Mess[17]) else
   BEGIN
   if TPiece.Count>0 then
      BEGIN
      if HM.Execute(24,'','')=mrYes then VisuPiecesGenere(TPiece,EcrGen,8) ;
      END ;
   END ;
end;

procedure TFReImput.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFReImput.BChercheClick(Sender: TObject);
begin
VideListe(ListeSel) ;
  inherited;
GereSelectionsGrid(FListe,Q) ;
end;

procedure TFReImput.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
VideListe(ListeSel) ; ListeSel.Free ;
VideListe(ListeAna) ; ListeAna.Free ;
VideListe(TPIECE) ; TPIECE.Free ;
end;

procedure TFReImput.BListePIECESClick(Sender: TObject);
begin
  inherited;
if TPiece.Count>0 then VisuPiecesGenere(TPiece,EcrGen,8) ;
end;

procedure TFReImput.FListeDblClick(Sender: TObject);
begin
  inherited;
TrouveEtLanceSaisie(Q,taConsult,E_QUALIFPIECE.Text) ;
end;

procedure TFReImput.BZoomClick(Sender: TObject);
begin
  inherited;
FListeDBLClick(Nil) ;
end;

procedure TFReImput.CONTRE_JOURNALExit(Sender: TObject);
Var Q : TQuery ;
begin
//  inherited;
if CONTRE_JOURNAL.Value='' then Exit ; MultiDEV:=False ;
Q:=OpenSQL('Select J_MULTIDEVISE from JOURNAl Where J_JOURNAL="'+CONTRE_JOURNAL.Value+'"',True) ;
if Not Q.EOF then MultiDEV:=(Q.Fields[0].AsString='X') ;
Ferme(Q) ;
end;

procedure TFReImput.E_DEVISEChange(Sender: TObject);
begin
  inherited;
if E_DEVISE.Value=V_PGI.DevisePivot then CModeS.Visible:=True else
   BEGIN
   CModeS.Visible:=False ; CModeS.Value:='MIX' ;
   END ;
end;

end.
