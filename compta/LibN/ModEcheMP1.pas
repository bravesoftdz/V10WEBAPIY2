unit ModEcheMP1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, Hctrls, Mask, ComCtrls, DB,
  DBTables, Hqry, ParamDBG, PrintDBG, Ent1, HEnt1, Paramdat,
  Saisutil, Saiscomm, Saisie, hmsgbox, Filtre, HDB, Menus,
  Hcompte, HSysMenu, HTB97, HPanel, UiUtil,EcheMPAMP,EcheMPA, EcheUnit,
  Mul, HRichOLE, HRichEdt, ColMemo ;

procedure ModifEcheMP1 ;

type
  TFMODIFECHEMP1 = class(TFMul)
    HLabel4: THLabel;
    HLabel1: THLabel;
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    E_AUXILIAIRE: THCpteEdit;
    E_NUMECHE: THCritMaskEdit;
    XX_WHERE: TEdit;
    E_ECHE: TEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHEREPOP: TEdit;
    E_GENERAL: THCpteEdit;
    TE_EXERCICE: THLabel;
    HLabel10: THLabel;
    HLabel3: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    HLabel6: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    HLabel7: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    E_EXERCICE: THValComboBox;
    Label1: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    Label2: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    E_NATUREPIECE: THValComboBox;
    TE_NATUREPIECE: THLabel;
    TE_QUALIFPIECE: THLabel;
    E_QUALIFPIECE: THValComboBox;
    TMP_CODEACCEPT: THLabel;
    MP_CODEACCEPT: THMultiValComboBox;
    Label12: THLabel;
    E_REFINTERNE: THCritMaskEdit;
    HLabel8: THLabel;
    E_DEVISE: THValComboBox;
    PModifS: TTabSheet;
    Bevel5: TBevel;
    Label3: TLabel;
    CMSMP: TCheckBox;
    CMSDateEche: TCheckBox;
    CMSRIB: TCheckBox;
    RIBNEW: TEdit;
    ZoomRib: TToolbarButton97;
    DateEcheNew: THCritMaskEdit;
    MPNew: THValComboBox;
    Label4: TLabel;
    HME: THMsgBox;
    BZoomPiece: TToolbarButton97;
    BModifSerie: TToolbarButton97;
    procedure E_AUXILIAIREChange(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure CMSMPClick(Sender: TObject);
    procedure CMSDateEcheClick(Sender: TObject);
    procedure CMSRIBClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RIBNEWKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ZoomRibClick(Sender: TObject);
    procedure BModifSerieClick(Sender: TObject);
  private
    { Déclarations privées }
    LastQualif  : String ;
    procedure ClickModifRIB ;
    procedure ClickModifMPA ;
    procedure ZoomSurRib ;
    procedure ClickPourMS ;
    procedure ModifSerieChamp ;
    procedure ModifUneEche ;
    procedure InitMulInteractif ;
  public
    { Déclarations publiques }
    ClickSurBChercher : Boolean ;
  end;

implementation

{$R *.DFM}

procedure ModifEcheMP1 ;
Var X  : TFModifEcheMP1 ;
    PP : THPanel ;
BEGIN
if Blocage(['nrCloture','nrBatch'],True,'nrSaisieModif') then Exit ;
X:=TFModifEcheMP1.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
     Bloqueur('nrSaisieModif',False) ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;


procedure TFMODIFECHEMP1.E_AUXILIAIREChange(Sender: TObject);
begin
  inherited;
CMSRIB.Enabled:=Trim(E_AUXILIAIRE.Text)<>'' ; ClickPourMS ;
end;

procedure TFMODIFECHEMP1.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ; 
end;

procedure TFMODIFECHEMP1.CMSMPClick(Sender: TObject);
begin
  inherited;
ClickPourMS ;
end;

procedure TFMODIFECHEMP1.CMSDateEcheClick(Sender: TObject);
begin
  inherited;
ClickPourMS ;
end;

procedure TFMODIFECHEMP1.CMSRIBClick(Sender: TObject);
begin
  inherited;
ClickPourMS ;
end;

procedure TFMODIFECHEMP1.InitMulInteractif ;
Var OkOk : Boolean ;
    St : String ;
BEGIN
XX_WHEREPOP.Text:='' ; VH^.MPModifFaite:=FALSE ;
OkOk:=(VH^.MPPop.MPAuxPop<>'') Or ((VH^.MPPop.MPAuxPop='') And (VH^.MPPop.MPNumEPop>0));
If OkOk Then
  BEGIN
  If VH^.MPPop.MPGenPop<>'' Then E_GENERAL.Text:=VH^.MPPop.MPGenPop ;
  If VH^.MPPop.MPAuxPop<>'' Then E_AUXILIAIRE.Text:=VH^.MPPop.MPAuxPop ;
  If VH^.MPPop.MPJalPop<>'' Then E_JOURNAL.Value:=VH^.MPPop.MPJalPop ;
  If VH^.MPPop.MPExoPop<>'' Then E_EXERCICE.Value:=VH^.MPPop.MPExoPop ;
  If VH^.MPPop.MPNumPop<>0 Then BEGIN E_NUMEROPIECE.Text:=IntToStr(VH^.MPPop.MPNumPop) ; E_NUMEROPIECE_.Text:=IntToStr(VH^.MPPop.MPNumPop) ; END ;
  If VH^.MPPop.MPDatePop<>0 Then BEGIN E_DATECOMPTABLE.Text:=DateToStr(VH^.MPPop.MPDatePop) ; E_DATECOMPTABLE_.Text:=DateToStr(VH^.MPPop.MPDatePop) ; END ;
  If VH^.MPPop.MPNumLPop<>0 Then XX_WHEREPOP.Text:=XX_WHEREPOP.Text+' AND E_NUMLIGNE='+IntToStr(VH^.MPPop.MPNumLPop)+' ' ;
  If VH^.MPPop.MPNumEPop<>0 Then XX_WHEREPOP.Text:=XX_WHEREPOP.Text+' AND E_NUMECHE='+IntToStr(VH^.MPPop.MPNumEPop)+' ' ;
  END ;
If Trim(XX_WHEREPOP.Text)<>'' Then BEGIN St:=XX_WHEREPOP.Text ; delete(St,1,4) ; XX_WHEREPOP.Text:=St ; END ;
END ;



procedure TFMODIFECHEMP1.FormShow(Sender: TObject);
begin
DateEcheNew.Text:=DateToStr(V_PGI.DateEntree) ;
E_QUALIFPIECE.Value:='N' ; LastQualif:='N' ; ClickSurBChercher:=FALSE ;
E_EXERCICE.Value:=VH^.Entree.Code ; E_EXERCICEChange(Nil) ;
E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ; E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
ChargeFiltre('CPMODECHEMP',FFiltres,Pages) ;
if ((E_JOURNAL.Value='') and (E_JOURNAL.Values.Count>0)) then
   BEGIN
   if Not E_JOURNAL.Vide then E_JOURNAL.Value:=E_JOURNAL.Values[0] else
    if E_JOURNAL.Values.Count>1 then E_JOURNAL.Value:=E_JOURNAL.Values[1] ;
   END ;
Q.Liste:='CPMODECHEMP1' ;
InitMulInteractif ;
  inherited;
(*
Q.Manuel:=FALSE ; CentreDBGrid(GL) ;
BChercherClick(Nil) ;
*)
end;

procedure TFMODIFECHEMP1.BZoomPieceClick(Sender: TObject);
begin
  inherited;
TrouveEtLanceSaisie(Q,taConsult,LastQualif) ;
end;

procedure TFMODIFECHEMP1.FListeDblClick(Sender: TObject);
begin
  inherited;
ClickModifMPA ;
end;

procedure TFMODIFECHEMP1.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
VH^.MPPop.MPGenPop:='' ; VH^.MPPop.MPAuxPop:='' ; VH^.MPPop.MPJalPop:='' ; VH^.MPPop.MPExoPop:='' ;
VH^.MPPop.MPNumPop:=0 ; VH^.MPPop.MPNumLPop:=0 ; VH^.MPPop.MPNumEPop:=0 ; VH^.MPPop.MPDatePop:=0 ;
end;

procedure TFMODIFECHEMP1.ClickModifRIB ;
Var M : RMVT ;
    Q1 : TQuery ;
    Trouv : boolean ;
    TAN   : String3 ;
    RIB,Aux,OldRib : String ;
    k     : integer ;
    Coll  : String ;
    RJal,RExo,RQualif : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
begin
if Q.EOF then Exit ;
RJal:=Q.FindField('E_JOURNAL').AsString ;
RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
RQualif:=LastQualif ;
RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_QUALIFPIECE="'+LastQualif+'"'
          +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
          +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString,FALSE) ;
Trouv:=Not Q1.EOF ;
if Trouv then
   BEGIN
   M:=MvtToIdent(Q1,fbGene,True) ;
   RIB:=Q1.FindField('E_RIB').AsString ;
   TAN:=Q1.FindField('E_ECRANOUVEAU').AsString ;
   Aux:=Q1.FindField('E_AUXILIAIRE').AsString ;
   If Aux='' Then Trouv:=FALSE ;
   if TAN='OAN' then
      BEGIN
      if M.CodeD<>V_PGI.DevisePivot then BEGIN HME.Execute(0,'','') ; Trouv:=FALSE ; END ;
      if ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then BEGIN HME.Execute(0,'','') ; Trouv:=FALSE ; END ;
      END ;
   END ;
If Trouv Then
   BEGIN
   OldRIB:=RIB ; ModifLeRIB(Rib,Aux) ;
   If RIB<>OldRib Then
      BEGIN
      Q1.Edit ;
      Q1.FindField('E_RIB').AsString:=RIB ;
      Q1.Post ;
      END ;
   Ferme(Q1) ;
   Application.ProcessMessages ; BChercheClick(Nil) ;
   Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                VarArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[])
   END Else Ferme(Q1) ;
end;

procedure TFMODIFECHEMP1.ClickModifMPA ;
Var M : RMVT ;
    Q1 : TQuery ;
    Trouv,OkModif : boolean ;
    TAN   : String3 ;
    RIB,Aux,OldRib,Gen : String ;
    k     : integer ;
    Coll  : String ;
    RJal,RExo,RQualif : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
    EcheMPA,OldEcheMPA : T_EcheMPAMP ;
    TIDTIC : Boolean ;
begin
if Q.EOF then Exit ;
TIDTIC:=FALSE ;
RJal:=Q.FindField('E_JOURNAL').AsString ;
RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
RQualif:=LastQualif ;
RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_QUALIFPIECE="'+LastQualif+'"'
          +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
          +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString,FALSE) ;
Trouv:=Not Q1.EOF ;
if Trouv then
   BEGIN
   M:=MvtToIdent(Q1,fbGene,True) ;
   RIB:=Q1.FindField('E_RIB').AsString ;
   TAN:=Q1.FindField('E_ECRANOUVEAU').AsString ;
   Aux:=Q1.FindField('E_AUXILIAIRE').AsString ;
   If Aux='' Then
     BEGIN
     Gen:=Q1.FindField('E_GENERAL').AsString ;
     If (Gen<>'') And (RNumEche>0) Then TIDTIC:=TRUE ELSE Trouv:=FALSE ;
     END ;
   if TAN='OAN' then
      BEGIN
      if M.CodeD<>V_PGI.DevisePivot then BEGIN HME.Execute(0,'','') ; Trouv:=FALSE ; END ;
      if ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then BEGIN HME.Execute(0,'','') ; Trouv:=FALSE ; END ;
      END ;
   END ;
If Trouv Then
   BEGIN
   EcheMPA.TIDTIC:=TIDTIC ;
   EcheMPA.DateEche:=Q1.Findfield('E_DATEECHEANCE').AsDateTime ;
   EcheMPA.DateComptable:=Q1.Findfield('E_DATECOMPTABLE').AsDateTime ;
   EcheMPA.ModePaie:=Q1.Findfield('E_MODEPAIE').AsString ;
   EcheMPA.Jal:=Q1.Findfield('E_JOURNAL').AsString ;
   EcheMPA.NatP:=Q1.Findfield('E_NATUREPIECE').AsString ;
   If EcheMPA.TIDTIC Then EcheMPA.Aux:=Gen Else EcheMPA.Aux:=Aux ;
   EcheMPA.Rib:=RIB ;
   EcheMPA.Montant:=Q1.Findfield('E_DEBIT').AsFloat+Q1.Findfield('E_CREDIT').AsFloat ;
   EcheMPA.NumP:=Q1.Findfield('E_NUMEROPIECE').AsInteger ;
   EcheMPA.RefExterne:=Q1.Findfield('E_REFEXTERNE').AsString ;
   EcheMPA.RefLibre:=Q1.Findfield('E_REFLIBRE').AsString ;
   EcheMPA.NumTraChq:=Q1.Findfield('E_NUMTRAITECHQ').AsString ;
   EcheMPA.CodeAcc:=Q1.Findfield('E_CODEACCEPT').AsString ;
   OldEcheMPA:=EcheMPA ;
   OkModif:=FALSE ;
   If ModifUneEcheanceMPAMP(EcheMPA) Then
      BEGIN
      Q1.Edit ;
      If EcheMPA.ModePaie<>OldEcheMPA.ModePaie Then BEGIN Q1.FindField('E_MODEPAIE').AsString:=EcheMPA.MODEPAIE ; OkModif:=TRUE ; END ;
      If EcheMPA.CodeAcc<>OldEcheMPA.CodeAcc Then BEGIN Q1.FindField('E_CODEACCEPT').AsString:=EcheMPA.CodeAcc ; OkModif:=TRUE ; END ;
      If EcheMPA.DateEche<>OldEcheMPA.DateEche Then BEGIN Q1.FindField('E_DATEECHEANCE').AsDateTime:=EcheMPA.DateEche ; OkModif:=TRUE ; END ;
      If EcheMPA.Rib<>OldEcheMPA.Rib Then BEGIN Q1.FindField('E_RIB').AsString:=EcheMPA.Rib ; OkModif:=TRUE ; END ;
      If EcheMPA.RefExterne<>OldEcheMPA.RefExterne Then BEGIN Q1.FindField('E_REFEXTERNE').AsString:=EcheMPA.RefExterne ; OkModif:=TRUE ; END ;
      If EcheMPA.RefLibre<>OldEcheMPA.RefLibre Then BEGIN Q1.FindField('E_REFLIBRE').AsString:=EcheMPA.RefLibre ; OkModif:=TRUE ; END ;
      If EcheMPA.NumTraChq<>OldEcheMPA.NumTraChq Then BEGIN Q1.FindField('E_NUMTRAITECHQ').AsString:=EcheMPA.NumTraChq ; OkModif:=TRUE ; END ;
      Q1.Post ;
      END ;
   Ferme(Q1) ;
   Application.ProcessMessages ; If OkModif Then BEGIN BChercheClick(Nil) ; VH^.MPModifFaite:=TRUE ; END ;
   Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                VarArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[])
   END Else Ferme(Q1) ;
end;



procedure TFMODIFECHEMP1.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : Boolean ;
begin
Inherited ;
Vide:=(Shift=[]) ;
if Vide AND (Key=VK_F5) then BEGIN Key:=0 ; ClickModifMPA ; END ;
if (ssCtrl in Shift) AND (Key=VK_F5) then BEGIN Key:=0 ; ClickModifRIB ; END ;
end;

procedure TFMODIFECHEMP1.ZoomSurRib ;
Var Rib,Aux : String ;
begin
Rib:='' ; Aux:=E_AUXILIAIRE.Text ;
If ModifLeRIB(Rib,Aux) Then RibNew.Text:=Rib ;
end;

procedure TFMODIFECHEMP1.RIBNEWKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : Boolean ;
begin
  inherited;
Vide:=(Shift=[]) ;
if Vide AND (Key=VK_F5) then BEGIN Key:=0 ; ZoomSurRIB ; END ;
Key:=0 ;
end;

procedure TFMODIFECHEMP1.ZoomRibClick(Sender: TObject);
begin
  inherited;
ZoomSurRib ;
end;

procedure TFMODIFECHEMP1.ClickPourMS ;
Var OkMS : Boolean ;
BEGIN
OkMS:=CMSMP.Checked or CMSDateEche.Checked Or CMSRib.Checked ;
BOuvrir.Visible:=Not OkMS ; BModifSerie.Visible:=OkMS ;
MPNew.Enabled:=CMSMP.Checked ; DateEcheNew.Enabled:=CMSDateEche.Checked ;
RIBNew.Enabled:=CMSRib.Checked ;
If OkMS Then
   BEGIN
   END Else
   BEGIN
   END ;
If Not CMSRib.Checked Then RIBNew.Text:='' ;
If Not CMSMP.Checked Then MPNew.ItemIndex:=-1 ;
FListe.MultiSelection:=OkMS ; BSelectAll.Visible:=OkMS ;
END ;


procedure TFMODIFECHEMP1.ModifSerieChamp ;
Var RJal,RExo,RQualif : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
    What,SQL : String ;
BEGIN
RJal:=Q.FindField('E_JOURNAL').AsString ;
RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
RQualif:=LastQualif ;
RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
What:='' ;
If CMSMP.Checked Then What:=What+',E_MODEPAIE="'+MPNew.Value+'" ' ;
If CMSDateEche.Checked Then What:=What+',E_DATEECHEANCE="'+USDATE(DateEcheNew)+'" ' ;
If CMSRIB.Checked Then What:=What+',E_RIB="'+RibNew.Text+'" ' ;
If Trim(What)='' Then Exit ;
What:=Copy(What,2,Length(What)-1) ;
SQL:='UPDATE Ecriture SET '+What
          +'where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_QUALIFPIECE="'+LastQualif+'"'
          +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
          +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString ;
ExecuteSQL(SQL) ;
END ;


procedure TFMODIFECHEMP1.BModifSerieClick(Sender: TObject);
Var BM : TBookmark;
    i : Integer ;
    What,SQL : String ;
begin
If CMSMP.Checked And (MPNEW.Text='') Then BEGIN HME.Execute(2,'','') ; Pages.ActivePage:=PModifS ; MPNew.SetFocus ; Exit ; END ;
If CMSDateEche.Checked And (Not IsValidDate(DateEcheNew.Text) ) Then BEGIN HME.Execute(3,'','') ; Pages.ActivePage:=PModifS ; DateEcheNew.SetFocus ; Exit ;END ;
If CMSRib.Checked And (RibNew.Text='') Then BEGIN HME.Execute(3,'','') ; Pages.ActivePage:=PModifS ; RIBNew.SetFocus ; Exit ; END ;
if (Fliste.NbSelected>0) or (Fliste.AllSelected) then
   BEGIN
//   BM:=NIL ; BM:=Q.GetBookmark;
   if Fliste.AllSelected then
     BEGIN
     Q.DisableControls ;
     Q.First ;
     While not Q.Eof do
       BEGIN
       If Transactions(ModifSerieChamp,2)<>oeOk Then ;
       Q.Next ;
       END ;
     Q.EnableControls ;
     END else
     BEGIN
     for i:=0 to Fliste.NbSelected-1 do
      BEGIN
      Fliste.GotoLeBookMark(i) ;
      If Transactions(ModifSerieChamp,2)<>oeOk Then ;
      END ;
     END ;
//   Q.GotoBookmark(BM); Q.FreeBookmark(BM);
   BChercheClick(NIL) ;
   END ;
end;

procedure TFMODIFECHEMP1.ModifUneEche ;
Var M : RMVT ;
    Q1 : TQuery ;
    Trouv : boolean ;
    EU    : T_ECHEUNIT ;
    TAN   : String3 ;
    k     : integer ;
    Coll  : String ;
    RJal,RExo,RQualif : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
begin
if Q.EOF then Exit ;
RJal:=Q.FindField('E_JOURNAL').AsString ;
RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
RQualif:=LastQualif ;
RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_QUALIFPIECE="'+LastQualif+'"'
          +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
          +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString,True) ;
Trouv:=Not Q1.EOF ;
if Trouv then
   BEGIN
   M:=MvtToIdent(Q1,fbGene,True) ; FillChar(EU,Sizeof(EU),#0) ;
   EU.DateEche:=Q1.FindField('E_DATEECHEANCE').AsDateTime ; EU.ModePaie:=Q1.FindField('E_MODEPAIE').AsString ;
   EU.DebitDEV:=Q1.FindField('E_DEBITDEV').AsFloat ; EU.CreditDEV:=Q1.FindField('E_CREDITDEV').AsFloat ;
   EU.Debit:=Q1.FindField('E_DEBIT').AsFloat ; EU.Credit:=Q1.FindField('E_CREDIT').AsFloat ;
   EU.DebitEuro:=Q1.FindField('E_DEBITEURO').AsFloat ; EU.CreditEuro:=Q1.FindField('E_CREDITEURO').AsFloat ;
   EU.DEVISE:=Q1.FindField('E_DEVISE').AsString ; EU.TauxDEV:=Q1.FindField('E_TAUXDEV').AsFloat ;
   EU.DateComptable:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ;
   EU.DateModif:=Q1.FindField('E_DATEMODIF').AsDateTime ;
   if EuroOK then EU.SaisieEuro:=(Q1.FindField('E_SAISIEEURO').AsString='X')
             else EU.SaisieEuro:=False ;
   TAN:=Q1.FindField('E_ECRANOUVEAU').AsString ;
   {#TVAENC}
   if VH^.OuiTvaEnc then
      BEGIN
      Coll:=Q1.FindField('E_GENERAL').AsString ;
      if EstCollFact(Coll) then
         BEGIN
         for k:=1 to 4 do EU.TabTva[k]:=Q1.FindField('E_ECHEENC'+IntToStr(k)).AsFloat ;
         EU.TabTva[5]:=Q1.FindField('E_ECHEDEBIT').AsFloat ;
         END ;
      END ;
   END ;
Ferme(Q1) ;
if Not Trouv then Exit ;
if TAN='OAN' then
   BEGIN
   if M.CodeD<>V_PGI.DevisePivot then BEGIN HME.Execute(0,'','') ; Exit ; END ;
   if ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then BEGIN HME.Execute(0,'','') ; Exit ; END ;
   END ;
if ModifUneEcheance(M,EU) then
  BEGIN
  Application.ProcessMessages ; BChercheClick(Nil) ;
  Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
               VarArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[])
  END ;
end;


end.
