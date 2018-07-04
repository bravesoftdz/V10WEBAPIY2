{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 30/04/2003
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par CPMULFACTEFF_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit MulFactEff;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB, Hqry,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, Saisie, SaisUtil,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil,
  EncUtil, LettUtil, UTOB, ed_Tools
  , SaisBor, HRichOLE, ADODB
  ;

Function PointeMvtCpt(Gene,Auxi,LibCpt : String ; LOBM : Tlist ; DateEcheDefaut : tDateTime ; NewFlux : Boolean ;
                      ModeSaisie : tModeSaisieEff ; ModeSR : tModeSR; O : TOB) : Boolean ;  // FQ 11665

Type tTotSel = Record
              XDebitP,XCreditP : Double ;
              NbP : Integer ;
              End ;

type
  TFMulFactEff = class(TFMul)
    TE_NUMEROPIECE: THLabel;
    HLabel2: THLabel;
    TE_ETABLISSEMENT: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    E_NUMEROPIECE_: THCritMaskEdit;
    E_ETABLISSEMENT: THValComboBox;
    HLabel4: THLabel;
    E_GENERAL: THCritMaskEdit;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCritMaskEdit;
    HM: THMsgBox;
    Label14: TLabel;
    E_MODEPAIE: THValComboBox;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    TE_DATEECHEANCE2: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    TE_DATEECHEANCE: THLabel;
    TE_DEVISE: THLabel;
    E_DEVISE: THValComboBox;
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    XX_WHEREAN: TEdit;
    E_NUMECHE: THCritMaskEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    E_ECHE: THCritMaskEdit;
    E_ETATLETTRAGE: THCritMaskEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHERENTP: TEdit;
    XX_WHEREMONTANT: TEdit;
    XX_WHERESEL: TEdit;
    WFliste2: TToolbar97;
    FTotSel: TToolbar97;
    FPied: TPanel;
    FSession: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    ISigneEuro1: TImage;
    FNbP: THNumEdit;
    FDebitPA: THNumEdit;
    FCreditPA: THNumEdit;
    FSoldeA: THNumEdit;
    bVoirTotSel: TToolbarButton97;
    Panel1: TPanel;
    Label3: TLabel;
    FListe2: THGrid;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FListeFlipSelection(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure bSelectAllClick(Sender: TObject);
    procedure FListeDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure FListeEndDrag(Sender, Target: TObject; X, Y: Integer);
    procedure FListeMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure Fliste2DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure Fliste2DragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListe2MouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure FListe2EndDrag(Sender, Target: TObject; X, Y: Integer);
    procedure FListeDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure bVoirTotSelClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private    { Déclarations privées }
    Gene,Auxi,LibCpt : String ;
    LOBM,LOBM1 : tList ;
    TOB1,TOB2 : TOB ;  // FQ 11665
    TOBFliste2 : TOB ;
    OkSel : Boolean ;
    (*
    XDebitP,XDebitE,XCreditP,XCreditE : Double ;
    NbP : Integer ;
    *)
    DateEcheDefaut : TDateTime ;
    NewFlux : Boolean ;
    ModeSaisie : tModeSaisieEff ;
    DragFListeEnCours,DragFListe2EnCours : Boolean ;
    OkDropFliste,OkDropFliste2 : Boolean ;
    totSel : tTotSel ;
    TotSelInit : tTotSel ;
    ModeSR : tModeSR ;
    procedure InitCriteres ;
//    procedure InitSelect ;
//    Function  EstSelectionne(Q : TQuery ; LOBM : TList) : Boolean ;
    procedure AfficheTotauxBasEcran ;
    Function  CtrlMontantSel(LOBM : TList) : Boolean ;
//    procedure FaitXX_WHERESEL ;
    procedure FAitOBMDrag ;
    procedure RefreshLOBM ;
    Function  ExistePasDansLOBM(O1 : TOBM ; TOBL1 : TOB) : Boolean ;
    procedure TraiteMajFliste2(Plus : Boolean) ;
    procedure RefreshFliste2 ;
    procedure CalculMontantSelect ;
    Procedure DeleteMvtFliste2 ;
    Procedure DeleteFromLOBM(TOBL : TOB) ;
    function  ExisteDansTOB(O1 : TOBM ; TOBL1 : TOB) : Boolean; // FQ 11665
 public
  end;

implementation

{$R *.DFM}


Uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus, ZoomfactEff ;

Const TypeSel : Byte = 0 ;


Function PointeMvtCpt(Gene,Auxi,LibCpt : String ; LOBM : Tlist ; DateEcheDefaut : tDateTime ; NewFlux : Boolean ;
                      ModeSaisie : tModeSaisieEff ; ModeSR : tModeSR; O : TOB) : Boolean ;
var X : TFMulFactEff ;
    PP : THPanel ;
begin
Result:=FALSE ;
PP:=FindInsidePanel ;
X:=TFMulFactEff.Create(Application) ;
X.Q.Manuel:=True ;
X.gene:=Gene ; X.Auxi:=Auxi ; X.LOBM:=LOBM ; X.LibCpt:=LibCpt ;
X.DateEcheDefaut:=DateEcheDefaut ;
X.NewFlux:=NewFlux ;
X.ModeSaisie:=ModeSaisie ;
X.ModeSR:=ModeSR ;
X.TOB2 := O;
Case X.ModeSaisie Of
  OnEffet : BEGIN
            X.FNomFiltre:='CPFACTCLIEFFT' ; X.Q.Liste:='CPMULFACTCLIEFFT' ;
            END ;
  OnChq   : BEGIN
            X.FNomFiltre:='CPMULFACTCLICHQ' ; X.Q.Liste:='CPMULFACTCLICHQ' ;
            END ;
  OnCB    : BEGIN
            X.FNomFiltre:='CPMULFACTCLICB' ; X.Q.Liste:='CPMULFACTCLICB' ;
            END ;
  OnBqe : BEGIN
            X.FNomFiltre:='CPMULFACTCLIBQE' ; X.Q.Liste:='CPMULFACTCLIBQE' ;
            END ;
  End ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     Result:=X.OkSel ;
     X.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

(*========================= METHODES DE LA FORM ==============================*)
procedure TFMulFactEff.FormCreate(Sender: TObject);
begin
  inherited;
MemoStyle:=msBook ;
TOB1:=TOB.Create('',Nil,-1) ;
TOBFliste2:=TOB.Create('',Nil,-1) ;
LOBM1:=TList.Create ;
RegLoadToolbarPos(Self,'MulFactEff') ;
end;

{Function TFMulFactEff.EstSelectionne(Q : TQuery ; LOBM : TList) : Boolean ;
Var i : Integer ;
//    D,C : Double ;
    O : TOBM ;
BEGIN
Result:=TRUE ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  If (O.GetMvt('E_JOURNAL')=Q.FindField('E_JOURNAL').AsString) And
     (O.GetMvt('E_DATECOMPTABLE')=Q.FindField('E_DATECOMPTABLE').AsDateTime) And
     (O.GetMvt('E_EXERCICE')=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))) And
     (O.GetMvt('E_NUMEROPIECE')=Q.FindField('E_NUMEROPIECE').AsInteger) And
     (O.GetMvt('E_NUMLIGNE')=Q.FindField('E_NUMLIGNE').AsInteger) And
     (O.GetMvt('E_NUMECHE')=Q.FindField('E_NUMECHE').AsInteger) Then
    BEGIN
    (*
    Inc(NbP) ;
    D:=O.GetMvt('E_DEBIT') ; C:=O.GetMvt('E_CREDIT') ;
    If Arrondi(D,V_PGI.OkDecV)=0 Then C:=C-O.GetMvt('E_COUVERTURE')
                                 Else D:=D-O.GetMvt('E_COUVERTURE') ;
    XDebitP:=Arrondi(XDebitP+D,V_PGI.OkDecV) ;
    XCreditP:=Arrondi(XCreditP+C,V_PGI.OkDecV) ;
    D:=O.GetMvt('E_DEBITEURO') ; C:=O.GetMvt('E_CREDITEURO') ;
    If Arrondi(D,V_PGI.OkDecV)=0 Then C:=C-O.GetMvt('E_COUVERTUREEURO')
                                 Else D:=D-O.GetMvt('E_COUVERTUREEURO') ;
    XDebitE:=Arrondi(XDebitE+D,V_PGI.OkDecE) ;
    XCreditE:=Arrondi(XCreditE+C,V_PGI.OkDecE) ;
    *)
    Exit ;
    END ;
  END ;
Result:=FALSE ;
END ;  }

{procedure TFMulFactEff.FaitXX_WHERESEL ;
Var St,St1 : String ;
    O : TOBM ;
    i : Integer ;
BEGIN
St:='' ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  St1:='(E_JOURNAL="'+O.GetMvt('E_JOURNAL')+'"  AND E_EXERCICE="'+QuelExo(DateToStr(O.GetMvt('E_DATECOMPTABLE')))+'" AND '+
       ' E_DATECOMPTABLE="'+UsDateTime(O.GetMvt('E_DATECOMPTABLE'))+'" AND '+
       ' E_NUMEROPIECE='+IntToStr(O.GetMvt('E_NUMEROPIECE'))+' AND '+
       ' E_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(O.GetMvt('E_NUMECHE'))+') ' ;
  If St='' Then St:=St1 Else St:=St+' OR '+St1 ;
  END ;
If St='' Then Exit ;
St:='OR ('+St+')' ;
XX_WHERESEL.Text:=St ;
END ;}

{procedure TFMulFactEff.InitSelect ;
//Var i : Integer ;
BEGIN
Exit ; // Modif GP 07/01/2002
(*
If LOBM=Nil Then Exit ;
If LOBM.Count<=0 Then Exit ; i:=0 ;
Q.Manuel:=TRUE ;
XX_WHERESEL.Text:='' ; FaitXX_WHERESEL ;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
Q.First ;
InitMove(RecordsCount(Q),'');
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  If EstSelectionne(Q,LOBM) Then BEGIN Inc(i) ; FListe.FlipSelection ; END ;
  If i=LOBM.Count Then Break ;
  Q.Next ;
  END ;
FiniMove ;
Q.First ;
*)
END ;}

procedure TFMulFactEff.FormShow(Sender: TObject);
begin
DragFListeEnCours:=FALSE ; DragFListe2EnCours:=FALSE ;
OkSel:=FALSE ;
InitCriteres ;
  inherited;
If LOBM=NIL Then BEGIN Q.Manuel:=FALSE ; Q.UpdateCriteres ; END ;
CentreDBGrid(FListe) ;
Caption:=HM.Mess[1]+' '+LibCpt+' ('+Gene+'/'+Auxi+')' ;
//InitSelect ;
//If NewFlux And ((LOBM=NIL) OR (LOBM.Count=0)) Then bSelectAllClick(NIL) ;
RefreshFliste2 ;
CalculMontantSelect ;
TotSelInit:=TotSel ;
AfficheTotauxBasEcran ;
FTotSel.Visible:=FALSE ;
end;

procedure TFMulFactEff.InitCriteres ;
Var dd : TDateTime ;
BEGIN
E_GENERAL.Text:=Gene ; E_Auxiliaire.Text:=Auxi ; E_EXERCICE.ItemIndex:=0 ; E_DEVISE.Value:=V_PGI.DevisePivot ;
Case ModeSaisie Of
  OnEffet : BEGIN
            DD:=DebutDeMois(DateEcheDefaut) ; E_DATEECHEANCE.Text:=DateToStr(DD) ;
            DD:=FinDeMois(DateEcheDefaut) ; E_DATEECHEANCE_.Text:=DateToStr(DD) ;
            END ;
  OnBqe   : BEGIN
            E_DEVISE.Enabled:=TRUE ; TE_DEVISE.Enabled:=TRUE ;
            END ;  
  End ;
Case ModeSR Of
  srFou : BEGIN
          XX_WHERENTP.Text:='E_NATUREPIECE="FF" OR E_NATUREPIECE="AF" OR E_NATUREPIECE="OD"  OR E_NATUREPIECE="OF"' ;
          END ;
  END ;
END ;

(*================================ CRITERES ==================================*)
procedure TFMulFactEff.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
If E_EXERCICE.ItemIndex>0 Then ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFMulFactEff.FListeDblClick(Sender: TObject);
Var sMode : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
sMode:=Q.FindField('E_MODESAISIE').AsString ;
if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q,TypeAction) else
TrouveEtLanceSaisie(Q,taConsult,'N') ;
end;

Function TFMulFactEff.CtrlMontantSel(LOBM : TList) : Boolean ;
Var O : TOBM ;
    D,C,Couv : Double ;
    i : integer ;
    M : Double ;
BEGIN
Result:=FALSE ;
D:=0 ; C:=0 ; Couv:=0 ;
If LOBM=NIL Then Exit ; If LOBM.Count<=0 Then Exit ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  D:=Arrondi(D+O.GetMvt('E_DEBIT'),V_PGI.OkDecV) ; C:=Arrondi(C+O.GetMvt('E_CREDIT'),V_PGI.OkDecE) ;
  Couv:=Arrondi(Couv+O.GetMvt('E_COUVERTURE'),V_PGI.OkDecV) ;
  END ;
Result:=TRUE ;
M:=Arrondi(D-C-Couv,V_PGI.OkDecV) ;
If Modesr=srFou Then
  BEGIN
  If M>=0 Then Result:=FALSE ;
  END Else
  BEGIN
  If M<=0 Then Result:=FALSE ;
  END ;
END ;

procedure TFMulFactEff.FAitOBMDrag ;
Var i : Integer ;
    Q1 : TQuery ;
    O : TOBM ;
begin
  inherited;
If (Gene<>E_GENERAL.Text) Or (Auxi<>E_AUXILIAIRE.Text) Then
  BEGIN
  HM.execute(0,Caption,'') ; Exit ;
  END ;
If TypeSel=0 Then
  BEGIN
  VideListe(LOBM1) ;
  LOBM1.Clear ;
  for i:=0 to FListe.NbSelected-1 do
      BEGIN
      FListe.GotoLeBookmark(i) ;
      Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
                +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
                +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
                +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
                +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
                +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString
                +' AND E_GENERAL="'+Gene+'" '
                +' AND E_AUXILIAIRE="'+Auxi+'" '
                +' AND E_QUALIFPIECE="N" ',True) ;
      If Not Q1.Eof Then
        BEGIN
        O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(Q1) ; LOBM1.Add(O) ;
        END ;
      Ferme(Q1) ;
      END ;
  END Else
  BEGIN
  TOB1.ClearDetail ;
  for i:=0 to FListe.NbSelected-1 do
      BEGIN
      FListe.GotoLeBookmark(i) ;
      Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
                +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
                +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
                +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
                +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
                +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString
                +' AND E_GENERAL="'+Gene+'" '
                +' AND E_AUXILIAIRE="'+Auxi+'" '
                +' AND E_QUALIFPIECE="N" ',True) ;
      TOB1.LoadDetailDB('ECRITURE','','',Q1,TRUE,True) ;
      Ferme(Q1) ;
      END ;
  END ;
end;

procedure TFMulFactEff.BOuvrirClick(Sender: TObject);
//Var i : Integer ;
//    Q1 : TQuery ;
//    O : TOBM ;
begin
  inherited;
If Sender<>NIL Then
  BEGIN
  If Fliste.NbSelected>0 Then TraiteMajFListe2(TRUE) ;
  END ;
If LOBM=NIL Then Exit ;
If (Gene<>E_GENERAL.Text) Or (Auxi<>E_AUXILIAIRE.Text) Then
  BEGIN
  HM.execute(0,Caption,'') ; Exit ;
  END ;
(*
VideListe(LOBM) ;
for i:=0 to FListe.NbSelected-1 do
    BEGIN
    FListe.GotoLeBookmark(i) ;
    Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
              +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
              +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
              +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
              +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
              +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString
              +' AND E_GENERAL="'+Gene+'" '
              +' AND E_AUXILIAIRE="'+Auxi+'" '
              +' AND E_QUALIFPIECE="N" ',True) ;
    If Not Q1.Eof Then
      BEGIN
      O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(Q1) ; LOBM.Add(O) ;
      END ;
    Ferme(Q1) ;
    END ;
*)

If Not CtrlMontantSel(LOBM) Then BEGIN HM.Execute(2,Caption,'') ; Exit ; END ;
ModalResult:=mrOk ;
OkSel:=TRUE ; //Close ;
end;

procedure TFMulFactEff.AfficheTotauxBasEcran ;
BEGIN
{ Totaux Pointés sur la référence (bas d'écran) }
FCreditPA.Value:=TotSel.XCreditP ; FDebitPA.Value:=TotSel.XDebitP ;
AfficheLeSolde(FSoldeA,FDebitPA.Value,FCreditPA.Value) ;
FNbP.Value:=TotSel.NbP ;
WFListe2.Caption:=HM.Mess[5]+'   '+HM.Mess[6]+' '+FloatToStrF(FDebitPA.Value,ffNumber,11,2)+
                             '   '+HM.Mess[7]+' '+FloatToStrF(FCreditPA.Value,ffNumber,11,2)+
                             '   '+HM.Mess[8]+' '+FloatToStrF(FSoldeA.Value,ffNumber,10,2)+
                             ' / ('+IntToStr(TotSel.NbP)+' échéance(s)'  ;
END ;

procedure TFMulFactEff.CalculMontantSelect ;
var i : Integer ;
    O : TOBM ;
    D,C : Double ;
BEGIN
Fillchar(TotSel,SizeOf(TotSel),#0) ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  TotSel.NbP:=TotSel.NbP+1 ;
  D:=O.GetMvt('E_DEBIT') ; C:=O.GetMvt('E_CREDIT') ;
  If Arrondi(D,V_PGI.OkDecV)=0 Then C:=C-O.GetMvt('E_COUVERTURE')
                               Else D:=D-O.GetMvt('E_COUVERTURE') ;
  TotSel.XDebitP:=Arrondi(TotSel.XDebitP+(D),V_PGI.OkDecV) ;
  TotSel.XCreditP:=Arrondi(TotSel.XCreditP+(C),V_PGI.OkDecV) ;
  END ;
END ;

procedure TFMulFactEff.FListeFlipSelection(Sender: TObject);
Var Coeff : Integer ;
    D,C : Double ;
begin
  inherited;
AfficheTotauxBasEcran ;
If FListe.IsCurrentSelected Then Coeff:=1 Else Coeff:=-1 ;
TotSel.NbP:=TotSel.NbP+Coeff ;
D:=Q.FindField('E_DEBIT').AsFloat ; C:=Q.FindField('E_CREDIT').AsFloat ;
If Arrondi(D,V_PGI.OkDecV)=0 Then C:=C-Q.FindField('E_COUVERTURE').AsFloat
                             Else D:=D-Q.FindField('E_COUVERTURE').AsFloat ;
TotSel.XDebitP:=Arrondi(TotSel.XDebitP+(D*Coeff),V_PGI.OkDecV) ;
TotSel.XCreditP:=Arrondi(TotSel.XCreditP+(C*Coeff),V_PGI.OkDecV) ;
AfficheTotauxBasEcran ;
end;

procedure TFMulFactEff.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : boolean ;
begin
Vide:=(Shift=[]) ;
if Vide then
   BEGIN
   Case Key of
        VK_F10 : BEGIN
                 Key:=0 ;
                 If Fliste.NbSelected>0 Then TraiteMajFListe2(TRUE) ;
                 BOuvrirClick(Nil) ; Exit ;
                 END ;
        107  :  BEGIN Key:=0 ; TraiteMajFListe2(TRUE) ; Exit ; END ; // "+"
     END ;
   END ;
Inherited ;
end;

procedure TFMulFactEff.bSelectAllClick(Sender: TObject);
begin
Q.First ;
While Not Q.Eof Do
  BEGIN
  FListe.FlipSelection ;
  Q.Next ;
  END
end;

procedure TFMulFactEff.FListeDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  inherited;
//Accept:=FALSE ;
If DragFListe2EnCours Then Accept:=True ;
end;

Function TFMulFactEff.ExistePasDansLOBM(O1 : TOBM ; TOBL1 : TOB) : Boolean ;
Var O : TOBM ;
    i : Integer ;
BEGIN
Result:=TRUE ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  If O1<>NIL Then
    BEGIN
    If (O.GetMvt('E_JOURNAL')=O1.GetMvt('E_JOURNAL')) And
       (O.GetMvt('E_EXERCICE')=O1.GetMvt('E_EXERCICE')) And
       (O.GetMvt('E_DATECOMPTABLE')=O1.GetMvt('E_DATECOMPTABLE')) And
       (O.GetMvt('E_NUMEROPIECE')=O1.GetMvt('E_NUMEROPIECE')) And
       (O.GetMvt('E_NUMLIGNE')=O1.GetMvt('E_NUMLIGNE')) And
       (O.GetMvt('E_NUMECHE')=O1.GetMvt('E_NUMECHE')) And
       (O.GetMvt('E_QUALIFPIECE')=O1.GetMvt('E_QUALIFPIECE')) Then
         BEGIN
         Result:=FALSE ; Exit ;
         END ;
    END Else
    BEGIN
    If (O.GetMvt('E_JOURNAL')=TOBL1.GetValue('E_JOURNAL')) And
       (O.GetMvt('E_EXERCICE')=TOBL1.GetValue('E_EXERCICE')) And
       (O.GetMvt('E_DATECOMPTABLE')=TOBL1.GetValue('E_DATECOMPTABLE')) And
       (O.GetMvt('E_NUMEROPIECE')=TOBL1.GetValue('E_NUMEROPIECE')) And
       (O.GetMvt('E_NUMLIGNE')=TOBL1.GetValue('E_NUMLIGNE')) And
       (O.GetMvt('E_NUMECHE')=TOBL1.GetValue('E_NUMECHE')) And
       (O.GetMvt('E_QUALIFPIECE')=TOBL1.GetValue('E_QUALIFPIECE')) Then
         BEGIN
         Result:=FALSE ; Exit ;
         END ;
    END ;
  END ;
END ;

// FQ 11665
function TFMulFactEff.ExisteDansTOB(O1: TOBM; TOBL1: TOB): Boolean;
Var
  O : TOB;
  i : Integer;
begin
  Result := False;
  For i := 0 To TOB2.Detail.Count-1 Do BEGIN
    O := TOB2.Detail[i];
    If O1<>NIL Then BEGIN
      If (O.GetValue('E_JOURNAL')      = O1.GetMvt('E_JOURNAL')) And
         (O.GetValue('E_EXERCICE')     = O1.GetMvt('E_EXERCICE')) And
         (O.GetValue('E_DATECOMPTABLE')= O1.GetMvt('E_DATECOMPTABLE')) And
         (O.GetValue('E_NUMEROPIECE')  = O1.GetMvt('E_NUMEROPIECE')) And
         (O.GetValue('E_NUMLIGNE')     = O1.GetMvt('E_NUMLIGNE')) And
         (O.GetValue('E_NUMECHE')      = O1.GetMvt('E_NUMECHE')) And
         (O.GetValue('E_QUALIFPIECE')  = O1.GetMvt('E_QUALIFPIECE')) Then BEGIN
        Result := True;
        Exit ;
      END ;
      END
    Else BEGIN
      If (O.GetValue('E_JOURNAL')       = TOBL1.GetValue('E_JOURNAL')) And
         (O.GetValue('E_EXERCICE')      = TOBL1.GetValue('E_EXERCICE')) And
         (O.GetValue('E_DATECOMPTABLE') = TOBL1.GetValue('E_DATECOMPTABLE')) And
         (O.GetValue('E_NUMEROPIECE')   = TOBL1.GetValue('E_NUMEROPIECE')) And
         (O.GetValue('E_NUMLIGNE')      = TOBL1.GetValue('E_NUMLIGNE')) And
         (O.GetValue('E_NUMECHE')       = TOBL1.GetValue('E_NUMECHE')) And
         (O.GetValue('E_QUALIFPIECE')   = TOBL1.GetValue('E_QUALIFPIECE')) Then BEGIN
        Result := True;
        Exit ;
      END ;
    END ;
  END ;
end;

Procedure TOBToOBM(TOBL1 : TOB ; Var O : TOBM) ;
Var i : Integer ;
BEGIN
O:=TOBM.Create(EcrGen,'',False) ;
for i:=1 to TOBL1.NbChamps do O.PutMvt(TOBL1.GetNomChamp(i), TOBL1.GetValeur(i)) ;
END ;

// FQ 11665
procedure TFMulFactEff.RefreshLOBM ;
var i : Integer ;
    O1,O : TOBM ;
    TObl1 : TOB ;
    bMessage : Boolean;
BEGIN
  bMessage := True;
If TypeSel=0 Then
  BEGIN
  For i:=0 To LOBM1.Count-1 Do
    BEGIN
    O1:=TOBM(LOBM1[i]) ; O:=NIL ;
    if ExisteDansTOB(O1,nil) then begin if bMessage then HM.Execute(10,Caption,''); bMessage := False; Continue; end;    
    If ExistePasDansLOBM(O1,NIL) Then BEGIN EgaliseOBM(O1,O) ; LOBM.Add(O) ; END ;
    END ;
  END Else
  BEGIN
  For i:=0 To TOB1.Detail.Count-1 Do
    BEGIN
    TOBL1:=TOB1.Detail[i] ; O:=NIL ;
    if ExisteDansTOB(nil,TOBL1) then begin if bMessage then HM.Execute(10,Caption,''); bMessage := False; Continue; end;    
    If ExistePasDansLOBM(Nil,TOBL1) Then BEGIN TOBToOBM(TOBL1,O) ; LOBM.Add(O) ; END ;
    END ;
  END ;
END ;

procedure TFMulFactEff.RefreshFliste2 ;
Var i,j : Integer ;
    TobMvt : TOB ;
    O : TOBM ;
    St : String ;
begin
//AfficheFliste2(True) ; NbFListe2Selected:=0 ;
//MeMoLig:=FListe.Row ;
TOBFliste2.ClearDetail ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  TOBMVT:=TOB.Create('ECRITURE',TOBFListe2,-1) ;
  for j:=1 to TOBMVT.NbChamps do
    BEGIN
    St:=TOBMVT.GetNomChamp(j) ;
    TOBMVT.PutValue(St, O.GetMvt(St)) ;
    END ;
  END ;
FListe2.SortEnabled:=FALSE ; FListe2.SortedCol:=-1 ;
TOBFListe2.PutGridDetail(Fliste2,TRUE,TRUE,'E_JOURNAL;E_GENERAL;E_AUXILIAIRE;E_DATECOMPTABLE;E_DATEECHEANCE;E_NUMEROPIECE;E_DEBIT;E_CREDIT;IND;',TRUE) ;
Fliste2.ColWidths[FListe2.ColCount-1]:=0 ;
For i:=1 To Fliste2.RowCount-1 Do
  BEGIN
  FListe2.Cells[Fliste2.ColCount-1,i]:=IntToStr(i-1) ;
  END ;
FListe2.SortEnabled:=TRUE ; FListe2.SortedCol:=0 ;
Fliste2.Refresh ;
If Fliste2.RowCount<=10 Then FListe2.Height:=23+((Fliste2.DefaultRowHeight+1)*FListe2.RowCount)
                       Else FListe2.Height:=23+((Fliste2.DefaultRowHeight+1)*10) ;
end;

procedure TFMulFactEff.TraiteMajFliste2(Plus : Boolean) ;
BEGIN
FaitOBMDrag ;
RefreshLOBM ;
If Plus Then OkDropFliste:=FALSE ;
RefreshFliste2 ;
FListe.ClearSelected ;
CalculMontantSelect ; AfficheTotauxBasEcran ;
If Plus Then
  BEGIN
  DragFListeEnCours:=FALSE ;
  FListe2.ClearSelected ;
  END ;
END ;

procedure TFMulFactEff.FListeEndDrag(Sender, Target: TObject; X, Y: Integer);
Var TC : TControl ;
begin
Inherited ;
If DragFListeEnCours Then
  BEGIN
  TC:=TControl(Target) ; If (TC=NIL) Or (TC.Name<>FListe2.Name) Then OkDropFliste:=FALSE ;
  If OkDropFliste Then TraiteMajFListe2(FALSE) ;
  DragFListeEnCours:=FALSE ;
  END ;
FListe2.ClearSelected ;
end;

procedure TFMulFactEff.FListeMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
(*
FListe2.MouseToCell(X,Y,ACol,ARow) ;
If ARow=0 Then Exit ;
*)
If Not DragFListe2EnCours Then
  BEGIN
  if ((ssLeft in Shift)) then
    BEGIN
    DragFListeEnCours:=TRUE ; Fliste.BeginDrag(True) ; OkDropFliste:=FALSE ;
    (*NbFListe2Selected:=FListe2.NbSelected ;*)
    END ;
  END ;
end;

procedure TFMulFactEff.FListe2DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  inherited;
If DragFListeEnCours Then
  BEGIN
  OkDropFliste:=FALSE ;
  If HM.Execute(3,Caption,'')=mrYes Then OkDropFliste:=TRUE  Else DragFListeEnCours:=FALSE ;
  END ;
end;

procedure TFMulFactEff.FListe2DragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
Inherited ;
If DragFListeEnCours Then Accept:=True ;
//If NbFListe2Selected>0 Then FListe.DragCursor:=crMultiDrag Else FListe.DragCursor:=crDrag ;
(*
FListe.MouseToCell(X,Y,ACol,ARow) ;
if ARow<FListe.FixedRows then BEGIN Accept:=False ; Exit ; END ;
If ARow>Fliste.RowCount-1 then BEGIN Accept:=False ; Exit ; END ;
Fliste.Row:=ARow ;
*)
end;

procedure TFMulFactEff.FormClose(Sender: TObject; var Action: TCloseAction);
begin
VideListe(LOBM1) ; LOBM1.Clear ; LOBM1.Free ;
TOB1.ClearDetail ; TOB1.Free ;
TOBFliste2.ClearDetail ; TOBFListe2.Free ;
RegSaveToolbarPos(Self,'MulFactEff') ;
  inherited;

end;

procedure TFMulFactEff.FListe2MouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
(*
FListe2.MouseToCell(X,Y,ACol,ARow) ;
If ARow=0 Then Exit ;
*)
If Not DragFListeEnCours Then
  BEGIN
  if ((ssLeft in Shift)) then
    BEGIN
    DragFListe2EnCours:=TRUE ; Fliste2.BeginDrag(True) ; OkDropFliste2:=FALSE ;
    (*NbFListe2Selected:=FListe2.NbSelected ;*)
    END ;
  END ;
end;

Procedure TFMulFactEff.DeleteFromLOBM(TOBL : TOB) ;
Var O,OToKill : TOBM ;
    i,Ind : Integer ;

BEGIN
Ind:=-1 ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  If (O.GetMvt('E_JOURNAL')=TOBL.GetValue('E_JOURNAL')) And
     (O.GetMvt('E_EXERCICE')=TOBL.GetValue('E_EXERCICE')) And
     (O.GetMvt('E_DATECOMPTABLE')=TOBL.GetValue('E_DATECOMPTABLE')) And
     (O.GetMvt('E_NUMEROPIECE')=TOBL.GetValue('E_NUMEROPIECE')) And
     (O.GetMvt('E_NUMLIGNE')=TOBL.GetValue('E_NUMLIGNE')) And
     (O.GetMvt('E_NUMECHE')=TOBL.GetValue('E_NUMECHE')) And
     (O.GetMvt('E_QUALIFPIECE')=TOBL.GetValue('E_QUALIFPIECE')) Then BEGIN Ind:=i ; Break ; END ;
  END ;
If Ind<0 Then Exit ;
OToKill:=TOBM(LOBM[Ind]) ;
OToKill.Free ;
LOBM.Delete(Ind) ;
END ;

Procedure TFMulFactEff.DeleteMvtFliste2 ;
Var i,j : Integer ;
    TOBL : TOB ;
BEGIN
Fliste2.SortEnabled:=FALSE ;
If FListe2.NbSelected>0 Then
  BEGIN
  For i:=0 To FListe2.RowCount-1 Do If FListe2.IsSelected(i) Then
    BEGIN
    j:=StrToInt(Fliste2.Cells[Fliste2.ColCount-1,i]) ;
    TOBL:=TobFListe2.Detail[j] ;
    If TOBL<>NIL Then DeleteFromLOBM(TOBL) ;
    END ;
  END Else
  BEGIN
  j:=StrToInt(Fliste2.Cells[Fliste2.ColCount-1,FListe2.Row]) ;
  TOBL:=TobFListe2.Detail[j] ;
  If TOBL<>NIL Then DeleteFromLOBM(TOBL) ;
  END ;
Fliste2.SortEnabled:=TRUE ;
END ;



procedure TFMulFactEff.FListe2EndDrag(Sender, Target: TObject; X, Y: Integer);
Var TC : TControl ;
begin
Inherited ;
If DragFListe2EnCours Then
  BEGIN
  TC:=TControl(Target) ; If (TC=NIL) Or (TC.Name<>FListe.Name) Then OkDropFliste2:=FALSE ;
  If OkDropFliste2 Then
    BEGIN
    (*
    FaitOBMDrag ;
    RefreshLOBM ;
    *)
    DeleteMvtFliste2 ;
    OkDropFliste2:=FALSE ;
    RefreshFliste2 ;
    CalculMontantSelect ; AfficheTotauxBasEcran ;
    END  ;
  DragFListe2EnCours:=FALSE ;
  END ;
FListe2.ClearSelected ;
end;

procedure TFMulFactEff.FListeDragDrop(Sender, Source: TObject; X,Y: Integer);
begin
  inherited;
If DragFListe2EnCours Then
  BEGIN
  OkDropFliste2:=FALSE ;
  If HM.Execute(4,Caption,'')=mrYes Then OkDropFliste2:=TRUE  Else DragFListe2EnCours:=FALSE ;
  END ;
end;

procedure TFMulFactEff.bVoirTotSelClick(Sender: TObject);
begin
  inherited;
FTotSel.Visible:=bVoirtotSel.Down ;
end;

procedure TFMulFactEff.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  inherited;
If (ModalResult=mrCancel) And (LOBM<>NIL) And (LOBM.Count>0) And
   (Arrondi(TotSelInit.XDEBITP+TotSelInit.XCREDITP-(TotSel.XDEBITP+TotSelInit.XCREDITP),V_PGI.OkDecV)<>0) Then
  Begin
  If HM.Execute(9,Caption,'')<>MrYes Then CanClose:=FALSE ;
  End ;
end;

end.

