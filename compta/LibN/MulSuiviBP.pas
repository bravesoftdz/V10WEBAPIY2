unit MulSuiviBP ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry,
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, Saisie, SaisUtil,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil,
  EncUtil, LettUtil, UTOB, ed_Tools
  ,SaisBor, HRichOLE, ADODB
  ;

Function SuiviBP(Client : Boolean ; TOBMvt : TOB) : Double ;

type
  TFMulSuiviBP = class(TFMul)
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
    HLabel17: THLabel;
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
    PzLibre: TTabSheet;
    Bevel5: TBevel;
    TT_TABLE0: THLabel;
    T_TABLE0: THCpteEdit;
    TT_TABLE5: THLabel;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    TT_TABLE6: THLabel;
    T_TABLE1: THCpteEdit;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    T_TABLE2: THCpteEdit;
    TT_TABLE7: THLabel;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    TT_TABLE8: THLabel;
    T_TABLE3: THCpteEdit;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    T_TABLE4: THCpteEdit;
    TT_TABLE9: THLabel;
    T_TABLE9: THCpteEdit;
    XX_WHEREENC: TEdit;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FListeFlipSelection(Sender: TObject);
    procedure BChercheClick(Sender: TObject); override;
    procedure bSelectAllClick(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject);
  private    { Déclarations privées }
    TOBMvt : TOB ;
    ModeOppose : Boolean ;
    Client : Boolean ;
    XDebitP,XCreditP : Double ;
    NbP : Integer ;
    TotSel : Double ;
    Manuel,PremFois : Boolean ;
    procedure InitEnc ;
    procedure InitDec ;
    procedure InitCriteres ;
    procedure InitSelect ;
    Function  EstSelectionne(Q : TQuery ; TobMvt : Tob) : Boolean ;
    procedure AfficheTotauxBasEcran ;
    procedure PrechargeOrigines(UneSeule : Boolean ; OkSel : Boolean) ;
    Function  MarqueOrigine(UneSeule,OkSel : Boolean) : Boolean ;
    procedure MajTobMvt ;
    procedure ToutMarquer ;
    procedure VireInutiles ;
    procedure InitTotaux ;
    procedure TrouveTot(Var D,C : Double ; Var Nb : Integer) ;
 public
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus ;

Function SuiviBP(Client : Boolean ; TOBMvt : TOB) : Double ;
var X : TFMulSuiviBP ;
    PP : THPanel ;
//    i : Integer ;
begin
Result:=-1 ;
PP:=FindInsidePanel ;
X:=TFMulSuiviBP.Create(Application) ;
X.Q.Manuel:=True ;
X.Client:=Client ;
X.TobMvt:=TobMvt ;
X.TotSel:=-1 ;
If Client Then
  BEGIN
  X.FNomFiltre:='CPENCBQPREVI' ; X.Q.Liste:='CPENCBQPREVI' ;
  END Else
  BEGIN
  X.FNomFiltre:='CPDECBQPREVI' ; X.Q.Liste:='CPDECBQPREVI' ;
  END ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     Result:=X.TotSel ;
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
procedure TFMulSuiviBP.FormCreate(Sender: TObject);
begin
  inherited;
MemoStyle:=msBook ;
end;


Function TFMulSuiviBP.EstSelectionne(Q : TQuery ; TobMvt : Tob) : Boolean ;
Var i : Integer ;
    TOBL : TOB ;
BEGIN
Result:=TRUE ;
For i:=0 To TobMvt.Detail.Count-1 Do
  BEGIN
  TOBL:=TOBMvt.Detail[i] ;
  If (TOBL.GetValue('E_JOURNAL')=Q.FindField('E_JOURNAL').AsString) And
     (TOBL.getvalue('E_DATECOMPTABLE')=Q.FindField('E_DATECOMPTABLE').AsDateTime) And
     (TOBL.getvalue('E_EXERCICE')=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))) And
     (TOBL.getvalue('E_NUMEROPIECE')=Q.FindField('E_NUMEROPIECE').AsInteger) And
     (TOBL.getvalue('E_NUMLIGNE')=Q.FindField('E_NUMLIGNE').AsInteger) And
     (TOBL.getvalue('E_NUMECHE')=Q.FindField('E_NUMECHE').AsInteger) Then
    BEGIN
    Exit ;
    END ;
  END ;
Result:=FALSE ;
END ;

procedure TFMulSuiviBP.InitSelect ;
Var i : Integer ;
BEGIN
If TobMvt=Nil Then Exit ;
If TobMvt.Detail=Nil Then Exit ;
If TobMvt.Detail.Count<=0 Then Exit ; i:=0 ;
Manuel:=FALSE ;
InitMove(RecordsCount(Q),'') ;
Q.First ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  If EstSelectionne(Q,TobMvt) Then BEGIN Inc(i) ; FListe.FlipSelection ; END ;
  If i>=TobMvt.Detail.Count Then Break ;
  Q.Next ;
  END ;
FiniMove ;
Q.First ;
AfficheTotauxBasEcran ;
Manuel:=TRUE ;
END ;

procedure TFMulSuiviBP.InitTotaux ;
BEGIN
XDebitP:=0 ; XCreditP:=0 ; NbP:=0 ;
END ;

procedure TFMulSuiviBP.FormShow(Sender: TObject);
begin
Manuel:=TRUE ; premFois:=TRUE ;
ModeOppose:=FALSE ; InitTotaux ;
InitCriteres ;
  inherited;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
If Client Then Caption:=HM.Mess[1] Else Caption:=HM.Mess[2] ;
InitSelect ;
end;

procedure TFMulSuiviBP.InitEnc ;
BEGIN
XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND E_NATUREPIECE="AC")' ;
E_GENERAL.DataType:='TZGENCAIS' ; E_AUXILIAIRE.DataType:='TZTTOUTDEBIT' ;
XX_WHERENTP.Text:='E_NATUREPIECE="FC" OR E_NATUREPIECE="AC" OR E_NATUREPIECE="OD"  OR E_NATUREPIECE="OC"' ;
END ;

procedure TFMulSuiviBP.InitDec ;
BEGIN
XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND E_NATUREPIECE="AF")' ;
E_GENERAL.DataType:='TZGDECAIS' ; E_AUXILIAIRE.DataType:='TZTTOUTCREDIT' ;
XX_WHERENTP.Text:='E_NATUREPIECE="FF" OR E_NATUREPIECE="AF" OR E_NATUREPIECE="OD"  OR E_NATUREPIECE="OF"' ;
END ;


procedure TFMulSuiviBP.InitCriteres ;
//Var i : integer ;
BEGIN
LibellesTableLibre(PzLibre,'TT_TABLE','T_TABLE','T') ;
E_EXERCICE.ItemIndex:=0 ; E_DEVISE.Value:=V_PGI.DevisePivot ;
If CLient Then InitEnc Else InitDec ;
END ;

(*================================ CRITERES ==================================*)
procedure TFMulSuiviBP.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
If E_EXERCICE.ItemIndex>0 Then ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFMulSuiviBP.FListeDblClick(Sender: TObject);
Var sMode : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
sMode:=Q.FindField('E_MODESAISIE').AsString ;
if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q,TypeAction) else
TrouveEtLanceSaisie(Q,taConsult,'N') ;
end;

Function TFMulSuiviBP.MarqueOrigine(UneSeule,OkSel : Boolean) : Boolean ;
Var TOBL : TOB ;
BEGIN
Result:=FALSE ;
TOBL:=TOBMvt.FindFirst(['E_JOURNAL','E_EXERCICE','E_DATECOMPTABLE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                        [Q.FindField('E_JOURNAL').AsString,Q.FindField('E_EXERCICE').AsString,
                         Q.FindField('E_DATECOMPTABLE').AsDateTime,Q.FindField('E_NUMEROPIECE').AsInteger,
                         Q.FindField('E_NUMLIGNE').AsInteger,Q.FindField('E_NUMECHE').AsInteger],False) ;
if TOBL<>Nil then
  BEGIN
  If UneSeule Then
    BEGIN
    If OkSel Then TOBL.PutValue('MARQUE','X') Else
      BEGIN
      TOBL.PutValue('MARQUE','-') ;
//      TOBL.Free ; TOBL:=Nil ;
      END ;
    END Else TOBL.PutValue('MARQUE','X') ;
    Result:=TRUE ;
  END ;
END ;

procedure TFMulSuiviBP.PrechargeOrigines(UneSeule : Boolean ; OkSel : Boolean) ;
Var QQ : TQuery ;
    sWhere : String ;
BEGIN
If UneSeule Then
  BEGIN
  If PremFois Then TOBMvt.ClearDetail ;
  If MarqueOrigine(UneSeule,OkSel) Then Exit ;
  SWhere:='WHERE E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'" AND E_EXERCICE="'+Q.FindField('E_EXERCICE').AsString+'" AND '+
          'E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'" AND E_NUMEROPIECE='+IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger)+' AND '+
          'E_NUMLIGNE='+IntToStr(Q.FindField('E_NUMLIGNE').AsInteger)+' AND E_NUMECHE='+IntToStr(Q.FindField('E_NUMECHE').AsInteger)+' AND '+
          'E_QUALIFPIECE="N" ' ;
  END Else
  BEGIN
  TOBMvt.ClearDetail ;
  SWhere:=RecupWhereCritere(Pages) ;
  END ;
QQ:=OpenSQL('SELECT * FROM ECRITURE '+SWhere,True) ;
TOBMvt.LoadDetailDB('ECRITURE','','',QQ,UneSeule,Not UneSeule) ;
Ferme(QQ) ;
if TOBMvt.Detail.Count>0 then
  BEGIN
  TOBMvt.Detail[0].AddChampSup('MARQUE',True) ;
  END ;
If UneSeule Then MarqueOrigine(UneSeule,OkSel) ;
END ;

procedure TFMulSuiviBP.VireInutiles ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=TOBMvt.Detail.Count-1 downto 0 do
    BEGIN
    TOBL:=TobMvt.Detail[i] ;
    if TOBL.GetValue('MARQUE')<>'X' then BEGIN TOBL.Free ; {TOBL:=Nil ; }END ;
    END ;
END ;

procedure TFMulSuiviBP.ToutMarquer ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBMvt.Detail.Count-1 do
    BEGIN
    TOBL:=TOBMvt.Detail[i] ;
    TOBL.PutValue('MARQUE','X') ;
    END ;
END ;

procedure TFMulSuiviBP.TrouveTot(Var D,C : Double ; Var Nb : Integer) ;
Var T : Tcontrol ;
    HN : THNumEdit ;
    i : Integer ;
    St : String ;
BEGIN
D:=0 ; C:=0 ;
For i:=0 To PCumul.ControlCount-1 Do
  BEGIN
  T:=TControl(PCumul.Controls[i]) ;
  If (T is THNumEdit)=FALSE Then Continue ;
  HN:=THNumEdit(T) ;
  If HN.Name='__QRYE_DEBIT' Then D:=HN.Value ;
  If HN.Hint='__QRYE_CREDIT' Then C:=HN.Value ;
  END ;
St:=PCumul.Caption ; Nb:=ValeurI(St) ;
END ;

procedure TFMulSuiviBP.MajTobMvt ;
{Var i : integer ;
    TOBL : TOB ;
    StBQE,StCAT,StMP,St : String ;}
BEGIN
TotSel:=0 ;
(*
PrechargeOrigines(FALSE,FALSE) ;
if Not FListe.AllSelected then
   BEGIN
   for i:=0 to FListe.NbSelected-1 do
       BEGIN
       FListe.GotoLeBookmark(i) ;
       MarqueOrigine ;
       END ;
*)
if FListe.AllSelected then
   BEGIN
   PrechargeOrigines(FALSE,FALSE) ;
   ToutMarquer ;
   END ;
VireInutiles ;
If FListe.AllSelected Then TrouveTot(XDebitP,XCreditP,NbP) ;
If CLient Then TotSel:=Arrondi(XDebitP-XCreditP,V_PGI.OkDecV)
          Else TotSel:=Arrondi(XCreditP-XDebitP,V_PGI.OkDecV) ;
END ;


procedure TFMulSuiviBP.BOuvrirClick(Sender: TObject);
{Var i : Integer ;
    Q1 : TQuery ;
    O : TOBM ;}
begin
  inherited;
If TobMvt=NIL Then Exit ; MajTobMvt ; Close ;
end;

procedure TFMulSuiviBP.AfficheTotauxBasEcran ;
BEGIN
{ Totaux Pointés sur la référence (bas d'écran) }
FCreditPA.Value:=XCreditP ; FDebitPA.Value:=XDebitP ;
AfficheLeSolde(FSoldeA,FDebitPA.Value,FCreditPA.Value) ;
FNbP.Value:=NbP ;
END ;


procedure TFMulSuiviBP.FListeFlipSelection(Sender: TObject);
Var Coeff : Integer ;
begin
  inherited;
AfficheTotauxBasEcran ;
If FListe.IsCurrentSelected Then Coeff:=1 Else Coeff:=-1 ;
NbP:=NbP+Coeff ;
XDebitP:=Arrondi(XDebitP+(Q.FindField('E_DEBIT').AsFloat*Coeff),V_PGI.OkDecV) ;
XCreditP:=Arrondi(XCreditP+(Q.FindField('E_CREDIT').AsFloat*Coeff),V_PGI.OkDecV) ;
If Manuel Then
  BEGIN
  PrechargeOrigines(TRUE,FListe.IsCurrentSelected) ;
  END ;
AfficheTotauxBasEcran ;
PremFois:=FALSE ;
end;

procedure TFMulSuiviBP.BChercheClick(Sender: TObject);
begin
  InitTotaux ;
  AfficheTotauxBasEcran ;
  inherited;
  FListeRowEnter(nil); // 10930
end;

procedure TFMulSuiviBP.bSelectAllClick(Sender: TObject);
begin
  inherited;
InitTotaux ;
TobMvt.ClearDetail ;
If bSelectAll.Down Then TrouveTot(XDebitP,XCreditP,NbP) ;
AfficheTotauxBasEcran ;
end;

procedure TFMulSuiviBP.FListeRowEnter(Sender: TObject);
begin
  // 10930
  inherited;
  VH^.MPModifFaite:=FALSE ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL Then VH^.MPPop.MPExoPop:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
  If Q.FindField('E_GENERAL')<>NIL Then VH^.MPPop.MPGenPop:=Q.FindField('E_GENERAL').AsString ;
  If Q.FindField('E_AUXILIAIRE')<>NIL Then VH^.MPPop.MPAuxPop:=Q.FindField('E_AUXILIAIRE').AsString ;
  If Q.FindField('E_JOURNAL')<>NIL Then VH^.MPPop.MPJalPop:=Q.FindField('E_JOURNAL').AsString ;
  If Q.FindField('E_NUMEROPIECE')<>NIL Then VH^.MPPop.MPNumPop:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  If Q.FindField('E_NUMLIGNE')<>NIL Then VH^.MPPop.MPNumLPop:=Q.FindField('E_NUMLIGNE').AsInteger ;
  If Q.FindField('E_NUMECHE')<>NIL Then VH^.MPPop.MPNumEPop:=Q.FindField('E_NUMECHE').AsInteger ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL Then VH^.MPPop.MPDatePop:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
end;

end.

