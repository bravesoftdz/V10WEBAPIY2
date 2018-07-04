unit RepDevEur;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,HPanel, UiUtil,
  HTB97, hmsgbox, HSysMenu, Mask, StdCtrls, Hctrls, Hcompte, ExtCtrls, Ent1, HEnt1,CritEdt,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  SaisUtil,SaisComm, HStatus, RapSuppr,RappType,ed_tools, db, UtilPGI, UTOB, Math ;

procedure ReparationMontantEuro ;
Procedure ChargeDevise(ListeDev : TStringList ; Code : String) ;
Procedure TraiteUnePiece(StPiece : String ; ListeDev : tStringList ; ListeErr : TList ; OkRepare,OkChancell,OkMess : Boolean ; TauxModifie : Double) ;
procedure ChangeLeTauxDevise(TOBOrig : TOB ; Comment : tActionFiche) ;
{$IFDEF EAGLCLIENT}
procedure UpdateEcriture(Pref : Char; Q : TOB);{JP 30/05/06 : FQ 18242}
{$ENDIF EAGLCLIENT}

Type tDevInOut = Record
                 Code : String ;
                 Decim : Integer ;
                 MonnaieIn : Boolean ;
                 TauxIn : Double ;
                 Quotite : Double ;
                 End ;
Type tClePiece = Record
                 Jal,Exo,Qualif,Dev : String ;
                 NumP : Integer ;
                 DateP : TDateTime ;
                 End ;

type
  TFRepareMontant = class(TForm)
    Dock: TDock97;
    HPB: TToolWindow97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    Panel1: TPanel;
    TFJal: THLabel;
    TFaJ: TLabel;
    TFExercice: THLabel;
    FExercice: THValComboBox;
    TFDateCpta1: THLabel;
    FDateCpta1: TMaskEdit;
    TFDateCpta2: TLabel;
    FDateCpta2: TMaskEdit;
    TFEtab: THLabel;
    FEtab: THValComboBox;
    TFTypeEcriture: THLabel;
    FTypeEcriture: THValComboBox;
    TFNumPiece1: THLabel;
    FNumPiece1: TMaskEdit;
    TFNumPiece2: TLabel;
    FNumPiece2: TMaskEdit;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    TE_DEVISE: THLabel;
    FDevise: THValComboBox;
    FOkRepare: TCheckBox;
    FChancell: TCheckBox;
    BChancelOut: TToolbarButton97;
    FJal1: THCpteEdit;
    FJal2: THCpteEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FExerciceChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BChancelOutClick(Sender: TObject);
  private
    { Déclarations privées }
    CritEdt : TCritEdt ;
    ListeDev : TStringList ;
    ListePiece : TStringList ;
    procedure RecupCrit ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  CPChancell_TOF,
  SaisTaux1 ;

procedure ReparationMontantEuro ;
var CDiv : TFRepareMontant ;
    PP : THPanel ;
BEGIN
if Not _BlocageMonoPoste(True) then Exit ;
CDiv:=TFRepareMontant.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     CDiv.ShowModal ;
    finally
     CDiv.Free ;
     _DeblocageMonoPoste(True) ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(CDiv,PP) ;
   CDiv.Show ;
   END ;
END ;

procedure TFRepareMontant.RecupCrit ;
BEGIN
Fillchar(CritEdt,SizeOf(CritEdt),#0) ;
With CritEdt Do
  BEGIN
  Cpt1:=Trim(FJal1.Text) ; Cpt2:=Trim(FJal2.Text) ;
  Exo.Code:=FExercice.Value ;
  Date1:=StrToDate(FDateCpta1.Text) ; Date2:=StrToDate(FDateCpta2.Text) ;
  If FEtab.ItemIndex>0 Then Etab:=FEtab.Value ;
  If FTypeEcriture.ItemIndex>0 Then QualifPiece:=FTypeEcriture.Value ;
  If FDevise.ItemIndex>0 Then DeviseSelect:=FDevise.Value ;
  GL.NumPiece1:=StrToInt(FNumPiece1.Text) ; GL.NumPiece2:=StrToInt(FNumPiece2.Text) ;
  END ;
END ;

Procedure ChargeDevise(ListeDev : TStringList ; Code : String) ;
Var Q : TQuery ;
    St,St1,StTaux,StQuot : String ;
    Taux,Quot : Double ;
BEGIN
ListeDev.Clear ; ListeDev.Sorted:=TRUE ; ListeDev.Duplicates:=DupIgnore ;
St1:='SELECT * FROM DEVISE WHERE D_DEVISE<>"'+V_PGI.DevisePivot+'" AND D_DEVISE<>"'+V_PGI.DeviseFongible+'" ' ;
If Code<>'' Then St1:=St1+' AND D_DEVISE="'+Code+'" ' ;
Q:=OpenSQL(St1,TRUE) ;
StTaux:='12346578' ; StQuot:=StTaux ;
While Not Q.Eof Do
  BEGIN
  Taux:=Q.FindField('D_PARITEEURO').AsFloat ; Move(Taux,StTaux[1],SizeOf(Double)) ;
  Quot:=Q.FindField('D_QUOTITE').AsFloat ; Move(Quot,StQuot[1],SizeOf(Double)) ;
  St:=Q.FindField('D_DEVISE').AsString+';'+
      InttoStr(Q.FindField('D_DECIMALE').AsInteger)+';'+
      Q.FindField('D_MONNAIEIN').AsString+';'+
      StTaux+';'+StQuot+';' ;
  ListeDev.Add(St) ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

Function TrouveDev(St : String ; ListeDev : tStringList ; Var DevInOut : TDevInOut) : Boolean ;
Var i,j : Integer ;
    TabSt : Array[1..5] Of String ;
    St1,StTaux : String ;
BEGIN
Result:=FALSE ; Fillchar(DevInout,SizeOf(DevInOut),#0) ; For i:=1 To 5 Do TabSt[i]:='' ;
For i:=0 To ListeDev.Count-1 Do
  BEGIN
  j:=1 ; St1:=ListeDev[i] ;
  While (St1<>'') Do BEGIN TabSt[j]:=ReadTokenSt(St1) ; Inc(j) ; END ;
  DevInOut.Code:=TabSt[1] ;
  If DevInOut.Code=St Then
    BEGIN
    Result:=TRUE ;
    With DevInOut Do
      BEGIN
      If TabSt[2]<>'' Then Decim:=StrToInt(TabSt[2]) ;
      If TabSt[3]<>'' Then MonnaieIn:=TabSt[3]='X' ;
      If TabSt[4]<>'' Then BEGIN StTaux:=TabSt[4] ; Move(StTaux[1],TauxIn,SizeOf(Double)) ; END ;
      If TabSt[5]<>'' Then BEGIN StTaux:=TabSt[5] ; Move(StTaux[1],Quotite,SizeOf(Double)) ; END ;
      END ;
    Break ;
    END ;
  END ;
END ;

Function ClePiece(ClePiece : tClePiece) : String ;
Var StDate : String ;
BEGIN
StDate:='12345678' ;
Move(ClePiece.DateP,StDate[1],SizeOf(Double)) ;
//Result:=Format_String(X.Jal,3)+Format_String(X.Qualif,1)+FormatFloat('00000000',X.NumP) ;
Result:=Format_String(ClePiece.Jal,3)+Format_String(ClePiece.Exo,3)+Format_String(ClePiece.Qualif,1)+
        FormatFloat('0000000000',ClePiece.NumP)+Format_String(ClePiece.Dev,3)+StDate ;
END ;

Procedure DechiffreClePiece(St : String ; Var ClePiece : tClePiece) ;
Var StDate : String ;
BEGIN
Fillchar(ClePiece,SizeOf(ClePiece),#0) ;
With ClePiece Do
  BEGIN
  Jal:=Trim(Copy(St,1,3)) ; Exo:=Trim(Copy(St,4,3)) ; Qualif:=Trim(Copy(St,7,1)) ;
  NumP:=StrToInt(Trim(Copy(St,8,10))) ; Dev:=Trim(Copy(St,18,3)) ;
  StDate:=Copy(St,21,SizeOf(Double)) ; Move(StDate[1],DateP,SizeOf(Double)) ;
  END ;
END ;

Function FaitClePieceDev(Q : TQuery) : String ;
Var Cle : tClePiece ;
BEGIN
Cle.Jal:=Q.FindField('E_JOURNAL').AsString ;
Cle.Exo:=Q.FindField('E_EXERCICE').AsString ;
Cle.Dev:=Q.FindField('E_DEVISE').AsString ;
Cle.Qualif:=Q.FindField('E_QUALIFPIECE').AsString ;
Cle.NumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
Cle.DateP:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
Result:=ClePiece(Cle) ;
END ;


Procedure ChargePiece(ListePiece : TStringList ; Var CritEdt : TCritEdt) ;
Var Q : TQuery ;
    St : String ;
    Cle : tClePiece ;
BEGIN
ListePiece.Clear ; ListePiece.Sorted:=FALSE ;
St:=' Select E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_DEVISE, E_QUALIFPIECE ';
St:=St+' From ECRITURE Where' ;
St:=St+' E_JOURNAL<>"'+w_w+'" ' ;
St:=St+' And E_EXERCICE="'+CritEdt.Exo.Code+'" ' ;
if CritEdt.QualifPiece<>'' then St:=St+' And E_QUALIFPIECE="'+CritEdt.QualifPiece+'" ' ;
if CritEdt.Cpt1<>'' then St:=St+' And E_JOURNAL>="'+CritEdt.Cpt1+'"' ;
if CritEdt.Cpt2<>'' then St:=St+' And E_JOURNAL<="'+CritEdt.Cpt2+'"' ;
If CritEdt.GL.NumPiece1<>0 Then St:=St+' And E_NUMEROPIECE>='+IntToStr(CritEdt.GL.NumPiece1)+' ' ;
If CritEdt.GL.NumPiece2<>0 Then St:=St+' and E_NUMEROPIECE<='+IntTostr(CritEdt.GL.NumPiece2)+' ' ;
if CritEdt.Etab<>'' then St:=St+' And E_ETABLISSEMENT="'+CritEdt.Etab+'" ' ;
St:=St+' and E_QUALIFPIECE<>"C" ' ;
St:=St+' AND E_DEVISE<>"'+V_PGI.DevisePivot+'" AND E_DEVISE<>"'+V_PGI.DeviseFongible+'" ' ;
St:=St+' And E_DATECOMPTABLE>="'+USDATETIME(CritEdt.Date1)+'" ' ;
St:=St+' And E_DATECOMPTABLE>="'+USDATETIME(V_PGI.DateDebutEuro)+'" ' ;
St:=St+' And E_DATECOMPTABLE<="'+USDATETIME(CritEdt.Date2)+'" ' ;
{ Construction de la clause GROUP By de la SQL }
St:=St+' GROUP BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_DEVISE,E_QUALIFPIECE ' ;
Q:=OpenSQL(St,TRUE) ;
While Not Q.Eof Do
  BEGIN
  Cle.Jal:=Q.FindField('E_JOURNAL').AsString ;
  Cle.Exo:=Q.FindField('E_EXERCICE').AsString ;
  Cle.Dev:=Q.FindField('E_DEVISE').AsString ;
  Cle.Qualif:=Q.FindField('E_QUALIFPIECE').AsString ;
  Cle.NumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  Cle.DateP:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
  ListePiece.Add(ClePiece(Cle)) ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

{JP 30/05/06 : FQ 18242 : Passage en eAGL
{---------------------------------------------------------------------------------------}
Procedure UpdateAna(Q1 : TQuery ; Var Tot,TotCalc,TotE,TotCalcE : Double ; OkRepare : Boolean) ;
{---------------------------------------------------------------------------------------}
Var
  Solde : Double ;
  OkRetouche : Boolean ;
  D, C : Double ;
begin
  Solde:=Arrondi(Tot-TotCalc,V_PGI.OkDecV) ; //OkRetouche:=FALSE ;
  If (Solde<>0) then begin
    Q1.Prior ;

    {$IFDEF EAGLCLIENT}
    {JP 30/05/06 : FQ 18242 : migration eagl !!}
     D := Q1.CurrentFille.GetDouble('Y_DEBIT');
     C := Q1.CurrentFille.GetDouble('Y_CREDIT');
     OkRetouche := False;
     if Arrondi(D, V_PGI.OkDecV) <> 0 then begin
       OkRetouche := True;
       D := D + Solde;
     end
     else if Arrondi(C, V_PGI.OkDecV) <> 0 then begin
       C := C + Solde;
       OkRetouche := True;
     end;

     if OkRetouche then begin
       if OkRepare then begin
         Q1.CurrentFille.SetDouble('Y_DEBIT', D);
         Q1.CurrentFille.SetDouble('Y_CREDIT', C);
       end ;
       UpdateEcriture('Y', Q1.CurrentFille);
     end ;
    {$ELSE}
     D:=Q1.Fields[0].AsFloat ; C:=Q1.Fields[1].AsFloat ; OkRetouche:=FALSE ;
     if Arrondi(D,V_PGI.OkDecV)<>0 then begin
       OkRetouche:=TRUE ;
       D:=D+Solde ;
     end
     else if Arrondi(C,V_PGI.OkDecV)<>0 then begin
       C:=C+Solde ;
       OkRetouche:=TRUE ;
     end;
     if OkRetouche then begin
       Q1.Edit ;
       if OkRepare then begin
         Q1.Fields[0].AsFloat:=D ;
         Q1.Fields[1].AsFloat:=C ;
       end ;
       Q1.Post ;
     end ;
    {$ENDIF EAGLCLIENT}
     Q1.Next ;
   end ;
  TotCalc:=0 ;
  TotCalcE:=0 ;
END ;

Procedure TraiteAna(ClePiece : tClePiece ; Lig : Integer ; Tot,TOTE,TauxD : Double ; DevInOut : tDevInOut ; OkRepare : Boolean) ;
Var St : String ;
    D,C,DD,CD : Double ;
    Cle,OldCle : String ;
    PremFois{,OkRetouche,OkOk} : Boolean ;
    TotCalc,TotCalcE{,Solde,SoldeE} : Double ;
    Dcalc,CCalc : Double ;
    Q1 : TQuery ;
BEGIN
//OkOk:=FALSE ;
St:='SELECT Y_DEBIT, Y_CREDIT, Y_TOTALECRITURE, Y_AXE, Y_DEBITDEV, Y_CREDITDEV, Y_TAUXDEV ';
{$IFDEF EAGLCLIENT}{JP 30/05/06 : FQ 18242 : migration eagl !!}
St:=St+',Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL, Y_QUALIFPIECE ';
{$ENDIF EAGLCLIENT}
St:=St+'FROM ANALYTIQ WHERE ' ;
St:=St+'Y_JOURNAL="'+ClePiece.JAl+'" AND Y_EXERCICE="'+ClePiece.exo+'" AND Y_DATECOMPTABLE="'+UsDateTime(ClePiece.DateP)+'" AND ' ;
St:=St+'Y_NUMEROPIECE='+IntToStr(ClePiece.NumP)+' AND Y_NUMLIGNE='+IntToStr(Lig)+' AND Y_QUALIFPIECE="'+ClePiece.Qualif+'" ' ;
St:=St+' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE ' ;
Q1:=OpenSQL(St,FALSE) ;
{$IFNDEF EAGLCLIENT}
Q1.UpdateMode:=UpWhereChanged ;
// YMO 05/01/2006 Je ne vois pas l'utilité vu que la clause WHERE est inscrite dans St,
// et qu'il n'y a pas d'Update ; ce n'est donc pas remplacé en CWAS
{$ENDIF}

(*Tot:=0 ;*) TotCalc:=0 ; TotCalcE:=0 ; PremFois:=TRUE ; //SoldeE:=0 ; Solde:=0 ;
  While Not Q1.EOF do
  BEGIN
    {JP 29/09/05 : FQ 16712 : J'ai une peu de mal à comprendre pourquoi la rupture sur l'axe avait été remplacée
                   par une rupture sur le taux de la devise puisque cette fonction est appelé par ligne d'écriture,
                   donc sauf erreur de ma part, Y_TAUXDEV est toujours le même.
     Cle:=Q1.Fields[6].AsString ;}
    Cle:=Q1.FindField('Y_AXE').AsString ;

    If PremFois Then OldCle:=Cle ;
    PremFois:=FALSE ; //OkOk:=TRUE ;
    If (Cle<>OldCle) Then UpdateAna(Q1,Tot,TotCalc,TotE,TotCalcE,OkRepare) ;
    D:=Arrondi(Q1.FindField('Y_DEBIT').AsFloat,V_PGI.OkDecV)    ; C:=Arrondi(Q1.FindField('Y_CREDIT').AsFloat,V_PGI.OkDecV) ;
    DD:=Arrondi(Q1.FindField('Y_DEBITDEV').AsFloat,DevInOut.Decim) ; CD:=Arrondi(Q1.FindField('Y_CREDITDEV').AsFloat,DevInOut.Decim) ;

    {$IFDEF EAGLCLIENT}
    {JP 30/05/06 : FQ 18242 : migration eagl !!}
    Q1.CurrentFille.SetDouble('Y_TOTALECRITURE', Tot);
    DCalc:=DeviseToEuro(DD,TauxD,DevInOut.Quotite) ; CCalc:=DeviseToEuro(CD,TauxD,DevInOut.Quotite) ;
    TotCalc:=Arrondi(TotCalc+DCalc+CCalc,V_PGI.OkDecV) ;
    If OkRepare then begin
      If Arrondi(DCalc-D,V_PGI.OkDecV)<>0 Then Q1.CurrentFille.SetDouble('Y_DEBIT' , DCalc) ;
      If Arrondi(CCalc-C,V_PGI.OkDecV)<>0 Then Q1.CurrentFille.SetDouble('Y_CREDIT', CCalc) ;
      Q1.CurrentFille.SetDouble('Y_TAUXDEV', TauxD);
    end ;
    UpdateEcriture('Y', Q1.CurrentFille);
    {$ELSE}
    Q1.Edit ;
    Q1.Fields[2].AsFloat:=Tot ;
    DCalc:=DeviseToEuro(DD,TauxD,DevInOut.Quotite) ; CCalc:=DeviseToEuro(CD,TauxD,DevInOut.Quotite) ;
    TotCalc:=Arrondi(TotCalc+DCalc+CCalc,V_PGI.OkDecV) ;
    If OkRepare then begin
      If Arrondi(DCalc-D,V_PGI.OkDecV)<>0 Then Q1.FindField('Y_DEBIT').AsFloat:=DCalc ;
      If Arrondi(CCalc-C,V_PGI.OkDecV)<>0 Then Q1.FindField('Y_CREDIT').AsFloat:=CCalc ;
      Q1.FindField('Y_TAUXDEV').AsFloat:=TauxD ;
    end ;
    Q1.Post ;
    {$ENDIF EAGLCLIENT}
    OldCle:=Cle ;
    Q1.Next ;
  end ;
UpdateAna(Q1,Tot,TotCalc,TotE,TotCalcE,OkRepare) ;
Ferme(Q1) ;
END ;

Function RecordErr(ClePiece : tClePiece ; Ref : String ; Lig : Integer ; Rem : String) : DelInfo ;
Var X : DelInfo ;
BEGIN
//Inc(NbError) ;
X:=DelInfo.Create ;
X.LeCod:=ClePiece.Jal ; X.LeLib:=Ref ;
X.LeMess:=IntToStr(ClePiece.Nump)+'/'+IntToStr(Lig) ;
X.LeMess2:=DateToStr(ClePiece.DateP) ;
X.LeMess3:=Rem ;
X.LeMess4:=ClePiece.Exo ;
Result:=X ;
END ;

{JP 30/05/06 : FQ 18242 : Passage en eAGL
{---------------------------------------------------------------------------------------}
Procedure TraiteUnePiece(StPiece : String ; ListeDev : tStringList ; ListeErr : TList ; OkRepare,OkChancell,OkMess : Boolean ; TauxModifie : Double) ;
{---------------------------------------------------------------------------------------}
Var EstIn : Boolean ;
    ClePiece : TClePiece ;
    Q : TQuery ;
    St,StErr,St1 : String ;
    PremFois : Boolean ;
    DevInOut : TDevInOut ;
    D,C,DE,CE,DD,CD,TauxD,Derr,Cerr,DEErr,CEErr : Double ;
    Dcalc,CCalc,DECalc,CECalc,TauxCalc : Double ;
    TotD,TotC,TotDE,TotCE : Double ;
    Delta,DeltaE : Double ;
    OnDeb : Boolean ;
    ModifD,ModifC,ModifDE,ModifCE : Boolean ;
    DateTaux : TDateTime ;
    G : DelInfo ;
begin
  TauxCalc:=0; TauxD:=0; OnDeb:= False; DE := 0; CE := 0;
  DechiffreClePiece(StPiece,ClePiece) ;
  If ClePiece.DateP<V_PGI.DateDebutEuro Then Exit ;
  St:='Select E_GENERAL, E_AUXILIAIRE, E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV, E_DEVISE, E_ANA,' ;
  St:=St+'E_COUVERTURE, E_COUVERTUREDEV, E_LETTRAGEDEV, E_TAUXDEV,E_COTATION, ' ;
  St:=St+'E_NUMLIGNE, E_REFINTERNE, E_DATETAUXDEV, E_NATUREPIECE ' ;
  {$IFDEF EAGLCLIENT} {JP 30/05/06 : FQ 18242 : migration eagl !!}
  St:=St+', E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMECHE, E_QUALIFPIECE ';
  {$ENDIF EAGLCLIENT}
  St:=St+'from Ecriture where E_JOURNAL="'+ClePiece.Jal+'"' ;
  St:=St+' AND E_EXERCICE="'+ClePiece.Exo+'"' ;
  St:=St+' AND E_DATECOMPTABLE="'+USDateTime(ClePiece.DateP)+'"' ;
  St:=St+' AND E_NUMEROPIECE='+IntToStr(ClePiece.NumP) ;
  St:=St+' AND E_QUALIFPIECE="'+ClePiece.Qualif+'" ' ;
  St:=St+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ' ;
  Q:=OpenSQL(St,FALSE) ; PremFois:=TRUE ;
  TotD:=0 ; TotC:=0 ; TotDE:=0 ; TotCE:=0 ;

  while not Q.Eof do begin
    If PremFois Then begin
      if not TrouveDev(Q.FindField('E_DEVISE').AsString,ListeDev,DevInOut) Then Break ;
      If Q.FindField('E_NATUREPIECE').AsString='ECC' Then Break ;
    end ;
    DD:=Arrondi(Q.FindField('E_DEBITDEV').AsFloat,DevInOut.Decim) ;
    CD:=Arrondi(Q.FindField('E_CREDITDEV').AsFloat,DevInOut.Decim) ;
    D:=Arrondi(Q.FindField('E_DEBIT').AsFloat,V_PGI.OkDecV) ;
    C:=Arrondi(Q.FindField('E_CREDIT').AsFloat,V_PGI.OkDecV) ;
    DErr:=0 ; CErr:=0 ; DEErr:=0 ; CEErr:=0 ;
    OnDeb:=DD<>0 ;
    TauxD:=Q.FindField('E_TAUXDEV').AsFloat ;
    If DevInOut.MonnaieIn Then
      TauxCalc := V_PGI.TauxEuro/DevInOut.TauxIn
    else begin
      If Arrondi(TauxModifie,9)<>0 Then TauxCalc:=TauxModifie Else
        If OkChancell then begin
          If PremFois Then TauxCalc:=GetTaux(ClePiece.Dev,DateTaux,ClePiece.DateP) ;
          If TauxCalc=1 Then TauxCalc:=V_PGI.TauxEuro*Q.FindField('E_COTATION').AsFloat ;
        end
        else
          TauxCalc:=V_PGI.TauxEuro*Q.FindField('E_COTATION').AsFloat ;
    end ;

    If Arrondi(TauxCalc,9)=0 Then Break ;
    DCalc  := DeviseToEuro(DD, TauxCalc, DevInOut.Quotite);
    CCalc  := DeviseToEuro(CD, TauxCalc, DevInOut.Quotite);
    ModifD := False; ModifC := False; ModifDE := False; ModifCE := False;

    {$IFDEF EAGLCLIENT}
    {JP 30/05/06 : FQ 18242 : migration eagl !!}
    if Arrondi(DCalc-D,V_PGI.OkDecV) <> 0 then begin
      ModifD := True;
      DErr   := D;
      D      := DCalc;
      if OkRepare Then Q.CurrentFille.SetDouble('E_DEBIT', D);
    end;

    if Arrondi(CCalc-C,V_PGI.OkDecV) <> 0 then begin
      ModifC := True;
      CErr   := C;
      C      := CCalc;
      if OkRepare Then Q.CurrentFille.SetDouble('E_CREDIT', C);
    end;

    if Arrondi(TauxD-TauxCalc,9) <> 0 then begin
      TauxD := TauxCalc;
      if OkRepare Then Q.CurrentFille.SetDouble('E_TAUXDEV', TauxD);
    end;
    SetCotationDB(ClePiece.DateP, Q.CurrentFille);
    UpdateEcriture('E', Q.CurrentFille);
    {$ELSE}
    Q.Edit;
    if Arrondi(DCalc-D,V_PGI.OkDecV) <> 0 then begin
      ModifD := True;
      DErr   := D;
      D      := DCalc;
      if OkRepare Then Q.FindField('E_DEBIT').AsFloat := D;
    end;

    if Arrondi(CCalc-C,V_PGI.OkDecV) <> 0 then begin
      ModifC := True;
      CErr   := C;
      C      := CCalc;
      if OkRepare then Q.FindField('E_CREDIT').AsFloat := C;
    end;

    if Arrondi(TauxD-TauxCalc,9) <> 0 then begin
      TauxD := TauxCalc;
      if OkRepare then Q.FindField('E_TAUXDEV').AsFloat := TauxD;
    end;
    SetCotationDB(ClePiece.DateP, Q);
    Q.Post ;
    {$ENDIF EAGLCLIENT}

    PremFois:=False ;
    TotD:=Arrondi(TotD+D,V_PGI.OkDecV) ; TotC:=Arrondi(TotC+C,V_PGI.OkDecV) ;
    TotDE:=Arrondi(TotDE+DE,V_PGI.OkDecV) ; TotCE:=Arrondi(TotCE+CE,V_PGI.OkDecV) ;
    If OkMess And (ListeErr<>NIL) And (ModifD Or ModifC Or ModifDE Or ModifCE) then begin
      StErr:='' ;
      If ModifD Or ModifC then begin
        St1:=' Montants lus : '+FloatToStr(Derr+CErr)+' / Montants calculés : '+FloatToStr(D+C) ;
        StErr:='Ecart sur les montants pivots '+St1 ;
        ListeErr.Add(RecordErr(ClePiece,Q.FindField('E_REFINTERNE').AsString,Q.FindField('E_NUMLIGNE').AsInteger,StErr));
      end ;

      StErr:='' ;
      If ModifDE Or ModifCE then begin
        St1:=' Montants lus : '+FloatToStr(DEerr+CEErr)+' / Montants calculés : '+FloatToStr(DE+CE) ;
        StErr:='Ecart sur les montants opposés'+St1 ;
        ListeErr.Add(RecordErr(ClePiece,Q.FindField('E_REFINTERNE').AsString,Q.FindField('E_NUMLIGNE').AsInteger,StErr));
      end ;
    end; {If OkMess And (}

    If (Q.FindField('E_ANA').AsString='X') And (ModifD Or ModifC Or ModifDE Or ModifCE) And OkRepare Then
      TraiteAna(ClePiece,Q.FindField('E_NUMLIGNE').Asinteger,D+C,DE+CE,TauxD,DevInOut,OkRepare) ;

    Q.Next ;
  end; {while not Q.Eof }

  Delta:=Arrondi(TotD-TotC,V_PGI.okDecV) ; DeltaE:=Arrondi(TotDE-TotCE,V_PGI.okDecE) ;
  if ((Delta<>0) or (DeltaE<>0)) and OkRepare then begin
    Q.Last ;
    {$IFDEF EAGLCLIENT}
    {JP 30/05/06 : FQ 18242 : migration eagl !!}
    if Delta <> 0 then begin
      if OnDeb then
        Q.CurrentFille.SetDouble('E_DEBIT', Q.FindField('E_DEBIT').AsFloat - Delta)
      else
        Q.CurrentFille.SetDouble('E_CREDIT', Q.FindField('E_CREDIT').AsFloat + Delta);
    end;
    UpdateEcriture('E', Q.CurrentFille);
    {$ELSE}
    Q.Edit ;
    if Delta <> 0 then begin
      if OnDeb then
        Q.FindField('E_DEBIT').AsFloat:=Q.FindField('E_DEBIT').AsFloat-Delta
      else
        Q.FindField('E_CREDIT').AsFloat:=Q.FindField('E_CREDIT').AsFloat+Delta ;
    end ;
    Q.Post ;
    {$ENDIF EAGLCLIENT}

    D:=Q.FindField('E_DEBIT').AsFloat ; C:=Q.FindField('E_CREDIT').AsFloat ;
    If (Q.FindField('E_ANA').AsString='X') And OkRepare Then
      TraiteAna(ClePiece,Q.FindField('E_NUMLIGNE').Asinteger,D+C,DE+CE,TauxD,DevInOut,OkRepare) ;
  end; {if ((Delta<>0) }

  Ferme(Q);
end ;

Procedure TraiteLesPieces(ListePiece,ListeDev : tStringList ; OkRepare,OkChancell,OkMess : Boolean) ;
Var i : Integer ;
    ListeErr : TList ;
    HH : Boolean ;
BEGIN
InitMove(ListePiece.Count,'') ; If ListePiece=NIL Then Exit ;
ListeErr:=TList.Create ;
For i:=0 To ListePiece.Count-1 Do
  BEGIN
  MoveCur(FALSE) ; TraiteUnePiece(ListePiece[i],ListeDev,ListeErr,OkRepare,OkChancell,OkMess,0) ;
  END ;
FiniMove ;
HH:=TRUE ; If OkMess And (ListeErr<>NIL) And (ListeErr.Count>0) Then RapportdErreurMvt(ListeErr,3,HH,False) ;
VideListe(ListeErr) ; ListeErr.Free ;
END ;


procedure TFRepareMontant.FormClose(Sender: TObject;  var Action: TCloseAction);
begin
ListePiece.Free ; ListeDev.Free ;
if Parent is THPanel then
   BEGIN
   _DeblocageMonoPoste(True) ;
   Action:=caFree ;
   END ;
end;

procedure TFRepareMontant.FExerciceChange(Sender: TObject);
begin
ExoToDates(FExercice.Value,FDateCpta1,FDateCpta2) ;
end;

procedure TFRepareMontant.BValiderClick(Sender: TObject);
Var i : Integer ;
begin
If FOkRepare.Checked Then i:=HM.Execute(1,'','') Else i:=HM.Execute(0,'','') ;
If i<>mrYes Then Exit ;
ChargeDevise(ListeDev,'') ;
RecupCrit ;
If (CritEdt.Date1<V_PGI.DateDebutEuro) And FOkRepare.Checked Then BEGIN HM.Execute(2,'','') ; FDateCpta1.SetFocus ; Exit ; END ;
ChargePiece(ListePiece,CritEdt) ;
TraiteLesPieces(ListePiece,ListeDev,FOkRepare.Checked,FChancell.Checked,TRUE) ;
end;

procedure TFRepareMontant.FormCreate(Sender: TObject);
begin
ListePiece:=TStringList.Create ;
ListeDev:=TStringList.Create ;
end;

procedure TFRepareMontant.BChancelOutClick(Sender: TObject);
Var St : String ;
begin
St:='' ;
If (FDevise.ItemIndex>0) And (FDevise.Value<>V_PGI.DevisePivot) And (FDevise.Value<>V_PGI.DeviseFongible) {And
   (Not EstMonnaieIn(FDevise.Value))} Then St:=FDevise.Value ;
FicheChancelSur2Dates(St,StrToDate(FDateCpta1.text),StrToDate(FDateCpta2.text),taModif,True) ;
end;



Function TrouvePieceDevise(TobOrig : TOB ; Var IdentPD : tIdentPieceDev ; Var Dev : RDevise ; Simul : String3 ; Var ListeCle : TStringList ; Var Comment : tActionFiche ; Var Mess : String) : Boolean ;
Var Q1 : TQuery ;
    Trouv : boolean ;
    StWhat : String ;
    OkV : Boolean ;
    i : integer ;
    TOBL : TOB ;
    FF   : TField ;
    DD : tDateTime ;

begin

Fillchar(IdentPD,SizeOf(IdentPD),#0) ; Fillchar(Dev,SizeOf(Dev),#0) ;
Result := False;          {FP 17/07/2006 Retire un avertissement lors de la compilation}

Mess:='' ;
For i:=0 to TOBOrig.Detail.Count-1 do
BEGIN
TOBL  := TOBOrig.Detail[i] ;

TrouvePieceDevise:=FALSE ;

StWhat:=' E_JOURNAL, E_EXERCICE, E_QUALIFPIECE, E_DATECOMPTABLE, E_REFINTERNE,'+
        ' E_LIBELLE, E_DEVISE, E_TAUXDEV, E_COTATION, E_NATUREPIECE, E_NUMEROPIECE,'+
        ' E_DATETAUXDEV, E_LETTRAGE, E_ETATLETTRAGE, E_REFPOINTAGE, E_DATEPOINTAGE,'+
        ' E_VALIDE, E_ECRANOUVEAU'; {FP 17/07/2006}
If Simul='' Then
  Q1:=OpenSQL('Select '+StWhat+' FROM ECRITURE WHERE E_JOURNAL="'+TOBL.GetString('E_JOURNAL')+'"'
            +' AND E_EXERCICE="'+QuelExo(TOBL.GetString('E_DATECOMPTABLE'))+'"'
            +' AND E_DATECOMPTABLE="'+USDATETIME(TOBL.GetDateTime('E_DATECOMPTABLE'))+'"'
            +' AND E_NUMEROPIECE='+TOBL.GetString('E_NUMEROPIECE'),True)
            Else
  Q1:=OpenSQL('Select '+StWhat+' FROM ECRITURE WHERE E_JOURNAL="'+TOBL.GetString('E_JOURNAL')+'"'
            +' AND E_EXERCICE="'+QuelExo(TOBL.GetString('E_DATECOMPTABLE'))+'"'
            +' AND E_DATECOMPTABLE="'+USDATETIME(TOBL.GetDateTime('E_DATECOMPTABLE'))+'"'
            +' AND E_NUMEROPIECE='+TOBL.GetString('E_NUMEROPIECE')
            +' AND E_QUALIFPIECE="'+Simul+'" ',True) ;

Trouv:=Not Q1.EOF ;
OkV:=False ;

if Trouv then
  BEGIN
  if (TOBL.GetIndex = 0) then
  begin       // on remplit les variables une fois, pour la 1ere ecriture
    IdentPD.JalP:=Q1.FindField('E_JOURNAL').AsString;
    IdentPD.DateP:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ;
    IdentPD.NatP:=Q1.FindField('E_NATUREPIECE').AsString ;
    IdentPD.RefP:=Q1.FindField('E_REFINTERNE').AsString ;
    IdentPD.LibP:=Q1.FindField('E_LIBELLE').AsString ;
    IdentPD.NumP:=Q1.FindField('E_NUMEROPIECE').AsInteger ;
    IdentPD.DateMin:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ;
    IdentPD.DateMax:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ;
    Dev.DateTaux:=Q1.FindField('E_DATETAUXDEV').AsDateTime ;
    Dev.Taux:=Q1.FindField('E_TAUXDEV').AsFloat ;
    Dev.Cotation:=Q1.FindField('E_COTATION').AsFloat ;
    Dev.Code:=Q1.FindField('E_DEVISE').AsString ;
  end
  else
  begin   // si les autres écritures ont des infos différentes, on initialise celles-ci
    if IdentPD.JalP<>Q1.FindField('E_JOURNAL').AsString then IdentPD.JalP := '';
    if IdentPD.DateP<>Q1.FindField('E_DATECOMPTABLE').AsDateTime then IdentPD.DateP := iDate1900 ;
    if IdentPD.NatP<>Q1.FindField('E_NATUREPIECE').AsString then IdentPD.NatP := '' ;
    if IdentPD.RefP<>Q1.FindField('E_REFINTERNE').AsString then IdentPD.RefP := '' ;
    if IdentPD.LibP<>Q1.FindField('E_LIBELLE').AsString then IdentPD.LibP := '' ;
    if IdentPD.NumP<>Q1.FindField('E_NUMEROPIECE').AsInteger then IdentPD.NumP := 0 ;
    IdentPD.DateMin:=Min(IdentPD.DateMin, Q1.FindField('E_DATECOMPTABLE').AsDateTime) ;
    IdentPD.DateMax:=Max(IdentPD.DateMax, Q1.FindField('E_DATECOMPTABLE').AsDateTime) ;
    if Dev.DateTaux<>Q1.FindField('E_DATETAUXDEV').AsDateTime then Dev.DateTaux := iDate1900 ;
    if Dev.Taux<>Q1.FindField('E_TAUXDEV').AsFloat then Dev.Taux := 0 ;
    if Dev.Cotation<>Q1.FindField('E_COTATION').AsFloat then Dev.Cotation := 0 ;
    if Dev.Code<>Q1.FindField('E_DEVISE').AsString then Dev.Code := '' ;
  end;

  ListeCle.Add(FaitClePieceDev(Q1));

  END ;
{
if (Q1.FindField('E_LETTRAGE').AsString<>'')
Or (Q1.FindField('E_ETATLETTRAGE').AsString='PL')
Or (Q1.FindField('E_ETATLETTRAGE').AsString='TL') Then
BEGIN
   Comment:=taConsult ; OkV:=True ;
   If Mess = '' then
      Mess:=TraduireMemoire('Taux non modifiable : l''un des mouvements est lettré')
   else
      Mess:=TraduireMemoire('Taux non modifiable : plusieurs mouvements sont lettrés')

END ;

if Not OKV then
BEGIN
}
   Q1.First ;
   While Not Q1.EOF do
   BEGIN
      FF:=Q1.FindField('E_DATECOMPTABLE') ; if FF=Nil then Exit ;
      DD:=FF.AsDateTime ;
      If (DD<VH^.EnCours.Deb) Or ((VH^.DateCloturePer>0) and (DD<=VH^.DateCloturePer))
      Or (Q1.FindField('E_LETTRAGE').AsString<>'')
      Or (Q1.FindField('E_ETATLETTRAGE').AsString='PL')
      Or (Q1.FindField('E_ETATLETTRAGE').AsString='TL')
      {b FP 17/07/2006: Ajoute la condition sur les reports à nouveau car la compensation n'exclue pas ces écritures contrairement à la modification de pièce}
      Or (Q1.FindField('E_ECRANOUVEAU').AsString='H')
      Or (Q1.FindField('E_ECRANOUVEAU').AsString='OAN')
      {e FP 17/07/2006}
      Or (Q1.FindField('E_REFPOINTAGE').AsString<>'')
      Or (Q1.FindField('E_VALIDE').AsString='X') Then
      BEGIN
          OKV:=True ;
          Break ;
      END ;
      Q1.Next ;
   END ;
//END ;

Ferme(Q1) ;
if OkV then
BEGIN
   Comment:=taConsult ;
   If Mess = '' then
      Mess:=TraduireMemoire('Votre sélection contient une pièce non modifiable')
   else
      Mess:=TraduireMemoire('Votre sélection contient des pièces non modifiables')
END ;

TrouvePieceDevise:=Trouv ;

END;

FreeAndNil(TOBL) ;
END;

procedure ChangeLeTauxDevise(TOBOrig : TOB ; Comment : tActionFiche ) ;
Var IdentPD : tIdentPieceDev ;
    Dev : RDevise ;
    ListeDev, ListeCle : TStringList ;
    Mess : String ;
    i : integer ;
begin

//YMO   16/12/2005  Plusieurs enregs passés en param
if TOBOrig<>nil then
//If TOBORig.Detail.Count>0 then
begin
  ListeCle:=TStringList.Create ;
  ListeCle.Clear ;  ListeCle.Sorted:=TRUE ;  ListeCle.Duplicates:=DupIgnore ;
  If TrouvePieceDevise(TOBOrig,IdentPD,Dev,'',ListeCle,Comment,Mess) then
  If (ListeCle.Count > 0) And SaisieNewTaux1(Dev,IdentPD,Comment,Mess) Then
  If Comment<>taConsult Then
  BEGIN
    ListeDev:=TStringList.Create ;
    ChargeDevise(ListeDev,Dev.Code) ;
    For i:=0 To ListeCle.Count-1 Do
    Begin // Maj de ttes les pieces de la tob
      TraiteUnePiece(ListeCle[i],ListeDev,Nil,TRUE,FALSE,FALSE,Dev.Taux) ;
    End;
    ListeDev.Free ;
  END ;
  ListeCle.Free ;
end;

end;

{$IFDEF EAGLCLIENT}
{JP 30/05/06 : FQ 18242 : migration eagl !!}
{---------------------------------------------------------------------------------------}
procedure UpdateEcriture(Pref : Char; Q : TOB);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  case Pref of
    'E' : begin
            SQL := 'UPDATE ECRITURE SET E_COTATION = ' + StrFPoint(Q.GetDouble('E_COTATION')) +
                   ', E_TAUXDEV = ' + StrFPoint(Q.GetDouble('E_TAUXDEV')) +
                   ', E_DEBIT = '   + StrFPoint(Q.GetDouble('E_DEBIT')) +
                   ', E_CREDIT = '  + StrFPoint(Q.GetDouble('E_CREDIT'));
            SQL := SQL + ' WHERE E_JOURNAL = "'     + Q.GetString('E_JOURNAL') + '"';
            SQL := SQL + ' AND E_EXERCICE = "'      + Q.GetString('E_EXERCICE') + '"';
            SQL := SQL + ' AND E_DATECOMPTABLE = "' + USDateTime(Q.GetDateTime('E_DATECOMPTABLE')) + '"';
            SQL := SQL + ' AND E_NUMEROPIECE = '    + Q.GetString('E_NUMEROPIECE') ;
            SQL := SQL + ' AND E_QUALIFPIECE = "'   + Q.GetString('E_QUALIFPIECE') + '"';
            SQL := SQL + ' AND E_NUMLIGNE = '       + Q.GetString('E_NUMLIGNE');
            SQL := SQL + ' AND E_NUMECHE = '        + Q.GetString('E_NUMECHE');
          end;
    'Y' : begin
            SQL := 'UPDATE ANALYTIQ SET Y_TOTALECRITURE = ' + StrFPoint(Q.GetDouble('Y_TOTALECRITURE')) +
                   ', Y_TAUXDEV = ' + StrFPoint(Q.GetDouble('Y_TAUXDEV')) +
                   ', Y_DEBIT = '   + StrFPoint(Q.GetDouble('Y_DEBIT')) +
                   ', Y_CREDIT = '  + StrFPoint(Q.GetDouble('Y_CREDIT'));
            SQL := SQL + ' WHERE Y_JOURNAL = "'     + Q.GetString('Y_JOURNAL') + '"';
            SQL := SQL + ' AND Y_EXERCICE = "'      + Q.GetString('Y_EXERCICE') + '"';
            SQL := SQL + ' AND Y_DATECOMPTABLE = "' + USDateTime(Q.GetDateTime('Y_DATECOMPTABLE')) + '"';
            SQL := SQL + ' AND Y_NUMEROPIECE = '    + Q.GetString('Y_NUMEROPIECE') ;
            SQL := SQL + ' AND Y_QUALIFPIECE = "'   + Q.GetString('Y_QUALIFPIECE') + '"';
            SQL := SQL + ' AND Y_AXE = "'           + Q.GetString('Y_AXE') + '"';
            SQL := SQL + ' AND Y_NUMLIGNE = '       + Q.GetString('Y_NUMLIGNE');
            SQL := SQL + ' AND Y_NUMVENTIL = '      + Q.GetString('Y_NUMVENTIL');
          end;
  end;
  ExecuteSQL(SQL);
end;
{$ENDIF EAGLCLIENT}

end.

