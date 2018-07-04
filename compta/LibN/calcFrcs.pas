unit calcFrcs;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,Saisutil,
  DB, DBTables, Mask, StdCtrls, Hctrls, Hcompte, Buttons, ExtCtrls,ENT1,HENT1, HStatus,
  HSysMenu, SoldeCpt, hmsgbox, HTB97, UiUtil, HPanel, SaisComm, ParamSoc, UtilPGI ;

procedure RecalculMontantFrancs ;
Function RecalculMontantFrancsRV : Boolean ;

type
  TFcalcFrcs = class(TForm)
    PCrit: TPanel;
    HPB: TToolWindow97;
    BStop: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    TFJal: THLabel;
    FJal1: THCpteEdit;
    TFaJ: TLabel;
    FJal2: THCpteEdit;
    FExercice: THValComboBox;
    TFExercice: THLabel;
    TFDateCpta1: THLabel;
    FDateCpta1: TMaskEdit;
    TFDateCpta2: TLabel;
    FDateCpta2: TMaskEdit;
    FEtab: THValComboBox;
    TFEtab: THLabel;
    Fcalc: TCheckBox;
    FEqui: TCheckBox;
    HMTrad: THSystemMenu;
    Panel3: TPanel;
    TDateDebEuro: THLabel;
    TTauxE: THLabel;
    DateDebEuro: THCritMaskEdit;
    TauxE: TEdit;
    HLabel1: THLabel;
    OldDev: TEdit;
    Hmess: THMsgBox;
    Dock: TDock97;
    FDelEc: TCheckBox;
    TTravail: TLabel;
    Timer1: TTimer;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Déclarations privées }
    PourRV,PbRV : Boolean ;
    Procedure RecalculEuro(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
    Procedure RecalculAnalEuro(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
    Procedure EquilibreAnalEuro(SJal,SExo,SNum,SLig,SQual : String ; DDate : TDateTime  ; TotalFrancsG : Double) ;
    Procedure BalayeEcrPourEquilibreAnal(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
    Procedure EquilibreEuro(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses CPTESAV ;

procedure RecalculMontantFrancs ;
var FFrcs : TFcalcFrcs ;
    PP : THPanel ;
BEGIN
if Not EuroOK then Exit ;
if Not BlocageMonoPoste(True) then Exit ;
FFrcs:=TFcalcFrcs.Create(Application) ;
FFrcs.PourRV:=FALSE ;
FFrcs.PbRV:=FALSE ;
FFrcs.BAide.Visible:=False ;
FFrcs.Timer1.Enabled:=FALSE ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FFrcs.ShowModal ;
    finally
     FFrcs.Free ;
     DeblocageMonoPoste(True) ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FFrcs,PP) ;
   FFrcs.Show ;
   END ;
END ;

Function RecalculMontantFrancsRV : Boolean ;
var FFrcs : TFcalcFrcs ;
BEGIN
Result:=FALSE ;
FFrcs:=TFcalcFrcs.Create(Application) ;
FFrcs.Timer1.Enabled:=TRUE ;
try
 FFrcs.PbRV:=FALSE ;
 FFrcs.PourRV:=TRUE ;
 FFrcs.BAide.Visible:=False ;
 FFrcs.ShowModal ;
finally
 Result:=FFrcs.PbRV ;
 FFrcs.Free ;
end ;
Screen.Cursor:=SyncrDefault ;
END ;

Function FormateTauxEuro(X : Double) : String ;
BEGIN
Result:=StrS(X,6) ;
END ;

procedure TFcalcFrcs.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FJal1.Text:='' ; FJal2.Text:='' ;
FEtab.ItemIndex:=0 ;
FExercice.Value:=VH^.Encours.Code ;
DateDebEuro.Text:=DateToStr(V_PGI.DateDebutEuro) ;
TauxE.Text:=FormateTauxEuro(V_PGI.TauxEuro) ;
OldDev.Enabled:=Not VH^.TenueEuro ;
If VH^.TenueEuro Then OldDev.Text:=V_PGI.DeviseFongible Else OldDev.Text:='' ;
OldDev.Visible:=FALSE ; HLabel1.Visible:=FALSE ;
PCrit.Visible:=TRUE ; Caption:=HMess.Mess[0] ;
BValider.Hint:=HMess.Mess[9] ;
If PourRV Then
  BEGIN
  FExercice.Value:=VH^.EnCours.Code ; FCalc.Checked:=FALSE ; FEqui.Checked:=TRUE ; FDelEc.Checked:=FALSE ;
  END ;
end;

Procedure DetruitEcart(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
Var St,Cpt1,Cpt2,St1,Jal : String ;
{$IFDEF SPEC302}
    Q : TQuery ;
{$ENDIF}
BEGIN
Cpt1:='' ; Cpt2:='' ; St1:='' ; Jal:='' ;
{$IFDEF SPEC302}
Q:=OpenSQL('SELECT SO_ECCEUROCREDIT,SO_ECCEURODEBIT,SO_JALECARTEURO FROM SOCIETE',TRUE) ;
If Not Q.Eof Then
  BEGIN
  Cpt1:=Q.Fields[0].AsString ; Cpt2:=Q.Fields[1].AsString ; Jal:=Q.Fields[2].AsString ;
  END ;
Ferme(Q) ;
{$ELSE}
Cpt1:=GetParamsocSecur('SO_ECCEURODEBIT','') ; Cpt2:=GetParamsocSecur('SO_ECCEUROCREDIT','') ;
Jal:=GetParamsocSecur('SO_JALECARTEURO','') ;
{$ENDIF}
If (Cpt1='') And (Cpt2='') Then Exit ;
If Jal='' Then Exit ;
If (Cpt1<>Cpt2) Then St1:='(E_GENERAL="'+Cpt1+'" OR E_GENERAL="'+Cpt2+'") '
                Else St1:='E_GENERAL="'+Cpt1+'" ' ;
St:='DELETE FROM ECRITURE WHERE '+St1+' AND E_JOURNAL<>"'+Jal+'" AND E_CREERPAR<>"GEN" '
    +' AND (E_EXERCICE="'+VH^.EnCours.Code+'" OR E_EXERCICE="'+VH^.Suivant.Code+'") ' ;
ExecuteSQL(St) ;
END ;

Procedure TFcalcFrcs.RecalculEuro(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
Var St : String ;
    DF,CF,LF,DE,CE,LE : Double ;
    Q : TQuery ;
    i : Integer ;
BEGIN
(*
St:='select * from ecriture Where E_JOURNAL<>"###" ' ;
If FJal1.Text<>'' Then St:=St+' AND E_JOURNAL>="'+FJal1.Text+'" ' ;
If FJal2.Text<>'' Then St:=St+' AND E_JOURNAL<="'+FJal2.Text+'" ' ;
If FExercice.ItemIndex>0 Then St:=ST+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
St:=St+'And E_DATECOMPTABLE>="'+usdate(FDateCpta1)+'" And E_DATECOMPTABLE<="'+usdate(FDateCpta2)+'" ' ;
If FEtab.ItemIndex>0 Then St:=ST+' AND E_ETABLISSEMENT="'+FETAB.Value+'" ' ;
St:=St+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ' ;
*)
St:='SELECT E_DEBIT, E_CREDIT, E_COUVERTURE, E_DEBITEURO, E_CREDITEURO, E_COUVERTUREEURO, ' ;
St:=St+'E_ECHEENC1, E_ECHEENC2, E_ECHEENC3, E_ECHEENC4, E_ECHEDEBIT, E_NUMEROPIECE, E_NUMLIGNE, E_GENERAL, E_NUMECHE ' ;
St:=St+'from ecriture Where E_JOURNAL<>"###" ' ;
If Jal1<>'' Then St:=St+' AND E_JOURNAL>="'+Jal1+'" ' ;
If Jal2<>'' Then St:=St+' AND E_JOURNAL<="'+Jal2+'" ' ;
If Exo<>'' Then St:=ST+' AND E_EXERCICE="'+Exo+'" ' ;
If Date1>0 Then St:=St+'AND E_DATECOMPTABLE>="'+usdateTime(Date1)+'" ' ;
If Date2>0 Then St:=St+'AND E_DATECOMPTABLE<="'+usdateTime(Date2)+'" ' ;
If Etab<>'' Then St:=ST+' AND E_ETABLISSEMENT="'+Etab+'" ' ;
St:=St+'AND (E_EXERCICE="'+VH^.EnCours.code+'" OR E_EXERCICE="'+VH^.Suivant.code+'") ' ;
// St:=St+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ' ;
InitMove(10000,'') ;
Q:=OpenSQL(St,FALSE) ; //Q.UpdateMode:=UpWhereChanged ;
Q.UpdateMode:=UpWhereKeyOnly ;
FetchSqlODBC(Q) ;
While Not Q.EOF do
   BEGIN
   MoveCur(FALSE) ;
   Q.Edit ;
   DE:=Q.Fields[0].AsFloat ; CE:=Q.Fields[1].AsFloat ; LE:=Q.Fields[2].AsFloat ;
   DF:=EuroToPivot(DE) ; CF:=EuroToPivot(CE) ; LF:=EuroToPivot(LE) ;
   If DF<>0 Then Q.Fields[3].AsFloat:=DF ;
   If CF<>0 Then Q.Fields[4].AsFloat:=CF ;
   If LF<>0 Then Q.Fields[5].AsFloat:=LF ;
   Q.Post ;
   Q.Next ;
   END ;
Ferme(Q) ; FiniMove ;
END ;

Procedure TFcalcFrcs.RecalculAnalEuro(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
Var St : String ;
    DF,CF,TF,DE,CE,TE : Double ;
    Q : TQuery ;
BEGIN
(*
St:='select * from ANALYTIQ Where Y_JOURNAL<>"###" ' ;
If FJal1.Text<>'' Then St:=St+' AND Y_JOURNAL>="'+FJal1.Text+'" ' ;
If FJal2.Text<>'' Then St:=St+' AND Y_JOURNAL<="'+FJal2.Text+'" ' ;
If FExercice.ItemIndex>0 Then St:=ST+' AND Y_EXERCICE="'+FExercice.Value+'" ' ;
St:=St+'And Y_DATECOMPTABLE>="'+usdate(FDateCpta1)+'" And Y_DATECOMPTABLE<="'+usdate(FDateCpta2)+'" ' ;
If FEtab.ItemIndex>0 Then St:=ST+' AND Y_ETABLISSEMENT="'+FETAB.Value+'" ' ;
St:=St+' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE ' ;
Q1.SQL.Add(St) ; ChangeSQL(Q1) ; Q1.Open ;
*)
St:='SELECT Y_DEBIT, Y_CREDIT, Y_TOTALECRITURE, Y_DEBITEURO, Y_CREDITEURO, Y_TOTALEURO, ' ;
St:=St+'Y_NUMEROPIECE,Y_NUMLIGNE,Y_SECTION FROM ANALYTIQ WHERE Y_JOURNAL<>"###" ' ;
If Jal1<>'' Then St:=St+' AND Y_JOURNAL>="'+Jal1+'" ' ;
If Jal2<>'' Then St:=St+' AND Y_JOURNAL<="'+Jal2+'" ' ;
If Exo<>'' Then St:=ST+' AND Y_EXERCICE="'+Exo+'" ' ;
If Date1>0 Then St:=St+'AND Y_DATECOMPTABLE>="'+usdateTime(Date1)+'" ' ;
If Date2>0 Then St:=St+'AND Y_DATECOMPTABLE<="'+usdateTime(Date2)+'" ' ;
If Etab<>'' Then St:=ST+' AND Y_ETABLISSEMENT="'+Etab+'" ' ;
St:=St+'AND (Y_EXERCICE="'+VH^.EnCours.code+'" OR Y_EXERCICE="'+VH^.Suivant.code+'") ' ;
St:=St+' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL, Y_QUALIFPIECE ' ;
InitMove(10000,'') ;
Q:=OpenSQL(St,FALSE) ; Q.UpdateMode:=UpWhereChanged ;
FetchSqlODBC(Q) ;
While Not Q.EOF do
   BEGIN
   If MoveCur(FALSE) Then ;
   DE:=Q.Fields[0].AsFloat ; CE:=Q.Fields[1].AsFloat ; TE:=Q.Fields[2].AsFloat ;
   DF:=EuroToPivot(DE) ; CF:=EuroToPivot(CE) ; TF:=EuroToPivot(TE) ;
   Q.Edit ;
   If DF<>0 Then Q.Fields[3].AsFloat:=DF ;
   If CF<>0 Then Q.Fields[4].AsFloat:=CF ;
   If TF<>0 Then Q.Fields[5].AsFloat:=TF ;
   Q.Post ;
   Q.Next ;
   END ;
Ferme(Q) ; FiniMove ;
END ;

Procedure TFcalcFrcs.EquilibreAnalEuro(SJal,SExo,SNum,SLig,SQual : String ; DDate : TDateTime  ; TotalFrancsG : Double) ;
Var St : String ;
    DF,CF : Double ;
    Cle,OldCle : String ;
    PremFois,OkRetouche,OkOk : Boolean ;
    TotalFrancsA,SoldeF : Double ;
    Q1 : TQuery ;
BEGIN
OkOk:=FALSE ;
St:='SELECT Y_DEBIT, Y_CREDIT, Y_TOTALECRITURE, Y_DEBITEURO, Y_CREDITEURO, Y_TOTALEURO, Y_AXE FROM ANALYTIQ WHERE ' ;
St:=St+'Y_JOURNAL="'+SJAl+'" AND Y_EXERCICE="'+SExo+'" AND Y_DATECOMPTABLE="'+UsDateTime(DDate)+'" AND ' ;
St:=St+'Y_NUMEROPIECE='+SNum+' AND Y_NUMLIGNE='+SLig+' AND Y_QUALIFPIECE="'+SQual+'" ' ;
St:=St+' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE ' ;
Q1:=OpenSQL(St,FALSE) ; Q1.UpdateMode:=UpWhereChanged ;
FetchSqlODBC(Q1) ;
TotalFrancsA:=0 ; PremFois:=TRUE ; SoldeF:=0 ;
While Not Q1.EOF do
   BEGIN
//   Cle:=Q1.FindField('Y_AXE').AsString ;
   Cle:=Q1.Fields[6].AsString ;
   If PremFois Then OldCle:=Cle ;
   PremFois:=FALSE ; OkOk:=TRUE ;
   If (Cle<>OldCle) Then
      BEGIN
      SoldeF:=Arrondi(TotalFrancsG-TotalFrancsA,V_PGI.OkDecE) ; OkRetouche:=FALSE ;
      If SoldeF<>0 Then
         BEGIN
         Q1.Prior ;
//         DE:=Q1.FindField('Y_DEBITEURO').AsFloat ; CE:=Q1.FindField('Y_CREDITEURO').AsFloat ;
         DF:=Q1.Fields[3].AsFloat ; CF:=Q1.Fields[4].AsFloat ;
         If Arrondi(DF,V_PGI.OkDecE)<>0 Then
            BEGIN
            OkRetouche:=TRUE ;DF:=DF+SoldeF ;
            END Else If Arrondi(CF,V_PGI.OkDecE)<>0 Then
            BEGIN
            CF:=CF+SoldeF ; OkRetouche:=TRUE ;
            END ;
         If OkRetouche Then
            BEGIN
            Q1.Edit ;
            Q1.Fields[3].AsFloat:=DF ; Q1.Fields[4].AsFloat:=CF ;
            Q1.Post ;
            END ;
         Q1.Next ;
         END ;
      TotalFrancsA:=0 ;
      END ;
//   DE:=Q1.FindField('Y_DEBITEURO').AsFloat ; CE:=Q1.FindField('Y_CREDITEURO').AsFloat ;
   DF:=Q1.Fields[3].AsFloat ; CF:=Q1.Fields[4].AsFloat ;
   TotalFrancsA:=Arrondi(TotalFrancsA+DF+CF,V_PGI.OkDecE) ;
   Q1.Edit ;
//   Q1.FindField('Y_TOTALEURO').AsFloat:=TotalEuroG ;
   Q1.Fields[5].AsFloat:=TotalFrancsG ;
   Q1.Post ;
   PremFois:=FALSE ;
   OldCle:=Cle ;
   Q1.Next ;
   END ;
SoldeF:=Arrondi(TotalFrancsG-TotalFrancsA,V_PGI.OkDecE) ; OkRetouche:=FALSE ;
If (SoldeF<>0) And OkOk Then
   BEGIN
   Q1.Prior ;
   If Not Q1.Eof Then
      BEGIN
//      DE:=Q1.FindField('Y_DEBITEURO').AsFloat ; CE:=Q1.FindField('Y_CREDITEURO').AsFloat ;
      DF:=Q1.Fields[3].AsFloat ; CF:=Q1.Fields[4].AsFloat ;
      If Arrondi(DF,V_PGI.OkDecE)<>0 Then
         BEGIN
         OkRetouche:=TRUE ;DF:=DF+SoldeF ;
         END Else If Arrondi(CF,V_PGI.OkDecE)<>0 Then
         BEGIN
         CF:=CF+SoldeF ; OkRetouche:=TRUE ;
         END ;
      If OkRetouche Then
         BEGIN
         Q1.Edit ;
//         Q1.FindField('Y_DEBITEURO').AsFloat:=DE ; Q1.FindField('Y_CREDITEURO').AsFloat:=CE ;
         Q1.Fields[3].AsFloat:=DF ; Q1.Fields[4].AsFloat:=CF ;
         Q1.Post ;
         END ;
      END ;
   END ;
Ferme(Q1) ;
END ;

Procedure TFcalcFrcs.BalayeEcrPourEquilibreAnal(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
Var St : String ;
    DF,CF,TotalFrancsG : Double ;
    OkAnal : Boolean ;
    SJal,SExo,SNum,SLig,SQual : String ;
    DDate : TDateTime  ;
    Q : TQuery ;
BEGIN
(*
St:='select * from ecriture Where E_JOURNAL<>"###" ' ;
If FJal1.Text<>'' Then St:=St+' AND E_JOURNAL>="'+FJal1.Text+'" ' ;
If FJal2.Text<>'' Then St:=St+' AND E_JOURNAL<="'+FJal2.Text+'" ' ;
If FExercice.ItemIndex>0 Then St:=ST+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
St:=St+'And E_DATECOMPTABLE>="'+usdate(FDateCpta1)+'" And E_DATECOMPTABLE<="'+usdate(FDateCpta2)+'" ' ;
If FEtab.ItemIndex>0 Then St:=ST+' AND E_ETABLISSEMENT="'+FETAB.Value+'" ' ;
St:=St+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ' ;
Q.SQL.Add(St) ; ChangeSQL(Q) ; Q.Open ;
*)
St:='SELECT E_DEBIT, E_CREDIT, E_DEBITEURO, E_CREDITEURO, ' ;
{         4          5              6              7               8            9           10        11   }
St:=St+'E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_ANA, E_GENERAL ' ;
St:=St+'FROM ECRITURE WHERE E_JOURNAL<>"###" ' ;
If Jal1<>'' Then St:=St+' AND E_JOURNAL>="'+Jal1+'" ' ;
If Jal2<>'' Then St:=St+' AND E_JOURNAL<="'+Jal2+'" ' ;
If Exo<>'' Then St:=ST+' AND E_EXERCICE="'+Exo+'" ' ;
If Date1>0 Then St:=St+'AND E_DATECOMPTABLE>="'+usdateTime(Date1)+'" ' ;
If Date2>0 Then St:=St+'AND E_DATECOMPTABLE<="'+usdateTime(Date2)+'" ' ;
If Etab<>'' Then St:=ST+' AND E_ETABLISSEMENT="'+Etab+'" ' ;
St:=St+'AND (E_EXERCICE="'+VH^.EnCours.code+'" OR E_EXERCICE="'+VH^.Suivant.code+'") ' ;
St:=St+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ' ;
InitMove(10000,'') ;
Q:=OpenSQL(St,FALSE) ; Q.UpdateMode:=UpWhereChanged ;
FetchSqlODBC(Q) ;
While Not Q.EOF do
   BEGIN
   If MoveCur(FALSE) Then ;
   (*
   DE:=Q.FindField('E_DEBITEURO').AsFloat ; CE:=Q.FindField('E_CREDITEURO').AsFloat ;
   OkAnal:=Q.FindField('E_ANA').AsString='X' ;
   *)
   DF:=Q.Fields[2].AsFloat ; CF:=Q.Fields[3].AsFloat ;
   OkAnal:=Q.Fields[11].AsString='X' ;
   If OkAnal Then
      BEGIN
      (*
      SJal:=Q.FindField('E_JOURNAL').AsString ;
      SExo:=Q.FindField('E_EXERCICE').AsString ;
      SNum:=IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger) ;
      SLig:=IntToStr(Q.FindField('E_NUMLIGNE').AsInteger) ;
      SQual:=Q.FindField('E_QUALIFPIECE').AsString ;
      DDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
      *)
      SJal:=Q.Fields[4].AsString ;
      SExo:=Q.Fields[5].AsString ;
      DDate:=Q.Fields[6].AsDateTime ;
      SNum:=IntToStr(Q.Fields[7].AsInteger) ;
      SLig:=IntToStr(Q.Fields[8].AsInteger) ;
      SQual:=Q.Fields[10].AsString ;
      TotalFrancsG:=DF+CF ;
      EquilibreAnalEuro(SJal,SExo,SNum,SLig,SQual,DDate,TotalFrancsG) ;
      END ;
   Q.Next ;
   END ;
Ferme(Q) ; FiniMove ;
END ;

Function OkTraitementResultat(ECRANOUVEAU,Jal,Exo,Qual : String ; DateC : TDateTime ; NumC : Integer ; SoldeF : Double) : Boolean ;
Var CptB,CptP,St : String ;
    Q : TQuery ;
    OkOk,OkRetouche : Boolean ;
    D,C,DF,CF : Double ;
BEGIN
Result:=FALSE ;
If EcrANouveau='N' Then Exit ;
CptB:='' ; CptP:='' ;
{$IFDEF SPEC302}
Q:=OpenSQL('SELECT SO_FERMEPERTE, SO_OUVREPERTE, SO_FERMEBEN, SO_OUVREBEN FROM SOCIETE',TRUE) ;
If Not Q.Eof Then
   BEGIN
   If EcrANouveau='C' Then
      BEGIN
      CptP:=Q.Fields[0].AsString ; CptB:=Q.Fields[2].AsString ;
      END Else
      BEGIN
      CptP:=Q.Fields[1].AsString ; CptB:=Q.Fields[3].AsString ;
      END ;
   END ;
Ferme(Q) ;
{$ELSE}
If EcrANouveau='C' Then
   BEGIN
   CptP:=GetParamsocSecur('SO_FERMEPERTE','') ; CptB:=GetParamsocSecur('SO_FERMEBEN','') ;
   END Else
   BEGIN
   CptP:=GetParamsocSecur('SO_OUVREPERTE','') ; CptB:=GetParamsocSecur('SO_OUVREBEN','') ;
   END ;
{$ENDIF}
If (CptP='') Or (CptB='') Then Exit ;
OkOk:=FALSE ;
St:='SELECT E_DEBIT, E_CREDIT, E_DEBITEURO, E_CREDITEURO, ' ;
{         4          5              6              7               8            9           10         11         12 }
St:=St+'E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_GENERAL, E_ECRANOUVEAU ' ;
St:=St+'FROM ECRITURE WHERE E_JOURNAL="'+Jal+'"  AND E_EXERCICE="'+Exo+'" AND ' ;
St:=St+'E_DATECOMPTABLE="'+usdateTime(DateC)+'" AND ' ;
St:=St+'E_NUMEROPIECE='+IntToStr(NumC)+' AND ' ;
St:=St+'E_QUALIFPIECE="'+Qual+'" AND ' ;
St:=St+'(E_GENERAL="'+CptB+'" OR E_GENERAL="'+CptP+'") ' ;
St:=St+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ' ;
Q:=OpenSQL(St,FALSE) ; Q.UpdateMode:=UpWhereChanged ;
FetchSqlODBC(Q) ;
If Not Q.Eof Then
   BEGIN
   OkRetouche:=FALSE ;
   D:=Q.Fields[0].AsFloat ; C:=Q.Fields[1].AsFloat ;
   DF:=Q.Fields[2].AsFloat ; CF:=Q.Fields[3].AsFloat ;
   If Arrondi(D,V_PGI.OkDecV)<>0 Then
      BEGIN
      OkRetouche:=TRUE ;DF:=DF-SoldeF ;
      END Else If Arrondi(C,V_PGI.OkDecV)<>0 Then
      BEGIN
      CF:=CF+SoldeF ; OkRetouche:=TRUE ;
      END ;
   If OkRetouche Then
      BEGIN
      (*
      Q.Edit ; Q.FindField('E_DEBITEURO').AsFloat:=DE ; Q.FindField('E_CREDITEURO').AsFloat:=CE ; Q.Post ;
      *)
      Q.Edit ; Q.Fields[2].AsFloat:=DF ; Q.Fields[3].AsFloat:=CF ; Q.Post ;
      OkOk:=TRUE ;
      END ;
   END ;
Ferme(Q) ;
Result:=OkOk ;
END ;

Procedure TFcalcFrcs.EquilibreEuro(Jal1,Jal2,Exo,Etab : String ; Date1,Date2 : TDateTime) ;
Var St : String ;
    D,C,DF,CF : Double ;
    Cle,OldCle : String ;
    PremFois,OkRetouche : Boolean ;
    TotalPivotD,TotalFrancsD,TotalPivotC,TotalFrancsC,SoldeF : Double ;
    Q : TQuery ;
BEGIN
(*
St:='select * from ecriture Where E_JOURNAL<>"###" ' ;
If FJal1.Text<>'' Then St:=St+' AND E_JOURNAL>="'+FJal1.Text+'" ' ;
If FJal2.Text<>'' Then St:=St+' AND E_JOURNAL<="'+FJal2.Text+'" ' ;
If FExercice.ItemIndex>0 Then St:=ST+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
St:=St+'And E_DATECOMPTABLE>="'+usdate(FDateCpta1)+'" And E_DATECOMPTABLE<="'+usdate(FDateCpta2)+'" ' ;
If FEtab.ItemIndex>0 Then St:=ST+' AND E_ETABLISSEMENT="'+FETAB.Value+'" ' ;
St:=St+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ' ;
Q.SQL.Add(St) ; ChangeSQL(Q) ; Q.Open ;
*)
St:='SELECT E_DEBIT, E_CREDIT, E_DEBITEURO, E_CREDITEURO, ' ;
{         4          5              6              7               8            9           10         11         12 }
St:=St+'E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_GENERAL, E_ECRANOUVEAU ' ;
St:=St+'FROM ECRITURE WHERE E_JOURNAL<>"###" ' ;
If Jal1<>'' Then St:=St+' AND E_JOURNAL>="'+Jal1+'" ' ;
If Jal2<>'' Then St:=St+' AND E_JOURNAL<="'+Jal2+'" ' ;
If Exo<>'' Then St:=ST+' AND E_EXERCICE="'+Exo+'" ' ;
If Date1>0 Then St:=St+' AND E_DATECOMPTABLE>="'+usdateTime(Date1)+'" ' ;
If Date2>0 Then St:=St+' AND E_DATECOMPTABLE<="'+usdateTime(Date2)+'" ' ;
If Etab<>'' Then St:=ST+' AND E_ETABLISSEMENT="'+Etab+'" ' ;
St:=St+'AND (E_EXERCICE="'+VH^.EnCours.code+'" OR E_EXERCICE="'+VH^.Suivant.code+'") ' ;
St:=St+' AND E_MODESAISIE<>"BOR" ' ;
St:=St+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ' ;
InitMove(10000,'') ;
Q:=OpenSQL(St,FALSE) ; Q.UpdateMode:=UpWhereChanged ;
FetchSqlODBC(Q) ;
PremFois:=TRUE ; TotalPivotD:=0 ; TotalFrancsD:=0 ; TotalPivotC:=0 ; TotalFrancsC:=0 ;
While Not Q.EOF do
   BEGIN
   If MoveCur(FALSE) Then ;
   (*
   Cle:=Format_String(Q.FindField('E_JOURNAL').AsString,3)+
        Format_String(Q.FindField('E_EXERCICE').AsString,3)+
        DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)+
        FormatFloat('0000000000',Q.FindField('E_NUMEROPIECE').AsInteger)+
        Format_String(Q.FindField('E_QUALIFPIECE').AsString,3) ;
   *)
   Cle:=Format_String(Q.Fields[4].AsString,3)+
        Format_String(Q.Fields[5].AsString,3)+
        DateToStr(Q.Fields[6].AsDateTime)+
        FormatFloat('0000000000',Q.Fields[7].AsInteger)+
        Format_String(Q.Fields[10].AsString,3) ;
   If PremFois Then OldCle:=Cle ;
   If (Cle<>OldCle) Then
      BEGIN
      SoldeF:=Arrondi(TotalFrancsD-TotalFrancsC,V_PGI.OkDecE) ; OkRetouche:=FALSE ;
      If SoldeF<>0 Then
         BEGIN
         Q.Prior ;
         (*
         D:=Q.FindField('E_DEBIT').AsFloat ; C:=Q.FindField('E_CREDIT').AsFloat ;
         DE:=Q.FindField('E_DEBITEURO').AsFloat ; CE:=Q.FindField('E_CREDITEURO').AsFloat ;
         *)
         D:=Q.Fields[0].AsFloat ; C:=Q.Fields[1].AsFloat ;
         DF:=Q.Fields[2].AsFloat ; CF:=Q.Fields[3].AsFloat ;
         If Arrondi(D,V_PGI.OkDecV)<>0 Then
            BEGIN
            OkRetouche:=TRUE ;DF:=DF-SoldeF ;
            END Else If Arrondi(C,V_PGI.OkDecV)<>0 Then
            BEGIN
            CF:=CF+SoldeF ; OkRetouche:=TRUE ;
            END ;
         If OkRetouche Then
            BEGIN
            (*
            Q.Edit ; Q.FindField('E_DEBITEURO').AsFloat:=DE ; Q.FindField('E_CREDITEURO').AsFloat:=CE ; Q.Post ;
            *)
            If Not OkTraitementResultat(Q.Fields[12].AsString,Q.Fields[4].AsString,Q.Fields[5].AsString,
                                        Q.Fields[10].AsString,Q.Fields[6].AsDateTime,Q.Fields[7].AsInteger,
                                        SoldeF) Then
               BEGIN
               Q.Edit ; Q.Fields[2].AsFloat:=DF ; Q.Fields[3].AsFloat:=CF ; Q.Post ;
               END ;
            END ;
         Q.Next ;
         END ;
      TotalPivotD:=0 ; TotalFrancsD:=0 ; TotalPivotC:=0 ; TotalFrancsC:=0 ;
      END ;
   DF:=Q.Fields[2].AsFloat ; CF:=Q.Fields[3].AsFloat ;
   TotalFrancsD:=Arrondi(TotalFrancsD+DF,V_PGI.OkDecE) ;
   TotalFrancsC:=Arrondi(TotalFrancsC+CF,V_PGI.OkDecE) ;
   PremFois:=FALSE ;
   OldCle:=Cle ;
   Q.Next ;
   END ;
Ferme(Q) ; FiniMove ;
END ;

procedure TFcalcFrcs.BValiderClick(Sender: TObject);
Var okok : Boolean ;
    StEtab,StDev : String ;
    D1,D2 : TDateTime ;
    Q : TQuery ;
    i : Integer ;
begin
TTravail.Visible:=TRUE ; OkOk:=TRUE ;
EnableControls(Self,False) ;
TTravail.Enabled:=TRUE ;
Application.ProcessMessages ;
StEtab:='' ; If FEtab.ItemIndex>0 Then StEtab:=FETAB.Value ;
D1:=StrToDate(FDateCpta1.Text) ; D2:=StrToDate(FDateCpta2.Text) ;
If FDelEc.Checked Then DetruitEcart(FJal1.Text,FJal2.Text,FExercice.Value,StEtab,D1,D2) ;
If FCalc.Checked Then
   BEGIN
    try
     BeginTrans ;
     RecalculEuro(FJal1.Text,FJal2.Text,FExercice.Value,StEtab,D1,D2) ;
     CommitTrans ;
     Except
     Rollback ; OkOk:=FALSE ;
    end ;
   If OkOk Then
      BEGIN
       try
        BeginTrans ;
        RecalculAnalEuro(FJal1.Text,FJal2.Text,FExercice.Value,StEtab,D1,D2) ;
        CommitTrans ;
        Except
        Rollback ; OkOk:=FALSE ;
        end ;
      end ;
   end ;
If OkOk And FEqui.Checked Then
   BEGIN
    try
     BeginTrans ;
     EquilibreEuro(FJal1.Text,FJal2.Text,FExercice.Value,StEtab,D1,D2) ;
     CommitTrans ;
     Except
     Rollback ; OkOk:=FALSE ;
    end ;
   If OkOk Then
      BEGIN
     try
        BeginTrans ;
        BalayeEcrPourEquilibreAnal(FJal1.Text,FJal2.Text,FExercice.Value,StEtab,D1,D2) ;
        CommitTrans ;
        Except
        Rollback ;
        end ;
      end ;
   END ;
EnableControls(Self,True) ;
TTravail.Visible:=FALSE ;
If PourRV Then PbRV:=Not OkOk ;
end;

procedure TFcalcFrcs.FExerciceChange(Sender: TObject);
begin
ExoToDates(FExercice.Value,FDateCpta1,FDateCpta2) ;
end;

procedure TFcalcFrcs.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then
   BEGIN
   DeblocageMonoPoste(True) ;
   Action:=caFree ;
   END ;
end;

procedure TFcalcFrcs.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFcalcFrcs.Timer1Timer(Sender: TObject);
begin
Timer1.Enabled:=FALSE ;
BValiderClick(Nil) ;
Close ;
end;

end.
