unit Qrjustbq;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, Ent1, Hent1, CritEdt, UtilEdt, HQry,
  CpteUtil,Choix, HTB97, HPanel, UiUtil;

procedure JustifPointage;

type
  TFJustifPointage = class(TFQR)
    HLabel3: THLabel;
    FRefP1: TEdit;
    HLabel9: THLabel;
    FRefP2: TEdit;
    TFBanque: THLabel;
    FBanque: THValComboBox;
    GEcritures: TRadioGroup;
    FLigneGenPied: TCheckBox;
    FRuptBanque: TCheckBox;
    FLigneBanque: TCheckBox;
    QRLabel1: TQRLabel;
    RRefP1: TQRLabel;
    QRLabel13: TQRLabel;
    RRefP2: TQRLabel;
    TRBanque: TQRLabel;
    RBanque: TQRLabel;
    TDateCompta: TQRLabel;
    TPiece: TQRLabel;
    TRefInterne: TQRLabel;
    TLibelle: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel19: TQRLabel;
    TDebit: TQRLabel;
    TCredit: TQRLabel;
    TSolde: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    HeadBanque: TQRBand;
    BQ_BANQUE: TQRDBText;
    QRLabel22: TQRLabel;
    TNomBanque: TQRLabel;
    HeadGene: TQRBand;
    FDateDebut: TQRLabel;
    FNonPointeDebut: TQRLabel;
    FPointeDebut: TQRLabel;
    FDebitNPDebut: TQRLabel;
    FDebitPDebut: TQRLabel;
    FCreditNPDebut: TQRLabel;
    FCreditPDebut: TQRLabel;
    FSoldeCptDeb: TQRLabel;
    FNbNPDebut: TQRLabel;
    E_GENERAL: TQRDBText;
    QRLabel23: TQRLabel;
    G_LIBELLE: TQRDBText;
    RRib: TQRLabel;
    E_DATECOMPTABLE: TQRDBText;
    E_NUMEROPIECE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    E_LIBELLE: TQRDBText;
    E_DATEECHEANCE: TQRDBText;
    E_MODEPAIE: TQRDBText;
    E_DEBIT: TQRLabel;
    E_CREDIT: TQRLabel;
    FootPointage: TQRBand;
    TCREDIT_REFP: TQRLabel;
    TDEBIT_REFP: TQRLabel;
    TCOUNT_REFP: TQRLabel;
    TTotalPointage: TQRLabel;
    LigneTot1: TQRShape;
    TSOLDE_REFP: TQRLabel;
    TSoldeBque: TQRLabel;
    FSoldeBque: TQRLabel;
    FootGene: TQRBand;
    TTOTCRE_P: TQRLabel;
    TTOTDEB_P: TQRLabel;
    TTOT_NP: TQRLabel;
    TCREDIT_P: TQRLabel;
    TDEBIT_P: TQRLabel;
    TDEBIT_NP: TQRLabel;
    TCOUNT_P: TQRLabel;
    TCOUNT_NP: TQRLabel;
    FDateFin: TQRLabel;
    FSoldeCptFin: TQRLabel;
    LigneTot2: TQRShape;
    FPointeFin: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    QRDBText1: TQRDBText;
    QRLabel27: TQRLabel;
    QRLabel31: TQRLabel;
    TTotalSelect: TQRLabel;
    TCREDIT_NP: TQRLabel;
    TopGene: TQRShape;
    REPORT2CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    QRBand1: TQRBand;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait0: TQRLigne;
    MsgBox: THMsgBox;
    GBanque: TQRGroup;
    GGene: TQRGroup;
    GPointage: TQRGroup;
    HeadPointage: TQRBand;
    QRLabel8: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel16: TQRLabel;
    CCToutes: TQRLabel;
    CCPointe: TQRLabel;
    CCNonPointe: TQRLabel;
    Trait4: TQRLigne;
    Ligne1: TQRLigne;
    TTOTDEB_NP: TQRLabel;
    TTOTCRE_NP: TQRLabel;
    FNbPDebut: TQRLabel;
    TTOT_P: TQRLabel;
    QRLabel10: TQRLabel;
    E_DEVISE: TQRDBText;
    TBQ_DEVISE: TQRLabel;
    procedure FormShow(Sender: TObject);
    procedure GEcrituresClick(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure HeadGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FootGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure HeadBanqueBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FootPointageBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FootPointageAfterPrint(BandPrinted: Boolean);
    procedure FootGeneAfterPrint(BandPrinted: Boolean);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FRefP1DblClick(Sender: TObject);
    procedure FRuptBanqueClick(Sender: TObject);
    procedure HeadPointageBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
  private
    { Déclarations privées }
    QR2E_BANQUE, QR2E_DEVISE,QR2E_DOMICILIATION, QR2E_ETABBQ,QR2E_EXERCICE,QR2E_JOURNAL,
    QR2E_GUICHET, QR2E_NUMEROCOMPTE, QR2E_CLERIB, QR2E_CODEBIC,
    QR2E_REFPOINTAGE, QR2E_MODEPAIE, QR2E_GENERAL, QR2E_REFINTERNE,
    QR2E_LIBELLE, QR2G_LIBELLE                                     : TStringField ;
    QR2E_DATECOMPTABLE, QR2E_DATEECHEANCE, QR2E_DATEPOINTAGE       : TDateTimeField ;
    QR2E_NUMEROPIECE,QR2E_NUMLIGNE                                 : TIntegerField ;
    QR2DEBIT, QR2CREDIT                                            : TFloatField ;
    COUNT_REFP, COUNT_P, COUNT_NP, TOT_NP, TOT_P                   : Integer ;
    DEBIT_REFP, CREDIT_REFP, SOLDE_REFP, DEBIT_P, CREDIT_P,
    DEBIT_NP, CREDIT_NP, TOTCRE_P, TOTDEB_P, TOTCRE_NP, TOTDEB_NP  : Double ;
    StReportGen, StReportRef                                       : string ;
    PrintPied,PrintTotaux                                          : boolean;
//    Function  QuoiMvt : String ;
//    Procedure GLGenZoom(Quoi : String) ;
    procedure InitSoldes(Gene: string);
    function OkSelect(DateP:TDateTime; RefP:String): boolean;
    procedure IncTotaux ;
    Function  QuoiGen : String ;
    Function  QuoiMvt : String ;
    Procedure JustifZoom(Quoi : String) ;
  public
    { Déclarations publiques }
    procedure FinirPrint ; Override ;
    procedure GenereSQL ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure InitDivers ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  end;

implementation

{$R *.DFM}

procedure JustifPointage;
var QR : TFJustifPointage;
    Edition : TEdition ;
    PP : THPanel ;
begin
PP:=FindInsidePanel ;
QR:=TFJustifPointage.Create(Application) ;
Edition.Etat:=etJustBq ;
QR.QRP.QRPrinter.OnSynZoom:=QR.JustifZoom ;
QR.InitType (nbGen,neJbq,msRien,'QRJUSTBQ','',TRUE,FALSE,FALSE,Edition) ;
if PP=Nil then
   BEGIN
   try
    QR.ShowModal ;
    finally
    QR.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(QR,PP) ;
   QR.Show ;
   END ;
end ;

procedure TFJustifPointage.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  JBq.Banque:='' ; if FBanque.ItemIndex<>0 then JBq.Banque:=FBanque.Value ;
  JBq.RefP1:=FRefP1.Text ; JBq.RefP2:=FRefP2.Text ;
  With JBq.FormatPrint Do
    BEGIN
    PrSepCompte[1]:=FLigneBanque.Checked ;
    PrSepCompte[2]:=FLigneCpt.Checked ;
    PrSepCompte[3]:=FLigneGenPied.Checked ;
    END ;
  END ;
END;

Function TFJustifPointage.QuoiGen : String ;
BEGIN
Inherited ;
Result:=QR2E_GENERAL.AsString+'#'+Qr2G_LIBELLE.AsString+'@'+'1' ;
END ;

Function TFJustifPointage.QuoiMvt : String ;
BEGIN
Inherited ;
Result:=QR2E_GENERAL.AsString+' '+Qr2G_LIBELLE.AsString+' '+{Le_Solde.Caption+}
        '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
        PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
       '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
        IntToStr(QR2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TFJustifPointage.JustifZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Inherited ;
LP:=Pos('@',Quoi) ; If LP=0 Then Exit ;
i:=StrToInt(Copy(Quoi,LP+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QRPrinter.FSynShiftDblClick Then i:=6 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFJustifPointage.GenereSQL;
BEGIN
Q.Close; Q.SQL.Clear;
Q.SQL.Add('Select G_LIBELLE, BQ_BANQUE, BQ_DEVISE,');
Q.SQL.Add(' BQ_DOMICILIATION, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEBIC,');
Q.SQL.Add(' E_GENERAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_LIBELLE, E_DATEECHEANCE, E_REFINTERNE,');
// BPY le 08/12/2003 => fiche 13100 : je dirais bienqu'il faut apprendre a coder mais bon ... pas de commentaire ... trop mdr
//
//Q.SQL.Add(' E_REFPOINTAGE, E_DATEPOINTAGE, E_MODEPAIE, E_JOURNAL, E_EXERCICE, E_NUMLIGNE, E_DEVISE,');
//  if CritEdt.DeviseSelect=V_PGI.DevisePivot then
//  begin
//    if CritEdt.Monnaie = 0 Then
//      Q.SQL.Add(' E_DEBIT DEBIT, E_CREDIT CREDIT ');
//  end
//  else
//    Q.SQL.Add(' E_DEBITDEV DEBIT, E_CREDITDEV CREDIT ');
//
// ==========
Q.SQL.Add(' E_REFPOINTAGE, E_DATEPOINTAGE, E_MODEPAIE, E_JOURNAL, E_EXERCICE, E_NUMLIGNE, E_DEVISE');
  if ((CritEdt.DeviseSelect=V_PGI.DevisePivot) or (CritEdt.Monnaie = 0)) then Q.SQL.Add(', E_DEBIT DEBIT, E_CREDIT CREDIT ')
  else Q.SQL.Add(', E_DEBITDEV DEBIT, E_CREDITDEV CREDIT ');
// fin BPY

Q.SQL.Add(' From ECRITURE E, GENERAUX G, BANQUECP B Where ');
Q.SQL.Add(' E_GENERAL=G.G_GENERAL and G_POINTABLE="X" and BQ_GENERAL=G.G_GENERAL AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" '); // 24/10/2006 YMO Multisociétés );
if VH^.ExoV8.Code<>'' then Q.SQL.Add(' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ') ;
if CritEdt.Joker then
   BEGIN
   Q.SQL.Add('  and E_GENERAL like "'+TraduitJoker(CritEdt.Cpt1)+'"  ' )
   END Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' And E_GENERAL>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' And E_GENERAL<="'+CritEdt.Cpt2+'"  ') ;
   END ;
if FBanque.ItemIndex<>0 then Q.SQL.Add(' And BQ_BANQUE="'+CritEdt.JBq.Banque+'"');
If CritEdt.DeviseSelect<>V_PGI.DevisePivot Then
  BEGIN
  Q.SQL.Add(' AND E_DEVISE="'+CritEdt.DeviseSelect+'"');
  END ;
//Q.SQL.Add(' and E_DEVISE="'+CritEdt.DeviseSelect+'"');
Q.SQL.Add(' and BQ_DEVISE="'+CritEdt.DeviseSelect+'"');
Q.SQL.Add(' and E_DATECOMPTABLE>="'+USDATETIME(CritEdt.DateDeb)+'" and E_DATECOMPTABLE<="'+USDATETIME(CritEdt.DateFin)+'"');
Q.SQL.Add(' and E_QUALIFPIECE="N" and E_ECRANOUVEAU<>"OAN" and E_ECRANOUVEAU<>"C"');
Q.SQL.Add(' Order by');
if FRuptBanque.Checked then Q.SQL.Add(' BQ_BANQUE,');
Q.SQL.Add(' E_GENERAL, E_DATEPOINTAGE, E_DATECOMPTABLE, E_NUMEROPIECE');
ChangeSQL(Q) ; Q.Open;
END;

procedure TFJustifPointage.RenseigneCritere;
BEGIN
Inherited ;
RRefP1.Caption:=FRefP1.Text; RRefP2.Caption:=FRefP2.Text;
Case GEcritures.ItemIndex of
  0 : BEGIN CCToutes.Caption:='þ' ; CCPointe.Caption:='o' ; CCNonPointe.Caption:='o' ; END ;
  1 : BEGIN CCToutes.Caption:='o' ; CCPointe.Caption:='þ' ; CCNonPointe.Caption:='o' ; END ;
  2 : BEGIN CCToutes.Caption:='o' ; CCPointe.Caption:='o' ; CCNonPointe.Caption:='þ' ; END ;
end ;
RBanque.Caption:=FBanque.Text ;
END;

procedure TFJustifPointage.InitDivers;
BEGIN
{Labels sur les bandes}
Inherited ;
FDateDebut.Caption:=MsgBox.Mess[5]+DateToStr(CritEdt.DateDeb-1);
FDateFin.Caption:=MsgBox.Mess[5]+DateToStr(CritEdt.DateFin);

{Séparateurs de bandes}
HeadBanque.Frame.DrawBottom:=CritEdt.JBq.FormatPrint.PrSepCompte[1] ;
HeadBanque.Frame.DrawTop:=CritEdt.JBq.FormatPrint.PrSepCompte[1] ;
HeadGene.Frame.DrawBottom:=CritEdt.JBq.FormatPrint.PrSepCompte[2] ;
FootGene.Frame.DrawBottom:=CritEdt.JBq.FormatPrint.PrSepCompte[3] ;
TopGene.Visible:=CritEdt.JBq.FormatPrint.PrSepCompte[3] ;
COUNT_REFP:=0 ; COUNT_P:=0 ; COUNT_NP:=0 ; TOT_NP:=0 ; TOT_P:=0 ;
DEBIT_REFP:=0 ; CREDIT_REFP:=0 ; SOLDE_REFP:=0 ;
DEBIT_P:=0 ; CREDIT_P:=0 ; DEBIT_NP:=0 ; CREDIT_NP:=0 ;
TOTDEB_P:=0 ; TOTCRE_P:=0 ; TOTDEB_NP:=0 ; TOTCRE_NP:=0 ;
PrintPied:=false; PrintTotaux:=false;
END;

procedure TFJustifPointage.ChoixEdition ;
BEGIN
Inherited ;
ChgMaskChamp(QR2DEBIT , CritEdt.Decimale, CritEdt.AfficheSymbole, CritEdt.Symbole, False) ;
ChgMaskChamp(QR2CREDIT, CritEdt.Decimale, CritEdt.AfficheSymbole, CritEdt.Symbole, False) ;
END ;

function  TFJustifPointage.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
END ;

Procedure TFJustifPointage.FinirPrint ;
BEGIN
InHerited ;
// ff
END ;

procedure TFJustifPointage.InitSoldes(Gene: string);
var i            : integer;
    D,C          : double;
    StSQL,NNP, NP : String;
    // Rony 17/07/97 StSQL,N : String;
    Q2            : TQuery;
BEGIN
For i:=1 to 2 do
    BEGIN
      if CritEdt.Monnaie = 0 then
        StSQL := 'Select SUM(E_CREDIT) CRED, SUM(E_DEBIT) DEB'
      else
        StSQL := 'Select SUM(E_CREDITDEV) CRED, SUM(E_DEBITDEV) DEB';

    StSQL:=StSQL+', COUNT(E_GENERAL) NB';
    // Fin Rony
    StSQL:=StSQL+' From ECRITURE';
    StSQL:=StSQL+' Where E_GENERAL="'+Gene+'" ' ;
    If CritEdt.Exo.Code<>'' Then StSQL:=StSQL+' and E_EXERCICE<="'+CritEdt.Exo.Code+'" ';
// rony 28/05/97    StSQL:=StSQL+' and E_DATECOMPTABLE>="'+USDATETIME(CritEdt.DateDeb)+'"';
    StSQL:=StSQL+' and E_DATECOMPTABLE<"'+USDATETIME(CritEdt.DateDeb)+'"';
    StSQL:=StSQL+' and E_QUALIFPIECE="N" and E_ECRANOUVEAU<>"OAN" and E_ECRANOUVEAU<>"C"';
    if i=1 then StSQL:=StSQL+' and E_REFPOINTAGE=""' else StSQL:=StSQL+' and E_REFPOINTAGE<>""';
    if VH^.ExoV8.Code<>'' then StSQL:=StSQL+' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ' ;
    If CritEdt.DeviseSelect<>V_PGI.DevisePivot Then
      BEGIN
      StSQL:=StSQL+' AND E_DEVISE="'+CritEdt.DeviseSelect+'"' ;
      END ;
    StSQL:=StSQL+' Group by E_GENERAL';
    Q2:=OpenSQL(StSQL,true);
    D:=Q2.FindField('Deb').AsFloat; C:=Q2.FindField('Cred').AsFloat;
    // Rony 17/07/97 if i=1 then BEGIN N:=Q2.FindField('Nb').AsString; if N='' then N:='0'; END;
    if i=1 then BEGIN NNP:=Q2.FindField('Nb').AsString; if NNP='' then NNP:='0'; END
           ELse BEGIN NP:=Q2.FindField('Nb').AsString; if NP='' then NP:='0'; END ;
    // Fin rony
    if i=1 then
       BEGIN
       FNbNPDebut.Caption:=NNP ; //FNbNPDebut.Caption:=N ;
       FCreditNPDebut.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
       FDebitNPDebut.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
       TOT_NP:= StrToInt(NNP) ; {TOT_NP:= StrToInt(N)} ; TOTCRE_NP:=C ; TOTDEB_NP:=D ;
       END else
       BEGIN
       FNbPDebut.Caption:=NP ;  //rony 17/07/97
       FCreditPDebut.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
       FDebitPDebut.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
       TOTCRE_P:=C ; TOTDEB_P:=D ;
       TOT_P:= StrToInt(NP) ; //rony 17/07/97
       END;
    Ferme(Q2);
    END;
SOLDE_REFP:=TOTDEB_P-TOTCRE_P ;
FSoldeCptDeb.Caption:=PrintSolde(TOTDEB_P+TOTDEB_NP,TOTCRE_P+TOTCRE_NP,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole);
END;

function TFJustifPointage.OkSelect(DateP:TDateTime; RefP:String): boolean;
BEGIN
Result:=FALSE ;
case GEcritures.ItemIndex of
   0 : BEGIN // ecr pointées et non pointées
       Result:=(RefP='');
       if Result then exit;
       if FRefP1.Text<>'' then Result:=(RefP>=FRefP1.Text) else Result:=true;
       if Result then BEGIN if FRefP2.Text<>'' then Result:=(RefP<=FRefP2.Text) else Result:=true;END;
       END;
   1 : BEGIN // ecr pointées
       Result:=(RefP<>'');
       if Result then if FRefP1.Text<>'' then Result:=(RefP>=FRefP1.Text);
       if Result then if FRefP2.Text<>'' then Result:=(RefP<=FRefP2.Text);
       END;
   2 : BEGIN //ecr non pointées
       Result:=(RefP='');
       END;
 end;
END;

procedure TFJustifPointage.IncTotaux ;
BEGIN
if QR2E_REFPOINTAGE.AsString<>''then
   BEGIN
   inc(TOT_P) ; //rony 17/07/97
   TOTDEB_P:=TOTDEB_P+QR2DEBIT.AsFloat ; TOTCRE_P:=TOTCRE_P+QR2CREDIT.AsFloat ;
   if OkSelect(QR2E_DATEPOINTAGE.AsDateTime,QR2E_REFPOINTAGE.AsString) then
      BEGIN
      inc(COUNT_REFP) ;
      DEBIT_REFP:=DEBIT_REFP+QR2DEBIT.AsFloat ; CREDIT_REFP:=CREDIT_REFP+QR2CREDIT.AsFloat ;
      SOLDE_REFP:=SOLDE_REFP+(QR2DEBIT.AsFloat-QR2CREDIT.AsFloat) ;
      inc(COUNT_P) ;
      DEBIT_P:=DEBIT_P+QR2DEBIT.AsFloat ; CREDIT_P:=CREDIT_P+QR2CREDIT.AsFloat ;
      END;
   END else
   BEGIN
   inc(TOT_NP) ; TOTDEB_NP:=TOTDEB_NP+QR2DEBIT.AsFloat ; TOTCRE_NP:=TOTCRE_NP+QR2CREDIT.AsFloat ;
   if OkSelect(QR2E_DATEPOINTAGE.AsDateTime,QR2E_REFPOINTAGE.AsString) then
      BEGIN
      inc(COUNT_REFP) ;
      DEBIT_REFP:=DEBIT_REFP+QR2DEBIT.AsFloat ; CREDIT_REFP:=CREDIT_REFP+QR2CREDIT.AsFloat ;
      //SOLDE_REFP:=SOLDE_REFP+(QR2DEBIT.AsFloat-QR2CREDIT.AsFloat) ;
      inc(COUNT_NP) ;
      DEBIT_NP:=DEBIT_NP+QR2DEBIT.AsFloat ;
      CREDIT_NP:=CREDIT_NP+QR2CREDIT.AsFloat ;
      END;
   END;
END ;

procedure TFJustifPointage.FormShow(Sender: TObject);
begin
HelpContext:=7622000 ;
//Standards.HelpContext:=7622001 ;
//Avances.HelpContext:=7622005 ;
//Mise.HelpContext:=7622010 ;
//Option.HelpContext:=7622020 ;

GEcritures.ItemIndex:=0; FBanque.ItemIndex:=0;
  inherited;
end;

procedure TFJustifPointage.GEcrituresClick(Sender: TObject);
begin
  inherited;
FRefP1.enabled:=(GEcritures.ItemIndex<>2); FRefP2.enabled:=(GEcritures.ItemIndex<>2);
end;

procedure TFJustifPointage.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
IncTotaux ;
PrintBand:=OkSelect(QR2E_DATEPOINTAGE.AsDateTime,QR2E_REFPOINTAGE.AsString);
if PrintBand then PrintPied:=true else Exit ;
Quoi:=QuoiMvt ;
if QR2DEBIT.AsFloat=0 then E_DEBIT.Caption:=''
                      else E_DEBIT.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, QR2DEBIT.AsFloat, CritEdt.AfficheSymbole ) ;
if QR2CREDIT.AsFloat=0 then E_CREDIT.Caption:=''
                       else E_CREDIT.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, QR2CREDIT.AsFloat, CritEdt.AfficheSymbole ) ;

AddReport([1,2,3],CritEdt.JBq.FormatPrint.Report,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale) ;
end;

procedure TFJustifPointage.HeadGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var St :string;
begin
  inherited;
Quoi:=QuoiGen ;
InitSoldes(QR2E_GENERAL.AsString);
// écriture du RIB du compte****
St:=QR2E_ETABBQ.AsString+' '+QR2E_GUICHET.AsString+' '+QR2E_NUMEROCOMPTE.AsString+' '+QR2E_CLERIB.AsString ;
if St='' then St:=QR2E_CODEBIC.AsString;
RRib.Caption:=QR2E_DOMICILIATION.AsString+' - '+St;
TBQ_DEVISE.Caption:=MsgBox.Mess[9]+' '+RechDom('TTDEVISETOUTES',QR2E_DEVISE.AsString,false);

StReportGen:= QR2E_GENERAL.AsString ;
InitReport([2],CritEdt.JBq.FormatPrint.Report) ;
end;

procedure TFJustifPointage.FootGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var i : integer;
begin
  inherited;
Quoi:=QuoiGen ;
LigneTot2.Left:=TTotalSelect.Left ;
LigneTot2.Width:=(FootGene.Width-TTotalSelect.Left) ;
TopGene.Left:=0 ;
TopGene.Width:=FootGene.Width ;
TCOUNT_NP.Caption:=IntToStr(COUNT_NP) ; TCOUNT_P.Caption:=IntToStr(COUNT_P) ;
TDEBIT_NP.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, DEBIT_NP, CritEdt.AfficheSymbole ) ;
TCREDIT_NP.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, CREDIT_NP, CritEdt.AfficheSymbole ) ;
TDEBIT_P.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, DEBIT_P, CritEdt.AfficheSymbole ) ;
TCREDIT_P.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, CREDIT_P, CritEdt.AfficheSymbole ) ;
TTOT_NP.Caption:=IntToStr(TOT_NP) ;
TTOT_P.Caption:=IntToStr(TOT_P) ; //rony17/07/97
TTOTDEB_NP.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TOTDEB_NP, CritEdt.AfficheSymbole ) ;
TTOTCRE_NP.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TOTCRE_NP, CritEdt.AfficheSymbole ) ;
TTOTDEB_P.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TOTDEB_P, CritEdt.AfficheSymbole ) ;
TTOTCRE_P.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TOTCRE_P, CritEdt.AfficheSymbole ) ;

FSoldeCptFin.Caption:=PrintSolde(TOTDEB_P+TOTDEB_NP,TOTCRE_P+TOTCRE_NP,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole);
for i:=0 to FootGene.ControlCount-1 do
    if FootGene.Controls[i].Tag=1 then FootGene.Controls[i].visible:=PrintTotaux;
//FootGene.Frame.DrawBottom:=FSepCompte.Checked;
end;

procedure TFJustifPointage.HeadBanqueBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
if Not FRuptBanque.Checked then BEGIN PrintBand:=false; exit; END;
TNomBanque.caption:=RechDom('ttBanque',QR2E_BANQUE.AsString,false);
//HeadBanque.Frame.DrawBottom:=FSepCompte.Checked;
end;

procedure TFJustifPointage.FootPointageBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var OkPointe : boolean;
    Q        : TQuery;
    D, C     : double;
    St       : string;
begin
  inherited;
if Not PrintPied then BEGIN PrintBand:=False ; Exit ; END ;
LigneTot1.Left:=TTotalPointage.Left ;
LigneTot1.Width:=(FootPointage.Width-TTotalPointage.Left) ;
OkPointe:=(QR2E_REFPOINTAGE.AsString<>'');
TSOLDE_REFP.Visible:=OkPointe ; TSoldeBque.visible:=OkPointe ; FSoldeBque.visible:=OkPointe;
TCOUNT_REFP.Caption:=IntToStr(COUNT_REFP) ;
TDEBIT_REFP.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, DEBIT_REFP, CritEdt.AfficheSymbole ) ;
TCREDIT_REFP.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, CREDIT_REFP, CritEdt.AfficheSymbole ) ;
if Not OkPointe then
   BEGIN
   TTotalPointage.Caption:=MsgBox.Mess[3];
   TSOLDE_REFP.Caption:='' ;
   END else
   BEGIN
   TTotalPointage.Caption:=MsgBox.Mess[4]+' '+QR2E_REFPOINTAGE.AsString ;
   TSoldeBque.Caption:=MsgBox.Mess[6]+' '+QR2E_DATEPOINTAGE.AsString ;

   if SOLDE_REFP<0 then TSOLDE_REFP.Caption:=PrintSolde(0,Abs(SOLDE_REFP),CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)
                   else TSOLDE_REFP.Caption:=PrintSolde(Abs(SOLDE_REFP),0,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   // renseignement des soldes bancaires sur pieds de groupes sur ref pointage
   St:='Select EE_NEWSOLDECRE, EE_NEWSOLDEDEB,EE_NEWSOLDECREEURO, EE_NEWSOLDEDEBEURO From EEXBQ ';
   St:=St+' Where EE_GENERAL="'+QR2E_GENERAL.AsString+'"';
   St:=St+' and EE_DATEPOINTAGE="'+USDATETIME(QR2E_DATEPOINTAGE.AsDateTime)+'"';
   St:=St+' and EE_REFPOINTAGE="'+QR2E_REFPOINTAGE.AsString+'"';
   Q:=OpenSQL(St,true);
   if Not Q.EOF then
      BEGIN
        if CritEdt.Monnaie = 0 then
        begin
          C := Q.FindField('EE_NEWSOLDECREEURO').AsFloat ;
          D := Q.FindField('EE_NEWSOLDEDEBEURO').AsFloat ;
        end
        else
        begin
          C := Q.FindField('EE_NEWSOLDECRE').AsFloat ;
          D := Q.FindField('EE_NEWSOLDEDEB').AsFloat ;
        end ;
      END Else BEGIN C:=0 ; D:=0 ; END ;
   Ferme(Q) ;
   FSoldeBque.Caption:=PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   END;
PrintTotaux:=true ;
end;

procedure TFJustifPointage.FootPointageAfterPrint(BandPrinted: Boolean);
begin
  inherited;
PrintPied:=false;
COUNT_REFP:=0 ; DEBIT_REFP:=0 ; CREDIT_REFP:=0 ;
InitReport([3],CritEdt.JBq.FormatPrint.Report) ;
end;

procedure TFJustifPointage.FootGeneAfterPrint(BandPrinted: Boolean);
begin
  inherited;
PrintTotaux:=false;
COUNT_P:=0 ; COUNT_NP:=0 ; DEBIT_P:=0 ; CREDIT_P:=0 ; DEBIT_NP:=0 ; CREDIT_NP:=0 ;
TOTDEB_P:=0 ; TOTCRE_P:=0 ; TOTDEB_NP:=0 ; TOTCRE_NP:=0 ;
InitReport([2],CritEdt.JBq.FormatPrint.Report) ;
end;

procedure TFJustifPointage.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2E_BANQUE        := TStringField(Q.FindField('BQ_BANQUE')) ;
QR2E_DOMICILIATION := TStringField(Q.FindField('BQ_DOMICILIATION')) ;
QR2E_ETABBQ        := TStringField(Q.FindField('BQ_ETABBQ')) ;
QR2E_GUICHET       := TStringField(Q.FindField('BQ_GUICHET')) ;
QR2E_NUMEROCOMPTE  := TStringField(Q.FindField('BQ_NUMEROCOMPTE')) ;
QR2E_CLERIB        := TStringField(Q.FindField('BQ_CLERIB')) ;
QR2E_CODEBIC       := TStringField(Q.FindField('BQ_CODEBIC')) ;
QR2E_DEVISE        := TStringField(Q.FindField('BQ_DEVISE')) ;
QR2E_REFPOINTAGE   :=TStringField(Q.FindField('E_REFPOINTAGE')) ;
QR2E_MODEPAIE      :=TStringField(Q.FindField('E_MODEPAIE')) ;
QR2E_GENERAL       :=TStringField(Q.FindField('E_GENERAL')) ;
QR2E_REFINTERNE    :=TStringField(Q.FindField('E_REFINTERNE')) ;
QR2E_LIBELLE       :=TStringField(Q.FindField('E_LIBELLE')) ;
QR2E_JOURNAL       :=TStringField(Q.FindField('E_JOURNAL')) ;
QR2G_LIBELLE       :=TStringField(Q.FindField('G_LIBELLE')) ;
QR2E_EXERCICE      :=TStringField(Q.FindField('E_EXERCICE')) ;
QR2E_DATECOMPTABLE :=TDateTimeField(Q.FindField('E_DATECOMPTABLE')) ;
QR2E_DATEECHEANCE  :=TDateTimeField(Q.FindField('E_DATEECHEANCE')) ;
QR2E_DATEPOINTAGE  :=TDateTimeField(Q.FindField('E_DATEPOINTAGE')) ;
QR2E_NUMEROPIECE   :=TIntegerField(Q.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE      :=TIntegerField(Q.FindField('E_NUMLIGNE')) ;
QR2DEBIT           :=TFloatField(Q.FindField('DEBIT')) ;
QR2CREDIT          :=TFloatField(Q.FindField('CREDIT')) ;

//ChgMaskChamp(QR2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
//ChgMaskChamp(QR2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFJustifPointage.FRefP1DblClick(Sender: TObject);
var St : string;
begin
  inherited;
St:='';
if FCpte1.Text<>'' then St:='EE_GENERAL>="'+FCpte1.Text+'"';
if FCpte2.Text<>'' then
   BEGIN
   if St='' then St:='EE_GENERAL<="'+FCpte2.Text+'"'
             else St:=St+' and EE_GENERAL<="'+FCpte2.Text+'"';
   END;
St:=Choisir(MsgBox.Mess[1],'EEXBQ','EE_REFPOINTAGE','EE_REFPOINTAGE',St,'EE_REFPOINTAGE');
if St='' then exit;
TEdit(Sender).Text:=St;
end;

procedure TFJustifPointage.FRuptBanqueClick(Sender: TObject);
begin
  inherited;
FLigneBanque.Checked:=false ;
FLigneBanque.Enabled:=(FRuptBanque.Checked);
end;

procedure TFJustifPointage.HeadPointageBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=False ;
StReportRef:= QR2E_REFPOINTAGE.AsString ;
InitReport([3],CritEdt.JBq.FormatPrint.Report) ;
end;

procedure TFJustifPointage.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TitreReportH.Caption:=TitreReportB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
end;

procedure TFJustifPointage.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
    R   : integer ;
begin
  inherited;
R:=QuelReport(CritEdt.JBq.FormatPrint.Report,D,C) ;
Case R Of
  1 : TitreReportB.Caption:='' ;
  2 : TitreReportB.Caption:=MsgBox.Mess[7]+' '+StReportGen ;
  3 : TitreReportB.Caption:=MsgBox.Mess[8]+' '+StReportRef ;
  END ;
if R=1 then
   BEGIN
   Report2Debit.Caption:='' ; Report2Credit.Caption:='' ;
   END else
   BEGIN
   Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
   Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
   END ;
end;

end.
