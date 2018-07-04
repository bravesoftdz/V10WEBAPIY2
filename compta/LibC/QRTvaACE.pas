unit QRTvaAcE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, HMsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons,
  Hctrls, HQry, ExtCtrls, Mask, Hcompte, ComCtrls, Ent1, Hent1,CritEdt,UtilEdt,QRRupt,
  CpteUtil,
//{$IFNDEF ESP(*}
  TVA,
//{$ELSE}
//{$ENDIF ESP} //XMG 19/01/04*)
  SaisComm, HTB97, SaisUtil, HStatus, HPanel, UiUtil, UtilPGI, tCalcCum ;

procedure EditionTvaAce ;

type
  TFQRTvaAcE = class(TFQR)
    QRLabel1: TQRLabel;
    TLE_DATECOMPTABLE: TQRLabel;
    TE_JOURNAL: TQRLabel;
    TE_PIECE: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TE_LIBELLE: TQRLabel;
    TLE_ENCAISSE: TQRLabel;
    MsgBox: THMsgBox;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    TITRE1REP: TQRLabel;
    TITRE2REP: TQRLabel;
    TLE_TVA: TQRLabel;
    REPORT1TVA: TQRLabel;
    TotalTva: TQRLabel;
    REPORT2TVA: TQRLabel;
    QRBand1: TQRBand;
    Trait0: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Ligne1: TQRLigne;
    FTva: THValComboBox;
    HLabel9: THLabel;
    TLE_BASE: TQRLabel;
    QRLigne1: TQRLigne;
    QRLigne2: TQRLigne;
    TotalBase: TQRLabel;
    FSautPage: TCheckBox;
    TypeTVA: TRadioGroup;
    TLE_AUXILIAIRE: TQRLabel;
    E_DATECOMPTABLE: TQRDBText;
    E_JOURNAL: TQRDBText;
    EE_PIECELIGECH: TQRLabel;
    E_REFINTERNE: TQRDBText;
    E_LIBELLE: TQRDBText;
    LE_TVA: TQRLabel;
    E_AUXILIAIRE: TQRDBText;
    LE_PAYE: TQRLabel;
    BFinAux: TQRBand;
    PayeTiers: TQRLabel;
    BaseTiers: TQRLabel;
    TVATiers: TQRLabel;
    QRLabel13: TQRLabel;
    NomTiers: TQRLabel;
    QRDLMulti: TQRDetailLink;
    QRDLRecapG: TQRDetailLink;
    BRecapG: TQRBand;
    Rappel: TQRLabel;
    LibRecapG: TQRLabel;
    QRLabel10: TQRLabel;
    BaseTiersG: TQRLabel;
    TVATiersG: TQRLabel;
    FRecapCpt: TCheckBox;
    FRecapG: TCheckBox;
    PayeTiersG: TQRLabel;
    TLE_TAUX: TQRLabel;
    LE_TAUXTVA: TQRLabel;
    TotalPaye: TQRLabel;
    LE_ECHEENC: TQRLabel;
    HLabel3: THLabel;
    FRegimeTva: THValComboBox;
    FPrevalidTva: TCheckBox;
    FValidTva: TCheckBox;
    QRLabel2: TQRLabel;
    LE_VALIDE: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    TRSituation: TQRLabel;
    RCollectee: TQRLabel;
    QRLabel16: TQRLabel;
    RDeductible: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel19: TQRLabel;
    RTauxTva: TQRLabel;
    QRLabel20: TQRLabel;
    RRegimeTVA: TQRLabel;
    procedure FormShow(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure QRDLRecapGNeedData(var MoreData: Boolean);
    procedure BRecapGAfterPrint(BandPrinted: Boolean);
    procedure BRecapGBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinAuxBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure EntetePageBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BValiderClick(Sender: TObject);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
    procedure InitDivers ; Override ;
    procedure ChoixEdition ; Override ;
  private
    AucunMvt : Boolean ;
    Qr1Debit,Qr1Credit,Qr1Couv,Qr1Base : TFloatField ;
    Qr1Regime,Qr1E_AUXILIAIRE,QR1T_LIBELLE,Qr1E_JOURNAL,Qr1E_EXERCICE,Qr1E_EDITEETATTVA : TStringField ;
    Qr1E_DATECOMPTABLE : TDateTimeField ;
    Qr1E_NUMEROPIECE,Qr1E_NUMLIGNE : TIntegerField ;
    TauxTva : Double ;
    TotEdt                : TabTot ;
    LTiers,LRecapG  : TStringList ;
    Function  QuoiMvt(i : integer) : String ;
    Function  QuoiAux(i : Integer) : String ;
    Procedure TvaFacZoom(Quoi : String) ;
    procedure MajOkEdition ;
    procedure LanceTousLesTaux;
  public
    { Déclarations publiques }
  end;

implementation

uses MsgQrTVA;

{$R *.DFM}

Function TFQRTvaAcE.QuoiAux(i : Integer) : String ;
BEGIN
Case i Of
  1 : Result:=QR1E_AUXILIAIRE.AsString+' '+'#'+QR1T_LIBELLE.AsString+'@'+'2' ;
  END ;
END ;

Function TFQRTvaAcE.QuoiMvt( i : integer) : String ;
var St : String ;
BEGIN
case i of
  1 : St:=QR1E_AUXILIAIRE.AsString+' '+QR1T_LIBELLE.AsString+' '+'#'+
          Qr1E_JOURNAL.AsString+' N° '+IntToStr(Qr1E_NUMEROPIECE.AsInteger)+
          ' '+DateToStr(Qr1E_DateComptable.AsDAteTime)+' . . . '+
          PrintSolde(Qr1DEBIT.AsFloat,Qr1CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)+
          '@'+'5;'+Qr1E_JOURNAL.AsString+';'+UsDateTime(Qr1E_DATECOMPTABLE.AsDateTime)+';'+
          IntToStr(Qr1E_NUMEROPIECE.AsInteger)+';'+Qr1E_EXERCICE.asString+';'+
          IntToStr(Qr1E_NumLigne.AsInteger)+';' ;
  END ;
Result:=St ;
END ;

Procedure TFQRTvaAcE.TvaFacZoom(Quoi : String) ;
Var Lp,i: Integer ;
    St : String ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
St:=Copy(Quoi,Lp+1,1) ;
If St='!' Then i:=100 Else i:=StrToInt(St) ;
If (i=5) Or (i=100) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QrPrinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure EditionTvaAce ;
var QR : TFQRTvaAcE ;
    Edition : TEdition ;
    PP : THPanel ;
    QC : TQuery ;
    OkTva : Boolean ;
BEGIN
If V_PGI.OutLook Then
  BEGIN
  OkTva:=FALSE ;
  QC:=OpenSQL('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="'+VH^.DefCatTVA+'" AND (CC_LIBRE="1" OR CC_LIBRE="2"'+
             ' OR CC_LIBRE="3" OR CC_LIBRE="4") ORDER BY CC_LIBRE',True) ;
  If not QC.Eof Then OkTva:=TRUE ;
  Ferme(QC) ;
  If Not OkTva Then BEGIN HShowMessage('9;Tva sur encaissements;Le paramétrage de la TVA est incorrect !;E;O;O;O;','','') ; Exit ; END ;
  if Not VH^.OuiTvaEnc then BEGIN HShowMessage('10;Tva sur encaissements;Le module n''est pas installé !;E;O;O;O;','','') ; Exit ; END ;
  END ;
PP:=FindInsidePanel ;
QR:= TFQRTvaAcE.Create(Application) ;
Edition.Etat:=etTvaFac ;
QR.QRP.QRPrinter.OnSynZoom:=QR.TvaFacZoom ;
QR.InitType (nbGen,neGL,msGenEcr,'QRTvaAce','',TRUE,FALSE,TRUE,Edition) ;
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
END;

procedure TFQRTvaAcE.GenereSQL ;
Var StValid : String ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('Select E_GENERAL,E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_QUALIFPIECE,E_NUMLIGNE,') ;
Q.SQL.Add(       'E_REFINTERNE,E_LIBELLE,E_DEVISE,E_LETTRAGE,E_JOURNAL,E_NATUREPIECE,E_ETATLETTRAGE,') ;
Q.SQL.Add(       'E_REGIMETVA,E_ECHEENC'+IntToStr(CritEdt.GL.NumTVA)+' ECHEENC,T_LIBELLE LIBAUX,E_EDITEETATTVA,') ;
Case CritEdt.Monnaie of
  0 : BEGIN Q.SQL.Add('E_DEBIT DEBIT, E_CREDIT CREDIT, E_COUVERTURE COUVERTURE') ; END ;
  1 : BEGIN Q.SQL.Add('E_DEBITDEV DEBIT, E_CREDITDEV CREDIT, E_COUVERTUREDEV COUVERTURE') ; END ;
end ;
Q.SQL.Add(' From ECRITURE ') ;
Q.SQL.Add(' Left Join TIERS on E_AUXILIAIRE=T_AUXILIAIRE ') ;
Q.SQL.Add('WHERE E_DATECOMPTABLE>="'+usDateTime(CritEdt.GL.Date21)+'" And E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'"') ;
if CritEdt.Exo.Code<>'' then Q.SQL.Add(' And E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
if CritEdt.GL.RegimeTva<>'' then Q.SQL.Add(' And E_REGIMETVA="'+CritEdt.GL.RegimeTva+'" ') ;
If CritEdt.GL.Deductible Then
   BEGIN
   Q.SQL.Add(SQLCollFouEncTVA('')) ;
   Q.SQL.Add(' AND (E_NATUREPIECE="OF")') ;
   Q.SQL.Add(' AND (T_NATUREAUXI="FOU" Or T_NATUREAUXI="AUC") ') ;
   END Else
   BEGIN
   Q.SQL.Add(SQLCollCliEncTVA('')) ;
   Q.SQL.Add(' AND (E_NATUREPIECE="OC")') ;
   Q.SQL.Add(' AND (T_NATUREAUXI="CLI" Or T_NATUREAUXI="AUD") ') ;
   END ;
//Q.SQL.Add(' AND (E_EDITEETATTVA="-" OR E_EDITEETATTVA="X")') ;
StValid:=' AND (E_EDITEETATTVA="-" ' ;
If CritEdt.GL.PrevalidTVA Then StValid:=StValid+' OR E_EDITEETATTVA="X" ' ;
If CritEdt.GL.ValidTVA Then StValid:=StValid+' OR E_EDITEETATTVA="#" ' ;
StValid:=StValid+') ' ;
Q.SQL.Add(StValid) ;
Q.SQL.Add(' AND E_ETATLETTRAGE="AL"') ;
{ Construction de la clause Order By de la SQL }
Q.SQL.Add(' Order By E_AUXILIAIRE,E_GENERAL,E_ETATLETTRAGE,E_LETTRAGE,E_DATECOMPTABLE') ;

ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
AucunMvt:=(Q.Eof) ;
END ;

procedure TFQRTvaAcE.FormShow(Sender: TObject);
var QC : TQuery ;
    OkTva : Boolean ;
begin
FTva.Values.Clear ; FTva.Items.Clear ; OkTva:=FALSE ;
FTva.Values.Add('') ;                         // 10538
FTva.Items.Add(TraduireMemoire('<<Tous>>')) ; // 10538
QC:=OpenSQL('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="'+VH^.DefCatTVA+'" AND (CC_LIBRE="1" OR CC_LIBRE="2"'+
           ' OR CC_LIBRE="3" OR CC_LIBRE="4") ORDER BY CC_LIBRE ',True) ;
While not QC.Eof do
  BEGIN
  OkTva:=TRUE ; FTva.Values.Add(QC.Fields[0].AsString) ; FTva.Items.Add(QC.Fields[1].AsString) ;
  QC.Next ;
  END ;
Ferme(QC) ;
  inherited;
TabSup.TabVisible:=False ;
FGroupChoixRupt.visible:=False ;
if Not VH^.OuiTvaEnc then BEGIN MsgRien.Execute(10,'','') ; MsgRien.Execute(11,'','') ; PostMessage(Handle,WM_CLOSE,0,0) ; END ;
If Not OkTva Then
  BEGIN
  MsgRien.Execute(9,'','') ; PostMessage(Handle,WM_CLOSE,0,0) ;
  END ;
if FTva.ItemIndex<0 then FTva.ItemIndex:=0 ;
end;

procedure TFQRTvaAcE.InitDivers ;
BEGIN
Inherited ;
{ init. tableaux de montants totaux et ruptures}
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
BRecapG.ForceNewPage:=(CritEdt.GL.FormatPrint.PrSepCompte[10]) ;
TLE_AUXILIAIRE.Caption:=MsgBox.Mess[11] ; TLE_DATECOMPTABLE.Caption:=MsgBox.Mess[12] ;
TE_JOURNAL.Caption:=MsgBox.Mess[13] ; TE_PIECE.Caption:=MsgBox.Mess[14] ;
TE_REFINTERNE.Caption:=MsgBox.Mess[15] ; TE_LIBELLE.Caption:=MsgBox.Mess[16] ;
TLE_ENCAISSE.Caption:=MsgBox.Mess[17] ; TLE_TAUX.Caption:=MsgBox.Mess[18] ;
TLE_BASE.Caption:=MsgBox.Mess[19] ; TLE_TVA.Caption:=MsgBox.Mess[20] ;
END ;

procedure TFQRTvaAcE.RenseigneCritere ;
BEGIN
Inherited ;
RCpte.Visible:=False ; TRa.Visible:=False ; RCpte2.Visible:=False ;
Case TypeTva.ItemIndex of
  0 : BEGIN RCollectee.Caption:='þ' ;    RDeductible.Caption:='o'  ; END ;
  1 : BEGIN RCollectee.Caption:='o'  ;   RDeductible.Caption:='þ' ; END ;
end ;
If CritEdt.GL.CodeTVA='' Then RTauxTVA.Caption:=Traduirememoire('<<Tous>>') Else RTauxTVA.Caption:=RechDom('TTTVA',CritEdt.GL.CodeTVA,FALSE) ;
If CritEdt.GL.RegimeTVA='' Then RRegimeTVA.Caption:=Traduirememoire('<<Tous>>') Else RTauxTVA.Caption:=RechDom('TTREGIMETVA',CritEdt.GL.RegimeTVA,FALSE) ;
END ;

procedure TFQRTvaAcE.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  //if TypeTva.ItemIndex=0 then NatureCpt:='CLI' else NatureCpt:='FOU' ;
  if FSautPage.State=cbGrayed then SautPage:=0 ;
  if FSautPage.State=cbChecked then SautPage:=1 ;
  if FSautPage.State=cbUnChecked then SautPage:=2 ;
  With GL.FormatPrint Do
    BEGIN
    PrSepCompte[8]:=FRecapCpt.Checked ;
    PrSepCompte[10]:=FRecapG.Checked ;
    END ;
  GL.Deductible:=TypeTva.ItemIndex=1 ;
  GL.NumTVA:=FTva.ItemIndex ; // 10538
  GL.CodeTVA:=FTva.Value ;
  If FRegimeTva.ItemIndex>0 Then GL.RegimeTva:=FRegimeTva.Value ;
  GL.LibRegimeTva:=FRegimeTva.Text ;
  GL.PrevalidTva:=FPrevalidTva.Checked ;
  GL.ValidTva:=FValidTva.Checked ;
  END ;
END ;

function TFQRTvaAcE.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
END ;

procedure TFQRTvaAcE.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;

ChargeGroup(LTiers,[MsgBox.Mess[7]]) ; ChargeRecap(LRecapG) ;
ChgMaskChamp(Qr1DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr1CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
END ;

procedure TFQRTvaAcE.MajOkEdition ;
BEGIN
BEGINTRANS ;
InitMove(100,'') ;
Q.Open ;
While not Q.Eof do
  BEGIN
  ExecuteSQL('UPDATE ECRITURE SET E_EDITEETATTVA="X" WHERE E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString
            +'" AND E_DATECOMPTABLE="'+USDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
            +' AND E_EXERCICE="'+Q.FindField('E_EXERCICE').AsString+'" AND E_NUMEROPIECE='+IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger)
            +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString+' '
            +' AND E_QUALIFPIECE="'+Q.FindField('E_QUALIFPIECE').AsString+'"') ;
  Q.Next ;
  END ;
Q.Close ;
FiniMove ;
COMMITTRANS ;
END ;

procedure TFQRTvaAcE.FinirPrint ;
begin
Inherited ;
VideGroup(LTiers) ; VideRecap(LRecapG) ;
if (not AucunMvt) and (NumErreurCrit=1) and MsgTva(2,CritEdt.GL.Deductible,CritEdt.GL.Date21,CritEdt.Date2,CritEdt.GL.LibRegimeTva,CritEdt.GL.CodeTva) then
  if Transactions(MajOkEdition,1)<>OeOk then MessageAlerte(MsgBox.Mess[21]) ;
end;


procedure TFQRTvaAcE.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotalPaye.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[0].TotDebit, CritEdt.AfficheSymbole) ;
TotalBase.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotDebit, CritEdt.AfficheSymbole) ;
TotalTva.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotCredit, CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRTvaAcE.BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
var Signe : Double ;
    MntBase,MntPaye,MntTva : Double ;

begin
  inherited;
EE_PIECELIGECH.Caption:=Q.FindField('E_NUMEROPIECE').AsString+'/'+Q.FindField('E_NUMLIGNE').AsString ;
LE_VALIDE.Caption:=ValideTva(Qr1E_EDITEETATTVA.AsString) ;
Signe:=1.0 ;
If CritEdt.GL.Deductible Then
   BEGIN
   If Arrondi(Qr1Debit.AsFloat,CritEdt.Decimale)=0 Then Signe:=-1.0 ;
   END Else
   BEGIN
   If Arrondi(Qr1Credit.AsFloat,CritEdt.Decimale)=0 Then Signe:=-1.0 ;
   END ;
If CritEdt.GL.Deductible Then MntPaye:=Arrondi(Qr1Debit.AsFloat-Qr1Credit.AsFloat,CritEdt.Decimale)
                         Else MntPaye:=Arrondi(Qr1Credit.AsFloat-Qr1Debit.AsFloat,CritEdt.Decimale) ;
MntBase:=Qr1Base.AsFloat ;
If EuroOk Then
   BEGIN
   If (Not VH^.TenueEuro) And (CritEdt.Monnaie=2) Then MntBase:=PivotToEuro(Qr1Base.AsFloat) ;
   If (VH^.TenueEuro) And (CritEdt.Monnaie=2) Then MntBase:=EuroToPivot(Qr1Base.AsFloat) ;
   END ;
MntBase:=MntBase*Signe ;
TauxTva:=TVA2TAUX(Qr1Regime.AsString,CritEdt.GL.CodeTva,CritEdt.GL.Deductible) ;
MntTva:=Arrondi(TauxTva*MntBase,CritEdt.Decimale) ;
LE_PAYE.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MntPaye, CritEdt.AfficheSymbole ) ;
LE_ECHEENC.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MntBase, CritEdt.AfficheSymbole ) ;
LE_TVA.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,MntTva, CritEdt.AfficheSymbole ) ;
LE_TAUXTVA.Caption:=StrS0(Arrondi(TauxTva*100.0,V_PGI.OkDecV))+' %' ;
AddGroup(LTiers, [Qr1E_AUXILIAIRE], [MntPaye,MntBase,MntTva]) ;
TotEdt[0].TotDebit:=TotEdt[0].TotDebit+MntPaye ;
TotEdt[1].TotDebit:=TotEdt[1].TotDebit+MntBase ;
TotEdt[1].TotCredit:=TotEdt[1].TotCredit+MntTva ;
Quoi:=QuoiMvt(1) ;
end;


procedure TFQRTvaAcE.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1Debit:=TFloatField(Q.FindField('DEBIT')) ;
Qr1Credit:=TFloatField(Q.FindField('CREDIT')) ;
Qr1Couv:=TFloatField(Q.FindField('E_COUVERTURE')) ;
Qr1Base:=TFloatField(Q.FindField('ECHEENC')) ;
Qr1Regime:=TStringField(Q.FindField('E_REGIMETVA')) ;
Qr1E_AUXILIAIRE:=TStringField(Q.FindField('E_AUXILIAIRE')) ;
Qr1T_LIBELLE:=TStringField(Q.FindField('LIBAUX')) ;
Qr1E_JOURNAL:=TStringField(Q.FindField('E_JOURNAL')) ;
Qr1E_NUMEROPIECE:=TIntegerField(Q.FindField('E_NUMEROPIECE')) ;
Qr1E_DATECOMPTABLE:=TDateTimeField(Q.FindField('E_DATECOMPTABLE')) ;
Qr1E_EXERCICE:=TStringField(Q.FindField('E_EXERCICE')) ;
Qr1E_NUMLIGNE:=TIntegerField(Q.FindField('E_NUMLIGNE')) ;
Qr1E_EDITEETATTVA:=TStringField(Q.FindField('E_EDITEETATTVA')) ;
end;

procedure TFQRTvaAcE.QRDLMultiNeedData(var MoreData: Boolean);
Var CodRupt, LibRupt : String ;
    Tot : Array[0..12] of Double ;
    QuelleRupt : integer ;

begin
  inherited;
MoreData:=PrintGroup(LTiers, Q, [Qr1E_AUXILIAIRE], CodRupt, LibRupt, Tot,Quellerupt) ;
If MoreData Then
   BEGIN
   NomTiers.Caption:=CodRupt ;
   PayeTiers.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[0], CritEDT.AfficheSymbole) ;
   BaseTiers.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[1], CritEDT.AfficheSymbole) ;
   TvaTiers.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[2], CritEdt.AfficheSymbole ) ;
   AddRecap(LRecapG,['Total'],[''],[Tot[0],Tot[1],Tot[2]]) ;
   END ;
end;

procedure TFQRTvaAcE.QRDLRecapGNeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    TotRecapG : Array[0..12] Of Double ;
begin
  inherited;
MoreData:=PrintRecap(LRecapG, Cod, Lib, TotRecapG) ;
If MoreData Then
   BEGIN
   LibRecapG.Caption:=Lib ;
   PayeTiersG.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotRecapG[0], CritEdt.AfficheSymbole) ;
   BaseTiersG.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotRecapG[1], CritEdt.AfficheSymbole) ;
   TvaTiersG.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotRecapG[2],CritEdt.AfficheSymbole) ;
   END ;
end;

procedure TFQRTvaAcE.BRecapGAfterPrint(BandPrinted: Boolean);
begin
  inherited;
if BRecapG.ForceNewPage then BRecapG.ForceNewPage:=FALSE ;
end;

procedure TFQRTvaAcE.BFinAuxBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if not CritEdt.GL.FormatPrint.PrSepCompte[8] then PrintBand:=False ;
Quoi:=QuoiAux(1) ;
end;

procedure TFQRTvaAcE.BRecapGBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
If Not (CritEdt.GL.FormatPrint.PrSepCompte[10]) Then PrintBand:=FALSE ;
end;

procedure TFQRTvaAcE.EntetePageBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if (CritEdt.GL.FormatPrint.PrSepCompte[10]) and (Q.Eof) then
   BEGIN
   TLE_AUXILIAIRE.Caption:='' ; TLE_DATECOMPTABLE.Caption:='' ; TE_JOURNAL.Caption:='' ;
   TE_PIECE.Caption:='' ; TE_REFINTERNE.Caption:='' ;
   TE_LIBELLE.Caption:='' ; TLE_TAUX.Caption:='' ;
   END ;
end;

// 10538
procedure TFQRTvaAcE.BValiderClick(Sender: TObject);
begin
if (FTva.ItemIndex = 0) then
	LanceTousLesTaux
else
  inherited;
end;

// 10538
procedure TFQRTvaAcE.LanceTousLesTaux;
var
   i : integer;
begin
	BValider.Enabled := False;
  for i := 1 to FTva.Items.Count - 1 do
  begin
  	FTva.ItemIndex := i;
  	BValiderClick(self);
  end;
  BValider.Enabled := True;
  FTva.ItemIndex := 0;
end;

end.
