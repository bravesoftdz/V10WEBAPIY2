unit QRTvaFac;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, HMsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  StdCtrls, Buttons,
  Hctrls, HQry, ExtCtrls, Mask, Hcompte, ComCtrls, Ent1, Hent1,CritEdt,UtilEdt,QRRupt,
  CpteUtil,
(*{$IFNDEF ESP}
  TVA,
{$ELSE}
{$ENDIF ESP}//XMG 19/01/04*)
  HTB97, HStatus, Saisutil, HPanel, UiUtil, UtilPGI, tCalcCum ;

procedure EditionTvaFac(Directe : boolean) ;

type
  TFQRTvaFac1 = class(TFQR)
    QRLabel1: TQRLabel;
    TLE_VALIDE: TQRLabel;
    TLE_DATECOMPTABLE: TQRLabel;
    TE_JOURNAL: TQRLabel;
    TE_PIECE: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TE_LIBELLE: TQRLabel;
    TLE_ENCAISSE: TQRLabel;
    TLE_FACTURE: TQRLabel;
    MsgBox: THMsgBox;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    TITRE1REP: TQRLabel;
    TITRE2REP: TQRLabel;
    QRLabel2: TQRLabel;
    REPORT1TVA: TQRLabel;
    TotalTva: TQRLabel;
    REPORT2TVA: TQRLabel;
    QRBand1: TQRBand;
    Trait0: TQRLigne;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Ligne1: TQRLigne;
    FDetail: TCheckBox;
    QRLabel3: TQRLabel;
    FTva: THValComboBox;
    HLabel9: THLabel;
    QRLabel6: TQRLabel;
    QRLabel8: TQRLabel;
    BSDMvt: TQRBand;
    QRLabel9: TQRLabel;
    QRLigne1: TQRLigne;
    QRLigne2: TQRLigne;
    TotalBase: TQRLabel;
    FSautPage: TCheckBox;
    FLigneRegime: TCheckBox;
    FLigneTaux: TCheckBox;
    TypeTVA: TRadioGroup;
    TLE_AUXILIAIRE: TQRLabel;
    E_DATECOMPTABLE: TQRDBText;
    E_JOURNAL: TQRDBText;
    EE_PIECELIGECH: TQRLabel;
    E_LETTRAGE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    E_LIBELLE: TQRDBText;
    E_AUXILIAIRE: TQRDBText;
    LE_PAYE: TQRLabel;
    Q2: TQuery;
    S2: TDataSource;
    DLMVTFact: TQRDetailLink;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    EE_PIECELIGECH2: TQRLabel;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    LE_FACTURE: TQRLabel;
    BFinFact: TQRBand;
    PayeFact: TQRLabel;
    FactureFact: TQRLabel;
    BaseFact: TQRLabel;
    TVAFact: TQRLabel;
    BFinAux: TQRBand;
    PayeTiers: TQRLabel;
    FactureTiers: TQRLabel;
    TVATiers: TQRLabel;
    BaseTiers: TQRLabel;
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
    QRLabel14: TQRLabel;
    QRLigne3: TQRLigne;
    LE_TTC: TQRLabel;
    TTCTiers: TQRLabel;
    TTCFact: TQRLabel;
    LE_ECHEENC: TQRLabel;
    RatioTiers: TQRLabel;
    RatioFact: TQRLabel;
    HLabel3: THLabel;
    FRegimeTva: THValComboBox;
    LE_VALIDE: TQRLabel;
    FPrevalidTva: TCheckBox;
    FValidTva: TCheckBox;
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
    LE_ECHEENC2: TQRDBText;
    procedure FormShow(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
    procedure Q2BeforeOpen(DataSet: TDataSet);
    procedure BFinFactBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDMvtBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure Q2AfterOpen(DataSet: TDataSet);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure QRDLRecapGNeedData(var MoreData: Boolean);
    procedure BRecapGAfterPrint(BandPrinted: Boolean);
    procedure BRecapGBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinAuxBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BValiderClick(Sender: TObject);
    procedure FinirPrint ; Override ;
    procedure GenereSQLSub ;
    procedure GenereSQL ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
  private
    AucunMvt,Directe : boolean ;
    Qr1Debit,Qr1Credit,Qr2Debit,Qr2Credit,Qr1Couv,Qr2Couv,Qr2Base : TFloatField ;
    Qr2Regime,Qr1E_AUXILIAIRE,QR1T_LIBELLE,Qr2E_EXERCICE,Qr2E_JOURNAL,Qr2E_QUALIFPIECE,
    //Qr1E_NATUREPIECE,Qr1E_JOURNAL,Qr1E_EXERCICE,Qr1E_EDITEETATTVA,QR1E_GENERAL,QR1E_ETATLETTRAGE,QR1E_LETTRAGE : TStringField ;
    Qr1E_NATUREPIECE,Qr1E_JOURNAL,Qr1E_EXERCICE,Qr1E_EDITEETATTVA,QR1E_GENERAL,QR1E_ETATLETTRAGE,QR1E_LETTRAGE : TStringField ;
    Qr2E_DATECOMPTABLE,Qr1E_DATECOMPTABLE,QR1E_DATEECHEANCE : TDateTimeField ;
    Qr2E_NUMEROPIECE,Qr2E_NUMLIGNE,Qr2E_NUMECHE,Qr1E_NUMEROPIECE,Qr1E_NUMLIGNE : TIntegerField ;
    TotFact,TotPaye,TotBaseFact,TauxTva,TotTTC : Double ;
    TotEdt                : TabTot ;
    LTiers,LRecapG{,LFactLu}  : TStringList ;
    QS : TQuery ;
    OkTiers  : Boolean ;
    TiersAEditer : Boolean ;
    DeuxReq : Boolean ;
    Function  QuoiMvt(i : Integer) : String ;
    Function  QuoiAux : String ;
    Procedure TvaFacZoom(Quoi : String) ;
    procedure MajOkEdition ;
//    PROCEDURE AddListeFact ;
//    Function  Clefact : String ;
    Function  AcompteAReguler : Boolean ;
		Procedure LanceTousLesTaux;
  public
    { Déclarations publiques }
  end;

implementation

uses MsgQrTVA;

{$R *.DFM}

Function TFQRTvaFac1.QuoiAux : String ;
BEGIN
Result:=QR1E_AUXILIAIRE.AsString+' '+'#'+QR1T_LIBELLE.AsString+'@'+'2' ;
END ;

Function TFQRTvaFac1.QuoiMvt(i : Integer) : String ;
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
  2 : St:=QR1E_AUXILIAIRE.AsString+' '+QR1T_LIBELLE.AsString+' '+'#'+
          Qr2E_JOURNAL.AsString+' N° '+IntToStr(Qr2E_NUMEROPIECE.AsInteger)+
          ' '+DateToStr(Qr2E_DateComptable.AsDAteTime)+' . . . '+
          PrintSolde(Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)+
          '@'+'5;'+Qr2E_JOURNAL.AsString+';'+UsDateTime(Qr2E_DATECOMPTABLE.AsDateTime)+';'+
          IntToStr(Qr2E_NUMEROPIECE.AsInteger)+';'+Qr2E_EXERCICE.asString+';'+
          IntToStr(Qr2E_NumLigne.AsInteger)+';' ;
  END ;
Result:=St ;
END ;

Procedure TFQRTvaFac1.TvaFacZoom(Quoi : String) ;
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

procedure EditionTvaFac(Directe : boolean) ;
var QR : TFQRTvaFac1 ;
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
  if Not VH^.OuiTvaEnc then
    BEGIN
    {JP 15/09/05 : FQ 16030 : Fusion des deux messages
    HShowMessage('10;Tva sur encaissements;Le module n''est pas installé !;E;O;O;O;','','') ;
    HShowMessage('11;Tva sur encaissements;Vous pouvez paramétrer ce module dans les paramètres société;E;O;O;O;','','') ; Exit ;}
    HShowMessage('11;Tva sur encaissements;Le module n''est pas installé.'#13#13 +
                 'Vous pouvez activer ce module dans les paramètres société.;E;O;O;O;','','');
    Exit ;
    END ;
  END ;
PP:=FindInsidePanel ;
QR:= TFQRTvaFac1.Create(Application) ;
Edition.Etat:=etTvaFac ;
QR.Directe:=Directe ;
QR.QRP.QRPrinter.OnSynZoom:=QR.TvaFacZoom ;
QR.InitType (nbGen,neGL,msGenEcr,'QRTvaFac','',TRUE,FALSE,TRUE,Edition) ;
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

procedure TFQRTvaFac1.GenereSQL ;
Var StValid : String ;
    FiltreCompenTP: String;         {FP 19/05/2006 FQ17981}
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
{b FP 19/05/2006 FQ17981}
FiltreCompenTP := ''; { YMO 28/06/2006 Mise en stand by des modifs pour fonctionnement tiers-payeurs
if CritEdt.GL.Deductible and (VH^.JalATP <> '') then          {Tiers payeur fournisseur
  FiltreCompenTP := FiltreCompenTP +
  ' OR (Q.E_JOURNAL="'+VH^.JalATP+'" AND Q.E_NATUREPIECE="FF" AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD" OR G_NATUREGENE="COS") AND T_PAYEUR<>"")';
if (not CritEdt.GL.Deductible) and (VH^.JalVTP <> '') then    {Tiers payeur client
  FiltreCompenTP := FiltreCompenTP +
  ' OR (Q.E_JOURNAL="'+VH^.JalVTP+'" AND Q.E_NATUREPIECE="FC" AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD" OR G_NATUREGENE="COS") AND T_PAYEUR<>"")';
if TCompensation.IsCompensation then
  FiltreCompenTP := FiltreCompenTP +
  ' OR (Q.E_JOURNAL="'+TCompensation.GetJalCompensation+'" AND Q.E_NATUREPIECE="OD" AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD" OR G_NATUREGENE="COS" OR G_NATUREGENE="COC"))';
{e FP 19/05/2006}
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('Select DISTINCT Q.E_GENERAL, Q.E_AUXILIAIRE, Q.E_EXERCICE, Q.E_DATECOMPTABLE, Q.E_NUMEROPIECE, Q.E_QUALIFPIECE, Q.E_NUMLIGNE,') ;
Q.SQL.Add(       ' Q.E_REFINTERNE, Q.E_LIBELLE, Q.E_DEVISE, Q.E_LETTRAGE, Q.E_JOURNAL, Q.E_NATUREPIECE, Q.E_ETATLETTRAGE,') ;
Q.SQL.Add(       ' Q.E_CONTREPARTIEGEN, Q.E_ECHEENC'+IntToStr(CritEdt.GL.NumTva)+' ECHEENC, T_LIBELLE LIBAUX, Q.E_EDITEETATTVA, Q.E_DATEECHEANCE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN Q.SQL.Add('Q.E_DEBIT DEBIT, Q.E_CREDIT CREDIT, Q.E_COUVERTURE COUVERTURE') ; END ;
  1 : BEGIN Q.SQL.Add('Q.E_DEBITDEV DEBIT, Q.E_CREDITDEV CREDIT, Q.E_COUVERTUREDEV COUVERTURE') ; END ;
  2 : BEGIN Q.SQL.Add('Q.E_DEBITEURO DEBIT, Q.E_CREDITEURO CREDIT, Q.E_COUVERTUREEURO COUVERTURE') ; END ;
end ;
Q.SQL.Add(' From ECRITURE Q ') ;
Q.SQL.Add(' Left Join GENERAUX on Q.E_CONTREPARTIEGEN=G_GENERAL ') ;
Q.SQL.Add(' Left Join ECRITURE Q2 on Q.E_AUXILIAIRE = Q2.E_AUXILIAIRE and Q.E_GENERAL = Q2.E_GENERAL and Q.E_ETATLETTRAGE= Q2.E_ETATLETTRAGE and Q.E_LETTRAGE = Q2.E_LETTRAGE ');
Q.SQL.Add(' Left Join TIERS on Q.E_AUXILIAIRE=T_AUXILIAIRE WHERE Q.E_AUXILIAIRE<>"" ') ;
If CritEdt.GL.Deductible Then Q.SQL.Add(SQLCollFouEncTVA('Q'))
                         Else Q.SQL.Add(SQLCollCliEncTVA('Q')) ;
// Condition sur exercice uniquement en facture directe
if Directe and (CritEdt.Exo.Code<>'') then Q.SQL.Add(' And Q.E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
// Factures directes : date comptable
if Directe then Q.SQL.Add('AND Q.E_DATECOMPTABLE>="'+usDateTime(CritEdt.GL.Date21)+'" And Q.E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'"')
// Factures indirectes : date échéance
           else Q.SQL.Add('AND Q.E_DATEECHEANCE>="'+usDateTime(CritEdt.GL.Date21)+'" And Q.E_DATEECHEANCE<="'+usdatetime(CritEdt.Date2)+'"') ;
//if CritEdt.GL.RegimeTva<>'' then Q.SQL.Add(' And Q.E_REGIMETVA="'+CritEdt.GL.RegimeTva+'" ') ;
//if CritEdt.Etab<>'' Then Q.SQL.Add(' And Q.E_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
//if Not Directe then Q.SQL.Add(' AND Q.E_JOURNAL<>"'+VH^.JalEcartEuro+'" ') ;
If CritEdt.GL.Deductible Then
   BEGIN
   Q.SQL.Add(' AND (T_NATUREAUXI="FOU" Or T_NATUREAUXI="AUC") ') ;
   If Directe then
     {b FP 19/05/2006 FQ17981}
     begin
     Q.SQL.Add(' AND ('+
     ' ((Q.E_NATUREPIECE="RF" OR Q.E_NATUREPIECE="OD" OR Q.E_NATUREPIECE="OF") AND (G_NATUREGENE="BQE" OR G_NATUREGENE="CAI"))'+
     FiltreCompenTP+
     ' )');
     end
     {e FP 19/05/2006}
     else Q.SQL.Add(' AND ((Q.E_NATUREPIECE="RF" OR Q.E_NATUREPIECE="OD" OR Q.E_NATUREPIECE="OF") AND G_NATUREGENE<>"BQE" AND G_NATUREGENE<>"CAI")') ;
   END Else
   BEGIN
//   Q.SQL.Add(SQLCollCliEncTVA('Q')) ;
   Q.SQL.Add(' AND (T_NATUREAUXI="CLI" Or T_NATUREAUXI="AUD") ') ;
   If Directe then
     {b FP 19/05/2006 FQ17981}
     begin
     Q.SQL.Add(' AND ('+
     ' ((Q.E_NATUREPIECE="RC" OR Q.E_NATUREPIECE="OD" OR Q.E_NATUREPIECE="OC") AND (G_NATUREGENE="BQE" OR G_NATUREGENE="CAI"))'+
     FiltreCompenTP+
     ' )');
     end
     {e FP 19/05/2006}
     else Q.SQL.Add(' AND ((Q.E_NATUREPIECE="RC" OR Q.E_NATUREPIECE="OD" OR Q.E_NATUREPIECE="OC") AND G_NATUREGENE<>"BQE" AND G_NATUREGENE<>"CAI")') ;
   END ;
StValid:=' AND (Q.E_EDITEETATTVA="-" ' ;
If CritEdt.GL.PrevalidTVA Then StValid:=StValid+' OR Q.E_EDITEETATTVA="X" ' ;
If CritEdt.GL.ValidTVA Then StValid:=StValid+' OR Q.E_EDITEETATTVA="#" ' ;
StValid:=StValid+') ' ;
Q.SQL.Add(StValid) ;
//Q.SQL.Add(' AND (E_EDITEETATTVA="-" OR E_EDITEETATTVA="X")' ) ;
Q.SQL.Add(' AND Q.E_LETTRAGE<>""') ;
If Not DeuxReq Then
  BEGIN
  If CritEdt.GL.Deductible Then
     BEGIN
     Q.SQL.Add('AND EXISTS(Select * from ECRITURE WHERE  E_AUXILIAIRE=Q.E_AUXILIAIRE And  E_ETATLETTRAGE=Q.E_ETATLETTRAGE And E_GENERAL=Q.E_GENERAL') ;
     Q.SQL.Add('AND E_LETTRAGE=Q.E_LETTRAGE AND (E_NATUREPIECE="FF" OR E_NATUREPIECE="AF") AND E_ECHEENC'+IntToStr(CritEdt.GL.NumTva)+'<>0)') ;
     END Else
     BEGIN
     Q.SQL.Add('AND EXISTS(Select * ') ;
     Q.SQL.Add('from ECRITURE Q1 WHERE  Q1.E_AUXILIAIRE=Q.E_AUXILIAIRE And Q1.E_GENERAL=Q.E_GENERAL And  Q1.E_ETATLETTRAGE=Q.E_ETATLETTRAGE ') ;
     Q.SQL.Add('AND Q1.E_LETTRAGE=Q.E_LETTRAGE AND (Q1.E_NATUREPIECE="FC" OR Q1.E_NATUREPIECE="AC") AND Q1.E_ECHEENC'+IntToStr(CritEdt.GL.NumTva)+'<>0)') ;
     END ;
  END ;
{ Construction de la clause Order By de la SQL }
If CritEdt.GL.Deductible Then // FQ 10911
  Q.SQL.Add(' AND (Q2.E_NATUREPIECE="FF" OR Q2.E_NATUREPIECE="AF") ')
else
  Q.SQL.Add(' AND (Q2.E_NATUREPIECE="FC" OR Q2.E_NATUREPIECE="AC") ');
Q.SQL.Add(' AND Q2.E_ECHEENC'+IntToStr(CritEdt.GL.NumTva)+'<>0 ');
Q.SQL.Add(' Order By Q.E_AUXILIAIRE,Q.E_GENERAL,Q.E_ETATLETTRAGE,Q.E_LETTRAGE,Q.E_DATECOMPTABLE') ;
StValid := Q.SQL.Text;

ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
AucunMvt:=(Q.Eof) ;
GenereSQLSub ;
END ;

procedure TFQRTvaFac1.GenereSQLSub ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Q2.Close ; Q2.SQL.Clear ;
Q2.SQL.Add('Select E_GENERAL,E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_REGIMETVA,') ;
Q2.SQL.Add(       'E_REFINTERNE,E_LIBELLE,E_DEVISE,E_LETTRAGE,E_JOURNAL,E_NATUREPIECE,E_CONTREPARTIEGEN,') ;
Q2.SQL.Add(       'E_ECHEENC'+IntToStr(CritEdt.GL.NumTva)+' AS ECHEENC2, E_QUALIFPIECE,E_NUMECHE, E_ETATLETTRAGE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN Q2.SQL.Add('E_DEBIT DEBIT2,E_CREDIT CREDIT2, E_COUVERTURE COUVERTURE2') ; END ;
  1 : BEGIN Q2.SQL.Add('E_DEBITDEV DEBIT2,E_CREDITDEV CREDIT2, E_COUVERTUREDEV COUVERTURE2') ; END ;
end ;
Q2.SQL.Add(' From ECRITURE ') ;
Q2.SQL.Add(' Where E_AUXILIAIRE=:AUX And E_GENERAL=:GEN And E_ETATLETTRAGE=:ETATL And E_LETTRAGE=:LETT') ;
If CritEdt.GL.Deductible Then Q2.SQL.Add(' AND (E_NATUREPIECE="FF" OR E_NATUREPIECE="AF") ')
                         else Q2.SQL.Add(' AND (E_NATUREPIECE="FC" OR E_NATUREPIECE="AC") ') ;
Q2.SQL.Add(' AND E_ECHEENC'+IntToStr(CritEdt.GL.NumTva)+'<>0 ') ;
if CritEdt.Etab<>'' Then Q2.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
if CritEdt.GL.RegimeTva<>'' then Q2.SQL.Add(' And E_REGIMETVA="'+CritEdt.GL.RegimeTva+'" ') ;
//YMO 28/06/2006 Mise en stand by des modifs pour fonctionnement tiers-payeurs
{b FP 19/05/2006 FQ17981
if Directe then
  begin
  if CritEdt.GL.Deductible and (VH^.JalATP <> '') then          {Tiers payeur fournisseur
    Q2.SQL.Add(' AND (E_JOURNAL<>"'+VH^.JalATP+'" OR (E_JOURNAL="'+VH^.JalATP+'" AND E_PIECETP="") )');
  if (not CritEdt.GL.Deductible) and (VH^.JalVTP <> '') then          {Tiers payeur fournisseur
    Q2.SQL.Add(' AND (E_JOURNAL<>"'+VH^.JalVTP+'" OR (E_JOURNAL="'+VH^.JalVTP+'" AND E_PIECETP="") )');
  end;
{e FP 19/05/2006}
{ Construction de la clause Order By de la SQL }
Q2.SQL.Add(' Order By E_AUXILIAIRE,E_GENERAL,E_ETATLETTRAGE,E_LETTRAGE,E_DATECOMPTABLE') ;
ChangeSQL(Q2) ; //Q2.Prepare ;
PrepareSQLODBC(Q2) ;
Q2.Open ;
END ;

procedure TFQRTvaFac1.FormShow(Sender: TObject);
var QC : TQuery ;
    StC : String ;
    OkTva : Boolean ;
begin
DeuxReq:=TRUE ;
if not Directe then
  BEGIN
  Caption:=MsgBox.Mess[10] ;
  HLabel6.Caption:=MsgBox.Mess[8] ;
  StC:=MsgBox.Mess[8] ;
  if Pos('&',StC)<>0 then Delete(StC,Pos('&',StC),1) ;
  QRLabel4.Caption:=StC ;
  END else Caption:=MsgBox.Mess[9] ;
  UpdateCaption(TForm(Self));{JP 18/08/04 FQ 10909}
If Directe Then
   BEGIN
   HelpContext:=7650000 ;
   //Standards.HelpContext:=7650100 ;
   //Avances.HelpContext:=7650200 ;
   //Mise.HelpContext:=7650300 ;
   //Option.HelpContext:=7650400 ;
   END Else
   BEGIN
   HelpContext:=7651050 ;
   //Standards.HelpContext:=7651150 ;
   //Avances.HelpContext:=7651250 ;
   //Mise.HelpContext:=7651300 ;
   //Option.HelpContext:=7651400 ;
   END ;

// Rempli le combo des TVA
FTva.Values.Clear ; FTva.Items.Clear ; OkTva:=FALSE ;
// Début VL 04/12/2002
FTva.Values.Add('') ;
FTva.Items.Add(TraduireMemoire('<<Tous>>')) ;
// Fin VL
QC:=OpenSQL('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="'+VH^.DefCatTVA+'" AND (CC_LIBRE="1" OR CC_LIBRE="2"'+
           ' OR CC_LIBRE="3" OR CC_LIBRE="4") ORDER BY CC_LIBRE',True) ;
While not QC.Eof do
  BEGIN
  OkTva:=TRUE ; FTva.Values.Add(QC.Fields[0].AsString) ; FTva.Items.Add(QC.Fields[1].AsString) ;
  QC.Next ;
  END ;
Ferme(QC) ;
  inherited;
TabSup.TabVisible:=False ;
FGroupChoixRupt.Enabled:=False ;
FMontant.ItemIndex:=0 ;
If Not OkTva Then BEGIN MsgRien.Execute(9,'','') ; PostMessage(Handle,WM_CLOSE,0,0) ; END ;
if Not VH^.OuiTvaEnc then BEGIN MsgRien.Execute(10,'','') ; PostMessage(Handle,WM_CLOSE,0,0) ; END ;
if FTva.ItemIndex<0 then FTva.ItemIndex:=0 ;
If Not Directe Then
  BEGIN
  FExercice.Vide:=TRUE ; FExercice.Reload ; FExercice.Value:=VH^.Entree.Code ;
  END ;
end;

procedure TFQRTvaFac1.InitDivers ;
BEGIN
Inherited ;
{ init. tableaux de montants totaux et ruptures}
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
TotFact:=0 ; TotPaye:=0 ; TotBaseFact:=0 ; TotTTC:=0 ;
Fillchar(CritEDT.GL.FormatPrint.Report,SizeOf(CritEDT.GL.FormatPrint.Report),#0) ;
BRecapG.ForceNewPage:=TRUE ;
If Directe Then E_DATECOMPTABLE.DataField:='E_DATECOMPTABLE' Else E_DATECOMPTABLE.DataField:='E_DATEECHEANCE' ;
END ;

procedure TFQRTvaFac1.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  if FSautPage.State=cbGrayed then SautPage:=0 ;
  if FSautPage.State=cbChecked then SautPage:=1 ;
  if FSautPage.State=cbUnChecked then SautPage:=2 ;
  With GL.FormatPrint Do
    BEGIN
    PrSepCompte[2]:=FDetail.Checked ;
    PrSepCompte[8]:=FRecapCpt.Checked ;
    PrSepCompte[10]:=FRecapG.Checked ;
    END ;
  GL.Deductible:=TypeTva.ItemIndex=1 ;
// VL 05/12/2002
  GL.NumTVA:=FTva.ItemIndex ;
//  GL.NumTVA:=FTva.ItemIndex+1 ;
  GL.CodeTVA:=FTva.Value ;
  If FRegimeTva.ItemIndex>0 Then GL.RegimeTva:=FRegimeTva.Value ;
  GL.LibRegimeTva:=FRegimeTva.Text ;
  GL.PrevalidTva:=FPrevalidTva.Checked ;
  GL.ValidTva:=FValidTva.Checked ;
  END ;
END ;

function TFQRTvaFac1.CritOk : Boolean ;
Var Q 						 : TQuery ;
		D1, DateSaisie : TDateTime ;
    Exo						 : String ;
BEGIN
Result:=Inherited CritOK ;

	// Test sur la date de début d'échéance uniquement en mode "facture par étape" (Directe = false)
  // De plus on courcircuite les éventuels erreurs 101 (Tests sur dates comptables)
	if (not Directe) and (Result or (NumErreurCrit=101)) then
  	begin
    Result 				:= True ;									// Réinit pour la cas erreur 101
    NumErreurCrit	:= 1 ;										// Réinit pour la cas erreur 101
    DateSaisie 		:= CritEdt.DateDeb ;
    Exo 					:= FExercice.Value ;
    if (IsValidDate(DateToStr(DateSaisie))) and (Exo<>'') then
    	begin
      Q:=OpenSQL('SELECT EX_DATEDEBUT FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"' ,TRUE) ;
      if Not Q.EOF
        then D1 := Q.FindField('EX_DATEDEBUT').asDateTime
        else D1 := iDate1900 ;
      Ferme(Q) ;
      // La date saisie de début d'échance doit être supérieur ou égale à la date de début d'exercice
      if D1 > DateSaisie then
        begin
        Result 				:= False ;
        NumErreurCrit := 111 ;
        end ;
      end ;
    end ;

END ;

procedure TFQRTvaFac1.RenseigneCritere ;
BEGIN
Inherited ;
Case TypeTva.ItemIndex of
  0 : BEGIN RCollectee.Caption:='þ' ;    RDeductible.Caption:='o'  ; END ;
  1 : BEGIN RCollectee.Caption:='o'  ;   RDeductible.Caption:='þ' ; END ;
end ;
If CritEdt.GL.CodeTVA='' Then RTauxTVA.Caption:=Traduirememoire('<<Tous>>') Else RTauxTVA.Caption:=RechDom('TTTVA',CritEdt.GL.CodeTVA,FALSE) ;
If CritEdt.GL.RegimeTVA='' Then RRegimeTVA.Caption:=Traduirememoire('<<Tous>>') Else RRegimeTVA.Caption:=RechDom('TTREGIMETVA',CritEdt.GL.RegimeTVA,FALSE) ;
END ;

procedure TFQRTvaFac1.ChoixEdition ;
{ Initialisation des options d'édition }
Var St : String ;
BEGIN
Inherited ;
ChargeGroup(LTiers,[MsgBox.Mess[7]]) ; ChargeRecap(LRecapG) ;
//LFactLu:=TStringList.Create ; LFactLu.Sorted:=TRUE ; LFactLu.Duplicates:=DupIgnore ;
ChgMaskChamp(Qr1DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr1CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr2DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr2CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
If DeuxReq Then
  BEGIN
  If CritEdt.GL.Deductible Then
    BEGIN
    St:='Select * From Ecriture where E_AUXILIAIRE=:AUX And E_GENERAL=:GEN And  E_ETATLETTRAGE=:EL  '
       +'AND E_LETTRAGE=:LE AND (E_NATUREPIECE="FF" OR E_NATUREPIECE="AF") AND E_ECHEENC'+IntToStr(CritEdt.GL.NumTva)+'<>0' ;
    END Else
    BEGIN
    St:='Select * From Ecriture where E_AUXILIAIRE=:AUX And E_GENERAL=:GEN And  E_ETATLETTRAGE=:EL  '
       +'AND E_LETTRAGE=:LE AND (E_NATUREPIECE="FC" OR E_NATUREPIECE="AC") AND E_ECHEENC'+IntToStr(CritEdt.GL.NumTva)+'<>0' ;
    END ;
  if CritEdt.Etab<>'' Then St:=St+' And E_ETABLISSEMENT="'+CritEdt.Etab+'" ' ;
  if CritEdt.GL.RegimeTva<>'' then St:=St+' And E_REGIMETVA="'+CritEdt.GL.RegimeTva+'" ' ;
  QS:=PrepareSQL(St,TRUE) ;
  END ;
END ;

procedure TFQRTvaFac1.MajOkEdition ;
BEGIN
BEGINTRANS ;
InitMove(100,'') ;
Q.Open ;
FetchSQLODBC(Q) ;
While not Q.Eof do
  BEGIN
  ExecuteSQL('UPDATE ECRITURE SET E_EDITEETATTVA="X" WHERE E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString
            +'" AND E_DATECOMPTABLE="'+USDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
            +' AND E_EXERCICE="'+Q.FindField('E_EXERCICE').AsString+'" AND E_NUMEROPIECE='+IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger)
            +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString+' '
            +' AND E_QUALIFPIECE="'+Q.FindField('E_QUALIFPIECE').AsString+'" '
            +' AND E_EDITEETATTVA<>"#"') ;
  Q.Next ;
  END ;
Q.Close ;
FiniMove ;
COMMITTRANS ;
END ;

procedure TFQRTvaFac1.FinirPrint ;
Var i : Integer ;
begin
Inherited ;
If DeuxReq Then Ferme(QS) ;
VideGroup(LTiers) ; VideRecap(LRecapG) ;
//LFactLu.Clear ; LFactLu.Free ;
If Directe Then i:=0 Else i:=1 ;
if (not AucunMvt) and (NumErreurCrit=1) and MsgTva(i,CritEdt.GL.Deductible,CritEdt.GL.Date21,CritEdt.Date2,CritEdt.GL.LibRegimeTva,CritEdt.GL.CodeTva) then
  if Transactions(MajOkEdition,1)<>OeOk then MessageAlerte(MsgBox.Mess[11]) ;
end;


procedure TFQRTvaFac1.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotalBase.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotDebit, CritEdt.AfficheSymbole) ;
TotalTva.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotCredit, CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

Function TFQRTvaFac1.AcompteAReguler : Boolean ;
BEGIN
Result:=FALSE ;
END ;

procedure TFQRTvaFac1.BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
//Var St : String ;
begin
  inherited;
TiersAEditer:=TRUE ;
LE_VALIDE.Caption:=ValideTva(Qr1E_EDITEETATTVA.AsString) ;
TotFact:=0 ; TotPaye:=0 ; TotTTC:=0 ; TotBaseFact:=0 ;
If DeuxReq Then
  BEGIN
  QS.Close ;
  QS.Params[0].AsString:=Qr1E_AUXILIAIRE.ASString ;
  QS.Params[1].AsString:=Qr1E_GENERAL.ASString ;
  QS.Params[2].AsString:=Qr1E_ETATLETTRAGE.ASString ;
  QS.Params[3].AsString:=Qr1E_LETTRAGE.ASString ;
  QS.Open ;
  If QS.Eof Then
  	BEGIN
    PrintBand:=FALSE ;
    TiersAEditer:=FALSE ;
    END ;
  QS.Close ;
  If Not PrintBand Then Exit ;
  END ;
EE_PIECELIGECH.Caption:=Q.FindField('E_NUMEROPIECE').AsString+'/'+Q.FindField('E_NUMLIGNE').AsString ;
If CritEdt.GL.Deductible Then TotPaye:=Arrondi(Qr1Debit.AsFloat-Qr1Credit.AsFloat,CritEdt.Decimale)
                         Else TotPaye:=Arrondi(Qr1Credit.AsFloat-Qr1Debit.AsFloat,CritEdt.Decimale) ;
LE_PAYE.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPaye, CritEdt.AfficheSymbole ) ;
LE_ECHEENC.Caption:='' ;
If ((Qr1E_NATUREPIECE.AsString='OF') Or (Qr1E_NATUREPIECE.AsString='OC')) And AcompteAReguler Then LE_ECHEENC.Caption:=MsgRien.Mess[8] ;
Quoi:=QuoiMvt(1) ;
AddGroup(LTiers, [Qr1E_AUXILIAIRE], [TotPaye,0,0,0,0]) ;
OkTiers:=TRUE ;
end;

{Function TFQRTvaFac1.Clefact : String ;
BEGIN
Result:=Format_String(Qr2E_EXERCICE.AsString,3)+
        Format_String(Qr2E_JOURNAL.AsString,3)+
        DateToStr(Qr2E_DATECOMPTABLE.AsDateTime)+
        FormatFloat('0000000000',Qr2E_NUMEROPIECE.AsInteger)+
        FormatFloat('00000',Qr2E_NUMLIGNE.AsInteger)+
        FormatFloat('00',Qr2E_NUMECHE.AsInteger)+
        Format_String(Qr2E_QUALIFPIECE.AsString,3) ;
END ;

PROCEDURE TFQRTvaFac1.AddListeFact ;
BEGIN
//LFactLu.Add(CleFact) ;
END ;
}

procedure TFQRTvaFac1.BSDMvtBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var MntFact,MntBaseFact,MntTTC,{RatioFacture,RatioEncaisse,RatioDef,} QBaseCalc : Double ;
    Signe : Double ;
//    i : Integer ;
begin
  inherited;
If Not TiersAEditer Then BEGIN PrintBand:=FALSE ; Exit ;END ;
LE_EcheEnc2.Caption:='' ;
Signe:=1.0 ;
If CritEdt.GL.Deductible Then
   BEGIN
   If Arrondi(Qr2Credit.AsFloat,CritEdt.Decimale)=0 Then Signe:=-1.0 ;
   END Else
   BEGIN
   If Arrondi(Qr2Debit.AsFloat,CritEdt.Decimale)=0 Then Signe:=-1.0 ;
   END ;
{ TTC FACTURE : }
MntTTC:=Arrondi((Qr2Debit.AsFloat+Qr2Credit.AsFloat)*Signe,CritEdt.Decimale) ;
{ PAYE FACTURE : }
MntFact:=Arrondi(Qr2Couv.AsFloat*Signe,CritEdt.Decimale) ;
{ Base HT à déclarer : }
QBaseCalc:=Qr2Base.AsFloat ;
If EuroOk Then
   BEGIN
   If (Not VH^.TenueEuro) And (CritEdt.Monnaie=2) Then QBaseCalc:=PivotToEuro(Qr2Base.AsFloat) ;
   If (VH^.TenueEuro) And (CritEdt.Monnaie=2) Then QBaseCalc:=EuroToPivot(Qr2Base.AsFloat) ;
   END ;
{ Base HT de la facture : }
MntBaseFact:=Arrondi(QBaseCalc,CritEdt.Decimale) ;
{ Taux TVA Attaché à la base HT : }
TauxTva:=Arrondi(TVA2TAUX(Qr2Regime.AsString,CritEdt.GL.CodeTva,CritEdt.GL.Deductible),4) ;
{ TVA de la facture : }
LE_Facture.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MntFact, CritEdt.AfficheSymbole ) ;
LE_TTC.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MntTTC, CritEdt.AfficheSymbole ) ;
LE_EcheEnc2.Caption:='['+AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MntBaseFact, CritEdt.AfficheSymbole )+']' ;
TotFact:=Arrondi(TotFact+MntFact,CritEdt.Decimale) ;
TotBaseFact:=Arrondi(TotBaseFact+MntBaseFact,CritEdt.Decimale) ;
TotTTC:=Arrondi(TotTTC+MntTTC,CritEdt.Decimale) ;
EE_PIECELIGECH2.Caption:=Q2.FindField('E_NUMEROPIECE').AsString+'/'+Q2.FindField('E_NUMLIGNE').AsString ;
//If LFactLu.Find(CleFact,i) Then BEGIN MntTTC:=0 ; MntFact:=0 ; END ;
//If OkTiers Then AddGroup(LTiers, [Qr1E_AUXILIAIRE], [0,MntFact,MntBaseCalc,MntBaseFact,MntTTC]) ;
If Not CritEdt.GL.FormatPrint.PrSepCompte[2] Then PrintBand:=FALSE ;
//AddListeFact ;
Quoi:=QuoiMvt(2) ;
end;

procedure TFQRTvaFac1.BFinFactBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var LaBaseFact,RatioDef,MntTva : Double ;
begin
  inherited;
If Not TiersAEditer Then BEGIN PrintBand:=FALSE ; Exit ;END ;
PayeFact.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPaye, CritEdt.AfficheSymbole ) ;
FactureFact.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotFact, CritEdt.AfficheSymbole ) ;
TTCFact.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotTTC, CritEdt.AfficheSymbole ) ;
RatioDef:=1 ; If TotTTC<>0 Then RatioDef:=TotPaye/TotTTC ; If RatioDef>1.0 Then RatioDef:=1 ; If RatioDef<-1.0 Then RatioDef:=1 ;
RatioFact.Caption:=StrfMontant(RatioDef*100,7,2,'',TRUE)+'%' ;
LaBaseFact:=TotBaseFact*RatioDef ;
(*
Ratio:=1 ; If TotFact<>0 Then Ratio:=TotPaye/TotFact ;
RatioFact.Caption:=StrfMontant(Ratio*100,7,2,'',TRUE)+'%' ;
BaseFact2.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotBase*Ratio, CritEdt.AfficheSymbole ) ;
*)

BaseFact.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, LaBaseFact, CritEdt.AfficheSymbole ) ;
MntTva:=Arrondi(TauxTva*LaBaseFact,CritEdt.Decimale) ;
TvaFact.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MntTva, CritEdt.AfficheSymbole ) ;
TotEdt[1].TotDebit :=Arrondi(TotEdt[1].TotDebit+LaBaseFact,CritEdt.Decimale) ;
TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit+MntTva,CritEdt.Decimale) ;
If Not CritEdt.GL.FormatPrint.PrSepCompte[2] Then PrintBand:=FALSE ;
Quoi:=QuoiAux ;
If OkTiers Then AddGroup(LTiers, [Qr1E_AUXILIAIRE], [0,TotFact,LaBaseFact,TotBaseFact,TotTTC]) ;
end;

procedure TFQRTvaFac1.Q2AfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr2Debit:=TFloatField(Q2.FindField('DEBIT2')) ;
Qr2Credit:=TFloatField(Q2.FindField('CREDIT2')) ;
Qr2Base:=TFloatField(Q2.FindField('ECHEENC2')) ;
Qr2Couv:=TFloatField(Q2.FindField('COUVERTURE2')) ;
Qr2Regime:=TStringField(Q2.FindField('E_REGIMETVA')) ;
Qr2E_EXERCICE:=TStringField(Q2.FindField('E_EXERCICE')) ;
Qr2E_JOURNAL:=TStringField(Q2.FindField('E_JOURNAL')) ;
Qr2E_NUMEROPIECE:=TIntegerField(Q2.FindField('E_NUMEROPIECE')) ;
Qr2E_NUMLIGNE:=TIntegerField(Q2.FindField('E_NUMLIGNE')) ;
Qr2E_NUMECHE:=TIntegerField(Q2.FindField('E_NUMECHE')) ;
Qr2E_DATECOMPTABLE:=TDateTimeField(Q2.FindField('E_DATECOMPTABLE')) ;
Qr2E_QUALIFPIECE:=TStringField(Q2.FindField('E_QUALIFPIECE')) ;
end;

procedure TFQRTvaFac1.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1Debit:=TFloatField(Q.FindField('DEBIT')) ;
Qr1Credit:=TFloatField(Q.FindField('CREDIT')) ;
Qr1Couv:=TFloatField(Q.FindField('COUVERTURE')) ;
Qr1E_AUXILIAIRE:=TStringField(Q.FindField('E_AUXILIAIRE')) ;
Qr1T_LIBELLE:=TStringField(Q.FindField('LIBAUX')) ;
Qr1E_NATUREPIECE:=TStringField(Q.FindField('E_NATUREPIECE')) ;
Qr1E_JOURNAL:=TStringField(Q.FindField('E_JOURNAL')) ;
Qr1E_NUMEROPIECE:=TIntegerField(Q.FindField('E_NUMEROPIECE')) ;
Qr1E_DATECOMPTABLE:=TDateTimeField(Q.FindField('E_DATECOMPTABLE')) ;
Qr1E_DATEECHEANCE:=TDateTimeField(Q.FindField('E_DATEECHEANCE')) ;
Qr1E_EXERCICE:=TStringField(Q.FindField('E_EXERCICE')) ;
Qr1E_NUMLIGNE:=TIntegerField(Q.FindField('E_NUMLIGNE')) ;
Qr1E_EDITEETATTVA:=TStringField(Q.FindField('E_EDITEETATTVA')) ;
Qr1E_GENERAL:=TStringField(Q.FindField('E_GENERAL')) ;
Qr1E_LETTRAGE:=TStringField(Q.FindField('E_LETTRAGE')) ;
Qr1E_ETATLETTRAGE:=TStringField(Q.FindField('E_ETATLETTRAGE')) ;
end;

procedure TFQRTvaFac1.Q2BeforeOpen(DataSet: TDataSet);
begin
  inherited;
Q2.Params[0].AsString:=Q.FindField('E_AUXILIAIRE').AsString ;
Q2.Params[1].AsString:=Q.FindField('E_GENERAL').AsString ;
Q2.Params[2].AsString:=Q.FindField('E_ETATLETTRAGE').AsString ;
Q2.Params[3].AsString:=Q.FindField('E_LETTRAGE').AsString ;
end;


procedure TFQRTvaFac1.QRDLMultiNeedData(var MoreData: Boolean);
Var CodRupt, LibRupt : String ;
    Tot : Array[0..12] of Double ;
    QuelleRupt : integer ;
    MntTvaTiers,RatioFacture,RatioEncaisse,RatioDef : Double ;
begin
  inherited;
//If Not OkTiers Then Exit ;
If Not CritEdt.GL.FormatPrint.PrSepCompte[8] Then BEGIN MoreData:=FALSE ; Exit ; END ;
MoreData:=PrintGroup(LTiers, Q, [Qr1E_AUXILIAIRE], CodRupt, LibRupt, Tot,Quellerupt) ;
If MoreData Then
   BEGIN
   OkTiers:=FALSE ;
   RatioFacture:=1.0 ; RatioEncaisse:=1 ;
   NomTiers.Caption:=Qr1T_LIBELLE.AsString ;
   PayeTiers.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[0], CritEDT.AfficheSymbole) ;
   FactureTiers.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[1], CritEDT.AfficheSymbole) ;
   TTCTiers.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[4], CritEDT.AfficheSymbole) ;
//   MntTvaTiers:=Arrondi(TauxTva*Tot[2],CritEdt.Decimale) ;
   If Tot[4]<>0  Then RatioFacture:=Tot[1]/Tot[4] ;
   If Tot[1]<>0 Then RatioEncaisse:=Tot[0]/Tot[1] ;
   RatioDef:=RatioFacture*RatioEncaisse ; If RatioDef>1.0 Then RatioDef:=1 ; If RatioDef<-1.0 Then RatioDef:=1 ;
   RatioTiers.Caption:=StrfMontant(RatioDef*100,7,2,'',TRUE)+'%' ;
   BaseTiers.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[2], CritEdt.AfficheSymbole ) ;
   MntTvaTiers:=Arrondi(TauxTva*Tot[2],CritEdt.Decimale) ;
   TvaTiers.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MntTvaTiers, CritEdt.AfficheSymbole ) ;
   AddRecap(LRecapG,[QR1E_AUXILIAIRE.AsString],[QR1T_LIBELLE.AsString],[Tot[2],MntTvaTiers]) ;
//   LFactLu.Clear ;
   END ;
end;

procedure TFQRTvaFac1.QRDLRecapGNeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    TotRecapG : Array[0..12] Of Double ;
begin
  inherited;
If Not (CritEdt.GL.FormatPrint.PrSepCompte[10]) Then BEGIN MoreData:=FALSE ; Exit ; END ;
MoreData:=PrintRecap(LRecapG, Cod, Lib, TotRecapG) ;
//CodCentr.Caption:='RIEN' ;
If MoreData Then
   BEGIN
//   Rappel.Caption:=Cod ; //CodCentr.Caption:=Lib ;
// LibRecapG.Caption:=FloatToStr(Arrondi(Tva2Taux(StRegime,StTva,StAchat='A')*100,2))+' %' ;
   LibRecapG.Caption:=Lib ;
   BaseTiersG.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotRecapG[0], CritEdt.AfficheSymbole) ;
   TvaTiersG.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotRecapG[1],CritEdt.AfficheSymbole) ;
   END ;
end;

procedure TFQRTvaFac1.BRecapGAfterPrint(BandPrinted: Boolean);
begin
  inherited;
BRecapG.ForceNewPage:=FALSE ;
end;

procedure TFQRTvaFac1.BRecapGBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
If Not (CritEdt.GL.FormatPrint.PrSepCompte[10]) Then PrintBand:=FALSE ;
end;

procedure TFQRTvaFac1.BFinAuxBeforePrint(var PrintBand: Boolean;var Quoi: string);
begin
  inherited;
If Not (CritEdt.GL.FormatPrint.PrSepCompte[8]) Then PrintBand:=FALSE ;
Quoi:=QuoiAux ;
end;

// VL 05/12/2002
procedure TFQRTvaFac1.BValiderClick(Sender: TObject);
begin
if (FTva.ItemIndex = 0) then
	LanceTousLesTaux
else
  inherited;
end;

// VL 04/12/2002 : Lance l'impression de tous les taux les uns après les autres
procedure TFQRTvaFac1.LanceTousLesTaux;
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
